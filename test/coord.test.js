/* Offline unit tests for the coordinator view. Run: node test/coord.test.js */
"use strict";
var C = require("../src/coord.js");
var F = require("../src/form.js");

var failures = 0;
function check(label, actual, expected) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    failures++;
    console.error("FAIL  " + label + "\n  expected: " + JSON.stringify(expected) + "\n  actual:   " + JSON.stringify(actual));
  }
}

var H = F.DEFAULT_PLANNER_HEADERS;
var NOW = new Date(2026, 7, 3, 9, 0, 0); // Mon 2026-08-03

/** Build a planner row BY HEADER NAME. Positional fixtures silently
 *  mis-map the moment a column is added, which is exactly what happened
 *  when the lifecycle columns landed. */
function row(fields) {
  return H.map(function (h) { return fields[h] == null ? "" : String(fields[h]); });
}

/* ------- what the writer writes, the reader must read back ------- */

var written = F.plannerRow(H, {
  name: "Matt Miller", division: "SOD", bureau: "Field", event: "TRB Annual",
  location: "Washington, DC", eventStart: "2027-01-11", attendeeRole: "attendee",
  costs: { travelMode: "600", registration: "400" },
  funding: "TEWD", comments: "annual",
  thirdParties: [{ name: "AASHTO", maxReimb: "500" }],
}, { fyStartMonth: 7, fyPrefix: "SFY" });

var back = C.mapRows(H, [written])[0];
check("round-trip traveler", back.traveler, "Matt Miller");
check("round-trip division", back.division, "SOD");
check("round-trip event", back.event, "TRB Annual");
check("round-trip destination", back.destination, "Washington, DC");
check("round-trip date", back.date, "2027-01-11");
check("round-trip cost is numeric", back.cost, 1000);
check("round-trip third party", back.thirdParty, "Yes - AASHTO");
check("round-trip fiscal year", back.fy, "SFY27");
check("round-trip status", back.status, "Requested");

/* ------------------------------ summarize ------------------------------ */

var rows = [
  row({ Traveler: "Ann Lee", Division: "MVD", Event: "TRB Annual", Destination: "DC",
        "Start date": "2026-09-10", Role: "attendee", "Estimated cost": 1200,
        "3rd Party": "No", Funding: "TEWD", "Fiscal year": "SFY27", Status: "Requested" }),
  row({ Traveler: "Bob Roe", Division: "MVD", Event: "TRB Annual", Destination: "DC",
        "Start date": "2026-09-10", Role: "attendee", "Estimated cost": 1200,
        "3rd Party": "No", Funding: "TEWD", "Fiscal year": "SFY27", Status: "Booked" }),
  row({ Traveler: "Sue Kim", Division: "TDD", Event: "Rail Summit", Destination: "Chicago",
        "Start date": "2026-06-01", Role: "speaker", "Estimated cost": 800,
        "% Reimbursed": "0.5", "3rd Party": "Yes - AAR", Funding: "TEWD",
        "Fiscal year": "SFY26", Status: "Booked" }),
  row({ Traveler: "Dan Poe", Division: "TDD", Event: "Bridge Conf", Destination: "Denver",
        "Start date": "2026-11-02", Role: "attendee", "Estimated cost": 1500,
        "3rd Party": "No", Funding: "Grant", "Fiscal year": "SFY27", Status: "Requested" }),
  row({}),
];
var recs = C.mapRows(H, rows);
check("blank rows dropped", recs.length, 4);

var sum = C.summarize(recs, { now: NOW });
check("upcoming excludes past trips", sum.upcoming.length, 3);
check("upcoming sorted by date", sum.upcoming[0].traveler, "Ann Lee");
check("unbooked upcoming only", sum.unbooked.map(function (r) { return r.traveler; }), ["Ann Lee", "Dan Poe"]);
check("finished trip with a third party surfaces", sum.endedThirdParty.map(function (r) { return r.traveler; }), ["Sue Kim"]);
check("total cost", sum.totalCost, 4700);
check("by division, biggest first (MVD 2400 > TDD 2300)", sum.byDivision.map(function (d) { return d.division; }), ["MVD", "TDD"]);
check("division totals", sum.byDivision[0].cost, 2400);
check("fiscal years newest first", sum.byFy.map(function (f) { return f.fy; }), ["SFY27", "SFY26"]);

var fy27 = C.summarize(recs, { now: NOW, fy: "SFY27" });
check("fy filter applies", fy27.count, 3);
check("fy filter changes the total", fy27.totalCost, 3900);

check("booked recognises Approved", C.isBooked({ status: "Approved" }), true);
check("booked recognises Requested as not booked", C.isBooked({ status: "Requested" }), false);
check("third party 'No' is not owed", C.hasThirdParty({ thirdParty: "No" }), false);

/* ---------------------------- reconciliation ---------------------------- */

check("parse a real subject",
  C.parseAuthSubject("Travel Auth - Miller - TRB Annual - 2027-01-11"),
  { last: "Miller", event: "TRB Annual", date: "2027-01-11" });
check("parse survives a reply prefix",
  C.parseAuthSubject("RE: Travel Auth - Lee - Rail Summit - 2026-06-01").last, "Lee");
check("parse keeps hyphens inside the event name",
  C.parseAuthSubject("Travel Auth - Poe - Bridge - Tunnel Conf - 2026-11-02").event, "Bridge - Tunnel Conf");
check("parse without a date", C.parseAuthSubject("Travel Auth - Kim - Rail Summit").event, "Rail Summit");
check("unrelated subject ignored", C.parseAuthSubject("Lunch tomorrow?"), null);

check("lastName from 'Matt Miller'", C.lastNameOf("Matt Miller"), "Miller");
check("lastName from 'Miller, Matt'", C.lastNameOf("Miller, Matt"), "Miller");

var rec = C.reconcile(recs, [
  "Travel Auth - Lee - TRB Annual - 2026-09-10",
  "Travel Auth - Roe - TRB Annual - 2026-09-10",
  "Travel Auth - Kim - Rail Summit - 2026-06-01",
  "Travel Auth - Ghost - Ferry Expo - 2026-12-01",   // authorised, never planned
  "Re: something else entirely",
]);
check("planner row with no authorization", rec.rowsWithoutAuth.map(function (r) { return r.traveler; }), ["Dan Poe"]);
check("authorization with no planner row", rec.authsWithoutRow.map(function (p) { return p.last; }), ["Ghost"]);
check("matched ones appear in neither",
  rec.rowsWithoutAuth.concat(rec.authsWithoutRow).length, 2);

/* ------------------- write-back: locating and settling ------------------- */

check("finds the right row", C.findRowIndex(recs, { traveler: "Dan Poe", event: "Bridge Conf" }), 3);
check("matches on last name alone", C.findRowIndex(recs, { traveler: "Poe", event: "Bridge Conf" }), 3);
check("matches 'Poe, Dan' form", C.findRowIndex(recs, { traveler: "Poe, Dan", event: "Bridge Conf" }), 3);
check("unknown trip -> -1", C.findRowIndex(recs, { traveler: "Nobody", event: "Nowhere" }), -1);

// two travellers to the same event: the date breaks the tie
var dupes = C.mapRows(H, [
  row({ Traveler: "Ann Lee", Division: "MVD", Event: "TRB Annual", Destination: "DC",
        "Start date": "2026-09-10", "Estimated cost": 1200, "Fiscal year": "SFY27", Status: "Requested" }),
  row({ Traveler: "Ann Lee", Division: "MVD", Event: "TRB Annual", Destination: "DC",
        "Start date": "2027-09-10", "Estimated cost": 1300, "Fiscal year": "SFY28", Status: "Requested" }),
]);
check("same person + event, later date picked",
  C.findRowIndex(dupes, { traveler: "Ann Lee", event: "TRB Annual", date: "2027-09-10" }), 1);
check("same person + event, earlier date picked",
  C.findRowIndex(dupes, { traveler: "Ann Lee", event: "TRB Annual", date: "2026-09-10" }), 0);

check("settling keeps the payer's name",
  C.markSettled("Yes - AASHTO", "2026-08-15"), "Yes - AASHTO (reimbursed 2026-08-15)");
check("settling twice doesn't stack markers",
  C.markSettled("Yes - AASHTO (reimbursed 2026-08-15)", "2026-09-01"),
  "Yes - AASHTO (reimbursed 2026-09-01)");
check("settled cells are recognised", C.isSettled("Yes - AASHTO (reimbursed 2026-08-15)"), true);
check("unsettled cells are not", C.isSettled("Yes - AASHTO"), false);
check("a settled trip stops showing as owing",
  C.hasThirdParty({ thirdParty: "Yes - AASHTO (reimbursed 2026-08-15)" }), false);
check("an unsettled one still does", C.hasThirdParty({ thirdParty: "Yes - AASHTO" }), true);

// end to end: settle Sue's row and she leaves the coordinator's owing list
var tpCol = H.indexOf("3rd Party");
var settled = C.mapRows(H, rows.map(function (r) {
  return r[0] === "Sue Kim"
    ? r.map(function (v, i) { return i === tpCol ? C.markSettled(v, "2026-07-01") : v; })
    : r;
}));
check("owing list clears once settled",
  C.summarize(settled, { now: NOW }).endedThirdParty.length, 0);

/* ---------------- trip lifecycle: estimate, actuals, approval ------------- */

var lifeRows = C.mapRows(H, [
  row({ Traveler: "Ann Lee", Division: "MVD", Event: "TRB", "Start date": "2026-09-01",
        "Estimated cost": 1000, Status: "Requested", "Fiscal year": "SFY27" }),
  row({ Traveler: "Bob Roe", Division: "MVD", Event: "Rail", "Start date": "2026-08-15",
        "Estimated cost": 2000, Status: "Requested", "Fiscal year": "SFY27" }),
  row({ Traveler: "Sue Kim", Division: "TDD", Event: "Bridge", "Start date": "2026-06-01",
        "Estimated cost": 1000, "Actual cost": 1450, Status: "Actuals submitted", "Fiscal year": "SFY27" }),
  row({ Traveler: "Dan Poe", Division: "TDD", Event: "Ferry", "Start date": "2026-05-01",
        "Estimated cost": 1000, "Actual cost": 1050, Status: "Actuals submitted", "Fiscal year": "SFY27" }),
  row({ Traveler: "Eve Fox", Division: "TDD", Event: "Done", "Start date": "2026-04-01",
        "Estimated cost": 900, "Actual cost": 880, Status: "Closed", "Fiscal year": "SFY27",
        "Approved by": "Keri", "Approved date": "2026-04-20" }),
]);

check("actual cost reads back as a number", lifeRows[2].actualCost, 1450);
check("approver reads back", lifeRows[4].approvedBy, "Keri");
check("approved date does not steal the trip date", lifeRows[4].date, "2026-04-01");

var q = C.queues(lifeRows);
check("approval queue holds only unapproved requests",
  q.awaitingApproval.map(function (r) { return r.traveler; }), ["Bob Roe", "Ann Lee"]);
check("approval queue is soonest-departing first", q.awaitingApproval[0].traveler, "Bob Roe");
check("close queue holds only submitted actuals",
  q.awaitingClose.map(function (r) { return r.traveler; }), ["Sue Kim", "Dan Poe"]);
check("worst overrun surfaces first", q.awaitingClose[0].variancePct, 45);
check("a modest overrun is not flagged", q.awaitingClose[1].overBudget, false);
check("a 45% overrun is flagged", q.awaitingClose[0].overBudget, true);
check("already-closed trips are in neither queue",
  q.awaitingApproval.concat(q.awaitingClose).some(function (r) { return r.traveler === "Eve Fox"; }), false);

var bt = C.budgetTruth(lifeRows, { fy: "SFY27" });
check("variance judges only trips that have actuals", bt.estimatedClosed, 2900);
check("actual total", bt.actual, 3380);
check("like-for-like variance percent", bt.variancePct, 17);
check("uncommitted estimates are reported separately", bt.estimatedAll - bt.estimatedClosed, 3000);
check("open trip count", bt.tripsWithout, 2);

check("variance helper", JSON.stringify(F.variance(1000, 1250)), '{"delta":250,"pct":25}');
check("variance unknown without an actual", F.variance(1000, 0), null);
check("tolerance respected", F.varianceFlag(1000, 1150), "ok");
check("over tolerance", F.varianceFlag(1000, 1400), "over");
check("under tolerance", F.varianceFlag(1000, 600), "under");
check("authorized states", [F.isAuthorized("Requested"), F.isAuthorized("Approved"),
  F.isAuthorized("Closed"), F.isAuthorized("Not approved")], [false, true, true, false]);

// a new request must never look like a completed one
var fresh = F.plannerRow(H, { name: "New Person", event: "Future", eventStart: "2027-03-01",
  costs: { registration: "500" } }, {});
check("new request has no actual cost", fresh[H.indexOf("Actual cost")], "");
check("new request has no approver", fresh[H.indexOf("Approved by")], "");
check("new request has no approval date", fresh[H.indexOf("Approved date")], "");
check("new request starts as Requested", fresh[H.indexOf("Status")], "Requested");

if (failures) {
  console.error("\n" + failures + " coordinator test(s) FAILED");
  process.exit(1);
}
console.log("All coordinator tests passed.");
