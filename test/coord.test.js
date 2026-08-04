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
  ["Ann Lee", "MVD", "", "TRB Annual", "DC", "2026-09-10", "attendee", "1200", "", "No", "TEWD", "SFY27", "Requested", "", ""],
  ["Bob Roe", "MVD", "", "TRB Annual", "DC", "2026-09-10", "attendee", "1200", "", "No", "TEWD", "SFY27", "Booked", "", ""],
  ["Sue Kim", "TDD", "", "Rail Summit", "Chicago", "2026-06-01", "speaker", "800", "0.5", "Yes - AAR", "TEWD", "SFY26", "Booked", "", ""],
  ["Dan Poe", "TDD", "", "Bridge Conf", "Denver", "2026-11-02", "attendee", "1500", "", "No", "Grant", "SFY27", "Requested", "", ""],
  ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
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
  ["Ann Lee", "MVD", "", "TRB Annual", "DC", "2026-09-10", "", "1200", "", "No", "", "SFY27", "Requested", "", ""],
  ["Ann Lee", "MVD", "", "TRB Annual", "DC", "2027-09-10", "", "1300", "", "No", "", "SFY28", "Requested", "", ""],
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
var settled = C.mapRows(H, rows.map(function (r) {
  return r[0] === "Sue Kim" ? r.map(function (v, i) { return i === 9 ? C.markSettled(v, "2026-07-01") : v; }) : r;
}));
check("owing list clears once settled",
  C.summarize(settled, { now: NOW }).endedThirdParty.length, 0);

if (failures) {
  console.error("\n" + failures + " coordinator test(s) FAILED");
  process.exit(1);
}
console.log("All coordinator tests passed.");
