/* Offline unit tests for Travel Desk form logic. Run: node test/form.test.js */
"use strict";
var F = require("../src/form.js");

var failures = 0;
function check(label, actual, expected) {
  var ok = actual === expected;
  if (!ok) {
    failures++;
    console.error("FAIL  " + label + "\n  expected: " + JSON.stringify(expected) + "\n  actual:   " + JSON.stringify(actual));
  }
}

var model = {
  name: "Matt Miller",
  costCenter: "471-0000",
  division: "System Ops",
  bureau: "TAS",
  event: "TRB Annual Meeting",
  location: "Washington, DC",
  eventStart: "2027-01-10",
  confDates: "Jan 10-14, 2027",
  departDate: "2027-01-09",
  returnDate: "2027-01-15",
  reason: "US Planning chair",
  attendeeRole: "committee member",
  funding: "Non-TEWD",
  comments: "BIM TPF",
  modes: { personal: false, state: false, air: true },
  costs: {
    travelMode: "550", luggage: "70", parking: "60",
    lodgingNights: "5", lodgingRate: "250",
    registration: "600", mealsB: "2", mealsL: "0", mealsD: "1",
    taxi: "120", additional: "0", additionalDesc: "",
  },
  thirdParties: [
    { name: "National Academies", project: "TPF-5(372)", packet: true, maxReimb: "2400",
      items: { registration: true, lodging: true, airfare: false, meals: false, ground: false }, notes: "" },
  ],
  meetingLink: "https://trb.org/annual",
};

// 1. Totals: lodging 5*250=1250; grand 550+70+60+1250+600+120 = 2650
var t = F.computeTotals(model);
check("lodging subtotal", t.lodging, 1250);
check("grand total", t.grand, 2650);

// 2. Money-ish strings with $ and commas still parse
check("num strips $ and commas", F._internals.num("$1,250.50"), 1250.5);

// 3. Structured subject
check("subject", F.subjectLine(model), "Travel Auth - Miller - TRB Annual Meeting - 2027-01-10");
check("subject last-first name", F.subjectLine({ name: "Miller, Matt", event: "X" }), "Travel Auth - Miller - X");

// 4. Fiscal-year labels for any org calendar
check("Iowa SFY for Jan 2027", F.fiscalLabel("2027-01-10", 7, "SFY"), "SFY27");
check("Iowa SFY for Aug 2026", F.fiscalLabel("2026-08-15", 7, "SFY"), "SFY27");
check("Federal FFY for Aug 2026", F.fiscalLabel("2026-08-15", 10, "FFY"), "FFY26");
check("Federal FFY for Oct 2026", F.fiscalLabel("2026-10-02", 10, "FFY"), "FFY27");
check("Calendar FY", F.fiscalLabel("2027-03-01", 1, "FY"), "FY27");
check("Default prefix", F.fiscalLabel("2027-03-01", 1, ""), "FY27");
check("bad date", F.fiscalLabel("", 7, "SFY"), "");

// 5. Planner row mapped against the real SFY27 template headers
var headers27 = ["Event Name", "City, State", "Start date", "Division", "Bureau",
  "Planned Attendee", "Estimated\nCost", "Third-party or grant reimburse?",
  "Reimburse %", "State fiscal year", "TEWD request", "Additional Comments"];
var IOWA = { fyStartMonth: 7, fyPrefix: "SFY" };
var row = F.plannerRow(headers27, model, IOWA);
check("row event", row[0], "TRB Annual Meeting");
check("row city", row[1], "Washington, DC");
check("row start", row[2], "2027-01-10");
check("row division", row[3], "System Ops");
check("row bureau", row[4], "TAS");
check("row attendee", row[5], "Matt Miller");
check("row cost", row[6], "2650");
check("row third party", row[7], "Yes - National Academies");
check("row reimburse pct", row[8], "0.91"); // 2400/2650 = 0.9056 -> 0.91
check("row fiscal year", row[9], "SFY27");
check("row tewd", row[10], "Non-TEWD");
check("row comments", row[11], "BIM TPF");

// 6. SFY26 template variant (role column, COO column stays blank)
var headers26 = ["Event Name", "City, State", "Start date", "Planned Attendee",
  "Attendee role", "Estimated\nCost", "Third-party or grant reimburse?",
  "Reimburse %", "TEWD", "COO approved"];
var row26 = F.plannerRow(headers26, model, IOWA);
check("row26 role", row26[4], "committee member");
check("row26 coo blank", row26[9], "");
check("row26 tewd", row26[8], "Non-TEWD");

// 7. No third party -> "No" and blank pct
check("no third party", F.thirdPartySummary({ thirdParties: [] }), "No");
check("no pct without reimb", F.reimbursePct({ costs: { registration: "100" }, thirdParties: [] }), "");

// 8. Generic-org planner headers (non-Iowa) map too
var headersGeneric = ["Traveler Name", "Conference", "Destination", "Trip Start Date",
  "Department", "Estimated Cost", "Funding Source", "Approved?"];
var rowG = F.plannerRow(headersGeneric, model, { fyStartMonth: 1, fyPrefix: "FY" });
check("generic traveler", rowG[0], "Matt Miller");
check("generic conference", rowG[1], "TRB Annual Meeting");
check("generic destination", rowG[2], "Washington, DC");
check("generic date", rowG[3], "2027-01-10");
check("generic department", rowG[4], "System Ops");
check("generic cost", rowG[5], "2650");
check("generic funding", rowG[6], "Non-TEWD");
check("generic approver column blank", rowG[7], "");

// 9. Form HTML: escaped, totals included, org branding
var html = F.formHtml(model, { orgName: "Iowa DOT", fundingLabel: "TEWD" });
check("html has total", html.indexOf("$2,650.00") !== -1, true);
check("html has lodging math", html.indexOf("5 nights @ $250.00 = $1,250.00") !== -1, true);
var xss = F.formHtml({ name: "<script>x</script>", costs: {}, modes: {}, thirdParties: [] });
check("html escapes", xss.indexOf("<script>") === -1, true);
check("html org name", html.indexOf("Iowa DOT") !== -1, true);
check("html funding label", html.indexOf("TEWD") !== -1, true);
var noOrg = F.formHtml(model, {});
check("html works with no org", noOrg.indexOf("Travel Authorization Request") !== -1, true);

// 10. Fiscal-year planner routing
var PLANNERS = {
  "SFY27": { tableName: "T27", wbRef: { name: "Planner 2027" } },
  "SFY26": { tableName: "T26", wbRef: { name: "Planner 2026" } },
  "*": { tableName: "TAll", wbRef: { name: "Catch-all" } },
};
check("fy exact match", F.pickPlanner(PLANNERS, "SFY27").planner.tableName, "T27");
check("fy other year", F.pickPlanner(PLANNERS, "SFY26").key, "SFY26");
check("fy fallback to catch-all", F.pickPlanner(PLANNERS, "SFY28").key, "*");
check("no planners -> null", F.pickPlanner({}, "SFY27"), null);
check("no match no catch-all -> null", F.pickPlanner({ "SFY26": PLANNERS["SFY26"] }, "SFY28"), null);

if (failures) {
  console.error("\n" + failures + " form test(s) FAILED");
  process.exit(1);
}
console.log("All Travel Desk form tests passed.");
