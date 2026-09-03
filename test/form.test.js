/* Offline unit tests for Travel Desk form logic. Run: node test/form.test.js */
"use strict";
var F = require("../src/form.js");

var T = require("./assert.js").suite("form");
var check = T.check, truthy = T.truthy;

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

// 11. Booking-email matching (Concur follow-up)
var trip = { event: "TRB Annual Meeting", location: "Washington, DC",
  createdAt: "2026-07-01T10:00:00Z", eventStart: "2027-01-10", returnDate: "2027-01-15" };
var mails = [
  { subject: "Your trip to Washington", bodyPreview: "Flight DSM-DCA Jan 9", receivedDateTime: "2026-07-20T09:00:00Z", webLink: "w1" },
  { subject: "Your trip to Denver", bodyPreview: "Flight DSM-DEN", receivedDateTime: "2026-07-21T09:00:00Z", webLink: "w2" },
  { subject: "Old itinerary", bodyPreview: "Washington", receivedDateTime: "2026-06-01T09:00:00Z", webLink: "w3" },
];
var mb = F.matchBooking(trip, mails);
check("confident match by city", mb.confident && mb.confident.webLink, "w1");
// Pre-request mail is now considered - you often register before filing the
// paperwork - but a confirmation issued AFTER the request still wins, so the
// look-back never overrides the obvious answer.
check("pre-request email is now a candidate", mb.candidates.some(function (e) { return e.webLink === "w3"; }), true);
check("but the post-request confirmation still wins", mb.confident && mb.confident.webLink, "w1");

// A conference registration from an arbitrary sender must match. Requiring a
// pre-listed sender domain made a real EDC-8 registration confirmation
// invisible - it was dropped before matching ever ran.
var regTrip = { event: "EDC-8 Midwest Peer Exchange", location: "Ames, IA",
  createdAt: "2026-08-01T10:00:00Z", eventStart: "2026-09-15", returnDate: "2026-09-17" };
var regMails = [
  { subject: "EDC-8 Midwest Peer Exchange Registration Confirmation", bodyPreview: "You are registered.",
    from: "someone@fhwa.dot.gov", receivedDateTime: "2026-08-23T12:00:00Z", webLink: "r1" },
];
check("registration from an unlisted sender matches",
  F.matchBooking(regTrip, regMails, ["concursolutions.com"]).confident.webLink, "r1");

// You register for a conference FIRST and file the authorisation afterwards,
// so the confirmation is already in the inbox when the request is raised.
// Requiring booking-after-request made that ordinary case invisible.
var lateFiled = { event: "EDC-8 Midwest Peer Exchange", location: "Ames, IA",
  createdAt: "2026-08-23T14:00:00Z", eventStart: "2026-09-15", returnDate: "2026-09-17" };
check("confirmation that arrived BEFORE the request still matches",
  F.matchBooking(lateFiled, regMails, []).confident.webLink, "r1");

// ...but not indefinitely far back.
var ancient = [{ subject: "EDC-8 Midwest Peer Exchange Registration Confirmation",
  bodyPreview: "registered", from: "x@y.gov",
  receivedDateTime: "2025-01-01T09:00:00Z", webLink: "old" }];
check("a confirmation from a year earlier is not matched",
  F.matchBooking(lateFiled, ancient, []).confident, null);

// Mentioning the city is not enough on its own - it must look like a booking.
var chatter = [
  { subject: "Lunch in Ames next week?", bodyPreview: "Are you around",
    from: "colleague@x.gov", receivedDateTime: "2026-08-23T12:00:00Z", webLink: "c1" },
];
check("ordinary mail mentioning the city is not a booking",
  F.matchBooking(regTrip, chatter, []).confident, null);

// A trusted sender is taken at its word even without booking language.
var terse = [
  { subject: "Ames", bodyPreview: "", from: "noreply@concursolutions.com",
    receivedDateTime: "2026-08-23T12:00:00Z", webLink: "t1" },
];
check("trusted sender needs no booking words",
  F.matchBooking(regTrip, terse, ["concursolutions.com"]).confident.webLink, "t1");

// ambiguous: two Washington emails in window -> no auto match, both candidates
var mails2 = mails.concat([{ subject: "Updated trip to Washington", bodyPreview: "", receivedDateTime: "2026-07-22T09:00:00Z", webLink: "w4" }]);
var mb2 = F.matchBooking(trip, mails2);
check("ambiguous -> no confident", mb2.confident, null);
check("candidates in window", mb2.candidates.length, 4);

// 12. Status column -> "Requested" (coordinator follow-up queue)
var stRow = F.plannerRow(["Event Name", "Status", "COO approved"], model, { fyStartMonth: 7, fyPrefix: "SFY" });
check("status column requested", stRow[1], "Requested");
check("approval column still blank", stRow[2], "");


/* ------------------------------------------- third-party receivables */

var NOWR = new Date("2026-08-01T12:00:00Z");
var tripsR = [
  { id: "t1", name: "Ann Lee", event: "TRB Annual", location: "Washington, DC",
    returnDate: "2026-06-15", thirdParties: [
      { name: "AASHTO", contact: "billing@aashto.org", project: "AG-42", maxReimb: "840" },
      { name: "FHWA Grant", contact: "", project: "", maxReimb: "" }],
    reimb: { 1: { status: "received", on: "2026-07-01" } } },
  { id: "t2", name: "Bob Ray", event: "Future Conf", location: "Denver, CO",
    returnDate: "2026-12-01", thirdParties: [{ name: "Org", maxReimb: "100" }] },
  { id: "t3", name: "Sue Kim", event: "No 3rd party trip", returnDate: "2026-05-01", thirdParties: [] },
];
var recv = F.receivables(tripsR, NOWR);
check("future trip + no-tp trip excluded", recv.length, 2);
check("open receivable first", recv[0].entity, "AASHTO");
check("age counts from the end of the return day, local", recv[0].ageDays, 46);
check("amount parsed", recv[0].amount, 840);
check("blank amount -> null", recv[1].amount, null);
check("received status carried", recv[1].status, "received");

var packet = F.workdayPacketHtml(recv[0], { funding: "TEWD", fundingLabel: "TEWD", costCenter: "CC-77" });
check("packet names the entity", packet.indexOf("AASHTO") !== -1, true);
check("packet shows expected amount", packet.indexOf("$840") !== -1, true);
check("packet has blank WD receivable line", packet.indexOf("Workday receivable #: ______") !== -1, true);
check("packet carries funding + cost center", packet.indexOf("TEWD") !== -1 && packet.indexOf("CC-77") !== -1, true);

var rem = F.reminderHtml(recv[0], { orgName: "Iowa DOT" });
check("reminder states amount + event", rem.indexOf("$840") !== -1 && rem.indexOf("TRB Annual") !== -1, true);
check("reminder references project", rem.indexOf("AG-42") !== -1, true);
check("reminder signs org", rem.indexOf("Iowa DOT") !== -1, true);

/* --- regressions found in the 2026-08-02 bug sweep --- */

// planners spell it "3rd Party" as often as "Third party"
check("3rd Party header populated",
  F.plannerRow(["3rd Party"], { thirdParties: [{ name: "AASHTO" }] }, {})[0], "Yes - AASHTO");
check("Reimbursed-by header populated",
  F.plannerRow(["Reimbursed by"], { thirdParties: [{ name: "AASHTO" }] }, {})[0], "Yes - AASHTO");
check("% Reimbursed still gets the percentage, not the names",
  F.plannerRow(["% Reimbursed"], { costs: { registration: "100" },
    thirdParties: [{ name: "AASHTO", maxReimb: "50" }] }, {})[0], "0.5");

// a genuine $0 cap is an answer, a blank is not
var zeroCap = F.receivables([{ id: "z", returnDate: "2026-01-01",
  thirdParties: [{ name: "X", maxReimb: "0" }] }], new Date(2026, 5, 1));
check("zero cap stays zero", zeroCap[0].amount, 0);
check("zero renders as a number, not TBD",
  F.workdayPacketHtml(zeroCap[0], {}).indexOf("$0") !== -1, true);
var blankCap = F.receivables([{ id: "b", returnDate: "2026-01-01",
  thirdParties: [{ name: "X", maxReimb: "" }] }], new Date(2026, 5, 1));
check("blank cap is unknown", blankCap[0].amount, null);
check("blank renders as TBD",
  F.workdayPacketHtml(blankCap[0], {}).indexOf("(amount TBD)") !== -1, true);

// a date-only return means the END of that day, locally
var sameDay = F.receivables([{ id: "s", returnDate: "2026-08-02",
  thirdParties: [{ name: "X", maxReimb: "10" }] }], new Date(2026, 7, 2, 9, 0, 0));
check("trip returning today is not yet a receivable", sameDay.length, 0);
var nextDay = F.receivables([{ id: "s", returnDate: "2026-08-02",
  thirdParties: [{ name: "X", maxReimb: "10" }] }], new Date(2026, 7, 3, 9, 0, 0));
check("the morning after, it is", nextDay.length, 1);
check("age starts at 0 the next morning", nextDay[0].ageDays, 0);

/* --- multi-traveler planner rows (one row per person per event) --- */

var G = require("../src/xlsxgen.js");
var MH = F.DEFAULT_PLANNER_HEADERS;
var group = {
  name: "Matt Miller", division: "SOD", bureau: "Field", attendeeRole: "attendee",
  event: "TRB", location: "DC", eventStart: "2027-01-11",
  costs: { travelMode: "600", registration: "400" },
  otherStaff: "Jane Doe, MVD, Ops, speaker\nBob Roe",
};
var perPerson = F.plannerRows(MH, group, { fyStartMonth: 7, fyPrefix: "SFY" });
check("one row per traveler", perPerson.length, 3);
check("primary first", perPerson[0][0], "Matt Miller");
check("second traveler name", perPerson[1][0], "Jane Doe");
check("second traveler division", perPerson[1][1], "MVD");
check("second traveler role", perPerson[1][6], "speaker");
check("blank fields inherit the primary's", perPerson[2][1], "SOD");
check("blank role inherits too", perPerson[2][6], "attendee");
check("per-person repeats the cost", perPerson[1][7], "1000");
check("shared trip details repeat", perPerson[2][3], "TRB");

var split = F.plannerRows(MH, group, { fyStartMonth: 7, fyPrefix: "SFY", costMode: "split" });
check("split divides across travelers", split[0][7], "333");
check("split applies to everyone", split[2][7], "333");

var solo = F.plannerRows(MH, { name: "Solo", costs: { registration: "500" } }, {});
check("no extra travelers -> one row", solo.length, 1);
check("solo split is a no-op", F.plannerRows(MH, { name: "Solo", costs: { registration: "500" } },
  { costMode: "split" })[0][7], "500");

check("parseTravelers ignores blank lines",
  F.parseTravelers("Jane Doe\n\n  \nBob Roe").length, 2);
check("parseTravelers trims parts",
  F.parseTravelers("  Jane Doe ,  MVD  ")[0].division, "MVD");
check("parseTravelers on empty input", F.parseTravelers("").length, 0);

/* --- generated workbook --- */

var wbg = G.buildWorkbook(MH, "Planner");
check("five OOXML parts", Object.keys(wbg.parts).length, 5);
// Graph's tables/add rejects a quoted sheet name. usedRange - the path
// "Make this sheet a table" uses successfully - returns Planner!A1:R1,
// so the generated address must match that shape exactly.
check("header range matches column count (18 cols -> R)", wbg.range, "Planner!A1:R1");
check("headers land in the sheet",
  wbg.parts["xl/worksheets/sheet1.xml"].indexOf("<t xml:space=\"preserve\">Traveler</t>") !== -1, true);
check("sheet name in workbook.xml",
  wbg.parts["xl/workbook.xml"].indexOf('name="Planner"') !== -1, true);
check("column letters past Z", G.colLetter(27), "AA");
// The sheet name is always ours ("Planner"), never user input, so the
// address is emitted unquoted. Documented as an assumption rather than
// silently relied on: a name needing quotes would not survive this.
check("simple sheet name is emitted unquoted", G.headerRange("Planner", 2), "Planner!A1:B1");
check("the address shape Graph returns and accepts",
  /^[A-Za-z0-9 ]+![A-Z]+\d+:[A-Z]+\d+$/.test(G.headerRange("Planner", 18)), true);
check("ampersand in a header is escaped",
  G.buildWorkbook(["A & B"]).parts["xl/worksheets/sheet1.xml"].indexOf("A &amp; B") !== -1, true);

/* ------------------- setup by invitation (no copy/paste) ------------------- */

check("subject carries the org", F.setupSubject("Iowa DOT"), "Travel Desk setup \u2014 Iowa DOT");
check("subject without an org", F.setupSubject(""), "Travel Desk setup");

var theCode = "eyJjb29yZEVtYWlsIjoia2VyaUBkb3QuZ292In0=";
var invite = F.setupInviteHtml({ code: theCode, orgName: "Iowa DOT",
  coordName: "Keri Greenfield", coordEmail: "keri@dot.gov" });
check("code round-trips out of the invitation", F.extractSetupCode(invite), theCode);
check("invitation names the org", invite.indexOf("Iowa DOT") !== -1, true);
check("invitation tells them what to click", invite.indexOf("Find my setup") !== -1, true);
check("invitation names the sender", invite.indexOf("Keri Greenfield") !== -1, true);

// mail clients wrap, re-encode and inject markup between the markers
var mangled = invite.replace(theCode,
  theCode.slice(0, 6) + '<span class="x">' + theCode.slice(6, 12) + "</span>\r\n   " + theCode.slice(12));
check("survives injected markup and wrapping", F.extractSetupCode(mangled), theCode);
check("a plain email yields nothing", F.extractSetupCode("Lunch on Friday?"), "");
check("truncated marker yields nothing", F.extractSetupCode("[[TD-SETUP]]abc"), "");

check("picks the newest invitation that has a code",
  F.pickInvite([
    { receivedDateTime: "2026-02-01", body: F.setupInviteHtml({ code: "b2xk" }) },
    { receivedDateTime: "2026-08-01", body: F.setupInviteHtml({ code: "bmV3" }) },
    { receivedDateTime: "2026-09-01", body: "<p>unrelated mail</p>" },
  ]).receivedDateTime, "2026-08-01");
check("no usable invitation -> null", F.pickInvite([{ receivedDateTime: "2026-01-01", body: "hi" }]), null);
check("empty inbox -> null", F.pickInvite([]), null);

// the code an invitation carries must restore through the same path a paste does
check("invitation code decodes to real settings",
  JSON.parse(decodeURIComponent(escape(Buffer.from(F.extractSetupCode(invite), "base64").toString("binary")))).coordEmail,
  "keri@dot.gov");

// -------------------------------- a trip as a calendar entry

var TRIP = {
  event: "EDC-8 Midwest Peer Exchange", location: "St. Louis, MO",
  eventStart: "2026-10-15", departDate: "2026-10-14", returnDate: "2026-10-16",
  confDates: "October 15, 2026, 8am-5pm",
  meetingLink: "https://www.eventleaf.com/e/edc8midwest",
};
var HOTEL = {
  name: "Hyatt Regency St. Louis at the Arch", checkIn: "2026-10-13",
  checkOut: "2026-10-17", address: "St. Louis, MO", confirmation: "MV59BR3A",
};

var withHotel = F.calendarEntry(TRIP, HOTEL);
check("the hotel is the location", withHotel.location, HOTEL.name);
check("a booking's dates beat the trip's own", withHotel.start, "2026-10-13");
check("at both ends", withHotel.end, "2026-10-17");
check("the confirmation number is in the body",
  /MV59BR3A/.test(withHotel.body), true);
check("so is the meeting link", /eventleaf/.test(withHotel.body), true);

var noHotel = F.calendarEntry(TRIP, null);
check("without a booking, the travel dates are used", noHotel.start, "2026-10-14");
check("and the destination is the location", noHotel.location, "St. Louis, MO");

check("a trip with no dates yields nothing rather than today",
  F.calendarEntry({ event: "X" }, null), null);
check("an end before the start is corrected, not emitted",
  F.calendarEntry({ event: "X", departDate: "2026-10-14", returnDate: "2026-10-10" }, null).end,
  "2026-10-14");
check("a single-day trip still produces an entry",
  F.calendarEntry({ event: "X", eventStart: "2026-10-15" }, null).start, "2026-10-15");

// The appointment is TIMED - AppointmentForm has no all-day flag - so the end
// is literal. This is what the pane hands Outlook.
// check() in this file compares with ===, so arrays never match; joined.
function span(e) {
  return e.start + "T00:00:00 -> " + e.end + "T23:59:00";
}
check("a trip ends on the day it returns, not the day after",
  span(F.calendarEntry({ event: "X", departDate: "2026-10-14", returnDate: "2026-10-16" }, null)),
  "2026-10-14T00:00:00 -> 2026-10-16T23:59:00");
check("a single-day trip is not zero length",
  span(F.calendarEntry({ event: "X", eventStart: "2026-10-15" }, null)),
  "2026-10-15T00:00:00 -> 2026-10-15T23:59:00");

// The coordinator pastes a To: line as often as a typed list, and the first
// statement of sendInvites depends on this.
check("semicolon separated", F.extractEmails("jane@x.gov; bob@x.gov").join(","), "jane@x.gov,bob@x.gov");
check("an Outlook To: line", F.extractEmails("Doe, Jane <jane.doe@iowadot.us>; Roe, Bob <bob.roe@iowadot.us>").join(","),
  "jane.doe@iowadot.us,bob.roe@iowadot.us");
check("one per line", F.extractEmails("a@x.gov\nb@x.gov\nc@x.gov").length, 3);
check("the same person twice is one invitation",
  F.extractEmails("jane@x.gov, jane@X.GOV").length, 1);
check("a trailing full stop is not part of the address",
  F.extractEmails("Send it to jane@x.gov.").join(","), "jane@x.gov");
check("nothing to send to", F.extractEmails("no addresses here").length, 0);
check("empty input", F.extractEmails("").length, 0);

T.done("All Travel Desk form tests passed.");
