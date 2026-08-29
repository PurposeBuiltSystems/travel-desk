/* Offline unit tests for reading a confirmation email. Run: node test/mailparse.test.js */
"use strict";
var M = require("../src/mailparse.js");

var failures = 0;
function check(label, actual, expected) {
  var a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a !== e) {
    failures++;
    console.error("FAIL  " + label + "\n  expected: " + e + "\n  actual:   " + a);
  }
}
function truthy(label, v) {
  if (!v) { failures++; console.error("FAIL  " + label + "\n  expected truthy, got: " + JSON.stringify(v)); }
}

// ---------------------------------------------------------------- subject

check("strips a single forward prefix",
  M.eventFromSubject("Fw: EDC-8 Midwest Peer Exchange Registration Confirmation"),
  "EDC-8 Midwest Peer Exchange");
check("strips stacked prefixes",
  M.eventFromSubject("RE: Fwd: FW: TRB Annual Meeting - Registration Confirmed"),
  "TRB Annual Meeting");
check("strips an [EXTERNAL] tag",
  M.eventFromSubject("[EXTERNAL] Your registration for AASHTO Spring Meeting"),
  "AASHTO Spring Meeting");
check("keeps a name that is only the event",
  M.eventFromSubject("Midwest Transportation Summit 2027"),
  "Midwest Transportation Summit 2027");
check("drops a trailing confirmation number",
  M.eventFromSubject("ITE Annual Meeting - Confirmation #A8F32K91"),
  "ITE Annual Meeting");
check("thank-you-for-registering phrasing",
  M.eventFromSubject("Thank you for registering for the Rural Transit Conference"),
  "the Rural Transit Conference");

// ------------------------------------------------------------------ dates

check("month-day range in one month",
  M.findDates("The conference runs January 10-14, 2027 in DC.")[0],
  { start: "2027-01-10", end: "2027-01-14", at: 20, kind: "range" });

var cross = M.findDates("Held September 29 - October 1, 2026.")[0];
check("range crossing a month", [cross.start, cross.end], ["2026-09-29", "2026-10-01"]);

var wrap = M.findDates("December 30 - January 2, 2027")[0];
check("range crossing the new year keeps the start in the prior year",
  [wrap.start, wrap.end], ["2026-12-30", "2027-01-02"]);

check("plain long date", M.findDates("March 3, 2027")[0].start, "2027-03-03");
check("day-first date", M.findDates("Meeting on 3 March 2027.")[0].start, "2027-03-03");
check("iso date", M.findDates("Starts 2027-04-05.")[0].start, "2027-04-05");
check("slashed date", M.findDates("Starts 4/5/2027.")[0].start, "2027-04-05");
check("two-digit year", M.findDates("Starts 4/5/27.")[0].start, "2027-04-05");

check("prefers the event date over the date you registered",
  M.eventDates(
    "Thank you — your registration was received on August 23, 2026.\n" +
    "Conference dates: October 6-8, 2026\n" +
    "Payment processed August 23, 2026.",
    "2026-08-23T12:00:00Z"),
  { start: "2026-10-06", end: "2026-10-08" });

check("ignores a date that is already past",
  M.eventDates(
    "Our 2025 event was held June 2, 2025. This year: June 9-11, 2027.",
    "2026-08-23T12:00:00Z").start,
  "2027-06-09");

check("ignores an RSVP deadline",
  M.eventDates(
    "Please respond by September 1, 2026.\nThe summit takes place November 4-6, 2026.",
    "2026-08-23T12:00:00Z"),
  { start: "2026-11-04", end: "2026-11-06" });

check("no dates at all", M.eventDates("See the attached agenda.", "2026-08-23T12:00:00Z"),
  { start: "", end: "" });

// --------------------------------------------------------------- location

check("plain city and state",
  M.findLocation("The event is held in Kansas City, MO at the convention center.").value,
  "Kansas City, MO");

check("full state name",
  M.findLocation("Location: Madison, Wisconsin").value, "Madison, WI");

check("skips the sender's own signature address",
  M.findLocation(
    "Venue: Hyatt Regency, Minneapolis, MN\n\n" +
    "Matthew Miller, CPM\nIowa Department of Transportation\n" +
    "800 Lincoln Way\nAmes, IA 50010").value,
  "Minneapolis, MN");

check("excludes the traveler's own city outright",
  M.findLocation(
    "Ames, IA is where we are.\nThe workshop venue is Omaha, NE.", "Ames").value,
  "Omaha, NE");

check("two-letter non-state is not a location",
  M.findLocation("Contact Miller, XX for details.").value, "");

check("a city name never straddles a line break",
  M.findLocation("Venue: Hilton Garden Inn, Columbus, OH\n800 Lincoln Way\nAmes, IA 50010"),
  { value: "Columbus, OH", alternates: ["Ames, IA"] });
check("a table label is not part of the city",
  M.findLocation("Location\tKansas City, MO").value, "Kansas City, MO");
check("a city whose own name contains 'City' survives",
  M.findLocation("The venue is in Sioux City, IA this year.").value, "Sioux City, IA");
check("Salt Lake City survives too",
  M.findLocation("Venue Salt Lake City, UT").value, "Salt Lake City, UT");
check("a leading preposition is dropped",
  M.findLocation("Conference at Des Moines, IA").value, "Des Moines, IA");

var alts = M.findLocation("Held in Denver, CO. Overflow hotel in Aurora, CO.");
check("offers the runners-up", alts.alternates.indexOf("Aurora, CO") >= 0, true);

// ------------------------------------------------------------------ money

check("labelled registration fee",
  M.findAmount("Registration fee: $675.00", /registration/).value, 675);
check("amount before the label (right-aligned table)",
  M.findAmount("$1,250.00 registration", /registration/).value, 1250);
check("takes the larger of two matches",
  M.findAmount("Early registration $400\nStandard registration $550", /registration/).value, 550);
check("unlabelled money is left alone",
  M.findAmount("Total due $900", /registration/).value, 0);
check("does not cross a line break",
  M.findAmount("Registration\nsomething else entirely $50", /registration/).value, 0);

// ------------------------------------------------------------------- misc

check("speaker role", M.findRole("You are confirmed as a speaker for the panel."), "speaker");
check("plain attendee", M.findRole("Your registration is confirmed."), "attendee");
check("no role stated", M.findRole("Hello there."), "");

check("explicit nights", M.findNights("3 nights at the Hyatt"), 3);
check("nights from the date range", M.findNights("no mention", "2027-01-10", "2027-01-13"), 3);

check("prefers a cued link",
  M.findLink("Unsubscribe: https://x.com/u\nFull agenda: https://edc8.example.org/agenda"),
  "https://edc8.example.org/agenda");

check("date range formatting, same month",
  M.confDatesLine("2027-01-10", "2027-01-14"), "January 10-14, 2027");
check("date range formatting, across months",
  M.confDatesLine("2026-09-29", "2026-10-01"), "September 29 - October 1, 2026");
check("single date formatting", M.confDatesLine("2027-03-03", ""), "March 3, 2027");

// -------------------------------------------------------------------- ics

var ICS = [
  "BEGIN:VCALENDAR", "BEGIN:VEVENT",
  "SUMMARY:EDC-8 Midwest Peer Exchange",
  "DTSTART;VALUE=DATE:20261006",
  "DTEND;VALUE=DATE:20261009",
  "LOCATION:Hyatt Regency\\, Kansas City\\, MO",
  "URL:https://edc8.example.org",
  "END:VEVENT", "END:VCALENDAR",
].join("\r\n");

var ics = M.parseIcs(ICS);
check("ics summary", ics.summary, "EDC-8 Midwest Peer Exchange");
check("ics start", ics.start, "2026-10-06");
check("ics end is made inclusive", ics.end, "2026-10-08");
check("not an ics", M.parseIcs("just some text"), null);

// -------------------------------------------------------------- end to end

var BODY = [
  "<html><body>",
  "<p>Thank you for registering! Your registration was received on August 23, 2026.</p>",
  "<table>",
  "<tr><td>Event</td><td>EDC-8 Midwest Peer Exchange</td></tr>",
  "<tr><td>Conference dates</td><td>October 6-8, 2026</td></tr>",
  "<tr><td>Location</td><td>Hyatt Regency, Kansas City, MO</td></tr>",
  "<tr><td>Registration fee</td><td>$675.00</td></tr>",
  "<tr><td>Room rate</td><td>$189.00 per night</td></tr>",
  "</table>",
  "<p>You are confirmed as a speaker on the Tuesday panel.</p>",
  "<p>Full agenda: https://edc8.example.org/agenda</p>",
  "<p>Please respond by September 1, 2026 if you need accessibility support.</p>",
  "<hr>",
  "<p>Matthew Miller, CPM<br>Iowa Department of Transportation<br>",
  "800 Lincoln Way<br>Ames, IA 50010</p>",
  "</body></html>",
].join("\n");

var pre = M.buildPrefill({
  subject: "Fw: EDC-8 Midwest Peer Exchange Registration Confirmation",
  body: BODY,
  receivedIso: "2026-08-23T12:20:00Z",
  homeCity: "Ames",
  attachments: [],
});

check("e2e event", pre.fields.event, "EDC-8 Midwest Peer Exchange");
check("e2e event start", pre.fields.eventStart, "2026-10-06");
check("e2e conference dates", pre.fields.confDates, "October 6-8, 2026");
check("e2e location, not the signature", pre.fields.location, "Kansas City, MO");
check("e2e registration", pre.fields.cRegistration, 675);
check("e2e lodging rate", pre.fields.cLodgingRate, 189);
check("e2e lodging nights", pre.fields.cLodgingNights, 2);
check("e2e role", pre.fields.attendeeRole, "speaker");
check("e2e link", pre.fields.meetingLink, "https://edc8.example.org/agenda");
check("e2e departure is the day before", pre.fields.departDate, "2026-10-05");
check("e2e return is the day after", pre.fields.returnDate, "2026-10-09");
truthy("e2e reports where the registration fee came from", pre.sources.cRegistration);
check("e2e departure is flagged as assumed",
  /assumed/.test(pre.sources.departDate), true);

// An .ics attachment must win over the body heuristics.
var pre2 = M.buildPrefill({
  subject: "Fw: Registration Confirmation",
  body: "Details are in the attached invite. Some other date: March 3, 2027.",
  receivedIso: "2026-08-23T12:20:00Z",
  attachments: [{ name: "invite.ics", text: ICS }],
});
check("ics wins on event", pre2.fields.event, "EDC-8 Midwest Peer Exchange");
check("ics wins on start", pre2.fields.eventStart, "2026-10-06");
check("ics gives the location", pre2.fields.location, "Kansas City, MO");
check("ics source is named", /invite\.ics/.test(pre2.sources.eventStart), true);

// An attachment fills what the body cannot.
var pre3 = M.buildPrefill({
  subject: "Fw: Peer Exchange Registration Confirmation",
  body: "You're registered. See the attached agenda for details.",
  receivedIso: "2026-08-23T12:20:00Z",
  attachments: [{
    name: "agenda.docx",
    text: "Midwest Peer Exchange Agenda\nDates: October 6-8, 2026\n" +
      "Venue: Hyatt Regency, Kansas City, MO\nRegistration fee\t$675.00",
  }],
});
check("attachment supplies the date", pre3.fields.eventStart, "2026-10-06");
check("attachment supplies the location", pre3.fields.location, "Kansas City, MO");
check("attachment supplies the fee", pre3.fields.cRegistration, 675);
check("attachment is named as the source", /agenda\.docx/.test(pre3.sources.eventStart), true);

// Nothing extractable: say so rather than inventing.
var pre4 = M.buildPrefill({
  subject: "Fw: Confirmation",
  body: "You're all set. Talk soon.",
  receivedIso: "2026-08-23T12:20:00Z",
  attachments: [],
});
check("empty prefill has no date", pre4.fields.eventStart, undefined);
check("empty prefill has no location", pre4.fields.location, undefined);
truthy("empty prefill explains itself", pre4.notes.length >= 3);

// The body must outrank an attachment when both have an amount.
var pre5 = M.buildPrefill({
  subject: "Fw: Conference Registration Confirmation",
  body: "Registration fee: $675.00. Event dates October 6-8, 2026.",
  receivedIso: "2026-08-23T12:20:00Z",
  attachments: [{ name: "policy.pdf", text: "Standard registration fee $2,400.00 for non-members." }],
});
check("the email beats an attachment", pre5.fields.cRegistration, 675);
check("and says it came from the email", /the email/.test(pre5.sources.cRegistration), true);

// ------------------------------------------------------------------ links

check("unwraps a Proofpoint v2 link",
  M.unwrapUrl("https://urldefense.proofpoint.com/v2/url?u=https-3A__www.eventleaf.com_e_edc8midwest&d=DwMGaQ&e="),
  "https://www.eventleaf.com/e/edc8midwest");
check("unwraps Proofpoint hyphen-escaped hyphens",
  M.unwrapUrl("https://urldefense.proofpoint.com/v2/url?u=https-3A__www.hyatt.com_hyatt-2Dregency_en-2DUS_stlrs&d=X"),
  "https://www.hyatt.com/hyatt-regency/en-US/stlrs");
check("unwraps Proofpoint v3",
  M.unwrapUrl("https://urldefense.com/v3/__https://example.org/agenda__;!!abc$"),
  "https://example.org/agenda");
check("unwraps Outlook Safe Links",
  M.unwrapUrl("https://na01.safelinks.protection.outlook.com/?url=https%3A%2F%2Fexample.org%2Fa&data=05"),
  "https://example.org/a");
check("leaves an ordinary URL alone",
  M.unwrapUrl("https://example.org/agenda"), "https://example.org/agenda");

check("finds links that only exist in href attributes",
  M.hrefs('<p>Register <a href="https://ex.org/e/x">Big Summit</a> now</p>')[0].url,
  "https://ex.org/e/x");

var ANCH = [
  { url: "https://www.eventleaf.com/e/edc8midwest", text: "EDC-8 Midwest Peer Exchange", at: 10 },
  { url: "https://www.hyatt.com/events/group-booking/STLRS", text: "Reserve your room", at: 500 },
  { url: "https://forms.gle/73ob3poyFYx5AgMA7", text: "Invitational Traveler form", at: 800 },
  { url: "https://www.google.com/calendar/render?action=TEMPLATE", text: "Google", at: 900 },
  { url: "https://www.eventleaf.com/Attendee/Unsubscribe?cId=x", text: "Unsubscribe", at: 999 },
];
check("picks the event's own page out of a wall of links",
  M.pickLink(ANCH, "EDC-8 Midwest Peer Exchange"),
  "https://www.eventleaf.com/e/edc8midwest");
check("no confident link rather than a wrong one",
  M.pickLink([{ url: "https://ex.org/unsubscribe", text: "Unsubscribe", at: 1 }], "Some Event"), "");

// ------------------------------------------------------------------ times

check("event times", M.findTimes("Time: 8:00 AM - 5:00 PM"), "8am-5pm");
check("times with minutes", M.findTimes("The session runs 9:30 am to 4:15 pm"), "9:30am-4:15pm");
check("ignores the registration desk's hours",
  M.findTimes("Registration opens 7:00 am to 8:00 am.\nThe event runs 8:00 am to 5:00 pm."),
  "8am-5pm");
check("a lone time is not a range", M.findTimes("Doors at 7:00 am."), "");

// --------------------------------------------------------- third-party payer

var INVITE = "Following the event, you may request reimbursement for hotel expenses and " +
  "authorized out-of-pocket expenses, including local transportation, parking, rideshares, " +
  "approved mileage, checked baggage fees, and applicable meal allowances. " +
  "FHWA's authorized travel agent will contact you to arrange and directly pay for approved " +
  "airfare, rail, and/or rental car reservations. " +
  "Questions? Contact Maliha Azmat (maliha.azmat@dot.gov). " +
  "FHWA | Office of Infrastructure | innovation@dot.gov";

var rb = M.findReimbursement(INVITE);
check("spots that someone else is paying", !!rb, true);
check("names the payer", rb.entity, "FHWA");
check("prefers the shared billing address", rb.contact, "innovation@dot.gov");
check("lodging is covered", rb.categories.lodging, true);
check("air/rail/baggage is covered", rb.categories.air, true);
check("meals are covered", rb.categories.meals, true);
check("ground transport is covered", rb.categories.ground, true);
check("registration is not claimed to be covered", rb.categories.reg, false);
check("quotes the sentence that lists what's covered",
  /request reimbursement for hotel expenses/.test(rb.snippet), true);
check("an ordinary confirmation has no third-party payer",
  M.findReimbursement("Your registration is confirmed. See you in June."), null);

// -------------------------------------------------- your own signature

var SIG = [
  "We look forward to seeing you in St. Louis!",
  "Accelerating Innovation Team",
  "FHWA | Office of Infrastructure",
  "U.S. Department of Transportation",
  "",
  "Matthew Miller, CPM",
  "Director of New and Emerging Transportation Technologies",
  "Systems Operations Division",
  "Iowa Department of Transportation",
  "800 Lincoln Way",
  "Ames, IA 50010",
].join("\n");

var sig = M.findSignature(SIG, "Matthew Miller");
check("division from your own signature", sig.division, "Systems Operations");
check("home city from your own signature", sig.city, "Ames");
check("does not adopt the sender's org", sig.bureau, "");
check("no name, no guessing", M.findSignature(SIG, ""), null);
check("a different person's signature is not yours",
  M.findSignature(SIG, "Keri Greenfield"), null);
check("surname-first signatures still match",
  M.findSignature("Miller, Matthew\nTraffic Operations Bureau\n", "Matthew Miller").bureau,
  "Traffic Operations");
check("Bureau of X form",
  M.findSignature("Jane Doe\nBureau of Local Systems\n", "Jane Doe").bureau,
  "Local Systems");

// ------------------------------------------------------- who it goes back to

check("the coordinator named on a travel form",
  M.findCoordinator("Please return this form & meeting link, if available, in email to Keri.Greenfield@iowadot.us"),
  "Keri.Greenfield@iowadot.us");
check("survives the letter-spacing a PDF adds inside an address",
  M.findCoordinator("Please return this form & meeting link, if available, in email to   Keri . Greenfield @iowadot.us"),
  "Keri.Greenfield@iowadot.us");
check("clean text is not mangled by the loose fallback",
  M.findCoordinator("Send it to jane@x.gov. Bob also helps."), "jane@x.gov");
check("ignores a no-reply address",
  M.findCoordinator("Please reply to noreply@eventleaf.com"), "");
check("nothing to offer", M.findCoordinator("Thanks for registering."), "");

// ------------------------------------------------ a filled-in paper form

// Exactly how the real Iowa DOT template extracts out of the PDF, kerning
// artefacts and all ("C ost Center", "Stat e Vehicle").
var BLANK_FORM = "Travel Authorization Form  Please return this form & meeting link, " +
  "if available, in email to   Keri . Greenfield @iowadot.us  " +
  "Name & C ost Center: __________________________________________  " +
  "Other Staff Attending: ________________________________________  " +
  "Name of Conference: ___________________________________________  " +
  "Location: _____________________________________________________  " +
  "Conference Dates & Times: _____________________________________  " +
  "Departure Date: _______________________________________________  " +
  "Return Date: __________________________________________________  " +
  "Reason for Travel (describe event and why you should be attending):  ______  " +
  "Mode of Travel (check ALL that apply):  Personal Vehicle:   Stat e Vehicle:  " +
  "Cost of Travel Mode (miles, estimated flight cost): $ ____________  " +
  "Luggage Fees: $ ________________  Parking: $ _________________  " +
  "Lodging: ________ nights @ $ __________   = $ ____________  " +
  "Registration Fee: $ ________________";

check("a blank template yields nothing at all",
  JSON.stringify(M.formFields(BLANK_FORM)), "{}");

var FILLED = "Travel Authorization Form  " +
  "Name & C ost Center: Matthew Miller, 471-0000  " +
  "Other Staff Attending: Wes Musgrove, Keri Greenfield  " +
  "Name of Conference: EDC-8 Midwest Peer Exchange  " +
  "Location: St. Louis, MO  " +
  "Conference Dates & Times: October 15, 2026, 8am-5pm  " +
  "Departure Date: October 14, 2026  " +
  "Return Date: October 16, 2026  " +
  "Cost of Travel Mode (miles, estimated flight cost): $ 410.00  " +
  "Luggage Fees: $ 70.00  Parking: $ 24.00  " +
  "Lodging: 2 nights @ $ 150.00   = $ 300.00  " +
  "Registration Fee: $ 0.00  Taxi/Uber Fees: $ 55.00  " +
  "Maximum Reimbursement Amount: $ 900.00";

var ff = M.formFields(FILLED);
check("cost center off the form", ff.costCenter, "471-0000");
check("and the name beside it", ff.name, "Matthew Miller");

// Surname-first with no comma before the code is at least as common, and
// splitting on the last comma there makes the cost center "Matthew 300000".
function nameCost(v) {
  return M.formFields(
    "Name & C ost Center: " + v +
    "  Other Staff Attending: none  Name of Conference: X Summit  Location: Ames, IA");
}
check("surname-first, space before the code", nameCost("Miller, Matthew 300000").costCenter, "300000");
check("and the name is kept whole", nameCost("Miller, Matthew 300000").name, "Miller, Matthew");
check("plain name then code", nameCost("Matthew Miller 300000").costCenter, "300000");
check("comma before the code", nameCost("Matthew Miller, 300000").costCenter, "300000");
check("a name with no code at all", nameCost("Matthew Miller").costCenter, undefined);
check("and that name still comes through", nameCost("Matthew Miller").name, "Matthew Miller");
check("other staff", ff.otherStaff, "Wes Musgrove, Keri Greenfield");
check("conference name", ff.event, "EDC-8 Midwest Peer Exchange");
check("location", ff.location, "St. Louis, MO");
check("departure date is parsed to a real date", ff.departDate, "2026-10-14");
check("return date", ff.returnDate, "2026-10-16");
check("travel mode cost", ff.cTravelMode, 410);
check("luggage", ff.cLuggage, 70);
check("parking", ff.cParking, 24);
check("taxi", ff.cTaxi, 55);
check("lodging nights", ff.cLodgingNights, 2);
check("nightly rate", ff.cLodgingRate, 150);
check("max reimbursement", ff.tp1Max, 900);
check("a zero registration fee is not treated as a value", ff.cRegistration, undefined);

// A filled form must beat anything inferred from the covering email's prose.
var pre6 = M.buildPrefill({
  subject: "Fw: Registration Confirmation",
  body: "Registration fee: $675.00. Event dates October 6-8, 2026 in Denver, CO.",
  receivedIso: "2026-08-23T12:20:00Z",
  attachments: [{ name: "Travel Authorization Request.pdf", text: FILLED }],
});
check("the filled form wins on location", pre6.fields.location, "St. Louis, MO");
check("the filled form wins on dates", pre6.fields.departDate, "2026-10-14");
check("and says the form is where it came from",
  /filled-in form/.test(pre6.sources.location), true);
check("the email still supplies what the form left out",
  pre6.fields.cRegistration, 675);

// The form carries TWO third-party entities with identical labels. Reading
// the second one's figures as the first misdirects a receivable.
var TWO_TP = [
  "Registration Fee: $ 450.00 (meals included? How many?) 2B 1L 0D",
  "Additional Fees (and what they are for): $ 85.00 conference banquet ticket",
  "3rd Party Spend Authorization",
  "First 3rd Party Entity",
  "Name: Federal Highway Administration",
  "Project Number: EDC8-IA-2026",
  "Do you have the 3rd party packet to attach? Yes X No ",
  "Maximum Reimbursement Amount: $ 900.00",
  "Reimbursement Items (Select all that apply):",
  "Lodging/Hotel Costs: checked",
  "Meals: checked",
  "Reimbursement Notes: Invitational travel, per diem rate only",
  "Second 3rd Party Entity (if applicable)",
  "Name: Midwest Transportation Consortium",
  "Project Number: MTC-4471",
  "Do you have the 3rd party packet to attach? Yes  No X",
  "Maximum Reimbursement Amount: $ 250.00",
  "Reimbursement Items (Select all that apply):",
  "Ground Transportation Costs: checked",
  "Reimbursement Notes: Covers the rental car only",
].join("\n");

var tp = M.formFields("Name of Conference: EDC-8 Peer Exchange\nLocation: St. Louis, MO\n" +
  "Departure Date: 10/14/2026\n" + TWO_TP);

check("first entity name", tp.tp1Name, "Federal Highway Administration");
check("first project number", tp.tp1Project, "EDC8-IA-2026");
check("first max reimbursement", tp.tp1Max, 900);
check("first entity's packet is marked yes", tp.tp1Packet, true);
check("first entity notes", tp.tp1Notes, "Invitational travel, per diem rate only");
check("first entity lodging", tp.tp1Lodging, true);
check("first entity meals", tp.tp1Meals, true);
check("second entity name", tp.tp2Name, "Midwest Transportation Consortium");
check("second project number", tp.tp2Project, "MTC-4471");
check("second max reimbursement", tp.tp2Max, 250);
check("second entity's packet is not marked yes", tp.tp2Packet, undefined);
check("second entity notes", tp.tp2Notes, "Covers the rental car only");
check("second entity ground transport", tp.tp2Ground, true);
check("the second entity does not inherit the first's categories", tp.tp2Lodging, undefined);

check("breakfasts included", tp.cMealsB, 2);
check("lunches included", tp.cMealsL, 1);
check("zero dinners is not a value", tp.cMealsD, undefined);
check("additional fees amount", tp.cAdditional, 85);
check("and what they were for", tp.cAdditionalDesc, "conference banquet ticket");

// ------------------------------------------------------------------- code

check("registration code", M.findCode("Code: MV59BR3A"), "MV59BR3A");
check("confirmation number", M.findCode("Confirmation number: A8F32K91"), "A8F32K91");
check("a ZIP code is not a confirmation code", M.findCode("Ames, IA 50010"), "");
check("no code present", M.findCode("Thanks for registering."), "");

// ------------------------------------ every field on the real DOT form
//
// The blank template's text, verbatim from the attachment in Matt's mailbox,
// with each printed field filled. This is a coverage test, not a parsing one:
// its job is to fail the day a field silently stops being read.

var BLANK = require("./fixtures-dot-form.js");
var FULL = BLANK
  .replace(/Name & C ost Center: _+/, "Name & C ost Center: Matthew Miller 300000")
  .replace(/Other Staff Attending: _+ _/, "Other Staff Attending: Cedric Wilkinson, Deanne Popp")
  .replace(/Name of Conference: _+/, "Name of Conference: EDC-8 Midwest Peer Exchange")
  .replace(/Location: _+ _+/, "Location: Hyatt Regency 315 Chestnut St, St. Louis, MO 63102")
  .replace(/Conference Dates & Times: _+ _/, "Conference Dates & Times: 10/15/26 7:00 am to 5:00 pm")
  .replace(/Departure Date: _+/, "Departure Date: 10/14/2026")
  .replace(/Return Date: _+ _+/, "Return Date: 10/15/2026")
  .replace(/attending\):  _+ _+  _+ _+/, "attending): Matt is lead for Connected Corridors for Iowa on EDC round 8.")
  .replace(/Stat e Vehicle:  /, "Stat e Vehicle: checked ")
  .replace(/flight cost\): \$ _+/, "flight cost): $ 770 miles")
  .replace(/Luggage Fees: \$ _+/, "Luggage Fees: $ 65")
  .replace(/Parking: \$ _+/, "Parking: $ 40")
  .replace(/Lodging: _+ nights @ \$ _+   = \$ _+/, "Lodging: 1 nights @ $ 176.90 = $ 176.90")
  .replace(/Registration Fee: \$ _+ \(meals included\? How many\?\) _+B_+L_+D/,
           "Registration Fee: $ 250 (meals included? How many?) 1B 1L 0D")
  .replace(/Taxi\/Uber Fees: \$ _+/, "Taxi/Uber Fees: $ 32")
  .replace(/for\): \$ _+ _+ _+  _+ _+/, "for): $ 85 conference banquet ticket")
  .replace(/First   3rd Party Entity  Name: _+/, "First   3rd Party Entity  Name: Federal Highway Administration")
  .replace(/Project Number: _+/, "Project Number: EDC8-IA-2026")
  .replace(/packet to attach\? Yes _+ No _+/, "packet to attach? Yes X No ")
  .replace(/Maximum Reimbursement Amount: \$ _+/, "Maximum Reimbursement Amount: $ 900") +
  "  Reimbursement Items (Select all that apply):  Lodging/Hotel Costs: checked  Meals: checked" +
  "  Reimbursement Notes: Per diem rate only" +
  "  Second   3rd Party Entity (if applicable)  Name: Midwest Transportation Consortium" +
  "  Project Number: MTC-4471  Do you have the 3rd party packet to attach? Yes  No X" +
  "  Maximum Reimbursement Amount: $ 250" +
  "  Reimbursement Items (Select all that apply):  Ground Transportation Costs: checked" +
  "  Reimbursement Notes: Rental car only";

var all = M.formFields(FULL);
[
  ["name", "Matthew Miller"], ["costCenter", "300000"],
  ["otherStaff", "Cedric Wilkinson, Deanne Popp"],
  ["event", "EDC-8 Midwest Peer Exchange"], ["location", "St. Louis, MO"],
  ["confDates", "10/15/26 7:00 am to 5:00 pm"],
  ["departDate", "2026-10-14"], ["returnDate", "2026-10-15"],
  ["modeState", true], ["cTravelMode", 770], ["cLuggage", 65], ["cParking", 40],
  ["cLodgingNights", 1], ["cLodgingRate", 176.9], ["cRegistration", 250],
  ["cMealsB", 1], ["cMealsL", 1], ["cTaxi", 32],
  ["cAdditional", 85], ["cAdditionalDesc", "conference banquet ticket"],
  ["tp1Name", "Federal Highway Administration"], ["tp1Project", "EDC8-IA-2026"],
  ["tp1Packet", true], ["tp1Max", 900], ["tp1Lodging", true], ["tp1Meals", true],
  ["tp1Notes", "Per diem rate only"],
  ["tp2Name", "Midwest Transportation Consortium"], ["tp2Project", "MTC-4471"],
  ["tp2Max", 250], ["tp2Ground", true], ["tp2Notes", "Rental car only"],
].forEach(function (pair) {
  check("real form field: " + pair[0], all[pair[0]], pair[1]);
});
check("reason for travel is read",
  /Connected Corridors/.test(all.reason || ""), true);
check("the same template with nothing filled in still yields nothing",
  JSON.stringify(M.formFields(BLANK)), "{}");

// ------------------------------- teaching it another org's wording

var OTHER_FORM = [
  "County Travel Request",
  "Traveler: Dana Reyes",
  "Acct String: 44-2210",
  "Trip Purpose: Presenting the culvert inventory results",
  "Meeting Name: Midwest Bridge Summit",
  "City/State: Omaha, NE",
  "Leave: 03/09/2027",
  "Back: 03/12/2027",
  "Conference Fee: $ 325.00",
  "Fund Source: Local Roads",
].join("\n");

var disc = M.discoverLabels(OTHER_FORM);
function labelled(list, name) {
  var hit = list.filter(function (x) { return x.label === name; })[0];
  return hit ? hit.field : "(not found)";
}
check("an unfamiliar form still yields its labels", disc.length >= 9, true);
check("wording it happens to share is matched", labelled(disc, "Meeting Name"), "");
check("and nothing is invented for the rest", labelled(disc, "Acct String"), "");
check("a label it does know is pre-assigned",
  labelled(M.discoverLabels("Departure Date: 1/2/2027\nLocation: Ames, IA\nReturn Date: 1/3/2027"),
    "Departure Date"), "departDate");

check("without a mapping, an unfamiliar form gives almost nothing",
  Object.keys(M.formFields(OTHER_FORM)).length < 3, true);

var ALIASES = {
  "Traveler": "name", "Acct String": "costCenter", "Trip Purpose": "reason",
  "Meeting Name": "event", "City/State": "location",
  "Leave": "departDate", "Back": "returnDate", "Conference Fee": "cRegistration",
  "Fund Source": "funding",
};
var mapped = M.formFields(OTHER_FORM, ALIASES);
check("mapped traveler", mapped.name, "Dana Reyes");
check("mapped cost center", mapped.costCenter, "44-2210");
check("mapped reason", mapped.reason, "Presenting the culvert inventory results");
check("mapped event", mapped.event, "Midwest Bridge Summit");
check("mapped location", mapped.location, "Omaha, NE");
check("mapped departure", mapped.departDate, "2027-03-09");
check("mapped return", mapped.returnDate, "2027-03-12");
check("mapped registration fee", mapped.cRegistration, 325);
check("mapped funding", mapped.funding, "Local Roads");

// A saved mapping is shown back as the current choice, not as unassigned.
check("saved choices come back selected",
  labelled(M.discoverLabels(OTHER_FORM, ALIASES), "Acct String"), "costCenter");

// The coordinator's wording wins where it disagrees with the built-in table.
check("an alias overrides the built-in meaning of a shared word",
  M.formFields("Location: Room 240B\nCity/State: Omaha, NE\nDeparture Date: 1/2/2027",
    { "Location": "comments", "City/State": "location" }).comments,
  "Room 240B");

// A tick box reached through an alias arrives as text, not as "checked".
check("an alias onto a tick box reads yes as ticked",
  M.formFields("Use Agency Car: Yes\nMeeting Name: X\nLeave: 1/2/2027",
    { "Use Agency Car": "modeState", "Meeting Name": "event", "Leave": "departDate" }).modeState,
  true);
check("and reads no as not ticked",
  M.formFields("Use Agency Car: No\nMeeting Name: X\nLeave: 1/2/2027",
    { "Use Agency Car": "modeState", "Meeting Name": "event", "Leave": "departDate" }).modeState,
  undefined);

check("every mappable target is a real field id",
  M.MAPPABLE.every(function (m) { return /^[A-Za-z_][A-Za-z0-9]*$/.test(m.field) && m.text; }),
  true);

// ------------------------------------ finding travel in a real-ish inbox
//
// The mailbox is overwhelmingly not travel, so the cost of being generous is
// a list a person has to read and reject. These are the shapes that actually
// sit in one.

var INBOX = [
  { id: "1", subject: "Fw: EDC-8 Midwest Peer Exchange Registration Confirmation",
    bodyPreview: "Thank you for registering. Date: October 15, 2026. Hyatt Regency St. Louis at the Arch. $150 per night.",
    from: "notify@eventleaf.com", receivedDateTime: "2026-08-20T10:00:00Z", hasAttachments: true },
  { id: "2", subject: "Your Roblox One-Time Code", bodyPreview: "Use this one-time code to sign in.",
    from: "no-reply@roblox.com", receivedDateTime: "2026-08-15T10:00:00Z" },
  { id: "3", subject: "AASHTO Spring Meeting - agenda now available",
    bodyPreview: "Join us May 4-6, 2027 in Nashville, TN. Room block at the Omni.",
    from: "events@aashto.org", receivedDateTime: "2026-08-10T10:00:00Z" },
  { id: "4", subject: "Call for Papers: TRB Annual Meeting",
    bodyPreview: "Submit your abstract by September 1, 2026.",
    from: "papers@trb.org", receivedDateTime: "2026-08-05T10:00:00Z" },
  { id: "5", subject: "Lottery Jackpot Fetching", bodyPreview: "You could win big today!",
    from: "promo@luckinnumbers.com", receivedDateTime: "2026-08-16T10:00:00Z" },
  { id: "6", subject: "Concur: Your trip to Washington",
    bodyPreview: "Flight confirmed. Depart March 3, 2027.",
    from: "noreply@concursolutions.com", receivedDateTime: "2026-08-12T10:00:00Z" },
  { id: "7", subject: "Monthly newsletter from the Transportation Institute",
    bodyPreview: "This month: research highlights and a webinar.",
    from: "news@example.org", receivedDateTime: "2026-08-18T10:00:00Z" },
  { id: "8", subject: "Save the date: 2028 Bridge Conference",
    bodyPreview: "More details to follow.", from: "info@bridges.example.org",
    receivedDateTime: "2026-08-19T10:00:00Z" },
];

var OPTS = { todayIso: "2026-08-29", homeCity: "Ames" };
var hits = M.findTravelEmails(INBOX, OPTS);
function has(id) { return hits.some(function (h) { return h.id === id; }); }

check("the real confirmation is found", has("1"), true);
check("a conference with dates and a room block is found", has("3"), true);
check("a booking system's trip email is found", has("6"), true);
check("a one-time passcode is not travel", has("2"), false);
check("a call for papers is not a trip", has("4"), false);
check("spam is not travel", has("5"), false);
check("a newsletter is not travel", has("7"), false);
check("save-the-date is not something to file yet", has("8"), false);
check("the strongest match sorts first", hits[0].id, "1");

check("a trip already filed is marked, not hidden",
  M.findTravelEmails(INBOX, { todayIso: "2026-08-29",
    filed: [{ event: "EDC-8 Midwest Peer Exchange" }] })
    .filter(function (h) { return h.id === "1"; })[0].alreadyFiled, true);
check("and one that is not filed is not marked",
  hits.filter(function (h) { return h.id === "3"; })[0].alreadyFiled, false);

check("every hit says why it is there", hits.every(function (h) { return h.why.length > 0; }), true);
check("the event name is pulled out for the list",
  hits.filter(function (h) { return h.id === "1"; })[0].event,
  "EDC-8 Midwest Peer Exchange");

// An organisation's own booking system is configured, not guessed.
check("an org's own booking sender lifts a plain subject",
  M.scoreTravelEmail({ subject: "Your reservation", bodyPreview: "Confirmed for October 6-8, 2026.",
                       from: "noreply@fleetio.example.gov" },
    { todayIso: "2026-08-29", trustedSenders: ["fleetio.example.gov"] }).score >
  M.scoreTravelEmail({ subject: "Your reservation", bodyPreview: "Confirmed for October 6-8, 2026.",
                       from: "noreply@fleetio.example.gov" }, { todayIso: "2026-08-29" }).score,
  true);

check("an event that has already happened scores lower than one ahead",
  M.scoreTravelEmail({ subject: "Registration confirmation", bodyPreview: "Held June 2, 2025 in Omaha, NE.",
                       from: "x@example.org" }, { todayIso: "2026-08-29" }).score <
  M.scoreTravelEmail({ subject: "Registration confirmation", bodyPreview: "Held June 2, 2027 in Omaha, NE.",
                       from: "x@example.org" }, { todayIso: "2026-08-29" }).score,
  true);

if (failures) {
  console.error("\n" + failures + " mail-parsing test(s) failed.");
  process.exit(1);
}
console.log("All Travel Desk mail-parsing tests passed.");
