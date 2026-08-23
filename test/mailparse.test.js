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

// ------------------------------------------------------------------- code

check("registration code", M.findCode("Code: MV59BR3A"), "MV59BR3A");
check("confirmation number", M.findCode("Confirmation number: A8F32K91"), "A8F32K91");
check("a ZIP code is not a confirmation code", M.findCode("Ames, IA 50010"), "");
check("no code present", M.findCode("Thanks for registering."), "");

if (failures) {
  console.error("\n" + failures + " mail-parsing test(s) failed.");
  process.exit(1);
}
console.log("All Travel Desk mail-parsing tests passed.");
