/*
 * Copying a filed request must give back what was filed.
 *
 * model() -> slimModel() -> applyModel() -> model() has to be the identity for
 * every box a person filled in. This is the one place where a quiet omission
 * is invisible: a copied request that is missing a field looks complete, gets
 * signed, and is wrong. So the test drives the REAL taskpane.html, and the two
 * functions are lifted out of the real taskpane.js rather than restated here.
 *
 * Run: node test/copytrip.test.js   (needs jsdom; skips cleanly without it)
 */
"use strict";
var fs = require("fs");
var path = require("path");

var JSDOM;
try { JSDOM = require("jsdom").JSDOM; }
catch (e) {
  console.log("copy-trip tests: jsdom not installed here — skipping.");
  process.exit(0);
}

var root = path.join(__dirname, "..");
var HTML = fs.readFileSync(path.join(root, "src/taskpane/taskpane.html"), "utf8");
var PANE = fs.readFileSync(path.join(root, "src/taskpane/taskpane.js"), "utf8");
var TravelForm = require("../src/form.js");

function lift(name, start) {
  var i = PANE.indexOf(start);
  if (i < 0) { throw new Error("could not find " + name); }
  var depth = 0, j = PANE.indexOf("{", i);
  for (var k = j; k < PANE.length; k++) {
    if (PANE[k] === "{") { depth++; }
    else if (PANE[k] === "}") { depth--; if (!depth) { return PANE.slice(i, k + 1); } }
  }
  throw new Error("unbalanced braces in " + name);
}

var SRC = [
  lift("tp", "function tp(n) {"),
  lift("model", "function model() {"),
  lift("applyModel", "function applyModel(m) {"),
].join("\n");

var dom = new JSDOM(HTML, { url: "https://purposebuiltsystems.github.io/travel-desk/" });
var D = dom.window.document;

var harness = new Function("document", "TravelForm", [
  "function byId(id) { var e = document.getElementById(id); if (!e) { throw new Error('no such element: ' + id); } return e; }",
  "function val(id) { var e = document.getElementById(id); return e ? e.value : ''; }",
  "function setVal(id, v) { var e = document.getElementById(id); if (e) { e.value = v; } }",
  "function setProp(id, k, v) { var e = document.getElementById(id); if (e) { e[k] = v; } }",
  "function setAttrIf(id, n, v) { var e = document.getElementById(id); if (e) { e.setAttribute(n, v); } }",
  "function refreshTotal() {}",
  "function refreshFyLine() {}",
  SRC,
  "return { model: model, applyModel: applyModel };",
].join("\n"))(D, TravelForm);

var failures = 0, passes = 0;
function check(label, actual, expected) {
  var a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a !== e) { failures++; console.error("FAIL  " + label + "\n  expected: " + e + "\n  actual:   " + a); }
  else { passes++; }
}

/* Every box a person could fill, filled with something distinguishable. */
var FILLED = {
  name: "Matthew Miller", costCenter: "300000", division: "Systems Operations",
  bureau: "Traffic Operations", otherStaff: "Cedric Wilkinson\nDeanne Popp",
  event: "EDC-8 Midwest Peer Exchange", location: "St. Louis, MO",
  eventStart: "2027-10-15", attendeeRole: "speaker",
  confDates: "October 15, 2027, 8am-5pm",
  departDate: "2027-10-14", returnDate: "2027-10-16",
  reason: "Lead for Connected Corridors on EDC round 8.",
  meetingLink: "https://www.eventleaf.com/e/edc8midwest",
  comments: "Registration code MV59BR3A", funding: "TEWD",
  modes: { personal: false, state: true, air: false },
  costs: {
    travelMode: "770", luggage: "65", parking: "40", taxi: "32",
    lodgingNights: "2", lodgingRate: "176.9", registration: "250",
    additional: "85", additionalDesc: "conference banquet ticket",
    mealsB: "1", mealsL: "1", mealsD: "0",
  },
  thirdParties: [
    { name: "Federal Highway Administration", contact: "innovation@dot.gov",
      project: "EDC8-IA-2026", packet: true, maxReimb: "900",
      items: { registration: false, lodging: true, airfare: true, meals: true, ground: true },
      notes: "Invitational travel, per diem rate only" },
    { name: "Midwest Transportation Consortium", contact: "billing@mtc.example.org",
      project: "MTC-4471", packet: false, maxReimb: "250",
      items: { registration: false, lodging: false, airfare: false, meals: false, ground: true },
      notes: "Rental car only" },
  ],
};

harness.applyModel(FILLED);
var read = harness.model();

Object.keys(FILLED).forEach(function (k) {
  if (k === "modes" || k === "costs" || k === "thirdParties") { return; }
  check("survives a copy: " + k, read[k], FILLED[k]);
});
check("travel modes survive", read.modes, FILLED.modes);
check("every cost line survives", read.costs, FILLED.costs);
check("both third parties survive", read.thirdParties, FILLED.thirdParties);

/* The tick boxes whose stored key differs from their element id — the two
   that a lowercase-the-id shortcut drops without a word. */
check("3rd party 'registration' tick is not confused with the tp1Reg box",
  read.thirdParties[0].items.registration, false);
check("3rd party 'airfare' tick reaches the tp1Air box",
  read.thirdParties[0].items.airfare, true);

/* Through storage: slimming must not lose anything that was filled in. */
var stored = TravelForm.slimModel(FILLED);
var dom2 = new JSDOM(HTML, { url: "https://x/" });
var h2 = new Function("document", "TravelForm", [
  "function byId(id) { var e = document.getElementById(id); if (!e) { throw new Error('no such element: ' + id); } return e; }",
  "function val(id) { var e = document.getElementById(id); return e ? e.value : ''; }",
  "function setVal(id, v) { var e = document.getElementById(id); if (e) { e.value = v; } }",
  "function setProp(id, k, v) { var e = document.getElementById(id); if (e) { e[k] = v; } }",
  "function setAttrIf(id, n, v) { var e = document.getElementById(id); if (e) { e.setAttribute(n, v); } }",
  "function refreshTotal() {}", "function refreshFyLine() {}",
  SRC, "return { model: model, applyModel: applyModel };",
].join("\n"))(dom2.window.document, TravelForm);

h2.applyModel(stored);
var afterStorage = h2.model();
check("nothing filled in is lost on the way through storage",
  afterStorage, read);
check("an untouched meal count comes back as the page's own default, not blank",
  afterStorage.costs.mealsD, "0");

/* A request from last year cannot keep last year's dates. */
var old = TravelForm.copyForNewTrip(FILLED, "2028-01-05");
check("past dates are cleared", old.dropped.sort(),
  ["confDates", "departDate", "eventStart", "returnDate"]);
check("and everything else is kept", old.model.event, FILLED.event);
check("costs are kept", old.model.costs.registration, "250");

var future = TravelForm.copyForNewTrip(FILLED, "2027-01-05");
check("dates still ahead are kept", future.dropped, []);
check("as are the conference dates", future.model.confDates, FILLED.confDates);

/* Slimming is what makes forty of these fit; it must not keep noise. */
check("empty boxes are dropped", TravelForm.slimModel({ a: "x", b: "", c: null }), { a: "x" });
check("a typed zero cost is carried forward, because it is an answer",
  TravelForm.slimModel({ costs: { parking: "0", taxi: "40" } }),
  { costs: { parking: "0", taxi: "40" } });
check("but an untouched meal count is not - that box ships showing 0",
  TravelForm.slimModel({ costs: { mealsB: "0", mealsL: "2" } }), { costs: { mealsL: "2" } });
check("an unticked box is dropped", TravelForm.slimModel({ modes: { air: false, state: true } }),
  { modes: { state: true } });
check("a third party with no name is dropped",
  TravelForm.slimModel({ thirdParties: [{ name: "" }, { name: "FHWA" }] }),
  { thirdParties: [{ name: "FHWA" }] });
check("slimming a filled request is a real saving",
  TravelForm.slimModel(FILLED).costs.mealsD, undefined);

if (failures) {
  console.error("\n" + failures + " copy-trip test(s) failed.");
  process.exit(1);
}
console.log("All " + passes + " copy-trip checks passed against the real pane markup.");
