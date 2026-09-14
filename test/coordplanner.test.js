/*
 * Which workbook the coordinator view reads.
 *
 * Reported from real use: "Load everyone's travel" did nothing. The cause
 * was one line -
 *
 *     var ref = wbRef || (st.wbUrl ? null : null);
 *
 * a ternary whose branches are identical, so it always collapsed to the
 * in-memory ref. That ref is only set by connecting a workbook in the
 * current session, so after a reload the button had nothing; and for anyone
 * set up with per-year planners st.wbUrl is empty, so the fallback had
 * nothing either. st.wbRef and st.planners - where the planner actually
 * lives - were never consulted.
 *
 * A resolution order is exactly the sort of thing unit tests skip because it
 * looks like plumbing. This drives the REAL coordPlanner out of taskpane.js.
 *
 * Run: npm test
 */
"use strict";
var fs = require("fs");
var path = require("path");
var t = require("./assert.js").suite("coord-planner");

var PANE = fs.readFileSync(path.join(__dirname, "..", "src", "taskpane", "taskpane.js"), "utf8");
var TravelForm = require("../src/form.js");

function lift(start) {
  var i = PANE.indexOf(start);
  if (i < 0) { throw new Error("not found: " + start); }
  var d = 0;
  for (var k = PANE.indexOf("{", i); k < PANE.length; k++) {
    if (PANE[k] === "{") { d++; }
    else if (PANE[k] === "}") { d--; if (!d) { return PANE.slice(i, k + 1); } }
  }
  throw new Error("unbalanced: " + start);
}

/** The real coordPlanner, with settings and Graph supplied. */
function makeResolver(st, sessionRef) {
  var resolved = [];
  var factory = new Function("settings", "wbRef", "TravelForm", "GraphData", "setStatus", [
    lift("async function coordPlanner(token, fy) {"),
    "return coordPlanner;",
  ].join("\n"));
  return {
    fn: factory(function () { return st; }, sessionRef || null, TravelForm, {
      resolveWorkbook: function (tok, url) {
        resolved.push(url);
        return Promise.resolve({ driveId: "d", itemId: "from-url", name: url });
      },
    }, function () {}),
    resolved: resolved,
  };
}

var REF = function (id) { return { driveId: "d", itemId: id, name: id }; };

(async function () {
  // --- the reported case: per-year planners, nothing in session ----------
  var r1 = makeResolver({
    planners: { "FY27": { wbRef: REF("fy27"), tableName: "Planner" } },
  }, null);
  var got = await r1.fn("tok", "FY27");
  t.check("the year's planner is found", got.ref.itemId, "fy27");
  t.check("with its own table name", got.tableName, "Planner");

  // --- the same planner, but the view is not filtered by year ------------
  var r2 = makeResolver({
    planners: { "FY27": { wbRef: REF("fy27"), tableName: "Planner" } },
  }, null);
  var got2 = await r2.fn("tok", "");
  t.check("an unfiltered view still finds it", got2.ref.itemId, "fy27");

  // --- a catch-all planner ------------------------------------------------
  var r3 = makeResolver({
    planners: { "*": { wbRef: REF("all"), tableName: "T" } },
  }, null);
  t.check("the catch-all is used for any year", (await r3.fn("tok", "FY29")).ref.itemId, "all");

  // --- the legacy single-planner setup ------------------------------------
  var r4 = makeResolver({ wbRef: REF("legacy"), tableName: "T" }, null);
  t.check("a saved single planner is used", (await r4.fn("tok", "")).ref.itemId, "legacy");
  t.check("and nothing is resolved over the network", r4.resolved.length, 0);

  // --- only a URL saved: resolve it, as before ----------------------------
  var r5 = makeResolver({ wbUrl: "https://x/planner.xlsx", tableName: "T" }, null);
  t.check("a URL-only setup still resolves",
    (await r5.fn("tok", "")).ref.itemId, "from-url");
  t.check("and it resolved the saved link", r5.resolved[0], "https://x/planner.xlsx");

  // --- a workbook connected in this session wins over a stale URL ---------
  var r6 = makeResolver({ wbUrl: "https://x/old.xlsx", tableName: "T" }, REF("session"));
  t.check("the session's workbook is preferred",
    (await r6.fn("tok", "")).ref.itemId, "session");

  // --- a year with no planner falls back rather than reporting nothing ----
  var r7 = makeResolver({
    planners: { "FY26": { wbRef: REF("fy26"), tableName: "T" } },
  }, null);
  t.check("an unmatched year still opens a planner",
    (await r7.fn("tok", "FY99")).ref.itemId, "fy26");

  // --- genuinely nothing connected ---------------------------------------
  var r8 = makeResolver({}, null);
  t.check("nothing connected returns no ref", (await r8.fn("tok", "")).ref, null);

  t.done();
})();
