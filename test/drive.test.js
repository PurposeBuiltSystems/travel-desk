/*
 * What the user is told when their OneDrive has never been opened.
 *
 * This is not hypothetical. A licensed Microsoft 365 account that has never
 * loaded OneDrive has no drive at all: GET /me/drive returns 404
 * itemNotFound and /me/drives returns an empty list. Confirmed against a
 * real account on 14 September 2026, which is what prompted this test.
 *
 * The failure mode being guarded against is a message that is technically
 * true and practically useless. uploadWorkbook already explained the 404
 * well - but only when it was the first call to hit it. Give the planner a
 * folder and ensureFolder ran first and said "Couldn't create the folder
 * 'Travel' (404)", which sends the reader looking for a permissions problem
 * that isn't there.
 *
 * Run: npm test
 */
"use strict";
var fs = require("fs");
var path = require("path");
var t = require("./assert.js").suite("onedrive");

var GRAPHJS = fs.readFileSync(path.join(__dirname, "..", "src", "graph.js"), "utf8");

/**
 * Load the real graph.js with a fetch that answers like a given account.
 * `routes` maps a URL fragment to {status, body}.
 */
function loadGraph(routes) {
  var calls = [];
  var self = {};
  var Office = {
    context: {
      roamingSettings: { get: function () { return false; }, set: function () {}, saveAsync: function (cb) { cb(); } },
      mailbox: { diagnostics: { hostName: "Outlook" } },
    },
  };
  var msal = { createNestablePublicClientApplication: function () { return Promise.resolve({}); } };

  function fakeFetch(url, opts) {
    calls.push(((opts && opts.method) || "GET") + " " + String(url).replace(/^https:\/\/graph\.microsoft\.com\/v1\.0/, ""));
    var hit = Object.keys(routes).find(function (frag) { return String(url).indexOf(frag) >= 0; });
    var r = hit ? routes[hit] : { status: 200, body: "{}" };
    return Promise.resolve({
      ok: r.status >= 200 && r.status < 300,
      status: r.status,
      text: function () { return Promise.resolve(r.body || ""); },
      json: function () { return Promise.resolve(JSON.parse(r.body || "{}")); },
      headers: { get: function () { return null; } },
    });
  }

  var factory = new Function("self", "Office", "msal", "fetch", "setTimeout", "clearTimeout",
    GRAPHJS + "\nreturn self.GraphData;");
  var api = factory(self, Office, msal, fakeFetch,
                    function (fn) { return fn(); }, function () {});
  return { api: api, calls: calls };
}

var NO_DRIVE = { "/me/drive?$select=id": { status: 404, body: '{"error":{"code":"itemNotFound"}}' } };

(async function () {
  // --- no OneDrive, and the planner is going in a folder -------------------
  var g = loadGraph(NO_DRIVE);
  var msg = "";
  try {
    await g.api.uploadWorkbook("tok", "Planner.xlsx", new Uint8Array(4), "Division/Travel");
  } catch (e) { msg = e.message; }

  t.has("names OneDrive as the problem", msg, "OneDrive isn't set up");
  t.has("tells the user what to do", msg, "onedrive.com");
  t.check("does not blame the folder", msg.indexOf("folder") >= 0, false);
  t.check("checked the drive before creating anything",
    g.calls[0], "GET /me/drive?$select=id");
  t.check("never attempted the folder", g.calls.some(function (c) {
    return c.indexOf("children") >= 0; }), false);
  t.check("never attempted the upload", g.calls.some(function (c) {
    return c.indexOf(":/content") >= 0; }), false);

  // --- no OneDrive, no folder: the same message, not a different one -------
  var g2 = loadGraph(NO_DRIVE);
  var msg2 = "";
  try { await g2.api.uploadWorkbook("tok", "Planner.xlsx", new Uint8Array(4), ""); }
  catch (e) { msg2 = e.message; }
  t.check("one message whether or not a folder was chosen", msg2, msg);

  // --- a normal account still works ---------------------------------------
  var ok = loadGraph({
    "/me/drive?$select=id": { status: 200, body: '{"id":"drive1"}' },
    "children": { status: 201, body: '{"id":"f1","name":"Travel"}' },
    ":/content": { status: 201, body: '{"id":"wb1","name":"Planner.xlsx","webUrl":"https://x/Planner.xlsx"}' },
  });
  var made = await ok.api.uploadWorkbook("tok", "Planner.xlsx", new Uint8Array(4), "Division/Travel");
  t.check("a provisioned account still gets its workbook", made && made.name, "Planner.xlsx");
  t.truthy("and a reference the planner can be reopened from", made && made.ref && made.ref.itemId);
  t.check("and the drive check did not replace the upload",
    ok.calls.some(function (c) { return c.indexOf(":/content") >= 0; }), true);

  // --- a real failure is not disguised as a provisioning problem -----------
  var boom = loadGraph({ "/me/drive?$select=id": { status: 500, body: "server error" } });
  var msg3 = "";
  try { await boom.api.uploadWorkbook("tok", "P.xlsx", new Uint8Array(4), ""); }
  catch (e) { msg3 = e.message; }
  t.has("a 500 is reported as a 500", msg3, "500");
  t.check("and is not called a setup problem", msg3.indexOf("isn't set up") >= 0, false);

  t.done();
})();
