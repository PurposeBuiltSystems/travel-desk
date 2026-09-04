/*
 * The build number in taskpane.js must match the ?v= in taskpane.html.
 *
 * The pane compares the two at startup to detect a page Outlook has cached.
 * If they disagree because somebody bumped one and not the other, every user
 * gets told their page is stale and gets reloaded once for nothing - so the
 * check that finds real staleness has to be right itself.
 */
"use strict";
var fs = require("fs");
var path = require("path");
var root = path.join(__dirname, "..");

var js = fs.readFileSync(path.join(root, "src/taskpane/taskpane.js"), "utf8");
var html = fs.readFileSync(path.join(root, "src/taskpane/taskpane.html"), "utf8");

var declared = /var PANE_BUILD = "(\d+)"/.exec(js);
var asked = html.match(/taskpane\.js\?v=(\d+)/);
var all = (html.match(/\?v=(\d+)/g) || []).map(function (v) { return v.slice(3); });
var distinct = all.filter(function (v, i) { return all.indexOf(v) === i; });

if (!declared) {
  console.error("check-build: taskpane.js has no PANE_BUILD constant.");
  process.exit(1);
}
if (!asked) {
  console.error("check-build: taskpane.html does not load taskpane.js with a ?v=.");
  process.exit(1);
}
if (distinct.length !== 1) {
  console.error("check-build: taskpane.html mixes build numbers: " + distinct.join(", ") +
    " — every asset must be bumped together or the browser caches half of them.");
  process.exit(1);
}
if (declared[1] !== asked[1]) {
  console.error("check-build: taskpane.js says build " + declared[1] +
    " but taskpane.html loads ?v=" + asked[1] + ". They must match, or every " +
    "user is told their page is stale.");
  process.exit(1);
}
console.log("check-build: pane build " + declared[1] + " consistent across html and js.");
