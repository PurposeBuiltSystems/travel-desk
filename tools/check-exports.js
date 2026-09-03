/*
 * Every Module.method() call must resolve to something the module exports.
 *
 * check-undeclared reads the AST for undeclared IDENTIFIERS, and a wrong
 * PROPERTY name is not one - TravelForm.extractEmails parses perfectly and
 * throws only when the line runs. That is how "Write the invitation" shipped
 * broken: the very first statement of sendInvites called a function that had
 * never existed, so the button failed for everyone, every time, and no test
 * touched it because the failure needed a browser.
 *
 * Usage: node tools/check-exports.js
 */
"use strict";
var fs = require("fs");
var path = require("path");

var ROOT = path.join(__dirname, "..");
var MODULES = {
  TravelForm: "src/form.js",
  GraphData: "src/graph.js",
  TdMail: "src/mailparse.js",
  TdAttach: "src/attachtext.js",
  TravelCoord: "src/coord.js",
  XlsxGen: "src/xlsxgen.js",
};

var exports_ = {};
Object.keys(MODULES).forEach(function (name) {
  var mod;
  try { mod = require(path.join(ROOT, MODULES[name])); }
  catch (e) {
    console.error("could not load " + MODULES[name] + ": " + e.message);
    process.exit(1);
  }
  // graph.js hands back { GraphData: api }; the others export the api itself.
  exports_[name] = Object.keys(mod[name] || mod);
});

var files = [];
["src", path.join("src", "taskpane"), path.join("src", "commands")].forEach(function (dir) {
  var full = path.join(ROOT, dir);
  var names = [];
  try { names = fs.readdirSync(full); } catch (e) { return; }
  names.filter(function (n) { return n.endsWith(".js"); })
       .forEach(function (n) { files.push(path.join(dir, n)); });
});

var bad = [], seen = {};
files.forEach(function (rel) {
  var src = fs.readFileSync(path.join(ROOT, rel), "utf8");
  var rx = new RegExp("\\b(" + Object.keys(MODULES).join("|") + ")\\.([A-Za-z_$][\\w$]*)", "g");
  var m;
  while ((m = rx.exec(src))) {
    var key = m[1] + "." + m[2] + "|" + rel;
    if (seen[key]) { continue; }
    seen[key] = 1;
    if (exports_[m[1]].indexOf(m[2]) < 0) {
      var line = src.slice(0, m.index).split("\n").length;
      bad.push("  " + rel + ":" + line + "  " + m[1] + "." + m[2] +
        " is not exported by " + MODULES[m[1]]);
    }
  }
});

if (bad.length) {
  console.error("\n" + bad.join("\n"));
  console.error("\n" + bad.length + " call(s) to something that does not exist — these throw when the line runs.");
  process.exit(1);
}
var total = Object.keys(seen).length;
console.log("check-exports: all " + total + " cross-module call(s) resolve.");
