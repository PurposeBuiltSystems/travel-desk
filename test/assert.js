/*
 * One assertion helper for the whole suite.
 *
 * There were five copies of check(), and they did not agree. Two compared with
 * ===, so any assertion on an array or an object could never pass however
 * right it was - and failed while printing an expected and an actual that
 * looked identical, which is worse than failing plainly. The other three
 * compared JSON.stringify output, which is closer but wrong in the other
 * direction: it is sensitive to key ORDER, so two equal objects built in a
 * different sequence are reported as different.
 *
 * This does structural comparison, and when it fails it names the PATH to the
 * first difference rather than printing two blobs and leaving the reader to
 * find it.
 */
"use strict";

function typeOf(v) {
  if (v === null) { return "null"; }
  if (Array.isArray(v)) { return "array"; }
  return typeof v;
}

function show(v) {
  if (typeof v === "string") { return JSON.stringify(v); }
  if (v === undefined) { return "undefined"; }
  if (typeof v === "object" && v !== null) {
    try { return JSON.stringify(v); } catch (e) { return String(v); }
  }
  return String(v);
}

/**
 * null when equal; otherwise a sentence saying where they diverge.
 *
 * NaN equals NaN here. Every other comparison in a test suite treats two NaNs
 * as the same value, and === is alone in disagreeing.
 */
function diff(a, b, path) {
  var at = path || "";
  var ta = typeOf(a), tb = typeOf(b);

  if (ta === "number" && tb === "number" && isNaN(a) && isNaN(b)) { return null; }
  if (ta !== tb) {
    return "at " + (at || "the value") + ": expected " + tb + " " + show(b) +
      ", got " + ta + " " + show(a);
  }

  if (ta === "array") {
    if (a.length !== b.length) {
      return "at " + (at || "the value") + ": expected " + b.length +
        " item(s), got " + a.length + " — " + show(a);
    }
    for (var i = 0; i < a.length; i++) {
      var d = diff(a[i], b[i], at + "[" + i + "]");
      if (d) { return d; }
    }
    return null;
  }

  if (ta === "object") {
    var ka = Object.keys(a).sort(), kb = Object.keys(b).sort();
    var missing = kb.filter(function (k) { return ka.indexOf(k) < 0; });
    var extra = ka.filter(function (k) { return kb.indexOf(k) < 0; });
    if (missing.length) {
      return "at " + (at || "the object") + ": missing " + missing.join(", ");
    }
    if (extra.length) {
      return "at " + (at || "the object") + ": unexpected " + extra.join(", ");
    }
    for (var j = 0; j < kb.length; j++) {
      var key = kb[j];
      var dd = diff(a[key], b[key], at + "." + key);
      if (dd) { return dd; }
    }
    return null;
  }

  if (a !== b) {
    return "at " + (at || "the value") + ": expected " + show(b) + ", got " + show(a);
  }
  return null;
}

/**
 * A counting check() bound to one test file.
 *
 *   var t = require("./assert.js").suite("form");
 *   t.check("label", actual, expected);
 *   t.done();
 */
function suite(name) {
  var failures = 0, passes = 0;
  return {
    check: function (label, actual, expected) {
      var d = diff(actual, expected, "");
      if (d) { failures++; console.error("FAIL  " + label + "\n  " + d); }
      else { passes++; }
    },
    /** Substring assertion - for extracted text, where an exact match is not
     *  the point and would break on a single stray space. */
    has: function (label, hay, needle) {
      if (String(hay).indexOf(needle) >= 0) { passes++; return; }
      failures++;
      console.error("FAIL  " + label + "\n  expected to contain " + show(needle) +
        "\n  got " + show(String(hay).slice(0, 300)));
    },
    truthy: function (label, v) {
      if (v) { passes++; return; }
      failures++;
      console.error("FAIL  " + label + "\n  expected something truthy, got " + show(v));
    },
    done: function (line) {
      if (failures) {
        console.error("\n" + failures + " " + name + " test(s) failed.");
        process.exit(1);
      }
      console.log(line || ("All " + passes + " " + name + " checks passed."));
    },
    get failures() { return failures; },
    get passes() { return passes; },
  };
}

module.exports = { diff: diff, suite: suite };
