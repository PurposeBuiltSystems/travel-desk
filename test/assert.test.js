/*
 * The comparator every other test now depends on.
 *
 * An untested assertion helper is a bad place for a bug: it fails silently in
 * the direction that matters, by agreeing that two different things are the
 * same, and every suite built on it inherits the blind spot.
 *
 * Run: node test/assert.test.js
 */
"use strict";
var A = require("./assert.js");

var failures = 0, passes = 0;
function equal(label, a, b) {
  var d = A.diff(a, b, "");
  if (d) { failures++; console.error("FAIL  " + label + "\n  said they differ: " + d); }
  else { passes++; }
}
function differs(label, a, b, expectPath) {
  var d = A.diff(a, b, "");
  if (!d) { failures++; console.error("FAIL  " + label + "\n  said they are equal, but they are not"); return; }
  if (expectPath && d.indexOf(expectPath) < 0) {
    failures++;
    console.error("FAIL  " + label + "\n  expected the message to point at " + expectPath + "\n  got: " + d);
    return;
  }
  passes++;
}

// --- the two failure modes of the helpers this replaced ---

// === could never match these, however right they were.
equal("two equal arrays", ["a", "b"], ["a", "b"]);
equal("two equal objects", { a: 1 }, { a: 1 });
equal("nested structures", { c: { x: [1, { y: "z" }] } }, { c: { x: [1, { y: "z" }] } });

// JSON.stringify said these differed, because it compares key ORDER.
equal("key order does not matter", { a: 1, b: 2 }, { b: 2, a: 1 });
equal("nor does it nested", { o: { p: 1, q: 2 } }, { o: { q: 2, p: 1 } });

// --- and it still says no when it should ---

differs("a changed leaf", { c: { x: [1, { y: "z" }] } }, { c: { x: [1, { y: "Q" }] } }, ".c.x[1].y");
differs("a missing key", { a: 1 }, { a: 1, b: 2 }, "missing b");
differs("an unexpected key", { a: 1, b: 2 }, { a: 1 }, "unexpected b");
differs("a shorter array", [1, 2], [1, 2, 3]);
differs("a different string", "yes", "no");
differs("a number that is a string", "1", 1);
differs("zero is not false", 0, false);
differs("null is not undefined", null, undefined);
differs("an object is not an array", {}, []);

// JSON.stringify DROPS undefined values, so it called these two the same.
differs("a key holding undefined is not a missing key", { a: undefined }, {});

// --- the awkward values ---

equal("NaN equals NaN", NaN, NaN);
equal("empty objects", {}, {});
equal("empty arrays", [], []);
equal("empty strings", "", "");
differs("NaN is not a number", NaN, 1);

// --- the message has to be usable ---

var msg = A.diff({ costs: { mealsB: "1", mealsD: "0" } }, { costs: { mealsB: "1", mealsD: "2" } }, "");
if (msg && msg.indexOf(".costs.mealsD") >= 0 && msg.indexOf('"2"') >= 0 && msg.indexOf('"0"') >= 0) { passes++; }
else { failures++; console.error("FAIL  the message names the path and both values\n  got: " + msg); }

var deep = A.diff({ a: [{ b: [{ c: 1 }] }] }, { a: [{ b: [{ c: 2 }] }] }, "");
if (deep && deep.indexOf(".a[0].b[0].c") >= 0) { passes++; }
else { failures++; console.error("FAIL  a deep path is reported in full\n  got: " + deep); }

if (failures) {
  console.error("\n" + failures + " assert-helper test(s) failed.");
  process.exit(1);
}
console.log("All " + passes + " assert-helper checks passed.");
