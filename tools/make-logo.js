/*
 * Regenerate src/logo.js from assets/icon-32.png.
 *
 * The logo is embedded in source so the generated email can carry it as an
 * inline attachment without fetching anything. Run this if the icon changes.
 */
"use strict";
var fs = require("fs");
var b = fs.readFileSync("assets/icon-32.png").toString("base64");
var src = fs.readFileSync("src/logo.js", "utf8");
var lines = b.match(/.{1,92}/g).map(function (l) { return '    "' + l + '" +'; });
lines[lines.length - 1] = lines[lines.length - 1].replace(/ \+$/, "");
var out = src.replace(/(\/\*\* base64 PNG, 32x32 \*\/\n    BYTES:\n)[\s\S]*?(,\n    CONTENT_TYPE)/,
  "$1" + lines.join("\n") + "$2");
fs.writeFileSync("src/logo.js", out);
console.log("src/logo.js regenerated from assets/icon-32.png (" + b.length + " chars)");
