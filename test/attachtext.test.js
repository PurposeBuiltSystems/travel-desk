/*
 * Offline tests for attachment text extraction. Run: node test/attachtext.test.js
 *
 * The fixtures are built here rather than committed as binaries so that what
 * each format actually contains is visible in the test, and so a change to the
 * extractor is checked against a real zip/PDF rather than a recorded string.
 */
"use strict";
global.JSZip = require("jszip");
var zlib = require("zlib");
var A = require("../src/attachtext.js");

var T = require("./assert.js").suite("attachment");
var check = T.check, truthy = T.truthy, has = T.has;

function bytes(buf) { return Uint8Array.from(buf); }
function latin1FromU8(u8) { return Buffer.from(u8).toString("latin1"); }

// ------------------------------------------------------------------- plain

(async function () {

  var ics = await A.attachmentText("invite.ics",
    bytes(Buffer.from("BEGIN:VCALENDAR\r\nSUMMARY:Peer Exchange\r\nEND:VCALENDAR", "utf8")));
  has("ics is read as text", ics.text, "SUMMARY:Peer Exchange");

  var txt = await A.attachmentText("notes.txt", bytes(Buffer.from("Kansas City, MO", "utf8")));
  check("txt", txt.text, "Kansas City, MO");

  var png = await A.attachmentText("logo.png", bytes(Buffer.from([137, 80, 78, 71])));
  check("an image yields no text", png.text, "");
  has("and says why", png.note, "image");

  var unknown = await A.attachmentText("thing.xyz", bytes(Buffer.from("hello")));
  has("an unknown type is reported, not silently dropped", unknown.note, "skipped");

  var empty = await A.attachmentText("empty.pdf", new Uint8Array(0));
  check("empty attachment", empty.text, "");

  // -------------------------------------------------------------- docx

  var docx = new global.JSZip();
  docx.file("word/document.xml",
    '<?xml version="1.0"?><w:document xmlns:w="x"><w:body>' +
    "<w:p><w:r><w:t>Midwest Peer Exchange Agenda</w:t></w:r></w:p>" +
    "<w:p><w:r><w:t>Dates: October 6-8, 2026</w:t></w:r></w:p>" +
    "<w:tbl><w:tr>" +
    "<w:tc><w:p><w:r><w:t>Registration fee</w:t></w:r></w:p></w:tc>" +
    "<w:tc><w:p><w:r><w:t>$675.00</w:t></w:r></w:p></w:tc>" +
    "</w:tr></w:tbl>" +
    "</w:body></w:document>");
  var docxBytes = await docx.generateAsync({ type: "uint8array" });
  var d = await A.attachmentText("agenda.docx", docxBytes);
  has("docx paragraph", d.text, "Midwest Peer Exchange Agenda");
  has("docx dates", d.text, "October 6-8, 2026");
  has("docx keeps a table row on one line", d.text.replace(/\n/g, "|"), "Registration fee");

  // The point of preserving line breaks: the money extractor must not read
  // across a paragraph boundary and pair a label with an unrelated number.
  var Mp = require("../src/mailparse.js");
  check("docx text still yields the fee",
    Mp.findAmount(d.text, /registration/).value, 675);

  // -------------------------------------------------------------- xlsx

  var xlsx = new global.JSZip();
  xlsx.file("xl/sharedStrings.xml",
    '<?xml version="1.0"?><sst><si><t>Event</t></si><si><t>EDC-8 Peer Exchange</t></si>' +
    "<si><t>Registration</t></si><si><t>Location</t></si><si><t>Kansas City, MO</t></si></sst>");
  xlsx.file("xl/worksheets/sheet1.xml",
    '<?xml version="1.0"?><worksheet><sheetData>' +
    '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>' +
    '<row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>675</v></c></row>' +
    '<row r="3"><c r="A3" t="s"><v>3</v></c><c r="B3" t="s"><v>4</v></c></row>' +
    "</sheetData></worksheet>");
  var xlsxBytes = await xlsx.generateAsync({ type: "uint8array" });
  var x = await A.attachmentText("costs.xlsx", xlsxBytes);
  has("xlsx shared string", x.text, "EDC-8 Peer Exchange");
  has("xlsx keeps a row together", x.text, "Registration\t675");
  has("xlsx location", x.text, "Kansas City, MO");
  check("a location can be read straight out of a spreadsheet",
    Mp.findLocation(x.text).value, "Kansas City, MO");

  // ---------------------------------------------------------------- pdf

  function makePdf(lines, compress) {
    var content = "BT /F1 12 Tf 72 720 Td\n" +
      lines.map(function (l) {
        return "(" + l.replace(/([()\\])/g, "\\$1") + ") Tj 0 -16 Td";
      }).join("\n") + "\nET";
    var streamBytes = compress
      ? zlib.deflateSync(Buffer.from(content, "latin1"))
      : Buffer.from(content, "latin1");
    var head = "%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n" +
      "4 0 obj<</Type/Page/Length " + streamBytes.length +
      (compress ? "/Filter/FlateDecode" : "") + ">>stream\n";
    return bytes(Buffer.concat([
      Buffer.from(head, "latin1"), streamBytes,
      Buffer.from("\nendstream endobj\n%%EOF", "latin1"),
    ]));
  }

  var flat = await A.attachmentText("agenda.pdf", makePdf([
    "EDC-8 Midwest Peer Exchange",
    "October 6-8, 2026",
    "Hyatt Regency, Kansas City, MO",
    "Registration fee: $675.00 per attendee",
    "Room rate $189.00 per night for three nights",
  ], false));
  has("uncompressed pdf text", flat.text, "EDC-8 Midwest Peer Exchange");
  has("pdf dates", flat.text, "October 6-8, 2026");

  var zipped = await A.attachmentText("agenda-flate.pdf", makePdf([
    "EDC-8 Midwest Peer Exchange",
    "October 6-8, 2026",
    "Hyatt Regency, Kansas City, MO",
    "Registration fee: $675.00 per attendee",
    "Room rate $189.00 per night for three nights",
  ], true));
  has("FlateDecode pdf is inflated", zipped.text, "Kansas City, MO");
  check("and the fee comes out of it",
    Mp.findAmount(zipped.text, /registration/).value, 675);
  check("and the dates come out of it",
    Mp.eventDates(zipped.text, "2026-08-23T12:00:00Z"),
    { start: "2026-10-06", end: "2026-10-08" });

  // Escapes and hex strings
  var esc = await A.attachmentText("esc.pdf", makePdf([
    "Registration (early bird) fee: $450.00 confirmed for the annual meeting",
    "Held in Des Moines, Iowa on May 4-6, 2027 at the downtown venue",
  ], true));
  has("escaped parens survive", esc.text, "(early bird)");

  // Garbage in, nothing out: a subset-font PDF decodes to glyph indices, and
  // handing those to the extractors is how a wrong date reaches a signed form.
  var junk = "BT /F1 12 Tf (" +
    Array.from({ length: 300 }, function (_, i) {
      return String.fromCharCode(1 + (i % 25));
    }).join("").replace(/([()\\])/g, "\\$1") + ") Tj ET";
  var junkPdf = bytes(Buffer.concat([
    Buffer.from("%PDF-1.4\n4 0 obj<</Type/Page/Length " + junk.length + ">>stream\n", "latin1"),
    Buffer.from(junk, "latin1"),
    Buffer.from("\nendstream endobj\n%%EOF", "latin1"),
  ]));
  var j = await A.attachmentText("scanned.pdf", junkPdf);
  check("glyph-index garbage is discarded", j.text, "");
  has("and reported as unreadable", j.note, "no readable text");

  // A subset CID font with /Encoding/Identity-H and a /ToUnicode CMap — what
  // Word and PDFsharp actually produce, and what the real FHWA invitational
  // travel guidelines attachment turned out to be. Without following the CMap
  // this decodes to glyph numbers and gets thrown away.
  function makeCidPdf(lines) {
    var text = lines.join("\n");
    var chars = [], seen = {};
    text.split("").forEach(function (ch) {
      if (ch === "\n" || seen[ch]) { return; }
      seen[ch] = chars.length + 1;       // glyph ids start at 1
      chars.push(ch);
    });
    var bf = chars.map(function (ch, i) {
      return "<" + ("000" + (i + 1).toString(16)).slice(-4) + "> <" +
        ("000" + ch.charCodeAt(0).toString(16)).slice(-4) + ">";
    }).join("\n");
    var cmap = "/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n" +
      "1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n" +
      chars.length + " beginbfchar\n" + bf + "\nendbfchar\nendcmap\nend\nend";

    var content = "BT /F0 12 Tf 72 720 Td\n" + lines.map(function (l) {
      var hex = l.split("").map(function (ch) {
        return ("000" + seen[ch].toString(16)).slice(-4);
      }).join("");
      return "<" + hex + "> Tj 0 -16 Td";
    }).join("\n") + "\nET";

    var cmapZ = zlib.deflateSync(Buffer.from(cmap, "latin1"));
    var contZ = zlib.deflateSync(Buffer.from(content, "latin1"));
    var parts = [Buffer.from(
      "%PDF-1.4\n" +
      "1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n" +
      "2 0 obj<</Type/Pages/Count 1/Kids[3 0 R]>>endobj\n" +
      "3 0 obj<</Type/Page/Parent 2 0 R/Resources<</Font<</F0 4 0 R>>>>/Contents 6 0 R>>endobj\n" +
      "4 0 obj<</Type/Font/Subtype/Type0/BaseFont/BXNAPB+Calibri/Encoding/Identity-H/ToUnicode 5 0 R>>endobj\n" +
      "5 0 obj<</Length " + cmapZ.length + "/Filter/FlateDecode>>stream\n", "latin1")];
    parts.push(cmapZ);
    parts.push(Buffer.from("\nendstream endobj\n6 0 obj<</Length " + contZ.length +
      "/Filter/FlateDecode>>stream\n", "latin1"));
    parts.push(contZ);
    parts.push(Buffer.from("\nendstream endobj\n%%EOF", "latin1"));
    return bytes(Buffer.concat(parts));
  }

  var cid = await A.attachmentText("guidelines.pdf", makeCidPdf([
    "Instructions for FHWA Invitational Travelers",
    "The peer exchange is held October 15, 2026 in St. Louis, MO.",
    "Lodging is reimbursed at the per diem rate of $150.00 per night.",
    "Contact innovation@dot.gov with any questions about reimbursement.",
  ]));
  has("a subset CID font is decoded through its ToUnicode CMap",
    cid.text, "Instructions for FHWA Invitational Travelers");
  has("and the body text comes through", cid.text, "St. Louis, MO");
  check("a date can be read out of a real-world PDF",
    Mp.eventDates(cid.text, "2026-08-23T12:00:00Z").start, "2026-10-15");
  check("and the room rate", Mp.findAmount(cid.text, /per\s*night|rate/).value, 150);
  check("and it is recognised as third-party paid",
    !!Mp.findReimbursement(cid.text), true);

  // Same file with the CMap removed: the glyph numbers must NOT be presented
  // as text, because a plausible wrong date on a travel form is the one
  // outcome worse than an empty field.
  var noCmap = await A.attachmentText("nocmap.pdf",
    bytes(Buffer.from(latin1FromU8(makeCidPdf([
      "Instructions for FHWA Invitational Travelers",
      "The peer exchange is held October 15, 2026 in St. Louis, MO.",
    ])).replace("/ToUnicode 5 0 R", "                 "), "latin1")));
  check("without a CMap the glyph numbers are discarded, not guessed at",
    noCmap.text, "");

  // A FILLED fillable form. The answers live in AcroForm field objects, never
  // in the page content — extracting only the page text of Matt's completed
  // Iowa DOT authorization returned the printed labels and underscore rules,
  // so a fully filled form read as blank.
  function makeAcroPdf(fields, pageText) {
    var content = "BT /F1 12 Tf 72 720 Td (" +
      pageText.replace(/([()\\])/g, "\\$1") + ") Tj ET";
    var objs = [
      "1 0 obj<</Type/Catalog/Pages 2 0 R/AcroForm<</Fields[" +
        fields.map(function (_, i) { return (5 + i) + " 0 R"; }).join(" ") + "]>>>>endobj",
      "2 0 obj<</Type/Pages/Count 1/Kids[3 0 R]>>endobj",
      "3 0 obj<</Type/Page/Parent 2 0 R/Contents 4 0 R/Annots[" +
        fields.map(function (_, i) { return (5 + i) + " 0 R"; }).join(" ") + "]>>endobj",
      "4 0 obj<</Length " + content.length + ">>stream\n" + content + "\nendstream endobj",
    ];
    fields.forEach(function (f, i) {
      objs.push((5 + i) + " 0 obj<</Type/Annot/Subtype/Widget/FT/" +
        (f.check ? "Btn" : "Tx") + "/T(" + f.name + ")" +
        (f.check ? "/AS/" + f.check : "/V" + (f.hex ? "<" + f.hex + ">" : "(" + f.value + ")")) +
        ">>endobj");
    });
    return bytes(Buffer.from("%PDF-1.4\n" + objs.join("\n") + "\n%%EOF", "latin1"));
  }

  var acro = await A.attachmentText("Travel Authorization Request.pdf", makeAcroPdf([
    { name: "Name & Cost Center", value: "Matthew Miller 30000" },
    { name: "Other Staff Attending", value: "Cedric Wilkinson, Brian Worrel, Tony Gustafson" },
    { name: "Name of Conference", value: "EDC-8 Midwest Peer Exchange" },
    { name: "Location", value: "Hyatt Regency 315 Chestnut St, St. Louis, MO 63102" },
    { name: "Conference Dates & Times", value: "10/15/26 7:00 am to 5:00 pm" },
    { name: "Departure Date", value: "10/14/2026" },
    { name: "Return Date", value: "10/15/2026" },
    { name: "Reason for Travel", value: "Matt is lead for Connected Corridors for Iowa." },
    { name: "Personal Vehicle", check: "Off" },
    { name: "State Vehicle", check: "Yes" },
    { name: "Commercial Air", check: "Off" },
    { name: "Cost of Travel Mode", value: "770 miles" },
    { name: "Parking", value: "40" },
    { name: "Luggage Fees", value: "0" },
  ], "Travel Authorization Form Name & Cost Center: ______ Departure Date: ______"));

  has("typed field values are read", acro.text, "Matthew Miller 30000");
  has("and the conference name", acro.text, "EDC-8 Midwest Peer Exchange");
  has("a ticked box is reported", acro.text, "State Vehicle: checked");
  check("an unticked box is not", /Personal Vehicle: checked/.test(acro.text), false);

  var af = Mp.formFields(acro.text);
  check("cost center out of a filled PDF", af.costCenter, "30000");
  check("name out of a filled PDF", af.name, "Matthew Miller");
  check("other staff", af.otherStaff, "Cedric Wilkinson, Brian Worrel, Tony Gustafson");
  check("conference", af.event, "EDC-8 Midwest Peer Exchange");
  check("departure date", af.departDate, "2026-10-14");
  check("return date", af.returnDate, "2026-10-15");
  check("reason for travel", af.reason, "Matt is lead for Connected Corridors for Iowa");
  check("state vehicle ticked", af.modeState, true);
  check("personal vehicle untouched", af.modePersonal, undefined);
  check("parking", af.cParking, 40);
  check("mileage figure is taken as entered", af.cTravelMode, 770);
  check("a venue's postal address is reduced to the destination",
    af.location, "St. Louis, MO");

  // Acrobat writes text as UTF-16BE with a byte-order mark often enough that
  // ignoring it would drop every value on some files.
  var utf16 = await A.attachmentText("utf16.pdf", makeAcroPdf([
    { name: "Name of Conference", hex: "FEFF00450044004300" + "2D0038" },
    { name: "Departure Date", value: "10/14/2026" },
    { name: "Parking", value: "40" },
  ], "form"));
  has("UTF-16BE field values decode", utf16.text, "EDC-8");

  check("prose passes the readability gate",
    A.looksLikeProse("The conference is held in Kansas City, Missouri on October 6 through 8."), true);
  check("a wall of symbols does not",
    A.looksLikeProse(Array(200).join("")), false);

  T.done("All Travel Desk attachment tests passed.");
})().catch(function (e) {
  console.error("attachment tests threw:", e);
  process.exit(1);
});
