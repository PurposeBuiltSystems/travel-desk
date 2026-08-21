/*
 * Travel Desk — minimal .xlsx generator (pure).
 *
 * An .xlsx is a ZIP of OOXML parts. This builds the parts as strings; the
 * task pane zips them with JSZip and uploads via Graph. It deliberately
 * writes ONLY a header row and no table part: creating the Excel Table is
 * left to the Graph workbook API (tables/add), which is the same call the
 * "make this sheet a table" action uses — one code path, and no hand-rolled
 * table XML to get subtly wrong.
 */
(function (root) {
  "use strict";

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;")
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
  }

  /** 1 -> A, 26 -> Z, 27 -> AA */
  function colLetter(n) {
    var s = "";
    while (n > 0) {
      var r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  /** Excel's A1-style range for a header row of n columns. */
  /**
   * Address for the header row, in the form Graph itself uses.
   *
   * This used to quote the sheet name - 'Planner'!A1:R1 - which is valid
   * Excel syntax and is NOT accepted by the workbook API's tables/add. The
   * proof was sitting in the same codebase: "Make this sheet a table" works,
   * and it passes the address returned by usedRange, which Graph writes
   * unquoted as Planner!A1:R1. Creating a planner failed at exactly that step
   * - the workbook uploaded fine and then had no table.
   *
   * The sheet name here is ours ("Planner"), so it never needs quoting. A
   * name containing a space or apostrophe would not survive this - that is
   * an accepted limit, not an oversight, because nothing here lets a user
   * name the sheet.
   */
  function headerRange(sheetName, n) {
    return String(sheetName) + "!A1:" + colLetter(n) + "1";
  }

  /**
   * Returns { parts: {path: xml}, sheetName, range } — feed parts to JSZip.
   */
  function buildWorkbook(headers, sheetName) {
    var cols = (headers && headers.length) ? headers : ["Column1"];
    var sheet = sheetName || "Planner";

    var cells = cols.map(function (h, i) {
      return '<c r="' + colLetter(i + 1) + '1" t="inlineStr"><is><t xml:space="preserve">' +
        esc(h) + "</t></is></c>";
    }).join("");

    var sheetXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData><row r="1">' + cells + "</row></sheetData>" +
      "</worksheet>";

    var workbookXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="' + esc(sheet) + '" sheetId="1" r:id="rId1"/></sheets>' +
      "</workbook>";

    var workbookRels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ' +
      'Target="worksheets/sheet1.xml"/>' +
      "</Relationships>";

    var pkgRels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" ' +
      'Target="xl/workbook.xml"/>' +
      "</Relationships>";

    var contentTypes =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ' +
      'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ' +
      'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      "</Types>";

    return {
      sheetName: sheet,
      range: headerRange(sheet, cols.length),
      parts: {
        "[Content_Types].xml": contentTypes,
        "_rels/.rels": pkgRels,
        "xl/workbook.xml": workbookXml,
        "xl/_rels/workbook.xml.rels": workbookRels,
        "xl/worksheets/sheet1.xml": sheetXml,
      },
    };
  }

  var api = { buildWorkbook: buildWorkbook, colLetter: colLetter, headerRange: headerRange };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.XlsxGen = api; }
})(typeof self !== "undefined" ? self : this);
