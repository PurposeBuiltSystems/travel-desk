/*
 * Travel Desk — get readable text out of an email attachment.
 *
 * Bytes in, plain text out. Nothing here knows about Office.js or Graph; the
 * taskpane fetches the attachment content and hands the bytes over, and the
 * tests hand over fixtures built on disk.
 *
 * Formats, and why each one is worth having:
 *   .ics   the organiser already wrote the dates and venue as structured
 *          fields, so this is the only attachment that is better than the
 *          email body
 *   .docx  agendas and invitation letters — where the real detail usually is
 *   .xlsx  cost worksheets and traveler rosters
 *   .pptx  occasionally the agenda deck
 *   .pdf   the most common and the least cooperative; see pdfText()
 *   .txt .csv .htm  trivial
 *
 * Anything else (images, .msg, .zip) returns "" and is reported as skipped
 * rather than silently ignored, because "it found nothing in the PDF" and "it
 * never looked at the PDF" are different problems and the person needs to know
 * which one they have.
 */
/* global JSZip, DecompressionStream, TextDecoder, atob, Buffer */
(function (root) {
  "use strict";

  var MAX_BYTES = 12 * 1024 * 1024;   // a 12 MB attachment is not an agenda

  function b64ToBytes(b64) {
    var raw = String(b64 || "").replace(/[\r\n\s]/g, "");
    var bin;
    if (typeof atob === "function") { bin = atob(raw); }
    else if (typeof Buffer !== "undefined") { return Uint8Array.from(Buffer.from(raw, "base64")); }
    else { return new Uint8Array(0); }
    var out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) { out[i] = bin.charCodeAt(i) & 255; }
    return out;
  }

  /** Bytes → a string where each char is one byte. Chunked: a 5 MB PDF
   *  spread-applied in one call overflows the argument stack. */
  function latin1(bytes) {
    var CH = 0x8000, parts = [];
    for (var i = 0; i < bytes.length; i += CH) {
      parts.push(String.fromCharCode.apply(null, bytes.subarray(i, i + CH)));
    }
    return parts.join("");
  }

  function utf8(bytes) {
    if (typeof TextDecoder === "function") {
      try { return new TextDecoder("utf-8", { fatal: false }).decode(bytes); }
      catch (e) { /* fall through */ }
    }
    return latin1(bytes);
  }

  function ext(name) {
    var m = /\.([A-Za-z0-9]{1,5})$/.exec(String(name || "").trim());
    return m ? m[1].toLowerCase() : "";
  }

  // ------------------------------------------------------------------ zip

  function stripXml(xml, breakOn) {
    var s = String(xml || "");
    if (breakOn) { s = s.replace(breakOn, "\n"); }
    return s
      .replace(/<[^>]+>/g, " ")
      .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/[ \t]{2,}/g, " ")
      .replace(/\n\s*\n+/g, "\n")
      .trim();
  }

  async function zipEntries(bytes, wanted) {
    if (typeof JSZip === "undefined") { throw new Error("JSZip not loaded"); }
    var zip = await JSZip.loadAsync(bytes);
    var names = Object.keys(zip.files).filter(wanted).sort();
    var out = [];
    for (var i = 0; i < names.length; i++) {
      out.push(await zip.files[names[i]].async("string"));
    }
    return out;
  }

  async function docxText(bytes) {
    var xs = await zipEntries(bytes, function (n) {
      return /^word\/(document|header\d*|footer\d*)\.xml$/.test(n);
    });
    // Cells become tabs and rows become newlines, in that order — a table row
    // reading "Registration fee | $675.00" has to survive as ONE line, because
    // that adjacency is exactly what the money extractor matches on. Breaking
    // every cell onto its own line would separate the label from its amount
    // and quietly lose every figure in every table.
    return xs.map(function (x) {
      return stripXml(x
        .replace(/<\/w:p>\s*(?=<\/w:tc>)/g, " ")
        .replace(/<\/w:tc>/g, "\t")
        .replace(/<\/w:tr>/g, "\n")
        .replace(/<\/w:p>/g, "\n"), null);
    }).join("\n");
  }

  async function xlsxText(bytes) {
    if (typeof JSZip === "undefined") { throw new Error("JSZip not loaded"); }
    var zip = await JSZip.loadAsync(bytes);
    var shared = [];
    if (zip.files["xl/sharedStrings.xml"]) {
      var ss = await zip.files["xl/sharedStrings.xml"].async("string");
      var rx = /<si>([\s\S]*?)<\/si>/g, m;
      while ((m = rx.exec(ss))) { shared.push(stripXml(m[1])); }
    }
    var sheets = Object.keys(zip.files)
      .filter(function (n) { return /^xl\/worksheets\/sheet\d+\.xml$/.test(n); }).sort();
    var lines = [];
    for (var i = 0; i < sheets.length; i++) {
      var xml = await zip.files[sheets[i]].async("string");
      var rowRx = /<row[^>]*>([\s\S]*?)<\/row>/g, r;
      while ((r = rowRx.exec(xml))) {
        // The type attribute has to be read out of the attribute string rather
        // than matched inline: an optional group next to <c[^>]*> is trivially
        // satisfied by skipping it, so every t="s" cell would fall through to
        // the numeric branch and the sheet would come back as its raw shared-
        // string indexes — "0 1 / 2 675" instead of the actual words.
        var cells = [], cRx = /<c\b([^>]*)>([\s\S]*?)<\/c>/g, c;
        while ((c = cRx.exec(r[1]))) {
          var tm = /\bt="([^"]+)"/.exec(c[1]);
          var type = tm ? tm[1] : "";
          var vm = /<v>([\s\S]*?)<\/v>/.exec(c[2]);
          var im = /<is>([\s\S]*?)<\/is>/.exec(c[2]);
          if (im) { cells.push(stripXml(im[1])); }
          else if (!vm) { continue; }
          else if (type === "s") { cells.push(shared[parseInt(vm[1], 10)] || ""); }
          else { cells.push(stripXml(vm[1])); }
        }
        // Tab-joined: a "Registration  450" row then reads as one line with the
        // label and the number together, which is what the extractors expect.
        if (cells.join("").trim()) { lines.push(cells.join("\t")); }
      }
    }
    return lines.join("\n");
  }

  async function pptxText(bytes) {
    var xs = await zipEntries(bytes, function (n) {
      return /^ppt\/(slides|notesSlides)\/[a-z]+\d+\.xml$/.test(n);
    });
    return xs.map(function (x) { return stripXml(x.replace(/<\/a:p>/g, "\n")); }).join("\n");
  }

  // ------------------------------------------------------------------ pdf

  /**
   * Inflate a PDF stream with the platform's own decompressor.
   *
   * Two things here are less obvious than they look. The writer's promises
   * must be caught explicitly — an unhandled write/close rejection escapes the
   * surrounding try/catch entirely and takes the process down, which is how a
   * single malformed attachment could kill the whole prefill. And a decoder
   * that errors partway through has still produced real text up to that point;
   * PDF streams routinely carry a stray EOL before "endstream", so partial
   * output is kept rather than thrown away.
   */
  async function inflate(bytes) {
    if (typeof DecompressionStream !== "function") { return null; }
    var modes = ["deflate", "deflate-raw"];
    for (var i = 0; i < modes.length; i++) {
      var chunks = [], total = 0;
      try {
        var ds = new DecompressionStream(modes[i]);
        var w = ds.writable.getWriter();
        var writing = w.write(bytes).then(function () { return w.close(); });
        if (writing && writing.catch) { writing.catch(function () {}); }
        var rd = ds.readable.getReader();
        try {
          for (;;) {
            var res = await rd.read();
            if (res.done) { break; }
            chunks.push(res.value); total += res.value.length;
            if (total > 8e6) { break; }
          }
        } catch (streamErr) { /* keep whatever decoded before it gave up */ }
        if (!total) { continue; }
        var out = new Uint8Array(total), at = 0;
        chunks.forEach(function (c) { out.set(c, at); at += c.length; });
        return out;
      } catch (e) { /* try the next mode */ }
    }
    return null;
  }

  function unescapePdfString(s) {
    return s.replace(/\\(\n|\r\n?|[nrtbf()\\]|[0-7]{1,3})/g, function (_, c) {
      if (c === "n") { return "\n"; }
      if (c === "r") { return "\r"; }
      if (c === "t") { return "\t"; }
      if (c === "b" || c === "f") { return " "; }
      if (c === "(" || c === ")" || c === "\\") { return c; }
      if (/^[\n\r]/.test(c)) { return ""; }          // line continuation
      return String.fromCharCode(parseInt(c, 8) & 255);
    });
  }

  function hexPdfString(h) {
    var s = h.replace(/[^0-9A-Fa-f]/g, "");
    if (s.length % 2) { s += "0"; }
    var out = "";
    for (var i = 0; i < s.length; i += 2) {
      out += String.fromCharCode(parseInt(s.substr(i, 2), 16));
    }
    return out;
  }

  /** Pull the shown strings out of one decoded content stream. */
  function opsToText(content) {
    var out = "", line = "";
    var rx = /\((?:\\[\s\S]|[^\\()])*\)|<[0-9A-Fa-f\s]*>|\bT[Jj]\b|\bT[dDm*]\b|\bTD\b|'|"|\bBT\b|\bET\b/g;
    var m, pending = [];
    while ((m = rx.exec(content))) {
      var tok = m[0];
      if (tok.charAt(0) === "(") { pending.push(unescapePdfString(tok.slice(1, -1))); }
      else if (tok.charAt(0) === "<" && tok.charAt(1) !== "<") { pending.push(hexPdfString(tok.slice(1, -1))); }
      else if (tok === "Tj" || tok === "TJ" || tok === "'" || tok === '"') {
        line += pending.join(""); pending = [];
        if (tok === "'" || tok === '"') { out += line + "\n"; line = ""; }
      } else if (tok === "Td" || tok === "TD" || tok === "T*" || tok === "ET") {
        pending = [];
        if (line) { out += line + "\n"; line = ""; }
      } else { pending = []; }
    }
    if (line) { out += line + "\n"; }
    return out;
  }

  /**
   * Text out of a PDF, without a PDF library.
   *
   * This works on PDFs whose text is stored as WinAnsi strings — which is what
   * Word, Google Docs and most registration systems produce. It does NOT work
   * on scanned documents (no text at all, just an image) or on PDFs using
   * subset CID fonts, where the "text" is glyph indices and decodes to
   * nonsense. Rather than hand that nonsense to the extractors and let them
   * find a date in it, looksLikeProse() throws it away and the caller reports
   * the PDF as unreadable. A blank field the person fills in is a much better
   * outcome than a plausible wrong one on a form they are about to sign.
   */
  async function pdfText(bytes) {
    var raw = latin1(bytes);
    var out = [];
    var rx = /stream\r\n|stream\n|stream\r/g;
    var m;
    while ((m = rx.exec(raw))) {
      var start = m.index + m[0].length;
      var end = raw.indexOf("endstream", start);
      if (end < 0) { break; }
      var dict = raw.slice(Math.max(0, m.index - 700), m.index);
      // Images, fonts and metadata are streams too, and inflating a 4 MB JPEG
      // to look for dates in it is pure waste.
      if (/\/Subtype\s*\/(Image|Type1C|TrueType|CIDFontType\dC)\b/.test(dict) ||
          /\/Type\s*\/(XObject|Font|Metadata|ObjStm)\b/.test(dict) && !/\/Type\s*\/Page\b/.test(dict)) {
        if (/\/Subtype\s*\/Image\b/.test(dict)) { rx.lastIndex = end; continue; }
      }
      // The EOL that PDF puts between the data and "endstream" is not part of
      // the compressed data, and zlib treats it as trailing junk.
      var chunk = raw.slice(start, end).replace(/[\r\n]+$/, "");
      var body = null;
      if (/\/FlateDecode/.test(dict)) {
        var u8 = new Uint8Array(chunk.length);
        for (var i = 0; i < chunk.length; i++) { u8[i] = chunk.charCodeAt(i) & 255; }
        var inf = await inflate(u8);
        if (inf) { body = latin1(inf); }
      } else if (!/\/(DCTDecode|JPXDecode|CCITTFaxDecode|RunLengthDecode|LZWDecode|ASCII85Decode)/.test(dict)) {
        body = chunk;
      }
      if (body && /\bT[Jj]\b/.test(body)) { out.push(opsToText(body)); }
      rx.lastIndex = end;
      if (out.join("").length > 400000) { break; }
    }
    var text = out.join("\n").replace(/[ \t]{2,}/g, " ").replace(/\n{3,}/g, "\n\n").trim();
    return looksLikeProse(text) ? text : "";
  }

  /**
   * Is this text, or is it glyph indices that happen to be printable?
   *
   * Two signals, both cheap: real English has spaces (roughly one in six
   * characters) and is overwhelmingly made of ASCII letters, digits and
   * punctuation. Subset-font garbage has almost no spaces and a high share of
   * bytes above 127.
   */
  function looksLikeProse(s) {
    if (!s || s.length < 40) { return false; }
    var spaces = 0, printable = 0, letters = 0;
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i);
      if (c === 32) { spaces++; }
      if (c === 9 || c === 10 || c === 13 || (c >= 32 && c < 127)) { printable++; }
      if ((c >= 65 && c <= 90) || (c >= 97 && c <= 122)) { letters++; }
    }
    return printable / s.length > 0.9 &&
      letters / s.length > 0.45 &&
      spaces / s.length > 0.07;
  }

  // ------------------------------------------------------------- dispatch

  var SKIP_NOTE = {
    png: "image", jpg: "image", jpeg: "image", gif: "image", bmp: "image",
    tif: "image", tiff: "image", svg: "image", heic: "image", webp: "image",
    msg: "an Outlook message", eml: "an email file", zip: "an archive",
    doc: "old Word format (.doc)", xls: "old Excel format (.xls)",
    ppt: "old PowerPoint format (.ppt)",
  };

  /**
   * One attachment → { name, text, kind, note }.
   *
   * Never throws: an attachment that cannot be read is a note on the result,
   * not a failure of the whole prefill. Ten attachments where one is a
   * corrupt PDF should still fill the form from the other nine.
   */
  async function attachmentText(name, bytes) {
    var e = ext(name);
    var res = { name: name, text: "", kind: e, note: "" };
    try {
      if (!bytes || !bytes.length) { res.note = "empty"; return res; }
      if (bytes.length > MAX_BYTES) {
        res.note = "too large to scan (" + Math.round(bytes.length / 1048576) + " MB)";
        return res;
      }
      if (e === "ics" || e === "vcs" || e === "ical") { res.text = utf8(bytes); return res; }
      if (e === "txt" || e === "csv" || e === "md" || e === "log") { res.text = utf8(bytes); return res; }
      if (e === "htm" || e === "html") { res.text = utf8(bytes); return res; }
      if (e === "docx") { res.text = await docxText(bytes); return res; }
      if (e === "xlsx" || e === "xlsm") { res.text = await xlsxText(bytes); return res; }
      if (e === "pptx") { res.text = await pptxText(bytes); return res; }
      if (e === "pdf") {
        res.text = await pdfText(bytes);
        if (!res.text) { res.note = "no readable text (scanned image, or an unsupported font)"; }
        return res;
      }
      res.note = SKIP_NOTE[e] ? "skipped — " + SKIP_NOTE[e] : "skipped — ." + (e || "unknown") + " isn't readable";
      return res;
    } catch (err) {
      res.text = "";
      res.note = "couldn't be read (" + (err && err.message ? err.message : "error") + ")";
      return res;
    }
  }

  var api = {
    attachmentText: attachmentText,
    b64ToBytes: b64ToBytes,
    pdfText: pdfText,
    docxText: docxText,
    xlsxText: xlsxText,
    pptxText: pptxText,
    looksLikeProse: looksLikeProse,
    opsToText: opsToText,
    _internals: { ext: ext, latin1: latin1, stripXml: stripXml },
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.TdAttach = api; }
})(typeof self !== "undefined" ? self : this);
