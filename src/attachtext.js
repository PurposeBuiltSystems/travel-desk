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

  /**
   * Pull the shown strings out of one decoded content stream.
   *
   * `fonts` maps a resource name ("F0") to a decoder for that font, because
   * with a subset CID font the bytes in the string are glyph numbers and mean
   * nothing without one. The current font is whatever the last "/Fx n Tf"
   * selected, exactly as the PDF viewer tracks it.
   */
  function opsToText(content, fonts) {
    var out = "", line = "";
    var rx = /\((?:\\[\s\S]|[^\\()])*\)|<[0-9A-Fa-f\s]*>|\/([A-Za-z0-9#+._-]+)\s+[\d.]+\s+Tf\b|\bT[Jj]\b|\bT[dDm*]\b|\bTD\b|'|"|\bBT\b|\bET\b/g;
    var m, pending = [], font = null;
    function decode(bytes) { return font ? font(bytes) : bytes; }
    while ((m = rx.exec(content))) {
      var tok = m[0];
      if (m[1] !== undefined) { font = (fonts && fonts[m[1]]) || null; continue; }
      if (tok.charAt(0) === "(") { pending.push(decode(unescapePdfString(tok.slice(1, -1)))); }
      else if (tok.charAt(0) === "<" && tok.charAt(1) !== "<") { pending.push(decode(hexPdfString(tok.slice(1, -1)))); }
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

  // ------------------------------------------------------- ToUnicode CMaps

  /**
   * Parse a /ToUnicode CMap into a decoder function.
   *
   * This is what makes a normal PDF readable. Word, PDFsharp and most modern
   * generators embed SUBSET fonts (/BXNAPB+Calibri) with /Encoding/Identity-H,
   * which means the "text" in the content stream is a list of glyph numbers
   * private to that one file — "October" might be bytes 0x2F 0x0C 0x11…, and
   * reading them as characters produces confident nonsense. The generator also
   * writes a /ToUnicode CMap mapping each glyph number back to real Unicode,
   * precisely so that text can be copied out. Following it is the difference
   * between reading agendas and refusing almost every PDF a person will
   * actually attach.
   *
   *   beginbfchar  <0044> <0054>            endbfchar
   *   beginbfrange <0003> <0008> <0020>     endbfrange
   *   beginbfrange <0010> <0012> [<0041> <0042> <0043>] endbfrange
   */
  function parseCMap(text) {
    var map = {}, srcBytes = 2, m, rx;

    function uni(hex) {
      var s = "";
      for (var i = 0; i + 3 < hex.length + 1; i += 4) {
        var cp = parseInt(hex.substr(i, 4), 16);
        if (!isNaN(cp)) { s += String.fromCharCode(cp); }
      }
      return s;
    }

    rx = /begincodespacerange([\s\S]*?)endcodespacerange/g;
    while ((m = rx.exec(text))) {
      var cs = /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/.exec(m[1]);
      if (cs) { srcBytes = Math.max(1, Math.round(cs[1].length / 2)); }
    }

    rx = /beginbfchar([\s\S]*?)endbfchar/g;
    while ((m = rx.exec(text))) {
      var crx = /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]*)>/g, c;
      while ((c = crx.exec(m[1]))) { map[parseInt(c[1], 16)] = uni(c[2]); }
    }

    rx = /beginbfrange([\s\S]*?)endbfrange/g;
    while ((m = rx.exec(text))) {
      var body = m[1];
      var lrx = /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(\[[\s\S]*?\]|<[0-9A-Fa-f]*>)/g, l;
      while ((l = lrx.exec(body))) {
        var lo = parseInt(l[1], 16), hi = parseInt(l[2], 16), dst = l[3];
        if (hi - lo > 65535) { continue; }
        if (dst.charAt(0) === "[") {
          var arr = dst.match(/<([0-9A-Fa-f]*)>/g) || [];
          for (var i = 0; i <= hi - lo && i < arr.length; i++) {
            map[lo + i] = uni(arr[i].slice(1, -1));
          }
        } else {
          var base = dst.slice(1, -1);
          var start = parseInt(base.substr(base.length - 4), 16);
          var prefix = base.slice(0, base.length - 4);
          for (var j = 0; j <= hi - lo; j++) {
            map[lo + j] = uni(prefix + ("000" + (start + j).toString(16)).slice(-4));
          }
        }
      }
    }

    if (!Object.keys(map).length) { return null; }
    return function (raw) {
      var out = "";
      if (srcBytes === 2) {
        for (var i = 0; i + 1 < raw.length; i += 2) {
          var code = (raw.charCodeAt(i) << 8) | raw.charCodeAt(i + 1);
          out += map[code] != null ? map[code] : "";
        }
      } else {
        for (var k = 0; k < raw.length; k++) {
          var c1 = raw.charCodeAt(k);
          out += map[c1] != null ? map[c1] : raw.charAt(k);
        }
      }
      return out;
    };
  }

  /** `12 0 obj … endobj` → { 12: {dict, streamStart, streamEnd} } */
  function parseObjects(raw) {
    var objs = {}, rx = /(\d+)\s+\d+\s+obj\b/g, m;
    while ((m = rx.exec(raw))) {
      var end = raw.indexOf("endobj", m.index);
      if (end < 0) { end = Math.min(raw.length, m.index + 200000); }
      var seg = raw.slice(m.index, end);
      var sm = /stream\r\n|stream\n|stream\r/.exec(seg);
      objs[m[1]] = {
        dict: sm ? seg.slice(0, sm.index) : seg,
        streamStart: sm ? m.index + sm.index + sm[0].length : -1,
        streamEnd: sm ? raw.indexOf("endstream", m.index + sm.index) : -1,
      };
    }
    return objs;
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
    var objs = parseObjects(raw);

    /** Decompress one object's stream, or null. */
    async function streamOf(num) {
      var o = objs[num];
      if (!o || o.streamStart < 0 || o.streamEnd < 0) { return null; }
      // The EOL that PDF puts between the data and "endstream" is not part of
      // the compressed data, and zlib treats it as trailing junk.
      var chunk = raw.slice(o.streamStart, o.streamEnd).replace(/[\r\n]+$/, "");
      if (/\/(DCTDecode|JPXDecode|CCITTFaxDecode|RunLengthDecode|LZWDecode|ASCII85Decode)/.test(o.dict)) { return null; }
      if (!/\/FlateDecode/.test(o.dict)) { return chunk; }
      var u8 = new Uint8Array(chunk.length);
      for (var i = 0; i < chunk.length; i++) { u8[i] = chunk.charCodeAt(i) & 255; }
      var inf = await inflate(u8);
      return inf ? latin1(inf) : null;
    }

    /** "/Resources 7 0 R" and "/Resources <<…>>" both have to work. */
    function deref(dict, key) {
      var ref = new RegExp("\\/" + key + "\\s+(\\d+)\\s+\\d+\\s+R\\b").exec(dict);
      if (ref) { return objs[ref[1]] ? objs[ref[1]].dict : ""; }
      var i = dict.indexOf("/" + key);
      if (i < 0) { return ""; }
      var j = dict.indexOf("<<", i);
      // The gap is measured from the END of the key, not its start — measuring
      // from the start makes every key longer than four characters look like a
      // miss, so "/Resources<<…>>" never resolved and no page had any fonts.
      if (j < 0 || j - (i + key.length + 1) > 4) { return ""; }
      var depth = 0;
      for (var k = j; k < dict.length - 1; k++) {
        if (dict.substr(k, 2) === "<<") { depth++; k++; }
        else if (dict.substr(k, 2) === ">>") { depth--; k++; if (!depth) { return dict.slice(j, k + 1); } }
      }
      return dict.slice(j);
    }

    var cmapCache = {};
    async function cmapFor(fontNum) {
      if (cmapCache[fontNum] !== undefined) { return cmapCache[fontNum]; }
      cmapCache[fontNum] = null;
      var fd = objs[fontNum] && objs[fontNum].dict;
      if (!fd) { return null; }
      var tu = /\/ToUnicode\s+(\d+)\s+\d+\s+R\b/.exec(fd);
      if (!tu) { return null; }
      var cm = await streamOf(tu[1]);
      if (cm) { cmapCache[fontNum] = parseCMap(cm); }
      return cmapCache[fontNum];
    }

    var out = [];
    var pageNums = Object.keys(objs).filter(function (n) {
      return /\/Type\s*\/Page\b/.test(objs[n].dict) && !/\/Type\s*\/Pages\b/.test(objs[n].dict);
    });

    for (var p = 0; p < pageNums.length; p++) {
      var page = objs[pageNums[p]].dict;

      var fonts = {};
      var fdict = deref(deref(page, "Resources") || page, "Font");
      var frx = /\/([A-Za-z0-9#+._-]+)\s+(\d+)\s+\d+\s+R\b/g, fm;
      while ((fm = frx.exec(fdict))) {
        var cm2 = await cmapFor(fm[2]);
        if (cm2) { fonts[fm[1]] = cm2; }
      }

      var contents = [];
      var single = /\/Contents\s+(\d+)\s+\d+\s+R\b/.exec(page);
      var arr = /\/Contents\s*\[([^\]]*)\]/.exec(page);
      if (single) { contents.push(single[1]); }
      else if (arr) {
        var arx = /(\d+)\s+\d+\s+R/g, am;
        while ((am = arx.exec(arr[1]))) { contents.push(am[1]); }
      }

      for (var c = 0; c < contents.length; c++) {
        var body = await streamOf(contents[c]);
        if (body && /\bT[Jj]\b/.test(body)) { out.push(opsToText(body, fonts)); }
      }
      if (out.join("").length > 400000) { break; }
    }

    // Fall back to a blind stream sweep for PDFs whose page tree we could not
    // follow — a linearised or object-stream file may hide /Type/Page from a
    // regex, and finding some text beats finding none.
    if (!out.join("").trim()) {
      var srx = /stream\r\n|stream\n|stream\r/g, sm2;
      while ((sm2 = srx.exec(raw))) {
        var st = sm2.index + sm2[0].length;
        var en = raw.indexOf("endstream", st);
        if (en < 0) { break; }
        var d = raw.slice(Math.max(0, sm2.index - 700), sm2.index);
        if (/\/Subtype\s*\/Image\b/.test(d)) { srx.lastIndex = en; continue; }
        var ch = raw.slice(st, en).replace(/[\r\n]+$/, "");
        var b2 = null;
        if (/\/FlateDecode/.test(d)) {
          var u2 = new Uint8Array(ch.length);
          for (var q = 0; q < ch.length; q++) { u2[q] = ch.charCodeAt(q) & 255; }
          var i2 = await inflate(u2);
          if (i2) { b2 = latin1(i2); }
        } else if (!/\/(DCTDecode|JPXDecode|CCITTFaxDecode|LZWDecode|ASCII85Decode)/.test(d)) { b2 = ch; }
        if (b2 && /\bT[Jj]\b/.test(b2)) { out.push(opsToText(b2, null)); }
        srx.lastIndex = en;
        if (out.join("").length > 400000) { break; }
      }
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
    parseCMap: parseCMap,
    parseObjects: parseObjects,
    _internals: { ext: ext, latin1: latin1, stripXml: stripXml },
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.TdAttach = api; }
})(typeof self !== "undefined" ? self : this);
