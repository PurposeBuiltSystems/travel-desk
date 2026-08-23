/*
 * Travel Desk — read a registration/confirmation email and turn it into a
 * filled-in travel request.
 *
 * Everything in here is PURE: text in, structured guess out. No Office.js, no
 * Graph, no DOM. The taskpane does the I/O (fetch the body, pull attachment
 * bytes) and hands the text to buildPrefill(); the tests hand it fixtures.
 *
 * Deterministic — no AI, no network. That is a deliberate constraint, not a
 * limitation we are working around: a travel authorization is a financial
 * document, and a wrong-but-fluent date or dollar figure is worse than a
 * blank field. So every extractor here either matches a real pattern in the
 * source text or returns nothing, and every value it does return is reported
 * with the exact snippet it came from, so the person can check it in one
 * glance before pressing Send.
 */
(function (root) {
  "use strict";

  var MONTHS = {
    jan: 1, january: 1, feb: 2, february: 2, mar: 3, march: 3, apr: 4, april: 4,
    may: 5, jun: 6, june: 6, jul: 7, july: 7, aug: 8, august: 8,
    sep: 9, sept: 9, september: 9, oct: 10, october: 10, nov: 11, november: 11,
    dec: 12, december: 12,
  };

  var STATES = {
    AL: 1, AK: 1, AZ: 1, AR: 1, CA: 1, CO: 1, CT: 1, DE: 1, DC: 1, FL: 1,
    GA: 1, HI: 1, ID: 1, IL: 1, IN: 1, IA: 1, KS: 1, KY: 1, LA: 1, ME: 1,
    MD: 1, MA: 1, MI: 1, MN: 1, MS: 1, MO: 1, MT: 1, NE: 1, NV: 1, NH: 1,
    NJ: 1, NM: 1, NY: 1, NC: 1, ND: 1, OH: 1, OK: 1, OR: 1, PA: 1, RI: 1,
    SC: 1, SD: 1, TN: 1, TX: 1, UT: 1, VT: 1, VA: 1, WA: 1, WV: 1, WI: 1,
    WY: 1, PR: 1,
  };

  var STATE_NAMES = {
    alabama: "AL", alaska: "AK", arizona: "AZ", arkansas: "AR",
    california: "CA", colorado: "CO", connecticut: "CT", delaware: "DE",
    florida: "FL", georgia: "GA", hawaii: "HI", idaho: "ID", illinois: "IL",
    indiana: "IN", iowa: "IA", kansas: "KS", kentucky: "KY", louisiana: "LA",
    maine: "ME", maryland: "MD", massachusetts: "MA", michigan: "MI",
    minnesota: "MN", mississippi: "MS", missouri: "MO", montana: "MT",
    nebraska: "NE", nevada: "NV", "new hampshire": "NH", "new jersey": "NJ",
    "new mexico": "NM", "new york": "NY", "north carolina": "NC",
    "north dakota": "ND", ohio: "OH", oklahoma: "OK", oregon: "OR",
    pennsylvania: "PA", "rhode island": "RI", "south carolina": "SC",
    "south dakota": "SD", tennessee: "TN", texas: "TX", utah: "UT",
    vermont: "VT", virginia: "VA", washington: "WA", "west virginia": "WV",
    wisconsin: "WI", wyoming: "WY",
  };

  function pad(n) { return (n < 10 ? "0" : "") + n; }

  function iso(y, m, d) {
    if (!y || !m || !d || m < 1 || m > 12 || d < 1 || d > 31) { return ""; }
    return y + "-" + pad(m) + "-" + pad(d);
  }

  /** Two-digit years in confirmations are always this century in practice. */
  function fullYear(y) {
    var n = parseInt(y, 10);
    if (isNaN(n)) { return 0; }
    return n < 100 ? 2000 + n : n;
  }

  function clean(s) {
    return String(s == null ? "" : s)
      .replace(/\r\n?/g, "\n")
      .replace(/ /g, " ")
      .replace(/[‐-―]/g, "-")     // every dash variant → hyphen
      .replace(/[‘’]/g, "'")
      .replace(/[“”]/g, '"')
      .replace(/[ \t]+/g, " ")
      .replace(/\n{3,}/g, "\n\n");
  }

  /** Strip HTML to text, keeping line structure so proximity scoring works. */
  function htmlToText(html) {
    return clean(String(html == null ? "" : html)
      .replace(/<(script|style)[\s\S]*?<\/\1>/gi, " ")
      .replace(/<!--[\s\S]*?-->/g, " ")
      .replace(/<\/(p|div|tr|li|h[1-6]|table)>/gi, "\n")
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<td[^>]*>/gi, "\t")
      .replace(/<[^>]+>/g, " ")
      .replace(/&nbsp;/gi, " ").replace(/&amp;/gi, "&")
      .replace(/&lt;/gi, "<").replace(/&gt;/gi, ">")
      .replace(/&quot;/gi, '"').replace(/&#39;/gi, "'")
      .replace(/&#(\d+);/g, function (_, d) {
        return String.fromCharCode(parseInt(d, 10));
      }));
  }

  // ---------------------------------------------------------------- subject

  var SUBJ_NOISE = [
    /\bregistration\s+(is\s+)?confirm(ed|ation)\b/i,
    /\bconfirmation\s+(of|for)\s+registration\b/i,
    /\byour\s+registration\s+(for|to)\b/i,
    /\bregistration\s+receipt\b/i,
    /\bthank\s+you\s+for\s+registering\s+(for|to)?\b/i,
    /\byou('|’)?re\s+registered\s+(for|to)\b/i,
    /\bbooking\s+confirm(ed|ation)\b/i,
    /\border\s+confirm(ed|ation)\b/i,
    /\bconfirmation\b/i,
    /\bregistration\b/i,
    /\breceipt\b/i,
    /\bitinerary\b/i,
    /\binvitation\s+to\b/i,
    /\bsave\s+the\s+date\b/i,
  ];

  /**
   * "Fw: EDC-8 Midwest Peer Exchange Registration Confirmation"
   *   → "EDC-8 Midwest Peer Exchange"
   *
   * Reply/forward prefixes come off first and repeatedly — a real forwarded
   * confirmation often carries "Fw: RE: Fwd:" stacked up. Then the boilerplate
   * that makes it a *confirmation* comes off, because what belongs in the
   * Event field is the name of the thing, not the fact that you registered.
   */
  function eventFromSubject(subject) {
    var s = clean(subject).replace(/\n/g, " ").trim();
    var prev = null;
    while (prev !== s) {
      prev = s;
      s = s.replace(/^\s*(re|fw|fwd|tr|aw)\s*:\s*/i, "").trim();
      s = s.replace(/^\s*\[[^\]]{1,40}\]\s*/, "").trim();   // [EXTERNAL] etc.
    }
    // The reference number goes before the boilerplate does: once the word
    // "Confirmation" has been stripped out of "Confirmation #A8F32K91", what
    // is left no longer looks like a reference and survives into the name.
    s = stripRef(s);
    SUBJ_NOISE.forEach(function (rx) { s = s.replace(rx, " "); });
    s = stripRef(s);
    return s.replace(/\s{2,}/g, " ")
      .replace(/\s*[-|–—:,]+\s*$/, "")
      .replace(/^\s*[-|–—:,]+\s*/, "")
      .trim();
  }

  /**
   * "- Confirmation #A8F32K91", "(Conf# 4820193)", "- REF: XY-88213"
   *
   * Two guards keep this off real event names. The keyword needs a word
   * boundary after it, or "Conference" reads as "Conf" + a reference; and the
   * reference itself has to contain a digit, because "…for the Rural Transit
   * Conference" is a title and "A8F32K91" is not.
   */
  function stripRef(s) {
    var REF = "(?=[A-Za-z0-9-]*\\d)[A-Za-z0-9][A-Za-z0-9-]{4,}";
    return String(s)
      .replace(new RegExp("[\\s\\-–—(,]*\\b(?:conf(?:irmation)?|ref(?:erence)?|order|reg(?:istration)?)\\b\\.?\\s*(?:#|no\\.?|number|id)?\\s*:?\\s*#?\\s*" + REF + "\\)?\\s*$", "i"), "")
      .replace(new RegExp("[\\s\\-–—(,]+#\\s*" + REF + "\\)?\\s*$"), "")
      .trim();
  }

  // ------------------------------------------------------------------ dates

  /**
   * Every date we can find, each with the character offset where it occurred
   * so callers can prefer the one nearest a "dates:" style label.
   *
   * Ranges matter more than single dates here: "September 29 - October 1,
   * 2026" is one event, and reading it as two unrelated dates would put the
   * departure a month from the return.
   */
  function findDates(text) {
    var t = clean(text);
    var out = [];
    var m, rx;

    function push(start, end, at, kind) {
      if (!start) { return; }
      out.push({ start: start, end: end || "", at: at, kind: kind });
    }

    // "September 29 - October 1, 2026" / "Sept 29-Oct 1, 2026"
    rx = /\b([A-Za-z]{3,9})\.?\s+(\d{1,2})\s*(?:-|to|through|thru|until)\s*([A-Za-z]{3,9})\.?\s+(\d{1,2})\s*,?\s*(\d{4}|\d{2})\b/g;
    while ((m = rx.exec(t))) {
      var m1 = MONTHS[m[1].toLowerCase()], m2 = MONTHS[m[3].toLowerCase()];
      if (!m1 || !m2) { continue; }
      var y = fullYear(m[5]);
      // A range that wraps the new year: "December 30 - January 2, 2027" means
      // the START is the previous year, not the end.
      var ys = m2 < m1 ? y - 1 : y;
      push(iso(ys, m1, +m[2]), iso(y, m2, +m[4]), m.index, "range");
    }

    // "January 10-14, 2027" / "Jan 10 to 14 2027"
    rx = /\b([A-Za-z]{3,9})\.?\s+(\d{1,2})\s*(?:-|to|through|thru|until|&)\s*(\d{1,2})\s*,?\s*(\d{4}|\d{2})\b/g;
    while ((m = rx.exec(t))) {
      var mo = MONTHS[m[1].toLowerCase()];
      if (!mo) { continue; }
      var yr = fullYear(m[4]);
      push(iso(yr, mo, +m[2]), iso(yr, mo, +m[3]), m.index, "range");
    }

    // "March 3, 2027" / "Mar. 3 2027" / "3 March 2027"
    rx = /\b([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?\s*,?\s*(\d{4})\b/g;
    while ((m = rx.exec(t))) {
      var mm = MONTHS[m[1].toLowerCase()];
      if (mm) { push(iso(+m[3], mm, +m[2]), "", m.index, "single"); }
    }
    rx = /\b(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]{3,9})\.?\s*,?\s*(\d{4})\b/g;
    while ((m = rx.exec(t))) {
      var mn = MONTHS[m[2].toLowerCase()];
      if (mn) { push(iso(+m[3], mn, +m[1]), "", m.index, "single"); }
    }

    // ISO and slashed forms
    rx = /\b(\d{4})-(\d{1,2})-(\d{1,2})\b/g;
    while ((m = rx.exec(t))) { push(iso(+m[1], +m[2], +m[3]), "", m.index, "single"); }
    rx = /\b(\d{1,2})\/(\d{1,2})\/(\d{4}|\d{2})\b/g;
    while ((m = rx.exec(t))) {
      push(iso(fullYear(m[3]), +m[1], +m[2]), "", m.index, "single");
    }

    // Same date found by two patterns (a range also matches "single") — keep
    // the richer one.
    var seen = {};
    var uniq = [];
    out.sort(function (a, b) {
      if (a.start !== b.start) { return a.start < b.start ? -1 : 1; }
      return (b.end ? 1 : 0) - (a.end ? 1 : 0);
    });
    out.forEach(function (d) {
      var k = d.start + "|" + (d.end || "");
      var loose = d.start;
      if (seen[k]) { return; }
      if (!d.end && seen["range:" + loose]) { return; }
      seen[k] = 1;
      if (d.end) { seen["range:" + d.start] = 1; seen["range:" + d.end] = 1; }
      uniq.push(d);
    });
    return uniq;
  }

  var DATE_CUE = /(event|conference|meeting|exchange|summit|workshop|session|held|takes?\s+place|when|dates?|schedule|agenda|arriv|depart|check-?in|check-?out|travel)/i;

  /**
   * Pick the event's dates out of everything the message mentions.
   *
   * Confirmations are full of dates that are not the event: the date you
   * registered, a payment date, an "respond by" deadline, the copyright year
   * in a footer. Two rules do most of the work — the event has not happened
   * yet relative to the message, and it is written near a word like
   * "conference" or "dates". A range beats a bare date, because someone
   * bothered to write both ends of it.
   */
  function eventDates(text, receivedIso) {
    var t = clean(text);
    var all = findDates(t);
    if (!all.length) { return { start: "", end: "" }; }

    var floor = "";
    if (receivedIso) {
      var d = new Date(receivedIso);
      if (!isNaN(d.getTime())) {
        d.setDate(d.getDate() - 3);       // registered the day it starts? fine.
        floor = iso(d.getFullYear(), d.getMonth() + 1, d.getDate());
      }
    }

    var best = null, bestScore = -1e9;
    all.forEach(function (c) {
      var score = 0;
      if (c.end) { score += 40; }
      if (floor && c.start >= floor) { score += 60; }
      else if (floor) { score -= 50; }
      var near = t.slice(Math.max(0, c.at - 90), c.at + 60);
      if (DATE_CUE.test(near)) { score += 35; }
      if (/(register|paid|payment|invoice|receipt|order|issued|sent|due|expires?|deadline|respond|rsvp|cancel)/i.test(near)) {
        score -= 45;
      }
      score -= c.at / 100000;             // earlier in the message, marginally
      if (score > bestScore) { bestScore = score; best = c; }
    });
    return { start: best.start, end: best.end || "" };
  }

  // --------------------------------------------------------------- location

  var LOC_CUE = /(location|venue|held|hosted|where|address|hotel|conference\s+center|convention|campus|site|city)/i;

  /**
   * Find "City, ST".
   *
   * The trap here is the signature block. A forwarded confirmation carries
   * the sender's own address — "Ames, IA 50010" — and a naive first-match
   * grabs it and confidently proposes that you are travelling to your own
   * office. So candidates are scored, addresses attached to a ZIP right after
   * a street line are penalised, and `homeCity` (the user's own, which we
   * already know from their settings) is excluded outright.
   */
  /**
   * "Location Kansas City" → "Kansas City".
   *
   * A spreadsheet cell or a table row puts the label and the value side by
   * side with only whitespace between them, and the multi-word city pattern
   * happily swallows the label. The match is lazy and requires whitespace
   * after the label word, so a city that legitimately contains one — Kansas
   * City, Sioux City, Salt Lake City — is left intact.
   */
  var CITY_LABEL = /^.*?\b(?:location|venue|address|city|where|held(?:\s+in)?|site|destination|hotel|host(?:ed)?(?:\s+(?:in|by))?|takes\s+place(?:\s+in)?|meeting|conference|event|in|at|to)\s*:?\s+/i;

  function trimCityLabel(city) {
    var out = String(city).replace(CITY_LABEL, "").trim();
    return out || String(city).trim();
  }

  function findLocation(text, homeCity) {
    var t = clean(text);
    var home = String(homeCity || "").toLowerCase().replace(/\s+/g, " ").trim();
    var cands = [];
    var m;

    // [ \t]+ rather than \s+ between the words of a city: a place name does
    // not straddle a line break, and allowing one lets a street line above an
    // address join its city, so "800 Lincoln Way\nAmes, IA" is offered as a
    // destination called "Lincoln Way Ames".
    var CITY = "\\b([A-Z][A-Za-z.'-]+(?:[ \\t]+[A-Z][A-Za-z.'-]+){0,3})[ \\t]*,[ \\t]*";
    var rx = new RegExp(CITY + "([A-Z]{2})\\b(?![ \\t]*\\d{5})", "g");
    while ((m = rx.exec(t))) {
      if (STATES[m[2]]) { cands.push({ city: m[1], st: m[2], at: m.index, zip: false }); }
    }
    rx = new RegExp(CITY + "([A-Z]{2})[ \\t]+\\d{5}(?:-\\d{4})?\\b", "g");
    while ((m = rx.exec(t))) {
      if (STATES[m[2]]) { cands.push({ city: m[1], st: m[2], at: m.index, zip: true }); }
    }
    rx = new RegExp(CITY + "([A-Za-z]{4,20}(?:[ \\t]+[A-Za-z]{4,20})?)\\b", "g");
    while ((m = rx.exec(t))) {
      var abbr = STATE_NAMES[m[2].toLowerCase()];
      if (abbr) { cands.push({ city: m[1], st: abbr, at: m.index, zip: false }); }
    }

    if (!cands.length) { return { value: "", alternates: [] }; }

    var scored = cands.map(function (c) {
      var label = trimCityLabel(c.city) + ", " + c.st;
      var score = 0;
      var before = t.slice(Math.max(0, c.at - 140), c.at);
      var near = before + t.slice(c.at, c.at + 60);
      if (LOC_CUE.test(near)) { score += 40; }
      if (c.zip) { score -= 12; }
      // A street number immediately above is a postal address, and postal
      // addresses in a confirmation are usually the sender's, not the venue's.
      if (/\d{3,5}\s+[A-Z][A-Za-z.'-]+\s+(St|Street|Ave|Avenue|Rd|Road|Blvd|Way|Dr|Drive|Ln|Lane|Pkwy|Parkway|Hwy|Highway|Ct|Court|Pl|Place|Sq|Square)\b[^\n]*\n?[^\n]*$/i.test(before)) {
        score -= 30;
      }
      if (/(hotel|marriott|hilton|hyatt|sheraton|westin|embassy|doubletree|convention|conference\s+cent)/i.test(near)) {
        score += 25;
      }
      if (home && label.toLowerCase().indexOf(home) === 0) { score -= 200; }
      score -= c.at / 100000;
      return { label: label, score: score };
    }).filter(function (c) { return c.score > -100; });

    if (!scored.length) { return { value: "", alternates: [] }; }
    scored.sort(function (a, b) { return b.score - a.score; });

    var seen = {}, order = [];
    scored.forEach(function (c) {
      if (seen[c.label]) { return; }
      seen[c.label] = 1; order.push(c.label);
    });
    return { value: order[0], alternates: order.slice(1, 4) };
  }

  // ------------------------------------------------------------------ money

  function toNum(s) {
    var n = parseFloat(String(s).replace(/[$,\s]/g, ""));
    return isNaN(n) ? 0 : n;
  }

  /**
   * Amounts labelled with something we recognise.
   *
   * Only labelled money is used. An unlabelled "$1,240.00" in a confirmation
   * could be the registration, the hotel total, or last year's price in a
   * comparison table — and guessing wrong writes a wrong number onto a
   * financial form. Better to leave the field blank and let the person type it.
   */
  /**
   * Capture groups in the caller's label pattern would shift the amount to an
   * unpredictable group number, and the failure is silent — every amount comes
   * back zero and the fields just look "not found". Neutralising the groups
   * keeps the amount at group 1 whatever the label looks like.
   */
  function noCapture(src) {
    return String(src).replace(/\((?!\?)/g, "(?:");
  }

  function findAmount(text, labelRx) {
    var t = clean(text);
    var label = noCapture(labelRx.source);
    var AMT = "([0-9][0-9,]*(?:\\.[0-9]{2})?)";
    var best = 0, snippet = "";
    // label ... $amount  (same line, or the next cell in a table row)
    var rx = new RegExp("(?:" + label + ")[^\\n$]{0,60}?\\$\\s*" + AMT, "gi");
    var m;
    while ((m = rx.exec(t))) {
      var v = toNum(m[1]);
      if (v > best) { best = v; snippet = m[0].replace(/\s+/g, " ").trim(); }
    }
    if (best) { return { value: best, snippet: snippet }; }
    // $amount ... label  (right-aligned columns put the number first)
    rx = new RegExp("\\$\\s*" + AMT + "[^\\n$]{0,40}?(?:" + label + ")", "gi");
    while ((m = rx.exec(t))) {
      var v2 = toNum(m[1]);
      if (v2 > best) { best = v2; snippet = m[0].replace(/\s+/g, " ").trim(); }
    }
    return { value: best, snippet: snippet };
  }

  var MONEY_FIELDS = [
    { field: "cRegistration", rx: /registration\s*(fee|cost|rate|total)?|conference\s*fee|attendee\s*fee|tuition|amount\s*(paid|charged|due)|total\s*(charged|paid|amount|due)|order\s*total|you\s*paid/ },
    { field: "cLodgingRate", rx: /(room|nightly|lodging|hotel)\s*rate|rate\s*(per|\/)\s*night|per\s*night|room\s*charge/ },
    { field: "cTravelMode", rx: /airfare|air\s*fare|flight\s*(cost|total|fare)?|ticket\s*(price|cost|total)/ },
    { field: "cParking", rx: /parking/ },
    { field: "cLuggage", rx: /bag(gage)?\s*fee|checked\s*bag|luggage/ },
  ];

  // ------------------------------------------------------------------- misc

  function findRole(text) {
    var t = clean(text);
    if (/\b(you\s+(are|will\s+be)\s+(a\s+)?)?(speak(er|ing)|present(er|ing)|panelis|panel\s+member|keynote)\b/i.test(t)) { return "speaker"; }
    if (/\bmoderat(or|ing)\b/i.test(t)) { return "moderator"; }
    if (/\bexhibitor|exhibit\s+booth\b/i.test(t)) { return "exhibitor"; }
    if (/\binstructor|trainer|facilitator\b/i.test(t)) { return "instructor"; }
    if (/\bregist(ered|ration)|attendee|participant\b/i.test(t)) { return "attendee"; }
    return "";
  }

  function findNights(text, start, end) {
    var m = /(\d{1,2})\s*(?:-|\s)?\s*night(?:s|\(s\))?\b/i.exec(clean(text));
    if (m) { return parseInt(m[1], 10); }
    if (start && end && end > start) {
      var a = new Date(start + "T00:00:00"), b = new Date(end + "T00:00:00");
      var n = Math.round((b - a) / 864e5);
      if (n > 0 && n < 30) { return n; }
    }
    return 0;
  }

  var LINK_CUE = /(agenda|event\s+(site|page|website)|conference\s+(site|website)|register|registration|more\s+information|details|program|venue|schedule|meeting)/i;

  function findLink(text) {
    var t = clean(text);
    var rx = /https?:\/\/[^\s<>")\]]+/g;
    var m, best = "", bestScore = -1e9;
    while ((m = rx.exec(t))) {
      var url = m[0].replace(/[.,;:)]+$/, "");
      if (/(unsubscribe|optout|opt-out|privacy|mailchimp|list-manage|\.png|\.jpg|\.gif|aka\.ms|safelinks)/i.test(url)) { continue; }
      var score = 0;
      // Scope both cues to lines. A label sits on the URL's own line or the
      // one above it; anything further away is a different part of the email,
      // and a footer 200 characters down should not veto the agenda link.
      var lineStart = t.lastIndexOf("\n", m.index) + 1;
      var prevStart = lineStart > 1 ? t.lastIndexOf("\n", lineStart - 2) + 1 : 0;
      var ownLine = t.slice(lineStart, m.index + url.length + 40);
      if (LINK_CUE.test(t.slice(prevStart, m.index + url.length + 40))) { score += 30; }
      // The word that disqualifies a link is usually beside it, not in it:
      // "Unsubscribe: https://x.com/u" has a perfectly innocent URL.
      if (/(unsubscribe|opt[\s-]?out|manage\s+(your\s+)?preferences|privacy\s+policy|view\s+in\s+browser|update\s+profile)/i.test(ownLine)) {
        score -= 80;
      }
      if (/^https?:\/\/[^/]*\/?$/.test(url)) { score += 10; }   // bare domain
      score -= url.length / 200;
      score -= m.index / 100000;
      if (score > bestScore) { bestScore = score; best = url; }
    }
    return best;
  }

  // ------------------------------------------------------ iCalendar (.ics)

  function unfoldIcs(text) {
    return String(text || "").replace(/\r\n?/g, "\n").replace(/\n[ \t]/g, "");
  }

  function icsDate(v) {
    var m = /^(\d{4})(\d{2})(\d{2})/.exec(String(v || "").trim());
    return m ? m[1] + "-" + m[2] + "-" + m[3] : "";
  }

  /**
   * An .ics attachment is the one part of a confirmation that is not prose —
   * the organiser has already told us the title, the start, the end and the
   * location as structured fields. When one is present it wins outright, and
   * everything the heuristics would have guessed is redundant.
   */
  function parseIcs(text) {
    var t = unfoldIcs(text);
    if (!/BEGIN:VEVENT/i.test(t)) { return null; }
    function field(name) {
      var rx = new RegExp("^" + name + "(?:;[^:\\n]*)?:(.*)$", "im");
      var m = rx.exec(t);
      return m ? m[1].replace(/\\,/g, ",").replace(/\\n/gi, "\n").replace(/\\\\/g, "\\").trim() : "";
    }
    var start = icsDate(field("DTSTART"));
    var end = icsDate(field("DTEND"));
    // DTEND on an all-day event is exclusive — the day after the last day.
    if (end && /^DTEND;VALUE=DATE/im.test(t) && end > start) {
      var d = new Date(end + "T00:00:00");
      d.setDate(d.getDate() - 1);
      end = iso(d.getFullYear(), d.getMonth() + 1, d.getDate());
    }
    return {
      summary: field("SUMMARY"), location: field("LOCATION"),
      start: start, end: end, url: field("URL"),
    };
  }

  // ------------------------------------------------------------- assemble

  function fmtLong(isoDate) {
    if (!isoDate) { return ""; }
    var p = isoDate.split("-");
    var d = new Date(+p[0], +p[1] - 1, +p[2]);
    if (isNaN(d.getTime())) { return ""; }
    return d.toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" });
  }

  function confDatesLine(start, end) {
    if (!start) { return ""; }
    if (!end || end === start) { return fmtLong(start); }
    var a = start.split("-"), b = end.split("-");
    var left = new Date(+a[0], +a[1] - 1, +a[2]);
    var right = new Date(+b[0], +b[1] - 1, +b[2]);
    if (a[0] === b[0] && a[1] === b[1]) {
      return left.toLocaleDateString("en-US", { month: "long", day: "numeric" }) +
        "-" + (+b[2]) + ", " + b[0];
    }
    if (a[0] === b[0]) {
      return left.toLocaleDateString("en-US", { month: "long", day: "numeric" }) +
        " - " + right.toLocaleDateString("en-US", { month: "long", day: "numeric" }) +
        ", " + b[0];
    }
    return fmtLong(start) + " - " + fmtLong(end);
  }

  function shiftDays(isoDate, n) {
    if (!isoDate) { return ""; }
    var p = isoDate.split("-");
    var d = new Date(+p[0], +p[1] - 1, +p[2]);
    if (isNaN(d.getTime())) { return ""; }
    d.setDate(d.getDate() + n);
    return iso(d.getFullYear(), d.getMonth() + 1, d.getDate());
  }

  /**
   * Everything above, applied to one message.
   *
   * input: {
   *   subject, body, receivedIso, homeCity,
   *   attachments: [{ name, text }]   // already decoded by the caller
   * }
   *
   * Returns { fields: {id: value}, sources: {id: "where it came from"},
   *           notes: [..], alternates: {id: [..]} }.
   *
   * `sources` is not decoration. The person is about to sign a form; being
   * able to see that "Registration $675" came from the line "Registration fee:
   * $675.00" in the body, and that the destination was a guess with two other
   * candidates, is the difference between a prefill they can trust and one
   * they have to re-check from scratch — which would be no saving at all.
   */
  function buildPrefill(input) {
    var inp = input || {};
    var subject = String(inp.subject || "");
    var bodyText = /<[a-z][\s\S]*>/i.test(inp.body || "")
      ? htmlToText(inp.body) : clean(inp.body || "");
    var atts = (inp.attachments || []).filter(function (a) { return a && a.text; });

    var fields = {}, sources = {}, alternates = {}, notes = [];
    function put(id, value, from) {
      if (value === "" || value === 0 || value == null) { return; }
      if (fields[id] != null) { return; }        // first (best) source wins
      fields[id] = value; sources[id] = from;
    }

    // Body first, then attachments, so the message itself outranks a generic
    // policy PDF that happens to mention a dollar figure.
    var pool = [{ name: "the email", text: bodyText }].concat(atts.map(function (a) {
      return { name: a.name, text: clean(a.text) };
    }));
    var allText = pool.map(function (p) { return p.text; }).join("\n\n");

    // --- an .ics beats every heuristic in here ---
    var ics = null;
    atts.forEach(function (a) {
      if (ics) { return; }
      var parsed = parseIcs(a.text);
      if (parsed && parsed.start) { ics = { data: parsed, name: a.name }; }
    });
    if (ics) {
      put("event", ics.data.summary, ics.name + " (calendar invite)");
      put("eventStart", ics.data.start, ics.name + " (calendar invite)");
      if (ics.data.location) {
        var lm = /([A-Z][A-Za-z.'-]+(?:\s+[A-Z][A-Za-z.'-]+){0,3})\s*,\s*([A-Z]{2})\b/.exec(ics.data.location);
        put("location", lm && STATES[lm[2]] ? lm[1] + ", " + lm[2] : ics.data.location,
          ics.name + " (calendar invite)");
      }
      put("confDates", confDatesLine(ics.data.start, ics.data.end), ics.name + " (calendar invite)");
      put("meetingLink", ics.data.url, ics.name + " (calendar invite)");
      notes.push("Used the calendar invite in " + ics.name + " — those fields are exact, not guessed.");
    }

    // --- event name ---
    var ev = eventFromSubject(subject);
    if (ev && ev.length >= 3) { put("event", ev, "the subject line"); }

    // --- dates ---
    var dates = { start: fields.eventStart || "", end: "" };
    if (!fields.eventStart) {
      pool.forEach(function (p) {
        if (dates.start) { return; }
        var d = eventDates(p.text, inp.receivedIso);
        if (d.start) { dates = d; put("eventStart", d.start, p.name); }
      });
    }
    if (dates.start) {
      put("confDates", confDatesLine(dates.start, dates.end), sources.eventStart || "the email");
      // Travel brackets the event by a day at each end — the usual shape, and
      // trivially adjusted. Marked as assumed so nobody mistakes it for stated.
      put("departDate", shiftDays(dates.start, -1), "assumed — the day before the event");
      put("returnDate", shiftDays(dates.end || dates.start, 1), "assumed — the day after the event");
    }

    // --- location ---
    if (!fields.location) {
      pool.forEach(function (p) {
        if (fields.location) { return; }
        var loc = findLocation(p.text, inp.homeCity);
        if (loc.value) {
          put("location", loc.value, p.name);
          if (loc.alternates.length) { alternates.location = loc.alternates; }
        }
      });
    }

    // --- money ---
    MONEY_FIELDS.forEach(function (spec) {
      pool.forEach(function (p) {
        if (fields[spec.field]) { return; }
        var hit = findAmount(p.text, spec.rx);
        if (hit.value) { put(spec.field, hit.value, p.name + ': "' + hit.snippet + '"'); }
      });
    });
    if (fields.cLodgingRate) {
      var nights = findNights(allText, dates.start, dates.end);
      if (nights) { put("cLodgingNights", nights, dates.end ? "the event dates" : "the email"); }
    }

    // --- role, link ---
    var role = findRole(bodyText);
    if (role) { put("attendeeRole", role, "the email"); }
    if (!fields.meetingLink) {
      var link = findLink(bodyText);
      if (link) { put("meetingLink", link, "a link in the email"); }
    }

    if (!fields.event) { notes.push("Couldn't tell what the event is called — the subject was all boilerplate."); }
    if (!fields.eventStart) { notes.push("No event date found. Confirmations that only say \"see attached\" usually need it typed in."); }
    if (!fields.location) { notes.push("No destination found."); }
    if (!fields.cRegistration) { notes.push("No labelled registration amount found — unlabelled dollar figures are left alone on purpose."); }

    return { fields: fields, sources: sources, alternates: alternates, notes: notes };
  }

  var api = {
    buildPrefill: buildPrefill,
    eventFromSubject: eventFromSubject,
    eventDates: eventDates,
    findDates: findDates,
    findLocation: findLocation,
    findAmount: findAmount,
    findRole: findRole,
    findNights: findNights,
    findLink: findLink,
    parseIcs: parseIcs,
    htmlToText: htmlToText,
    confDatesLine: confDatesLine,
    shiftDays: shiftDays,
    MONEY_FIELDS: MONEY_FIELDS,
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.TdMail = api; }
})(typeof self !== "undefined" ? self : this);
