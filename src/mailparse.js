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

  // ------------------------------------------------------------------ links

  /**
   * Undo the link rewriting that sits between a government mailbox and every
   * URL in it.
   *
   * Proofpoint URL Defense and Outlook Safe Links both replace the real
   * address with a 400-character wrapper of their own. Store that in "Meeting
   * / event link" and it is worthless to a human and worse in a printed
   * authorization form. Both wrappers carry the original inside them, so it
   * can simply be taken back out.
   *
   *   urldefense.proofpoint.com/v2/url?u=https-3A__x.org_a-2Db&d=...
   *   urldefense.com/v3/__https://x.org/a__;!!token$
   *   xx.safelinks.protection.outlook.com/?url=https%3A%2F%2Fx.org%2Fa&data=...
   */
  function unwrapUrl(url) {
    var u = String(url || "");
    for (var pass = 0; pass < 3; pass++) {
      var before = u;
      var m = /[?&]u=([^&]+)/.exec(u);
      if (/urldefense\.(proofpoint\.)?com\/v2\//i.test(u) && m) {
        // Proofpoint v2: "_" is "/" and "-XX" is a percent-escape with the
        // "%" written as "-". Restore both, then decode normally.
        var s = m[1].replace(/_/g, "/").replace(/-([0-9A-Fa-f]{2})/g, "%$1");
        try { u = decodeURIComponent(s); } catch (e) { u = s; }
      } else if (/urldefense\.com\/v3\//i.test(u)) {
        var v3 = /\/v3\/__(.+?)__;/.exec(u);
        if (v3) { try { u = decodeURIComponent(v3[1]); } catch (e) { u = v3[1]; } }
      } else if (/safelinks\.protection\.outlook\.com/i.test(u)) {
        var sl = /[?&]url=([^&]+)/.exec(u);
        if (sl) { try { u = decodeURIComponent(sl[1]); } catch (e) { u = sl[1]; } }
      }
      if (u === before) { break; }
    }
    return u.replace(/&amp;/g, "&");
  }

  /**
   * The links in an HTML email live in href attributes, which stripping tags
   * throws away — so scanning the visible text for a URL finds nothing at all
   * in a professionally built confirmation, where every link is a word like
   * "Reserve your room". Anchors are collected before the tags come off, each
   * with the words it was wrapped around, which is also the best available
   * label for deciding which link is the event.
   */
  function hrefs(html) {
    var out = [];
    var rx = /<a\b[^>]*?href\s*=\s*["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
    var m;
    while ((m = rx.exec(String(html || "")))) {
      var url = unwrapUrl(m[1].replace(/&amp;/g, "&").trim());
      if (!/^https?:/i.test(url)) { continue; }
      out.push({
        url: url,
        text: m[2].replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim(),
        at: m.index,
      });
    }
    return out;
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

  var LINK_JUNK = /(unsubscribe|optout|opt-out|privacy|mailchimp|list-manage|\.png|\.jpg|\.gif|aka\.ms|bookwithme|play\.google|itunes\.apple|microsoft\.com\/.*\/store|calendar\.yahoo|google\.com\/calendar|outlook\.(live|office)\.com|addtocalendar|forms\.gle)/i;

  /**
   * The event's own page, chosen from the anchors in the message.
   *
   * A confirmation is mostly links that are not the event: add-to-calendar for
   * five different calendars, app-store badges, the unsubscribe, a hotel
   * booking portal, a Google Form. What distinguishes the real one is that its
   * anchor text is the event's name — so the event name, once known, is the
   * strongest signal available, ahead of any keyword list.
   */
  function pickLink(anchors, eventName) {
    var ev = String(eventName || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
    var evWords = ev ? ev.split(" ").filter(function (w) { return w.length > 2; }) : [];
    var best = "", bestScore = -1e9;
    (anchors || []).forEach(function (a) {
      if (LINK_JUNK.test(a.url) || LINK_JUNK.test(a.text)) { return; }
      var score = 0;
      var txt = a.text.toLowerCase();
      if (evWords.length) {
        var hits = evWords.filter(function (w) { return txt.indexOf(w) >= 0; }).length;
        score += 60 * (hits / evWords.length);
      }
      if (LINK_CUE.test(a.text)) { score += 25; }
      if (/^https?:\/\/[^/]+\/?$/.test(a.url)) { score += 8; }
      score -= a.url.length / 300;
      score -= a.at / 200000;
      if (score > bestScore) { bestScore = score; best = a.url; }
    });
    return bestScore > 5 ? best : "";
  }

  // ------------------------------------------------------------------ times

  var TIME = "(\\d{1,2})(?::(\\d{2}))?\\s*([ap])\\.?m\\.?";

  /**
   * "8:00 AM - 5:00 PM" out of the message, for the free-text dates field.
   *
   * The field's own placeholder is "Jan 10-14, 8am-5pm", so the times are part
   * of what it is asking for, and a confirmation almost always states them.
   * A lone time is ignored — "registration opens at 7:00 am" is not the event
   * running from 7am to nothing.
   */
  function findTimes(text) {
    var t = clean(text);
    var rx = new RegExp(TIME + "\\s*(?:-|to|until|till|through)\\s*" + TIME, "gi");
    var m, best = null, bestScore = -1e9;
    while ((m = rx.exec(t))) {
      // Scoped to the line the times sit on. A wider window sees the whole
      // paragraph, so "Registration opens 7-8am" and "the event runs 8-5" on
      // consecutive lines score identically and the wrong one wins on
      // position alone.
      var lineStart = t.lastIndexOf("\n", m.index) + 1;
      var lineEnd = t.indexOf("\n", m.index);
      var line = t.slice(lineStart, lineEnd < 0 ? t.length : lineEnd);
      var score = 0;
      if (/(time|event|conference|session|runs?|agenda|schedule|program|all\s+day)/i.test(line)) { score += 30; }
      if (/(registration|check-?in|badge|breakfast|reception|desk\s+opens|doors)/i.test(line)) { score -= 45; }
      score -= m.index / 100000;
      if (score > bestScore) { bestScore = score; best = m; }
    }
    if (!best) { return ""; }
    function fmt(h, mm, ap) {
      return String(parseInt(h, 10)) + (mm && mm !== "00" ? ":" + mm : "") +
        (ap.toLowerCase() === "a" ? "am" : "pm");
    }
    return fmt(best[1], best[2], best[3]) + "-" + fmt(best[4], best[5], best[6]);
  }

  // -------------------------------------------------- third-party payer

  var REIMB_CUE = /(reimburs|invitational\s+travel|at\s+no\s+cost\s+to|covered\s+by|will\s+pay|directly\s+pay|sponsor(ed|ship)?\s+(your\s+)?travel|we\s+will\s+cover|travel\s+(is\s+)?(fully\s+)?(funded|paid))/i;

  var ACRONYM_STOP = {
    CAUTION: 1, MUST: 1, NOTE: 1, PDH: 1, FYI: 1, RSVP: 1, HTML: 1, PDF: 1,
    USA: 1, AM: 1, PM: 1, THE: 1, AND: 1, FOR: 1, YOU: 1, NOT: 1, ALL: 1,
    NEW: 1, ARE: 1, WILL: 1, FROM: 1, WITH: 1, YOUR: 1, THIS: 1, THAT: 1,
    INC: 1, LLC: 1, ONLY: 1, FREE: 1,
  };

  /**
   * Who else is paying, and for what.
   *
   * On invitational or sponsored travel this is the most consequential thing
   * in the message and the most tedious to re-key: the categories are spelled
   * out in a sentence, the billing contact is three paragraphs further down,
   * and getting it wrong means the receivable never gets raised.
   *
   * The entity is the shakiest part, so it is derived conservatively — an
   * acronym that appears near the reimbursement language AND more than once in
   * the message, which is what distinguishes a sponsor from a passing
   * reference. Everything else here is quoted directly from the text.
   */
  function findReimbursement(text) {
    var t = clean(text);
    if (!REIMB_CUE.test(t)) { return null; }

    var cats = {
      reg: /\bregistration\s+(fee|cost)?s?\b[^.]{0,80}reimburs|reimburs[^.]{0,120}\bregistration\b/i.test(t),
      lodging: /(hotel|lodging|room|accommodation)[^.]{0,120}(reimburs|cover|paid|pay)|(reimburs|cover|pay)[^.]{0,120}(hotel|lodging|room)/i.test(t),
      air: /(airfare|air\s+fare|flight|rail|rental\s+car|checked\s+bag|baggage)[^.]{0,140}(reimburs|cover|paid|pay|arrange)|(reimburs|cover|pay|arrange|directly\s+pay)[^.]{0,160}(airfare|rail|rental\s+car|baggage)/i.test(t),
      meals: /(meal|per\s+diem|m&ie)[^.]{0,120}(allowance|reimburs|cover|paid)|(reimburs|cover|pay)[^.]{0,140}(meal|per\s+diem)/i.test(t),
      ground: /(local\s+transport|parking|rideshare|taxi|mileage|ground\s+transport)/i.test(t),
    };

    // The billing contact: a mailbox named where reimbursement or questions
    // are discussed. A shared address beats a person's - it survives them
    // changing jobs, which is exactly when a receivable goes unchased.
    var mails = emailsIn(t), mm;
    var cueAt = [];
    var crx = new RegExp(REIMB_CUE.source + "|questions?|contact", "gi");
    while ((mm = crx.exec(t))) { cueAt.push(mm.index); }
    var contact = "";
    var bestD = 1e9;
    mails.forEach(function (e) {
      if (/(noreply|no-reply|donotreply|notify|unsubscribe|@eventleaf|@mailchimp)/i.test(e.addr)) { return; }
      cueAt.forEach(function (c) {
        var d = Math.abs(c - e.at);
        var shared = !/\./.test(e.addr.split("@")[0]) ? 400 : 0;   // innovation@ beats jane.doe@
        if (d - shared < bestD) { bestD = d - shared; contact = e.addr; }
      });
    });

    // Entity: an acronym used near the money language and more than once.
    var counts = {}, arx = /\b([A-Z]{2,6})\b(?!-\d)/g;
    while ((mm = arx.exec(t))) {
      if (ACRONYM_STOP[mm[1]]) { continue; }
      counts[mm[1]] = (counts[mm[1]] || 0) + 1;
    }
    var entity = "", entScore = 0;
    arx.lastIndex = 0;
    while ((mm = arx.exec(t))) {
      var tok = mm[1];
      if (ACRONYM_STOP[tok] || counts[tok] < 2) { continue; }
      var nearCue = cueAt.some(function (c) { return Math.abs(c - mm.index) < 700; });
      if (!nearCue) { continue; }
      if (counts[tok] > entScore) { entScore = counts[tok]; entity = tok; }
    }

    var any = cats.reg || cats.lodging || cats.air || cats.meals || cats.ground;
    if (!any && !contact) { return null; }
    return {
      entity: entity, contact: contact, categories: cats,
      snippet: quoteSentence(t, /reimburs\w*/i) ||
        quoteSentence(t, /directly\s+pay|covered\s+by|at\s+no\s+cost\s+to/i) ||
        quoteSentence(t, REIMB_CUE),
    };
  }

  /**
   * The sentence a match sits in, trimmed to something quotable.
   *
   * Taking a fixed window around the first cue quotes whatever happens to be
   * next to it — here, a section heading two paragraphs above the sentence
   * that actually promises the money. The source line is the whole point of
   * showing it, so it has to be the sentence a person can find in the email.
   */
  var COST_WORD = /(hotel|lodging|room|airfare|rail|rental|baggage|bag|meal|per\s+diem|parking|rideshare|mileage|transport|registration|expense)/gi;

  function quoteSentence(text, rx) {
    var t = String(text);
    var g = new RegExp(rx.source, rx.flags.indexOf("g") >= 0 ? rx.flags : rx.flags + "g");
    var m, best = null, bestScore = -1;
    while ((m = g.exec(t))) {
      // Several sentences may promise reimbursement; the useful one is the
      // one that lists what is covered, not the one that says it exists.
      var win = t.slice(m.index, m.index + 200);
      var n = (win.match(COST_WORD) || []).length;
      if (n > bestScore) { bestScore = n; best = m; }
      if (g.lastIndex === m.index) { g.lastIndex++; }
    }
    if (!best) { return ""; }
    m = best;
    var start = Math.max(
      t.lastIndexOf(". ", m.index) + 1,
      t.lastIndexOf("\n", m.index) + 1,
      t.lastIndexOf(": ", m.index) + 1
    );
    var stop = t.length;
    [". ", "\n"].forEach(function (d) {
      var i = t.indexOf(d, m.index);
      if (i >= 0 && i < stop) { stop = i + (d === ". " ? 1 : 0); }
    });
    var s = t.slice(start, Math.min(stop, start + 160)).replace(/\s+/g, " ").trim();
    return s.length > 150 ? s.slice(0, 147) + "…" : s;
  }

  // ------------------------------------------------------- your own details

  /**
   * Your division, bureau and home city, read out of your own signature block.
   *
   * These are the fields a confirmation email never contains, because they are
   * facts about you rather than the trip — and they are also the ones a person
   * is most irritated to retype. They do appear, in every message you have
   * ever forwarded, in the signature underneath.
   *
   *   Matthew Miller, CPM
   *   Director of New and Emerging Transportation Technologies
   *   Systems Operations Division      <- division
   *   Iowa Department of Transportation
   *   800 Lincoln Way
   *   Ames, IA 50010                   <- home city
   *
   * It is anchored on YOUR name and looks only at the handful of lines below
   * it. A confirmation carries the sender's signature too — "FHWA | Office of
   * Infrastructure" — and without the anchor this would confidently report
   * that you work for the people who invited you.
   */
  function findSignature(text, myName) {
    var name = String(myName || "").trim();
    if (!name) { return null; }
    var t = clean(text);
    var lines = t.split("\n");

    // Match on surname too: signatures alternate "Matthew Miller" and
    // "Miller, Matthew" within a single forwarded thread.
    var parts = name.split(/[\s,]+/).filter(Boolean);
    var last = parts.length ? parts[parts.length - 1] : "";
    if (/^(jr|sr|ii|iii|iv|cpm|pe|phd|aicp)\.?$/i.test(last) && parts.length > 1) {
      last = parts[parts.length - 2];
    }
    var first = parts[0] || "";

    var out = { division: "", bureau: "", city: "" };
    for (var i = 0; i < lines.length; i++) {
      var L = lines[i].trim();
      var isMe = last && L.length < 60 &&
        new RegExp("\\b" + last.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\b", "i").test(L) &&
        (!first || new RegExp("\\b" + first.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\b", "i").test(L));
      if (!isMe) { continue; }
      for (var j = i + 1; j < Math.min(lines.length, i + 9); j++) {
        var s = lines[j].trim().replace(/[|•·]+/g, " ").replace(/\s{2,}/g, " ").trim();
        if (!s) { continue; }
        if (/^(from|sent|to|cc|subject|date)\s*:/i.test(s)) { break; }   // next header block
        var m;
        if (!out.division && (m = /^(.{2,45}?)\s+Division$/i.exec(s))) { out.division = m[1].trim(); }
        if (!out.bureau && (m = /^(?:Bureau\s+of\s+)(.{2,45})$/i.exec(s))) { out.bureau = m[1].trim(); }
        if (!out.bureau && (m = /^(.{2,45}?)\s+Bureau$/i.exec(s))) { out.bureau = m[1].trim(); }
        if (!out.bureau && (m = /^Office\s+of\s+(.{2,45})$/i.exec(s))) { out.bureau = m[1].trim(); }
        if (!out.city && (m = /^([A-Z][A-Za-z .'-]{1,30}),\s*([A-Z]{2})\s+\d{5}(?:-\d{4})?$/.exec(s)) &&
            STATES[m[2]]) { out.city = m[1].trim(); }
      }
      if (out.division || out.bureau || out.city) { break; }
    }
    return (out.division || out.bureau || out.city) ? out : null;
  }

  /**
   * "Please return this form ... in email to Keri.Greenfield@iowadot.us"
   *
   * Every organization's travel form names the person it goes back to, and a
   * new traveler otherwise has to be told that address by someone. Worth
   * offering — never applying silently, because the coordinator is a setting
   * that outlives this one trip.
   */
  function findCoordinator(text, myDomain) {
    var t = clean(text);
    var dom = String(myDomain || "").toLowerCase();
    var best = "", bestScore = 0;
    emailsIn(t).forEach(function (e) {
      if (/(noreply|no-reply|donotreply|notify|unsubscribe)/i.test(e.addr)) { return; }
      var before = t.slice(Math.max(0, e.at - 140), e.at);
      var score = 0;
      // "return this FORM to X" is the coordinator. A bare "email X" is
      // whoever runs the event, and addressing your authorization paperwork
      // to the conference organizer is worse than leaving it unset.
      if (/\b(return|submit|forward|send)\b[^.]{0,60}\bform\b/i.test(before)) { score += 60; }
      else if (/\b(return|submit)\b/i.test(before)) { score += 30; }
      else if (/\b(send|forward)\b/i.test(before)) { score += 12; }
      else if (/\b(email|e-mail|reply)\b/i.test(before)) { score += 6; }
      else { return; }
      // Your travel coordinator works where you work.
      if (dom && e.addr.toLowerCase().indexOf("@" + dom) > 0) { score += 45; }
      else if (dom) { score -= 25; }
      if (score > bestScore) { bestScore = score; best = e.addr; }
    });
    // 10 clears a plain "send it to X" but not a bare "email X" at an outside
    // domain, which is the event's organizer rather than your coordinator.
    return bestScore >= 10 ? best : "";
  }

  var MAIL_STRICT = /\b([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})\b/g;
  // Text lifted out of a PDF carries the letter-spacing the layout used, so a
  // perfectly ordinary address arrives as "Keri . Greenfield @iowadot.us" and
  // the strict pattern misses it entirely. The loose pattern is only tried
  // when the strict one found nothing at all — on clean text it could staple
  // a sentence-ending full stop onto the address in front of it.
  var MAIL_LOOSE = /([A-Za-z0-9_%+-]+(?:\s*\.\s*[A-Za-z0-9_%+-]+)*\s*@\s*[A-Za-z0-9-]+(?:\s*\.\s*[A-Za-z0-9-]+)+)/g;

  function emailsIn(text) {
    var t = String(text), out = [], m;
    MAIL_STRICT.lastIndex = 0;
    while ((m = MAIL_STRICT.exec(t))) { out.push({ addr: m[1], at: m.index }); }
    if (out.length) { return out; }
    MAIL_LOOSE.lastIndex = 0;
    while ((m = MAIL_LOOSE.exec(t))) {
      var addr = m[1].replace(/\s+/g, "");
      if (/\.[A-Za-z]{2,}$/.test(addr)) { out.push({ addr: addr, at: m.index }); }
    }
    return out;
  }

  // ------------------------------------------------- a filled-in paper form

  /**
   * Labels, tolerant of the letter-spacing a PDF leaves behind.
   *
   * The real Iowa DOT form extracts as "Name & C ost Center:" — the layout's
   * kerning becomes literal spaces, and an exact match finds nothing. Allowing
   * optional whitespace between every character costs nothing and makes the
   * labels matchable however the generator spaced them.
   */
  function looseLabel(s) {
    return String(s).split("").map(function (c) {
      if (/\s/.test(c)) { return "\\s*"; }
      if (/[A-Za-z0-9]/.test(c)) { return c + "\\s*"; }
      return "\\" + c + "\\s*";
    }).join("");
  }

  var FORM_LABELS = [
    { label: "Name & Cost Center", field: "_nameCost" },
    { label: "Cost Center", field: "costCenter" },
    { label: "Other Staff Attending", field: "otherStaff" },
    { label: "Name of Conference", field: "event" },
    { label: "Conference Dates & Times", field: "confDates" },
    { label: "Departure Date", field: "departDate" },
    { label: "Return Date", field: "returnDate" },
    { label: "Reason for Travel", field: "reason" },
    { label: "Cost of Travel Mode", field: "cTravelMode" },
    { label: "Luggage Fees", field: "cLuggage" },
    { label: "Registration Fee", field: "cRegistration" },
    { label: "Taxi/Uber Fees", field: "cTaxi" },
    { label: "Additional Fees", field: "cAdditional" },
    { label: "Project Number", field: "tp1Project" },
    { label: "Maximum Reimbursement Amount", field: "tp1Max" },
    { label: "Location", field: "location" },
    { label: "Parking", field: "cParking" },
  ];

  var MONEY_FIELD = /^(cTravelMode|cLuggage|cParking|cRegistration|cTaxi|cAdditional|tp1Max)$/;
  var DATE_FIELD = /^(departDate|returnDate)$/;

  /**
   * Read a filled-in travel form that somebody attached.
   *
   * The blank template is all underscores and yields nothing, which is
   * correct. But the same form comes back filled — a colleague sends theirs
   * for the planner, or you complete the paper version first — and everything
   * on it maps onto a field here. A value that is still just underscores or
   * empty is skipped, so a half-completed form contributes the half that was
   * filled and stays silent about the rest.
   */
  function formFields(text) {
    var t = clean(text).replace(/_{2,}/g, " ");   // erase the blank rules
    var out = {};

    // A value runs until the NEXT LABEL, full stop. Relying on the run of
    // spaces the form puts between fields does not survive clean(), which
    // collapses whitespace - so every value swallowed the entire rest of the
    // form. The optional parenthetical is needed because the real template
    // writes "Cost of Travel Mode (miles, estimated flight cost):".
    var PAREN = "\\s*(?:\\([^)]{0,60}\\))?\\s*";
    var STOP = "(?:" +
      FORM_LABELS.map(function (o) { return looseLabel(o.label); }).join("|") +
      "|" + looseLabel("Lodging") +
      "|" + looseLabel("Mode of Travel") +
      "|" + looseLabel("Reimbursement Items") +
      "|" + looseLabel("Reimbursement Notes") +
      ")" + PAREN + "[:\\-]";

    FORM_LABELS.forEach(function (spec) {
      var rx = new RegExp(looseLabel(spec.label) + PAREN + "[:\\-]\\s*([\\s\\S]{0,140}?)(?=" +
        STOP + "|\\n|$)", "i");
      var m = rx.exec(t);
      if (!m) { return; }
      var v = m[1].replace(/^[\s$]+/, "").replace(/[\s.:;,\-=]+$/, "").trim();
      if (!v || v.length < 2) { return; }
      out[spec.field] = v;
    });

    // Only trust this if the text really is a form. "Location" and "Parking"
    // are ordinary words, and a calendar invite's "LOCATION:" line matched one
    // of them and overruled the invite's own parser. Several labels together
    // are a form; one on its own is a coincidence.
    if (Object.keys(out).length < 3) { return {}; }

    // "Name & Cost Center: Matthew Miller, 471-0000"
    if (out._nameCost) {
      // The combined label owns both halves. The standalone "Cost Center"
      // pattern also matches inside "Name & Cost Center:", so leaving its
      // result in place puts the person's name in the cost-center box.
      //
      // Split on the CODE, not on the last comma: "Miller, Matthew 300000" is
      // written surname-first as often as not, and taking everything after the
      // final comma makes the cost center "Matthew 300000".
      var v = out._nameCost.trim();
      var code = /(?:^|[\s,;:])([0-9][0-9.\-]*[0-9])\s*$/.exec(v);
      if (code) {
        out.costCenter = code[1];
        out.name = v.slice(0, code.index).replace(/[\s,;:\-]+$/, "").trim();
      } else {
        var parts = v.split(/\s*[,;]\s*|\s+-\s+/);
        if (parts.length > 1) {
          out.costCenter = parts[parts.length - 1].trim();
          out.name = parts.slice(0, -1).join(", ").trim();
        } else {
          out.name = v;
          if (out.costCenter === v) { delete out.costCenter; }
        }
      }
      if (!out.name) { delete out.name; }
      delete out._nameCost;
    }

    // "Lodging: 2 nights @ $150.00 = $300.00"
    var lodge = new RegExp(looseLabel("Lodging") + "[:\\-]?\\s*(\\d{1,2})\\s*n\\s*i\\s*g\\s*h\\s*t", "i").exec(t);
    if (lodge) { out.cLodgingNights = parseInt(lodge[1], 10); }
    var rate = /night[^$\n]{0,20}\$\s*([0-9][0-9,]*(?:\.\d{2})?)/i.exec(t);
    if (rate) { out.cLodgingRate = toNum(rate[1]); }

    Object.keys(out).forEach(function (k) {
      if (MONEY_FIELD.test(k)) {
        var n = toNum(String(out[k]).replace(/^\$/, ""));
        if (n) { out[k] = n; } else { delete out[k]; }
      } else if (DATE_FIELD.test(k)) {
        var d = findDates(String(out[k]));
        if (d.length) { out[k] = d[0].start; } else { delete out[k]; }
      }
    });
    return out;
  }

  /** "Code: MV59BR3A" / "Confirmation number 4820193" — worth keeping. */
  function findCode(text) {
    var t = clean(text);
    var rx = /\b(?:confirmation|registration|reference|booking|event|attendee)?\s*(?:code|number|#|no\.?|id)\s*:?\s*\*?\*?\s*([A-Z0-9][A-Z0-9-]{5,17})\b/gi;
    var m;
    while ((m = rx.exec(t))) {
      var c = m[1];
      if (!/\d/.test(c) || !/[A-Z]/i.test(c)) { continue; }   // needs both
      if (/^\d{5}(-\d{4})?$/.test(c)) { continue; }           // a ZIP code
      return c;
    }
    return "";
  }

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
    var isHtml = /<[a-z][\s\S]*>/i.test(inp.body || "");
    var bodyText = isHtml ? htmlToText(inp.body) : clean(inp.body || "");
    var anchors = isHtml ? hrefs(inp.body) : [];
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

    // --- a filled-in travel form beats every heuristic in here ---
    // Somebody typed these values into the fields on purpose. Nothing derived
    // from prose should be allowed to overrule that.
    atts.forEach(function (a) {
      var ff = formFields(a.text);
      Object.keys(ff).forEach(function (k) { put(k, ff[k], a.name + " (the filled-in form)"); });
    });

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
      var times = "";
      pool.forEach(function (p) { if (!times) { times = findTimes(p.text); } });
      var line = confDatesLine(dates.start, dates.end) + (times ? ", " + times : "");
      put("confDates", line, (sources.eventStart || "the email") + (times ? " + the stated times" : ""));
      // Travel brackets the event by a day at each end — the usual shape, and
      // trivially adjusted. Marked as assumed so nobody mistakes it for stated.
      put("departDate", shiftDays(dates.start, -1), "assumed — the day before the event");
      put("returnDate", shiftDays(dates.end || dates.start, 1), "assumed — the day after the event");
    }

    // --- your own details, from your signature block ---
    var sig = null;
    pool.forEach(function (p) { if (!sig) { sig = findSignature(p.text, inp.myName); } });
    if (sig) {
      put("division", sig.division, "your signature block");
      put("bureau", sig.bureau, "your signature block");
    }
    // Knowing where you are based is what stops your own office being read as
    // the destination, and the signature states it.
    var home = inp.homeCity || (sig && sig.city) || "";

    // --- location ---
    if (!fields.location) {
      pool.forEach(function (p) {
        if (fields.location) { return; }
        var loc = findLocation(p.text, home);
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
      // Nights follow the TRAVEL dates, not the event's. A one-day event you
      // fly in for is two hotel nights, and taking the event's own span gives
      // zero — which then multiplies the room rate out to nothing.
      var nights = findNights(allText, dates.start, dates.end);
      if (!nights && fields.departDate && fields.returnDate) {
        var span = Math.round(
          (new Date(fields.returnDate + "T00:00:00") -
           new Date(fields.departDate + "T00:00:00")) / 864e5);
        if (span > 0 && span < 30) { nights = span; }
      }
      if (nights) { put("cLodgingNights", nights, "the travel dates"); }
    }

    // --- role, link ---
    var role = findRole(bodyText);
    if (role) { put("attendeeRole", role, "the email"); }
    if (!fields.meetingLink) {
      // Anchors first: in a real confirmation every link is a phrase, so the
      // URL never appears in the visible text at all.
      var link = pickLink(anchors, fields.event) || findLink(bodyText);
      if (link) {
        put("meetingLink", link,
          /urldefense|safelinks/i.test(link) ? "a link in the email"
            : "a link in the email" + (anchors.length ? " (unwrapped from the mail filter)" : ""));
      }
    }

    // --- who else is paying ---
    var reimb = null;
    pool.forEach(function (p) { if (!reimb) { reimb = findReimbursement(p.text); } });
    if (reimb) {
      var where = 'the email: "' + reimb.snippet + '"';
      put("tp1Name", reimb.entity, reimb.entity ? where : "");
      put("tp1Contact", reimb.contact, where);
      var boxes = {
        tp1Reg: reimb.categories.reg, tp1Lodging: reimb.categories.lodging,
        tp1Air: reimb.categories.air, tp1Meals: reimb.categories.meals,
        tp1Ground: reimb.categories.ground,
      };
      Object.keys(boxes).forEach(function (id) { if (boxes[id]) { put(id, true, where); } });
      notes.push("Someone else is paying for part of this trip — check " +
        "“Third-party reimbursement #1”, and put the amount in " +
        "“Max reimbursement” if the email states one." +
        (reimb.entity ? "" : " It couldn't tell who the payer is; type their name."));
    }

    // --- registration code, kept as a planner comment ---
    var code = findCode(bodyText);
    if (code) { put("comments", "Registration code " + code, "the email"); }

    // --- offered, never applied: settings outlive this one trip ---
    var suggest = {};
    var coord = "";
    // Attachments first here, unlike everywhere else: the travel FORM names
    // the coordinator, while the event email names the event's organizer.
    pool.slice(1).concat(pool[0]).forEach(function (p) {
      if (!coord) { coord = findCoordinator(p.text, inp.myDomain); }
    });
    if (coord) { suggest.coordEmail = coord; }
    if (home) { suggest.homeCity = home; }

    if (!fields.event) { notes.push("Couldn't tell what the event is called — the subject was all boilerplate."); }
    if (!fields.eventStart) { notes.push("No event date found. Confirmations that only say \"see attached\" usually need it typed in."); }
    if (!fields.location) { notes.push("No destination found."); }
    if (!fields.cRegistration) { notes.push("No labelled registration amount found — unlabelled dollar figures are left alone on purpose."); }

    return {
      fields: fields, sources: sources, alternates: alternates,
      notes: notes, suggest: suggest,
    };
  }

  var api = {
    buildPrefill: buildPrefill,
    eventFromSubject: eventFromSubject,
    eventDates: eventDates,
    findDates: findDates,
    findLocation: findLocation,
    findAmount: findAmount,
    findRole: findRole,
    findSignature: findSignature,
    findCoordinator: findCoordinator,
    formFields: formFields,
    findNights: findNights,
    findLink: findLink,
    findTimes: findTimes,
    findReimbursement: findReimbursement,
    findCode: findCode,
    unwrapUrl: unwrapUrl,
    hrefs: hrefs,
    pickLink: pickLink,
    parseIcs: parseIcs,
    htmlToText: htmlToText,
    confDatesLine: confDatesLine,
    shiftDays: shiftDays,
    MONEY_FIELDS: MONEY_FIELDS,
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.TdMail = api; }
})(typeof self !== "undefined" ? self : this);
