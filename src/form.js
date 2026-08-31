/*
 * Travel Desk — pure form logic (no Office/Graph dependencies).
 *
 * Turns one filled-out trip into BOTH artifacts the Iowa DOT process needs:
 *   1. a printable Travel Authorization form (HTML, mirrors the PDF) that
 *      goes to the travel coordinator as an email draft, and
 *   2. a row for the Division out-of-state travel planner workbook,
 *      mapped by header name so year-to-year template drift doesn't break it.
 *
 * Deterministic — no AI. Works in the browser (global `TravelForm`) and in
 * Node (module.exports) so the tests run offline.
 */
(function (root) {
  "use strict";

  function num(v) {
    var n = parseFloat(String(v == null ? "" : v).replace(/[$,\s]/g, ""));
    return isNaN(n) ? 0 : n;
  }

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function money(n) {
    return "$" + (Math.round(n * 100) / 100).toLocaleString("en-US",
      { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  /** Lodging subtotal and grand total across all cost lines. */
  function computeTotals(model) {
    var c = model.costs || {};
    var lodging = num(c.lodgingNights) * num(c.lodgingRate);
    var grand = num(c.travelMode) + num(c.luggage) + num(c.parking) + lodging +
      num(c.registration) + num(c.taxi) + num(c.additional);
    return { lodging: lodging, grand: grand };
  }

  /** "Travel Auth - Miller - TRB Annual Meeting - 2027-01-10" */
  function subjectLine(model) {
    var name = String(model.name || "").trim();
    var last = name.split(/[\s,]+/).filter(Boolean);
    // "Miller, Matt" → Miller; "Matt Miller" → Miller
    var lastName = /,/.test(name) ? (last[0] || "") : (last[last.length - 1] || "");
    var parts = ["Travel Auth", lastName, String(model.event || "").trim()];
    if (model.eventStart) { parts.push(model.eventStart); }
    return parts.filter(Boolean).join(" - ");
  }

  /**
   * Fiscal-year label for a date, for any org's fiscal calendar.
   * startMonth 1 = calendar year ("FY27" = 2027); otherwise the fiscal year
   * is numbered by its ENDING year (start 7: Jul 2026–Jun 2027 = 27; start
   * 10 = US federal). Prefix is the org's convention ("FY", "SFY", "FFY").
   */
  function fiscalLabel(isoDate, startMonth, prefix) {
    var d = new Date(String(isoDate || "") + "T00:00:00");
    if (isNaN(d)) { return ""; }
    var start = Number(startMonth) || 1;
    var y = d.getFullYear(), m = d.getMonth() + 1;
    var fy = start === 1 ? y : (m >= start ? y + 1 : y);
    return (prefix == null || prefix === "" ? "FY" : prefix) + String(fy).slice(-2);
  }

  /** Summary of third-party reimbursement for the planner column. */
  function thirdPartySummary(model) {
    var tps = (model.thirdParties || []).filter(function (t) { return t && (t.name || "").trim(); });
    if (!tps.length) { return "No"; }
    return "Yes - " + tps.map(function (t) { return t.name.trim(); }).join(", ");
  }

  /** Estimated reimbursed fraction (0–1, 2dp) like the planner uses. */
  function reimbursePct(model) {
    var grand = computeTotals(model).grand;
    if (!grand) { return ""; }
    var covered = (model.thirdParties || []).reduce(function (s, t) {
      return s + num(t && t.maxReimb);
    }, 0);
    if (!covered) { return ""; }
    var f = Math.min(1, covered / grand);
    return String(Math.round(f * 100) / 100);
  }

  /**
   * Map the trip onto a planner table row by HEADER NAME (fuzzy, lowercase
   * contains) so the add-in survives template drift between fiscal years —
   * and works against ANY org's planner columns. Unknown headers get ""
   * (never invents data). opts: {fyStartMonth, fyPrefix} per org settings.
   */
  function plannerRow(headers, model, opts) {
    opts = opts || {};
    var totals = computeTotals(model);
    var fy = fiscalLabel(model.eventStart || model.departDate, opts.fyStartMonth, opts.fyPrefix);
    var funding = model.funding || model.tewd || "";
    return (headers || []).map(function (h) {
      var k = String(h || "").toLowerCase();
      if (!k) { return ""; }
      // These are filled later in the trip's life (close-out and approval),
      // never when the request is first filed. They must be tested FIRST:
      // "Approved date" contains "date" and "Actual cost" contains "cost",
      // so the generic tests below would otherwise claim them.
      if (k.indexOf("actual") !== -1) { return ""; }
      if (k.indexOf("approved") !== -1) { return ""; }
      if (k.indexOf("event") !== -1 || k.indexOf("conference") !== -1) { return model.event || ""; }
      if (k.indexOf("city") !== -1 || k.indexOf("location") !== -1 || k.indexOf("destination") !== -1) { return model.location || ""; }
      if (k.indexOf("start") !== -1 || k.indexOf("date") !== -1) { return model.eventStart || model.departDate || ""; }
      if (k.indexOf("division") !== -1 || k.indexOf("department") !== -1) { return model.division || ""; }
      if (k.indexOf("bureau") !== -1 || k.indexOf("office") !== -1 || k.indexOf("team") !== -1) { return model.bureau || ""; }
      if ((k.indexOf("attendee") !== -1 || k.indexOf("traveler") !== -1 || k.indexOf("name") !== -1) && k.indexOf("role") === -1) { return model.name || ""; }
      if (k.indexOf("role") !== -1) { return model.attendeeRole || ""; }
      if (k.indexOf("cost") !== -1 || k.indexOf("estimate") !== -1) {
        var amt = opts.costOverride != null ? opts.costOverride : totals.grand;
        return amt ? String(Math.round(amt)) : "";
      }
      // "% Reimbursed" must be tested BEFORE the third-party names below,
      // or the shared word "reimburs" steals the percentage column.
      if (k.indexOf("%") !== -1) { return reimbursePct(model); }
      // real planners spell it "3rd Party" as often as "Third party"
      if (k.indexOf("third") !== -1 || k.indexOf("3rd") !== -1 ||
          k.indexOf("grant") !== -1 || k.indexOf("reimburs") !== -1) { return thirdPartySummary(model); }
      if (k.indexOf("fiscal") !== -1) { return fy; }
      if (k.indexOf("tewd") !== -1 || k.indexOf("funding") !== -1 || k.indexOf("program") !== -1) { return funding; }
      if (k.indexOf("comment") !== -1 || k.indexOf("note") !== -1) { return model.comments || ""; }
      if (k.indexOf("status") !== -1 || k.indexOf("booked") !== -1) { return "Requested"; } // coordinator's follow-up queue
      if (k.indexOf("coo") !== -1 || k.indexOf("approv") !== -1) { return ""; } // approver's column, never ours
      return "";
    });
  }

  function box(v) { return v ? "&#9746;" : "&#9744;"; } // ☒ / ☐

  function fieldRow(label, value) {
    return '<tr><td class="lbl">' + esc(label) + '</td><td class="val">' + esc(value || "") + "</td></tr>";
  }

  function thirdPartyHtml(t, ordinal) {
    if (!t || !(t.name || "").trim()) { return ""; }
    var items = t.items || {};
    return '<h3>' + ordinal + ' 3rd Party Entity</h3><table class="fields">' +
      fieldRow("Name", t.name) +
      fieldRow("Project Number", t.project) +
      fieldRow("3rd party packet to attach?", t.packet ? "Yes" : "No") +
      fieldRow("Maximum Reimbursement Amount", t.maxReimb ? money(num(t.maxReimb)) : "") +
      "</table><p class=\"items\">Reimbursement items: " +
      box(items.registration) + " Registration Fees &nbsp; " +
      box(items.lodging) + " Lodging/Hotel &nbsp; " +
      box(items.airfare) + " Airfare/Luggage &nbsp; " +
      box(items.meals) + " Meals &nbsp; " +
      box(items.ground) + " Ground Transportation</p>" +
      (t.notes ? "<p><strong>Reimbursement notes:</strong> " + esc(t.notes) + "</p>" : "");
  }

  /**
   * Printable Travel Authorization form. Field set mirrors the common
   * travel-auth shape (originally modeled on Iowa DOT's form). opts:
   * {orgName, fundingLabel} from org settings.
   */
  function formHtml(model, opts) {
    opts = opts || {};
    var c = model.costs || {};
    var m = model.modes || {};
    var totals = computeTotals(model);
    var funding = model.funding || model.tewd || "";
    var meals = "B: " + esc(c.mealsB || "0") + " &nbsp; L: " + esc(c.mealsL || "0") + " &nbsp; D: " + esc(c.mealsD || "0");
    return [
      '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:640px;color:#242424">',
      '<h2 style="color:#0f6cbd;margin-bottom:2px">' +
        (opts.orgName ? esc(opts.orgName) + " &mdash; " : "") + "Travel Authorization Request</h2>",
      '<style>table.fields{border-collapse:collapse;width:100%;margin:6px 0}table.fields td{border-bottom:1px solid #e1e1e1;padding:4px 6px;vertical-align:top}td.lbl{width:42%;color:#616161}td.val{font-weight:600}h3{margin:14px 0 2px;color:#0f6cbd}p.items{margin:4px 0}</style>',
      '<table class="fields">',
      fieldRow("Name & Cost Center", (model.name || "") + (model.costCenter ? " — " + model.costCenter : "")),
      fieldRow("Other Staff Attending", model.otherStaff),
      fieldRow("Name of Conference", model.event),
      fieldRow("Location", model.location),
      fieldRow("Conference Dates & Times", model.confDates),
      fieldRow("Departure Date", model.departDate),
      fieldRow("Return Date", model.returnDate),
      "</table>",
      "<p><strong>Reason for Travel:</strong> " + esc(model.reason || "") + "</p>",
      '<p class="items"><strong>Mode of Travel:</strong> ' +
        box(m.personal) + " Personal Vehicle &nbsp; " + box(m.state) + " State Vehicle &nbsp; " + box(m.air) + " Commercial Air</p>",
      '<table class="fields">',
      fieldRow("Cost of Travel Mode", c.travelMode ? money(num(c.travelMode)) : ""),
      fieldRow("Luggage Fees", c.luggage ? money(num(c.luggage)) : ""),
      fieldRow("Parking", c.parking ? money(num(c.parking)) : ""),
      fieldRow("Lodging", (c.lodgingNights || "0") + " nights @ " + money(num(c.lodgingRate)) + " = " + money(totals.lodging)),
      fieldRow("Registration Fee", (c.registration ? money(num(c.registration)) : "") + "  (meals included: " + meals + ")"),
      fieldRow("Taxi/Uber Fees", c.taxi ? money(num(c.taxi)) : ""),
      fieldRow("Additional Fees", (c.additional ? money(num(c.additional)) : "") + (c.additionalDesc ? " — " + c.additionalDesc : "")),
      fieldRow("ESTIMATED TOTAL", money(totals.grand)),
      "</table>",
      thirdPartyHtml((model.thirdParties || [])[0], "First"),
      thirdPartyHtml((model.thirdParties || [])[1], "Second"),
      funding ? "<p><strong>" + esc(opts.fundingLabel || "Funding") + ":</strong> " + esc(funding) + "</p>" : "",
      model.meetingLink ? "<p><strong>Meeting / event link:</strong> " + esc(model.meetingLink) + "</p>" : "",
      '<p style="color:#616161;font-size:12px">Generated by Travel Desk. The matching row was added to the shared travel planner.</p>',
      "</div>",
    ].join("\n");
  }

  /**
   * Choose the planner for a trip's fiscal year. planners is a map like
   * {"SFY27": {...}, "SFY26": {...}, "*": {...}} — exact FY match first,
   * then the "*" catch-all. Returns {key, planner} or null.
   */
  function pickPlanner(planners, fyLabel) {
    var p = planners || {};
    if (fyLabel && p[fyLabel]) { return { key: fyLabel, planner: p[fyLabel] }; }
    if (p["*"]) { return { key: "*", planner: p["*"] }; }
    return null;
  }

  /**
   * Match booking-confirmation emails (already filtered to booking senders,
   * e.g. Concur) to a trip. A candidate must arrive after the request and
   * before the trip ends; it's a CONFIDENT match when the subject/preview
   * mentions the destination city or the event name.
   * Returns {confident: email|null, candidates: [...]}.
   */
  // Words that mark a message as being ABOUT a booking. Deliberately broad:
  // a conference registration is as much a booking as a flight, and the
  // add-in should not decide which kinds of travel count.
  // How far back of a request to still accept a confirmation. Long enough
  // for "registered in June, filed the paperwork in August".
  var LOOKBACK_DAYS = 120;

  var BOOKING_WORDS = new RegExp("\\b(" + [
    "confirm(ed|ation|s)?", "itinerar(y|ies)", "reservation", "registration",
    "registered", "book(ed|ing)", "e-?ticket", "ticket", "check-?in",
    "trip", "flight", "hotel", "lodging", "airfare", "conference", "agenda",
  ].join("|") + ")\\b", "i");

  function matchBooking(trip, emails, trustedDomains) {
    var reqT = Date.parse(trip.createdAt || "") || 0;
    var endT = Date.parse(trip.returnDate || trip.eventStart || "") || (reqT + 90 * 864e5);
    endT += 864e5; // through the end of the return day
    // Confirmations that PREDATE the request are normal, not anomalous. For a
    // conference you register first and seek authorisation afterwards, so the
    // registration email is already sitting in the inbox when the request is
    // filed. Requiring booking-after-request made exactly that case invisible.
    var startT = reqT - LOOKBACK_DAYS * 864e5;
    var city = String(trip.location || "").split(",")[0].trim().toLowerCase();
    var eventTok = String(trip.event || "").toLowerCase();
    var candidates = (emails || []).filter(function (e) {
      var t = Date.parse(e.receivedDateTime || "") || 0;
      return t >= startT && t <= endT;
    });
    var trusted = (trustedDomains || []).map(function (d) {
      return String(d).toLowerCase().trim();
    }).filter(Boolean);

    // A match needs to be about THIS trip - the city or the event by name -
    // and to look like a booking at all. Requiring both keeps a wide net from
    // dragging in every message that happens to mention the city.
    var confident = candidates.filter(function (e) {
      var hay = ((e.subject || "") + " " + (e.bodyPreview || "")).toLowerCase();
      var aboutTrip = (city && city.length >= 3 && hay.indexOf(city) !== -1) ||
                      (eventTok && eventTok.length >= 6 && hay.indexOf(eventTok) !== -1);
      if (!aboutTrip) { return false; }
      var from = String(e.from || "").toLowerCase();
      var fromTrusted = trusted.some(function (d) { return d && from.indexOf(d) !== -1; });
      // A trusted sender is taken at its word; anyone else has to look like a
      // booking. That keeps Concur working exactly as before while letting a
      // conference registration through on its own merits.
      return fromTrusted || BOOKING_WORDS.test((e.subject || "") + " " + (e.bodyPreview || ""));
    });
    // Looking back past the request means an old confirmation can compete with
    // the real one. A confirmation issued AFTER the paperwork is the more
    // likely operative booking, so those win outright; the look-back only
    // decides when nothing arrived afterwards - which is precisely the
    // register-first case it exists for.
    var after = confident.filter(function (e) {
      return (Date.parse(e.receivedDateTime || "") || 0) >= reqT;
    });
    var pool = after.length ? after : confident;
    return {
      confident: pool.length === 1 ? pool[0] : null,
      candidates: candidates,
    };
  }

  /* ------------------------------------------------------------------
   * Third-party reimbursement follow-through. The money itself moves in
   * Workday; the add-in covers everything AROUND that: which receivables
   * exist, how old they are, the handoff packet finance needs to key the
   * Workday receivable, and the polite reminder to the third party.
   * ------------------------------------------------------------------ */

  /**
   * One receivable per third party on every ENDED trip. Status lives on
   * trip.reimb[tpIndex] = {status: "open"|"invoiced"|"received",
   * wdRef, on}. Unresolved sorted oldest-return first; received last.
   */
  /** Date-only strings parse as UTC midnight, which makes a trip look
   *  "ended" hours early (and skews ageDays) west of Greenwich. Treat a
   *  bare YYYY-MM-DD as the END of that day, locally. */
  function endOfLocalDay(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || "").trim());
    if (m) { return new Date(+m[1], +m[2] - 1, +m[3], 23, 59, 59).getTime(); }
    return Date.parse(iso) || 0;
  }

  function receivables(tripsList, now) {
    var nowT = (now instanceof Date ? now : new Date(now || Date.now())).getTime();
    var out = [];
    (tripsList || []).forEach(function (trip) {
      var tps = trip.thirdParties || [];
      if (!tps.length) { return; }
      var endIso = trip.returnDate || trip.eventStart || "";
      var endT = endOfLocalDay(endIso);
      if (!endT || endT > nowT) { return; } // trip not over yet
      tps.forEach(function (tp, i) {
        if (!tp || !(tp.name || "").trim()) { return; }
        var st = (trip.reimb && trip.reimb[i]) || {};
        out.push({
          tripId: trip.id, tpIndex: i,
          entity: tp.name.trim(), contact: tp.contact || "",
          project: tp.project || "",
          // a genuine 0 is an answer; only a blank is unknown
          amount: (tp.maxReimb === "" || tp.maxReimb == null) ? null : num(tp.maxReimb),
          traveler: trip.name || "", event: trip.event || "",
          location: trip.location || "", returnDate: endIso,
          ageDays: Math.max(0, Math.floor((nowT - endT) / 864e5)),
          status: st.status || "open", wdRef: st.wdRef || "", on: st.on || "",
        });
      });
    });
    out.sort(function (a, b) {
      var ra = a.status === "received" ? 1 : 0;
      var rb = b.status === "received" ? 1 : 0;
      if (ra !== rb) { return ra - rb; }
      return b.ageDays - a.ageDays;
    });
    return out;
  }

  function moneyRounded(n) { // distinct from money(): whole dollars, TBD-aware
    return n == null ? "(amount TBD)" :
      "$" + String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  }

  /**
   * Handoff email for finance/AR: everything needed to key the Workday
   * receivable, plus blank lines for what only Workday knows. The entry
   * itself stays human — this kills the retyping and the forgetting.
   */
  function workdayPacketHtml(r, opts) {
    opts = opts || {};
    var row = function (l, v) {
      return "<tr><td style=\"color:#616161;padding:3px 8px 3px 0;vertical-align:top\">" + esc(l) +
        "</td><td style=\"font-weight:600;padding:3px 0\">" + esc(v || "\u2014") + "</td></tr>";
    };
    return '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:640px;color:#242424">' +
      "<p>Please set up a third-party travel reimbursement receivable in Workday:</p>" +
      "<table style=\"border-collapse:collapse\">" +
      row("Third party", r.entity) +
      row("Billing contact", r.contact) +
      row("Project / agreement #", r.project) +
      row("Expected amount", moneyRounded(r.amount)) +
      row("Traveler", r.traveler) +
      row("Trip", r.event + (r.location ? " \u2014 " + r.location : "")) +
      row("Trip ended", (r.returnDate || "").slice(0, 10)) +
      (opts.funding ? row(opts.fundingLabel || "Funding", opts.funding) : "") +
      (opts.costCenter ? row("Cost center", opts.costCenter) : "") +
      "</table>" +
      "<p style=\"margin-top:14px\">For Workday (please reply with these once created):</p>" +
      "<p>Workday receivable #: ______________<br>" +
      "Cost center / fund to credit: ______________</p>" +
      "<p style=\"color:#616161;font-size:12px\">Sent from Travel Desk \u2014 details as captured on the travel authorization.</p>" +
      "</div>";
  }

  /** Polite follow-up to the third party itself. */
  function reminderHtml(r, opts) {
    opts = opts || {};
    return '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:640px;color:#242424">' +
      "<p>Hello" + (r.contact ? "" : " " + esc(r.entity)) + ",</p>" +
      "<p>Following up on the agreed travel reimbursement of <strong>" + moneyRounded(r.amount) +
      "</strong> for <strong>" + esc(r.traveler) + "</strong>'s attendance at <strong>" +
      esc(r.event) + "</strong>" + (r.location ? " (" + esc(r.location) + ")" : "") +
      ", which concluded on " + esc((r.returnDate || "").slice(0, 10)) + "." +
      (r.project ? " Reference: " + esc(r.project) + "." : "") + "</p>" +
      "<p>Could you let us know the status of the payment, or where to direct an invoice if one is needed?</p>" +
      "<p>Thank you!" + (opts.orgName ? "<br>" + esc(opts.orgName) : "") + "</p>" +
      "</div>";
  }

  /** Columns a generated planner starts with — chosen so every one of them
   *  is filled by the header matcher above. */
  var DEFAULT_PLANNER_HEADERS = ["Traveler", "Division", "Bureau", "Event", "Destination",
    "Start date", "Role", "Estimated cost", "Actual cost", "% Reimbursed", "3rd Party",
    "Funding", "Fiscal year", "Status", "Approved by", "Approved date", "Comments", "Approver"];

  /* ------------------------------------------------------------------
   * A trip is one row that moves through states — never two rows, or the
   * division totals count the same trip twice. The estimate is what gets
   * authorised; the actual is what it really cost. Keeping both on the row
   * makes the variance visible, and the variance is what defends next
   * year's budget request.
   *
   * The add-in RECORDS approvals; it cannot enforce them. There is no
   * backend to gate anything, and pretending otherwise would be dishonest
   * to anyone relying on it.
   * ------------------------------------------------------------------ */
  var STATUS = {
    REQUESTED: "Requested",           // traveller filed an estimate
    APPROVED: "Approved",             // coordinator authorised the trip
    BOOKED: "Booked",                 // travel arranged
    TRAVELED: "Traveled",             // trip happened, actuals not in yet
    ACTUALS: "Actuals submitted",     // traveller entered real costs
    CLOSED: "Closed",                 // coordinator accepted the actuals
    DECLINED: "Not approved",
  };

  /** Does this status mean the trip is authorised to happen? */
  function isAuthorized(status) {
    var s = String(status || "").toLowerCase();
    return ["approved", "booked", "traveled", "actuals submitted", "closed"]
      .indexOf(s) !== -1;
  }

  /** Estimate vs actual. Returns null when either side is unknown. */
  function variance(estimate, actual) {
    var e = num(estimate), a = num(actual);
    if (!e || !a) { return null; }
    return { delta: a - e, pct: Math.round(((a - e) / e) * 100) };
  }

  /**
   * How much attention a closed-out trip deserves. Small drift is noise;
   * a big overrun is the thing a coordinator actually needs to see.
   */
  function varianceFlag(estimate, actual, tolerancePct) {
    var v = variance(estimate, actual);
    if (!v) { return "unknown"; }
    var tol = tolerancePct == null ? 20 : tolerancePct;
    if (Math.abs(v.pct) <= tol) { return "ok"; }
    return v.pct > 0 ? "over" : "under";
  }

  /**
   * "Other travelers" box -> structured travelers. One per line:
   *   Jane Doe, MVD, Field Ops, speaker
   * Only the name is required; blank parts inherit the primary traveler's.
   */
  function parseTravelers(text) {
    return String(text || "").split(/[\n;]+/).map(function (line) {
      var parts = line.split(",").map(function (x) { return x.trim(); });
      if (!parts[0]) { return null; }
      return { name: parts[0], division: parts[1] || "", bureau: parts[2] || "", role: parts[3] || "" };
    }).filter(Boolean);
  }

  /**
   * The planner convention is one row per PERSON per event, so a delegation
   * has to become several rows or the overall view (and any budget total)
   * undercounts the trip. Returns one row per traveler.
   * opts.costMode: "per-person" (each row carries the entered cost, the
   * usual case) or "split" (the entered cost is the group's, divided).
   */
  function plannerRows(headers, model, opts) {
    opts = opts || {};
    var extra = parseTravelers(model.otherStaff);
    var people = [{ name: model.name, division: model.division, bureau: model.bureau, role: model.attendeeRole }]
      .concat(extra);
    var grand = computeTotals(model).grand;
    var override = null;
    if (opts.costMode === "split" && people.length > 1 && grand) {
      override = grand / people.length;
    }
    return people.map(function (p) {
      var m = {};
      Object.keys(model).forEach(function (k) { m[k] = model[k]; });
      m.name = p.name || model.name;
      m.division = p.division || model.division;
      m.bureau = p.bureau || model.bureau;
      m.attendeeRole = p.role || model.attendeeRole;
      var o = {};
      Object.keys(opts).forEach(function (k) { o[k] = opts[k]; });
      if (override != null) { o.costOverride = override; }
      return plannerRow(headers, m, o);
    });
  }

  /* ------------------------------------------------------------------
   * Setup by invitation. Copy-a-code-and-paste-it is the part coordinators
   * dread explaining, so the coordinator addresses people instead and each
   * traveller's own mailbox carries the setup to them. The code rides in a
   * marker the add-in can find again; nothing here is secret (planner link,
   * coordinator addresses, fiscal-year rules).
   * ------------------------------------------------------------------ */

  var SETUP_SUBJECT = "Travel Desk setup";
  var CODE_OPEN = "[[TD-SETUP]]";
  var CODE_CLOSE = "[[/TD-SETUP]]";

  function setupSubject(orgName) {
    var org = String(orgName || "").trim();
    return SETUP_SUBJECT + (org ? " \u2014 " + org : "");
  }

  /** The invitation a coordinator sends. Readable first, machine-readable second. */
  function setupInviteHtml(opts) {
    opts = opts || {};
    var org = esc(opts.orgName || "your organization");
    var who = esc(opts.coordName || "Your travel coordinator");
    return '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:640px;color:#242424">' +
      "<p>You've been set up to file travel requests for <strong>" + org + "</strong> " +
      "using the Travel Desk add-in in Outlook.</p>" +
      "<p><strong>To finish, open any email, click Travel Desk on the ribbon, and choose " +
      "&ldquo;Find my setup&rdquo;.</strong> It reads this message and configures itself \u2014 " +
      "the planner, the right addresses, and your fiscal-year rules. Nothing to copy or paste.</p>" +
      "<p style=\"color:#616161;font-size:12px\">Keep this message; the add-in looks for it. " +
      "The setup details below are just configuration \u2014 no passwords or personal data.</p>" +
      '<div style="background:#f3f2f1;border-radius:6px;padding:8px;font-family:ui-monospace,monospace;' +
      'font-size:11px;word-break:break-all;color:#616161">' +
      CODE_OPEN + esc(opts.code || "") + CODE_CLOSE + "</div>" +
      "<p style=\"color:#616161;font-size:12px\">Sent by " + who +
      (opts.coordEmail ? " (" + esc(opts.coordEmail) + ")" : "") + ".</p>" +
      "</div>";
  }

  /**
   * Pull the code back out of an invitation. Mail clients wrap and re-encode
   * bodies, so tolerate injected tags, entities and line breaks between the
   * markers.
   */
  function extractSetupCode(body) {
    var text = String(body || "");
    var i = text.indexOf(CODE_OPEN);
    var j = text.indexOf(CODE_CLOSE, i + 1);
    if (i === -1 || j === -1) { return ""; }
    return text.slice(i + CODE_OPEN.length, j)
      .replace(/<[^>]*>/g, "")
      .replace(/&amp;/g, "&")
      .replace(/[^A-Za-z0-9+/=]/g, "");
  }

  /** Newest invitation that actually carries a code. */
  function pickInvite(messages) {
    var withCode = (messages || []).filter(function (m) {
      return extractSetupCode(m.body || m.bodyPreview || "");
    });
    withCode.sort(function (a, b) {
      return String(b.receivedDateTime || "").localeCompare(String(a.receivedDateTime || ""));
    });
    return withCode[0] || null;
  }

  /**
   * A filed request, small enough to keep 40 of them.
   *
   * Everything on the form is worth carrying into the next request, but
   * roamingSettings caps this add-in at 32 KB across every setting it has, and
   * forty full models would exceed that on their own - at which point saveAsync
   * fails and the trip list stops recording anything. Empty boxes carry no
   * information, so they are dropped: a typical request loses two thirds of its
   * size and nothing a person would notice.
   */
  // The only inputs the pane ships with value="0".
  var ZERO_IS_BLANK = { mealsB: 1, mealsL: 1, mealsD: 1 };

  function slimModel(m) {
    if (!m || typeof m !== "object") { return {}; }
    var out = {};
    Object.keys(m).forEach(function (k) {
      var v = m[k];
      if (v === "" || v === null || v === undefined || v === false) { return; }
      if (Array.isArray(v)) {
        var arr = v.map(slimModel).filter(function (x) { return Object.keys(x).length; });
        if (arr.length) { out[k] = arr; }
        return;
      }
      if (typeof v === "object") {
        var sub = slimModel(v);
        if (Object.keys(sub).length) { out[k] = sub; }
        return;
      }
      // Only the meal counts start life at 0 on a blank form; every other
      // number box starts empty. So a zero elsewhere was typed, and a typed
      // zero is an answer worth carrying - "$0 luggage" is a checked box, not
      // an untouched one.
      if ((v === 0 || v === "0") && ZERO_IS_BLANK[k]) { return; }
      out[k] = v;
    });
    return out;
  }

  /**
   * What a copied request should carry, and what it must not.
   *
   * Dates already in the past cannot be right for a new trip, and a stale
   * departure date is the one mistake here with real consequences - it would
   * be signed, filed, and reconciled against a planner row that says the
   * traveller left last March. They are dropped and named, rather than
   * carried across and hopefully noticed.
   */
  function copyForNewTrip(saved, todayIso) {
    var m = JSON.parse(JSON.stringify(saved || {}));
    var today = String(todayIso || "").slice(0, 10);
    var dropped = [];
    ["eventStart", "departDate", "returnDate"].forEach(function (k) {
      if (m[k] && today && m[k] < today) { delete m[k]; dropped.push(k); }
    });
    if (dropped.length) { delete m.confDates; dropped.push("confDates"); }
    return { model: m, dropped: dropped };
  }

  /**
   * A short, stable fingerprint of a setup code.
   *
   * A traveler's settings are a COPY, taken once. Move the planner, change the
   * fiscal year or the coordinator address, and every traveler carries on
   * writing to the old workbook with the old rules - silently, because nothing
   * about their pane looks wrong. Storing what they applied is what makes
   * "yours is older than mine" answerable at all.
   *
   * FNV-1a: not a security hash, and does not need to be. It only has to
   * differ when the code differs.
   */
  function profileStamp(code) {
    var s = String(code || "");
    if (!s) { return ""; }
    var h = 0x811c9dc5;
    for (var i = 0; i < s.length; i++) {
      h ^= s.charCodeAt(i);
      h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
    }
    return ("0000000" + h.toString(16)).slice(-8);
  }

  /**
   * Is there a newer invitation than the one this traveler applied?
   *
   * Compared on the date first because that is free - the search returns it as
   * metadata. Only a message that is genuinely newer is worth opening to see
   * whether its code actually differs; a coordinator who resends an identical
   * profile should not make anybody do anything.
   */
  function newerInvite(messages, appliedAtIso) {
    var best = null;
    (messages || []).forEach(function (m) {
      if (!m || !m.receivedDateTime) { return; }
      if (appliedAtIso && m.receivedDateTime <= appliedAtIso) { return; }
      if (!best || m.receivedDateTime > best.receivedDateTime) { best = m; }
    });
    return best;
  }

  /**
   * A trip as a calendar entry, ready for Outlook's own appointment form.
   *
   * Deliberately NOT written to the calendar. Writing would need
   * Calendars.ReadWrite - a new scope, a changed consent screen, and another
   * pass over the listing and the privacy policy - to save one click. It also
   * breaks the rule the rest of this add-in follows: the draft opens, and you
   * press send. Same here: the appointment opens, and you press save.
   *
   * Hotel dates win over the event's when both exist, because a booking is
   * the trip's real shape - usually a day wider at each end than the
   * conference it is for.
   */
  function calendarEntry(trip, hotel) {
    var t = trip || {}, h = hotel || {};
    var start = h.checkIn || t.departDate || t.eventStart || "";
    var end = h.checkOut || t.returnDate || t.eventStart || start;
    if (!start) { return null; }
    if (end < start) { end = start; }

    var where = h.name
      ? h.name + (h.address && h.address !== t.location ? ", " + h.address : "")
      : (t.location || "");

    var lines = [];
    if (t.event) { lines.push(t.event); }
    if (t.location) { lines.push(t.location); }
    if (h.name) { lines.push("Staying at " + h.name + (h.confirmation ? " (" + h.confirmation + ")" : "")); }
    if (t.confDates) { lines.push(t.confDates); }
    if (t.meetingLink) { lines.push(t.meetingLink); }
    lines.push("");
    lines.push("Added from Travel Desk.");

    return {
      subject: (t.event || "Travel") + (t.location ? " \u2014 " + t.location : ""),
      location: where,
      start: start,
      end: end,
      body: lines.join("\n"),
    };
  }

  /** ISO date plus n days. */
  function shiftIso(isoDate, n) {
    var p = String(isoDate || "").split("-");
    if (p.length !== 3) { return isoDate; }
    var d = new Date(+p[0], +p[1] - 1, +p[2]);
    if (isNaN(d.getTime())) { return isoDate; }
    d.setDate(d.getDate() + n);
    var mm = d.getMonth() + 1, dd = d.getDate();
    return d.getFullYear() + "-" + (mm < 10 ? "0" : "") + mm + "-" + (dd < 10 ? "0" : "") + dd;
  }

  var api = {
    pickPlanner: pickPlanner,
    STATUS: STATUS,
    isAuthorized: isAuthorized,
    variance: variance,
    varianceFlag: varianceFlag,
    profileStamp: profileStamp,
    newerInvite: newerInvite,
    setupSubject: setupSubject,
    setupInviteHtml: setupInviteHtml,
    extractSetupCode: extractSetupCode,
    pickInvite: pickInvite,
    DEFAULT_PLANNER_HEADERS: DEFAULT_PLANNER_HEADERS,
    parseTravelers: parseTravelers,
    plannerRows: plannerRows,
    receivables: receivables,
    workdayPacketHtml: workdayPacketHtml,
    reminderHtml: reminderHtml,
    matchBooking: matchBooking,
    BOOKING_LOOKBACK_DAYS: LOOKBACK_DAYS,
    calendarEntry: calendarEntry,
    shiftIso: shiftIso,
    slimModel: slimModel,
    copyForNewTrip: copyForNewTrip,
    computeTotals: computeTotals,
    subjectLine: subjectLine,
    fiscalLabel: fiscalLabel,
    plannerRow: plannerRow,
    thirdPartySummary: thirdPartySummary,
    reimbursePct: reimbursePct,
    formHtml: formHtml,
    _internals: { num: num, esc: esc },
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.TravelForm = api; }
})(typeof self !== "undefined" ? self : this);
