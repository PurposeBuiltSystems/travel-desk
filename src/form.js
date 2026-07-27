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
      if (k.indexOf("event") !== -1 || k.indexOf("conference") !== -1) { return model.event || ""; }
      if (k.indexOf("city") !== -1 || k.indexOf("location") !== -1 || k.indexOf("destination") !== -1) { return model.location || ""; }
      if (k.indexOf("start") !== -1 || k.indexOf("date") !== -1) { return model.eventStart || model.departDate || ""; }
      if (k.indexOf("division") !== -1 || k.indexOf("department") !== -1) { return model.division || ""; }
      if (k.indexOf("bureau") !== -1 || k.indexOf("office") !== -1 || k.indexOf("team") !== -1) { return model.bureau || ""; }
      if ((k.indexOf("attendee") !== -1 || k.indexOf("traveler") !== -1 || k.indexOf("name") !== -1) && k.indexOf("role") === -1) { return model.name || ""; }
      if (k.indexOf("role") !== -1) { return model.attendeeRole || ""; }
      if (k.indexOf("cost") !== -1 || k.indexOf("estimate") !== -1) { return totals.grand ? String(Math.round(totals.grand)) : ""; }
      if (k.indexOf("third") !== -1 || k.indexOf("grant") !== -1) { return thirdPartySummary(model); }
      if (k.indexOf("%") !== -1) { return reimbursePct(model); }
      if (k.indexOf("fiscal") !== -1) { return fy; }
      if (k.indexOf("tewd") !== -1 || k.indexOf("funding") !== -1 || k.indexOf("program") !== -1) { return funding; }
      if (k.indexOf("comment") !== -1 || k.indexOf("note") !== -1) { return model.comments || ""; }
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

  var api = {
    pickPlanner: pickPlanner,
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
