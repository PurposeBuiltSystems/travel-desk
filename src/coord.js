/*
 * Travel Desk — coordinator logic (pure).
 *
 * Travellers work from their own trips, which live in their own mailbox
 * settings. The coordinator works from the PLANNER, which is the shared
 * truth: everybody's trips, every division. This module reads planner rows
 * back into records, summarises them, and reconciles them against the
 * Travel Authorization emails sitting in the coordinator's inbox.
 *
 * Header matching mirrors form.js's writer (same keywords, same order, so a
 * column written by the add-in is read back into the same field).
 */
(function (root) {
  "use strict";

  function norm(s) { return String(s == null ? "" : s).trim(); }
  function low(s) { return norm(s).toLowerCase(); }

  function num(v) {
    var n = parseFloat(String(v == null ? "" : v).replace(/[$,\s]/g, ""));
    return isNaN(n) ? 0 : n;
  }

  /** Which column holds which field. Same tests as the writer, same order. */
  function fieldIndex(headers) {
    var idx = {};
    (headers || []).forEach(function (h, i) {
      var k = low(h);
      if (!k) { return; }
      function set(f) { if (idx[f] == null) { idx[f] = i; } }
      if (k.indexOf("event") !== -1 || k.indexOf("conference") !== -1) { return set("event"); }
      if (k.indexOf("city") !== -1 || k.indexOf("location") !== -1 || k.indexOf("destination") !== -1) { return set("destination"); }
      if (k.indexOf("start") !== -1 || k.indexOf("date") !== -1) { return set("date"); }
      if (k.indexOf("division") !== -1 || k.indexOf("department") !== -1) { return set("division"); }
      if (k.indexOf("bureau") !== -1 || k.indexOf("office") !== -1 || k.indexOf("team") !== -1) { return set("bureau"); }
      if ((k.indexOf("attendee") !== -1 || k.indexOf("traveler") !== -1 || k.indexOf("name") !== -1) &&
          k.indexOf("role") === -1) { return set("traveler"); }
      if (k.indexOf("role") !== -1) { return set("role"); }
      if (k.indexOf("cost") !== -1 || k.indexOf("estimate") !== -1) { return set("cost"); }
      if (k.indexOf("%") !== -1) { return set("pctReimb"); }
      if (k.indexOf("third") !== -1 || k.indexOf("3rd") !== -1 ||
          k.indexOf("grant") !== -1 || k.indexOf("reimburs") !== -1) { return set("thirdParty"); }
      if (k.indexOf("fiscal") !== -1) { return set("fy"); }
      if (k.indexOf("tewd") !== -1 || k.indexOf("funding") !== -1 || k.indexOf("program") !== -1) { return set("funding"); }
      if (k.indexOf("comment") !== -1 || k.indexOf("note") !== -1) { return set("comments"); }
      if (k.indexOf("status") !== -1 || k.indexOf("booked") !== -1) { return set("status"); }
    });
    return idx;
  }

  /** Planner rows (arrays of cell values) -> records. Blank rows dropped. */
  function mapRows(headers, rows) {
    var idx = fieldIndex(headers);
    var out = [];
    (rows || []).forEach(function (r, n) {
      var get = function (f) { return idx[f] == null ? "" : norm(r[idx[f]]); };
      var rec = {
        rowNumber: n + 1,
        traveler: get("traveler"), division: get("division"), bureau: get("bureau"),
        event: get("event"), destination: get("destination"), date: get("date"),
        role: get("role"), cost: num(get("cost")), pctReimb: num(get("pctReimb")),
        thirdParty: get("thirdParty"), funding: get("funding"), fy: get("fy"),
        status: get("status"), comments: get("comments"),
      };
      if (!rec.traveler && !rec.event) { return; } // spacer / sample row
      out.push(rec);
    });
    return out;
  }

  /** "Yes - AASHTO" means a third party owes something; "No"/"" doesn't. */
  function hasThirdParty(rec) {
    var t = low(rec.thirdParty);
    return !!t && t !== "no" && t !== "none" && t !== "n/a";
  }

  function isBooked(rec) {
    var s = low(rec.status);
    return s.indexOf("book") !== -1 || s.indexOf("approv") !== -1 ||
           s.indexOf("complete") !== -1 || s.indexOf("done") !== -1;
  }

  function dayMs(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(norm(iso));
    if (m) { return new Date(+m[1], +m[2] - 1, +m[3], 12, 0, 0).getTime(); }
    var t = Date.parse(iso);
    return isNaN(t) ? null : t;
  }

  /**
   * The coordinator's working set: what's coming, what still isn't booked,
   * where the money is going, and which finished trips owe reimbursement.
   * opts: {now, fy} — fy filters to one fiscal year when supplied.
   */
  function summarize(records, opts) {
    opts = opts || {};
    var nowT = (opts.now instanceof Date ? opts.now : new Date(opts.now || Date.now())).getTime();
    var recs = (records || []).filter(function (r) {
      return !opts.fy || low(r.fy) === low(opts.fy);
    });

    var upcoming = [], unbooked = [], endedThirdParty = [];
    var byDivision = {}, byFy = {}, total = 0;

    recs.forEach(function (r) {
      var t = dayMs(r.date);
      var future = t != null && t >= nowT;
      if (future) { upcoming.push(r); }
      if (future && !isBooked(r)) { unbooked.push(r); }
      if (t != null && t < nowT && hasThirdParty(r)) { endedThirdParty.push(r); }

      var d = r.division || "(none)";
      var slot = byDivision[d] || (byDivision[d] = { division: d, trips: 0, cost: 0 });
      slot.trips++; slot.cost += r.cost;
      total += r.cost;

      var f = r.fy || "(no FY)";
      var fslot = byFy[f] || (byFy[f] = { fy: f, trips: 0, cost: 0 });
      fslot.trips++; fslot.cost += r.cost;
    });

    var byDate = function (a, b) { return (dayMs(a.date) || 0) - (dayMs(b.date) || 0); };
    upcoming.sort(byDate);
    unbooked.sort(byDate);
    endedThirdParty.sort(byDate);

    return {
      count: recs.length, totalCost: total,
      upcoming: upcoming, unbooked: unbooked, endedThirdParty: endedThirdParty,
      byDivision: Object.keys(byDivision).sort(function (a, b) {
        return byDivision[b].cost - byDivision[a].cost;
      }).map(function (k) { return byDivision[k]; }),
      byFy: Object.keys(byFy).sort().reverse().map(function (k) { return byFy[k]; }),
    };
  }

  /* ------------------------------------------------------- reconciliation */

  function slug(s) { return low(s).replace(/[^a-z0-9]+/g, ""); }

  /** Last name the same way the authorization subject builds it. */
  function lastNameOf(name) {
    var n = norm(name);
    var parts = n.split(/[\s,]+/).filter(Boolean);
    return /,/.test(n) ? (parts[0] || "") : (parts[parts.length - 1] || "");
  }

  /** "Travel Auth - Miller - TRB Annual - 2027-01-11" -> parts, or null. */
  function parseAuthSubject(subject) {
    var s = norm(subject).replace(/^(re|fw|fwd)\s*:\s*/i, "");
    if (!/^travel auth\s*-/i.test(s)) { return null; }
    var parts = s.split("-").map(function (x) { return x.trim(); });
    parts.shift(); // "Travel Auth"
    var date = "";
    if (parts.length >= 4 && /^\d{4}$/.test(parts[parts.length - 3])) {
      date = parts.slice(-3).join("-");        // the event date got split on its hyphens
      parts = parts.slice(0, -3);
    }
    var last = parts.shift() || "";
    return { last: last, event: parts.join(" - "), date: date };
  }

  function keyOf(last, event) { return slug(last) + "|" + slug(event); }

  /**
   * Cross-check the planner against the authorization emails the
   * coordinator actually received. Catches both directions: a request that
   * never reached the planner, and a planner row nobody authorised.
   */
  function reconcile(records, subjects) {
    var authKeys = {};
    var parsedAuths = [];
    (subjects || []).forEach(function (subj) {
      var p = parseAuthSubject(subj);
      if (!p || !p.last) { return; }
      var k = keyOf(p.last, p.event);
      if (!authKeys[k]) { authKeys[k] = true; parsedAuths.push(p); }
    });

    var rowKeys = {};
    (records || []).forEach(function (r) {
      rowKeys[keyOf(lastNameOf(r.traveler), r.event)] = true;
    });

    return {
      rowsWithoutAuth: (records || []).filter(function (r) {
        return !authKeys[keyOf(lastNameOf(r.traveler), r.event)];
      }),
      authsWithoutRow: parsedAuths.filter(function (p) {
        return !rowKeys[keyOf(p.last, p.event)];
      }),
    };
  }

  var api = {
    fieldIndex: fieldIndex, mapRows: mapRows, summarize: summarize,
    reconcile: reconcile, parseAuthSubject: parseAuthSubject,
    lastNameOf: lastNameOf, hasThirdParty: hasThirdParty, isBooked: isBooked,
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.TravelCoord = api; }
})(typeof self !== "undefined" ? self : this);
