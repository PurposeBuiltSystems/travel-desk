/*
 * Travel Desk — task pane UI wiring.
 *
 * Setup once (workbook link + table + coordinator), then each trip: fill the
 * form, click "Create travel request" — the add-in creates the Travel
 * Authorization email draft AND appends the matching planner row.
 */
/* global Office, GraphData, TravelForm, TravelCoord, XlsxGen, JSZip, document, window */
(function () {
  "use strict";

  var SETTINGS_KEY = "traveldesk.settings";
  var wbRef = null; // {driveId, itemId, name} cached after connect

  function byId(id) { return document.getElementById(id); }

  /**
   * Outlook caches the pane HTML but the ?v= query string makes it fetch
   * JavaScript fresh, so a returning user can run today's JS against
   * yesterday's page. Binding through this helper means a missing element
   * costs one feature instead of throwing and leaving EVERY button dead.
   */
  function on(id, ev, fn) {
    var el = byId(id);
    if (el) { el.addEventListener(ev, fn); }
    return el;
  }
  function val(id) { var el = byId(id); return el ? el.value : ""; }

  function setStatus(kind, text) {
    var el = byId("status");
    if (!text) { el.hidden = true; return; }
    el.hidden = false;
    el.className = "status " + kind;
    el.textContent = text;
  }

  function settings() {
    try {
      return JSON.parse(Office.context.roamingSettings.get(SETTINGS_KEY) || "{}");
    } catch (e) { return {}; }
  }

  function saveSettings(patch) {
    var s = settings();
    Object.keys(patch).forEach(function (k) { s[k] = patch[k]; });
    Office.context.roamingSettings.set(SETTINGS_KEY, JSON.stringify(s));
    Office.context.roamingSettings.saveAsync(function () {});
    try { renderFirstRun(); refreshCoordVisibility(); } catch (e) { /* pre-DOM calls are fine */ }
    return s;
  }

  // Settings fields that mirror a Setup input 1:1 (id === settings key).
  var ORG_FIELDS = ["coordEmail", "arEmail", "orgName", "fyStartMonth", "fyPrefix", "fundingLabel", "bookingSenders"];
  // What an org profile carries (coordinator → travelers). wbUrl/table ride
  // along so one Apply fully configures a traveler.
  var PROFILE_FIELDS = ORG_FIELDS.concat(["wbUrl", "tableName", "fundingOptions", "planners"]);

  function applyOrgLabels() {
    var s = settings();
    var label = (s.fundingLabel || "").trim();
    byId("fundingHead").hidden = byId("fundingWrap").hidden = !label && !byId("funding").value;
    byId("fundingLabelText").textContent = label || "Funding program";
    var dl = byId("fundingOptions");
    dl.innerHTML = "";
    String(s.fundingOptionsCsv || s.fundingOptions || "").split(",").forEach(function (v) {
      v = v.trim();
      if (!v) { return; }
      var o = document.createElement("option");
      o.value = v;
      dl.appendChild(o);
    });
  }

  Office.onReady(function () {
    var s = settings();
    if (s.wbUrl) { byId("wbUrl").value = s.wbUrl; }
    if (s.tableName) {
      var opt = document.createElement("option");
      opt.value = opt.textContent = s.tableName;
      byId("tableName").appendChild(opt);
    }
    ORG_FIELDS.forEach(function (k) {
      var el = byId(k);
      if (el && s[k] != null && s[k] !== "") { el.value = s[k]; }
    });
    ["name", "costCenter", "division", "bureau"].forEach(function (k) {
      var el = byId(k);
      if (el && s[k]) { el.value = s[k]; }
    });
    if (!s.name) {
      var prof = Office.context.mailbox.userProfile;
      if (prof && prof.displayName) { byId("name").value = prof.displayName; }
    }
    if (!s.wbUrl) { byId("setup").setAttribute("open", "open"); }
    applyOrgLabels();

    on("connect", "click", connectWorkbook);
    on("makeTable", "click", makeTable);
    on("createPlanner", "click", createPlanner);
    on("coordLoad", "click", coordLoad);
    on("wbBrowse", "click", browseWorkbooks);
    on("wbSearch", "keydown", function (ev) {
      if (ev.key === "Enter") { ev.preventDefault(); browseWorkbooks(); }
    });
    on("wbPick", "change", pickWorkbook);
    on("checkBookings", "click", checkBookings);
    on("profileApplyFast", "click", function () {
      byId("profileBlob").value = byId("profileBlobFast").value;
      profileApply();
      byId("profileBlobFast").value = "";
    });
    // returning users: lead with their trips; first-timers see the form
    var tripsEl = byId("trips").querySelector("details");
    if (trips().length === 0 && tripsEl) { tripsEl.removeAttribute("open"); }
    try { renderReimb(); renderTrips(); renderFirstRun(); refreshCoordVisibility(); }
    catch (e) { /* stale cached page — wiring below still binds */ }
    on("savePlanner", "click", savePlanner);
    on("plannerList", "click", function (ev) {
      var k = ev.target && ev.target.getAttribute && ev.target.getAttribute("data-del");
      if (!k) { return; }
      var st = settings();
      var pl = st.planners || {};
      delete pl[k];
      saveSettings({ planners: pl });
      renderPlannerList();
      setStatus("info", (k === "*" ? "Catch-all planner" : k + " planner") + " removed.");
    });
    if (!byId("wbFy").value) {
      var today = new Date();
      var p2 = function (n) { return (n < 10 ? "0" : "") + n; };
      // MUST be zero-padded: "2026-8-2" is an Invalid Date, which silently
      // blanked the fiscal year and saved the planner as the catch-all.
      var iso = today.getFullYear() + "-" + p2(today.getMonth() + 1) + "-" + p2(today.getDate());
      byId("wbFy").value = TravelForm.fiscalLabel(iso, s.fyStartMonth, s.fyPrefix) || "";
    }
    renderPlannerList();
    on("justChips", "click", function (ev) {
      var t = ev.target && ev.target.getAttribute && ev.target.getAttribute("data-txt");
      if (!t) { return; }
      var box = byId("reason");
      box.value = (box.value.trim() ? box.value.trim() + " " : "") + t;
    });
    on("submit", "click", submit);
    on("profileCopy", "click", profileCopy);
    on("profileApply", "click", profileApply);
    ORG_FIELDS.forEach(function (k) {
      on(k, "change", function () {
        var patch = {};
        patch[k] = byId(k).value;
        saveSettings(patch);
        applyOrgLabels();
      });
    });
    ["cTravelMode", "cLuggage", "cParking", "cTaxi", "cLodgingNights",
     "cLodgingRate", "cRegistration", "cAdditional"].forEach(function (id) {
      on(id, "input", refreshTotal);
    });
    ["eventStart", "departDate"].forEach(function (id) {
      on(id, "change", refreshFyLine);
    });
    on("eventStart", "change", function () {
      // prefill the free-text conference dates from the event start
      if (!byId("confDates").value.trim() && byId("eventStart").value) {
        var d = new Date(byId("eventStart").value + "T00:00:00");
        byId("confDates").value = d.toLocaleDateString(undefined, { month: "long", day: "numeric", year: "numeric" });
      }
    });
    refreshFyLine();
  });

  // ---------- org profile (share setup with the team) ----------

  function profileCopy() {
    var s = settings();
    var out = {};
    PROFILE_FIELDS.forEach(function (k) { if (s[k]) { out[k] = s[k]; } });
    var blob = btoa(unescape(encodeURIComponent(JSON.stringify(out))));
    byId("profileBlob").value = blob;
    byId("profileBlob").select();
    try { document.execCommand("copy"); } catch (e) { /* user copies manually */ }
    setStatus("info", "Profile code is in the box (and copied). Send it to your travelers.");
  }

  function profileApply() {
    try {
      var raw = byId("profileBlob").value.trim();
      if (!raw) { setStatus("error", "Paste the profile code from your coordinator first."); return; }
      var p = JSON.parse(decodeURIComponent(escape(atob(raw))));
      var patch = {};
      PROFILE_FIELDS.forEach(function (k) { if (p[k] != null) { patch[k] = p[k]; } });
      saveSettings(patch);
      ORG_FIELDS.forEach(function (k) {
        var el = byId(k);
        if (el && patch[k] != null) { el.value = patch[k]; }
      });
      if (patch.wbUrl) { byId("wbUrl").value = patch.wbUrl; }
      if (patch.tableName) {
        var sel = byId("tableName");
        sel.innerHTML = "";
        var opt = document.createElement("option");
        opt.value = opt.textContent = patch.tableName;
        sel.appendChild(opt);
      }
      applyOrgLabels();
      setStatus("info", "Profile applied" + (patch.wbUrl ? " — click Connect workbook to finish." : "."));
    } catch (e) {
      setStatus("error", "That doesn't look like a valid profile code.");
    }
  }

  function renderPlannerList() {
    var el = byId("plannerList");
    var pl = settings().planners || {};
    var keys = Object.keys(pl);
    if (!keys.length) {
      el.innerHTML = "No year-specific planners saved yet \u2014 connect a workbook, set its fiscal year above, and click Save.";
      return;
    }
    el.innerHTML = keys.sort().map(function (k) {
      var name = (pl[k].wbRef && pl[k].wbRef.name) || pl[k].wbUrl || "workbook";
      return "<b>" + (k === "*" ? "All years" : k) + "</b>: " + name + " \u00b7 " + (pl[k].tableName || "?") +
        ' <button type="button" class="chip-del" data-del="' + k + '">remove</button>';
    }).join("<br>");
  }

  function savePlanner() {
    var st = settings();
    if (!wbRef && !(st.wbRef)) {
      setStatus("error", "Connect the workbook first, then save it for a year."); return;
    }
    var key = byId("wbFy").value.trim() || "*";
    var pl = st.planners || {};
    pl[key] = {
      wbUrl: byId("wbUrl").value.trim(),
      tableName: byId("tableName").value,
      wbRef: wbRef || st.wbRef,
    };
    saveSettings({ planners: pl });
    renderPlannerList();
    setStatus("info", (key === "*" ? "Saved as the catch-all planner." :
      "Saved as the " + key + " planner \u2014 trips dated in " + key + " will go there automatically."));
  }

  function refreshFyLine() {
    var el = byId("fyLine");
    if (!el) { return; }
    var st = settings();
    var date = byId("eventStart").value || byId("departDate").value;
    if (!date) { el.textContent = ""; return; }
    var fy = TravelForm.fiscalLabel(date, st.fyStartMonth, st.fyPrefix);
    var picked = TravelForm.pickPlanner(st.planners, fy);
    if (picked) {
      el.innerHTML = "Files to: <b>" + (picked.key === "*" ? "your planner" : picked.key + " planner") + "</b> (" +
        ((picked.planner.wbRef && picked.planner.wbRef.name) || "workbook") + ") \u2713";
      el.className = "hint fy-ok";
    } else if (st.planners && Object.keys(st.planners).length) {
      el.innerHTML = "\u26a0 No planner saved for <b>" + fy + "</b> \u2014 connect that year's workbook in Setup before submitting.";
      el.className = "hint fy-warn";
    } else if (st.wbRef) {
      el.innerHTML = "Files to: <b>" + (st.wbRef.name || "your planner") + "</b> \u2713";
      el.className = "hint fy-ok";
    } else { el.textContent = ""; }
  }

  function refreshTotal() {
    byId("grandTotal").textContent = "$" +
      TravelForm.computeTotals(model()).grand.toLocaleString("en-US",
        { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  var TRIPS_KEY = "traveldesk.trips";

  function trips() {
    try { return JSON.parse(Office.context.roamingSettings.get(TRIPS_KEY) || "[]"); }
    catch (e) { return []; }
  }

  function saveTrips(list) {
    try {
      Office.context.roamingSettings.set(TRIPS_KEY, JSON.stringify(list.slice(-40)));
      Office.context.roamingSettings.saveAsync(function () {});
    } catch (e) { /* best-effort */ }
  }

  function addTrip(m) {
    var list = trips();
    list.push({
      // unique for the life of the mailbox: list.length repeats forever once
      // the 40-trip cap kicks in, which collided for recurring events
      id: "t" + Date.now().toString(36) + "-" + (m.event || "").slice(0, 8).replace(/\W/g, ""),
      name: m.name, event: m.event, location: m.location,
      eventStart: m.eventStart, departDate: m.departDate, returnDate: m.returnDate,
      createdAt: new Date().toISOString(),
      status: "requested", bookingLink: "", bookingSubject: "",
      thirdParties: (m.thirdParties || []).map(function (t) {
        return { name: t.name, contact: t.contact || "", project: t.project || "", maxReimb: t.maxReimb };
      }),
      funding: m.funding || m.tewd || "", costCenter: m.costCenter || "",
    });
    saveTrips(list);
    renderTrips();
    renderReimb();
    renderFirstRun();
  }

  /**
   * First-run guidance. A brand-new traveler opens a 40-field form with no
   * idea where to start; this says it in one line and disappears for good
   * once they're configured and have filed something.
   */
  function renderFirstRun() {
    var el = byId("firstRun");
    if (!el) { return; }
    var s = settings();
    var configured = !!(s.wbUrl || s.coordEmail ||
      (s.planners && Object.keys(s.planners).length));
    if (configured && trips().length) { el.hidden = true; return; }

    el.hidden = false;
    el.innerHTML = "";
    var h = document.createElement("p");
    h.className = "firstrun-h";
    var body = document.createElement("p");
    body.className = "firstrun-b";

    if (!configured) {
      h.textContent = "\ud83d\udc4b First time here?";
      body.innerHTML = "If your travel coordinator sent you a <b>setup code</b>, paste it into " +
        "Setup below \u2014 that is the entire setup. No code? Open <b>Coordinator setup</b> " +
        "to connect the planner workbook yourself. Either way you only do this once.";
      var btn = document.createElement("button");
      btn.textContent = "Open setup";
      btn.addEventListener("click", function () {
        var setup = byId("setup");
        setup.setAttribute("open", "open");
        try { setup.scrollIntoView({ behavior: "smooth", block: "start" }); }
        catch (e) { setup.scrollIntoView(); }
      });
      el.appendChild(h); el.appendChild(body); el.appendChild(btn);
      return;
    }
    h.textContent = "\u2705 You're set up \u2014 here's the routine";
    body.innerHTML = "Fill in <b>Traveler</b>, <b>Event</b>, and <b>Travel &amp; costs</b>, then click " +
      "<b>Create travel request</b> at the bottom. Only the fields marked " +
      "<i>required</i> must be filled. The email opens as a draft for you to review and send.";
    el.appendChild(h); el.appendChild(body);
  }

  /** The coordinator section only makes sense once a planner is connected —
   *  travellers who pasted a profile code never see it. */
  function refreshCoordVisibility() {
    var el = byId("coord");
    if (!el) { return; }
    var st = settings();
    el.hidden = !(st.tableName && (st.wbUrl || (st.planners && Object.keys(st.planners).length)));
  }

  function money0(n) {
    return "$" + String(Math.round(n || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  }

  /**
   * Load the whole planner and show the coordinator's working set. This is
   * the one view that reads everybody's rows rather than the signed-in
   * user's own trips.
   */
  async function coordLoad() {
    var st = settings();
    var ref = wbRef || (st.wbUrl ? null : null);
    byId("coordLoad").disabled = true;
    try {
      var token = await GraphData.getToken();
      if (!ref) {
        setStatus("work", "Opening the planner\u2026");
        ref = await GraphData.resolveWorkbook(token, st.wbUrl);
      }
      var tableName = st.tableName || (await GraphData.listTables(token, ref))[0];
      if (!tableName) { throw new Error("no table on that workbook yet"); }

      setStatus("work", "Reading everyone's trips\u2026");
      var headers = await GraphData.tableHeaders(token, ref, tableName);
      var rows = await GraphData.tableRows(token, ref, tableName);
      var records = TravelCoord.mapRows(headers, rows);
      var sum = TravelCoord.summarize(records, { fy: val("coordFy").trim() });

      var recon = null;
      if (byId("coordReconcile").checked) {
        setStatus("work", "Checking your inbox for authorizations\u2026");
        try {
          var mails = await GraphData.authEmails(token, 365);
          recon = TravelCoord.reconcile(records, mails.map(function (m) { return m.subject; }));
        } catch (e) { /* reconciliation is a bonus; the planner view still stands */ }
      }
      renderCoord(sum, recon);
      setStatus("info", sum.count + " trip(s) in the planner \u00b7 " + money0(sum.totalCost) +
        " total \u00b7 " + sum.unbooked.length + " upcoming still unbooked.");
    } catch (e) {
      setStatus("error", "Couldn't load the planner: " + ((e && e.message) || e));
    } finally {
      byId("coordLoad").disabled = false;
    }
  }

  function renderCoord(sum, recon) {
    var host = byId("coordView");
    host.innerHTML = "";
    function section(title) {
      var h = document.createElement("p");
      h.style.cssText = "font-weight:700;margin:12px 0 4px";
      h.textContent = title;
      host.appendChild(h);
      return h;
    }
    function line(text, tone) {
      var p = document.createElement("p");
      p.style.cssText = "margin:2px 0;font-size:12.5px" +
        (tone === "warn" ? ";color:var(--err-fg);font-weight:600" : "");
      p.textContent = text;
      host.appendChild(p);
    }
    function trip(r) {
      return (r.date || "(no date)") + " \u00b7 " + r.traveler +
        (r.division ? " (" + r.division + ")" : "") + " \u00b7 " + r.event +
        (r.cost ? " \u00b7 " + money0(r.cost) : "");
    }

    section("Still to book (" + sum.unbooked.length + ")");
    if (!sum.unbooked.length) { line("Everything upcoming is booked."); }
    sum.unbooked.slice(0, 15).forEach(function (r) { line(trip(r), "warn"); });

    section("Coming up (" + sum.upcoming.length + ")");
    if (!sum.upcoming.length) { line("Nothing upcoming in this view."); }
    sum.upcoming.slice(0, 15).forEach(function (r) { line(trip(r)); });

    section("Spend by division");
    sum.byDivision.forEach(function (d) {
      line(d.division + " \u00b7 " + d.trips + " trip(s) \u00b7 " + money0(d.cost));
    });
    if (sum.byFy.length > 1) {
      section("By fiscal year");
      sum.byFy.forEach(function (f) { line(f.fy + " \u00b7 " + f.trips + " trip(s) \u00b7 " + money0(f.cost)); });
    }

    section("Finished trips with a third party owing (" + sum.endedThirdParty.length + ")");
    if (!sum.endedThirdParty.length) { line("Nothing outstanding."); }
    sum.endedThirdParty.slice(0, 15).forEach(function (r) {
      line(trip(r) + " \u00b7 " + r.thirdParty, "warn");
    });

    if (recon) {
      section("Authorizations vs planner");
      if (!recon.rowsWithoutAuth.length && !recon.authsWithoutRow.length) {
        line("Every planner row has a matching authorization email. \u2713");
      }
      recon.rowsWithoutAuth.slice(0, 15).forEach(function (r) {
        line("No authorization email found for: " + trip(r), "warn");
      });
      recon.authsWithoutRow.slice(0, 15).forEach(function (p) {
        line("Authorized but never added to the planner: " + p.last + " \u00b7 " + p.event +
          (p.date ? " \u00b7 " + p.date : ""), "warn");
      });
    }
  }

  /** Resolve which planner a trip belongs to (same rule as filing it). */
  async function plannerFor(token, trip) {
    var st = settings();
    var fy = TravelForm.fiscalLabel(trip.eventStart || trip.departDate || trip.date,
      st.fyStartMonth, st.fyPrefix);
    var picked = TravelForm.pickPlanner(st.planners, fy);
    if (picked) {
      return { ref: picked.planner.wbRef, tableName: picked.planner.tableName };
    }
    var ref = wbRef || st.wbRef;
    if (!ref && st.wbUrl) { ref = await GraphData.resolveWorkbook(token, st.wbUrl); }
    return { ref: ref, tableName: st.tableName };
  }

  /**
   * Update this traveller's own row in the shared planner. Without this the
   * spreadsheet is append-only: a trip stays "Requested" forever and the
   * coordinator chases people who already booked. Best-effort by design —
   * the local state change stands even if the write-back can't happen.
   */
  async function writeBackRow(trip, mutate, label) {
    try {
      var token = await GraphData.getToken();
      var pl = await plannerFor(token, trip);
      if (!pl.ref || !pl.tableName) { return false; }
      var headers = await GraphData.tableHeaders(token, pl.ref, pl.tableName);
      var raw = await GraphData.tableRows(token, pl.ref, pl.tableName);
      var records = TravelCoord.mapRows(headers, raw);
      var i = TravelCoord.findRowIndex(records, {
        traveler: trip.name || trip.traveler, event: trip.event,
        date: trip.eventStart || trip.departDate || trip.date,
      });
      if (i < 0) { return false; }
      // mapRows drops blank rows, so translate back to the raw row position
      var rawIndex = records[i].rowNumber - 1;
      var values = (raw[rawIndex] || []).slice();
      var idx = TravelCoord.fieldIndex(headers);
      if (!mutate(values, idx)) { return false; }
      await GraphData.updateTableRow(token, pl.ref, pl.tableName, rawIndex, values);
      if (label) { setStatus("info", label); }
      return true;
    } catch (e) {
      setStatus("info", "Saved here, but the planner row couldn't be updated (" +
        ((e && e.message) || e) + "). You can change it in the spreadsheet.");
      return false;
    }
  }

  function markBookedEverywhere(trip) {
    return writeBackRow(trip, function (values, idx) {
      if (idx.status == null) { return false; }
      values[idx.status] = "Booked";
      return true;
    }, "Marked booked here and in the shared planner.");
  }

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function updateReimb(tripId, tpIndex, patch) {
    var list = trips();
    list.some(function (t) {
      if (t.id !== tripId) { return false; }
      t.reimb = t.reimb || {};
      var st = t.reimb[tpIndex] || {};
      Object.keys(patch).forEach(function (k) { st[k] = patch[k]; });
      t.reimb[tpIndex] = st;
      return true; // ids are unique — never touch a second trip
    });
    saveTrips(list);
    renderReimb();
  }

  async function draftReimb(r, kind, btn) {
    btn.disabled = true;
    try {
      var s = settings();
      var token = await GraphData.getToken();
      var to, subject, html;
      if (kind === "packet") {
        to = (s.arEmail || s.coordEmail || "").trim();
        subject = "Workday receivable — " + r.entity + " — " + r.event;
        var trip = trips().filter(function (t) { return t.id === r.tripId; })[0] || {};
        html = TravelForm.workdayPacketHtml(r, {
          funding: trip.funding || "", costCenter: trip.costCenter || "",
          fundingLabel: s.fundingLabel || "Funding",
        });
      } else {
        to = r.contact;
        subject = "Reimbursement follow-up — " + r.event + " — " + r.traveler;
        html = TravelForm.reminderHtml(r, { orgName: s.orgName || "" });
      }
      var draft = await GraphData.createDraft(token, to, subject, html);
      setStatus("info", (kind === "packet" ? "Workday packet" : "Reminder") +
        " draft created" + (to ? "" : " — add the recipient before sending") + ".");
      try {
        if (draft && draft.webLink) {
          Office.context.ui.openBrowserWindow
            ? Office.context.ui.openBrowserWindow(draft.webLink)
            : window.open(draft.webLink, "_blank");
        }
      } catch (e) { /* Drafts folder note above covers it */ }
      if (kind === "packet" && r.status === "open") {
        updateReimb(r.tripId, r.tpIndex, { status: "invoiced", on: new Date().toISOString().slice(0, 10) });
      }
    } catch (e) {
      setStatus("error", "Draft failed: " + ((e && e.message) || e));
    } finally {
      btn.disabled = false;
    }
  }

  function renderReimb() {
    var el = byId("reimbList");
    if (!el) { return; }
    var list = TravelForm.receivables(trips(), new Date());
    if (!list.length) { el.innerHTML = "No third-party reimbursements to chase."; return; }
    el.innerHTML = "";
    list.forEach(function (r) {
      var box = document.createElement("div");
      box.style.cssText = "border:1px solid var(--line);border-radius:8px;padding:8px 10px;margin-bottom:8px" +
        (r.status === "received" ? ";opacity:0.55" : "");
      var head = document.createElement("div");
      var amt = r.amount == null ? "amount TBD" : "$" + String(Math.round(r.amount)).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
      var stateText = r.status === "received" ? "\u2713 received " + r.on :
        r.status === "invoiced" ? "invoiced" + (r.wdRef ? " (WD " + r.wdRef + ")" : "") + " \u00b7 " + r.ageDays + "d" :
        r.ageDays + "d";
      head.innerHTML = "<strong>" + esc(r.entity) + "</strong> owes " + esc(amt) +
        " \u2014 " + esc(r.traveler) + ", " + esc(r.event) +
        " <span style=\"color:var(--warn-fg);font-weight:700\">" + esc(stateText) + "</span>";
      box.appendChild(head);
      var row = document.createElement("div");
      row.style.cssText = "display:flex;gap:6px;flex-wrap:wrap;margin-top:6px";
      function act(label, fn, primary) {
        var b = document.createElement("button");
        b.textContent = label;
        if (primary) { b.className = "primary"; }
        b.addEventListener("click", function () { fn(b); });
        row.appendChild(b);
        return b;
      }
      if (r.status !== "received") {
        act("Workday packet", function (b) { draftReimb(r, "packet", b); }, true);
        act("Remind " + r.entity.split(" ")[0], function (b) { draftReimb(r, "remind", b); });
        act("WD #…", function () {
          var wd = window.prompt("Workday receivable number for " + r.entity + ":", r.wdRef || "");
          if (wd !== null) { updateReimb(r.tripId, r.tpIndex, { wdRef: wd.trim(), status: "invoiced" }); }
        });
        act("Received ✓", function () {
          var on = new Date().toISOString().slice(0, 10);
          updateReimb(r.tripId, r.tpIndex, { status: "received", on: on });
          // settle it in the shared planner too, or the coordinator keeps
          // chasing money that already arrived
          var trip = trips().filter(function (t) { return t.id === r.tripId; })[0];
          if (trip) {
            writeBackRow(trip, function (values, idx) {
              if (idx.thirdParty == null) { return false; }
              values[idx.thirdParty] = TravelCoord.markSettled(values[idx.thirdParty], on);
              return true;
            }, "Marked received here and settled in the shared planner.");
          }
        });
      } else {
        act("Reopen", function () { updateReimb(r.tripId, r.tpIndex, { status: "invoiced" }); });
      }
      box.appendChild(row);
      el.appendChild(box);
    });
  }

  function renderTrips() {
    var el = byId("tripsList");
    if (!el) { return; }
    var list = trips();
    if (!list.length) { el.innerHTML = "No requests tracked yet \u2014 submit one and it appears here."; return; }
    el.innerHTML = list.slice().reverse().map(function (t, ri) {
      var i = list.length - 1 - ri;
      // "departing soon" = still ahead of us AND inside 14 days; the old
      // test was also true for every trip already in the past, so finished
      // trips stayed flagged red forever.
      var untilDepart = t.departDate ? Date.parse(t.departDate + "T23:59:59") - Date.now() : null;
      var chip = t.status === "booked"
        ? '<span style="color:var(--ok-fg);font-weight:700">\u2713 Booked</span>'
        : (untilDepart !== null && untilDepart > 0 && untilDepart < 14 * 864e5
            ? '<span style="color:var(--err-fg);font-weight:700">\u23f3 Requested \u2014 no booking email yet</span>'
            : '<span style="color:var(--warn-fg);font-weight:700">\u23f3 Requested</span>');
      return "<div style=\"border-top:1px solid var(--line);padding:6px 0\"><b>" + (t.event || "trip") + "</b> \u00b7 " +
        (t.eventStart || t.departDate || "") + " \u00b7 " + chip +
        (t.bookingLink ? ' \u00b7 <a href="' + t.bookingLink + '" target="_blank" rel="noopener">open booking email</a>' : "") +
        (t.status !== "booked" ? ' \u00b7 <button type="button" class="chip-del" data-book="' + i + '">mark booked</button>' : "") +
        "</div>";
    }).join("");
    el.querySelectorAll("[data-book]").forEach(function (b) {
      b.addEventListener("click", function () {
        var list2 = trips();
        var bi = Number(b.getAttribute("data-book"));
        list2[bi].status = "booked";
        markBookedEverywhere(list2[bi]);
        saveTrips(list2);
        renderTrips();
      });
    });
  }

  async function checkBookings() {
    byId("checkBookings").disabled = true;
    try {
      var list = trips();
      var open = list.filter(function (t) { return t.status !== "booked"; });
      if (!open.length) { setStatus("info", "No trips waiting on bookings. \u2708\ufe0f"); return; }
      setStatus("work", "Checking your inbox for booking confirmations\u2026");
      var token = await GraphData.getToken();
      var since = open.reduce(function (min, t) { return t.createdAt < min ? t.createdAt : min; }, open[0].createdAt);
      var senders = (byId("bookingSenders").value || "concursolutions.com").split(",");
      var emails = await GraphData.bookingEmails(token, since, senders);
      var booked = 0, unsure = 0;
      open.forEach(function (t) {
        var m = TravelForm.matchBooking(t, emails);
        if (m.confident) {
          t.status = "booked";
          markBookedEverywhere(t);
          t.bookingLink = m.confident.webLink || "";
          t.bookingSubject = m.confident.subject || "";
          booked++;
        } else if (m.candidates.length) { unsure++; }
      });
      saveTrips(list);
      renderTrips();
      setStatus(booked || !unsure ? "info" : "error",
        booked + " trip(s) matched to booking emails" +
        (unsure ? " \u00b7 " + unsure + " trip(s) have possible matches \u2014 check your inbox and use \u201cmark booked\u201d" : "") +
        (!booked && !unsure ? " \u2014 none found yet. Booking emails usually take a few days after the request." : "") + ".");
    } catch (e) {
      setStatus("error", "Booking check failed: " + ((e && e.message) || e));
    } finally {
      byId("checkBookings").disabled = false;
    }
  }

  var pickerRefs = [];

  async function browseWorkbooks() {
    byId("wbBrowse").disabled = true;
    try {
      var q = byId("wbSearch").value.trim();
      setStatus("work", q ? 'Searching your files for \u201c' + q + '\u201d\u2026' : "Loading your recent workbooks\u2026");
      var token = await GraphData.getToken();
      var found;
      if (q) {
        found = await GraphData.searchWorkbooks(token, q);
      } else {
        var recent = await GraphData.recentWorkbooks(token);
        var shared = [];
        try { shared = await GraphData.sharedWorkbooks(token); } catch (e) { /* optional */ }
        var seen = {};
        found = [];
        recent.concat(shared).forEach(function (r) {
          var k = r.driveId + "|" + r.itemId;
          if (seen[k]) { return; }
          seen[k] = true;
          found.push(r);
        });
      }
      pickerRefs = found;
      var sel = byId("wbPick");
      sel.innerHTML = "";
      var o0 = document.createElement("option");
      o0.value = ""; o0.textContent = found.length ? "Pick a workbook (" + found.length + ")\u2026" : "No workbooks found \u2014 search or paste a link";
      sel.appendChild(o0);
      found.forEach(function (r, i) {
        var o = document.createElement("option");
        o.value = String(i);
        o.textContent = r.name;
        sel.appendChild(o);
      });
      setStatus(found.length ? "info" : "error",
        found.length ? found.length + " workbook(s) \u2014 pick your planner." :
        "Nothing matched. Try a search term, or paste the workbook's Copy-link below.");
    } catch (e) {
      setStatus("error", "Couldn't list your files: " + ((e && e.message) || e));
    } finally {
      byId("wbBrowse").disabled = false;
    }
  }

  async function pickWorkbook() {
    var idx = byId("wbPick").value;
    if (idx === "") { return; }
    var ref = pickerRefs[Number(idx)];
    if (!ref) { return; }
    try {
      setStatus("work", 'Opening \u201c' + ref.name + '\u201d\u2026');
      var token = await GraphData.getToken();
      var tables = await GraphData.listTables(token, ref);
      if (!tables.length) {
        setStatus("error", '\u201c' + ref.name + '\u201d has no Excel Table. In Excel: select the planner\u2019s header row and data, Insert > Table, then pick it again.');
        return;
      }
      wbRef = ref;
      byId("wbUrl").value = ref.webUrl || "";
      var sel = byId("tableName");
      sel.innerHTML = "";
      tables.forEach(function (t) {
        var opt = document.createElement("option");
        opt.value = opt.textContent = t;
        sel.appendChild(opt);
      });
      saveSettings({ wbUrl: ref.webUrl || "", tableName: tables[0], wbRef: ref });
      setStatus("info", "Connected: " + ref.name + " \u2014 table \u201c" + tables[0] + "\u201d. Set its fiscal year above and click Save planner.");
    } catch (e) {
      setStatus("error", "Couldn't open that workbook: " + ((e && e.message) || e));
    }
  }

  /** Populate the sheet picker so a table-less workbook can be converted. */
  async function offerTableConversion(token, ref) {
    var box = byId("noTableBox");
    var sel = byId("tableSheet");
    if (!box || !sel) { return; }
    try {
      var sheets = await GraphData.listWorksheets(token, ref);
      sel.innerHTML = "";
      sheets.forEach(function (n) {
        var o = document.createElement("option");
        o.value = o.textContent = n;
        sel.appendChild(o);
      });
      box.hidden = !sheets.length;
    } catch (e) { box.hidden = true; }
  }

  /**
   * Convert the chosen sheet's used range into a real Excel Table. The data
   * is untouched — a table is a wrapper, so the overall view keeps every row
   * and gains filters, an auto-expanding range, and something Power BI and
   * this add-in can both bind to.
   */
  async function makeTable() {
    var sheet = val("tableSheet");
    if (!wbRef) { setStatus("error", "Connect the workbook first."); return; }
    if (!sheet) { setStatus("error", "Pick which sheet the planner is on."); return; }
    byId("makeTable").disabled = true;
    try {
      setStatus("work", "Reading \u201c" + sheet + "\u201d\u2026");
      var token = await GraphData.getToken();
      var used = await GraphData.usedRange(token, wbRef, sheet);
      var address = used && used.address;
      if (!address) { throw new Error("that sheet looks empty — add your header row first"); }
      setStatus("work", "Creating the table\u2026");
      var t = await GraphData.addTable(token, wbRef, address, "TravelPlanner");
      var name = (t && t.name) || "TravelPlanner";
      var sel = byId("tableName");
      sel.innerHTML = "";
      var o = document.createElement("option");
      o.value = o.textContent = name;
      sel.appendChild(o);
      byId("noTableBox").hidden = true;
      saveSettings({ wbUrl: byId("wbUrl").value.trim(), tableName: name });
      setStatus("info", "\u201c" + sheet + "\u201d is now the table \u201c" + name +
        "\u201d (" + address + ") \u2014 every existing row is still there. Set the fiscal year and save the planner.");
    } catch (e) {
      setStatus("error", "Couldn't create the table: " + ((e && e.message) || e));
    } finally {
      byId("makeTable").disabled = false;
    }
  }

  /** Generate a correctly-shaped planner in OneDrive and connect to it. */
  async function createPlanner() {
    var name = (val("newPlannerName") || "Division travel planner").trim();
    if (!/\.xlsx$/i.test(name)) { name += ".xlsx"; }
    byId("createPlanner").disabled = true;
    try {
      setStatus("work", "Building the workbook\u2026");
      var built = XlsxGen.buildWorkbook(TravelForm.DEFAULT_PLANNER_HEADERS, "Planner");
      var zip = new JSZip();
      Object.keys(built.parts).forEach(function (path) { zip.file(path, built.parts[path]); });
      var bytes = await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });

      setStatus("work", "Saving it to your OneDrive\u2026");
      var token = await GraphData.getToken();
      var made = await GraphData.uploadWorkbook(token, name, bytes);
      wbRef = { driveId: made.ref.driveId, itemId: made.ref.itemId, name: made.name };

      setStatus("work", "Formatting it as a table\u2026");
      var t = await GraphData.addTable(token, wbRef, built.range, "TravelPlanner");
      var tableName = (t && t.name) || "TravelPlanner";

      var sel = byId("tableName");
      sel.innerHTML = "";
      var o = document.createElement("option");
      o.value = o.textContent = tableName;
      sel.appendChild(o);
      byId("wbUrl").value = made.webUrl || "";
      byId("noTableBox").hidden = true;
      saveSettings({ wbUrl: made.webUrl || "", tableName: tableName });
      setStatus("info", "Created \u201c" + made.name + "\u201d in your OneDrive and connected it. " +
        "Set the fiscal year it covers, then Save planner for this year.");
    } catch (e) {
      setStatus("error", "Couldn't create the planner: " + ((e && e.message) || e));
    } finally {
      byId("createPlanner").disabled = false;
    }
  }

  async function connectWorkbook() {
    var url = byId("wbUrl").value.trim();
    if (!url) { setStatus("error", "Paste the planner workbook's link first (Copy link in SharePoint/OneDrive)."); return; }
    byId("connect").disabled = true;
    try {
      setStatus("work", "Connecting to the planner workbook…");
      var token = await GraphData.getToken();
      wbRef = await GraphData.resolveWorkbook(token, url);
      var tables = await GraphData.listTables(token, wbRef);
      if (!tables.length) {
        await offerTableConversion(token, wbRef);
        setStatus("error", "\"" + wbRef.name + "\" has no Excel table yet — pick the sheet below and " +
          "click \u201cMake this sheet a table\u201d.");
        return;
      }
      var box = byId("noTableBox");
      if (box) { box.hidden = true; }
      var sel = byId("tableName");
      sel.innerHTML = "";
      tables.forEach(function (t) {
        var opt = document.createElement("option");
        opt.value = opt.textContent = t;
        sel.appendChild(opt);
      });
      saveSettings({ wbUrl: url, tableName: tables[0], wbRef: wbRef });
      setStatus("info", "Connected: " + wbRef.name + " — table \"" + tables[0] + "\". You're set.");
      sel.addEventListener("change", function () { saveSettings({ tableName: sel.value }); });
    } catch (e) {
      setStatus("error", "Couldn't open the workbook: " + ((e && e.message) || e));
    } finally {
      byId("connect").disabled = false;
    }
  }

  function tp(n) {
    var name = byId("tp" + n + "Name").value.trim();
    return {
      name: name,
      contact: val("tp" + n + "Contact").trim(),
      project: byId("tp" + n + "Project").value.trim(),
      packet: byId("tp" + n + "Packet").checked,
      maxReimb: byId("tp" + n + "Max").value,
      items: {
        registration: byId("tp" + n + "Reg").checked,
        lodging: byId("tp" + n + "Lodging").checked,
        airfare: byId("tp" + n + "Air").checked,
        meals: byId("tp" + n + "Meals").checked,
        ground: byId("tp" + n + "Ground").checked,
      },
      notes: byId("tp" + n + "Notes").value.trim(),
    };
  }

  function model() {
    return {
      name: byId("name").value.trim(),
      costCenter: byId("costCenter").value.trim(),
      division: byId("division").value.trim(),
      bureau: byId("bureau").value.trim(),
      otherStaff: byId("otherStaff").value.trim(),
      event: byId("event").value.trim(),
      location: byId("location").value.trim(),
      eventStart: byId("eventStart").value,
      attendeeRole: byId("attendeeRole").value.trim(),
      confDates: byId("confDates").value.trim(),
      departDate: byId("departDate").value,
      returnDate: byId("returnDate").value,
      reason: byId("reason").value.trim(),
      meetingLink: byId("meetingLink").value.trim(),
      comments: byId("comments").value.trim(),
      funding: byId("funding").value.trim(),
      modes: {
        personal: byId("modePersonal").checked,
        state: byId("modeState").checked,
        air: byId("modeAir").checked,
      },
      costs: {
        travelMode: byId("cTravelMode").value,
        luggage: byId("cLuggage").value,
        parking: byId("cParking").value,
        taxi: byId("cTaxi").value,
        lodgingNights: byId("cLodgingNights").value,
        lodgingRate: byId("cLodgingRate").value,
        registration: byId("cRegistration").value,
        additional: byId("cAdditional").value,
        additionalDesc: byId("cAdditionalDesc").value.trim(),
        mealsB: byId("cMealsB").value,
        mealsL: byId("cMealsL").value,
        mealsD: byId("cMealsD").value,
      },
      thirdParties: [tp(1), tp(2)].filter(function (t) { return t.name; }),
    };
  }

  async function submit() {
    var m = model();
    var doDraft = byId("doDraft").checked;
    var doRow = byId("doRow").checked;
    if (!doDraft && !doRow) { setStatus("error", "Pick at least one action."); return; }
    if (!m.name || !m.event || !m.location) {
      setStatus("error", "Name, event, and location are required.");
      return;
    }
    var s = settings();
    if (doRow && !(s.wbRef || wbRef)) {
      setStatus("error", "Connect the planner workbook first (Setup section), or untick the planner action.");
      return;
    }

    byId("submit").disabled = true;
    try {
      var token = await GraphData.getToken();
      var done = [];
      var orgOpts = {
        orgName: s.orgName || "",
        fundingLabel: s.fundingLabel || "",
        fyStartMonth: Number(s.fyStartMonth) || 1,
        fyPrefix: s.fyPrefix || "FY",
        costMode: val("costMode") || "per-person",
      };

      if (doDraft) {
        setStatus("work", "Creating the Travel Authorization draft…");
        var draft = await GraphData.createDraft(token, byId("coordEmail").value.trim(),
          TravelForm.subjectLine(m), TravelForm.formHtml(m, orgOpts));
        done.push("draft OPENED for you — review it and press Send");
        // Open the draft immediately: "draft created" is not "request sent",
        // and this is the moment the traveler must press Send.
        try {
          if (draft && draft.webLink) {
            Office.context.ui.openBrowserWindow
              ? Office.context.ui.openBrowserWindow(draft.webLink)
              : window.open(draft.webLink, "_blank");
          }
        } catch (e) { done.push("(open your Drafts folder to send it)"); }
      }

      if (doRow) {
        setStatus("work", "Adding the planner row…");
        var tripFy = TravelForm.fiscalLabel(m.eventStart || m.departDate, s.fyStartMonth, s.fyPrefix);
        var picked = TravelForm.pickPlanner(s.planners, tripFy);
        var ref, tableName;
        if (picked) {
          ref = picked.planner.wbRef;
          tableName = picked.planner.tableName;
        } else if (s.planners && Object.keys(s.planners).length) {
          throw new Error("No planner saved for " + (tripFy || "that date") +
            " — in Setup, connect that year's workbook and save it for " + tripFy +
            " (or save one planner as 'all years').");
        } else {
          ref = wbRef || s.wbRef; // legacy single-planner setup
          tableName = s.tableName;
        }
        var headers = await GraphData.tableHeaders(token, ref, tableName);
        // One row per traveler: the planner's convention is one line per
        // person per event, and anything less undercounts a delegation.
        var rows = TravelForm.plannerRows(headers, m, orgOpts);
        for (var ri = 0; ri < rows.length; ri++) {
          if (rows.length > 1) { setStatus("work", "Adding planner row " + (ri + 1) + " of " + rows.length + "…"); }
          await GraphData.addTableRow(token, ref, tableName, rows[ri]);
        }
        done.push(rows.length + (rows.length === 1 ? " row" : " rows (one per traveler)") +
          " added to " + (picked && picked.key !== "*" ? picked.key + " planner (" : "") +
          (ref.name || "the planner") + (picked && picked.key !== "*" ? ")" : ""));
      }

      saveSettings({ name: m.name, costCenter: m.costCenter, division: m.division, bureau: m.bureau });
      addTrip(m);
      setStatus("info", "Done: " + done.join(" + ") + ".");
    } catch (e) {
      setStatus("error", "Travel request failed: " + ((e && e.message) || e) +
        (byId("doDraft").checked && byId("doRow").checked
          ? " — if the draft was created, it is still in Drafts; fix and retry with only the failed action ticked."
          : ""));
    } finally {
      byId("submit").disabled = false;
    }
  }
})();
