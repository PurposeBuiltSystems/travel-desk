/*
 * Travel Desk — task pane UI wiring.
 *
 * Setup once (workbook link + table + coordinator), then each trip: fill the
 * form, click "Create travel request" — the add-in creates the Travel
 * Authorization email draft AND appends the matching planner row.
 */
/* global Office, GraphData, TravelForm, document */
(function () {
  "use strict";

  var SETTINGS_KEY = "traveldesk.settings";
  var wbRef = null; // {driveId, itemId, name} cached after connect

  function byId(id) { return document.getElementById(id); }

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
    return s;
  }

  // Settings fields that mirror a Setup input 1:1 (id === settings key).
  var ORG_FIELDS = ["coordEmail", "orgName", "fyStartMonth", "fyPrefix", "fundingLabel"];
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
    ORG_FIELDS.forEach(function (k) { if (s[k] != null && s[k] !== "") { byId(k).value = s[k]; } });
    ["name", "costCenter", "division", "bureau"].forEach(function (k) {
      if (s[k]) { byId(k).value = s[k]; }
    });
    if (!s.name) {
      var prof = Office.context.mailbox.userProfile;
      if (prof && prof.displayName) { byId("name").value = prof.displayName; }
    }
    if (!s.wbUrl) { byId("setup").setAttribute("open", "open"); }
    applyOrgLabels();

    byId("connect").addEventListener("click", connectWorkbook);
    byId("wbBrowse").addEventListener("click", browseWorkbooks);
    byId("wbSearch").addEventListener("keydown", function (ev) {
      if (ev.key === "Enter") { ev.preventDefault(); browseWorkbooks(); }
    });
    byId("wbPick").addEventListener("change", pickWorkbook);
    byId("savePlanner").addEventListener("click", savePlanner);
    byId("plannerList").addEventListener("click", function (ev) {
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
      var iso = today.getFullYear() + "-" + (today.getMonth() + 1) + "-" + today.getDate();
      byId("wbFy").value = TravelForm.fiscalLabel(iso, s.fyStartMonth, s.fyPrefix) || "";
    }
    renderPlannerList();
    byId("justChips").addEventListener("click", function (ev) {
      var t = ev.target && ev.target.getAttribute && ev.target.getAttribute("data-txt");
      if (!t) { return; }
      var box = byId("reason");
      box.value = (box.value.trim() ? box.value.trim() + " " : "") + t;
    });
    byId("submit").addEventListener("click", submit);
    byId("profileCopy").addEventListener("click", profileCopy);
    byId("profileApply").addEventListener("click", profileApply);
    ORG_FIELDS.forEach(function (k) {
      byId(k).addEventListener("change", function () {
        var patch = {};
        patch[k] = byId(k).value;
        saveSettings(patch);
        applyOrgLabels();
      });
    });
    ["cTravelMode", "cLuggage", "cParking", "cTaxi", "cLodgingNights",
     "cLodgingRate", "cRegistration", "cAdditional"].forEach(function (id) {
      byId(id).addEventListener("input", refreshTotal);
    });
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
      ORG_FIELDS.forEach(function (k) { if (patch[k] != null) { byId(k).value = patch[k]; } });
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

  function refreshTotal() {
    byId("grandTotal").textContent = "$" +
      TravelForm.computeTotals(model()).grand.toLocaleString("en-US",
        { minimumFractionDigits: 2, maximumFractionDigits: 2 });
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
        setStatus("error", "\"" + wbRef.name + "\" has no Excel Table. In Excel: select the planner's header row and data, Insert > Table, then reconnect. (Tables are what keep row-adds reliable.)");
        return;
      }
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
      };

      if (doDraft) {
        setStatus("work", "Creating the Travel Authorization draft…");
        await GraphData.createDraft(token, byId("coordEmail").value.trim(),
          TravelForm.subjectLine(m), TravelForm.formHtml(m, orgOpts));
        done.push("draft in your Drafts folder (review and send)");
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
        await GraphData.addTableRow(token, ref, tableName, TravelForm.plannerRow(headers, m, orgOpts));
        done.push("row added to " + (picked && picked.key !== "*" ? picked.key + " planner (" : "") +
          (ref.name || "the planner") + (picked && picked.key !== "*" ? ")" : ""));
      }

      saveSettings({ name: m.name, costCenter: m.costCenter, division: m.division, bureau: m.bureau });
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
