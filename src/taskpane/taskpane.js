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

  Office.onReady(function () {
    var s = settings();
    if (s.wbUrl) { byId("wbUrl").value = s.wbUrl; }
    if (s.coordEmail) { byId("coordEmail").value = s.coordEmail; }
    if (s.tableName) {
      var opt = document.createElement("option");
      opt.value = opt.textContent = s.tableName;
      byId("tableName").appendChild(opt);
    }
    ["name", "costCenter", "division", "bureau"].forEach(function (k) {
      if (s[k]) { byId(k).value = s[k]; }
    });
    if (!s.name) {
      var prof = Office.context.mailbox.userProfile;
      if (prof && prof.displayName) { byId("name").value = prof.displayName; }
    }
    if (!s.wbUrl) { byId("setup").setAttribute("open", "open"); }

    byId("connect").addEventListener("click", connectWorkbook);
    byId("submit").addEventListener("click", submit);
    ["cTravelMode", "cLuggage", "cParking", "cTaxi", "cLodgingNights",
     "cLodgingRate", "cRegistration", "cAdditional"].forEach(function (id) {
      byId(id).addEventListener("input", refreshTotal);
    });
  });

  function refreshTotal() {
    byId("grandTotal").textContent = "$" +
      TravelForm.computeTotals(model()).grand.toLocaleString("en-US",
        { minimumFractionDigits: 2, maximumFractionDigits: 2 });
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
      tewd: byId("tewd").value,
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

      if (doDraft) {
        setStatus("work", "Creating the Travel Authorization draft…");
        await GraphData.createDraft(token, byId("coordEmail").value.trim(),
          TravelForm.subjectLine(m), TravelForm.formHtml(m));
        done.push("draft in your Drafts folder (review and send)");
      }

      if (doRow) {
        setStatus("work", "Adding the planner row…");
        var ref = wbRef || s.wbRef;
        var headers = await GraphData.tableHeaders(token, ref, s.tableName);
        await GraphData.addTableRow(token, ref, s.tableName, TravelForm.plannerRow(headers, m));
        done.push("row added to " + (ref.name || "the planner"));
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
