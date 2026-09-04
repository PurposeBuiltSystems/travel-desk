/*
 * Travel Desk — task pane UI wiring.
 *
 * Setup once (workbook link + table + coordinator), then each trip: fill the
 * form, click "Create travel request" — the add-in creates the Travel
 * Authorization email draft AND appends the matching planner row.
 */
/* global Office, GraphData, TravelForm, TravelCoord, XlsxGen, TdMail, TdAttach, JSZip, document, window */
(function () {
  "use strict";

  var SETTINGS_KEY = "traveldesk.settings";
  var wbRef = null; // {driveId, itemId, name} cached after connect

  function byId(id) { return document.getElementById(id); }

  /**
   * Guarded element access. Outlook desktop caches the pane HTML far harder
   * than the web client while ?v= still fetches today's JavaScript, so startup
   * routinely runs new code against an old page. One unguarded
   * `byId(x).value` there throws inside Office.onReady, and Outlook reports
   * that as "Add-in Error" - the whole pane, not one field. This is the exact
   * cause of certification finding 1120.3.7.8 on a sibling add-in.
   */
  function setVal(id, v) { var el = byId(id); if (el) { el.value = v; } }
  function setProp(id, k, v) { var el = byId(id); if (el) { el[k] = v; } }
  function setText(id, t) { var el = byId(id); if (el) { el.textContent = t; } }
  function setAttrIf(id, n, v) { var el = byId(id); if (el) { el.setAttribute(n, v); } }
  function rmAttrIf(id, n) { var el = byId(id); if (el) { el.removeAttribute(n); } }
  function isChecked(id) { var el = byId(id); return !!(el && el.checked); }

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

  /**
   * roamingSettings caps at 32 KB across everything this add-in stores. Over
   * that, saveAsync FAILS - and a callback that ignores asyncResult.status
   * turns a failure into silent data loss: the user believes it saved. Every
   * write of real user data goes through here so a failure is at least said
   * out loud.
   */
  function persistSettings(what) {
    Office.context.roamingSettings.saveAsync(function (r) {
      if (r && r.status !== Office.AsyncResultStatus.Succeeded) {
        setStatus("error", "Couldn't save " + (what || "your settings") +
          " \u2014 you may be at the 32 KB limit Outlook allows an add-in. " +
          "Recent changes may not survive a restart.");
      }
    });
  }


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
    persistSettings("your settings");
    try { renderFirstRun(); refreshCoordVisibility(); } catch (e) { /* pre-DOM calls are fine */ }
    return s;
  }

  // Settings fields that mirror a Setup input 1:1 (id === settings key).
  var ORG_FIELDS = ["coordEmail", "arEmail", "orgName", "fyStartMonth", "fyPrefix", "fundingLabel", "bookingSenders"];
  // What an org profile carries (coordinator → travelers). wbUrl/table ride
  // along so one Apply fully configures a traveler.
  var PROFILE_FIELDS = ORG_FIELDS.concat(["wbUrl", "tableName", "fundingOptions", "planners", "formLabels"]);

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

  /**
   * Account row. Certification policy 1100.5.7.1 requires a visible way out
   * wherever an add-in signs a user in. Every element access is guarded:
   * Outlook desktop caches the pane HTML while ?v= fetches fresh JS, so this
   * code can run against a page that predates these controls, and an
   * unguarded dereference here would throw inside Office.onReady and take
   * the whole pane down as "Add-in Error".
   */
  function authSet(id, k, v) { var e = document.getElementById(id); if (e) { e[k] = v; } }

  async function renderAuthState() {
    var who = null;
    try { who = await GraphData.currentAccount(); } catch (e) { who = null; }
    authSet("authWho", "textContent", who ? ("Signed in as " + who) : "Not signed in");
    authSet("signOut", "hidden", !who);
    authSet("signIn", "hidden", !!who);
  }

  async function doSignIn() {
    authSet("signIn", "disabled", true);
    try { await GraphData.getToken(); }
    catch (e) { authSet("authWho", "textContent", "Sign-in failed: " + ((e && e.message) || e)); }
    finally { authSet("signIn", "disabled", false); renderAuthState(); }
  }

  async function doSignOut() {
    // Respond immediately: awaiting a broker handshake before changing
    // anything on screen is what certification saw as the button doing nothing.
    authSet("signOut", "disabled", true);
    authSet("signOut", "hidden", true);
    authSet("signIn", "hidden", false);
    authSet("authWho", "textContent", "Signed out. This add-in will ask you to sign in " +
      "again before its next action. Your Outlook session is separate and is not affected.");
    try {
      await GraphData.signOut();
    } catch (e) {
      // The enforced state is already set; a failed cache clear doesn't undo it.
    } finally {
      authSet("signOut", "disabled", false);
    }
  }



  /**
   * Run one piece of startup without letting it take the pane with it.
   *
   * Everything below used to be a single unguarded block, so ONE bad line -
   * a render that trips over a stored trip, a field a cached page does not
   * have - threw inside Office.onReady, and Outlook answers that by showing
   * nothing at all. A blank pane tells the person nothing and tells me less.
   * Isolated, a failure costs one feature and says which one.
   */
  var startupFailures = [];
  function step(what, fn) {
    try { fn(); }
    catch (e) {
      startupFailures.push(what + " (" + ((e && e.message) || e) + ")");
      try { if (window.console && console.error) { console.error("Travel Desk startup:", what, e); } }
      catch (e2) { /* nothing left to try */ }
    }
  }

  Office.onReady(function () {
   try {
    // Certification 1100.5.7.1 - sign-out must be reachable.
    var _si = document.getElementById("signIn");
    if (_si) { _si.addEventListener("click", doSignIn); }
    var _so = document.getElementById("signOut");
    if (_so) { _so.addEventListener("click", doSignOut); }
    step("the sign-in state", renderAuthState);
    var s = settings();
    if (s.wbUrl) { setVal("wbUrl", s.wbUrl); }
    if (s.tableName) {
      var tableSel = byId("tableName");
      if (tableSel) {
        var opt = document.createElement("option");
        opt.value = opt.textContent = s.tableName;
        tableSel.appendChild(opt);
      }
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
      if (prof && prof.displayName) { setVal("name", prof.displayName); }
    }
    // Remembered per person, not shared in the org profile - where you keep
    // your own file is not a decision to push onto everyone else.
    if (s.plannerFolder) { setVal("newPlannerFolder", s.plannerFolder); }
    if (!s.wbUrl) { setAttrIf("setup", "open", "open"); }
    applyOrgLabels();

    on("connect", "click", connectWorkbook);
    on("makeTable", "click", makeTable);
    on("createPlanner", "click", createPlanner);
    on("coordLoad", "click", coordLoad);
    on("closeSubmit", "click", closeOutTrip);
    on("wbBrowse", "click", browseWorkbooks);
    on("wbSearch", "keydown", function (ev) {
      if (ev.key === "Enter") { ev.preventDefault(); browseWorkbooks(); }
    });
    on("wbPick", "change", pickWorkbook);
    on("checkBookings", "click", checkBookings);
    on("profileApplyFast", "click", function () {
      profileApply(val("profileBlobFast"));
      setVal("profileBlobFast", "");
    });
    // returning users: lead with their trips; first-timers see the form
    var tripsHost = byId("trips");
    var tripsEl = tripsHost && tripsHost.querySelector("details");
    if (trips().length === 0 && tripsEl) { tripsEl.removeAttribute("open"); }
    // These were already guarded, but silently: a swallowed failure is exactly
    // why a half-working pane gives you nothing to go on.
    step("your trips", renderTrips);
    step("the reimbursement list", renderReimb);
    step("the first-run panel", renderFirstRun);
    step("the coordinator section", refreshCoordVisibility);
    on("savePlanner", "click", savePlanner);
    on("addCols", "click", addLifecycleColumns);
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
    if (!val("wbFy")) {
      var today = new Date();
      var p2 = function (n) { return (n < 10 ? "0" : "") + n; };
      // MUST be zero-padded: "2026-8-2" is an Invalid Date, which silently
      // blanked the fiscal year and saved the planner as the catch-all.
      var iso = today.getFullYear() + "-" + p2(today.getMonth() + 1) + "-" + p2(today.getDate());
      setVal("wbFy", TravelForm.fiscalLabel(iso, s.fyStartMonth, s.fyPrefix) || "");
    }
    step("the planner list", renderPlannerList);
    step("the planner link", renderPlannerLink);
    step("the build indicator", renderVersion);
    // Only when it has never verified: this must not become a Graph call on
    // every single open.
    step("the setup freshness check", checkSetupFreshness);
    step("the stale-invitation notice", noteInvitesStale);
    step("the setup checklist", renderCoordSteps);
    if (s.wbRef && s.plannerWriteOk !== true) {
      step("the planner access check", function () { verifyPlannerAccess(s.wbRef, "traveler"); });
    }
    // Diagnostic only: never awaited, never blocks startup, and swallows its
    // own failure inside checkPlannerColumns().
    if (s.wbRef && s.tableName) { checkPlannerColumns(); }
    on("justChips", "click", function (ev) {
      var t = ev.target && ev.target.getAttribute && ev.target.getAttribute("data-txt");
      if (!t) { return; }
      var box = byId("reason");
      if (!box) { return; }
      box.value = (box.value.trim() ? box.value.trim() + " " : "") + t;
    });
    on("submit", "click", submit);
    on("fillFromEmail", "click", fillFromEmail);
    on("findTravel", "click", findTravelEmails);
    on("learnScan", "click", learnScan);
    on("learnSave", "click", learnSave);
    on("learnClear", "click", learnClear);
    on("profileCopy", "click", profileCopy);
    on("inviteSend", "click", sendInvites);
    /*
     * Who you are sticks the moment you type it, not when you submit.
     *
     * Name, cost center, division and bureau were only persisted inside
     * submit(). So typing your cost center and then doing anything else -
     * reopening the pane, reloading after an update, or simply not finishing
     * that request - silently discarded it, and the next visit showed the old
     * value with no sign anything had been lost. These are four facts about
     * you that never change; there is no reason to make completing a whole
     * travel request the price of recording them.
     */
    ["name", "costCenter", "division", "bureau"].forEach(function (k) {
      on(k, "change", function () {
        var patch = {};
        patch[k] = val(k).trim();
        saveSettings(patch);
      });
    });
    ORG_FIELDS.forEach(function (k) {
      on(k, "change", function () {
        var patch = {};
        patch[k] = byId(k).value;
        saveSettings(patch);
        applyOrgLabels();
        noteInvitesStale();
        renderCoordSteps();
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
      if (!val("confDates").trim() && val("eventStart")) {
        var d = new Date(val("eventStart") + "T00:00:00");
        setVal("confDates", d.toLocaleDateString(undefined, { month: "long", day: "numeric", year: "numeric" }));
      }
    });
    step("the fiscal-year line", refreshFyLine);
   } catch (fatal) {
    // Nothing above should reach here now, but if it does the pane must still
    // be usable and must say what happened rather than going white.
    startupFailures.push("startup (" + ((fatal && fatal.message) || fatal) + ")");
   }
   if (startupFailures.length) {
    setStatus("error", "Travel Desk started with " + startupFailures.length +
      " problem(s): " + startupFailures.join("; ") +
      ". The rest of the pane still works. If this persists, close and reopen it, " +
      "and check the build number at the top.");
   }
  });

  // ---------- org profile (share setup with the team) ----------

  /**
   * The coordinator's three stages, in the order they have to happen.
   *
   * The order was always real - you cannot invite people to a planner you
   * have not saved, and a profile without your address reaches nobody - but
   * it was invisible, spread across two levels of collapsed sections with no
   * indication of how far through you were or what was left. Every one of
   * these steps existed already; this only says which one is next.
   *
   * Pure, so the sequencing can be tested without a DOM.
   */
  function coordStages(st) {
    var s = st || {};
    var savedPlanner = !!(s.planners && Object.keys(s.planners).some(function (k) {
      return s.planners[k] && s.planners[k].wbRef;
    }));
    var org = !!String(s.coordEmail || "").trim();
    var invited = !!s.invitesSentAt;
    return [
      { key: "org", name: "Say who you are",
        why: "Your travel coordinator's address \u2014 where every request is sent.",
        done: org, action: "coordEmail" },
      { key: "planner", name: "Set up the planner",
        why: savedPlanner ? "" :
          (s.wbUrl ? "Connected, but not saved yet \u2014 set its fiscal year and save it."
                   : "Create one, or connect the workbook you already use."),
        done: savedPlanner, action: "planner" },
      { key: "invite", name: "Invite your travelers",
        why: "They click one button and are set up from it.",
        done: invited, action: "inviteTo", needs: org && savedPlanner },
    ];
  }

  /**
   * What is still missing before a setup is worth handing to anybody.
   *
   * Both of these let a coordinator finish successfully and ship an
   * incomplete setup to their whole team, which is the worst kind of
   * unfinished - it looks done.
   *
   * CONNECTING a workbook is not SAVING it. Connect stores the link, which is
   * enough for the invitation to be allowed, but not the resolved reference -
   * so every traveler who applies that code is told to "click Connect
   * workbook once to finish", a technical instruction they were never meant
   * to see and cannot be expected to understand.
   *
   * And nothing ever required the coordinator's own address, so a profile
   * could go out without one and every traveler's authorization draft would
   * be addressed to nobody.
   */
  function setupGaps() {
    var st = settings();
    var gaps = [];
    var armed = st.planners && Object.keys(st.planners).some(function (k) {
      return st.planners[k] && st.planners[k].wbRef;
    });
    if (!armed) {
      gaps.push(st.wbUrl
        ? "the planner is connected but not saved \u2014 set its fiscal year and click " +
          "\u201cSave planner for this year\u201d, or your travelers each have to connect it themselves"
        : "no planner is connected yet");
    }
    if (!String(st.coordEmail || "").trim()) {
      gaps.push("no travel coordinator address \u2014 without it, every traveler\u2019s " +
        "authorization draft is addressed to nobody");
    }
    return gaps;
  }

  /** The setup code as it stands right now. */
  function currentProfileCode() {
    var st = settings();
    var out = {};
    PROFILE_FIELDS.forEach(function (k) { if (st[k]) { out[k] = st[k]; } });
    try { return btoa(unescape(encodeURIComponent(JSON.stringify(out)))); }
    catch (e) { return ""; }
  }

  /**
   * Tell the coordinator when what they have sent no longer matches what they
   * have.
   *
   * The traveler's side of this notices too, but only the next time each of
   * them opens the pane - which for somebody who travels twice a year is
   * months. The person who made the change is the one who can fix it today,
   * and they are the only one who knows they made it.
   */
  function noteInvitesStale() {
    var st = settings();
    if (!st.invitesSentAt) { return; }
    var box = byId("inviteInfo");
    if (!box) { return; }
    var now = TravelForm.profileStamp(currentProfileCode());
    if (!now || now === st.invitesStamp) { return; }
    box.textContent = "\u26a0\ufe0f Your settings have changed since you last invited anyone. " +
      "Until you send the invitation again, everyone already set up keeps using the old " +
      "planner and the old rules \u2014 and their trips will not appear in your view.";
    box.style.color = "var(--err-fg)";
  }

  function profileCopy() {
    var gaps = setupGaps();
    if (gaps.length) {
      setStatus("error", "That code would not fully set anybody up: " + gaps.join("; ") + ".");
      return;
    }
    var s = settings();
    var out = {};
    PROFILE_FIELDS.forEach(function (k) { if (s[k]) { out[k] = s[k]; } });
    var blob = btoa(unescape(encodeURIComponent(JSON.stringify(out))));
    var box = byId("profileBlob");
    box.value = blob;
    box.select();
    var done = function () {
      setStatus("info", "Setup code generated and copied. Email it to your travelers \u2014 they paste it " +
        "into \u201cI'm a traveler\u2026\u201d at the top of Setup.");
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(blob).then(done, function () {
        try { document.execCommand("copy"); } catch (e) { /* select-and-copy by hand */ }
        done();
      });
    } else {
      try { document.execCommand("copy"); } catch (e) { /* select-and-copy by hand */ }
      done();
    }
  }

  function profileApply(rawText) {
    try {
      var raw = String(rawText != null ? rawText : val("profileBlob")).trim();
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
      // Record what was applied, so a later invitation can be recognised as
      // newer rather than identical.
      saveSettings({
        profileStamp: TravelForm.profileStamp(raw),
        profileAppliedAt: new Date().toISOString(),
      });
      renderCoordSteps();
      // a saved planner carries its resolved ref, so there's nothing left to
      // click; only a legacy single-workbook code needs Connect
      var ready = patch.planners && Object.keys(patch.planners).some(function (k) {
        return patch.planners[k] && patch.planners[k].wbRef;
      });
      setStatus("info", ready
        ? "You're set up — the planner and your coordinator's settings all came across. " +
          "Fill in the form below and click Create travel request."
        : (patch.wbUrl ? "Setup applied — click Connect workbook once to finish." : "Setup applied."));
      // Not awaited: setup is finished either way, and this only ever speaks
      // up to say the rows will not land.
      var firstRef = null;
      Object.keys(patch.planners || {}).some(function (k) {
        if (patch.planners[k] && patch.planners[k].wbRef) { firstRef = patch.planners[k].wbRef; return true; }
        return false;
      });
      if (firstRef) { verifyPlannerAccess(firstRef, "traveler"); }
    } catch (e) {
      setStatus("error", "That doesn't look like a valid profile code.");
    }
  }

  /**
   * Planners created before the trip lifecycle shipped have the original 15
   * columns. Close-out writes Actual cost, and approval writes Approved by /
   * Approved date, so against an older planner those silently have nowhere to
   * go - the close-out already fails with "the planner has no Actual cost
   * column yet - ask your coordinator to add one (Setup can do it)". This is
   * the part of Setup that can actually do it.
   *
   * Repair in place rather than asking anyone to rebuild: a planner in use has
   * a year of trips in it, and columns can be appended to an Excel table
   * without touching a single existing row.
   */
  var LIFECYCLE_COLUMNS = ["Actual cost", "Approved by", "Approved date"];
  var missingCols = [];

  async function checkPlannerColumns() {
    var st = settings();
    var ref = st.wbRef;
    var table = st.tableName;
    if (!ref || !table) { setText("colCheck", ""); setProp("addCols", "hidden", true); return; }
    try {
      var token = await GraphData.getToken();
      var have = await GraphData.tableColumns(token, ref, table);
      var lower = have.map(function (h) { return String(h).toLowerCase().trim(); });
      missingCols = LIFECYCLE_COLUMNS.filter(function (c) {
        return lower.indexOf(c.toLowerCase()) === -1;
      });
      if (!missingCols.length) {
        setText("colCheck", "\u2713 This planner has the lifecycle columns \u2014 close-out and approvals will work.");
        setProp("addCols", "hidden", true);
      } else {
        setText("colCheck", "This planner predates the trip lifecycle. Missing: " +
          missingCols.join(", ") + ". Close-out and approvals need them.");
        setProp("addCols", "hidden", false);
      }
    } catch (e) {
      // Never block Setup on this check; it is diagnostic, not required.
      setText("colCheck", "");
      setProp("addCols", "hidden", true);
    }
  }

  async function addLifecycleColumns() {
    var st = settings();
    if (!st.wbRef || !st.tableName || !missingCols.length) { return; }
    setProp("addCols", "disabled", true);
    var added = [];
    try {
      var token = await GraphData.getToken();
      for (var i = 0; i < missingCols.length; i++) {
        setStatus("work", "Adding \u201c" + missingCols[i] + "\u201d\u2026");
        await GraphData.addTableColumn(token, st.wbRef, st.tableName, missingCols[i]);
        added.push(missingCols[i]);
      }
      setStatus("ok", "Added " + added.join(", ") + " to the planner. Existing rows are untouched \u2014 " +
        "they simply have those cells empty until each trip is closed out.");
      await checkPlannerColumns();
    } catch (e) {
      // Partial success is the likely failure here, so say which ones landed.
      setStatus("error", (added.length ? "Added " + added.join(", ") + ", then stopped: " : "Couldn't add the columns: ") +
        ((e && e.message) || e) + (added.length ? " \u2014 click again to finish the rest." : ""));
      await checkPlannerColumns();
    } finally {
      setProp("addCols", "disabled", false);
    }
  }

  /** Where a workbook actually sits, in words a person recognises. */
  function plannerFolder(url) {
    try {
      var u = decodeURIComponent(String(url || ""));
      var m = /\/(?:Documents|Shared Documents)\/(.*)$/.exec(u);
      if (!m) { return ""; }
      var parts = m[1].split("/");
      parts.pop();                                   // drop the file itself
      var host = /-my\.sharepoint\.com/.test(u) ? "Your OneDrive" : "SharePoint";
      return host + (parts.length ? " \u203a " + parts.join(" \u203a ") : "");
    } catch (e) { return ""; }
  }

  function renderPlannerList() {
    var el = byId("plannerList");
    if (!el) { return; }
    var pl = settings().planners || {};
    var keys = Object.keys(pl);
    el.innerHTML = "";
    if (!keys.length) {
      el.textContent = "No year-specific planners saved yet \u2014 connect a workbook, " +
        "set its fiscal year above, and click Save.";
      return;
    }
    keys.sort().forEach(function (k) {
      var p = pl[k] || {};
      var name = (p.wbRef && p.wbRef.name) || "workbook";
      var row = document.createElement("div");
      row.style.cssText = "padding:4px 0";

      var head = document.createElement("div");
      head.innerHTML = "<b>" + (k === "*" ? "All years" : k) + "</b>: " + name +
        " \u00b7 table " + (p.tableName || "?");
      row.appendChild(head);

      // Naming the file was never enough - "I don't know where it went" is
      // about the FOLDER, and about being able to go and look.
      var folder = plannerFolder(p.wbUrl);
      if (folder) {
        var f = document.createElement("div");
        f.className = "hint";
        f.style.margin = "0";
        f.textContent = folder;
        row.appendChild(f);
      }

      if (p.wbUrl) {
        var open = document.createElement("button");
        open.type = "button";
        open.className = "chip-del";
        open.textContent = "open it";
        open.addEventListener("click", function () {
          if (!openWorkbookUrl(p.wbUrl)) {
            setStatus("info", "Couldn't open it from here. Its address is: " + p.wbUrl);
          }
        });
        row.appendChild(open);
      }
      var del = document.createElement("button");
      del.type = "button";
      del.className = "chip-del";
      del.setAttribute("data-del", k);
      del.textContent = "remove";
      row.appendChild(del);
      el.appendChild(row);
    });
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

    // Say it saved BEFORE doing anything that can fail. Every one of the calls
    // below is a nicety - a column check, an access probe, a note about where
    // the file lives - and any of them throwing used to swallow the
    // confirmation that came after them. The save had happened; the button
    // just looked dead, which is the one thing it must never look.
    var where = (pl[key].wbRef && pl[key].wbRef.name) || pl[key].wbUrl || "the workbook";
    setStatus("info", (key === "*" ? "Saved \u2014 " + where + " is now your catch-all planner."
      : "Saved \u2014 trips dated in " + key + " go to " + where + ".") +
      " It's listed below with a link to open it.");

    [renderPlannerList, checkPlannerColumns, renderCoordSteps, noteInvitesStale,
     function () { noteWhereThePlannerLives(pl[key].wbUrl); },
     function () { verifyPlannerAccess(pl[key].wbRef, "coordinator"); },
    ].forEach(function (fn) {
      try { fn(); } catch (e) {
        try { if (window.console && console.error) { console.error("savePlanner:", e); } }
        catch (e2) { /* nothing left to try */ }
      }
    });
  }

  function refreshFyLine() {
    var el = byId("fyLine");
    if (!el) { return; }
    var st = settings();
    var date = val("eventStart") || val("departDate");
    if (!date) { el.textContent = ""; return; }
    var fy = TravelForm.fiscalLabel(date, st.fyStartMonth, st.fyPrefix);
    var picked = TravelForm.pickPlanner(st.planners, fy);
    if (picked) {
      // Workbook and fiscal-year names come from the user's own files and can
      // contain < > &. Unescaped, "Travel <Draft> FY27.xlsx" silently loses
      // the middle of its own name in this line.
      el.innerHTML = "Files to: <b>" + esc(picked.key === "*" ? "your planner" : picked.key + " planner") + "</b> (" +
        esc((picked.planner.wbRef && picked.planner.wbRef.name) || "workbook") + ") \u2713";
      el.className = "hint fy-ok";
    } else if (st.planners && Object.keys(st.planners).length) {
      el.innerHTML = "\u26a0 No planner saved for <b>" + esc(fy) + "</b> \u2014 connect that year's workbook in Setup before submitting.";
      el.className = "hint fy-warn";
    } else if (st.wbRef) {
      el.innerHTML = "Files to: <b>" + esc(st.wbRef.name || "your planner") + "</b> \u2713";
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

  // Trips get at most this much of the 32 KB roamingSettings ceiling. The rest
  // belongs to the planner list, the org profile and the form-label mapping,
  // and a saveAsync that fails takes the whole trip list with it.
  var TRIPS_BUDGET = 16 * 1024;

  // Forty trips of bare history cost 15 KB on their own, which left nothing
  // for the forms and made "copy to a new request" quietly do nothing on a
  // busy mailbox. Twenty-five is still years of one person's travel, and it
  // leaves room for the thing that is actually useful.
  var TRIPS_KEPT = 25;

  /**
   * Keep every trip; keep as many COPYABLE ones as will fit.
   *
   * The full form is what makes "copy to a new request" work, and it is also
   * the only large thing here. A fixed count was the obvious rule and the
   * wrong one - requests vary several-fold in size, so any number that is safe
   * for a long one wastes most of the budget on a short one. Newest first,
   * because that is the one somebody will want to copy.
   */
  function saveTrips(list) {
    try {
      var keep = list.slice(-TRIPS_KEPT);
      var stubs = keep.map(function (t) {
        var c = {};
        Object.keys(t).forEach(function (k) { if (k !== "form") { c[k] = t[k]; } });
        return c;
      });
      var size = JSON.stringify(stubs).length;
      var out = stubs;
      for (var i = keep.length - 1; i >= 0; i--) {
        if (!keep[i].form) { continue; }
        var cost = JSON.stringify(keep[i].form).length + 8;
        if (size + cost > TRIPS_BUDGET) { break; }
        out[i].form = keep[i].form;
        size += cost;
      }
      Office.context.roamingSettings.set(TRIPS_KEY, JSON.stringify(out));
      persistSettings("your trip list");
    } catch (e) { /* best-effort */ }
  }

  /** How many of the stored trips can still be copied from. */
  function copyableCount() {
    return trips().filter(function (t) { return t && t.form; }).length;
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
      // The whole request, minus its empty boxes, so "copy this one" gives
      // back everything typed rather than the handful of fields the list
      // happens to display.
      form: TravelForm.slimModel(m),
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
    // A coordinator part-way through setup is NOT "set up", whatever the
    // presence of a workbook link suggests. Saying both at once - "You're set
    // up" above "1 of 3 done" - is worse than saying neither.
    if (!s.profileStamp && coordStages(s).some(function (x) { return !x.done; }) &&
        (s.isCoordinator || s.coordEmail || s.wbUrl || s.orgName || s.planners)) {
      el.hidden = true; return;
    }
    if (configured && trips().length) { el.hidden = true; return; }

    el.hidden = false;
    el.innerHTML = "";
    var h = document.createElement("p");
    h.className = "firstrun-h";
    var body = document.createElement("p");
    body.className = "firstrun-b";

    if (!configured) {
      h.textContent = "\ud83d\udc4b First time here?";
      body.innerHTML = "If your travel coordinator has added you, everything you need is already " +
        "in your mailbox \u2014 click below and Travel Desk sets itself up. If you're the " +
        "coordinator, open Setup and connect your planner.";
      var find = document.createElement("button");
      find.className = "primary";
      find.textContent = "Find my setup";
      find.addEventListener("click", function () { findMySetup(); });
      var btn = document.createElement("button");
      btn.textContent = "I'm the coordinator";
      btn.style.marginLeft = "6px";
      btn.addEventListener("click", function () {
        // Declaring the role is what lets the checklist appear for a
        // coordinator without ever appearing for a traveler.
        saveSettings({ isCoordinator: true });
        var setup = byId("setup");
        setup.setAttribute("open", "open");
        try { setup.scrollIntoView({ behavior: "smooth", block: "start" }); }
        catch (e) { setup.scrollIntoView(); }
        renderCoordSteps();
        gotoStage("coordEmail");
      });
      el.appendChild(h); el.appendChild(body); el.appendChild(find); el.appendChild(btn);
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
  /**
   * Reveal the workbook controls and put them on screen.
   *
   * They live inside <details id="setup"> and then <details id="coordSetup">,
   * so a coordinator who is told to "connect the planner in Setup" has to know
   * to open two collapsed sections, the outer of which used to be labelled
   * "travelers: just paste your code". That is how a correct instruction still
   * leaves someone with nowhere to go.
   */
  /**
   * Open the planner workbook itself.
   *
   * "Create my planner" writes a file into OneDrive that the user has never
   * seen and cannot picture - the add-in knows where it went and the person
   * does not. Opening it once, at the moment it is created, answers "where is
   * it?" before it is asked. Connecting an existing workbook does not need
   * this: the user pasted the link, so they already know.
   *
   * openBrowserWindow is the sanctioned route and behaves on the web hosts;
   * window.open is the desktop fallback. If both are blocked the persistent
   * link below is still there, so this never becomes the only way in.
   */
  function openWorkbookUrl(url) {
    if (!url) { return false; }
    try {
      if (Office.context.ui && Office.context.ui.openBrowserWindow) {
        Office.context.ui.openBrowserWindow(url);
        return true;
      }
      window.open(url, "_blank");
      return true;
    } catch (e) { return false; }
  }

  /** A standing "open it" link, so the planner is never more than one click away. */
  function renderPlannerLink() {
    var host = byId("plannerLink");
    if (!host) { return; }
    var st = settings();
    var url = (st.wbUrl || "").trim();
    host.innerHTML = "";
    if (!url) { host.hidden = true; return; }
    host.hidden = false;
    var a = document.createElement("a");
    a.href = "#";
    a.textContent = "\u2197 Open the planner workbook";
    a.addEventListener("click", function (ev) {
      ev.preventDefault();
      if (!openWorkbookUrl(url)) {
        setStatus("info", "Couldn't open it from here \u2014 the link is in the box above.");
      }
    });
    host.appendChild(a);
  }

  /**
   * Which build is actually running.
   *
   * Outlook caches the pane HTML for far longer than the server asks (the
   * files are served max-age=600). The ?v= that busts the JavaScript lives
   * INSIDE that HTML, so a stale page quietly loads stale script - and a fix
   * that is live on the server does nothing for the user, with no way for
   * either of us to tell. Reading the version back out of the script tag and
   * showing it turns that into something checkable in one glance.
   */
  function paneVersion() {
    try {
      var tags = document.querySelectorAll('script[src*="taskpane.js"]');
      var src = tags && tags.length ? tags[tags.length - 1].getAttribute("src") : "";
      var m = /\?v=(\d+)/.exec(src || "");
      return m ? m[1] : "?";
    } catch (e) { return "?"; }
  }

  function renderVersion() {
    var el = byId("paneVersion");
    if (!el) { return; }
    el.textContent = "Travel Desk build " + paneVersion();
  }

  // ------------------------------------------------ fill from this email

  var FILL_LABELS = {
    name: "Name", costCenter: "Cost center",
    division: "Division", bureau: "Bureau",
    otherStaff: "Other travelers", reason: "Reason for travel",
    modePersonal: "Personal vehicle", modeState: "State vehicle",
    modeAir: "Commercial air", cTaxi: "Taxi / rideshare $",
    cMealsB: "Breakfasts included", cMealsL: "Lunches included",
    cMealsD: "Dinners included", cAdditional: "Additional $",
    cAdditionalDesc: "Additional fees are for",
    tp1Project: "3rd party — project number", tp1Max: "3rd party — max reimbursement $",
    tp1Packet: "3rd party packet to attach", tp1Notes: "3rd party — notes",
    tp2Name: "3rd party #2 — entity", tp2Contact: "3rd party #2 — billing contact",
    tp2Project: "3rd party #2 — project number", tp2Max: "3rd party #2 — max reimbursement $",
    tp2Packet: "3rd party #2 packet to attach", tp2Notes: "3rd party #2 — notes",
    tp2Reg: "3rd party #2 covers registration", tp2Lodging: "3rd party #2 covers lodging",
    tp2Air: "3rd party #2 covers airfare/luggage", tp2Meals: "3rd party #2 covers meals",
    tp2Ground: "3rd party #2 covers ground transport",
    event: "Conference / event name", location: "Location",
    eventStart: "Event start", confDates: "Conference dates",
    departDate: "Departure", returnDate: "Return",
    attendeeRole: "Role", meetingLink: "Meeting / event link",
    cRegistration: "Registration $", cLodgingRate: "Rate / night $",
    cLodgingNights: "Lodging nights", cTravelMode: "Travel mode cost $",
    cParking: "Parking $", cLuggage: "Luggage $",
    comments: "Planner comments",
    tp1Name: "3rd party — entity", tp1Contact: "3rd party — billing contact",
    tp1Reg: "3rd party covers registration", tp1Lodging: "3rd party covers lodging",
    tp1Air: "3rd party covers airfare/luggage", tp1Meals: "3rd party covers meals",
    tp1Ground: "3rd party covers ground transport",
  };

  /** Promise wrapper — Office.js is callback-only and this reads better. */
  function officeCall(fn) {
    return new Promise(function (resolve, reject) {
      try {
        fn(function (r) {
          if (r && r.status === Office.AsyncResultStatus.Succeeded) { resolve(r.value); }
          else { reject(new Error((r && r.error && r.error.message) || "Outlook returned no value")); }
        });
      } catch (e) { reject(e); }
    });
  }

  /**
   * Attachment bytes, whatever shape Outlook hands them over in.
   *
   * getAttachmentContentAsync returns one of four formats and they are not
   * interchangeable: base64 for ordinary files, but raw text for a calendar
   * invite or an attached message, and for a cloud attachment just a URL that
   * we cannot follow from here (it is a different origin, and there is no
   * token for it). Treating "url" as a failure rather than an empty file is
   * the difference between "OneDrive attachments can't be scanned" and a
   * silent blank.
   */
  async function attachmentBytes(item, att) {
    var res = await officeCall(function (cb) { item.getAttachmentContentAsync(att.id, cb); });
    var F = Office.MailboxEnums.AttachmentContentFormat;
    if (res.format === F.Base64) { return { bytes: TdAttach.b64ToBytes(res.content) }; }
    if (res.format === F.ICalendar || res.format === F.Eml) { return { text: res.content }; }
    if (res.format === F.Url) {
      return { note: "stored in the cloud, not attached — download it and attach it to scan it" };
    }
    return { note: "unrecognised attachment format" };
  }

  /**
   * Read the open message and fill the form from it.
   *
   * Two rules make this safe to press. Nothing already typed is overwritten
   * unless the box is ticked, and every value that lands is reported with
   * where it came from. A prefill that silently replaced a figure you had
   * checked, or that you could not trace back to a line in the email, would
   * be worse than typing it — you would have to verify the whole form anyway.
   */
  async function fillFromEmail() {
    var item = Office.context.mailbox && Office.context.mailbox.item;
    if (!item) {
      setStatus("error", "Open an email first — Travel Desk reads the message you're looking at.");
      return;
    }
    if (typeof TdMail === "undefined" || typeof TdAttach === "undefined") {
      setStatus("error", "This pane is a cached older version. Close it, reopen it, and check the " +
        "build number at the top.");
      return;
    }
    var btn = byId("fillFromEmail");
    if (btn) { btn.disabled = true; }
    setStatus("work", "Reading the message…");
    try {
      var body = "";
      try {
        body = await officeCall(function (cb) { item.body.getAsync(Office.CoercionType.Html, cb); });
      } catch (e) {
        try { body = await officeCall(function (cb) { item.body.getAsync(Office.CoercionType.Text, cb); }); }
        catch (e2) { body = ""; }
      }

      var got = await scanAttachments(item, function (name, n, total) {
        setStatus("work", "Reading " + name + " (" + n + " of " + total + ")…");
      });
      var scanned = got.scanned, skipped = got.skipped;

      var received = "";
      try {
        var dt = item.dateTimeCreated || item.dateTimeModified;
        if (dt) { received = new Date(dt).toISOString(); }
      } catch (e) { /* a missing timestamp only costs the past-date filter */ }

      // The traveler is almost always the person reading the mail, and Outlook
      // already knows their name - there is no reason to make them type it.
      // Settings win when they exist, because a coordinator filing on someone
      // else's behalf has already said whose trip it is.
      var s = settings();
      var me = "";
      try { me = (Office.context.mailbox.userProfile || {}).displayName || ""; } catch (e) { me = ""; }
      var myName = s.name || me;
      // Your own domain is what separates your travel coordinator from the
      // conference's organizer when both are named in the same message.
      var myDomain = "";
      try {
        var addr = (Office.context.mailbox.userProfile || {}).emailAddress || "";
        myDomain = (addr.split("@")[1] || "").toLowerCase();
      } catch (e) { myDomain = ""; }

      var pre = TdMail.buildPrefill({
        subject: item.subject || "",
        body: body,
        receivedIso: received,
        homeCity: (s.homeCity || ""),
        myName: myName,
        myDomain: myDomain,
        formLabels: s.formLabels || null,
        attachments: scanned,
      });

      if (myName) { pre.fields.name = myName; pre.sources.name = s.name ? "your settings" : "your Outlook profile"; }
      // Saved settings outrank the signature block: you have already corrected
      // these once, and a signature is only ever an inference.
      ["costCenter", "division", "bureau"].forEach(function (k) {
        if (s[k]) { pre.fields[k] = s[k]; pre.sources[k] = "your settings"; }
      });

      applyPrefill(pre, scanned, skipped);
    } catch (err) {
      setStatus("error", "Couldn't read this message: " + (err && err.message ? err.message : err));
    } finally {
      if (btn) { btn.disabled = false; }
    }
  }

  function applyPrefill(pre, scanned, skipped) {
    var overwrite = isChecked("fillOverwrite");
    var filled = [], kept = [];

    Object.keys(pre.fields).forEach(function (id) {
      var el = byId(id);
      if (!el) { return; }
      var next = pre.fields[id];

      // Checkboxes are booleans, and assigning .value to one sets an attribute
      // nobody reads while leaving the box unticked - the third-party cost
      // categories would have looked filled in the report and been empty on
      // the form.
      if (el.type === "checkbox") {
        if (el.checked === !!next) { return; }
        if (el.checked && !overwrite) { return; }
        el.checked = !!next;
        filled.push({ id: id, value: "yes", source: pre.sources[id] || "" });
        return;
      }

      var current = String(el.value || "").trim();
      if (current && !overwrite) {
        if (current !== String(next)) { kept.push({ id: id, mine: current, theirs: next }); }
        return;
      }
      el.value = next;
      filled.push({ id: id, value: next, source: pre.sources[id] || "" });
    });

    // A filled section that is still collapsed reads as not filled at all.
    ["tp1", "tp2"].forEach(function (sec) {
      if (filled.some(function (f) { return f.id.indexOf(sec) === 0; })) {
        setAttrIf(sec, "open", "open");
      }
    });

    // The costs only re-total on 'input', which assigning .value does not fire.
    try { refreshTotal(); refreshFyLine(); } catch (e) { /* cosmetic */ }

    renderFillReport(pre, filled, kept, scanned, skipped);
    if (filled.length) {
      setStatus("info", "Filled " + filled.length + " field" + (filled.length === 1 ? "" : "s") +
        " from this email. Check them against the message before you submit — " +
        "each one shows where it came from.");
    } else if (kept.length) {
      setStatus("info", "Everything it found was already filled in. Tick “Replace what I've " +
        "already typed” if you want it to overwrite.");
    } else {
      setStatus("error", "Nothing usable in this message. See the notes below — the fields it " +
        "couldn't find are listed with why.");
    }
  }

  function renderFillReport(pre, filled, kept, scanned, skipped) {
    var box = byId("fillReport");
    if (!box) { return; }
    box.hidden = false;
    box.innerHTML = "";

    function head(t) {
      var p = document.createElement("p");
      p.className = "f-head"; p.textContent = t; box.appendChild(p); return p;
    }
    function note(t) {
      var p = document.createElement("p");
      p.className = "f-note"; p.textContent = t; box.appendChild(p); return p;
    }

    if (filled.length) {
      head("Filled in — check each against the email:");
      var ul = document.createElement("ul");
      filled.forEach(function (f) {
        var li = document.createElement("li");
        var n = document.createElement("span");
        n.className = "f-name";
        n.textContent = (FILL_LABELS[f.id] || f.id) + ": " + f.value;
        li.appendChild(n);
        if (f.source) {
          var s = document.createElement("span");
          s.className = "f-src";
          s.textContent = "from " + f.source;
          li.appendChild(s);
        }
        ul.appendChild(li);
      });
      box.appendChild(ul);
    }

    // A second candidate for the destination is common and cheap to offer:
    // "Denver, CO" and the overflow hotel in "Aurora, CO" both look like the
    // venue to a regex, and only the person reading it knows which.
    if (pre.alternates && pre.alternates.location && pre.alternates.location.length) {
      var wrap = document.createElement("div");
      wrap.className = "f-alt";
      var lbl = document.createElement("span");
      lbl.textContent = "Other places mentioned: ";
      wrap.appendChild(lbl);
      pre.alternates.location.forEach(function (alt) {
        var b = document.createElement("button");
        b.type = "button";
        b.textContent = alt;
        b.addEventListener("click", function () {
          setVal("location", alt);
          setStatus("info", "Destination set to " + alt + ".");
        });
        wrap.appendChild(b);
      });
      box.appendChild(wrap);
    }

    if (kept.length) {
      head("Left alone — you'd already typed something:");
      var ul2 = document.createElement("ul");
      kept.forEach(function (k) {
        var li = document.createElement("li");
        li.textContent = (FILL_LABELS[k.id] || k.id) + ": kept “" + k.mine +
          "” (the email says “" + k.theirs + "”)";
        ul2.appendChild(li);
      });
      box.appendChild(ul2);
    }

    renderSuggestions(pre.suggest || {});

    if (scanned.length) { note("Read: " + scanned.map(function (a) { return a.name; }).join(", ")); }
    if (skipped.length) { note("Not read: " + skipped.join("; ")); }
    (pre.notes || []).forEach(function (n) { note(n); });
  }

  /**
   * Settings the message revealed, offered as a button rather than applied.
   *
   * The coordinator's address and your home city are not facts about this
   * trip — they persist and shape every future one. Writing them silently
   * from one email would mean a stray "send this to Jane" line quietly
   * redirecting your travel paperwork, and you would have no reason to look.
   */
  function renderSuggestions(sug) {
    var box = byId("fillReport");
    if (!box) { return; }
    var s = settings();
    var offers = [];
    if (sug.coordEmail && sug.coordEmail !== s.coordEmail) {
      offers.push({
        key: "coordEmail", value: sug.coordEmail,
        label: (s.coordEmail ? "Change your travel coordinator to " : "Set your travel coordinator to ") +
          sug.coordEmail,
      });
    }
    if (sug.homeCity && sug.homeCity !== s.homeCity) {
      offers.push({
        key: "homeCity", value: sug.homeCity,
        label: "Remember that you're based in " + sug.homeCity,
      });
    }
    if (!offers.length) { return; }

    var h = document.createElement("p");
    h.className = "f-head";
    h.textContent = "Worth saving — this changes settings, so it's your call:";
    box.appendChild(h);
    var wrap = document.createElement("div");
    wrap.className = "f-alt";
    offers.forEach(function (o) {
      var b = document.createElement("button");
      b.type = "button";
      b.textContent = o.label;
      b.addEventListener("click", function () {
        var patch = {};
        patch[o.key] = o.value;
        saveSettings(patch);
        if (o.key === "coordEmail") { setVal("coordEmail", o.value); }
        b.disabled = true;
        b.textContent = "Saved ✓";
        setStatus("info", "Saved. " + (o.key === "coordEmail"
          ? "Travel Authorization drafts will be addressed to " + o.value + "."
          : "Your own city won't be offered as a destination again."));
      });
      wrap.appendChild(b);
    });
    box.appendChild(wrap);
  }

  /**
   * Show the stages until they are done, and never to a traveler.
   *
   * Somebody set up from a code is not configuring anything - showing them a
   * coordinator's checklist would be noise at best and alarming at worst.
   */
  function renderCoordSteps() {
    var box = byId("coordSteps");
    if (!box) { return; }
    var st = settings();
    if (st.profileStamp) { box.hidden = true; return; }   // set up from a code
    // On a fresh install we do not yet know which role this is, and a
    // coordinator's checklist in front of somebody who is about to click
    // "Find my setup" is worse than nothing.
    var started = st.isCoordinator || st.coordEmail || st.wbUrl || st.orgName ||
      st.planners || st.invitesSentAt;
    if (!started) { box.hidden = true; return; }
    var stages = coordStages(st);
    if (stages.every(function (x) { return x.done; })) { box.hidden = true; return; }

    box.hidden = false;
    box.innerHTML = "";
    var h = document.createElement("h3");
    h.textContent = "Setting up for your team";
    var sub = document.createElement("p");
    sub.className = "steps-sub";
    var doneCount = stages.filter(function (x) { return x.done; }).length;
    sub.textContent = doneCount + " of " + stages.length + " done \u00b7 " +
      "each one has to come before the next.";
    box.appendChild(h); box.appendChild(sub);

    var ol = document.createElement("ol");
    var firstOpen = true;
    stages.forEach(function (stage, i) {
      var li = document.createElement("li");
      var isNow = !stage.done && firstOpen && stage.needs !== false;
      if (stage.done) { li.className = "done"; }
      else if (isNow) { li.className = "now"; firstOpen = false; }

      var mark = document.createElement("span");
      mark.className = "step-mark";
      mark.textContent = stage.done ? "\u2713" : String(i + 1);

      var body = document.createElement("div");
      body.className = "step-body";
      var nm = document.createElement("span");
      nm.className = "step-name";
      nm.textContent = stage.name;
      body.appendChild(nm);
      if (!stage.done && stage.why) {
        var why = document.createElement("span");
        why.className = "step-why";
        why.textContent = stage.needs === false
          ? "Ready once the two above are done."
          : stage.why;
        body.appendChild(why);
      }
      if (!stage.done && stage.needs !== false) {
        var go = document.createElement("button");
        go.type = "button";
        go.textContent = "Take me there";
        go.addEventListener("click", function () { gotoStage(stage.action); });
        body.appendChild(go);
      }
      li.appendChild(mark); li.appendChild(body);
      ol.appendChild(li);
    });
    box.appendChild(ol);
  }

  /** Open whatever is collapsed around a control, then put it on screen. */
  function gotoStage(action) {
    setAttrIf("setup", "open", "open");
    setAttrIf("coordSetup", "open", "open");
    var target = action === "planner" ? byId("coordSetup") : byId(action);
    if (!target) { return; }
    try { target.scrollIntoView({ behavior: "smooth", block: "center" }); }
    catch (e) { target.scrollIntoView(); }
    if (target.focus && action !== "planner") {
      try { target.focus({ preventScroll: true }); } catch (e) { target.focus(); }
    }
  }

  // ------------------------------------------- is my setup out of date?

  var STALE_CHECK_HOURS = 24;

  /**
   * Notice when the coordinator has moved on without you.
   *
   * A traveler's settings are a copy taken the day they were set up. Move the
   * planner and they keep writing to the old workbook - no error, no clue,
   * and the coordinator simply stops seeing them. This is the same silent
   * shape as the write-permission gap, and the only way to close it is for
   * the traveler's own pane to go and look.
   *
   * Cheap on purpose: at most once a day, only for people set up FROM a code,
   * and the date comparison is free metadata - a body is fetched only when
   * something genuinely newer exists.
   */
  async function checkSetupFreshness() {
    var s = settings();
    if (!s.profileStamp) { return; }          // coordinators set themselves up
    var last = Date.parse(s.staleCheckedAt || 0) || 0;
    if (Date.now() - last < STALE_CHECK_HOURS * 36e5) { return; }
    saveSettings({ staleCheckedAt: new Date().toISOString() });

    try {
      var token = await GraphData.getToken();
      var msgs = await GraphData.setupInvites(token, "Travel Desk setup");
      var newer = TravelForm.newerInvite(msgs, s.profileAppliedAt);
      if (!newer) { return; }

      var body = await GraphData.messageBody(token, newer.id);
      var code = TravelForm.extractSetupCode(body);
      if (!code || TravelForm.profileStamp(code) === s.profileStamp) { return; }

      offerSetupUpdate(code, newer);
    } catch (e) { /* a check that cannot run is not news */ }
  }

  function offerSetupUpdate(code, invite) {
    var host = byId("firstRun");
    if (!host) { return; }
    host.hidden = false;
    host.innerHTML = "";
    var h = document.createElement("p");
    h.className = "firstrun-h";
    h.textContent = "\u26a0\ufe0f Your travel coordinator has sent newer settings";
    var b = document.createElement("p");
    b.className = "firstrun-b";
    b.textContent = "The planner, the coordinator address or the fiscal-year rules have " +
      "changed since you were set up" +
      (invite && invite.receivedDateTime ? " (" + invite.receivedDateTime.slice(0, 10) + ")" : "") +
      ". Until you take them, your requests keep going to the old planner \u2014 where " +
      "nobody is looking.";
    var go = document.createElement("button");
    go.className = "primary";
    go.textContent = "Use the new settings";
    go.addEventListener("click", function () {
      profileApply(code);
      saveSettings({
        profileStamp: TravelForm.profileStamp(code),
        profileAppliedAt: (invite && invite.receivedDateTime) || new Date().toISOString(),
      });
      host.hidden = true;
      renderFirstRun();
    });
    host.appendChild(h); host.appendChild(b); host.appendChild(go);
  }

  // ------------------------------------- find travel emails in the mailbox

  var TRAVEL_LOOKBACK_DAYS = 180;

  /**
   * Search the mailbox rather than only reading what is open.
   *
   * "Fill from this email" assumes you already found the confirmation. Half
   * the time the trip you are filing for is three weeks down the inbox, and
   * hunting for it is the part nobody wants to do.
   */
  async function findTravelEmails() {
    if (typeof TdMail === "undefined") {
      setStatus("error", "This pane is a cached older version \u2014 close it, reopen it, " +
        "and check the build number at the top.");
      return;
    }
    var btn = byId("findTravel");
    if (btn) { btn.disabled = true; }
    setStatus("work", "Looking through your inbox\u2026");
    try {
      var token = await GraphData.getToken();
      var since = new Date(Date.now() - TRAVEL_LOOKBACK_DAYS * 864e5).toISOString();
      var msgs = await GraphData.bookingEmails(token, since);
      var s = settings();
      var hits = TdMail.findTravelEmails(msgs, {
        todayIso: new Date().toISOString(),
        homeCity: s.homeCity || "",
        trustedSenders: (val("bookingSenders") || "").split(","),
        filed: trips(),
      });
      renderTravelFinds(hits, msgs.length);
    } catch (e) {
      setStatus("error", "Couldn't search your inbox: " + ((e && e.message) || e));
    } finally {
      if (btn) { btn.disabled = false; }
    }
  }

  function renderTravelFinds(hits, scanned) {
    var box = byId("travelFinds");
    if (!box) { return; }
    box.hidden = false;
    box.innerHTML = "";

    var head = document.createElement("p");
    head.className = "f-head";
    head.textContent = hits.length
      ? hits.length + " look like travel, out of " + scanned + " messages"
      : "Nothing in the last " + TRAVEL_LOOKBACK_DAYS + " days looks like a trip";
    box.appendChild(head);

    if (!hits.length) {
      var p = document.createElement("p");
      p.className = "f-note";
      p.textContent = "Open the confirmation yourself and use \u201cFill from this email\u201d \u2014 " +
        "that reads whatever you are looking at, whether or not this recognises it.";
      box.appendChild(p);
      setStatus("info", "No travel emails found in the last " + TRAVEL_LOOKBACK_DAYS + " days.");
      return;
    }

    var ul = document.createElement("ul");
    ul.className = "learn-list";
    hits.slice(0, 25).forEach(function (h) {
      var li = document.createElement("li");
      var t = document.createElement("span");
      t.className = "f-name";
      t.textContent = h.subject;
      if (h.alreadyFiled) {
        var tag = document.createElement("span");
        tag.className = "learn-new";
        tag.textContent = "already filed";
        t.appendChild(tag);
      }
      var meta = document.createElement("span");
      meta.className = "f-src";
      meta.textContent = (h.receivedDateTime ? h.receivedDateTime.slice(0, 10) + " \u00b7 " : "") +
        h.from + " \u00b7 " + h.why.slice(0, 2).join(", ");
      var use = document.createElement("button");
      use.type = "button";
      use.className = "chip-del";
      use.textContent = "Use this";
      use.addEventListener("click", function () { useTravelEmail(h); });
      li.appendChild(t); li.appendChild(meta); li.appendChild(use);
      ul.appendChild(li);
    });
    box.appendChild(ul);
    setStatus("info", "Pick one and it fills the form the same way the open message would. " +
      "Ones you have already filed are marked, not hidden \u2014 they are the ones worth copying.");
  }

  /** Read a message the user is not looking at, then prefill from it. */
  async function useTravelEmail(hit) {
    setStatus("work", "Reading \u201c" + hit.subject + "\u201d\u2026");
    try {
      var token = await GraphData.getToken();
      var msg = await GraphData.messageFull(token, hit.id);
      var scanned = [], skipped = [];
      if (msg.hasAttachments) {
        setStatus("work", "Reading its attachments\u2026");
        var atts = await GraphData.messageAttachments(token, hit.id);
        for (var i = 0; i < atts.length; i++) {
          var a = atts[i];
          if (!a.contentBytes) { skipped.push(a.name + (a.note ? " \u2014 " + a.note : "")); continue; }
          try {
            var out = await TdAttach.attachmentText(a.name, TdAttach.b64ToBytes(a.contentBytes));
            if (out.text) { scanned.push({ name: a.name, text: out.text }); }
            else { skipped.push(a.name + (out.note ? " \u2014 " + out.note : "")); }
          } catch (e) { skipped.push(a.name + " \u2014 couldn't be read"); }
        }
      }

      var s = settings();
      var me = "";
      try { me = (Office.context.mailbox.userProfile || {}).displayName || ""; } catch (e) { me = ""; }
      var myDomain = "";
      try {
        var addr = (Office.context.mailbox.userProfile || {}).emailAddress || "";
        myDomain = (addr.split("@")[1] || "").toLowerCase();
      } catch (e) { myDomain = ""; }

      var pre = TdMail.buildPrefill({
        subject: msg.subject, body: msg.body,
        receivedIso: msg.receivedDateTime,
        homeCity: s.homeCity || "", myName: s.name || me, myDomain: myDomain,
        formLabels: s.formLabels || null,
        attachments: scanned,
      });
      if (s.name || me) {
        pre.fields.name = s.name || me;
        pre.sources.name = s.name ? "your settings" : "your Outlook profile";
      }
      ["costCenter", "division", "bureau"].forEach(function (k) {
        if (s[k]) { pre.fields[k] = s[k]; pre.sources[k] = "your settings"; }
      });
      applyPrefill(pre, scanned, skipped);
    } catch (e) {
      setStatus("error", "Couldn't read that message: " + ((e && e.message) || e));
    }
  }

  // -------------------------------------------- teach it your own form

  /**
   * Every attachment on the open message, decoded to text.
   *
   * Shared by "Fill from this email" and "Learn from this email's attachment",
   * because the two need exactly the same thing and a second copy of this
   * would drift out of step with the first.
   */
  async function scanAttachments(item, onProgress) {
    var atts = (item.attachments || []).filter(function (a) { return a && !a.isInline; });
    var scanned = [], skipped = [];
    for (var i = 0; i < atts.length; i++) {
      var a = atts[i];
      if (onProgress) { onProgress(a.name, i + 1, atts.length); }
      try {
        var got = await attachmentBytes(item, a);
        if (got.text) { scanned.push({ name: a.name, text: got.text }); continue; }
        if (got.note) { skipped.push(a.name + " — " + got.note); continue; }
        var out = await TdAttach.attachmentText(a.name, got.bytes);
        if (out.text) { scanned.push({ name: a.name, text: out.text }); }
        else { skipped.push(a.name + (out.note ? " — " + out.note : "")); }
      } catch (err) {
        skipped.push(a.name + " — couldn't be opened");
      }
    }
    return { scanned: scanned, skipped: skipped };
  }

  var learnFound = [];   // [{label, field, known}] from the last scan

  async function learnScan() {
    var item = Office.context.mailbox && Office.context.mailbox.item;
    if (!item) { setStatus("error", "Open the email your blank form is attached to first."); return; }
    if (typeof TdMail === "undefined" || typeof TdAttach === "undefined") {
      setStatus("error", "This pane is a cached older version — close it, reopen it, and check the build number at the top.");
      return;
    }
    var btn = byId("learnScan");
    if (btn) { btn.disabled = true; }
    setStatus("work", "Reading the attachment…");
    try {
      var res = await scanAttachments(item, function (name, n, total) {
        setStatus("work", "Reading " + name + " (" + n + " of " + total + ")…");
      });
      if (!res.scanned.length) {
        setStatus("error", "Nothing readable was attached." +
          (res.skipped.length ? " Skipped: " + res.skipped.join("; ") : ""));
        return;
      }
      // The form is the attachment with the most labels in it, not the
      // biggest one — an agenda can be far longer and have none.
      var aliases = settings().formLabels || {};
      var best = null;
      res.scanned.forEach(function (a) {
        var found = TdMail.discoverLabels(a.text, aliases);
        if (!best || found.length > best.found.length) { best = { name: a.name, found: found }; }
      });
      learnFound = best.found;
      renderLearn(best.name);
      setStatus("info", "Found " + best.found.length + " labels in " + best.name +
        ". Check the ones marked new, then save.");
    } catch (e) {
      setStatus("error", "Couldn't read the attachment: " + ((e && e.message) || e));
    } finally {
      if (btn) { btn.disabled = false; }
    }
  }

  function renderLearn(fileName) {
    var box = byId("learnList");
    if (!box) { return; }
    box.innerHTML = "";
    if (!learnFound.length) { box.textContent = "No labels found in " + fileName + "."; return; }

    var head = document.createElement("p");
    head.className = "f-head";
    var newCount = learnFound.filter(function (f) { return !f.known; }).length;
    head.textContent = fileName + " — " + learnFound.length + " labels, " +
      (learnFound.length - newCount) + " recognised, " + newCount + " for you to assign";
    box.appendChild(head);

    var ul = document.createElement("ul");
    ul.className = "learn-list";
    learnFound.forEach(function (f, i) {
      var li = document.createElement("li");
      var name = document.createElement("span");
      name.className = "f-name";
      name.textContent = f.label;
      if (!f.known) {
        var tag = document.createElement("span");
        tag.className = "learn-new";
        tag.textContent = "new";
        name.appendChild(tag);
      }
      var sel = document.createElement("select");
      sel.setAttribute("data-idx", String(i));
      var none = document.createElement("option");
      none.value = "";
      none.textContent = "(ignore this label)";
      sel.appendChild(none);
      TdMail.MAPPABLE.forEach(function (m) {
        var o = document.createElement("option");
        o.value = m.field;
        o.textContent = m.text;
        if (m.field === f.field) { o.selected = true; }
        sel.appendChild(o);
      });
      sel.addEventListener("change", function () { learnFound[i].field = sel.value; });
      li.appendChild(name);
      li.appendChild(sel);
      ul.appendChild(li);
    });
    box.appendChild(ul);
    setProp("learnSave", "hidden", false);
    setProp("learnClear", "hidden", false);
  }

  /**
   * Save only what DIFFERS from what Travel Desk already worked out.
   *
   * Storing all 33 labels would mean this org's copy of the built-in wording
   * silently overrides future improvements to it, and roamingSettings has a
   * 32 KB ceiling shared with everything else. Only the corrections are kept.
   */
  function learnSave() {
    var map = {};
    learnFound.forEach(function (f) {
      // Only a DIFFERENCE from the built-in table is worth keeping - including
      // an explicit blank, which means "stop reading this label".
      if (f.field !== f.builtin) { map[f.label] = f.field; }
    });
    saveSettings({ formLabels: map });
    var n = Object.keys(map).length;
    setStatus("info", n
      ? "Saved " + n + " label" + (n === 1 ? "" : "s") + " for your form. " +
        "Regenerate your setup code so travelers get it too."
      : "Nothing to save — every label already matched the built-in wording.");
  }

  function learnClear() {
    saveSettings({ formLabels: {} });
    learnFound = [];
    setText("learnList", "Cleared — back to the built-in wording.");
    setProp("learnSave", "hidden", true);
    setProp("learnClear", "hidden", true);
    setStatus("info", "Your form's custom wording was removed.");
  }

  function openPlannerSetup() {
    setAttrIf("setup", "open", "open");
    setAttrIf("coordSetup", "open", "open");
    var el = byId("coordSetup") || byId("setup");
    if (el && el.scrollIntoView) {
      try { el.scrollIntoView({ behavior: "smooth", block: "start" }); }
      catch (e) { el.scrollIntoView(); }   // older webviews take no options
    }
  }

  function refreshCoordVisibility() {
    var el = byId("coord");
    if (!el) { return; }
    var st = settings();
    el.hidden = !(st.tableName && (st.wbUrl || (st.planners && Object.keys(st.planners).length)));
  }

  var coordRecords = [];

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
      coordRecords = records;
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

  /** One queue row with its approve action. */
  function approvalRow(host, r, stage) {
    var box = document.createElement("div");
    box.style.cssText = "border:1px solid var(--line);border-radius:8px;padding:8px 10px;margin:6px 0";
    var head = document.createElement("div");
    var detail = (r.date || "(no date)") + " \u00b7 " + r.traveler +
      (r.division ? " (" + r.division + ")" : "") + " \u00b7 " + r.event;
    if (stage === "close") {
      detail += " \u00b7 est " + money0(r.cost) + " \u2192 actual " + money0(r.actualCost) +
        (r.variancePct != null ? " (" + (r.variancePct > 0 ? "+" : "") + r.variancePct + "%)" : "");
    } else if (r.cost) {
      detail += " \u00b7 " + money0(r.cost);
    }
    head.textContent = detail;
    if (r.overBudget) { head.style.cssText = "color:var(--err-fg);font-weight:600"; }
    box.appendChild(head);

    var btn = document.createElement("button");
    btn.className = "primary";
    btn.style.marginTop = "6px";
    btn.textContent = stage === "close" ? "Mark seen" : "Approve travel";
    btn.addEventListener("click", async function () {
      btn.disabled = true;
      var me = (Office.context.mailbox.userProfile || {}).displayName || "";
      var ok = await approveRow({
        name: r.traveler, event: r.event, eventStart: r.date,
      }, stage, me);
      if (ok) { btn.textContent = stage === "close" ? "Seen \u2713" : "Approved \u2713"; }
      else { btn.disabled = false; }
    });
    box.appendChild(btn);
    host.appendChild(box);
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

    var q = TravelCoord.queues(coordRecords);
    var bt = TravelCoord.budgetTruth(coordRecords, { fy: val("coordFy").trim() });

    // The gauge first: the whole point is "is travel costing roughly what we
    // planned", and that answer should not sit below two work queues.
    section("Gauge" + (val("coordFy").trim() ? " \u00b7 " + val("coordFy").trim() : ""));
    if (bt.tripsWithActuals) {
      line("Planned " + money0(bt.estimatedClosed) + " on " + bt.tripsWithActuals +
        " completed trip(s); actually spent " + money0(bt.actual) + " \u2014 " +
        (bt.variancePct === 0 ? "on the nose."
          : Math.abs(bt.variancePct) + "% " + (bt.variancePct > 0 ? "over." : "under.")),
        Math.abs(bt.variancePct) > 20 ? "warn" : null);
    } else {
      line("No completed trips yet \u2014 the gauge fills in as trips are closed out.");
    }
    line(money0(bt.estimatedAll - bt.estimatedClosed) + " planned on " + bt.tripsWithout +
      " trip(s) still to come.");

    section("Waiting on you \u2014 approve travel (" + q.awaitingApproval.length + ")");
    if (!q.awaitingApproval.length) { line("Nothing awaiting approval."); }
    q.awaitingApproval.slice(0, 15).forEach(function (r) {
      approvalRow(host, r, "travel");
    });

    if (q.awaitingClose.length) {
      section("Closed out recently (" + q.awaitingClose.length + ")");
      line("Nothing to do here unless something looks wrong \u2014 mark them seen to tidy the list.");
      q.awaitingClose.slice(0, 15).forEach(function (r) {
        approvalRow(host, r, "close");
      });
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

  /**
   * Traveller closes out a trip: real costs replace the guess on the SAME
   * planner row. The estimate is deliberately left in place — the gap
   * between the two is the number that defends next year's budget.
   */
  async function submitActuals(trip, actualTotal, note) {
    return writeBackRow(trip, function (values, idx) {
      if (idx.actualCost == null) { return false; }   // planner lacks the column
      values[idx.actualCost] = String(Math.round(actualTotal));
      if (idx.status != null) { values[idx.status] = TravelForm.STATUS.ACTUALS; }
      if (note && idx.comments != null) {
        values[idx.comments] = (values[idx.comments] ? values[idx.comments] + " | " : "") + note;
      }
      return true;
    }, "Actual costs recorded on the planner — your coordinator sees them next.");
  }

  /** Coordinator records an approval on the row (recorded, not enforced). */
  async function approveRow(trip, stage, approver) {
    var status = stage === "close" ? TravelForm.STATUS.CLOSED : TravelForm.STATUS.APPROVED;
    return writeBackRow(trip, function (values, idx) {
      if (idx.status == null) { return false; }
      values[idx.status] = status;
      if (idx.approvedBy != null) { values[idx.approvedBy] = approver || ""; }
      if (idx.approvedDate != null) {
        values[idx.approvedDate] = new Date().toISOString().slice(0, 10);
      }
      return true;
    }, stage === "close" ? "Closed out on the planner." : "Approved on the planner.");
  }

  function markBookedEverywhere(trip) {
    return writeBackRow(trip, function (values, idx) {
      if (idx.status == null) { return false; }
      values[idx.status] = "Booked";
      return true;
    }, "Marked booked here and in the shared planner.");
  }

  /** Coordinator: address people instead of handing out a code. */
  async function sendInvites() {
    var emails = TravelForm.extractEmails(val("inviteTo"));
    if (!emails.length) {
      byId("inviteInfo").textContent = "Add at least one email address first.";
      return;
    }
    byId("inviteSend").disabled = true;
    try {
      var st = settings();
      var out = {};
      PROFILE_FIELDS.forEach(function (k) { if (st[k]) { out[k] = st[k]; } });
      var gaps = setupGaps();
      if (gaps.length) {
        byId("inviteInfo").style.color = "var(--err-fg)";
        byId("inviteInfo").textContent = "Not ready to invite anyone yet: " + gaps.join("; ") + ".";
        return;
      }
      var code = btoa(unescape(encodeURIComponent(JSON.stringify(out))));
      var prof = Office.context.mailbox.userProfile || {};
      setStatus("work", "Writing the invitation\u2026");
      var token = await GraphData.getToken();
      var draft = await GraphData.createDraft(token, emails,
        TravelForm.setupSubject(st.orgName),
        TravelForm.setupInviteHtml({
          code: code, orgName: st.orgName,
          coordName: prof.displayName, coordEmail: prof.emailAddress,
        }));
      saveSettings({ invitesSentAt: new Date().toISOString(), invitesStamp: TravelForm.profileStamp(code) });
      renderCoordSteps();
      byId("inviteInfo").style.color = "";
      byId("inviteInfo").textContent = "Invitation drafted for " + emails.length +
        " traveler(s) — review it and press Send.";
      setStatus("info", "Invitation ready in your Drafts for " + emails.join(", ") +
        ". Nothing is sent until you press Send.");
      try {
        if (draft && draft.webLink) {
          Office.context.ui.openBrowserWindow
            ? Office.context.ui.openBrowserWindow(draft.webLink)
            : window.open(draft.webLink, "_blank");
        }
      } catch (e) { /* it's in Drafts either way */ }
    } catch (e) {
      setStatus("error", "Couldn't write the invitation: " + ((e && e.message) || e));
    } finally {
      byId("inviteSend").disabled = false;
    }
  }

  /** Traveller: let their own mailbox carry the setup to them. */
  async function findMySetup() {
    setStatus("work", "Looking for a setup from your coordinator\u2026");
    try {
      var token = await GraphData.getToken();
      var msgs = await GraphData.setupInvites(token, "Travel Desk setup");
      if (!msgs.length) {
        setStatus("error", "No setup invitation found in your mailbox. Ask your coordinator to " +
          "send one \u2014 or paste a setup code in Setup if you were given one.");
        return;
      }
      // search results carry metadata only; fetch bodies newest-first
      for (var i = 0; i < Math.min(msgs.length, 5); i++) {
        msgs[i].body = await GraphData.messageBody(token, msgs[i].id);
      }
      var invite = TravelForm.pickInvite(msgs);
      if (!invite) {
        setStatus("error", "Found a setup message but no settings inside it — ask your " +
          "coordinator to resend.");
        return;
      }
      var from = ((invite.from || {}).emailAddress || {}).name ||
                 ((invite.from || {}).emailAddress || {}).address || "your coordinator";
      var code = TravelForm.extractSetupCode(invite.body);
      profileApply(code);
      saveSettings({
        profileStamp: TravelForm.profileStamp(code),
        profileAppliedAt: invite.receivedDateTime || new Date().toISOString(),
      });
      setStatus("info", "Set up from " + from + "'s invitation. Fill in the form below and " +
        "click Create travel request.");
    } catch (e) {
      setStatus("error", "Couldn't read your setup: " + ((e && e.message) || e));
    }
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

  /** Keep the close-out picker in step with the traveller's own trips. */
  function renderCloseOptions() {
    var sel = byId("closeTrip");
    if (!sel) { return; }
    var list = trips();
    sel.innerHTML = "";
    var o0 = document.createElement("option");
    o0.value = "";
    o0.textContent = list.length ? "Pick one of your trips\u2026" : "No trips filed yet";
    sel.appendChild(o0);
    list.slice().reverse().forEach(function (t) {
      var o = document.createElement("option");
      o.value = t.id;
      o.textContent = (t.event || "trip") + (t.eventStart ? " \u00b7 " + t.eventStart : "");
      sel.appendChild(o);
    });
  }

  async function closeOutTrip() {
    var id = val("closeTrip");
    var amount = Number(val("closeActual"));
    if (!id) { setStatus("error", "Pick which trip you're closing out."); return; }
    if (!amount || amount <= 0) { setStatus("error", "Enter what the trip actually cost."); return; }
    var trip = trips().filter(function (t) { return t.id === id; })[0];
    if (!trip) { setStatus("error", "That trip is no longer in your list."); return; }
    byId("closeSubmit").disabled = true;
    try {
      var ok = await submitActuals(trip, amount, val("closeNote").trim());
      if (ok) {
        byId("closeActual").value = "";
        byId("closeNote").value = "";
      } else {
        setStatus("error", "Couldn't find your row, or this planner predates the trip lifecycle. " +
          "A coordinator can fix that in one click: Setup \u2192 \u201cAdd the missing " +
          "lifecycle columns\u201d. Existing rows aren't touched.");
      }
    } finally {
      byId("closeSubmit").disabled = false;
    }
  }

  function renderTrips() {
    var el = byId("tripsList");
    if (!el) { return; }
    renderCloseOptions();
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
      // A trip with no date printed an empty separator - "Winter Ops · · Requested".
      var when = t.eventStart || t.departDate || "";
      return "<div style=\"border-top:1px solid var(--line);padding:6px 0\"><b>" + (t.event || "trip") + "</b> \u00b7 " +
        (when ? when + " \u00b7 " : "") + chip +
        (t.bookingLink ? ' \u00b7 <a href="' + t.bookingLink + '" target="_blank" rel="noopener">open booking email</a>' : "") +
        (t.status !== "booked" ? ' \u00b7 <button type="button" class="chip-del" data-book="' + i + '">mark booked</button>' : "") +
        ' \u00b7 <button type="button" class="chip-del" data-copy="' + i + '">copy to a new request</button>' +
        ' \u00b7 <button type="button" class="chip-del" data-cal="' + i + '">add to calendar</button>' +
        "</div>";
    }).join("");
    el.querySelectorAll("[data-cal]").forEach(function (b) {
      b.addEventListener("click", function () {
        var t = trips()[Number(b.getAttribute("data-cal"))];
        if (t) { addTripToCalendar(t); }
      });
    });
    el.querySelectorAll("[data-copy]").forEach(function (b) {
      b.addEventListener("click", function () {
        copyTrip(Number(b.getAttribute("data-copy")));
      });
    });
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

  var PARTIAL_NOTE = "Only the basics came across \u2014 this trip is old enough that " +
    "the rest of its form was let go to stay inside what Outlook lets an add-in store.";

  var COPY_LABELS = {
    eventStart: "Event start", departDate: "Departure", returnDate: "Return",
    confDates: "Conference dates",
  };

  /**
   * Start a new request from one already filed.
   *
   * Most trips are a variation on a previous one - the same conference a year
   * later, the same site visit with different dates - and retyping twenty
   * fields to change three is the tax this was built to remove.
   */
  function copyTrip(i) {
    var list = trips();
    var t = list[i];
    if (!t) { return; }

    // Two kinds of trip copy only partly: one filed before the full form was
    // ever stored, and one old enough that its form was dropped to stay inside
    // the settings budget. Both still give the basics, and both say so rather
    // than handing back a form that looks complete and is not.
    var partial = !t.form;
    var saved = t.form || {
      name: t.name, event: t.event, location: t.location,
      eventStart: t.eventStart, departDate: t.departDate, returnDate: t.returnDate,
      funding: t.funding, costCenter: t.costCenter,
      thirdParties: (t.thirdParties || []).map(function (x) {
        return { name: x.name, contact: x.contact, project: x.project, maxReimb: x.maxReimb };
      }),
    };

    var res = TravelForm.copyForNewTrip(saved, new Date().toISOString());
    applyModel(res.model);

    var where = byId("event") || byId("name");
    if (where && where.scrollIntoView) {
      try { where.scrollIntoView({ behavior: "smooth", block: "center" }); }
      catch (e) { where.scrollIntoView(); }
    }

    if (res.dropped.length) {
      res.dropped.forEach(function (k) {
        var el = byId(k);
        if (el) { el.style.outline = "2px solid var(--err-fg)"; el.style.outlineOffset = "1px"; }
      });
      setStatus("info", "Copied \u201c" + (t.event || "that trip") + "\u201d. " +
        res.dropped.map(function (k) { return COPY_LABELS[k] || k; }).join(", ") +
        " \u2014 cleared, because they were in the past. Fill them in, check the costs, " +
        "then Create travel request." + (partial ? " " + PARTIAL_NOTE : ""));
    } else {
      setStatus("info", "Copied \u201c" + (t.event || "that trip") + "\u201d into the form. " +
        "Check the dates and the costs before you submit." + (partial ? " " + PARTIAL_NOTE : ""));
    }
  }

  /** Write a saved model back onto the form. The inverse of model(). */
  function applyModel(m) {
    [
      "name", "costCenter", "division", "bureau", "otherStaff", "event", "location",
      "eventStart", "attendeeRole", "confDates", "departDate", "returnDate",
      "reason", "meetingLink", "comments", "funding",
    ].forEach(function (k) { setVal(k, m[k] || ""); });

    var modes = m.modes || {};
    setProp("modePersonal", "checked", !!modes.personal);
    setProp("modeState", "checked", !!modes.state);
    setProp("modeAir", "checked", !!modes.air);

    var c = m.costs || {};
    var COSTS = {
      cTravelMode: "travelMode", cLuggage: "luggage", cParking: "parking",
      cTaxi: "taxi", cLodgingNights: "lodgingNights", cLodgingRate: "lodgingRate",
      cRegistration: "registration", cAdditional: "additional",
      cAdditionalDesc: "additionalDesc", cMealsB: "mealsB", cMealsL: "mealsL",
      cMealsD: "mealsD",
    };
    Object.keys(COSTS).forEach(function (id) {
      // The three meal counts start life at 0 on a blank form, and slimming
      // drops a zero as "nothing was entered". Restoring them empty would
      // leave a copied request looking subtly unlike a fresh one, so they go
      // back to the same default the page ships with.
      var dflt = /^cMeals[BLD]$/.test(id) ? "0" : "";
      setVal(id, c[COSTS[id]] || dflt);
    });

    var tps = m.thirdParties || [];
    [1, 2].forEach(function (n) {
      var t = tps[n - 1] || {};
      setVal("tp" + n + "Name", t.name || "");
      setVal("tp" + n + "Contact", t.contact || "");
      setVal("tp" + n + "Project", t.project || "");
      setVal("tp" + n + "Max", t.maxReimb || "");
      setProp("tp" + n + "Packet", "checked", !!t.packet);
      // Spelled out rather than lowercased: the boxes are tp1Reg and tp1Air
      // but the stored keys are "registration" and "airfare", so deriving one
      // from the other silently loses those two ticks on every copy.
      var ITEMS = { Reg: "registration", Lodging: "lodging", Air: "airfare",
                    Meals: "meals", Ground: "ground" };
      Object.keys(ITEMS).forEach(function (k) {
        setProp("tp" + n + k, "checked", !!(t.items && t.items[ITEMS[k]]));
      });
      setVal("tp" + n + "Notes", t.notes || "");
      if (t.name) { setAttrIf("tp" + n, "open", "open"); }
    });

    try { refreshTotal(); refreshFyLine(); } catch (e) { /* cosmetic */ }
  }

  /**
   * Open a pre-filled appointment for a trip. Outlook saves it, not us.
   *
   * displayNewAppointmentForm needs no Graph scope at all - it hands Outlook a
   * form and steps back. Writing the event directly would need
   * Calendars.ReadWrite, which is a new consent screen and another pass over
   * the listing, to save the person one keystroke.
   *
   * If the trip has a matched hotel confirmation, its dates and the property
   * name are read out of that email first: a booking is the trip's real
   * shape, usually a day wider at each end than the conference itself.
   */
  /**
   * Ask for the dates rather than refusing over them.
   *
   * An older trip, or one copied from a request that never had them, has no
   * departure date - and "add a departure date and try again" is a dead end
   * dressed up as an error: it names the problem, offers no way to fix it,
   * and sends the person off to find a form. Asking here takes four seconds,
   * and what they type is kept on the trip so it is never asked twice.
   *
   * Resolves to null if they close it. A cancelled dialog is an answer.
   */
  function askForTripDates(trip) {
    return new Promise(function (resolve) {
      var host = byId("tripsList");
      if (!host) { resolve(null); return; }
      var old = byId("askDates");
      if (old) { old.remove(); }

      var box = document.createElement("div");
      box.id = "askDates";
      box.className = "warnbox";
      box.style.marginTop = "8px";

      var p = document.createElement("p");
      p.style.margin = "0 0 6px";
      p.innerHTML = "<b>When is this trip?</b> \u201c" +
        String(trip.event || "That trip").replace(/[<>&]/g, "") +
        "\u201d has no dates recorded, so there is nothing to put in the calendar yet.";

      var row = document.createElement("div");
      row.className = "row2";
      function field(label, id, value) {
        var l = document.createElement("label");
        l.textContent = label;
        var i = document.createElement("input");
        i.type = "date"; i.id = id;
        if (value) { i.value = value; }
        l.appendChild(i);
        row.appendChild(l);
        return i;
      }
      var dep = field("Departure", "askDepart", trip.eventStart || "");
      var ret = field("Return (optional)", "askReturn", "");

      var actions = document.createElement("p");
      actions.style.margin = "6px 0 0";
      var ok = document.createElement("button");
      ok.className = "primary";
      ok.textContent = "Add to calendar";
      var no = document.createElement("button");
      no.textContent = "Cancel";
      no.style.marginLeft = "6px";

      function done(value) { box.remove(); resolve(value); }
      no.addEventListener("click", function () { done(null); });
      ok.addEventListener("click", function () {
        var d = dep.value;
        if (!d) {
          setStatus("error", "A departure date is the one thing needed \u2014 the rest can follow.");
          dep.focus();
          return;
        }
        done({ departDate: d, returnDate: ret.value || d });
      });

      actions.appendChild(ok); actions.appendChild(no);
      box.appendChild(p); box.appendChild(row); box.appendChild(actions);
      host.appendChild(box);
      try { box.scrollIntoView({ behavior: "smooth", block: "center" }); } catch (e) { /* fine */ }
      dep.focus();
    });
  }

  /** Keep what the person just told us, so it is asked once and not again. */
  function rememberTripDates(trip, dates) {
    var list = trips();
    var hit = null;
    list.forEach(function (t) { if (t && t.id === trip.id) { hit = t; } });
    if (!hit) { return; }
    hit.departDate = dates.departDate;
    if (!hit.returnDate) { hit.returnDate = dates.returnDate; }
    if (!hit.eventStart) { hit.eventStart = dates.departDate; }
    saveTrips(list);
    renderTrips();
  }

  async function addTripToCalendar(trip) {
    var mailbox = Office.context.mailbox;
    if (!mailbox || !mailbox.displayNewAppointmentForm) {
      setStatus("error", "This version of Outlook can't open a new appointment from an add-in. " +
        "The trip dates are on the request if you want to add it by hand.");
      return;
    }
    setStatus("work", "Building the appointment\u2026");
    var hotel = null;
    try {
      // Only if a booking email was actually matched to this trip.
      if (trip.bookingId) {
        var token = await GraphData.getToken();
        var msg = await GraphData.messageFull(token, trip.bookingId);
        var text = /<[a-z][\s\S]*>/i.test(msg.body || "")
          ? TdMail.htmlToText(msg.body) : (msg.body || "");
        hotel = TdMail.hotelDetails(text);
      }
    } catch (e) { hotel = null; }   // the trip's own dates are enough

    var entry = TravelForm.calendarEntry(trip, hotel);
    if (!entry) {
      setStatus("info", "");
      var asked = await askForTripDates(trip);
      if (!asked) { setStatus("info", "No appointment created."); return; }
      trip = Object.assign({}, trip, asked);
      rememberTripDates(trip, asked);
      entry = TravelForm.calendarEntry(trip, hotel);
      if (!entry) { setStatus("error", "Those dates didn't work \u2014 try again."); return; }
    }
    try {
      mailbox.displayNewAppointmentForm({
        requiredAttendees: [],
        subject: entry.subject,
        location: entry.location,
        // A TIMED appointment, not an all-day one - AppointmentForm has no
        // all-day flag, so the end is literal rather than exclusive. Shifting
        // it by a day, as an all-day event would need, made every trip run a
        // day past the return. Midnight to 23:59 covers the whole span and
        // ends on the right date.
        start: new Date(entry.start + "T00:00:00"),
        end: new Date(entry.end + "T23:59:00"),
        body: entry.body,
      });
      setStatus("info", "Appointment opened" +
        (hotel && hotel.name ? " with " + hotel.name + " and your booking dates" : "") +
        ". Check it and save it \u2014 Travel Desk doesn't write to your calendar.");
    } catch (e) {
      setStatus("error", "Couldn't open the appointment: " + ((e && e.message) || e));
    }
  }

  async function checkBookings() {
    byId("checkBookings").disabled = true;
    try {
      var list = trips();
      var open = list.filter(function (t) { return t.status !== "booked"; });
      if (!open.length) { setStatus("info", "No trips waiting on bookings. \u2708\ufe0f"); return; }
      setStatus("work", "Checking your inbox for booking confirmations\u2026");
      var token = await GraphData.getToken();
      var earliest = open.reduce(function (min, t) { return t.createdAt < min ? t.createdAt : min; }, open[0].createdAt);
      // Reach back before the earliest request, or the confirmation that
      // arrived BEFORE the paperwork is never even fetched.
      var lookback = (TravelForm.BOOKING_LOOKBACK_DAYS || 120) * 864e5;
      var since = new Date((Date.parse(earliest) || Date.now()) - lookback).toISOString();
      var senders = (val("bookingSenders") || "").split(",");
      var emails = await GraphData.bookingEmails(token, since);
      var booked = 0, unsure = 0;
      open.forEach(function (t) {
        var m = TravelForm.matchBooking(t, emails, senders);
        if (m.confident) {
          t.status = "booked";
          markBookedEverywhere(t);
          t.bookingLink = m.confident.webLink || "";
          t.bookingSubject = m.confident.subject || "";
          t.bookingId = m.confident.id || "";
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
      var res = await ensurePlannerTable(token, ref);
      if (!res.name) {
        wbRef = ref;
        byId("wbUrl").value = ref.webUrl || "";
        await offerTableConversion(token, ref);
        setStatus("error", "\u201c" + ref.name + "\u201d has no Excel table" +
          (res.reason === "empty"
            ? ", and no sheet in it has a header row yet. Add your column headings first."
            : ", and more than one sheet has data \u2014 pick the planner\u2019s sheet below.") +
          " Then click \u201cMake this sheet a table\u201d.");
        return;
      }
      wbRef = ref;
      byId("wbUrl").value = ref.webUrl || "";
      var box2 = byId("noTableBox");
      if (box2) { box2.hidden = true; }
      var sel = byId("tableName");
      sel.innerHTML = "";
      (res.tables || [res.name]).forEach(function (t) {
        var opt = document.createElement("option");
        opt.value = opt.textContent = t;
        sel.appendChild(opt);
      });
      saveSettings({ wbUrl: ref.webUrl || "", tableName: res.name, wbRef: ref });
      setStatus("info", "Connected: " + ref.name + (res.created
        ? " \u2014 it had no table, so \u201c" + res.sheet + "\u201d (" + res.address +
          ") is now the table \u201c" + res.name + "\u201d, with every existing row intact."
        : " \u2014 table \u201c" + res.name + "\u201d") +
        ". Set its fiscal year above and click Save planner.");
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
  /**
   * A planner workbook is useless without a table, so make one.
   *
   * Travel Desk writes rows through an Excel table. When "Create my planner"
   * failed at the tables/add step the file was left with headers and no
   * table, and nothing ever retried: every later submit failed at the row,
   * the coordinator view read zero, and the only clue was a collapsed box in
   * Setup asking the person to convert a sheet they did not know was
   * unconverted. Connecting now repairs that instead of reporting it.
   *
   * It only acts when the answer is unambiguous - one sheet with data, or a
   * sheet actually called Planner. A workbook with several populated sheets
   * is somebody else's spreadsheet, and guessing which one to restructure
   * would be worse than asking.
   */
  async function ensurePlannerTable(token, ref) {
    var tables = await GraphData.listTables(token, ref);
    if (tables.length) { return { name: tables[0], created: false, tables: tables }; }

    var sheets = await GraphData.listWorksheets(token, ref);
    var withData = [];
    for (var i = 0; i < sheets.length; i++) {
      try {
        var used = await GraphData.usedRange(token, ref, sheets[i]);
        if (used && used.address && (used.columnCount || 0) > 1) {
          withData.push({ sheet: sheets[i], address: used.address, rows: used.rowCount || 0 });
        }
      } catch (e) { /* an empty sheet has no used range; that is not an error */ }
    }
    if (!withData.length) { return { name: "", created: false, reason: "empty", sheets: sheets }; }

    var named = withData.filter(function (w) { return /^planner$/i.test(w.sheet); });
    var pick = named.length === 1 ? named[0] : (withData.length === 1 ? withData[0] : null);
    if (!pick) { return { name: "", created: false, reason: "ambiguous", sheets: sheets }; }

    var t = await GraphData.addTable(token, ref, pick.address, "TravelPlanner");
    return {
      name: (t && t.name) || "TravelPlanner", created: true,
      sheet: pick.sheet, address: pick.address,
    };
  }

  /**
   * Tell someone at SETUP time whether their rows will actually land.
   *
   * Never blocks and never nags: it runs after a planner is configured, and
   * only says anything when the answer is bad. A traveler who cannot write
   * has one thing to do - ask their coordinator - and no way to discover it
   * on their own, because everything else about the add-in works fine for
   * them right up until the row silently does not appear.
   */
  /**
   * Say once, to the coordinator, where the planner actually lives.
   *
   * A personal-OneDrive planner and a site planner look identical in here -
   * both are sharepoint.com URLs, both connect the same way, both work
   * perfectly for the person who created them. The difference only shows up
   * as other people failing to write, and as the file vanishing when its
   * owner changes role. Said once at setup it costs a minute to fix; found
   * later it costs the division its travel record.
   */
  function noteWhereThePlannerLives(url) {
    if (!url || !TravelCoord.isPersonalDrive(url)) { return; }
    if (settings().plannerHomeNoted) { return; }
    saveSettings({ plannerHomeNoted: true });
    var box = byId("plannerList");
    if (!box) { return; }
    var p = document.createElement("p");
    p.className = "warnbox";
    p.style.marginTop = "8px";
    p.innerHTML =
      "<b>This planner is in your personal OneDrive.</b> It works, and you can stop here \u2014 " +
      "but two things follow from it. Every traveler needs you to share the file with them " +
      "individually before their rows will land, and the file goes with your account if you " +
      "change roles. A workbook on a Teams or SharePoint site takes its permissions from the " +
      "site\u2019s membership instead, so anyone already on the site can file. Travel Desk " +
      "connects to either one identically \u2014 <b>Show my files</b> lists both.";
    box.appendChild(p);
  }

  async function verifyPlannerAccess(ref, who) {
    if (!ref || !ref.driveId || !ref.itemId) { return; }
    try {
      var token = await GraphData.getToken();
      var res = await GraphData.canWriteWorkbook(token, ref);
      if (res.ok) {
        if (!res.unverified) { saveSettings({ plannerWriteOk: true }); }
        return;
      }
      saveSettings({ plannerWriteOk: false });
      if (res.reason === "missing") {
        setStatus("error", "The planner workbook isn't there any more \u2014 it has been moved, " +
          "renamed or deleted. " + (who === "coordinator"
            ? "Connect the current one in Setup."
            : "Ask your travel coordinator for a new setup code."));
        return;
      }
      setStatus("error", who === "coordinator"
        ? "You can open the planner but not write to it. Travel Desk adds each trip as a row, " +
          "so requests will fail at the row step. Give yourself edit access to the workbook."
        : "You can see the planner but not write to it, so your trips won't reach the shared " +
          "spreadsheet. Ask your travel coordinator to give you edit access \u2014 everything " +
          "else here works in the meantime, and your requests are kept locally.");
    } catch (e) { /* a probe that fails proves nothing; say nothing */ }
  }

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
    // Where it goes is the coordinator's call. Blank keeps the old behaviour,
    // the OneDrive root, which is a reasonable default and a poor assumption
    // to have baked in permanently.
    var folder = (val("newPlannerFolder") || "").trim().replace(/^\/+|\/+$/g, "");
    byId("createPlanner").disabled = true;
    try {
      setStatus("work", "Building the workbook\u2026");
      var built = XlsxGen.buildWorkbook(TravelForm.DEFAULT_PLANNER_HEADERS, "Planner");
      var zip = new JSZip();
      Object.keys(built.parts).forEach(function (path) { zip.file(path, built.parts[path]); });
      var bytes = await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });

      setStatus("work", "Saving it to your OneDrive\u2026");
      var token = await GraphData.getToken();
      if (folder) { setStatus("work", "Making sure \u201c" + folder + "\u201d exists\u2026"); }
      var made = await GraphData.uploadWorkbook(token, name, bytes, folder);
      saveSettings({ plannerFolder: folder });
      wbRef = { driveId: made.ref.driveId, itemId: made.ref.itemId, name: made.name };

      setStatus("work", "Formatting it as a table\u2026");
      // The workbook is already saved by this point. If the table step fails,
      // the file must not be abandoned - that is what made this look like
      // "create my planner doesn't work" while leaving an orphaned workbook in
      // OneDrive. Try Graph's own usedRange address as a fallback, which is
      // the path "Make this sheet a table" uses and is known to work.
      var tableName = null;
      try {
        var t = await GraphData.addTable(token, wbRef, built.range, "TravelPlanner");
        tableName = (t && t.name) || "TravelPlanner";
      } catch (tableErr) {
        try {
          var used = await GraphData.usedRange(token, wbRef, "Planner");
          if (used && used.address) {
            var t2 = await GraphData.addTable(token, wbRef, used.address, "TravelPlanner");
            tableName = (t2 && t2.name) || "TravelPlanner";
          }
        } catch (e2) {
          // Last resort before giving up: the same repair a Connect does.
          try {
            var res3 = await ensurePlannerTable(token, wbRef);
            if (res3.name) { tableName = res3.name; }
          } catch (e3) { /* fall through to the guidance below */ }
        }
        if (!tableName) {
          // Keep the workbook connected so one click finishes the job.
          setVal("wbUrl", made.webUrl || "");
          saveSettings({ wbUrl: made.webUrl || "" });
          setStatus("error", "Created \u201c" + made.name + "\u201d in " +
            (plannerFolder(made.webUrl) || "your OneDrive") + ", but couldn't " +
            "format it as a table (" + ((tableErr && tableErr.message) || tableErr) + "). " +
            "It's connected \u2014 pick the Planner sheet and click \u201cMake this sheet a table\u201d.");
          setProp("noTableBox", "hidden", false);
          return;
        }
      }

      var sel = byId("tableName");
      if (sel) {
        sel.innerHTML = "";
        var o = document.createElement("option");
        o.value = o.textContent = tableName;
        sel.appendChild(o);
      }
      setVal("wbUrl", made.webUrl || "");
      setProp("noTableBox", "hidden", true);
      saveSettings({ wbUrl: made.webUrl || "", tableName: tableName });
      renderPlannerLink();
      var shown = openWorkbookUrl(made.webUrl);
      // "in your OneDrive" is not a location. Name the folder it went into -
      // especially when one was typed, since that is the moment a file
      // becomes hard to find again.
      var at = plannerFolder(made.webUrl) || "your OneDrive";
      setStatus("info", "Created \u201c" + made.name + "\u201d in " + at + " and connected it." +
        (shown ? " I've opened it so you can see it." :
                 " Use \u201cOpen the planner workbook\u201d above to open it.") +
        " Set the fiscal year it covers, then Save planner for this year.");
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
      var res = await ensurePlannerTable(token, wbRef);
      if (!res.name) {
        await offerTableConversion(token, wbRef);
        setStatus("error", "\u201c" + wbRef.name + "\u201d has no Excel table" +
          (res.reason === "empty"
            ? ", and no sheet in it has a header row yet. Add your column headings first."
            : ", and it has more than one sheet with data \u2014 pick the planner\u2019s sheet below.") +
          " Then click \u201cMake this sheet a table\u201d.");
        return;
      }
      var box = byId("noTableBox");
      if (box) { box.hidden = true; }
      var sel = byId("tableName");
      sel.innerHTML = "";
      (res.tables || [res.name]).forEach(function (t) {
        var opt = document.createElement("option");
        opt.value = opt.textContent = t;
        sel.appendChild(opt);
      });
      saveSettings({ wbUrl: url, tableName: res.name, wbRef: wbRef });
      setStatus("info", res.created
        ? "Connected: " + wbRef.name + " \u2014 it had no table, so \u201c" + res.sheet +
          "\u201d (" + res.address + ") is now the table \u201c" + res.name +
          "\u201d. Every row that was there is still there. You're set."
        : "Connected: " + wbRef.name + " \u2014 table \u201c" + res.name + "\u201d. You're set.");
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
    var doDraft = isChecked("doDraft");
    var doRow = isChecked("doRow");
    if (!doDraft && !doRow) { setStatus("error", "Pick at least one action."); return; }
    if (!m.name || !m.event || !m.location) {
      setStatus("error", "Name, event, and location are required.");
      return;
    }
    var s = settings();
    // The email and the planner row are two actions the user ticked
    // independently. A missing planner used to abort BOTH with an early
    // return, so a request could produce no email at all while the traveler
    // believed one had been sent. Now whatever can be done is done, and
    // whatever cannot is named.
    var plannerBlocked = doRow && !(s.wbRef || wbRef);

    setProp("submit", "disabled", true);
    var done = [];
    var failed = [];
    try {
      var token = await GraphData.getToken();
      var orgOpts = {
        orgName: s.orgName || "",
        fundingLabel: s.fundingLabel || "",
        fyStartMonth: Number(s.fyStartMonth) || 1,
        fyPrefix: s.fyPrefix || "FY",
        costMode: val("costMode") || "per-person",
      };

      if (doDraft) {
        try {
          setStatus("work", "Creating the Travel Authorization draft…");
          var draft = await GraphData.createDraft(token, val("coordEmail").trim(),
            TravelForm.subjectLine(m), TravelForm.formHtml(m, orgOpts));
          done.push("draft OPENED for you — review it and press Send");
          // "Draft created" is not "request sent". Open it, because this is
          // the moment the traveler has to press Send.
          try {
            if (draft && draft.webLink) {
              if (Office.context.ui && Office.context.ui.openBrowserWindow) {
                Office.context.ui.openBrowserWindow(draft.webLink);
              } else {
                window.open(draft.webLink, "_blank");
              }
            }
          } catch (openErr) { done.push("(it's in your Drafts folder)"); }
        } catch (draftErr) {
          failed.push("the email draft (" + ((draftErr && draftErr.message) || draftErr) + ")");
        }
      }

      if (plannerBlocked) {
        openPlannerSetup();
        failed.push("the planner row — no planner is connected. I've opened Setup below: use " +
          "“Create my planner”, or “Show my files” to connect an existing workbook");
      } else if (doRow) {
        try {
          setStatus("work", "Adding the planner row…");
          var tripFy = TravelForm.fiscalLabel(m.eventStart || m.departDate, s.fyStartMonth, s.fyPrefix);
          var picked = TravelForm.pickPlanner(s.planners, tripFy);
          var ref, tableName;
          if (picked) {
            ref = picked.planner.wbRef;
            tableName = picked.planner.tableName;
          } else if (s.planners && Object.keys(s.planners).length) {
            throw new Error("no planner saved for " + (tripFy || "that date") +
              " — in Setup, connect that year's workbook and save it for " + tripFy +
              " (or save one planner as 'all years')");
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
        } catch (rowErr) {
          failed.push("the planner row (" + ((rowErr && rowErr.message) || rowErr) + ")");
        }
      }

      saveSettings({ name: m.name, costCenter: m.costCenter, division: m.division, bureau: m.bureau });
      // Record the trip if ANY part succeeded. This previously ran only on a
      // completely clean pass, so a created draft could leave no trace in
      // My trips - the traveler had a request the add-in had no memory of.
      if (done.length) { addTrip(m); }

      if (done.length && !failed.length) {
        setStatus("ok", "Done: " + done.join(" + ") + ".");
      } else if (done.length) {
        setStatus("info", "Done: " + done.join(" + ") + ". NOT done: " + failed.join("; ") + ".");
      } else {
        setStatus("error", "Nothing was created. Couldn't do " + failed.join("; ") + ".");
      }
    } catch (e) {
      setStatus("error", "Travel request failed: " + ((e && e.message) || e) +
        (done.length ? " — but this did happen: " + done.join(" + ") + "." : ""));
    } finally {
      setProp("submit", "disabled", false);
    }
  }
})();
