/*
 * Travel Desk — Microsoft Graph data layer.
 *
 * AUTH: Nested App Authentication (NAA) via MSAL — no backend; same pattern
 * as the rest of the PurposeBuilt suite. Delegated scopes:
 *   Files.ReadWrite.All — append rows to the shared Division travel planner
 *                         workbook on SharePoint/OneDrive (files the signed-in
 *                         user can already reach; no elevation)
 *   Mail.ReadWrite      — create the Travel Authorization email DRAFT in the
 *                         user's own mailbox (nothing is ever sent)
 *
 * Exposes a global `GraphData` object.
 */
/* global msal */
(function (root) {
  "use strict";

  var CLIENT_ID = "b54eb66f-221f-4ed0-bcdc-857b4ae6111d"; // "Travel Desk" Entra app (purposebuilt.systems tenant)
  var GRAPH = "https://graph.microsoft.com/v1.0";
  var SCOPES = ["Files.ReadWrite.All", "Mail.ReadWrite"];

  var pcaPromise = null;

  function getPca() {
    if (!pcaPromise) {
      pcaPromise = msal.createNestablePublicClientApplication({
        auth: {
          clientId: CLIENT_ID,
          authority: "https://login.microsoftonline.com/common",
        },
      });
    }
    return pcaPromise;
  }

  // --- add-in sign-out state -------------------------------------------
  //
  // Certification rejected the naive version on a sibling add-in: "after
  // clicking sign-out there is no response or not signed out." The reason is
  // structural. Under nested app authentication Outlook owns the session and
  // getAllAccounts() reports the HOST's account, not a cache this add-in
  // controls - so clearing MSAL's cache changes nothing visible, the next
  // silent acquisition succeeds anyway, and the pane redraws as signed in.
  //
  // A sign-out this add-in cannot deliver should not be offered. What it CAN
  // deliver is refusing to act until the user authenticates again: while
  // signed out it reports itself signed out and will not use a silent token,
  // so the next action raises a real prompt. Outlook's own session is
  // untouched, and the pane says so.
  var SIGNED_OUT_KEY = "addinSignedOut";
  var signedOut = false;
  try { signedOut = Office.context.roamingSettings.get(SIGNED_OUT_KEY) === true; } catch (e) { signedOut = false; }

  function setSignedOut(v) {
    signedOut = !!v;
    try {
      Office.context.roamingSettings.set(SIGNED_OUT_KEY, signedOut);
      Office.context.roamingSettings.saveAsync(function () {});
    } catch (e) { /* in-memory is still correct for this session */ }
  }

  /** Signed-in account, or null. Reports null while signed out, by design. */

  /**
   * Microsoft Graph throttles, and this add-in makes bursts of calls - a
   * records bundle or a bulk post is dozens to hundreds. An unretried 429
   * aborts the whole run part-way, which is the worst possible failure for
   * work that is half-written. One respectful retry honouring Retry-After
   * absorbs the overwhelming majority of throttling without hammering the
   * service; anything past that is a real outage and should surface.
   */
  async function fetchRetry(url, opts) {
    var res = await fetch(url, opts);
    if (res.status === 429 || res.status === 503) {
      var wait = Number(res.headers.get("Retry-After") || 3) * 1000;
      await new Promise(function (r) { setTimeout(r, Math.min(wait, 15000)); });
      res = await fetch(url, opts);
    }
    return res;
  }

  async function currentAccount() {
    if (signedOut) { return null; }
    try {
      var pca = await getPca();
      var accts = (pca.getAllAccounts && pca.getAllAccounts()) || [];
      return accts.length ? (accts[0].username || accts[0].name || "signed in") : null;
    } catch (e) { return null; }
  }

  /**
   * Sign out of the add-in. The state flips SYNCHRONOUSLY before any awaiting
   * so the pane can respond instantly - awaiting a broker handshake first is
   * the "no response" half of the finding. Cache clearing is best-effort on
   * top; the enforced state is what makes this real.
   */
  function signOut() {
    setSignedOut(true);
    var pending = pcaPromise;      // only clear what exists; never start a handshake here
    pcaPromise = null;
    if (!pending) { return Promise.resolve(true); }
    return Promise.resolve(pending).then(function (pca) {
      var accts = (pca && pca.getAllAccounts && pca.getAllAccounts()) || [];
      var chain = Promise.resolve();
      accts.forEach(function (a) {
        chain = chain.then(function () {
          if (pca.clearCache) { return pca.clearCache({ account: a }); }
          if (pca.logoutPopup) { return pca.logoutPopup({ account: a }); }
        }).catch(function () { /* best effort; the enforced state stands */ });
      });
      return chain.then(function () { return true; });
    }).catch(function () { return true; });
  }

  /** True while the user has signed the add-in out. */
  function isSignedOut() { return signedOut; }




  /**
   * Sign-in must never hang the pane. An un-timed await on the popup flow
   * leaves a button disabled and nothing visible happening — which reads
   * to the user as "the button does nothing".
   */
  function withTimeout(promise, ms, message) {
    var timer;
    return Promise.race([
      promise.then(function (v) { clearTimeout(timer); return v; },
                   function (e) { clearTimeout(timer); throw e; }),
      new Promise(function (_, reject) {
        timer = setTimeout(function () { reject(new Error(message)); }, ms);
      }),
    ]);
  }

  async function getToken() {
    var pca = await withTimeout(getPca(), 20000,
      "Sign-in didn't start. Fully quit Outlook (Cmd+Q) and reopen, then try again.");
    try {
      // Signed out means signed out: skip silent so the user must re-authenticate.
      if (signedOut) { throw new Error("signed out of the add-in"); }
      return (await withTimeout(pca.acquireTokenSilent({ scopes: SCOPES }), 20000, "silent timeout")).accessToken;
    } catch (e) {
      var interactive = await withTimeout(
        pca.acquireTokenPopup({ scopes: SCOPES }), 120000,
        "Sign-in didn't finish. A Microsoft sign-in window may have opened behind Outlook — " +
        "check for it (or Mission Control), finish signing in, and click again. If no window " +
        "appeared at all, fully quit Outlook (Cmd+Q), reopen, and retry.");
      setSignedOut(false);   // a real interactive sign-in ends the signed-out state
      return interactive.accessToken;
    }
  }

  async function graphJson(token, method, path, body) {
    var res = await fetchRetry(GRAPH + path, {
      method: method,
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: body ? JSON.stringify(body) : undefined,
    });
    if (!res.ok) { throw new Error("Graph " + method + " " + path.split("?")[0] + " -> " + res.status + " " + (await res.text())); }
    return res.status === 204 ? null : res.json();
  }

  // ---------- workbook (Division travel planner on SharePoint/OneDrive) ----------

  /** Base64url share-token for the Graph /shares API. */
  function shareToken(url) {
    var b64 = btoa(unescape(encodeURIComponent(String(url).trim())));
    return "u!" + b64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  }

  /**
   * Resolve a pasted SharePoint/OneDrive link to {driveId, itemId, name}.
   * Works with the "Copy link" URL of the workbook — no site browsing needed.
   */
  async function resolveWorkbook(token, shareUrl) {
    var item = await graphJson(token, "GET",
      "/shares/" + shareToken(shareUrl) + "/driveItem?$select=id,name,parentReference");
    return {
      driveId: item.parentReference && item.parentReference.driveId,
      itemId: item.id,
      name: item.name,
    };
  }

  function wbBase(ref) {
    return "/drives/" + ref.driveId + "/items/" + ref.itemId + "/workbook";
  }

  /** Table names in the workbook (the planner should be one Excel Table). */
  async function listTables(token, ref) {
    var res = await graphJson(token, "GET", wbBase(ref) + "/tables?$select=name");
    return (res.value || []).map(function (t) { return t.name; });
  }

  /** Header row of a table, as an array of strings. */
  async function tableHeaders(token, ref, tableName) {
    var res = await graphJson(token, "GET",
      wbBase(ref) + "/tables/" + encodeURIComponent(tableName) + "/headerRowRange?$select=values");
    return ((res.values || [])[0] || []).map(function (h) { return String(h == null ? "" : h); });
  }

  /** Worksheet names in a connected workbook. */
  async function listWorksheets(token, ref) {
    var res = await graphJson(token, "GET", wbBase(ref) + "/worksheets?$select=name");
    return (res.value || []).map(function (w) { return w.name; });
  }

  /** Used range of a sheet — tells us where the headers actually are. */
  async function usedRange(token, ref, sheetName) {
    return graphJson(token, "GET", wbBase(ref) + "/worksheets/" +
      encodeURIComponent(sheetName) + "/usedRange?$select=address,rowCount,columnCount,values");
  }

  /**
   * Turn a plain sheet range into a real Excel Table — the thing Travel
   * Desk appends rows through. Address is Excel A1 style, e.g.
   * "Planner!A1:O40". Optionally renames it so the picker is readable.
   */
  async function addTable(token, ref, address, name) {
    var t = await graphJson(token, "POST", wbBase(ref) + "/tables/add",
      { address: address, hasHeaders: true });
    if (name && t && t.id) {
      try {
        var renamed = await graphJson(token, "PATCH",
          wbBase(ref) + "/tables/" + encodeURIComponent(t.id), { name: name });
        return renamed || t;
      } catch (e) { return t; } // a name clash is cosmetic; the table exists
    }
    return t;
  }

  /** Create a workbook in the user's OneDrive root. Returns a {driveId,itemId} ref. */
  /**
   * Create a folder chain in OneDrive if it isn't there. 409 means it already
   * exists, which is the outcome we want, so it is not an error.
   */
  var ensuredFolders = {};
  async function ensureFolder(token, path) {
    var segs = String(path || "").split("/").map(function (x) { return x.trim(); }).filter(Boolean);
    var soFar = "";
    for (var i = 0; i < segs.length; i++) {
      var parent = soFar;
      soFar = soFar ? soFar + "/" + segs[i] : segs[i];
      if (ensuredFolders[soFar]) { continue; }
      var url = parent
        ? "/me/drive/root:/" + parent.split("/").map(encodeURIComponent).join("/") + ":/children"
        : "/me/drive/root/children";
      var res = await fetchRetry(GRAPH + url, {
        method: "POST",
        headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
        body: JSON.stringify({ name: segs[i], folder: {}, "@microsoft.graph.conflictBehavior": "fail" }),
      });
      if (!res.ok && res.status !== 409) {
        throw new Error("Couldn't create the folder \u201c" + soFar + "\u201d (" + res.status + ")");
      }
      ensuredFolders[soFar] = true;
    }
    return segs.join("/");
  }

  /**
   * folder is optional. Empty means the OneDrive root, which is where this
   * always used to put the planner - fine for a personal file, less so for
   * something a division shares, which is why the caller can now choose.
   */
  async function uploadWorkbook(token, filename, bytes, folder) {
    var dir = "";
    if (folder && String(folder).trim()) { dir = await ensureFolder(token, folder); }
    var target = dir
      ? dir.split("/").map(encodeURIComponent).join("/") + "/" + encodeURIComponent(filename)
      : encodeURIComponent(filename);
    var res = await fetchRetry(GRAPH + "/me/drive/root:/" + target + ":/content", {
      method: "PUT",
      headers: {
        Authorization: "Bearer " + token,
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
      body: bytes,
    });
    if (!res.ok) {
      var txt = await res.text();
      if (res.status === 404) {
        throw new Error("Your OneDrive isn't set up yet — open onedrive.com once, then try again.");
      }
      throw new Error("Couldn't create the workbook (" + res.status + ") " + txt.slice(0, 200));
    }
    var item = await res.json();
    return {
      ref: { driveId: (item.parentReference || {}).driveId, itemId: item.id },
      name: item.name, webUrl: item.webUrl,
    };
  }

  /** Every data row of the planner table — the coordinator's whole picture. */
  async function tableRows(token, ref, tableName) {
    var res = await graphJson(token, "GET",
      wbBase(ref) + "/tables/" + encodeURIComponent(tableName) + "/dataBodyRange?$select=values");
    return (res && res.values) || [];
  }

  /** Travel Authorization emails the coordinator has received. */
  async function authEmails(token, daysBack) {
    var since = new Date(Date.now() - (daysBack || 365) * 864e5).toISOString();
    var res = await graphJson(token, "GET",
      "/me/messages?$select=subject,receivedDateTime,from,webLink" +
      "&$filter=receivedDateTime ge " + since +
      " and startswith(subject,'Travel Auth')&$top=200");
    return res.value || [];
  }

  /** Replace one existing row (0-based index within the data body). */
  async function updateTableRow(token, ref, tableName, index, values) {
    return graphJson(token, "PATCH",
      wbBase(ref) + "/tables/" + encodeURIComponent(tableName) +
      "/rows/itemAt(index=" + Number(index) + ")", { values: [values] });
  }

  /** Column names on a table (to see whether the lifecycle columns exist). */
  async function tableColumns(token, ref, tableName) {
    var res = await graphJson(token, "GET",
      wbBase(ref) + "/tables/" + encodeURIComponent(tableName) + "/columns?$select=name");
    return (res.value || []).map(function (c) { return c.name; });
  }

  /** Add a column to an existing planner table (for planners built before
   *  close-out existed). Appends at the end so nothing shifts. */
  async function addTableColumn(token, ref, tableName, name) {
    return graphJson(token, "POST",
      wbBase(ref) + "/tables/" + encodeURIComponent(tableName) + "/columns", { name: name });
  }

  /** Append one row to the planner table. */
  async function addTableRow(token, ref, tableName, rowValues) {
    return graphJson(token, "POST",
      wbBase(ref) + "/tables/" + encodeURIComponent(tableName) + "/rows/add",
      { values: [rowValues] });
  }

  /** Normalize a driveItem (handles remoteItem for shared/SharePoint files). */
  function itemRef(it) {
    var r = it.remoteItem || it;
    var pr = r.parentReference || {};
    if (!r.id || !pr.driveId) { return null; }
    return { driveId: pr.driveId, itemId: r.id, name: r.name || it.name || "workbook",
             webUrl: r.webUrl || it.webUrl || "" };
  }

  function xlsxOnly(items) {
    var out = [];
    (items || []).forEach(function (it) {
      var name = (it.remoteItem && it.remoteItem.name) || it.name || "";
      if (!/\.xlsx$/i.test(name)) { return; }
      var ref = itemRef(it);
      if (ref) { out.push(ref); }
    });
    return out;
  }

  /** Recently used workbooks — spans OneDrive AND SharePoint files you've opened. */
  async function recentWorkbooks(token) {
    var res = await graphJson(token, "GET", "/me/drive/recent?$top=50");
    return xlsxOnly(res.value);
  }

  /** Workbooks shared with the signed-in user. */
  async function sharedWorkbooks(token) {
    var res = await graphJson(token, "GET", "/me/drive/sharedWithMe");
    return xlsxOnly(res.value);
  }

  /** Search the user's OneDrive by name. */
  async function searchWorkbooks(token, q) {
    var res = await graphJson(token, "GET",
      "/me/drive/root/search(q='" + encodeURIComponent(String(q).replace(/'/g, "")) + "')?$top=25");
    return xlsxOnly(res.value);
  }

  /** Inbox messages from booking senders (e.g. Concur) since a date. */
  /**
   * Inbox mail since the request was raised.
   *
   * This used to discard anything not from a listed sender domain, defaulting
   * to Concur. That only ever finds the confirmations you predicted: a
   * conference registration, a hotel, an airline, or a colleague forwarding
   * you the booking all arrive from somewhere else and were dropped before
   * matching ever ran. A real case - an EDC-8 registration confirmation
   * forwarded from a work address - was invisible for exactly that reason.
   *
   * Sender is now a signal, not a gate. Everything in the window is returned
   * and matchBooking() decides, using the trip's own city and event name.
   */
  async function bookingEmails(token, sinceIso) {
    var filt = "receivedDateTime ge " + sinceIso;
    var res = await graphJson(token, "GET", "/me/mailFolders/inbox/messages?$filter=" +
      encodeURIComponent(filt) +
      "&$select=id,subject,bodyPreview,from,receivedDateTime,webLink&$orderby=receivedDateTime desc&$top=200");
    return (res.value || []).map(function (m) {
      return {
        id: m.id, subject: m.subject, bodyPreview: m.bodyPreview,
        receivedDateTime: m.receivedDateTime, webLink: m.webLink,
        from: (((m.from || {}).emailAddress || {}).address || ""),
      };
    });
  }

  // ---------- mail ----------

  /** Create the Travel Authorization email as a DRAFT (never sent). */
  /**
   * Setup invitations sitting in this user's mailbox. Searched by subject
   * prefix so a traveller never has to find, open, or copy anything.
   */
  async function setupInvites(token, subjectPrefix) {
    var res = await graphJson(token, "GET",
      "/me/messages?$select=id,subject,from,receivedDateTime" +
      "&$filter=" + encodeURIComponent("startswith(subject,'" +
        String(subjectPrefix || "Travel Desk setup").replace(/'/g, "''") + "')") +
      // NO $orderby: Graph rejects it alongside a startswith() filter with
      // "InefficientFilter". pickInvite() sorts newest-first client-side.
      "&$top=50");
    return res.value || [];
  }

  /** Full body of one message (search results carry only metadata). */
  async function messageBody(token, id) {
    var m = await graphJson(token, "GET", "/me/messages/" + encodeURIComponent(id) + "?$select=body");
    return (m && m.body && m.body.content) || "";
  }

  async function createDraft(token, to, subject, html) {
    var body = {
      subject: subject,
      body: { contentType: "HTML", content: html },
    };
    // An empty address is rejected outright (400 ErrorInvalidRecipients).
    // A third party with no billing contact on file, or a blank coordinator
    // address, should still get a draft — the user just fills the To line.
    var list = (Array.isArray(to) ? to : [to])
      .map(function (a) { return String(a || "").trim(); })
      .filter(Boolean);
    if (list.length) {
      body.toRecipients = list.map(function (a) { return { emailAddress: { address: a } }; });
    }
    return graphJson(token, "POST", "/me/messages", body);
  }

  root.GraphData = {
    signOut: signOut,
    currentAccount: currentAccount,
    isSignedOut: isSignedOut,
    getToken: getToken,
    resolveWorkbook: resolveWorkbook,
    recentWorkbooks: recentWorkbooks,
    sharedWorkbooks: sharedWorkbooks,
    searchWorkbooks: searchWorkbooks,
    listTables: listTables,
    listWorksheets: listWorksheets,
    usedRange: usedRange,
    addTable: addTable,
    uploadWorkbook: uploadWorkbook,
    ensureFolder: ensureFolder,
    tableHeaders: tableHeaders,
    addTableRow: addTableRow,
    tableRows: tableRows,
    updateTableRow: updateTableRow,
    tableColumns: tableColumns,
    addTableColumn: addTableColumn,
    authEmails: authEmails,
    createDraft: createDraft,
    setupInvites: setupInvites,
    messageBody: messageBody,
    bookingEmails: bookingEmails,
    _config: { clientId: CLIENT_ID },
  };
})(typeof self !== "undefined" ? self : this);
