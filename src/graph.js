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

  async function getToken() {
    var pca = await getPca();
    try {
      var silent = await pca.acquireTokenSilent({ scopes: SCOPES });
      return silent.accessToken;
    } catch (e) {
      var interactive = await pca.acquireTokenPopup({ scopes: SCOPES });
      return interactive.accessToken;
    }
  }

  async function graphJson(token, method, path, body) {
    var res = await fetch(GRAPH + path, {
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
  async function bookingEmails(token, sinceIso, senderDomains) {
    var filt = "receivedDateTime ge " + sinceIso;
    var res = await graphJson(token, "GET", "/me/mailFolders/inbox/messages?$filter=" +
      encodeURIComponent(filt) +
      "&$select=id,subject,bodyPreview,from,receivedDateTime,webLink&$orderby=receivedDateTime desc&$top=100");
    var domains = (senderDomains || []).map(function (d) { return String(d).toLowerCase().trim(); }).filter(Boolean);
    return (res.value || []).filter(function (m) {
      var a = ((m.from || {}).emailAddress || {}).address || "";
      var al = a.toLowerCase();
      return domains.some(function (d) { return al.indexOf(d) !== -1; });
    });
  }

  // ---------- mail ----------

  /** Create the Travel Authorization email as a DRAFT (never sent). */
  async function createDraft(token, to, subject, html) {
    var body = {
      subject: subject,
      body: { contentType: "HTML", content: html },
    };
    // An empty address is rejected outright (400 ErrorInvalidRecipients).
    // A third party with no billing contact on file, or a blank coordinator
    // address, should still get a draft — the user just fills the To line.
    var addr = String(to || "").trim();
    if (addr) { body.toRecipients = [{ emailAddress: { address: addr } }]; }
    return graphJson(token, "POST", "/me/messages", body);
  }

  root.GraphData = {
    getToken: getToken,
    resolveWorkbook: resolveWorkbook,
    recentWorkbooks: recentWorkbooks,
    sharedWorkbooks: sharedWorkbooks,
    searchWorkbooks: searchWorkbooks,
    listTables: listTables,
    tableHeaders: tableHeaders,
    addTableRow: addTableRow,
    createDraft: createDraft,
    bookingEmails: bookingEmails,
    _config: { clientId: CLIENT_ID },
  };
})(typeof self !== "undefined" ? self : this);
