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

  // ---------- mail ----------

  /** Create the Travel Authorization email as a DRAFT (never sent). */
  async function createDraft(token, to, subject, html) {
    return graphJson(token, "POST", "/me/messages", {
      subject: subject,
      body: { contentType: "HTML", content: html },
      toRecipients: [{ emailAddress: { address: to } }],
    });
  }

  root.GraphData = {
    getToken: getToken,
    resolveWorkbook: resolveWorkbook,
    listTables: listTables,
    tableHeaders: tableHeaders,
    addTableRow: addTableRow,
    createDraft: createDraft,
    _config: { clientId: CLIENT_ID },
  };
})(typeof self !== "undefined" ? self : this);
