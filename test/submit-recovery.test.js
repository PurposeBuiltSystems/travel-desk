"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const raw = fs.readFileSync(path.join(__dirname, "../src/taskpane/taskpane.js"), "utf8");
const end = raw.lastIndexOf("})();");
const source = raw.slice(0, end) + `
model=()=>self.form;settings=()=>self.settings;saveSettings=()=>{};
addTrip=()=>self.recorded++;logoAttachment=()=>[];
self.workflow={submit};` + raw.slice(end);
function load() {
  const elements = {}, calls = { draft: 0, rows: [], recorded: 0 };
  function element() { return { value: "", checked: false, disabled: false, hidden: true,
    textContent: "", children: [], addEventListener(event, fn) { this[event] = fn; },
    appendChild(child) { this.children.push(child); },
    set innerHTML(value) { this.children = []; } }; }
  const el = id => elements[id] || (elements[id] = element());
  const self = { form: { name: "Traveler", event: "Conference", location: "City" },
    settings: { wbRef: { driveId: "d", itemId: "w", name: "Planner" }, tableName: "Travel" }, recorded: 0 };
  const graph = { getToken: async () => "token", currentAccount: async () => "owner@example.org",
    createDraft: async () => { calls.draft++; return { id: "draft" }; },
    tableHeaders: async () => ["Name"], addTableRow: async (token, ref, table, row) => { calls.rows.push(row[0]); } };
  const context = { self, window: {}, document: { getElementById: el, createElement: element },
    Office: { onReady() {}, context: { ui: {} } }, GraphData: graph,
    TravelForm: { subjectLine: () => "Travel Auth", formHtml: () => "<p>request</p>",
      fiscalLabel: () => "FY27", pickPlanner: () => null, plannerRows: () => [["Traveler"], ["Other traveler"]] } };
  vm.runInNewContext(source, context);
  el("doDraft").checked = el("doRow").checked = true;
  el("coordEmail").value = "coordinator@example.org";
  return { api: self.workflow, self, elements, calls, graph };
}
(async () => {
  const retry = load(); let rejectSecond = true;
  retry.graph.addTableRow = async (token, ref, table, row) => {
    if (row[0] === "Other traveler" && rejectSecond) { const e = Error("access denied"); e.status = 403; throw e; }
    retry.calls.rows.push(row[0]);
  };
  await retry.api.submit();
  assert.deepEqual(retry.calls.rows, ["Traveler"]);
  assert.match(retry.elements.status.textContent, /1 planner row\(s\) already saved/);
  rejectSecond = false; await retry.api.submit(); await retry.api.submit();
  assert.equal(retry.calls.draft, 1, "do not duplicate a saved draft on retry");
  assert.deepEqual(retry.calls.rows, ["Traveler", "Other traveler"], "resume after confirmed rows");
  assert.equal(retry.self.recorded, 1, "one trip-history entry across retries");
  retry.elements.coordEmail.value = "corrected@example.org";
  retry.graph.tableHeaders = async () => ["Different header order"];
  await retry.api.submit();
  assert.equal(retry.calls.draft, 2, "a corrected coordinator gets a new draft");
  assert.equal(retry.calls.rows.length, 2, "coordinator or header changes do not duplicate saved planner rows");

  const uncertain = load(); let lost = true;
  uncertain.graph.addTableRow = async (token, ref, table, row) => {
    uncertain.calls.rows.push(row[0]);
    if (lost) { lost = false; throw Error("response lost after append"); }
  };
  await uncertain.api.submit(); await uncertain.api.submit();
  assert.deepEqual(uncertain.calls.rows, ["Traveler"], "unknown outcome blocks blind retries");
  assert.equal(uncertain.elements.submissionRecovery.hidden, false);
  uncertain.elements.submissionRecovery.children[1].click(); // traveler found the saved row
  await new Promise(setImmediate);
  assert.deepEqual(uncertain.calls.rows, ["Traveler", "Other traveler"]);

  const draftLost = load();
  draftLost.graph.createDraft = async () => { draftLost.calls.draft++; throw Error("draft response lost"); };
  await draftLost.api.submit(); await draftLost.api.submit();
  assert.equal(draftLost.calls.draft, 1);
  assert.equal(draftLost.calls.rows.length, 0, "resolve uncertain draft before planner writes");
  draftLost.elements.submissionRecovery.children[1].click(); // draft is present
  await new Promise(setImmediate);
  assert.equal(draftLost.calls.draft, 1);
  assert.equal(draftLost.calls.rows.length, 2);

  const busy = load(); let release;
  busy.graph.getToken = () => new Promise(resolve => { release = resolve; });
  const active = busy.api.submit(); await busy.api.submit(); release("token"); await active;
  assert.equal(busy.calls.draft, 1, "ignore overlapping submits");
  assert.equal(busy.elements.submit.disabled, false);
  console.log("Travel submission recovery regressions passed (all Graph writes mocked).");
})().catch(error => { console.error(error); process.exitCode = 1; });
