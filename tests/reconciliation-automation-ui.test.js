const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
const appMain = fs.readFileSync(path.join(root, "app-main.js"), "utf8");

function appFunctionSource(name) {
  const start = appMain.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `${name} should be defined in app-main.js`);
  const bodyStart = appMain.indexOf("{", appMain.indexOf(")", start));
  let depth = 0;
  for (let index = bodyStart; index < appMain.length; index += 1) {
    if (appMain[index] === "{") depth += 1;
    if (appMain[index] === "}") depth -= 1;
    if (depth === 0) return appMain.slice(start, index + 1);
  }
  throw new Error(`Could not extract ${name} from app-main.js`);
}

function managedRules() {
  return [
    {
      ruleKey: "rule-b",
      ruleVersion: 2,
      displayName: "Second managed rule",
      baseSourceType: "financial_documents",
      destinationSourceTypes: ["import_cgd_extrato_ordem"],
      logicDescription: "Second rule logic",
      definition: { thresholds: { score: 0.8 } },
      enabled: false,
      allowManualExecution: true,
      includeInScheduledBatch: false,
      differenceAllowed: "2.50",
      maxDifferenceDays: 4,
      priority: 2,
    },
    {
      ruleKey: "rule-a",
      ruleVersion: 1,
      displayName: "First managed rule",
      baseSourceType: "financial_documents",
      destinationSourceTypes: ["import_cgd_extrato_ordem"],
      logicDescription: "First rule logic",
      definition: { thresholds: { score: 0.6 } },
      enabled: true,
      allowManualExecution: false,
      includeInScheduledBatch: true,
      differenceAllowed: "1.25",
      maxDifferenceDays: 7,
      priority: 1,
    },
  ];
}

function automationSettings(overrides = {}) {
  return {
    loaded: true,
    loading: false,
    activeTab: "automatic",
    schedule: { enabled: true, timeOfDay: "02:15", timeZone: "Europe/Lisbon" },
    rules: managedRules(),
    lastScheduledRun: null,
    ...overrides,
  };
}

function compilePayload(state) {
  return new Function(
    "state",
    `${appFunctionSource("clean")}
     ${appFunctionSource("reconciliationAutomationSettingsPayload")}
     return reconciliationAutomationSettingsPayload;`,
  )(state);
}

function fakeClassList() {
  const values = new Set();
  return {
    toggle(name, enabled) {
      if (enabled) values.add(name);
      else values.delete(name);
    },
    contains(name) {
      return values.has(name);
    },
  };
}

function renderAutomationSettings(settings) {
  const state = { reconciliationAutomationSettings: settings };
  const els = {
    financialReconciliationSettingsSourceTab: { classList: fakeClassList(), setAttribute(name, value) { this[name] = value; } },
    financialReconciliationSettingsAutomaticTab: { classList: fakeClassList(), setAttribute(name, value) { this[name] = value; } },
    financialReconciliationSettingsSourcePanel: { hidden: false },
    financialReconciliationSettingsAutomaticPanel: { hidden: true },
    financialReconciliationAutomationScheduleEnabled: { checked: false, disabled: false },
    financialReconciliationAutomationScheduleTime: { value: "", disabled: false },
    financialReconciliationAutomationLastExecution: { textContent: "" },
    financialReconciliationAutomationLastResult: { textContent: "" },
    financialReconciliationAutomationNextExecution: { textContent: "" },
    financialReconciliationAutomationRules: { innerHTML: "" },
    financialReconciliationAutomationSave: { disabled: false },
    financialReconciliationAutomationRunBatchNow: { disabled: false },
  };
  const payload = compilePayload(state);
  const render = new Function(
    "state",
    "els",
    "clean",
    "escape",
    "formatDateTimeShort",
    "financialReconciliationSourceLabel",
    "reconciliationAutomationSettingsPayload",
    `${appFunctionSource("renderReconciliationAutomationSettings")}
     return renderReconciliationAutomationSettings;`,
  )(
    state,
    els,
    (value) => String(value ?? "").trim(),
    (value) => String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;"),
    (value) => value ? "2026-08-14 02:00" : "-",
    (value) => ({
      financial_documents: "Financial Documents",
      import_cgd_extrato_ordem: "CGD Bank Statement",
    })[value] || value,
    payload,
  );
  render();
  return els;
}

test("reconciliation settings expose accessible source and automatic tabs", () => {
  for (const id of [
    "financial-reconciliation-settings-source-tab",
    "financial-reconciliation-settings-automatic-tab",
    "financial-reconciliation-settings-source-panel",
    "financial-reconciliation-settings-automatic-panel",
    "financial-reconciliation-automation-schedule-enabled",
    "financial-reconciliation-automation-schedule-time",
    "financial-reconciliation-automation-rules",
    "financial-reconciliation-automation-save",
    "financial-reconciliation-automation-run-batch-now",
  ]) {
    assert.match(html, new RegExp(`id="${id}"`));
  }
  assert.match(html, /role="tablist"[\s\S]*aria-controls="financial-reconciliation-settings-source-panel"[\s\S]*aria-controls="financial-reconciliation-settings-automatic-panel"/);
  assert.match(html, /id="financial-reconciliation-automation-time-zone"[^>]*>Europe\/Lisbon</);
});

test("actual renderer keeps managed definition text, versions, and thresholds read only and escaped", () => {
  const rule = {
    ...managedRules()[1],
    displayName: '<img src=x onerror="alert(1)">',
    logicDescription: "Use <script>alert(2)</script> and fixed thresholds.",
    definition: { descriptionSimilarity: 0.6, supplierSimilarity: 0.7, note: "<svg onload=alert(3)>" },
  };
  const els = renderAutomationSettings(automationSettings({ rules: [rule] }));
  const markup = els.financialReconciliationAutomationRules.innerHTML;

  assert.match(markup, /&lt;img src=x onerror=&quot;alert\(1\)&quot;&gt;/);
  assert.match(markup, /Use &lt;script&gt;alert\(2\)&lt;\/script&gt; and fixed thresholds\./);
  assert.match(markup, /&lt;svg onload=alert\(3\)&gt;/);
  assert.doesNotMatch(markup, /<img|<script|<svg/);
  assert.match(markup, /Version 1/);
  assert.match(markup, /descriptionSimilarity[\s\S]*0\.6[\s\S]*supplierSimilarity[\s\S]*0\.7/);
  assert.match(markup, /Managed definition \(read only\)/);
  assert.match(markup, /Financial Documents[\s\S]*financial_documents[\s\S]*CGD Bank Statement[\s\S]*import_cgd_extrato_ordem/);
  assert.doesNotMatch(markup, /data-reconciliation-automation-rule-field="(?:definition|logicDescription|ruleVersion)"/);
});

test("actual renderer summarizes the last scheduled result separately from execution timing", () => {
  const els = renderAutomationSettings(automationSettings({
    lastScheduledRun: {
      status: "partial",
      startedAt: "2026-08-14T02:00:00Z",
      finishedAt: "2026-08-14T02:01:00Z",
      counts: { proposed: 4, completed: 3, failed: 1 },
    },
  }));

  assert.equal(els.financialReconciliationAutomationLastExecution.textContent, "Last execution: 2026-08-14 02:00 · partial");
  assert.equal(els.financialReconciliationAutomationLastResult.textContent, "Last result: proposed 4 · completed 3 · failed 1");
});

test("actual serializer emits only approved schedule and managed-rule configuration fields", () => {
  const settings = automationSettings();
  settings.rules[0].logicDescription = "must not be sent";
  settings.rules[0].definition = { threshold: 0.8 };
  settings.rules[0].displayName = "must not be sent";
  const payload = compilePayload({ reconciliationAutomationSettings: settings })();

  assert.deepEqual(payload, {
    schedule: { enabled: true, timeOfDay: "02:15", timeZone: "Europe/Lisbon" },
    rules: [
      {
        ruleKey: "rule-a",
        ruleVersion: 1,
        enabled: true,
        allowManualExecution: false,
        includeInScheduledBatch: true,
        differenceAllowed: "1.25",
        maxDifferenceDays: 7,
        priority: 1,
      },
      {
        ruleKey: "rule-b",
        ruleVersion: 2,
        enabled: false,
        allowManualExecution: true,
        includeInScheduledBatch: false,
        differenceAllowed: "2.50",
        maxDifferenceDays: 4,
        priority: 2,
      },
    ],
  });
});

test("actual reorder helper produces stable unique consecutive priorities", () => {
  const state = { reconciliationAutomationSettings: automationSettings() };
  const move = new Function(
    "state",
    "renderReconciliationAutomationSettings",
    `${appFunctionSource("clean")}
     ${appFunctionSource("moveReconciliationAutomationRule")}
     return moveReconciliationAutomationRule;`,
  )(state, () => {});

  move("rule-b", "up");
  assert.deepEqual(
    state.reconciliationAutomationSettings.rules.map(({ ruleKey, priority }) => ({ ruleKey, priority })),
    [{ ruleKey: "rule-b", priority: 1 }, { ruleKey: "rule-a", priority: 2 }],
  );
  move("rule-b", "up");
  assert.deepEqual(state.reconciliationAutomationSettings.rules.map((rule) => rule.priority), [1, 2]);
});

test("invalid local automation values disable Save", () => {
  for (const invalid of [
    automationSettings({ schedule: { enabled: true, timeOfDay: "25:00", timeZone: "Europe/Lisbon" } }),
    automationSettings({ rules: [{ ...managedRules()[1], differenceAllowed: "-0.01" }] }),
    automationSettings({ rules: [{ ...managedRules()[1], differenceAllowed: "1.234" }] }),
    automationSettings({ rules: [{ ...managedRules()[1], differenceAllowed: "90071992547410.00" }] }),
    automationSettings({ rules: [{ ...managedRules()[1], maxDifferenceDays: "" }] }),
    automationSettings({ rules: [{ ...managedRules()[1], maxDifferenceDays: 366 }] }),
  ]) {
    const payload = compilePayload({ reconciliationAutomationSettings: invalid });
    assert.equal(payload(), null);
    assert.equal(renderAutomationSettings(invalid).financialReconciliationAutomationSave.disabled, true);
  }
});

test("automation inputs change only the local draft until one atomic Save", async () => {
  class FakeHTMLElement {
    constructor(dataset, values) {
      this.dataset = dataset;
      Object.assign(this, values);
    }
  }
  const settings = automationSettings();
  const state = { reconciliationAutomationSettings: settings };
  const requests = [];
  const input = new Function(
    "HTMLElement",
    "state",
    "clean",
    "renderReconciliationAutomationSettings",
    `${appFunctionSource("onReconciliationAutomationSettingsInput")}
     return onReconciliationAutomationSettingsInput;`,
  )(FakeHTMLElement, state, (value) => String(value ?? "").trim(), () => {});
  input({ target: new FakeHTMLElement({ reconciliationAutomationRuleKey: "rule-a", reconciliationAutomationRuleField: "differenceAllowed" }, { value: "3.25" }) });
  assert.equal(settings.rules.find((rule) => rule.ruleKey === "rule-a").differenceAllowed, "3.25");
  assert.deepEqual(requests, []);

  const payload = compilePayload(state);
  const save = new Function(
    "state",
    "api",
    "reconciliationAutomationSettingsPayload",
    "renderReconciliationAutomationSettings",
    "setReconciliationAutomationSettingsStatus",
    "applyReconciliationAutomationSettingsResult",
    `${appFunctionSource("saveReconciliationAutomationSettings").replace(/^function /, "async function ")}
     return saveReconciliationAutomationSettings;`,
  )(
    state,
    async (url, options) => {
      requests.push({ url, options });
      return { schedule: options.body.schedule, rules: settings.rules, lastScheduledRun: null };
    },
    payload,
    () => {},
    () => {},
    () => {},
  );
  await save();

  assert.equal(requests.length, 1);
  assert.equal(requests[0].url, "/api/reconciliation-automation-settings");
  assert.equal(requests[0].options.method, "PUT");
  assert.equal(requests[0].options.body.rules[0].differenceAllowed, "3.25");
});

test("Run batch now stores analysis, navigates to Reconciliation, renders, and never executes", async () => {
  const requests = [];
  const sequence = [];
  const run = { runId: "00000000-0000-0000-0000-000000000001", proposals: [{ id: "proposal-1", status: "proposed" }] };
  const current = { automation: { rules: [], run: null, selectedProposalIds: new Set(), pendingAction: "", loaded: false } };
  const state = { reconciliationAutomationSettings: automationSettings() };
  const runBatchNow = new Function(
    "state",
    "api",
    "crypto",
    "financialReconciliationState",
    "setView",
    "renderFinancialReconciliation",
    "renderReconciliationAutomationSettings",
    "setReconciliationAutomationSettingsStatus",
    `${appFunctionSource("runReconciliationAutomationBatchNow").replace(/^function /, "async function ")}
     return runReconciliationAutomationBatchNow;`,
  )(
    state,
    async (url, options) => {
      requests.push({ url, options });
      sequence.push("api");
      return run;
    },
    { randomUUID: () => "00000000-0000-0000-0000-000000000099" },
    () => current,
    async (view) => sequence.push(`view:${view}`),
    () => sequence.push(`render:${current.automation.run?.runId || "missing"}`),
    () => {},
    () => {},
  );

  await runBatchNow();

  assert.deepEqual(requests, [{
    url: "/api/reconciliation-automation",
    options: {
      method: "POST",
      body: { action: "analyze_batch", clientRequestId: "00000000-0000-0000-0000-000000000099" },
    },
  }]);
  assert.equal(current.automation.run, run);
  assert.deepEqual(sequence, ["api", "view:financial-reconciliation", `render:${run.runId}`]);
  assert.doesNotMatch(JSON.stringify(requests), /execute_selected|proposalIds/);
});
