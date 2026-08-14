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

function renderAutomationSettings(settings, { dirty = false, canOpenWorkbench = true } = {}) {
  const state = {
    reconciliationAutomationSettings: settings,
    reconciliationAutomationSettingsDirty: dirty,
  };
  const els = {
    financialReconciliationSettingsSourceTab: { classList: fakeClassList(), setAttribute(name, value) { this[name] = value; } },
    financialReconciliationSettingsAutomaticTab: { classList: fakeClassList(), setAttribute(name, value) { this[name] = value; } },
    financialReconciliationSettingsSourcePanel: { hidden: false },
    financialReconciliationSettingsAutomaticPanel: { hidden: true },
    financialReconciliationAutomationScheduleEnabled: { checked: false, disabled: false },
    financialReconciliationAutomationScheduleTime: {
      value: "",
      disabled: false,
      setAttribute(name, value) { this[name] = value; },
      removeAttribute(name) { delete this[name]; },
    },
    financialReconciliationAutomationLastExecution: { textContent: "" },
    financialReconciliationAutomationLastResult: { textContent: "" },
    financialReconciliationAutomationNextExecution: { textContent: "" },
    financialReconciliationAutomationRules: { innerHTML: "", querySelectorAll: () => [] },
    financialReconciliationAutomationSave: { disabled: false },
    financialReconciliationAutomationRunBatchNow: { disabled: false },
    financialReconciliationAutomationRunHint: { textContent: "" },
  };
  const payload = compilePayload(state);
  const clean = new Function(`${appFunctionSource("clean")}; return clean;`)();
  const updateControls = new Function(
    "state",
    "els",
    "clean",
    "canAppFinancialReconciliation",
    "reconciliationAutomationSettingsPayload",
    `${appFunctionSource("updateReconciliationAutomationControls")}
     return updateReconciliationAutomationControls;`,
  )(state, els, clean, () => canOpenWorkbench, payload);
  const render = new Function(
    "state",
    "els",
    "clean",
    "escape",
    "formatReconciliationAutomationDateTime",
    "financialReconciliationSourceLabel",
    "reconciliationAutomationSettingsPayload",
    "updateReconciliationAutomationControls",
    `${appFunctionSource("renderReconciliationAutomationSettings")}
     return renderReconciliationAutomationSettings;`,
  )(
    state,
    els,
    clean,
    new Function(`${appFunctionSource("escape")}; return escape;`)(),
    new Function(
      "clean",
      `${appFunctionSource("formatReconciliationAutomationDateTime")}; return formatReconciliationAutomationDateTime;`,
    )(clean),
    (value) => ({
      financial_documents: "Financial Documents",
      import_cgd_extrato_ordem: "CGD Bank Statement",
    })[value] || value,
    payload,
    updateControls,
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
  assert.match(html, /id="financial-reconciliation-automation-schedule-time"[^>]*aria-describedby="financial-reconciliation-automation-status"/);
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

  assert.equal(els.financialReconciliationAutomationLastExecution.textContent, "Last execution: 2026-08-14 03:01 Europe/Lisbon · partial");
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
  const state = { reconciliationAutomationSettings: automationSettings(), reconciliationAutomationSettingsDirty: false };
  const focusCalls = [];
  const move = new Function(
    "state",
    "renderReconciliationAutomationSettings",
    "focusReconciliationAutomationRuleMove",
    `${appFunctionSource("clean")}
     ${appFunctionSource("moveReconciliationAutomationRule")}
     return moveReconciliationAutomationRule;`,
  )(state, () => {}, (...args) => focusCalls.push(args));

  move("rule-b", "up");
  assert.deepEqual(
    state.reconciliationAutomationSettings.rules.map(({ ruleKey, priority }) => ({ ruleKey, priority })),
    [{ ruleKey: "rule-b", priority: 1 }, { ruleKey: "rule-a", priority: 2 }],
  );
  assert.equal(state.reconciliationAutomationSettingsDirty, true);
  assert.deepEqual(focusCalls[0], ["rule-b", "up"]);
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

test("empty time remains visible as invalid instead of being replaced with the default", () => {
  const invalid = automationSettings({
    schedule: { enabled: true, timeOfDay: "", timeZone: "Europe/Lisbon" },
  });
  const els = renderAutomationSettings(invalid);

  assert.equal(els.financialReconciliationAutomationScheduleTime.value, "");
  assert.equal(els.financialReconciliationAutomationScheduleTime["aria-invalid"], "true");
  assert.equal(els.financialReconciliationAutomationSave.disabled, true);
});

test("Run batch control requires saved settings, an eligible rule, and workbench access", () => {
  const dirtyEls = renderAutomationSettings(automationSettings(), { dirty: true });
  assert.equal(dirtyEls.financialReconciliationAutomationRunBatchNow.disabled, true);
  assert.match(dirtyEls.financialReconciliationAutomationRunHint.textContent, /save configuration changes/i);

  const noAccessEls = renderAutomationSettings(automationSettings(), { canOpenWorkbench: false });
  assert.equal(noAccessEls.financialReconciliationAutomationRunBatchNow.disabled, true);
  assert.match(noAccessEls.financialReconciliationAutomationRunHint.textContent, /reconciliation app access/i);

  const noRuleEls = renderAutomationSettings(automationSettings({
    rules: managedRules().map((rule) => ({ ...rule, includeInScheduledBatch: false })),
  }));
  assert.equal(noRuleEls.financialReconciliationAutomationRunBatchNow.disabled, true);
  assert.match(noRuleEls.financialReconciliationAutomationRunHint.textContent, /at least one rule/i);
});

test("automation inputs change only the local draft until one atomic Save", async () => {
  class FakeHTMLElement {
    constructor(dataset, values) {
      this.dataset = dataset;
      Object.assign(this, values);
    }
  }
  const settings = automationSettings();
  const state = {
    reconciliationAutomationSettings: settings,
    reconciliationAutomationSettingsDirty: false,
    financialReconciliation: { automation: { loaded: true } },
  };
  const requests = [];
  let validationUpdates = 0;
  const input = new Function(
    "HTMLElement",
    "state",
    "clean",
    "updateReconciliationAutomationNextExecution",
    "updateReconciliationAutomationControls",
    `${appFunctionSource("onReconciliationAutomationSettingsInput")}
     return onReconciliationAutomationSettingsInput;`,
  )(FakeHTMLElement, state, (value) => String(value ?? "").trim(), () => {}, () => { validationUpdates += 1; });
  input({ target: new FakeHTMLElement({ reconciliationAutomationRuleKey: "rule-a", reconciliationAutomationRuleField: "differenceAllowed" }, { value: "3.25" }) });
  assert.equal(settings.rules.find((rule) => rule.ruleKey === "rule-a").differenceAllowed, "3.25");
  assert.equal(state.reconciliationAutomationSettingsDirty, true);
  assert.equal(validationUpdates, 1);
  assert.deepEqual(requests, []);

  const payload = compilePayload(state);
  const statuses = [];
  const applyResult = new Function(
    "state",
    "clean",
    "clone",
    `${appFunctionSource("applyReconciliationAutomationSettingsResult")}
     return applyReconciliationAutomationSettingsResult;`,
  )(
    state,
    new Function(`${appFunctionSource("clean")}; return clean;`)(),
    new Function(`${appFunctionSource("clone")}; return clone;`)(),
  );
  const financialState = new Function(
    "state",
    `${appFunctionSource("financialReconciliationState")}
     return financialReconciliationState;`,
  )(state);
  const save = new Function(
    "state",
    "api",
    "reconciliationAutomationSettingsPayload",
    "renderReconciliationAutomationSettings",
    "setReconciliationAutomationSettingsStatus",
    "applyReconciliationAutomationSettingsResult",
    "financialReconciliationState",
    `${appFunctionSource("saveReconciliationAutomationSettings").replace(/^function /, "async function ")}
     return saveReconciliationAutomationSettings;`,
  )(
    state,
    async (url, options) => {
      requests.push({ url, options });
      return {
        schedule: options.body.schedule,
        rules: settings.rules.map((rule) => rule.ruleKey === "rule-a"
          ? { ...rule, ruleVersion: 9, displayName: "Authoritative name", definition: { authoritative: true } }
          : rule),
        lastScheduledRun: null,
      };
    },
    payload,
    () => {},
    (message, isError) => statuses.push({ message, isError }),
    applyResult,
    financialState,
  );
  await save();

  assert.equal(requests.length, 1);
  assert.equal(requests[0].url, "/api/reconciliation-automation-settings");
  assert.equal(requests[0].options.method, "PUT");
  assert.equal(requests[0].options.body.rules[0].differenceAllowed, "3.25");
  assert.equal(state.reconciliationAutomationSettingsDirty, false);
  assert.equal(state.reconciliationAutomationSettings.rules[0].ruleVersion, 9);
  assert.equal(state.reconciliationAutomationSettings.rules[0].displayName, "Authoritative name");
  assert.deepEqual(state.reconciliationAutomationSettings.rules[0].definition, { authoritative: true });
  assert.equal(state.financialReconciliation.automation.loaded, false);
  assert.deepEqual(statuses, [{ message: "Automatic reconciliation configuration saved.", isError: undefined }]);
});

test("Run batch now stores analysis, navigates to Reconciliation, renders, and never executes", async () => {
  const requests = [];
  const sequence = [];
  const run = { runId: "00000000-0000-0000-0000-000000000001", proposals: [{ id: "proposal-1", status: "proposed" }] };
  const existingRules = [{ ruleKey: "authoritative-workbench-rule" }];
  const current = { automation: { rules: existingRules, run: null, selectedProposalIds: new Set(), pendingAction: "", loaded: false } };
  const state = { reconciliationAutomationSettings: automationSettings(), reconciliationAutomationSettingsDirty: false };
  const runBatchNow = new Function(
    "state",
    "api",
    "crypto",
    "financialReconciliationState",
    "setView",
    "renderFinancialReconciliation",
    "renderReconciliationAutomationSettings",
    "setReconciliationAutomationSettingsStatus",
    "canAppFinancialReconciliation",
    "reconciliationAutomationSettingsPayload",
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
    () => true,
    compilePayload(state),
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
  assert.equal(current.automation.rules, existingRules);
  assert.deepEqual(sequence, ["api", "view:financial-reconciliation", `render:${run.runId}`]);
  assert.doesNotMatch(JSON.stringify(requests), /execute_selected|proposalIds/);
});

test("Run batch now refuses dispatch for dirty settings or missing workbench access", async () => {
  for (const scenario of [
    { dirty: true, canOpenWorkbench: true, expected: /save.*before running/i },
    { dirty: false, canOpenWorkbench: false, expected: /reconciliation app access/i },
  ]) {
    const requests = [];
    const statuses = [];
    const state = {
      reconciliationAutomationSettings: automationSettings(),
      reconciliationAutomationSettingsDirty: scenario.dirty,
    };
    const runBatchNow = new Function(
      "state",
      "api",
      "crypto",
      "financialReconciliationState",
      "setView",
      "renderFinancialReconciliation",
      "renderReconciliationAutomationSettings",
      "setReconciliationAutomationSettingsStatus",
      "canAppFinancialReconciliation",
      "reconciliationAutomationSettingsPayload",
      `${appFunctionSource("runReconciliationAutomationBatchNow").replace(/^function /, "async function ")}
       return runReconciliationAutomationBatchNow;`,
    )(
      state,
      async (...args) => requests.push(args),
      { randomUUID: () => "00000000-0000-0000-0000-000000000099" },
      () => ({ automation: {} }),
      async () => {},
      () => {},
      () => {},
      (message, isError) => statuses.push({ message, isError }),
      () => scenario.canOpenWorkbench,
      compilePayload(state),
    );

    await runBatchNow();

    assert.deepEqual(requests, []);
    assert.match(statuses.at(-1).message, scenario.expected);
    assert.equal(statuses.at(-1).isError, true);
  }
});

test("settings tabs support roving focus and horizontal keyboard activation", () => {
  const activations = [];
  const sourceTab = { focusCalled: 0, focus() { this.focusCalled += 1; } };
  const automaticTab = { focusCalled: 0, focus() { this.focusCalled += 1; } };
  const els = {
    financialReconciliationSettingsSourceTab: sourceTab,
    financialReconciliationSettingsAutomaticTab: automaticTab,
  };
  const onKeydown = new Function(
    "els",
    "setReconciliationSettingsTab",
    `${appFunctionSource("onReconciliationSettingsTabKeydown")}
     return onReconciliationSettingsTabKeydown;`,
  )(els, (tab) => activations.push(tab));
  let prevented = 0;

  onKeydown({ key: "ArrowRight", currentTarget: sourceTab, preventDefault: () => { prevented += 1; } });
  onKeydown({ key: "Home", currentTarget: automaticTab, preventDefault: () => { prevented += 1; } });

  assert.deepEqual(activations, ["automatic", "source-rules"]);
  assert.equal(automaticTab.focusCalled, 1);
  assert.equal(sourceTab.focusCalled, 1);
  assert.equal(prevented, 2);
});
