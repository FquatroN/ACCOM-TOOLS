const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
const appMain = fs.readFileSync(path.join(root, "app-main.js"), "utf8");

function appFunctionSource(name) {
  const functionStart = appMain.indexOf(`function ${name}(`);
  assert.notEqual(functionStart, -1, `${name} should be defined in app-main.js`);
  const start = appMain.slice(Math.max(0, functionStart - 6), functionStart) === "async " ? functionStart - 6 : functionStart;
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

test("Reconciliation entry defaults to Manual and accepts only an explicit Automatic handoff", () => {
  const entryTab = new Function(
    `${appFunctionSource("clean")}\n${appFunctionSource("financialReconciliationEntryTab")}\nreturn financialReconciliationEntryTab;`,
  )();
  assert.equal(entryTab(), "manual");
  assert.equal(entryTab({}), "manual");
  assert.equal(entryTab({ financialReconciliationTab: "manual" }), "manual");
  assert.equal(entryTab({ financialReconciliationTab: "automatic" }), "automatic");
  assert.equal(entryTab({ financialReconciliationTab: "AUTOMATIC" }), "manual");
});

test("Run batch now stores analysis, navigates to Reconciliation Automatic, and never executes", async () => {
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
    async (view, options) => sequence.push(`view:${view}:${options?.financialReconciliationTab || "manual"}`),
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
  assert.equal(current.automation.rules, existingRules);
  assert.deepEqual(sequence, ["api", "view:financial-reconciliation:automatic"]);
  assert.equal(current.automation.run, run);
  assert.deepEqual([...current.automation.selectedProposalIds], ["proposal-1"]);
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

const WORKBENCH_RUN_ID = "00000000-0000-0000-0000-000000000101";
const WORKBENCH_PROPOSAL_1 = "00000000-0000-0000-0000-000000000102";
const WORKBENCH_PROPOSAL_2 = "00000000-0000-0000-0000-000000000103";
const WORKBENCH_PROPOSAL_3 = "00000000-0000-0000-0000-000000000104";

function workbenchRules() {
  return [
    {
      ruleKey: "manual-enabled",
      ruleVersion: 3,
      displayName: "Manual enabled",
      baseSourceType: "financial_documents",
      destinationSourceTypes: ["import_cgd_extrato_ordem"],
      enabled: true,
      allowManualExecution: true,
      differenceAllowed: "1.00",
      maxDifferenceDays: 7,
      priority: 1,
    },
    {
      ruleKey: "disabled",
      ruleVersion: 1,
      displayName: "Disabled rule",
      enabled: false,
      allowManualExecution: true,
      differenceAllowed: "0.00",
      maxDifferenceDays: 7,
      priority: 2,
    },
    {
      ruleKey: "not-manual",
      ruleVersion: 1,
      displayName: "Scheduled only",
      enabled: true,
      allowManualExecution: false,
      differenceAllowed: "0.00",
      maxDifferenceDays: 7,
      priority: 3,
    },
  ];
}

function workbenchRun(proposals) {
  return {
    runId: WORKBENCH_RUN_ID,
    trigger: "manual",
    scope: "rule",
    status: "ready",
    definitions: [{
      ruleKey: "manual-enabled",
      ruleVersion: 3,
      operator: "-",
      differenceAllowed: "1.00",
      maxDifferenceDays: 7,
    }],
    counts: {},
    proposals,
  };
}

function compileVisibleAutomationProposals() {
  return new Function(
    "clean",
    `${appFunctionSource("financialReconciliationAutomationVisibleProposals")}
     return financialReconciliationAutomationVisibleProposals;`,
  )((value) => String(value ?? "").trim());
}

function compileAutomationOutcomeCounts() {
  return new Function(
    "clean",
    `${appFunctionSource("financialReconciliationAutomationOutcomeCounts")}
     return financialReconciliationAutomationOutcomeCounts;`,
  )((value) => String(value ?? "").trim());
}

function renderAutomationWorkbench(run, selectedProposalIds = new Set()) {
  const state = { automation: { rules: [], run, selectedProposalIds, pendingAction: "", loaded: true } };
  const proposalContainer = { innerHTML: "", querySelectorAll: () => [] };
  const els = {
    financialReconciliationWorkbenchAutomationRules: { innerHTML: "" },
    financialReconciliationWorkbenchAutomationProposals: proposalContainer,
    financialReconciliationWorkbenchAutomationResults: { innerHTML: "" },
    financialReconciliationWorkbenchAutomationSelectAll: { disabled: false },
    financialReconciliationWorkbenchAutomationClearAll: { disabled: false },
    financialReconciliationWorkbenchAutomationExecute: { disabled: false, textContent: "" },
  };
  const render = new Function(
    "financialReconciliationState",
    "clean",
    "els",
    "financialReconciliationAutomationRulesMarkup",
    "financialReconciliationAutomationProposalMarkup",
    "financialReconciliationAutomationResultsMarkup",
    "escape",
    `${appFunctionSource("financialReconciliationAutomationVisibleProposals")}
     ${appFunctionSource("financialReconciliationAutomationEmptyMessage")}
     ${appFunctionSource("renderFinancialReconciliationAutomation")}
     return renderFinancialReconciliationAutomation;`,
  )(
    () => state,
    (value) => String(value ?? "").trim(),
    els,
    () => "",
    (proposal) => `<article data-proposal-id="${proposal.id}">${proposal.id}</article>`,
    () => "",
    (value) => String(value ?? "").replace(/[&<>"']/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[character]),
  );
  render();
  return { els, state };
}

test("active automation runs show proposed and ambiguous rows only", () => {
  const visible = compileVisibleAutomationProposals();
  const proposals = [
    { id: "checked", status: "proposed" },
    { id: "unchecked", status: "proposed" },
    { id: "ambiguous", status: "ambiguous" },
    { id: "skipped", status: "skipped" },
    { id: "completed", status: "completed" },
  ];
  const run = Object.freeze({ finishedAt: null, proposals: Object.freeze(proposals) });

  assert.deepEqual(visible(run).map((proposal) => proposal.id), ["checked", "unchecked", "ambiguous"]);
  assert.equal(run.proposals, proposals);
});

test("finished automation runs show selected persisted outcomes only", () => {
  const visible = compileVisibleAutomationProposals();
  const run = {
    finishedAt: "2026-08-16T10:00:00.000Z",
    proposals: [
      { id: "completed", status: "completed" },
      { id: "stale", status: "stale" },
      { id: "failed", status: "failed" },
      { id: "ambiguous", status: "ambiguous" },
      { id: "skipped", status: "skipped" },
      { id: "deselected", status: "deselected" },
    ],
  };

  assert.deepEqual(visible(run).map((proposal) => proposal.id), ["completed", "stale", "failed"]);
});

test("active proposal rendering retains unchecked proposals and counts hidden outcomes", () => {
  const run = workbenchRun([
    { id: "checked", status: "proposed" },
    { id: "unchecked", status: "proposed" },
    { id: "ambiguous", status: "ambiguous" },
    { id: "skipped", status: "skipped" },
  ]);
  const { els } = renderAutomationWorkbench(run, new Set(["checked"]));

  assert.match(els.financialReconciliationWorkbenchAutomationProposals.innerHTML, /checked/);
  assert.match(els.financialReconciliationWorkbenchAutomationProposals.innerHTML, /unchecked/);
  assert.deepEqual(compileAutomationOutcomeCounts()(run), {
    completed: 0, stale: 0, failed: 0, ambiguous: 1, skipped: 1, deselected: 0, attemptFailures: 0,
  });
});

test("finished proposal rendering shows execution outcomes and counts hidden rows", () => {
  const run = workbenchRun([
    { id: "completed", status: "completed" },
    { id: "stale", status: "stale" },
    { id: "failed", status: "failed" },
    { id: "ambiguous", status: "ambiguous" },
    { id: "skipped", status: "skipped" },
    { id: "deselected", status: "deselected" },
  ]);
  run.finishedAt = "2026-08-16T10:00:00.000Z";
  const { els } = renderAutomationWorkbench(run);
  const markup = els.financialReconciliationWorkbenchAutomationProposals.innerHTML;

  assert.match(markup, /completed/);
  assert.match(markup, /stale/);
  assert.match(markup, /failed/);
  assert.doesNotMatch(markup, /ambiguous|skipped|deselected/);
  assert.deepEqual(compileAutomationOutcomeCounts()(run), {
    completed: 1, stale: 1, failed: 1, ambiguous: 1, skipped: 1, deselected: 1, attemptFailures: 0,
  });
});

function compileWorkbenchProposalMarkup() {
  return new Function(
    "clean",
    "escape",
    "financialReconciliationSourceLabel",
    "formatDateOnly",
    "formatMoney",
    `${appFunctionSource("financialReconciliationAutomationReasonLabel")}
     ${appFunctionSource("financialReconciliationAutomationRunDefinition")}
     ${appFunctionSource("financialReconciliationAutomationIdentityEvidenceMarkup")}
     ${appFunctionSource("financialReconciliationAutomationItemMarkup")}
     ${appFunctionSource("financialReconciliationAutomationProposalMarkup")}
     return financialReconciliationAutomationProposalMarkup;`,
  )(
    (value) => String(value ?? "").trim(),
    new Function(`${appFunctionSource("escape")}; return escape;`)(),
    (value) => ({
      financial_documents: "Financial Documents",
      import_cgd_extrato_ordem: "CGD Bank Statement",
    })[value] || value,
    (value) => String(value || "").slice(0, 10),
    (value) => `${Number(value).toFixed(2)} â‚¬`,
  );
}

test("workbench shows Analyze only for enabled manual rules from the app catalog", () => {
  const rulesMarkup = new Function(
    "clean",
    "escape",
    "financialReconciliationSourceLabel",
    "formatMoney",
    `${appFunctionSource("financialReconciliationAutomationRulesMarkup")}
     return financialReconciliationAutomationRulesMarkup;`,
  )(
    (value) => String(value ?? "").trim(),
    new Function(`${appFunctionSource("escape")}; return escape;`)(),
    (value) => ({ financial_documents: "Financial Documents", import_cgd_extrato_ordem: "CGD Bank Statement" })[value] || value,
    (value) => `${Number(value).toFixed(2)} â‚¬`,
  );

  const markup = rulesMarkup(workbenchRules(), "");
  assert.equal((markup.match(/data-financial-reconciliation-automation-analyze/g) || []).length, 1);
  assert.match(markup, /data-financial-reconciliation-automation-rule-key="manual-enabled"/);
  assert.match(markup, /Manual enabled[\s\S]*Financial Documents[\s\S]*CGD Bank Statement[\s\S]*1\.00 â‚¬[\s\S]*7 days/);
  assert.doesNotMatch(markup, /Disabled rule|Scheduled only/);
});

test("manual rule loader uses the app-authorized catalog without schedule administration", async () => {
  const current = {
    automation: { rules: [], run: null, selectedProposalIds: new Set(), pendingAction: "", loaded: false },
  };
  const calls = [];
  const loadRules = new Function(
    "financialReconciliationState",
    "api",
    "clean",
    "clone",
    "renderFinancialReconciliationAutomation",
    "setFinancialReconciliationAutomationStatus",
    `${appFunctionSource("loadFinancialReconciliationAutomationRules").replace(/^function /, "async function ")}
     return loadFinancialReconciliationAutomationRules;`,
  )(
    () => current,
    async (url) => {
      calls.push(url);
      return { rules: workbenchRules(), schedule: { enabled: true }, diagnostic: "hidden" };
    },
    (value) => String(value ?? "").trim(),
    (value) => JSON.parse(JSON.stringify(value)),
    () => {},
    () => {},
  );

  await loadRules();

  assert.deepEqual(calls, ["/api/reconciliation-automation?view=rules"]);
  assert.deepEqual(current.automation.rules, workbenchRules());
  assert.equal(current.automation.loaded, true);
  assert.equal(Object.hasOwn(current.automation, "schedule"), false);
});

test("failed manual catalog reload clears stale Analyze rules without clearing the retained run or selection", async () => {
  const retainedRun = { runId: "retained-run" };
  const retainedSelections = new Set(["retained-proposal"]);
  const current = {
    automation: { rules: workbenchRules(), run: retainedRun, selectedProposalIds: retainedSelections, pendingAction: "", loaded: false },
  };
  const loadRules = new Function(
    "financialReconciliationState",
    "api",
    "clean",
    "clone",
    "renderFinancialReconciliationAutomation",
    "setFinancialReconciliationAutomationStatus",
    `${appFunctionSource("loadFinancialReconciliationAutomationRules").replace(/^function /, "async function ")}
     return loadFinancialReconciliationAutomationRules;`,
  )(
    () => current,
    async () => { throw new Error("catalog unavailable"); },
    (value) => String(value ?? "").trim(),
    (value) => JSON.parse(JSON.stringify(value)),
    () => {},
    () => {},
  );

  await loadRules();

  assert.deepEqual(current.automation.rules, []);
  assert.equal(current.automation.loaded, false);
  assert.strictEqual(current.automation.run, retainedRun);
  assert.strictEqual(current.automation.selectedProposalIds, retainedSelections);
});

function compileEnsureFinancialReconciliationData({ current, loadWorkspace, loadRules, render }) {
  return new Function(
    "financialReconciliationState",
    "loadFinancialReconciliationWorkspace",
    "loadFinancialReconciliationAutomationRules",
    "renderFinancialReconciliation",
    `${appFunctionSource("ensureFinancialReconciliationData")}
     return ensureFinancialReconciliationData;`,
  )(
    () => current,
    loadWorkspace,
    loadRules,
    render,
  );
}

test("Reconciliation loads Manual first and lazily loads Automatic rules once", async () => {
  const calls = [];
  const current = { loaded: false, activeTab: "manual", automation: { loaded: false } };
  const ensure = compileEnsureFinancialReconciliationData({
    current,
    loadWorkspace: async () => { calls.push("workspace"); current.loaded = true; },
    loadRules: async () => { calls.push("rules"); current.automation.loaded = true; },
    render: () => calls.push("render"),
  });

  await ensure();
  assert.deepEqual(calls, ["workspace", "render"]);

  calls.length = 0;
  current.activeTab = "automatic";
  await ensure();
  assert.deepEqual(calls, ["rules", "render"]);

  calls.length = 0;
  await ensure();
  assert.deepEqual(calls, ["render"]);
});

test("Analyze sends one rule and a fresh UUID, clears old selection, and selects only proposals", async () => {
  const current = {
    automation: {
      rules: workbenchRules(),
      run: null,
      selectedProposalIds: new Set(["old-proposal"]),
      pendingAction: "",
      loaded: true,
    },
  };
  const calls = [];
  const run = workbenchRun([
    { id: WORKBENCH_PROPOSAL_1, status: "proposed" },
    { id: WORKBENCH_PROPOSAL_2, status: "ambiguous", reason: "multiple_combinations" },
  ]);
  const analyze = new Function(
    "financialReconciliationState",
    "api",
    "crypto",
    "clean",
    "renderFinancialReconciliationAutomation",
    "setFinancialReconciliationAutomationStatus",
    "showToast",
    `${appFunctionSource("analyzeFinancialReconciliationAutomationRule").replace(/^function /, "async function ")}
     return analyzeFinancialReconciliationAutomationRule;`,
  )(
    () => current,
    async (url, options) => {
      calls.push({ url, options });
      return run;
    },
    { randomUUID: () => "00000000-0000-0000-0000-000000000199" },
    (value) => String(value ?? "").trim(),
    () => {},
    () => {},
    () => {},
  );

  await analyze("manual-enabled");

  assert.deepEqual(calls, [{
    url: "/api/reconciliation-automation",
    options: {
      method: "POST",
      body: {
        action: "analyze_rule",
        ruleKeys: ["manual-enabled"],
        clientRequestId: "00000000-0000-0000-0000-000000000199",
      },
    },
  }]);
  assert.doesNotMatch(JSON.stringify(calls), /proposalIds/);
  assert.equal(current.automation.run, run);
  assert.deepEqual([...current.automation.selectedProposalIds], [WORKBENCH_PROPOSAL_1]);
  assert.equal(current.automation.pendingAction, "");
});

test("failed analysis preserves the displayed run and its prior selected proposals", async () => {
  const retainedRun = workbenchRun([{ id: WORKBENCH_PROPOSAL_1, status: "proposed" }]);
  const retainedSelections = new Set([WORKBENCH_PROPOSAL_1]);
  const current = {
    automation: {
      rules: workbenchRules(),
      run: retainedRun,
      selectedProposalIds: retainedSelections,
      pendingAction: "",
      loaded: true,
    },
  };
  const analyze = new Function(
    "financialReconciliationState",
    "api",
    "crypto",
    "clean",
    "renderFinancialReconciliationAutomation",
    "setFinancialReconciliationAutomationStatus",
    "showToast",
    `${appFunctionSource("analyzeFinancialReconciliationAutomationRule").replace(/^function /, "async function ")}
     return analyzeFinancialReconciliationAutomationRule;`,
  )(
    () => current,
    async () => { throw new Error("analysis unavailable"); },
    { randomUUID: () => "00000000-0000-0000-0000-000000000299" },
    (value) => String(value ?? "").trim(),
    () => {},
    () => {},
    () => {},
  );

  await analyze("manual-enabled");

  assert.strictEqual(current.automation.run, retainedRun);
  assert.strictEqual(current.automation.selectedProposalIds, retainedSelections);
  assert.deepEqual([...current.automation.selectedProposalIds], [WORKBENCH_PROPOSAL_1]);
  assert.equal(current.automation.pendingAction, "");
});

test("proposal markup starts executable rows selected and audits every ambiguous candidate group", () => {
  const proposalMarkup = compileWorkbenchProposalMarkup();
  const baseSnapshot = {
    sourceType: "financial_documents",
    sourceId: "document-1",
    sourceDate: "2026-08-01",
    docNumber: "FT 2026/55",
    description: "Invoice <script>alert(1)</script>",
    supplierName: "Safe Supplier",
    amount: 101,
  };
  const bank = (id, description, score) => ({
    sourceType: "import_cgd_extrato_ordem",
    sourceId: id,
    sourceDate: "2026-08-02",
    description,
    amount: -100,
    evidence: {
      documentNumber: { matched: true, normalized: "FT202655" },
      description: { matched: true, score, threshold: 0.6 },
      supplier: { matched: false, score: 0.4, threshold: 0.7 },
    },
  });
  const proposed = {
    id: WORKBENCH_PROPOSAL_1,
    ruleKey: "manual-enabled",
    ruleVersion: 3,
    status: "proposed",
    baseSnapshot,
    items: [bank("bank-1", "Primary bank row", 0.75)],
    candidateGroups: [[bank("bank-1", "Primary bank row", 0.75)]],
    calculatedDifference: 1,
    allowedDifference: 1,
  };
  const ambiguous = {
    ...proposed,
    id: WORKBENCH_PROPOSAL_2,
    status: "ambiguous",
    reason: "multiple_combinations",
    items: [],
    candidateGroups: [
      [bank("bank-a", "Candidate group A", 0.61)],
      [bank("bank-b", "Candidate group B", 0.62)],
    ],
  };
  const run = workbenchRun([proposed, ambiguous]);

  const proposedMarkup = proposalMarkup(proposed, run, workbenchRules(), new Set([WORKBENCH_PROPOSAL_1]), false);
  const ambiguousMarkup = proposalMarkup(ambiguous, run, workbenchRules(), new Set(), false);

  assert.match(proposedMarkup, /type="checkbox"[^>]*checked/);
  assert.doesNotMatch(proposedMarkup, /type="checkbox"[^>]*disabled/);
  assert.match(proposedMarkup, /Financial Documents[\s\S]*2026-08-01[\s\S]*FT 2026\/55[\s\S]*Safe Supplier[\s\S]*101\.00 â‚¬/);
  assert.match(proposedMarkup, /CGD Bank Statement[\s\S]*Primary bank row[\s\S]*-100\.00 â‚¬[\s\S]*Operator -/);
  assert.match(proposedMarkup, new RegExp(`Document number matched[\\s\\S]*FT202655[\\s\\S]*Description score 0\\.750 ${"\u2265"} 0\\.600`));
  assert.match(proposedMarkup, /Difference 1\.00 â‚¬[\s\S]*Allowed 1\.00 â‚¬[\s\S]*Manual enabled[\s\S]*version 3/i);
  assert.match(proposedMarkup, /Invoice &lt;script&gt;alert\(1\)&lt;\/script&gt;/);
  assert.doesNotMatch(proposedMarkup, /<script>/);

  assert.match(ambiguousMarkup, /type="checkbox"[^>]*disabled/);
  assert.match(ambiguousMarkup, /Multiple qualifying combinations/);
  assert.match(ambiguousMarkup, /Candidate group 1[\s\S]*Candidate group A[\s\S]*Candidate group 2[\s\S]*Candidate group B/);
  assert.match(proposedMarkup, new RegExp(`aria-label="Execute automatic proposal for Financial Documents record ${baseSnapshot.sourceId}"`));
});

test("scheduled-only batch proposals use the immutable friendly rule name", () => {
  const proposalMarkup = compileWorkbenchProposalMarkup();
  const proposal = {
    id: WORKBENCH_PROPOSAL_1,
    ruleKey: "scheduled-only-rule",
    ruleVersion: 8,
    status: "proposed",
    baseSnapshot: { sourceType: "financial_documents", sourceId: "document-8", sourceDate: "2026-08-08", description: "Batch", amount: 8 },
    items: [],
    candidateGroups: [],
    calculatedDifference: 0,
    allowedDifference: 0,
  };
  const run = workbenchRun([proposal]);
  run.definitions = [{ ruleKey: "scheduled-only-rule", ruleVersion: 8, displayName: "Scheduled friendly name", operator: "-", differenceAllowed: "0.00" }];

  const markup = proposalMarkup(proposal, run, [], new Set([WORKBENCH_PROPOSAL_1]), false);

  assert.match(markup, /Scheduled friendly name[\s\S]*version 8/i);
});

test("candidate-limit ambiguity labels a flat candidate list without inventing combinations", () => {
  const proposalMarkup = compileWorkbenchProposalMarkup();
  const candidate = (id) => ({ sourceType: "import_cgd_extrato_ordem", sourceId: id, sourceDate: "2026-08-02", description: id, amount: -1 });
  const proposal = {
    id: WORKBENCH_PROPOSAL_2,
    ruleKey: "manual-enabled",
    ruleVersion: 3,
    status: "ambiguous",
    reason: "candidate_limit",
    baseSnapshot: { sourceType: "financial_documents", sourceId: "document-limit", sourceDate: "2026-08-01", description: "Limit", amount: 13 },
    items: [],
    candidateGroups: [candidate("candidate-a"), candidate("candidate-b")],
    allowedDifference: 1,
  };

  const markup = proposalMarkup(proposal, workbenchRun([proposal]), workbenchRules(), new Set(), false);

  assert.match(markup, /Candidate 1[\s\S]*candidate-a[\s\S]*Candidate 2[\s\S]*candidate-b/);
  assert.doesNotMatch(markup, /Candidate group/);
});

test("completed proposals do not invent a generic failure reason", () => {
  const proposalMarkup = compileWorkbenchProposalMarkup();
  const proposal = {
    id: WORKBENCH_PROPOSAL_1,
    ruleKey: "manual-enabled",
    ruleVersion: 3,
    status: "completed",
    reason: "",
    baseSnapshot: { sourceType: "financial_documents", sourceId: "document-complete", sourceDate: "2026-08-01", description: "Complete", amount: 1 },
    items: [],
    candidateGroups: [],
    calculatedDifference: 0,
    allowedDifference: 1,
  };

  const markup = proposalMarkup(proposal, workbenchRun([proposal]), workbenchRules(), new Set(), false);

  assert.doesNotMatch(markup, /Proposal is not executable/);
});

test("all authoritative proposal reasons have safe auditable labels", () => {
  const reasonLabel = new Function(
    "clean",
    `${appFunctionSource("financialReconciliationAutomationReasonLabel")}
     return financialReconciliationAutomationReasonLabel;`,
  )((value) => String(value ?? "").trim());
  const expected = {
    no_qualifying_combination: "No qualifying destination combination",
    rule_snapshot_changed: "Rule configuration changed after analysis",
    operator_changed: "Signed operator changed after analysis",
    tolerance_changed: "Allowed tolerance changed after analysis",
    combination_changed: "Candidate combination changed after analysis",
    proposal_evidence_changed: "Identity evidence changed after analysis",
    not_selected: "Not selected for execution",
  };
  for (const [reason, label] of Object.entries(expected)) assert.equal(reasonLabel(reason), label);
});

test("select-all and clear-all change executable proposals only", () => {
  const current = {
    automation: {
      rules: workbenchRules(),
      run: workbenchRun([
        { id: WORKBENCH_PROPOSAL_1, status: "proposed" },
        { id: WORKBENCH_PROPOSAL_2, status: "ambiguous" },
        { id: WORKBENCH_PROPOSAL_3, status: "stale" },
      ]),
      selectedProposalIds: new Set(),
      pendingAction: "",
      loaded: true,
    },
  };
  const functions = new Function(
    "financialReconciliationState",
    "clean",
    "renderFinancialReconciliationAutomation",
    `${appFunctionSource("setFinancialReconciliationAutomationSelection")}
     ${appFunctionSource("toggleFinancialReconciliationAutomationProposal")}
     return { setFinancialReconciliationAutomationSelection, toggleFinancialReconciliationAutomationProposal };`,
  )(
    () => current,
    (value) => String(value ?? "").trim(),
    () => {},
  );

  functions.setFinancialReconciliationAutomationSelection("all");
  assert.deepEqual([...current.automation.selectedProposalIds], [WORKBENCH_PROPOSAL_1]);
  functions.toggleFinancialReconciliationAutomationProposal(WORKBENCH_PROPOSAL_2, true);
  assert.deepEqual([...current.automation.selectedProposalIds], [WORKBENCH_PROPOSAL_1]);
  let restoredFocusId = "";
  const focusAwareFunctions = new Function(
    "financialReconciliationState",
    "clean",
    "renderFinancialReconciliationAutomation",
    `${appFunctionSource("toggleFinancialReconciliationAutomationProposal")}
     return { toggleFinancialReconciliationAutomationProposal };`,
  )(
    () => current,
    (value) => String(value ?? "").trim(),
    (value) => { restoredFocusId = value; },
  );
  focusAwareFunctions.toggleFinancialReconciliationAutomationProposal(WORKBENCH_PROPOSAL_1, false);
  assert.equal(restoredFocusId, WORKBENCH_PROPOSAL_1);
  functions.setFinancialReconciliationAutomationSelection("clear");
  assert.deepEqual([...current.automation.selectedProposalIds], []);
});

test("Execute selected copies checked IDs once, is pending-safe, refreshes history, and blocks zero selection", async () => {
  const proposals = [WORKBENCH_PROPOSAL_1, WORKBENCH_PROPOSAL_2, WORKBENCH_PROPOSAL_3]
    .map((id) => ({ id, status: "proposed" }));
  const current = {
    loaded: true,
    automation: {
      rules: workbenchRules(),
      run: workbenchRun(proposals),
      selectedProposalIds: new Set(proposals.map((proposal) => proposal.id)),
      pendingAction: "",
      loaded: true,
    },
  };
  const refreshedRun = workbenchRun([
    { id: WORKBENCH_PROPOSAL_1, status: "completed" },
    { id: WORKBENCH_PROPOSAL_2, status: "stale", reason: "source_snapshot_changed" },
    { id: WORKBENCH_PROPOSAL_3, status: "deselected", reason: "not_selected" },
  ]);
  const executionOutcomes = [
    { proposalId: WORKBENCH_PROPOSAL_1, status: "completed" },
    { proposalId: WORKBENCH_PROPOSAL_2, status: "stale", reason: "source_snapshot_changed" },
    { proposalId: WORKBENCH_PROPOSAL_3, status: "failed", reason: "execution_failed" },
  ];
  const calls = [];
  const statuses = [];
  let resolveRequest;
  let refreshCount = 0;
  const execute = new Function(
    "financialReconciliationState",
    "api",
    "clean",
    "renderFinancialReconciliationAutomation",
    "setFinancialReconciliationAutomationStatus",
    "loadFinancialReconciliationWorkspace",
    "showToast",
    "financialReconciliationAutomationOutcomeCounts",
    `${appFunctionSource("executeFinancialReconciliationAutomationSelection").replace(/^function /, "async function ")}
     return executeFinancialReconciliationAutomationSelection;`,
  )(
    () => current,
    (url, options) => {
      calls.push({ url, options });
      return new Promise((resolve) => { resolveRequest = () => resolve({ run: refreshedRun, outcomes: executionOutcomes }); });
    },
    (value) => String(value ?? "").trim(),
    () => {},
    (message, tone) => statuses.push({ message, tone }),
    async () => { refreshCount += 1; current.loaded = true; return true; },
    () => {},
    new Function(
      "clean",
      `${appFunctionSource("financialReconciliationAutomationOutcomeCounts")}
       return financialReconciliationAutomationOutcomeCounts;`,
    )((value) => String(value ?? "").trim()),
  );

  const first = execute();
  const duplicate = execute();
  assert.equal(calls.length, 1);
  assert.equal(current.automation.pendingAction, "execute");
  assert.deepEqual(calls[0], {
    url: "/api/reconciliation-automation",
    options: {
      method: "POST",
      body: {
        action: "execute_selected",
        runId: WORKBENCH_RUN_ID,
        proposalIds: [WORKBENCH_PROPOSAL_1, WORKBENCH_PROPOSAL_2, WORKBENCH_PROPOSAL_3],
      },
    },
  });
  resolveRequest();
  await Promise.all([first, duplicate]);

  assert.notEqual(current.automation.run, refreshedRun);
  assert.deepEqual(current.automation.run.proposals, refreshedRun.proposals, "persisted proposal lifecycle remains authoritative");
  assert.deepEqual(current.automation.run.executionOutcomes, executionOutcomes, "attempt outcomes remain separately visible");
  assert.deepEqual([...current.automation.selectedProposalIds], []);
  assert.equal(current.automation.pendingAction, "");
  assert.equal(current.loaded, true, "the workspace loader owns its loaded state after refresh");
  assert.equal(refreshCount, 1);
  assert.match(statuses.at(-1).message, /1 completed, 1 stale, 0 failed/i);
  assert.match(statuses.at(-1).message, /1 execution attempt failure/i);
  assert.equal(statuses.at(-1).tone, "error");

  await execute();
  assert.equal(calls.length, 1);
  assert.match(statuses.at(-1).message, /select at least one executable proposal/i);
});

test("execution reports a shared-history refresh failure instead of a false success", async () => {
  const current = {
    loaded: true,
    automation: {
      run: workbenchRun([{ id: WORKBENCH_PROPOSAL_1, status: "proposed" }]),
      selectedProposalIds: new Set([WORKBENCH_PROPOSAL_1]),
      pendingAction: "",
    },
  };
  const statuses = [];
  const toasts = [];
  const execute = new Function(
    "financialReconciliationState",
    "api",
    "clean",
    "renderFinancialReconciliationAutomation",
    "setFinancialReconciliationAutomationStatus",
    "loadFinancialReconciliationWorkspace",
    "showToast",
    "financialReconciliationAutomationOutcomeCounts",
    `${appFunctionSource("executeFinancialReconciliationAutomationSelection").replace(/^function /, "async function ")}
     return executeFinancialReconciliationAutomationSelection;`,
  )(
    () => current,
    async () => ({ run: workbenchRun([{ id: WORKBENCH_PROPOSAL_1, status: "completed" }]), outcomes: [{ proposalId: WORKBENCH_PROPOSAL_1, status: "completed" }] }),
    (value) => String(value ?? "").trim(),
    () => {},
    (message, tone) => statuses.push({ message, tone }),
    async () => false,
    (message, tone) => toasts.push({ message, tone }),
    new Function(
      "clean",
      `${appFunctionSource("financialReconciliationAutomationOutcomeCounts")}
       return financialReconciliationAutomationOutcomeCounts;`,
    )((value) => String(value ?? "").trim()),
  );

  await execute();

  assert.match(statuses.at(-1).message, /history.*failed/i);
  assert.equal(statuses.at(-1).tone, "error");
  assert.equal(toasts.at(-1).tone, "error");
});

test("history Open from Automatic activates and focuses Manual before loading the selected reconciliation", async () => {
  class FakeElement {
    constructor(dataset) { this.dataset = dataset; }
    closest(selector) { return selector === "[data-financial-reconciliation-select]" ? this : null; }
  }
  const record = { id: "history-1", base_source_type: "financial_documents" };
  const current = {
    activeTab: "automatic",
    selectedReconciliationId: "",
    candidateSourceType: "import_cgd_extrato_ordem",
    workspace: { history: [record] },
    loaded: true,
  };
  const activations = [];
  const openHistory = new Function(
    "HTMLElement",
    "financialReconciliationState",
    "clean",
    "setFinancialReconciliationTab",
    "loadFinancialReconciliationWorkspace",
    `${appFunctionSource("onFinancialReconciliationHistoryClick").replace(/^function /, "async function ")}
     return onFinancialReconciliationHistoryClick;`,
  )(
    FakeElement,
    () => current,
    (value) => String(value ?? "").trim(),
    async (tab, options) => { activations.push({ tab, options }); current.activeTab = tab; },
    async () => {},
  );

  await openHistory({ target: new FakeElement({ financialReconciliationSelect: "history-1" }) });

  assert.deepEqual(activations, [{ tab: "manual", options: { focus: true } }]);
  assert.equal(current.activeTab, "manual");
  assert.equal(current.selectedReconciliationId, "history-1");
});

test("completed, stale, failed, skipped, ambiguous, and deselected outcomes render as separate summaries", () => {
  const resultsMarkup = new Function(
    "clean",
    "escape",
    `${appFunctionSource("financialReconciliationAutomationOutcomeCounts")}
     ${appFunctionSource("financialReconciliationAutomationResultsMarkup")}
     return financialReconciliationAutomationResultsMarkup;`,
  )(
    (value) => String(value ?? "").trim(),
    new Function(`${appFunctionSource("escape")}; return escape;`)(),
  );
  const run = workbenchRun([
    { id: WORKBENCH_PROPOSAL_1, status: "completed" },
    { id: WORKBENCH_PROPOSAL_2, status: "stale", reason: "source_snapshot_changed" },
    { id: WORKBENCH_PROPOSAL_3, status: "deselected", reason: "not_selected" },
    { id: "ambiguous-id", status: "ambiguous", reason: "multiple_combinations" },
    { id: "skipped-id", status: "skipped", reason: "no_qualifying_combination" },
  ]);
  run.executionOutcomes = [
    { proposalId: WORKBENCH_PROPOSAL_1, status: "completed" },
    { proposalId: WORKBENCH_PROPOSAL_2, status: "stale" },
    { proposalId: WORKBENCH_PROPOSAL_3, status: "failed", reason: "execution_failed" },
  ];
  const markup = resultsMarkup(run);

  assert.match(markup, new RegExp(`Run ${WORKBENCH_RUN_ID}[\\s\\S]*ready[\\s\\S]*manual`, "i"));
  assert.match(markup, /financial-reconciliation-automation-result--completed[\s\S]*Completed[\s\S]*1/);
  assert.match(markup, /financial-reconciliation-automation-result--stale[\s\S]*Stale[\s\S]*1/);
  assert.match(markup, /financial-reconciliation-automation-result--failed[\s\S]*Failed[\s\S]*0/);
  assert.match(markup, /financial-reconciliation-automation-result--ambiguous[\s\S]*Ambiguous[\s\S]*1/);
  assert.match(markup, /financial-reconciliation-automation-result--skipped[\s\S]*Skipped[\s\S]*1/);
  assert.match(markup, /financial-reconciliation-automation-result--deselected[\s\S]*Skipped \/ deselected[\s\S]*1/);
  assert.match(markup, /financial-reconciliation-automation-result--attempt-failed[\s\S]*Execution attempt failures[\s\S]*1/);
});

test("persisted completion remains authoritative when the execution response is transport-uncertain", () => {
  const outcomeCounts = new Function(
    "clean",
    `${appFunctionSource("financialReconciliationAutomationOutcomeCounts")}
     return financialReconciliationAutomationOutcomeCounts;`,
  )((value) => String(value ?? "").trim());
  const run = workbenchRun([{ id: WORKBENCH_PROPOSAL_1, status: "completed" }]);
  run.executionOutcomes = [{ proposalId: WORKBENCH_PROPOSAL_1, status: "failed", reason: "execution_failed" }];

  assert.deepEqual(outcomeCounts(run), {
    completed: 1,
    stale: 0,
    failed: 0,
    ambiguous: 0,
    skipped: 0,
    deselected: 0,
    attemptFailures: 1,
  });
});

test("persisted failures are not double-counted as transport-uncertain attempts", () => {
  const outcomeCounts = new Function(
    "clean",
    `${appFunctionSource("financialReconciliationAutomationOutcomeCounts")}
     return financialReconciliationAutomationOutcomeCounts;`,
  )((value) => String(value ?? "").trim());
  const run = workbenchRun([{ id: WORKBENCH_PROPOSAL_1, status: "failed", reason: "execution_failed" }]);
  run.executionOutcomes = [{ proposalId: WORKBENCH_PROPOSAL_1, status: "failed", reason: "execution_failed" }];

  assert.deepEqual(outcomeCounts(run), {
    completed: 0,
    stale: 0,
    failed: 1,
    ambiguous: 0,
    skipped: 0,
    deselected: 0,
    attemptFailures: 0,
  });
});
