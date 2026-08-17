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
    {
      ruleKey: "financial_documents_cgd_bank_statement_amount_only",
      ruleVersion: 1,
      displayName: "Bank statement amount-only <rule>",
      baseSourceType: "financial_documents",
      destinationSourceTypes: ["import_cgd_extrato_ordem"],
      logicDescription: "Amount-only bank statement rule",
      definition: { matching: "amount-only" },
      enabled: false,
      allowManualExecution: true,
      includeInScheduledBatch: false,
      differenceAllowed: "0.00",
      maxDifferenceDays: 2,
      priority: 3,
    },
    {
      ruleKey: "financial_documents_cgd_credit_card_amount_only",
      ruleVersion: 1,
      displayName: "Credit card amount-only rule",
      baseSourceType: "financial_documents",
      destinationSourceTypes: ["import_fdm_accounts"],
      logicDescription: "Amount-only credit card rule",
      definition: { matching: "amount-only" },
      enabled: false,
      allowManualExecution: true,
      includeInScheduledBatch: false,
      differenceAllowed: "0.00",
      maxDifferenceDays: 3,
      priority: 4,
    },
  ];
}

function isAmountOnlyRuleKey(ruleKey) {
  return [
    "financial_documents_cgd_bank_statement_amount_only",
    "financial_documents_cgd_credit_card_amount_only",
  ].includes(String(ruleKey ?? "").trim());
}

function automationSettings(overrides = {}) {
  return {
    loaded: true,
    loading: false,
    activeTab: "automatic",
    schedule: { enabled: true, timeOfDay: "02:15", timeZone: "Europe/Lisbon" },
    rules: managedRules(),
    lastScheduledBatch: null,
    ...overrides,
  };
}

function compilePayload(state) {
  return new Function(
    "state",
    "isReconciliationAutomationAmountOnlyRule",
    `${appFunctionSource("clean")}
     ${appFunctionSource("reconciliationAutomationSettingsPayload")}
     return reconciliationAutomationSettingsPayload;`,
  )(state, isAmountOnlyRuleKey);
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
    financialReconciliationAutomationOpenWorkbench: { disabled: false },
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
    "isReconciliationAutomationAmountOnlyRule",
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
    isAmountOnlyRuleKey,
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
    "financial-reconciliation-automation-open-workbench",
  ]) {
    assert.match(html, new RegExp(`id="${id}"`));
  }
  assert.match(html, /role="tablist"[\s\S]*aria-controls="financial-reconciliation-settings-source-panel"[\s\S]*aria-controls="financial-reconciliation-settings-automatic-panel"/);
  assert.match(html, /id="financial-reconciliation-automation-time-zone"[^>]*>Europe\/Lisbon</);
  assert.match(html, /id="financial-reconciliation-automation-schedule-time"[^>]*aria-describedby="financial-reconciliation-automation-status"/);
  assert.match(html, /id="financial-reconciliation-automation-open-workbench"[^>]*>Open automatic reconciliation<\/button>/);
  assert.doesNotMatch(html, /financial-reconciliation-automation-run-batch-now|>Run batch now<\/button>/);
  assert.match(html, /id="financial-reconciliation-workbench-automation-rule"[^>]*aria-describedby="financial-reconciliation-workbench-automation-status"/);
  assert.match(html, /id="financial-reconciliation-workbench-automation-analyze"[^>]*>Analyze<\/button>/);
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

test("amount-only Settings cards render a fixed zero while their other controls stay editable", () => {
  const markup = renderAutomationSettings(automationSettings()).financialReconciliationAutomationRules.innerHTML;
  const amountOnlyKeys = [
    "financial_documents_cgd_bank_statement_amount_only",
    "financial_documents_cgd_credit_card_amount_only",
  ];

  assert.equal((markup.match(/data-reconciliation-automation-rule-field="differenceAllowed"/g) || []).length, 2);
  assert.equal((markup.match(/financial-reconciliation-automation-fixed-value/g) || []).length, 2);
  for (const key of ["rule-a", "rule-b"]) {
    const card = markup.match(new RegExp(`<article[^>]*data-reconciliation-automation-rule-card="${key}"[\\s\\S]*?<\\/article>`))?.[0] || "";
    assert.match(card, /<input type="number"[^>]*data-reconciliation-automation-rule-field="differenceAllowed"/);
    assert.doesNotMatch(card, /financial-reconciliation-automation-fixed-value/);
  }
  for (const key of amountOnlyKeys) {
    const card = markup.match(new RegExp(`<article[^>]*data-reconciliation-automation-rule-card="${key}"[\\s\\S]*?<\\/article>`))?.[0] || "";
    assert.match(card, /<output class="financial-reconciliation-automation-fixed-value"[^>]*aria-label="Difference allowed, fixed"[^>]*>0\.00 €<\/output>/);
    assert.doesNotMatch(card, /data-reconciliation-automation-rule-field="differenceAllowed"/);
    for (const field of ["enabled", "allowManualExecution", "includeInScheduledBatch", "maxDifferenceDays"]) {
      assert.match(card, new RegExp(`data-reconciliation-automation-rule-field="${field}"`));
      assert.doesNotMatch(card.match(new RegExp(`data-reconciliation-automation-rule-field="${field}"[^>]*`, "g"))?.join("") || "", /disabled/);
    }
  }
});

test("server-loaded amount-only tolerances canonicalize to zero without changing identity tolerances", () => {
  const rules = managedRules();
  const result = {
    schedule: {
      enabled: true,
      timeOfDay: "04:45",
      timeZone: "Europe/Lisbon",
      updatedBy: "admin@example.com",
      updatedAt: "2026-08-17T10:00:00.000Z",
    },
    rules: [
      {
        ...rules[0],
        ruleKey: "financial_documents_cgd_bank_statement",
        ruleVersion: 2,
        displayName: "Financial Documents to CGD Bank Statement",
        differenceAllowed: 7.65,
        priority: 1,
        updatedBy: "admin@example.com",
        updatedAt: "2026-08-17T10:00:01.000Z",
      },
      {
        ...rules[1],
        ruleKey: "financial_documents_cgd_credit_card",
        ruleVersion: 1,
        displayName: "Financial Documents to CGD Credit Card",
        destinationSourceTypes: ["import_cgd_cartao_credito"],
        differenceAllowed: 4.32,
        priority: 2,
        updatedBy: "admin@example.com",
        updatedAt: "2026-08-17T10:00:02.000Z",
      },
      {
        ...rules[2],
        differenceAllowed: 9.87,
        priority: 3,
        updatedBy: "admin@example.com",
        updatedAt: "2026-08-17T10:00:03.000Z",
      },
      {
        ...rules[3],
        destinationSourceTypes: ["import_cgd_cartao_credito"],
        differenceAllowed: 6.54,
        priority: 4,
        updatedBy: "admin@example.com",
        updatedAt: "2026-08-17T10:00:04.000Z",
      },
    ],
    lastScheduledBatch: {
      id: "00000000-0000-0000-0000-000000000777",
      scheduledSlot: "2026-08-17",
      status: "completed",
      counts: {
        ruleCount: 4,
        childCount: 4,
        completedChildren: 4,
        partialChildren: 0,
        failedChildren: 0,
        unfinishedChildren: 0,
      },
      ruleCount: 4,
      childCount: 4,
      startedAt: "2026-08-17T04:45:00.000Z",
      finishedAt: "2026-08-17T04:49:00.000Z",
      updatedAt: "2026-08-17T04:49:00.000Z",
    },
  };
  const state = {
    reconciliationAutomationSettings: automationSettings({ loaded: false, rules: [] }),
    reconciliationAutomationSettingsDirty: true,
  };
  const applyResult = new Function(
    "state",
    "clean",
    "clone",
    "isReconciliationAutomationAmountOnlyRule",
    `${appFunctionSource("applyReconciliationAutomationSettingsResult")}
     return applyReconciliationAutomationSettingsResult;`,
  )(
    state,
    new Function(`${appFunctionSource("clean")}; return clean;`)(),
    new Function(`${appFunctionSource("clone")}; return clone;`)(),
    isAmountOnlyRuleKey,
  );

  applyResult(result);

  const loadedDifference = (ruleKey) => state.reconciliationAutomationSettings.rules
    .find((rule) => rule.ruleKey === ruleKey)?.differenceAllowed;
  assert.equal(loadedDifference("financial_documents_cgd_bank_statement"), "7.65");
  assert.equal(loadedDifference("financial_documents_cgd_credit_card"), "4.32");
  assert.equal(loadedDifference("financial_documents_cgd_bank_statement_amount_only"), "0.00");
  assert.equal(loadedDifference("financial_documents_cgd_credit_card_amount_only"), "0.00");
});

test("actual renderer summarizes the last scheduled batch without presenting one child run", () => {
  const els = renderAutomationSettings(automationSettings({
    lastScheduledBatch: {
      batchId: "00000000-0000-0000-0000-000000000777",
      status: "partial",
      startedAt: "2026-08-14T02:00:00Z",
      finishedAt: "2026-08-14T02:01:00Z",
      counts: { rules: 2, completed: 1, failed: 1 },
      childRunId: "must-not-render",
    },
  }));

  assert.equal(els.financialReconciliationAutomationLastExecution.textContent, "Last batch: 2026-08-14 03:01 Europe/Lisbon · partial");
  assert.equal(els.financialReconciliationAutomationLastResult.textContent, "Batch result: rules 2 · completed 1 · failed 1");
  assert.doesNotMatch(`${els.financialReconciliationAutomationLastExecution.textContent} ${els.financialReconciliationAutomationLastResult.textContent}`, /must-not-render/);
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
      {
        ruleKey: "financial_documents_cgd_bank_statement_amount_only",
        ruleVersion: 1,
        enabled: false,
        allowManualExecution: true,
        includeInScheduledBatch: false,
        differenceAllowed: "0.00",
        maxDifferenceDays: 2,
        priority: 3,
      },
      {
        ruleKey: "financial_documents_cgd_credit_card_amount_only",
        ruleVersion: 1,
        enabled: false,
        allowManualExecution: true,
        includeInScheduledBatch: false,
        differenceAllowed: "0.00",
        maxDifferenceDays: 3,
        priority: 4,
      },
    ],
  });
});

test("amount-only settings serialize authoritative zero despite a tampered local tolerance", () => {
  const settings = automationSettings();
  for (const rule of settings.rules.filter((item) => isAmountOnlyRuleKey(item.ruleKey))) rule.differenceAllowed = "9.99";

  const payload = compilePayload({ reconciliationAutomationSettings: settings })();

  assert.ok(payload);
  assert.deepEqual(
    payload.rules.filter((rule) => isAmountOnlyRuleKey(rule.ruleKey)).map((rule) => rule.differenceAllowed),
    ["0.00", "0.00"],
  );
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
    [
      { ruleKey: "rule-b", priority: 1 },
      { ruleKey: "rule-a", priority: 2 },
      { ruleKey: "financial_documents_cgd_bank_statement_amount_only", priority: 3 },
      { ruleKey: "financial_documents_cgd_credit_card_amount_only", priority: 4 },
    ],
  );
  assert.equal(state.reconciliationAutomationSettingsDirty, true);
  assert.deepEqual(focusCalls[0], ["rule-b", "up"]);
  move("financial_documents_cgd_bank_statement_amount_only", "up");
  assert.equal(
    state.reconciliationAutomationSettings.rules.find((rule) => rule.ruleKey === "financial_documents_cgd_bank_statement_amount_only").priority,
    2,
  );
  move("rule-b", "up");
  assert.deepEqual(state.reconciliationAutomationSettings.rules.map((rule) => rule.priority), [1, 2, 3, 4]);
});

test("invalid local automation values disable Save", () => {
  for (const invalid of [
    automationSettings({ schedule: { enabled: true, timeOfDay: "25:00", timeZone: "Europe/Lisbon" } }),
    automationSettings({ rules: [{ ...managedRules()[1], differenceAllowed: "-0.01" }] }),
    automationSettings({ rules: [{ ...managedRules()[1], differenceAllowed: "1.234" }] }),
    automationSettings({ rules: [{ ...managedRules()[1], differenceAllowed: "90071992547410.00" }] }),
    automationSettings({ rules: [{ ...managedRules()[1], maxDifferenceDays: "" }] }),
    automationSettings({ rules: [{ ...managedRules()[1], maxDifferenceDays: 91 }] }),
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

test("Open automatic reconciliation requires app access but ignores unsaved Settings drafts", () => {
  const dirtyEls = renderAutomationSettings(automationSettings(), { dirty: true });
  assert.equal(dirtyEls.financialReconciliationAutomationOpenWorkbench.disabled, false);
  assert.match(dirtyEls.financialReconciliationAutomationRunHint.textContent, /saved manual-enabled rules/i);
  assert.match(dirtyEls.financialReconciliationAutomationRunHint.textContent, /unsaved changes are not applied/i);

  const noAccessEls = renderAutomationSettings(automationSettings(), { canOpenWorkbench: false });
  assert.equal(noAccessEls.financialReconciliationAutomationOpenWorkbench.disabled, true);
  assert.match(noAccessEls.financialReconciliationAutomationRunHint.textContent, /reconciliation app access/i);
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
    "isReconciliationAutomationAmountOnlyRule",
    "updateReconciliationAutomationNextExecution",
    "updateReconciliationAutomationControls",
    `${appFunctionSource("onReconciliationAutomationSettingsInput")}
     return onReconciliationAutomationSettingsInput;`,
  )(FakeHTMLElement, state, (value) => String(value ?? "").trim(), isAmountOnlyRuleKey, () => {}, () => { validationUpdates += 1; });
  input({ target: new FakeHTMLElement({ reconciliationAutomationRuleKey: "rule-a", reconciliationAutomationRuleField: "differenceAllowed" }, { value: "3.25" }) });
  assert.equal(settings.rules.find((rule) => rule.ruleKey === "rule-a").differenceAllowed, "3.25");
  assert.equal(state.reconciliationAutomationSettingsDirty, true);
  assert.equal(validationUpdates, 1);
  assert.deepEqual(requests, []);

  const amountOnlyRule = settings.rules.find((rule) => isAmountOnlyRuleKey(rule.ruleKey));
  input({ target: new FakeHTMLElement({ reconciliationAutomationRuleKey: amountOnlyRule.ruleKey, reconciliationAutomationRuleField: "differenceAllowed" }, { value: "4.50" }) });
  assert.equal(amountOnlyRule.differenceAllowed, "0.00");
  assert.equal(validationUpdates, 1);

  input({ target: new FakeHTMLElement({ reconciliationAutomationRuleKey: amountOnlyRule.ruleKey, reconciliationAutomationRuleField: "maxDifferenceDays" }, { value: "8" }) });
  assert.equal(amountOnlyRule.maxDifferenceDays, "8");
  assert.equal(validationUpdates, 2);

  for (const [field, checked] of [["enabled", true], ["allowManualExecution", false], ["includeInScheduledBatch", true]]) {
    input({ target: new FakeHTMLElement({ reconciliationAutomationRuleKey: amountOnlyRule.ruleKey, reconciliationAutomationRuleField: field }, { checked }) });
    assert.equal(amountOnlyRule[field], checked);
  }
  assert.equal(validationUpdates, 5);

  const payload = compilePayload(state);
  const statuses = [];
  const applyResult = new Function(
    "state",
    "clean",
    "clone",
    "isReconciliationAutomationAmountOnlyRule",
    `${appFunctionSource("applyReconciliationAutomationSettingsResult")}
     return applyReconciliationAutomationSettingsResult;`,
  )(
    state,
    new Function(`${appFunctionSource("clean")}; return clean;`)(),
    new Function(`${appFunctionSource("clone")}; return clone;`)(),
    isAmountOnlyRuleKey,
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
        lastScheduledBatch: null,
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
  assert.deepEqual(
    requests[0].options.body.rules.filter((rule) => isAmountOnlyRuleKey(rule.ruleKey)).map((rule) => rule.differenceAllowed),
    ["0.00", "0.00"],
  );
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

test("Open automatic reconciliation only navigates to the Automatic tab and never leaks the Settings draft", async () => {
  const calls = [];
  const statuses = [];
  const openWorkbench = new Function(
    "canAppFinancialReconciliation",
    "setReconciliationAutomationSettingsStatus",
    "setView",
    `${appFunctionSource("openFinancialReconciliationAutomation").replace(/^function /, "async function ")}
     return openFinancialReconciliationAutomation;`,
  )(
    () => true,
    (message, isError) => statuses.push({ message, isError }),
    async (view, options) => calls.push({ view, options }),
  );

  await openWorkbench();

  assert.deepEqual(calls, [{
    view: "financial-reconciliation",
    options: { financialReconciliationTab: "automatic" },
  }]);
  assert.deepEqual(statuses, []);
  assert.doesNotMatch(appFunctionSource("openFinancialReconciliationAutomation"), /api\(|analyze_|execute_|proposal|reconciliationAutomationSettingsPayload/);
});

test("Open automatic reconciliation fails closed without Reconciliation app access", async () => {
  const calls = [];
  const statuses = [];
  const openWorkbench = new Function(
    "canAppFinancialReconciliation",
    "setReconciliationAutomationSettingsStatus",
    "setView",
    `${appFunctionSource("openFinancialReconciliationAutomation").replace(/^function /, "async function ")}
     return openFinancialReconciliationAutomation;`,
  )(
    () => false,
    (message, isError) => statuses.push({ message, isError }),
    async (...args) => calls.push(args),
  );

  await openWorkbench();

  assert.deepEqual(calls, []);
  assert.deepEqual(statuses, [{ message: "Reconciliation app access is required.", isError: true }]);
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
    analysisCompletedAt: "2026-08-16T09:00:00.000Z",
    analysisProcessed: 1,
    analysisTotal: 1,
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
    `${appFunctionSource("financialReconciliationAutomationIsAnalyzing")}
     ${appFunctionSource("financialReconciliationAutomationVisibleProposals")}
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

function renderAutomationWorkbench(run, selectedProposalIds = new Set(), { continuationRetry = false, rules = workbenchRules(), selectedRuleKey = "manual-enabled", pendingAction = "" } = {}) {
  const state = { automation: { rules, run, selectedRuleKey, selectedProposalIds, pendingAction, loaded: true, continuationRetry } };
  const proposalContainer = { innerHTML: "", querySelectorAll: () => [] };
  const els = {
    financialReconciliationWorkbenchAutomationRule: { innerHTML: "", value: "", disabled: false },
    financialReconciliationWorkbenchAutomationAnalyze: { disabled: false, textContent: "" },
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
    "financialReconciliationAutomationRuleOptions",
    "financialReconciliationAutomationProposalMarkup",
    "financialReconciliationAutomationResultsMarkup",
    "escape",
    `${appFunctionSource("financialReconciliationAutomationOpenRun")}
     ${appFunctionSource("financialReconciliationAutomationIsAnalyzing")}
     ${appFunctionSource("financialReconciliationAutomationProgressLabel")}
     ${appFunctionSource("financialReconciliationAutomationVisibleProposals")}
     ${appFunctionSource("financialReconciliationAutomationEmptyMessage")}
     ${appFunctionSource("renderFinancialReconciliationAutomation")}
     return renderFinancialReconciliationAutomation;`,
  )(
    () => state,
    (value) => String(value ?? "").trim(),
    els,
    new Function(
      "clean",
      "escape",
      `${appFunctionSource("financialReconciliationAutomationRuleOptions")}
       return financialReconciliationAutomationRuleOptions;`,
    )(
      (value) => String(value ?? "").trim(),
      new Function(`${appFunctionSource("escape")}; return escape;`)(),
    ),
    (proposal) => `<article data-proposal-id="${proposal.id}">${proposal.id}</article>`,
    () => "",
    (value) => String(value ?? "").replace(/[&<>"']/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[character]),
  );
  render();
  return { els, state };
}

test("analysis progress renders processed and total records while review controls stay disabled", () => {
  const progressLabel = new Function(
    `${appFunctionSource("financialReconciliationAutomationProgressLabel")}
     return financialReconciliationAutomationProgressLabel;`,
  )();
  const run = {
    ...workbenchRun([{ id: WORKBENCH_PROPOSAL_1, status: "proposed" }]),
    status: "analyzing",
    analysisCompletedAt: null,
    analysisProcessed: 25,
    analysisTotal: 876,
  };
  const { els } = renderAutomationWorkbench(run, new Set([WORKBENCH_PROPOSAL_1]));

  assert.match(progressLabel(run), /Analyzing 25 of 876 records/i);
  assert.match(els.financialReconciliationWorkbenchAutomationProposals.innerHTML, /Analyzing 25 of 876 records/i);
  assert.equal(els.financialReconciliationWorkbenchAutomationSelectAll.disabled, true);
  assert.equal(els.financialReconciliationWorkbenchAutomationClearAll.disabled, true);
  assert.equal(els.financialReconciliationWorkbenchAutomationExecute.disabled, true);
});

test("serial continuation replaces progress and selects proposals only after analysis completes", async () => {
  const current = {
    automation: {
      run: {
        runId: WORKBENCH_RUN_ID,
        status: "analyzing",
        analysisCompletedAt: null,
        analysisProcessed: 0,
        analysisTotal: 50,
        proposals: [],
      },
      selectedProposalIds: new Set(),
      pendingAction: "",
      continuationToken: 1,
      continuationRetry: false,
    },
  };
  const calls = [];
  const responses = [
    { ...current.automation.run, analysisProcessed: 25 },
    workbenchRun([{ id: WORKBENCH_PROPOSAL_1, status: "proposed" }]),
  ];
  const statuses = [];
  const continueAnalysis = new Function(
    "financialReconciliationState",
    "api",
    "clean",
    "renderFinancialReconciliationAutomation",
    "setFinancialReconciliationAutomationStatus",
    "showToast",
    "finalizeFinancialReconciliationAutomationAnalysis",
    `${appFunctionSource("financialReconciliationAutomationIsAnalyzing")}
     ${appFunctionSource("financialReconciliationAutomationProgressLabel")}
     ${appFunctionSource("finalizeFinancialReconciliationAutomationAnalysis")}
     ${appFunctionSource("continueFinancialReconciliationAutomationAnalysis").replace(/^function /, "async function ")}
     return continueFinancialReconciliationAutomationAnalysis;`,
  )(
    () => current,
    async (url, options) => { calls.push({ url, options }); return responses.shift(); },
    (value) => String(value ?? "").trim(),
    () => {},
    (message, tone) => statuses.push({ message, tone }),
    () => {},
  );

  await continueAnalysis(1);

  assert.deepEqual(calls, [1, 2].map(() => ({
    url: "/api/reconciliation-automation",
    options: { method: "POST", body: { action: "continue_analysis", runId: WORKBENCH_RUN_ID } },
  })));
  assert.deepEqual([...current.automation.selectedProposalIds], [WORKBENCH_PROPOSAL_1]);
  assert.equal(current.automation.pendingAction, "");
  assert.equal(current.automation.continuationRetry, false);
  assert.match(statuses.at(-1).message, /Analysis ready: 1 executable proposal/i);
});

test("automatic tab restores an unfinished actor run before continuing it", async () => {
  const current = {
    automation: {
      rules: [
        ...workbenchRules(),
        { ...workbenchRules()[0], ruleKey: "other-manual", displayName: "Other manual", priority: 4 },
      ],
      run: null,
      selectedRuleKey: "other-manual",
      selectedProposalIds: new Set(["old"]),
      pendingAction: "",
      continuationToken: 0,
      continuationRetry: true,
    },
  };
  const activeRun = {
    runId: WORKBENCH_RUN_ID,
    status: "analyzing",
    analysisCompletedAt: null,
    analysisProcessed: 25,
    analysisTotal: 75,
    definitions: [{ ruleKey: "manual-enabled", ruleVersion: 3 }],
    proposals: [],
  };
  const calls = [];
  const continuedTokens = [];
  const restore = new Function(
    "financialReconciliationState",
    "api",
    "clean",
    "renderFinancialReconciliationAutomation",
    "setFinancialReconciliationAutomationStatus",
    "continueFinancialReconciliationAutomationAnalysis",
    `${appFunctionSource("financialReconciliationAutomationIsAnalyzing")}
     ${appFunctionSource("restoreFinancialReconciliationAutomationAnalysis").replace(/^function /, "async function ")}
     return restoreFinancialReconciliationAutomationAnalysis;`,
  )(
    () => current,
    async (url) => { calls.push(url); return activeRun; },
    (value) => String(value ?? "").trim(),
    () => {},
    () => {},
    async (token) => { continuedTokens.push(token); },
  );

  await restore();

  assert.deepEqual(calls, ["/api/reconciliation-automation?view=active_run"]);
  assert.strictEqual(current.automation.run, activeRun);
  assert.equal(current.automation.selectedRuleKey, "manual-enabled");
  assert.deepEqual([...current.automation.selectedProposalIds], []);
  assert.equal(current.automation.continuationRetry, false);
  assert.deepEqual(continuedTokens, [1]);
});

test("automatic tab restores a ready manual run for proposal review without continuing it", async () => {
  const readyRun = workbenchRun([{ id: WORKBENCH_PROPOSAL_1, status: "proposed" }]);
  const current = {
    automation: {
      run: null,
      selectedProposalIds: new Set(),
      pendingAction: "",
      continuationToken: 0,
      continuationRetry: false,
    },
  };
  const continuedTokens = [];
  const statuses = [];
  const restore = new Function(
    "financialReconciliationState",
    "api",
    "clean",
    "renderFinancialReconciliationAutomation",
    "setFinancialReconciliationAutomationStatus",
    "continueFinancialReconciliationAutomationAnalysis",
    "finalizeFinancialReconciliationAutomationAnalysis",
    `${appFunctionSource("financialReconciliationAutomationIsAnalyzing")}
     ${appFunctionSource("restoreFinancialReconciliationAutomationAnalysis").replace(/^function /, "async function ")}
     return restoreFinancialReconciliationAutomationAnalysis;`,
  )(
    () => current,
    async () => readyRun,
    (value) => String(value ?? "").trim(),
    () => {},
    (message, tone) => statuses.push({ message, tone }),
    async (token) => { continuedTokens.push(token); },
    () => {
      current.automation.selectedProposalIds = new Set([WORKBENCH_PROPOSAL_1]);
      statuses.push({ message: "Analysis ready", tone: "success" });
    },
  );

  await restore();

  assert.strictEqual(current.automation.run, readyRun);
  assert.deepEqual(continuedTokens, []);
  assert.deepEqual([...current.automation.selectedProposalIds], [WORKBENCH_PROPOSAL_1]);
  assert.match(statuses.at(-1).message, /Analysis ready/i);
});

test("failed initial restore exposes a retry that retries the active-run lookup", async () => {
  const current = {
    automation: {
      run: null,
      selectedProposalIds: new Set(),
      pendingAction: "",
      continuationToken: 0,
      continuationRetry: false,
    },
  };
  const restore = new Function(
    "financialReconciliationState",
    "api",
    "clean",
    "renderFinancialReconciliationAutomation",
    "setFinancialReconciliationAutomationStatus",
    "continueFinancialReconciliationAutomationAnalysis",
    "finalizeFinancialReconciliationAutomationAnalysis",
    `${appFunctionSource("financialReconciliationAutomationIsAnalyzing")}
     ${appFunctionSource("restoreFinancialReconciliationAutomationAnalysis").replace(/^function /, "async function ")}
     return restoreFinancialReconciliationAutomationAnalysis;`,
  )(
    () => current,
    async () => { throw new Error("network unavailable"); },
    (value) => String(value ?? "").trim(),
    () => {},
    () => {},
    async () => {},
    () => {},
  );

  await restore();

  assert.equal(current.automation.continuationRetry, true);
  const { els } = renderAutomationWorkbench(null, new Set(), { continuationRetry: true });
  assert.equal((els.financialReconciliationWorkbenchAutomationProposals.innerHTML.match(/data-financial-reconciliation-automation-retry/g) || []).length, 1);
});

test("active-run restoration retry locks Analyze until the authoritative lookup settles", async () => {
  const current = {
    automation: {
      rules: workbenchRules(),
      run: null,
      selectedRuleKey: "manual-enabled",
      selectedProposalIds: new Set(),
      pendingAction: "",
      loaded: true,
      continuationToken: 0,
      continuationRetry: true,
    },
  };
  let resolveLookup;
  const lookup = new Promise((resolve) => { resolveLookup = resolve; });
  const renders = [];
  const retry = new Function(
    "financialReconciliationState",
    "api",
    "clean",
    "renderFinancialReconciliationAutomation",
    "setFinancialReconciliationAutomationStatus",
    "continueFinancialReconciliationAutomationAnalysis",
    "finalizeFinancialReconciliationAutomationAnalysis",
    `${appFunctionSource("financialReconciliationAutomationIsAnalyzing")}
     ${appFunctionSource("restoreFinancialReconciliationAutomationAnalysis").replace(/^function /, "async function ")}
     ${appFunctionSource("retryFinancialReconciliationAutomationAnalysis")}
     return retryFinancialReconciliationAutomationAnalysis;`,
  )(
    () => current,
    async () => lookup,
    (value) => String(value ?? "").trim(),
    () => renders.push(current.automation.pendingAction),
    () => {},
    async () => {},
    () => {},
  );

  retry();

  assert.equal(current.automation.continuationRetry, false);
  assert.equal(current.automation.pendingAction, "restore");
  const pending = renderAutomationWorkbench(null, new Set(), { pendingAction: current.automation.pendingAction });
  assert.equal(pending.els.financialReconciliationWorkbenchAutomationRule.disabled, true);
  assert.equal(pending.els.financialReconciliationWorkbenchAutomationAnalyze.disabled, true);
  assert.deepEqual(renders, ["restore"]);

  resolveLookup(null);
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(current.automation.pendingAction, "");
  assert.deepEqual(renders, ["restore", ""]);
});

test("analysis completion timestamp is the authoritative review boundary", () => {
  const isAnalyzing = new Function(
    "clean",
    `${appFunctionSource("financialReconciliationAutomationIsAnalyzing")}
     return financialReconciliationAutomationIsAnalyzing;`,
  )((value) => String(value ?? "").trim());

  assert.equal(isAnalyzing({ status: "ready", analysisCompletedAt: null, finishedAt: null }), true);
  assert.equal(isAnalyzing({ status: "analyzing", analysisCompletedAt: "2026-08-16T10:00:00.000Z", finishedAt: null }), false);
  assert.equal(isAnalyzing({ status: "failed", analysisCompletedAt: null, finishedAt: "2026-08-16T10:00:01.000Z" }), false);
  assert.match(appFunctionSource("analyzeFinancialReconciliationAutomationRule"), /if \(financialReconciliationAutomationIsAnalyzing\(run\)\)/);
});

test("uncertain continuation reloads persisted progress and exposes one retry action", async () => {
  const persistedRun = {
    runId: WORKBENCH_RUN_ID,
    status: "analyzing",
    analysisCompletedAt: null,
    analysisProcessed: 25,
    analysisTotal: 75,
    proposals: [],
  };
  const current = {
    automation: {
      run: { ...persistedRun, analysisProcessed: 0 },
      selectedProposalIds: new Set(),
      pendingAction: "",
      continuationToken: 2,
      continuationRetry: false,
    },
  };
  const calls = [];
  const statuses = [];
  const continueAnalysis = new Function(
    "financialReconciliationState",
    "api",
    "clean",
    "renderFinancialReconciliationAutomation",
    "setFinancialReconciliationAutomationStatus",
    "showToast",
    `${appFunctionSource("financialReconciliationAutomationIsAnalyzing")}
     ${appFunctionSource("financialReconciliationAutomationProgressLabel")}
     ${appFunctionSource("finalizeFinancialReconciliationAutomationAnalysis")}
     ${appFunctionSource("continueFinancialReconciliationAutomationAnalysis").replace(/^function /, "async function ")}
     return continueFinancialReconciliationAutomationAnalysis;`,
  )(
    () => current,
    async (url) => {
      calls.push(url);
      if (url.endsWith("view=active_run")) return persistedRun;
      throw new Error("connection lost");
    },
    (value) => String(value ?? "").trim(),
    () => {},
    (message, tone) => statuses.push({ message, tone }),
    () => {},
  );

  await continueAnalysis(2);

  assert.deepEqual(calls, [
    "/api/reconciliation-automation",
    "/api/reconciliation-automation?view=active_run",
  ]);
  assert.strictEqual(current.automation.run, persistedRun);
  assert.equal(current.automation.pendingAction, "");
  assert.equal(current.automation.continuationRetry, true);
  assert.match(statuses.at(-1).message, /Analysis paused after a connection error/i);
  const { els } = renderAutomationWorkbench(current.automation.run, new Set(), { continuationRetry: true });
  assert.equal((els.financialReconciliationWorkbenchAutomationProposals.innerHTML.match(/data-financial-reconciliation-automation-retry/g) || []).length, 1);
  assert.match(els.financialReconciliationWorkbenchAutomationProposals.innerHTML, /Analyzing 25 of 75 records/i);
});

test("active automation runs show proposed and ambiguous rows only", () => {
  const visible = compileVisibleAutomationProposals();
  const proposals = [
    { id: "checked", status: "proposed" },
    { id: "unchecked", status: "proposed" },
    { id: "ambiguous", status: "ambiguous" },
    { id: "skipped", status: "skipped" },
    { id: "completed", status: "completed" },
  ];
  const run = Object.freeze({ analysisCompletedAt: "2026-08-16T09:00:00.000Z", finishedAt: null, proposals: Object.freeze(proposals) });

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

test("workbench rule LOV filters, sorts, escapes, and preserves only a valid selection", () => {
  const ruleOptions = new Function(
    "clean",
    "escape",
    `${appFunctionSource("financialReconciliationAutomationRuleOptions")}
     return financialReconciliationAutomationRuleOptions;`,
  )(
    (value) => String(value ?? "").trim(),
    new Function(`${appFunctionSource("escape")}; return escape;`)(),
  );
  const eligibleLater = {
    ...workbenchRules()[0],
    ruleKey: "manual-&-later",
    displayName: "<Unsafe & later>",
    priority: 9,
  };
  const eligibleFirst = {
    ...workbenchRules()[0],
    ruleKey: 'manual-"-first',
    displayName: "First eligible",
    priority: 0,
  };
  const rules = [eligibleLater, ...workbenchRules(), eligibleFirst];

  const retainedMarkup = ruleOptions(rules, "manual-&-later");
  assert.equal((retainedMarkup.match(/<option /g) || []).length, 3);
  assert.ok(retainedMarkup.indexOf("First eligible") < retainedMarkup.indexOf("Manual enabled"));
  assert.ok(retainedMarkup.indexOf("Manual enabled") < retainedMarkup.indexOf("&lt;Unsafe &amp; later&gt;"));
  assert.match(retainedMarkup, /value="manual-&amp;-later"[^>]*selected[^>]*>&lt;Unsafe &amp; later&gt;<\/option>/);
  assert.doesNotMatch(retainedMarkup, /Disabled rule|Scheduled only/);

  const defaultMarkup = ruleOptions(rules, "missing-rule");
  assert.match(defaultMarkup, /value="manual-&quot;-first"[^>]*selected[^>]*>First eligible<\/option>/);
  assert.doesNotMatch(defaultMarkup, /value="manual-enabled"[^>]*selected/);
});

test("enabled manual amount-only rules appear escaped in the workbench selector", () => {
  const amountOnly = {
    ...managedRules()[2],
    enabled: true,
    allowManualExecution: true,
    displayName: "<Amount-only & manual>",
  };
  const { els } = renderAutomationWorkbench(null, new Set(), {
    rules: [amountOnly],
    selectedRuleKey: amountOnly.ruleKey,
  });

  assert.match(
    els.financialReconciliationWorkbenchAutomationRule.innerHTML,
    /value="financial_documents_cgd_bank_statement_amount_only" selected>&lt;Amount-only &amp; manual&gt;<\/option>/,
  );
  assert.equal(els.financialReconciliationWorkbenchAutomationRule.disabled, false);
});

test("workbench rule change keeps the user selection and refuses changes while a run is open", () => {
  const otherRule = { ...workbenchRules()[0], ruleKey: "other-manual", displayName: "Other manual", priority: 2 };
  const current = {
    automation: {
      rules: [workbenchRules()[0], otherRule],
      run: null,
      selectedRuleKey: "manual-enabled",
      pendingAction: "",
    },
  };
  let renders = 0;
  const onChange = new Function(
    "financialReconciliationState",
    "clean",
    "renderFinancialReconciliationAutomation",
    `${appFunctionSource("financialReconciliationAutomationOpenRun")}
     ${appFunctionSource("onFinancialReconciliationAutomationRuleChange")}
     return onFinancialReconciliationAutomationRuleChange;`,
  )(
    () => current,
    (value) => String(value ?? "").trim(),
    () => { renders += 1; },
  );

  onChange({ target: { value: "other-manual" } });
  assert.equal(current.automation.selectedRuleKey, "other-manual");

  current.automation.run = { ...workbenchRun([]), status: "ready", finishedAt: null };
  onChange({ target: { value: "manual-enabled" } });
  assert.equal(current.automation.selectedRuleKey, "other-manual");
  assert.equal(renders, 2);
});

test("workbench selector locks only for unfinished analyzing, ready, or running runs", () => {
  const openRun = new Function(
    "clean",
    `${appFunctionSource("financialReconciliationAutomationOpenRun")}
     return financialReconciliationAutomationOpenRun;`,
  )((value) => String(value ?? "").trim());

  for (const status of ["analyzing", "ready", "running"]) {
    assert.equal(openRun({ status, finishedAt: null }), true, status);
    const { els } = renderAutomationWorkbench({ ...workbenchRun([]), status, finishedAt: null });
    assert.equal(els.financialReconciliationWorkbenchAutomationRule.disabled, true, status);
    assert.equal(els.financialReconciliationWorkbenchAutomationAnalyze.disabled, true, status);
  }
  assert.equal(openRun({ status: "completed", finishedAt: "2026-08-16T10:00:00Z" }), false);
  const terminal = renderAutomationWorkbench({ ...workbenchRun([]), status: "completed", finishedAt: "2026-08-16T10:00:00Z" });
  assert.equal(terminal.els.financialReconciliationWorkbenchAutomationRule.disabled, false);
  assert.equal(terminal.els.financialReconciliationWorkbenchAutomationAnalyze.disabled, false);
});

test("workbench keeps an unavailable active-run rule visible instead of silently selecting another rule", () => {
  const activeRun = {
    ...workbenchRun([]),
    status: "ready",
    finishedAt: null,
    definitions: [{
      ruleKey: "disabled-after-analysis",
      ruleVersion: 1,
      displayName: "Disabled Credit Card rule",
    }],
  };

  const { els, state } = renderAutomationWorkbench(activeRun, new Set(), {
    selectedRuleKey: "manual-enabled",
  });

  assert.equal(state.automation.selectedRuleKey, "disabled-after-analysis");
  assert.equal(els.financialReconciliationWorkbenchAutomationRule.value, "disabled-after-analysis");
  assert.equal(els.financialReconciliationWorkbenchAutomationRule.disabled, true);
  assert.equal(els.financialReconciliationWorkbenchAutomationAnalyze.disabled, true);
  assert.match(
    els.financialReconciliationWorkbenchAutomationRule.innerHTML,
    /value="disabled-after-analysis" selected>Disabled Credit Card rule \(active run\)<\/option>/,
  );
  assert.doesNotMatch(els.financialReconciliationWorkbenchAutomationRule.innerHTML, /Manual enabled/);
});

test("manual rule loader uses the app-authorized catalog without schedule administration", async () => {
  const current = {
    automation: { rules: [], run: null, selectedRuleKey: "manual-enabled", selectedProposalIds: new Set(), pendingAction: "", loaded: false },
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
  assert.equal(current.automation.selectedRuleKey, "manual-enabled");
  assert.equal(current.automation.loaded, true);
  assert.equal(Object.hasOwn(current.automation, "schedule"), false);
});

test("authoritative rule reload keeps a valid selection and defaults only an invalid one", async () => {
  const current = {
    activeTab: "manual",
    automation: { rules: [], run: null, selectedRuleKey: "missing", selectedProposalIds: new Set(), pendingAction: "", loaded: false },
  };
  const orderedRules = [
    { ...workbenchRules()[0], ruleKey: "second", displayName: "Second", priority: 2 },
    { ...workbenchRules()[0], ruleKey: "first", displayName: "First", priority: 1 },
  ];
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
    async () => ({ rules: orderedRules }),
    (value) => String(value ?? "").trim(),
    (value) => JSON.parse(JSON.stringify(value)),
    () => {},
    () => {},
  );

  await loadRules();
  assert.equal(current.automation.selectedRuleKey, "first");

  current.automation.loaded = false;
  current.automation.selectedRuleKey = "second";
  await loadRules();
  assert.equal(current.automation.selectedRuleKey, "second");

  current.automation.loaded = false;
  current.automation.run = {
    ...workbenchRun([]),
    status: "ready",
    finishedAt: null,
    definitions: [{ ruleKey: "disabled-after-analysis", ruleVersion: 1 }],
  };
  await loadRules();
  assert.equal(current.automation.selectedRuleKey, "disabled-after-analysis");
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

test("Analyze sends only the selected amount-only rule and a fresh UUID", async () => {
  const amountOnly = { ...managedRules()[2], enabled: true, allowManualExecution: true };
  const current = {
    automation: {
      rules: [...workbenchRules(), amountOnly],
      run: null,
      selectedRuleKey: amountOnly.ruleKey,
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
    "finalizeFinancialReconciliationAutomationAnalysis",
    `${appFunctionSource("financialReconciliationAutomationOpenRun")}
     ${appFunctionSource("financialReconciliationAutomationIsAnalyzing")}
     ${appFunctionSource("analyzeFinancialReconciliationAutomationRule").replace(/^function /, "async function ")}
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
    () => {
      current.automation.selectedProposalIds = new Set(current.automation.run.proposals
        .filter((proposal) => proposal.status === "proposed")
        .map((proposal) => proposal.id));
    },
  );

  await analyze();

  assert.deepEqual(calls, [{
    url: "/api/reconciliation-automation",
    options: {
      method: "POST",
      body: {
        action: "analyze_rule",
        ruleKeys: ["financial_documents_cgd_bank_statement_amount_only"],
        clientRequestId: "00000000-0000-0000-0000-000000000199",
      },
    },
  }]);
  assert.doesNotMatch(JSON.stringify(calls), /proposalIds/);
  assert.equal(current.automation.run, run);
  assert.deepEqual([...current.automation.selectedProposalIds], [WORKBENCH_PROPOSAL_1]);
  assert.equal(current.automation.pendingAction, "");
});

test("Analyze reports a terminal first-page failure instead of announcing ready proposals", async () => {
  const failedRun = {
    ...workbenchRun([]),
    status: "failed",
    analysisCompletedAt: null,
    finishedAt: "2026-08-16T09:00:01.000Z",
    analysisErrorCode: "analysis_continuation_failed",
  };
  const current = {
    automation: {
      rules: workbenchRules(),
      run: null,
      selectedProposalIds: new Set(),
      pendingAction: "",
      loaded: true,
      continuationToken: 0,
      continuationRetry: false,
    },
  };
  const finalized = [];
  const analyze = new Function(
    "financialReconciliationState",
    "api",
    "crypto",
    "clean",
    "renderFinancialReconciliationAutomation",
    "setFinancialReconciliationAutomationStatus",
    "showToast",
    "finalizeFinancialReconciliationAutomationAnalysis",
    `${appFunctionSource("financialReconciliationAutomationOpenRun")}
     ${appFunctionSource("financialReconciliationAutomationIsAnalyzing")}
     ${appFunctionSource("analyzeFinancialReconciliationAutomationRule").replace(/^function /, "async function ")}
     return analyzeFinancialReconciliationAutomationRule;`,
  )(
    () => current,
    async () => failedRun,
    { randomUUID: () => "00000000-0000-0000-0000-000000000198" },
    (value) => String(value ?? "").trim(),
    () => {},
    () => {},
    () => {},
    () => finalized.push(current.automation.run?.status),
  );

  await analyze("manual-enabled");

  assert.strictEqual(current.automation.run, failedRun);
  assert.deepEqual(finalized, ["failed"]);
  assert.deepEqual([...current.automation.selectedProposalIds], []);
});

test("failed analysis preserves the displayed terminal run and its prior selected proposals", async () => {
  const retainedRun = {
    ...workbenchRun([{ id: WORKBENCH_PROPOSAL_1, status: "completed" }]),
    status: "completed",
    finishedAt: "2026-08-16T09:05:00.000Z",
  };
  const retainedSelections = new Set([WORKBENCH_PROPOSAL_1]);
  let calls = 0;
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
    `${appFunctionSource("financialReconciliationAutomationOpenRun")}
     ${appFunctionSource("analyzeFinancialReconciliationAutomationRule").replace(/^function /, "async function ")}
     return analyzeFinancialReconciliationAutomationRule;`,
  )(
    () => current,
    async () => { calls += 1; throw new Error("analysis unavailable"); },
    { randomUUID: () => "00000000-0000-0000-0000-000000000299" },
    (value) => String(value ?? "").trim(),
    () => {},
    () => {},
    () => {},
  );

  await analyze("manual-enabled");

  assert.strictEqual(current.automation.run, retainedRun);
  assert.strictEqual(current.automation.selectedProposalIds, retainedSelections);
  assert.equal(calls, 1);
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

  assert.match(proposedMarkup, /financial-reconciliation-automation-proposal-meta/);
  assert.match(proposedMarkup, /type="checkbox"[^>]*checked/);
  assert.doesNotMatch(proposedMarkup, /type="checkbox"[^>]*disabled/);
  assert.match(proposedMarkup, /financial-reconciliation-automation-proposal-status[^>]*>proposed</);
  assert.doesNotMatch(proposedMarkup, /<header>/);
  assert.doesNotMatch(proposedMarkup, /<footer>/);
  assert.match(proposedMarkup, /Financial Documents[\s\S]*2026-08-01[\s\S]*Document FT 2026\/55[\s\S]*Supplier Safe Supplier/);
  assert.match(proposedMarkup, /financial-reconciliation-automation-item-amount">101\.00 â‚¬/);
  assert.match(proposedMarkup, /CGD Bank Statement[\s\S]*financial-reconciliation-automation-item-operator">-[\s\S]*-100\.00 â‚¬[\s\S]*Primary bank row/);
  assert.match(proposedMarkup, new RegExp(`Document number matched[\\s\\S]*FT202655[\\s\\S]*Description score 0\\.750 ${"\u2265"} 0\\.600`));
  assert.match(proposedMarkup, /Difference 1\.00 â‚¬[\s\S]*Allowed 1\.00 â‚¬[\s\S]*Manual enabled[\s\S]*version 3/i);
  assert.match(proposedMarkup, /Invoice &lt;script&gt;alert\(1\)&lt;\/script&gt;/);
  assert.doesNotMatch(proposedMarkup, /<script>/);

  assert.doesNotMatch(ambiguousMarkup, /type="checkbox"/);
  assert.match(ambiguousMarkup, /financial-reconciliation-automation-proposal-status[^>]*>ambiguous</);
  assert.match(ambiguousMarkup, /Multiple qualifying combinations/);
  assert.match(ambiguousMarkup, /Candidate group 1[\s\S]*Candidate group A[\s\S]*Candidate group 2[\s\S]*Candidate group B/);
  assert.match(proposedMarkup, new RegExp(`aria-label="Execute automatic proposal for Financial Documents record ${baseSnapshot.sourceId}"`));

  for (const status of ["completed", "stale", "failed"]) {
    const lifecycleMarkup = proposalMarkup({ ...proposed, status }, run, workbenchRules(), new Set(), false);
    assert.doesNotMatch(lifecycleMarkup, /type="checkbox"/);
    assert.match(lifecycleMarkup, new RegExp(`financial-reconciliation-automation-proposal-status[^>]*>${status}<`));
    assert.match(lifecycleMarkup, /Difference[\s\S]*Allowed[\s\S]*version 3/i);
  }

  const hostileBase = {
    sourceType: "financial_documents",
    sourceId: "document-<one>",
    sourceDate: "2026-08-16",
    docNumber: "FT <42>",
    supplierName: "Supplier & Sons",
    description: "Invoice <script>alert(1)</script>",
    amount: 100,
  };
  const compactMarkup = proposalMarkup(
    {
      id: WORKBENCH_PROPOSAL_1,
      ruleKey: "manual-enabled",
      ruleVersion: 3,
      status: "proposed",
      baseSnapshot: hostileBase,
      items: [{
        sourceType: "import_cgd_extrato_ordem",
        sourceId: "bank-<one>",
        sourceDate: "2026-08-16",
        description: "Bank <match>",
        amount: -100,
        evidence: { documentNumber: { matched: true, normalized: "FT42" } },
      }],
      calculatedDifference: 0,
      allowedDifference: 1,
    },
    workbenchRun([]),
    workbenchRules(),
    new Set([WORKBENCH_PROPOSAL_1]),
    false,
  );

  assert.match(compactMarkup, /financial-reconciliation-automation-item-meta/);
  assert.match(compactMarkup, /financial-reconciliation-automation-item-description/);
  assert.match(compactMarkup, /financial-reconciliation-automation-item-operator/);
  assert.match(compactMarkup, /financial-reconciliation-automation-item-id/);
  assert.match(compactMarkup, /Supplier &amp; Sons/);
  assert.match(compactMarkup, /FT &lt;42&gt;/);
  assert.match(compactMarkup, /Invoice &lt;script&gt;alert\(1\)&lt;\/script&gt;/);
  assert.doesNotMatch(compactMarkup, /<script>/);
  assert.match(compactMarkup, /Document number matched: FT42/);
  assert.match(compactMarkup, /Difference[\s\S]*Allowed[\s\S]*version 3/);
  assert.match(compactMarkup, new RegExp(`type="checkbox"[^>]*aria-label="Execute automatic proposal for Financial Documents record document-&lt;one&gt;"[^>]*data-financial-reconciliation-automation-proposal-id="${WORKBENCH_PROPOSAL_1}"[^>]*checked`));
  assert.match(compactMarkup, /<article class="financial-reconciliation-automation-proposal[^>]*tabindex="-1"/);
  assert.match(compactMarkup, /Record ID<\/summary><code>document-&lt;one&gt;<\/code>/);
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
