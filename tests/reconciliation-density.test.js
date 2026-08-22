const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
const css = fs.readFileSync(path.join(root, "styles.css"), "utf8");
const appMain = fs.readFileSync(path.join(root, "app-main.js"), "utf8");
const reconciliationSettingsApi = fs.readFileSync(path.join(root, "api", "reconciliation-settings.js"), "utf8");
const supabaseApi = fs.readFileSync(path.join(root, "api", "_supabase.js"), "utf8");

function appFunctionSource(name) {
  const functionStart = appMain.indexOf(`function ${name}(`);
  assert.notEqual(functionStart, -1, `${name} should be defined in app-main.js`);
  const start = appMain.slice(Math.max(0, functionStart - 6), functionStart) === "async " ? functionStart - 6 : functionStart;
  const bodyStart = appMain.indexOf("{", appMain.indexOf(")", functionStart));
  let depth = 0;
  for (let index = bodyStart; index < appMain.length; index += 1) {
    if (appMain[index] === "{") depth += 1;
    if (appMain[index] === "}") depth -= 1;
    if (depth === 0) return appMain.slice(start, index + 1);
  }
  throw new Error(`Could not extract ${name} from app-main.js`);
}

function appConstantSource(name) {
  const start = appMain.indexOf(`const ${name} = `);
  assert.notEqual(start, -1, `${name} should be defined in app-main.js`);
  const end = appMain.indexOf("function ", start);
  assert.notEqual(end, -1, `${name} should end before a function declaration`);
  return appMain.slice(start, end);
}

const financialReconciliationItemDetails = new Function(`${appFunctionSource("clean")}\n${appFunctionSource("formatDateOnly")}\n${appFunctionSource("financialReconciliationItemDetails")}\nreturn financialReconciliationItemDetails;`)();
const financialReconciliationSummaryMarkup = new Function(
  "financialReconciliationStatusMarkup",
  "formatMoney",
  "escape",
  `${appFunctionSource("financialReconciliationSummaryMarkup")}\nreturn financialReconciliationSummaryMarkup;`,
)(
  (status) => `<span class="financial-reconciliation-status">${status}</span>`,
  (amount) => `${Number(amount).toFixed(2)} €`,
  (value) => String(value),
);

const historySourceHelpers = new Function(
  "FINANCIAL_RECONCILIATION_SOURCES",
  "financialReconciliationSourceLabel",
  "formatMoney",
  "clean",
  `${appFunctionSource("financialReconciliationHistorySourceSummary")}
   ${appFunctionSource("financialReconciliationHistorySourceText")}
   return { financialReconciliationHistorySourceSummary, financialReconciliationHistorySourceText };`,
)(
  {
    financial_documents: "Financial Documents",
    import_cgd_extrato_ordem: "CGD Bank Statement",
  },
  (value) => ({
    financial_documents: "Financial Documents",
    import_cgd_extrato_ordem: "CGD Bank Statement",
  })[value] || value,
  (value) => `${Number(value).toFixed(2)} €`,
  (value) => String(value || "").trim(),
);

function renderCurrentSummary({ status, difference, items, origin, automaticTrigger, audit = [] }) {
  const current = { workspace: { reconciliation: { id: "rec-1", status, difference_amount: difference, completion_type: "normal", base_source_type: "financial_documents", matching_source_types: ["import_cgd_extrato_ordem"], origin, automaticTrigger }, items, audit } };
  const els = {
    financialReconciliationCurrent: { innerHTML: "" },
    financialReconciliationNew: { hidden: true },
    financialReconciliationReopen: { hidden: true },
    financialReconciliationDelete: { hidden: true },
  };
  const render = new Function(
    "financialReconciliationState", "normalizeFinancialReconciliationWorkspace", "financialReconciliationActiveRecord",
    "financialReconciliationDifference", "clean", "financialReconciliationCompletionDraft",
    "financialReconciliationCompletionPresentation", "financialReconciliationItemDetails", "escape",
    "financialReconciliationSourceLabel", "formatMoney", "formatDateTimeShort",
    "financialReconciliationStatusMarkup", "financialReconciliationSummaryMarkup", "els", "financialReconciliationOriginMarkup", "financialReconciliationAutomaticAuditMarkup",
    `${appFunctionSource("renderFinancialReconciliationCurrent")}\nreturn renderFinancialReconciliationCurrent;`,
  )(
    () => current,
    (value) => value,
    () => current.workspace.reconciliation,
    (value) => Number(value.difference_amount),
    (value) => String(value || "").trim(),
    () => "",
    () => ({ required: false, disabled: false, label: "Complete reconciliation" }),
    () => "",
    (value) => String(value),
    (value) => ({ financial_documents: "Financial Documents", import_cgd_extrato_ordem: "CGD Bank Statement" })[value] || value,
    (value) => `${Number(value).toFixed(2)} €`,
    () => "",
    (value) => `<span class="financial-reconciliation-status">${value}</span>`,
    financialReconciliationSummaryMarkup,
    els,
    (value) => `<span class="origin">${value.origin === "automatic" ? "Automatic" : "User"}</span>`,
    (entry) => entry.action === "automatic_complete" ? '<details data-automatic-audit-evidence></details>' : "",
  );
  render();
  return els.financialReconciliationCurrent.innerHTML;
}

test("reloaded automatic reconciliations display escaped structured audit evidence", () => {
  const automaticAuditMarkup = new Function(
    "clean",
    "escape",
    "formatMoney",
    `${appFunctionSource("financialReconciliationAutomaticAuditMarkup")}
     return financialReconciliationAutomaticAuditMarkup;`,
  )(
    (value) => String(value ?? "").trim(),
    new Function(`${appFunctionSource("escape")}; return escape;`)(),
    (value) => `${Number(value).toFixed(2)} â‚¬`,
  );
  const entry = {
    action: "automatic_complete",
    metadata: {
      ruleSnapshot: { ruleKey: "financial_documents_cgd_credit_card", ruleVersion: 1 },
      configSnapshot: { differenceAllowed: 0, maxDifferenceDays: 10 },
      operatorSnapshot: { import_cgd_cartao_credito: "+" },
      identityEvidence: [{
        documentNumber: { matched: true, normalized: "FT<script>" },
        description: { matched: true, score: 0.55, threshold: 0.55 },
      }],
      trigger: "manual",
      runId: "run-123",
      calculatedDifference: 0,
    },
  };

  const evidence = automaticAuditMarkup(entry);
  assert.match(evidence, /Automatic evidence/);
  assert.match(evidence, /financial_documents_cgd_credit_card[\s\S]*version 1/);
  assert.match(evidence, /manual[\s\S]*run-123[\s\S]*0\.00 â‚¬[\s\S]*0\.00 â‚¬/i);
  assert.match(evidence, /CGD Credit Card operator[\s\S]*\+/);
  assert.match(evidence, /Document number FT&lt;script&gt;[\s\S]*Description 0\.550 ≥ 0\.550/);
  assert.doesNotMatch(evidence, /<script>/);

  const currentMarkup = renderCurrentSummary({
    status: "complete",
    difference: 0,
    items: [{ source_type: "financial_documents", amount_snapshot: 10 }],
    origin: "automatic",
    automaticTrigger: "manual",
    audit: [entry],
  });
  assert.match(currentMarkup, /data-automatic-audit-evidence/);
});
const reconciliationSettingsDestinationHelpers = new Function("appFeatures", "lastMainView", `
const state = { lastMainView, access: { appFeatures } };
${appFunctionSource("clean")}
${appFunctionSource("canApp")}
${appFunctionSource("canAppFinancialDocs")}
${appFunctionSource("canAppImportData")}
${appFunctionSource("canAppBankAccounts")}
${appFunctionSource("canAppFinancialReconciliation")}
${appFunctionSource("canUseGuestsBi")}
${appFunctionSource("canUseBookingsBi")}
${appFunctionSource("canUseFinancialBi")}
${appFunctionSource("canUseSalesBi")}
${appFunctionSource("preferredMainAppView")}
${appFunctionSource("reconciliationSettingsAppDestination")}
return { preferredMainAppView, reconciliationSettingsAppDestination };
`);

test("settings exposes a reconciliation rule editor", () => {
  assert.match(html, /id="settings-menu-financial-reconciliation"/);
  assert.match(html, /id="settings-view-financial-reconciliation"/);
  assert.match(html, /id="financial-reconciliation-settings-base-source"/);
  assert.match(html, /id="financial-reconciliation-settings-rules-body"/);
  assert.match(html, /id="financial-reconciliation-settings-save"/);
});

test("automatic reconciliation settings stay dense, wrapping, and reachable on narrow screens", () => {
  assert.match(html, /class="[^"]*financial-reconciliation-settings-tabs[^"]*"/);
  const sourcePanel = html.slice(
    html.indexOf('id="financial-reconciliation-settings-source-panel"'),
    html.indexOf('id="financial-reconciliation-settings-automatic-panel"'),
  );
  for (const id of [
    "financial-reconciliation-settings-base-source",
    "financial-reconciliation-settings-rules-body",
    "financial-reconciliation-settings-save",
  ]) assert.match(sourcePanel, new RegExp(`id="${id}"`));
  assert.match(css, /\.financial-reconciliation-automation-rule-list\s*\{[\s\S]*display:\s*grid;[\s\S]*gap:\s*\.65rem;/);
  assert.match(css, /\.financial-reconciliation-automation-schedule\s*\{[\s\S]*grid-template-columns:\s*repeat\(auto-fit,\s*minmax\(10rem,\s*1fr\)\);/);
  assert.match(css, /\.financial-reconciliation-automation-logic\s*\{[\s\S]*overflow-wrap:\s*anywhere;[\s\S]*white-space:\s*pre-wrap;/);
  assert.match(css, /\.financial-reconciliation-automation-rule-controls\s*\{[\s\S]*grid-template-columns:\s*repeat\(5,\s*minmax\(0,\s*1fr\)\);/);
  assert.match(css, /\.financial-reconciliation-automation-fixed-value\s*\{[\s\S]*min-height:[\s\S]*border:[\s\S]*background:/);
  assert.match(css, /@media \(max-width:\s*768px\)[\s\S]*\.financial-reconciliation-automation-rule-controls\s*\{[\s\S]*grid-template-columns:\s*1fr;/);
  assert.match(css, /@media \(max-width:\s*768px\)[\s\S]*\.financial-reconciliation-automation-fixed-value\s*\{[\s\S]*font-size:\s*16px;/);
  assert.match(css, /@media \(max-width:\s*768px\)[\s\S]*\.financial-reconciliation-settings-tabs\s*>\s*button\s*\{[\s\S]*flex:\s*1 1 12rem;/);
  assert.match(css, /\.financial-reconciliation-automation-rule-card\s*:focus-within\s*\{/);
  assert.match(css, /\.financial-reconciliation-automation-settings\s+input\[aria-invalid="true"\]\s*\{[\s\S]*border-color:/);
  assert.match(html, /id="financial-reconciliation-automation-open-workbench"[^>]*>Open automatic reconciliation<\/button>/);
  assert.match(appMain, /<input type="number" min="0" max="90"[^>]*data-reconciliation-automation-rule-field="maxDifferenceDays"/);
});

test("reconciliation settings validates and sends one atomic replacement RPC", () => {
  assert.match(reconciliationSettingsApi, /requireFeature\(req, "settings", "financial-reconciliation"\)/);
  assert.match(supabaseApi, /SETTINGS_FEATURES = \[[^\]]*"financial-reconciliation"/);
  assert.match(appMain, /SETTINGS_FEATURE_OPTIONS = \[[^\]]*"financial-reconciliation"/);
  const validation = reconciliationSettingsApi.indexOf("const input = normalizeReconciliationRules");
  const replacement = reconciliationSettingsApi.indexOf('restQuery("rpc/replace_financial_reconciliation_source_rules"');
  assert.ok(validation >= 0 && validation < replacement, "validation must complete before replacement begins");
  assert.doesNotMatch(reconciliationSettingsApi, /restQuery\("financial_reconciliation_source_rules", \{ method: "DELETE" \}\)/);
  assert.doesNotMatch(reconciliationSettingsApi, /restQuery\("financial_reconciliation_source_rules", \{ method: "POST"/);
});

test("reconciliation settings keeps its contextual tab and backs out only to authorized app views", () => {
  assert.match(appMain, /previousView === "financial-reconciliation" && canSettings\("financial-reconciliation"\)/);
  assert.match(appMain, /function reconciliationSettingsAppDestination\(\)/);
  assert.match(appMain, /els\.closeSettingsFinancialReconciliation\?\.addEventListener\("click", closeReconciliationSettings\)/);
  assert.doesNotMatch(appMain, /closeSettingsFinancialReconciliation\?\.addEventListener\([^\n]*"communications"/);
});

test("reconciliation settings destination uses actual app feature guards", () => {
  const settingsOnly = reconciliationSettingsDestinationHelpers([], "communications");
  assert.equal(settingsOnly.preferredMainAppView(), "");
  assert.equal(settingsOnly.reconciliationSettingsAppDestination(), "");

  const noCommunications = reconciliationSettingsDestinationHelpers(["guests"], "communications");
  assert.equal(noCommunications.preferredMainAppView(), "guests");
  assert.equal(noCommunications.reconciliationSettingsAppDestination(), "guests");
  assert.notEqual(noCommunications.reconciliationSettingsAppDestination(), "communications");

  const reconciliationWithoutBackoffice = reconciliationSettingsDestinationHelpers(["financial-reconciliation"], "");
  assert.equal(reconciliationWithoutBackoffice.reconciliationSettingsAppDestination(), "");

  const reconciliationApp = reconciliationSettingsDestinationHelpers(["backoffice", "financial-reconciliation"], "");
  assert.equal(reconciliationApp.reconciliationSettingsAppDestination(), "financial-reconciliation");
});

test("workbench uses one source selector and displays saved rule hints", () => {
  assert.match(html, /<label>Source<select id="financial-reconciliation-source"><\/select><\/label>/);
  assert.match(html, /id="financial-reconciliation-rule-hint"/);
  assert.doesNotMatch(html, /id="financial-reconciliation-base-source"/);
  assert.doesNotMatch(html, /id="financial-reconciliation-matching-sources"/);
  assert.doesNotMatch(html, /id="financial-reconciliation-candidate-source"/);
  assert.match(appMain, /function onFinancialReconciliationSourceChange\(\)/);
  assert.match(appMain, /action: "start", sourceType, sourceId/);
});

test("workbench exposes source controls before the dynamic filter row", () => {
  assert.match(
    html,
    /class="financial-reconciliation-source-row"[\s\S]*id="financial-reconciliation-source"[\s\S]*id="financial-reconciliation-rule-hint"[\s\S]*<\/div>\s*<div id="financial-reconciliation-dynamic-filters"/,
  );
  assert.match(appMain, /class="financial-reconciliation-filter-\$\{escape\(field\)\}"/);
});

test("reconciliation separates manual and automatic work into accessible tabs with shared history", () => {
  assert.match(html, /class="[^"]*financial-reconciliation-view-tabs[^"]*"[^>]*role="tablist"/);
  assert.match(html, /id="financial-reconciliation-manual-tab"[^>]*role="tab"[^>]*aria-controls="financial-reconciliation-manual-panel"/);
  assert.match(html, /id="financial-reconciliation-automatic-tab"[^>]*role="tab"[^>]*aria-controls="financial-reconciliation-automatic-panel"/);
  assert.match(html, /id="financial-reconciliation-manual-panel"[^>]*role="tabpanel"[^>]*aria-labelledby="financial-reconciliation-manual-tab"/);
  assert.match(html, /id="financial-reconciliation-automatic-panel"[^>]*role="tabpanel"[^>]*aria-labelledby="financial-reconciliation-automatic-tab"[^>]*hidden/);

  const manualPanel = html.indexOf('id="financial-reconciliation-manual-panel"');
  const automaticPanel = html.indexOf('id="financial-reconciliation-automatic-panel"');
  const history = html.indexOf('class="card financial-reconciliation-history-card"');
  assert.ok(manualPanel >= 0 && automaticPanel > manualPanel && history > automaticPanel);
  assert.match(html.slice(manualPanel, automaticPanel), /id="financial-reconciliation-filters"[\s\S]*id="financial-reconciliation-current"/);
  assert.match(html.slice(automaticPanel, history), /class="financial-reconciliation-workbench-automation-rule-picker"[\s\S]*for="financial-reconciliation-workbench-automation-rule"[\s\S]*id="financial-reconciliation-workbench-automation-rule"[\s\S]*id="financial-reconciliation-workbench-automation-analyze"[\s\S]*id="financial-reconciliation-workbench-automation-proposals"/);
  assert.match(html.slice(automaticPanel, history), /id="financial-reconciliation-status"[\s\S]*role="status"[\s\S]*aria-live="polite"/);
  assert.equal((html.match(/id="financial-reconciliation-history-rows"/g) || []).length, 1);
  assert.ok(html.indexOf('id="financial-reconciliation-history-rows"') > html.indexOf('id="financial-reconciliation-automatic-panel"'));
});

test("reconciliation provides a dedicated searchable History tab while retaining compact history", () => {
  assert.match(html, /id="financial-reconciliation-history-tab"[^>]*role="tab"[^>]*aria-controls="financial-reconciliation-history-panel"/);
  assert.match(html, /id="financial-reconciliation-history-panel"[^>]*role="tabpanel"[^>]*aria-labelledby="financial-reconciliation-history-tab"[^>]*hidden/);
  assert.match(html, /id="financial-reconciliation-history-filters"[\s\S]*id="financial-reconciliation-history-created-from"[\s\S]*id="financial-reconciliation-history-created-to"/);
  assert.match(html, /id="financial-reconciliation-history-origin"[\s\S]*<option value="user">User<\/option>[\s\S]*<option value="automatic">Automatic<\/option>/);
  assert.match(html, /id="financial-reconciliation-history-status-filter"[\s\S]*<option value="not_started">Not started<\/option>[\s\S]*<option value="started">Started<\/option>[\s\S]*<option value="complete">Complete<\/option>/);
  assert.match(html, /id="financial-reconciliation-history-difference-from"[\s\S]*id="financial-reconciliation-history-difference-to"/);
  assert.match(html, /<th>Created<\/th>[\s\S]*<th>Source<\/th>[\s\S]*<th>Destination<\/th>[\s\S]*<th># records<\/th>[\s\S]*<th>Origin<\/th>[\s\S]*<th>Status<\/th>[\s\S]*<th>Source total<\/th>[\s\S]*<th>Destination total<\/th>[\s\S]*<th>Difference<\/th>[\s\S]*<th>Completion comment<\/th>/);
  assert.match(html, /id="financial-reconciliation-history-search-rows"/);
  assert.match(html, /id="financial-reconciliation-history-card"[^>]*class="[^"]*financial-reconciliation-history-card/);
});

test("reconciliation application tabs remain visible and usable on narrow screens", () => {
  assert.match(css, /\.financial-reconciliation-view-tabs\s*\{[\s\S]*flex-wrap:\s*wrap;[\s\S]*margin-bottom:\s*1rem;/);
  assert.match(css, /\.financial-reconciliation-view-tabs\s*>\s*button:focus-visible\s*\{[\s\S]*outline:/);
  assert.match(css, /\.financial-reconciliation-manual-panel:not\(\[hidden\]\),[\s\S]*\.financial-reconciliation-automatic-panel:not\(\[hidden\]\),[\s\S]*\.financial-reconciliation-history-panel:not\(\[hidden\]\)\s*\{[\s\S]*display:\s*grid;/);
  assert.match(css, /@media \(max-width:\s*768px\)[\s\S]*\.financial-reconciliation-view-tabs\s*>\s*button\s*\{[\s\S]*flex:\s*1 1 12rem;/);
  assert.match(html, /id="financial-reconciliation-manual-panel"[^>]*class="[^"]*financial-reconciliation-manual-panel/);
  assert.match(html, /id="financial-reconciliation-automatic-panel"[^>]*class="[^"]*financial-reconciliation-automatic-panel/);
  assert.match(css, /\.financial-reconciliation-workbench-automation-rule-picker\s*\{[\s\S]*display:\s*grid;[\s\S]*grid-template-columns:\s*minmax\(0,\s*1fr\)\s+auto;/);
  assert.match(css, /\.financial-reconciliation-workbench-automation-rule-picker\s+(?:select|button):focus-visible\s*\{[\s\S]*outline:/);
  assert.match(css, /@media \(max-width:\s*768px\)[\s\S]*\.financial-reconciliation-workbench-automation-rule-picker\s*\{[\s\S]*grid-template-columns:\s*1fr;/);
  assert.match(css, /@media \(max-width:\s*768px\)[\s\S]*\.financial-reconciliation-workbench-automation-rule-picker\s+(?:select|button)\s*\{[\s\S]*width:\s*100%;/);
});

function reconciliationTabClassList() {
  const values = new Set();
  return {
    toggle(name, force) {
      if (force) values.add(name);
      else values.delete(name);
    },
    contains: (name) => values.has(name),
  };
}

function reconciliationTabElements() {
  const tab = () => ({
    classList: reconciliationTabClassList(),
    focused: false,
    focusCalls: 0,
    setAttribute(name, value) { this[name] = value; },
    focus() { this.focused = true; this.focusCalls += 1; },
  });
  return {
    financialReconciliationManualTab: tab(),
    financialReconciliationAutomaticTab: tab(),
    financialReconciliationHistoryTab: tab(),
    financialReconciliationManualPanel: { hidden: false },
    financialReconciliationAutomaticPanel: { hidden: true },
    financialReconciliationHistoryPanel: { hidden: true },
    financialReconciliationHistoryCard: { hidden: false },
  };
}

function compileReconciliationTabController({ current, els, calls, loadRules = async () => { calls.push("load-rules"); current.automation.loaded = true; }, loadHistory = async () => { calls.push("load-history"); current.historySearch.loaded = true; } }) {
  let controller;
  controller = new Function(
    "clean",
    "financialReconciliationState",
    "renderFinancialReconciliation",
    "loadFinancialReconciliationAutomationRules",
    "loadFinancialReconciliationHistory",
    "els",
    `${appFunctionSource("normalizeFinancialReconciliationTab")}
     ${appFunctionSource("renderFinancialReconciliationTabs")}
     ${appFunctionSource("setFinancialReconciliationTab")}
     ${appFunctionSource("onFinancialReconciliationTabKeydown")}
     return { normalizeFinancialReconciliationTab, renderFinancialReconciliationTabs, setFinancialReconciliationTab, onFinancialReconciliationTabKeydown };`,
  )(
    (value) => String(value ?? "").trim(),
    () => current,
    () => { calls.push(`render:${current.activeTab}`); controller.renderFinancialReconciliationTabs(); },
    loadRules,
    loadHistory,
    els,
  );
  return controller;
}

test("reconciliation tab controller defaults to Manual and supports arrow activation", async () => {
  const current = {
    activeTab: "manual",
    filters: { description: "retained" },
    automation: { loaded: true, run: { runId: "retained-run" } },
    historySearch: { loaded: false },
  };
  const calls = [];
  const els = reconciliationTabElements();
  const controller = compileReconciliationTabController({ current, els, calls });

  controller.renderFinancialReconciliationTabs();
  assert.equal(els.financialReconciliationManualPanel.hidden, false);
  assert.equal(els.financialReconciliationAutomaticPanel.hidden, true);
  assert.equal(els.financialReconciliationManualTab["aria-selected"], "true");
  assert.equal(els.financialReconciliationAutomaticTab["aria-selected"], "false");
  assert.equal(els.financialReconciliationManualTab.tabindex, "0");
  assert.equal(els.financialReconciliationAutomaticTab.tabindex, "-1");

  await controller.setFinancialReconciliationTab("automatic", { focus: true });
  assert.equal(current.activeTab, "automatic");
  assert.equal(current.filters.description, "retained");
  assert.equal(current.automation.run.runId, "retained-run");
  assert.equal(els.financialReconciliationAutomaticPanel.hidden, false);
  assert.equal(els.financialReconciliationManualTab["aria-selected"], "false");
  assert.equal(els.financialReconciliationAutomaticTab["aria-selected"], "true");
  assert.equal(els.financialReconciliationManualTab.tabindex, "-1");
  assert.equal(els.financialReconciliationAutomaticTab.tabindex, "0");
  assert.equal(els.financialReconciliationAutomaticTab.focused, true);

  controller.onFinancialReconciliationTabKeydown({ key: "ArrowRight", preventDefault() { calls.push("prevent"); } });
  await Promise.resolve();
  assert.equal(current.activeTab, "history");
  assert.equal(els.financialReconciliationHistoryPanel.hidden, false);
  assert.equal(els.financialReconciliationHistoryCard.hidden, true);
  assert.ok(calls.includes("load-history"));

  controller.onFinancialReconciliationTabKeydown({ key: "ArrowRight", preventDefault() { calls.push("prevent"); } });
  await Promise.resolve();
  assert.equal(current.activeTab, "manual");
  assert.ok(calls.includes("prevent"));
});

test("late Automatic rule loading cannot steal focus after returning to Manual", async () => {
  const current = { activeTab: "manual", automation: { loaded: false } };
  const els = reconciliationTabElements();
  const calls = [];
  let resolveRules;
  const controller = compileReconciliationTabController({
    current,
    els,
    calls,
    loadRules: () => new Promise((resolve) => { resolveRules = () => { current.automation.loaded = true; resolve(); }; }),
  });

  const automaticActivation = controller.setFinancialReconciliationTab("automatic", { focus: true });
  await Promise.resolve();
  assert.equal(els.financialReconciliationAutomaticTab.focusCalls, 1, "Automatic receives focus when it is activated");
  await controller.setFinancialReconciliationTab("manual", { focus: true });
  resolveRules();
  await automaticActivation;

  assert.equal(current.activeTab, "manual");
  assert.equal(els.financialReconciliationManualTab.focusCalls, 1);
  assert.equal(els.financialReconciliationAutomaticTab.focusCalls, 1, "the settled Automatic load does not refocus its stale tab");
});

test("late Settings Automatic handoff focus respects the tab selected while data loads", async () => {
  const state = { access: { settingsFeatures: [] }, currentView: "settings", financialReconciliation: { activeTab: "manual" } };
  const els = reconciliationTabElements();
  let releaseDataLoad;
  const setView = new Function(
    "state",
    "els",
    "showToast",
    "canAppFinancialReconciliation",
    "setMobileNavOpen",
    "syncAppRoute",
    "renderLayout",
    "renderSettingsSection",
    "render",
    "ensureCurrentViewData",
    `${appFunctionSource("financialReconciliationState")}
     ${appFunctionSource("clean")}
     ${appFunctionSource("normalizeFinancialReconciliationTab")}
     ${appFunctionSource("financialReconciliationEntryTab")}
     ${appFunctionSource("focusFinancialReconciliationTab")}
     ${appFunctionSource("setView")}
     return setView;`,
  )(
    state,
    els,
    () => {},
    () => true,
    () => {},
    () => {},
    () => {},
    () => {},
    () => {},
    () => new Promise((resolve) => { releaseDataLoad = resolve; }),
  );

  const handoff = setView("financial-reconciliation", { financialReconciliationTab: "automatic" });
  await Promise.resolve();

  state.financialReconciliation.activeTab = "manual";
  releaseDataLoad();
  await handoff;

  assert.equal(els.financialReconciliationAutomaticTab.focusCalls, 0);
});

test("automatic proposals render as three divided open desktop columns", () => {
  assert.match(css, /^:root\s*\{[^}]*--line\s*:[^;}]+;[^}]*\}/);
  assert.match(css, /\.financial-reconciliation-workbench-automation-proposals\s*\{[^}]*gap:\s*0;[^}]*border-top:\s*1px solid var\(--line\)[^}]*\}/);
  assert.match(css, /\.financial-reconciliation-automation-proposal\s*\{[^}]*grid-template-columns:\s*minmax\([^;]+\)\s+minmax\(0,\s*1fr\)\s+minmax\(0,\s*1fr\)[^}]*border:\s*0;[^}]*border-bottom:\s*1px solid var\(--line\)[^}]*border-left:\s*4px solid var\(--brand\)[^}]*border-radius:\s*0[^}]*background:\s*transparent[^}]*\}/);
  assert.match(css, /\.financial-reconciliation-automation-proposal-meta\s*\{[^}]*border-right:\s*1px solid var\(--line\)[^}]*background:\s*transparent[^}]*\}/);
  assert.match(css, /\.financial-reconciliation-automation-item\s*\{[^}]*background:\s*transparent[^}]*\}/);
});

test("monthly automatic proposals keep three separated desktop columns and accessible group controls", () => {
  assert.match(css, /\.financial-reconciliation-automation-proposal--monthly\s*\{[^}]*grid-template-columns:\s*minmax\(11rem,\s*\.55fr\)\s+minmax\(0,\s*1fr\)\s+minmax\(0,\s*1fr\)[^}]*\}/);
  assert.match(css, /\.financial-reconciliation-automation-member-group\s*\{[^}]*min-width:\s*0[^}]*border-left:\s*1px solid var\(--line\)[^}]*\}/);
  assert.match(css, /\.financial-reconciliation-automation-member-group\s*>\s*summary\s*\{[^}]*cursor:\s*pointer[^}]*overflow-wrap:\s*anywhere[^}]*\}/);
  assert.match(css, /\.financial-reconciliation-automation-member-group\s*>\s*summary:focus-visible\s*\{[^}]*outline:\s*3px solid[^}]*\}/);
  assert.match(css, /\.financial-reconciliation-automation-member-load-more\s*\{[^}]*min-height:[^}]*\}/);
  assert.match(css, /\.financial-reconciliation-automation-member-error\s*\{[^}]*color:\s*var\(--danger\)[^}]*\}/);
});

test("automatic proposal desktop keeps every destination and candidate group in the third column", () => {
  assert.match(css, /\.financial-reconciliation-automation-proposal-records\s*>\s*:first-child\s*\{[^}]*grid-column:\s*1[^}]*\}/);
  assert.match(css, /\.financial-reconciliation-automation-proposal-records\s*>\s*:not\(:first-child\)\s*\{[^}]*grid-column:\s*2[^}]*border-left:\s*1px solid var\(--line\)[^}]*\}/);
  assert.match(css, /\.financial-reconciliation-automation-proposal-records\s*>\s*:not\(:first-child\):not\(:nth-child\(2\)\)\s*\{[^}]*border-top:\s*1px solid var\(--line\)[^}]*\}/);
  assert.match(css, /\.financial-reconciliation-automation-candidate-group\s*\{[^}]*grid-template-columns:\s*minmax\(0,\s*1fr\)[^}]*border:\s*0[^}]*background:\s*transparent[^}]*\}/);
});

test("automatic proposal dividers become section separators on narrow screens", () => {
  const narrowStart = css.indexOf("@media (max-width: 700px)");
  const narrowEnd = css.indexOf("@media (max-width: 620px)", narrowStart);
  assert.notEqual(narrowStart, -1);
  assert.notEqual(narrowEnd, -1);
  const narrowCss = css.slice(narrowStart, narrowEnd);

  assert.match(narrowCss, /\.financial-reconciliation-automation-proposal\s*\{[^}]*grid-template-columns:\s*minmax\(0,\s*1fr\)[^}]*\}/);
  assert.match(narrowCss, /\.financial-reconciliation-automation-proposal-meta\s*\{[^}]*border-right:\s*0[^}]*border-bottom:\s*1px solid var\(--line\)[^}]*\}/);
  assert.match(narrowCss, /\.financial-reconciliation-automation-proposal-records\s*>\s*:first-child\s*,\s*\.financial-reconciliation-automation-proposal-records\s*>\s*:not\(:first-child\)\s*\{[^}]*grid-column:\s*1[^}]*\}/);
  assert.match(narrowCss, /\.financial-reconciliation-automation-proposal-records\s*>\s*:not\(:first-child\)\s*\{[^}]*border-left:\s*0[^}]*border-top:\s*1px solid var\(--line\)[^}]*\}/);
});

test("monthly proposal groups stack with horizontal separators and usable controls at 768px", () => {
  const narrowStart = css.indexOf("@media (max-width: 768px)", css.indexOf(".financial-reconciliation-automation-proposal--monthly"));
  const narrowEnd = css.indexOf("@media", narrowStart + 1);
  assert.notEqual(narrowStart, -1);
  const narrowCss = css.slice(narrowStart, narrowEnd === -1 ? css.length : narrowEnd);

  assert.match(narrowCss, /\.financial-reconciliation-automation-proposal--monthly\s*\{[^}]*grid-template-columns:\s*minmax\(0,\s*1fr\)[^}]*\}/);
  assert.match(narrowCss, /\.financial-reconciliation-automation-proposal--monthly\s+\.financial-reconciliation-automation-proposal-meta\s*\{[^}]*border-right:\s*0[^}]*border-bottom:\s*1px solid var\(--line\)[^}]*\}/);
  assert.match(narrowCss, /\.financial-reconciliation-automation-member-group\s*\{[^}]*border-left:\s*0[^}]*border-top:\s*1px solid var\(--line\)[^}]*\}/);
  assert.match(narrowCss, /\.financial-reconciliation-automation-member-group\s*>\s*summary\s*,\s*\.financial-reconciliation-automation-member-load-more\s*\{[^}]*font-size:\s*16px[^}]*\}/);
  assert.match(narrowCss, /\.financial-reconciliation-automation-member-row-description\s*\{[^}]*overflow-wrap:\s*anywhere[^}]*white-space:\s*normal[^}]*\}/);
  assert.match(narrowCss, /\.financial-reconciliation-automation-member-actions\s*\{[^}]*flex-wrap:\s*wrap[^}]*\}/);
});

test("workbench rule snapshots keep only valid unique signed sources", () => {
  const normalizeRuleSnapshot = new Function("FINANCIAL_RECONCILIATION_SOURCES", `
${appFunctionSource("clean")}
${appFunctionSource("normalizeFinancialReconciliationRuleSnapshot")}
return normalizeFinancialReconciliationRuleSnapshot;
`)({ financial_documents: "Financial Documents", import_cgd_extrato_ordem: "CGD Bank Statement" });

  assert.deepEqual(normalizeRuleSnapshot([
    { sourceType: "import_cgd_extrato_ordem", operator: "-" },
    { sourceType: "import_cgd_extrato_ordem", operator: "+" },
    { sourceType: "financial_documents", operator: "?" },
    { sourceType: "unknown", operator: "+" },
    null,
  ]), [{ sourceType: "import_cgd_extrato_ordem", operator: "-" }]);
});

test("source changes reset pagination and trigger a silent workspace refresh", () => {
  const current = { candidateSourceType: "financial_documents", page: 4, loaded: true };
  const els = { financialReconciliationSource: { value: "import_cgd_extrato_ordem" } };
  const calls = [];
  const onSourceChange = new Function("clean", "els", "financialReconciliationState", "loadFinancialReconciliationWorkspace", `
${appFunctionSource("onFinancialReconciliationSourceChange")}
return onFinancialReconciliationSourceChange;
`)((value) => String(value || "").trim(), els, () => current, (options) => calls.push(options));

  onSourceChange();

  assert.deepEqual(current, { candidateSourceType: "import_cgd_extrato_ordem", page: 1, loaded: false });
  assert.deepEqual(calls, [{ silent: true }]);
});

test("workspace loading replays a source refresh requested while a prior load is in flight", async () => {
  const current = {
    candidateSourceType: "financial_documents",
    loaded: true,
    loading: false,
    reloadRequested: false,
    workspace: { sourceConfig: { sourceType: "financial_documents" } },
  };
  let resolveFirst;
  const requests = [];
  const api = (url) => {
    requests.push(url);
    if (requests.length === 1) return new Promise((resolve) => { resolveFirst = resolve; });
    return Promise.resolve({ sourceConfig: { sourceType: "import_cgd_extrato_ordem" }, reconciliation: { id: "reconciliation-1" } });
  };
  const loadSource = appFunctionSource("loadFinancialReconciliationWorkspace").replace(/^function /, "async function ");
  const loadWorkspace = new Function(
    "canAppFinancialReconciliation", "financialReconciliationState", "clean", "api",
    "buildFinancialReconciliationWorkspaceUrl", "normalizeFinancialReconciliationWorkspace",
    "financialReconciliationActiveRecord", "reconciliationRulesFor", "renderFinancialReconciliation",
    "setFinancialReconciliationStatus",
    `${loadSource}\nreturn loadFinancialReconciliationWorkspace;`,
  )(
    () => true,
    () => current,
    (value) => String(value || "").trim(),
    api,
    () => `${current.candidateSourceType}:${current.workspace?.reconciliation?.id || ""}`,
    (value) => ({ candidates: [], sourceConfig: {}, ...value }),
    () => null,
    () => [],
    () => {},
    () => {},
  );

  const firstLoad = loadWorkspace({ silent: true });
  current.candidateSourceType = "import_cgd_extrato_ordem";
  current.workspace = { ...current.workspace, reconciliation: { id: "reconciliation-1" } };
  await loadWorkspace({ silent: true });
  resolveFirst({ sourceConfig: { sourceType: "financial_documents" } });
  await firstLoad;

  assert.deepEqual(requests, ["financial_documents:", "import_cgd_extrato_ordem:reconciliation-1"]);
  assert.equal(current.workspace.sourceConfig.sourceType, "import_cgd_extrato_ordem");
  assert.equal(current.workspace.reconciliation.id, "reconciliation-1");
});

test("workspace loading reports failure to callers while retaining the shared history workspace", async () => {
  const retainedWorkspace = { sourceConfig: { sourceType: "financial_documents" }, history: [{ id: "history-1" }] };
  const current = {
    candidateSourceType: "financial_documents",
    loaded: true,
    loading: false,
    reloadRequested: false,
    workspace: retainedWorkspace,
  };
  const statuses = [];
  const loadWorkspace = new Function(
    "canAppFinancialReconciliation", "financialReconciliationState", "clean", "api",
    "buildFinancialReconciliationWorkspaceUrl", "normalizeFinancialReconciliationWorkspace",
    "financialReconciliationActiveRecord", "reconciliationRulesFor", "renderFinancialReconciliation",
    "setFinancialReconciliationStatus",
    `${appFunctionSource("loadFinancialReconciliationWorkspace").replace(/^function /, "async function ")}\nreturn loadFinancialReconciliationWorkspace;`,
  )(
    () => true,
    () => current,
    (value) => String(value || "").trim(),
    async () => { throw new Error("history unavailable"); },
    () => "workspace-url",
    (value) => value,
    () => null,
    () => [],
    () => {},
    (message, tone) => statuses.push({ message, tone }),
  );

  const refreshed = await loadWorkspace({ silent: true });

  assert.equal(refreshed, false);
  assert.strictEqual(current.workspace, retainedWorkspace);
  assert.match(statuses.at(-1).message, /failed to load reconciliation data/i);
  assert.equal(statuses.at(-1).tone, "error");
});

test("successful actions refresh the selected source after preserving the returned reconciliation", async () => {
  const current = {
    candidateSourceType: "import_cgd_extrato_ordem",
    pendingAction: "",
    workspace: { sourceConfig: { sourceType: "import_cgd_extrato_ordem" }, candidates: [{ id: "bank-1" }] },
  };
  const returnedReconciliation = { id: "reconciliation-1", matching_source_rules: [{ sourceType: "import_cgd_extrato_ordem", operator: "-" }] };
  let refresh;
  const actionSource = appFunctionSource("runFinancialReconciliationAction").replace(/^function /, "async function ");
  const runAction = new Function(
    "financialReconciliationState", "api", "normalizeFinancialReconciliationWorkspace", "clean",
    "loadFinancialReconciliationWorkspace", "showToast", "setFinancialReconciliationStatus",
    "renderFinancialReconciliation",
    `${actionSource}\nreturn runFinancialReconciliationAction;`,
  )(
    () => current,
    async () => ({ sourceConfig: { sourceType: "financial_documents" }, reconciliation: returnedReconciliation, items: [], audit: [], history: [] }),
    (value) => ({ candidates: [], items: [], audit: [], history: [], ...value }),
    (value) => String(value || "").trim(),
    async (options) => { refresh = { options, workspace: current.workspace }; },
    () => {},
    () => {},
    () => {},
  );

  await runAction({ action: "complete", reconciliationId: "reconciliation-1" });

  assert.deepEqual(refresh.options, { silent: true });
  assert.equal(refresh.workspace.sourceConfig.sourceType, "import_cgd_extrato_ordem");
  assert.equal(refresh.workspace.reconciliation, returnedReconciliation);
});

test("row actions use the source rendered with that row while a source reload is pending", () => {
  class FakeHTMLElement {
    constructor(dataset) {
      this.dataset = dataset;
      this.disabled = false;
    }

    closest(selector) {
      return selector === "button[data-financial-reconciliation-row-action]" ? this : null;
    }
  }

  const actions = [];
  const current = { candidateSourceType: "import_cgd_cartao_credito" };
  const onRowsClick = new Function(
    "HTMLElement", "clean", "financialReconciliationState", "runFinancialReconciliationAction", "financialReconciliationActiveRecord",
    `${appFunctionSource("onFinancialReconciliationRowsClick")}\nreturn onFinancialReconciliationRowsClick;`,
  )(
    FakeHTMLElement,
    (value) => String(value || "").trim(),
    () => current,
    (payload) => actions.push(payload),
    () => null,
  );

  onRowsClick({ target: new FakeHTMLElement({ financialReconciliationRowAction: "start", sourceId: "document-1", sourceType: "financial_documents" }) });

  assert.deepEqual(actions, [{ action: "start", sourceType: "financial_documents", sourceId: "document-1" }]);
});

test("reconciliation density rules are scoped to workbench and eligible records", () => {
  assert.match(html, /class="card financial-reconciliation-workbench-card"/);
  assert.match(html, /class="card financial-reconciliation-eligible-card"/);
  assert.match(css, /\.financial-reconciliation-workbench-card h2,\s*\.financial-reconciliation-eligible-card h2\s*\{\s*font-size:\s*\.94rem;/);
  assert.match(css, /\.financial-reconciliation-filters label\s*\{\s*font-size:\s*\.72rem;/);
  assert.match(css, /\.financial-reconciliation-filters input,\s*\.financial-reconciliation-filters select\s*\{\s*font-size:\s*\.70rem;/);
  assert.match(css, /\.financial-reconciliation-table th\s*\{\s*font-size:\s*\.70rem;\s*padding:\s*\.54rem;/);
  assert.match(css, /\.financial-reconciliation-table td\s*\{\s*font-size:\s*\.74rem;\s*padding:\s*\.54rem;/);
  assert.match(css, /\.financial-reconciliation-table button\s*\{\s*font-size:\s*\.72rem;/);
  assert.match(css, /@media \(max-width:\s*768px\)\s*\{\s*\.financial-reconciliation-filters input,\s*\.financial-reconciliation-filters select\s*\{\s*font-size:\s*16px;/);
  assert.match(css, /\.financial-reconciliation-table th\s*\{\s*font-size:\s*\.70rem;\s*padding:\s*\.46rem;/);
  assert.match(css, /\.financial-reconciliation-table td\s*\{\s*font-size:\s*\.70rem;\s*padding:\s*\.46rem;/);
  assert.match(css, /\.financial-reconciliation-table button\s*\{\s*font-size:\s*\.70rem;/);
  assert.match(appMain, /<th class="financial-reconciliation-action"><\/th><th class="financial-reconciliation-date">Date<\/th>/);
  assert.match(appMain, /<td class="financial-reconciliation-action"><button/);
  assert.match(appMain, /<td class="financial-reconciliation-date">\$\{escape\(formatDateOnly\(row\.source_date\) \|\| "-"\)\}<\/td>/);
  assert.match(appMain, /<td class="financial-reconciliation-detail">/);
  assert.match(appMain, /<td class="financial-reconciliation-amount">/);
  assert.match(appMain, /<td class="financial-reconciliation-status-cell">/);
  assert.match(css, /\.financial-reconciliation-action\s*\{\s*width:\s*4\.5rem;/);
  assert.match(css, /\.financial-reconciliation-date\s*\{\s*width:\s*6rem;/);
  assert.match(css, /\.financial-reconciliation-amount\s*\{\s*width:\s*5\.8rem;/);
  assert.match(css, /\.financial-reconciliation-status-cell\s*\{\s*width:\s*6\.8rem;/);
  assert.match(css, /overflow-wrap:\s*anywhere;/);
  assert.match(css, /#financial-reconciliation-start\s*\{\s*font-size:\s*\.84rem;/);
  assert.match(css, /\.financial-reconciliation-table button\s*\{\s*font-size:\s*\.72rem;/);
  assert.match(css, /\.financial-reconciliation-table \.financial-reconciliation-status\s*\{\s*font-size:\s*\.70rem;/);
  assert.match(appMain, /function financialReconciliationItemDetails\(item\)/);
  assert.match(appMain, /\[clean\(item\.source_date\) \? formatDateOnly\(item\.source_date\) : "", clean\(item\.supplier\), clean\(item\.description\)\]\.filter\(Boolean\)\.join\(" · "\)/);
  assert.match(appMain, /class="financial-reconciliation-item-details"/);
  assert.match(css, /\.financial-reconciliation-item-details\s*\{\s*font-size:\s*\.68rem;/);
  assert.match(css, /\.financial-reconciliation-current\s*\{\s*font-size:\s*\.86rem;/);
  assert.match(css, /\.financial-reconciliation-current h3\s*\{[\s\S]*font-size:\s*\.82rem;/);
  assert.match(css, /\.financial-reconciliation-completion\s*\{[\s\S]*display:\s*grid;[\s\S]*gap:\s*\.35rem;/);
  assert.match(css, /\.financial-reconciliation-completion textarea\s*\{[\s\S]*min-height:\s*4\.5rem;/);
  assert.match(css, /\.financial-reconciliation-items li\s*\{[\s\S]*column-gap:\s*\.45rem;[\s\S]*row-gap:\s*\.12rem;/);
  assert.match(css, /\.financial-reconciliation-item-details\s*\{[\s\S]*line-height:\s*1\.15;[\s\S]*margin-top:\s*0;/);
  assert.match(css, /\.financial-reconciliation-summary\s*\{[\s\S]*grid-template-columns:\s*auto 1fr auto;[\s\S]*align-items:\s*center;/);
  assert.match(css, /\.financial-reconciliation-record-count\s*\{[\s\S]*text-align:\s*center;[\s\S]*white-space:\s*nowrap;/);
  assert.doesNotMatch(css, /\.financial-reconciliation-summary p\s*\{/);
});

test("Current reconciliation summary shows status record count and short difference", () => {
  assert.equal(
    financialReconciliationSummaryMarkup("started", 2, 0),
    '<div class="financial-reconciliation-summary"><span class="financial-reconciliation-status">started</span><strong class="financial-reconciliation-record-count">#records: 2</strong><strong class="financial-reconciliation-difference">Dif: 0.00 €</strong></div>',
  );
  assert.equal(
    financialReconciliationSummaryMarkup("complete", 3, -10),
    '<div class="financial-reconciliation-summary"><span class="financial-reconciliation-status">complete</span><strong class="financial-reconciliation-record-count">#records: 3</strong><strong class="financial-reconciliation-difference financial-reconciliation-forced-difference">Dif: -10.00 €</strong></div>',
  );
});

test("Started and Complete Current summaries omit source prose and count every locked source", () => {
  const items = [
    { source_type: "financial_documents", source_id: "doc-1", amount_snapshot: 10 },
    { source_type: "import_cgd_extrato_ordem", source_id: "bank-1", amount_snapshot: -10 },
  ];
  const started = renderCurrentSummary({ status: "started", difference: 0, items });
  const complete = renderCurrentSummary({ status: "complete", difference: 0, items });
  for (const markup of [started, complete]) {
    assert.match(markup, /#records: 2/);
    assert.match(markup, /Dif: 0\.00 €/);
    assert.doesNotMatch(markup, /Financial Documents with CGD Bank Statement/);
    assert.doesNotMatch(markup, /Difference:/);
  }
});

test("Current reconciliation renders its structured origin beside status", () => {
  const markup = renderCurrentSummary({
    status: "complete",
    difference: 0,
    items: [],
    origin: "automatic",
    automaticTrigger: "manual",
  });
  assert.match(markup, /financial-reconciliation-summary[\s\S]*class="origin">Automatic<\/span>[\s\S]*#records: 0/);
});

test("history source text uses raw ordered source aggregates", () => {
  const record = {
    sourceSummary: [
      { sourceType: "financial_documents", recordCount: 4, amountTotal: 450 },
      { sourceType: "import_cgd_extrato_ordem", recordCount: 4, amountTotal: -450 },
    ],
  };
  assert.equal(
    historySourceHelpers.financialReconciliationHistorySourceText(record),
    "Financial Documents (#4; 450.00 €), CGD Bank Statement (#4; -450.00 €)",
  );
});

test("history source text distinguishes missing malformed and empty summaries", () => {
  assert.equal(historySourceHelpers.financialReconciliationHistorySourceText({}), "Source details unavailable");
  assert.equal(historySourceHelpers.financialReconciliationHistorySourceText({ sourceSummary: "invalid" }), "Source details unavailable");
  assert.equal(historySourceHelpers.financialReconciliationHistorySourceText({ sourceSummary: [] }), "No records");
  assert.deepEqual(
    historySourceHelpers.financialReconciliationHistorySourceSummary({
      sourceSummary: [
        { sourceType: "financial_documents", recordCount: 2, amountTotal: 20 },
        { sourceType: "unknown", recordCount: 1, amountTotal: 10 },
        { sourceType: "financial_documents", recordCount: 2, amountTotal: 20 },
        { sourceType: "import_cgd_extrato_ordem", recordCount: 0, amountTotal: -20 },
      ],
    }),
    [{ sourceType: "financial_documents", recordCount: 2, amountTotal: 20 }],
  );
});

test("history source summary rejects coercible non-numbers and preserves numeric zero totals", () => {
  const malformedEntries = [
    { sourceType: "financial_documents", recordCount: "2", amountTotal: 20 },
    { sourceType: "financial_documents", recordCount: true, amountTotal: 20 },
    { sourceType: "financial_documents", recordCount: 2, amountTotal: null },
    { sourceType: "financial_documents", recordCount: 2, amountTotal: "" },
    { sourceType: "financial_documents", recordCount: 2, amountTotal: false },
  ];
  for (const entry of malformedEntries) {
    assert.deepEqual(
      historySourceHelpers.financialReconciliationHistorySourceSummary({ sourceSummary: [entry] }),
      [],
    );
  }
  assert.deepEqual(
    historySourceHelpers.financialReconciliationHistorySourceSummary({
      sourceSummary: [{ sourceType: "financial_documents", recordCount: 2, amountTotal: 0 }],
    }),
    [{ sourceType: "financial_documents", recordCount: 2, amountTotal: 0 }],
  );
});

function renderHistory(record, selectedReconciliationId = "") {
  const els = { financialReconciliationHistoryRows: { innerHTML: "" } };
  const current = { selectedReconciliationId, workspace: { history: [record] } };
  const render = new Function(
    "financialReconciliationState",
    "clean",
    "escape",
    "formatDateTimeShort",
    "financialReconciliationHistorySourceText",
    "financialReconciliationStatusMarkup",
    "financialReconciliationDifference",
    "formatMoney",
    "els",
    "financialReconciliationOriginMarkup",
    `${appFunctionSource("renderFinancialReconciliationHistory")}
     return renderFinancialReconciliationHistory;`,
  )(
    () => current,
    (value) => String(value || "").trim(),
    (value) => String(value),
    () => "2026-08-12 10:00",
    historySourceHelpers.financialReconciliationHistorySourceText,
    (status) => status === "complete" ? "Complete" : "Started",
    (value) => Number(value.difference_amount),
    (value) => `${Number(value).toFixed(2)} €`,
    els,
    (value) => `<span class="origin">${value.origin === "automatic" ? `Automatic Â· ${value.automaticTrigger === "scheduled" ? "Scheduled" : "Manual"}` : "User"}</span>`,
  );
  render();
  return els.financialReconciliationHistoryRows.innerHTML;
}

test("history renders one wrapping source summary and preserves row behavior", () => {
  const markup = renderHistory({
    id: "rec-1",
    created_at: "2026-08-12T10:00:00Z",
    status: "complete",
    difference_amount: 0,
    sourceSummary: [
      { sourceType: "financial_documents", recordCount: 4, amountTotal: 450 },
      { sourceType: "import_cgd_extrato_ordem", recordCount: 4, amountTotal: -450 },
    ],
  }, "rec-1");

  assert.match(html, /<th>Created<\/th><th>Source<\/th><th>Origin<\/th><th>Status<\/th><th>Difference<\/th><th><\/th>/);
  assert.doesNotMatch(html, /<th>Base source<\/th>|<th>Matching sources<\/th>/);
  assert.match(markup, /class="selected"/);
  assert.match(markup, /class="financial-reconciliation-history-source"/);
  assert.match(markup, /Financial Documents \(#4; 450\.00 €\), CGD Bank Statement \(#4; -450\.00 €\)/);
  assert.match(markup, /Complete/);
  assert.match(markup, /0\.00 €/);
  assert.match(markup, /data-financial-reconciliation-select="rec-1">Open<\/button>/);
});

test("dedicated history renders source and destination totals with completion comments", () => {
  const current = {
    rows: [{
      id: "rec-history-1",
      created_at: "2026-08-22T12:00:00Z",
      base_source_type: "financial_documents",
      origin: "automatic",
      automaticTrigger: "scheduled",
      status: "complete",
      difference_amount: 0,
      totalRecords: 5,
      sourceAmountTotal: 450,
      destinationAmountTotal: -450,
      completionComment: "Matched by scheduled rule",
      sourceSummary: [
        { sourceType: "financial_documents", recordCount: 2, amountTotal: 450 },
        { sourceType: "import_cgd_extrato_ordem", recordCount: 2, amountTotal: -400 },
        { sourceType: "import_cgd_cartao_credito", recordCount: 1, amountTotal: -50 },
      ],
    }],
    page: 1,
    pageSize: 50,
    total: 1,
    loading: false,
    error: "",
  };
  const els = {
    financialReconciliationHistorySearchRows: { innerHTML: "" },
    financialReconciliationHistoryCount: { textContent: "" },
    financialReconciliationHistoryPage: { textContent: "" },
    financialReconciliationHistoryPrevious: { disabled: false },
    financialReconciliationHistoryNext: { disabled: false },
    financialReconciliationHistorySearchStatus: { textContent: "" },
  };
  const render = new Function(
    "clean", "escape", "formatMoney", "formatDateTimeShort", "financialReconciliationSourceLabel",
    "financialReconciliationHistorySearchState", "financialReconciliationOriginMarkup",
    "financialReconciliationStatusMarkup", "financialReconciliationDifference", "FINANCIAL_RECONCILIATION_SOURCES", "els",
    `${appFunctionSource("financialReconciliationHistorySourceSummary")}
     ${appFunctionSource("financialReconciliationHistorySummaryMarkup")}
     ${appFunctionSource("renderFinancialReconciliationHistorySearch")}
     return renderFinancialReconciliationHistorySearch;`,
  )(
    (value) => String(value ?? "").trim(),
    (value) => String(value),
    (value) => `${Number(value).toFixed(2)} €`,
    () => "2026-08-22 12:00",
    (source) => ({ financial_documents: "Financial Documents", import_cgd_extrato_ordem: "CGD Bank Statement", import_cgd_cartao_credito: "CGD Credit Card" }[source] || source),
    () => current,
    () => "Automatic",
    () => "Complete",
    (record) => Number(record.difference_amount),
    { financial_documents: {}, import_cgd_extrato_ordem: {}, import_cgd_cartao_credito: {} },
    els,
  );
  render();
  const markup = els.financialReconciliationHistorySearchRows.innerHTML;
  assert.match(markup, /Financial Documents <small>\(#2; 450\.00 €\)<\/small>/);
  assert.match(markup, /CGD Bank Statement <small>\(#2; -400\.00 €\)<\/small>/);
  assert.match(markup, /CGD Credit Card <small>\(#1; -50\.00 €\)<\/small>/);
  assert.match(markup, />5<\/td>/);
  assert.match(markup, /Matched by scheduled rule/);
  assert.match(markup, /data-financial-reconciliation-select="rec-history-1"/);
});

test("origin presentation is backward compatible and distinguishes automatic triggers", () => {
  const originPresentation = new Function(
    "clean",
    `${appFunctionSource("financialReconciliationOriginPresentation")}
     return financialReconciliationOriginPresentation;`,
  )((value) => String(value ?? "").trim());

  assert.deepEqual(originPresentation({}), {
    key: "user",
    label: "User",
    className: "financial-reconciliation-origin--user",
  });
  assert.deepEqual(originPresentation({ origin: "user" }), {
    key: "user",
    label: "User",
    className: "financial-reconciliation-origin--user",
  });
  assert.deepEqual(originPresentation({ origin: "automatic", automaticTrigger: "manual" }), {
    key: "automatic-manual",
    label: "Automatic \u00b7 Manual",
    className: "financial-reconciliation-origin--automatic-manual",
  });
  assert.deepEqual(originPresentation({ origin: "automatic", automatic_trigger: "scheduled" }), {
    key: "automatic-scheduled",
    label: "Automatic \u00b7 Scheduled",
    className: "financial-reconciliation-origin--automatic-scheduled",
  });
});

test("reconciliation item details omit empty fields", () => {
  assert.equal(financialReconciliationItemDetails({ source_date: "", supplier: " ", description: "" }), "");
});

test("manual reconciliation renders source LOVs and escapes every option", () => {
  const financialReconciliationFilterFieldMarkup = new Function(
    "clean",
    "escape",
    `${appFunctionSource("financialReconciliationFilterFieldMarkup")}\nreturn financialReconciliationFilterFieldMarkup;`,
  )(
    (value) => String(value ?? "").trim(),
    (value) => String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;"),
  );
  const markup = financialReconciliationFilterFieldMarkup(
    "payment",
    "Visa",
    ["Banco", "Visa", '<script data-x="1">'],
  );
  assert.match(markup, /^<label class="financial-reconciliation-filter-payment">Payment<select/);
  assert.match(markup, /<option value="">All payments<\/option>/);
  assert.match(markup, /<option value="Visa" selected>Visa<\/option>/);
  assert.match(markup, /&lt;script data-x=&quot;1&quot;&gt;/);
  assert.doesNotMatch(markup, /<script/);

  for (const [field, allLabel] of [
    ["payment", "All payments"],
    ["category", "All categories"],
    ["account", "All accounts"],
  ]) {
    assert.match(
      financialReconciliationFilterFieldMarkup(field, "", []),
      new RegExp(`<option value="">${allLabel}<\\/option>`),
    );
  }
});

test("manual reconciliation combines Financial Documents description and supplier search", () => {
  const financialReconciliationFilterFieldMarkup = new Function(
    "clean",
    "escape",
    `${appFunctionSource("financialReconciliationFilterFieldMarkup")}\nreturn financialReconciliationFilterFieldMarkup;`,
  )(
    (value) => String(value ?? "").trim(),
    (value) => String(value),
  );
  const combined = financialReconciliationFilterFieldMarkup("description", "Acme", null, "financial_documents");
  const ordinary = financialReconciliationFilterFieldMarkup("description", "guest", null, "import_fdm_accounts");

  assert.match(combined, /^<label class="financial-reconciliation-filter-description">Description \/ Supplier Search<input type="search"/);
  assert.match(combined, /placeholder="Search description or supplier"/);
  assert.match(ordinary, /^<label class="financial-reconciliation-filter-description">Description<input type="search"/);
  assert.match(ordinary, /placeholder="Search description"/);
});

test("manual reconciliation normalizes source filter options without resorting them", () => {
  const financialReconciliationFilterOptions = new Function(
    "clean",
    `${appFunctionSource("financialReconciliationFilterOptions")}\nreturn financialReconciliationFilterOptions;`,
  )((value) => String(value ?? "").trim());

  assert.deepEqual(financialReconciliationFilterOptions({
    sourceConfig: {
      filterOptions: {
        payment: [" Visa ", "", "Visa", null, "Banco"],
        category: "not-an-array",
      },
    },
  }), { payment: ["Visa", "Banco"] });
  assert.deepEqual(financialReconciliationFilterOptions({ sourceConfig: {} }), {});
});

test("manual reconciliation request filters omit stale Financial Documents accounts and retain FDM fields", () => {
  const state = {
    workspace: { sourceConfig: {} },
    filters: {
      dateFrom: "2026-01-01",
      dateTo: "",
      amountMin: "",
      amountMax: "",
      description: "EDP",
      supplier: "stale-supplier",
      payment: "Visa",
      account: "stale-account",
      category: "Food",
    },
  };
  const requestFilters = new Function(
    "financialReconciliationState",
    "clean",
    "currentFinancialReconciliationFilters",
    `${appConstantSource("FINANCIAL_RECONCILIATION_FILTER_FIELDS")}
     ${appFunctionSource("financialReconciliationRequestFilters")}
     return financialReconciliationRequestFilters;`,
  )(
    () => state,
    (value) => String(value ?? "").trim(),
    () => state.filters,
  );

  assert.deepEqual(requestFilters("financial_documents"), {
    dateFrom: "2026-01-01",
    dateTo: "",
    amountMin: "",
    amountMax: "",
    description: "EDP",
    payment: "Visa",
    category: "Food",
  });
  assert.deepEqual(requestFilters("import_fdm_accounts"), {
    dateFrom: "2026-01-01",
    dateTo: "",
    amountMin: "",
    amountMax: "",
    description: "EDP",
    account: "stale-account",
    category: "Food",
  });
});

test("manual reconciliation selects do not schedule a duplicate delayed reload", () => {
  let scheduled = 0;
  const onFinancialReconciliationFilterInput = new Function(
    "window",
    "clean",
    "onFinancialReconciliationFilterChange",
    `let financialReconciliationFilterTimer = 0;\n${appFunctionSource("onFinancialReconciliationFilterInput")}\nreturn onFinancialReconciliationFilterInput;`,
  )(
    {
      clearTimeout: () => {},
      setTimeout: () => {
        scheduled += 1;
        return scheduled;
      },
    },
    (value) => String(value ?? "").trim(),
    () => {},
  );

  onFinancialReconciliationFilterInput({ target: { tagName: "SELECT" } });
  assert.equal(scheduled, 0);
  onFinancialReconciliationFilterInput({ target: { tagName: "INPUT" } });
  assert.equal(scheduled, 1);
});

test("manual reconciliation renders declared LOVs as selects in the dynamic filter row", () => {
  const current = {
    workspace: {
      sourceConfig: {
        sourceType: "financial_documents",
        filterFields: ["description", "payment", "category"],
        filterOptions: { payment: ["Visa"], category: [] },
      },
    },
    filters: {},
  };
  const els = { financialReconciliationDynamicFilters: { innerHTML: "" } };
  const renderFinancialReconciliationFilters = new Function(
    "financialReconciliationState",
    "els",
    "document",
    `${appFunctionSource("clean")}
     ${appFunctionSource("escape")}
     ${appFunctionSource("currentFinancialReconciliationFilters")}
     ${appFunctionSource("financialReconciliationFilterOptions")}
     ${appFunctionSource("financialReconciliationFilterFieldMarkup")}
     ${appFunctionSource("renderFinancialReconciliationFilters")}
     return renderFinancialReconciliationFilters;`,
  )(
    () => current,
    els,
    { querySelectorAll: () => [] },
  );

  renderFinancialReconciliationFilters();

  const markup = els.financialReconciliationDynamicFilters.innerHTML;
  assert.match(markup, /financial-reconciliation-filter-description">Description \/ Supplier Search<input type="search"/);
  assert.match(markup, /placeholder="Search description or supplier"/);
  assert.match(markup, /financial-reconciliation-filter-payment">Payment<select data-financial-reconciliation-filter="payment"><option value="">All payments<\/option><option value="Visa">Visa<\/option><\/select>/);
  assert.match(markup, /financial-reconciliation-filter-category">Category<select data-financial-reconciliation-filter="category"><option value="">All categories<\/option><\/select>/);
  assert.doesNotMatch(markup, /financial-reconciliation-filter-supplier/);
});

test("reconciliation item details order date supplier and description", () => {
  assert.equal(
    financialReconciliationItemDetails({ source_date: "2026-08-10T09:30:00Z", supplier: "Acme Supplies", description: "August invoice" }),
    "2026-08-10 · Acme Supplies · August invoice",
  );
});
