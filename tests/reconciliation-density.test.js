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

const financialReconciliationItemDetails = new Function(`${appFunctionSource("clean")}\n${appFunctionSource("formatDateOnly")}\n${appFunctionSource("financialReconciliationItemDetails")}\nreturn financialReconciliationItemDetails;`)();
const reconciliationSettingsDestinationHelpers = new Function("appFeatures", "lastMainView", `
const state = { lastMainView, access: { appFeatures } };
${appFunctionSource("clean")}
${appFunctionSource("canApp")}
${appFunctionSource("canAppFinancialDocs")}
${appFunctionSource("canAppImportData")}
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
});

test("reconciliation item details omit empty fields", () => {
  assert.equal(financialReconciliationItemDetails({ source_date: "", supplier: " ", description: "" }), "");
});

test("reconciliation item details order date supplier and description", () => {
  assert.equal(
    financialReconciliationItemDetails({ source_date: "2026-08-10T09:30:00Z", supplier: "Acme Supplies", description: "August invoice" }),
    "2026-08-10 · Acme Supplies · August invoice",
  );
});
