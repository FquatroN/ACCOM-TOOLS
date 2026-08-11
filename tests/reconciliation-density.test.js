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
  const bodyStart = appMain.indexOf("{", start);
  let depth = 0;
  for (let index = bodyStart; index < appMain.length; index += 1) {
    if (appMain[index] === "{") depth += 1;
    if (appMain[index] === "}") depth -= 1;
    if (depth === 0) return appMain.slice(start, index + 1);
  }
  throw new Error(`Could not extract ${name} from app-main.js`);
}

const financialReconciliationItemDetails = new Function(`${appFunctionSource("clean")}\n${appFunctionSource("formatDateOnly")}\n${appFunctionSource("financialReconciliationItemDetails")}\nreturn financialReconciliationItemDetails;`)();
const reconciliationSettingsAppDestinationFor = new Function("capabilities", `
function canAppFinancialReconciliation() { return Boolean(capabilities.financialReconciliation); }
function canAppFinancialDocs() { return Boolean(capabilities.financialDocs); }
function canAppImportData() { return Boolean(capabilities.importData); }
function preferredMainAppView() { return capabilities.permittedMainView || ""; }
function canUseGuestsBi() { return Boolean(capabilities.guestsBi); }
function canUseBookingsBi() { return Boolean(capabilities.bookingsBi); }
function canUseFinancialBi() { return Boolean(capabilities.financialBi); }
function canUseSalesBi() { return Boolean(capabilities.salesBi); }
${appFunctionSource("reconciliationSettingsAppDestination")}
return reconciliationSettingsAppDestination();
`);

test("settings exposes a reconciliation rule editor", () => {
  assert.match(html, /id="settings-menu-financial-reconciliation"/);
  assert.match(html, /id="settings-view-financial-reconciliation"/);
  assert.match(html, /id="financial-reconciliation-settings-base-source"/);
  assert.match(html, /id="financial-reconciliation-settings-rules-body"/);
  assert.match(html, /id="financial-reconciliation-settings-save"/);
});

test("reconciliation settings authorizes and validates a complete replacement before deleting rules", () => {
  assert.match(reconciliationSettingsApi, /requireFeature\(req, "settings", "financial-reconciliation"\)/);
  assert.match(supabaseApi, /SETTINGS_FEATURES = \[[^\]]*"financial-reconciliation"/);
  assert.match(appMain, /SETTINGS_FEATURE_OPTIONS = \[[^\]]*"financial-reconciliation"/);
  const validation = reconciliationSettingsApi.indexOf("const input = normalizeReconciliationRules");
  const replacement = reconciliationSettingsApi.indexOf('restQuery("financial_reconciliation_source_rules", { method: "DELETE" })');
  assert.ok(validation >= 0 && validation < replacement, "validation must complete before replacement begins");
});

test("reconciliation settings keeps its contextual tab and backs out only to authorized app views", () => {
  assert.match(appMain, /previousView === "financial-reconciliation" && canSettings\("financial-reconciliation"\)/);
  assert.match(appMain, /function reconciliationSettingsAppDestination\(\)/);
  assert.match(appMain, /els\.closeSettingsFinancialReconciliation\?\.addEventListener\("click", closeReconciliationSettings\)/);
  assert.doesNotMatch(appMain, /closeSettingsFinancialReconciliation\?\.addEventListener\([^\n]*"communications"/);
});

test("reconciliation settings destination gives a settings-only administrator no app route", () => {
  assert.equal(reconciliationSettingsAppDestinationFor({}), "");
  assert.equal(reconciliationSettingsAppDestinationFor({ financialReconciliation: true }), "financial-reconciliation");
  assert.equal(reconciliationSettingsAppDestinationFor({ financialDocs: true }), "financial-docs");
  assert.equal(reconciliationSettingsAppDestinationFor({ importData: true }), "import-data");
  assert.equal(reconciliationSettingsAppDestinationFor({ guestsBi: true }), "guests-bi");
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
