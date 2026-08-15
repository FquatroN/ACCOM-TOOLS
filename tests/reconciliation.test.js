const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const {
  MIN_RECONCILIATION_DATE,
  calculateDifference,
  mapRpcError,
  normalizeReconciliationRules,
  normalizeRuleSnapshot,
  normalizeSourceType,
  validateMutation,
  validateWorkspaceQuery,
} = require("../api/_reconciliation");

const oldestFirstMigrationPath = path.join(
  __dirname,
  "..",
  "supabase-migrations",
  "2026-08-12-financial-reconciliation-oldest-first-candidates.sql",
);
const oldestFirstMigration = fs.existsSync(oldestFirstMigrationPath)
  ? fs.readFileSync(oldestFirstMigrationPath, "utf8")
  : "";

const historySourceSummaryMigrationPath = path.join(
  __dirname,
  "..",
  "supabase-migrations",
  "2026-08-12-financial-reconciliation-history-source-summary.sql",
);
const historySourceSummaryMigration = fs.existsSync(historySourceSummaryMigrationPath)
  ? fs.readFileSync(historySourceSummaryMigrationPath, "utf8")
  : "";

const filterLovMigrationPath = path.join(
  __dirname,
  "..",
  "supabase-migrations",
  "2026-08-15-financial-reconciliation-workspace-filter-lovs.sql",
);
const filterLovMigration = fs.existsSync(filterLovMigrationPath)
  ? fs.readFileSync(filterLovMigrationPath, "utf8")
  : "";

test("workspace filter LOV migration preserves the RPC while adding source metadata", () => {
  assert.match(filterLovMigration, /pg_get_functiondef\('public\.get_financial_reconciliation_workspace\(uuid,text,jsonb,integer,integer\)'::regprocedure\)/);
  assert.match(filterLovMigration, /'filterOptions',v_filter_options/);
  assert.match(filterLovMigration, /jsonb_build_array\('dateFrom','dateTo','amountMin','amountMax','description','supplier','payment','category'\)/);
  assert.doesNotMatch(
    filterLovMigration,
    /jsonb_build_array\('dateFrom','dateTo','amountMin','amountMax','description','supplier','payment','account','category'\)/,
  );
  assert.match(filterLovMigration, /s\.supplier ilike[\s\S]+or s\.supplier_nif ilike/);
  assert.match(filterLovMigration, /select distinct btrim\(payment\)/i);
  assert.match(filterLovMigration, /select distinct btrim\(category\)/i);
  assert.match(filterLovMigration, /select distinct btrim\(account\)/i);
  assert.match(filterLovMigration, /order by lower\(option_value\),option_value/i);
  assert.match(filterLovMigration, /btrim\(s\.payment\) = p_filters->>'payment'/);
  assert.match(filterLovMigration, /btrim\(s\.account\) = p_filters->>'account'/);
  assert.match(filterLovMigration, /btrim\(s\.category\) = p_filters->>'category'/);
  assert.match(filterLovMigration, /security definer/i);
  assert.match(filterLovMigration, /revoke all on function public\.get_financial_reconciliation_workspace\(uuid,text,jsonb,integer,integer\) from public, anon, authenticated;/);
  assert.match(filterLovMigration, /grant execute on function public\.get_financial_reconciliation_workspace\(uuid,text,jsonb,integer,integer\) to service_role;/);

  const rpcSmoke = fs.readFileSync(path.join(__dirname, "reconciliation-rpc.smoke.sql"), "utf8");
  assert.equal(
    (rpcSmoke.match(/2026-08-15-financial-reconciliation-workspace-filter-lovs\.sql/g) || []).length,
    2,
  );
  assert.match(rpcSmoke, /Supplier Search did not match Supplier NIF/);
  assert.match(rpcSmoke, /Supplier Search did not match Supplier Name/);
  assert.match(rpcSmoke, /payment LOV is not trimmed, distinct, and nonblank/);
  assert.match(rpcSmoke, /FDM Account LOV is not trimmed, distinct, and nonblank/);
  assert.match(rpcSmoke, /Payment filter did not match trimmed stored data/);
  assert.match(rpcSmoke, /Financial Documents category filter did not match trimmed stored data/);
  assert.match(rpcSmoke, /FDM Account filter did not match trimmed stored data/);
  assert.match(rpcSmoke, /FDM category filter did not match trimmed stored data/);
  assert.match(rpcSmoke, /Financial Documents LOVs did not include ineligible, out-of-date, and locked values/);
  assert.match(rpcSmoke, /FDM LOVs did not include ineligible, out-of-date, and locked values/);
  assert.match(rpcSmoke, /Workspace RPC security definer or search path is incorrect/);
  assert.match(rpcSmoke, /Workspace RPC permissions are not service-role-only/);
});

test("history source-summary migration safely enriches the workspace function", () => {
  assert.match(historySourceSummaryMigration, /pg_get_functiondef\('public\.get_financial_reconciliation_workspace\(uuid,text,jsonb,integer,integer\)'::regprocedure\)/);
  assert.match(historySourceSummaryMigration, /'sourceSummary'/);
  assert.match(historySourceSummaryMigration, /count\(\*\)/i);
  assert.match(historySourceSummaryMigration, /sum\(i\.amount_snapshot\)/i);
  assert.match(historySourceSummaryMigration, /jsonb_array_elements_text\(h\.matching_source_types\)\s+with ordinality/i);
  assert.match(historySourceSummaryMigration, /old_history_count = 1\s+and new_history_count = 0/is);
  assert.match(historySourceSummaryMigration, /old_history_count = 0\s+and new_history_count = 1/is);
  assert.match(historySourceSummaryMigration, /unexpected reconciliation workspace function definition/i);
});

test("oldest-first migration targets the workspace function and declares deterministic clauses", () => {
  assert.match(oldestFirstMigration, /pg_get_functiondef\('public\.get_financial_reconciliation_workspace\(uuid,text,jsonb,integer,integer\)'::regprocedure\)/);
  assert.match(oldestFirstMigration, /new_page_order constant text := \$\$order by source_date asc, id asc offset v_offset limit p_page_size\$\$/i);
  assert.match(oldestFirstMigration, /new_json_order constant text := \$\$order by x\.source_date asc, x\.id asc\$\$/i);
  assert.match(oldestFirstMigration, /old_page_count = 1\s+and new_page_count = 0\s+and old_json_count = 1\s+and new_json_count = 0/is);
  assert.match(oldestFirstMigration, /old_page_count = 0\s+and new_page_count = 1\s+and old_json_count = 0\s+and new_json_count = 1/is);
  assert.match(oldestFirstMigration, /unexpected reconciliation workspace function definition/i);
  assert.match(oldestFirstMigration, /could not verify deterministic oldest-first candidate ordering/i);
});

test("Financial Documents-led groups sum all sources", () => {
  assert.equal(calculateDifference("financial_documents", [
    { sourceType: "import_cgd_extrato_ordem", operator: "+" },
    { sourceType: "import_cgd_cartao_credito", operator: "+" },
    { sourceType: "import_fdm_accounts", operator: "+" },
  ], [
    { sourceType: "financial_documents", amountSnapshot: 100 },
    { sourceType: "import_cgd_extrato_ordem", amountSnapshot: -40 },
    { sourceType: "import_cgd_cartao_credito", amountSnapshot: -30 },
    { sourceType: "import_fdm_accounts", amountSnapshot: -30 },
  ]), 0);
});

test("FDM-led groups subtract bank values", () => {
  assert.equal(calculateDifference("import_fdm_accounts", [{ sourceType: "import_cgd_extrato_ordem", operator: "-" }], [
    { sourceType: "import_fdm_accounts", amountSnapshot: 42 },
    { sourceType: "import_cgd_extrato_ordem", amountSnapshot: 42 },
  ]), 0);
});

test("calculations reject items outside the selected reconciliation mode", () => {
  assert.throws(() => calculateDifference("import_fdm_accounts", [{ sourceType: "import_cgd_extrato_ordem", operator: "-" }], [
    { sourceType: "import_fdm_accounts", amountSnapshot: 42 },
    { sourceType: "import_cgd_extrato_ordem", amountSnapshot: 42 },
    { sourceType: "financial_documents", amountSnapshot: 10 },
  ]), /not allowed/i);
});

test("card-led groups use the selected approved pairing", () => {
  assert.equal(calculateDifference("import_cgd_cartao_credito", [{ sourceType: "financial_documents", operator: "+" }], [
    { sourceType: "import_cgd_cartao_credito", amountSnapshot: -30.115 },
    { sourceType: "financial_documents", amountSnapshot: 30.11 },
  ]), 0);
  assert.equal(calculateDifference("import_cgd_cartao_credito", [{ sourceType: "import_cgd_extrato_ordem", operator: "+" }], [
    { sourceType: "import_cgd_cartao_credito", amountSnapshot: -30 },
    { sourceType: "import_cgd_extrato_ordem", amountSnapshot: 30 },
  ]), 0);
});

test("bank-led groups use the matching source formula", () => {
  assert.equal(calculateDifference("import_cgd_extrato_ordem", [{ sourceType: "financial_documents", operator: "+" }], [
    { sourceType: "import_cgd_extrato_ordem", amountSnapshot: -20 },
    { sourceType: "financial_documents", amountSnapshot: 20 },
  ]), 0);
  assert.equal(calculateDifference("import_cgd_extrato_ordem", [{ sourceType: "import_cgd_cartao_credito", operator: "+" }], [
    { sourceType: "import_cgd_extrato_ordem", amountSnapshot: -20 },
    { sourceType: "import_cgd_cartao_credito", amountSnapshot: 20 },
  ]), 0);
  assert.equal(calculateDifference("import_cgd_extrato_ordem", [{ sourceType: "import_fdm_accounts", operator: "-" }], [
    { sourceType: "import_cgd_extrato_ordem", amountSnapshot: 20 },
    { sourceType: "import_fdm_accounts", amountSnapshot: 20 },
  ]), 0);
});

test("rules preserve independent directions with different reverse operators", () => {
  assert.deepEqual(normalizeReconciliationRules([
    { baseSourceType: "financial_documents", matchingSourceType: "import_cgd_extrato_ordem", operator: "+" },
    { baseSourceType: "import_cgd_extrato_ordem", matchingSourceType: "financial_documents", operator: "-" },
  ]), [
    { baseSourceType: "financial_documents", matchingSourceType: "import_cgd_extrato_ordem", operator: "+" },
    { baseSourceType: "import_cgd_extrato_ordem", matchingSourceType: "financial_documents", operator: "-" },
  ]);
});

test("rules reject self-pairs, duplicates, and unknown operators", () => {
  assert.throws(() => normalizeReconciliationRules([{ baseSourceType: "financial_documents", matchingSourceType: "financial_documents", operator: "+" }]), /different/i);
  assert.throws(() => normalizeReconciliationRules([
    { baseSourceType: "financial_documents", matchingSourceType: "import_fdm_accounts", operator: "+" },
    { baseSourceType: "financial_documents", matchingSourceType: "import_fdm_accounts", operator: "-" },
  ]), /duplicate/i);
  assert.throws(() => normalizeReconciliationRules([{ baseSourceType: "financial_documents", matchingSourceType: "import_fdm_accounts", operator: "*" }]), /operator/i);
});

test("difference uses the base amount and each directional snapshot operator", () => {
  const rules = normalizeRuleSnapshot("financial_documents", [
    { sourceType: "import_cgd_extrato_ordem", operator: "+" },
    { sourceType: "import_cgd_cartao_credito", operator: "+" },
  ]);
  assert.equal(calculateDifference("financial_documents", rules, [
    { sourceType: "financial_documents", amountSnapshot: 100 },
    { sourceType: "import_cgd_extrato_ordem", amountSnapshot: -60 },
    { sourceType: "import_cgd_cartao_credito", amountSnapshot: 40 },
  ]), 80);
});

test("source names are restricted to the four configured sources", () => {
  assert.throws(() => normalizeSourceType("financial-documents"), /source type/i);
});

test("force complete requires comment", () => {
  assert.throws(() => validateMutation("force_complete", {
    reconciliationId: "a", comment: " ",
  }), /comment is required/i);
});

test("mutations require their reconciliation and source identifiers", () => {
  assert.throws(() => validateMutation("add_item", {
    sourceType: "financial_documents", sourceId: "source-1",
  }), /reconciliation id is required/i);
  assert.throws(() => validateMutation("start", {
    sourceType: "financial_documents",
  }), /source id is required/i);
});

test("a Start action contains no client-selected matching sources", () => {
  assert.deepEqual(validateMutation("start", {
    sourceType: "financial_documents", sourceId: "record-1",
  }), { action: "start", reconciliationId: "", sourceType: "financial_documents", sourceId: "record-1", comment: "" });
});

test("workspace validation enforces source names and page bounds", () => {
  assert.deepEqual(validateWorkspaceQuery({
    source_type: "financial_documents",
    page: "2",
    page_size: "100",
    filters: '{"dateFrom":"2026-01-01"}',
  }), {
    reconciliationId: "",
    sourceType: "financial_documents",
    page: 2,
    pageSize: 100,
    filters: { dateFrom: "2026-01-01" },
  });
  assert.throws(
    () => validateWorkspaceQuery({ source_type: "financial_documents", page_size: "101" }),
    /page size/i,
  );
  assert.throws(
    () => validateWorkspaceQuery({ source_type: "financial_documents", page: "0" }),
    /page/i,
  );
  assert.throws(
    () => validateWorkspaceQuery({ source_type: "financial_documents", page: "101" }),
    /page/i,
  );
});

test("workspace query no longer accepts a matching-source mode", () => {
  assert.deepEqual(validateWorkspaceQuery({ source_type: "import_cgd_extrato_ordem", page: "1" }), {
    reconciliationId: "", sourceType: "import_cgd_extrato_ordem", page: 1, pageSize: 50, filters: {},
  });
});

test("workspace rejects invalid source and oversized page", () => {
  assert.throws(
    () => validateWorkspaceQuery({ source_type: "wrong", page_size: "101" }),
    /source type|page size/i,
  );
});

test("workspace rejects malformed date and amount filters before SQL casts", () => {
  assert.throws(
    () => validateWorkspaceQuery({ source_type: "financial_documents", filters: '{"dateFrom":"2026-99-99"}' }),
    /valid iso date/i,
  );
  assert.throws(
    () => validateWorkspaceQuery({ source_type: "financial_documents", filters: '{"amountMin":"not-a-number"}' }),
    /valid number/i,
  );
});

test("RPC error mapping makes source locks conflicts and validation errors client-safe", () => {
  assert.equal(mapRpcError(new Error("This record is already reconciled.")).statusCode, 409);
  assert.equal(mapRpcError(new Error("Only started reconciliations can be edited or completed.")).statusCode, 400);
});

test("minimum reconciliation date is the documented 2026 eligibility floor", () => {
  assert.equal(MIN_RECONCILIATION_DATE, "2026-01-01");
});

test("failed browse-source reload preserves the active reconciliation workspace", () => {
  const source = fs.readFileSync(path.join(__dirname, "..", "app-main.js"), "utf8");
  assert.match(
    source,
    /const previousWorkspace = current\.workspace;\s*const previousCandidateSourceType = clean\(previousWorkspace\?\.sourceConfig\?\.sourceType\) \|\| current\.candidateSourceType;/,
  );
  assert.match(
    source,
    /catch \(error\) \{\s*if \(!current\.reloadRequested\) \{\s*current\.workspace = previousWorkspace \|\| normalizeFinancialReconciliationWorkspace\(\{\}\);\s*current\.candidateSourceType = previousCandidateSourceType;/,
  );
});
