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
