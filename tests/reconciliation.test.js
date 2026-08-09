const test = require("node:test");
const assert = require("node:assert/strict");
const {
  MIN_RECONCILIATION_DATE,
  calculateDifference,
  normalizeMatchingSourceTypes,
  normalizeSourceType,
  validateMutation,
  validateWorkspaceQuery,
} = require("../api/_reconciliation");

test("Financial Documents-led groups sum all sources", () => {
  assert.equal(calculateDifference("financial_documents", [
    "import_cgd_extrato_ordem", "import_cgd_cartao_credito", "import_fdm_accounts",
  ], [
    { sourceType: "financial_documents", amountSnapshot: 100 },
    { sourceType: "import_cgd_extrato_ordem", amountSnapshot: -40 },
    { sourceType: "import_cgd_cartao_credito", amountSnapshot: -30 },
    { sourceType: "import_fdm_accounts", amountSnapshot: -30 },
  ]), 0);
});

test("FDM-led groups subtract bank values", () => {
  assert.equal(calculateDifference("import_fdm_accounts", ["import_cgd_extrato_ordem"], [
    { sourceType: "import_fdm_accounts", amountSnapshot: 42 },
    { sourceType: "import_cgd_extrato_ordem", amountSnapshot: 42 },
  ]), 0);
});

test("card-led groups use the selected approved pairing", () => {
  assert.equal(calculateDifference("import_cgd_cartao_credito", ["financial_documents"], [
    { sourceType: "import_cgd_cartao_credito", amountSnapshot: -30.115 },
    { sourceType: "financial_documents", amountSnapshot: 30.11 },
  ]), 0);
  assert.equal(calculateDifference("import_cgd_cartao_credito", ["import_cgd_extrato_ordem"], [
    { sourceType: "import_cgd_cartao_credito", amountSnapshot: -30 },
    { sourceType: "import_cgd_extrato_ordem", amountSnapshot: 30 },
  ]), 0);
});

test("bank-led groups use the matching source formula", () => {
  assert.equal(calculateDifference("import_cgd_extrato_ordem", ["financial_documents"], [
    { sourceType: "import_cgd_extrato_ordem", amountSnapshot: -20 },
    { sourceType: "financial_documents", amountSnapshot: 20 },
  ]), 0);
  assert.equal(calculateDifference("import_cgd_extrato_ordem", ["import_cgd_cartao_credito"], [
    { sourceType: "import_cgd_extrato_ordem", amountSnapshot: -20 },
    { sourceType: "import_cgd_cartao_credito", amountSnapshot: 20 },
  ]), 0);
  assert.equal(calculateDifference("import_cgd_extrato_ordem", ["import_fdm_accounts"], [
    { sourceType: "import_cgd_extrato_ordem", amountSnapshot: 20 },
    { sourceType: "import_fdm_accounts", amountSnapshot: 20 },
  ]), 0);
});

test("non-document bases permit one matching source", () => {
  assert.throws(() => normalizeMatchingSourceTypes("import_cgd_extrato_ordem", [
    "financial_documents", "import_cgd_cartao_credito",
  ]), /exactly one/i);
});

test("matching source types reject unsupported pairings", () => {
  assert.throws(
    () => normalizeMatchingSourceTypes("import_fdm_accounts", ["financial_documents"]),
    /not allowed/i,
  );
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
    baseSourceType: "financial_documents", matchingSourceTypes: ["import_fdm_accounts"],
    sourceType: "financial_documents",
  }), /source id is required/i);
});

test("workspace validation enforces source names and page bounds", () => {
  assert.deepEqual(validateWorkspaceQuery({
    source_type: "financial_documents",
    matching_source_types: "import_fdm_accounts,import_cgd_cartao_credito",
    page: "2",
    page_size: "100",
    filters: '{"dateFrom":"2026-01-01"}',
  }), {
    reconciliationId: "",
    sourceType: "financial_documents",
    matchingSourceTypes: ["import_fdm_accounts", "import_cgd_cartao_credito"],
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
});

test("minimum reconciliation date is the documented 2026 eligibility floor", () => {
  assert.equal(MIN_RECONCILIATION_DATE, "2026-01-01");
});
