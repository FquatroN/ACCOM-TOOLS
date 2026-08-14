const MIN_RECONCILIATION_DATE = "2026-01-01";

const SOURCE_TYPES = Object.freeze([
  "financial_documents",
  "import_fdm_accounts",
  "import_cgd_cartao_credito",
  "import_cgd_extrato_ordem",
]);

const MUTATION_ACTIONS = new Set([
  "start", "add_item", "remove_item", "complete", "force_complete", "reopen", "delete",
]);

function inputError(message) {
  const error = new Error(message);
  error.statusCode = 400;
  return error;
}

function normalizeSourceType(value) {
  if (typeof value !== "string" || !SOURCE_TYPES.includes(value)) {
    throw inputError("Source type is invalid.");
  }
  return value;
}

function normalizeOperator(value) {
  if (value !== "+" && value !== "-") throw inputError("Rule operator must be '+' or '-'.");
  return value;
}

function normalizeReconciliationRules(value) {
  if (!Array.isArray(value)) throw inputError("Reconciliation rules must be an array.");
  const seen = new Set();
  return value.map((rule) => {
    const baseSourceType = normalizeSourceType(rule?.baseSourceType || rule?.base_source_type);
    const matchingSourceType = normalizeSourceType(rule?.matchingSourceType || rule?.matching_source_type);
    if (baseSourceType === matchingSourceType) throw inputError("Rule sources must be different.");
    const key = `${baseSourceType}:${matchingSourceType}`;
    if (seen.has(key)) throw inputError("Duplicate reconciliation rule.");
    seen.add(key);
    return { baseSourceType, matchingSourceType, operator: normalizeOperator(rule?.operator) };
  });
}

function normalizeRuleSnapshot(baseSourceType, value) {
  const base = normalizeSourceType(baseSourceType);
  const rows = Array.isArray(value) ? value : [];
  const seen = new Set();
  return rows.map((rule) => {
    const sourceType = normalizeSourceType(rule?.sourceType || rule?.source_type);
    if (sourceType === base) throw inputError("Snapshot source must differ from the base source.");
    if (seen.has(sourceType)) throw inputError("Duplicate reconciliation rule snapshot.");
    seen.add(sourceType);
    return { sourceType, operator: normalizeOperator(rule?.operator) };
  });
}

function roundMoney(value) {
  const rounded = Math.round((Number(value) + Number.EPSILON) * 100) / 100;
  return Object.is(rounded, -0) ? 0 : rounded;
}

function amountFor(item) {
  const amount = Number(item && item.amountSnapshot);
  if (!Number.isFinite(amount)) {
    throw inputError("Item amount snapshot is invalid.");
  }
  return amount;
}

function calculateDifference(baseSourceType, matchingSourceRules, items) {
  const base = normalizeSourceType(baseSourceType);
  const rules = normalizeRuleSnapshot(base, matchingSourceRules);
  const operators = new Map(rules.map((rule) => [rule.sourceType, rule.operator]));
  return roundMoney((Array.isArray(items) ? items : []).reduce((total, item) => {
    const sourceType = normalizeSourceType(item?.sourceType);
    const sign = sourceType === base ? 1 : operators.get(sourceType) === "-" ? -1 : operators.get(sourceType) === "+" ? 1 : null;
    if (sign === null) throw inputError("Item source type is not allowed for this reconciliation.");
    return total + (sign * amountFor(item));
  }, 0));
}

function identifier(value, label, required) {
  if (value === undefined || value === null || String(value).trim() === "") {
    if (required) throw inputError(`${label} is required.`);
    return "";
  }
  return String(value).trim();
}

function positiveInteger(value, label, maximum) {
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < 1 || parsed > maximum) {
    throw inputError(`${label} must be between 1 and ${maximum}.`);
  }
  return parsed;
}

function parseFilters(value) {
  if (value === undefined || value === null || value === "") return {};
  let filters = value;
  if (typeof value === "string") {
    try {
      filters = JSON.parse(value);
    } catch {
      throw inputError("Filters must be valid JSON.");
    }
  }
  if (!filters || Array.isArray(filters) || typeof filters !== "object") {
    throw inputError("Filters must be an object.");
  }
  for (const key of ["dateFrom", "dateTo"]) {
    const date = filters[key];
    if (date === undefined || date === null || String(date).trim() === "") continue;
    const text = String(date).trim();
    const parsed = new Date(`${text}T00:00:00.000Z`);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(text) || Number.isNaN(parsed.getTime()) || parsed.toISOString().slice(0, 10) !== text) {
      throw inputError(`${key === "dateFrom" ? "Date from" : "Date to"} must be a valid ISO date.`);
    }
  }
  for (const key of ["amountMin", "amountMax"]) {
    const amount = filters[key];
    if (amount === undefined || amount === null || String(amount).trim() === "") continue;
    if (!/^[+-]?(?:\d+(?:\.\d+)?|\.\d+)$/.test(String(amount).trim())) {
      throw inputError(`${key === "amountMin" ? "Amount from" : "Amount to"} must be a valid number.`);
    }
  }
  return filters;
}

function validateWorkspaceQuery(query) {
  const input = query && typeof query === "object" ? query : {};
  const reconciliationId = identifier(input.reconciliation_id || input.reconciliationId, "Reconciliation ID", false);
  const sourceType = normalizeSourceType(input.source_type || input.sourceType || "financial_documents");
  return {
    reconciliationId,
    sourceType,
    page: input.page === undefined || input.page === "" ? 1 : positiveInteger(input.page, "Page", 100),
    pageSize: input.page_size === undefined || input.page_size === "" ? 50 : positiveInteger(input.page_size, "Page size", 100),
    filters: parseFilters(input.filters),
  };
}

function validateMutation(action, payload) {
  if (!MUTATION_ACTIONS.has(action)) throw inputError("Reconciliation action is invalid.");
  const input = payload && typeof payload === "object" ? payload : {};
  const reconciliationId = identifier(input.reconciliationId || input.reconciliation_id, "Reconciliation ID", action !== "start");
  const sourceAction = action === "start" || action === "add_item" || action === "remove_item";
  const sourceType = sourceAction
    ? normalizeSourceType(input.sourceType || input.source_type)
    : "";
  const sourceId = identifier(input.sourceId || input.source_id, "Source ID", sourceAction);
  const comment = identifier(input.comment, "Comment", action === "force_complete");

  const result = { action, reconciliationId, sourceType, sourceId, comment };
  return result;
}

function mapRpcError(error) {
  const mapped = error instanceof Error ? error : new Error("Unexpected server error.");
  const message = mapped.message || "Unexpected server error.";
  if (/already reconciled|unique_violation|duplicate key|conflict/i.test(message)) {
    mapped.statusCode = 409;
  } else if (/scheduled slot|stale proposal/i.test(message)) {
    mapped.statusCode = 409;
  } else if (mapped.statusCode) {
    mapped.statusCode = mapped.statusCode;
  } else if (/automatic rule|automation proposal|candidate limit|ambiguous|invalid|required|not allowed|exactly one|must be|cannot|only|zero difference|reconciliation item not found|reconciliation not found|eligible|started reconciliations|complete reconciliations/i.test(message)) {
    mapped.statusCode = 400;
  } else {
    mapped.statusCode = 500;
  }
  return mapped;
}

module.exports = {
  MIN_RECONCILIATION_DATE,
  SOURCE_TYPES,
  calculateDifference,
  mapRpcError,
  normalizeOperator,
  normalizeReconciliationRules,
  normalizeRuleSnapshot,
  normalizeSourceType,
  validateMutation,
  validateWorkspaceQuery,
};
