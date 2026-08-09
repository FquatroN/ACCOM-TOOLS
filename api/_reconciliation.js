const MIN_RECONCILIATION_DATE = "2026-01-01";

const SOURCE_TYPES = Object.freeze([
  "financial_documents",
  "import_fdm_accounts",
  "import_cgd_cartao_credito",
  "import_cgd_extrato_ordem",
]);

const MATCHING_SOURCE_TYPES = Object.freeze({
  financial_documents: [
    "import_fdm_accounts",
    "import_cgd_cartao_credito",
    "import_cgd_extrato_ordem",
  ],
  import_fdm_accounts: ["import_cgd_extrato_ordem"],
  import_cgd_cartao_credito: ["financial_documents", "import_cgd_extrato_ordem"],
  import_cgd_extrato_ordem: [
    "financial_documents",
    "import_fdm_accounts",
    "import_cgd_cartao_credito",
  ],
});

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

function normalizeMatchingSourceTypes(baseSourceType, matchingSourceTypes) {
  const base = normalizeSourceType(baseSourceType);
  if (!Array.isArray(matchingSourceTypes) || matchingSourceTypes.length === 0) {
    throw inputError("At least one matching source type is required.");
  }
  if (matchingSourceTypes.length > 3) {
    throw inputError("Matching source types are invalid.");
  }

  const normalized = matchingSourceTypes.map(normalizeSourceType);
  if (new Set(normalized).size !== normalized.length) {
    throw inputError("Matching source types must not contain duplicates.");
  }
  if (base !== "financial_documents" && normalized.length !== 1) {
    throw inputError("Non-document bases require exactly one matching source type.");
  }
  for (const sourceType of normalized) {
    if (!MATCHING_SOURCE_TYPES[base].includes(sourceType)) {
      throw inputError(`Matching source type '${sourceType}' is not allowed for '${base}'.`);
    }
  }
  return normalized;
}

function validateReconciliationMode(baseSourceType, matchingSourceTypes) {
  return {
    baseSourceType: normalizeSourceType(baseSourceType),
    matchingSourceTypes: normalizeMatchingSourceTypes(baseSourceType, matchingSourceTypes),
  };
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

function calculateDifference(baseSourceType, matchingSourceTypes, items) {
  const { baseSourceType: base, matchingSourceTypes: matching } = validateReconciliationMode(
    baseSourceType,
    matchingSourceTypes,
  );
  if (!Array.isArray(items)) throw inputError("Reconciliation items must be an array.");

  const totals = Object.fromEntries(SOURCE_TYPES.map((sourceType) => [sourceType, 0]));
  for (const item of items) {
    const sourceType = normalizeSourceType(item && item.sourceType);
    if (sourceType !== base && !matching.includes(sourceType)) {
      throw inputError("Item source type is not allowed for the selected reconciliation mode.");
    }
    totals[sourceType] += amountFor(item);
  }

  const document = totals.financial_documents;
  const fdm = totals.import_fdm_accounts;
  const card = totals.import_cgd_cartao_credito;
  const bank = totals.import_cgd_extrato_ordem;

  if (base === "financial_documents") return roundMoney(document + fdm + card + bank);
  if (base === "import_fdm_accounts") return roundMoney(fdm - bank);
  if (base === "import_cgd_cartao_credito") {
    return roundMoney(card + (matching[0] === "financial_documents" ? document : bank));
  }
  if (matching[0] === "import_fdm_accounts") return roundMoney(fdm - bank);
  return roundMoney(bank + (matching[0] === "financial_documents" ? document : card));
}

function identifier(value, label, required) {
  if (value === undefined || value === null || String(value).trim() === "") {
    if (required) throw inputError(`${label} is required.`);
    return "";
  }
  return String(value).trim();
}

function parseMatchingSourceTypes(value) {
  if (Array.isArray(value)) return value;
  if (typeof value !== "string") return [];
  const text = value.trim();
  if (!text) return [];
  if (text.startsWith("[")) {
    try {
      return JSON.parse(text);
    } catch {
      throw inputError("Matching source types must be a valid array.");
    }
  }
  return text.split(",").map((sourceType) => sourceType.trim());
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
  return filters;
}

function validateWorkspaceQuery(query) {
  const input = query && typeof query === "object" ? query : {};
  const sourceType = normalizeSourceType(input.source_type || input.sourceType || "financial_documents");
  const matching = parseMatchingSourceTypes(input.matching_source_types || input.matchingSourceTypes);
  const matchingSourceTypes = normalizeMatchingSourceTypes(
    sourceType,
    matching.length ? matching : MATCHING_SOURCE_TYPES[sourceType].slice(0, 1),
  );
  return {
    reconciliationId: identifier(input.reconciliation_id || input.reconciliationId, "Reconciliation ID", false),
    sourceType,
    matchingSourceTypes,
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
  if (action === "start") {
    const mode = validateReconciliationMode(
      input.baseSourceType || input.base_source_type,
      parseMatchingSourceTypes(input.matchingSourceTypes || input.matching_source_types),
    );
    if (mode.baseSourceType !== sourceType) {
      throw inputError("Start source type must match the base source type.");
    }
    result.baseSourceType = mode.baseSourceType;
    result.matchingSourceTypes = mode.matchingSourceTypes;
  }
  return result;
}

function mapRpcError(error) {
  const mapped = error instanceof Error ? error : new Error("Unexpected server error.");
  if (mapped.statusCode) return mapped;
  const message = mapped.message || "Unexpected server error.";
  if (/already reconciled|unique|conflict/i.test(message)) mapped.statusCode = 409;
  else if (/invalid|required|not allowed|exactly one|must be|cannot|only|zero difference/i.test(message)) mapped.statusCode = 400;
  else mapped.statusCode = 500;
  return mapped;
}

module.exports = {
  MIN_RECONCILIATION_DATE,
  SOURCE_TYPES,
  calculateDifference,
  mapRpcError,
  normalizeMatchingSourceTypes,
  normalizeSourceType,
  validateMutation,
  validateReconciliationMode,
  validateWorkspaceQuery,
};
