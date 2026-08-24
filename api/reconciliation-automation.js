const { cleanText, parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const { mapRpcError } = require("./_reconciliation");
const {
  AUTOMATIC_RULE_DISPLAY_NAMES,
  AUTOMATIC_RULE_VERSIONS,
  ADYEN_MONTHLY_RULE_KEY,
  BANK_RESERVATION_RULE_KEY,
  BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY,
  BANK_STATEMENT_RULE_KEY,
  CREDIT_CARD_AMOUNT_ONLY_RULE_KEY,
  CREDIT_CARD_RULE_KEY,
  MONTHLY_INCOME_RULE_KEY,
  normalizeAnalyzePayload,
  normalizeAutomationAction,
  normalizeContinueAnalysisPayload,
  normalizeExecutePayload,
  toAutomationPublicResult,
} = require("./_reconciliation-automation");

const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const DATE_PATTERN = /^\d{4}-\d{2}-\d{2}$/;
const RUN_FIELDS = new Set([
  "runId", "trigger", "scope", "status", "actor", "clientRequestId", "scheduledSlot",
  "batchId", "batchRuleKey", "batchRulePosition", "batchRuleCount", "definitions", "counts",
  "analysisCursorDate", "analysisCursorId", "analysisProcessed", "analysisTotal", "analysisErrorCode",
  "analysisErrorAt", "analysisUnit", "analysisComplete", "analysisCompletedAt", "startedAt",
  "finishedAt", "proposals",
]);
const DEFINITION_FIELDS = new Set([
  "ruleKey", "ruleVersion", "displayName", "priority", "differenceAllowed", "maxDifferenceDays",
  "destinationSourceType", "definition", "operator",
]);
const PROPOSAL_FIELDS = new Set([
  "id", "runId", "ruleKey", "ruleVersion", "baseSourceType", "baseSourceId", "baseSourceDate",
  "baseSnapshot", "items", "evidence", "candidateGroups", "groupingKey", "summarySnapshot",
  "calculatedDifference", "allowedDifference", "status", "reason", "signature", "reconciliationId",
  "createdAt", "updatedAt", "displayName",
]);
const COUNT_FIELDS = new Set([
  "bases", "proposed", "ambiguous", "skipped", "deselected", "executing", "completed", "stale", "failed",
]);
const CATALOG_FIELDS = new Set([
  "ruleKey", "ruleVersion", "displayName", "baseSourceType", "destinationSourceTypes",
  "logicDescription", "definition", "enabled", "allowManualExecution", "differenceAllowed",
  "maxDifferenceDays", "priority", "operator",
]);
const GROUPED_RULES = new Set([MONTHLY_INCOME_RULE_KEY, BANK_RESERVATION_RULE_KEY, ADYEN_MONTHLY_RULE_KEY]);
const CLASSIC_RULES = new Set([
  BANK_STATEMENT_RULE_KEY,
  CREDIT_CARD_RULE_KEY,
  BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY,
  CREDIT_CARD_AMOUNT_ONLY_RULE_KEY,
]);
const AMOUNT_ONLY_RULES = new Set([
  BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY,
  CREDIT_CARD_AMOUNT_ONLY_RULE_KEY,
]);
const RUN_STATUSES = new Set(["analyzing", "ready", "running", "completed", "partial", "failed"]);
const PROPOSAL_STATUSES = new Set([
  "proposed", "ambiguous", "skipped", "deselected", "executing", "completed", "stale", "failed",
]);
const GROUPED_BASE_SNAPSHOT_FIELDS = new Set(["sourceType", "sourceId", "sourceDate"]);
const BANK_GROUPED_SUMMARY_FIELDS = new Set([
  "classification", "reason", "candidateCount", "bankAnchorDate",
  "sourceCount", "sourceTotal", "destinationCount", "destinationTotal",
]);
const MONTHLY_GROUPED_SUMMARY_FIELDS = new Set([
  "calendarMonth", "sourceCount", "sourceTotal", "destinationCount", "destinationTotal",
]);
const GROUPED_STALE_REASONS = new Set([
  "analysis_population_changed", "operator_changed", "rule_snapshot_changed",
  "rule_version_changed", "source_snapshot_changed", "tolerance_changed",
]);
const MANAGED_RULE_CONTRACTS = Object.freeze({
  [BANK_STATEMENT_RULE_KEY]: {
    baseSourceType: "financial_documents",
    destinationSourceType: "import_cgd_extrato_ordem",
    logicDescription: "Payment must equal exactly Banco. A bank candidate must match at least one of three OR identity branches: compact document-number containment, document-description similarity, or supplier-to-bank-description word similarity. A base record is executable only when exactly one complete destination combination is valid; multiple combinations are reported as ambiguous and are never selected automatically.",
    definition: {
      baseSourceType: "financial_documents",
      destinationSourceTypes: ["import_cgd_extrato_ordem"],
      baseEligibility: { payment: { operator: "exact_text_equal", value: "Banco", caseSensitive: true, trim: false } },
      identityBranches: {
        document_number: { algorithm: "compact_containment" },
        description_similarity: { algorithm: "similarity" },
        supplier_similarity: { algorithm: "word_similarity" },
      },
      documentNumberMinimumCompactLength: 4,
      descriptionSimilarityThreshold: 0.60,
      supplierWordSimilarityThreshold: 0.70,
      maxDestinationRecords: 4,
      maxIdentityCandidatesPerBase: 12,
    },
  },
  [CREDIT_CARD_RULE_KEY]: {
    baseSourceType: "financial_documents",
    destinationSourceType: "import_cgd_cartao_credito",
    logicDescription: "Payment must equal exactly Visa. Each credit-card candidate must satisfy invoice containment, description similarity, or supplier word similarity. Exactly one one-to-four-record amount combination is executable.",
    definition: {
      baseEligibility: { payment: { operator: "exact_text_equal", value: "Visa", caseSensitive: true, trim: false } },
      identityBranches: {
        document_number: { algorithm: "symmetric_compact_containment" },
        description_similarity: { algorithm: "similarity" },
        supplier_similarity: { algorithm: "word_similarity" },
      },
      documentNumberMinimumCompactLength: 4,
      descriptionSimilarityThreshold: 0.55,
      supplierWordSimilarityThreshold: 0.60,
      maxDestinationRecords: 4,
      maxIdentityCandidatesPerBase: 12,
    },
  },
  [BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY]: {
    baseSourceType: "financial_documents",
    destinationSourceType: "import_cgd_extrato_ordem",
    logicDescription: "Payment must equal exactly Banco. Exactly one CGD Bank Account destination record must make the signed amounts sum to zero within the inclusive configured date window; identity fields and similarity are not used.",
    definition: {
      baseSourceType: "financial_documents",
      destinationSourceTypes: ["import_cgd_extrato_ordem"],
      baseEligibility: { payment: { operator: "exact_text_equal", value: "Banco", caseSensitive: true, trim: false } },
      matchingMode: "amount_only_one_to_one",
      fixedDifferenceAllowed: 0,
      maxDifferenceDays: { minimum: 0, maximum: 90, default: 1 },
      maxDestinationRecords: 1,
    },
  },
  [CREDIT_CARD_AMOUNT_ONLY_RULE_KEY]: {
    baseSourceType: "financial_documents",
    destinationSourceType: "import_cgd_cartao_credito",
    logicDescription: "Payment must equal exactly Visa. Exactly one CGD Credit Card destination record must make the signed amounts sum to zero within the inclusive configured date window; identity fields and similarity are not used.",
    definition: {
      baseSourceType: "financial_documents",
      destinationSourceTypes: ["import_cgd_cartao_credito"],
      baseEligibility: { payment: { operator: "exact_text_equal", value: "Visa", caseSensitive: true, trim: false } },
      matchingMode: "amount_only_one_to_one",
      fixedDifferenceAllowed: 0,
      maxDifferenceDays: { minimum: 0, maximum: 90, default: 1 },
      maxDestinationRecords: 1,
    },
  },
  [MONTHLY_INCOME_RULE_KEY]: {
    baseSourceType: "import_cgd_extrato_ordem",
    destinationSourceType: "import_fdm_accounts",
    logicDescription: "Every unlocked CGD Bank Statement POS VENDAS record is reconciled against every unlocked FDM Credit Card record in the same closed calendar month, except FDM records categorized as TransferOutToAccount; the difference is Bank Statement total minus FDM Accounts total.",
    definition: {
      matchingMode: "monthly_aggregate",
      sourceDescriptionPattern: "%POS VENDAS%",
      destinationAccount: "Credit Card",
      destinationExcludedCategory: "TransferOutToAccount",
      calendarGrouping: "closed_month",
      fixedMaxDifferenceDays: 31,
      eligibilityFloor: "2026-01-01",
      requiresNonNullAmount: true,
    },
  },
  [BANK_RESERVATION_RULE_KEY]: {
    baseSourceType: "import_fdm_accounts",
    destinationSourceType: "import_cgd_extrato_ordem",
    logicDescription: "Exactly one CGD Bank Statement record is matched to one through ten eligible FDM Bank Transfer records with opposite signed totals that equal zero exactly in integer cents within the inclusive configured date window.",
    definition: {
      strategy: "bounded_exact_combination",
      sourceAccount: "Bank Transfer",
      maxSourceRecords: 10,
      candidatePoolLimit: 60,
      stateLimit: 250000,
      evidenceGroupLimit: 12,
      amountMode: "signed_integer_cents",
      dateMode: "inclusive_days",
    },
  },
  [ADYEN_MONTHLY_RULE_KEY]: {
    baseSourceType: "import_cgd_extrato_ordem",
    destinationSourceType: "import_fdm_accounts",
    logicDescription: "Every eligible unlocked CGD Bank Statement and FDM Adyen record in the same closed calendar month forms one proposal; both sides are required and the signed difference must be within the configured allowance.",
    definition: {
      strategy: "closed_calendar_month",
      bankDescriptionContains: "Adyen",
      fdmAccount: "Adyen",
      requiresBothSides: true,
      monthMarkerDays: 31,
    },
  },
});

function inputError(message) {
  const error = new Error(message);
  error.statusCode = 400;
  return error;
}

function isPlainRecord(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function failUnexpected(label = "run") {
  throw new Error(`Unexpected reconciliation ${label} response.`);
}

function rejectUnsafeOwnData(value, seen = new Set()) {
  if (value === null || typeof value !== "object") return;
  if (seen.has(value)) failUnexpected();
  seen.add(value);
  if (Array.isArray(value)) {
    for (const nested of value) rejectUnsafeOwnData(nested, seen);
    return;
  }
  if (!isPlainRecord(value)) failUnexpected();
  for (const [key, nested] of Object.entries(value)) {
    if (/(?:^|_)(?:private|diagnostic|internal|stack|error_detail|error_summary)(?:_|$)/i.test(key)
      || /private|diagnostic|internal|stack/i.test(key)) failUnexpected();
    rejectUnsafeOwnData(nested, seen);
  }
}

function requireExactFields(value, fields, label = "run") {
  if (!isPlainRecord(value)) failUnexpected(label);
  const keys = Object.keys(value);
  if (keys.length !== fields.size || keys.some((key) => !fields.has(key))) failUnexpected(label);
}

function isTimestamp(value) {
  return typeof value === "string" && !Number.isNaN(Date.parse(value));
}

function canonicalJson(value) {
  if (value === null || typeof value === "string" || typeof value === "boolean") return JSON.stringify(value);
  if (typeof value === "number") return Number.isFinite(value) ? JSON.stringify(value) : null;
  if (Array.isArray(value)) {
    const entries = value.map(canonicalJson);
    return entries.some((entry) => entry === null) ? null : `[${entries.join(",")}]`;
  }
  if (!isPlainRecord(value)) return null;
  const entries = Object.keys(value).sort().map((key) => {
    const nested = canonicalJson(value[key]);
    return nested === null ? null : `${JSON.stringify(key)}:${nested}`;
  });
  return entries.some((entry) => entry === null) ? null : `{${entries.join(",")}}`;
}

function isDecimal(value) {
  return (typeof value === "number" && Number.isFinite(value))
    || (typeof value === "string" && /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(value));
}

function decimalCents(value) {
  const text = typeof value === "number" && Number.isFinite(value) ? String(value) : value;
  if (typeof text !== "string") return null;
  const match = /^(-?)(0|[1-9]\d{0,11})(?:\.(\d{1,2}))?$/.exec(text);
  if (!match) return null;
  const cents = BigInt(match[2]) * 100n
    + BigInt((match[3] || "").padEnd(2, "0"));
  return match[1] ? -cents : cents;
}

function isIsoDate(value) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) return false;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  if (year < 1 || month < 1 || month > 12 || day < 1) return false;
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  const days = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  return day <= days[month - 1];
}

function isoDayNumber(value) {
  if (!isIsoDate(value)) return null;
  return Date.parse(`${value}T00:00:00Z`) / 86400000;
}

function requireManagedTuple(ruleKey, ruleVersion, label = "run") {
  if (!Object.hasOwn(AUTOMATIC_RULE_VERSIONS, ruleKey)
    || ruleVersion !== AUTOMATIC_RULE_VERSIONS[ruleKey]) failUnexpected(label);
}

function requireManagedDefinition(ruleKey, definition, label) {
  const expected = MANAGED_RULE_CONTRACTS[ruleKey]?.definition;
  if (!expected || canonicalJson(definition) !== canonicalJson(expected)) failUnexpected(label);
}

function requireRuleValues(rule, label) {
  requireManagedTuple(rule.ruleKey, rule.ruleVersion, label);
  const contract = MANAGED_RULE_CONTRACTS[rule.ruleKey];
  if (rule.displayName !== AUTOMATIC_RULE_DISPLAY_NAMES[rule.ruleKey]
    || !contract
    || !Number.isSafeInteger(rule.priority) || rule.priority < 1
    || !isDecimal(rule.differenceAllowed)
    || !Number.isSafeInteger(rule.maxDifferenceDays)
    || rule.maxDifferenceDays < 0 || rule.maxDifferenceDays > 90
    || !isPlainRecord(rule.definition)) failUnexpected(label);
  requireManagedDefinition(rule.ruleKey, rule.definition, label);
  const expectedOperator = rule.ruleKey === MONTHLY_INCOME_RULE_KEY
    || rule.ruleKey === ADYEN_MONTHLY_RULE_KEY ? "-" : "+";
  if (rule.operator !== expectedOperator
    || (AMOUNT_ONLY_RULES.has(rule.ruleKey) && Number(rule.differenceAllowed) !== 0)
    || (rule.ruleKey === BANK_RESERVATION_RULE_KEY
      && (Number(rule.differenceAllowed) !== 0
        || rule.maxDifferenceDays < 0 || rule.maxDifferenceDays > 90))
    || (rule.ruleKey === ADYEN_MONTHLY_RULE_KEY
      && (Number(rule.differenceAllowed) < 0 || rule.maxDifferenceDays !== 31))
    || (rule.ruleKey === MONTHLY_INCOME_RULE_KEY
      && rule.maxDifferenceDays !== 31)) failUnexpected(label);
}

function requireManualRuleCatalog(value) {
  rejectUnsafeOwnData(value);
  const result = toAutomationPublicResult(value);
  requireExactFields(result, new Set(["rules"]), "catalog");
  if (!Array.isArray(result.rules) || result.rules.length > 7) failUnexpected("catalog");
  const seen = new Set();
  const priorities = new Set();
  for (const rule of result.rules) {
    requireExactFields(rule, CATALOG_FIELDS, "catalog");
    requireRuleValues(rule, "catalog");
    if (rule.enabled !== true || rule.allowManualExecution !== true
      || rule.baseSourceType !== MANAGED_RULE_CONTRACTS[rule.ruleKey]?.baseSourceType
      || !Array.isArray(rule.destinationSourceTypes) || rule.destinationSourceTypes.length !== 1
      || rule.destinationSourceTypes[0] !== MANAGED_RULE_CONTRACTS[rule.ruleKey]?.destinationSourceType
      || rule.logicDescription !== MANAGED_RULE_CONTRACTS[rule.ruleKey]?.logicDescription
      || seen.has(rule.ruleKey) || priorities.has(rule.priority)) failUnexpected("catalog");
    seen.add(rule.ruleKey);
    priorities.add(rule.priority);
  }
  return result;
}

function requireBankGroupedProposal(value, run) {
  const summary = value.summarySnapshot;
  requireExactFields(summary, BANK_GROUPED_SUMMARY_FIELDS);
  const sourceTotal = decimalCents(summary.sourceTotal);
  const destinationTotal = decimalCents(summary.destinationTotal);
  const calculatedDifference = decimalCents(value.calculatedDifference);
  const allowedDifference = decimalCents(value.allowedDifference);
  const unique = summary.classification === "proposed"
    && summary.reason === "unique_qualifying_combination";
  const candidateLimit = summary.classification === "ambiguous"
    && summary.reason === "candidate_limit";
  const multiple = summary.classification === "ambiguous"
    && summary.reason === "multiple_qualifying_combinations";
  const baseDay = isoDayNumber(value.baseSourceDate);
  const anchorDay = isoDayNumber(summary.bankAnchorDate);
  const maxDifferenceDays = run.definitions[0].maxDifferenceDays;
  if (!UUID_PATTERN.test(value.groupingKey)
    || value.groupingKey !== value.groupingKey.toLowerCase()
    || value.baseSourceType !== "import_fdm_accounts"
    || baseDay === null || anchorDay === null
    || value.baseSourceDate < "2026-01-01" || summary.bankAnchorDate < "2026-01-01"
    || Math.abs(baseDay - anchorDay) > maxDifferenceDays
    || !Number.isSafeInteger(summary.candidateCount)
    || !Number.isSafeInteger(summary.sourceCount)
    || summary.destinationCount !== 1
    || sourceTotal === null || destinationTotal === null
    || calculatedDifference !== 0n || allowedDifference !== 0n
    || (!unique && !candidateLimit && !multiple)) failUnexpected();

  if (unique) {
    if (summary.sourceCount < 1 || summary.sourceCount > 10
      || summary.candidateCount < summary.sourceCount || summary.candidateCount > 60
      || sourceTotal === 0n || destinationTotal === 0n
      || (sourceTotal < 0n) === (destinationTotal < 0n)
      || sourceTotal + destinationTotal !== 0n) failUnexpected();
  } else if (summary.sourceCount < 1 || summary.sourceCount > 60
    || summary.sourceCount > summary.candidateCount
    || summary.candidateCount < (multiple ? 2 : 1)
    || summary.candidateCount > (candidateLimit ? 61 : 60)
    || sourceTotal === 0n || destinationTotal === 0n
    || (sourceTotal < 0n) === (destinationTotal < 0n)) failUnexpected();

  const lifecycleIsValid = value.status === "proposed"
    ? unique && value.reason === "unique_qualifying_combination"
    : value.status === "ambiguous"
      ? value.reason === "overlapping_records" ? unique
        : (candidateLimit || multiple) && value.reason === summary.reason
      : new Set(["executing", "completed"]).has(value.status)
        ? unique && value.reason === ""
        : value.status === "deselected"
          ? unique && value.reason === "not_selected"
          : value.status === "failed"
            ? unique && value.reason === "execution_failed"
            : value.status === "stale" && GROUPED_STALE_REASONS.has(value.reason);
  if (!lifecycleIsValid
    || ((value.status === "completed") !== (value.reconciliationId !== null))) failUnexpected();
}

function requireAdyenGroupedProposal(value, run) {
  const summary = value.summarySnapshot;
  requireExactFields(summary, MONTHLY_GROUPED_SUMMARY_FIELDS);
  const sourceTotal = decimalCents(summary.sourceTotal);
  const destinationTotal = decimalCents(summary.destinationTotal);
  const calculatedDifference = decimalCents(value.calculatedDifference);
  const allowedDifference = decimalCents(value.allowedDifference);
  const configuredDifference = decimalCents(run.definitions[0].differenceAllowed);
  const calendarMonth = typeof summary.calendarMonth === "string"
    ? summary.calendarMonth : "";
  const groupingKey = calendarMonth.slice(0, 7);
  const currentMonth = new Date().toISOString().slice(0, 7);
  if (!isIsoDate(calendarMonth) || !calendarMonth.endsWith("-01")
    || value.groupingKey !== groupingKey || !/^\d{4}-\d{2}$/.test(value.groupingKey)
    || value.groupingKey < "2026-01" || value.groupingKey >= currentMonth
    || value.baseSourceType !== "import_cgd_extrato_ordem"
    || !isIsoDate(value.baseSourceDate)
    || !value.baseSourceDate.startsWith(`${value.groupingKey}-`)
    || !Number.isSafeInteger(summary.sourceCount) || summary.sourceCount < 1
    || !Number.isSafeInteger(summary.destinationCount) || summary.destinationCount < 1
    || sourceTotal === null || destinationTotal === null
    || calculatedDifference === null || allowedDifference === null
    || configuredDifference === null || allowedDifference < 0n
    || allowedDifference !== configuredDifference
    || calculatedDifference !== sourceTotal - destinationTotal) failUnexpected();

  const withinAllowance = calculatedDifference < 0n
    ? -calculatedDifference <= allowedDifference
    : calculatedDifference <= allowedDifference;
  const lifecycleIsValid = value.status === "proposed"
    ? value.reason === "" && withinAllowance
    : value.status === "ambiguous"
      ? value.reason === "monthly_difference_exceeded" && !withinAllowance
      : new Set(["executing", "completed"]).has(value.status)
        ? value.reason === "" && withinAllowance
        : value.status === "deselected"
          ? value.reason === "not_selected" && withinAllowance
          : value.status === "failed"
            ? value.reason === "execution_failed" && withinAllowance
            : value.status === "stale" && GROUPED_STALE_REASONS.has(value.reason);
  if (!lifecycleIsValid
    || ((value.status === "completed") !== (value.reconciliationId !== null))) failUnexpected();
}

function requireGroupedSummary(value, run) {
  if (value.ruleKey === BANK_RESERVATION_RULE_KEY) {
    requireBankGroupedProposal(value, run);
    return;
  }
  if (value.ruleKey === ADYEN_MONTHLY_RULE_KEY) {
    requireAdyenGroupedProposal(value, run);
    return;
  }
  requireExactFields(value.summarySnapshot, MONTHLY_GROUPED_SUMMARY_FIELDS);
  if (!isIsoDate(value.summarySnapshot.calendarMonth)
    || !value.summarySnapshot.calendarMonth.endsWith("-01")
    || !Number.isSafeInteger(value.summarySnapshot.sourceCount) || value.summarySnapshot.sourceCount < 0
    || !isDecimal(value.summarySnapshot.sourceTotal)
    || !Number.isSafeInteger(value.summarySnapshot.destinationCount)
    || value.summarySnapshot.destinationCount < 0
    || !isDecimal(value.summarySnapshot.destinationTotal)) failUnexpected();
}

function requireProposal(value, run) {
  requireExactFields(value, PROPOSAL_FIELDS);
  requireManagedTuple(value.ruleKey, value.ruleVersion);
  const groupingKeyIsValid = GROUPED_RULES.has(value.ruleKey)
    ? typeof value.groupingKey === "string" && Boolean(value.groupingKey)
    : CLASSIC_RULES.has(value.ruleKey) && value.groupingKey === null;
  if (!UUID_PATTERN.test(value.id) || value.runId !== run.runId
    || value.ruleKey !== run.definitions[0].ruleKey || value.ruleVersion !== run.definitions[0].ruleVersion
    || value.displayName !== AUTOMATIC_RULE_DISPLAY_NAMES[value.ruleKey]
    || typeof value.baseSourceType !== "string" || !value.baseSourceType
    || !UUID_PATTERN.test(value.baseSourceId) || !DATE_PATTERN.test(value.baseSourceDate)
    || !isPlainRecord(value.baseSnapshot) || !Array.isArray(value.items) || !Array.isArray(value.evidence)
    || !Array.isArray(value.candidateGroups) || !groupingKeyIsValid
    || !isPlainRecord(value.summarySnapshot) || !isDecimal(value.calculatedDifference)
    || !isDecimal(value.allowedDifference) || !PROPOSAL_STATUSES.has(value.status)
    || typeof value.reason !== "string" || typeof value.signature !== "string" || !value.signature
    || (value.reconciliationId !== null && !UUID_PATTERN.test(value.reconciliationId))
    || !isTimestamp(value.createdAt) || !isTimestamp(value.updatedAt)) failUnexpected();
  if (GROUPED_RULES.has(value.ruleKey)) {
    if (value.items.length || value.evidence.length || value.candidateGroups.length) failUnexpected();
    requireExactFields(value.baseSnapshot, GROUPED_BASE_SNAPSHOT_FIELDS);
    if (value.baseSnapshot.sourceType !== value.baseSourceType
      || value.baseSnapshot.sourceId !== value.baseSourceId
      || value.baseSnapshot.sourceDate !== value.baseSourceDate) failUnexpected();
    requireGroupedSummary(value, run);
  }
}

function requireManualRun(value, expected = {}, allowNull = false) {
  if (value === null && allowNull) return null;
  rejectUnsafeOwnData(value);
  const run = toAutomationPublicResult(value);
  requireExactFields(run, RUN_FIELDS);
  if (!UUID_PATTERN.test(run.runId) || run.trigger !== "manual" || run.scope !== "rule"
    || !RUN_STATUSES.has(run.status) || typeof run.actor !== "string" || !run.actor
    || !UUID_PATTERN.test(run.clientRequestId)
    || run.scheduledSlot !== null || run.batchId !== null || run.batchRuleKey !== null
    || run.batchRulePosition !== null || run.batchRuleCount !== null
    || !Array.isArray(run.definitions) || run.definitions.length !== 1
    || !isPlainRecord(run.counts) || !Array.isArray(run.proposals)
    || !Number.isSafeInteger(run.analysisProcessed) || run.analysisProcessed < 0
    || !Number.isSafeInteger(run.analysisTotal) || run.analysisTotal < 0
    || run.analysisProcessed > run.analysisTotal
    || typeof run.analysisComplete !== "boolean" || !isTimestamp(run.startedAt)
    || (run.finishedAt !== null && !isTimestamp(run.finishedAt))) failUnexpected();
  if (expected.runId && run.runId.toLowerCase() !== expected.runId.toLowerCase()) failUnexpected();
  if (expected.actor && run.actor !== expected.actor) failUnexpected();
  if (expected.ruleKey && run.definitions[0]?.ruleKey !== expected.ruleKey) failUnexpected();
  if (expected.clientRequestId
    && run.clientRequestId.toLowerCase() !== expected.clientRequestId.toLowerCase()) failUnexpected();
  requireExactFields(run.definitions[0], DEFINITION_FIELDS);
  requireRuleValues(run.definitions[0], "run");
  const destination = run.definitions[0].destinationSourceType;
  if (destination !== MANAGED_RULE_CONTRACTS[run.definitions[0].ruleKey]?.destinationSourceType) {
    failUnexpected();
  }
  for (const [key, count] of Object.entries(run.counts)) {
    if (!COUNT_FIELDS.has(key) || !Number.isSafeInteger(count) || count < 0) failUnexpected();
  }
  const cursorPair = run.analysisCursorDate === null && run.analysisCursorId === null;
  if (!cursorPair && (!DATE_PATTERN.test(run.analysisCursorDate) || !UUID_PATTERN.test(run.analysisCursorId))) {
    failUnexpected();
  }
  if (run.analysisComplete !== (run.analysisCompletedAt !== null)
    || (run.analysisCompletedAt !== null && !isTimestamp(run.analysisCompletedAt))
    || (run.analysisErrorCode !== null
      && (typeof run.analysisErrorCode !== "string" || !run.analysisErrorCode))
    || ((run.analysisErrorCode === null) !== (run.analysisErrorAt === null))
    || (run.analysisErrorAt !== null && !isTimestamp(run.analysisErrorAt))) failUnexpected();
  const isTerminal = new Set(["completed", "partial", "failed"]).has(run.status);
  if (isTerminal !== (run.finishedAt !== null)) failUnexpected();
  if (run.status === "analyzing" && (run.analysisComplete || run.proposals.length)) failUnexpected();
  if (new Set(["ready", "running"]).has(run.status) && !run.analysisComplete) failUnexpected();
  if (new Set(["completed", "partial"]).has(run.status)
    && (!run.analysisComplete || run.analysisProcessed !== run.analysisTotal)) failUnexpected();
  if (run.status === "failed" && !run.analysisComplete
    && (typeof run.analysisErrorCode !== "string" || !run.analysisErrorCode || !run.analysisErrorAt)) failUnexpected();
  const expectedUnit = run.definitions[0].ruleKey === BANK_RESERVATION_RULE_KEY ? "bank_anchors"
    : GROUPED_RULES.has(run.definitions[0].ruleKey) ? "calendar_months" : "records";
  if (run.analysisUnit !== expectedUnit) failUnexpected();
  for (const proposal of run.proposals) requireProposal(proposal, run);
  return run;
}

function statusError(message, statusCode) {
  const error = new Error(message);
  error.statusCode = statusCode;
  return error;
}

function safePublicError(error) {
  const mapped = mapRpcError(error);
  if (!error?.supabasePayload && mapped.statusCode < 500) return mapped;
  const safe = new Error(mapped.statusCode === 409
    ? "The reconciliation automation state changed. Refresh and try again."
    : mapped.statusCode >= 500
      ? "Unexpected server error."
      : "Reconciliation automation request could not be completed.");
  safe.statusCode = mapped.statusCode;
  return safe;
}

function actorFor(auth) {
  return cleanText(auth.user?.email) || cleanText(auth.user?.id);
}

async function requireManagedFeature(req, area) {
  const auth = await requireFeature(req, area, "financial-reconciliation");
  if (!cleanText(auth.access?.profile?.id)) {
    throw statusError("You do not have permission for this feature.", 403);
  }
  return auth;
}

function normalizeRunId(value) {
  if (typeof value !== "string" || !UUID_PATTERN.test(value)) {
    throw inputError("Run ID must be a valid UUID.");
  }
  return value.toLowerCase();
}

function requireClientRequestId(input) {
  if (!input.clientRequestId) throw inputError("Client request ID is required.");
  return input;
}

async function createAnalysis(input, actor, mode) {
  const result = await restQuery("rpc/create_financial_reconciliation_automatic_analysis", {
    method: "POST",
    body: {
      p_rule_keys: input.ruleKeys,
      p_mode: mode,
      p_actor: actor,
      p_client_request_id: input.clientRequestId,
    },
  });
  return requireManualRun(result, {
    actor,
    ruleKey: input.ruleKeys[0],
    clientRequestId: input.clientRequestId,
  });
}

async function analyzeRule(req, body) {
  const auth = await requireManagedFeature(req, "app");
  const input = requireClientRequestId(normalizeAnalyzePayload(body));
  if (input.action !== "analyze_rule" || input.ruleKeys.length !== 1) {
    throw inputError("Analyze rule requires exactly one manually enabled rule.");
  }
  const [selectedRuleKey] = input.ruleKeys;
  return createAnalysis({
    action: "analyze_rule",
    ruleKeys: [selectedRuleKey],
    clientRequestId: input.clientRequestId,
  }, actorFor(auth), "manual_rule");
}

async function continueAnalysis(req, body) {
  const auth = await requireManagedFeature(req, "app");
  const input = normalizeContinueAnalysisPayload(body);
  const result = await restQuery("rpc/continue_financial_reconciliation_automatic_analysis", {
    method: "POST",
    body: { p_run_id: input.runId, p_actor: actorFor(auth) },
  });
  return requireManualRun(result, { runId: input.runId, actor: actorFor(auth) });
}

async function executeSelected(req, body) {
  const auth = await requireManagedFeature(req, "app");
  const normalizedInput = normalizeExecutePayload(body);
  const input = {
    ...normalizedInput,
    runId: normalizedInput.runId.toLowerCase(),
    proposalIds: normalizedInput.proposalIds.map((proposalId) => proposalId.toLowerCase()),
  };
  if (new Set(input.proposalIds).size !== input.proposalIds.length) {
    throw inputError("Proposal IDs must contain up to 100 unique proposal IDs.");
  }
  const actor = actorFor(auth);
  const outcomes = [];
  const directOutcomes = new Map();
  const run = requireManualRun(await restQuery(
    "rpc/get_financial_reconciliation_automatic_run",
    { method: "POST", body: { p_run_id: input.runId } },
  ), { runId: input.runId, actor });
  const proposals = new Map((Array.isArray(run.proposals) ? run.proposals : [])
    .map((proposal) => [cleanText(proposal?.id).toLowerCase(), proposal]));
  if (input.proposalIds.some((proposalId) => !proposals.has(proposalId))) {
    throw inputError("Selected proposals must belong to the requested manual run.");
  }
  if (run.finishedAt) {
    return {
      run,
      outcomes: input.proposalIds.map((proposalId) => {
        const proposal = proposals.get(proposalId);
        const outcome = {
          proposalId,
          runId: input.runId,
          status: proposal.status,
        };
        if (proposal.reason) outcome.reason = proposal.reason;
        if (proposal.reconciliationId) outcome.reconciliationId = proposal.reconciliationId;
        return outcome;
      }),
    };
  }
  if (!input.proposalIds.length) {
    if (!run.analysisCompletedAt || cleanText(run.status).toLowerCase() === "analyzing") {
      throw inputError("Automatic analysis must be ready before it can be finished.");
    }
    const finalizedRun = await restQuery("rpc/finish_financial_reconciliation_automatic_run", {
      method: "POST",
      body: { p_run_id: input.runId },
    });
    return { run: requireManualRun(finalizedRun, { runId: input.runId, actor }), outcomes };
  }

  for (const proposalId of input.proposalIds) {
    let result;
    try {
      result = await restQuery("rpc/execute_financial_reconciliation_automatic_proposal", {
        method: "POST",
        body: { p_proposal_id: proposalId, p_actor: actor },
      });
    } catch {
      outcomes.push({ proposalId, status: "failed", reason: "execution_failed" });
      continue;
    }
    rejectUnsafeOwnData(result);
    const outcome = toAutomationPublicResult(result);
    if (!isPlainRecord(outcome)
      || outcome.proposalId !== proposalId
      || (outcome.runId !== undefined && outcome.runId !== input.runId)
      || !new Set(["completed", "failed", "stale", "deselected"]).has(outcome.status)
      || (outcome.reason !== undefined && typeof outcome.reason !== "string")
      || (outcome.reconciliationId !== undefined
        && (!UUID_PATTERN.test(outcome.reconciliationId)))
      || (outcome.status === "completed" && !UUID_PATTERN.test(outcome.reconciliationId))
      || Object.keys(outcome).some((key) => !new Set([
        "proposalId", "runId", "status", "reason", "reconciliationId",
      ]).has(key))) failUnexpected("proposal execution");
    directOutcomes.set(proposalId, outcome);
    outcomes.push(outcome);
  }

  const refreshedRun = requireManualRun(await restQuery(
    "rpc/get_financial_reconciliation_automatic_run",
    { method: "POST", body: { p_run_id: input.runId } },
  ), { runId: input.runId, actor });
  const refreshedProposals = new Map((Array.isArray(refreshedRun?.proposals) ? refreshedRun.proposals : [])
    .map((proposal) => [cleanText(proposal?.id).toLowerCase(), proposal]));
  for (const [proposalId, outcome] of directOutcomes) {
    const persisted = refreshedProposals.get(proposalId);
    if (!persisted || persisted.status !== outcome.status
      || (outcome.reason !== undefined && persisted.reason !== outcome.reason)
      || (outcome.reconciliationId !== undefined
        && persisted.reconciliationId !== outcome.reconciliationId)) failUnexpected("proposal execution");
  }
  const hasUnresolvedSelection = input.proposalIds.some((proposalId) => {
    const status = cleanText(refreshedProposals.get(proposalId)?.status).toLowerCase();
    return !status || status === "proposed" || status === "executing";
  });
  if (hasUnresolvedSelection) return { run: refreshedRun, outcomes };

  const finalizedRun = await restQuery("rpc/finish_financial_reconciliation_automatic_run", {
    method: "POST",
    body: { p_run_id: input.runId },
  });
  return { run: requireManualRun(finalizedRun, { runId: input.runId, actor }), outcomes };
}

module.exports = async function handler(req, res) {
  try {
    if (req.method === "GET") {
      const auth = await requireManagedFeature(req, "app");
      const view = cleanText(req.query?.view);
      if (view) {
        if (!new Set(["rules", "active_run"]).has(view) || cleanText(req.query?.run_id)) {
          throw inputError("Automation view is invalid.");
        }
        const resource = view === "rules"
          ? "rpc/get_financial_reconciliation_automatic_manual_rules"
          : "rpc/get_financial_reconciliation_automatic_active_run";
        const body = view === "rules" ? {} : { p_actor: actorFor(auth) };
        const result = await restQuery(resource, {
          method: "POST",
          body,
        });
        const publicResult = view === "rules"
          ? requireManualRuleCatalog(result)
          : requireManualRun(result, { actor: actorFor(auth) }, true);
        return res.status(200).json(publicResult);
      }
      const runId = normalizeRunId(req.query?.run_id);
      const run = await restQuery("rpc/get_financial_reconciliation_automatic_run", {
        method: "POST",
        body: { p_run_id: runId },
      });
      return res.status(200).json(requireManualRun(run, {
        runId,
        actor: actorFor(auth),
      }));
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const action = normalizeAutomationAction(body.action);
      const result = action === "analyze_rule"
        ? await analyzeRule(req, body)
        : action === "continue_analysis"
          ? await continueAnalysis(req, body)
          : await executeSelected(req, body);
      return res.status(200).json(result);
    }

    await requireManagedFeature(req, "app");
    res.setHeader("Allow", "GET, POST");
    return res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    return sendError(res, safePublicError(error));
  }
};
