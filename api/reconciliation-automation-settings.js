const { cleanText, parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const { mapRpcError } = require("./_reconciliation");
const {
  AUTOMATIC_RULE_DISPLAY_NAMES,
  AUTOMATIC_TIME_ZONE,
  AUTOMATIC_RULE_VERSIONS,
  ADYEN_MONTHLY_RULE_KEY,
  BANK_RESERVATION_RULE_KEY,
  BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY,
  BANK_STATEMENT_RULE_KEY,
  CREDIT_CARD_AMOUNT_ONLY_RULE_KEY,
  CREDIT_CARD_RULE_KEY,
  MONTHLY_INCOME_RULE_KEY,
  normalizeAutomationSettingsPayload,
  toAutomationPublicResult,
  toAutomationSettingsRpcPayload,
} = require("./_reconciliation-automation");

const SETTINGS_FIELDS = new Set(["schedule", "rules", "lastScheduledBatch"]);
const SCHEDULE_FIELDS = new Set(["enabled", "timeOfDay", "timeZone", "updatedBy", "updatedAt"]);
const RULE_FIELDS = new Set([
  "ruleKey", "ruleVersion", "displayName", "baseSourceType", "destinationSourceTypes",
  "logicDescription", "definition", "enabled", "allowManualExecution", "includeInScheduledBatch",
  "differenceAllowed", "maxDifferenceDays", "priority", "operator", "updatedBy", "updatedAt",
]);
const LAST_BATCH_FIELDS = new Set([
  "id", "scheduledSlot", "status", "counts", "ruleCount", "childCount", "startedAt", "finishedAt", "updatedAt",
]);
const BATCH_STATUSES = new Set(["pending", "running", "completed", "partial", "failed"]);
const BATCH_COUNT_FIELDS = new Set([
  "ruleCount", "childCount", "completedChildren", "partialChildren", "failedChildren", "unfinishedChildren",
  "bases", "proposed", "ambiguous", "skipped", "deselected", "executing", "completed", "stale", "failed",
]);
const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
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
    logicDescription: "Every eligible unlocked CGD Bank Statement and FDM Adyen record whose category is not TransferOutToAccount in the same closed calendar month forms one proposal; both sides are required and the signed difference must be within the configured allowance.",
    definition: {
      strategy: "closed_calendar_month",
      bankDescriptionContains: "Adyen",
      fdmAccount: "Adyen",
      fdmExcludedCategory: "TransferOutToAccount",
      requiresBothSides: true,
      monthMarkerDays: 31,
    },
  },
});

function isPlainRecord(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function failUnexpected() {
  throw new Error("Unexpected reconciliation settings response.");
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
    if (/private|diagnostic|internal|stack|error_detail|error_summary/i.test(key)) failUnexpected();
    rejectUnsafeOwnData(nested, seen);
  }
}

function requireExactFields(value, fields) {
  if (!isPlainRecord(value) || Object.keys(value).length !== fields.size
    || Object.keys(value).some((key) => !fields.has(key))) failUnexpected();
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

function requireSettingsResult(value) {
  rejectUnsafeOwnData(value);
  const result = toAutomationPublicResult(value);
  requireExactFields(result, SETTINGS_FIELDS);
  requireExactFields(result.schedule, SCHEDULE_FIELDS);
  if (typeof result.schedule.enabled !== "boolean"
    || !/^(?:[01]\d|2[0-3]):[0-5]\d$/.test(result.schedule.timeOfDay)
    || result.schedule.timeZone !== AUTOMATIC_TIME_ZONE
    || (result.schedule.updatedBy !== null && typeof result.schedule.updatedBy !== "string")
    || !isTimestamp(result.schedule.updatedAt)
    || !Array.isArray(result.rules) || result.rules.length !== 7) failUnexpected();
  const ruleKeys = new Set();
  const priorities = new Set();
  for (const rule of result.rules) {
    const ruleKey = rule?.ruleKey;
    const contract = MANAGED_RULE_CONTRACTS[ruleKey];
    requireExactFields(rule, RULE_FIELDS);
    if (!isPlainRecord(rule)
      || !Object.hasOwn(AUTOMATIC_RULE_VERSIONS, ruleKey)
      || rule.ruleVersion !== AUTOMATIC_RULE_VERSIONS[ruleKey]
      || rule.displayName !== AUTOMATIC_RULE_DISPLAY_NAMES[ruleKey]
      || !contract || rule.baseSourceType !== contract.baseSourceType
      || !Array.isArray(rule.destinationSourceTypes) || rule.destinationSourceTypes.length !== 1
      || rule.destinationSourceTypes[0] !== contract.destinationSourceType
      || rule.logicDescription !== contract.logicDescription
      || canonicalJson(rule.definition) !== canonicalJson(contract.definition)
      || typeof rule.enabled !== "boolean" || typeof rule.allowManualExecution !== "boolean"
      || typeof rule.includeInScheduledBatch !== "boolean"
      || (rule.updatedBy !== null && typeof rule.updatedBy !== "string") || !isTimestamp(rule.updatedAt)
      || !Number.isSafeInteger(rule.priority)
      || rule.priority < 1
      || ruleKeys.has(ruleKey)
      || priorities.has(rule.priority)) {
      failUnexpected();
    }
    const expectedOperator = ruleKey === MONTHLY_INCOME_RULE_KEY || ruleKey === ADYEN_MONTHLY_RULE_KEY
      ? "-" : "+";
    if (rule.operator !== expectedOperator
      || (ruleKey === BANK_RESERVATION_RULE_KEY
        && (Number(rule.differenceAllowed) !== 0 || rule.maxDifferenceDays < 0 || rule.maxDifferenceDays > 90))
      || (ruleKey === ADYEN_MONTHLY_RULE_KEY
        && (Number(rule.differenceAllowed) < 0 || rule.maxDifferenceDays !== 31))) failUnexpected();
    ruleKeys.add(ruleKey);
    priorities.add(rule.priority);
  }
  if (ruleKeys.size !== Object.keys(AUTOMATIC_RULE_VERSIONS).length) failUnexpected();
  try {
    normalizeAutomationSettingsPayload({
      schedule: {
        enabled: result.schedule.enabled,
        timeOfDay: result.schedule.timeOfDay,
        timeZone: result.schedule.timeZone,
      },
      rules: result.rules.map((rule) => ({
        ruleKey: rule.ruleKey,
        ruleVersion: rule.ruleVersion,
        enabled: rule.enabled,
        allowManualExecution: rule.allowManualExecution,
        includeInScheduledBatch: rule.includeInScheduledBatch,
        differenceAllowed: rule.differenceAllowed,
        maxDifferenceDays: rule.maxDifferenceDays,
        priority: rule.priority,
      })),
    });
  } catch {
    failUnexpected();
  }
  if (result.lastScheduledBatch !== null) {
    requireExactFields(result.lastScheduledBatch, LAST_BATCH_FIELDS);
    const batch = result.lastScheduledBatch;
    if (!UUID_PATTERN.test(batch.id)
      || typeof batch.scheduledSlot !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(batch.scheduledSlot)
      || !BATCH_STATUSES.has(batch.status) || !isPlainRecord(batch.counts)
      || !Number.isSafeInteger(batch.ruleCount) || batch.ruleCount < 0
      || !Number.isSafeInteger(batch.childCount) || batch.childCount < 0
      || !isTimestamp(batch.startedAt) || (batch.finishedAt !== null && !isTimestamp(batch.finishedAt))
      || !isTimestamp(batch.updatedAt)) failUnexpected();
    const countEntries = Object.entries(batch.counts);
    if (countEntries.some(([key, count]) => !BATCH_COUNT_FIELDS.has(key)
      || !Number.isSafeInteger(count) || count < 0)
      || (Object.hasOwn(batch.counts, "ruleCount") && batch.counts.ruleCount !== batch.ruleCount)
      || (Object.hasOwn(batch.counts, "childCount") && batch.counts.childCount !== batch.childCount)
      || (new Set(["completed", "partial", "failed"]).has(batch.status) !== (batch.finishedAt !== null))) {
      failUnexpected();
    }
  }
  return result;
}

function actorFor(auth) {
  return cleanText(auth.user?.email) || cleanText(auth.user?.id);
}

function permissionError() {
  const error = new Error("You do not have permission for this feature.");
  error.statusCode = 403;
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

async function requireManagedSettingsFeature(req) {
  const auth = await requireFeature(req, "settings", "financial-reconciliation");
  if (!cleanText(auth.access?.profile?.id)) throw permissionError();
  return auth;
}

module.exports = async function handler(req, res) {
  try {
    const auth = await requireManagedSettingsFeature(req);

    if (req.method === "GET") {
      const result = await restQuery("rpc/get_financial_reconciliation_automation_settings", {
        method: "POST",
        body: {},
      });
      return res.status(200).json(requireSettingsResult(result));
    }

    if (req.method === "PUT") {
      const settings = normalizeAutomationSettingsPayload(await parseBody(req));
      const result = await restQuery("rpc/replace_financial_reconciliation_automation_settings", {
        method: "POST",
        body: toAutomationSettingsRpcPayload(settings, actorFor(auth)),
      });
      return res.status(200).json(requireSettingsResult(result));
    }

    res.setHeader("Allow", "GET, PUT");
    return res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    return sendError(res, safePublicError(error));
  }
};
