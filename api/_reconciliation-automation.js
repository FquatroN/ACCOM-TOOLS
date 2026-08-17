const { SOURCE_TYPES, normalizeSourceType } = require("./_reconciliation");

const BANK_STATEMENT_RULE_KEY = "financial_documents_cgd_bank_statement";
const BANK_STATEMENT_RULE_VERSION = 2;
const CREDIT_CARD_RULE_KEY = "financial_documents_cgd_credit_card";
const CREDIT_CARD_RULE_VERSION = 1;
const BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY = "financial_documents_cgd_bank_statement_amount_only";
const BANK_STATEMENT_AMOUNT_ONLY_RULE_VERSION = 1;
const CREDIT_CARD_AMOUNT_ONLY_RULE_KEY = "financial_documents_cgd_credit_card_amount_only";
const CREDIT_CARD_AMOUNT_ONLY_RULE_VERSION = 1;
const AUTOMATIC_RULE_VERSIONS = Object.freeze({
  [BANK_STATEMENT_RULE_KEY]: BANK_STATEMENT_RULE_VERSION,
  [CREDIT_CARD_RULE_KEY]: CREDIT_CARD_RULE_VERSION,
  [BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY]: BANK_STATEMENT_AMOUNT_ONLY_RULE_VERSION,
  [CREDIT_CARD_AMOUNT_ONLY_RULE_KEY]: CREDIT_CARD_AMOUNT_ONLY_RULE_VERSION,
});
const AMOUNT_ONLY_RULE_KEYS = Object.freeze([
  BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY,
  CREDIT_CARD_AMOUNT_ONLY_RULE_KEY,
]);
const AMOUNT_ONLY_RULE_KEY_SET = new Set(AMOUNT_ONLY_RULE_KEYS);
const AUTOMATIC_RULE_KEY = BANK_STATEMENT_RULE_KEY;
const AUTOMATIC_RULE_VERSION = BANK_STATEMENT_RULE_VERSION;
const AUTOMATIC_TIME_ZONE = "Europe/Lisbon";
const AUTOMATION_ACTIONS = new Set(["analyze_rule", "continue_analysis", "execute_selected"]);
const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const EDITABLE_SCHEDULE_FIELDS = new Set(["enabled", "timeOfDay", "timeZone"]);
const EDITABLE_RULE_FIELDS = new Set([
  "ruleKey",
  "ruleVersion",
  "enabled",
  "allowManualExecution",
  "includeInScheduledBatch",
  "differenceAllowed",
  "maxDifferenceDays",
  "priority",
]);
const DEFINITION_FIELDS = new Set([
  "baseSourceType",
  "destinationSourceTypes",
  "logic",
  "definition",
  "thresholds",
]);
const PUBLIC_KEY_MAP = Object.freeze({
  rule_key: "ruleKey",
  rule_version: "ruleVersion",
  display_name: "displayName",
  base_source_type: "baseSourceType",
  destination_source_types: "destinationSourceTypes",
  logic_description: "logicDescription",
  allow_manual_execution: "allowManualExecution",
  include_in_scheduled_batch: "includeInScheduledBatch",
  difference_allowed: "differenceAllowed",
  max_difference_days: "maxDifferenceDays",
  time_of_day: "timeOfDay",
  time_zone: "timeZone",
  last_scheduled_run: "lastScheduledRun",
  batch_id: "batchId",
  batch_rule_key: "batchRuleKey",
  batch_rule_position: "batchRulePosition",
  batch_rule_count: "batchRuleCount",
  last_scheduled_batch: "lastScheduledBatch",
  client_request_id: "clientRequestId",
  scheduled_slot: "scheduledSlot",
  definition_config_snapshot: "definitionConfigSnapshot",
  analysis_completed_at: "analysisCompletedAt",
  analysis_cursor_date: "analysisCursorDate",
  analysis_cursor_id: "analysisCursorId",
  analysis_processed: "analysisProcessed",
  analysis_total: "analysisTotal",
  analysis_error_code: "analysisErrorCode",
  analysis_error_at: "analysisErrorAt",
  started_at: "startedAt",
  finished_at: "finishedAt",
  run_id: "runId",
  base_source_id: "baseSourceId",
  base_source_date: "baseSourceDate",
  candidate_groups: "candidateGroups",
  calculated_difference: "calculatedDifference",
  allowed_difference: "allowedDifference",
  reconciliation_id: "reconciliationId",
  automatic_trigger: "automaticTrigger",
  automatic_rule_key: "automaticRuleKey",
  automatic_rule_version: "automaticRuleVersion",
  automatic_run_id: "automaticRunId",
  automatic_proposal_id: "automaticProposalId",
  source_type: "sourceType",
  source_id: "sourceId",
  source_date: "sourceDate",
  amount_snapshot: "amountSnapshot",
  created_at: "createdAt",
  updated_at: "updatedAt",
});
const PRIVATE_PUBLIC_RESULT_KEYS = new Set(["error_detail", "internal_error", "error_summary", "diagnostic", "stack"]);

function inputError(message) {
  const error = new Error(message);
  error.statusCode = 400;
  return error;
}

function isPlainObject(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function requirePlainObject(value, label) {
  if (!isPlainObject(value)) throw inputError(`${label} must be an object.`);
  return value;
}

function requireOnlyKeys(value, allowedKeys, label) {
  for (const key of Object.keys(value)) {
    if (!allowedKeys.has(key)) throw inputError(`${label} contains an unsupported field.`);
  }
}

function requireBoolean(value, label) {
  if (typeof value !== "boolean") throw inputError(`${label} must be true or false.`);
  return value;
}

function requireInteger(value, label, minimum, maximum = Number.MAX_SAFE_INTEGER) {
  if (typeof value !== "number" || !Number.isInteger(value) || value < minimum || value > maximum) {
    throw inputError(`${label} must be between ${minimum} and ${maximum}.`);
  }
  return value;
}

function moneyToCents(value, label) {
  const text = String(value ?? "").trim();
  if (!/^\d+(?:\.\d{1,2})?$/.test(text)) throw inputError(`${label} must be a non-negative amount with at most two decimals.`);
  const [whole, fraction = ""] = text.split(".");
  const cents = (Number(whole) * 100) + Number(fraction.padEnd(2, "0"));
  if (!Number.isSafeInteger(cents)) throw inputError(`${label} is too large.`);
  return cents;
}

function centsToMoney(cents) {
  if (!Number.isSafeInteger(cents) || cents < 0) throw inputError("Difference allowed cents is invalid.");
  return `${Math.floor(cents / 100)}.${String(cents % 100).padStart(2, "0")}`;
}

function normalizeTimeOfDay(value) {
  if (typeof value !== "string" || !/^(?:[01]\d|2[0-3]):[0-5]\d$/.test(value)) {
    throw inputError("Time of day must be a valid HH:MM value.");
  }
  return value;
}

function normalizeRuleKey(value) {
  if (!Object.hasOwn(AUTOMATIC_RULE_VERSIONS, value)) throw inputError("Rule key is invalid.");
  return value;
}

function normalizeRuleVersion(value, ruleKey) {
  const normalizedRuleKey = normalizeRuleKey(ruleKey);
  if (value !== AUTOMATIC_RULE_VERSIONS[normalizedRuleKey]) throw inputError("Rule version is invalid.");
  return value;
}

function isAmountOnlyRuleKey(ruleKey) {
  return AMOUNT_ONLY_RULE_KEY_SET.has(ruleKey);
}

function normalizeSchedule(value) {
  const schedule = requirePlainObject(value, "Schedule");
  requireOnlyKeys(schedule, EDITABLE_SCHEDULE_FIELDS, "Schedule");
  if (schedule.timeZone !== AUTOMATIC_TIME_ZONE) throw inputError("Time zone must be Europe/Lisbon.");
  return {
    enabled: requireBoolean(schedule.enabled, "Schedule enabled"),
    timeOfDay: normalizeTimeOfDay(schedule.timeOfDay),
    timeZone: AUTOMATIC_TIME_ZONE,
  };
}

function normalizeManagedRule(value) {
  const rule = requirePlainObject(value, "Automation rule");
  for (const key of Object.keys(rule)) {
    if (DEFINITION_FIELDS.has(key) || !EDITABLE_RULE_FIELDS.has(key)) {
      throw inputError("Automation rules accept only editable managed-rule fields.");
    }
  }
  const ruleKey = normalizeRuleKey(rule.ruleKey);
  const differenceAllowedCents = moneyToCents(rule.differenceAllowed, "Difference allowed");
  if (isAmountOnlyRuleKey(ruleKey) && differenceAllowedCents !== 0) {
    throw inputError("Amount-only rules require a zero difference allowed.");
  }
  return {
    ruleKey,
    ruleVersion: normalizeRuleVersion(rule.ruleVersion, ruleKey),
    enabled: requireBoolean(rule.enabled, "Rule enabled"),
    allowManualExecution: requireBoolean(rule.allowManualExecution, "Allow manual execution"),
    includeInScheduledBatch: requireBoolean(rule.includeInScheduledBatch, "Include in scheduled batch"),
    differenceAllowedCents,
    maxDifferenceDays: requireInteger(rule.maxDifferenceDays, "Max difference days", 0, 90),
    priority: requireInteger(rule.priority, "Priority", 1),
  };
}

function normalizeAutomationSettingsPayload(value) {
  const input = requirePlainObject(value, "Automation settings");
  requireOnlyKeys(input, new Set(["schedule", "rules"]), "Automation settings");
  if (!Array.isArray(input.rules)) throw inputError("Rules must be an array.");
  const priorities = new Set();
  const ruleKeys = new Set();
  const rules = input.rules.map((rule) => {
    const normalized = normalizeManagedRule(rule);
    if (priorities.has(normalized.priority)) throw inputError("Duplicate rule priority.");
    if (ruleKeys.has(normalized.ruleKey)) throw inputError("Duplicate rule key.");
    priorities.add(normalized.priority);
    ruleKeys.add(normalized.ruleKey);
    return normalized;
  });
  return { schedule: normalizeSchedule(input.schedule), rules };
}

function normalizeAutomationAction(value) {
  if (typeof value !== "string" || !AUTOMATION_ACTIONS.has(value)) {
    throw inputError("Automation action is invalid.");
  }
  return value;
}

function normalizeUuid(value, label) {
  if (typeof value !== "string" || !UUID_PATTERN.test(value)) throw inputError(`${label} must be a valid UUID.`);
  return value;
}

function normalizeRuleKeys(value) {
  if (!Array.isArray(value)) throw inputError("Rule keys must be an array.");
  const keys = [];
  const seen = new Set();
  for (const ruleKey of value) {
    normalizeRuleKey(ruleKey);
    if (!seen.has(ruleKey)) {
      seen.add(ruleKey);
      keys.push(ruleKey);
    }
  }
  if (keys.length === 0) throw inputError("Rule keys must contain at least one managed rule key.");
  return keys;
}

function normalizeAnalyzePayload(value) {
  const input = requirePlainObject(value, "Analyze payload");
  requireOnlyKeys(input, new Set(["action", "ruleKeys", "clientRequestId"]), "Analyze payload");
  const action = normalizeAutomationAction(input.action);
  if (action !== "analyze_rule") throw inputError("Analysis action is invalid.");
  const clientRequestId = input.clientRequestId === undefined
    ? ""
    : normalizeUuid(input.clientRequestId, "Client request ID");
  const ruleKeys = normalizeRuleKeys(input.ruleKeys);
  if (ruleKeys.length !== 1) throw inputError("Analyze rule requires exactly one selected rule.");
  return { action, ruleKeys, clientRequestId };
}

function normalizeContinueAnalysisPayload(value) {
  const input = requirePlainObject(value, "Continue analysis payload");
  requireOnlyKeys(input, new Set(["action", "runId"]), "Continue analysis payload");
  if (normalizeAutomationAction(input.action) !== "continue_analysis") {
    throw inputError("Continue analysis action is invalid.");
  }
  return { action: "continue_analysis", runId: normalizeUuid(input.runId, "Run ID") };
}

function normalizeExecutePayload(value) {
  const input = requirePlainObject(value, "Execute payload");
  requireOnlyKeys(input, new Set(["action", "runId", "proposalIds"]), "Execute payload");
  const action = normalizeAutomationAction(input.action);
  if (action !== "execute_selected") throw inputError("Execution action is invalid.");
  if (!Array.isArray(input.proposalIds)) throw inputError("Proposal IDs must contain between 1 and 100 unique proposal IDs.");
  const proposalIds = [];
  const seen = new Set();
  for (const proposalId of input.proposalIds) {
    normalizeUuid(proposalId, "Proposal ID");
    if (seen.has(proposalId)) throw inputError("Proposal IDs must contain between 1 and 100 unique proposal IDs.");
    seen.add(proposalId);
    proposalIds.push(proposalId);
  }
  if (proposalIds.length < 1 || proposalIds.length > 100) {
    throw inputError("Proposal IDs must contain between 1 and 100 unique proposal IDs.");
  }
  return { action, runId: normalizeUuid(input.runId, "Run ID"), proposalIds };
}

function normalizeRpcSettings(settings) {
  if (!isPlainObject(settings) || !Array.isArray(settings.rules) || !settings.rules.some((rule) => Object.hasOwn(rule || {}, "differenceAllowedCents"))) {
    return normalizeAutomationSettingsPayload(settings);
  }
  const schedule = normalizeSchedule(settings.schedule);
  const priorities = new Set();
  const ruleKeys = new Set();
  const rules = settings.rules.map((rule) => {
    const input = requirePlainObject(rule, "Automation rule");
    requireOnlyKeys(input, new Set([
      "ruleKey",
      "ruleVersion",
      "enabled",
      "allowManualExecution",
      "includeInScheduledBatch",
      "differenceAllowedCents",
      "maxDifferenceDays",
      "priority",
    ]), "Automation rule");
    const ruleKey = normalizeRuleKey(input.ruleKey);
    const differenceAllowedCents = requireInteger(input.differenceAllowedCents, "Difference allowed cents", 0);
    if (isAmountOnlyRuleKey(ruleKey) && differenceAllowedCents !== 0) {
      throw inputError("Amount-only rules require a zero difference allowed.");
    }
    const normalized = {
      ruleKey,
      ruleVersion: normalizeRuleVersion(input.ruleVersion, ruleKey),
      enabled: requireBoolean(input.enabled, "Rule enabled"),
      allowManualExecution: requireBoolean(input.allowManualExecution, "Allow manual execution"),
      includeInScheduledBatch: requireBoolean(input.includeInScheduledBatch, "Include in scheduled batch"),
      differenceAllowedCents,
      maxDifferenceDays: requireInteger(input.maxDifferenceDays, "Max difference days", 0, 90),
      priority: requireInteger(input.priority, "Priority", 1),
    };
    if (priorities.has(normalized.priority)) throw inputError("Duplicate rule priority.");
    if (ruleKeys.has(normalized.ruleKey)) throw inputError("Duplicate rule key.");
    priorities.add(normalized.priority);
    ruleKeys.add(normalized.ruleKey);
    return normalized;
  });
  return { schedule, rules };
}

function toAutomationSettingsRpcPayload(settings, actor) {
  const normalized = normalizeRpcSettings(settings);
  return {
    p_schedule: {
      enabled: normalized.schedule.enabled,
      time_of_day: normalized.schedule.timeOfDay,
      time_zone: normalized.schedule.timeZone,
    },
    p_rules: normalized.rules.map((rule) => ({
      rule_key: rule.ruleKey,
      rule_version: rule.ruleVersion,
      enabled: rule.enabled,
      allow_manual_execution: rule.allowManualExecution,
      include_in_scheduled_batch: rule.includeInScheduledBatch,
      difference_allowed: centsToMoney(rule.differenceAllowedCents),
      max_difference_days: rule.maxDifferenceDays,
      priority: rule.priority,
    })),
    p_actor: actor,
  };
}

function headerValue(headers, name) {
  if (!headers || typeof headers !== "object") return undefined;
  const matchedKey = Object.keys(headers).find((key) => key.toLowerCase() === name);
  return matchedKey === undefined ? undefined : headers[matchedKey];
}

function isCronRequest(req, cronSecret) {
  const vercelCron = headerValue(req && req.headers, "x-vercel-cron");
  if (vercelCron === "1") return true;
  const authorization = headerValue(req && req.headers, "authorization");
  return typeof cronSecret === "string" && cronSecret !== "" && authorization === `Bearer ${cronSecret}`;
}

function toAutomationPublicResult(value) {
  if (Array.isArray(value)) return value.map(toAutomationPublicResult);
  if (!isPlainObject(value)) return value;
  const result = {};
  for (const [key, nestedValue] of Object.entries(value)) {
    if (PRIVATE_PUBLIC_RESULT_KEYS.has(key)) continue;
    result[PUBLIC_KEY_MAP[key] || key] = toAutomationPublicResult(nestedValue);
  }
  return result;
}

module.exports = {
  AUTOMATIC_RULE_KEY,
  AUTOMATIC_RULE_VERSION,
  AUTOMATIC_TIME_ZONE,
  AUTOMATIC_RULE_VERSIONS,
  AMOUNT_ONLY_RULE_KEYS,
  AUTOMATION_ACTIONS,
  BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY,
  BANK_STATEMENT_AMOUNT_ONLY_RULE_VERSION,
  BANK_STATEMENT_RULE_KEY,
  BANK_STATEMENT_RULE_VERSION,
  CREDIT_CARD_AMOUNT_ONLY_RULE_KEY,
  CREDIT_CARD_AMOUNT_ONLY_RULE_VERSION,
  CREDIT_CARD_RULE_KEY,
  CREDIT_CARD_RULE_VERSION,
  SOURCE_TYPES,
  isAmountOnlyRuleKey,
  isCronRequest,
  normalizeAnalyzePayload,
  normalizeAutomationAction,
  normalizeAutomationSettingsPayload,
  normalizeContinueAnalysisPayload,
  normalizeExecutePayload,
  normalizeRpcSettings,
  normalizeRuleKey,
  normalizeRuleVersion,
  normalizeSourceType,
  toAutomationPublicResult,
  toAutomationSettingsRpcPayload,
};
