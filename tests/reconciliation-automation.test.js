const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const {
  AUTOMATIC_RULE_KEY,
  AUTOMATIC_RULE_VERSION,
  AUTOMATIC_RULE_VERSIONS,
  AUTOMATIC_TIME_ZONE,
  AMOUNT_ONLY_RULE_KEYS,
  BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY,
  BANK_STATEMENT_AMOUNT_ONLY_RULE_VERSION,
  BANK_STATEMENT_RULE_KEY,
  BANK_STATEMENT_RULE_VERSION,
  CREDIT_CARD_AMOUNT_ONLY_RULE_KEY,
  CREDIT_CARD_AMOUNT_ONLY_RULE_VERSION,
  CREDIT_CARD_RULE_KEY,
  CREDIT_CARD_RULE_VERSION,
  isAmountOnlyRuleKey,
  isMonthlyAggregateRule,
  isCronRequest,
  MONTHLY_INCOME_RULE_KEY,
  normalizeAnalyzePayload,
  normalizeAutomationAction,
  normalizeAutomationSettingsPayload,
  normalizeContinueAnalysisPayload,
  normalizeExecutePayload,
  normalizeRpcSettings,
  normalizeRuleVersion,
  toAutomationPublicResult,
  toAutomationSettingsRpcPayload,
} = require("../api/_reconciliation-automation");
const { mapRpcError } = require("../api/_reconciliation");

const RUN_ID = "00000000-0000-0000-0000-000000000001";
const PROPOSAL_ID = "00000000-0000-0000-0000-000000000002";
const REQUEST_ID = "00000000-0000-0000-0000-000000000003";
const BATCH_ID = "00000000-0000-0000-0000-000000000006";
const SCHEMA_MIGRATION_PATH = path.join(
  __dirname,
  "..",
  "supabase-migrations",
  "2026-08-14-financial-reconciliation-automation-schema.sql",
);
const ANALYSIS_MIGRATION_PATH = path.join(
  __dirname,
  "..",
  "supabase-migrations",
  "2026-08-14-financial-reconciliation-automation-analysis.sql",
);
const ANALYSIS_PERFORMANCE_MIGRATION_PATH = path.join(
  __dirname,
  "..",
  "supabase-migrations",
  "2026-08-15-financial-reconciliation-automation-analysis-performance.sql",
);
const ANALYSIS_INDEX_LOOKUP_MIGRATION_PATH = path.join(
  __dirname,
  "..",
  "supabase-migrations",
  "2026-08-15-financial-reconciliation-automation-candidate-index-lookup.sql",
);
const AUTOMATION_90_DAY_MIGRATION_PATH = path.join(
  __dirname,
  "..",
  "supabase-migrations",
  "2026-08-16-financial-reconciliation-automation-90-day-performance.sql",
);
const CREDIT_CARD_MIGRATION_PATH = path.join(
  __dirname,
  "..",
  "supabase-migrations",
  "2026-08-16-financial-reconciliation-automation-credit-card-rule.sql",
);
const EXECUTION_MIGRATION_PATH = path.join(
  __dirname,
  "..",
  "supabase-migrations",
  "2026-08-14-financial-reconciliation-automation-execution.sql",
);
const SOURCE_RULE_MIGRATION_PATH = path.join(
  __dirname,
  "..",
  "supabase-migrations",
  "2026-08-11-financial-reconciliation-source-rules.sql",
);
const RPC_SMOKE_PATH = path.join(__dirname, "reconciliation-automation-rpc.smoke.sql");
const MANUAL_RPC_SMOKE_PATH = path.join(__dirname, "reconciliation-rpc.smoke.sql");
const SETTINGS_HANDLER_PATH = path.join(__dirname, "..", "api", "reconciliation-automation-settings.js");
const MANUAL_HANDLER_PATH = path.join(__dirname, "..", "api", "reconciliation-automation.js");
const MEMBERS_HANDLER_PATH = path.join(__dirname, "..", "api", "reconciliation-automation-members.js");
const CRON_HANDLER_PATH = path.join(__dirname, "..", "api", "reconciliation-automation-cron.js");
const VERCEL_CONFIG_PATH = path.join(__dirname, "..", "vercel.json");
const README_PATH = path.join(__dirname, "..", "README.md");
const POS_INCOME_MIGRATION_PATH = path.join(
  __dirname,
  "..",
  "supabase-migrations",
  "2026-08-22-financial-reconciliation-automation-pos-income.sql",
);
const SUPABASE_MODULE_PATH = require.resolve("../api/_supabase");
const PROPOSAL_ID_2 = "00000000-0000-0000-0000-000000000004";
const PROPOSAL_ID_3 = "00000000-0000-0000-0000-000000000005";
const CASE_UUID = "abcdefab-cdef-abcd-efab-cdefabcdefab";
const CRON_SECRET = "test-cron-secret";
const SCHEDULE_ACTOR = "system:reconciliation";

function responseRecorder() {
  return {
    statusCode: 0,
    body: null,
    headers: {},
    status(code) {
      this.statusCode = code;
      return this;
    },
    json(body) {
      this.body = body;
      return this;
    },
    setHeader(name, value) {
      this.headers[name] = value;
      return this;
    },
  };
}

async function withMockedHandler(handlerPath, supabase, run) {
  const previousHandler = require.cache[handlerPath];
  const previousSupabase = require.cache[SUPABASE_MODULE_PATH];
  delete require.cache[handlerPath];
  require.cache[SUPABASE_MODULE_PATH] = {
    id: SUPABASE_MODULE_PATH,
    filename: SUPABASE_MODULE_PATH,
    loaded: true,
    exports: supabase,
  };

  try {
    await run(require(handlerPath));
  } finally {
    delete require.cache[handlerPath];
    if (previousHandler) require.cache[handlerPath] = previousHandler;
    if (previousSupabase) require.cache[SUPABASE_MODULE_PATH] = previousSupabase;
    else delete require.cache[SUPABASE_MODULE_PATH];
  }
}

async function withCronEnvironment(nowIso, run) {
  const NativeDate = global.Date;
  const previousSecret = process.env.CRON_SECRET;
  class FixedDate extends NativeDate {
    constructor(...args) {
      super(...(args.length ? args : [nowIso]));
    }

    static now() {
      return new NativeDate(nowIso).getTime();
    }
  }

  global.Date = FixedDate;
  process.env.CRON_SECRET = CRON_SECRET;
  try {
    await run();
  } finally {
    global.Date = NativeDate;
    if (previousSecret === undefined) delete process.env.CRON_SECRET;
    else process.env.CRON_SECRET = previousSecret;
  }
}

function uuidFor(value) {
  return `00000000-0000-0000-0000-${String(value).padStart(12, "0")}`;
}

function compareCodeUnits(left, right) {
  return left < right ? -1 : left > right ? 1 : 0;
}

function scheduledRun(overrides = {}) {
  const run = {
    runId: RUN_ID,
    trigger: "scheduled",
    scope: "rule",
    status: "ready",
    actor: SCHEDULE_ACTOR,
    clientRequestId: null,
    scheduledSlot: "2026-08-15",
    batchId: BATCH_ID,
    batchRuleKey: AUTOMATIC_RULE_KEY,
    batchRulePosition: 1,
    batchRuleCount: 2,
    analysisCursorDate: null,
    analysisCursorId: null,
    analysisProcessed: 1,
    analysisTotal: 1,
    analysisErrorCode: null,
    analysisErrorAt: null,
    analysisComplete: true,
    analysisCompletedAt: "2026-08-15T02:00:01.000Z",
    counts: { bases: 1, proposed: 0, ambiguous: 0, skipped: 0 },
    startedAt: "2026-08-15T02:00:00.000Z",
    finishedAt: null,
    definitions: [{ ruleKey: AUTOMATIC_RULE_KEY, priority: 1 }],
    proposals: [],
    ...overrides,
  };
  run.proposals = run.proposals.map((proposal) => ({ runId: run.runId, ...proposal }));
  return run;
}

function scheduledClaim(run, overrides = {}) {
  return {
    claimed: true,
    resumed: true,
    batchId: run.batchId,
    batchRulePosition: run.batchRulePosition,
    batchRuleCount: run.batchRuleCount,
    run,
    ...overrides,
  };
}

function mockedSupabase(overrides = {}) {
  const { exposeOldestAnalysis = false, restQuery: restQueryOverride, ...otherOverrides } = overrides;
  const fallbackRestQuery = restQueryOverride || (async () => ({}));
  return {
    cleanText: (value) => String(value ?? "").trim(),
    parseBody: async (request) => request.body || {},
    requireFeature: async () => ({
      user: { email: "user@example.com", id: "user-1" },
      access: { profile: { id: "profile-1" } },
    }),
    restQuery: async (resource, options) => {
      if (resource === "rpc/continue_financial_reconciliation_automatic_oldest_analysis"
        && !exposeOldestAnalysis) {
        return { continued: false };
      }
      return fallbackRestQuery(resource, options);
    },
    sendError: (response, error) => response.status(error.statusCode || 500).json({ error: error.message }),
    ...otherOverrides,
  };
}

function managedSettings(overrides = {}) {
  return {
    schedule: { enabled: true, timeOfDay: "02:15", timeZone: AUTOMATIC_TIME_ZONE },
    rules: [{
      ruleKey: AUTOMATIC_RULE_KEY,
      ruleVersion: AUTOMATIC_RULE_VERSION,
      enabled: true,
      allowManualExecution: true,
      includeInScheduledBatch: false,
      differenceAllowed: "1.25",
      maxDifferenceDays: 7,
      priority: 1,
    }],
    ...overrides,
  };
}

const creditCardRule = {
  ruleKey: CREDIT_CARD_RULE_KEY,
  ruleVersion: CREDIT_CARD_RULE_VERSION,
  enabled: false,
  allowManualExecution: false,
  includeInScheduledBatch: false,
  differenceAllowed: "0.00",
  maxDifferenceDays: 10,
  priority: 2,
};

function fourRuleSettings({
  amountOnlyDifferenceAllowed = "0.00",
  bankStatementAmountOnlyDifferenceAllowed = amountOnlyDifferenceAllowed,
  creditCardAmountOnlyDifferenceAllowed = amountOnlyDifferenceAllowed,
  ...overrides
} = {}) {
  return managedSettings({
    rules: [
      managedSettings().rules[0],
      creditCardRule,
      {
        ruleKey: BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY,
        ruleVersion: BANK_STATEMENT_AMOUNT_ONLY_RULE_VERSION,
        enabled: false,
        allowManualExecution: false,
        includeInScheduledBatch: false,
        differenceAllowed: bankStatementAmountOnlyDifferenceAllowed,
        maxDifferenceDays: 1,
        priority: 3,
      },
      {
        ruleKey: CREDIT_CARD_AMOUNT_ONLY_RULE_KEY,
        ruleVersion: CREDIT_CARD_AMOUNT_ONLY_RULE_VERSION,
        enabled: false,
        allowManualExecution: false,
        includeInScheduledBatch: false,
        differenceAllowed: creditCardAmountOnlyDifferenceAllowed,
        maxDifferenceDays: 1,
        priority: 4,
      },
    ],
    ...overrides,
  });
}

const monthlyIncomeRule = {
  ruleKey: MONTHLY_INCOME_RULE_KEY,
  ruleVersion: 2,
  enabled: true,
  allowManualExecution: true,
  includeInScheduledBatch: true,
  differenceAllowed: "7500.00",
  maxDifferenceDays: 31,
  priority: 5,
};

function fiveRuleSettings(overrides = {}) {
  return fourRuleSettings({
    rules: [...fourRuleSettings().rules, monthlyIncomeRule],
    ...overrides,
  });
}

function amountOnlyScheduledDefinition(ruleKey, priority) {
  const bank = ruleKey === BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY;
  return {
    ruleKey,
    ruleVersion: 1,
    displayName: bank
      ? "Financial Documents to CGD Bank Account – AMOUNT ONLY"
      : "Financial Documents to CGD Credit Card – AMOUNT ONLY",
    priority,
    differenceAllowed: 0,
    maxDifferenceDays: 1,
    destinationSourceType: bank ? "import_cgd_extrato_ordem" : "import_cgd_cartao_credito",
    definition: {
      baseSourceType: "financial_documents",
      destinationSourceTypes: [bank ? "import_cgd_extrato_ordem" : "import_cgd_cartao_credito"],
      baseEligibility: {
        payment: {
          operator: "exact_text_equal",
          value: bank ? "Banco" : "Visa",
          caseSensitive: true,
          trim: false,
        },
      },
      matchingMode: "amount_only_one_to_one",
      fixedDifferenceAllowed: 0,
      maxDifferenceDays: { minimum: 0, maximum: 90, default: 1 },
      maxDestinationRecords: 1,
    },
    operator: "+",
  };
}

function monthlyIncomeScheduledDefinition(priority) {
  return {
    ruleKey: MONTHLY_INCOME_RULE_KEY,
    ruleVersion: 2,
    displayName: "Card Payments - POS - Income",
    priority,
    differenceAllowed: 7500,
    maxDifferenceDays: 31,
    destinationSourceType: "import_fdm_accounts",
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
    operator: "-",
  };
}

function fourRuleRpcSettings({
  amountOnlyDifferenceAllowedCents = 0,
  bankStatementAmountOnlyDifferenceAllowedCents = amountOnlyDifferenceAllowedCents,
  creditCardAmountOnlyDifferenceAllowedCents = amountOnlyDifferenceAllowedCents,
  ...overrides
} = {}) {
  return {
    schedule: { enabled: true, timeOfDay: "02:15", timeZone: AUTOMATIC_TIME_ZONE },
    rules: [
      {
        ruleKey: BANK_STATEMENT_RULE_KEY,
        ruleVersion: BANK_STATEMENT_RULE_VERSION,
        enabled: true,
        allowManualExecution: true,
        includeInScheduledBatch: false,
        differenceAllowedCents: 125,
        maxDifferenceDays: 7,
        priority: 1,
      },
      {
        ruleKey: CREDIT_CARD_RULE_KEY,
        ruleVersion: CREDIT_CARD_RULE_VERSION,
        enabled: false,
        allowManualExecution: false,
        includeInScheduledBatch: false,
        differenceAllowedCents: 0,
        maxDifferenceDays: 10,
        priority: 2,
      },
      {
        ruleKey: BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY,
        ruleVersion: BANK_STATEMENT_AMOUNT_ONLY_RULE_VERSION,
        enabled: false,
        allowManualExecution: false,
        includeInScheduledBatch: false,
        differenceAllowedCents: bankStatementAmountOnlyDifferenceAllowedCents,
        maxDifferenceDays: 1,
        priority: 3,
      },
      {
        ruleKey: CREDIT_CARD_AMOUNT_ONLY_RULE_KEY,
        ruleVersion: CREDIT_CARD_AMOUNT_ONLY_RULE_VERSION,
        enabled: false,
        allowManualExecution: false,
        includeInScheduledBatch: false,
        differenceAllowedCents: creditCardAmountOnlyDifferenceAllowedCents,
        maxDifferenceDays: 1,
        priority: 4,
      },
    ],
    ...overrides,
  };
}

function fiveRuleRpcSettings(overrides = {}) {
  return fourRuleRpcSettings({
    rules: [...fourRuleRpcSettings().rules, {
      ruleKey: MONTHLY_INCOME_RULE_KEY,
      ruleVersion: 2,
      enabled: true,
      allowManualExecution: true,
      includeInScheduledBatch: true,
      differenceAllowedCents: 750000,
      maxDifferenceDays: 31,
      priority: 5,
    }],
    ...overrides,
  });
}

function productionSettingsRules() {
  return [
    {
      ruleKey: BANK_STATEMENT_RULE_KEY,
      ruleVersion: BANK_STATEMENT_RULE_VERSION,
      displayName: "Financial Documents to CGD Bank Statement",
      baseSourceType: "financial_documents",
      destinationSourceTypes: ["import_cgd_extrato_ordem"],
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
      enabled: true,
      allowManualExecution: true,
      includeInScheduledBatch: true,
      differenceAllowed: 1.25,
      maxDifferenceDays: 7,
      priority: 1,
      updatedBy: "admin@example.com",
      updatedAt: "2026-08-17T09:00:00.000Z",
    },
    {
      ruleKey: CREDIT_CARD_RULE_KEY,
      ruleVersion: CREDIT_CARD_RULE_VERSION,
      displayName: "Financial Documents to CGD Credit Card",
      baseSourceType: "financial_documents",
      destinationSourceTypes: ["import_cgd_cartao_credito"],
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
      enabled: true,
      allowManualExecution: false,
      includeInScheduledBatch: true,
      differenceAllowed: 0,
      maxDifferenceDays: 10,
      priority: 2,
      updatedBy: "admin@example.com",
      updatedAt: "2026-08-17T09:00:01.000Z",
    },
    {
      ruleKey: BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY,
      ruleVersion: BANK_STATEMENT_AMOUNT_ONLY_RULE_VERSION,
      displayName: "Financial Documents to CGD Bank Account – AMOUNT ONLY",
      baseSourceType: "financial_documents",
      destinationSourceTypes: ["import_cgd_extrato_ordem"],
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
      enabled: false,
      allowManualExecution: false,
      includeInScheduledBatch: false,
      differenceAllowed: 0,
      maxDifferenceDays: 1,
      priority: 3,
      updatedBy: "",
      updatedAt: "2026-08-17T09:00:02.000Z",
    },
    {
      ruleKey: CREDIT_CARD_AMOUNT_ONLY_RULE_KEY,
      ruleVersion: CREDIT_CARD_AMOUNT_ONLY_RULE_VERSION,
      displayName: "Financial Documents to CGD Credit Card – AMOUNT ONLY",
      baseSourceType: "financial_documents",
      destinationSourceTypes: ["import_cgd_cartao_credito"],
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
      enabled: false,
      allowManualExecution: false,
      includeInScheduledBatch: false,
      differenceAllowed: 0,
      maxDifferenceDays: 1,
      priority: 4,
      updatedBy: "",
      updatedAt: "2026-08-17T09:00:03.000Z",
    },
  ];
}

function productionSettingsRpcResult(ruleCount) {
  return {
    schedule: {
      enabled: true,
      timeOfDay: "02:15",
      timeZone: AUTOMATIC_TIME_ZONE,
      updatedBy: "admin@example.com",
      updatedAt: "2026-08-17T09:00:04.000Z",
    },
    rules: productionSettingsRules().slice(0, ruleCount),
    last_scheduled_batch: {
      id: BATCH_ID,
      scheduledSlot: "2026-08-17",
      status: "partial",
      counts: {
        ruleCount,
        childCount: ruleCount,
        completedChildren: ruleCount - 1,
        partialChildren: 0,
        failedChildren: 1,
        unfinishedChildren: 0,
      },
      ruleCount,
      childCount: ruleCount,
      startedAt: "2026-08-17T02:15:00.000Z",
      finishedAt: "2026-08-17T02:19:00.000Z",
      updatedAt: "2026-08-17T02:19:00.000Z",
    },
  };
}

function expectedPublicSettings(ruleCount) {
  const rpcResult = productionSettingsRpcResult(ruleCount);
  return {
    schedule: rpcResult.schedule,
    rules: rpcResult.rules,
    lastScheduledBatch: rpcResult.last_scheduled_batch,
  };
}

test("managed settings accept the two explicit rule/version pairs", () => {
  const input = managedSettings({
    rules: [managedSettings().rules[0], creditCardRule],
  });
  assert.deepEqual(
    normalizeAutomationSettingsPayload(input).rules.map(({ ruleKey, ruleVersion }) => ({ ruleKey, ruleVersion })),
    [
      { ruleKey: BANK_STATEMENT_RULE_KEY, ruleVersion: 2 },
      { ruleKey: CREDIT_CARD_RULE_KEY, ruleVersion: 1 },
    ],
  );
  assert.throws(() => normalizeAutomationSettingsPayload({
    ...input,
    rules: [{ ...creditCardRule, ruleVersion: 2 }],
  }), /rule version/i);
});

test("managed automation accepts the four explicit key/version pairs", () => {
  const normalized = normalizeAutomationSettingsPayload(fourRuleSettings());
  assert.deepEqual(normalized.rules.map(({ ruleKey, ruleVersion }) => [ruleKey, ruleVersion]), [
    [BANK_STATEMENT_RULE_KEY, 2],
    [CREDIT_CARD_RULE_KEY, 1],
    [BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY, 1],
    [CREDIT_CARD_AMOUNT_ONLY_RULE_KEY, 1],
  ]);
  assert.equal(isAmountOnlyRuleKey(BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY), true);
  assert.equal(isAmountOnlyRuleKey(CREDIT_CARD_RULE_KEY), false);

  assert.throws(() => normalizeAutomationSettingsPayload(fourRuleSettings({
    rules: [{ ...fourRuleSettings().rules[2], ruleVersion: 2 }],
  })), /rule version/i);
  assert.throws(() => normalizeAutomationSettingsPayload(fourRuleSettings({
    rules: [{ ...fourRuleSettings().rules[2], ruleKey: `${BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY}_injected` }],
  })), /rule key/i);
  assert.throws(() => normalizeAutomationSettingsPayload(fourRuleSettings({
    rules: [
      ...fourRuleSettings().rules.slice(0, 3),
      { ...fourRuleSettings().rules[2], priority: 4 },
    ],
  })), /duplicate rule key/i);
});

test("managed automation exposes a five-rule response with a fixed monthly income window", () => {
  const settingsPayload = fiveRuleSettings();
  const normalized = toAutomationPublicResult({
    rules: toAutomationSettingsRpcPayload(settingsPayload, "user@example.com").p_rules,
  });

  assert.equal(AUTOMATIC_RULE_VERSIONS[MONTHLY_INCOME_RULE_KEY], 2);
  assert.equal(normalized.rules[4].differenceAllowed, "7500.00");
  assert.equal(normalized.rules[4].maxDifferenceDays, 31);
  assert.equal(isMonthlyAggregateRule(MONTHLY_INCOME_RULE_KEY), true);

  const tampered = structuredClone(settingsPayload);
  tampered.rules.find((rule) => rule.ruleKey === MONTHLY_INCOME_RULE_KEY).maxDifferenceDays = 30;
  assert.throws(
    () => normalizeAutomationSettingsPayload(tampered),
    /Maximum difference in days is invalid/,
  );

  const authoritative = fiveRuleRpcSettings();
  assert.equal(normalizeRpcSettings(authoritative).rules[4].maxDifferenceDays, 31);
  authoritative.rules[4].maxDifferenceDays = 30;
  assert.throws(() => normalizeRpcSettings(authoritative), /Maximum difference in days is invalid/);
});

test("amount-only tolerance is fixed at zero in both settings shapes", () => {
  assert.throws(() => normalizeAutomationSettingsPayload(fourRuleSettings({
    amountOnlyDifferenceAllowed: "0.01",
  })), /amount-only.*zero/i);
  assert.throws(() => normalizeRpcSettings(fourRuleRpcSettings({
    amountOnlyDifferenceAllowedCents: 1,
  })), /amount-only.*zero/i);
  assert.throws(() => normalizeAutomationSettingsPayload(fourRuleSettings({
    creditCardAmountOnlyDifferenceAllowed: "0.01",
  })), /amount-only.*zero/i);
  assert.throws(() => normalizeRpcSettings(fourRuleRpcSettings({
    creditCardAmountOnlyDifferenceAllowedCents: 1,
  })), /amount-only.*zero/i);
});

test("public amount-only keys cannot disable amount-only validation", () => {
  const restoreCreditCardKey = typeof AMOUNT_ONLY_RULE_KEYS.add === "function";
  try {
    const mutationResult = typeof AMOUNT_ONLY_RULE_KEYS.delete === "function"
      ? AMOUNT_ONLY_RULE_KEYS.delete(CREDIT_CARD_AMOUNT_ONLY_RULE_KEY)
      : Reflect.deleteProperty(AMOUNT_ONLY_RULE_KEYS, 1);
    assert.equal(mutationResult, false);

    const copiedKeys = [...AMOUNT_ONLY_RULE_KEYS];
    copiedKeys.length = 0;
    assert.equal(isAmountOnlyRuleKey(CREDIT_CARD_AMOUNT_ONLY_RULE_KEY), true);
    assert.throws(() => normalizeRpcSettings(fourRuleRpcSettings({
      creditCardAmountOnlyDifferenceAllowedCents: 1,
    })), /amount-only.*zero/i);
  } finally {
    if (restoreCreditCardKey) AMOUNT_ONLY_RULE_KEYS.add(CREDIT_CARD_AMOUNT_ONLY_RULE_KEY);
  }
});

test("rule-version validation rejects an unknown key even with an undefined version", () => {
  assert.throws(
    () => normalizeRuleVersion(undefined, "not-a-managed-rule"),
    /rule key/i,
  );
});

test("manual analysis accepts exactly one allowlisted rule and has no batch action", () => {
  assert.deepEqual(normalizeAnalyzePayload({
    action: "analyze_rule",
    ruleKeys: [CREDIT_CARD_RULE_KEY],
    clientRequestId: REQUEST_ID,
  }).ruleKeys, [CREDIT_CARD_RULE_KEY]);
  assert.throws(() => normalizeAnalyzePayload({
    action: "analyze_rule",
    ruleKeys: [BANK_STATEMENT_RULE_KEY, CREDIT_CARD_RULE_KEY],
    clientRequestId: REQUEST_ID,
  }), /exactly one/i);
  assert.throws(() => normalizeAutomationAction("analyze_batch"), /automation action/i);
});

test("two-key manual analysis rejects before any RPC", async () => {
  let rpcCalled = false;
  const response = responseRecorder();
  await withMockedHandler(MANUAL_HANDLER_PATH, mockedSupabase({
    restQuery: async () => {
      rpcCalled = true;
      return {};
    },
  }), async (handler) => {
    await handler({
      method: "POST",
      body: {
        action: "analyze_rule",
        ruleKeys: [BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY, CREDIT_CARD_AMOUNT_ONLY_RULE_KEY],
        clientRequestId: REQUEST_ID,
      },
    }, response);
  });
  assert.equal(response.statusCode, 400);
  assert.equal(rpcCalled, false);
});

test("batch lifecycle keys map without leaking diagnostic keys", () => {
  assert.deepEqual(toAutomationPublicResult({
    batch_id: RUN_ID,
    batch_rule_key: CREDIT_CARD_RULE_KEY,
    batch_rule_position: 2,
    batch_rule_count: 3,
    last_scheduled_batch: { error_detail: "hidden", status: "partial" },
  }), {
    batchId: RUN_ID,
    batchRuleKey: CREDIT_CARD_RULE_KEY,
    batchRulePosition: 2,
    batchRuleCount: 3,
    lastScheduledBatch: { status: "partial" },
  });
});

test("automation settings accept only editable managed-rule fields", () => {
  assert.deepEqual(normalizeAutomationSettingsPayload(managedSettings()), {
    schedule: { enabled: true, timeOfDay: "02:15", timeZone: "Europe/Lisbon" },
    rules: [{
      ruleKey: AUTOMATIC_RULE_KEY,
      ruleVersion: 2,
      enabled: true,
      allowManualExecution: true,
      includeInScheduledBatch: false,
      differenceAllowedCents: 125,
      maxDifferenceDays: 7,
      priority: 1,
    }],
  });
});

test("automation accepts current version 2 settings and rejects legacy version 1 edits", () => {
  assert.equal(normalizeAutomationSettingsPayload(managedSettings()).rules[0].ruleVersion, 2);
  assert.throws(
    () => normalizeAutomationSettingsPayload({
      ...managedSettings(),
      rules: [{ ...managedSettings().rules[0], ruleVersion: 1 }],
    }),
    /rule version/i,
  );
});

test("automation settings reject invalid managed schedule and editable rule values", () => {
  const cases = [
    ["invalid time", { schedule: { enabled: true, timeOfDay: "24:00", timeZone: AUTOMATIC_TIME_ZONE } }, /time of day/i],
    ["non-Lisbon time zone", { schedule: { enabled: true, timeOfDay: "02:15", timeZone: "UTC" } }, /time zone/i],
    ["negative tolerance", { rules: [{ ...managedSettings().rules[0], differenceAllowed: "-0.01" }] }, /non-negative amount/i],
    ["three-decimal tolerance", { rules: [{ ...managedSettings().rules[0], differenceAllowed: "1.234" }] }, /non-negative amount/i],
    ["day below zero", { rules: [{ ...managedSettings().rules[0], maxDifferenceDays: -1 }] }, /between 0 and 90/i],
    ["day above limit", { rules: [{ ...managedSettings().rules[0], maxDifferenceDays: 91 }] }, /between 0 and 90/i],
    ["unknown rule key", { rules: [{ ...managedSettings().rules[0], ruleKey: "other" }] }, /rule key/i],
    ["unknown rule version", { rules: [{ ...managedSettings().rules[0], ruleVersion: 1 }] }, /rule version/i],
    ["definition field", { rules: [{ ...managedSettings().rules[0], definition: {} }] }, /editable managed-rule fields/i],
    ["threshold field", { rules: [{ ...managedSettings().rules[0], thresholds: {} }] }, /editable managed-rule fields/i],
    ["unsupported top-level field", { unexpected: true }, /unsupported field/i],
    ["unsupported schedule field", { schedule: { ...managedSettings().schedule, unexpected: true } }, /unsupported field/i],
    ["unsupported rule field", { rules: [{ ...managedSettings().rules[0], unexpected: true }] }, /editable managed-rule fields/i],
  ];

  for (const [name, override, expected] of cases) {
    const settings = managedSettings(override);
    assert.throws(() => normalizeAutomationSettingsPayload(settings), expected, name);
  }
});

test("automation settings reject duplicate priorities and return copies", () => {
  const input = managedSettings({
    rules: [
      managedSettings().rules[0],
      { ...managedSettings().rules[0], priority: 1 },
    ],
  });
  assert.throws(() => normalizeAutomationSettingsPayload(input), /duplicate (rule )?priority/i);

  const valid = managedSettings();
  const normalized = normalizeAutomationSettingsPayload(valid);
  normalized.schedule.enabled = false;
  normalized.rules[0].priority = 9;
  assert.equal(valid.schedule.enabled, true);
  assert.equal(valid.rules[0].priority, 1);
});

test("automatic reconciliation caps managed date windows at 90 days", () => {
  assert.equal(normalizeAutomationSettingsPayload(managedSettings({
    rules: [{ ...managedSettings().rules[0], maxDifferenceDays: 90 }],
  })).rules[0].maxDifferenceDays, 90);
  assert.throws(() => normalizeAutomationSettingsPayload(managedSettings({
    rules: [{ ...managedSettings().rules[0], maxDifferenceDays: 91 }],
  })), /between 0 and 90/i);
});

test("automation actions are restricted to their public contract", () => {
  assert.equal(normalizeAutomationAction("analyze_rule"), "analyze_rule");
  assert.equal(normalizeAutomationAction("continue_analysis"), "continue_analysis");
  assert.equal(normalizeAutomationAction("execute_selected"), "execute_selected");
  assert.throws(() => normalizeAutomationAction("analyze_batch"), /automation action/i);
  assert.throws(() => normalizeAutomationAction("start"), /automation action/i);
  assert.throws(() => normalizeAutomationAction(1), /automation action/i);
});

test("continue analysis accepts only its action and run ID", () => {
  assert.deepEqual(normalizeContinueAnalysisPayload({
    action: "continue_analysis",
    runId: RUN_ID,
  }), { action: "continue_analysis", runId: RUN_ID });
  assert.throws(() => normalizeContinueAnalysisPayload({
    action: "continue_analysis",
    runId: RUN_ID,
    pageSize: 1000,
  }), /unsupported field/i);
});

test("automation run mapping exposes resumable analysis progress", () => {
  assert.deepEqual(toAutomationPublicResult({
    analysis_cursor_date: "2026-04-30",
    analysis_cursor_id: RUN_ID,
    analysis_processed: 25,
    analysis_total: 876,
    analysis_error_code: "",
    analysis_error_at: null,
  }), {
    analysisCursorDate: "2026-04-30",
    analysisCursorId: RUN_ID,
    analysisProcessed: 25,
    analysisTotal: 876,
    analysisErrorCode: "",
    analysisErrorAt: null,
  });
});

test("analysis normalizes managed rule keys and validates request UUIDs", () => {
  assert.deepEqual(normalizeAnalyzePayload({
    action: "analyze_rule",
    ruleKeys: [AUTOMATIC_RULE_KEY, AUTOMATIC_RULE_KEY],
    clientRequestId: REQUEST_ID,
  }), {
    action: "analyze_rule",
    ruleKeys: [AUTOMATIC_RULE_KEY],
    clientRequestId: REQUEST_ID,
  });

  for (const payload of [
    { action: "analyze_rule" },
    { action: "analyze_rule", ruleKeys: [] },
    { action: "analyze_rule", ruleKeys: ["other"] },
    { action: "analyze_rule", ruleKeys: [AUTOMATIC_RULE_KEY], clientRequestId: "not-a-uuid" },
    { action: "execute_selected", ruleKeys: [AUTOMATIC_RULE_KEY] },
  ]) {
    assert.throws(() => normalizeAnalyzePayload(payload), /rule key|client request id|analysis action/i);
  }
});

test("execution rejects duplicate or oversized proposal selections", () => {
  assert.throws(() => normalizeExecutePayload({
    action: "execute_selected",
    runId: RUN_ID,
    proposalIds: Array.from({ length: 101 }, (_, index) => `10000000-0000-4000-8000-${String(index + 1).padStart(12, "0")}`),
  }), /up to 100 unique proposal IDs/);
});

test("execution requires valid unique run and proposal UUIDs", () => {
  assert.deepEqual(normalizeExecutePayload({
    action: "execute_selected",
    runId: RUN_ID,
    proposalIds: [PROPOSAL_ID],
  }), { action: "execute_selected", runId: RUN_ID, proposalIds: [PROPOSAL_ID] });
  assert.deepEqual(normalizeExecutePayload({
    action: "execute_selected",
    runId: RUN_ID,
    proposalIds: [],
  }), { action: "execute_selected", runId: RUN_ID, proposalIds: [] });

  for (const payload of [
    { action: "analyze_batch", runId: RUN_ID, proposalIds: [PROPOSAL_ID] },
    { action: "execute_selected", runId: "invalid", proposalIds: [PROPOSAL_ID] },
    { action: "execute_selected", runId: RUN_ID, proposalIds: [PROPOSAL_ID, PROPOSAL_ID] },
    { action: "execute_selected", runId: RUN_ID, proposalIds: ["invalid"] },
  ]) {
    assert.throws(() => normalizeExecutePayload(payload), /automation action|execution action|run id|proposal id/i);
  }
});

test("settings RPC payload uses only managed snake-case fields and integer cents", () => {
  const settings = normalizeAutomationSettingsPayload(managedSettings());
  assert.deepEqual(toAutomationSettingsRpcPayload(settings, "user-1"), {
    p_schedule: { enabled: true, time_of_day: "02:15", time_zone: "Europe/Lisbon" },
    p_rules: [{
      rule_key: AUTOMATIC_RULE_KEY,
      rule_version: 2,
      enabled: true,
      allow_manual_execution: true,
      include_in_scheduled_batch: false,
      difference_allowed: "1.25",
      max_difference_days: 7,
      priority: 1,
    }],
    p_actor: "user-1",
  });
});

test("cron authentication accepts Vercel cron and the configured bearer secret", () => {
  assert.equal(isCronRequest({ headers: { "x-vercel-cron": "1" } }, "secret"), true);
  assert.equal(isCronRequest({ headers: { authorization: "Bearer secret" } }, "secret"), true);
  assert.equal(isCronRequest({ headers: { authorization: "Bearer wrong" } }, "secret"), false);
  assert.equal(isCronRequest({ headers: {} }, "secret"), false);
});

test("scheduled heartbeat accepts only protected GET and POST requests before any RPC", async () => {
  await withCronEnvironment("2026-08-15T02:00:00.000Z", async () => {
    for (const method of ["GET", "POST"]) {
      let rpcCalled = false;
      const response = responseRecorder();
      await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
        restQuery: async () => {
          rpcCalled = true;
          return {};
        },
      }), async (handler) => {
        await handler({ method, headers: { authorization: "Bearer wrong" } }, response);
      });
      assert.equal(response.statusCode, 401, method);
      assert.deepEqual(response.body, { error: "Unauthorized." }, method);
      assert.equal(rpcCalled, false, method);
    }

    for (const request of [
      { method: "GET", headers: { "x-vercel-cron": "1" } },
      { method: "DELETE", headers: {} },
    ]) {
      let rpcCalled = false;
      const response = responseRecorder();
      await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
        restQuery: async () => {
          rpcCalled = true;
          return {};
        },
      }), async (handler) => {
        await handler(request, response);
      });
      assert.equal(response.statusCode, 401, `${request.method} must authenticate first`);
      assert.deepEqual(response.body, { error: "Unauthorized." });
      assert.equal(rpcCalled, false);
    }

    let methodRpcCalled = false;
    const methodResponse = responseRecorder();
    await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
      restQuery: async () => {
        methodRpcCalled = true;
        return {};
      },
    }), async (handler) => {
      await handler({ method: "DELETE", headers: { authorization: `Bearer ${CRON_SECRET}` } }, methodResponse);
    });
    assert.equal(methodResponse.statusCode, 405);
    assert.equal(methodResponse.headers.Allow, "GET, POST");
    assert.equal(methodRpcCalled, false);
  });
});

test("scheduled heartbeat returns safe database reasons when disabled, not due, or complete", async () => {
  const cases = [
    ["schedule_disabled", "2026-08-15T01:00:00.000Z", {}],
    ["before_scheduled_time", "2026-08-15T01:59:00.000Z", {}],
    ["no_enabled_rules", "2026-08-15T02:00:00.000Z", {}],
    ["unsupported_rule_set", "2026-08-15T02:00:00.000Z", {}],
    ["slot_failed", "2026-08-15T02:00:00.000Z", {}],
    ["batch_complete", "2026-08-15T02:00:00.000Z", { batch_id: BATCH_ID }],
  ];

  for (const [reason, nowIso, claimFields] of cases) {
    const calls = [];
    const response = responseRecorder();
    await withCronEnvironment(nowIso, async () => {
      await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
        restQuery: async (resource, options) => {
          calls.push({ resource, options });
          return { claimed: false, reason, ...claimFields, diagnostic: "hidden schedule state" };
        },
      }), async (handler) => {
        await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, response);
      });
    });

    assert.deepEqual(calls, [{
      resource: "rpc/claim_financial_reconciliation_automatic_schedule",
      options: {
        method: "POST",
        body: { p_now: nowIso, p_actor: SCHEDULE_ACTOR },
      },
    }], reason);
    assert.equal(response.statusCode, 200, reason);
    assert.deepEqual(response.body, {
      ok: true,
      claimed: false,
      reason,
      ...(reason === "batch_complete" ? { batchId: BATCH_ID } : {}),
      hasMore: false,
    }, reason);
    assert.doesNotMatch(JSON.stringify(response.body), /hidden schedule state/);
  }
});

test("scheduled heartbeat advances one unfinished analysis page before claiming or executing", async () => {
  const calls = [];
  const response = responseRecorder();
  const progressRun = scheduledRun({
    trigger: "manual",
    scope: "rule",
    actor: "user@example.com",
    status: "analyzing",
    analysisCursorDate: "2026-01-31",
    analysisCursorId: uuidFor(31),
    analysisProcessed: 25,
    analysisTotal: 100,
    analysisComplete: false,
    analysisCompletedAt: null,
  });

  await withCronEnvironment("2026-08-15T02:00:00.000Z", async () => {
    await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
      exposeOldestAnalysis: true,
      restQuery: async (resource, options) => {
        calls.push({ resource, options });
        if (resource === "rpc/continue_financial_reconciliation_automatic_oldest_analysis") {
          return { continued: true, run: progressRun, diagnostic: "hidden" };
        }
        throw new Error(`Unexpected RPC ${resource}`);
      },
    }), async (handler) => {
      await handler({ method: "POST", headers: { authorization: `Bearer ${CRON_SECRET}` } }, response);
    });
  });

  assert.deepEqual(calls, [{
    resource: "rpc/continue_financial_reconciliation_automatic_oldest_analysis",
    options: { method: "POST", body: { p_worker: SCHEDULE_ACTOR } },
  }]);
  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, {
    ok: true,
    claimed: false,
    continuedAnalysis: true,
    runId: RUN_ID,
    status: "analyzing",
    analysisProcessed: 25,
    analysisTotal: 100,
    hasMore: true,
  });
});

test("scheduled heartbeat returns parent and rule progress when continuing a scheduled child", async () => {
  const progressRun = scheduledRun({
    status: "analyzing",
    analysisCursorDate: "2026-01-31",
    analysisCursorId: uuidFor(32),
    analysisProcessed: 25,
    analysisTotal: 100,
    analysisComplete: false,
    analysisCompletedAt: null,
  });
  const response = responseRecorder();

  await withCronEnvironment("2026-08-15T02:00:00.000Z", async () => {
    await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
      exposeOldestAnalysis: true,
      restQuery: async (resource) => {
        if (resource === "rpc/continue_financial_reconciliation_automatic_oldest_analysis") {
          return { continued: true, run: progressRun };
        }
        throw new Error(`Unexpected RPC ${resource}`);
      },
    }), async (handler) => {
      await handler({ method: "POST", headers: { authorization: `Bearer ${CRON_SECRET}` } }, response);
    });
  });

  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, {
    ok: true,
    claimed: false,
    continuedAnalysis: true,
    batchId: BATCH_ID,
    ruleKey: AUTOMATIC_RULE_KEY,
    rulePosition: 1,
    ruleCount: 2,
    runId: RUN_ID,
    status: "analyzing",
    analysisProcessed: 25,
    analysisTotal: 100,
    hasMore: true,
  });
});

test("scheduled heartbeat reports a persisted continuation failure without retrying or returning 500", async () => {
  const failedRun = scheduledRun({
    status: "failed",
    analysisComplete: false,
    analysisCompletedAt: null,
    analysisErrorCode: "analysis_continuation_failed",
    analysisErrorAt: "2026-08-15T02:00:01.000Z",
    finishedAt: "2026-08-15T02:00:01.000Z",
    proposals: [],
  });
  const calls = [];
  const response = responseRecorder();
  await withCronEnvironment("2026-08-15T02:00:00.000Z", async () => {
    await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
      exposeOldestAnalysis: true,
      restQuery: async (resource) => {
        calls.push(resource);
        if (resource === "rpc/continue_financial_reconciliation_automatic_oldest_analysis") {
          return { continued: true, run: failedRun };
        }
        throw new Error(`Unexpected RPC ${resource}`);
      },
    }), async (handler) => {
      await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, response);
    });
  });

  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, {
    ok: true,
    claimed: false,
    continuedAnalysis: true,
    batchId: BATCH_ID,
    ruleKey: AUTOMATIC_RULE_KEY,
    rulePosition: 1,
    ruleCount: 2,
    runId: failedRun.runId,
    status: "failed",
    analysisProcessed: failedRun.analysisProcessed,
    analysisTotal: failedRun.analysisTotal,
    hasMore: false,
  });
  assert.deepEqual(calls, ["rpc/continue_financial_reconciliation_automatic_oldest_analysis"]);
});

test("first scheduled analysis page can fail terminally without returning 500 or more work", async () => {
  const claimedRun = scheduledRun({
    status: "analyzing",
    analysisComplete: false,
    analysisCompletedAt: null,
    analysisProcessed: 0,
    analysisTotal: 1,
  });
  const failedRun = scheduledRun({
    status: "failed",
    analysisComplete: false,
    analysisCompletedAt: null,
    analysisProcessed: 0,
    analysisTotal: 1,
    analysisErrorCode: "analysis_continuation_failed",
    analysisErrorAt: "2026-08-15T02:00:01.000Z",
    finishedAt: "2026-08-15T02:00:01.000Z",
  });
  const calls = [];
  const response = responseRecorder();
  await withCronEnvironment("2026-08-15T02:00:00.000Z", async () => {
    await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
      restQuery: async (resource) => {
        calls.push(resource);
        if (resource === "rpc/claim_financial_reconciliation_automatic_schedule") {
          return scheduledClaim(claimedRun, { resumed: false });
        }
        if (resource === "rpc/continue_financial_reconciliation_automatic_analysis") return failedRun;
        throw new Error(`Unexpected RPC ${resource}`);
      },
    }), async (handler) => {
      await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, response);
    });
  });

  assert.equal(response.statusCode, 200);
  assert.equal(response.body.status, "failed");
  assert.equal(response.body.hasMore, false);
  assert.deepEqual(calls, [
    "rpc/claim_financial_reconciliation_automatic_schedule",
    "rpc/continue_financial_reconciliation_automatic_analysis",
  ]);
});

test("first scheduled child populates analysis and executes proposals in stable base order", async () => {
  const proposalA = uuidFor(11);
  const proposalB = uuidFor(12);
  const proposalC = uuidFor(13);
  const proposalSkipped = uuidFor(14);
  const baseA = uuidFor(101);
  const baseB = uuidFor(102);
  const baseC = uuidFor(103);
  const baseSkipped = uuidFor(104);
  const pendingRun = scheduledRun({
    status: "analyzing",
    analysisComplete: false,
    analysisCompletedAt: null,
  });
  const analyzedRun = scheduledRun({
    definitions: pendingRun.definitions,
    proposals: [
      { id: proposalA, ruleKey: AUTOMATIC_RULE_KEY, baseSourceDate: "2026-08-01", baseSourceId: baseA, status: "proposed" },
      { id: proposalB, ruleKey: AUTOMATIC_RULE_KEY, baseSourceDate: "2026-08-03", baseSourceId: baseB, status: "proposed" },
      { id: proposalC, ruleKey: AUTOMATIC_RULE_KEY, baseSourceDate: "2026-08-02", baseSourceId: baseC, status: "proposed" },
      { id: proposalSkipped, ruleKey: AUTOMATIC_RULE_KEY, baseSourceDate: "2026-08-04", baseSourceId: baseSkipped, status: "skipped", reason: "no_qualifying_combination" },
    ],
  });
  const completedRun = scheduledRun({
    definitions: pendingRun.definitions,
    status: "running",
    proposals: [
      { ...analyzedRun.proposals[0], status: "completed" },
      { ...analyzedRun.proposals[1], status: "stale", reason: "source_snapshot_changed" },
      { ...analyzedRun.proposals[2], status: "completed" },
      analyzedRun.proposals[3],
    ],
  });
  const finalizedRun = {
    ...completedRun,
    status: "partial",
    finishedAt: "2026-08-15T02:00:04.000Z",
    diagnostic: "hidden run detail",
  };
  const calls = [];
  const response = responseRecorder();

  await withCronEnvironment("2026-08-15T02:00:00.000Z", async () => {
    await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
      restQuery: async (resource, options) => {
        calls.push({ resource, options });
        if (resource === "rpc/claim_financial_reconciliation_automatic_schedule") {
          return scheduledClaim(pendingRun, { resumed: false, internal_error: "hidden claim detail" });
        }
        if (resource === "rpc/continue_financial_reconciliation_automatic_analysis") return analyzedRun;
        if (resource === "rpc/execute_financial_reconciliation_automatic_proposal") {
          const proposalId = options.body.p_proposal_id;
          return proposalId === proposalB
            ? { proposalId, runId: RUN_ID, status: "stale", reason: "source_snapshot_changed", error_detail: "hidden" }
            : { proposalId, runId: RUN_ID, status: "completed", diagnostic: "hidden" };
        }
        if (resource === "rpc/get_financial_reconciliation_automatic_run") return completedRun;
        if (resource === "rpc/finish_financial_reconciliation_automatic_run") return finalizedRun;
        throw new Error(`Unexpected RPC ${resource}`);
      },
    }), async (handler) => {
      await handler({ method: "POST", headers: { authorization: `Bearer ${CRON_SECRET}` } }, response);
    });
  });

  assert.deepEqual(calls, [
    {
      resource: "rpc/claim_financial_reconciliation_automatic_schedule",
      options: {
        method: "POST",
        body: { p_now: "2026-08-15T02:00:00.000Z", p_actor: SCHEDULE_ACTOR },
      },
    },
    {
      resource: "rpc/continue_financial_reconciliation_automatic_analysis",
      options: { method: "POST", body: { p_run_id: RUN_ID, p_actor: SCHEDULE_ACTOR } },
    },
    ...[proposalA, proposalC, proposalB].map((proposalId) => ({
      resource: "rpc/execute_financial_reconciliation_automatic_proposal",
      options: { method: "POST", body: { p_proposal_id: proposalId, p_actor: SCHEDULE_ACTOR } },
    })),
    {
      resource: "rpc/get_financial_reconciliation_automatic_run",
      options: { method: "POST", body: { p_run_id: RUN_ID } },
    },
    {
      resource: "rpc/finish_financial_reconciliation_automatic_run",
      options: { method: "POST", body: { p_run_id: RUN_ID } },
    },
  ]);
  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, {
    ok: true,
    claimed: true,
    resumed: false,
    batchId: BATCH_ID,
    ruleKey: AUTOMATIC_RULE_KEY,
    rulePosition: 1,
    ruleCount: 2,
    runId: RUN_ID,
    status: "partial",
    counts: {
      bases: 4,
      proposed: 0,
      ambiguous: 0,
      deselected: 0,
      executing: 0,
      completed: 2,
      stale: 1,
      failed: 0,
      skipped: 1,
    },
    attemptedCount: 3,
    hasMore: false,
  });
  assert.doesNotMatch(JSON.stringify(response.body), /hidden|diagnostic|error_detail|internal_error/);
});

test("terminal scheduled child stops the request and the next heartbeat claims the next rule", async () => {
  const firstRun = scheduledRun();
  const firstFinishedRun = scheduledRun({
    status: "completed",
    finishedAt: "2026-08-15T02:00:01.000Z",
  });
  const secondRun = scheduledRun({
    runId: uuidFor(702),
    batchRuleKey: CREDIT_CARD_RULE_KEY,
    batchRulePosition: 2,
    definitions: [{ ruleKey: CREDIT_CARD_RULE_KEY, priority: 2 }],
    status: "completed",
    finishedAt: "2026-08-15T02:01:01.000Z",
  });
  const calls = [];
  let heartbeat = 0;
  const firstResponse = responseRecorder();
  const secondResponse = responseRecorder();

  await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
    restQuery: async (resource) => {
      calls.push({ heartbeat, resource });
      if (resource === "rpc/claim_financial_reconciliation_automatic_schedule") {
        return heartbeat === 0
          ? scheduledClaim(firstRun, { resumed: false })
          : scheduledClaim(secondRun, { resumed: false });
      }
      if (resource === "rpc/finish_financial_reconciliation_automatic_run") return firstFinishedRun;
      throw new Error(`Unexpected RPC ${resource}`);
    },
  }), async (handler) => {
    await withCronEnvironment("2026-08-15T02:00:00.000Z", async () => {
      await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, firstResponse);
    });
    heartbeat += 1;
    await withCronEnvironment("2026-08-15T02:01:00.000Z", async () => {
      await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, secondResponse);
    });
  });

  assert.deepEqual(calls, [
    { heartbeat: 0, resource: "rpc/claim_financial_reconciliation_automatic_schedule" },
    { heartbeat: 0, resource: "rpc/finish_financial_reconciliation_automatic_run" },
    { heartbeat: 1, resource: "rpc/claim_financial_reconciliation_automatic_schedule" },
  ]);
  assert.equal(firstResponse.statusCode, 200);
  assert.deepEqual({
    batchId: firstResponse.body.batchId,
    ruleKey: firstResponse.body.ruleKey,
    rulePosition: firstResponse.body.rulePosition,
    ruleCount: firstResponse.body.ruleCount,
    hasMore: firstResponse.body.hasMore,
  }, {
    batchId: BATCH_ID,
    ruleKey: AUTOMATIC_RULE_KEY,
    rulePosition: 1,
    ruleCount: 2,
    hasMore: false,
  });
  assert.equal(secondResponse.statusCode, 200);
  assert.deepEqual({
    batchId: secondResponse.body.batchId,
    ruleKey: secondResponse.body.ruleKey,
    rulePosition: secondResponse.body.rulePosition,
    ruleCount: secondResponse.body.ruleCount,
    hasMore: secondResponse.body.hasMore,
  }, {
    batchId: BATCH_ID,
    ruleKey: CREDIT_CARD_RULE_KEY,
    rulePosition: 2,
    ruleCount: 2,
    hasMore: false,
  });
});

test("failed scheduled child returns 200 and the next heartbeat can claim the next rule", async () => {
  const failedRun = scheduledRun({
    status: "failed",
    analysisComplete: false,
    analysisCompletedAt: null,
    analysisErrorCode: "analysis_continuation_failed",
    analysisErrorAt: "2026-08-15T02:00:01.000Z",
    finishedAt: "2026-08-15T02:00:01.000Z",
  });
  const nextRun = scheduledRun({
    runId: uuidFor(703),
    batchRuleKey: CREDIT_CARD_RULE_KEY,
    batchRulePosition: 2,
    definitions: [{ ruleKey: CREDIT_CARD_RULE_KEY, priority: 2 }],
    status: "completed",
    finishedAt: "2026-08-15T02:01:01.000Z",
  });
  let heartbeat = 0;
  let claimCount = 0;
  const firstResponse = responseRecorder();
  const secondResponse = responseRecorder();

  await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
    restQuery: async (resource) => {
      if (resource !== "rpc/claim_financial_reconciliation_automatic_schedule") {
        throw new Error(`Unexpected RPC ${resource}`);
      }
      claimCount += 1;
      return heartbeat === 0 ? scheduledClaim(failedRun) : scheduledClaim(nextRun, { resumed: false });
    },
  }), async (handler) => {
    await withCronEnvironment("2026-08-15T02:00:00.000Z", async () => {
      await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, firstResponse);
    });
    heartbeat += 1;
    await withCronEnvironment("2026-08-15T02:01:00.000Z", async () => {
      await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, secondResponse);
    });
  });

  assert.equal(claimCount, 2);
  assert.equal(firstResponse.statusCode, 200);
  assert.equal(firstResponse.body.status, "failed");
  assert.equal(firstResponse.body.rulePosition, 1);
  assert.equal(firstResponse.body.hasMore, false);
  assert.equal(secondResponse.statusCode, 200);
  assert.equal(secondResponse.body.ruleKey, CREDIT_CARD_RULE_KEY);
  assert.equal(secondResponse.body.rulePosition, 2);
});

test("scheduled heartbeat finishes one resumed amount-only child before the next heartbeat claims position four", async () => {
  const amountOnlyRunId = uuidFor(722);
  const amountOnlyProposalId = uuidFor(724);
  const amountOnlyBaseId = uuidFor(725);
  const amountOnlyDefinition = amountOnlyScheduledDefinition(BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY, 3);
  const proposed = {
    id: amountOnlyProposalId,
    runId: amountOnlyRunId,
    ruleKey: BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY,
    ruleVersion: 1,
    baseSourceType: "financial_documents",
    baseSourceId: amountOnlyBaseId,
    baseSourceDate: "2026-08-14",
    baseSnapshot: { payment: "Banco", amount: 42.5 },
    items: [{ sourceType: "import_cgd_extrato_ordem", sourceId: uuidFor(726), amount: -42.5 }],
    evidence: { matchingMode: "amount_only_one_to_one" },
    candidateGroups: [{ sourceId: uuidFor(726), amount: -42.5 }],
    calculatedDifference: 0,
    allowedDifference: 0,
    status: "proposed",
    reason: null,
    signature: "amount-only-proposal-signature",
    reconciliationId: null,
    createdAt: "2026-08-15T02:02:01.000Z",
    updatedAt: "2026-08-15T02:02:01.000Z",
  };
  const analyzingRun = scheduledRun({
    runId: amountOnlyRunId,
    batchRuleKey: BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY,
    batchRulePosition: 3,
    batchRuleCount: 4,
    status: "analyzing",
    definitions: [amountOnlyDefinition],
    counts: { bases: 0, proposed: 0, ambiguous: 0, skipped: 0 },
    analysisCursorDate: "2026-08-13",
    analysisCursorId: uuidFor(727),
    analysisProcessed: 25,
    analysisTotal: 26,
    analysisComplete: false,
    analysisCompletedAt: null,
    proposals: [],
  });
  const readyRun = scheduledRun({
    ...analyzingRun,
    status: "ready",
    counts: { bases: 1, proposed: 1, ambiguous: 0, skipped: 0 },
    analysisCursorDate: null,
    analysisCursorId: null,
    analysisProcessed: 26,
    analysisTotal: 26,
    analysisComplete: true,
    analysisCompletedAt: "2026-08-15T02:02:02.000Z",
    proposals: [proposed],
  });
  const runningRun = scheduledRun({
    ...readyRun,
    status: "running",
    counts: { bases: 1, proposed: 0, ambiguous: 0, skipped: 0, completed: 1 },
    proposals: [{
      ...proposed,
      status: "completed",
      reconciliationId: uuidFor(728),
      updatedAt: "2026-08-15T02:02:03.000Z",
    }],
  });
  const completedRun = scheduledRun({
    ...runningRun,
    status: "completed",
    finishedAt: "2026-08-15T02:02:04.000Z",
  });
  const nextRun = scheduledRun({
    runId: uuidFor(723),
    batchRuleKey: CREDIT_CARD_AMOUNT_ONLY_RULE_KEY,
    batchRulePosition: 4,
    batchRuleCount: 4,
    status: "completed",
    definitions: [amountOnlyScheduledDefinition(CREDIT_CARD_AMOUNT_ONLY_RULE_KEY, 4)],
    counts: { bases: 0, proposed: 0, ambiguous: 0, skipped: 0 },
    analysisProcessed: 0,
    analysisTotal: 0,
    finishedAt: "2026-08-15T02:03:01.000Z",
  });
  let heartbeat = 0;
  const calls = [];
  const responses = [];

  await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
    restQuery: async (resource, options) => {
      calls.push({ heartbeat, resource, body: options.body });
      if (resource === "rpc/claim_financial_reconciliation_automatic_schedule") {
        return heartbeat === 0
          ? scheduledClaim(analyzingRun, { resumed: true })
          : scheduledClaim(nextRun, { resumed: false });
      }
      if (resource === "rpc/continue_financial_reconciliation_automatic_analysis") return readyRun;
      if (resource === "rpc/execute_financial_reconciliation_automatic_proposal") {
        return {
          proposalId: amountOnlyProposalId,
          runId: amountOnlyRunId,
          status: "completed",
          reconciliationId: uuidFor(728),
        };
      }
      if (resource === "rpc/get_financial_reconciliation_automatic_run") return runningRun;
      if (resource === "rpc/finish_financial_reconciliation_automatic_run") return completedRun;
      throw new Error(`Unexpected RPC ${resource}`);
    },
  }), async (handler) => {
    for (heartbeat = 0; heartbeat < 2; heartbeat += 1) {
      const response = responseRecorder();
      await withCronEnvironment(`2026-08-15T02:0${heartbeat}:00.000Z`, async () => {
        await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, response);
      });
      responses.push(response);
    }
  });

  assert.deepEqual(calls, [
    {
      heartbeat: 0,
      resource: "rpc/claim_financial_reconciliation_automatic_schedule",
      body: { p_now: "2026-08-15T02:00:00.000Z", p_actor: SCHEDULE_ACTOR },
    },
    {
      heartbeat: 0,
      resource: "rpc/continue_financial_reconciliation_automatic_analysis",
      body: { p_run_id: amountOnlyRunId, p_actor: SCHEDULE_ACTOR },
    },
    {
      heartbeat: 0,
      resource: "rpc/execute_financial_reconciliation_automatic_proposal",
      body: { p_proposal_id: amountOnlyProposalId, p_actor: SCHEDULE_ACTOR },
    },
    {
      heartbeat: 0,
      resource: "rpc/get_financial_reconciliation_automatic_run",
      body: { p_run_id: amountOnlyRunId },
    },
    {
      heartbeat: 0,
      resource: "rpc/finish_financial_reconciliation_automatic_run",
      body: { p_run_id: amountOnlyRunId },
    },
    {
      heartbeat: 1,
      resource: "rpc/claim_financial_reconciliation_automatic_schedule",
      body: { p_now: "2026-08-15T02:01:00.000Z", p_actor: SCHEDULE_ACTOR },
    },
  ]);
  assert.deepEqual(responses.map((response) => response.body), [
    {
      ok: true,
      claimed: true,
      resumed: true,
      batchId: BATCH_ID,
      ruleKey: BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY,
      rulePosition: 3,
      ruleCount: 4,
      runId: amountOnlyRunId,
      status: "completed",
      counts: {
        bases: 1,
        proposed: 0,
        ambiguous: 0,
        skipped: 0,
        deselected: 0,
        executing: 0,
        completed: 1,
        stale: 0,
        failed: 0,
      },
      attemptedCount: 1,
      hasMore: false,
    },
    {
      ok: true,
      claimed: true,
      resumed: false,
      batchId: BATCH_ID,
      ruleKey: CREDIT_CARD_AMOUNT_ONLY_RULE_KEY,
      rulePosition: 4,
      ruleCount: 4,
      runId: uuidFor(723),
      status: "completed",
      counts: {
        bases: 0,
        proposed: 0,
        ambiguous: 0,
        skipped: 0,
        deselected: 0,
        executing: 0,
        completed: 0,
        stale: 0,
        failed: 0,
      },
      attemptedCount: 0,
      hasMore: false,
    },
  ]);
});

test("five-rule heartbeat resumes monthly analysis across midnight before execute, finalize, and the next child", async () => {
  const monthlyRunId = uuidFor(741);
  const monthlyProposalId = uuidFor(742);
  const monthlyBaseId = uuidFor(743);
  const monthlyDefinition = monthlyIncomeScheduledDefinition(4);
  const monthlyProposal = {
    id: monthlyProposalId,
    ruleKey: MONTHLY_INCOME_RULE_KEY,
    ruleVersion: 2,
    baseSourceType: "import_cgd_extrato_ordem",
    baseSourceId: monthlyBaseId,
    baseSourceDate: "2026-03-02",
    groupingKey: "2026-03",
    summarySnapshot: {
      calendarMonth: "2026-03-01",
      sourceCount: 1000,
      sourceTotal: 2000,
      destinationCount: 1000,
      destinationTotal: 1250,
      totalCount: 2000,
      calculatedDifference: 750,
    },
    items: [],
    candidateGroups: [],
    calculatedDifference: 750,
    allowedDifference: 7500,
    status: "proposed",
    reason: null,
    signature: "monthly-scheduled-signature",
    reconciliationId: null,
    createdAt: "2026-08-15T02:00:01.000Z",
    updatedAt: "2026-08-15T02:00:01.000Z",
  };
  const claimedAnalyzingRun = scheduledRun({
    runId: monthlyRunId,
    batchRuleKey: MONTHLY_INCOME_RULE_KEY,
    batchRulePosition: 4,
    batchRuleCount: 5,
    status: "analyzing",
    definitions: [monthlyDefinition],
    analysisCursorDate: null,
    analysisCursorId: null,
    analysisProcessed: 0,
    analysisTotal: 3,
    analysisComplete: false,
    analysisCompletedAt: null,
  });
  const firstPageRun = scheduledRun({
    ...claimedAnalyzingRun,
    analysisCursorDate: "2026-01-01",
    analysisCursorId: uuidFor(744),
    analysisProcessed: 1,
  });
  const crossMidnightRun = scheduledRun({
    ...firstPageRun,
    analysisCursorDate: "2026-02-01",
    analysisCursorId: uuidFor(745),
    analysisProcessed: 2,
  });
  const readyRun = scheduledRun({
    ...crossMidnightRun,
    status: "ready",
    analysisCursorDate: null,
    analysisCursorId: null,
    analysisProcessed: 3,
    analysisComplete: true,
    analysisCompletedAt: "2026-08-16T00:02:00.000Z",
    proposals: [monthlyProposal],
  });
  const runningRun = scheduledRun({
    ...readyRun,
    status: "running",
    proposals: [{
      ...monthlyProposal,
      status: "completed",
      reconciliationId: uuidFor(746),
      updatedAt: "2026-08-16T00:03:01.000Z",
    }],
  });
  const completedRun = scheduledRun({
    ...runningRun,
    status: "completed",
    finishedAt: "2026-08-16T00:03:02.000Z",
  });
  const nextRun = scheduledRun({
    runId: uuidFor(747),
    batchRuleKey: BANK_STATEMENT_RULE_KEY,
    batchRulePosition: 5,
    batchRuleCount: 5,
    status: "completed",
    definitions: [{
      ruleKey: BANK_STATEMENT_RULE_KEY,
      ruleVersion: BANK_STATEMENT_RULE_VERSION,
      priority: 5,
    }],
    analysisProcessed: 0,
    analysisTotal: 0,
    finishedAt: "2026-08-16T00:04:00.000Z",
  });
  const oldestResults = [
    { continued: false },
    { continued: true, run: crossMidnightRun, diagnostic: "hidden oldest cursor" },
    { continued: true, run: readyRun, error_detail: "hidden ready detail" },
    { continued: false },
    { continued: false },
  ];
  const calls = [];
  const responses = [];
  let heartbeat = 0;

  await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
    exposeOldestAnalysis: true,
    restQuery: async (resource, options) => {
      calls.push({ heartbeat, resource, body: options.body });
      if (resource === "rpc/continue_financial_reconciliation_automatic_oldest_analysis") {
        return oldestResults[heartbeat];
      }
      if (resource === "rpc/claim_financial_reconciliation_automatic_schedule") {
        if (heartbeat === 0) return scheduledClaim(claimedAnalyzingRun, { resumed: false });
        if (heartbeat === 3) return scheduledClaim(readyRun, { resumed: true });
        if (heartbeat === 4) return scheduledClaim(nextRun, { resumed: false });
      }
      if (resource === "rpc/continue_financial_reconciliation_automatic_analysis") return firstPageRun;
      if (resource === "rpc/execute_financial_reconciliation_automatic_proposal") {
        return {
          proposalId: monthlyProposalId,
          runId: monthlyRunId,
          status: "completed",
          reconciliationId: uuidFor(746),
        };
      }
      if (resource === "rpc/get_financial_reconciliation_automatic_run") return runningRun;
      if (resource === "rpc/finish_financial_reconciliation_automatic_run") return completedRun;
      throw new Error(`Unexpected RPC ${resource}`);
    },
  }), async (handler) => {
    for (const nowIso of [
      "2026-08-15T23:58:00.000Z",
      "2026-08-16T00:01:00.000Z",
      "2026-08-16T00:02:00.000Z",
      "2026-08-16T00:03:00.000Z",
      "2026-08-16T00:04:00.000Z",
    ]) {
      const response = responseRecorder();
      await withCronEnvironment(nowIso, async () => {
        await handler({ method: "POST", headers: { authorization: `Bearer ${CRON_SECRET}` } }, response);
      });
      responses.push(response);
      heartbeat += 1;
    }
  });

  assert.deepEqual(
    calls.filter((call) => call.resource === "rpc/claim_financial_reconciliation_automatic_schedule")
      .map((call) => call.heartbeat),
    [0, 3, 4],
    "no second child may be claimed while monthly analysis is unfinished or merely ready",
  );
  assert.deepEqual(calls.map((call) => `${call.heartbeat}:${call.resource}`), [
    "0:rpc/continue_financial_reconciliation_automatic_oldest_analysis",
    "0:rpc/claim_financial_reconciliation_automatic_schedule",
    "0:rpc/continue_financial_reconciliation_automatic_analysis",
    "1:rpc/continue_financial_reconciliation_automatic_oldest_analysis",
    "2:rpc/continue_financial_reconciliation_automatic_oldest_analysis",
    "3:rpc/continue_financial_reconciliation_automatic_oldest_analysis",
    "3:rpc/claim_financial_reconciliation_automatic_schedule",
    "3:rpc/execute_financial_reconciliation_automatic_proposal",
    "3:rpc/get_financial_reconciliation_automatic_run",
    "3:rpc/finish_financial_reconciliation_automatic_run",
    "4:rpc/continue_financial_reconciliation_automatic_oldest_analysis",
    "4:rpc/claim_financial_reconciliation_automatic_schedule",
  ]);
  assert.deepEqual(responses.map((response) => response.statusCode), [200, 200, 200, 200, 200]);
  assert.deepEqual(responses.map((response) => ({
    claimed: response.body.claimed,
    continuedAnalysis: response.body.continuedAnalysis || false,
    ruleKey: response.body.ruleKey,
    rulePosition: response.body.rulePosition,
    ruleCount: response.body.ruleCount,
    runId: response.body.runId,
    status: response.body.status,
    attemptedCount: response.body.attemptedCount,
    hasMore: response.body.hasMore,
  })), [
    {
      claimed: true,
      continuedAnalysis: false,
      ruleKey: MONTHLY_INCOME_RULE_KEY,
      rulePosition: 4,
      ruleCount: 5,
      runId: monthlyRunId,
      status: "analyzing",
      attemptedCount: 0,
      hasMore: true,
    },
    {
      claimed: false,
      continuedAnalysis: true,
      ruleKey: MONTHLY_INCOME_RULE_KEY,
      rulePosition: 4,
      ruleCount: 5,
      runId: monthlyRunId,
      status: "analyzing",
      attemptedCount: undefined,
      hasMore: true,
    },
    {
      claimed: false,
      continuedAnalysis: true,
      ruleKey: MONTHLY_INCOME_RULE_KEY,
      rulePosition: 4,
      ruleCount: 5,
      runId: monthlyRunId,
      status: "ready",
      attemptedCount: undefined,
      hasMore: false,
    },
    {
      claimed: true,
      continuedAnalysis: false,
      ruleKey: MONTHLY_INCOME_RULE_KEY,
      rulePosition: 4,
      ruleCount: 5,
      runId: monthlyRunId,
      status: "completed",
      attemptedCount: 1,
      hasMore: false,
    },
    {
      claimed: true,
      continuedAnalysis: false,
      ruleKey: BANK_STATEMENT_RULE_KEY,
      rulePosition: 5,
      ruleCount: 5,
      runId: nextRun.runId,
      status: "completed",
      attemptedCount: 0,
      hasMore: false,
    },
  ]);
  assert.deepEqual(responses[3].body.counts, {
    bases: 1,
    proposed: 0,
    ambiguous: 0,
    skipped: 0,
    deselected: 0,
    executing: 0,
    completed: 1,
    stale: 0,
    failed: 0,
  });
  assert.doesNotMatch(JSON.stringify(responses.map((response) => response.body)), /hidden|diagnostic|error_detail/);
});

test("monthly scheduled continuation exposes only its sanitized terminal failure", async () => {
  const failedRun = scheduledRun({
    runId: uuidFor(751),
    batchRuleKey: MONTHLY_INCOME_RULE_KEY,
    batchRulePosition: 2,
    batchRuleCount: 5,
    definitions: [monthlyIncomeScheduledDefinition(2)],
    status: "failed",
    analysisCursorDate: "2026-02-01",
    analysisCursorId: uuidFor(752),
    analysisProcessed: 2,
    analysisTotal: 3,
    analysisComplete: false,
    analysisCompletedAt: null,
    analysisErrorCode: "analysis_continuation_failed",
    analysisErrorAt: "2026-08-16T00:05:00.000Z",
    finishedAt: "2026-08-16T00:05:00.000Z",
    errorSummary: "secret monthly database failure",
    errorDetail: "secret monthly stack",
  });
  const calls = [];
  const response = responseRecorder();

  await withCronEnvironment("2026-08-16T00:05:00.000Z", async () => {
    await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
      exposeOldestAnalysis: true,
      restQuery: async (resource, options) => {
        calls.push({ resource, options });
        if (resource === "rpc/continue_financial_reconciliation_automatic_oldest_analysis") {
          return { continued: true, run: failedRun, diagnostic: "secret continuation diagnostic" };
        }
        throw new Error(`Unexpected RPC ${resource}`);
      },
    }), async (handler) => {
      await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, response);
    });
  });

  assert.deepEqual(calls, [{
    resource: "rpc/continue_financial_reconciliation_automatic_oldest_analysis",
    options: { method: "POST", body: { p_worker: SCHEDULE_ACTOR } },
  }]);
  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, {
    ok: true,
    claimed: false,
    continuedAnalysis: true,
    batchId: BATCH_ID,
    ruleKey: MONTHLY_INCOME_RULE_KEY,
    rulePosition: 2,
    ruleCount: 5,
    runId: failedRun.runId,
    status: "failed",
    analysisProcessed: 2,
    analysisTotal: 3,
    hasMore: false,
  });
  assert.doesNotMatch(JSON.stringify(response.body), /secret|diagnostic|stack|errorSummary|errorDetail/);
});

test("scheduled heartbeat rejects a monthly child whose snapshot is not the exact managed version", async () => {
  const response = responseRecorder();
  const wrongVersionRun = scheduledRun({
    runId: uuidFor(761),
    batchRuleKey: MONTHLY_INCOME_RULE_KEY,
    batchRulePosition: 4,
    batchRuleCount: 5,
    definitions: [{ ...monthlyIncomeScheduledDefinition(4), ruleVersion: 1 }],
    status: "completed",
    finishedAt: "2026-08-16T00:06:00.000Z",
  });

  await withCronEnvironment("2026-08-16T00:06:00.000Z", async () => {
    await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
      restQuery: async (resource) => {
        if (resource === "rpc/claim_financial_reconciliation_automatic_schedule") {
          return scheduledClaim(wrongVersionRun, { resumed: false });
        }
        throw new Error(`Unexpected RPC ${resource}`);
      },
    }), async (handler) => {
      await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, response);
    });
  });

  assert.equal(response.statusCode, 500);
  assert.deepEqual(response.body, { error: "Unexpected server error." });
});

test("scheduled heartbeat fails closed on malformed parent or one-rule child metadata", async () => {
  const proposal = {
    id: uuidFor(704),
    ruleKey: AUTOMATIC_RULE_KEY,
    baseSourceDate: "2026-08-01",
    baseSourceId: uuidFor(705),
    status: "proposed",
  };
  const validRun = scheduledRun({ proposals: [proposal] });
  const cases = [
    ["legacy batch scope", scheduledRun({ scope: "batch", proposals: [proposal] }), {}],
    ["invalid batch id", scheduledRun({ batchId: "not-a-uuid", proposals: [proposal] }), {}],
    ["invalid batch position", scheduledRun({ batchRulePosition: 0, proposals: [proposal] }), {}],
    ["position exceeds count", scheduledRun({ batchRulePosition: 3, proposals: [proposal] }), {}],
    ["batch rule differs from definition", scheduledRun({ batchRuleKey: CREDIT_CARD_RULE_KEY, proposals: [proposal] }), {}],
    ["more than one definition", scheduledRun({
      definitions: [
        { ruleKey: AUTOMATIC_RULE_KEY, priority: 1 },
        { ruleKey: CREDIT_CARD_RULE_KEY, priority: 2 },
      ],
      proposals: [proposal],
    }), {}],
    ["claim batch differs from run", validRun, { batchId: uuidFor(706) }],
    ["claim position differs from run", validRun, { batchRulePosition: 2 }],
    ["claim count differs from run", validRun, { batchRuleCount: 3 }],
  ];

  for (const [name, run, claimOverrides] of cases) {
    let executionCalled = false;
    const response = responseRecorder();
    await withCronEnvironment("2026-08-15T02:00:00.000Z", async () => {
      await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
        restQuery: async (resource) => {
          if (resource === "rpc/claim_financial_reconciliation_automatic_schedule") {
            return scheduledClaim(run, claimOverrides);
          }
          if (resource === "rpc/execute_financial_reconciliation_automatic_proposal") {
            executionCalled = true;
          }
          return run;
        },
      }), async (handler) => {
        await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, response);
      });
    });
    assert.equal(response.statusCode, 500, name);
    assert.deepEqual(response.body, { error: "Unexpected server error." }, name);
    assert.equal(executionCalled, false, name);
  }
});

test("scheduled heartbeat rejects foreign-run or unknown-rule proposals before execution", async () => {
  const cases = [
    ["foreign run", {
      id: uuidFor(20),
      runId: uuidFor(900),
      ruleKey: AUTOMATIC_RULE_KEY,
      baseSourceDate: "2026-08-01",
      baseSourceId: uuidFor(120),
      status: "proposed",
    }],
    ["unknown rule", {
      id: uuidFor(21),
      runId: RUN_ID,
      ruleKey: "unknown_rule",
      baseSourceDate: "2026-08-01",
      baseSourceId: uuidFor(121),
      status: "proposed",
    }],
  ];

  for (const [name, proposal] of cases) {
    let executionCalled = false;
    const response = responseRecorder();
    const run = scheduledRun({ proposals: [proposal] });
    await withCronEnvironment("2026-08-15T02:00:00.000Z", async () => {
      await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
        restQuery: async (resource) => {
          if (resource === "rpc/claim_financial_reconciliation_automatic_schedule") {
            return scheduledClaim(run);
          }
          if (resource === "rpc/execute_financial_reconciliation_automatic_proposal") executionCalled = true;
          return run;
        },
      }), async (handler) => {
        await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, response);
      });
    });
    assert.equal(response.statusCode, 500, name);
    assert.deepEqual(response.body, { error: "Unexpected server error." }, name);
    assert.equal(executionCalled, false, name);
  }
});

test("scheduled heartbeat preserves the claimed run identity across every follow-up RPC", async () => {
  const otherRun = scheduledRun({
    runId: uuidFor(901),
    status: "completed",
    finishedAt: "2026-08-15T02:00:10.000Z",
  });
  const cases = [
    ["continue", scheduledRun({ status: "analyzing", analysisComplete: false, analysisCompletedAt: null }), "rpc/continue_financial_reconciliation_automatic_analysis"],
    ["refresh", scheduledRun({ proposals: [{
      id: uuidFor(22),
      ruleKey: AUTOMATIC_RULE_KEY,
      baseSourceDate: "2026-08-01",
      baseSourceId: uuidFor(122),
      status: "proposed",
    }] }), "rpc/get_financial_reconciliation_automatic_run"],
    ["finalize", scheduledRun(), "rpc/finish_financial_reconciliation_automatic_run"],
  ];

  for (const [name, claimedRun, driftingRpc] of cases) {
    const response = responseRecorder();
    await withCronEnvironment("2026-08-15T02:00:00.000Z", async () => {
      await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
        restQuery: async (resource, options) => {
          if (resource === "rpc/claim_financial_reconciliation_automatic_schedule") {
            return scheduledClaim(claimedRun);
          }
          if (resource === driftingRpc) return otherRun;
          if (resource === "rpc/execute_financial_reconciliation_automatic_proposal") {
            return { proposalId: options.body.p_proposal_id, status: "completed" };
          }
          throw new Error(`Unexpected RPC ${resource}`);
        },
      }), async (handler) => {
        await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, response);
      });
    });
    assert.equal(response.statusCode, 500, name);
    assert.deepEqual(response.body, { error: "Unexpected server error." }, name);
  }
});

test("scheduled heartbeat rejects inconsistent run lifecycle state before mutation", async () => {
  const cases = [
    ["unknown status", scheduledRun({ status: "unknown" })],
    ["invalid analysis timestamp", scheduledRun({ analysisCompletedAt: "not-a-timestamp" })],
    ["non-contract analysis timestamp", scheduledRun({ analysisCompletedAt: "1" })],
    ["finished run with pending work", scheduledRun({
      status: "completed",
      finishedAt: "2026-08-15T02:00:10.000Z",
      proposals: [{
        id: uuidFor(23),
        ruleKey: AUTOMATIC_RULE_KEY,
        baseSourceDate: "2026-08-01",
        baseSourceId: uuidFor(123),
        status: "proposed",
      }],
    })],
  ];

  for (const [name, run] of cases) {
    let mutationCalled = false;
    const response = responseRecorder();
    await withCronEnvironment("2026-08-15T02:00:00.000Z", async () => {
      await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
        restQuery: async (resource) => {
          if (resource === "rpc/claim_financial_reconciliation_automatic_schedule") {
            return scheduledClaim(run);
          }
          mutationCalled = true;
          return run;
        },
      }), async (handler) => {
        await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, response);
      });
    });
    assert.equal(response.statusCode, 500, name);
    assert.deepEqual(response.body, { error: "Unexpected server error." }, name);
    assert.equal(mutationCalled, false, name);
  }
});

test("scheduled heartbeat enforces refresh and finalize phase postconditions", async () => {
  const proposal = {
    id: uuidFor(24),
    ruleKey: AUTOMATIC_RULE_KEY,
    baseSourceDate: "2026-08-01",
    baseSourceId: uuidFor(124),
    status: "proposed",
  };
  const analyzingRun = scheduledRun({ status: "analyzing", analysisComplete: false, analysisCompletedAt: null });
  const terminalRun = scheduledRun({
    status: "completed",
    finishedAt: "2026-08-15T02:00:10.000Z",
  });
  const cases = [
    {
      name: "refresh loses analysis completion",
      claimedRun: scheduledRun({ proposals: [proposal] }),
      responses: {
        "rpc/execute_financial_reconciliation_automatic_proposal": { proposalId: proposal.id, status: "completed" },
        "rpc/get_financial_reconciliation_automatic_run": analyzingRun,
        "rpc/finish_financial_reconciliation_automatic_run": terminalRun,
      },
      expectedCalls: [
        "rpc/claim_financial_reconciliation_automatic_schedule",
        "rpc/execute_financial_reconciliation_automatic_proposal",
        "rpc/get_financial_reconciliation_automatic_run",
      ],
    },
    {
      name: "finalize remains unfinished",
      claimedRun: scheduledRun(),
      responses: {
        "rpc/finish_financial_reconciliation_automatic_run": scheduledRun(),
      },
      expectedCalls: [
        "rpc/claim_financial_reconciliation_automatic_schedule",
        "rpc/finish_financial_reconciliation_automatic_run",
      ],
    },
  ];

  for (const { name, claimedRun, responses, expectedCalls } of cases) {
    const calls = [];
    const response = responseRecorder();
    await withCronEnvironment("2026-08-15T02:00:00.000Z", async () => {
      await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
        restQuery: async (resource) => {
          calls.push(resource);
          if (resource === "rpc/claim_financial_reconciliation_automatic_schedule") {
            return scheduledClaim(claimedRun);
          }
          if (Object.hasOwn(responses, resource)) return responses[resource];
          throw new Error(`Unexpected RPC ${resource}`);
        },
      }), async (handler) => {
        await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, response);
      });
    });
    assert.equal(response.statusCode, 500, name);
    assert.deepEqual(response.body, { error: "Unexpected server error." }, name);
    assert.deepEqual(calls, expectedCalls, name);
  }
});

test("scheduled heartbeat rejects prototype-backed run fields from an RPC response", async () => {
  const inheritedRunFields = {
    runId: RUN_ID,
    trigger: "scheduled",
    scope: "rule",
    status: "completed",
    actor: SCHEDULE_ACTOR,
    batchId: BATCH_ID,
    batchRuleKey: AUTOMATIC_RULE_KEY,
    batchRulePosition: 1,
    batchRuleCount: 2,
    analysisCompletedAt: "2026-08-15T02:00:01.000Z",
    finishedAt: "2026-08-15T02:00:10.000Z",
  };
  const poisonedRun = Object.assign(Object.create(inheritedRunFields), {
    definitions: [{ ruleKey: AUTOMATIC_RULE_KEY, priority: 1 }],
    proposals: [],
  });
  let followUpCalled = false;
  const response = responseRecorder();

  await withCronEnvironment("2026-08-15T02:00:00.000Z", async () => {
    await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
      restQuery: async (resource) => {
        if (resource === "rpc/claim_financial_reconciliation_automatic_schedule") {
          return scheduledClaim(poisonedRun);
        }
        followUpCalled = true;
        return {};
      },
    }), async (handler) => {
      await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, response);
    });
  });

  assert.equal(response.statusCode, 500);
  assert.deepEqual(response.body, { error: "Unexpected server error." });
  assert.equal(followUpCalled, false);
});

test("scheduled heartbeat rejects prototype-backed claim fields from an RPC response", async () => {
  const finishedRun = scheduledRun({
    status: "completed",
    finishedAt: "2026-08-15T02:00:10.000Z",
  });
  const poisonedClaim = Object.create(scheduledClaim(finishedRun));
  let followUpCalled = false;
  const response = responseRecorder();

  await withCronEnvironment("2026-08-15T02:00:00.000Z", async () => {
    await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
      restQuery: async (resource) => {
        if (resource === "rpc/claim_financial_reconciliation_automatic_schedule") return poisonedClaim;
        followUpCalled = true;
        return {};
      },
    }), async (handler) => {
      await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, response);
    });
  });

  assert.equal(response.statusCode, 500);
  assert.deepEqual(response.body, { error: "Unexpected server error." });
  assert.equal(followUpCalled, false);
});

test("scheduled heartbeat caps sequential proposal attempts at 25 and leaves the run resumable", async () => {
  const proposals = Array.from({ length: 27 }, (_, index) => ({
    id: uuidFor(200 + index),
    ruleKey: AUTOMATIC_RULE_KEY,
    baseSourceDate: "2026-08-10",
    baseSourceId: index === 0
      ? "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"
      : index === 1
        ? "BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB"
        : uuidFor(300 + index),
    status: "proposed",
  })).reverse();
  const expectedIds = [...proposals]
    .sort((left, right) => compareCodeUnits(left.baseSourceId, right.baseSourceId)
      || compareCodeUnits(left.id, right.id))
    .slice(0, 25)
    .map((proposal) => proposal.id);
  const completedIds = new Set();
  const calls = [];
  let activeExecutions = 0;
  let maximumActiveExecutions = 0;
  const response = responseRecorder();

  await withCronEnvironment("2026-08-15T02:01:00.000Z", async () => {
    await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
      restQuery: async (resource, options) => {
        calls.push({ resource, options });
        if (resource === "rpc/claim_financial_reconciliation_automatic_schedule") {
          const run = scheduledRun({ proposals });
          return scheduledClaim(run);
        }
        if (resource === "rpc/execute_financial_reconciliation_automatic_proposal") {
          activeExecutions += 1;
          maximumActiveExecutions = Math.max(maximumActiveExecutions, activeExecutions);
          await new Promise((resolve) => setImmediate(resolve));
          completedIds.add(options.body.p_proposal_id);
          activeExecutions -= 1;
          return { proposalId: options.body.p_proposal_id, status: "completed" };
        }
        if (resource === "rpc/get_financial_reconciliation_automatic_run") {
          return scheduledRun({
            status: "running",
            proposals: proposals.map((proposal) => ({
              ...proposal,
              status: completedIds.has(proposal.id) ? "completed" : "proposed",
            })),
          });
        }
        if (resource === "rpc/finish_financial_reconciliation_automatic_run") {
          throw new Error("Run with pending work must not be finalized.");
        }
        throw new Error(`Unexpected RPC ${resource}`);
      },
    }), async (handler) => {
      await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, response);
    });
  });

  const executionIds = calls
    .filter(({ resource }) => resource === "rpc/execute_financial_reconciliation_automatic_proposal")
    .map(({ options }) => options.body.p_proposal_id);
  assert.deepEqual(executionIds, expectedIds);
  assert.equal(executionIds.length, 25);
  assert.equal(maximumActiveExecutions, 1);
  assert.equal(calls.some(({ resource }) => resource === "rpc/finish_financial_reconciliation_automatic_run"), false);
  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body.counts, {
    bases: 27,
    proposed: 2,
    ambiguous: 0,
    skipped: 0,
    deselected: 0,
    executing: 0,
    completed: 25,
    stale: 0,
    failed: 0,
  });
  assert.equal(response.body.attemptedCount, 25);
  assert.equal(response.body.hasMore, true);
});

test("scheduled heartbeat continues after an isolated failure and finalizes only after a safe resume", async () => {
  const firstProposal = uuidFor(401);
  const secondProposal = uuidFor(402);
  let heartbeat = 0;
  let firstProposalAttempt = 0;
  let secondCompleted = false;
  let firstFailed = false;
  const calls = [];
  const firstResponse = responseRecorder();
  const secondResponse = responseRecorder();

  await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
    restQuery: async (resource, options) => {
      calls.push({ heartbeat, resource, options });
      const currentProposals = [
        {
          id: firstProposal,
          ruleKey: AUTOMATIC_RULE_KEY,
          baseSourceDate: "2026-08-10",
          baseSourceId: uuidFor(501),
          status: firstFailed ? "failed" : "proposed",
          error_detail: "database stack hidden",
        },
        {
          id: secondProposal,
          ruleKey: AUTOMATIC_RULE_KEY,
          baseSourceDate: "2026-08-11",
          baseSourceId: uuidFor(502),
          status: secondCompleted ? "completed" : "proposed",
        },
      ];
      if (resource === "rpc/claim_financial_reconciliation_automatic_schedule") {
        const run = scheduledRun({ proposals: currentProposals });
        return scheduledClaim(run);
      }
      if (resource === "rpc/execute_financial_reconciliation_automatic_proposal") {
        if (options.body.p_proposal_id === firstProposal) {
          firstProposalAttempt += 1;
          if (firstProposalAttempt === 1) throw new Error("secret transport diagnostic");
          firstFailed = true;
          return { proposalId: firstProposal, status: "failed", reason: "execution_failed", stack: "hidden" };
        }
        secondCompleted = true;
        return { proposalId: secondProposal, status: "completed" };
      }
      if (resource === "rpc/get_financial_reconciliation_automatic_run") {
        return scheduledRun({ status: "running", proposals: currentProposals.map((proposal) => {
          if (proposal.id === firstProposal) return { ...proposal, status: firstFailed ? "failed" : "proposed" };
          return { ...proposal, status: secondCompleted ? "completed" : "proposed" };
        }) });
      }
      if (resource === "rpc/finish_financial_reconciliation_automatic_run") {
        return scheduledRun({
          status: "partial",
          finishedAt: "2026-08-15T02:02:05.000Z",
          proposals: [
            { ...currentProposals[0], status: "failed", reason: "execution_failed" },
            { ...currentProposals[1], status: "completed" },
          ],
        });
      }
      throw new Error(`Unexpected RPC ${resource}`);
    },
  }), async (handler) => {
    await withCronEnvironment("2026-08-15T02:01:00.000Z", async () => {
      await handler({ method: "POST", headers: { authorization: `Bearer ${CRON_SECRET}` } }, firstResponse);
    });
    heartbeat += 1;
    await withCronEnvironment("2026-08-15T02:02:00.000Z", async () => {
      await handler({ method: "POST", headers: { authorization: `Bearer ${CRON_SECRET}` } }, secondResponse);
    });
  });

  const firstHeartbeatCalls = calls.filter((call) => call.heartbeat === 0).map((call) => call.resource);
  const secondHeartbeatCalls = calls.filter((call) => call.heartbeat === 1).map((call) => call.resource);
  assert.deepEqual(firstHeartbeatCalls, [
    "rpc/claim_financial_reconciliation_automatic_schedule",
    "rpc/execute_financial_reconciliation_automatic_proposal",
    "rpc/execute_financial_reconciliation_automatic_proposal",
    "rpc/get_financial_reconciliation_automatic_run",
  ]);
  assert.deepEqual(secondHeartbeatCalls, [
    "rpc/claim_financial_reconciliation_automatic_schedule",
    "rpc/execute_financial_reconciliation_automatic_proposal",
    "rpc/get_financial_reconciliation_automatic_run",
    "rpc/finish_financial_reconciliation_automatic_run",
  ]);
  assert.equal(firstResponse.body.hasMore, true);
  assert.equal(firstResponse.body.counts.proposed, 1);
  assert.equal(firstResponse.body.counts.completed, 1);
  assert.equal(secondResponse.body.hasMore, false);
  assert.equal(secondResponse.body.counts.failed, 1);
  assert.equal(secondResponse.body.counts.completed, 1);
  assert.doesNotMatch(JSON.stringify([firstResponse.body, secondResponse.body]), /secret|diagnostic|stack|error_detail/);
});

test("scheduled heartbeat delegates Lisbon DST slot identity to the database claim", async () => {
  const calls = [];
  const firstResponse = responseRecorder();
  const secondResponse = responseRecorder();
  let claimCount = 0;
  const finishedRun = scheduledRun({
    status: "completed",
    scheduledSlot: "2026-03-29",
    finishedAt: "2026-03-29T00:31:00.000Z",
    diagnostic: "hidden",
  });

  await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
    restQuery: async (resource, options) => {
      calls.push({ resource, options });
      claimCount += 1;
      return scheduledClaim(finishedRun, { resumed: claimCount > 1 });
    },
  }), async (handler) => {
    await withCronEnvironment("2026-03-29T00:30:00.000Z", async () => {
      await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, firstResponse);
    });
    await withCronEnvironment("2026-03-29T01:30:00.000Z", async () => {
      await handler({ method: "POST", headers: { authorization: `Bearer ${CRON_SECRET}` } }, secondResponse);
    });
  });

  assert.deepEqual(calls, [
    {
      resource: "rpc/claim_financial_reconciliation_automatic_schedule",
      options: {
        method: "POST",
        body: { p_now: "2026-03-29T00:30:00.000Z", p_actor: SCHEDULE_ACTOR },
      },
    },
    {
      resource: "rpc/claim_financial_reconciliation_automatic_schedule",
      options: {
        method: "POST",
        body: { p_now: "2026-03-29T01:30:00.000Z", p_actor: SCHEDULE_ACTOR },
      },
    },
  ]);
  assert.equal(firstResponse.body.runId, RUN_ID);
  assert.equal(firstResponse.body.resumed, false);
  assert.equal(secondResponse.body.runId, RUN_ID);
  assert.equal(secondResponse.body.resumed, true);
  assert.equal(firstResponse.body.hasMore, false);
  assert.equal(secondResponse.body.hasMore, false);
  assert.doesNotMatch(JSON.stringify([firstResponse.body, secondResponse.body]), /diagnostic|scheduledSlot/);
});

test("scheduled heartbeat never exposes unexpected database diagnostics", async () => {
  const response = responseRecorder();
  await withCronEnvironment("2026-08-15T02:00:00.000Z", async () => {
    await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
      restQuery: async () => {
        const error = new Error("relation internal_schedule_secret does not exist");
        error.supabasePayload = { details: "database credentials hidden" };
        throw error;
      },
    }), async (handler) => {
      await handler({ method: "GET", headers: { authorization: `Bearer ${CRON_SECRET}` } }, response);
    });
  });
  assert.equal(response.statusCode, 500);
  assert.deepEqual(response.body, { error: "Unexpected server error." });
  assert.doesNotMatch(JSON.stringify(response.body), /internal_schedule_secret|credentials|relation/);
});

test("deployment config and README describe the protected reconciliation heartbeat rollout", () => {
  const vercelConfig = JSON.parse(fs.readFileSync(VERCEL_CONFIG_PATH, "utf8"));
  assert.equal(vercelConfig.crons.filter((cron) => cron.path === "/api/reconciliation-automation-cron").length, 1);
  assert.deepEqual(
    vercelConfig.crons.find((cron) => cron.path === "/api/reconciliation-automation-cron"),
    { path: "/api/reconciliation-automation-cron", schedule: "* * * * *" },
  );

  const readme = fs.readFileSync(README_PATH, "utf8");
  const migrationNames = [
    "2026-08-14-financial-reconciliation-automation-schema.sql",
    "2026-08-14-financial-reconciliation-automation-analysis.sql",
    "2026-08-14-financial-reconciliation-automation-execution.sql",
  ];
  let previousIndex = -1;
  for (const migrationName of migrationNames) {
    const migrationIndex = readme.indexOf(migrationName);
    assert.ok(migrationIndex > previousIndex, `${migrationName} must be documented in migration order`);
    previousIndex = migrationIndex;
  }
  assert.match(readme, /CRON_SECRET/);
  assert.match(readme, /Authorization[^\n]*Bearer|Bearer[^\n]*Authorization/i);
  assert.doesNotMatch(readme, /signed cron header|x-vercel-cron/i);
  assert.match(readme, /every minute/i);
  assert.match(readme, /once[^\n]*daily[^\n]*Europe\/Lisbon|Europe\/Lisbon[^\n]*once[^\n]*daily/i);
  assert.match(readme, /disabled by default/i);
  assert.match(readme, /manual[^\n]*validat[^\n]*before[^\n]*scheduled|before[^\n]*scheduled[^\n]*manual[^\n]*validat/i);
  assert.match(readme, /Settings[^\n]*last[^\n]*batch|last[^\n]*batch[^\n]*Settings/i);
});

test("automation public result recursively maps known fields and strips diagnostics", () => {
  const input = {
    schedule: { time_of_day: "02:15", time_zone: "Europe/Lisbon", diagnostic: "remove" },
    rules: [{ rule_key: AUTOMATIC_RULE_KEY, rule_version: 1, unknown_key: "kept" }],
    proposals: [{
      run_id: RUN_ID,
      candidate_groups: [{ base_source_id: "doc-1", items: [{ source_id: "bank-1", error_detail: "remove" }] }],
      error_summary: "remove",
    }],
    definition: { destination_source_types: ["import_cgd_extrato_ordem"], internal_error: "remove" },
    reason: "public",
  };

  assert.deepEqual(toAutomationPublicResult(input), {
    schedule: { timeOfDay: "02:15", timeZone: "Europe/Lisbon" },
    rules: [{ ruleKey: AUTOMATIC_RULE_KEY, ruleVersion: 1, unknown_key: "kept" }],
    proposals: [{
      runId: RUN_ID,
      candidateGroups: [{ baseSourceId: "doc-1", items: [{ sourceId: "bank-1" }] }],
    }],
    definition: { destinationSourceTypes: ["import_cgd_extrato_ordem"] },
    reason: "public",
  });
  assert.equal(input.schedule.diagnostic, "remove");
  assert.equal(input.proposals[0].candidate_groups[0].items[0].error_detail, "remove");
});

test("automation public result maps monthly proposal summaries and paged members", () => {
  const input = {
    proposals: [{
      grouping_key: "fdm-credit-card:2026-08",
      summary_snapshot: {
        calendar_month: "2026-08",
        source_count: 3,
        source_total: "7500.00",
        destination_count: 2,
        destination_total: "7500.00",
        total_count: 5,
        diagnostic: "remove",
      },
      error_detail: "remove",
    }],
    page: {
      members: [{
        source_type: "financial_documents",
        source_id: "document-1",
        amount_snapshot: "2500.00",
        internal_error: "remove",
      }],
      diagnostic: "remove",
    },
  };

  assert.deepEqual(toAutomationPublicResult(input), {
    proposals: [{
      groupingKey: "fdm-credit-card:2026-08",
      summarySnapshot: {
        calendarMonth: "2026-08",
        sourceCount: 3,
        sourceTotal: "7500.00",
        destinationCount: 2,
        destinationTotal: "7500.00",
        totalCount: 5,
      },
    }],
    page: {
      members: [{
        sourceType: "financial_documents",
        sourceId: "document-1",
        amountSnapshot: "2500.00",
      }],
    },
  });
});

test("automation public result preserves primitives and null without coercion", () => {
  for (const value of [null, undefined, "text", 0, false]) {
    assert.equal(toAutomationPublicResult(value), value);
  }
});

test("automation public result exhaustively maps the controller contract without leaking diagnostics", () => {
  const mappings = [
    ["rule_key", "ruleKey"],
    ["rule_version", "ruleVersion"],
    ["display_name", "displayName"],
    ["base_source_type", "baseSourceType"],
    ["destination_source_types", "destinationSourceTypes"],
    ["logic_description", "logicDescription"],
    ["allow_manual_execution", "allowManualExecution"],
    ["include_in_scheduled_batch", "includeInScheduledBatch"],
    ["difference_allowed", "differenceAllowed"],
    ["max_difference_days", "maxDifferenceDays"],
    ["time_of_day", "timeOfDay"],
    ["time_zone", "timeZone"],
    ["last_scheduled_run", "lastScheduledRun"],
    ["client_request_id", "clientRequestId"],
    ["scheduled_slot", "scheduledSlot"],
    ["definition_config_snapshot", "definitionConfigSnapshot"],
    ["analysis_completed_at", "analysisCompletedAt"],
    ["started_at", "startedAt"],
    ["finished_at", "finishedAt"],
    ["run_id", "runId"],
    ["base_source_id", "baseSourceId"],
    ["base_source_date", "baseSourceDate"],
    ["candidate_groups", "candidateGroups"],
    ["calculated_difference", "calculatedDifference"],
    ["allowed_difference", "allowedDifference"],
    ["reconciliation_id", "reconciliationId"],
    ["automatic_trigger", "automaticTrigger"],
    ["automatic_rule_key", "automaticRuleKey"],
    ["automatic_rule_version", "automaticRuleVersion"],
    ["automatic_run_id", "automaticRunId"],
    ["automatic_proposal_id", "automaticProposalId"],
    ["proposal_id", "proposalId"],
    ["source_type", "sourceType"],
    ["source_id", "sourceId"],
    ["source_date", "sourceDate"],
    ["amount_snapshot", "amountSnapshot"],
    ["row_snapshot", "rowSnapshot"],
    ["created_at", "createdAt"],
    ["updated_at", "updatedAt"],
  ];
  const strippedKeys = ["error_detail", "internal_error", "error_summary", "diagnostic", "stack"];
  const mappedFields = Object.fromEntries(mappings.map(([inputKey], index) => [inputKey, {
    nested_value: index,
    diagnostics: Object.fromEntries(strippedKeys.map((key) => [key, "remove"])),
  }]));
  const input = {
    payload: [mappedFields],
    unknown_public_field: { items: [{ still_unknown: true }] },
    diagnostics: Object.fromEntries(strippedKeys.map((key) => [key, "remove"])),
  };
  const before = JSON.parse(JSON.stringify(input));
  const result = toAutomationPublicResult(input);

  assert.deepEqual(input, before);
  assert.deepEqual(result.unknown_public_field, { items: [{ still_unknown: true }] });
  assert.deepEqual(result.diagnostics, {});
  for (const [inputKey, outputKey] of mappings) {
    assert.deepEqual(result.payload[0][outputKey], { nested_value: mappings.findIndex(([key]) => key === inputKey), diagnostics: {} });
    assert.equal(Object.hasOwn(result.payload[0], inputKey), false);
  }
  for (const value of [null, undefined, "text", 0, false]) {
    assert.equal(toAutomationPublicResult(value), value);
  }
});

test("automation RPC errors expose safe client statuses", () => {
  const cases = [
    ["Automatic rule is invalid.", 400],
    ["Automation proposal cannot be executed.", 400],
    ["Scheduled slot already exists.", 409],
    ["Stale proposal selected.", 409],
    ["Candidate limit exceeded.", 400],
    ["Ambiguous candidate selection.", 400],
  ];
  for (const [message, statusCode] of cases) {
    assert.equal(mapRpcError(new Error(message)).statusCode, statusCode, message);
  }
});

test("automation settings GET preserves complete pre- and post-migration RPC responses", async () => {
  const authorizations = [];
  const calls = [];
  for (const [label, ruleCount] of [["pre-migration", 2], ["post-migration", 4]]) {
    const response = responseRecorder();
    await withMockedHandler(SETTINGS_HANDLER_PATH, mockedSupabase({
      requireFeature: async (_request, area, feature) => {
        authorizations.push({ label, area, feature });
        return {
          user: { email: "admin@example.com", id: "admin-1" },
          access: { profile: { id: "admin-profile" } },
        };
      },
      restQuery: async (resource, options) => {
        calls.push({ label, resource, options });
        return productionSettingsRpcResult(ruleCount);
      },
    }), async (handler) => {
      await handler({ method: "GET" }, response);
    });

    assert.equal(response.statusCode, 200, label);
    assert.deepEqual(response.body, expectedPublicSettings(ruleCount), label);
  }

  assert.deepEqual(authorizations, [
    { label: "pre-migration", area: "settings", feature: "financial-reconciliation" },
    { label: "post-migration", area: "settings", feature: "financial-reconciliation" },
  ]);
  assert.deepEqual(calls, ["pre-migration", "post-migration"].map((label) => ({
    label,
    resource: "rpc/get_financial_reconciliation_automation_settings",
    options: { method: "POST", body: {} },
  })));
});

test("automation settings GET supplies the managed display name in its five-rule public response", async () => {
  const response = responseRecorder();
  await withMockedHandler(SETTINGS_HANDLER_PATH, mockedSupabase({
    restQuery: async () => ({
      ...productionSettingsRpcResult(4),
      rules: [...productionSettingsRules(), {
        rule_key: MONTHLY_INCOME_RULE_KEY,
        rule_version: 2,
        enabled: false,
        allow_manual_execution: false,
        include_in_scheduled_batch: false,
        difference_allowed: "7500.00",
        max_difference_days: 31,
        priority: 5,
      }],
    }),
  }), async (handler) => {
    await handler({ method: "GET" }, response);
  });

  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body.rules[4], {
    ruleKey: MONTHLY_INCOME_RULE_KEY,
    ruleVersion: 2,
    displayName: "Card Payments - POS - Income",
    enabled: false,
    allowManualExecution: false,
    includeInScheduledBatch: false,
    differenceAllowed: "7500.00",
    maxDifferenceDays: 31,
    priority: 5,
  });
});

test("automation settings has no action that creates an analysis run", async () => {
  let rpcCalled = false;
  const response = responseRecorder();
  await withMockedHandler(SETTINGS_HANDLER_PATH, mockedSupabase({
    restQuery: async () => {
      rpcCalled = true;
      return {};
    },
  }), async (handler) => {
    await handler({
      method: "POST",
      body: { action: "analyze_rule", ruleKeys: [CREDIT_CARD_RULE_KEY] },
    }, response);
  });

  assert.equal(response.statusCode, 405);
  assert.equal(response.headers.Allow, "GET, PUT");
  assert.equal(rpcCalled, false);
});

test("automation settings PUT normalizes the complete two-rule payload and preserves the complete RPC response", async () => {
  const calls = [];
  const response = responseRecorder();
  const input = managedSettings({ rules: [managedSettings().rules[0], creditCardRule] });
  await withMockedHandler(SETTINGS_HANDLER_PATH, mockedSupabase({
    requireFeature: async () => ({
      user: { email: " admin@example.com ", id: "admin-1" },
      access: { profile: { id: "admin-profile" } },
    }),
    restQuery: async (resource, options) => {
      calls.push({ resource, options });
      return productionSettingsRpcResult(2);
    },
  }), async (handler) => {
    await handler({ method: "PUT", body: input }, response);
  });

  assert.deepEqual(calls, [{
    resource: "rpc/replace_financial_reconciliation_automation_settings",
    options: {
      method: "POST",
      body: {
        p_schedule: { enabled: true, time_of_day: "02:15", time_zone: AUTOMATIC_TIME_ZONE },
        p_rules: [
          {
            rule_key: AUTOMATIC_RULE_KEY,
            rule_version: 2,
            enabled: true,
            allow_manual_execution: true,
            include_in_scheduled_batch: false,
            difference_allowed: "1.25",
            max_difference_days: 7,
            priority: 1,
          },
          {
            rule_key: CREDIT_CARD_RULE_KEY,
            rule_version: 1,
            enabled: false,
            allow_manual_execution: false,
            include_in_scheduled_batch: false,
            difference_allowed: "0.00",
            max_difference_days: 10,
            priority: 2,
          },
        ],
        p_actor: "admin@example.com",
      },
    },
  }]);
  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, expectedPublicSettings(2));
});

test("automation settings PUT supports pre-migration two-rule and post-migration four-rule catalogs", async () => {
  for (const [label, settings] of [
    ["pre-migration", managedSettings({ rules: [managedSettings().rules[0], creditCardRule] })],
    ["post-migration", fourRuleSettings()],
  ]) {
    const calls = [];
    const response = responseRecorder();
    await withMockedHandler(SETTINGS_HANDLER_PATH, mockedSupabase({
      restQuery: async (resource, options) => {
        calls.push({ resource, options });
        return productionSettingsRpcResult(settings.rules.length);
      },
    }), async (handler) => {
      await handler({ method: "PUT", body: settings }, response);
    });

    assert.equal(response.statusCode, 200, label);
    assert.equal(calls.length, 1, label);
    assert.equal(calls[0].resource, "rpc/replace_financial_reconciliation_automation_settings", label);
    assert.deepEqual(
      calls[0].options.body.p_rules.map(({ rule_key, rule_version }) => [rule_key, rule_version]),
      settings.rules.map(({ ruleKey, ruleVersion }) => [ruleKey, ruleVersion]),
      label,
    );
    assert.deepEqual(response.body, expectedPublicSettings(settings.rules.length), label);
  }
});

test("automation settings PUT rejects nonzero amount-only tolerance before replacement RPC", async () => {
  let rpcCalled = false;
  const response = responseRecorder();
  await withMockedHandler(SETTINGS_HANDLER_PATH, mockedSupabase({
    restQuery: async () => {
      rpcCalled = true;
      return {};
    },
  }), async (handler) => {
    await handler({
      method: "PUT",
      body: fourRuleSettings({ bankStatementAmountOnlyDifferenceAllowed: "0.01" }),
    }, response);
  });

  assert.equal(response.statusCode, 400);
  assert.equal(rpcCalled, false);
  assert.match(response.body.error, /amount-only.*zero/i);
});

test("manual automation GET authorizes app access and validates the run detail RPC", async () => {
  const authorizations = [];
  const calls = [];
  const response = responseRecorder();
  await withMockedHandler(MANUAL_HANDLER_PATH, mockedSupabase({
    requireFeature: async (_request, area, feature) => {
      authorizations.push({ area, feature });
      return {
        user: { email: "user@example.com", id: "user-1" },
        access: { profile: { id: "profile-1" } },
      };
    },
    restQuery: async (resource, options) => {
      calls.push({ resource, options });
      return { run_id: RUN_ID, proposals: [{ source_id: "bank-1", internal_error: "hidden" }] };
    },
  }), async (handler) => {
    await handler({ method: "GET", query: { run_id: RUN_ID } }, response);
  });

  assert.deepEqual(authorizations, [{ area: "app", feature: "financial-reconciliation" }]);
  assert.deepEqual(calls, [{
    resource: "rpc/get_financial_reconciliation_automatic_run",
    options: { method: "POST", body: { p_run_id: RUN_ID } },
  }]);
  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, { runId: RUN_ID, proposals: [{ sourceId: "bank-1" }] });

  const invalidResponse = responseRecorder();
  let invalidRpcCalled = false;
  await withMockedHandler(MANUAL_HANDLER_PATH, mockedSupabase({
    restQuery: async () => {
      invalidRpcCalled = true;
      return {};
    },
  }), async (handler) => {
    await handler({ method: "GET", query: { run_id: "not-a-uuid" } }, invalidResponse);
  });
  assert.equal(invalidResponse.statusCode, 400);
  assert.equal(invalidRpcCalled, false);
});

test("monthly member paging binds the authenticated app actor to the only data RPC", async () => {
  const authorizations = [];
  const calls = [];
  const response = responseRecorder();
  await withMockedHandler(MEMBERS_HANDLER_PATH, mockedSupabase({
    requireFeature: async (_request, area, feature) => {
      authorizations.push({ area, feature });
      return {
        user: { email: " owner@example.com ", id: "user-1" },
        access: { profile: { id: "profile-1" } },
      };
    },
    restQuery: async (resource, options) => {
      calls.push({ resource, options });
      return {
        run_id: RUN_ID,
        proposal_id: PROPOSAL_ID,
        role: "source",
        offset: 50,
        limit: 50,
        total_count: 73,
        members: [{
          role: "source",
          source_type: "import_cgd_extrato_ordem",
          source_id: "00000000-0000-0000-0000-000000000073",
          ordinal: 51,
          source_date: "2026-08-10",
          amount: "125.00",
          description: "POS VENDAS",
          account: "",
          row_snapshot: { row_key: "bank-73", diagnostic: "remove" },
          internal_error: "remove",
        }],
        diagnostic: "remove",
      };
    },
  }), async (handler) => {
    await handler({
      method: "GET",
      query: {
        run_id: RUN_ID.toUpperCase(),
        proposal_id: PROPOSAL_ID.toUpperCase(),
        role: "source",
        offset: "50",
        limit: "50",
      },
    }, response);
  });

  assert.deepEqual(authorizations, [{ area: "app", feature: "financial-reconciliation" }]);
  assert.deepEqual(calls, [{
    resource: "rpc/get_financial_reconciliation_automatic_proposal_members",
    options: {
      method: "POST",
      body: {
        p_run_id: RUN_ID,
        p_proposal_id: PROPOSAL_ID,
        p_role: "source",
        p_offset: 50,
        p_limit: 50,
        p_actor: "owner@example.com",
      },
    },
  }]);
  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, {
    runId: RUN_ID,
    proposalId: PROPOSAL_ID,
    role: "source",
    offset: 50,
    limit: 50,
    totalCount: 73,
    members: [{
      role: "source",
      sourceType: "import_cgd_extrato_ordem",
      sourceId: "00000000-0000-0000-0000-000000000073",
      ordinal: 51,
      sourceDate: "2026-08-10",
      amount: "125.00",
      description: "POS VENDAS",
      account: "",
      rowSnapshot: { row_key: "bank-73" },
    }],
  });
});

test("monthly member paging rejects malformed query values and absent actor before its RPC", async () => {
  const invalidQueries = [
    { run_id: "not-a-uuid", proposal_id: PROPOSAL_ID, role: "source", offset: "0", limit: "50" },
    { run_id: RUN_ID, proposal_id: "not-a-uuid", role: "source", offset: "0", limit: "50" },
    { run_id: RUN_ID, proposal_id: PROPOSAL_ID, role: "Source", offset: "0", limit: "50" },
    { run_id: RUN_ID, proposal_id: PROPOSAL_ID, role: "source ", offset: "0", limit: "50" },
    { run_id: RUN_ID, proposal_id: PROPOSAL_ID, role: "source", offset: "-1", limit: "50" },
    { run_id: RUN_ID, proposal_id: PROPOSAL_ID, role: "source", offset: "1.5", limit: "50" },
    { run_id: RUN_ID, proposal_id: PROPOSAL_ID, role: "source", offset: "0", limit: "0" },
    { run_id: RUN_ID, proposal_id: PROPOSAL_ID, role: "source", offset: "0", limit: "51" },
    { run_id: RUN_ID, proposal_id: PROPOSAL_ID, role: "source", offset: "0", limit: ["50"] },
  ];

  for (const query of invalidQueries) {
    let rpcCalled = false;
    const response = responseRecorder();
    await withMockedHandler(MEMBERS_HANDLER_PATH, mockedSupabase({
      restQuery: async () => {
        rpcCalled = true;
        return {};
      },
    }), async (handler) => handler({ method: "GET", query }, response));
    assert.equal(response.statusCode, 400, JSON.stringify(query));
    assert.equal(rpcCalled, false, JSON.stringify(query));
  }

  let rpcCalled = false;
  const noActorResponse = responseRecorder();
  await withMockedHandler(MEMBERS_HANDLER_PATH, mockedSupabase({
    requireFeature: async () => ({
      user: { email: "", id: "" },
      access: { profile: { id: "profile-1" } },
    }),
    restQuery: async () => {
      rpcCalled = true;
      return {};
    },
  }), async (handler) => handler({
    method: "GET",
    query: { run_id: RUN_ID, proposal_id: PROPOSAL_ID, role: "source", offset: "0", limit: "50" },
  }, noActorResponse));
  assert.equal(noActorResponse.statusCode, 403);
  assert.equal(rpcCalled, false);
});

test("monthly member paging is GET-only, denies fallback access, and sanitizes database errors", async () => {
  for (const [request, expectedStatus, expectedAllow] of [
    [{ method: "POST", query: {} }, 405, "GET"],
    [{ method: "DELETE", query: {} }, 405, "GET"],
  ]) {
    let rpcCalled = false;
    const response = responseRecorder();
    await withMockedHandler(MEMBERS_HANDLER_PATH, mockedSupabase({
      restQuery: async () => {
        rpcCalled = true;
        return {};
      },
    }), async (handler) => handler(request, response));
    assert.equal(response.statusCode, expectedStatus);
    assert.equal(response.headers.Allow, expectedAllow);
    assert.equal(rpcCalled, false);
  }

  let deniedRpcCalled = false;
  const deniedResponse = responseRecorder();
  await withMockedHandler(MEMBERS_HANDLER_PATH, mockedSupabase({
    requireFeature: async () => ({
      user: { email: "user@example.com", id: "user-1" },
      access: { profile: { id: "", name: "Full access (fallback)" } },
    }),
    restQuery: async () => {
      deniedRpcCalled = true;
      return {};
    },
  }), async (handler) => handler({
    method: "GET",
    query: { run_id: RUN_ID, proposal_id: PROPOSAL_ID, role: "source", offset: "0", limit: "50" },
  }, deniedResponse));
  assert.equal(deniedResponse.statusCode, 403);
  assert.equal(deniedRpcCalled, false);

  const errorResponse = responseRecorder();
  await withMockedHandler(MEMBERS_HANDLER_PATH, mockedSupabase({
    restQuery: async () => {
      const error = new Error("Automation proposal was not found. relation task5_private_4f92 does not exist");
      error.statusCode = 400;
      error.supabasePayload = { details: "relation task5_private_4f92 does not exist" };
      throw error;
    },
  }), async (handler) => handler({
    method: "GET",
    query: { run_id: RUN_ID, proposal_id: PROPOSAL_ID, role: "destination", offset: "0", limit: "50" },
  }, errorResponse));
  assert.equal(errorResponse.statusCode, 400);
  assert.deepEqual(errorResponse.body, { error: "Reconciliation automation request could not be completed." });
  assert.doesNotMatch(JSON.stringify(errorResponse.body), /task5_private_4f92/);
});

test("manual automation GET exposes only enabled manual rules from the workbench catalog", async () => {
  const authorizations = [];
  const calls = [];
  const response = responseRecorder();
  await withMockedHandler(MANUAL_HANDLER_PATH, mockedSupabase({
    requireFeature: async (_request, area, feature) => {
      authorizations.push({ area, feature });
      return {
        user: { email: "user@example.com", id: "user-1" },
        access: { profile: { id: "profile-1" } },
      };
    },
    restQuery: async (resource, options) => {
      calls.push({ resource, options });
      return {
        rules: [
          {
            rule_key: AUTOMATIC_RULE_KEY,
            rule_version: 2,
            display_name: "Financial Documents to CGD Bank Statement",
            enabled: true,
            allow_manual_execution: true,
            difference_allowed: "1.00",
            max_difference_days: 7,
            diagnostic: "hidden",
          },
          {
            rule_key: CREDIT_CARD_RULE_KEY,
            rule_version: 1,
            display_name: "Financial Documents to CGD Credit Card",
            enabled: false,
            allow_manual_execution: true,
            difference_allowed: "0.00",
            max_difference_days: 10,
          },
          {
            rule_key: BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY,
            rule_version: 1,
            display_name: "Financial Documents to CGD Bank Statement (Amount Only)",
            enabled: true,
            allow_manual_execution: false,
            difference_allowed: "0.00",
            max_difference_days: 1,
          },
          {
            rule_key: CREDIT_CARD_AMOUNT_ONLY_RULE_KEY,
            rule_version: 1,
            display_name: "Financial Documents to CGD Credit Card (Amount Only)",
            enabled: true,
            allow_manual_execution: true,
            difference_allowed: "0.00",
            max_difference_days: 1,
          },
        ],
      };
    },
  }), async (handler) => {
    await handler({ method: "GET", query: { view: "rules" } }, response);
  });

  assert.deepEqual(authorizations, [{ area: "app", feature: "financial-reconciliation" }]);
  assert.deepEqual(calls, [{
    resource: "rpc/get_financial_reconciliation_automatic_manual_rules",
    options: { method: "POST", body: {} },
  }]);
  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, {
    rules: [
      {
        ruleKey: AUTOMATIC_RULE_KEY,
        ruleVersion: 2,
        displayName: "Financial Documents to CGD Bank Statement",
        enabled: true,
        allowManualExecution: true,
        differenceAllowed: "1.00",
        maxDifferenceDays: 7,
      },
      {
        ruleKey: CREDIT_CARD_AMOUNT_ONLY_RULE_KEY,
        ruleVersion: 1,
        displayName: "Financial Documents to CGD Credit Card (Amount Only)",
        enabled: true,
        allowManualExecution: true,
        differenceAllowed: "0.00",
        maxDifferenceDays: 1,
      },
    ],
  });
  assert.equal(Object.hasOwn(response.body, "schedule"), false);
});

test("manual active-run lookup and continuation bind the authenticated actor", async () => {
  const calls = [];
  const auth = {
    user: { email: "admin@example.com", id: "admin-1" },
    access: { profile: { id: "profile-1" } },
  };
  const supabase = mockedSupabase({
    requireFeature: async () => auth,
    restQuery: async (resource, options) => {
      calls.push({ resource, options });
      return { run_id: RUN_ID, status: "analyzing", analysis_processed: 25, analysis_total: 100 };
    },
  });

  const activeResponse = responseRecorder();
  await withMockedHandler(MANUAL_HANDLER_PATH, supabase, async (handler) => {
    await handler({ method: "GET", query: { view: "active_run" } }, activeResponse);
  });
  const continueResponse = responseRecorder();
  await withMockedHandler(MANUAL_HANDLER_PATH, supabase, async (handler) => {
    await handler({ method: "POST", body: { action: "continue_analysis", runId: RUN_ID } }, continueResponse);
  });

  assert.deepEqual(calls, [
    {
      resource: "rpc/get_financial_reconciliation_automatic_active_run",
      options: { method: "POST", body: { p_actor: "admin@example.com" } },
    },
    {
      resource: "rpc/continue_financial_reconciliation_automatic_analysis",
      options: { method: "POST", body: { p_run_id: RUN_ID, p_actor: "admin@example.com" } },
    },
  ]);
  for (const response of [activeResponse, continueResponse]) {
    assert.equal(response.statusCode, 200);
    assert.deepEqual(response.body, {
      runId: RUN_ID,
      status: "analyzing",
      analysisProcessed: 25,
      analysisTotal: 100,
    });
  }

  let invalidRpcCalled = false;
  const invalidResponse = responseRecorder();
  await withMockedHandler(MANUAL_HANDLER_PATH, mockedSupabase({
    restQuery: async () => { invalidRpcCalled = true; return {}; },
  }), async (handler) => {
    await handler({
      method: "POST",
      body: { action: "continue_analysis", runId: RUN_ID, pageSize: 50 },
    }, invalidResponse);
  });
  assert.equal(invalidResponse.statusCode, 400);
  assert.equal(invalidRpcCalled, false);
});

test("analyze_rule sends exactly one selected rule for each amount-only key/version", async () => {
  for (const [ruleKey, ruleVersion] of [
    [BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY, BANK_STATEMENT_AMOUNT_ONLY_RULE_VERSION],
    [CREDIT_CARD_AMOUNT_ONLY_RULE_KEY, CREDIT_CARD_AMOUNT_ONLY_RULE_VERSION],
  ]) {
    const calls = [];
    const response = responseRecorder();
    await withMockedHandler(MANUAL_HANDLER_PATH, mockedSupabase({
      restQuery: async (resource, options) => {
        calls.push({ resource, options });
        return { run_id: RUN_ID, status: "ready", definitions: [{ rule_key: ruleKey, rule_version: ruleVersion }] };
      },
    }), async (handler) => {
      await handler({
        method: "POST",
        body: { action: "analyze_rule", ruleKeys: [ruleKey], clientRequestId: REQUEST_ID },
      }, response);
    });

    assert.equal(response.statusCode, 200, ruleKey);
    assert.deepEqual(calls, [{
      resource: "rpc/create_financial_reconciliation_automatic_analysis",
      options: {
        method: "POST",
        body: {
          p_rule_keys: [ruleKey],
          p_mode: "manual_rule",
          p_actor: "user@example.com",
          p_client_request_id: REQUEST_ID,
        },
      },
    }], ruleKey);
    assert.deepEqual(response.body.definitions, [{ ruleKey, ruleVersion }], ruleKey);
  }
});

test("analyze_rule rejects an unknown fifth rule before analysis RPC", async () => {
  let rpcCalled = false;
  const response = responseRecorder();
  await withMockedHandler(MANUAL_HANDLER_PATH, mockedSupabase({
    restQuery: async () => {
      rpcCalled = true;
      return {};
    },
  }), async (handler) => {
    await handler({
      method: "POST",
      body: {
        action: "analyze_rule",
        ruleKeys: ["financial_documents_cgd_cash_amount_only"],
        clientRequestId: REQUEST_ID,
      },
    }, response);
  });

  assert.equal(response.statusCode, 400);
  assert.equal(rpcCalled, false);
  assert.match(response.body.error, /rule key/i);
});

test("analyze_batch returns 400 before any RPC", async () => {
  let authorizationCalled = false;
  let rpcCalled = false;
  const response = responseRecorder();
  await withMockedHandler(MANUAL_HANDLER_PATH, mockedSupabase({
    requireFeature: async () => {
      authorizationCalled = true;
      return {};
    },
    restQuery: async () => {
      rpcCalled = true;
      return {};
    },
  }), async (handler) => {
    await handler({ method: "POST", body: { action: "analyze_batch", clientRequestId: REQUEST_ID } }, response);
  });

  assert.equal(response.statusCode, 400);
  assert.equal(authorizationCalled, false);
  assert.equal(rpcCalled, false);
  assert.match(response.body.error, /automation action/i);
});

test("execute_selected runs proposal RPCs sequentially, retains partial failures, and finalizes after the loop", async () => {
  const authorizations = [];
  const calls = [];
  let runReads = 0;
  let activeExecutions = 0;
  let maximumActiveExecutions = 0;
  const response = responseRecorder();
  await withMockedHandler(MANUAL_HANDLER_PATH, mockedSupabase({
    requireFeature: async (_request, area, feature) => {
      authorizations.push({ area, feature });
      return {
        user: { email: "user@example.com", id: "user-1" },
        access: { profile: { id: "profile-1" } },
      };
    },
    restQuery: async (resource, options) => {
      calls.push({ resource, options });
      if (resource === "rpc/get_financial_reconciliation_automatic_run") {
        runReads += 1;
        if (runReads > 1) {
          return {
            runId: RUN_ID,
            trigger: "manual",
            actor: "user@example.com",
            finishedAt: null,
            proposals: [
              { id: PROPOSAL_ID, status: "completed" },
              { id: PROPOSAL_ID_2, status: "failed", reason: "execution_failed" },
              { id: PROPOSAL_ID_3, status: "stale", reason: "source_snapshot_changed" },
            ],
          };
        }
        return {
          runId: RUN_ID,
          trigger: "manual",
          actor: "user@example.com",
          finishedAt: null,
          proposals: [
            { id: PROPOSAL_ID, status: "proposed" },
            { id: PROPOSAL_ID_2, status: "proposed" },
            { id: PROPOSAL_ID_3, status: "proposed" },
          ],
        };
      }
      if (resource === "rpc/execute_financial_reconciliation_automatic_proposal") {
        activeExecutions += 1;
        maximumActiveExecutions = Math.max(maximumActiveExecutions, activeExecutions);
        await new Promise((resolve) => setImmediate(resolve));
        activeExecutions -= 1;
        if (options.body.p_proposal_id === PROPOSAL_ID_2) throw new Error("secret database diagnostic");
        if (options.body.p_proposal_id === PROPOSAL_ID_3) {
          return { proposalId: PROPOSAL_ID_3, runId: RUN_ID, status: "stale", reason: "source_snapshot_changed" };
        }
        return { proposalId: PROPOSAL_ID, runId: RUN_ID, status: "completed" };
      }
      return { run_id: RUN_ID, status: "partial", diagnostic: "hidden" };
    },
  }), async (handler) => {
    await handler({
      method: "POST",
      body: {
        action: "execute_selected",
        runId: RUN_ID,
        proposalIds: [PROPOSAL_ID, PROPOSAL_ID_2, PROPOSAL_ID_3],
      },
    }, response);
  });

  assert.deepEqual(authorizations, [{ area: "app", feature: "financial-reconciliation" }]);
  assert.equal(maximumActiveExecutions, 1);
  assert.deepEqual(calls, [
    {
      resource: "rpc/get_financial_reconciliation_automatic_run",
      options: { method: "POST", body: { p_run_id: RUN_ID } },
    },
    {
      resource: "rpc/execute_financial_reconciliation_automatic_proposal",
      options: { method: "POST", body: { p_proposal_id: PROPOSAL_ID, p_actor: "user@example.com" } },
    },
    {
      resource: "rpc/execute_financial_reconciliation_automatic_proposal",
      options: { method: "POST", body: { p_proposal_id: PROPOSAL_ID_2, p_actor: "user@example.com" } },
    },
    {
      resource: "rpc/execute_financial_reconciliation_automatic_proposal",
      options: { method: "POST", body: { p_proposal_id: PROPOSAL_ID_3, p_actor: "user@example.com" } },
    },
    {
      resource: "rpc/get_financial_reconciliation_automatic_run",
      options: { method: "POST", body: { p_run_id: RUN_ID } },
    },
    {
      resource: "rpc/finish_financial_reconciliation_automatic_run",
      options: { method: "POST", body: { p_run_id: RUN_ID } },
    },
  ]);
  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, {
    run: { runId: RUN_ID, status: "partial" },
    outcomes: [
      { proposalId: PROPOSAL_ID, runId: RUN_ID, status: "completed" },
      { proposalId: PROPOSAL_ID_2, status: "failed", reason: "execution_failed" },
      { proposalId: PROPOSAL_ID_3, runId: RUN_ID, status: "stale", reason: "source_snapshot_changed" },
    ],
  });
  assert.doesNotMatch(JSON.stringify(response.body), /secret database diagnostic/);
});

test("execute_selected with zero proposals finalizes the ready manual run without executing proposals", async () => {
  const calls = [];
  const response = responseRecorder();
  const readyRun = {
    runId: RUN_ID,
    trigger: "manual",
    actor: "user@example.com",
    status: "ready",
    analysisCompletedAt: "2026-08-23T12:00:00.000Z",
    finishedAt: null,
    proposals: [{ id: PROPOSAL_ID, status: "proposed" }],
  };
  await withMockedHandler(MANUAL_HANDLER_PATH, mockedSupabase({
    restQuery: async (resource, options) => {
      calls.push({ resource, options });
      if (resource === "rpc/get_financial_reconciliation_automatic_run") return readyRun;
      if (resource === "rpc/finish_financial_reconciliation_automatic_run") {
        return {
          ...readyRun,
          status: "completed",
          finishedAt: "2026-08-23T12:01:00.000Z",
          proposals: [{ id: PROPOSAL_ID, status: "deselected", reason: "not_selected" }],
        };
      }
      throw new Error(`Unexpected RPC ${resource}`);
    },
  }), async (handler) => {
    await handler({
      method: "POST",
      body: { action: "execute_selected", runId: RUN_ID, proposalIds: [] },
    }, response);
  });

  assert.deepEqual(calls, [
    {
      resource: "rpc/get_financial_reconciliation_automatic_run",
      options: { method: "POST", body: { p_run_id: RUN_ID } },
    },
    {
      resource: "rpc/finish_financial_reconciliation_automatic_run",
      options: { method: "POST", body: { p_run_id: RUN_ID } },
    },
  ]);
  assert.equal(response.statusCode, 200);
  assert.equal(response.body.run.status, "completed");
  assert.equal(response.body.run.proposals[0].status, "deselected");
  assert.deepEqual(response.body.outcomes, []);
});

test("execute_selected with zero proposals cannot finish an analysis still in progress", async () => {
  const calls = [];
  const response = responseRecorder();
  await withMockedHandler(MANUAL_HANDLER_PATH, mockedSupabase({
    restQuery: async (resource, options) => {
      calls.push({ resource, options });
      return {
        runId: RUN_ID,
        trigger: "manual",
        actor: "user@example.com",
        status: "analyzing",
        analysisCompletedAt: null,
        finishedAt: null,
        proposals: [],
      };
    },
  }), async (handler) => {
    await handler({
      method: "POST",
      body: { action: "execute_selected", runId: RUN_ID, proposalIds: [] },
    }, response);
  });

  assert.equal(response.statusCode, 400);
  assert.match(response.body.error, /analysis.*ready/i);
  assert.deepEqual(calls, [{
    resource: "rpc/get_financial_reconciliation_automatic_run",
    options: { method: "POST", body: { p_run_id: RUN_ID } },
  }]);
});

test("execute_selected leaves a transport-uncertain selected proposal resumable", async () => {
  const calls = [];
  let runReads = 0;
  const response = responseRecorder();
  const unresolvedRun = {
    runId: RUN_ID,
    trigger: "manual",
    actor: "user@example.com",
    status: "ready",
    finishedAt: null,
    proposals: [{ id: PROPOSAL_ID, status: "proposed" }],
  };
  await withMockedHandler(MANUAL_HANDLER_PATH, mockedSupabase({
    restQuery: async (resource, options) => {
      calls.push({ resource, options });
      if (resource === "rpc/get_financial_reconciliation_automatic_run") {
        runReads += 1;
        return unresolvedRun;
      }
      if (resource === "rpc/execute_financial_reconciliation_automatic_proposal") {
        throw new Error("transport outcome unknown");
      }
      throw new Error(`Unexpected RPC ${resource}`);
    },
  }), async (handler) => {
    await handler({
      method: "POST",
      body: { action: "execute_selected", runId: RUN_ID, proposalIds: [PROPOSAL_ID] },
    }, response);
  });

  assert.equal(runReads, 2, "authoritative state is reloaded after the uncertain attempt");
  assert.equal(calls.some(({ resource }) => resource === "rpc/finish_financial_reconciliation_automatic_run"), false);
  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, {
    run: unresolvedRun,
    outcomes: [{ proposalId: PROPOSAL_ID, status: "failed", reason: "execution_failed" }],
  });
});

test("execute_selected rejects mixed, scheduled, finished, or foreign-actor runs before mutation", async () => {
  const cases = [
    ["mixed selection", {
      runId: RUN_ID,
      trigger: "manual",
      actor: "user@example.com",
      finishedAt: null,
      proposals: [{ id: PROPOSAL_ID_2, status: "proposed" }],
    }, 400],
    ["scheduled run", {
      runId: RUN_ID,
      trigger: "scheduled",
      actor: "user@example.com",
      finishedAt: null,
      proposals: [{ id: PROPOSAL_ID, status: "proposed" }],
    }, 400],
    ["foreign actor", {
      runId: RUN_ID,
      trigger: "manual",
      actor: "other@example.com",
      finishedAt: null,
      proposals: [{ id: PROPOSAL_ID, status: "proposed" }],
    }, 403],
  ];

  for (const [name, run, expectedStatus] of cases) {
    const calls = [];
    const response = responseRecorder();
    await withMockedHandler(MANUAL_HANDLER_PATH, mockedSupabase({
      restQuery: async (resource, options) => {
        calls.push({ resource, options });
        return run;
      },
    }), async (handler) => {
      await handler({
        method: "POST",
        body: { action: "execute_selected", runId: RUN_ID, proposalIds: [PROPOSAL_ID] },
      }, response);
    });
    assert.equal(response.statusCode, expectedStatus, name);
    assert.deepEqual(calls, [{
      resource: "rpc/get_financial_reconciliation_automatic_run",
      options: { method: "POST", body: { p_run_id: RUN_ID } },
    }], name);
  }
});

test("execute_selected returns authoritative persisted outcomes when a finished manual run is retried", async () => {
  const calls = [];
  const response = responseRecorder();
  const finishedRun = {
    runId: RUN_ID,
    trigger: "manual",
    actor: "user@example.com",
    status: "partial",
    finishedAt: "2026-08-14T10:00:00Z",
    proposals: [
      { id: PROPOSAL_ID, status: "completed", reconciliationId: "00000000-0000-0000-0000-000000000006" },
      { id: PROPOSAL_ID_2, status: "stale", reason: "source_snapshot_changed" },
      { id: PROPOSAL_ID_3, status: "failed", reason: "execution_failed" },
    ],
  };
  await withMockedHandler(MANUAL_HANDLER_PATH, mockedSupabase({
    restQuery: async (resource, options) => {
      calls.push({ resource, options });
      return finishedRun;
    },
  }), async (handler) => {
    await handler({
      method: "POST",
      body: {
        action: "execute_selected",
        runId: RUN_ID,
        proposalIds: [PROPOSAL_ID, PROPOSAL_ID_2, PROPOSAL_ID_3],
      },
    }, response);
  });

  assert.deepEqual(calls, [{
    resource: "rpc/get_financial_reconciliation_automatic_run",
    options: { method: "POST", body: { p_run_id: RUN_ID } },
  }]);
  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, {
    run: finishedRun,
    outcomes: [
      {
        proposalId: PROPOSAL_ID,
        runId: RUN_ID,
        status: "completed",
        reconciliationId: "00000000-0000-0000-0000-000000000006",
      },
      { proposalId: PROPOSAL_ID_2, runId: RUN_ID, status: "stale", reason: "source_snapshot_changed" },
      { proposalId: PROPOSAL_ID_3, runId: RUN_ID, status: "failed", reason: "execution_failed" },
    ],
  });
});

test("automation endpoints reject fallback access before any RPC", async () => {
  for (const [handlerPath, request] of [
    [SETTINGS_HANDLER_PATH, { method: "GET" }],
    [MANUAL_HANDLER_PATH, { method: "GET", query: { run_id: RUN_ID } }],
    [MANUAL_HANDLER_PATH, {
      method: "POST",
      body: { action: "analyze_rule", ruleKeys: [AUTOMATIC_RULE_KEY], clientRequestId: REQUEST_ID },
    }],
    [MANUAL_HANDLER_PATH, {
      method: "POST",
      body: { action: "execute_selected", runId: RUN_ID, proposalIds: [PROPOSAL_ID] },
    }],
    [MANUAL_HANDLER_PATH, { method: "DELETE" }],
  ]) {
    let rpcCalled = false;
    const response = responseRecorder();
    await withMockedHandler(handlerPath, mockedSupabase({
      requireFeature: async () => ({
        user: { email: "user@example.com", id: "user-1" },
        access: { profile: { id: "", name: "Full access (fallback)" } },
      }),
      restQuery: async () => {
        rpcCalled = true;
        return {};
      },
    }), async (handler) => {
      await handler(request, response);
    });
    assert.equal(response.statusCode, 403);
    assert.equal(rpcCalled, false);
  }
});

test("manual automation canonicalizes valid UUID spellings and rejects case-variant duplicates", async () => {
  const calls = [];
  const response = responseRecorder();
  await withMockedHandler(MANUAL_HANDLER_PATH, mockedSupabase({
    restQuery: async (resource, options) => {
      calls.push({ resource, options });
      return { runId: RUN_ID };
    },
  }), async (handler) => {
    await handler({ method: "GET", query: { run_id: CASE_UUID.toUpperCase() } }, response);
  });
  assert.equal(response.statusCode, 200);
  assert.deepEqual(calls, [{
    resource: "rpc/get_financial_reconciliation_automatic_run",
    options: { method: "POST", body: { p_run_id: CASE_UUID } },
  }]);

  const duplicateResponse = responseRecorder();
  let duplicateRpcCalled = false;
  await withMockedHandler(MANUAL_HANDLER_PATH, mockedSupabase({
    restQuery: async () => {
      duplicateRpcCalled = true;
      return {};
    },
  }), async (handler) => {
    await handler({
      method: "POST",
      body: {
        action: "execute_selected",
        runId: RUN_ID,
        proposalIds: [CASE_UUID, CASE_UUID.toUpperCase()],
      },
    }, duplicateResponse);
  });
  assert.equal(duplicateResponse.statusCode, 400);
  assert.equal(duplicateRpcCalled, false);
  assert.match(duplicateResponse.body.error, /up to 100 unique proposal IDs/i);
});

test("automation handlers set Allow, reject invalid payloads before RPCs, and safely map RPC errors", async () => {
  const settingsMethodResponse = responseRecorder();
  await withMockedHandler(SETTINGS_HANDLER_PATH, mockedSupabase(), async (handler) => {
    await handler({ method: "POST" }, settingsMethodResponse);
  });
  assert.equal(settingsMethodResponse.statusCode, 405);
  assert.equal(settingsMethodResponse.headers.Allow, "GET, PUT");

  const manualMethodResponse = responseRecorder();
  await withMockedHandler(MANUAL_HANDLER_PATH, mockedSupabase(), async (handler) => {
    await handler({ method: "DELETE" }, manualMethodResponse);
  });
  assert.equal(manualMethodResponse.statusCode, 405);
  assert.equal(manualMethodResponse.headers.Allow, "GET, POST");

  let invalidRpcCalled = false;
  const invalidSettingsResponse = responseRecorder();
  await withMockedHandler(SETTINGS_HANDLER_PATH, mockedSupabase({
    restQuery: async () => {
      invalidRpcCalled = true;
      return {};
    },
  }), async (handler) => {
    await handler({ method: "PUT", body: managedSettings({ unexpected: true }) }, invalidSettingsResponse);
  });
  assert.equal(invalidSettingsResponse.statusCode, 400);
  assert.equal(invalidRpcCalled, false);

  const mappedErrorResponse = responseRecorder();
  await withMockedHandler(MANUAL_HANDLER_PATH, mockedSupabase({
    restQuery: async () => {
      const error = new Error("Stale proposal selected. relation internal_secret does not exist");
      error.statusCode = 400;
      error.supabasePayload = { details: "relation internal_secret does not exist" };
      throw error;
    },
  }), async (handler) => {
    await handler({ method: "GET", query: { run_id: RUN_ID } }, mappedErrorResponse);
  });
  assert.equal(mappedErrorResponse.statusCode, 409);
  assert.deepEqual(mappedErrorResponse.body, { error: "The reconciliation automation state changed. Refresh and try again." });
  assert.doesNotMatch(JSON.stringify(mappedErrorResponse.body), /internal_secret/);
});

test("automation schema migration pins the managed catalog and execution provenance contract", () => {
  assert.equal(fs.existsSync(SCHEMA_MIGRATION_PATH), true, "automation schema migration must exist");
  const schemaMigration = fs.readFileSync(SCHEMA_MIGRATION_PATH, "utf8");

  for (const table of [
    "financial_reconciliation_automatic_rule_definitions",
    "financial_reconciliation_automatic_rule_configs",
    "financial_reconciliation_automatic_schedule",
    "financial_reconciliation_automatic_runs",
    "financial_reconciliation_automatic_proposals",
  ]) {
    assert.match(schemaMigration, new RegExp(`create table if not exists public\\.${table}`));
    assert.match(schemaMigration, new RegExp(`alter table public\\.${table} enable row level security;`));
  }

  assert.match(schemaMigration, /create extension if not exists pgcrypto;/);
  assert.match(schemaMigration, /create extension if not exists unaccent;/);
  assert.match(schemaMigration, /create extension if not exists pg_trgm;/);
  assert.match(schemaMigration, /'financial_documents_cgd_bank_statement',\s*1/);
  assert.match(schemaMigration, /"destinationSourceTypes": \["import_cgd_extrato_ordem"\]/);
  assert.match(schemaMigration, /"identityBranches"/);
  assert.match(schemaMigration, /"document_number"/);
  assert.match(schemaMigration, /"description_similarity"/);
  assert.match(schemaMigration, /"supplier_similarity"/);
  assert.match(schemaMigration, /"documentNumberMinimumCompactLength": 4/);
  assert.match(schemaMigration, /"descriptionSimilarityThreshold": 0\.60/);
  assert.match(schemaMigration, /"supplierWordSimilarityThreshold": 0\.70/);
  assert.match(schemaMigration, /"maxDestinationRecords": 4/);
  assert.match(schemaMigration, /"maxIdentityCandidatesPerBase": 12/);
  assert.match(schemaMigration, /enabled boolean not null default false/);
  assert.match(schemaMigration, /'financial_documents_cgd_bank_statement', 1, false, false, false, 0\.00, 7, 1/);
  assert.match(schemaMigration, /on conflict \(rule_key\) do nothing;/);
  assert.match(schemaMigration, /on conflict \(id\) do nothing;/);

  assert.match(schemaMigration, /id boolean primary key default true check \(id\)/);
  assert.match(schemaMigration, /time_of_day time without time zone not null default '02:00'/);
  assert.match(schemaMigration, /time_zone text not null default 'Europe\/Lisbon' check \(time_zone = 'Europe\/Lisbon'\)/);
  assert.match(schemaMigration, /unique \(actor, client_request_id\)/);
  assert.match(schemaMigration, /create unique index if not exists financial_reconciliation_automatic_runs_scheduled_slot_uidx[\s\S]*where scheduled_slot is not null;/);
  assert.match(schemaMigration, /status text not null default 'analyzing' check \(status in \('analyzing','ready','running','completed','partial','failed'\)\)/);
  assert.match(schemaMigration, /status text not null default 'proposed',[\s\S]*financial_reconciliation_automatic_proposals_status_check[\s\S]*status in \('proposed','ambiguous','skipped','deselected','executing','completed','stale','failed'\)/);
  assert.match(schemaMigration, /unique \(run_id, rule_key, base_source_type, base_source_id, signature\)/);
  assert.match(schemaMigration, /check \(origin in \('user','automatic'\)\)/);
  assert.match(schemaMigration, /automatic_trigger text null check \(automatic_trigger in \('manual','scheduled'\)\)/);
  assert.match(schemaMigration, /automatic_run_id uuid null references public\.financial_reconciliation_automatic_runs\(id\)/);
  assert.match(schemaMigration, /automatic_proposal_id uuid null references public\.financial_reconciliation_automatic_proposals\(id\)/);
  assert.match(schemaMigration, /origin = 'user'[\s\S]*automatic_proposal_id is null/);
  assert.match(schemaMigration, /origin = 'automatic'[\s\S]*automatic_proposal_id is not null/);
  for (const table of [
    "financial_reconciliation_automatic_rule_definitions",
    "financial_reconciliation_automatic_rule_configs",
    "financial_reconciliation_automatic_schedule",
    "financial_reconciliation_automatic_runs",
    "financial_reconciliation_automatic_proposals",
  ]) {
    assert.match(schemaMigration, new RegExp(`revoke all on table public\\.${table} from public, anon, authenticated, service_role;`));
    assert.doesNotMatch(schemaMigration, new RegExp(`grant [^;]+ on table public\\.${table} to service_role;`));
  }
  assert.match(schemaMigration, /unique \(priority\) deferrable initially deferred/);
  assert.match(schemaMigration, /status in \('proposed','ambiguous','skipped','deselected','executing','completed','stale','failed'\)/);
  assert.match(schemaMigration, /financial_reconciliation_automatic_proposals_status_check[\s\S]*skipped/);
});

test("POS income migration never gives a temporary table a foreign key to a permanent table", () => {
  const migration = fs.readFileSync(POS_INCOME_MIGRATION_PATH, "utf8");
  const temporaryTableNames = new Set(
    [...migration.matchAll(/create\s+temporary\s+table\s+([a-z0-9_]+)/gi)]
      .map((match) => match[1].toLowerCase()),
  );
  const temporaryTableBlocks = [
    ...migration.matchAll(
      /create\s+temporary\s+table\s+([a-z0-9_]+)\s*\(([\s\S]*?)\)\s+on\s+commit\s+drop\s*;/gi,
    ),
  ];

  assert.ok(temporaryTableBlocks.length > 0, "migration must retain its temporary schema checks");
  for (const [, temporaryTableName, body] of temporaryTableBlocks) {
    for (const reference of body.matchAll(/references\s+((?:[a-z0-9_]+\.)?[a-z0-9_]+)/gi)) {
      const referencedTableName = reference[1].split(".").at(-1).toLowerCase();
      assert.equal(
        temporaryTableNames.has(referencedTableName),
        true,
        `temporary table ${temporaryTableName} illegally references permanent table ${reference[1]}`,
      );
    }
  }

  assert.match(migration, /fr_auto_proposal_memberships_proposal_fkey[\s\S]*confrelid/);
  assert.match(migration, /confdeltype\s+is\s+distinct\s+from\s+'c'/);
  assert.match(migration, /conkey\s+is\s+distinct\s+from/);
  assert.match(migration, /confkey\s+is\s+distinct\s+from/);
});

test("automation analysis migration fixes deterministic matching, ambiguity, and RPC security contracts", () => {
  assert.equal(fs.existsSync(ANALYSIS_MIGRATION_PATH), true, "automation analysis migration must exist");
  const analysisMigration = fs.readFileSync(ANALYSIS_MIGRATION_PATH, "utf8");
  const compactAnalysisMigration = analysisMigration.replace(/\s+/g, " ");

  for (const signature of [
    "financial_reconciliation_match_normalize(p_value text)",
    "financial_reconciliation_match_compact(p_value text)",
    "financial_reconciliation_automatic_build_combinations(p_base jsonb, p_candidates jsonb, p_operators jsonb, p_tolerance numeric, p_max_group_size integer)",
    "financial_reconciliation_automatic_rule_candidates(p_rule_key text, p_rule_version integer, p_difference_allowed numeric, p_max_difference_days integer)",
    "populate_financial_reconciliation_automatic_run(p_run_id uuid)",
    "create_financial_reconciliation_automatic_analysis(p_rule_keys text[], p_mode text, p_actor text, p_client_request_id uuid)",
    "claim_financial_reconciliation_automatic_schedule(p_now timestamptz, p_actor text)",
    "get_financial_reconciliation_automatic_run(p_run_id uuid)",
  ]) {
    const escapedSignature = signature
      .replace(/[.*+?^${}|[\]\\]/g, "\\$&")
      .replace(/\(/g, "\\(\\s*")
      .replace(/\)/g, "\\s*\\)");
    assert.match(compactAnalysisMigration, new RegExp(`create or replace function public\\.${escapedSignature}`));
  }

  assert.match(analysisMigration, /language\s+sql\s+stable\s+strict/);
  for (const extension of ["pgcrypto", "unaccent", "pg_trgm"]) {
    assert.match(analysisMigration, new RegExp(`e\\.extname = '${extension}'`));
  }
  assert.match(analysisMigration, /from pg_catalog\.pg_extension e[\s\S]*join pg_catalog\.pg_namespace n on n\.oid = e\.extnamespace/);
  for (const helper of ["unaccent", "similarity", "word_similarity", "sha256"]) {
    assert.match(analysisMigration, new RegExp(`create or replace function public\\.financial_reconciliation_extension_${helper}\\(`));
    assert.match(analysisMigration, new RegExp(`public\\.financial_reconciliation_extension_${helper}\\(`));
  }
  assert.match(analysisMigration, /select %I\.unaccent\(p_value\)/);
  assert.match(analysisMigration, /select %I\.similarity\(p_left, p_right\)/);
  assert.match(analysisMigration, /select %I\.word_similarity\(p_left, p_right\)/);
  assert.match(analysisMigration, /select pg_catalog\.encode\(%I\.digest\(p_value, 'sha256'::text\), 'hex'::text\)/);
  assert.equal((analysisMigration.match(/public\.financial_reconciliation_extension_sha256\(/g) || []).length, 7);
  assert.doesNotMatch(analysisMigration, /extensions\.(digest|unaccent|similarity|word_similarity)\(/);
  assert.match(analysisMigration, /regexp_split_to_table\([\s\S]*'\[\[:space:\]\]\+'/);
  assert.doesNotMatch(analysisMigration, /'\\\\s\+'/);
  assert.match(analysisMigration, /financial_reconciliation_match_compact[\s\S]*unaccent\(lower\(p_value\)\)[\s\S]*'\[\^\[:alnum:\]\]'/);
  assert.doesNotMatch(analysisMigration, /financial_reconciliation_match_compact[\s\S]{0,240}financial_reconciliation_match_normalize\(p_value\)/);
  assert.match(analysisMigration, /similarity\(normalized_document_description, normalized_bank_description\)/);
  assert.match(analysisMigration, /word_similarity\(normalized_supplier_name, normalized_bank_description\)/);
  assert.match(analysisMigration, /description_score >= 0\.60/);
  assert.match(analysisMigration, /supplier_score >= 0\.70/);
  assert.match(analysisMigration, /document_number.*matched/si);
  assert.match(analysisMigration, /order by base_date, base_id/);
  assert.match(analysisMigration, /round\(\(p_base->>'amount'\)::numeric \* 100\)::bigint/);
  assert.match(analysisMigration, /abs\(calculated_difference_cents\) <= tolerance_cents/);
  assert.match(analysisMigration, /candidate_limit/);
  assert.match(analysisMigration, /cross_base_overlap/);
  assert.match(analysisMigration, /'skipped', 'no_qualifying_combination'/);
  assert.match(analysisMigration, /'skipped', count\(\*\) filter \(where status = 'skipped'\)/);
  assert.match(analysisMigration, /jsonb_array_elements\(p\.candidate_groups\)/);
  assert.match(analysisMigration, /counts = \(select jsonb_build_object\([\s\S]*from public\.financial_reconciliation_automatic_proposals/);
  assert.match(analysisMigration, /Europe\/Lisbon/);
  assert.match(analysisMigration, /where trigger = 'scheduled' and finished_at is null\s+order by scheduled_slot, started_at for update;/);
  assert.doesNotMatch(analysisMigration, /where trigger = 'scheduled' and scheduled_slot = v_slot and finished_at is null/);
  assert.ok(
    analysisMigration.indexOf("where trigger = 'scheduled' and finished_at is null")
      < analysisMigration.indexOf("if v_local::time < v_schedule.time_of_day"),
    "unfinished scheduled runs must resume before a new local-time slot is considered",
  );
  assert.match(analysisMigration, /security definer set search_path = public, pg_temp/g);
  assert.match(analysisMigration, /revoke all on function public\.populate_financial_reconciliation_automatic_run\(uuid\) from public, anon, authenticated;/);
  assert.match(analysisMigration, /revoke all on function public\.financial_reconciliation_match_normalize\(text\) from public, anon, authenticated;/);
  assert.match(analysisMigration, /revoke all on function public\.financial_reconciliation_match_compact\(text\) from public, anon, authenticated;/);
  assert.match(analysisMigration, /grant execute on function public\.financial_reconciliation_match_normalize\(text\) to service_role;/);
  assert.match(analysisMigration, /grant execute on function public\.financial_reconciliation_match_compact\(text\) to service_role;/);
  assert.match(analysisMigration, /grant execute on function public\.get_financial_reconciliation_automatic_run\(uuid\) to service_role;/);
  assert.doesNotMatch(analysisMigration, /grant execute on function public\.[^;]+ to (?:anon|authenticated);/);
});

test("automation performance migration materializes matching work without changing rule semantics", () => {
  assert.equal(
    fs.existsSync(ANALYSIS_PERFORMANCE_MIGRATION_PATH),
    true,
    "automation analysis performance migration must exist",
  );
  const migration = fs.readFileSync(ANALYSIS_PERFORMANCE_MIGRATION_PATH, "utf8");

  assert.match(migration, /create index[^;]+financial_documents[^;]+\(document_date\)/i);
  assert.match(migration, /create index[^;]+import_cgd_extrato_ordem[^;]+\(data\)/i);
  assert.match(migration, /create or replace function public\.financial_reconciliation_automatic_rule_candidates\(/);
  for (const stage of ["bases", "bank_rows", "qualified", "scored"]) {
    assert.match(migration, new RegExp(`${stage}\\s+as\\s+materialized`, "i"));
  }
  assert.match(migration, /b\.data between d\.document_date - p_max_difference_days and d\.document_date \+ p_max_difference_days/);
  assert.match(migration, /d\.document_date >= date '2026-01-01'/);
  assert.match(migration, /b\.data >= date '2026-01-01'/);
  assert.match(migration, /description_score >= 0\.60/);
  assert.match(migration, /supplier_score >= 0\.70/);
  assert.match(migration, /order by base_date, base_id/);
  assert.match(migration, /security definer set search_path = public, pg_temp/);
  assert.match(migration, /revoke all on function public\.financial_reconciliation_automatic_rule_candidates\(text,integer,numeric,integer\) from public, anon, authenticated;/);
  assert.match(migration, /grant execute on function public\.financial_reconciliation_automatic_rule_candidates\(text,integer,numeric,integer\) to service_role;/);
  assert.match(migration, /notify pgrst, 'reload schema';/);
  assert.doesNotMatch(migration, /statement_timeout/i);
});

test("automation candidate lookup keeps the bank date index available for every base record", () => {
  assert.equal(
    fs.existsSync(ANALYSIS_INDEX_LOOKUP_MIGRATION_PATH),
    true,
    "automation analysis index-lookup migration must exist",
  );
  const migration = fs.readFileSync(ANALYSIS_INDEX_LOOKUP_MIGRATION_PATH, "utf8");

  assert.match(migration, /create or replace function public\.financial_reconciliation_automatic_rule_candidates\(/);
  assert.match(migration, /left join lateral \([\s\S]+from public\.import_cgd_extrato_ordem bank/);
  assert.match(migration, /bank\.data between d\.document_date - p_max_difference_days and d\.document_date \+ p_max_difference_days/);
  assert.match(migration, /not exists \([\s\S]+i\.source_type = 'import_cgd_extrato_ordem'[\s\S]+i\.source_id = bank\.id/);
  assert.doesNotMatch(migration, /bank_rows\s+as\s+materialized/i);
  assert.match(migration, /description_score >= 0\.60/);
  assert.match(migration, /supplier_score >= 0\.70/);
  assert.match(migration, /revoke all on function public\.financial_reconciliation_automatic_rule_candidates\(text,integer,numeric,integer\) from public, anon, authenticated;/);
  assert.match(migration, /grant execute on function public\.financial_reconciliation_automatic_rule_candidates\(text,integer,numeric,integer\) to service_role;/);
  assert.doesNotMatch(migration, /statement_timeout/i);
});

test("90-day automation migration installs resumable indexed analysis after Banco v2", () => {
  assert.equal(fs.existsSync(AUTOMATION_90_DAY_MIGRATION_PATH), true,
    "90-day automatic reconciliation migration must exist in the normal migration folder");
  const migration = fs.readFileSync(AUTOMATION_90_DAY_MIGRATION_PATH, "utf8");
  const smokeSql = fs.readFileSync(RPC_SMOKE_PATH, "utf8");

  assert.match(migration, /create table if not exists public\.financial_reconciliation_cgd_match_search/i);
  assert.match(migration, /references public\.import_cgd_extrato_ordem\(id\)\s+on update cascade on delete cascade/i);
  assert.match(migration, /if new\.data is null then[\s\S]*delete from public\.financial_reconciliation_cgd_match_search/i);
  assert.match(migration, /from public\.import_cgd_extrato_ordem bank\s+where bank\.data is not null/i);
  assert.match(migration, /create or replace function public\.continue_financial_reconciliation_automatic_analysis\(/i);
  assert.match(migration, /continue_financial_reconciliation_automatic_analysis\([\s\S]*if v_run\.analysis_total = 0 then[\s\S]*count\(\*\)[\s\S]*analysis_total = v_total/i);
  assert.match(migration, /analysis_total = greatest\(analysis_total, analysis_processed \+ v_page_count\)/i);
  assert.match(migration, /exception when others then[\s\S]*status = 'failed'[\s\S]*analysis_error_code = 'analysis_continuation_failed'[\s\S]*return public\.get_financial_reconciliation_automatic_run/i);
  assert.match(migration, /if v_run\.actor <> p_actor then raise exception[^;]+; end if;[\s\S]*if v_run\.status <> 'analyzing' then[\s\S]*end if;\s+begin[\s\S]*exception when others then/i);
  assert.match(migration, /financial_reconciliation_automatic_candidate_page\([\s\S]*with page as materialized[\s\S]*array\(select page\.id/i);
  assert.match(migration, /create or replace function public\.get_financial_reconciliation_automatic_active_run\(/i);
  assert.match(migration, /trigger = 'manual'[\s\S]*finished_at is null[\s\S]*status in \('analyzing', 'ready'\)/i);
  assert.match(migration, /create or replace function public\.continue_financial_reconciliation_automatic_oldest_analysis\(/i);
  assert.match(migration, /create or replace function public\.claim_financial_reconciliation_automatic_schedule\([\s\S]*unsupported_rule_set/i);
  assert.match(migration, /if v_existing_status = 'failed' then[\s\S]*'slot_failed'/i);
  assert.match(migration, /replace\(v_settings_definition, 'not between 0 and 365', 'not between 0 and 90'\)/i);
  assert.match(migration, /max_difference_days between 0 and 90/i);
  assert.match(smokeSql,
    /2026-08-16-financial-reconciliation-automation-banco-v2\.sql[\s\S]*2026-08-16-financial-reconciliation-automation-90-day-performance\.sql/i);
  assert.match(smokeSql, /Projection sync fixture updated[\s\S]*did not synchronize a delete/i);
  assert.match(smokeSql, /Unauthorized continuation changed the owned run state/i);
  assert.match(smokeSql, /analysis_continuation_failed/i);
  assert.match(smokeSql, /90-day boundary did not include day 90 and exclude day 91/i);
});

test("credit-card automation migration preserves Banco v2 and installs an explicit indexed adapter", () => {
  assert.equal(fs.existsSync(CREDIT_CARD_MIGRATION_PATH), true,
    "credit-card automatic reconciliation migration must exist in the normal migration folder");
  const sql = fs.readFileSync(CREDIT_CARD_MIGRATION_PATH, "utf8");
  const previousSql = fs.readFileSync(AUTOMATION_90_DAY_MIGRATION_PATH, "utf8");
  const smokeSql = fs.readFileSync(RPC_SMOKE_PATH, "utf8");

  assert.match(sql, /financial_documents_cgd_credit_card/);
  assert.match(sql, /payment\s*=\s*'Visa'/);
  assert.match(sql, /description_score\s*>=\s*0\.55/);
  assert.match(sql, /supplier_score\s*>=\s*0\.60/);
  assert.match(sql, /import_cgd_cartao_credito/);
  assert.doesNotMatch(sql, /extensions\.unaccent|extensions\.digest/);

  for (const signature of [
    "financial_reconciliation_automatic_rule_contract",
    "financial_reconciliation_automatic_bank_candidates_for_base_ids",
    "financial_reconciliation_automatic_credit_card_candidates_for_base_ids",
    "financial_reconciliation_automatic_candidates_for_base_ids",
    "financial_reconciliation_automatic_base_page",
    "financial_reconciliation_automatic_base_count",
    "financial_reconciliation_automatic_candidate_page",
    "financial_reconciliation_automatic_single_base_candidates",
    "financial_reconciliation_automatic_rule_candidates",
  ]) {
    assert.match(sql, new RegExp(`create or replace function public\\.${signature}\\(`, "i"), signature);
  }
  for (const indexName of [
    "financial_reconciliation_cgd_credit_card_match_search_date_id_idx",
    "financial_reconciliation_cgd_credit_card_match_search_normalized_trgm_idx",
    "financial_reconciliation_cgd_credit_card_match_search_compact_trgm_idx",
  ]) {
    assert.match(sql, new RegExp(`create index if not exists ${indexName}`, "i"), indexName);
  }
  assert.match(sql, /old\.id is distinct from new\.id[\s\S]*delete from public\.financial_reconciliation_cgd_credit_card_match_search where source_id = old\.id/i);
  assert.match(sql, /new\.id, new\.data, new\.valor, new\.descricao/);
  assert.match(sql, /financial_reconciliation_automatic_rule_contract\([\s\S]*'payment','Banco'[\s\S]*'payment','Visa'/i);
  assert.match(sql, /char_length\(q\.compact_document_number\) >= 4[\s\S]*position\(q\.compact_document_number in q\.compact_description\)[\s\S]*position\(q\.compact_description in q\.compact_document_number\)/i);
  assert.match(sql, /alter table public\.financial_reconciliation_cgd_credit_card_match_search enable row level security/i);
  assert.match(sql, /revoke all on table public\.financial_reconciliation_cgd_credit_card_match_search[\s\S]*grant select on table public\.financial_reconciliation_cgd_credit_card_match_search to service_role/i);

  const functionSource = (migration, functionName) => {
    const match = migration.match(new RegExp(
      `create or replace function public\\.${functionName}\\([\\s\\S]*?\\$body\\$;`,
      "i",
    ));
    assert.ok(match, `${functionName} SQL body must exist`);
    return match[0]
      .replaceAll(functionName, "candidate_adapter")
      .replaceAll("\r\n", "\n");
  };
  assert.equal(
    functionSource(sql, "financial_reconciliation_automatic_bank_candidates_for_base_ids"),
    functionSource(previousSql, "financial_reconciliation_automatic_candidates_for_base_ids"),
    "the named Banco v2 adapter must remain byte-for-byte identical to the indexed base-ID query",
  );

  const ninetyDayInclude = smokeSql.indexOf(
    "\\ir ../supabase-migrations/2026-08-16-financial-reconciliation-automation-90-day-performance.sql",
  );
  const creditCardInclude = smokeSql.indexOf(
    "\\ir ../supabase-migrations/2026-08-16-financial-reconciliation-automation-credit-card-rule.sql",
  );
  assert.ok(creditCardInclude > ninetyDayInclude,
    "credit-card smoke migration must run after the 90-day migration");
  assert.equal((smokeSql.match(/2026-08-16-financial-reconciliation-automation-credit-card-rule\.sql/g) || []).length, 2);
  for (const contract of [
    "credit-card immutable definition and first config",
    "credit-card source rule",
    "credit-card projection INSERT UPDATE ID-change DELETE and data_valor isolation",
    "credit-card exact Visa eligibility and exclusions",
    "credit-card dates exactly 10 and 11 days apart",
    "credit-card symmetric compact document-number containment with four-character minimum",
    "credit-card description score immediately below and at 0.55",
    "credit-card supplier word score immediately below and at 0.60",
    "credit-card independent identity branches",
    "Banco v2 dispatcher IDs and evidence remain byte-for-byte unchanged",
    "credit-card migration reapply is idempotent and preserves administrator settings",
  ]) {
    assert.match(smokeSql, new RegExp(`-- ${contract}`));
  }
});

test("credit-card automation analyzes one immutable managed rule with resumable terminal lifecycle", () => {
  const sql = fs.readFileSync(CREDIT_CARD_MIGRATION_PATH, "utf8");
  const smokeSql = fs.readFileSync(RPC_SMOKE_PATH, "utf8");
  const functionSource = (functionName) => {
    const match = sql.match(new RegExp(
      `create or replace function public\\.${functionName}\\([\\s\\S]*?\\n\\$\\$;`,
      "i",
    ));
    assert.ok(match, `${functionName} replacement must exist in the credit-card migration`);
    return match[0];
  };

  const createSource = functionSource("create_financial_reconciliation_automatic_analysis");
  assert.equal(
    mapRpcError(new Error("Automatic analysis conflict: an unfinished manual run already exists for this actor.")).statusCode,
    409,
    "the one-open-run error must map to a safe conflict response",
  );
  assert.match(createSource,
    /if p_mode <> 'manual_rule' or cardinality\(p_rule_keys\) <> 1 then[\s\S]*Manual automatic analysis requires exactly one selected rule\./i);
  assert.doesNotMatch(createSource, /manual_batch|payment\s*=\s*'Banco'/i);
  assert.match(createSource, /jsonb_array_elements_text\(\s*definition\.destination_source_types\s*\)/i);
  assert.match(createSource, /'destinationSourceType', destination\.source_type/i);
  assert.match(createSource, /source_rule\.matching_source_type\s*=\s*destination\.source_type/i);
  assert.match(createSource, /financial_reconciliation_automatic_base_count\(/i);
  assert.match(createSource, /pg_advisory_xact_lock[\s\S]*p_actor/i);
  assert.match(createSource, /finished_at is null[\s\S]*Automatic analysis conflict:[^']*unfinished manual run already exists/i);
  assert.ok(
    createSource.indexOf("run.trigger = 'manual' and run.finished_at is null")
      < createSource.indexOf("run.client_request_id = p_client_request_id"),
    "the current unfinished run must win over an older idempotency key",
  );

  assert.match(sql,
    /row_number\(\)\s+over\s*\(\s*partition by actor\s+order by started_at desc, created_at desc, id desc\s*\)[\s\S]*status = 'failed'/i);
  assert.match(sql,
    /create unique index if not exists financial_reconciliation_automatic_runs_open_manual_actor_uidx[\s\S]*where trigger = 'manual' and finished_at is null/i);

  const continueSource = functionSource("continue_financial_reconciliation_automatic_analysis");
  assert.match(continueSource, /jsonb_array_length\(v_run\.definition_config_snapshot\) <> 1/i);
  assert.match(continueSource, /financial_reconciliation_automatic_rule_contract\(/i);
  assert.match(continueSource, /v_max_candidates[^;]*v_contract->>'maxCandidates'/i);
  assert.match(continueSource, /v_max_destination_records[^;]*v_contract->>'maxDestinationRecords'/i);
  assert.match(continueSource, /coalesce\(v_rule->>'operator', ''\) not in \('\+', '-'\)/i);
  assert.match(continueSource, /coalesce\(v_max_candidates, 0\) < 1/i);
  assert.match(continueSource, /coalesce\(v_max_destination_records, 0\) < 1/i);
  assert.match(continueSource,
    /jsonb_build_object\(v_destination_source_type, v_rule->>'operator'\)[\s\S]*v_max_destination_records/i);
  assert.doesNotMatch(continueSource, /payment\s*=\s*'Banco'|import_cgd_extrato_ordem', v_operator|candidate_count > 12|numeric, 4\s*\)/i);
  assert.match(continueSource, /financial_reconciliation_automatic_base_count\(/i);
  assert.match(continueSource, /financial_reconciliation_automatic_candidate_page\([\s\S]*25/i);

  const finalizeSource = functionSource("financial_reconciliation_finalize_automatic_analysis");
  assert.match(finalizeSource,
    /status = case when exists \([\s\S]*status = 'proposed'[\s\S]*then 'ready' else 'completed' end/i);
  assert.match(finalizeSource,
    /finished_at = case when exists \([\s\S]*status = 'proposed'[\s\S]*then null else now\(\) end/i);
  assert.ok(finalizeSource.indexOf("cross_base_overlap") < finalizeSource.indexOf("counts ="),
    "overlap resolution must run before persisted proposal counts are recalculated");

  const activeSource = functionSource("get_financial_reconciliation_automatic_active_run");
  assert.match(activeSource, /actor = p_actor[\s\S]*trigger = 'manual'[\s\S]*finished_at is null/i);
  assert.doesNotMatch(activeSource, /status in \('analyzing', 'ready'\)/i);
  const oldestSource = functionSource("continue_financial_reconciliation_automatic_oldest_analysis");
  assert.match(oldestSource, /p_worker is distinct from 'system:reconciliation'/i);
  assert.match(oldestSource,
    /jsonb_typeof\(v_snapshot\) = 'array'[\s\S]*jsonb_array_length\(v_snapshot\) = 1/i);
  assert.match(oldestSource,
    /financial_reconciliation_automatic_rule_contract\([\s\S]*is not null[\s\S]*financial_reconciliation_automatic_base_count\(/i);
  assert.match(oldestSource, /financial_reconciliation_automatic_base_count\(/i);

  for (const contract of [
    "one-rule manual creation validation and immutable snapshots",
    "one open manual run per actor",
    "credit-card 25-base resumable lifecycle and retry idempotency",
    "credit-card one-to-four exact combinations and five-card skip",
    "credit-card ambiguity and candidate limit",
    "zero-executable analysis terminates without visible executable rows",
    "Banco paging and counts remain unchanged",
  ]) {
    assert.match(smokeSql, new RegExp(`-- ${contract}`));
  }
});

test("automatic proposal execution dispatches explicit managed adapters and retains audit evidence", () => {
  const sql = fs.readFileSync(CREDIT_CARD_MIGRATION_PATH, "utf8");
  const smokeSql = fs.readFileSync(RPC_SMOKE_PATH, "utf8");
  const functionSource = (functionName) => {
    const match = sql.match(new RegExp(
      `create or replace function public\\.${functionName}\\([\\s\\S]*?\\n\\$\\$;`,
      "i",
    ));
    assert.ok(match, `${functionName} replacement must exist in the credit-card migration`);
    return match[0];
  };

  const lockSource = functionSource("financial_reconciliation_automatic_lock_destination_items");
  assert.match(lockSource,
    /if p_source_type = 'import_cgd_extrato_ordem' then[\s\S]*join public\.import_cgd_extrato_ordem[\s\S]*elsif p_source_type = 'import_cgd_cartao_credito' then[\s\S]*join public\.import_cgd_cartao_credito/i);
  assert.match(lockSource, /order by bank\.data, bank\.id[\s\S]*for update of bank/i);
  assert.match(lockSource, /order by card\.data, card\.id[\s\S]*for update of card/i);
  assert.match(lockSource, /Automatic reconciliation destination source is unsupported\./i);
  assert.match(lockSource, /get diagnostics v_count = row_count[\s\S]*return v_count/i);
  assert.doesNotMatch(lockSource, /\bexecute\b|\bformat\s*\(/i,
    "destination locking must not use dynamic SQL");

  const executeSource = functionSource("execute_financial_reconciliation_automatic_proposal");
  assert.match(executeSource,
    /v_contract := public\.financial_reconciliation_automatic_rule_contract\(\s*v_proposal\.rule_key,\s*v_proposal\.rule_version\s*\)/i);
  assert.match(executeSource, /v_destination_source_type := v_contract->>'destinationSourceType'/i);
  assert.match(executeSource,
    /v_proposal\.base_source_type <> 'financial_documents'[\s\S]*jsonb_array_length\(v_run\.definition_config_snapshot\) <> 1/i);
  assert.match(executeSource,
    /jsonb_typeof\(v_rule_snapshot->'definition'\) is distinct from 'object'/i,
    "a missing immutable definition must fail closed under SQL NULL semantics");
  assert.match(executeSource,
    /value->>'sourceType' is distinct from v_destination_source_type/i);
  assert.match(executeSource,
    /jsonb_array_length\(v_proposal\.items\)[\s\S]*v_max_destination_records/i);
  assert.match(executeSource,
    /financial_reconciliation_automatic_lock_destination_items\(\s*v_destination_source_type,\s*v_proposal\.items\s*\)/i);
  assert.match(executeSource,
    /financial_reconciliation_automatic_single_base_candidates\([\s\S]*v_proposal\.base_source_id/i);
  assert.match(executeSource,
    /financial_reconciliation_automatic_build_combinations\([\s\S]*jsonb_build_object\(v_destination_source_type, v_rule_snapshot->>'operator'\)[\s\S]*v_max_destination_records/i);
  assert.match(executeSource, /v_combination\.signature is distinct from v_proposal\.signature/i);
  assert.match(executeSource, /v_combination\.items is distinct from v_proposal\.items/i);
  assert.match(executeSource, /v_current_evidence is distinct from v_proposal\.evidence/i);
  assert.match(executeSource,
    /v_combination\.calculated_difference is distinct from v_proposal\.calculated_difference/i);
  assert.match(executeSource,
    /jsonb_build_object\(\s*'sourceType', v_destination_source_type,\s*'operator', v_rule_snapshot->>'operator'\s*\)/i);
  assert.match(executeSource,
    /Automatically completed by rule ['|\s]*\|\|\s*v_rule_snapshot->>'displayName'[\s\S]*v_proposal\.rule_version::text/i);
  assert.match(executeSource,
    /'operatorSnapshot', jsonb_build_object\(\s*v_destination_source_type, v_rule_snapshot->>'operator'\s*\)/i);
  for (const metadataKey of [
    "ruleSnapshot",
    "configSnapshot",
    "operatorSnapshot",
    "baseSnapshot",
    "destinationSnapshots",
    "identityEvidence",
    "proposalSignature",
    "trigger",
    "runId",
    "proposalId",
    "tolerance",
    "calculatedDifference",
  ]) {
    assert.match(executeSource, new RegExp(`'${metadataKey}'`));
  }
  for (const action of ["start", "add_item", "complete", "force_complete"]) {
    assert.match(executeSource,
      new RegExp(`financial_reconciliation_action\\(\\s*'${action}'`, "i"));
  }
  assert.match(executeSource,
    /begin[\s\S]*set status = 'executing'[\s\S]*exception when others then[\s\S]*set status = 'failed'/i);
  assert.match(executeSource,
    /if v_run\.analysis_completed_at is null or v_run\.status = 'analyzing' then[\s\S]*Automatic analysis must finish before proposals can be executed/i);
  assert.match(executeSource,
    /status in \('ambiguous',\s*'skipped',\s*'deselected',\s*'failed'\)/i);
  assert.doesNotMatch(executeSource,
    /financial_documents_cgd_bank_statement|financial_documents_cgd_credit_card|import_cgd_extrato_ordem|import_cgd_cartao_credito|Financial Documents to CGD Bank Statement/i,
    "execution must derive rule and destination details from the allowlisted contract and immutable snapshot");

  assert.match(sql,
    /revoke all on function public\.financial_reconciliation_automatic_lock_destination_items\(text,jsonb\)[\s\S]*from public, anon, authenticated, service_role/i);
  assert.match(sql,
    /revoke all on function public\.execute_financial_reconciliation_automatic_proposal\(uuid,text\)[\s\S]*grant execute on function public\.execute_financial_reconciliation_automatic_proposal\(uuid,text\)[\s\S]*to service_role/i);

  for (const contract of [
    "automatic destination lock helper privileges and dispatch",
    "credit-card automatic execution and audit evidence",
    "credit-card repeated execution is idempotent",
    "credit-card execution stale source and proposal paths",
    "automatic execution rejects unfinished and non-executable proposals",
    "Banco v2 execution evidence remains unchanged",
  ]) {
    assert.match(smokeSql, new RegExp(`-- ${contract}`));
  }
});

test("scheduled automation snapshots deterministic parent batches and advances one rule child at a time", () => {
  const sql = fs.readFileSync(CREDIT_CARD_MIGRATION_PATH, "utf8");
  const smokeSql = fs.readFileSync(RPC_SMOKE_PATH, "utf8");
  const functionSource = (functionName) => {
    const match = sql.match(new RegExp(
      `create or replace function public\\.${functionName}\\([\\s\\S]*?\\n\\$\\$;`,
      "i",
    ));
    assert.ok(match, `${functionName} replacement must exist in the credit-card migration`);
    return match[0];
  };

  assert.match(sql,
    /create table if not exists public\.financial_reconciliation_automatic_batches[\s\S]*scheduled_slot text not null[\s\S]*rule_snapshot jsonb not null[\s\S]*unique \(scheduled_slot\)/i);
  for (const column of ["batch_id", "batch_rule_key", "batch_rule_position", "batch_rule_count"]) {
    assert.match(sql, new RegExp(`add column if not exists ${column}`, "i"), column);
  }
  assert.match(sql, /drop index if exists public\.financial_reconciliation_automatic_runs_scheduled_slot_uidx/i);
  assert.match(sql,
    /create unique index if not exists financial_reconciliation_automatic_runs_legacy_scheduled_slot_uidx[\s\S]*where scheduled_slot is not null and batch_id is null/i);
  assert.match(sql,
    /create unique index if not exists financial_reconciliation_automatic_runs_batch_position_uidx[\s\S]*\(batch_id, batch_rule_position\)/i);
  assert.match(sql,
    /create unique index if not exists financial_reconciliation_automatic_runs_batch_rule_uidx[\s\S]*\(batch_id, batch_rule_key\)/i);
  assert.match(sql,
    /trigger = 'scheduled'\s+and scope = 'rule'[\s\S]*batch_id is not null[\s\S]*batch_rule_key is not null/i);
  assert.match(sql,
    /trigger = 'scheduled'\s+and scope = 'batch'[\s\S]*batch_id is not null/i);
  assert.match(sql,
    /Analysis must be restarted after the 90-day performance upgrade\.[\s\S]*analysis_upgrade_restart_required/i);

  const refreshSource = functionSource("financial_reconciliation_refresh_automatic_batch");
  assert.match(refreshSource, /for update/i);
  assert.match(refreshSource, /jsonb_array_length\(v_batch\.rule_snapshot\)/i);
  assert.match(refreshSource, /count\(\*\) filter \(where run\.status = 'completed'\)/i);
  assert.match(refreshSource, /then 'failed'[\s\S]*then 'partial'[\s\S]*else 'completed'/i);
  assert.doesNotMatch(refreshSource, /error_summary|error_detail/i);

  const claimSource = functionSource("claim_financial_reconciliation_automatic_schedule");
  assert.match(claimSource,
    /jsonb_agg\(jsonb_build_object\([\s\S]*'destinationSourceType', destination\.source_type[\s\S]*order by config\.priority, config\.rule_key/i);
  assert.match(claimSource,
    /jsonb_array_elements\(v_batch\.rule_snapshot\)\s+with ordinality/i);
  assert.match(claimSource,
    /from public\.financial_reconciliation_automatic_batches batch[\s\S]*for update/i);
  assert.match(claimSource,
    /run\.batch_id = v_batch\.id[\s\S]*run\.finished_at is null[\s\S]*order by run\.batch_rule_position/i);
  assert.match(claimSource,
    /'scheduled', 'rule'[\s\S]*jsonb_build_array\(v_selected_rule\)[\s\S]*financial_reconciliation_automatic_base_count/i);
  assert.match(claimSource, /'reason', 'batch_complete'/i);
  assert.doesNotMatch(claimSource,
    /v_enabled_rule_count\s*<>\s*1|matching_source_type\s*=\s*'import_cgd_extrato_ordem'|payment\s*=\s*'Banco'/i);
  assert.doesNotMatch(claimSource, /\bexecute\b|\bformat\s*\(/i,
    "scheduled claiming must not use dynamic SQL dispatch");

  for (const serializer of [
    "get_financial_reconciliation_automatic_run",
    "financial_reconciliation_automatic_progress_or_run",
  ]) {
    const source = functionSource(serializer);
    for (const field of ["batchId", "batchRuleKey", "batchRulePosition", "batchRuleCount"]) {
      assert.match(source, new RegExp(`'${field}'`), `${serializer} ${field}`);
    }
  }

  const settingsSource = functionSource("get_financial_reconciliation_automation_settings");
  assert.match(settingsSource, /'last_scheduled_batch', v_last_scheduled_batch/i);
  assert.match(settingsSource, /from public\.financial_reconciliation_automatic_batches/i);
  assert.doesNotMatch(settingsSource, /error_summary|error_detail/i);
  assert.match(sql,
    /create trigger financial_reconciliation_refresh_automatic_batch_trigger[\s\S]*execute function public\.financial_reconciliation_refresh_automatic_batch_from_run\(\)/i);
  assert.match(sql,
    /alter table public\.financial_reconciliation_automatic_batches enable row level security[\s\S]*revoke all on table public\.financial_reconciliation_automatic_batches[\s\S]*from public, anon, authenticated, service_role/i);
  assert.match(sql,
    /revoke all on function public\.financial_reconciliation_refresh_automatic_batch\(uuid\)[\s\S]*grant execute on function public\.financial_reconciliation_refresh_automatic_batch\(uuid\)[\s\S]*to service_role/i);
  for (const signature of [
    "claim_financial_reconciliation_automatic_schedule\\(timestamptz,text\\)",
    "get_financial_reconciliation_automatic_run\\(uuid\\)",
    "financial_reconciliation_automatic_progress_or_run\\(uuid\\)",
    "get_financial_reconciliation_automation_settings\\(\\)",
  ]) {
    assert.match(sql,
      new RegExp(`revoke all on function public\\.${signature}[\\s\\S]*from public, anon, authenticated, service_role`, "i"));
    assert.match(sql,
      new RegExp(`grant execute on function public\\.${signature}[\\s\\S]*to service_role`, "i"));
  }

  for (const contract of [
    "scheduled parent batch schema security and legacy backfill",
    "scheduled batch snapshots all rules in deterministic priority order",
    "scheduled child resumes before the next rule starts",
    "scheduled snapshot survives settings changes and tomorrow uses new settings",
    "failed scheduled child advances and aggregate batch becomes partial",
    "equal scheduled priorities use the rule-key tie-breaker while Settings rejects duplicates",
    "scheduled retries and cross-midnight heartbeats are idempotent",
    "completed scheduled batch returns stable no-work state",
    "historical scheduled runs remain readable and cannot execute again",
  ]) {
    assert.match(smokeSql, new RegExp(`-- ${contract}`));
  }
});

test("credit-card SQL and production code pin RPC ACLs, reapply, and fixed dispatch", () => {
  const migration = fs.readFileSync(CREDIT_CARD_MIGRATION_PATH, "utf8");
  const smokeSql = fs.readFileSync(RPC_SMOKE_PATH, "utf8");
  const ninetyDayName = "2026-08-16-financial-reconciliation-automation-90-day-performance.sql";
  const creditCardName = "2026-08-16-financial-reconciliation-automation-credit-card-rule.sql";

  const creditCardIncludes = [...smokeSql.matchAll(new RegExp(
    `^\\\\ir \\.\\./supabase-migrations/${creditCardName.replaceAll(".", "\\.")}$`,
    "gm",
  ))];
  assert.equal(creditCardIncludes.length, 2, "smoke must apply the migration once and explicitly reapply it once");
  const ninetyDayInclude = smokeSql.indexOf(`\\ir ../supabase-migrations/${ninetyDayName}`);
  assert.ok(creditCardIncludes[0].index > ninetyDayInclude, "normal credit-card migration must follow the 90-day migration");
  assert.ok(creditCardIncludes[1].index > creditCardIncludes[0].index, "credit-card reapply must follow its normal application");
  assert.match(
    smokeSql.slice(creditCardIncludes[1].index),
    /credit-card migration reapply is idempotent and preserves administrator settings/i,
  );

  const serviceRoleRpcSignatures = [
    "financial_reconciliation_automatic_rule_contract\\(text,integer\\)",
    "financial_reconciliation_automatic_bank_candidates_for_base_ids\\(text,integer,numeric,integer,uuid\\[\\]\\)",
    "financial_reconciliation_automatic_credit_card_candidates_for_base_ids\\(text,integer,numeric,integer,uuid\\[\\]\\)",
    "financial_reconciliation_automatic_candidates_for_base_ids\\(text,integer,numeric,integer,uuid\\[\\]\\)",
    "financial_reconciliation_automatic_base_page\\(text,integer,date,uuid,integer\\)",
    "financial_reconciliation_automatic_base_count\\(text,integer\\)",
    "financial_reconciliation_automatic_candidate_page\\(text,integer,numeric,integer,date,uuid,integer\\)",
    "financial_reconciliation_automatic_single_base_candidates\\(text,integer,numeric,integer,uuid\\)",
    "financial_reconciliation_automatic_rule_candidates\\(text,integer,numeric,integer\\)",
    "continue_financial_reconciliation_automatic_analysis\\(uuid,text\\)",
    "create_financial_reconciliation_automatic_analysis\\(text\\[\\],text,text,uuid\\)",
    "get_financial_reconciliation_automatic_active_run\\(text\\)",
    "continue_financial_reconciliation_automatic_oldest_analysis\\(text\\)",
    "execute_financial_reconciliation_automatic_proposal\\(uuid,text\\)",
    "financial_reconciliation_refresh_automatic_batch\\(uuid\\)",
    "claim_financial_reconciliation_automatic_schedule\\(timestamptz,text\\)",
    "get_financial_reconciliation_automatic_run\\(uuid\\)",
    "financial_reconciliation_automatic_progress_or_run\\(uuid\\)",
    "get_financial_reconciliation_automation_settings\\(\\)",
    "replace_financial_reconciliation_source_rules\\(jsonb\\)",
  ];
  for (const signature of serviceRoleRpcSignatures) {
    assert.match(
      migration,
      new RegExp(`revoke all on function public\\.${signature}\\s+from public, anon, authenticated, service_role;`, "i"),
      `${signature} revoke`,
    );
    assert.match(
      migration,
      new RegExp(`grant execute on function public\\.${signature}\\s+to service_role;`, "i"),
      `${signature} service-role grant`,
    );
  }
  assert.match(
    migration,
    /managed Credit Card source rule must remain enabled with operator \+\./i,
  );
  assert.match(
    migration,
    /managed Bank Statement source rule must remain enabled with operator \+\./i,
  );
  assert.match(
    smokeSql,
    /managed automatic source rules reject operator changes and deletion/i,
  );
  for (const signature of [
    "financial_reconciliation_finalize_automatic_analysis\\(uuid\\)",
    "financial_reconciliation_automatic_lock_destination_items\\(text,jsonb\\)",
    "financial_reconciliation_refresh_automatic_batch_from_run\\(\\)",
  ]) {
    assert.match(
      migration,
      new RegExp(`revoke all on function public\\.${signature}\\s+from public, anon, authenticated, service_role;`, "i"),
      `${signature} internal revoke`,
    );
    assert.doesNotMatch(migration, new RegExp(`grant execute on function public\\.${signature}`, "i"));
  }

  const productionPaths = [
    path.join(__dirname, "..", "api", "_reconciliation-automation.js"),
    MANUAL_HANDLER_PATH,
    SETTINGS_HANDLER_PATH,
    CRON_HANDLER_PATH,
    path.join(__dirname, "..", "app-main.js"),
    path.join(__dirname, "..", "index.html"),
    path.join(__dirname, "..", "styles.css"),
    CREDIT_CARD_MIGRATION_PATH,
  ];
  for (const productionPath of productionPaths) {
    const source = fs.readFileSync(productionPath, "utf8");
    assert.doesNotMatch(source, /analyze_batch/i, `${path.basename(productionPath)} analyze_batch`);
    assert.doesNotMatch(source, /Run batch now/i, `${path.basename(productionPath)} Run batch now`);
    assert.doesNotMatch(source, /\bextensions\./i, `${path.basename(productionPath)} extension-schema call`);
  }

  const functionSource = (functionName) => {
    const match = migration.match(new RegExp(
      `create or replace function public\\.${functionName}\\([\\s\\S]*?\\n\\$\\$;`,
      "i",
    ));
    assert.ok(match, `${functionName} body must exist`);
    return match[0];
  };
  for (const functionName of [
    "financial_reconciliation_automatic_candidates_for_base_ids",
    "financial_reconciliation_automatic_lock_destination_items",
    "execute_financial_reconciliation_automatic_proposal",
    "claim_financial_reconciliation_automatic_schedule",
  ]) {
    assert.doesNotMatch(
      functionSource(functionName),
      /\bexecute\b|\bformat\s*\(/i,
      `${functionName} must not dynamically dispatch SQL, tables, or functions`,
    );
  }
  for (const apiPath of [MANUAL_HANDLER_PATH, SETTINGS_HANDLER_PATH, CRON_HANDLER_PATH]) {
    const source = fs.readFileSync(apiPath, "utf8");
    const rpcFirstArguments = [...source.matchAll(/restQuery\(\s*([^,\r\n]+)/g)]
      .map((match) => match[1].trim());
    for (const firstArgument of rpcFirstArguments) {
      assert.ok(
        /^["']/.test(firstArgument)
          || (apiPath === MANUAL_HANDLER_PATH && firstArgument === "resource"),
        `${path.basename(apiPath)} must use literal or explicitly allowlisted RPC resources`,
      );
    }
  }
  const manualApi = fs.readFileSync(MANUAL_HANDLER_PATH, "utf8");
  assert.match(manualApi, /new Set\(\["rules", "active_run"\]\)\.has\(view\)/);
  assert.match(manualApi,
    /const resource = view === "rules"\s*\? "rpc\/get_financial_reconciliation_automatic_manual_rules"\s*:\s*"rpc\/get_financial_reconciliation_automatic_active_run"/);
});

test("automation execution migration revalidates and completes each proposal atomically", () => {
  assert.equal(fs.existsSync(EXECUTION_MIGRATION_PATH), true, "automation execution migration must exist");
  const executionMigration = fs.readFileSync(EXECUTION_MIGRATION_PATH, "utf8");
  const compactExecutionMigration = executionMigration.replace(/\s+/g, " ");

  for (const signature of [
    "execute_financial_reconciliation_automatic_proposal(p_proposal_id uuid, p_actor text)",
    "finish_financial_reconciliation_automatic_run(p_run_id uuid)",
  ]) {
    const escapedSignature = signature
      .replace(/[.*+?^${}|[\]\\]/g, "\\$&")
      .replace(/\(/g, "\\(\\s*")
      .replace(/\)/g, "\\s*\\)");
    assert.match(compactExecutionMigration, new RegExp(`create or replace function public\\.${escapedSignature}`));
  }

  assert.match(executionMigration, /security definer set search_path = public, pg_temp/g);
  assert.match(executionMigration, /financial_reconciliation_audit_action_check[\s\S]*'automatic_complete'/);
  assert.match(executionMigration, /from public\.financial_reconciliation_automatic_runs[\s\S]*for update/);
  assert.match(executionMigration, /from public\.financial_reconciliation_automatic_proposals[\s\S]*for update/);
  assert.match(executionMigration, /from public\.financial_documents[\s\S]*for update/);
  assert.match(executionMigration, /(?:from|join) public\.import_cgd_extrato_ordem[\s\S]*for update/);
  assert.match(executionMigration, /for share of definition, config, source_rule/);
  assert.match(executionMigration, /if v_proposal\.status = 'completed'[\s\S]*v_proposal\.reconciliation_id/);
  assert.match(executionMigration, /status in \('ambiguous',\s*'skipped',\s*'deselected',\s*'failed'\)/);
  assert.match(executionMigration, /financial_reconciliation_automatic_rule_candidates\(/);
  assert.match(executionMigration, /financial_reconciliation_automatic_build_combinations\(/);
  assert.match(executionMigration, /v_combination\.signature\s+is distinct from\s+v_proposal\.signature/);
  assert.match(executionMigration, /v_combination\.items\s+is distinct from\s+v_proposal\.items/);
  assert.match(executionMigration, /v_current_evidence\s+is distinct from\s+v_proposal\.evidence/);
  assert.match(executionMigration, /v_proposal\.allowed_difference\s+is distinct from\s+\(v_rule_snapshot->>'differenceAllowed'\)::numeric/);
  assert.match(executionMigration, /order by[\s\S]*value->>'sourceType'[\s\S]*value->>'sourceDate'[\s\S]*value->>'sourceId'/);

  const startCall = executionMigration.search(/public\.financial_reconciliation_action\(\s*'start'/);
  const provenanceUpdate = executionMigration.indexOf("origin = 'automatic'");
  const completionCall = Math.max(
    executionMigration.search(/public\.financial_reconciliation_action\(\s*'complete'/),
    executionMigration.search(/public\.financial_reconciliation_action\(\s*'force_complete'/),
  );
  assert.ok(startCall >= 0 && provenanceUpdate > startCall && completionCall > provenanceUpdate,
    "provenance must be set after start and before completion");
  assert.match(executionMigration, /Automatically completed by rule Financial Documents to CGD Bank Statement v1; difference/);
  assert.equal((executionMigration.match(/chr\(8364\)/g) || []).length, 2, "stable comment must contain two euro symbols");
  assert.match(executionMigration, /within allowed tolerance/);
  assert.match(executionMigration, /trigger [^;]+; batch [^;]+\./);
  assert.match(executionMigration, /'automatic_complete'/);
  for (const metadataKey of [
    "ruleSnapshot",
    "configSnapshot",
    "operatorSnapshot",
    "identityEvidence",
    "proposalSignature",
    "trigger",
    "runId",
    "tolerance",
  ]) {
    assert.match(executionMigration, new RegExp(`'${metadataKey}'`));
  }
  assert.match(executionMigration, /set status = 'completed'[\s\S]*reconciliation_id = v_reconciliation_id[\s\S]*completed_at =/);
  assert.match(executionMigration, /v_actual_matching_source_rules\s+is distinct from\s+v_expected_matching_source_rules/);
  assert.match(executionMigration, /v_actual_difference\s+is distinct from\s+v_proposal\.calculated_difference/);
  assert.match(executionMigration, /abs\(v_actual_difference\) > v_proposal\.allowed_difference/);
  assert.match(executionMigration, /exception when others[\s\S]*get stacked diagnostics[\s\S]*set status = 'failed'/);

  assert.match(executionMigration, /set status = 'deselected'[\s\S]*where run_id = v_run\.id[\s\S]*and status = 'proposed'/);
  assert.match(executionMigration, /'completed', count\(\*\) filter \(where status = 'completed'\)/);
  assert.match(executionMigration, /'stale', count\(\*\) filter \(where status = 'stale'\)/);
  assert.match(executionMigration, /'failed', count\(\*\) filter \(where status = 'failed'\)/);
  assert.match(executionMigration, /'skipped', count\(\*\) filter \(where status = 'skipped'\)/);
  assert.match(executionMigration, /status = case[\s\S]*'partial'[\s\S]*'failed'[\s\S]*'completed'/);
  assert.match(executionMigration, /finished_at = now\(\)/);

  assert.match(executionMigration, /pg_get_functiondef\('public\.get_financial_reconciliation_workspace\(uuid,text,jsonb,integer,integer\)'::regprocedure\)/);
  assert.match(executionMigration, /Unexpected reconciliation workspace function definition; could not install automatic provenance\./);
  for (const field of ["origin", "automaticTrigger", "automaticRuleKey", "automaticRuleVersion", "automaticRunId"]) {
    assert.match(executionMigration, new RegExp(`'${field}'`));
  }
  assert.match(executionMigration, /'sourceSummary'/);

  for (const signature of [
    "execute_financial_reconciliation_automatic_proposal(uuid,text)",
    "finish_financial_reconciliation_automatic_run(uuid)",
  ]) {
    const escaped = signature.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    assert.match(executionMigration, new RegExp(`revoke all on function public\\.${escaped}\\s+from public, anon, authenticated, service_role;`));
    assert.match(executionMigration, new RegExp(`grant execute on function public\\.${escaped}\\s+to service_role;`));
  }
  assert.doesNotMatch(executionMigration, /grant [^;]+ on table public\.financial_reconciliation_automatic_/);
});

test("automatic proposals persist immutable base snapshots and expose an RPC-only manual catalog", () => {
  const schemaMigration = fs.readFileSync(SCHEMA_MIGRATION_PATH, "utf8");
  const analysisMigration = fs.readFileSync(ANALYSIS_MIGRATION_PATH, "utf8");
  const executionMigration = fs.readFileSync(EXECUTION_MIGRATION_PATH, "utf8");

  assert.match(schemaMigration, /base_snapshot jsonb not null default '\{\}'::jsonb check \(jsonb_typeof\(base_snapshot\) = 'object'\)/);
  assert.match(schemaMigration, /alter table public\.financial_reconciliation_automatic_proposals[\s\S]*add column if not exists base_snapshot jsonb not null default '\{\}'::jsonb/);
  assert.match(schemaMigration, /financial_reconciliation_automatic_proposals_base_snapshot_object_check[\s\S]*jsonb_typeof\(base_snapshot\) = 'object'/);
  assert.match(schemaMigration, /create or replace function public\.prevent_financial_reconciliation_automatic_proposal_snapshot_change\(\)/);
  assert.match(schemaMigration, /new\.base_snapshot is distinct from old\.base_snapshot[\s\S]*Automatic proposal base snapshot is immutable\./);
  assert.match(schemaMigration, /create trigger financial_reconciliation_automatic_proposal_snapshot_immutable/);

  assert.match(analysisMigration, /'baseSnapshot', p\.base_snapshot/);
  assert.equal((analysisMigration.match(/base_snapshot,\s*candidate_groups/g) || []).length, 3,
    "ambiguous and skipped proposal paths must snapshot their base record");
  assert.match(analysisMigration, /base_snapshot,\s*items,\s*evidence,\s*candidate_groups/);
  assert.equal((analysisMigration.match(/v_base\.base_snapshot/g) || []).length >= 6, true,
    "every proposal path and combination calculation must use the authoritative base snapshot");
  assert.equal((analysisMigration.match(/'displayName', d\.display_name/g) || []).length, 2,
    "manual and scheduled run snapshots must retain the managed friendly rule name");
  assert.match(executionMigration, /v_base\.base_snapshot is distinct from v_proposal\.base_snapshot/);

  assert.match(schemaMigration, /create or replace function public\.get_financial_reconciliation_automatic_manual_rules\(\)/);
  assert.match(schemaMigration, /where c\.enabled\s+and c\.allow_manual_execution/);
  assert.match(schemaMigration, /'rules', v_rules/);
  assert.doesNotMatch(
    schemaMigration.slice(
      schemaMigration.indexOf("create or replace function public.get_financial_reconciliation_automatic_manual_rules()"),
      schemaMigration.indexOf("create or replace function public.replace_financial_reconciliation_automation_settings"),
    ),
    /financial_reconciliation_automatic_schedule|'schedule'|'lastScheduledRun'|updated_by|error_summary/,
  );
  assert.match(schemaMigration, /revoke all on function public\.get_financial_reconciliation_automatic_manual_rules\(\) from public, anon, authenticated, service_role;/);
  assert.match(schemaMigration, /grant execute on function public\.get_financial_reconciliation_automatic_manual_rules\(\) to service_role;/);
  assert.doesNotMatch(schemaMigration, /grant [^;]+ on table public\.financial_reconciliation_automatic_/);
});

test("automation settings RPCs validate and replace the complete payload atomically", () => {
  assert.equal(fs.existsSync(SCHEMA_MIGRATION_PATH), true, "automation schema migration must exist");
  const schemaMigration = fs.readFileSync(SCHEMA_MIGRATION_PATH, "utf8");
  const sourceRuleMigration = fs.readFileSync(SOURCE_RULE_MIGRATION_PATH, "utf8");

  assert.match(schemaMigration, /create or replace function public\.get_financial_reconciliation_automation_settings\(\)/);
  assert.match(schemaMigration, /create or replace function public\.replace_financial_reconciliation_automation_settings\(p_schedule jsonb, p_rules jsonb, p_actor text\)/);
  assert.match(schemaMigration, /security definer set search_path = public, pg_temp/);
  assert.match(schemaMigration, /lock table public\.financial_reconciliation_automatic_rule_configs in share row exclusive mode;/);
  assert.match(schemaMigration, /lock table public\.financial_reconciliation_automatic_schedule in share row exclusive mode;/);
  assert.match(schemaMigration, /lock table public\.financial_reconciliation_source_rules in share row exclusive mode;/);
  assert.match(sourceRuleMigration, /lock table public\.financial_reconciliation_source_rules in share row exclusive mode;/);
  const sourceRuleLock = schemaMigration.indexOf("lock table public.financial_reconciliation_source_rules in share row exclusive mode;");
  const configLock = schemaMigration.indexOf("lock table public.financial_reconciliation_automatic_rule_configs in share row exclusive mode;");
  const scheduleLock = schemaMigration.indexOf("lock table public.financial_reconciliation_automatic_schedule in share row exclusive mode;");
  const managedVersionCheck = schemaMigration.indexOf("Submitted automatic rule version does not match managed configuration.");
  const sourceRuleRecheck = schemaMigration.lastIndexOf("No directional source rule exists for an enabled automatic rule.");
  const settingsUpdate = schemaMigration.indexOf("update public.financial_reconciliation_automatic_rule_configs c");
  const lockedValidation = schemaMigration.slice(sourceRuleLock, settingsUpdate);
  assert.ok(sourceRuleLock < configLock && configLock < scheduleLock, "settings locks must use source-rules/config/schedule order");
  assert.ok(scheduleLock < managedVersionCheck && managedVersionCheck < sourceRuleRecheck, "managed version equality must be checked under lock before source rules");
  assert.ok(scheduleLock < sourceRuleRecheck, "directional source rules must be rechecked after locking");
  assert.match(lockedValidation, /join public\.financial_reconciliation_automatic_rule_configs managed_config/);
  assert.match(lockedValidation, /d\.version = managed_config\.rule_version/);
  assert.doesNotMatch(lockedValidation, /d\.version = \(rule->>'rule_version'\)::integer/);
  assert.match(schemaMigration, /Automation settings require every managed rule exactly once\./);
  assert.match(schemaMigration, /Duplicate automatic rule priority\./);
  assert.match(schemaMigration, /Automatic rule version is invalid\./);
  assert.match(schemaMigration, /No directional source rule exists for an enabled automatic rule\./);
  assert.match(schemaMigration, /'timeOfDay'/);
  assert.match(schemaMigration, /'destinationSourceTypes'/);
  assert.match(schemaMigration, /'lastScheduledRun'/);
  assert.match(schemaMigration, /revoke all on function public\.get_financial_reconciliation_automation_settings\(\) from public, anon, authenticated;/);
  assert.match(schemaMigration, /grant execute on function public\.replace_financial_reconciliation_automation_settings\(jsonb,jsonb,text\) to service_role;/);
  assert.doesNotMatch(schemaMigration, /set rule_version = input\.rule_version/);
  assert.match(schemaMigration, /notify pgrst, 'reload schema';/);
});

test("automation SQL smoke transaction covers reapply, security, validation, rollback, and provenance", () => {
  assert.equal(fs.existsSync(RPC_SMOKE_PATH), true, "automation SQL smoke transaction must exist");
  const smokeSql = fs.readFileSync(RPC_SMOKE_PATH, "utf8");

  assert.match(smokeSql, /^begin;/m);
  assert.match(smokeSql, /\\ir \.\.\/supabase-migrations\/2026-08-14-financial-reconciliation-automation-schema\.sql/g);
  assert.equal((smokeSql.match(/\\ir \.\.\/supabase-migrations\/2026-08-14-financial-reconciliation-automation-schema\.sql/g) || []).length, 2);
  assert.match(
    smokeSql,
    /2026-08-14-financial-reconciliation-automation-analysis\.sql[\s\S]*2026-08-15-financial-reconciliation-automation-analysis-performance\.sql/,
  );
  assert.equal(
    (smokeSql.match(/\\ir \.\.\/supabase-migrations\/2026-08-15-financial-reconciliation-automation-analysis-performance\.sql/g) || []).length,
    2,
  );
  assert.equal(
    (smokeSql.match(/\\ir \.\.\/supabase-migrations\/2026-08-15-financial-reconciliation-automation-candidate-index-lookup\.sql/g) || []).length,
    2,
  );
  assert.equal((smokeSql.match(/\\ir \.\.\/supabase-migrations\/2026-08-14-financial-reconciliation-automation-execution\.sql/g) || []).length, 2);
  assert.match(
    smokeSql,
    /pg_get_functiondef\('public\.financial_reconciliation_automatic_rule_candidates\(text,integer,numeric,integer\)'::regprocedure\)/,
  );
  for (const stage of ["bases", "qualified", "scored"]) {
    assert.match(smokeSql, new RegExp(`${stage}\\\\s\\+as\\\\s\\+materialized`, "i"));
  }
  assert.match(smokeSql, /left join lateral\\s\+\\\(\[\\s\\S\]\+from public\\\.import_cgd_extrato_ordem bank/);
  assert.match(
    smokeSql,
    /if v_candidate_definition ~\* 'bank_rows\\s\+as\\s\+materialized' then[\s\S]+raise exception 'Bank candidate rows must remain index-driven\.'/i,
  );
  for (const contract of [
    "definition/config preservation",
    "RLS and privileges",
    "table constraints",
    "unknown-rule rejection",
    "duplicate-priority rejection",
    "atomic rollback",
    "provenance checks",
    "RPC-only privileges",
    "managed rule version",
    "source-rule lock recheck",
    "priority swap",
    "document-number containment",
    "description score immediately below and at 0.60",
    "supplier word score immediately below and at 0.70",
    "blank identity fields",
    "dates exactly 7 and 8 days apart",
    "differences exactly at and above tolerance",
    "one-to-one and one-to-many sums",
    "independent operators",
    "two valid combinations for one base",
    "cross-base overlap",
    "candidate_limit",
    "Lisbon DST slot claim",
    "cross-midnight scheduled resume",
    "Description threshold below fixture was accepted",
    "Supplier threshold below fixture was accepted",
    "Blank identity fixture produced a candidate",
    "Cross-base overlap did not mark every affected proposal ambiguous",
    "Candidate-limit run did not persist exactly one ambiguous proposal",
    "Description below fixture was not boundary-adjacent",
    "Supplier below fixture was not boundary-adjacent",
    "automatic execution RPC privileges",
    "non-zero automatic completion and idempotency",
    "zero-difference automatic completion",
    "stale amount/date/lock/rule/operator/evidence",
    "post-write rollback and later-proposal isolation",
    "automatic reopen/delete provenance",
    "automatic run finalization",
    "all automatic items were not locked",
    "automatic provenance was not persisted",
    "generated automatic completion comment was not stable",
    "zero-difference automatic completion did not retain structured audit metadata",
    "repeated automatic execution duplicated items or audit rows",
    "stale automatic proposal created a reconciliation",
    "failed proposal left partial lifecycle mutations",
    "later proposal was blocked by an earlier failed RPC transaction",
    "automatic lifecycle action changed provenance",
    "automatic lifecycle snapshots were not rechecked before completion",
    "failed proposal was not persisted as failed",
    "manual rule catalog is filtered and RPC-only",
    "proposal base snapshot is complete and immutable",
    "execution rejects a changed base snapshot",
  ]) {
    assert.match(smokeSql, new RegExp(`-- ${contract}`));
  }
  assert.match(smokeSql, /similarity\('abcdefg', candidate\)[\s\S]*where score < 0\.60[\s\S]*order by 0\.60 - score/);
  assert.match(smokeSql, /word_similarity\('abcdefg', candidate\)[\s\S]*where score < 0\.70[\s\S]*order by 0\.70 - score/);
  assert.match(smokeSql, /0\.60 - v_description_below_score > 0\.05/);
  assert.match(smokeSql, /0\.70 - v_supplier_below_score > 0\.05/);
  assert.match(smokeSql, /Submitted automatic rule version does not match managed configuration\./);
  assert.match(smokeSql, /Mismatched managed rule version partially changed settings\./);
  assert.doesNotMatch(smokeSql, /Settings PUT did not update approved editable fields\./);
  assert.match(smokeSql, /^rollback;/m);

  const manualSmokeSql = fs.readFileSync(MANUAL_RPC_SMOKE_PATH, "utf8");
  assert.match(manualSmokeSql, /\\ir \.\.\/supabase-migrations\/2026-08-14-financial-reconciliation-automation-execution\.sql/);
  assert.match(manualSmokeSql, /Manual lifecycle did not preserve user origin and null automation provenance/);
  assert.match(manualSmokeSql, /automaticTrigger/);
  assert.match(manualSmokeSql, /automaticRuleKey/);
  assert.match(manualSmokeSql, /automaticRuleVersion/);
  assert.match(manualSmokeSql, /automaticRunId/);
});
