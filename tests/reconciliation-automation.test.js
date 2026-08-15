const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const {
  AUTOMATIC_RULE_KEY,
  AUTOMATIC_RULE_VERSION,
  AUTOMATIC_TIME_ZONE,
  isCronRequest,
  normalizeAnalyzePayload,
  normalizeAutomationAction,
  normalizeAutomationSettingsPayload,
  normalizeExecutePayload,
  toAutomationPublicResult,
  toAutomationSettingsRpcPayload,
} = require("../api/_reconciliation-automation");
const { mapRpcError } = require("../api/_reconciliation");

const RUN_ID = "00000000-0000-0000-0000-000000000001";
const PROPOSAL_ID = "00000000-0000-0000-0000-000000000002";
const REQUEST_ID = "00000000-0000-0000-0000-000000000003";
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
const CRON_HANDLER_PATH = path.join(__dirname, "..", "api", "reconciliation-automation-cron.js");
const VERCEL_CONFIG_PATH = path.join(__dirname, "..", "vercel.json");
const README_PATH = path.join(__dirname, "..", "README.md");
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
    scope: "batch",
    status: "ready",
    actor: SCHEDULE_ACTOR,
    analysisCompletedAt: "2026-08-15T02:00:01.000Z",
    finishedAt: null,
    definitions: [{ ruleKey: AUTOMATIC_RULE_KEY, priority: 1 }],
    proposals: [],
    ...overrides,
  };
  run.proposals = run.proposals.map((proposal) => ({ runId: run.runId, ...proposal }));
  return run;
}

function mockedSupabase(overrides = {}) {
  return {
    cleanText: (value) => String(value ?? "").trim(),
    parseBody: async (request) => request.body || {},
    requireFeature: async () => ({
      user: { email: "user@example.com", id: "user-1" },
      access: { profile: { id: "profile-1" } },
    }),
    restQuery: async () => ({}),
    sendError: (response, error) => response.status(error.statusCode || 500).json({ error: error.message }),
    ...overrides,
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

test("automation settings accept only editable managed-rule fields", () => {
  assert.deepEqual(normalizeAutomationSettingsPayload(managedSettings()), {
    schedule: { enabled: true, timeOfDay: "02:15", timeZone: "Europe/Lisbon" },
    rules: [{
      ruleKey: AUTOMATIC_RULE_KEY,
      ruleVersion: 1,
      enabled: true,
      allowManualExecution: true,
      includeInScheduledBatch: false,
      differenceAllowedCents: 125,
      maxDifferenceDays: 7,
      priority: 1,
    }],
  });
});

test("automation settings reject invalid managed schedule and editable rule values", () => {
  const cases = [
    ["invalid time", { schedule: { enabled: true, timeOfDay: "24:00", timeZone: AUTOMATIC_TIME_ZONE } }, /time of day/i],
    ["non-Lisbon time zone", { schedule: { enabled: true, timeOfDay: "02:15", timeZone: "UTC" } }, /time zone/i],
    ["negative tolerance", { rules: [{ ...managedSettings().rules[0], differenceAllowed: "-0.01" }] }, /non-negative amount/i],
    ["three-decimal tolerance", { rules: [{ ...managedSettings().rules[0], differenceAllowed: "1.234" }] }, /non-negative amount/i],
    ["day below zero", { rules: [{ ...managedSettings().rules[0], maxDifferenceDays: -1 }] }, /between 0 and 365/i],
    ["day above limit", { rules: [{ ...managedSettings().rules[0], maxDifferenceDays: 366 }] }, /between 0 and 365/i],
    ["unknown rule key", { rules: [{ ...managedSettings().rules[0], ruleKey: "other" }] }, /rule key/i],
    ["unknown rule version", { rules: [{ ...managedSettings().rules[0], ruleVersion: 2 }] }, /rule version/i],
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

test("automation actions are restricted to their public contract", () => {
  assert.equal(normalizeAutomationAction("analyze_rule"), "analyze_rule");
  assert.equal(normalizeAutomationAction("analyze_batch"), "analyze_batch");
  assert.equal(normalizeAutomationAction("execute_selected"), "execute_selected");
  assert.throws(() => normalizeAutomationAction("start"), /automation action/i);
  assert.throws(() => normalizeAutomationAction(1), /automation action/i);
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
    proposalIds: Array(101).fill(PROPOSAL_ID),
  }), /between 1 and 100 unique proposal IDs/);
});

test("execution requires valid unique run and proposal UUIDs", () => {
  assert.deepEqual(normalizeExecutePayload({
    action: "execute_selected",
    runId: RUN_ID,
    proposalIds: [PROPOSAL_ID],
  }), { action: "execute_selected", runId: RUN_ID, proposalIds: [PROPOSAL_ID] });

  for (const payload of [
    { action: "analyze_batch", runId: RUN_ID, proposalIds: [PROPOSAL_ID] },
    { action: "execute_selected", runId: "invalid", proposalIds: [PROPOSAL_ID] },
    { action: "execute_selected", runId: RUN_ID, proposalIds: [] },
    { action: "execute_selected", runId: RUN_ID, proposalIds: [PROPOSAL_ID, PROPOSAL_ID] },
    { action: "execute_selected", runId: RUN_ID, proposalIds: ["invalid"] },
  ]) {
    assert.throws(() => normalizeExecutePayload(payload), /execution action|run id|proposal id/i);
  }
});

test("settings RPC payload uses only managed snake-case fields and integer cents", () => {
  const settings = normalizeAutomationSettingsPayload(managedSettings());
  assert.deepEqual(toAutomationSettingsRpcPayload(settings, "user-1"), {
    p_schedule: { enabled: true, time_of_day: "02:15", time_zone: "Europe/Lisbon" },
    p_rules: [{
      rule_key: AUTOMATIC_RULE_KEY,
      rule_version: 1,
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

test("scheduled heartbeat returns safe database reasons when disabled or not due", async () => {
  const cases = [
    ["schedule_disabled", "2026-08-15T01:00:00.000Z"],
    ["before_scheduled_time", "2026-08-15T01:59:00.000Z"],
    ["no_enabled_rules", "2026-08-15T02:00:00.000Z"],
  ];

  for (const [reason, nowIso] of cases) {
    const calls = [];
    const response = responseRecorder();
    await withCronEnvironment(nowIso, async () => {
      await withMockedHandler(CRON_HANDLER_PATH, mockedSupabase({
        restQuery: async (resource, options) => {
          calls.push({ resource, options });
          return { claimed: false, reason, diagnostic: "hidden schedule state" };
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
    assert.deepEqual(response.body, { ok: true, claimed: false, reason, hasMore: false }, reason);
    assert.doesNotMatch(JSON.stringify(response.body), /hidden schedule state/);
  }
});

test("first scheduled claim populates analysis and executes proposals in stable priority and base order", async () => {
  const lowPriorityRule = "future_low_priority_rule";
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
    analysisCompletedAt: null,
    definitions: [
      { ruleKey: lowPriorityRule, priority: 2 },
      { ruleKey: AUTOMATIC_RULE_KEY, priority: 1 },
    ],
  });
  const analyzedRun = scheduledRun({
    definitions: pendingRun.definitions,
    proposals: [
      { id: proposalA, ruleKey: lowPriorityRule, baseSourceDate: "2026-08-01", baseSourceId: baseA, status: "proposed" },
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
          return { claimed: true, resumed: false, run: pendingRun, internal_error: "hidden claim detail" };
        }
        if (resource === "rpc/populate_financial_reconciliation_automatic_run") return analyzedRun;
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
      resource: "rpc/populate_financial_reconciliation_automatic_run",
      options: { method: "POST", body: { p_run_id: RUN_ID } },
    },
    ...[proposalC, proposalB, proposalA].map((proposalId) => ({
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
            return { claimed: true, resumed: true, run };
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
    ["populate", scheduledRun({ status: "analyzing", analysisCompletedAt: null }), "rpc/populate_financial_reconciliation_automatic_run"],
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
            return { claimed: true, resumed: true, run: claimedRun };
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
            return { claimed: true, resumed: true, run };
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

test("scheduled heartbeat enforces populate, refresh, and finalize phase postconditions", async () => {
  const proposal = {
    id: uuidFor(24),
    ruleKey: AUTOMATIC_RULE_KEY,
    baseSourceDate: "2026-08-01",
    baseSourceId: uuidFor(124),
    status: "proposed",
  };
  const analyzingRun = scheduledRun({
    status: "analyzing",
    analysisCompletedAt: null,
  });
  const terminalRun = scheduledRun({
    status: "completed",
    finishedAt: "2026-08-15T02:00:10.000Z",
  });
  const cases = [
    {
      name: "populate remains analyzing",
      claimedRun: analyzingRun,
      responses: {
        "rpc/populate_financial_reconciliation_automatic_run": analyzingRun,
        "rpc/finish_financial_reconciliation_automatic_run": terminalRun,
      },
      expectedCalls: [
        "rpc/claim_financial_reconciliation_automatic_schedule",
        "rpc/populate_financial_reconciliation_automatic_run",
      ],
    },
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
            return { claimed: true, resumed: true, run: claimedRun };
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
    scope: "batch",
    status: "completed",
    actor: SCHEDULE_ACTOR,
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
          return { claimed: true, resumed: true, run: poisonedRun };
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
  const poisonedClaim = Object.create({ claimed: true, resumed: true, run: finishedRun });
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
          return { claimed: true, resumed: true, run: scheduledRun({ proposals }) };
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
        return { claimed: true, resumed: true, run: scheduledRun({ proposals: currentProposals }) };
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
      return { claimed: true, resumed: claimCount > 1, run: finishedRun };
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
    ["source_type", "sourceType"],
    ["source_id", "sourceId"],
    ["source_date", "sourceDate"],
    ["amount_snapshot", "amountSnapshot"],
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

test("automation settings GET authorizes and calls only the settings RPC", async () => {
  const authorizations = [];
  const calls = [];
  const response = responseRecorder();
  await withMockedHandler(SETTINGS_HANDLER_PATH, mockedSupabase({
    requireFeature: async (_request, area, feature) => {
      authorizations.push({ area, feature });
      return {
        user: { email: "admin@example.com", id: "admin-1" },
        access: { profile: { id: "admin-profile" } },
      };
    },
    restQuery: async (resource, options) => {
      calls.push({ resource, options });
      return {
        schedule: { enabled: true, time_of_day: "02:15", time_zone: AUTOMATIC_TIME_ZONE },
        rules: [{ rule_key: AUTOMATIC_RULE_KEY, rule_version: 1, diagnostic: "hidden" }],
      };
    },
  }), async (handler) => {
    await handler({ method: "GET" }, response);
  });

  assert.deepEqual(authorizations, [{ area: "settings", feature: "financial-reconciliation" }]);
  assert.deepEqual(calls, [{
    resource: "rpc/get_financial_reconciliation_automation_settings",
    options: { method: "POST", body: {} },
  }]);
  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, {
    schedule: { enabled: true, timeOfDay: "02:15", timeZone: AUTOMATIC_TIME_ZONE },
    rules: [{ ruleKey: AUTOMATIC_RULE_KEY, ruleVersion: 1 }],
  });
});

test("automation settings PUT normalizes the complete payload and actor into one replacement RPC", async () => {
  const calls = [];
  const response = responseRecorder();
  await withMockedHandler(SETTINGS_HANDLER_PATH, mockedSupabase({
    requireFeature: async () => ({
      user: { email: " admin@example.com ", id: "admin-1" },
      access: { profile: { id: "admin-profile" } },
    }),
    restQuery: async (resource, options) => {
      calls.push({ resource, options });
      return { schedule: { time_of_day: "02:15" }, rules: [{ rule_key: AUTOMATIC_RULE_KEY }] };
    },
  }), async (handler) => {
    await handler({ method: "PUT", body: managedSettings() }, response);
  });

  assert.deepEqual(calls, [{
    resource: "rpc/replace_financial_reconciliation_automation_settings",
    options: {
      method: "POST",
      body: {
        p_schedule: { enabled: true, time_of_day: "02:15", time_zone: AUTOMATIC_TIME_ZONE },
        p_rules: [{
          rule_key: AUTOMATIC_RULE_KEY,
          rule_version: 1,
          enabled: true,
          allow_manual_execution: true,
          include_in_scheduled_batch: false,
          difference_allowed: "1.25",
          max_difference_days: 7,
          priority: 1,
        }],
        p_actor: "admin@example.com",
      },
    },
  }]);
  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, {
    schedule: { timeOfDay: "02:15" },
    rules: [{ ruleKey: AUTOMATIC_RULE_KEY }],
  });
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

test("manual automation GET exposes only the app-authorized enabled manual rule catalog", async () => {
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
        rules: [{
          rule_key: AUTOMATIC_RULE_KEY,
          rule_version: 1,
          display_name: "Financial Documents to CGD Bank Statement",
          enabled: true,
          allow_manual_execution: true,
          difference_allowed: "1.00",
          max_difference_days: 7,
          diagnostic: "hidden",
        }],
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
    rules: [{
      ruleKey: AUTOMATIC_RULE_KEY,
      ruleVersion: 1,
      displayName: "Financial Documents to CGD Bank Statement",
      enabled: true,
      allowManualExecution: true,
      differenceAllowed: "1.00",
      maxDifferenceDays: 7,
    }],
  });
  assert.equal(Object.hasOwn(response.body, "schedule"), false);
});

test("analyze_rule authorizes app access and sends exactly one manually enabled rule", async () => {
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
      return { run_id: RUN_ID, status: "ready" };
    },
  }), async (handler) => {
    await handler({
      method: "POST",
      body: { action: "analyze_rule", ruleKeys: [AUTOMATIC_RULE_KEY], clientRequestId: REQUEST_ID },
    }, response);
  });

  assert.deepEqual(authorizations, [{ area: "app", feature: "financial-reconciliation" }]);
  assert.deepEqual(calls, [{
    resource: "rpc/create_financial_reconciliation_automatic_analysis",
    options: {
      method: "POST",
      body: {
        p_rule_keys: [AUTOMATIC_RULE_KEY],
        p_mode: "manual_rule",
        p_actor: "user@example.com",
        p_client_request_id: REQUEST_ID,
      },
    },
  }]);
  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, { runId: RUN_ID, status: "ready" });
});

test("analyze_batch authorizes settings access and analyzes every enabled batch rule without execution", async () => {
  const authorizations = [];
  const calls = [];
  const response = responseRecorder();
  await withMockedHandler(MANUAL_HANDLER_PATH, mockedSupabase({
    requireFeature: async (_request, area, feature) => {
      authorizations.push({ area, feature });
      return {
        user: { email: "admin@example.com", id: "admin-1" },
        access: { profile: { id: "admin-profile" } },
      };
    },
    restQuery: async (resource, options) => {
      calls.push({ resource, options });
      if (resource === "rpc/get_financial_reconciliation_automation_settings") {
        return {
          rules: [{
            rule_key: AUTOMATIC_RULE_KEY,
            enabled: true,
            include_in_scheduled_batch: true,
          }],
        };
      }
      return { run_id: RUN_ID, scope: "batch", status: "ready" };
    },
  }), async (handler) => {
    await handler({ method: "POST", body: { action: "analyze_batch", clientRequestId: REQUEST_ID } }, response);
  });

  assert.deepEqual(authorizations, [{ area: "settings", feature: "financial-reconciliation" }]);
  assert.deepEqual(calls, [
    {
      resource: "rpc/get_financial_reconciliation_automation_settings",
      options: { method: "POST", body: {} },
    },
    {
      resource: "rpc/create_financial_reconciliation_automatic_analysis",
      options: {
        method: "POST",
        body: {
          p_rule_keys: [AUTOMATIC_RULE_KEY],
          p_mode: "manual_batch",
          p_actor: "admin@example.com",
          p_client_request_id: REQUEST_ID,
        },
      },
    },
  ]);
  assert.equal(calls.some(({ resource }) => resource.includes("execute_financial")), false);
  assert.deepEqual(response.body, { runId: RUN_ID, scope: "batch", status: "ready" });
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
    [MANUAL_HANDLER_PATH, { method: "POST", body: { action: "analyze_batch", clientRequestId: REQUEST_ID } }],
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
  assert.equal((smokeSql.match(/\\ir \.\.\/supabase-migrations\/2026-08-14-financial-reconciliation-automation-execution\.sql/g) || []).length, 2);
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
