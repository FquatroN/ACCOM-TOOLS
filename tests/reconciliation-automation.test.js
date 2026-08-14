const test = require("node:test");
const assert = require("node:assert/strict");

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
