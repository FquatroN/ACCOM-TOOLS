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
const SOURCE_RULE_MIGRATION_PATH = path.join(
  __dirname,
  "..",
  "supabase-migrations",
  "2026-08-11-financial-reconciliation-source-rules.sql",
);
const RPC_SMOKE_PATH = path.join(__dirname, "reconciliation-automation-rpc.smoke.sql");

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
  assert.match(schemaMigration, /status text not null default 'proposed' check \(status in \('proposed','ambiguous','deselected','executing','completed','stale','failed'\)\)/);
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
  ]) {
    assert.match(smokeSql, new RegExp(`-- ${contract}`));
  }
  assert.match(smokeSql, /Submitted automatic rule version does not match managed configuration\./);
  assert.match(smokeSql, /Mismatched managed rule version partially changed settings\./);
  assert.doesNotMatch(smokeSql, /Settings PUT did not update approved editable fields\./);
  assert.match(smokeSql, /^rollback;/m);
});
