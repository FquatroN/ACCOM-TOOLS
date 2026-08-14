# Automatic Financial Reconciliation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add managed, deterministic reconciliation rules that users can analyze and selectively execute, and that one protected daily batch can execute automatically with complete provenance and audit evidence.

**Architecture:** Keep rule definitions and the authoritative matching/execution engine in versioned Supabase migrations. Add a small CommonJS contract module for API validation, dedicated settings/manual/scheduler endpoints, and focused UI panels in the existing single-page application. Analysis persists immutable proposals without locks; execution revalidates and atomically uses the existing reconciliation lifecycle one proposal per transaction.

**Tech Stack:** Vanilla JavaScript/CommonJS, Node.js built-in test runner, Vercel Functions and Cron, Supabase/PostgreSQL PL/pgSQL, `pgcrypto`, `unaccent`, `pg_trgm`, existing HTML/CSS SPA.

## Global Constraints

- Follow `docs/superpowers/specs/2026-08-13-automatic-financial-reconciliation-design.md` exactly.
- Eligible source records must be dated `2026-01-01` or later.
- Matching is deterministic; do not use AI, embeddings, or nondeterministic ranking.
- A managed definition is read-only and versioned; only enabled state, manual/scheduled modes, difference allowance, date allowance, and priority are editable.
- Store all money as `numeric(14,2)` in PostgreSQL and validate/calculate client inputs as integer cents; never compare binary floating-point money inside the authoritative engine.
- Interpret `difference_allowed` as an inclusive absolute tolerance and `max_difference_days` as an inclusive symmetric calendar-day window.
- Manual analysis does not lock or mutate source records. Every selected proposal is revalidated during execution.
- Ambiguous groups are never automatically executed and similarity scores never break ambiguity ties.
- The first definition is `financial_documents_cgd_bank_statement`, version `1`, disabled by default, with default tolerance `0.00` and default date window `7`.
- Version 1 uses a maximum of four CGD destination records per group and at most twelve identity-qualified CGD candidates per base. If the candidate limit is exceeded, mark the base ambiguous with reason `candidate_limit`; do not truncate and execute a partial search.
- Version 1 document-number containment requires at least four compact alphanumeric characters; description trigram similarity is `>= 0.60`; supplier-to-bank word similarity is `>= 0.70`; blank inputs never pass.
- Every destination record in a proposed multi-record group must pass at least one identity branch against the base record.
- Cross-base overlap is ambiguous: if one source record appears in otherwise unique proposals for two base records in the same analysis, mark all affected proposals ambiguous.
- Scheduled execution uses `Europe/Lisbon`, claims at most one slot per Lisbon calendar day, resumes unfinished work, and executes at most 25 proposals per cron invocation.
- Manual execution accepts at most 100 proposal IDs per request.
- `origin` values are `user` and `automatic`; automatic triggers are `manual` and `scheduled`.
- Existing manual reconciliations, source-rule snapshots, item locks, reopen/delete behavior, and audit history must remain compatible.
- All new automation tables use RLS, grant no direct access to `anon` or `authenticated`, and are mutated only through validated service-role RPCs.
- Preserve all existing untracked user files; stage only files named by each task.

## File Structure

### New files

- `api/_reconciliation-automation.js` — request/config normalization, result shaping, cron authentication, and stable public constants; contains no database access.
- `api/reconciliation-automation-settings.js` — authorized GET/PUT for definitions, configurations, and shared schedule.
- `api/reconciliation-automation.js` — authorized manual Analyze, Run batch now, Execute selected, and run-detail endpoint.
- `api/reconciliation-automation-cron.js` — protected heartbeat that claims/resumes the daily scheduled batch and executes bounded work.
- `supabase-migrations/2026-08-14-financial-reconciliation-automation-schema.sql` — catalog, configuration, schedule, run/proposal tables, provenance columns, constraints, RLS, and settings RPCs.
- `supabase-migrations/2026-08-14-financial-reconciliation-automation-analysis.sql` — deterministic normalization, identity, combination, slot-claim, and proposal-analysis functions.
- `supabase-migrations/2026-08-14-financial-reconciliation-automation-execution.sql` — atomic proposal execution, generated comments, provenance, run summaries, and workspace enrichment.
- `tests/reconciliation-automation.test.js` — pure contracts and mocked API endpoint tests.
- `tests/reconciliation-automation-rpc.smoke.sql` — authoritative database schema, analysis, ambiguity, execution, idempotency, and provenance smoke contract.
- `tests/reconciliation-automation-ui.test.js` — executable extraction tests for settings/workbench renderers and event flows.

### Existing files modified

- `api/_reconciliation.js` — accept and map automatic lifecycle/provenance validation errors without changing manual request contracts.
- `app-main.js` — automation settings state/render/save, analysis state/render/actions, proposal execution, and provenance badges.
- `index.html` — Reconciliation Settings subtabs and Automatic reconciliation workbench/results markup.
- `styles.css` — scoped responsive styles for the rule list, schedule, proposals, evidence, and provenance badges.
- `vercel.json` — one one-minute heartbeat for the reconciliation automation scheduler; the database still claims only one daily Lisbon slot.
- `tests/reconciliation-density.test.js` — structural safeguards for existing reconciliation layout and new provenance/history markup.
- `tests/reconciliation-rpc.smoke.sql` — verify manual reconciliation behavior remains unchanged after all automation migrations.
- `README.md` — migration order, required `CRON_SECRET`, scheduler behavior, and staged rollout instructions.

---

### Task 1: Define Automation Contracts and Validation

**Files:**
- Create: `api/_reconciliation-automation.js`
- Create: `tests/reconciliation-automation.test.js`
- Modify: `api/_reconciliation.js`

**Interfaces:**
- Consumes: `SOURCE_TYPES`, `normalizeSourceType`, and `mapRpcError` from `api/_reconciliation.js`.
- Produces:
  - `AUTOMATIC_RULE_KEY = "financial_documents_cgd_bank_statement"`.
  - `AUTOMATIC_RULE_VERSION = 1`.
  - `AUTOMATIC_TIME_ZONE = "Europe/Lisbon"`.
  - `normalizeAutomationSettingsPayload(value)` returning `{ schedule, rules }`.
  - `normalizeAutomationAction(value)` returning `analyze_rule`, `analyze_batch`, or `execute_selected`.
  - `normalizeAnalyzePayload(value)` returning `{ action, ruleKeys, clientRequestId }`.
  - `normalizeExecutePayload(value)` returning `{ action, runId, proposalIds }`.
  - `toAutomationSettingsRpcPayload(settings, actor)` returning `{ p_schedule, p_rules, p_actor }` with snake_case rows and a fixed two-decimal tolerance string.
  - `isCronRequest(req, cronSecret)` returning boolean.
  - `toAutomationPublicResult(value)` returning the stable camelCase API payload.

- [ ] **Step 1: Write failing contract tests**

Add table-driven tests that require the exact constraints:

```js
test("automation settings accept only editable managed-rule fields", () => {
  assert.deepEqual(normalizeAutomationSettingsPayload({
    schedule: { enabled: true, timeOfDay: "02:15", timeZone: "Europe/Lisbon" },
    rules: [{
      ruleKey: AUTOMATIC_RULE_KEY,
      ruleVersion: 1,
      enabled: true,
      allowManualExecution: true,
      includeInScheduledBatch: false,
      differenceAllowed: "1.25",
      maxDifferenceDays: 7,
      priority: 1,
    }],
  }), {
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

test("execution rejects duplicate or oversized proposal selections", () => {
  assert.throws(() => normalizeExecutePayload({
    action: "execute_selected",
    runId: "00000000-0000-0000-0000-000000000001",
    proposalIds: Array(101).fill("00000000-0000-0000-0000-000000000002"),
  }), /between 1 and 100 unique proposal IDs/);
});
```

Cover invalid times, non-Lisbon timezone, negative or more-than-two-decimal tolerance, day values outside `0..365`, duplicate priorities, unknown keys/versions, editable definition fields, invalid UUIDs, missing rule keys, duplicate IDs, and cron authentication through `x-vercel-cron` or `Authorization: Bearer <CRON_SECRET>`.

- [ ] **Step 2: Run the focused test to verify RED**

Run: `node --test --test-isolation=none tests/reconciliation-automation.test.js`

Expected: FAIL because `api/_reconciliation-automation.js` does not exist.

- [ ] **Step 3: Implement the pure contract module**

Use strict parsing rather than coercion. The public shapes must be:

```js
const AUTOMATIC_RULE_KEY = "financial_documents_cgd_bank_statement";
const AUTOMATIC_RULE_VERSION = 1;
const AUTOMATIC_TIME_ZONE = "Europe/Lisbon";
const AUTOMATION_ACTIONS = new Set(["analyze_rule", "analyze_batch", "execute_selected"]);

function moneyToCents(value, label) {
  const text = String(value ?? "").trim();
  if (!/^\d+(?:\.\d{1,2})?$/.test(text)) throw inputError(`${label} must be a non-negative amount with at most two decimals.`);
  const [whole, fraction = ""] = text.split(".");
  const cents = (Number(whole) * 100) + Number(fraction.padEnd(2, "0"));
  if (!Number.isSafeInteger(cents)) throw inputError(`${label} is too large.`);
  return cents;
}
```

Validate UUIDs with one shared anchored expression, deduplicate before returning, require `ruleKeys` to contain only the version-1 key, reject any rule object containing `baseSourceType`, `destinationSourceTypes`, `logic`, `definition`, or thresholds, and return copies rather than mutating input. `toAutomationSettingsRpcPayload` must convert cents without floating-point division: build its decimal string from integer quotient/remainder and emit `rule_key`, `rule_version`, `enabled`, `allow_manual_execution`, `include_in_scheduled_batch`, `difference_allowed`, `max_difference_days`, and `priority` only.

- [ ] **Step 4: Extend RPC error mapping**

In `api/_reconciliation.js`, map messages containing `automatic rule`, `automation proposal`, `scheduled slot`, `stale proposal`, `candidate limit`, or `ambiguous` to safe `400`/`409` statuses. Keep existing manual mappings unchanged and add focused assertions to `tests/reconciliation-automation.test.js`.

- [ ] **Step 5: Run focused and full Node tests**

Run:

```powershell
node --check api/_reconciliation-automation.js
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none tests/*.test.js
```

Expected: all pass.

- [ ] **Step 6: Commit**

```powershell
git add -- api/_reconciliation-automation.js api/_reconciliation.js tests/reconciliation-automation.test.js
git commit -m "feat: define reconciliation automation contracts"
```

---

### Task 2: Add Managed Definitions, Configuration, Schedule, and Provenance Schema

**Files:**
- Create: `supabase-migrations/2026-08-14-financial-reconciliation-automation-schema.sql`
- Create: `tests/reconciliation-automation-rpc.smoke.sql`
- Modify: `tests/reconciliation-automation.test.js`

**Interfaces:**
- Consumes: the four existing source names and existing `financial_reconciliations` table.
- Produces:
  - `get_financial_reconciliation_automation_settings()`.
  - `replace_financial_reconciliation_automation_settings(jsonb,jsonb,text)`.
  - Tables `financial_reconciliation_automatic_rule_definitions`, `financial_reconciliation_automatic_rule_configs`, `financial_reconciliation_automatic_schedule`, `financial_reconciliation_automatic_runs`, and `financial_reconciliation_automatic_proposals`.
  - Provenance columns on `financial_reconciliations`.

- [ ] **Step 1: Add failing migration source-contract tests**

Read the migration in `tests/reconciliation-automation.test.js` and assert exact table names, RLS statements, service-role grants, definition seed, default-disabled config, provenance checks, RPC signatures, and notification:

```js
assert.match(schemaMigration, /create table if not exists public\.financial_reconciliation_automatic_rule_definitions/);
assert.match(schemaMigration, /'financial_documents_cgd_bank_statement',\s*1/);
assert.match(schemaMigration, /enabled boolean not null default false/);
assert.match(schemaMigration, /check \(origin in \('user','automatic'\)\)/);
assert.match(schemaMigration, /create or replace function public\.replace_financial_reconciliation_automation_settings\(p_schedule jsonb, p_rules jsonb, p_actor text\)/);
assert.match(schemaMigration, /notify pgrst, 'reload schema';/);
```

- [ ] **Step 2: Run the focused test to verify RED**

Run: `node --test --test-isolation=none tests/reconciliation-automation.test.js`

Expected: FAIL because the schema migration does not exist.

- [ ] **Step 3: Create the schema migration**

Create extensions and tables with these authoritative columns and constraints:

```sql
create extension if not exists pgcrypto;
create extension if not exists unaccent;
create extension if not exists pg_trgm;

create table if not exists public.financial_reconciliation_automatic_rule_definitions (
  rule_key text not null,
  version integer not null check (version > 0),
  display_name text not null,
  base_source_type text not null,
  destination_source_types jsonb not null check (jsonb_typeof(destination_source_types) = 'array'),
  logic_description text not null,
  definition jsonb not null check (jsonb_typeof(definition) = 'object'),
  created_at timestamptz not null default now(),
  primary key (rule_key, version)
);

create table if not exists public.financial_reconciliation_automatic_rule_configs (
  rule_key text primary key,
  rule_version integer not null,
  enabled boolean not null default false,
  allow_manual_execution boolean not null default false,
  include_in_scheduled_batch boolean not null default false,
  difference_allowed numeric(14,2) not null default 0 check (difference_allowed >= 0),
  max_difference_days integer not null default 7 check (max_difference_days between 0 and 365),
  priority integer not null check (priority > 0),
  updated_by text not null default '',
  updated_at timestamptz not null default now(),
  foreign key (rule_key, rule_version) references public.financial_reconciliation_automatic_rule_definitions(rule_key, version),
  unique (priority)
);
```

Add the singleton schedule with `id boolean primary key default true check (id)`, `enabled`, `time_of_day time not null default '02:00'`, fixed timezone check, update actor/time. Add runs with `trigger` (`manual`/`scheduled`), `scope` (`rule`/`batch`), status checks, UUID `client_request_id`, nullable unique `scheduled_slot`, `analysis_completed_at`, snapshots and counts. Enforce manual idempotency with unique `(actor, client_request_id)` and scheduled idempotency with a partial unique index on non-null `scheduled_slot`. Add proposals with rule identity, base identity, `items/evidence/candidate_groups` JSON, numeric difference/tolerance, status check, signature, reconciliation link, safe error, and unique `(run_id, rule_key, base_source_type, base_source_id, signature)`.

Add these provenance columns to `financial_reconciliations`:

```sql
origin text not null default 'user' check (origin in ('user','automatic')),
automatic_trigger text null check (automatic_trigger in ('manual','scheduled')),
automatic_rule_key text null,
automatic_rule_version integer null,
automatic_run_id uuid null references public.financial_reconciliation_automatic_runs(id),
automatic_proposal_id uuid null references public.financial_reconciliation_automatic_proposals(id)
```

Add one check requiring all automatic fields to be null for `origin='user'`, and requiring trigger/rule/version/run/proposal for `origin='automatic'`.

- [ ] **Step 4: Seed the disabled definition and configuration**

Seed key/version 1 with source arrays, the three identity branches, thresholds `0.60` and `0.70`, compact document minimum `4`, maximum destination group `4`, and maximum candidate count `12` in the definition JSON. Seed its config disabled, manual false, scheduled false, tolerance `0.00`, days `7`, priority `1`. Use `on conflict ... do update` only for read-only definition metadata; use `do nothing` for configuration so a reapplication cannot overwrite administrator values.

- [ ] **Step 5: Add RLS and atomic settings RPCs**

Enable RLS on all five tables, revoke from `public, anon, authenticated`, and grant required table access only to `service_role`. `get_financial_reconciliation_automation_settings()` returns camelCase definition/config objects plus schedule and last scheduled run. `replace_financial_reconciliation_automation_settings(...)` validates the entire payload before locking config and schedule tables, rejects unknown versions/duplicates/bad priorities/bad values, verifies every destination has an existing directional source rule, then replaces config values and updates the singleton schedule atomically.

Revoke RPC execution from `public, anon, authenticated`; grant to `service_role`.

- [ ] **Step 6: Write and run the schema smoke transaction**

In `tests/reconciliation-automation-rpc.smoke.sql`, start a transaction, apply the schema migration twice, assert definition/config preservation, RLS/privileges, constraints, unknown-rule rejection, duplicate-priority rejection, atomic rollback, and provenance checks, then `rollback`.

Run when a disposable database is configured:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

Expected: exit `0`. If `psql` or `SUPABASE_DB_URL` is unavailable, record the command as a required rollout gate; do not claim database verification.

- [ ] **Step 7: Run Node tests and commit**

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none tests/*.test.js
git add -- supabase-migrations/2026-08-14-financial-reconciliation-automation-schema.sql tests/reconciliation-automation-rpc.smoke.sql tests/reconciliation-automation.test.js
git commit -m "feat: add reconciliation automation schema"
```

---

### Task 3: Implement Deterministic Analysis and Ambiguity Detection

**Files:**
- Create: `supabase-migrations/2026-08-14-financial-reconciliation-automation-analysis.sql`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`
- Modify: `tests/reconciliation-automation.test.js`

**Interfaces:**
- Consumes: Task 2 catalog/config/run/proposal tables and existing source-rule operators.
- Produces:
  - `financial_reconciliation_match_normalize(text)`.
  - `financial_reconciliation_match_compact(text)`.
  - `financial_reconciliation_automatic_build_combinations(jsonb,jsonb,jsonb,numeric,integer)`.
  - `financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer)`.
  - `populate_financial_reconciliation_automatic_run(uuid)`.
  - `create_financial_reconciliation_automatic_analysis(text[],text,text,uuid)`.
  - `claim_financial_reconciliation_automatic_schedule(timestamptz,text)`.
  - `get_financial_reconciliation_automatic_run(uuid)`.

- [ ] **Step 1: Add failing analysis migration assertions and SQL fixtures**

Require the exact function signatures, thresholds, oldest-first order, integer-cent comparison, candidate-limit reason, overlap ambiguity, and service-role grants. Add smoke fixtures for:

- document-number containment;
- description score at both sides of `0.60`;
- supplier word score at both sides of `0.70`;
- blank identity fields;
- dates exactly 7 and 8 days away;
- differences exactly at and above tolerance;
- one-to-one and one-to-many sums;
- a pure combination fixture containing CGD Bank Statement and CGD Credit Card candidates with independent operators;
- two valid combinations for one base;
- one bank row reused by two bases;
- thirteen identity-qualified candidates causing `candidate_limit`.
- two UTC timestamps around Lisbon daylight-saving transitions that map to the same configured local slot/date, proving only one scheduled claim.

- [ ] **Step 2: Run focused Node test to verify RED**

Run: `node --test --test-isolation=none tests/reconciliation-automation.test.js`

Expected: FAIL because the analysis migration is absent.

- [ ] **Step 3: Implement deterministic text helpers**

`financial_reconciliation_match_normalize` must lower-case, unaccent, replace punctuation with spaces, collapse whitespace, split tokens, retain numeric tokens and alphabetic tokens of length at least three, and rejoin in source order. `financial_reconciliation_match_compact` removes all remaining non-alphanumeric characters. Mark wrappers that call `unaccent` as `stable` and `strict` rather than falsely declaring them immutable; set a safe `search_path` on every security-definer function.

Identity evidence for every bank candidate is:

```json
{
  "documentNumber": { "matched": true, "normalized": "FT2026001234" },
  "description": { "matched": false, "score": 0.42, "threshold": 0.60 },
  "supplier": { "matched": true, "score": 0.81, "threshold": 0.70 }
}
```

Use `similarity(normalized_document_description, normalized_bank_description)` and `word_similarity(normalized_supplier_name, normalized_bank_description)`. A candidate qualifies when document containment passes or either score reaches its inclusive threshold.

- [ ] **Step 4: Implement bounded candidate combinations**

For rule key/version 1, select unlocked `financial_documents` using `document_date`, `doc_number`, `description`, `supplier_name`, and `amount` with `fat='S'`. Select unlocked `import_cgd_extrato_ordem` using `data`, `descritivo`, and `montante`. Both dates must be on/after the eligibility floor, `montante` must be non-null, and unlocked means no row exists in `financial_reconciliation_items` for that source type/ID. Join by inclusive date window and identity. If more than 12 identity-qualified bank rows exist for one document, return one `candidate_limit` ambiguous result and generate no subsets.

Put subset construction in `financial_reconciliation_automatic_build_combinations`. It accepts a base snapshot, heterogeneous destination candidate snapshots, one operator per destination source, tolerance, and maximum group size. Use a recursive CTE that creates stable source-type/ID-ordered subsets. Sum integer cents and retain only subsets inside tolerance. Generate a SHA-256 signature over sorted `{sourceType,sourceId,amountCents,sourceDate}` entries. The smoke fixture must prove one group can contain Bank Statement and Credit Card candidates with independent operators even though the first production rule supplies only Bank Statement candidates.

For 12 or fewer version-1 candidates, call the generic builder with maximum group size 4 and the current `financial_documents -> import_cgd_extrato_ordem` source-rule operator captured into the analysis snapshot.

- [ ] **Step 5: Persist analysis runs and proposals**

`populate_financial_reconciliation_automatic_run` locks an unanalysed run, reads its snapshotted ordered rules, inserts the proposals below, sets `analysis_completed_at`, and returns the public run. Repeated calls return the existing analysed run without duplicating proposals.

`create_financial_reconciliation_automatic_analysis` treats its second parameter as mode: `manual_rule` requires every requested rule to allow manual execution; `manual_batch` requires every requested rule to be enabled for the scheduled batch; any other mode is rejected. It writes `trigger='manual'` with scope `rule` or `batch`, snapshots definitions/configs/operators in priority order, inserts one run idempotently by `(actor,client_request_id)`, and calls `populate_financial_reconciliation_automatic_run`.

The population step inserts:

- status `proposed` when exactly one complete combination exists;
- status `ambiguous` with all candidate groups when zero uniqueness exists because two or more combinations qualify;
- status `ambiguous`/reason `cross_base_overlap` when a source item occurs in proposals for multiple bases;
- status `ambiguous`/reason `candidate_limit` when the fixed candidate bound is exceeded.

Return `get_financial_reconciliation_automatic_run(run_id)` with definitions, proposals, evidence, counters, and no internal diagnostic fields.

- [ ] **Step 6: Implement once-per-Lisbon-day claiming**

`claim_financial_reconciliation_automatic_schedule(p_now,p_actor)` locks the singleton schedule, computes Lisbon local date/time, returns `claimed=false` before the configured time or when disabled, resumes the existing unfinished scheduled run, and otherwise snapshots all enabled batch rules/operators in priority order and inserts exactly one unanalysed run with `trigger='scheduled'`, scope `batch`, and `scheduled_slot = to_char(lisbon_local_date,'YYYY-MM-DD')`. A unique constraint is the final duplicate guard. The scheduler calls `populate_financial_reconciliation_automatic_run` after claiming so a retried heartbeat can safely resume either before or after analysis.

- [ ] **Step 7: Run smoke and Node verification**

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none tests/*.test.js
```

Expected: all available checks pass; document any unavailable database check explicitly.

- [ ] **Step 8: Commit**

```powershell
git add -- supabase-migrations/2026-08-14-financial-reconciliation-automation-analysis.sql tests/reconciliation-automation-rpc.smoke.sql tests/reconciliation-automation.test.js
git commit -m "feat: analyze automatic reconciliation rules"
```

---

### Task 4: Execute Proposals Atomically with Provenance

**Files:**
- Create: `supabase-migrations/2026-08-14-financial-reconciliation-automation-execution.sql`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`
- Modify: `tests/reconciliation-automation.test.js`
- Modify: `tests/reconciliation-rpc.smoke.sql`

**Interfaces:**
- Consumes: Task 3 proposal rows and existing `financial_reconciliation_action(text,text,uuid,text,uuid,text)`.
- Produces:
  - `execute_financial_reconciliation_automatic_proposal(uuid,text)`.
  - `finish_financial_reconciliation_automatic_run(uuid)`.
  - Workspace/history provenance fields `origin`, `automaticTrigger`, `automaticRuleKey`, `automaticRuleVersion`, and `automaticRunId`.

- [ ] **Step 1: Add failing execution and compatibility fixtures**

Extend automation smoke coverage to prove:

- a proposed group completes and locks every item;
- `origin='automatic'` and trigger/rule/run/proposal links are populated;
- non-zero tolerated completion uses `force_complete` with the generated comment;
- zero difference completes normally but retains structured audit metadata;
- stale amount/date/lock/rule evidence marks the proposal `stale` and creates no reconciliation;
- a second execution returns the same reconciliation and does not duplicate audit/items;
- one failed proposal does not prevent a later proposal executed in a separate RPC transaction;
- reopening/deleting an automatic reconciliation preserves provenance and uses existing lifecycle behavior;
- existing manual start/add/complete tests still produce `origin='user'`.

- [ ] **Step 2: Run focused tests to verify RED**

Run the Node source-contract test and the SQL smoke file. Expected: failures for the absent execution migration/function.

- [ ] **Step 3: Implement one-proposal transactional execution**

`execute_financial_reconciliation_automatic_proposal` must:

1. Lock the proposal row and its run.
2. Return the existing reconciliation for `status='completed'`.
3. Reject ambiguous/deselected/failed proposals.
4. Re-run the same rule/version/config/operator analysis for that base and require the signature, item snapshots, identity evidence, and tolerance to match exactly.
5. Call the existing action RPC internally to start with the base item, add destination items in stable source/date/ID order, and complete or force-complete.
6. Before completion, update the new reconciliation provenance columns.
7. Generate this stable non-zero comment:

```text
Automatically completed by rule Financial Documents to CGD Bank Statement v1; difference €0.35 within allowed tolerance €1.00; trigger Scheduled; batch <run-uuid>.
```

8. Add an `automatic_complete` audit row whose metadata contains rule/config/operator snapshots, identity evidence, proposal signature, trigger, run ID, and tolerance.
9. Mark the proposal completed with reconciliation ID and timestamps.

Do not catch errors inside this function in a way that commits partial lifecycle mutations. The API supplies per-proposal isolation by calling this RPC once for each proposal.

- [ ] **Step 4: Implement run finalization and workspace enrichment**

`finish_financial_reconciliation_automatic_run` aggregates proposal statuses into counters, sets `completed`, `partial`, or `failed`, and stores finish time. Replace/enrich the current workspace function using the same `pg_get_functiondef` count-validated migration pattern already used by `2026-08-12-financial-reconciliation-history-source-summary.sql`. Return camelCase provenance on current and history records without changing existing source summaries.

- [ ] **Step 5: Run all database and Node checks**

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-rpc.smoke.sql
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
node --test --test-isolation=none tests/*.test.js
```

- [ ] **Step 6: Commit**

```powershell
git add -- supabase-migrations/2026-08-14-financial-reconciliation-automation-execution.sql tests/reconciliation-automation-rpc.smoke.sql tests/reconciliation-automation.test.js tests/reconciliation-rpc.smoke.sql
git commit -m "feat: execute automatic reconciliations atomically"
```

---

### Task 5: Add Settings and Manual Automation APIs

**Files:**
- Create: `api/reconciliation-automation-settings.js`
- Create: `api/reconciliation-automation.js`
- Modify: `tests/reconciliation-automation.test.js`

**Interfaces:**
- Consumes: Tasks 1–4 normalizers and RPCs.
- Produces:
  - `GET/PUT /api/reconciliation-automation-settings`.
  - `GET/POST /api/reconciliation-automation`.

- [ ] **Step 1: Write failing mocked-handler tests**

Use the existing require-cache replacement pattern from `tests/reconciliation-settings.test.js`. Assert exact authorization and RPC calls:

```js
assert.deepEqual(calls[0], {
  resource: "rpc/create_financial_reconciliation_automatic_analysis",
  options: { method: "POST", body: {
    p_rule_keys: [AUTOMATIC_RULE_KEY],
    p_mode: "manual_rule",
    p_actor: "user@example.com",
    p_client_request_id: clientRequestId,
  } },
});
```

Cover Settings GET/PUT, single-rule Analyze, administrator-only Run batch now, Execute selected with one execution RPC per proposal, finalization after the loop, partial failures retained in the response, run-detail GET, wrong methods, and safe mapped errors.

- [ ] **Step 2: Run focused tests to verify RED**

Run: `node --test --test-isolation=none tests/reconciliation-automation.test.js`

Expected: FAIL because both handlers are absent.

- [ ] **Step 3: Implement the settings handler**

Require `settings/financial-reconciliation` for every method. GET calls `rpc/get_financial_reconciliation_automation_settings`. PUT parses/normalizes the complete schedule/rule payload, passes it through `toAutomationSettingsRpcPayload(settings, actor)`, and calls `rpc/replace_financial_reconciliation_automation_settings`. Support only GET/PUT and set `Allow` on 405.

- [ ] **Step 4: Implement the manual automation handler**

- GET requires `app/financial-reconciliation`, validates `run_id`, and calls `rpc/get_financial_reconciliation_automatic_run`.
- `analyze_rule` requires app access, sends mode `manual_rule`, and permits only a rule configured for manual execution.
- `analyze_batch` requires `settings/financial-reconciliation`, sends mode `manual_batch`, and analyzes all enabled batch rules without executing them.
- `execute_selected` requires app access, validates at most 100 IDs, calls `rpc/execute_financial_reconciliation_automatic_proposal` separately for each selected ID, then calls `rpc/finish_financial_reconciliation_automatic_run` and returns the refreshed run plus per-proposal outcomes.
- Use `Promise` sequencing, not parallel execution, so overlapping proposals encounter deterministic locks and results.

- [ ] **Step 5: Verify handlers and full suite**

```powershell
node --check api/reconciliation-automation-settings.js
node --check api/reconciliation-automation.js
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none tests/*.test.js
```

- [ ] **Step 6: Commit**

```powershell
git add -- api/reconciliation-automation-settings.js api/reconciliation-automation.js tests/reconciliation-automation.test.js
git commit -m "feat: expose reconciliation automation APIs"
```

---

### Task 6: Build the Automatic Reconciliation Settings Tab

**Files:**
- Modify: `index.html`
- Modify: `styles.css`
- Modify: `app-main.js`
- Create: `tests/reconciliation-automation-ui.test.js`
- Modify: `tests/reconciliation-density.test.js`

**Interfaces:**
- Consumes: Task 5 settings API response `{ schedule, rules, lastScheduledRun }`.
- Produces: settings state, rendering, atomic save payload, priority reorder, and Run batch now navigation/analysis.

- [ ] **Step 1: Write failing structural and behavior tests**

Require IDs for:

```text
financial-reconciliation-settings-source-tab
financial-reconciliation-settings-automatic-tab
financial-reconciliation-settings-source-panel
financial-reconciliation-settings-automatic-panel
financial-reconciliation-automation-schedule-enabled
financial-reconciliation-automation-schedule-time
financial-reconciliation-automation-rules
financial-reconciliation-automation-save
financial-reconciliation-automation-run-batch-now
```

Extract and execute actual renderer/serializer helpers from `app-main.js`. Prove definition text and thresholds render read-only, only approved config fields enter PUT, reorder produces unique consecutive priorities, invalid local values disable Save, and Run batch now switches to Reconciliation and performs analysis rather than execution.

- [ ] **Step 2: Run UI tests to verify RED**

Run: `node --test --test-isolation=none tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js`

Expected: FAIL for absent markup/helpers.

- [ ] **Step 3: Add settings subtabs and automatic panel**

Keep current Source rules controls inside the source panel. Add an automatic panel containing schedule enabled/time fields, fixed `Europe/Lisbon` label, last/next execution text, ordered rule cards/table, Move up/down controls, enabled/manual/scheduled checkboxes, tolerance/day inputs, expandable read-only definition, version, Save, and Run batch now.

- [ ] **Step 4: Implement settings state and rendering**

Add state:

```js
reconciliationAutomationSettings: {
  loaded: false,
  loading: false,
  activeTab: "source-rules",
  schedule: { enabled: false, timeOfDay: "02:00", timeZone: "Europe/Lisbon" },
  rules: [],
  lastScheduledRun: null,
}
```

Implement `loadReconciliationAutomationSettings`, `renderReconciliationAutomationSettings`, `reconciliationAutomationSettingsPayload`, `saveReconciliationAutomationSettings`, and `moveReconciliationAutomationRule`. Escape every server-provided label/explanation. Inputs mutate only the local draft until one atomic PUT succeeds.

- [ ] **Step 5: Implement Run batch now as analysis**

The action POSTs `{ action:"analyze_batch", clientRequestId: crypto.randomUUID() }`, stores the returned run in automation workbench state, navigates to the Reconciliation page, and renders proposals. It must never call Execute automatically.

- [ ] **Step 6: Verify UI and responsive CSS**

Run Node UI/full suites, start the local server, and use an authenticated browser fixture or mocks to check desktop and narrow widths. Confirm long read-only logic wraps, controls remain reachable, source rules are unchanged, and keyboard focus/order works.

- [ ] **Step 7: Commit**

```powershell
git add -- index.html styles.css app-main.js tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
git commit -m "feat: configure automatic reconciliation rules"
```

---

### Task 7: Add Workbench Analysis, Selection, Results, and Origin Badges

**Files:**
- Modify: `index.html`
- Modify: `styles.css`
- Modify: `app-main.js`
- Modify: `tests/reconciliation-automation-ui.test.js`
- Modify: `tests/reconciliation-density.test.js`

**Interfaces:**
- Consumes: Task 5 manual API and workspace provenance from Task 4.
- Produces: rule Analyze controls, proposal selection/execution, result summaries, and provenance rendering.

- [ ] **Step 1: Write failing end-user behavior tests**

Extract actual production helpers and assert:

- only enabled/manual rules show Analyze;
- Analyze sends one rule key plus a fresh UUID and never sends proposal IDs;
- proposed rows start selected; ambiguous rows are disabled and display all candidate groups/reason;
- select-all/clear-all affect only executable proposals;
- Execute selected sends exactly checked proposal IDs once and disables while pending;
- zero selected prevents execution;
- returned `completed/stale/failed` results render separately;
- origin badges render `User`, `Automatic · Manual`, and `Automatic · Scheduled` in Current and History;
- missing legacy provenance defaults to `User`.

- [ ] **Step 2: Run UI tests to verify RED**

Run: `node --test --test-isolation=none tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js`

- [ ] **Step 3: Add workbench markup and state**

Add an Automatic reconciliation section after the workbench filters and before history, with rule cards, Analyze buttons, analysis status, proposal list, selection controls, Execute selected, and result summary. Add state:

```js
automation: {
  rules: [],
  run: null,
  selectedProposalIds: new Set(),
  pendingAction: "",
  loaded: false,
}
```

Because state serialization cannot preserve `Set`, do not place it in persisted/local storage.

- [ ] **Step 4: Render auditable proposals**

Each proposal group shows base/destination source, date, description, supplier when present, amount, applied operator, passed identity evidence with scores/thresholds, difference/tolerance, rule/version, and safe ambiguity/stale reason. Reuse `financialReconciliationSourceLabel`, `formatMoney`, `formatDateOnly`, and `escape`.

- [ ] **Step 5: Implement Analyze and Execute selected flows**

Analyze resets prior selection, POSTs `analyze_rule`, stores the run, and selects only status `proposed`. Execute copies the current selected IDs, disables controls, POSTs `execute_selected`, replaces the run with the refreshed response, refreshes the main reconciliation workspace/history, and shows summary/toasts. Repeated clicks while pending dispatch only once.

- [ ] **Step 6: Add provenance badges without regressing history**

Create `financialReconciliationOriginPresentation(record)` returning `{ key, label, className }`. Render it beside the status in Current and as a new Origin column in History. Update empty-row colspan and preserve the existing consolidated Source column and Open action.

- [ ] **Step 7: Verify behavior and browser layout**

Run focused/full Node suites and authenticated browser checks for long evidence, multi-record proposals, deselection, partial failure, and narrow widths. Confirm analysis creates no locks by refreshing Eligible records before execution.

- [ ] **Step 8: Commit**

```powershell
git add -- index.html styles.css app-main.js tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
git commit -m "feat: review and execute automatic matches"
```

---

### Task 8: Add the Protected Daily Scheduler

**Files:**
- Create: `api/reconciliation-automation-cron.js`
- Modify: `vercel.json`
- Modify: `tests/reconciliation-automation.test.js`
- Modify: `README.md`

**Interfaces:**
- Consumes: Task 1 cron authentication and Tasks 3–4 claim/analyze/execute/finalize RPCs.
- Produces: `GET/POST /api/reconciliation-automation-cron` heartbeat and deployment configuration.

- [ ] **Step 1: Write failing scheduler handler tests**

Mock time, auth, and RPCs. Cover unauthorized 401, disabled/not-due 200, first daily claim, duplicate heartbeat resume, 25-proposal limit, sequential proposal calls, failure continuation, run finalization when no pending work remains, and no second slot across Lisbon DST boundaries.

- [ ] **Step 2: Run focused tests to verify RED**

Run: `node --test --test-isolation=none tests/reconciliation-automation.test.js`

Expected: FAIL because the cron handler is absent.

- [ ] **Step 3: Implement the protected heartbeat**

Accept only requests passing `isCronRequest(req, process.env.CRON_SECRET)`. Call `claim_financial_reconciliation_automatic_schedule(now,"system:reconciliation")`. If not claimed, return its reason. For a claimed/resumed run whose `analysisCompletedAt` is empty, call `rpc/populate_financial_reconciliation_automatic_run` with its run ID. Load pending proposed IDs in stable rule-priority/base-date/ID order, execute at most 25 sequentially, then finalize only when no pending proposals remain. Return counts and `hasMore`; never return internal diagnostic metadata.

- [ ] **Step 4: Configure Vercel and deployment documentation**

Add:

```json
{ "path": "/api/reconciliation-automation-cron", "schedule": "* * * * *" }
```

Document that the heartbeat is every minute but the database claims only the configured once-daily Lisbon slot. Document `CRON_SECRET`, the three migration files in order, disabled-by-default rollout, manual validation before scheduled enablement, and how to inspect the last batch in Settings.

- [ ] **Step 5: Verify and commit**

```powershell
node --check api/reconciliation-automation-cron.js
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none tests/*.test.js
git add -- api/reconciliation-automation-cron.js vercel.json tests/reconciliation-automation.test.js README.md
git commit -m "feat: schedule automatic reconciliation batches"
```

---

### Task 9: Run Full Verification and Stage the Rollout

**Files:**
- Modify only if verification exposes a defect: files already named in Tasks 1–8.

**Interfaces:**
- Consumes: the complete feature.
- Produces: evidence that the implementation is safe to merge and a precise list of external rollout gates.

- [ ] **Step 1: Apply migrations to a disposable Supabase environment in order**

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f supabase-migrations/2026-08-14-financial-reconciliation-automation-schema.sql
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f supabase-migrations/2026-08-14-financial-reconciliation-automation-analysis.sql
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f supabase-migrations/2026-08-14-financial-reconciliation-automation-execution.sql
```

Reapply all three once to prove idempotency without resetting administrator configuration.

- [ ] **Step 2: Run authoritative SQL smoke contracts**

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-rpc.smoke.sql
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

Expected: both exit `0` and roll back their fixtures.

- [ ] **Step 3: Run syntax and full Node verification**

```powershell
node --check api/_reconciliation-automation.js
node --check api/reconciliation-automation-settings.js
node --check api/reconciliation-automation.js
node --check api/reconciliation-automation-cron.js
node --check app-main.js
node --test --test-isolation=none tests/*.test.js
git diff --check
```

Expected: all commands exit `0`, zero failed tests, and no whitespace errors.

- [ ] **Step 4: Perform authenticated browser verification**

Verify:

1. Existing Source rules still load/save.
2. Automatic rule logic is visible and noneditable.
3. Tolerance/days/modes/priority and Lisbon schedule save atomically.
4. Analyze creates proposals but leaves records unlocked.
5. Evidence, scores, thresholds, multi-record totals, and ambiguity reasons are accurate.
6. Deselect and Execute selected affect only chosen groups.
7. Origin badges and audit metadata appear in Current and History.
8. Run batch now analyzes but does not auto-execute.
9. A protected scheduled heartbeat claims once, resumes, and reports results.
10. Desktop and mobile layouts wrap without overlap or inaccessible controls.

- [ ] **Step 5: Exercise rollout safety**

Keep the seeded rule disabled. Enable manual only in non-production, compare proposals with known true/false samples, and record threshold-boundary results. Enable scheduled mode only after acceptance. Confirm disabling schedule/rule prevents future runs and leaves completed reconciliations unchanged.

- [ ] **Step 6: Review the final diff and route defects back to their owning task**

Run `git status --short`, `git diff --check`, and `git diff`. If verification exposes a defect, reopen the task that owns that file, repeat its specified RED/GREEN cycle, use that task's explicit `git add -- ...` command, and commit with `fix: harden reconciliation automation verification`. If no files changed, do not create an empty commit.
