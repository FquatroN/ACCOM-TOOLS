# FDM Bank and Adyen Automatic Reconciliation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add two disabled-by-default managed automatic reconciliation rules: a bounded exact-cents one-Bank-to-1–10-FDM combination rule and a closed-calendar-month Adyen aggregation rule.

**Architecture:** Extend the existing explicit managed-strategy registry from five to seven key/version pairs. One forward Supabase migration installs immutable definitions, disabled configurations, literal analysis and execution dispatch, bounded combination search, calendar-month aggregation, seven-child scheduling, and reapply-safe database guards. Existing authorized Node endpoints, proposal memberships, three-column review UI, audit/history, and shared daily scheduler remain the only public lifecycle.

**Tech Stack:** PostgreSQL/Supabase RPC and PostgREST, Vercel Node functions, browser JavaScript, HTML/CSS, Node's built-in test runner, transactional PostgreSQL smoke tests.

**Spec:** `docs/superpowers/specs/2026-08-23-fdm-bank-adyen-automatic-reconciliation-design.md`

## Global Constraints

- Add `fdm_bank_transfer_cgd_bank_statement_combination` version `1` and `cgd_bank_statement_fdm_adyen_monthly_payments` version `1`; preserve the existing five managed rules unchanged.
- Seed both new configurations with `enabled = false`, `allow_manual_execution = false`, and `include_in_scheduled_batch = false`. Deployment never activates a rule.
- Bank Reservation Payments uses exact `import_fdm_accounts.account = 'Bank Transfer'`, one Bank Statement destination, 1–10 FDM source memberships, opposite signs, exact integer cents, and a fixed zero difference allowance.
- Bank Reservation Payments uses a configurable inclusive `0..90` day window, default `3`, measured between every selected FDM date and the Bank anchor date.
- Bound Bank combination analysis at 60 eligible FDM candidates per Bank, 250,000 evaluated states, and 12 persisted evidence groups. Every ceiling produces non-executable `candidate_limit`; it never produces a false no-match.
- Adyen uses Bank Description containing `Adyen` case-insensitively and exact FDM Account `Adyen`. It groups every eligible unlocked record from both sides in one closed calendar month and requires both sides to be nonempty.
- Adyen difference allowance defaults to `2000.00` and remains administrator-editable; its managed day value is fixed/read-only `31` as the calendar-month marker.
- Use the configured directional source rule for each business direction. Snapshot its identity/operator and make operator or source-rule drift stale.
- Use exact numeric/integer-cent arithmetic; never JavaScript floating-point comparison or SQL approximate numeric matching.
- Use explicit key/version allowlists and literal strategy dispatch. Do not introduce dynamic SQL, editable predicates, or a general rule language.
- Reuse immutable runs, proposals, proposal memberships, reconciliation items, source locks, audit events, history summaries, and the shared schedule.
- Keep manual analysis to one selected rule and one unfinished run per actor. Keep scheduled execution sequential: one child completes or fails before the next child is claimed.
- Lock proposal, run, definition/config/source-rule rows, and live members in deterministic global order before execution revalidation.
- Treat changed, deleted, already consumed, newly overlapping, or eligibility-breaking members as stale with no reconciliation writes.
- All SECURITY DEFINER functions use `SET search_path = public, pg_temp`, schema-qualified objects, sanitized errors, and explicit ACLs. Private helpers deny `public`, `anon`, and `authenticated`; public mutation RPCs grant only `service_role`.
- The migration is safe to apply and reapply, compares same-named indexes/constraints/definitions exactly, preserves historical snapshots, and contains no transaction-control statements because the smoke suite owns its outer transaction.
- Add behavior-bearing Node and transactional SQL tests before production changes. Source-text assertions may guard file presence or immutable literal boundaries only; they never substitute for database behavior.
- Do not claim PostgreSQL, authenticated browser, or protected heartbeat verification unless each gate actually ran.

---

### Task 1: Seven-rule application registry and rule-specific settings validation

**Files:**
- Modify: `api/_reconciliation-automation.js`
- Modify: `tests/reconciliation-automation.test.js`

**Interfaces:**
- Consumes: `AUTOMATIC_RULE_VERSIONS`, `AUTOMATIC_RULE_DISPLAY_NAMES`, `normalizeAutomationSettingsPayload`, `normalizeRpcSettings`, `normalizeAnalyzePayload`, and existing amount-only/monthly predicates.
- Produces:
  - `BANK_RESERVATION_RULE_KEY = "fdm_bank_transfer_cgd_bank_statement_combination"`, version `1`;
  - `ADYEN_MONTHLY_RULE_KEY = "cgd_bank_statement_fdm_adyen_monthly_payments"`, version `1`;
  - seven-entry immutable rule registry and display-name map;
  - `isCombinationAggregateRule(ruleKey)` and expanded `isMonthlyAggregateRule(ruleKey)`;
  - fixed-zero validation for Bank Reservation and fixed-31 validation for Adyen;
  - complete camel-case mapping for membership/group summary fields.

- [ ] **Step 1: Add failing behavior tests for the seven explicit rule pairs**

In `tests/reconciliation-automation.test.js`, import the two new constants and add a production-shaped seven-rule settings fixture. Assert behavior equivalent to:

```js
test("managed settings accept exactly the seven supported rule versions", () => {
  const result = normalizeAutomationSettingsPayload(sevenRuleSettings());
  assert.deepEqual(result.rules.map(({ ruleKey, ruleVersion }) => [ruleKey, ruleVersion]), [
    [BANK_STATEMENT_RULE_KEY, 2],
    [CREDIT_CARD_RULE_KEY, 1],
    [BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY, 1],
    [CREDIT_CARD_AMOUNT_ONLY_RULE_KEY, 1],
    [MONTHLY_INCOME_RULE_KEY, 2],
    [BANK_RESERVATION_RULE_KEY, 1],
    [ADYEN_MONTHLY_RULE_KEY, 1],
  ]);
});

test("Bank Reservation fixes zero tolerance while Adyen fixes calendar-month mode", () => {
  assert.throws(
    () => normalizeAutomationSettingsPayload(sevenRuleSettings({ bankReservationDifference: "0.01" })),
    /zero difference/i,
  );
  assert.throws(
    () => normalizeAutomationSettingsPayload(sevenRuleSettings({ adyenMaxDays: 30 })),
    /calendar.month/i,
  );
});
```

Also prove that Bank Reservation accepts day values `0` and `90` but rejects `91`; Adyen accepts non-negative allowance including `0` and `2000.00`; unsupported version `2`, near-name keys, prototype-backed objects, duplicate priorities, and multi-key manual analyze payloads fail before RPC invocation.

- [ ] **Step 2: Run focused tests and capture RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: FAIL only on absent new key/version mappings and rule-specific validation.

- [ ] **Step 3: Extend the explicit JavaScript registry minimally**

In `api/_reconciliation-automation.js`, add explicit constants and predicates:

```js
const BANK_RESERVATION_RULE_KEY = "fdm_bank_transfer_cgd_bank_statement_combination";
const BANK_RESERVATION_RULE_VERSION = 1;
const ADYEN_MONTHLY_RULE_KEY = "cgd_bank_statement_fdm_adyen_monthly_payments";
const ADYEN_MONTHLY_RULE_VERSION = 1;

const COMBINATION_AGGREGATE_RULE_KEYS = new Set([BANK_RESERVATION_RULE_KEY]);
const MONTHLY_AGGREGATE_RULE_KEYS = new Set([
  MONTHLY_INCOME_RULE_KEY,
  ADYEN_MONTHLY_RULE_KEY,
]);

function isCombinationAggregateRule(ruleKey) {
  return COMBINATION_AGGREGATE_RULE_KEYS.has(normalizeRuleKey(ruleKey));
}
```

Extend `AUTOMATIC_RULE_VERSIONS`, `AUTOMATIC_RULE_DISPLAY_NAMES`, exports, and public key mapping. In both editable and RPC-shaped settings normalization:

```js
if (isCombinationAggregateRule(ruleKey) && differenceAllowedCents !== 0) {
  throw inputError("Bank Reservation rules require a zero difference allowed.");
}
if (ruleKey === ADYEN_MONTHLY_RULE_KEY && maxDifferenceDays !== 31) {
  throw inputError("Adyen monthly rules require calendar-month mode.");
}
```

Do not require all seven rows at the Node mapping boundary; rollout must tolerate the old five-row database until migration application.

- [ ] **Step 4: Verify GREEN and commit**

Run:

```powershell
node --check api/_reconciliation-automation.js
node --test --test-isolation=none tests/reconciliation-automation.test.js
git diff --check
```

Expected: all commands exit `0`.

Commit:

```powershell
git add -- api/_reconciliation-automation.js tests/reconciliation-automation.test.js
git commit -m "feat: register FDM Bank and Adyen rules"
```

---

### Task 2: Immutable definitions, disabled configurations, indexes, and Settings RPC

**Files:**
- Create: `supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`
- Modify: `tests/reconciliation-settings.test.js`
- Modify: `api/reconciliation-settings.js`

**Interfaces:**
- Consumes: automatic definition/config tables, proposal memberships, directional source rules, `get_financial_reconciliation_automation_settings`, and `replace_financial_reconciliation_automation_settings`.
- Produces: two immutable definitions, two disabled configs, exact supporting indexes, seven-row atomic Settings replacement, and schema-cache reload notification.

- [ ] **Step 1: Add transactional catalog and reapply fixtures before the migration**

At the new section of `tests/reconciliation-automation-rpc.smoke.sql`, include the new migration twice inside the existing outer transaction:

```sql
\ir ../supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql
\ir ../supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql
```

Add executable assertions that both immutable definition JSON values equal their approved contracts, both configurations are disabled/manual false/scheduled false, Bank Reservation has zero/3, Adyen has 2000/31, and existing five configs retain values and relative priority. Change the new rows, reapply, and prove administrator-controlled flags/priority/allowed editable values are preserved.

Add failing Settings replacement cases for:

- missing either managed key;
- duplicate key or priority;
- unsupported version;
- Bank Reservation nonzero allowance or day `91`;
- Adyen day other than `31` or negative allowance;
- attempted mutation of definition/source types/strategy/cardinality/search limits.

Assert every failed replacement leaves schedule and all seven configuration rows byte-equivalent.

- [ ] **Step 2: Capture RED at the available layers**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-settings.test.js tests/reconciliation-automation.test.js
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

Expected: Node fails on the missing seven-rule API contract. PostgreSQL fails because the dated migration and rows do not exist. If `psql` or the URL is unavailable, record that exact gate instead of claiming RED execution.

- [ ] **Step 3: Install exact immutable definitions and disabled configs**

In the dated migration, insert definitions only when absent. Use stable JSON strategy metadata such as:

```sql
jsonb_build_object(
  'strategy', 'bounded_exact_combination',
  'sourceAccount', 'Bank Transfer',
  'maxSourceRecords', 10,
  'candidatePoolLimit', 60,
  'stateLimit', 250000,
  'evidenceGroupLimit', 12,
  'amountMode', 'signed_integer_cents',
  'dateMode', 'inclusive_days'
)
```

and:

```sql
jsonb_build_object(
  'strategy', 'closed_calendar_month',
  'bankDescriptionContains', 'Adyen',
  'fdmAccount', 'Adyen',
  'requiresBothSides', true,
  'monthMarkerDays', 31
)
```

For a same key/version already present, compare the complete immutable row and raise a deterministic exception on any mismatch. Append absent configs at deterministic unused priorities after the existing five without rewriting existing order.

- [ ] **Step 4: Add exact lookup indexes and fail-closed reapply checks**

Add only indexes justified by the approved predicates, with names and definitions pinned by `pg_get_indexdef`/`pg_index` checks. Use source columns confirmed from the existing adapters, for example the equivalent of:

```sql
create index financial_reconciliation_fdm_bank_transfer_lookup_idx
  on public.import_fdm_accounts (date, amount, id)
  where account = 'Bank Transfer' and date >= date '2026-01-01' and amount is not null;

create index financial_reconciliation_fdm_adyen_lookup_idx
  on public.import_fdm_accounts (date, id)
  include (amount)
  where account = 'Adyen' and date >= date '2026-01-01' and amount is not null;
```

Use the real date/amount column names already normalized by the current FDM adapter. Add a Bank date/amount index usable by the bounded search and a trigram/functional description index only if the current extension-safe wrapper and query planner contract justify it. A conflicting same-named object raises; reapply is a no-op.

- [ ] **Step 5: Replace the Settings RPC with seven-rule atomic validation**

Replace `replace_financial_reconciliation_automation_settings` using its existing signature, fixed search path, authorization, schedule lock order, and all-row transaction. Validate the exact tuple set:

```sql
values
  ('financial_documents_cgd_bank_statement', 2),
  ('financial_documents_cgd_credit_card', 1),
  ('financial_documents_cgd_bank_statement_amount_only', 1),
  ('financial_documents_cgd_credit_card_amount_only', 1),
  ('cgd_bank_statement_fdm_credit_card_monthly_income', 2),
  ('fdm_bank_transfer_cgd_bank_statement_combination', 1),
  ('cgd_bank_statement_fdm_adyen_monthly_payments', 1)
```

Keep logic JSON read-only. Bank Reservation accepts only allowance zero and days 0–90. Adyen accepts non-negative allowance and only day 31. Revoke execute from broad roles and grant the public RPC only to `service_role`.

- [ ] **Step 6: Extend the source-rule mutation guard**

In `api/reconciliation-settings.js`, prevent removal or operator mutation of the managed FDM→Bank and Bank→FDM directional pairs when required by installed automatic definitions. Add actual handler tests in `tests/reconciliation-settings.test.js` proving invalid requests return `400` before RPC, while unrelated source-rule edits still call the RPC.

- [ ] **Step 7: Verify GREEN and commit**

Run:

```powershell
node --check api/reconciliation-settings.js
node --test --test-isolation=none tests/reconciliation-settings.test.js tests/reconciliation-automation.test.js
node --test --test-isolation=none
git diff --check
```

Run the SQL smoke when available. Commit only the scoped files:

```powershell
git add -- supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql tests/reconciliation-automation-rpc.smoke.sql tests/reconciliation-settings.test.js api/reconciliation-settings.js
git commit -m "feat: configure FDM Bank and Adyen rules"
```

---

### Task 3: Bounded Bank Reservation combination analysis

**Files:**
- Modify: `supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`
- Modify: `tests/reconciliation-automation.test.js`

**Interfaces:**
- Consumes: unlocked Bank/FDM workspace adapters, run snapshot, proposal/membership schema, `continue_financial_reconciliation_automatic_analysis`.
- Produces:
  - stable Bank-anchor count/page helpers;
  - deterministic bounded exact-cents combination search;
  - Bank Reservation proposal/member snapshots;
  - literal continuation dispatch and correct run counters.

- [ ] **Step 1: Write SQL behavior fixtures for every classification**

Create isolated fixture dates/amounts after the eligibility floor. Exercise:

- unique 1-, 2-, and 10-FDM exact zero groups;
- same-sign, one-cent mismatch, Account near-match, null amount/date, and day-outside exclusions;
- inclusive day boundary at configured days;
- 11-member sum rejection;
- two qualifying combinations as `multiple_qualifying_combinations`;
- candidate pool 61, state ceiling, and thirteenth qualifying group as `candidate_limit` with at most 12 evidence groups;
- shared Bank and shared FDM overlaps as `overlapping_records`;
- no qualifying group omitted from visible proposals but counted in skipped/no-match accounting;
- stable continuation across multiple 25-Bank pages and retry idempotency.

Assert membership roles: all selected FDM rows are `source`, the Bank row is `destination`, ordinals are stable, and canonical `base_source_id` equals the first selected FDM `(date,id)`.

- [ ] **Step 2: Capture RED**

Run the transactional smoke. Expected: missing Bank Reservation continuation/combination behavior. If PostgreSQL is unavailable, run the focused Node contract and record SQL as the external behavior gate.

- [ ] **Step 3: Implement stable Bank anchor count/page helpers**

Add private helpers with fixed signatures:

```sql
public.financial_reconciliation_automatic_bank_reservation_count()
  returns bigint

public.financial_reconciliation_automatic_bank_reservation_page(
  p_after_date date,
  p_after_id uuid,
  p_limit integer
) returns table(bank_id uuid, bank_date date, bank_amount numeric)
```

They return eligible unlocked Bank Statement anchors in `(date,id)` order and enforce the 2026-01-01 floor, non-null date/amount, and source-lock exclusions.

- [ ] **Step 4: Implement a deterministic bounded exact-cents search**

Add a private helper returning classification plus bounded evidence:

```sql
public.financial_reconciliation_automatic_bank_reservation_groups(
  p_bank_id uuid,
  p_max_difference_days integer,
  p_candidate_pool_limit integer,
  p_state_limit integer,
  p_evidence_limit integer
) returns jsonb
```

Load at most 61 candidates to detect the 60-record ceiling. Convert signed amounts to integer cents. Reject wrong signs before recursion. Enumerate candidates in `(date,id)` order with a deterministic depth-first or dynamic-programming state machine, maximum depth 10, and an explicit evaluated-state counter. Stop only with a classified limit, never a silent empty result.

Return JSON shaped like:

```json
{
  "classification": "proposed",
  "reason": "unique_qualifying_combination",
  "evaluatedStates": 84,
  "candidateGroups": [
    {"fdmIds": ["..."], "fdmTotalCents": 1250, "bankAmountCents": -1250, "equationCents": 0}
  ]
}
```

- [ ] **Step 5: Persist immutable proposals and memberships**

Implement:

```sql
public.financial_reconciliation_continue_automatic_bank_reservation(
  p_run_id uuid,
  p_actor text
) returns jsonb
```

Lock the run, validate its exact definition/config/source-rule snapshot, fetch one page, search each anchor, and insert proposal/membership evidence atomically. Snapshot member source type, ID, role, ordinal, date, amount, description, Account, and source row. After all pages, run an authoritative membership overlap pass across otherwise-proposed rows and update every affected row to ambiguous.

Keep counters monotonic and derive `analysis_total` from Bank anchors. Finalize to ready only when at least one proposed/ambiguous/stale/failed review row exists; otherwise complete the run with skipped accounting.

- [ ] **Step 6: Extend literal continuation dispatch**

Replace `continue_financial_reconciliation_automatic_analysis` without changing its signature. Dispatch explicitly:

```sql
elsif v_rule_key = 'fdm_bank_transfer_cgd_bank_statement_combination'
  and v_rule_version = 1 then
  return public.financial_reconciliation_continue_automatic_bank_reservation(
    p_run_id,
    p_actor
  );
```

Unknown tuples fail closed. No configured function name enters SQL execution.

- [ ] **Step 7: Verify analysis and commit**

Run focused Node tests, the full Node suite, SQL smoke when available, and `git diff --check`. Commit:

```powershell
git add -- supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql tests/reconciliation-automation-rpc.smoke.sql tests/reconciliation-automation.test.js
git commit -m "feat: analyze Bank Reservation combinations"
```

---

### Task 4: Closed-calendar-month Adyen analysis

**Files:**
- Modify: `supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`
- Modify: `tests/reconciliation-automation.test.js`

**Interfaces:**
- Consumes: existing monthly-income month cursor/summary patterns and proposal memberships.
- Produces: Adyen month count/page/continuation, exact monthly totals, allowance classification, and literal dispatch.

- [ ] **Step 1: Add Adyen monthly behavior fixtures first**

Cover Bank descriptions `Adyen`, `ADYEN settlement`, and a nonmatching description; FDM Account exact `Adyen` and near matches; closed previous months and excluded current month; both sides required; the 2026-01-01 floor; one month with many members; zero, nonzero-under, exact-boundary, and over-allowance differences.

Assert:

- one proposal per `YYYY-MM` grouping key;
- every eligible member is included exactly once;
- within allowance is proposed and above allowance is ambiguous with `monthly_difference_exceeded`;
- empty-side months have no review proposal;
- month cursor resumes in ascending order and retries do not duplicate proposals;
- existing POS monthly fixtures remain byte-equivalent in behavior.

- [ ] **Step 2: Capture RED**

Run the SQL smoke and focused automation tests. Expected: missing Adyen monthly strategy and fifth-to-seventh registry scheduling support.

- [ ] **Step 3: Add month count/page helpers**

Implement explicit private functions:

```sql
public.financial_reconciliation_automatic_adyen_month_count()
  returns bigint

public.financial_reconciliation_automatic_adyen_month_page(
  p_after_month date,
  p_limit integer
) returns table(calendar_month date)
```

The month set is the union of eligible Bank and FDM months, filtered to `< date_trunc('month', current_date)` and `>= date '2026-01-01'`, then filtered during proposal creation to require both sides.

- [ ] **Step 4: Implement monthly continuation**

Add:

```sql
public.financial_reconciliation_continue_automatic_adyen_monthly(
  p_run_id uuid,
  p_actor text
) returns jsonb
```

For each month, select all unlocked Bank rows via the current extension-safe case-insensitive predicate and all unlocked FDM rows with exact Account. Snapshot totals and compute the configured Bank→FDM operator result. Persist all Bank members as source-role and all FDM members as destination-role, with stable `(date,id)` ordinal order.

Classify with:

```sql
case
  when abs(v_difference) <= v_allowed_difference then 'proposed'
  else 'ambiguous'
end
```

Above-allowance reason is exactly `monthly_difference_exceeded`. Snapshot day marker 31 as calendar-month mode, not a rolling window.

- [ ] **Step 5: Extend continuation/finalization dispatch and progress totals**

Add the exact Adyen key/version branch to the generic continuation and finalization functions. `analysis_total` and `analysis_processed` use calendar months. Completed no-proposal runs retain skipped counts without visible no-match rows.

- [ ] **Step 6: Verify and commit**

Run the focused/full Node suites, transactional SQL smoke when available, and diff check. Commit:

```powershell
git add -- supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql tests/reconciliation-automation-rpc.smoke.sql tests/reconciliation-automation.test.js
git commit -m "feat: analyze monthly Adyen payments"
```

---

### Task 5: Atomic execution, stale detection, audit, and history

**Files:**
- Modify: `supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`
- Modify: `tests/reconciliation-automation.test.js`

**Interfaces:**
- Consumes: immutable proposal memberships, source locks, source-rule snapshots, reconciliation lifecycle/audit helpers, `execute_financial_reconciliation_automatic_proposal`.
- Produces: specialized Bank Reservation and Adyen execution helpers plus literal top-level dispatch.

- [ ] **Step 1: Add successful execution and immutable audit fixtures**

For Bank Reservation, execute a 10-FDM/one-Bank proposal and assert one reconciliation, 11 items, correct source totals/history summary, all locks, zero difference, normal completion, origin `automatic`, rule/run/proposal IDs, and idempotent retry.

For Adyen, execute zero and nonzero-within-allowance months. Assert zero completes normally; nonzero force-completes with a deterministic comment containing rule display name, `YYYY-MM`, actual difference, and configured allowance. Assert every member and signed total appears in history.

- [ ] **Step 2: Add stale, contention, and rollback fixtures**

Mutate each snapshotted property independently after analysis: member ID/type/date/amount/description/Account, group count, Bank/FDM eligibility, config days/allowance/priority, definition version/strategy, directional source-rule operator, source lock, deletion, reconciliation consumption, and overlap. Each case must return sanitized stale, create no reconciliation/items/locks, and preserve proposal evidence.

Add malformed integer/numeric snapshot cases, simulated post-start write failure, competing execution, and unexpected database failure. Assert nested rollback removes partial reconciliation writes while only a generic failure code/status persists.

- [ ] **Step 3: Implement deterministic live-row locking helpers**

Add allowlisted helpers that lock all members in one global order `(source_type, source_id)` and return normalized live snapshots. Use literal branches for `import_fdm_accounts` and `import_cgd_extrato_ordem`; no dynamic relation names.

- [ ] **Step 4: Implement Bank Reservation execution**

Add:

```sql
public.financial_reconciliation_execute_bank_reservation_proposal(
  p_proposal_id uuid,
  p_actor text
) returns jsonb
```

Within a rollback-capable subtransaction, revalidate exact run/config/definition/operator snapshots; exactly 1–10 FDM source memberships and one Bank destination; dates; exact account; descriptions/amounts; opposite signs; zero integer-cent equation; configured operator zero difference; locks; and no cross-proposal overlap. Create the reconciliation directly from the immutable membership set rather than inferring from canonical base ID.

- [ ] **Step 5: Implement Adyen execution**

Add:

```sql
public.financial_reconciliation_execute_adyen_monthly_proposal(
  p_proposal_id uuid,
  p_actor text
) returns jsonb
```

Revalidate a single closed month, both source types, every membership, Description/Account predicates, exact totals, operator, and allowance. Select normal or forced completion based on exact current difference. Generate the forced comment inside SQL from immutable fields so retries are deterministic.

- [ ] **Step 6: Replace literal top-level execution dispatch**

Extend `execute_financial_reconciliation_automatic_proposal` with two explicit key/version branches, retaining old branches unchanged. Validate actor/run ownership before strategy dispatch and reload/finalize the authoritative run afterward.

- [ ] **Step 7: Verify and commit**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none
git diff --check
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

Commit:

```powershell
git add -- supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql tests/reconciliation-automation-rpc.smoke.sql tests/reconciliation-automation.test.js
git commit -m "feat: execute FDM Bank and Adyen proposals"
```

---

### Task 6: Manual API, proposal membership paging, and seven-child scheduler

**Files:**
- Modify: `api/reconciliation-automation.js`
- Modify: `api/reconciliation-automation-members.js`
- Modify: `api/reconciliation-automation-cron.js`
- Modify: `api/reconciliation-automation-settings.js`
- Modify: `supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql`
- Modify: `tests/reconciliation-automation.test.js`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`

**Interfaces:**
- Consumes: existing analyze/continue/execute/finish endpoints, member paging endpoint, settings endpoint, scheduled claim RPC, and batch aggregate lifecycle.
- Produces: complete seven-rule public contracts, one-rule manual analysis, grouped-member paging, seven-child sequential batches, and strategy-specific progress.

- [ ] **Step 1: Add handler-level RED tests with production-shaped RPC responses**

Test `/api/reconciliation-automation` GET manual rules returns the new rules only when both enabled and manual-enabled. Test POST Analyze passes exactly one selected new rule key, returns the locked run, and refuses another analysis while unfinished. Test Execute Selected empty and nonempty paths remain valid for both rules.

Test `/api/reconciliation-automation-members` maps member role, ordinal, row snapshot, grouping key, source totals, description, and Account without exposing private diagnostics.

- [ ] **Step 2: Add seven-child scheduled lifecycle tests**

Build a full seven-rule scheduled response ordered by priority/key. Prove:

- one heartbeat continues or executes only the current child;
- Bank Reservation progress labels/counts are Bank anchors;
- Adyen progress labels/counts are calendar months;
- the next child is first claimable on a later heartbeat;
- one failed child does not prevent later children;
- same-slot retry and cross-midnight resume use the same batch;
- terminal batch totals include all seven children;
- malformed tuple, duplicate position, wrong rule count, or invalid progress fails closed without a second claim.

- [ ] **Step 3: Capture RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: handler and cron validators reject or omit the new strategies.

- [ ] **Step 4: Extend API mapping and lifecycle validation**

Keep handlers RPC-only. Reuse `toAutomationPublicResult` and add only missing strategy fields. In cron validation, recognize the two exact key/version pairs and validate progress invariants by strategy. A terminal analysis failure with `finishedAt` and no `analysisCompletedAt` remains a valid failed child rather than causing repeated slot errors.

- [ ] **Step 5: Replace SQL manual creation and scheduled claim allowlists**

Extend `create_financial_reconciliation_automatic_analysis` and `claim_financial_reconciliation_automatic_schedule` to seven tuples. Manual creation snapshots exactly one definition/config/source rule. Scheduled claim snapshots enabled batch configs in `priority, rule_key` order, counts Bank anchors or Adyen closed months correctly, creates only one unfinished child, and preserves oldest unfinished cross-midnight behavior.

Extend `get_financial_reconciliation_automatic_manual_rules`, `get_financial_reconciliation_automatic_run`, Settings getter, batch serializer, and aggregate refresh without changing public signatures.

- [ ] **Step 6: Verify and commit**

Run syntax checks for all four handlers, focused/full Node tests, SQL smoke when available, and diff check. Commit:

```powershell
git add -- api/reconciliation-automation.js api/reconciliation-automation-members.js api/reconciliation-automation-cron.js api/reconciliation-automation-settings.js supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql tests/reconciliation-automation.test.js tests/reconciliation-automation-rpc.smoke.sql
git commit -m "feat: orchestrate seven reconciliation rules"
```

---

### Task 7: Settings and three-column proposal review UI

**Files:**
- Modify: `app-main.js`
- Modify: `styles.css`
- Modify: `tests/reconciliation-automation-ui.test.js`
- Modify: `tests/reconciliation-density.test.js`

**Interfaces:**
- Consumes: seven-rule Settings payloads, manual rule LOV, proposal summary/memberships, member paging endpoint, existing three-column layout.
- Produces: rule-specific read-only/editable controls, grouped Bank Reservation/Adyen review rows, strategy progress labels, and accessible refresh-safe interactions.

- [ ] **Step 1: Add executable UI RED tests**

Extract and execute the actual production render/normalization helpers. Assert Settings renders:

- Bank Reservation difference as read-only `0.00`, days editable `0..90`, max FDM records `10` read-only, and logic read-only;
- Adyen difference editable, days `31` read-only, and calendar-month logic read-only;
- enable/manual/scheduled/priority controls for both rules;
- saving sends only editable fields and authoritative fixed values.

Assert the manual LOV includes enabled/manual new rules, preserves selected open-run identity if later disabled, and stays locked during restore/analyze.

- [ ] **Step 2: Add three-column proposal rendering tests**

For Bank Reservation, assert first column status/reason/difference/rule, second column every FDM source member stacked in ordinal order with group total, and third column the single Bank destination. For Adyen, assert Bank members stack in column two and FDM members stack in column three. Verify descriptions, Account/supplier-equivalent detail, signed amounts, dates, and collapsible record IDs are escaped.

Test proposed selectability, ambiguous disabled selection, deselect/reselect, Execute Selected (0), completed-run filtering, pending member paging, network error recovery, keyboard focus preservation, and desktop/narrow responsive ordering.

- [ ] **Step 3: Capture RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
```

Expected: failures on new rule classifiers, fixed-field rendering, progress units, and grouped member columns.

- [ ] **Step 4: Implement rule classifiers and Settings controls**

In `app-main.js`, extend the exact key classifiers rather than inferring behavior from display text:

```js
function isBankReservationAutomationRule(ruleKey) {
  return clean(ruleKey) === "fdm_bank_transfer_cgd_bank_statement_combination";
}

function isAdyenMonthlyAutomationRule(ruleKey) {
  return clean(ruleKey) === "cgd_bank_statement_fdm_adyen_monthly_payments";
}
```

Reuse existing form rows. Disable the fixed inputs while serializing authoritative `0.00` or `31`; do not trust disabled DOM values. Preserve input focus by patching the changed rule row instead of rerendering the entire Settings panel on each keystroke.

- [ ] **Step 5: Render grouped members in the existing three-column grid**

Normalize proposal memberships once, partition by role, sort by ordinal/date/id, and render each role into its dedicated column. Keep row separators and compact density. Use visible labels such as `FDM source 1`, `Bank destination`, `Bank source 1`, and `FDM destination 1`; do not allow a second destination to wrap under the source column.

Use `aria-describedby` for reason/summary, unique checkbox names containing the proposal identity, live status for analysis progress, and retained focus/Record ID disclosure across refreshes.

- [ ] **Step 6: Verify and commit**

Run:

```powershell
node --check app-main.js
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
node --test --test-isolation=none
git diff --check
```

If an authenticated in-app session exists, verify Automatic Reconciliation and Settings at desktop and narrow widths. Record the exact limitation otherwise.

Commit:

```powershell
git add -- app-main.js styles.css tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
git commit -m "feat: review FDM Bank and Adyen proposals"
```

---

### Task 8: Rollout documentation, complete verification, and release hold

**Files:**
- Modify: `README.md`
- Modify: `docs/superpowers/plans/2026-08-23-fdm-bank-adyen-automatic-reconciliation.md`

**Interfaces:**
- Consumes: completed Tasks 1–7 and every required local/external gate.
- Produces: exact migration order, reapply/smoke commands, disabled-by-default rollout, verification evidence, and production activation hold.

- [x] **Step 1: Document migration order and rollback-safe smoke**

Add the dated migration after the currently documented migrations. Document applying it once, reapplying it once, then running:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-rpc.smoke.sql
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

State explicitly that the migration contains no `BEGIN`/`COMMIT`, both new rules remain disabled, and no administrator flags are overwritten on reapply.

- [x] **Step 2: Run final local verification from a clean committed tree**

Run:

```powershell
node --check api/_reconciliation-automation.js
node --check api/reconciliation-automation.js
node --check api/reconciliation-automation-members.js
node --check api/reconciliation-automation-settings.js
node --check api/reconciliation-automation-cron.js
node --check app-main.js
node --test --test-isolation=none tests/reconciliation-automation.test.js tests/reconciliation-settings.test.js
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
node --test --test-isolation=none
node -e "JSON.parse(require('fs').readFileSync('vercel.json','utf8')); console.log('vercel.json valid')"
git diff --check
git status --short
```

Expected: every command exits `0`; status contains only intentionally untracked user files and the scoped documentation change before its commit.

**Recorded 2026-08-24:** on hardened head `0a1c12c5acb73f9b700c6a1d58deff0e8c233c0c`, with only the two scoped Task 8 documentation changes present, all six requested syntax checks passed; the automation/Settings focused suite passed `144/144`; the UI/density focused suite passed `123/123`; the complete Node suite passed `323/323`; `vercel.json` parsed as valid JSON; and `git diff --check` passed. The working tree status showed only those scoped documentation changes before their commit.

- [ ] **Step 3: Run authoritative PostgreSQL gates**

Apply and reapply the migration against non-production, run both SQL smokes with `ON_ERROR_STOP=1`, and record exact command output. Add a real two-session contention check proving deterministic locks do not deadlock and only one executor consumes shared members.

Production rollout remains blocked if PostgreSQL is unavailable or any smoke assertion fails.

**Not run / production hold retained:** this workstation has no `psql`, `pg_isready`, `supabase`, `docker`, or `podman` command, and both `SUPABASE_DB_URL` and `DATABASE_URL` are unset. Therefore migration apply/reapply, both SQL smokes, PostgreSQL parsing/catalog/ACL/lock validation, and real two-session contention remain external gates rather than claimed successes.

- [ ] **Step 4: Run authenticated browser and protected scheduler gates**

When a signed-in non-production session is available, verify:

1. Both rules appear disabled in Settings with correct fixed/editable fields.
2. Enabling manual Bank Reservation exposes it in the LOV; unique/ambiguous groups render correctly; Execute Selected and Execute Selected (0) finish correctly.
3. Enabling manual Adyen exposes it in the LOV; a closed month renders all members; zero and forced completion create correct history/audit.
4. Desktop and narrow layouts retain three columns/stacking semantics, focus, status announcements, and accessible selection.
5. Enabling scheduled participation snapshots seven children in configured order; protected heartbeat/retry finishes one child before advancing and resumes across midnight.

Disable both rules again after verification unless the administrator explicitly authorizes activation.

**Not run / production hold retained:** the default in-app browser connection has no open tabs (`[]`), so no authenticated non-production application session was available. No Settings/manual flow, responsive browser check, protected heartbeat, retry, or seven-child live schedule was exercised.

- [x] **Step 5: Request final independent review**

Use `superpowers:requesting-code-review` across the full implementation range. Require the reviewer to inspect exact-cents search bounds, candidate-limit behavior, membership overlap, monthly completeness, lock order, stale detection, reapply safety, ACLs, seven-child scheduling, UI selection/focus, and behavior-bearing tests. Resolve every Critical or Important finding with a new RED/GREEN cycle and repeat review.

**Recorded 2026-08-24:** the initial full-range independent review found three Important defects. Hardening commit `0a1c12c5acb73f9b700c6a1d58deff0e8c233c0c` added RED/GREEN regressions for live Bank-combination drift, non-executable ambiguity overlap, and strategy-specific public grouped-proposal validation. The repeated independent review of `5a22eb06886d59076c0d2f5220ee93a2672c44b7..0a1c12c5acb73f9b700c6a1d58deff0e8c233c0c` returned `0 Critical`, `0 Important`, and `0 Minor` findings. It did not claim any PostgreSQL, browser, or heartbeat result.

- [x] **Step 6: Commit release documentation**

```powershell
git add -- README.md docs/superpowers/plans/2026-08-23-fdm-bank-adyen-automatic-reconciliation.md
git commit -m "docs: document FDM Bank and Adyen rollout"
```

**Recorded 2026-08-24:** the scoped commit uses the required message and includes only `README.md` and this plan. It does not merge, push, publish, apply SQL, or enable a rule.

- [x] **Step 7: Keep activation as a separate administrator action**

Report separately:

- local verification result;
- PostgreSQL smoke result or exact unavailable gate;
- browser/heartbeat result or exact unavailable gate;
- both rules' final disabled/manual/scheduled flags;
- commit range ready for local merge.

Do not merge, push, publish, or enable either rule unless the user separately requests that action.

**Recorded 2026-08-24:** local verification passed, while PostgreSQL and authenticated browser/protected-heartbeat results remain unavailable exactly as recorded above. The migration's static seed configuration is disabled for both new rules' enabled/manual/scheduled flags; no live database flag state was observed or changed. Activation remains a separate administrator action and no merge, push, publication, SQL application, or rule enablement occurred.
