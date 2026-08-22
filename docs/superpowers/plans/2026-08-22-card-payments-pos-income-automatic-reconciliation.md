# Card Payments - POS - Income Automatic Reconciliation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a fifth managed automatic reconciliation rule, **Card Payments - POS - Income**, which creates one auditable proposal per closed calendar month by reconciling every unlocked CGD Bank Statement `POS VENDAS` record against every unlocked FDM `Credit Card` record in that month.

**Architecture:** Preserve the existing single-rule manual analysis and sequential scheduled-batch lifecycle. Add a literal `monthly_aggregate` branch to the existing database analysis and execution dispatchers, persist immutable proposal summaries plus immutable member snapshots, and expose members only through a bounded app-authorized paging RPC/API. The browser keeps the existing three-column proposal layout: proposal summary, collapsed Bank Statement group, and collapsed FDM group.

**Tech Stack:** PostgreSQL/Supabase migrations and RPCs, Vercel Node functions, browser JavaScript, HTML/CSS, Node's built-in test runner, transactional PostgreSQL smoke tests, authenticated browser verification, Git.

> **Approved correction (2026-08-23):** The current managed contract is version
> `2`. Its FDM destination predicate is `account = 'Credit Card' AND category IS
> DISTINCT FROM 'TransferOutToAccount'`, so an exact `TransferOutToAccount` is
> excluded and a `NULL` category remains eligible. Version `1` snapshots and
> completed history remain immutable/readable. This correction supersedes the
> version-1 destination-predicate references in the original execution steps
> below.

**Spec:** `docs/superpowers/specs/2026-08-22-card-payments-pos-income-automatic-reconciliation-design.md`

## Global Constraints

- Add exactly this managed rule/version contract:
  - key: `cgd_bank_statement_fdm_credit_card_monthly_income`
  - version: `1`
  - display name: `Card Payments - POS - Income`
  - matching mode: `monthly_aggregate`
  - base: `import_cgd_extrato_ordem`
  - destination: `import_fdm_accounts`
  - source predicate: `descritivo ilike '%POS VENDAS%'`
  - destination predicate: `account = 'Credit Card'`
  - operator: Bank Statement → FDM Accounts `-`
  - default editable tolerance: `7500.00`
  - immutable maximum difference in days: `31`
  - eligibility floor: `2026-01-01`
  - both manual execution and scheduled execution disabled by default.
- Only closed calendar months are eligible. Exclude the current month using `record_date < date_trunc('month', current_date)::date`, not a rolling-day approximation.
- A month is represented only when both sides contain at least one eligible unlocked record. Do not persist or render a proposal for a one-sided or empty month.
- Aggregate all eligible unlocked records in the month. Never truncate a proposal's membership to an arbitrary candidate limit.
- Calculate `sourceTotal - destinationTotal`. Status is `proposed` when `abs(difference) <= differenceAllowed`; otherwise it is `ambiguous` with reason `monthly_difference_exceeded`.
- Process months oldest first and use a stable month cursor. Existing four rules retain their current record/date/ID cursors and behavior.
- Use the earliest eligible Bank Statement row ordered by `(data, id)` as the technical base record, but do not make that one row look like the whole source group in the UI.
- Proposal summary and member snapshots are immutable audit evidence. Execution re-reads and locks the live rows, requires exact membership equality, and marks the proposal stale if config, operator, membership, dates, amounts, or eligibility changed.
- Complete zero-difference reconciliations normally. Force-complete non-zero differences only when they remain within the snapshotted tolerance and write the generated monthly audit comment required by the design.
- Protect the managed directional source rule `import_cgd_extrato_ordem -> import_fdm_accounts (-)` in both the Settings API and the atomic database replacement RPC.
- All browser-facing database access remains RPC-only. Do not expose tables through REST. Use literal allowlisted dispatch only; do not introduce dynamic SQL.
- Every new or replaced `SECURITY DEFINER` function must use `SET search_path = public, pg_temp`, revoke execution from `public`, `anon`, and `authenticated`, and grant only the intended function to `service_role`.
- Add one forward, reapply-safe migration after migration 11: `supabase-migrations/2026-08-22-financial-reconciliation-automation-pos-income.sql`. Do not edit historical migrations.
- Do not wrap that migration in top-level `BEGIN`/`COMMIT`; the transactional smoke imports it inside its own outer transaction and must still be able to `ROLLBACK` all fixtures.
- PostgreSQL behavior must be tested in `tests/reconciliation-automation-rpc.smoke.sql`. Do not replace behavior tests with SQL-source regex tests.
- Preserve the current clean baseline: `node --test --test-isolation=none` passes 228/228 before implementation.
- Never claim SQL smoke, authenticated browser, or protected non-production heartbeat success unless those gates actually ran.

---

### Task 1: Extend the managed public contract to five rules

**Files:**
- Modify: `api/_reconciliation-automation.js`
- Modify: `tests/reconciliation-automation.test.js`

**Interfaces:**
- Add `MONTHLY_INCOME_RULE_KEY = "cgd_bank_statement_fdm_credit_card_monthly_income"` and version `1` to `AUTOMATIC_RULE_VERSIONS`.
- Accept the existing editable fields only: `enabled`, `allowManualExecution`, `includeInScheduledBatch`, `differenceAllowed`, and `priority`.
- Require `maxDifferenceDays === 31` for this rule in client payloads and authoritative RPC results.
- Map the new run/proposal fields from snake case to camel case: `grouping_key`, `summary_snapshot`, `calendar_month`, `source_count`, `source_total`, `destination_count`, `destination_total`, `total_count`, and `members`.

- [x] **Step 1: Add failing helper and handler-contract tests**

Add tests that build a five-rule Settings response and prove:

```js
assert.equal(AUTOMATIC_RULE_VERSIONS[MONTHLY_INCOME_RULE_KEY], 1);
assert.equal(normalized.rules[4].differenceAllowed, "7500.00");
assert.equal(normalized.rules[4].maxDifferenceDays, 31);
const tampered = structuredClone(settingsPayload);
tampered.rules.find((rule) => rule.ruleKey === MONTHLY_INCOME_RULE_KEY).maxDifferenceDays = 30;
assert.throws(
  () => normalizeAutomationSettingsPayload(tampered),
  /Maximum difference in days is invalid/,
);
```

Also assert that recursive public mapping converts a monthly proposal summary and paged members without leaking snake-case or diagnostic keys.

- [x] **Step 2: Run the focused test and verify RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: only the new fifth-rule/fixed-31/public-mapping assertions fail.

- [x] **Step 3: Add the minimal allowlisted normalization**

Implement explicit constants and a helper:

```js
const MONTHLY_INCOME_RULE_KEY = "cgd_bank_statement_fdm_credit_card_monthly_income";
const MONTHLY_AGGREGATE_RULE_KEYS = new Set([MONTHLY_INCOME_RULE_KEY]);

const AUTOMATIC_RULE_VERSIONS = Object.freeze({
  [BANK_STATEMENT_RULE_KEY]: BANK_STATEMENT_RULE_VERSION,
  [CREDIT_CARD_RULE_KEY]: CREDIT_CARD_RULE_VERSION,
  [BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY]: BANK_STATEMENT_AMOUNT_ONLY_RULE_VERSION,
  [CREDIT_CARD_AMOUNT_ONLY_RULE_KEY]: CREDIT_CARD_AMOUNT_ONLY_RULE_VERSION,
  [MONTHLY_INCOME_RULE_KEY]: 1,
});

const isMonthlyAggregateRule = (ruleKey) =>
  MONTHLY_AGGREGATE_RULE_KEYS.has(normalizeRuleKey(ruleKey));
```

In both `normalizeManagedRule` and `normalizeRpcSettings`, parse the value as an integer and reject anything other than `31` when `isMonthlyAggregateRule(ruleKey)` is true. Keep the generic `0..90` validation for the existing four rules. Add the approved public-key mappings and export the new constant/helper.

- [x] **Step 4: Run focused and full tests**

Run:

```powershell
node --check api/_reconciliation-automation.js
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none
```

Expected: all commands exit `0`; the full suite is at least 228 passing tests.

- [x] **Step 5: Commit Task 1**

```powershell
git add api/_reconciliation-automation.js tests/reconciliation-automation.test.js
git commit -m "feat: add monthly reconciliation rule contract"
```

---

### Task 2: Install the catalog, immutable membership model, and source-rule guard

**Files:**
- Create: `supabase-migrations/2026-08-22-financial-reconciliation-automation-pos-income.sql`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`
- Modify: `api/reconciliation-settings.js`
- Modify: `tests/reconciliation-settings.test.js`

**Interfaces:**
- Extend `financial_reconciliation_automatic_rule_definitions` and `financial_reconciliation_automatic_rule_configs` with the fifth managed entry.
- Add nullable monthly-only proposal columns `grouping_key text` and `summary_snapshot jsonb not null default '{}'::jsonb`.
- Add `financial_reconciliation_automatic_proposal_memberships` with immutable source/destination snapshots.
- Preserve the existing `replace_financial_reconciliation_source_rules(p_rules jsonb)` signature while requiring the Bank Statement → FDM Accounts `-` pair.

- [x] **Step 1: Add transactional RED fixtures**

At the end of the existing smoke, import migration 12 and add assertions for:

- exactly one v1 definition with the approved immutable definition JSON;
- config defaults `7500.00`, `31`, and all three execution flags false except normal `enabled` remains false;
- stable next priority without changing priorities of the existing four rules;
- a second `\ir` leaves definition, config, proposal schema, constraints, indexes, functions, privileges, and data unchanged;
- the membership table rejects invalid roles, duplicate membership, invalid snapshot objects, and updates;
- RLS is enabled, direct table privileges are absent, and only service-role RPC execution is granted;
- both API and database source-rule replacement reject changing/removing `import_cgd_extrato_ordem -> import_fdm_accounts (-)`.

Add the API regression:

```js
assert.equal(response.statusCode, 400);
assert.match(response.body.error, /managed POS income source rule must remain enabled with operator -/i);
assert.equal(rpcCalls.length, 0);
```

- [ ] **Step 2: Run RED gates**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-settings.test.js tests/reconciliation-automation.test.js
psql "$env:SUPABASE_DB_URL" -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

Expected: Node fails only on the missing source-rule guard. PostgreSQL fails because migration 12/schema/catalog do not yet exist. If `psql` or the database URL is unavailable, record this as an external RED/GREEN gate; do not report it as run.

- [x] **Step 3: Add reapply-safe DDL and managed rows**

Create the migration with exact-definition checks before accepting same-named constraints/indexes. The membership table contract is:

```sql
create table if not exists public.financial_reconciliation_automatic_proposal_memberships (
  proposal_id uuid not null
    references public.financial_reconciliation_automatic_proposals(id) on delete cascade,
  role text not null check (role in ('source','destination')),
  source_type text not null check (source_type in ('import_cgd_extrato_ordem','import_fdm_accounts')),
  source_id uuid not null,
  ordinal integer not null check (ordinal > 0),
  source_date date not null,
  amount numeric(14,2) not null,
  description text not null default '',
  account text not null default '',
  row_snapshot jsonb not null check (jsonb_typeof(row_snapshot) = 'object'),
  created_at timestamptz not null default now(),
  primary key (proposal_id, role, source_type, source_id),
  unique (proposal_id, role, ordinal),
  unique (proposal_id, source_type, source_id)
);
```

Add indexes `(proposal_id, role, ordinal)` and source lock lookup indexes on `(data, id)` for Bank Statement with the POS predicate and `(event_date, id)` for FDM with `account = 'Credit Card'`. Compare `pg_get_indexdef`/predicates on reapply and fail closed if an existing same-named object differs.

Insert the definition using immutable JSON such as:

```sql
jsonb_build_object(
  'matchingMode','monthly_aggregate',
  'sourceDescriptionPattern','%POS VENDAS%',
  'destinationAccount','Credit Card',
  'calendarGrouping','closed_month',
  'fixedMaxDifferenceDays',31,
  'eligibilityFloor','2026-01-01'
)
```

Insert the config only when missing, with `enabled=false`, both execution flags false, `difference_allowed=7500.00`, `max_difference_days=31`, and `priority=max(priority)+1`. Never overwrite an administrator's later tolerance/flags/priority on reapply.

- [x] **Step 4: Protect the directional source rule twice**

Extend `requireManagedAutomaticSourceRules` in `api/reconciliation-settings.js` with:

```js
{
  baseSourceType: "import_cgd_extrato_ordem",
  matchingSourceType: "import_fdm_accounts",
  operator: "-",
  displayName: "POS income",
}
```

Replace the database source-rule RPC with the same atomic validation before its delete/insert. Preserve its signature, fixed search path, owner, revokes, and service-role grant.

- [ ] **Step 5: Run focused, SQL, and full GREEN gates**

Run the commands from Step 2, then:

```powershell
node --test --test-isolation=none
git diff --check
```

Expected: Node and SQL fixtures pass, including applying migration 12 twice. If SQL remains unavailable, leave the task locally verified but explicitly gated from rollout.

- [x] **Step 6: Commit Task 2**

```powershell
git add supabase-migrations/2026-08-22-financial-reconciliation-automation-pos-income.sql tests/reconciliation-automation-rpc.smoke.sql api/reconciliation-settings.js tests/reconciliation-settings.test.js
git commit -m "feat: store monthly reconciliation memberships"
```

---

### Task 3: Analyze closed months as immutable aggregate proposals

**Files:**
- Modify: `supabase-migrations/2026-08-22-financial-reconciliation-automation-pos-income.sql`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`

**Interfaces:**
- Extend `financial_reconciliation_automatic_rule_contract(text, integer)` with a literal fifth branch.
- Add `financial_reconciliation_automatic_monthly_income_count()` and `financial_reconciliation_automatic_monthly_income_page(p_after_month date, p_limit integer)`.
- Replace the latest `create_financial_reconciliation_automatic_analysis`, `continue_financial_reconciliation_automatic_analysis`, finalizer, progress, and run-detail functions without changing their public signatures.
- For the monthly rule, `analysis_total` and `analysis_processed` count eligible months, `analysis_cursor_date` stores the calendar month's first day, and `analysis_cursor_id` stores the earliest Bank Statement ID for stable existing response compatibility.

- [x] **Step 1: Add behavior-bearing analysis fixtures**

Create deterministic 2026 fixtures covering:

- records before `2026-01-01` excluded;
- December/January year boundary and February leap/non-leap month grouping;
- current-month rows excluded on both sides;
- description matching case-insensitively anywhere in Bank Statement text;
- FDM account matching exactly `Credit Card`;
- locked source or destination rows excluded;
- months missing either side produce no proposal row;
- `abs(difference) = tolerance` is `proposed`;
- `abs(difference) = tolerance + 0.01` is `ambiguous` with `monthly_difference_exceeded`;
- every eligible row is stored once, ordered by `(source_date, source_id)` per role;
- the technical base is the earliest Bank row;
- month proposals are oldest first across at least 30 months and continuation resumes without duplicates;
- 1,000 Bank plus 1,000 FDM rows produce one proposal with exact counts/totals and complete memberships;
- reapply changes no proposal summary, membership, timestamp, signature, or existing four-rule output.

- [ ] **Step 2: Capture PostgreSQL RED**

Run:

```powershell
psql "$env:SUPABASE_DB_URL" -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

Expected: monthly analysis fixtures fail because the fifth dispatcher branch does not exist.

- [x] **Step 3: Implement the literal monthly page query**

Use materialized source/destination aggregates and an inner join by month:

```sql
with source_rows as materialized (
  select bank.id, bank.data, bank.montante, bank.descritivo,
         date_trunc('month', bank.data)::date as calendar_month
  from public.import_cgd_extrato_ordem bank
  where bank.data >= date '2026-01-01'
    and bank.data < date_trunc('month', current_date)::date
    and bank.descritivo ilike '%POS VENDAS%'
    and not exists (
      select 1 from public.financial_reconciliation_items locked
      where locked.source_type = 'import_cgd_extrato_ordem' and locked.source_id = bank.id
    )
), destination_rows as materialized (
  select fdm.id, fdm.event_date, fdm.amount, fdm.description, fdm.account,
         date_trunc('month', fdm.event_date)::date as calendar_month
  from public.import_fdm_accounts fdm
  where fdm.event_date >= date '2026-01-01'
    and fdm.event_date < date_trunc('month', current_date)::date
    and fdm.account = 'Credit Card'
    and not exists (
      select 1 from public.financial_reconciliation_items locked
      where locked.source_type = 'import_fdm_accounts' and locked.source_id = fdm.id
    )
), source_months as (
  select
    calendar_month,
    count(*)::integer as source_count,
    sum(montante)::numeric(14,2) as source_total,
    (array_agg(id order by data, id))[1] as technical_base_source_id,
    min(data) as technical_base_source_date
  from source_rows
  group by calendar_month
), destination_months as (
  select
    calendar_month,
    count(*)::integer as destination_count,
    sum(amount)::numeric(14,2) as destination_total
  from destination_rows
  group by calendar_month
)
select
  source.calendar_month,
  source.source_count,
  source.source_total,
  destination.destination_count,
  destination.destination_total,
  (source.source_total - destination.destination_total)::numeric(14,2) as calculated_difference,
  source.technical_base_source_id,
  source.technical_base_source_date
from source_months source
join destination_months destination using (calendar_month)
where source.calendar_month > coalesce(p_after_month, date '0001-01-01')
order by source.calendar_month
limit p_limit;
```

Keep all amounts numeric until the final `numeric(14,2)` summary. Build `summary_snapshot` with `calendarMonth`, counts, totals, calculated difference, and technical base ID. Insert proposal plus all source/destination memberships in one subtransaction per month; any partial insert must roll back.

- [x] **Step 4: Add a dedicated continuation branch**

Before the existing record-based dispatcher, branch only on the exact rule/version. Page at most 25 months, advance the month cursor monotonically, update processed/total using nondecreasing values, and finalize only after the page is exhausted. Preserve the four existing branches byte-for-byte where practical.

Set:

```sql
v_status := case when abs(v_difference) <= v_difference_allowed
  then 'proposed' else 'ambiguous' end;
v_reason := case when v_status = 'ambiguous'
  then 'monthly_difference_exceeded' else '' end;
```

Hash a deterministic signature containing rule key/version, calendar month, ordered source IDs, ordered destination IDs, totals, tolerance, and operator snapshot.

- [x] **Step 5: Serialize summaries, not thousands of rows**

Replace `get_financial_reconciliation_automatic_run` so every proposal includes `groupingKey` and `summarySnapshot`, but monthly proposals keep `items` and `candidateGroups` empty. Member snapshots must only be available through the paging RPC added later. Preserve response shape and ordering for the existing four rules.

- [ ] **Step 6: Run SQL and regression GREEN**

```powershell
psql "$env:SUPABASE_DB_URL" -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
node --test --test-isolation=none
git diff --check
```

Expected: the transactional smoke passes on first and second apply, the 1,000-row fixture has exact totals/counts, and the full Node suite remains green.

- [x] **Step 7: Commit Task 3**

```powershell
git add supabase-migrations/2026-08-22-financial-reconciliation-automation-pos-income.sql tests/reconciliation-automation-rpc.smoke.sql
git commit -m "feat: analyze monthly POS income groups"
```

---

### Task 4: Execute a monthly proposal atomically and auditably

**Files:**
- Modify: `supabase-migrations/2026-08-22-financial-reconciliation-automation-pos-income.sql`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`

**Interfaces:**
- Replace `execute_financial_reconciliation_automatic_proposal(uuid, text)` with an exact literal monthly branch that delegates existing four rule/version pairs to the prior implementation.
- Preserve the normal reconciliation tables, locks, origin fields, and history behavior.
- Expand the internal FDM source lookup used by `financial_reconciliation_action` to accept `account = 'Credit Card'` for this managed path, while retaining existing FDM eligibility.

- [x] **Step 1: Add execution RED fixtures first**

Cover:

- zero difference completes normally with every membership locked;
- nonzero difference within tolerance force-completes with the generated comment;
- above-tolerance ambiguous proposals cannot execute;
- a second execution call returns the same completed reconciliation and adds no rows/audit;
- membership gained/lost, source predicate changed, destination account changed, date moved to another month, amount changed, source rule/operator changed, definition/config/tolerance/priority changed, or any source already reconciled returns sanitized `stale` and creates no reconciliation;
- concurrent destination consumption and deterministic lock ordering do not partially write;
- failure after reconciliation start rolls back reconciliation, items, audit, origin links, and proposal completion, then persists only the sanitized failed/stale proposal outcome;
- exactly one `automatic_complete` audit record contains rule/config/operator/summary/member snapshots and the generated comment;
- 1,000 + 1,000 member execution completes and history reports both source counts/totals;
- all four existing rule execution fixtures remain unchanged.

- [ ] **Step 2: Capture PostgreSQL RED**

Run the transactional smoke and confirm only monthly execution cases fail.

- [x] **Step 3: Implement deterministic revalidation and locking**

In this order:

1. Lock run, proposal, current definition/config, and required directional source rule.
2. Parse and validate the run's immutable rule snapshot before any casts.
3. Lock stored memberships ordered by `(source_type, source_date, source_id)`.
4. Lock live Bank and FDM rows in the same deterministic order.
5. Re-run the exact monthly eligibility query.
6. Compare the current ordered `(role, source_type, source_id, source_date, amount)` set and totals against stored memberships and `summary_snapshot` in both directions.
7. Return `stale` before writes on any mismatch.

- [x] **Step 4: Create and complete one ordinary reconciliation**

Use the earliest Bank member as the technical `start`, add every remaining Bank member, then add every FDM member. Verify the final locked-item count is `sourceCount + destinationCount`, matching-source rules contain FDM with `-`, and `difference_amount` equals the snapshotted monthly difference.

For zero difference call normal completion. Otherwise build this nonblank generated comment and force-complete:

```text
Automatic monthly reconciliation for YYYY-MM: Bank Statement total X.XX EUR; FDM Credit Card total Y.YY EUR; difference Z.ZZ EUR within allowed A.AA EUR; run <run-id>; proposal <proposal-id>.
```

Set `origin='automatic'`, trigger, rule/version, run ID, and proposal ID before completion. Persist the complete immutable audit metadata, then mark the proposal completed inside the same subtransaction.

- [ ] **Step 5: Run SQL/full GREEN and commit**

```powershell
psql "$env:SUPABASE_DB_URL" -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
node --test --test-isolation=none
git diff --check
git add supabase-migrations/2026-08-22-financial-reconciliation-automation-pos-income.sql tests/reconciliation-automation-rpc.smoke.sql
git commit -m "feat: execute monthly POS income reconciliations"
```

Expected: all lifecycle, rollback, stale, idempotency, large-membership, history, and existing-rule cases pass.

---

### Task 5: Add app-authorized member paging

**Files:**
- Create: `api/reconciliation-automation-members.js`
- Modify: `api/_reconciliation-automation.js`
- Modify: `tests/reconciliation-automation.test.js`
- Modify: `supabase-migrations/2026-08-22-financial-reconciliation-automation-pos-income.sql`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`

**Interfaces:**
- HTTP: `GET /api/reconciliation-automation-members?run_id=<uuid>&proposal_id=<uuid>&role=source|destination&offset=<n>&limit=50`.
- RPC: `get_financial_reconciliation_automatic_proposal_members(p_run_id uuid, p_proposal_id uuid, p_role text, p_offset integer, p_limit integer, p_actor text) returns jsonb`.
- Response: `{ runId, proposalId, role, offset, limit, totalCount, members }`.

- [x] **Step 1: Add failing API and RPC tests**

API tests must prove app feature authorization, actor binding, UUID validation, exact role allowlist, `offset >= 0`, `1 <= limit <= 50`, GET-only behavior, RPC-only data access, camelCase response mapping, and sanitized database errors.

SQL tests must prove manual-run ownership, run/proposal relationship, monthly-rule restriction, role isolation, stable ordinal ordering, no duplicate/skip across pages `0`, `50`, `100`, a short final page, invalid/foreign-run rejection, completed snapshot readability, ACLs, and no direct table reads by browser roles.

- [ ] **Step 2: Capture RED**

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
psql "$env:SUPABASE_DB_URL" -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

- [x] **Step 3: Implement the narrow handler**

Use the existing `requireFeature(req, "app", "financial-reconciliation")`, authenticated actor identity, and `toAutomationPublicResult`. Call only:

```js
restQuery("rpc/get_financial_reconciliation_automatic_proposal_members", {
  method: "POST",
  body: {
    p_run_id: runId,
    p_proposal_id: proposalId,
    p_role: role,
    p_offset: offset,
    p_limit: limit,
    p_actor: actor,
  },
});
```

Do not query the membership table from Node.

The RPC must validate before querying and return only the stored snapshot projection:

```sql
jsonb_build_object(
  'runId', p_run_id,
  'proposalId', p_proposal_id,
  'role', p_role,
  'offset', p_offset,
  'limit', p_limit,
  'totalCount', v_total,
  'members', coalesce(v_members, '[]'::jsonb)
)
```

For unfinished manual runs, require `run.actor = p_actor`. Keep service-role-only execution and fixed search path.

- [ ] **Step 4: Run GREEN and commit**

```powershell
node --check api/reconciliation-automation-members.js
node --test --test-isolation=none tests/reconciliation-automation.test.js
psql "$env:SUPABASE_DB_URL" -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
node --test --test-isolation=none
git diff --check
git add api/reconciliation-automation-members.js api/_reconciliation-automation.js tests/reconciliation-automation.test.js supabase-migrations/2026-08-22-financial-reconciliation-automation-pos-income.sql tests/reconciliation-automation-rpc.smoke.sql
git commit -m "feat: page monthly reconciliation members"
```

---

### Task 6: Render Settings and the three-column monthly review

**Files:**
- Modify: `app-main.js`
- Modify: `styles.css`
- Modify: `tests/reconciliation-automation-ui.test.js`
- Modify: `tests/reconciliation-density.test.js`

**Interfaces:**
- Settings adds a fifth managed card with editable tolerance, flags, and priority, fixed/read-only `31 days`, and read-only definition text.
- Workbench selector includes the rule only when it is enabled and manual execution is allowed.
- Monthly proposal renderer uses summary/meta, source group, destination group columns.
- Paging state is independent for `${proposalId}:source` and `${proposalId}:destination`; page size is exactly 50.

- [x] **Step 1: Add UI RED tests using actual extracted production helpers**

Assert:

- the fifth card renders escaped immutable logic and `<output class="financial-reconciliation-automation-fixed-value" aria-label="Maximum difference in days, fixed">31 days</output>`, with no editable max-days control;
- serialization reasserts `31` despite DOM/state tampering while preserving edited tolerance/flags/priority;
- the rule selector uses the approved display name and single-rule Analyze payload;
- `monthly_difference_exceeded` renders a safe audit label and has no execution checkbox;
- the monthly proposal has exactly three top-level columns and does not render a fake single base/destination item;
- both `<details>` groups start closed and show count/total in their summaries;
- opening source loads only source page 0, opening destination loads only destination page 0;
- `Load more` appends the next 50 in ordinal order without replacing, collapsing, or reloading the other group;
- pending and error states remain local to the selected group;
- member descriptions/accounts are escaped and record IDs remain in disclosure controls;
- proposal selection, execute-selected, retained-run restoration, finished-run summaries, and existing four proposal layouts remain unchanged.

- [x] **Step 2: Capture focused RED**

```powershell
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
```

Expected: failures are confined to the missing fifth Settings card/monthly renderer/paging behavior.

- [x] **Step 3: Add explicit monthly rendering/state helpers**

Add:

```js
function isFinancialReconciliationMonthlyAggregateRule(ruleKey) {
  return clean(ruleKey) === "cgd_bank_statement_fdm_credit_card_monthly_income";
}

function financialReconciliationAutomationMembershipKey(proposalId, role) {
  return `${clean(proposalId)}:${role}`;
}
```

Keep membership state outside proposal snapshots:

```js
automation.memberships[key] = {
  role,
  members: [],
  offset: 0,
  totalCount: 0,
  loaded: false,
  loading: false,
  error: "",
};
```

Clear this map when a new run replaces the current run, but retain it through unrelated renders of the same run.

- [x] **Step 4: Render the approved three columns**

Produce this semantic structure:

```html
<article class="financial-reconciliation-automation-proposal financial-reconciliation-automation-proposal--monthly">
  <section class="financial-reconciliation-automation-proposal-meta">Monthly summary</section>
  <details class="financial-reconciliation-automation-member-group" data-role="source"><summary>CGD Bank Statement members</summary></details>
  <details class="financial-reconciliation-automation-member-group" data-role="destination"><summary>FDM Accounts members</summary></details>
</article>
```

The meta column shows status/reason, month, difference, tolerance, rule/version, and execution checkbox only for `proposed`. The source summary reads `CGD Bank Statement (#N; X.XX €)`; destination reads `FDM Accounts (#N; Y.YY €)`. Each member row shows date, description, amount, and FDM account where present. Do not display all memberships until their group opens.

- [x] **Step 5: Implement asynchronous group paging**

On the first `toggle` to open, call the new API with `offset=0&limit=50`. `Load more` calls with `offset=members.length&limit=50`, validates the returned run/proposal/role/offset, rejects duplicate IDs, and appends. Disable only that group's controls while pending.

- [x] **Step 6: Add responsive/accessibility CSS**

Desktop uses:

```css
.financial-reconciliation-automation-proposal--monthly {
  grid-template-columns: minmax(11rem, .55fr) minmax(0, 1fr) minmax(0, 1fr);
}
```

Keep vertical separators and proposal row dividers. At `max-width: 768px`, stack meta/source/destination in DOM order, remove vertical separators, retain horizontal separators, use 16px controls, visible focus rings, wrapping descriptions, and non-overlapping buttons.

- [x] **Step 7: Run focused/full GREEN and commit**

```powershell
node --check app-main.js
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
node --test --test-isolation=none
git diff --check
git add app-main.js styles.css tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
git commit -m "feat: review monthly reconciliation groups"
```

---

### Task 7: Extend sequential scheduling to the fifth managed rule

**Files:**
- Modify: `api/reconciliation-automation-cron.js` only if production validation needs the fifth contract explicitly
- Modify: `tests/reconciliation-automation.test.js`
- Modify: `supabase-migrations/2026-08-22-financial-reconciliation-automation-pos-income.sql`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`

**Interfaces:**
- The shared once-daily schedule snapshots all enabled scheduled rules ordered by administrator priority.
- One child rule runs at a time; the next child is not claimed until the prior child is terminal.
- Monthly analysis may span multiple heartbeats but resumes the same run/month cursor.

- [x] **Step 1: Add five-rule scheduler RED tests**

Node tests use a production-shaped five-rule batch and prove claim → continue → ready → execute → finalize → next-rule sequencing, exact counts, no second claim while monthly analysis is unfinished, sanitized monthly failure continuation, and cross-midnight resume of the oldest unfinished batch.

SQL fixtures prove disabled-by-default exclusion, administrator priority reorder, optional scheduled enablement, immutable five-rule snapshot, one child per position, retry idempotency, failed-child continuation, aggregate parent status, and unchanged four-rule behavior when the new rule remains disabled.

- [ ] **Step 2: Capture RED, implement the minimum, and run GREEN**

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
psql "$env:SUPABASE_DB_URL" -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

If adding `MONTHLY_INCOME_RULE_KEY` to `AUTOMATIC_RULE_VERSIONS` already makes the Node cron generic, change no cron production code; keep only the regression tests. In SQL, replace the latest settings/claim functions only as needed to validate five exact managed contracts and fixed `31` for this rule.

- [ ] **Step 3: Verify and commit**

```powershell
node --check api/reconciliation-automation-cron.js
node --test --test-isolation=none tests/reconciliation-automation.test.js
psql "$env:SUPABASE_DB_URL" -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
node --test --test-isolation=none
git diff --check
git add api/reconciliation-automation-cron.js tests/reconciliation-automation.test.js supabase-migrations/2026-08-22-financial-reconciliation-automation-pos-income.sql tests/reconciliation-automation-rpc.smoke.sql
git commit -m "test: verify five-rule reconciliation batches"
```

Omit `api/reconciliation-automation-cron.js` from the commit if it did not need a production change.

---

### Task 8: Document rollout and pass all release gates

**Files:**
- Modify: `README.md`
- Modify: `docs/superpowers/plans/2026-08-22-card-payments-pos-income-automatic-reconciliation.md` only to mark completed checkboxes during execution

**Interfaces:**
- README migration order includes migration 12 after history migration 11.
- Rollout keeps the fifth rule disabled for both manual and scheduled execution until SQL/browser/non-production gates pass.

- [x] **Step 1: Add migration and rollout instructions**

Document:

1. apply migrations 1–11 in the existing order;
2. apply `2026-08-22-financial-reconciliation-automation-pos-income.sql`;
3. apply it a second time to prove reapply safety;
4. run both reconciliation SQL smokes;
5. verify the rule appears disabled in Settings;
6. enable manual only in protected non-production, analyze and execute a closed month;
7. enable scheduled only after the protected heartbeat passes.

- [x] **Step 2: Run final static and Node verification**

```powershell
node --check api/_reconciliation-automation.js
node --check api/reconciliation-automation.js
node --check api/reconciliation-automation-members.js
node --check api/reconciliation-automation-settings.js
node --check api/reconciliation-automation-cron.js
node --check api/reconciliation-settings.js
node --check app-main.js
node --test --test-isolation=none
node -e "JSON.parse(require('fs').readFileSync('vercel.json','utf8')); console.log('vercel.json OK')"
git diff --check
git status --short
```

Expected: all syntax checks pass, full Node suite has zero fail/skip/cancel/todo, JSON parses, diff check is clean, and only intended feature files are changed.

- [ ] **Step 3: Run mandatory PostgreSQL gates**

```powershell
psql "$env:SUPABASE_DB_URL" -v ON_ERROR_STOP=1 -f tests/reconciliation-rpc.smoke.sql
psql "$env:SUPABASE_DB_URL" -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

The automation smoke must apply and reapply migration 12 inside its rollback transaction. No production publish is cleared if this gate is unavailable or fails.

- [ ] **Step 4: Verify authenticated browser behavior**

At desktop and `<=768px`, verify:

1. Settings shows the fifth card disabled, tolerance editable, days fixed at 31, definition read-only.
2. Saving flags/priority/tolerance reloads authoritatively.
3. Automatic Reconciliation defaults to the manual tab and lists the rule only after manual enablement.
4. Analyze displays closed months oldest first.
5. An above-tolerance month is ambiguous and cannot be selected.
6. Source and destination groups start collapsed.
7. Opening each group loads 50 rows independently; Load more appends without duplicates.
8. A zero-difference proposal completes normally.
9. A nonzero in-tolerance proposal force-completes with the generated comment.
10. History shows Automatic · Manual/Scheduled origin and source counts/totals.

Record screenshots and console/network errors. If no authenticated fixture exists, document the exact limitation and leave the gate open.

- [ ] **Step 5: Verify one protected scheduled heartbeat**

In non-production, enable the fifth rule for the shared schedule, invoke one authorized heartbeat, and verify sequential child creation/resume/finalization plus idempotent retry. Disable it again unless production enablement was separately approved.

- [ ] **Step 6: Request independent review**

Use `superpowers:requesting-code-review` for:

- SQL correctness, reapply safety, ACL/RLS, lock ordering, stale/rollback behavior, and 1,000-row query plans;
- API authorization and response validation;
- UI state, escaping, accessibility, independent paging, desktop/narrow layout;
- exact preservation of the four existing managed rules.

Resolve every Critical/Important finding with a RED/GREEN regression before integration.

- [x] **Step 7: Commit rollout documentation**

```powershell
git add README.md docs/superpowers/specs/2026-08-22-card-payments-pos-income-automatic-reconciliation-design.md docs/superpowers/plans/2026-08-22-card-payments-pos-income-automatic-reconciliation.md
git commit -m "docs: add POS income reconciliation rollout"
```

Only after every available gate is green should `superpowers:finishing-a-development-branch` be used to offer local merge/publish choices.
