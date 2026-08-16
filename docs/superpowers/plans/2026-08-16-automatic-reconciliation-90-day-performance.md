# Automatic Reconciliation 90-Day Performance Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the Financial Documents to CGD Bank Statement rule complete reliably with date windows up to 90 days by using an indexed CGD search projection and resumable 25-record analysis pages.

**Architecture:** One forward-only Supabase migration creates and synchronizes a pre-normalized CGD search projection, installs indexed exact-semantic candidate paging and single-base revalidation, and makes automatic-analysis runs resumable. The API and cron advance one atomic page at a time, while the browser restores unfinished manual runs, displays progress, and serially continues analysis until proposals are ready.

**Tech Stack:** PostgreSQL/Supabase RPC and PostgREST, `pg_trgm`, Vercel Node functions, browser JavaScript, Node's built-in test runner, transactional SQL smoke tests.

## Global Constraints

- The configured maximum date difference is inclusive and must be between 0 and 90 days.
- Keep rule key `financial_documents_cgd_bank_statement` and immutable rule version `2`.
- Preserve exact `fat = 'S'`, exact `payment = 'Banco'`, and the `2026-01-01` eligibility floor.
- Preserve document-number containment, description similarity `>= 0.60`, supplier word similarity `>= 0.70`, candidate cap 12, group size 4, integer-cent arithmetic, source operators, ambiguity, locks, stale checks, and audit history.
- Analysis page size is the server-owned constant 25 and is never accepted from a browser payload.
- Partial runs and proposals are never executable.
- All new SQL helpers are `SECURITY DEFINER SET search_path = public, pg_temp`; `anon` and `authenticated` receive no execution rights.
- Existing completed runs are not rewritten. Unfinished pre-migration runs become failed with `analysis_upgrade_restart_required`.
- Do not expose Supabase exception detail in public API responses.

---

### Task 1: Public contract and 90-day validation

**Files:**
- Modify: `api/_reconciliation-automation.js`
- Modify: `tests/reconciliation-automation.test.js`

**Interfaces:**
- Consumes: existing `normalizeAutomationAction`, `normalizeAnalyzePayload`, `toAutomationPublicResult`, and settings normalization.
- Produces: `normalizeContinueAnalysisPayload(value) -> { action: "continue_analysis", runId: string }`, action membership for `continue_analysis`, a 0–90 day validator, and snake-to-camel mappings for the six analysis progress fields.

- [ ] **Step 1: Add failing behavior tests**

Add literal assertions that catch these regressions:

```js
test("automatic reconciliation caps managed date windows at 90 days", () => {
  assert.equal(normalizeManagedSettings({
    ...managedSettings(),
    rules: [{ ...managedSettings().rules[0], maxDifferenceDays: 90 }],
  }).rules[0].maxDifferenceDays, 90);
  assert.throws(() => normalizeManagedSettings({
    ...managedSettings(),
    rules: [{ ...managedSettings().rules[0], maxDifferenceDays: 91 }],
  }), /between 0 and 90/i);
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
```

Import `normalizeContinueAnalysisPayload` from the real helper. Expected values are literals and do not reuse production normalization.

- [ ] **Step 2: Run the focused test and verify RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: FAIL because 91 is currently allowed, `continue_analysis` is unknown, the continue normalizer is absent, and progress snake-case keys are not mapped.

- [ ] **Step 3: Implement the minimal helper contract**

In `api/_reconciliation-automation.js`:

```js
const AUTOMATION_ACTIONS = new Set([
  "analyze_rule",
  "analyze_batch",
  "continue_analysis",
  "execute_selected",
]);

function normalizeContinueAnalysisPayload(value) {
  const input = requirePlainObject(value, "Continue analysis payload");
  requireOnlyKeys(input, new Set(["action", "runId"]), "Continue analysis payload");
  if (normalizeAutomationAction(input.action) !== "continue_analysis") {
    throw inputError("Continue analysis action is invalid.");
  }
  return { action: "continue_analysis", runId: normalizeUuid(input.runId, "Run ID") };
}
```

Change both managed-rule validators from maximum 365 to 90. Add mappings:

```js
analysis_cursor_date: "analysisCursorDate",
analysis_cursor_id: "analysisCursorId",
analysis_processed: "analysisProcessed",
analysis_total: "analysisTotal",
analysis_error_code: "analysisErrorCode",
analysis_error_at: "analysisErrorAt",
```

Export `normalizeContinueAnalysisPayload`.

- [ ] **Step 4: Run focused and full tests to verify GREEN**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none
```

Expected: both exit 0 with no failures.

- [ ] **Step 5: Commit Task 1**

```powershell
git add -- api/_reconciliation-automation.js tests/reconciliation-automation.test.js
git commit -m "feat: cap automatic reconciliation at 90 days"
```

---

### Task 2: Indexed projection and resumable database analysis

**Files:**
- Create: `supabase-migrations/2026-08-16-financial-reconciliation-automation-90-day-performance.sql`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`
- Modify: `tests/reconciliation-automation.test.js`

**Interfaces:**
- Consumes: rule version 2, existing normalization/score wrappers, `financial_reconciliation_automatic_build_combinations`, proposal schema/signatures, source rules, and run lifecycle.
- Produces:
  - `financial_reconciliation_cgd_match_search` and synchronization trigger;
  - `financial_reconciliation_automatic_candidate_page(text,integer,numeric,integer,date,uuid,integer) -> existing candidate columns`;
  - `financial_reconciliation_automatic_single_base_candidates(text,integer,numeric,integer,uuid)`;
  - `continue_financial_reconciliation_automatic_analysis(uuid,text) -> jsonb`;
  - `financial_reconciliation_automatic_progress_or_run(uuid) -> jsonb`;
  - `get_financial_reconciliation_automatic_active_run(text) -> jsonb`;
  - `continue_financial_reconciliation_automatic_oldest_analysis(text) -> jsonb` for the cron worker;
  - resumable `create_financial_reconciliation_automatic_analysis(...)`;
  - existing compatibility candidate and execution functions backed by the optimized helpers.

- [ ] **Step 1: Add failing migration and SQL behavior coverage**

Add `AUTOMATION_90_DAY_MIGRATION_PATH` to `tests/reconciliation-automation.test.js`. Add a deployment-contract test that requires the migration file, includes it after the Banco-v2 migration in the smoke suite, and checks that the normal migration path is named exactly as above.

Extend `tests/reconciliation-automation-rpc.smoke.sql` with transactional fixtures that:

1. insert one projected CGD row and assert normalized/compact values;
2. update its description/date/amount and assert the projection changes;
3. delete it and assert the projection disappears;
4. insert bases and bank rows on day 90 and day 91, then assert only day 90 qualifies;
5. insert 30 ordered bases, continue twice, and assert processed counts are 25 then 30 with no duplicate proposals;
6. call continuation again and assert the ready run is unchanged;
7. create an incomplete run and prove execution rejects it;
8. retrieve the actor's active run and prove another actor cannot retrieve it;
9. compare candidate IDs, evidence booleans, and exact scores from the compatibility function with hand-checked v2 fixtures;
10. inspect function privileges and `proconfig` for every new helper.

Use fixed UUIDs and dates. Derive expected candidate IDs and counts literally in the smoke assertions.

- [ ] **Step 2: Run Node RED and record the SQL external gate**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: FAIL because the migration does not exist and the smoke suite does not include it.

If `psql` and `SUPABASE_DB_URL` are available, also run:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

Expected before implementation: FAIL on the missing projection/helper contracts. If the tools are unavailable, record this exact PostgreSQL smoke as a mandatory external gate rather than claiming it ran.

- [ ] **Step 3: Create the projection, trigger, indexes, and 90-day constraint**

Start the migration with idempotent DDL:

```sql
create table if not exists public.financial_reconciliation_cgd_match_search (
  source_id uuid primary key,
  source_date date not null,
  amount numeric,
  description text,
  normalized_description text not null,
  compact_description text not null,
  updated_at timestamptz not null default now()
);

create index if not exists financial_reconciliation_cgd_match_search_date_idx
  on public.financial_reconciliation_cgd_match_search (source_date, source_id);
create index if not exists financial_reconciliation_cgd_match_search_normalized_trgm_idx
  on public.financial_reconciliation_cgd_match_search using gin (normalized_description gin_trgm_ops);
create index if not exists financial_reconciliation_cgd_match_search_compact_trgm_idx
  on public.financial_reconciliation_cgd_match_search using gin (compact_description gin_trgm_ops);
```

Install `sync_financial_reconciliation_cgd_match_search()` as a fixed-search-path trigger function. DELETE removes `old.id`; INSERT/UPDATE upserts `new.id`, `new.data`, `new.montante`, `new.descritivo`, `coalesce(financial_reconciliation_match_normalize(new.descritivo), '')`, and `coalesce(financial_reconciliation_match_compact(new.descritivo), '')`. If an UPDATE changes ID, remove the old projection row first. Backfill with `INSERT ... SELECT ... ON CONFLICT DO UPDATE` before enabling the trigger.

Discover the installed `pg_trgm` extension schema from `pg_extension`/`pg_namespace`. Build both GIN indexes with its schema-qualified `gin_trgm_ops`, and dynamically install candidate helper SQL with schema-qualified `OPERATOR(schema.%)` and `OPERATOR(schema.<%)`. This preserves fixed function search paths and works whether Supabase installed the extension in `public` or `extensions`.

Clamp existing rule configs above 90 to 90, drop the existing max-days check by its catalog-discovered name, and add:

```sql
constraint financial_reconciliation_automatic_rule_configs_max_days_check
  check (max_difference_days between 0 and 90)
```

Add the six resumable columns with the types/defaults from the design, and transition unfinished pre-migration rows to `failed`, `finished_at = now()`, and `error_summary = 'analysis_upgrade_restart_required'`.

- [ ] **Step 4: Implement indexed paged and single-base candidate helpers**

Page eligible bases first:

```sql
where d.fat = 'S'
  and d.payment = 'Banco'
  and d.document_date >= date '2026-01-01'
  and (p_after_date is null or (d.document_date, d.id) > (p_after_date, p_after_id))
order by d.document_date, d.id
limit least(greatest(p_page_size, 1), 25)
```

For each page base, union index-assisted source IDs from the projection:

```sql
s.compact_description like '%' || d.compact_document_number || '%'
s.normalized_description % d.normalized_document_description
d.normalized_supplier_name <% s.normalized_description
```

Limit every branch to the inclusive date range and unlocked source rows. Set `pg_trgm.similarity_threshold=0.60` and `pg_trgm.word_similarity_threshold=0.70` in the function configuration. Recompute the existing exact containment, `similarity`, and `word_similarity` values after the indexed prefilter and accept only the exact OR expression. Preserve stable candidate ordering and evidence JSON.

The single-base helper uses the same query with `d.id = p_base_source_id`. The compatibility function enumerates page helpers for smoke/backward compatibility but is not used by normal create/continue/execute paths.

- [ ] **Step 5: Implement atomic run continuation and finalization**

Install `continue_financial_reconciliation_automatic_analysis(p_run_id uuid, p_actor text)` with this lifecycle:

```sql
select * into strict v_run
from public.financial_reconciliation_automatic_runs
where id = p_run_id
for update;

if v_run.analysis_completed_at is not null then
  return public.get_financial_reconciliation_automatic_run(v_run.id);
end if;

if v_run.trigger = 'manual' and v_run.actor <> trim(p_actor) then
  raise exception 'Automatic analysis run belongs to another actor.';
end if;
```

Call the page helper with the stored cursor and constant 25. For every returned base, execute the existing candidate-limit / one-combination / multiple-combination / skipped decision tree and existing signature formulas. Persist the page's last `(base_source_date, base_source_id)` and increment `analysis_processed` by the number of bases returned in the same transaction.

When the page is empty, run the existing cross-base overlap CTE, recalculate counts, set `status='ready'`, and set `analysis_completed_at=now()`. Return `get_financial_reconciliation_automatic_run` after every page.

Replace create-analysis population with creation/lookup plus one continuation call. Add active-run lookup constrained to `trigger='manual'`, actor equality, and `analysis_completed_at is null`. Add an internal progress serializer that returns lifecycle/progress fields with an empty proposal array while analysis is incomplete, avoiding repeated aggregation of accumulated skipped rows; return the existing full run only at Ready. Extend `get_financial_reconciliation_automatic_run` with all six progress fields.

Add `continue_financial_reconciliation_automatic_oldest_analysis(p_worker text)`. It accepts only the existing fixed worker identity `system:reconciliation`, locks the oldest unfinished run with `FOR UPDATE SKIP LOCKED`, advances one page using the persisted run actor for ownership, and returns `jsonb_build_object('continued', false)` when none exists or `jsonb_build_object('continued', true, 'run', public.financial_reconciliation_automatic_progress_or_run(v_run_id))` after one page.

Patch execution revalidation to call the single-base helper rather than scanning the compatibility function, and reject runs without `analysis_completed_at`.

- [ ] **Step 6: Secure and expose only required RPCs**

Revoke projection/helper access from `public`, `anon`, and `authenticated`. Grant required execution to `service_role`. Keep the trigger helper unavailable even to direct `service_role` calls. End with:

```sql
notify pgrst, 'reload schema';
```

- [ ] **Step 7: Run focused/full Node GREEN and SQL smoke when available**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none
```

Then run the PostgreSQL smoke command from Step 2 when `psql` and the URL are available. Expected: all tests exit 0. If SQL cannot run locally, retain it as an explicit release gate.

- [ ] **Step 8: Commit Task 2**

```powershell
git add -- supabase-migrations/2026-08-16-financial-reconciliation-automation-90-day-performance.sql tests/reconciliation-automation-rpc.smoke.sql tests/reconciliation-automation.test.js
git commit -m "feat: resume 90-day reconciliation analysis"
```

---

### Task 3: Manual API and cron continuation

**Files:**
- Modify: `api/reconciliation-automation.js`
- Modify: `api/reconciliation-automation-cron.js`
- Modify: `tests/reconciliation-automation.test.js`

**Interfaces:**
- Consumes: `normalizeContinueAnalysisPayload`, progress-mapped run JSON, `continue_financial_reconciliation_automatic_analysis`, and `get_financial_reconciliation_automatic_active_run`.
- Produces: POST `continue_analysis`, GET `view=active_run`, actor-bound continuation, and cron page advancement before proposal execution.

- [ ] **Step 1: Add failing endpoint behavior tests**

Add real handler tests with the existing mocked external `restQuery` boundary:

```js
test("manual continuation binds the authenticated actor and uses the fixed RPC", async () => {
  // POST { action: "continue_analysis", runId: RUN_ID }
  // Assert status 200 and literal public run progress.
  // Assert the only continuation RPC body is
  // { p_run_id: RUN_ID, p_actor: "admin@example.com" }.
});

test("active run lookup is actor-scoped", async () => {
  // GET ?view=active_run
  // Assert RPC get_financial_reconciliation_automatic_active_run
  // receives { p_actor: "admin@example.com" }.
});

test("cron advances one unfinished analysis page before execution", async () => {
  // Return an oldest unfinished manual run from the worker continuation RPC
  // and assert schedule claim, execute, and finish RPCs are not called.
});
```

Add malformed/unauthorized tests proving browser-supplied `pageSize`, another actor, and fallback profiles cannot reach an RPC.

- [ ] **Step 2: Run focused RED**

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: FAIL because neither API route calls the new RPCs and cron still calls the all-at-once populate RPC.

- [ ] **Step 3: Implement manual continuation and active-run GET**

In `api/reconciliation-automation.js`, add:

```js
async function continueAnalysis(req, body) {
  const auth = await requireManagedFeature(req, "app");
  const input = normalizeContinueAnalysisPayload(body);
  return toAutomationPublicResult(await restQuery(
    "rpc/continue_financial_reconciliation_automatic_analysis",
    { method: "POST", body: { p_run_id: input.runId, p_actor: actorFor(auth) } },
  ));
}
```

Route `continue_analysis` before execution. In GET, accept exactly `view=rules` or `view=active_run`; active-run calls its RPC with the authenticated actor. Keep sanitized error behavior.

- [ ] **Step 4: Replace cron all-at-once population with one continuation page**

At the start of the heartbeat, call `rpc/continue_financial_reconciliation_automatic_oldest_analysis` with `p_worker: SCHEDULE_ACTOR`. If it returns `continued: true`, return a sanitized heartbeat response immediately; one heartbeat performs one analysis page.

When no unfinished analysis was continued, claim the normal scheduled slot. If the claimed scheduled run has no `analysisCompletedAt`, call:

```js
run = requireScheduledRun(await restQuery(
  "rpc/continue_financial_reconciliation_automatic_analysis",
  { method: "POST", body: { p_run_id: run.runId, p_actor: SCHEDULE_ACTOR } },
));
if (!run.analysisCompletedAt) {
  return res.status(200).json(publicRunResponse(claim, run, 0, true));
}
```

Only select/execute proposals after analysis completes. Extend scheduled-run validation to require non-negative integer `analysisProcessed`/`analysisTotal` and valid nullable cursor/error timestamps.

- [ ] **Step 5: Run focused and full GREEN**

```powershell
node --check api/reconciliation-automation.js
node --check api/reconciliation-automation-cron.js
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none
```

Expected: all commands exit 0.

- [ ] **Step 6: Commit Task 3**

```powershell
git add -- api/reconciliation-automation.js api/reconciliation-automation-cron.js tests/reconciliation-automation.test.js
git commit -m "feat: continue reconciliation analysis in pages"
```

---

### Task 4: Browser progress, resume, and retry

**Files:**
- Modify: `app-main.js`
- Modify: `tests/reconciliation-automation-ui.test.js`

**Interfaces:**
- Consumes: `analysisCompletedAt`, `analysisProcessed`, `analysisTotal`, active-run GET, and continue-analysis POST.
- Produces: actor-safe automatic continuation loop, progress text, reload restoration, disabled review/execution during analysis, and inline retry behavior.

- [ ] **Step 1: Add failing browser behavior tests**

Use the existing real-function extraction harness to test observable behavior:

```js
test("analysis progress renders processed and total records", () => {
  assert.match(financialReconciliationAutomationEmptyMessage({
    status: "analyzing",
    analysisCompletedAt: null,
    analysisProcessed: 25,
    analysisTotal: 876,
  }), /Analyzing 25 of 876 records/i);
});

test("proposal controls remain disabled until analysis completes", () => {
  // Render a run containing a proposed row but analysisCompletedAt null.
  // Assert select-all, clear, and execute are disabled.
});

test("automatic tab restores and continues an unfinished actor run", async () => {
  // Mock only the external api boundary.
  // Return active run -> continued run -> ready run.
  // Assert calls are sequential and final proposed IDs are selected once ready.
});

test("uncertain continuation reloads before retrying", async () => {
  // Make continuation reject once, then active-run GET return persisted progress.
  // Assert no duplicate concurrent continuation and inline Retry state.
});
```

The expected call order and progress strings are literal.

- [ ] **Step 2: Run UI RED**

```powershell
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js
```

Expected: FAIL because incomplete runs currently render the generic empty message and no restoration/continuation loop exists.

- [ ] **Step 3: Add state and pure lifecycle helpers**

Extend automation state with:

```js
continuationToken: 0,
continuationRetry: false,
```

Add pure helpers:

```js
function financialReconciliationAutomationIsAnalyzing(run) {
  return Boolean(run && !clean(run.analysisCompletedAt) && clean(run.status) !== "failed");
}

function financialReconciliationAutomationProgressLabel(run) {
  const processed = Math.max(0, Number(run?.analysisProcessed) || 0);
  const total = Math.max(processed, Number(run?.analysisTotal) || 0);
  return `Analyzing ${processed} of ${total} records…`;
}
```

Use these helpers in empty/results markup and control disabling. Do not render partial proposals as selectable.

- [ ] **Step 4: Implement restoration and serial continuation**

After rules load, GET `/api/reconciliation-automation?view=active_run`. If it returns an unfinished run, store it and call `continueFinancialReconciliationAutomationAnalysis(token)`.

The continuation loop:

1. captures the current incremented token;
2. posts `{ action: "continue_analysis", runId }` once;
3. replaces the run from the response;
4. renders progress;
5. awaits the next iteration only when the token/run are still current;
6. when ready, selects all `proposed` IDs and announces success;
7. on network uncertainty, reloads active run before setting `continuationRetry=true` and showing an inline Retry action.

Starting a new Analyze action increments the token, clears retry state, and uses the returned run as the first loop value. Tab changes do not cancel server state; a new tab load restores it.

- [ ] **Step 5: Run UI, focused automation, and full GREEN**

```powershell
node --check app-main.js
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js
node --test --test-isolation=none tests/reconciliation-automation.test.js tests/reconciliation-automation-ui.test.js
node --test --test-isolation=none
```

Expected: all commands exit 0.

- [ ] **Step 6: Commit Task 4**

```powershell
git add -- app-main.js tests/reconciliation-automation-ui.test.js
git commit -m "feat: show resumable reconciliation progress"
```

---

### Task 5: Migration ordering, production-size verification, and release gate

**Files:**
- Modify: `README.md`
- Modify: `tests/reconciliation-automation.test.js`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`

**Interfaces:**
- Consumes: the complete Task 1–4 implementation.
- Produces: authoritative migration order, idempotent reapplication proof, final regression evidence, and a measured 90-day production gate.

- [ ] **Step 1: Add failing release-contract tests**

Require README and both SQL smoke bootstrap sections to place migrations in this order:

```text
2026-08-14-financial-reconciliation-automation-schema.sql
2026-08-14-financial-reconciliation-automation-analysis.sql
2026-08-14-financial-reconciliation-automation-execution.sql
2026-08-15-financial-reconciliation-automation-analysis-performance.sql
2026-08-15-financial-reconciliation-automation-candidate-index-lookup.sql
2026-08-16-financial-reconciliation-automation-banco-v2.sql
2026-08-16-financial-reconciliation-automation-90-day-performance.sql
```

The smoke suite must apply the new migration twice and rerun projection synchronization plus continuation assertions after the second application.

- [ ] **Step 2: Run focused RED**

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: FAIL until README and smoke ordering/reapplication are complete.

- [ ] **Step 3: Document the exact operator procedure**

Update README with:

1. the migration order above;
2. a statement that only the final new migration is needed on a database already current through Banco v2;
3. the supported settings range 0–90 days;
4. the required post-migration SQL smoke command;
5. the production verification steps: save 90 days, Analyze, observe progress, wait for Ready, compare proposal/counter semantics, and confirm no 500/timeout in logs.

- [ ] **Step 4: Run complete local verification**

```powershell
node --check api/_reconciliation-automation.js
node --check api/reconciliation-automation.js
node --check api/reconciliation-automation-cron.js
node --check app-main.js
node --test --test-isolation=none
git diff --check
```

Parse `vercel.json` with:

```powershell
Get-Content -Raw vercel.json | ConvertFrom-Json | Out-Null
```

Expected: every command exits 0 and the full Node suite has zero failures.

- [ ] **Step 5: Run the database and production-size gates**

When PostgreSQL access is available:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

Apply the new migration to the target Supabase project. Configure the rule to 90 days and run Analyze in the authenticated app. Require:

- HTTP 200 for create and every continuation;
- monotonically increasing processed counts ending at total;
- Ready within two minutes on the current dataset;
- no statement timeout or HTTP 500;
- no duplicate proposals;
- unchanged exact candidate/evidence fixtures and reconciliation history.

If live database or authenticated browser access is unavailable, report those gates as not run and do not claim production rollout readiness.

- [ ] **Step 6: Request independent code/spec review**

Provide the reviewer with the base SHA before Task 1, final HEAD, this plan, the design specification, complete diff, and verification output. Resolve every Critical and Important finding with another strict red-green cycle.

- [ ] **Step 7: Commit Task 5**

```powershell
git add -- README.md tests/reconciliation-automation.test.js tests/reconciliation-automation-rpc.smoke.sql
git commit -m "docs: verify 90-day reconciliation rollout"
```
