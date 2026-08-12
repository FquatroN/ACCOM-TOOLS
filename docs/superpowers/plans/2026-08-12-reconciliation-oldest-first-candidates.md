# Reconciliation Oldest-First Candidates Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make eligible reconciliation candidates deterministically oldest-first across pagination and replace the Current reconciliation source sentence with a compact status, locked-record count, and shortened difference row.

**Architecture:** Add an idempotent forward PostgreSQL migration that transforms only the installed workspace function's two candidate-order clauses, preserving all previously deployed source-rule and filter fixes. In the browser, extract a pure summary-markup helper so Started and Complete panels share one tested rendering path while the history table remains unchanged.

**Tech Stack:** PostgreSQL/PLpgSQL, browser JavaScript, CSS, Node.js built-in test runner, Supabase SQL smoke test.

## Global Constraints

- Eligible candidates are ordered before pagination by `source_date ASC, id ASC` for every configured source.
- Candidate JSON aggregation uses the same `source_date ASC, id ASC` ordering.
- The browser does not apply a second candidate sort.
- The 2026-01-01 eligibility floor, filters, source rules, record locks, page size, calculations, locked-record order, audit order, and history order remain unchanged.
- The forward migration targets `public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)`, preserves the rest of its installed definition, fails clearly on an unexpected definition, and is safe to run again after success.
- Current reconciliation renders one compact row in this order: status badge, `#records: N`, `Dif: AMOUNT`.
- `N` is the complete `workspace.items.length` across all locked source types.
- The source-summary sentence is absent for both Started and Complete reconciliations.
- Reconciliation history keeps its Base source and Matching sources columns.
- No API request/response contract or database table changes are introduced.

---

## File structure

- Create `supabase-migrations/2026-08-12-financial-reconciliation-oldest-first-candidates.sql` — idempotently transform and verify only the installed workspace function's candidate ordering.
- Modify `tests/reconciliation.test.js` — statically verify the forward migration's exact signature, deterministic ordering, verification guard, and idempotent branch.
- Modify `tests/reconciliation-rpc.smoke.sql` — apply the new migration inside the disposable transaction and prove oldest-first, same-date ID ordering from the real workspace RPC.
- Modify `app-main.js` — add the compact summary helper, use the full locked-item count, remove the source-summary rendering and its now-unused rule-label lookup.
- Modify `styles.css` — align the three compact summary values on one row and remove the obsolete summary-paragraph rule.
- Modify `tests/reconciliation-density.test.js` — execute the summary helper for Started and Complete fixtures and preserve the history source-column contract.

### Task 1: Deterministic oldest-first workspace migration

**Files:**
- Create: `supabase-migrations/2026-08-12-financial-reconciliation-oldest-first-candidates.sql`
- Modify: `tests/reconciliation.test.js`
- Modify: `tests/reconciliation-rpc.smoke.sql`

**Interfaces:**
- Consumes: installed PostgreSQL function `public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)`.
- Produces: the same function signature and JSON contract, with candidate selection and aggregation ordered by `source_date ASC, id ASC`.
- Produces migration constants for the exact old and new candidate-order fragments; no new RPC or table interface.

- [ ] **Step 1: Write the failing migration-contract test**

At the top of `tests/reconciliation.test.js`, after the existing imports, read the planned migration without turning the initial RED run into a file-loading error:

```js
const oldestFirstMigrationPath = path.join(
  __dirname,
  "..",
  "supabase-migrations",
  "2026-08-12-financial-reconciliation-oldest-first-candidates.sql",
);
const oldestFirstMigration = fs.existsSync(oldestFirstMigrationPath)
  ? fs.readFileSync(oldestFirstMigrationPath, "utf8")
  : "";
```

Add this test with literal expectations:

```js
test("oldest-first migration deterministically orders candidates before and after pagination", () => {
  assert.match(oldestFirstMigration, /pg_get_functiondef\('public\.get_financial_reconciliation_workspace\(uuid,text,jsonb,integer,integer\)'::regprocedure\)/);
  assert.match(oldestFirstMigration, /old_page_order constant text := \$\$order by source_date desc offset v_offset limit p_page_size\$\$/i);
  assert.match(oldestFirstMigration, /new_page_order constant text := \$\$order by source_date asc, id asc offset v_offset limit p_page_size\$\$/i);
  assert.match(oldestFirstMigration, /old_json_order constant text := \$\$order by x\.source_date desc\$\$/i);
  assert.match(oldestFirstMigration, /new_json_order constant text := \$\$order by x\.source_date asc, x\.id asc\$\$/i);
  assert.match(oldestFirstMigration, /if position\(new_page_order in definition\) = 0 or position\(new_json_order in definition\) = 0/i);
  assert.match(oldestFirstMigration, /if position\(old_page_order in definition\) > 0 then[\s\S]*replace\(definition, old_page_order, new_page_order\)/i);
  assert.match(oldestFirstMigration, /if position\(old_json_order in definition\) > 0 then[\s\S]*replace\(definition, old_json_order, new_json_order\)/i);
  assert.match(oldestFirstMigration, /if definition <> original_definition then\s*execute definition;/i);
  assert.match(oldestFirstMigration, /could not verify deterministic oldest-first candidate ordering/i);
  assert.doesNotMatch(oldestFirstMigration, /create table|alter table|drop table/i);
});
```

This catches reintroducing descending page order, omitting the deterministic ID tie-breaker, executing unnecessarily on a second run, or targeting the wrong overload.

- [ ] **Step 2: Run the focused test and verify the red state**

Run: `node --test --test-isolation=none tests/reconciliation.test.js`

Expected: FAIL on the first `assert.match` because the missing migration is represented by an empty string. The test runner itself must load normally; an `ENOENT` setup error is not an acceptable RED state.

- [ ] **Step 3: Create the idempotent forward migration**

Create the migration with this complete structure:

```sql
-- Orders eligible reconciliation candidates from oldest to newest before pagination.
do $fix$
declare
  definition text;
  original_definition text;
  old_page_order constant text := $$order by source_date desc offset v_offset limit p_page_size$$;
  new_page_order constant text := $$order by source_date asc, id asc offset v_offset limit p_page_size$$;
  old_json_order constant text := $$order by x.source_date desc$$;
  new_json_order constant text := $$order by x.source_date asc, x.id asc$$;
begin
  select pg_get_functiondef('public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)'::regprocedure)
    into definition;

  original_definition := definition;

  if position(old_page_order in definition) > 0 then
    definition := replace(definition, old_page_order, new_page_order);
  end if;
  if position(old_json_order in definition) > 0 then
    definition := replace(definition, old_json_order, new_json_order);
  end if;

  if position(new_page_order in definition) = 0
     or position(new_json_order in definition) = 0
     or position(old_page_order in definition) > 0
     or position(old_json_order in definition) > 0 then
    raise exception 'Could not verify deterministic oldest-first candidate ordering in the reconciliation workspace function.';
  end if;

  if definition <> original_definition then
    execute definition;
  end if;
end $fix$;
```

Do not add `notify pgrst`: the function signature and PostgREST schema do not change.

- [ ] **Step 4: Run the migration-contract test and verify green**

Run: `node --test --test-isolation=none tests/reconciliation.test.js`

Expected: all tests in `tests/reconciliation.test.js` PASS.

- [ ] **Step 5: Write the real SQL ordering smoke test**

In `tests/reconciliation-rpc.smoke.sql`, apply the new migration immediately after the existing source-rules migration line:

```sql
\ir ../supabase-migrations/2026-08-12-financial-reconciliation-oldest-first-candidates.sql
```

Extend the `do` block declarations with:

```sql
old_doc_id uuid := '00000000-0000-0000-0000-000000000101';
same_date_low_id uuid := '00000000-0000-0000-0000-000000000102';
same_date_high_id uuid := '00000000-0000-0000-0000-000000000103';
new_doc_id uuid := '00000000-0000-0000-0000-000000000104';
candidate_ids uuid[];
```

Insert four eligible, unlocked Financial Documents after the existing fixtures:

```sql
insert into financial_documents(id,document_date,amount,fat,created_by,description) values
  (old_doc_id,'2026-02-01',1,'S','smoke','oldest ordering fixture'),
  (same_date_high_id,'2026-02-02',1,'S','smoke','same date high id fixture'),
  (same_date_low_id,'2026-02-02',1,'S','smoke','same date low id fixture'),
  (new_doc_id,'2026-02-03',1,'S','smoke','newest ordering fixture');
```

Request a four-record page filtered to those dates and assert the literal order:

```sql
r := get_financial_reconciliation_workspace(
  null,
  'financial_documents',
  '{"dateFrom":"2026-02-01","dateTo":"2026-02-03","amountMin":"1","amountMax":"1","description":"ordering fixture"}'::jsonb,
  1,
  4
);
select array_agg((candidate->>'id')::uuid order by ordinal)
  into candidate_ids
  from jsonb_array_elements(r->'candidates') with ordinality candidates(candidate, ordinal);
if candidate_ids is distinct from array[old_doc_id, same_date_low_id, same_date_high_id, new_doc_id] then
  raise exception 'Candidates were not returned oldest-first with deterministic same-date ordering: %', candidate_ids;
end if;
```

Run the migration a second time before the assertion to prove idempotency in PostgreSQL:

```sql
\ir ../supabase-migrations/2026-08-12-financial-reconciliation-oldest-first-candidates.sql
```

Place the second `\ir` at the top SQL level before the `do` block; psql meta-commands cannot appear inside PL/pgSQL.

- [ ] **Step 6: Execute or explicitly report the SQL smoke environment**

If a disposable Supabase/PostgreSQL connection is configured, run:

```bash
psql "$SUPABASE_DB_URL" -v ON_ERROR_STOP=1 -f tests/reconciliation-rpc.smoke.sql
```

Expected: transaction completes and rolls back without an exception.

If no database connection or `psql` is available, do not claim the SQL was executed. Record this exact verification gap in the task report and rely on the migration source test until the user runs the migration in Supabase.

- [ ] **Step 7: Run focused checks**

Run:

```bash
node --test --test-isolation=none tests/reconciliation.test.js
git diff --check
```

Expected: Node tests PASS and `git diff --check` produces no output.

- [ ] **Step 8: Commit Task 1**

```bash
git add supabase-migrations/2026-08-12-financial-reconciliation-oldest-first-candidates.sql tests/reconciliation.test.js tests/reconciliation-rpc.smoke.sql
git commit -m "feat: order reconciliation candidates oldest first"
```

### Task 2: Compact Current reconciliation summary

**Files:**
- Modify: `app-main.js:21548-21581`
- Modify: `styles.css:6388-6407`
- Modify: `styles.css:6470-6474`
- Modify: `tests/reconciliation-density.test.js`

**Interfaces:**
- Produces: `financialReconciliationSummaryMarkup(status: string, itemCount: number, difference: number): string`.
- Consumes: existing `financialReconciliationStatusMarkup`, `formatMoney`, `escape`, and the complete `workspace.items.length`.
- Preserves: `renderFinancialReconciliationHistory()` including Base source and Matching sources cells.

- [ ] **Step 1: Write the failing executable summary-helper tests**

In `tests/reconciliation-density.test.js`, add this harness after the existing `financialReconciliationItemDetails` harness:

```js
const financialReconciliationSummaryMarkup = new Function(
  "financialReconciliationStatusMarkup",
  "formatMoney",
  "escape",
  `${appFunctionSource("financialReconciliationSummaryMarkup")}\nreturn financialReconciliationSummaryMarkup;`,
)(
  (status) => `<span class="financial-reconciliation-status">${status}</span>`,
  (amount) => `${Number(amount).toFixed(2)} €`,
  (value) => String(value),
);
```

Add literal behavior tests for both lifecycle states and the multi-source total:

```js
test("Current reconciliation summary shows status record count and short difference", () => {
  assert.equal(
    financialReconciliationSummaryMarkup("started", 2, 0),
    '<div class="financial-reconciliation-summary"><span class="financial-reconciliation-status">started</span><strong class="financial-reconciliation-record-count">#records: 2</strong><strong class="financial-reconciliation-difference">Dif: 0.00 €</strong></div>',
  );
  assert.equal(
    financialReconciliationSummaryMarkup("complete", 3, -10),
    '<div class="financial-reconciliation-summary"><span class="financial-reconciliation-status">complete</span><strong class="financial-reconciliation-record-count">#records: 3</strong><strong class="financial-reconciliation-difference financial-reconciliation-forced-difference">Dif: -10.00 €</strong></div>',
  );
});
```

The count fixtures intentionally represent records from different sources; the helper accepts only the already-combined total and therefore cannot count a single source subset.

- [ ] **Step 2: Write the failing integration/source-removal contract test**

Add:

```js
test("Current summary omits source prose while history keeps source columns", () => {
  const currentSource = appFunctionSource("renderFinancialReconciliationCurrent");
  const historySource = appFunctionSource("renderFinancialReconciliationHistory");
  assert.match(currentSource, /financialReconciliationSummaryMarkup\(reconciliation\.status, workspace\.items\.length, difference\)/);
  assert.doesNotMatch(currentSource, /matchingSources|base_source_type\)\)\} with/);
  assert.doesNotMatch(currentSource, />Difference:/);
  assert.match(historySource, /financialReconciliationSourceLabel\(record\.base_source_type\)/);
  assert.match(historySource, /record\.matching_source_types/);
});
```

- [ ] **Step 3: Run the density test and verify the red state**

Run: `node --test --test-isolation=none tests/reconciliation-density.test.js`

Expected: FAIL because `financialReconciliationSummaryMarkup` is undefined.

- [ ] **Step 4: Add the pure summary helper**

Add this immediately before `renderFinancialReconciliationCurrent`:

```js
function financialReconciliationSummaryMarkup(status, itemCount, difference) {
  const numericDifference = Number(difference || 0);
  return `<div class="financial-reconciliation-summary">${financialReconciliationStatusMarkup(status)}<strong class="financial-reconciliation-record-count">#records: ${escape(Number(itemCount) || 0)}</strong><strong class="financial-reconciliation-difference${numericDifference === 0 ? "" : " financial-reconciliation-forced-difference"}">Dif: ${escape(formatMoney(numericDifference))}</strong></div>`;
}
```

- [ ] **Step 5: Use the helper and remove source-summary logic**

In `renderFinancialReconciliationCurrent`:

- delete the `matchingSources` declaration;
- create `const summary = financialReconciliationSummaryMarkup(reconciliation.status, workspace.items.length, difference);`;
- replace the existing summary HTML at the start of `innerHTML` with `${summary}`;
- keep items, completion controls/completed summary, audit, and lifecycle buttons unchanged.

Do not change `renderFinancialReconciliationHistory`.

- [ ] **Step 6: Update compact summary CSS**

Change `.financial-reconciliation-summary` to a single three-value row:

```css
.financial-reconciliation-summary {
  display: grid;
  grid-template-columns: auto 1fr auto;
  gap: .55rem;
  align-items: center;
  padding: .75rem;
  border: 1px solid var(--border);
  border-radius: 12px;
  background: var(--surface-soft, rgba(255,255,255,.55));
}

.financial-reconciliation-record-count {
  color: var(--muted);
  text-align: center;
  white-space: nowrap;
}
```

Delete `.financial-reconciliation-summary p` because the source paragraph no longer exists. Replace the `@media (max-width: 620px)` summary rule with:

```css
.financial-reconciliation-summary {
  grid-template-columns: auto 1fr auto;
  gap: .35rem;
}
```

This keeps `#records` physically between status and `Dif` even at the narrow width shown in the user's reference screenshot. Keep the difference right-aligned and do not affect other Current reconciliation content.

- [ ] **Step 7: Add CSS contract assertions**

In the existing reconciliation density test, add:

```js
assert.match(css, /\.financial-reconciliation-summary\s*\{[\s\S]*grid-template-columns:\s*auto 1fr auto;[\s\S]*align-items:\s*center;/);
assert.match(css, /\.financial-reconciliation-record-count\s*\{[\s\S]*text-align:\s*center;[\s\S]*white-space:\s*nowrap;/);
assert.doesNotMatch(css, /\.financial-reconciliation-summary p\s*\{/);
```

- [ ] **Step 8: Run focused tests and verify green**

Run:

```bash
node --test --test-isolation=none tests/reconciliation-density.test.js tests/reconciliation-inline-completion.test.js
node --check app-main.js
```

Expected: all focused tests PASS and JavaScript syntax is valid.

- [ ] **Step 9: Commit Task 2**

```bash
git add app-main.js styles.css tests/reconciliation-density.test.js
git commit -m "feat: compact current reconciliation summary"
```

### Task 3: Full regression and rollout verification

**Files:**
- Verify only; no planned production edits.
- If a verification defect is found, return to the task that owns the affected file and add a failing regression test before fixing it.

**Interfaces:**
- Consumes: Task 1 migration and Task 2 summary helper.
- Produces: verified feature branch ready for review and integration.

- [ ] **Step 1: Run all automated checks**

Run:

```bash
node --check app-main.js
node --test --test-isolation=none tests/*.test.js
git diff --check
```

Expected: syntax check exits 0, every Node test passes, and diff check produces no output.

- [ ] **Step 2: Verify the exact branch scope**

Run:

```bash
git status --short
git diff --stat main...HEAD
git diff main...HEAD -- app-main.js styles.css tests/reconciliation-density.test.js tests/reconciliation.test.js tests/reconciliation-rpc.smoke.sql supabase-migrations/2026-08-12-financial-reconciliation-oldest-first-candidates.sql
```

Expected: only the six planned implementation files differ from the execution base; existing unrelated untracked files remain untouched.

- [ ] **Step 3: Verify the browser after applying the SQL in a non-production environment**

Apply migrations in this order if the target environment does not already include the earlier reconciliation fixes:

1. `supabase-migrations/2026-08-11-financial-reconciliation-source-rules.sql`
2. `supabase-migrations/2026-08-11-financial-reconciliation-source-rules-workspace-filter-fix.sql`
3. `supabase-migrations/2026-08-12-financial-reconciliation-action-overload-fix.sql`
4. `supabase-migrations/2026-08-12-financial-reconciliation-oldest-first-candidates.sql`

If the first three have already succeeded in that environment, run only step 4.

Then verify:

1. Select Financial Documents and confirm the earliest eligible date is first.
2. Select a second source and confirm the earliest eligible date is first.
3. Confirm equal-date records are consistently ordered by ID across refreshes.
4. Open a Started reconciliation and confirm the row reads status, `#records: N`, `Dif: AMOUNT`, with no source sentence.
5. Open a Complete reconciliation and confirm the same summary contract.
6. Confirm `N` equals the number of locked-record cards across all source types.
7. Confirm Reconciliation history still shows Base source and Matching sources.

- [ ] **Step 4: Record the database execution boundary**

The code can be merged and published without automatically mutating Supabase. In the final handoff, state clearly:

- whether the SQL migration was actually executed against Supabase;
- the exact migration filename;
- that candidate ordering changes only after this migration succeeds;
- that no data backfill is required.

- [ ] **Step 5: Final commit check**

Run: `git log --oneline main..HEAD`

Expected: the Task 1 and Task 2 commits are present, with no uncommitted planned files.
