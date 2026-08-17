# Financial Reconciliation Combined Search Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the separate Financial Documents description and supplier filters with one combined field that matches description, supplier name, or supplier NIF.

**Architecture:** Keep the existing `description` request key and broaden only its Financial Documents database predicate. The workspace RPC remains authoritative; the client changes only source-specific filter metadata, presentation, and request-field selection.

**Tech Stack:** Vanilla JavaScript, Node.js test runner, PostgreSQL/PLpgSQL, Supabase RPC.

## Global Constraints

- The field label is exactly **Description / Supplier Search** for Financial Documents.
- Matching is case-insensitive partial matching with OR semantics across description, supplier name, and supplier NIF.
- Other reconciliation sources retain their current description-filter behavior.
- Payment and Category remain LOV filters.
- The migration is re-runnable, fail-closed on an unknown workspace-function definition, and preserves service-role-only RPC execution.
- Do not change reconciliation eligibility, locking, paging, ordering, or lifecycle behavior.

---

### Task 1: Combined Financial Documents search

**Files:**
- Modify: `app-main.js:21301-21306`
- Modify: `app-main.js:22546-22591`
- Create: `supabase-migrations/2026-08-17-financial-reconciliation-combined-search.sql`
- Modify: `tests/reconciliation-density.test.js:893-1035`
- Modify: `tests/reconciliation.test.js:1-80`
- Modify: `tests/reconciliation-rpc.smoke.sql:20-27`
- Modify: `tests/reconciliation-rpc.smoke.sql:230-275`

**Interfaces:**
- Consumes: `financialReconciliationFilterFieldMarkup(field, value, optionValues, sourceType)` and `get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)`.
- Produces: Financial Documents `sourceConfig.filterFields` without `supplier`; a source-aware combined-search label; an OR-based Financial Documents `description` predicate.

- [ ] **Step 1: Write failing client and migration-contract tests**

Update `tests/reconciliation-density.test.js` so the real extracted renderer proves the Financial Documents field is combined while other sources keep the ordinary label:

```js
test("manual reconciliation combines Financial Documents description and supplier search", () => {
  const render = extractedFinancialReconciliationFilterFieldMarkup();
  const combined = render("description", "EDP", null, "financial_documents");
  const ordinary = render("description", "guest", null, "import_fdm_accounts");

  assert.match(combined, /^<label class="financial-reconciliation-filter-description">Description \/ Supplier Search<input type="search"/);
  assert.match(combined, /placeholder="Search description or supplier"/);
  assert.match(ordinary, /^<label class="financial-reconciliation-filter-description">Description<input type="search"/);
});
```

Update the existing request-filter fixture so Financial Documents expects:

```js
{
  dateFrom: "2026-01-01",
  dateTo: "",
  amountMin: "",
  amountMax: "",
  description: "EDP",
  payment: "Visa",
  category: "Food",
}
```

and assert the rendered Financial Documents filter row contains the combined description field and no `financial-reconciliation-filter-supplier` element.

In `tests/reconciliation.test.js`, load `2026-08-17-financial-reconciliation-combined-search.sql` and assert it:

```js
assert.match(combinedSearchMigration, /pg_get_functiondef\('public\.get_financial_reconciliation_workspace\(uuid,text,jsonb,integer,integer\)'::regprocedure\)/);
assert.match(combinedSearchMigration, /p_source_type = 'financial_documents'[\s\S]*s\.supplier ilike[\s\S]*s\.supplier_nif ilike/);
assert.match(combinedSearchMigration, /jsonb_build_array\('dateFrom','dateTo','amountMin','amountMax','description','payment','category'\)/);
assert.doesNotMatch(combinedSearchMigration, /jsonb_build_array\('dateFrom','dateTo','amountMin','amountMax','description','supplier','payment','category'\)/);
```

- [ ] **Step 2: Extend the SQL smoke test before production SQL exists**

Include the new migration twice after the existing filter-LOV migration to prove reapplication:

```sql
\ir ../supabase-migrations/2026-08-17-financial-reconciliation-combined-search.sql
\ir ../supabase-migrations/2026-08-17-financial-reconciliation-combined-search.sql
```

Change the Financial Documents `filterFields` expectation to:

```sql
'["dateFrom","dateTo","amountMin","amountMax","description","payment","category"]'::jsonb
```

Call the workspace with the same `description` key four times and assert independent matching by description, supplier name, supplier NIF, and no match:

```sql
r := public.get_financial_reconciliation_workspace(
  null, 'financial_documents',
  '{"dateFrom":"2026-03-01","dateTo":"2026-03-03","description":"lov supplier-name fixture"}'::jsonb,
  1, 50
);
if jsonb_array_length(r->'candidates') <> 1
   or r->'candidates'->0->>'description' <> 'lov supplier-name fixture' then
  raise exception 'Combined Search did not match Description.';
end if;

r := public.get_financial_reconciliation_workspace(
  null, 'financial_documents',
  '{"dateFrom":"2026-03-01","dateTo":"2026-03-03","description":"LOV Supplier Name"}'::jsonb,
  1, 50
);
if jsonb_array_length(r->'candidates') <> 1
   or r->'candidates'->0->>'supplier' <> 'LOV Supplier Name' then
  raise exception 'Combined Search did not match Supplier Name.';
end if;

r := public.get_financial_reconciliation_workspace(
  null, 'financial_documents',
  '{"dateFrom":"2026-03-01","dateTo":"2026-03-03","description":"PT-LOV-SEARCH"}'::jsonb,
  1, 50
);
if jsonb_array_length(r->'candidates') <> 1
   or r->'candidates'->0->>'supplier_nif' <> 'PT-LOV-SEARCH' then
  raise exception 'Combined Search did not match Supplier NIF.';
end if;

r := public.get_financial_reconciliation_workspace(
  null, 'financial_documents',
  '{"dateFrom":"2026-03-01","dateTo":"2026-03-03","description":"NO-COMBINED-SEARCH-MATCH"}'::jsonb,
  1, 50
);
if jsonb_array_length(r->'candidates') <> 0 then
  raise exception 'Combined Search returned unrelated candidates.';
end if;
```

- [ ] **Step 3: Run focused tests and verify RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-density.test.js tests/reconciliation.test.js
```

Expected: failures show the old Supplier Search field, the old Financial Documents request shape, and the missing combined-search migration.

- [ ] **Step 4: Implement the minimal client change**

Remove `supplier` only from the Financial Documents fallback declaration:

```js
financial_documents: ["dateFrom", "dateTo", "amountMin", "amountMax", "description", "payment", "category"],
```

Extend `financialReconciliationFilterFieldMarkup` with `sourceType = ""`. When `field === "description"` and `sourceType === "financial_documents"`, use:

```js
const combinedSearch = field === "description" && clean(sourceType) === "financial_documents";
const label = combinedSearch ? "Description / Supplier Search" : labels[field] || field;
const placeholder = combinedSearch
  ? ' placeholder="Search description or supplier"'
  : field === "description" ? ' placeholder="Search description"' : "";
```

Pass `workspace.sourceConfig.sourceType || current.candidateSourceType` from `renderFinancialReconciliationFilters` into the renderer. Do not alter LOV handling.

- [ ] **Step 5: Add the re-runnable workspace migration**

Create `supabase-migrations/2026-08-17-financial-reconciliation-combined-search.sql`. Read the installed workspace function with `pg_get_functiondef`, then replace exactly two recognized description predicates:

```sql
and (
  nullif(p_filters->>'description','') is null
  or s.description ilike '%' || (p_filters->>'description') || '%'
  or (
    p_source_type = 'financial_documents'
    and (
      s.supplier ilike '%' || (p_filters->>'description') || '%'
      or s.supplier_nif ilike '%' || (p_filters->>'description') || '%'
    )
  )
)
```

Replace exactly one Financial Documents filter-field array with:

```sql
jsonb_build_array('dateFrom','dateTo','amountMin','amountMax','description','payment','category')
```

Accept either the complete old state or complete new state on reapply; reject mixed/unrecognized occurrence counts. Execute the reconstructed definition only when it changed, then reapply:

```sql
revoke all on function public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer) from public, anon, authenticated;
grant execute on function public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer) to service_role;
notify pgrst, 'reload schema';
```

Keep the legacy `supplier` predicate in the function for compatibility with already-open old clients, even though the new UI no longer advertises or sends it.

- [ ] **Step 6: Run focused tests and verify GREEN**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-density.test.js tests/reconciliation.test.js
```

Expected: all focused tests pass.

- [ ] **Step 7: Run full verification**

Run:

```powershell
node --test --test-isolation=none
git diff --check
```

If PostgreSQL is available, also run:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-rpc.smoke.sql
```

Expected: all Node tests pass; diff check is clean; SQL smoke passes when the database gate is available. If PostgreSQL or `SUPABASE_DB_URL` is unavailable, report that exact external gate instead of claiming the SQL ran.

- [ ] **Step 8: Review and commit**

Request a read-only code/spec review of the complete task diff. Fix every Critical or Important finding, rerun Step 7, then commit only the scoped files:

```powershell
git add -- app-main.js supabase-migrations/2026-08-17-financial-reconciliation-combined-search.sql tests/reconciliation-density.test.js tests/reconciliation.test.js tests/reconciliation-rpc.smoke.sql
git commit -m "feat: combine reconciliation text search"
```
