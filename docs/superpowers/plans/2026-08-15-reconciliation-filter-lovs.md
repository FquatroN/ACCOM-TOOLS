# Manual Reconciliation Filter LOVs Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add source-specific LOV filters to Manual Reconciliation, make Financial Documents supplier search cover name and NIF, and remove its Account filter.

**Architecture:** Keep the existing `/api/reconciliation` request and `get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)` RPC. A new idempotent migration enriches `sourceConfig` with complete-table `filterOptions` while preserving every existing workspace enrichment; the frontend renders a `<select>` whenever the selected source advertises an option array.

**Tech Stack:** Vanilla JavaScript, Node.js built-in test runner, PostgreSQL/PLpgSQL, Supabase PostgREST RPC, HTML/CSS already present in the application.

## Global Constraints

- The workspace RPC signature must remain `get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)`.
- Financial Documents filters are exactly `dateFrom`, `dateTo`, `amountMin`, `amountMax`, `description`, `supplier`, `payment`, and `category` in that order.
- FDM Accounts filters are exactly `dateFrom`, `dateTo`, `amountMin`, `amountMax`, `description`, `account`, and `category` in that order.
- Financial Documents Payment and Category and FDM Accounts Account and Category are single-select LOVs with an `All ...` option.
- LOV data comes from all distinct, trimmed, nonblank values in the complete relevant source table; it is not limited by paging, active filters, locks, or eligibility.
- Supplier Search is one case-insensitive partial-match field covering both supplier name and supplier NIF.
- CGD Credit Card and CGD Bank Statement filter behavior remains unchanged.
- Existing authorization, minimum date `2026-01-01`, `fat = 'S'`, FDM eligibility, locks, oldest-first ordering, counts, history summaries, current-item details, and automatic-reconciliation provenance must remain unchanged.
- No new API endpoint or direct browser database access.
- Use strict RED/GREEN TDD for every production change.

## File Structure

- Create `supabase-migrations/2026-08-15-financial-reconciliation-workspace-filter-lovs.sql`: idempotently patch the latest workspace RPC with LOV metadata and supplier name/NIF matching.
- Modify `tests/reconciliation.test.js`: pin the migration contract, security, source-specific fields, trimmed exact LOV predicates, and supplier predicate.
- Modify `tests/reconciliation-rpc.smoke.sql`: reapply the migration and exercise LOV/supplier behavior against controlled database fixtures.
- Modify `app-main.js`: normalize workspace option metadata, render LOV selects, rename Supplier Search, remove the Financial Documents Account fallback, and prevent duplicate select reloads.
- Modify `tests/reconciliation-density.test.js`: execute the real frontend helpers and verify markup, escaping, source-specific request keys, and event behavior.

---

### Task 1: Enrich the workspace RPC with source-specific LOV metadata

**Files:**
- Create: `supabase-migrations/2026-08-15-financial-reconciliation-workspace-filter-lovs.sql`
- Modify: `tests/reconciliation.test.js`
- Modify: `tests/reconciliation-rpc.smoke.sql`

**Interfaces:**
- Consumes: `public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)` and its existing `sourceConfig` JSON object.
- Produces: `sourceConfig.filterOptions: Record<string,string[]>`; updated `sourceConfig.filterFields`; supplier filtering across `s.supplier` and `s.supplier_nif`.
- Preserves: the RPC signature, result envelope, security-definer search path, and service-role-only execution.

- [ ] **Step 1: Add the failing migration contract test**

In `tests/reconciliation.test.js`, define and load the new migration next to the existing reconciliation migration constants:

```js
const filterLovMigrationPath = path.join(
  __dirname,
  "..",
  "supabase-migrations",
  "2026-08-15-financial-reconciliation-workspace-filter-lovs.sql",
);
const filterLovMigration = fs.existsSync(filterLovMigrationPath)
  ? fs.readFileSync(filterLovMigrationPath, "utf8")
  : "";
```

Add a test named `workspace filter LOV migration preserves the RPC while adding source metadata` that asserts:

```js
assert.match(filterLovMigration, /pg_get_functiondef\('public\.get_financial_reconciliation_workspace\(uuid,text,jsonb,integer,integer\)'::regprocedure\)/);
assert.match(filterLovMigration, /'filterOptions',v_filter_options/);
assert.match(filterLovMigration, /jsonb_build_array\('dateFrom','dateTo','amountMin','amountMax','description','supplier','payment','category'\)/);
assert.doesNotMatch(
  filterLovMigration,
  /jsonb_build_array\('dateFrom','dateTo','amountMin','amountMax','description','supplier','payment','account','category'\)/,
);
assert.match(filterLovMigration, /s\.supplier ilike[\s\S]+or s\.supplier_nif ilike/);
assert.match(filterLovMigration, /select distinct btrim\(payment\)/i);
assert.match(filterLovMigration, /select distinct btrim\(category\)/i);
assert.match(filterLovMigration, /select distinct btrim\(account\)/i);
assert.match(filterLovMigration, /order by lower\(option_value\),option_value/i);
assert.match(filterLovMigration, /btrim\(s\.payment\) = p_filters->>'payment'/);
assert.match(filterLovMigration, /btrim\(s\.account\) = p_filters->>'account'/);
assert.match(filterLovMigration, /btrim\(s\.category\) = p_filters->>'category'/);
assert.match(filterLovMigration, /security definer/i);
assert.match(filterLovMigration, /revoke all on function public\.get_financial_reconciliation_workspace\(uuid,text,jsonb,integer,integer\) from public, anon, authenticated;/);
assert.match(filterLovMigration, /grant execute on function public\.get_financial_reconciliation_workspace\(uuid,text,jsonb,integer,integer\) to service_role;/);
```

The mutation this catches is reinstalling a workspace definition that exposes Account for Financial Documents, searches supplier name only, or omits complete-table LOV metadata.

- [ ] **Step 2: Run the focused test and verify RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation.test.js
```

Expected: FAIL because `2026-08-15-financial-reconciliation-workspace-filter-lovs.sql` does not exist.

- [ ] **Step 3: Add failing SQL smoke coverage**

In `tests/reconciliation-rpc.smoke.sql`, apply the new migration twice immediately after the automation execution migration so it always patches the latest provenance-enriched workspace function:

```sql
\ir ../supabase-migrations/2026-08-15-financial-reconciliation-workspace-filter-lovs.sql
\ir ../supabase-migrations/2026-08-15-financial-reconciliation-workspace-filter-lovs.sql
```

Extend the existing smoke transaction with unique fixtures. Use the current transaction identifier in values to prevent collisions with pre-existing rows:

```sql
insert into financial_documents(
  id, document_date, amount, fat, created_by, description,
  supplier_name, supplier_nif, payment, category
) values
  (gen_random_uuid(), '2026-03-01', 11, 'S', 'smoke', 'lov supplier-name fixture',
   'LOV Supplier Name', 'PT-LOV-NIF', '  LOV Visa  ', '  LOV Food  '),
  (gen_random_uuid(), '2026-03-02', 12, 'S', 'smoke', 'lov supplier-nif fixture',
   'Different Supplier', 'PT-LOV-SEARCH', 'LOV Visa', 'LOV Food'),
  (gen_random_uuid(), '2026-03-03', 13, 'S', 'smoke', 'lov blank fixture',
   'Blank Supplier', 'PT-BLANK', '   ', '');

insert into import_fdm_accounts(
  id, import_batch, account, date_time_raw, event_date,
  category, amount, invoice_flag, description
) values
  (gen_random_uuid(), 'smoke', '  LOV Main Account  ', '2026-03-01', '2026-03-01',
   '  LOV Purchases  ', -11, true, 'lov fdm fixture'),
  (gen_random_uuid(), 'smoke', 'LOV Main Account', '2026-03-02', '2026-03-02',
   'LOV Purchases', -12, true, 'lov fdm duplicate fixture'),
  (gen_random_uuid(), 'smoke', ' ', '2026-03-03', '2026-03-03',
   '', -13, true, 'lov fdm blank fixture');
```

Call the workspace RPC for Financial Documents and assert the literal contract:

```sql
r := public.get_financial_reconciliation_workspace(
  null, 'financial_documents',
  '{"dateFrom":"2026-03-01","dateTo":"2026-03-03","supplier":"PT-LOV-SEARCH"}'::jsonb,
  1, 50
);
if r->'sourceConfig'->'filterFields' is distinct from
   '["dateFrom","dateTo","amountMin","amountMax","description","supplier","payment","category"]'::jsonb then
  raise exception 'Financial Documents filter fields are invalid: %', r->'sourceConfig'->'filterFields';
end if;
if (select count(*) from jsonb_array_elements_text(r->'sourceConfig'->'filterOptions'->'payment') value where value = 'LOV Visa') <> 1
   or (r->'sourceConfig'->'filterOptions'->'payment') @> '[""]'::jsonb then
  raise exception 'Financial Documents payment LOV is not trimmed, distinct, and nonblank.';
end if;
if jsonb_array_length(r->'candidates') <> 1
   or r->'candidates'->0->>'supplier_nif' <> 'PT-LOV-SEARCH' then
  raise exception 'Supplier Search did not match Supplier NIF.';
end if;
```

Prove supplier-name matching separately:

```sql
r := public.get_financial_reconciliation_workspace(
  null, 'financial_documents',
  '{"dateFrom":"2026-03-01","dateTo":"2026-03-03","supplier":"LOV Supplier Name"}'::jsonb,
  1, 50
);
if jsonb_array_length(r->'candidates') <> 1
   or r->'candidates'->0->>'supplier' <> 'LOV Supplier Name' then
  raise exception 'Supplier Search did not match Supplier Name.';
end if;
```

Call the FDM workspace and verify the complete arrays against independent SQL results:

```sql
r := public.get_financial_reconciliation_workspace(
  null, 'import_fdm_accounts',
  '{"dateFrom":"2026-03-01","dateTo":"2026-03-03"}'::jsonb,
  1, 50
);
if r->'sourceConfig'->'filterFields' is distinct from
   '["dateFrom","dateTo","amountMin","amountMax","description","account","category"]'::jsonb then
  raise exception 'FDM filter fields are invalid: %', r->'sourceConfig'->'filterFields';
end if;
if r->'sourceConfig'->'filterOptions'->'account' is distinct from (
  select coalesce(jsonb_agg(option_value order by lower(option_value), option_value), '[]'::jsonb)
  from (
    select distinct btrim(account) option_value
    from public.import_fdm_accounts
    where nullif(btrim(account), '') is not null
  ) options
) then
  raise exception 'FDM Account LOV ordering or values are invalid.';
end if;
if r->'sourceConfig'->'filterOptions'->'category' is distinct from (
  select coalesce(jsonb_agg(option_value order by lower(option_value), option_value), '[]'::jsonb)
  from (
    select distinct btrim(category) option_value
    from public.import_fdm_accounts
    where nullif(btrim(category), '') is not null
  ) options
) then
  raise exception 'FDM Category LOV ordering or values are invalid.';
end if;
if (select count(*) from jsonb_array_elements_text(r->'sourceConfig'->'filterOptions'->'account') value where value = 'LOV Main Account') <> 1
   or exists (
     select 1
     from jsonb_array_elements_text(r->'sourceConfig'->'filterOptions'->'account') value
     where nullif(btrim(value), '') is null
   ) then
  raise exception 'FDM Account LOV is not trimmed, distinct, and nonblank.';
end if;
```

Verify the complete Financial Documents arrays independently as well:

```sql
r := public.get_financial_reconciliation_workspace(
  null, 'financial_documents', '{}'::jsonb, 1, 50
);
if r->'sourceConfig'->'filterOptions'->'payment' is distinct from (
  select coalesce(jsonb_agg(option_value order by lower(option_value), option_value), '[]'::jsonb)
  from (
    select distinct btrim(payment) option_value
    from public.financial_documents
    where nullif(btrim(payment), '') is not null
  ) options
) then
  raise exception 'Financial Documents Payment LOV ordering or values are invalid.';
end if;
if r->'sourceConfig'->'filterOptions'->'category' is distinct from (
  select coalesce(jsonb_agg(option_value order by lower(option_value), option_value), '[]'::jsonb)
  from (
    select distinct btrim(category) option_value
    from public.financial_documents
    where nullif(btrim(category), '') is not null
  ) options
) then
  raise exception 'Financial Documents Category LOV ordering or values are invalid.';
end if;
```

- [ ] **Step 4: Extend the source-level smoke contract and verify RED**

Add assertions to `tests/reconciliation.test.js` confirming that `tests/reconciliation-rpc.smoke.sql` includes the new migration twice and contains the name, NIF, distinct-value, blank-value, and deterministic-order assertions:

```js
const rpcSmoke = fs.readFileSync(path.join(__dirname, "reconciliation-rpc.smoke.sql"), "utf8");
assert.equal(
  (rpcSmoke.match(/2026-08-15-financial-reconciliation-workspace-filter-lovs\.sql/g) || []).length,
  2,
);
assert.match(rpcSmoke, /Supplier Search did not match Supplier NIF/);
assert.match(rpcSmoke, /Supplier Search did not match Supplier Name/);
assert.match(rpcSmoke, /payment LOV is not trimmed, distinct, and nonblank/);
assert.match(rpcSmoke, /FDM Account LOV is not trimmed, distinct, and nonblank/);
```

Run:

```powershell
node --test --test-isolation=none tests/reconciliation.test.js
```

Expected: FAIL until the complete smoke assertions and migration are present.

- [ ] **Step 5: Implement the idempotent workspace migration**

Create `supabase-migrations/2026-08-15-financial-reconciliation-workspace-filter-lovs.sql` as a guarded `DO` block. Read the installed function using:

```sql
select pg_get_functiondef(
  'public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)'::regprocedure
) into strict definition;
```

Patch the existing definition instead of copying an older full function. This preserves history summaries, oldest-first ordering, item details, and automatic provenance. Use exact old/new snippet counts and accept only these two states for every replacement: old snippet present at its expected count and new absent, or old absent and new present at its expected count. Raise `Unexpected reconciliation workspace function definition; could not install filter LOVs.` for every other state.

Add `v_filter_options jsonb := '{}'::jsonb` to the declaration and inject the following after `v_offset := (p_page - 1) * p_page_size;`:

```sql
if p_source_type = 'financial_documents' then
  select jsonb_build_object(
    'payment', coalesce((
      select jsonb_agg(option_value order by lower(option_value), option_value)
      from (
        select distinct btrim(payment) as option_value
        from public.financial_documents
        where nullif(btrim(payment), '') is not null
      ) options
    ), '[]'::jsonb),
    'category', coalesce((
      select jsonb_agg(option_value order by lower(option_value), option_value)
      from (
        select distinct btrim(category) as option_value
        from public.financial_documents
        where nullif(btrim(category), '') is not null
      ) options
    ), '[]'::jsonb)
  ) into v_filter_options;
elsif p_source_type = 'import_fdm_accounts' then
  select jsonb_build_object(
    'account', coalesce((
      select jsonb_agg(option_value order by lower(option_value), option_value)
      from (
        select distinct btrim(account) as option_value
        from public.import_fdm_accounts
        where nullif(btrim(account), '') is not null
      ) options
    ), '[]'::jsonb),
    'category', coalesce((
      select jsonb_agg(option_value order by lower(option_value), option_value)
      from (
        select distinct btrim(category) as option_value
        from public.import_fdm_accounts
        where nullif(btrim(category), '') is not null
      ) options
    ), '[]'::jsonb)
  ) into v_filter_options;
end if;
```

Replace both name-only supplier predicates with this parenthesized predicate:

```sql
and (
  nullif(p_filters->>'supplier','') is null
  or s.supplier ilike '%' || (p_filters->>'supplier') || '%'
  or s.supplier_nif ilike '%' || (p_filters->>'supplier') || '%'
)
```

Replace both Payment, Account, and Category exact predicates so the trimmed LOV value matches trimmed stored data:

```sql
and (nullif(p_filters->>'payment','') is null or btrim(s.payment) = p_filters->>'payment')
and (nullif(p_filters->>'account','') is null or btrim(s.account) = p_filters->>'account')
and (nullif(p_filters->>'category','') is null or btrim(s.category) = p_filters->>'category')
```

Replace the Financial Documents `filterFields` array with the approved eight fields and inject `'filterOptions',v_filter_options` inside `sourceConfig` immediately after `'amountColumn','amount'`.

Execute the patched function only when `definition <> original_definition`. Finish with the existing permission boundary:

```sql
revoke all on function public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)
  from public, anon, authenticated;
grant execute on function public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)
  to service_role;
notify pgrst, 'reload schema';
```

- [ ] **Step 6: Run focused tests and SQL smoke**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation.test.js
```

Expected: all reconciliation tests PASS.

When `psql` and `SUPABASE_DB_URL` are available, run:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-rpc.smoke.sql
```

Expected: exit 0 and the transaction rolls back. If the local database connection is unavailable, record this as a required external verification gate; do not claim the SQL smoke passed.

- [ ] **Step 7: Commit Task 1**

```powershell
git add -- 'supabase-migrations/2026-08-15-financial-reconciliation-workspace-filter-lovs.sql' 'tests/reconciliation.test.js' 'tests/reconciliation-rpc.smoke.sql'
git commit -m "feat: add reconciliation filter LOV metadata"
```

---

### Task 2: Render safe source-specific LOV controls in Manual Reconciliation

**Files:**
- Modify: `app-main.js:21258-21263`
- Modify: `app-main.js:21870-21925`
- Modify: `app-main.js:22304-22323`
- Modify: `app-main.js:22553-22563`
- Modify: `tests/reconciliation-density.test.js`

**Interfaces:**
- Consumes: `workspace.sourceConfig.filterOptions: Record<string,string[]>` from Task 1.
- Produces: `financialReconciliationFilterOptions(workspace): Record<string,string[]>` and `financialReconciliationFilterFieldMarkup(field,value,optionValues): string`.
- Preserves: `data-financial-reconciliation-filter`, `currentFinancialReconciliationFilters()`, pagination reset, and the existing workspace reload path.

- [ ] **Step 1: Add failing executable frontend tests**

In `tests/reconciliation-density.test.js`, extract the actual new helpers with `appFunctionSource`. Add a test named `manual reconciliation renders source LOVs and escapes every option` using literal expectations:

```js
const markup = financialReconciliationFilterFieldMarkup(
  "payment",
  "Visa",
  ["Banco", "Visa", '<script data-x="1">'],
);
assert.match(markup, /^<label class="financial-reconciliation-filter-payment">Payment<select/);
assert.match(markup, /<option value="">All payments<\/option>/);
assert.match(markup, /<option value="Visa" selected>Visa<\/option>/);
assert.match(markup, /&lt;script data-x=&quot;1&quot;&gt;/);
assert.doesNotMatch(markup, /<script/);
```

Cover all four LOV labels with a literal table:

```js
for (const [field, allLabel] of [
  ["payment", "All payments"],
  ["category", "All categories"],
  ["account", "All accounts"],
]) {
  assert.match(
    financialReconciliationFilterFieldMarkup(field, "", []),
    new RegExp(`<option value="">${allLabel}<\\/option>`),
  );
}
```

Add a separate test proving `supplier` remains `<input type="search">`, is labeled `Supplier Search`, and never becomes a select when no option array is provided.

Add a normalization test:

```js
assert.deepEqual(financialReconciliationFilterOptions({
  sourceConfig: {
    filterOptions: {
      payment: [" Visa ", "", "Visa", null, "Banco"],
      category: "not-an-array",
    },
  },
}), { payment: ["Visa", "Banco"] });
assert.deepEqual(financialReconciliationFilterOptions({ sourceConfig: {} }), {});
```

The mutation these tests catch is rendering free-text controls for LOV fields, trusting malformed option metadata, or inserting unescaped option values.

- [ ] **Step 2: Add failing request-key and event tests**

Compile the actual `financialReconciliationRequestFilters` with controlled state and assert the removed Account value is not sent for Financial Documents:

```js
assert.deepEqual(requestFilters("financial_documents"), {
  dateFrom: "2026-01-01",
  dateTo: "",
  amountMin: "",
  amountMax: "",
  description: "",
  supplier: "",
  payment: "Visa",
  category: "Food",
});
```

Use a state fixture containing `account: "stale-account"` so the assertion proves stale Account is dropped. Add the inverse FDM assertion showing Account and Category are included.

Compile `onFinancialReconciliationFilterInput(event)` and use a fake timer counter. Assert a select event schedules zero delayed reloads while a search input schedules exactly one. The select's `change` event remains handled by `onFinancialReconciliationFilterChange`, so this prevents one selection from dispatching twice.

- [ ] **Step 3: Run the focused UI test and verify RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-density.test.js
```

Expected: FAIL because the helpers do not exist, Financial Documents still includes Account, and all dynamic fields render as inputs.

- [ ] **Step 4: Implement option normalization and approved fallback fields**

In `app-main.js`, change the Financial Documents fallback to:

```js
financial_documents: ["dateFrom", "dateTo", "amountMin", "amountMax", "description", "supplier", "payment", "category"],
```

Add this defensive normalizer near the existing workspace helpers:

```js
function financialReconciliationFilterOptions(workspace) {
  const raw = workspace?.sourceConfig?.filterOptions;
  if (!raw || Array.isArray(raw) || typeof raw !== "object") return {};
  return Object.fromEntries(Object.entries(raw).flatMap(([field, values]) => {
    if (!Array.isArray(values)) return [];
    const seen = new Set();
    const normalized = values.reduce((result, value) => {
      const option = clean(value);
      if (!option || seen.has(option)) return result;
      seen.add(option);
      result.push(option);
      return result;
    }, []);
    return [[field, normalized]];
  }));
}
```

Do not resort the server response in the browser; Task 1 owns deterministic ordering.

- [ ] **Step 5: Implement the pure field renderer and integrate it**

Add a pure helper immediately before `renderFinancialReconciliationFilters`:

```js
function financialReconciliationFilterFieldMarkup(field, value, optionValues = null) {
  const labels = {
    dateFrom: "Date from", dateTo: "Date to", amountMin: "Amount from",
    amountMax: "Amount to", description: "Description", supplier: "Supplier Search",
    payment: "Payment", account: "Account", category: "Category",
  };
  const label = labels[field] || field;
  const normalizedValue = clean(value);
  if (Array.isArray(optionValues)) {
    const allLabels = { payment: "All payments", account: "All accounts", category: "All categories" };
    const options = [
      `<option value="">${escape(allLabels[field] || "All")}</option>`,
      ...optionValues.map((option) => {
        const selected = option === normalizedValue ? " selected" : "";
        return `<option value="${escape(option)}"${selected}>${escape(option)}</option>`;
      }),
    ];
    return `<label class="financial-reconciliation-filter-${escape(field)}">${escape(label)}<select data-financial-reconciliation-filter="${escape(field)}">${options.join("")}</select></label>`;
  }
  const isDate = field === "dateFrom" || field === "dateTo";
  const isAmount = field === "amountMin" || field === "amountMax";
  const type = isDate ? "date" : isAmount ? "number" : "search";
  const min = isDate ? ' min="2026-01-01"' : "";
  const step = isAmount ? ' step="0.01"' : "";
  const placeholder = field === "description" ? ' placeholder="Search description"' : "";
  const renderedValue = field === "dateFrom" ? (normalizedValue || "2026-01-01") : normalizedValue;
  return `<label class="financial-reconciliation-filter-${escape(field)}">${escape(label)}<input type="${type}" data-financial-reconciliation-filter="${escape(field)}" value="${escape(renderedValue)}"${min}${step}${placeholder} /></label>`;
}
```

Refactor `renderFinancialReconciliationFilters()` to obtain `const filterOptions = financialReconciliationFilterOptions(workspace)` and render each field with:

```js
const hasOptions = Object.prototype.hasOwnProperty.call(filterOptions, field);
return financialReconciliationFilterFieldMarkup(
  field,
  filters[field],
  hasOptions ? filterOptions[field] : null,
);
```

This `hasOwnProperty` check is required so an empty LOV still renders a select with its `All ...` option.

- [ ] **Step 6: Prevent duplicate select reloads**

Change the input handler to accept the event and leave selects to the existing `change` handler:

```js
function onFinancialReconciliationFilterInput(event) {
  if (clean(event?.target?.tagName).toLowerCase() === "select") return;
  window.clearTimeout(financialReconciliationFilterTimer);
  financialReconciliationFilterTimer = window.setTimeout(onFinancialReconciliationFilterChange, 250);
}
```

- [ ] **Step 7: Run focused frontend tests and verify GREEN**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-density.test.js
node --test --test-isolation=none tests/reconciliation.test.js
```

Expected: both focused suites PASS with no failures.

- [ ] **Step 8: Commit Task 2**

```powershell
git add -- 'app-main.js' 'tests/reconciliation-density.test.js'
git commit -m "feat: render reconciliation filter LOVs"
```

---

## Final Verification and Handoff

- [ ] Run syntax checks:

```powershell
node --check app-main.js
node --check api/reconciliation.js
node --check api/_reconciliation.js
```

- [ ] Run the complete Node regression suite:

```powershell
node --test --test-isolation=none
```

Expected: all tests pass with zero failures, skips, cancellations, or todos.

- [ ] Check patch hygiene and scope:

```powershell
git diff --check main..HEAD
git status --short
git diff --stat main..HEAD
```

Expected tracked product scope: the new migration, `app-main.js`, `tests/reconciliation.test.js`, and `tests/reconciliation-density.test.js`, plus the SQL smoke update and approved design/plan documents. The worktree must be clean.

- [ ] If PostgreSQL access is available, run `tests/reconciliation-rpc.smoke.sql` and record the actual result. Otherwise state that applying the new migration and running the SQL smoke in Supabase remain mandatory external gates.

- [ ] Request an independent code/spec review before merging or publishing. Review the migration's idempotency counts, preservation of the latest workspace enrichments, supplier predicate parentheses, trimmed LOV equality, HTML escaping, empty-LOV rendering, and double-dispatch prevention.
