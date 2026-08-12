# Reconciliation Workbench and History Source Summary Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Place Source and its rule hint on the first workbench row, place all source-specific filters on a compact second row, and replace the history Base source/Matching sources columns with one record-derived Source summary.

**Architecture:** Keep the existing single reconciliation workspace request. A forward, idempotent Supabase migration enriches each returned history row with an ordered `sourceSummary` array derived from `financial_reconciliation_items.amount_snapshot`; focused client helpers validate and format that payload without re-aggregating source tables. Scoped HTML/CSS creates explicit source and filter rows while retaining responsive wrapping.

**Tech Stack:** Vanilla HTML/CSS/JavaScript, Node.js `node:test`, PostgreSQL/PLpgSQL, Supabase RPC/PostgREST.

## Global Constraints

- The Source selector and existing reconciliation-rule hint occupy the first workbench row.
- All filters applicable to the selected source occupy the second row on a wide desktop display.
- Description receives flexible remaining width; Date from, Date to, Amount from, Amount to, Supplier, Payment, Account, and Category are narrower.
- Narrow viewports may wrap filters; existing 16px mobile input sizing and usable touch targets remain intact.
- Reconciliation history columns are exactly Created, Source, Status, Difference, and Open action.
- Source summaries use `financial_reconciliation_items.amount_snapshot` raw values; never apply matching-rule `+` or `-` operators.
- Source text uses `<Source label> (#<record count>; <raw amount total>)`, separated by comma-space.
- A used base source appears first, followed by used matching sources in saved `matching_source_types` order.
- Configured sources with no records are omitted.
- A present empty summary renders `No records`; a missing or malformed summary container renders `Source details unavailable`.
- Existing candidate eligibility, oldest-first ordering, filters, rules, locks, calculations, pagination, lifecycle actions, audit ordering, history newest-first ordering/100-row limit, selection, status, difference, and Open behavior remain unchanged.
- The workspace remains one HTTP request; do not add per-history-row client or API calls.
- No new table, persistent aggregate column, or data backfill.
- The new migration must be idempotent, detect unexpected installed-function drift, and live in the normal `supabase-migrations/` folder.
- SQL smoke behavior is authoritative. If PostgreSQL is unavailable, retain a minimal static migration safeguard and report the unexecuted SQL gap honestly.
- Automated layout coverage is limited to the structural DOM contract and real source-switch/render behavior. Wide/narrow field fitting, wrapping, and relative widths are browser-verification requirements; do not use CSS source-pattern assertions as their authority.

## File Structure

- Modify `index.html` — add explicit workbench source/filter row wrappers and replace the two history source headers with one Source header.
- Modify `styles.css` — implement wide-screen two-row control layout, compact field widths, responsive wrapping, and wrapping Source history cells.
- Modify `app-main.js` — normalize/format history source summaries and render the new single history cell.
- Create `supabase-migrations/2026-08-12-financial-reconciliation-history-source-summary.sql` — safely enrich workspace history rows with ordered per-source aggregates.
- Modify `tests/reconciliation-density.test.js` — executable client renderer/helper coverage plus scoped HTML/CSS layout contracts.
- Modify `tests/reconciliation.test.js` — minimal static migration installation and defensive-definition checks.
- Modify `tests/reconciliation-rpc.smoke.sql` — authoritative multi-source grouping, ordering, raw-total, omission, removal, and idempotency coverage.
- Do not modify `api/reconciliation.js` — it already passes the workspace RPC JSON through unchanged.

---

### Task 1: Two-row Workbench Control Layout

**Files:**
- Modify: `index.html:3843-3847`
- Modify: `styles.css:6249-6268, 6295-6318, 6320-6326`
- Test: `tests/reconciliation-density.test.js:129-137, 282-297`

**Interfaces:**
- Consumes: existing element IDs `financial-reconciliation-source`, `financial-reconciliation-rule-hint`, and `financial-reconciliation-dynamic-filters` used by `app-main.js`.
- Produces: `.financial-reconciliation-source-row` containing Source plus hint; `.financial-reconciliation-dynamic-filters` as the explicit second-row flex container; existing per-field classes `.financial-reconciliation-filter-<field>` remain the sizing hooks.

- [ ] **Step 1: Write the failing layout contract test**

Add this focused test to `tests/reconciliation-density.test.js` after the existing workbench source-selector test:

```js
test("workbench exposes source controls before the dynamic filter row", () => {
  assert.match(
    html,
    /class="financial-reconciliation-source-row"[\s\S]*id="financial-reconciliation-source"[\s\S]*id="financial-reconciliation-rule-hint"[\s\S]*<\/div>\s*<div id="financial-reconciliation-dynamic-filters"/,
  );
  assert.match(appMain, /class="financial-reconciliation-filter-\$\{escape\(field\)\}"/);
});
```

- [ ] **Step 2: Run the focused test to verify RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-density.test.js
```

Expected: FAIL in `workbench exposes source controls before the dynamic filter row` because the source-row wrapper does not exist.

- [ ] **Step 3: Add explicit source and filter rows**

Replace the filter container inside `index.html` with:

```html
<div id="financial-reconciliation-filters" class="grid filters financial-reconciliation-filters">
  <div class="financial-reconciliation-source-row">
    <label>Source<select id="financial-reconciliation-source"></select></label>
    <p id="financial-reconciliation-rule-hint" class="field-hint financial-reconciliation-rule-hint"></p>
  </div>
  <div id="financial-reconciliation-dynamic-filters" class="financial-reconciliation-dynamic-filters"></div>
</div>
```

Do not rename any IDs; the existing event binding and render functions must continue working.

- [ ] **Step 4: Implement compact wide-screen sizing and responsive wrapping**

Replace the current workbench filter-grid and `display: contents` rules with the following scoped CSS, retaining the existing label/input font-size rules:

```css
.financial-reconciliation-filters {
  display: grid;
  grid-template-columns: minmax(0, 1fr);
  gap: .65rem;
}

.financial-reconciliation-source-row {
  display: grid;
  grid-template-columns: minmax(12rem, 18rem) minmax(0, 1fr);
  gap: .75rem;
  align-items: end;
}

.financial-reconciliation-rule-hint {
  align-self: end;
  margin: 0 0 .6rem;
  font-size: .68rem;
}

.financial-reconciliation-dynamic-filters {
  display: flex;
  flex-wrap: nowrap;
  gap: .6rem;
  align-items: end;
  min-width: 0;
}

.financial-reconciliation-dynamic-filters > label { min-width: 0; }
.financial-reconciliation-filter-description { flex: 1 1 10rem; }
.financial-reconciliation-filter-dateFrom,
.financial-reconciliation-filter-dateTo { flex: 0 1 8.75rem; }
.financial-reconciliation-filter-amountMin,
.financial-reconciliation-filter-amountMax { flex: 0 1 7.5rem; }
.financial-reconciliation-filter-supplier,
.financial-reconciliation-filter-payment,
.financial-reconciliation-filter-account,
.financial-reconciliation-filter-category { flex: 0 1 8rem; }

@media (max-width: 1200px) {
  .financial-reconciliation-dynamic-filters { flex-wrap: wrap; }
  .financial-reconciliation-dynamic-filters > label { flex-grow: 1; }
}

@media (max-width: 768px) {
  .financial-reconciliation-source-row { grid-template-columns: 1fr; }
  .financial-reconciliation-rule-hint { margin: 0; }
}
```

Remove the obsolete `.financial-reconciliation-workbench-card .financial-reconciliation-rule-hint { grid-column: ... }` declarations and `.financial-reconciliation-dynamic-filters { display: contents; }`. Keep the existing 16px input/select rule inside the 768px media query.

- [ ] **Step 5: Run focused and full Node verification**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-density.test.js
node --test --test-isolation=none tests/*.test.js
git diff --check
```

Expected: all tests pass; no whitespace errors. Confirm the existing source-change tests still pass without changes.

Do not add CSS source-pattern assertions for field fitting. Task 4's browser checks are authoritative for the wide-screen single-row fit, compact relative widths, and responsive wrapping.

- [ ] **Step 6: Commit Task 1**

```powershell
git add -- index.html styles.css tests/reconciliation-density.test.js
git commit -m "style: simplify reconciliation filter layout"
```

---

### Task 2: Database History Source Aggregates

**Files:**
- Create: `supabase-migrations/2026-08-12-financial-reconciliation-history-source-summary.sql`
- Modify: `tests/reconciliation.test.js:14-34`
- Modify: `tests/reconciliation-rpc.smoke.sql:1-16, 16-170`

**Interfaces:**
- Consumes: installed RPC signature `public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)` after `2026-08-12-financial-reconciliation-oldest-first-candidates.sql`.
- Produces: every object in workspace `history` has `sourceSummary: Array<{sourceType: string, recordCount: number, amountTotal: number}>`, ordered base-first then by saved `matching_source_types` position.
- Preserves: `history` remains newest-first and limited to 100 active reconciliation rows.

- [ ] **Step 1: Write the failing static migration safeguard**

Add a safe read at the top of `tests/reconciliation.test.js`:

```js
const historySourceSummaryMigrationPath = path.join(
  __dirname,
  "..",
  "supabase-migrations",
  "2026-08-12-financial-reconciliation-history-source-summary.sql",
);
const historySourceSummaryMigration = fs.existsSync(historySourceSummaryMigrationPath)
  ? fs.readFileSync(historySourceSummaryMigrationPath, "utf8")
  : "";

test("history source-summary migration safely enriches the workspace function", () => {
  assert.match(historySourceSummaryMigration, /pg_get_functiondef\('public\.get_financial_reconciliation_workspace\(uuid,text,jsonb,integer,integer\)'::regprocedure\)/);
  assert.match(historySourceSummaryMigration, /'sourceSummary'/);
  assert.match(historySourceSummaryMigration, /count\(\*\)/i);
  assert.match(historySourceSummaryMigration, /sum\(i\.amount_snapshot\)/i);
  assert.match(historySourceSummaryMigration, /jsonb_array_elements_text\(h\.matching_source_types\)\s+with ordinality/i);
  assert.match(historySourceSummaryMigration, /old_history_count = 1\s+and new_history_count = 0/is);
  assert.match(historySourceSummaryMigration, /old_history_count = 0\s+and new_history_count = 1/is);
  assert.match(historySourceSummaryMigration, /unexpected reconciliation workspace function definition/i);
});
```

- [ ] **Step 2: Run the focused Node test to verify RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation.test.js
```

Expected: FAIL in the new migration safeguard because the migration file is absent. The safe read must prevent an `ENOENT` failure so RED points at the missing contract.

- [ ] **Step 3: Create the idempotent forward migration**

Create `supabase-migrations/2026-08-12-financial-reconciliation-history-source-summary.sql` with this structure and exact aggregate semantics:

```sql
-- Adds raw per-source record counts and totals to reconciliation history rows.
do $fix$
declare
  definition text;
  original_definition text;
  old_history constant text := $$'history',coalesce((select jsonb_agg(to_jsonb(h) order by h.created_at desc) from (select * from public.financial_reconciliations where deleted_at is null order by created_at desc limit 100) h),'[]'::jsonb)$$;
  new_history constant text := $$'history',coalesce((
    select jsonb_agg(
      to_jsonb(h) || jsonb_build_object('sourceSummary',coalesce(summary.source_summary,'[]'::jsonb))
      order by h.created_at desc
    )
    from (
      select *
      from public.financial_reconciliations
      where deleted_at is null
      order by created_at desc
      limit 100
    ) h
    left join lateral (
      select jsonb_agg(
        jsonb_build_object(
          'sourceType',grouped.source_type,
          'recordCount',grouped.record_count,
          'amountTotal',grouped.amount_total
        )
        order by
          case when grouped.source_type=h.base_source_type then 0 else 1 end,
          coalesce((
            select matching.position
            from jsonb_array_elements_text(h.matching_source_types)
              with ordinality matching(source_type,position)
            where matching.source_type=grouped.source_type
            limit 1
          ),2147483647),
          grouped.source_type
      ) as source_summary
      from (
        select
          i.source_type,
          count(*) as record_count,
          coalesce(sum(i.amount_snapshot),0) as amount_total
        from public.financial_reconciliation_items i
        where i.reconciliation_id=h.id
        group by i.source_type
      ) grouped
    ) summary on true
  ),'[]'::jsonb)$$;
  old_history_count integer;
  new_history_count integer;
begin
  select pg_get_functiondef(
    'public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)'::regprocedure
  ) into definition;
  original_definition := definition;

  old_history_count := (length(definition)-length(replace(definition,old_history,''))) / length(old_history);
  new_history_count := (length(definition)-length(replace(definition,new_history,''))) / length(new_history);

  if not (
    (old_history_count = 1 and new_history_count = 0)
    or (old_history_count = 0 and new_history_count = 1)
  ) then
    raise exception 'Unexpected reconciliation workspace function definition; could not install history source summaries.';
  end if;

  if old_history_count = 1 then
    definition := replace(definition,old_history,new_history);
  end if;

  old_history_count := (length(definition)-length(replace(definition,old_history,''))) / length(old_history);
  new_history_count := (length(definition)-length(replace(definition,new_history,''))) / length(new_history);
  if old_history_count <> 0 or new_history_count <> 1 then
    raise exception 'Unexpected reconciliation workspace function definition; could not verify history source summaries.';
  end if;

  if definition <> original_definition then
    execute definition;
  end if;
end $fix$;
```

Keep the new history fragment semantically identical to this code. If `pg_get_functiondef` normalizes whitespace differently in the target definition, adjust both exact constants using an observed disposable-database definition; do not weaken the count/drift validation to an unrestricted replacement.

- [ ] **Step 4: Extend authoritative SQL smoke coverage**

Immediately after the oldest-first migration includes at the top of `tests/reconciliation-rpc.smoke.sql`, apply the new migration twice:

```sql
\ir ../supabase-migrations/2026-08-12-financial-reconciliation-history-source-summary.sql
\ir ../supabase-migrations/2026-08-12-financial-reconciliation-history-source-summary.sql
```

Extend the `do` block declaration with:

```sql
history_rid uuid := gen_random_uuid();
history_row jsonb;
history_source_ids text[];
history_card_item_id uuid := gen_random_uuid();
```

Before the existing action-flow fixtures, insert a reconciliation and raw snapshot items that cannot be confused with source-table values:

```sql
insert into financial_reconciliations(
  id,status,base_source_type,matching_source_types,matching_source_rules,created_by
) values (
  history_rid,'started','financial_documents',
  '["import_cgd_extrato_ordem","import_cgd_cartao_credito"]'::jsonb,
  '[{"sourceType":"import_cgd_extrato_ordem","operator":"+"},{"sourceType":"import_cgd_cartao_credito","operator":"-"}]'::jsonb,
  'smoke-history-summary'
);

insert into financial_reconciliation_items(
  reconciliation_id,source_type,source_id,amount_snapshot,created_by
) values
  (history_rid,'financial_documents',gen_random_uuid(),200,'smoke-history-summary'),
  (history_rid,'financial_documents',gen_random_uuid(),250,'smoke-history-summary'),
  (history_rid,'import_cgd_extrato_ordem',gen_random_uuid(),-200,'smoke-history-summary'),
  (history_rid,'import_cgd_extrato_ordem',gen_random_uuid(),-250,'smoke-history-summary'),
  (history_rid,'import_cgd_cartao_credito',history_card_item_id,25,'smoke-history-summary');

r := get_financial_reconciliation_workspace(
  history_rid,'financial_documents','{}'::jsonb,1,50
);
select history_record
into history_row
from jsonb_array_elements(r->'history') history_records(history_record)
where history_record->>'id'=history_rid::text;

select array_agg(source_entry->>'sourceType' order by position)
into history_source_ids
from jsonb_array_elements(history_row->'sourceSummary')
  with ordinality source_entries(source_entry,position);

if history_source_ids is distinct from array[
  'financial_documents','import_cgd_extrato_ordem','import_cgd_cartao_credito'
] then
  raise exception 'History sources were not ordered base-first then by saved matching source order: %', history_source_ids;
end if;

if history_row->'sourceSummary'->0 is distinct from
   '{"sourceType":"financial_documents","recordCount":2,"amountTotal":450}'::jsonb
   or history_row->'sourceSummary'->1 is distinct from
   '{"sourceType":"import_cgd_extrato_ordem","recordCount":2,"amountTotal":-450}'::jsonb
   or history_row->'sourceSummary'->2 is distinct from
   '{"sourceType":"import_cgd_cartao_credito","recordCount":1,"amountTotal":25}'::jsonb then
  raise exception 'History source summary did not preserve raw counts and amount snapshots: %', history_row->'sourceSummary';
end if;

delete from financial_reconciliation_items
where reconciliation_id=history_rid
  and source_type='import_cgd_cartao_credito'
  and source_id=history_card_item_id;

r := get_financial_reconciliation_workspace(
  history_rid,'financial_documents','{}'::jsonb,1,50
);
select history_record
into history_row
from jsonb_array_elements(r->'history') history_records(history_record)
where history_record->>'id'=history_rid::text;

if jsonb_array_length(history_row->'sourceSummary') <> 2
   or exists (
     select 1
     from jsonb_array_elements(history_row->'sourceSummary') entry
     where entry->>'sourceType'='import_cgd_cartao_credito'
   ) then
  raise exception 'Removed or unused history source was not omitted: %', history_row->'sourceSummary';
end if;
```

This fixture proves raw totals despite the card rule being `-`, proves negative values remain negative, proves base-first and matching order, and proves removal/omission.

- [ ] **Step 5: Run Node GREEN and SQL smoke when available**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation.test.js
node --check app-main.js
node --test --test-isolation=none tests/*.test.js
git diff --check
```

Expected: all Node tests pass and diff check is clean.

If `psql` and a confirmed disposable Supabase/PostgreSQL connection are available, run from `tests/`:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f reconciliation-rpc.smoke.sql
```

Expected: exit 0 and transaction rolls back. If either requirement is unavailable, record `SQL smoke not executed` in the task report; do not treat the static Node safeguard as database execution.

- [ ] **Step 6: Commit Task 2**

```powershell
git add -- supabase-migrations/2026-08-12-financial-reconciliation-history-source-summary.sql tests/reconciliation.test.js tests/reconciliation-rpc.smoke.sql
git commit -m "feat: summarize reconciliation history sources"
```

---

### Task 3: Single History Source Column and Client Formatting

**Files:**
- Modify: `index.html:3875-3878`
- Modify: `app-main.js:21334-21348, 21589-21593`
- Modify: `styles.css:6455-6458`
- Test: `tests/reconciliation-density.test.js:339-375`

**Interfaces:**
- Consumes: Task 2 history field `sourceSummary` with ordered `{sourceType, recordCount, amountTotal}` objects.
- Produces:
  - `financialReconciliationHistorySourceSummary(record): Array<{sourceType:string,recordCount:number,amountTotal:number}> | null`
  - `financialReconciliationHistorySourceText(record): string`
  - one history Source cell with class `.financial-reconciliation-history-source`.
- `null` means missing/malformed summary container; `[]` means a valid present empty summary.

- [ ] **Step 1: Write failing executable helper and renderer tests**

Replace `Reconciliation history still renders base and matching source labels` in `tests/reconciliation-density.test.js` with tests that extract and execute the real new functions:

```js
const historySourceHelpers = new Function(
  "FINANCIAL_RECONCILIATION_SOURCES",
  "financialReconciliationSourceLabel",
  "formatMoney",
  "clean",
  `${appFunctionSource("financialReconciliationHistorySourceSummary")}
   ${appFunctionSource("financialReconciliationHistorySourceText")}
   return { financialReconciliationHistorySourceSummary, financialReconciliationHistorySourceText };`,
)(
  {
    financial_documents: "Financial Documents",
    import_cgd_extrato_ordem: "CGD Bank Statement",
  },
  (value) => ({
    financial_documents: "Financial Documents",
    import_cgd_extrato_ordem: "CGD Bank Statement",
  })[value] || value,
  (value) => `${Number(value).toFixed(2)} €`,
  (value) => String(value || "").trim(),
);

test("history source text uses raw ordered source aggregates", () => {
  const record = {
    sourceSummary: [
      { sourceType: "financial_documents", recordCount: 4, amountTotal: 450 },
      { sourceType: "import_cgd_extrato_ordem", recordCount: 4, amountTotal: -450 },
    ],
  };
  assert.equal(
    historySourceHelpers.financialReconciliationHistorySourceText(record),
    "Financial Documents (#4; 450.00 €), CGD Bank Statement (#4; -450.00 €)",
  );
});

test("history source text distinguishes missing malformed and empty summaries", () => {
  assert.equal(historySourceHelpers.financialReconciliationHistorySourceText({}), "Source details unavailable");
  assert.equal(historySourceHelpers.financialReconciliationHistorySourceText({ sourceSummary: "invalid" }), "Source details unavailable");
  assert.equal(historySourceHelpers.financialReconciliationHistorySourceText({ sourceSummary: [] }), "No records");
  assert.deepEqual(
    historySourceHelpers.financialReconciliationHistorySourceSummary({
      sourceSummary: [
        { sourceType: "financial_documents", recordCount: 2, amountTotal: 20 },
        { sourceType: "unknown", recordCount: 1, amountTotal: 10 },
        { sourceType: "financial_documents", recordCount: 2, amountTotal: 20 },
        { sourceType: "import_cgd_extrato_ordem", recordCount: 0, amountTotal: -20 },
      ],
    }),
    [{ sourceType: "financial_documents", recordCount: 2, amountTotal: 20 }],
  );
});
```

Create an actual renderer harness using `appFunctionSource("renderFinancialReconciliationHistory")`:

```js
function renderHistory(record, selectedReconciliationId = "") {
  const els = { financialReconciliationHistoryRows: { innerHTML: "" } };
  const current = { selectedReconciliationId, workspace: { history: [record] } };
  const render = new Function(
    "financialReconciliationState",
    "clean",
    "escape",
    "formatDateTimeShort",
    "financialReconciliationHistorySourceText",
    "financialReconciliationStatusMarkup",
    "financialReconciliationDifference",
    "formatMoney",
    "els",
    `${appFunctionSource("renderFinancialReconciliationHistory")}
     return renderFinancialReconciliationHistory;`,
  )(
    () => current,
    (value) => String(value || "").trim(),
    (value) => String(value),
    () => "2026-08-12 10:00",
    historySourceHelpers.financialReconciliationHistorySourceText,
    (status) => status === "complete" ? "Complete" : "Started",
    (value) => Number(value.difference_amount),
    (value) => `${Number(value).toFixed(2)} €`,
    els,
  );
  render();
  return els.financialReconciliationHistoryRows.innerHTML;
}

test("history renders one wrapping source summary and preserves row behavior", () => {
  const markup = renderHistory({
    id: "rec-1",
    created_at: "2026-08-12T10:00:00Z",
    status: "complete",
    difference_amount: 0,
    sourceSummary: [
      { sourceType: "financial_documents", recordCount: 4, amountTotal: 450 },
      { sourceType: "import_cgd_extrato_ordem", recordCount: 4, amountTotal: -450 },
    ],
  }, "rec-1");

  assert.match(html, /<th>Created<\/th><th>Source<\/th><th>Status<\/th><th>Difference<\/th><th><\/th>/);
  assert.doesNotMatch(html, /<th>Base source<\/th>|<th>Matching sources<\/th>/);
  assert.match(markup, /class="selected"/);
  assert.match(markup, /class="financial-reconciliation-history-source"/);
  assert.match(markup, /Financial Documents \(#4; 450\.00 €\), CGD Bank Statement \(#4; -450\.00 €\)/);
  assert.match(markup, /Complete/);
  assert.match(markup, /0\.00 €/);
  assert.match(markup, /data-financial-reconciliation-select="rec-1">Open<\/button>/);
});
```

- [ ] **Step 2: Run the focused tests to verify RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-density.test.js
```

Expected: FAIL because the helper functions and single Source header do not yet exist.

- [ ] **Step 3: Implement summary validation and formatting helpers**

Add these functions immediately before `renderFinancialReconciliationHistory` in `app-main.js`:

```js
function financialReconciliationHistorySourceSummary(record) {
  if (!record || !Object.prototype.hasOwnProperty.call(record, "sourceSummary")) return null;
  if (!Array.isArray(record.sourceSummary)) return null;
  const seen = new Set();
  return record.sourceSummary.flatMap((entry) => {
    const sourceType = clean(entry?.sourceType);
    const recordCount = Number(entry?.recordCount);
    const amountTotal = Number(entry?.amountTotal);
    if (!Object.prototype.hasOwnProperty.call(FINANCIAL_RECONCILIATION_SOURCES, sourceType)
      || seen.has(sourceType)
      || !Number.isInteger(recordCount)
      || recordCount <= 0
      || !Number.isFinite(amountTotal)) return [];
    seen.add(sourceType);
    return [{ sourceType, recordCount, amountTotal }];
  });
}

function financialReconciliationHistorySourceText(record) {
  const summary = financialReconciliationHistorySourceSummary(record);
  if (summary === null) return "Source details unavailable";
  if (!summary.length) return "No records";
  return summary.map((entry) => (
    `${financialReconciliationSourceLabel(entry.sourceType)} (#${entry.recordCount}; ${formatMoney(entry.amountTotal)})`
  )).join(", ");
}
```

Do not sort in these helpers. Task 2's database order is authoritative and must be preserved.

- [ ] **Step 4: Replace the history headers and row cells**

Change the history table header in `index.html` to:

```html
<thead><tr><th>Created</th><th>Source</th><th>Status</th><th>Difference</th><th></th></tr></thead>
```

Update `renderFinancialReconciliationHistory` so each row renders:

```js
`<tr class="${clean(record.id) === financialReconciliationState().selectedReconciliationId ? "selected" : ""}">
  <td>${escape(formatDateTimeShort(record.created_at) || "-")}</td>
  <td class="financial-reconciliation-history-source">${escape(financialReconciliationHistorySourceText(record))}</td>
  <td>${financialReconciliationStatusMarkup(record.status)}</td>
  <td>${escape(formatMoney(financialReconciliationDifference(record)))}</td>
  <td><button type="button" class="ghost" data-financial-reconciliation-select="${escape(record.id)}">Open</button></td>
</tr>`
```

Change the empty-history row from `colspan="6"` to `colspan="5"`. Preserve selected-row class logic and all existing formatting helpers.

- [ ] **Step 5: Add wrapping Source-cell CSS**

Add near the existing history selected-row rule:

```css
.financial-reconciliation-history-source {
  min-width: 18rem;
  white-space: normal;
  overflow-wrap: anywhere;
}
```

- [ ] **Step 6: Run focused and full verification**

Run:

```powershell
node --check app-main.js
node --test --test-isolation=none tests/reconciliation-density.test.js
node --test --test-isolation=none tests/*.test.js
git diff --check
```

Expected: syntax check and all tests pass, with no whitespace errors. Confirm the actual-renderer test covers selected styling, Status, Difference, and Open instead of only testing helper output.

- [ ] **Step 7: Commit Task 3**

```powershell
git add -- index.html app-main.js styles.css tests/reconciliation-density.test.js
git commit -m "feat: show reconciliation history source totals"
```

---

### Task 4: Integrated Verification and Rollout Handoff

**Files:**
- Verify only; no planned source edits.
- Reference: `docs/superpowers/specs/2026-08-12-reconciliation-workbench-history-source-summary-design.md`
- Reference: `supabase-migrations/2026-08-12-financial-reconciliation-history-source-summary.sql`

**Interfaces:**
- Consumes: Tasks 1-3 completed commits.
- Produces: evidence that automated checks pass and an exact migration/browser handoff identifying any environment-dependent gaps.

- [ ] **Step 1: Run fresh automated verification on the final branch**

Run:

```powershell
node --check app-main.js
node --test --test-isolation=none tests/*.test.js
git diff --check
git status --short --branch
git log --oneline main..HEAD
```

Expected: JavaScript syntax passes, all Node tests pass with zero failures, diff check is clean, only planned files differ from the branch base, and the three task commits are present.

- [ ] **Step 2: Execute SQL behavior verification when safely configured**

Check availability without printing credential values:

```powershell
Get-Command psql -ErrorAction SilentlyContinue
Get-ChildItem Env: | Where-Object { $_.Name -eq 'SUPABASE_DB_URL' } | Select-Object -ExpandProperty Name
```

When both `psql` and a confirmed disposable `SUPABASE_DB_URL` are present, run:

```powershell
Set-Location tests
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f reconciliation-rpc.smoke.sql
Set-Location ..
```

Expected: exit 0; the smoke transaction rolls back. Otherwise report exactly which prerequisite was absent and mark SQL behavior verification pending.

- [ ] **Step 3: Document exact migration order**

For a target that already has all previously published reconciliation migrations through oldest-first candidates, apply only:

1. `supabase-migrations/2026-08-12-financial-reconciliation-history-source-summary.sql`

For a target missing the recent prerequisites, preserve this order:

1. `supabase-migrations/2026-08-11-financial-reconciliation-source-rules.sql`
2. `supabase-migrations/2026-08-11-financial-reconciliation-source-rules-workspace-filter-fix.sql`
3. `supabase-migrations/2026-08-12-financial-reconciliation-action-overload-fix.sql`
4. `supabase-migrations/2026-08-12-financial-reconciliation-oldest-first-candidates.sql`
5. `supabase-migrations/2026-08-12-financial-reconciliation-history-source-summary.sql`

State explicitly: publishing application code does not apply Supabase SQL; until migration item 5 is applied, history displays `Source details unavailable`; no backfill is required.

- [ ] **Step 4: Perform authenticated browser verification when available**

In a non-production environment with the new migration applied and a signed-in account authorized for Reconciliation, verify:

1. At a wide viewport, Source plus the rule hint occupy the first workbench row.
2. All Financial Documents filters occupy the second row without overlap.
3. Date, amount, Supplier, Payment, Account, and Category are visibly narrower; Description uses remaining width.
4. At 1200px and below, filters wrap cleanly; at 768px and below, inputs remain 16px.
5. Switching Source refreshes the applicable filter set and candidate records.
6. A history reconciliation with Financial Documents totals `450` across four items and bank totals `-450` across four items displays `Financial Documents (#4; 450,00 €), CGD Bank Statement (#4; -450,00 €)` using the application's locale formatting.
7. A reconciliation whose saved matching rule uses `-` still displays the raw source total, not an operator-adjusted total.
8. Sources with no remaining records are absent; an explicit empty summary says `No records`.
9. Created, Status, Difference, selected highlighting, and Open continue working.

If authentication or a migrated non-production target is unavailable, list these checks as pending rather than claiming visual verification.

- [ ] **Step 5: Review final scope and prepare integration choice**

Run:

```powershell
git diff --name-status main...HEAD
git diff --stat main...HEAD
```

Expected implementation scope:

```text
M app-main.js
M index.html
M styles.css
A supabase-migrations/2026-08-12-financial-reconciliation-history-source-summary.sql
M tests/reconciliation-density.test.js
M tests/reconciliation-rpc.smoke.sql
M tests/reconciliation.test.js
```

After independent final review and fresh verification, use `superpowers:finishing-a-development-branch` and offer the standard merge/push/keep choices. Do not merge, publish, or apply SQL without the user's chosen next action.
