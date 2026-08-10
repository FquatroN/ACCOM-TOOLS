# Reconciliation Current Details Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Enrich locked reconciliation items with live source details and show a compact details line beneath each record in the Current reconciliation panel.

**Architecture:** The workspace RPC will enrich each stored reconciliation item through a left lateral call to `financial_reconciliation_source`, so existing reconciliations receive source date, description, and supplier without schema changes. The client will compose the requested date/supplier/description detail line and use tightly scoped CSS for the right-panel body content.

**Tech Stack:** PostgreSQL/Supabase SQL, vanilla JavaScript, CSS, Node.js built-in test runner.

## Global Constraints

- Resolve `source_date`, `description`, and `supplier` at workspace-load time through `financial_reconciliation_source(source_type, source_id)`.
- A missing source record must not hide or fail the locked item; its detail values are empty.
- Do not change reconciliation-item persistence, lock rules, amounts, calculations, audit data, filters, or actions.
- Detail order is `date · supplier · description` when supplier exists; otherwise `date · description`. Omit every unavailable part without a placeholder.
- Scope smaller typography to Current reconciliation body content; retain the card title and action buttons at their current size.
- Add a normal forward migration in `supabase-migrations/` and update the original reconciliation migration so fresh and existing databases agree.

---

### Task 1: Enrich workspace items through the source function

**Files:**
- Modify: `supabase-migrations/2026-08-09-financial-reconciliation.sql:151`
- Create: `supabase-migrations/2026-08-10-financial-reconciliation-current-details.sql`
- Modify: `tests/reconciliation-rpc.smoke.sql:18-19`

**Interfaces:**
- Consumes: `financial_reconciliation_items i` and `financial_reconciliation_source(i.source_type, i.source_id) s`.
- Produces: every workspace `items[]` entry includes `source_date`, `description`, and `supplier`, each nullable.

- [ ] **Step 1: Add a failing smoke assertion for enriched locked items**

In `tests/reconciliation-rpc.smoke.sql`, after the workspace call, add this assertion using the already-created `doc_id` record:

```sql
if not exists (
  select 1
  from jsonb_array_elements(r->'items') item
  where (item->>'source_id')::uuid = doc_id
    and item->>'source_date' = '2026-01-02'
    and item->>'description' = 'smoke document'
    and item->>'supplier' = 'Smoke Supplier'
) then
  raise exception 'Workspace item details were not returned';
end if;
```

Set the smoke document insert values to `description = 'smoke document'` and `supplier_name = 'Smoke Supplier'` if they are not already set.

- [ ] **Step 2: Verify the smoke assertion fails against the current RPC**

Run the SQL smoke script in Supabase after applying the current reconciliation migration.

Expected: FAIL with `Workspace item details were not returned`, because `items` is currently built from `to_jsonb(i)` alone.

- [ ] **Step 3: Update the source migration and add the forward migration**

In the `items` value of the `jsonb_build_object` return expression, replace the current aggregation:

```sql
coalesce((select jsonb_agg(to_jsonb(i) order by i.created_at)
  from financial_reconciliation_items i
  where i.reconciliation_id=p_reconciliation_id),'[]'::jsonb)
```

with this aggregation:

```sql
coalesce((select jsonb_agg(
  to_jsonb(i) || jsonb_build_object(
    'source_date', s.source_date,
    'description', s.description,
    'supplier', nullif(s.supplier, '')
  ) order by i.created_at
) from financial_reconciliation_items i
left join lateral financial_reconciliation_source(i.source_type, i.source_id) s on true
where i.reconciliation_id=p_reconciliation_id),'[]'::jsonb)
```

Create `supabase-migrations/2026-08-10-financial-reconciliation-current-details.sql` with a `create or replace function public.get_financial_reconciliation_workspace(...)` definition copied from the updated source migration. Preserve all existing parameters, candidate queries, filter predicates, grants, and source configuration; only the `items` aggregation changes as shown above.

- [ ] **Step 4: Run the updated SQL smoke script**

Run: `tests/reconciliation-rpc.smoke.sql` after the original migration and the new forward migration.

Expected: the smoke script completes without exception and the document locked item includes its date, description, and supplier.

- [ ] **Step 5: Commit the RPC deliverable**

```bash
git add supabase-migrations/2026-08-09-financial-reconciliation.sql supabase-migrations/2026-08-10-financial-reconciliation-current-details.sql tests/reconciliation-rpc.smoke.sql
git commit -m "feat: include reconciliation item details"
```

### Task 2: Render compact locked-record details in the panel

**Files:**
- Modify: `app-main.js:21369-21374`
- Modify: `styles.css:6286-6350`
- Modify: `tests/reconciliation-density.test.js:1-45`

**Interfaces:**
- Consumes: enriched item properties `source_date`, `supplier`, and `description` from Task 1.
- Produces: `financialReconciliationItemDetails(item)`, `.financial-reconciliation-item-details`, and compact Current reconciliation body selectors.

- [ ] **Step 1: Write failing client source-contract assertions**

Add these assertions to `tests/reconciliation-density.test.js` after reading `app-main.js`:

```js
assert.match(appMain, /function financialReconciliationItemDetails\(item\)/);
assert.match(appMain, /\[formatDateOnly\(item\.source_date\), clean\(item\.supplier\), clean\(item\.description\)\]\.filter\(Boolean\)\.join\(" · "\)/);
assert.match(appMain, /class="financial-reconciliation-item-details"/);
assert.match(css, /\.financial-reconciliation-item-details\s*\{\s*font-size:\s*\.68rem;/);
assert.match(css, /\.financial-reconciliation-current\s*\{\s*font-size:\s*\.86rem;/);
assert.match(css, /\.financial-reconciliation-current h3\s*\{[\s\S]*font-size:\s*\.82rem;/);
```

- [ ] **Step 2: Run the focused test to verify it fails**

Run: `node --test tests/reconciliation-density.test.js`

Expected: FAIL because the item-details formatter and current-panel rules do not exist.

- [ ] **Step 3: Add detail composition and markup**

Add this helper immediately before `renderFinancialReconciliationCurrent`:

```js
function financialReconciliationItemDetails(item) {
  return [formatDateOnly(item.source_date), clean(item.supplier), clean(item.description)].filter(Boolean).join(" · ");
}
```

Replace the locked-items map with markup that keeps the primary grid and adds a full-width detail line only when `financialReconciliationItemDetails(item)` is non-empty:

```js
const items = workspace.items.map((item) => {
  const details = financialReconciliationItemDetails(item);
  return `<li><span>${escape(financialReconciliationSourceLabel(item.source_type))}</span><strong>${escape(formatMoney(Number(item.amount_snapshot || 0)))}</strong>${complete ? "" : `<button type="button" class="ghost" data-financial-reconciliation-remove data-source-type="${escape(item.source_type)}" data-source-id="${escape(item.source_id)}">Remove</button>`}${details ? `<small class="financial-reconciliation-item-details">${escape(details)}</small>` : ""}</li>`;
}).join("") || "<li>No locked records.</li>";
```

- [ ] **Step 4: Scope compact Current reconciliation typography**

Add these CSS rules in the reconciliation block:

```css
.financial-reconciliation-current { font-size: .86rem; }

.financial-reconciliation-current h3 {
  margin: .8rem 0 .35rem;
  font-size: .82rem;
}

.financial-reconciliation-items li { font-size: .78rem; }

.financial-reconciliation-item-details {
  grid-column: 1 / -1;
  color: var(--muted);
  font-size: .68rem;
  line-height: 1.3;
  overflow-wrap: anywhere;
}

.financial-reconciliation-audit li { font-size: .74rem; }
```

Do not add a `font-size` selector for `.financial-reconciliation-basket h2`, `.financial-reconciliation-actions`, `.financial-reconciliation-items button`, or the complete/force-complete button.

- [ ] **Step 5: Run the Node suite and formatting check**

Run: `node --test tests/*.test.js`

Expected: all tests PASS.

Run: `git diff --check`

Expected: no whitespace errors.

- [ ] **Step 6: Commit the panel deliverable**

```bash
git add app-main.js styles.css tests/reconciliation-density.test.js
git commit -m "style: add reconciliation current details"
```
