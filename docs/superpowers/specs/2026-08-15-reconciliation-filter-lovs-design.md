# Manual Reconciliation Filter LOVs Design

## Objective

Make the Manual Reconciliation filters source-specific and easier to use by replacing selected free-text fields with list-of-values (LOV) dropdowns, expanding Financial Documents supplier search to cover both name and NIF, and removing the Financial Documents Account filter.

## Approved user experience

### Financial Documents

Show these filters in this order:

1. Date from
2. Date to
3. Amount from
4. Amount to
5. Description
6. Supplier Search
7. Payment
8. Category

Payment and Category are single-select LOV dropdowns. Each starts with an `All payments` or `All categories` option and then lists every distinct, trimmed, nonblank value currently present in `financial_documents`, not merely values from the visible page.

Supplier Search remains one text field. It performs a case-insensitive partial match against either `financial_documents.supplier_name` or `financial_documents.supplier_nif`.

The Account filter is not shown and is not sent for Financial Documents.

### FDM Accounts

Show these filters in this order:

1. Date from
2. Date to
3. Amount from
4. Amount to
5. Description
6. Account
7. Category

Account and Category are single-select LOV dropdowns. Each starts with an `All accounts` or `All categories` option and then lists every distinct, trimmed, nonblank value currently present in `import_fdm_accounts`, not merely values from the visible page.

### Other reconciliation sources

CGD Credit Card and CGD Bank Statement filters remain unchanged.

## Workspace contract

The existing `get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)` RPC remains the only Manual Reconciliation data request and retains its signature.

Its `sourceConfig` object gains a `filterOptions` object:

```json
{
  "sourceType": "financial_documents",
  "filterFields": ["dateFrom", "dateTo", "amountMin", "amountMax", "description", "supplier", "payment", "category"],
  "filterOptions": {
    "payment": ["Banco", "Caixa", "Visa"],
    "category": ["Energy", "Food"]
  }
}
```

For FDM Accounts, `filterOptions` contains `account` and `category`. For sources without LOV filters it is an empty object. Values are strings, trimmed first, deduplicated by the resulting trimmed value, and sorted case-insensitively with the original value as a deterministic tie-breaker.

LOV options are derived from the complete relevant source table and are independent of the current page, current filters, lock state, and reconciliation status. Eligibility and locking continue to determine candidate records only.

## Filtering behavior

The existing exact-match behavior is retained for Payment, Category, and Account selections. An empty LOV selection means no restriction.

Financial Documents supplier filtering changes from name-only matching to:

```sql
s.supplier ilike '%' || supplier_search || '%'
or s.supplier_nif ilike '%' || supplier_search || '%'
```

The existing minimum reconciliation date, `fat = 'S'` eligibility, locking rules, paging, oldest-first ordering, status counts, source rules, and reconciliation lifecycle behavior are unchanged.

Only filter keys advertised by the selected source's `filterFields` are sent by the client. This prevents the removed Financial Documents Account value, or any stale source-specific value, from reaching the RPC after a source change.

## Frontend rendering

The dynamic filter renderer continues to use `sourceConfig.filterFields` for ordering. When `sourceConfig.filterOptions[field]` is an array, it renders an escaped `<select>` with an `All ...` option and the returned values. Otherwise it renders the existing date, number, or search input.

The `supplier` field label becomes `Supplier Search`. Its control remains a search input.

Changing any filter keeps the existing behavior: update the stored filter value, reset pagination to page 1, and reload the workspace. Switching source also retains the existing pagination reset and source reload behavior.

If an older workspace response omits `filterOptions`, the client treats it as an empty object. If an LOV has no values, its dropdown still renders the corresponding `All ...` option.

## Error handling and security

- The existing `financial-reconciliation` application authorization remains mandatory.
- No new endpoint, direct table permission, or browser-side database access is added.
- LOV values are escaped before insertion into HTML.
- The migration preserves the RPC's security-definer search path and existing execute grants.
- Failure to load the workspace uses the existing error presentation and preserves any open reconciliation.

## Migration strategy

Add one migration in `supabase-migrations` that safely replaces the current workspace RPC without changing its signature. The migration must be compatible with the latest workspace enrichments, including oldest-first candidates, history source summaries, current item details, reconciliation origins, and automatic-reconciliation provenance.

The migration updates the source-specific `filterFields`, adds `filterOptions`, and changes the supplier predicate. It does not recreate tables or alter stored reconciliation data.

## Verification

Automated tests must cover:

- Financial Documents advertises Payment and Category LOVs, omits Account, and labels Supplier Search correctly.
- FDM Accounts advertises Account and Category LOVs.
- LOV values come from complete source tables, exclude blank/null values, are distinct, and have deterministic case-insensitive ordering.
- Supplier Search matches either supplier name or supplier NIF and rejects unrelated records.
- LOV controls render as single-select dropdowns with `All` options and escaped values.
- Sources without LOVs retain their existing text/date/number filters.
- Source changes and filter changes continue to reset pagination and reload candidates.
- Existing authorization, eligibility, locking, counts, oldest-first ordering, history, lifecycle, and automatic-reconciliation tests remain green.

The SQL smoke test should apply the migration twice to prove reapplication safety and exercise the workspace response against controlled Financial Documents and FDM Accounts fixtures.

## Non-goals

- Multi-select LOVs.
- Type-ahead/autocomplete controls.
- Dependent LOVs that change based on other active filters.
- Changes to eligible-record columns or reconciliation matching rules.
- Changes to Settings > Reconciliation or Automatic Reconciliation.
