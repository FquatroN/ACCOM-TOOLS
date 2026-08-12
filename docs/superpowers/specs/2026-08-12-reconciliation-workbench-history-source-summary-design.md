# Reconciliation Workbench and History Source Summary Design

**Date:** 2026-08-12  
**Status:** Approved

## Objective

Make the Reconciliation workbench filters easier to scan and replace the history table's configuration-oriented source columns with one record-oriented Source summary.

## Scope

This change covers:

- the two-row layout of the Reconciliation workbench controls;
- narrower widths for compact filter fields;
- one Source column in Reconciliation history;
- per-source record counts and raw amount totals in each history row;
- the workspace database response needed to render those summaries;
- regression coverage for layout, aggregation, formatting, and preserved behavior.

It does not change reconciliation calculations, matching rules, eligible-record filtering, candidate ordering, item locks, lifecycle actions, audit history, or the current reconciliation card.

## Workbench Layout

The workbench control card has two logical rows.

### First row

- The Source selector appears first.
- The existing reconciliation-rule hint appears after the selector.
- The hint continues to show the selected source's configured matches and operators.

### Second row

- Every filter applicable to the selected source appears in this row on wide screens.
- Description receives the flexible remaining width.
- Date from, Date to, Amount from, Amount to, Supplier, Payment, Account, and Category use narrower fixed or bounded widths so all Financial Documents filters fit on a normal wide desktop display.
- Source-specific filter visibility remains driven by the workspace `filterFields` response.
- Changing Source continues to reset candidate pagination and refresh eligible records.

At narrower breakpoints the second row may wrap. Mobile controls retain their existing accessible font sizing and usable touch targets.

## Reconciliation History Table

The history columns become:

1. Created
2. Source
3. Status
4. Difference
5. Open action

The former Base source and Matching sources columns are removed only from the Reconciliation history table. The reconciliation model may continue storing those fields because they remain necessary for rules, calculations, and deterministic ordering.

### Source summary format

Each used source is rendered as:

`<Source label> (#<record count>; <raw amount total>)`

Multiple sources are separated by a comma and a space. Example:

`Financial Documents (#4; 450,00 €), CGD Bank Statement (#4; -450,00 €)`

The cell may wrap naturally so long summaries remain readable without widening the table beyond its container.

### Aggregation rules

- Aggregate the current `financial_reconciliation_items` belonging to each reconciliation.
- Group by `source_type`.
- `recordCount` is the number of item rows in that source group.
- `amountTotal` is the raw sum of `amount_snapshot` for the group.
- Do not apply the reconciliation rule's `+` or `-` operator to `amountTotal`.
- Negative stored amounts remain negative.
- Show the reconciliation's `base_source_type` first when it has records.
- Order the remaining used sources according to the saved `matching_source_types` order.
- Omit configured sources that have no records.
- Return `No records` when the updated database explicitly returns an empty summary.
- Return `Source details unavailable` when the history row does not contain the new summary field, which safely distinguishes a missing migration from a truly empty reconciliation.

Existing Created formatting, newest-first history order, 100-row history limit, selected-row highlighting, Status badge, Difference calculation/display, and Open behavior remain unchanged.

## Database and API Contract

The existing reconciliation workspace response remains the single data request used to render history. Each object in `history` gains a camel-cased client-facing field normalized from the database payload:

```text
sourceSummary: [
  {
    sourceType: "financial_documents",
    recordCount: 4,
    amountTotal: 450.00
  }
]
```

The database may emit the corresponding snake-cased JSON keys if that matches the existing RPC convention; the application normalization layer must accept the exact deployed representation and expose one consistent rendering shape.

The summary is produced inside `get_financial_reconciliation_workspace` for all returned history rows. The implementation must avoid an HTTP request per history row. A lateral or pre-aggregated query should group `financial_reconciliation_items` for the limited history result and attach one ordered JSON array to each reconciliation.

No new table, persistent summary column, or historical backfill is needed because counts and totals are derived from stored reconciliation item snapshots.

## Migration Strategy

Add a forward migration in `supabase-migrations/` after `2026-08-12-financial-reconciliation-oldest-first-candidates.sql`.

The migration must:

- target the exact installed `get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)` signature;
- preserve candidate ordering, filters, rules, item detail enrichment, audit ordering, and unrelated workspace behavior;
- replace only the history payload construction needed for `sourceSummary`;
- validate exact expected fragments or otherwise install an equivalently safe complete function definition;
- fail clearly if the installed function has unexpected drift;
- be idempotent when run more than once;
- require no data backfill.

The web application may be published before the migration, but history will show `Source details unavailable` until the migration is applied. Applying the migration enables the new summaries immediately for existing and future reconciliations.

## Rendering Boundaries

Use small focused helpers rather than embedding aggregation assumptions in the table renderer:

- Normalize a history row's source-summary payload.
- Format one source entry with the existing source-label and euro-format helpers.
- Format the complete Source cell in the prescribed stable order.
- Render history using the normalized summary while preserving its other cells and interactions.

The client must not recompute totals from current source tables. `amount_snapshot` from reconciliation items is authoritative.

## Error Handling

- Invalid or malformed source-summary entries are ignored rather than breaking the entire history table.
- A present but empty valid summary renders `No records`.
- A missing summary property renders `Source details unavailable`.
- Migration definition drift raises an explicit SQL error instead of partially rewriting the workspace function.
- Existing workspace load errors continue using the page's current error status behavior and must not discard an open reconciliation.

## Testing

### Database behavior

Extend the reconciliation SQL smoke coverage to prove:

- multiple sources are grouped separately;
- counts are correct;
- totals use raw `amount_snapshot` values;
- negative totals remain negative;
- the base source appears first;
- remaining sources follow saved matching-source order;
- unused configured sources are omitted;
- an item removal changes the derived history aggregate appropriately;
- the migration can be applied twice safely.

If PostgreSQL is unavailable locally, retain a minimal static installation safeguard, but treat executable SQL smoke behavior as authoritative and disclose the execution gap.

### Client behavior

Executable tests must call the actual history rendering/formatting functions with controlled history fixtures and verify:

- the exact example-style Source text;
- raw totals are not operator-adjusted;
- missing summary and empty summary have distinct messages;
- Source replaces both former headers;
- Created, Status, Difference, selected styling, and Open remain present.

### Layout behavior

Tests and browser verification must confirm:

- Source and rule hint occupy the first row;
- dynamic filters occupy the second row;
- all Financial Documents filters fit on a wide desktop viewport;
- the named compact fields are narrower than before;
- filters wrap without overlap at narrower widths;
- source switching still refreshes the correct filter set and candidates.

## Deployment and Verification

1. Publish the application and migration file.
2. Apply the new migration after all earlier reconciliation migrations, including the oldest-first candidates migration.
3. Run the reconciliation SQL smoke test in a safe non-production environment.
4. Verify a Started and a Complete reconciliation with at least two used sources.
5. Confirm counts and raw totals against the locked records shown in Current reconciliation.
6. Confirm history layout and workbench filter layout at wide and narrow viewport widths.

Publishing application code does not apply the Supabase migration automatically.
