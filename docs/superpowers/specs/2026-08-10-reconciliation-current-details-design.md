# Reconciliation current-details panel

## Goal

Show concise source details beneath every locked record in the Current reconciliation panel, while making the panel’s body content more compact.

## Data design

- Enrich the existing `get_financial_reconciliation_workspace` response with `source_date`, `description`, and `supplier` for every reconciliation item.
- Resolve these values at workspace-load time through the existing `financial_reconciliation_source(source_type, source_id)` function.
- Use a left lateral lookup so a historical locked item still appears if its source record is unavailable; its detail values are then empty rather than causing the workspace load to fail.
- This live lookup applies to existing and newly created reconciliations. It does not change lock records, amounts, reconciliation calculations, or audit data.

## Panel layout

- Keep each primary locked-record row as source label, amount, and Remove action.
- Add a smaller, full-width line beneath it in the order `date · supplier · description` when a supplier exists; otherwise use `date · description`.
- Omit any unavailable detail rather than displaying placeholder text.
- Allow the details line to wrap cleanly.
- Reduce typography only for Current reconciliation body content: summary copy, locked-record rows/details, section headings, and audit entries. Keep the card title and action buttons at their current readable size.

## Verification

- Add a migration that updates the workspace RPC and preserves the existing source-unavailable behaviour.
- Add source-contract tests for the enriched item response and client detail rendering, including supplier omission.
- Run the complete Node test suite and SQL smoke checks.
