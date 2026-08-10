# Reconciliation eligible-records table layout

## Goal

Make the Eligible records table readable at the current desktop layout by wrapping long record text, preventing Start/Add actions from overlapping dates, and reducing the visual weight of Start actions and status pills.

## Scope

- Apply only to the Eligible records table and its header action in the Reconciliation view.
- Do not change reconciliation data, filtering, selection, locking, or completion behaviour.
- Do not alter the Current reconciliation card or reconciliation history table.

## Layout and wrapping

- Give the first action column a fixed `4.5rem` width. The row-level Start/Add button must stay inside this column and may not overlap the date cell.
- Mark the generated action, date, amount, and status cells with reconciliation-specific classes. This keeps table styling stable when the source adds optional columns.
- Let description and optional source-detail cells wrap at word boundaries and, where needed, within long unbroken strings.
- Keep date and amount cells on one line. Give Date a `6rem` width, Amount a `5.8rem` width, and Status a `6.8rem` width.
- Remove the existing `1%` first-column sizing rule, which is the source of the overlap.

## Compact controls

- Reduce the header “Choose Start…” button to `0.84rem`, with proportionally smaller padding.
- Reduce row-level Start/Add actions to `0.72rem` on desktop and `0.70rem` at `768px` and below.
- Reduce status-pill text from `0.78rem` to `0.70rem`, with matching smaller padding; keep each status label on one line.
- Preserve the existing mobile-safe `16px` filter input/select text size.

## Verification

- Add a source-contract test that verifies the semantic cell classes and the action/date/amount/status layout declarations.
- Confirm the generated row keeps the Start/Add button in its own action cell.
- Run the complete Node test suite and `git diff --check`.
