# Compact Automatic Reconciliation Proposal Columns

## Goal

Reduce the vertical height of automatic reconciliation proposal results by moving proposal-level information into a dedicated first column and removing the separate proposal header and footer rows.

## Scope

This is a presentation-only change to the automatic reconciliation proposal review area. It does not change proposal selection rules, reconciliation calculations, persistence, APIs, database functions, or which proposal statuses are displayed.

## Desktop layout

Each proposal renders as one three-column row:

1. **Proposal column**: selection control when applicable, lifecycle status, difference, allowed difference, and rule version.
2. **Base record column**: source label, amount, date, optional document and supplier metadata, description, identity evidence when present, and collapsible record ID.
3. **Destination column**: destination record or candidate-group details, amount and operator, date and other metadata, matching evidence, and collapsible record ID.

The current proposal `<header>` and `<footer>` bands are removed. Their information moves into the proposal column so no lifecycle or audit context is lost.

## Selection behavior

- An enabled checkbox is shown only for an executable proposal whose status is `proposed`.
- Ambiguous proposals do not show a checkbox; the first column shows `Ambiguous`.
- Completed, stale, and failed proposals do not show a checkbox; the first column shows the persisted status and reconciliation information.
- Existing selection state, checkbox data attributes, and accessible names remain unchanged for executable proposals.

## Responsive behavior

At viewport widths below 700 px, the proposal column becomes a compact top strip. Base and destination records then stack vertically below it. Text, evidence, descriptions, and record IDs continue to wrap without horizontal overflow.

## Accessibility

- Preserve focus indication on each proposal.
- Preserve the accessible checkbox label for executable proposals.
- Keep lifecycle status available as visible text rather than color alone.
- Keep immutable record IDs accessible through the existing collapsed details control.

## Testing

Behavior tests will prove that:

- only executable `proposed` entries render a checkbox;
- ambiguous, completed, stale, and failed entries render no checkbox;
- status, difference, allowed difference, rule version, evidence, and record IDs remain present;
- the old proposal header and footer bands are absent.

CSS contract tests will prove the desktop three-column structure and the below-700-pixel stacked layout. The focused automatic reconciliation UI and density tests, followed by the complete Node test suite, must pass before integration.

## Out of scope

- Changing automatic reconciliation rules or matching logic.
- Changing which proposal statuses appear in active or completed runs.
- Changing reconciliation execution behavior or API contracts.
- Altering the general page layout outside proposal cards.
