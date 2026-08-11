# Reconciliation source rules design

**Date:** 2026-08-11  
**Status:** Approved design

## Goal

Simplify the Reconciliation workbench to one Source selector while allowing administrators to define the allowed, directional source matches and the arithmetic operator for each direction.

## Workbench

- Replace **Base source**, **Match with**, and **Browse source** with one **Source** selector.
- Selecting a source reloads Eligible records for that source, retaining the existing date and other applicable filters.
- Show a read-only compatibility hint beneath the selector, for example: `Matches: CGD Bank Statement (-), CGD Credit Card (-)`.
- Before a reconciliation starts, Source contains every configured source. Clicking **Start** on an eligible record creates a reconciliation using the selected source as the base and the saved outbound rules for that source.
- While a reconciliation is started, Source contains the reconciliation base source and the sources configured by its captured rules. Selecting one refreshes Eligible records without a dialog.
- A source unavailable for the current reconciliation cannot be selected or added.

## Settings: Reconciliation

- Add a **Reconciliation** area in Settings for administrators.
- The editor presents one source at a time and allows the administrator to select compatible target sources.
- Each selected source-to-source rule stores a required `+` or `-` operator.
- Rules are fully directional and independent. Creating `A -> B (-)` does not create or modify `B -> A`; the reverse rule must be entered separately and may use either operator.
- Rules cannot point from a source to itself.
- Removing a rule affects future reconciliations only. Existing reconciliations remain intact.

## Calculation and audit behavior

- At reconciliation creation, capture the base source and the configured outbound rules, including each operator, as a rule snapshot.
- The difference is calculated from the base-record amount plus or minus each added item according to the captured operator for that item's source.
- Starting a reconciliation requires at least one configured outbound rule for the selected base source.
- Adding an item requires an applicable captured directional rule; API validation enforces this independently of the UI.
- Settings changes never recalculate, invalidate, or change the audit history of started or completed reconciliations.

## Data and API design

- Add a dedicated reconciliation-source-rules table keyed by base source and matching source, with an operator constrained to `+` or `-`.
- Extend the settings API to read and save these rules.
- Replace the hard-coded source-compatibility map in the Reconciliation UI and API with the stored rules.
- Persist the rules snapshot with each reconciliation so historical behavior is deterministic.
- Preserve the existing record-locking, complete/force-complete, reopen, and audit-trail behavior.

## Error handling

- Show a clear empty state when a source has no configured outbound rules; Start remains unavailable.
- Reject invalid sources, missing rules, invalid operators, self-rules, and attempts to add a source outside the reconciliation snapshot.
- Keep the current reconciliation open if a refresh or action fails, and show the server error in the existing reconciliation error area.

## Verification

- Test Settings rule validation and persistence for independent directions and independent operators.
- Test automatic reconciliation creation from a selected source with no pairing dialog.
- Test eligible-record refresh when the Source selector changes.
- Test difference calculations for both `+` and `-` rules.
- Test API rejection of missing or unauthorized directional rules.
- Test that a settings-rule change does not modify a reconciliation created under an earlier snapshot.
