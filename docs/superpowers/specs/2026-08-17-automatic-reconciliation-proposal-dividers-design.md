# Automatic Reconciliation Proposal Dividers Design

## Goal

Make automatic-reconciliation proposals easier to scan by presenting each proposal as an open, divided row while preserving the existing three-column desktop structure and all current behavior.

## Approved layout

Use the approved Option C, **Open rows with separators**:

- Keep three desktop columns in this order: proposal status/action, base record, destination records.
- Remove the rounded card treatment and enclosing proposal box.
- Retain the colored left status accent.
- Add one clear top boundary to the proposal list and one bottom separator to every proposal, avoiding doubled lines while ensuring adjacent proposals cannot visually blend together.
- Add light vertical separators between the status/action, base-record, and destination-record columns.
- Keep amounts right-aligned and retain the current compact type sizes and wrapping.
- Keep candidate groups and multiple destination records inside the destination area without changing their data or ordering.

## Responsive behavior

At the existing narrow-screen breakpoint, preserve the current single-column stack. Horizontal separators continue to distinguish proposals, while desktop-only vertical separators become horizontal section separators between metadata, base, and destination content.

## Behavior and accessibility

This is a presentation-only change. It must not change:

- proposal selection or default checked state;
- rule selection, analysis, execution, or lifecycle behavior;
- visible proposal statuses or audit evidence;
- keyboard focus, checkbox labels, or semantic proposal articles;
- ambiguous candidate grouping;
- history or manual reconciliation screens.

Focus indication must remain clearly visible around the active proposal row.

## Verification

Add or update executable CSS/markup contract tests to verify:

- the three-column desktop grid remains present;
- proposal rows use horizontal separators without rounded card framing;
- vertical dividers exist between all three desktop columns;
- the narrow layout remains one column with section separators;
- existing automatic-reconciliation UI behavior tests remain green.

Perform a desktop and narrow visual check when an authenticated browser session is available. If not available, record that limitation explicitly.

## Out of scope

- Reconciliation rules, matching logic, API behavior, and database migrations.
- New proposal actions, including closing a run without execution.
- Changes to the content or order of proposal details.
