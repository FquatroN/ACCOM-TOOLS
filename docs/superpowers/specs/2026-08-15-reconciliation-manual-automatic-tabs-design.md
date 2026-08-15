# Reconciliation Manual and Automatic Tabs Design

## Objective

Simplify the Reconciliation screen by separating manual work from automatic rule execution and proposal review. The existing Reconciliation application remains one top-level option. Inside it, users switch between **Manual reconciliation** and **Automatic reconciliation** tabs, while a single shared reconciliation history remains visible below both.

## Screen Structure

The Reconciliation title and introductory text remain at the top of the page. Immediately below them, add an accessible horizontal tab list with two tabs:

1. **Manual reconciliation**
2. **Automatic reconciliation**

The Manual reconciliation panel contains:

- source selection and rule hint;
- all eligible-record filters;
- eligible records and manual Start/Add actions;
- the current reconciliation basket and lifecycle actions.

The Automatic reconciliation panel contains:

- enabled rules that allow manual analysis;
- Analyze actions;
- proposal review and evidence;
- proposal selection controls;
- Execute selected;
- automatic run status, outcome counts, and errors.

The existing Reconciliation history card is outside both tab panels and directly below them. It is therefore visible and usable regardless of the selected tab. The history remains a combined history of user and automatic reconciliations, distinguished by the existing Origin presentation.

Rule configuration, scheduling, and **Run batch now** remain under **Settings → Reconciliation → Automatic reconciliation**. The new application tab is only for executing rules and reviewing proposals.

## Navigation and State

Normal navigation into Reconciliation always selects **Manual reconciliation**, even if Automatic reconciliation was selected during the previous visit.

The Settings **Run batch now** flow is an explicit exception. After it creates an analysis run, it navigates to Reconciliation with **Automatic reconciliation** selected so the new proposals are immediately visible.

Switching tabs does not discard state:

- Manual filters, selected source, page, current reconciliation, and completion-comment draft remain intact.
- Automatic rule data, current run, proposals, checkbox selections, outcome summary, and status remain intact.
- Tab switching alone does not issue a new API request.

The selected application tab is session state only. It is not persisted across navigation away from Reconciliation and does not require a new URL route.

## Loading and Data Flow

On normal entry, the Manual workspace loads and renders first. Automatic rule data is loaded lazily when the user first selects Automatic reconciliation.

When Settings **Run batch now** creates a run, the returned run and its default executable-proposal selection are stored before navigation. Reconciliation then opens on Automatic reconciliation and renders that run. The automatic rules catalog may load separately if it is not already available.

After a successful manual or automatic reconciliation mutation, the shared workspace/history refreshes. A history refresh must not clear Manual or Automatic tab-local state.

## Error Handling

Errors stay with the feature that produced them:

- Manual workspace errors appear in the Manual panel and preserve any open reconciliation.
- Automatic rule-loading, analysis, and execution errors appear in the Automatic panel and preserve the current run and selections when safe.
- A failure in one panel must not clear or replace state in the other panel.
- Shared-history refresh failures use the existing reconciliation status behavior without hiding either panel's retained work.

## Accessibility and Responsive Layout

The tab controls use `role="tablist"`, `role="tab"`, `role="tabpanel"`, `aria-selected`, `aria-controls`, and managed `tabindex` values. Left and Right Arrow keys move focus and activate the adjacent tab. Click activation and keyboard activation use the same state transition.

Only the active tab panel is exposed. The inactive panel uses the `hidden` attribute. Focus remains predictable after tab changes and after the Settings handoff.

On narrow screens, the two tabs may wrap or share the available width, while retaining visible focus styles and touch-friendly targets. The current responsive layouts inside the Manual and Automatic panels remain in effect.

## Implementation Boundaries

This change is a client-side information-architecture change. It does not alter database tables, reconciliation calculations, source rules, automation definitions, RPC contracts, API endpoints, permissions, or scheduled execution.

The existing manual and automatic renderers should remain independently testable. A small tab-state controller determines which panel is visible and which data loader is required. Shared-history rendering remains independent of the selected panel.

## Acceptance Criteria

- Reconciliation displays Manual reconciliation and Automatic reconciliation tabs.
- Normal entry always selects Manual reconciliation.
- The Manual panel contains the existing manual workbench and current reconciliation.
- The Automatic panel contains rule execution and proposal review only.
- Configuration and scheduling remain in Settings.
- Shared reconciliation history is visible from either tab.
- Switching tabs preserves both panels' state and does not reload already loaded data.
- First activation of Automatic reconciliation lazily loads its rule catalog.
- Settings **Run batch now** navigates directly to the Automatic panel with the returned run visible.
- Manual and automatic errors remain scoped to their respective panels.
- Tabs support mouse, touch, Left Arrow, and Right Arrow activation with correct ARIA state.
- The layout remains usable at desktop and narrow-screen widths.
- Automated tests cover default selection, tab switching, keyboard behavior, lazy loading, state retention, Settings handoff, shared history visibility, and responsive structure.

## Non-goals

- Separate application routes or deep links for the two tabs.
- Separate manual and automatic history tables.
- Moving automatic configuration or scheduling out of Settings.
- Changing reconciliation matching, execution, locking, audit, or completion behavior.
