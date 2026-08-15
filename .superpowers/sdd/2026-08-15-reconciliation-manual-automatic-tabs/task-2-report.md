# Task 2 Report: Reconciliation Manual/Automatic tabs

## Files changed

- `app-main.js`
  - Added explicit `financialReconciliationEntryTab(options)` handling.
  - Made `setView()` accept an entry option, choose Manual by default, focus Automatic after an explicit handoff, and kept the option out of routing/storage.
  - Deferred automatic-rule loading until the Automatic tab is active.
  - Routed Settings “Run batch now” to the Automatic tab without a redundant direct render.
- `tests/reconciliation-automation-ui.test.js`
  - Added entry-option and production `ensureFinancialReconciliationData()` execution coverage.
  - Updated the Settings batch handoff contract and pinned retained run/selection identity through failed rule reloads.
  - Updated the source extractor to preserve an `async` prefix so the lazy-loading test executes the production async function.
- `tests/reconciliation-density.test.js`
  - No Task 2 edit; run as the required regression suite.

## RED evidence

Command:

```powershell
node --test --test-isolation=none --test-name-pattern="entry defaults|loads Manual first|Run batch now stores" tests/reconciliation-automation-ui.test.js
```

Result: exit 1; 0 passed, 3 failed.

- Entry test failed because `financialReconciliationEntryTab` was not defined.
- Run batch test received `view:financial-reconciliation:manual` followed by a direct render instead of the Automatic handoff.
- Lazy-load test received `workspace`, `rules`, `render` for the initial Manual entry, proving the eager rules request.

## GREEN verification

```powershell
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js
```

Result: exit 0; 28 passed, 0 failed.

```powershell
node --test --test-isolation=none tests/reconciliation-density.test.js
```

Result: exit 0; 27 passed, 0 failed.

## Self-review

- Every ordinary Reconciliation entry resolves to Manual; only the exact `"automatic"` option selects Automatic.
- `setView()` applies the tab only after authorization and focuses the Automatic tab only for that explicit handoff.
- Initial Manual loading fetches the workspace but does not fetch rules; Automatic fetches the rules once; subsequent Automatic renders do not refetch.
- Settings retains the returned run and selected proposal IDs while invalidating only the rule catalog, then hands off to Automatic.
- The existing rule-loader failure path only clears rules/loading and leaves the workspace, filters, draft, run, and selection intact; the test now asserts identity retention for run and selected IDs.
- `git diff --check` completed without whitespace errors before commit.

## Commit

`feat: load reconciliation tabs on demand` (the final commit ID is reported at handoff).

## Concerns

None. The requested scoped suites are green. The full repository suite was not run because Task 2 specifies the two focused suites above.
