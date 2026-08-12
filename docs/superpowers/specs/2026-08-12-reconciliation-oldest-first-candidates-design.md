# Reconciliation Oldest-First Candidates Design

## Goal

Always show eligible reconciliation records from the oldest date to the newest date, and simplify the Current reconciliation summary by removing its source sentence and adding a locked-record count.

## Eligible-record ordering

The reconciliation workspace database function remains the authoritative source for candidate ordering. Eligible records are ordered before pagination using:

1. `source_date ASC`;
2. `id ASC` when two records have the same date.

The same ordering is used when aggregating the selected page into the returned candidates JSON. This makes the order deterministic and guarantees that page 1 contains the oldest eligible records across every configured source. The browser renders candidates in the order returned by the server and does not apply a second client-side sort.

Filters, the 2026-01-01 eligibility floor, source rules, record locks, page size, and reconciliation calculations remain unchanged.

## Forward database migration

Add a new file to the normal `supabase-migrations` folder. It updates the currently deployed `public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)` function rather than editing an older migration that may already have run.

The migration reads the installed function definition, replaces both candidate-specific descending order clauses with deterministic ascending clauses, verifies that the expected ascending clauses exist and the candidate-specific descending clauses no longer exist, and then installs the updated definition.

If the installed function does not contain the expected clauses, the migration raises a clear exception and makes no silent partial change. The migration remains safe to run again after a successful application: an already-ascending function passes verification without being changed again.

The migration must preserve every other part of the installed workspace function, including prior source-rule and filter fixes.

## Current reconciliation summary

Remove the sentence that lists the base source and matching sources beneath the status. The sentence is omitted for both Started and Complete reconciliations.

Display one compact summary row in this order:

1. status badge;
2. `#records: N`, where `N` is the total number of currently locked records across every source in the reconciliation;
3. `Dif: AMOUNT`, using the existing formatted difference amount.

The locked-record count comes from the loaded Current reconciliation items and requires no new API or database field. It is displayed for both Started and Complete reconciliations. Replace the existing `Difference:` label with `Dif:`.

Keep all other Current reconciliation content unchanged:

- locked records and their details;
- inline completion controls or completed summary;
- audit trail;
- New reconciliation, Reopen, and Delete actions.

The Reconciliation history table continues to show its Base source and Matching sources columns. Only the duplicate source summary inside Current reconciliation is removed.

Remove any JavaScript variable or rendering logic used only by the deleted summary. Remove summary-paragraph CSS only if it has no remaining consumer in that component.

## Error handling and rollout

No API contract changes are required. Before the database migration runs, the published browser change removes the Current reconciliation source summary, while candidate ordering remains as currently installed. After the migration runs successfully in Supabase, candidate pages become oldest-first.

The migration error must identify that the expected candidate ordering could not be verified so the deployed function can be inspected safely.

## Verification

Automated tests will verify:

- the new migration targets the exact workspace-function signature;
- both candidate pagination and candidate JSON aggregation use `source_date ASC, id ASC`;
- the candidate-specific descending order is removed;
- the migration verifies its result and is idempotent for an already-ascending function;
- Current reconciliation no longer renders the base/matching-source summary;
- Current reconciliation renders status, `#records: N`, and `Dif: AMOUNT` in that order for both Started and Complete states;
- the count matches the complete locked-item list, including records from every source;
- the history table still renders Base source and Matching sources;
- existing reconciliation behavior tests continue to pass.

Manual verification will confirm that eligible records display oldest-to-newest for at least two source selections after the SQL migration is applied, that neither Started nor Complete Current reconciliation panels show the removed source summary, and that their compact summary rows show the correct locked-record counts.

## Out of scope

- User-selectable sort controls.
- Sorting the candidate list in JavaScript.
- Reordering locked records, audit entries, or reconciliation history.
- Changing source rules, calculations, filters, pagination size, or eligibility rules.
- Removing source information from Reconciliation history.
