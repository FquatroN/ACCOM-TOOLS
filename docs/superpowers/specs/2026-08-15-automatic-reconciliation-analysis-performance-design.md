# Automatic Reconciliation Analysis Performance Design

## Problem

Pressing **Analyze** for the Financial Documents to CGD Bank Statement rule fails with the sanitized UI message `Unexpected server error.` A read-only production reproduction of `financial_reconciliation_automatic_rule_candidates` returned PostgreSQL `57014: canceling statement due to statement timeout`.

The production input currently contains 1,347 eligible financial documents and 2,461 dated CGD bank-statement records. The current query repeatedly normalizes text and calculates fuzzy similarity while building the seven-day candidate join, exceeding the Supabase statement timeout.

## Decision

Add a new, idempotent Supabase migration that replaces only the candidate-query implementation. Do not edit or depend on reapplying an already-deployed migration.

The replacement must preserve the approved rule contract exactly:

- source direction: Financial Documents to CGD Bank Statement;
- maximum date difference: configurable, currently seven days;
- allowed amount difference: configurable, currently zero euros;
- at least one identity signal: invoice/document number, description similarity, or supplier-name similarity;
- existing record locks and the 2026-01-01 eligibility floor;
- deterministic proposal ordering and evidence payloads.

## Query Design

1. Add idempotent B-tree indexes for `financial_documents(document_date)` and `import_cgd_extrato_ordem(data)` where an equivalent usable index is not already present.
2. Build eligible document and bank source sets separately.
3. Materialize normalized document number, description, and supplier text once per source row instead of recalculating those expressions for every reference.
4. Join the materialized sets by the configured date window before evaluating fuzzy identity scores.
5. Materialize the scored candidate set so the same similarity values feed filtering, evidence, and aggregation without repeated function calls.
6. Preserve the current output columns, JSON shapes, thresholds, ordering, and security-definer grants so callers require no API or UI changes.

## Alternatives Considered

- **Increase `statement_timeout`: rejected.** It hides the scaling defect and makes scheduled batches increasingly fragile.
- **Precompute permanent normalized columns: deferred.** It may be useful at much larger scale, but adds triggers/generated-column lifecycle and broader schema coupling that are unnecessary for the current volume.
- **Materialized source/scoring CTEs plus date indexes: selected.** It is the smallest change that removes repeated work while preserving the public RPC contract.

## Failure Handling and Rollout

The migration is safe to re-run and ends with a PostgREST schema reload notification. If migration execution fails, the existing function remains available because the SQL editor transaction rolls back. No reconciliation runs or proposals are created by the migration itself.

After applying the migration, run a read-only timed call to `financial_reconciliation_automatic_rule_candidates` with the production rule parameters. It must return successfully within the configured Supabase statement timeout before testing **Analyze** in the application.

## Testing

- Add source-contract tests that fail against the current implementation and require materialized source/scoring stages, date indexes, the unchanged thresholds and eligibility rules, stable ordering, grants, and schema reload.
- Run the focused automation migration tests and the full Node test suite.
- Apply the new migration in Supabase.
- Re-run the same read-only production candidate RPC used to reproduce `57014` and record its elapsed time.
- Press **Analyze** and confirm proposals or auditable skipped/ambiguous results render without a server error.

## Scope

In scope: one new Supabase migration and its automated/source-contract coverage. Out of scope: UI changes, API response changes, rule-threshold changes, increased database timeouts, and automatic execution behavior.
