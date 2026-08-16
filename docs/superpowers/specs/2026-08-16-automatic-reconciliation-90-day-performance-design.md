# Automatic Reconciliation 90-Day Performance Design

**Date:** 2026-08-16

## Problem

The Financial Documents to CGD Bank Statement rule accepts a configurable date window, but its synchronous analysis no longer finishes reliably when the window is widened. Production measurements against the current dataset show:

- 876 eligible Financial Documents for both measurements;
- the candidate phase at 5 days completes in about 5.9 seconds and returns 220 identity-qualified candidates across 195 base records;
- the candidate phase at 7 days completes in about 7.7 seconds and returns 245 identity-qualified candidates across 205 base records;
- the complete 7-day Analyze request returns HTTP 500 because proposal construction, inserts, overlap processing, and response serialization still follow the already expensive candidate phase.

The value 7 is valid. The failure is caused by repeatedly normalizing and scoring overlapping bank rows and by attempting the entire run in one database transaction.

## Goals

- Support configured date windows from 0 through 90 days.
- Preserve the exact rule-version-2 matching results and execution behavior.
- Prevent any one database or HTTP request from containing the complete analysis workload.
- Allow interrupted manual and scheduled analyses to resume without duplicate proposals.
- Show deterministic progress in the Automatic reconciliation tab.
- Keep existing completed runs, reconciliations, provenance, locks, and audit history unchanged.

## Non-Goals

- Do not change identity thresholds, tolerance arithmetic, source operators, candidate limits, or maximum destination group size.
- Do not add new editable rule logic.
- Do not support date windows above 90 days.
- Do not change Manual reconciliation.
- Do not make partial proposals executable.

## Matching Invariants

The optimized implementation must retain all current rule-version-2 behavior:

- base source is `financial_documents`;
- destination source is `import_cgd_extrato_ordem`;
- `fat = 'S'`;
- `payment = 'Banco'` exactly, with no trimming or case folding;
- both source dates are on or after `2026-01-01`;
- date difference is inclusive in both directions;
- locked records are excluded;
- a bank record qualifies through document-number containment, description similarity at `0.60`, or supplier word similarity at `0.70`;
- candidate evidence stores the existing exact scores and thresholds;
- more than 12 identity-qualified candidates produces `candidate_limit` ambiguity;
- combinations contain at most four destination records;
- difference calculation uses the directional operator snapshot and integer-cent tolerance;
- multiple qualifying combinations and cross-base overlap remain ambiguous;
- execution revalidates rule version, configuration snapshot, source snapshots, locks, and calculated difference.

## Search Projection

Add `public.financial_reconciliation_cgd_match_search` with one row per `import_cgd_extrato_ordem` record:

- `source_id uuid primary key`;
- `source_date date not null`;
- `amount numeric`;
- `description text`;
- `normalized_description text not null`;
- `compact_description text not null`;
- `updated_at timestamptz not null default now()`.

The migration backfills the projection from current CGD records. An `AFTER INSERT OR UPDATE OR DELETE` trigger on `import_cgd_extrato_ordem` keeps the projection transactionally synchronized. Updates that change the source ID, date, amount, or description replace the projected row. Deletes remove it.

Indexes:

- B-tree on `(source_date, source_id)` for deterministic date-window scans;
- GIN `gin_trgm_ops` on `normalized_description` for similarity and word-similarity prefilters;
- GIN `gin_trgm_ops` on `compact_description` for document-number containment.

The projection is internal. Revoke access from `public`, `anon`, and `authenticated`; grant only the minimum table privileges required by `service_role` and security-definer functions. Trigger execution remains owned by the migration/function owner.

Index-assisted searches must be supersets of the existing exact identity tests. The final candidate row must still be accepted only after the current exact containment, `similarity(...) >= 0.60`, and `word_similarity(...) >= 0.70` checks. This prevents an index prefilter from changing reconciliation results.

## Paged Candidate Analysis

Add an internal security-definer helper that analyzes one deterministic page of base records. It consumes:

- rule key and version;
- difference tolerance;
- maximum date difference;
- cursor date and cursor UUID;
- page size.

It returns at most the requested number of base records ordered by `(document_date, id)`, plus each base snapshot, exact candidate snapshots, candidate count, and the next cursor. Page size is fixed by server code at 25 records and is not user-configurable.

The existing `financial_reconciliation_automatic_rule_candidates(text, integer, numeric, integer)` contract remains available for compatibility and SQL smoke tests, but it delegates to the same optimized projection logic. Execution revalidation uses a dedicated single-base lookup so executing one selected proposal never rescans every eligible Financial Document.

## Resumable Run State

Extend `financial_reconciliation_automatic_runs` with:

- `analysis_cursor_date date`;
- `analysis_cursor_id uuid`;
- `analysis_processed integer not null default 0`;
- `analysis_total integer not null default 0`;
- `analysis_error_code text`;
- `analysis_error_at timestamptz`.

The existing run status represents lifecycle state. A new run is returned as `running` until all pages and finalization are complete. `analysis_completed_at` remains the authoritative Ready boundary.

`create_financial_reconciliation_automatic_analysis(...)` remains idempotent by actor and client request ID. It snapshots rule configuration, creates or retrieves the run, calculates the total eligible base count, processes at most the first page, and returns the current run. It no longer attempts the complete dataset in one call.

Add `continue_financial_reconciliation_automatic_analysis(p_run_id uuid, p_actor text)`:

1. Lock the run row.
2. Verify that the run is unfinished and that a manual caller owns the run, unless the caller is the internal scheduler identity.
3. Select the next 25 base records after the persisted cursor.
4. Generate and upsert proposals with the existing signatures and statuses.
5. Commit proposal rows, processed count, and cursor together.
6. When no bases remain, run cross-base overlap resolution, persist final counters, set status `ready`, and set `analysis_completed_at`.
7. Return the sanitized public run representation and progress fields.

The cursor advances even when a base produces no identity candidate. A transaction failure advances neither proposals nor cursor. Repeating the same continuation is safe because the run lock, cursor, and existing proposal uniqueness constraints prevent duplicate work.

## API and Background Continuation

The existing POST Analyze action creates the run and returns its current state. Add the POST action `continue_analysis`, accepting only a run UUID. It authorizes Financial Reconciliation app access, binds manual runs to the authenticated actor, and processes one page.

Add GET view `active_run` to return the authenticated actor's newest unfinished manual analysis, if any. This lets a browser reload restore progress without trusting local state.

While an Automatic reconciliation tab remains open, the browser sends sequential continuation requests. It starts the next request only after the preceding response completes. A transient request failure leaves the run resumable and shows an inline Retry action.

The existing cron heartbeat also advances one page of the oldest unfinished analysis before attempting normal scheduled-run work. Manual runs remain manual in provenance; background continuation does not convert their trigger or actor.

## User Interface

While `analysis_completed_at` is empty:

- show `Analyzing {analysisProcessed} of {analysisTotal} records…`;
- keep proposal review, selection, and execution disabled;
- keep the run visible across tab changes;
- continue automatically while the tab is open;
- show an inline retry message if continuation fails.

When the run becomes ready, render the existing filtered proposal review without layout or matching changes. Partial proposals are never presented as executable.

The public run mapper adds `analysisCursorDate`, `analysisCursorId`, `analysisProcessed`, `analysisTotal`, `analysisErrorCode`, and `analysisErrorAt`. Cursor fields are returned only for lifecycle coordination and are never editable by the browser.

## Configuration Limit

Change the API and database validation range for `maxDifferenceDays` from `0–365` to `0–90`. The settings UI uses the same maximum and validation message. The migration converts any existing value above 90 to 90 before installing the new constraint and records the configuration update timestamp. Current production configuration is below the new maximum.

## Error Handling and Concurrency

- Continuation holds a row lock on the run, so UI and cron requests cannot process the same page concurrently.
- Each page is atomic.
- Network timeouts are treated as uncertain outcomes; the client reloads the run before continuing.
- Database exception text remains server-side. Public responses use stable sanitized codes.
- A terminal analysis failure records a sanitized `analysis_error_code` and timestamp. A new Analyze action creates a new idempotency key and run.
- Execution rejects any run whose `analysis_completed_at` is null.
- Source changes after analysis continue to use the existing stale-proposal path.

## Migration and Compatibility

Implement one forward-only, idempotent migration in the normal `supabase-migrations` folder. It must:

- create/backfill/index the projection;
- install synchronization triggers;
- add resumable-run columns and the 90-day constraint;
- install paged, single-base, create, continue, active-run, and cron-support RPC behavior;
- preserve existing function privileges and fixed `search_path` protections;
- notify PostgREST to reload its schema.

Completed automatic runs require no data rewrite and remain complete. The migration transitions each unfinished pre-migration run to `failed` with the stable reason `analysis_upgrade_restart_required`, because it has no trustworthy page cursor; the user can then start a new analysis.

## Verification

Automated Node source/contract tests and transactional PostgreSQL smoke tests must prove:

- 0 and 90 are accepted; 91 is rejected by API and database validation;
- day 90 is inclusive and day 91 is excluded;
- rule-version-2 candidate and evidence results match the pre-optimization implementation on the same fixtures;
- exact `Banco` eligibility remains unchanged;
- projection backfill and insert/update/delete synchronization are correct;
- pages are ordered by date then UUID and contain no more than 25 bases;
- retries do not duplicate proposals or skip a cursor position;
- concurrent continuations cannot process the same page;
- incomplete runs cannot execute proposals;
- closing/reopening can retrieve and resume the actor's active run;
- cron and UI continuation safely share the same run lock;
- final overlap, ambiguity, skipped audit rows, counts, and history remain correct;
- execution revalidation uses the single-base path;
- `anon` and `authenticated` cannot call internal helpers and `service_role` retains required execution privileges.

Production verification uses the current dataset with a 90-day window. It must complete the run without HTTP 500, statement timeout, duplicate proposal, or matching-result regression. Each continuation request must remain below the active database/API timeout; the complete current-dataset run should finish within two minutes while the tab remains open.
