# Automatic Reconciliation Banco Rule Version 2 Design

## Purpose

Restrict the managed **Financial Documents to CGD Bank Statement** automatic reconciliation rule to Financial Documents whose stored Payment value is exactly `Banco`. Preserve historical version-1 audit evidence, introduce the restriction as managed rule version 2, and make automatic-analysis proposal rows more compact and focused.

## Scope

This change covers:

- version 2 of the existing `financial_documents_cgd_bank_statement` managed rule;
- exact base eligibility `financial_documents.payment = 'Banco'`;
- version-aware analysis and execution validation;
- proposal-list visibility rules for active and finished runs;
- the approved compact paired-card analysis layout;
- database, source-contract, frontend behavior, density, and regression tests.

It does not add a configurable Payment condition, a generic eligibility-expression engine, a new rule key, a new API endpoint, or direct browser database access.

## Managed Rule Versioning

The stable rule key remains `financial_documents_cgd_bank_statement`. A new definition row with version `2` is added; version `1` remains immutable for historical runs and completed reconciliations.

Version 2 retains all existing behavior except the new Payment eligibility condition:

- base source: `financial_documents`;
- destination source: `import_cgd_extrato_ordem`;
- existing difference allowance and maximum date difference configuration;
- existing date floor, `fat = 'S'`, unlocked-record, identity, similarity, ambiguity, amount, directional-operator, and maximum-combination behavior;
- exact additional base predicate: `d.payment = 'Banco'`.

The predicate is intentionally case-sensitive and whitespace-sensitive. Values such as `BANCO`, ` banco `, blank, or any other Payment value do not qualify.

The read-only managed definition and business explanation explicitly state that the Financial Document Payment must equal `Banco`. The migration copies the current administrator configuration to version 2 without changing enabled state, manual/scheduled flags, difference allowance, maximum date difference, or priority. Reapplication is idempotent.

## Analysis Data Flow

Manual and scheduled analyses snapshot version 2 and its existing administrator configuration. Candidate loading filters Financial Documents by `payment = 'Banco'` before date-window joins, identity scoring, candidate-limit evaluation, and combination building.

Consequences:

- a qualifying `Banco` Financial Document follows the existing matching algorithm unchanged;
- a non-`Banco` Financial Document is absent from candidate processing;
- excluded Financial Documents create neither automatic proposals nor skipped/no-match proposal records;
- eligible `Banco` documents for which no destination combination qualifies may still be stored as skipped for audit, but their individual rows are not displayed in the workbench;
- counts continue to reflect the complete persisted run, including skipped and deselected records.

## Execution and Stale Handling

Execution accepts version 2 of this managed rule and retains the existing transactional revalidation, row locking, snapshot checks, operator checks, and reconciliation creation flow.

Before creating a reconciliation, execution re-runs version-2 candidate eligibility. If the base Payment changes away from exact `Banco` after analysis, revalidation cannot reproduce the proposal; the proposal becomes `stale`, no reconciliation is created, and the existing safe stale reason is returned.

Existing unexecuted version-1 proposals cannot reconcile after the rollout. An attempt to execute one atomically changes it to `stale` with `rule_version_changed`. The user must run a new version-2 analysis. Completed version-1 reconciliations, definitions, snapshots, evidence, and audit history remain unchanged.

The migration preserves current service-role-only execution privileges, security-definer search paths, RLS boundaries, and existing public API/RPC response contracts.

## Proposal Visibility

One pure frontend selector derives visible proposal rows from the authoritative persisted run without mutating it.

For an active, unfinished analysis run:

- show `proposed` rows, whether currently checked or unchecked;
- show `ambiguous` rows;
- hide `skipped` rows.

For a finished run:

- show only selected execution outcomes whose persisted status is `completed`, `stale`, or `failed`;
- hide `ambiguous`, `skipped`, and finalized `deselected` rows.

Unchecked proposed rows remain visible and can be selected again until the user activates **Execute selected**. After the run is finalized, deselected proposals remain in the audit data and aggregate counts but have no individual row.

The result-summary counters remain derived from all persisted proposals and execution outcomes. Skipped and deselected counts therefore remain visible even though their detailed rows are hidden.

## Compact Paired-Card Layout

The selected visual direction is **Option A: compact paired cards**.

Each visible proposal retains:

- selection control and status;
- side-by-side base and destination cards when width permits;
- source label;
- compact date, supplier/document, and description summary;
- amount and signed operator;
- matching evidence;
- difference, allowed tolerance, managed rule name, and version;
- access to immutable record identifiers for audit/troubleshooting.

Density is reduced through smaller outer and inner padding, tighter gaps, smaller secondary text, compact line height, and collapsed metadata presentation. Long descriptions wrap safely. Narrow screens stack the paired cards without horizontal clipping, and checkbox labels and focus states remain accessible.

The compact styling applies consistently to every proposal row that passes the visibility selector.

## Error Handling

- Unknown rule keys and unsupported versions retain the existing safe validation failures.
- Version-1 execution attempts become stale rather than failing unsafely or creating a reconciliation.
- Payment changes after analysis become stale through normal source/candidate revalidation.
- A failed analysis or execution retains the currently displayed authoritative run and safe error mapping.
- UI filtering changes presentation only; it does not delete proposals, rewrite counts, or weaken audit persistence.
- Migration guards reject an unexpected installed function definition instead of silently overwriting incompatible logic.

## Migration and Compatibility Strategy

Add a new normal migration in `supabase-migrations/`. It must patch or replace the latest candidate and execution functions, including the current indexed candidate lookup behavior, rather than restoring an older implementation.

The migration must:

1. seed the immutable version-2 managed definition;
2. atomically move the current rule configuration from version 1 to version 2 while preserving editable values;
3. update analysis candidates to recognize version 2 and require exact `payment = 'Banco'`;
4. update execution to accept version 2 and make version-1 proposals stale;
5. retain function signatures, search paths, role grants, and schema reload behavior;
6. be safe to run more than once.

No frontend/database contract change is required beyond the rule version and existing definition data already returned by the automation APIs.

## Testing

Strict RED/GREEN TDD covers production changes.

Database and source-contract coverage proves:

- version 2 is seeded with the exact Payment condition and explanation;
- configuration values move to version 2 unchanged;
- reapplication is safe;
- exact `Banco` bases are eligible;
- `BANCO`, ` banco `, blank, and other values are excluded before candidate generation;
- non-`Banco` documents create no proposals or skipped rows;
- changing Payment after analysis makes execution stale and creates no reconciliation;
- pending version-1 execution becomes stale with `rule_version_changed`;
- existing identity thresholds, ambiguity, locking, date, amount, operator, and indexed lookup behavior remain intact;
- security-definer settings and service-role-only privileges remain intact.

Executable frontend tests prove:

- active runs render proposed and ambiguous rows only;
- checked and unchecked proposed rows stay visible before execution;
- finished runs render completed, stale, and failed selected outcomes only;
- skipped, ambiguous-after-finish, and deselected rows are hidden;
- aggregate counters continue to include hidden skipped and deselected records;
- compact paired-card markup escapes all source values and retains required audit details;
- desktop and narrow-screen CSS preserve readability, wrapping, focus, and reachable controls.

Run the complete Node regression suite, syntax checks, and diff checks. Run the automation SQL smoke against Supabase/PostgreSQL when database access is available; if not available locally, record it as a mandatory external gate.

## Rollout

1. Apply the new Supabase migration.
2. Publish the application changes.
3. Run the SQL smoke contract in a disposable or development environment.
4. Run one authenticated manual version-2 analysis and confirm non-`Banco` documents are absent.
5. Review compact active and finished run results on desktop and narrow screens.
6. Only then rely on the next scheduled batch.
