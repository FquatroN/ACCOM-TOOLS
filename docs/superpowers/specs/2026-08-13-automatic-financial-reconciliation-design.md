# Automatic Financial Reconciliation Design

## Purpose

Add deterministic, auditable automatic reconciliation without weakening the existing manual reconciliation controls. Administrators will operate managed matching rules whose logic is implemented, tested, versioned, and displayed read-only. They may change only approved operational parameters. A rule can be analyzed interactively by a user or included in one shared daily batch.

## Goals

- Automatically complete only deterministic matches produced by managed rules.
- Support one base record matched with one or more records from one or more destination source types.
- Require at least one rule-defined identity signal in addition to date and amount controls.
- Let administrators analyze a manually triggered rule, review proposals, deselect groups, and execute the remaining selection.
- Run selected rules once per day at an administrator-configured time in `Europe/Lisbon`.
- Distinguish user-created reconciliations from automatic reconciliations and distinguish manual from scheduled automatic triggers.
- Preserve item locks, lifecycle behavior, comments, arithmetic, and the existing audit trail.

## Non-goals

- A free-form or no-code rule builder.
- AI or nondeterministic matching.
- Automatically resolving ambiguous combinations.
- Changing existing manual reconciliation source rules or historical reconciliations.
- Allowing an administrator to edit the implemented identity predicates or similarity algorithm.

## Terminology

- **Managed rule definition:** Read-only, versioned matching logic delivered through an application release and database migration.
- **Rule configuration:** Administrator-editable operational values for a managed definition.
- **Analysis run:** A read-only evaluation that creates proposals but does not lock or mutate source records.
- **Proposal:** One unique group of source records that satisfies a rule at analysis time.
- **Execution batch:** A manual or scheduled attempt to execute proposals using snapshotted rule definitions and configurations.
- **Identity signal:** A deterministic comparison, such as document-number containment or normalized description similarity, defined by a managed rule.

## Managed Rule Model

Each definition has a stable rule key and an integer version. A new implementation of matching behavior creates a new version; it never silently changes the meaning of an older execution snapshot.

The read-only definition contains:

- Rule key, version, display name, and business description.
- Base source type.
- One or more permitted destination source types.
- Required source eligibility conditions.
- Identity predicates and their explicit `AND`/`OR` structure.
- Normalization and similarity algorithms with fixed thresholds.
- Combination limits, uniqueness criteria, and ambiguity handling.
- A human-readable explanation suitable for Settings and proposal evidence.

The editable configuration contains:

- `enabled`.
- `allow_manual_execution`.
- `include_in_scheduled_batch`.
- `difference_allowed`, stored as a non-negative currency value with two decimal places.
- `max_difference_days`, stored as a non-negative whole number.
- `priority`, unique and reorderable among enabled rules.

The difference allowance is an absolute tolerance: `1.00` accepts calculated differences from `-1.00` through `+1.00`, inclusive. The date allowance is symmetric and inclusive: `7` accepts destination dates from seven calendar days before through seven calendar days after the base date.

Every analysis and execution snapshots both the definition version and effective configuration. Later configuration or implementation changes cannot change historical evidence.

## Initial Managed Rule

The first managed rule is **Financial Documents to CGD Bank Statement**:

- Base: `financial_documents`.
- Destination: `import_cgd_extrato_ordem`.
- Default difference allowance: `0.00`.
- Default maximum date difference: `7` days.
- At least one of these identity branches must pass:
  1. The normalized CGD description contains the normalized Financial Document `doc_number`.
  2. The normalized Financial Document description and normalized CGD description meet the rule's fixed similarity threshold.
  3. The normalized CGD description and normalized Financial Document supplier name meet the rule's fixed similarity threshold.
- A missing or blank value cannot satisfy its branch.
- The matching evidence identifies the exact branch or branches that passed.

Normalization is deterministic: Unicode case-folding; diacritic removal; punctuation-to-space conversion; whitespace collapse; and comparison of significant alphanumeric tokens. Document-number comparison also removes common separators before containment testing. The implementation must publish the fixed significant-token and similarity thresholds in the read-only rule explanation and cover them with boundary tests before the rule is enabled. Thresholds are part of the managed definition, not editable configuration.

Additional rules may use several destination source types, but each rule must explicitly define those sources and its identity predicates. No rule may search arbitrary configured sources.

## Candidate and Combination Algorithm

For each enabled rule, the engine processes eligible base records from oldest to newest with a stable ID tie-breaker.

1. Exclude records dated before `2026-01-01`, locked records, deleted records, and records already included in another reconciliation.
2. Load eligible records only from the destination sources named by the managed definition.
3. Keep destination records inside the inclusive configured date window.
4. Evaluate the managed identity predicates. At least one required identity branch must pass.
5. Build combinations subject to the definition's fixed maximum group size and source cardinality limits.
6. Apply the existing snapshotted directional `+` or `-` operators and calculate using integer cents.
7. Keep combinations whose absolute difference is less than or equal to `difference_allowed`.
8. Produce a proposal only when exactly one complete combination qualifies for the base record. Two or more qualifying combinations make the base record ambiguous; all such combinations are reported but none is automatically executable.
9. Process rules in ascending administrator-defined priority. Records claimed by a higher-priority executed rule are unavailable to lower-priority rules.

The engine never chooses a "best" candidate from multiple qualifying combinations. Similarity scores determine whether an identity predicate passes; they do not break ambiguity ties.

## Manual Analysis and Execution

Rules with manual execution enabled appear in the Reconciliation app.

1. The user clicks **Analyze** for one rule. Analysis does not lock or change source records.
2. The result displays executable proposals, ambiguous groups, and skipped records with reasons.
3. Executable proposals are selected by default. The user may deselect individual proposals or use select-all and clear-all controls.
4. The user clicks **Execute selected matches**.
5. Each selected proposal is re-evaluated against current records, locks, rule version, configuration snapshot, and arithmetic.
6. A valid proposal starts, adds and locks its items, recalculates, completes, comments, and audits in one database transaction.
7. A stale or no-longer-valid proposal is skipped with a safe explanation and no partial mutation.

An administrator's **Run batch now** action performs a combined manual analysis of the currently batch-enabled rules. It still requires proposal review and **Execute selected matches**; only the unattended scheduled trigger executes without review.

## Scheduled Batch

Settings contains one shared schedule:

- Enabled or disabled.
- One daily local time in `Europe/Lisbon`.
- Last execution, next expected execution, and last result.

The scheduler invokes one protected application entry point. The database atomically claims at most one execution for a Lisbon calendar date and configured schedule time, so retries cannot create duplicate batches and overlapping runs cannot occur. The exact scheduler adapter may wake more frequently than once daily, but only one eligible daily batch is claimed and executed.

The scheduled batch:

- Includes only enabled rules marked for scheduled execution.
- Snapshots the ordered definitions and configurations at batch start.
- Analyzes and executes rules in priority order.
- Automatically executes only unique, valid proposals.
- Skips ambiguous and stale proposals.
- Continues with other proposals after an isolated proposal failure.
- Records counts and reasons for proposed, completed, ambiguous, stale, failed, and skipped work.

## Atomicity, Locking, and Idempotency

Each proposal executes in its own database transaction. Starting the reconciliation, locking and adding all records, calculating the difference, completing, storing the comment, and writing audit entries either all succeed or all roll back.

Execution uses the existing database reconciliation action contract rather than duplicating lifecycle rules in the browser or scheduler. It must recheck record eligibility and acquire the same source-record locks used by manual reconciliation.

Each batch and proposal has an idempotency key enforced by a unique database constraint. Retrying a request returns the existing result or safely resumes eligible pending work; it cannot execute a proposal twice.

## Difference Comments

When the executed difference is non-zero but inside the configured tolerance, the system supplies the mandatory comment automatically. Its stable content includes:

- Managed rule key, name, and version.
- Actual difference.
- Allowed difference.
- Manual or scheduled trigger.
- Execution batch identifier.

For a zero difference, this automatic explanation is still retained in structured audit metadata even though the completion comment is optional.

## Origin and Audit Data

Reconciliations gain structured provenance rather than inferring origin from the actor string:

- `origin`: `user` or `automatic`.
- `automatic_trigger`: `manual` or `scheduled` when origin is automatic.
- Automatic rule key and version.
- Execution batch identifier.

The visible badges are:

- **User**.
- **Automatic · Manual**.
- **Automatic · Scheduled**.

The origin badge appears in Current reconciliation and Reconciliation history. The audit trail retains the actor, trigger, rule and configuration snapshots, proposal evidence, difference, tolerance, generated comment, timestamps, and final outcome.

## Persistence Boundaries

The feature uses four bounded persistence concerns:

1. A managed definition catalog seeded and versioned by migrations; application matching strategies are keyed by rule key and version.
2. Administrator configurations and one shared schedule.
3. Analysis/execution batches and proposal records containing immutable snapshots and outcomes.
4. Reconciliation provenance linked to an execution batch and rule version.

Direct access remains restricted. Authenticated application users use authorized API endpoints; service-role database functions perform validated mutations. Scheduled calls use a protected service credential and never expose it to the browser.

## Settings Interface

Settings → Reconciliation gains two tabs:

- **Source rules** keeps the existing directional source/operator editor.
- **Automatic reconciliation** contains the shared schedule and managed automatic rules.

The automatic tab shows:

- Schedule enablement, Lisbon time, last execution, next expected execution, and **Run batch now**.
- Rules in priority order with reorder controls.
- Enabled, Manual, and Scheduled switches.
- Editable difference and day allowances.
- Read-only friendly source names, technical source names as secondary text, full rule explanation, and version.

Configuration updates are atomic and validated. Invalid tolerances, duplicate priorities, unknown definitions, or enabling a rule without a corresponding directional source rule are rejected without partially saving changes.

## Reconciliation Interface

The Reconciliation app adds an **Automatic reconciliation** area listing enabled, manually executable rules and their current parameters.

Analysis results show:

- One selectable group per executable proposal.
- Base and destination source, date, description, supplier where relevant, and amount.
- Applied source operators.
- Identity branch or branches that passed.
- Calculated difference and configured allowance.
- Rule name and version.
- Ambiguity, stale, or skip explanations.
- Select all, clear all, individual selection, and **Execute selected matches**.

After execution, a summary reports completed, skipped, stale, ambiguous, and failed proposals. Individual failures do not erase successful results from the same batch.

## Authorization

- Users with existing reconciliation access may view automatic-origin reconciliations and their audit evidence.
- Only administrators with Reconciliation Settings permission may change configurations, reorder priorities, edit the shared schedule, or use **Run batch now**.
- Manual Analyze and Execute permissions follow the existing reconciliation-action authorization unless a narrower automatic-execution permission is introduced during implementation; they must never be broader.
- Only the protected scheduler identity may create scheduled executions.

## Error Handling

- Analysis returns safe per-record or per-proposal skip reasons rather than failing the entire run for expected ambiguity or stale data.
- Unexpected rule-level errors mark that rule failed and allow later rules to continue when safe.
- Transaction failures roll back only the affected proposal.
- Scheduled and manual execution records retain diagnostic metadata inaccessible to unauthorized users.
- The interface never reports a proposal completed until the database transaction confirms completion.

## Testing and Verification

Automated tests cover:

- Every identity branch and normalization boundary.
- Similarity-threshold pass/fail boundaries and blank values.
- Eligibility floor and inclusive date-window boundaries.
- Integer-cent arithmetic, positive and negative operators, and inclusive difference tolerance.
- One-to-one, one-to-many, and explicitly defined multi-source combinations.
- Stable oldest-first ordering and combination limits.
- Ambiguous-combination rejection without score tie-breaking.
- Administrator priority conflicts.
- Manual analysis without locks or source mutations.
- Deselection, select-all, stale proposals, and revalidation.
- Scheduled once-per-Lisbon-day claiming, retry idempotency, and overlap prevention.
- Per-proposal rollback with batch continuation.
- Generated comments and structured provenance.
- Settings, action, and scheduler authorization failures.
- Database migrations, constraints, RPC contracts, and repeated migration application.

Authenticated browser verification covers Settings configuration, priority reordering, manual analysis, proposal evidence, selection and execution, result summaries, provenance badges, history, and audit details. A non-production scheduled run verifies the daily claim and generated batch report before production enablement.

## Deployment and Rollout

1. Apply schema and function migrations with all automatic rules disabled.
2. Publish the application and scheduler entry point.
3. Verify Settings and manual analysis in an authenticated non-production environment.
4. Validate the first managed rule against representative known matches and known non-matches, including threshold boundaries and ambiguities.
5. Enable manual execution for the first rule and review real proposals.
6. Enable scheduled execution only after the manual results are accepted.
7. Monitor the batch report and audit trail; disabling the schedule or a rule stops future work without changing completed reconciliations.
