# FDM Bank and Adyen Automatic Reconciliation Design

**Date:** 2026-08-23

**Status:** Approved design

**Scope:** Add two disabled-by-default managed automatic reconciliation rules to the existing manual-review and shared scheduled-batch system.

## 1. Purpose

Add these automatic reconciliation rules without creating a separate rule engine or bypassing the existing proposal, review, execution, locking, audit, and scheduling lifecycle:

1. **FDM Accounts – Bank Reservation Payments**
2. **FDM Accounts – Adyen Reservation Payments**

The implementation extends the controlled managed-strategy registry. Rule logic remains immutable and read-only in Settings. Administrators control activation, manual availability, scheduled-batch participation, priority, and the specifically approved editable thresholds.

## 2. Existing Architecture to Reuse

The feature must reuse the current automatic reconciliation system:

- Immutable versioned rule definitions.
- Administrator-owned rule configurations.
- One manual rule selection and one unfinished manual run per actor.
- Sequential child runs inside one shared scheduled batch.
- Immutable run and proposal snapshots.
- Proposal memberships for grouped source records.
- Stable proposal statuses: proposed, ambiguous, stale, failed, completed, and skipped/deselected.
- User review before manual execution.
- Empty selection as an explicit way to finish a reviewed run.
- Transactional execution, locking, stale detection, audit trail, and history origin.
- Service-role-only mutation RPCs behind authorized API handlers.

The new work must not introduce editable SQL expressions, dynamic table names, or a general user-authored rule language.

## 3. Managed Rule Catalog

### 3.1 Rule A: FDM Accounts – Bank Reservation Payments

| Property | Contract |
| --- | --- |
| Proposed rule key | `fdm_bank_transfer_cgd_bank_statement_combination` |
| Version | `1` |
| Display name | `FDM Accounts – Bank Reservation Payments` |
| Business direction | FDM Accounts → CGD Bank Statement |
| FDM source type | `import_fdm_accounts` |
| Bank source type | `import_cgd_extrato_ordem` |
| FDM eligibility | `Account = 'Bank Transfer'` exactly |
| Amount equation | Bank amount plus the selected FDM signed amounts must equal exactly `0.00` |
| Sign requirement | Bank and selected FDM total use opposite signs |
| Date rule | Every selected FDM record is within the configured inclusive day distance from the bank record |
| Default maximum difference in days | `3` |
| Editable days | Yes, from 0 through 90 inclusive |
| Difference allowed | Fixed `0.00` and read-only |
| Group cardinality | Exactly one bank record and between 1 and 10 FDM records |
| Directional operator | Snapshot and use the configured FDM Accounts → CGD Bank Statement source rule |
| Default activation | Disabled |

The candidate equation is evaluated in exact integer cents. Floating-point arithmetic is not allowed.

The configured directional operator remains authoritative for reconciliation calculation. A candidate group is executable only when both the opposite-sign amount contract and the operator-derived reconciliation difference are exactly zero. A later operator change makes an existing proposal stale.

### 3.2 Rule B: FDM Accounts – Adyen Reservation Payments

| Property | Contract |
| --- | --- |
| Proposed rule key | `cgd_bank_statement_fdm_adyen_monthly_payments` |
| Version | `1` |
| Display name | `FDM Accounts – Adyen Reservation Payments` |
| Business direction | CGD Bank Statement → FDM Accounts |
| Bank source type | `import_cgd_extrato_ordem` |
| FDM source type | `import_fdm_accounts` |
| Bank eligibility | Description contains `Adyen`, case-insensitively |
| FDM eligibility | `Account = 'Adyen'` exactly |
| Grouping | One proposal per closed calendar month |
| Default difference allowed | `2000.00` |
| Editable difference allowed | Yes; non-negative and within existing numeric limits |
| Maximum difference in days | Fixed/read-only `31`; it identifies calendar-month mode and is not a rolling date window |
| Directional operator | Snapshot and use the configured CGD Bank Statement → FDM Accounts source rule |
| Default activation | Disabled |

Each monthly proposal contains every eligible, unlocked Bank and FDM record in that calendar month. Both sides must contain at least one record. The current incomplete month is excluded.

The difference is calculated from signed decimal amounts using the configured directional operator. A proposal is executable when the absolute difference is less than or equal to the configured allowance.

- Difference `0.00`: complete normally.
- Non-zero difference within the allowance: force-complete with a deterministic generated audit comment identifying the rule, month, actual difference, and configured allowance.
- Difference above the allowance: ambiguous with reason `monthly_difference_exceeded`.

## 4. Configuration and Activation

Both configurations are inserted only when absent with `enabled = false`, `allow_manual_execution = false`, and `include_in_scheduled_batch = false`. Deployment must not activate either rule.

Settings → Reconciliation continues to expose the existing controls:

- Enabled.
- Allow manual execution.
- Include in the shared scheduled batch.
- Priority/order.

Administrators independently choose manual availability and scheduled-batch participation after enabling a rule.

Rule-specific fields:

### Bank Reservation Payments

- Difference allowed: `0.00`, fixed and read-only.
- Maximum difference in days: editable, default `3`.
- Maximum FDM records per bank record: `10`, fixed and read-only.
- Matching logic: read-only explanation.

### Adyen Reservation Payments

- Difference allowed: editable, default `2000.00`.
- Calendar-month value: `31`, fixed and read-only.
- Matching logic: read-only explanation.

The settings replacement RPC must validate the complete seven-rule managed catalog, reject duplicate keys and unsupported versions, preserve rule-definition immutability, and remain atomic.

## 5. Bank Combination Analysis

### 5.1 Stable analysis anchor

Analysis pages through eligible unlocked CGD Bank Statement records in stable `(date, id)` order. Each bank record is the stable search anchor. The proposal retains the business direction FDM Accounts → CGD Bank Statement and records all FDM memberships as source-role members and the bank membership as the destination-role member.

Where the existing proposal schema requires one `base_source_id`, the canonical base is the first selected FDM member in stable `(date, id)` order. The complete group is authoritative in proposal memberships and the summary snapshot; execution must never infer group membership from only the canonical base row.

### 5.2 Candidate pool

For each bank record, eligible FDM candidates must:

- Have Account exactly `Bank Transfer`.
- Satisfy the global reconciliation eligibility floor of 2026-01-01.
- Be unlocked by manual and automatic reconciliation ownership rules.
- Be within the configured inclusive absolute day difference from the bank date.
- Have a non-null valid amount.
- Have a sign compatible with the approved opposite-sign contract.
- Have an absolute amount no larger than the remaining target during combination search.

Candidates use stable `(event_date, id)` order.

### 5.3 Bounded exact combination search

The search considers combinations of 1 through 10 distinct FDM records. It uses exact cents and deterministic ordering.

The implementation must be bounded:

- Maximum eligible FDM candidate pool per bank: 60 records. A larger pool is classified as `candidate_limit` rather than truncated into a false answer.
- Maximum evaluated combination-search states per bank: 250,000. Reaching the limit is classified as `candidate_limit`.
- Maximum persisted qualifying candidate groups per bank: 12 evidence groups.
- The thirteenth qualifying group changes the proposal reason to `candidate_limit`.
- Reaching a safety ceiling must produce an ambiguous `candidate_limit` result. It must never silently report no match.
- Fully exhausting the bounded search with no qualifying group produces no visible proposal and increments only the run's skipped/no-match accounting.

### 5.4 Proposal classification

- Exactly one qualifying combination: `proposed`.
- More than one qualifying combination: `ambiguous`, reason `multiple_qualifying_combinations`.
- More qualifying combinations or search states than the safe bound: `ambiguous`, reason `candidate_limit`.
- A bank or FDM record appearing in more than one otherwise-proposed group across bank anchors: every affected proposal becomes `ambiguous`, reason `overlapping_records`.
- No two-bank group and no more than ten FDM records are allowed.

Ambiguous evidence is immutable and bounded. It includes stable IDs, dates, signed amounts, descriptions/accounts where available, the equation total, and the reason.

## 6. Adyen Monthly Analysis

Analysis pages through closed calendar months in stable ascending month order, beginning at the 2026 eligibility floor.

For each month:

1. Load every unlocked Bank record whose description contains `Adyen` case-insensitively.
2. Load every unlocked FDM record whose Account equals `Adyen` exactly.
3. Require at least one record from each source.
4. Persist immutable memberships and source totals.
5. Calculate the directional monthly difference using exact numeric arithmetic.
6. Classify against the snapshotted allowance.

The current month is never analyzed until it closes. The value `31` is a managed calendar-month marker and is not used to create a rolling 31-day window or to reject a valid month whose records span more than 31 elapsed days across boundaries; cross-month records are never combined.

Months with one empty side or no eligible records do not appear in proposal review.

## 7. Immutable Snapshots

Every run snapshots the selected definition and configuration, including:

- Rule key and version.
- Display name.
- Strategy/matching mode.
- Source and destination source types.
- Directional operator and source-rule identity/version.
- Difference allowance.
- Date mode and configured days.
- Maximum combination size.
- Candidate/evidence limits.
- Manual or scheduled trigger.
- Batch ID, position, and rule count when scheduled.

Every proposal snapshots:

- Grouping key or calendar month.
- Classification and reason.
- Signed source totals, destination totals, difference, and allowance.
- Every member's source type, ID, role, ordinal, date, amount, description, Account, and the minimal complete source-row audit snapshot.
- Candidate groups/evidence for ambiguous results.

Snapshots are append-only audit evidence. Reapplying the migration or editing live source rows must not rewrite them.

## 8. Execution and Stale Detection

Execution uses specialized allowlisted strategy dispatch. It must not use dynamic SQL.

Before any reconciliation write, execution locks the proposal, run, rule/config/source-rule rows, and every live source member in a deterministic global order. It then revalidates:

- Run ownership, trigger, lifecycle, and selected rule.
- Proposal status and selection.
- Rule definition/version and managed strategy.
- Configuration and source-rule operator identity.
- Exact membership count and membership identities.
- Every member's source type, ID, date, signed amount, description, and Account.
- Rule-specific eligibility predicates.
- No member is already locked, deleted, or consumed.
- Bank combination equation/cardinality/date constraints.
- Adyen calendar month, totals, and configured allowance.
- No proposal overlap has appeared since analysis.

Any mismatch returns a sanitized stale result and writes no reconciliation or lock.

Successful execution creates one reconciliation containing every snapshotted member. History displays source counts and signed totals through the existing reconciliation history summary.

Unexpected failures roll back the proposal's reconciliation writes and persist only a sanitized failure code/message through the existing safe failure lifecycle.

## 9. Manual Review and UI

The Automatic Reconciliation tab continues to start with no rule running. The rule LOV shows only enabled rules that allow manual execution. A user selects one rule and analyzes it; the selector remains locked until the run is finished.

The existing three-column proposal layout is reused:

- First column: selection, status, reason, difference, allowance, and rule/version.
- Second column: FDM/source members for the Bank combination rule or Bank/source members for Adyen.
- Third column: the single Bank destination for the combination rule or FDM destination members for Adyen.

All members are shown below one another with date, description, Account/supplier-equivalent details where relevant, signed amount, and collapsible record ID. Group totals remain visible.

Only proposed rows are selectable. Ambiguous rows remain visible for review but cannot be executed. The user may deselect and reselect proposals, or finish with Execute Selected (0). Completed runs retain only selected/executed completed proposal details, consistent with the existing behavior.

## 10. Shared Scheduled Batch

The shared schedule remains once per day at the administrator-configured time.

The scheduler snapshots all enabled batch rules in priority then rule-key order. It processes one child rule at a time and must finish or fail the current child before claiming the next child on a later continuation/heartbeat.

The batch supports seven managed rule keys after this change. It preserves:

- One unfinished child per batch.
- Oldest unfinished cross-midnight resume.
- Idempotent same-slot retry.
- Failure continuation to the next rule.
- Accurate aggregate totals and terminal status.
- Strategy-specific progress units: Bank records for the combination rule and calendar months for Adyen.

A disabled rule or a rule not marked for scheduled execution is absent from the batch snapshot.

## 11. Data Model and Migration

No new general-purpose rule-language table is introduced. Reuse:

- `financial_reconciliation_automatic_rule_definitions`
- `financial_reconciliation_automatic_rule_configs`
- `financial_reconciliation_automatic_runs`
- `financial_reconciliation_automatic_batches`
- `financial_reconciliation_automatic_proposals`
- `financial_reconciliation_automatic_proposal_memberships`
- Existing reconciliations, items, audit, source rules, and locking constraints

The new dated migration in `supabase-migrations/` must:

- Be safe to apply and reapply.
- Insert immutable definitions/configurations only when absent.
- Fail closed when a same-key/version definition differs from the approved immutable contract.
- Extend exact allowlists, serializers, dispatchers, settings replacement, manual analysis creation, scheduled claim, analysis continuation, finalization, execution, and public run mapping from five to seven rules.
- Add only justified indexes for FDM Account/date/amount and Bank date/description/amount lookup.
- Compare the exact expected index/constraint definitions on reapply and fail on a conflicting same-named object.
- Preserve historical rows and snapshots unchanged.
- Avoid transaction-control statements so the transactional SQL smoke can include the migration inside its outer rollback.

## 12. Authorization and Security

- Rule Settings GET/PUT remain administrator-only through the current feature authorization.
- Manual analysis/execution remains restricted to authorized Financial Reconciliation users.
- Scheduled execution remains service-role-only through the protected cron route.
- Private SQL helpers revoke execution from public, anonymous, and authenticated roles.
- Public mutation RPCs revoke execution from those roles and grant only service role.
- Every SECURITY DEFINER function uses a fixed safe search path and schema-qualified objects.
- No live SQL error text, table contents, or internal details are returned to the browser.
- No dynamic SQL chooses a source table or rule function.

## 13. API Contracts

The existing endpoints remain authoritative:

- `/api/reconciliation-automation-settings`
- `/api/reconciliation-automation`
- `/api/reconciliation-automation-cron`

Shared response mapping must recognize both new keys and preserve complete run/proposal/member snapshots. Unknown keys, unsupported versions, prototype-backed values, malformed lifecycle states, and invalid rule-specific settings fail closed.

No endpoint may bypass the RPC lifecycle with direct proposal or reconciliation mutations.

## 14. Testing and Verification

### 14.1 Node/API contract tests

- Seven-rule settings GET/PUT round trips.
- Immutable fields and rule-specific editable/fixed fields.
- Manual catalog visibility from enabled/manual flags.
- Manual analysis creation and one-rule locking.
- Empty-selection finalization.
- Seven-child scheduled ordering and one-child-per-heartbeat behavior.
- Retry, failure continuation, cross-midnight resume, and aggregate terminal status.
- Unknown key/version and malformed response rejection.

### 14.2 Transactional PostgreSQL smoke tests

Bank combination fixtures:

- Exact opposite one-to-one match.
- Unique combinations of 2 through 10 FDM records.
- Same-sign and one-cent mismatch exclusion.
- Inclusive date boundary and one-day-outside exclusion.
- Account exact-match exclusion.
- Multiple qualifying combinations.
- Candidate-limit behavior with bounded evidence.
- Shared-bank and shared-FDM overlap ambiguity.
- No two-bank or more-than-ten-FDM execution.
- Source/operator/config/member drift.
- Lock contention, idempotency, atomic rollback, and sanitized failure.

Adyen fixtures:

- Case-insensitive Bank description inclusion.
- Non-Adyen description exclusion.
- Exact FDM Account inclusion and near-match exclusion.
- Closed-month grouping and current-month exclusion.
- Both sides required.
- Zero, within-allowance non-zero, exact allowance boundary, and over-allowance outcomes.
- Normal and forced completion/audit evidence.
- Member/source-rule/config drift and overlap.
- Large monthly membership execution and reapply preservation.

Cross-layer fixtures:

- Apply and reapply the migration.
- Preserve all existing five-rule behavior.
- Seven-rule settings and schedule snapshots.
- Manual and scheduled complete flows for both new rules.
- Service-role ACLs and private-helper denial.

### 14.3 Local and deployment gates

- `node --check` for changed API/client files.
- Focused rule/API/UI tests.
- Full Node test suite.
- `git diff --check`.
- Transactional PostgreSQL smoke with `ON_ERROR_STOP=1` before production activation.
- Authenticated desktop and narrow-layout browser verification when a session is available.
- Protected non-production scheduled heartbeat/retry verification.

## 15. Rollout

1. Deploy the reapply-safe migration.
2. Reapply it once and run the complete transactional SQL smoke.
3. Deploy API/client changes.
4. Verify Settings shows both rules disabled with correct immutable/editable fields.
5. Enable manual execution for one rule in non-production and validate analysis, review, execution, history, and audit.
6. Validate the second rule similarly.
7. Enable scheduled participation only after the protected batch heartbeat and retry checks pass.
8. Administrators choose final priority/order.

No rule is activated by the migration or application deployment.

## 16. Acceptance Criteria

The feature is accepted when:

- Both rules exist as immutable versioned definitions and disabled configurations.
- Administrators can independently enable manual and scheduled execution and reorder them.
- Bank Reservation Payments uniquely matches one Bank record to 1–10 opposite-signed `Bank Transfer` FDM records totaling exactly zero, with ambiguity and overlap handled safely.
- Adyen creates one complete closed-calendar-month proposal from all eligible Bank/FDM records and applies the editable €2,000 default allowance.
- Every execution revalidates immutable snapshots and fails stale without partial writes.
- Manual review, Execute Selected (0), history, origin, and audit behavior remain consistent.
- Scheduled processing handles seven rules sequentially and resumes safely.
- Existing five rules remain unchanged.
- Migration reapply, PostgreSQL smoke, full Node tests, and available browser/deployment gates pass before activation.
