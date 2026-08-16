# Financial Documents to CGD Credit Card Automatic Reconciliation Rule Design

**Date:** 2026-08-16

## Purpose

Add a second managed automatic reconciliation rule that matches eligible Visa Financial Documents to one or more CGD Credit Card records. At the same time, generalize the automatic reconciliation orchestration so additional managed rules can be added without duplicating the paging, proposal, execution, scheduling, audit, or user-interface lifecycle.

The new rule is independent from the existing **Financial Documents to CGD Bank Statement** rule. Manual users select and finish one rule at a time. The shared daily scheduler processes configured rules sequentially in administrator-defined priority order, with a separate auditable run for every rule.

## Scope

This change covers:

- a new managed rule, **Financial Documents to CGD Credit Card**;
- exact base eligibility `financial_documents.payment = 'Visa'`;
- matching one Financial Document to a combination of one through four `import_cgd_cartao_credito` records;
- the new rule's immutable identity logic and similarity thresholds;
- a reusable, allowlisted managed-rule adapter architecture;
- an Automatic reconciliation rule selector for manual analysis;
- independent manual-analysis and scheduled-batch enablement per rule;
- editable scheduled-rule priority;
- deterministic, sequential multi-rule daily batches;
- independent run, proposal, execution, failure, and audit records per rule;
- database, API, frontend, scheduling, security, and regression tests.

This change does not provide a generic SQL-expression builder, allow administrators to edit managed matching logic, combine different rule definitions in one run, queue several rules from the manual interface, or alter the Manual reconciliation workflow.

## Managed Rule Definition

Add managed rule key `financial_documents_cgd_credit_card`, version `1`, with the display name **Financial Documents to CGD Credit Card**.

Its immutable definition is:

- base source: `financial_documents`;
- destination source: `import_cgd_cartao_credito`;
- directional amount operator: `+`;
- base Payment predicate: exact stored value `Visa`;
- default difference allowed: `0.00 EUR`;
- default maximum date difference: `10` calendar days;
- maximum destination records in a combination: `4`;
- maximum identity-qualified destination candidates per base: `12`;
- description similarity threshold: `0.55`;
- supplier-name word-similarity threshold: `0.60`;
- document-number match: normalized compact containment with a minimum of four characters.

The exact Payment comparison is case-sensitive and whitespace-sensitive. `Visa` qualifies; values such as `VISA`, `visa`, ` Visa `, blank, or null do not.

The new rule is seeded disabled by default. Administrators explicitly choose whether it is enabled for manual analysis, enabled for scheduled execution, or both. The existing Bank Statement rule retains its current configuration and behavior.

The managed definition and explanation are read-only. Administrators may edit only:

- enabled state;
- manual-analysis availability;
- scheduled-batch availability;
- difference allowed;
- maximum difference in days, within the existing `0` through `90` limit;
- scheduled priority.

Changing those values updates the rule's editable configuration and timestamp. Every subsequent run snapshots that configuration for audit and execution; the change never rewrites an active or historical run snapshot and does not create an editable managed-definition version.

## Reusable Rule-Adapter Architecture

The automation engine remains responsible for shared lifecycle behavior:

- authenticated run creation and restoration;
- deterministic base paging and progress;
- proposal persistence;
- candidate and combination limits;
- overlap and ambiguity handling;
- proposal selection;
- atomic execution and source locking;
- stale detection;
- sanitized errors;
- audit evidence;
- manual and scheduled orchestration;
- result mapping and rendering.

An explicit server-side adapter registry supplies the parts that differ by managed rule:

- supported rule key and version;
- base eligibility predicate;
- destination source and indexed search projection;
- source snapshot mapping;
- date and amount fields;
- identity scoring and evidence;
- amount operator;
- threshold and combination limits.

Adapter keys are allowlisted in code and SQL. Configuration data cannot name a table, function, SQL fragment, or adapter implementation. An unknown rule key or unsupported version fails closed with a sanitized validation code.

Every analysis run contains exactly one immutable rule-and-configuration snapshot. The engine never mixes Bank Statement and Credit Card candidates or proposals inside the same run. Adding a future managed rule requires a definition, an adapter and search projection where needed, and focused contract tests; it does not require another copy of the run lifecycle.

## Eligibility and Candidate Search

### Base records

A Financial Document is eligible only when all of the following are true:

- `fat = 'S'`;
- `payment = 'Visa'` exactly;
- its reconciliation date is on or after `2026-01-01`;
- it is not locked by another reconciliation;
- it falls after the current analysis cursor and within the rule's configured analysis scope.

### Destination records

A CGD Credit Card record is eligible only when:

- `import_cgd_cartao_credito.data` is not null and is on or after `2026-01-01`;
- `valor` is not null;
- it is not locked by another reconciliation;
- its `data` is within the configured inclusive date window around the Financial Document date.

The rule uses `data`, not `data_valor`, as the destination reconciliation date.

### Identity signals

Every destination record admitted to a candidate combination must independently satisfy at least one of these signals:

1. Its normalized compact `descricao` contains the Financial Document invoice number, or the normalized invoice number contains the compact description, with at least four normalized characters.
2. Normalized descriptions have PostgreSQL trigram similarity greater than or equal to `0.55`.
3. The destination description and normalized Financial Document supplier name have word similarity greater than or equal to `0.60`.

Evidence stores the signal or signals that qualified each candidate, the measured scores, the thresholds, and the immutable source snapshots. Index-assisted prefilters must be supersets of these exact final checks and must not change their results.

If more than 12 identity-qualified destination records exist for one base record, the base is classified as `candidate_limit`; the engine does not enumerate combinations for it.

## Combination and Difference Rules

For each eligible Financial Document, build combinations containing one through four distinct qualified CGD Credit Card records. The calculation is:

`financial_documents.amount + sum(import_cgd_cartao_credito.valor)`

A combination qualifies when the absolute calculated difference is within the configured difference allowance. With the default `0.00 EUR`, the result must be exactly zero using the existing integer-cent comparison rules.

Outcomes are deterministic:

- exactly one valid combination creates a `proposed` proposal;
- more than one valid combination creates an `ambiguous` proposal and none is automatically executable;
- no valid combination is retained in run totals and audit summaries but does not appear as an individual review row;
- candidate-limit and cross-base destination overlap use the engine's existing safe ambiguity behavior.

The same destination record cannot be consumed by two completed reconciliations. Execution locks and revalidates every source record atomically before creating the reconciliation.

## Execution and Historical Determinism

Execution revalidates:

- the managed rule key and version;
- the immutable run configuration snapshot;
- the directional `+` operator;
- exact `Visa` eligibility;
- source dates and the configured date window;
- source IDs, amounts, descriptions, supplier, and document-number snapshots;
- identity scores and thresholds;
- the one-to-four destination limit;
- the calculated difference and tolerance;
- all reconciliation locks.

If any source value, eligibility condition, rule version, configuration, operator, or lock has changed such that the proposal can no longer be reproduced, the proposal becomes `stale` and no reconciliation is created. Existing completed runs and reconciliations retain their original snapshots and audit evidence.

## Manual Automatic-Reconciliation Experience

The Automatic reconciliation tab shows one LOV/dropdown populated from configured rules that are enabled for manual analysis. It replaces the current set of per-rule Analyze cards.

The workflow is:

1. The user selects one enabled rule.
2. The user clicks **Analyze**.
3. The selected rule creates or resumes one paged, auditable run.
4. While that run is unfinished, the rule selector is locked so the user cannot start a second rule.
5. The user reviews proposed and ambiguous matches, may select or deselect executable proposals, and clicks **Execute selected**.
6. After execution finishes, or after analysis finishes with no executable proposals, the selector becomes available for another rule.

There is no manual multi-rule queue and no automatic movement to another rule. The user deliberately chooses the next rule after finishing the previous one.

The review list retains the established visibility rules:

- active runs show proposed rows whether checked or unchecked, plus ambiguous rows;
- no-match/skipped details are hidden but remain included in summary counts and audit data;
- completed runs show only selected outcomes persisted as `completed`, `stale`, or `failed`;
- deselected proposals remain selectable until execution, then remain in audit data without an individual completed-run row.

Reconciliation history remains shared between the Manual and Automatic tabs. Automatic history retains manual-versus-scheduled origin, the managed rule identity and version, source summaries, status, and difference.

## Settings Experience

Settings -> Reconciliation continues to contain the managed automatic-rule configuration and shared schedule.

Each rule displays its read-only managed logic and editable operational settings. The new Credit Card rule explains its exact Visa predicate, source pair, identity alternatives, thresholds, maximum four-card combination, default tolerance, and default 10-day window.

Scheduled priority is editable. The default order is:

1. Financial Documents to CGD Bank Statement.
2. Financial Documents to CGD Credit Card.

The Settings action **Run batch now** is removed. It is replaced with **Open automatic reconciliation**, which navigates to the Automatic reconciliation tab without starting analysis or preselecting execution. Rule execution remains in the Reconciliation application, while configuration and the shared schedule remain in Settings.

## Daily Batch Scheduling

The system keeps one shared schedule, running once per day at the administrator-configured time.

To make multiple rules deterministic, add a daily automatic-batch record. At the beginning of a scheduled slot in the configured schedule timezone, the batch snapshots the enabled scheduled rules, their configuration, and their priority order. Later Settings changes apply to the next daily batch, not the batch already in progress. Priority order is ascending; equal priorities are resolved by stable managed rule key so the batch order is always deterministic.

Each batch has one child automatic run per rule. The uniqueness boundary is the batch and managed rule identity, so rerunning a heartbeat cannot create a duplicate child run. Each child still contains exactly one immutable rule snapshot.

The scheduler processes rules sequentially:

1. Resume the oldest unfinished child run, if one exists.
2. Otherwise claim the next unstarted rule in the batch priority snapshot.
3. Analyze that rule in deterministic pages.
4. For a scheduled run, automatically execute only unique executable proposals after analysis completes.
5. Mark the child run terminal and advance to the next rule on a later heartbeat.
6. Return a safe no-work result after every batch rule is terminal.

A failed rule records a sanitized terminal failure but does not block later rules. Retries resume the same unfinished run and cannot duplicate proposals or reconciliations. One slow rule never causes two rules to be processed concurrently inside the same batch.

## Data Model

Introduce an automatic batch entity with, at minimum:

- batch ID;
- scheduled slot/date and trigger identity;
- immutable ordered rule snapshot;
- lifecycle status and timestamps;
- aggregate safe outcome counts.

Automatic runs gain a nullable batch reference and a stable rule position/key within that batch. A uniqueness constraint prevents more than one child run for the same batch and rule snapshot. Manual runs have no scheduled batch parent.

The existing proposal, execution, source-lock, reconciliation-origin, and audit structures remain authoritative. Schema changes are forward-only and idempotent. Migration logic preserves completed historical data and safely handles any unfinished pre-migration scheduled run according to the existing upgrade/restart policy.

## API and RPC Boundaries

All database access remains behind authenticated application endpoints and service-role RPCs. The browser never reads automation tables directly.

The API adds or generalizes only the contracts needed to:

- list enabled manual rules for the selector;
- create or resume an analysis for one selected rule;
- continue its pages;
- return the actor's active manual run;
- claim or resume the next scheduled batch child;
- expose sanitized batch and run progress.

The server binds manual runs to the authenticated actor. Administrator-only Settings operations remain separate from app-authorized analysis and review operations. Internal helpers and projections revoke access from `public`, `anon`, and `authenticated`; only the necessary service role and security-definer functions receive privileges. All security-definer functions retain a fixed safe `search_path`.

## Error Handling

- Unsupported rules or versions fail closed.
- Public errors use stable sanitized messages and codes; database exception details remain server-side.
- Page processing and proposal persistence are atomic. A failed page advances neither cursor nor proposals.
- Network uncertainty causes the client or scheduler to reload authoritative run state before retrying.
- Analysis failures persist a safe terminal marker without leaving a permanently claimable run.
- A scheduled child failure advances the batch to the next rule.
- A manual failure leaves the rule available for a fresh deliberate analysis after the failed run is acknowledged.
- Partial analyses cannot execute proposals.
- Configuration changes never rewrite active or completed snapshots.

## Verification

Strict RED/GREEN TDD covers every production change.

### Rule behavior

Database and source-contract tests prove:

- only exact `Visa` Payment values qualify;
- `VISA`, `visa`, padded values, blank, and null are completely excluded before proposal creation;
- source dates before `2026-01-01`, null card dates/amounts, and locked records are excluded;
- `data` is used and `data_valor` is not used for the date window;
- a date difference of 10 days is included and 11 days is excluded under defaults;
- configured windows continue to support the engine's `0` through `90` range;
- document-number matching requires at least four normalized characters;
- description scores immediately below `0.55` fail and scores at or above it pass;
- supplier-name word scores immediately below `0.60` fail and scores at or above it pass;
- each card in a combination independently satisfies at least one identity signal;
- combinations of one, two, three, and four cards can qualify;
- a five-card-only solution does not qualify;
- the `+` calculation uses integer cents and exact default tolerance;
- multiple valid combinations, cross-base overlap, and more than 12 candidates are safely ambiguous;
- execution revalidation, locks, stale handling, and historical snapshots remain correct.

### Multi-rule orchestration

Tests prove:

- every run contains exactly one rule snapshot;
- the default daily order is Bank Statement then Credit Card;
- administrators can change future batch priority;
- a daily batch snapshots configuration and order once;
- each batch creates at most one child run per rule;
- a child finishes before the next starts;
- failure of the first rule does not prevent the second;
- retries and cross-midnight heartbeats resume the authoritative batch without duplicates;
- manual and scheduled enablement are independent;
- scheduled runs automatically execute only unique proposals;
- manual runs never automatically start another rule.

### User interface and security

Executable frontend and API tests prove:

- the LOV lists only enabled manual rules and escapes labels;
- Analyze submits exactly the selected allowlisted rule key;
- the selector is locked during an unfinished run and restored afterward;
- active and completed proposal visibility remains correct;
- Settings shows immutable logic and only permitted editable fields;
- scheduled priority saves and reloads authoritatively;
- **Open automatic reconciliation** navigates without starting a run;
- shared history appears in both reconciliation tabs with rule identity and origin;
- app-only and administrator-only authorization boundaries remain intact;
- browser payloads cannot inject source names, SQL, functions, operators, or thresholds.

Run syntax checks, focused Node tests, the complete Node regression suite, migration source-contract checks, diff checks, and transactional PostgreSQL smoke tests. The SQL smoke must cover migration reapplication, projection synchronization, RPC privileges, RLS, threshold boundaries, combination limits, multi-rule batch claiming, failure continuation, execution, locking, and audit history.

## Rollout

1. Apply the new forward-only Supabase migration in the normal migration folder after all existing automatic-reconciliation migrations.
2. Run the transactional automation SQL smoke in a disposable or development database.
3. Publish the API and frontend changes.
4. In Settings, verify the new rule is present but disabled and that the Bank Statement rule remains unchanged.
5. Enable the new rule for manual analysis, keep scheduled execution disabled, and run an authenticated Visa analysis against known fixtures.
6. Verify one-to-four-card proposals, ambiguity, audit evidence, execution, history, and stale handling.
7. Enable scheduled execution and confirm a test batch processes Bank Statement first and Credit Card second as separate runs.
8. Confirm a deliberately failed first child does not block the second before relying on the production daily schedule.

## Acceptance Criteria

The feature is accepted when an authorized user can select the enabled Credit Card rule from the Automatic reconciliation LOV, analyze only exact-Visa Financial Documents, review auditable one-to-four-card proposals using the approved thresholds and calculation, and execute selected proposals safely. Administrators can independently enable manual and scheduled use and change daily rule priority. The shared scheduler processes each enabled rule sequentially as a separate immutable run, continues after an individual rule failure, and supports adding future managed adapters without duplicating the lifecycle engine.
