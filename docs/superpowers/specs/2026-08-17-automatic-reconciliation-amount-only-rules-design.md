# Automatic Reconciliation Amount-Only Rules Design

**Date:** 2026-08-17

## Purpose

Add two independent managed automatic reconciliation rules that match one eligible Financial Document to exactly one destination record using only signed amount equality and an inclusive date window. The rules reuse the existing automatic-reconciliation lifecycle, manual rule selector, scheduled batch, proposal review, execution, history, and audit framework.

The new rules are conservative fallbacks after the existing identity-based Bank Statement and Credit Card rules. They are disabled by default and must be explicitly enabled by an administrator.

## Scope

This change covers:

- **Financial Documents to CGD Bank Account – AMOUNT ONLY**;
- **Financial Documents to CGD Credit Card – AMOUNT ONLY**;
- exact `Banco` and `Visa` base eligibility;
- strictly one-to-one destination matching;
- signed integer-cent equality with a fixed zero difference allowance;
- an editable inclusive date window, defaulting to one day;
- safe ambiguity for duplicate and cross-base matches;
- manual analysis through the existing rule selector;
- sequential scheduled execution through the shared daily batch;
- read-only managed logic and fixed tolerance in Settings;
- immutable rule/configuration snapshots, stale detection, audit evidence, and shared reconciliation history;
- database, API, frontend, security, and regression coverage.

This change does not add multi-record destination combinations, use invoice numbers or text similarity, allow a nonzero difference tolerance, create a generic administrator-authored rule language, change Manual Reconciliation, or modify the behavior of the two existing identity-based automatic rules.

## Managed Rule Definitions

### Bank Account amount-only rule

- rule key: `financial_documents_cgd_bank_statement_amount_only`;
- version: `1`;
- display name: **Financial Documents to CGD Bank Account – AMOUNT ONLY**;
- base source: `financial_documents`;
- destination source: `import_cgd_extrato_ordem`;
- directional operator: `+`;
- exact base Payment: `Banco`;
- fixed difference allowed: `0.00 EUR`;
- default maximum date difference: `1` calendar day;
- maximum destination records per proposal: `1`;
- default priority: immediately after the two existing managed rules, normally `3`.

### Credit Card amount-only rule

- rule key: `financial_documents_cgd_credit_card_amount_only`;
- version: `1`;
- display name: **Financial Documents to CGD Credit Card – AMOUNT ONLY**;
- base source: `financial_documents`;
- destination source: `import_cgd_cartao_credito`;
- directional operator: `+`;
- exact base Payment: `Visa`;
- fixed difference allowed: `0.00 EUR`;
- default maximum date difference: `1` calendar day;
- maximum destination records per proposal: `1`;
- default priority: immediately after the Bank Account amount-only rule, normally `4`.

Both rules are seeded disabled for manual and scheduled execution. Their managed definitions and zero difference allowance are immutable. Administrators may edit only:

- enabled state;
- manual-analysis availability;
- scheduled-batch availability;
- maximum difference in days, within the existing `0` through `90` limit;
- scheduled priority.

Settings displays `Difference allowed` as a read-only `0.00 €` value for these two rules. The API and database reject any nonzero submitted value even if a browser payload is tampered with. The existing identity-based rules retain their current editable difference allowance.

## Eligibility

### Financial Documents

A Financial Document is eligible for an amount-only rule only when all of the following are true:

- `fat = 'S'`;
- `payment` is exactly `Banco` for the Bank Account rule or exactly `Visa` for the Credit Card rule;
- `document_date` is not null and is on or after `2026-01-01`;
- `amount` is not null;
- the document is not locked by another non-deleted reconciliation;
- it falls after the run's deterministic analysis cursor and within the run's analysis scope.

Payment matching is case-sensitive and whitespace-sensitive. `Banco` and `Visa` qualify; casing variants, padded values, blank values, and null do not.

### Bank Account destinations

A Bank Account destination is eligible only when:

- `import_cgd_extrato_ordem.data` is not null and is on or after `2026-01-01`;
- `montante` is not null;
- it is not locked by another non-deleted reconciliation;
- `data` is between the Financial Document date minus and plus the configured maximum difference in days, inclusively.

### Credit Card destinations

A Credit Card destination is eligible only when:

- `import_cgd_cartao_credito.data` is not null and is on or after `2026-01-01`;
- `valor` is not null;
- it is not locked by another non-deleted reconciliation;
- `data` is between the Financial Document date minus and plus the configured maximum difference in days, inclusively.

The Credit Card rule uses `data`, not `data_valor`, as its reconciliation date.

Neither destination adapter reads invoice number, description, supplier, supplier NIF, or similarity scores when deciding whether a candidate qualifies.

## Amount and Date Matching

Amounts are converted through the existing integer-cent path. A destination qualifies only when:

`Financial Document amount + destination amount = 0 cents`

For example, `100.00 €` matches `-100.00 €`. The calculation does not use floating-point tolerance or absolute-value equality, and the fixed difference allowance cannot be increased.

The date window is symmetric and inclusive. With the default one-day setting, a destination dated one day before, the same day, or one day after the Financial Document qualifies; a destination two days away does not. Administrators may change the window from `0` through `90` days for future runs. Every run snapshots the configured window.

The adapters use indexed equality/range candidate lookups on destination amount, date, and stable ID. The migration adds or verifies supporting indexes without changing source data.

## Proposal and Ambiguity Rules

For each eligible Financial Document:

- zero exact destination matches: record skipped/no-match totals and audit evidence, but do not display an individual review row;
- exactly one exact destination match: create one selectable `proposed` proposal;
- more than one exact destination match: create an `ambiguous` proposal and expose its candidate evidence; none is executable;
- more than the engine's safe candidate-evidence limit: classify the result as `candidate_limit`, retain safe audit evidence, and do not choose a destination.

The engine performs cross-base overlap detection after candidate analysis. If the same destination is the sole qualifying match for more than one Financial Document, every affected proposal becomes ambiguous. The engine never resolves duplicates using nearest date, insertion order, record ID, supplier, description, or invoice number.

No proposal contains more than one destination record, and the engine never enumerates or accepts combinations of two or more destinations for these rules.

## Shared Adapter Architecture

The existing automatic-reconciliation engine remains responsible for:

- authenticated run creation and restoration;
- deterministic base paging and progress;
- proposal persistence and visibility;
- overlap and ambiguity handling;
- manual proposal selection;
- atomic source locking and execution;
- stale detection;
- immutable source and configuration snapshots;
- audit evidence and reconciliation origin;
- manual and scheduled orchestration;
- sanitized errors and public result mapping.

Two explicit allowlisted adapters supply only the rule-specific contract:

- supported key and version;
- exact Payment eligibility;
- destination source, date field, and amount field;
- one-to-one candidate lookup;
- fixed `+` operator;
- fixed zero tolerance;
- maximum one destination record;
- evidence explaining exact signed amount equality and date distance.

Configuration cannot name tables, columns, functions, SQL fragments, operators, or adapter implementations. Unknown keys or unsupported versions fail closed. Adding future managed rules follows the same adapter registry rather than copying the run lifecycle.

## Execution and Historical Determinism

Execution atomically locks and revalidates both source records. It verifies:

- managed rule key and version;
- immutable run configuration snapshot;
- exact `Banco` or `Visa` Payment eligibility;
- fixed `+` source operator;
- fixed zero difference allowance;
- configured inclusive date window;
- one-to-one destination count;
- source IDs, dates, signed amounts, and eligibility snapshots;
- exact zero-cent calculation;
- source and reconciliation locks.

If any value, rule contract, configuration, eligibility condition, or lock changed such that the proposal cannot be reproduced, the proposal becomes `stale` and no reconciliation is created. Existing completed runs and reconciliations retain their original snapshots and evidence.

The same destination cannot be consumed by two completed reconciliations. Concurrent or repeated execution remains idempotent through the current locking and uniqueness boundaries.

## Manual Experience

The Automatic Reconciliation rule selector lists the two new rules only when they are enabled for manual analysis. The workflow remains:

1. The user selects one enabled rule.
2. The user clicks **Analyze**.
3. The selected rule creates or resumes one paged, auditable run.
4. The selector remains locked while the run is unfinished.
5. The user reviews proposed and ambiguous matches in the existing three-column layout.
6. The user may select or deselect executable proposals and clicks **Execute selected**.
7. After the run is terminal, the user deliberately selects another rule.

There is no manual multi-rule queue and no automatic transition to the next rule. Active and completed proposal visibility retains the current behavior, and no-match details remain in totals/audit data without creating review rows.

## Settings Experience

Settings -> Reconciliation -> Automatic shows all four managed rules. The new rules display their one-to-one amount/date logic read-only.

For each amount-only rule:

- `Difference allowed` is rendered read-only as `0.00 €`;
- enabled/manual/scheduled controls remain editable;
- maximum difference in days remains editable from `0` through `90`;
- priority remains editable;
- the source pair, exact Payment, fixed operator, one-destination maximum, and absence of identity signals are read-only.

Atomic Save continues to replace the complete authoritative Settings payload. API and database validation require every currently installed managed rule exactly once, reject duplicate priorities, and enforce the fixed zero tolerance for amount-only definitions.

## Scheduled Execution

The shared daily schedule remains unchanged. On an installation whose existing rules retain their default order, new scheduled batches default to:

1. Financial Documents to CGD Bank Statement;
2. Financial Documents to CGD Credit Card;
3. Financial Documents to CGD Bank Account – AMOUNT ONLY;
4. Financial Documents to CGD Credit Card – AMOUNT ONLY.

Administrators may reorder future batches. At claim time, a batch snapshots enabled scheduled rules, priorities, versions, and editable configuration. Each rule receives a separate child run and is processed sequentially. A child finishes before a later rule starts, and a failed child records a sanitized terminal result without blocking later rules.

If administrators already reordered the two existing identity rules before this migration, the migration preserves that relative order and appends the two new amount-only rules after them. It never resets an existing administrator choice merely to force numeric priorities `1` and `2`; it assigns the new rules the next two available priorities deterministically.

Running the stronger identity-based rules first lets them lock high-confidence matches before the amount-only fallbacks. Manual runs remain independent and never start another rule automatically.

## Data and Migration

One forward-only, re-runnable migration in the normal Supabase migration folder will:

- insert immutable version-1 definitions for both amount-only rules;
- seed disabled configurations at the next two deterministic priorities after the existing rules, producing priorities 3 and 4 on the standard configuration without changing existing rule settings or relative order;
- extend managed rule/version allowlists from two to four entries;
- extend generic analysis, continuation, execution, Settings, and scheduled-batch contracts to recognize both new adapters;
- enforce fixed zero tolerance for the two amount-only rules;
- preserve fixed `+` directional source rules;
- add or verify destination date/amount indexes;
- retain existing tables, runs, proposals, history, and identity-rule behavior;
- reapply RLS, fixed `search_path`, and service-role-only helper/RPC privileges;
- notify PostgREST to reload its schema cache.

The rollout uses compatibility-tolerant application code first: before the migration, it accepts the existing two-rule catalog; after the migration, it accepts and requires the complete four-rule catalog returned by the database. This avoids a Settings outage between application deployment and manual migration execution. After the migration is applied, atomic Settings replacement requires all four rules exactly once.

## API and Security Boundaries

The browser continues to use authenticated application endpoints; it never reads automation tables directly.

Application validation will:

- allow only the four explicit managed rule/version pairs;
- keep manual analysis restricted to exactly one selected rule;
- reject nonzero amount-only difference allowance;
- reject duplicate or missing installed-rule configurations after migration;
- map only approved public fields and strip diagnostics;
- bind manual runs to the authenticated actor;
- keep Settings operations administrator-only and analysis operations app-authorized.

Database helpers, projections, and mutation RPCs remain unavailable to `public`, `anon`, and `authenticated`. Necessary execution remains limited to `service_role` through security-definer functions with fixed safe search paths.

## Error Handling

- Unknown or unsupported rule contracts fail closed.
- Invalid Payment, operator, tolerance, window, destination count, or snapshot state cannot create proposals or reconciliations.
- Page processing is atomic; a failed page advances neither cursor nor proposals.
- Network uncertainty reloads authoritative state before retrying.
- Partial analyses cannot execute proposals.
- Manual failures become terminal and allow a later deliberate fresh analysis.
- Scheduled child failures do not block later rules.
- Public errors remain sanitized while database diagnostics stay server-side.
- Migration reapplication is a verified no-op; mixed or unrecognized installed definitions raise instead of being partially rewritten.

## Verification

Strict RED/GREEN TDD covers every production change.

### Rule behavior

Tests prove:

- only exact `Banco` or `Visa` Payment values qualify;
- casing variants, padded values, blank values, and null are excluded;
- `fat <> 'S'`, pre-2026 dates, null dates/amounts, and locked records are excluded;
- Bank Account uses `data`/`montante` and Credit Card uses `data`/`valor`;
- day minus one, same day, and day plus one qualify under defaults; day two does not;
- configured windows support the existing `0` through `90` range;
- `100.00 € + -100.00 €` qualifies and a one-cent mismatch does not;
- descriptions, supplier values, and invoice numbers neither admit nor reject candidates;
- zero, one, multiple, candidate-limit, and cross-base overlap outcomes are safe and deterministic;
- every proposal contains at most one destination;
- a match requiring two destinations never qualifies;
- execution revalidates exact eligibility, amounts, dates, fixed tolerance, locks, and snapshots;
- stale or already-locked proposals cannot create reconciliations.

### Multi-rule behavior

Tests prove:

- all four rule/version pairs are allowlisted;
- every run contains exactly one rule snapshot;
- new rules are disabled at priorities 3 and 4 by default;
- amount-only difference allowance is read-only and rejected when nonzero;
- administrators can change future date windows and priorities;
- manual analysis sends exactly the selected amount-only rule;
- a scheduled batch snapshots all enabled rules in configured order;
- each child finishes before the next starts;
- a failed rule does not block later rules;
- retries do not duplicate child runs, proposals, or reconciliations;
- existing Bank Statement and Credit Card identity-rule behavior remains unchanged.

### UI, API, migration, and security

Tests prove:

- the selector lists enabled amount-only rules with escaped labels;
- the selector stays locked during an unfinished run;
- proposal layout and history behavior remain unchanged;
- Settings renders zero difference as read-only only for amount-only rules;
- Settings saves the complete authoritative catalog and preserves priorities;
- pre-migration two-rule and post-migration four-rule rollout states are handled safely;
- browser payloads cannot inject rule keys, versions, sources, operators, SQL, functions, or tolerance;
- migration reapplication, RLS, RPC privileges, fixed search paths, and schema reload are correct.

Run focused Node tests, the complete Node suite, syntax checks, diff checks, and transactional PostgreSQL smoke tests. SQL smoke coverage must include both migration applications, exact eligibility, one-to-one matching, ambiguity, overlap, execution, stale handling, scheduling order, failure continuation, locks, and audit history.

## Rollout

1. Publish compatibility-tolerant API and frontend code; before migration it continues to show the existing two-rule catalog.
2. Apply the new forward-only Supabase migration after all existing reconciliation migrations.
3. Reapply the migration once in a disposable/development database and run the transactional automation SQL smoke.
4. Confirm Settings now shows four managed rules, with both amount-only rules disabled and zero tolerance read-only.
5. Enable only the Bank Account amount-only rule for manual analysis and validate known unique, duplicate, and cross-base fixtures.
6. Repeat for the Credit Card amount-only rule.
7. Enable scheduled execution in a controlled environment and verify the four-rule priority order and failure continuation.
8. Only then enable the desired amount-only rules in the production daily schedule.

## Acceptance Criteria

The feature is accepted when authorized users can deliberately analyze either enabled amount-only rule and receive only one-to-one proposals whose signed amounts net to exactly zero within the configured inclusive date window. Duplicate candidates and cross-base destination overlap are always ambiguous. No identity signal or destination combination affects matching. Administrators can configure enablement, date window, and priority while the zero tolerance remains immutable. Scheduled execution processes enabled rules sequentially after the existing identity rules, and all runs remain secure, auditable, deterministic, and historically reproducible.
