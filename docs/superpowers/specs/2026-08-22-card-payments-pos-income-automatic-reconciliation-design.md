# Card Payments - POS - Income Automatic Reconciliation Design

## Purpose

Add a fifth managed automatic reconciliation rule that reconciles card-payment
income by closed calendar month. Unlike the existing identity and amount-only
rules, this rule intentionally aggregates hundreds of records on each side into
one proposal and one completed reconciliation.

The rule display name is **Card Payments - POS - Income**.

## Approved Business Contract

The immutable rule definition is:

- Rule key: `cgd_bank_statement_fdm_credit_card_monthly_income`
- Current version: `2` (version `1` snapshots remain readable and immutable)
- Matching mode: `monthly_aggregate`
- Base source: `import_cgd_extrato_ordem` (CGD Bank Statement)
- Base predicate: `descritivo ILIKE '%POS VENDAS%'`
- Destination source: `import_fdm_accounts` (FDM Accounts)
- Destination predicate: `account = 'Credit Card' AND category IS DISTINCT FROM
  'TransferOutToAccount'` (`NULL` category remains eligible)
- Grouping: the same closed calendar month, using Bank Statement `data` and
  FDM Accounts `event_date`
- Operator: the configured directional source rule
  `CGD Bank Statement -> FDM Accounts (-)`
- Difference: `CGD Bank Statement total - FDM Accounts total`
- Initial allowed absolute difference: EUR 7,500.00
- Maximum difference in days: `31`, immutable and read-only because calendar
  month membership is the actual date rule
- Eligibility floor: `2026-01-01`
- Default manual execution: disabled
- Default scheduled execution: disabled

Administrators may edit the difference tolerance, execution flags, and priority.
They cannot edit the rule definition, operator, predicates, grouping mode,
version, or 31-day display property.

## Matching Semantics

Analysis considers only records that are currently unlocked and otherwise
eligible for reconciliation.

For each calendar month strictly earlier than the current calendar month:

1. Collect every unlocked Bank Statement record on or after `2026-01-01` whose
   description contains `POS VENDAS`, case-insensitively.
2. Collect every unlocked FDM Accounts record in the same month whose account is
   exactly `Credit Card` and whose category is not exactly
   `TransferOutToAccount`. A `NULL` category remains eligible.
3. If either collection is empty, persist no visible proposal for that month.
4. Sum each collection with exact decimal arithmetic and calculate Bank total
   minus FDM total.
5. Create one executable `proposed` result when the absolute difference is less
   than or equal to the snapshotted configured tolerance.
6. Create one non-executable `ambiguous` result with reason
   `monthly_difference_exceeded` when the absolute difference is greater than
   the tolerance.

The current month is never analyzed, even if its apparent totals balance.
Months are analyzed oldest first. A continuation page advances by month rather
than by individual record.

Records already reconciled or locked elsewhere are excluded. Per the approved
business decision, a later analysis may reconcile the remaining unlocked
records from that month rather than requiring the original whole-month set.

## Proposal Identity and Snapshots

Each month produces at most one proposal per run and rule. The earliest eligible
Bank Statement record, ordered by `data` then ID, is the technical base record
needed by the existing proposal model. The UI does not expose this distinction:
all Bank records are presented as one source group.

The proposal stores an immutable monthly summary containing:

- rule key and version;
- calendar month;
- predicates and directional operator;
- tolerance and read-only 31-day property;
- source and destination counts;
- source and destination totals;
- calculated difference;
- technical base record identity;
- analysis timestamp and signature.

Individual immutable memberships are stored in a child proposal-membership
table rather than in one oversized JSON array. Each membership stores the
proposal ID, role (`source` or `destination`), source type, source ID, stable
ordinal, date, amount, description, account where applicable, and a compact row
snapshot. A unique key prevents a source record from appearing twice in one
proposal.

The membership rows and proposal snapshot cannot be rewritten after the
proposal reaches a terminal state. Completed historical snapshots remain
unchanged on migration reapply.

## Analysis Lifecycle

The existing manual and scheduled run orchestration remains authoritative. The
new rule is one additional allowlisted dispatch branch, not a second automation
system.

Analysis must:

- snapshot exactly one managed rule and its current configuration;
- require the directional Bank-to-FDM source rule to exist with `-`;
- count and page closed months deterministically;
- persist monthly summaries and memberships atomically per proposal;
- avoid duplicate proposals for the same run, rule, and month;
- expose progress in month units;
- finalize the run as ready when executable or ambiguous proposals exist;
- finish cleanly with no visible proposals when no two-sided month exists.

An unfinished manual run continues to block choosing another rule, consistent
with current behavior. Scheduled batches continue to execute one child rule at
a time in configured priority order.

## Execution and Concurrency

Only `proposed` monthly results are executable. Ambiguous results remain visible
for review but cannot be selected.

Execution runs in one security-definer database transaction:

1. Lock the run, proposal, managed configuration, and directional source rule.
2. Lock all proposal members in deterministic source-type and source-ID order.
3. Re-query the closed month and verify that every snapshotted member still
   exists, remains unlocked, still satisfies its source predicate, and retains
   the snapshotted date, amount, description, and account values.
4. Verify that the complete currently eligible unlocked membership for the month
   exactly equals the proposal membership. A new, edited, deleted, or competing
   record makes the proposal stale rather than silently changing it.
5. Recalculate counts, totals, and difference from the locked database rows.
6. Require the rule definition, version, operator, tolerance, and immutable
   properties to match the run snapshot.
7. Create one financial reconciliation containing every member record.
8. Complete it atomically and link the proposal to the reconciliation.

The reconciliation origin is `Automatic - Manual` or `Automatic - Scheduled`.
For a non-zero difference within tolerance, execution uses force completion with
a generated audit comment. The generated comment always includes the rule name,
calendar month, both counts and totals, tolerance, and final difference.

Execution is idempotent. Retrying an already completed proposal returns its
existing reconciliation. Any unexpected failure rolls back all reconciliation
writes and locks, persists only a sanitized proposal/run failure through the
established lifecycle, and allows later scheduled rules to continue safely.

## Source-Rule Protection

The rule depends on `import_cgd_extrato_ordem -> import_fdm_accounts (-)`.
The source-rule Settings API and its atomic replacement RPC must reject changing
or removing that direction while this managed automation rule is installed.
The analysis and execution paths independently fail closed if the dependency
drifts.

## Settings Interface

Settings -> Reconciliation -> Automatic reconciliation shows a fifth managed
rule card with:

- display name **Card Payments - POS - Income**;
- read-only source, destination, predicates, monthly matching description,
  operator, version, and 31-day property;
- editable difference tolerance, initialized to EUR 7,500.00;
- editable manual and scheduled execution flags;
- existing priority reordering controls.

The rule is installed disabled for both execution modes. Settings replacement
continues to require every installed managed rule exactly once, now including
this fifth rule.

## Proposal Review Interface

When manual execution is enabled, the rule appears in the existing Automatic
reconciliation rule selector. Proposal rows retain the existing three-column
desktop layout:

1. Status, selection, month, counts, totals, tolerance, and difference.
2. Complete CGD Bank Statement monthly group.
3. Complete FDM Accounts monthly group.

Both record groups start collapsed. Expanding a group calls a paginated,
app-authorized API/RPC that returns 50 immutable membership snapshots at a time,
oldest first. `Load more` appends the next page without collapsing or rerendering
the other group. The response includes a total count and rejects invalid,
oversized, foreign-run, or unauthorized requests.

Each row shows date, description, amount, account where relevant, and record ID.
All text is escaped. The narrow layout stacks the same three semantic sections
without changing ordering or accessibility.

Executable proposals are selected by default and can be deselected and selected
again. Ambiguous proposals show both groups but have no execution checkbox.
Finished runs retain only selected persisted outcomes, matching current behavior.

## Reconciliation and History Presentation

The completed reconciliation stores the actual Bank and FDM items, so existing
history aggregation shows:

- CGD Bank Statement count and total;
- FDM Accounts count and total;
- total record count;
- automatic origin;
- completed status;
- final difference;
- generated completion comment.

Opening the reconciliation continues to use the existing Manual reconciliation
details flow. The compact automatic proposal review remains the primary scalable
place to inspect the original monthly membership snapshot.

## API and Security Boundary

All database access remains RPC-only through service-role server handlers. New
tables use RLS and grant no direct access to anonymous or authenticated roles.
New security-definer functions use a fixed `search_path`, literal allowlisted
dispatch, bounded page sizes, strict UUID/rule/version validation, and no dynamic
SQL.

Public API mapping exposes only approved camel-case fields and strips database
diagnostics recursively. Manual actions bind the authenticated actor; scheduled
actions require the existing protected cron authentication and parent/child run
identity checks.

## Performance

Indexes must support:

- Bank Statement closed-month scanning with the `POS VENDAS` predicate;
- FDM Accounts month scanning for exact account `Credit Card`, excluding exact
  category `TransferOutToAccount` while retaining `NULL` categories;
- unlocked-record exclusion by source type and source ID;
- proposal membership paging by proposal, role, ordinal;
- deterministic proposal uniqueness by run, rule, and month.

Analysis aggregates before materializing member snapshots. The main run-detail
response returns monthly summaries only; individual memberships are loaded only
when expanded. Tests must exercise at least 1,000 records in one month.

## Error Handling

Expected business changes produce stable outcomes rather than raw errors:

- changed membership or row data -> `stale`;
- difference above tolerance during analysis -> `ambiguous`;
- no records on one side -> no visible proposal;
- changed rule definition, operator, or configuration -> stale/fail closed;
- competing execution -> one transaction succeeds and the other returns the
  persisted outcome or a safe conflict;
- unexpected database error -> sanitized failed outcome with no partial
  reconciliation.

The UI retains the current run and selections on transport uncertainty, reloads
the authoritative persisted state, and never reports success from an uncertain
response.

## Migration and Rollout

Ship one normal, reapply-safe migration after migration 11. It installs the fifth
managed rule, membership storage, indexes, RPC replacements, dispatch branches,
source-rule protection, ACLs, and safe legacy adaptation. Reapplying it must not
modify completed runs, proposals, reconciliations, or comments.

The migration is applied manually to Supabase before enabling the rule. The rule
remains disabled until an administrator explicitly enables manual or scheduled
execution.

## Verification

Automated Node and transactional PostgreSQL tests cover:

- five-rule settings and allowlists;
- closed-month, year-boundary, and leap-year behavior;
- case-insensitive `POS VENDAS`, exact `Credit Card`, exact
  `TransferOutToAccount` exclusion, and `NULL` category eligibility;
- missing-side invisibility;
- proposed and ambiguous EUR 7,500.00 boundaries;
- remaining-unlocked-record behavior;
- deterministic month identity and duplicate prevention;
- 1,000-record proposal membership and 50-row paging;
- reanalysis after other reconciliations consume records;
- stale detection after insert, edit, delete, source-rule change, configuration
  change, and competing reconciliation;
- atomic completion, generated comment, history totals, audit snapshots,
  idempotency, and rollback;
- manual execution and sequential scheduled execution;
- failed-child continuation and cross-slot resume;
- migration reapply and ACL behavior;
- unchanged behavior for all existing four managed rules.

Release requires:

1. Node syntax and full-suite success.
2. Transactional PostgreSQL smoke success after applying and reapplying the new
   migration.
3. Authenticated desktop and narrow-screen browser verification of Settings,
   analysis, collapsed groups, pagination, execution, and History.
4. Protected non-production scheduled-heartbeat verification before production
   enablement.
