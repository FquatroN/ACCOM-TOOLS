# Financial Reconciliation Design

## Purpose

Add a **Reconciliation** option to the ACCOMTOOLS Backoffice. It enables an authenticated user to create auditable many-to-many reconciliations between financial documents, FDM Accounts, CGD credit-card imports, and CGD bank-statement imports.

The module supports an explicit reconciliation workflow. Users may complete a balanced group only after clicking **Complete**. They may force-complete an unbalanced group only after providing a mandatory comment. Existing reconciliations may be reopened without losing the audit record.

## Confirmed rules

- Only source records dated on or after **2026-01-01** are eligible for reconciliation.
- A source record may belong to only one reconciliation at a time, whether that reconciliation is Started or Complete.
- Adding an item locks it immediately. Removing it deletes the current membership and makes the source record available again.
- A reconciliation is Started after its first source item is added. It becomes Complete only when the user explicitly completes it.
- Exact-zero difference is required for normal completion. A non-zero difference requires **Force complete** and a non-empty comment.
- Reopen changes a Complete reconciliation back to Started, keeps its current items attached and locked, and writes an audit event. The user can then add or remove items.
- `financial_documents` are eligible only when `fat = 'S'`.
- `import_fdm_accounts` are eligible when `invoice_flag = true` or `category = 'Compras'`.
- Reconciliation uses the existing `import_cgd_extrato_ordem.montante` field as the bank-statement amount. The business shorthand `amount` refers to this field.
- Only Financial Documents-led reconciliations may combine more than two source types. Their group difference is `Σ financial_documents.amount + Σ import_cgd_extrato_ordem.montante + Σ import_cgd_cartao_credito.valor + Σ import_fdm_accounts.amount`.
- FDM Account-, Credit Card-, and Bank Statement-led reconciliations select one matching source type only and use that pair's approved formula.

## Data model

### `financial_reconciliations`

One row per reconciliation group.

| Field | Purpose |
| --- | --- |
| `id` | UUID primary key. |
| `status` | `started` or `complete`. Source-row Not started status is derived, not stored here. |
| `base_source_type` | Source type selected when the reconciliation started. It controls the allowed matching modes. |
| `matching_source_types` | JSON array of matching source types. It may contain all three non-document sources only for a Financial Documents-led group; otherwise it contains exactly one type. |
| `completion_type` | `normal` or `forced` after completion; `null` while Started. |
| `difference_amount` | Persisted calculated difference at the latest state-changing action. |
| `forced_completion_comment` | Required for forced completion and otherwise `null`. |
| `created_by`, `created_at`, `updated_at` | Ownership and lifecycle timestamps. |
| `completed_by`, `completed_at` | Finalization attribution, cleared only when reopened. |
| `deleted_by`, `deleted_at` | Soft-deletion attribution. Deleted groups are hidden from normal history. |

### `financial_reconciliation_items`

One row per source record included in a reconciliation.

| Field | Purpose |
| --- | --- |
| `id` | UUID primary key. |
| `reconciliation_id` | Foreign key to `financial_reconciliations`. |
| `source_type` | One of `financial_documents`, `import_fdm_accounts`, `import_cgd_cartao_credito`, or `import_cgd_extrato_ordem`. |
| `source_id` | UUID of the row in the typed source table. |
| `amount_snapshot` | The amount used in the reconciliation calculation when the record was added. |
| `created_by`, `created_at` | Membership attribution. |

The child table has a unique constraint on `(source_type, source_id)`. This is the authoritative lock against double reconciliation. It deliberately uses a typed reference rather than a database foreign key because the source may be in any of four tables.

### `financial_reconciliation_audit`

Append-only event rows containing the reconciliation ID, event type, acting user, timestamp, difference after the event, optional comment, and JSON metadata for item IDs or prior values. Events are written for create, item added, item removed, normal complete, forced complete, reopen, and delete. No update or delete API is exposed for audit rows. Reconciliations use soft deletion so an audit row always retains its reconciliation reference.

## Compatibility and calculations

The source adapter controls which tables can be selected and which rows appear in the workbench. All rows must meet the date rule and must have no `financial_reconciliation_items` membership.

| Base source | Allowed matching source | Additional eligibility | Calculation that must equal zero |
| --- | --- | --- | --- |
| Financial document | Bank statement, credit card, and/or FDM Account | Financial document `fat = 'S'`; FDM `category = 'Compras'` | `Σ financial_documents.amount + Σ import_cgd_extrato_ordem.montante + Σ import_cgd_cartao_credito.valor + Σ import_fdm_accounts.amount` |
| FDM Account | Bank statement | FDM `invoice_flag = true` or `category = 'Compras'` | `import_fdm_accounts.amount - import_cgd_extrato_ordem.montante` |
| Credit card | Financial document | Financial document `fat = 'S'` | `financial_documents.amount + import_cgd_cartao_credito.valor` |
| Credit card | Bank statement | None beyond the common date and lock rules | `import_cgd_cartao_credito.valor + import_cgd_extrato_ordem.montante` |
| Bank statement | Financial document, credit card, or FDM Account | Select exactly one matching source type; the matching source's eligibility rule applies | `financial_documents.amount + import_cgd_extrato_ordem.montante`, `import_cgd_cartao_credito.valor + import_cgd_extrato_ordem.montante`, or `import_fdm_accounts.amount - import_cgd_extrato_ordem.montante`, respectively. |

The calculation service applies the approved formula for the selected source combination and displays the running difference. A Financial Documents-led group may hold multiple eligible records from every listed source; every other group may hold multiple records but from only its base source and one selected matching source. It stores each source amount in `amount_snapshot` so the confirmed reconciliation remains reproducible if later import data changes.

## User experience

### Navigation and workbench

Reconciliation is a new Backoffice navigation entry and uses the existing ACCOMTOOLS visual system: warm background, teal actions, rounded white cards, and the existing filter/table patterns.

The workbench has:

- a header with the 2026 eligibility rule and a **Start reconciliation** action;
- derived counts for Not started, Started, and Complete;
- a dynamic left-hand source table with the exact fields and filters defined in the feature brief;
- a right-hand basket containing the reconciliation members, their source types, amounts, running difference, Save draft, Complete, and Force complete actions; and
- recent reconciliation history with status, difference, details, audit history, and Reopen action.

The left table changes fields and filters per source:

- Financial documents: document date, description, supplier NIF, supplier name, payment, amount; date, amount, contains, and payment filters.
- FDM Accounts: account, event date, category, reservation ID, amount; date, amount, account, and category filters.
- Credit cards: date, description, value; date, value, and description filters.
- Bank statements: date, description, amount (`montante`); date, amount, and description filters.

### Status presentation

- **Not started (red):** an eligible source record has no current reconciliation item.
- **Started (yellow):** a reconciliation group has members but has not been completed.
- **Complete (green):** a group was explicitly completed, normally or forcibly. Forced groups visibly retain their non-zero difference and comment.

## Operations and validation

All mutation endpoints execute in a database transaction and validate server-side. The UI may hide invalid choices but cannot bypass these checks.

1. Start creates a Started reconciliation and adds the chosen base record atomically.
2. Add item validates date, lock, compatibility, source-specific eligibility, and the base-source matching-mode limit before creating a membership item and audit event.
3. Remove item deletes the membership row and writes an audit event. The source record becomes available immediately.
4. Complete validates a zero difference and changes the group to Complete with `completion_type = normal`.
5. Force complete validates a non-zero difference and a non-blank comment, then sets `completion_type = forced`.
6. Reopen changes the group to Started, preserves its memberships, clears completion metadata and any forced-completion comment, and writes an audit event.
7. Delete soft-deletes the reconciliation and removes its current membership rows in one transaction, writes a deletion audit event, and releases every member.

If a concurrent request has already linked an item, the transaction fails without partial writes. The UI shows a clear conflict message and refreshes the affected source list while retaining the rest of the user's basket.

## Verification

Database and API tests cover the eligibility date, `fat = 'S'`, `invoice_flag = true`, `category = 'Compras'`, the Financial Documents-led four-source sum, each approved two-source calculation, the non-document matching-mode limit, membership uniqueness, forced-comment validation, normal completion, reopening, removal, deletion, and concurrent attempts to link the same source record.

UI tests cover table-specific fields and filters, locked-record exclusion, status counts, running difference, completion dialogs, audit display, and the responsive ACCOMTOOLS workbench layout.

## Out of scope

Automated match suggestions, importing new financial data, changing the four source schemas, and introducing new user roles are outside this feature. The module follows the application's existing authenticated-user access model.
