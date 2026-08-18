# Automatic Reconciliation Proposal Details — Design

## Goal

Show useful source-record details in the existing three-column automatic reconciliation proposal layout. Financial Documents base records must show their document number, supplier, supplier NIF, and description. CGD Bank Statement and CGD Credit Card destination or candidate records must show their description. Supplier data must not be inferred for destination records because those source tables do not contain a separate supplier field.

## Scope

This change applies to the two amount-only automatic reconciliation rules:

- `financial_documents_cgd_bank_statement_amount_only` version 1
- `financial_documents_cgd_credit_card_amount_only` version 1

The existing identity-based rules already persist and display these fields and must remain unchanged.

## Snapshot Contract

New amount-only proposal snapshots will contain the following display fields in addition to their existing identifiers, dates, and amounts.

### Financial Documents base snapshot

- `docNumber` from `financial_documents.doc_number`
- `description` from `financial_documents.description`
- `supplierName` from `financial_documents.supplier_name`
- `supplierNif` from `financial_documents.supplier_nif`

### CGD Bank Statement destination snapshot

- `description` from `import_cgd_extrato_ordem.descritivo`

### CGD Credit Card destination snapshot

- `description` from `import_cgd_cartao_credito.descricao`

No supplier value will be generated for either destination source.

## Display Behavior

The existing proposal renderer remains the presentation boundary. It already renders:

- document number and supplier name in the item metadata line;
- record description in the description line;
- source date and amount in their current locations;
- base records in the second column;
- every destination and ambiguous candidate group in the third column.

The renderer will add supplier NIF to the base-record metadata line when it is present. The migration enriches the proposal snapshots so that the renderer receives the required data. Empty source values continue to render with the existing fallback behavior. All values remain escaped before insertion into the page.

## Unfinished-Run Backfill

The forward migration will enrich stored amount-only proposals whose parent run has `finished_at is null`. This includes normal proposal `items` and every candidate record in `candidate_groups`, including nested groups.

The backfill will:

- join records by their persisted `sourceType` and `sourceId`;
- enrich only the display fields defined in this specification;
- preserve proposal IDs, lifecycle statuses, amounts, evidence, configuration, signatures, and reconciliation links;
- leave completed runs and completed reconciliation audit metadata unchanged;
- leave an individual snapshot unchanged when its referenced source record no longer exists.

Updating unfinished snapshots is required for compatibility with execution-time snapshot revalidation. Once the adapters return the richer snapshot contract, an existing unfinished proposal must carry the same enriched values or normal execution would classify it as stale.

## Migration Design

Add one dated, forward-only migration after the existing 2026-08-17 amount-only migration. The migration will:

1. Replace only the amount-only Bank Statement and Credit Card candidate adapters so future analysis emits the richer snapshots.
2. Enrich unfinished amount-only proposal `base_snapshot`, `items`, and `candidate_groups` atomically.
3. Preserve array order and nested candidate-group structure.
4. Be idempotent and safe to reapply.
5. Preserve the existing function signatures, ownership, fixed search paths, privileges, and literal rule dispatch.

Prior migrations will not be edited.

## Failure and Audit Behavior

- Missing source records do not receive invented or derived display information.
- A proposal with a missing or changed source record remains subject to the existing stale revalidation behavior.
- Completed runs are immutable for this change, even if their older snapshots lack display details.
- Completed reconciliation audit entries are never rewritten.
- Backfill failure aborts the migration transaction rather than leaving a partially enriched run.

## Verification

Transactional PostgreSQL smoke coverage will prove:

- new Banco and Visa amount-only proposals include all required base details;
- Bank Statement and Credit Card destinations include their descriptions;
- proposed items and flat or nested ambiguous candidate groups use the same enriched snapshot contract;
- an unfinished pre-migration proposal is enriched and remains executable;
- completed proposal snapshots and audit metadata remain byte-for-byte unchanged;
- missing referenced source records remain unchanged;
- migration reapplication is idempotent.

UI behavior tests will use amount-only-shaped proposals to prove the existing renderer:

- displays document number, supplier name, supplier NIF, and base description;
- displays destination descriptions without inventing a supplier;
- places all destinations and candidate groups in the third column;
- escapes every enriched field.

The transactional PostgreSQL smoke is a mandatory rollout gate. If no local PostgreSQL runtime is available during implementation, that limitation must be reported explicitly and the smoke must be run in Supabase before publication.
