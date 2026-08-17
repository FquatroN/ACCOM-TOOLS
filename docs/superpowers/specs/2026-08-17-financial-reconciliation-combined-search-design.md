# Financial Reconciliation Combined Search Design

## Goal

Replace the separate **Description** and **Supplier Search** filters for Manual Reconciliation → Financial Documents with one **Description / Supplier Search** field.

## Behavior

- The combined field performs a case-insensitive partial search.
- A Financial Documents record matches when the entered term appears in any one of:
  - `financial_documents.description`
  - the workspace supplier-name projection
  - the workspace supplier-NIF projection
- The three conditions use OR semantics.
- The separate Supplier Search field is removed from the Financial Documents filter list.
- Payment and Category remain LOV filters.
- Filters for every other reconciliation source remain unchanged.

## Implementation Shape

Reuse the existing `description` filter key to preserve the client/API contract. For the Financial Documents source only:

1. Render the `description` field with the label **Description / Supplier Search** and an appropriate search placeholder.
2. Remove `supplier` from the source-provided `filterFields` list.
3. Replace the Financial Documents description predicate with one grouped OR predicate covering description, supplier name, and supplier NIF.

The workspace RPC remains the authoritative filtering boundary. The browser continues to submit one JSON filter value and does not attempt client-side filtering.

## Compatibility and Errors

- Existing empty description values continue to mean “no text filter.”
- Existing callers that send `description` remain compatible and gain the broader Financial Documents search behavior.
- The migration must be re-runnable and must fail closed if the installed workspace-function definition is not a recognized version.
- No change is made to reconciliation eligibility, locking, paging, ordering, or lifecycle behavior.

## Verification

- UI contract: Financial Documents renders one combined text field and no separate Supplier Search field.
- Request contract: Financial Documents sends `description` and omits `supplier`; other sources retain their declared filters.
- Database smoke coverage: the same combined field independently matches description, supplier name, and supplier NIF, while an unrelated term returns no candidates.
- Existing reconciliation and full Node suites remain green.
