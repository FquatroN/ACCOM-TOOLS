# Communications deletion audit trail

## Intent

Make every future Communications deletion attributable to the authenticated user and preserve a recoverable copy of the deleted record. The prior 29 October deletions cannot be retroactively attributed because no user audit data was written at the time.

## Goals

- Record the actor's Supabase user ID and email, deletion timestamp, original Communication ID, and the complete pre-delete record.
- Keep the audit write and delete atomic: a deletion succeeds only if its audit event is stored.
- Let users with Communications Settings access review the audit entries in the app.
- Keep ordinary Communications users from reading or altering audit records directly.

## Non-goals

- Rebuild the records already deleted without a backup.
- Change the existing Communications delete button or introduce soft deletes.
- Provide self-service restoration in this change; administrators can use the preserved snapshot for controlled recovery.

## Approaches considered

1. **Append-only audit trail — selected.** A dedicated table stores the actor and record snapshot before deletion. It preserves the current user experience and supports later recovery.
2. **Soft delete.** Leave deleted rows in the main table. This complicates every existing Communications query and does not give a clear immutable actor history.
3. **Platform logs only.** The current server-side service-role requests cannot reliably identify the signed-in app user.

## Design

Add a protected `communications_audit_events` table and a database RPC that, in one transaction, reads the target Communication, records a `deleted` event with its JSON snapshot and supplied verified actor details, and deletes the original row. The API continues to authorize `app:communications`, but it passes the identity returned by token verification to this RPC rather than using an anonymous server-side action.

Add an authenticated read endpoint guarded by `settings:communications`. In Communications Settings, add a read-only Audit tab/table showing the deletion time, actor email, and original record details. The audit data has no direct public/authenticated-table permissions; the API is the only application path to it.

## Verification

Automated tests will prove that deletion invokes the atomic audit RPC with the verified user identity, that unauthorized audit reads are rejected, and that the settings UI exposes the read-only audit view. Existing Communications behavior remains covered by the full test suite.
