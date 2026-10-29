# Communications deletion audit trail Implementation Plan

> **For Codex:** REQUIRED SUB-SKILL: Use `superpowers:executing-plans` to implement this plan task-by-task.

**Goal:** Make every future Communications deletion attributable to the verified app user, retain a recoverable snapshot, and expose a read-only audit history to Communications Settings users.

**Architecture:** Replace the current service-role REST `DELETE` with a service-role call to one security-definer PostgreSQL RPC. That RPC locks the communication, writes an append-only `deleted` event containing the verified actor and full row snapshot, then deletes the row in the same transaction. A separate, settings-authorized API endpoint returns audit records for a read-only Communications Settings tab.

**Tech Stack:** Vanilla HTML/CSS/JavaScript, Vercel Node API handlers, Supabase PostgREST/RPC, PostgreSQL migration SQL, Node's built-in test runner.

## Scope and invariants

- This applies only to deletions performed after the migration is deployed. The 29 October 2026 deletions cannot be retrospectively attributed or restored without an independent backup.
- The public Communications delete interaction and its existing undo/re-create behavior stay unchanged.
- The audit insert and original-row delete must be atomic: if audit recording fails, no row is deleted.
- The actor comes only from `requireFeature`'s verified Supabase user, never from browser request input.
- Audit records are immutable to browser roles. Only the API using the service-role key can access them.
- The audit view is available only to users with `settings:communications`; it is read-only and deliberately has no restore action in this change.

## Task 1: Add the protected, atomic database audit primitive

**Files:**
- Create: `supabase-migrations/2026-10-29-communications-delete-audit.sql`
- Create: `tests/communications-audit.test.js`

**Step 1: Write the migration-shape tests first.**

In `tests/communications-audit.test.js`, read the migration and assert that it:

- creates `public.communications_audit_events` with an event ID, `action`, `communication_id`, `actor_user_id`, `actor_email`, `occurred_at`, and `record_snapshot jsonb`;
- enables RLS and does not grant browser roles (`anon` or `authenticated`) any table access;
- defines `public.delete_communication_with_audit` as `security definer` and revokes public/browser execution while granting only `service_role` access;
- selects the target communication `for update`, inserts the JSON snapshot before deleting, and returns the original deleted row.

**Step 2: Run the new test to confirm it fails.**

Run: `node --test tests/communications-audit.test.js`

Expected: FAIL because the migration does not yet exist.

**Step 3: Implement the migration.**

Create the append-only event table and supporting descending `occurred_at` index. Use a check constraint that only permits `action = 'deleted'`. Enable RLS, revoke table permissions from `public`, `anon`, and `authenticated`, and create no browser-facing policies.

Implement this database interface:

```sql
public.delete_communication_with_audit(
  p_communication_id uuid,
  p_actor_user_id uuid,
  p_actor_email text
) returns setof public.communications
```

The `plpgsql`, `security definer` function must set a safe `search_path`, lock exactly one target row with `FOR UPDATE`, return no rows without creating an event if the communication no longer exists, insert a snapshot via `to_jsonb(v_communication)`, and only then delete and return that same row. Do not split this into API-side read/insert/delete requests. Revoke execute from `public`, `anon`, and `authenticated`; grant execute to `service_role` only.

**Step 4: Run the focused test.**

Run: `node --test tests/communications-audit.test.js`

Expected: PASS.

## Task 2: Route deletes through the audit RPC and add the audit read API

**Files:**
- Modify: `api/communications.js`
- Create: `api/communications-audit.js`
- Modify: `tests/communications-audit.test.js`

**Step 1: Add failing handler tests.**

Extend the test file with module-mocked `_supabase` dependencies, following the existing handler test pattern. Cover:

- `DELETE /api/communications?id=<id>` invokes `requireFeature(req, "app", "communications")`, then exactly one `POST` to `rpc/delete_communication_with_audit` with `p_communication_id`, the verified user's `id`, and verified email; it must not make a direct `DELETE communications?...` call.
- A missing ID returns `400` before calling the RPC.
- A non-existent ID receives the RPC's empty result and returns the existing-compatible `{ rows: [] }` response without fabricating an audit event.
- The handler uses the identity returned by `requireFeature`, so a forged actor field in a request body cannot affect the RPC payload.
- `GET /api/communications-audit` requires `requireFeature(req, "settings", "communications")`, requests newest-first audit rows, and returns only the intended audit fields.
- A permissions failure from either handler is surfaced through `sendError`, and non-GET methods for the new audit endpoint receive `405`.

**Step 2: Run the focused test to confirm it fails.**

Run: `node --test tests/communications-audit.test.js`

Expected: FAIL because the handlers still use the direct delete path and the audit endpoint does not exist.

**Step 3: Implement the API changes.**

In `api/communications.js`, retain the existing `app:communications` authorization but keep its returned `{ user }`. For a delete, normalize and validate the ID, normalize the verified email, and invoke only:

```js
restQuery("rpc/delete_communication_with_audit", {
  method: "POST",
  body: {
    p_communication_id: id,
    p_actor_user_id: user.id,
    p_actor_email: user.email,
  },
})
```

Preserve the current successful `{ rows }` response shape so the browser’s deletion and undo flow continues to work.

Create `api/communications-audit.js` as a GET-only handler. Authorize it with `requireFeature(req, "settings", "communications")`, query a bounded newest-first event list through `restQuery`, and return `id`, action, original communication ID, actor user ID/email, timestamp, and snapshot. Do not accept write operations, filters that expand access, or any client-supplied actor data.

**Step 4: Run the focused test.**

Run: `node --test tests/communications-audit.test.js`

Expected: PASS.

## Task 3: Add the read-only audit tab to Communications Settings

**Files:**
- Modify: `index.html`
- Modify: `app-main.js`
- Modify: `tests/communications-audit.test.js`

**Step 1: Add UI presence/escaping tests first.**

Extend the test file to verify that the Communications Settings markup has Configuration and Audit tab controls plus an audit table/panel, and that `app-main.js` contains a dedicated audit fetch/render path. Test the renderer's mapping or its shared escaping helper with an audit snapshot containing HTML characters in person/message values so the view cannot inject stored record content as markup.

**Step 2: Run the focused test to confirm it fails.**

Run: `node --test tests/communications-audit.test.js`

Expected: FAIL because the audit tab and loader do not exist.

**Step 3: Implement the settings UI.**

Within the existing `settings-view-communications` section:

- Add two local tabs: **Configuration** (the current Categories and Automatic Emails content) and **Audit**.
- Keep Configuration as the default screen so existing settings behavior remains unchanged.
- Add a read-only Audit panel with a newest-first table. Columns: Deleted at, Deleted by, Original date/time, Person, Status, Category, and Message.
- Extend `state`, `els`, settings-section data loading, and event listeners so selecting Audit lazily calls `/api/communications-audit`, handles loading/error/empty states, and re-renders without a full page reload.
- Render actor information from `actor_email` (with user ID as a fallback only if an old/incomplete row lacks an email) and render snapshot fields through the existing HTML-escaping helper.
- Do not add edit, delete, or restore controls to the audit table.

**Step 4: Run the focused test.**

Run: `node --test tests/communications-audit.test.js`

Expected: PASS.

## Task 4: Verify the integrated behavior and prepare release

**Files:**
- Verify only: `supabase-migrations/2026-10-29-communications-delete-audit.sql`
- Verify only: `api/communications.js`
- Verify only: `api/communications-audit.js`
- Verify only: `index.html`
- Verify only: `app-main.js`
- Verify only: `tests/communications-audit.test.js`

**Step 1: Run targeted tests.**

Run: `node --test tests/communications-audit.test.js`

Expected: PASS.

**Step 2: Run the repository test suite.**

Run: `node --test tests/*.test.js`

Expected: the new audit tests pass. If the known pre-existing `reconciliation-density.test.js` `canAppFdmAccounts is not defined` failure remains, record it separately rather than changing unrelated reconciliation code.

**Step 3: Review the diff and migration safety.**

Run: `git diff --check` and `git diff -- supabase-migrations/2026-10-29-communications-delete-audit.sql api/communications.js api/communications-audit.js index.html app-main.js tests/communications-audit.test.js`

Confirm the database function is the only production delete path, audit data has no browser-role grants/policies, and the UI has no mutation controls.

**Step 4: Apply and deploy in the correct order when publishing is requested.**

Apply the Supabase migration before deploying the Vercel code, because the new delete route depends on the RPC. Then publish the committed branch and verify the production deployment is Ready. Smoke-test one disposable Communications deletion: confirm it disappears from the list and one audit entry shows the signed-in user's email and intact snapshot.
