# Supplier Invoice Email Schedules Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add supplier-specific monthly invoice email schedules that deliver the previous calendar month's original documents and invoice summary through Resend.

**Architecture:** Dedicated Supabase schedule and delivery-run records keep configuration distinct from operational history. A protected Vercel Cron controller evaluates Portugal-local due dates, atomically claims a supplier-period run, downloads the existing Google Drive attachments, and sends one Resend email using a stable idempotency key. The existing Financial Documents settings screen receives a Supplier Emails tab for schedule management and failure retry.

**Tech Stack:** Vanilla HTML/CSS/JavaScript, Vercel serverless functions and Cron, Supabase/PostgREST and PostgreSQL functions, Google Drive API, Resend Email API, Node built-in test runner.

**Spec:** `docs/superpowers/specs/2026-09-24-financial-documents-supplier-invoice-emails-design.md`

## Global Constraints

- The schedule day is an integer from 1 through 31 and always selects the previous calendar month's `document_date` range.
- Select every matching supplier document regardless of its status.
- Attach original Google Drive files and include Invoice date, Supplier, Description, and Amount in HTML and plain-text email bodies.
- Use the configured subject exactly as entered; recipient input is comma-, semicolon-, or newline-separated, normalized, validated, and deduplicated.
- Do not send a partial invoice pack: missing files, download errors, unsupported types, or an oversized attachment bundle fail before Resend is called.
- Use `Europe/Lisbon`, process normal scheduled runs at 09:00 local time, and protect Cron using `CRON_SECRET`/Vercel Cron authentication.
- Use a stable Resend `Idempotency-Key` derived from the schedule and period; never allow a duplicate-risk retry after an uncertain provider outcome.
- Keep credentials server-only and expose only sanitized provider errors to users.
- Preserve the existing unrelated `reconciliation-density.test.js` baseline failure (`canAppFdmAccounts is not defined`); all newly added tests must pass.

## Review Focus

- A schedule set to day 31 must not silently run in a 30- or 28-day month; it runs only on the configured calendar day when that day exists. Test in Task 1.
- Supplier names with HTML characters and descriptions with user-provided markup must be escaped in both email forms. Test in Task 1.
- Duplicate, mixed-case recipient addresses supplied across commas, semicolons, and lines must yield one normalized recipient list. Test in Task 1.
- A record without `drive_file_id` or a bundle whose Base64 representation exceeds Resend's 40 MB ceiling must reach no Resend call. Test in Task 3.
- A timeout after email submission must become `uncertain`, retain the stable provider idempotency key, and expose no Retry control. Test in Tasks 3 and 4.

---

## File Structure

- Create: `supabase-migrations/2026-09-24-financial-document-supplier-email-schedules.sql` — tables, indexes, RLS, and RPCs that atomically claim/finalize delivery attempts.
- Create: `api/_financial-docs-supplier-emails.js` — pure schedule validation, Portuguese-time period/due calculation, recipient normalization, HTML/text email generation, and attachment-size validation.
- Create: `api/_financial-docs-supplier-email-service.js` — PostgREST/RPC access for schedules, runs, supplier-period documents, and safe client projections.
- Create: `api/_financial-docs-supplier-email-automation.js` — orchestration that claims a run, preflights Drive attachments, calls Resend, and finalizes safe statuses.
- Create: `api/financial-docs-supplier-email-schedules.js` — authenticated Financial Documents settings CRUD/list/retry endpoint.
- Create: `api/financial-docs-supplier-email-automation.js` — protected Cron endpoint for due schedules.
- Modify: `api/_financial-docs-service.js` — reuse its Drive token refresh and file download helpers from the automation layer; do not duplicate Drive OAuth behavior.
- Modify: `app-main.js` — load/render/edit schedules, integrate the fourth settings tab, and render safe last-run/retry state.
- Modify: `index.html` — add Supplier Emails tab and accessible schedule-table panel.
- Modify: `styles.css` — compact responsive settings-table styles and error/success status presentation.
- Modify: `vercel.json` — register the protected minute-level cron route used to evaluate 09:00 Lisbon schedules through DST.
- Modify: `README.md` — document the route, existing Resend/Drive/Cron prerequisites, and monthly invoice-email behavior.
- Create: `tests/financial-docs-supplier-email.test.js` — unit and source-contract tests for the new pure module, endpoint guards, automation, migration, and UI wiring.

### Task 1: Define the schedule contract and database-safe delivery lifecycle

**Files:**
- Create: `supabase-migrations/2026-09-24-financial-document-supplier-email-schedules.sql`
- Create: `api/_financial-docs-supplier-emails.js`
- Test: `tests/financial-docs-supplier-email.test.js`

**Interfaces:**
- Consumes: `cleanText`-compatible primitive values and the Financial Document entity shape `{ nif, name }`.
- Produces: `normalizeSupplierEmailSchedule(input, entities)`, `previousCalendarMonth(now, timeZone)`, `isSupplierEmailScheduleDue(schedule, now, timeZone)`, `buildSupplierInvoiceEmail(input)`, `validateAttachmentBundle(files, rawByteLimit)`, and database RPCs `claim_financial_document_supplier_email_run`, `complete_financial_document_supplier_email_run`, and `retry_financial_document_supplier_email_run`.

- [ ] **Step 1: Write failing pure-contract and migration tests**

```js
test('normalizes one supplier schedule and rejects invalid values', () => {
  assert.deepEqual(
    normalizeSupplierEmailSchedule(
      { supplierNif: 'PT 123', supplierName: 'A & B', recipients: 'A@EXAMPLE.COM; a@example.com\nb@example.com', dayOfMonth: '5', subject: 'Invoices' },
      [{ nif: '123', name: 'A & B' }]
    ),
    { supplierNif: '123', supplierName: 'A & B', recipients: ['a@example.com', 'b@example.com'], dayOfMonth: 5, subject: 'Invoices', enabled: true }
  );
  assert.throws(() => normalizeSupplierEmailSchedule({ dayOfMonth: 32 }, []), /day/i);
});

test('uses the complete prior calendar month and exact existing day in Lisbon', () => {
  assert.deepEqual(previousCalendarMonth(new Date('2026-03-05T09:00:00Z'), 'Europe/Lisbon'), { start: '2026-02-01', end: '2026-02-28', key: '2026-02' });
  assert.equal(isSupplierEmailScheduleDue({ enabled: true, dayOfMonth: 31, recipients: ['a@example.com'] }, new Date('2026-04-30T08:00:00Z'), 'Europe/Lisbon').due, false);
});

test('migration provides atomic schedule-period claiming and no duplicate sent run', () => {
  const sql = readFileSync('supabase-migrations/2026-09-24-financial-document-supplier-email-schedules.sql', 'utf8');
  assert.match(sql, /financial_document_supplier_email_schedules/);
  assert.match(sql, /claim_financial_document_supplier_email_run/);
  assert.match(sql, /unique/i);
});
```

- [ ] **Step 2: Run the focused test to verify it fails**

Run: `node --test tests/financial-docs-supplier-email.test.js`

Expected: FAIL because the module and migration do not exist.

- [ ] **Step 3: Implement the pure schedule/email helpers**

```js
function previousCalendarMonth(now, timeZone) {
  const { year, month } = getClockParts(now, timeZone);
  const end = new Date(Date.UTC(year, month - 1, 0));
  const start = new Date(Date.UTC(end.getUTCFullYear(), end.getUTCMonth(), 1));
  return { start: isoDate(start), end: isoDate(end), key: isoDate(start).slice(0, 7) };
}

function isSupplierEmailScheduleDue(schedule, now, timeZone) {
  const clock = getClockParts(now, timeZone);
  return {
    due: !!schedule.enabled && schedule.recipients.length > 0 && clock.hour === 9 && clock.minute === 0 && clock.day === schedule.dayOfMonth,
    slotKey: `${clock.year}-${String(clock.month).padStart(2, '0')}-${String(clock.day).padStart(2, '0')}:09:00`,
  };
}
```

Implement `escapeHtml` before producing the table. Build text and HTML from the same ordered document projection. Preserve the configured subject without interpolation. Calculate Base64-safe payload capacity from raw bytes rather than trusting stored metadata alone.

- [ ] **Step 4: Add the migration and lifecycle RPCs**

Create schedule and run tables, RLS policies, supplier/period indexes, and server-only RPC execution privileges. Store immutable schedule/recipient/subject/document snapshots on a run. Make the claim RPC reject an already `running`, `sent`, `skipped`, or `uncertain` schedule-period and return a typed `claimed`/`already_processed` result. Make retry create a new attempt only for a known-safe `failed` run; it must reject `sent`, `skipped`, and `uncertain` runs.

```sql
create unique index financial_document_supplier_email_active_delivery_idx
  on public.financial_document_supplier_email_runs (schedule_id, period_start)
  where status in ('running', 'sent', 'skipped', 'uncertain');
```

- [ ] **Step 5: Run the focused test to verify it passes**

Run: `node --test tests/financial-docs-supplier-email.test.js`

Expected: PASS for the pure helper and migration-contract tests.

- [ ] **Step 6: Commit**

```bash
git add supabase-migrations/2026-09-24-financial-document-supplier-email-schedules.sql api/_financial-docs-supplier-emails.js tests/financial-docs-supplier-email.test.js
git commit -m "feat: add supplier invoice email schedule contract"
```

### Task 2: Expose authenticated schedule management and supplier-period data

**Files:**
- Create: `api/_financial-docs-supplier-email-service.js`
- Create: `api/financial-docs-supplier-email-schedules.js`
- Modify: `api/_financial-docs-service.js`
- Modify: `tests/financial-docs-supplier-email.test.js`

**Interfaces:**
- Consumes: Task 1's normalized schedule shape and claim/finalize/retry RPCs; `_financial-docs-service.js` exports `listFinancialDocumentEntities`, `loadFinancialDocsSettings`, `refreshDriveAccessToken`, and `downloadDriveFile`.
- Produces: `listSupplierEmailSchedules()`, `createSupplierEmailSchedule(input)`, `updateSupplierEmailSchedule(id, input)`, `deleteSupplierEmailSchedule(id)`, `listSupplierPeriodDocuments(schedule, period)`, `listSupplierEmailRuns(scheduleId)`, and the authenticated JSON endpoint `/api/financial-docs-supplier-email-schedules`.

- [ ] **Step 1: Write failing endpoint and query contract tests**

```js
test('schedule endpoint requires Financial Documents settings access and validates suppliers', async () => {
  const response = await invokeScheduleHandler({ method: 'POST', body: { supplierNif: '999', recipients: 'a@example.com', dayOfMonth: 5, subject: 'Invoices' } });
  assert.equal(response.statusCode, 400);
  assert.match(response.body.error, /supplier/i);
});

test('supplier-period lookup uses supplier identity and document_date only, not status', () => {
  const path = buildSupplierPeriodDocumentsPath({ supplierNif: '123', start: '2026-09-01', end: '2026-09-30' });
  assert.match(path, /supplier_nif=eq\.123/);
  assert.match(path, /document_date=gte\.2026-09-01/);
  assert.doesNotMatch(path, /status=/);
});
```

- [ ] **Step 2: Run the focused test to verify it fails**

Run: `node --test tests/financial-docs-supplier-email.test.js`

Expected: FAIL because the data service and endpoint do not exist.

- [ ] **Step 3: Implement the service layer and schedule CRUD endpoint**

Use `restQuery` for the schedule/run projections and the Task 1 RPCs for state transitions. Before creating or updating a schedule, load `financial_document_entities` and require the selected NIF/name to resolve to one existing entity. Return only safe fields: no Drive access token, Resend API key, or unfiltered provider diagnostic.

```js
async function listSupplierPeriodDocuments({ supplierNif, periodStart, periodEnd }) {
  return restQuery(
    `financial_documents?select=id,document_date,supplier_nif,supplier_name,description,amount,drive_file_id,stored_filename,mime_type,file_size&supplier_nif=eq.${encodeURIComponent(supplierNif)}&document_date=gte.${periodStart}&document_date=lte.${periodEnd}&order=document_date.asc,id.asc`,
    { method: 'GET' }
  );
}
```

Implement `GET`, `POST`, `PUT`, and `DELETE` with `requireFeature(req, 'settings', 'financial-docs')`. Reserve retry dispatch for Task 3's orchestration path; this endpoint may validate and request it but must not expose a Cron bypass.

- [ ] **Step 4: Run the focused test to verify it passes**

Run: `node --test tests/financial-docs-supplier-email.test.js`

Expected: PASS for permission, validation, and exact supplier-period query tests.

- [ ] **Step 5: Commit**

```bash
git add api/_financial-docs-supplier-email-service.js api/financial-docs-supplier-email-schedules.js api/_financial-docs-service.js tests/financial-docs-supplier-email.test.js
git commit -m "feat: add supplier invoice email schedule API"
```

### Task 3: Send an all-or-nothing invoice pack through protected automation

**Files:**
- Create: `api/_financial-docs-supplier-email-automation.js`
- Create: `api/financial-docs-supplier-email-automation.js`
- Modify: `api/_financial-docs-supplier-email-service.js`
- Modify: `tests/financial-docs-supplier-email.test.js`

**Interfaces:**
- Consumes: Task 1 schedule/date/email helpers and run RPCs; Task 2 schedule/document service; existing `refreshDriveAccessToken`, `downloadDriveFile`, and Financial Documents settings loader.
- Produces: `processSupplierEmailSchedule({ schedule, now, trigger })`, `processDueSupplierEmailSchedules({ now })`, `sendSupplierInvoiceEmail(payload)`, and protected route `/api/financial-docs-supplier-email-automation`.

- [ ] **Step 1: Write failing automation tests**

```js
test('automation fails before Resend when one selected document has no Drive attachment', async () => {
  const result = await processSupplierEmailSchedule({ schedule, now, documents: [{ id: 'a', drive_file_id: '' }] });
  assert.equal(result.status, 'failed');
  assert.equal(fetchCalls.filter((call) => call.url === 'https://api.resend.com/emails').length, 0);
});

test('automation sends one provider request with document table, attachments, and stable idempotency key', async () => {
  const result = await processSupplierEmailSchedule({ schedule, now, documents: [attachedDocument] });
  assert.equal(result.status, 'sent');
  assert.equal(resendRequest.headers['Idempotency-Key'], `financial-document-supplier-email/${schedule.id}/2026-09`);
  assert.match(JSON.parse(resendRequest.body).html, /Invoice date/);
});

test('timeout after provider submission is uncertain and cannot enter retry flow', async () => {
  const result = await processSupplierEmailSchedule({ schedule, now, fetchImpl: timeoutFetch });
  assert.equal(result.status, 'uncertain');
  await assert.rejects(() => retrySupplierEmailRun(result.runId), /uncertain/i);
});
```

- [ ] **Step 2: Run the focused test to verify it fails**

Run: `node --test tests/financial-docs-supplier-email.test.js`

Expected: FAIL because the automation controller and protected route do not exist.

- [ ] **Step 3: Implement preflight, delivery, and finalization**

Claim the run before loading documents. If the prior-month query is empty, finalize as `skipped` and do not contact Resend. Refresh the existing Drive credential once per run, download each `drive_file_id`, check the real buffer size and allowed MIME type, then call `validateAttachmentBundle` before the first Resend request.

```js
const resendResponse = await fetch('https://api.resend.com/emails', {
  method: 'POST',
  headers: {
    Authorization: `Bearer ${process.env.RESEND_API_KEY}`,
    'Content-Type': 'application/json',
    'Idempotency-Key': `financial-document-supplier-email/${schedule.id}/${period.key}`,
  },
  body: JSON.stringify({ from, to: schedule.recipients, subject: schedule.subject, html, text, attachments }),
});
```

Finalize `sent` with the Resend message ID only after a successful response. Finalize confirmed rejections and preflight failures as retryable `failed`; finalize transport ambiguity as `uncertain` and never expose a normal retry for it. Include only sanitized failure messages in API responses.

- [ ] **Step 4: Implement Cron and manual-retry authorization boundaries**

The automation route accepts only `GET`/`POST`. It recognizes Vercel Cron or a valid `Authorization: Bearer ${CRON_SECRET}` exactly as existing automation routes do; otherwise it requires Financial Documents settings access only for an explicit retry action. A normal user request must identify an existing failed run and must never process arbitrary schedule IDs or bypass the server-side period calculation.

- [ ] **Step 5: Run the focused test to verify it passes**

Run: `node --test tests/financial-docs-supplier-email.test.js`

Expected: PASS for all-or-nothing preflight, idempotent Resend request, provider uncertainty, due selection, and route authorization tests.

- [ ] **Step 6: Commit**

```bash
git add api/_financial-docs-supplier-email-automation.js api/financial-docs-supplier-email-automation.js api/_financial-docs-supplier-email-service.js tests/financial-docs-supplier-email.test.js
git commit -m "feat: automate supplier invoice email delivery"
```

### Task 4: Build the Financial Documents settings tab and safe retry UX

**Files:**
- Modify: `index.html`
- Modify: `app-main.js`
- Modify: `styles.css`
- Modify: `tests/financial-docs-supplier-email.test.js`

**Interfaces:**
- Consumes: `/api/financial-docs-supplier-email-schedules` list/CRUD responses and Task 3's explicit retry action result.
- Produces: a `supplier-emails` Financial Documents settings tab, `loadFinancialDocsSupplierEmailSchedules()`, `renderFinancialDocsSupplierEmailSchedules()`, `saveFinancialDocsSupplierEmailSchedule()`, and `retryFinancialDocsSupplierEmailRun(runId)`.

- [ ] **Step 1: Write failing UI wiring tests**

```js
test('Financial Documents settings exposes Supplier Emails as a fourth tab', () => {
  assert.match(indexHtml, /id="financial-docs-settings-supplier-emails-tab"/);
  assert.match(appMain, /setFinancialDocsSettingsTab\("supplier-emails"\)/);
});

test('schedule UI requires supplier, recipients, day, and subject and hides retry for uncertain results', () => {
  const invalid = validateSupplierEmailScheduleDraft({ supplierNif: '', recipients: '', dayOfMonth: 0, subject: '' });
  assert.equal(invalid.valid, false);
  assert.doesNotMatch(renderSupplierEmailRun({ status: 'uncertain' }), /Retry/);
  assert.match(renderSupplierEmailRun({ status: 'failed', retryable: true }), /Retry/);
});
```

- [ ] **Step 2: Run the focused test to verify it fails**

Run: `node --test tests/financial-docs-supplier-email.test.js`

Expected: FAIL because the fourth tab and client helpers do not exist.

- [ ] **Step 3: Add accessible markup and client state**

Add the tab button beside Attributes, Rules, and Drive settings. Extend `state.financialDocsSettingsTab`, the `els` registry, tab rendering, and tab activation to accept `supplier-emails` without changing existing tab behavior. Use a semantic table with an add/edit form containing a supplier `<select>`, recipient `<textarea>`, day `<input type="number" min="1" max="31">`, required subject input, enabled checkbox, status cell, and row actions.

```js
function setFinancialDocsSettingsTab(tab) {
  state.financialDocsSettingsTab = ['attributes', 'rules', 'drive', 'supplier-emails'].includes(tab) ? tab : 'attributes';
  renderFinancialDocsSettingsTabs();
}
```

Populate supplier options from the already loaded Financial Document Entities, but defer saving authority to the server. Escape all server-returned user fields with existing `escapeHtml`. Preserve keyboard focus after saving/deleting a row and show the API's safe result in the existing settings status/toast conventions.

- [ ] **Step 4: Wire CRUD and retry requests**

Load schedules on entering the Supplier Emails tab and after each successful mutation. Send typed JSON to the schedule endpoint; keep unsaved drafts local. The Retry button must appear only when `status === 'failed' && retryable === true`, request the server-side retry action by run ID, then reload the row/result. Never offer Retry for `sent`, `skipped`, `running`, or `uncertain`.

- [ ] **Step 5: Run the focused test to verify it passes**

Run: `node --test tests/financial-docs-supplier-email.test.js`

Expected: PASS for tab accessibility, required inputs, safe result rendering, and retry visibility.

- [ ] **Step 6: Commit**

```bash
git add index.html app-main.js styles.css tests/financial-docs-supplier-email.test.js
git commit -m "feat: add supplier invoice email settings tab"
```

### Task 5: Register deployment behavior and complete end-to-end verification

**Files:**
- Modify: `vercel.json`
- Modify: `README.md`
- Modify: `tests/financial-docs-supplier-email.test.js`

**Interfaces:**
- Consumes: protected automation route from Task 3 and user configuration from Task 4.
- Produces: Vercel Cron registration for `/api/financial-docs-supplier-email-automation`, operational deployment documentation, and a green feature test suite.

- [ ] **Step 1: Write failing deployment/documentation contract tests**

```js
test('deployment registers the protected supplier invoice email cron and documents prerequisites', () => {
  const vercel = JSON.parse(readFileSync('vercel.json', 'utf8'));
  assert.ok(vercel.crons.some((cron) => cron.path === '/api/financial-docs-supplier-email-automation' && cron.schedule === '* * * * *'));
  assert.match(readFileSync('README.md', 'utf8'), /supplier invoice email/i);
  assert.match(readFileSync('README.md', 'utf8'), /RESEND_API_KEY/);
  assert.match(readFileSync('README.md', 'utf8'), /CRON_SECRET/);
});
```

- [ ] **Step 2: Run the focused test to verify it fails**

Run: `node --test tests/financial-docs-supplier-email.test.js`

Expected: FAIL because deployment and README entries do not exist.

- [ ] **Step 3: Register and document the feature**

Add a minute-level Vercel cron entry so server code—not UTC cron syntax—can apply 09:00 `Europe/Lisbon` correctly across daylight-saving transitions. Document that Financial Documents Drive must be connected, Resend values (`RESEND_API_KEY`, `EMAIL_FROM`) and `CRON_SECRET` must be configured, the system selects `document_date` from the prior month, no empty email is sent, and the 40 MB Resend attachment limit is protected by preflight.

- [ ] **Step 4: Run focused feature tests**

Run: `node --test tests/financial-docs-supplier-email.test.js`

Expected: PASS.

- [ ] **Step 5: Run the full suite and inspect the diff**

Run: `node --test tests/*.test.js`

Expected: All supplier-email tests pass; the only known baseline failure remains `reconciliation settings destination uses actual app feature guards` with `canAppFdmAccounts is not defined`.

Run: `git diff main...HEAD --check`

Expected: no whitespace errors.

- [ ] **Step 6: Commit**

```bash
git add vercel.json README.md tests/financial-docs-supplier-email.test.js
git commit -m "docs: configure supplier invoice email automation"
```

## Plan Self-Review

### Spec coverage

- Supplier, recipient list, day, subject, and enabled configuration are implemented in Tasks 1, 2, and 4.
- Previous-calendar-month selection without status filtering is implemented and tested in Task 2.
- Body table and original attachments are implemented in Tasks 1 and 3.
- Resend delivery, size handling, all-or-nothing behavior, provider idempotency, and uncertain delivery protection are implemented and tested in Task 3.
- Dedicated audit records, duplicate protection, and safe retry are implemented in Task 1 and surfaced in Task 4.
- Lisbon scheduling, protected Cron, deployment configuration, and operational documentation are implemented in Tasks 3 and 5.

### Placeholder scan

Searched this plan for unresolved planning placeholders and vague implementation directives; none are present.

### Type consistency

The schedule contract from Task 1 (`supplierNif`, `supplierName`, `recipients`, `dayOfMonth`, `subject`, `enabled`) is the only configuration shape used by Tasks 2–4. The task output statuses (`running`, `sent`, `failed`, `skipped`, `uncertain`) match the migration and client rendering contract. The supplier-period shape (`periodStart`, `periodEnd`, `key`) is used consistently by selection, delivery, and idempotency-key generation.

### Review Focus coverage

Each listed failure mode has an explicit focused test in its referenced task: day 31 in Task 1, escaping and recipient normalization in Task 1, attachment preflight in Task 3, and timeout/uncertainty in Tasks 3 and 4.
