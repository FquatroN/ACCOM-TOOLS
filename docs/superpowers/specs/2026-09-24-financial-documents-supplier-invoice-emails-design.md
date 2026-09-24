# Supplier invoice email schedules — design

## Purpose

Financial Documents users need to send each supplier's invoice pack automatically at the beginning of the following month. A schedule specifies the supplier, recipients, day of the month, and email subject. On its scheduled day, the system sends every financial document for that supplier whose **document date** falls in the previous calendar month.

For example, a schedule configured for day 5 runs on 5 October and sends documents dated 1–30 September. The configured day may be any day from 1 through 31; it never changes the selected invoice period.

## Confirmed requirements

- Add a **Supplier Emails** tab to Financial Documents settings.
- A schedule has a selectable supplier, a list of email addresses, a day of the month, and a subject.
- Send all matching financial documents regardless of their status.
- Attach the original stored document files to the email.
- Include an email-body table with invoice date, supplier, description, and amount.
- Deliver through Resend.
- Use the existing Google Drive connection to retrieve stored document files.

## Chosen approach

Use dedicated database records for schedules and delivery runs, rather than embedding this operational state in the existing Financial Documents settings JSON.

This approach gives each monthly delivery an immutable audit trail and permits an atomic claim before calling Resend. Retries, concurrent Vercel Cron invocations, and a manually retried failed run therefore cannot send the same supplier-period pack twice.

## Settings experience

The new tab contains an addable schedule table. Each row exposes:

- **Supplier**: select from the existing Financial Document Entities list. The saved identity is the entity NIF and normalized name, so renamed display names do not change the intended supplier.
- **Emails**: a comma-, semicolon-, or newline-separated list. Saving normalizes, validates, deduplicates, and stores the addresses.
- **Day of the month**: integer 1–31.
- **Subject**: required text, used exactly as entered for the outgoing email.
- **Enabled**: permits a schedule to be paused without deleting its configuration.
- **Last result**: latest sent, failed, or skipped result with its timestamp.
- **Retry**: available for a failed run after the data or configuration has been corrected.

Only one active schedule is allowed for a supplier. This avoids overlapping monthly packs being delivered to different recipient lists. A user can pause, edit, or delete a schedule from the same tab.

## Data model

### `financial_document_supplier_email_schedules`

One configuration per supplier. Key fields:

- `id` UUID
- `supplier_nif`, `supplier_name`
- `recipients` JSONB array of normalized addresses
- `day_of_month` integer, constrained to 1–31
- `subject` text
- `enabled` boolean
- timestamps and user audit fields

A unique supplier identity constraint prevents duplicate active configurations and supports clear validation in the settings API.

### `financial_document_supplier_email_runs`

An append-only delivery audit. Key fields:

- `id`, `schedule_id`, `period_start`, `period_end`
- `status`: `running`, `sent`, `failed`, or `skipped`
- a snapshot of recipients and subject used for the run
- selected document IDs and attachment metadata
- attachment count and raw-byte total
- Resend message ID, sent timestamp, and safe error information
- timestamps and retry linkage

A uniqueness constraint on the schedule and the calendar period makes each scheduled delivery idempotent. The database record is created/claimed before fetching attachments or calling Resend.

## Scheduling and selection

A protected Vercel Cron endpoint runs continuously using the same `CRON_SECRET` guard as existing automations. It evaluates time in `Europe/Lisbon` and only processes due schedules at 09:00 local time on their configured calendar day.

For a due schedule it calculates the prior calendar month's inclusive range, then selects every Financial Document with the saved supplier identity and a `document_date` in that range. Status is deliberately not filtered.

If the request is retried or two cron requests overlap, only the successful database claim may continue. A completed run is never resent automatically.

## Email assembly and delivery

The outbound email contains:

- the configured subject, unchanged;
- a short statement identifying the supplier and covered month;
- an HTML and plain-text invoice table with **Invoice date**, **Supplier**, **Description**, and **Amount**; and
- every selected document's original stored file as an attachment.

The service refreshes the existing Google Drive token, downloads the selected Drive files server-side, and submits their Base64 content and stored filenames to Resend.

Resend limits an email, including Base64-encoded attachments, to 40 MB. The implementation will enforce a conservative raw-file bundle limit before calling Resend. This leaves room for Base64 expansion and email content.

## Failure behavior

An invoice pack is all-or-nothing. The service does not send a partial pack:

- If any selected document lacks a Drive attachment, the run fails before sending and records the affected documents.
- If a Drive download fails, the file is unsupported, or the bundle exceeds the safe attachment threshold, the run fails before sending and records the reason.
- If Resend rejects the email, the run is marked failed with a sanitized provider error.
- If no matching documents exist, the run is marked skipped and no empty email is sent.

The settings tab surfaces the most recent result. A failed run may be retried explicitly after correction; successful runs remain immutable.

## Security and observability

- Settings endpoints require the existing Financial Documents settings feature permission.
- The automation endpoint accepts only protected Vercel Cron or valid bearer-secret requests; normal users cannot invoke it.
- The browser never receives Drive or Resend credentials.
- Recipients, subject, selected IDs, provider ID, and sanitized errors are retained in the run record for support and audit.
- User-facing status never exposes provider credentials or raw response diagnostics.

## Verification plan

- Unit-test settings validation, prior-month period calculation, supplier identity selection, and the day-1–31 schedule rules.
- Unit-test email-body table escaping, recipient normalization, no-document behavior, missing-file failures, attachment-size rejection, and Resend payload construction.
- Unit-test cron authentication and idempotent schedule-period claiming, including retry behavior.
- Add UI tests for the Supplier Emails tab, required subject, supplier selector, recipient validation, and failure/retry rendering.
- Run the full Node test suite. The current baseline contains one unrelated failure in `reconciliation-density.test.js` because `canAppFdmAccounts` is undefined; it must remain the only pre-existing failure unless separately addressed.
