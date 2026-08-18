# Communications Log App

Web app for team communications with:

- Existing fields: `Data`, `Hora`, `Pessoa`, `O que aconteceu?`
- New fields: `Status`, `Category`
- Categories: `Warning`, `Maintenance`, `Information`, `very important`
- Inline add row and inline row editing

## What is implemented

- Supabase database integration
- Login page (`/gate.html`) with email/password authentication
- Protected app page (`/index.html`) only for authenticated users
- Server-side API layer on Vercel (`/api/communications`) for CRUD
- Excel import from sheet `Comunicações` into database
- CSV export
- Backoffice → Reconciliation workbench for Financial Documents, FDM Accounts,
  CGD Credit Card, and CGD Bank Statement records

## Files to configure

- `supabase.sql`: SQL schema and policies to run in Supabase
- `config.js`: frontend credentials used for authentication only

## 1) Create Supabase project

1. Create a project in Supabase.
2. Open SQL Editor and run `supabase.sql`.
3. If you changed policies before and see `new row violates row-level security policy`, run `supabase.sql` again to reset all old policies.
4. Copy `Project URL` and `anon public key` from Settings -> API.

### Financial reconciliation migration

Run `supabase-migrations/2026-08-09-financial-reconciliation.sql` in the
Supabase SQL Editor before using Backoffice → Reconciliation.

The migration creates the reconciliation, item, and audit tables, plus the
atomic RPCs that enforce the eligibility floor (`2026-01-01`), record locks,
completion rules, reopening history, and the supported source combinations.

### Automatic financial reconciliation rollout

After the base reconciliation migrations, apply the automatic reconciliation
migrations in this exact order:

1. `supabase-migrations/2026-08-14-financial-reconciliation-automation-schema.sql`
2. `supabase-migrations/2026-08-14-financial-reconciliation-automation-analysis.sql`
3. `supabase-migrations/2026-08-14-financial-reconciliation-automation-execution.sql`
4. `supabase-migrations/2026-08-15-financial-reconciliation-automation-analysis-performance.sql`
5. `supabase-migrations/2026-08-15-financial-reconciliation-automation-candidate-index-lookup.sql`
6. `supabase-migrations/2026-08-16-financial-reconciliation-automation-banco-v2.sql`
7. `supabase-migrations/2026-08-16-financial-reconciliation-automation-90-day-performance.sql`
8. `supabase-migrations/2026-08-16-financial-reconciliation-automation-credit-card-rule.sql`
9. `supabase-migrations/2026-08-17-financial-reconciliation-automation-amount-only-rules.sql`
10. `supabase-migrations/2026-08-18-financial-reconciliation-automation-proposal-details.sql`

If the database is already current through Banco v2, apply migrations 7 and 8
and then migration 9 in that order. If it is already current through the
90-day migration, apply migrations 8 and 9 in that order. If it is already
current through the Credit Card migration, apply only migration 9. The Credit
Card and amount-only migrations are reapply-safe: they verify immutable managed
definitions without overwriting saved administrator flags, editable day window,
or priority. The supported **Max difference in days** range is 0-90 days.
Installations current through migration 9 apply only migration 10.

### Amount-only rollout sequence

Publish compatibility-tolerant application code **before** applying migration 9.
Before the manual database migration, Settings safely accepts and shows the
existing two managed rules. After migration 9, the same Settings screen shows
all four managed rules and requires all four in an atomic Settings save. This
ordering avoids a Settings outage while the application and database catalogs
briefly differ.

Supabase migrations are a manual database operation; Vercel deploys the
application and invokes the scheduled HTTP endpoint, but does **not** apply SQL
migrations. In the Supabase SQL Editor, run migration 9 after the preceding
migrations. With a controlled disposable or development database and `psql`,
the equivalent manual apply and explicit reapply are:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f supabase-migrations/2026-08-17-financial-reconciliation-automation-amount-only-rules.sql
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f supabase-migrations/2026-08-17-financial-reconciliation-automation-amount-only-rules.sql
```

Both new amount-only rules start disabled for manual and scheduled execution.
Their `Difference allowed` values are fixed, read-only `0.00 €`; only enabled
state, manual/scheduled participation, maximum day difference, and priority are
administratively editable. Do not change the existing identity-rule configuration
as part of this rollout.

After applying the migration, run the transaction-safe database smoke suite:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

Run the same command a second time against the disposable or development
database. The smoke rolls back its fixtures, applies migration 9 once in normal
order and once as an explicit reapply, and must complete cleanly on both
invocations before enabling the new rules.

### Proposal details migration 10 rollout

Publish compatibility-tolerant application code **before** applying migration 10.
Supabase migrations are a manual database operation; Vercel deploys the
application and invokes the scheduled HTTP endpoint, but does **not** apply SQL
migrations. In the Supabase SQL Editor, run migration 10 after the preceding
migrations. With a controlled disposable or development database and `psql`,
the equivalent manual apply and explicit reapply are:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f supabase-migrations/2026-08-18-financial-reconciliation-automation-proposal-details.sql
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f supabase-migrations/2026-08-18-financial-reconciliation-automation-proposal-details.sql
```

Migration 10 immediately enriches unfinished amount-only proposals with their
source display details. It leaves completed history and audit snapshots
unchanged, and it does not infer supplier data for CGD Bank Statement or CGD
Credit Card records.

After applying migration 10, run the transaction-safe database smoke suite
twice against the disposable or development database. The smoke rolls back its
fixtures, applies migration 10 once in normal order and once as an explicit
reapply, and must complete cleanly on both invocations before continuing
automatic execution:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

The SQL smoke and the protected authenticated-browser scenarios below are
mandatory rollout gates. Use a disposable or development environment for their
fixtures and manual execution checks.

For migration 10, use an authenticated browser session to verify all of the
following before continuing automatic execution:

1. Analyze the Bank amount-only rule. In column two, confirm the Financial
   Documents base record displays its document number, supplier name, supplier
   NIF, and description.
2. In column three, confirm every Bank destination and ambiguous candidate
   description displays without a supplier label.
3. Repeat the two display checks for the Credit Card amount-only rule.
4. Open an unfinished pre-migration run and confirm migration 10 enriches it
   immediately; execute a backfilled unfinished unique proposal and confirm it
   completes rather than becoming stale.
5. Open completed history and confirm its prior audit snapshot is unchanged.
6. Verify desktop three-column and narrow stacked layouts, with escaped
   punctuation in every detail field.

1. Confirm pre-migration Settings loads safely with the two existing rules.
2. After migration, confirm Settings lists four rules, both amount-only rules
   are disabled, and both fixed tolerances display read-only `0.00 €`.
3. Submit a tampered nonzero amount-only tolerance and confirm it is rejected
   atomically, with no partial Settings write.
4. Enable **Financial Documents to CGD Bank Account – AMOUNT ONLY** for manual
   analysis only. Validate unique, duplicate, cross-base-overlap, and no-match
   Banco fixtures; select it alone, confirm the selector locks during the run,
   execute a selected unique proposal, and confirm zero-difference history and
   audit evidence. Change a source before execution and confirm it becomes stale
   without creating a reconciliation.
5. Disable the Bank amount-only rule again. Enable **Financial Documents to CGD
   Credit Card – AMOUNT ONLY** for manual analysis only and repeat the equivalent
   Visa fixture, selector-locking, execution, history/audit, and stale-source
   checks.
6. Only after both controlled manual validations pass, enable scheduled
   participation for the required rules. Confirm the scheduled four-rule order
   (Bank Statement, Credit Card, Bank Account – AMOUNT ONLY, Credit Card –
   AMOUNT ONLY) and verify that a failed child still allows later children to
   run. Then inspect the resulting batch and child history before any production
   scheduled enablement.

For production-size verification, save the rule with 90 days, press Analyze,
and observe the processed/total progress until Ready. Require Ready within two minutes
on the current dataset. Compare proposal counts, ambiguity,
evidence, and reconciliation-history semantics with the existing rule. Confirm
the API and Vercel logs contain no HTTP 500 and no statement timeout before
enabling scheduled execution.

The Credit Card rule is disabled by default; its initial tolerance is `0.00`,
maximum date difference is 10 days, and priority is 2. The migration preserves
the existing Bank Statement configuration. Enable Credit Card for manual analysis with scheduled execution disabled. In an authenticated
non-production session, select Credit Card from the Automatic reconciliation
rule list and verify exact-`Visa` eligibility, one-to-four-card proposals,
ambiguity, hidden no-match rows with accurate counts, execution, history,
identity evidence, stale outcomes, selector locking, and reload/resume. Use
**Open automatic reconciliation** from Settings to confirm that navigation does
not start analysis. Complete this manual validation before enabling scheduled
execution.

For scheduled validation, enable the required managed rules in a disposable or
development environment and confirm their configured order appears as separate
child runs and separate history entries. Deliberately create a failed child and
confirm it does not block the next child. Do not rely on the production daily
schedule until the configured order and failure continuation have been observed.

Vercel calls `/api/reconciliation-automation-cron` every minute as a heartbeat.
The database—not the Vercel schedule—atomically claims at most one configured
once-daily slot in `Europe/Lisbon`, including across retries and daylight-saving
changes. Each heartbeat advances no more than 25 analysis records or processes
no more than 25 proposals, and a later heartbeat safely resumes remaining work.

In Settings → Reconciliation → Automatic reconciliation, inspect the last batch result
before and after scheduled enablement. Disabling the shared schedule or
a managed rule prevents future scheduled work without changing completed
reconciliations.

## 2) Configure app (`config.js`)

Fill `config.js`:

```js
window.APP_CONFIG = {
  SUPABASE_URL: "https://YOUR_PROJECT_ID.supabase.co",
  SUPABASE_ANON_KEY: "YOUR_SUPABASE_ANON_KEY",
};
```

Then open `index.html` in browser.
Users sign in via `gate.html`.

## 3) Configure Vercel environment variables

In Vercel project settings -> Environment Variables, add:

- `SUPABASE_URL`
- `SUPABASE_ANON_KEY`
- `SUPABASE_SERVICE_ROLE_KEY`
- `RESEND_API_KEY`
- `EMAIL_FROM` (example: `notifications@yourdomain.com`)
- `AUTOMATION_TIMEZONE` (optional, default: `Europe/Lisbon`)
- `CRON_SECRET` (a long random secret used to protect scheduled endpoints)

`SUPABASE_SERVICE_ROLE_KEY` is used only server-side by `/api/communications`.
Automatic email delivery is executed by `/api/email-automation` via Vercel Cron.
Vercel sends `CRON_SECRET` as `Authorization: Bearer <CRON_SECRET>` for each
automatic reconciliation heartbeat. The endpoint requires that exact value and
never exposes the secret to browser code.
The Reconciliation workbench calls the server-side `/api/reconciliation` route;
it is available to users with the `financial-reconciliation` app feature.

## 4) Deploy online

### Option A: Netlify

1. Push this folder to GitHub.
2. In Netlify: New site from Git.
3. Build command: leave empty.
4. Publish directory: project root.

### Option B: Vercel

1. Push this folder to GitHub.
2. In Vercel: Add New Project.
3. Framework preset: Other.
4. No build command needed.

## Reconciliation verification

Before deploying the reconciliation feature, run the local static checks:

```powershell
node --check api/_reconciliation.js
node --check api/reconciliation.js
node --test tests/reconciliation.test.js
git diff --check
```

In the development Supabase project, run
`tests/reconciliation-rpc.smoke.sql` in SQL Editor. It rolls back its fixtures
and should confirm that lifecycle actions retain audit history, while only
removing an item or deleting a reconciliation releases a locked source record.
Repeat the Backoffice → Reconciliation browser checks in a clean session before
production deployment.

## Important deployment note

Never put service-role keys or raw Postgres connection strings in frontend code.

## Fallback mode

If `config.js` is empty, app runs in local mode (localStorage) and seeded sample rows.
