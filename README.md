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
migrations in this order:

1. `supabase-migrations/2026-08-14-financial-reconciliation-automation-schema.sql`
2. `supabase-migrations/2026-08-14-financial-reconciliation-automation-analysis.sql`
3. `supabase-migrations/2026-08-14-financial-reconciliation-automation-execution.sql`
4. `supabase-migrations/2026-08-15-financial-reconciliation-automation-analysis-performance.sql`
5. `supabase-migrations/2026-08-15-financial-reconciliation-automation-candidate-index-lookup.sql`
6. `supabase-migrations/2026-08-16-financial-reconciliation-automation-banco-v2.sql`
7. `supabase-migrations/2026-08-16-financial-reconciliation-automation-90-day-performance.sql`

If the database is already current through Banco v2, only
`2026-08-16-financial-reconciliation-automation-90-day-performance.sql` is
needed. The supported **Max difference in days** range is 0-90 days.

After applying the migration, run the transaction-safe database smoke suite:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

For production-size verification, save the rule with 90 days, press Analyze,
and observe the processed/total progress until Ready. Require Ready within two minutes
on the current dataset. Compare proposal counts, ambiguity,
evidence, and reconciliation-history semantics with the existing rule. Confirm
the API and Vercel logs contain no HTTP 500 and no statement timeout before
enabling scheduled execution.

The managed rule and shared schedule are disabled by default. Keep them
disabled while deploying, then validate representative matches, non-matches,
threshold boundaries, and ambiguities through manual analysis and execution in
a non-production environment. Complete manual validation before enabling scheduled execution.

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
