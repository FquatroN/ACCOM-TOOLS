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

## Files to configure

- `supabase.sql`: SQL schema and policies to run in Supabase
- `config.js`: frontend credentials used for authentication only

## 1) Create Supabase project

1. Create a project in Supabase.
2. Open SQL Editor and run `supabase.sql`.
3. If you changed policies before and see `new row violates row-level security policy`, run `supabase.sql` again to reset all old policies.
4. Copy `Project URL` and `anon public key` from Settings -> API.

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

`SUPABASE_SERVICE_ROLE_KEY` is used only server-side by `/api/communications`.
Automatic email delivery is executed by `/api/email-automation` via Vercel Cron.

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

## Important deployment note

Never put service-role keys or raw Postgres connection strings in frontend code.

## Fallback mode

If `config.js` is empty, app runs in local mode (localStorage) and seeded sample rows.
