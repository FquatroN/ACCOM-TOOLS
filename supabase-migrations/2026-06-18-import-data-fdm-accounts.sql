create table if not exists public.import_fdm_accounts (
  id uuid primary key default gen_random_uuid(),
  import_batch text not null,
  source_type text not null default 'fdm-accounts',
  source_name text not null default '',
  source_row_number integer not null default 0,
  account text not null,
  date_time_raw text not null,
  event_date date null,
  event_time text null,
  category text not null,
  amount numeric(14,2) not null,
  reservation_id text not null default '',
  guest text not null default '',
  reporting_date_raw text not null default '',
  reporting_date date null,
  user_name text not null default '',
  description text not null default '',
  bill_number text not null default '',
  item text not null default '',
  invoice_number text not null default '',
  currency text not null default 'EUR',
  invoice_amount numeric(14,2) null,
  designation text not null default '',
  invoice text not null default '',
  invoice_flag boolean null,
  raw_payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create index if not exists import_fdm_accounts_created_at_idx on public.import_fdm_accounts (created_at desc);
create index if not exists import_fdm_accounts_import_batch_idx on public.import_fdm_accounts (import_batch);
create index if not exists import_fdm_accounts_event_date_idx on public.import_fdm_accounts (event_date desc);
create index if not exists import_fdm_accounts_category_idx on public.import_fdm_accounts (category);
create index if not exists import_fdm_accounts_account_idx on public.import_fdm_accounts (account);
create index if not exists import_fdm_accounts_reservation_id_idx on public.import_fdm_accounts (reservation_id);
