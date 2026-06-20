create table if not exists public.import_fdm_sales (
  id uuid primary key default gen_random_uuid(),
  import_batch text not null,
  source_type text not null default 'fdm-sales',
  source_name text not null default '',
  source_row_number integer not null default 0,
  reservation_id text not null default '',
  sale_date_raw text not null default '',
  sale_date date not null,
  sale_time text not null default '',
  sale_item text not null default '',
  quantity numeric(14,2) not null,
  price numeric(14,2) null,
  net_price numeric(14,2) null,
  tax numeric(14,2) null,
  total numeric(14,2) null,
  total_net numeric(14,2) null,
  total_tax numeric(14,2) null,
  user_name text not null default '',
  guest text not null default '',
  financial_account text not null default '',
  note text not null default '',
  raw_payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create unique index if not exists import_fdm_sales_key_uidx
  on public.import_fdm_sales (sale_date, sale_time, sale_item, quantity, guest);
create index if not exists import_fdm_sales_created_at_idx on public.import_fdm_sales (created_at desc);
create index if not exists import_fdm_sales_import_batch_idx on public.import_fdm_sales (import_batch);
create index if not exists import_fdm_sales_sale_date_idx on public.import_fdm_sales (sale_date desc);
create index if not exists import_fdm_sales_sale_item_idx on public.import_fdm_sales (sale_item);
create index if not exists import_fdm_sales_reservation_id_idx on public.import_fdm_sales (reservation_id);
create index if not exists import_fdm_sales_user_name_idx on public.import_fdm_sales (user_name);
