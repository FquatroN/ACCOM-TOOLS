create table if not exists public.import_cgd_extrato_ordem (
  id uuid primary key default gen_random_uuid(),
  import_batch text not null,
  source_type text not null default 'cgd-extrato-ordem',
  source_name text not null default '',
  source_row_number integer not null default 0,
  row_key text not null,
  data_raw text not null default '',
  data date null,
  data_valor_raw text not null default '',
  data_valor date null,
  descritivo text not null default '',
  montante numeric(14,2) null,
  saldo numeric(14,2) null,
  raw_payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create unique index if not exists import_cgd_extrato_ordem_row_key_uidx
  on public.import_cgd_extrato_ordem (row_key);

create index if not exists import_cgd_extrato_ordem_created_at_idx
  on public.import_cgd_extrato_ordem (created_at desc);

create index if not exists import_cgd_extrato_ordem_import_batch_idx
  on public.import_cgd_extrato_ordem (import_batch);

create index if not exists import_cgd_extrato_ordem_data_idx
  on public.import_cgd_extrato_ordem (data desc);

create index if not exists import_cgd_extrato_ordem_data_valor_idx
  on public.import_cgd_extrato_ordem (data_valor desc);
