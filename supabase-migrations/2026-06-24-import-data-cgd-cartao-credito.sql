create table if not exists public.import_cgd_cartao_credito (
  id uuid primary key default gen_random_uuid(),
  import_batch text not null default '',
  source_type text not null default 'cgd-cartao-credito',
  source_name text not null default '',
  source_row_number integer not null default 0,
  row_key text not null,
  data_raw text not null default '',
  data date,
  data_valor_raw text not null default '',
  data_valor date,
  descricao text not null default '',
  debito numeric(14,2),
  credito numeric(14,2),
  valor numeric(14,2) generated always as (coalesce(credito, 0) - coalesce(debito, 0)) stored,
  raw_payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default timezone('utc', now())
);

create unique index if not exists import_cgd_cartao_credito_row_key_uidx
  on public.import_cgd_cartao_credito (row_key);

create index if not exists import_cgd_cartao_credito_created_at_idx
  on public.import_cgd_cartao_credito (created_at desc);

create index if not exists import_cgd_cartao_credito_import_batch_idx
  on public.import_cgd_cartao_credito (import_batch);

create index if not exists import_cgd_cartao_credito_data_idx
  on public.import_cgd_cartao_credito (data desc);

create index if not exists import_cgd_cartao_credito_data_valor_idx
  on public.import_cgd_cartao_credito (data_valor desc);
