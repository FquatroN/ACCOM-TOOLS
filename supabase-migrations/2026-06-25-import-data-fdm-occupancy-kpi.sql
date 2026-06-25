create table if not exists public.import_fdm_occupancy_kpi (
  id uuid primary key default gen_random_uuid(),
  import_batch text not null default '',
  source_type text not null default 'fdm-occupancy-kpi',
  source_name text not null default '',
  source_row_number integer not null default 0,
  row_key text not null,
  mes_raw text not null default '',
  mes date,
  room_type_orig text not null default '',
  occup_percent numeric(8,2),
  charge numeric(14,2),
  average numeric(14,2),
  revpar numeric(14,2),
  booked integer,
  oc_combinada numeric(8,2),
  property text not null default '',
  type_orig text not null default '',
  type text not null default '',
  room_type text not null default '',
  room_type_orig_transf text not null default '',
  raw_payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default timezone('utc', now())
);

create unique index if not exists import_fdm_occupancy_kpi_row_key_uidx
  on public.import_fdm_occupancy_kpi (row_key);

create index if not exists import_fdm_occupancy_kpi_created_at_idx
  on public.import_fdm_occupancy_kpi (created_at desc);

create index if not exists import_fdm_occupancy_kpi_import_batch_idx
  on public.import_fdm_occupancy_kpi (import_batch);

create index if not exists import_fdm_occupancy_kpi_mes_idx
  on public.import_fdm_occupancy_kpi (mes desc);

create index if not exists import_fdm_occupancy_kpi_room_type_orig_idx
  on public.import_fdm_occupancy_kpi (room_type_orig);
