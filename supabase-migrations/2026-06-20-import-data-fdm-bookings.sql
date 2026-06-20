create table if not exists public.import_fdm_bookings (
  id uuid primary key default gen_random_uuid(),
  import_batch text not null,
  source_type text not null default 'fdm-bookings',
  source_name text not null default '',
  source_row_number integer not null default 0,
  booking_number text not null,
  room_type text not null default '',
  room text not null default '',
  rate text not null default '',
  guest_name text not null default '',
  arrival_raw text not null default '',
  arrival_time text null,
  check_in_raw text not null default '',
  check_in_date date null,
  check_out_raw text not null default '',
  check_out_date date null,
  nights integer null,
  guests integer null,
  room_assigned text not null default '',
  status text not null default '',
  payment_status text not null default '',
  balance_due numeric(14,2) null,
  channel text not null default '',
  booking_date_raw text not null default '',
  booking_date date null,
  booking_time text null,
  country text not null default '',
  city text not null default '',
  invoice_total numeric(14,2) null,
  currency text not null default 'EUR',
  raw_payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create unique index if not exists import_fdm_bookings_booking_number_uidx on public.import_fdm_bookings (booking_number);
create index if not exists import_fdm_bookings_created_at_idx on public.import_fdm_bookings (created_at desc);
create index if not exists import_fdm_bookings_import_batch_idx on public.import_fdm_bookings (import_batch);
create index if not exists import_fdm_bookings_check_in_date_idx on public.import_fdm_bookings (check_in_date desc);
create index if not exists import_fdm_bookings_check_out_date_idx on public.import_fdm_bookings (check_out_date desc);
create index if not exists import_fdm_bookings_status_idx on public.import_fdm_bookings (status);
create index if not exists import_fdm_bookings_channel_idx on public.import_fdm_bookings (channel);
create index if not exists import_fdm_bookings_country_idx on public.import_fdm_bookings (country);
