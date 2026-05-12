create extension if not exists pgcrypto;

create or replace function public.set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = timezone('utc', now());
  return new;
end;
$$;

create table if not exists public.guest_records (
  id uuid primary key default gen_random_uuid(),
  ha text not null default 'H' check (ha in ('H', 'A')),
  name text not null default '',
  nationality text not null default '',
  nationality_code text not null default '',
  birth_date date not null,
  birth_place text not null default '',
  doc_number text not null default '',
  doc_type text not null default 'P' check (doc_type in ('P', 'O', 'B')),
  issuer_country text not null default '',
  issuer_country_code text not null default '',
  residence_country text not null default '',
  residence_country_code text not null default '',
  residence_city text not null default '',
  check_in date not null,
  check_out date not null,
  sent_status text not null default 'pending' check (sent_status in ('pending', 'sent', 'error')),
  sent_at timestamptz null,
  send_error text not null default '',
  send_batch_number integer not null default 0,
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now())
);

create unique index if not exists guest_records_doc_number_check_in_idx
  on public.guest_records (doc_number, check_in);

drop trigger if exists guest_records_set_updated_at on public.guest_records;
create trigger guest_records_set_updated_at
before update on public.guest_records
for each row execute procedure public.set_updated_at();

alter table public.guest_records enable row level security;

drop policy if exists "guest_records authenticated select" on public.guest_records;
create policy "guest_records authenticated select"
on public.guest_records
for select
to authenticated
using (true);

drop policy if exists "guest_records authenticated insert" on public.guest_records;
create policy "guest_records authenticated insert"
on public.guest_records
for insert
to authenticated
with check (true);

drop policy if exists "guest_records authenticated update" on public.guest_records;
create policy "guest_records authenticated update"
on public.guest_records
for update
to authenticated
using (true)
with check (true);

drop policy if exists "guest_records authenticated delete" on public.guest_records;
create policy "guest_records authenticated delete"
on public.guest_records
for delete
to authenticated
using (true);

grant select, insert, update, delete on public.guest_records to authenticated;
grant select, insert, update, delete on public.guest_records to anon;

create table if not exists public.guests_blacklist (
  id uuid primary key default gen_random_uuid(),
  name text not null default '',
  nationality text not null default '',
  nationality_code text not null default '',
  birth_date date null,
  doc_number text not null default '',
  what_happened text not null default '',
  occurrence_date date not null,
  who_reported text not null default '',
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now())
);

drop trigger if exists guests_blacklist_set_updated_at on public.guests_blacklist;
create trigger guests_blacklist_set_updated_at
before update on public.guests_blacklist
for each row execute procedure public.set_updated_at();

alter table public.guests_blacklist enable row level security;

drop policy if exists "guests_blacklist authenticated select" on public.guests_blacklist;
create policy "guests_blacklist authenticated select"
on public.guests_blacklist
for select
to authenticated
using (true);

drop policy if exists "guests_blacklist authenticated insert" on public.guests_blacklist;
create policy "guests_blacklist authenticated insert"
on public.guests_blacklist
for insert
to authenticated
with check (true);

drop policy if exists "guests_blacklist authenticated update" on public.guests_blacklist;
create policy "guests_blacklist authenticated update"
on public.guests_blacklist
for update
to authenticated
using (true)
with check (true);

drop policy if exists "guests_blacklist authenticated delete" on public.guests_blacklist;
create policy "guests_blacklist authenticated delete"
on public.guests_blacklist
for delete
to authenticated
using (true);

grant select, insert, update, delete on public.guests_blacklist to authenticated;
grant select, insert, update, delete on public.guests_blacklist to anon;

insert into public.guest_records (
  id,
  ha,
  name,
  nationality,
  nationality_code,
  birth_date,
  birth_place,
  doc_number,
  doc_type,
  issuer_country,
  issuer_country_code,
  residence_country,
  residence_country_code,
  residence_city,
  check_in,
  check_out,
  sent_status,
  sent_at,
  send_error,
  send_batch_number,
  created_at,
  updated_at
)
select
  case
    when coalesce(item ->> 'id', '') ~* '^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$'
      then (item ->> 'id')::uuid
    else gen_random_uuid()
  end as id,
  case when upper(coalesce(item ->> 'ha', 'H')) = 'A' then 'A' else 'H' end as ha,
  coalesce(item ->> 'name', '') as name,
  coalesce(item ->> 'nationality', '') as nationality,
  coalesce(item ->> 'nationalityCode', item ->> 'nationality_code', '') as nationality_code,
  nullif(coalesce(item ->> 'birthDate', item ->> 'birth_date', ''), '')::date as birth_date,
  coalesce(item ->> 'birthPlace', item ->> 'birth_place', '') as birth_place,
  coalesce(item ->> 'docNumber', item ->> 'doc_number', '') as doc_number,
  case
    when upper(coalesce(item ->> 'docType', item ->> 'doc_type', 'P')) in ('P', 'O', 'B')
      then upper(coalesce(item ->> 'docType', item ->> 'doc_type', 'P'))
    else 'P'
  end as doc_type,
  coalesce(item ->> 'issuerCountry', item ->> 'issuer_country', '') as issuer_country,
  coalesce(item ->> 'issuerCountryCode', item ->> 'issuer_country_code', '') as issuer_country_code,
  coalesce(item ->> 'residenceCountry', item ->> 'residence_country', '') as residence_country,
  coalesce(item ->> 'residenceCountryCode', item ->> 'residence_country_code', '') as residence_country_code,
  coalesce(item ->> 'residenceCity', item ->> 'residence_city', '') as residence_city,
  nullif(coalesce(item ->> 'checkIn', item ->> 'check_in', ''), '')::date as check_in,
  nullif(coalesce(item ->> 'checkOut', item ->> 'check_out', ''), '')::date as check_out,
  case
    when lower(coalesce(item ->> 'sentStatus', item ->> 'sent_status', 'pending')) in ('sent', 'error', 'pending')
      then lower(coalesce(item ->> 'sentStatus', item ->> 'sent_status', 'pending'))
    else 'pending'
  end as sent_status,
  nullif(coalesce(item ->> 'sentAt', item ->> 'sent_at', ''), '')::timestamptz as sent_at,
  coalesce(item ->> 'sendError', item ->> 'send_error', '') as send_error,
  greatest(0, coalesce(nullif(item ->> 'sendBatchNumber', '')::integer, nullif(item ->> 'send_batch_number', '')::integer, 0)) as send_batch_number,
  coalesce(nullif(item ->> 'createdAt', '')::timestamptz, nullif(item ->> 'created_at', '')::timestamptz, timezone('utc', now())) as created_at,
  coalesce(nullif(item ->> 'updatedAt', '')::timestamptz, nullif(item ->> 'updated_at', '')::timestamptz, timezone('utc', now())) as updated_at
from public.app_settings settings_row
cross join lateral jsonb_array_elements(coalesce(settings_row.payload -> 'rows', '[]'::jsonb)) as item
where settings_row.setting_key = 'guests'
  and coalesce(item ->> 'name', '') <> ''
  and nullif(coalesce(item ->> 'birthDate', item ->> 'birth_date', ''), '') is not null
  and nullif(coalesce(item ->> 'docNumber', item ->> 'doc_number', ''), '') is not null
  and nullif(coalesce(item ->> 'checkIn', item ->> 'check_in', ''), '') is not null
  and nullif(coalesce(item ->> 'checkOut', item ->> 'check_out', ''), '') is not null
on conflict (doc_number, check_in) do update set
  ha = excluded.ha,
  name = excluded.name,
  nationality = excluded.nationality,
  nationality_code = excluded.nationality_code,
  birth_date = excluded.birth_date,
  birth_place = excluded.birth_place,
  doc_type = excluded.doc_type,
  issuer_country = excluded.issuer_country,
  issuer_country_code = excluded.issuer_country_code,
  residence_country = excluded.residence_country,
  residence_country_code = excluded.residence_country_code,
  residence_city = excluded.residence_city,
  check_out = excluded.check_out,
  sent_status = excluded.sent_status,
  sent_at = excluded.sent_at,
  send_error = excluded.send_error,
  send_batch_number = excluded.send_batch_number,
  updated_at = excluded.updated_at;

insert into public.guests_blacklist (
  id,
  name,
  nationality,
  nationality_code,
  birth_date,
  doc_number,
  what_happened,
  occurrence_date,
  who_reported,
  created_at,
  updated_at
)
select
  case
    when coalesce(item ->> 'id', '') ~* '^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$'
      then (item ->> 'id')::uuid
    else gen_random_uuid()
  end as id,
  coalesce(item ->> 'name', '') as name,
  coalesce(item ->> 'nationality', '') as nationality,
  coalesce(item ->> 'nationalityCode', item ->> 'nationality_code', '') as nationality_code,
  nullif(coalesce(item ->> 'birthDate', item ->> 'birth_date', ''), '')::date as birth_date,
  coalesce(item ->> 'docNumber', item ->> 'doc_number', '') as doc_number,
  coalesce(item ->> 'whatHappened', item ->> 'what_happened', '') as what_happened,
  nullif(coalesce(item ->> 'occurrenceDate', item ->> 'occurrence_date', ''), '')::date as occurrence_date,
  coalesce(item ->> 'whoReported', item ->> 'who_reported', '') as who_reported,
  coalesce(nullif(item ->> 'createdAt', '')::timestamptz, nullif(item ->> 'created_at', '')::timestamptz, timezone('utc', now())) as created_at,
  coalesce(nullif(item ->> 'updatedAt', '')::timestamptz, nullif(item ->> 'updated_at', '')::timestamptz, timezone('utc', now())) as updated_at
from public.app_settings settings_row
cross join lateral jsonb_array_elements(coalesce(settings_row.payload -> 'blacklist', '[]'::jsonb)) as item
where settings_row.setting_key = 'guests'
  and (
    coalesce(item ->> 'name', '') <> ''
    or coalesce(item ->> 'docNumber', item ->> 'doc_number', '') <> ''
  )
  and nullif(coalesce(item ->> 'occurrenceDate', item ->> 'occurrence_date', ''), '') is not null
on conflict (id) do update set
  name = excluded.name,
  nationality = excluded.nationality,
  nationality_code = excluded.nationality_code,
  birth_date = excluded.birth_date,
  doc_number = excluded.doc_number,
  what_happened = excluded.what_happened,
  occurrence_date = excluded.occurrence_date,
  who_reported = excluded.who_reported,
  updated_at = excluded.updated_at;

update public.app_settings
set payload = jsonb_set(
  jsonb_set(coalesce(payload, '{}'::jsonb), '{rows}', '[]'::jsonb, true),
  '{blacklist}',
  '[]'::jsonb,
  true
)
where setting_key = 'guests';
