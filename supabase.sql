-- Run in Supabase SQL editor

create extension if not exists pgcrypto;

create table if not exists public.communications (
  id uuid primary key default gen_random_uuid(),
  date date not null,
  time time not null,
  person text not null,
  status text not null default 'Open',
  category text not null default 'Information',
  message text not null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.lost_found (
  id uuid primary key default gen_random_uuid(),
  item_number bigint generated always as identity unique,
  who_found text not null default '',
  who_recorded text not null default '',
  location_found text not null default '',
  object_description text not null default '',
  notes text not null default '',
  stored_location text not null default 'Receção',
  status text not null default 'Open',
  closed_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table public.communications
  add column if not exists updated_at timestamptz not null default now();

update public.communications
set updated_at = coalesce(updated_at, created_at, now())
where updated_at is null;

create or replace function public.set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists communications_set_updated_at on public.communications;
create trigger communications_set_updated_at
before update on public.communications
for each row execute function public.set_updated_at();

drop trigger if exists lost_found_set_updated_at on public.lost_found;
create trigger lost_found_set_updated_at
before update on public.lost_found
for each row execute function public.set_updated_at();

create table if not exists public.app_settings (
  id uuid primary key default gen_random_uuid(),
  setting_key text not null unique,
  payload jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now()
);

create table if not exists public.app_profiles (
  id uuid primary key default gen_random_uuid(),
  name text not null unique,
  app_features jsonb not null default '[]'::jsonb,
  settings_features jsonb not null default '[]'::jsonb,
  created_at timestamptz not null default now()
);

create table if not exists public.user_profile_assignments (
  user_id uuid primary key,
  profile_id uuid not null references public.app_profiles(id) on delete restrict,
  created_at timestamptz not null default now()
);

create table if not exists public.group_proposals (
  id uuid primary key default gen_random_uuid(),
  reservation_number text not null default '',
  creation_date date not null default current_date,
  name text not null,
  email text not null,
  check_in date not null,
  check_out date not null,
  guests integer not null check (guests between 1 and 60),
  guest_groups jsonb not null default '[]'::jsonb,
  room_items jsonb not null default '[]'::jsonb,
  total_value numeric(12, 2) not null default 0,
  option_date date,
  status text not null default 'Proposal',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.services (
  id uuid primary key default gen_random_uuid(),
  request_number bigint generated always as identity unique,
  service_type text not null,
  customer_name text not null,
  customer_email text not null default '',
  customer_phone text not null default '',
  pax integer not null check (pax between 1 and 60),
  notes text not null default '',
  service_date date not null,
  service_time time not null,
  pickup_location text not null default '',
  dropoff_location text not null default '',
  flight_number text not null default '',
  has_return boolean not null default false,
  return_pickup_location text not null default '',
  return_dropoff_location text not null default '',
  return_date date,
  return_time time,
  return_flight_number text not null default '',
  price numeric(12, 2) not null default 0,
  status text not null default 'Submitted',
  provider_user_id uuid,
  provider_email text not null default '',
  audit_log jsonb not null default '[]'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

drop trigger if exists group_proposals_set_updated_at on public.group_proposals;
create trigger group_proposals_set_updated_at
before update on public.group_proposals
for each row execute function public.set_updated_at();

drop trigger if exists services_set_updated_at on public.services;
create trigger services_set_updated_at
before update on public.services
for each row execute function public.set_updated_at();

create table if not exists public.properties (
  id uuid primary key default gen_random_uuid(),
  name text not null unique,
  active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.review_import_runs (
  id uuid primary key default gen_random_uuid(),
  property_id uuid not null references public.properties(id) on delete restrict,
  source text not null,
  file_name text not null default '',
  file_type text not null default '',
  upload_kind text not null default 'csv',
  status text not null default 'uploaded',
  row_count_detected integer not null default 0,
  row_count_imported integer not null default 0,
  error_summary text not null default '',
  created_by uuid,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.review_import_staging (
  id uuid primary key default gen_random_uuid(),
  import_run_id uuid not null references public.review_import_runs(id) on delete cascade,
  property_id uuid not null references public.properties(id) on delete restrict,
  source text not null,
  source_review_id text not null default '',
  source_reservation_id text not null default '',
  review_date date,
  reviewer_name text not null default '',
  reviewer_country text not null default '',
  language text not null default '',
  rating_raw numeric(8, 2),
  rating_scale numeric(8, 2),
  rating_normalized_100 numeric(8, 2),
  title text not null default '',
  positive_review_text text not null default '',
  negative_review_text text not null default '',
  body text not null default '',
  subscores jsonb not null default '{}'::jsonb,
  host_reply_text text not null default '',
  host_reply_date date,
  raw_text text not null default '',
  parse_confidence numeric(5, 2),
  dedupe_fingerprint text not null default '',
  warning_flags jsonb not null default '[]'::jsonb,
  raw_payload jsonb not null default '{}'::jsonb,
  selected_for_import boolean not null default true,
  is_valid boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.reviews (
  id uuid primary key default gen_random_uuid(),
  property_id uuid not null references public.properties(id) on delete restrict,
  import_run_id uuid references public.review_import_runs(id) on delete set null,
  source text not null,
  source_review_id text not null default '',
  source_reservation_id text not null default '',
  review_date date,
  reviewer_name text not null default '',
  reviewer_country text not null default '',
  language text not null default '',
  rating_raw numeric(8, 2),
  rating_scale numeric(8, 2),
  rating_normalized_100 numeric(8, 2),
  title text not null default '',
  positive_review_text text not null default '',
  negative_review_text text not null default '',
  body text not null default '',
  subscores jsonb not null default '{}'::jsonb,
  host_reply_text text not null default '',
  host_reply_date date,
  raw_text text not null default '',
  parse_confidence numeric(5, 2),
  dedupe_fingerprint text not null default '',
  raw_payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create unique index if not exists reviews_source_source_review_id_unique
on public.reviews (source, source_review_id)
where source_review_id <> '';

create unique index if not exists reviews_dedupe_fingerprint_unique
on public.reviews (dedupe_fingerprint)
where dedupe_fingerprint <> '';

alter table public.review_import_staging
  add column if not exists source_reservation_id text not null default '',
  add column if not exists positive_review_text text not null default '',
  add column if not exists negative_review_text text not null default '',
  add column if not exists subscores jsonb not null default '{}'::jsonb;

alter table public.reviews
  add column if not exists source_reservation_id text not null default '',
  add column if not exists positive_review_text text not null default '',
  add column if not exists negative_review_text text not null default '',
  add column if not exists subscores jsonb not null default '{}'::jsonb;

drop trigger if exists properties_set_updated_at on public.properties;
create trigger properties_set_updated_at
before update on public.properties
for each row execute function public.set_updated_at();

drop trigger if exists review_import_runs_set_updated_at on public.review_import_runs;
create trigger review_import_runs_set_updated_at
before update on public.review_import_runs
for each row execute function public.set_updated_at();

drop trigger if exists review_import_staging_set_updated_at on public.review_import_staging;
create trigger review_import_staging_set_updated_at
before update on public.review_import_staging
for each row execute function public.set_updated_at();

drop trigger if exists reviews_set_updated_at on public.reviews;
create trigger reviews_set_updated_at
before update on public.reviews
for each row execute function public.set_updated_at();

alter table public.communications enable row level security;
alter table public.lost_found enable row level security;
alter table public.app_settings enable row level security;
alter table public.app_profiles enable row level security;
alter table public.user_profile_assignments enable row level security;
alter table public.group_proposals enable row level security;
alter table public.services enable row level security;
alter table public.properties enable row level security;
alter table public.review_import_runs enable row level security;
alter table public.review_import_staging enable row level security;
alter table public.reviews enable row level security;

-- Clean reset: remove any existing policies on this table (including old custom names)
do $$
declare
  p record;
begin
  for p in
    select policyname
    from pg_policies
    where schemaname = 'public'
      and tablename = 'group_proposals'
  loop
    execute format('drop policy if exists %I on public.group_proposals', p.policyname);
  end loop;
end
$$;

do $$
declare
  p record;
begin
  for p in
    select policyname
    from pg_policies
    where schemaname = 'public'
      and tablename = 'services'
  loop
    execute format('drop policy if exists %I on public.services', p.policyname);
  end loop;
end
$$;

do $$
declare
  p record;
begin
  for p in
    select policyname
    from pg_policies
    where schemaname = 'public'
      and tablename = 'communications'
  loop
    execute format('drop policy if exists %I on public.communications', p.policyname);
  end loop;
end
$$;

do $$
declare
  p record;
begin
  for p in
    select policyname
    from pg_policies
    where schemaname = 'public'
      and tablename = 'properties'
  loop
    execute format('drop policy if exists %I on public.properties', p.policyname);
  end loop;
end
$$;

do $$
declare
  p record;
begin
  for p in
    select policyname
    from pg_policies
    where schemaname = 'public'
      and tablename = 'review_import_runs'
  loop
    execute format('drop policy if exists %I on public.review_import_runs', p.policyname);
  end loop;
end
$$;

do $$
declare
  p record;
begin
  for p in
    select policyname
    from pg_policies
    where schemaname = 'public'
      and tablename = 'review_import_staging'
  loop
    execute format('drop policy if exists %I on public.review_import_staging', p.policyname);
  end loop;
end
$$;

do $$
declare
  p record;
begin
  for p in
    select policyname
    from pg_policies
    where schemaname = 'public'
      and tablename = 'reviews'
  loop
    execute format('drop policy if exists %I on public.reviews', p.policyname);
  end loop;
end
$$;

do $$
declare
  p record;
begin
  for p in
    select policyname
    from pg_policies
    where schemaname = 'public'
      and tablename = 'app_settings'
  loop
    execute format('drop policy if exists %I on public.app_settings', p.policyname);
  end loop;
end
$$;

do $$
declare
  p record;
begin
  for p in
    select policyname
    from pg_policies
    where schemaname = 'public'
      and tablename = 'app_profiles'
  loop
    execute format('drop policy if exists %I on public.app_profiles', p.policyname);
  end loop;
end
$$;

do $$
declare
  p record;
begin
  for p in
    select policyname
    from pg_policies
    where schemaname = 'public'
      and tablename = 'user_profile_assignments'
  loop
    execute format('drop policy if exists %I on public.user_profile_assignments', p.policyname);
  end loop;
end
$$;

-- Ensure authenticated users can access the table at the SQL permission layer
grant usage on schema public to authenticated;
grant select, insert, update, delete on table public.communications to authenticated;
grant select, insert, update, delete on table public.lost_found to authenticated;
grant select, insert, update, delete on table public.app_settings to authenticated;
grant select, insert, update, delete on table public.app_profiles to authenticated;
grant select, insert, update, delete on table public.user_profile_assignments to authenticated;
grant select, insert, update, delete on table public.group_proposals to authenticated;
grant select, insert, update, delete on table public.services to authenticated;
grant select, insert, update, delete on table public.properties to authenticated;
grant select, insert, update, delete on table public.review_import_runs to authenticated;
grant select, insert, update, delete on table public.review_import_staging to authenticated;
grant select, insert, update, delete on table public.reviews to authenticated;
grant usage on schema public to anon;
grant select, insert, update, delete on table public.communications to anon;
grant select, insert, update, delete on table public.lost_found to anon;
grant select, insert, update, delete on table public.app_settings to anon;
grant select, insert, update, delete on table public.app_profiles to anon;
grant select, insert, update, delete on table public.user_profile_assignments to anon;
grant select, insert, update, delete on table public.group_proposals to anon;
grant select, insert, update, delete on table public.services to anon;
grant select, insert, update, delete on table public.properties to anon;
grant select, insert, update, delete on table public.review_import_runs to anon;
grant select, insert, update, delete on table public.review_import_staging to anon;
grant select, insert, update, delete on table public.reviews to anon;

-- Authenticated-only access
create policy "communications_select_authenticated"
on public.communications
for select
to public
using (auth.uid() is not null);

create policy "communications_insert_authenticated"
on public.communications
for insert
to public
with check (auth.uid() is not null);

create policy "communications_update_authenticated"
on public.communications
for update
to public
using (auth.uid() is not null)
with check (auth.uid() is not null);

create policy "communications_delete_authenticated"
on public.communications
for delete
to public
using (auth.uid() is not null);

drop policy if exists "lost_found_select_authenticated" on public.lost_found;
drop policy if exists "lost_found_insert_authenticated" on public.lost_found;
drop policy if exists "lost_found_update_authenticated" on public.lost_found;
drop policy if exists "lost_found_delete_authenticated" on public.lost_found;

create policy "lost_found_select_authenticated"
on public.lost_found
for select
to public
using (auth.uid() is not null);

create policy "lost_found_insert_authenticated"
on public.lost_found
for insert
to public
with check (auth.uid() is not null);

create policy "lost_found_update_authenticated"
on public.lost_found
for update
to public
using (auth.uid() is not null)
with check (auth.uid() is not null);

create policy "lost_found_delete_authenticated"
on public.lost_found
for delete
to public
using (auth.uid() is not null);

create policy "services_select_authenticated"
on public.services
for select
to public
using (auth.uid() is not null);

create policy "services_insert_authenticated"
on public.services
for insert
to public
with check (auth.uid() is not null);

create policy "services_update_authenticated"
on public.services
for update
to public
using (auth.uid() is not null)
with check (auth.uid() is not null);

create policy "services_delete_authenticated"
on public.services
for delete
to public
using (auth.uid() is not null);

create policy "app_settings_select_authenticated"
on public.app_settings
for select
to public
using (auth.uid() is not null);

create policy "app_settings_insert_authenticated"
on public.app_settings
for insert
to public
with check (auth.uid() is not null);

create policy "app_settings_update_authenticated"
on public.app_settings
for update
to public
using (auth.uid() is not null)
with check (auth.uid() is not null);

create policy "app_settings_delete_authenticated"
on public.app_settings
for delete
to public
using (auth.uid() is not null);

create policy "app_profiles_select_authenticated"
on public.app_profiles
for select
to public
using (auth.uid() is not null);

create policy "app_profiles_insert_authenticated"
on public.app_profiles
for insert
to public
with check (auth.uid() is not null);

create policy "app_profiles_update_authenticated"
on public.app_profiles
for update
to public
using (auth.uid() is not null)
with check (auth.uid() is not null);

create policy "app_profiles_delete_authenticated"
on public.app_profiles
for delete
to public
using (auth.uid() is not null);

create policy "user_profile_assignments_select_authenticated"
on public.user_profile_assignments
for select
to public
using (auth.uid() is not null);

create policy "user_profile_assignments_insert_authenticated"
on public.user_profile_assignments
for insert
to public
with check (auth.uid() is not null);

create policy "user_profile_assignments_update_authenticated"
on public.user_profile_assignments
for update
to public
using (auth.uid() is not null)
with check (auth.uid() is not null);

create policy "user_profile_assignments_delete_authenticated"
on public.user_profile_assignments
for delete
to public
using (auth.uid() is not null);

create policy "group_proposals_select_authenticated"
on public.group_proposals
for select
to public
using (auth.uid() is not null);

create policy "group_proposals_insert_authenticated"
on public.group_proposals
for insert
to public
with check (auth.uid() is not null);

create policy "group_proposals_update_authenticated"
on public.group_proposals
for update
to public
using (auth.uid() is not null)
with check (auth.uid() is not null);

create policy "group_proposals_delete_authenticated"
on public.group_proposals
for delete
to public
using (auth.uid() is not null);

create policy "properties_select_authenticated"
on public.properties
for select
to public
using (auth.uid() is not null);

create policy "properties_insert_authenticated"
on public.properties
for insert
to public
with check (auth.uid() is not null);

create policy "properties_update_authenticated"
on public.properties
for update
to public
using (auth.uid() is not null)
with check (auth.uid() is not null);

create policy "properties_delete_authenticated"
on public.properties
for delete
to public
using (auth.uid() is not null);

create policy "review_import_runs_select_authenticated"
on public.review_import_runs
for select
to public
using (auth.uid() is not null);

create policy "review_import_runs_insert_authenticated"
on public.review_import_runs
for insert
to public
with check (auth.uid() is not null);

create policy "review_import_runs_update_authenticated"
on public.review_import_runs
for update
to public
using (auth.uid() is not null)
with check (auth.uid() is not null);

create policy "review_import_runs_delete_authenticated"
on public.review_import_runs
for delete
to public
using (auth.uid() is not null);

create policy "review_import_staging_select_authenticated"
on public.review_import_staging
for select
to public
using (auth.uid() is not null);

create policy "review_import_staging_insert_authenticated"
on public.review_import_staging
for insert
to public
with check (auth.uid() is not null);

create policy "review_import_staging_update_authenticated"
on public.review_import_staging
for update
to public
using (auth.uid() is not null)
with check (auth.uid() is not null);

create policy "review_import_staging_delete_authenticated"
on public.review_import_staging
for delete
to public
using (auth.uid() is not null);

create policy "reviews_select_authenticated"
on public.reviews
for select
to public
using (auth.uid() is not null);

create policy "reviews_insert_authenticated"
on public.reviews
for insert
to public
with check (auth.uid() is not null);

create policy "reviews_update_authenticated"
on public.reviews
for update
to public
using (auth.uid() is not null)
with check (auth.uid() is not null);

create policy "reviews_delete_authenticated"
on public.reviews
for delete
to public
using (auth.uid() is not null);

insert into public.app_profiles (name, app_features, settings_features)
values ('Administrator', '["communications","lost-found","reviews","groups","services"]'::jsonb, '["communications","reviews","groups","services","admin-users"]'::jsonb)
on conflict (name) do nothing;

update public.app_profiles
set
  app_features = (
    case when app_features ? 'groups' then app_features else app_features || '["groups"]'::jsonb end
  ) || case when app_features ? 'services' then '[]'::jsonb else '["services"]'::jsonb end
    || case when app_features ? 'lost-found' then '[]'::jsonb else '["lost-found"]'::jsonb end,
  settings_features = (case when settings_features ? 'groups' then settings_features else settings_features || '["groups"]'::jsonb end) || case when settings_features ? 'services' then '[]'::jsonb else '["services"]'::jsonb end
where name = 'Administrator';

insert into public.app_profiles (name, app_features, settings_features)
values ('Service Provider', '["services"]'::jsonb, '[]'::jsonb)
on conflict (name) do nothing;

insert into public.properties (name)
values ('Lisboa Central Hostel'), ('Cruz Apartments')
on conflict (name) do nothing;

-- One-time repair for Expedia/Hotels.com imports.
-- Hotels.com is now treated as Expedia; keep the original brand in the review text.
update public.reviews
set
  source = 'expedia',
  body = case
    when coalesce(body, '') ilike '%Brand type:%' then body
    when coalesce(body, '') = '' then 'Brand type: Hotels'
    else trim(body || E'\n\nBrand type: Hotels')
  end
where source = 'hotels';

update public.reviews
set body = case
  when coalesce(body, '') ilike '%Brand type:%' then body
  when coalesce(body, '') = '' then 'Brand type: ' || coalesce(raw_payload->>'brandType', raw_payload->'row'->>1)
  else trim(body || E'\n\nBrand type: ' || coalesce(raw_payload->>'brandType', raw_payload->'row'->>1))
end
where source = 'expedia'
  and coalesce(raw_payload->>'brandType', raw_payload->'row'->>1, '') <> ''
  and coalesce(body, '') not ilike '%Brand type:%';
