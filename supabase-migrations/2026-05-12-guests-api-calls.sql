create extension if not exists pgcrypto;

create table if not exists public.guest_api_calls (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default timezone('utc', now()),
  endpoint text not null default '',
  request_method text not null default 'POST',
  soap_action text not null default '',
  http_status integer not null default 0,
  file_number integer not null default 0,
  guest_count integer not null default 0,
  success boolean not null default false,
  response_message text not null default '',
  error_message text not null default '',
  request_details jsonb not null default '{}'::jsonb,
  request_body text not null default '',
  response_body text not null default ''
);

create index if not exists guest_api_calls_created_at_idx
  on public.guest_api_calls (created_at desc);

alter table public.guest_api_calls enable row level security;

drop policy if exists "guest_api_calls authenticated select" on public.guest_api_calls;
create policy "guest_api_calls authenticated select"
on public.guest_api_calls
for select
to authenticated
using (true);

drop policy if exists "guest_api_calls authenticated insert" on public.guest_api_calls;
create policy "guest_api_calls authenticated insert"
on public.guest_api_calls
for insert
to authenticated
with check (true);

grant select, insert on public.guest_api_calls to authenticated;
grant select, insert on public.guest_api_calls to anon;
