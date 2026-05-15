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

create table if not exists public.maintenance_logs (
  id uuid primary key default gen_random_uuid(),
  task_id text not null default '',
  task_name text not null default '',
  where_value text not null default '',
  done_date date not null,
  type text not null default '',
  who text not null default '',
  note text not null default '',
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now())
);

create index if not exists maintenance_logs_task_date_idx
  on public.maintenance_logs (task_id, done_date desc);

create index if not exists maintenance_logs_task_where_date_idx
  on public.maintenance_logs (task_id, where_value, done_date desc);

drop trigger if exists maintenance_logs_set_updated_at on public.maintenance_logs;
create trigger maintenance_logs_set_updated_at
before update on public.maintenance_logs
for each row execute procedure public.set_updated_at();

alter table public.maintenance_logs enable row level security;

drop policy if exists "maintenance_logs authenticated select" on public.maintenance_logs;
create policy "maintenance_logs authenticated select"
on public.maintenance_logs
for select
to authenticated
using (true);

drop policy if exists "maintenance_logs authenticated insert" on public.maintenance_logs;
create policy "maintenance_logs authenticated insert"
on public.maintenance_logs
for insert
to authenticated
with check (true);

drop policy if exists "maintenance_logs authenticated update" on public.maintenance_logs;
create policy "maintenance_logs authenticated update"
on public.maintenance_logs
for update
to authenticated
using (true)
with check (true);

drop policy if exists "maintenance_logs authenticated delete" on public.maintenance_logs;
create policy "maintenance_logs authenticated delete"
on public.maintenance_logs
for delete
to authenticated
using (true);

grant select, insert, update, delete on public.maintenance_logs to authenticated;
grant select, insert, update, delete on public.maintenance_logs to anon;
grant select, insert, update, delete on public.maintenance_logs to service_role;
