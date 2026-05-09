create extension if not exists pgcrypto;

create or replace function public.set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

create table if not exists public.cash_control_records (
  id uuid primary key default gen_random_uuid(),
  record_day date not null,
  shift_id text not null,
  shift_name text not null default '',
  name text not null default '',
  denominations jsonb not null default '{}'::jsonb,
  card_pos numeric(12, 2) not null default 0,
  cash_fdm numeric(12, 2) not null default 0,
  card_fdm numeric(12, 2) not null default 0,
  justification text not null default '',
  item_counts jsonb not null default '{}'::jsonb,
  item_justifications jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint cash_control_records_day_shift_unique unique (record_day, shift_id)
);

drop trigger if exists cash_control_records_set_updated_at on public.cash_control_records;
create trigger cash_control_records_set_updated_at
before update on public.cash_control_records
for each row execute function public.set_updated_at();

grant select, insert, update, delete on table public.cash_control_records to authenticated;
grant select, insert, update, delete on table public.cash_control_records to anon;

alter table public.cash_control_records enable row level security;

drop policy if exists "cash_control_records_select_authenticated" on public.cash_control_records;
drop policy if exists "cash_control_records_insert_authenticated" on public.cash_control_records;
drop policy if exists "cash_control_records_update_authenticated" on public.cash_control_records;
drop policy if exists "cash_control_records_delete_authenticated" on public.cash_control_records;

create policy "cash_control_records_select_authenticated"
on public.cash_control_records
for select
to public
using (auth.uid() is not null);

create policy "cash_control_records_insert_authenticated"
on public.cash_control_records
for insert
to public
with check (auth.uid() is not null);

create policy "cash_control_records_update_authenticated"
on public.cash_control_records
for update
to public
using (auth.uid() is not null)
with check (auth.uid() is not null);

create policy "cash_control_records_delete_authenticated"
on public.cash_control_records
for delete
to public
using (auth.uid() is not null);
