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

create table if not exists public.financial_document_entities (
  id uuid primary key default gen_random_uuid(),
  nif text not null default '',
  name text not null default '',
  address text not null default '',
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now()),
  constraint financial_document_entities_nif_required check (regexp_replace(btrim(nif), '\D', '', 'g') <> ''),
  constraint financial_document_entities_name_required check (btrim(name) <> '')
);

create unique index if not exists financial_document_entities_nif_norm_uidx
  on public.financial_document_entities ((regexp_replace(btrim(nif), '\D', '', 'g')));

create unique index if not exists financial_document_entities_name_norm_uidx
  on public.financial_document_entities ((lower(regexp_replace(btrim(name), '\s+', ' ', 'g'))));

create index if not exists financial_document_entities_name_idx
  on public.financial_document_entities (name);

create index if not exists financial_document_entities_nif_idx
  on public.financial_document_entities (nif);

drop trigger if exists financial_document_entities_set_updated_at on public.financial_document_entities;
create trigger financial_document_entities_set_updated_at
before update on public.financial_document_entities
for each row execute procedure public.set_updated_at();

alter table public.financial_document_entities enable row level security;

drop policy if exists "financial_document_entities authenticated select" on public.financial_document_entities;
create policy "financial_document_entities authenticated select"
on public.financial_document_entities
for select
to authenticated
using (true);

drop policy if exists "financial_document_entities authenticated insert" on public.financial_document_entities;
create policy "financial_document_entities authenticated insert"
on public.financial_document_entities
for insert
to authenticated
with check (true);

drop policy if exists "financial_document_entities authenticated update" on public.financial_document_entities;
create policy "financial_document_entities authenticated update"
on public.financial_document_entities
for update
to authenticated
using (true)
with check (true);

drop policy if exists "financial_document_entities authenticated delete" on public.financial_document_entities;
create policy "financial_document_entities authenticated delete"
on public.financial_document_entities
for delete
to authenticated
using (true);

grant select, insert, update, delete on public.financial_document_entities to authenticated;
grant select, insert, update, delete on public.financial_document_entities to anon;
grant select, insert, update, delete on public.financial_document_entities to service_role;
