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

create table if not exists public.financial_documents (
  id uuid primary key default gen_random_uuid(),
  cc text not null default '',
  document_date date not null,
  doc_number text not null default '',
  description text not null default '',
  supplier_nif text not null default '',
  supplier_name text not null default '',
  amount numeric(14,2) not null default 0,
  vat_amount numeric(14,2),
  payment text not null default '',
  document_type text not null default '',
  fat text not null default '',
  category text not null default '',
  status text not null default 'Draft',
  drive_file_id text not null default '',
  drive_folder_id text not null default '',
  drive_file_url text not null default '',
  original_filename text not null default '',
  stored_filename text not null default '',
  mime_type text not null default '',
  file_size bigint not null default 0,
  file_hash text not null default '',
  uploaded_by text not null default '',
  uploaded_at timestamptz,
  ocr_fields jsonb not null default '{}'::jsonb,
  ocr_raw_text text not null default '',
  created_by text not null default '',
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now())
);

create table if not exists public.financial_document_history (
  id uuid primary key default gen_random_uuid(),
  document_id uuid not null references public.financial_documents(id) on delete cascade,
  action_type text not null default '',
  field_name text not null default '',
  message text not null default '',
  old_value jsonb,
  new_value jsonb,
  metadata jsonb not null default '{}'::jsonb,
  created_by text not null default '',
  created_at timestamptz not null default timezone('utc', now())
);

create index if not exists financial_documents_created_at_idx
  on public.financial_documents (created_at desc);

create index if not exists financial_documents_document_date_idx
  on public.financial_documents (document_date desc);

create index if not exists financial_documents_supplier_nif_idx
  on public.financial_documents (supplier_nif);

create index if not exists financial_documents_doc_number_idx
  on public.financial_documents (doc_number);

create index if not exists financial_documents_file_hash_idx
  on public.financial_documents (file_hash);

create index if not exists financial_document_history_document_created_idx
  on public.financial_document_history (document_id, created_at desc);

drop trigger if exists financial_documents_set_updated_at on public.financial_documents;
create trigger financial_documents_set_updated_at
before update on public.financial_documents
for each row execute procedure public.set_updated_at();

alter table public.financial_documents enable row level security;
alter table public.financial_document_history enable row level security;

drop policy if exists "financial_documents authenticated select" on public.financial_documents;
create policy "financial_documents authenticated select"
on public.financial_documents
for select
to authenticated
using (true);

drop policy if exists "financial_documents authenticated insert" on public.financial_documents;
create policy "financial_documents authenticated insert"
on public.financial_documents
for insert
to authenticated
with check (true);

drop policy if exists "financial_documents authenticated update" on public.financial_documents;
create policy "financial_documents authenticated update"
on public.financial_documents
for update
to authenticated
using (true)
with check (true);

drop policy if exists "financial_documents authenticated delete" on public.financial_documents;
create policy "financial_documents authenticated delete"
on public.financial_documents
for delete
to authenticated
using (true);

drop policy if exists "financial_document_history authenticated select" on public.financial_document_history;
create policy "financial_document_history authenticated select"
on public.financial_document_history
for select
to authenticated
using (true);

drop policy if exists "financial_document_history authenticated insert" on public.financial_document_history;
create policy "financial_document_history authenticated insert"
on public.financial_document_history
for insert
to authenticated
with check (true);

grant select, insert, update, delete on public.financial_documents to authenticated;
grant select, insert, update, delete on public.financial_documents to anon;
grant select, insert, update, delete on public.financial_documents to service_role;

grant select, insert on public.financial_document_history to authenticated;
grant select, insert on public.financial_document_history to anon;
grant select, insert on public.financial_document_history to service_role;
