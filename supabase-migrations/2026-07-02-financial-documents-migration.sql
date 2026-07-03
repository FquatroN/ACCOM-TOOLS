-- Financial documents migration support for:
--   Movimentos Financeiros - Migration.xlsx
--
-- Suggested flow:
--   1. Export the Excel file to CSV.
--   2. Run this script in Supabase.
--   3. Import the CSV into public.financial_documents_migration_raw.
--      The table uses the exact CSV header names.
--   4. Review:
--        select * from public.financial_documents_migration_validation_summary;
--        select * from public.financial_documents_migration_invalid_rows;
--        select * from public.financial_documents_migration_normalized limit 100;
--   5. Insert:
--        select * from public.financial_documents_migration_insert();

create table if not exists public.financial_documents_migration_raw (
  staging_id bigint generated always as identity primary key,
  "Create Date" text not null default '',
  "CC" text not null default '',
  "Data" text not null default '',
  "Doc Number" text not null default '',
  "Description" text not null default '',
  "Name" text not null default '',
  "Supplier NIF" text not null default '',
  "Amount" text not null default '',
  "VAT Amount" text not null default '',
  "Pagamento" text not null default '',
  "Type" text not null default '',
  "Fat" text not null default '',
  "Categoria" text not null default '',
  "Status" text not null default '',
  import_batch text not null default 'movimentos-financeiros-2026-07-02',
  source_file text not null default 'Movimentos Financeiros - Migration.xlsx',
  loaded_at timestamptz not null default timezone('utc', now())
);

create index if not exists financial_documents_migration_raw_batch_idx
  on public.financial_documents_migration_raw (import_batch);

create index if not exists financial_documents_migration_raw_doc_idx
  on public.financial_documents_migration_raw ("Doc Number");

create or replace function public.financial_docs_migration_clean_text(raw_value text)
returns text
language sql
immutable
as $$
  select btrim(coalesce(raw_value, ''), E' \t\n\r"');
$$;

create or replace function public.financial_docs_migration_normalize_cc(raw_value text)
returns text
language sql
immutable
as $$
  select upper(public.financial_docs_migration_clean_text(raw_value));
$$;

create or replace function public.financial_docs_migration_normalize_payment(raw_value text)
returns text
language sql
immutable
as $$
  select case lower(public.financial_docs_migration_clean_text(raw_value))
    when '' then ''
    when 'banco' then 'Banco'
    when 'caixa' then 'Caixa'
    when 'cash' then 'Cash'
    when 'visa' then 'Visa'
    when 'miguel' then 'Miguel'
    when 'carlos' then 'Carlos'
    when 'odete' then 'Odete'
    else public.financial_docs_migration_clean_text(raw_value)
  end;
$$;

create or replace function public.financial_docs_migration_normalize_type(raw_value text)
returns text
language sql
immutable
as $$
  select upper(public.financial_docs_migration_clean_text(raw_value));
$$;

create or replace function public.financial_docs_migration_normalize_fat(raw_value text)
returns text
language sql
immutable
as $$
  select upper(public.financial_docs_migration_clean_text(raw_value));
$$;

create or replace function public.financial_docs_migration_parse_date(raw_value text)
returns date
language plpgsql
immutable
as $$
declare
  value text := public.financial_docs_migration_clean_text(raw_value);
  parsed date;
begin
  value := regexp_replace(value, '\s+[0-9]{1,2}:[0-9]{2}(:[0-9]{2})?$', '');

  if value = '' or value = '-' then
    return null;
  end if;

  if value ~ '^[0-9]+(\.0+)?$' then
    return date '1899-12-30' + floor(value::numeric)::integer;
  end if;

  if value ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$' then
    parsed := to_date(value, 'YYYY-MM-DD');
    if to_char(parsed, 'YYYY-MM-DD') = value then
      return parsed;
    end if;
  end if;

  if value ~ '^[0-9]{1,2}/[0-9]{1,2}/[0-9]{4}$' then
    parsed := to_date(value, 'DD/MM/YYYY');
    if to_char(parsed, 'DD/MM/YYYY') = value
       or to_char(parsed, 'FMDD/FMMM/YYYY') = value then
      return parsed;
    end if;
  end if;

  if value ~ '^[0-9]{1,2}-[0-9]{1,2}-[0-9]{4}$' then
    parsed := to_date(value, 'DD-MM-YYYY');
    if to_char(parsed, 'DD-MM-YYYY') = value
       or to_char(parsed, 'FMDD-FMMM-YYYY') = value then
      return parsed;
    end if;
  end if;

  return null;
exception
  when others then
    return null;
end;
$$;

create or replace function public.financial_docs_migration_parse_numeric(raw_value text)
returns numeric
language plpgsql
immutable
as $$
declare
  value text := public.financial_docs_migration_clean_text(raw_value);
begin
  value := replace(value, chr(160), '');
  value := regexp_replace(value, '\s+', '', 'g');
  value := replace(value, chr(8364), '');
  value := regexp_replace(value, '[^0-9,.\-]', '', 'g');

  if value = '' then
    return null;
  end if;

  if value = '-' or (value !~ '[0-9]' and value like '%-%') then
    return 0;
  end if;

  if value ~ '^-?[0-9]{1,3}(\.[0-9]{3})+,[0-9]+$'
     or value ~ '^-?[0-9]+,[0-9]+$' then
    value := replace(value, '.', '');
    value := replace(value, ',', '.');
  elsif value ~ '^-?[0-9]{1,3}(,[0-9]{3})+(\.[0-9]+)?$' then
    value := replace(value, ',', '');
  else
    value := replace(value, ',', '.');
  end if;

  return value::numeric;
exception
  when others then
    return null;
end;
$$;

drop view if exists public.financial_documents_migration_invalid_rows;
drop view if exists public.financial_documents_migration_validation_summary;
drop view if exists public.financial_documents_migration_normalized;

create view public.financial_documents_migration_normalized as
select
  staging_id,
  import_batch,
  source_file,
  loaded_at,
  public.financial_docs_migration_parse_date("Create Date") as create_date,
  public.financial_docs_migration_normalize_cc("CC") as cc,
  public.financial_docs_migration_parse_date("Data") as document_date,
  public.financial_docs_migration_clean_text("Doc Number") as doc_number,
  public.financial_docs_migration_clean_text("Description") as description,
  public.financial_docs_migration_clean_text("Name") as supplier_name,
  public.financial_docs_migration_clean_text("Supplier NIF") as supplier_nif,
  public.financial_docs_migration_parse_numeric("Amount") as amount,
  public.financial_docs_migration_parse_numeric("VAT Amount") as vat_amount,
  public.financial_docs_migration_normalize_payment("Pagamento") as payment,
  public.financial_docs_migration_normalize_type("Type") as document_type,
  public.financial_docs_migration_normalize_fat("Fat") as fat,
  public.financial_docs_migration_clean_text("Categoria") as category,
  public.financial_docs_migration_clean_text("Status") as status,
  public.financial_docs_migration_clean_text("Create Date") as create_date_raw,
  public.financial_docs_migration_clean_text("Data") as document_date_raw,
  public.financial_docs_migration_clean_text("Amount") as amount_raw,
  public.financial_docs_migration_clean_text("VAT Amount") as vat_amount_raw,
  public.financial_docs_migration_clean_text("CC") as cc_raw,
  public.financial_docs_migration_clean_text("Pagamento") as payment_raw,
  public.financial_docs_migration_clean_text("Type") as document_type_raw,
  public.financial_docs_migration_clean_text("Fat") as fat_raw
from public.financial_documents_migration_raw;

create view public.financial_documents_migration_invalid_rows as
select
  *,
  array_remove(array[
    case when create_date is null then 'invalid_create_date' end,
    case when document_date is null then 'invalid_document_date' end,
    case when cc not in ('A', 'H') then 'invalid_cc' end,
    case when description = '' then 'missing_description' end,
    case when supplier_name = '' then 'missing_name' end,
    case when amount is null then 'invalid_amount' end,
    case when payment = '' then 'missing_payment' end,
    case when document_type not in ('R', 'F') then 'invalid_type' end,
    case when fat not in ('S', 'N') then 'invalid_fat' end,
    case when category = '' then 'missing_category' end,
    case when status = '' then 'missing_status' end
  ], null) as validation_errors
from public.financial_documents_migration_normalized
where create_date is null
   or document_date is null
   or cc not in ('A', 'H')
   or description = ''
   or supplier_name = ''
   or amount is null
   or payment = ''
   or document_type not in ('R', 'F')
   or fat not in ('S', 'N')
   or category = ''
   or status = '';

create view public.financial_documents_migration_validation_summary as
select
  import_batch,
  count(*) as total_rows,
  count(*) filter (where create_date is null) as invalid_create_date,
  count(*) filter (where document_date is null) as invalid_document_date,
  count(*) filter (where cc not in ('A', 'H')) as invalid_cc,
  count(*) filter (where description = '') as missing_description,
  count(*) filter (where supplier_name = '') as missing_name,
  count(*) filter (where amount is null) as invalid_amount,
  count(*) filter (where payment = '') as missing_payment,
  count(*) filter (where document_type not in ('R', 'F')) as invalid_type,
  count(*) filter (where fat not in ('S', 'N')) as invalid_fat,
  count(*) filter (where category = '') as missing_category,
  count(*) filter (where status = '') as missing_status,
  count(*) filter (where doc_number = '') as blank_doc_number_accepted,
  count(*) filter (where supplier_nif = '') as blank_supplier_nif_accepted,
  count(*) filter (where vat_amount is null and vat_amount_raw = '') as blank_vat_amount_accepted,
  count(*) filter (
    where create_date is not null
      and document_date is not null
      and cc in ('A', 'H')
      and description <> ''
      and supplier_name <> ''
      and amount is not null
      and payment <> ''
      and document_type in ('R', 'F')
      and fat in ('S', 'N')
      and category <> ''
      and status <> ''
  ) as valid_rows
from public.financial_documents_migration_normalized
group by import_batch;

create or replace function public.financial_documents_migration_insert(
  p_import_batch text default 'movimentos-financeiros-2026-07-02',
  p_created_by text default 'migration'
)
returns table(inserted_rows integer, skipped_existing_rows integer, invalid_rows integer)
language plpgsql
as $$
declare
  inserted_count integer := 0;
  skipped_count integer := 0;
  invalid_count integer := 0;
begin
  select count(*)
    into invalid_count
  from public.financial_documents_migration_invalid_rows
  where import_batch = p_import_batch;

  if invalid_count > 0 then
    raise exception 'Import batch % has % invalid row(s). Review public.financial_documents_migration_invalid_rows first.',
      p_import_batch,
      invalid_count;
  end if;

  select count(*)
    into skipped_count
  from public.financial_documents_migration_normalized n
  where n.import_batch = p_import_batch
    and exists (
      select 1
      from public.financial_documents fd
      where fd.ocr_fields->>'migration_batch' = n.import_batch
        and fd.ocr_fields->>'source_staging_id' = n.staging_id::text
    );

  insert into public.financial_documents (
    cc,
    document_date,
    doc_number,
    description,
    supplier_nif,
    supplier_name,
    amount,
    vat_amount,
    payment,
    document_type,
    fat,
    category,
    status,
    ocr_fields,
    created_by,
    created_at,
    updated_at
  )
  select
    n.cc,
    n.document_date,
    n.doc_number,
    n.description,
    n.supplier_nif,
    n.supplier_name,
    round(n.amount, 2),
    case when n.vat_amount is null then null else round(n.vat_amount, 2) end,
    n.payment,
    n.document_type,
    n.fat,
    n.category,
    n.status,
    jsonb_build_object(
      'migration_batch', n.import_batch,
      'source_file', n.source_file,
      'source_staging_id', n.staging_id,
      'source_create_date_raw', n.create_date_raw,
      'source_document_date_raw', n.document_date_raw,
      'source_amount_raw', n.amount_raw,
      'source_vat_amount_raw', n.vat_amount_raw
    ),
    p_created_by,
    n.create_date::timestamptz,
    timezone('utc', now())
  from public.financial_documents_migration_normalized n
  where n.import_batch = p_import_batch
    and not exists (
      select 1
      from public.financial_documents fd
      where fd.ocr_fields->>'migration_batch' = n.import_batch
        and fd.ocr_fields->>'source_staging_id' = n.staging_id::text
    );

  get diagnostics inserted_count = row_count;

  return query select inserted_count, skipped_count, invalid_count;
end;
$$;

comment on table public.financial_documents_migration_raw is
  'Raw staging table for Movimentos Financeiros migration. Uses exact CSV headers and normalizes CC, Pagamento, Type, and Fat through the migration views/functions.';
