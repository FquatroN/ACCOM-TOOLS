-- Step 3 for the large guest history migration.
-- Builds cleaned candidate views from public.guest_records_import_csv_raw.
--
-- Business rules applied here:
-- - final guest_records rows will be imported with sent_status = 'sent'
-- - rows with blocker issues are rejected
-- - exact duplicate rows are collapsed
-- - duplicate (doc_number, check_in) groups keep a single best row
-- - document numbers that look like scientific notation are kept as-is
--   for now, but remain visible in the scientific-number validation view

drop view if exists public.guest_records_import_csv_rejected_rows;
drop view if exists public.guest_records_import_csv_clean_base;
drop view if exists public.guest_records_import_csv_final_candidates;

create view public.guest_records_import_csv_rejected_rows as
select *
from public.guest_records_import_csv_blocker_rows;

create view public.guest_records_import_csv_clean_base as
with base as (
  select
    r.staging_id,
    r.import_batch,
    r.source_file,
    r.loaded_at,
    btrim(r."HA", E' \t\n\r"') as ha_raw,
    btrim(r."Name", E' \t\n\r"') as name_raw,
    btrim(r."Nationality", E' \t\n\r"') as nationality_raw,
    btrim(r."Birth Date", E' \t\n\r"') as birth_date_raw,
    btrim(r."Doc. Number", E' \t\n\r"') as doc_number_raw,
    upper(btrim(r."Doc Type", E' \t\n\r"')) as doc_type_raw,
    btrim(r."Issuer Country", E' \t\n\r"') as issuer_country_raw,
    btrim(r."Check-in", E' \t\n\r"') as check_in_raw,
    btrim(r."Check-out", E' \t\n\r"') as check_out_raw
  from public.guest_records_import_csv_raw r
),
filtered as (
  select b.*
  from base b
  left join public.guest_records_import_csv_blocker_rows x
    on x.staging_id = b.staging_id
  where x.staging_id is null
),
normalized as (
  select
    staging_id,
    import_batch,
    source_file,
    loaded_at,
    case when upper(ha_raw) = 'A' then 'A' else 'H' end as ha,
    name_raw as name,
    nationality_raw as nationality,
    public.guest_country_resolve_code(nationality_raw) as nationality_code,
    case
      when birth_date_raw ~ '^[0-9]{2}/[0-9]{2}/[0-9]{4}$'
        then to_date(birth_date_raw, 'DD/MM/YYYY')
      when birth_date_raw ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
        then to_date(birth_date_raw, 'YYYY-MM-DD')
      else null
    end as birth_date,
    ''::text as birth_place,
    upper(replace(doc_number_raw, ' ', '')) as doc_number,
    case
      when doc_type_raw in ('P', 'B', 'O') then doc_type_raw
      when doc_type_raw = 'I' then 'B'
      else 'O'
    end as doc_type,
    issuer_country_raw as issuer_country,
    public.guest_country_resolve_code(issuer_country_raw) as issuer_country_code,
    ''::text as residence_country,
    ''::text as residence_country_code,
    ''::text as residence_city,
    case
      when check_in_raw ~ '^[0-9]{2}/[0-9]{2}/[0-9]{4}$'
        then to_date(check_in_raw, 'DD/MM/YYYY')
      when check_in_raw ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
        then to_date(check_in_raw, 'YYYY-MM-DD')
      else null
    end as check_in,
    case
      when check_out_raw ~ '^[0-9]{2}/[0-9]{2}/[0-9]{4}$'
        then to_date(check_out_raw, 'DD/MM/YYYY')
      when check_out_raw ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
        then to_date(check_out_raw, 'YYYY-MM-DD')
      else null
    end as check_out,
    'sent'::text as sent_status,
    null::timestamptz as sent_at,
    ''::text as send_error,
    0::integer as send_batch_number
  from filtered
),
scored as (
  select
    *,
    (
      case when name <> '' then 1 else 0 end +
      case when nationality <> '' then 1 else 0 end +
      case when nationality_code <> '' then 1 else 0 end +
      case when issuer_country <> '' then 1 else 0 end +
      case when issuer_country_code <> '' then 1 else 0 end +
      case when residence_country <> '' then 1 else 0 end +
      case when residence_country_code <> '' then 1 else 0 end +
      case when residence_city <> '' then 1 else 0 end +
      case when birth_place <> '' then 1 else 0 end
    ) as completeness_score
  from normalized
)
select *
from scored;

create view public.guest_records_import_csv_final_candidates as
with ranked as (
  select
    c.*,
    row_number() over (
      partition by c.doc_number, c.check_in
      order by
        c.completeness_score desc,
        c.check_out desc nulls last,
        c.staging_id desc
    ) as pick_rank,
    count(*) over (
      partition by c.doc_number, c.check_in
    ) as duplicate_group_size
  from public.guest_records_import_csv_clean_base c
)
select
  gen_random_uuid() as id,
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
  timezone('utc', now()) as created_at,
  timezone('utc', now()) as updated_at,
  import_batch,
  source_file,
  staging_id as kept_staging_id,
  duplicate_group_size,
  completeness_score
from ranked
where pick_rank = 1;

comment on view public.guest_records_import_csv_rejected_rows is
  'Rows rejected from the large guest CSV import because of blocker issues.';

comment on view public.guest_records_import_csv_clean_base is
  'Cleaned guest CSV staging rows after blocker removal, before duplicate-group collapse.';

comment on view public.guest_records_import_csv_final_candidates is
  'Final deduplicated candidate rows ready for upsert into public.guest_records.';
