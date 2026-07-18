create or replace function public.get_bi_financial_analysis_sales(
  p_year_from integer default extract(year from current_date)::integer - 4,
  p_year_to integer default extract(year from current_date)::integer,
  p_cc text default null
)
returns table (
  type text,
  year integer,
  month integer,
  year_month text,
  cc text,
  category text,
  total_amount numeric
)
language sql
stable
as $$
with filter_bounds as (
  select
    make_date(least(coalesce(p_year_from, extract(year from current_date)::integer - 4), coalesce(p_year_to, extract(year from current_date)::integer)), 1, 1) as date_from,
    make_date(greatest(coalesce(p_year_from, extract(year from current_date)::integer - 4), coalesce(p_year_to, extract(year from current_date)::integer)) + 1, 1, 1) as date_to,
    nullif(upper(btrim(coalesce(p_cc, ''))), '') as filter_cc
),
bi_sale_config as (
  select distinct on (lower(btrim(config.sale_item)))
    config.sale_item,
    config.sale_category
  from (
    select
      coalesce(item ->> 'saleItem', item ->> 'sale_item', item ->> 'SALE_ITEM') as sale_item,
      coalesce(item ->> 'saleCategory', item ->> 'sale_category', item ->> 'SALE_CATEGORY') as sale_category
    from public.app_settings settings
    cross join lateral jsonb_array_elements((settings.payload::jsonb) -> 'saleItemCategories') as item
    where settings.setting_key in ('bi_settings', 'bi-settings')
  ) config
  where nullif(btrim(config.sale_item), '') is not null
  order by lower(btrim(config.sale_item))
),
filtered_sales as materialized (
  select
    sales.sale_date,
    sales.reservation_id,
    sales.sale_item,
    sales.total
  from public.import_fdm_sales sales
  cross join filter_bounds bounds
  where sales.sale_date >= bounds.date_from
    and sales.sale_date < bounds.date_to
)
select
  'EXPENSE'::text as type,
  extract(year from documents.document_date)::integer as year,
  extract(month from documents.document_date)::integer as month,
  to_char(documents.document_date, 'YYYY-MM') as year_month,
  documents.cc::text as cc,
  documents.category::text as category,
  sum(coalesce(documents.amount, 0))::numeric as total_amount
from public.financial_documents documents
cross join filter_bounds bounds
where documents.document_date >= bounds.date_from
  and documents.document_date < bounds.date_to
  and (bounds.filter_cc is null or upper(btrim(documents.cc)) = bounds.filter_cc)
group by
  extract(year from documents.document_date),
  extract(month from documents.document_date),
  to_char(documents.document_date, 'YYYY-MM'),
  documents.cc,
  documents.category

union all

select
  'INCOME'::text as type,
  extract(year from occupancy.mes)::integer as year,
  extract(month from occupancy.mes)::integer as month,
  to_char(occupancy.mes, 'YYYY-MM') as year_month,
  case
    when occupancy.property = 'Hostel' then 'H'
    when occupancy.property = 'Apartamentos' then 'A'
  end as cc,
  'Accomodation'::text as category,
  sum(coalesce(occupancy.charge, 0))::numeric as total_amount
from public.import_fdm_occupancy_kpi occupancy
cross join filter_bounds bounds
where occupancy.mes >= bounds.date_from
  and occupancy.mes < bounds.date_to
  and occupancy.property in ('Hostel', 'Apartamentos')
  and (
    bounds.filter_cc is null
    or (bounds.filter_cc = 'H' and occupancy.property = 'Hostel')
    or (bounds.filter_cc = 'A' and occupancy.property = 'Apartamentos')
  )
group by
  extract(year from occupancy.mes),
  extract(month from occupancy.mes),
  to_char(occupancy.mes, 'YYYY-MM'),
  case
    when occupancy.property = 'Hostel' then 'H'
    when occupancy.property = 'Apartamentos' then 'A'
  end

union all

select
  'INCOME'::text as type,
  extract(year from sales.sale_date)::integer as year,
  extract(month from sales.sale_date)::integer as month,
  to_char(sales.sale_date, 'YYYY-MM') as year_month,
  coalesce(booking.cc, 'H') as cc,
  coalesce(config.sale_category, '') as category,
  sum(coalesce(sales.total, 0))::numeric as total_amount
from filtered_sales sales
cross join filter_bounds bounds
left join bi_sale_config config
  on lower(btrim(config.sale_item)) = lower(btrim(sales.sale_item))
left join lateral (
  select case when booking_row.room_type ilike '%Cruz%' then 'A' else 'H' end as cc
  from public.import_fdm_bookings booking_row
  where booking_row.booking_number = sales.reservation_id
  order by booking_row.created_at desc
  limit 1
) booking on true
where bounds.filter_cc is null or coalesce(booking.cc, 'H') = bounds.filter_cc
group by
  extract(year from sales.sale_date),
  extract(month from sales.sale_date),
  to_char(sales.sale_date, 'YYYY-MM'),
  coalesce(booking.cc, 'H'),
  coalesce(config.sale_category, '');
$$;

create or replace function public.get_bi_financial_analysis_sales_payload(
  p_year_from integer default extract(year from current_date)::integer - 4,
  p_year_to integer default extract(year from current_date)::integer,
  p_cc text default null
)
returns table (rows jsonb)
language sql
stable
as $$
select coalesce(
  jsonb_agg(
    to_jsonb(aggregate_row)
    order by aggregate_row.year, aggregate_row.month, aggregate_row.type, aggregate_row.category, aggregate_row.cc
  ),
  '[]'::jsonb
) as rows
from public.get_bi_financial_analysis_sales(p_year_from, p_year_to, p_cc) aggregate_row;
$$;

grant execute on function public.get_bi_financial_analysis_sales(integer, integer, text)
  to anon, authenticated, service_role;
grant execute on function public.get_bi_financial_analysis_sales_payload(integer, integer, text)
  to anon, authenticated, service_role;

create index if not exists financial_documents_document_date_cc_idx
  on public.financial_documents (document_date, cc);
create index if not exists import_fdm_occupancy_kpi_mes_property_idx
  on public.import_fdm_occupancy_kpi (mes, property);
create index if not exists import_fdm_sales_sale_date_reservation_id_idx
  on public.import_fdm_sales (sale_date, reservation_id);
create index if not exists import_fdm_bookings_booking_number_created_at_idx
  on public.import_fdm_bookings (booking_number, created_at desc);
