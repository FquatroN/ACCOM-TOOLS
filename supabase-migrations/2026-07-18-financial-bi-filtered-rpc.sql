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
  select
    coalesce(item ->> 'saleItem', item ->> 'sale_item', item ->> 'SALE_ITEM') as sale_item,
    coalesce(item ->> 'saleCategory', item ->> 'sale_category', item ->> 'SALE_CATEGORY') as sale_category
  from public.app_settings s
  cross join lateral jsonb_array_elements((s.payload::jsonb) -> 'saleItemCategories') as item
  where s.setting_key in ('bi_settings', 'bi-settings')
),
bookings_by_reservation as (
  select distinct on (booking_number)
    booking_number,
    case when room_type ilike '%Cruz%' then 'A' else 'H' end as cc
  from public.import_fdm_bookings
  where coalesce(booking_number, '') <> ''
  order by booking_number, created_at desc
)
select
  'EXPENSE'::text as type,
  extract(year from fd.document_date)::integer as year,
  extract(month from fd.document_date)::integer as month,
  to_char(fd.document_date, 'YYYY-MM') as year_month,
  fd.cc::text as cc,
  fd.category::text as category,
  sum(coalesce(fd.amount, 0))::numeric as total_amount
from public.financial_documents fd
cross join filter_bounds fb
where fd.document_date >= fb.date_from
  and fd.document_date < fb.date_to
  and (fb.filter_cc is null or upper(btrim(fd.cc)) = fb.filter_cc)
group by
  extract(year from fd.document_date),
  extract(month from fd.document_date),
  to_char(fd.document_date, 'YYYY-MM'),
  fd.cc,
  fd.category

union all

select
  'INCOME'::text as type,
  extract(year from kpi.mes)::integer as year,
  extract(month from kpi.mes)::integer as month,
  to_char(kpi.mes, 'YYYY-MM') as year_month,
  case
    when kpi.property = 'Hostel' then 'H'
    when kpi.property = 'Apartamentos' then 'A'
  end as cc,
  'Accomodation'::text as category,
  sum(coalesce(kpi.charge, 0))::numeric as total_amount
from public.import_fdm_occupancy_kpi kpi
cross join filter_bounds fb
where kpi.mes >= fb.date_from
  and kpi.mes < fb.date_to
  and kpi.property in ('Hostel', 'Apartamentos')
  and (
    fb.filter_cc is null
    or (fb.filter_cc = 'H' and kpi.property = 'Hostel')
    or (fb.filter_cc = 'A' and kpi.property = 'Apartamentos')
  )
group by
  extract(year from kpi.mes),
  extract(month from kpi.mes),
  to_char(kpi.mes, 'YYYY-MM'),
  case
    when kpi.property = 'Hostel' then 'H'
    when kpi.property = 'Apartamentos' then 'A'
  end

union all

select
  'INCOME'::text as type,
  extract(year from s.sale_date)::integer as year,
  extract(month from s.sale_date)::integer as month,
  to_char(s.sale_date, 'YYYY-MM') as year_month,
  coalesce(b.cc, 'H') as cc,
  coalesce(c.sale_category, '') as category,
  sum(coalesce(s.total, 0))::numeric as total_amount
from public.import_fdm_sales s
cross join filter_bounds fb
left join bi_sale_config c
  on lower(btrim(c.sale_item)) = lower(btrim(s.sale_item))
left join bookings_by_reservation b
  on b.booking_number = s.reservation_id
where s.sale_date >= fb.date_from
  and s.sale_date < fb.date_to
  and (fb.filter_cc is null or coalesce(b.cc, 'H') = fb.filter_cc)
group by
  extract(year from s.sale_date),
  extract(month from s.sale_date),
  to_char(s.sale_date, 'YYYY-MM'),
  coalesce(b.cc, 'H'),
  coalesce(c.sale_category, '');
$$;

grant execute on function public.get_bi_financial_analysis_sales(integer, integer, text) to anon, authenticated, service_role;

create index if not exists financial_documents_document_date_cc_idx
  on public.financial_documents (document_date, cc);

create index if not exists import_fdm_occupancy_kpi_mes_property_idx
  on public.import_fdm_occupancy_kpi (mes, property);

create index if not exists import_fdm_sales_sale_date_reservation_id_idx
  on public.import_fdm_sales (sale_date, reservation_id);

create index if not exists import_fdm_bookings_booking_number_created_at_idx
  on public.import_fdm_bookings (booking_number, created_at desc);
