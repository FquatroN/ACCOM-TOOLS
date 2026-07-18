create or replace view public.bi_financial_analysis_sales as
with bi_sale_config as (
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
  'EXPENSE' as type,
  extract(year from document_date)::int as year,
  extract(month from document_date)::int as month,
  to_char(document_date, 'YYYY-MM') as year_month,
  cc,
  category,
  sum(coalesce(amount, 0)) as total_amount
from public.financial_documents
where document_date is not null
group by
  extract(year from document_date),
  extract(month from document_date),
  to_char(document_date, 'YYYY-MM'),
  cc,
  category

union all

select
  'INCOME' as type,
  extract(year from mes)::int as year,
  extract(month from mes)::int as month,
  to_char(mes, 'YYYY-MM') as year_month,
  case
    when property = 'Hostel' then 'H'
    when property = 'Apartamentos' then 'A'
  end as cc,
  'Accomodation' as category,
  sum(coalesce(charge, 0)) as total_amount
from public.import_fdm_occupancy_kpi
where mes is not null
  and property in ('Hostel', 'Apartamentos')
group by
  extract(year from mes),
  extract(month from mes),
  to_char(mes, 'YYYY-MM'),
  case
    when property = 'Hostel' then 'H'
    when property = 'Apartamentos' then 'A'
  end

union all

select
  'INCOME' as type,
  extract(year from s.sale_date)::int as year,
  extract(month from s.sale_date)::int as month,
  to_char(s.sale_date, 'YYYY-MM') as year_month,
  coalesce(b.cc, 'H') as cc,
  coalesce(c.sale_category, '') as category,
  sum(coalesce(s.total, 0)) as total_amount
from public.import_fdm_sales s
left join bi_sale_config c
  on lower(btrim(c.sale_item)) = lower(btrim(s.sale_item))
left join bookings_by_reservation b
  on b.booking_number = s.reservation_id
where s.sale_date is not null
group by
  extract(year from s.sale_date),
  extract(month from s.sale_date),
  to_char(s.sale_date, 'YYYY-MM'),
  coalesce(b.cc, 'H'),
  coalesce(c.sale_category, '');
