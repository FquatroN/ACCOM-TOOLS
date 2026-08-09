create or replace function public.get_bi_sales_pivot_payload(
  p_year_from integer default extract(year from current_date)::integer - 4,
  p_year_to integer default extract(year from current_date)::integer
)
returns table (rows jsonb)
language sql
stable
as $$
with bounds as (
  select
    make_date(
      least(
        coalesce(p_year_from, extract(year from current_date)::integer - 4),
        coalesce(p_year_to, extract(year from current_date)::integer)
      ),
      1,
      1
    ) as date_from,
    least(
      make_date(
        greatest(
          coalesce(p_year_from, extract(year from current_date)::integer - 4),
          coalesce(p_year_to, extract(year from current_date)::integer)
        ) + 1,
        1,
        1
      ),
      date_trunc('month', current_date)::date
    ) as date_to
),
sale_category_config as (
  select distinct on (sale_item_key)
    sale_item_key,
    sale_category
  from (
    select
      lower(btrim(coalesce(item ->> 'saleItem', item ->> 'sale_item', item ->> 'SALE_ITEM', ''))) as sale_item_key,
      btrim(coalesce(item ->> 'saleCategory', item ->> 'sale_category', item ->> 'SALE_CATEGORY', '')) as sale_category
    from public.app_settings settings
    cross join lateral jsonb_array_elements(
      coalesce((settings.payload::jsonb) -> 'saleItemCategories', '[]'::jsonb)
    ) as item
    where settings.setting_key in ('bi_settings', 'bi-settings')
  ) configured_items
  where sale_item_key <> ''
  order by sale_item_key, sale_category
),
aggregated as (
  select
    coalesce(config.sale_category, '') as category,
    btrim(coalesce(sale.sale_item, '')) as sale_item,
    extract(year from sale.sale_date)::integer as year,
    extract(month from sale.sale_date)::integer as month,
    to_char(sale.sale_date, 'YYYY-MM') as year_month,
    sum(coalesce(sale.total, 0))::numeric as total
  from public.import_fdm_sales sale
  cross join bounds
  left join sale_category_config config
    on config.sale_item_key = lower(btrim(coalesce(sale.sale_item, '')))
  where sale.sale_date >= bounds.date_from
    and sale.sale_date < bounds.date_to
    and btrim(coalesce(sale.sale_item, '')) <> ''
  group by 1, 2, 3, 4, 5
)
select coalesce(
  jsonb_agg(to_jsonb(aggregate_row) order by aggregate_row.category, aggregate_row.sale_item, aggregate_row.year, aggregate_row.month),
  '[]'::jsonb
)
from aggregated aggregate_row;
$$;
