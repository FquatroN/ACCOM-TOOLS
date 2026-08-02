drop function if exists public.get_bi_financial_items_revenue_payload(integer, integer);
drop function if exists public.get_bi_financial_items_revenue(integer, integer);

create function public.get_bi_financial_items_revenue(
  p_year_from integer default extract(year from current_date)::integer - 5,
  p_year_to integer default extract(year from current_date)::integer
)
returns table (
  analysis text,
  year integer,
  month integer,
  year_month text,
  sales numeric,
  buy numeric,
  revenue numeric
)
language sql
stable
as $$
with bounds as (
  select
    make_date(least(coalesce(p_year_from, extract(year from current_date)::integer - 5), coalesce(p_year_to, extract(year from current_date)::integer)), 1, 1) as date_from,
    make_date(greatest(coalesce(p_year_from, extract(year from current_date)::integer - 5), coalesce(p_year_to, extract(year from current_date)::integer)) + 1, 1, 1) as date_to
),
analyses(analysis) as (
  values ('nespresso'), ('nescafe_machine'), ('souvenirs')
),
months as (
  select analyses.analysis, month_start::date as month_start
  from bounds
  cross join analyses
  cross join lateral generate_series(bounds.date_from, bounds.date_to - interval '1 month', interval '1 month') as month_start
),
sales_by_month as (
  select
    'nespresso'::text as analysis,
    date_trunc('month', sale.sale_date)::date as month_start,
    sum(coalesce(sale.total, 0))::numeric as sales
  from public.import_fdm_sales sale
  cross join bounds
  where sale.sale_date >= bounds.date_from
    and sale.sale_date < bounds.date_to
    and lower(btrim(coalesce(sale.sale_item, ''))) = 'nespresso'
  group by 2

  union all

  select
    'nescafe_machine'::text as analysis,
    date_trunc('month', account.event_date)::date as month_start,
    sum(coalesce(account.amount, 0))::numeric as sales
  from public.import_fdm_accounts account
  cross join bounds
  where account.event_date >= bounds.date_from
    and account.event_date < bounds.date_to
    and lower(btrim(coalesce(account.category, ''))) = 'coffee machine'
  group by 2

  union all

  select
    'souvenirs'::text as analysis,
    date_trunc('month', sale.sale_date)::date as month_start,
    sum(coalesce(sale.total, 0))::numeric as sales
  from public.import_fdm_sales sale
  cross join bounds
  where sale.sale_date >= bounds.date_from
    and sale.sale_date < bounds.date_to
    and lower(coalesce(sale.sale_item, '')) like '%souvenir%'
  group by 2
),
buy_by_month as (
  select
    'nespresso'::text as analysis,
    date_trunc('month', document.document_date)::date as month_start,
    sum(coalesce(document.amount, 0))::numeric as buy
  from public.financial_documents document
  cross join bounds
  where document.document_date >= bounds.date_from
    and document.document_date < bounds.date_to
    and upper(regexp_replace(coalesce(document.supplier_nif, ''), E'\\s+', '', 'g')) in ('500201307', 'PT500201307')
    and (
      document.description ilike '%capsulas%'
      or document.description ilike U&'%c\00E1psulas%'
      or document.description ilike '%nespresso%'
    )
  group by 2

  union all

  select
    'nescafe_machine'::text as analysis,
    date_trunc('month', document.document_date)::date as month_start,
    sum(coalesce(document.amount, 0))::numeric as buy
  from public.financial_documents document
  cross join bounds
  where document.document_date >= bounds.date_from
    and document.document_date < bounds.date_to
    and upper(regexp_replace(coalesce(document.supplier_nif, ''), E'\\s+', '', 'g')) in ('500201307', 'PT500201307')
    and coalesce(document.description, '') not ilike '%capsulas%'
    and coalesce(document.description, '') not ilike U&'%c\00E1psulas%'
    and coalesce(document.description, '') not ilike '%nespresso%'
  group by 2

  union all

  select
    'souvenirs'::text as analysis,
    date_trunc('month', document.document_date)::date as month_start,
    sum(coalesce(document.amount, 0))::numeric as buy
  from public.financial_documents document
  cross join bounds
  where document.document_date >= bounds.date_from
    and document.document_date < bounds.date_to
    and upper(regexp_replace(coalesce(document.supplier_nif, ''), E'\\s+', '', 'g')) in ('509905099', 'PT509905099')
  group by 2
)
select
  months.analysis,
  extract(year from months.month_start)::integer as year,
  extract(month from months.month_start)::integer as month,
  to_char(months.month_start, 'YYYY-MM') as year_month,
  coalesce(sales_by_month.sales, 0)::numeric as sales,
  coalesce(buy_by_month.buy, 0)::numeric as buy,
  (coalesce(sales_by_month.sales, 0) - coalesce(buy_by_month.buy, 0))::numeric as revenue
from months
left join sales_by_month
  on sales_by_month.analysis = months.analysis
  and sales_by_month.month_start = months.month_start
left join buy_by_month
  on buy_by_month.analysis = months.analysis
  and buy_by_month.month_start = months.month_start
order by months.analysis, months.month_start;
$$;

create function public.get_bi_financial_items_revenue_payload(
  p_year_from integer default extract(year from current_date)::integer - 5,
  p_year_to integer default extract(year from current_date)::integer
)
returns table (rows jsonb)
language sql
stable
as $$
select coalesce(
  jsonb_agg(to_jsonb(aggregate_row) order by aggregate_row.analysis, aggregate_row.year, aggregate_row.month),
  '[]'::jsonb
) as rows
from public.get_bi_financial_items_revenue(p_year_from, p_year_to) aggregate_row;
$$;

grant execute on function public.get_bi_financial_items_revenue(integer, integer)
  to anon, authenticated, service_role;
grant execute on function public.get_bi_financial_items_revenue_payload(integer, integer)
  to anon, authenticated, service_role;

create index if not exists import_fdm_accounts_coffee_machine_items_revenue_idx
  on public.import_fdm_accounts (event_date)
  include (amount)
  where lower(btrim(coalesce(category, ''))) = 'coffee machine';

create index if not exists import_fdm_sales_souvenirs_items_revenue_idx
  on public.import_fdm_sales (sale_date)
  include (total)
  where lower(coalesce(sale_item, '')) like '%souvenir%';

create index if not exists financial_documents_souvenirs_items_revenue_idx
  on public.financial_documents (document_date, supplier_nif)
  include (amount)
  where upper(regexp_replace(coalesce(supplier_nif, ''), E'\\s+', '', 'g')) in ('509905099', 'PT509905099');

analyze public.import_fdm_accounts;
analyze public.import_fdm_sales;
analyze public.financial_documents;
