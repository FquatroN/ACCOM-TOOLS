drop function if exists public.get_bi_financial_utilities_payload(integer, integer, text);
drop function if exists public.get_bi_financial_utilities(integer, integer, text);

create function public.get_bi_financial_utilities(
  p_year_from integer default extract(year from current_date)::integer - 10,
  p_year_to integer default extract(year from current_date)::integer,
  p_supplier_nifs text[] default null
)
returns table (
  year integer,
  month integer,
  year_month text,
  cc text,
  sum_amount numeric,
  document_count bigint
)
language sql
stable
as $$
with bounds as (
  select
    make_date(least(coalesce(p_year_from, extract(year from current_date)::integer - 10), coalesce(p_year_to, extract(year from current_date)::integer)), 1, 1) as date_from,
    make_date(greatest(coalesce(p_year_from, extract(year from current_date)::integer - 10), coalesce(p_year_to, extract(year from current_date)::integer)) + 1, 1, 1) as date_to
)
select
  extract(year from document.document_date)::integer as year,
  extract(month from document.document_date)::integer as month,
  to_char(document.document_date, 'YYYY-MM') as year_month,
  upper(btrim(document.cc)) as cc,
  sum(coalesce(document.amount, 0))::numeric as sum_amount,
  count(document.amount)::bigint as document_count
from public.financial_documents document
cross join bounds
where document.document_date >= bounds.date_from
  and document.document_date < bounds.date_to
  and lower(btrim(coalesce(document.category, ''))) = 'utility'
  and upper(btrim(coalesce(document.cc, ''))) in ('H', 'A')
  and (coalesce(cardinality(p_supplier_nifs), 0) = 0 or btrim(coalesce(document.supplier_nif, '')) = any(p_supplier_nifs))
group by
  extract(year from document.document_date),
  extract(month from document.document_date),
  to_char(document.document_date, 'YYYY-MM'),
  upper(btrim(document.cc))
order by year, month, cc;
$$;

create function public.get_bi_financial_utilities_payload(
  p_year_from integer default extract(year from current_date)::integer - 10,
  p_year_to integer default extract(year from current_date)::integer,
  p_supplier_nifs text[] default null
)
returns table (
  rows jsonb,
  suppliers jsonb
)
language sql
stable
as $$
select
  coalesce((
    select jsonb_agg(to_jsonb(aggregate_row) order by aggregate_row.year, aggregate_row.month, aggregate_row.cc)
    from public.get_bi_financial_utilities(p_year_from, p_year_to, p_supplier_nifs) aggregate_row
  ), '[]'::jsonb) as rows,
  coalesce((
    select jsonb_agg(jsonb_build_object('nif', supplier_nif, 'name', supplier_name) order by supplier_nif)
    from (
      select
        btrim(document.supplier_nif) as supplier_nif,
        coalesce(min(nullif(btrim(document.supplier_name), '')), '') as supplier_name
      from public.financial_documents document
      where lower(btrim(coalesce(document.category, ''))) = 'utility'
        and nullif(btrim(coalesce(document.supplier_nif, '')), '') is not null
      group by btrim(document.supplier_nif)
    ) suppliers
  ), '[]'::jsonb) as suppliers;
$$;

grant execute on function public.get_bi_financial_utilities(integer, integer, text[])
  to anon, authenticated, service_role;
grant execute on function public.get_bi_financial_utilities_payload(integer, integer, text[])
  to anon, authenticated, service_role;

create index if not exists financial_documents_utilities_bi_cover_idx
  on public.financial_documents (document_date, cc, supplier_nif)
  include (amount)
  where lower(btrim(coalesce(category, ''))) = 'utility';

analyze public.financial_documents;
