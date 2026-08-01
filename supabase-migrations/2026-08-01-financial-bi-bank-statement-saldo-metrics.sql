drop function if exists public.get_bi_financial_bank_statement_payload(integer, integer);
drop function if exists public.get_bi_financial_bank_statement(integer, integer);

create function public.get_bi_financial_bank_statement(
  p_year_from integer default extract(year from current_date)::integer - 10,
  p_year_to integer default extract(year from current_date)::integer
)
returns table (
  year integer,
  month integer,
  year_month text,
  sum_amount numeric,
  saldo_sum numeric,
  average_saldo numeric,
  min_saldo numeric,
  max_saldo numeric,
  saldo_count bigint
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
  extract(year from statement.data)::integer as year,
  extract(month from statement.data)::integer as month,
  to_char(statement.data, 'YYYY-MM') as year_month,
  sum(statement.montante)::numeric as sum_amount,
  coalesce(sum(statement.saldo), 0)::numeric as saldo_sum,
  avg(statement.saldo)::numeric as average_saldo,
  min(statement.saldo)::numeric as min_saldo,
  max(statement.saldo)::numeric as max_saldo,
  count(statement.saldo)::bigint as saldo_count
from public.import_cgd_extrato_ordem statement
cross join bounds
where statement.data >= bounds.date_from
  and statement.data < bounds.date_to
  and statement.montante is not null
group by
  extract(year from statement.data),
  extract(month from statement.data),
  to_char(statement.data, 'YYYY-MM')
order by year, month;
$$;

create function public.get_bi_financial_bank_statement_payload(
  p_year_from integer default extract(year from current_date)::integer - 10,
  p_year_to integer default extract(year from current_date)::integer
)
returns table (rows jsonb)
language sql
stable
as $$
select coalesce(
  jsonb_agg(to_jsonb(aggregate_row) order by aggregate_row.year, aggregate_row.month),
  '[]'::jsonb
) as rows
from public.get_bi_financial_bank_statement(p_year_from, p_year_to) aggregate_row;
$$;

grant execute on function public.get_bi_financial_bank_statement(integer, integer)
  to anon, authenticated, service_role;
grant execute on function public.get_bi_financial_bank_statement_payload(integer, integer)
  to anon, authenticated, service_role;

create index if not exists import_cgd_extrato_ordem_bank_statement_saldo_cover_idx
  on public.import_cgd_extrato_ordem (data)
  include (montante, saldo)
  where data is not null and montante is not null;

analyze public.import_cgd_extrato_ordem;
