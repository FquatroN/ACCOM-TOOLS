create or replace function public.get_bi_financial_cash_analysis(
  p_year_from integer default extract(year from current_date)::integer - 5,
  p_year_to integer default extract(year from current_date)::integer
)
returns table (
  year integer,
  month integer,
  year_month text,
  total_withdrawals numeric,
  total_deposits numeric,
  nao_faturado numeric,
  analysis numeric
)
language sql
stable
as $$
with bounds as (
  select
    make_date(least(coalesce(p_year_from, extract(year from current_date)::integer - 5), coalesce(p_year_to, extract(year from current_date)::integer)), 1, 1) as date_from,
    make_date(greatest(coalesce(p_year_from, extract(year from current_date)::integer - 5), coalesce(p_year_to, extract(year from current_date)::integer)) + 1, 1, 1) as date_to
),
configured_account_categories as (
  select distinct on (lower(btrim(config.category)))
    config.category,
    config.is_result
  from (
    select
      coalesce(item ->> 'category', item ->> 'Category', item ->> 'CATEGORY') as category,
      lower(coalesce(item ->> 'isResult', item ->> 'is_result', item ->> 'IsResult', item ->> 'IS_RESULT', '')) in ('yes', 'true', '1', 'y') as is_result
    from public.app_settings settings
    cross join lateral jsonb_array_elements(coalesce((settings.payload::jsonb) -> 'fdmAccountCategories', '[]'::jsonb)) as item
    where settings.setting_key in ('bi_settings', 'bi-settings')
  ) config
  where nullif(btrim(config.category), '') is not null
  order by lower(btrim(config.category))
),
default_account_categories(category, is_result) as (
  values
    ('POSSales', true),
    ('ReservationSales', true),
    ('ReservationPayments', true),
    ('DepositReturn', false),
    ('RefundsAndReturns', true),
    ('DepositTaken', false),
    ('TransferInFromAccount', false),
    ('TransferOutToAccount', false),
    ('StaffWithdrawals', false),
    ('Compras', false),
    ('StaffDeposit', false),
    ('CancellationCharges', true),
    ('Tip', true),
    ('Coffee Machine', true),
    ('NoShowCharges', true)
),
account_categories as (
  select category, is_result from configured_account_categories
  union all
  select category, is_result from default_account_categories
  where not exists (select 1 from configured_account_categories)
),
months as (
  select month_start::date
  from bounds,
  lateral generate_series(bounds.date_from, bounds.date_to - interval '1 month', interval '1 month') as month_start
),
withdrawals as (
  select
    date_trunc('month', account.event_date)::date as month_start,
    sum(abs(account.amount))::numeric as total_withdrawals
  from public.import_fdm_accounts account
  cross join bounds
  where account.event_date >= bounds.date_from
    and account.event_date < bounds.date_to
    and lower(btrim(account.account)) = 'cash box'
    and lower(btrim(account.category)) = 'staffwithdrawals'
  group by 1
),
deposits as (
  select
    date_trunc('month', statement.data)::date as month_start,
    sum(statement.montante)::numeric as total_deposits
  from public.import_cgd_extrato_ordem statement
  cross join bounds
  where statement.data >= bounds.date_from
    and statement.data < bounds.date_to
    and statement.descritivo like '%DEP%'
    and statement.montante > 0
  group by 1
),
unbilled as (
  select
    date_trunc('month', account.event_date)::date as month_start,
    sum(account.amount)::numeric as nao_faturado
  from public.import_fdm_accounts account
  cross join bounds
  inner join account_categories category_config
    on lower(btrim(category_config.category)) = lower(btrim(account.category))
    and category_config.is_result
  where account.event_date >= bounds.date_from
    and account.event_date < bounds.date_to
    and upper(btrim(account.invoice)) = 'N'
  group by 1
)
select
  extract(year from months.month_start)::integer as year,
  extract(month from months.month_start)::integer as month,
  to_char(months.month_start, 'YYYY-MM') as year_month,
  coalesce(withdrawals.total_withdrawals, 0)::numeric as total_withdrawals,
  coalesce(deposits.total_deposits, 0)::numeric as total_deposits,
  coalesce(unbilled.nao_faturado, 0)::numeric as nao_faturado,
  (coalesce(deposits.total_deposits, 0) + coalesce(unbilled.nao_faturado, 0) - coalesce(withdrawals.total_withdrawals, 0))::numeric as analysis
from months
left join withdrawals on withdrawals.month_start = months.month_start
left join deposits on deposits.month_start = months.month_start
left join unbilled on unbilled.month_start = months.month_start
order by months.month_start;
$$;

create or replace function public.get_bi_financial_cash_analysis_payload(
  p_year_from integer default extract(year from current_date)::integer - 5,
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
from public.get_bi_financial_cash_analysis(p_year_from, p_year_to) aggregate_row;
$$;

grant execute on function public.get_bi_financial_cash_analysis(integer, integer)
  to anon, authenticated, service_role;
grant execute on function public.get_bi_financial_cash_analysis_payload(integer, integer)
  to anon, authenticated, service_role;

create index if not exists import_fdm_accounts_cash_analysis_cover_idx
  on public.import_fdm_accounts (event_date, account, category, invoice)
  include (amount);

create index if not exists import_cgd_extrato_ordem_cash_analysis_cover_idx
  on public.import_cgd_extrato_ordem (data)
  include (montante, descritivo);

analyze public.import_fdm_accounts;
analyze public.import_cgd_extrato_ordem;
