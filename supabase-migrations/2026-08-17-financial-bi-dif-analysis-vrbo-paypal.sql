create or replace function public.get_bi_financial_dif_analysis(
  p_year_from integer default extract(year from current_date)::integer - 1,
  p_year_to integer default extract(year from current_date)::integer
)
returns table (
  analysis text,
  year integer,
  month integer,
  year_month text,
  fdm_amount numeric,
  extrato_amount numeric,
  difference numeric
)
language sql
stable
as $$
with bounds as (
  select
    make_date(least(coalesce(p_year_from, extract(year from current_date)::integer - 1), coalesce(p_year_to, extract(year from current_date)::integer)), 1, 1) as date_from,
    make_date(greatest(coalesce(p_year_from, extract(year from current_date)::integer - 1), coalesce(p_year_to, extract(year from current_date)::integer)) + 1, 1, 1) as date_to
),
analyses(analysis) as (
  values ('tpa'), ('adyen'), ('vrbo'), ('hw_cruz_direct')
),
months as (
  select analyses.analysis, month_start::date as month_start
  from bounds
  cross join analyses
  cross join lateral generate_series(bounds.date_from, bounds.date_to - interval '1 month', interval '1 month') as month_start
),
fdm_by_month(analysis, month_start, amount) as (
  select 'tpa'::text, date_trunc('month', account.event_date)::date, sum(coalesce(account.amount, 0))::numeric
  from public.import_fdm_accounts account cross join bounds
  where account.event_date >= bounds.date_from and account.event_date < bounds.date_to
    and lower(btrim(coalesce(account.account, ''))) = 'credit card'
    and lower(btrim(coalesce(account.category, ''))) not in ('transferouttoaccount')
  group by 2
  union all
  select 'adyen'::text, date_trunc('month', account.event_date)::date, sum(coalesce(account.amount, 0))::numeric
  from public.import_fdm_accounts account cross join bounds
  where account.event_date >= bounds.date_from and account.event_date < bounds.date_to
    and lower(btrim(coalesce(account.account, ''))) = 'adyen'
    and lower(btrim(coalesce(account.category, ''))) not in ('transferouttoaccount', 'deposittaken')
  group by 2
  union all
  select 'vrbo'::text, date_trunc('month', account.event_date)::date, sum(coalesce(account.amount, 0))::numeric
  from public.import_fdm_accounts account cross join bounds
  where account.event_date >= bounds.date_from and account.event_date < bounds.date_to
    and lower(btrim(coalesce(account.account, ''))) = 'bank transfer'
    and lower(btrim(coalesce(account.category, ''))) not in ('transferouttoaccount')
    and exists (
      select 1
      from public.import_fdm_bookings booking
      where btrim(coalesce(booking.booking_number, '')) = btrim(coalesce(account.reservation_id, ''))
        and lower(btrim(coalesce(booking.channel, ''))) = 'vrbo'
    )
  group by 2
  union all
  select 'hw_cruz_direct'::text, date_trunc('month', account.event_date)::date, sum(coalesce(account.amount, 0))::numeric
  from public.import_fdm_accounts account cross join bounds
  where account.event_date >= bounds.date_from and account.event_date < bounds.date_to
    and lower(btrim(coalesce(account.category, ''))) not in ('transferouttoaccount')
    and (
      (
        lower(btrim(coalesce(account.account, ''))) = 'bank transfer'
        and exists (
          select 1
          from public.import_fdm_bookings booking
          where btrim(coalesce(booking.booking_number, '')) = btrim(coalesce(account.reservation_id, ''))
            and lower(btrim(coalesce(booking.channel, ''))) = 'hostelworld'
        )
      )
      or lower(btrim(coalesce(account.account, ''))) = 'stripe'
    )
  group by 2
),
extrato_by_month(analysis, month_start, amount) as (
  select 'tpa'::text, date_trunc('month', statement.data)::date, sum(coalesce(statement.montante, 0))::numeric
  from public.import_cgd_extrato_ordem statement cross join bounds
  where statement.data >= bounds.date_from and statement.data < bounds.date_to
    and upper(coalesce(statement.descritivo, '')) like '%POS VENDAS%'
  group by 2
  union all
  select 'adyen'::text, date_trunc('month', statement.data)::date, sum(coalesce(statement.montante, 0))::numeric
  from public.import_cgd_extrato_ordem statement cross join bounds
  where statement.data >= bounds.date_from and statement.data < bounds.date_to
    and upper(coalesce(statement.descritivo, '')) like '%ADYEN%'
  group by 2
  union all
  select 'vrbo'::text, date_trunc('month', statement.data)::date, sum(coalesce(statement.montante, 0))::numeric
  from public.import_cgd_extrato_ordem statement cross join bounds
  where statement.data >= bounds.date_from and statement.data < bounds.date_to
    and upper(coalesce(statement.descritivo, '')) like '%TRF PAYPAL%'
  group by 2
  union all
  select 'hw_cruz_direct'::text, date_trunc('month', statement.data)::date, sum(coalesce(statement.montante, 0))::numeric
  from public.import_cgd_extrato_ordem statement cross join bounds
  where statement.data >= bounds.date_from and statement.data < bounds.date_to
    and upper(coalesce(statement.descritivo, '')) like '%STRIPE%'
  group by 2
)
select
  months.analysis,
  extract(year from months.month_start)::integer as year,
  extract(month from months.month_start)::integer as month,
  to_char(months.month_start, 'YYYY-MM') as year_month,
  coalesce(fdm_by_month.amount, 0)::numeric as fdm_amount,
  coalesce(extrato_by_month.amount, 0)::numeric as extrato_amount,
  (coalesce(extrato_by_month.amount, 0) - coalesce(fdm_by_month.amount, 0))::numeric as difference
from months
left join fdm_by_month on fdm_by_month.analysis = months.analysis and fdm_by_month.month_start = months.month_start
left join extrato_by_month on extrato_by_month.analysis = months.analysis and extrato_by_month.month_start = months.month_start
order by months.analysis, months.month_start;
$$;

grant execute on function public.get_bi_financial_dif_analysis(integer, integer)
  to anon, authenticated, service_role;
