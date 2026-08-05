drop function if exists public.get_bi_financial_income_book_vs_accounts_payload(integer, integer);
drop function if exists public.get_bi_financial_income_book_vs_accounts(integer, integer);

create function public.get_bi_financial_income_book_vs_accounts(
  p_year_from integer default extract(year from current_date)::integer - 5,
  p_year_to integer default extract(year from current_date)::integer
)
returns table (
  year integer,
  month integer,
  year_month text,
  accommodation numeric,
  drinks numeric,
  booking_sales numeric,
  tmt numeric,
  tours numeric,
  income_bookings numeric,
  reservation numeric,
  account_sales numeric,
  income_accounts numeric,
  difference numeric
)
language sql
stable
as $$
with bounds as (
  select
    make_date(least(coalesce(p_year_from, extract(year from current_date)::integer - 5), coalesce(p_year_to, extract(year from current_date)::integer)), 1, 1) as date_from,
    make_date(greatest(coalesce(p_year_from, extract(year from current_date)::integer - 5), coalesce(p_year_to, extract(year from current_date)::integer)) + 1, 1, 1) as date_to
),
months as (
  select month_start::date as month_start
  from bounds
  cross join lateral generate_series(bounds.date_from, bounds.date_to - interval '1 month', interval '1 month') as month_start
),
bookings as (
  select
    source.year,
    source.month,
    max(source.year_month) as year_month,
    sum(source.total_amount) filter (where lower(btrim(source.category)) = 'accomodation')::numeric as accommodation,
    sum(source.total_amount) filter (where lower(btrim(source.category)) = 'drinks')::numeric as drinks,
    sum(source.total_amount) filter (where lower(btrim(source.category)) = 'sales')::numeric as booking_sales,
    sum(source.total_amount) filter (where lower(btrim(source.category)) = 'tmt')::numeric as tmt,
    sum(source.total_amount) filter (where lower(btrim(source.category)) = 'tours')::numeric as tours
  from bounds
  cross join lateral public.get_bi_financial_analysis_sales(
    extract(year from bounds.date_from)::integer,
    extract(year from bounds.date_to)::integer - 1,
    null
  ) source
  where source.type = 'INCOME'
    and lower(btrim(source.category)) in ('accomodation', 'drinks', 'sales', 'tmt', 'tours')
  group by source.year, source.month
),
accounts as (
  select
    source.year,
    source.month,
    max(source.year_month) as year_month,
    sum(source.total_amount) filter (where lower(btrim(source.category)) = 'reservation')::numeric as reservation,
    sum(source.total_amount) filter (where lower(btrim(source.category)) = 'sales')::numeric as account_sales
  from bounds
  cross join lateral public.get_bi_financial_analysis_fdm_accounts(
    extract(year from bounds.date_from)::integer,
    extract(year from bounds.date_to)::integer - 1,
    null
  ) source
  where source.type = 'INCOME'
    and lower(btrim(source.category)) in ('reservation', 'sales')
  group by source.year, source.month
),
monthly as (
  select
    extract(year from months.month_start)::integer as year,
    extract(month from months.month_start)::integer as month,
    to_char(months.month_start, 'YYYY-MM') as year_month,
    coalesce(bookings.accommodation, 0)::numeric as accommodation,
    coalesce(bookings.drinks, 0)::numeric as drinks,
    coalesce(bookings.booking_sales, 0)::numeric as booking_sales,
    coalesce(bookings.tmt, 0)::numeric as tmt,
    coalesce(bookings.tours, 0)::numeric as tours,
    coalesce(accounts.reservation, 0)::numeric as reservation,
    coalesce(accounts.account_sales, 0)::numeric as account_sales
  from months
  left join bookings
    on bookings.year = extract(year from months.month_start)::integer
    and bookings.month = extract(month from months.month_start)::integer
  left join accounts
    on accounts.year = extract(year from months.month_start)::integer
    and accounts.month = extract(month from months.month_start)::integer
)
select
  year,
  month,
  year_month,
  accommodation,
  drinks,
  booking_sales,
  tmt,
  tours,
  (accommodation + drinks + booking_sales + tmt + tours)::numeric as income_bookings,
  reservation,
  account_sales,
  (reservation + account_sales)::numeric as income_accounts,
  ((reservation + account_sales) - (accommodation + drinks + booking_sales + tmt + tours))::numeric as difference
from monthly
order by year, month;
$$;

create function public.get_bi_financial_income_book_vs_accounts_payload(
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
  from public.get_bi_financial_income_book_vs_accounts(p_year_from, p_year_to) aggregate_row;
$$;

grant execute on function public.get_bi_financial_income_book_vs_accounts(integer, integer)
  to anon, authenticated, service_role;
grant execute on function public.get_bi_financial_income_book_vs_accounts_payload(integer, integer)
  to anon, authenticated, service_role;
