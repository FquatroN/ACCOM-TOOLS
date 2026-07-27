create or replace function public.get_bi_financial_analysis_fdm_accounts(
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
configured_account_categories as (
  select distinct on (lower(btrim(config.category)))
    config.category,
    config.macro_category,
    config.is_result
  from (
    select
      coalesce(item ->> 'category', item ->> 'Category', item ->> 'CATEGORY') as category,
      coalesce(item ->> 'macroCategory', item ->> 'macro_category', item ->> 'Macro Category', item ->> 'MACRO_CATEGORY') as macro_category,
      lower(coalesce(item ->> 'isResult', item ->> 'is_result', item ->> 'IsResult', item ->> 'IS_RESULT', '')) in ('yes', 'true', '1', 'y') as is_result
    from public.app_settings settings
    cross join lateral jsonb_array_elements(coalesce((settings.payload::jsonb) -> 'fdmAccountCategories', '[]'::jsonb)) as item
    where settings.setting_key in ('bi_settings', 'bi-settings')
  ) config
  where nullif(btrim(config.category), '') is not null
  order by lower(btrim(config.category)), config.macro_category
),
default_account_categories(category, macro_category, is_result) as (
  values
    ('POSSales', 'Sales', true),
    ('ReservationSales', 'Sales', true),
    ('ReservationPayments', 'Reservation', true),
    ('DepositReturn', 'Deposit', false),
    ('RefundsAndReturns', 'Reservation', true),
    ('DepositTaken', 'Deposit', false),
    ('TransferInFromAccount', 'Transfer', false),
    ('TransferOutToAccount', 'Transfer', false),
    ('StaffWithdrawals', 'Deposit', false),
    ('Compras', 'Transfer', false),
    ('StaffDeposit', 'Deposit', false),
    ('CancellationCharges', 'Reservation', true),
    ('Tip', 'Sales', true),
    ('Coffee Machine', 'Sales', true),
    ('NoShowCharges', 'Reservation', true)
),
account_categories as (
  select category, macro_category, is_result
  from configured_account_categories
  union all
  select category, macro_category, is_result
  from default_account_categories
  where not exists (select 1 from configured_account_categories)
),
filtered_accounts as materialized (
  select
    account.event_date,
    account.reservation_id,
    account.category,
    account.amount
  from public.import_fdm_accounts account
  cross join filter_bounds bounds
  where account.event_date >= bounds.date_from
    and account.event_date < bounds.date_to
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
  extract(year from account.event_date)::integer as year,
  extract(month from account.event_date)::integer as month,
  to_char(account.event_date, 'YYYY-MM') as year_month,
  coalesce(booking.cc, 'H') as cc,
  coalesce(config.macro_category, '') as category,
  sum(coalesce(account.amount, 0))::numeric as total_amount
from filtered_accounts account
cross join filter_bounds bounds
inner join account_categories config
  on lower(btrim(config.category)) = lower(btrim(account.category))
  and config.is_result
left join lateral (
  select case when booking_row.room_type ilike '%Cruz%' then 'A' else 'H' end as cc
  from public.import_fdm_bookings booking_row
  where booking_row.booking_number = account.reservation_id
  order by booking_row.created_at desc
  limit 1
) booking on true
where bounds.filter_cc is null or coalesce(booking.cc, 'H') = bounds.filter_cc
group by
  extract(year from account.event_date),
  extract(month from account.event_date),
  to_char(account.event_date, 'YYYY-MM'),
  coalesce(booking.cc, 'H'),
  coalesce(config.macro_category, '');
$$;

create or replace function public.get_bi_financial_analysis_fdm_accounts_payload(
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
from public.get_bi_financial_analysis_fdm_accounts(p_year_from, p_year_to, p_cc) aggregate_row;
$$;

grant execute on function public.get_bi_financial_analysis_fdm_accounts(integer, integer, text)
  to anon, authenticated, service_role;
grant execute on function public.get_bi_financial_analysis_fdm_accounts_payload(integer, integer, text)
  to anon, authenticated, service_role;

create index if not exists import_fdm_accounts_bi_cover_idx
  on public.import_fdm_accounts (event_date, reservation_id, category)
  include (amount);

analyze public.import_fdm_accounts;
