-- Adds a five-year H/A booking-window comparison that is independent of the selected year and HA filters.
create or replace function public.bookings_bi_booking_window(
  p_year integer default extract(year from now())::integer,
  p_ha text default null,
  p_statuses text[] default array['Checked Out', 'Confirmed', 'Arriving', 'Late', 'Leaving', 'Checked-in', 'Checked In']
)
returns jsonb
language sql
stable
as $$
  with filtered as (
    select booking_date, coalesce(nullif(btrim(channel), ''), 'Unknown') as channel_label,
      (check_in_date - booking_date)::integer as booking_window_days
    from public.import_fdm_bookings
    where booking_date is not null and check_in_date is not null
      and extract(year from booking_date)::integer = p_year and check_in_date >= booking_date
      and (nullif(btrim(coalesce(p_ha, '')), '') is null or case when coalesce(room_type, '') ilike '%Cruz%' then 'A' else 'H' end = upper(btrim(p_ha)))
      and (coalesce(array_length(p_statuses, 1), 0) = 0 or coalesce(nullif(btrim(status), ''), 'Unknown') = any(p_statuses))
  ),
  summary as (
    select count(*)::bigint as booking_count, round(avg(booking_window_days)::numeric, 1) as average_days,
      round(percentile_cont(0.5) within group (order by booking_window_days)::numeric, 1) as median_days,
      count(*) filter (where booking_window_days = 0)::bigint as same_day_count,
      count(*) filter (where booking_window_days <= 7)::bigint as within_7_days_count,
      count(*) filter (where booking_window_days >= 31)::bigint as over_30_days_count
    from filtered
  ),
  distribution as (
    select case when booking_window_days = 0 then 'Same day' when booking_window_days <= 3 then '1–3 days' when booking_window_days <= 7 then '4–7 days' when booking_window_days <= 14 then '8–14 days' when booking_window_days <= 30 then '15–30 days' when booking_window_days <= 60 then '31–60 days' when booking_window_days <= 90 then '61–90 days' else '91+ days' end as label,
      case when booking_window_days = 0 then 1 when booking_window_days <= 3 then 2 when booking_window_days <= 7 then 3 when booking_window_days <= 14 then 4 when booking_window_days <= 30 then 5 when booking_window_days <= 60 then 6 when booking_window_days <= 90 then 7 else 8 end as sort_order,
      count(*)::bigint as booking_count
    from filtered group by 1, 2
  ),
  months as (
    select to_char(date_trunc('month', booking_date), 'YYYY-MM') as booking_month, count(*)::bigint as booking_count,
      round(avg(booking_window_days)::numeric, 1) as average_days,
      round(percentile_cont(0.5) within group (order by booking_window_days)::numeric, 1) as median_days
    from filtered group by 1
  ),
  channels as (
    select channel_label, count(*)::bigint as booking_count, round(avg(booking_window_days)::numeric, 1) as average_days,
      round(percentile_cont(0.5) within group (order by booking_window_days)::numeric, 1) as median_days,
      count(*) filter (where booking_window_days = 0)::bigint as same_day_count,
      count(*) filter (where booking_window_days <= 7)::bigint as within_7_days_count,
      count(*) filter (where booking_window_days >= 31)::bigint as over_30_days_count
    from filtered group by 1
  ),
  years as (
    select distinct extract(year from booking_date)::integer as booking_year from public.import_fdm_bookings where booking_date is not null
  ),
  trend_years as (
    select booking_year from years order by booking_year desc limit 5
  ),
  year_trend as (
    select extract(year from booking_date)::integer as booking_year,
      case when coalesce(room_type, '') ilike '%Cruz%' then 'A' else 'H' end as ha,
      round(avg((check_in_date - booking_date)::integer)::numeric, 1) as average_days
    from public.import_fdm_bookings
    where booking_date is not null and check_in_date is not null and check_in_date >= booking_date
      and extract(year from booking_date)::integer in (select booking_year from trend_years)
      and (coalesce(array_length(p_statuses, 1), 0) = 0 or coalesce(nullif(btrim(status), ''), 'Unknown') = any(p_statuses))
    group by 1, 2
  )
  select jsonb_build_object(
    'years', coalesce((select jsonb_agg(booking_year order by booking_year desc) from years), '[]'::jsonb),
    'summary', coalesce((select to_jsonb(summary) from summary), '{}'::jsonb),
    'distribution', coalesce((select jsonb_agg(jsonb_build_object('label', label, 'bookingCount', booking_count) order by sort_order) from distribution), '[]'::jsonb),
    'months', coalesce((select jsonb_agg(jsonb_build_object('bookingMonth', booking_month, 'bookingCount', booking_count, 'averageDays', average_days, 'medianDays', median_days) order by booking_month) from months), '[]'::jsonb),
    'channels', coalesce((select jsonb_agg(jsonb_build_object('channel', channel_label, 'bookingCount', booking_count, 'averageDays', average_days, 'medianDays', median_days, 'sameDayCount', same_day_count, 'within7DaysCount', within_7_days_count, 'over30DaysCount', over_30_days_count) order by booking_count desc, channel_label) from channels), '[]'::jsonb),
    'yearTrend', coalesce((select jsonb_agg(jsonb_build_object('bookingYear', booking_year, 'ha', ha, 'averageDays', average_days) order by booking_year, ha) from year_trend), '[]'::jsonb)
  );
$$;
