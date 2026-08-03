create or replace function public.bookings_bi_channel_share_last_12_months(
  p_ha text default null,
  p_statuses text[] default array['Checked Out', 'Confirmed', 'Arriving', 'Late', 'Leaving', 'Checked-in', 'Checked In']
)
returns table (
  year_month text,
  channel_label text,
  booking_count bigint
)
language sql
stable
as $$
with bounds as (
  select
    (date_trunc('month', current_date)::date - interval '12 months')::date as date_from,
    date_trunc('month', current_date)::date as date_to
)
select
  to_char(date_trunc('month', booking.check_in_date), 'YYYY-MM') as year_month,
  coalesce(nullif(btrim(booking.channel), ''), 'Unknown') as channel_label,
  count(*)::bigint as booking_count
from public.import_fdm_bookings booking
cross join bounds
where booking.check_in_date >= bounds.date_from
  and booking.check_in_date < bounds.date_to
  and (
    nullif(upper(btrim(coalesce(p_ha, ''))), '') is null
    or case when coalesce(booking.room_type, '') ilike '%Cruz%' then 'A' else 'H' end = upper(btrim(p_ha))
  )
  and (
    coalesce(array_length(p_statuses, 1), 0) = 0
    or coalesce(nullif(btrim(booking.status), ''), 'Unknown') = any(p_statuses)
  )
group by
  date_trunc('month', booking.check_in_date),
  coalesce(nullif(btrim(booking.channel), ''), 'Unknown')
order by
  date_trunc('month', booking.check_in_date),
  booking_count desc,
  channel_label;
$$;

grant execute on function public.bookings_bi_channel_share_last_12_months(text, text[])
  to anon, authenticated, service_role;

analyze public.import_fdm_bookings;
