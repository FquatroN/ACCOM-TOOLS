create or replace function public.bookings_bi_channel_monthly_share(
  p_year integer default extract(year from now())::integer,
  p_ha text default null,
  p_statuses text[] default array['Checked Out', 'Confirmed', 'Arriving', 'Late', 'Leaving', 'Checked-in', 'Checked In']
)
returns table (
  month integer,
  channel_label text,
  booking_count bigint
)
language sql
stable
as $$
select
  extract(month from booking.check_in_date)::integer as month,
  coalesce(nullif(btrim(booking.channel), ''), 'Unknown') as channel_label,
  count(*)::bigint as booking_count
from public.import_fdm_bookings booking
where booking.check_in_date >= make_date(p_year, 1, 1)
  and booking.check_in_date < make_date(p_year + 1, 1, 1)
  and (
    nullif(upper(btrim(coalesce(p_ha, ''))), '') is null
    or case when coalesce(booking.room_type, '') ilike '%Cruz%' then 'A' else 'H' end = upper(btrim(p_ha))
  )
  and (
    coalesce(array_length(p_statuses, 1), 0) = 0
    or coalesce(nullif(btrim(booking.status), ''), 'Unknown') = any(p_statuses)
  )
group by
  extract(month from booking.check_in_date),
  coalesce(nullif(btrim(booking.channel), ''), 'Unknown')
order by month, booking_count desc, channel_label;
$$;

grant execute on function public.bookings_bi_channel_monthly_share(integer, text, text[])
  to anon, authenticated, service_role;

create index if not exists import_fdm_bookings_monthly_channel_bi_idx
  on public.import_fdm_bookings (check_in_date, status)
  include (room_type, channel);

analyze public.import_fdm_bookings;
