create or replace view public.bookings_bi_channels_agg as
select
  extract(year from check_in_date)::integer as chart_year,
  case
    when coalesce(room_type, '') ilike '%Cruz%' then 'A'
    else 'H'
  end as ha,
  coalesce(nullif(btrim(status), ''), 'Unknown') as status,
  coalesce(nullif(btrim(channel), ''), 'Unknown') as channel_label,
  count(*)::bigint as booking_count
from public.import_fdm_bookings
where check_in_date is not null
group by 1, 2, 3, 4;

create or replace function public.bookings_bi_channels(
  p_year integer default extract(year from now())::integer,
  p_ha text default null,
  p_statuses text[] default array['Checked Out', 'Confirmed', 'Arriving', 'Late', 'Leaving', 'Checked-in']
)
returns table (
  chart_year integer,
  ha text,
  status text,
  channel_label text,
  booking_count bigint
)
language sql
stable
as $$
  with selected_years as (
    select generate_series(p_year - 3, p_year)::integer as chart_year
  )
  select
    agg.chart_year,
    agg.ha,
    agg.status,
    agg.channel_label,
    sum(agg.booking_count)::bigint as booking_count
  from public.bookings_bi_channels_agg agg
  join selected_years years on years.chart_year = agg.chart_year
  where (
      nullif(btrim(coalesce(p_ha, '')), '') is null
      or upper(agg.ha) = upper(btrim(p_ha))
    )
    and (
      coalesce(array_length(p_statuses, 1), 0) = 0
      or agg.status = any(p_statuses)
    )
  group by
    agg.chart_year,
    agg.ha,
    agg.status,
    agg.channel_label
  order by
    agg.chart_year desc,
    sum(agg.booking_count) desc,
    agg.channel_label asc;
$$;

grant select on public.bookings_bi_channels_agg to anon, authenticated, service_role;
grant execute on function public.bookings_bi_channels(integer, text, text[]) to anon, authenticated, service_role;
