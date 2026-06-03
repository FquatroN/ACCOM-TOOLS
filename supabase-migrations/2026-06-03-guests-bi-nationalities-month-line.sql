create or replace function public.guests_bi_nationality_month_line()
returns table(
  chart_month integer,
  country_label text,
  guest_count bigint,
  sort_order integer
)
language sql
stable
as $$
  with base_rows as (
    select
      extract(month from gr.check_in)::int as chart_month,
      public.guests_bi_nationality_key(gr.nationality, gr.nationality_code) as country_key,
      count(*)::bigint as guest_count,
      coalesce(
        mode() within group (order by nullif(trim(gr.nationality), ''))
          filter (where nullif(trim(gr.nationality), '') is not null),
        initcap(lower(public.guests_bi_nationality_key(gr.nationality, gr.nationality_code)))
      ) as country_label
    from public.guest_records gr
    where gr.check_in is not null
    group by 1, 2
  ),
  month_list as (
    select generate_series(1, 12)::int as chart_month
  ),
  top_countries as (
    select
      country_key,
      min(country_label) as country_label,
      sum(guest_count) as total_guests
    from base_rows
    group by country_key
    order by total_guests desc, min(country_label)
    limit 20
  ),
  expanded_top as (
    select
      m.chart_month,
      t.country_key,
      t.country_label,
      t.total_guests,
      coalesce(b.guest_count, 0)::bigint as guest_count
    from month_list m
    cross join top_countries t
    left join base_rows b
      on b.chart_month = m.chart_month
     and b.country_key = t.country_key
  ),
  top_rows as (
    select
      chart_month,
      country_label,
      guest_count,
      row_number() over (order by total_guests desc, country_label) as sort_order
    from expanded_top
  ),
  others_rows as (
    select
      m.chart_month,
      'Others'::text as country_label,
      coalesce(sum(b.guest_count), 0)::bigint as guest_count,
      21 as sort_order
    from month_list m
    left join base_rows b
      on b.chart_month = m.chart_month
     and b.country_key not in (select country_key from top_countries)
    group by m.chart_month
  )
  select chart_month, country_label, guest_count, sort_order
  from (
    select * from top_rows
    union all
    select * from others_rows
  ) rows
  order by chart_month asc, sort_order asc, country_label asc;
$$;
