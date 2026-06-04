create or replace function public.guests_bi_nationality_pies()
returns table(
  chart_year integer,
  country_label text,
  guest_count bigint,
  sort_order integer
)
language sql
stable
as $$
  with current_year as (
    select extract(year from current_date)::int as year_value
  ),
  target_years as (
    select generate_series(
      (select year_value from current_year) - 3,
      (select year_value from current_year)
    )::int as chart_year
  ),
  base_rows as (
    select
      extract(year from gr.check_in)::int as chart_year,
      public.guests_bi_nationality_key(gr.nationality, gr.nationality_code) as country_key,
      count(*)::bigint as guest_count,
      coalesce(
        mode() within group (
          order by public.guests_bi_nationality_label(gr.nationality, gr.nationality_code)
        ) filter (where public.guests_bi_nationality_label(gr.nationality, gr.nationality_code) <> ''),
        initcap(lower(public.guests_bi_nationality_key(gr.nationality, gr.nationality_code)))
      ) as country_label
    from public.guest_records gr
    where gr.check_in is not null
    group by 1, 2
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
  top_rows as (
    select
      b.chart_year,
      t.country_label,
      b.guest_count,
      row_number() over (partition by b.chart_year order by b.guest_count desc, t.country_label) as sort_order
    from base_rows b
    join top_countries t on t.country_key = b.country_key
    where b.chart_year in (select chart_year from target_years)
  ),
  others_rows as (
    select
      y.chart_year,
      'Others'::text as country_label,
      coalesce(sum(b.guest_count), 0)::bigint as guest_count,
      21 as sort_order
    from target_years y
    left join base_rows b
      on b.chart_year = y.chart_year
     and b.country_key not in (select country_key from top_countries)
    group by y.chart_year
  )
  select chart_year, country_label, guest_count, sort_order
  from (
    select * from top_rows
    union all
    select * from others_rows
  ) rows
  where guest_count > 0
  order by chart_year desc, sort_order, country_label;
$$;
