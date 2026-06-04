drop function if exists public.guests_bi_years();
drop function if exists public.guests_bi_tmt(integer);
drop function if exists public.guests_bi_nationality_pies();
drop function if exists public.guests_bi_nationality_line();
drop function if exists public.guests_bi_nationality_month_line();

create or replace function public.guests_bi_years(p_ha text default null)
returns table(year integer)
language sql
stable
as $$
  with ha_filter as (
    select nullif(upper(trim(coalesce(p_ha, ''))), '') as ha_value
  )
  select distinct extract(year from gr.check_in)::int as year
  from public.guest_records gr
  cross join ha_filter hf
  where gr.check_in is not null
    and (hf.ha_value is null or upper(trim(coalesce(gr.ha, ''))) = hf.ha_value)
  order by year desc;
$$;

create or replace function public.guests_bi_tmt(p_year integer default null, p_ha text default null)
returns table(
  year_month text,
  total_nights bigint,
  exempt_7days bigint,
  exempt_13_year bigint
)
language sql
stable
as $$
  with ha_filter as (
    select nullif(upper(trim(coalesce(p_ha, ''))), '') as ha_value
  ),
  selected_year as (
    select coalesce(
      p_year,
      (
        select max(extract(year from gr.check_in))::int
        from public.guest_records gr
        cross join ha_filter hf
        where gr.check_in is not null
          and (hf.ha_value is null or upper(trim(coalesce(gr.ha, ''))) = hf.ha_value)
      ),
      extract(year from current_date)::int
    ) as year_value
  ),
  months as (
    select
      make_date((select year_value from selected_year), month_num, 1) as month_start,
      to_char(make_date((select year_value from selected_year), month_num, 1), 'YYYY-MM') as year_month
    from generate_series(1, 12) as month_num
  ),
  guest_rows as (
    select
      to_char(gr.check_in, 'YYYY-MM') as year_month,
      greatest((gr.check_out - gr.check_in), 0)::bigint as nights,
      greatest((gr.check_out - gr.check_in) - 7, 0)::bigint as exempt_7days,
      case
        when extract(year from age(gr.check_in, gr.birth_date)) < 13
          then greatest((gr.check_out - gr.check_in), 0)::bigint
        else 0::bigint
      end as exempt_13_year
    from public.guest_records gr
    cross join ha_filter hf
    where gr.check_in is not null
      and gr.check_out is not null
      and (hf.ha_value is null or upper(trim(coalesce(gr.ha, ''))) = hf.ha_value)
      and gr.check_in >= make_date((select year_value from selected_year), 1, 1)
      and gr.check_in < make_date((select year_value from selected_year) + 1, 1, 1)
  )
  select
    m.year_month,
    coalesce(sum(g.nights), 0)::bigint as total_nights,
    coalesce(sum(g.exempt_7days), 0)::bigint as exempt_7days,
    coalesce(sum(g.exempt_13_year), 0)::bigint as exempt_13_year
  from months m
  left join guest_rows g on g.year_month = m.year_month
  group by m.month_start, m.year_month
  order by m.month_start;
$$;

create or replace function public.guests_bi_nationality_pies(p_ha text default null)
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
  ha_filter as (
    select nullif(upper(trim(coalesce(p_ha, ''))), '') as ha_value
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
    cross join ha_filter hf
    where gr.check_in is not null
      and (hf.ha_value is null or upper(trim(coalesce(gr.ha, ''))) = hf.ha_value)
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

create or replace function public.guests_bi_nationality_line(p_ha text default null)
returns table(
  chart_year integer,
  country_label text,
  guest_count bigint,
  sort_order integer
)
language sql
stable
as $$
  with ha_filter as (
    select nullif(upper(trim(coalesce(p_ha, ''))), '') as ha_value
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
    cross join ha_filter hf
    where gr.check_in is not null
      and (hf.ha_value is null or upper(trim(coalesce(gr.ha, ''))) = hf.ha_value)
    group by 1, 2
  ),
  year_list as (
    select distinct chart_year
    from base_rows
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
      y.chart_year,
      t.country_key,
      t.country_label,
      t.total_guests,
      coalesce(b.guest_count, 0)::bigint as guest_count
    from year_list y
    cross join top_countries t
    left join base_rows b
      on b.chart_year = y.chart_year
     and b.country_key = t.country_key
  ),
  top_rows as (
    select
      chart_year,
      country_label,
      guest_count,
      row_number() over (order by total_guests desc, country_label) as sort_order
    from expanded_top
  ),
  others_rows as (
    select
      y.chart_year,
      'Others'::text as country_label,
      coalesce(sum(b.guest_count), 0)::bigint as guest_count,
      21 as sort_order
    from year_list y
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
  order by chart_year asc, sort_order asc, country_label asc;
$$;

create or replace function public.guests_bi_nationality_month_line(p_ha text default null)
returns table(
  chart_month integer,
  country_label text,
  guest_count bigint,
  sort_order integer
)
language sql
stable
as $$
  with ha_filter as (
    select nullif(upper(trim(coalesce(p_ha, ''))), '') as ha_value
  ),
  base_rows as (
    select
      extract(month from gr.check_in)::int as chart_month,
      public.guests_bi_nationality_key(gr.nationality, gr.nationality_code) as country_key,
      count(*)::bigint as guest_count,
      coalesce(
        mode() within group (
          order by public.guests_bi_nationality_label(gr.nationality, gr.nationality_code)
        ) filter (where public.guests_bi_nationality_label(gr.nationality, gr.nationality_code) <> ''),
        initcap(lower(public.guests_bi_nationality_key(gr.nationality, gr.nationality_code)))
      ) as country_label
    from public.guest_records gr
    cross join ha_filter hf
    where gr.check_in is not null
      and (hf.ha_value is null or upper(trim(coalesce(gr.ha, ''))) = hf.ha_value)
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
