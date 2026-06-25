create or replace function public.guests_bi_details(
  p_year integer,
  p_ha text default null
)
returns table (
  section text,
  sort_order integer,
  chart_year text,
  age_segment text,
  guest_count bigint,
  average_age numeric,
  average_stay numeric
)
language sql
stable
as $$
  with normalized as (
    select
      gr.ha,
      gr.check_in::date as check_in,
      gr.check_out::date as check_out,
      gr.birth_date::date as birth_date,
      extract(year from gr.check_in::date)::integer as check_in_year,
      greatest(coalesce(gr.check_out::date - gr.check_in::date, 0), 0)::numeric as stay_nights,
      case
        when gr.birth_date is null or gr.check_in is null then null
        else extract(year from age(gr.check_in::date, gr.birth_date::date))::integer
      end as age_at_check_in
    from public.guest_records gr
    where gr.check_in is not null
      and (nullif(btrim(coalesce(p_ha, '')), '') is null or gr.ha = upper(btrim(p_ha)))
  ),
  selected as (
    select *
    from normalized
    where check_in_year = p_year
  ),
  segment_seed as (
    select *
    from (values
      (1, '0-12'),
      (2, '13-17'),
      (3, '18-25'),
      (4, '26-35'),
      (5, '36-45'),
      (6, '46-55'),
      (7, '56-65'),
      (8, '66+'),
      (9, 'Unknown')
    ) as seed(sort_order, age_segment)
  ),
  selected_segments as (
    select
      case
        when age_at_check_in is null then 'Unknown'
        when age_at_check_in <= 12 then '0-12'
        when age_at_check_in <= 17 then '13-17'
        when age_at_check_in <= 25 then '18-25'
        when age_at_check_in <= 35 then '26-35'
        when age_at_check_in <= 45 then '36-45'
        when age_at_check_in <= 55 then '46-55'
        when age_at_check_in <= 65 then '56-65'
        else '66+'
      end as age_segment,
      count(*)::bigint as guest_count
    from selected
    group by 1
  )
  select
    'summary'::text as section,
    0::integer as sort_order,
    p_year::text as chart_year,
    null::text as age_segment,
    count(*)::bigint as guest_count,
    round(avg(age_at_check_in)::numeric, 2) as average_age,
    round(avg(stay_nights)::numeric, 2) as average_stay
  from selected

  union all

  select
    'segment'::text as section,
    seed.sort_order,
    p_year::text as chart_year,
    seed.age_segment,
    coalesce(seg.guest_count, 0)::bigint as guest_count,
    null::numeric as average_age,
    null::numeric as average_stay
  from segment_seed seed
  left join selected_segments seg on seg.age_segment = seed.age_segment

  union all

  select
    'trend'::text as section,
    check_in_year::integer as sort_order,
    check_in_year::text as chart_year,
    null::text as age_segment,
    count(*)::bigint as guest_count,
    round(avg(age_at_check_in)::numeric, 2) as average_age,
    round(avg(stay_nights)::numeric, 2) as average_stay
  from normalized
  group by check_in_year
  order by section, sort_order;
$$;
