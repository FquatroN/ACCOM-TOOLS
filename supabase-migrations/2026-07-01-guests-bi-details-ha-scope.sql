drop function if exists public.guests_bi_details(integer, text);

create or replace function public.guests_bi_details(
  p_year integer,
  p_ha text default null
)
returns table (
  section text,
  sort_order integer,
  ha_scope text,
  chart_year text,
  age_segment text,
  guest_count bigint,
  average_age numeric,
  average_stay numeric
)
language sql
stable
as $$
  with scope_seed as (
    select *
    from (values
      (0, 'All'::text, null::text),
      (1, 'H'::text, 'H'::text),
      (2, 'A'::text, 'A'::text)
    ) as seed(sort_order, ha_scope, ha_value)
    where nullif(btrim(coalesce(p_ha, '')), '') is null
       or seed.ha_scope = upper(btrim(p_ha))
  ),
  normalized as (
    select
      upper(btrim(coalesce(gr.ha, ''))) as ha,
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
      and (nullif(btrim(coalesce(p_ha, '')), '') is null or upper(btrim(coalesce(gr.ha, ''))) = upper(btrim(p_ha)))
  ),
  scoped as (
    select
      scope.sort_order as scope_sort_order,
      scope.ha_scope,
      normalized.*
    from scope_seed scope
    join normalized on scope.ha_value is null or normalized.ha = scope.ha_value
  ),
  selected as (
    select *
    from scoped
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
      ha_scope,
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
    group by 1, 2
  )
  select
    'summary'::text as section,
    scope.sort_order::integer as sort_order,
    scope.ha_scope,
    p_year::text as chart_year,
    null::text as age_segment,
    count(selected.check_in)::bigint as guest_count,
    round(avg(selected.age_at_check_in)::numeric, 2) as average_age,
    round(avg(selected.stay_nights)::numeric, 2) as average_stay
  from scope_seed scope
  left join selected on selected.ha_scope = scope.ha_scope
  group by scope.sort_order, scope.ha_scope

  union all

  select
    'segment'::text as section,
    (scope.sort_order * 100 + seed.sort_order)::integer as sort_order,
    scope.ha_scope,
    p_year::text as chart_year,
    seed.age_segment,
    coalesce(seg.guest_count, 0)::bigint as guest_count,
    null::numeric as average_age,
    null::numeric as average_stay
  from scope_seed scope
  cross join segment_seed seed
  left join selected_segments seg on seg.ha_scope = scope.ha_scope and seg.age_segment = seed.age_segment

  union all

  select
    'trend'::text as section,
    (scope_sort_order * 10000 + check_in_year)::integer as sort_order,
    ha_scope,
    check_in_year::text as chart_year,
    null::text as age_segment,
    count(*)::bigint as guest_count,
    round(avg(age_at_check_in)::numeric, 2) as average_age,
    round(avg(stay_nights)::numeric, 2) as average_stay
  from scoped
  group by scope_sort_order, ha_scope, check_in_year
  order by section, sort_order;
$$;
