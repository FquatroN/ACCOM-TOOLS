create or replace function public.guests_bi_ine(p_year_month text, p_ha text default null)
returns table(
  section text,
  row_label text,
  guests_entered integer,
  guests_slept integer,
  nights integer,
  sort_order integer
)
language sql
stable
as $$
with params as (
  select
    to_date(p_year_month || '-01', 'YYYY-MM-DD') as month_start,
    (to_date(p_year_month || '-01', 'YYYY-MM-DD') + interval '1 month')::date as month_end,
    nullif(trim(upper(coalesce(p_ha, ''))), '') as ha_filter
),
base as (
  select
    coalesce(nullif(trim(gr.nationality), ''), nullif(trim(gr.nationality_code), ''), 'Unknown') as nationality_label,
    coalesce(nullif(trim(gr.issuer_country), ''), nullif(trim(gr.issuer_country_code), ''), 'Unknown') as issuer_label,
    upper(trim(coalesce(gr.nationality_code, ''))) as nationality_code,
    upper(trim(coalesce(gr.issuer_country_code, ''))) as issuer_country_code,
    case
      when gr.check_in >= p.month_start and gr.check_in < p.month_end then 1
      else 0
    end as guests_entered,
    greatest(
      least(gr.check_out, p.month_end) - greatest(gr.check_in, p.month_start),
      0
    )::int as nights
  from public.guest_records gr
  cross join params p
  where gr.check_in is not null
    and gr.check_out is not null
    and gr.check_in < p.month_end
    and gr.check_out > p.month_start
    and (p.ha_filter is null or upper(trim(coalesce(gr.ha, ''))) = p.ha_filter)
),
classified as (
  select
    *,
    case
      when issuer_country_code in ('PRT', 'PTR') or upper(issuer_label) = 'PORTUGAL' then true
      else false
    end as issuer_portugal,
    case
      when nationality_code in ('PRT', 'PTR') or upper(nationality_label) = 'PORTUGAL' then true
      else false
    end as nationality_portugal,
    case
      when nights > 0 then 1
      else 0
    end as guests_slept
  from base
),
summary as (
  select
    'summary'::text as section,
    'Portugueses residentes em Portugal'::text as row_label,
    coalesce(sum(guests_entered) filter (where issuer_portugal and nationality_portugal), 0)::int as guests_entered,
    coalesce(sum(guests_slept) filter (where issuer_portugal and nationality_portugal), 0)::int as guests_slept,
    coalesce(sum(nights) filter (where issuer_portugal and nationality_portugal), 0)::int as nights,
    1 as sort_order
  from classified
  union all
  select
    'summary'::text as section,
    'Estrangeiros residentes em Portugal'::text as row_label,
    coalesce(sum(guests_entered) filter (where issuer_portugal and not nationality_portugal), 0)::int as guests_entered,
    coalesce(sum(guests_slept) filter (where issuer_portugal and not nationality_portugal), 0)::int as guests_slept,
    coalesce(sum(nights) filter (where issuer_portugal and not nationality_portugal), 0)::int as nights,
    2 as sort_order
  from classified
),
detail as (
  select
    'detail'::text as section,
    nationality_label as row_label,
    sum(guests_entered)::int as guests_entered,
    sum(guests_slept)::int as guests_slept,
    sum(nights)::int as nights,
    row_number() over (order by sum(nights) desc, nationality_label) + 100 as sort_order
  from classified
  where not nationality_portugal
    and nights > 0
    and nullif(trim(nationality_label), '') is not null
  group by nationality_label
)
select section, row_label, guests_entered, guests_slept, nights, sort_order
from summary
union all
select section, row_label, guests_entered, guests_slept, nights, sort_order
from detail
order by sort_order, row_label;
$$;
