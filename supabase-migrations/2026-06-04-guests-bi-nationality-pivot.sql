create or replace function public.guests_bi_nationality_key(p_nationality text, p_nationality_code text)
returns text
language sql
immutable
as $$
  select coalesce(
    nullif(upper(trim(coalesce(public.guest_country_canonical_code(p_nationality_code), ''))), ''),
    nullif(upper(trim(coalesce(public.guest_country_resolve_code(p_nationality), ''))), ''),
    public.guest_country_lookup_key(p_nationality),
    'UNKNOWN'
  );
$$;

create or replace function public.guests_bi_nationality_label(p_nationality text, p_nationality_code text)
returns text
language sql
stable
as $$
  select coalesce(
    nullif(public.guest_country_designacao(p_nationality_code), ''),
    nullif(public.guest_country_designacao(public.guest_country_resolve_code(p_nationality)), ''),
    nullif(trim(coalesce(p_nationality, '')), ''),
    'Unknown'
  );
$$;

create or replace function public.guests_bi_nationality_pivot(p_ha text default null)
returns table(
  country_label text,
  chart_year integer,
  guest_count bigint,
  row_total bigint
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
  totals as (
    select
      country_key,
      min(country_label) as country_label,
      sum(guest_count)::bigint as row_total
    from base_rows
    group by country_key
  )
  select
    t.country_label,
    b.chart_year,
    b.guest_count,
    t.row_total
  from base_rows b
  join totals t on t.country_key = b.country_key
  order by t.row_total desc, t.country_label asc, b.chart_year desc;
$$;
