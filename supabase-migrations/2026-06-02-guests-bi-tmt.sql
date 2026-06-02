create or replace function public.guests_bi_years()
returns table(year integer)
language sql
stable
as $$
  select distinct extract(year from gr.check_in)::int as year
  from public.guest_records gr
  where gr.check_in is not null
  order by year desc;
$$;

create or replace function public.guests_bi_tmt(p_year integer default null)
returns table(
  year_month text,
  total_nights bigint,
  exempt_7days bigint,
  exempt_13_year bigint
)
language sql
stable
as $$
  with selected_year as (
    select coalesce(
      p_year,
      (
        select max(extract(year from gr.check_in))::int
        from public.guest_records gr
        where gr.check_in is not null
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
    where gr.check_in is not null
      and gr.check_out is not null
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
