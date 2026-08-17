create or replace function public.get_bi_financial_dif_analysis_payload(
  p_year_from integer default extract(year from current_date)::integer - 1,
  p_year_to integer default extract(year from current_date)::integer
)
returns table (rows jsonb)
language sql
stable
as $$
  select coalesce(
    jsonb_agg(
      to_jsonb(aggregate_row)
        || jsonb_build_object('difference', aggregate_row.extrato_amount - aggregate_row.fdm_amount)
      order by aggregate_row.analysis, aggregate_row.year, aggregate_row.month
    ),
    '[]'::jsonb
  ) as rows
  from public.get_bi_financial_dif_analysis(p_year_from, p_year_to) aggregate_row;
$$;

grant execute on function public.get_bi_financial_dif_analysis_payload(integer, integer)
  to anon, authenticated, service_role;
