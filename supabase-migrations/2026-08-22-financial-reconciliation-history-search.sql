-- Dedicated, paginated reconciliation history for the History workbench tab.
create or replace function public.get_financial_reconciliation_history(
  p_created_from date default null,
  p_created_to date default null,
  p_origin text default null,
  p_status text default null,
  p_difference_from numeric default null,
  p_difference_to numeric default null,
  p_page integer default 1,
  p_page_size integer default 50
)
returns jsonb
language plpgsql
stable
security definer
set search_path = public
as $$
declare
  v_offset integer;
  v_total bigint;
  v_rows jsonb;
begin
  if p_origin is not null and p_origin not in ('user', 'automatic') then
    raise exception 'History origin is invalid.';
  end if;
  if p_status is not null and p_status not in ('not_started', 'started', 'complete') then
    raise exception 'History status is invalid.';
  end if;
  if p_created_from is not null and p_created_to is not null and p_created_from > p_created_to then
    raise exception 'History created date range is invalid.';
  end if;
  if p_difference_from is not null and p_difference_to is not null and p_difference_from > p_difference_to then
    raise exception 'History difference range is invalid.';
  end if;
  if p_page < 1 or p_page_size < 1 or p_page_size > 100 then
    raise exception 'History pagination is invalid.';
  end if;

  v_offset := (p_page - 1) * p_page_size;

  with filtered as (
    select reconciliation.*
    from public.financial_reconciliations reconciliation
    where reconciliation.deleted_at is null
      and (p_created_from is null or reconciliation.created_at >= (p_created_from::timestamp at time zone 'Europe/Lisbon'))
      and (p_created_to is null or reconciliation.created_at < ((p_created_to + 1)::timestamp at time zone 'Europe/Lisbon'))
      and (p_origin is null or reconciliation.origin = p_origin)
      and (
        p_status is null
        or (p_status = 'not_started' and false)
        or reconciliation.status = p_status
      )
      and (p_difference_from is null or reconciliation.difference_amount >= p_difference_from)
      and (p_difference_to is null or reconciliation.difference_amount <= p_difference_to)
  )
  select count(*) into v_total from filtered;

  with filtered as (
    select reconciliation.*
    from public.financial_reconciliations reconciliation
    where reconciliation.deleted_at is null
      and (p_created_from is null or reconciliation.created_at >= (p_created_from::timestamp at time zone 'Europe/Lisbon'))
      and (p_created_to is null or reconciliation.created_at < ((p_created_to + 1)::timestamp at time zone 'Europe/Lisbon'))
      and (p_origin is null or reconciliation.origin = p_origin)
      and (
        p_status is null
        or (p_status = 'not_started' and false)
        or reconciliation.status = p_status
      )
      and (p_difference_from is null or reconciliation.difference_amount >= p_difference_from)
      and (p_difference_to is null or reconciliation.difference_amount <= p_difference_to)
    order by reconciliation.created_at desc, reconciliation.id desc
    offset v_offset limit p_page_size
  )
  select coalesce(jsonb_agg(
    to_jsonb(filtered)
    || jsonb_build_object(
      'sourceSummary', coalesce(summary.source_summary, '[]'::jsonb),
      'totalRecords', coalesce(summary.total_records, 0),
      'sourceAmountTotal', coalesce(summary.source_amount_total, 0),
      'destinationAmountTotal', coalesce(summary.destination_amount_total, 0),
      'completionComment', completion.comment
    ) order by filtered.created_at desc, filtered.id desc
  ), '[]'::jsonb)
  into v_rows
  from filtered
  left join lateral (
    select
      jsonb_agg(
        jsonb_build_object(
          'sourceType', grouped.source_type,
          'recordCount', grouped.record_count,
          'amountTotal', grouped.amount_total
        ) order by case when grouped.source_type = filtered.base_source_type then 0 else 1 end, grouped.source_type
      ) as source_summary,
      sum(grouped.record_count)::integer as total_records,
      coalesce(sum(grouped.amount_total) filter (where grouped.source_type = filtered.base_source_type), 0) as source_amount_total,
      coalesce(sum(grouped.amount_total) filter (where grouped.source_type <> filtered.base_source_type), 0) as destination_amount_total
    from (
      select item.source_type, count(*)::integer as record_count, sum(item.amount_snapshot)::numeric as amount_total
      from public.financial_reconciliation_items item
      where item.reconciliation_id = filtered.id
      group by item.source_type
    ) grouped
  ) summary on true
  left join lateral (
    select audit.comment
    from public.financial_reconciliation_audit audit
    where audit.reconciliation_id = filtered.id
      and audit.action in ('complete', 'force_complete', 'automatic_complete')
    order by audit.created_at desc, audit.id desc
    limit 1
  ) completion on true;

  return jsonb_build_object(
    'rows', v_rows,
    'page', p_page,
    'pageSize', p_page_size,
    'total', v_total
  );
end $$;

revoke all on function public.get_financial_reconciliation_history(date,date,text,text,numeric,numeric,integer,integer)
  from public, anon, authenticated;
grant execute on function public.get_financial_reconciliation_history(date,date,text,text,numeric,numeric,integer,integer)
  to service_role;

notify pgrst, 'reload schema';
