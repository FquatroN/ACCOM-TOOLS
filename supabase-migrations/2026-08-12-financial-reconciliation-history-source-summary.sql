-- Adds raw per-source record counts and totals to reconciliation history rows.
do $fix$
declare
  definition text;
  original_definition text;
  old_history constant text := $$'history',coalesce((select jsonb_agg(to_jsonb(h) order by h.created_at desc) from (select * from public.financial_reconciliations where deleted_at is null order by created_at desc limit 100) h),'[]'::jsonb)$$;
  new_history constant text := $$'history',coalesce((
    select jsonb_agg(
      to_jsonb(h) || jsonb_build_object('sourceSummary',coalesce(summary.source_summary,'[]'::jsonb))
      order by h.created_at desc
    )
    from (
      select *
      from public.financial_reconciliations
      where deleted_at is null
      order by created_at desc
      limit 100
    ) h
    left join lateral (
      select jsonb_agg(
        jsonb_build_object(
          'sourceType',grouped.source_type,
          'recordCount',grouped.record_count,
          'amountTotal',grouped.amount_total
        )
        order by
          case when grouped.source_type=h.base_source_type then 0 else 1 end,
          coalesce((
            select matching.position
            from jsonb_array_elements_text(h.matching_source_types)
              with ordinality matching(source_type,position)
            where matching.source_type=grouped.source_type
            limit 1
          ),2147483647),
          grouped.source_type
      ) as source_summary
      from (
        select
          i.source_type,
          count(*) as record_count,
          coalesce(sum(i.amount_snapshot),0) as amount_total
        from public.financial_reconciliation_items i
        where i.reconciliation_id=h.id
        group by i.source_type
      ) grouped
    ) summary on true
  ),'[]'::jsonb)$$;
  old_history_count integer;
  new_history_count integer;
begin
  select pg_get_functiondef('public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)'::regprocedure) into definition;
  original_definition := definition;

  old_history_count := (length(definition)-length(replace(definition,old_history,''))) / length(old_history);
  new_history_count := (length(definition)-length(replace(definition,new_history,''))) / length(new_history);

  if not (
    (old_history_count = 1 and new_history_count = 0)
    or (old_history_count = 0 and new_history_count = 1)
  ) then
    raise exception 'Unexpected reconciliation workspace function definition; could not install history source summaries.';
  end if;

  if old_history_count = 1 then
    definition := replace(definition,old_history,new_history);
  end if;

  old_history_count := (length(definition)-length(replace(definition,old_history,''))) / length(old_history);
  new_history_count := (length(definition)-length(replace(definition,new_history,''))) / length(new_history);
  if old_history_count <> 0 or new_history_count <> 1 then
    raise exception 'Unexpected reconciliation workspace function definition; could not verify history source summaries.';
  end if;

  if definition <> original_definition then
    execute definition;
  end if;
end $fix$;
