-- Orders eligible reconciliation candidates from oldest to newest before pagination.
do $fix$
declare
  definition text;
  original_definition text;
  old_page_order constant text := $$order by source_date desc offset v_offset limit p_page_size$$;
  new_page_order constant text := $$order by source_date asc, id asc offset v_offset limit p_page_size$$;
  old_json_order constant text := $$order by x.source_date desc$$;
  new_json_order constant text := $$order by x.source_date asc, x.id asc$$;
begin
  select pg_get_functiondef('public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)'::regprocedure)
    into definition;

  original_definition := definition;

  if position(old_page_order in definition) > 0 then
    definition := replace(definition, old_page_order, new_page_order);
  end if;
  if position(old_json_order in definition) > 0 then
    definition := replace(definition, old_json_order, new_json_order);
  end if;

  if position(new_page_order in definition) = 0
     or position(new_json_order in definition) = 0
     or position(old_page_order in definition) > 0
     or position(old_json_order in definition) > 0 then
    raise exception 'Could not verify deterministic oldest-first candidate ordering in the reconciliation workspace function.';
  end if;

  if definition <> original_definition then
    execute definition;
  end if;
end $fix$;
