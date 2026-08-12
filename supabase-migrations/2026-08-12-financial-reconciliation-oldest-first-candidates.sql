-- Orders eligible reconciliation candidates from oldest to newest before pagination.
do $fix$
declare
  definition text;
  original_definition text;
  old_page_order constant text := $$order by source_date desc offset v_offset limit p_page_size$$;
  new_page_order constant text := $$order by source_date asc, id asc offset v_offset limit p_page_size$$;
  old_json_order constant text := $$order by x.source_date desc$$;
  new_json_order constant text := $$order by x.source_date asc, x.id asc$$;
  old_page_count integer;
  new_page_count integer;
  old_json_count integer;
  new_json_count integer;
begin
  select pg_get_functiondef('public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)'::regprocedure)
    into definition;

  original_definition := definition;

  old_page_count := (length(definition) - length(replace(definition, old_page_order, ''))) / length(old_page_order);
  new_page_count := (length(definition) - length(replace(definition, new_page_order, ''))) / length(new_page_order);
  old_json_count := (length(definition) - length(replace(definition, old_json_order, ''))) / length(old_json_order);
  new_json_count := (length(definition) - length(replace(definition, new_json_order, ''))) / length(new_json_order);

  if not (
    (old_page_count = 1
     and new_page_count = 0
     and old_json_count = 1
     and new_json_count = 0)
    or
    (old_page_count = 0
     and new_page_count = 1
     and old_json_count = 0
     and new_json_count = 1)
  ) then
    raise exception 'Unexpected reconciliation workspace function definition; could not verify deterministic oldest-first candidate ordering.';
  end if;

  if old_page_count = 1 then
    definition := replace(definition, old_page_order, new_page_order);
    definition := replace(definition, old_json_order, new_json_order);
  end if;

  old_page_count := (length(definition) - length(replace(definition, old_page_order, ''))) / length(old_page_order);
  new_page_count := (length(definition) - length(replace(definition, new_page_order, ''))) / length(new_page_order);
  old_json_count := (length(definition) - length(replace(definition, old_json_order, ''))) / length(old_json_order);
  new_json_count := (length(definition) - length(replace(definition, new_json_order, ''))) / length(new_json_order);

  if old_page_count <> 0
     or new_page_count <> 1
     or old_json_count <> 0
     or new_json_count <> 1 then
    raise exception 'Unexpected reconciliation workspace function definition; could not verify deterministic oldest-first candidate ordering.';
  end if;

  if definition <> original_definition then
    execute definition;
  end if;
end $fix$;
