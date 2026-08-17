-- Combines the Financial Documents description, supplier name, and supplier NIF search.
-- Other source types keep their existing description-only behavior.
do $combined_search$
declare
  definition text;
  original_definition text;
  old_description constant text := $$and (nullif(p_filters->>'description','') is null or s.description ilike '%' || (p_filters->>'description') || '%')$$;
  new_description constant text := $$and (
  nullif(p_filters->>'description','') is null
  or s.description ilike '%' || (p_filters->>'description') || '%'
  or (
    p_source_type = 'financial_documents'
    and (
      s.supplier ilike '%' || (p_filters->>'description') || '%'
      or s.supplier_nif ilike '%' || (p_filters->>'description') || '%'
    )
  )
)$$;
  old_filter_fields constant text := $$jsonb_build_array('dateFrom','dateTo','amountMin','amountMax','description','supplier','payment','category')$$;
  new_filter_fields constant text := $$jsonb_build_array('dateFrom','dateTo','amountMin','amountMax','description','payment','category')$$;
  old_description_count integer;
  new_description_count integer;
  old_filter_fields_count integer;
  new_filter_fields_count integer;
begin
  select pg_get_functiondef('public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)'::regprocedure)
    into strict definition;
  original_definition := definition;

  old_description_count := (length(definition) - length(replace(definition, old_description, '')))
    / length(old_description);
  new_description_count := (length(definition) - length(replace(definition, new_description, '')))
    / length(new_description);
  old_filter_fields_count := (length(definition) - length(replace(definition, old_filter_fields, '')))
    / length(old_filter_fields);
  new_filter_fields_count := (length(definition) - length(replace(definition, new_filter_fields, '')))
    / length(new_filter_fields);

  if old_description_count = 2 and new_description_count = 0
     and old_filter_fields_count = 1 and new_filter_fields_count = 0 then
    definition := replace(definition, old_description, new_description);
    definition := replace(definition, old_filter_fields, new_filter_fields);
  elsif old_description_count = 0 and new_description_count = 2
        and old_filter_fields_count = 0 and new_filter_fields_count = 1 then
    null;
  else
    raise exception 'Unexpected reconciliation workspace function definition; could not install combined search.';
  end if;

  old_description_count := (length(definition) - length(replace(definition, old_description, '')))
    / length(old_description);
  new_description_count := (length(definition) - length(replace(definition, new_description, '')))
    / length(new_description);
  old_filter_fields_count := (length(definition) - length(replace(definition, old_filter_fields, '')))
    / length(old_filter_fields);
  new_filter_fields_count := (length(definition) - length(replace(definition, new_filter_fields, '')))
    / length(new_filter_fields);

  if old_description_count <> 0 or new_description_count <> 2
     or old_filter_fields_count <> 0 or new_filter_fields_count <> 1 then
    raise exception 'Unexpected reconciliation workspace function definition; could not verify combined search.';
  end if;

  if definition <> original_definition then
    execute definition;
  end if;
end $combined_search$;

revoke all on function public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer) from public, anon, authenticated;
grant execute on function public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer) to service_role;
notify pgrst, 'reload schema';
