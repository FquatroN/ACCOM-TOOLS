-- Repairs the workspace function after the source-rules migration was applied
-- with unparenthesized jsonb text extraction in ILIKE filters.
do $fix$
declare
  definition text;
  bad_description constant text := $$s.description ilike '%' || p_filters->>'description' || '%'$$;
  good_description constant text := $$s.description ilike '%' || (p_filters->>'description') || '%'$$;
  bad_supplier constant text := $$s.supplier ilike '%' || p_filters->>'supplier' || '%'$$;
  good_supplier constant text := $$s.supplier ilike '%' || (p_filters->>'supplier') || '%'$$;
begin
  select pg_get_functiondef('public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)'::regprocedure)
    into definition;

  definition := replace(definition, bad_description, good_description);
  definition := replace(definition, bad_supplier, good_supplier);

  if position(good_description in definition) = 0 or position(good_supplier in definition) = 0 then
    raise exception 'Source-rules workspace filter fix could not verify the function body.';
  end if;

  execute definition;
end $fix$;
