-- Adds source-specific filter LOV metadata without replacing later workspace enhancements.
-- The installed definition retains its security definer search path and result envelope.
do $filter_lovs$
declare
  definition text;
  original_definition text;
  old_snippets constant text[] := array[
    $$declare v_rec public.financial_reconciliations%rowtype; v_rules jsonb; v_offset int; v_candidates jsonb; v_count int; v_started_count int; v_complete_count int;$$,
    $$v_offset := (p_page - 1) * p_page_size;
  with candidate_rows as ($$,
    $$and (nullif(p_filters->>'supplier','') is null or s.supplier ilike '%' || (p_filters->>'supplier') || '%')$$,
    $$and (nullif(p_filters->>'payment','') is null or s.payment = p_filters->>'payment')$$,
    $$and (nullif(p_filters->>'account','') is null or s.account = p_filters->>'account')$$,
    $$and (nullif(p_filters->>'category','') is null or s.category = p_filters->>'category')$$,
    $$'amountColumn','amount','columns'$$,
    $$jsonb_build_array('dateFrom','dateTo','amountMin','amountMax','description','supplier','payment',$$ ||
      $$'account','category')$$
  ];
  new_snippets constant text[] := array[
    $$declare v_rec public.financial_reconciliations%rowtype; v_rules jsonb; v_filter_options jsonb := '{}'::jsonb; v_offset int; v_candidates jsonb; v_count int; v_started_count int; v_complete_count int;$$,
    $$v_offset := (p_page - 1) * p_page_size;
  if p_source_type = 'financial_documents' then
    select jsonb_build_object(
      'payment', coalesce((
        select jsonb_agg(option_value order by lower(option_value),option_value)
        from (
          select distinct btrim(payment) as option_value
          from public.financial_documents
          where nullif(btrim(payment), '') is not null
        ) options
      ), '[]'::jsonb),
      'category', coalesce((
        select jsonb_agg(option_value order by lower(option_value),option_value)
        from (
          select distinct btrim(category) as option_value
          from public.financial_documents
          where nullif(btrim(category), '') is not null
        ) options
      ), '[]'::jsonb)
    ) into v_filter_options;
  elsif p_source_type = 'import_fdm_accounts' then
    select jsonb_build_object(
      'account', coalesce((
        select jsonb_agg(option_value order by lower(option_value),option_value)
        from (
          select distinct btrim(account) as option_value
          from public.import_fdm_accounts
          where nullif(btrim(account), '') is not null
        ) options
      ), '[]'::jsonb),
      'category', coalesce((
        select jsonb_agg(option_value order by lower(option_value),option_value)
        from (
          select distinct btrim(category) as option_value
          from public.import_fdm_accounts
          where nullif(btrim(category), '') is not null
        ) options
      ), '[]'::jsonb)
    ) into v_filter_options;
  end if;
  with candidate_rows as ($$,
    $$and (
  nullif(p_filters->>'supplier','') is null
  or s.supplier ilike '%' || (p_filters->>'supplier') || '%'
  or s.supplier_nif ilike '%' || (p_filters->>'supplier') || '%'
)$$,
    $$and (nullif(p_filters->>'payment','') is null or btrim(s.payment) = p_filters->>'payment')$$,
    $$and (nullif(p_filters->>'account','') is null or btrim(s.account) = p_filters->>'account')$$,
    $$and (nullif(p_filters->>'category','') is null or btrim(s.category) = p_filters->>'category')$$,
    $$'amountColumn','amount','filterOptions',v_filter_options,'columns'$$,
    $$jsonb_build_array('dateFrom','dateTo','amountMin','amountMax','description','supplier','payment','category')$$
  ];
  expected_counts constant integer[] := array[1, 1, 2, 2, 2, 2, 1, 1];
  old_count integer;
  new_count integer;
  snippet_index integer;
begin
  select pg_get_functiondef('public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)'::regprocedure)
    into strict definition;
  original_definition := definition;

  for snippet_index in 1..array_length(old_snippets, 1) loop
    old_count := (length(definition) - length(replace(definition, old_snippets[snippet_index], '')))
      / length(old_snippets[snippet_index]);
    new_count := (length(definition) - length(replace(definition, new_snippets[snippet_index], '')))
      / length(new_snippets[snippet_index]);

    if not (
      (old_count = expected_counts[snippet_index] and new_count = 0)
      or (old_count = 0 and new_count = expected_counts[snippet_index])
    ) then
      raise exception 'Unexpected reconciliation workspace function definition; could not install filter LOVs.';
    end if;

    if old_count = expected_counts[snippet_index] then
      definition := replace(definition, old_snippets[snippet_index], new_snippets[snippet_index]);
    end if;
  end loop;

  for snippet_index in 1..array_length(old_snippets, 1) loop
    old_count := (length(definition) - length(replace(definition, old_snippets[snippet_index], '')))
      / length(old_snippets[snippet_index]);
    new_count := (length(definition) - length(replace(definition, new_snippets[snippet_index], '')))
      / length(new_snippets[snippet_index]);

    if old_count <> 0 or new_count <> expected_counts[snippet_index] then
      raise exception 'Unexpected reconciliation workspace function definition; could not install filter LOVs.';
    end if;
  end loop;

  if definition <> original_definition then
    execute definition;
  end if;
end $filter_lovs$;

revoke all on function public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer) from public, anon, authenticated;
grant execute on function public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer) to service_role;
notify pgrst, 'reload schema';
