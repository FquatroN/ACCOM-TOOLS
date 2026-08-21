-- Keep source-rule replacement compatible with Supabase safe-update guards,
-- which reject DELETE statements that do not include a WHERE clause.

create or replace function public.replace_financial_reconciliation_source_rules(p_rules jsonb)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
begin
  if p_rules is null or jsonb_typeof(p_rules) <> 'array' then
    raise exception 'Reconciliation rules must be an array.';
  end if;
  if exists (
    select 1 from jsonb_array_elements(p_rules) rule
    where jsonb_typeof(rule) <> 'object'
  ) then
    raise exception 'Each reconciliation rule must be an object.';
  end if;
  if exists (
    select 1 from jsonb_array_elements(p_rules) rule
    where coalesce(rule->>'base_source_type', '') not in (
      'financial_documents', 'import_fdm_accounts',
      'import_cgd_cartao_credito', 'import_cgd_extrato_ordem'
    ) or coalesce(rule->>'matching_source_type', '') not in (
      'financial_documents', 'import_fdm_accounts',
      'import_cgd_cartao_credito', 'import_cgd_extrato_ordem'
    )
  ) then
    raise exception 'Rule source type is invalid.';
  end if;
  if exists (
    select 1 from jsonb_array_elements(p_rules) rule
    where rule->>'base_source_type' = rule->>'matching_source_type'
  ) then
    raise exception 'Rule sources must be different.';
  end if;
  if exists (
    select 1 from jsonb_array_elements(p_rules) rule
    where coalesce(rule->>'operator', '') not in ('+', '-')
  ) then
    raise exception 'Rule operator must be ''+'' or ''-''.';
  end if;
  if exists (
    select 1
    from jsonb_to_recordset(p_rules) as rule(
      base_source_type text, matching_source_type text, operator text
    )
    group by rule.base_source_type, rule.matching_source_type
    having count(*) > 1
  ) then
    raise exception 'Duplicate reconciliation rule.';
  end if;
  if (
    select count(*)
    from jsonb_to_recordset(p_rules) as rule(
      base_source_type text, matching_source_type text, operator text
    )
    where rule.base_source_type = 'financial_documents'
      and rule.matching_source_type = 'import_cgd_cartao_credito'
      and rule.operator = '+'
  ) <> 1 then
    raise exception 'The managed Credit Card source rule must remain enabled with operator +.';
  end if;
  if (
    select count(*)
    from jsonb_to_recordset(p_rules) as rule(
      base_source_type text, matching_source_type text, operator text
    )
    where rule.base_source_type = 'financial_documents'
      and rule.matching_source_type = 'import_cgd_extrato_ordem'
      and rule.operator = '+'
  ) <> 1 then
    raise exception 'The managed Bank Statement source rule must remain enabled with operator +.';
  end if;

  lock table public.financial_reconciliation_source_rules in share row exclusive mode;
  delete from public.financial_reconciliation_source_rules
  where base_source_type in (
    'financial_documents', 'import_fdm_accounts',
    'import_cgd_cartao_credito', 'import_cgd_extrato_ordem'
  );

  insert into public.financial_reconciliation_source_rules (
    base_source_type, matching_source_type, operator
  )
  select rule.base_source_type, rule.matching_source_type, rule.operator
  from jsonb_to_recordset(p_rules) as rule(
    base_source_type text, matching_source_type text, operator text
  );

  return coalesce((
    select jsonb_agg(jsonb_build_object(
      'base_source_type', base_source_type,
      'matching_source_type', matching_source_type,
      'operator', operator
    ) order by base_source_type, matching_source_type)
    from public.financial_reconciliation_source_rules
  ), '[]'::jsonb);
end
$$;

revoke all on function public.replace_financial_reconciliation_source_rules(jsonb) from public, anon, authenticated;
grant execute on function public.replace_financial_reconciliation_source_rules(jsonb) to service_role;

notify pgrst, 'reload schema';
