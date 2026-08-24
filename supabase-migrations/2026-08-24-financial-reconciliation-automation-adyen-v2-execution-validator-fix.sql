-- Keep the Adyen v2 execution validator aligned with the immutable definition
-- stored by analysis. The original category-exclusion migration used an exact
-- whitespace-sensitive replacement and could leave this one local constant on
-- the v1 definition even though the rest of the executor was upgraded to v2.

do $migration$
declare
  v_procedure constant regprocedure :=
    'public.financial_reconciliation_execute_adyen_monthly_proposal(uuid,text)'::regprocedure;
  v_source text;
  v_original text;
begin
  v_source := pg_get_functiondef(v_procedure);
  v_original := v_source;

  if position(
      $expected$'fdmExcludedCategory', 'TransferOutToAccount'$expected$
      in v_source
    ) = 0 then
    v_source := regexp_replace(
      v_source,
      '(''fdmAccount''\s*,\s*''Adyen''\s*,)(\s*)(''requiresBothSides''\s*,\s*true\s*,)',
      E'\\1\\2''fdmExcludedCategory'', ''TransferOutToAccount'',\\2\\3',
      'g'
    );
  end if;

  if position(
      $expected$'fdmExcludedCategory', 'TransferOutToAccount'$expected$
      in v_source
    ) = 0 then
    raise exception
      'Adyen v2 execution validator definition could not be upgraded safely.';
  end if;

  if v_source is distinct from v_original then
    execute v_source;
  end if;

  v_source := pg_get_functiondef(v_procedure);
  if position(
      $expected$'fdmExcludedCategory', 'TransferOutToAccount'$expected$
      in v_source
    ) = 0
    or position(
      $expected$v_proposal.rule_version <> 2$expected$
      in v_source
    ) = 0
    or position(
      $expected$and fdm.category <> 'TransferOutToAccount'$expected$
      in v_source
    ) = 0 then
    raise exception 'Adyen v2 execution validator upgrade verification failed.';
  end if;
end
$migration$;

notify pgrst, 'reload schema';
