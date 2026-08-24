-- Completed automatic proposals are immutable audit history, not active locks.
-- The reconciliation item table remains the authority for completed records
-- that are still locked by an existing reconciliation.

do $migration$
declare
  v_signature text;
  v_oid oid;
  v_source text;
  v_old_count integer;
  v_new_count integer;
  v_old_pattern constant text :=
    $pattern$overlap_proposal\.status\s+in\s*\(\s*'proposed'\s*,\s*'executing'\s*,\s*'completed'\s*\)$pattern$;
  v_new_pattern constant text :=
    $pattern$overlap_proposal\.status\s+in\s*\(\s*'proposed'\s*,\s*'executing'\s*\)$pattern$;
  v_new_fragment constant text := $fragment$overlap_proposal.status in (
            'proposed','executing'
          )$fragment$;
begin
  foreach v_signature in array array[
    'public.financial_reconciliation_execute_bank_reservation_proposal(uuid,text)',
    'public.financial_reconciliation_execute_adyen_monthly_proposal(uuid,text)'
  ] loop
    v_oid := to_regprocedure(v_signature);
    if v_oid is null then
      raise exception 'Required automatic reconciliation executor is missing: %.',
        v_signature;
    end if;

    v_source := pg_get_functiondef(v_oid);
    select count(*) into v_old_count
    from regexp_matches(v_source, v_old_pattern, 'g');
    select count(*) into v_new_count
    from regexp_matches(v_source, v_new_pattern, 'g');

    if v_old_count = 1 and v_new_count = 0 then
      v_source := regexp_replace(
        v_source, v_old_pattern, v_new_fragment, 'g'
      );
      execute v_source;
    elsif v_old_count = 0 and v_new_count = 1 then
      null;
    else
      raise exception
        'Unexpected automatic overlap predicate definition in % (old %, current %).',
        v_signature, v_old_count, v_new_count;
    end if;

    v_source := pg_get_functiondef(v_oid);
    select count(*) into v_old_count
    from regexp_matches(v_source, v_old_pattern, 'g');
    select count(*) into v_new_count
    from regexp_matches(v_source, v_new_pattern, 'g');
    if v_old_count <> 0 or v_new_count <> 1 then
      raise exception
        'Completed automatic proposal history still blocks unlocked records in %.',
        v_signature;
    end if;
  end loop;
end
$migration$;
