\set ON_ERROR_STOP on

begin;

\ir ../supabase-migrations/2026-08-14-financial-reconciliation-automation-schema.sql
\ir ../supabase-migrations/2026-08-14-financial-reconciliation-automation-analysis.sql

-- definition/config preservation
update public.financial_reconciliation_automatic_rule_definitions
set display_name = 'tampered definition'
where rule_key = 'financial_documents_cgd_bank_statement' and version = 1;

update public.financial_reconciliation_automatic_rule_configs
set enabled = false,
    allow_manual_execution = true,
    include_in_scheduled_batch = true,
    difference_allowed = 12.34,
    max_difference_days = 21,
    priority = 1,
    updated_by = 'smoke:administrator'
where rule_key = 'financial_documents_cgd_bank_statement';

update public.financial_reconciliation_automatic_schedule
set enabled = true,
    time_of_day = '04:30',
    updated_by = 'smoke:administrator'
where id = true;

\ir ../supabase-migrations/2026-08-14-financial-reconciliation-automation-schema.sql
\ir ../supabase-migrations/2026-08-14-financial-reconciliation-automation-analysis.sql

do $$
declare
  v_definition public.financial_reconciliation_automatic_rule_definitions%rowtype;
  v_config public.financial_reconciliation_automatic_rule_configs%rowtype;
  v_schedule public.financial_reconciliation_automatic_schedule%rowtype;
begin
  select * into strict v_definition
  from public.financial_reconciliation_automatic_rule_definitions
  where rule_key = 'financial_documents_cgd_bank_statement' and version = 1;
  if v_definition.display_name <> 'Financial Documents to CGD Bank Statement'
    or v_definition.definition->>'maxIdentityCandidatesPerBase' <> '12' then
    raise exception 'Managed definition was not restored by migration reapply.';
  end if;

  select * into strict v_config
  from public.financial_reconciliation_automatic_rule_configs
  where rule_key = 'financial_documents_cgd_bank_statement';
  if v_config.enabled
    or not v_config.allow_manual_execution
    or not v_config.include_in_scheduled_batch
    or v_config.difference_allowed <> 12.34
    or v_config.max_difference_days <> 21
    or v_config.updated_by <> 'smoke:administrator' then
    raise exception 'Administrator rule configuration was overwritten by migration reapply.';
  end if;

  select * into strict v_schedule
  from public.financial_reconciliation_automatic_schedule
  where id = true;
  if not v_schedule.enabled
    or v_schedule.time_of_day <> time '04:30'
    or v_schedule.updated_by <> 'smoke:administrator' then
    raise exception 'Administrator schedule was overwritten by migration reapply.';
  end if;
end $$;

-- RLS and privileges
-- RPC-only privileges
do $$
declare
  v_table text;
begin
  foreach v_table in array array[
    'financial_reconciliation_automatic_rule_definitions',
    'financial_reconciliation_automatic_rule_configs',
    'financial_reconciliation_automatic_schedule',
    'financial_reconciliation_automatic_runs',
    'financial_reconciliation_automatic_proposals'
  ] loop
    if not exists (
      select 1 from pg_class c join pg_namespace n on n.oid = c.relnamespace
      where n.nspname = 'public' and c.relname = v_table and c.relrowsecurity
    ) then
      raise exception 'RLS is not enabled for %.', v_table;
    end if;
    if has_table_privilege('anon', format('public.%I', v_table), 'SELECT')
      or has_table_privilege('authenticated', format('public.%I', v_table), 'SELECT') then
      raise exception 'Application roles retain direct privileges on %.', v_table;
    end if;
    if has_table_privilege('service_role', format('public.%I', v_table), 'SELECT')
      or has_table_privilege('service_role', format('public.%I', v_table), 'INSERT')
      or has_table_privilege('service_role', format('public.%I', v_table), 'UPDATE')
      or has_table_privilege('service_role', format('public.%I', v_table), 'DELETE') then
      raise exception 'service_role retains direct table privileges on %.', v_table;
    end if;
  end loop;

  if has_function_privilege('anon', 'public.get_financial_reconciliation_automation_settings()', 'EXECUTE')
    or has_function_privilege('authenticated', 'public.replace_financial_reconciliation_automation_settings(jsonb,jsonb,text)', 'EXECUTE')
    or not has_function_privilege('service_role', 'public.get_financial_reconciliation_automation_settings()', 'EXECUTE')
    or not has_function_privilege('service_role', 'public.replace_financial_reconciliation_automation_settings(jsonb,jsonb,text)', 'EXECUTE') then
    raise exception 'Automation RPC privileges are invalid.';
  end if;
  if has_function_privilege('anon', 'public.financial_reconciliation_match_normalize(text)', 'EXECUTE')
    or has_function_privilege('authenticated', 'public.financial_reconciliation_match_compact(text)', 'EXECUTE')
    or not has_function_privilege('service_role', 'public.financial_reconciliation_match_normalize(text)', 'EXECUTE')
    or not has_function_privilege('service_role', 'public.financial_reconciliation_match_compact(text)', 'EXECUTE') then
    raise exception 'Automatic matching helper privileges inherit from PUBLIC.';
  end if;
end $$;

-- table constraints
do $$
begin
  begin
    update public.financial_reconciliation_automatic_schedule set time_zone = 'UTC' where id = true;
    raise exception 'Schedule time-zone constraint accepted UTC.';
  exception when check_violation then null;
  end;

  begin
    insert into public.financial_reconciliation_automatic_runs (
      trigger, scope, actor, scheduled_slot
    ) values (
      'scheduled', 'rule', 'smoke', '2026-08-14'
    );
    raise exception 'Scheduled run constraint accepted rule scope.';
  exception when check_violation then null;
  end;

  begin
    update public.financial_reconciliation_automatic_rule_configs
    set difference_allowed = -0.01
    where rule_key = 'financial_documents_cgd_bank_statement';
    raise exception 'Rule config constraint accepted a negative tolerance.';
  exception when check_violation then null;
  end;

  if not exists (
    select 1 from pg_constraint
    where conrelid = 'public.financial_reconciliation_automatic_rule_configs'::regclass
      and conname = 'financial_reconciliation_automatic_rule_configs_priority_key'
      and condeferrable
      and condeferred
  ) then
    raise exception 'Automatic rule priority uniqueness is not initially deferred.';
  end if;
end $$;

-- unknown-rule rejection
do $$
begin
  begin
    perform public.replace_financial_reconciliation_automation_settings(
      '{"enabled":false,"time_of_day":"02:00","time_zone":"Europe/Lisbon"}'::jsonb,
      '[{"rule_key":"unknown","rule_version":1,"enabled":false,"allow_manual_execution":false,"include_in_scheduled_batch":false,"difference_allowed":"0.00","max_difference_days":7,"priority":1}]'::jsonb,
      'smoke:administrator'
    );
    raise exception 'Unknown automatic rule was accepted.';
  exception when raise_exception then
    if sqlerrm <> 'Automatic rule is invalid.' then raise; end if;
  end;
end $$;

-- duplicate-priority rejection
do $$
begin
  begin
    perform public.replace_financial_reconciliation_automation_settings(
      '{"enabled":false,"time_of_day":"02:00","time_zone":"Europe/Lisbon"}'::jsonb,
      '[{"rule_key":"financial_documents_cgd_bank_statement","rule_version":1,"enabled":false,"allow_manual_execution":false,"include_in_scheduled_batch":false,"difference_allowed":"0.00","max_difference_days":7,"priority":1},{"rule_key":"financial_documents_cgd_bank_statement","rule_version":1,"enabled":false,"allow_manual_execution":false,"include_in_scheduled_batch":false,"difference_allowed":"0.00","max_difference_days":7,"priority":1}]'::jsonb,
      'smoke:administrator'
    );
    raise exception 'Duplicate automatic rule priority was accepted.';
  exception when raise_exception then
    if sqlerrm <> 'Duplicate automatic rule priority.' then raise; end if;
  end;
end $$;

-- managed rule version
do $$
declare
  v_before jsonb;
  v_after jsonb;
begin
  insert into public.financial_reconciliation_automatic_rule_definitions (
    rule_key, version, display_name, base_source_type, destination_source_types,
    logic_description, definition
  )
  select rule_key, 2, display_name, base_source_type, destination_source_types,
         logic_description, definition
  from public.financial_reconciliation_automatic_rule_definitions
  where rule_key = 'financial_documents_cgd_bank_statement' and version = 1;

  select public.get_financial_reconciliation_automation_settings() into v_before;
  begin
    perform public.replace_financial_reconciliation_automation_settings(
      '{"enabled":false,"time_of_day":"23:45","time_zone":"Europe/Lisbon"}'::jsonb,
      '[{"rule_key":"financial_documents_cgd_bank_statement","rule_version":2,"enabled":false,"allow_manual_execution":false,"include_in_scheduled_batch":false,"difference_allowed":"3.21","max_difference_days":9,"priority":1}]'::jsonb,
      'smoke:managed-version'
    );
    raise exception 'Settings PUT accepted a mismatched managed rule version.';
  exception when raise_exception then
    if sqlerrm <> 'Submitted automatic rule version does not match managed configuration.' then raise; end if;
  end;
  select public.get_financial_reconciliation_automation_settings() into v_after;
  if v_after <> v_before then
    raise exception 'Mismatched managed rule version partially changed settings.';
  end if;
end $$;

-- source-rule lock recheck
do $$
begin
  delete from public.financial_reconciliation_source_rules
  where base_source_type = 'financial_documents'
    and matching_source_type = 'import_cgd_extrato_ordem';

  begin
    perform public.replace_financial_reconciliation_automation_settings(
      '{"enabled":true,"time_of_day":"04:30","time_zone":"Europe/Lisbon"}'::jsonb,
      '[{"rule_key":"financial_documents_cgd_bank_statement","rule_version":1,"enabled":true,"allow_manual_execution":true,"include_in_scheduled_batch":true,"difference_allowed":"0.00","max_difference_days":7,"priority":1}]'::jsonb,
      'smoke:missing-source-rule'
    );
    raise exception 'Enabled automatic rule without a directional source rule was accepted.';
  exception when raise_exception then
    if sqlerrm <> 'No directional source rule exists for an enabled automatic rule.' then raise; end if;
  end;

  insert into public.financial_reconciliation_source_rules (
    base_source_type, matching_source_type, operator
  ) values (
    'financial_documents', 'import_cgd_extrato_ordem', '+'
  );
end $$;

-- atomic rollback
do $$
declare
  v_before jsonb;
  v_after jsonb;
begin
  select public.get_financial_reconciliation_automation_settings() into v_before;
  begin
    perform public.replace_financial_reconciliation_automation_settings(
      '{"enabled":false,"time_of_day":"23:59","time_zone":"Europe/Lisbon"}'::jsonb,
      '[{"rule_key":"unknown","rule_version":1,"enabled":false,"allow_manual_execution":false,"include_in_scheduled_batch":false,"difference_allowed":"0.00","max_difference_days":7,"priority":1}]'::jsonb,
      'smoke:failed-update'
    );
  exception when raise_exception then
    if sqlerrm <> 'Automatic rule is invalid.' then raise; end if;
  end;
  select public.get_financial_reconciliation_automation_settings() into v_after;
  if v_after <> v_before then
    raise exception 'Rejected settings payload partially changed persisted settings.';
  end if;
end $$;

-- priority swap
do $$
begin
  insert into public.financial_reconciliation_automatic_rule_definitions (
    rule_key, version, display_name, base_source_type, destination_source_types,
    logic_description, definition
  ) values (
    'smoke_second_rule', 1, 'Smoke second rule', 'financial_documents',
    '["import_cgd_extrato_ordem"]'::jsonb, 'Smoke fixture.', '{}'::jsonb
  );
  insert into public.financial_reconciliation_automatic_rule_configs (
    rule_key, rule_version, enabled, allow_manual_execution,
    include_in_scheduled_batch, difference_allowed, max_difference_days, priority
  ) values (
    'smoke_second_rule', 1, false, false, false, 0.00, 7, 2
  );

  perform public.replace_financial_reconciliation_automation_settings(
    '{"enabled":true,"time_of_day":"04:30","time_zone":"Europe/Lisbon"}'::jsonb,
    '[{"rule_key":"financial_documents_cgd_bank_statement","rule_version":1,"enabled":false,"allow_manual_execution":false,"include_in_scheduled_batch":false,"difference_allowed":"3.21","max_difference_days":9,"priority":2},{"rule_key":"smoke_second_rule","rule_version":1,"enabled":false,"allow_manual_execution":false,"include_in_scheduled_batch":false,"difference_allowed":"0.00","max_difference_days":7,"priority":1}]'::jsonb,
    'smoke:priority-swap'
  );

  if not exists (
      select 1 from public.financial_reconciliation_automatic_rule_configs
      where rule_key = 'financial_documents_cgd_bank_statement' and priority = 2
    ) or not exists (
      select 1 from public.financial_reconciliation_automatic_rule_configs
      where rule_key = 'smoke_second_rule' and priority = 1
    ) then
    raise exception 'Automatic rule priorities did not swap atomically.';
  end if;
end $$;

-- provenance checks
do $$
declare
  v_run_id uuid;
  v_proposal_id uuid;
  v_reconciliation_id uuid;
begin
  insert into public.financial_reconciliations (
    status, base_source_type, matching_source_types, created_by
  ) values (
    'started', 'financial_documents', '["import_cgd_extrato_ordem"]'::jsonb, 'smoke:user'
  ) returning id into v_reconciliation_id;

  if exists (
    select 1 from public.financial_reconciliations
    where id = v_reconciliation_id
      and (origin <> 'user' or automatic_trigger is not null or automatic_run_id is not null or automatic_proposal_id is not null)
  ) then
    raise exception 'Manual reconciliation defaults do not preserve null automatic provenance.';
  end if;

  begin
    insert into public.financial_reconciliations (
      status, base_source_type, matching_source_types, created_by, origin
    ) values (
      'started', 'financial_documents', '["import_cgd_extrato_ordem"]'::jsonb, 'smoke:automatic', 'automatic'
    );
    raise exception 'Incomplete automatic provenance was accepted.';
  exception when check_violation then null;
  end;

  insert into public.financial_reconciliation_automatic_runs (
    trigger, scope, actor, client_request_id
  ) values (
    'manual', 'rule', 'smoke:automatic', gen_random_uuid()
  ) returning id into v_run_id;

  insert into public.financial_reconciliation_automatic_proposals (
    run_id, rule_key, rule_version, base_source_type, base_source_id,
    base_source_date, allowed_difference, signature
  ) values (
    v_run_id, 'financial_documents_cgd_bank_statement', 1,
    'financial_documents', gen_random_uuid(), date '2026-08-14', 0.00, 'smoke-signature'
  ) returning id into v_proposal_id;

  insert into public.financial_reconciliations (
    status, base_source_type, matching_source_types, created_by,
    origin, automatic_trigger, automatic_rule_key, automatic_rule_version,
    automatic_run_id, automatic_proposal_id
  ) values (
    'started', 'financial_documents', '["import_cgd_extrato_ordem"]'::jsonb, 'smoke:automatic',
    'automatic', 'manual', 'financial_documents_cgd_bank_statement', 1,
    v_run_id, v_proposal_id
  ) returning id into v_reconciliation_id;

  if not exists (
    select 1 from public.financial_reconciliations
    where id = v_reconciliation_id
      and origin = 'automatic'
      and automatic_trigger = 'manual'
      and automatic_run_id = v_run_id
      and automatic_proposal_id = v_proposal_id
  ) then
    raise exception 'Complete automatic provenance was not persisted.';
  end if;
end $$;

-- document-number containment
-- description score immediately below and at 0.60
-- supplier word score immediately below and at 0.70
-- blank identity fields
do $$
declare
  v_candidate_count integer;
  v_description_score real;
  v_supplier_score real;
  v_supplier_at text;
begin
  if public.financial_reconciliation_match_normalize(' Fatura Nº 12, Árvore! ') <> 'fatura 12 arvore'
    or public.financial_reconciliation_match_compact('FT-2026/001234') <> 'ft2026001234' then
    raise exception 'Deterministic normalization did not preserve significant text and numeric tokens.';
  end if;
  if public.financial_reconciliation_match_normalize('a de 12') <> '12' then
    raise exception 'Short alphabetic identity tokens were retained.';
  end if;
  if public.financial_reconciliation_match_normalize('  ') <> '' then
    raise exception 'Blank identity fields produced a normalized value.';
  end if;
  insert into public.financial_documents (id, document_date, doc_number, description, supplier_name, amount, fat)
  values ('00000000-0000-0000-0000-000000000d01', date '2026-03-20', 'FT-2026/001234', 'Invoice service', 'Supplier', 100.00, 'S');
  insert into public.import_cgd_extrato_ordem (id, import_batch, row_key, data, descritivo, montante) values
    ('00000000-0000-0000-0000-000000000b01', 'smoke-analysis', 'smoke-analysis-seven', date '2026-03-27', 'Settlement FT2026001234', -100.00),
    ('00000000-0000-0000-0000-000000000b02', 'smoke-analysis', 'smoke-analysis-eight', date '2026-03-28', 'Settlement FT2026001234', -100.00);
  select candidate_count into strict v_candidate_count
  from public.financial_reconciliation_automatic_rule_candidates(
    'financial_documents_cgd_bank_statement', 1, 0.00, 7
  ) where base_source_id = '00000000-0000-0000-0000-000000000d01';
  if v_candidate_count <> 1 then raise exception 'Inclusive seven-day date boundary did not exclude the eighth day.'; end if;
  insert into public.financial_documents (id, document_date, doc_number, description, supplier_name, amount, fat) values
    ('00000000-0000-0000-0000-000000000d03', date '2026-05-01', '', 'abcdefg', '', 10.00, 'S'),
    ('00000000-0000-0000-0000-000000000d04', date '2026-06-01', '', 'abcdefg', '', 10.00, 'S'),
    ('00000000-0000-0000-0000-000000000d05', date '2026-07-01', '', '', 'abcdefg', 10.00, 'S'),
    ('00000000-0000-0000-0000-000000000d06', date '2026-08-01', '', '', 'abcdefg', 10.00, 'S'),
    ('00000000-0000-0000-0000-000000000d07', date '2026-09-01', '', '', '', 10.00, 'S');
  select candidate into v_supplier_at from (
    select left('abcdefg', prefix_length) || repeat('z', suffix_length) as candidate,
           word_similarity('abcdefg', left('abcdefg', prefix_length) || repeat('z', suffix_length)) as score
    from generate_series(1, 7) prefix_length cross join generate_series(1, 12) suffix_length
  ) scores where abs(score - 0.70) < 0.000001 order by candidate limit 1;
  if v_supplier_at is null then raise exception 'No repeatable supplier word-similarity fixture reached 0.70.'; end if;
  insert into public.import_cgd_extrato_ordem (id, import_batch, row_key, data, descritivo, montante) values
    ('00000000-0000-0000-0000-000000000b03', 'smoke-analysis', 'smoke-description-at', date '2026-05-01', 'abcdefx', -10.00),
    ('00000000-0000-0000-0000-000000000b04', 'smoke-analysis', 'smoke-description-below', date '2026-06-01', 'abcdeyx', -10.00),
    ('00000000-0000-0000-0000-000000000b05', 'smoke-analysis', 'smoke-supplier-at', date '2026-07-01', v_supplier_at, -10.00),
    ('00000000-0000-0000-0000-000000000b06', 'smoke-analysis', 'smoke-supplier-below', date '2026-08-01', 'zzzzzzzzzz', -10.00),
    ('00000000-0000-0000-0000-000000000b07', 'smoke-analysis', 'smoke-blank-identity', date '2026-09-01', '', -10.00);
  select ((candidates->0->'evidence'->'description'->>'score')::real) into strict v_description_score
  from public.financial_reconciliation_automatic_rule_candidates('financial_documents_cgd_bank_statement', 1, 0.00, 7)
  where base_source_id = '00000000-0000-0000-0000-000000000d03';
  if abs(v_description_score - 0.60) >= 0.000001 then raise exception 'Description threshold fixture did not measure exactly 0.60.'; end if;
  select candidate_count into strict v_candidate_count
  from public.financial_reconciliation_automatic_rule_candidates('financial_documents_cgd_bank_statement', 1, 0.00, 7)
  where base_source_id = '00000000-0000-0000-0000-000000000d04';
  -- Description threshold below fixture was accepted
  if v_candidate_count <> 0 then raise exception 'Description threshold below fixture was accepted.'; end if;
  select ((candidates->0->'evidence'->'supplier'->>'score')::real) into strict v_supplier_score
  from public.financial_reconciliation_automatic_rule_candidates('financial_documents_cgd_bank_statement', 1, 0.00, 7)
  where base_source_id = '00000000-0000-0000-0000-000000000d05';
  if abs(v_supplier_score - 0.70) >= 0.000001 then raise exception 'Supplier threshold fixture did not measure exactly 0.70.'; end if;
  select candidate_count into strict v_candidate_count
  from public.financial_reconciliation_automatic_rule_candidates('financial_documents_cgd_bank_statement', 1, 0.00, 7)
  where base_source_id = '00000000-0000-0000-0000-000000000d06';
  -- Supplier threshold below fixture was accepted
  if v_candidate_count <> 0 then raise exception 'Supplier threshold below fixture was accepted.'; end if;
  select candidate_count into strict v_candidate_count
  from public.financial_reconciliation_automatic_rule_candidates('financial_documents_cgd_bank_statement', 1, 0.00, 7)
  where base_source_id = '00000000-0000-0000-0000-000000000d07';
  -- Blank identity fixture produced a candidate
  if v_candidate_count <> 0 then raise exception 'Blank identity fixture produced a candidate.'; end if;
end $$;

-- cross-base overlap includes ambiguous candidate groups and persisted counters
do $$
declare v_overlap_run uuid; v_overlap_result jsonb; v_expected_counts jsonb;
begin
  insert into public.financial_documents (id, document_date, doc_number, description, supplier_name, amount, fat)
  values ('00000000-0000-0000-0000-000000000d08', date '2026-03-21', 'FT-2026/001234', 'Invoice service', 'Supplier', 100.00, 'S');
  insert into public.financial_reconciliation_automatic_runs (
    trigger, scope, actor, client_request_id, definition_config_snapshot
  ) values (
    'manual', 'rule', 'smoke:overlap', '00000000-0000-0000-0000-000000000d09',
    '[{"ruleKey":"financial_documents_cgd_bank_statement","ruleVersion":1,"priority":1,"differenceAllowed":0.00,"maxDifferenceDays":7,"operator":"+"}]'::jsonb
  ) returning id into v_overlap_run;
  select public.populate_financial_reconciliation_automatic_run(v_overlap_run) into v_overlap_result;
  if (select count(*) from public.financial_reconciliation_automatic_proposals
      where run_id = v_overlap_run and base_source_id in ('00000000-0000-0000-0000-000000000d01', '00000000-0000-0000-0000-000000000d08')
        and status = 'ambiguous' and reason = 'cross_base_overlap') <> 2 then
    -- Cross-base overlap did not mark every affected proposal ambiguous
    raise exception 'Cross-base overlap did not mark every affected proposal ambiguous.';
  end if;
  select jsonb_build_object(
    'bases', count(distinct base_source_id),
    'proposed', count(*) filter (where status = 'proposed'),
    'ambiguous', count(*) filter (where status = 'ambiguous')
  ) into v_expected_counts from public.financial_reconciliation_automatic_proposals where run_id = v_overlap_run;
  if v_overlap_result->'counts' <> v_expected_counts then
    raise exception 'Persisted proposal counters were not recomputed after overlap ambiguity.';
  end if;
end $$;

-- dates exactly 7 and 8 days apart
-- differences exactly at and above tolerance
-- one-to-one and one-to-many sums
-- independent operators
do $$
declare v_count integer;
begin
  select count(*) into v_count from public.financial_reconciliation_automatic_build_combinations(
    '{"amount":"100.00"}'::jsonb,
    '[{"sourceType":"import_cgd_extrato_ordem","sourceId":"00000000-0000-0000-0000-000000000101","sourceDate":"2026-08-07","amount":"-99.00"}]'::jsonb,
    '{"import_cgd_extrato_ordem":"+"}'::jsonb, 1.00, 4
  );
  if v_count <> 1 then raise exception 'Inclusive integer-cent tolerance rejected exact boundary.'; end if;
  select count(*) into v_count from public.financial_reconciliation_automatic_build_combinations(
    '{"amount":"100.00"}'::jsonb,
    '[{"sourceType":"import_cgd_extrato_ordem","sourceId":"00000000-0000-0000-0000-000000000102","sourceDate":"2026-08-08","amount":"-98.99"}]'::jsonb,
    '{"import_cgd_extrato_ordem":"+"}'::jsonb, 1.00, 4
  );
  if v_count <> 0 then raise exception 'Integer-cent tolerance accepted amount above boundary.'; end if;
  select count(*) into v_count from public.financial_reconciliation_automatic_build_combinations(
    '{"amount":"100.00"}'::jsonb,
    '[{"sourceType":"import_cgd_extrato_ordem","sourceId":"00000000-0000-0000-0000-000000000103","sourceDate":"2026-08-14","amount":"-60.00"},{"sourceType":"import_cgd_cartao_credito","sourceId":"00000000-0000-0000-0000-000000000104","sourceDate":"2026-08-14","amount":"40.00"}]'::jsonb,
    '{"import_cgd_extrato_ordem":"+","import_cgd_cartao_credito":"-"}'::jsonb, 0.00, 4
  );
  if v_count <> 1 then raise exception 'Heterogeneous candidates did not retain independent operators.'; end if;
  select count(*) into v_count from public.financial_reconciliation_automatic_build_combinations(
    '{"amount":"100.00"}'::jsonb,
    '[{"sourceType":"import_cgd_extrato_ordem","sourceId":"00000000-0000-0000-0000-000000000105","sourceDate":"2026-08-14","amount":"-40.00"},{"sourceType":"import_cgd_extrato_ordem","sourceId":"00000000-0000-0000-0000-000000000106","sourceDate":"2026-08-14","amount":"-60.00"}]'::jsonb,
    '{"import_cgd_extrato_ordem":"+"}'::jsonb, 0.00, 4
  );
  if v_count <> 1 then raise exception 'One-to-many integer-cent sum did not produce one complete group.'; end if;
  select count(*) into v_count from public.financial_reconciliation_automatic_build_combinations(
    '{"amount":"100.00"}'::jsonb,
    '[{"sourceType":"import_cgd_extrato_ordem","sourceId":"00000000-0000-0000-0000-000000000107","sourceDate":"2026-08-14","amount":"-100.00"},{"sourceType":"import_cgd_extrato_ordem","sourceId":"00000000-0000-0000-0000-000000000108","sourceDate":"2026-08-14","amount":"-100.00"}]'::jsonb,
    '{"import_cgd_extrato_ordem":"+"}'::jsonb, 0.00, 4
  );
  if v_count <> 2 then raise exception 'Two valid combinations did not remain ambiguous candidates.'; end if;
end $$;

-- two valid combinations for one base
-- cross-base overlap
-- candidate_limit
-- Lisbon DST slot claim
do $$
declare v_first jsonb; v_second jsonb; v_candidate_count integer; v_limit_run uuid; v_limit_result jsonb;
begin
  insert into public.financial_documents (id, document_date, doc_number, description, supplier_name, amount, fat)
  values ('00000000-0000-0000-0000-000000000d02', date '2026-04-01', 'FT-2026/009999', 'Limit fixture', 'Supplier', 100.00, 'S');
  insert into public.import_cgd_extrato_ordem (id, import_batch, row_key, data, descritivo, montante)
  select ('00000000-0000-0000-0000-' || lpad(n::text, 12, '0'))::uuid, 'smoke-analysis', 'smoke-limit-' || n,
         date '2026-04-01', 'Settlement FT2026009999', -100.00
  from generate_series(1, 13) n;
  select candidate_count into strict v_candidate_count
  from public.financial_reconciliation_automatic_rule_candidates(
    'financial_documents_cgd_bank_statement', 1, 0.00, 7
  ) where base_source_id = '00000000-0000-0000-0000-000000000d02';
  if v_candidate_count <> 13 then raise exception 'Thirteen identity-qualified candidates did not trigger the candidate limit fixture.'; end if;
  insert into public.financial_reconciliation_automatic_runs (
    trigger, scope, actor, client_request_id, definition_config_snapshot
  ) values (
    'manual', 'rule', 'smoke:candidate-limit', '00000000-0000-0000-0000-000000000d03',
    '[{"ruleKey":"financial_documents_cgd_bank_statement","ruleVersion":1,"priority":1,"differenceAllowed":0.00,"maxDifferenceDays":7,"operator":"+"}]'::jsonb
  ) returning id into v_limit_run;
  select public.populate_financial_reconciliation_automatic_run(v_limit_run) into v_limit_result;
  if (select count(*) from public.financial_reconciliation_automatic_proposals
      where run_id = v_limit_run and base_source_id = '00000000-0000-0000-0000-000000000d02'
        and status = 'ambiguous' and reason = 'candidate_limit' and candidate_groups <> '[]'::jsonb and items = '[]'::jsonb) <> 1 then
    -- Candidate-limit run did not persist exactly one ambiguous proposal
    raise exception 'Candidate-limit run did not persist exactly one ambiguous proposal.';
  end if;
  update public.financial_reconciliation_automatic_schedule set enabled = true, time_of_day = time '00:00' where id = true;
  update public.financial_reconciliation_automatic_rule_configs
  set enabled = true, include_in_scheduled_batch = true, allow_manual_execution = true
  where rule_key = 'financial_documents_cgd_bank_statement';
  select public.claim_financial_reconciliation_automatic_schedule('2026-03-29 00:30:00+00', 'smoke:schedule') into v_first;
  select public.claim_financial_reconciliation_automatic_schedule('2026-03-29 01:30:00+00', 'smoke:schedule') into v_second;
  if not (v_first->>'claimed')::boolean or not (v_second->>'claimed')::boolean
    or v_first#>>'{run,runId}' <> v_second#>>'{run,runId}' then
    raise exception 'Lisbon DST slot claim did not produce exactly one scheduled run.';
  end if;
end $$;

rollback;
