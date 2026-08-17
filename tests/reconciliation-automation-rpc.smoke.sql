\set ON_ERROR_STOP on

begin;

\ir ../supabase-migrations/2026-08-14-financial-reconciliation-automation-schema.sql
\ir ../supabase-migrations/2026-08-14-financial-reconciliation-automation-analysis.sql
\ir ../supabase-migrations/2026-08-14-financial-reconciliation-automation-execution.sql
\ir ../supabase-migrations/2026-08-15-financial-reconciliation-automation-analysis-performance.sql
\ir ../supabase-migrations/2026-08-15-financial-reconciliation-automation-candidate-index-lookup.sql

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
\ir ../supabase-migrations/2026-08-14-financial-reconciliation-automation-execution.sql
\ir ../supabase-migrations/2026-08-15-financial-reconciliation-automation-analysis-performance.sql
\ir ../supabase-migrations/2026-08-15-financial-reconciliation-automation-candidate-index-lookup.sql

-- optimized analysis definition and privileges
do $$
declare
  v_candidate_definition text;
begin
  select pg_get_functiondef('public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer)'::regprocedure)
  into strict v_candidate_definition;

  if v_candidate_definition !~* 'bases\s+as\s+materialized'
    or v_candidate_definition !~* 'qualified\s+as\s+materialized'
    or v_candidate_definition !~* 'scored\s+as\s+materialized'
    or v_candidate_definition !~* 'left join lateral\s+\([\s\S]+from public\.import_cgd_extrato_ordem bank' then
    raise exception 'Optimized automatic candidate stages were not installed.';
  end if;

  if v_candidate_definition ~* 'bank_rows\s+as\s+materialized' then
    raise exception 'Bank candidate rows must remain index-driven.';
  end if;

  if has_function_privilege(
      'anon',
      'public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer)',
      'EXECUTE'
    )
    or has_function_privilege(
      'authenticated',
      'public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer)',
      'EXECUTE'
    )
    or not has_function_privilege(
      'service_role',
      'public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer)',
      'EXECUTE'
    ) then
    raise exception 'Optimized automatic candidate privileges are invalid.';
  end if;
end $$;

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
    or has_function_privilege('anon', 'public.get_financial_reconciliation_automatic_manual_rules()', 'EXECUTE')
    or has_function_privilege('authenticated', 'public.get_financial_reconciliation_automatic_manual_rules()', 'EXECUTE')
    or not has_function_privilege('service_role', 'public.get_financial_reconciliation_automation_settings()', 'EXECUTE')
    or not has_function_privilege('service_role', 'public.get_financial_reconciliation_automatic_manual_rules()', 'EXECUTE')
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

-- manual rule catalog is filtered and RPC-only
do $$
declare
  v_catalog jsonb;
  v_rule jsonb;
begin
  update public.financial_reconciliation_automatic_rule_configs
  set enabled = true, allow_manual_execution = true
  where rule_key = 'financial_documents_cgd_bank_statement';

  select public.get_financial_reconciliation_automatic_manual_rules() into v_catalog;
  if jsonb_typeof(v_catalog) <> 'object'
    or jsonb_array_length(v_catalog->'rules') <> 1
    or v_catalog ? 'schedule'
    or v_catalog ? 'lastScheduledRun' then
    raise exception 'Manual automatic rule catalog exposed the wrong public envelope.';
  end if;
  v_rule := v_catalog->'rules'->0;
  if v_rule->>'ruleKey' <> 'financial_documents_cgd_bank_statement'
    or v_rule->>'enabled' <> 'true'
    or v_rule->>'allowManualExecution' <> 'true'
    or not (v_rule ?& array[
      'ruleKey','ruleVersion','displayName','baseSourceType','destinationSourceTypes',
      'logicDescription','definition','enabled','allowManualExecution',
      'differenceAllowed','maxDifferenceDays','priority'
    ])
    or v_rule ?| array['includeInScheduledBatch','updatedBy','updatedAt','errorSummary','diagnostic'] then
    raise exception 'Manual automatic rule catalog exposed unsafe or incomplete fields.';
  end if;

  update public.financial_reconciliation_automatic_rule_configs
  set enabled = false, allow_manual_execution = false
  where rule_key = 'financial_documents_cgd_bank_statement';
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

  delete from public.financial_reconciliation_automatic_rule_configs
  where rule_key = 'smoke_second_rule';
  delete from public.financial_reconciliation_automatic_rule_definitions
  where rule_key = 'smoke_second_rule';
  update public.financial_reconciliation_automatic_rule_configs
  set priority = 1
  where rule_key = 'financial_documents_cgd_bank_statement';
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
  v_description_below text;
  v_description_below_score real;
  v_supplier_score real;
  v_supplier_at text;
  v_supplier_below text;
  v_supplier_below_score real;
begin
  if public.financial_reconciliation_match_normalize(' Fatura Nº 12, Árvore! ') <> 'fatura 12 arvore'
    or public.financial_reconciliation_match_compact('FT-2026/001234') <> 'ft2026001234'
    or public.financial_reconciliation_match_compact('AB-12') <> 'ab12' then
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
  insert into public.financial_documents (id, document_date, doc_number, description, supplier_name, amount, fat)
  values ('00000000-0000-0000-0000-000000000d0a', date '2026-03-20', 'FT-1234', '', '', 25.00, 'S');
  insert into public.import_cgd_extrato_ordem (id, import_batch, row_key, data, descritivo, montante)
  values ('00000000-0000-0000-0000-000000000b0a', 'smoke-analysis', 'smoke-short-prefix', date '2026-03-20', 'Unrelated settlement 1234', -25.00);
  select candidate_count into strict v_candidate_count
  from public.financial_reconciliation_automatic_rule_candidates('financial_documents_cgd_bank_statement', 1, 0.00, 7)
  where base_source_id = '00000000-0000-0000-0000-000000000d0a';
  if v_candidate_count <> 0 then
    raise exception 'Document-number matching discarded a short alphabetic prefix.';
  end if;
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
  select candidate, score into v_description_below, v_description_below_score from (
    select candidate, similarity('abcdefg', candidate) as score
    from (values ('abc'), ('abcd'), ('abcde'), ('abcdey'), ('abcdeyx'), ('abcdef'), ('abcdefx')) corpus(candidate)
  ) scores where score < 0.60 order by 0.60 - score, candidate limit 1;
  -- Description below fixture was not boundary-adjacent
  if v_description_below is null or v_description_below_score >= 0.60 or 0.60 - v_description_below_score > 0.05 then
    raise exception 'Description below fixture was not boundary-adjacent.';
  end if;
  select candidate, score into v_supplier_below, v_supplier_below_score from (
    select candidate, word_similarity('abcdefg', candidate) as score
    from (
      select left('abcdefg', prefix_length) || repeat('z', suffix_length) as candidate
      from generate_series(1, 7) prefix_length cross join generate_series(1, 12) suffix_length
    ) corpus
  ) scores where score < 0.70 order by 0.70 - score, candidate limit 1;
  -- Supplier below fixture was not boundary-adjacent
  if v_supplier_below is null or v_supplier_below_score >= 0.70 or 0.70 - v_supplier_below_score > 0.05 then
    raise exception 'Supplier below fixture was not boundary-adjacent.';
  end if;
  insert into public.import_cgd_extrato_ordem (id, import_batch, row_key, data, descritivo, montante) values
    ('00000000-0000-0000-0000-000000000b03', 'smoke-analysis', 'smoke-description-at', date '2026-05-01', 'abcdefx', -10.00),
    ('00000000-0000-0000-0000-000000000b04', 'smoke-analysis', 'smoke-description-below', date '2026-06-01', v_description_below, -10.00),
    ('00000000-0000-0000-0000-000000000b05', 'smoke-analysis', 'smoke-supplier-at', date '2026-07-01', v_supplier_at, -10.00),
    ('00000000-0000-0000-0000-000000000b06', 'smoke-analysis', 'smoke-supplier-below', date '2026-08-01', v_supplier_below, -10.00),
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
  if not exists (
    select 1 from public.financial_reconciliation_automatic_proposals
    where run_id = v_overlap_run
      and base_source_id = '00000000-0000-0000-0000-000000000d04'
      and status = 'skipped'
      and reason = 'no_qualifying_combination'
  ) then
    raise exception 'A base without a qualifying combination was not persisted as skipped.';
  end if;
  select jsonb_build_object(
    'bases', count(distinct base_source_id),
    'proposed', count(*) filter (where status = 'proposed'),
    'ambiguous', count(*) filter (where status = 'ambiguous'),
    'skipped', count(*) filter (where status = 'skipped')
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
-- cross-midnight scheduled resume
do $$
declare v_first jsonb; v_second jsonb; v_next_day jsonb; v_candidate_count integer; v_limit_run uuid; v_limit_result jsonb;
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
  select public.claim_financial_reconciliation_automatic_schedule('2026-03-30 00:30:00+00', 'smoke:schedule') into v_next_day;
  if not (v_first->>'claimed')::boolean or not (v_second->>'claimed')::boolean
    or v_first#>>'{run,runId}' <> v_second#>>'{run,runId}' then
    raise exception 'Lisbon DST slot claim did not produce exactly one scheduled run.';
  end if;
  if not (v_next_day->>'claimed')::boolean
    or not (v_next_day->>'resumed')::boolean
    or v_first#>>'{run,runId}' <> v_next_day#>>'{run,runId}'
    or (select count(*) from public.financial_reconciliation_automatic_runs where trigger = 'scheduled') <> 1 then
    raise exception 'Unfinished scheduled run was not resumed across Lisbon dates.';
  end if;
end $$;

-- automatic execution RPC privileges
do $$
begin
  if has_function_privilege('anon', 'public.execute_financial_reconciliation_automatic_proposal(uuid,text)', 'EXECUTE')
    or has_function_privilege('authenticated', 'public.execute_financial_reconciliation_automatic_proposal(uuid,text)', 'EXECUTE')
    or not has_function_privilege('service_role', 'public.execute_financial_reconciliation_automatic_proposal(uuid,text)', 'EXECUTE')
    or has_function_privilege('anon', 'public.finish_financial_reconciliation_automatic_run(uuid)', 'EXECUTE')
    or has_function_privilege('authenticated', 'public.finish_financial_reconciliation_automatic_run(uuid)', 'EXECUTE')
    or not has_function_privilege('service_role', 'public.finish_financial_reconciliation_automatic_run(uuid)', 'EXECUTE') then
    raise exception 'Automatic execution RPC privileges are invalid.';
  end if;
end $$;

update public.financial_reconciliation_automatic_rule_configs
set enabled = true,
    allow_manual_execution = true,
    include_in_scheduled_batch = true,
    difference_allowed = 1.00,
    max_difference_days = 7
where rule_key = 'financial_documents_cgd_bank_statement';

update public.financial_reconciliation_source_rules
set operator = '+'
where base_source_type = 'financial_documents'
  and matching_source_type = 'import_cgd_extrato_ordem';

create or replace function pg_temp.make_automatic_proposal(
  p_base_source_id uuid,
  p_actor text,
  p_client_request_id uuid
)
returns uuid
language plpgsql
as $$
declare
  v_run jsonb;
  v_run_id uuid;
  v_proposal_id uuid;
begin
  v_run := public.create_financial_reconciliation_automatic_analysis(
    array['financial_documents_cgd_bank_statement'],
    'manual_rule',
    p_actor,
    p_client_request_id
  );
  v_run_id := (v_run->>'runId')::uuid;
  select proposal.id into strict v_proposal_id
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.run_id = v_run_id
    and proposal.base_source_id = p_base_source_id
    and proposal.status = 'proposed';
  return v_proposal_id;
end $$;

-- proposal base snapshot is complete and immutable
-- execution rejects a changed base snapshot
do $$
declare
  v_document_id uuid := '10000000-0000-0000-0000-000000000100';
  v_bank_id uuid := '20000000-0000-0000-0000-000000000100';
  v_proposal_id uuid;
  v_run_id uuid;
  v_snapshot jsonb;
  v_public_run jsonb;
  v_public_proposal jsonb;
  v_result jsonb;
begin
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, amount, fat
  ) values (
    v_document_id, date '2026-09-30', 'AUTO-SNAPSHOT-0100',
    'Original invoice description', 'Snapshot Supplier', 90.00, 'S'
  );
  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  ) values (
    v_bank_id, 'smoke-snapshot', 'smoke-auto-snapshot-0100', date '2026-09-30',
    'Payment AUTOSNAPSHOT0100', -90.00
  );

  v_proposal_id := pg_temp.make_automatic_proposal(
    v_document_id, 'smoke:base-snapshot', '30000000-0000-0000-0000-000000000100'
  );
  select run_id, base_snapshot into strict v_run_id, v_snapshot
  from public.financial_reconciliation_automatic_proposals
  where id = v_proposal_id;
  v_public_run := public.get_financial_reconciliation_automatic_run(v_run_id);
  if v_public_run#>>'{definitions,0,displayName}' <> 'Financial Documents to CGD Bank Statement' then
    raise exception 'Automatic run snapshot omitted the managed friendly rule name.';
  end if;
  if v_snapshot is distinct from jsonb_build_object(
    'sourceType', 'financial_documents',
    'sourceId', v_document_id,
    'sourceDate', date '2026-09-30',
    'amount', 90.00,
    'docNumber', 'AUTO-SNAPSHOT-0100',
    'description', 'Original invoice description',
    'supplierName', 'Snapshot Supplier'
  ) then
    raise exception 'Automatic proposal did not persist the complete base snapshot.';
  end if;

  select proposal.value into strict v_public_proposal
  from jsonb_array_elements(v_public_run->'proposals') proposal(value)
  where proposal.value->>'id' = v_proposal_id::text;
  if v_public_proposal->'baseSnapshot' is distinct from v_snapshot then
    raise exception 'Automatic run detail omitted or changed the base snapshot.';
  end if;

  begin
    update public.financial_reconciliation_automatic_proposals
    set base_snapshot = jsonb_build_object('sourceType', 'tampered')
    where id = v_proposal_id;
    raise exception 'Automatic proposal base snapshot accepted mutation.';
  exception when raise_exception then
    if sqlerrm <> 'Automatic proposal base snapshot is immutable.' then raise; end if;
  end;

  update public.financial_documents
  set description = 'Changed after analysis'
  where id = v_document_id;
  v_result := public.execute_financial_reconciliation_automatic_proposal(
    v_proposal_id, 'smoke:base-snapshot'
  );
  if v_result->>'status' <> 'stale'
    or v_result->>'reason' <> 'source_snapshot_changed'
    or exists (
      select 1 from public.financial_reconciliation_automatic_proposals
      where id = v_proposal_id and reconciliation_id is not null
    ) then
    raise exception 'Execution accepted a source record that differed from its base snapshot.';
  end if;
end $$;

-- non-zero automatic completion and idempotency
-- all automatic items were not locked
-- automatic provenance was not persisted
-- generated automatic completion comment was not stable
-- repeated automatic execution duplicated items or audit rows
-- automatic lifecycle snapshots were not rechecked before completion
-- automatic reopen/delete provenance
-- automatic lifecycle action changed provenance
do $$
declare
  v_document_id uuid := '10000000-0000-0000-0000-000000000101';
  v_bank_id uuid := '20000000-0000-0000-0000-000000000101';
  v_proposal_id uuid;
  v_run_id uuid;
  v_reconciliation_id uuid;
  v_repeated_id uuid;
  v_result jsonb;
  v_workspace jsonb;
  v_history jsonb;
  v_expected_comment text;
  v_items_before integer;
  v_audit_before integer;
begin
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, amount, fat
  ) values (
    v_document_id, date '2026-10-01', 'AUTO-NONZERO-0101', '', '', 100.35, 'S'
  );
  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  ) values (
    v_bank_id, 'smoke-execution', 'smoke-auto-nonzero-0101', date '2026-10-01',
    'Payment AUTONONZERO0101', -100.00
  );

  v_proposal_id := pg_temp.make_automatic_proposal(
    v_document_id, 'smoke:automatic-nonzero', '30000000-0000-0000-0000-000000000101'
  );
  select run_id into strict v_run_id
  from public.financial_reconciliation_automatic_proposals
  where id = v_proposal_id;
  update public.financial_reconciliation_automatic_runs
  set trigger = 'scheduled', scope = 'batch', client_request_id = null,
      scheduled_slot = '2026-10-01'
  where id = v_run_id;

  v_result := public.execute_financial_reconciliation_automatic_proposal(
    v_proposal_id, 'smoke:automatic-nonzero'
  );
  v_reconciliation_id := (v_result->>'reconciliationId')::uuid;
  v_expected_comment := 'Automatically completed by rule Financial Documents to CGD Bank Statement v1; difference '
    || chr(8364) || '0.35 within allowed tolerance ' || chr(8364)
    || '1.00; trigger Scheduled; batch ' || v_run_id::text || '.';

  if (select count(*) from public.financial_reconciliation_items where reconciliation_id = v_reconciliation_id) <> 2
    or not exists (
      select 1 from public.financial_reconciliation_items
      where reconciliation_id = v_reconciliation_id
        and source_type = 'financial_documents' and source_id = v_document_id
    )
    or not exists (
      select 1 from public.financial_reconciliation_items
      where reconciliation_id = v_reconciliation_id
        and source_type = 'import_cgd_extrato_ordem' and source_id = v_bank_id
    ) then
    raise exception 'All automatic items were not locked.';
  end if;

  if not exists (
    select 1 from public.financial_reconciliations reconciliation
    where reconciliation.id = v_reconciliation_id
      and reconciliation.status = 'complete'
      and reconciliation.completion_type = 'forced'
      and reconciliation.difference_amount = 0.35
      and reconciliation.forced_completion_comment = v_expected_comment
      and reconciliation.origin = 'automatic'
      and reconciliation.automatic_trigger = 'scheduled'
      and reconciliation.automatic_rule_key = 'financial_documents_cgd_bank_statement'
      and reconciliation.automatic_rule_version = 1
      and reconciliation.automatic_run_id = v_run_id
      and reconciliation.automatic_proposal_id = v_proposal_id
  ) then
    raise exception 'Automatic provenance was not persisted.';
  end if;
  if (select forced_completion_comment from public.financial_reconciliations where id = v_reconciliation_id)
      is distinct from v_expected_comment then
    raise exception 'Generated automatic completion comment was not stable.';
  end if;

  if not exists (
    select 1 from public.financial_reconciliation_audit audit
    where audit.reconciliation_id = v_reconciliation_id
      and audit.action = 'automatic_complete'
      and audit.metadata @> jsonb_build_object(
        'trigger', 'scheduled',
        'runId', v_run_id,
        'proposalSignature', (select signature from public.financial_reconciliation_automatic_proposals where id = v_proposal_id),
        'tolerance', 1.00
      )
      and audit.metadata ?& array[
        'ruleSnapshot','configSnapshot','operatorSnapshot','identityEvidence',
        'proposalSignature','trigger','runId','tolerance'
      ]
  ) then
    raise exception 'Automatic completion audit metadata was incomplete.';
  end if;

  v_workspace := public.get_financial_reconciliation_workspace(
    v_reconciliation_id, 'financial_documents', '{}'::jsonb, 1, 50
  );
  select history.value into strict v_history
  from jsonb_array_elements(v_workspace->'history') history(value)
  where history.value->>'id' = v_reconciliation_id::text;
  if v_workspace#>>'{reconciliation,origin}' <> 'automatic'
    or v_workspace#>>'{reconciliation,automaticTrigger}' <> 'scheduled'
    or v_workspace#>>'{reconciliation,automaticRuleKey}' <> 'financial_documents_cgd_bank_statement'
    or v_workspace#>>'{reconciliation,automaticRuleVersion}' <> '1'
    or v_workspace#>>'{reconciliation,automaticRunId}' <> v_run_id::text
    or v_history->>'origin' <> 'automatic'
    or v_history->>'automaticTrigger' <> 'scheduled'
    or v_history->>'automaticRuleKey' <> 'financial_documents_cgd_bank_statement'
    or v_history->>'automaticRuleVersion' <> '1'
    or v_history->>'automaticRunId' <> v_run_id::text
    or jsonb_typeof(v_history->'sourceSummary') <> 'array' then
    raise exception 'Workspace/history automatic provenance or source summaries were not preserved.';
  end if;

  select count(*) into v_items_before
  from public.financial_reconciliation_items where reconciliation_id = v_reconciliation_id;
  select count(*) into v_audit_before
  from public.financial_reconciliation_audit where reconciliation_id = v_reconciliation_id;
  v_result := public.execute_financial_reconciliation_automatic_proposal(
    v_proposal_id, 'smoke:automatic-nonzero'
  );
  v_repeated_id := (v_result->>'reconciliationId')::uuid;
  if v_repeated_id <> v_reconciliation_id
    or (select count(*) from public.financial_reconciliation_items where reconciliation_id = v_reconciliation_id) <> v_items_before
    or (select count(*) from public.financial_reconciliation_audit where reconciliation_id = v_reconciliation_id) <> v_audit_before then
    raise exception 'Repeated automatic execution duplicated items or audit rows.';
  end if;

  perform public.financial_reconciliation_action(
    'reopen', 'smoke:automatic-lifecycle', v_reconciliation_id, null, null, null
  );
  if not exists (
    select 1 from public.financial_reconciliations
    where id = v_reconciliation_id and status = 'started'
      and origin = 'automatic' and automatic_run_id = v_run_id
      and automatic_proposal_id = v_proposal_id
  ) or (select count(*) from public.financial_reconciliation_items where reconciliation_id = v_reconciliation_id) <> 2 then
    raise exception 'Automatic lifecycle action changed provenance.';
  end if;
  perform public.financial_reconciliation_action(
    'force_complete', 'smoke:automatic-lifecycle', v_reconciliation_id, null, null,
    'Smoke lifecycle completion.'
  );
  perform public.financial_reconciliation_action(
    'delete', 'smoke:automatic-lifecycle', v_reconciliation_id, null, null, null
  );
  if not exists (
    select 1 from public.financial_reconciliations
    where id = v_reconciliation_id and deleted_at is not null
      and origin = 'automatic' and automatic_run_id = v_run_id
      and automatic_proposal_id = v_proposal_id
  ) or exists (
    select 1 from public.financial_reconciliation_items where reconciliation_id = v_reconciliation_id
  ) then
    raise exception 'Automatic lifecycle action changed provenance.';
  end if;
end $$;

-- zero-difference automatic completion
-- zero-difference automatic completion did not retain structured audit metadata
do $$
declare
  v_document_id uuid := '10000000-0000-0000-0000-000000000102';
  v_bank_one_id uuid := '20000000-0000-0000-0000-000000000102';
  v_bank_two_id uuid := '20000000-0000-0000-0000-000000000103';
  v_proposal_id uuid;
  v_reconciliation_id uuid;
  v_result jsonb;
begin
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, amount, fat
  ) values (
    v_document_id, date '2026-10-02', 'AUTO-ZERO-0102', '', '', 200.00, 'S'
  );
  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  ) values
    (v_bank_one_id, 'smoke-execution', 'smoke-auto-zero-0102-a', date '2026-10-02', 'Payment AUTOZERO0102', -80.00),
    (v_bank_two_id, 'smoke-execution', 'smoke-auto-zero-0102-b', date '2026-10-03', 'Payment AUTOZERO0102', -120.00);

  v_proposal_id := pg_temp.make_automatic_proposal(
    v_document_id, 'smoke:automatic-zero', '30000000-0000-0000-0000-000000000102'
  );
  v_result := public.execute_financial_reconciliation_automatic_proposal(
    v_proposal_id, 'smoke:automatic-zero'
  );
  v_reconciliation_id := (v_result->>'reconciliationId')::uuid;

  if not exists (
    select 1 from public.financial_reconciliations
    where id = v_reconciliation_id and status = 'complete'
      and completion_type = 'normal' and difference_amount = 0
      and forced_completion_comment is null and origin = 'automatic'
  )
    or (select count(*) from public.financial_reconciliation_items where reconciliation_id = v_reconciliation_id) <> 3
    or not exists (
      select 1 from public.financial_reconciliation_audit
      where reconciliation_id = v_reconciliation_id and action = 'automatic_complete'
        and comment is null
        and metadata ?& array[
          'ruleSnapshot','configSnapshot','operatorSnapshot','identityEvidence',
          'proposalSignature','trigger','runId','tolerance'
        ]
    ) then
    raise exception 'Zero-difference automatic completion did not retain structured audit metadata.';
  end if;
end $$;

-- stale amount/date/lock/rule/operator/evidence
-- stale automatic proposal created a reconciliation
do $$
declare
  v_kind text;
  v_document_id uuid;
  v_bank_id uuid;
  v_manual_reconciliation_id uuid;
  v_proposal_id uuid;
  v_result jsonb;
  v_doc_number text;
  v_original_definition jsonb;
begin
  select definition into strict v_original_definition
  from public.financial_reconciliation_automatic_rule_definitions
  where rule_key = 'financial_documents_cgd_bank_statement' and version = 1;

  foreach v_kind in array array['amount','date','lock','rule','operator','evidence'] loop
    v_document_id := gen_random_uuid();
    v_bank_id := gen_random_uuid();
    v_doc_number := 'STALE' || upper(v_kind) || replace(v_document_id::text, '-', '');
    insert into public.financial_documents (
      id, document_date, doc_number, description, supplier_name, amount, fat
    ) values (
      v_document_id, date '2026-11-01', v_doc_number, '',
      case when v_kind = 'evidence' then 'Evidence Supplier' else '' end,
      50.00, 'S'
    );
    insert into public.import_cgd_extrato_ordem (
      id, import_batch, row_key, data, descritivo, montante
    ) values (
      v_bank_id, 'smoke-stale', 'smoke-stale-' || v_kind || '-' || v_bank_id,
      date '2026-11-01', 'Payment ' || v_doc_number || case when v_kind = 'evidence' then ' Evidence Supplier' else '' end,
      -50.00
    );
    v_proposal_id := pg_temp.make_automatic_proposal(
      v_document_id, 'smoke:stale-' || v_kind, gen_random_uuid()
    );

    if v_kind = 'amount' then
      update public.financial_documents set amount = 51.00 where id = v_document_id;
    elsif v_kind = 'date' then
      update public.financial_documents set document_date = date '2026-12-01' where id = v_document_id;
    elsif v_kind = 'lock' then
      v_result := public.financial_reconciliation_action(
        'start', 'smoke:stale-lock', null, 'import_cgd_extrato_ordem', v_bank_id, null
      );
      v_manual_reconciliation_id := (v_result#>>'{reconciliation,id}')::uuid;
    elsif v_kind = 'rule' then
      update public.financial_reconciliation_automatic_rule_definitions
      set definition = definition || '{"smokeChanged":true}'::jsonb
      where rule_key = 'financial_documents_cgd_bank_statement' and version = 1;
    elsif v_kind = 'operator' then
      update public.financial_reconciliation_source_rules set operator = '-'
      where base_source_type = 'financial_documents'
        and matching_source_type = 'import_cgd_extrato_ordem';
    elsif v_kind = 'evidence' then
      update public.financial_documents set supplier_name = 'Changed Identity'
      where id = v_document_id;
    end if;

    v_result := public.execute_financial_reconciliation_automatic_proposal(
      v_proposal_id, 'smoke:stale-' || v_kind
    );
    if v_result->>'status' <> 'stale'
      or not exists (
        select 1 from public.financial_reconciliation_automatic_proposals
        where id = v_proposal_id and status = 'stale'
      )
      or exists (
        select 1 from public.financial_reconciliations
        where automatic_proposal_id = v_proposal_id
      ) then
      raise exception 'Stale automatic proposal created a reconciliation.';
    end if;

    if v_kind = 'lock' then
      perform public.financial_reconciliation_action(
        'delete', 'smoke:stale-lock', v_manual_reconciliation_id, null, null, null
      );
    elsif v_kind = 'rule' then
      update public.financial_reconciliation_automatic_rule_definitions
      set definition = v_original_definition
      where rule_key = 'financial_documents_cgd_bank_statement' and version = 1;
    elsif v_kind = 'operator' then
      update public.financial_reconciliation_source_rules set operator = '+'
      where base_source_type = 'financial_documents'
        and matching_source_type = 'import_cgd_extrato_ordem';
    end if;
  end loop;
end $$;

-- post-write rollback and later-proposal isolation
-- failed proposal left partial lifecycle mutations
-- failed proposal was not persisted as failed
-- later proposal was blocked by an earlier failed RPC transaction
create or replace function pg_temp.reject_automatic_complete()
returns trigger language plpgsql as $trigger$
begin
  if new.action = 'automatic_complete' and new.actor = 'smoke:rollback' then
    raise exception 'Smoke forced automatic audit failure.';
  end if;
  return new;
end $trigger$;

create trigger reconciliation_automatic_rollback_smoke
  before insert on public.financial_reconciliation_audit
  for each row execute function pg_temp.reject_automatic_complete();

do $$
declare
  v_failed_document_id uuid := '10000000-0000-0000-0000-000000000103';
  v_failed_bank_id uuid := '20000000-0000-0000-0000-000000000104';
  v_later_document_id uuid := '10000000-0000-0000-0000-000000000104';
  v_later_bank_id uuid := '20000000-0000-0000-0000-000000000105';
  v_failed_proposal_id uuid;
  v_later_proposal_id uuid;
  v_run_id uuid;
  v_result jsonb;
begin
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, amount, fat
  ) values
    (v_failed_document_id, date '2026-12-10', 'AUTO-ROLLBACK-0103', '', '', 75.00, 'S'),
    (v_later_document_id, date '2026-12-11', 'AUTO-LATER-0104', '', '', 80.00, 'S');
  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  ) values
    (v_failed_bank_id, 'smoke-rollback', 'smoke-auto-rollback-0103', date '2026-12-10', 'Payment AUTOROLLBACK0103', -75.00),
    (v_later_bank_id, 'smoke-rollback', 'smoke-auto-later-0104', date '2026-12-11', 'Payment AUTOLATER0104', -80.00);

  v_failed_proposal_id := pg_temp.make_automatic_proposal(
    v_failed_document_id, 'smoke:rollback-run', '30000000-0000-0000-0000-000000000103'
  );
  select run_id into strict v_run_id
  from public.financial_reconciliation_automatic_proposals
  where id = v_failed_proposal_id;
  select id into strict v_later_proposal_id
  from public.financial_reconciliation_automatic_proposals
  where run_id = v_run_id and base_source_id = v_later_document_id and status = 'proposed';

  v_result := public.execute_financial_reconciliation_automatic_proposal(
    v_failed_proposal_id, 'smoke:rollback'
  );

  if v_result->>'status' <> 'failed'
    or exists (
      select 1 from public.financial_reconciliations
      where automatic_proposal_id = v_failed_proposal_id
    )
    or exists (
      select 1 from public.financial_reconciliation_items
      where (source_type = 'financial_documents' and source_id = v_failed_document_id)
         or (source_type = 'import_cgd_extrato_ordem' and source_id = v_failed_bank_id)
    )
    or not exists (
      select 1 from public.financial_reconciliation_automatic_proposals
      where id = v_failed_proposal_id and status = 'failed'
    ) then
    raise exception 'Failed proposal left partial lifecycle mutations.';
  end if;

  v_result := public.execute_financial_reconciliation_automatic_proposal(
    v_later_proposal_id, 'smoke:later'
  );
  if v_result->>'status' <> 'completed'
    or not exists (
      select 1 from public.financial_reconciliation_automatic_proposals
      where id = v_later_proposal_id and status = 'completed'
    ) then
    raise exception 'Later proposal was blocked by an earlier failed RPC transaction.';
  end if;

  v_result := public.finish_financial_reconciliation_automatic_run(v_run_id);
  if v_result->>'status' <> 'partial'
    or v_result#>>'{counts,completed}' <> '1'
    or v_result#>>'{counts,failed}' <> '1'
    or v_result#>>'{counts,deselected}' <> '0' then
    raise exception 'Failed proposal was not persisted as failed.';
  end if;
end $$;

drop trigger reconciliation_automatic_rollback_smoke on public.financial_reconciliation_audit;

-- automatic run finalization
do $$
declare
  v_partial_run_id uuid;
  v_failed_run_id uuid;
  v_completed_run_id uuid;
  v_result jsonb;
begin
  insert into public.financial_reconciliation_automatic_runs (
    trigger, scope, actor, client_request_id, analysis_completed_at
  ) values (
    'manual', 'rule', 'smoke:finish-partial', gen_random_uuid(), now()
  ) returning id into v_partial_run_id;
  insert into public.financial_reconciliation_automatic_proposals (
    run_id, rule_key, rule_version, base_source_type, base_source_id,
    base_source_date, allowed_difference, status, signature
  ) values
    (v_partial_run_id, 'financial_documents_cgd_bank_statement', 1, 'financial_documents', gen_random_uuid(), date '2026-12-20', 0, 'completed', 'finish-partial-completed'),
    (v_partial_run_id, 'financial_documents_cgd_bank_statement', 1, 'financial_documents', gen_random_uuid(), date '2026-12-20', 0, 'stale', 'finish-partial-stale');
  v_result := public.finish_financial_reconciliation_automatic_run(v_partial_run_id);
  if v_result->>'status' <> 'partial'
    or v_result#>>'{counts,completed}' <> '1'
    or v_result#>>'{counts,stale}' <> '1'
    or v_result->>'finishedAt' is null then
    raise exception 'Mixed automatic outcomes did not finalize as partial.';
  end if;

  insert into public.financial_reconciliation_automatic_runs (
    trigger, scope, actor, client_request_id, analysis_completed_at
  ) values (
    'manual', 'rule', 'smoke:finish-failed', gen_random_uuid(), now()
  ) returning id into v_failed_run_id;
  insert into public.financial_reconciliation_automatic_proposals (
    run_id, rule_key, rule_version, base_source_type, base_source_id,
    base_source_date, allowed_difference, status, signature
  ) values (
    v_failed_run_id, 'financial_documents_cgd_bank_statement', 1, 'financial_documents',
    gen_random_uuid(), date '2026-12-21', 0, 'failed', 'finish-failed'
  );
  v_result := public.finish_financial_reconciliation_automatic_run(v_failed_run_id);
  if v_result->>'status' <> 'failed' or v_result#>>'{counts,failed}' <> '1' then
    raise exception 'Failed-only automatic outcomes did not finalize as failed.';
  end if;

  insert into public.financial_reconciliation_automatic_runs (
    trigger, scope, actor, client_request_id, analysis_completed_at
  ) values (
    'manual', 'rule', 'smoke:finish-completed', gen_random_uuid(), now()
  ) returning id into v_completed_run_id;
  insert into public.financial_reconciliation_automatic_proposals (
    run_id, rule_key, rule_version, base_source_type, base_source_id,
    base_source_date, allowed_difference, status, signature
  ) values
    (v_completed_run_id, 'financial_documents_cgd_bank_statement', 1, 'financial_documents', gen_random_uuid(), date '2026-12-22', 0, 'ambiguous', 'finish-completed-ambiguous'),
    (v_completed_run_id, 'financial_documents_cgd_bank_statement', 1, 'financial_documents', gen_random_uuid(), date '2026-12-22', 0, 'proposed', 'finish-completed-deselected');
  v_result := public.finish_financial_reconciliation_automatic_run(v_completed_run_id);
  if v_result->>'status' <> 'completed'
    or v_result#>>'{counts,ambiguous}' <> '1'
    or v_result#>>'{counts,deselected}' <> '1' then
    raise exception 'Expected skips did not finalize as completed.';
  end if;
end $$;

update public.financial_reconciliation_automatic_rule_configs
set enabled = true,
    allow_manual_execution = true,
    include_in_scheduled_batch = false,
    difference_allowed = 4.56,
    max_difference_days = 11,
    priority = 1,
    updated_by = 'smoke:banco-v2'
where rule_key = 'financial_documents_cgd_bank_statement';

\ir ../supabase-migrations/2026-08-16-financial-reconciliation-automation-banco-v2.sql
\ir ../supabase-migrations/2026-08-16-financial-reconciliation-automation-banco-v2.sql
\ir ../supabase-migrations/2026-08-16-financial-reconciliation-automation-90-day-performance.sql
\ir ../supabase-migrations/2026-08-16-financial-reconciliation-automation-90-day-performance.sql

insert into public.financial_documents (
  id, document_date, doc_number, description, supplier_name, payment, amount, fat
) values (
  '44000000-0000-0000-0000-000000000901', date '2027-08-01',
  'BANK-DISPATCH-901', '', '', 'Banco', 901.00, 'S'
);
insert into public.import_cgd_extrato_ordem (
  id, import_batch, row_key, data, descritivo, montante
) values (
  '45000000-0000-0000-0000-000000000901', 'smoke-credit-card',
  'bank-dispatch-901', date '2027-08-01', 'Payment BANKDISPATCH901', -901.00
);
create temporary table credit_card_bank_v2_baseline on commit drop as
select to_jsonb(candidate) as row_snapshot
from public.financial_reconciliation_automatic_candidates_for_base_ids(
  'financial_documents_cgd_bank_statement', 2, 0, 10,
  array['44000000-0000-0000-0000-000000000901'::uuid]
) candidate;

\ir ../supabase-migrations/2026-08-16-financial-reconciliation-automation-credit-card-rule.sql

-- credit-card immutable definition and first config
-- credit-card source rule
do $$
declare
  v_definition jsonb := '{
    "baseEligibility":{"payment":{"operator":"exact_text_equal","value":"Visa","caseSensitive":true,"trim":false}},
    "identityBranches":{"document_number":{"algorithm":"symmetric_compact_containment"},"description_similarity":{"algorithm":"similarity"},"supplier_similarity":{"algorithm":"word_similarity"}},
    "documentNumberMinimumCompactLength":4,
    "descriptionSimilarityThreshold":0.55,
    "supplierWordSimilarityThreshold":0.60,
    "maxDestinationRecords":4,
    "maxIdentityCandidatesPerBase":12
  }'::jsonb;
  v_logic text := 'Payment must equal exactly Visa. Each credit-card candidate must satisfy invoice containment, description similarity, or supplier word similarity. Exactly one one-to-four-record amount combination is executable.';
begin
  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_definitions definition
    where definition.rule_key = 'financial_documents_cgd_credit_card'
      and definition.version = 1
      and definition.display_name = 'Financial Documents to CGD Credit Card'
      and definition.base_source_type = 'financial_documents'
      and definition.destination_source_types = '["import_cgd_cartao_credito"]'::jsonb
      and definition.logic_description = v_logic
      and definition.definition = v_definition
  ) then
    raise exception 'Credit-card immutable definition differs from the approved literal.';
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs config
    where config.rule_key = 'financial_documents_cgd_credit_card'
      and config.rule_version = 1
      and not config.enabled
      and not config.allow_manual_execution
      and not config.include_in_scheduled_batch
      and config.difference_allowed = 0
      and config.max_difference_days = 10
      and config.priority = 2
  ) or not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs config
    where config.rule_key = 'financial_documents_cgd_bank_statement'
      and config.priority = 1
  ) then
    raise exception 'Credit-card config was not inserted disabled at priority 2 after Banco.';
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_source_rules source_rule
    where source_rule.base_source_type = 'financial_documents'
      and source_rule.matching_source_type = 'import_cgd_cartao_credito'
      and source_rule.operator = '+'
  ) then
    raise exception 'Credit-card directional source rule is not financial_documents to import_cgd_cartao_credito (+).';
  end if;
end $$;

-- managed automatic source rules reject operator changes and deletion
do $$
declare
  v_rules jsonb;
  v_rejected boolean;
begin
  select coalesce(jsonb_agg(jsonb_build_object(
    'base_source_type', base_source_type,
    'matching_source_type', matching_source_type,
    'operator', operator
  ) order by base_source_type, matching_source_type), '[]'::jsonb)
  into v_rules
  from public.financial_reconciliation_source_rules;

  v_rejected := false;
  begin
    perform public.replace_financial_reconciliation_source_rules((
      select jsonb_agg(case
        when rule->>'base_source_type' = 'financial_documents'
         and rule->>'matching_source_type' = 'import_cgd_cartao_credito'
          then jsonb_set(rule, '{operator}', '"-"'::jsonb)
        else rule
      end)
      from jsonb_array_elements(v_rules) rule
    ));
  exception when others then
    v_rejected := sqlerrm =
      'The managed Credit Card source rule must remain enabled with operator +.';
  end;
  if not v_rejected then
    raise exception 'Managed Credit Card source-rule operator change was accepted.';
  end if;

  v_rejected := false;
  begin
    perform public.replace_financial_reconciliation_source_rules((
      select coalesce(jsonb_agg(rule), '[]'::jsonb)
      from jsonb_array_elements(v_rules) rule
      where not (
        rule->>'base_source_type' = 'financial_documents'
        and rule->>'matching_source_type' = 'import_cgd_cartao_credito'
      )
    ));
  exception when others then
    v_rejected := sqlerrm =
      'The managed Credit Card source rule must remain enabled with operator +.';
  end;
  if not v_rejected then
    raise exception 'Managed Credit Card source-rule deletion was accepted.';
  end if;

  v_rejected := false;
  begin
    perform public.replace_financial_reconciliation_source_rules((
      select jsonb_agg(case
        when rule->>'base_source_type' = 'financial_documents'
         and rule->>'matching_source_type' = 'import_cgd_extrato_ordem'
          then jsonb_set(rule, '{operator}', '"-"'::jsonb)
        else rule
      end)
      from jsonb_array_elements(v_rules) rule
    ));
  exception when others then
    v_rejected := sqlerrm =
      'The managed Bank Statement source rule must remain enabled with operator +.';
  end;
  if not v_rejected then
    raise exception 'Managed Bank Statement source-rule operator change was accepted.';
  end if;

  v_rejected := false;
  begin
    perform public.replace_financial_reconciliation_source_rules((
      select coalesce(jsonb_agg(rule), '[]'::jsonb)
      from jsonb_array_elements(v_rules) rule
      where not (
        rule->>'base_source_type' = 'financial_documents'
        and rule->>'matching_source_type' = 'import_cgd_extrato_ordem'
      )
    ));
  exception when others then
    v_rejected := sqlerrm =
      'The managed Bank Statement source rule must remain enabled with operator +.';
  end;
  if not v_rejected then
    raise exception 'Managed Bank Statement source-rule deletion was accepted.';
  end if;
end $$;

insert into public.import_cgd_cartao_credito (
  id, import_batch, row_key, data, data_valor, descricao, debito
) values (
  '46000000-0000-0000-0000-000000000990', 'smoke-credit-card',
  'credit-card-reapply-990', date '2027-08-02', date '2027-08-03',
  'Credit card reapply projection', 9.90
);
update public.financial_reconciliation_automatic_rule_configs
set enabled = true,
    allow_manual_execution = true,
    include_in_scheduled_batch = true,
    difference_allowed = 3.21,
    max_difference_days = 12,
    priority = 3,
    updated_by = 'smoke:credit-card-admin'
where rule_key = 'financial_documents_cgd_credit_card';

\ir ../supabase-migrations/2026-08-16-financial-reconciliation-automation-credit-card-rule.sql

-- Banco v2 dispatcher IDs and evidence remain byte-for-byte unchanged
-- credit-card migration reapply is idempotent and preserves administrator settings
do $$
declare
  v_before text;
  v_after text;
begin
  select row_snapshot::text into strict v_before
  from credit_card_bank_v2_baseline;
  select to_jsonb(candidate)::text into strict v_after
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_bank_statement', 2, 0, 10,
    array['44000000-0000-0000-0000-000000000901'::uuid]
  ) candidate;
  if convert_to(v_before, 'UTF8') is distinct from convert_to(v_after, 'UTF8') then
    raise exception 'Banco v2 dispatcher IDs or evidence changed byte-for-byte.';
  end if;
  if (select row_snapshot#>>'{base_source_id}' from credit_card_bank_v2_baseline)
      <> '44000000-0000-0000-0000-000000000901'
    or (select row_snapshot#>>'{candidates,0,sourceId}' from credit_card_bank_v2_baseline)
      <> '45000000-0000-0000-0000-000000000901'
    or (select (row_snapshot#>>'{candidates,0,evidence,documentNumber,matched}')::boolean
        from credit_card_bank_v2_baseline) is not true
    or (select row_snapshot#>>'{candidates,0,evidence,documentNumber,normalized}'
        from credit_card_bank_v2_baseline) <> 'bankdispatch901'
    or (select row_snapshot#>>'{candidates,0,evidence,description,threshold}'
        from credit_card_bank_v2_baseline) <> '0.60'
    or (select row_snapshot#>>'{candidates,0,evidence,supplier,threshold}'
        from credit_card_bank_v2_baseline) <> '0.70' then
    raise exception 'Banco v2 dispatcher fixture lost its literal IDs or evidence.';
  end if;

  if (select count(*) from public.financial_reconciliation_automatic_rule_definitions
      where rule_key = 'financial_documents_cgd_credit_card' and version = 1) <> 1
    or (select count(*) from public.financial_reconciliation_automatic_rule_configs
        where rule_key = 'financial_documents_cgd_credit_card') <> 1
    or (select count(*) from pg_trigger
        where tgrelid = 'public.import_cgd_cartao_credito'::regclass
          and tgname = 'financial_reconciliation_sync_cgd_credit_card_match_search_trigger'
          and not tgisinternal) <> 1
    or (select count(*) from pg_indexes
        where schemaname = 'public' and indexname in (
          'financial_reconciliation_cgd_credit_card_match_search_date_id_idx',
          'financial_reconciliation_cgd_credit_card_match_search_normalized_trgm_idx',
          'financial_reconciliation_cgd_credit_card_match_search_compact_trgm_idx'
        )) <> 3
    or (select count(*) from public.financial_reconciliation_cgd_credit_card_match_search
        where source_id = '46000000-0000-0000-0000-000000000990') <> 1 then
    raise exception 'Credit-card migration reapply duplicated a managed row, projection, trigger, or index.';
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs config
    where config.rule_key = 'financial_documents_cgd_credit_card'
      and config.rule_version = 1
      and config.enabled
      and config.allow_manual_execution
      and config.include_in_scheduled_batch
      and config.difference_allowed = 3.21
      and config.max_difference_days = 12
      and config.priority = 3
      and config.updated_by = 'smoke:credit-card-admin'
  ) then
    raise exception 'Credit-card migration reapply overwrote administrator settings or priority.';
  end if;

  update public.financial_reconciliation_automatic_rule_configs
  set enabled = false,
      allow_manual_execution = false,
      include_in_scheduled_batch = false,
      difference_allowed = 0,
      max_difference_days = 10,
      priority = 2,
      updated_by = ''
  where rule_key = 'financial_documents_cgd_credit_card';
end $$;

-- amount-only managed definitions, deterministic configs, and fixed-zero Settings contract
create temporary table amount_only_existing_priority_baseline on commit drop as
select config.rule_key, config.priority
from public.financial_reconciliation_automatic_rule_configs config;

\ir ../supabase-migrations/2026-08-17-financial-reconciliation-automation-amount-only-rules.sql
\ir ../supabase-migrations/2026-08-17-financial-reconciliation-automation-amount-only-rules.sql

do $$
declare
  v_bank_definition jsonb := '{
    "baseSourceType":"financial_documents",
    "destinationSourceTypes":["import_cgd_extrato_ordem"],
    "baseEligibility":{"payment":{"operator":"exact_text_equal","value":"Banco","caseSensitive":true,"trim":false}},
    "matchingMode":"amount_only_one_to_one",
    "fixedDifferenceAllowed":0,
    "maxDifferenceDays":{"minimum":0,"maximum":90,"default":1},
    "maxDestinationRecords":1
  }'::jsonb;
  v_card_definition jsonb := '{
    "baseSourceType":"financial_documents",
    "destinationSourceTypes":["import_cgd_cartao_credito"],
    "baseEligibility":{"payment":{"operator":"exact_text_equal","value":"Visa","caseSensitive":true,"trim":false}},
    "matchingMode":"amount_only_one_to_one",
    "fixedDifferenceAllowed":0,
    "maxDifferenceDays":{"minimum":0,"maximum":90,"default":1},
    "maxDestinationRecords":1
  }'::jsonb;
  v_settings jsonb;
  v_schedule jsonb;
  v_rules jsonb;
  v_before jsonb;
  v_after jsonb;
  v_key text;
  v_rejected boolean;
  v_signature text;
begin
  if not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_definitions definition
      where definition.rule_key = 'financial_documents_cgd_bank_statement_amount_only'
        and definition.version = 1
        and definition.display_name = 'Financial Documents to CGD Bank Account – AMOUNT ONLY'
        and definition.base_source_type = 'financial_documents'
        and definition.destination_source_types = '["import_cgd_extrato_ordem"]'::jsonb
        and definition.logic_description = 'Payment must equal exactly Banco. Exactly one CGD Bank Account destination record must make the signed amounts sum to zero within the inclusive configured date window; identity fields and similarity are not used.'
        and definition.definition = v_bank_definition
    ) or not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_definitions definition
      where definition.rule_key = 'financial_documents_cgd_credit_card_amount_only'
        and definition.version = 1
        and definition.display_name = 'Financial Documents to CGD Credit Card – AMOUNT ONLY'
        and definition.base_source_type = 'financial_documents'
        and definition.destination_source_types = '["import_cgd_cartao_credito"]'::jsonb
        and definition.logic_description = 'Payment must equal exactly Visa. Exactly one CGD Credit Card destination record must make the signed amounts sum to zero within the inclusive configured date window; identity fields and similarity are not used.'
        and definition.definition = v_card_definition
    ) then
    raise exception 'Amount-only immutable definitions differ from the approved literals.';
  end if;

  if not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_configs config
      where config.rule_key = 'financial_documents_cgd_bank_statement_amount_only'
        and config.rule_version = 1
        and not config.enabled
        and not config.allow_manual_execution
        and not config.include_in_scheduled_batch
        and config.difference_allowed = 0
        and config.max_difference_days = 1
        and config.priority = 3
    ) or not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_configs config
      where config.rule_key = 'financial_documents_cgd_credit_card_amount_only'
        and config.rule_version = 1
        and not config.enabled
        and not config.allow_manual_execution
        and not config.include_in_scheduled_batch
        and config.difference_allowed = 0
        and config.max_difference_days = 1
        and config.priority = 4
    ) then
    raise exception 'Amount-only configs were not appended disabled at standard priorities 3 and 4.';
  end if;

  if exists (
    select baseline.rule_key
    from amount_only_existing_priority_baseline baseline
    join public.financial_reconciliation_automatic_rule_configs config
      on config.rule_key = baseline.rule_key
    where config.priority <> baseline.priority
  ) then
    raise exception 'Amount-only migration rewrote a pre-existing administrator priority.';
  end if;

  if exists (
      select 1
      from public.financial_reconciliation_automatic_rule_configs amount_config
      join public.financial_reconciliation_automatic_rule_configs existing_config
        on existing_config.rule_key not in (
          'financial_documents_cgd_bank_statement_amount_only',
          'financial_documents_cgd_credit_card_amount_only'
        )
      where amount_config.rule_key in (
          'financial_documents_cgd_bank_statement_amount_only',
          'financial_documents_cgd_credit_card_amount_only'
        )
        and amount_config.priority <= existing_config.priority
    ) or (select priority from public.financial_reconciliation_automatic_rule_configs
          where rule_key = 'financial_documents_cgd_bank_statement_amount_only') >=
         (select priority from public.financial_reconciliation_automatic_rule_configs
          where rule_key = 'financial_documents_cgd_credit_card_amount_only') then
    raise exception 'Amount-only configs were not appended in Bank-then-Card order.';
  end if;

  if (select count(*)
      from public.financial_reconciliation_source_rules source_rule
      where source_rule.base_source_type = 'financial_documents'
        and source_rule.operator = '+'
        and source_rule.matching_source_type in (
          'import_cgd_extrato_ordem', 'import_cgd_cartao_credito'
        )) <> 2 then
    raise exception 'Amount-only managed directional source operators are not fixed to +.';
  end if;

  if (select count(*) from pg_indexes
      where schemaname = 'public'
        and indexname in (
          'import_cgd_extrato_ordem_reconciliation_amount_date_id_idx',
          'import_cgd_cartao_credito_reconciliation_amount_date_id_idx'
        )) <> 2 then
    raise exception 'Amount-only destination lookup indexes are missing.';
  end if;
  if not exists (
      select 1
      from pg_index index_row
      where index_row.indexrelid = 'public.import_cgd_extrato_ordem_reconciliation_amount_date_id_idx'::regclass
        and pg_get_indexdef(index_row.indexrelid, 1, true) = 'montante'
        and pg_get_indexdef(index_row.indexrelid, 2, true) = 'data'
        and pg_get_indexdef(index_row.indexrelid, 3, true) = 'id'
    ) or not exists (
      select 1
      from pg_index index_row
      where index_row.indexrelid = 'public.import_cgd_cartao_credito_reconciliation_amount_date_id_idx'::regclass
        and pg_get_indexdef(index_row.indexrelid, 1, true) = 'valor'
        and pg_get_indexdef(index_row.indexrelid, 2, true) = 'data'
        and pg_get_indexdef(index_row.indexrelid, 3, true) = 'id'
    ) then
    raise exception 'Amount-only destination lookup indexes use the wrong key order.';
  end if;

  v_settings := public.get_financial_reconciliation_automation_settings();
  if jsonb_array_length(v_settings->'rules') <> 4
    or exists (
      select 1
      from jsonb_array_elements(v_settings->'rules') rule
      group by rule->>'ruleKey'
      having count(*) <> 1
    )
    or (select count(*) from jsonb_array_elements(v_settings->'rules') rule
        where rule->>'ruleKey' in (
          'financial_documents_cgd_bank_statement',
          'financial_documents_cgd_credit_card',
          'financial_documents_cgd_bank_statement_amount_only',
          'financial_documents_cgd_credit_card_amount_only'
        )) <> 4 then
    raise exception 'Settings getter did not return each of the four managed rules exactly once.';
  end if;

  foreach v_signature in array array[
    'public.get_financial_reconciliation_automation_settings()',
    'public.replace_financial_reconciliation_automation_settings(jsonb,jsonb,text)'
  ] loop
    if not (
      select procedure.prosecdef
        and coalesce(procedure.proconfig, '{}'::text[]) @> array['search_path=public, pg_temp']
      from pg_proc procedure
      where procedure.oid = v_signature::regprocedure
    )
      or has_function_privilege('anon', v_signature, 'EXECUTE')
      or has_function_privilege('authenticated', v_signature, 'EXECUTE')
      or not has_function_privilege('service_role', v_signature, 'EXECUTE') then
      raise exception 'Amount-only Settings function security is invalid for %.', v_signature;
    end if;
  end loop;

  select jsonb_build_object(
    'enabled', schedule.enabled,
    'time_of_day', to_char(schedule.time_of_day, 'HH24:MI'),
    'time_zone', schedule.time_zone
  ) into strict v_schedule
  from public.financial_reconciliation_automatic_schedule schedule
  where schedule.id = true;

  select jsonb_agg(jsonb_build_object(
    'rule_key', config.rule_key,
    'rule_version', config.rule_version,
    'enabled', case when config.rule_key = 'financial_documents_cgd_bank_statement_amount_only' then true else config.enabled end,
    'allow_manual_execution', case when config.rule_key = 'financial_documents_cgd_bank_statement_amount_only' then true else config.allow_manual_execution end,
    'include_in_scheduled_batch', case when config.rule_key = 'financial_documents_cgd_bank_statement_amount_only' then true else config.include_in_scheduled_batch end,
    'difference_allowed', to_char(config.difference_allowed, 'FM999999999990.00'),
    'max_difference_days', case config.rule_key
      when 'financial_documents_cgd_bank_statement_amount_only' then 15
      when 'financial_documents_cgd_credit_card_amount_only' then 90
      else config.max_difference_days end,
    'priority', case config.rule_key
      when 'financial_documents_cgd_bank_statement_amount_only' then 4
      when 'financial_documents_cgd_credit_card_amount_only' then 3
      else config.priority end
  ) order by config.priority, config.rule_key) into strict v_rules
  from public.financial_reconciliation_automatic_rule_configs config;

  perform public.replace_financial_reconciliation_automation_settings(
    v_schedule, v_rules, 'smoke:amount-only-admin'
  );
  if not exists (
      select 1 from public.financial_reconciliation_automatic_rule_configs
      where rule_key = 'financial_documents_cgd_bank_statement_amount_only'
        and enabled and allow_manual_execution and include_in_scheduled_batch
        and difference_allowed = 0 and max_difference_days = 15 and priority = 4
        and updated_by = 'smoke:amount-only-admin'
    ) or not exists (
      select 1 from public.financial_reconciliation_automatic_rule_configs
      where rule_key = 'financial_documents_cgd_credit_card_amount_only'
        and difference_allowed = 0 and max_difference_days = 90 and priority = 3
        and updated_by = 'smoke:amount-only-admin'
    ) then
    raise exception 'Settings replacement did not accept amount-only days, priority, and enablement edits.';
  end if;

  foreach v_key in array array[
    'financial_documents_cgd_bank_statement_amount_only',
    'financial_documents_cgd_credit_card_amount_only'
  ] loop
    select public.get_financial_reconciliation_automation_settings() into v_before;
    v_rejected := false;
    begin
      perform public.replace_financial_reconciliation_automation_settings(
        v_schedule,
        (
          select jsonb_agg(case when rule->>'rule_key' = v_key
            then jsonb_set(rule, '{difference_allowed}', '"0.01"'::jsonb)
            else rule end)
          from jsonb_array_elements(v_rules) rule
        ),
        'smoke:amount-only-nonzero'
      );
    exception when raise_exception then
      if sqlerrm = 'Amount-only automatic rules require zero difference allowed.' then
        v_rejected := true;
      else
        raise;
      end if;
    end;
    select public.get_financial_reconciliation_automation_settings() into v_after;
    if not v_rejected or v_after is distinct from v_before then
      raise exception 'Nonzero tolerance for % was accepted or partially persisted.', v_key;
    end if;

    v_rejected := false;
    begin
      perform public.replace_financial_reconciliation_automation_settings(
        v_schedule,
        (select jsonb_agg(rule) from jsonb_array_elements(v_rules) rule
         where rule->>'rule_key' <> v_key),
        'smoke:amount-only-missing'
      );
    exception when raise_exception then
      if sqlerrm = 'Automation settings require every managed rule exactly once.' then
        v_rejected := true;
      else
        raise;
      end if;
    end;
    select public.get_financial_reconciliation_automation_settings() into v_after;
    if not v_rejected or v_after is distinct from v_before then
      raise exception 'Missing managed rule % was accepted or partially persisted.', v_key;
    end if;
  end loop;

  select public.get_financial_reconciliation_automation_settings() into v_before;
  v_rejected := false;
  begin
    perform public.replace_financial_reconciliation_automation_settings(
      v_schedule,
      (
        select jsonb_agg(case
          when rule->>'rule_key' = 'financial_documents_cgd_bank_statement_amount_only'
            then jsonb_set(rule, '{max_difference_days}', '91'::jsonb)
          else rule end)
        from jsonb_array_elements(v_rules) rule
      ),
      'smoke:amount-only-days'
    );
  exception when raise_exception then
    if sqlerrm = 'Automatic rule values are invalid.' then
      v_rejected := true;
    else
      raise;
    end if;
  end;
  select public.get_financial_reconciliation_automation_settings() into v_after;
  if not v_rejected or v_after is distinct from v_before then
    raise exception 'Amount-only Settings accepted a date window above 90 or partially persisted it.';
  end if;
end $$;

\ir ../supabase-migrations/2026-08-17-financial-reconciliation-automation-amount-only-rules.sql

do $$
begin
  if not exists (
    select 1 from public.financial_reconciliation_automatic_rule_configs
    where rule_key = 'financial_documents_cgd_bank_statement_amount_only'
      and enabled and allow_manual_execution and include_in_scheduled_batch
      and difference_allowed = 0 and max_difference_days = 15 and priority = 4
      and updated_by = 'smoke:amount-only-admin'
  ) or not exists (
    select 1 from public.financial_reconciliation_automatic_rule_configs
    where rule_key = 'financial_documents_cgd_credit_card_amount_only'
      and difference_allowed = 0 and max_difference_days = 90 and priority = 3
      and updated_by = 'smoke:amount-only-admin'
  ) then
    raise exception 'Amount-only migration reapply overwrote administrator settings.';
  end if;

  set constraints financial_reconciliation_automatic_rule_configs_priority_key deferred;
  update public.financial_reconciliation_automatic_rule_configs config
  set enabled = false,
      allow_manual_execution = false,
      include_in_scheduled_batch = false,
      difference_allowed = 0,
      max_difference_days = 1,
      priority = case config.rule_key
        when 'financial_documents_cgd_bank_statement_amount_only' then 3
        else 4 end,
      updated_by = ''
  where config.rule_key in (
    'financial_documents_cgd_bank_statement_amount_only',
    'financial_documents_cgd_credit_card_amount_only'
  );
end $$;

-- amount-only four-rule dispatch and exact one-to-one candidate behavior
do $$
declare
  v_expected jsonb;
begin
  v_expected := jsonb_build_object(
    'payment','Banco','destinationSourceType','import_cgd_extrato_ordem',
    'descriptionThreshold',0.60,'supplierThreshold',0.70,
    'maxDestinationRecords',4,'maxCandidates',12
  );
  if public.financial_reconciliation_automatic_rule_contract(
      'financial_documents_cgd_bank_statement', 2
    ) is distinct from v_expected then
    raise exception 'The Banco identity adapter contract changed.';
  end if;

  v_expected := jsonb_build_object(
    'payment','Visa','destinationSourceType','import_cgd_cartao_credito',
    'descriptionThreshold',0.55,'supplierThreshold',0.60,
    'maxDestinationRecords',4,'maxCandidates',12
  );
  if public.financial_reconciliation_automatic_rule_contract(
      'financial_documents_cgd_credit_card', 1
    ) is distinct from v_expected then
    raise exception 'The Visa identity adapter contract changed.';
  end if;

  v_expected := jsonb_build_object(
    'payment','Banco','destinationSourceType','import_cgd_extrato_ordem',
    'matchingMode','amount_only_one_to_one','maxDestinationRecords',1,
    'maxCandidates',12,'fixedDifferenceAllowed',0
  );
  if public.financial_reconciliation_automatic_rule_contract(
      'financial_documents_cgd_bank_statement_amount_only', 1
    ) is distinct from v_expected then
    raise exception 'The Banco amount-only adapter contract is invalid.';
  end if;

  v_expected := jsonb_build_object(
    'payment','Visa','destinationSourceType','import_cgd_cartao_credito',
    'matchingMode','amount_only_one_to_one','maxDestinationRecords',1,
    'maxCandidates',12,'fixedDifferenceAllowed',0
  );
  if public.financial_reconciliation_automatic_rule_contract(
      'financial_documents_cgd_credit_card_amount_only', 1
    ) is distinct from v_expected
    or public.financial_reconciliation_automatic_rule_contract(
      'financial_documents_cgd_bank_statement_amount_only', 2
    ) is not null then
    raise exception 'The Visa amount-only adapter contract or version allowlist is invalid.';
  end if;
end $$;

update public.financial_documents
set payment = case payment
  when 'Banco' then 'smoke:before-amount-only-banco'
  when 'Visa' then 'smoke:before-amount-only-visa'
  else payment end
where payment in ('Banco', 'Visa');

alter table public.financial_documents
  alter column document_date drop not null,
  alter column amount drop not null,
  alter column payment drop not null;

do $$
declare
  v_reconciliation_id uuid;
  v_bank_base_ids uuid[];
  v_card_base_ids uuid[];
  v_candidate jsonb;
  v_candidate_count integer;
begin
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, payment, amount, fat
  ) values
    ('60000000-0000-0000-0000-000000000001', date '2030-06-15', 'BANK-EXACT', 'never similar', 'unrelated supplier', 'Banco', 100.00, 'S'),
    ('60000000-0000-0000-0000-000000000002', date '2030-06-15', 'BANK-CASE', '', '', 'banco', 100.00, 'S'),
    ('60000000-0000-0000-0000-000000000003', date '2030-06-15', 'BANK-PAD', '', '', ' Banco ', 100.00, 'S'),
    ('60000000-0000-0000-0000-000000000004', date '2030-06-15', 'BANK-BLANK', '', '', '', 100.00, 'S'),
    ('60000000-0000-0000-0000-000000000005', date '2030-06-15', 'BANK-NULL-PAYMENT', '', '', null, 100.00, 'S'),
    ('60000000-0000-0000-0000-000000000006', date '2030-06-15', 'BANK-FAT', '', '', 'Banco', 100.00, 'N'),
    ('60000000-0000-0000-0000-000000000007', date '2025-12-31', 'BANK-PRE-2026', '', '', 'Banco', 100.00, 'S'),
    ('60000000-0000-0000-0000-000000000008', null, 'BANK-NULL-DATE', '', '', 'Banco', 100.00, 'S'),
    ('60000000-0000-0000-0000-000000000009', date '2030-06-15', 'BANK-NULL-AMOUNT', '', '', 'Banco', null, 'S'),
    ('60000000-0000-0000-0000-000000000010', date '2030-06-15', 'BANK-LOCKED', '', '', 'Banco', 100.00, 'S'),
    ('60000000-0000-0000-0000-000000000020', date '2026-01-01', 'BANK-DEST-EXCLUSIONS', '', '', 'Banco', 130.00, 'S'),
    ('62000000-0000-0000-0000-000000000001', date '2030-06-15', 'CARD-EXACT', 'never similar', 'unrelated supplier', 'Visa', 100.00, 'S'),
    ('62000000-0000-0000-0000-000000000002', date '2030-06-15', 'CARD-CASE', '', '', 'visa', 100.00, 'S'),
    ('62000000-0000-0000-0000-000000000003', date '2030-06-15', 'CARD-PAD', '', '', ' Visa ', 100.00, 'S'),
    ('62000000-0000-0000-0000-000000000004', date '2030-06-15', 'CARD-BLANK', '', '', '', 100.00, 'S'),
    ('62000000-0000-0000-0000-000000000005', date '2030-06-15', 'CARD-NULL-PAYMENT', '', '', null, 100.00, 'S'),
    ('62000000-0000-0000-0000-000000000006', date '2030-06-15', 'CARD-FAT', '', '', 'Visa', 100.00, 'N'),
    ('62000000-0000-0000-0000-000000000007', date '2025-12-31', 'CARD-PRE-2026', '', '', 'Visa', 100.00, 'S'),
    ('62000000-0000-0000-0000-000000000008', null, 'CARD-NULL-DATE', '', '', 'Visa', 100.00, 'S'),
    ('62000000-0000-0000-0000-000000000009', date '2030-06-15', 'CARD-NULL-AMOUNT', '', '', 'Visa', null, 'S'),
    ('62000000-0000-0000-0000-000000000010', date '2030-06-15', 'CARD-LOCKED', '', '', 'Visa', 100.00, 'S'),
    ('62000000-0000-0000-0000-000000000020', date '2026-01-01', 'CARD-DEST-EXCLUSIONS', '', '', 'Visa', 130.00, 'S');

  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, data_valor, descritivo, montante
  ) values
    ('61000000-0000-0000-0000-000000000001', 'smoke-amount-only', 'amount-bank-exact', date '2030-06-15', date '1999-01-01', 'nothing in common', -100.00),
    ('61000000-0000-0000-0000-000000000020', 'smoke-amount-only', 'amount-bank-pre', date '2025-12-31', date '2026-01-01', 'BANK-DEST-EXCLUSIONS', -130.00),
    ('61000000-0000-0000-0000-000000000021', 'smoke-amount-only', 'amount-bank-null-date', null, date '2026-01-01', 'BANK-DEST-EXCLUSIONS', -130.00),
    ('61000000-0000-0000-0000-000000000022', 'smoke-amount-only', 'amount-bank-null-amount', date '2026-01-01', date '2026-01-01', 'BANK-DEST-EXCLUSIONS', null),
    ('61000000-0000-0000-0000-000000000023', 'smoke-amount-only', 'amount-bank-locked', date '2026-01-01', date '2026-01-01', 'BANK-DEST-EXCLUSIONS', -130.00),
    ('61000000-0000-0000-0000-000000000024', 'smoke-amount-only', 'amount-bank-wrong-date-field', date '2026-01-03', date '2026-01-01', 'BANK-DEST-EXCLUSIONS', -130.00);

  insert into public.import_cgd_cartao_credito (
    id, import_batch, row_key, data, data_valor, descricao, debito, credito
  ) values
    ('63000000-0000-0000-0000-000000000001', 'smoke-amount-only', 'amount-card-exact', date '2030-06-15', date '1999-01-01', 'nothing in common', 100.00, null),
    ('63000000-0000-0000-0000-000000000020', 'smoke-amount-only', 'amount-card-pre', date '2025-12-31', date '2026-01-01', 'CARD-DEST-EXCLUSIONS', 130.00, null),
    ('63000000-0000-0000-0000-000000000021', 'smoke-amount-only', 'amount-card-null-date', null, date '2026-01-01', 'CARD-DEST-EXCLUSIONS', 130.00, null),
    ('63000000-0000-0000-0000-000000000022', 'smoke-amount-only', 'amount-card-empty-source-amount', date '2026-01-01', date '2026-01-01', 'CARD-DEST-EXCLUSIONS', null, null),
    ('63000000-0000-0000-0000-000000000023', 'smoke-amount-only', 'amount-card-locked', date '2026-01-01', date '2026-01-01', 'CARD-DEST-EXCLUSIONS', 130.00, null),
    ('63000000-0000-0000-0000-000000000024', 'smoke-amount-only', 'amount-card-wrong-date-field', date '2026-01-03', date '2026-01-01', 'CARD-DEST-EXCLUSIONS', 130.00, null);

  if (select valor from public.import_cgd_cartao_credito
      where id = '63000000-0000-0000-0000-000000000022') <> 0 then
    raise exception 'Empty card debit/credit inputs no longer generate the schema-defined zero approved amount.';
  end if;

  insert into public.financial_reconciliations (
    status, base_source_type, matching_source_types, created_by
  ) values (
    'started', 'financial_documents',
    '["import_cgd_extrato_ordem","import_cgd_cartao_credito"]'::jsonb,
    'smoke:amount-only-locks'
  ) returning id into v_reconciliation_id;

  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  ) values
    (v_reconciliation_id, 'financial_documents', '60000000-0000-0000-0000-000000000010', 100.00, 'smoke:amount-only-locks'),
    (v_reconciliation_id, 'financial_documents', '62000000-0000-0000-0000-000000000010', 100.00, 'smoke:amount-only-locks'),
    (v_reconciliation_id, 'import_cgd_extrato_ordem', '61000000-0000-0000-0000-000000000023', -130.00, 'smoke:amount-only-locks'),
    (v_reconciliation_id, 'import_cgd_cartao_credito', '63000000-0000-0000-0000-000000000023', -130.00, 'smoke:amount-only-locks');

  select coalesce(array_agg(candidate.base_source_id order by candidate.base_source_id), '{}'::uuid[])
  into v_bank_base_ids
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_bank_statement_amount_only', 1, 0, 1,
    array(select ('60000000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid
          from generate_series(1, 10) series)
  ) candidate;
  if v_bank_base_ids is distinct from array['60000000-0000-0000-0000-000000000001'::uuid] then
    raise exception 'Banco amount-only base eligibility admitted case, space, blank, null, fat, date, amount, or lock exclusions: %', v_bank_base_ids;
  end if;

  select coalesce(array_agg(candidate.base_source_id order by candidate.base_source_id), '{}'::uuid[])
  into v_card_base_ids
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_credit_card_amount_only', 1, 0, 1,
    array(select ('62000000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid
          from generate_series(1, 10) series)
  ) candidate;
  if v_card_base_ids is distinct from array['62000000-0000-0000-0000-000000000001'::uuid] then
    raise exception 'Visa amount-only base eligibility admitted case, space, blank, null, fat, date, amount, or lock exclusions: %', v_card_base_ids;
  end if;

  select coalesce(array_agg(page.id order by page.document_date, page.id), '{}'::uuid[])
  into v_bank_base_ids
  from public.financial_reconciliation_automatic_base_page(
    'financial_documents_cgd_bank_statement_amount_only', 1, null, null, 25
  ) page
  where page.id between '60000000-0000-0000-0000-000000000001'::uuid
                    and '60000000-0000-0000-0000-000000000020'::uuid;
  if v_bank_base_ids is distinct from array[
      '60000000-0000-0000-0000-000000000001'::uuid,
      '60000000-0000-0000-0000-000000000020'::uuid
    ] or public.financial_reconciliation_automatic_base_count(
      'financial_documents_cgd_bank_statement_amount_only', 1
    ) <> 2 then
    raise exception 'Banco amount-only base paging/counting admitted a null amount or other ineligible base: %', v_bank_base_ids;
  end if;

  select coalesce(array_agg(page.id order by page.document_date, page.id), '{}'::uuid[])
  into v_card_base_ids
  from public.financial_reconciliation_automatic_base_page(
    'financial_documents_cgd_credit_card_amount_only', 1, null, null, 25
  ) page
  where page.id between '62000000-0000-0000-0000-000000000001'::uuid
                    and '62000000-0000-0000-0000-000000000020'::uuid;
  if v_card_base_ids is distinct from array[
      '62000000-0000-0000-0000-000000000001'::uuid,
      '62000000-0000-0000-0000-000000000020'::uuid
    ] or public.financial_reconciliation_automatic_base_count(
      'financial_documents_cgd_credit_card_amount_only', 1
    ) <> 2 then
    raise exception 'Visa amount-only base paging/counting admitted a null amount or other ineligible base: %', v_card_base_ids;
  end if;

  select candidates, candidate_count into strict v_candidate, v_candidate_count
  from public.financial_reconciliation_automatic_single_base_candidates(
    'financial_documents_cgd_bank_statement_amount_only', 1, 0, 1,
    '60000000-0000-0000-0000-000000000001'
  );
  if v_candidate_count <> 1
    or jsonb_array_length(v_candidate) <> 1
    or v_candidate->0->>'sourceType' <> 'import_cgd_extrato_ordem'
    or v_candidate->0->>'sourceId' <> '61000000-0000-0000-0000-000000000001'
    or v_candidate->0->>'sourceDate' <> '2030-06-15'
    or (v_candidate->0->>'amount')::numeric <> -100.00
    or v_candidate#>>'{0,evidence,amount,baseAmountCents}' <> '10000'
    or v_candidate#>>'{0,evidence,amount,destinationAmountCents}' <> '-10000'
    or v_candidate#>>'{0,evidence,amount,signedDifferenceCents}' <> '0'
    or v_candidate#>>'{0,evidence,date,distanceDays}' <> '0'
    or v_candidate#>'{0,evidence}' ?| array['documentNumber','description','supplier','similarity'] then
    raise exception 'Banco amount-only candidate identity, exact-cent/date evidence, or no-similarity contract is invalid: %', v_candidate;
  end if;

  select candidates, candidate_count into strict v_candidate, v_candidate_count
  from public.financial_reconciliation_automatic_single_base_candidates(
    'financial_documents_cgd_credit_card_amount_only', 1, 0, 1,
    '62000000-0000-0000-0000-000000000001'
  );
  if v_candidate_count <> 1
    or jsonb_array_length(v_candidate) <> 1
    or v_candidate->0->>'sourceType' <> 'import_cgd_cartao_credito'
    or v_candidate->0->>'sourceId' <> '63000000-0000-0000-0000-000000000001'
    or v_candidate->0->>'sourceDate' <> '2030-06-15'
    or (v_candidate->0->>'amount')::numeric <> -100.00
    or v_candidate#>>'{0,evidence,amount,baseAmountCents}' <> '10000'
    or v_candidate#>>'{0,evidence,amount,destinationAmountCents}' <> '-10000'
    or v_candidate#>>'{0,evidence,amount,signedDifferenceCents}' <> '0'
    or v_candidate#>>'{0,evidence,date,distanceDays}' <> '0'
    or v_candidate#>'{0,evidence}' ?| array['documentNumber','description','supplier','similarity'] then
    raise exception 'Visa amount-only candidate identity, exact-cent/date evidence, or no-similarity contract is invalid: %', v_candidate;
  end if;

  select candidates, candidate_count into strict v_candidate, v_candidate_count
  from public.financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(
    'financial_documents_cgd_bank_statement_amount_only', 1, 0, 1,
    array['60000000-0000-0000-0000-000000000020'::uuid]
  );
  if v_candidate_count <> 0 or v_candidate <> '[]'::jsonb then
    raise exception 'Banco amount-only admitted a pre-2026, null, locked, or wrong approved-field destination: %', v_candidate;
  end if;

  select candidates, candidate_count into strict v_candidate, v_candidate_count
  from public.financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(
    'financial_documents_cgd_credit_card_amount_only', 1, 0, 1,
    array['62000000-0000-0000-0000-000000000020'::uuid]
  );
  if v_candidate_count <> 0 or v_candidate <> '[]'::jsonb then
    raise exception 'Visa amount-only admitted a pre-2026, empty-source, locked, or wrong approved-field destination: %', v_candidate;
  end if;

  if exists (
    select 1 from public.financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(
      'financial_documents_cgd_bank_statement_amount_only', 1, 0.01, 1,
      array['60000000-0000-0000-0000-000000000001'::uuid]
    )
  ) or exists (
    select 1 from public.financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(
      'financial_documents_cgd_credit_card_amount_only', 1, 0.01, 1,
      array['62000000-0000-0000-0000-000000000001'::uuid]
    )
  ) then
    raise exception 'An amount-only adapter accepted a nonzero difference allowance.';
  end if;
end $$;

update public.financial_documents
set document_date = coalesce(document_date, date '2030-06-15'),
    amount = coalesce(amount, 0),
    payment = coalesce(payment, 'smoke:null-payment')
where id between '60000000-0000-0000-0000-000000000001'::uuid
             and '62000000-0000-0000-0000-000000000020'::uuid;

alter table public.financial_documents
  alter column document_date set not null,
  alter column amount set not null,
  alter column payment set not null;

-- amount-only inclusive date windows, signed integer cents, stable order, and identity independence
do $$
declare
  v_ids uuid[];
  v_candidate jsonb;
  v_count integer;
begin
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, payment, amount, fat
  ) values
    ('60000000-0000-0000-0000-000000000030', date '2031-06-15', 'BANK-WINDOW', 'window identity', 'window supplier', 'Banco', 110.00, 'S'),
    ('60000000-0000-0000-0000-000000000031', date '2031-09-01', 'BANK-90', 'boundary identity', 'boundary supplier', 'Banco', 111.00, 'S'),
    ('60000000-0000-0000-0000-000000000032', date '2031-12-01', 'BANK-AMOUNT-032', 'matching text', 'same supplier', 'Banco', 120.00, 'S'),
    ('62000000-0000-0000-0000-000000000030', date '2031-06-15', 'CARD-WINDOW', 'window identity', 'window supplier', 'Visa', 110.00, 'S'),
    ('62000000-0000-0000-0000-000000000031', date '2031-09-01', 'CARD-90', 'boundary identity', 'boundary supplier', 'Visa', 111.00, 'S'),
    ('62000000-0000-0000-0000-000000000032', date '2031-12-01', 'CARD-AMOUNT-032', 'matching text', 'same supplier', 'Visa', 120.00, 'S');

  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  ) values
    ('61000000-0000-0000-0000-000000000030', 'smoke-amount-only', 'bank-window-minus-2', date '2031-06-13', '', -110.00),
    ('61000000-0000-0000-0000-000000000031', 'smoke-amount-only', 'bank-window-minus-1', date '2031-06-14', '', -110.00),
    ('61000000-0000-0000-0000-000000000032', 'smoke-amount-only', 'bank-window-zero', date '2031-06-15', '', -110.00),
    ('61000000-0000-0000-0000-000000000033', 'smoke-amount-only', 'bank-window-plus-1', date '2031-06-16', '', -110.00),
    ('61000000-0000-0000-0000-000000000034', 'smoke-amount-only', 'bank-window-plus-2', date '2031-06-17', '', -110.00),
    ('61000000-0000-0000-0000-000000000035', 'smoke-amount-only', 'bank-window-90', date '2031-11-30', '', -111.00),
    ('61000000-0000-0000-0000-000000000036', 'smoke-amount-only', 'bank-window-91', date '2031-12-01', '', -111.00),
    ('61000000-0000-0000-0000-000000000040', 'smoke-amount-only', 'bank-exact-cents', date '2031-12-01', 'totally unrelated', -120.00),
    ('61000000-0000-0000-0000-000000000041', 'smoke-amount-only', 'bank-one-cent', date '2031-12-01', 'BANKAMOUNT032 matching text same supplier', -119.99),
    ('61000000-0000-0000-0000-000000000042', 'smoke-amount-only', 'bank-same-sign', date '2031-12-01', 'BANKAMOUNT032 matching text same supplier', 120.00);

  insert into public.import_cgd_cartao_credito (
    id, import_batch, row_key, data, descricao, debito, credito
  ) values
    ('63000000-0000-0000-0000-000000000030', 'smoke-amount-only', 'card-window-minus-2', date '2031-06-13', '', 110.00, null),
    ('63000000-0000-0000-0000-000000000031', 'smoke-amount-only', 'card-window-minus-1', date '2031-06-14', '', 110.00, null),
    ('63000000-0000-0000-0000-000000000032', 'smoke-amount-only', 'card-window-zero', date '2031-06-15', '', 110.00, null),
    ('63000000-0000-0000-0000-000000000033', 'smoke-amount-only', 'card-window-plus-1', date '2031-06-16', '', 110.00, null),
    ('63000000-0000-0000-0000-000000000034', 'smoke-amount-only', 'card-window-plus-2', date '2031-06-17', '', 110.00, null),
    ('63000000-0000-0000-0000-000000000035', 'smoke-amount-only', 'card-window-90', date '2031-11-30', '', 111.00, null),
    ('63000000-0000-0000-0000-000000000036', 'smoke-amount-only', 'card-window-91', date '2031-12-01', '', 111.00, null),
    ('63000000-0000-0000-0000-000000000040', 'smoke-amount-only', 'card-exact-cents', date '2031-12-01', 'totally unrelated', 120.00, null),
    ('63000000-0000-0000-0000-000000000041', 'smoke-amount-only', 'card-one-cent', date '2031-12-01', 'CARDAMOUNT032 matching text same supplier', 119.99, null),
    ('63000000-0000-0000-0000-000000000042', 'smoke-amount-only', 'card-same-sign', date '2031-12-01', 'CARDAMOUNT032 matching text same supplier', null, 120.00);

  select array_agg((item->>'sourceId')::uuid order by ordinal)
  into v_ids
  from public.financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(
    'financial_documents_cgd_bank_statement_amount_only', 1, 0, 1,
    array['60000000-0000-0000-0000-000000000030'::uuid]
  ) candidate
  cross join lateral jsonb_array_elements(candidate.candidates) with ordinality entry(item, ordinal);
  if v_ids is distinct from array[
    '61000000-0000-0000-0000-000000000031'::uuid,
    '61000000-0000-0000-0000-000000000032'::uuid,
    '61000000-0000-0000-0000-000000000033'::uuid
  ] then
    raise exception 'Banco amount-only default window did not include -1/0/+1 and exclude -2/+2 in stable order: %', v_ids;
  end if;

  select candidates into strict v_candidate
  from public.financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(
    'financial_documents_cgd_bank_statement_amount_only', 1, 0, 0,
    array['60000000-0000-0000-0000-000000000030'::uuid]
  );
  if jsonb_array_length(v_candidate) <> 1
    or v_candidate->0->>'sourceId' <> '61000000-0000-0000-0000-000000000032' then
    raise exception 'Banco amount-only zero-day window admitted a non-same-day row.';
  end if;

  select candidates into strict v_candidate
  from public.financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(
    'financial_documents_cgd_bank_statement_amount_only', 1, 0, 90,
    array['60000000-0000-0000-0000-000000000031'::uuid]
  );
  if jsonb_array_length(v_candidate) <> 1
    or v_candidate->0->>'sourceId' <> '61000000-0000-0000-0000-000000000035'
    or v_candidate#>>'{0,evidence,date,distanceDays}' <> '90' then
    raise exception 'Banco amount-only 90-day boundary was not inclusive or admitted day 91.';
  end if;

  select candidates, candidate_count into strict v_candidate, v_count
  from public.financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(
    'financial_documents_cgd_bank_statement_amount_only', 1, 0, 0,
    array['60000000-0000-0000-0000-000000000032'::uuid]
  );
  if v_count <> 1 or v_candidate->0->>'sourceId' <> '61000000-0000-0000-0000-000000000040'
    or v_candidate#>>'{0,evidence,amount,signedDifferenceCents}' <> '0' then
    raise exception 'Banco amount-only allowed a one-cent mismatch, same sign, or identity override: %', v_candidate;
  end if;

  select array_agg((item->>'sourceId')::uuid order by ordinal)
  into v_ids
  from public.financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(
    'financial_documents_cgd_credit_card_amount_only', 1, 0, 1,
    array['62000000-0000-0000-0000-000000000030'::uuid]
  ) candidate
  cross join lateral jsonb_array_elements(candidate.candidates) with ordinality entry(item, ordinal);
  if v_ids is distinct from array[
    '63000000-0000-0000-0000-000000000031'::uuid,
    '63000000-0000-0000-0000-000000000032'::uuid,
    '63000000-0000-0000-0000-000000000033'::uuid
  ] then
    raise exception 'Visa amount-only default window did not include -1/0/+1 and exclude -2/+2 in stable order: %', v_ids;
  end if;

  select candidates into strict v_candidate
  from public.financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(
    'financial_documents_cgd_credit_card_amount_only', 1, 0, 0,
    array['62000000-0000-0000-0000-000000000030'::uuid]
  );
  if jsonb_array_length(v_candidate) <> 1
    or v_candidate->0->>'sourceId' <> '63000000-0000-0000-0000-000000000032' then
    raise exception 'Visa amount-only zero-day window admitted a non-same-day row.';
  end if;

  select candidates into strict v_candidate
  from public.financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(
    'financial_documents_cgd_credit_card_amount_only', 1, 0, 90,
    array['62000000-0000-0000-0000-000000000031'::uuid]
  );
  if jsonb_array_length(v_candidate) <> 1
    or v_candidate->0->>'sourceId' <> '63000000-0000-0000-0000-000000000035'
    or v_candidate#>>'{0,evidence,date,distanceDays}' <> '90' then
    raise exception 'Visa amount-only 90-day boundary was not inclusive or admitted day 91.';
  end if;

  select candidates, candidate_count into strict v_candidate, v_count
  from public.financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(
    'financial_documents_cgd_credit_card_amount_only', 1, 0, 0,
    array['62000000-0000-0000-0000-000000000032'::uuid]
  );
  if v_count <> 1 or v_candidate->0->>'sourceId' <> '63000000-0000-0000-0000-000000000040'
    or v_candidate#>>'{0,evidence,amount,signedDifferenceCents}' <> '0' then
    raise exception 'Visa amount-only allowed a one-cent mismatch, same sign, or identity override: %', v_candidate;
  end if;
end $$;

update public.financial_documents
set payment = 'smoke:amount-only-candidate-covered'
where id between '60000000-0000-0000-0000-000000000001'::uuid
             and '62000000-0000-0000-0000-000000000032'::uuid;

-- amount-only skipped, duplicate, candidate-limit, cross-base, and one-row proposal lifecycle
do $$
declare
  v_run jsonb;
  v_run_id uuid;
  v_candidates jsonb;
  v_candidate_count integer;
  v_base_ids uuid[];
begin
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, payment, amount, fat
  ) values
    ('70000000-0000-0000-0000-000000000001', date '2135-01-01', 'BANK-NONE', '', '', 'Banco', 200.00, 'S'),
    ('70000000-0000-0000-0000-000000000002', date '2135-01-01', 'BANK-ONE', 'does not match', 'does not match', 'Banco', 201.00, 'S'),
    ('70000000-0000-0000-0000-000000000003', date '2135-01-01', 'BANK-TWO', '', '', 'Banco', 202.00, 'S'),
    ('70000000-0000-0000-0000-000000000004', date '2135-01-01', 'BANK-LIMIT', '', '', 'Banco', 203.00, 'S'),
    ('70000000-0000-0000-0000-000000000005', date '2135-01-01', 'BANK-CROSS-A', '', '', 'Banco', 204.00, 'S'),
    ('70000000-0000-0000-0000-000000000006', date '2135-01-01', 'BANK-CROSS-B', '', '', 'Banco', 204.00, 'S'),
    ('70000000-0000-0000-0000-000000000007', date '2135-01-01', 'BANK-SUM', 'SUM-IDENTITY', 'SUM-SUPPLIER', 'Banco', 205.00, 'S'),
    ('72000000-0000-0000-0000-000000000001', date '2135-01-01', 'CARD-NONE', '', '', 'Visa', 200.00, 'S'),
    ('72000000-0000-0000-0000-000000000002', date '2135-01-01', 'CARD-ONE', 'does not match', 'does not match', 'Visa', 201.00, 'S'),
    ('72000000-0000-0000-0000-000000000003', date '2135-01-01', 'CARD-TWO', '', '', 'Visa', 202.00, 'S'),
    ('72000000-0000-0000-0000-000000000004', date '2135-01-01', 'CARD-LIMIT', '', '', 'Visa', 203.00, 'S'),
    ('72000000-0000-0000-0000-000000000005', date '2135-01-01', 'CARD-CROSS-A', '', '', 'Visa', 204.00, 'S'),
    ('72000000-0000-0000-0000-000000000006', date '2135-01-01', 'CARD-CROSS-B', '', '', 'Visa', 204.00, 'S'),
    ('72000000-0000-0000-0000-000000000007', date '2135-01-01', 'CARD-SUM', 'SUM-IDENTITY', 'SUM-SUPPLIER', 'Visa', 205.00, 'S');

  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  ) values
    ('71000000-0000-0000-0000-000000000002', 'smoke-amount-only-lifecycle', 'bank-one', date '2135-01-01', 'unrelated', -201.00),
    ('71000000-0000-0000-0000-000000000003', 'smoke-amount-only-lifecycle', 'bank-two-a', date '2135-01-01', '', -202.00),
    ('71000000-0000-0000-0000-000000000004', 'smoke-amount-only-lifecycle', 'bank-two-b', date '2135-01-01', '', -202.00),
    ('71000000-0000-0000-0000-000000000005', 'smoke-amount-only-lifecycle', 'bank-cross', date '2135-01-01', '', -204.00),
    ('71000000-0000-0000-0000-000000000006', 'smoke-amount-only-lifecycle', 'bank-sum-a', date '2135-01-01', 'SUMIDENTITY SUMSUPPLIER', -100.00),
    ('71000000-0000-0000-0000-000000000007', 'smoke-amount-only-lifecycle', 'bank-sum-b', date '2135-01-01', 'SUMIDENTITY SUMSUPPLIER', -105.00);
  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  )
  select
    ('71000000-0000-0000-0000-' || lpad((100 + series)::text, 12, '0'))::uuid,
    'smoke-amount-only-lifecycle', 'bank-limit-' || series,
    date '2135-01-01', '', -203.00
  from generate_series(1, 13) series;

  insert into public.import_cgd_cartao_credito (
    id, import_batch, row_key, data, descricao, debito
  ) values
    ('73000000-0000-0000-0000-000000000002', 'smoke-amount-only-lifecycle', 'card-one', date '2135-01-01', 'unrelated', 201.00),
    ('73000000-0000-0000-0000-000000000003', 'smoke-amount-only-lifecycle', 'card-two-a', date '2135-01-01', '', 202.00),
    ('73000000-0000-0000-0000-000000000004', 'smoke-amount-only-lifecycle', 'card-two-b', date '2135-01-01', '', 202.00),
    ('73000000-0000-0000-0000-000000000005', 'smoke-amount-only-lifecycle', 'card-cross', date '2135-01-01', '', 204.00),
    ('73000000-0000-0000-0000-000000000006', 'smoke-amount-only-lifecycle', 'card-sum-a', date '2135-01-01', 'SUMIDENTITY SUMSUPPLIER', 100.00),
    ('73000000-0000-0000-0000-000000000007', 'smoke-amount-only-lifecycle', 'card-sum-b', date '2135-01-01', 'SUMIDENTITY SUMSUPPLIER', 105.00);
  insert into public.import_cgd_cartao_credito (
    id, import_batch, row_key, data, descricao, debito
  )
  select
    ('73000000-0000-0000-0000-' || lpad((100 + series)::text, 12, '0'))::uuid,
    'smoke-amount-only-lifecycle', 'card-limit-' || series,
    date '2135-01-01', '', 203.00
  from generate_series(1, 13) series;

  select candidates, candidate_count into strict v_candidates, v_candidate_count
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_bank_statement_amount_only', 1, 0, 1,
    array['70000000-0000-0000-0000-000000000004'::uuid]
  );
  if v_candidate_count <> 13 or jsonb_array_length(v_candidates) <> 12 then
    raise exception 'Banco amount-only did not retain unbounded count with twelve bounded evidence rows: %, %', v_candidate_count, v_candidates;
  end if;

  select candidates, candidate_count into strict v_candidates, v_candidate_count
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_credit_card_amount_only', 1, 0, 1,
    array['72000000-0000-0000-0000-000000000004'::uuid]
  );
  if v_candidate_count <> 13 or jsonb_array_length(v_candidates) <> 12 then
    raise exception 'Visa amount-only did not retain unbounded count with twelve bounded evidence rows: %, %', v_candidate_count, v_candidates;
  end if;

  update public.financial_reconciliation_automatic_rule_configs
  set enabled = true,
      allow_manual_execution = true,
      include_in_scheduled_batch = false,
      difference_allowed = 0,
      max_difference_days = 1
  where rule_key = 'financial_documents_cgd_bank_statement_amount_only';

  if public.financial_reconciliation_automatic_base_count(
      'financial_documents_cgd_bank_statement_amount_only', 1
    ) <> 7 then
    raise exception 'Banco amount-only base count did not use exact managed eligibility.';
  end if;
  select array_agg(page.id order by page.document_date, page.id) into v_base_ids
  from public.financial_reconciliation_automatic_base_page(
    'financial_documents_cgd_bank_statement_amount_only', 1, null, null, 25
  ) page;
  if v_base_ids is distinct from array[
    '70000000-0000-0000-0000-000000000001'::uuid,
    '70000000-0000-0000-0000-000000000002'::uuid,
    '70000000-0000-0000-0000-000000000003'::uuid,
    '70000000-0000-0000-0000-000000000004'::uuid,
    '70000000-0000-0000-0000-000000000005'::uuid,
    '70000000-0000-0000-0000-000000000006'::uuid,
    '70000000-0000-0000-0000-000000000007'::uuid
  ] then
    raise exception 'Banco amount-only base page did not preserve (document_date,id) ordering: %', v_base_ids;
  end if;

  v_run := public.create_financial_reconciliation_automatic_analysis(
    array['financial_documents_cgd_bank_statement_amount_only'], 'manual_rule',
    'smoke:bank-amount-only-lifecycle', '74000000-0000-0000-0000-000000000001'
  );
  v_run_id := (v_run->>'runId')::uuid;
  if v_run->>'status' <> 'ready'
    or v_run->>'analysisComplete' <> 'true'
    or v_run#>>'{counts,bases}' <> '7'
    or v_run#>>'{counts,proposed}' <> '1'
    or v_run#>>'{counts,ambiguous}' <> '4'
    or v_run#>>'{counts,skipped}' <> '2' then
    raise exception 'Banco amount-only lifecycle counts or terminal state are invalid: %', v_run;
  end if;
  if not exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = v_run_id
        and proposal.base_source_id = '70000000-0000-0000-0000-000000000001'
        and proposal.status = 'skipped' and proposal.reason = 'no_qualifying_combination'
        and proposal.items = '[]'::jsonb
    ) or not exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = v_run_id
        and proposal.base_source_id = '70000000-0000-0000-0000-000000000002'
        and proposal.status = 'proposed' and jsonb_array_length(proposal.items) = 1
        and proposal.calculated_difference = 0
        and not (proposal.evidence->0 ?| array['documentNumber','description','supplier','similarity'])
    ) or not exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = v_run_id
        and proposal.base_source_id = '70000000-0000-0000-0000-000000000003'
        and proposal.status = 'ambiguous' and proposal.reason = 'multiple_combinations'
        and jsonb_array_length(proposal.candidate_groups) = 2
    ) or not exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = v_run_id
        and proposal.base_source_id = '70000000-0000-0000-0000-000000000004'
        and proposal.status = 'ambiguous' and proposal.reason = 'candidate_limit'
        and jsonb_array_length(proposal.candidate_groups) = 12
    ) or (select count(*) from public.financial_reconciliation_automatic_proposals proposal
          where proposal.run_id = v_run_id
            and proposal.base_source_id in (
              '70000000-0000-0000-0000-000000000005',
              '70000000-0000-0000-0000-000000000006'
            ) and proposal.status = 'ambiguous' and proposal.reason = 'cross_base_overlap') <> 2
    or not exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = v_run_id
        and proposal.base_source_id = '70000000-0000-0000-0000-000000000007'
        and proposal.status = 'skipped' and proposal.reason = 'no_qualifying_combination'
        and proposal.candidate_groups = '[]'::jsonb
    ) then
    raise exception 'Banco amount-only proposal, ambiguity, overlap, candidate-limit, or two-row-sum behavior is invalid.';
  end if;

  update public.financial_reconciliation_automatic_rule_configs
  set enabled = false, allow_manual_execution = false
  where rule_key = 'financial_documents_cgd_bank_statement_amount_only';
  update public.financial_documents
  set payment = 'smoke:bank-amount-only-lifecycle-covered'
  where id between '70000000-0000-0000-0000-000000000001'::uuid
               and '70000000-0000-0000-0000-000000000007'::uuid;

  update public.financial_reconciliation_automatic_rule_configs
  set enabled = true,
      allow_manual_execution = true,
      include_in_scheduled_batch = false,
      difference_allowed = 0,
      max_difference_days = 1
  where rule_key = 'financial_documents_cgd_credit_card_amount_only';

  if public.financial_reconciliation_automatic_base_count(
      'financial_documents_cgd_credit_card_amount_only', 1
    ) <> 7 then
    raise exception 'Visa amount-only base count did not use exact managed eligibility.';
  end if;
  select array_agg(page.id order by page.document_date, page.id) into v_base_ids
  from public.financial_reconciliation_automatic_base_page(
    'financial_documents_cgd_credit_card_amount_only', 1, null, null, 25
  ) page;
  if v_base_ids is distinct from array[
    '72000000-0000-0000-0000-000000000001'::uuid,
    '72000000-0000-0000-0000-000000000002'::uuid,
    '72000000-0000-0000-0000-000000000003'::uuid,
    '72000000-0000-0000-0000-000000000004'::uuid,
    '72000000-0000-0000-0000-000000000005'::uuid,
    '72000000-0000-0000-0000-000000000006'::uuid,
    '72000000-0000-0000-0000-000000000007'::uuid
  ] then
    raise exception 'Visa amount-only base page did not preserve (document_date,id) ordering: %', v_base_ids;
  end if;

  v_run := public.create_financial_reconciliation_automatic_analysis(
    array['financial_documents_cgd_credit_card_amount_only'], 'manual_rule',
    'smoke:card-amount-only-lifecycle', '74000000-0000-0000-0000-000000000002'
  );
  v_run_id := (v_run->>'runId')::uuid;
  if v_run->>'status' <> 'ready'
    or v_run->>'analysisComplete' <> 'true'
    or v_run#>>'{counts,bases}' <> '7'
    or v_run#>>'{counts,proposed}' <> '1'
    or v_run#>>'{counts,ambiguous}' <> '4'
    or v_run#>>'{counts,skipped}' <> '2' then
    raise exception 'Visa amount-only lifecycle counts or terminal state are invalid: %', v_run;
  end if;
  if not exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = v_run_id
        and proposal.base_source_id = '72000000-0000-0000-0000-000000000001'
        and proposal.status = 'skipped' and proposal.reason = 'no_qualifying_combination'
        and proposal.items = '[]'::jsonb
    ) or not exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = v_run_id
        and proposal.base_source_id = '72000000-0000-0000-0000-000000000002'
        and proposal.status = 'proposed' and jsonb_array_length(proposal.items) = 1
        and proposal.calculated_difference = 0
        and not (proposal.evidence->0 ?| array['documentNumber','description','supplier','similarity'])
    ) or not exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = v_run_id
        and proposal.base_source_id = '72000000-0000-0000-0000-000000000003'
        and proposal.status = 'ambiguous' and proposal.reason = 'multiple_combinations'
        and jsonb_array_length(proposal.candidate_groups) = 2
    ) or not exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = v_run_id
        and proposal.base_source_id = '72000000-0000-0000-0000-000000000004'
        and proposal.status = 'ambiguous' and proposal.reason = 'candidate_limit'
        and jsonb_array_length(proposal.candidate_groups) = 12
    ) or (select count(*) from public.financial_reconciliation_automatic_proposals proposal
          where proposal.run_id = v_run_id
            and proposal.base_source_id in (
              '72000000-0000-0000-0000-000000000005',
              '72000000-0000-0000-0000-000000000006'
            ) and proposal.status = 'ambiguous' and proposal.reason = 'cross_base_overlap') <> 2
    or not exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = v_run_id
        and proposal.base_source_id = '72000000-0000-0000-0000-000000000007'
        and proposal.status = 'skipped' and proposal.reason = 'no_qualifying_combination'
        and proposal.candidate_groups = '[]'::jsonb
    ) then
    raise exception 'Visa amount-only proposal, ambiguity, overlap, candidate-limit, or two-row-sum behavior is invalid.';
  end if;

  update public.financial_reconciliation_automatic_rule_configs
  set enabled = false, allow_manual_execution = false
  where rule_key = 'financial_documents_cgd_credit_card_amount_only';
  update public.financial_documents
  set payment = 'smoke:card-amount-only-lifecycle-covered'
  where id between '72000000-0000-0000-0000-000000000001'::uuid
               and '72000000-0000-0000-0000-000000000007'::uuid;
end $$;

-- credit-card projection INSERT UPDATE ID-change DELETE and data_valor isolation
do $$
declare
  v_projected_date date;
begin
  insert into public.import_cgd_cartao_credito (
    id, import_batch, row_key, data, data_valor, descricao, credito
  ) values (
    '46000000-0000-0000-0000-000000000901', 'smoke-credit-card',
    'credit-card-sync-901', date '2027-08-10', date '2027-08-11',
    'Credit card projection insert', 12.34
  );
  if not exists (
    select 1 from public.financial_reconciliation_cgd_credit_card_match_search
    where source_id = '46000000-0000-0000-0000-000000000901'
      and source_date = date '2027-08-10'
      and amount = 12.34
      and description = 'Credit card projection insert'
  ) then
    raise exception 'Credit-card projection INSERT did not synchronize data, valor, and descricao.';
  end if;

  update public.import_cgd_cartao_credito
  set data = date '2027-08-12',
      data_valor = date '2027-08-13',
      descricao = 'Credit card projection update',
      credito = 23.45
  where id = '46000000-0000-0000-0000-000000000901';
  if not exists (
    select 1 from public.financial_reconciliation_cgd_credit_card_match_search
    where source_id = '46000000-0000-0000-0000-000000000901'
      and source_date = date '2027-08-12'
      and amount = 23.45
      and description = 'Credit card projection update'
  ) then
    raise exception 'Credit-card projection UPDATE did not synchronize data, valor, and descricao.';
  end if;

  select source_date into strict v_projected_date
  from public.financial_reconciliation_cgd_credit_card_match_search
  where source_id = '46000000-0000-0000-0000-000000000901';
  update public.import_cgd_cartao_credito
  set data_valor = date '2027-09-30'
  where id = '46000000-0000-0000-0000-000000000901';
  if (select source_date from public.financial_reconciliation_cgd_credit_card_match_search
      where source_id = '46000000-0000-0000-0000-000000000901') is distinct from v_projected_date then
    raise exception 'A data_valor-only change altered the projected reconciliation date.';
  end if;

  update public.import_cgd_cartao_credito
  set id = '46000000-0000-0000-0000-000000000902'
  where id = '46000000-0000-0000-0000-000000000901';
  if exists (
      select 1 from public.financial_reconciliation_cgd_credit_card_match_search
      where source_id = '46000000-0000-0000-0000-000000000901'
    ) or not exists (
      select 1 from public.financial_reconciliation_cgd_credit_card_match_search
      where source_id = '46000000-0000-0000-0000-000000000902'
        and source_date = date '2027-08-12'
        and amount = 23.45
        and description = 'Credit card projection update'
    ) then
    raise exception 'Credit-card projection did not synchronize a source ID change.';
  end if;

  delete from public.import_cgd_cartao_credito
  where id = '46000000-0000-0000-0000-000000000902';
  if exists (
    select 1 from public.financial_reconciliation_cgd_credit_card_match_search
    where source_id = '46000000-0000-0000-0000-000000000902'
  ) then
    raise exception 'Credit-card projection did not synchronize DELETE.';
  end if;
end $$;

-- credit-card exact Visa eligibility and exclusions
do $$
declare
  v_reconciliation_id uuid;
  v_base_ids uuid[];
begin
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, payment, amount, fat
  ) values
    ('47000000-0000-0000-0000-000000000001', date '2027-09-01', 'CC-EXACT-0001', '', '', 'Visa', 101.00, 'S'),
    ('47000000-0000-0000-0000-000000000002', date '2027-09-02', 'CC-UPPER-0002', '', '', 'VISA', 102.00, 'S'),
    ('47000000-0000-0000-0000-000000000003', date '2027-09-03', 'CC-LOWER-0003', '', '', 'visa', 103.00, 'S'),
    ('47000000-0000-0000-0000-000000000004', date '2027-09-04', 'CC-PADDED-0004', '', '', ' Visa ', 104.00, 'S'),
    ('47000000-0000-0000-0000-000000000005', date '2027-09-05', 'CC-NULL-0005', '', '', null, 105.00, 'S'),
    ('47000000-0000-0000-0000-000000000006', date '2025-12-31', 'CC-PRE-0006', '', '', 'Visa', 106.00, 'S'),
    ('47000000-0000-0000-0000-000000000007', date '2027-09-07', 'CC-LOCKED-0007', '', '', 'Visa', 107.00, 'S');

  insert into public.import_cgd_cartao_credito (
    id, import_batch, row_key, data, descricao, debito
  ) values
    ('48000000-0000-0000-0000-000000000001', 'smoke-credit-card', 'cc-exact-0001', date '2027-09-01', 'Payment CCEXACT0001', 101.00),
    ('48000000-0000-0000-0000-000000000002', 'smoke-credit-card', 'cc-upper-0002', date '2027-09-02', 'Payment CCUPPER0002', 102.00),
    ('48000000-0000-0000-0000-000000000003', 'smoke-credit-card', 'cc-lower-0003', date '2027-09-03', 'Payment CCLOWER0003', 103.00),
    ('48000000-0000-0000-0000-000000000004', 'smoke-credit-card', 'cc-padded-0004', date '2027-09-04', 'Payment CCPADDED0004', 104.00),
    ('48000000-0000-0000-0000-000000000005', 'smoke-credit-card', 'cc-null-0005', date '2027-09-05', 'Payment CCNULL0005', 105.00),
    ('48000000-0000-0000-0000-000000000006', 'smoke-credit-card', 'cc-pre-0006', date '2025-12-31', 'Payment CCPRE0006', 106.00),
    ('48000000-0000-0000-0000-000000000007', 'smoke-credit-card', 'cc-locked-0007', date '2027-09-07', 'Payment CCLOCKED0007', 107.00);

  insert into public.financial_reconciliations (
    status, base_source_type, matching_source_types, created_by
  ) values (
    'started', 'financial_documents', '["import_cgd_cartao_credito"]'::jsonb,
    'smoke:credit-card-lock'
  ) returning id into v_reconciliation_id;
  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  ) values (
    v_reconciliation_id, 'financial_documents',
    '47000000-0000-0000-0000-000000000007', 107.00,
    'smoke:credit-card-lock'
  );

  select coalesce(array_agg(candidate.base_source_id order by candidate.base_source_id), '{}'::uuid[])
  into v_base_ids
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_credit_card', 1, 0, 10,
    array[
      '47000000-0000-0000-0000-000000000001'::uuid,
      '47000000-0000-0000-0000-000000000002'::uuid,
      '47000000-0000-0000-0000-000000000003'::uuid,
      '47000000-0000-0000-0000-000000000004'::uuid,
      '47000000-0000-0000-0000-000000000005'::uuid,
      '47000000-0000-0000-0000-000000000006'::uuid,
      '47000000-0000-0000-0000-000000000007'::uuid
    ]
  ) candidate;
  if v_base_ids is distinct from array['47000000-0000-0000-0000-000000000001'::uuid] then
    raise exception 'Exact Visa eligibility returned uppercase, lowercase, padded, null, pre-2026, or locked bases: %', v_base_ids;
  end if;
end $$;

-- credit-card dates exactly 10 and 11 days apart
do $$
declare
  v_candidates jsonb;
begin
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, payment, amount, fat
  ) values (
    '47000000-0000-0000-0000-000000000010', date '2027-10-01',
    'CC-DAY-0010', '', '', 'Visa', 110.00, 'S'
  );
  insert into public.import_cgd_cartao_credito (
    id, import_batch, row_key, data, descricao, debito
  ) values
    ('48000000-0000-0000-0000-000000000010', 'smoke-credit-card', 'cc-day-10', date '2027-10-11', 'Payment CCDAY0010', 110.00),
    ('48000000-0000-0000-0000-000000000011', 'smoke-credit-card', 'cc-day-11', date '2027-10-12', 'Payment CCDAY0010', 110.00);

  select candidates into strict v_candidates
  from public.financial_reconciliation_automatic_credit_card_candidates_for_base_ids(
    'financial_documents_cgd_credit_card', 1, 0, 10,
    array['47000000-0000-0000-0000-000000000010'::uuid]
  );
  if jsonb_array_length(v_candidates) <> 1
    or v_candidates->0->>'sourceId' <> '48000000-0000-0000-0000-000000000010' then
    raise exception 'Credit-card date boundary did not include day 10 and exclude day 11.';
  end if;
end $$;

-- credit-card symmetric compact document-number containment with four-character minimum
do $$
declare
  v_forward jsonb;
  v_reverse jsonb;
  v_short_count integer;
begin
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, payment, amount, fat
  ) values
    ('47000000-0000-0000-0000-000000000020', date '2027-11-01', 'AB-12', '', '', 'Visa', 120.00, 'S'),
    ('47000000-0000-0000-0000-000000000021', date '2027-11-20', 'LONG-1234', '', '', 'Visa', 121.00, 'S'),
    ('47000000-0000-0000-0000-000000000022', date '2027-12-10', 'A-12', '', '', 'Visa', 122.00, 'S');
  insert into public.import_cgd_cartao_credito (
    id, import_batch, row_key, data, descricao, debito
  ) values
    ('48000000-0000-0000-0000-000000000020', 'smoke-credit-card', 'cc-containment-forward', date '2027-11-01', 'Payment AB12 reference', 120.00),
    ('48000000-0000-0000-0000-000000000021', 'smoke-credit-card', 'cc-containment-reverse', date '2027-11-20', 'LONG', 121.00),
    ('48000000-0000-0000-0000-000000000022', 'smoke-credit-card', 'cc-containment-short', date '2027-12-10', 'A12', 122.00);

  select candidates into strict v_forward
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_credit_card', 1, 0, 10,
    array['47000000-0000-0000-0000-000000000020'::uuid]
  );
  select candidates into strict v_reverse
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_credit_card', 1, 0, 10,
    array['47000000-0000-0000-0000-000000000021'::uuid]
  );
  select candidate_count into strict v_short_count
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_credit_card', 1, 0, 10,
    array['47000000-0000-0000-0000-000000000022'::uuid]
  );
  if jsonb_array_length(v_forward) <> 1
    or (v_forward#>>'{0,evidence,documentNumber,matched}')::boolean is not true
    or jsonb_array_length(v_reverse) <> 1
    or (v_reverse#>>'{0,evidence,documentNumber,matched}')::boolean is not true
    or v_short_count <> 0 then
    raise exception 'Symmetric compact containment or its four-character invoice minimum changed.';
  end if;
end $$;

-- credit-card description score immediately below and at 0.55
-- credit-card supplier word score immediately below and at 0.60
-- credit-card independent identity branches
do $$
declare
  v_description_at text := 'abcdefghijk';
  v_description_at_score real;
  v_description_below text;
  v_description_below_score real;
  v_supplier_at text;
  v_supplier_at_score real;
  v_supplier_below text;
  v_supplier_below_score real;
  v_document_candidates jsonb;
  v_description_candidates jsonb;
  v_supplier_candidates jsonb;
  v_below_count integer;
begin
  v_description_at_score := public.financial_reconciliation_extension_similarity(
    'abcdefghijklmnopqr', v_description_at
  );
  select candidate, score into v_description_below, v_description_below_score
  from (
    select left('abcdefghijklmnopqr', prefix_length) as candidate,
           public.financial_reconciliation_extension_similarity(
             'abcdefghijklmnopqr', left('abcdefghijklmnopqr', prefix_length)
           ) as score
    from generate_series(1, 17) prefix_length
  ) scores
  where score < 0.55
  order by 0.55 - score, candidate
  limit 1;
  if abs(v_description_at_score - 0.55) >= 0.000001 then
    raise exception 'Credit-card description threshold fixture did not measure exactly 0.55: %', v_description_at_score;
  end if;
  if v_description_below is null or v_description_below_score >= 0.55
    or 0.55 - v_description_below_score > 0.051 then
    raise exception 'Credit-card description below fixture was not boundary-adjacent: %', v_description_below_score;
  end if;

  select candidate, score into v_supplier_at, v_supplier_at_score
  from (
    select candidate,
           public.financial_reconciliation_extension_word_similarity('abcdefg', candidate) as score
    from (
      select left('abcdefg', prefix_length) || repeat('z', suffix_length) as candidate
      from generate_series(1, 7) prefix_length
      cross join generate_series(1, 24) suffix_length
    ) corpus
  ) scores
  where abs(score - 0.60) < 0.000001
  order by candidate
  limit 1;
  select candidate, score into v_supplier_below, v_supplier_below_score
  from (
    select candidate,
           public.financial_reconciliation_extension_word_similarity('abcdefg', candidate) as score
    from (
      select left('abcdefg', prefix_length) || repeat('z', suffix_length) as candidate
      from generate_series(1, 7) prefix_length
      cross join generate_series(1, 24) suffix_length
    ) corpus
  ) scores
  where score < 0.60
  order by 0.60 - score, candidate
  limit 1;
  if v_supplier_at is null or abs(v_supplier_at_score - 0.60) >= 0.000001 then
    raise exception 'Credit-card supplier threshold fixture did not measure exactly 0.60.';
  end if;
  if v_supplier_below is null or v_supplier_below_score >= 0.60
    or 0.60 - v_supplier_below_score > 0.051 then
    raise exception 'Credit-card supplier below fixture was not boundary-adjacent: %', v_supplier_below_score;
  end if;

  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, payment, amount, fat
  ) values
    ('47000000-0000-0000-0000-000000000030', date '2028-01-01', 'ONLY-DOC-0030', '', '', 'Visa', 130.00, 'S'),
    ('47000000-0000-0000-0000-000000000031', date '2028-01-20', '', 'abcdefghijklmnopqr', '', 'Visa', 131.00, 'S'),
    ('47000000-0000-0000-0000-000000000032', date '2028-02-10', '', 'abcdefghijklmnopqr', '', 'Visa', 132.00, 'S'),
    ('47000000-0000-0000-0000-000000000033', date '2028-03-01', '', '', 'abcdefg', 'Visa', 133.00, 'S'),
    ('47000000-0000-0000-0000-000000000034', date '2028-03-20', '', '', 'abcdefg', 'Visa', 134.00, 'S');
  insert into public.import_cgd_cartao_credito (
    id, import_batch, row_key, data, descricao, debito
  ) values
    ('48000000-0000-0000-0000-000000000030', 'smoke-credit-card', 'cc-only-document', date '2028-01-01', 'ONLYDOC0030', 130.00),
    ('48000000-0000-0000-0000-000000000031', 'smoke-credit-card', 'cc-description-at', date '2028-01-20', v_description_at, 131.00),
    ('48000000-0000-0000-0000-000000000032', 'smoke-credit-card', 'cc-description-below', date '2028-02-10', v_description_below, 132.00),
    ('48000000-0000-0000-0000-000000000033', 'smoke-credit-card', 'cc-supplier-at', date '2028-03-01', v_supplier_at, 133.00),
    ('48000000-0000-0000-0000-000000000034', 'smoke-credit-card', 'cc-supplier-below', date '2028-03-20', v_supplier_below, 134.00);

  select candidates into strict v_document_candidates
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_credit_card', 1, 0, 10,
    array['47000000-0000-0000-0000-000000000030'::uuid]
  );
  select candidates into strict v_description_candidates
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_credit_card', 1, 0, 10,
    array['47000000-0000-0000-0000-000000000031'::uuid]
  );
  select candidate_count into strict v_below_count
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_credit_card', 1, 0, 10,
    array['47000000-0000-0000-0000-000000000032'::uuid]
  );
  if v_below_count <> 0 then
    raise exception 'Credit-card description score immediately below 0.55 qualified.';
  end if;
  select candidates into strict v_supplier_candidates
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_credit_card', 1, 0, 10,
    array['47000000-0000-0000-0000-000000000033'::uuid]
  );
  select candidate_count into strict v_below_count
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_credit_card', 1, 0, 10,
    array['47000000-0000-0000-0000-000000000034'::uuid]
  );
  if v_below_count <> 0 then
    raise exception 'Credit-card supplier score immediately below 0.60 qualified.';
  end if;

  if jsonb_array_length(v_document_candidates) <> 1
    or (v_document_candidates#>>'{0,evidence,documentNumber,matched}')::boolean is not true
    or (v_document_candidates#>>'{0,evidence,description,matched}')::boolean is not false
    or (v_document_candidates#>>'{0,evidence,supplier,matched}')::boolean is not false
    or jsonb_array_length(v_description_candidates) <> 1
    or (v_description_candidates#>>'{0,evidence,documentNumber,matched}')::boolean is not false
    or (v_description_candidates#>>'{0,evidence,description,matched}')::boolean is not true
    or (v_description_candidates#>>'{0,evidence,supplier,matched}')::boolean is not false
    or abs((v_description_candidates#>>'{0,evidence,description,score}')::real - 0.55) >= 0.000001
    or jsonb_array_length(v_supplier_candidates) <> 1
    or (v_supplier_candidates#>>'{0,evidence,documentNumber,matched}')::boolean is not false
    or (v_supplier_candidates#>>'{0,evidence,description,matched}')::boolean is not false
    or (v_supplier_candidates#>>'{0,evidence,supplier,matched}')::boolean is not true
    or abs((v_supplier_candidates#>>'{0,evidence,supplier,score}')::real - 0.60) >= 0.000001 then
    raise exception 'Credit-card identity branches did not qualify independently at exact thresholds.';
  end if;
end $$;

do $$
begin
  if not exists (
    select 1 from public.financial_reconciliation_automatic_rule_definitions
    where rule_key = 'financial_documents_cgd_bank_statement' and version = 1
  ) or not exists (
    select 1 from public.financial_reconciliation_automatic_rule_definitions
    where rule_key = 'financial_documents_cgd_bank_statement'
      and version = 2
      and definition#>>'{baseEligibility,payment,value}' = 'Banco'
      and (definition#>>'{baseEligibility,payment,caseSensitive}')::boolean
      and not (definition#>>'{baseEligibility,payment,trim}')::boolean
  ) then
    raise exception 'Managed Banco rule versions are invalid.';
  end if;

  if not exists (
    select 1 from public.financial_reconciliation_automatic_rule_configs
    where rule_key = 'financial_documents_cgd_bank_statement'
      and rule_version = 2
      and enabled
      and allow_manual_execution
      and not include_in_scheduled_batch
      and difference_allowed = 4.56
      and max_difference_days = 11
      and priority = 1
      and updated_by = 'smoke:banco-v2'
  ) then
    raise exception 'Version 2 migration changed administrator configuration.';
  end if;

  update public.financial_reconciliation_automatic_rule_configs
  set max_difference_days = 90
  where rule_key = 'financial_documents_cgd_bank_statement';
  begin
    update public.financial_reconciliation_automatic_rule_configs
    set max_difference_days = 91
    where rule_key = 'financial_documents_cgd_bank_statement';
    raise exception 'Expected the 90-day database cap to reject 91.';
  exception when check_violation then null;
  end;
  update public.financial_reconciliation_automatic_rule_configs
  set max_difference_days = 11
  where rule_key = 'financial_documents_cgd_bank_statement';
end $$;

do $$
declare
  v_exact_id uuid := '41000000-0000-0000-0000-000000000201';
  v_excluded_ids uuid[] := array[
    '41000000-0000-0000-0000-000000000202'::uuid,
    '41000000-0000-0000-0000-000000000203'::uuid,
    '41000000-0000-0000-0000-000000000204'::uuid,
    '41000000-0000-0000-0000-000000000205'::uuid
  ];
  v_candidate_ids uuid[];
  v_run jsonb;
  v_run_id uuid;
  v_continue_guard integer := 0;
begin
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name,
    payment, amount, fat
  ) values
    (v_exact_id, date '2027-01-20', 'BANCO-V2-201', '', '', 'Banco', 101.00, 'S'),
    (v_excluded_ids[1], date '2027-01-20', 'BANCO-V2-202', '', '', 'BANCO', 102.00, 'S'),
    (v_excluded_ids[2], date '2027-01-20', 'BANCO-V2-203', '', '', ' banco ', 103.00, 'S'),
    (v_excluded_ids[3], date '2027-01-20', 'BANCO-V2-204', '', '', '', 104.00, 'S'),
    (v_excluded_ids[4], date '2027-01-20', 'BANCO-V2-205', '', '', null, 105.00, 'S');

  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  ) values
    ('42000000-0000-0000-0000-000000000201', 'smoke-banco-v2', 'banco-v2-201', date '2027-01-20', 'Payment BANCOV2201', -101.00),
    ('42000000-0000-0000-0000-000000000202', 'smoke-banco-v2', 'banco-v2-202', date '2027-01-20', 'Payment BANCOV2202', -102.00),
    ('42000000-0000-0000-0000-000000000203', 'smoke-banco-v2', 'banco-v2-203', date '2027-01-20', 'Payment BANCOV2203', -103.00),
    ('42000000-0000-0000-0000-000000000204', 'smoke-banco-v2', 'banco-v2-204', date '2027-01-20', 'Payment BANCOV2204', -104.00),
    ('42000000-0000-0000-0000-000000000205', 'smoke-banco-v2', 'banco-v2-205', date '2027-01-20', 'Payment BANCOV2205', -105.00),
    ('42000000-0000-0000-0000-000000000298', 'smoke-banco-v2', 'banco-v2-sync', date '2027-01-20', 'Projection sync fixture', -298.00),
    ('42000000-0000-0000-0000-000000000299', 'smoke-banco-v2', 'banco-v2-undated', null, 'Undated ineligible row', -999.00);

  if (select count(*) from public.financial_reconciliation_cgd_match_search
      where source_id between '42000000-0000-0000-0000-000000000201'::uuid
                          and '42000000-0000-0000-0000-000000000205'::uuid) <> 5 then
    raise exception 'CGD match projection trigger did not synchronize inserted rows.';
  end if;
  if exists (
    select 1 from public.financial_reconciliation_cgd_match_search
    where source_id = '42000000-0000-0000-0000-000000000299'::uuid
  ) then
    raise exception 'CGD match projection retained an ineligible undated row.';
  end if;

  update public.import_cgd_extrato_ordem
  set id = '42000000-0000-0000-0000-000000000297',
      descritivo = 'Projection sync fixture updated'
  where id = '42000000-0000-0000-0000-000000000298';
  if exists (
      select 1 from public.financial_reconciliation_cgd_match_search
      where source_id = '42000000-0000-0000-0000-000000000298'
    ) or not exists (
      select 1 from public.financial_reconciliation_cgd_match_search
      where source_id = '42000000-0000-0000-0000-000000000297'
        and description = 'Projection sync fixture updated'
    ) then
    raise exception 'CGD match projection did not synchronize an ID/content update.';
  end if;
  delete from public.import_cgd_extrato_ordem
  where id = '42000000-0000-0000-0000-000000000297';
  if exists (
    select 1 from public.financial_reconciliation_cgd_match_search
    where source_id = '42000000-0000-0000-0000-000000000297'
  ) then
    raise exception 'CGD match projection did not synchronize a delete.';
  end if;

  select coalesce(array_agg(candidate.base_source_id order by candidate.base_source_id), '{}'::uuid[])
  into v_candidate_ids
  from public.financial_reconciliation_automatic_rule_candidates(
    'financial_documents_cgd_bank_statement', 2, 4.56, 11
  ) candidate
  where candidate.base_source_id = v_exact_id
     or candidate.base_source_id = any(v_excluded_ids);

  if v_candidate_ids is distinct from array[v_exact_id] then
    raise exception 'Exact Banco eligibility returned an unexpected base set: %', v_candidate_ids;
  end if;

  v_run := public.create_financial_reconciliation_automatic_analysis(
    array['financial_documents_cgd_bank_statement'],
    'manual_rule',
    'smoke:banco-v2',
    '43000000-0000-0000-0000-000000000201'
  );
  v_run_id := (v_run->>'runId')::uuid;
  while not coalesce((v_run->>'analysisComplete')::boolean, false) loop
    v_continue_guard := v_continue_guard + 1;
    if v_continue_guard > 100 then
      raise exception 'Resumable Banco analysis did not finish.';
    end if;
    v_run := public.continue_financial_reconciliation_automatic_analysis(
      v_run_id,
      'smoke:banco-v2'
    );
  end loop;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_proposals
    where run_id = v_run_id
      and base_source_id = v_exact_id
      and status = 'proposed'
  ) or exists (
    select 1
    from public.financial_reconciliation_automatic_proposals
    where run_id = v_run_id
      and base_source_id = any(v_excluded_ids)
  ) then
    raise exception 'Non-Banco bases created proposal or skipped rows.';
  end if;
end $$;

do $$
declare
  v_failed_run_id uuid;
  v_failed_run jsonb;
begin
  insert into public.financial_reconciliation_automatic_runs (
    trigger, scope, actor, client_request_id, definition_config_snapshot,
    analysis_processed, analysis_total
  ) values (
    'manual', 'rule', 'smoke:continuation-failure',
    '43000000-0000-0000-0000-000000000208',
    jsonb_build_array(jsonb_build_object(
      'ruleKey', 'financial_documents_cgd_bank_statement',
      'ruleVersion', 2,
      'displayName', 'Failure fixture',
      'priority', 1,
      'differenceAllowed', 0,
      'maxDifferenceDays', 7,
      'operator', '?'
    )),
    0, 1
  ) returning id into v_failed_run_id;

  begin
    perform public.continue_financial_reconciliation_automatic_analysis(
      v_failed_run_id,
      'smoke:continuation-intruder'
    );
    raise exception 'Expected continuation actor ownership validation.';
  exception when others then
    if sqlerrm not like '%belongs to another actor%' then raise; end if;
  end;
  if exists (
    select 1 from public.financial_reconciliation_automatic_runs
    where id = v_failed_run_id
      and (status <> 'analyzing' or analysis_error_code is not null)
  ) then
    raise exception 'Unauthorized continuation changed the owned run state.';
  end if;

  v_failed_run := public.continue_financial_reconciliation_automatic_analysis(
    v_failed_run_id,
    'smoke:continuation-failure'
  );
  if v_failed_run->>'status' <> 'failed'
    or v_failed_run->>'analysisErrorCode' <> 'analysis_continuation_failed'
    or nullif(v_failed_run->>'finishedAt', '') is null
    or exists (
      select 1 from public.financial_reconciliation_automatic_proposals
      where run_id = v_failed_run_id
    ) then
    raise exception 'Continuation failure did not persist one sanitized terminal run state.';
  end if;
end $$;

do $$
declare
  v_drift_document_id uuid := '41000000-0000-0000-0000-000000000206';
  v_run jsonb;
  v_run_id uuid;
  v_drift_proposal_id uuid;
  v_result jsonb;
  v_continue_guard integer := 0;
begin
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name,
    payment, amount, fat
  ) values (
    v_drift_document_id, date '2027-02-20', 'BANCO-V2-206', '', '',
    'Banco', 106.00, 'S'
  );
  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  ) values (
    '42000000-0000-0000-0000-000000000206',
    'smoke-banco-v2', 'banco-v2-206', date '2027-02-20',
    'Payment BANCOV2206', -106.00
  );

  v_run := public.create_financial_reconciliation_automatic_analysis(
    array['financial_documents_cgd_bank_statement'],
    'manual_rule',
    'smoke:banco-drift',
    '43000000-0000-0000-0000-000000000206'
  );
  v_run_id := (v_run->>'runId')::uuid;
  while not coalesce((v_run->>'analysisComplete')::boolean, false) loop
    v_continue_guard := v_continue_guard + 1;
    if v_continue_guard > 100 then
      raise exception 'Resumable Banco drift analysis did not finish.';
    end if;
    v_run := public.continue_financial_reconciliation_automatic_analysis(
      v_run_id,
      'smoke:banco-drift'
    );
  end loop;
  select id into strict v_drift_proposal_id
  from public.financial_reconciliation_automatic_proposals
  where run_id = v_run_id
    and base_source_id = v_drift_document_id
    and status = 'proposed';

  update public.financial_documents
  set payment = 'BANCO'
  where id = v_drift_document_id;

  v_result := public.execute_financial_reconciliation_automatic_proposal(
    v_drift_proposal_id,
    'smoke:banco-drift'
  );

  if v_result->>'status' <> 'stale'
    or v_result->>'reason' <> 'source_snapshot_changed'
    or exists (
      select 1 from public.financial_reconciliation_automatic_proposals
      where id = v_drift_proposal_id and reconciliation_id is not null
    ) then
    raise exception 'Payment drift created an automatic reconciliation.';
  end if;
end $$;

do $$
declare
  v_window_candidates jsonb;
begin
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name,
    payment, amount, fat
  ) values (
    '41000000-0000-0000-0000-000000000209', date '2027-04-01',
    'WINDOW-90-209', '', '', 'Banco', 209.00, 'S'
  );
  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  ) values
    ('42000000-0000-0000-0000-000000000209', 'smoke-window', 'window-day-90', date '2027-06-30', 'Payment WINDOW90209', -209.00),
    ('42000000-0000-0000-0000-000000000210', 'smoke-window', 'window-day-91', date '2027-07-01', 'Payment WINDOW90209', -209.00);

  select candidates into strict v_window_candidates
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_bank_statement', 2, 0, 90,
    array['41000000-0000-0000-0000-000000000209'::uuid]
  );
  if jsonb_array_length(v_window_candidates) <> 1
    or v_window_candidates->0->>'sourceId' <> '42000000-0000-0000-0000-000000000209' then
    raise exception 'The 90-day boundary did not include day 90 and exclude day 91.';
  end if;
end $$;

do $$
declare
  v_version_one_run_id uuid;
  v_version_one_proposal_id uuid;
  v_result jsonb;
begin
  insert into public.financial_reconciliation_automatic_runs (
    trigger, scope, actor, client_request_id, status, analysis_completed_at
  ) values (
    'manual', 'rule', 'smoke:legacy-version',
    '43000000-0000-0000-0000-000000000207', 'ready', now()
  ) returning id into v_version_one_run_id;

  insert into public.financial_reconciliation_automatic_proposals (
    run_id, rule_key, rule_version, base_source_type,
    base_source_id, base_source_date, allowed_difference, status, signature
  ) values (
    v_version_one_run_id,
    'financial_documents_cgd_bank_statement',
    1,
    'financial_documents',
    '41000000-0000-0000-0000-000000000201',
    date '2027-01-20',
    4.56,
    'proposed',
    'smoke-banco-v1-pending'
  ) returning id into v_version_one_proposal_id;

  v_result := public.execute_financial_reconciliation_automatic_proposal(
    v_version_one_proposal_id,
    'smoke:legacy-version'
  );

  if v_result->>'status' <> 'stale'
    or v_result->>'reason' <> 'rule_version_changed' then
    raise exception 'Pending version 1 proposal did not become stale.';
  end if;
end $$;

-- one-rule manual creation validation and immutable snapshots
do $$
declare
  v_before bigint;
  v_after bigint;
  v_run jsonb;
  v_snapshot jsonb;
begin
  update public.financial_documents
  set payment = case payment when 'Visa' then 'smoke:hidden-visa' when 'Banco' then 'smoke:hidden-banco' else payment end
  where payment in ('Visa', 'Banco');

  update public.financial_reconciliation_automatic_rule_configs
  set enabled = true, allow_manual_execution = true, include_in_scheduled_batch = false,
      difference_allowed = 0, max_difference_days = 10
  where rule_key in (
    'financial_documents_cgd_bank_statement',
    'financial_documents_cgd_credit_card'
  );

  select count(*) into v_before
  from public.financial_reconciliation_automatic_runs
  where actor like 'smoke:one-rule-validation%';

  begin
    perform public.create_financial_reconciliation_automatic_analysis(
      array['financial_documents_cgd_credit_card'], 'manual_batch',
      'smoke:one-rule-validation-mode', '53000000-0000-0000-0000-000000000001'
    );
    raise exception 'Expected manual_batch rejection.';
  exception when others then
    if sqlerrm not like 'Manual automatic analysis requires exactly one selected rule.%' then raise; end if;
  end;
  begin
    perform public.create_financial_reconciliation_automatic_analysis(
      array['financial_documents_cgd_bank_statement', 'financial_documents_cgd_credit_card'], 'manual_rule',
      'smoke:one-rule-validation-two', '53000000-0000-0000-0000-000000000002'
    );
    raise exception 'Expected two-rule rejection.';
  exception when others then
    if sqlerrm not like 'Manual automatic analysis requires exactly one selected rule.%' then raise; end if;
  end;
  begin
    perform public.create_financial_reconciliation_automatic_analysis(
      array['smoke_unknown_rule'], 'manual_rule',
      'smoke:one-rule-validation-unknown', '53000000-0000-0000-0000-000000000003'
    );
    raise exception 'Expected unknown-rule rejection.';
  exception when others then
    if sqlerrm not like 'Automatic rule is not enabled for manual analysis.%' then raise; end if;
  end;

  update public.financial_reconciliation_automatic_rule_configs
  set enabled = false, allow_manual_execution = true
  where rule_key = 'financial_documents_cgd_credit_card';
  begin
    perform public.create_financial_reconciliation_automatic_analysis(
      array['financial_documents_cgd_credit_card'], 'manual_rule',
      'smoke:one-rule-validation-disabled', '53000000-0000-0000-0000-000000000004'
    );
    raise exception 'Expected disabled-rule rejection.';
  exception when others then
    if sqlerrm not like 'Automatic rule is not enabled for manual analysis.%' then raise; end if;
  end;

  update public.financial_reconciliation_automatic_rule_configs
  set enabled = true, allow_manual_execution = false
  where rule_key = 'financial_documents_cgd_credit_card';
  begin
    perform public.create_financial_reconciliation_automatic_analysis(
      array['financial_documents_cgd_credit_card'], 'manual_rule',
      'smoke:one-rule-validation-not-manual', '53000000-0000-0000-0000-000000000005'
    );
    raise exception 'Expected manual-disabled rule rejection.';
  exception when others then
    if sqlerrm not like 'Automatic rule is not enabled for manual analysis.%' then raise; end if;
  end;

  select count(*) into v_after
  from public.financial_reconciliation_automatic_runs
  where actor like 'smoke:one-rule-validation%';
  if v_after <> v_before then
    raise exception 'Invalid one-rule requests inserted an automatic run.';
  end if;

  update public.financial_reconciliation_automatic_rule_configs
  set enabled = true, allow_manual_execution = true
  where rule_key = 'financial_documents_cgd_credit_card';
  v_run := public.create_financial_reconciliation_automatic_analysis(
    array['financial_documents_cgd_credit_card'], 'manual_rule',
    'smoke:one-rule-credit-snapshot', '53000000-0000-0000-0000-000000000006'
  );
  v_snapshot := v_run->'definitions';
  if jsonb_array_length(v_snapshot) <> 1
    or v_snapshot->0->>'ruleKey' <> 'financial_documents_cgd_credit_card'
    or v_snapshot->0->>'destinationSourceType' <> 'import_cgd_cartao_credito'
    or v_snapshot->0->>'operator' <> '+'
    or v_run->>'status' <> 'completed' then
    raise exception 'Credit-card manual run did not snapshot exactly one immutable directional definition.';
  end if;

  v_run := public.create_financial_reconciliation_automatic_analysis(
    array['financial_documents_cgd_bank_statement'], 'manual_rule',
    'smoke:one-rule-bank-snapshot', '53000000-0000-0000-0000-000000000007'
  );
  v_snapshot := v_run->'definitions';
  if jsonb_array_length(v_snapshot) <> 1
    or v_snapshot->0->>'ruleKey' <> 'financial_documents_cgd_bank_statement'
    or v_snapshot->0->>'destinationSourceType' <> 'import_cgd_extrato_ordem'
    or v_snapshot->0->>'operator' <> '+'
    or v_run->>'status' <> 'completed' then
    raise exception 'Banco manual run did not snapshot exactly one immutable directional definition.';
  end if;
end $$;

-- credit-card 25-base resumable lifecycle and retry idempotency
-- credit-card one-to-four exact combinations and five-card skip
-- credit-card ambiguity and candidate limit
do $$
declare
  v_run jsonb;
  v_retry jsonb;
  v_run_id uuid;
  v_proposal_count integer;
  v_cursor uuid;
begin
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, payment, amount, fat
  ) values
    ('51000000-0000-0000-0000-000000000001', date '2090-01-01', 'CC-ONE-001', '', '', 'Visa', 100, 'S'),
    ('51000000-0000-0000-0000-000000000002', date '2090-01-01', 'CC-TWO-002', '', '', 'Visa', 100, 'S'),
    ('51000000-0000-0000-0000-000000000003', date '2090-01-01', 'CC-THREE-003', '', '', 'Visa', 100, 'S'),
    ('51000000-0000-0000-0000-000000000004', date '2090-01-01', 'CC-FOUR-004', '', '', 'Visa', 100, 'S'),
    ('51000000-0000-0000-0000-000000000005', date '2090-01-01', 'CC-FIVE-005', '', '', 'Visa', 50, 'S'),
    ('51000000-0000-0000-0000-000000000006', date '2090-01-01', 'CC-MULTI-006', '', '', 'Visa', 20, 'S'),
    ('51000000-0000-0000-0000-000000000007', date '2090-01-01', 'CC-LIMIT-007', '', '', 'Visa', 130, 'S'),
    ('51000000-0000-0000-0000-000000000008', date '2090-01-01', 'CC-NONE-008', '', '', 'Visa', 80, 'S');

  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, payment, amount, fat
  )
  select
    ('51000000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
    date '2090-01-01', 'CC-FILL-' || lpad(series::text, 3, '0'), '', '', 'Visa', 100 + series, 'S'
  from generate_series(9, 28) series;

  insert into public.import_cgd_cartao_credito (
    id, import_batch, row_key, data, descricao, debito
  ) values
    ('52000000-0000-0000-0000-000000000001', 'smoke-one-rule', 'cc-one-1', date '2090-01-01', 'CCONE001', 100),
    ('52000000-0000-0000-0000-000000000002', 'smoke-one-rule', 'cc-two-1', date '2090-01-01', 'CCTWO002', 40),
    ('52000000-0000-0000-0000-000000000003', 'smoke-one-rule', 'cc-two-2', date '2090-01-01', 'CCTWO002', 60),
    ('52000000-0000-0000-0000-000000000004', 'smoke-one-rule', 'cc-three-1', date '2090-01-01', 'CCTHREE003', 20),
    ('52000000-0000-0000-0000-000000000005', 'smoke-one-rule', 'cc-three-2', date '2090-01-01', 'CCTHREE003', 30),
    ('52000000-0000-0000-0000-000000000006', 'smoke-one-rule', 'cc-three-3', date '2090-01-01', 'CCTHREE003', 50),
    ('52000000-0000-0000-0000-000000000007', 'smoke-one-rule', 'cc-four-1', date '2090-01-01', 'CCFOUR004', 10),
    ('52000000-0000-0000-0000-000000000008', 'smoke-one-rule', 'cc-four-2', date '2090-01-01', 'CCFOUR004', 20),
    ('52000000-0000-0000-0000-000000000009', 'smoke-one-rule', 'cc-four-3', date '2090-01-01', 'CCFOUR004', 30),
    ('52000000-0000-0000-0000-000000000010', 'smoke-one-rule', 'cc-four-4', date '2090-01-01', 'CCFOUR004', 40),
    ('52000000-0000-0000-0000-000000000011', 'smoke-one-rule', 'cc-five-1', date '2090-01-01', 'CCFIVE005', 10),
    ('52000000-0000-0000-0000-000000000012', 'smoke-one-rule', 'cc-five-2', date '2090-01-01', 'CCFIVE005', 10),
    ('52000000-0000-0000-0000-000000000013', 'smoke-one-rule', 'cc-five-3', date '2090-01-01', 'CCFIVE005', 10),
    ('52000000-0000-0000-0000-000000000014', 'smoke-one-rule', 'cc-five-4', date '2090-01-01', 'CCFIVE005', 10),
    ('52000000-0000-0000-0000-000000000015', 'smoke-one-rule', 'cc-five-5', date '2090-01-01', 'CCFIVE005', 10),
    ('52000000-0000-0000-0000-000000000016', 'smoke-one-rule', 'cc-multi-1', date '2090-01-01', 'CCMULTI006', 20),
    ('52000000-0000-0000-0000-000000000017', 'smoke-one-rule', 'cc-multi-2', date '2090-01-01', 'CCMULTI006', 10),
    ('52000000-0000-0000-0000-000000000018', 'smoke-one-rule', 'cc-multi-3', date '2090-01-01', 'CCMULTI006', 10);

  insert into public.import_cgd_cartao_credito (
    id, import_batch, row_key, data, descricao, debito
  )
  select
    ('52000000-0000-0000-0000-' || lpad((100 + series)::text, 12, '0'))::uuid,
    'smoke-one-rule', 'cc-limit-' || series, date '2090-01-01', 'CCLIMIT007', 10
  from generate_series(1, 13) series;

  v_run := public.create_financial_reconciliation_automatic_analysis(
    array['financial_documents_cgd_credit_card'], 'manual_rule',
    'smoke:credit-card-lifecycle', '53000000-0000-0000-0000-000000000010'
  );
  v_run_id := (v_run->>'runId')::uuid;
  if v_run->>'status' <> 'analyzing'
    or (v_run->>'analysisProcessed')::integer <> 25
    or (v_run->>'analysisTotal')::integer <> 28
    or v_run->>'analysisCursorId' <> '51000000-0000-0000-0000-000000000025'
    or (select count(*) from public.financial_reconciliation_automatic_proposals where run_id = v_run_id) <> 25 then
    raise exception 'Credit-card first analysis page was not the ordered 25-base page.';
  end if;
  v_cursor := (v_run->>'analysisCursorId')::uuid;

  v_run := public.continue_financial_reconciliation_automatic_analysis(
    v_run_id, 'smoke:credit-card-lifecycle'
  );
  if v_run->>'status' <> 'ready'
    or (v_run->>'analysisProcessed')::integer <> 28
    or v_run->>'analysisCursorId' <> '51000000-0000-0000-0000-000000000028'
    or (v_run->'counts'->>'bases')::integer <> 28
    or (v_run->'counts'->>'proposed')::integer <> 4
    or (v_run->'counts'->>'ambiguous')::integer <> 2
    or (v_run->'counts'->>'skipped')::integer <> 22
    or v_run->>'finishedAt' is not null then
    raise exception 'Credit-card final analysis state or persisted counts are invalid.';
  end if;
  if v_cursor = (v_run->>'analysisCursorId')::uuid then
    raise exception 'Credit-card continuation did not advance the cursor exactly to the remaining page.';
  end if;

  select count(*) into v_proposal_count
  from public.financial_reconciliation_automatic_proposals
  where run_id = v_run_id;
  v_retry := public.continue_financial_reconciliation_automatic_analysis(
    v_run_id, 'smoke:credit-card-lifecycle'
  );
  if (select count(*) from public.financial_reconciliation_automatic_proposals where run_id = v_run_id) <> v_proposal_count
    or v_retry->>'analysisCursorId' <> v_run->>'analysisCursorId'
    or (v_retry->>'analysisProcessed')::integer <> 28 then
    raise exception 'Credit-card continuation retry duplicated proposals or advanced the cursor.';
  end if;

  if exists (
    select 1
    from public.financial_reconciliation_automatic_proposals proposal
    where proposal.run_id = v_run_id
      and proposal.base_source_id between '51000000-0000-0000-0000-000000000001'::uuid
                                      and '51000000-0000-0000-0000-000000000004'::uuid
      and (proposal.status <> 'proposed'
        or jsonb_array_length(proposal.items) <> substring(proposal.base_source_id::text from 36 for 1)::integer)
  ) or (select count(*) from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = v_run_id and proposal.status = 'proposed') <> 4 then
    raise exception 'Credit-card one-to-four exact-zero combinations were not proposed.';
  end if;
  if not exists (
    select 1 from public.financial_reconciliation_automatic_proposals proposal
    where proposal.run_id = v_run_id
      and proposal.base_source_id = '51000000-0000-0000-0000-000000000005'
      and proposal.status = 'skipped'
      and proposal.reason = 'no_qualifying_combination'
      and jsonb_array_length(proposal.candidate_groups) = 5
  ) then
    raise exception 'A solution requiring five credit-card rows was not skipped.';
  end if;
  if not exists (
    select 1 from public.financial_reconciliation_automatic_proposals proposal
    where proposal.run_id = v_run_id
      and proposal.base_source_id = '51000000-0000-0000-0000-000000000006'
      and proposal.status = 'ambiguous' and proposal.reason = 'multiple_combinations'
      and jsonb_array_length(proposal.candidate_groups) = 2
  ) then
    raise exception 'Two valid credit-card combinations did not become ambiguous.';
  end if;
  if not exists (
    select 1 from public.financial_reconciliation_automatic_proposals proposal
    where proposal.run_id = v_run_id
      and proposal.base_source_id = '51000000-0000-0000-0000-000000000007'
      and proposal.status = 'ambiguous' and proposal.reason = 'candidate_limit'
      and jsonb_array_length(proposal.candidate_groups) = 13
  ) then
    raise exception 'Thirteen credit-card identity candidates did not hit the managed candidate limit.';
  end if;

  if public.get_financial_reconciliation_automatic_active_run('smoke:credit-card-lifecycle')->>'runId'
      <> v_run_id::text then
    raise exception 'Ready credit-card run was not retained as the actor active run.';
  end if;

  insert into public.financial_reconciliation_automatic_runs (
    trigger, scope, actor, client_request_id, status,
    definition_config_snapshot, analysis_completed_at, finished_at
  )
  select
    'manual', 'rule', run.actor, '53000000-0000-0000-0000-000000000012', 'completed',
    run.definition_config_snapshot, now(), now()
  from public.financial_reconciliation_automatic_runs run
  where run.id = v_run_id;
  begin
    perform public.create_financial_reconciliation_automatic_analysis(
      array['financial_documents_cgd_credit_card'], 'manual_rule',
      'smoke:credit-card-lifecycle', '53000000-0000-0000-0000-000000000012'
    );
    raise exception 'Expected the current open run to win over an older idempotency key.';
  exception when others then
    if sqlerrm not like 'Automatic analysis conflict: an unfinished manual run already exists for this actor.%' then raise; end if;
  end;

  begin
    perform public.create_financial_reconciliation_automatic_analysis(
      array['financial_documents_cgd_credit_card'], 'manual_rule',
      'smoke:credit-card-lifecycle', '53000000-0000-0000-0000-000000000011'
    );
    raise exception 'Expected one-open-manual-run conflict.';
  exception when others then
    if sqlerrm not like 'Automatic analysis conflict: an unfinished manual run already exists for this actor.%' then raise; end if;
  end;
  if (select count(*) from public.financial_reconciliation_automatic_runs
      where actor = 'smoke:credit-card-lifecycle' and finished_at is null) <> 1 then
    raise exception 'Conflicting manual creation inserted or changed the current open run.';
  end if;
end $$;

-- one open manual run per actor
do $$
begin
  if not exists (
    select 1 from pg_indexes
    where schemaname = 'public'
      and indexname = 'financial_reconciliation_automatic_runs_open_manual_actor_uidx'
  ) then
    raise exception 'The one-open-manual-run actor index was not installed.';
  end if;
end $$;

-- zero-executable analysis terminates without visible executable rows
do $$
declare
  v_run jsonb;
  v_run_id uuid;
begin
  update public.financial_documents set payment = 'smoke:analyzed-visa' where payment = 'Visa';
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, payment, amount, fat
  ) values (
    '51000000-0000-0000-0000-000000000029', date '2091-01-01',
    'CC-ZERO-029', '', '', 'Visa', 29, 'S'
  );

  v_run := public.create_financial_reconciliation_automatic_analysis(
    array['financial_documents_cgd_credit_card'], 'manual_rule',
    'smoke:zero-executable', '53000000-0000-0000-0000-000000000020'
  );
  v_run_id := (v_run->>'runId')::uuid;
  if v_run->>'status' <> 'completed'
    or v_run->>'finishedAt' is null
    or (v_run->'counts'->>'bases')::integer <> 1
    or (v_run->'counts'->>'skipped')::integer <> 1
    or (v_run->'counts'->>'proposed')::integer <> 0
    or public.get_financial_reconciliation_automatic_active_run('smoke:zero-executable') is not null then
    raise exception 'Zero-executable analysis did not finish terminally and release the active selector.';
  end if;
  if (select count(*) from public.financial_reconciliation_automatic_proposals
      where run_id = v_run_id and status = 'proposed') <> 0
    or not exists (
      select 1 from public.financial_reconciliation_automatic_proposals
      where run_id = v_run_id and status = 'skipped'
        and reason = 'no_qualifying_combination'
        and items = '[]'::jsonb
    ) then
    raise exception 'Skipped audit evidence leaked a visible executable proposal row contract.';
  end if;
end $$;

-- Banco paging and counts remain unchanged
do $$
declare
  v_count bigint;
  v_first_ids uuid[];
  v_second_ids uuid[];
  v_candidate_rows integer;
begin
  update public.financial_documents set payment = 'smoke:analyzed-banco' where payment = 'Banco';
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, payment, amount, fat
  )
  select
    ('54000000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
    date '2092-01-01', 'BANK-PAGE-' || lpad(series::text, 3, '0'), '', '', 'Banco', 200 + series, 'S'
  from generate_series(1, 26) series;

  select public.financial_reconciliation_automatic_base_count(
    'financial_documents_cgd_bank_statement', 2
  ) into v_count;
  select array_agg(page.id order by page.document_date, page.id)
  into v_first_ids
  from public.financial_reconciliation_automatic_base_page(
    'financial_documents_cgd_bank_statement', 2, null, null, 25
  ) page;
  select array_agg(page.id order by page.document_date, page.id)
  into v_second_ids
  from public.financial_reconciliation_automatic_base_page(
    'financial_documents_cgd_bank_statement', 2, date '2092-01-01',
    '54000000-0000-0000-0000-000000000025', 25
  ) page;
  select count(*) into v_candidate_rows
  from public.financial_reconciliation_automatic_candidate_page(
    'financial_documents_cgd_bank_statement', 2, 0, 10, null, null, 25
  );

  if v_count <> 26
    or cardinality(v_first_ids) <> 25
    or v_first_ids[1] <> '54000000-0000-0000-0000-000000000001'
    or v_first_ids[25] <> '54000000-0000-0000-0000-000000000025'
    or v_second_ids <> array['54000000-0000-0000-0000-000000000026'::uuid]
    or v_candidate_rows <> 25 then
    raise exception 'Banco base count or ordered 25-row paging behavior changed.';
  end if;
end $$;

-- automatic destination lock helper privileges and dispatch
do $$
declare
  v_definition text;
begin
  select pg_get_functiondef(
    'public.financial_reconciliation_automatic_lock_destination_items(text,jsonb)'::regprocedure
  ) into strict v_definition;

  if v_definition !~* 'security definer'
    or v_definition !~* 'if p_source_type = ''import_cgd_extrato_ordem'''
    or v_definition !~* 'elsif p_source_type = ''import_cgd_cartao_credito'''
    or v_definition ~* '\mexecute\M' then
    raise exception 'Automatic destination locking is not explicit and fixed-search-path.';
  end if;
  if not (
    select procedure.prosecdef
      and coalesce(procedure.proconfig, '{}'::text[]) @> array['search_path=public, pg_temp']
    from pg_proc procedure
    where procedure.oid =
      'public.financial_reconciliation_automatic_lock_destination_items(text,jsonb)'::regprocedure
  ) then
    raise exception 'Automatic destination locking did not retain its fixed search path.';
  end if;
  if has_function_privilege(
      'anon',
      'public.financial_reconciliation_automatic_lock_destination_items(text,jsonb)',
      'EXECUTE'
    )
    or has_function_privilege(
      'authenticated',
      'public.financial_reconciliation_automatic_lock_destination_items(text,jsonb)',
      'EXECUTE'
    )
    or has_function_privilege(
      'service_role',
      'public.financial_reconciliation_automatic_lock_destination_items(text,jsonb)',
      'EXECUTE'
    ) then
    raise exception 'The internal automatic destination lock helper is publicly executable.';
  end if;
  begin
    perform public.financial_reconciliation_automatic_lock_destination_items(
      'smoke_unsupported_source', '[]'::jsonb
    );
    raise exception 'Expected unsupported automatic destination source rejection.';
  exception when others then
    if sqlerrm <> 'Automatic reconciliation destination source is unsupported.' then raise; end if;
  end;
end $$;

create or replace function pg_temp.make_task4_proposal(
  p_rule_key text,
  p_base_source_id uuid,
  p_actor text,
  p_client_request_id uuid
)
returns jsonb
language plpgsql
as $$
declare
  v_run jsonb;
  v_run_id uuid;
  v_proposal_id uuid;
  v_guard integer := 0;
begin
  update public.financial_reconciliation_automatic_rule_configs
  set enabled = true,
      allow_manual_execution = true,
      include_in_scheduled_batch = false,
      difference_allowed = 0,
      max_difference_days = 10
  where rule_key = p_rule_key;

  v_run := public.create_financial_reconciliation_automatic_analysis(
    array[p_rule_key], 'manual_rule', p_actor, p_client_request_id
  );
  v_run_id := (v_run->>'runId')::uuid;
  while not coalesce((v_run->>'analysisComplete')::boolean, false) loop
    v_guard := v_guard + 1;
    if v_guard > 100 then
      raise exception 'Task 4 proposal analysis did not finish.';
    end if;
    v_run := public.continue_financial_reconciliation_automatic_analysis(v_run_id, p_actor);
  end loop;
  if v_run->>'status' <> 'ready' then
    raise exception 'Task 4 proposal analysis was not executable.';
  end if;
  select proposal.id into strict v_proposal_id
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.run_id = v_run_id
    and proposal.base_source_id = p_base_source_id
    and proposal.status = 'proposed';
  return jsonb_build_object('runId', v_run_id, 'proposalId', v_proposal_id);
end $$;

-- credit-card automatic execution and audit evidence
-- credit-card repeated execution is idempotent
do $$
declare
  v_document_id uuid := '61000000-0000-0000-0000-000000000001';
  v_run_id uuid;
  v_proposal_id uuid;
  v_reconciliation_id uuid;
  v_result jsonb;
  v_repeated jsonb;
  v_base_snapshot jsonb;
  v_destination_snapshots jsonb;
  v_evidence jsonb;
  v_signature text;
  v_item_count integer;
  v_audit_count integer;
begin
  update public.financial_documents
  set payment = 'smoke:before-task4-visa'
  where payment = 'Visa';
  update public.financial_reconciliation_source_rules
  set operator = '+'
  where base_source_type = 'financial_documents'
    and matching_source_type = 'import_cgd_cartao_credito';

  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, payment, amount, fat
  ) values (
    v_document_id, date '2093-01-10', 'CC-T4-EXEC-001',
    'Task 4 Visa execution', 'Task 4 Supplier', 'Visa', 100, 'S'
  );
  insert into public.import_cgd_cartao_credito (
    id, import_batch, row_key, data, descricao, debito
  ) values
    ('62000000-0000-0000-0000-000000000001', 'smoke-task4', 'task4-exec-1', date '2093-01-10', 'CCT4EXEC001 part one', 10),
    ('62000000-0000-0000-0000-000000000002', 'smoke-task4', 'task4-exec-2', date '2093-01-10', 'CCT4EXEC001 part two', 20),
    ('62000000-0000-0000-0000-000000000003', 'smoke-task4', 'task4-exec-3', date '2093-01-10', 'CCT4EXEC001 part three', 30),
    ('62000000-0000-0000-0000-000000000004', 'smoke-task4', 'task4-exec-4', date '2093-01-10', 'CCT4EXEC001 part four', 40);

  v_result := pg_temp.make_task4_proposal(
    'financial_documents_cgd_credit_card', v_document_id,
    'smoke:task4-execute', '63000000-0000-0000-0000-000000000001'
  );
  v_run_id := (v_result->>'runId')::uuid;
  v_proposal_id := (v_result->>'proposalId')::uuid;
  select base_snapshot, items, evidence, signature
  into strict v_base_snapshot, v_destination_snapshots, v_evidence, v_signature
  from public.financial_reconciliation_automatic_proposals
  where id = v_proposal_id;
  if jsonb_array_length(v_destination_snapshots) <> 4 then
    raise exception 'Task 4 fixture did not produce the maximum four-card proposal.';
  end if;

  v_result := public.execute_financial_reconciliation_automatic_proposal(
    v_proposal_id, 'smoke:task4-execute'
  );
  v_reconciliation_id := (v_result->>'reconciliationId')::uuid;
  if v_result is distinct from jsonb_build_object(
      'proposalId', v_proposal_id,
      'runId', v_run_id,
      'status', 'completed',
      'reconciliationId', v_reconciliation_id
    )
    or v_reconciliation_id is null then
    raise exception 'Credit-card execution changed the public completion outcome shape.';
  end if;
  if not exists (
    select 1
    from public.financial_reconciliations reconciliation
    where reconciliation.id = v_reconciliation_id
      and reconciliation.status = 'complete'
      and reconciliation.completion_type = 'normal'
      and reconciliation.difference_amount = 0
      and reconciliation.origin = 'automatic'
      and reconciliation.automatic_trigger = 'manual'
      and reconciliation.automatic_rule_key = 'financial_documents_cgd_credit_card'
      and reconciliation.automatic_rule_version = 1
      and reconciliation.automatic_run_id = v_run_id
      and reconciliation.automatic_proposal_id = v_proposal_id
      and reconciliation.matching_source_rules @> jsonb_build_array(jsonb_build_object(
        'sourceType', 'import_cgd_cartao_credito', 'operator', '+'
      ))
  ) then
    raise exception 'Credit-card reconciliation provenance, zero difference, or directional rule was not retained.';
  end if;
  if (select count(*) from public.financial_reconciliation_items
      where reconciliation_id = v_reconciliation_id) <> 5
    or (select count(*) from public.financial_reconciliation_items
        where reconciliation_id = v_reconciliation_id
          and source_type = 'financial_documents'
          and source_id = v_document_id) <> 1
    or (select count(*) from public.financial_reconciliation_items
        where reconciliation_id = v_reconciliation_id
          and source_type = 'import_cgd_cartao_credito') <> 4 then
    raise exception 'Credit-card execution did not lock one document and one through four card items.';
  end if;
  if not exists (
    select 1
    from public.financial_reconciliation_audit audit
    where audit.reconciliation_id = v_reconciliation_id
      and audit.action = 'automatic_complete'
      and audit.metadata @> jsonb_build_object(
        'ruleSnapshot', jsonb_build_object(
          'ruleKey', 'financial_documents_cgd_credit_card',
          'ruleVersion', 1,
          'definition', (
            select definition.definition
            from public.financial_reconciliation_automatic_rule_definitions definition
            where definition.rule_key = 'financial_documents_cgd_credit_card'
              and definition.version = 1
          )
        ),
        'configSnapshot', jsonb_build_object(
          'differenceAllowed', 0,
          'maxDifferenceDays', 10,
          'priority', 2
        ),
        'operatorSnapshot', jsonb_build_object('import_cgd_cartao_credito', '+'),
        'baseSnapshot', v_base_snapshot,
        'destinationSnapshots', v_destination_snapshots,
        'identityEvidence', v_evidence,
        'proposalSignature', v_signature,
        'trigger', 'manual',
        'runId', v_run_id,
        'proposalId', v_proposal_id,
        'tolerance', 0,
        'calculatedDifference', 0
      )
  ) then
    raise exception 'Credit-card automatic audit evidence was incomplete or changed.';
  end if;
  if not exists (
    select 1
    from public.financial_reconciliation_automatic_proposals proposal
    where proposal.id = v_proposal_id
      and proposal.status = 'completed'
      and proposal.reconciliation_id = v_reconciliation_id
      and proposal.base_snapshot = v_base_snapshot
      and proposal.items = v_destination_snapshots
      and proposal.evidence = v_evidence
      and proposal.signature = v_signature
      and proposal.calculated_difference = 0
  ) then
    raise exception 'Completed credit-card proposal lost immutable source snapshots or evidence.';
  end if;

  select count(*) into v_item_count
  from public.financial_reconciliation_items
  where reconciliation_id = v_reconciliation_id;
  select count(*) into v_audit_count
  from public.financial_reconciliation_audit
  where reconciliation_id = v_reconciliation_id;
  v_repeated := public.execute_financial_reconciliation_automatic_proposal(
    v_proposal_id, 'smoke:task4-execute'
  );
  if v_repeated is distinct from v_result
    or (select count(*) from public.financial_reconciliations
        where automatic_proposal_id = v_proposal_id) <> 1
    or (select count(*) from public.financial_reconciliation_items
        where reconciliation_id = v_reconciliation_id) <> v_item_count
    or (select count(*) from public.financial_reconciliation_audit
        where reconciliation_id = v_reconciliation_id) <> v_audit_count then
    raise exception 'Repeated credit-card execution was not idempotent.';
  end if;
end $$;

-- credit-card execution stale source and proposal paths
do $$
declare
  v_kinds text[] := array[
    'payment', 'card_data', 'card_value', 'card_description',
    'selected_ids', 'rule_version', 'definition', 'operator',
    'tolerance', 'evidence', 'locked', 'deleted'
  ];
  v_kind text;
  v_index integer := 0;
  v_document_id uuid;
  v_card_id uuid;
  v_request_id uuid;
  v_run_id uuid;
  v_proposal_id uuid;
  v_lock_id uuid;
  v_original_definition jsonb;
  v_result jsonb;
begin
  update public.financial_documents
  set payment = 'smoke:before-task4-stale'
  where payment = 'Visa';
  select definition into strict v_original_definition
  from public.financial_reconciliation_automatic_rule_definitions
  where rule_key = 'financial_documents_cgd_credit_card' and version = 1;

  foreach v_kind in array v_kinds loop
    v_index := v_index + 1;
    v_document_id := ('64000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;
    v_card_id := ('65000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;
    v_request_id := ('66000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;

    insert into public.financial_documents (
      id, document_date, doc_number, description, supplier_name, payment, amount, fat
    ) values (
      v_document_id, date '2093-02-01', 'CCT4STALE' || lpad(v_index::text, 3, '0'),
      '', '', 'Visa', 50, 'S'
    );
    insert into public.import_cgd_cartao_credito (
      id, import_batch, row_key, data, descricao, debito
    ) values (
      v_card_id, 'smoke-task4-stale', 'task4-stale-' || v_index,
      date '2093-02-01', 'CCT4STALE' || lpad(v_index::text, 3, '0'), 50
    );
    v_result := pg_temp.make_task4_proposal(
      'financial_documents_cgd_credit_card', v_document_id,
      'smoke:task4-stale-' || v_kind, v_request_id
    );
    v_run_id := (v_result->>'runId')::uuid;
    v_proposal_id := (v_result->>'proposalId')::uuid;
    v_lock_id := null;

    if v_kind = 'payment' then
      update public.financial_documents set payment = 'VISA' where id = v_document_id;
    elsif v_kind = 'card_data' then
      update public.import_cgd_cartao_credito set data = date '2093-03-01' where id = v_card_id;
    elsif v_kind = 'card_value' then
      update public.import_cgd_cartao_credito set debito = 51 where id = v_card_id;
    elsif v_kind = 'card_description' then
      update public.import_cgd_cartao_credito set descricao = 'Identity removed' where id = v_card_id;
    elsif v_kind = 'selected_ids' then
      update public.financial_reconciliation_automatic_proposals
      set items = jsonb_set(
        items, '{0,sourceId}', to_jsonb('65000000-0000-0000-0000-999999999999'::text)
      )
      where id = v_proposal_id;
    elsif v_kind = 'rule_version' then
      update public.financial_reconciliation_automatic_runs
      set definition_config_snapshot = jsonb_set(
        definition_config_snapshot, '{0,ruleVersion}', '2'::jsonb
      )
      where id = v_run_id;
    elsif v_kind = 'definition' then
      update public.financial_reconciliation_automatic_rule_definitions
      set definition = definition || '{"task4Drift":true}'::jsonb
      where rule_key = 'financial_documents_cgd_credit_card' and version = 1;
    elsif v_kind = 'operator' then
      update public.financial_reconciliation_source_rules
      set operator = '-'
      where base_source_type = 'financial_documents'
        and matching_source_type = 'import_cgd_cartao_credito';
    elsif v_kind = 'tolerance' then
      update public.financial_reconciliation_automatic_proposals
      set allowed_difference = 1
      where id = v_proposal_id;
    elsif v_kind = 'evidence' then
      update public.financial_reconciliation_automatic_proposals
      set evidence = evidence || jsonb_build_array(jsonb_build_object('task4Tampered', true))
      where id = v_proposal_id;
    elsif v_kind = 'locked' then
      v_result := public.financial_reconciliation_action(
        'start', 'smoke:task4-lock', null,
        'import_cgd_cartao_credito', v_card_id, null
      );
      v_lock_id := (v_result#>>'{reconciliation,id}')::uuid;
    elsif v_kind = 'deleted' then
      delete from public.import_cgd_cartao_credito where id = v_card_id;
    end if;

    v_result := public.execute_financial_reconciliation_automatic_proposal(
      v_proposal_id, 'smoke:task4-stale-' || v_kind
    );
    if v_result->>'status' <> 'stale'
      or not exists (
        select 1
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.id = v_proposal_id
          and proposal.status = 'stale'
          and proposal.reconciliation_id is null
      )
      or exists (
        select 1
        from public.financial_reconciliations reconciliation
        where reconciliation.automatic_proposal_id = v_proposal_id
      ) then
      raise exception 'Task 4 stale path % created an automatic reconciliation.', v_kind;
    end if;

    if v_kind = 'definition' then
      update public.financial_reconciliation_automatic_rule_definitions
      set definition = v_original_definition
      where rule_key = 'financial_documents_cgd_credit_card' and version = 1;
    elsif v_kind = 'operator' then
      update public.financial_reconciliation_source_rules
      set operator = '+'
      where base_source_type = 'financial_documents'
        and matching_source_type = 'import_cgd_cartao_credito';
    elsif v_kind = 'locked' then
      perform public.financial_reconciliation_action(
        'delete', 'smoke:task4-lock', v_lock_id, null, null, null
      );
    end if;
  end loop;
end $$;

-- automatic execution rejects unfinished and non-executable proposals
do $$
declare
  v_status text;
  v_index integer := 0;
  v_document_id uuid;
  v_card_id uuid;
  v_run_id uuid;
  v_proposal_id uuid;
  v_fixture jsonb;
begin
  update public.financial_documents
  set payment = 'smoke:before-task4-status'
  where payment = 'Visa';

  foreach v_status in array array['analyzing', 'ambiguous', 'skipped', 'deselected'] loop
    v_index := v_index + 1;
    v_document_id := ('67000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;
    v_card_id := ('68000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;
    insert into public.financial_documents (
      id, document_date, doc_number, description, supplier_name, payment, amount, fat
    ) values (
      v_document_id, date '2093-04-01', 'CCT4STATUS' || lpad(v_index::text, 3, '0'),
      '', '', 'Visa', 25, 'S'
    );
    insert into public.import_cgd_cartao_credito (
      id, import_batch, row_key, data, descricao, debito
    ) values (
      v_card_id, 'smoke-task4-status', 'task4-status-' || v_index,
      date '2093-04-01', 'CCT4STATUS' || lpad(v_index::text, 3, '0'), 25
    );
    v_fixture := pg_temp.make_task4_proposal(
      'financial_documents_cgd_credit_card', v_document_id,
      'smoke:task4-status-' || v_status,
      ('69000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid
    );
    v_run_id := (v_fixture->>'runId')::uuid;
    v_proposal_id := (v_fixture->>'proposalId')::uuid;
    if v_status = 'analyzing' then
      update public.financial_reconciliation_automatic_runs
      set status = 'analyzing', analysis_completed_at = null
      where id = v_run_id;
    else
      update public.financial_reconciliation_automatic_proposals
      set status = v_status
      where id = v_proposal_id;
    end if;

    begin
      perform public.execute_financial_reconciliation_automatic_proposal(
        v_proposal_id, 'smoke:task4-status-' || v_status
      );
      raise exception 'Expected Task 4 non-executable status rejection for %.', v_status;
    exception when others then
      if v_status = 'analyzing'
        and sqlerrm not like 'Automatic analysis must finish before proposals can be executed.%' then
        raise;
      elsif v_status <> 'analyzing'
        and sqlerrm not like 'Automation proposal with status % cannot be executed.%' then
        raise;
      end if;
    end;
    if exists (
      select 1
      from public.financial_reconciliations reconciliation
      where reconciliation.automatic_proposal_id = v_proposal_id
    ) then
      raise exception 'Non-executable Task 4 proposal % created a reconciliation.', v_status;
    end if;
  end loop;
end $$;

-- Banco v2 execution evidence remains unchanged
do $$
declare
  v_document_id uuid := '6a000000-0000-0000-0000-000000000001';
  v_run_id uuid;
  v_proposal_id uuid;
  v_reconciliation_id uuid;
  v_fixture jsonb;
  v_result jsonb;
begin
  update public.financial_documents
  set payment = 'smoke:before-task4-banco'
  where payment = 'Banco';
  update public.financial_reconciliation_source_rules
  set operator = '+'
  where base_source_type = 'financial_documents'
    and matching_source_type = 'import_cgd_extrato_ordem';

  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, payment, amount, fat
  ) values (
    v_document_id, date '2093-05-01', 'BANK-T4-001',
    'Banco Task 4 regression', 'Banco Supplier', 'Banco', 77, 'S'
  );
  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  ) values (
    '6b000000-0000-0000-0000-000000000001', 'smoke-task4-banco',
    'task4-banco-1', date '2093-05-01', 'Payment BANKT4001', -77
  );
  v_fixture := pg_temp.make_task4_proposal(
    'financial_documents_cgd_bank_statement', v_document_id,
    'smoke:task4-banco', '6c000000-0000-0000-0000-000000000001'
  );
  v_run_id := (v_fixture->>'runId')::uuid;
  v_proposal_id := (v_fixture->>'proposalId')::uuid;
  v_result := public.execute_financial_reconciliation_automatic_proposal(
    v_proposal_id, 'smoke:task4-banco'
  );
  v_reconciliation_id := (v_result->>'reconciliationId')::uuid;

  if v_result->>'status' <> 'completed'
    or not exists (
      select 1
      from public.financial_reconciliations reconciliation
      where reconciliation.id = v_reconciliation_id
        and reconciliation.difference_amount = 0
        and reconciliation.automatic_rule_key = 'financial_documents_cgd_bank_statement'
        and reconciliation.automatic_rule_version = 2
        and reconciliation.matching_source_rules @> jsonb_build_array(jsonb_build_object(
          'sourceType', 'import_cgd_extrato_ordem', 'operator', '+'
        ))
    )
    or not exists (
      select 1
      from public.financial_reconciliation_audit audit
      where audit.reconciliation_id = v_reconciliation_id
        and audit.action = 'automatic_complete'
        and audit.metadata @> jsonb_build_object(
          'operatorSnapshot', jsonb_build_object('import_cgd_extrato_ordem', '+'),
          'trigger', 'manual',
          'runId', v_run_id,
          'proposalId', v_proposal_id
        )
        and audit.metadata ?& array[
          'ruleSnapshot', 'configSnapshot', 'operatorSnapshot',
          'baseSnapshot', 'destinationSnapshots', 'identityEvidence',
          'proposalSignature', 'trigger', 'runId', 'proposalId',
          'tolerance', 'calculatedDifference'
        ]
    ) then
    raise exception 'Banco v2 execution or historical evidence changed under adapter dispatch.';
  end if;
  if not exists (
    select 1
    from public.financial_reconciliation_audit audit
    where audit.actor = 'smoke:automatic-nonzero'
      and audit.action = 'automatic_complete'
      and audit.metadata->'operatorSnapshot' = jsonb_build_object(
        'import_cgd_extrato_ordem', '+'
      )
  ) then
    raise exception 'Historical Banco automatic audit evidence was rewritten.';
  end if;
end $$;

-- scheduled parent batch schema security and legacy backfill
-- historical scheduled runs remain readable and cannot execute again
do $$
declare
  v_legacy_run_id uuid;
  v_legacy_batch_id uuid;
  v_public_run jsonb;
  v_claim jsonb;
  v_run_count integer;
  v_signature text;
begin
  select run.id, run.batch_id
  into strict v_legacy_run_id, v_legacy_batch_id
  from public.financial_reconciliation_automatic_runs run
  where run.trigger = 'scheduled'
    and run.scope = 'batch'
    and run.scheduled_slot = '2026-03-29';

  if not exists (
      select 1
      from public.financial_reconciliation_automatic_batches batch
      where batch.id = v_legacy_batch_id
        and batch.scheduled_slot = '2026-03-29'
        and batch.status = 'failed'
        and batch.finished_at is not null
    )
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_runs run
      where run.id = v_legacy_run_id
        and run.status = 'failed'
        and run.finished_at is not null
        and run.analysis_error_code = 'analysis_upgrade_restart_required'
        and run.error_summary = 'Analysis must be restarted after the 90-day performance upgrade.'
    )
    or (select batch.rule_snapshot
        from public.financial_reconciliation_automatic_batches batch
        where batch.id = v_legacy_batch_id) is distinct from (
      select run.definition_config_snapshot
      from public.financial_reconciliation_automatic_runs run
      where run.id = v_legacy_run_id
    )
    or (select batch.counts
        from public.financial_reconciliation_automatic_batches batch
        where batch.id = v_legacy_batch_id) is distinct from (
      select run.counts
      from public.financial_reconciliation_automatic_runs run
      where run.id = v_legacy_run_id
    ) then
    raise exception 'Legacy scheduled run was not safely backfilled into one terminal batch.';
  end if;

  v_public_run := public.get_financial_reconciliation_automatic_run(v_legacy_run_id);
  if v_public_run->>'runId' <> v_legacy_run_id::text
    or v_public_run->>'scope' <> 'batch'
    or v_public_run->>'batchId' <> v_legacy_batch_id::text
    or v_public_run->>'batchRuleKey' is not null
    or v_public_run->>'batchRulePosition' is not null
    or v_public_run->>'batchRuleCount' is not null
    or v_public_run->'definitions' is distinct from (
      select run.definition_config_snapshot
      from public.financial_reconciliation_automatic_runs run
      where run.id = v_legacy_run_id
    ) then
    raise exception 'Historical scheduled run is no longer readable with its immutable snapshot.';
  end if;

  select public.get_financial_reconciliation_automatic_run(run.id)
  into strict v_public_run
  from public.financial_reconciliation_automatic_runs run
  where run.trigger = 'manual'
  order by run.started_at, run.id
  limit 1;
  if v_public_run->>'batchId' is not null
    or v_public_run->>'batchRuleKey' is not null
    or v_public_run->>'batchRulePosition' is not null
    or v_public_run->>'batchRuleCount' is not null then
    raise exception 'Manual automatic run exposed scheduled batch metadata.';
  end if;

  if not exists (
      select 1
      from pg_class relation
      join pg_namespace namespace on namespace.oid = relation.relnamespace
      where namespace.nspname = 'public'
        and relation.relname = 'financial_reconciliation_automatic_batches'
        and relation.relrowsecurity
    )
    or has_table_privilege('anon', 'public.financial_reconciliation_automatic_batches', 'SELECT')
    or has_table_privilege('authenticated', 'public.financial_reconciliation_automatic_batches', 'SELECT')
    or has_table_privilege('service_role', 'public.financial_reconciliation_automatic_batches', 'SELECT') then
    raise exception 'Scheduled parent batch RLS or direct table privileges are invalid.';
  end if;

  foreach v_signature in array array[
    'public.financial_reconciliation_refresh_automatic_batch(uuid)',
    'public.claim_financial_reconciliation_automatic_schedule(timestamptz,text)',
    'public.get_financial_reconciliation_automatic_run(uuid)',
    'public.financial_reconciliation_automatic_progress_or_run(uuid)',
    'public.get_financial_reconciliation_automation_settings()',
    'public.replace_financial_reconciliation_source_rules(jsonb)'
  ] loop
    if not (
      select procedure.prosecdef
        and coalesce(procedure.proconfig, '{}'::text[]) @> array['search_path=public, pg_temp']
      from pg_proc procedure
      where procedure.oid = v_signature::regprocedure
    )
      or has_function_privilege('anon', v_signature, 'EXECUTE')
      or has_function_privilege('authenticated', v_signature, 'EXECUTE')
      or not has_function_privilege('service_role', v_signature, 'EXECUTE') then
      raise exception 'Scheduled batch RPC security changed for %.', v_signature;
    end if;
  end loop;
  if has_function_privilege(
      'service_role',
      'public.financial_reconciliation_refresh_automatic_batch_from_run()',
      'EXECUTE'
    )
    or (select count(*) from public.financial_reconciliation_automatic_batches
        where scheduled_slot = '2026-03-29') <> 1
    or (select count(*) from pg_trigger
        where tgrelid = 'public.financial_reconciliation_automatic_runs'::regclass
          and tgname = 'financial_reconciliation_refresh_automatic_batch_trigger'
          and not tgisinternal) <> 1
    or (select count(*) from pg_indexes
        where schemaname = 'public'
          and indexname in (
            'financial_reconciliation_automatic_runs_legacy_scheduled_slot_uidx',
            'financial_reconciliation_automatic_runs_batch_position_uidx',
            'financial_reconciliation_automatic_runs_batch_rule_uidx'
          )) <> 3 then
    raise exception 'Scheduled batch migration reapply duplicated or exposed an internal object.';
  end if;

  update public.financial_reconciliation_automatic_schedule
  set enabled = true, time_of_day = time '00:00', time_zone = 'Europe/Lisbon'
  where id = true;
  select count(*) into v_run_count
  from public.financial_reconciliation_automatic_runs;
  v_claim := public.claim_financial_reconciliation_automatic_schedule(
    '2026-03-29 02:00:00+00', 'smoke:historical-schedule'
  );
  if v_claim->>'reason' <> 'batch_complete'
    or v_claim->>'batchId' <> v_legacy_batch_id::text
    or (select count(*) from public.financial_reconciliation_automatic_runs) <> v_run_count then
    raise exception 'Historical scheduled batch was re-executed by a later heartbeat.';
  end if;
end $$;

-- scheduled batch snapshots all rules in deterministic priority order
-- scheduled child resumes before the next rule starts
-- scheduled snapshot survives settings changes and tomorrow uses new settings
-- scheduled retries and cross-midnight heartbeats are idempotent
-- completed scheduled batch returns stable no-work state
do $$
declare
  v_first jsonb;
  v_retry jsonb;
  v_cross_midnight jsonb;
  v_second jsonb;
  v_complete jsonb;
  v_complete_retry jsonb;
  v_tomorrow jsonb;
  v_batch_id uuid;
  v_first_run_id uuid;
  v_second_run_id uuid;
  v_tomorrow_batch_id uuid;
  v_snapshot jsonb;
  v_counts jsonb;
  v_proposal_count integer;
  v_reconciliation_count integer;
begin
  set constraints financial_reconciliation_automatic_rule_configs_priority_key deferred;
  update public.financial_reconciliation_automatic_rule_configs config
  set enabled = true,
      include_in_scheduled_batch = true,
      difference_allowed = case config.rule_key
        when 'financial_documents_cgd_bank_statement' then 0.10 else 0.20 end,
      max_difference_days = case config.rule_key
        when 'financial_documents_cgd_bank_statement' then 10 else 11 end,
      priority = case config.rule_key
        when 'financial_documents_cgd_bank_statement' then 1 else 2 end
  where config.rule_key in (
    'financial_documents_cgd_bank_statement',
    'financial_documents_cgd_credit_card'
  );

  v_first := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-01 01:00:00+00', 'smoke:scheduled-batch'
  );
  v_batch_id := (v_first->>'batchId')::uuid;
  v_first_run_id := (v_first#>>'{run,runId}')::uuid;
  select batch.rule_snapshot into strict v_snapshot
  from public.financial_reconciliation_automatic_batches batch
  where batch.id = v_batch_id;

  if not (v_first->>'claimed')::boolean
    or (v_first->>'resumed')::boolean
    or v_first->>'batchRulePosition' <> '1'
    or v_first->>'batchRuleCount' <> '2'
    or v_first#>>'{run,scope}' <> 'rule'
    or v_first#>>'{run,batchId}' <> v_batch_id::text
    or v_first#>>'{run,batchRuleKey}' <> 'financial_documents_cgd_bank_statement'
    or v_first#>>'{run,batchRulePosition}' <> '1'
    or v_first#>>'{run,batchRuleCount}' <> '2'
    or jsonb_array_length(v_first#>'{run,definitions}') <> 1
    or jsonb_array_length(v_snapshot) <> 2
    or v_snapshot#>>'{0,ruleKey}' <> 'financial_documents_cgd_bank_statement'
    or v_snapshot#>>'{0,ruleVersion}' <> '2'
    or v_snapshot#>>'{0,destinationSourceType}' <> 'import_cgd_extrato_ordem'
    or v_snapshot#>>'{0,operator}' <> '+'
    or v_snapshot#>>'{0,priority}' <> '1'
    or v_snapshot#>>'{0,differenceAllowed}' <> '0.10'
    or v_snapshot#>>'{0,maxDifferenceDays}' <> '10'
    or jsonb_typeof(v_snapshot#>'{0,definition}') <> 'object'
    or v_snapshot#>>'{1,ruleKey}' <> 'financial_documents_cgd_credit_card'
    or v_snapshot#>>'{1,ruleVersion}' <> '1'
    or v_snapshot#>>'{1,destinationSourceType}' <> 'import_cgd_cartao_credito'
    or v_snapshot#>>'{1,operator}' <> '+'
    or v_snapshot#>>'{1,priority}' <> '2'
    or v_snapshot#>>'{1,differenceAllowed}' <> '0.20'
    or v_snapshot#>>'{1,maxDifferenceDays}' <> '11'
    or jsonb_typeof(v_snapshot#>'{1,definition}') <> 'object' then
    raise exception 'Scheduled batch did not snapshot both managed rules in deterministic order.';
  end if;

  select count(*) into v_proposal_count
  from public.financial_reconciliation_automatic_proposals
  where run_id = v_first_run_id;
  select count(*) into v_reconciliation_count
  from public.financial_reconciliations
  where automatic_run_id = v_first_run_id;

  v_retry := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-01 01:01:00+00', 'smoke:scheduled-batch'
  );
  v_cross_midnight := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-02 00:30:00+00', 'smoke:scheduled-batch'
  );
  if not (v_retry->>'resumed')::boolean
    or not (v_cross_midnight->>'resumed')::boolean
    or v_retry#>>'{run,runId}' <> v_first_run_id::text
    or v_cross_midnight#>>'{run,runId}' <> v_first_run_id::text
    or (select count(*) from public.financial_reconciliation_automatic_batches
        where scheduled_slot = '2094-01-01') <> 1
    or (select count(*) from public.financial_reconciliation_automatic_runs
        where batch_id = v_batch_id) <> 1
    or (select count(*) from public.financial_reconciliation_automatic_proposals
        where run_id = v_first_run_id) <> v_proposal_count
    or (select count(*) from public.financial_reconciliations
        where automatic_run_id = v_first_run_id) <> v_reconciliation_count then
    raise exception 'Scheduled retry or cross-midnight heartbeat duplicated batch child work.';
  end if;

  set constraints financial_reconciliation_automatic_rule_configs_priority_key deferred;
  update public.financial_reconciliation_automatic_rule_configs config
  set difference_allowed = case config.rule_key
        when 'financial_documents_cgd_credit_card' then 0.30 else 0.40 end,
      max_difference_days = case config.rule_key
        when 'financial_documents_cgd_credit_card' then 12 else 13 end,
      priority = case config.rule_key
        when 'financial_documents_cgd_credit_card' then 1 else 2 end
  where config.rule_key in (
    'financial_documents_cgd_bank_statement',
    'financial_documents_cgd_credit_card'
  );
  if (select batch.rule_snapshot from public.financial_reconciliation_automatic_batches batch
      where batch.id = v_batch_id) is distinct from v_snapshot then
    raise exception 'Settings changes rewrote the active scheduled batch snapshot.';
  end if;

  update public.financial_reconciliation_automatic_runs
  set status = 'completed',
      analysis_completed_at = coalesce(analysis_completed_at, now()),
      counts = '{"bases":0,"completed":0,"failed":0}'::jsonb,
      finished_at = now(),
      updated_at = now()
  where id = v_first_run_id;
  v_second := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-02 00:31:00+00', 'smoke:scheduled-batch'
  );
  v_second_run_id := (v_second#>>'{run,runId}')::uuid;
  if (v_second->>'resumed')::boolean
    or v_second->>'batchId' <> v_batch_id::text
    or v_second->>'batchRulePosition' <> '2'
    or v_second#>>'{run,batchRuleKey}' <> 'financial_documents_cgd_credit_card'
    or (select count(*) from public.financial_reconciliation_automatic_runs
        where batch_id = v_batch_id) <> 2 then
    raise exception 'Scheduled batch started the wrong second child.';
  end if;

  update public.financial_reconciliation_automatic_runs
  set status = 'completed',
      analysis_completed_at = coalesce(analysis_completed_at, now()),
      counts = '{"bases":0,"completed":0,"failed":0}'::jsonb,
      finished_at = now(),
      updated_at = now()
  where id = v_second_run_id;
  v_complete := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-01 02:00:00+00', 'smoke:scheduled-batch'
  );
  select batch.counts into strict v_counts
  from public.financial_reconciliation_automatic_batches batch
  where batch.id = v_batch_id and batch.status = 'completed';
  v_complete_retry := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-01 02:01:00+00', 'smoke:scheduled-batch'
  );
  if v_complete->>'reason' <> 'batch_complete'
    or v_complete_retry->>'reason' <> 'batch_complete'
    or v_complete_retry->>'batchId' <> v_batch_id::text
    or (select batch.counts from public.financial_reconciliation_automatic_batches batch
        where batch.id = v_batch_id) is distinct from v_counts then
    raise exception 'Completed scheduled batch did not return stable no-work state.';
  end if;

  v_tomorrow := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-02 01:00:00+00', 'smoke:scheduled-batch'
  );
  v_tomorrow_batch_id := (v_tomorrow->>'batchId')::uuid;
  if v_tomorrow_batch_id = v_batch_id
    or v_tomorrow#>>'{run,batchRuleKey}' <> 'financial_documents_cgd_credit_card'
    or (select batch.rule_snapshot#>>'{0,priority}'
        from public.financial_reconciliation_automatic_batches batch
        where batch.id = v_tomorrow_batch_id) <> '1'
    or (select batch.rule_snapshot#>>'{0,differenceAllowed}'
        from public.financial_reconciliation_automatic_batches batch
        where batch.id = v_tomorrow_batch_id) <> '0.30' then
    raise exception 'Tomorrow scheduled batch did not use the changed settings snapshot.';
  end if;
end $$;

-- failed scheduled child advances and aggregate batch becomes partial
do $$
declare
  v_first jsonb;
  v_second jsonb;
  v_complete jsonb;
  v_batch_id uuid;
  v_first_run_id uuid;
  v_second_run_id uuid;
  v_settings jsonb;
begin
  v_first := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-02 01:01:00+00', 'smoke:scheduled-batch'
  );
  v_batch_id := (v_first->>'batchId')::uuid;
  v_first_run_id := (v_first#>>'{run,runId}')::uuid;
  update public.financial_reconciliation_automatic_runs
  set status = 'failed',
      error_summary = 'internal scheduled fixture failure',
      analysis_error_code = 'analysis_continuation_failed',
      analysis_error_at = now(),
      finished_at = now(),
      updated_at = now()
  where id = v_first_run_id;

  v_second := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-02 01:02:00+00', 'smoke:scheduled-batch'
  );
  v_second_run_id := (v_second#>>'{run,runId}')::uuid;
  if v_second->>'batchId' <> v_batch_id::text
    or v_second#>>'{run,batchRuleKey}' <> 'financial_documents_cgd_bank_statement'
    or v_second->>'batchRulePosition' <> '2' then
    raise exception 'A failed scheduled child blocked the next snapshotted rule.';
  end if;

  update public.financial_reconciliation_automatic_runs
  set status = 'completed',
      analysis_completed_at = coalesce(analysis_completed_at, now()),
      counts = '{"bases":1,"completed":1,"failed":0}'::jsonb,
      finished_at = now(),
      updated_at = now()
  where id = v_second_run_id;
  v_complete := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-02 01:03:00+00', 'smoke:scheduled-batch'
  );
  if v_complete->>'reason' <> 'batch_complete'
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_batches batch
      where batch.id = v_batch_id
        and batch.status = 'partial'
        and batch.counts @> '{"ruleCount":2,"childCount":2,"completedChildren":1,"failedChildren":1}'::jsonb
    ) then
    raise exception 'Mixed scheduled child outcomes did not produce a terminal partial batch.';
  end if;

  v_settings := public.get_financial_reconciliation_automation_settings();
  if v_settings#>>'{last_scheduled_batch,id}' <> v_batch_id::text
    or v_settings#>>'{last_scheduled_batch,status}' <> 'partial'
    or v_settings#>'{last_scheduled_batch,counts}' is null
    or v_settings::text like '%internal scheduled fixture failure%'
    or v_settings::text like '%analysis_continuation_failed%' then
    raise exception 'Settings did not expose the safe latest scheduled batch aggregate.';
  end if;
end $$;

-- equal scheduled priorities use the rule-key tie-breaker while Settings rejects duplicates
do $$
declare
  v_claim jsonb;
  v_next jsonb;
  v_complete jsonb;
  v_batch_id uuid;
  v_run_id uuid;
  v_rejected boolean := false;
  v_schedule jsonb;
  v_rules jsonb;
begin
  alter table public.financial_reconciliation_automatic_rule_configs
    drop constraint financial_reconciliation_automatic_rule_configs_priority_key;
  update public.financial_reconciliation_automatic_rule_configs
  set priority = 7
  where rule_key in (
    'financial_documents_cgd_bank_statement',
    'financial_documents_cgd_credit_card'
  );

  v_claim := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-03 01:00:00+00', 'smoke:scheduled-batch'
  );
  v_batch_id := (v_claim->>'batchId')::uuid;
  v_run_id := (v_claim#>>'{run,runId}')::uuid;
  if v_claim#>>'{run,batchRuleKey}' <> 'financial_documents_cgd_bank_statement'
    or (select batch.rule_snapshot#>>'{0,ruleKey}'
        from public.financial_reconciliation_automatic_batches batch
        where batch.id = v_batch_id) <> 'financial_documents_cgd_bank_statement'
    or (select batch.rule_snapshot#>>'{1,ruleKey}'
        from public.financial_reconciliation_automatic_batches batch
        where batch.id = v_batch_id) <> 'financial_documents_cgd_credit_card' then
    raise exception 'Equal scheduled priorities did not use the stable rule-key tie-breaker.';
  end if;

  update public.financial_reconciliation_automatic_runs
  set status = 'failed', finished_at = now(), updated_at = now()
  where id = v_run_id;
  v_next := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-03 01:01:00+00', 'smoke:scheduled-batch'
  );
  v_run_id := (v_next#>>'{run,runId}')::uuid;
  update public.financial_reconciliation_automatic_runs
  set status = 'failed', finished_at = now(), updated_at = now()
  where id = v_run_id;
  v_complete := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-03 01:02:00+00', 'smoke:scheduled-batch'
  );
  if v_complete->>'reason' <> 'batch_complete'
    or not exists (
      select 1 from public.financial_reconciliation_automatic_batches batch
      where batch.id = v_batch_id and batch.status = 'failed'
    ) then
    raise exception 'All-failed scheduled children did not produce a failed parent batch.';
  end if;

  update public.financial_reconciliation_automatic_rule_configs config
  set priority = case config.rule_key
    when 'financial_documents_cgd_bank_statement' then 1 else 2 end
  where config.rule_key in (
    'financial_documents_cgd_bank_statement',
    'financial_documents_cgd_credit_card'
  );
  alter table public.financial_reconciliation_automatic_rule_configs
    add constraint financial_reconciliation_automatic_rule_configs_priority_key
    unique (priority) deferrable initially deferred;

  select jsonb_build_object(
    'enabled', schedule.enabled,
    'time_of_day', to_char(schedule.time_of_day, 'HH24:MI'),
    'time_zone', schedule.time_zone
  ) into strict v_schedule
  from public.financial_reconciliation_automatic_schedule schedule
  where schedule.id = true;
  select jsonb_agg(jsonb_build_object(
    'rule_key', config.rule_key,
    'rule_version', config.rule_version,
    'enabled', config.enabled,
    'allow_manual_execution', config.allow_manual_execution,
    'include_in_scheduled_batch', config.include_in_scheduled_batch,
    'difference_allowed', to_char(config.difference_allowed, 'FM999999999990.00'),
    'max_difference_days', config.max_difference_days,
    'priority', 1
  ) order by config.rule_key) into v_rules
  from public.financial_reconciliation_automatic_rule_configs config;
  begin
    perform public.replace_financial_reconciliation_automation_settings(
      v_schedule, v_rules, 'smoke:duplicate-priority'
    );
  exception when raise_exception then
    if sqlerrm = 'Duplicate automatic rule priority.' then
      v_rejected := true;
    else
      raise;
    end if;
  end;
  if not v_rejected then
    raise exception 'Settings accepted duplicate automatic rule priorities.';
  end if;
end $$;

do $$
declare
  v_candidates jsonb;
  v_candidate_count integer;
  v_signature text;
begin
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, payment, amount, fat
  ) values
    ('76000000-0000-0000-0000-000000009001', date '2140-01-01', 'IDENTITY-BANK-9001', '', '', 'Banco', 301.00, 'S'),
    ('78000000-0000-0000-0000-000000009001', date '2140-01-01', 'IDENTITY-CARD-9001', '', '', 'Visa', 302.00, 'S');
  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  ) values (
    '77000000-0000-0000-0000-000000009001', 'smoke-final-identity-dispatch',
    'final-identity-bank-9001', date '2140-01-01', 'Payment IDENTITYBANK9001', -301.00
  );
  insert into public.import_cgd_cartao_credito (
    id, import_batch, row_key, data, descricao, debito
  ) values (
    '79000000-0000-0000-0000-000000009001', 'smoke-final-identity-dispatch',
    'final-identity-card-9001', date '2140-01-01', 'Payment IDENTITYCARD9001', 302.00
  );

  select candidates, candidate_count into strict v_candidates, v_candidate_count
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_bank_statement', 2, 0, 1,
    array['76000000-0000-0000-0000-000000009001'::uuid]
  );
  if v_candidate_count <> 1
    or jsonb_array_length(v_candidates) <> 1
    or v_candidates->0->>'sourceType' <> 'import_cgd_extrato_ordem'
    or v_candidates->0->>'sourceId' <> '77000000-0000-0000-0000-000000009001'
    or v_candidates#>>'{0,evidence,documentNumber,matched}' <> 'true' then
    raise exception 'The Banco identity adapter no longer dispatches with its executable candidate evidence: %', v_candidates;
  end if;

  select candidates, candidate_count into strict v_candidates, v_candidate_count
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_credit_card', 1, 0, 1,
    array['78000000-0000-0000-0000-000000009001'::uuid]
  );
  if v_candidate_count <> 1
    or jsonb_array_length(v_candidates) <> 1
    or v_candidates->0->>'sourceType' <> 'import_cgd_cartao_credito'
    or v_candidates->0->>'sourceId' <> '79000000-0000-0000-0000-000000009001'
    or v_candidates#>>'{0,evidence,documentNumber,matched}' <> 'true' then
    raise exception 'The Visa identity adapter no longer dispatches with its executable candidate evidence: %', v_candidates;
  end if;

  foreach v_signature in array array[
    'public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer)',
    'public.execute_financial_reconciliation_automatic_proposal(uuid,text)'
  ] loop
    if not (
      select procedure.prosecdef
        and coalesce(procedure.proconfig, '{}'::text[]) @> array['search_path=public, pg_temp']
      from pg_proc procedure
      where procedure.oid = v_signature::regprocedure
    )
      or has_function_privilege('anon', v_signature, 'EXECUTE')
      or has_function_privilege('authenticated', v_signature, 'EXECUTE')
      or not has_function_privilege('service_role', v_signature, 'EXECUTE') then
      raise exception 'Automatic reconciliation function security changed for %.', v_signature;
    end if;
  end loop;
end $$;

rollback;
