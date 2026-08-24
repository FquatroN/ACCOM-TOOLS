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
  v_constraint_type "char";
  v_constraint_validated boolean;
  v_installed_definition text;
  v_expected_definition text;
begin
  create temporary table financial_reconciliation_amount_only_zero_check_expected (
    rule_key text,
    difference_allowed numeric,
    constraint financial_reconciliation_amount_only_zero_check_expected_check
      check (
        rule_key not in (
          'financial_documents_cgd_bank_statement_amount_only',
          'financial_documents_cgd_credit_card_amount_only'
        )
        or difference_allowed = 0
      )
  ) on commit drop;

  select regexp_replace(pg_get_constraintdef(constraint_row.oid, true), '\s+NOT VALID$', '')
  into strict v_expected_definition
  from pg_constraint constraint_row
  where constraint_row.conrelid = 'financial_reconciliation_amount_only_zero_check_expected'::regclass
    and constraint_row.conname = 'financial_reconciliation_amount_only_zero_check_expected_check';

  select
    constraint_row.contype,
    constraint_row.convalidated,
    regexp_replace(pg_get_constraintdef(constraint_row.oid, true), '\s+NOT VALID$', '')
  into v_constraint_type, v_constraint_validated, v_installed_definition
  from pg_constraint constraint_row
  where constraint_row.conrelid = 'public.financial_reconciliation_automatic_rule_configs'::regclass
    and constraint_row.conname = 'financial_reconciliation_automatic_rule_configs_amount_only_zero_check';

  drop table financial_reconciliation_amount_only_zero_check_expected;

  if v_constraint_type is distinct from 'c'
    or not coalesce(v_constraint_validated, false)
    or v_installed_definition is distinct from v_expected_definition then
    raise exception 'Amount-only migration reapply did not leave the exact validated fixed-zero constraint installed.';
  end if;
end $$;

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

  select coalesce(array_agg(page.id order by page.document_date, page.id), '{}'::uuid[])
  into v_bank_base_ids
  from public.financial_reconciliation_automatic_base_page(
    'financial_documents_cgd_bank_statement', 2, null, null, 25
  ) page
  where page.id between '60000000-0000-0000-0000-000000000001'::uuid
                    and '60000000-0000-0000-0000-000000000020'::uuid;
  if v_bank_base_ids is distinct from array[
      '60000000-0000-0000-0000-000000000001'::uuid,
      '60000000-0000-0000-0000-000000000009'::uuid,
      '60000000-0000-0000-0000-000000000020'::uuid
    ] or public.financial_reconciliation_automatic_base_count(
      'financial_documents_cgd_bank_statement', 2
    ) <> 3 then
    raise exception 'Banco identity base paging/counting no longer preserves legacy nullable-amount eligibility: %', v_bank_base_ids;
  end if;

  select coalesce(array_agg(page.id order by page.document_date, page.id), '{}'::uuid[])
  into v_card_base_ids
  from public.financial_reconciliation_automatic_base_page(
    'financial_documents_cgd_credit_card', 1, null, null, 25
  ) page
  where page.id between '62000000-0000-0000-0000-000000000001'::uuid
                    and '62000000-0000-0000-0000-000000000020'::uuid;
  if v_card_base_ids is distinct from array[
      '62000000-0000-0000-0000-000000000001'::uuid,
      '62000000-0000-0000-0000-000000000009'::uuid,
      '62000000-0000-0000-0000-000000000020'::uuid
    ] or public.financial_reconciliation_automatic_base_count(
      'financial_documents_cgd_credit_card', 1
    ) <> 3 then
    raise exception 'Visa identity base paging/counting no longer preserves legacy nullable-amount eligibility: %', v_card_base_ids;
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

create or replace function pg_temp.make_task4_amount_only_run(
  p_rule_key text,
  p_actor text,
  p_client_request_id uuid
)
returns uuid
language plpgsql
as $$
declare
  v_run jsonb;
  v_run_id uuid;
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
      raise exception 'Task 4 amount-only analysis did not finish.';
    end if;
    v_run := public.continue_financial_reconciliation_automatic_analysis(v_run_id, p_actor);
  end loop;
  return v_run_id;
end $$;

do $$
begin
  if has_function_privilege(
      'anon',
      'public.financial_reconciliation_execute_identity_proposal(uuid,text)',
      'EXECUTE'
    )
    or has_function_privilege(
      'authenticated',
      'public.financial_reconciliation_execute_identity_proposal(uuid,text)',
      'EXECUTE'
    )
    or has_function_privilege(
      'service_role',
      'public.financial_reconciliation_execute_identity_proposal(uuid,text)',
      'EXECUTE'
    )
    or not (
      select procedure.prosecdef
        and coalesce(procedure.proconfig, '{}'::text[]) @> array['search_path=public, pg_temp']
      from pg_proc procedure
      where procedure.oid =
        'public.financial_reconciliation_execute_identity_proposal(uuid,text)'::regprocedure
    ) then
    raise exception 'The preserved identity execution body is directly exposed or unsafe.';
  end if;
end $$;

-- amount-only execution under the normal three-rule source catalog creates one
-- exact reconciliation with immutable evidence and preserves the full snapshot
-- amount-only repeated execution is idempotent
do $$
declare
  v_kind text;
  v_index integer := 0;
  v_rule_key text;
  v_source_type text;
  v_payment text;
  v_document_id uuid;
  v_destination_id uuid;
  v_run_id uuid;
  v_proposal_id uuid;
  v_reconciliation_id uuid;
  v_result jsonb;
  v_repeated jsonb;
  v_rule_snapshot jsonb;
  v_base_snapshot jsonb;
  v_destination_snapshots jsonb;
  v_evidence jsonb;
  v_signature text;
  v_item_count integer;
  v_audit_count integer;
  v_expected_trigger text;
  v_claim jsonb;
  v_run jsonb;
  v_guard integer;
  v_expected_matching_source_rules jsonb := '[
    {"sourceType":"import_cgd_cartao_credito","operator":"+"},
    {"sourceType":"import_cgd_extrato_ordem","operator":"+"},
    {"sourceType":"import_fdm_accounts","operator":"+"}
  ]'::jsonb;
begin
  if (select coalesce(jsonb_agg(jsonb_build_object(
        'sourceType', source_rule.matching_source_type,
        'operator', source_rule.operator
      ) order by source_rule.matching_source_type), '[]'::jsonb)
      from public.financial_reconciliation_source_rules source_rule
      where source_rule.base_source_type = 'financial_documents')
      is distinct from v_expected_matching_source_rules then
    raise exception 'Amount-only success fixture does not have the normal three-rule source catalog.';
  end if;

  foreach v_kind in array array['bank', 'card'] loop
    v_index := v_index + 1;
    v_rule_key := case v_kind
      when 'bank' then 'financial_documents_cgd_bank_statement_amount_only'
      else 'financial_documents_cgd_credit_card_amount_only' end;
    v_source_type := case v_kind
      when 'bank' then 'import_cgd_extrato_ordem'
      else 'import_cgd_cartao_credito' end;
    v_payment := case v_kind when 'bank' then 'Banco' else 'Visa' end;
    v_document_id := ('8a000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;
    v_destination_id := ('8b000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;

    update public.financial_reconciliation_source_rules
    set operator = '+'
    where base_source_type = 'financial_documents'
      and matching_source_type = v_source_type;
    insert into public.financial_documents (
      id, document_date, doc_number, description, supplier_name, payment, amount, fat
    ) values (
      v_document_id, date '2160-01-10' + v_index,
      'AMOUNT-EXEC-' || v_kind, 'identity deliberately unrelated', '',
      v_payment, 811 + v_index, 'S'
    );
    if v_kind = 'bank' then
      insert into public.import_cgd_extrato_ordem (
        id, import_batch, row_key, data, descritivo, montante
      ) values (
        v_destination_id, 'smoke-amount-task4', 'amount-exec-bank',
        date '2160-01-10' + v_index, 'unrelated bank text', -(811 + v_index)
      );
    else
      insert into public.import_cgd_cartao_credito (
        id, import_batch, row_key, data, descricao, debito
      ) values (
        v_destination_id, 'smoke-amount-task4', 'amount-exec-card',
        date '2160-01-10' + v_index, 'unrelated card text', 811 + v_index
      );
    end if;

    if v_kind = 'bank' then
      v_run_id := pg_temp.make_task4_amount_only_run(
        v_rule_key, 'smoke:amount-execute-' || v_kind,
        ('8c000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid
      );
    else
      update public.financial_reconciliation_automatic_rule_configs
      set include_in_scheduled_batch = false;
      update public.financial_reconciliation_automatic_rule_configs
      set enabled = true,
          allow_manual_execution = false,
          include_in_scheduled_batch = true,
          difference_allowed = 0,
          max_difference_days = 10
      where rule_key = v_rule_key;
      update public.financial_reconciliation_automatic_schedule
      set enabled = true,
          time_of_day = time '00:00',
          time_zone = 'Europe/Lisbon'
      where id = true;
      v_claim := public.claim_financial_reconciliation_automatic_schedule(
        '2080-01-12 01:00:00+00', 'smoke:amount-execute-card'
      );
      if not coalesce((v_claim->>'claimed')::boolean, false)
        or coalesce((v_claim->>'resumed')::boolean, true)
        or v_claim#>>'{run,batchRuleKey}' <> v_rule_key then
        raise exception 'Amount-only card fixture did not claim an authentic scheduled child: %', v_claim;
      end if;
      v_run := v_claim->'run';
      v_run_id := (v_run->>'runId')::uuid;
      v_guard := 0;
      while not coalesce((v_run->>'analysisComplete')::boolean, false) loop
        v_guard := v_guard + 1;
        if v_guard > 100 then
          raise exception 'Scheduled amount-only card analysis did not finish.';
        end if;
        v_run := public.continue_financial_reconciliation_automatic_analysis(
          v_run_id, 'smoke:amount-execute-card'
        );
      end loop;
    end if;
    select proposal.id, proposal.base_snapshot, proposal.items,
           proposal.evidence, proposal.signature, run.definition_config_snapshot->0
    into strict v_proposal_id, v_base_snapshot, v_destination_snapshots,
                v_evidence, v_signature, v_rule_snapshot
    from public.financial_reconciliation_automatic_proposals proposal
    join public.financial_reconciliation_automatic_runs run on run.id = proposal.run_id
    where proposal.run_id = v_run_id
      and proposal.base_source_id = v_document_id
      and proposal.status = 'proposed';
    if jsonb_array_length(v_destination_snapshots) <> 1
      or v_destination_snapshots#>>'{0,sourceId}' <> v_destination_id::text then
      raise exception 'Amount-only % success fixture was not exactly one-to-one.', v_kind;
    end if;

    v_expected_trigger := case v_kind when 'bank' then 'manual' else 'scheduled' end;

    v_result := public.execute_financial_reconciliation_automatic_proposal(
      v_proposal_id, 'smoke:amount-execute-' || v_kind
    );
    v_reconciliation_id := (v_result->>'reconciliationId')::uuid;
    if v_result is distinct from jsonb_build_object(
        'proposalId', v_proposal_id,
        'runId', v_run_id,
        'status', 'completed',
        'reconciliationId', v_reconciliation_id
      )
      or v_reconciliation_id is null
      or not exists (
        select 1
        from public.financial_reconciliations reconciliation
        where reconciliation.id = v_reconciliation_id
          and reconciliation.status = 'complete'
          and reconciliation.completion_type = 'normal'
          and reconciliation.difference_amount = 0
          and reconciliation.origin = 'automatic'
          and reconciliation.automatic_trigger = v_expected_trigger
          and reconciliation.automatic_rule_key = v_rule_key
          and reconciliation.automatic_rule_version = 1
          and reconciliation.automatic_run_id = v_run_id
          and reconciliation.automatic_proposal_id = v_proposal_id
          and reconciliation.matching_source_rules = v_expected_matching_source_rules
      )
      or (select count(*) from public.financial_reconciliation_items item
          where item.reconciliation_id = v_reconciliation_id) <> 2
      or (select count(*) from public.financial_reconciliation_items item
          where item.reconciliation_id = v_reconciliation_id
            and item.source_type = 'financial_documents'
            and item.source_id = v_document_id) <> 1
      or (select count(*) from public.financial_reconciliation_items item
          where item.reconciliation_id = v_reconciliation_id
            and item.source_type = v_source_type
            and item.source_id = v_destination_id) <> 1 then
      raise exception 'Amount-only % execution did not create one exact automatic reconciliation.', v_kind;
    end if;

    if not exists (
      select 1
      from public.financial_reconciliation_audit audit
      where audit.reconciliation_id = v_reconciliation_id
        and audit.action = 'automatic_complete'
        and audit.metadata @> jsonb_build_object(
          'ruleSnapshot', jsonb_build_object(
            'ruleKey', v_rule_key,
            'ruleVersion', 1,
            'definition', v_rule_snapshot->'definition'
          ),
          'configSnapshot', jsonb_build_object(
            'differenceAllowed', 0,
            'maxDifferenceDays', 10,
            'priority', (v_rule_snapshot->>'priority')::integer
          ),
          'operatorSnapshot', jsonb_build_object(v_source_type, '+'),
          'baseSnapshot', v_base_snapshot,
          'destinationSnapshots', v_destination_snapshots,
          'identityEvidence', v_evidence,
          'proposalSignature', v_signature,
          'trigger', v_expected_trigger,
          'runId', v_run_id,
          'proposalId', v_proposal_id,
          'tolerance', 0,
          'calculatedDifference', 0
        )
    ) or not exists (
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
      raise exception 'Amount-only % immutable audit or proposal evidence changed.', v_kind;
    end if;

    select count(*) into v_item_count
    from public.financial_reconciliation_items item
    where item.reconciliation_id = v_reconciliation_id;
    select count(*) into v_audit_count
    from public.financial_reconciliation_audit audit
    where audit.reconciliation_id = v_reconciliation_id;
    v_repeated := public.execute_financial_reconciliation_automatic_proposal(
      v_proposal_id, 'smoke:amount-execute-' || v_kind
    );
    if v_repeated is distinct from v_result
      or (select count(*) from public.financial_reconciliations reconciliation
          where reconciliation.automatic_proposal_id = v_proposal_id) <> 1
      or (select count(*) from public.financial_reconciliation_items item
          where item.reconciliation_id = v_reconciliation_id) <> v_item_count
      or (select count(*) from public.financial_reconciliation_audit audit
          where audit.reconciliation_id = v_reconciliation_id) <> v_audit_count then
      raise exception 'Repeated amount-only % execution was not idempotent.', v_kind;
    end if;
    if v_kind = 'card' then
      v_run := public.finish_financial_reconciliation_automatic_run(v_run_id);
      if v_run->>'status' <> 'completed' then
        raise exception 'Scheduled amount-only card child did not finish cleanly: %', v_run;
      end if;
    end if;
  end loop;
end $$;

-- oversized integer snapshot fields become sanitized stale results without writes
do $$
declare
  v_fields text[] := array['ruleVersion', 'maxDifferenceDays', 'priority'];
  v_field text;
  v_index integer := 0;
  v_document_id uuid;
  v_destination_id uuid;
  v_run_id uuid;
  v_proposal_id uuid;
  v_result jsonb;
begin
  foreach v_field in array v_fields loop
    v_index := v_index + 1;
    v_document_id := ('a1000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;
    v_destination_id := ('a2000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;

    insert into public.financial_documents (
      id, document_date, doc_number, description, supplier_name, payment, amount, fat
    ) values (
      v_document_id, date '2160-06-01' + v_index,
      'AMOUNT-OVERSIZED-' || v_field, '', '', 'Banco', 98700 + v_index, 'S'
    );
    insert into public.import_cgd_extrato_ordem (
      id, import_batch, row_key, data, descritivo, montante
    ) values (
      v_destination_id, 'smoke-amount-oversized', 'amount-oversized-' || v_index,
      date '2160-06-01' + v_index, '', -(98700 + v_index)
    );

    v_run_id := pg_temp.make_task4_amount_only_run(
      'financial_documents_cgd_bank_statement_amount_only',
      'smoke:amount-oversized-' || v_field,
      ('a3000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid
    );
    select proposal.id into strict v_proposal_id
    from public.financial_reconciliation_automatic_proposals proposal
    where proposal.run_id = v_run_id
      and proposal.base_source_id = v_document_id
      and proposal.status = 'proposed';

    update public.financial_reconciliation_automatic_runs
    set definition_config_snapshot = jsonb_set(
      definition_config_snapshot,
      array['0', v_field],
      '2147483648'::jsonb
    )
    where id = v_run_id;

    v_result := public.execute_financial_reconciliation_automatic_proposal(
      v_proposal_id, 'smoke:amount-oversized-' || v_field
    );
    if v_result is distinct from jsonb_build_object(
        'proposalId', v_proposal_id,
        'runId', v_run_id,
        'status', 'stale',
        'reason', 'rule_snapshot_changed'
      )
      or not exists (
        select 1
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.id = v_proposal_id
          and proposal.status = 'stale'
          and proposal.reason = 'rule_snapshot_changed'
          and proposal.reconciliation_id is null
          and proposal.completed_at is null
          and proposal.error = ''
          and proposal.error_detail = ''
      )
      or exists (
        select 1
        from public.financial_reconciliations reconciliation
        where reconciliation.automatic_proposal_id = v_proposal_id
      )
      or exists (
        select 1
        from public.financial_reconciliation_audit audit
        where audit.metadata->>'proposalId' = v_proposal_id::text
      ) then
      raise exception 'Oversized % snapshot was not sanitized as stale: %',
        v_field, v_result;
    end if;
  end loop;
end $$;

-- amount-only execution fails closed on rule config source item and lock drift
do $$
declare
  v_kinds text[] := array[
    'payment', 'base_amount', 'base_date', 'destination_amount',
    'destination_date', 'window', 'definition', 'rule_version',
    'operator', 'item_count', 'snapshot_tolerance', 'locked'
  ];
  v_source_kind text;
  v_kind text;
  v_index integer := 0;
  v_rule_key text;
  v_source_type text;
  v_payment text;
  v_document_id uuid;
  v_destination_id uuid;
  v_run_id uuid;
  v_proposal_id uuid;
  v_lock_id uuid;
  v_original_definition jsonb;
  v_result jsonb;
begin
  foreach v_source_kind in array array['bank', 'card'] loop
    v_rule_key := case v_source_kind
      when 'bank' then 'financial_documents_cgd_bank_statement_amount_only'
      else 'financial_documents_cgd_credit_card_amount_only' end;
    v_source_type := case v_source_kind
      when 'bank' then 'import_cgd_extrato_ordem'
      else 'import_cgd_cartao_credito' end;
    v_payment := case v_source_kind when 'bank' then 'Banco' else 'Visa' end;
    select definition into strict v_original_definition
    from public.financial_reconciliation_automatic_rule_definitions
    where rule_key = v_rule_key and version = 1;

    foreach v_kind in array v_kinds loop
      v_index := v_index + 1;
      v_document_id := ('8d000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;
      v_destination_id := ('8e000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;
      v_lock_id := null;

      update public.financial_reconciliation_source_rules
      set operator = '+'
      where base_source_type = 'financial_documents'
        and matching_source_type = v_source_type;
      insert into public.financial_documents (
        id, document_date, doc_number, description, supplier_name, payment, amount, fat
      ) values (
        v_document_id, date '2161-01-01' + v_index,
        'AMOUNT-STALE-' || v_index, '', '', v_payment, 830 + v_index, 'S'
      );
      if v_source_kind = 'bank' then
        insert into public.import_cgd_extrato_ordem (
          id, import_batch, row_key, data, descritivo, montante
        ) values (
          v_destination_id, 'smoke-amount-stale', 'amount-stale-' || v_index,
          date '2161-01-11' + v_index, '', -(830 + v_index)
        );
      else
        insert into public.import_cgd_cartao_credito (
          id, import_batch, row_key, data, descricao, debito
        ) values (
          v_destination_id, 'smoke-amount-stale', 'amount-stale-' || v_index,
          date '2161-01-11' + v_index, '', 830 + v_index
        );
      end if;

      v_run_id := pg_temp.make_task4_amount_only_run(
        v_rule_key, 'smoke:amount-stale-' || v_source_kind || '-' || v_kind,
        ('8f000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid
      );
      select proposal.id into strict v_proposal_id
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = v_run_id
        and proposal.base_source_id = v_document_id
        and proposal.status = 'proposed';

      if v_kind = 'payment' then
        update public.financial_documents set payment = upper(v_payment) where id = v_document_id;
      elsif v_kind = 'base_amount' then
        update public.financial_documents set amount = amount + 1 where id = v_document_id;
      elsif v_kind = 'base_date' then
        update public.financial_documents set document_date = document_date + 1 where id = v_document_id;
      elsif v_kind = 'destination_amount' and v_source_kind = 'bank' then
        update public.import_cgd_extrato_ordem set montante = montante - 1 where id = v_destination_id;
      elsif v_kind = 'destination_amount' then
        update public.import_cgd_cartao_credito set debito = debito + 1 where id = v_destination_id;
      elsif v_kind = 'destination_date' and v_source_kind = 'bank' then
        update public.import_cgd_extrato_ordem set data = data + 1 where id = v_destination_id;
      elsif v_kind = 'destination_date' then
        update public.import_cgd_cartao_credito set data = data + 1 where id = v_destination_id;
      elsif v_kind = 'window' then
        update public.financial_reconciliation_automatic_rule_configs
        set max_difference_days = 9 where rule_key = v_rule_key;
      elsif v_kind = 'definition' then
        update public.financial_reconciliation_automatic_rule_definitions
        set definition = definition || '{"task4ExecutionDrift":true}'::jsonb
        where rule_key = v_rule_key and version = 1;
      elsif v_kind = 'rule_version' then
        update public.financial_reconciliation_automatic_runs
        set definition_config_snapshot = jsonb_set(
          definition_config_snapshot, '{0,ruleVersion}', '2'::jsonb
        ) where id = v_run_id;
      elsif v_kind = 'operator' then
        update public.financial_reconciliation_source_rules
        set operator = '-'
        where base_source_type = 'financial_documents'
          and matching_source_type = v_source_type;
      elsif v_kind = 'item_count' then
        update public.financial_reconciliation_automatic_proposals
        set items = items || items where id = v_proposal_id;
      elsif v_kind = 'snapshot_tolerance' then
        update public.financial_reconciliation_automatic_runs
        set definition_config_snapshot = jsonb_set(
          definition_config_snapshot, '{0,differenceAllowed}', '1'::jsonb
        ) where id = v_run_id;
        update public.financial_reconciliation_automatic_proposals
        set allowed_difference = 1 where id = v_proposal_id;
      elsif v_kind = 'locked' then
        v_result := public.financial_reconciliation_action(
          'start', 'smoke:amount-destination-lock', null,
          v_source_type, v_destination_id, null
        );
        v_lock_id := (v_result#>>'{reconciliation,id}')::uuid;
      end if;

      v_result := public.execute_financial_reconciliation_automatic_proposal(
        v_proposal_id, 'smoke:amount-stale-' || v_source_kind || '-' || v_kind
      );
      if v_result->>'status' <> 'stale'
        or v_result->>'reason' not in (
          'rule_version_changed', 'rule_snapshot_changed', 'operator_changed',
          'tolerance_changed', 'source_snapshot_changed', 'combination_changed',
          'proposal_evidence_changed'
        )
        or jsonb_object_length(v_result) <> 4
        or v_result ?| array['error', 'errorDetail', 'detail', 'message']
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
        raise exception 'Amount-only % stale path % did not fail closed: %',
          v_source_kind, v_kind, v_result;
      end if;

      if v_kind = 'window' then
        update public.financial_reconciliation_automatic_rule_configs
        set max_difference_days = 10 where rule_key = v_rule_key;
      elsif v_kind = 'definition' then
        update public.financial_reconciliation_automatic_rule_definitions
        set definition = v_original_definition
        where rule_key = v_rule_key and version = 1;
      elsif v_kind = 'operator' then
        update public.financial_reconciliation_source_rules
        set operator = '+'
        where base_source_type = 'financial_documents'
          and matching_source_type = v_source_type;
      elsif v_kind = 'locked' then
        perform public.financial_reconciliation_action(
          'delete', 'smoke:amount-destination-lock', v_lock_id, null, null, null
        );
      end if;
    end loop;
  end loop;
end $$;

-- competing amount-only proposals cannot consume the same destination
do $$
declare
  v_kind text;
  v_index integer := 0;
  v_rule_key text;
  v_source_type text;
  v_payment text;
  v_first_document_id uuid;
  v_second_document_id uuid;
  v_destination_id uuid;
  v_first_run_id uuid;
  v_second_run_id uuid;
  v_first_proposal_id uuid;
  v_second_proposal_id uuid;
  v_first_result jsonb;
  v_second_result jsonb;
begin
  foreach v_kind in array array['bank', 'card'] loop
    v_index := v_index + 1;
    v_rule_key := case v_kind
      when 'bank' then 'financial_documents_cgd_bank_statement_amount_only'
      else 'financial_documents_cgd_credit_card_amount_only' end;
    v_source_type := case v_kind
      when 'bank' then 'import_cgd_extrato_ordem'
      else 'import_cgd_cartao_credito' end;
    v_payment := case v_kind when 'bank' then 'Banco' else 'Visa' end;
    v_first_document_id := ('9a000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;
    v_second_document_id := ('9b000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;
    v_destination_id := ('9c000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;

    update public.financial_reconciliation_source_rules
    set operator = '+'
    where base_source_type = 'financial_documents'
      and matching_source_type = v_source_type;
    insert into public.financial_documents (
      id, document_date, doc_number, description, supplier_name, payment, amount, fat
    ) values (
      v_first_document_id, date '2162-01-10' + v_index,
      'AMOUNT-COMPETE-A-' || v_kind, '', '', v_payment, 920 + v_index, 'S'
    );
    if v_kind = 'bank' then
      insert into public.import_cgd_extrato_ordem (
        id, import_batch, row_key, data, descritivo, montante
      ) values (
        v_destination_id, 'smoke-amount-compete', 'amount-compete-bank',
        date '2162-01-10' + v_index, '', -(920 + v_index)
      );
    else
      insert into public.import_cgd_cartao_credito (
        id, import_batch, row_key, data, descricao, debito
      ) values (
        v_destination_id, 'smoke-amount-compete', 'amount-compete-card',
        date '2162-01-10' + v_index, '', 920 + v_index
      );
    end if;

    v_first_run_id := pg_temp.make_task4_amount_only_run(
      v_rule_key, 'smoke:amount-compete-first-' || v_kind,
      ('9d000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid
    );
    select id into strict v_first_proposal_id
    from public.financial_reconciliation_automatic_proposals
    where run_id = v_first_run_id
      and base_source_id = v_first_document_id
      and status = 'proposed';

    update public.financial_documents
    set payment = 'smoke:hidden-from-second-analysis'
    where id = v_first_document_id;
    insert into public.financial_documents (
      id, document_date, doc_number, description, supplier_name, payment, amount, fat
    ) values (
      v_second_document_id, date '2162-01-10' + v_index,
      'AMOUNT-COMPETE-B-' || v_kind, '', '', v_payment, 920 + v_index, 'S'
    );
    v_second_run_id := pg_temp.make_task4_amount_only_run(
      v_rule_key, 'smoke:amount-compete-second-' || v_kind,
      ('9e000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid
    );
    select id into strict v_second_proposal_id
    from public.financial_reconciliation_automatic_proposals
    where run_id = v_second_run_id
      and base_source_id = v_second_document_id
      and status = 'proposed';
    update public.financial_documents set payment = v_payment where id = v_first_document_id;

    v_first_result := public.execute_financial_reconciliation_automatic_proposal(
      v_first_proposal_id, 'smoke:amount-compete-first-' || v_kind
    );
    v_second_result := public.execute_financial_reconciliation_automatic_proposal(
      v_second_proposal_id, 'smoke:amount-compete-second-' || v_kind
    );
    if v_first_result->>'status' <> 'completed'
      or v_second_result->>'status' <> 'stale'
      or jsonb_object_length(v_second_result) <> 4
      or (select count(*)
          from public.financial_reconciliation_items item
          where item.source_type = v_source_type
            and item.source_id = v_destination_id) <> 1
      or (select count(*)
          from public.financial_reconciliations reconciliation
          where reconciliation.automatic_proposal_id in (
            v_first_proposal_id, v_second_proposal_id
          )) <> 1 then
      raise exception 'Competing amount-only % proposals consumed one destination more than once.', v_kind;
    end if;
  end loop;
end $$;

-- amount-only ambiguous and candidate-limit proposals cannot execute
do $$
declare
  v_kind text;
  v_case text;
  v_index integer := 0;
  v_destination_index integer;
  v_destination_count integer;
  v_rule_key text;
  v_payment text;
  v_document_id uuid;
  v_destination_id uuid;
  v_run_id uuid;
  v_proposal_id uuid;
  v_status text;
  v_reason text;
  v_rejected boolean;
begin
  foreach v_kind in array array['bank', 'card'] loop
    v_rule_key := case v_kind
      when 'bank' then 'financial_documents_cgd_bank_statement_amount_only'
      else 'financial_documents_cgd_credit_card_amount_only' end;
    v_payment := case v_kind when 'bank' then 'Banco' else 'Visa' end;
    foreach v_case in array array['ambiguous', 'candidate_limit'] loop
      v_index := v_index + 1;
      v_destination_count := case v_case when 'ambiguous' then 2 else 13 end;
      v_document_id := ('aa000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;
      insert into public.financial_documents (
        id, document_date, doc_number, description, supplier_name, payment, amount, fat
      ) values (
        v_document_id, date '2163-01-01' + v_index,
        'AMOUNT-NONEXEC-' || v_kind || '-' || v_case,
        '', '', v_payment, 950 + v_index, 'S'
      );
      for v_destination_index in 1..v_destination_count loop
        v_destination_id := (
          'ab000000-0000-' || lpad(v_index::text, 4, '0') || '-0000-'
          || lpad(v_destination_index::text, 12, '0')
        )::uuid;
        if v_kind = 'bank' then
          insert into public.import_cgd_extrato_ordem (
            id, import_batch, row_key, data, descritivo, montante
          ) values (
            v_destination_id, 'smoke-amount-nonexec',
            'amount-nonexec-' || v_kind || '-' || v_case || '-' || v_destination_index,
            date '2163-01-01' + v_index, '', -(950 + v_index)
          );
        else
          insert into public.import_cgd_cartao_credito (
            id, import_batch, row_key, data, descricao, debito
          ) values (
            v_destination_id, 'smoke-amount-nonexec',
            'amount-nonexec-' || v_kind || '-' || v_case || '-' || v_destination_index,
            date '2163-01-01' + v_index, '', 950 + v_index
          );
        end if;
      end loop;

      v_run_id := pg_temp.make_task4_amount_only_run(
        v_rule_key, 'smoke:amount-nonexec-' || v_kind || '-' || v_case,
        ('ac000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid
      );
      select proposal.id, proposal.status, proposal.reason
      into strict v_proposal_id, v_status, v_reason
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = v_run_id
        and proposal.base_source_id = v_document_id;
      if v_status <> 'ambiguous'
        or v_reason <> case v_case
          when 'ambiguous' then 'multiple_combinations'
          else 'candidate_limit' end then
        raise exception 'Amount-only % % fixture had unexpected lifecycle %/%',
          v_kind, v_case, v_status, v_reason;
      end if;

      v_rejected := false;
      begin
        perform public.execute_financial_reconciliation_automatic_proposal(
          v_proposal_id, 'smoke:amount-nonexec-' || v_kind || '-' || v_case
        );
      exception when others then
        if sqlerrm like 'Automation proposal with status ambiguous cannot be executed.%' then
          v_rejected := true;
        else
          raise;
        end if;
      end;
      if not v_rejected or exists (
        select 1
        from public.financial_reconciliations reconciliation
        where reconciliation.automatic_proposal_id = v_proposal_id
      ) then
        raise exception 'Amount-only % % proposal was executable.', v_kind, v_case;
      end if;
    end loop;
  end loop;
end $$;

-- amount-only post-write verification rolls back a changed completion snapshot
create or replace function pg_temp.tamper_task4_amount_only_completion()
returns trigger language plpgsql as $$
begin
  if new.status = 'complete'
    and new.automatic_rule_key = 'financial_documents_cgd_bank_statement_amount_only'
    and exists (
      select 1
      from public.financial_reconciliation_automatic_runs run
      where run.id = new.automatic_run_id
        and run.actor like 'smoke:amount-postwrite-%'
    ) then
    new.matching_source_rules := new.matching_source_rules || jsonb_build_array(
      jsonb_build_object('sourceType', 'import_fdm_accounts', 'operator', '+')
    );
  end if;
  return new;
end $$;

create or replace function pg_temp.tamper_task4_amount_only_item_snapshot()
returns trigger language plpgsql as $$
begin
  if new.action = 'automatic_complete'
    and new.actor = 'smoke:amount-postwrite-card' then
    update public.financial_reconciliation_items
    set amount_snapshot = amount_snapshot + 1
    where reconciliation_id = new.reconciliation_id
      and source_type = 'import_cgd_cartao_credito';
  end if;
  return new;
end $$;

create trigger reconciliation_task4_amount_only_postwrite_smoke
  before update on public.financial_reconciliations
  for each row execute function pg_temp.tamper_task4_amount_only_completion();

create trigger reconciliation_task4_amount_only_postwrite_item_smoke
  before insert on public.financial_reconciliation_audit
  for each row execute function pg_temp.tamper_task4_amount_only_item_snapshot();

do $$
declare
  v_kind text;
  v_index integer := 0;
  v_rule_key text;
  v_source_type text;
  v_payment text;
  v_document_id uuid;
  v_destination_id uuid;
  v_run_id uuid;
  v_proposal_id uuid;
  v_result jsonb;
begin
  foreach v_kind in array array['bank', 'card'] loop
    v_index := v_index + 1;
    v_rule_key := case v_kind
      when 'bank' then 'financial_documents_cgd_bank_statement_amount_only'
      else 'financial_documents_cgd_credit_card_amount_only' end;
    v_source_type := case v_kind
      when 'bank' then 'import_cgd_extrato_ordem'
      else 'import_cgd_cartao_credito' end;
    v_payment := case v_kind when 'bank' then 'Banco' else 'Visa' end;
    v_document_id := ('c0000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;
    v_destination_id := ('c1000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;

    insert into public.financial_documents (
      id, document_date, doc_number, description, supplier_name, payment, amount, fat
    ) values (
      v_document_id, date '2163-06-10' + v_index,
      'AMOUNT-POSTWRITE-' || v_kind, '', '', v_payment, 970 + v_index, 'S'
    );
    if v_kind = 'bank' then
      insert into public.import_cgd_extrato_ordem (
        id, import_batch, row_key, data, descritivo, montante
      ) values (
        v_destination_id, 'smoke-amount-postwrite', 'amount-postwrite-bank',
        date '2163-06-10' + v_index, '', -(970 + v_index)
      );
    else
      insert into public.import_cgd_cartao_credito (
        id, import_batch, row_key, data, descricao, debito
      ) values (
        v_destination_id, 'smoke-amount-postwrite', 'amount-postwrite-card',
        date '2163-06-10' + v_index, '', 970 + v_index
      );
    end if;

    v_run_id := pg_temp.make_task4_amount_only_run(
      v_rule_key, 'smoke:amount-postwrite-' || v_kind,
      ('c2000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid
    );
    select id into strict v_proposal_id
    from public.financial_reconciliation_automatic_proposals
    where run_id = v_run_id
      and base_source_id = v_document_id
      and status = 'proposed';
    v_result := public.execute_financial_reconciliation_automatic_proposal(
      v_proposal_id, 'smoke:amount-postwrite-' || v_kind
    );
    if v_result is distinct from jsonb_build_object(
        'proposalId', v_proposal_id,
        'runId', v_run_id,
        'status', 'stale',
        'reason', 'source_snapshot_changed'
      )
      or exists (
        select 1 from public.financial_reconciliations reconciliation
        where reconciliation.automatic_proposal_id = v_proposal_id
      )
      or exists (
        select 1 from public.financial_reconciliation_items item
        where (item.source_type = 'financial_documents' and item.source_id = v_document_id)
           or (item.source_type = v_source_type and item.source_id = v_destination_id)
      )
      or not exists (
        select 1 from public.financial_reconciliation_automatic_proposals proposal
        where proposal.id = v_proposal_id
          and proposal.status = 'stale'
          and proposal.reason = 'source_snapshot_changed'
          and proposal.reconciliation_id is null
          and proposal.error = ''
          and proposal.error_detail = ''
      ) then
      raise exception 'Amount-only % post-write verification did not roll back changed state.', v_kind;
    end if;
  end loop;
end $$;

drop trigger reconciliation_task4_amount_only_postwrite_smoke
  on public.financial_reconciliations;
drop trigger reconciliation_task4_amount_only_postwrite_item_smoke
  on public.financial_reconciliation_audit;

-- amount-only post-write failure rolls back and later proposals remain isolated
-- amount-only failure results are sanitized
create or replace function pg_temp.reject_task4_amount_only_audit()
returns trigger language plpgsql as $$
begin
  if new.action = 'automatic_complete'
    and new.actor like 'smoke:amount-rollback-%' then
    raise exception 'Smoke forced amount-only audit failure.';
  end if;
  return new;
end $$;

create trigger reconciliation_task4_amount_only_rollback_smoke
  before insert on public.financial_reconciliation_audit
  for each row execute function pg_temp.reject_task4_amount_only_audit();

do $$
declare
  v_kind text;
  v_index integer := 0;
  v_rule_key text;
  v_source_type text;
  v_payment text;
  v_document_id uuid;
  v_later_document_id uuid;
  v_destination_id uuid;
  v_later_destination_id uuid;
  v_run_id uuid;
  v_later_run_id uuid;
  v_proposal_id uuid;
  v_later_proposal_id uuid;
  v_result jsonb;
begin
  foreach v_kind in array array['bank', 'card'] loop
    v_index := v_index + 1;
    v_rule_key := case v_kind
      when 'bank' then 'financial_documents_cgd_bank_statement_amount_only'
      else 'financial_documents_cgd_credit_card_amount_only' end;
    v_source_type := case v_kind
      when 'bank' then 'import_cgd_extrato_ordem'
      else 'import_cgd_cartao_credito' end;
    v_payment := case v_kind when 'bank' then 'Banco' else 'Visa' end;
    v_document_id := ('bd000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;
    v_later_document_id := ('bd100000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;
    v_destination_id := ('be000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;
    v_later_destination_id := ('be100000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid;

    insert into public.financial_documents (
      id, document_date, doc_number, description, supplier_name, payment, amount, fat
    ) values
      (v_document_id, date '2164-01-10' + v_index,
       'AMOUNT-ROLLBACK-' || v_kind, '', '', v_payment, 980 + v_index, 'S'),
      (v_later_document_id, date '2164-02-10' + v_index,
       'AMOUNT-LATER-' || v_kind, '', '', v_payment, 990 + v_index, 'S');
    if v_kind = 'bank' then
      insert into public.import_cgd_extrato_ordem (
        id, import_batch, row_key, data, descritivo, montante
      ) values
        (v_destination_id, 'smoke-amount-rollback', 'amount-rollback-bank',
         date '2164-01-10' + v_index, '', -(980 + v_index)),
        (v_later_destination_id, 'smoke-amount-rollback', 'amount-later-bank',
         date '2164-02-10' + v_index, '', -(990 + v_index));
    else
      insert into public.import_cgd_cartao_credito (
        id, import_batch, row_key, data, descricao, debito
      ) values
        (v_destination_id, 'smoke-amount-rollback', 'amount-rollback-card',
         date '2164-01-10' + v_index, '', 980 + v_index),
        (v_later_destination_id, 'smoke-amount-rollback', 'amount-later-card',
         date '2164-02-10' + v_index, '', 990 + v_index);
    end if;

    v_run_id := pg_temp.make_task4_amount_only_run(
      v_rule_key, 'smoke:amount-rollback-run-' || v_kind,
      ('bf000000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid
    );
    select id into strict v_proposal_id
    from public.financial_reconciliation_automatic_proposals
    where run_id = v_run_id
      and base_source_id = v_document_id
      and status = 'proposed';
    v_result := public.execute_financial_reconciliation_automatic_proposal(
      v_proposal_id, 'smoke:amount-rollback-' || v_kind
    );
    if v_result is distinct from jsonb_build_object(
        'proposalId', v_proposal_id,
        'runId', v_run_id,
        'status', 'failed',
        'reason', 'execution_failed'
      )
      or exists (
        select 1 from public.financial_reconciliations reconciliation
        where reconciliation.automatic_proposal_id = v_proposal_id
      )
      or exists (
        select 1 from public.financial_reconciliation_items item
        where (item.source_type = 'financial_documents' and item.source_id = v_document_id)
           or (item.source_type = v_source_type and item.source_id = v_destination_id)
      )
      or not exists (
        select 1 from public.financial_reconciliation_automatic_proposals proposal
        where proposal.id = v_proposal_id
          and proposal.status = 'failed'
          and proposal.reason = 'execution_failed'
          and proposal.error = 'Automatic reconciliation execution failed.'
          and proposal.reconciliation_id is null
      ) then
      raise exception 'Amount-only % post-write failure was not isolated and sanitized.', v_kind;
    end if;

    v_later_run_id := pg_temp.make_task4_amount_only_run(
      v_rule_key, 'smoke:amount-later-run-' || v_kind,
      ('bf100000-0000-0000-0000-' || lpad(v_index::text, 12, '0'))::uuid
    );
    select id into strict v_later_proposal_id
    from public.financial_reconciliation_automatic_proposals
    where run_id = v_later_run_id
      and base_source_id = v_later_document_id
      and status = 'proposed';
    v_result := public.execute_financial_reconciliation_automatic_proposal(
      v_later_proposal_id, 'smoke:amount-later-' || v_kind
    );
    if v_result->>'status' <> 'completed' then
      raise exception 'Later amount-only % proposal was blocked by a rolled-back failure.', v_kind;
    end if;
  end loop;
end $$;

drop trigger reconciliation_task4_amount_only_rollback_smoke
  on public.financial_reconciliation_audit;

update public.financial_documents
set payment = 'smoke:task4-amount-execution-covered'
where doc_number like 'AMOUNT-%';

-- amount-only candidate-limit overlap remains authoritative beyond bounded evidence
do $$
declare
  v_candidates jsonb;
  v_candidate_count integer;
  v_run jsonb;
  v_run_id uuid;
begin
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, payment, amount, fat
  ) values
    ('80000000-0000-0000-0000-000000000001', date '2136-02-01', 'BANK-HIDDEN-OVERLAP-LIMIT', '', '', 'Banco', 206.00, 'S'),
    ('80000000-0000-0000-0000-000000000002', date '2136-02-03', 'BANK-HIDDEN-OVERLAP-ONLY', '', '', 'Banco', 206.00, 'S'),
    ('82000000-0000-0000-0000-000000000001', date '2136-02-01', 'CARD-HIDDEN-OVERLAP-LIMIT', '', '', 'Visa', 206.00, 'S'),
    ('82000000-0000-0000-0000-000000000002', date '2136-02-03', 'CARD-HIDDEN-OVERLAP-ONLY', '', '', 'Visa', 206.00, 'S');

  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  )
  select
    ('81000000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
    'smoke-hidden-overlap', 'bank-hidden-overlap-private-' || series,
    date '2136-01-31', '', -206.00
  from generate_series(1, 12) series;
  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  ) values (
    '81000000-0000-0000-0000-000000000999', 'smoke-hidden-overlap',
    'bank-hidden-overlap-shared', date '2136-02-02', '', -206.00
  );

  insert into public.import_cgd_cartao_credito (
    id, import_batch, row_key, data, descricao, debito
  )
  select
    ('83000000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
    'smoke-hidden-overlap', 'card-hidden-overlap-private-' || series,
    date '2136-01-31', '', 206.00
  from generate_series(1, 12) series;
  insert into public.import_cgd_cartao_credito (
    id, import_batch, row_key, data, descricao, debito
  ) values (
    '83000000-0000-0000-0000-000000000999', 'smoke-hidden-overlap',
    'card-hidden-overlap-shared', date '2136-02-02', '', 206.00
  );

  select candidates, candidate_count into strict v_candidates, v_candidate_count
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_bank_statement_amount_only', 1, 0, 1,
    array['80000000-0000-0000-0000-000000000001'::uuid]
  );
  if v_candidate_count <> 13
    or jsonb_array_length(v_candidates) <> 12
    or v_candidates @> jsonb_build_array(jsonb_build_object(
      'sourceId', '81000000-0000-0000-0000-000000000999'
    )) then
    raise exception 'Banco hidden-overlap fixture did not place the shared destination beyond bounded evidence: %', v_candidates;
  end if;
  select candidates, candidate_count into strict v_candidates, v_candidate_count
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_bank_statement_amount_only', 1, 0, 1,
    array['80000000-0000-0000-0000-000000000002'::uuid]
  );
  if v_candidate_count <> 1
    or v_candidates->0->>'sourceId' <> '81000000-0000-0000-0000-000000000999' then
    raise exception 'Banco hidden-overlap fixture did not retain the shared destination as the second base sole candidate: %', v_candidates;
  end if;

  select candidates, candidate_count into strict v_candidates, v_candidate_count
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_credit_card_amount_only', 1, 0, 1,
    array['82000000-0000-0000-0000-000000000001'::uuid]
  );
  if v_candidate_count <> 13
    or jsonb_array_length(v_candidates) <> 12
    or v_candidates @> jsonb_build_array(jsonb_build_object(
      'sourceId', '83000000-0000-0000-0000-000000000999'
    )) then
    raise exception 'Visa hidden-overlap fixture did not place the shared destination beyond bounded evidence: %', v_candidates;
  end if;
  select candidates, candidate_count into strict v_candidates, v_candidate_count
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    'financial_documents_cgd_credit_card_amount_only', 1, 0, 1,
    array['82000000-0000-0000-0000-000000000002'::uuid]
  );
  if v_candidate_count <> 1
    or v_candidates->0->>'sourceId' <> '83000000-0000-0000-0000-000000000999' then
    raise exception 'Visa hidden-overlap fixture did not retain the shared destination as the second base sole candidate: %', v_candidates;
  end if;

  update public.financial_reconciliation_automatic_rule_configs
  set enabled = true,
      allow_manual_execution = true,
      include_in_scheduled_batch = false,
      difference_allowed = 0,
      max_difference_days = 1
  where rule_key = 'financial_documents_cgd_bank_statement_amount_only';
  v_run := public.create_financial_reconciliation_automatic_analysis(
    array['financial_documents_cgd_bank_statement_amount_only'], 'manual_rule',
    'smoke:bank-hidden-overlap', '84000000-0000-0000-0000-000000000001'
  );
  v_run_id := (v_run->>'runId')::uuid;
  if v_run->>'status' <> 'completed'
    or v_run#>>'{counts,bases}' <> '2'
    or v_run#>>'{counts,proposed}' <> '0'
    or v_run#>>'{counts,ambiguous}' <> '2'
    or (select count(*) from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = v_run_id
          and proposal.base_source_id in (
            '80000000-0000-0000-0000-000000000001',
            '80000000-0000-0000-0000-000000000002'
          )
          and proposal.status = 'ambiguous'
          and proposal.reason = 'cross_base_overlap') <> 2
    or not exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = v_run_id
        and proposal.base_source_id = '80000000-0000-0000-0000-000000000001'
        and jsonb_array_length(proposal.candidate_groups) = 12
        and not proposal.candidate_groups @> jsonb_build_array(jsonb_build_object(
          'sourceId', '81000000-0000-0000-0000-000000000999'
        ))
    ) then
    raise exception 'Banco hidden candidate-limit overlap did not make both bases non-executable while preserving bounded evidence: %', v_run;
  end if;
  update public.financial_reconciliation_automatic_rule_configs
  set enabled = false, allow_manual_execution = false
  where rule_key = 'financial_documents_cgd_bank_statement_amount_only';
  update public.financial_documents
  set payment = 'smoke:bank-hidden-overlap-covered'
  where id in (
    '80000000-0000-0000-0000-000000000001',
    '80000000-0000-0000-0000-000000000002'
  );

  update public.financial_reconciliation_automatic_rule_configs
  set enabled = true,
      allow_manual_execution = true,
      include_in_scheduled_batch = false,
      difference_allowed = 0,
      max_difference_days = 1
  where rule_key = 'financial_documents_cgd_credit_card_amount_only';
  v_run := public.create_financial_reconciliation_automatic_analysis(
    array['financial_documents_cgd_credit_card_amount_only'], 'manual_rule',
    'smoke:card-hidden-overlap', '84000000-0000-0000-0000-000000000002'
  );
  v_run_id := (v_run->>'runId')::uuid;
  if v_run->>'status' <> 'completed'
    or v_run#>>'{counts,bases}' <> '2'
    or v_run#>>'{counts,proposed}' <> '0'
    or v_run#>>'{counts,ambiguous}' <> '2'
    or (select count(*) from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = v_run_id
          and proposal.base_source_id in (
            '82000000-0000-0000-0000-000000000001',
            '82000000-0000-0000-0000-000000000002'
          )
          and proposal.status = 'ambiguous'
          and proposal.reason = 'cross_base_overlap') <> 2
    or not exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = v_run_id
        and proposal.base_source_id = '82000000-0000-0000-0000-000000000001'
        and jsonb_array_length(proposal.candidate_groups) = 12
        and not proposal.candidate_groups @> jsonb_build_array(jsonb_build_object(
          'sourceId', '83000000-0000-0000-0000-000000000999'
        ))
    ) then
    raise exception 'Visa hidden candidate-limit overlap did not make both bases non-executable while preserving bounded evidence: %', v_run;
  end if;
  update public.financial_reconciliation_automatic_rule_configs
  set enabled = false, allow_manual_execution = false
  where rule_key = 'financial_documents_cgd_credit_card_amount_only';
  update public.financial_documents
  set payment = 'smoke:card-hidden-overlap-covered'
  where id in (
    '82000000-0000-0000-0000-000000000001',
    '82000000-0000-0000-0000-000000000002'
  );
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
  v_third jsonb;
  v_fourth jsonb;
  v_complete jsonb;
  v_complete_retry jsonb;
  v_tomorrow jsonb;
  v_batch_id uuid;
  v_first_run_id uuid;
  v_second_run_id uuid;
  v_third_run_id uuid;
  v_fourth_run_id uuid;
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
        when 'financial_documents_cgd_bank_statement' then 0.10
        when 'financial_documents_cgd_credit_card' then 0.20
        else 0 end,
      max_difference_days = case config.rule_key
        when 'financial_documents_cgd_bank_statement' then 10
        when 'financial_documents_cgd_credit_card' then 11
        else 1 end,
      priority = case config.rule_key
        when 'financial_documents_cgd_bank_statement' then 1
        when 'financial_documents_cgd_credit_card' then 2
        when 'financial_documents_cgd_bank_statement_amount_only' then 3
        else 4 end
  where config.rule_key in (
    'financial_documents_cgd_bank_statement',
    'financial_documents_cgd_credit_card',
    'financial_documents_cgd_bank_statement_amount_only',
    'financial_documents_cgd_credit_card_amount_only'
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
    or v_first->>'batchRuleCount' <> '4'
    or v_first#>>'{run,scope}' <> 'rule'
    or v_first#>>'{run,batchId}' <> v_batch_id::text
    or v_first#>>'{run,batchRuleKey}' <> 'financial_documents_cgd_bank_statement'
    or v_first#>>'{run,batchRulePosition}' <> '1'
    or v_first#>>'{run,batchRuleCount}' <> '4'
    or jsonb_array_length(v_first#>'{run,definitions}') <> 1
    or jsonb_array_length(v_snapshot) <> 4
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
    or jsonb_typeof(v_snapshot#>'{1,definition}') <> 'object'
    or v_snapshot#>>'{2,ruleKey}' <> 'financial_documents_cgd_bank_statement_amount_only'
    or v_snapshot#>>'{2,ruleVersion}' <> '1'
    or v_snapshot#>>'{2,destinationSourceType}' <> 'import_cgd_extrato_ordem'
    or v_snapshot#>>'{2,operator}' <> '+'
    or v_snapshot#>>'{2,priority}' <> '3'
    or v_snapshot#>>'{2,differenceAllowed}' <> '0.00'
    or v_snapshot#>>'{2,maxDifferenceDays}' <> '1'
    or jsonb_typeof(v_snapshot#>'{2,definition}') <> 'object'
    or v_snapshot#>>'{3,ruleKey}' <> 'financial_documents_cgd_credit_card_amount_only'
    or v_snapshot#>>'{3,ruleVersion}' <> '1'
    or v_snapshot#>>'{3,destinationSourceType}' <> 'import_cgd_cartao_credito'
    or v_snapshot#>>'{3,operator}' <> '+'
    or v_snapshot#>>'{3,priority}' <> '4'
    or v_snapshot#>>'{3,differenceAllowed}' <> '0.00'
    or v_snapshot#>>'{3,maxDifferenceDays}' <> '1'
    or jsonb_typeof(v_snapshot#>'{3,definition}') <> 'object' then
    raise exception 'Scheduled batch did not snapshot all four managed rules in deterministic order.';
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
        when 'financial_documents_cgd_credit_card' then 0.30
        when 'financial_documents_cgd_bank_statement' then 0.40
        else 0 end,
      max_difference_days = case config.rule_key
        when 'financial_documents_cgd_credit_card_amount_only' then 2
        when 'financial_documents_cgd_bank_statement_amount_only' then 3
        when 'financial_documents_cgd_credit_card' then 12
        else 13 end,
      priority = case config.rule_key
        when 'financial_documents_cgd_credit_card_amount_only' then 1
        when 'financial_documents_cgd_bank_statement_amount_only' then 2
        when 'financial_documents_cgd_credit_card' then 3
        else 4 end
  where config.rule_key in (
    'financial_documents_cgd_bank_statement',
    'financial_documents_cgd_credit_card',
    'financial_documents_cgd_bank_statement_amount_only',
    'financial_documents_cgd_credit_card_amount_only'
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
  v_third := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-02 00:32:00+00', 'smoke:scheduled-batch'
  );
  v_third_run_id := (v_third#>>'{run,runId}')::uuid;
  if (v_third->>'resumed')::boolean
    or v_third->>'batchRulePosition' <> '3'
    or v_third#>>'{run,batchRuleKey}' <> 'financial_documents_cgd_bank_statement_amount_only'
    or (select count(*) from public.financial_reconciliation_automatic_runs
        where batch_id = v_batch_id) <> 3 then
    raise exception 'Scheduled batch started a later child before the third rule became current.';
  end if;

  update public.financial_reconciliation_automatic_runs
  set status = 'completed',
      analysis_completed_at = coalesce(analysis_completed_at, now()),
      counts = '{"bases":0,"completed":0,"failed":0}'::jsonb,
      finished_at = now(),
      updated_at = now()
  where id = v_third_run_id;
  v_fourth := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-02 00:33:00+00', 'smoke:scheduled-batch'
  );
  v_fourth_run_id := (v_fourth#>>'{run,runId}')::uuid;
  if (v_fourth->>'resumed')::boolean
    or v_fourth->>'batchRulePosition' <> '4'
    or v_fourth#>>'{run,batchRuleKey}' <> 'financial_documents_cgd_credit_card_amount_only'
    or (select count(*) from public.financial_reconciliation_automatic_runs
        where batch_id = v_batch_id) <> 4 then
    raise exception 'Scheduled batch started a later child before the fourth rule became current.';
  end if;

  update public.financial_reconciliation_automatic_runs
  set status = 'completed',
      analysis_completed_at = coalesce(analysis_completed_at, now()),
      counts = '{"bases":0,"completed":0,"failed":0}'::jsonb,
      finished_at = now(),
      updated_at = now()
  where id = v_fourth_run_id;
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
    or v_tomorrow#>>'{run,batchRuleKey}' <> 'financial_documents_cgd_credit_card_amount_only'
    or (select batch.rule_snapshot#>>'{0,priority}'
        from public.financial_reconciliation_automatic_batches batch
        where batch.id = v_tomorrow_batch_id) <> '1'
    or (select batch.rule_snapshot#>>'{0,differenceAllowed}'
        from public.financial_reconciliation_automatic_batches batch
        where batch.id = v_tomorrow_batch_id) <> '0.00' then
    raise exception 'Tomorrow scheduled batch did not use the changed settings snapshot.';
  end if;
end $$;

-- failed scheduled child advances and aggregate batch becomes partial
do $$
declare
  v_first jsonb;
  v_second jsonb;
  v_third jsonb;
  v_fourth jsonb;
  v_complete jsonb;
  v_batch_id uuid;
  v_first_run_id uuid;
  v_second_run_id uuid;
  v_third_run_id uuid;
  v_fourth_run_id uuid;
  v_settings jsonb;
begin
  v_first := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-02 01:01:00+00', 'smoke:scheduled-batch'
  );
  v_batch_id := (v_first->>'batchId')::uuid;
  v_first_run_id := (v_first#>>'{run,runId}')::uuid;
  if not (v_first->>'resumed')::boolean
    or v_first#>>'{run,batchRuleKey}' <> 'financial_documents_cgd_credit_card_amount_only'
    or v_first->>'batchRulePosition' <> '1'
    or (select count(*) from public.financial_reconciliation_automatic_runs
        where batch_id = v_batch_id) <> 1 then
    raise exception 'Configured amount-only child was not resumed alone before failure.';
  end if;
  update public.financial_reconciliation_automatic_runs
  set status = 'failed',
      error_summary = 'internal scheduled fixture failure',
      analysis_error_code = 'analysis_continuation_failed',
      analysis_error_at = now(),
      finished_at = now(),
      updated_at = now()
  where id = v_first_run_id;
  if not exists (
    select 1
    from public.financial_reconciliation_automatic_runs run
    where run.id = v_first_run_id
      and run.status = 'failed'
      and run.analysis_error_code = 'analysis_continuation_failed'
      and run.finished_at is not null
  ) then
    raise exception 'Failed amount-only scheduled child did not persist a terminal failure.';
  end if;

  v_second := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-02 01:02:00+00', 'smoke:scheduled-batch'
  );
  v_second_run_id := (v_second#>>'{run,runId}')::uuid;
  if v_second->>'batchId' <> v_batch_id::text
    or v_second#>>'{run,batchRuleKey}' <> 'financial_documents_cgd_bank_statement_amount_only'
    or v_second->>'batchRulePosition' <> '2' then
    raise exception 'A failed amount-only scheduled child blocked the next snapshotted rule.';
  end if;

  update public.financial_reconciliation_automatic_runs
  set status = 'completed',
      analysis_completed_at = coalesce(analysis_completed_at, now()),
      counts = '{"bases":1,"completed":1,"failed":0}'::jsonb,
      finished_at = now(),
      updated_at = now()
  where id = v_second_run_id;
  v_third := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-02 01:03:00+00', 'smoke:scheduled-batch'
  );
  v_third_run_id := (v_third#>>'{run,runId}')::uuid;
  if v_third->>'batchId' <> v_batch_id::text
    or v_third#>>'{run,batchRuleKey}' <> 'financial_documents_cgd_credit_card'
    or v_third->>'batchRulePosition' <> '3' then
    raise exception 'The configured third scheduled child was not selected.';
  end if;
  update public.financial_reconciliation_automatic_runs
  set status = 'completed',
      analysis_completed_at = coalesce(analysis_completed_at, now()),
      counts = '{"bases":0,"completed":0,"failed":0}'::jsonb,
      finished_at = now(),
      updated_at = now()
  where id = v_third_run_id;

  v_fourth := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-02 01:04:00+00', 'smoke:scheduled-batch'
  );
  v_fourth_run_id := (v_fourth#>>'{run,runId}')::uuid;
  if v_fourth->>'batchId' <> v_batch_id::text
    or v_fourth#>>'{run,batchRuleKey}' <> 'financial_documents_cgd_bank_statement'
    or v_fourth->>'batchRulePosition' <> '4' then
    raise exception 'The configured fourth scheduled child was not selected.';
  end if;
  update public.financial_reconciliation_automatic_runs
  set status = 'completed',
      analysis_completed_at = coalesce(analysis_completed_at, now()),
      counts = '{"bases":0,"completed":0,"failed":0}'::jsonb,
      finished_at = now(),
      updated_at = now()
  where id = v_fourth_run_id;

  v_complete := public.claim_financial_reconciliation_automatic_schedule(
    '2094-01-02 01:05:00+00', 'smoke:scheduled-batch'
  );
  if v_complete->>'reason' <> 'batch_complete'
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_batches batch
      where batch.id = v_batch_id
        and batch.status = 'partial'
        and batch.counts @> '{"ruleCount":4,"childCount":4,"completedChildren":3,"failedChildren":1}'::jsonb
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

  set constraints financial_reconciliation_automatic_rule_configs_priority_key deferred;
  update public.financial_reconciliation_automatic_rule_configs config
  set enabled = config.rule_key in (
        'financial_documents_cgd_bank_statement',
        'financial_documents_cgd_credit_card'
      ),
      include_in_scheduled_batch = config.rule_key in (
        'financial_documents_cgd_bank_statement',
        'financial_documents_cgd_credit_card'
      ),
      priority = case config.rule_key
        when 'financial_documents_cgd_bank_statement' then 1
        when 'financial_documents_cgd_credit_card' then 2
        when 'financial_documents_cgd_bank_statement_amount_only' then 3
        else 4 end
  where config.rule_key in (
    'financial_documents_cgd_bank_statement',
    'financial_documents_cgd_credit_card',
    'financial_documents_cgd_bank_statement_amount_only',
    'financial_documents_cgd_credit_card_amount_only'
  );
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

-- amount-only proposal display details enrich future snapshots and only unfinished legacy rows
do $$
declare
  v_bank_rule_snapshot jsonb;
  v_card_rule_snapshot jsonb;
  v_bank_base_snapshot jsonb;
  v_bank_candidates jsonb;
  v_bank_candidate_count integer;
  v_bank_combination record;
  v_bank_evidence jsonb;
  v_card_base_snapshot jsonb;
  v_card_candidates jsonb;
  v_card_candidate_count integer;
  v_card_combination record;
  v_card_evidence jsonb;
  v_preserve_base_snapshot jsonb;
  v_preserve_candidates jsonb;
  v_preserve_candidate_count integer;
  v_preserve_combination record;
  v_preserve_evidence jsonb;
begin
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name, supplier_nif,
    payment, amount, fat
  ) values
    (
      '81000000-0000-0000-0000-000000000001', date '2145-01-10',
      'FT <DETAIL/1>', 'Base & detail', 'Supplier <Detail>', 'PT500000001',
      'Banco', 411.00, 'S'
    ),
    (
      '81000000-0000-0000-0000-000000000002', date '2145-01-11',
      'FT <DETAIL/1>', 'Base & detail', 'Supplier <Detail>', 'PT500000001',
      'Visa', 412.00, 'S'
    ),
    (
      '81000000-0000-0000-0000-000000000003', date '2145-01-13',
      'FT <PRESERVE/1>', 'Base preserve original', 'Supplier preserve original', '',
      'Banco', 413.00, 'S'
    ),
    (
      '81000000-0000-0000-0000-000000000004', date '2145-01-14',
      'FT <CANDIDATE/1>', 'Candidate base original', 'Candidate supplier original',
      'PT500000004', 'Banco', 414.00, 'S'
    );

  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  ) values
    (
      '81000000-0000-0000-0000-000000000011', 'smoke-proposal-details',
      'smoke-proposal-details-bank', date '2145-01-10', 'Bank <detail>', -411.00
    ),
    (
      '81000000-0000-0000-0000-000000000013', 'smoke-proposal-details',
      'smoke-proposal-details-preserve', date '2145-01-13',
      'Bank preserve original', -413.00
    ),
    (
      '81000000-0000-0000-0000-000000000014', 'smoke-proposal-details',
      'smoke-proposal-details-candidate', date '2145-01-14',
      'Candidate preserve original', -414.00
    );

  insert into public.import_cgd_cartao_credito (
    id, import_batch, row_key, data, descricao, debito
  ) values (
    '81000000-0000-0000-0000-000000000012', 'smoke-proposal-details',
    'smoke-proposal-details-card', date '2145-01-11', 'Card & detail', 412.00
  );

  select jsonb_build_object(
    'ruleKey', config.rule_key,
    'ruleVersion', config.rule_version,
    'displayName', definition.display_name,
    'priority', config.priority,
    'differenceAllowed', config.difference_allowed,
    'maxDifferenceDays', config.max_difference_days,
    'destinationSourceType', 'import_cgd_extrato_ordem',
    'definition', definition.definition,
    'operator', source_rule.operator
  )
  into strict v_bank_rule_snapshot
  from public.financial_reconciliation_automatic_rule_configs config
  join public.financial_reconciliation_automatic_rule_definitions definition
    on definition.rule_key = config.rule_key
   and definition.version = config.rule_version
  join public.financial_reconciliation_source_rules source_rule
    on source_rule.base_source_type = definition.base_source_type
   and source_rule.matching_source_type = 'import_cgd_extrato_ordem'
  where config.rule_key = 'financial_documents_cgd_bank_statement_amount_only'
    and config.rule_version = 1;

  select jsonb_build_object(
    'ruleKey', config.rule_key,
    'ruleVersion', config.rule_version,
    'displayName', definition.display_name,
    'priority', config.priority,
    'differenceAllowed', config.difference_allowed,
    'maxDifferenceDays', config.max_difference_days,
    'destinationSourceType', 'import_cgd_cartao_credito',
    'definition', definition.definition,
    'operator', source_rule.operator
  )
  into strict v_card_rule_snapshot
  from public.financial_reconciliation_automatic_rule_configs config
  join public.financial_reconciliation_automatic_rule_definitions definition
    on definition.rule_key = config.rule_key
   and definition.version = config.rule_version
  join public.financial_reconciliation_source_rules source_rule
    on source_rule.base_source_type = definition.base_source_type
   and source_rule.matching_source_type = 'import_cgd_cartao_credito'
  where config.rule_key = 'financial_documents_cgd_credit_card_amount_only'
    and config.rule_version = 1;

  select base_snapshot, candidates, candidate_count
  into strict v_bank_base_snapshot, v_bank_candidates, v_bank_candidate_count
  from public.financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(
    'financial_documents_cgd_bank_statement_amount_only', 1, 0,
    (v_bank_rule_snapshot->>'maxDifferenceDays')::integer,
    array['81000000-0000-0000-0000-000000000001'::uuid]
  );
  if v_bank_candidate_count <> 1
    or v_bank_base_snapshot ?| array['docNumber','description','supplierName','supplierNif']
    or v_bank_candidates->0 ? 'description' then
    raise exception 'The Bank proposal-detail fixture did not begin with the legacy minimal snapshot.';
  end if;
  select * into strict v_bank_combination
  from public.financial_reconciliation_automatic_build_combinations(
    v_bank_base_snapshot, v_bank_candidates,
    '{"import_cgd_extrato_ordem":"+"}'::jsonb, 0, 1
  );
  select coalesce(jsonb_agg(item.value->'evidence' order by item.ordinality), '[]'::jsonb)
  into v_bank_evidence
  from jsonb_array_elements(v_bank_combination.items) with ordinality item(value, ordinality);

  select base_snapshot, candidates, candidate_count
  into strict v_card_base_snapshot, v_card_candidates, v_card_candidate_count
  from public.financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(
    'financial_documents_cgd_credit_card_amount_only', 1, 0,
    (v_card_rule_snapshot->>'maxDifferenceDays')::integer,
    array['81000000-0000-0000-0000-000000000002'::uuid]
  );
  if v_card_candidate_count <> 1
    or v_card_base_snapshot ?| array['docNumber','description','supplierName','supplierNif']
    or v_card_candidates->0 ? 'description' then
    raise exception 'The Credit Card proposal-detail fixture did not begin with the legacy minimal snapshot.';
  end if;
  select * into strict v_card_combination
  from public.financial_reconciliation_automatic_build_combinations(
    v_card_base_snapshot, v_card_candidates,
    '{"import_cgd_cartao_credito":"+"}'::jsonb, 0, 1
  );
  select coalesce(jsonb_agg(item.value->'evidence' order by item.ordinality), '[]'::jsonb)
  into v_card_evidence
  from jsonb_array_elements(v_card_combination.items) with ordinality item(value, ordinality);

  select base_snapshot, candidates, candidate_count
  into strict v_preserve_base_snapshot, v_preserve_candidates, v_preserve_candidate_count
  from public.financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(
    'financial_documents_cgd_bank_statement_amount_only', 1, 0,
    (v_bank_rule_snapshot->>'maxDifferenceDays')::integer,
    array['81000000-0000-0000-0000-000000000003'::uuid]
  );
  if v_preserve_candidate_count <> 1
    or v_preserve_base_snapshot ?| array['docNumber','description','supplierName','supplierNif']
    or v_preserve_candidates->0 ? 'description' then
    raise exception 'The reapply-preservation fixture did not begin with the legacy minimal snapshot.';
  end if;
  select * into strict v_preserve_combination
  from public.financial_reconciliation_automatic_build_combinations(
    v_preserve_base_snapshot, v_preserve_candidates,
    '{"import_cgd_extrato_ordem":"+"}'::jsonb, 0, 1
  );
  select coalesce(jsonb_agg(item.value->'evidence' order by item.ordinality), '[]'::jsonb)
  into v_preserve_evidence
  from jsonb_array_elements(v_preserve_combination.items) with ordinality item(value, ordinality);

  insert into public.financial_reconciliation_automatic_runs (
    id, trigger, scope, status, actor, client_request_id,
    definition_config_snapshot, analysis_completed_at, finished_at
  ) values
    (
      '81000000-0000-0000-0000-000000000101', 'manual', 'rule', 'ready',
      'smoke:proposal-details-bank', '81000000-0000-0000-0000-000000000111',
      jsonb_build_array(v_bank_rule_snapshot), now(), null
    ),
    (
      '81000000-0000-0000-0000-000000000102', 'manual', 'rule', 'ready',
      'smoke:proposal-details-flat', '81000000-0000-0000-0000-000000000112',
      '[]'::jsonb, now(), null
    ),
    (
      '81000000-0000-0000-0000-000000000103', 'manual', 'rule', 'ready',
      'smoke:proposal-details-nested', '81000000-0000-0000-0000-000000000113',
      '[]'::jsonb, now(), null
    ),
    (
      '81000000-0000-0000-0000-000000000104', 'manual', 'rule', 'ready',
      'smoke:proposal-details-card', '81000000-0000-0000-0000-000000000114',
      jsonb_build_array(v_card_rule_snapshot), now(), null
    ),
    (
      '81000000-0000-0000-0000-000000000105', 'manual', 'rule', 'completed',
      'smoke:proposal-details-completed', '81000000-0000-0000-0000-000000000115',
      '[]'::jsonb, now(), now()
    ),
    (
      '81000000-0000-0000-0000-000000000106', 'manual', 'rule', 'ready',
      'smoke:proposal-details-missing', '81000000-0000-0000-0000-000000000116',
      '[]'::jsonb, now(), null
    ),
    (
      '81000000-0000-0000-0000-000000000107', 'manual', 'rule', 'ready',
      'smoke:proposal-details-preserve', '81000000-0000-0000-0000-000000000117',
      jsonb_build_array(v_bank_rule_snapshot), now(), null
    ),
    (
      '81000000-0000-0000-0000-000000000108', 'manual', 'rule', 'ready',
      'smoke:proposal-details-completed-open', '81000000-0000-0000-0000-000000000118',
      '[]'::jsonb, now(), null
    ),
    (
      '81000000-0000-0000-0000-000000000109', 'manual', 'rule', 'ready',
      'smoke:proposal-details-wrong-base-type', '81000000-0000-0000-0000-000000000119',
      '[]'::jsonb, now(), null
    ),
    (
      '81000000-0000-0000-0000-000000000110', 'manual', 'rule', 'ready',
      'smoke:proposal-details-preserve-candidate', '81000000-0000-0000-0000-000000000120',
      '[]'::jsonb, now(), null
    );

  insert into public.financial_reconciliations (
    id, status, base_source_type, matching_source_types, completion_type,
    created_by, completed_by, completed_at
  ) values (
    '81000000-0000-0000-0000-000000000301', 'complete',
    'financial_documents', '["import_cgd_extrato_ordem"]'::jsonb, 'normal',
    'smoke:proposal-details-completed', 'smoke:proposal-details-completed', now()
  );

  insert into public.financial_reconciliation_automatic_proposals (
    id, run_id, rule_key, rule_version, base_source_type, base_source_id,
    base_source_date, base_snapshot, items, evidence, candidate_groups,
    calculated_difference, allowed_difference, status, reason, signature,
    reconciliation_id, completed_at, updated_at
  ) values
    (
      '81000000-0000-0000-0000-000000000201',
      '81000000-0000-0000-0000-000000000101',
      'financial_documents_cgd_bank_statement_amount_only', 1,
      'financial_documents', '81000000-0000-0000-0000-000000000001',
      date '2145-01-10', v_bank_base_snapshot,
      v_bank_combination.items, v_bank_evidence, '[]'::jsonb,
      v_bank_combination.calculated_difference, 0, 'proposed', '',
      v_bank_combination.signature, null, null, '2000-01-01 00:00:00+00'
    ),
    (
      '81000000-0000-0000-0000-000000000202',
      '81000000-0000-0000-0000-000000000102',
      'financial_documents_cgd_bank_statement_amount_only', 1,
      'financial_documents', '81000000-0000-0000-0000-000000000001',
      date '2145-01-10',
      jsonb_build_object(
        'sourceType', 'financial_documents',
        'sourceId', '81000000-0000-0000-0000-000000000001',
        'sourceDate', '2145-01-10', 'amount', 411.00,
        'futureBaseKey', 'keep-flat'
      ),
      '[]'::jsonb,
      '[{"proposalEvidence":"keep-flat"}]'::jsonb,
      jsonb_build_array(
        jsonb_build_object(
          'sourceType', 'import_cgd_extrato_ordem',
          'sourceId', '81000000-0000-0000-0000-000000000011',
          'sourceDate', '2145-01-10', 'amount', -411.00,
          'evidence', jsonb_build_object('marker', 'flat-first'),
          'futureItemKey', 'keep-flat-first'
        ),
        jsonb_build_object(
          'sourceType', 'future_destination_type',
          'sourceId', '81000000-0000-0000-0000-000000000091',
          'sourceDate', '2145-01-12', 'amount', -1,
          'evidence', jsonb_build_object('marker', 'flat-second'),
          'futureItemKey', 'keep-flat-second'
        )
      ),
      0, 0, 'ambiguous', 'candidate_limit', 'proposal-details-flat',
      null, null, '2000-01-01 00:00:00+00'
    ),
    (
      '81000000-0000-0000-0000-000000000203',
      '81000000-0000-0000-0000-000000000103',
      'financial_documents_cgd_bank_statement_amount_only', 1,
      'financial_documents', '81000000-0000-0000-0000-000000000001',
      date '2145-01-10',
      jsonb_build_object(
        'sourceType', 'financial_documents',
        'sourceId', '81000000-0000-0000-0000-000000000001',
        'sourceDate', '2145-01-10', 'amount', 411.00,
        'futureBaseKey', 'keep-nested'
      ),
      '[]'::jsonb,
      '[{"proposalEvidence":"keep-nested"}]'::jsonb,
      jsonb_build_array(
        jsonb_build_array(
          jsonb_build_object(
            'sourceType', 'import_cgd_extrato_ordem',
            'sourceId', '81000000-0000-0000-0000-000000000011',
            'sourceDate', '2145-01-10', 'amount', -411.00,
            'evidence', jsonb_build_object('marker', 'nested-first'),
            'futureItemKey', 'keep-nested-first'
          ),
          jsonb_build_object(
            'sourceType', 'import_cgd_extrato_ordem',
            'sourceId', 'not-a-uuid',
            'sourceDate', '2145-01-10', 'amount', -411.00,
            'evidence', jsonb_build_object('marker', 'nested-malformed'),
            'futureItemKey', 'keep-nested-malformed'
          )
        ),
        jsonb_build_array(
          jsonb_build_object(
            'sourceType', 'future_destination_type',
            'sourceId', '81000000-0000-0000-0000-000000000092',
            'sourceDate', '2145-01-13', 'amount', -2,
            'evidence', jsonb_build_object('marker', 'nested-second-group'),
            'futureItemKey', 'keep-nested-second-group'
          )
        )
      ),
      0, 0, 'ambiguous', 'multiple_combinations', 'proposal-details-nested',
      null, null, '2000-01-01 00:00:00+00'
    ),
    (
      '81000000-0000-0000-0000-000000000204',
      '81000000-0000-0000-0000-000000000104',
      'financial_documents_cgd_credit_card_amount_only', 1,
      'financial_documents', '81000000-0000-0000-0000-000000000002',
      date '2145-01-11', v_card_base_snapshot,
      v_card_combination.items, v_card_evidence, '[]'::jsonb,
      v_card_combination.calculated_difference, 0, 'proposed', '',
      v_card_combination.signature, null, null, '2000-01-01 00:00:00+00'
    ),
    (
      '81000000-0000-0000-0000-000000000205',
      '81000000-0000-0000-0000-000000000105',
      'financial_documents_cgd_bank_statement_amount_only', 1,
      'financial_documents', '81000000-0000-0000-0000-000000000001',
      date '2145-01-10',
      jsonb_build_object(
        'sourceType', 'financial_documents',
        'sourceId', '81000000-0000-0000-0000-000000000001',
        'sourceDate', '2145-01-10', 'amount', 411.00,
        'futureBaseKey', 'keep-completed'
      ),
      v_bank_combination.items, v_bank_evidence, '[]'::jsonb,
      0, 0, 'completed', '', 'proposal-details-completed',
      '81000000-0000-0000-0000-000000000301', now(), '2000-01-01 00:00:00+00'
    ),
    (
      '81000000-0000-0000-0000-000000000206',
      '81000000-0000-0000-0000-000000000106',
      'financial_documents_cgd_bank_statement_amount_only', 1,
      'financial_documents', '81000000-0000-0000-0000-000000000099',
      date '2145-01-12',
      jsonb_build_object(
        'sourceType', 'financial_documents',
        'sourceId', '81000000-0000-0000-0000-000000000099',
        'sourceDate', '2145-01-12', 'amount', 499.00,
        'futureBaseKey', 'keep-missing'
      ),
      jsonb_build_array(jsonb_build_object(
        'sourceType', 'import_cgd_extrato_ordem',
        'sourceId', '81000000-0000-0000-0000-000000000098',
        'sourceDate', '2145-01-12', 'amount', -499.00,
        'evidence', jsonb_build_object('marker', 'missing-item'),
        'futureItemKey', 'keep-missing-item'
      )),
      '[{"proposalEvidence":"keep-missing"}]'::jsonb,
      jsonb_build_array(jsonb_build_object(
        'sourceType', 'import_cgd_extrato_ordem',
        'sourceId', '81000000-0000-0000-0000-000000000098',
        'sourceDate', '2145-01-12', 'amount', -499.00,
        'evidence', jsonb_build_object('marker', 'missing-candidate'),
        'futureItemKey', 'keep-missing-candidate'
      )),
      0, 0, 'proposed', '', 'proposal-details-missing',
      null, null, '2000-01-01 00:00:00+00'
    ),
    (
      '81000000-0000-0000-0000-000000000207',
      '81000000-0000-0000-0000-000000000107',
      'financial_documents_cgd_bank_statement_amount_only', 1,
      'financial_documents', '81000000-0000-0000-0000-000000000003',
      date '2145-01-13',
      v_preserve_base_snapshot || jsonb_build_object('supplierNif', 'null'::jsonb),
      v_preserve_combination.items, v_preserve_evidence, '[]'::jsonb,
      v_preserve_combination.calculated_difference, 0, 'proposed', '',
      v_preserve_combination.signature, null, null, '2000-01-01 00:00:00+00'
    ),
    (
      '81000000-0000-0000-0000-000000000208',
      '81000000-0000-0000-0000-000000000108',
      'financial_documents_cgd_bank_statement_amount_only', 1,
      'financial_documents', '81000000-0000-0000-0000-000000000001',
      date '2145-01-10',
      jsonb_build_object(
        'sourceType', 'financial_documents',
        'sourceId', '81000000-0000-0000-0000-000000000001',
        'sourceDate', '2145-01-10', 'amount', 411.00,
        'futureBaseKey', 'keep-completed-open'
      ),
      v_bank_combination.items, v_bank_evidence, '[]'::jsonb,
      0, 0, 'completed', '', 'proposal-details-completed-open',
      '81000000-0000-0000-0000-000000000301', now(), '2000-01-01 00:00:00+00'
    ),
    (
      '81000000-0000-0000-0000-000000000209',
      '81000000-0000-0000-0000-000000000109',
      'financial_documents_cgd_bank_statement_amount_only', 1,
      'financial_documents', '81000000-0000-0000-0000-000000000001',
      date '2145-01-10',
      jsonb_build_object(
        'sourceType', 'future_financial_documents_adapter',
        'sourceId', '81000000-0000-0000-0000-000000000001',
        'sourceDate', '2145-01-10', 'amount', 411.00,
        'futureBaseKey', 'keep-wrong-base-type'
      ),
      '[]'::jsonb, '[]'::jsonb, '[]'::jsonb,
      0, 0, 'proposed', '', 'proposal-details-wrong-base-type',
      null, null, '2000-01-01 00:00:00+00'
    ),
    (
      '81000000-0000-0000-0000-000000000210',
      '81000000-0000-0000-0000-000000000110',
      'financial_documents_cgd_bank_statement_amount_only', 1,
      'financial_documents', '81000000-0000-0000-0000-000000000004',
      date '2145-01-14',
      jsonb_build_object(
        'sourceType', 'financial_documents',
        'sourceId', '81000000-0000-0000-0000-000000000004',
        'sourceDate', '2145-01-14', 'amount', 414.00,
        'futureBaseKey', 'keep-candidate-preservation'
      ),
      '[]'::jsonb, '[{"proposalEvidence":"keep-candidate-preservation"}]'::jsonb,
      jsonb_build_array(jsonb_build_object(
        'sourceType', 'import_cgd_extrato_ordem',
        'sourceId', '81000000-0000-0000-0000-000000000014',
        'sourceDate', '2145-01-14', 'amount', -414.00,
        'evidence', jsonb_build_object('marker', 'preserve-candidate'),
        'futureItemKey', 'keep-preserve-candidate'
      )),
      0, 0, 'ambiguous', 'candidate_limit', 'proposal-details-preserve-candidate',
      null, null, '2000-01-01 00:00:00+00'
    );

  insert into public.financial_reconciliation_audit (
    id, reconciliation_id, action, actor, comment, difference_amount, metadata
  )
  select
    '81000000-0000-0000-0000-000000000401',
    '81000000-0000-0000-0000-000000000301',
    'complete', 'smoke:proposal-details-completed',
    'Completed proposal snapshots must remain immutable.', 0,
    jsonb_build_object(
      'baseSnapshot', proposal.base_snapshot,
      'items', proposal.items,
      'evidence', proposal.evidence,
      'candidateGroups', proposal.candidate_groups,
      'futureAuditKey', 'keep-audit'
    )
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.id = '81000000-0000-0000-0000-000000000205';
end $$;

create temporary table automatic_proposal_detail_baseline as
select fixture.fixture_name, proposal.id as proposal_id, to_jsonb(proposal) as proposal_json
from (values
  ('bank', '81000000-0000-0000-0000-000000000201'::uuid),
  ('flat', '81000000-0000-0000-0000-000000000202'::uuid),
  ('nested', '81000000-0000-0000-0000-000000000203'::uuid),
  ('card', '81000000-0000-0000-0000-000000000204'::uuid),
  ('completed', '81000000-0000-0000-0000-000000000205'::uuid),
  ('missing', '81000000-0000-0000-0000-000000000206'::uuid),
  ('preserve-reapply', '81000000-0000-0000-0000-000000000207'::uuid),
  ('completed-open', '81000000-0000-0000-0000-000000000208'::uuid),
  ('wrong-base-type', '81000000-0000-0000-0000-000000000209'::uuid),
  ('preserve-candidate', '81000000-0000-0000-0000-000000000210'::uuid)
) fixture(fixture_name, proposal_id)
join public.financial_reconciliation_automatic_proposals proposal
  on proposal.id = fixture.proposal_id;

create temporary table automatic_proposal_detail_audit_baseline as
select audit.id as audit_id, audit.metadata
from public.financial_reconciliation_audit audit
where audit.id = '81000000-0000-0000-0000-000000000401';

\ir ../supabase-migrations/2026-08-18-financial-reconciliation-automation-proposal-details.sql

do $$
begin
  if not exists (
    select 1
    from pg_trigger trigger_row
    where trigger_row.tgrelid =
        'public.financial_reconciliation_automatic_proposals'::regclass
      and trigger_row.tgname =
        'financial_reconciliation_automatic_proposal_snapshot_immutable'
      and not trigger_row.tgisinternal
      and trigger_row.tgenabled = 'O'
  ) then
    raise exception 'Proposal-detail migration did not restore the immutable base-snapshot trigger.';
  end if;
end $$;

create temporary table automatic_proposal_detail_after_first as
select proposal.id as proposal_id, to_jsonb(proposal) as proposal_json
from public.financial_reconciliation_automatic_proposals proposal
where proposal.id in (
  '81000000-0000-0000-0000-000000000201',
  '81000000-0000-0000-0000-000000000202',
  '81000000-0000-0000-0000-000000000203',
  '81000000-0000-0000-0000-000000000204',
  '81000000-0000-0000-0000-000000000205',
  '81000000-0000-0000-0000-000000000206',
  '81000000-0000-0000-0000-000000000207',
  '81000000-0000-0000-0000-000000000208',
  '81000000-0000-0000-0000-000000000209',
  '81000000-0000-0000-0000-000000000210'
);

update public.financial_documents
set doc_number = 'FT <PRESERVE/CHANGED>',
    description = 'Base preserve changed',
    supplier_name = 'Supplier preserve changed',
    supplier_nif = 'PT500000003'
where id = '81000000-0000-0000-0000-000000000003';

update public.import_cgd_extrato_ordem
set descritivo = 'Bank preserve changed'
where id = '81000000-0000-0000-0000-000000000013';

update public.import_cgd_extrato_ordem
set descritivo = 'Candidate preserve changed'
where id = '81000000-0000-0000-0000-000000000014';

create or replace function pg_temp.reject_automatic_proposal_detail_reapply_update()
returns trigger
language plpgsql
as $$
begin
  raise exception 'Proposal-detail migration reapply updated fixture proposal %.', old.id;
end
$$;

create trigger automatic_proposal_detail_reapply_guard
before update on public.financial_reconciliation_automatic_proposals
for each row
when (old.id in (
  '81000000-0000-0000-0000-000000000201',
  '81000000-0000-0000-0000-000000000202',
  '81000000-0000-0000-0000-000000000203',
  '81000000-0000-0000-0000-000000000204',
  '81000000-0000-0000-0000-000000000205',
  '81000000-0000-0000-0000-000000000206',
  '81000000-0000-0000-0000-000000000207',
  '81000000-0000-0000-0000-000000000208',
  '81000000-0000-0000-0000-000000000209',
  '81000000-0000-0000-0000-000000000210'
))
execute function pg_temp.reject_automatic_proposal_detail_reapply_update();

\ir ../supabase-migrations/2026-08-18-financial-reconciliation-automation-proposal-details.sql

drop trigger automatic_proposal_detail_reapply_guard
  on public.financial_reconciliation_automatic_proposals;

do $$
declare
  v_bank_base_snapshot jsonb;
  v_bank_candidates jsonb;
  v_bank_candidate_count integer;
  v_card_base_snapshot jsonb;
  v_card_candidates jsonb;
  v_card_candidate_count integer;
  v_result jsonb;
  v_signature text;
begin
  if not exists (
    select 1
    from pg_trigger trigger_row
    where trigger_row.tgrelid =
        'public.financial_reconciliation_automatic_proposals'::regclass
      and trigger_row.tgname =
        'financial_reconciliation_automatic_proposal_snapshot_immutable'
      and not trigger_row.tgisinternal
      and trigger_row.tgenabled = 'O'
  ) then
    raise exception 'Proposal-detail migration reapply did not restore the immutable base-snapshot trigger.';
  end if;

  if exists (
    select 1
    from automatic_proposal_detail_baseline baseline
    join public.financial_reconciliation_automatic_proposals proposal
      on proposal.id = baseline.proposal_id
    where baseline.fixture_name in ('bank', 'flat', 'nested', 'card')
      and (
        (proposal.base_snapshot - array['docNumber','description','supplierName','supplierNif'])
          is distinct from
        ((baseline.proposal_json->'base_snapshot') - array['docNumber','description','supplierName','supplierNif'])
        or proposal.evidence is distinct from baseline.proposal_json->'evidence'
        or proposal.updated_at is not distinct from
          (baseline.proposal_json->>'updated_at')::timestamptz
      )
  ) then
    raise exception 'Unfinished proposal enrichment changed unrelated base JSON, evidence, or no-op timestamps.';
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_proposals proposal
    join automatic_proposal_detail_baseline baseline on baseline.proposal_id = proposal.id
    where baseline.fixture_name = 'bank'
      and proposal.base_snapshot->>'docNumber' = 'FT <DETAIL/1>'
      and proposal.base_snapshot->>'description' = 'Base & detail'
      and proposal.base_snapshot->>'supplierName' = 'Supplier <Detail>'
      and proposal.base_snapshot->>'supplierNif' = 'PT500000001'
      and proposal.items->0->>'description' = 'Bank <detail>'
      and (proposal.items->0 - 'description')
        is not distinct from (baseline.proposal_json#>'{items,0}')
      and proposal.candidate_groups = baseline.proposal_json->'candidate_groups'
  ) then
    raise exception 'Unfinished unique Bank proposal did not preserve and enrich its legacy snapshots.';
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_proposals proposal
    join automatic_proposal_detail_baseline baseline on baseline.proposal_id = proposal.id
    where baseline.fixture_name = 'flat'
      and proposal.base_snapshot->>'docNumber' = 'FT <DETAIL/1>'
      and proposal.base_snapshot->>'description' = 'Base & detail'
      and proposal.base_snapshot->>'supplierName' = 'Supplier <Detail>'
      and proposal.base_snapshot->>'supplierNif' = 'PT500000001'
      and proposal.candidate_groups->0->>'description' = 'Bank <detail>'
      and proposal.candidate_groups#>>'{0,evidence,marker}' = 'flat-first'
      and proposal.candidate_groups#>>'{1,evidence,marker}' = 'flat-second'
      and (proposal.candidate_groups->0 - 'description')
        is not distinct from (baseline.proposal_json#>'{candidate_groups,0}')
      and proposal.candidate_groups->1
        is not distinct from (baseline.proposal_json#>'{candidate_groups,1}')
      and proposal.items = baseline.proposal_json->'items'
  ) then
    raise exception 'Flat candidate_limit groups lost ordering, future keys, or Bank details.';
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_proposals proposal
    join automatic_proposal_detail_baseline baseline on baseline.proposal_id = proposal.id
    where baseline.fixture_name = 'nested'
      and proposal.base_snapshot->>'docNumber' = 'FT <DETAIL/1>'
      and proposal.candidate_groups->0->0->>'description' = 'Bank <detail>'
      and proposal.candidate_groups#>>'{0,0,evidence,marker}' = 'nested-first'
      and proposal.candidate_groups#>>'{0,1,evidence,marker}' = 'nested-malformed'
      and proposal.candidate_groups#>>'{1,0,evidence,marker}' = 'nested-second-group'
      and (proposal.candidate_groups->0->0 - 'description')
        is not distinct from (baseline.proposal_json#>'{candidate_groups,0,0}')
      and proposal.candidate_groups->0->1
        is not distinct from (baseline.proposal_json#>'{candidate_groups,0,1}')
      and proposal.candidate_groups->1
        is not distinct from (baseline.proposal_json#>'{candidate_groups,1}')
      and jsonb_array_length(proposal.candidate_groups) = 2
      and jsonb_array_length(proposal.candidate_groups->0) = 2
  ) then
    raise exception 'Nested multiple_combinations groups lost shape, ordering, or missing-source behavior.';
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_proposals proposal
    join automatic_proposal_detail_baseline baseline on baseline.proposal_id = proposal.id
    where baseline.fixture_name = 'card'
      and proposal.base_snapshot->>'docNumber' = 'FT <DETAIL/1>'
      and proposal.base_snapshot->>'description' = 'Base & detail'
      and proposal.base_snapshot->>'supplierName' = 'Supplier <Detail>'
      and proposal.base_snapshot->>'supplierNif' = 'PT500000001'
      and proposal.items->0->>'description' = 'Card & detail'
      and (proposal.items->0 - 'description')
        is not distinct from (baseline.proposal_json#>'{items,0}')
      and proposal.candidate_groups = baseline.proposal_json->'candidate_groups'
  ) then
    raise exception 'Unfinished unique Credit Card proposal did not preserve and enrich its legacy snapshots.';
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_proposals proposal
    join automatic_proposal_detail_after_first first_apply
      on first_apply.proposal_id = proposal.id
    where proposal.id = '81000000-0000-0000-0000-000000000207'
      and proposal.base_snapshot->>'docNumber' = 'FT <PRESERVE/1>'
      and proposal.base_snapshot->>'description' = 'Base preserve original'
      and proposal.base_snapshot->>'supplierName' = 'Supplier preserve original'
      and proposal.base_snapshot ? 'supplierNif'
      and proposal.base_snapshot->'supplierNif' = 'null'::jsonb
      and proposal.items->0->>'description' = 'Bank preserve original'
      and to_jsonb(proposal) is not distinct from first_apply.proposal_json
  ) then
    raise exception 'Migration reapply overwrote an existing base/item display key or its JSON null.';
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_proposals proposal
    join automatic_proposal_detail_after_first first_apply
      on first_apply.proposal_id = proposal.id
    where proposal.id = '81000000-0000-0000-0000-000000000210'
      and proposal.candidate_groups->0->>'description' = 'Candidate preserve original'
      and proposal.candidate_groups#>>'{0,evidence,marker}' = 'preserve-candidate'
      and to_jsonb(proposal) is not distinct from first_apply.proposal_json
  ) then
    raise exception 'Migration reapply overwrote an existing candidate display key or timestamp.';
  end if;

  if exists (
    select 1
    from automatic_proposal_detail_baseline baseline
    join public.financial_reconciliation_automatic_proposals proposal
      on proposal.id = baseline.proposal_id
    where baseline.fixture_name in (
      'completed', 'missing', 'completed-open', 'wrong-base-type'
    )
      and to_jsonb(proposal) is distinct from baseline.proposal_json
  ) then
    raise exception 'Protected completed, missing-source, or wrong-base-type proposal JSON changed during the backfill.';
  end if;

  if exists (
    select 1
    from automatic_proposal_detail_audit_baseline baseline
    join public.financial_reconciliation_audit audit on audit.id = baseline.audit_id
    where audit.metadata is distinct from baseline.metadata
  ) then
    raise exception 'Completed reconciliation audit metadata changed during the backfill.';
  end if;

  if exists (
    select 1
    from automatic_proposal_detail_after_first first_apply
    join public.financial_reconciliation_automatic_proposals proposal
      on proposal.id = first_apply.proposal_id
    where to_jsonb(proposal) is distinct from first_apply.proposal_json
  ) then
    raise exception 'Proposal-detail migration reapply was not a JSON no-op.';
  end if;

  foreach v_signature in array array[
    'public.financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])',
    'public.financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])'
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
      raise exception 'Proposal-detail adapter security changed for %.', v_signature;
    end if;
  end loop;

  select base_snapshot, candidates, candidate_count
  into strict v_bank_base_snapshot, v_bank_candidates, v_bank_candidate_count
  from public.financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(
    'financial_documents_cgd_bank_statement_amount_only', 1, 0, 1,
    array['81000000-0000-0000-0000-000000000001'::uuid]
  );
  if v_bank_candidate_count <> 1
    or v_bank_base_snapshot->>'docNumber' <> 'FT <DETAIL/1>'
    or v_bank_base_snapshot->>'description' <> 'Base & detail'
    or v_bank_base_snapshot->>'supplierName' <> 'Supplier <Detail>'
    or v_bank_base_snapshot->>'supplierNif' <> 'PT500000001'
    or v_bank_candidates->0->>'description' <> 'Bank <detail>' then
    raise exception 'New Bank amount-only snapshots do not contain the proposal display details.';
  end if;

  select base_snapshot, candidates, candidate_count
  into strict v_card_base_snapshot, v_card_candidates, v_card_candidate_count
  from public.financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(
    'financial_documents_cgd_credit_card_amount_only', 1, 0, 1,
    array['81000000-0000-0000-0000-000000000002'::uuid]
  );
  if v_card_candidate_count <> 1
    or v_card_base_snapshot->>'docNumber' <> 'FT <DETAIL/1>'
    or v_card_base_snapshot->>'description' <> 'Base & detail'
    or v_card_base_snapshot->>'supplierName' <> 'Supplier <Detail>'
    or v_card_base_snapshot->>'supplierNif' <> 'PT500000001'
    or v_card_candidates->0->>'description' <> 'Card & detail' then
    raise exception 'New Credit Card amount-only snapshots do not contain the proposal display details.';
  end if;

  v_result := public.execute_financial_reconciliation_automatic_proposal(
    '81000000-0000-0000-0000-000000000201', 'smoke:proposal-details-execute'
  );
  if v_result->>'status' <> 'completed'
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.id = '81000000-0000-0000-0000-000000000201'
        and proposal.status = 'completed'
        and proposal.reconciliation_id is not null
    ) then
    raise exception 'Backfilled unique Bank proposal became stale instead of completing: %', v_result;
  end if;

  v_result := public.execute_financial_reconciliation_automatic_proposal(
    '81000000-0000-0000-0000-000000000204', 'smoke:proposal-details-execute'
  );
  if v_result->>'status' <> 'completed'
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.id = '81000000-0000-0000-0000-000000000204'
        and proposal.status = 'completed'
        and proposal.reconciliation_id is not null
    ) then
    raise exception 'Backfilled unique Credit Card proposal became stale instead of completing: %', v_result;
  end if;

  v_result := public.execute_financial_reconciliation_automatic_proposal(
    '81000000-0000-0000-0000-000000000207', 'smoke:proposal-details-reapply-stale'
  );
  if v_result->>'status' <> 'stale'
    or v_result->>'reason' <> 'source_snapshot_changed'
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.id = '81000000-0000-0000-0000-000000000207'
        and proposal.status = 'stale'
        and proposal.reason = 'source_snapshot_changed'
        and proposal.reconciliation_id is null
        and proposal.completed_at is null
    ) then
    raise exception 'Migration reapply masked changed-source stale semantics: %', v_result;
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

-- Card Payments - POS - Income catalog, immutable memberships, and source-rule guard.
create temporary table pos_income_existing_priorities as
select rule_key, priority
from public.financial_reconciliation_automatic_rule_configs;

create temporary table pos_income_expected_priority as
select coalesce(max(priority), 0) + 1 as priority
from public.financial_reconciliation_automatic_rule_configs;

create temporary table pos_income_source_rule_rpc_owner as
select procedure.proowner
from pg_proc procedure
where procedure.oid =
  'public.replace_financial_reconciliation_source_rules(jsonb)'::regprocedure;

create temporary table pos_income_task3_four_rule_output_baseline as
select run.id as run_id,
       public.get_financial_reconciliation_automatic_run(run.id) as detail
from public.financial_reconciliation_automatic_runs run
where exists (
  select 1
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.run_id = run.id
    and (proposal.rule_key, proposal.rule_version) in (
      ('financial_documents_cgd_bank_statement', 2),
      ('financial_documents_cgd_credit_card', 1),
      ('financial_documents_cgd_bank_statement_amount_only', 1),
      ('financial_documents_cgd_credit_card_amount_only', 1)
    )
);

create temporary table pos_income_task3_four_rule_contract_baseline as
select expected.rule_key,
       expected.rule_version,
       public.financial_reconciliation_automatic_rule_contract(
         expected.rule_key,
         expected.rule_version
       ) as contract
from (values
  ('financial_documents_cgd_bank_statement', 2),
  ('financial_documents_cgd_credit_card', 1),
  ('financial_documents_cgd_bank_statement_amount_only', 1),
  ('financial_documents_cgd_credit_card_amount_only', 1)
) expected(rule_key, rule_version);

\ir ../supabase-migrations/2026-08-22-financial-reconciliation-automation-pos-income.sql

create or replace function pg_temp.pos_income_task3_normalized_run(p_run_id uuid)
returns jsonb
language sql
stable
set search_path = public, pg_temp
as $$
  select jsonb_set(
    detail,
    '{proposals}',
    coalesce((
      select jsonb_agg(
        proposal.value - 'groupingKey' - 'summarySnapshot'
        order by proposal.ordinality
      )
      from jsonb_array_elements(detail->'proposals')
        with ordinality proposal(value, ordinality)
    ), '[]'::jsonb)
  )
  from (
    select public.get_financial_reconciliation_automatic_run(p_run_id) as detail
  ) current_run
$$;

do $$
declare
  v_baseline record;
begin
  if exists (
    select 1
    from pos_income_task3_four_rule_contract_baseline baseline
    where public.financial_reconciliation_automatic_rule_contract(
            baseline.rule_key,
            baseline.rule_version
          ) is distinct from baseline.contract
  ) then
    raise exception 'Installing POS income analysis changed an existing dispatcher contract.';
  end if;

  for v_baseline in
    select * from pos_income_task3_four_rule_output_baseline
  loop
    if pg_temp.pos_income_task3_normalized_run(v_baseline.run_id)
        is distinct from v_baseline.detail then
      raise exception 'Installing POS income analysis changed an existing four-rule run response for %.',
        v_baseline.run_id;
    end if;
  end loop;
end $$;

do $$
declare
  v_definition jsonb := jsonb_build_object(
    'matchingMode', 'monthly_aggregate',
    'sourceDescriptionPattern', '%POS VENDAS%',
    'destinationAccount', 'Credit Card',
    'destinationExcludedCategory', 'TransferOutToAccount',
    'calendarGrouping', 'closed_month',
    'fixedMaxDifferenceDays', 31,
    'eligibilityFloor', '2026-01-01',
    'requiresNonNullAmount', true
  );
begin
  if (select count(*)
      from public.financial_reconciliation_automatic_rule_definitions
      where rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income'
        and version = 2) <> 1
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_definitions definition
      where definition.rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income'
        and definition.version = 2
        and definition.display_name = 'Card Payments - POS - Income'
        and definition.base_source_type = 'import_cgd_extrato_ordem'
        and definition.destination_source_types = '["import_fdm_accounts"]'::jsonb
        and definition.definition = v_definition
    ) then
    raise exception 'POS income immutable definition differs from the approved v2 literal.';
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs config
    cross join pos_income_expected_priority expected
    where config.rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income'
      and config.rule_version = 2
      and not config.enabled
      and not config.allow_manual_execution
      and not config.include_in_scheduled_batch
      and config.difference_allowed = 7500.00
      and config.max_difference_days = 31
      and config.priority = expected.priority
  ) then
    raise exception 'POS income config defaults or stable next priority are invalid.';
  end if;

  if exists (
    select 1
    from pos_income_existing_priorities expected
    join public.financial_reconciliation_automatic_rule_configs config
      using (rule_key)
    where config.priority is distinct from expected.priority
  ) then
    raise exception 'Installing POS income changed an existing managed rule priority.';
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_source_rules source_rule
    where source_rule.base_source_type = 'import_cgd_extrato_ordem'
      and source_rule.matching_source_type = 'import_fdm_accounts'
      and source_rule.operator = '-'
  ) then
    raise exception 'POS income directional source rule is not Bank Statement to FDM Accounts (-).';
  end if;
end $$;

do $$
declare
  v_proposal_before jsonb;
  v_rejected boolean;
  v_rules jsonb;
begin
  if not exists (
      select 1
      from information_schema.columns column_row
      where column_row.table_schema = 'public'
        and column_row.table_name = 'financial_reconciliation_automatic_proposals'
        and column_row.column_name = 'grouping_key'
        and column_row.data_type = 'text'
        and column_row.is_nullable = 'YES'
    ) or not exists (
      select 1
      from information_schema.columns column_row
      where column_row.table_schema = 'public'
        and column_row.table_name = 'financial_reconciliation_automatic_proposals'
        and column_row.column_name = 'summary_snapshot'
        and column_row.data_type = 'jsonb'
        and column_row.is_nullable = 'NO'
        and column_row.column_default = '''{}''::jsonb'
  ) then
    raise exception 'POS income proposal columns do not match the monthly-only schema.';
  end if;

  if not exists (
    select 1
    from pg_constraint constraint_row
    where constraint_row.conrelid =
        'public.financial_reconciliation_automatic_proposals'::regclass
      and constraint_row.conname =
        'financial_reconciliation_proposals_summary_snapshot_check'
      and constraint_row.contype = 'c'
      and constraint_row.convalidated
  ) then
    raise exception 'POS income proposal summary object constraint is missing or unvalidated.';
  end if;

  v_rejected := false;
  begin
    insert into public.financial_reconciliation_automatic_runs (
      id, trigger, scope, status, actor, client_request_id
    ) values (
      '82000000-0000-0000-0000-000000000001', 'manual', 'rule', 'ready',
      'smoke:pos-income-memberships',
      '82000000-0000-0000-0000-000000000002'
    );
    insert into public.financial_reconciliation_automatic_proposals (
      id, run_id, rule_key, rule_version, base_source_type, base_source_id,
      base_source_date, base_snapshot, allowed_difference, signature,
      grouping_key, summary_snapshot
    ) values (
      '82000000-0000-0000-0000-000000000003',
      '82000000-0000-0000-0000-000000000001',
      'cgd_bank_statement_fdm_credit_card_monthly_income', 2,
      'import_cgd_extrato_ordem',
      '82000000-0000-0000-0000-000000000004', date '2026-01-01',
      '{}'::jsonb, 7500.00, 'smoke:pos-income:2026-01',
      '2026-01', '{"calendarMonth":"2026-01"}'::jsonb
    );
  exception when others then
    raise exception 'Could not create POS income membership fixture: %', sqlerrm;
  end;

  select to_jsonb(proposal)
  into strict v_proposal_before
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.id = '82000000-0000-0000-0000-000000000003';

  v_rejected := false;
  begin
    update public.financial_reconciliation_automatic_proposals
    set grouping_key = '2026-02'
    where id = '82000000-0000-0000-0000-000000000003';
  exception when others then
    v_rejected := sqlerrm =
      'Automatic proposal monthly audit snapshot is immutable.';
  end;
  if not v_rejected or exists (
    select 1
    from public.financial_reconciliation_automatic_proposals proposal
    where proposal.id = '82000000-0000-0000-0000-000000000003'
      and to_jsonb(proposal) is distinct from v_proposal_before
  ) then
    raise exception 'POS income grouping key update changed the stored proposal value or timestamps.';
  end if;

  v_rejected := false;
  begin
    update public.financial_reconciliation_automatic_proposals
    set summary_snapshot = '{"calendarMonth":"2026-02"}'::jsonb
    where id = '82000000-0000-0000-0000-000000000003';
  exception when others then
    v_rejected := sqlerrm =
      'Automatic proposal monthly audit snapshot is immutable.';
  end;
  if not v_rejected or exists (
    select 1
    from public.financial_reconciliation_automatic_proposals proposal
    where proposal.id = '82000000-0000-0000-0000-000000000003'
      and to_jsonb(proposal) is distinct from v_proposal_before
  ) then
    raise exception 'POS income summary snapshot update changed the stored proposal value or timestamps.';
  end if;

  insert into public.financial_reconciliation_automatic_proposal_memberships (
    proposal_id, role, source_type, source_id, ordinal, source_date,
    amount, description, account, row_snapshot
  ) values (
    '82000000-0000-0000-0000-000000000003', 'source',
    'import_cgd_extrato_ordem', '82000000-0000-0000-0000-000000000005',
    1, date '2026-01-02', 100.00, 'POS VENDAS', '',
    '{"sourceType":"import_cgd_extrato_ordem"}'::jsonb
  );

  v_rejected := false;
  begin
    insert into public.financial_reconciliation_automatic_proposal_memberships (
      proposal_id, role, source_type, source_id, ordinal, source_date,
      amount, row_snapshot
    ) values (
      '82000000-0000-0000-0000-000000000003', 'other',
      'import_fdm_accounts', '82000000-0000-0000-0000-000000000006',
      2, date '2026-01-03', 100.00, '{}'::jsonb
    );
  exception when check_violation then
    v_rejected := true;
  end;
  if not v_rejected then
    raise exception 'POS income memberships accepted an invalid role.';
  end if;

  v_rejected := false;
  begin
    insert into public.financial_reconciliation_automatic_proposal_memberships (
      proposal_id, role, source_type, source_id, ordinal, source_date,
      amount, row_snapshot
    ) values (
      '82000000-0000-0000-0000-000000000003', 'source',
      'import_fdm_accounts', '82000000-0000-0000-0000-000000000007',
      1, date '2026-01-03', 100.00, '{}'::jsonb
    );
  exception when unique_violation then
    v_rejected := true;
  end;
  if not v_rejected then
    raise exception 'POS income memberships accepted a duplicate role ordinal.';
  end if;

  v_rejected := false;
  begin
    insert into public.financial_reconciliation_automatic_proposal_memberships (
      proposal_id, role, source_type, source_id, ordinal, source_date,
      amount, row_snapshot
    ) values (
      '82000000-0000-0000-0000-000000000003', 'destination',
      'import_cgd_extrato_ordem', '82000000-0000-0000-0000-000000000005',
      2, date '2026-01-03', 100.00, '{}'::jsonb
    );
  exception when unique_violation then
    v_rejected := true;
  end;
  if not v_rejected then
    raise exception 'POS income memberships accepted a duplicate source membership.';
  end if;

  v_rejected := false;
  begin
    insert into public.financial_reconciliation_automatic_proposal_memberships (
      proposal_id, role, source_type, source_id, ordinal, source_date,
      amount, row_snapshot
    ) values (
      '82000000-0000-0000-0000-000000000003', 'destination',
      'import_fdm_accounts', '82000000-0000-0000-0000-000000000008',
      2, date '2026-01-03', 100.00, '[]'::jsonb
    );
  exception when check_violation then
    v_rejected := true;
  end;
  if not v_rejected then
    raise exception 'POS income memberships accepted a non-object row snapshot.';
  end if;

  v_rejected := false;
  begin
    update public.financial_reconciliation_automatic_proposal_memberships
    set amount = 101.00
    where proposal_id = '82000000-0000-0000-0000-000000000003'
      and source_id = '82000000-0000-0000-0000-000000000005';
  exception when others then
    v_rejected := sqlerrm = 'Automatic proposal memberships are immutable.';
  end;
  if not v_rejected then
    raise exception 'POS income membership snapshot update was accepted.';
  end if;

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
        when rule->>'base_source_type' = 'import_cgd_extrato_ordem'
         and rule->>'matching_source_type' = 'import_fdm_accounts'
          then jsonb_set(rule, '{operator}', '"+"'::jsonb)
        else rule
      end)
      from jsonb_array_elements(v_rules) rule
    ));
  exception when others then
    v_rejected := sqlerrm =
      'The managed POS income source rule must remain enabled with operator -.';
  end;
  if not v_rejected then
    raise exception 'Managed POS income source-rule operator change was accepted.';
  end if;

  v_rejected := false;
  begin
    perform public.replace_financial_reconciliation_source_rules((
      select coalesce(jsonb_agg(rule), '[]'::jsonb)
      from jsonb_array_elements(v_rules) rule
      where not (
        rule->>'base_source_type' = 'import_cgd_extrato_ordem'
        and rule->>'matching_source_type' = 'import_fdm_accounts'
      )
    ));
  exception when others then
    v_rejected := sqlerrm =
      'The managed POS income source rule must remain enabled with operator -.';
  end;
  if not v_rejected then
    raise exception 'Managed POS income source-rule deletion was accepted.';
  end if;
end $$;

do $$
declare
  v_index_name text;
  v_signature text :=
    'public.replace_financial_reconciliation_source_rules(jsonb)';
begin
  if not (
    select table_row.relrowsecurity
    from pg_class table_row
    where table_row.oid =
      'public.financial_reconciliation_automatic_proposal_memberships'::regclass
  ) then
    raise exception 'POS income membership RLS is not enabled.';
  end if;

  if has_table_privilege(
      'anon',
      'public.financial_reconciliation_automatic_proposal_memberships',
      'SELECT,INSERT,UPDATE,DELETE'
    ) or has_table_privilege(
      'authenticated',
      'public.financial_reconciliation_automatic_proposal_memberships',
      'SELECT,INSERT,UPDATE,DELETE'
    ) or has_table_privilege(
      'service_role',
      'public.financial_reconciliation_automatic_proposal_memberships',
      'SELECT,INSERT,UPDATE,DELETE'
  ) then
    raise exception 'POS income membership table retains direct API-role privileges.';
  end if;

  if has_function_privilege('anon', v_signature, 'EXECUTE')
    or has_function_privilege('authenticated', v_signature, 'EXECUTE')
    or not has_function_privilege('service_role', v_signature, 'EXECUTE')
    or not (
      select procedure.prosecdef
        and coalesce(procedure.proconfig, '{}'::text[])
          @> array['search_path=public, pg_temp']
      from pg_proc procedure
      where procedure.oid = v_signature::regprocedure
    ) then
    raise exception 'POS income source-rule RPC security is invalid.';
  end if;

  if not exists (
    select 1
    from pg_proc procedure
    cross join pos_income_source_rule_rpc_owner expected
    where procedure.oid = v_signature::regprocedure
      and procedure.proowner = expected.proowner
  ) then
    raise exception 'POS income source-rule RPC owner changed during replacement.';
  end if;

  if has_function_privilege(
      'anon',
      'public.prevent_financial_reconciliation_automatic_membership_change()',
      'EXECUTE'
    ) or has_function_privilege(
      'authenticated',
      'public.prevent_financial_reconciliation_automatic_membership_change()',
      'EXECUTE'
    ) or has_function_privilege(
      'service_role',
      'public.prevent_financial_reconciliation_automatic_membership_change()',
      'EXECUTE'
    ) or has_function_privilege(
      'anon',
      'public.prevent_financial_reconciliation_monthly_snapshot_change()',
      'EXECUTE'
    ) or has_function_privilege(
      'authenticated',
      'public.prevent_financial_reconciliation_monthly_snapshot_change()',
      'EXECUTE'
    ) or has_function_privilege(
      'service_role',
      'public.prevent_financial_reconciliation_monthly_snapshot_change()',
      'EXECUTE'
  ) then
    raise exception 'POS income immutable-snapshot trigger function is directly executable by an API role.';
  end if;

  foreach v_index_name in array array[
    'financial_reconciliation_automatic_memberships_role_ordinal_idx',
    'import_cgd_extrato_ordem_pos_income_lock_idx',
    'import_fdm_accounts_credit_card_eligible_v2_lock_idx'
  ] loop
    if to_regclass('public.' || v_index_name) is null then
      raise exception 'POS income required index is missing: %.', v_index_name;
    end if;
  end loop;
end $$;

update public.financial_reconciliation_automatic_rule_configs
set enabled = true,
    allow_manual_execution = true,
    include_in_scheduled_batch = true,
    difference_allowed = 4321.00,
    priority = (select max(priority) + 10
                from public.financial_reconciliation_automatic_rule_configs),
    updated_by = 'smoke:pos-income-administrator'
where rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income';

create or replace function pg_temp.pos_income_task2_state()
returns jsonb
language sql
stable
set search_path = public, pg_temp
as $$
  select jsonb_build_object(
    'definitions', (
      select jsonb_agg(to_jsonb(definition)
                       order by definition.rule_key, definition.version)
      from public.financial_reconciliation_automatic_rule_definitions definition
    ),
    'configs', (
      select jsonb_agg(to_jsonb(config) order by config.rule_key)
      from public.financial_reconciliation_automatic_rule_configs config
    ),
    'proposalColumns', (
      select jsonb_agg(to_jsonb(column_row) order by column_row.column_name)
      from information_schema.columns column_row
      where column_row.table_schema = 'public'
        and column_row.table_name = 'financial_reconciliation_automatic_proposals'
        and column_row.column_name in ('grouping_key', 'summary_snapshot')
    ),
    'constraints', (
      select jsonb_agg(jsonb_build_object(
        'table', constraint_row.conrelid::regclass::text,
        'name', constraint_row.conname,
        'type', constraint_row.contype,
        'validated', constraint_row.convalidated,
        'definition', pg_get_constraintdef(constraint_row.oid, true)
      ) order by constraint_row.conrelid::regclass::text, constraint_row.conname)
      from pg_constraint constraint_row
      where constraint_row.conrelid in (
        'public.financial_reconciliation_automatic_proposal_memberships'::regclass,
        'public.financial_reconciliation_automatic_proposals'::regclass,
        'public.financial_reconciliation_automatic_rule_configs'::regclass
      )
        and (
          constraint_row.conrelid =
            'public.financial_reconciliation_automatic_proposal_memberships'::regclass
          or constraint_row.conname in (
            'financial_reconciliation_proposals_summary_snapshot_check',
            'financial_reconciliation_rule_configs_pos_income_days_check'
          )
        )
    ),
    'indexes', (
      select jsonb_agg(jsonb_build_object(
        'name', index_row.indexrelid::regclass::text,
        'definition', pg_get_indexdef(index_row.indexrelid),
        'predicate', pg_get_expr(index_row.indpred, index_row.indrelid)
      ) order by index_row.indexrelid::regclass::text)
      from pg_index index_row
      where index_row.indexrelid in (
        'public.financial_reconciliation_automatic_memberships_role_ordinal_idx'::regclass,
        'public.import_cgd_extrato_ordem_pos_income_lock_idx'::regclass,
        'public.import_fdm_accounts_credit_card_eligible_v2_lock_idx'::regclass
      )
    ),
    'functions', (
      select jsonb_agg(jsonb_build_object(
        'signature', procedure.oid::regprocedure::text,
        'definition', pg_get_functiondef(procedure.oid),
        'owner', procedure.proowner::regrole::text,
        'acl', coalesce(procedure.proacl::text, ''),
        'config', coalesce(to_jsonb(procedure.proconfig), '[]'::jsonb)
      ) order by procedure.oid::regprocedure::text)
      from pg_proc procedure
      where procedure.oid in (
        'public.replace_financial_reconciliation_source_rules(jsonb)'::regprocedure,
        'public.prevent_financial_reconciliation_automatic_membership_change()'::regprocedure,
        'public.prevent_financial_reconciliation_monthly_snapshot_change()'::regprocedure
      )
    ),
    'triggers', (
      select jsonb_agg(pg_get_triggerdef(trigger_row.oid, true)
                       order by trigger_row.tgrelid::regclass::text,
                                trigger_row.tgname)
      from pg_trigger trigger_row
      where trigger_row.tgrelid in (
          'public.financial_reconciliation_automatic_proposal_memberships'::regclass,
          'public.financial_reconciliation_automatic_proposals'::regclass
        )
        and not trigger_row.tgisinternal
    ),
    'tableAcl', (
      select coalesce(table_row.relacl::text, '')
      from pg_class table_row
      where table_row.oid =
        'public.financial_reconciliation_automatic_proposal_memberships'::regclass
    ),
    'rowSecurity', (
      select table_row.relrowsecurity
      from pg_class table_row
      where table_row.oid =
        'public.financial_reconciliation_automatic_proposal_memberships'::regclass
    ),
    'sourceRules', (
      select jsonb_agg(to_jsonb(source_rule)
                       order by base_source_type, matching_source_type)
      from public.financial_reconciliation_source_rules source_rule
    ),
    'memberships', (
      select jsonb_agg(to_jsonb(membership)
                       order by role, ordinal, source_id)
      from public.financial_reconciliation_automatic_proposal_memberships membership
      where membership.proposal_id =
        '82000000-0000-0000-0000-000000000003'
    ),
    'proposal', (
      select to_jsonb(proposal)
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.id =
        '82000000-0000-0000-0000-000000000003'
    )
  )
$$;

create temporary table pos_income_task2_reapply_baseline as
select pg_temp.pos_income_task2_state() as state;

\ir ../supabase-migrations/2026-08-22-financial-reconciliation-automation-pos-income.sql

do $$
begin
  if (select state from pos_income_task2_reapply_baseline)
      is distinct from pg_temp.pos_income_task2_state() then
    raise exception 'POS income migration second apply changed schema, catalog, functions, privileges, or data.';
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs config
    where config.rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income'
      and config.rule_version = 2
      and config.enabled
      and config.allow_manual_execution
      and config.include_in_scheduled_batch
      and config.difference_allowed = 4321.00
      and config.max_difference_days = 31
      and config.updated_by = 'smoke:pos-income-administrator'
  ) then
    raise exception 'POS income migration reapply overwrote administrator configuration.';
  end if;
end $$;

savepoint pos_income_conflicting_index_fixture;

drop index public.import_cgd_extrato_ordem_pos_income_lock_idx;
create index import_cgd_extrato_ordem_pos_income_lock_idx
  on public.import_cgd_extrato_ordem (data, id)
  with (fillfactor = 70)
  where descritivo ilike '%POS VENDAS%';

\set ON_ERROR_STOP off
\ir ../supabase-migrations/2026-08-22-financial-reconciliation-automation-pos-income.sql
\set pos_income_conflicting_index_rejected :ERROR
\set ON_ERROR_STOP on

rollback to savepoint pos_income_conflicting_index_fixture;

\if :pos_income_conflicting_index_rejected
\else
  \echo 'POS income migration accepted a same-named index with incompatible storage options.'
  \quit 1
\endif

\ir ../supabase-migrations/2026-08-22-financial-reconciliation-automation-pos-income.sql

do $$
begin
  if (select state from pos_income_task2_reapply_baseline)
      is distinct from pg_temp.pos_income_task2_state() then
    raise exception 'POS income migration did not restore and reapply safely after the conflicting-index fixture.';
  end if;
end $$;

-- Card Payments - POS - Income closed-month analysis and serialization.
update public.financial_reconciliation_automatic_rule_configs
set enabled = true,
    allow_manual_execution = true,
    include_in_scheduled_batch = true,
    difference_allowed = 7500.00,
    updated_by = 'smoke:pos-income-analysis'
where rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income';

insert into public.import_cgd_extrato_ordem (
  id, import_batch, row_key, data, descritivo, montante
) values
  ('83000000-0000-0000-0000-000000000001', 'smoke-pos-income-analysis',
   'pos-income-pre-floor', date '2025-12-31', 'POS VENDAS pre-floor', 25.00),
  ('83000000-0000-0000-0000-000000000002', 'smoke-pos-income-analysis',
   'pos-income-leap-floor', date '2024-02-29', 'POS VENDAS leap-floor', 25.00),
  ('83000000-0000-0000-0000-000000000010', 'smoke-pos-income-analysis',
   'pos-income-jan-a', date '2026-01-05', 'prefix pos vendas suffix', 4000.00),
  ('83000000-0000-0000-0000-000000000011', 'smoke-pos-income-analysis',
   'pos-income-jan-b', date '2026-01-05', 'POS VENDAS January', 3600.00),
  ('83000000-0000-0000-0000-000000000020', 'smoke-pos-income-analysis',
   'pos-income-feb', date '2026-02-28', 'POS VENDAS February', 7600.01),
  ('83000000-0000-0000-0000-000000000040', 'smoke-pos-income-analysis',
   'pos-income-source-only', date '2026-04-15', 'POS VENDAS source-only', 10.00),
  ('83000000-0000-0000-0000-000000000060', 'smoke-pos-income-analysis',
   'pos-income-jun-good', date '2026-06-05', 'xx PoS VeNdAs yy', 75.00),
  ('83000000-0000-0000-0000-000000000061', 'smoke-pos-income-analysis',
   'pos-income-jun-wrong-description', date '2026-06-06', 'POS VENDA', 500.00),
  ('83000000-0000-0000-0000-000000000062', 'smoke-pos-income-analysis',
   'pos-income-jun-locked', date '2026-06-07', 'POS VENDAS locked', 600.00),
  ('83000000-0000-0000-0000-000000000063', 'smoke-pos-income-analysis',
   'pos-income-jun-null-amount', date '2026-06-08',
   'POS VENDAS null amount', null),
  ('83000000-0000-0000-0000-000000000090', 'smoke-pos-income-analysis',
   'pos-income-current-month', date_trunc('month', current_date)::date + 1,
   'POS VENDAS current-month', 700.00);

-- The production FDM schema rejects null amounts at ingestion. Temporarily relax
-- that guard inside this rollback-only smoke transaction to prove the analysis
-- layer still treats a malformed legacy row as ineligible rather than failing.
alter table public.import_fdm_accounts alter column amount drop not null;

insert into public.import_fdm_accounts (
  id, import_batch, account, date_time_raw, event_date, category,
  amount, description
) values
  ('84000000-0000-0000-0000-000000000001', 'smoke-pos-income-analysis',
   'Credit Card', '2025-12-31', date '2025-12-31', 'POS income',
   25.00, 'pre-floor'),
  ('84000000-0000-0000-0000-000000000002', 'smoke-pos-income-analysis',
   'Credit Card', '2024-02-29', date '2024-02-29', 'POS income',
   25.00, 'leap-floor'),
  ('84000000-0000-0000-0000-000000000010', 'smoke-pos-income-analysis',
   'Credit Card', '2026-01-31', date '2026-01-31', 'POS income',
   100.00, 'January'),
  ('84000000-0000-0000-0000-000000000020', 'smoke-pos-income-analysis',
   'Credit Card', '2026-02-01', date '2026-02-01', 'POS income',
   100.00, 'February'),
  ('84000000-0000-0000-0000-000000000050', 'smoke-pos-income-analysis',
   'Credit Card', '2026-05-15', date '2026-05-15', 'POS income',
   10.00, 'destination-only'),
  ('84000000-0000-0000-0000-000000000060', 'smoke-pos-income-analysis',
   'Credit Card', '2026-06-05', date '2026-06-05', null,
   75.00, 'June eligible'),
  ('84000000-0000-0000-0000-000000000061', 'smoke-pos-income-analysis',
   'credit card', '2026-06-06', date '2026-06-06', 'POS income',
   500.00, 'wrong account case'),
  ('84000000-0000-0000-0000-000000000062', 'smoke-pos-income-analysis',
   'Credit Card', '2026-06-07', date '2026-06-07', 'POS income',
   600.00, 'locked destination'),
  ('84000000-0000-0000-0000-000000000063', 'smoke-pos-income-analysis',
   'Credit Card', '2026-06-08', date '2026-06-08', 'POS income',
   null, 'null amount'),
  ('84000000-0000-0000-0000-000000000064', 'smoke-pos-income-analysis',
   'Credit Card', '2026-06-09', date '2026-06-09', 'TransferOutToAccount',
   999.00, 'excluded transfer-out category'),
  ('84000000-0000-0000-0000-000000000090', 'smoke-pos-income-analysis',
   'Credit Card', current_date::text,
   date_trunc('month', current_date)::date + 1, 'POS income',
   700.00, 'current-month');

insert into public.import_cgd_extrato_ordem (
  id, import_batch, row_key, source_row_number, data, descritivo, montante
)
select
  ('83100000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
  'smoke-pos-income-large',
  'pos-income-large-bank-' || series,
  series,
  date '2026-03-01' + ((series - 1) % 31),
  'bulk POS VENDAS row ' || series,
  2.00
from generate_series(1, 1000) series;

insert into public.import_fdm_accounts (
  id, import_batch, source_row_number, account, date_time_raw, event_date,
  category, amount, description
)
select
  ('84100000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
  'smoke-pos-income-large',
  series,
  'Credit Card',
  '2026-03-' || lpad((((series - 1) % 31) + 1)::text, 2, '0'),
  date '2026-03-01' + ((series - 1) % 31),
  'POS income',
  1.25,
  'bulk FDM row ' || series
from generate_series(1, 1000) series;

do $$
declare
  v_reconciliation_id uuid;
begin
  insert into public.financial_reconciliations (
    status, base_source_type, matching_source_types, created_by
  ) values (
    'started', 'import_cgd_extrato_ordem', '["import_fdm_accounts"]'::jsonb,
    'smoke:pos-income-analysis-locks'
  ) returning id into v_reconciliation_id;

  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  ) values
    (v_reconciliation_id, 'import_cgd_extrato_ordem',
     '83000000-0000-0000-0000-000000000062', 600.00,
     'smoke:pos-income-analysis-locks'),
    (v_reconciliation_id, 'import_fdm_accounts',
     '84000000-0000-0000-0000-000000000062', 600.00,
     'smoke:pos-income-analysis-locks');
end $$;

do $$
declare
  v_contract jsonb;
  v_june record;
  v_page_one date[];
  v_page_two date[];
  v_page_three date[];
  v_rejected boolean := false;
  v_signature text;
begin
  v_contract := public.financial_reconciliation_automatic_rule_contract(
    'cgd_bank_statement_fdm_credit_card_monthly_income', 2
  );
  if v_contract is distinct from jsonb_build_object(
      'destinationSourceType', 'import_fdm_accounts',
      'matchingMode', 'monthly_aggregate',
      'sourceDescriptionPattern', '%POS VENDAS%',
      'destinationAccount', 'Credit Card',
      'destinationExcludedCategory', 'TransferOutToAccount',
      'calendarGrouping', 'closed_month',
      'eligibilityFloor', '2026-01-01',
      'fixedMaxDifferenceDays', 31,
      'requiresNonNullAmount', true
    ) then
    raise exception 'POS income dispatcher contract is not the approved literal: %',
      v_contract;
  end if;

  if public.financial_reconciliation_automatic_monthly_income_count() <> 4 then
    raise exception 'POS income month count admitted a floor, current-month, locked, null-amount, predicate, or one-sided row.';
  end if;

  select page.* into strict v_june
  from public.financial_reconciliation_automatic_monthly_income_page(
    date '2026-03-01', 25
  ) page
  where page.calendar_month = date '2026-06-01';
  if v_june.source_count <> 1
    or v_june.source_total <> 75.00
    or v_june.destination_count <> 1
    or v_june.destination_total <> 75.00
    or v_june.calculated_difference <> 0.00 then
    raise exception 'POS income category exclusion or null eligibility changed June counts or totals: %',
      to_jsonb(v_june);
  end if;

  select array_agg(page.calendar_month order by page.calendar_month)
  into v_page_one
  from public.financial_reconciliation_automatic_monthly_income_page(null, 2) page;
  select array_agg(page.calendar_month order by page.calendar_month)
  into v_page_two
  from public.financial_reconciliation_automatic_monthly_income_page(
    date '2026-02-01', 2
  ) page;
  select array_agg(page.calendar_month order by page.calendar_month)
  into v_page_three
  from public.financial_reconciliation_automatic_monthly_income_page(
    date '2026-06-01', 2
  ) page;
  if v_page_one is distinct from array[date '2026-01-01', date '2026-02-01']
    or v_page_two is distinct from array[date '2026-03-01', date '2026-06-01']
    or coalesce(v_page_three, '{}'::date[]) <> '{}'::date[] then
    raise exception 'POS income month paging is not stable and oldest first: %, %, %',
      v_page_one, v_page_two, v_page_three;
  end if;

  v_rejected := false;
  begin
    perform public.financial_reconciliation_automatic_monthly_income_page(null, 26);
  exception when others then
    v_rejected := sqlerrm =
      'Automatic monthly analysis page size must be between 1 and 25.';
  end;
  if not v_rejected then
    raise exception 'POS income monthly page accepted an oversized limit.';
  end if;

  v_rejected := false;
  begin
    perform public.financial_reconciliation_automatic_monthly_income_page(null, null);
  exception when others then
    v_rejected := sqlerrm =
      'Automatic monthly analysis page size must be between 1 and 25.';
  end;
  if not v_rejected then
    raise exception 'POS income monthly page accepted a null limit.';
  end if;

  foreach v_signature in array array[
    'public.financial_reconciliation_automatic_rule_contract(text,integer)',
    'public.financial_reconciliation_automatic_monthly_income_count()',
    'public.financial_reconciliation_automatic_monthly_income_page(date,integer)',
    'public.continue_financial_reconciliation_automatic_analysis(uuid,text)',
    'public.create_financial_reconciliation_automatic_analysis(text[],text,text,uuid)',
    'public.financial_reconciliation_finalize_automatic_analysis(uuid)',
    'public.financial_reconciliation_automatic_progress_or_run(uuid)',
    'public.get_financial_reconciliation_automatic_run(uuid)'
  ] loop
    if has_function_privilege('anon', v_signature, 'EXECUTE')
      or has_function_privilege('authenticated', v_signature, 'EXECUTE')
      or not has_function_privilege('service_role', v_signature, 'EXECUTE')
      or not (
        select procedure.prosecdef
          and coalesce(procedure.proconfig, '{}'::text[])
            @> array['search_path=public, pg_temp']
        from pg_proc procedure
        where procedure.oid = v_signature::regprocedure
      ) then
      raise exception 'POS income analysis function security is invalid for %.',
        v_signature;
    end if;
  end loop;

  v_signature :=
    'public.financial_reconciliation_continue_automatic_monthly_income(uuid,jsonb)';
  if has_function_privilege('anon', v_signature, 'EXECUTE')
    or has_function_privilege('authenticated', v_signature, 'EXECUTE')
    or has_function_privilege('service_role', v_signature, 'EXECUTE')
    or not (
      select not procedure.prosecdef
        and coalesce(procedure.proconfig, '{}'::text[])
          @> array['search_path=public, pg_temp']
      from pg_proc procedure
      where procedure.oid = v_signature::regprocedure
    ) then
    raise exception 'POS income internal continuation helper is exposed or unsafe.';
  end if;
end $$;

create or replace function pg_temp.reject_pos_income_membership_insert()
returns trigger
language plpgsql
set search_path = public, pg_temp
as $$
begin
  if new.source_id = '84000000-0000-0000-0000-000000000060'::uuid then
    raise exception 'smoke POS income membership insert failure';
  end if;
  return new;
end
$$;

create trigger reject_pos_income_membership_insert
before insert on public.financial_reconciliation_automatic_proposal_memberships
for each row execute function pg_temp.reject_pos_income_membership_insert();

do $$
declare
  v_result jsonb;
  v_run_id uuid;
begin
  v_result := public.create_financial_reconciliation_automatic_analysis(
    array['cgd_bank_statement_fdm_credit_card_monthly_income'],
    'manual_rule',
    'smoke:pos-income-atomicity',
    '85000000-0000-0000-0000-000000000001'
  );
  v_run_id := (v_result->>'runId')::uuid;
  if v_result->>'status' <> 'failed'
    or v_result->>'analysisErrorCode' <> 'analysis_continuation_failed'
    or exists (
      select 1
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = v_run_id
    )
    or exists (
      select 1
      from public.financial_reconciliation_automatic_proposal_memberships membership
      join public.financial_reconciliation_automatic_proposals proposal
        on proposal.id = membership.proposal_id
      where proposal.run_id = v_run_id
    ) then
    raise exception 'POS income proposal and memberships did not roll back atomically: %',
      v_result;
  end if;
end $$;

drop trigger reject_pos_income_membership_insert
on public.financial_reconciliation_automatic_proposal_memberships;
drop function pg_temp.reject_pos_income_membership_insert();

-- Simulate a same-transaction source mutation after the proposal summary is
-- inserted but before its memberships are materialized. Analysis must reassert
-- the exact destination eligibility predicate and roll the whole month back;
-- otherwise it can persist a born-stale mixed snapshot.
create or replace function pg_temp.mutate_pos_income_before_memberships()
returns trigger
language plpgsql
set search_path = public, pg_temp
as $$
begin
  if new.rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income'
    and new.rule_version = 2
    and new.grouping_key = '2026-06'
    and exists (
      select 1
      from public.financial_reconciliation_automatic_runs run
      where run.id = new.run_id
        and run.actor = 'smoke:pos-income-born-stale'
    ) then
    update public.import_fdm_accounts
    set account = 'Cash'
    where id = '84000000-0000-0000-0000-000000000060';
  end if;
  return new;
end
$$;

create trigger mutate_pos_income_before_memberships
after insert on public.financial_reconciliation_automatic_proposals
for each row execute function pg_temp.mutate_pos_income_before_memberships();

do $$
declare
  v_result jsonb;
  v_run_id uuid;
begin
  v_result := public.create_financial_reconciliation_automatic_analysis(
    array['cgd_bank_statement_fdm_credit_card_monthly_income'],
    'manual_rule',
    'smoke:pos-income-born-stale',
    '85000000-0000-0000-0000-000000000003'
  );
  v_run_id := (v_result->>'runId')::uuid;
  if v_result->>'status' <> 'failed'
    or v_result->>'analysisErrorCode' <> 'analysis_continuation_failed'
    or exists (
      select 1
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = v_run_id
    )
    or (select account
        from public.import_fdm_accounts
        where id = '84000000-0000-0000-0000-000000000060')
       is distinct from 'Credit Card' then
    raise exception 'POS income analysis persisted or retained a born-stale source mutation: %',
      v_result;
  end if;
end $$;

drop trigger mutate_pos_income_before_memberships
on public.financial_reconciliation_automatic_proposals;
drop function pg_temp.mutate_pos_income_before_memberships();

create temporary table pos_income_task3_success_run (
  run_id uuid primary key
);

do $$
declare
  v_result jsonb;
  v_retry jsonb;
  v_run_id uuid;
  v_january_proposal_id uuid;
  v_large_proposal_id uuid;
  v_expected_months text[] := array['2026-01','2026-02','2026-03','2026-06'];
  v_actual_months text[];
begin
  v_result := public.create_financial_reconciliation_automatic_analysis(
    array['cgd_bank_statement_fdm_credit_card_monthly_income'],
    'manual_rule',
    'smoke:pos-income-analysis',
    '85000000-0000-0000-0000-000000000002'
  );
  v_run_id := (v_result->>'runId')::uuid;
  insert into pos_income_task3_success_run values (v_run_id);

  if v_result->>'status' <> 'ready'
    or v_result->>'analysisComplete' <> 'true'
    or (v_result->>'analysisProcessed')::integer <> 4
    or (v_result->>'analysisTotal')::integer <> 4
    or v_result->>'analysisCursorDate' <> '2026-06-01'
    or v_result->>'analysisCursorId' <>
       '83000000-0000-0000-0000-000000000060'
    or jsonb_array_length(v_result->'proposals') <> 4 then
    raise exception 'POS income analysis lifecycle or month cursor is invalid: %',
      v_result;
  end if;

  select array_agg(proposal.value->>'groupingKey' order by proposal.ordinality)
  into v_actual_months
  from jsonb_array_elements(v_result->'proposals')
    with ordinality proposal(value, ordinality);
  if v_actual_months is distinct from v_expected_months then
    raise exception 'POS income proposals are not serialized oldest first: %',
      v_actual_months;
  end if;

  if exists (
    select 1
    from jsonb_array_elements(v_result->'proposals') proposal(value)
    where proposal.value->'items' <> '[]'::jsonb
       or proposal.value->'candidateGroups' <> '[]'::jsonb
       or jsonb_typeof(proposal.value->'summarySnapshot') <> 'object'
  ) then
    raise exception 'POS income run detail embedded members or omitted a summary: %',
      v_result->'proposals';
  end if;

  select proposal.id into strict v_january_proposal_id
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.run_id = v_run_id and proposal.grouping_key = '2026-01';
  if not exists (
      select 1
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.id = v_january_proposal_id
        and proposal.base_source_id =
          '83000000-0000-0000-0000-000000000010'
        and proposal.base_source_date = date '2026-01-05'
        and proposal.status = 'proposed'
        and proposal.reason = ''
        and proposal.calculated_difference = 7500.00
        and proposal.allowed_difference = 7500.00
        and proposal.summary_snapshot @> jsonb_build_object(
          'calendarMonth', date '2026-01-01',
          'sourceCount', 2,
          'sourceTotal', 7600.00,
          'destinationCount', 1,
          'destinationTotal', 100.00,
          'calculatedDifference', 7500.00,
          'technicalBaseSourceId',
            '83000000-0000-0000-0000-000000000010'
        )
    ) then
    raise exception 'POS income exact-tolerance or technical-base proposal is invalid.';
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_proposals proposal
    where proposal.run_id = v_run_id
      and proposal.grouping_key = '2026-02'
      and proposal.status = 'ambiguous'
      and proposal.reason = 'monthly_difference_exceeded'
      and proposal.calculated_difference = 7500.01
      and proposal.allowed_difference = 7500.00
  ) then
    raise exception 'POS income tolerance-plus-one-cent boundary is invalid.';
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_proposals proposal
    where proposal.run_id = v_run_id
      and proposal.grouping_key = '2026-06'
      and proposal.status = 'proposed'
      and proposal.reason = ''
      and proposal.calculated_difference = 0.00
      and proposal.summary_snapshot @> jsonb_build_object(
        'calendarMonth', date '2026-06-01',
        'sourceCount', 1,
        'sourceTotal', 75.00,
        'destinationCount', 1,
        'destinationTotal', 75.00,
        'calculatedDifference', 0.00
      )
  ) then
    raise exception 'POS income valid June rows did not produce the null-safe proposal.';
  end if;

  if exists (
    select 1
    from public.financial_reconciliation_automatic_proposals proposal
    where proposal.run_id = v_run_id
      and proposal.grouping_key in ('2025-12','2024-02','2026-04','2026-05')
  ) then
    raise exception 'POS income analysis persisted a floor, leap-floor, or one-sided month.';
  end if;

  select proposal.id into strict v_large_proposal_id
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.run_id = v_run_id and proposal.grouping_key = '2026-03';
  if not exists (
      select 1
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.id = v_large_proposal_id
        and proposal.calculated_difference = 750.00
        and proposal.summary_snapshot @> jsonb_build_object(
          'sourceCount', 1000,
          'sourceTotal', 2000.00,
          'destinationCount', 1000,
          'destinationTotal', 1250.00,
          'calculatedDifference', 750.00
        )
    )
    or (select count(*)
        from public.financial_reconciliation_automatic_proposal_memberships
        where proposal_id = v_large_proposal_id and role = 'source') <> 1000
    or (select count(*)
        from public.financial_reconciliation_automatic_proposal_memberships
        where proposal_id = v_large_proposal_id and role = 'destination') <> 1000
  then
    raise exception 'POS income 1,000-row aggregates or memberships are incomplete.';
  end if;

  if (select count(*)
      from public.financial_reconciliation_automatic_proposal_memberships membership
      join public.financial_reconciliation_automatic_proposals proposal
        on proposal.id = membership.proposal_id
      where proposal.run_id = v_run_id) <> 2007
    or exists (
      select 1
      from (
        select membership.*,
               row_number() over (
                 partition by membership.proposal_id, membership.role
                 order by membership.source_date, membership.source_id
               )::integer as expected_ordinal
        from public.financial_reconciliation_automatic_proposal_memberships membership
        join public.financial_reconciliation_automatic_proposals proposal
          on proposal.id = membership.proposal_id
        where proposal.run_id = v_run_id
      ) ordered_membership
      where ordered_membership.ordinal <> ordered_membership.expected_ordinal
    ) then
    raise exception 'POS income memberships are incomplete, duplicated, or unstably ordered.';
  end if;

  if exists (
    select 1
    from public.financial_reconciliation_automatic_proposal_memberships membership
    join public.financial_reconciliation_automatic_proposals proposal
      on proposal.id = membership.proposal_id
    where proposal.run_id = v_run_id
      and membership.source_id in (
        '83000000-0000-0000-0000-000000000001'::uuid,
        '83000000-0000-0000-0000-000000000002'::uuid,
        '83000000-0000-0000-0000-000000000061'::uuid,
        '83000000-0000-0000-0000-000000000062'::uuid,
        '83000000-0000-0000-0000-000000000063'::uuid,
        '83000000-0000-0000-0000-000000000090'::uuid,
        '84000000-0000-0000-0000-000000000001'::uuid,
        '84000000-0000-0000-0000-000000000002'::uuid,
        '84000000-0000-0000-0000-000000000061'::uuid,
        '84000000-0000-0000-0000-000000000062'::uuid,
        '84000000-0000-0000-0000-000000000063'::uuid,
        '84000000-0000-0000-0000-000000000090'::uuid
      )
  ) then
    raise exception 'POS income memberships admitted an ineligible row.';
  end if;

  v_retry := public.continue_financial_reconciliation_automatic_analysis(
    v_run_id,
    'smoke:pos-income-analysis'
  );
  if v_retry is distinct from v_result
    or (select count(*)
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = v_run_id) <> 4
    or (select count(*)
        from public.financial_reconciliation_automatic_proposal_memberships membership
        join public.financial_reconciliation_automatic_proposals proposal
          on proposal.id = membership.proposal_id
        where proposal.run_id = v_run_id) <> 2007 then
    raise exception 'POS income completed continuation changed output or duplicated proposals/memberships.';
  end if;
end $$;

delete from public.import_fdm_accounts
where id = '84000000-0000-0000-0000-000000000063';
alter table public.import_fdm_accounts alter column amount set not null;

create or replace function pg_temp.pos_income_task3_state(p_run_id uuid)
returns jsonb
language sql
stable
set search_path = public, pg_temp
as $$
  select jsonb_build_object(
    'run', (
      select to_jsonb(run)
      from public.financial_reconciliation_automatic_runs run
      where run.id = p_run_id
    ),
    'detail', public.get_financial_reconciliation_automatic_run(p_run_id),
    'proposals', (
      select jsonb_agg(to_jsonb(proposal)
                       order by proposal.base_source_date,
                                proposal.base_source_id,
                                proposal.signature)
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = p_run_id
    ),
    'memberships', (
      select jsonb_agg(to_jsonb(membership)
                       order by proposal.base_source_date,
                                membership.role,
                                membership.ordinal,
                                membership.source_id)
      from public.financial_reconciliation_automatic_proposal_memberships membership
      join public.financial_reconciliation_automatic_proposals proposal
        on proposal.id = membership.proposal_id
      where proposal.run_id = p_run_id
    )
  )
$$;

create temporary table pos_income_task3_reapply_baseline as
select pg_temp.pos_income_task3_state(run_id) as state
from pos_income_task3_success_run;

\ir ../supabase-migrations/2026-08-22-financial-reconciliation-automation-pos-income.sql

do $$
declare
  v_baseline record;
begin
  if (select state from pos_income_task3_reapply_baseline)
      is distinct from (
        select pg_temp.pos_income_task3_state(run_id)
        from pos_income_task3_success_run
      ) then
    raise exception 'POS income migration reapply changed a run, proposal summary, membership, timestamp, or signature.';
  end if;

  if exists (
    select 1
    from pos_income_task3_four_rule_contract_baseline baseline
    where public.financial_reconciliation_automatic_rule_contract(
            baseline.rule_key,
            baseline.rule_version
          ) is distinct from baseline.contract
  ) then
    raise exception 'POS income migration reapply changed an existing dispatcher contract.';
  end if;

  for v_baseline in
    select * from pos_income_task3_four_rule_output_baseline
  loop
    if pg_temp.pos_income_task3_normalized_run(v_baseline.run_id)
        is distinct from v_baseline.detail then
      raise exception 'POS income migration reapply changed an existing four-rule run response for %.',
        v_baseline.run_id;
    end if;
  end loop;
end $$;

-- Card Payments - POS - Income execution, stale revalidation, rollback, and
-- existing four-rule dispatch regression.
create or replace function pg_temp.pos_income_task4_clone_proposal(
  p_grouping_key text,
  p_run_id uuid,
  p_proposal_id uuid,
  p_actor text
)
returns uuid
language plpgsql
set search_path = public, pg_temp
as $$
declare
  v_source_run_id uuid;
  v_source_proposal_id uuid;
begin
  select proposal.run_id, proposal.id
  into strict v_source_run_id, v_source_proposal_id
  from public.financial_reconciliation_automatic_proposals proposal
  join pos_income_task3_success_run successful
    on successful.run_id = proposal.run_id
  where proposal.grouping_key = p_grouping_key;

  insert into public.financial_reconciliation_automatic_runs (
    id, trigger, scope, status, actor, client_request_id,
    definition_config_snapshot, counts, analysis_completed_at
  )
  select
    p_run_id, 'manual', 'rule', 'ready', p_actor, p_run_id,
    source_run.definition_config_snapshot, source_run.counts, now()
  from public.financial_reconciliation_automatic_runs source_run
  where source_run.id = v_source_run_id;

  insert into public.financial_reconciliation_automatic_proposals (
    id, run_id, rule_key, rule_version, base_source_type, base_source_id,
    base_source_date, base_snapshot, items, evidence, candidate_groups,
    calculated_difference, allowed_difference, status, reason, signature,
    reconciliation_id, error, error_detail, completed_at,
    grouping_key, summary_snapshot
  )
  select
    p_proposal_id, p_run_id, proposal.rule_key, proposal.rule_version,
    proposal.base_source_type, proposal.base_source_id,
    proposal.base_source_date, proposal.base_snapshot, proposal.items,
    proposal.evidence, proposal.candidate_groups,
    proposal.calculated_difference, proposal.allowed_difference,
    proposal.status, proposal.reason, proposal.signature,
    null, '', '', null, proposal.grouping_key, proposal.summary_snapshot
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.id = v_source_proposal_id;

  insert into public.financial_reconciliation_automatic_proposal_memberships (
    proposal_id, role, source_type, source_id, ordinal, source_date,
    amount, description, account, row_snapshot
  )
  select
    p_proposal_id, membership.role, membership.source_type,
    membership.source_id, membership.ordinal, membership.source_date,
    membership.amount, membership.description, membership.account,
    membership.row_snapshot
  from public.financial_reconciliation_automatic_proposal_memberships membership
  where membership.proposal_id = v_source_proposal_id
  order by membership.source_type, membership.source_date, membership.source_id;

  return p_proposal_id;
end
$$;

create or replace function pg_temp.pos_income_task4_assert_stale(
  p_proposal_id uuid,
  p_reason text
)
returns void
language plpgsql
set search_path = public, pg_temp
as $$
declare
  v_actor text;
  v_run_id uuid;
  v_result jsonb;
begin
  select run.actor, run.id
  into strict v_actor, v_run_id
  from public.financial_reconciliation_automatic_proposals proposal
  join public.financial_reconciliation_automatic_runs run
    on run.id = proposal.run_id
  where proposal.id = p_proposal_id;

  v_result := public.execute_financial_reconciliation_automatic_proposal(
    p_proposal_id, v_actor
  );
  if v_result is distinct from jsonb_build_object(
      'proposalId', p_proposal_id,
      'runId', v_run_id,
      'status', 'stale',
      'reason', p_reason
    )
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.id = p_proposal_id
        and proposal.status = 'stale'
        and proposal.reason = p_reason
        and proposal.reconciliation_id is null
        and proposal.completed_at is null
        and proposal.error = ''
        and proposal.error_detail = ''
    )
    or exists (
      select 1
      from public.financial_reconciliations reconciliation
      where reconciliation.automatic_proposal_id = p_proposal_id
         or reconciliation.created_by = v_actor
    )
    or exists (
      select 1
      from public.financial_reconciliation_items item
      where item.created_by = v_actor
    )
    or exists (
      select 1
      from public.financial_reconciliation_audit audit
      where audit.actor = v_actor
    ) then
    raise exception 'POS income stale outcome was not atomic or sanitized for %: %.',
      p_proposal_id, v_result;
  end if;
end
$$;

-- A monthly run containing only an above-tolerance aggregate must remain ready
-- so its immutable source and destination snapshots stay reviewable.
select pg_temp.pos_income_task4_clone_proposal(
  '2026-02', '85a00000-0000-0000-0000-000000000001',
  '85a10000-0000-0000-0000-000000000001',
  'smoke:pos-income-ambiguous-only'
);
update public.financial_reconciliation_automatic_runs
set status = 'analyzing', analysis_completed_at = null, finished_at = null,
    counts = '{}'::jsonb
where id = '85a00000-0000-0000-0000-000000000001';

do $$
declare
  v_result jsonb;
begin
  v_result := public.financial_reconciliation_finalize_automatic_analysis(
    '85a00000-0000-0000-0000-000000000001'
  );
  if v_result->>'status' is distinct from 'ready'
    or v_result->>'finishedAt' is not null
    or v_result#>>'{counts,proposed}' is distinct from '0'
    or v_result#>>'{counts,ambiguous}' is distinct from '1'
    or jsonb_array_length(v_result->'proposals') <> 1
    or v_result#>>'{proposals,0,status}' is distinct from 'ambiguous'
    or v_result#>>'{proposals,0,reason}' is distinct from
      'monthly_difference_exceeded' then
    raise exception 'POS income ambiguous-only run was not retained ready: %',
      v_result;
  end if;
end $$;

do $$
declare
  v_rejected boolean := false;
  v_signature text;
  v_eligible boolean;
begin
  insert into public.import_fdm_accounts (
    id, import_batch, account, date_time_raw, event_date, category,
    amount, description
  ) values (
    '84200000-0000-0000-0000-000000000001',
    'smoke-pos-income-execution-eligibility', 'Credit Card', '2026-07-01',
    date '2026-07-01', 'POS income', 10.00,
    'managed-only eligibility probe'
  );

  select source.eligible into strict v_eligible
  from public.financial_reconciliation_source(
    'import_fdm_accounts', '84200000-0000-0000-0000-000000000001'
  ) source;
  if v_eligible then
    raise exception 'POS income execution broadened public FDM workbench eligibility.';
  end if;

  begin
    perform public.financial_reconciliation_action(
      'start', 'smoke:pos-income-managed-only', null,
      'import_fdm_accounts', '84200000-0000-0000-0000-000000000001', null
    );
  exception when others then
    v_rejected := sqlerrm =
      'Source record is not eligible for reconciliation.';
  end;
  if not v_rejected then
    raise exception 'Public reconciliation action admitted a managed-only FDM row.';
  end if;
  delete from public.import_fdm_accounts
  where id = '84200000-0000-0000-0000-000000000001';

  foreach v_signature in array array[
    'public.financial_reconciliation_execute_prior_proposal(uuid,text)',
    'public.financial_reconciliation_execute_monthly_income_proposal(uuid,text)'
  ] loop
    if has_function_privilege('anon', v_signature, 'EXECUTE')
      or has_function_privilege('authenticated', v_signature, 'EXECUTE')
      or has_function_privilege('service_role', v_signature, 'EXECUTE')
      or not (
        select procedure.prosecdef
          and coalesce(procedure.proconfig, '{}'::text[])
            @> array['search_path=public, pg_temp']
        from pg_proc procedure
        where procedure.oid = v_signature::regprocedure
      ) then
      raise exception 'POS income private execution helper is exposed or unsafe: %.',
        v_signature;
    end if;
  end loop;

  v_signature :=
    'public.execute_financial_reconciliation_automatic_proposal(uuid,text)';
  if has_function_privilege('anon', v_signature, 'EXECUTE')
    or has_function_privilege('authenticated', v_signature, 'EXECUTE')
    or not has_function_privilege('service_role', v_signature, 'EXECUTE')
    or not (
      select procedure.prosecdef
        and coalesce(procedure.proconfig, '{}'::text[])
          @> array['search_path=public, pg_temp']
      from pg_proc procedure
      where procedure.oid = v_signature::regprocedure
    ) then
    raise exception 'POS income public execution dispatcher ACL or search path changed.';
  end if;
end
$$;

-- Clone every execution fixture before any success consumes a shared member.
select pg_temp.pos_income_task4_clone_proposal(
  '2026-06', '86000000-0000-0000-0000-000000000001',
  '86100000-0000-0000-0000-000000000001', 'smoke:pos-income-zero'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000002',
  '86100000-0000-0000-0000-000000000002', 'smoke:pos-income-forced'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000003',
  '86100000-0000-0000-0000-000000000003', 'smoke:pos-income-gained'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000004',
  '86100000-0000-0000-0000-000000000004', 'smoke:pos-income-lost'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000005',
  '86100000-0000-0000-0000-000000000005', 'smoke:pos-income-predicate'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000006',
  '86100000-0000-0000-0000-000000000006', 'smoke:pos-income-account'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000007',
  '86100000-0000-0000-0000-000000000007', 'smoke:pos-income-source-date'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000008',
  '86100000-0000-0000-0000-000000000008', 'smoke:pos-income-destination-date'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000009',
  '86100000-0000-0000-0000-000000000009', 'smoke:pos-income-source-amount'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000010',
  '86100000-0000-0000-0000-000000000010', 'smoke:pos-income-destination-amount'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000011',
  '86100000-0000-0000-0000-000000000011', 'smoke:pos-income-rule-missing'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000012',
  '86100000-0000-0000-0000-000000000012', 'smoke:pos-income-operator'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000013',
  '86100000-0000-0000-0000-000000000013', 'smoke:pos-income-definition'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000014',
  '86100000-0000-0000-0000-000000000014', 'smoke:pos-income-disabled'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000015',
  '86100000-0000-0000-0000-000000000015', 'smoke:pos-income-tolerance'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000016',
  '86100000-0000-0000-0000-000000000016', 'smoke:pos-income-priority'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000017',
  '86100000-0000-0000-0000-000000000017', 'smoke:pos-income-consumed'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000018',
  '86100000-0000-0000-0000-000000000018', 'smoke:pos-income-null'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000019',
  '86100000-0000-0000-0000-000000000019', 'smoke:pos-income-malformed'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-06', '86000000-0000-0000-0000-000000000020',
  '86100000-0000-0000-0000-000000000020', 'smoke:pos-income-competing'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-06', '86000000-0000-0000-0000-000000000021',
  '86100000-0000-0000-0000-000000000021', 'smoke:pos-income-rollback'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000022',
  '86100000-0000-0000-0000-000000000022',
  'smoke:pos-income-source-description'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-01', '86000000-0000-0000-0000-000000000023',
  '86100000-0000-0000-0000-000000000023',
  'smoke:pos-income-destination-description'
);
select pg_temp.pos_income_task4_clone_proposal(
  '2026-06', '86000000-0000-0000-0000-000000000024',
  '86100000-0000-0000-0000-000000000024',
  'smoke:pos-income-unexpected-unique'
);

-- Observe the serialization protocol at the first write after exact stale
-- revalidation and at both ends of a successful lifecycle. PostgreSQL relation
-- locks are transaction-scoped, so these checkpoints prove the three writer-
-- conflicting locks are already held after revalidation and remain held
-- through proposal completion.
create temporary table pos_income_task4_serialization_observations (
  proposal_id uuid not null,
  observed_status text not null,
  locked_relations text[] not null,
  primary key (proposal_id, observed_status)
);

create or replace function pg_temp.pos_income_task4_observe_serialization()
returns trigger
language plpgsql
set search_path = public, pg_temp
as $$
declare
  v_locked_relations text[];
begin
  if (new.id = '86100000-0000-0000-0000-000000000003'
      and new.status = 'stale')
    or (new.id = '86100000-0000-0000-0000-000000000001'
        and new.status in ('executing', 'completed')) then
    select array_agg(target.relation_name order by target.lock_ordinal)
    into v_locked_relations
    from (values
      (1, 'import_cgd_extrato_ordem'::text,
       'public.import_cgd_extrato_ordem'::regclass),
      (2, 'import_fdm_accounts'::text,
       'public.import_fdm_accounts'::regclass),
      (3, 'financial_reconciliation_items'::text,
       'public.financial_reconciliation_items'::regclass)
    ) target(lock_ordinal, relation_name, relation_oid)
    where exists (
      select 1
      from pg_catalog.pg_locks held
      where held.pid = pg_backend_pid()
        and held.locktype = 'relation'
        and held.relation = target.relation_oid
        and held.mode = 'ShareRowExclusiveLock'
        and held.granted
    );

    if v_locked_relations is distinct from array[
      'import_cgd_extrato_ordem',
      'import_fdm_accounts',
      'financial_reconciliation_items'
    ]::text[] then
      raise exception
        'POS income serialization locks were not held at %: %.',
        new.status, v_locked_relations;
    end if;

    insert into pg_temp.pos_income_task4_serialization_observations (
      proposal_id, observed_status, locked_relations
    ) values (new.id, new.status, v_locked_relations);
  end if;
  return new;
end
$$;

create trigger pos_income_task4_observe_serialization
before update on public.financial_reconciliation_automatic_proposals
for each row execute function pg_temp.pos_income_task4_observe_serialization();

-- Above-tolerance proposals remain non-executable and make no lifecycle writes.
do $$
declare
  v_proposal_id uuid;
  v_rejected boolean := false;
  v_reconciliations_before bigint;
begin
  select proposal.id into strict v_proposal_id
  from public.financial_reconciliation_automatic_proposals proposal
  join pos_income_task3_success_run successful
    on successful.run_id = proposal.run_id
  where proposal.grouping_key = '2026-02';
  select count(*) into v_reconciliations_before
  from public.financial_reconciliations;

  begin
    perform public.execute_financial_reconciliation_automatic_proposal(
      v_proposal_id, 'smoke:pos-income-ambiguous'
    );
  exception when others then
    v_rejected := sqlerrm =
      'Automation proposal with status ambiguous cannot be executed.';
  end;
  if not v_rejected
    or (select count(*) from public.financial_reconciliations) <>
       v_reconciliations_before
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.id = v_proposal_id
        and proposal.status = 'ambiguous'
        and proposal.reason = 'monthly_difference_exceeded'
        and proposal.reconciliation_id is null
    ) then
    raise exception 'POS income above-tolerance proposal executed or mutated.';
  end if;
end
$$;

-- Entire live source/destination membership and every date/amount are compared
-- against the immutable snapshot in both directions.
do $$
declare
  v_lock_id uuid;
  v_result jsonb;
  v_definition jsonb;
  v_priority integer;
begin
  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  ) values (
    '83200000-0000-0000-0000-000000000001',
    'smoke-pos-income-execution-drift', 'pos-income-execution-gained',
    date '2026-01-06', 'POS VENDAS gained after analysis', 1.00
  );
  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000003',
    'source_snapshot_changed'
  );
  delete from public.import_cgd_extrato_ordem
  where id = '83200000-0000-0000-0000-000000000001';

  delete from public.import_fdm_accounts
  where id = '84000000-0000-0000-0000-000000000010';
  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000004',
    'source_snapshot_changed'
  );
  insert into public.import_fdm_accounts (
    id, import_batch, account, date_time_raw, event_date, category,
    amount, description
  ) values (
    '84000000-0000-0000-0000-000000000010',
    'smoke-pos-income-analysis', 'Credit Card', '2026-01-31',
    date '2026-01-31', 'POS income', 100.00, 'January'
  );

  update public.import_cgd_extrato_ordem
  set descritivo = 'CARD SALES January'
  where id = '83000000-0000-0000-0000-000000000010';
  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000005',
    'source_snapshot_changed'
  );
  update public.import_cgd_extrato_ordem
  set descritivo = 'prefix pos vendas suffix'
  where id = '83000000-0000-0000-0000-000000000010';

  update public.import_fdm_accounts set account = 'Cash'
  where id = '84000000-0000-0000-0000-000000000010';
  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000006',
    'source_snapshot_changed'
  );
  update public.import_fdm_accounts set account = 'Credit Card'
  where id = '84000000-0000-0000-0000-000000000010';

  update public.import_cgd_extrato_ordem set data = date '2026-02-01'
  where id = '83000000-0000-0000-0000-000000000010';
  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000007',
    'source_snapshot_changed'
  );
  update public.import_cgd_extrato_ordem set data = date '2026-01-05'
  where id = '83000000-0000-0000-0000-000000000010';

  update public.import_fdm_accounts set event_date = date '2026-02-01'
  where id = '84000000-0000-0000-0000-000000000010';
  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000008',
    'source_snapshot_changed'
  );
  update public.import_fdm_accounts set event_date = date '2026-01-31'
  where id = '84000000-0000-0000-0000-000000000010';

  update public.import_cgd_extrato_ordem set montante = 4000.01
  where id = '83000000-0000-0000-0000-000000000010';
  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000009',
    'source_snapshot_changed'
  );
  update public.import_cgd_extrato_ordem set montante = 4000.00
  where id = '83000000-0000-0000-0000-000000000010';

  update public.import_fdm_accounts set amount = 100.01
  where id = '84000000-0000-0000-0000-000000000010';
  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000010',
    'source_snapshot_changed'
  );
  update public.import_fdm_accounts set amount = 100.00
  where id = '84000000-0000-0000-0000-000000000010';

  update public.import_cgd_extrato_ordem
  set descritivo = 'POS VENDAS January description drift'
  where id = '83000000-0000-0000-0000-000000000010';
  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000022',
    'source_snapshot_changed'
  );
  update public.import_cgd_extrato_ordem
  set descritivo = 'prefix pos vendas suffix'
  where id = '83000000-0000-0000-0000-000000000010';

  update public.import_fdm_accounts
  set description = 'January destination description drift'
  where id = '84000000-0000-0000-0000-000000000010';
  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000023',
    'source_snapshot_changed'
  );
  update public.import_fdm_accounts set description = 'January'
  where id = '84000000-0000-0000-0000-000000000010';

  delete from public.financial_reconciliation_source_rules
  where base_source_type = 'import_cgd_extrato_ordem'
    and matching_source_type = 'import_fdm_accounts';
  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000011',
    'rule_snapshot_changed'
  );
  insert into public.financial_reconciliation_source_rules (
    base_source_type, matching_source_type, operator
  ) values ('import_cgd_extrato_ordem', 'import_fdm_accounts', '-');

  update public.financial_reconciliation_source_rules set operator = '+'
  where base_source_type = 'import_cgd_extrato_ordem'
    and matching_source_type = 'import_fdm_accounts';
  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000012',
    'operator_changed'
  );
  update public.financial_reconciliation_source_rules set operator = '-'
  where base_source_type = 'import_cgd_extrato_ordem'
    and matching_source_type = 'import_fdm_accounts';

  select definition into strict v_definition
  from public.financial_reconciliation_automatic_rule_definitions
  where rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income'
    and version = 2;
  update public.financial_reconciliation_automatic_rule_definitions
  set definition = definition || '{"executionDrift":true}'::jsonb
  where rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income'
    and version = 2;
  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000013',
    'rule_snapshot_changed'
  );
  update public.financial_reconciliation_automatic_rule_definitions
  set definition = v_definition
  where rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income'
    and version = 2;

  update public.financial_reconciliation_automatic_rule_configs
  set enabled = false
  where rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income';
  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000014',
    'rule_snapshot_changed'
  );
  update public.financial_reconciliation_automatic_rule_configs
  set enabled = true
  where rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income';

  update public.financial_reconciliation_automatic_rule_configs
  set difference_allowed = 7499.99
  where rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income';
  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000015',
    'tolerance_changed'
  );
  update public.financial_reconciliation_automatic_rule_configs
  set difference_allowed = 7500.00
  where rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income';

  select priority into strict v_priority
  from public.financial_reconciliation_automatic_rule_configs
  where rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income';
  update public.financial_reconciliation_automatic_rule_configs
  set priority = 1000000
  where rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income';
  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000016',
    'rule_snapshot_changed'
  );
  update public.financial_reconciliation_automatic_rule_configs
  set priority = v_priority
  where rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income';

  v_result := public.financial_reconciliation_action(
    'start', 'smoke:pos-income-consumed-holder', null,
    'import_cgd_extrato_ordem',
    '83000000-0000-0000-0000-000000000010', null
  );
  v_lock_id := (v_result#>>'{reconciliation,id}')::uuid;
  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000017',
    'source_snapshot_changed'
  );
  perform public.financial_reconciliation_action(
    'delete', 'smoke:pos-income-consumed-holder',
    v_lock_id, null, null, null
  );

  update public.import_cgd_extrato_ordem set montante = null
  where id = '83000000-0000-0000-0000-000000000010';
  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000018',
    'source_snapshot_changed'
  );
  update public.import_cgd_extrato_ordem set montante = 4000.00
  where id = '83000000-0000-0000-0000-000000000010';

  update public.financial_reconciliation_automatic_runs
  set definition_config_snapshot = jsonb_set(
    definition_config_snapshot,
    '{0,ruleVersion}',
    '"999999999999999999999999999999999999999999999"'::jsonb
  )
  where id = '86000000-0000-0000-0000-000000000019';
  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000019',
    'rule_snapshot_changed'
  );
end
$$;

-- Simulate a destination becoming consumed after the monthly start audit but
-- before the remaining item locks are inserted. The whole lifecycle must roll
-- back and leave only the proposal's stale outcome.
create temporary table pos_income_task4_competing_lock (
  reconciliation_id uuid not null,
  source_id uuid not null
);

create or replace function pg_temp.pos_income_task4_consume_after_start()
returns trigger
language plpgsql
set search_path = public, pg_temp
as $$
begin
  if new.action = 'start' and new.actor = 'smoke:pos-income-competing' then
    insert into public.financial_reconciliation_items (
      reconciliation_id, source_type, source_id, amount_snapshot, created_by
    )
    select reconciliation_id, 'import_fdm_accounts', source_id, 75.00,
           'smoke:pos-income-competing'
    from pos_income_task4_competing_lock;
  end if;
  return new;
end
$$;

create trigger pos_income_task4_consume_after_start
after insert on public.financial_reconciliation_audit
for each row execute function pg_temp.pos_income_task4_consume_after_start();

do $$
declare
  v_competing_reconciliation_id uuid;
begin
  insert into public.financial_reconciliations (
    status, base_source_type, matching_source_types,
    matching_source_rules, created_by
  ) values (
    'started', 'import_cgd_extrato_ordem',
    '["import_fdm_accounts"]'::jsonb,
    '[{"sourceType":"import_fdm_accounts","operator":"-"}]'::jsonb,
    'smoke:pos-income-competing-holder'
  ) returning id into v_competing_reconciliation_id;
  insert into pos_income_task4_competing_lock values (
    v_competing_reconciliation_id,
    '84000000-0000-0000-0000-000000000060'
  );

  perform pg_temp.pos_income_task4_assert_stale(
    '86100000-0000-0000-0000-000000000020',
    'source_snapshot_changed'
  );
  if exists (
    select 1
    from public.financial_reconciliation_items item
    where item.reconciliation_id = v_competing_reconciliation_id
  ) then
    raise exception 'POS income competing-consumption rollback left a partial lock.';
  end if;
end
$$;

drop trigger pos_income_task4_consume_after_start
on public.financial_reconciliation_audit;
drop function pg_temp.pos_income_task4_consume_after_start();

-- A failure after reconciliation start rolls back reconciliation, item locks,
-- lifecycle audit, provenance, and proposal completion before persisting failed.
create or replace function pg_temp.pos_income_task4_reject_automatic_complete()
returns trigger
language plpgsql
set search_path = public, pg_temp
as $$
begin
  if new.action = 'automatic_complete'
    and new.actor = 'smoke:pos-income-rollback' then
    raise exception using
      message = 'task4-secret-message-77fbd4d6',
      detail = 'task4-secret-detail-77fbd4d6';
  end if;
  return new;
end
$$;

create trigger pos_income_task4_reject_automatic_complete
before insert on public.financial_reconciliation_audit
for each row execute function pg_temp.pos_income_task4_reject_automatic_complete();

do $$
declare
  v_result jsonb;
  v_public_run jsonb;
  v_secret_pattern text := '%77fbd4d6%';
begin
  v_result := public.execute_financial_reconciliation_automatic_proposal(
    '86100000-0000-0000-0000-000000000021',
    'smoke:pos-income-rollback'
  );
  v_public_run := public.get_financial_reconciliation_automatic_run(
    '86000000-0000-0000-0000-000000000021'
  );
  if v_result is distinct from jsonb_build_object(
      'proposalId', '86100000-0000-0000-0000-000000000021'::uuid,
      'runId', '86000000-0000-0000-0000-000000000021'::uuid,
      'status', 'failed',
      'reason', 'execution_failed'
    )
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.id = '86100000-0000-0000-0000-000000000021'
        and proposal.status = 'failed'
        and proposal.reason = 'execution_failed'
        and proposal.reconciliation_id is null
        and proposal.completed_at is null
        and proposal.error = 'Automatic reconciliation execution failed.'
        and proposal.error_detail = ''
    )
    or v_result::text ilike v_secret_pattern
    or v_public_run::text ilike v_secret_pattern
    or exists (
      select 1
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.id = '86100000-0000-0000-0000-000000000021'
        and to_jsonb(proposal)::text ilike v_secret_pattern
    )
    or exists (
      select 1
      from public.financial_reconciliation_automatic_runs run
      where run.id = '86000000-0000-0000-0000-000000000021'
        and to_jsonb(run)::text ilike v_secret_pattern
    )
    or exists (
      select 1 from public.financial_reconciliations reconciliation
      where reconciliation.automatic_proposal_id =
        '86100000-0000-0000-0000-000000000021'
         or reconciliation.created_by = 'smoke:pos-income-rollback'
    )
    or exists (
      select 1 from public.financial_reconciliation_items item
      where item.created_by = 'smoke:pos-income-rollback'
    )
    or exists (
      select 1 from public.financial_reconciliation_audit audit
      where audit.actor = 'smoke:pos-income-rollback'
         or to_jsonb(audit)::text ilike v_secret_pattern
    ) then
    raise exception 'POS income failure leaked diagnostics or left lifecycle writes: %.',
      v_result;
  end if;
end
$$;

drop trigger pos_income_task4_reject_automatic_complete
on public.financial_reconciliation_audit;
drop function pg_temp.pos_income_task4_reject_automatic_complete();

-- Only the source-lock uniqueness constraint is expected drift. An unrelated
-- unique violation must take the sanitized internal-failure path and roll back
-- every reconciliation lifecycle write.
create temporary table pos_income_task4_unexpected_unique (
  token text primary key
);
insert into pos_income_task4_unexpected_unique values ('occupied');

create or replace function pg_temp.pos_income_task4_raise_unexpected_unique()
returns trigger
language plpgsql
set search_path = public, pg_temp
as $$
begin
  if new.action = 'automatic_complete'
    and new.actor = 'smoke:pos-income-unexpected-unique' then
    insert into pg_temp.pos_income_task4_unexpected_unique values ('occupied');
  end if;
  return new;
end
$$;

create trigger pos_income_task4_raise_unexpected_unique
before insert on public.financial_reconciliation_audit
for each row execute function pg_temp.pos_income_task4_raise_unexpected_unique();

do $$
declare
  v_result jsonb;
begin
  v_result := public.execute_financial_reconciliation_automatic_proposal(
    '86100000-0000-0000-0000-000000000024',
    'smoke:pos-income-unexpected-unique'
  );
  if v_result is distinct from jsonb_build_object(
      'proposalId', '86100000-0000-0000-0000-000000000024'::uuid,
      'runId', '86000000-0000-0000-0000-000000000024'::uuid,
      'status', 'failed',
      'reason', 'execution_failed'
    )
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.id = '86100000-0000-0000-0000-000000000024'
        and proposal.status = 'failed'
        and proposal.reason = 'execution_failed'
        and proposal.reconciliation_id is null
        and proposal.completed_at is null
        and proposal.error = 'Automatic reconciliation execution failed.'
        and proposal.error_detail = ''
    )
    or exists (
      select 1
      from public.financial_reconciliations reconciliation
      where reconciliation.automatic_proposal_id =
        '86100000-0000-0000-0000-000000000024'
        or reconciliation.created_by = 'smoke:pos-income-unexpected-unique'
    )
    or exists (
      select 1
      from public.financial_reconciliation_items item
      where item.created_by = 'smoke:pos-income-unexpected-unique'
    )
    or exists (
      select 1
      from public.financial_reconciliation_audit audit
      where audit.actor = 'smoke:pos-income-unexpected-unique'
    ) then
    raise exception 'POS income unexpected unique violation was misclassified or partially persisted: %',
      v_result;
  end if;
end $$;

drop trigger pos_income_task4_raise_unexpected_unique
on public.financial_reconciliation_audit;
drop function pg_temp.pos_income_task4_raise_unexpected_unique();

-- Normal and forced completion preserve every lock, provenance field, ordinary
-- lifecycle audit record, immutable snapshots, generated comment, and retry ID.
do $$
declare
  v_case record;
  v_result jsonb;
  v_retry jsonb;
  v_reconciliation_id uuid;
  v_expected_comment text;
  v_items_before bigint;
  v_audit_before bigint;
begin
  for v_case in
    select * from (values
      ('86100000-0000-0000-0000-000000000001'::uuid,
       '86000000-0000-0000-0000-000000000001'::uuid,
       'smoke:pos-income-zero'::text, 'normal'::text, 0.00::numeric,
       null::text, 2::integer),
      ('86100000-0000-0000-0000-000000000002'::uuid,
       '86000000-0000-0000-0000-000000000002'::uuid,
       'smoke:pos-income-forced'::text, 'forced'::text, 7500.00::numeric,
       'Automatic monthly reconciliation for 2026-01: Bank Statement total 7600.00 EUR; FDM Credit Card total 100.00 EUR; difference 7500.00 EUR within allowed 7500.00 EUR; run 86000000-0000-0000-0000-000000000002; proposal 86100000-0000-0000-0000-000000000002.'::text,
       3::integer)
    ) expected(
      proposal_id, run_id, actor, completion_type, difference_amount,
      completion_comment, item_count
    )
  loop
    v_result := public.execute_financial_reconciliation_automatic_proposal(
      v_case.proposal_id, v_case.actor
    );
    v_reconciliation_id := (v_result->>'reconciliationId')::uuid;
    v_expected_comment := v_case.completion_comment;

    if v_result is distinct from jsonb_build_object(
        'proposalId', v_case.proposal_id,
        'runId', v_case.run_id,
        'status', 'completed',
        'reconciliationId', v_reconciliation_id
      )
      or v_reconciliation_id is null
      or not exists (
        select 1
        from public.financial_reconciliations reconciliation
        where reconciliation.id = v_reconciliation_id
          and reconciliation.status = 'complete'
          and reconciliation.completion_type = v_case.completion_type
          and reconciliation.difference_amount = v_case.difference_amount
          and reconciliation.forced_completion_comment is not distinct from
            v_expected_comment
          and reconciliation.origin = 'automatic'
          and reconciliation.automatic_trigger = 'manual'
          and reconciliation.automatic_rule_key =
            'cgd_bank_statement_fdm_credit_card_monthly_income'
          and reconciliation.automatic_rule_version = 2
          and reconciliation.automatic_run_id = v_case.run_id
          and reconciliation.automatic_proposal_id = v_case.proposal_id
          and reconciliation.matching_source_rules @> jsonb_build_array(
            jsonb_build_object(
              'sourceType', 'import_fdm_accounts', 'operator', '-'
            )
          )
      )
      or (select count(*)
          from public.financial_reconciliation_items item
          where item.reconciliation_id = v_reconciliation_id) <>
         v_case.item_count
      or exists (
        select membership.source_type, membership.source_id, membership.amount
        from public.financial_reconciliation_automatic_proposal_memberships membership
        where membership.proposal_id = v_case.proposal_id
        except
        select item.source_type, item.source_id, item.amount_snapshot
        from public.financial_reconciliation_items item
        where item.reconciliation_id = v_reconciliation_id
      )
      or exists (
        select item.source_type, item.source_id, item.amount_snapshot
        from public.financial_reconciliation_items item
        where item.reconciliation_id = v_reconciliation_id
        except
        select membership.source_type, membership.source_id, membership.amount
        from public.financial_reconciliation_automatic_proposal_memberships membership
        where membership.proposal_id = v_case.proposal_id
      ) then
      raise exception 'POS income ordinary reconciliation lifecycle is incomplete for %: %.',
        v_case.proposal_id, v_result;
    end if;

    if (select count(*)
        from public.financial_reconciliation_audit audit
        where audit.reconciliation_id = v_reconciliation_id
          and audit.action = 'automatic_complete') <> 1
      or (select count(*)
          from public.financial_reconciliation_audit audit
          where audit.reconciliation_id = v_reconciliation_id) <>
         v_case.item_count + 2
      or not exists (
        select 1
        from public.financial_reconciliation_audit audit
        join public.financial_reconciliation_automatic_proposals proposal
          on proposal.id = v_case.proposal_id
        where audit.reconciliation_id = v_reconciliation_id
          and audit.action = 'automatic_complete'
          and audit.actor = v_case.actor
          and audit.comment is not distinct from v_expected_comment
          and audit.difference_amount = v_case.difference_amount
          and audit.metadata @> jsonb_build_object(
            'ruleSnapshot', jsonb_build_object(
              'ruleKey', proposal.rule_key,
              'ruleVersion', proposal.rule_version
            ),
            'configSnapshot', jsonb_build_object(
              'differenceAllowed', proposal.allowed_difference,
              'maxDifferenceDays', 31
            ),
            'operatorSnapshot', jsonb_build_object(
              'import_fdm_accounts', '-'
            ),
            'summarySnapshot', proposal.summary_snapshot,
            'proposalSignature', proposal.signature,
            'trigger', 'manual',
            'runId', v_case.run_id,
            'proposalId', v_case.proposal_id,
            'tolerance', proposal.allowed_difference,
            'calculatedDifference', v_case.difference_amount
          )
          and jsonb_typeof(audit.metadata->'membershipSnapshots') = 'array'
          and jsonb_array_length(audit.metadata->'membershipSnapshots') =
            v_case.item_count
          and audit.metadata->'membershipSnapshots' = (
            select jsonb_agg(jsonb_build_object(
              'role', membership.role,
              'sourceType', membership.source_type,
              'sourceId', membership.source_id,
              'ordinal', membership.ordinal,
              'sourceDate', membership.source_date,
              'amount', membership.amount,
              'description', membership.description,
              'account', membership.account,
              'rowSnapshot', membership.row_snapshot
            ) order by membership.source_type, membership.source_date,
                       membership.source_id)
            from public.financial_reconciliation_automatic_proposal_memberships membership
            where membership.proposal_id = v_case.proposal_id
          )
      ) then
      raise exception 'POS income automatic audit is duplicated or incomplete for %.',
        v_case.proposal_id;
    end if;

    select count(*) into v_items_before
    from public.financial_reconciliation_items item
    where item.reconciliation_id = v_reconciliation_id;
    select count(*) into v_audit_before
    from public.financial_reconciliation_audit audit
    where audit.reconciliation_id = v_reconciliation_id;
    v_retry := public.execute_financial_reconciliation_automatic_proposal(
      v_case.proposal_id, v_case.actor
    );
    if v_retry is distinct from v_result
      or (select count(*)
          from public.financial_reconciliation_items item
          where item.reconciliation_id = v_reconciliation_id) <>
         v_items_before
      or (select count(*)
          from public.financial_reconciliation_audit audit
          where audit.reconciliation_id = v_reconciliation_id) <>
         v_audit_before then
      raise exception 'POS income retry duplicated lifecycle rows for %.',
        v_case.proposal_id;
    end if;
  end loop;
end
$$;

do $$
begin
  if not exists (
      select 1
      from pg_temp.pos_income_task4_serialization_observations observation
      where observation.proposal_id =
          '86100000-0000-0000-0000-000000000003'
        and observation.observed_status = 'stale'
    )
    or not exists (
      select 1
      from pg_temp.pos_income_task4_serialization_observations observation
      where observation.proposal_id =
          '86100000-0000-0000-0000-000000000001'
        and observation.observed_status = 'executing'
    )
    or not exists (
      select 1
      from pg_temp.pos_income_task4_serialization_observations observation
      where observation.proposal_id =
          '86100000-0000-0000-0000-000000000001'
        and observation.observed_status = 'completed'
    ) then
    raise exception
      'POS income serialization protocol missed a revalidation/lifecycle checkpoint.';
  end if;
end
$$;

drop trigger pos_income_task4_observe_serialization
on public.financial_reconciliation_automatic_proposals;
drop function pg_temp.pos_income_task4_observe_serialization();
drop table pg_temp.pos_income_task4_serialization_observations;

-- Execute the exact 1,000 + 1,000 immutable membership set and verify the
-- paginated history summary reports both raw counts and totals.
do $$
declare
  v_proposal_id uuid;
  v_run_id uuid;
  v_reconciliation_id uuid;
  v_result jsonb;
  v_history jsonb;
  v_row jsonb;
begin
  select proposal.id, proposal.run_id
  into strict v_proposal_id, v_run_id
  from public.financial_reconciliation_automatic_proposals proposal
  join pos_income_task3_success_run successful
    on successful.run_id = proposal.run_id
  where proposal.grouping_key = '2026-03';

  v_result := public.execute_financial_reconciliation_automatic_proposal(
    v_proposal_id, 'smoke:pos-income-large-execution'
  );
  v_reconciliation_id := (v_result->>'reconciliationId')::uuid;
  if v_result->>'status' <> 'completed'
    or v_reconciliation_id is null
    or (select count(*)
        from public.financial_reconciliation_items item
        where item.reconciliation_id = v_reconciliation_id) <> 2000
    or (select count(*)
        from public.financial_reconciliation_audit audit
        where audit.reconciliation_id = v_reconciliation_id
          and audit.action = 'automatic_complete') <> 1 then
    raise exception 'POS income 1,000 + 1,000 execution was incomplete: %.',
      v_result;
  end if;

  v_history := public.get_financial_reconciliation_history(
    null, null, 'automatic', 'complete', null, null, 1, 100
  );
  select history_row.value into strict v_row
  from jsonb_array_elements(v_history->'rows') history_row(value)
  where history_row.value->>'id' = v_reconciliation_id::text;
  if (v_row->>'totalRecords')::integer <> 2000
    or (v_row->>'sourceAmountTotal')::numeric <> 2000.00
    or (v_row->>'destinationAmountTotal')::numeric <> 1250.00
    or not (v_row->'sourceSummary') @> jsonb_build_array(
      jsonb_build_object(
        'sourceType', 'import_cgd_extrato_ordem',
        'recordCount', 1000,
        'amountTotal', 2000.00
      ),
      jsonb_build_object(
        'sourceType', 'import_fdm_accounts',
        'recordCount', 1000,
        'amountTotal', 1250.00
      )
    ) then
    raise exception 'POS income large history summary lost counts or totals: %.',
      v_row;
  end if;
end
$$;

-- The exact four pre-existing rule/version adapters still own their execution.
do $$
declare
  v_fixture record;
  v_result jsonb;
  v_items_before bigint;
  v_audit_before bigint;
  v_seen integer := 0;
begin
  for v_fixture in
    select distinct on (proposal.rule_key, proposal.rule_version)
      proposal.id, proposal.run_id, proposal.reconciliation_id,
      proposal.rule_key, proposal.rule_version
    from public.financial_reconciliation_automatic_proposals proposal
    where proposal.status = 'completed'
      and proposal.reconciliation_id is not null
      and (proposal.rule_key, proposal.rule_version) in (
        ('financial_documents_cgd_bank_statement', 2),
        ('financial_documents_cgd_credit_card', 1),
        ('financial_documents_cgd_bank_statement_amount_only', 1),
        ('financial_documents_cgd_credit_card_amount_only', 1)
      )
    order by proposal.rule_key, proposal.rule_version, proposal.id
  loop
    v_seen := v_seen + 1;
    select count(*) into v_items_before
    from public.financial_reconciliation_items item
    where item.reconciliation_id = v_fixture.reconciliation_id;
    select count(*) into v_audit_before
    from public.financial_reconciliation_audit audit
    where audit.reconciliation_id = v_fixture.reconciliation_id;

    v_result := public.execute_financial_reconciliation_automatic_proposal(
      v_fixture.id, 'smoke:pos-income-four-rule-regression'
    );
    if v_result is distinct from jsonb_build_object(
        'proposalId', v_fixture.id,
        'runId', v_fixture.run_id,
        'status', 'completed',
        'reconciliationId', v_fixture.reconciliation_id
      )
      or (select count(*)
          from public.financial_reconciliation_items item
          where item.reconciliation_id = v_fixture.reconciliation_id) <>
         v_items_before
      or (select count(*)
          from public.financial_reconciliation_audit audit
          where audit.reconciliation_id = v_fixture.reconciliation_id) <>
         v_audit_before then
      raise exception 'POS income dispatcher changed completed % v% execution.',
        v_fixture.rule_key, v_fixture.rule_version;
    end if;
  end loop;

  if v_seen <> 4 then
    raise exception 'POS income four-rule execution regression covered % adapters.',
      v_seen;
  end if;
end
$$;

-- App-authorized monthly snapshot member paging uses only the service-role RPC.
select pg_temp.pos_income_task4_clone_proposal(
  '2026-03', '87000000-0000-0000-0000-000000000001',
  '87100000-0000-0000-0000-000000000001', 'smoke:task5-owner'
);
delete from public.financial_reconciliation_automatic_proposal_memberships
where proposal_id = '87100000-0000-0000-0000-000000000001'
  and ((role = 'source' and ordinal > 123)
    or (role = 'destination' and ordinal > 3));

select pg_temp.pos_income_task4_clone_proposal(
  '2026-06', '87000000-0000-0000-0000-000000000002',
  '87100000-0000-0000-0000-000000000002', 'smoke:task5-owner'
);

insert into public.financial_reconciliation_automatic_runs (
  id, trigger, scope, status, actor, client_request_id,
  definition_config_snapshot, analysis_completed_at
) values (
  '87000000-0000-0000-0000-000000000003', 'manual', 'rule', 'ready',
  'smoke:task5-owner', '87000000-0000-0000-0000-000000000003',
  '[]'::jsonb, now()
);
insert into public.financial_reconciliation_automatic_proposals (
  id, run_id, rule_key, rule_version, base_source_type, base_source_id,
  base_source_date, base_snapshot, allowed_difference, signature
) values (
  '87100000-0000-0000-0000-000000000003',
  '87000000-0000-0000-0000-000000000003',
  'financial_documents_cgd_bank_statement', 2, 'financial_documents',
  '87100000-0000-0000-0000-000000000004', date '2026-03-01',
  '{}'::jsonb, 0.00, 'smoke:task5-non-monthly'
);

create or replace function pg_temp.pos_income_task5_assert_rejected(
  p_run_id uuid,
  p_proposal_id uuid,
  p_role text,
  p_offset integer,
  p_limit integer,
  p_actor text,
  p_expected_error text
)
returns void
language plpgsql
set search_path = public, pg_temp
as $$
declare
  v_rejected boolean := false;
begin
  begin
    perform public.get_financial_reconciliation_automatic_proposal_members(
      p_run_id, p_proposal_id, p_role, p_offset, p_limit, p_actor
    );
  exception when others then
    v_rejected := sqlerrm = p_expected_error;
  end;
  if not v_rejected then
    raise exception 'Task 5 member paging did not reject with expected error: %.',
      p_expected_error;
  end if;
end
$$;

do $$
declare
  v_page_zero jsonb;
  v_page_fifty jsonb;
  v_page_hundred jsonb;
  v_destination jsonb;
  v_completed_page jsonb;
  v_ordinals integer[];
  v_signature text :=
    'public.get_financial_reconciliation_automatic_proposal_members(uuid,uuid,text,integer,integer,text)';
begin
  perform pg_temp.pos_income_task5_assert_rejected(
    null, '87100000-0000-0000-0000-000000000001', 'source', 0, 50,
    'smoke:task5-owner', 'Automatic run ID is required.'
  );
  perform pg_temp.pos_income_task5_assert_rejected(
    '87000000-0000-0000-0000-000000000001', null, 'source', 0, 50,
    'smoke:task5-owner', 'Automatic proposal ID is required.'
  );
  perform pg_temp.pos_income_task5_assert_rejected(
    '87000000-0000-0000-0000-000000000001',
    '87100000-0000-0000-0000-000000000001', 'Source', 0, 50,
    'smoke:task5-owner', 'Automatic proposal member role is invalid.'
  );
  perform pg_temp.pos_income_task5_assert_rejected(
    '87000000-0000-0000-0000-000000000001',
    '87100000-0000-0000-0000-000000000001', 'source', -1, 50,
    'smoke:task5-owner', 'Automatic proposal member offset must be zero or greater.'
  );
  perform pg_temp.pos_income_task5_assert_rejected(
    '87000000-0000-0000-0000-000000000001',
    '87100000-0000-0000-0000-000000000001', 'source', null, 50,
    'smoke:task5-owner', 'Automatic proposal member offset must be zero or greater.'
  );
  perform pg_temp.pos_income_task5_assert_rejected(
    '87000000-0000-0000-0000-000000000001',
    '87100000-0000-0000-0000-000000000001', 'source', 0, 0,
    'smoke:task5-owner', 'Automatic proposal member limit must be between 1 and 50.'
  );
  perform pg_temp.pos_income_task5_assert_rejected(
    '87000000-0000-0000-0000-000000000001',
    '87100000-0000-0000-0000-000000000001', 'source', 0, 51,
    'smoke:task5-owner', 'Automatic proposal member limit must be between 1 and 50.'
  );
  perform pg_temp.pos_income_task5_assert_rejected(
    '87000000-0000-0000-0000-000000000001',
    '87100000-0000-0000-0000-000000000001', 'source', 0, 50, '  ',
    'Automatic proposal member actor is required.'
  );
  perform pg_temp.pos_income_task5_assert_rejected(
    '87000000-0000-0000-0000-000000000001',
    '87100000-0000-0000-0000-000000000001', 'source', 0, 50,
    'smoke:task5-foreign', 'You do not have permission for this automation run.'
  );
  perform pg_temp.pos_income_task5_assert_rejected(
    '87000000-0000-0000-0000-000000000001',
    '87100000-0000-0000-0000-000000000002', 'source', 0, 50,
    'smoke:task5-owner',
    'Automatic monthly proposal was not found for the requested run.'
  );
  perform pg_temp.pos_income_task5_assert_rejected(
    '87000000-0000-0000-0000-000000000099',
    '87100000-0000-0000-0000-000000000099', 'source', 0, 50,
    'smoke:task5-owner',
    'Automatic monthly proposal was not found for the requested run.'
  );
  perform pg_temp.pos_income_task5_assert_rejected(
    '87000000-0000-0000-0000-000000000003',
    '87100000-0000-0000-0000-000000000003', 'source', 0, 50,
    'smoke:task5-owner',
    'Automation proposal does not use the monthly income rule.'
  );

  v_page_zero := public.get_financial_reconciliation_automatic_proposal_members(
    '87000000-0000-0000-0000-000000000001',
    '87100000-0000-0000-0000-000000000001', 'source', 0, 50,
    'smoke:task5-owner'
  );
  v_page_fifty := public.get_financial_reconciliation_automatic_proposal_members(
    '87000000-0000-0000-0000-000000000001',
    '87100000-0000-0000-0000-000000000001', 'source', 50, 50,
    'smoke:task5-owner'
  );
  v_page_hundred := public.get_financial_reconciliation_automatic_proposal_members(
    '87000000-0000-0000-0000-000000000001',
    '87100000-0000-0000-0000-000000000001', 'source', 100, 50,
    'smoke:task5-owner'
  );
  v_destination := public.get_financial_reconciliation_automatic_proposal_members(
    '87000000-0000-0000-0000-000000000001',
    '87100000-0000-0000-0000-000000000001', 'destination', 0, 50,
    'smoke:task5-owner'
  );

  select array_agg((member.value->>'ordinal')::integer order by page_number,
                   member.ordinality)
  into v_ordinals
  from (values
    (0, v_page_zero),
    (1, v_page_fifty),
    (2, v_page_hundred)
  ) page(page_number, payload)
  cross join lateral jsonb_array_elements(page.payload->'members')
    with ordinality member(value, ordinality);

  if (v_page_zero - array[
      'runId','proposalId','role','offset','limit','totalCount','members'
    ]::text[]) <> '{}'::jsonb
    or v_page_zero->>'runId' <> '87000000-0000-0000-0000-000000000001'
    or v_page_zero->>'proposalId' <> '87100000-0000-0000-0000-000000000001'
    or v_page_zero->>'role' <> 'source'
    or (v_page_zero->>'offset')::integer <> 0
    or (v_page_zero->>'limit')::integer <> 50
    or (v_page_zero->>'totalCount')::integer <> 123
    or jsonb_array_length(v_page_zero->'members') <> 50
    or jsonb_array_length(v_page_fifty->'members') <> 50
    or jsonb_array_length(v_page_hundred->'members') <> 23
    or v_ordinals is distinct from array(select generate_series(1, 123))
    or exists (
      select 1
      from (values (v_page_zero), (v_page_fifty), (v_page_hundred)) page(payload)
      cross join lateral jsonb_array_elements(page.payload->'members') member(value)
      where member.value->>'role' <> 'source'
        or member.value->>'sourceType' <> 'import_cgd_extrato_ordem'
        or member.value - array[
          'role','sourceType','sourceId','ordinal','sourceDate','amount',
          'description','account','rowSnapshot'
        ]::text[] <> '{}'::jsonb
    )
    or (select count(distinct member.value->>'sourceId')
        from (values (v_page_zero), (v_page_fifty), (v_page_hundred)) page(payload)
        cross join lateral jsonb_array_elements(page.payload->'members') member(value)) <> 123
  then
    raise exception 'Task 5 source paging skipped, duplicated, reordered, or leaked snapshot fields: %, %, %.',
      v_page_zero, v_page_fifty, v_page_hundred;
  end if;

  if (v_destination->>'role' <> 'destination'
    or (v_destination->>'totalCount')::integer <> 3
    or jsonb_array_length(v_destination->'members') <> 3
    or exists (
      select 1 from jsonb_array_elements(v_destination->'members') member(value)
      where member.value->>'role' <> 'destination'
        or member.value->>'sourceType' <> 'import_fdm_accounts'
    ) then
    raise exception 'Task 5 member paging mixed source and destination roles: %.',
      v_destination;
  end if;

  perform public.finish_financial_reconciliation_automatic_run(
    '87000000-0000-0000-0000-000000000001'
  );
  v_completed_page :=
    public.get_financial_reconciliation_automatic_proposal_members(
      '87000000-0000-0000-0000-000000000001',
      '87100000-0000-0000-0000-000000000001', 'source', 100, 50,
      'smoke:task5-completed-reader'
    );
  if v_completed_page is distinct from v_page_hundred then
    raise exception 'Task 5 completed snapshot page changed or remained owner-locked.';
  end if;

  if has_function_privilege('anon', v_signature, 'EXECUTE')
    or has_function_privilege('authenticated', v_signature, 'EXECUTE')
    or not has_function_privilege('service_role', v_signature, 'EXECUTE')
    or not (
      select procedure.prosecdef
        and coalesce(procedure.proconfig, '{}'::text[])
          @> array['search_path=public, pg_temp']
      from pg_proc procedure
      where procedure.oid = v_signature::regprocedure
    )
    or has_table_privilege(
      'anon',
      'public.financial_reconciliation_automatic_proposal_memberships',
      'SELECT'
    )
    or has_table_privilege(
      'authenticated',
      'public.financial_reconciliation_automatic_proposal_memberships',
      'SELECT'
    )
    or not (
      select table_row.relrowsecurity
      from pg_class table_row
      where table_row.oid =
        'public.financial_reconciliation_automatic_proposal_memberships'::regclass
    ) then
    raise exception 'Task 5 member paging RPC ACL, search path, or table RLS is unsafe.';
  end if;
end
$$;

-- Task 7 assertion guard: retry and new-child response contracts reject
-- missing, unclaimed, cross-batch, cross-position, or terminal child metadata.
do $$
declare
  v_expected_run_id uuid := '00000000-0000-0000-0000-000000000001';
  v_expected_batch_id uuid := '00000000-0000-0000-0000-000000000002';
  v_expected_position integer := 2;
  v_expected_resumed boolean;
  v_valid_response jsonb;
  v_response jsonb;
  v_rejected boolean;
begin
  foreach v_expected_resumed in array array[true, false] loop
    v_valid_response := jsonb_build_object(
      'claimed', true,
      'resumed', v_expected_resumed,
      'batchId', v_expected_batch_id,
      'batchRulePosition', v_expected_position,
      'run', jsonb_build_object(
        'runId', v_expected_run_id,
        'status', 'analyzing',
        'batchId', v_expected_batch_id,
        'batchRulePosition', v_expected_position
      )
    );

    foreach v_response in array array[
      '{}'::jsonb,
      jsonb_set(v_valid_response, '{claimed}', 'false'::jsonb),
      v_valid_response - 'resumed',
      v_valid_response - 'batchId',
      jsonb_set(v_valid_response, '{batchRulePosition}', '99'::jsonb),
      v_valid_response #- '{run,runId}',
      jsonb_set(v_valid_response, '{run,status}', '"ready"'::jsonb),
      jsonb_set(
        v_valid_response,
        '{run,batchId}',
        '"00000000-0000-0000-0000-000000000099"'::jsonb
      ),
      jsonb_set(v_valid_response, '{run,batchRulePosition}', '99'::jsonb)
    ] loop
      v_rejected := false;
      begin
        if (v_response->>'claimed')::boolean is distinct from true
          or (v_response->>'resumed')::boolean is distinct from
            v_expected_resumed
          or v_response->>'batchId' is distinct from
            v_expected_batch_id::text
          or (v_response->>'batchRulePosition')::integer is distinct from
            v_expected_position
          or v_response#>>'{run,runId}' is distinct from
            v_expected_run_id::text
          or v_response#>>'{run,status}' is distinct from 'analyzing'
          or v_response#>>'{run,batchId}' is distinct from
            v_expected_batch_id::text
          or (v_response#>>'{run,batchRulePosition}')::integer is distinct from
            v_expected_position then
          raise check_violation using message =
            'Task 7 child response contract rejected malformed fixture.';
        end if;
      exception when check_violation then
        v_rejected := true;
      end;

      if v_rejected is distinct from true then
        raise exception 'Task 7 child response assertions no longer fail closed: %.',
          v_response;
      end if;
    end loop;

    if (v_valid_response->>'claimed')::boolean is distinct from true
      or (v_valid_response->>'resumed')::boolean is distinct from
        v_expected_resumed
      or v_valid_response->>'batchId' is distinct from
        v_expected_batch_id::text
      or (v_valid_response->>'batchRulePosition')::integer is distinct from
        v_expected_position
      or v_valid_response#>>'{run,runId}' is distinct from
        v_expected_run_id::text
      or v_valid_response#>>'{run,status}' is distinct from 'analyzing'
      or v_valid_response#>>'{run,batchId}' is distinct from
        v_expected_batch_id::text
      or (v_valid_response#>>'{run,batchRulePosition}')::integer is distinct from
        v_expected_position then
      raise exception 'Task 7 child response assertion guard rejected a valid fixture.';
    end if;
  end loop;
end
$$;

-- Task 7: the installed fifth rule stays out of unchanged four-rule batches
-- until an administrator explicitly enables scheduled execution.
do $$
declare
  v_rules jsonb := '[
    {"rule_key":"financial_documents_cgd_bank_statement","rule_version":2,"enabled":true,"allow_manual_execution":true,"include_in_scheduled_batch":true,"difference_allowed":"0.10","max_difference_days":10,"priority":1},
    {"rule_key":"financial_documents_cgd_credit_card","rule_version":1,"enabled":true,"allow_manual_execution":true,"include_in_scheduled_batch":true,"difference_allowed":"0.20","max_difference_days":11,"priority":2},
    {"rule_key":"financial_documents_cgd_bank_statement_amount_only","rule_version":1,"enabled":true,"allow_manual_execution":true,"include_in_scheduled_batch":true,"difference_allowed":"0.00","max_difference_days":1,"priority":3},
    {"rule_key":"financial_documents_cgd_credit_card_amount_only","rule_version":1,"enabled":true,"allow_manual_execution":true,"include_in_scheduled_batch":true,"difference_allowed":"0.00","max_difference_days":1,"priority":4},
    {"rule_key":"cgd_bank_statement_fdm_credit_card_monthly_income","rule_version":2,"enabled":false,"allow_manual_execution":false,"include_in_scheduled_batch":false,"difference_allowed":"7500.00","max_difference_days":31,"priority":5}
  ]'::jsonb;
  v_expected_rules text[] := array[
    'financial_documents_cgd_bank_statement',
    'financial_documents_cgd_credit_card',
    'financial_documents_cgd_bank_statement_amount_only',
    'financial_documents_cgd_credit_card_amount_only'
  ];
  v_settings jsonb;
  v_claim jsonb;
  v_retry jsonb;
  v_complete jsonb;
  v_batch_id uuid;
  v_run_id uuid;
  v_position integer;
begin
  v_settings := public.replace_financial_reconciliation_automation_settings(
    '{"enabled":true,"time_of_day":"00:00","time_zone":"Europe/Lisbon"}'::jsonb,
    v_rules,
    'smoke:task7-four-rule-settings'
  );
  if jsonb_array_length(v_settings->'rules') <> 5
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_configs config
      where config.rule_key =
          'cgd_bank_statement_fdm_credit_card_monthly_income'
        and config.rule_version = 2
        and not config.enabled
        and not config.allow_manual_execution
        and not config.include_in_scheduled_batch
        and config.difference_allowed = 7500.00
        and config.max_difference_days = 31
        and config.priority = 5
    ) then
    raise exception 'Task 7 five-rule Settings replacement changed the disabled monthly contract.';
  end if;

  for v_position in 1..4 loop
    v_claim := public.claim_financial_reconciliation_automatic_schedule(
      timestamptz '2095-01-01 01:00:00+00'
        + make_interval(mins => v_position),
      'smoke:task7-four-rule-batch'
    );
    v_batch_id := coalesce(v_batch_id, (v_claim->>'batchId')::uuid);
    v_run_id := (v_claim#>>'{run,runId}')::uuid;
    if not (v_claim->>'claimed')::boolean
      or (v_claim->>'resumed')::boolean
      or v_claim->>'batchId' <> v_batch_id::text
      or (v_claim->>'batchRulePosition')::integer <> v_position
      or (v_claim->>'batchRuleCount')::integer <> 4
      or v_claim#>>'{run,batchRuleKey}' <>
        v_expected_rules[v_position]
      or (select jsonb_array_length(batch.rule_snapshot)
          from public.financial_reconciliation_automatic_batches batch
          where batch.id = v_batch_id) <> 4
      or exists (
        select 1
        from public.financial_reconciliation_automatic_batches batch,
             jsonb_array_elements(batch.rule_snapshot) snapshot(value)
        where batch.id = v_batch_id
          and snapshot.value->>'ruleKey' =
            'cgd_bank_statement_fdm_credit_card_monthly_income'
      ) then
      raise exception 'Task 7 disabled monthly rule changed four-rule child position %: %.',
        v_position, v_claim;
    end if;

    v_retry := public.claim_financial_reconciliation_automatic_schedule(
      timestamptz '2095-01-01 01:10:00+00'
        + make_interval(mins => v_position),
      'smoke:task7-four-rule-batch'
    );
    if (v_retry->>'claimed')::boolean is distinct from true
      or (v_retry->>'resumed')::boolean is distinct from true
      or v_retry->>'batchId' is distinct from v_batch_id::text
      or (v_retry->>'batchRulePosition')::integer is distinct from v_position
      or v_retry#>>'{run,runId}' is distinct from v_run_id::text
      or v_retry#>>'{run,status}' is distinct from 'analyzing'
      or v_retry#>>'{run,batchId}' is distinct from v_batch_id::text
      or (v_retry#>>'{run,batchRulePosition}')::integer is distinct from
        v_position
      or (select count(*)
          from public.financial_reconciliation_automatic_runs run
          where run.batch_id = v_batch_id) is distinct from v_position then
      raise exception 'Task 7 four-rule retry duplicated child position %.',
        v_position;
    end if;

    update public.financial_reconciliation_automatic_runs
    set status = 'completed',
        analysis_completed_at = coalesce(analysis_completed_at, now()),
        counts = '{"bases":0,"completed":0,"failed":0}'::jsonb,
        finished_at = now(),
        updated_at = now()
    where id = v_run_id;
  end loop;

  v_complete := public.claim_financial_reconciliation_automatic_schedule(
    '2095-01-01 02:00:00+00', 'smoke:task7-four-rule-batch'
  );
  if (v_complete->>'claimed')::boolean is distinct from false
    or v_complete->>'reason' is distinct from 'batch_complete'
    or v_complete->>'batchId' is distinct from v_batch_id::text
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_batches batch
      where batch.id = v_batch_id
        and batch.status = 'completed'
        and batch.counts @> '{"ruleCount":4,"childCount":4,"completedChildren":4,"failedChildren":0}'::jsonb
    )
    or (select count(*)
        from public.financial_reconciliation_automatic_runs run
        where run.batch_id = v_batch_id) is distinct from 4
    or (select count(distinct run.batch_rule_position)
        from public.financial_reconciliation_automatic_runs run
        where run.batch_id = v_batch_id) is distinct from 4 then
    raise exception 'Task 7 disabled monthly rule no longer preserves the four-rule batch lifecycle.';
  end if;
end
$$;

-- Task 7: an administrator can reorder and schedule all five exact managed
-- contracts; the immutable batch advances only after each child is terminal.
do $$
declare
  v_rules jsonb := '[
    {"rule_key":"financial_documents_cgd_credit_card_amount_only","rule_version":1,"enabled":true,"allow_manual_execution":true,"include_in_scheduled_batch":true,"difference_allowed":"0.00","max_difference_days":1,"priority":1},
    {"rule_key":"cgd_bank_statement_fdm_credit_card_monthly_income","rule_version":2,"enabled":true,"allow_manual_execution":false,"include_in_scheduled_batch":true,"difference_allowed":"7500.00","max_difference_days":31,"priority":2},
    {"rule_key":"financial_documents_cgd_bank_statement_amount_only","rule_version":1,"enabled":true,"allow_manual_execution":true,"include_in_scheduled_batch":true,"difference_allowed":"0.00","max_difference_days":1,"priority":3},
    {"rule_key":"financial_documents_cgd_credit_card","rule_version":1,"enabled":true,"allow_manual_execution":true,"include_in_scheduled_batch":true,"difference_allowed":"0.20","max_difference_days":11,"priority":4},
    {"rule_key":"financial_documents_cgd_bank_statement","rule_version":2,"enabled":true,"allow_manual_execution":true,"include_in_scheduled_batch":true,"difference_allowed":"0.10","max_difference_days":10,"priority":5}
  ]'::jsonb;
  v_changed_rules jsonb := '[
    {"rule_key":"financial_documents_cgd_bank_statement","rule_version":2,"enabled":true,"allow_manual_execution":true,"include_in_scheduled_batch":true,"difference_allowed":"0.40","max_difference_days":13,"priority":1},
    {"rule_key":"financial_documents_cgd_credit_card","rule_version":1,"enabled":true,"allow_manual_execution":true,"include_in_scheduled_batch":true,"difference_allowed":"0.30","max_difference_days":12,"priority":2},
    {"rule_key":"financial_documents_cgd_bank_statement_amount_only","rule_version":1,"enabled":true,"allow_manual_execution":true,"include_in_scheduled_batch":true,"difference_allowed":"0.00","max_difference_days":3,"priority":3},
    {"rule_key":"financial_documents_cgd_credit_card_amount_only","rule_version":1,"enabled":true,"allow_manual_execution":true,"include_in_scheduled_batch":true,"difference_allowed":"0.00","max_difference_days":2,"priority":4},
    {"rule_key":"cgd_bank_statement_fdm_credit_card_monthly_income","rule_version":2,"enabled":false,"allow_manual_execution":false,"include_in_scheduled_batch":false,"difference_allowed":"7000.00","max_difference_days":31,"priority":5}
  ]'::jsonb;
  v_expected_rules text[] := array[
    'financial_documents_cgd_credit_card_amount_only',
    'cgd_bank_statement_fdm_credit_card_monthly_income',
    'financial_documents_cgd_bank_statement_amount_only',
    'financial_documents_cgd_credit_card',
    'financial_documents_cgd_bank_statement'
  ];
  v_settings jsonb;
  v_claim jsonb;
  v_retry jsonb;
  v_cursor_retry jsonb;
  v_complete jsonb;
  v_batch_id uuid;
  v_run_id uuid;
  v_monthly_run_id uuid;
  v_snapshot jsonb;
  v_month_count bigint;
  v_first_month record;
  v_position integer;
  v_runs_before bigint;
  v_proposals_before bigint;
  v_reconciliations_before bigint;
  v_fixed_days_rejected boolean := false;
begin
  begin
    perform public.replace_financial_reconciliation_automation_settings(
      '{"enabled":true,"time_of_day":"00:00","time_zone":"Europe/Lisbon"}'::jsonb,
      jsonb_set(v_rules, '{1,max_difference_days}', '30'::jsonb),
      'smoke:task7-fixed-days-tamper'
    );
  exception when others then
    v_fixed_days_rejected := sqlerrm =
      'POS income automatic rule requires the fixed 31-day display property.';
  end;
  if not v_fixed_days_rejected then
    raise exception 'Task 7 Settings accepted a monthly maximum difference other than 31.';
  end if;

  v_settings := public.replace_financial_reconciliation_automation_settings(
    '{"enabled":true,"time_of_day":"00:00","time_zone":"Europe/Lisbon"}'::jsonb,
    v_rules,
    'smoke:task7-five-rule-settings'
  );
  if jsonb_array_length(v_settings->'rules') <> 5
    or not exists (
      select 1
      from jsonb_array_elements(v_settings->'rules') rule(value)
      where rule.value->>'ruleKey' =
          'cgd_bank_statement_fdm_credit_card_monthly_income'
        and rule.value->>'ruleVersion' = '2'
        and rule.value->>'includeInScheduledBatch' = 'true'
        and rule.value->>'priority' = '2'
        and rule.value->>'maxDifferenceDays' = '31'
    ) then
    raise exception 'Task 7 Settings did not persist scheduled monthly priority and fixed days.';
  end if;

  v_claim := public.claim_financial_reconciliation_automatic_schedule(
    '2095-01-02 01:00:00+00', 'smoke:task7-five-rule-batch'
  );
  v_batch_id := (v_claim->>'batchId')::uuid;
  v_run_id := (v_claim#>>'{run,runId}')::uuid;
  select batch.rule_snapshot into strict v_snapshot
  from public.financial_reconciliation_automatic_batches batch
  where batch.id = v_batch_id;
  if not (v_claim->>'claimed')::boolean
    or (v_claim->>'resumed')::boolean
    or v_claim->>'batchRulePosition' <> '1'
    or v_claim->>'batchRuleCount' <> '5'
    or v_claim#>>'{run,batchRuleKey}' <> v_expected_rules[1]
    or jsonb_array_length(v_snapshot) <> 5
    or v_snapshot#>>'{0,ruleKey}' <> v_expected_rules[1]
    or v_snapshot#>>'{1,ruleKey}' <> v_expected_rules[2]
    or v_snapshot#>>'{1,ruleVersion}' <> '2'
    or v_snapshot#>>'{1,displayName}' <> 'Card Payments - POS - Income'
    or v_snapshot#>>'{1,priority}' <> '2'
    or v_snapshot#>>'{1,differenceAllowed}' <> '7500.00'
    or v_snapshot#>>'{1,maxDifferenceDays}' <> '31'
    or v_snapshot#>>'{1,destinationSourceType}' <> 'import_fdm_accounts'
    or v_snapshot#>>'{1,operator}' <> '-'
    or v_snapshot#>>'{1,definition,matchingMode}' <> 'monthly_aggregate'
    or v_snapshot#>>'{2,ruleKey}' <> v_expected_rules[3]
    or v_snapshot#>>'{3,ruleKey}' <> v_expected_rules[4]
    or v_snapshot#>>'{4,ruleKey}' <> v_expected_rules[5] then
    raise exception 'Task 7 scheduled batch did not snapshot all five exact managed contracts: %.',
      v_snapshot;
  end if;

  v_retry := public.claim_financial_reconciliation_automatic_schedule(
    '2095-01-02 01:01:00+00', 'smoke:task7-five-rule-batch'
  );
  if (v_retry->>'claimed')::boolean is distinct from true
    or (v_retry->>'resumed')::boolean is distinct from true
    or v_retry->>'batchId' is distinct from v_batch_id::text
    or (v_retry->>'batchRulePosition')::integer is distinct from 1
    or v_retry#>>'{run,runId}' is distinct from v_run_id::text
    or v_retry#>>'{run,status}' is distinct from 'analyzing'
    or v_retry#>>'{run,batchId}' is distinct from v_batch_id::text
    or (v_retry#>>'{run,batchRulePosition}')::integer is distinct from 1
    or (select count(*)
        from public.financial_reconciliation_automatic_runs run
        where run.batch_id = v_batch_id) is distinct from 1 then
    raise exception 'Task 7 retry duplicated the first five-rule child.';
  end if;

  perform public.replace_financial_reconciliation_automation_settings(
    '{"enabled":true,"time_of_day":"00:00","time_zone":"Europe/Lisbon"}'::jsonb,
    v_changed_rules,
    'smoke:task7-settings-after-snapshot'
  );
  if (select batch.rule_snapshot
      from public.financial_reconciliation_automatic_batches batch
      where batch.id = v_batch_id) is distinct from v_snapshot then
    raise exception 'Task 7 administrator changes rewrote the five-rule batch snapshot.';
  end if;

  update public.financial_reconciliation_automatic_runs
  set status = 'completed',
      analysis_completed_at = coalesce(analysis_completed_at, now()),
      counts = '{"bases":0,"completed":0,"failed":0}'::jsonb,
      finished_at = now(),
      updated_at = now()
  where id = v_run_id;

  select public.financial_reconciliation_automatic_monthly_income_count()
  into v_month_count;
  v_claim := public.claim_financial_reconciliation_automatic_schedule(
    '2095-01-03 00:30:00+00', 'smoke:task7-five-rule-batch'
  );
  v_monthly_run_id := (v_claim#>>'{run,runId}')::uuid;
  if (v_claim->>'resumed')::boolean
    or v_claim->>'batchId' <> v_batch_id::text
    or v_claim->>'batchRulePosition' <> '2'
    or v_claim->>'batchRuleCount' <> '5'
    or v_claim#>>'{run,batchRuleKey}' <> v_expected_rules[2]
    or (v_claim#>>'{run,analysisTotal}')::bigint <> v_month_count
    or jsonb_array_length(v_claim#>'{run,definitions}') <> 1
    or v_claim#>>'{run,definitions,0,maxDifferenceDays}' <> '31'
    or (select count(*)
        from public.financial_reconciliation_automatic_runs run
        where run.batch_id = v_batch_id) <> 2
    or exists (
      select 1
      from public.financial_reconciliation_automatic_batches batch
      where batch.scheduled_slot = '2095-01-03'
    ) then
    raise exception 'Task 7 cross-midnight claim did not start only the monthly snapshot child: %.',
      v_claim;
  end if;

  select count(*) into v_runs_before
  from public.financial_reconciliation_automatic_runs run
  where run.batch_id = v_batch_id;
  select count(*) into v_proposals_before
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.run_id = v_monthly_run_id;
  select count(*) into v_reconciliations_before
  from public.financial_reconciliations reconciliation
  where reconciliation.automatic_run_id = v_monthly_run_id;
  v_retry := public.claim_financial_reconciliation_automatic_schedule(
    '2095-01-03 00:31:00+00', 'smoke:task7-five-rule-batch'
  );
  if (v_retry->>'claimed')::boolean is distinct from true
    or (v_retry->>'resumed')::boolean is distinct from true
    or v_retry->>'batchId' is distinct from v_batch_id::text
    or (v_retry->>'batchRulePosition')::integer is distinct from 2
    or v_retry#>>'{run,runId}' is distinct from v_monthly_run_id::text
    or v_retry#>>'{run,status}' is distinct from 'analyzing'
    or v_retry#>>'{run,batchId}' is distinct from v_batch_id::text
    or (v_retry#>>'{run,batchRulePosition}')::integer is distinct from 2
    or (select count(*)
        from public.financial_reconciliation_automatic_runs run
        where run.batch_id = v_batch_id) is distinct from v_runs_before
    or (select count(*)
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = v_monthly_run_id) is distinct from
      v_proposals_before
    or (select count(*)
        from public.financial_reconciliations reconciliation
        where reconciliation.automatic_run_id = v_monthly_run_id) is distinct from
      v_reconciliations_before then
    raise exception 'Task 7 monthly retry duplicated batch, proposal, or reconciliation work.';
  end if;

  if v_month_count > 0 then
    select month_page.calendar_month,
           month_page.technical_base_source_id
    into strict v_first_month
    from public.financial_reconciliation_automatic_monthly_income_page(
      null, 1
    ) month_page;
    update public.financial_reconciliation_automatic_runs
    set analysis_cursor_date = v_first_month.calendar_month,
        analysis_cursor_id = v_first_month.technical_base_source_id,
        analysis_processed = 1,
        analysis_total = greatest(analysis_total, 1),
        updated_at = now()
    where id = v_monthly_run_id;
    v_cursor_retry := public.claim_financial_reconciliation_automatic_schedule(
      '2095-01-03 00:32:00+00', 'smoke:task7-five-rule-batch'
    );
    if (v_cursor_retry->>'claimed')::boolean is distinct from true
      or (v_cursor_retry->>'resumed')::boolean is distinct from true
      or v_cursor_retry->>'batchId' is distinct from v_batch_id::text
      or (v_cursor_retry->>'batchRulePosition')::integer is distinct from 2
      or v_cursor_retry#>>'{run,runId}' is distinct from
        v_monthly_run_id::text
      or v_cursor_retry#>>'{run,status}' is distinct from 'analyzing'
      or v_cursor_retry#>>'{run,batchId}' is distinct from v_batch_id::text
      or (v_cursor_retry#>>'{run,batchRulePosition}')::integer is distinct from
        2
      or v_cursor_retry#>>'{run,analysisCursorDate}' is distinct from
        v_first_month.calendar_month::text
      or v_cursor_retry#>>'{run,analysisCursorId}' is distinct from
        v_first_month.technical_base_source_id::text
      or v_cursor_retry#>>'{run,analysisProcessed}' is distinct from '1'
      or (select count(*)
          from public.financial_reconciliation_automatic_runs run
          where run.batch_id = v_batch_id) is distinct from 2 then
      raise exception 'Task 7 monthly retry did not preserve the stable month cursor.';
    end if;
  end if;

  update public.financial_reconciliation_automatic_runs
  set status = 'failed',
      error_summary = 'secret task7 monthly database failure',
      error_detail = 'secret task7 monthly stack',
      analysis_error_code = 'analysis_continuation_failed',
      analysis_error_at = now(),
      finished_at = now(),
      updated_at = now()
  where id = v_monthly_run_id;

  for v_position in 3..5 loop
    v_claim := public.claim_financial_reconciliation_automatic_schedule(
      timestamptz '2095-01-03 00:33:00+00'
        + make_interval(mins => v_position),
      'smoke:task7-five-rule-batch'
    );
    select run.id into strict v_run_id
    from public.financial_reconciliation_automatic_runs run
    where run.batch_id = v_batch_id
      and run.batch_rule_position = v_position;
    if (v_claim->>'claimed')::boolean is distinct from true
      or (v_claim->>'resumed')::boolean is distinct from false
      or v_claim->>'batchId' is distinct from v_batch_id::text
      or (v_claim->>'batchRulePosition')::integer is distinct from v_position
      or (v_claim->>'batchRuleCount')::integer is distinct from 5
      or v_claim#>>'{run,runId}' is distinct from v_run_id::text
      or v_claim#>>'{run,status}' is distinct from 'analyzing'
      or v_claim#>>'{run,batchId}' is distinct from v_batch_id::text
      or (v_claim#>>'{run,batchRulePosition}')::integer is distinct from
        v_position
      or v_claim#>>'{run,batchRuleKey}' is distinct from
        v_expected_rules[v_position]
      or (select count(*)
          from public.financial_reconciliation_automatic_runs run
          where run.batch_id = v_batch_id) is distinct from v_position then
      raise exception 'Task 7 failed monthly child did not advance to position %: %.',
        v_position, v_claim;
    end if;
    update public.financial_reconciliation_automatic_runs
    set status = 'completed',
        analysis_completed_at = coalesce(analysis_completed_at, now()),
        counts = '{"bases":0,"completed":0,"failed":0}'::jsonb,
        finished_at = now(),
        updated_at = now()
    where id = v_run_id;
  end loop;

  v_complete := public.claim_financial_reconciliation_automatic_schedule(
    '2095-01-03 01:00:00+00', 'smoke:task7-five-rule-batch'
  );
  v_settings := public.get_financial_reconciliation_automation_settings();
  if (v_complete->>'claimed')::boolean is distinct from false
    or v_complete->>'reason' is distinct from 'batch_complete'
    or v_complete->>'batchId' is distinct from v_batch_id::text
    or (select count(*)
        from public.financial_reconciliation_automatic_runs run
        where run.batch_id = v_batch_id) is distinct from 5
    or (select count(distinct run.batch_rule_position)
        from public.financial_reconciliation_automatic_runs run
        where run.batch_id = v_batch_id) is distinct from 5
    or (select count(distinct run.batch_rule_key)
        from public.financial_reconciliation_automatic_runs run
        where run.batch_id = v_batch_id) is distinct from 5
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_batches batch
      where batch.id = v_batch_id
        and batch.status = 'partial'
        and batch.rule_snapshot = v_snapshot
        and batch.counts @> '{"ruleCount":5,"childCount":5,"completedChildren":4,"failedChildren":1}'::jsonb
    )
    or v_settings#>>'{last_scheduled_batch,id}' is distinct from v_batch_id::text
    or v_settings#>>'{last_scheduled_batch,status}' is distinct from 'partial'
    or v_settings::text like '%secret task7 monthly%' then
    raise exception 'Task 7 five-rule parent did not finish as a sanitized partial batch.';
  end if;
end
$$;

do $$
declare
  v_signature text;
begin
  foreach v_signature in array array[
    'public.replace_financial_reconciliation_automation_settings(jsonb,jsonb,text)',
    'public.claim_financial_reconciliation_automatic_schedule(timestamptz,text)'
  ] loop
    if not (
      select procedure.prosecdef
        and coalesce(procedure.proconfig, '{}'::text[])
          @> array['search_path=public, pg_temp']
      from pg_proc procedure
      where procedure.oid = v_signature::regprocedure
    )
      or has_function_privilege('anon', v_signature, 'EXECUTE')
      or has_function_privilege('authenticated', v_signature, 'EXECUTE')
      or not has_function_privilege('service_role', v_signature, 'EXECUTE') then
      raise exception 'Task 7 five-rule scheduler RPC security changed for %.',
        v_signature;
    end if;
  end loop;
end
$$;

-- Task 2: install and reapply the two immutable FDM/Bank definitions,
-- disabled configurations, exact supporting indexes, and seven-rule Settings RPC.
create temporary table task2_existing_config_baseline on commit drop as
select config.rule_key, to_jsonb(config) as row_snapshot
from public.financial_reconciliation_automatic_rule_configs config;

\ir ../supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql
\ir ../supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql

do $$
declare
  v_bank_definition jsonb := jsonb_build_object(
    'strategy', 'bounded_exact_combination',
    'sourceAccount', 'Bank Transfer',
    'maxSourceRecords', 10,
    'candidatePoolLimit', 60,
    'stateLimit', 250000,
    'evidenceGroupLimit', 12,
    'amountMode', 'signed_integer_cents',
    'dateMode', 'inclusive_days'
  );
  v_adyen_definition jsonb := jsonb_build_object(
    'strategy', 'closed_calendar_month',
    'bankDescriptionContains', 'Adyen',
    'fdmAccount', 'Adyen',
    'requiresBothSides', true,
    'monthMarkerDays', 31
  );
  v_bank_logic text := 'Exactly one CGD Bank Statement record is matched to one through ten eligible FDM Bank Transfer records with opposite signed totals that equal zero exactly in integer cents within the inclusive configured date window.';
  v_adyen_logic text := 'Every eligible unlocked CGD Bank Statement and FDM Adyen record in the same closed calendar month forms one proposal; both sides are required and the signed difference must be within the configured allowance.';
  v_existing_max_priority integer;
  v_bank_priority integer;
  v_adyen_priority integer;
begin
  if (select count(*)
      from public.financial_reconciliation_automatic_rule_definitions definition
      where definition.rule_key =
          'fdm_bank_transfer_cgd_bank_statement_combination'
        and definition.version = 1) <> 1
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_definitions definition
      where definition.rule_key =
          'fdm_bank_transfer_cgd_bank_statement_combination'
        and definition.version = 1
        and definition.display_name =
          'FDM Accounts – Bank Reservation Payments'
        and definition.base_source_type = 'import_fdm_accounts'
        and definition.destination_source_types =
          '["import_cgd_extrato_ordem"]'::jsonb
        and definition.logic_description = v_bank_logic
        and definition.definition = v_bank_definition
    ) then
    raise exception 'Task 2 Bank Reservation immutable definition differs from the approved contract.';
  end if;

  if (select count(*)
      from public.financial_reconciliation_automatic_rule_definitions definition
      where definition.rule_key =
          'cgd_bank_statement_fdm_adyen_monthly_payments'
        and definition.version = 1) <> 1
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_definitions definition
      where definition.rule_key =
          'cgd_bank_statement_fdm_adyen_monthly_payments'
        and definition.version = 1
        and definition.display_name =
          'FDM Accounts – Adyen Reservation Payments'
        and definition.base_source_type = 'import_cgd_extrato_ordem'
        and definition.destination_source_types =
          '["import_fdm_accounts"]'::jsonb
        and definition.logic_description = v_adyen_logic
        and definition.definition = v_adyen_definition
    ) then
    raise exception 'Task 2 Adyen immutable definition differs from the approved contract.';
  end if;

  if not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_configs config
      where config.rule_key =
          'fdm_bank_transfer_cgd_bank_statement_combination'
        and config.rule_version = 1
        and not config.enabled
        and not config.allow_manual_execution
        and not config.include_in_scheduled_batch
        and config.difference_allowed = 0
        and config.max_difference_days = 3
    ) or not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_configs config
      where config.rule_key =
          'cgd_bank_statement_fdm_adyen_monthly_payments'
        and config.rule_version = 1
        and not config.enabled
        and not config.allow_manual_execution
        and not config.include_in_scheduled_batch
        and config.difference_allowed = 2000
        and config.max_difference_days = 31
    ) then
    raise exception 'Task 2 new configurations were not installed disabled with approved defaults.';
  end if;

  if exists (
      select 1
      from task2_existing_config_baseline baseline
      join public.financial_reconciliation_automatic_rule_configs config
        on config.rule_key = baseline.rule_key
      where to_jsonb(config) is distinct from baseline.row_snapshot
    ) or (select count(*) from task2_existing_config_baseline) <> 5 then
    raise exception 'Task 2 migration changed an existing managed configuration.';
  end if;

  select max((baseline.row_snapshot->>'priority')::integer)
  into strict v_existing_max_priority
  from task2_existing_config_baseline baseline;
  select config.priority into strict v_bank_priority
  from public.financial_reconciliation_automatic_rule_configs config
  where config.rule_key =
    'fdm_bank_transfer_cgd_bank_statement_combination';
  select config.priority into strict v_adyen_priority
  from public.financial_reconciliation_automatic_rule_configs config
  where config.rule_key =
    'cgd_bank_statement_fdm_adyen_monthly_payments';
  if v_bank_priority <= v_existing_max_priority
    or v_adyen_priority <= v_existing_max_priority
    or v_bank_priority >= v_adyen_priority then
    raise exception 'Task 2 configs were not appended after the existing five in deterministic relative order.';
  end if;

  if (select count(*)
      from pg_index index_row
      join pg_class index_class on index_class.oid = index_row.indexrelid
      where index_class.relname in (
        'financial_reconciliation_fdm_bank_transfer_lookup_idx',
        'financial_reconciliation_fdm_adyen_lookup_idx',
        'financial_reconciliation_bank_date_amount_lookup_idx'
      )
        and index_row.indisvalid
        and index_row.indisready) <> 3 then
    raise exception 'Task 2 exact supporting indexes are missing, invalid, or not ready.';
  end if;
end
$$;

create or replace function pg_temp.task2_installation_state()
returns jsonb
language sql
stable
set search_path = public, pg_temp
as $$
  select jsonb_build_object(
    'definitions', coalesce((
      select jsonb_agg(to_jsonb(definition)
                       order by definition.rule_key, definition.version)
      from public.financial_reconciliation_automatic_rule_definitions definition
      where definition.rule_key in (
        'fdm_bank_transfer_cgd_bank_statement_combination',
        'cgd_bank_statement_fdm_adyen_monthly_payments'
      )
    ), '[]'::jsonb),
    'configs', coalesce((
      select jsonb_agg(to_jsonb(config) order by config.rule_key)
      from public.financial_reconciliation_automatic_rule_configs config
    ), '[]'::jsonb),
    'schedule', (
      select to_jsonb(schedule)
      from public.financial_reconciliation_automatic_schedule schedule
      where schedule.id = true
    ),
    'constraint', (
      select jsonb_build_object(
        'type', constraint_row.contype,
        'validated', constraint_row.convalidated,
        'definition', pg_get_constraintdef(constraint_row.oid, true)
      )
      from pg_constraint constraint_row
      where constraint_row.conrelid =
          'public.financial_reconciliation_automatic_rule_configs'::regclass
        and constraint_row.conname =
          'financial_reconciliation_rule_configs_fdm_bank_adyen_check'
    ),
    'indexes', coalesce((
      select jsonb_agg(jsonb_build_object(
        'name', index_class.relname,
        'definition', pg_get_indexdef(index_row.indexrelid),
        'valid', index_row.indisvalid,
        'ready', index_row.indisready
      ) order by index_class.relname)
      from pg_index index_row
      join pg_class index_class on index_class.oid = index_row.indexrelid
      join pg_namespace namespace_row on namespace_row.oid = index_class.relnamespace
      where namespace_row.nspname = 'public'
        and index_class.relname in (
          'financial_reconciliation_fdm_bank_transfer_lookup_idx',
          'financial_reconciliation_fdm_adyen_lookup_idx',
          'financial_reconciliation_bank_date_amount_lookup_idx'
        )
    ), '[]'::jsonb),
    'settingsFunction', (
      select jsonb_build_object(
        'definition', pg_get_functiondef(procedure.oid),
        'securityDefiner', procedure.prosecdef,
        'configuration', procedure.proconfig,
        'acl', procedure.proacl
      )
      from pg_proc procedure
      where procedure.oid =
        'public.replace_financial_reconciliation_automation_settings(jsonb,jsonb,text)'::regprocedure
    )
  )
$$;

create temporary table task2_installation_baseline on commit drop as
select pg_temp.task2_installation_state() as state;

savepoint task2_conflicting_definition_fixture;

update public.financial_reconciliation_automatic_rule_definitions definition
set definition = jsonb_set(definition.definition, '{stateLimit}', '250001'::jsonb)
where definition.rule_key =
    'fdm_bank_transfer_cgd_bank_statement_combination'
  and definition.version = 1;

\set ON_ERROR_STOP off
\ir ../supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql
\set task2_conflicting_definition_rejected :ERROR
\set ON_ERROR_STOP on

rollback to savepoint task2_conflicting_definition_fixture;

\if :task2_conflicting_definition_rejected
\else
  \echo 'Task 2 migration accepted a conflicting immutable definition for the managed Bank Reservation key/version.'
  \quit 1
\endif

\ir ../supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql

do $$
begin
  if (select state from task2_installation_baseline)
      is distinct from pg_temp.task2_installation_state() then
    raise exception 'Task 2 definition conflict did not reject and restore the complete installation state.';
  end if;
end
$$;

savepoint task2_conflicting_constraint_fixture;

alter table public.financial_reconciliation_automatic_rule_configs
  drop constraint financial_reconciliation_rule_configs_fdm_bank_adyen_check;
alter table public.financial_reconciliation_automatic_rule_configs
  add constraint financial_reconciliation_rule_configs_fdm_bank_adyen_check
  check (
    rule_key <> 'fdm_bank_transfer_cgd_bank_statement_combination'
    or max_difference_days between 0 and 89
  ) not valid;

\set ON_ERROR_STOP off
\ir ../supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql
\set task2_conflicting_constraint_rejected :ERROR
\set ON_ERROR_STOP on

rollback to savepoint task2_conflicting_constraint_fixture;

\if :task2_conflicting_constraint_rejected
\else
  \echo 'Task 2 migration accepted a conflicting same-named managed configuration constraint.'
  \quit 1
\endif

\ir ../supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql

do $$
begin
  if (select state from task2_installation_baseline)
      is distinct from pg_temp.task2_installation_state() then
    raise exception 'Task 2 constraint conflict did not reject and restore the complete installation state.';
  end if;
end
$$;

savepoint task2_conflicting_index_fixture;

drop index public.financial_reconciliation_fdm_bank_transfer_lookup_idx;
create index financial_reconciliation_fdm_bank_transfer_lookup_idx
  on public.import_fdm_accounts (event_date, id)
  where account = 'Bank Transfer';

\set ON_ERROR_STOP off
\ir ../supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql
\set task2_conflicting_index_rejected :ERROR
\set ON_ERROR_STOP on

rollback to savepoint task2_conflicting_index_fixture;

\if :task2_conflicting_index_rejected
\else
  \echo 'Task 2 migration accepted a conflicting same-named FDM Bank Transfer lookup index.'
  \quit 1
\endif

\ir ../supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql

do $$
begin
  if (select state from task2_installation_baseline)
      is distinct from pg_temp.task2_installation_state() then
    raise exception 'Task 2 index conflict did not reject and restore the complete installation state.';
  end if;
end
$$;

set constraints financial_reconciliation_automatic_rule_configs_priority_key deferred;
update public.financial_reconciliation_automatic_rule_configs config
set enabled = true,
    allow_manual_execution = true,
    include_in_scheduled_batch = true,
    difference_allowed = case
      when config.rule_key =
          'fdm_bank_transfer_cgd_bank_statement_combination' then 0
      else 1234.56
    end,
    max_difference_days = case
      when config.rule_key =
          'fdm_bank_transfer_cgd_bank_statement_combination' then 90
      else 31
    end,
    priority = case
      when config.rule_key =
          'fdm_bank_transfer_cgd_bank_statement_combination' then 7
      else 6
    end,
    updated_by = 'smoke:task2-administrator'
where config.rule_key in (
  'fdm_bank_transfer_cgd_bank_statement_combination',
  'cgd_bank_statement_fdm_adyen_monthly_payments'
);

\ir ../supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql

do $$
begin
  if not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_configs config
      where config.rule_key =
          'fdm_bank_transfer_cgd_bank_statement_combination'
        and config.rule_version = 1
        and config.enabled
        and config.allow_manual_execution
        and config.include_in_scheduled_batch
        and config.difference_allowed = 0
        and config.max_difference_days = 90
        and config.priority = 7
        and config.updated_by = 'smoke:task2-administrator'
    ) or not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_configs config
      where config.rule_key =
          'cgd_bank_statement_fdm_adyen_monthly_payments'
        and config.rule_version = 1
        and config.enabled
        and config.allow_manual_execution
        and config.include_in_scheduled_batch
        and config.difference_allowed = 1234.56
        and config.max_difference_days = 31
        and config.priority = 6
        and config.updated_by = 'smoke:task2-administrator'
    ) then
    raise exception 'Task 2 migration reapply overwrote administrator-controlled settings.';
  end if;

  if exists (
    select 1
    from task2_existing_config_baseline baseline
    join public.financial_reconciliation_automatic_rule_configs config
      on config.rule_key = baseline.rule_key
    where to_jsonb(config) is distinct from baseline.row_snapshot
  ) then
    raise exception 'Task 2 migration reapply changed an existing managed rule or its relative priority.';
  end if;
end
$$;

create or replace function pg_temp.task2_assert_settings_rejected(
  p_rules jsonb,
  p_expected_error text,
  p_label text
)
returns void
language plpgsql
set search_path = public, pg_temp
as $$
declare
  v_before_schedule jsonb;
  v_after_schedule jsonb;
  v_before_configs jsonb;
  v_after_configs jsonb;
  v_rejected boolean := false;
begin
  select to_jsonb(schedule) into strict v_before_schedule
  from public.financial_reconciliation_automatic_schedule schedule
  where schedule.id = true;
  select jsonb_agg(to_jsonb(config) order by config.rule_key)
  into strict v_before_configs
  from public.financial_reconciliation_automatic_rule_configs config;

  begin
    perform public.replace_financial_reconciliation_automation_settings(
      jsonb_build_object(
        'enabled', not (v_before_schedule->>'enabled')::boolean,
        'time_of_day', '05:45',
        'time_zone', 'Europe/Lisbon'
      ),
      p_rules,
      'smoke:task2-invalid'
    );
  exception when others then
    v_rejected := true;
    if sqlerrm is distinct from p_expected_error then
      raise exception 'Task 2 % failed with unexpected error: %', p_label, sqlerrm;
    end if;
  end;

  select to_jsonb(schedule) into strict v_after_schedule
  from public.financial_reconciliation_automatic_schedule schedule
  where schedule.id = true;
  select jsonb_agg(to_jsonb(config) order by config.rule_key)
  into strict v_after_configs
  from public.financial_reconciliation_automatic_rule_configs config;

  if not v_rejected then
    raise exception 'Task 2 % was accepted.', p_label;
  end if;
  if v_after_schedule is distinct from v_before_schedule
    or v_after_configs is distinct from v_before_configs then
    raise exception 'Task 2 % changed schedule or managed config rows despite rejection.', p_label;
  end if;
end
$$;

do $$
declare
  v_rules jsonb;
  v_invalid jsonb;
  v_settings jsonb;
  v_key text;
  v_field text;
begin
  select jsonb_agg(jsonb_build_object(
    'rule_key', config.rule_key,
    'rule_version', config.rule_version,
    'enabled', config.enabled,
    'allow_manual_execution', config.allow_manual_execution,
    'include_in_scheduled_batch', config.include_in_scheduled_batch,
    'difference_allowed', to_char(config.difference_allowed,
      'FM999999999999990.00'),
    'max_difference_days', config.max_difference_days,
    'priority', config.priority
  ) order by config.priority, config.rule_key)
  into strict v_rules
  from public.financial_reconciliation_automatic_rule_configs config;

  v_settings := public.replace_financial_reconciliation_automation_settings(
    '{"enabled":true,"time_of_day":"04:15","time_zone":"Europe/Lisbon"}'::jsonb,
    v_rules,
    'smoke:task2-seven-rule-success'
  );
  if jsonb_array_length(v_settings->'rules') <> 7
    or (select count(*)
        from public.financial_reconciliation_automatic_rule_configs) <> 7 then
    raise exception 'Task 2 Settings replacement did not atomically return all seven managed rules.';
  end if;

  foreach v_key in array array[
    'fdm_bank_transfer_cgd_bank_statement_combination',
    'cgd_bank_statement_fdm_adyen_monthly_payments'
  ] loop
    select coalesce(jsonb_agg(rule.value order by rule.ordinality), '[]'::jsonb)
    into v_invalid
    from jsonb_array_elements(v_rules) with ordinality rule(value, ordinality)
    where rule.value->>'rule_key' <> v_key;
    perform pg_temp.task2_assert_settings_rejected(
      v_invalid,
      'Automatic rules payload must contain the seven managed rule objects.',
      format('missing managed key %s', v_key)
    );
  end loop;

  select jsonb_agg(case
    when rule.value->>'rule_key' =
        'cgd_bank_statement_fdm_adyen_monthly_payments'
      then jsonb_set(rule.value, '{rule_key}',
        '"fdm_bank_transfer_cgd_bank_statement_combination"'::jsonb)
    else rule.value
  end order by rule.ordinality)
  into v_invalid
  from jsonb_array_elements(v_rules) with ordinality rule(value, ordinality);
  perform pg_temp.task2_assert_settings_rejected(
    v_invalid, 'Duplicate automatic rule.', 'duplicate key'
  );

  select jsonb_agg(case
    when rule.value->>'rule_key' =
        'cgd_bank_statement_fdm_adyen_monthly_payments'
      then jsonb_set(rule.value, '{priority}', to_jsonb((
        select (bank_rule.value->>'priority')::integer
        from jsonb_array_elements(v_rules) bank_rule(value)
        where bank_rule.value->>'rule_key' =
          'fdm_bank_transfer_cgd_bank_statement_combination'
      )))
    else rule.value
  end order by rule.ordinality)
  into v_invalid
  from jsonb_array_elements(v_rules) with ordinality rule(value, ordinality);
  perform pg_temp.task2_assert_settings_rejected(
    v_invalid, 'Duplicate automatic rule priority.', 'duplicate priority'
  );

  select jsonb_agg(case
    when rule.value->>'rule_key' =
        'fdm_bank_transfer_cgd_bank_statement_combination'
      then jsonb_set(rule.value, '{rule_version}', '2'::jsonb)
    else rule.value
  end order by rule.ordinality)
  into v_invalid
  from jsonb_array_elements(v_rules) with ordinality rule(value, ordinality);
  perform pg_temp.task2_assert_settings_rejected(
    v_invalid, 'Automatic rule/version is invalid.', 'unsupported version'
  );

  select jsonb_agg(case
    when rule.value->>'rule_key' =
        'fdm_bank_transfer_cgd_bank_statement_combination'
      then jsonb_set(rule.value, '{difference_allowed}', '"0.01"'::jsonb)
    else rule.value
  end order by rule.ordinality)
  into v_invalid
  from jsonb_array_elements(v_rules) with ordinality rule(value, ordinality);
  perform pg_temp.task2_assert_settings_rejected(
    v_invalid,
    'Bank Reservation automatic rule requires zero difference allowed.',
    'Bank Reservation nonzero allowance'
  );

  select jsonb_agg(case
    when rule.value->>'rule_key' =
        'fdm_bank_transfer_cgd_bank_statement_combination'
      then jsonb_set(rule.value, '{max_difference_days}', '91'::jsonb)
    else rule.value
  end order by rule.ordinality)
  into v_invalid
  from jsonb_array_elements(v_rules) with ordinality rule(value, ordinality);
  perform pg_temp.task2_assert_settings_rejected(
    v_invalid, 'Automatic rule values are invalid.', 'Bank Reservation day 91'
  );

  select jsonb_agg(case
    when rule.value->>'rule_key' =
        'cgd_bank_statement_fdm_adyen_monthly_payments'
      then jsonb_set(rule.value, '{max_difference_days}', '30'::jsonb)
    else rule.value
  end order by rule.ordinality)
  into v_invalid
  from jsonb_array_elements(v_rules) with ordinality rule(value, ordinality);
  perform pg_temp.task2_assert_settings_rejected(
    v_invalid,
    'Adyen automatic rule requires the fixed 31-day calendar-month property.',
    'Adyen non-calendar day'
  );

  select jsonb_agg(case
    when rule.value->>'rule_key' =
        'cgd_bank_statement_fdm_adyen_monthly_payments'
      then jsonb_set(rule.value, '{difference_allowed}', '"-0.01"'::jsonb)
    else rule.value
  end order by rule.ordinality)
  into v_invalid
  from jsonb_array_elements(v_rules) with ordinality rule(value, ordinality);
  perform pg_temp.task2_assert_settings_rejected(
    v_invalid, 'Automatic rule values are invalid.', 'negative Adyen allowance'
  );

  foreach v_field in array array[
    'definition', 'base_source_type', 'destination_source_types', 'strategy',
    'maxSourceRecords', 'candidatePoolLimit', 'stateLimit',
    'evidenceGroupLimit'
  ] loop
    select jsonb_agg(case
      when rule.value->>'rule_key' =
          'fdm_bank_transfer_cgd_bank_statement_combination'
        then rule.value || jsonb_build_object(v_field, 'attempted mutation')
      else rule.value
    end order by rule.ordinality)
    into v_invalid
    from jsonb_array_elements(v_rules) with ordinality rule(value, ordinality);
    perform pg_temp.task2_assert_settings_rejected(
      v_invalid, 'Automatic rule fields are invalid.',
      format('immutable field %s mutation', v_field)
    );
  end loop;
end
$$;

do $$
declare
  v_signature text :=
    'public.replace_financial_reconciliation_automation_settings(jsonb,jsonb,text)';
begin
  if not (
      select procedure.prosecdef
        and coalesce(procedure.proconfig, '{}'::text[])
          @> array['search_path=public, pg_temp']
      from pg_proc procedure
      where procedure.oid = v_signature::regprocedure
    )
    or exists (
      select 1
      from pg_proc procedure,
           lateral aclexplode(coalesce(
             procedure.proacl,
             acldefault('f', procedure.proowner)
           )) privilege
      where procedure.oid = v_signature::regprocedure
        and privilege.grantee = 0
        and privilege.privilege_type = 'EXECUTE'
    )
    or has_function_privilege('anon', v_signature, 'EXECUTE')
    or has_function_privilege('authenticated', v_signature, 'EXECUTE')
    or not has_function_privilege('service_role', v_signature, 'EXECUTE') then
    raise exception 'Task 2 seven-rule Settings RPC security is invalid.';
  end if;
end
$$;

-- Task 3 Bank Reservation helper classifications and hard bounds
-- Isolate the bounded-search fixtures from every source row exercised above by
-- taking rollback-only reconciliation locks on the remaining unlocked rows.
do $$
declare
  v_reconciliation_id uuid;
begin
  insert into public.financial_reconciliations (
    status, base_source_type, matching_source_types, created_by
  ) values (
    'started', 'import_fdm_accounts', '["import_cgd_extrato_ordem"]'::jsonb,
    'smoke:task3-isolation'
  ) returning id into v_reconciliation_id;

  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  )
  select v_reconciliation_id, 'import_cgd_extrato_ordem', bank.id,
         bank.montante, 'smoke:task3-isolation'
  from public.import_cgd_extrato_ordem bank
  where bank.montante is not null
    and not exists (
      select 1 from public.financial_reconciliation_items locked
      where locked.source_type = 'import_cgd_extrato_ordem'
        and locked.source_id = bank.id
    );

  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  )
  select v_reconciliation_id, 'import_fdm_accounts', fdm.id,
         fdm.amount, 'smoke:task3-isolation'
  from public.import_fdm_accounts fdm
  where fdm.amount is not null
    and not exists (
      select 1 from public.financial_reconciliation_items locked
      where locked.source_type = 'import_fdm_accounts'
        and locked.source_id = fdm.id
    );
end
$$;

alter table public.import_fdm_accounts alter column amount drop not null;

insert into public.import_cgd_extrato_ordem (
  id, import_batch, row_key, data, descritivo, montante
) values
  ('b3100000-0000-0000-0000-000000000001', 'smoke-task3-bank',
   'task3-bank-unique-1', date '2100-01-10', 'unique one', -10.00),
  ('b3100000-0000-0000-0000-000000000002', 'smoke-task3-bank',
   'task3-bank-unique-2', date '2100-02-10', 'unique two', -30.00),
  ('b3100000-0000-0000-0000-000000000003', 'smoke-task3-bank',
   'task3-bank-unique-10', date '2100-03-10', 'unique ten', -10.00),
  ('b3100000-0000-0000-0000-000000000004', 'smoke-task3-bank',
   'task3-bank-exclusions', date '2100-04-10', 'inclusive boundary', -50.00),
  ('b3100000-0000-0000-0000-000000000005', 'smoke-task3-bank',
   'task3-bank-eleven', date '2100-05-10', 'eleven rejected', -11.00),
  ('b3100000-0000-0000-0000-000000000006', 'smoke-task3-bank',
   'task3-bank-multiple', date '2100-06-10', 'multiple groups', -10.00),
  ('b3100000-0000-0000-0000-000000000007', 'smoke-task3-bank',
   'task3-bank-pool-limit', date '2120-01-10', 'pool ceiling', -100.00),
  ('b3100000-0000-0000-0000-000000000008', 'smoke-task3-bank',
   'task3-bank-group-limit', date '2110-01-10', 'group ceiling', -10.00),
  ('b3100000-0000-0000-0000-000000000009', 'smoke-task3-bank',
   'task3-bank-state-limit', date '2130-01-10', 'state ceiling', -1000.00),
  ('b3100000-0000-0000-0000-000000000010', 'smoke-task3-bank',
   'task3-bank-overlap-a', date '2160-01-10', 'shared FDM a', -25.00),
  ('b3100000-0000-0000-0000-000000000011', 'smoke-task3-bank',
   'task3-bank-overlap-b', date '2160-01-11', 'shared FDM b', -25.00),
  ('b3100000-0000-0000-0000-000000000012', 'smoke-task3-bank',
   'task3-bank-locked', date '2100-07-10', 'locked bank', -10.00),
  ('b3100000-0000-0000-0000-000000000013', 'smoke-task3-bank',
   'task3-bank-null-date', null, 'null date', -10.00),
  ('b3100000-0000-0000-0000-000000000014', 'smoke-task3-bank',
   'task3-bank-null-amount', date '2100-08-10', 'null amount', null),
  ('b3100000-0000-0000-0000-000000000015', 'smoke-task3-bank',
   'task3-bank-pre-floor', date '2025-12-31', 'pre floor', -10.00);

insert into public.import_cgd_extrato_ordem (
  id, import_batch, row_key, source_row_number, data, descritivo, montante
)
select
  ('b3900000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
  'smoke-task3-bank', 'task3-bank-page-' || series, series,
  date '2150-01-01' + (series - 1), 'stable page ' || series,
  -(1000 + series)::numeric
from generate_series(1, 30) series;

insert into public.import_fdm_accounts (
  id, import_batch, account, date_time_raw, event_date, category,
  amount, description
) values
  ('c3100000-0000-0000-0000-000000000001', 'smoke-task3-fdm',
   'Bank Transfer', '2100-01-10', date '2100-01-10', 'Reservation',
   10.00, 'unique one'),
  ('c3100000-0000-0000-0000-000000000010', 'smoke-task3-fdm',
   'Bank Transfer', '2100-02-10', date '2100-02-10', 'Reservation',
   20.00, 'unique two later'),
  ('c3100000-0000-0000-0000-000000000011', 'smoke-task3-fdm',
   'Bank Transfer', '2100-02-09', date '2100-02-09', 'Reservation',
   10.00, 'unique two canonical'),
  ('c3120000-0000-0000-0000-000000000001', 'smoke-task3-fdm',
   'Bank Transfer', '2100-04-13', date '2100-04-13', 'Reservation',
   50.00, 'inclusive day boundary'),
  ('c3120000-0000-0000-0000-000000000002', 'smoke-task3-fdm',
   'Bank Transfer', '2100-04-10', date '2100-04-10', 'Reservation',
   -50.00, 'same sign'),
  ('c3120000-0000-0000-0000-000000000003', 'smoke-task3-fdm',
   'Bank Transfer', '2100-04-10', date '2100-04-10', 'Reservation',
   49.99, 'one cent short'),
  ('c3120000-0000-0000-0000-000000000004', 'smoke-task3-fdm',
   'Bank Transfers', '2100-04-10', date '2100-04-10', 'Reservation',
   50.00, 'near account'),
  ('c3120000-0000-0000-0000-000000000005', 'smoke-task3-fdm',
   'Bank Transfer', '2100-04-10', date '2100-04-10', 'Reservation',
   null, 'null amount'),
  ('c3120000-0000-0000-0000-000000000006', 'smoke-task3-fdm',
   'Bank Transfer', '', null, 'Reservation', 50.00, 'null date'),
  ('c3120000-0000-0000-0000-000000000007', 'smoke-task3-fdm',
   'Bank Transfer', '2100-04-14', date '2100-04-14', 'Reservation',
   50.00, 'outside day'),
  ('c3120000-0000-0000-0000-000000000008', 'smoke-task3-fdm',
   'Bank Transfer', '2100-04-10', date '2100-04-10', 'Reservation',
   50.00, 'locked FDM'),
  ('c3140000-0000-0000-0000-000000000001', 'smoke-task3-fdm',
   'Bank Transfer', '2100-06-10', date '2100-06-10', 'Reservation',
   10.00, 'multiple singleton'),
  ('c3140000-0000-0000-0000-000000000002', 'smoke-task3-fdm',
   'Bank Transfer', '2100-06-10', date '2100-06-10', 'Reservation',
   4.00, 'multiple pair a'),
  ('c3140000-0000-0000-0000-000000000003', 'smoke-task3-fdm',
   'Bank Transfer', '2100-06-10', date '2100-06-10', 'Reservation',
   6.00, 'multiple pair b'),
  ('c315f000-0000-0000-0000-000000000001', 'smoke-task3-fdm',
   'Bank Transfer', '2110-01-10', date '2110-01-10', 'Reservation',
   3.00, 'earlier non-qualifying group candidate'),
  ('c3180000-0000-0000-0000-000000000001', 'smoke-task3-fdm',
   'Bank Transfer', '2160-01-10', date '2160-01-10', 'Reservation',
   25.00, 'shared FDM');

insert into public.import_fdm_accounts (
  id, import_batch, source_row_number, account, date_time_raw, event_date,
  category, amount, description
)
select
  ('c3110000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
  'smoke-task3-fdm', series, 'Bank Transfer', '2100-03-10',
  date '2100-03-10', 'Reservation', 1.00, 'ten member ' || series
from generate_series(1, 10) series
union all
select
  ('c3130000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
  'smoke-task3-fdm', 100 + series, 'Bank Transfer', '2100-05-10',
  date '2100-05-10', 'Reservation', 1.00, 'eleven member ' || series
from generate_series(1, 11) series
union all
select
  ('c3150000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
  'smoke-task3-fdm', 200 + series, 'Bank Transfer', '2120-01-10',
  date '2120-01-10', 'Reservation', 1.00, 'pool member ' || series
from generate_series(1, 61) series
union all
select
  ('c3160000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
  'smoke-task3-fdm', 300 + series, 'Bank Transfer', '2110-01-10',
  date '2110-01-10', 'Reservation', 10.00, 'group member ' || series
from generate_series(1, 13) series
union all
select
  ('c3170000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
  'smoke-task3-fdm', 400 + series, 'Bank Transfer', '2130-01-10',
  date '2130-01-10', 'Reservation', 1.00, 'state member ' || series
from generate_series(1, 19) series
union all
select
  ('c3900000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
  'smoke-task3-fdm', 500 + series, 'Bank Transfer',
  (date '2150-01-01' + (series - 1))::text,
  date '2150-01-01' + (series - 1), 'Reservation',
  (1000 + series)::numeric, 'stable page ' || series
from generate_series(1, 30) series;

do $$
declare
  v_reconciliation_id uuid;
begin
  insert into public.financial_reconciliations (
    status, base_source_type, matching_source_types, created_by
  ) values (
    'started', 'import_fdm_accounts', '["import_cgd_extrato_ordem"]'::jsonb,
    'smoke:task3-explicit-locks'
  ) returning id into v_reconciliation_id;
  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  ) values
    (v_reconciliation_id, 'import_cgd_extrato_ordem',
     'b3100000-0000-0000-0000-000000000012', -10.00,
     'smoke:task3-explicit-locks'),
    (v_reconciliation_id, 'import_fdm_accounts',
     'c3120000-0000-0000-0000-000000000008', 50.00,
     'smoke:task3-explicit-locks');
end
$$;

do $$
declare
  v_group jsonb;
  v_repeat_group jsonb;
  v_last_page_one_date date;
  v_last_page_one_id uuid;
  v_page_one_ids uuid[];
  v_page_two_ids uuid[];
  v_rejected boolean;
  v_signature text;
begin
  if public.financial_reconciliation_automatic_bank_reservation_count()
      is distinct from 41 then
    raise exception 'Task 3 Bank anchor count admitted a locked, null, or pre-floor source.';
  end if;

  select array_agg(page.bank_id order by page.bank_date, page.bank_id)
  into v_page_one_ids
  from public.financial_reconciliation_automatic_bank_reservation_page(
    null, null, 25
  ) page;
  select page.bank_date, page.bank_id
  into strict v_last_page_one_date, v_last_page_one_id
  from public.financial_reconciliation_automatic_bank_reservation_page(
    null, null, 25
  ) page
  order by page.bank_date desc, page.bank_id desc
  limit 1;
  select array_agg(page.bank_id order by page.bank_date, page.bank_id)
  into v_page_two_ids
  from public.financial_reconciliation_automatic_bank_reservation_page(
    v_last_page_one_date, v_last_page_one_id, 25
  ) page;
  if cardinality(v_page_one_ids) is distinct from 25
    or cardinality(v_page_two_ids) is distinct from 16
    or cardinality(v_page_one_ids || v_page_two_ids) is distinct from 41
    or (select count(distinct page_id)
        from unnest(v_page_one_ids || v_page_two_ids) page_id)
      is distinct from 41 then
    raise exception 'Task 3 Bank anchor pages skipped, duplicated, or reordered rows.';
  end if;

  foreach v_signature in array array[
    'public.financial_reconciliation_automatic_bank_reservation_count()',
    'public.financial_reconciliation_automatic_bank_reservation_page(date,uuid,integer)',
    'public.financial_reconciliation_automatic_bank_reservation_groups(uuid,integer,integer,integer,integer)',
    'public.financial_reconciliation_continue_automatic_bank_reservation(uuid,text)'
  ] loop
    if has_function_privilege('anon', v_signature, 'EXECUTE')
        is distinct from false
      or has_function_privilege('authenticated', v_signature, 'EXECUTE')
        is distinct from false
      or has_function_privilege('service_role', v_signature, 'EXECUTE')
        is distinct from false then
      raise exception 'Task 3 private helper unexpectedly exposes EXECUTE on %.',
        v_signature;
    end if;
  end loop;

  v_group := public.financial_reconciliation_automatic_bank_reservation_groups(
    'b3100000-0000-0000-0000-000000000001', 3, 60, 250000, 12
  );
  if jsonb_typeof(v_group) is distinct from 'object'
    or (v_group ?& array[
      'classification','reason','evaluatedStates','candidateCount',
      'canonicalFdmId','canonicalFdmDate','candidateGroups'
    ]) is not true
    or v_group->'classification' is distinct from '"proposed"'::jsonb
    or v_group->'reason' is distinct from
      '"unique_qualifying_combination"'::jsonb
    or jsonb_typeof(v_group->'candidateGroups') is distinct from 'array'
    or jsonb_array_length(v_group->'candidateGroups') is distinct from 1
    or v_group#>'{candidateGroups,0,fdmTotalCents}'
      is distinct from '1000'::jsonb
    or v_group#>'{candidateGroups,0,bankAmountCents}'
      is distinct from '-1000'::jsonb
    or v_group#>'{candidateGroups,0,equationCents}'
      is distinct from '0'::jsonb then
    raise exception 'Task 3 one-member exact-cents classification is invalid: %.',
      v_group;
  end if;

  v_group := public.financial_reconciliation_automatic_bank_reservation_groups(
    'b3100000-0000-0000-0000-000000000004', 3, 60, 250000, 12
  );
  if jsonb_typeof(v_group) is distinct from 'object'
    or v_group->'classification' is distinct from '"proposed"'::jsonb
    or jsonb_typeof(v_group#>'{candidateGroups,0,fdmIds}')
      is distinct from 'array'
    or v_group#>'{candidateGroups,0,fdmIds,0}' is distinct from
      '"c3120000-0000-0000-0000-000000000001"'::jsonb
    or jsonb_array_length(v_group#>'{candidateGroups,0,fdmIds}')
      is distinct from 1 then
    raise exception 'Task 3 sign, cent, Account, null, lock, or inclusive-day filtering failed: %.',
      v_group;
  end if;

  v_group := public.financial_reconciliation_automatic_bank_reservation_groups(
    'b3100000-0000-0000-0000-000000000005', 3, 60, 250000, 12
  );
  if jsonb_typeof(v_group) is distinct from 'object'
    or v_group->'classification' is distinct from '"skipped"'::jsonb
    or v_group->'reason' is distinct from
      '"no_qualifying_combination"'::jsonb
    or v_group->'candidateGroups' is distinct from '[]'::jsonb then
    raise exception 'Task 3 search admitted an eleven-member group: %.', v_group;
  end if;

  v_group := public.financial_reconciliation_automatic_bank_reservation_groups(
    'b3100000-0000-0000-0000-000000000006', 3, 60, 250000, 12
  );
  if jsonb_typeof(v_group) is distinct from 'object'
    or v_group->'classification' is distinct from '"ambiguous"'::jsonb
    or v_group->'reason' is distinct from
      '"multiple_qualifying_combinations"'::jsonb
    or jsonb_typeof(v_group->'candidateGroups') is distinct from 'array'
    or jsonb_array_length(v_group->'candidateGroups') is distinct from 2 then
    raise exception 'Task 3 multiple-combination classification is invalid: %.',
      v_group;
  end if;

  -- Pool overflow must observe exactly the sixty-first stable candidate.
  v_group := public.financial_reconciliation_automatic_bank_reservation_groups(
    'b3100000-0000-0000-0000-000000000007', 3, 60, 250000, 12
  );
  v_repeat_group :=
    public.financial_reconciliation_automatic_bank_reservation_groups(
      'b3100000-0000-0000-0000-000000000007', 3, 60, 250000, 12
    );
  if jsonb_typeof(v_group) is distinct from 'object'
    or v_group->'classification' is distinct from '"ambiguous"'::jsonb
    or v_group->'reason' is distinct from '"candidate_limit"'::jsonb
    or v_group->'candidateCount' is distinct from '61'::jsonb
    or v_group->'evaluatedStates' is distinct from '0'::jsonb
    or v_group->'candidateGroups' is distinct from '[]'::jsonb
    or v_repeat_group is distinct from v_group then
    raise exception 'Task 3 candidate pool boundary is not exact or stable: %, %.',
      v_group, v_repeat_group;
  end if;

  -- The production state ceiling must classify after exactly 250,000 states.
  v_group := public.financial_reconciliation_automatic_bank_reservation_groups(
    'b3100000-0000-0000-0000-000000000009', 3, 60, 250000, 12
  );
  v_repeat_group :=
    public.financial_reconciliation_automatic_bank_reservation_groups(
      'b3100000-0000-0000-0000-000000000009', 3, 60, 250000, 12
    );
  if jsonb_typeof(v_group) is distinct from 'object'
    or v_group->'classification' is distinct from '"ambiguous"'::jsonb
    or v_group->'reason' is distinct from '"candidate_limit"'::jsonb
    or v_group->'candidateCount' is distinct from '19'::jsonb
    or v_group->'evaluatedStates' is distinct from '250000'::jsonb
    or v_group->'candidateGroups' is distinct from '[]'::jsonb
    or v_repeat_group is distinct from v_group then
    raise exception 'Task 3 evaluated-state boundary is not exact or stable: %, %.',
      v_group, v_repeat_group;
  end if;

  -- Twelve groups are retained; discovering the thirteenth fails closed.
  v_group := public.financial_reconciliation_automatic_bank_reservation_groups(
    'b3100000-0000-0000-0000-000000000008', 3, 60, 250000, 12
  );
  v_repeat_group :=
    public.financial_reconciliation_automatic_bank_reservation_groups(
      'b3100000-0000-0000-0000-000000000008', 3, 60, 250000, 12
    );
  if jsonb_typeof(v_group) is distinct from 'object'
    or v_group->'classification' is distinct from '"ambiguous"'::jsonb
    or v_group->'reason' is distinct from '"candidate_limit"'::jsonb
    or v_group->'candidateCount' is distinct from '14'::jsonb
    or v_group->'evaluatedStates' is distinct from '14'::jsonb
    or jsonb_typeof(v_group->'candidateGroups') is distinct from 'array'
    or jsonb_array_length(v_group->'candidateGroups') is distinct from 12
    or v_group#>'{candidateGroups,0,fdmIds,0}' is distinct from
      '"c3160000-0000-0000-0000-000000000001"'::jsonb
    or v_group->'canonicalFdmId'
      is distinct from v_group#>'{candidateGroups,0,fdmIds,0}'
    or exists (
      select 1
      from jsonb_array_elements(v_group->'candidateGroups') evidence(value)
      where evidence.value->'equationCents' is distinct from '0'::jsonb
        or jsonb_array_length(evidence.value->'fdmIds') is distinct from 1
    )
    or v_repeat_group is distinct from v_group then
    raise exception 'Task 3 thirteenth-group boundary is not exact or stable: %, %.',
      v_group, v_repeat_group;
  end if;

  v_rejected := false;
  begin
    perform public.financial_reconciliation_automatic_bank_reservation_page(
      null, null, 26
    );
  exception when others then
    v_rejected := sqlerrm =
      'Automatic Bank Reservation page size must be between 1 and 25.';
  end;
  if v_rejected is not true then
    raise exception 'Task 3 Bank page accepted an oversized limit.';
  end if;
end
$$;

create or replace function pg_temp.task3_bank_snapshot(p_days integer)
returns jsonb
language sql
stable
set search_path = public, pg_temp
as $$
  select jsonb_build_array(jsonb_build_object(
    'ruleKey', definition.rule_key,
    'ruleVersion', definition.version,
    'displayName', definition.display_name,
    'priority', config.priority,
    'differenceAllowed', 0,
    'maxDifferenceDays', p_days,
    'destinationSourceType', 'import_cgd_extrato_ordem',
    'definition', definition.definition,
    'operator', source_rule.operator
  ))
  from public.financial_reconciliation_automatic_rule_definitions definition
  join public.financial_reconciliation_automatic_rule_configs config
    on config.rule_key = definition.rule_key
   and config.rule_version = definition.version
  join public.financial_reconciliation_source_rules source_rule
    on source_rule.base_source_type = definition.base_source_type
   and source_rule.matching_source_type = 'import_cgd_extrato_ordem'
   and source_rule.operator = '+'
  where definition.rule_key =
      'fdm_bank_transfer_cgd_bank_statement_combination'
    and definition.version = 1
$$;

-- Task 3 Bank Reservation population projection reapply security and cascade
insert into public.financial_reconciliation_automatic_runs (
  id, trigger, scope, status, actor, client_request_id,
  definition_config_snapshot, analysis_processed, analysis_total
) values (
  'b3250000-0000-0000-0000-000000000001', 'manual', 'rule', 'analyzing',
  'smoke:task3-population-contract',
  'b3250000-0000-0000-0000-000000000001',
  pg_temp.task3_bank_snapshot(3), 0, 1
);
insert into public.financial_reconciliation_automatic_bank_reservation_population (
  run_id, bank_id, ordinal, bank_date
) values (
  'b3250000-0000-0000-0000-000000000001',
  'b3250000-0000-0000-0000-000000000002', 1, date '2160-01-01'
);
create temporary table task3_population_reapply_baseline on commit drop as
select to_jsonb(population) as row_snapshot
from public.financial_reconciliation_automatic_bank_reservation_population population
where population.run_id = 'b3250000-0000-0000-0000-000000000001';

\ir ../supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql

do $$
declare
  v_role text;
  v_privilege text;
begin
  if (select row_snapshot from task3_population_reapply_baseline)
      is distinct from (
        select to_jsonb(population)
        from public.financial_reconciliation_automatic_bank_reservation_population population
        where population.run_id = 'b3250000-0000-0000-0000-000000000001'
      )
    or not (select relation.relkind = 'r'
                 and relation.relrowsecurity
                 and not relation.relforcerowsecurity
        from pg_class relation
        where relation.oid =
          'public.financial_reconciliation_automatic_bank_reservation_population'::regclass)
    or exists (
      select 1
      from pg_policy policy_row
      where policy_row.polrelid =
        'public.financial_reconciliation_automatic_bank_reservation_population'::regclass
    ) then
    raise exception 'Task 3 population projection changed on reapply or lacks RLS.';
  end if;

  if exists (
    select 1
    from information_schema.table_privileges grant_row
    where grant_row.table_schema = 'public'
      and grant_row.table_name =
        'financial_reconciliation_automatic_bank_reservation_population'
      and grant_row.grantee = 'PUBLIC'
  ) then
    raise exception 'Task 3 population projection grants table access to PUBLIC.';
  end if;

  if exists (
    select 1
    from information_schema.column_privileges grant_row
    where grant_row.table_schema = 'public'
      and grant_row.table_name =
        'financial_reconciliation_automatic_bank_reservation_population'
      and grant_row.grantee in (
        'PUBLIC','anon','authenticated','service_role'
      )
  ) then
    raise exception 'Task 3 population projection grants direct column access.';
  end if;

  foreach v_role in array array['anon','authenticated','service_role']
  loop
    foreach v_privilege in array array['SELECT','INSERT','UPDATE','DELETE']
    loop
      if has_table_privilege(
          v_role,
          'public.financial_reconciliation_automatic_bank_reservation_population',
          v_privilege
        ) then
        raise exception 'Task 3 population projection grants % to %.',
          v_privilege, v_role;
      end if;
    end loop;
  end loop;

  delete from public.financial_reconciliation_automatic_runs run
  where run.id = 'b3250000-0000-0000-0000-000000000001';
  if exists (
    select 1
    from public.financial_reconciliation_automatic_bank_reservation_population population
    where population.run_id = 'b3250000-0000-0000-0000-000000000001'
  ) then
    raise exception 'Task 3 population projection did not cascade with its run.';
  end if;
end
$$;

-- Task 3 Bank Reservation continuation pages and retry idempotency
insert into public.financial_reconciliation_automatic_runs (
  id, trigger, scope, status, actor, client_request_id,
  definition_config_snapshot, analysis_processed, analysis_total
) values (
  'b3300000-0000-0000-0000-000000000001', 'manual', 'rule', 'analyzing',
  'smoke:task3-owner', 'b3300000-0000-0000-0000-000000000001',
  pg_temp.task3_bank_snapshot(3), 0, 0
);

do $$
declare
  v_first jsonb;
  v_second jsonb;
  v_retry jsonb;
begin
  v_first := public.continue_financial_reconciliation_automatic_analysis(
    'b3300000-0000-0000-0000-000000000001', 'smoke:task3-owner'
  );
  if jsonb_typeof(v_first) is distinct from 'object'
    or (v_first ?& array[
      'status','analysisComplete','analysisProcessed','analysisTotal','counts'
    ]) is not true
    or v_first->'status' is distinct from '"analyzing"'::jsonb
    or v_first->'analysisComplete' is distinct from 'false'::jsonb
    or v_first->'analysisProcessed' is distinct from '25'::jsonb
    or v_first->'analysisTotal' is distinct from '41'::jsonb
    or (select count(*)
        from public.financial_reconciliation_automatic_bank_reservation_population population
        where population.run_id = 'b3300000-0000-0000-0000-000000000001')
      is distinct from 41
    or (select min(population.ordinal)
        from public.financial_reconciliation_automatic_bank_reservation_population population
        where population.run_id = 'b3300000-0000-0000-0000-000000000001')
      is distinct from 1
    or (select max(population.ordinal)
        from public.financial_reconciliation_automatic_bank_reservation_population population
        where population.run_id = 'b3300000-0000-0000-0000-000000000001')
      is distinct from 41 then
    raise exception 'Task 3 first 25-Bank continuation page is invalid: %.',
      v_first;
  end if;

  v_second := public.continue_financial_reconciliation_automatic_analysis(
    'b3300000-0000-0000-0000-000000000001', 'smoke:task3-owner'
  );
  if jsonb_typeof(v_second) is distinct from 'object'
    or (v_second ?& array[
      'status','analysisComplete','analysisProcessed','analysisTotal','counts'
    ]) is not true
    or v_second->'status' is distinct from '"ready"'::jsonb
    or v_second->'analysisComplete' is distinct from 'true'::jsonb
    or v_second->'analysisProcessed' is distinct from '41'::jsonb
    or v_second->'analysisTotal' is distinct from '41'::jsonb
    or v_second#>'{counts,bases}' is distinct from '41'::jsonb then
    raise exception 'Task 3 final continuation page did not finalize review work: %.',
      v_second;
  end if;

  v_retry := public.continue_financial_reconciliation_automatic_analysis(
    'b3300000-0000-0000-0000-000000000001', 'smoke:task3-owner'
  );
  if v_retry is distinct from v_second then
    raise exception 'Task 3 completed continuation retry changed the immutable run: %, %.',
      v_second, v_retry;
  end if;
end
$$;

-- Task 3 Bank Reservation memberships and omitted no-match accounting
do $$
declare
  v_case record;
  v_proposal public.financial_reconciliation_automatic_proposals%rowtype;
  v_source_ids uuid[];
  v_source_ordinals integer[];
begin
  for v_case in
    select * from (values
      ('b3100000-0000-0000-0000-000000000001'::uuid, 1),
      ('b3100000-0000-0000-0000-000000000002'::uuid, 2),
      ('b3100000-0000-0000-0000-000000000003'::uuid, 10),
      ('b3100000-0000-0000-0000-000000000004'::uuid, 1)
    ) fixture(bank_id, source_count)
  loop
    select proposal.* into strict v_proposal
    from public.financial_reconciliation_automatic_proposals proposal
    join public.financial_reconciliation_automatic_proposal_memberships bank_member
      on bank_member.proposal_id = proposal.id
     and bank_member.role = 'destination'
     and bank_member.source_type = 'import_cgd_extrato_ordem'
     and bank_member.source_id = v_case.bank_id
    where proposal.run_id = 'b3300000-0000-0000-0000-000000000001';

    select array_agg(member.source_id order by member.ordinal),
           array_agg(member.ordinal order by member.ordinal)
    into v_source_ids, v_source_ordinals
    from public.financial_reconciliation_automatic_proposal_memberships member
    where member.proposal_id = v_proposal.id and member.role = 'source';

    if v_proposal.status is distinct from 'proposed'
      or v_proposal.reason is distinct from 'unique_qualifying_combination'
      or cardinality(v_source_ids) is distinct from v_case.source_count
      or v_source_ordinals is distinct from
        array(select generate_series(1, v_case.source_count))
      or v_proposal.base_source_id is distinct from v_source_ids[1]
      or (select count(*)
          from public.financial_reconciliation_automatic_proposal_memberships member
          where member.proposal_id = v_proposal.id
            and member.role = 'destination'
            and member.ordinal = 1) is distinct from 1
      or exists (
        select 1
        from public.financial_reconciliation_automatic_proposal_memberships member
        left join public.import_fdm_accounts fdm
          on member.source_type = 'import_fdm_accounts'
         and member.source_id = fdm.id
        left join public.import_cgd_extrato_ordem bank
          on member.source_type = 'import_cgd_extrato_ordem'
         and member.source_id = bank.id
        where member.proposal_id = v_proposal.id
          and (
            (member.role = 'source'
              and (member.source_type is distinct from 'import_fdm_accounts'
                or member.row_snapshot is distinct from to_jsonb(fdm)))
            or
            (member.role = 'destination'
              and (member.source_type is distinct from 'import_cgd_extrato_ordem'
                or member.row_snapshot is distinct from to_jsonb(bank)))
          )
      ) then
      raise exception 'Task 3 immutable membership/base contract failed for bank %.',
        v_case.bank_id;
    end if;
  end loop;

  if exists (
      select 1
      from public.financial_reconciliation_automatic_proposals proposal
      join public.financial_reconciliation_automatic_proposal_memberships member
        on member.proposal_id = proposal.id
      where proposal.run_id = 'b3300000-0000-0000-0000-000000000001'
        and member.source_id = 'b3100000-0000-0000-0000-000000000005'
    )
    or exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = 'b3300000-0000-0000-0000-000000000001'
        and proposal.status = 'skipped'
    )
    or (select run.counts->'skipped'
        from public.financial_reconciliation_automatic_runs run
        where run.id = 'b3300000-0000-0000-0000-000000000001')
      is distinct from '1'::jsonb then
    raise exception 'Task 3 no-match anchor was visible or missing from skipped accounting.';
  end if;

  if (exists (
      select 1
      from public.financial_reconciliation_automatic_proposals proposal
      join public.financial_reconciliation_automatic_proposal_memberships member
        on member.proposal_id = proposal.id
      where proposal.run_id = 'b3300000-0000-0000-0000-000000000001'
        and member.source_id = 'b3100000-0000-0000-0000-000000000006'
        and proposal.status = 'ambiguous'
        and proposal.reason = 'multiple_qualifying_combinations'
        and jsonb_array_length(proposal.candidate_groups) = 2
    )) is not true
    or (select count(*)
          from public.financial_reconciliation_automatic_proposals proposal
          where proposal.run_id = 'b3300000-0000-0000-0000-000000000001'
            and proposal.status = 'ambiguous'
            and proposal.reason = 'candidate_limit'
            and jsonb_array_length(proposal.candidate_groups) <= 12)
      is distinct from 3
    or (select count(*)
        from public.financial_reconciliation_automatic_proposals proposal
        join public.financial_reconciliation_automatic_proposal_memberships member
          on member.proposal_id = proposal.id
        where proposal.run_id = 'b3300000-0000-0000-0000-000000000001'
          and member.source_id in (
            'b3100000-0000-0000-0000-000000000010',
            'b3100000-0000-0000-0000-000000000011'
          )
          and proposal.status = 'ambiguous'
          and proposal.reason = 'overlapping_records') is distinct from 2 then
    raise exception 'Task 3 ambiguity, evidence bound, or shared-FDM overlap classification failed.';
  end if;

  for v_case in
    select * from (values
      ('b3100000-0000-0000-0000-000000000007'::uuid, 61, 0, 0),
      ('b3100000-0000-0000-0000-000000000009'::uuid, 19, 250000, 0),
      ('b3100000-0000-0000-0000-000000000008'::uuid, 14, 14, 12)
    ) boundary(bank_id, candidate_count, evaluated_states, evidence_count)
  loop
    select proposal.* into strict v_proposal
    from public.financial_reconciliation_automatic_proposals proposal
    join public.financial_reconciliation_automatic_proposal_memberships bank_member
      on bank_member.proposal_id = proposal.id
     and bank_member.role = 'destination'
     and bank_member.source_type = 'import_cgd_extrato_ordem'
     and bank_member.source_id = v_case.bank_id
    where proposal.run_id = 'b3300000-0000-0000-0000-000000000001';

    if v_proposal.status is distinct from 'ambiguous'
      or v_proposal.reason is distinct from 'candidate_limit'
      or jsonb_typeof(v_proposal.candidate_groups) is distinct from 'array'
      or jsonb_array_length(v_proposal.candidate_groups)
        is distinct from v_case.evidence_count
      or jsonb_typeof(v_proposal.summary_snapshot) is distinct from 'object'
      or v_proposal.summary_snapshot->'classification'
        is distinct from '"ambiguous"'::jsonb
      or v_proposal.summary_snapshot->'reason'
        is distinct from '"candidate_limit"'::jsonb
      or v_proposal.summary_snapshot->'candidateCount'
        is distinct from to_jsonb(v_case.candidate_count)
      or v_proposal.summary_snapshot->'evaluatedStates'
        is distinct from to_jsonb(v_case.evaluated_states) then
      raise exception 'Task 3 persisted hard boundary is not exact for Bank %: %.',
        v_case.bank_id, to_jsonb(v_proposal);
    end if;
  end loop;
end
$$;

-- A run with only an exhausted no-match anchor completes without a visible row.
do $$
declare
  v_reconciliation_id uuid;
begin
  insert into public.financial_reconciliations (
    status, base_source_type, matching_source_types, created_by
  ) values (
    'started', 'import_fdm_accounts', '["import_cgd_extrato_ordem"]'::jsonb,
    'smoke:task3-post-analysis-isolation'
  ) returning id into v_reconciliation_id;
  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  )
  select v_reconciliation_id, 'import_cgd_extrato_ordem', bank.id,
         bank.montante, 'smoke:task3-post-analysis-isolation'
  from public.import_cgd_extrato_ordem bank
  where bank.import_batch = 'smoke-task3-bank'
    and bank.montante is not null
    and not exists (
      select 1 from public.financial_reconciliation_items locked
      where locked.source_type = 'import_cgd_extrato_ordem'
        and locked.source_id = bank.id
    );
  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  )
  select v_reconciliation_id, 'import_fdm_accounts', fdm.id,
         fdm.amount, 'smoke:task3-post-analysis-isolation'
  from public.import_fdm_accounts fdm
  where fdm.import_batch = 'smoke-task3-fdm'
    and fdm.amount is not null
    and not exists (
      select 1 from public.financial_reconciliation_items locked
      where locked.source_type = 'import_fdm_accounts'
        and locked.source_id = fdm.id
    );
end
$$;

insert into public.import_cgd_extrato_ordem (
  id, import_batch, row_key, data, descritivo, montante
) values (
  'b3400000-0000-0000-0000-000000000001', 'smoke-task3-zero-review',
  'task3-zero-review', date '2180-01-10', 'zero review rows', -11.00
);
insert into public.import_fdm_accounts (
  id, import_batch, source_row_number, account, date_time_raw, event_date,
  category, amount, description
)
select
  ('c3400000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
  'smoke-task3-zero-review', series, 'Bank Transfer', '2180-01-10',
  date '2180-01-10', 'Reservation', 1.00, 'zero review ' || series
from generate_series(1, 11) series;
insert into public.financial_reconciliation_automatic_runs (
  id, trigger, scope, status, actor, client_request_id,
  definition_config_snapshot, analysis_processed, analysis_total
) values (
  'b3400000-0000-0000-0000-000000000010', 'manual', 'rule', 'analyzing',
  'smoke:task3-zero-review', 'b3400000-0000-0000-0000-000000000010',
  pg_temp.task3_bank_snapshot(3), 0, 0
);

do $$
declare
  v_result jsonb;
begin
  v_result := public.continue_financial_reconciliation_automatic_analysis(
    'b3400000-0000-0000-0000-000000000010', 'smoke:task3-zero-review'
  );
  if jsonb_typeof(v_result) is distinct from 'object'
    or (v_result ?& array[
      'status','analysisComplete','analysisProcessed','analysisTotal',
      'counts','proposals'
    ]) is not true
    or v_result->'status' is distinct from '"completed"'::jsonb
    or v_result->'analysisComplete' is distinct from 'true'::jsonb
    or v_result->'analysisProcessed' is distinct from '1'::jsonb
    or v_result->'analysisTotal' is distinct from '1'::jsonb
    or v_result#>'{counts,skipped}' is distinct from '1'::jsonb
    or jsonb_typeof(v_result->'proposals') is distinct from 'array'
    or jsonb_array_length(v_result->'proposals') is distinct from 0 then
    raise exception 'Task 3 zero-review run did not complete with skipped accounting: %.',
      v_result;
  end if;
end
$$;

-- Task 3 Bank Reservation run population excludes behind and ahead inter-page inserts.
do $$
declare
  v_reconciliation_id uuid;
begin
  insert into public.financial_reconciliations (
    status, base_source_type, matching_source_types, created_by
  ) values (
    'started', 'import_fdm_accounts', '["import_cgd_extrato_ordem"]'::jsonb,
    'smoke:task3-zero-population-isolation'
  ) returning id into v_reconciliation_id;
  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  )
  select v_reconciliation_id, 'import_cgd_extrato_ordem', bank.id,
         bank.montante, 'smoke:task3-zero-population-isolation'
  from public.import_cgd_extrato_ordem bank
  where bank.import_batch = 'smoke-task3-zero-review';
end
$$;

insert into public.import_cgd_extrato_ordem (
  id, import_batch, row_key, data, descritivo, montante
)
select
  ('b3500000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
  'smoke-task3-population-insert', 'task3-population-insert-' || series,
  date '2230-01-01' + (series - 1), 'population anchor ' || series, -17.00
from generate_series(1, 26) series;
insert into public.financial_reconciliation_automatic_runs (
  id, trigger, scope, status, actor, client_request_id,
  definition_config_snapshot, analysis_processed, analysis_total
) values (
  'b3500000-0000-0000-0000-000000000100', 'manual', 'rule', 'analyzing',
  'smoke:task3-population-insert',
  'b3500000-0000-0000-0000-000000000100',
  pg_temp.task3_bank_snapshot(3), 0, 0
);

do $$
declare
  v_first jsonb;
begin
  v_first := public.continue_financial_reconciliation_automatic_analysis(
    'b3500000-0000-0000-0000-000000000100',
    'smoke:task3-population-insert'
  );
  if jsonb_typeof(v_first) is distinct from 'object'
    or v_first->'status' is distinct from '"analyzing"'::jsonb
    or v_first->'analysisProcessed' is distinct from '25'::jsonb
    or v_first->'analysisTotal' is distinct from '26'::jsonb
    or (select count(*)
        from public.financial_reconciliation_automatic_bank_reservation_population population
        where population.run_id = 'b3500000-0000-0000-0000-000000000100')
      is distinct from 26
    or exists (
      select 1
      from (
        select population.ordinal,
               row_number() over (
                 order by population.bank_date, population.bank_id
               )::integer as expected_ordinal
        from public.financial_reconciliation_automatic_bank_reservation_population population
        where population.run_id = 'b3500000-0000-0000-0000-000000000100'
      ) ordered
      where ordered.ordinal is distinct from ordered.expected_ordinal
    ) then
    raise exception 'Task 3 population projection first page is invalid: %.',
      v_first;
  end if;
end
$$;

insert into public.import_cgd_extrato_ordem (
  id, import_batch, row_key, data, descritivo, montante, created_at
) values
  (
    'b3500000-0000-0000-0000-999999999999',
    'smoke-task3-population-insert', 'task3-population-insert-behind-cursor',
    date '2230-01-05', 'post-snapshot anchor behind cursor', -17.00,
    timestamptz '2000-01-01 00:00:00+00'
  ),
  (
    'b3500000-0000-0000-0000-888888888888',
    'smoke-task3-population-insert', 'task3-population-insert-ahead-cursor',
    date '2300-01-05', 'post-snapshot anchor ahead of cursor', -17.00,
    timestamptz '2000-01-01 00:00:00+00'
  );

do $$
declare
  v_second jsonb;
begin
  v_second := public.continue_financial_reconciliation_automatic_analysis(
    'b3500000-0000-0000-0000-000000000100',
    'smoke:task3-population-insert'
  );
  if jsonb_typeof(v_second) is distinct from 'object'
    or v_second->'status' is distinct from '"completed"'::jsonb
    or v_second->'analysisComplete' is distinct from 'true'::jsonb
    or v_second->'analysisProcessed' is distinct from '26'::jsonb
    or v_second->'analysisTotal' is distinct from '26'::jsonb
    or v_second#>'{counts,bases}' is distinct from '26'::jsonb
    or v_second#>'{counts,skipped}' is distinct from '26'::jsonb
    or jsonb_typeof(v_second->'proposals') is distinct from 'array'
    or jsonb_array_length(v_second->'proposals') is distinct from 0
    or public.financial_reconciliation_automatic_bank_reservation_count()
      is distinct from 28
    or exists (
      select 1
      from public.financial_reconciliation_automatic_bank_reservation_population population
      where population.run_id = 'b3500000-0000-0000-0000-000000000100'
        and population.bank_id in (
          'b3500000-0000-0000-0000-999999999999',
          'b3500000-0000-0000-0000-888888888888'
        )
    ) then
    raise exception 'Task 3 post-snapshot Banks did not wait for the next run: %.',
      v_second;
  end if;
end
$$;

do $$
declare
  v_reconciliation_id uuid;
begin
  insert into public.financial_reconciliations (
    status, base_source_type, matching_source_types, created_by
  ) values (
    'started', 'import_fdm_accounts', '["import_cgd_extrato_ordem"]'::jsonb,
    'smoke:task3-population-insert-isolation'
  ) returning id into v_reconciliation_id;
  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  )
  select v_reconciliation_id, 'import_cgd_extrato_ordem', bank.id,
         bank.montante, 'smoke:task3-population-insert-isolation'
  from public.import_cgd_extrato_ordem bank
  where bank.import_batch = 'smoke-task3-population-insert';
end
$$;

-- Task 3 Bank Reservation run population fails closed on consumed and deleted members.
insert into public.import_cgd_extrato_ordem (
  id, import_batch, row_key, data, descritivo, montante
)
select
  ('b3600000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
  'smoke-task3-population-consumed', 'task3-population-consumed-' || series,
  date '2240-01-01' + (series - 1), 'consumed anchor ' || series, -19.00
from generate_series(1, 26) series;
insert into public.import_fdm_accounts (
  id, import_batch, source_row_number, account, date_time_raw, event_date,
  category, amount, description
) values (
  'c3600000-0000-0000-0000-000000000001',
  'smoke-task3-population-consumed', 1, 'Bank Transfer', '2240-01-01',
  date '2240-01-01', 'Reservation', 19.00,
  'proposal made stale by population mutation'
);
insert into public.financial_reconciliation_automatic_runs (
  id, trigger, scope, status, actor, client_request_id,
  definition_config_snapshot, analysis_processed, analysis_total
) values (
  'b3600000-0000-0000-0000-000000000100', 'manual', 'rule', 'analyzing',
  'smoke:task3-population-consumed',
  'b3600000-0000-0000-0000-000000000100',
  pg_temp.task3_bank_snapshot(0), 0, 0
);

do $$
declare
  v_first jsonb;
  v_second jsonb;
  v_reconciliation_id uuid;
begin
  v_first := public.continue_financial_reconciliation_automatic_analysis(
    'b3600000-0000-0000-0000-000000000100',
    'smoke:task3-population-consumed'
  );
  if jsonb_typeof(v_first) is distinct from 'object'
    or v_first->'status' is distinct from '"analyzing"'::jsonb
    or v_first->'analysisProcessed' is distinct from '25'::jsonb
    or v_first->'analysisTotal' is distinct from '26'::jsonb then
    raise exception 'Task 3 consumed-population first page is invalid: %.',
      v_first;
  end if;

  insert into public.financial_reconciliations (
    status, base_source_type, matching_source_types, created_by
  ) values (
    'started', 'import_fdm_accounts', '["import_cgd_extrato_ordem"]'::jsonb,
    'smoke:task3-population-consumed'
  ) returning id into v_reconciliation_id;
  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  ) values (
    v_reconciliation_id, 'import_cgd_extrato_ordem',
    'b3600000-0000-0000-0000-000000000026', -19.00,
    'smoke:task3-population-consumed'
  );

  v_second := public.continue_financial_reconciliation_automatic_analysis(
    'b3600000-0000-0000-0000-000000000100',
    'smoke:task3-population-consumed'
  );
  if jsonb_typeof(v_second) is distinct from 'object'
    or v_second->'status' is distinct from '"failed"'::jsonb
    or v_second->'analysisComplete' is distinct from 'true'::jsonb
    or v_second->'analysisProcessed' is distinct from '25'::jsonb
    or v_second->'analysisTotal' is distinct from '26'::jsonb
    or v_second->'analysisErrorCode' is distinct from
      '"analysis_population_changed"'::jsonb
    or jsonb_typeof(v_second->'analysisCompletedAt') is distinct from 'string'
    or jsonb_typeof(v_second->'finishedAt') is distinct from 'string'
    or (select count(*)
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = 'b3600000-0000-0000-0000-000000000100'
          and proposal.status = 'stale'
          and proposal.reason = 'analysis_population_changed')
      is distinct from 1
    or exists (
      select 1
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = 'b3600000-0000-0000-0000-000000000100'
        and proposal.status in ('proposed','executing')
    ) then
    raise exception 'Task 3 consumed population did not fail closed coherently: %.',
      v_second;
  end if;
end
$$;

do $$
declare
  v_reconciliation_id uuid;
begin
  insert into public.financial_reconciliations (
    status, base_source_type, matching_source_types, created_by
  ) values (
    'started', 'import_fdm_accounts', '["import_cgd_extrato_ordem"]'::jsonb,
    'smoke:task3-population-consumed-isolation'
  ) returning id into v_reconciliation_id;
  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  )
  select v_reconciliation_id, 'import_cgd_extrato_ordem', bank.id,
         bank.montante, 'smoke:task3-population-consumed-isolation'
  from public.import_cgd_extrato_ordem bank
  where bank.import_batch = 'smoke-task3-population-consumed'
    and not exists (
      select 1
      from public.financial_reconciliation_items locked
      where locked.source_type = 'import_cgd_extrato_ordem'
        and locked.source_id = bank.id
    );
end
$$;

insert into public.import_cgd_extrato_ordem (
  id, import_batch, row_key, data, descritivo, montante
)
select
  ('b3700000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
  'smoke-task3-population-deleted', 'task3-population-deleted-' || series,
  date '2250-01-01' + (series - 1), 'deleted anchor ' || series, -23.00
from generate_series(1, 26) series;
insert into public.financial_reconciliation_automatic_runs (
  id, trigger, scope, status, actor, client_request_id,
  definition_config_snapshot, analysis_processed, analysis_total
) values (
  'b3700000-0000-0000-0000-000000000100', 'manual', 'rule', 'analyzing',
  'smoke:task3-population-deleted',
  'b3700000-0000-0000-0000-000000000100',
  pg_temp.task3_bank_snapshot(3), 0, 0
);

do $$
declare
  v_first jsonb;
  v_second jsonb;
begin
  v_first := public.continue_financial_reconciliation_automatic_analysis(
    'b3700000-0000-0000-0000-000000000100',
    'smoke:task3-population-deleted'
  );
  if jsonb_typeof(v_first) is distinct from 'object'
    or v_first->'status' is distinct from '"analyzing"'::jsonb
    or v_first->'analysisProcessed' is distinct from '25'::jsonb
    or v_first->'analysisTotal' is distinct from '26'::jsonb then
    raise exception 'Task 3 deleted-population first page is invalid: %.',
      v_first;
  end if;

  delete from public.import_cgd_extrato_ordem bank
  where bank.id = 'b3700000-0000-0000-0000-000000000026';

  v_second := public.continue_financial_reconciliation_automatic_analysis(
    'b3700000-0000-0000-0000-000000000100',
    'smoke:task3-population-deleted'
  );
  if jsonb_typeof(v_second) is distinct from 'object'
    or v_second->'status' is distinct from '"failed"'::jsonb
    or v_second->'analysisComplete' is distinct from 'true'::jsonb
    or v_second->'analysisProcessed' is distinct from '25'::jsonb
    or v_second->'analysisTotal' is distinct from '26'::jsonb
    or v_second->'analysisErrorCode' is distinct from
      '"analysis_population_changed"'::jsonb
    or jsonb_typeof(v_second->'analysisCompletedAt') is distinct from 'string'
    or jsonb_typeof(v_second->'finishedAt') is distinct from 'string' then
    raise exception 'Task 3 deleted population did not fail closed coherently: %.',
      v_second;
  end if;
end
$$;

do $$
declare
  v_reconciliation_id uuid;
begin
  insert into public.financial_reconciliations (
    status, base_source_type, matching_source_types, created_by
  ) values (
    'started', 'import_fdm_accounts', '["import_cgd_extrato_ordem"]'::jsonb,
    'smoke:task3-population-deleted-isolation'
  ) returning id into v_reconciliation_id;
  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  )
  select v_reconciliation_id, 'import_cgd_extrato_ordem', bank.id,
         bank.montante, 'smoke:task3-population-deleted-isolation'
  from public.import_cgd_extrato_ordem bank
  where bank.import_batch = 'smoke-task3-population-deleted';
end
$$;

-- Task 3 Bank Reservation authoritative shared-bank overlap
insert into public.import_cgd_extrato_ordem (
  id, import_batch, row_key, data, descritivo, montante
) values (
  'b3200000-0000-0000-0000-000000000001', 'smoke-task3-shared-bank',
  'task3-shared-bank', date '2170-01-10', 'shared bank', -20.00
);
insert into public.import_fdm_accounts (
  id, import_batch, account, date_time_raw, event_date, category,
  amount, description
) values
  ('c3200000-0000-0000-0000-000000000001', 'smoke-task3-shared-bank',
   'Bank Transfer', '2170-01-10', date '2170-01-10', 'Reservation',
   20.00, 'shared bank source a'),
  ('c3200000-0000-0000-0000-000000000002', 'smoke-task3-shared-bank',
   'Bank Transfer', '2170-01-10', date '2170-01-10', 'Reservation',
   20.00, 'shared bank source b');

insert into public.financial_reconciliation_automatic_runs (
  id, trigger, scope, status, actor, client_request_id,
  definition_config_snapshot, analysis_cursor_date, analysis_cursor_id,
  analysis_processed, analysis_total
) values (
  'b3200000-0000-0000-0000-000000000010', 'manual', 'rule', 'analyzing',
  'smoke:task3-shared-bank', 'b3200000-0000-0000-0000-000000000010',
  pg_temp.task3_bank_snapshot(3), date '2170-01-10',
  'b3200000-0000-0000-0000-000000000001', 1, 1
);
insert into public.financial_reconciliation_automatic_bank_reservation_population (
  run_id, bank_id, ordinal, bank_date
) values (
  'b3200000-0000-0000-0000-000000000010',
  'b3200000-0000-0000-0000-000000000001', 1, date '2170-01-10'
);
insert into public.financial_reconciliation_automatic_proposals (
  id, run_id, rule_key, rule_version, base_source_type, base_source_id,
  base_source_date, base_snapshot, allowed_difference, status, reason,
  signature, grouping_key, summary_snapshot
) values
  ('b3200000-0000-0000-0000-000000000011',
   'b3200000-0000-0000-0000-000000000010',
   'fdm_bank_transfer_cgd_bank_statement_combination', 1,
   'import_fdm_accounts', 'c3200000-0000-0000-0000-000000000001',
   date '2170-01-10', '{}'::jsonb, 0, 'proposed',
   'unique_qualifying_combination', 'task3-shared-bank-a',
   'b3200000-0000-0000-0000-000000000001', '{}'::jsonb),
  ('b3200000-0000-0000-0000-000000000012',
   'b3200000-0000-0000-0000-000000000010',
   'fdm_bank_transfer_cgd_bank_statement_combination', 1,
   'import_fdm_accounts', 'c3200000-0000-0000-0000-000000000002',
   date '2170-01-10', '{}'::jsonb, 0, 'proposed',
   'unique_qualifying_combination', 'task3-shared-bank-b',
   'b3200000-0000-0000-0000-000000000001', '{}'::jsonb);
insert into public.financial_reconciliation_automatic_proposal_memberships (
  proposal_id, role, source_type, source_id, ordinal, source_date,
  amount, description, account, row_snapshot
)
select proposal.id, 'source', 'import_fdm_accounts', proposal.base_source_id,
       1, fdm.event_date, fdm.amount, fdm.description, fdm.account, to_jsonb(fdm)
from public.financial_reconciliation_automatic_proposals proposal
join public.import_fdm_accounts fdm on fdm.id = proposal.base_source_id
where proposal.id in (
  'b3200000-0000-0000-0000-000000000011',
  'b3200000-0000-0000-0000-000000000012'
);
insert into public.financial_reconciliation_automatic_proposal_memberships (
  proposal_id, role, source_type, source_id, ordinal, source_date,
  amount, description, account, row_snapshot
)
select proposal.id, 'destination', 'import_cgd_extrato_ordem', bank.id,
       1, bank.data, bank.montante, bank.descritivo, '', to_jsonb(bank)
from public.financial_reconciliation_automatic_proposals proposal
cross join public.import_cgd_extrato_ordem bank
where proposal.id in (
    'b3200000-0000-0000-0000-000000000011',
    'b3200000-0000-0000-0000-000000000012'
  )
  and bank.id = 'b3200000-0000-0000-0000-000000000001';

do $$
declare
  v_result jsonb;
begin
  v_result := public.continue_financial_reconciliation_automatic_analysis(
    'b3200000-0000-0000-0000-000000000010',
    'smoke:task3-shared-bank'
  );
  if jsonb_typeof(v_result) is distinct from 'object'
    or v_result->'status' is distinct from '"ready"'::jsonb
    or v_result->'analysisComplete' is distinct from 'true'::jsonb
    or v_result->'analysisProcessed' is distinct from '1'::jsonb
    or v_result->'analysisTotal' is distinct from '1'::jsonb
    or (select count(*)
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = 'b3200000-0000-0000-0000-000000000010'
          and proposal.status = 'ambiguous'
          and proposal.reason = 'overlapping_records') is distinct from 2 then
    raise exception 'Task 3 shared Bank overlap did not update every affected proposal: %.',
      v_result;
  end if;
end
$$;

-- Task 4 Adyen month union, eligibility, and cursor
-- Isolate Adyen from any earlier smoke or target-database rows while retaining
-- every non-Adyen fixture for the existing managed strategies.
do $$
declare
  v_reconciliation_id uuid;
begin
  insert into public.financial_reconciliations (
    status, base_source_type, matching_source_types, created_by
  ) values (
    'started', 'import_cgd_extrato_ordem', '["import_fdm_accounts"]'::jsonb,
    'smoke:task4-isolation'
  ) returning id into v_reconciliation_id;

  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  )
  select v_reconciliation_id, 'import_cgd_extrato_ordem', bank.id,
         bank.montante, 'smoke:task4-isolation'
  from public.import_cgd_extrato_ordem bank
  where bank.montante is not null
    and bank.descritivo ilike '%Adyen%'
    and not exists (
      select 1 from public.financial_reconciliation_items locked
      where locked.source_type = 'import_cgd_extrato_ordem'
        and locked.source_id = bank.id
    );

  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  )
  select v_reconciliation_id, 'import_fdm_accounts', fdm.id,
         fdm.amount, 'smoke:task4-isolation'
  from public.import_fdm_accounts fdm
  where fdm.amount is not null
    and fdm.account = 'Adyen'
    and not exists (
      select 1 from public.financial_reconciliation_items locked
      where locked.source_type = 'import_fdm_accounts'
        and locked.source_id = fdm.id
    );
end
$$;

insert into public.import_cgd_extrato_ordem (
  id, import_batch, row_key, data, descritivo, montante
)
select
  ('d4010000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
  'smoke-task4-adyen', 'task4-many-bank-' || series,
  date '2026-01-01' + ((series - 1) % 28),
  case when series % 2 = 0 then 'ADYEN settlement' else 'Adyen' end,
  1.00
from generate_series(1, 30) series;

insert into public.import_cgd_extrato_ordem (
  id, import_batch, row_key, data, descritivo, montante
) values
  ('d4020000-0000-0000-0000-000000000001', 'smoke-task4-adyen',
   'task4-within-bank', date '2026-02-10', 'Adyen within', 110.00),
  ('d4020000-0000-0000-0000-000000000002', 'smoke-task4-adyen',
   'task4-nonmatching-bank', date '2026-02-11', 'payment provider', 999.00),
  ('d4030000-0000-0000-0000-000000000001', 'smoke-task4-adyen',
   'task4-boundary-bank', date '2026-03-10', 'ADYEN settlement', 120.00),
  ('d4040000-0000-0000-0000-000000000001', 'smoke-task4-adyen',
   'task4-over-bank', date '2026-04-10', 'Adyen over', 121.00),
  ('d4050000-0000-0000-0000-000000000001', 'smoke-task4-adyen',
   'task4-bank-only', date '2026-05-10', 'Adyen bank only', 50.00),
  ('d4080000-0000-0000-0000-000000000001', 'smoke-task4-adyen',
   'task4-null-date-bank', null, 'Adyen null date', 25.00),
  ('d4080000-0000-0000-0000-000000000002', 'smoke-task4-adyen',
   'task4-null-amount-bank', date '2026-02-15', 'Adyen null amount', null),
  ('d4000000-0000-0000-0000-000000000001', 'smoke-task4-adyen',
   'task4-pre-floor-bank', date '2025-12-31', 'Adyen pre-floor', 25.00),
  ('d4090000-0000-0000-0000-000000000001', 'smoke-task4-adyen',
   'task4-current-bank', date_trunc('month', current_date)::date,
   'Adyen current', 25.00);

insert into public.import_fdm_accounts (
  id, import_batch, account, date_time_raw, event_date, category,
  amount, description
)
select
  ('e4010000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
  'smoke-task4-adyen', 'Adyen',
  to_char(date '2026-01-01' + ((series - 1) % 28), 'YYYY-MM-DD'),
  date '2026-01-01' + ((series - 1) % 28), 'Reservation',
  case when series <= 30 then 1.00 else 0.00 end,
  'many FDM member ' || series
from generate_series(1, 35) series;

insert into public.import_fdm_accounts (
  id, import_batch, account, date_time_raw, event_date, category,
  amount, description
) values
  ('e4020000-0000-0000-0000-000000000001', 'smoke-task4-adyen',
   'Adyen', '2026-02-10', date '2026-02-10', 'Reservation',
   100.00, 'within FDM'),
  ('e4020000-0000-0000-0000-000000000002', 'smoke-task4-adyen',
   'adyen', '2026-02-11', date '2026-02-11', 'Reservation',
   999.00, 'lowercase near match'),
  ('e4020000-0000-0000-0000-000000000003', 'smoke-task4-adyen',
   'Adyen ', '2026-02-12', date '2026-02-12', 'Reservation',
   999.00, 'space near match'),
  ('e4030000-0000-0000-0000-000000000001', 'smoke-task4-adyen',
   'Adyen', '2026-03-10', date '2026-03-10', 'Reservation',
   100.00, 'boundary FDM'),
  ('e4040000-0000-0000-0000-000000000001', 'smoke-task4-adyen',
   'Adyen', '2026-04-10', date '2026-04-10', 'Reservation',
   100.00, 'over FDM'),
  ('e4060000-0000-0000-0000-000000000001', 'smoke-task4-adyen',
   'Adyen', '2026-06-10', date '2026-06-10', 'Reservation',
   50.00, 'FDM only'),
  ('e4080000-0000-0000-0000-000000000001', 'smoke-task4-adyen',
   'Adyen', '', null, 'Reservation', 25.00, 'null date FDM'),
  ('e4080000-0000-0000-0000-000000000002', 'smoke-task4-adyen',
   'Adyen', '2026-02-15', date '2026-02-15', 'Reservation',
   null, 'null amount FDM'),
  ('e4000000-0000-0000-0000-000000000001', 'smoke-task4-adyen',
   'Adyen', '2025-12-31', date '2025-12-31', 'Reservation',
   25.00, 'pre-floor FDM'),
  ('e4090000-0000-0000-0000-000000000001', 'smoke-task4-adyen',
   'Adyen', to_char(date_trunc('month', current_date), 'YYYY-MM-DD'),
   date_trunc('month', current_date)::date, 'Reservation',
   25.00, 'current FDM');

do $$
declare
  v_count bigint;
  v_first date[];
  v_second date[];
  v_end date[];
  v_rejected boolean := false;
begin
  select public.financial_reconciliation_automatic_adyen_month_count()
  into v_count;
  select array_agg(page.calendar_month order by page.calendar_month)
  into v_first
  from public.financial_reconciliation_automatic_adyen_month_page(null, 3) page;
  select array_agg(page.calendar_month order by page.calendar_month)
  into v_second
  from public.financial_reconciliation_automatic_adyen_month_page(
    date '2026-03-01', 3
  ) page;
  select array_agg(page.calendar_month order by page.calendar_month)
  into v_end
  from public.financial_reconciliation_automatic_adyen_month_page(
    date '2026-06-01', 3
  ) page;

  if v_count is distinct from 6
    or v_first is distinct from array[
      date '2026-01-01', date '2026-02-01', date '2026-03-01'
    ]
    or v_second is distinct from array[
      date '2026-04-01', date '2026-05-01', date '2026-06-01'
    ]
    or v_end is not null then
    raise exception 'Task 4 month union/cursor excluded or reordered a literal fixture: %, %, %, %.',
      v_count, v_first, v_second, v_end;
  end if;

  begin
    perform public.financial_reconciliation_automatic_adyen_month_page(
      null, 26
    );
  exception when others then
    v_rejected := sqlerrm =
      'Automatic Adyen month page size must be between 1 and 25.';
  end;
  if v_rejected is not true then
    raise exception 'Task 4 Adyen month page accepted an oversized limit.';
  end if;
end
$$;

create or replace function pg_temp.task4_adyen_snapshot(p_allowance numeric)
returns jsonb
language sql
stable
set search_path = public, pg_temp
as $$
  select jsonb_build_array(jsonb_build_object(
    'ruleKey', definition.rule_key,
    'ruleVersion', definition.version,
    'displayName', definition.display_name,
    'priority', config.priority,
    'differenceAllowed', p_allowance,
    'maxDifferenceDays', 31,
    'destinationSourceType', 'import_fdm_accounts',
    'definition', definition.definition,
    'operator', '-'
  ))
  from public.financial_reconciliation_automatic_rule_definitions definition
  join public.financial_reconciliation_automatic_rule_configs config
    on config.rule_key = definition.rule_key
   and config.rule_version = definition.version
  where definition.rule_key =
      'cgd_bank_statement_fdm_adyen_monthly_payments'
    and definition.version = 1
$$;

insert into public.financial_reconciliation_automatic_runs (
  id, trigger, scope, status, actor, client_request_id,
  definition_config_snapshot, analysis_processed, analysis_total
) values (
  'd4400000-0000-0000-0000-000000000001', 'manual', 'rule', 'analyzing',
  'smoke:task4-adyen', 'd4400000-0000-0000-0000-000000000001',
  pg_temp.task4_adyen_snapshot(20.00), 0, 0
);

-- Task 4 Adyen closed-month classifications and omitted empty sides
do $$
declare
  v_result jsonb;
begin
  v_result := public.continue_financial_reconciliation_automatic_analysis(
    'd4400000-0000-0000-0000-000000000001',
    'smoke:task4-adyen'
  );

  if jsonb_typeof(v_result) is distinct from 'object'
    or v_result->'status' is distinct from '"ready"'::jsonb
    or v_result->'analysisComplete' is distinct from 'true'::jsonb
    or v_result->'analysisProcessed' is distinct from '6'::jsonb
    or v_result->'analysisTotal' is distinct from '6'::jsonb
    or v_result#>'{counts,bases}' is distinct from '6'::jsonb
    or v_result#>'{counts,proposed}' is distinct from '3'::jsonb
    or v_result#>'{counts,ambiguous}' is distinct from '1'::jsonb
    or v_result#>'{counts,skipped}' is distinct from '2'::jsonb then
    raise exception 'Task 4 Adyen run lifecycle/counts are invalid: %.', v_result;
  end if;

  if (select count(*)
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = 'd4400000-0000-0000-0000-000000000001')
        is distinct from 4
    or (select count(distinct proposal.grouping_key)
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = 'd4400000-0000-0000-0000-000000000001')
          is distinct from 4
    or not exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = 'd4400000-0000-0000-0000-000000000001'
        and proposal.grouping_key = '2026-01'
        and proposal.status = 'proposed' and proposal.reason = ''
        and proposal.calculated_difference = 0
    )
    or not exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = 'd4400000-0000-0000-0000-000000000001'
        and proposal.grouping_key = '2026-02'
        and proposal.status = 'proposed' and proposal.reason = ''
        and proposal.calculated_difference = 10
    )
    or not exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = 'd4400000-0000-0000-0000-000000000001'
        and proposal.grouping_key = '2026-03'
        and proposal.status = 'proposed' and proposal.reason = ''
        and proposal.calculated_difference = 20
    )
    or not exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = 'd4400000-0000-0000-0000-000000000001'
        and proposal.grouping_key = '2026-04'
        and proposal.status = 'ambiguous'
        and proposal.reason = 'monthly_difference_exceeded'
        and proposal.calculated_difference = 21
    )
    or exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = 'd4400000-0000-0000-0000-000000000001'
        and proposal.grouping_key in ('2025-12','2026-05','2026-06')
    ) then
    raise exception 'Task 4 Adyen monthly classifications or empty-side omission are invalid.';
  end if;

  if exists (
    select 1
    from public.financial_reconciliation_automatic_proposals proposal
    where proposal.run_id = 'd4400000-0000-0000-0000-000000000001'
      and (
        proposal.rule_key is distinct from
          'cgd_bank_statement_fdm_adyen_monthly_payments'
        or proposal.rule_version is distinct from 1
        or proposal.base_source_type is distinct from 'import_cgd_extrato_ordem'
        or proposal.allowed_difference is distinct from 20
        or proposal.grouping_key is distinct from to_char(
          (proposal.summary_snapshot->>'calendarMonth')::date, 'YYYY-MM'
        )
        or proposal.summary_snapshot->>'sourceDescriptionContains'
          is distinct from 'Adyen'
        or proposal.summary_snapshot->>'destinationAccount'
          is distinct from 'Adyen'
        or proposal.summary_snapshot->>'operator' is distinct from '-'
        or proposal.summary_snapshot->'maxDifferenceDays'
          is distinct from '31'::jsonb
        or proposal.summary_snapshot->'differenceAllowed'
          is distinct from '20.00'::jsonb
        or proposal.summary_snapshot->'calculatedDifference' is distinct from
          to_jsonb(proposal.calculated_difference)
      )
  ) then
    raise exception 'Task 4 Adyen proposal snapshot contract is incomplete.';
  end if;
end
$$;

-- Task 4 Adyen memberships, retry, and reapply idempotency
do $$
declare
  v_source_ids uuid[];
  v_destination_ids uuid[];
begin
  if (select count(*)
      from public.financial_reconciliation_automatic_adyen_population population
      where population.run_id = 'd4400000-0000-0000-0000-000000000001')
      is distinct from 73
    or (select count(distinct population.calendar_month)
        from public.financial_reconciliation_automatic_adyen_population population
        where population.run_id = 'd4400000-0000-0000-0000-000000000001')
      is distinct from 6
    or exists (
      select 1
      from public.financial_reconciliation_automatic_adyen_population population
      left join public.import_cgd_extrato_ordem bank
        on population.role = 'source' and bank.id = population.source_id
      left join public.import_fdm_accounts fdm
        on population.role = 'destination' and fdm.id = population.source_id
      where population.run_id = 'd4400000-0000-0000-0000-000000000001'
        and (
          (population.role = 'source' and (
            population.source_type is distinct from 'import_cgd_extrato_ordem'
            or population.source_date is distinct from bank.data
            or population.amount is distinct from bank.montante
            or population.description is distinct from bank.descritivo
            or population.account is distinct from ''
            or population.row_snapshot is distinct from to_jsonb(bank)
          ))
          or (population.role = 'destination' and (
            population.source_type is distinct from 'import_fdm_accounts'
            or population.source_date is distinct from fdm.event_date
            or population.amount is distinct from fdm.amount
            or population.description is distinct from fdm.description
            or population.account is distinct from fdm.account
            or population.row_snapshot is distinct from to_jsonb(fdm)
          ))
        )
    )
    or exists (
      select 1
      from public.financial_reconciliation_automatic_adyen_population population
      where population.run_id = 'd4400000-0000-0000-0000-000000000001'
      group by population.calendar_month, population.role
      having min(population.member_ordinal) is distinct from 1
        or max(population.member_ordinal) is distinct from count(*)
    ) then
    raise exception 'Task 4 Adyen projected identities, ordinals, or audit snapshots diverged.';
  end if;

  select array_agg(member.source_id order by member.source_id)
  into v_source_ids
  from public.financial_reconciliation_automatic_proposal_memberships member
  join public.financial_reconciliation_automatic_proposals proposal
    on proposal.id = member.proposal_id
  where proposal.run_id = 'd4400000-0000-0000-0000-000000000001'
    and member.role = 'source';

  select array_agg(member.source_id order by member.source_id)
  into v_destination_ids
  from public.financial_reconciliation_automatic_proposal_memberships member
  join public.financial_reconciliation_automatic_proposals proposal
    on proposal.id = member.proposal_id
  where proposal.run_id = 'd4400000-0000-0000-0000-000000000001'
    and member.role = 'destination';

  if v_source_ids is distinct from array(
      select expected.id
      from (
        select ('d4010000-0000-0000-0000-' ||
          lpad(series::text, 12, '0'))::uuid as id
        from generate_series(1, 30) series
        union all values
          ('d4020000-0000-0000-0000-000000000001'::uuid),
          ('d4030000-0000-0000-0000-000000000001'::uuid),
          ('d4040000-0000-0000-0000-000000000001'::uuid)
      ) expected order by expected.id
    )
    or v_destination_ids is distinct from array(
      select expected.id
      from (
        select ('e4010000-0000-0000-0000-' ||
          lpad(series::text, 12, '0'))::uuid as id
        from generate_series(1, 35) series
        union all values
          ('e4020000-0000-0000-0000-000000000001'::uuid),
          ('e4030000-0000-0000-0000-000000000001'::uuid),
          ('e4040000-0000-0000-0000-000000000001'::uuid)
      ) expected order by expected.id
    ) then
    raise exception 'Task 4 Adyen memberships skipped or included the wrong literal rows.';
  end if;

  if exists (
    select 1
    from (
      select member.*,
             row_number() over (
               partition by member.proposal_id, member.role
               order by member.source_date, member.source_id
             )::integer as expected_ordinal
      from public.financial_reconciliation_automatic_proposal_memberships member
      join public.financial_reconciliation_automatic_proposals proposal
        on proposal.id = member.proposal_id
      where proposal.run_id = 'd4400000-0000-0000-0000-000000000001'
    ) member
    where member.ordinal is distinct from member.expected_ordinal
      or (member.role = 'source'
          and member.source_type is distinct from 'import_cgd_extrato_ordem')
      or (member.role = 'destination'
          and member.source_type is distinct from 'import_fdm_accounts')
  )
    or (select count(*)
        from public.financial_reconciliation_automatic_proposal_memberships member
        join public.financial_reconciliation_automatic_proposals proposal
          on proposal.id = member.proposal_id
        where proposal.run_id = 'd4400000-0000-0000-0000-000000000001'
          and proposal.grouping_key = '2026-01'
          and member.role = 'source') is distinct from 30
    or (select count(*)
        from public.financial_reconciliation_automatic_proposal_memberships member
        join public.financial_reconciliation_automatic_proposals proposal
          on proposal.id = member.proposal_id
        where proposal.run_id = 'd4400000-0000-0000-0000-000000000001'
          and proposal.grouping_key = '2026-01'
          and member.role = 'destination') is distinct from 35 then
    raise exception 'Task 4 Adyen membership roles, ordinals, or large month are invalid.';
  end if;

  if exists (
    select 1
    from public.financial_reconciliation_automatic_proposals proposal
    left join public.financial_reconciliation_automatic_proposal_memberships member
      on member.proposal_id = proposal.id
     and member.role = 'source'
     and member.ordinal = 1
    where proposal.run_id = 'd4400000-0000-0000-0000-000000000001'
      and (
        member.source_id is null
        or proposal.base_source_id is distinct from member.source_id
        or proposal.base_source_date is distinct from member.source_date
        or proposal.base_snapshot is distinct from member.row_snapshot
      )
  ) then
    raise exception 'Task 4 Adyen canonical Bank is not the first source member.';
  end if;
end
$$;

-- A run whose month union contains only one-sided months completes with no
-- review proposal while preserving both processed months as skipped accounting.
do $$
declare
  v_reconciliation_id uuid;
begin
  insert into public.financial_reconciliations (
    status, base_source_type, matching_source_types, created_by
  ) values (
    'started', 'import_cgd_extrato_ordem', '["import_fdm_accounts"]'::jsonb,
    'smoke:task4-empty-only-isolation'
  ) returning id into v_reconciliation_id;

  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  )
  select distinct v_reconciliation_id, member.source_type, member.source_id,
         member.amount, 'smoke:task4-empty-only-isolation'
  from public.financial_reconciliation_automatic_proposal_memberships member
  join public.financial_reconciliation_automatic_proposals proposal
    on proposal.id = member.proposal_id
  where proposal.run_id = 'd4400000-0000-0000-0000-000000000001';
end
$$;

insert into public.financial_reconciliation_automatic_runs (
  id, trigger, scope, status, actor, client_request_id,
  definition_config_snapshot, analysis_processed, analysis_total
) values (
  'd4400000-0000-0000-0000-000000000002', 'manual', 'rule', 'analyzing',
  'smoke:task4-empty-only', 'd4400000-0000-0000-0000-000000000002',
  pg_temp.task4_adyen_snapshot(20.00), 0, 0
);

do $$
declare
  v_result jsonb;
begin
  v_result := public.continue_financial_reconciliation_automatic_analysis(
    'd4400000-0000-0000-0000-000000000002',
    'smoke:task4-empty-only'
  );
  if v_result->'status' is distinct from '"completed"'::jsonb
    or v_result->'analysisComplete' is distinct from 'true'::jsonb
    or jsonb_typeof(v_result->'finishedAt') is distinct from 'string'
    or v_result->'analysisProcessed' is distinct from '2'::jsonb
    or v_result->'analysisTotal' is distinct from '2'::jsonb
    or v_result#>'{counts,bases}' is distinct from '2'::jsonb
    or v_result#>'{counts,proposed}' is distinct from '0'::jsonb
    or v_result#>'{counts,ambiguous}' is distinct from '0'::jsonb
    or v_result#>'{counts,skipped}' is distinct from '2'::jsonb
    or exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = 'd4400000-0000-0000-0000-000000000002'
    ) then
    raise exception 'Task 4 empty-side-only run did not complete invisibly: %.',
      v_result;
  end if;
end
$$;

-- Task 4 Adyen frozen population pages and inter-page inserts
-- Task 4 Adyen frozen population cursor pair and retry
-- The projection is seeded directly here so this transactional fixture remains
-- a deterministic >25-month paging test even when run before 2028.
insert into public.import_cgd_extrato_ordem (
  id, import_batch, row_key, data, descritivo, montante
)
select
  ('d4500000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
  'smoke-task4-frozen', 'task4-frozen-bank-' || series,
  (date '2026-01-10' + make_interval(months => (series - 1) * 2))::date,
  'Adyen frozen month ' || series, 10.00
from generate_series(1, 26) series;

insert into public.import_fdm_accounts (
  id, import_batch, source_row_number, account, date_time_raw, event_date,
  category, amount, description
)
select
  ('e4500000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
  'smoke-task4-frozen', series, 'Adyen',
  to_char(
    date '2026-01-10' + make_interval(months => (series - 1) * 2),
    'YYYY-MM-DD'
  ),
  (date '2026-01-10' + make_interval(months => (series - 1) * 2))::date,
  'Reservation', 10.00, 'Adyen frozen month ' || series
from generate_series(1, 26) series;

insert into public.financial_reconciliation_automatic_runs (
  id, trigger, scope, status, actor, client_request_id,
  definition_config_snapshot, analysis_processed, analysis_total
) values (
  'd4500000-0000-0000-0000-000000000100', 'manual', 'rule', 'analyzing',
  'smoke:task4-frozen', 'd4500000-0000-0000-0000-000000000100',
  pg_temp.task4_adyen_snapshot(0.00), 0, 26
);

insert into public.financial_reconciliation_automatic_adyen_population (
  run_id, calendar_month, month_ordinal, role, source_type, source_id,
  member_ordinal, source_date, amount, description, account, row_snapshot
)
select
  'd4500000-0000-0000-0000-000000000100',
  date_trunc('month', bank.data)::date, series, 'source',
  'import_cgd_extrato_ordem', bank.id, 1, bank.data, bank.montante,
  bank.descritivo, '', to_jsonb(bank)
from generate_series(1, 26) series
join public.import_cgd_extrato_ordem bank
  on bank.id = ('d4500000-0000-0000-0000-' ||
    lpad(series::text, 12, '0'))::uuid
union all
select
  'd4500000-0000-0000-0000-000000000100',
  date_trunc('month', fdm.event_date)::date, series, 'destination',
  'import_fdm_accounts', fdm.id, 1, fdm.event_date, fdm.amount,
  fdm.description, fdm.account, to_jsonb(fdm)
from generate_series(1, 26) series
join public.import_fdm_accounts fdm
  on fdm.id = ('e4500000-0000-0000-0000-' ||
    lpad(series::text, 12, '0'))::uuid;

do $$
declare
  v_first jsonb;
begin
  v_first := public.continue_financial_reconciliation_automatic_analysis(
    'd4500000-0000-0000-0000-000000000100', 'smoke:task4-frozen'
  );
  if v_first->'status' is distinct from '"analyzing"'::jsonb
    or v_first->'analysisComplete' is distinct from 'false'::jsonb
    or v_first->'analysisProcessed' is distinct from '25'::jsonb
    or v_first->'analysisTotal' is distinct from '26'::jsonb
    or (select run.analysis_cursor_date
        from public.financial_reconciliation_automatic_runs run
        where run.id = 'd4500000-0000-0000-0000-000000000100')
      is distinct from date '2030-01-01'
    or (select run.analysis_cursor_id
        from public.financial_reconciliation_automatic_runs run
        where run.id = 'd4500000-0000-0000-0000-000000000100')
      is distinct from 'd4500000-0000-0000-0000-000000000025'::uuid
    or (select count(distinct population.month_ordinal)
        from public.financial_reconciliation_automatic_adyen_population population
        where population.run_id = 'd4500000-0000-0000-0000-000000000100')
      is distinct from 26
    or (select count(*)
        from public.financial_reconciliation_automatic_adyen_population population
        where population.run_id = 'd4500000-0000-0000-0000-000000000100')
      is distinct from 52 then
    raise exception 'Task 4 frozen Adyen first page/counters are invalid: %.',
      v_first;
  end if;
end
$$;

create temporary table task4_adyen_cursor_checkpoint on commit drop as
select run.analysis_cursor_date, run.analysis_cursor_id
from public.financial_reconciliation_automatic_runs run
where run.id = 'd4500000-0000-0000-0000-000000000100';

insert into public.import_cgd_extrato_ordem (
  id, import_batch, row_key, data, descritivo, montante
) values
  ('d4510000-0000-0000-0000-000000000001', 'smoke-task4-frozen-new',
   'task4-new-month-behind', date '2026-02-10', 'Adyen new behind', 3.00),
  ('d4510000-0000-0000-0000-000000000002', 'smoke-task4-frozen-new',
   'task4-new-month-ahead', date '2030-02-10', 'Adyen new ahead', 4.00),
  ('d4510000-0000-0000-0000-000000000003', 'smoke-task4-frozen-new',
   'task4-new-member-processed', date '2026-01-20',
   'Adyen new processed member', 5.00);
insert into public.import_fdm_accounts (
  id, import_batch, source_row_number, account, date_time_raw, event_date,
  category, amount, description
) values
  ('e4510000-0000-0000-0000-000000000001', 'smoke-task4-frozen-new', 1,
   'Adyen', '2026-02-10', date '2026-02-10', 'Reservation', 3.00,
   'new month behind'),
  ('e4510000-0000-0000-0000-000000000002', 'smoke-task4-frozen-new', 2,
   'Adyen', '2030-02-10', date '2030-02-10', 'Reservation', 4.00,
   'new month ahead'),
  ('e4510000-0000-0000-0000-000000000003', 'smoke-task4-frozen-new', 3,
   'Adyen', '2026-01-20', date '2026-01-20', 'Reservation', 5.00,
   'new processed member');

do $$
declare
  v_second jsonb;
  v_retry jsonb;
begin
  v_second := public.continue_financial_reconciliation_automatic_analysis(
    'd4500000-0000-0000-0000-000000000100', 'smoke:task4-frozen'
  );
  v_retry := public.continue_financial_reconciliation_automatic_analysis(
    'd4500000-0000-0000-0000-000000000100', 'smoke:task4-frozen'
  );
  if v_second->'status' is distinct from '"ready"'::jsonb
    or v_second->'analysisComplete' is distinct from 'true'::jsonb
    or v_second->'analysisProcessed' is distinct from '26'::jsonb
    or v_second->'analysisTotal' is distinct from '26'::jsonb
    or v_second#>'{counts,bases}' is distinct from '26'::jsonb
    or v_second#>'{counts,proposed}' is distinct from '26'::jsonb
    or v_retry is distinct from v_second
    or (select checkpoint.analysis_cursor_date
        from task4_adyen_cursor_checkpoint checkpoint) is null
    or (select checkpoint.analysis_cursor_id
        from task4_adyen_cursor_checkpoint checkpoint) is null
    or (select run.analysis_cursor_date
        from public.financial_reconciliation_automatic_runs run
        where run.id = 'd4500000-0000-0000-0000-000000000100')
      is distinct from date '2030-03-01'
    or (select run.analysis_cursor_id
        from public.financial_reconciliation_automatic_runs run
        where run.id = 'd4500000-0000-0000-0000-000000000100')
      is distinct from 'd4500000-0000-0000-0000-000000000026'::uuid
    or ((select run.analysis_cursor_date
         from public.financial_reconciliation_automatic_runs run
         where run.id = 'd4500000-0000-0000-0000-000000000100')
        > (select checkpoint.analysis_cursor_date
           from task4_adyen_cursor_checkpoint checkpoint))
      is distinct from true
    or (select count(*)
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = 'd4500000-0000-0000-0000-000000000100')
      is distinct from 26
    or (select count(*)
        from public.financial_reconciliation_automatic_proposal_memberships member
        join public.financial_reconciliation_automatic_proposals proposal
          on proposal.id = member.proposal_id
        where proposal.run_id = 'd4500000-0000-0000-0000-000000000100')
      is distinct from 52
    or exists (
      select 1
      from public.financial_reconciliation_automatic_adyen_population population
      where population.run_id = 'd4500000-0000-0000-0000-000000000100'
        and population.source_id in (
          'd4510000-0000-0000-0000-000000000001',
          'd4510000-0000-0000-0000-000000000002',
          'd4510000-0000-0000-0000-000000000003',
          'e4510000-0000-0000-0000-000000000001',
          'e4510000-0000-0000-0000-000000000002',
          'e4510000-0000-0000-0000-000000000003'
        )
    ) then
    raise exception 'Task 4 frozen Adyen second page/retry changed population: %, %.',
      v_second, v_retry;
  end if;
end
$$;

-- Task 4 Adyen frozen population fails closed on changed future members
insert into public.financial_reconciliation_automatic_runs (
  id, trigger, scope, status, actor, client_request_id,
  definition_config_snapshot, analysis_processed, analysis_total
)
select fixture.run_id, 'manual', 'rule', 'analyzing', fixture.actor,
       fixture.run_id, pg_temp.task4_adyen_snapshot(0.00), 0, 26
from (values
  ('d4600000-0000-0000-0000-000000000101'::uuid,
   'smoke:task4-frozen-consumed'),
  ('d4600000-0000-0000-0000-000000000102'::uuid,
   'smoke:task4-frozen-mutated'),
  ('d4600000-0000-0000-0000-000000000103'::uuid,
   'smoke:task4-frozen-deleted')
) fixture(run_id, actor);

insert into public.financial_reconciliation_automatic_adyen_population (
  run_id, calendar_month, month_ordinal, role, source_type, source_id,
  member_ordinal, source_date, amount, description, account, row_snapshot
)
select fixture.run_id, population.calendar_month, population.month_ordinal,
       population.role, population.source_type, population.source_id,
       population.member_ordinal, population.source_date, population.amount,
       population.description, population.account, population.row_snapshot
from public.financial_reconciliation_automatic_adyen_population population
cross join (values
  ('d4600000-0000-0000-0000-000000000101'::uuid),
  ('d4600000-0000-0000-0000-000000000102'::uuid),
  ('d4600000-0000-0000-0000-000000000103'::uuid)
) fixture(run_id)
where population.run_id = 'd4500000-0000-0000-0000-000000000100';

do $$
declare
  v_case record;
  v_first jsonb;
begin
  for v_case in
    select * from (values
      ('d4600000-0000-0000-0000-000000000101'::uuid,
       'smoke:task4-frozen-consumed'),
      ('d4600000-0000-0000-0000-000000000102'::uuid,
       'smoke:task4-frozen-mutated'),
      ('d4600000-0000-0000-0000-000000000103'::uuid,
       'smoke:task4-frozen-deleted')
    ) fixture(run_id, actor)
  loop
    v_first := public.continue_financial_reconciliation_automatic_analysis(
      v_case.run_id, v_case.actor
    );
    if v_first->'status' is distinct from '"analyzing"'::jsonb
      or v_first->'analysisProcessed' is distinct from '25'::jsonb
      or v_first->'analysisTotal' is distinct from '26'::jsonb then
      raise exception 'Task 4 mutation setup page is invalid for %: %.',
        v_case.run_id, v_first;
    end if;
  end loop;
end
$$;

do $$
declare
  v_result jsonb;
  v_reconciliation_id uuid;
begin
  insert into public.financial_reconciliations (
    status, base_source_type, matching_source_types, created_by
  ) values (
    'started', 'import_cgd_extrato_ordem', '["import_fdm_accounts"]'::jsonb,
    'smoke:task4-frozen-consumed'
  ) returning id into v_reconciliation_id;
  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  ) values (
    v_reconciliation_id, 'import_cgd_extrato_ordem',
    'd4500000-0000-0000-0000-000000000026', 10.00,
    'smoke:task4-frozen-consumed'
  );

  v_result := public.continue_financial_reconciliation_automatic_analysis(
    'd4600000-0000-0000-0000-000000000101',
    'smoke:task4-frozen-consumed'
  );
  if v_result->'status' is distinct from '"failed"'::jsonb
    or v_result->'analysisComplete' is distinct from 'true'::jsonb
    or v_result->'analysisProcessed' is distinct from '25'::jsonb
    or v_result->'analysisTotal' is distinct from '26'::jsonb
    or v_result->'analysisErrorCode' is distinct from
      '"analysis_population_changed"'::jsonb
    or jsonb_typeof(v_result->'finishedAt') is distinct from 'string'
    or exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = 'd4600000-0000-0000-0000-000000000101'
        and proposal.status in ('proposed','ambiguous','executing')
    )
    or (select count(*)
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = 'd4600000-0000-0000-0000-000000000101'
          and proposal.status = 'stale'
          and proposal.reason = 'analysis_population_changed')
      is distinct from 25 then
    raise exception 'Task 4 consumed projected member did not fail safely: %.',
      v_result;
  end if;
  delete from public.financial_reconciliation_items item
  where item.reconciliation_id = v_reconciliation_id;
end
$$;

do $$
declare
  v_result jsonb;
begin
  update public.import_fdm_accounts fdm
  set amount = 11.00
  where fdm.id = 'e4500000-0000-0000-0000-000000000026';
  v_result := public.continue_financial_reconciliation_automatic_analysis(
    'd4600000-0000-0000-0000-000000000102',
    'smoke:task4-frozen-mutated'
  );
  if v_result->'status' is distinct from '"failed"'::jsonb
    or v_result->'analysisErrorCode' is distinct from
      '"analysis_population_changed"'::jsonb
    or exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = 'd4600000-0000-0000-0000-000000000102'
        and proposal.status in ('proposed','ambiguous','executing')
    ) then
    raise exception 'Task 4 mutated projected member did not fail safely: %.',
      v_result;
  end if;
  update public.import_fdm_accounts fdm
  set amount = 10.00
  where fdm.id = 'e4500000-0000-0000-0000-000000000026';
end
$$;

do $$
declare
  v_result jsonb;
begin
  delete from public.import_fdm_accounts fdm
  where fdm.id = 'e4500000-0000-0000-0000-000000000026';
  v_result := public.continue_financial_reconciliation_automatic_analysis(
    'd4600000-0000-0000-0000-000000000103',
    'smoke:task4-frozen-deleted'
  );
  if v_result->'status' is distinct from '"failed"'::jsonb
    or v_result->'analysisErrorCode' is distinct from
      '"analysis_population_changed"'::jsonb
    or exists (
      select 1 from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = 'd4600000-0000-0000-0000-000000000103'
        and proposal.status in ('proposed','ambiguous','executing')
    ) then
    raise exception 'Task 4 deleted projected member did not fail safely: %.',
      v_result;
  end if;
end
$$;

-- Task 4 finalizer literal seven-tuple allowlist
insert into public.financial_reconciliation_automatic_runs (
  id, trigger, scope, status, actor, client_request_id,
  definition_config_snapshot, analysis_processed, analysis_total
) values
  (
    'd4800000-0000-0000-0000-000000000001', 'manual', 'rule', 'analyzing',
    'smoke:task4-finalizer-unknown',
    'd4800000-0000-0000-0000-000000000001',
    '[{"ruleKey":"unknown_rule","ruleVersion":1}]'::jsonb, 0, 0
  ),
  (
    'd4800000-0000-0000-0000-000000000002', 'manual', 'rule', 'analyzing',
    'smoke:task4-finalizer-bank',
    'd4800000-0000-0000-0000-000000000002',
    pg_temp.task3_bank_snapshot(3), 0, 0
  );

do $$
declare
  v_unknown_rejected boolean := false;
  v_bank_rejected boolean := false;
  v_completed_bank jsonb;
begin
  begin
    perform public.financial_reconciliation_finalize_automatic_analysis(
      'd4800000-0000-0000-0000-000000000001'
    );
  exception when others then
    v_unknown_rejected := sqlerrm =
      'Automatic reconciliation rule is unsupported.';
  end;
  begin
    perform public.financial_reconciliation_finalize_automatic_analysis(
      'd4800000-0000-0000-0000-000000000002'
    );
  exception when others then
    v_bank_rejected := sqlerrm =
      'Automatic Bank Reservation analysis uses strategy-specific finalization.';
  end;
  v_completed_bank := public.financial_reconciliation_finalize_automatic_analysis(
    'b3300000-0000-0000-0000-000000000001'
  );
  if not v_unknown_rejected or not v_bank_rejected
    or v_completed_bank->'status' is distinct from '"ready"'::jsonb
    or v_completed_bank->'analysisComplete' is distinct from 'true'::jsonb
    or exists (
      select 1 from public.financial_reconciliation_automatic_runs run
      where run.id in (
          'd4800000-0000-0000-0000-000000000001',
          'd4800000-0000-0000-0000-000000000002'
        )
        and (
          run.status is distinct from 'analyzing'
          or run.analysis_completed_at is not null
        )
    ) then
    raise exception 'Task 4 finalizer did not fail closed for unknown/Bank tuples.';
  end if;
end
$$;

-- Task 4 Adyen population reapply security and run cascade
insert into public.financial_reconciliation_automatic_runs (
  id, trigger, scope, status, actor, client_request_id,
  definition_config_snapshot, analysis_processed, analysis_total
) values (
  'd4700000-0000-0000-0000-000000000001', 'manual', 'rule', 'analyzing',
  'smoke:task4-cascade', 'd4700000-0000-0000-0000-000000000001',
  pg_temp.task4_adyen_snapshot(0.00), 0, 1
);
insert into public.financial_reconciliation_automatic_adyen_population (
  run_id, calendar_month, month_ordinal, role, source_type, source_id,
  member_ordinal, source_date, amount, description, account, row_snapshot
)
select 'd4700000-0000-0000-0000-000000000001', date_trunc('month', bank.data),
       1, 'source', 'import_cgd_extrato_ordem', bank.id, 1, bank.data,
       bank.montante, bank.descritivo, '', to_jsonb(bank)
from public.import_cgd_extrato_ordem bank
where bank.id = 'd4500000-0000-0000-0000-000000000001';

create temporary table task4_adyen_before_reapply on commit drop as
select
  (select count(*) from public.financial_reconciliation_automatic_proposals
   where run_id = 'd4400000-0000-0000-0000-000000000001') as proposal_count,
  (select count(*)
   from public.financial_reconciliation_automatic_proposal_memberships member
   join public.financial_reconciliation_automatic_proposals proposal
     on proposal.id = member.proposal_id
   where proposal.run_id = 'd4400000-0000-0000-0000-000000000001')
    as membership_count;

-- Task 4 POS v2 helper definitions remain byte-equivalent
create temporary table task4_pos_v2_before_reapply on commit drop as
select procedure.oid::regprocedure::text as signature,
       pg_get_functiondef(procedure.oid) as definition
from pg_proc procedure
where procedure.oid in (
  'public.financial_reconciliation_automatic_monthly_income_count()'::regprocedure,
  'public.financial_reconciliation_automatic_monthly_income_page(date,integer)'::regprocedure,
  'public.financial_reconciliation_continue_automatic_monthly_income(uuid,jsonb)'::regprocedure
);

create temporary table task4_bank_before_reapply on commit drop as
select run.status, run.analysis_processed, run.analysis_total, run.counts,
       (select count(*)
        from public.financial_reconciliation_automatic_bank_reservation_population population
        where population.run_id = run.id) as population_count
from public.financial_reconciliation_automatic_runs run
where run.id = 'b3300000-0000-0000-0000-000000000001';

\ir ../supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql

do $$
declare
  v_retry jsonb;
  v_before record;
  v_role text;
  v_privilege text;
begin
  select * into strict v_before from task4_adyen_before_reapply;
  v_retry := public.continue_financial_reconciliation_automatic_analysis(
    'd4400000-0000-0000-0000-000000000001',
    'smoke:task4-adyen'
  );

  if v_retry->'status' is distinct from '"ready"'::jsonb
    or v_retry->'analysisProcessed' is distinct from '6'::jsonb
    or v_retry->'analysisTotal' is distinct from '6'::jsonb
    or (select count(*) from public.financial_reconciliation_automatic_proposals
        where run_id = 'd4400000-0000-0000-0000-000000000001')
      is distinct from v_before.proposal_count
    or (select count(*)
        from public.financial_reconciliation_automatic_proposal_memberships member
        join public.financial_reconciliation_automatic_proposals proposal
          on proposal.id = member.proposal_id
        where proposal.run_id = 'd4400000-0000-0000-0000-000000000001')
      is distinct from v_before.membership_count then
    raise exception 'Task 4 retry or migration reapply duplicated Adyen evidence: %.',
      v_retry;
  end if;

  if exists (
    select before.signature
    from task4_pos_v2_before_reapply before
    full join (
      select procedure.oid::regprocedure::text as signature,
             pg_get_functiondef(procedure.oid) as definition
      from pg_proc procedure
      where procedure.oid in (
        'public.financial_reconciliation_automatic_monthly_income_count()'::regprocedure,
        'public.financial_reconciliation_automatic_monthly_income_page(date,integer)'::regprocedure,
        'public.financial_reconciliation_continue_automatic_monthly_income(uuid,jsonb)'::regprocedure
      )
    ) after using (signature)
    where before.signature is null or after.signature is null
      or before.definition is distinct from after.definition
  ) then
    raise exception 'Task 4 changed an installed POS v2 monthly helper definition.';
  end if;

  if exists (
    select 1
    from (values
      ('public.financial_reconciliation_automatic_adyen_month_count()'),
      ('public.financial_reconciliation_automatic_adyen_month_page(date,integer)'),
      ('public.financial_reconciliation_continue_automatic_adyen_monthly(uuid,text)'),
      ('public.financial_reconciliation_finalize_automatic_prior_analysis(uuid)')
    ) expected(signature)
    where has_function_privilege('anon', expected.signature, 'EXECUTE')
      or has_function_privilege('authenticated', expected.signature, 'EXECUTE')
      or has_function_privilege('service_role', expected.signature, 'EXECUTE')
  ) then
    raise exception 'Task 4 private helper unexpectedly exposes EXECUTE.';
  end if;

  if (select count(*)
      from information_schema.columns column_row
      where column_row.table_schema = 'public'
        and column_row.table_name =
          'financial_reconciliation_automatic_adyen_population')
      is distinct from 12
    or (select count(*)
        from pg_constraint constraint_row
        where constraint_row.conrelid =
          'public.financial_reconciliation_automatic_adyen_population'::regclass)
      is distinct from 8
    or (select count(*)
        from pg_index index_row
        where index_row.indrelid =
          'public.financial_reconciliation_automatic_adyen_population'::regclass)
      is distinct from 2
    or not (select relation.relrowsecurity and not relation.relforcerowsecurity
            from pg_class relation
            where relation.oid =
              'public.financial_reconciliation_automatic_adyen_population'::regclass)
    or exists (
      select 1 from pg_policy policy_row
      where policy_row.polrelid =
        'public.financial_reconciliation_automatic_adyen_population'::regclass
    )
    or (select count(*)
        from public.financial_reconciliation_automatic_adyen_population population
        where population.run_id = 'd4700000-0000-0000-0000-000000000001')
      is distinct from 1 then
    raise exception 'Task 4 Adyen population reapply/catalog contract changed.';
  end if;

  foreach v_role in array array['anon','authenticated','service_role']
  loop
    foreach v_privilege in array array['SELECT','INSERT','UPDATE','DELETE']
    loop
      if has_table_privilege(
          v_role,
          'public.financial_reconciliation_automatic_adyen_population',
          v_privilege
        ) then
        raise exception 'Task 4 Adyen population grants % to %.',
          v_privilege, v_role;
      end if;
    end loop;
  end loop;

  if exists (
    select 1
    from task4_bank_before_reapply before
    full join (
      select run.status, run.analysis_processed, run.analysis_total, run.counts,
             (select count(*)
              from public.financial_reconciliation_automatic_bank_reservation_population population
              where population.run_id = run.id) as population_count
      from public.financial_reconciliation_automatic_runs run
      where run.id = 'b3300000-0000-0000-0000-000000000001'
    ) after using (
      status, analysis_processed, analysis_total, counts, population_count
    )
    where before.status is null or after.status is null
  ) then
    raise exception 'Task 4 reapply changed the Task 3 Bank population lifecycle.';
  end if;

  delete from public.financial_reconciliation_automatic_runs run
  where run.id = 'd4700000-0000-0000-0000-000000000001';
  if exists (
    select 1
    from public.financial_reconciliation_automatic_adyen_population population
    where population.run_id = 'd4700000-0000-0000-0000-000000000001'
  ) then
    raise exception 'Task 4 Adyen population did not cascade with its run.';
  end if;
end
$$;

-- Task 4 Adyen population rejects same-named wrong-column FK
savepoint task4_adyen_wrong_column_fk_fixture;

delete from public.financial_reconciliation_automatic_adyen_population;
alter table public.financial_reconciliation_automatic_adyen_population
  drop constraint fr_auto_adyen_population_run_fkey;
alter table public.financial_reconciliation_automatic_adyen_population
  add constraint fr_auto_adyen_population_run_fkey
  foreign key (source_id)
  references public.financial_reconciliation_automatic_runs(id)
  on delete cascade;

\set ON_ERROR_STOP off
\ir ../supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql
\set task4_adyen_wrong_column_fk_rejected :ERROR
\set ON_ERROR_STOP on

rollback to savepoint task4_adyen_wrong_column_fk_fixture;

\if :task4_adyen_wrong_column_fk_rejected
\else
  \echo 'Task 4 migration accepted a same-named Adyen population FK on source_id instead of run_id.'
  \quit 1
\endif

\ir ../supabase-migrations/2026-08-23-financial-reconciliation-automation-fdm-bank-adyen-rules.sql

do $$
declare
  v_foreign_key record;
begin
  select constraint_row.conkey, constraint_row.confkey
  into v_foreign_key
  from pg_constraint constraint_row
  where constraint_row.conrelid =
      'public.financial_reconciliation_automatic_adyen_population'::regclass
    and constraint_row.conname = 'fr_auto_adyen_population_run_fkey';

  if not found
    or v_foreign_key.conkey is distinct from array[
      (select attribute_row.attnum
       from pg_attribute attribute_row
       where attribute_row.attrelid =
         'public.financial_reconciliation_automatic_adyen_population'::regclass
         and attribute_row.attname = 'run_id'
         and not attribute_row.attisdropped)
    ]::smallint[]
    or v_foreign_key.confkey is distinct from array[
      (select attribute_row.attnum
       from pg_attribute attribute_row
       where attribute_row.attrelid =
         'public.financial_reconciliation_automatic_runs'::regclass
         and attribute_row.attname = 'id'
         and not attribute_row.attisdropped)
    ]::smallint[]
    or (select count(*)
        from public.financial_reconciliation_automatic_adyen_population population
        where population.run_id = 'd4500000-0000-0000-0000-000000000100')
      is distinct from 52 then
    raise exception 'Task 4 Adyen wrong-column FK fixture did not restore exact projection state.';
  end if;
end
$$;

-- Task 5 Bank Reservation executes all eleven immutable members and retries idempotently
-- Task 5 Adyen zero and allowed nonzero execution preserve history and audit
-- Task 5 member identity type date amount description and Account drift
-- Task 5 group eligibility configuration definition operator and overlap drift
-- Task 5 deletion consumption and source-lock stale outcomes are atomic
-- Task 5 malformed snapshots fail stale without unsafe casts
-- Task 5 post-start and competing writes roll back every lifecycle row
-- Task 5 unexpected database failures persist only sanitized diagnostics
-- Task 5 execution helpers are private and top-level dispatch is literal

-- Task 4 intentionally left analysis-only proposals reviewable. They have
-- already served their paging assertions; deselect them so Task 5 can prove a
-- fresh successful execution without manufacturing a cross-proposal overlap.
update public.financial_reconciliation_automatic_proposals
set status = 'deselected', updated_at = now()
where run_id = 'd4500000-0000-0000-0000-000000000100'
  and status in ('proposed','ambiguous');

update public.financial_reconciliation_source_rules source_rule
set operator = case
      when source_rule.base_source_type = 'import_fdm_accounts' then '+'
      else '-'
    end
where (source_rule.base_source_type, source_rule.matching_source_type) in (
  ('import_fdm_accounts', 'import_cgd_extrato_ordem'),
  ('import_cgd_extrato_ordem', 'import_fdm_accounts')
);

update public.financial_reconciliation_automatic_rule_configs config
set enabled = true,
    allow_manual_execution = true,
    difference_allowed = case
      when config.rule_key =
        'fdm_bank_transfer_cgd_bank_statement_combination' then 0
      else 20
    end,
    max_difference_days = case
      when config.rule_key =
        'fdm_bank_transfer_cgd_bank_statement_combination' then 3
      else 31
    end
where config.rule_key in (
  'fdm_bank_transfer_cgd_bank_statement_combination',
  'cgd_bank_statement_fdm_adyen_monthly_payments'
);

insert into public.import_cgd_extrato_ordem (
  id, import_batch, row_key, data, descritivo, montante
) values (
  'f5100000-0000-0000-0000-000000000001',
  'smoke-task5-bank', 'task5-bank-reservation', date '2099-11-15',
  'Reservation settlement', 100.00
);

insert into public.import_fdm_accounts (
  id, import_batch, source_row_number, account, date_time_raw, event_date,
  category, amount, description
)
select
  ('f5110000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
  'smoke-task5-bank', series, 'Bank Transfer', '2099-11-15',
  date '2099-11-15', 'Reservation',
  case when series < 10 then -1.00 else -91.00 end,
  'Task 5 reservation member ' || series
from generate_series(1, 10) series;

insert into public.financial_reconciliation_automatic_runs (
  id, trigger, scope, status, actor, client_request_id,
  definition_config_snapshot, analysis_processed, analysis_total
) values (
  'f5200000-0000-0000-0000-000000000001', 'manual', 'rule', 'analyzing',
  'smoke:task5-bank', 'f5200000-0000-0000-0000-000000000001',
  pg_temp.task3_bank_snapshot(3), 0, 0
);

do $$
declare
  v_result jsonb;
  v_attempt integer := 0;
begin
  loop
    v_attempt := v_attempt + 1;
    v_result := public.continue_financial_reconciliation_automatic_analysis(
      'f5200000-0000-0000-0000-000000000001', 'smoke:task5-bank'
    );
    exit when v_result->>'status' <> 'analyzing';
    if v_attempt >= 100 then
      raise exception 'Task 5 Bank analysis did not terminate: %.', v_result;
    end if;
  end loop;
  if v_result->>'status' <> 'ready' then
    raise exception 'Task 5 Bank execution fixture was not reviewable: %.', v_result;
  end if;
end
$$;

create temporary table task5_bank_target on commit drop as
select proposal.id as proposal_id, proposal.run_id,
       proposal.candidate_groups, proposal.evidence,
       proposal.summary_snapshot, proposal.base_snapshot
from public.financial_reconciliation_automatic_proposals proposal
where proposal.run_id = 'f5200000-0000-0000-0000-000000000001'
  and proposal.grouping_key = 'f5100000-0000-0000-0000-000000000001'
  and proposal.status = 'proposed';

do $$
begin
  if (select count(*) from task5_bank_target) <> 1
    or (select count(*)
        from public.financial_reconciliation_automatic_proposal_memberships member
        join task5_bank_target target on target.proposal_id = member.proposal_id)
      <> 11
    or (select run.definition_config_snapshot#>>'{0,operator}'
        from public.financial_reconciliation_automatic_runs run
        where run.id = 'f5200000-0000-0000-0000-000000000001')
      is distinct from '+'
    or (select target.summary_snapshot->>'operator'
        from task5_bank_target target) is distinct from '+'
    or not exists (
      select 1 from public.financial_reconciliation_source_rules source_rule
      where source_rule.base_source_type = 'import_fdm_accounts'
        and source_rule.matching_source_type = 'import_cgd_extrato_ordem'
        and source_rule.operator = '+'
    ) then
    raise exception 'Task 5 Bank analysis did not produce the exact 10 + 1 group.';
  end if;
end
$$;

create or replace function pg_temp.task5_clone_grouped_proposal(
  p_source_proposal_id uuid,
  p_run_id uuid,
  p_proposal_id uuid,
  p_actor text,
  p_membership_mutation text default ''
)
returns uuid
language plpgsql
set search_path = public, pg_temp
as $$
declare
  v_source_run_id uuid;
begin
  select proposal.run_id into strict v_source_run_id
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.id = p_source_proposal_id;

  insert into public.financial_reconciliation_automatic_runs (
    id, trigger, scope, status, actor, client_request_id,
    definition_config_snapshot, counts, analysis_completed_at
  )
  select p_run_id, 'manual', 'rule', 'ready', p_actor, p_run_id,
         source_run.definition_config_snapshot, source_run.counts, now()
  from public.financial_reconciliation_automatic_runs source_run
  where source_run.id = v_source_run_id;

  insert into public.financial_reconciliation_automatic_proposals (
    id, run_id, rule_key, rule_version, base_source_type, base_source_id,
    base_source_date, base_snapshot, items, evidence, candidate_groups,
    calculated_difference, allowed_difference, status, reason, signature,
    reconciliation_id, error, error_detail, completed_at,
    grouping_key, summary_snapshot
  )
  select p_proposal_id, p_run_id, proposal.rule_key, proposal.rule_version,
         proposal.base_source_type, proposal.base_source_id,
         proposal.base_source_date, proposal.base_snapshot, proposal.items,
         proposal.evidence, proposal.candidate_groups,
          proposal.calculated_difference, proposal.allowed_difference,
          'proposed', proposal.reason, proposal.signature, null, '', '', null,
         proposal.grouping_key, proposal.summary_snapshot
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.id = p_source_proposal_id;

  insert into public.financial_reconciliation_automatic_proposal_memberships (
    proposal_id, role, source_type, source_id, ordinal, source_date,
    amount, description, account, row_snapshot
  )
  select
    p_proposal_id,
    membership.role,
    case
      when p_membership_mutation = 'member_type'
        and membership.role = 'source' and membership.ordinal = 1
        then case membership.source_type
          when 'import_fdm_accounts' then 'import_cgd_extrato_ordem'
          else 'import_fdm_accounts'
        end
      else membership.source_type
    end,
    case
      when p_membership_mutation = 'member_id'
        and membership.role = 'source' and membership.ordinal = 1
        then 'f5ff0000-0000-0000-0000-000000000001'::uuid
      else membership.source_id
    end,
    membership.ordinal, membership.source_date, membership.amount,
    membership.description, membership.account, membership.row_snapshot
  from public.financial_reconciliation_automatic_proposal_memberships membership
  where membership.proposal_id = p_source_proposal_id
    and not (
      p_membership_mutation = 'group_count'
      and membership.role = 'source'
      and membership.ordinal = (
        select max(last_member.ordinal)
        from public.financial_reconciliation_automatic_proposal_memberships last_member
        where last_member.proposal_id = p_source_proposal_id
          and last_member.role = 'source'
      )
    )
  order by membership.source_type, membership.source_id;

  return p_proposal_id;
end
$$;

create or replace function pg_temp.task5_assert_stale(
  p_proposal_id uuid,
  p_reason text
)
returns void
language plpgsql
set search_path = public, pg_temp
as $$
declare
  v_actor text;
  v_run_id uuid;
  v_candidate_groups jsonb;
  v_evidence jsonb;
  v_summary_snapshot jsonb;
  v_base_snapshot jsonb;
  v_result jsonb;
  v_reconciliation_count bigint;
  v_item_count bigint;
  v_audit_count bigint;
begin
  select run.actor, run.id,
         proposal.candidate_groups, proposal.evidence,
         proposal.summary_snapshot, proposal.base_snapshot
  into strict v_actor, v_run_id,
       v_candidate_groups, v_evidence,
       v_summary_snapshot, v_base_snapshot
  from public.financial_reconciliation_automatic_proposals proposal
  join public.financial_reconciliation_automatic_runs run
    on run.id = proposal.run_id
  where proposal.id = p_proposal_id;

  select count(*) into v_reconciliation_count
  from public.financial_reconciliations;
  select count(*) into v_item_count
  from public.financial_reconciliation_items;
  select count(*) into v_audit_count
  from public.financial_reconciliation_audit;

  v_result := public.execute_financial_reconciliation_automatic_proposal(
    p_proposal_id, v_actor
  );
  if v_result is distinct from jsonb_build_object(
      'proposalId', p_proposal_id,
      'runId', v_run_id,
      'status', 'stale',
      'reason', p_reason
    )
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.id = p_proposal_id
        and proposal.status = 'stale'
        and proposal.reason = p_reason
        and proposal.reconciliation_id is null
        and proposal.completed_at is null
        and proposal.error = ''
        and proposal.error_detail = ''
        and proposal.candidate_groups is not distinct from
          v_candidate_groups
        and proposal.evidence is not distinct from v_evidence
        and proposal.summary_snapshot is not distinct from
          v_summary_snapshot
        and proposal.base_snapshot is not distinct from v_base_snapshot
    )
    or (select count(*) from public.financial_reconciliations) <>
      v_reconciliation_count
    or (select count(*) from public.financial_reconciliation_items) <>
      v_item_count
    or (select count(*) from public.financial_reconciliation_audit) <>
      v_audit_count then
    raise exception 'Task 5 stale execution was not atomic/sanitized for %: %.',
      p_proposal_id, v_result;
  end if;
end
$$;

create or replace function pg_temp.task5_assert_failed(p_proposal_id uuid)
returns void
language plpgsql
set search_path = public, pg_temp
as $$
declare
  v_actor text;
  v_run_id uuid;
  v_result jsonb;
  v_reconciliation_count bigint;
  v_item_count bigint;
  v_audit_count bigint;
begin
  select run.actor, run.id into strict v_actor, v_run_id
  from public.financial_reconciliation_automatic_proposals proposal
  join public.financial_reconciliation_automatic_runs run
    on run.id = proposal.run_id
  where proposal.id = p_proposal_id;
  select count(*) into v_reconciliation_count
  from public.financial_reconciliations;
  select count(*) into v_item_count
  from public.financial_reconciliation_items;
  select count(*) into v_audit_count
  from public.financial_reconciliation_audit;

  v_result := public.execute_financial_reconciliation_automatic_proposal(
    p_proposal_id, v_actor
  );
  if v_result is distinct from jsonb_build_object(
      'proposalId', p_proposal_id,
      'runId', v_run_id,
      'status', 'failed',
      'reason', 'execution_failed'
    )
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.id = p_proposal_id
        and proposal.status = 'failed'
        and proposal.reason = 'execution_failed'
        and proposal.reconciliation_id is null
        and proposal.completed_at is null
        and proposal.error = 'Automatic reconciliation execution failed.'
        and proposal.error_detail = ''
    )
    or v_result::text like '%task5 secret%'
    or (select proposal.error || proposal.error_detail
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.id = p_proposal_id) like '%task5 secret%'
    or (select count(*) from public.financial_reconciliations) <>
      v_reconciliation_count
    or (select count(*) from public.financial_reconciliation_items) <>
      v_item_count
    or (select count(*) from public.financial_reconciliation_audit) <>
      v_audit_count then
    raise exception 'Task 5 failure did not roll back/sanitize: %.', v_result;
  end if;
end
$$;

-- bank execution becomes stale when a new live candidate creates a second qualifying combination
do $$
declare
  v_source_proposal_id uuid :=
    (select target.proposal_id from task5_bank_target target);
begin
  begin
    update public.financial_reconciliation_automatic_proposals
    set status = 'deselected' where id = v_source_proposal_id;
    perform pg_temp.task5_clone_grouped_proposal(
      v_source_proposal_id,
      'f5400000-0000-0000-0000-000000000001',
      'f5410000-0000-0000-0000-000000000001',
      'smoke:task5-live-second-group', ''
    );
    insert into public.import_fdm_accounts (
      id, import_batch, source_row_number, account, date_time_raw, event_date,
      category, amount, description
    ) values (
      'f5120000-0000-0000-0000-000000000001',
      'smoke-task5-live-second-group', 1, 'Bank Transfer', '2099-11-15',
      date '2099-11-15', 'Reservation', -100.00,
      'Task 5 new independently qualifying candidate'
    );
    perform pg_temp.task5_assert_stale(
      'f5410000-0000-0000-0000-000000000001',
      'source_snapshot_changed'
    );
    raise sqlstate 'T5001' using message = 'rollback task5 live second group';
  exception when sqlstate 'T5001' then null;
  end;
end
$$;

-- bank execution becomes stale when a new live candidate pool reaches candidate_limit
do $$
declare
  v_source_proposal_id uuid :=
    (select target.proposal_id from task5_bank_target target);
begin
  begin
    update public.financial_reconciliation_automatic_proposals
    set status = 'deselected' where id = v_source_proposal_id;
    perform pg_temp.task5_clone_grouped_proposal(
      v_source_proposal_id,
      'f5400000-0000-0000-0000-000000000002',
      'f5410000-0000-0000-0000-000000000002',
      'smoke:task5-live-candidate-limit', ''
    );
    insert into public.import_fdm_accounts (
      id, import_batch, source_row_number, account, date_time_raw, event_date,
      category, amount, description
    )
    select
      ('f5130000-0000-0000-0000-' || lpad(series::text, 12, '0'))::uuid,
      'smoke-task5-live-candidate-limit', series, 'Bank Transfer',
      '2099-11-15', date '2099-11-15', 'Reservation', -0.01,
      'Task 5 candidate-limit addition ' || series
    from generate_series(1, 51) series;
    perform pg_temp.task5_assert_stale(
      'f5410000-0000-0000-0000-000000000002',
      'source_snapshot_changed'
    );
    raise sqlstate 'T5001' using message = 'rollback task5 live candidate limit';
  exception when sqlstate 'T5001' then null;
  end;
end
$$;

-- bank execution accepts a harmless new live candidate when the immutable group stays unique
do $$
declare
  v_source_proposal_id uuid :=
    (select target.proposal_id from task5_bank_target target);
  v_result jsonb;
begin
  begin
    update public.financial_reconciliation_automatic_proposals
    set status = 'deselected' where id = v_source_proposal_id;
    perform pg_temp.task5_clone_grouped_proposal(
      v_source_proposal_id,
      'f5400000-0000-0000-0000-000000000003',
      'f5410000-0000-0000-0000-000000000003',
      'smoke:task5-live-harmless-candidate', ''
    );
    insert into public.import_fdm_accounts (
      id, import_batch, source_row_number, account, date_time_raw, event_date,
      category, amount, description
    ) values (
      'f5140000-0000-0000-0000-000000000001',
      'smoke-task5-live-harmless-candidate', 1, 'Bank Transfer',
      '2099-11-15', date '2099-11-15', 'Reservation', -0.50,
      'Task 5 harmless nonqualifying candidate'
    );
    v_result := public.execute_financial_reconciliation_automatic_proposal(
      'f5410000-0000-0000-0000-000000000003',
      'smoke:task5-live-harmless-candidate'
    );
    if v_result->>'status' is distinct from 'completed'
      or jsonb_typeof(v_result->'reconciliationId') is distinct from 'string'
      or not exists (
        select 1
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.id = 'f5410000-0000-0000-0000-000000000003'
          and proposal.status = 'completed'
          and proposal.reconciliation_id =
            (v_result->>'reconciliationId')::uuid
      ) then
      raise exception 'Task 5 harmless live candidate changed the unique immutable group: %.',
        v_result;
    end if;
    raise sqlstate 'T5001' using message = 'rollback task5 harmless live candidate';
  exception when sqlstate 'T5001' then null;
  end;
end
$$;

-- bank execution becomes stale when changed live candidates leave no qualifying combination
do $$
declare
  v_source_proposal_id uuid :=
    (select target.proposal_id from task5_bank_target target);
begin
  begin
    update public.financial_reconciliation_automatic_proposals
    set status = 'deselected' where id = v_source_proposal_id;
    perform pg_temp.task5_clone_grouped_proposal(
      v_source_proposal_id,
      'f5400000-0000-0000-0000-000000000006',
      'f5410000-0000-0000-0000-000000000006',
      'smoke:task5-live-no-match', ''
    );
    update public.import_fdm_accounts
    set amount = -90.99
    where id = 'f5110000-0000-0000-0000-000000000010';
    perform pg_temp.task5_assert_stale(
      'f5410000-0000-0000-0000-000000000006',
      'source_snapshot_changed'
    );
    raise sqlstate 'T5001' using message = 'rollback task5 live no match';
  exception when sqlstate 'T5001' then null;
  end;
end
$$;

-- candidate-limit evidence does not overlap an otherwise-proposed group
do $$
declare
  v_source_proposal_id uuid :=
    (select target.proposal_id from task5_bank_target target);
  v_result jsonb;
begin
  begin
    update public.financial_reconciliation_automatic_proposals
    set status = 'deselected' where id = v_source_proposal_id;
    perform pg_temp.task5_clone_grouped_proposal(
      v_source_proposal_id,
      'f5400000-0000-0000-0000-000000000004',
      'f5410000-0000-0000-0000-000000000004',
      'smoke:task5-proposed-overlap-authority', ''
    );
    perform pg_temp.task5_clone_grouped_proposal(
      v_source_proposal_id,
      'f5400000-0000-0000-0000-000000000005',
      'f5410000-0000-0000-0000-000000000005',
      'smoke:task5-candidate-limit-evidence', ''
    );
    update public.financial_reconciliation_automatic_proposals
    set status = 'ambiguous', reason = 'candidate_limit',
        summary_snapshot = jsonb_set(
          jsonb_set(summary_snapshot, '{classification}', '"ambiguous"'),
          '{reason}', '"candidate_limit"'
        )
    where id = 'f5410000-0000-0000-0000-000000000005';

    v_result := public.execute_financial_reconciliation_automatic_proposal(
      'f5410000-0000-0000-0000-000000000004',
      'smoke:task5-proposed-overlap-authority'
    );
    if v_result->>'status' is distinct from 'completed'
      or jsonb_typeof(v_result->'reconciliationId') is distinct from 'string'
      or not exists (
        select 1
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.id = 'f5410000-0000-0000-0000-000000000004'
          and proposal.status = 'completed'
      )
      or not exists (
        select 1
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.id = 'f5410000-0000-0000-0000-000000000005'
          and proposal.status = 'ambiguous'
          and proposal.reason = 'candidate_limit'
      ) then
      raise exception 'Task 5 candidate-limit evidence falsely blocked proposed execution: %.',
        v_result;
    end if;
    raise sqlstate 'T5001' using message = 'rollback task5 overlap authority';
  exception when sqlstate 'T5001' then null;
  end;
end
$$;

-- Every case runs in its own subtransaction and deliberately rolls its fixture
-- back after assertions, so the exact same analyzed evidence can prove that
-- each independent mismatch is sufficient to produce stale.
do $$
declare
  v_source_proposal_id uuid :=
    (select target.proposal_id from task5_bank_target target);
  v_mutation text;
  v_member_id uuid;
  v_reconciliation_id uuid;
begin
  foreach v_mutation in array array['member_id','member_type','group_count']
  loop
    begin
      update public.financial_reconciliation_automatic_proposals
      set status = 'deselected' where id = v_source_proposal_id;
      perform pg_temp.task5_clone_grouped_proposal(
        v_source_proposal_id,
        'f5300000-0000-0000-0000-000000000001',
        'f5310000-0000-0000-0000-000000000001',
        'smoke:task5-bank-stale', v_mutation
      );
      perform pg_temp.task5_assert_stale(
        'f5310000-0000-0000-0000-000000000001',
        'source_snapshot_changed'
      );
      raise sqlstate 'T5001' using message = 'rollback task5 fixture';
    exception when sqlstate 'T5001' then null;
    end;
  end loop;

  foreach v_mutation in array array[
    'date','amount','description','account','bank_eligibility','fdm_eligibility'
  ]
  loop
    begin
      update public.financial_reconciliation_automatic_proposals
      set status = 'deselected' where id = v_source_proposal_id;
      perform pg_temp.task5_clone_grouped_proposal(
        v_source_proposal_id,
        'f5300000-0000-0000-0000-000000000001',
        'f5310000-0000-0000-0000-000000000001',
        'smoke:task5-bank-stale', ''
      );
      select member.source_id into strict v_member_id
      from public.financial_reconciliation_automatic_proposal_memberships member
      where member.proposal_id =
          'f5310000-0000-0000-0000-000000000001'
        and member.role = 'source'
      order by member.ordinal limit 1;
      if v_mutation = 'date' then
        update public.import_fdm_accounts set event_date = event_date + 1
        where id = v_member_id;
      elsif v_mutation = 'amount' then
        update public.import_fdm_accounts set amount = amount - 0.01
        where id = v_member_id;
      elsif v_mutation = 'description' then
        update public.import_fdm_accounts set description = description || ' drift'
        where id = v_member_id;
      elsif v_mutation in ('account','fdm_eligibility') then
        update public.import_fdm_accounts set account = 'Bank transfer'
        where id = v_member_id;
      else
        update public.import_cgd_extrato_ordem set data = null
        where id = 'f5100000-0000-0000-0000-000000000001';
      end if;
      perform pg_temp.task5_assert_stale(
        'f5310000-0000-0000-0000-000000000001',
        'source_snapshot_changed'
      );
      raise sqlstate 'T5001' using message = 'rollback task5 fixture';
    exception when sqlstate 'T5001' then null;
    end;
  end loop;

  -- An active source lock and a completed consumption are independently stale.
  foreach v_mutation in array array['source_lock','consumption']
  loop
    begin
      update public.financial_reconciliation_automatic_proposals
      set status = 'deselected' where id = v_source_proposal_id;
      perform pg_temp.task5_clone_grouped_proposal(
        v_source_proposal_id,
        'f5300000-0000-0000-0000-000000000001',
        'f5310000-0000-0000-0000-000000000001',
        'smoke:task5-bank-stale', ''
      );
      select member.source_id into strict v_member_id
      from public.financial_reconciliation_automatic_proposal_memberships member
      where member.proposal_id = 'f5310000-0000-0000-0000-000000000001'
      order by member.source_type, member.source_id limit 1;
      insert into public.financial_reconciliations (
        status, base_source_type, matching_source_types, completion_type,
        difference_amount, created_by, completed_by, completed_at
      ) values (
        case when v_mutation = 'source_lock' then 'started' else 'complete' end,
        'import_fdm_accounts', '["import_cgd_extrato_ordem"]'::jsonb,
        case when v_mutation = 'consumption' then 'normal' else null end,
        0, 'smoke:task5-lock-owner',
        case when v_mutation = 'consumption'
          then 'smoke:task5-lock-owner' else null end,
        case when v_mutation = 'consumption'
          then timezone('utc', now()) else null end
      ) returning id into v_reconciliation_id;
      insert into public.financial_reconciliation_items (
        reconciliation_id, source_type, source_id, amount_snapshot, created_by
      )
      select v_reconciliation_id, member.source_type, member.source_id,
             member.amount, 'smoke:task5-lock-owner'
      from public.financial_reconciliation_automatic_proposal_memberships member
      where member.proposal_id = 'f5310000-0000-0000-0000-000000000001'
        and member.source_id = v_member_id;
      perform pg_temp.task5_assert_stale(
        'f5310000-0000-0000-0000-000000000001',
        'source_snapshot_changed'
      );
      raise sqlstate 'T5001' using message = 'rollback task5 fixture';
    exception when sqlstate 'T5001' then null;
    end;
  end loop;

  -- Deleted live membership.
  begin
    update public.financial_reconciliation_automatic_proposals
    set status = 'deselected' where id = v_source_proposal_id;
    perform pg_temp.task5_clone_grouped_proposal(
      v_source_proposal_id,
      'f5300000-0000-0000-0000-000000000001',
      'f5310000-0000-0000-0000-000000000001',
      'smoke:task5-bank-stale', ''
    );
    select member.source_id into strict v_member_id
    from public.financial_reconciliation_automatic_proposal_memberships member
    where member.proposal_id = 'f5310000-0000-0000-0000-000000000001'
      and member.role = 'source'
    order by member.ordinal limit 1;
    delete from public.import_fdm_accounts where id = v_member_id;
    perform pg_temp.task5_assert_stale(
      'f5310000-0000-0000-0000-000000000001',
      'source_snapshot_changed'
    );
    raise sqlstate 'T5001' using message = 'rollback task5 fixture';
  exception when sqlstate 'T5001' then null;
  end;

  -- An independently proposed overlapping group makes this clone stale.
  begin
    perform pg_temp.task5_clone_grouped_proposal(
      v_source_proposal_id,
      'f5300000-0000-0000-0000-000000000001',
      'f5310000-0000-0000-0000-000000000001',
      'smoke:task5-bank-stale', ''
    );
    perform pg_temp.task5_assert_stale(
      'f5310000-0000-0000-0000-000000000001',
      'source_snapshot_changed'
    );
    raise sqlstate 'T5001' using message = 'rollback task5 fixture';
  exception when sqlstate 'T5001' then null;
  end;
end
$$;

-- Rule/config/operator and malformed textual snapshots are each independently
-- rejected before any numeric cast or reconciliation mutation.
do $$
declare
  v_source_proposal_id uuid :=
    (select target.proposal_id from task5_bank_target target);
  v_mutation text;
begin
  foreach v_mutation in array array[
    'days','priority','definition_strategy','definition_version','operator',
    'malformed_integer','malformed_numeric'
  ]
  loop
    begin
      update public.financial_reconciliation_automatic_proposals
      set status = 'deselected' where id = v_source_proposal_id;
      perform pg_temp.task5_clone_grouped_proposal(
        v_source_proposal_id,
        'f5300000-0000-0000-0000-000000000001',
        'f5310000-0000-0000-0000-000000000001',
        'smoke:task5-bank-stale', ''
      );
      if v_mutation = 'days' then
        update public.financial_reconciliation_automatic_rule_configs
        set max_difference_days = 4
        where rule_key = 'fdm_bank_transfer_cgd_bank_statement_combination';
      elsif v_mutation = 'priority' then
        update public.financial_reconciliation_automatic_rule_configs
        set priority = priority + 100
        where rule_key = 'fdm_bank_transfer_cgd_bank_statement_combination';
      elsif v_mutation = 'definition_strategy' then
        update public.financial_reconciliation_automatic_rule_definitions
        set definition = jsonb_set(
          definition, '{strategy}', '"drifted_strategy"'::jsonb
        )
        where rule_key = 'fdm_bank_transfer_cgd_bank_statement_combination'
          and version = 1;
      elsif v_mutation = 'definition_version' then
        insert into public.financial_reconciliation_automatic_rule_definitions (
          rule_key, version, display_name, base_source_type,
          destination_source_types, logic_description, definition
        )
        select definition.rule_key, 2, definition.display_name,
               definition.base_source_type, definition.destination_source_types,
               definition.logic_description, definition.definition
        from public.financial_reconciliation_automatic_rule_definitions definition
        where definition.rule_key =
            'fdm_bank_transfer_cgd_bank_statement_combination'
          and definition.version = 1;
        update public.financial_reconciliation_automatic_rule_configs
        set rule_version = 2
        where rule_key = 'fdm_bank_transfer_cgd_bank_statement_combination';
      elsif v_mutation = 'operator' then
        update public.financial_reconciliation_source_rules
        set operator = '-'
        where base_source_type = 'import_fdm_accounts'
          and matching_source_type = 'import_cgd_extrato_ordem';
      elsif v_mutation = 'malformed_integer' then
        update public.financial_reconciliation_automatic_runs
        set definition_config_snapshot = jsonb_set(
          definition_config_snapshot, '{0,priority}', '"2147483648x"'::jsonb
        )
        where id = 'f5300000-0000-0000-0000-000000000001';
      else
        update public.financial_reconciliation_automatic_runs
        set definition_config_snapshot = jsonb_set(
          definition_config_snapshot,
          '{0,differenceAllowed}', '"999999999999999999999"'::jsonb
        )
        where id = 'f5300000-0000-0000-0000-000000000001';
      end if;
      perform pg_temp.task5_assert_stale(
        'f5310000-0000-0000-0000-000000000001',
        case
          when v_mutation = 'operator' then 'operator_changed'
          else 'rule_snapshot_changed'
        end
      );
      raise sqlstate 'T5001' using message = 'rollback task5 fixture';
    exception when sqlstate 'T5001' then null;
    end;
  end loop;
end
$$;

insert into public.import_cgd_extrato_ordem (
  id, import_batch, row_key, data, descritivo, montante
) values
  (
    'f5400000-0000-0000-0000-000000000001',
    'smoke-task5-adyen', 'task5-adyen-june', date '2026-06-20',
    'Merchant ADYEN settlement', 60.00
  ),
  (
    'f5400000-0000-0000-0000-000000000002',
    'smoke-task5-adyen', 'task5-adyen-july', date '2026-07-20',
    'Adyen settlement', 25.00
  );

insert into public.import_fdm_accounts (
  id, import_batch, account, date_time_raw, event_date, category,
  amount, description
) values (
  'f5410000-0000-0000-0000-000000000001',
  'smoke-task5-adyen', 'Adyen', '2026-07-20', date '2026-07-20',
  'Reservation', 25.00, 'Task 5 Adyen July'
);

insert into public.financial_reconciliation_automatic_runs (
  id, trigger, scope, status, actor, client_request_id,
  definition_config_snapshot, analysis_processed, analysis_total
) values (
  'f5500000-0000-0000-0000-000000000001', 'manual', 'rule', 'analyzing',
  'smoke:task5-adyen', 'f5500000-0000-0000-0000-000000000001',
  pg_temp.task4_adyen_snapshot(20.00), 0, 0
);

do $$
declare
  v_result jsonb;
  v_attempt integer := 0;
begin
  loop
    v_attempt := v_attempt + 1;
    v_result := public.continue_financial_reconciliation_automatic_analysis(
      'f5500000-0000-0000-0000-000000000001', 'smoke:task5-adyen'
    );
    exit when v_result->>'status' <> 'analyzing';
    if v_attempt >= 10 then
      raise exception 'Task 5 Adyen analysis did not terminate: %.', v_result;
    end if;
  end loop;
  if v_result->>'status' <> 'ready'
    or (select count(*)
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = 'f5500000-0000-0000-0000-000000000001'
          and proposal.grouping_key in ('2026-06','2026-07')
          and proposal.status = 'proposed') <> 2 then
    raise exception 'Task 5 Adyen execution months were not proposed: %.', v_result;
  end if;
end
$$;

create temporary table task5_adyen_targets on commit drop as
select proposal.grouping_key, proposal.id as proposal_id
from public.financial_reconciliation_automatic_proposals proposal
where proposal.run_id = 'f5500000-0000-0000-0000-000000000001'
  and proposal.grouping_key in ('2026-06','2026-07');

-- Adyen Description, Account, group expansion, and allowance drift each stale
-- an otherwise unchanged month and preserve its immutable evidence.
do $$
declare
  v_source_proposal_id uuid := (
    select target.proposal_id from task5_adyen_targets target
    where target.grouping_key = '2026-07'
  );
  v_mutation text;
begin
  foreach v_mutation in array array[
    'bank_description','fdm_account','new_member','allowance'
  ]
  loop
    begin
      update public.financial_reconciliation_automatic_proposals
      set status = 'deselected' where id = v_source_proposal_id;
      perform pg_temp.task5_clone_grouped_proposal(
        v_source_proposal_id,
        'f5600000-0000-0000-0000-000000000001',
        'f5610000-0000-0000-0000-000000000001',
        'smoke:task5-adyen-stale', ''
      );
      if v_mutation = 'bank_description' then
        update public.import_cgd_extrato_ordem
        set descritivo = 'Payment provider'
        where id = 'f5400000-0000-0000-0000-000000000002';
      elsif v_mutation = 'fdm_account' then
        update public.import_fdm_accounts set account = 'adyen'
        where id = 'f5410000-0000-0000-0000-000000000001';
      elsif v_mutation = 'new_member' then
        insert into public.import_fdm_accounts (
          id, import_batch, account, date_time_raw, event_date, category,
          amount, description
        ) values (
          'f5410000-0000-0000-0000-000000000002',
          'smoke-task5-adyen', 'Adyen', '2026-07-21', date '2026-07-21',
          'Reservation', 1.00, 'Task 5 gained member'
        );
      else
        update public.financial_reconciliation_automatic_rule_configs
        set difference_allowed = 21
        where rule_key = 'cgd_bank_statement_fdm_adyen_monthly_payments';
      end if;
      perform pg_temp.task5_assert_stale(
        'f5610000-0000-0000-0000-000000000001',
        case
          when v_mutation = 'allowance' then 'tolerance_changed'
          else 'source_snapshot_changed'
        end
      );
      raise sqlstate 'T5001' using message = 'rollback task5 fixture';
    exception when sqlstate 'T5001' then null;
    end;
  end loop;
end
$$;

create or replace function pg_temp.task5_competing_item_write()
returns trigger
language plpgsql
set search_path = public, pg_temp
as $$
declare
  v_competing_reconciliation_id uuid;
begin
  if new.created_by <> 'smoke:task5-competing' then
    return new;
  end if;
  insert into public.financial_reconciliations (
    status, base_source_type, matching_source_types, created_by
  ) values (
    'started', 'import_fdm_accounts',
    '["import_cgd_extrato_ordem"]'::jsonb,
    'smoke:task5-competing-owner'
  ) returning id into v_competing_reconciliation_id;
  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  )
  select v_competing_reconciliation_id, member.source_type,
         member.source_id, member.amount, 'smoke:task5-competing-owner'
  from public.financial_reconciliation_automatic_proposal_memberships member
  where member.proposal_id = 'f5310000-0000-0000-0000-000000000001'
  order by member.source_type, member.source_id
  limit 1;
  return new;
end
$$;

create or replace function pg_temp.task5_secret_audit_failure()
returns trigger
language plpgsql
set search_path = public, pg_temp
as $$
begin
  if new.actor = 'smoke:task5-unexpected'
    and new.action = 'automatic_complete' then
    raise exception 'task5 secret database diagnostics must never escape';
  end if;
  return new;
end
$$;

do $$
declare
  v_source_proposal_id uuid :=
    (select target.proposal_id from task5_bank_target target);
begin
  -- The trigger performs a competing ownership write after reconciliation
  -- creation. The executor must classify the unique source lock as stale and
  -- roll both its lifecycle rows and the simulated competing rows back.
  begin
    update public.financial_reconciliation_automatic_proposals
    set status = 'deselected' where id = v_source_proposal_id;
    perform pg_temp.task5_clone_grouped_proposal(
      v_source_proposal_id,
      'f5300000-0000-0000-0000-000000000001',
      'f5310000-0000-0000-0000-000000000001',
      'smoke:task5-competing', ''
    );
    create trigger task5_competing_item_write
      after insert on public.financial_reconciliations
      for each row execute function pg_temp.task5_competing_item_write();
    perform pg_temp.task5_assert_stale(
      'f5310000-0000-0000-0000-000000000001',
      'source_snapshot_changed'
    );
    raise sqlstate 'T5001' using message = 'rollback task5 fixture';
  exception when sqlstate 'T5001' then null;
  end;

  -- A late unexpected audit failure occurs after all items and normal
  -- completion have been attempted. Only the generic failed proposal remains.
  begin
    update public.financial_reconciliation_automatic_proposals
    set status = 'deselected' where id = v_source_proposal_id;
    perform pg_temp.task5_clone_grouped_proposal(
      v_source_proposal_id,
      'f5300000-0000-0000-0000-000000000001',
      'f5310000-0000-0000-0000-000000000001',
      'smoke:task5-unexpected', ''
    );
    create trigger task5_secret_audit_failure
      before insert on public.financial_reconciliation_audit
      for each row execute function pg_temp.task5_secret_audit_failure();
    perform pg_temp.task5_assert_failed(
      'f5310000-0000-0000-0000-000000000001'
    );
    raise sqlstate 'T5001' using message = 'rollback task5 fixture';
  exception when sqlstate 'T5001' then null;
  end;
end
$$;

-- Execute the Bank 10 + 1 group and both Adyen outcomes, then prove exact
-- membership/item identity, automatic provenance, immutable audit evidence,
-- history totals, deterministic forced comment, and retry idempotency.
do $$
declare
  v_bank_proposal_id uuid :=
    (select target.proposal_id from task5_bank_target target);
  v_adyen_zero_proposal_id uuid := (
    select target.proposal_id from task5_adyen_targets target
    where target.grouping_key = '2026-07'
  );
  v_adyen_forced_proposal_id uuid := (
    select target.proposal_id from task5_adyen_targets target
    where target.grouping_key = '2026-06'
  );
  v_bank_result jsonb;
  v_zero_result jsonb;
  v_forced_result jsonb;
  v_retry jsonb;
  v_reconciliation_id uuid;
  v_history jsonb;
  v_history_row jsonb;
  v_items_before bigint;
  v_audit_before bigint;
begin
  v_bank_result := public.execute_financial_reconciliation_automatic_proposal(
    v_bank_proposal_id, 'smoke:task5-bank'
  );
  v_reconciliation_id := (v_bank_result->>'reconciliationId')::uuid;
  if v_bank_result->>'status' <> 'completed'
    or not exists (
      select 1 from public.financial_reconciliations reconciliation
      where reconciliation.id = v_reconciliation_id
        and reconciliation.status = 'complete'
        and reconciliation.completion_type = 'normal'
        and reconciliation.difference_amount = 0
        and reconciliation.base_source_type = 'import_fdm_accounts'
        and reconciliation.origin = 'automatic'
        and reconciliation.automatic_rule_key =
          'fdm_bank_transfer_cgd_bank_statement_combination'
        and reconciliation.automatic_run_id =
          'f5200000-0000-0000-0000-000000000001'
        and reconciliation.automatic_proposal_id = v_bank_proposal_id
    )
    or (select count(*) from public.financial_reconciliation_items item
        where item.reconciliation_id = v_reconciliation_id) <> 11
    or exists (
      select member.source_type, member.source_id, member.amount
      from public.financial_reconciliation_automatic_proposal_memberships member
      where member.proposal_id = v_bank_proposal_id
      except
      select item.source_type, item.source_id, item.amount_snapshot
      from public.financial_reconciliation_items item
      where item.reconciliation_id = v_reconciliation_id
    )
    or exists (
      select item.source_type, item.source_id, item.amount_snapshot
      from public.financial_reconciliation_items item
      where item.reconciliation_id = v_reconciliation_id
      except
      select member.source_type, member.source_id, member.amount
      from public.financial_reconciliation_automatic_proposal_memberships member
      where member.proposal_id = v_bank_proposal_id
    )
    or (select count(*) from public.financial_reconciliation_audit audit
        where audit.reconciliation_id = v_reconciliation_id) <> 13
    or not exists (
      select 1 from public.financial_reconciliation_audit audit
      where audit.reconciliation_id = v_reconciliation_id
        and audit.action = 'automatic_complete'
        and jsonb_array_length(audit.metadata->'membershipSnapshots') = 11
    ) then
    raise exception 'Task 5 Bank execution lost lifecycle or audit state: %.',
      v_bank_result;
  end if;

  select count(*) into v_items_before
  from public.financial_reconciliation_items item
  where item.reconciliation_id = v_reconciliation_id;
  select count(*) into v_audit_before
  from public.financial_reconciliation_audit audit
  where audit.reconciliation_id = v_reconciliation_id;
  v_retry := public.execute_financial_reconciliation_automatic_proposal(
    v_bank_proposal_id, 'smoke:task5-bank'
  );
  if v_retry is distinct from v_bank_result
    or (select count(*) from public.financial_reconciliation_items item
        where item.reconciliation_id = v_reconciliation_id) <> v_items_before
    or (select count(*) from public.financial_reconciliation_audit audit
        where audit.reconciliation_id = v_reconciliation_id) <> v_audit_before then
    raise exception 'Task 5 Bank retry was not idempotent: %.', v_retry;
  end if;

  v_zero_result := public.execute_financial_reconciliation_automatic_proposal(
    v_adyen_zero_proposal_id, 'smoke:task5-adyen'
  );
  v_forced_result := public.execute_financial_reconciliation_automatic_proposal(
    v_adyen_forced_proposal_id, 'smoke:task5-adyen'
  );
  if v_zero_result->>'status' <> 'completed'
    or v_forced_result->>'status' <> 'completed'
    or not exists (
      select 1 from public.financial_reconciliations reconciliation
      where reconciliation.id = (v_zero_result->>'reconciliationId')::uuid
        and reconciliation.completion_type = 'normal'
        and reconciliation.difference_amount = 0
        and reconciliation.forced_completion_comment is null
        and reconciliation.origin = 'automatic'
        and reconciliation.automatic_rule_key =
          'cgd_bank_statement_fdm_adyen_monthly_payments'
        and reconciliation.automatic_run_id =
          'f5500000-0000-0000-0000-000000000001'
        and reconciliation.automatic_proposal_id = v_adyen_zero_proposal_id
    )
    or not exists (
      select 1 from public.financial_reconciliations reconciliation
      where reconciliation.id = (v_forced_result->>'reconciliationId')::uuid
        and reconciliation.completion_type = 'forced'
        and reconciliation.difference_amount = 10
        and reconciliation.forced_completion_comment =
          'FDM Accounts – Adyen Reservation Payments | month 2026-06 | difference 10.00 EUR | allowance 20.00 EUR.'
        and reconciliation.origin = 'automatic'
        and reconciliation.automatic_rule_key =
          'cgd_bank_statement_fdm_adyen_monthly_payments'
        and reconciliation.automatic_run_id =
          'f5500000-0000-0000-0000-000000000001'
        and reconciliation.automatic_proposal_id = v_adyen_forced_proposal_id
    ) then
    raise exception 'Task 5 Adyen normal/forced outcomes diverged: %, %.',
      v_zero_result, v_forced_result;
  end if;

  v_history := public.get_financial_reconciliation_history(
    null, null, 'automatic', 'complete', null, null, 1, 100
  );
  foreach v_reconciliation_id in array array[
    (v_zero_result->>'reconciliationId')::uuid,
    (v_forced_result->>'reconciliationId')::uuid
  ]
  loop
    if exists (
      select member.source_type, member.source_id, member.amount
      from public.financial_reconciliation_automatic_proposal_memberships member
      join public.financial_reconciliation_automatic_proposals proposal
        on proposal.id = member.proposal_id
      where proposal.reconciliation_id = v_reconciliation_id
      except
      select item.source_type, item.source_id, item.amount_snapshot
      from public.financial_reconciliation_items item
      where item.reconciliation_id = v_reconciliation_id
    )
      or exists (
        select item.source_type, item.source_id, item.amount_snapshot
        from public.financial_reconciliation_items item
        where item.reconciliation_id = v_reconciliation_id
        except
        select member.source_type, member.source_id, member.amount
        from public.financial_reconciliation_automatic_proposal_memberships member
        join public.financial_reconciliation_automatic_proposals proposal
          on proposal.id = member.proposal_id
        where proposal.reconciliation_id = v_reconciliation_id
      )
      or not exists (
        select 1
        from public.financial_reconciliation_audit audit
        join public.financial_reconciliation_automatic_proposals proposal
          on proposal.reconciliation_id = audit.reconciliation_id
        where audit.reconciliation_id = v_reconciliation_id
          and audit.action = 'automatic_complete'
          and jsonb_array_length(audit.metadata->'membershipSnapshots') = (
            select count(*)
            from public.financial_reconciliation_automatic_proposal_memberships member
            where member.proposal_id = proposal.id
          )
      ) then
      raise exception 'Task 5 Adyen reconciliation omitted immutable members.';
    end if;

    select history.value into strict v_history_row
    from jsonb_array_elements(v_history->'rows') history(value)
    where history.value->>'id' = v_reconciliation_id::text;
    if (v_history_row->>'totalRecords')::bigint is distinct from (
        select count(*)
        from public.financial_reconciliation_items item
        where item.reconciliation_id = v_reconciliation_id
      )
      or (v_history_row->>'sourceAmountTotal')::numeric is distinct from (
        select sum(member.amount)
        from public.financial_reconciliation_automatic_proposal_memberships member
        join public.financial_reconciliation_automatic_proposals proposal
          on proposal.id = member.proposal_id
        where proposal.reconciliation_id = v_reconciliation_id
          and member.role = 'source'
      )
      or (v_history_row->>'destinationAmountTotal')::numeric is distinct from (
        select sum(member.amount)
        from public.financial_reconciliation_automatic_proposal_memberships member
        join public.financial_reconciliation_automatic_proposals proposal
          on proposal.id = member.proposal_id
        where proposal.reconciliation_id = v_reconciliation_id
          and member.role = 'destination'
      )
      or not coalesce((v_history_row->'sourceSummary') @> (
        select jsonb_agg(jsonb_build_object(
          'sourceType', expected.source_type,
          'recordCount', expected.record_count,
          'amountTotal', expected.amount_total
        ))
        from (
          select member.source_type, count(*)::integer as record_count,
                 sum(member.amount)::numeric as amount_total
          from public.financial_reconciliation_automatic_proposal_memberships member
          join public.financial_reconciliation_automatic_proposals proposal
            on proposal.id = member.proposal_id
          where proposal.reconciliation_id = v_reconciliation_id
          group by member.source_type
        ) expected
      ), false
      ) then
      raise exception 'Task 5 Adyen history lost signed members/totals: %.',
        v_history_row;
    end if;
  end loop;

  select history.value into strict v_history_row
  from jsonb_array_elements(v_history->'rows') history(value)
  where history.value->>'id' =
    (v_bank_result->>'reconciliationId');
  if (v_history_row->>'totalRecords')::integer is distinct from 11
    or (v_history_row->>'sourceAmountTotal')::numeric
      is distinct from -100.00::numeric
    or (v_history_row->>'destinationAmountTotal')::numeric
      is distinct from 100.00::numeric
    or not coalesce((v_history_row->'sourceSummary') @> jsonb_build_array(
      jsonb_build_object(
        'sourceType', 'import_fdm_accounts',
        'recordCount', 10, 'amountTotal', -100.00
      ),
      jsonb_build_object(
        'sourceType', 'import_cgd_extrato_ordem',
        'recordCount', 1, 'amountTotal', 100.00
      )
    ), false) then
    raise exception 'Task 5 Bank history lost signed totals: %.', v_history_row;
  end if;

  select count(*) into v_items_before
  from public.financial_reconciliation_items item
  where item.reconciliation_id = (v_forced_result->>'reconciliationId')::uuid;
  select count(*) into v_audit_before
  from public.financial_reconciliation_audit audit
  where audit.reconciliation_id = (v_forced_result->>'reconciliationId')::uuid;
  v_retry := public.execute_financial_reconciliation_automatic_proposal(
    v_adyen_forced_proposal_id, 'smoke:task5-adyen'
  );
  if v_retry is distinct from v_forced_result
    or (select count(*) from public.financial_reconciliation_items item
        where item.reconciliation_id =
          (v_forced_result->>'reconciliationId')::uuid) <> v_items_before
    or (select count(*) from public.financial_reconciliation_audit audit
        where audit.reconciliation_id =
          (v_forced_result->>'reconciliationId')::uuid) <> v_audit_before then
    raise exception 'Task 5 Adyen retry changed immutable lifecycle state.';
  end if;
end
$$;

-- Task 6 manual catalog and one-rule actor lifecycle
do $$
declare
  v_catalog jsonb;
  v_bank_run jsonb;
  v_bank_retry jsonb;
  v_adyen_run jsonb;
  v_locked_run_id uuid := 'f6000000-0000-0000-0000-000000000001';
begin
  update public.financial_reconciliation_automatic_rule_configs config
  set enabled = true,
      allow_manual_execution = case
        when config.rule_key =
          'fdm_bank_transfer_cgd_bank_statement_combination' then true
        when config.rule_key =
          'cgd_bank_statement_fdm_adyen_monthly_payments' then false
        else config.allow_manual_execution
      end
  where config.rule_key in (
    'fdm_bank_transfer_cgd_bank_statement_combination',
    'cgd_bank_statement_fdm_adyen_monthly_payments'
  );

  v_catalog := public.get_financial_reconciliation_automatic_manual_rules();
  if jsonb_typeof(v_catalog->'rules') is distinct from 'array'
    or not coalesce(v_catalog->'rules' @> jsonb_build_array(
      jsonb_build_object(
        'ruleKey', 'fdm_bank_transfer_cgd_bank_statement_combination',
        'ruleVersion', 1, 'enabled', true,
        'allowManualExecution', true
      )
    ), false)
    or coalesce(v_catalog->'rules' @> jsonb_build_array(
      jsonb_build_object(
        'ruleKey', 'cgd_bank_statement_fdm_adyen_monthly_payments'
      )
    ), false) then
    raise exception 'Task 6 manual catalog did not require enabled and manual flags: %.',
      v_catalog;
  end if;

  update public.financial_reconciliation_automatic_rule_configs config
  set allow_manual_execution = true
  where config.rule_key in (
    'fdm_bank_transfer_cgd_bank_statement_combination',
    'cgd_bank_statement_fdm_adyen_monthly_payments'
  );

  v_bank_run := public.create_financial_reconciliation_automatic_analysis(
    array['fdm_bank_transfer_cgd_bank_statement_combination'],
    'manual_rule', 'smoke:task6-bank',
    'f6000000-0000-0000-0000-000000000002'
  );
  v_bank_retry := public.create_financial_reconciliation_automatic_analysis(
    array['fdm_bank_transfer_cgd_bank_statement_combination'],
    'manual_rule', 'smoke:task6-bank',
    'f6000000-0000-0000-0000-000000000002'
  );
  if v_bank_run->>'runId' is distinct from v_bank_retry->>'runId'
    or v_bank_run->>'trigger' is distinct from 'manual'
    or v_bank_run->>'scope' is distinct from 'rule'
    or jsonb_array_length(v_bank_run->'definitions') is distinct from 1
    or v_bank_run#>>'{definitions,0,ruleKey}' is distinct from
      'fdm_bank_transfer_cgd_bank_statement_combination'
    or (v_bank_run#>>'{definitions,0,ruleVersion}')::integer
      is distinct from 1
    or v_bank_run#>>'{definitions,0,operator}' is distinct from '+' then
    raise exception 'Task 6 Bank manual creation lost its one-rule + snapshot: %.',
      v_bank_run;
  end if;

  v_adyen_run := public.create_financial_reconciliation_automatic_analysis(
    array['cgd_bank_statement_fdm_adyen_monthly_payments'],
    'manual_rule', 'smoke:task6-adyen',
    'f6000000-0000-0000-0000-000000000003'
  );
  if v_adyen_run->>'trigger' is distinct from 'manual'
    or jsonb_array_length(v_adyen_run->'definitions') is distinct from 1
    or v_adyen_run#>>'{definitions,0,ruleKey}' is distinct from
      'cgd_bank_statement_fdm_adyen_monthly_payments'
    or (v_adyen_run#>>'{definitions,0,ruleVersion}')::integer
      is distinct from 1
    or v_adyen_run#>>'{definitions,0,operator}' is distinct from '-' then
    raise exception 'Task 6 Adyen manual creation lost its one-rule - snapshot: %.',
      v_adyen_run;
  end if;

  insert into public.financial_reconciliation_automatic_runs (
    id, trigger, scope, actor, client_request_id, status,
    definition_config_snapshot, analysis_processed, analysis_total
  ) values (
    v_locked_run_id, 'manual', 'rule', 'smoke:task6-locked-actor',
    'f6000000-0000-0000-0000-000000000004', 'analyzing',
    jsonb_build_array(jsonb_build_object(
      'ruleKey', 'fdm_bank_transfer_cgd_bank_statement_combination',
      'ruleVersion', 1, 'priority', 6, 'differenceAllowed', 0,
      'maxDifferenceDays', 3,
      'destinationSourceType', 'import_cgd_extrato_ordem',
      'definition', jsonb_build_object('strategy', 'bounded_exact_combination'),
      'operator', '+'
    )), 0, 1
  );
  begin
    perform public.create_financial_reconciliation_automatic_analysis(
      array['cgd_bank_statement_fdm_adyen_monthly_payments'],
      'manual_rule', 'smoke:task6-locked-actor',
      'f6000000-0000-0000-0000-000000000005'
    );
    raise exception 'Task 6 accepted a second unfinished manual run.';
  exception when others then
    if sqlerrm = 'Task 6 accepted a second unfinished manual run.'
      or sqlerrm not ilike '%unfinished manual run already exists%' then
      raise;
    end if;
  end;
end
$$;

-- Task 6 grouped member paging exposes complete public snapshots
do $$
declare
  v_proposal record;
  v_page jsonb;
begin
  select proposal.* into strict v_proposal
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.run_id = 'f5200000-0000-0000-0000-000000000001'
  order by proposal.created_at, proposal.id
  limit 1;
  v_page := public.get_financial_reconciliation_automatic_proposal_members(
    v_proposal.run_id, v_proposal.id, 'source', 0, 50,
    'smoke:task5-bank'
  );
  if v_page->>'runId' is distinct from v_proposal.run_id::text
    or v_page->>'proposalId' is distinct from v_proposal.id::text
    or v_page->>'ruleKey' is distinct from v_proposal.rule_key
    or (v_page->>'ruleVersion')::integer is distinct from v_proposal.rule_version
    or v_page->>'groupingKey' is distinct from v_proposal.grouping_key
    or v_page->'summarySnapshot' is distinct from v_proposal.summary_snapshot
    or (v_page->>'sourceCount')::integer is distinct from (
      select count(*)::integer
      from public.financial_reconciliation_automatic_proposal_memberships member
      where member.proposal_id = v_proposal.id and member.role = 'source'
    )
    or (v_page->>'sourceTotal')::numeric is distinct from (
      select sum(member.amount)
      from public.financial_reconciliation_automatic_proposal_memberships member
      where member.proposal_id = v_proposal.id and member.role = 'source'
    )
    or (v_page->>'totalCount')::integer is distinct from (
      select count(*)::integer
      from public.financial_reconciliation_automatic_proposal_memberships member
      where member.proposal_id = v_proposal.id and member.role = 'source'
    )
    or jsonb_typeof(v_page->'members') is distinct from 'array'
    or jsonb_array_length(v_page->'members') < 1
    or not (v_page->'members'->0 ?& array[
      'role', 'sourceType', 'sourceId', 'ordinal', 'sourceDate',
      'amount', 'description', 'account', 'rowSnapshot'
    ])
    or jsonb_typeof(v_page#>'{members,0,rowSnapshot}')
      is distinct from 'object' then
    raise exception 'Task 6 Bank member page lost public grouped evidence: %.',
      v_page;
  end if;

  select proposal.* into strict v_proposal
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.run_id = 'f5500000-0000-0000-0000-000000000001'
  order by proposal.created_at, proposal.id
  limit 1;
  v_page := public.get_financial_reconciliation_automatic_proposal_members(
    v_proposal.run_id, v_proposal.id, 'destination', 0, 50,
    'smoke:task5-adyen'
  );
  if v_page->>'ruleKey' is distinct from
      'cgd_bank_statement_fdm_adyen_monthly_payments'
    or v_page->>'groupingKey' is distinct from v_proposal.grouping_key
    or v_page->'destinationCount' is distinct from
      v_proposal.summary_snapshot->'destinationCount'
    or v_page->'destinationTotal' is distinct from
      v_proposal.summary_snapshot->'destinationTotal'
    or jsonb_array_length(v_page->'members') < 1
    or v_page#>>'{members,0,role}' is distinct from 'destination'
    or v_page#>>'{members,0,account}' is distinct from 'Adyen' then
    raise exception 'Task 6 Adyen member page lost public grouped evidence: %.',
      v_page;
  end if;
end
$$;

-- Task 6 seven-child priority snapshot and strategy totals
-- Task 6 same-slot retry and oldest cross-midnight child
-- Task 6 terminal child failure continues and aggregates all seven children
do $$
declare
  v_expected_keys text[] := array[
    'financial_documents_cgd_bank_statement',
    'financial_documents_cgd_credit_card',
    'financial_documents_cgd_bank_statement_amount_only',
    'financial_documents_cgd_credit_card_amount_only',
    'cgd_bank_statement_fdm_credit_card_monthly_income',
    'fdm_bank_transfer_cgd_bank_statement_combination',
    'cgd_bank_statement_fdm_adyen_monthly_payments'
  ];
  v_expected_versions integer[] := array[2, 1, 1, 1, 2, 1, 1];
  v_claim jsonb;
  v_retry jsonb;
  v_cross_midnight_retry jsonb;
  v_batch jsonb;
  v_batch_id uuid;
  v_run_id uuid;
  v_position integer;
  v_expected_total bigint;
begin
  update public.financial_reconciliation_automatic_batches batch
  set status = case when batch.status in ('pending', 'running')
      then 'completed' else batch.status end,
      finished_at = case when batch.status in ('pending', 'running')
        then coalesce(batch.finished_at, now()) else batch.finished_at end;
  update public.financial_reconciliation_automatic_schedule
  set enabled = true, time_of_day = '00:00', time_zone = 'Europe/Lisbon'
  where id = true;
  update public.financial_reconciliation_automatic_rule_configs config
  set enabled = true,
      include_in_scheduled_batch = true,
      priority = case config.rule_key
        when 'financial_documents_cgd_bank_statement' then 1
        when 'financial_documents_cgd_credit_card' then 2
        when 'financial_documents_cgd_bank_statement_amount_only' then 3
        when 'financial_documents_cgd_credit_card_amount_only' then 4
        when 'cgd_bank_statement_fdm_credit_card_monthly_income' then 5
        when 'fdm_bank_transfer_cgd_bank_statement_combination' then 6
        when 'cgd_bank_statement_fdm_adyen_monthly_payments' then 7
      end
  where config.rule_key = any(v_expected_keys);

  v_claim := public.claim_financial_reconciliation_automatic_schedule(
    '2099-01-01 02:00:00+00', 'system:reconciliation'
  );
  v_batch_id := (v_claim->>'batchId')::uuid;
  v_run_id := (v_claim#>>'{run,runId}')::uuid;
  v_retry := public.claim_financial_reconciliation_automatic_schedule(
    '2099-01-01 03:00:00+00', 'system:reconciliation'
  );
  v_cross_midnight_retry :=
    public.claim_financial_reconciliation_automatic_schedule(
    '2099-01-02 02:00:00+00', 'system:reconciliation'
  );
  if v_claim->>'claimed' is distinct from 'true'
    or v_claim->>'resumed' is distinct from 'false'
    or v_retry->>'claimed' is distinct from 'true'
    or v_retry->>'resumed' is distinct from 'true'
    or v_retry->>'batchId' is distinct from v_batch_id::text
    or v_retry#>>'{run,runId}' is distinct from v_run_id::text
    or (v_retry->>'batchRulePosition')::integer is distinct from 1
    or (v_retry->>'batchRuleCount')::integer is distinct from 7 then
    raise exception 'Task 6 same-slot retry changed batch child: %, %.',
      v_claim, v_retry;
  end if;
  if v_cross_midnight_retry->>'claimed' is distinct from 'true'
    or v_cross_midnight_retry->>'resumed' is distinct from 'true'
    or v_cross_midnight_retry->>'batchId' is distinct from v_batch_id::text
    or v_cross_midnight_retry#>>'{run,runId}' is distinct from v_run_id::text
    or (v_cross_midnight_retry->>'batchRulePosition')::integer
      is distinct from 1
    or (v_cross_midnight_retry->>'batchRuleCount')::integer
      is distinct from 7 then
    raise exception 'Task 6 cross-midnight retry changed batch child: %, %.',
      v_claim, v_cross_midnight_retry;
  end if;

  for v_position in 1..7 loop
    if v_position > 1 then
      v_claim := public.claim_financial_reconciliation_automatic_schedule(
        ('2099-01-' || lpad((v_position + 1)::text, 2, '0') ||
          ' 02:00:00+00')::timestamptz,
        'system:reconciliation'
      );
      v_run_id := (v_claim#>>'{run,runId}')::uuid;
    end if;
    if v_claim->>'claimed' is distinct from 'true'
      or (v_claim->>'batchRulePosition')::integer is distinct from v_position
      or (v_claim->>'batchRuleCount')::integer is distinct from 7
      or v_claim#>>'{run,batchRuleKey}' is distinct from
        v_expected_keys[v_position]
      or v_claim#>>'{run,definitions,0,ruleKey}' is distinct from
        v_expected_keys[v_position]
      or (v_claim#>>'{run,definitions,0,ruleVersion}')::integer
        is distinct from v_expected_versions[v_position]
      or (v_claim#>>'{run,batchRuleCount}')::integer is distinct from 7 then
      raise exception 'Task 6 scheduled child order/tuple diverged at %: %.',
        v_position, v_claim;
    end if;
    if v_position = 6 then
      select public.financial_reconciliation_automatic_bank_reservation_count()
      into v_expected_total;
      if (v_claim#>>'{run,analysisTotal}')::bigint is distinct from v_expected_total
        or v_claim#>>'{run,analysisUnit}' is distinct from 'bank_anchors' then
        raise exception 'Task 6 Bank scheduled progress is not Bank anchors: %.',
          v_claim;
      end if;
    elsif v_position = 7 then
      select public.financial_reconciliation_automatic_adyen_month_count()
      into v_expected_total;
      if (v_claim#>>'{run,analysisTotal}')::bigint is distinct from v_expected_total
        or v_claim#>>'{run,analysisUnit}' is distinct from 'calendar_months' then
        raise exception 'Task 6 Adyen scheduled progress is not calendar months: %.',
          v_claim;
      end if;
    end if;

    update public.financial_reconciliation_automatic_runs run
    set status = case when v_position = 3 then 'failed' else 'completed' end,
        analysis_processed = analysis_total,
        analysis_cursor_date = case
          when v_position = 6 and analysis_total > 0 then date '2099-01-06'
          when v_position = 7 and analysis_total > 0 then date '2099-01-01'
          else analysis_cursor_date
        end,
        analysis_cursor_id = case
          when v_position = 6 and analysis_total > 0
            then 'f6000000-0000-0000-0000-000000000006'::uuid
          when v_position = 7 and analysis_total > 0
            then 'f6000000-0000-0000-0000-000000000007'::uuid
          else analysis_cursor_id
        end,
        analysis_completed_at = case when v_position = 3 then null else now() end,
        analysis_error_code = case when v_position = 3
          then 'analysis_continuation_failed' else null end,
        analysis_error_at = case when v_position = 3 then now() else null end,
        finished_at = now(),
        counts = jsonb_build_object(
          'bases', v_position,
          'proposed', v_position,
          'ambiguous', 0,
          'skipped', 0,
          'deselected', 0,
          'completed', v_position,
          'stale', 0,
          'failed', case when v_position = 3 then 1 else 0 end
        )
    where run.id = v_run_id;
  end loop;

  v_batch := public.financial_reconciliation_refresh_automatic_batch(v_batch_id);
  if v_batch->>'status' is distinct from 'partial'
    or (v_batch->>'ruleCount')::integer is distinct from 7
    or (v_batch->>'childCount')::integer is distinct from 7
    or (v_batch#>>'{counts,completedChildren}')::integer is distinct from 6
    or (v_batch#>>'{counts,failedChildren}')::integer is distinct from 1
    or (v_batch#>>'{counts,unfinishedChildren}')::integer is distinct from 0
    or (v_batch#>>'{counts,bases}')::integer is distinct from 28
    or (v_batch#>>'{counts,completed}')::integer is distinct from 28
    or (v_batch#>>'{counts,failed}')::integer is distinct from 1 then
    raise exception 'Task 6 terminal seven-child aggregate is incomplete: %.',
      v_batch;
  end if;
  v_claim := public.claim_financial_reconciliation_automatic_schedule(
    '2099-01-01 23:00:00+00', 'system:reconciliation'
  );
  if v_claim->>'claimed' is distinct from 'false'
    or v_claim->>'reason' is distinct from 'batch_complete'
    or v_claim->>'batchId' is distinct from v_batch_id::text then
    raise exception 'Task 6 terminal batch retry did not remain idempotent: %.',
      v_claim;
  end if;
end
$$;

-- Task 6 malformed batch metadata and progress fail closed
do $$
declare
  v_batch_id uuid;
  v_run_id uuid;
  v_duplicate_id uuid := 'f6000000-0000-0000-0000-000000000099';
begin
  select batch.id into strict v_batch_id
  from public.financial_reconciliation_automatic_batches batch
  where batch.scheduled_slot = '2099-01-01';
  select run.id into strict v_run_id
  from public.financial_reconciliation_automatic_runs run
  where run.batch_id = v_batch_id and run.batch_rule_position = 1;

  begin
    update public.financial_reconciliation_automatic_batches batch
    set rule_snapshot = jsonb_set(
      batch.rule_snapshot, '{0,ruleVersion}', '99'::jsonb
    )
    where batch.id = v_batch_id;
    perform public.financial_reconciliation_refresh_automatic_batch(v_batch_id);
    raise exception 'Task 6 malformed rule tuple was accepted.';
  exception when others then
    if sqlerrm = 'Task 6 malformed rule tuple was accepted.'
      or sqlerrm not ilike '%batch snapshot is invalid%' then
      raise;
    end if;
  end;

  begin
    insert into public.financial_reconciliation_automatic_runs (
      id, trigger, scope, actor, scheduled_slot, status,
      definition_config_snapshot, analysis_processed, analysis_total,
      batch_id, batch_rule_key, batch_rule_position, batch_rule_count
    )
    select v_duplicate_id, run.trigger, run.scope, run.actor,
      run.scheduled_slot, run.status, run.definition_config_snapshot,
      run.analysis_processed, run.analysis_total, run.batch_id,
      run.batch_rule_key, run.batch_rule_position, run.batch_rule_count
    from public.financial_reconciliation_automatic_runs run
    where run.id = v_run_id;
    raise exception 'Task 6 duplicate batch position was accepted.';
  exception when unique_violation then
    null;
  end;

  begin
    update public.financial_reconciliation_automatic_runs run
    set batch_rule_count = 6
    where run.id = v_run_id;
    perform public.financial_reconciliation_refresh_automatic_batch(v_batch_id);
    raise exception 'Task 6 wrong child rule count was accepted.';
  exception when others then
    if sqlerrm = 'Task 6 wrong child rule count was accepted.'
      or sqlerrm not ilike '%child metadata is invalid%' then
      raise;
    end if;
  end;

  begin
    update public.financial_reconciliation_automatic_runs run
    set analysis_processed = analysis_total + 1
    where run.id = v_run_id;
    perform public.financial_reconciliation_refresh_automatic_batch(v_batch_id);
    raise exception 'Task 6 invalid child progress was accepted.';
  exception when others then
    if sqlerrm = 'Task 6 invalid child progress was accepted.'
      or sqlerrm not ilike '%batch progress is invalid%' then
      raise;
    end if;
  end;
end
$$;

-- Task 6 terminal manual retry and request rule binding
do $$
declare
  v_client_request_id uuid := 'f6100000-0000-0000-0000-000000000001';
  v_first jsonb;
  v_retry jsonb;
  v_run_id uuid;
  v_error_at timestamptz := '2026-08-23 14:00:00+00';
begin
  update public.financial_reconciliation_automatic_rule_configs config
  set enabled = true, allow_manual_execution = true
  where config.rule_key in (
    'fdm_bank_transfer_cgd_bank_statement_combination',
    'cgd_bank_statement_fdm_adyen_monthly_payments'
  );

  v_first := public.create_financial_reconciliation_automatic_analysis(
    array['fdm_bank_transfer_cgd_bank_statement_combination'],
    'manual_rule', 'smoke:task6-terminal-retry', v_client_request_id
  );
  v_run_id := (v_first->>'runId')::uuid;
  update public.financial_reconciliation_automatic_runs run
  set status = 'failed',
      analysis_processed = 0,
      analysis_cursor_date = null,
      analysis_cursor_id = null,
      analysis_completed_at = null,
      analysis_error_code = 'analysis_continuation_failed',
      analysis_error_at = v_error_at,
      error_summary = 'Automatic analysis could not be completed.',
      finished_at = v_error_at,
      updated_at = v_error_at
  where run.id = v_run_id;

  v_retry := public.create_financial_reconciliation_automatic_analysis(
    array['fdm_bank_transfer_cgd_bank_statement_combination'],
    'manual_rule', 'smoke:task6-terminal-retry', v_client_request_id
  );
  if v_retry->>'runId' is distinct from v_run_id::text
    or v_retry->>'status' is distinct from 'failed'
    or v_retry->>'analysisComplete' is distinct from 'false'
    or v_retry->>'analysisErrorCode' is distinct from
      'analysis_continuation_failed'
    or (v_retry->>'analysisErrorAt')::timestamptz is distinct from v_error_at
    or (v_retry->>'finishedAt')::timestamptz is distinct from v_error_at
    or (select run.updated_at from public.financial_reconciliation_automatic_runs run
        where run.id = v_run_id) is distinct from v_error_at then
    raise exception 'Task 6 terminal manual retry was not authoritative/idempotent: %.',
      v_retry;
  end if;

  begin
    perform public.create_financial_reconciliation_automatic_analysis(
      array['cgd_bank_statement_fdm_adyen_monthly_payments'],
      'manual_rule', 'smoke:task6-terminal-retry', v_client_request_id
    );
    raise exception 'Task 6 reused one client request for another rule.';
  exception when others then
    if sqlerrm = 'Task 6 reused one client request for another rule.'
      or sqlerrm not ilike '%client request ID is already bound to another automatic rule%' then
      raise;
    end if;
  end;
end
$$;

-- Adyen v2 excludes TransferOutToAccount and NULL destination categories
-- during both analysis and execution-time live membership revalidation.
\ir ../supabase-migrations/2026-08-24-financial-reconciliation-automation-adyen-category-exclusion.sql
\ir ../supabase-migrations/2026-08-24-financial-reconciliation-automation-adyen-category-exclusion.sql
\ir ../supabase-migrations/2026-08-24-financial-reconciliation-automation-adyen-v2-execution-validator-fix.sql
\ir ../supabase-migrations/2026-08-24-financial-reconciliation-automation-adyen-v2-execution-validator-fix.sql

do $$
declare
  v_actor text := 'smoke:adyen-v2-category-exclusion';
  v_result jsonb;
  v_run_id uuid;
  v_proposal_id uuid;
  v_attempt integer := 0;
begin
  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_definitions definition
    where definition.rule_key = 'cgd_bank_statement_fdm_adyen_monthly_payments'
      and definition.version = 2
      and definition.definition = jsonb_build_object(
        'strategy', 'closed_calendar_month',
        'bankDescriptionContains', 'Adyen',
        'fdmAccount', 'Adyen',
        'fdmExcludedCategory', 'TransferOutToAccount',
        'requiresBothSides', true,
        'monthMarkerDays', 31
      )
  ) or not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs config
    where config.rule_key = 'cgd_bank_statement_fdm_adyen_monthly_payments'
      and config.rule_version = 2
  ) then
    raise exception 'Adyen v2 managed definition/config was not installed.';
  end if;

  update public.import_cgd_extrato_ordem
  set descritivo = 'non-Adyen prior smoke row'
  where descritivo ilike '%Adyen%';
  update public.import_fdm_accounts
  set category = 'TransferOutToAccount'
  where account = 'Adyen';

  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  ) values (
    'a2400000-0000-0000-0000-000000000001',
    'smoke-adyen-v2-category', 'adyen-v2-april-bank', date '2026-04-15',
    'Adyen April settlement', 50.00
  );
  insert into public.import_fdm_accounts (
    id, import_batch, source_row_number, account, date_time_raw, event_date,
    category, amount, description
  ) values
    (
      'a2410000-0000-0000-0000-000000000001',
      'smoke-adyen-v2-category', 1, 'Adyen', '2026-04-15', date '2026-04-15',
      'Reservation', 50.00, 'eligible Adyen destination'
    ),
    (
      'a2410000-0000-0000-0000-000000000002',
      'smoke-adyen-v2-category', 2, 'Adyen', '2026-04-16', date '2026-04-16',
      'TransferOutToAccount', 99.00, 'excluded transfer-out destination'
    ),
    (
      'a2410000-0000-0000-0000-000000000003',
      'smoke-adyen-v2-category', 3, 'Adyen', '2026-04-17', date '2026-04-17',
      null, 77.00, 'excluded null-category destination'
    ),
    (
      'a2410000-0000-0000-0000-000000000006',
      'smoke-adyen-v2-category', 6, 'Adyen', '2026-05-17', date '2026-05-17',
      'TransferOutToAccount', 88.00,
      'excluded-only month must not affect analysis total'
    );

  update public.financial_reconciliation_automatic_rule_configs
  set enabled = true,
      allow_manual_execution = true,
      difference_allowed = 0,
      max_difference_days = 31,
      updated_at = now()
  where rule_key = 'cgd_bank_statement_fdm_adyen_monthly_payments';

  v_result := public.create_financial_reconciliation_automatic_analysis(
    array['cgd_bank_statement_fdm_adyen_monthly_payments'],
    'manual_rule', v_actor, 'a2420000-0000-0000-0000-000000000001'
  );
  v_run_id := (v_result->>'runId')::uuid;
  while v_result->>'status' = 'analyzing' loop
    v_attempt := v_attempt + 1;
    if v_attempt > 10 then
      raise exception 'Adyen v2 analysis did not terminate: %.', v_result;
    end if;
    v_result := public.continue_financial_reconciliation_automatic_analysis(
      v_run_id, v_actor
    );
  end loop;

  select proposal.id into strict v_proposal_id
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.run_id = v_run_id
    and proposal.rule_key = 'cgd_bank_statement_fdm_adyen_monthly_payments'
    and proposal.rule_version = 2
    and proposal.grouping_key = '2026-04'
    and proposal.status = 'proposed';

  if exists (
    select 1
    from public.financial_reconciliation_automatic_proposals proposal
    where proposal.run_id = v_run_id
      and proposal.grouping_key = '2026-05'
  ) then
    raise exception 'Adyen v2 analyzed an excluded-only calendar month.';
  end if;

  if (select count(*)
      from public.financial_reconciliation_automatic_proposal_memberships member
      where member.proposal_id = v_proposal_id
        and member.source_type = 'import_fdm_accounts') <> 1
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_proposal_memberships member
      where member.proposal_id = v_proposal_id
        and member.source_id = 'a2410000-0000-0000-0000-000000000001'
    )
    or exists (
      select 1
      from public.financial_reconciliation_automatic_proposal_memberships member
      where member.proposal_id = v_proposal_id
        and member.source_id in (
          'a2410000-0000-0000-0000-000000000002',
          'a2410000-0000-0000-0000-000000000003'
        )
    ) then
    raise exception 'Adyen v2 analysis included an ineligible FDM category.';
  end if;

  -- Rows appearing after analysis must also be ignored by live execution.
  insert into public.import_fdm_accounts (
    id, import_batch, source_row_number, account, date_time_raw, event_date,
    category, amount, description
  ) values
    (
      'a2410000-0000-0000-0000-000000000004',
      'smoke-adyen-v2-category', 4, 'Adyen', '2026-04-18', date '2026-04-18',
      'TransferOutToAccount', 20.00, 'late excluded transfer-out destination'
    ),
    (
      'a2410000-0000-0000-0000-000000000005',
      'smoke-adyen-v2-category', 5, 'Adyen', '2026-04-19', date '2026-04-19',
      null, 20.00, 'late excluded null-category destination'
    );

  v_result := public.execute_financial_reconciliation_automatic_proposal(
    v_proposal_id, v_actor
  );
  if v_result->>'status' is distinct from 'completed' then
    raise exception 'Adyen v2 execution did not ignore excluded live rows: %.',
      v_result;
  end if;
end
$$;

-- Bank Reservation v2 uses FDM minus Bank and both source directions stay '-'.
-- Completed v1 details remain readable and an unfinished v1 scheduled child
-- terminalizes its parent batch instead of stranding the schedule.
insert into public.financial_reconciliation_automatic_batches (
  id, scheduled_slot, actor, status, rule_snapshot
) values (
  'b2430000-0000-0000-0000-000000000001', '2099-12-30',
  'smoke-bank-reservation-v1-batch', 'pending',
  (select run.definition_config_snapshot
   from public.financial_reconciliation_automatic_runs run
   where run.id = 'f5200000-0000-0000-0000-000000000001')
);
insert into public.financial_reconciliation_automatic_runs (
  id, trigger, scope, status, actor, scheduled_slot, batch_id,
  batch_rule_key, batch_rule_position, batch_rule_count,
  definition_config_snapshot, analysis_processed, analysis_total
) values (
  'b2430000-0000-0000-0000-000000000002', 'scheduled', 'rule',
  'analyzing', 'smoke-bank-reservation-v1-batch', '2099-12-30',
  'b2430000-0000-0000-0000-000000000001',
  'fdm_bank_transfer_cgd_bank_statement_combination', 1, 1,
  (select run.definition_config_snapshot
   from public.financial_reconciliation_automatic_runs run
   where run.id = 'f5200000-0000-0000-0000-000000000001'),
  0, 1
);
\ir ../supabase-migrations/2026-08-24-financial-reconciliation-automation-bank-reservation-minus.sql
\ir ../supabase-migrations/2026-08-24-financial-reconciliation-automation-bank-reservation-minus.sql
\ir ../supabase-migrations/2026-08-24-financial-reconciliation-automation-completed-overlap-fix.sql
\ir ../supabase-migrations/2026-08-24-financial-reconciliation-automation-completed-overlap-fix.sql
\ir ../supabase-migrations/2026-08-24-financial-reconciliation-automation-adyen-v2-analysis-gate-fix.sql
\ir ../supabase-migrations/2026-08-24-financial-reconciliation-automation-adyen-v2-analysis-gate-fix.sql

-- The manual analysis gate must accept the immutable Adyen v2 definition,
-- including its excluded FDM category, after all shared rule upgrades run.
do $$
declare
  v_result jsonb;
begin
  update public.financial_reconciliation_automatic_rule_configs config
  set enabled = true,
      allow_manual_execution = true,
      updated_at = now()
  where config.rule_key =
    'cgd_bank_statement_fdm_adyen_monthly_payments';

  v_result := public.create_financial_reconciliation_automatic_analysis(
    array['cgd_bank_statement_fdm_adyen_monthly_payments'],
    'manual_rule', 'smoke:adyen-v2-analysis-gate',
    'a2420000-0000-0000-0000-000000000001'
  );

  if v_result#>>'{definitions,0,ruleKey}' is distinct from
      'cgd_bank_statement_fdm_adyen_monthly_payments'
    or v_result#>>'{definitions,0,ruleVersion}' is distinct from '2'
    or v_result#>>'{definitions,0,operator}' is distinct from '-'
    or v_result#>>'{definitions,0,definition,fdmExcludedCategory}'
      is distinct from 'TransferOutToAccount' then
    raise exception 'Adyen v2 manual analysis gate rejected its managed definition: %.',
      v_result;
  end if;
end
$$;

-- Completed automatic proposal history does not block unlocked records
do $$
declare
  v_actor text := 'smoke-bank-reservation-v2@example.com';
  v_result jsonb;
  v_run_id uuid;
  v_proposal_id uuid;
  v_reconciliation_id uuid;
  v_second_run_id uuid;
  v_second_proposal_id uuid;
  v_second_reconciliation_id uuid;
  v_attempt integer := 0;
  v_completed_proposal_snapshot jsonb;
  v_historical jsonb;
  v_historical_page jsonb;
  v_historical_proposal_id uuid :=
    (select target.proposal_id from task5_bank_target target);
begin
  if (select count(*)
      from public.financial_reconciliation_source_rules source_rule
      where source_rule.base_source_type = 'import_cgd_extrato_ordem'
        and source_rule.matching_source_type = 'import_fdm_accounts'
        and source_rule.operator = '-') <> 1
    or (select count(*)
      from public.financial_reconciliation_source_rules source_rule
      where source_rule.base_source_type = 'import_fdm_accounts'
        and source_rule.matching_source_type = 'import_cgd_extrato_ordem'
        and source_rule.operator = '-') <> 1 then
    raise exception 'Bank/FDM managed source-rule directions are not both subtraction.';
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs config
    where config.rule_key =
      'fdm_bank_transfer_cgd_bank_statement_combination'
      and config.rule_version = 2
  ) then
    raise exception 'Bank Reservation v2 managed configuration is missing.';
  end if;

  v_historical := public.get_financial_reconciliation_automatic_run(
    'f5200000-0000-0000-0000-000000000001'
  );
  if v_historical#>>'{definitions,0,ruleVersion}' is distinct from '1'
    or v_historical#>>'{definitions,0,operator}' is distinct from '+'
    or v_historical->>'finishedAt' is null
    or not exists (
      select 1
      from jsonb_array_elements(v_historical->'proposals') proposal(value)
      where proposal.value->>'id' = v_historical_proposal_id::text
        and proposal.value->>'ruleVersion' = '1'
        and proposal.value->>'status' = 'completed'
    ) then
    raise exception 'Completed Bank Reservation v1 audit details were not preserved: %.',
      v_historical;
  end if;
  v_historical_page :=
    public.get_financial_reconciliation_automatic_proposal_members(
      'f5200000-0000-0000-0000-000000000001',
      v_historical_proposal_id, 'source', 0, 50, 'historical-reader'
    );
  if v_historical_page->>'ruleVersion' is distinct from '1'
    or jsonb_array_length(v_historical_page->'members') < 1 then
    raise exception 'Completed Bank Reservation v1 member evidence is unreadable: %.',
      v_historical_page;
  end if;
  begin
    update public.financial_reconciliation_automatic_runs run
    set status = 'ready', finished_at = null
    where run.id = 'f5200000-0000-0000-0000-000000000001';
    perform public.get_financial_reconciliation_automatic_proposal_members(
      'f5200000-0000-0000-0000-000000000001',
      v_historical_proposal_id, 'source', 0, 50, 'smoke:task5-bank'
    );
    raise exception 'Unfinished Bank Reservation v1 member paging was accepted.';
  exception when others then
    if sqlerrm = 'Unfinished Bank Reservation v1 member paging was accepted.'
      or sqlerrm not ilike
        '%Historical Bank Reservation proposal members require a finished run%' then
      raise;
    end if;
  end;
  if not exists (
    select 1
    from public.financial_reconciliation_automatic_runs run
    where run.id = 'b2430000-0000-0000-0000-000000000002'
      and run.status = 'failed'
      and run.finished_at is not null
      and run.analysis_error_code = 'rule_version_changed'
  ) or not exists (
    select 1
    from public.financial_reconciliation_automatic_batches batch
    where batch.id = 'b2430000-0000-0000-0000-000000000001'
      and batch.status = 'failed'
      and batch.finished_at is not null
  ) then
    raise exception 'Unfinished Bank Reservation v1 scheduled work was stranded.';
  end if;

  begin
    perform public.replace_financial_reconciliation_source_rules((
      select jsonb_agg(jsonb_build_object(
        'base_source_type', source_rule.base_source_type,
        'matching_source_type', source_rule.matching_source_type,
        'operator', case
          when source_rule.base_source_type = 'import_fdm_accounts'
            and source_rule.matching_source_type = 'import_cgd_extrato_ordem'
            then '+'
          else source_rule.operator
        end
      ) order by source_rule.base_source_type, source_rule.matching_source_type)
      from public.financial_reconciliation_source_rules source_rule
    ));
    raise exception 'Managed Bank Reservation source rule accepted operator +.';
  exception when others then
    if sqlerrm = 'Managed Bank Reservation source rule accepted operator +.'
      or sqlerrm not ilike '%managed Bank Reservation source rule%operator -%' then
      raise;
    end if;
  end;

  if not exists (
    select 1 from public.financial_reconciliation_source_rules source_rule
    where source_rule.base_source_type = 'financial_documents'
      and source_rule.matching_source_type = 'import_cgd_cartao_credito'
      and source_rule.operator = '+'
  ) then
    raise exception 'Bank Reservation upgrade changed an unrelated source operator.';
  end if;

  update public.import_fdm_accounts
  set account = 'Bank Transfer prior smoke'
  where account = 'Bank Transfer';

  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  ) values (
    'b2400000-0000-0000-0000-000000000001',
    'smoke-bank-reservation-v2', 'bank-reservation-v2-bank',
    date '2026-12-31', 'Bank Reservation v2 anchor', 98765.43
  );
  insert into public.import_fdm_accounts (
    id, import_batch, source_row_number, account, date_time_raw, event_date,
    category, amount, description
  ) values (
    'b2410000-0000-0000-0000-000000000001',
    'smoke-bank-reservation-v2', 1, 'Bank Transfer', '2026-12-31',
    date '2026-12-31', 'Reservation', 98765.43,
    'Bank Reservation v2 same-sign destination'
  );

  update public.financial_reconciliation_automatic_rule_configs
  set enabled = true,
      allow_manual_execution = true,
      difference_allowed = 0,
      max_difference_days = 0,
      updated_at = now()
  where rule_key = 'fdm_bank_transfer_cgd_bank_statement_combination';

  v_result := public.create_financial_reconciliation_automatic_analysis(
    array['fdm_bank_transfer_cgd_bank_statement_combination'],
    'manual_rule', v_actor, 'b2420000-0000-0000-0000-000000000001'
  );
  v_run_id := (v_result->>'runId')::uuid;
  while v_result->>'status' = 'analyzing' loop
    v_attempt := v_attempt + 1;
    if v_attempt > 1000 then
      raise exception 'Bank Reservation v2 analysis did not terminate: %.',
        v_result;
    end if;
    v_result := public.continue_financial_reconciliation_automatic_analysis(
      v_run_id, v_actor
    );
  end loop;

  select proposal.id into strict v_proposal_id
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.run_id = v_run_id
    and proposal.rule_key =
      'fdm_bank_transfer_cgd_bank_statement_combination'
    and proposal.rule_version = 2
    and proposal.grouping_key =
      'b2400000-0000-0000-0000-000000000001'
    and proposal.status = 'proposed';

  if (select proposal.summary_snapshot->>'operator'
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.id = v_proposal_id) is distinct from '-'
    or (select proposal.summary_snapshot#>>'{candidateGroups,0,equationCents}'
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.id = v_proposal_id) is distinct from '0' then
    raise exception 'Bank Reservation v2 proposal did not snapshot subtraction.';
  end if;

  v_result := public.execute_financial_reconciliation_automatic_proposal(
    v_proposal_id, v_actor
  );
  if v_result->>'status' is distinct from 'completed' then
    raise exception 'Bank Reservation v2 execution did not complete: %.',
      v_result;
  end if;
  v_reconciliation_id := (v_result->>'reconciliationId')::uuid;

  if (select reconciliation.difference_amount
      from public.financial_reconciliations reconciliation
      where reconciliation.id = v_reconciliation_id) is distinct from 0
    or (select reconciliation.matching_source_rules#>>'{0,operator}'
      from public.financial_reconciliations reconciliation
      where reconciliation.id = v_reconciliation_id) is distinct from '-' then
    raise exception 'Bank Reservation v2 persisted the wrong arithmetic.';
  end if;

  select to_jsonb(proposal) - array['reconciliation_id','updated_at']::text[]
  into strict v_completed_proposal_snapshot
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.id = v_proposal_id;

  perform public.finish_financial_reconciliation_automatic_run(v_run_id);
  perform public.financial_reconciliation_action(
    'delete', v_actor, v_reconciliation_id, null, null, null
  );

  if exists (
    select 1
    from public.financial_reconciliation_items item
    where (item.source_type = 'import_cgd_extrato_ordem'
        and item.source_id = 'b2400000-0000-0000-0000-000000000001')
       or (item.source_type = 'import_fdm_accounts'
        and item.source_id = 'b2410000-0000-0000-0000-000000000001')
  ) or not exists (
    select 1
    from public.financial_reconciliation_automatic_proposals proposal
    where proposal.id = v_proposal_id
      and proposal.status = 'completed'
  ) then
    raise exception
      'Deleting the reconciliation did not unlock items while retaining completed audit history.';
  end if;

  v_result := public.create_financial_reconciliation_automatic_analysis(
    array['fdm_bank_transfer_cgd_bank_statement_combination'],
    'manual_rule', v_actor, 'b2420000-0000-0000-0000-000000000002'
  );
  v_second_run_id := (v_result->>'runId')::uuid;
  v_attempt := 0;
  while v_result->>'status' = 'analyzing' loop
    v_attempt := v_attempt + 1;
    if v_attempt > 1000 then
      raise exception
        'Second Bank Reservation v2 analysis did not terminate: %.', v_result;
    end if;
    v_result := public.continue_financial_reconciliation_automatic_analysis(
      v_second_run_id, v_actor
    );
  end loop;

  select proposal.id into strict v_second_proposal_id
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.run_id = v_second_run_id
    and proposal.rule_key =
      'fdm_bank_transfer_cgd_bank_statement_combination'
    and proposal.rule_version = 2
    and proposal.grouping_key =
      'b2400000-0000-0000-0000-000000000001'
    and proposal.status = 'proposed';

  v_result := public.execute_financial_reconciliation_automatic_proposal(
    v_second_proposal_id, v_actor
  );
  if v_result->>'status' is distinct from 'completed' then
    raise exception
      'Completed automatic proposal history still blocks unlocked records: %.',
      v_result;
  end if;
  v_second_reconciliation_id := (v_result->>'reconciliationId')::uuid;

  if v_second_reconciliation_id is null
    or (select to_jsonb(proposal) - array['reconciliation_id','updated_at']::text[]
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.id = v_proposal_id)
      is distinct from v_completed_proposal_snapshot then
    raise exception
      'Re-execution did not preserve the first completed proposal audit snapshot.';
  end if;
end
$$;

rollback;
