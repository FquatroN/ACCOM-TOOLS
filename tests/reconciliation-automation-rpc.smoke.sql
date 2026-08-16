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

do $$
declare
  v_candidate_definition text;
  v_signature text;
begin
  select pg_get_functiondef(
    'public.financial_reconciliation_automatic_candidates_for_base_ids(text,integer,numeric,integer,uuid[])'::regprocedure
  ) into strict v_candidate_definition;

  if v_candidate_definition !~* 'left join lateral\s+\('
    or v_candidate_definition !~* 'financial_reconciliation_cgd_match_search'
    or v_candidate_definition !~* 'd\.payment\s*=\s*''Banco'''
    or v_candidate_definition !~* 'p_rule_version\s*=\s*2' then
    raise exception 'Version 2 candidate function lost indexed Banco semantics.';
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
