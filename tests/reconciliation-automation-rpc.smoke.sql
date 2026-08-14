\set ON_ERROR_STOP on

begin;

\ir ../supabase-migrations/2026-08-14-financial-reconciliation-automation-schema.sql

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
  v_rule_version integer;
begin
  insert into public.financial_reconciliation_automatic_rule_definitions (
    rule_key, version, display_name, base_source_type, destination_source_types,
    logic_description, definition
  )
  select rule_key, 2, display_name, base_source_type, destination_source_types,
         logic_description, definition
  from public.financial_reconciliation_automatic_rule_definitions
  where rule_key = 'financial_documents_cgd_bank_statement' and version = 1;

  perform public.replace_financial_reconciliation_automation_settings(
    '{"enabled":true,"time_of_day":"04:30","time_zone":"Europe/Lisbon"}'::jsonb,
    '[{"rule_key":"financial_documents_cgd_bank_statement","rule_version":2,"enabled":false,"allow_manual_execution":false,"include_in_scheduled_batch":false,"difference_allowed":"3.21","max_difference_days":9,"priority":1}]'::jsonb,
    'smoke:managed-version'
  );

  select rule_version into strict v_rule_version
  from public.financial_reconciliation_automatic_rule_configs
  where rule_key = 'financial_documents_cgd_bank_statement';
  if v_rule_version <> 1 then
    raise exception 'Settings PUT changed managed rule identity.';
  end if;
  if not exists (
    select 1 from public.financial_reconciliation_automatic_rule_configs
    where rule_key = 'financial_documents_cgd_bank_statement'
      and difference_allowed = 3.21
      and max_difference_days = 9
      and updated_by = 'smoke:managed-version'
  ) then
    raise exception 'Settings PUT did not update approved editable fields.';
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

rollback;
