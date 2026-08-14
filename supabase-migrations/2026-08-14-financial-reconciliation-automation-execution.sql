alter table public.financial_reconciliation_audit
  drop constraint if exists financial_reconciliation_audit_action_check;

alter table public.financial_reconciliation_audit
  add constraint financial_reconciliation_audit_action_check check (
    action in (
      'start', 'add_item', 'remove_item', 'complete', 'force_complete',
      'reopen', 'delete', 'automatic_complete'
    )
  );

create or replace function public.execute_financial_reconciliation_automatic_proposal(
  p_proposal_id uuid,
  p_actor text
)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
declare
  v_run_id uuid;
  v_run public.financial_reconciliation_automatic_runs%rowtype;
  v_proposal public.financial_reconciliation_automatic_proposals%rowtype;
  v_rule_snapshot jsonb;
  v_rule_snapshot_count integer;
  v_current_definition jsonb;
  v_current_base_source_type text;
  v_current_destination_source_types jsonb;
  v_current_rule_version integer;
  v_current_operator text;
  v_locked_destination_count integer;
  v_base record;
  v_combination record;
  v_combination_count integer;
  v_current_evidence jsonb;
  v_action_result jsonb;
  v_item record;
  v_reconciliation_id uuid;
  v_expected_item_count integer;
  v_actual_item_count integer;
  v_expected_matching_source_rules jsonb;
  v_actual_matching_source_rules jsonb;
  v_actual_difference numeric;
  v_completion_action text;
  v_comment text;
  v_trigger_label text;
  v_failure_message text;
  v_failure_detail text;
begin
  if p_proposal_id is null then
    raise exception 'Automation proposal ID is required.';
  end if;
  if nullif(trim(coalesce(p_actor, '')), '') is null then
    raise exception 'Actor is required.';
  end if;

  select proposal.run_id into v_run_id
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.id = p_proposal_id;
  if not found then
    raise exception 'Automation proposal was not found.';
  end if;

  select * into strict v_run
  from public.financial_reconciliation_automatic_runs
  where id = v_run_id
  for update;

  select * into strict v_proposal
  from public.financial_reconciliation_automatic_proposals
  where id = p_proposal_id
  for update;
  if v_proposal.run_id <> v_run.id then
    raise exception 'Automation proposal run changed during execution.';
  end if;

  if v_proposal.status = 'completed' then
    if v_proposal.reconciliation_id is null then
      raise exception 'Completed automation proposal has no reconciliation.';
    end if;
    return jsonb_build_object(
      'proposalId', v_proposal.id,
      'runId', v_run.id,
      'status', 'completed',
      'reconciliationId', v_proposal.reconciliation_id
    );
  end if;
  if v_proposal.status in ('ambiguous', 'deselected', 'failed') then
    raise exception 'Automation proposal with status % cannot be executed.', v_proposal.status;
  end if;
  if v_proposal.status <> 'proposed' then
    raise exception 'Automation proposal is already being executed.';
  end if;
  if v_run.finished_at is not null then
    raise exception 'Automation proposal belongs to a finished run.';
  end if;
  if v_proposal.rule_key <> 'financial_documents_cgd_bank_statement'
    or v_proposal.rule_version <> 1
    or v_proposal.base_source_type <> 'financial_documents' then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_version_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_version_changed'
    );
  end if;

  select count(*) into v_rule_snapshot_count
  from jsonb_array_elements(v_run.definition_config_snapshot) snapshot(value)
  where value->>'ruleKey' = v_proposal.rule_key
    and (value->>'ruleVersion')::integer = v_proposal.rule_version;
  if v_rule_snapshot_count <> 1 then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_snapshot_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_snapshot_changed'
    );
  end if;
  select value into strict v_rule_snapshot
  from jsonb_array_elements(v_run.definition_config_snapshot) snapshot(value)
  where value->>'ruleKey' = v_proposal.rule_key
    and (value->>'ruleVersion')::integer = v_proposal.rule_version;

  select
    definition.definition,
    definition.base_source_type,
    definition.destination_source_types,
    config.rule_version,
    source_rule.operator
  into
    v_current_definition,
    v_current_base_source_type,
    v_current_destination_source_types,
    v_current_rule_version,
    v_current_operator
  from public.financial_reconciliation_automatic_rule_definitions definition
  join public.financial_reconciliation_automatic_rule_configs config
    on config.rule_key = definition.rule_key
  join public.financial_reconciliation_source_rules source_rule
    on source_rule.base_source_type = definition.base_source_type
   and source_rule.matching_source_type = 'import_cgd_extrato_ordem'
  where definition.rule_key = v_proposal.rule_key
    and definition.version = v_proposal.rule_version
  for share of definition, config, source_rule;

  if not found
    or v_current_rule_version is distinct from v_proposal.rule_version
    or v_current_definition is distinct from v_rule_snapshot->'definition'
    or v_current_base_source_type is distinct from v_proposal.base_source_type
    or v_current_destination_source_types is distinct from '["import_cgd_extrato_ordem"]'::jsonb then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_snapshot_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_snapshot_changed'
    );
  end if;
  if v_current_operator is distinct from v_rule_snapshot->>'operator' then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'operator_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'operator_changed'
    );
  end if;
  if v_proposal.allowed_difference is distinct from (v_rule_snapshot->>'differenceAllowed')::numeric then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'tolerance_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'tolerance_changed'
    );
  end if;

  perform document.id
  from public.financial_documents document
  where document.id = v_proposal.base_source_id
  for update;
  if not found then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'source_snapshot_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'source_snapshot_changed'
    );
  end if;
  if jsonb_array_length(v_proposal.items) = 0
    or exists (
      select 1 from jsonb_array_elements(v_proposal.items) item(value)
      where value->>'sourceType' <> 'import_cgd_extrato_ordem'
        or coalesce(value->>'sourceId', '') !~
          '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
        or coalesce(value->>'sourceDate', '') !~ '^\d{4}-\d{2}-\d{2}$'
    ) then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'source_snapshot_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'source_snapshot_changed'
    );
  end if;
  perform bank.id
  from jsonb_array_elements(v_proposal.items) item(value)
  join public.import_cgd_extrato_ordem bank
    on bank.id = (item.value->>'sourceId')::uuid
  order by item.value->>'sourceType', bank.data, bank.id
  for update of bank;
  get diagnostics v_locked_destination_count = row_count;
  if v_locked_destination_count <> jsonb_array_length(v_proposal.items) then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'source_snapshot_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'source_snapshot_changed'
    );
  end if;

  select * into v_base
  from public.financial_reconciliation_automatic_rule_candidates(
    v_proposal.rule_key,
    v_proposal.rule_version,
    (v_rule_snapshot->>'differenceAllowed')::numeric,
    (v_rule_snapshot->>'maxDifferenceDays')::integer
  ) candidates
  where candidates.base_source_id = v_proposal.base_source_id;
  if not found
    or v_base.base_source_date is distinct from v_proposal.base_source_date
    or v_base.base_snapshot is distinct from v_proposal.base_snapshot
    or v_base.candidate_count > 12 then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'source_snapshot_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'source_snapshot_changed'
    );
  end if;

  select count(*) into v_combination_count
  from public.financial_reconciliation_automatic_build_combinations(
    v_base.base_snapshot,
    v_base.candidates,
    jsonb_build_object('import_cgd_extrato_ordem', v_rule_snapshot->>'operator'),
    (v_rule_snapshot->>'differenceAllowed')::numeric,
    4
  );
  if v_combination_count <> 1 then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'combination_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'combination_changed'
    );
  end if;
  select * into strict v_combination
  from public.financial_reconciliation_automatic_build_combinations(
    v_base.base_snapshot,
    v_base.candidates,
    jsonb_build_object('import_cgd_extrato_ordem', v_rule_snapshot->>'operator'),
    (v_rule_snapshot->>'differenceAllowed')::numeric,
    4
  );
  select coalesce(jsonb_agg(value->'evidence'), '[]'::jsonb)
  into v_current_evidence
  from jsonb_array_elements(v_combination.items) item(value);

  if v_combination.signature is distinct from v_proposal.signature
    or v_combination.items is distinct from v_proposal.items
    or v_current_evidence is distinct from v_proposal.evidence
    or v_combination.calculated_difference is distinct from v_proposal.calculated_difference then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'proposal_evidence_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'proposal_evidence_changed'
    );
  end if;

  begin
    update public.financial_reconciliation_automatic_proposals
    set status = 'executing', reason = '', error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;

    v_action_result := public.financial_reconciliation_action(
      'start', p_actor, null, v_proposal.base_source_type, v_proposal.base_source_id, null
    );
    v_reconciliation_id := (v_action_result#>>'{reconciliation,id}')::uuid;
    if v_reconciliation_id is null then
      raise exception 'Automatic reconciliation start returned no reconciliation.';
    end if;

    for v_item in
      select value
      from jsonb_array_elements(v_proposal.items) item(value)
      order by value->>'sourceType', (value->>'sourceDate')::date, value->>'sourceId'
    loop
      perform public.financial_reconciliation_action(
        'add_item', p_actor, v_reconciliation_id,
        v_item.value->>'sourceType', (v_item.value->>'sourceId')::uuid, null
      );
    end loop;

    v_expected_item_count := 1 + jsonb_array_length(v_proposal.items);
    select count(*) into v_actual_item_count
    from public.financial_reconciliation_items item
    where item.reconciliation_id = v_reconciliation_id;
    if v_actual_item_count <> v_expected_item_count
      or exists (
        select 1
        from public.financial_reconciliation_items locked_item
        where locked_item.reconciliation_id = v_reconciliation_id
          and (
            (
              locked_item.source_type = v_proposal.base_source_type
              and (
                locked_item.source_id <> v_proposal.base_source_id
                or locked_item.amount_snapshot is distinct from (v_base.base_snapshot->>'amount')::numeric
              )
            )
            or
            (
              locked_item.source_type <> v_proposal.base_source_type
              and not exists (
                select 1
                from jsonb_array_elements(v_proposal.items) proposal_item(value)
                where proposal_item.value->>'sourceType' = locked_item.source_type
                  and (proposal_item.value->>'sourceId')::uuid = locked_item.source_id
                  and (proposal_item.value->>'amount')::numeric = locked_item.amount_snapshot
              )
            )
          )
      ) then
      raise exception 'Automatic reconciliation lifecycle snapshots changed after revalidation.';
    end if;

    v_expected_matching_source_rules := jsonb_build_object(
      'sourceType', 'import_cgd_extrato_ordem',
      'operator', v_rule_snapshot->>'operator'
    );
    select matching_rule.value, reconciliation.difference_amount
    into v_actual_matching_source_rules, v_actual_difference
    from public.financial_reconciliations reconciliation
    join lateral jsonb_array_elements(reconciliation.matching_source_rules) matching_rule(value)
      on matching_rule.value->>'sourceType' = 'import_cgd_extrato_ordem'
    where reconciliation.id = v_reconciliation_id;
    if not found
      or v_actual_matching_source_rules is distinct from v_expected_matching_source_rules
      or v_actual_difference is distinct from v_proposal.calculated_difference
      or abs(v_actual_difference) > v_proposal.allowed_difference then
      raise exception 'Automatic reconciliation lifecycle snapshots changed after revalidation.';
    end if;

    update public.financial_reconciliations
    set origin = 'automatic',
        automatic_trigger = v_run.trigger,
        automatic_rule_key = v_proposal.rule_key,
        automatic_rule_version = v_proposal.rule_version,
        automatic_run_id = v_run.id,
        automatic_proposal_id = v_proposal.id,
        updated_at = timezone('utc', now())
    where id = v_reconciliation_id;

    if v_actual_difference = 0 then
      v_completion_action := 'complete';
      v_comment := null;
    else
      v_completion_action := 'force_complete';
      v_trigger_label := case v_run.trigger when 'manual' then 'Manual' else 'Scheduled' end;
      v_comment := 'Automatically completed by rule Financial Documents to CGD Bank Statement v1; difference '
        || chr(8364) || to_char(v_actual_difference, 'FM999999999990.00')
        || ' within allowed tolerance ' || chr(8364)
        || to_char(v_proposal.allowed_difference, 'FM999999999990.00')
        || '; trigger ' || v_trigger_label || '; batch ' || v_run.id::text || '.';
    end if;

    if v_completion_action = 'complete' then
      perform public.financial_reconciliation_action(
        'complete', p_actor, v_reconciliation_id, null, null, null
      );
    else
      perform public.financial_reconciliation_action(
        'force_complete', p_actor, v_reconciliation_id, null, null, v_comment
      );
    end if;

    insert into public.financial_reconciliation_audit (
      reconciliation_id, action, actor, comment, difference_amount, metadata
    ) values (
      v_reconciliation_id,
      'automatic_complete',
      p_actor,
      v_comment,
      v_actual_difference,
      jsonb_build_object(
        'ruleSnapshot', jsonb_build_object(
          'ruleKey', v_proposal.rule_key,
          'ruleVersion', v_proposal.rule_version,
          'definition', v_rule_snapshot->'definition'
        ),
        'configSnapshot', jsonb_build_object(
          'differenceAllowed', (v_rule_snapshot->>'differenceAllowed')::numeric,
          'maxDifferenceDays', (v_rule_snapshot->>'maxDifferenceDays')::integer,
          'priority', (v_rule_snapshot->>'priority')::integer
        ),
        'operatorSnapshot', jsonb_build_object(
          'import_cgd_extrato_ordem', v_rule_snapshot->>'operator'
        ),
        'identityEvidence', v_proposal.evidence,
        'proposalSignature', v_proposal.signature,
        'trigger', v_run.trigger,
        'runId', v_run.id,
        'proposalId', v_proposal.id,
        'tolerance', v_proposal.allowed_difference,
        'calculatedDifference', v_actual_difference
      )
    );

    update public.financial_reconciliation_automatic_proposals
    set status = 'completed',
        reconciliation_id = v_reconciliation_id,
        completed_at = now(),
        reason = '',
        error = '',
        error_detail = '',
        updated_at = now()
    where id = v_proposal.id;
  exception when others then
    get stacked diagnostics
      v_failure_message = message_text,
      v_failure_detail = pg_exception_detail;
    if v_failure_message in (
      'Automatic reconciliation lifecycle snapshots changed after revalidation.',
      'This record is already reconciled.'
    ) then
      update public.financial_reconciliation_automatic_proposals
      set status = 'stale',
          reason = 'source_snapshot_changed',
          reconciliation_id = null,
          completed_at = null,
          error = '',
          error_detail = '',
          updated_at = now()
      where id = v_proposal.id;
      return jsonb_build_object(
        'proposalId', v_proposal.id,
        'runId', v_run.id,
        'status', 'stale',
        'reason', 'source_snapshot_changed'
      );
    end if;
    update public.financial_reconciliation_automatic_proposals
    set status = 'failed',
        reason = 'execution_failed',
        reconciliation_id = null,
        completed_at = null,
        error = 'Automatic reconciliation execution failed.',
        error_detail = left(concat_ws(' ', v_failure_message, nullif(v_failure_detail, '')), 2000),
        updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id,
      'runId', v_run.id,
      'status', 'failed',
      'reason', 'execution_failed'
    );
  end;

  return jsonb_build_object(
    'proposalId', v_proposal.id,
    'runId', v_run.id,
    'status', 'completed',
    'reconciliationId', v_reconciliation_id
  );
end $$;

create or replace function public.finish_financial_reconciliation_automatic_run(
  p_run_id uuid
)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
declare
  v_run public.financial_reconciliation_automatic_runs%rowtype;
  v_counts jsonb;
  v_status text;
  v_completed integer;
  v_stale integer;
  v_failed integer;
begin
  if p_run_id is null then
    raise exception 'Automatic run ID is required.';
  end if;
  select * into v_run
  from public.financial_reconciliation_automatic_runs
  where id = p_run_id
  for update;
  if not found then
    raise exception 'Automatic analysis run was not found.';
  end if;
  if v_run.finished_at is not null then
    return public.get_financial_reconciliation_automatic_run(v_run.id);
  end if;
  if exists (
    select 1 from public.financial_reconciliation_automatic_proposals
    where run_id = v_run.id and status = 'executing'
  ) then
    raise exception 'Automatic run still has an executing proposal.';
  end if;

  update public.financial_reconciliation_automatic_proposals
  set status = 'deselected', reason = 'not_selected', updated_at = now()
  where run_id = v_run.id
    and status = 'proposed';

  select jsonb_build_object(
    'bases', count(distinct base_source_id),
    'proposed', count(*) filter (where status = 'proposed'),
    'ambiguous', count(*) filter (where status = 'ambiguous'),
    'deselected', count(*) filter (where status = 'deselected'),
    'executing', count(*) filter (where status = 'executing'),
    'completed', count(*) filter (where status = 'completed'),
    'stale', count(*) filter (where status = 'stale'),
    'failed', count(*) filter (where status = 'failed')
  ) into v_counts
  from public.financial_reconciliation_automatic_proposals
  where run_id = v_run.id;

  v_completed := coalesce((v_counts->>'completed')::integer, 0);
  v_stale := coalesce((v_counts->>'stale')::integer, 0);
  v_failed := coalesce((v_counts->>'failed')::integer, 0);
  v_status := case
    when v_failed > 0 and v_completed = 0 then 'failed'
    when v_failed > 0 or v_stale > 0 then 'partial'
    else 'completed'
  end;

  update public.financial_reconciliation_automatic_runs
  set status = case
        when v_status = 'partial' then 'partial'
        when v_status = 'failed' then 'failed'
        else 'completed'
      end,
      counts = v_counts,
      finished_at = now(),
      updated_at = now()
  where id = v_run.id;

  return public.get_financial_reconciliation_automatic_run(v_run.id);
end $$;

do $provenance$
declare
  definition text;
  original_definition text;
  old_current constant text := $$to_jsonb(v_rec) || jsonb_build_object('matchingSourceRules',v_rules)$$;
  new_current constant text := $$to_jsonb(v_rec) || jsonb_build_object(
    'matchingSourceRules',v_rules,
    'origin',v_rec.origin,
    'automaticTrigger',v_rec.automatic_trigger,
    'automaticRuleKey',v_rec.automatic_rule_key,
    'automaticRuleVersion',v_rec.automatic_rule_version,
    'automaticRunId',v_rec.automatic_run_id
  )$$;
  old_history constant text := $$to_jsonb(h) || jsonb_build_object('sourceSummary',coalesce(summary.source_summary,'[]'::jsonb))$$;
  new_history constant text := $$to_jsonb(h) || jsonb_build_object(
        'sourceSummary',coalesce(summary.source_summary,'[]'::jsonb),
        'origin',h.origin,
        'automaticTrigger',h.automatic_trigger,
        'automaticRuleKey',h.automatic_rule_key,
        'automaticRuleVersion',h.automatic_rule_version,
        'automaticRunId',h.automatic_run_id
      )$$;
  old_current_count integer;
  new_current_count integer;
  old_history_count integer;
  new_history_count integer;
begin
  select pg_get_functiondef('public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer)'::regprocedure)
  into definition;
  original_definition := definition;

  old_current_count := (length(definition)-length(replace(definition,old_current,''))) / length(old_current);
  new_current_count := (length(definition)-length(replace(definition,new_current,''))) / length(new_current);
  old_history_count := (length(definition)-length(replace(definition,old_history,''))) / length(old_history);
  new_history_count := (length(definition)-length(replace(definition,new_history,''))) / length(new_history);

  if not (
    (old_current_count = 1 and new_current_count = 0 and old_history_count = 1 and new_history_count = 0)
    or
    (old_current_count = 0 and new_current_count = 1 and old_history_count = 0 and new_history_count = 1)
  ) then
    raise exception 'Unexpected reconciliation workspace function definition; could not install automatic provenance.';
  end if;

  if old_current_count = 1 then
    definition := replace(definition,old_current,new_current);
    definition := replace(definition,old_history,new_history);
  end if;

  old_current_count := (length(definition)-length(replace(definition,old_current,''))) / length(old_current);
  new_current_count := (length(definition)-length(replace(definition,new_current,''))) / length(new_current);
  old_history_count := (length(definition)-length(replace(definition,old_history,''))) / length(old_history);
  new_history_count := (length(definition)-length(replace(definition,new_history,''))) / length(new_history);
  if old_current_count <> 0 or new_current_count <> 1
    or old_history_count <> 0 or new_history_count <> 1 then
    raise exception 'Unexpected reconciliation workspace function definition; could not verify automatic provenance.';
  end if;

  if definition <> original_definition then
    execute definition;
  end if;
end $provenance$;

revoke all on function public.execute_financial_reconciliation_automatic_proposal(uuid,text)
  from public, anon, authenticated, service_role;
revoke all on function public.finish_financial_reconciliation_automatic_run(uuid)
  from public, anon, authenticated, service_role;
grant execute on function public.execute_financial_reconciliation_automatic_proposal(uuid,text)
  to service_role;
grant execute on function public.finish_financial_reconciliation_automatic_run(uuid)
  to service_role;

notify pgrst, 'reload schema';
