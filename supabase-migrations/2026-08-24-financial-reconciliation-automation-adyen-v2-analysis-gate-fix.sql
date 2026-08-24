-- Keep manual automatic-analysis creation aligned with the managed v2 Adyen
-- definition. The earlier dynamic upgrade could leave this gate comparing the
-- immutable v2 row with the pre-category-exclusion v1 JSON shape.

create or replace function public.create_financial_reconciliation_automatic_analysis(
  p_rule_keys text[],
  p_mode text,
  p_actor text,
  p_client_request_id uuid
)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
declare
  v_snapshot jsonb;
  v_run_id uuid;
  v_current_run_id uuid;
  v_current_request_id uuid;
  v_existing_rule_key text;
  v_existing_rule_version integer;
  v_existing_finished_at timestamptz;
  v_existing_snapshot jsonb;
  v_rule_key text;
  v_rule_version integer;
  v_total bigint;
begin
  if nullif(trim(coalesce(p_actor, '')), '') is null then
    raise exception 'Actor is required.';
  end if;
  if p_client_request_id is null then
    raise exception 'Client request ID is required.';
  end if;
  if p_mode is null or p_rule_keys is null
    or p_mode <> 'manual_rule'
    or cardinality(p_rule_keys) <> 1 then
    raise exception 'Manual automatic analysis requires exactly one selected rule.';
  end if;

  perform pg_advisory_xact_lock(
    hashtextextended('financial_reconciliation_automatic_manual:' || p_actor, 0)
  );

  select run.id, run.client_request_id, run.definition_config_snapshot,
         run.definition_config_snapshot#>>'{0,ruleKey}',
         case when coalesce(run.definition_config_snapshot#>>'{0,ruleVersion}', '')
           ~ '^[0-9]+$' then
           (run.definition_config_snapshot#>>'{0,ruleVersion}')::integer
         else null end
  into v_current_run_id, v_current_request_id, v_existing_snapshot,
       v_existing_rule_key, v_existing_rule_version
  from public.financial_reconciliation_automatic_runs run
  where run.actor = p_actor
    and run.trigger = 'manual'
    and run.finished_at is null
  for update;

  if v_current_run_id is not null then
    if v_current_request_id = p_client_request_id then
      if jsonb_typeof(v_existing_snapshot) is distinct from 'array'
        or jsonb_array_length(v_existing_snapshot) <> 1
        or v_existing_rule_key is distinct from p_rule_keys[1]
        or not coalesce((v_existing_rule_key, v_existing_rule_version) in (
          ('financial_documents_cgd_bank_statement', 2),
          ('financial_documents_cgd_credit_card', 1),
          ('financial_documents_cgd_bank_statement_amount_only', 1),
          ('financial_documents_cgd_credit_card_amount_only', 1),
          ('cgd_bank_statement_fdm_credit_card_monthly_income', 2),
          ('fdm_bank_transfer_cgd_bank_statement_combination', 1),
          ('fdm_bank_transfer_cgd_bank_statement_combination', 2),
          ('cgd_bank_statement_fdm_adyen_monthly_payments', 1),
          ('cgd_bank_statement_fdm_adyen_monthly_payments', 2)
        ), false) then
        raise exception 'Client request ID is already bound to another automatic rule.';
      end if;
      return public.continue_financial_reconciliation_automatic_analysis(
        v_current_run_id, p_actor
      );
    end if;
    raise exception 'Automatic analysis conflict: an unfinished manual run already exists for this actor.';
  end if;

  select run.id, run.definition_config_snapshot,
         run.definition_config_snapshot#>>'{0,ruleKey}',
         case when coalesce(run.definition_config_snapshot#>>'{0,ruleVersion}', '')
           ~ '^[0-9]+$' then
           (run.definition_config_snapshot#>>'{0,ruleVersion}')::integer
         else null end,
         run.finished_at
  into v_run_id, v_existing_snapshot, v_existing_rule_key,
       v_existing_rule_version, v_existing_finished_at
  from public.financial_reconciliation_automatic_runs run
  where run.actor = p_actor
    and run.client_request_id = p_client_request_id;

  if v_run_id is not null then
    if jsonb_typeof(v_existing_snapshot) is distinct from 'array'
      or jsonb_array_length(v_existing_snapshot) <> 1
      or v_existing_rule_key is distinct from p_rule_keys[1]
      or not coalesce((v_existing_rule_key, v_existing_rule_version) in (
        ('financial_documents_cgd_bank_statement', 2),
        ('financial_documents_cgd_credit_card', 1),
        ('financial_documents_cgd_bank_statement_amount_only', 1),
        ('financial_documents_cgd_credit_card_amount_only', 1),
        ('cgd_bank_statement_fdm_credit_card_monthly_income', 2),
        ('fdm_bank_transfer_cgd_bank_statement_combination', 1),
        ('fdm_bank_transfer_cgd_bank_statement_combination', 2),
        ('cgd_bank_statement_fdm_adyen_monthly_payments', 1),
        ('cgd_bank_statement_fdm_adyen_monthly_payments', 2)
      ), false) then
      raise exception 'Client request ID is already bound to another automatic rule.';
    end if;
    if v_existing_finished_at is not null then
      return public.get_financial_reconciliation_automatic_run(v_run_id);
    end if;
    return public.continue_financial_reconciliation_automatic_analysis(
      v_run_id, p_actor
    );
  end if;

  lock table public.financial_reconciliation_source_rules
    in share row exclusive mode;
  lock table public.financial_reconciliation_automatic_rule_configs
    in share row exclusive mode;

  select coalesce(jsonb_agg(jsonb_build_object(
    'ruleKey', config.rule_key,
    'ruleVersion', config.rule_version,
    'displayName', definition.display_name,
    'priority', config.priority,
    'differenceAllowed', config.difference_allowed,
    'maxDifferenceDays', config.max_difference_days,
    'destinationSourceType', destination.source_type,
    'definition', definition.definition,
    'operator', source_rule.operator
  )), '[]'::jsonb)
  into v_snapshot
  from public.financial_reconciliation_automatic_rule_configs config
  join public.financial_reconciliation_automatic_rule_definitions definition
    on definition.rule_key = config.rule_key
   and definition.version = config.rule_version
  cross join lateral jsonb_array_elements_text(
    definition.destination_source_types
  ) destination(source_type)
  join public.financial_reconciliation_source_rules source_rule
    on source_rule.base_source_type = definition.base_source_type
   and source_rule.matching_source_type = destination.source_type
  where config.rule_key = p_rule_keys[1]
    and config.enabled
    and config.allow_manual_execution
    and config.difference_allowed >= 0
    and jsonb_array_length(definition.destination_source_types) = 1
    and (config.rule_key, config.rule_version) in (
      ('financial_documents_cgd_bank_statement', 2),
      ('financial_documents_cgd_credit_card', 1),
      ('financial_documents_cgd_bank_statement_amount_only', 1),
      ('financial_documents_cgd_credit_card_amount_only', 1),
      ('cgd_bank_statement_fdm_credit_card_monthly_income', 2),
      ('fdm_bank_transfer_cgd_bank_statement_combination', 2),
      ('cgd_bank_statement_fdm_adyen_monthly_payments', 2)
    )
    and (
      (
        config.rule_key =
          'fdm_bank_transfer_cgd_bank_statement_combination'
        and config.difference_allowed = 0
        and config.max_difference_days between 0 and 90
        and destination.source_type = 'import_cgd_extrato_ordem'
        and definition.definition = jsonb_build_object(
          'strategy', 'bounded_exact_combination',
          'sourceAccount', 'Bank Transfer',
          'maxSourceRecords', 10,
          'candidatePoolLimit', 60,
          'stateLimit', 250000,
          'evidenceGroupLimit', 12,
          'amountMode', 'signed_integer_cents',
          'dateMode', 'inclusive_days'
        )
      )
      or (
        config.rule_key =
          'cgd_bank_statement_fdm_adyen_monthly_payments'
        and config.max_difference_days = 31
        and destination.source_type = 'import_fdm_accounts'
        and definition.definition = jsonb_build_object(
          'strategy', 'closed_calendar_month',
          'bankDescriptionContains', 'Adyen',
          'fdmAccount', 'Adyen',
          'fdmExcludedCategory', 'TransferOutToAccount',
          'requiresBothSides', true,
          'monthMarkerDays', 31
        )
      )
      or (
        config.rule_key =
          'cgd_bank_statement_fdm_credit_card_monthly_income'
        and config.max_difference_days = 31
        and public.financial_reconciliation_automatic_rule_contract(
          config.rule_key, config.rule_version
        )->>'destinationSourceType' = destination.source_type
      )
      or (
        config.rule_key in (
          'financial_documents_cgd_bank_statement',
          'financial_documents_cgd_credit_card',
          'financial_documents_cgd_bank_statement_amount_only',
          'financial_documents_cgd_credit_card_amount_only'
        )
        and config.max_difference_days between 0 and 90
        and public.financial_reconciliation_automatic_rule_contract(
          config.rule_key, config.rule_version
        )->>'destinationSourceType' = destination.source_type
      )
    )
    and source_rule.operator = case
      when config.rule_key in (
        'cgd_bank_statement_fdm_credit_card_monthly_income',
        'fdm_bank_transfer_cgd_bank_statement_combination',
        'cgd_bank_statement_fdm_adyen_monthly_payments'
      ) then '-'
      else '+'
    end;

  if jsonb_array_length(v_snapshot) <> 1 then
    raise exception 'Automatic rule is not enabled for manual analysis.';
  end if;

  v_rule_key := v_snapshot->0->>'ruleKey';
  v_rule_version := (v_snapshot->0->>'ruleVersion')::integer;
  if v_rule_key = 'fdm_bank_transfer_cgd_bank_statement_combination'
    and v_snapshot->0->>'operator' is distinct from '-' then
    raise exception 'Automatic rule source operator is invalid.';
  end if;
  if v_rule_key = 'cgd_bank_statement_fdm_adyen_monthly_payments'
    and v_snapshot->0->>'operator' is distinct from '-' then
    raise exception 'Automatic rule source operator is invalid.';
  end if;

  if v_rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income'
    and v_rule_version = 2 then
    select public.financial_reconciliation_automatic_monthly_income_count()
    into v_total;
  elsif v_rule_key =
      'fdm_bank_transfer_cgd_bank_statement_combination'
    and v_rule_version = 2 then
    select public.financial_reconciliation_automatic_bank_reservation_count()
    into v_total;
  elsif v_rule_key =
      'cgd_bank_statement_fdm_adyen_monthly_payments'
    and v_rule_version = 2 then
    select public.financial_reconciliation_automatic_adyen_month_count()
    into v_total;
  else
    select public.financial_reconciliation_automatic_base_count(
      v_rule_key, v_rule_version
    ) into v_total;
  end if;

  insert into public.financial_reconciliation_automatic_runs (
    trigger, scope, actor, client_request_id, definition_config_snapshot,
    analysis_processed, analysis_total
  ) values (
    'manual', 'rule', p_actor, p_client_request_id, v_snapshot, 0, v_total
  ) returning id into v_run_id;

  return public.continue_financial_reconciliation_automatic_analysis(
    v_run_id, p_actor
  );
end
$$;

revoke all on function public.create_financial_reconciliation_automatic_analysis(
  text[], text, text, uuid
) from public, anon, authenticated;
grant execute on function public.create_financial_reconciliation_automatic_analysis(
  text[], text, text, uuid
) to service_role;

do $migration$
declare
  v_matching_rules integer;
begin
  select count(*)::integer
  into v_matching_rules
  from public.financial_reconciliation_automatic_rule_configs config
  join public.financial_reconciliation_automatic_rule_definitions definition
    on definition.rule_key = config.rule_key
   and definition.version = config.rule_version
  cross join lateral jsonb_array_elements_text(
    definition.destination_source_types
  ) destination(source_type)
  join public.financial_reconciliation_source_rules source_rule
    on source_rule.base_source_type = definition.base_source_type
   and source_rule.matching_source_type = destination.source_type
  where config.rule_key =
      'cgd_bank_statement_fdm_adyen_monthly_payments'
    and config.rule_version = 2
    and config.difference_allowed >= 0
    and config.max_difference_days = 31
    and destination.source_type = 'import_fdm_accounts'
    and source_rule.operator = '-'
    and definition.definition = jsonb_build_object(
      'strategy', 'closed_calendar_month',
      'bankDescriptionContains', 'Adyen',
      'fdmAccount', 'Adyen',
      'fdmExcludedCategory', 'TransferOutToAccount',
      'requiresBothSides', true,
      'monthMarkerDays', 31
    );

  if v_matching_rules <> 1 then
    raise exception
      'The managed Adyen v2 rule is not eligible for manual analysis.';
  end if;
end
$migration$;

notify pgrst, 'reload schema';
