-- Upgrade the managed Adyen monthly rule to v2.
-- Eligible FDM destinations require Account = 'Adyen', a non-null Category,
-- and Category <> 'TransferOutToAccount'. Completed v1 audit rows are retained.

do $migration$
declare
  v_rule_key constant text :=
    'cgd_bank_statement_fdm_adyen_monthly_payments';
  v_display_name constant text :=
    'FDM Accounts – Adyen Reservation Payments';
  v_logic constant text :=
    'Every eligible unlocked CGD Bank Statement and FDM Adyen record whose category is not TransferOutToAccount in the same closed calendar month forms one proposal; both sides are required and the signed difference must be within the configured allowance.';
  v_definition constant jsonb := jsonb_build_object(
    'strategy', 'closed_calendar_month',
    'bankDescriptionContains', 'Adyen',
    'fdmAccount', 'Adyen',
    'fdmExcludedCategory', 'TransferOutToAccount',
    'requiresBothSides', true,
    'monthMarkerDays', 31
  );
  v_proc record;
  v_source text;
  v_original text;
begin
  lock table public.financial_reconciliation_source_rules
    in share row exclusive mode;
  lock table public.financial_reconciliation_automatic_rule_configs
    in share row exclusive mode;

  insert into public.financial_reconciliation_automatic_rule_definitions (
    rule_key, version, display_name, base_source_type,
    destination_source_types, logic_description, definition
  ) values (
    v_rule_key, 2, v_display_name, 'import_cgd_extrato_ordem',
    '["import_fdm_accounts"]'::jsonb, v_logic, v_definition
  ) on conflict (rule_key, version) do nothing;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_definitions definition
    where definition.rule_key = v_rule_key
      and definition.version = 2
      and definition.display_name = v_display_name
      and definition.base_source_type = 'import_cgd_extrato_ordem'
      and definition.destination_source_types =
        '["import_fdm_accounts"]'::jsonb
      and definition.logic_description = v_logic
      and definition.definition = v_definition
  ) then
    raise exception
      'Installed Adyen automatic reconciliation definition differs from the approved immutable v2 definition.';
  end if;

  -- Patch only functions that already contain the managed Adyen rule. The
  -- transformation is idempotent and retains v1 in read/lifecycle allowlists
  -- so completed historical snapshots remain readable.
  for v_proc in
    select procedure.oid, procedure.proname
    from pg_proc procedure
    join pg_namespace namespace on namespace.oid = procedure.pronamespace
    where namespace.nspname = 'public'
      and (
        position(v_rule_key in pg_get_functiondef(procedure.oid)) > 0
        or procedure.proname in (
          'financial_reconciliation_automatic_adyen_month_count',
          'financial_reconciliation_automatic_adyen_month_page'
        )
      )
    order by procedure.oid
  loop
    v_source := pg_get_functiondef(v_proc.oid);
    v_original := v_source;

    if position($old$('cgd_bank_statement_fdm_adyen_monthly_payments', 1)$old$
        in v_source) > 0
      and position($new$('cgd_bank_statement_fdm_adyen_monthly_payments', 2)$new$
        in v_source) = 0 then
      v_source := replace(
        v_source,
        $old$('cgd_bank_statement_fdm_adyen_monthly_payments', 1)$old$,
        $new$('cgd_bank_statement_fdm_adyen_monthly_payments', 1),
          ('cgd_bank_statement_fdm_adyen_monthly_payments', 2)$new$
      );
    end if;

    if position($old$'fdmAccount', 'Adyen',
      'requiresBothSides'$old$ in v_source) > 0
      and position($new$'fdmExcludedCategory', 'TransferOutToAccount'$new$
        in v_source) = 0 then
      v_source := replace(
        v_source,
        $old$'fdmAccount', 'Adyen',
      'requiresBothSides'$old$,
        $new$'fdmAccount', 'Adyen',
      'fdmExcludedCategory', 'TransferOutToAccount',
      'requiresBothSides'$new$
      );
    end if;

    if position($old$and fdm.account = 'Adyen'$old$ in v_source) > 0
      and position($new$and fdm.category is not null
          and fdm.category <> 'TransferOutToAccount'$new$ in v_source) = 0 then
      v_source := replace(
        v_source,
        $old$and fdm.account = 'Adyen'$old$,
        $new$and fdm.account = 'Adyen'
          and fdm.category is not null
          and fdm.category <> 'TransferOutToAccount'$new$
      );
    end if;

    if position($old$'destinationAccount', 'Adyen',$old$ in v_source) > 0
      and position($new$'destinationExcludedCategory',
      'TransferOutToAccount',$new$ in v_source) = 0 then
      v_source := replace(
        v_source,
        $old$'destinationAccount', 'Adyen',$old$,
        $new$'destinationAccount', 'Adyen',
      'destinationExcludedCategory', 'TransferOutToAccount',$new$
      );
    end if;

    if v_proc.proname in (
      'financial_reconciliation_continue_automatic_adyen_monthly',
      'financial_reconciliation_continue_automatic_adyen_mutable_prior'
    ) then
      v_source := replace(
        v_source,
        $old$v_rule->>'ruleVersion' is distinct from '1'$old$,
        $new$v_rule->>'ruleVersion' is distinct from '2'$new$
      );
      v_source := replace(
        v_source,
        $old$'ruleVersion', 1$old$,
        $new$'ruleVersion', 2$new$
      );
      v_source := replace(
        v_source,
        $old$'cgd_bank_statement_fdm_adyen_monthly_payments', 1,$old$,
        $new$'cgd_bank_statement_fdm_adyen_monthly_payments', 2,$new$
      );
      v_source := replace(
        v_source,
        $old$proposal.rule_version = 1$old$,
        $new$proposal.rule_version = 2$new$
      );
    elsif v_proc.proname =
        'financial_reconciliation_execute_adyen_monthly_proposal' then
      v_source := replace(
        v_source,
        $old$v_proposal.rule_version <> 1$old$,
        $new$v_proposal.rule_version <> 2$new$
      );
      v_source := replace(
        v_source,
        $old$v_proposal.summary_snapshot->>'ruleVersion' is distinct from '1'$old$,
        $new$v_proposal.summary_snapshot->>'ruleVersion' is distinct from '2'$new$
      );
      if position(
          $old$v_proposal.summary_snapshot->>'destinationExcludedCategory'$old$
          in v_source) = 0 then
        v_source := replace(
          v_source,
          $old$v_proposal.summary_snapshot->>'destinationAccount' is distinct from
        'Adyen'$old$,
          $new$v_proposal.summary_snapshot->>'destinationAccount' is distinct from
        'Adyen'
      or v_proposal.summary_snapshot->>'destinationExcludedCategory' is distinct from
        'TransferOutToAccount'$new$
        );
      end if;
      if position($old$membership.account <> 'Adyen'$old$ in v_source) > 0
        and position($new$membership.row_snapshot->>'category' is null$new$
          in v_source) = 0 then
        v_source := replace(
          v_source,
          $old$membership.account <> 'Adyen'$old$,
          $new$membership.account <> 'Adyen'
              or membership.row_snapshot->>'category' is null
              or membership.row_snapshot->>'category' =
                'TransferOutToAccount'$new$
        );
      end if;
    elsif v_proc.proname =
        'financial_reconciliation_finalize_automatic_analysis' then
      v_source := replace(
        v_source,
        $old$v_rule_version is distinct from 1$old$,
        $new$v_rule_version is distinct from 2$new$
      );
      v_source := replace(
        v_source,
        $old$proposal.rule_version = 1$old$,
        $new$proposal.rule_version = 2$new$
      );
    end if;

    v_source := regexp_replace(
      v_source,
      '(v_rule_key\s*=\s*''cgd_bank_statement_fdm_adyen_monthly_payments''\s*and\s+v_rule_version\s*=\s*)1(\s+then)',
      E'\\1 2\\2',
      'g'
    );
    v_source := regexp_replace(
      v_source,
      '(v_selected_rule_key\s*=\s*''cgd_bank_statement_fdm_adyen_monthly_payments''\s*and\s+v_selected_rule_version\s*=\s*)1(\s+then)',
      E'\\1 2\\2',
      'g'
    );

    if v_source is distinct from v_original then
      execute v_source;
    end if;
  end loop;

  if position(
      $expected$v_rule->>'ruleVersion' is distinct from '2'$expected$
      in pg_get_functiondef(
        'public.financial_reconciliation_continue_automatic_adyen_monthly(uuid,text)'::regprocedure
      )
    ) = 0
    or position(
      $expected$and fdm.category <> 'TransferOutToAccount'$expected$
      in pg_get_functiondef(
        'public.financial_reconciliation_continue_automatic_adyen_monthly(uuid,text)'::regprocedure
      )
    ) = 0
    or position(
      $expected$v_proposal.rule_version <> 2$expected$
      in pg_get_functiondef(
        'public.financial_reconciliation_execute_adyen_monthly_proposal(uuid,text)'::regprocedure
      )
    ) = 0
    or position(
      $expected$and fdm.category <> 'TransferOutToAccount'$expected$
      in pg_get_functiondef(
        'public.financial_reconciliation_execute_adyen_monthly_proposal(uuid,text)'::regprocedure
      )
    ) = 0
    or position(
      $expected$v_rule_version is distinct from 2$expected$
      in pg_get_functiondef(
        'public.financial_reconciliation_finalize_automatic_analysis(uuid)'::regprocedure
      )
    ) = 0
    or position(
      $expected$proposal.rule_version = 2$expected$
      in pg_get_functiondef(
        'public.financial_reconciliation_finalize_automatic_analysis(uuid)'::regprocedure
      )
    ) = 0
    or position(
      $expected$and fdm.category <> 'TransferOutToAccount'$expected$
      in pg_get_functiondef(
        'public.financial_reconciliation_automatic_adyen_month_count()'::regprocedure
      )
    ) = 0
    or position(
      $expected$and fdm.category <> 'TransferOutToAccount'$expected$
      in pg_get_functiondef(
        'public.financial_reconciliation_automatic_adyen_month_page(date,integer)'::regprocedure
      )
    ) = 0 then
    raise exception 'Adyen v2 analysis/execution functions were not upgraded.';
  end if;

  -- The config changes only after the v2 functions are installed. All
  -- administrator flags, tolerances, schedule choices, and priority remain.
  update public.financial_reconciliation_automatic_rule_configs config
  set rule_version = 2,
      updated_at = now()
  where config.rule_key = v_rule_key
    and config.rule_version = 1;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs config
    where config.rule_key = v_rule_key
      and config.rule_version = 2
      and config.max_difference_days = 31
      and config.difference_allowed >= 0
  ) then
    raise exception 'Adyen v2 managed configuration was not installed.';
  end if;

  -- Existing unfinished v1 analyses must be reviewed again under v2. Their
  -- immutable snapshots remain untouched; completed v1 history is unchanged.
  update public.financial_reconciliation_automatic_proposals proposal
  set status = 'stale',
      reason = 'rule_version_changed',
      reconciliation_id = null,
      completed_at = null,
      error = '',
      error_detail = '',
      updated_at = now()
  from public.financial_reconciliation_automatic_runs run
  where proposal.run_id = run.id
    and run.finished_at is null
    and proposal.rule_key = v_rule_key
    and proposal.rule_version = 1
    and proposal.status in ('proposed', 'ambiguous', 'executing');

  update public.financial_reconciliation_automatic_runs run
  set status = 'failed',
      analysis_completed_at = coalesce(run.analysis_completed_at, now()),
      analysis_error_code = 'rule_version_changed',
      analysis_error_at = now(),
      error_summary = 'The automatic rule changed. Analyze it again.',
      finished_at = now(),
      updated_at = now()
  where run.finished_at is null
    and jsonb_typeof(run.definition_config_snapshot) = 'array'
    and jsonb_array_length(run.definition_config_snapshot) = 1
    and run.definition_config_snapshot#>>'{0,ruleKey}' = v_rule_key
    and run.definition_config_snapshot#>>'{0,ruleVersion}' = '1';
end
$migration$;

create index if not exists
  import_fdm_accounts_adyen_v2_month_id_idx
on public.import_fdm_accounts (event_date, id)
where event_date >= date '2026-01-01'
  and amount is not null
  and account = 'Adyen'
  and category is not null
  and category <> 'TransferOutToAccount';

do $migration$
declare
  v_definition text;
begin
  select pg_get_indexdef(index_row.indexrelid)
  into v_definition
  from pg_index index_row
  where index_row.indexrelid =
    'public.import_fdm_accounts_adyen_v2_month_id_idx'::regclass
    and index_row.indisvalid
    and index_row.indisready;

  if v_definition is null
    or v_definition not ilike '%(event_date, id)%'
    or v_definition not ilike '%account = ''Adyen''%'
    or v_definition not ilike '%category is not null%'
    or v_definition not ilike
      '%category <> ''TransferOutToAccount''%' then
    raise exception 'Installed Adyen v2 eligibility index is invalid.';
  end if;
end
$migration$;
