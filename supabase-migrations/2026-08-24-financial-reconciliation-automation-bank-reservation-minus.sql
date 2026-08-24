-- Upgrade Bank Reservation to v2 and use subtraction in both Bank/FDM
-- source-rule directions. Completed v1 audit rows remain immutable; unfinished
-- v1 analyses are closed and must be analyzed again under the v2 contract.

do $migration$
declare
  v_rule_key constant text :=
    'fdm_bank_transfer_cgd_bank_statement_combination';
  v_display_name constant text :=
    'FDM Accounts – Bank Reservation Payments';
  v_logic constant text :=
    'Exactly one CGD Bank Statement record is matched to one through ten eligible FDM Bank Transfer records whose total minus the Bank Statement amount equals zero exactly in integer cents within the inclusive configured date window.';
  v_definition constant jsonb := jsonb_build_object(
    'strategy', 'bounded_exact_combination',
    'sourceAccount', 'Bank Transfer',
    'maxSourceRecords', 10,
    'candidatePoolLimit', 60,
    'stateLimit', 250000,
    'evidenceGroupLimit', 12,
    'amountMode', 'signed_integer_cents',
    'dateMode', 'inclusive_days'
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
    v_rule_key, 2, v_display_name, 'import_fdm_accounts',
    '["import_cgd_extrato_ordem"]'::jsonb, v_logic, v_definition
  ) on conflict (rule_key, version) do nothing;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_definitions definition
    where definition.rule_key = v_rule_key
      and definition.version = 2
      and definition.display_name = v_display_name
      and definition.base_source_type = 'import_fdm_accounts'
      and definition.destination_source_types =
        '["import_cgd_extrato_ordem"]'::jsonb
      and definition.logic_description = v_logic
      and definition.definition = v_definition
  ) then
    raise exception
      'Installed Bank Reservation definition differs from the approved immutable v2 definition.';
  end if;

  insert into public.financial_reconciliation_source_rules (
    base_source_type, matching_source_type, operator
  ) values (
    'import_fdm_accounts', 'import_cgd_extrato_ordem', '-'
  ) on conflict (base_source_type, matching_source_type)
  do update set operator = excluded.operator;

  -- Patch only ordinary functions. This avoids pg_get_functiondef aggregate
  -- errors and keeps historical migration files byte-for-byte unchanged.
  for v_proc in
    select procedure.oid, procedure.proname
    from pg_proc procedure
    join pg_namespace namespace on namespace.oid = procedure.pronamespace
    where namespace.nspname = 'public'
      and procedure.prokind = 'f'
      and (
        position(v_rule_key in procedure.prosrc) > 0
        or procedure.proname in (
          'financial_reconciliation_automatic_bank_reservation_groups',
          'financial_reconciliation_commit_fdm_bank_automatic_proposal',
          'replace_financial_reconciliation_source_rules'
        )
      )
    order by procedure.oid
  loop
    v_source := pg_get_functiondef(v_proc.oid);
    v_original := v_source;

    -- Historical readers retain v1 while current catalogs, dispatchers and
    -- execution paths move to v2. This keeps completed v1 snapshots readable.
    if v_proc.proname in (
      'get_financial_reconciliation_automatic_run',
      'financial_reconciliation_automatic_progress_or_run',
      'get_financial_reconciliation_automatic_proposal_members',
      'financial_reconciliation_refresh_automatic_batch'
    ) then
      if position(
          $old$('fdm_bank_transfer_cgd_bank_statement_combination', 1)$old$
          in v_source) > 0
        and position(
          $new$('fdm_bank_transfer_cgd_bank_statement_combination', 2)$new$
          in v_source) = 0 then
        v_source := replace(
          v_source,
          $old$('fdm_bank_transfer_cgd_bank_statement_combination', 1)$old$,
          $new$('fdm_bank_transfer_cgd_bank_statement_combination', 1),
            ('fdm_bank_transfer_cgd_bank_statement_combination', 2)$new$
        );
      end if;
      if v_proc.proname in (
        'get_financial_reconciliation_automatic_run',
        'financial_reconciliation_automatic_progress_or_run'
      ) then
        if v_proc.proname = 'get_financial_reconciliation_automatic_run' then
          v_source := regexp_replace(
            v_source,
            '=\s*\(''fdm_bank_transfer_cgd_bank_statement_combination'',\s*1\),\s*\(''fdm_bank_transfer_cgd_bank_statement_combination'',\s*2\)\s+then',
            $replacement$in (('fdm_bank_transfer_cgd_bank_statement_combination', 1),
        ('fdm_bank_transfer_cgd_bank_statement_combination', 2)) then$replacement$,
            'g'
          );
        end if;
        v_source := regexp_replace(
          v_source,
          '(v_rule_key\s*=\s*''fdm_bank_transfer_cgd_bank_statement_combination''\s+and\s+v_rule_version\s*=\s*)1(\s+then\s+''bank_anchors'')',
          E'v_rule_key = ''fdm_bank_transfer_cgd_bank_statement_combination''\n      and v_rule_version in (1, 2)\\2',
          'g'
        );
      elsif v_proc.proname =
          'financial_reconciliation_refresh_automatic_batch' then
        v_source := replace(
          v_source,
          $old$or snapshot.value->>'operator' is distinct from '+'
            or snapshot.value->>'destinationSourceType' is distinct from
              'import_cgd_extrato_ordem'$old$,
          $new$or (
              (snapshot.value->>'ruleVersion')::integer = 1
              and snapshot.value->>'operator' is distinct from '+'
            )
            or (
              (snapshot.value->>'ruleVersion')::integer = 2
              and snapshot.value->>'operator' is distinct from '-'
            )
            or snapshot.value->>'destinationSourceType' is distinct from
              'import_cgd_extrato_ordem'$new$
        );
      elsif v_proc.proname =
          'get_financial_reconciliation_automatic_proposal_members'
        and position(
          'Historical Bank Reservation proposal members require a finished run.'
          in v_source
        ) = 0 then
        v_source := replace(
          v_source,
          $old$if v_run.finished_at is null and v_run.actor <> p_actor then
    raise exception 'You do not have permission for this automation run.';
  end if;$old$,
          $new$if v_run.rule_key =
      'fdm_bank_transfer_cgd_bank_statement_combination'
    and v_run.rule_version = 1
    and v_run.finished_at is null then
    raise exception
      'Historical Bank Reservation proposal members require a finished run.';
  end if;
  if v_run.finished_at is null and v_run.actor <> p_actor then
    raise exception 'You do not have permission for this automation run.';
  end if;$new$
        );
      end if;
    else
      v_source := replace(
        v_source,
        $old$('fdm_bank_transfer_cgd_bank_statement_combination', 1)$old$,
        $new$('fdm_bank_transfer_cgd_bank_statement_combination', 2)$new$
      );
      v_source := regexp_replace(
        v_source,
        '(v_rule_key\s*=\s*''fdm_bank_transfer_cgd_bank_statement_combination''\s*and\s+v_rule_version\s*=\s*)1(\s+then)',
        E'\\1 2\\2',
        'g'
      );
      v_source := regexp_replace(
        v_source,
        '(v_selected_rule_key\s*=\s*''fdm_bank_transfer_cgd_bank_statement_combination''\s*and\s+v_selected_rule_version\s*=\s*)1(\s+then)',
        E'\\1 2\\2',
        'g'
      );
    end if;

    if v_proc.proname =
        'financial_reconciliation_automatic_bank_reservation_groups' then
      v_source := regexp_replace(
        v_source,
        '=\s*-sign\(v_bank_amount_cents\)',
        '= sign(v_bank_amount_cents)',
        'g'
      );
      v_source := replace(
        v_source,
        $old$'equationCents', v_path_signed_total_cents + v_bank_amount_cents$old$,
        $new$'equationCents', v_path_signed_total_cents - v_bank_amount_cents$new$
      );
    elsif v_proc.proname =
        'financial_reconciliation_continue_automatic_bank_reservation' then
      v_source := replace(
        v_source,
        $old$v_rule->>'ruleVersion' is distinct from '1'$old$,
        $new$v_rule->>'ruleVersion' is distinct from '2'$new$
      );
      v_source := replace(v_source, $old$'ruleVersion', 1$old$,
        $new$'ruleVersion', 2$new$);
      v_source := replace(
        v_source,
        $old$p_run_id, 'fdm_bank_transfer_cgd_bank_statement_combination', 1,$old$,
        $new$p_run_id, 'fdm_bank_transfer_cgd_bank_statement_combination', 2,$new$
      );
      v_source := replace(v_source, $old$proposal.rule_version = 1$old$,
        $new$proposal.rule_version = 2$new$);
      v_source := replace(v_source, $old$v_operator is distinct from '+'$old$,
        $new$v_operator is distinct from '-'$new$);
    elsif v_proc.proname =
        'financial_reconciliation_execute_bank_reservation_proposal' then
      v_source := replace(v_source, $old$v_proposal.rule_version <> 1$old$,
        $new$v_proposal.rule_version <> 2$new$);
      v_source := replace(
        v_source,
        $old$v_rule_snapshot->>'operator' is distinct from '+'$old$,
        $new$v_rule_snapshot->>'operator' is distinct from '-'$new$
      );
      v_source := replace(v_source, $old$v_current_operator is distinct from '+'$old$,
        $new$v_current_operator is distinct from '-'$new$);
      v_source := replace(
        v_source,
        $old$v_proposal.summary_snapshot->>'ruleVersion' is distinct from '1'$old$,
        $new$v_proposal.summary_snapshot->>'ruleVersion' is distinct from '2'$new$
      );
      v_source := replace(
        v_source,
        $old$v_proposal.summary_snapshot->>'operator' is distinct from '+'$old$,
        $new$v_proposal.summary_snapshot->>'operator' is distinct from '-'$new$
      );
      v_source := replace(
        v_source,
        $old$v_equation_cents := v_source_total_cents + v_bank_amount_cents$old$,
        $new$v_equation_cents := v_source_total_cents - v_bank_amount_cents$new$
      );
      v_source := replace(v_source, $old$v_source_total + v_bank_amount <> 0$old$,
        $new$v_source_total - v_bank_amount <> 0$new$);
      v_source := regexp_replace(
        v_source,
        '<>\s*-sign\(v_bank_amount_cents\)',
        '<> sign(v_bank_amount_cents)',
        'g'
      );
      v_source := replace(v_source, $old$'operator', '+'$old$,
        $new$'operator', '-'$new$);
      v_source := regexp_replace(
        v_source,
        '(''import_fdm_accounts''\s*,\s*''import_cgd_extrato_ordem''\s*,\s*)''\+''(\s*,)',
        E'\\1''-''\\2',
        'g'
      );
    elsif v_proc.proname =
        'financial_reconciliation_commit_fdm_bank_automatic_proposal' then
      v_source := replace(
        v_source,
        $old$('import_fdm_accounts', 'import_cgd_extrato_ordem', '+')$old$,
        $new$('import_fdm_accounts', 'import_cgd_extrato_ordem', '-')$new$
      );
    elsif v_proc.proname = 'replace_financial_reconciliation_source_rules'
      and position('managed Bank Reservation source rule' in v_source) = 0 then
      v_source := regexp_replace(
        v_source,
        '(raise exception ''The managed POS income source rule must remain enabled with operator -\.'';\s*end if;)',
        E'\\1\n  if (\n    select count(*)\n    from jsonb_to_recordset(p_rules) as rule(\n      base_source_type text, matching_source_type text, operator text\n    )\n    where rule.base_source_type = ''import_fdm_accounts''\n      and rule.matching_source_type = ''import_cgd_extrato_ordem''\n      and rule.operator = ''-''\n  ) <> 1 then\n    raise exception ''The managed Bank Reservation source rule must remain enabled with operator -.'';\n  end if;',
        'g'
      );
    else
      -- Shared serializers, settings, create/claim/finalize and dispatch paths.
      v_source := regexp_replace(
        v_source,
        '(when\s+config\.rule_key\s*=\s*''fdm_bank_transfer_cgd_bank_statement_combination''\s+then\s*)''\+''',
        E'\\1''-''',
        'g'
      );
      v_source := regexp_replace(
        v_source,
        '(snapshot\.value->>''ruleKey''\s*=\s*''fdm_bank_transfer_cgd_bank_statement_combination''\s+and\s+\(\s+\(snapshot\.value->>''differenceAllowed''\)::numeric\s*<>\s*0\s+or\s+snapshot\.value->>''operator''\s+is distinct from\s*)''\+''',
        E'\\1''-''',
        'g'
      );
      if position(
          $current$'fdm_bank_transfer_cgd_bank_statement_combination'$current$
          in v_source) > 0 then
        v_source := regexp_replace(
          v_source,
          '(''cgd_bank_statement_fdm_credit_card_monthly_income''\s*,)(\s*)(''cgd_bank_statement_fdm_adyen_monthly_payments''\s*\)\s*then\s*''-'')',
          E'\\1\\2''fdm_bank_transfer_cgd_bank_statement_combination'',\\2\\3',
          'g'
        );
      end if;
      v_source := replace(
        v_source,
        $old$v_snapshot->0->>'operator' is distinct from '+'$old$,
        $new$v_snapshot->0->>'operator' is distinct from '-'$new$
      );
    end if;

    if v_source is distinct from v_original then
      execute v_source;
    end if;
  end loop;

  -- Fail closed if any current Bank Reservation path retained v1 or plus.
  if position(
      $expected$v_rule->>'ruleVersion' is distinct from '2'$expected$
      in pg_get_functiondef(
        'public.financial_reconciliation_continue_automatic_bank_reservation(uuid,text)'::regprocedure
      )
    ) = 0
    or position(
      $expected$v_operator is distinct from '-'$expected$
      in pg_get_functiondef(
        'public.financial_reconciliation_continue_automatic_bank_reservation(uuid,text)'::regprocedure
      )
    ) = 0
    or position(
      $expected$= sign(v_bank_amount_cents)$expected$
      in pg_get_functiondef(
        'public.financial_reconciliation_automatic_bank_reservation_groups(uuid,integer,integer,integer,integer)'::regprocedure
      )
    ) = 0
    or position(
      $expected$'equationCents', v_path_signed_total_cents - v_bank_amount_cents$expected$
      in pg_get_functiondef(
        'public.financial_reconciliation_automatic_bank_reservation_groups(uuid,integer,integer,integer,integer)'::regprocedure
      )
    ) = 0
    or position(
      $expected$v_proposal.rule_version <> 2$expected$
      in pg_get_functiondef(
        'public.financial_reconciliation_execute_bank_reservation_proposal(uuid,text)'::regprocedure
      )
    ) = 0
    or position(
      $expected$v_equation_cents := v_source_total_cents - v_bank_amount_cents$expected$
      in pg_get_functiondef(
        'public.financial_reconciliation_execute_bank_reservation_proposal(uuid,text)'::regprocedure
      )
    ) = 0
    or position(
      $expected$('import_fdm_accounts', 'import_cgd_extrato_ordem', '-')$expected$
      in pg_get_functiondef(
        'public.financial_reconciliation_commit_fdm_bank_automatic_proposal(uuid,text,text,text,text,numeric,text,jsonb,jsonb,jsonb,jsonb)'::regprocedure
      )
    ) = 0
    or position(
      $expected$managed Bank Reservation source rule must remain enabled with operator -$expected$
      in pg_get_functiondef(
        'public.replace_financial_reconciliation_source_rules(jsonb)'::regprocedure
      )
    ) = 0
    or pg_get_functiondef(
      'public.get_financial_reconciliation_automatic_manual_rules()'::regprocedure
    ) !~ 'case when config\.rule_key in\s*\([^)]*fdm_bank_transfer_cgd_bank_statement_combination[^)]*\)\s*then\s*''-'''
    or pg_get_functiondef(
      'public.create_financial_reconciliation_automatic_analysis(text[],text,text,uuid)'::regprocedure
    ) !~ 'when config\.rule_key in\s*\([^)]*fdm_bank_transfer_cgd_bank_statement_combination[^)]*\)\s*then\s*''-'''
    or position(
      $expected$when config.rule_key =
              'fdm_bank_transfer_cgd_bank_statement_combination'
            then '-'$expected$
      in pg_get_functiondef(
        'public.claim_financial_reconciliation_automatic_schedule(text)'::regprocedure
      )
    ) = 0
    or position(
      $expected$('fdm_bank_transfer_cgd_bank_statement_combination', 1)$expected$
      in pg_get_functiondef(
        'public.get_financial_reconciliation_automatic_run(uuid)'::regprocedure
      )
    ) = 0
    or position(
      $expected$('fdm_bank_transfer_cgd_bank_statement_combination', 2)$expected$
      in pg_get_functiondef(
        'public.get_financial_reconciliation_automatic_run(uuid)'::regprocedure
      )
    ) = 0
    or position(
      $expected$Historical Bank Reservation proposal members require a finished run.$expected$
      in pg_get_functiondef(
        'public.get_financial_reconciliation_automatic_proposal_members(uuid,uuid,text,integer,integer,text)'::regprocedure
      )
    ) = 0
    or pg_get_functiondef(
      'public.financial_reconciliation_automatic_progress_or_run(uuid)'::regprocedure
    ) !~ 'v_rule_version in\s*\(1,\s*2\)'
    or position(
      $expected$('fdm_bank_transfer_cgd_bank_statement_combination', 1)$expected$
      in pg_get_functiondef(
        'public.get_financial_reconciliation_automatic_proposal_members(uuid,uuid,text,integer,integer,text)'::regprocedure
      )
    ) = 0
    or position(
      $expected$('fdm_bank_transfer_cgd_bank_statement_combination', 2)$expected$
      in pg_get_functiondef(
        'public.get_financial_reconciliation_automatic_proposal_members(uuid,uuid,text,integer,integer,text)'::regprocedure
      )
    ) = 0
    or pg_get_functiondef(
      'public.financial_reconciliation_refresh_automatic_batch(uuid)'::regprocedure
    ) !~ 'ruleVersion''\)::integer = 1\s+and snapshot\.value->>''operator'' is distinct from ''\+'''
    or pg_get_functiondef(
      'public.financial_reconciliation_refresh_automatic_batch(uuid)'::regprocedure
    ) !~ 'ruleVersion''\)::integer = 2\s+and snapshot\.value->>''operator'' is distinct from ''-'''
    then
    raise exception 'Bank Reservation v2 subtraction functions were not upgraded.';
  end if;

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
      and config.difference_allowed = 0
      and config.max_difference_days between 0 and 90
  ) then
    raise exception 'Bank Reservation v2 managed configuration was not installed.';
  end if;

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
