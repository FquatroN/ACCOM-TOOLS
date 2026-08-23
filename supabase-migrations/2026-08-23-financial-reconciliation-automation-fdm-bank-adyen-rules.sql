do $migration$
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
begin
  insert into public.financial_reconciliation_automatic_rule_definitions (
    rule_key, version, display_name, base_source_type,
    destination_source_types, logic_description, definition
  ) values
  (
    'fdm_bank_transfer_cgd_bank_statement_combination',
    1,
    'FDM Accounts – Bank Reservation Payments',
    'import_fdm_accounts',
    '["import_cgd_extrato_ordem"]'::jsonb,
    v_bank_logic,
    v_bank_definition
  ),
  (
    'cgd_bank_statement_fdm_adyen_monthly_payments',
    1,
    'FDM Accounts – Adyen Reservation Payments',
    'import_cgd_extrato_ordem',
    '["import_fdm_accounts"]'::jsonb,
    v_adyen_logic,
    v_adyen_definition
  )
  on conflict (rule_key, version) do nothing;

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
    raise exception 'Installed Bank Reservation automatic reconciliation definition differs from the approved immutable v1 definition.';
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
    raise exception 'Installed Adyen automatic reconciliation definition differs from the approved immutable v1 definition.';
  end if;
end
$migration$;

do $migration$
declare
  v_next_priority integer;
begin
  lock table public.financial_reconciliation_automatic_rule_configs
    in share row exclusive mode;

  if exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs config
    where (config.rule_key, config.rule_version) not in (
      ('financial_documents_cgd_bank_statement', 2),
      ('financial_documents_cgd_credit_card', 1),
      ('financial_documents_cgd_bank_statement_amount_only', 1),
      ('financial_documents_cgd_credit_card_amount_only', 1),
      ('cgd_bank_statement_fdm_credit_card_monthly_income', 2),
      ('fdm_bank_transfer_cgd_bank_statement_combination', 1),
      ('cgd_bank_statement_fdm_adyen_monthly_payments', 1)
    )
  ) then
    raise exception 'Installed automatic reconciliation configuration is not in the seven-rule managed allowlist.';
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs config
    where config.rule_key =
      'fdm_bank_transfer_cgd_bank_statement_combination'
  ) then
    select coalesce(max(config.priority), 0) + 1
    into v_next_priority
    from public.financial_reconciliation_automatic_rule_configs config;

    insert into public.financial_reconciliation_automatic_rule_configs (
      rule_key, rule_version, enabled, allow_manual_execution,
      include_in_scheduled_batch, difference_allowed,
      max_difference_days, priority
    ) values (
      'fdm_bank_transfer_cgd_bank_statement_combination',
      1, false, false, false, 0.00, 3, v_next_priority
    );
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs config
    where config.rule_key =
      'cgd_bank_statement_fdm_adyen_monthly_payments'
  ) then
    select coalesce(max(config.priority), 0) + 1
    into v_next_priority
    from public.financial_reconciliation_automatic_rule_configs config;

    insert into public.financial_reconciliation_automatic_rule_configs (
      rule_key, rule_version, enabled, allow_manual_execution,
      include_in_scheduled_batch, difference_allowed,
      max_difference_days, priority
    ) values (
      'cgd_bank_statement_fdm_adyen_monthly_payments',
      1, false, false, false, 2000.00, 31, v_next_priority
    );
  end if;

  if not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_configs config
      where config.rule_key =
          'fdm_bank_transfer_cgd_bank_statement_combination'
        and config.rule_version = 1
        and config.difference_allowed = 0
        and config.max_difference_days between 0 and 90
    ) or not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_configs config
      where config.rule_key =
          'cgd_bank_statement_fdm_adyen_monthly_payments'
        and config.rule_version = 1
        and config.difference_allowed >= 0
        and config.max_difference_days = 31
    ) then
    raise exception 'Installed FDM Bank/Adyen automatic reconciliation configuration violates its managed fixed values.';
  end if;
end
$migration$;

do $migration$
begin
  if not exists (
    select 1
    from pg_constraint constraint_row
    where constraint_row.conrelid =
        'public.financial_reconciliation_automatic_rule_configs'::regclass
      and constraint_row.conname =
        'financial_reconciliation_rule_configs_fdm_bank_adyen_check'
  ) then
    alter table public.financial_reconciliation_automatic_rule_configs
      add constraint financial_reconciliation_rule_configs_fdm_bank_adyen_check
      check (
        (
          rule_key <> 'fdm_bank_transfer_cgd_bank_statement_combination'
          or (difference_allowed = 0 and max_difference_days between 0 and 90)
        )
        and (
          rule_key <> 'cgd_bank_statement_fdm_adyen_monthly_payments'
          or max_difference_days = 31
        )
      ) not valid;
  end if;
end
$migration$;

do $migration$
declare
  v_constraint_type "char";
  v_installed_definition text;
  v_expected_definition text;
begin
  create temporary table task2_config_constraint_expected (
    rule_key text,
    difference_allowed numeric,
    max_difference_days integer,
    constraint task2_config_constraint_expected_check check (
      (
        rule_key <> 'fdm_bank_transfer_cgd_bank_statement_combination'
        or (difference_allowed = 0 and max_difference_days between 0 and 90)
      )
      and (
        rule_key <> 'cgd_bank_statement_fdm_adyen_monthly_payments'
        or max_difference_days = 31
      )
    )
  ) on commit drop;

  select constraint_row.contype,
         pg_get_constraintdef(constraint_row.oid, true)
  into strict v_constraint_type, v_expected_definition
  from pg_constraint constraint_row
  where constraint_row.conrelid =
      'task2_config_constraint_expected'::regclass
    and constraint_row.conname =
      'task2_config_constraint_expected_check';

  select constraint_row.contype,
         regexp_replace(
           pg_get_constraintdef(constraint_row.oid, true),
           '\s+NOT VALID$', ''
         )
  into v_constraint_type, v_installed_definition
  from pg_constraint constraint_row
  where constraint_row.conrelid =
      'public.financial_reconciliation_automatic_rule_configs'::regclass
    and constraint_row.conname =
      'financial_reconciliation_rule_configs_fdm_bank_adyen_check';

  drop table task2_config_constraint_expected;

  if v_constraint_type is distinct from 'c'
    or v_installed_definition is distinct from v_expected_definition then
    raise exception 'Installed FDM Bank/Adyen config constraint differs from the required definition.';
  end if;
end
$migration$;

alter table public.financial_reconciliation_automatic_rule_configs
  validate constraint financial_reconciliation_rule_configs_fdm_bank_adyen_check;

do $migration$
begin
  if to_regclass(
      'public.financial_reconciliation_fdm_bank_transfer_lookup_idx'
    ) is null then
    create index financial_reconciliation_fdm_bank_transfer_lookup_idx
      on public.import_fdm_accounts (event_date, amount, id)
      where account = 'Bank Transfer'
        and event_date >= date '2026-01-01'
        and amount is not null;
  end if;

  if to_regclass(
      'public.financial_reconciliation_fdm_adyen_lookup_idx'
    ) is null then
    create index financial_reconciliation_fdm_adyen_lookup_idx
      on public.import_fdm_accounts (event_date, id)
      include (amount)
      where account = 'Adyen'
        and event_date >= date '2026-01-01'
        and amount is not null;
  end if;

  if to_regclass(
      'public.financial_reconciliation_bank_date_amount_lookup_idx'
    ) is null then
    create index financial_reconciliation_bank_date_amount_lookup_idx
      on public.import_cgd_extrato_ordem (data, montante, id)
      where data >= date '2026-01-01'
        and montante is not null;
  end if;
end
$migration$;

do $migration$
declare
  v_pair record;
  v_actual record;
  v_expected record;
begin
  create temporary table task2_fdm_index_expected (
    event_date date,
    amount numeric(14,2),
    id uuid,
    account text
  ) on commit drop;
  create temporary table task2_bank_index_expected (
    data date,
    montante numeric(14,2),
    id uuid
  ) on commit drop;

  create index task2_expected_fdm_bank_transfer_idx
    on task2_fdm_index_expected (event_date, amount, id)
    where account = 'Bank Transfer'
      and event_date >= date '2026-01-01'
      and amount is not null;
  create index task2_expected_fdm_adyen_idx
    on task2_fdm_index_expected (event_date, id)
    include (amount)
    where account = 'Adyen'
      and event_date >= date '2026-01-01'
      and amount is not null;
  create index task2_expected_bank_date_amount_idx
    on task2_bank_index_expected (data, montante, id)
    where data >= date '2026-01-01'
      and montante is not null;

  for v_pair in
    select *
    from (values
      (
        'financial_reconciliation_fdm_bank_transfer_lookup_idx',
        'task2_expected_fdm_bank_transfer_idx',
        'public.import_fdm_accounts'::regclass
      ),
      (
        'financial_reconciliation_fdm_adyen_lookup_idx',
        'task2_expected_fdm_adyen_idx',
        'public.import_fdm_accounts'::regclass
      ),
      (
        'financial_reconciliation_bank_date_amount_lookup_idx',
        'task2_expected_bank_date_amount_idx',
        'public.import_cgd_extrato_ordem'::regclass
      )
    ) expected(actual_name, expected_name, expected_table)
  loop
    select index_row.indrelid,
           index_row.indisunique,
           index_row.indisvalid,
           index_row.indisready,
           index_row.indnkeyatts,
           index_row.indnatts,
           access_method.amname,
           pg_get_expr(index_row.indpred, index_row.indrelid, true) as predicate,
           array(
             select pg_get_indexdef(index_row.indexrelid, ordinal, true)
             from generate_series(1, index_row.indnatts) ordinal
             order by ordinal
           ) as index_columns
    into strict v_actual
    from pg_index index_row
    join pg_class index_class on index_class.oid = index_row.indexrelid
    join pg_am access_method on access_method.oid = index_class.relam
    join pg_namespace namespace_row on namespace_row.oid = index_class.relnamespace
    where namespace_row.nspname = 'public'
      and index_class.relname = v_pair.actual_name;

    select index_row.indisunique,
           index_row.indnkeyatts,
           index_row.indnatts,
           access_method.amname,
           pg_get_expr(index_row.indpred, index_row.indrelid, true) as predicate,
           array(
             select pg_get_indexdef(index_row.indexrelid, ordinal, true)
             from generate_series(1, index_row.indnatts) ordinal
             order by ordinal
           ) as index_columns
    into strict v_expected
    from pg_index index_row
    join pg_class index_class on index_class.oid = index_row.indexrelid
    join pg_am access_method on access_method.oid = index_class.relam
    where index_class.relname = v_pair.expected_name
      and index_class.relnamespace = pg_my_temp_schema();

    if v_actual.indrelid is distinct from v_pair.expected_table
      or v_actual.indisunique is distinct from v_expected.indisunique
      or not v_actual.indisvalid
      or not v_actual.indisready
      or v_actual.indnkeyatts is distinct from v_expected.indnkeyatts
      or v_actual.indnatts is distinct from v_expected.indnatts
      or v_actual.amname is distinct from v_expected.amname
      or v_actual.predicate is distinct from v_expected.predicate
      or v_actual.index_columns is distinct from v_expected.index_columns then
      raise exception 'Installed index % conflicts with the required definition.',
        v_pair.actual_name;
    end if;
  end loop;

  drop table task2_fdm_index_expected;
  drop table task2_bank_index_expected;
end
$migration$;

create or replace function public.replace_financial_reconciliation_automation_settings(
  p_schedule jsonb,
  p_rules jsonb,
  p_actor text
)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
begin
  if nullif(trim(coalesce(p_actor, '')), '') is null then
    raise exception 'Automation settings actor is required.';
  end if;

  if p_schedule is null or jsonb_typeof(p_schedule) <> 'object'
    or (select count(*) from jsonb_object_keys(p_schedule)) <> 3
    or not (p_schedule ?& array['enabled','time_of_day','time_zone'])
    or exists (
      select 1 from jsonb_object_keys(p_schedule) key
      where key not in ('enabled','time_of_day','time_zone')
    ) then
    raise exception 'Automatic schedule payload is invalid.';
  end if;

  if jsonb_typeof(p_schedule->'enabled') <> 'boolean'
    or jsonb_typeof(p_schedule->'time_of_day') <> 'string'
    or coalesce(p_schedule->>'time_of_day', '')
      !~ '^(?:[01][0-9]|2[0-3]):[0-5][0-9]$'
    or jsonb_typeof(p_schedule->'time_zone') <> 'string'
    or p_schedule->>'time_zone' <> 'Europe/Lisbon' then
    raise exception 'Automatic schedule values are invalid.';
  end if;

  if p_rules is null or jsonb_typeof(p_rules) <> 'array'
    or jsonb_array_length(p_rules) <> 7
    or exists (
      select 1 from jsonb_array_elements(p_rules) rule
      where jsonb_typeof(rule) <> 'object'
    ) then
    raise exception 'Automatic rules payload must contain the seven managed rule objects.';
  end if;

  if exists (
    select 1
    from jsonb_array_elements(p_rules) rule
    where (select count(*) from jsonb_object_keys(rule)) <> 8
       or not (rule ?& array[
         'rule_key','rule_version','enabled','allow_manual_execution',
         'include_in_scheduled_batch','difference_allowed',
         'max_difference_days','priority'
       ])
       or exists (
         select 1 from jsonb_object_keys(rule) key
         where key not in (
           'rule_key','rule_version','enabled','allow_manual_execution',
           'include_in_scheduled_batch','difference_allowed',
           'max_difference_days','priority'
         )
       )
  ) then
    raise exception 'Automatic rule fields are invalid.';
  end if;

  if exists (
    select 1
    from jsonb_array_elements(p_rules) rule
    where jsonb_typeof(rule->'rule_key') <> 'string'
       or nullif(trim(rule->>'rule_key'), '') is null
       or jsonb_typeof(rule->'rule_version') <> 'number'
       or coalesce(rule->>'rule_version', '') !~ '^[0-9]+$'
       or (rule->>'rule_version')::numeric not between 1 and 2147483647
       or jsonb_typeof(rule->'enabled') <> 'boolean'
       or jsonb_typeof(rule->'allow_manual_execution') <> 'boolean'
       or jsonb_typeof(rule->'include_in_scheduled_batch') <> 'boolean'
       or jsonb_typeof(rule->'difference_allowed') <> 'string'
       or coalesce(rule->>'difference_allowed', '')
         !~ '^(0|[0-9]+)(\.[0-9]{1,2})?$'
       or (rule->>'difference_allowed')::numeric
         not between 0 and 999999999999.99
       or jsonb_typeof(rule->'max_difference_days') <> 'number'
       or coalesce(rule->>'max_difference_days', '') !~ '^[0-9]+$'
       or (rule->>'max_difference_days')::numeric not between 0 and 90
       or jsonb_typeof(rule->'priority') <> 'number'
       or coalesce(rule->>'priority', '') !~ '^[0-9]+$'
       or (rule->>'priority')::numeric not between 1 and 2147483647
  ) then
    raise exception 'Automatic rule values are invalid.';
  end if;

  if exists (
    select 1
    from jsonb_array_elements(p_rules) rule
    group by (rule->>'priority')::integer
    having count(*) > 1
  ) then
    raise exception 'Duplicate automatic rule priority.';
  end if;

  if exists (
    select 1
    from jsonb_array_elements(p_rules) rule
    group by rule->>'rule_key'
    having count(*) > 1
  ) then
    raise exception 'Duplicate automatic rule.';
  end if;

  if exists (
    select 1
    from jsonb_array_elements(p_rules) rule
    where (rule->>'rule_key', (rule->>'rule_version')::integer) not in (
      ('financial_documents_cgd_bank_statement', 2),
      ('financial_documents_cgd_credit_card', 1),
      ('financial_documents_cgd_bank_statement_amount_only', 1),
      ('financial_documents_cgd_credit_card_amount_only', 1),
      ('cgd_bank_statement_fdm_credit_card_monthly_income', 2),
      ('fdm_bank_transfer_cgd_bank_statement_combination', 1),
      ('cgd_bank_statement_fdm_adyen_monthly_payments', 1)
    )
  ) then
    raise exception 'Automatic rule/version is invalid.';
  end if;

  if exists (
    select 1
    from jsonb_array_elements(p_rules) rule
    where rule->>'rule_key' in (
      'financial_documents_cgd_bank_statement_amount_only',
      'financial_documents_cgd_credit_card_amount_only'
    )
      and (rule->>'difference_allowed')::numeric <> 0
  ) then
    raise exception 'Amount-only automatic rules require zero difference allowed.';
  end if;

  if exists (
    select 1
    from jsonb_array_elements(p_rules) rule
    where rule->>'rule_key' =
        'cgd_bank_statement_fdm_credit_card_monthly_income'
      and (rule->>'max_difference_days')::integer <> 31
  ) then
    raise exception
      'POS income automatic rule requires the fixed 31-day display property.';
  end if;

  if exists (
    select 1
    from jsonb_array_elements(p_rules) rule
    where rule->>'rule_key' =
        'fdm_bank_transfer_cgd_bank_statement_combination'
      and (rule->>'difference_allowed')::numeric <> 0
  ) then
    raise exception
      'Bank Reservation automatic rule requires zero difference allowed.';
  end if;

  if exists (
    select 1
    from jsonb_array_elements(p_rules) rule
    where rule->>'rule_key' =
        'cgd_bank_statement_fdm_adyen_monthly_payments'
      and (rule->>'max_difference_days')::integer <> 31
  ) then
    raise exception
      'Adyen automatic rule requires the fixed 31-day calendar-month property.';
  end if;

  lock table public.financial_reconciliation_source_rules
    in share row exclusive mode;
  lock table public.financial_reconciliation_automatic_rule_configs
    in share row exclusive mode;
  lock table public.financial_reconciliation_automatic_schedule
    in share row exclusive mode;
  set constraints financial_reconciliation_automatic_rule_configs_priority_key
    deferred;

  if (select count(*)
      from public.financial_reconciliation_automatic_rule_configs) <> 7
    or exists (
      select 1
      from public.financial_reconciliation_automatic_rule_configs config
      where not exists (
        select 1
        from jsonb_array_elements(p_rules) rule
        where rule->>'rule_key' = config.rule_key
      )
    ) then
    raise exception 'Automation settings require every managed rule exactly once.';
  end if;

  if exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs config
    where (config.rule_key, config.rule_version) not in (
      ('financial_documents_cgd_bank_statement', 2),
      ('financial_documents_cgd_credit_card', 1),
      ('financial_documents_cgd_bank_statement_amount_only', 1),
      ('financial_documents_cgd_credit_card_amount_only', 1),
      ('cgd_bank_statement_fdm_credit_card_monthly_income', 2),
      ('fdm_bank_transfer_cgd_bank_statement_combination', 1),
      ('cgd_bank_statement_fdm_adyen_monthly_payments', 1)
    )
  ) then
    raise exception
      'Installed automatic reconciliation configuration is not in the managed rule/version allowlist.';
  end if;

  if exists (
    select 1
    from jsonb_array_elements(p_rules) rule
    join public.financial_reconciliation_automatic_rule_configs config
      on config.rule_key = rule->>'rule_key'
    where config.rule_version <> (rule->>'rule_version')::integer
  ) then
    raise exception
      'Submitted automatic rule version does not match managed configuration.';
  end if;

  if exists (
    select 1
    from jsonb_array_elements(p_rules) rule
    join public.financial_reconciliation_automatic_rule_configs config
      on config.rule_key = rule->>'rule_key'
    join public.financial_reconciliation_automatic_rule_definitions definition
      on definition.rule_key = config.rule_key
     and definition.version = config.rule_version
    cross join lateral jsonb_array_elements_text(
      definition.destination_source_types
    ) destination(source_type)
    where (rule->>'enabled')::boolean
      and not exists (
        select 1
        from public.financial_reconciliation_source_rules source_rule
        where source_rule.base_source_type = definition.base_source_type
          and source_rule.matching_source_type = destination.source_type
          and source_rule.operator = case
            when config.rule_key in (
              'cgd_bank_statement_fdm_credit_card_monthly_income',
              'fdm_bank_transfer_cgd_bank_statement_combination',
              'cgd_bank_statement_fdm_adyen_monthly_payments'
            ) then '-'
            else '+'
          end
      )
  ) then
    raise exception
      'No fixed directional source rule exists for an enabled automatic rule.';
  end if;

  update public.financial_reconciliation_automatic_rule_configs config
  set enabled = input.enabled,
      allow_manual_execution = input.allow_manual_execution,
      include_in_scheduled_batch = input.include_in_scheduled_batch,
      difference_allowed = input.difference_allowed::numeric(14,2),
      max_difference_days = input.max_difference_days,
      priority = input.priority,
      updated_by = trim(p_actor),
      updated_at = now()
  from jsonb_to_recordset(p_rules) as input(
    rule_key text,
    rule_version integer,
    enabled boolean,
    allow_manual_execution boolean,
    include_in_scheduled_batch boolean,
    difference_allowed text,
    max_difference_days integer,
    priority integer
  )
  where config.rule_key = input.rule_key;

  update public.financial_reconciliation_automatic_schedule
  set enabled = (p_schedule->>'enabled')::boolean,
      time_of_day = (p_schedule->>'time_of_day')::time,
      time_zone = p_schedule->>'time_zone',
      updated_by = trim(p_actor),
      updated_at = now()
  where id = true;

  return public.get_financial_reconciliation_automation_settings();
end
$$;

revoke all on function public.replace_financial_reconciliation_automation_settings(jsonb,jsonb,text)
  from public, anon, authenticated, service_role;
grant execute on function public.replace_financial_reconciliation_automation_settings(jsonb,jsonb,text)
  to service_role;

notify pgrst, 'reload schema';
