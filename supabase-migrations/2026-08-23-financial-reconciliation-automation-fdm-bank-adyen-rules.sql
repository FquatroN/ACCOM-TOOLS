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

create table if not exists public.financial_reconciliation_automatic_bank_reservation_population (
  run_id uuid not null,
  bank_id uuid not null,
  ordinal integer not null,
  bank_date date not null,
  constraint fr_auto_bank_res_population_run_fkey
    foreign key (run_id)
    references public.financial_reconciliation_automatic_runs(id)
    on delete cascade,
  constraint fr_auto_bank_res_population_ordinal_check
    check (ordinal > 0),
  constraint fr_auto_bank_res_population_pkey
    primary key (run_id, bank_id),
  constraint fr_auto_bank_res_population_run_ordinal_key
    unique (run_id, ordinal)
);

do $migration$
declare
  v_column_count integer;
  v_constraint_count integer;
  v_constraint record;
  v_actual_type "char";
  v_actual_definition text;
  v_expected_type "char";
  v_expected_definition text;
  v_foreign_key record;
  v_actual_index record;
  v_expected_index record;
begin
  select count(*) into v_column_count
  from information_schema.columns column_row
  where column_row.table_schema = 'public'
    and column_row.table_name =
      'financial_reconciliation_automatic_bank_reservation_population';

  if v_column_count is distinct from 4
    or not exists (
      select 1 from information_schema.columns
      where table_schema = 'public'
        and table_name =
          'financial_reconciliation_automatic_bank_reservation_population'
        and ordinal_position = 1 and column_name = 'run_id'
        and data_type = 'uuid' and is_nullable = 'NO'
        and column_default is null
    )
    or not exists (
      select 1 from information_schema.columns
      where table_schema = 'public'
        and table_name =
          'financial_reconciliation_automatic_bank_reservation_population'
        and ordinal_position = 2 and column_name = 'bank_id'
        and data_type = 'uuid' and is_nullable = 'NO'
        and column_default is null
    )
    or not exists (
      select 1 from information_schema.columns
      where table_schema = 'public'
        and table_name =
          'financial_reconciliation_automatic_bank_reservation_population'
        and ordinal_position = 3 and column_name = 'ordinal'
        and data_type = 'integer' and is_nullable = 'NO'
        and column_default is null
    )
    or not exists (
      select 1 from information_schema.columns
      where table_schema = 'public'
        and table_name =
          'financial_reconciliation_automatic_bank_reservation_population'
        and ordinal_position = 4 and column_name = 'bank_date'
        and data_type = 'date' and is_nullable = 'NO'
        and column_default is null
    ) then
    raise exception 'Installed Bank Reservation population columns differ from the required contract.';
  end if;

  select count(*) into v_constraint_count
  from pg_constraint constraint_row
  where constraint_row.conrelid =
    'public.financial_reconciliation_automatic_bank_reservation_population'::regclass;
  if v_constraint_count is distinct from 4 then
    raise exception 'Installed Bank Reservation population constraint count differs from the required contract.';
  end if;

  select constraint_row.contype,
         constraint_row.confrelid,
         constraint_row.confdeltype,
         constraint_row.confupdtype,
         constraint_row.confmatchtype,
         constraint_row.condeferrable,
         constraint_row.condeferred,
         constraint_row.convalidated,
         constraint_row.conkey,
         constraint_row.confkey
  into v_foreign_key
  from pg_constraint constraint_row
  where constraint_row.conrelid =
      'public.financial_reconciliation_automatic_bank_reservation_population'::regclass
    and constraint_row.conname = 'fr_auto_bank_res_population_run_fkey';

  if not found
    or v_foreign_key.contype is distinct from 'f'
    or v_foreign_key.confrelid is distinct from
      'public.financial_reconciliation_automatic_runs'::regclass
    or v_foreign_key.confdeltype is distinct from 'c'
    or v_foreign_key.confupdtype is distinct from 'a'
    or v_foreign_key.confmatchtype is distinct from 's'
    or v_foreign_key.condeferrable is distinct from false
    or v_foreign_key.condeferred is distinct from false
    or v_foreign_key.convalidated is distinct from true
    or v_foreign_key.conkey is distinct from array[
      (select attribute_row.attnum
       from pg_attribute attribute_row
       where attribute_row.attrelid =
         'public.financial_reconciliation_automatic_bank_reservation_population'::regclass
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
    ]::smallint[] then
    raise exception 'Installed Bank Reservation population constraint % differs from the required definition.',
      'fr_auto_bank_res_population_run_fkey';
  end if;

  create temporary table task3_population_expected (
    run_id uuid not null,
    bank_id uuid not null,
    ordinal integer not null,
    bank_date date not null,
    constraint task3_population_expected_ordinal_check
      check (ordinal > 0),
    constraint task3_population_expected_pkey
      primary key (run_id, bank_id),
    constraint task3_population_expected_run_ordinal_key
      unique (run_id, ordinal)
  ) on commit drop;

  for v_constraint in
    select * from (values
      ('fr_auto_bank_res_population_ordinal_check',
       'task3_population_expected_ordinal_check'),
      ('fr_auto_bank_res_population_pkey',
       'task3_population_expected_pkey'),
      ('fr_auto_bank_res_population_run_ordinal_key',
       'task3_population_expected_run_ordinal_key')
    ) expected(actual_name, expected_name)
  loop
    select constraint_row.contype,
           pg_get_constraintdef(constraint_row.oid, true)
    into v_actual_type, v_actual_definition
    from pg_constraint constraint_row
    where constraint_row.conrelid =
        'public.financial_reconciliation_automatic_bank_reservation_population'::regclass
      and constraint_row.conname = v_constraint.actual_name;

    select constraint_row.contype,
           pg_get_constraintdef(constraint_row.oid, true)
    into strict v_expected_type, v_expected_definition
    from pg_constraint constraint_row
    where constraint_row.conrelid = 'task3_population_expected'::regclass
      and constraint_row.conname = v_constraint.expected_name;

    if v_actual_type is distinct from v_expected_type
      or v_actual_definition is distinct from v_expected_definition then
      raise exception 'Installed Bank Reservation population constraint % differs from the required definition.',
        v_constraint.actual_name;
    end if;
  end loop;

  if (select count(*)
      from pg_index index_row
      where index_row.indrelid =
        'public.financial_reconciliation_automatic_bank_reservation_population'::regclass)
      is distinct from 2 then
    raise exception 'Installed Bank Reservation population index count differs from the required contract.';
  end if;

  for v_constraint in
    select * from (values
      ('fr_auto_bank_res_population_pkey',
       'task3_population_expected_pkey'),
      ('fr_auto_bank_res_population_run_ordinal_key',
       'task3_population_expected_run_ordinal_key')
    ) expected(actual_name, expected_name)
  loop
    select index_row.indisunique,
           index_row.indisprimary,
           index_row.indisexclusion,
           index_row.indimmediate,
           index_row.indisvalid,
           index_row.indisready,
           index_row.indislive,
           index_row.indnkeyatts,
           index_row.indnatts,
           array(
             select pg_get_indexdef(index_row.indexrelid, ordinal, true)
             from generate_series(1, index_row.indnatts) ordinal
             order by ordinal
           ) as index_columns
    into v_actual_index
    from pg_index index_row
    where index_row.indexrelid =
      to_regclass('public.' || v_constraint.actual_name);

    select index_row.indisunique,
           index_row.indisprimary,
           index_row.indisexclusion,
           index_row.indimmediate,
           index_row.indisvalid,
           index_row.indisready,
           index_row.indislive,
           index_row.indnkeyatts,
           index_row.indnatts,
           array(
             select pg_get_indexdef(index_row.indexrelid, ordinal, true)
             from generate_series(1, index_row.indnatts) ordinal
             order by ordinal
           ) as index_columns
    into strict v_expected_index
    from pg_index index_row
    where index_row.indexrelid = to_regclass(v_constraint.expected_name);

    if v_actual_index.indisunique is distinct from v_expected_index.indisunique
      or v_actual_index.indisprimary is distinct from v_expected_index.indisprimary
      or v_actual_index.indisexclusion is distinct from v_expected_index.indisexclusion
      or v_actual_index.indimmediate is distinct from v_expected_index.indimmediate
      or v_actual_index.indisvalid is distinct from v_expected_index.indisvalid
      or v_actual_index.indisready is distinct from v_expected_index.indisready
      or v_actual_index.indislive is distinct from v_expected_index.indislive
      or v_actual_index.indnkeyatts is distinct from v_expected_index.indnkeyatts
      or v_actual_index.indnatts is distinct from v_expected_index.indnatts
      or v_actual_index.index_columns is distinct from v_expected_index.index_columns then
      raise exception 'Installed Bank Reservation population index % differs from the required definition.',
        v_constraint.actual_name;
    end if;
  end loop;

  drop table task3_population_expected;
end
$migration$;

alter table public.financial_reconciliation_automatic_bank_reservation_population
  enable row level security;

revoke all on table public.financial_reconciliation_automatic_bank_reservation_population
  from public, anon, authenticated, service_role;

do $migration$
declare
  v_role text;
  v_privilege text;
begin
  if not (select relation.relkind = 'r'
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
    )
    or exists (
      select 1
      from information_schema.table_privileges grant_row
      where grant_row.table_schema = 'public'
        and grant_row.table_name =
          'financial_reconciliation_automatic_bank_reservation_population'
        and grant_row.grantee = 'PUBLIC'
    )
    or exists (
      select 1
      from information_schema.column_privileges grant_row
      where grant_row.table_schema = 'public'
        and grant_row.table_name =
          'financial_reconciliation_automatic_bank_reservation_population'
        and grant_row.grantee in (
          'PUBLIC','anon','authenticated','service_role'
        )
    ) then
    raise exception 'Installed Bank Reservation population security differs from the required contract.';
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
        raise exception 'Installed Bank Reservation population security grants % to %.',
          v_privilege, v_role;
      end if;
    end loop;
  end loop;
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

create or replace function public.financial_reconciliation_automatic_bank_reservation_count()
returns bigint
language sql
stable
security definer set search_path = public, pg_temp
as $$
  select count(*)
  from public.import_cgd_extrato_ordem bank
  where bank.data is not null
    and bank.data >= date '2026-01-01'
    and bank.montante is not null
    and not exists (
      select 1
      from public.financial_reconciliation_items locked
      where locked.source_type = 'import_cgd_extrato_ordem'
        and locked.source_id = bank.id
    )
$$;

create or replace function public.financial_reconciliation_automatic_bank_reservation_page(
  p_after_date date,
  p_after_id uuid,
  p_limit integer
)
returns table (
  bank_id uuid,
  bank_date date,
  bank_amount numeric
)
language plpgsql
stable
security definer set search_path = public, pg_temp
as $$
begin
  if p_limit is null or p_limit not between 1 and 25 then
    raise exception
      'Automatic Bank Reservation page size must be between 1 and 25.';
  end if;
  if (p_after_date is null) <> (p_after_id is null) then
    raise exception
      'Automatic Bank Reservation page cursor must contain both date and ID.';
  end if;

  return query
  select bank.id, bank.data, bank.montante
  from public.import_cgd_extrato_ordem bank
  where bank.data is not null
    and bank.data >= date '2026-01-01'
    and bank.montante is not null
    and (
      p_after_date is null
      or (bank.data, bank.id) > (p_after_date, p_after_id)
    )
    and not exists (
      select 1
      from public.financial_reconciliation_items locked
      where locked.source_type = 'import_cgd_extrato_ordem'
        and locked.source_id = bank.id
    )
  order by bank.data, bank.id
  limit p_limit;
end
$$;

create or replace function public.financial_reconciliation_automatic_bank_reservation_groups(
  p_bank_id uuid,
  p_max_difference_days integer,
  p_candidate_pool_limit integer,
  p_state_limit integer,
  p_evidence_limit integer
)
returns jsonb
language plpgsql
stable
security definer set search_path = public, pg_temp
as $$
declare
  v_bank_date date;
  v_bank_amount_cents bigint;
  v_bank_record jsonb;
  v_target_cents bigint;
  v_candidate_ids uuid[];
  v_candidate_dates date[];
  v_candidate_cents bigint[];
  v_candidate_count integer := 0;
  v_path integer[] := '{}'::integer[];
  v_depth integer;
  v_current_position integer;
  v_next_position integer;
  v_prefix_total_cents bigint;
  v_path_total_cents bigint := 0;
  v_path_signed_total_cents bigint;
  v_evaluated_states integer := 0;
  v_qualifying_count integer := 0;
  v_candidate_groups jsonb := '[]'::jsonb;
  v_group_ids uuid[];
  v_group_records jsonb;
  v_descended boolean;
  v_advanced boolean;
  v_state_limited boolean := false;
  v_group_limited boolean := false;
begin
  if p_bank_id is null then
    raise exception 'Automatic Bank Reservation Bank ID is required.';
  end if;
  if p_max_difference_days is null
    or p_max_difference_days not between 0 and 90 then
    raise exception
      'Automatic Bank Reservation day limit must be between 0 and 90.';
  end if;
  if p_candidate_pool_limit is null
    or p_candidate_pool_limit not between 1 and 60 then
    raise exception
      'Automatic Bank Reservation candidate limit must be between 1 and 60.';
  end if;
  if p_state_limit is null or p_state_limit not between 1 and 250000 then
    raise exception
      'Automatic Bank Reservation state limit must be between 1 and 250000.';
  end if;
  if p_evidence_limit is null or p_evidence_limit not between 1 and 12 then
    raise exception
      'Automatic Bank Reservation evidence limit must be between 1 and 12.';
  end if;

  select
    bank.data,
    round(bank.montante * 100)::bigint,
    jsonb_build_object(
      'sourceType', 'import_cgd_extrato_ordem',
      'sourceId', bank.id,
      'sourceDate', bank.data,
      'amount', bank.montante,
      'description', bank.descritivo,
      'account', '',
      'rowSnapshot', to_jsonb(bank)
    )
  into v_bank_date, v_bank_amount_cents, v_bank_record
  from public.import_cgd_extrato_ordem bank
  where bank.id = p_bank_id
    and bank.data is not null
    and bank.data >= date '2026-01-01'
    and bank.montante is not null
    and not exists (
      select 1
      from public.financial_reconciliation_items locked
      where locked.source_type = 'import_cgd_extrato_ordem'
        and locked.source_id = bank.id
    );
  if not found then
    raise exception 'Automatic Bank Reservation Bank anchor is not eligible.';
  end if;
  v_target_cents := abs(v_bank_amount_cents);

  select
    array_agg(candidate.id order by candidate.event_date, candidate.id),
    array_agg(candidate.event_date order by candidate.event_date, candidate.id),
    array_agg(candidate.amount_cents order by candidate.event_date, candidate.id)
  into
    v_candidate_ids,
    v_candidate_dates,
    v_candidate_cents
  from (
    select
      fdm.id,
      fdm.event_date,
      fdm.amount,
      round(fdm.amount * 100)::bigint as amount_cents
    from public.import_fdm_accounts fdm
    where fdm.account = 'Bank Transfer'
      and fdm.event_date is not null
      and fdm.event_date >= date '2026-01-01'
      and fdm.event_date between
          v_bank_date - p_max_difference_days
          and v_bank_date + p_max_difference_days
      and fdm.amount is not null
      and round(fdm.amount * 100)::bigint <> 0
      and sign(round(fdm.amount * 100)::bigint) =
          -sign(v_bank_amount_cents)
      and abs(round(fdm.amount * 100)::bigint) <= v_target_cents
      and not exists (
        select 1
        from public.financial_reconciliation_items locked
        where locked.source_type = 'import_fdm_accounts'
          and locked.source_id = fdm.id
      )
    order by fdm.event_date, fdm.id
    limit p_candidate_pool_limit + 1
  ) candidate;

  v_candidate_count := coalesce(cardinality(v_candidate_ids), 0);
  if v_candidate_count > p_candidate_pool_limit then
    return jsonb_build_object(
      'classification', 'ambiguous',
      'reason', 'candidate_limit',
      'evaluatedStates', 0,
      'candidateCount', v_candidate_count,
      'canonicalFdmId', v_candidate_ids[1],
      'canonicalFdmDate', v_candidate_dates[1],
      'candidateGroups', v_candidate_groups
    );
  end if;

  if v_candidate_count > 0 and v_target_cents > 0 then
    v_path := array[1];
    v_path_total_cents := abs(v_candidate_cents[1]);
  end if;

  while cardinality(v_path) > 0 loop
    if v_evaluated_states >= p_state_limit then
      v_state_limited := true;
      exit;
    end if;
    v_evaluated_states := v_evaluated_states + 1;

    if v_path_total_cents = v_target_cents then
      v_path_signed_total_cents := case
        when v_candidate_cents[v_path[1]] < 0 then -v_path_total_cents
        else v_path_total_cents
      end;
      v_qualifying_count := v_qualifying_count + 1;
      if v_qualifying_count > p_evidence_limit then
        v_group_limited := true;
        exit;
      end if;

      select array_agg(v_candidate_ids[path.position] order by path.ordinality)
      into v_group_ids
      from unnest(v_path) with ordinality path(position, ordinality);

      select coalesce(jsonb_agg(jsonb_build_object(
        'sourceType', 'import_fdm_accounts',
        'sourceId', fdm.id,
        'sourceDate', fdm.event_date,
        'amount', fdm.amount,
        'description', fdm.description,
        'account', fdm.account,
        'rowSnapshot', to_jsonb(fdm)
      ) order by fdm.event_date, fdm.id), '[]'::jsonb)
      into v_group_records
      from public.import_fdm_accounts fdm
      where fdm.id = any(v_group_ids);

      v_candidate_groups := v_candidate_groups || jsonb_build_array(
        jsonb_build_object(
          'fdmIds', to_jsonb(v_group_ids),
          'fdmTotalCents', v_path_signed_total_cents,
          'bankAmountCents', v_bank_amount_cents,
          'equationCents', v_path_signed_total_cents + v_bank_amount_cents,
          'fdmRecords', v_group_records,
          'bankRecord', v_bank_record
        )
      );
    end if;

    v_descended := false;
    v_depth := cardinality(v_path);
    v_current_position := v_path[v_depth];
    if v_path_total_cents < v_target_cents
      and cardinality(v_path) < 10
      and v_current_position < v_candidate_count then
      for v_next_position in v_current_position + 1..v_candidate_count loop
        if v_path_total_cents + abs(v_candidate_cents[v_next_position])
            <= v_target_cents then
          v_path := array_append(v_path, v_next_position);
          v_path_total_cents :=
            v_path_total_cents + abs(v_candidate_cents[v_next_position]);
          v_descended := true;
          exit;
        end if;
      end loop;
    end if;
    if v_descended then
      continue;
    end if;

    v_advanced := false;
    while cardinality(v_path) > 0 loop
      v_depth := cardinality(v_path);
      v_current_position := v_path[v_depth];
      v_prefix_total_cents :=
        v_path_total_cents - abs(v_candidate_cents[v_current_position]);

      if v_current_position < v_candidate_count then
        for v_next_position in v_current_position + 1..v_candidate_count loop
          if v_prefix_total_cents + abs(v_candidate_cents[v_next_position])
              <= v_target_cents then
            v_path[v_depth] := v_next_position;
            v_path_total_cents :=
              v_prefix_total_cents + abs(v_candidate_cents[v_next_position]);
            v_advanced := true;
            exit;
          end if;
        end loop;
      end if;
      if v_advanced then
        exit;
      end if;

      if v_depth = 1 then
        v_path := '{}'::integer[];
        v_path_total_cents := 0;
      else
        v_path := v_path[1:v_depth - 1];
        v_path_total_cents := v_prefix_total_cents;
      end if;
    end loop;
  end loop;

  if v_state_limited or v_group_limited then
    return jsonb_build_object(
      'classification', 'ambiguous',
      'reason', 'candidate_limit',
      'evaluatedStates', v_evaluated_states,
      'candidateCount', v_candidate_count,
      'canonicalFdmId', coalesce(
        (v_candidate_groups#>>'{0,fdmIds,0}')::uuid,
        v_candidate_ids[1]
      ),
      'canonicalFdmDate', (
        select fdm.event_date
        from public.import_fdm_accounts fdm
        where fdm.id = coalesce(
          (v_candidate_groups#>>'{0,fdmIds,0}')::uuid,
          v_candidate_ids[1]
        )
      ),
      'candidateGroups', v_candidate_groups
    );
  elsif v_qualifying_count = 0 then
    return jsonb_build_object(
      'classification', 'skipped',
      'reason', 'no_qualifying_combination',
      'evaluatedStates', v_evaluated_states,
      'candidateCount', v_candidate_count,
      'canonicalFdmId', v_candidate_ids[1],
      'canonicalFdmDate', v_candidate_dates[1],
      'candidateGroups', v_candidate_groups
    );
  elsif v_qualifying_count = 1 then
    return jsonb_build_object(
      'classification', 'proposed',
      'reason', 'unique_qualifying_combination',
      'evaluatedStates', v_evaluated_states,
      'candidateCount', v_candidate_count,
      'canonicalFdmId', (v_candidate_groups#>>'{0,fdmIds,0}')::uuid,
      'canonicalFdmDate', (
        select fdm.event_date
        from public.import_fdm_accounts fdm
        where fdm.id = (v_candidate_groups#>>'{0,fdmIds,0}')::uuid
      ),
      'candidateGroups', v_candidate_groups
    );
  end if;

  return jsonb_build_object(
    'classification', 'ambiguous',
    'reason', 'multiple_qualifying_combinations',
    'evaluatedStates', v_evaluated_states,
    'candidateCount', v_candidate_count,
    'canonicalFdmId', (v_candidate_groups#>>'{0,fdmIds,0}')::uuid,
    'canonicalFdmDate', (
      select fdm.event_date
      from public.import_fdm_accounts fdm
      where fdm.id = (v_candidate_groups#>>'{0,fdmIds,0}')::uuid
    ),
    'candidateGroups', v_candidate_groups
  );
end
$$;

create or replace function public.financial_reconciliation_continue_automatic_bank_reservation(
  p_run_id uuid,
  p_actor text
)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
declare
  v_run public.financial_reconciliation_automatic_runs%rowtype;
  v_rule jsonb;
  v_bank record;
  v_groups jsonb;
  v_candidate_groups jsonb;
  v_classification text;
  v_reason text;
  v_canonical_fdm_id uuid;
  v_canonical_fdm_date date;
  v_base_snapshot jsonb;
  v_bank_snapshot jsonb;
  v_summary_snapshot jsonb;
  v_signature text;
  v_proposal_id uuid;
  v_inserted boolean;
  v_expected_source_count integer;
  v_inserted_source_count integer;
  v_page_count integer := 0;
  v_population_total integer := 0;
  v_population_max_ordinal integer;
  v_population_invalid boolean := false;
  v_population_error_summary text;
  v_invalid_bank_id uuid;
  v_last_ordinal integer;
  v_last_date date;
  v_last_id uuid;
  v_processed_after integer;
begin
  if nullif(trim(coalesce(p_actor, '')), '') is null then
    raise exception 'Actor is required.';
  end if;

  select * into v_run
  from public.financial_reconciliation_automatic_runs run
  where run.id = p_run_id
  for update;
  if not found then
    raise exception 'Automatic analysis run was not found.';
  end if;
  if v_run.actor <> p_actor then
    raise exception 'Automatic analysis run belongs to another actor.';
  end if;
  if v_run.analysis_completed_at is not null then
    return public.get_financial_reconciliation_automatic_run(p_run_id);
  end if;
  if v_run.status <> 'analyzing' then
    raise exception 'Automatic analysis run is not resumable.';
  end if;
  if jsonb_array_length(v_run.definition_config_snapshot) <> 1 then
    raise exception
      'Resumable automatic analysis requires exactly one snapshotted rule.';
  end if;
  select snapshot.value into strict v_rule
  from jsonb_array_elements(v_run.definition_config_snapshot) snapshot(value);

  if jsonb_typeof(v_rule) is distinct from 'object'
    or v_rule - array[
      'ruleKey','ruleVersion','displayName','priority','differenceAllowed',
      'maxDifferenceDays','destinationSourceType','definition','operator'
    ]::text[] <> '{}'::jsonb
    or not (v_rule ?& array[
      'ruleKey','ruleVersion','displayName','priority','differenceAllowed',
      'maxDifferenceDays','destinationSourceType','definition','operator'
    ])
    or v_rule->>'ruleKey' is distinct from
      'fdm_bank_transfer_cgd_bank_statement_combination'
    or v_rule->>'ruleVersion' is distinct from '1'
    or v_rule->>'displayName' is distinct from
      'FDM Accounts – Bank Reservation Payments'
    or coalesce(v_rule->>'priority', '') !~ '^[1-9][0-9]*$'
    or length(v_rule->>'priority') > 10
    or v_rule->>'destinationSourceType' is distinct from
      'import_cgd_extrato_ordem'
    or v_rule->>'operator' is distinct from '-'
    or coalesce(v_rule->>'differenceAllowed', '') !~ '^0(?:\.0+)?$'
    or coalesce(v_rule->>'maxDifferenceDays', '') !~ '^[0-9]+$'
    or (v_rule->>'maxDifferenceDays')::integer not between 0 and 90
    or v_rule->'definition' is distinct from jsonb_build_object(
      'strategy', 'bounded_exact_combination',
      'sourceAccount', 'Bank Transfer',
      'maxSourceRecords', 10,
      'candidatePoolLimit', 60,
      'stateLimit', 250000,
      'evidenceGroupLimit', 12,
      'amountMode', 'signed_integer_cents',
      'dateMode', 'inclusive_days'
    ) then
    raise exception 'Automatic Bank Reservation rule snapshot contract is invalid.';
  end if;

  lock table
    public.import_cgd_extrato_ordem,
    public.import_fdm_accounts,
    public.financial_reconciliation_items
  in share row exclusive mode;

  select count(*)::integer, max(population.ordinal)
  into v_population_total, v_population_max_ordinal
  from public.financial_reconciliation_automatic_bank_reservation_population population
  where population.run_id = p_run_id;

  if v_population_total = 0 then
    if v_run.analysis_cursor_date is not null
      or v_run.analysis_cursor_id is not null
      or v_run.analysis_processed <> 0 then
      v_population_invalid := true;
      v_population_error_summary :=
        'The Bank Reservation analysis population projection is missing.';
    else
      insert into public.financial_reconciliation_automatic_bank_reservation_population (
        run_id, bank_id, ordinal, bank_date
      )
      select
        p_run_id,
        bank.id,
        row_number() over (order by bank.data, bank.id)::integer,
        bank.data
      from public.import_cgd_extrato_ordem bank
      where bank.data is not null
        and bank.data >= date '2026-01-01'
        and bank.montante is not null
        and not exists (
          select 1
          from public.financial_reconciliation_items locked
          where locked.source_type = 'import_cgd_extrato_ordem'
            and locked.source_id = bank.id
        )
      order by bank.data, bank.id;
      get diagnostics v_population_total = row_count;
      v_population_max_ordinal := nullif(v_population_total, 0);

      if v_run.analysis_total > v_population_total then
        v_population_invalid := true;
        v_population_error_summary :=
          'The Bank Reservation analysis population changed before projection.';
      else
        update public.financial_reconciliation_automatic_runs run
        set analysis_total = v_population_total,
            updated_at = now()
        where run.id = p_run_id;
        v_run.analysis_total := v_population_total;
      end if;
    end if;
  elsif v_population_max_ordinal is distinct from v_population_total
    or v_population_total is distinct from v_run.analysis_total
    or v_run.analysis_processed not between 0 and v_population_total
    or (v_run.analysis_processed = 0 and (
      v_run.analysis_cursor_date is not null
      or v_run.analysis_cursor_id is not null
    ))
    or (v_run.analysis_processed > 0 and not exists (
      select 1
      from public.financial_reconciliation_automatic_bank_reservation_population population
      where population.run_id = p_run_id
        and population.ordinal = v_run.analysis_processed
        and population.bank_date = v_run.analysis_cursor_date
        and population.bank_id = v_run.analysis_cursor_id
    )) then
    v_population_invalid := true;
    v_population_error_summary :=
      'The Bank Reservation analysis population projection diverged.';
  end if;

  if not v_population_invalid then
    select population.bank_id
    into v_invalid_bank_id
    from public.financial_reconciliation_automatic_bank_reservation_population population
    left join public.import_cgd_extrato_ordem bank
      on bank.id = population.bank_id
    where population.run_id = p_run_id
      and (
        bank.id is null
        or bank.data is distinct from population.bank_date
        or bank.data < date '2026-01-01'
        or bank.montante is null
        or exists (
          select 1
          from public.financial_reconciliation_items locked
          where locked.source_type = 'import_cgd_extrato_ordem'
            and locked.source_id = population.bank_id
        )
      )
    order by population.ordinal
    limit 1;
    if found then
      v_population_invalid := true;
      v_population_error_summary :=
        'The Bank Reservation analysis population changed before completion.';
    end if;
  end if;

  if v_population_invalid then
    update public.financial_reconciliation_automatic_proposals proposal
    set status = 'stale',
        reason = 'analysis_population_changed',
        updated_at = now()
    where proposal.run_id = p_run_id
      and proposal.status in ('proposed','ambiguous');

    update public.financial_reconciliation_automatic_runs run
    set status = 'failed',
        error_summary = v_population_error_summary,
        analysis_completed_at = coalesce(run.analysis_completed_at, now()),
        analysis_error_code = 'analysis_population_changed',
        analysis_error_at = now(),
        finished_at = coalesce(run.finished_at, now()),
        updated_at = now(),
        counts = jsonb_build_object(
          'bases', run.analysis_total,
          'proposed', 0,
          'ambiguous', 0,
          'skipped', greatest(
            run.analysis_processed - (
              select count(*)::integer
              from public.financial_reconciliation_automatic_proposals proposal
              where proposal.run_id = p_run_id
                and proposal.status in ('stale','failed')
            ),
            0
          ),
          'stale', (
            select count(*)
            from public.financial_reconciliation_automatic_proposals proposal
            where proposal.run_id = p_run_id
              and proposal.status = 'stale'
          )
        )
    where run.id = p_run_id;
    return public.get_financial_reconciliation_automatic_run(p_run_id);
  end if;

  for v_bank in
    select
      population.ordinal,
      bank.id as bank_id,
      bank.data as bank_date,
      bank.montante as bank_amount
    from public.financial_reconciliation_automatic_bank_reservation_population population
    left join public.import_cgd_extrato_ordem bank
      on bank.id = population.bank_id
    where population.run_id = p_run_id
      and population.ordinal > v_run.analysis_processed
    order by population.ordinal
    limit 25
  loop
    v_page_count := v_page_count + 1;
    v_last_ordinal := v_bank.ordinal;
    v_last_date := v_bank.bank_date;
    v_last_id := v_bank.bank_id;
    v_groups :=
      public.financial_reconciliation_automatic_bank_reservation_groups(
        v_bank.bank_id,
        (v_rule->>'maxDifferenceDays')::integer,
        60,
        250000,
        12
      );
    v_classification := v_groups->>'classification';
    v_reason := v_groups->>'reason';

    if v_classification = 'skipped' then
      continue;
    end if;
    if v_classification not in ('proposed', 'ambiguous')
      or v_reason not in (
        'unique_qualifying_combination',
        'multiple_qualifying_combinations',
        'candidate_limit'
      ) then
      raise exception 'Automatic Bank Reservation search classification is invalid.';
    end if;

    v_candidate_groups := v_groups->'candidateGroups';
    v_canonical_fdm_id := (v_groups->>'canonicalFdmId')::uuid;
    v_canonical_fdm_date := (v_groups->>'canonicalFdmDate')::date;
    if v_canonical_fdm_id is null or v_canonical_fdm_date is null then
      raise exception 'Automatic Bank Reservation canonical FDM member is missing.';
    end if;

    select jsonb_build_object(
      'sourceType', 'import_fdm_accounts',
      'sourceId', fdm.id,
      'sourceDate', fdm.event_date,
      'amount', fdm.amount,
      'description', fdm.description,
      'account', fdm.account,
      'rowSnapshot', to_jsonb(fdm)
    )
    into strict v_base_snapshot
    from public.import_fdm_accounts fdm
    where fdm.id = v_canonical_fdm_id
      and fdm.event_date = v_canonical_fdm_date
      and fdm.account = 'Bank Transfer'
      and fdm.amount is not null;

    select jsonb_build_object(
      'sourceType', 'import_cgd_extrato_ordem',
      'sourceId', bank.id,
      'sourceDate', bank.data,
      'amount', bank.montante,
      'description', bank.descritivo,
      'account', '',
      'rowSnapshot', to_jsonb(bank)
    )
    into strict v_bank_snapshot
    from public.import_cgd_extrato_ordem bank
    where bank.id = v_bank.bank_id
      and bank.data = v_bank.bank_date
      and bank.montante = v_bank.bank_amount;

    v_signature := public.financial_reconciliation_extension_sha256(
      jsonb_build_object(
        'ruleKey', 'fdm_bank_transfer_cgd_bank_statement_combination',
        'ruleVersion', 1,
        'bankId', v_bank.bank_id,
        'bankDate', v_bank.bank_date,
        'bankAmount', v_bank.bank_amount,
        'classification', v_classification,
        'reason', v_reason,
        'candidateGroups', v_candidate_groups,
        'operator', '-',
        'maxDifferenceDays', (v_rule->>'maxDifferenceDays')::integer
      )::text
    );
    v_summary_snapshot := jsonb_build_object(
      'ruleKey', 'fdm_bank_transfer_cgd_bank_statement_combination',
      'ruleVersion', 1,
      'bankAnchor', v_bank_snapshot,
      'candidateGroups', v_candidate_groups,
      'classification', v_classification,
      'reason', v_reason,
      'evaluatedStates', (v_groups->>'evaluatedStates')::integer,
      'candidateCount', (v_groups->>'candidateCount')::integer,
      'operator', '-',
      'differenceAllowed', 0,
      'maxDifferenceDays', (v_rule->>'maxDifferenceDays')::integer,
      'maxSourceRecords', 10,
      'analysisTimestamp', now(),
      'signature', v_signature
    );

    v_proposal_id := null;
    insert into public.financial_reconciliation_automatic_proposals (
      run_id, rule_key, rule_version, base_source_type,
      base_source_id, base_source_date, base_snapshot, items, evidence,
      candidate_groups, calculated_difference, allowed_difference,
      status, reason, signature, grouping_key, summary_snapshot
    ) values (
      p_run_id, 'fdm_bank_transfer_cgd_bank_statement_combination', 1,
      'import_fdm_accounts', v_canonical_fdm_id, v_canonical_fdm_date,
      v_base_snapshot, '[]'::jsonb, v_candidate_groups,
      v_candidate_groups, 0, 0, v_classification, v_reason, v_signature,
      v_bank.bank_id::text, v_summary_snapshot
    )
    on conflict do nothing
    returning id into v_proposal_id;
    v_inserted := v_proposal_id is not null;

    if not v_inserted then
      select proposal.id into strict v_proposal_id
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = p_run_id
        and proposal.rule_key =
          'fdm_bank_transfer_cgd_bank_statement_combination'
        and proposal.rule_version = 1
        and proposal.base_source_type = 'import_fdm_accounts'
        and proposal.base_source_id = v_canonical_fdm_id
        and proposal.signature = v_signature;
    else
      with selected_ids as (
        select distinct selected.id
        from (
          select fdm_id.value::uuid as id
          from jsonb_array_elements(v_candidate_groups) candidate_group(value)
          cross join lateral jsonb_array_elements_text(
            candidate_group.value->'fdmIds'
          )
            fdm_id(value)
          union all
          select v_canonical_fdm_id
        ) selected
      ), source_members as (
        select
          fdm.*,
          row_number() over (order by fdm.event_date, fdm.id)::integer
            as ordinal
        from public.import_fdm_accounts fdm
        join selected_ids selected on selected.id = fdm.id
      )
      insert into public.financial_reconciliation_automatic_proposal_memberships (
        proposal_id, role, source_type, source_id, ordinal, source_date,
        amount, description, account, row_snapshot
      )
      select
        v_proposal_id, 'source', 'import_fdm_accounts', source_member.id,
        source_member.ordinal, source_member.event_date, source_member.amount,
        source_member.description, source_member.account,
        to_jsonb(source_member) - 'ordinal'
      from source_members source_member
      order by source_member.ordinal;
      get diagnostics v_inserted_source_count = row_count;

      select count(distinct selected.id)::integer
      into v_expected_source_count
      from (
        select fdm_id.value::uuid as id
        from jsonb_array_elements(v_candidate_groups) candidate_group(value)
        cross join lateral jsonb_array_elements_text(
          candidate_group.value->'fdmIds'
        )
          fdm_id(value)
        union all
        select v_canonical_fdm_id
      ) selected;
      if v_inserted_source_count <> v_expected_source_count then
        raise exception
          'Automatic Bank Reservation source membership insert was incomplete.';
      end if;

      insert into public.financial_reconciliation_automatic_proposal_memberships (
        proposal_id, role, source_type, source_id, ordinal, source_date,
        amount, description, account, row_snapshot
      )
      select
        v_proposal_id, 'destination', 'import_cgd_extrato_ordem', bank.id,
        1, bank.data, bank.montante, bank.descritivo, '', to_jsonb(bank)
      from public.import_cgd_extrato_ordem bank
      where bank.id = v_bank.bank_id;

      if (select member.source_id
          from public.financial_reconciliation_automatic_proposal_memberships member
          where member.proposal_id = v_proposal_id
            and member.role = 'source'
          order by member.ordinal
          limit 1) is distinct from v_canonical_fdm_id
        or (select count(*)
            from public.financial_reconciliation_automatic_proposal_memberships member
            where member.proposal_id = v_proposal_id
              and member.role = 'destination') <> 1 then
        raise exception
          'Automatic Bank Reservation proposal and memberships diverged.';
      end if;
    end if;
  end loop;

  if v_page_count > 0 then
    update public.financial_reconciliation_automatic_runs run
    set analysis_cursor_date = v_last_date,
        analysis_cursor_id = v_last_id,
        analysis_processed = greatest(
          run.analysis_processed,
          v_last_ordinal
        ),
        analysis_total = v_population_total,
        updated_at = now(),
        analysis_error_code = null,
        analysis_error_at = null
    where run.id = p_run_id
      and run.analysis_processed < v_last_ordinal;
  end if;

  select run.analysis_processed into strict v_processed_after
  from public.financial_reconciliation_automatic_runs run
  where run.id = p_run_id;

  if v_processed_after < v_population_total then
    return public.financial_reconciliation_automatic_progress_or_run(p_run_id);
  end if;

  with shared_members as (
    select member.source_type, member.source_id
    from public.financial_reconciliation_automatic_proposal_memberships member
    join public.financial_reconciliation_automatic_proposals proposal
      on proposal.id = member.proposal_id
    where proposal.run_id = p_run_id
      and proposal.rule_key =
        'fdm_bank_transfer_cgd_bank_statement_combination'
      and proposal.rule_version = 1
      and proposal.status = 'proposed'
    group by member.source_type, member.source_id
    having count(distinct member.proposal_id) > 1
  ), affected as (
    select distinct proposal.id
    from public.financial_reconciliation_automatic_proposals proposal
    join public.financial_reconciliation_automatic_proposal_memberships member
      on member.proposal_id = proposal.id
    join shared_members shared
      on shared.source_type = member.source_type
     and shared.source_id = member.source_id
    where proposal.run_id = p_run_id
      and proposal.rule_key =
        'fdm_bank_transfer_cgd_bank_statement_combination'
      and proposal.rule_version = 1
      and proposal.status = 'proposed'
  )
  update public.financial_reconciliation_automatic_proposals proposal
  set status = 'ambiguous',
      reason = 'overlapping_records',
      updated_at = now()
  where proposal.id in (select affected.id from affected);

  if v_processed_after is distinct from v_population_total then
    raise exception
      'Automatic Bank Reservation population processing did not reach its exact total.';
  end if;

  update public.financial_reconciliation_automatic_runs run
  set status = case when exists (
        select 1
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = p_run_id
          and proposal.status in ('proposed','ambiguous','stale','failed')
      ) then 'ready' else 'completed' end,
      finished_at = case when exists (
        select 1
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = p_run_id
          and proposal.status in ('proposed','ambiguous','stale','failed')
      ) then null else now() end,
      analysis_completed_at = now(),
      updated_at = now(),
      analysis_error_code = null,
      analysis_error_at = null,
      counts = (
        select jsonb_build_object(
          'bases', v_population_total,
          'proposed', count(*) filter (where proposal.status = 'proposed'),
          'ambiguous', count(*) filter (where proposal.status = 'ambiguous'),
          'skipped', greatest(
            v_processed_after - count(*) filter (
              where proposal.status in ('proposed','ambiguous','stale','failed')
            ),
            0
          )
        )
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = p_run_id
      )
  where run.id = p_run_id and run.analysis_completed_at is null;

  return public.get_financial_reconciliation_automatic_run(p_run_id);
end
$$;

do $migration$
begin
  if to_regprocedure(
      'public.financial_reconciliation_continue_automatic_prior_analysis(uuid,text)'
    ) is null then
    alter function public.continue_financial_reconciliation_automatic_analysis(uuid,text)
      rename to financial_reconciliation_continue_automatic_prior_analysis;
  end if;
end
$migration$;

create or replace function public.continue_financial_reconciliation_automatic_analysis(
  p_run_id uuid,
  p_actor text
)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
declare
  v_run public.financial_reconciliation_automatic_runs%rowtype;
  v_rule jsonb;
  v_rule_key text;
  v_rule_version integer;
begin
  if nullif(trim(coalesce(p_actor, '')), '') is null then
    raise exception 'Actor is required.';
  end if;

  select * into v_run
  from public.financial_reconciliation_automatic_runs run
  where run.id = p_run_id
  for update;
  if not found then
    raise exception 'Automatic analysis run was not found.';
  end if;
  if v_run.actor <> p_actor then
    raise exception 'Automatic analysis run belongs to another actor.';
  end if;
  if v_run.analysis_completed_at is not null then
    return public.get_financial_reconciliation_automatic_run(p_run_id);
  end if;
  if v_run.status <> 'analyzing' then
    raise exception 'Automatic analysis run is not resumable.';
  end if;

  begin
    if jsonb_array_length(v_run.definition_config_snapshot) <> 1 then
      raise exception
        'Resumable automatic analysis requires exactly one snapshotted rule.';
    end if;
    select snapshot.value into strict v_rule
    from jsonb_array_elements(v_run.definition_config_snapshot) snapshot(value);
    if coalesce(v_rule->>'ruleVersion', '') !~ '^[0-9]+$'
      or length(v_rule->>'ruleVersion') > 10 then
      raise exception 'Automatic reconciliation rule is unsupported.';
    end if;
    v_rule_key := v_rule->>'ruleKey';
    v_rule_version := (v_rule->>'ruleVersion')::integer;

    if (v_rule_key, v_rule_version) in (
        ('financial_documents_cgd_bank_statement', 2),
        ('financial_documents_cgd_credit_card', 1),
        ('financial_documents_cgd_bank_statement_amount_only', 1),
        ('financial_documents_cgd_credit_card_amount_only', 1),
        ('cgd_bank_statement_fdm_credit_card_monthly_income', 2)
      ) then
      return public.financial_reconciliation_continue_automatic_prior_analysis(
        p_run_id,
        p_actor
      );
    elsif v_rule_key = 'fdm_bank_transfer_cgd_bank_statement_combination'
      and v_rule_version = 1 then
      return public.financial_reconciliation_continue_automatic_bank_reservation(
        p_run_id,
        p_actor
      );
    end if;

    raise exception 'Automatic reconciliation rule is unsupported.';
  exception when others then
    update public.financial_reconciliation_automatic_runs run
    set status = 'failed',
        error_summary = 'Automatic analysis could not be completed.',
        analysis_error_code = 'analysis_continuation_failed',
        analysis_error_at = now(),
        finished_at = coalesce(run.finished_at, now()),
        updated_at = now()
    where run.id = p_run_id;
    return public.get_financial_reconciliation_automatic_run(p_run_id);
  end;
end
$$;

revoke all on function public.financial_reconciliation_automatic_bank_reservation_count()
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_automatic_bank_reservation_page(date,uuid,integer)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_automatic_bank_reservation_groups(uuid,integer,integer,integer,integer)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_continue_automatic_bank_reservation(uuid,text)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_continue_automatic_prior_analysis(uuid,text)
  from public, anon, authenticated, service_role;
revoke all on function public.continue_financial_reconciliation_automatic_analysis(uuid,text)
  from public, anon, authenticated, service_role;
grant execute on function public.continue_financial_reconciliation_automatic_analysis(uuid,text)
  to service_role;

notify pgrst, 'reload schema';
