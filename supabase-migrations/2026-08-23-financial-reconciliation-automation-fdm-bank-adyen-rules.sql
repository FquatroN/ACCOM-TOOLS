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

-- Bank Reservation uses opposite-signed FDM and Bank amounts, so its
-- authoritative directional rule is additive. Adyen retains the historical
-- Bank-to-FDM subtraction rule.
insert into public.financial_reconciliation_source_rules (
  base_source_type, matching_source_type, operator
) values (
  'import_fdm_accounts', 'import_cgd_extrato_ordem', '+'
)
on conflict (base_source_type, matching_source_type)
do update set operator = excluded.operator;

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
            when config.rule_key =
              'fdm_bank_transfer_cgd_bank_statement_combination' then '+'
            when config.rule_key in (
              'cgd_bank_statement_fdm_credit_card_monthly_income',
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
  v_source_count integer;
  v_source_total numeric(14,2);
  v_destination_count integer;
  v_destination_total numeric(14,2);
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
  v_operator text;
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

  select source_rule.operator into v_operator
  from public.financial_reconciliation_source_rules source_rule
  where source_rule.base_source_type = 'import_fdm_accounts'
    and source_rule.matching_source_type = 'import_cgd_extrato_ordem'
  for share of source_rule;

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
    or v_operator is distinct from '+'
    or v_rule->>'operator' is distinct from v_operator
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

    with selected_ids as (
      select distinct selected.id
      from (
        select fdm_id.value::uuid as id
        from jsonb_array_elements(v_candidate_groups) candidate_group(value)
        cross join lateral jsonb_array_elements_text(
          candidate_group.value->'fdmIds'
        ) fdm_id(value)
        union all
        select v_canonical_fdm_id
      ) selected
    )
    select count(*)::integer, sum(fdm.amount)::numeric(14,2)
    into v_source_count, v_source_total
    from public.import_fdm_accounts fdm
    join selected_ids selected on selected.id = fdm.id;
    v_destination_count := 1;
    v_destination_total := v_bank.bank_amount;
    if v_source_count is null or v_source_count < 1
      or v_source_total is null or v_destination_total is null then
      raise exception 'Automatic Bank Reservation immutable proposal totals are invalid.';
    end if;

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
        'operator', v_operator,
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
      'bankAnchorDate', v_bank.bank_date,
      'sourceCount', v_source_count,
      'sourceTotal', v_source_total,
      'destinationCount', v_destination_count,
      'destinationTotal', v_destination_total,
      'evaluatedStates', (v_groups->>'evaluatedStates')::integer,
      'candidateCount', (v_groups->>'candidateCount')::integer,
      'operator', v_operator,
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
            and member.role = 'destination') <> v_destination_count
        or (select count(*)::integer
            from public.financial_reconciliation_automatic_proposal_memberships member
            where member.proposal_id = v_proposal_id
              and member.role = 'source') <> v_source_count
        or (select sum(member.amount)::numeric(14,2)
            from public.financial_reconciliation_automatic_proposal_memberships member
            where member.proposal_id = v_proposal_id
              and member.role = 'source') is distinct from v_source_total
        or (select sum(member.amount)::numeric(14,2)
            from public.financial_reconciliation_automatic_proposal_memberships member
            where member.proposal_id = v_proposal_id
              and member.role = 'destination') is distinct from v_destination_total then
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

create table if not exists public.financial_reconciliation_automatic_adyen_population (
  run_id uuid not null,
  calendar_month date not null,
  month_ordinal integer not null,
  role text not null,
  source_type text not null,
  source_id uuid not null,
  member_ordinal integer not null,
  source_date date not null,
  amount numeric(14,2) not null,
  description text not null,
  account text not null,
  row_snapshot jsonb not null,
  constraint fr_auto_adyen_population_run_fkey
    foreign key (run_id)
    references public.financial_reconciliation_automatic_runs(id)
    on delete cascade,
  constraint fr_auto_adyen_population_month_ordinal_check
    check (month_ordinal > 0),
  constraint fr_auto_adyen_population_member_ordinal_check
    check (member_ordinal > 0),
  constraint fr_auto_adyen_population_role_source_check
    check (
      (role = 'source' and source_type = 'import_cgd_extrato_ordem')
      or (role = 'destination' and source_type = 'import_fdm_accounts')
    ),
  constraint fr_auto_adyen_population_month_date_check
    check (
      calendar_month = date_trunc('month', calendar_month)::date
      and source_date >= calendar_month
      and source_date < (calendar_month + interval '1 month')::date
    ),
  constraint fr_auto_adyen_population_snapshot_check
    check (jsonb_typeof(row_snapshot) = 'object'),
  constraint fr_auto_adyen_population_pkey
    primary key (run_id, role, source_id),
  constraint fr_auto_adyen_population_ordinal_key
    unique (run_id, month_ordinal, role, member_ordinal)
);

create or replace function public.financial_reconciliation_continue_automatic_adyen_monthly(
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
  v_month record;
  v_page_count integer := 0;
  v_population_member_total integer := 0;
  v_population_total integer := 0;
  v_population_min_ordinal integer;
  v_population_max_ordinal integer;
  v_population_invalid boolean := false;
  v_population_error_summary text;
  v_invalid_source_id uuid;
  v_last_ordinal integer;
  v_last_month date;
  v_last_cursor_id uuid;
  v_processed_after integer;
  v_allowed_difference numeric(14,2);
  v_operator text;
  v_source_ids uuid[];
  v_destination_ids uuid[];
  v_source_count integer;
  v_destination_count integer;
  v_source_total numeric(14,2);
  v_destination_total numeric(14,2);
  v_difference numeric(14,2);
  v_base_id uuid;
  v_base_date date;
  v_base_snapshot jsonb;
  v_summary_snapshot jsonb;
  v_signature text;
  v_status text;
  v_reason text;
  v_proposal_id uuid;
  v_inserted boolean;
  v_inserted_count integer;
  v_inserted_source_ids uuid[];
  v_inserted_destination_ids uuid[];
  v_inserted_source_total numeric(14,2);
  v_inserted_destination_total numeric(14,2);
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
      'cgd_bank_statement_fdm_adyen_monthly_payments'
    or v_rule->>'ruleVersion' is distinct from '1'
    or v_rule->>'displayName' is distinct from
      'FDM Accounts – Adyen Reservation Payments'
    or coalesce(v_rule->>'priority', '') !~ '^[1-9][0-9]*$'
    or length(v_rule->>'priority') > 10
    or v_rule->>'destinationSourceType' is distinct from
      'import_fdm_accounts'
    or v_rule->>'operator' is distinct from '-'
    or coalesce(v_rule->>'differenceAllowed', '') !~
      '^[0-9]+(\.[0-9]+)?$'
    or length(v_rule->>'differenceAllowed') > 20
    or (v_rule->>'differenceAllowed')::numeric < 0
    or v_rule->>'maxDifferenceDays' is distinct from '31'
    or v_rule->'definition' is distinct from jsonb_build_object(
      'strategy', 'closed_calendar_month',
      'bankDescriptionContains', 'Adyen',
      'fdmAccount', 'Adyen',
      'requiresBothSides', true,
      'monthMarkerDays', 31
    ) then
    raise exception 'Automatic Adyen monthly rule snapshot contract is invalid.';
  end if;

  v_allowed_difference :=
    (v_rule->>'differenceAllowed')::numeric(14,2);
  v_operator := v_rule->>'operator';

  lock table
    public.import_cgd_extrato_ordem,
    public.import_fdm_accounts,
    public.financial_reconciliation_items
  in share row exclusive mode;

  select count(*)::integer,
         count(distinct population.calendar_month)::integer,
         min(population.month_ordinal),
         max(population.month_ordinal)
  into v_population_member_total, v_population_total,
       v_population_min_ordinal, v_population_max_ordinal
  from public.financial_reconciliation_automatic_adyen_population population
  where population.run_id = p_run_id;

  if v_population_member_total = 0 then
    if v_run.analysis_cursor_date is not null
      or v_run.analysis_cursor_id is not null
      or v_run.analysis_processed <> 0 then
      v_population_invalid := true;
      v_population_error_summary :=
        'The Adyen analysis population projection is missing.';
    else
      with bank_eligible as (
        select
          date_trunc('month', bank.data)::date as calendar_month,
          'source'::text as role,
          'import_cgd_extrato_ordem'::text as source_type,
          bank.id as source_id,
          row_number() over (
            partition by date_trunc('month', bank.data)
            order by bank.data, bank.id
          )::integer as member_ordinal,
          bank.data as source_date,
          bank.montante::numeric(14,2) as amount,
          bank.descritivo as description,
          ''::text as account,
          to_jsonb(bank) as row_snapshot
        from public.import_cgd_extrato_ordem bank
        where bank.data >= date '2026-01-01'
          and bank.data < date_trunc('month', current_date)::date
          and bank.montante is not null
          and bank.descritivo ilike '%Adyen%'
          and not exists (
            select 1
            from public.financial_reconciliation_items locked
            where locked.source_type = 'import_cgd_extrato_ordem'
              and locked.source_id = bank.id
          )
      ),
      fdm_eligible as (
        select
          date_trunc('month', fdm.event_date)::date as calendar_month,
          'destination'::text as role,
          'import_fdm_accounts'::text as source_type,
          fdm.id as source_id,
          row_number() over (
            partition by date_trunc('month', fdm.event_date)
            order by fdm.event_date, fdm.id
          )::integer as member_ordinal,
          fdm.event_date as source_date,
          fdm.amount::numeric(14,2) as amount,
          fdm.description,
          fdm.account,
          to_jsonb(fdm) as row_snapshot
        from public.import_fdm_accounts fdm
        where fdm.event_date >= date '2026-01-01'
          and fdm.event_date < date_trunc('month', current_date)::date
          and fdm.amount is not null
          and fdm.account = 'Adyen'
          and not exists (
            select 1
            from public.financial_reconciliation_items locked
            where locked.source_type = 'import_fdm_accounts'
              and locked.source_id = fdm.id
          )
      ),
      eligible as (
        select * from bank_eligible
        union all
        select * from fdm_eligible
      ),
      months as (
        select
          eligible.calendar_month,
          row_number() over (order by eligible.calendar_month)::integer
            as month_ordinal
        from (
          select distinct member.calendar_month
          from eligible member
        ) eligible
      )
      insert into public.financial_reconciliation_automatic_adyen_population (
        run_id, calendar_month, month_ordinal, role, source_type, source_id,
        member_ordinal, source_date, amount, description, account, row_snapshot
      )
      select
        p_run_id, eligible.calendar_month, months.month_ordinal,
        eligible.role, eligible.source_type, eligible.source_id,
        eligible.member_ordinal, eligible.source_date, eligible.amount,
        eligible.description, eligible.account, eligible.row_snapshot
      from eligible
      join months using (calendar_month)
      order by months.month_ordinal, eligible.role, eligible.member_ordinal;
      get diagnostics v_population_member_total = row_count;

      select count(distinct population.calendar_month)::integer,
             min(population.month_ordinal), max(population.month_ordinal)
      into v_population_total, v_population_min_ordinal,
           v_population_max_ordinal
      from public.financial_reconciliation_automatic_adyen_population population
      where population.run_id = p_run_id;

      update public.financial_reconciliation_automatic_runs run
      set analysis_total = v_population_total,
          updated_at = now()
      where run.id = p_run_id;
      v_run.analysis_total := v_population_total;
    end if;
  end if;

  if not v_population_invalid and v_population_member_total > 0 then
    if v_population_min_ordinal is distinct from 1
      or v_population_max_ordinal is distinct from v_population_total
      or v_population_total is distinct from v_run.analysis_total
      or v_run.analysis_processed not between 0 and v_population_total
      or (v_run.analysis_processed = 0 and (
        v_run.analysis_cursor_date is not null
        or v_run.analysis_cursor_id is not null
      ))
      or (v_run.analysis_processed > 0 and (
        not exists (
          select 1
          from public.financial_reconciliation_automatic_adyen_population population
          where population.run_id = p_run_id
            and population.month_ordinal = v_run.analysis_processed
            and population.calendar_month = v_run.analysis_cursor_date
        )
        or v_run.analysis_cursor_id is distinct from (
          select cursor_member.source_id
          from public.financial_reconciliation_automatic_adyen_population cursor_member
          where cursor_member.run_id = p_run_id
            and cursor_member.month_ordinal = v_run.analysis_processed
          order by case cursor_member.role
                     when 'source' then 1 else 2
                   end,
                   cursor_member.member_ordinal,
                   cursor_member.source_id
          limit 1
        )
      ))
      or exists (
        select 1
        from public.financial_reconciliation_automatic_adyen_population population
        where population.run_id = p_run_id
        group by population.month_ordinal
        having count(distinct population.calendar_month) <> 1
      )
      or exists (
        select 1
        from public.financial_reconciliation_automatic_adyen_population population
        where population.run_id = p_run_id
        group by population.calendar_month
        having count(distinct population.month_ordinal) <> 1
      )
      or exists (
        select 1
        from public.financial_reconciliation_automatic_adyen_population population
        where population.run_id = p_run_id
        group by population.calendar_month, population.role
        having min(population.member_ordinal) <> 1
          or max(population.member_ordinal) <> count(*)
      ) then
      v_population_invalid := true;
      v_population_error_summary :=
        'The Adyen analysis population projection diverged.';
    end if;
  elsif not v_population_invalid
    and (v_run.analysis_total <> 0 or v_run.analysis_processed <> 0) then
    v_population_invalid := true;
    v_population_error_summary :=
      'The Adyen analysis population projection diverged.';
  end if;

  if not v_population_invalid then
    select population.source_id
    into v_invalid_source_id
    from public.financial_reconciliation_automatic_adyen_population population
    left join public.import_cgd_extrato_ordem bank
      on population.role = 'source'
     and bank.id = population.source_id
    left join public.import_fdm_accounts fdm
      on population.role = 'destination'
     and fdm.id = population.source_id
    where population.run_id = p_run_id
      and (
        (
          population.role = 'source'
          and (
            bank.id is null
            or bank.data is distinct from population.source_date
            or bank.montante is distinct from population.amount
            or bank.descritivo is distinct from population.description
            or population.account is distinct from ''
            or to_jsonb(bank) is distinct from population.row_snapshot
            or exists (
              select 1
              from public.financial_reconciliation_items locked
              where locked.source_type = 'import_cgd_extrato_ordem'
                and locked.source_id = population.source_id
            )
          )
        )
        or (
          population.role = 'destination'
          and (
            fdm.id is null
            or fdm.event_date is distinct from population.source_date
            or fdm.amount is distinct from population.amount
            or fdm.description is distinct from population.description
            or fdm.account is distinct from population.account
            or to_jsonb(fdm) is distinct from population.row_snapshot
            or exists (
              select 1
              from public.financial_reconciliation_items locked
              where locked.source_type = 'import_fdm_accounts'
                and locked.source_id = population.source_id
            )
          )
        )
      )
    order by population.month_ordinal, population.role,
             population.member_ordinal
    limit 1;
    if found then
      v_population_invalid := true;
      v_population_error_summary :=
        'The Adyen analysis population changed before completion.';
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

  for v_month in
    select population.month_ordinal, population.calendar_month
    from public.financial_reconciliation_automatic_adyen_population population
    where population.run_id = p_run_id
      and population.month_ordinal > v_run.analysis_processed
    group by population.month_ordinal, population.calendar_month
    order by population.month_ordinal
    limit 25
  loop
    v_page_count := v_page_count + 1;
    v_last_ordinal := v_month.month_ordinal;
    v_last_month := v_month.calendar_month;

    select array_agg(population.source_id order by population.member_ordinal),
           count(*)::integer, sum(population.amount)::numeric(14,2),
           (array_agg(population.source_id
              order by population.member_ordinal))[1],
           (array_agg(population.source_date
              order by population.member_ordinal))[1],
           (array_agg(population.row_snapshot
              order by population.member_ordinal))[1]
    into v_source_ids, v_source_count, v_source_total,
         v_base_id, v_base_date, v_base_snapshot
    from public.financial_reconciliation_automatic_adyen_population population
    where population.run_id = p_run_id
      and population.month_ordinal = v_month.month_ordinal
      and population.role = 'source';

    select array_agg(population.source_id order by population.member_ordinal),
           count(*)::integer, sum(population.amount)::numeric(14,2)
    into v_destination_ids, v_destination_count, v_destination_total
    from public.financial_reconciliation_automatic_adyen_population population
    where population.run_id = p_run_id
      and population.month_ordinal = v_month.month_ordinal
      and population.role = 'destination';

    v_last_cursor_id := coalesce(v_base_id, v_destination_ids[1]);

    if v_source_count = 0 or v_destination_count = 0 then
      continue;
    end if;

    v_difference := case v_operator
      when '-' then
        (v_source_total - v_destination_total)::numeric(14,2)
      when '+' then
        (v_source_total + v_destination_total)::numeric(14,2)
      else null
    end;
    if v_difference is null then
      raise exception 'Automatic Adyen monthly operator is invalid.';
    end if;

    v_status := case
      when abs(v_difference) <= v_allowed_difference then 'proposed'
      else 'ambiguous'
    end;
    v_reason := case
      when v_status = 'ambiguous' then 'monthly_difference_exceeded'
      else ''
    end;
    v_signature := public.financial_reconciliation_extension_sha256(
      jsonb_build_object(
        'ruleKey', 'cgd_bank_statement_fdm_adyen_monthly_payments',
        'ruleVersion', 1,
        'calendarMonth', v_month.calendar_month,
        'sourceIds', to_jsonb(v_source_ids),
        'destinationIds', to_jsonb(v_destination_ids),
        'sourceTotal', v_source_total,
        'destinationTotal', v_destination_total,
        'calculatedDifference', v_difference,
        'differenceAllowed', v_allowed_difference,
        'operator', v_operator
      )::text
    );
    v_summary_snapshot := jsonb_build_object(
      'ruleKey', 'cgd_bank_statement_fdm_adyen_monthly_payments',
      'ruleVersion', 1,
      'strategy', 'closed_calendar_month',
      'calendarMonth', v_month.calendar_month,
      'sourceDescriptionContains', 'Adyen',
      'destinationAccount', 'Adyen',
      'operator', v_operator,
      'differenceAllowed', v_allowed_difference,
      'maxDifferenceDays', 31,
      'sourceCount', v_source_count,
      'sourceTotal', v_source_total,
      'destinationCount', v_destination_count,
      'destinationTotal', v_destination_total,
      'calculatedDifference', v_difference,
      'technicalBaseSourceId', v_base_id,
      'technicalBaseSourceDate', v_base_date,
      'analysisTimestamp', now(),
      'signature', v_signature
    );

    v_proposal_id := null;
    insert into public.financial_reconciliation_automatic_proposals (
      run_id, rule_key, rule_version, base_source_type,
      base_source_id, base_source_date, base_snapshot,
      calculated_difference, allowed_difference, status, reason, signature,
      grouping_key, summary_snapshot
    ) values (
      p_run_id, 'cgd_bank_statement_fdm_adyen_monthly_payments', 1,
      'import_cgd_extrato_ordem', v_base_id, v_base_date, v_base_snapshot,
      v_difference, v_allowed_difference, v_status, v_reason, v_signature,
      to_char(v_month.calendar_month, 'YYYY-MM'), v_summary_snapshot
    )
    on conflict do nothing
    returning id into v_proposal_id;
    v_inserted := v_proposal_id is not null;

    if not v_inserted then
      select proposal.id into strict v_proposal_id
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = p_run_id
        and proposal.rule_key =
          'cgd_bank_statement_fdm_adyen_monthly_payments'
        and proposal.rule_version = 1
        and proposal.base_source_type = 'import_cgd_extrato_ordem'
        and proposal.base_source_id = v_base_id
        and proposal.signature = v_signature;
    else
      insert into public.financial_reconciliation_automatic_proposal_memberships (
        proposal_id, role, source_type, source_id, ordinal, source_date,
        amount, description, account, row_snapshot
      )
      select
        v_proposal_id, population.role, population.source_type,
        population.source_id, population.member_ordinal,
        population.source_date, population.amount, population.description,
        population.account, population.row_snapshot
      from public.financial_reconciliation_automatic_adyen_population population
      where population.run_id = p_run_id
        and population.month_ordinal = v_month.month_ordinal
        and population.role = 'source'
      order by population.member_ordinal;
      get diagnostics v_inserted_count = row_count;
      if v_inserted_count <> v_source_count then
        raise exception 'Automatic Adyen source membership insert was incomplete.';
      end if;

      insert into public.financial_reconciliation_automatic_proposal_memberships (
        proposal_id, role, source_type, source_id, ordinal, source_date,
        amount, description, account, row_snapshot
      )
      select
        v_proposal_id, population.role, population.source_type,
        population.source_id, population.member_ordinal,
        population.source_date, population.amount, population.description,
        population.account, population.row_snapshot
      from public.financial_reconciliation_automatic_adyen_population population
      where population.run_id = p_run_id
        and population.month_ordinal = v_month.month_ordinal
        and population.role = 'destination'
      order by population.member_ordinal;
      get diagnostics v_inserted_count = row_count;
      if v_inserted_count <> v_destination_count then
        raise exception
          'Automatic Adyen destination membership insert was incomplete.';
      end if;

      select array_agg(member.source_id order by member.ordinal),
             sum(member.amount)::numeric(14,2)
      into v_inserted_source_ids, v_inserted_source_total
      from public.financial_reconciliation_automatic_proposal_memberships member
      where member.proposal_id = v_proposal_id
        and member.role = 'source';
      select array_agg(member.source_id order by member.ordinal),
             sum(member.amount)::numeric(14,2)
      into v_inserted_destination_ids, v_inserted_destination_total
      from public.financial_reconciliation_automatic_proposal_memberships member
      where member.proposal_id = v_proposal_id
        and member.role = 'destination';

      if v_inserted_source_ids is distinct from v_source_ids
        or v_inserted_destination_ids is distinct from v_destination_ids
        or v_inserted_source_total is distinct from v_source_total
        or v_inserted_destination_total is distinct from v_destination_total
        or v_base_id is distinct from v_inserted_source_ids[1] then
        raise exception
          'Automatic Adyen proposal summary and memberships diverged.';
      end if;
    end if;
  end loop;

  v_processed_after := v_run.analysis_processed;
  if v_page_count > 0 then
    update public.financial_reconciliation_automatic_runs run
    set analysis_cursor_date = v_last_month,
        analysis_cursor_id = v_last_cursor_id,
        analysis_processed = v_last_ordinal,
        analysis_total = v_population_total,
        updated_at = now(),
        analysis_error_code = null,
        analysis_error_at = null
    where run.id = p_run_id
    returning run.analysis_processed into v_processed_after;
  end if;

  if v_processed_after is distinct from v_population_total then
    return public.financial_reconciliation_automatic_progress_or_run(p_run_id);
  end if;
  return public.financial_reconciliation_finalize_automatic_analysis(p_run_id);
end
$$;

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

create table if not exists public.financial_reconciliation_automatic_adyen_population (
  run_id uuid not null,
  calendar_month date not null,
  month_ordinal integer not null,
  role text not null,
  source_type text not null,
  source_id uuid not null,
  member_ordinal integer not null,
  source_date date not null,
  amount numeric(14,2) not null,
  description text not null,
  account text not null,
  row_snapshot jsonb not null,
  constraint fr_auto_adyen_population_run_fkey
    foreign key (run_id)
    references public.financial_reconciliation_automatic_runs(id)
    on delete cascade,
  constraint fr_auto_adyen_population_month_ordinal_check
    check (month_ordinal > 0),
  constraint fr_auto_adyen_population_member_ordinal_check
    check (member_ordinal > 0),
  constraint fr_auto_adyen_population_role_source_check
    check (
      (role = 'source' and source_type = 'import_cgd_extrato_ordem')
      or (role = 'destination' and source_type = 'import_fdm_accounts')
    ),
  constraint fr_auto_adyen_population_month_date_check
    check (
      calendar_month = date_trunc('month', calendar_month)::date
      and source_date >= calendar_month
      and source_date < (calendar_month + interval '1 month')::date
    ),
  constraint fr_auto_adyen_population_snapshot_check
    check (jsonb_typeof(row_snapshot) = 'object'),
  constraint fr_auto_adyen_population_pkey
    primary key (run_id, role, source_id),
  constraint fr_auto_adyen_population_ordinal_key
    unique (run_id, month_ordinal, role, member_ordinal)
);

do $migration$
declare
  v_column_count integer;
  v_constraint_count integer;
  v_pair record;
  v_actual_type "char";
  v_actual_definition text;
  v_expected_type "char";
  v_expected_definition text;
  v_foreign_key record;
begin
  select count(*) into v_column_count
  from information_schema.columns column_row
  where column_row.table_schema = 'public'
    and column_row.table_name =
      'financial_reconciliation_automatic_adyen_population';

  if v_column_count is distinct from 12
    or exists (
      select 1
      from (values
        (1, 'run_id', 'uuid', 'NO'),
        (2, 'calendar_month', 'date', 'NO'),
        (3, 'month_ordinal', 'integer', 'NO'),
        (4, 'role', 'text', 'NO'),
        (5, 'source_type', 'text', 'NO'),
        (6, 'source_id', 'uuid', 'NO'),
        (7, 'member_ordinal', 'integer', 'NO'),
        (8, 'source_date', 'date', 'NO'),
        (9, 'amount', 'numeric', 'NO'),
        (10, 'description', 'text', 'NO'),
        (11, 'account', 'text', 'NO'),
        (12, 'row_snapshot', 'jsonb', 'NO')
      ) expected(ordinal_position, column_name, data_type, is_nullable)
      left join information_schema.columns actual
        on actual.table_schema = 'public'
       and actual.table_name =
         'financial_reconciliation_automatic_adyen_population'
       and actual.ordinal_position = expected.ordinal_position
       and actual.column_name = expected.column_name
       and actual.data_type = expected.data_type
       and actual.is_nullable = expected.is_nullable
       and actual.column_default is null
      where actual.column_name is null
    )
    or not exists (
      select 1
      from information_schema.columns actual
      where actual.table_schema = 'public'
        and actual.table_name =
          'financial_reconciliation_automatic_adyen_population'
        and actual.column_name = 'amount'
        and actual.numeric_precision = 14
        and actual.numeric_scale = 2
    ) then
    raise exception 'Installed Adyen population columns differ from the required contract.';
  end if;

  select count(*) into v_constraint_count
  from pg_constraint constraint_row
  where constraint_row.conrelid =
    'public.financial_reconciliation_automatic_adyen_population'::regclass;
  if v_constraint_count is distinct from 8 then
    raise exception 'Installed Adyen population constraint count differs from the required contract.';
  end if;

  select constraint_row.contype, constraint_row.confrelid,
         constraint_row.confdeltype, constraint_row.confupdtype,
         constraint_row.confmatchtype, constraint_row.condeferrable,
         constraint_row.condeferred, constraint_row.convalidated,
         constraint_row.conkey, constraint_row.confkey
  into v_foreign_key
  from pg_constraint constraint_row
  where constraint_row.conrelid =
      'public.financial_reconciliation_automatic_adyen_population'::regclass
    and constraint_row.conname = 'fr_auto_adyen_population_run_fkey';
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
    ]::smallint[] then
    raise exception 'Installed Adyen population constraint % differs from the required definition.',
      'fr_auto_adyen_population_run_fkey';
  end if;

  create temporary table task4_adyen_population_expected (
    run_id uuid not null,
    calendar_month date not null,
    month_ordinal integer not null,
    role text not null,
    source_type text not null,
    source_id uuid not null,
    member_ordinal integer not null,
    source_date date not null,
    amount numeric(14,2) not null,
    description text not null,
    account text not null,
    row_snapshot jsonb not null,
    constraint task4_adyen_population_month_ordinal_check
      check (month_ordinal > 0),
    constraint task4_adyen_population_member_ordinal_check
      check (member_ordinal > 0),
    constraint task4_adyen_population_role_source_check
      check (
        (role = 'source' and source_type = 'import_cgd_extrato_ordem')
        or (role = 'destination' and source_type = 'import_fdm_accounts')
      ),
    constraint task4_adyen_population_month_date_check
      check (
        calendar_month = date_trunc('month', calendar_month)::date
        and source_date >= calendar_month
        and source_date < (calendar_month + interval '1 month')::date
      ),
    constraint task4_adyen_population_snapshot_check
      check (jsonb_typeof(row_snapshot) = 'object'),
    constraint task4_adyen_population_pkey
      primary key (run_id, role, source_id),
    constraint task4_adyen_population_ordinal_key
      unique (run_id, month_ordinal, role, member_ordinal)
  ) on commit drop;

  for v_pair in
    select * from (values
      ('fr_auto_adyen_population_month_ordinal_check',
       'task4_adyen_population_month_ordinal_check'),
      ('fr_auto_adyen_population_member_ordinal_check',
       'task4_adyen_population_member_ordinal_check'),
      ('fr_auto_adyen_population_role_source_check',
       'task4_adyen_population_role_source_check'),
      ('fr_auto_adyen_population_month_date_check',
       'task4_adyen_population_month_date_check'),
      ('fr_auto_adyen_population_snapshot_check',
       'task4_adyen_population_snapshot_check'),
      ('fr_auto_adyen_population_pkey',
       'task4_adyen_population_pkey'),
      ('fr_auto_adyen_population_ordinal_key',
       'task4_adyen_population_ordinal_key')
    ) expected(actual_name, expected_name)
  loop
    select constraint_row.contype,
           pg_get_constraintdef(constraint_row.oid, true)
    into v_actual_type, v_actual_definition
    from pg_constraint constraint_row
    where constraint_row.conrelid =
        'public.financial_reconciliation_automatic_adyen_population'::regclass
      and constraint_row.conname = v_pair.actual_name;
    select constraint_row.contype,
           pg_get_constraintdef(constraint_row.oid, true)
    into strict v_expected_type, v_expected_definition
    from pg_constraint constraint_row
    where constraint_row.conrelid =
        'task4_adyen_population_expected'::regclass
      and constraint_row.conname = v_pair.expected_name;
    if v_actual_type is distinct from v_expected_type
      or v_actual_definition is distinct from v_expected_definition then
      raise exception 'Installed Adyen population constraint % differs from the required definition.',
        v_pair.actual_name;
    end if;
  end loop;

  if (select count(*) from pg_index index_row
      where index_row.indrelid =
        'public.financial_reconciliation_automatic_adyen_population'::regclass)
      is distinct from 2 then
    raise exception 'Installed Adyen population index count differs from the required contract.';
  end if;
  drop table task4_adyen_population_expected;
end
$migration$;

alter table public.financial_reconciliation_automatic_adyen_population
  enable row level security;

revoke all on table public.financial_reconciliation_automatic_adyen_population
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
        'public.financial_reconciliation_automatic_adyen_population'::regclass)
    or exists (
      select 1 from pg_policy policy_row
      where policy_row.polrelid =
        'public.financial_reconciliation_automatic_adyen_population'::regclass
    )
    or exists (
      select 1 from information_schema.table_privileges grant_row
      where grant_row.table_schema = 'public'
        and grant_row.table_name =
          'financial_reconciliation_automatic_adyen_population'
        and grant_row.grantee = 'PUBLIC'
    )
    or exists (
      select 1 from information_schema.column_privileges grant_row
      where grant_row.table_schema = 'public'
        and grant_row.table_name =
          'financial_reconciliation_automatic_adyen_population'
        and grant_row.grantee in (
          'PUBLIC','anon','authenticated','service_role'
        )
    ) then
    raise exception 'Installed Adyen population security differs from the required contract.';
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
        raise exception 'Installed Adyen population security grants % to %.',
          v_privilege, v_role;
      end if;
    end loop;
  end loop;
end
$migration$;

create or replace function public.financial_reconciliation_automatic_adyen_month_count()
returns bigint
language sql
stable
security definer set search_path = public, pg_temp
as $$
  with eligible_months as (
    select date_trunc('month', bank.data)::date as calendar_month
    from public.import_cgd_extrato_ordem bank
    where bank.data >= date '2026-01-01'
      and bank.data < date_trunc('month', current_date)::date
      and bank.montante is not null
      and bank.descritivo ilike '%Adyen%'
      and not exists (
        select 1
        from public.financial_reconciliation_items locked
        where locked.source_type = 'import_cgd_extrato_ordem'
          and locked.source_id = bank.id
      )
    union
    select date_trunc('month', fdm.event_date)::date as calendar_month
    from public.import_fdm_accounts fdm
    where fdm.event_date >= date '2026-01-01'
      and fdm.event_date < date_trunc('month', current_date)::date
      and fdm.amount is not null
      and fdm.account = 'Adyen'
      and not exists (
        select 1
        from public.financial_reconciliation_items locked
        where locked.source_type = 'import_fdm_accounts'
          and locked.source_id = fdm.id
      )
  )
  select count(*) from eligible_months
$$;

create or replace function public.financial_reconciliation_automatic_adyen_month_page(
  p_after_month date,
  p_limit integer
)
returns table (
  calendar_month date
)
language plpgsql
stable
security definer set search_path = public, pg_temp
as $$
begin
  if p_limit is null or p_limit not between 1 and 25 then
    raise exception 'Automatic Adyen month page size must be between 1 and 25.';
  end if;

  return query
  with eligible_months as (
    select date_trunc('month', bank.data)::date as calendar_month
    from public.import_cgd_extrato_ordem bank
    where bank.data >= date '2026-01-01'
      and bank.data < date_trunc('month', current_date)::date
      and bank.montante is not null
      and bank.descritivo ilike '%Adyen%'
      and not exists (
        select 1
        from public.financial_reconciliation_items locked
        where locked.source_type = 'import_cgd_extrato_ordem'
          and locked.source_id = bank.id
      )
    union
    select date_trunc('month', fdm.event_date)::date as calendar_month
    from public.import_fdm_accounts fdm
    where fdm.event_date >= date '2026-01-01'
      and fdm.event_date < date_trunc('month', current_date)::date
      and fdm.amount is not null
      and fdm.account = 'Adyen'
      and not exists (
        select 1
        from public.financial_reconciliation_items locked
        where locked.source_type = 'import_fdm_accounts'
          and locked.source_id = fdm.id
      )
  )
  select eligible.calendar_month
  from eligible_months eligible
  where eligible.calendar_month >
    coalesce(p_after_month, date '0001-01-01')
  order by eligible.calendar_month
  limit p_limit;
end
$$;

do $migration$
begin
  if to_regprocedure(
      'public.financial_reconciliation_finalize_automatic_prior_analysis(uuid)'
    ) is null then
    alter function public.financial_reconciliation_finalize_automatic_analysis(uuid)
      rename to financial_reconciliation_finalize_automatic_prior_analysis;
  end if;
end
$migration$;

create or replace function public.financial_reconciliation_finalize_automatic_analysis(
  p_run_id uuid
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
  select * into v_run
  from public.financial_reconciliation_automatic_runs run
  where run.id = p_run_id
  for update;
  if not found then
    raise exception 'Automatic analysis run was not found.';
  end if;

  if jsonb_array_length(v_run.definition_config_snapshot) = 1 then
    select snapshot.value into strict v_rule
    from jsonb_array_elements(v_run.definition_config_snapshot) snapshot(value);
    if coalesce(v_rule->>'ruleVersion', '') ~ '^[0-9]+$'
      and length(v_rule->>'ruleVersion') <= 10 then
      v_rule_key := v_rule->>'ruleKey';
      v_rule_version := (v_rule->>'ruleVersion')::integer;
    end if;
  end if;

  if (v_rule_key, v_rule_version) in (
      ('financial_documents_cgd_bank_statement', 2),
      ('financial_documents_cgd_credit_card', 1),
      ('financial_documents_cgd_bank_statement_amount_only', 1),
      ('financial_documents_cgd_credit_card_amount_only', 1),
      ('cgd_bank_statement_fdm_credit_card_monthly_income', 2)
    ) then
    return public.financial_reconciliation_finalize_automatic_prior_analysis(
      p_run_id
    );
  elsif v_rule_key = 'fdm_bank_transfer_cgd_bank_statement_combination'
    and v_rule_version = 1 then
    if v_run.analysis_completed_at is not null then
      return public.get_financial_reconciliation_automatic_run(p_run_id);
    end if;
    raise exception
      'Automatic Bank Reservation analysis uses strategy-specific finalization.';
  elsif v_rule_key is distinct from
      'cgd_bank_statement_fdm_adyen_monthly_payments'
    or v_rule_version is distinct from 1 then
    raise exception 'Automatic reconciliation rule is unsupported.';
  end if;

  if v_run.analysis_completed_at is not null then
    return public.get_financial_reconciliation_automatic_run(p_run_id);
  end if;
  if v_run.status <> 'analyzing'
    or v_run.analysis_processed is distinct from v_run.analysis_total then
    raise exception 'Automatic Adyen analysis is not ready to finalize.';
  end if;

  update public.financial_reconciliation_automatic_runs run
  set status = case when exists (
        select 1
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = p_run_id
          and proposal.rule_key =
            'cgd_bank_statement_fdm_adyen_monthly_payments'
          and proposal.rule_version = 1
          and proposal.status in ('proposed','ambiguous','stale','failed')
      ) then 'ready' else 'completed' end,
      finished_at = case when exists (
        select 1
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = p_run_id
          and proposal.rule_key =
            'cgd_bank_statement_fdm_adyen_monthly_payments'
          and proposal.rule_version = 1
          and proposal.status in ('proposed','ambiguous','stale','failed')
      ) then null else now() end,
      analysis_completed_at = now(),
      updated_at = now(),
      analysis_error_code = null,
      analysis_error_at = null,
      counts = (
        select jsonb_build_object(
          'bases', v_run.analysis_total,
          'proposed', count(*) filter (
            where proposal.status = 'proposed'
          ),
          'ambiguous', count(*) filter (
            where proposal.status = 'ambiguous'
          ),
          'skipped', greatest(
            v_run.analysis_processed - count(*) filter (
              where proposal.status in (
                'proposed','ambiguous','stale','failed'
              )
            ),
            0
          )
        )
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = p_run_id
          and proposal.rule_key =
            'cgd_bank_statement_fdm_adyen_monthly_payments'
          and proposal.rule_version = 1
      )
  where run.id = p_run_id and run.analysis_completed_at is null;

  return public.get_financial_reconciliation_automatic_run(p_run_id);
end
$$;

create or replace function public.financial_reconciliation_continue_automatic_adyen_mutable_prior(
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
  v_month record;
  v_page_count integer := 0;
  v_total bigint;
  v_last_month date;
  v_last_cursor_id uuid;
  v_allowed_difference numeric(14,2);
  v_operator text;
  v_source_ids uuid[];
  v_destination_ids uuid[];
  v_source_count integer;
  v_destination_count integer;
  v_source_total numeric(14,2);
  v_destination_total numeric(14,2);
  v_difference numeric(14,2);
  v_base_id uuid;
  v_base_date date;
  v_base_snapshot jsonb;
  v_summary_snapshot jsonb;
  v_signature text;
  v_status text;
  v_reason text;
  v_proposal_id uuid;
  v_inserted boolean;
  v_inserted_count integer;
  v_inserted_source_ids uuid[];
  v_inserted_destination_ids uuid[];
  v_inserted_source_total numeric(14,2);
  v_inserted_destination_total numeric(14,2);
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
      'cgd_bank_statement_fdm_adyen_monthly_payments'
    or v_rule->>'ruleVersion' is distinct from '1'
    or v_rule->>'displayName' is distinct from
      'FDM Accounts – Adyen Reservation Payments'
    or coalesce(v_rule->>'priority', '') !~ '^[1-9][0-9]*$'
    or length(v_rule->>'priority') > 10
    or v_rule->>'destinationSourceType' is distinct from
      'import_fdm_accounts'
    or v_rule->>'operator' is distinct from '-'
    or coalesce(v_rule->>'differenceAllowed', '') !~
      '^[0-9]+(\.[0-9]+)?$'
    or length(v_rule->>'differenceAllowed') > 20
    or (v_rule->>'differenceAllowed')::numeric < 0
    or v_rule->>'maxDifferenceDays' is distinct from '31'
    or v_rule->'definition' is distinct from jsonb_build_object(
      'strategy', 'closed_calendar_month',
      'bankDescriptionContains', 'Adyen',
      'fdmAccount', 'Adyen',
      'requiresBothSides', true,
      'monthMarkerDays', 31
    ) then
    raise exception 'Automatic Adyen monthly rule snapshot contract is invalid.';
  end if;

  v_allowed_difference :=
    (v_rule->>'differenceAllowed')::numeric(14,2);
  v_operator := v_rule->>'operator';

  lock table
    public.import_cgd_extrato_ordem,
    public.import_fdm_accounts,
    public.financial_reconciliation_items
  in share row exclusive mode;

  select public.financial_reconciliation_automatic_adyen_month_count()
  into v_total;
  update public.financial_reconciliation_automatic_runs run
  set analysis_total = greatest(
        run.analysis_total, run.analysis_processed, v_total
      ),
      updated_at = now()
  where run.id = p_run_id;

  for v_month in
    select page.calendar_month
    from public.financial_reconciliation_automatic_adyen_month_page(
      v_run.analysis_cursor_date,
      25
    ) page
  loop
    v_page_count := v_page_count + 1;
    v_last_month := v_month.calendar_month;

    select
      array_agg(bank.id order by bank.data, bank.id),
      count(*)::integer,
      sum(bank.montante)::numeric(14,2),
      (array_agg(bank.id order by bank.data, bank.id))[1],
      (array_agg(bank.data order by bank.data, bank.id))[1]
    into
      v_source_ids, v_source_count, v_source_total, v_base_id, v_base_date
    from public.import_cgd_extrato_ordem bank
    where bank.data >= v_month.calendar_month
      and bank.data < v_month.calendar_month + interval '1 month'
      and bank.data >= date '2026-01-01'
      and bank.data < date_trunc('month', current_date)::date
      and bank.montante is not null
      and bank.descritivo ilike '%Adyen%'
      and not exists (
        select 1
        from public.financial_reconciliation_items locked
        where locked.source_type = 'import_cgd_extrato_ordem'
          and locked.source_id = bank.id
      );

    select
      array_agg(fdm.id order by fdm.event_date, fdm.id),
      count(*)::integer,
      sum(fdm.amount)::numeric(14,2)
    into v_destination_ids, v_destination_count, v_destination_total
    from public.import_fdm_accounts fdm
    where fdm.event_date >= v_month.calendar_month
      and fdm.event_date < v_month.calendar_month + interval '1 month'
      and fdm.event_date >= date '2026-01-01'
      and fdm.event_date < date_trunc('month', current_date)::date
      and fdm.amount is not null
      and fdm.account = 'Adyen'
      and not exists (
        select 1
        from public.financial_reconciliation_items locked
        where locked.source_type = 'import_fdm_accounts'
          and locked.source_id = fdm.id
      );

    v_last_cursor_id := coalesce(v_base_id, v_destination_ids[1]);

    if cardinality(v_source_ids) is null
      or cardinality(v_destination_ids) is null then
      continue;
    end if;

    v_difference := case v_operator
      when '-' then
        (v_source_total - v_destination_total)::numeric(14,2)
      when '+' then
        (v_source_total + v_destination_total)::numeric(14,2)
      else null
    end;
    if v_difference is null then
      raise exception 'Automatic Adyen monthly operator is invalid.';
    end if;

    select to_jsonb(bank)
    into strict v_base_snapshot
    from public.import_cgd_extrato_ordem bank
    where bank.id = v_base_id
      and bank.id = any(v_source_ids);

    v_status := case
      when abs(v_difference) <= v_allowed_difference then 'proposed'
      else 'ambiguous'
    end;
    v_reason := case
      when v_status = 'ambiguous' then 'monthly_difference_exceeded'
      else ''
    end;
    v_signature := public.financial_reconciliation_extension_sha256(
      jsonb_build_object(
        'ruleKey', 'cgd_bank_statement_fdm_adyen_monthly_payments',
        'ruleVersion', 1,
        'calendarMonth', v_month.calendar_month,
        'sourceIds', to_jsonb(v_source_ids),
        'destinationIds', to_jsonb(v_destination_ids),
        'sourceTotal', v_source_total,
        'destinationTotal', v_destination_total,
        'calculatedDifference', v_difference,
        'differenceAllowed', v_allowed_difference,
        'operator', v_operator
      )::text
    );
    v_summary_snapshot := jsonb_build_object(
      'ruleKey', 'cgd_bank_statement_fdm_adyen_monthly_payments',
      'ruleVersion', 1,
      'strategy', 'closed_calendar_month',
      'calendarMonth', v_month.calendar_month,
      'sourceDescriptionContains', 'Adyen',
      'destinationAccount', 'Adyen',
      'operator', v_operator,
      'differenceAllowed', v_allowed_difference,
      'maxDifferenceDays', 31,
      'sourceCount', v_source_count,
      'sourceTotal', v_source_total,
      'destinationCount', v_destination_count,
      'destinationTotal', v_destination_total,
      'calculatedDifference', v_difference,
      'technicalBaseSourceId', v_base_id,
      'technicalBaseSourceDate', v_base_date,
      'analysisTimestamp', now(),
      'signature', v_signature
    );

    v_proposal_id := null;
    insert into public.financial_reconciliation_automatic_proposals (
      run_id, rule_key, rule_version, base_source_type,
      base_source_id, base_source_date, base_snapshot,
      calculated_difference, allowed_difference, status, reason, signature,
      grouping_key, summary_snapshot
    ) values (
      p_run_id, 'cgd_bank_statement_fdm_adyen_monthly_payments', 1,
      'import_cgd_extrato_ordem', v_base_id, v_base_date, v_base_snapshot,
      v_difference, v_allowed_difference, v_status, v_reason, v_signature,
      to_char(v_month.calendar_month, 'YYYY-MM'), v_summary_snapshot
    )
    on conflict do nothing
    returning id into v_proposal_id;
    v_inserted := v_proposal_id is not null;

    if not v_inserted then
      select proposal.id into strict v_proposal_id
      from public.financial_reconciliation_automatic_proposals proposal
      where proposal.run_id = p_run_id
        and proposal.rule_key =
          'cgd_bank_statement_fdm_adyen_monthly_payments'
        and proposal.rule_version = 1
        and proposal.base_source_type = 'import_cgd_extrato_ordem'
        and proposal.base_source_id = v_base_id
        and proposal.signature = v_signature;
    else
      insert into public.financial_reconciliation_automatic_proposal_memberships (
        proposal_id, role, source_type, source_id, ordinal, source_date,
        amount, description, account, row_snapshot
      )
      select
        v_proposal_id, 'source', 'import_cgd_extrato_ordem',
        source_member.id, source_member.ordinal, source_member.data,
        source_member.montante, source_member.descritivo, '',
        source_member.row_snapshot
      from (
        select bank.id, bank.data, bank.montante, bank.descritivo,
               to_jsonb(bank) as row_snapshot,
               row_number() over (order by bank.data, bank.id)::integer
                 as ordinal
        from public.import_cgd_extrato_ordem bank
        where bank.id = any(v_source_ids)
      ) source_member
      order by source_member.ordinal;
      get diagnostics v_inserted_count = row_count;
      if v_inserted_count <> v_source_count then
        raise exception 'Automatic Adyen source membership insert was incomplete.';
      end if;

      insert into public.financial_reconciliation_automatic_proposal_memberships (
        proposal_id, role, source_type, source_id, ordinal, source_date,
        amount, description, account, row_snapshot
      )
      select
        v_proposal_id, 'destination', 'import_fdm_accounts',
        destination_member.id, destination_member.ordinal,
        destination_member.event_date, destination_member.amount,
        destination_member.description, destination_member.account,
        destination_member.row_snapshot
      from (
        select fdm.id, fdm.event_date, fdm.amount, fdm.description,
               fdm.account, to_jsonb(fdm) as row_snapshot,
               row_number() over (order by fdm.event_date, fdm.id)::integer
                 as ordinal
        from public.import_fdm_accounts fdm
        where fdm.id = any(v_destination_ids)
      ) destination_member
      order by destination_member.ordinal;
      get diagnostics v_inserted_count = row_count;
      if v_inserted_count <> v_destination_count then
        raise exception
          'Automatic Adyen destination membership insert was incomplete.';
      end if;

      select
        array_agg(member.source_id order by member.ordinal),
        sum(member.amount)::numeric(14,2)
      into v_inserted_source_ids, v_inserted_source_total
      from public.financial_reconciliation_automatic_proposal_memberships member
      where member.proposal_id = v_proposal_id
        and member.role = 'source';

      select
        array_agg(member.source_id order by member.ordinal),
        sum(member.amount)::numeric(14,2)
      into v_inserted_destination_ids, v_inserted_destination_total
      from public.financial_reconciliation_automatic_proposal_memberships member
      where member.proposal_id = v_proposal_id
        and member.role = 'destination';

      if v_inserted_source_ids is distinct from v_source_ids
        or v_inserted_destination_ids is distinct from v_destination_ids
        or v_inserted_source_total is distinct from v_source_total
        or v_inserted_destination_total is distinct from v_destination_total
        or v_base_id is distinct from v_inserted_source_ids[1] then
        raise exception
          'Automatic Adyen proposal summary and memberships diverged.';
      end if;
    end if;
  end loop;

  if v_page_count > 0 then
    update public.financial_reconciliation_automatic_runs run
    set analysis_cursor_date = v_last_month,
        analysis_cursor_id = v_last_cursor_id,
        analysis_processed = greatest(
          run.analysis_processed,
          run.analysis_processed + v_page_count
        ),
        analysis_total = greatest(
          run.analysis_total,
          run.analysis_processed + v_page_count
        ),
        updated_at = now(),
        analysis_error_code = null,
        analysis_error_at = null
    where run.id = p_run_id
      and (
        run.analysis_cursor_date is null
        or v_last_month > run.analysis_cursor_date
      );
  end if;

  if v_page_count < 25 then
    return public.financial_reconciliation_finalize_automatic_analysis(
      p_run_id
    );
  end if;
  return public.financial_reconciliation_automatic_progress_or_run(p_run_id);
end
$$;

drop function public.financial_reconciliation_continue_automatic_adyen_mutable_prior(uuid,text);

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
    elsif v_rule_key = 'cgd_bank_statement_fdm_adyen_monthly_payments'
      and v_rule_version = 1 then
      return public.financial_reconciliation_continue_automatic_adyen_monthly(
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

create or replace function public.financial_reconciliation_lock_fdm_bank_automatic_members(
  p_proposal_id uuid
)
returns table (
  role text,
  source_type text,
  source_id uuid,
  ordinal integer,
  source_date date,
  amount numeric(14,2),
  description text,
  account text,
  row_snapshot jsonb,
  live_exists boolean
)
language plpgsql
security definer set search_path = public, pg_temp
as $$
declare
  v_member record;
begin
  if p_proposal_id is null then
    raise exception 'Automation proposal ID is required.';
  end if;

  for v_member in
    select membership.*
    from public.financial_reconciliation_automatic_proposal_memberships membership
    where membership.proposal_id = p_proposal_id
    order by membership.source_type, membership.source_id
    for share of membership
  loop
    role := v_member.role;
    source_type := v_member.source_type;
    source_id := v_member.source_id;
    ordinal := v_member.ordinal;
    source_date := null;
    amount := null;
    description := null;
    account := null;
    row_snapshot := null;
    live_exists := false;

    if v_member.source_type = 'import_cgd_extrato_ordem' then
      select
        bank.data,
        bank.montante,
        bank.descritivo,
        ''::text,
        to_jsonb(bank)
      into source_date, amount, description, account, row_snapshot
      from public.import_cgd_extrato_ordem bank
      where bank.id = v_member.source_id
      for update of bank;
      live_exists := found;
    elsif v_member.source_type = 'import_fdm_accounts' then
      select
        fdm.event_date,
        fdm.amount,
        fdm.description,
        fdm.account,
        to_jsonb(fdm)
      into source_date, amount, description, account, row_snapshot
      from public.import_fdm_accounts fdm
      where fdm.id = v_member.source_id
      for update of fdm;
      live_exists := found;
    else
      raise exception 'Automatic proposal membership source type is unsupported.';
    end if;

    return next;
  end loop;
end
$$;

create or replace function public.financial_reconciliation_commit_fdm_bank_automatic_proposal(
  p_proposal_id uuid,
  p_actor text,
  p_base_source_type text,
  p_matching_source_type text,
  p_matching_operator text,
  p_expected_difference numeric,
  p_forced_comment text,
  p_rule_snapshot jsonb,
  p_config_snapshot jsonb,
  p_operator_snapshot jsonb,
  p_membership_snapshots jsonb
)
returns uuid
language plpgsql
security definer set search_path = public, pg_temp
as $$
declare
  v_run public.financial_reconciliation_automatic_runs%rowtype;
  v_proposal public.financial_reconciliation_automatic_proposals%rowtype;
  v_matching_source_types jsonb;
  v_matching_source_rules jsonb;
  v_technical_base_id uuid;
  v_technical_base_amount numeric(14,2);
  v_reconciliation_id uuid;
  v_actual_item_count bigint;
  v_expected_item_count bigint;
  v_actual_difference numeric(14,2);
  v_completion_type text;
begin
  if (p_base_source_type, p_matching_source_type, p_matching_operator)
      not in (
        ('import_fdm_accounts', 'import_cgd_extrato_ordem', '+'),
        ('import_cgd_extrato_ordem', 'import_fdm_accounts', '-')
      ) then
    raise exception 'Automatic grouped reconciliation lifecycle is unsupported.';
  end if;
  if p_expected_difference is null
    or jsonb_typeof(p_rule_snapshot) is distinct from 'object'
    or jsonb_typeof(p_config_snapshot) is distinct from 'object'
    or jsonb_typeof(p_operator_snapshot) is distinct from 'object'
    or jsonb_typeof(p_membership_snapshots) is distinct from 'array' then
    raise exception
      'Automatic grouped reconciliation lifecycle snapshots changed after revalidation.';
  end if;

  select * into strict v_proposal
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.id = p_proposal_id
  for update;
  select * into strict v_run
  from public.financial_reconciliation_automatic_runs run
  where run.id = v_proposal.run_id
  for update;
  if v_proposal.status <> 'proposed'
    or v_run.actor is distinct from p_actor then
    raise exception
      'Automatic grouped reconciliation lifecycle snapshots changed after revalidation.';
  end if;

  select membership.source_id, membership.amount
  into strict v_technical_base_id, v_technical_base_amount
  from public.financial_reconciliation_automatic_proposal_memberships membership
  where membership.proposal_id = p_proposal_id
    and membership.role = 'source'
    and membership.source_type = p_base_source_type
  order by membership.ordinal, membership.source_id
  limit 1;

  select count(*) into v_expected_item_count
  from public.financial_reconciliation_automatic_proposal_memberships membership
  where membership.proposal_id = p_proposal_id;
  if v_expected_item_count < 2
    or jsonb_array_length(p_membership_snapshots) <> v_expected_item_count then
    raise exception
      'Automatic grouped reconciliation lifecycle snapshots changed after revalidation.';
  end if;

  v_matching_source_types := jsonb_build_array(p_matching_source_type);
  v_matching_source_rules := jsonb_build_array(jsonb_build_object(
    'sourceType', p_matching_source_type,
    'operator', p_matching_operator
  ));

  update public.financial_reconciliation_automatic_proposals
  set status = 'executing', reason = '', error = '', error_detail = '',
      updated_at = now()
  where id = p_proposal_id;

  insert into public.financial_reconciliations (
    status, base_source_type, matching_source_types,
    matching_source_rules, created_by
  ) values (
    'started', p_base_source_type, v_matching_source_types,
    v_matching_source_rules, p_actor
  ) returning id into v_reconciliation_id;

  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  ) values (
    v_reconciliation_id, p_base_source_type, v_technical_base_id,
    v_technical_base_amount, p_actor
  );

  update public.financial_reconciliations
  set difference_amount = v_technical_base_amount,
      updated_at = timezone('utc', now())
  where id = v_reconciliation_id;

  insert into public.financial_reconciliation_audit (
    reconciliation_id, action, actor, difference_amount, metadata
  ) values (
    v_reconciliation_id, 'start', p_actor, v_technical_base_amount,
    jsonb_build_object(
      'sourceType', p_base_source_type,
      'sourceId', v_technical_base_id,
      'differenceAmount', v_technical_base_amount,
      'matchingSourceRules', v_matching_source_rules
    )
  );

  insert into public.financial_reconciliation_items (
    reconciliation_id, source_type, source_id, amount_snapshot, created_by
  )
  select v_reconciliation_id, membership.source_type,
         membership.source_id, membership.amount, p_actor
  from public.financial_reconciliation_automatic_proposal_memberships membership
  where membership.proposal_id = p_proposal_id
    and not (
      membership.source_type = p_base_source_type
      and membership.source_id = v_technical_base_id
    )
  order by membership.source_type, membership.source_id;

  with ordered as (
    select
      membership.*,
      row_number() over (
        order by
          case
            when membership.source_type = p_base_source_type
              and membership.source_id = v_technical_base_id then 0
            else 1
          end,
          membership.source_type,
          membership.source_id
      ) as sequence,
      sum(case
        when membership.source_type = p_base_source_type
          then membership.amount
        when p_matching_operator = '+' then membership.amount
        else -membership.amount
      end) over (
        order by
          case
            when membership.source_type = p_base_source_type
              and membership.source_id = v_technical_base_id then 0
            else 1
          end,
          membership.source_type,
          membership.source_id
        rows between unbounded preceding and current row
      )::numeric(14,2) as running_difference
    from public.financial_reconciliation_automatic_proposal_memberships membership
    where membership.proposal_id = p_proposal_id
  )
  insert into public.financial_reconciliation_audit (
    reconciliation_id, action, actor, difference_amount, metadata
  )
  select
    v_reconciliation_id, 'add_item', p_actor,
    ordered.running_difference,
    jsonb_build_object(
      'sourceType', ordered.source_type,
      'sourceId', ordered.source_id,
      'differenceAmount', ordered.running_difference
    )
  from ordered
  where ordered.sequence > 1
  order by ordered.sequence;

  select count(*) into v_actual_item_count
  from public.financial_reconciliation_items item
  where item.reconciliation_id = v_reconciliation_id;
  select public.financial_reconciliation_difference(
    p_base_source_type, v_matching_source_rules, v_reconciliation_id
  ) into v_actual_difference;

  update public.financial_reconciliations
  set difference_amount = v_actual_difference,
      updated_at = timezone('utc', now())
  where id = v_reconciliation_id;

  if v_actual_item_count <> v_expected_item_count
    or v_actual_difference is distinct from p_expected_difference
    or exists (
      select membership.source_type, membership.source_id, membership.amount
      from public.financial_reconciliation_automatic_proposal_memberships membership
      where membership.proposal_id = p_proposal_id
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
      where membership.proposal_id = p_proposal_id
    ) then
    raise exception
      'Automatic grouped reconciliation lifecycle snapshots changed after revalidation.';
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

  if v_actual_difference = 0 and p_forced_comment is null then
    v_completion_type := 'normal';
    perform public.financial_reconciliation_action(
      'complete', p_actor, v_reconciliation_id, null, null, null
    );
  elsif v_actual_difference <> 0 and p_forced_comment is not null then
    v_completion_type := 'forced';
    perform public.financial_reconciliation_action(
      'force_complete', p_actor, v_reconciliation_id,
      null, null, p_forced_comment
    );
  else
    raise exception
      'Automatic grouped reconciliation lifecycle snapshots changed after revalidation.';
  end if;

  insert into public.financial_reconciliation_audit (
    reconciliation_id, action, actor, comment, difference_amount, metadata
  ) values (
    v_reconciliation_id, 'automatic_complete', p_actor,
    p_forced_comment, v_actual_difference,
    jsonb_build_object(
      'ruleSnapshot', p_rule_snapshot,
      'configSnapshot', p_config_snapshot,
      'operatorSnapshot', p_operator_snapshot,
      'summarySnapshot', v_proposal.summary_snapshot,
      'membershipSnapshots', p_membership_snapshots,
      'proposalSignature', v_proposal.signature,
      'trigger', v_run.trigger,
      'runId', v_run.id,
      'proposalId', v_proposal.id,
      'tolerance', v_proposal.allowed_difference,
      'calculatedDifference', v_actual_difference
    )
  );

  if not exists (
    select 1
    from public.financial_reconciliations reconciliation
    where reconciliation.id = v_reconciliation_id
      and reconciliation.status = 'complete'
      and reconciliation.completion_type = v_completion_type
      and reconciliation.difference_amount = v_actual_difference
      and reconciliation.forced_completion_comment is not distinct from
        p_forced_comment
      and reconciliation.origin = 'automatic'
      and reconciliation.automatic_trigger = v_run.trigger
      and reconciliation.automatic_rule_key = v_proposal.rule_key
      and reconciliation.automatic_rule_version = v_proposal.rule_version
      and reconciliation.automatic_run_id = v_run.id
      and reconciliation.automatic_proposal_id = v_proposal.id
      and reconciliation.matching_source_rules = v_matching_source_rules
  ) or (select count(*)
        from public.financial_reconciliation_audit audit
        where audit.reconciliation_id = v_reconciliation_id
          and audit.action = 'automatic_complete') <> 1 then
    raise exception
      'Automatic grouped reconciliation lifecycle snapshots changed after revalidation.';
  end if;

  update public.financial_reconciliation_automatic_proposals
  set status = 'completed', reconciliation_id = v_reconciliation_id,
      completed_at = now(), reason = '', error = '', error_detail = '',
      updated_at = now()
  where id = p_proposal_id;

  return v_reconciliation_id;
end
$$;

create or replace function public.financial_reconciliation_execute_bank_reservation_proposal(
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
  v_expected_definition jsonb := jsonb_build_object(
    'strategy', 'bounded_exact_combination',
    'sourceAccount', 'Bank Transfer',
    'maxSourceRecords', 10,
    'candidatePoolLimit', 60,
    'stateLimit', 250000,
    'evidenceGroupLimit', 12,
    'amountMode', 'signed_integer_cents',
    'dateMode', 'inclusive_days'
  );
  v_current_definition jsonb;
  v_current_display_name text;
  v_current_base_source_type text;
  v_current_destination_source_types jsonb;
  v_current_rule_version integer;
  v_current_enabled boolean;
  v_current_allow_manual boolean;
  v_current_include_scheduled boolean;
  v_current_difference_allowed numeric(14,2);
  v_current_max_difference_days integer;
  v_current_priority integer;
  v_current_operator text;
  v_snapshot_rule_version integer;
  v_snapshot_difference_allowed numeric(14,2);
  v_snapshot_max_difference_days integer;
  v_snapshot_priority integer;
  v_locked_members jsonb;
  v_membership_snapshots jsonb;
  v_stored_count bigint;
  v_locked_count bigint;
  v_source_count bigint;
  v_destination_count bigint;
  v_source_ids uuid[];
  v_source_total numeric(14,2);
  v_source_total_cents bigint;
  v_bank_id uuid;
  v_bank_date date;
  v_bank_amount numeric(14,2);
  v_bank_amount_cents bigint;
  v_equation_cents bigint;
  v_membership_mismatch boolean;
  v_current_signature text;
  v_reconciliation_id uuid;
  v_failure_message text;
  v_failure_constraint text;
  v_stale_reason text;
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
  from public.financial_reconciliation_automatic_runs run
  where run.id = v_run_id
  for update;
  select * into strict v_proposal
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.id = p_proposal_id
  for update;

  if v_proposal.run_id <> v_run.id then
    raise exception 'Automation proposal run changed during execution.';
  end if;
  if v_run.actor is distinct from p_actor then
    raise exception 'Automatic analysis run belongs to another actor.';
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
  if v_proposal.status in (
    'ambiguous','skipped','deselected','stale','failed'
  ) then
    raise exception 'Automation proposal with status % cannot be executed.',
      v_proposal.status;
  end if;
  if v_proposal.status <> 'proposed' then
    raise exception 'Automation proposal is already being executed.';
  end if;
  if v_run.trigger not in ('manual','scheduled')
    or v_run.status not in ('ready','running','partial')
    or v_run.analysis_completed_at is null
    or v_run.finished_at is not null then
    raise exception 'Automatic proposal run is not executable.';
  end if;
  if v_proposal.rule_key <>
      'fdm_bank_transfer_cgd_bank_statement_combination'
    or v_proposal.rule_version <> 1
    or v_proposal.base_source_type <> 'import_fdm_accounts' then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_version_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_version_changed'
    );
  end if;

  begin
    select
      definition.definition,
      definition.display_name,
      definition.base_source_type,
      definition.destination_source_types,
      config.rule_version,
      config.enabled,
      config.allow_manual_execution,
      config.include_in_scheduled_batch,
      config.difference_allowed,
      config.max_difference_days,
      config.priority
    into
      v_current_definition,
      v_current_display_name,
      v_current_base_source_type,
      v_current_destination_source_types,
      v_current_rule_version,
      v_current_enabled,
      v_current_allow_manual,
      v_current_include_scheduled,
      v_current_difference_allowed,
      v_current_max_difference_days,
      v_current_priority
    from public.financial_reconciliation_automatic_rule_definitions definition
    join public.financial_reconciliation_automatic_rule_configs config
      on config.rule_key = definition.rule_key
     and config.rule_version = definition.version
    where definition.rule_key = v_proposal.rule_key
      and definition.version = v_proposal.rule_version
    for share of definition, config;

    select source_rule.operator into v_current_operator
    from public.financial_reconciliation_source_rules source_rule
    where source_rule.base_source_type = 'import_fdm_accounts'
      and source_rule.matching_source_type = 'import_cgd_extrato_ordem'
    for share of source_rule;

    if v_current_definition is null then
      raise exception 'Automatic specialized rule snapshot changed.';
    end if;
    if v_current_operator is null then
      raise exception 'Automatic specialized operator changed.';
    end if;
    if jsonb_typeof(v_run.definition_config_snapshot) is distinct from
        'array' then
      raise exception 'Automatic specialized rule snapshot changed.';
    end if;
    if jsonb_array_length(v_run.definition_config_snapshot) <> 1
      or jsonb_typeof(v_run.definition_config_snapshot->0) is distinct from
        'object' then
      raise exception 'Automatic specialized rule snapshot changed.';
    end if;
    v_rule_snapshot := v_run.definition_config_snapshot->0;
    if v_rule_snapshot - array[
        'ruleKey','ruleVersion','displayName','priority','differenceAllowed',
        'maxDifferenceDays','destinationSourceType','definition','operator'
      ]::text[] <> '{}'::jsonb
      or not (v_rule_snapshot ?& array[
        'ruleKey','ruleVersion','displayName','priority','differenceAllowed',
        'maxDifferenceDays','destinationSourceType','definition','operator'
      ])
      or v_rule_snapshot->>'ruleKey' is distinct from v_proposal.rule_key
      or coalesce(v_rule_snapshot->>'ruleVersion', '') !~ '^[1-9][0-9]*$'
      or length(v_rule_snapshot->>'ruleVersion') > 10
      or v_rule_snapshot->>'displayName' is distinct from
        'FDM Accounts – Bank Reservation Payments'
      or coalesce(v_rule_snapshot->>'priority', '') !~ '^[1-9][0-9]*$'
      or length(v_rule_snapshot->>'priority') > 10
      or v_rule_snapshot->>'destinationSourceType' is distinct from
        'import_cgd_extrato_ordem'
      or v_rule_snapshot->>'operator' is distinct from '+'
      or coalesce(v_rule_snapshot->>'differenceAllowed', '') !~
        '^0(?:\.0+)?$'
      or coalesce(v_rule_snapshot->>'maxDifferenceDays', '') !~ '^[0-9]+$'
      or length(v_rule_snapshot->>'maxDifferenceDays') > 10
      or v_rule_snapshot->'definition' is distinct from
        v_expected_definition then
      raise exception 'Automatic specialized rule snapshot changed.';
    end if;
    if (v_rule_snapshot->>'ruleVersion')::bigint not between 1 and 2147483647
      or (v_rule_snapshot->>'priority')::bigint not between 1 and 2147483647
      or (v_rule_snapshot->>'maxDifferenceDays')::bigint not between 0 and 90
      then
      raise exception 'Automatic specialized rule snapshot changed.';
    end if;

    v_snapshot_rule_version :=
      (v_rule_snapshot->>'ruleVersion')::integer;
    v_snapshot_difference_allowed :=
      (v_rule_snapshot->>'differenceAllowed')::numeric(14,2);
    v_snapshot_max_difference_days :=
      (v_rule_snapshot->>'maxDifferenceDays')::integer;
    v_snapshot_priority := (v_rule_snapshot->>'priority')::integer;

    if v_snapshot_rule_version <> v_proposal.rule_version
      or v_current_rule_version is distinct from v_snapshot_rule_version
      or not v_current_enabled
      or (v_run.trigger = 'manual' and not v_current_allow_manual)
      or (v_run.trigger = 'scheduled' and not v_current_include_scheduled)
      or v_current_definition is distinct from v_expected_definition
      or v_current_definition is distinct from v_rule_snapshot->'definition'
      or v_current_display_name is distinct from
        v_rule_snapshot->>'displayName'
      or v_current_base_source_type is distinct from 'import_fdm_accounts'
      or v_current_destination_source_types is distinct from
        '["import_cgd_extrato_ordem"]'::jsonb
      or v_current_max_difference_days is distinct from
        v_snapshot_max_difference_days
      or v_current_priority is distinct from v_snapshot_priority then
      raise exception 'Automatic specialized rule snapshot changed.';
    end if;
    if v_current_operator is distinct from '+' then
      raise exception 'Automatic specialized operator changed.';
    end if;
    if v_current_difference_allowed is distinct from 0::numeric
      or v_snapshot_difference_allowed is distinct from 0::numeric
      or v_proposal.allowed_difference is distinct from 0::numeric then
      raise exception 'Automatic specialized tolerance changed.';
    end if;

    if jsonb_typeof(v_proposal.summary_snapshot) is distinct from 'object'
      or jsonb_typeof(v_proposal.candidate_groups) is distinct from 'array'
      or jsonb_array_length(v_proposal.candidate_groups) <> 1
      or v_proposal.summary_snapshot->>'ruleKey' is distinct from
        v_proposal.rule_key
      or v_proposal.summary_snapshot->>'ruleVersion' is distinct from '1'
      or v_proposal.summary_snapshot->>'classification' is distinct from
        'proposed'
      or v_proposal.summary_snapshot->>'reason' is distinct from
        'unique_qualifying_combination'
      or v_proposal.reason is distinct from 'unique_qualifying_combination'
      or v_proposal.summary_snapshot->>'operator' is distinct from '+'
      or v_proposal.summary_snapshot->>'differenceAllowed' is distinct from '0'
      or v_proposal.summary_snapshot->>'maxDifferenceDays' is distinct from
        v_snapshot_max_difference_days::text
      or v_proposal.summary_snapshot->>'maxSourceRecords' is distinct from '10'
      or coalesce(
        v_proposal.summary_snapshot#>>'{bankAnchor,amount}', ''
      ) !~ '^-?[0-9]{1,12}(\.[0-9]{1,2})?$'
      or v_proposal.summary_snapshot->'candidateGroups' is distinct from
        v_proposal.candidate_groups
      or v_proposal.evidence is distinct from v_proposal.candidate_groups
      or v_proposal.summary_snapshot->>'signature' is distinct from
        v_proposal.signature
      or v_proposal.grouping_key is distinct from
        (v_proposal.summary_snapshot#>>'{bankAnchor,sourceId}') then
      raise exception 'Automatic specialized source snapshot changed.';
    end if;

    lock table
      public.import_cgd_extrato_ordem,
      public.import_fdm_accounts,
      public.financial_reconciliation_items
    in share row exclusive mode;

    select coalesce(jsonb_agg(jsonb_build_object(
      'role', locked_member.role,
      'source_type', locked_member.source_type,
      'source_id', locked_member.source_id,
      'ordinal', locked_member.ordinal,
      'source_date', locked_member.source_date,
      'amount', locked_member.amount,
      'description', locked_member.description,
      'account', locked_member.account,
      'row_snapshot', locked_member.row_snapshot,
      'live_exists', locked_member.live_exists
    ) order by locked_member.source_type, locked_member.source_id), '[]'::jsonb),
      count(*)
    into v_locked_members, v_locked_count
    from public.financial_reconciliation_lock_fdm_bank_automatic_members(
      v_proposal.id
    ) locked_member;

    select
      count(*),
      count(*) filter (where membership.role = 'source'),
      count(*) filter (where membership.role = 'destination'),
      array_agg(membership.source_id order by membership.ordinal)
        filter (where membership.role = 'source'),
      sum(membership.amount) filter (where membership.role = 'source'),
      sum(round(membership.amount * 100)::bigint)
        filter (where membership.role = 'source'),
      (array_agg(membership.source_id order by membership.ordinal)
        filter (where membership.role = 'destination'))[1],
      (array_agg(membership.source_date order by membership.ordinal)
        filter (where membership.role = 'destination'))[1],
      (array_agg(membership.amount order by membership.ordinal)
        filter (where membership.role = 'destination'))[1],
      jsonb_agg(jsonb_build_object(
        'role', membership.role,
        'sourceType', membership.source_type,
        'sourceId', membership.source_id,
        'ordinal', membership.ordinal,
        'sourceDate', membership.source_date,
        'amount', membership.amount,
        'description', membership.description,
        'account', membership.account,
        'rowSnapshot', membership.row_snapshot
      ) order by membership.source_type, membership.source_id)
    into
      v_stored_count, v_source_count, v_destination_count,
      v_source_ids, v_source_total, v_source_total_cents,
      v_bank_id, v_bank_date, v_bank_amount, v_membership_snapshots
    from public.financial_reconciliation_automatic_proposal_memberships membership
    where membership.proposal_id = v_proposal.id;
    v_bank_amount_cents := round(v_bank_amount * 100)::bigint;
    v_equation_cents := v_source_total_cents + v_bank_amount_cents;

    with live_members as (
      select *
      from jsonb_to_recordset(v_locked_members) as locked_member(
        role text, source_type text, source_id uuid, ordinal integer,
        source_date date, amount numeric, description text, account text,
        row_snapshot jsonb, live_exists boolean
      )
    )
    select exists (
      select 1
      from public.financial_reconciliation_automatic_proposal_memberships membership
      left join live_members locked_member
        on locked_member.role = membership.role
       and locked_member.source_type = membership.source_type
       and locked_member.source_id = membership.source_id
      where membership.proposal_id = v_proposal.id
        and (
          not coalesce(locked_member.live_exists, false)
          or membership.ordinal is distinct from locked_member.ordinal
          or membership.source_date is distinct from locked_member.source_date
          or membership.amount is distinct from locked_member.amount
          or membership.description is distinct from locked_member.description
          or membership.account is distinct from locked_member.account
          or membership.row_snapshot is distinct from
            locked_member.row_snapshot
        )
    ) into v_membership_mismatch;

    if v_stored_count <> v_locked_count
      or v_source_count not between 1 and 10
      or v_destination_count <> 1
      or v_bank_id is null or v_bank_date is null or v_bank_amount is null
      or v_membership_mismatch
      or v_equation_cents <> 0
      or v_source_total + v_bank_amount <> 0
      or v_proposal.calculated_difference <> 0
      or v_proposal.base_source_id is distinct from v_source_ids[1]
      or v_proposal.base_source_date is distinct from (
        select membership.source_date
        from public.financial_reconciliation_automatic_proposal_memberships membership
        where membership.proposal_id = v_proposal.id
          and membership.role = 'source'
        order by membership.ordinal limit 1
      )
      or exists (
        select 1
        from (
          select membership.*,
                 row_number() over (
                   partition by membership.role
                   order by membership.source_date, membership.source_id
                 )::integer as expected_ordinal
          from public.financial_reconciliation_automatic_proposal_memberships membership
          where membership.proposal_id = v_proposal.id
        ) membership
        where membership.ordinal <> membership.expected_ordinal
          or membership.source_date < date '2026-01-01'
          or (
            membership.role = 'source'
            and (
              membership.source_type <> 'import_fdm_accounts'
              or membership.account <> 'Bank Transfer'
              or abs(membership.source_date - v_bank_date) >
                v_snapshot_max_difference_days
              or membership.amount = 0
              or sign(round(membership.amount * 100)::bigint) <>
                -sign(v_bank_amount_cents)
            )
          )
          or (
            membership.role = 'destination'
            and (
              membership.source_type <> 'import_cgd_extrato_ordem'
              or membership.account <> ''
              or membership.ordinal <> 1
              or membership.source_id <> v_bank_id
            )
          )
      )
      or exists (
        select 1
        from public.financial_reconciliation_items locked
        join public.financial_reconciliation_automatic_proposal_memberships membership
          on membership.source_type = locked.source_type
         and membership.source_id = locked.source_id
        where membership.proposal_id = v_proposal.id
      )
      or exists (
        select 1
        from public.financial_reconciliation_automatic_proposal_memberships membership
        join public.financial_reconciliation_automatic_proposal_memberships overlap_member
          on overlap_member.source_type = membership.source_type
         and overlap_member.source_id = membership.source_id
         and overlap_member.proposal_id <> membership.proposal_id
        join public.financial_reconciliation_automatic_proposals overlap_proposal
          on overlap_proposal.id = overlap_member.proposal_id
        where membership.proposal_id = v_proposal.id
          and overlap_proposal.status in (
            'proposed','ambiguous','executing','completed'
          )
      )
      or v_proposal.summary_snapshot#>'{candidateGroups,0,fdmIds}'
        is distinct from to_jsonb(v_source_ids)
      or v_proposal.summary_snapshot#>'{candidateGroups,0,fdmTotalCents}'
        is distinct from to_jsonb(v_source_total_cents)
      or v_proposal.summary_snapshot#>'{candidateGroups,0,bankAmountCents}'
        is distinct from to_jsonb(v_bank_amount_cents)
      or v_proposal.summary_snapshot#>'{candidateGroups,0,equationCents}'
        is distinct from '0'::jsonb
      or v_proposal.summary_snapshot#>>'{bankAnchor,sourceId}' is distinct from
        v_bank_id::text
      or v_proposal.summary_snapshot#>>'{bankAnchor,sourceDate}' is distinct from
        v_bank_date::text
      or (v_proposal.summary_snapshot#>>'{bankAnchor,amount}')::numeric(14,2)
        is distinct from v_bank_amount
      or v_proposal.summary_snapshot#>'{bankAnchor,rowSnapshot}' is distinct from
        (select membership.row_snapshot
         from public.financial_reconciliation_automatic_proposal_memberships membership
         where membership.proposal_id = v_proposal.id
           and membership.role = 'destination')
      or v_proposal.base_snapshot is distinct from (
        select jsonb_build_object(
          'sourceType', membership.source_type,
          'sourceId', membership.source_id,
          'sourceDate', membership.source_date,
          'amount', membership.amount,
          'description', membership.description,
          'account', membership.account,
          'rowSnapshot', membership.row_snapshot
        )
        from public.financial_reconciliation_automatic_proposal_memberships membership
        where membership.proposal_id = v_proposal.id
          and membership.role = 'source'
        order by membership.ordinal limit 1
      ) then
      raise exception 'Automatic specialized source snapshot changed.';
    end if;

    v_current_signature := public.financial_reconciliation_extension_sha256(
      jsonb_build_object(
        'ruleKey', v_proposal.rule_key,
        'ruleVersion', v_proposal.rule_version,
        'bankId', v_bank_id,
        'bankDate', v_bank_date,
        'bankAmount', v_bank_amount,
        'classification', 'proposed',
        'reason', 'unique_qualifying_combination',
        'candidateGroups', v_proposal.candidate_groups,
        'operator', '+',
        'maxDifferenceDays', v_snapshot_max_difference_days
      )::text
    );
    if v_current_signature is distinct from v_proposal.signature then
      raise exception 'Automatic specialized source snapshot changed.';
    end if;

    v_reconciliation_id :=
      public.financial_reconciliation_commit_fdm_bank_automatic_proposal(
        v_proposal.id,
        p_actor,
        'import_fdm_accounts',
        'import_cgd_extrato_ordem',
        '+',
        0,
        null,
        jsonb_build_object(
          'ruleKey', v_proposal.rule_key,
          'ruleVersion', v_proposal.rule_version,
          'displayName', v_rule_snapshot->>'displayName',
          'definition', v_rule_snapshot->'definition'
        ),
        jsonb_build_object(
          'differenceAllowed', v_snapshot_difference_allowed,
          'maxDifferenceDays', v_snapshot_max_difference_days,
          'priority', v_snapshot_priority
        ),
        jsonb_build_object(
          'import_cgd_extrato_ordem', v_current_operator
        ),
        v_membership_snapshots
      );
  exception
    when unique_violation then
      get stacked diagnostics v_failure_constraint = constraint_name;
      if v_failure_constraint =
          'financial_reconciliation_items_source_type_source_id_key' then
        v_stale_reason := 'source_snapshot_changed';
      else
        update public.financial_reconciliation_automatic_proposals
        set status = 'failed', reason = 'execution_failed',
            reconciliation_id = null, completed_at = null,
            error = 'Automatic reconciliation execution failed.',
            error_detail = '', updated_at = now()
        where id = v_proposal.id;
        return jsonb_build_object(
          'proposalId', v_proposal.id, 'runId', v_run.id,
          'status', 'failed', 'reason', 'execution_failed'
        );
      end if;
    when others then
      get stacked diagnostics v_failure_message = message_text;
      v_stale_reason := case v_failure_message
        when 'Automatic specialized rule snapshot changed.'
          then 'rule_snapshot_changed'
        when 'Automatic specialized operator changed.'
          then 'operator_changed'
        when 'Automatic specialized tolerance changed.'
          then 'tolerance_changed'
        when 'Automatic specialized source snapshot changed.'
          then 'source_snapshot_changed'
        when 'Automatic grouped reconciliation lifecycle snapshots changed after revalidation.'
          then 'source_snapshot_changed'
        when 'This record is already reconciled.'
          then 'source_snapshot_changed'
        else null
      end;
      if v_stale_reason is null then
        update public.financial_reconciliation_automatic_proposals
        set status = 'failed', reason = 'execution_failed',
            reconciliation_id = null, completed_at = null,
            error = 'Automatic reconciliation execution failed.',
            error_detail = '', updated_at = now()
        where id = v_proposal.id;
        return jsonb_build_object(
          'proposalId', v_proposal.id, 'runId', v_run.id,
          'status', 'failed', 'reason', 'execution_failed'
        );
      end if;
  end;

  if v_stale_reason is not null then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = v_stale_reason,
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', v_stale_reason
    );
  end if;

  return jsonb_build_object(
    'proposalId', v_proposal.id,
    'runId', v_run.id,
    'status', 'completed',
    'reconciliationId', v_reconciliation_id
  );
end
$$;

create or replace function public.financial_reconciliation_execute_adyen_monthly_proposal(
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
  v_expected_definition jsonb := jsonb_build_object(
    'strategy', 'closed_calendar_month',
    'bankDescriptionContains', 'Adyen',
    'fdmAccount', 'Adyen',
    'requiresBothSides', true,
    'monthMarkerDays', 31
  );
  v_current_definition jsonb;
  v_current_display_name text;
  v_current_base_source_type text;
  v_current_destination_source_types jsonb;
  v_current_rule_version integer;
  v_current_enabled boolean;
  v_current_allow_manual boolean;
  v_current_include_scheduled boolean;
  v_current_difference_allowed numeric(14,2);
  v_current_max_difference_days integer;
  v_current_priority integer;
  v_current_operator text;
  v_snapshot_rule_version integer;
  v_snapshot_difference_allowed numeric(14,2);
  v_snapshot_max_difference_days integer;
  v_snapshot_priority integer;
  v_calendar_month date;
  v_locked_members jsonb;
  v_membership_snapshots jsonb;
  v_stored_count bigint;
  v_locked_count bigint;
  v_source_count bigint;
  v_destination_count bigint;
  v_source_ids uuid[];
  v_destination_ids uuid[];
  v_source_total numeric(14,2);
  v_destination_total numeric(14,2);
  v_live_difference numeric(14,2);
  v_technical_base_id uuid;
  v_technical_base_date date;
  v_technical_base_snapshot jsonb;
  v_membership_mismatch boolean;
  v_month_membership_mismatch boolean;
  v_current_signature text;
  v_comment text;
  v_reconciliation_id uuid;
  v_failure_message text;
  v_failure_constraint text;
  v_stale_reason text;
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
  from public.financial_reconciliation_automatic_runs run
  where run.id = v_run_id
  for update;
  select * into strict v_proposal
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.id = p_proposal_id
  for update;

  if v_proposal.run_id <> v_run.id then
    raise exception 'Automation proposal run changed during execution.';
  end if;
  if v_run.actor is distinct from p_actor then
    raise exception 'Automatic analysis run belongs to another actor.';
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
  if v_proposal.status in (
    'ambiguous','skipped','deselected','stale','failed'
  ) then
    raise exception 'Automation proposal with status % cannot be executed.',
      v_proposal.status;
  end if;
  if v_proposal.status <> 'proposed' then
    raise exception 'Automation proposal is already being executed.';
  end if;
  if v_run.trigger not in ('manual','scheduled')
    or v_run.status not in ('ready','running','partial')
    or v_run.analysis_completed_at is null
    or v_run.finished_at is not null then
    raise exception 'Automatic proposal run is not executable.';
  end if;
  if v_proposal.rule_key <>
      'cgd_bank_statement_fdm_adyen_monthly_payments'
    or v_proposal.rule_version <> 1
    or v_proposal.base_source_type <> 'import_cgd_extrato_ordem' then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_version_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_version_changed'
    );
  end if;

  begin
    select
      definition.definition,
      definition.display_name,
      definition.base_source_type,
      definition.destination_source_types,
      config.rule_version,
      config.enabled,
      config.allow_manual_execution,
      config.include_in_scheduled_batch,
      config.difference_allowed,
      config.max_difference_days,
      config.priority
    into
      v_current_definition,
      v_current_display_name,
      v_current_base_source_type,
      v_current_destination_source_types,
      v_current_rule_version,
      v_current_enabled,
      v_current_allow_manual,
      v_current_include_scheduled,
      v_current_difference_allowed,
      v_current_max_difference_days,
      v_current_priority
    from public.financial_reconciliation_automatic_rule_definitions definition
    join public.financial_reconciliation_automatic_rule_configs config
      on config.rule_key = definition.rule_key
     and config.rule_version = definition.version
    where definition.rule_key = v_proposal.rule_key
      and definition.version = v_proposal.rule_version
    for share of definition, config;

    select source_rule.operator into v_current_operator
    from public.financial_reconciliation_source_rules source_rule
    where source_rule.base_source_type = 'import_cgd_extrato_ordem'
      and source_rule.matching_source_type = 'import_fdm_accounts'
    for share of source_rule;

    if v_current_definition is null then
      raise exception 'Automatic specialized rule snapshot changed.';
    end if;
    if v_current_operator is null then
      raise exception 'Automatic specialized operator changed.';
    end if;
    if jsonb_typeof(v_run.definition_config_snapshot) is distinct from
        'array' then
      raise exception 'Automatic specialized rule snapshot changed.';
    end if;
    if jsonb_array_length(v_run.definition_config_snapshot) <> 1
      or jsonb_typeof(v_run.definition_config_snapshot->0) is distinct from
        'object' then
      raise exception 'Automatic specialized rule snapshot changed.';
    end if;
    v_rule_snapshot := v_run.definition_config_snapshot->0;
    if v_rule_snapshot - array[
        'ruleKey','ruleVersion','displayName','priority','differenceAllowed',
        'maxDifferenceDays','destinationSourceType','definition','operator'
      ]::text[] <> '{}'::jsonb
      or not (v_rule_snapshot ?& array[
        'ruleKey','ruleVersion','displayName','priority','differenceAllowed',
        'maxDifferenceDays','destinationSourceType','definition','operator'
      ])
      or v_rule_snapshot->>'ruleKey' is distinct from v_proposal.rule_key
      or coalesce(v_rule_snapshot->>'ruleVersion', '') !~ '^[1-9][0-9]*$'
      or length(v_rule_snapshot->>'ruleVersion') > 10
      or v_rule_snapshot->>'displayName' is distinct from
        'FDM Accounts – Adyen Reservation Payments'
      or coalesce(v_rule_snapshot->>'priority', '') !~ '^[1-9][0-9]*$'
      or length(v_rule_snapshot->>'priority') > 10
      or v_rule_snapshot->>'destinationSourceType' is distinct from
        'import_fdm_accounts'
      or v_rule_snapshot->>'operator' is distinct from '-'
      or coalesce(v_rule_snapshot->>'differenceAllowed', '') !~
        '^[0-9]{1,12}(\.[0-9]{1,2})?$'
      or coalesce(v_rule_snapshot->>'maxDifferenceDays', '') !~ '^[0-9]+$'
      or length(v_rule_snapshot->>'maxDifferenceDays') > 10
      or v_rule_snapshot->'definition' is distinct from
        v_expected_definition then
      raise exception 'Automatic specialized rule snapshot changed.';
    end if;
    if (v_rule_snapshot->>'ruleVersion')::bigint not between 1 and 2147483647
      or (v_rule_snapshot->>'priority')::bigint not between 1 and 2147483647
      or (v_rule_snapshot->>'maxDifferenceDays')::bigint <> 31 then
      raise exception 'Automatic specialized rule snapshot changed.';
    end if;

    v_snapshot_rule_version :=
      (v_rule_snapshot->>'ruleVersion')::integer;
    v_snapshot_difference_allowed :=
      (v_rule_snapshot->>'differenceAllowed')::numeric(14,2);
    v_snapshot_max_difference_days :=
      (v_rule_snapshot->>'maxDifferenceDays')::integer;
    v_snapshot_priority := (v_rule_snapshot->>'priority')::integer;

    if v_snapshot_rule_version <> v_proposal.rule_version
      or v_current_rule_version is distinct from v_snapshot_rule_version
      or not v_current_enabled
      or (v_run.trigger = 'manual' and not v_current_allow_manual)
      or (v_run.trigger = 'scheduled' and not v_current_include_scheduled)
      or v_current_definition is distinct from v_expected_definition
      or v_current_definition is distinct from v_rule_snapshot->'definition'
      or v_current_display_name is distinct from
        v_rule_snapshot->>'displayName'
      or v_current_base_source_type is distinct from
        'import_cgd_extrato_ordem'
      or v_current_destination_source_types is distinct from
        '["import_fdm_accounts"]'::jsonb
      or v_current_max_difference_days is distinct from 31
      or v_current_priority is distinct from v_snapshot_priority then
      raise exception 'Automatic specialized rule snapshot changed.';
    end if;
    if v_current_operator is distinct from '-' then
      raise exception 'Automatic specialized operator changed.';
    end if;
    if v_current_difference_allowed is distinct from
        v_snapshot_difference_allowed
      or v_proposal.allowed_difference is distinct from
        v_snapshot_difference_allowed then
      raise exception 'Automatic specialized tolerance changed.';
    end if;

    if jsonb_typeof(v_proposal.summary_snapshot) is distinct from 'object'
      or coalesce(v_proposal.grouping_key, '') !~
        '^\d{4}-(0[1-9]|1[0-2])$'
      or substring(v_proposal.grouping_key from 1 for 4) = '0000'
      or coalesce(v_proposal.summary_snapshot->>'calendarMonth', '') !~
        '^\d{4}-(0[1-9]|1[0-2])-01$'
      or v_proposal.summary_snapshot->>'calendarMonth' is distinct from
        v_proposal.grouping_key || '-01'
      or v_proposal.summary_snapshot->>'ruleKey' is distinct from
        v_proposal.rule_key
      or v_proposal.summary_snapshot->>'ruleVersion' is distinct from '1'
      or v_proposal.summary_snapshot->>'strategy' is distinct from
        'closed_calendar_month'
      or v_proposal.summary_snapshot->>'sourceDescriptionContains' is distinct
        from 'Adyen'
      or v_proposal.summary_snapshot->>'destinationAccount' is distinct from
        'Adyen'
      or v_proposal.summary_snapshot->>'operator' is distinct from '-'
      or v_proposal.summary_snapshot->>'maxDifferenceDays' is distinct from '31'
      or coalesce(v_proposal.summary_snapshot->>'sourceCount', '') !~
        '^[1-9][0-9]*$'
      or length(v_proposal.summary_snapshot->>'sourceCount') > 10
      or coalesce(v_proposal.summary_snapshot->>'destinationCount', '') !~
        '^[1-9][0-9]*$'
      or length(v_proposal.summary_snapshot->>'destinationCount') > 10
      or coalesce(v_proposal.summary_snapshot->>'sourceTotal', '') !~
        '^-?[0-9]{1,12}(\.[0-9]{1,2})?$'
      or coalesce(v_proposal.summary_snapshot->>'destinationTotal', '') !~
        '^-?[0-9]{1,12}(\.[0-9]{1,2})?$'
      or coalesce(v_proposal.summary_snapshot->>'calculatedDifference', '') !~
        '^-?[0-9]{1,12}(\.[0-9]{1,2})?$'
      or coalesce(v_proposal.summary_snapshot->>'differenceAllowed', '') !~
        '^[0-9]{1,12}(\.[0-9]{1,2})?$'
      or coalesce(v_proposal.summary_snapshot->>'technicalBaseSourceId', '') !~
        '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
      or coalesce(v_proposal.summary_snapshot->>'technicalBaseSourceDate', '') !~
        '^\d{4}-\d{2}-\d{2}$'
      or v_proposal.summary_snapshot->>'signature' is distinct from
        v_proposal.signature then
      raise exception 'Automatic specialized source snapshot changed.';
    end if;
    v_calendar_month :=
      (v_proposal.summary_snapshot->>'calendarMonth')::date;
    if v_calendar_month < date '2026-01-01'
      or v_calendar_month >= date_trunc('month', current_date)::date
      or v_calendar_month <> date_trunc('month', v_calendar_month)::date then
      raise exception 'Automatic specialized source snapshot changed.';
    end if;

    lock table
      public.import_cgd_extrato_ordem,
      public.import_fdm_accounts,
      public.financial_reconciliation_items
    in share row exclusive mode;

    select coalesce(jsonb_agg(jsonb_build_object(
      'role', locked_member.role,
      'source_type', locked_member.source_type,
      'source_id', locked_member.source_id,
      'ordinal', locked_member.ordinal,
      'source_date', locked_member.source_date,
      'amount', locked_member.amount,
      'description', locked_member.description,
      'account', locked_member.account,
      'row_snapshot', locked_member.row_snapshot,
      'live_exists', locked_member.live_exists
    ) order by locked_member.source_type, locked_member.source_id), '[]'::jsonb),
      count(*)
    into v_locked_members, v_locked_count
    from public.financial_reconciliation_lock_fdm_bank_automatic_members(
      v_proposal.id
    ) locked_member;

    select
      count(*),
      count(*) filter (where membership.role = 'source'),
      count(*) filter (where membership.role = 'destination'),
      array_agg(membership.source_id order by membership.ordinal)
        filter (where membership.role = 'source'),
      array_agg(membership.source_id order by membership.ordinal)
        filter (where membership.role = 'destination'),
      sum(membership.amount) filter (where membership.role = 'source'),
      sum(membership.amount) filter (where membership.role = 'destination'),
      (array_agg(membership.source_id order by membership.ordinal)
        filter (where membership.role = 'source'))[1],
      (array_agg(membership.source_date order by membership.ordinal)
        filter (where membership.role = 'source'))[1],
      (array_agg(membership.row_snapshot order by membership.ordinal)
        filter (where membership.role = 'source'))[1],
      jsonb_agg(jsonb_build_object(
        'role', membership.role,
        'sourceType', membership.source_type,
        'sourceId', membership.source_id,
        'ordinal', membership.ordinal,
        'sourceDate', membership.source_date,
        'amount', membership.amount,
        'description', membership.description,
        'account', membership.account,
        'rowSnapshot', membership.row_snapshot
      ) order by membership.source_type, membership.source_id)
    into
      v_stored_count, v_source_count, v_destination_count,
      v_source_ids, v_destination_ids, v_source_total,
      v_destination_total, v_technical_base_id,
      v_technical_base_date, v_technical_base_snapshot,
      v_membership_snapshots
    from public.financial_reconciliation_automatic_proposal_memberships membership
    where membership.proposal_id = v_proposal.id;

    with live_members as (
      select *
      from jsonb_to_recordset(v_locked_members) as locked_member(
        role text, source_type text, source_id uuid, ordinal integer,
        source_date date, amount numeric, description text, account text,
        row_snapshot jsonb, live_exists boolean
      )
    )
    select exists (
      select 1
      from public.financial_reconciliation_automatic_proposal_memberships membership
      left join live_members locked_member
        on locked_member.role = membership.role
       and locked_member.source_type = membership.source_type
       and locked_member.source_id = membership.source_id
      where membership.proposal_id = v_proposal.id
        and (
          not coalesce(locked_member.live_exists, false)
          or membership.ordinal is distinct from locked_member.ordinal
          or membership.source_date is distinct from locked_member.source_date
          or membership.amount is distinct from locked_member.amount
          or membership.description is distinct from locked_member.description
          or membership.account is distinct from locked_member.account
          or membership.row_snapshot is distinct from
            locked_member.row_snapshot
        )
    ) into v_membership_mismatch;

    with stored as (
      select membership.role, membership.source_type,
             membership.source_id, membership.source_date,
             membership.amount, membership.description, membership.account
      from public.financial_reconciliation_automatic_proposal_memberships membership
      where membership.proposal_id = v_proposal.id
    ), live as (
      select 'source'::text as role,
             'import_cgd_extrato_ordem'::text as source_type,
             bank.id as source_id, bank.data as source_date,
             bank.montante as amount, bank.descritivo as description,
             ''::text as account
      from public.import_cgd_extrato_ordem bank
      where bank.data >= v_calendar_month
        and bank.data < (v_calendar_month + interval '1 month')::date
        and bank.data >= date '2026-01-01'
        and bank.data < date_trunc('month', current_date)::date
        and bank.montante is not null
        and bank.descritivo ilike '%Adyen%'
        and not exists (
          select 1 from public.financial_reconciliation_items locked
          where locked.source_type = 'import_cgd_extrato_ordem'
            and locked.source_id = bank.id
        )
      union all
      select 'destination'::text, 'import_fdm_accounts'::text,
             fdm.id, fdm.event_date, fdm.amount,
             fdm.description, fdm.account
      from public.import_fdm_accounts fdm
      where fdm.event_date >= v_calendar_month
        and fdm.event_date < (v_calendar_month + interval '1 month')::date
        and fdm.event_date >= date '2026-01-01'
        and fdm.event_date < date_trunc('month', current_date)::date
        and fdm.amount is not null
        and fdm.account = 'Adyen'
        and not exists (
          select 1 from public.financial_reconciliation_items locked
          where locked.source_type = 'import_fdm_accounts'
            and locked.source_id = fdm.id
        )
    ), lost as (
      select * from stored except select * from live
    ), gained as (
      select * from live except select * from stored
    )
    select exists (select 1 from lost) or exists (select 1 from gained)
    into v_month_membership_mismatch;

    v_live_difference :=
      (v_source_total - v_destination_total)::numeric(14,2);
    if v_stored_count <> v_locked_count
      or v_source_count < 1
      or v_destination_count < 1
      or v_membership_mismatch
      or v_month_membership_mismatch
      or exists (
        select 1
        from (
          select membership.*,
                 row_number() over (
                   partition by membership.role
                   order by membership.source_date, membership.source_id
                 )::integer as expected_ordinal
          from public.financial_reconciliation_automatic_proposal_memberships membership
          where membership.proposal_id = v_proposal.id
        ) membership
        where membership.ordinal <> membership.expected_ordinal
          or membership.source_date < v_calendar_month
          or membership.source_date >=
            (v_calendar_month + interval '1 month')::date
          or (
            membership.role = 'source'
            and (
              membership.source_type <> 'import_cgd_extrato_ordem'
              or membership.account <> ''
              or membership.description not ilike '%Adyen%'
            )
          )
          or (
            membership.role = 'destination'
            and (
              membership.source_type <> 'import_fdm_accounts'
              or membership.account <> 'Adyen'
            )
          )
      )
      or exists (
        select 1
        from public.financial_reconciliation_items locked
        join public.financial_reconciliation_automatic_proposal_memberships membership
          on membership.source_type = locked.source_type
         and membership.source_id = locked.source_id
        where membership.proposal_id = v_proposal.id
      )
      or exists (
        select 1
        from public.financial_reconciliation_automatic_proposal_memberships membership
        join public.financial_reconciliation_automatic_proposal_memberships overlap_member
          on overlap_member.source_type = membership.source_type
         and overlap_member.source_id = membership.source_id
         and overlap_member.proposal_id <> membership.proposal_id
        join public.financial_reconciliation_automatic_proposals overlap_proposal
          on overlap_proposal.id = overlap_member.proposal_id
        where membership.proposal_id = v_proposal.id
          and overlap_proposal.status in (
            'proposed','ambiguous','executing','completed'
          )
      )
      or (v_proposal.summary_snapshot->>'sourceCount')::bigint <>
        v_source_count
      or (v_proposal.summary_snapshot->>'destinationCount')::bigint <>
        v_destination_count
      or (v_proposal.summary_snapshot->>'sourceTotal')::numeric(14,2)
        is distinct from v_source_total
      or (v_proposal.summary_snapshot->>'destinationTotal')::numeric(14,2)
        is distinct from v_destination_total
      or (v_proposal.summary_snapshot->>'calculatedDifference')::numeric(14,2)
        is distinct from v_live_difference
      or (v_proposal.summary_snapshot->>'differenceAllowed')::numeric(14,2)
        is distinct from v_snapshot_difference_allowed
      or v_proposal.calculated_difference is distinct from v_live_difference
      or abs(v_live_difference) > v_snapshot_difference_allowed
      or v_proposal.base_source_id is distinct from v_technical_base_id
      or v_proposal.base_source_date is distinct from v_technical_base_date
      or v_proposal.base_snapshot is distinct from v_technical_base_snapshot
      or v_proposal.summary_snapshot->>'technicalBaseSourceId' is distinct from
        v_technical_base_id::text
      or v_proposal.summary_snapshot->>'technicalBaseSourceDate' is distinct from
        v_technical_base_date::text then
      raise exception 'Automatic specialized source snapshot changed.';
    end if;

    v_current_signature := public.financial_reconciliation_extension_sha256(
      jsonb_build_object(
        'ruleKey', v_proposal.rule_key,
        'ruleVersion', v_proposal.rule_version,
        'calendarMonth', v_calendar_month,
        'sourceIds', to_jsonb(v_source_ids),
        'destinationIds', to_jsonb(v_destination_ids),
        'sourceTotal', v_source_total,
        'destinationTotal', v_destination_total,
        'calculatedDifference', v_live_difference,
        'differenceAllowed', v_snapshot_difference_allowed,
        'operator', '-'
      )::text
    );
    if v_current_signature is distinct from v_proposal.signature then
      raise exception 'Automatic specialized source snapshot changed.';
    end if;

    if v_live_difference = 0 then
      v_comment := null;
    else
      v_comment := 'FDM Accounts – Adyen Reservation Payments'
        || ' | month ' || to_char(v_calendar_month, 'YYYY-MM')
        || ' | difference '
        || to_char(v_live_difference, 'FM999999999990.00')
        || ' EUR | allowance '
        || to_char(v_snapshot_difference_allowed, 'FM999999999990.00')
        || ' EUR.';
    end if;

    v_reconciliation_id :=
      public.financial_reconciliation_commit_fdm_bank_automatic_proposal(
        v_proposal.id,
        p_actor,
        'import_cgd_extrato_ordem',
        'import_fdm_accounts',
        '-',
        v_live_difference,
        v_comment,
        jsonb_build_object(
          'ruleKey', v_proposal.rule_key,
          'ruleVersion', v_proposal.rule_version,
          'displayName', v_rule_snapshot->>'displayName',
          'definition', v_rule_snapshot->'definition'
        ),
        jsonb_build_object(
          'differenceAllowed', v_snapshot_difference_allowed,
          'maxDifferenceDays', v_snapshot_max_difference_days,
          'priority', v_snapshot_priority
        ),
        jsonb_build_object('import_fdm_accounts', v_current_operator),
        v_membership_snapshots
      );
  exception
    when unique_violation then
      get stacked diagnostics v_failure_constraint = constraint_name;
      if v_failure_constraint =
          'financial_reconciliation_items_source_type_source_id_key' then
        v_stale_reason := 'source_snapshot_changed';
      else
        update public.financial_reconciliation_automatic_proposals
        set status = 'failed', reason = 'execution_failed',
            reconciliation_id = null, completed_at = null,
            error = 'Automatic reconciliation execution failed.',
            error_detail = '', updated_at = now()
        where id = v_proposal.id;
        return jsonb_build_object(
          'proposalId', v_proposal.id, 'runId', v_run.id,
          'status', 'failed', 'reason', 'execution_failed'
        );
      end if;
    when others then
      get stacked diagnostics v_failure_message = message_text;
      v_stale_reason := case v_failure_message
        when 'Automatic specialized rule snapshot changed.'
          then 'rule_snapshot_changed'
        when 'Automatic specialized operator changed.'
          then 'operator_changed'
        when 'Automatic specialized tolerance changed.'
          then 'tolerance_changed'
        when 'Automatic specialized source snapshot changed.'
          then 'source_snapshot_changed'
        when 'Automatic grouped reconciliation lifecycle snapshots changed after revalidation.'
          then 'source_snapshot_changed'
        when 'This record is already reconciled.'
          then 'source_snapshot_changed'
        else null
      end;
      if v_stale_reason is null then
        update public.financial_reconciliation_automatic_proposals
        set status = 'failed', reason = 'execution_failed',
            reconciliation_id = null, completed_at = null,
            error = 'Automatic reconciliation execution failed.',
            error_detail = '', updated_at = now()
        where id = v_proposal.id;
        return jsonb_build_object(
          'proposalId', v_proposal.id, 'runId', v_run.id,
          'status', 'failed', 'reason', 'execution_failed'
        );
      end if;
  end;

  if v_stale_reason is not null then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = v_stale_reason,
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', v_stale_reason
    );
  end if;

  return jsonb_build_object(
    'proposalId', v_proposal.id,
    'runId', v_run.id,
    'status', 'completed',
    'reconciliationId', v_reconciliation_id
  );
end
$$;

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
  v_run_actor text;
  v_rule_key text;
  v_rule_version integer;
  v_result jsonb;
  v_authoritative_status text;
  v_authoritative_reconciliation_id uuid;
  v_authoritative_run_status text;
begin
  if p_proposal_id is null then
    raise exception 'Automation proposal ID is required.';
  end if;
  if nullif(trim(coalesce(p_actor, '')), '') is null then
    raise exception 'Actor is required.';
  end if;

  select proposal.run_id, run.actor,
         proposal.rule_key, proposal.rule_version
  into v_run_id, v_run_actor, v_rule_key, v_rule_version
  from public.financial_reconciliation_automatic_proposals proposal
  join public.financial_reconciliation_automatic_runs run
    on run.id = proposal.run_id
  where proposal.id = p_proposal_id
  for update of run, proposal;
  if not found then
    raise exception 'Automation proposal was not found.';
  end if;
  if v_run_actor is distinct from p_actor then
    raise exception 'Automatic analysis run belongs to another actor.';
  end if;

  if v_rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income'
    and v_rule_version = 2 then
    v_result := public.financial_reconciliation_execute_monthly_income_proposal(
      p_proposal_id, p_actor
    );
  elsif (v_rule_key, v_rule_version) in (
    ('financial_documents_cgd_bank_statement', 2),
    ('financial_documents_cgd_credit_card', 1),
    ('financial_documents_cgd_bank_statement_amount_only', 1),
    ('financial_documents_cgd_credit_card_amount_only', 1)
  ) then
    v_result := public.financial_reconciliation_execute_prior_proposal(
      p_proposal_id, p_actor
    );
  elsif v_rule_key =
      'fdm_bank_transfer_cgd_bank_statement_combination'
    and v_rule_version = 1 then
    v_result :=
      public.financial_reconciliation_execute_bank_reservation_proposal(
        p_proposal_id, p_actor
      );
  elsif v_rule_key =
      'cgd_bank_statement_fdm_adyen_monthly_payments'
    and v_rule_version = 1 then
    v_result :=
      public.financial_reconciliation_execute_adyen_monthly_proposal(
        p_proposal_id, p_actor
      );
  else
    raise exception 'Automation proposal rule/version is not executable.';
  end if;

  -- Reload both rows after strategy execution. The caller still owns the
  -- existing finish_financial_reconciliation_automatic_run step, but receives
  -- only an outcome that agrees with authoritative persisted lifecycle state.
  select proposal.status, proposal.reconciliation_id
  into strict v_authoritative_status, v_authoritative_reconciliation_id
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.id = p_proposal_id;
  select run.status into strict v_authoritative_run_status
  from public.financial_reconciliation_automatic_runs run
  where run.id = v_run_id;

  if v_result->>'status' is distinct from v_authoritative_status
    or (
      v_authoritative_status = 'completed'
      and (
        v_authoritative_reconciliation_id is null
        or v_result->>'reconciliationId' is distinct from
          v_authoritative_reconciliation_id::text
      )
    )
    or v_authoritative_run_status not in (
      'ready','running','completed','partial','failed'
    ) then
    raise exception 'Automatic proposal authoritative outcome changed.';
  end if;

  return v_result;
end
$$;

revoke all on function public.financial_reconciliation_automatic_adyen_month_count()
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_automatic_adyen_month_page(date,integer)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_continue_automatic_adyen_monthly(uuid,text)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_finalize_automatic_prior_analysis(uuid)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_finalize_automatic_analysis(uuid)
  from public, anon, authenticated, service_role;
revoke all on function public.continue_financial_reconciliation_automatic_analysis(uuid,text)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_lock_fdm_bank_automatic_members(uuid)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_commit_fdm_bank_automatic_proposal(uuid,text,text,text,text,numeric,text,jsonb,jsonb,jsonb,jsonb)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_execute_bank_reservation_proposal(uuid,text)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_execute_adyen_monthly_proposal(uuid,text)
  from public, anon, authenticated, service_role;
revoke all on function public.execute_financial_reconciliation_automatic_proposal(uuid,text)
  from public, anon, authenticated, service_role;
grant execute on function public.financial_reconciliation_finalize_automatic_analysis(uuid)
  to service_role;
grant execute on function public.continue_financial_reconciliation_automatic_analysis(uuid,text)
  to service_role;
grant execute on function public.execute_financial_reconciliation_automatic_proposal(uuid,text)
  to service_role;

create or replace function public.get_financial_reconciliation_automatic_manual_rules()
returns jsonb
language plpgsql
stable
security definer set search_path = public, pg_temp
as $$
declare
  v_rules jsonb;
  v_config_count integer;
begin
  select count(*)::integer into v_config_count
  from public.financial_reconciliation_automatic_rule_configs config;
  if v_config_count <> 7 or exists (
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
    raise exception 'Automatic reconciliation managed catalog is invalid.';
  end if;

  select coalesce(jsonb_agg(jsonb_build_object(
    'ruleKey', definition.rule_key,
    'ruleVersion', config.rule_version,
    'displayName', definition.display_name,
    'baseSourceType', definition.base_source_type,
    'destinationSourceTypes', definition.destination_source_types,
    'logicDescription', definition.logic_description,
    'definition', definition.definition,
    'enabled', config.enabled,
    'allowManualExecution', config.allow_manual_execution,
    'differenceAllowed', config.difference_allowed,
    'maxDifferenceDays', config.max_difference_days,
    'priority', config.priority,
    'operator', source_rule.operator
  ) order by config.priority, definition.rule_key), '[]'::jsonb)
  into v_rules
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
  where config.enabled
    and config.allow_manual_execution
    and jsonb_array_length(definition.destination_source_types) = 1
    and source_rule.operator = case when config.rule_key in (
      'cgd_bank_statement_fdm_credit_card_monthly_income',
      'cgd_bank_statement_fdm_adyen_monthly_payments'
    ) then '-' else '+' end
    and (config.rule_key, config.rule_version) in (
      ('financial_documents_cgd_bank_statement', 2),
      ('financial_documents_cgd_credit_card', 1),
      ('financial_documents_cgd_bank_statement_amount_only', 1),
      ('financial_documents_cgd_credit_card_amount_only', 1),
      ('cgd_bank_statement_fdm_credit_card_monthly_income', 2),
      ('fdm_bank_transfer_cgd_bank_statement_combination', 1),
      ('cgd_bank_statement_fdm_adyen_monthly_payments', 1)
    );

  return jsonb_build_object('rules', v_rules);
end
$$;

create or replace function public.get_financial_reconciliation_automation_settings()
returns jsonb
language plpgsql
stable
security definer set search_path = public, pg_temp
as $$
declare
  v_schedule jsonb;
  v_rules jsonb;
  v_config_count integer;
  v_last_scheduled_batch jsonb;
begin
  select jsonb_build_object(
    'enabled', schedule.enabled,
    'timeOfDay', to_char(schedule.time_of_day, 'HH24:MI'),
    'timeZone', schedule.time_zone,
    'updatedBy', schedule.updated_by,
    'updatedAt', schedule.updated_at
  )
  into v_schedule
  from public.financial_reconciliation_automatic_schedule schedule
  where schedule.id = true;

  select count(*)::integer into v_config_count
  from public.financial_reconciliation_automatic_rule_configs config;
  if v_config_count <> 7 or exists (
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
    raise exception 'Automatic reconciliation managed catalog is invalid.';
  end if;

  select coalesce(jsonb_agg(jsonb_build_object(
    'ruleKey', definition.rule_key,
    'ruleVersion', config.rule_version,
    'displayName', definition.display_name,
    'baseSourceType', definition.base_source_type,
    'destinationSourceTypes', definition.destination_source_types,
    'logicDescription', definition.logic_description,
    'definition', definition.definition,
    'enabled', config.enabled,
    'allowManualExecution', config.allow_manual_execution,
    'includeInScheduledBatch', config.include_in_scheduled_batch,
    'differenceAllowed', config.difference_allowed,
    'maxDifferenceDays', config.max_difference_days,
    'priority', config.priority,
    'operator', source_rule.operator,
    'updatedBy', config.updated_by,
    'updatedAt', config.updated_at
  ) order by config.priority, definition.rule_key), '[]'::jsonb)
  into v_rules
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
  where jsonb_array_length(definition.destination_source_types) = 1
    and source_rule.operator = case when config.rule_key in (
      'cgd_bank_statement_fdm_credit_card_monthly_income',
      'cgd_bank_statement_fdm_adyen_monthly_payments'
    ) then '-' else '+' end;
  if jsonb_array_length(v_rules) <> 7 then
    raise exception 'Automatic reconciliation managed catalog is invalid.';
  end if;

  select jsonb_build_object(
    'id', batch.id,
    'scheduledSlot', batch.scheduled_slot,
    'status', batch.status,
    'counts', batch.counts,
    'ruleCount', jsonb_array_length(batch.rule_snapshot),
    'childCount', coalesce((batch.counts->>'childCount')::integer, 0),
    'startedAt', batch.started_at,
    'finishedAt', batch.finished_at,
    'updatedAt', batch.updated_at
  )
  into v_last_scheduled_batch
  from public.financial_reconciliation_automatic_batches batch
  order by batch.scheduled_slot desc, batch.started_at desc, batch.id desc
  limit 1;

  return jsonb_build_object(
    'schedule', v_schedule,
    'rules', v_rules,
    'last_scheduled_batch', v_last_scheduled_batch
  );
end
$$;

create or replace function public.get_financial_reconciliation_automatic_run(
  p_run_id uuid
)
returns jsonb
language plpgsql
stable
security definer set search_path = public, pg_temp
as $$
declare
  v_run public.financial_reconciliation_automatic_runs%rowtype;
  v_proposals jsonb;
  v_rule_key text;
  v_rule_version integer;
  v_analysis_unit text;
begin
  select * into v_run
  from public.financial_reconciliation_automatic_runs
  where id = p_run_id;
  if not found then
    raise exception 'Automatic analysis run was not found.';
  end if;
  if jsonb_typeof(v_run.definition_config_snapshot) = 'array'
    and jsonb_array_length(v_run.definition_config_snapshot) = 1
    and coalesce(v_run.definition_config_snapshot#>>'{0,ruleVersion}', '')
      ~ '^[0-9]+$' then
    v_rule_key := v_run.definition_config_snapshot#>>'{0,ruleKey}';
    v_rule_version :=
      (v_run.definition_config_snapshot#>>'{0,ruleVersion}')::integer;
  end if;
  v_analysis_unit := case
    when v_rule_key = 'fdm_bank_transfer_cgd_bank_statement_combination'
      and v_rule_version = 1 then 'bank_anchors'
    when v_rule_key in (
      'cgd_bank_statement_fdm_credit_card_monthly_income',
      'cgd_bank_statement_fdm_adyen_monthly_payments'
    ) then 'calendar_months'
    else 'records'
  end;

  select coalesce(jsonb_agg(jsonb_build_object(
    'id', proposal.id,
    'runId', proposal.run_id,
    'ruleKey', proposal.rule_key,
    'ruleVersion', proposal.rule_version,
    'baseSourceType', proposal.base_source_type,
    'baseSourceId', proposal.base_source_id,
    'baseSourceDate', proposal.base_source_date,
    'baseSnapshot', case
      when (proposal.rule_key, proposal.rule_version) in (
        ('cgd_bank_statement_fdm_credit_card_monthly_income', 1),
        ('cgd_bank_statement_fdm_credit_card_monthly_income', 2),
        ('fdm_bank_transfer_cgd_bank_statement_combination', 1),
        ('cgd_bank_statement_fdm_adyen_monthly_payments', 1)
      ) then jsonb_build_object(
        'sourceType', proposal.base_source_type,
        'sourceId', proposal.base_source_id,
        'sourceDate', proposal.base_source_date
      )
      else proposal.base_snapshot
    end,
    'items', case
      when (proposal.rule_key, proposal.rule_version) in (
        ('cgd_bank_statement_fdm_credit_card_monthly_income', 1),
        ('cgd_bank_statement_fdm_credit_card_monthly_income', 2),
        ('fdm_bank_transfer_cgd_bank_statement_combination', 1),
        ('cgd_bank_statement_fdm_adyen_monthly_payments', 1)
      ) then '[]'::jsonb
      else proposal.items
    end,
    'evidence', case
      when (proposal.rule_key, proposal.rule_version) in (
        ('cgd_bank_statement_fdm_credit_card_monthly_income', 1),
        ('cgd_bank_statement_fdm_credit_card_monthly_income', 2),
        ('fdm_bank_transfer_cgd_bank_statement_combination', 1),
        ('cgd_bank_statement_fdm_adyen_monthly_payments', 1)
      ) then '[]'::jsonb
      else proposal.evidence
    end,
    'candidateGroups', case
      when (proposal.rule_key, proposal.rule_version) in (
        ('cgd_bank_statement_fdm_credit_card_monthly_income', 1),
        ('cgd_bank_statement_fdm_credit_card_monthly_income', 2),
        ('fdm_bank_transfer_cgd_bank_statement_combination', 1),
        ('cgd_bank_statement_fdm_adyen_monthly_payments', 1)
      ) then '[]'::jsonb
      else proposal.candidate_groups
    end,
    'groupingKey', proposal.grouping_key,
    'summarySnapshot', case
      when (proposal.rule_key, proposal.rule_version) =
        ('fdm_bank_transfer_cgd_bank_statement_combination', 1) then
        jsonb_build_object(
          'classification', proposal.summary_snapshot->'classification',
          'reason', proposal.summary_snapshot->'reason',
          'candidateCount', proposal.summary_snapshot->'candidateCount',
          'bankAnchorDate', coalesce(
            proposal.summary_snapshot->'bankAnchorDate',
            (select to_jsonb(member.source_date)
             from public.financial_reconciliation_automatic_proposal_memberships member
             where member.proposal_id = proposal.id
               and member.role = 'destination'
             order by member.ordinal
             limit 1)
          ),
          'sourceCount', coalesce(
            proposal.summary_snapshot->'sourceCount',
            (select to_jsonb(count(*)::integer)
             from public.financial_reconciliation_automatic_proposal_memberships member
             where member.proposal_id = proposal.id and member.role = 'source')
          ),
          'sourceTotal', coalesce(
            proposal.summary_snapshot->'sourceTotal',
            (select to_jsonb(coalesce(sum(member.amount), 0::numeric))
             from public.financial_reconciliation_automatic_proposal_memberships member
             where member.proposal_id = proposal.id and member.role = 'source')
          ),
          'destinationCount', coalesce(
            proposal.summary_snapshot->'destinationCount',
            (select to_jsonb(count(*)::integer)
             from public.financial_reconciliation_automatic_proposal_memberships member
             where member.proposal_id = proposal.id and member.role = 'destination')
          ),
          'destinationTotal', coalesce(
            proposal.summary_snapshot->'destinationTotal',
            (select to_jsonb(coalesce(sum(member.amount), 0::numeric))
             from public.financial_reconciliation_automatic_proposal_memberships member
             where member.proposal_id = proposal.id and member.role = 'destination')
          )
        )
      when (proposal.rule_key, proposal.rule_version) in (
        ('cgd_bank_statement_fdm_credit_card_monthly_income', 1),
        ('cgd_bank_statement_fdm_credit_card_monthly_income', 2),
        ('cgd_bank_statement_fdm_adyen_monthly_payments', 1)
      ) then jsonb_build_object(
        'calendarMonth', proposal.summary_snapshot->'calendarMonth',
        'sourceCount', proposal.summary_snapshot->'sourceCount',
        'sourceTotal', proposal.summary_snapshot->'sourceTotal',
        'destinationCount', proposal.summary_snapshot->'destinationCount',
        'destinationTotal', proposal.summary_snapshot->'destinationTotal'
      )
      else proposal.summary_snapshot
    end,
    'calculatedDifference', proposal.calculated_difference,
    'allowedDifference', proposal.allowed_difference,
    'status', proposal.status,
    'reason', proposal.reason,
    'signature', proposal.signature,
    'reconciliationId', proposal.reconciliation_id,
    'createdAt', proposal.created_at,
    'updatedAt', proposal.updated_at
  ) order by proposal.base_source_date, proposal.base_source_id,
             proposal.signature), '[]'::jsonb)
  into v_proposals
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.run_id = v_run.id;

  return jsonb_build_object(
    'runId', v_run.id,
    'trigger', v_run.trigger,
    'scope', v_run.scope,
    'status', v_run.status,
    'actor', v_run.actor,
    'clientRequestId', v_run.client_request_id,
    'scheduledSlot', v_run.scheduled_slot,
    'batchId', v_run.batch_id,
    'batchRuleKey', v_run.batch_rule_key,
    'batchRulePosition', v_run.batch_rule_position,
    'batchRuleCount', v_run.batch_rule_count,
    'definitions', v_run.definition_config_snapshot,
    'counts', v_run.counts,
    'analysisCursorDate', v_run.analysis_cursor_date,
    'analysisCursorId', v_run.analysis_cursor_id,
    'analysisProcessed', v_run.analysis_processed,
    'analysisTotal', v_run.analysis_total,
    'analysisErrorCode', v_run.analysis_error_code,
    'analysisErrorAt', v_run.analysis_error_at,
    'analysisUnit', v_analysis_unit,
    'analysisComplete', v_run.analysis_completed_at is not null,
    'analysisCompletedAt', v_run.analysis_completed_at,
    'startedAt', v_run.started_at,
    'finishedAt', v_run.finished_at,
    'proposals', v_proposals
  );
end
$$;

create or replace function public.financial_reconciliation_automatic_progress_or_run(
  p_run_id uuid
)
returns jsonb
language plpgsql
stable
security definer set search_path = public, pg_temp
as $$
declare
  v_run public.financial_reconciliation_automatic_runs%rowtype;
  v_rule_key text;
  v_rule_version integer;
  v_analysis_unit text;
begin
  select * into v_run
  from public.financial_reconciliation_automatic_runs
  where id = p_run_id;
  if not found then
    raise exception 'Automatic analysis run was not found.';
  end if;
  if v_run.analysis_completed_at is not null then
    return public.get_financial_reconciliation_automatic_run(p_run_id);
  end if;
  if jsonb_typeof(v_run.definition_config_snapshot) = 'array'
    and jsonb_array_length(v_run.definition_config_snapshot) = 1
    and coalesce(v_run.definition_config_snapshot#>>'{0,ruleVersion}', '')
      ~ '^[0-9]+$' then
    v_rule_key := v_run.definition_config_snapshot#>>'{0,ruleKey}';
    v_rule_version :=
      (v_run.definition_config_snapshot#>>'{0,ruleVersion}')::integer;
  end if;
  v_analysis_unit := case
    when v_rule_key = 'fdm_bank_transfer_cgd_bank_statement_combination'
      and v_rule_version = 1 then 'bank_anchors'
    when v_rule_key in (
      'cgd_bank_statement_fdm_credit_card_monthly_income',
      'cgd_bank_statement_fdm_adyen_monthly_payments'
    ) then 'calendar_months'
    else 'records'
  end;

  return jsonb_build_object(
    'runId', v_run.id,
    'trigger', v_run.trigger,
    'scope', v_run.scope,
    'status', v_run.status,
    'actor', v_run.actor,
    'clientRequestId', v_run.client_request_id,
    'scheduledSlot', v_run.scheduled_slot,
    'batchId', v_run.batch_id,
    'batchRuleKey', v_run.batch_rule_key,
    'batchRulePosition', v_run.batch_rule_position,
    'batchRuleCount', v_run.batch_rule_count,
    'definitions', v_run.definition_config_snapshot,
    'counts', v_run.counts,
    'analysisCursorDate', v_run.analysis_cursor_date,
    'analysisCursorId', v_run.analysis_cursor_id,
    'analysisProcessed', v_run.analysis_processed,
    'analysisTotal', v_run.analysis_total,
    'analysisErrorCode', v_run.analysis_error_code,
    'analysisErrorAt', v_run.analysis_error_at,
    'analysisUnit', v_analysis_unit,
    'analysisComplete', false,
    'analysisCompletedAt', null,
    'startedAt', v_run.started_at,
    'finishedAt', v_run.finished_at,
    'proposals', '[]'::jsonb
  );
end
$$;

create or replace function public.claim_financial_reconciliation_automatic_schedule(
  p_now timestamptz,
  p_actor text
)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
declare
  v_schedule public.financial_reconciliation_automatic_schedule%rowtype;
  v_batch public.financial_reconciliation_automatic_batches%rowtype;
  v_local timestamp;
  v_slot text;
  v_snapshot jsonb;
  v_enabled_rule_count integer;
  v_batch_rule_count integer;
  v_selected_rule jsonb;
  v_selected_position integer;
  v_selected_rule_key text;
  v_selected_rule_version integer;
  v_analysis_total bigint;
  v_run_id uuid;
begin
  if p_now is null then
    raise exception 'Schedule time is required.';
  end if;
  if nullif(trim(coalesce(p_actor, '')), '') is null then
    raise exception 'Actor is required.';
  end if;

  lock table public.financial_reconciliation_source_rules
    in share row exclusive mode;
  lock table public.financial_reconciliation_automatic_rule_configs
    in share row exclusive mode;
  select * into strict v_schedule
  from public.financial_reconciliation_automatic_schedule
  where id = true
  for update;
  if not v_schedule.enabled then
    return jsonb_build_object('claimed', false, 'reason', 'schedule_disabled');
  end if;

  v_local := p_now at time zone v_schedule.time_zone;
  v_slot := to_char(v_local::date, 'YYYY-MM-DD');

  select * into v_batch
  from public.financial_reconciliation_automatic_batches batch
  where batch.status in ('pending', 'running')
  order by batch.scheduled_slot, batch.started_at, batch.id
  for update
  limit 1;

  if not found then
    if v_local::time < v_schedule.time_of_day then
      return jsonb_build_object(
        'claimed', false, 'reason', 'before_scheduled_time'
      );
    end if;

    select * into v_batch
    from public.financial_reconciliation_automatic_batches batch
    where batch.scheduled_slot = v_slot
    for update;

    if not found then
      select count(*)::integer into v_enabled_rule_count
      from public.financial_reconciliation_automatic_rule_configs config
      where config.enabled
        and config.include_in_scheduled_batch;
      if v_enabled_rule_count = 0 then
        return jsonb_build_object(
          'claimed', false, 'reason', 'no_enabled_rules'
        );
      end if;

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
      ) order by config.priority, config.rule_key), '[]'::jsonb)
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
      where config.enabled
        and config.include_in_scheduled_batch
        and config.difference_allowed >= 0
        and jsonb_array_length(definition.destination_source_types) = 1
        and (config.rule_key, config.rule_version) in (
          ('financial_documents_cgd_bank_statement', 2),
          ('financial_documents_cgd_credit_card', 1),
          ('financial_documents_cgd_bank_statement_amount_only', 1),
          ('financial_documents_cgd_credit_card_amount_only', 1),
          ('cgd_bank_statement_fdm_credit_card_monthly_income', 2),
          ('fdm_bank_transfer_cgd_bank_statement_combination', 1),
          ('cgd_bank_statement_fdm_adyen_monthly_payments', 1)
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
          when config.rule_key =
              'fdm_bank_transfer_cgd_bank_statement_combination'
            then '+'
          when config.rule_key =
              'cgd_bank_statement_fdm_adyen_monthly_payments'
            then '-'
          when config.rule_key =
              'cgd_bank_statement_fdm_credit_card_monthly_income'
            then '-'
          else '+'
        end;

      if jsonb_array_length(v_snapshot) <> v_enabled_rule_count then
        return jsonb_build_object(
          'claimed', false, 'reason', 'unsupported_rule_set'
        );
      end if;

      insert into public.financial_reconciliation_automatic_batches (
        scheduled_slot, actor, status, rule_snapshot
      ) values (
        v_slot, trim(p_actor), 'pending', v_snapshot
      )
      on conflict (scheduled_slot) do nothing
      returning * into v_batch;
      if not found then
        select * into strict v_batch
        from public.financial_reconciliation_automatic_batches batch
        where batch.scheduled_slot = v_slot
        for update;
      end if;
    end if;
  end if;

  perform public.financial_reconciliation_refresh_automatic_batch(v_batch.id);
  select * into strict v_batch
  from public.financial_reconciliation_automatic_batches batch
  where batch.id = v_batch.id
  for update;
  if v_batch.status in ('completed', 'partial', 'failed') then
    return jsonb_build_object(
      'claimed', false,
      'reason', 'batch_complete',
      'batchId', v_batch.id
    );
  end if;

  v_batch_rule_count := jsonb_array_length(v_batch.rule_snapshot);
  select run.id into v_run_id
  from public.financial_reconciliation_automatic_runs run
  where run.batch_id = v_batch.id
    and run.scope = 'rule'
    and run.status in ('analyzing', 'ready', 'running')
    and run.finished_at is null
  order by run.batch_rule_position, run.started_at, run.id
  limit 1;
  if found then
    return jsonb_build_object(
      'claimed', true,
      'resumed', true,
      'batchId', v_batch.id,
      'batchRulePosition', (
        select run.batch_rule_position
        from public.financial_reconciliation_automatic_runs run
        where run.id = v_run_id
      ),
      'batchRuleCount', v_batch_rule_count,
      'run',
        public.financial_reconciliation_automatic_progress_or_run(v_run_id)
    );
  end if;

  select snapshot.value, snapshot.ordinality::integer
  into v_selected_rule, v_selected_position
  from jsonb_array_elements(v_batch.rule_snapshot)
    with ordinality snapshot(value, ordinality)
  where not exists (
    select 1
    from public.financial_reconciliation_automatic_runs run
    where run.batch_id = v_batch.id
      and run.batch_rule_position = snapshot.ordinality::integer
  )
  order by snapshot.ordinality
  limit 1;

  if not found then
    perform public.financial_reconciliation_refresh_automatic_batch(v_batch.id);
    select * into strict v_batch
    from public.financial_reconciliation_automatic_batches batch
    where batch.id = v_batch.id
    for update;
    if v_batch.status in ('completed', 'partial', 'failed') then
      return jsonb_build_object(
        'claimed', false,
        'reason', 'batch_complete',
        'batchId', v_batch.id
      );
    end if;
    raise exception 'Automatic scheduled batch has no resumable rule.';
  end if;

  v_selected_rule_key := v_selected_rule->>'ruleKey';
  v_selected_rule_version := (v_selected_rule->>'ruleVersion')::integer;
  if v_selected_rule_key =
      'cgd_bank_statement_fdm_credit_card_monthly_income'
    and v_selected_rule_version = 2 then
    select public.financial_reconciliation_automatic_monthly_income_count()
    into v_analysis_total;
  elsif v_selected_rule_key =
      'fdm_bank_transfer_cgd_bank_statement_combination'
    and v_selected_rule_version = 1 then
    select public.financial_reconciliation_automatic_bank_reservation_count()
    into v_analysis_total;
  elsif v_selected_rule_key =
      'cgd_bank_statement_fdm_adyen_monthly_payments'
    and v_selected_rule_version = 1 then
    select public.financial_reconciliation_automatic_adyen_month_count()
    into v_analysis_total;
  else
    select public.financial_reconciliation_automatic_base_count(
      v_selected_rule_key, v_selected_rule_version
    ) into v_analysis_total;
  end if;

  insert into public.financial_reconciliation_automatic_runs (
    trigger, scope, actor, scheduled_slot, definition_config_snapshot,
    analysis_processed, analysis_total,
    batch_id, batch_rule_key, batch_rule_position, batch_rule_count
  ) values (
    'scheduled', 'rule', v_batch.actor, v_batch.scheduled_slot,
    jsonb_build_array(v_selected_rule), 0, v_analysis_total,
    v_batch.id, v_selected_rule_key, v_selected_position,
    v_batch_rule_count
  )
  on conflict do nothing
  returning id into v_run_id;
  if v_run_id is null then
    select run.id into strict v_run_id
    from public.financial_reconciliation_automatic_runs run
    where run.batch_id = v_batch.id
      and run.batch_rule_position = v_selected_position
    limit 1;
    return jsonb_build_object(
      'claimed', true,
      'resumed', true,
      'batchId', v_batch.id,
      'batchRulePosition', v_selected_position,
      'batchRuleCount', v_batch_rule_count,
      'run',
        public.financial_reconciliation_automatic_progress_or_run(v_run_id)
    );
  end if;

  return jsonb_build_object(
    'claimed', true,
    'resumed', false,
    'batchId', v_batch.id,
    'batchRulePosition', v_selected_position,
    'batchRuleCount', v_batch_rule_count,
    'run', public.financial_reconciliation_automatic_progress_or_run(v_run_id)
  );
end
$$;

create or replace function public.financial_reconciliation_refresh_automatic_batch(
  p_batch_id uuid
)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
declare
  v_batch public.financial_reconciliation_automatic_batches%rowtype;
  v_rule_count integer;
  v_child_count integer;
  v_completed_children integer;
  v_partial_children integer;
  v_failed_children integer;
  v_unfinished_children integer;
  v_counts jsonb;
  v_children jsonb;
  v_status text;
begin
  if p_batch_id is null then
    raise exception 'Automatic batch ID is required.';
  end if;
  select * into v_batch
  from public.financial_reconciliation_automatic_batches
  where id = p_batch_id
  for update;
  if not found then
    raise exception 'Automatic scheduled batch was not found.';
  end if;

  if jsonb_typeof(v_batch.rule_snapshot) is distinct from 'array' then
    raise exception 'Automatic scheduled batch snapshot is invalid.';
  end if;
  v_rule_count := jsonb_array_length(v_batch.rule_snapshot);
  if exists (
    select 1
    from public.financial_reconciliation_automatic_runs run
    where run.batch_id = v_batch.id
      and run.scope = 'batch'
  ) then
    return jsonb_build_object(
      'id', v_batch.id,
      'scheduledSlot', v_batch.scheduled_slot,
      'actor', v_batch.actor,
      'status', v_batch.status,
      'ruleCount', v_rule_count,
      'childCount', 0,
      'counts', v_batch.counts,
      'children', '[]'::jsonb,
      'startedAt', v_batch.started_at,
      'finishedAt', v_batch.finished_at,
      'updatedAt', v_batch.updated_at
    );
  end if;
  if v_rule_count not between 1 and 7
    or exists (
      select 1
      from jsonb_array_elements(v_batch.rule_snapshot)
        with ordinality snapshot(value, position)
      where jsonb_typeof(snapshot.value) is distinct from 'object'
        or (select count(*) from jsonb_object_keys(snapshot.value)) <> 9
        or not (snapshot.value ?& array[
          'ruleKey','ruleVersion','displayName','priority',
          'differenceAllowed','maxDifferenceDays','destinationSourceType',
          'definition','operator'
        ])
        or jsonb_typeof(snapshot.value->'ruleKey') is distinct from 'string'
        or jsonb_typeof(snapshot.value->'ruleVersion') is distinct from 'number'
        or jsonb_typeof(snapshot.value->'displayName') is distinct from 'string'
        or jsonb_typeof(snapshot.value->'priority') is distinct from 'number'
        or jsonb_typeof(snapshot.value->'differenceAllowed')
          is distinct from 'number'
        or jsonb_typeof(snapshot.value->'maxDifferenceDays')
          is distinct from 'number'
        or jsonb_typeof(snapshot.value->'destinationSourceType')
          is distinct from 'string'
        or jsonb_typeof(snapshot.value->'operator') is distinct from 'string'
        or coalesce(snapshot.value->>'ruleKey', '') = ''
        or coalesce(snapshot.value->>'displayName', '') = ''
        or coalesce(snapshot.value->>'ruleVersion', '') !~ '^[0-9]+$'
        or coalesce(snapshot.value->>'priority', '') !~ '^[1-9][0-9]*$'
        or coalesce(snapshot.value->>'differenceAllowed', '')
          !~ '^[0-9]+(\.[0-9]+)?$'
        or coalesce(snapshot.value->>'maxDifferenceDays', '')
          !~ '^[0-9]+$'
        or coalesce(snapshot.value->>'destinationSourceType', '') = ''
        or jsonb_typeof(snapshot.value->'definition') is distinct from 'object'
        or coalesce(snapshot.value->>'operator', '') not in ('+', '-')
        or coalesce(
          (
            snapshot.value->>'ruleKey',
            (snapshot.value->>'ruleVersion')::integer
          ) in (
            ('financial_documents_cgd_bank_statement', 2),
            ('financial_documents_cgd_credit_card', 1),
            ('financial_documents_cgd_bank_statement_amount_only', 1),
            ('financial_documents_cgd_credit_card_amount_only', 1),
            ('cgd_bank_statement_fdm_credit_card_monthly_income', 2),
            ('fdm_bank_transfer_cgd_bank_statement_combination', 1),
            ('cgd_bank_statement_fdm_adyen_monthly_payments', 1)
          ),
          false
        ) = false
        or (snapshot.value->>'maxDifferenceDays')::integer > 90
        or (
          snapshot.value->>'ruleKey' =
            'fdm_bank_transfer_cgd_bank_statement_combination'
          and (
            (snapshot.value->>'differenceAllowed')::numeric <> 0
            or snapshot.value->>'operator' is distinct from '+'
            or snapshot.value->>'destinationSourceType' is distinct from
              'import_cgd_extrato_ordem'
            or snapshot.value->'definition' is distinct from jsonb_build_object(
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
        )
        or (
          snapshot.value->>'ruleKey' =
            'cgd_bank_statement_fdm_adyen_monthly_payments'
          and (
            (snapshot.value->>'maxDifferenceDays')::integer <> 31
            or snapshot.value->>'operator' is distinct from '-'
            or snapshot.value->>'destinationSourceType' is distinct from
              'import_fdm_accounts'
            or snapshot.value->'definition' is distinct from jsonb_build_object(
              'strategy', 'closed_calendar_month',
              'bankDescriptionContains', 'Adyen',
              'fdmAccount', 'Adyen',
              'requiresBothSides', true,
              'monthMarkerDays', 31
            )
          )
        )
        or (
          snapshot.value->>'ruleKey' =
            'cgd_bank_statement_fdm_credit_card_monthly_income'
          and (
            (snapshot.value->>'maxDifferenceDays')::integer <> 31
            or snapshot.value->>'operator' is distinct from '-'
          )
        )
        or (
          snapshot.value->>'ruleKey' in (
            'financial_documents_cgd_bank_statement',
            'financial_documents_cgd_credit_card',
            'financial_documents_cgd_bank_statement_amount_only',
            'financial_documents_cgd_credit_card_amount_only'
          )
          and snapshot.value->>'operator' is distinct from '+'
        )
        or (
          snapshot.value->>'ruleKey' not in (
            'fdm_bank_transfer_cgd_bank_statement_combination',
            'cgd_bank_statement_fdm_adyen_monthly_payments'
          )
          and public.financial_reconciliation_automatic_rule_contract(
            snapshot.value->>'ruleKey',
            (snapshot.value->>'ruleVersion')::integer
          )->>'destinationSourceType' is distinct from
            snapshot.value->>'destinationSourceType'
        )
        or not exists (
          select 1
          from public.financial_reconciliation_automatic_rule_definitions definition
          where definition.rule_key = snapshot.value->>'ruleKey'
            and definition.version =
              (snapshot.value->>'ruleVersion')::integer
            and definition.display_name = snapshot.value->>'displayName'
            and definition.destination_source_types = jsonb_build_array(
              snapshot.value->>'destinationSourceType'
            )
            and definition.definition = snapshot.value->'definition'
        )
    )
    or exists (
      select 1
      from jsonb_array_elements(v_batch.rule_snapshot) snapshot(value)
      group by snapshot.value->>'ruleKey'
      having count(*) <> 1
    )
    or exists (
      select 1
      from jsonb_array_elements(v_batch.rule_snapshot) snapshot(value)
      group by (snapshot.value->>'priority')::integer
      having count(*) <> 1
    )
    or v_batch.rule_snapshot is distinct from (
      select jsonb_agg(snapshot.value order by
        (snapshot.value->>'priority')::integer,
        snapshot.value->>'ruleKey'
      )
      from jsonb_array_elements(v_batch.rule_snapshot) snapshot(value)
    ) then
    raise exception 'Automatic scheduled batch snapshot is invalid.';
  end if;

  if exists (
    select 1
    from public.financial_reconciliation_automatic_runs run
    left join lateral (
      select snapshot.value
      from jsonb_array_elements(v_batch.rule_snapshot)
        with ordinality snapshot(value, position)
      where snapshot.position::integer = run.batch_rule_position
    ) expected on true
    where run.batch_id = v_batch.id
      and (
        run.trigger is distinct from 'scheduled'
        or run.scope is distinct from 'rule'
        or run.actor is distinct from v_batch.actor
        or run.scheduled_slot is distinct from v_batch.scheduled_slot
        or run.batch_rule_position is null
        or run.batch_rule_position not between 1 and v_rule_count
        or run.batch_rule_count is distinct from v_rule_count
        or run.batch_rule_key is distinct from expected.value->>'ruleKey'
        or jsonb_typeof(run.definition_config_snapshot) is distinct from 'array'
        or jsonb_array_length(run.definition_config_snapshot) <> 1
        or run.definition_config_snapshot->0 is distinct from expected.value
        or (
          run.status in ('completed', 'partial', 'failed')
          and run.finished_at is null
        )
        or (
          run.status not in ('completed', 'partial', 'failed')
          and run.finished_at is not null
        )
      )
  ) or exists (
    select 1
    from (
      select
        run.batch_rule_position,
        row_number() over (order by run.batch_rule_position) as expected_position
      from public.financial_reconciliation_automatic_runs run
      where run.batch_id = v_batch.id
        and run.scope = 'rule'
    ) ordered
    where ordered.batch_rule_position is distinct from ordered.expected_position
  ) or (
    select count(*)
    from public.financial_reconciliation_automatic_runs run
    where run.batch_id = v_batch.id
      and run.scope = 'rule'
      and run.status not in ('completed', 'partial', 'failed')
  ) > 1 then
    raise exception 'Automatic scheduled batch child metadata is invalid.';
  end if;

  if exists (
    select 1
    from public.financial_reconciliation_automatic_runs run
    left join lateral (
      select snapshot.value
      from jsonb_array_elements(v_batch.rule_snapshot)
        with ordinality snapshot(value, position)
      where snapshot.position::integer = run.batch_rule_position
    ) expected on true
    where run.batch_id = v_batch.id
      and run.scope = 'rule'
      and (
        run.analysis_processed is null
        or run.analysis_total is null
        or run.analysis_processed < 0
        or run.analysis_total < 0
        or run.analysis_processed > run.analysis_total
        or ((run.analysis_error_code is null) <>
          (run.analysis_error_at is null))
        or (
          run.status = 'analyzing'
          and run.analysis_completed_at is not null
        )
        or (
          run.status in ('ready', 'running', 'completed', 'partial')
          and run.analysis_completed_at is null
        )
        or (
          run.status in ('completed', 'partial')
          and run.analysis_processed <> run.analysis_total
        )
        or (
          run.status = 'failed'
          and run.analysis_completed_at is null
          and (
            nullif(run.analysis_error_code, '') is null
            or run.analysis_error_at is null
          )
        )
        or (
          expected.value->>'ruleKey' in (
            'fdm_bank_transfer_cgd_bank_statement_combination',
            'cgd_bank_statement_fdm_adyen_monthly_payments'
          )
          and (
            (run.analysis_processed = 0 and (
              run.analysis_cursor_date is not null
              or run.analysis_cursor_id is not null
            ))
            or (run.analysis_processed > 0 and (
              run.analysis_cursor_date is null
              or run.analysis_cursor_id is null
            ))
            or (
              run.analysis_completed_at is not null
              and run.analysis_processed <> run.analysis_total
            )
            or (
              expected.value->>'ruleKey' =
                'cgd_bank_statement_fdm_adyen_monthly_payments'
              and run.analysis_cursor_date is not null
              and extract(day from run.analysis_cursor_date) <> 1
            )
          )
        )
        or jsonb_typeof(run.counts) is distinct from 'object'
        or exists (
          select 1
          from unnest(array[
            'bases','proposed','ambiguous','skipped','deselected',
            'completed','stale','failed'
          ]) as count_key(value)
          where coalesce(run.counts->>count_key.value, '0') !~ '^[0-9]+$'
        )
      )
  ) then
    raise exception 'Automatic scheduled batch progress is invalid.';
  end if;

  select
    count(*)::integer,
    count(*) filter (where run.status = 'completed')::integer,
    count(*) filter (where run.status = 'partial')::integer,
    count(*) filter (where run.status = 'failed')::integer,
    count(*) filter (
      where run.status not in ('completed', 'partial', 'failed')
    )::integer,
    jsonb_build_object(
      'ruleCount', v_rule_count,
      'childCount', count(*),
      'completedChildren', count(*) filter (where run.status = 'completed'),
      'partialChildren', count(*) filter (where run.status = 'partial'),
      'failedChildren', count(*) filter (where run.status = 'failed'),
      'unfinishedChildren', count(*) filter (
        where run.status not in ('completed', 'partial', 'failed')
      ),
      'bases', coalesce(sum((run.counts->>'bases')::bigint), 0),
      'proposed', coalesce(sum((run.counts->>'proposed')::bigint), 0),
      'ambiguous', coalesce(sum((run.counts->>'ambiguous')::bigint), 0),
      'skipped', coalesce(sum((run.counts->>'skipped')::bigint), 0),
      'deselected', coalesce(sum((run.counts->>'deselected')::bigint), 0),
      'completed', coalesce(sum((run.counts->>'completed')::bigint), 0),
      'stale', coalesce(sum((run.counts->>'stale')::bigint), 0),
      'failed', coalesce(sum((run.counts->>'failed')::bigint), 0)
    )
  into
    v_child_count,
    v_completed_children,
    v_partial_children,
    v_failed_children,
    v_unfinished_children,
    v_counts
  from public.financial_reconciliation_automatic_runs run
  where run.batch_id = v_batch.id
    and run.scope = 'rule';

  select coalesce(jsonb_agg(jsonb_build_object(
    'runId', run.id,
    'ruleKey', run.batch_rule_key,
    'position', run.batch_rule_position,
    'ruleCount', run.batch_rule_count,
    'status', run.status,
    'counts', run.counts,
    'analysisProcessed', run.analysis_processed,
    'analysisTotal', run.analysis_total,
    'analysisCompletedAt', run.analysis_completed_at,
    'finishedAt', run.finished_at
  ) order by run.batch_rule_position), '[]'::jsonb)
  into v_children
  from public.financial_reconciliation_automatic_runs run
  where run.batch_id = v_batch.id
    and run.scope = 'rule';

  v_status := case
    when v_child_count = 0 then 'pending'
    when v_child_count < v_rule_count or v_unfinished_children > 0
      then 'running'
    when v_failed_children = v_child_count then 'failed'
    when v_failed_children > 0 or v_partial_children > 0 then 'partial'
    else 'completed'
  end;

  update public.financial_reconciliation_automatic_batches batch
  set status = v_status,
      counts = v_counts,
      finished_at = case
        when v_status in ('completed', 'partial', 'failed')
          then coalesce(batch.finished_at, now())
        else null
      end,
      updated_at = now()
  where batch.id = v_batch.id
  returning * into v_batch;

  return jsonb_build_object(
    'id', v_batch.id,
    'scheduledSlot', v_batch.scheduled_slot,
    'actor', v_batch.actor,
    'status', v_batch.status,
    'ruleCount', v_rule_count,
    'childCount', v_child_count,
    'counts', v_batch.counts,
    'children', v_children,
    'startedAt', v_batch.started_at,
    'finishedAt', v_batch.finished_at,
    'updatedAt', v_batch.updated_at
  );
end
$$;

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
          ('cgd_bank_statement_fdm_adyen_monthly_payments', 1)
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
  into v_run_id, v_existing_snapshot, v_existing_rule_key, v_existing_rule_version,
       v_existing_finished_at
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
        ('cgd_bank_statement_fdm_adyen_monthly_payments', 1)
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
      ('fdm_bank_transfer_cgd_bank_statement_combination', 1),
      ('cgd_bank_statement_fdm_adyen_monthly_payments', 1)
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
    and v_snapshot->0->>'operator' is distinct from '+' then
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
    and v_rule_version = 1 then
    select public.financial_reconciliation_automatic_bank_reservation_count()
    into v_total;
  elsif v_rule_key =
      'cgd_bank_statement_fdm_adyen_monthly_payments'
    and v_rule_version = 1 then
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

create or replace function public.get_financial_reconciliation_automatic_proposal_members(
  p_run_id uuid,
  p_proposal_id uuid,
  p_role text,
  p_offset integer,
  p_limit integer,
  p_actor text
)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
declare
  v_run record;
  v_total integer;
  v_source_count integer;
  v_destination_count integer;
  v_source_total numeric;
  v_destination_total numeric;
  v_members jsonb;
begin
  if p_run_id is null then
    raise exception 'Automatic run ID is required.';
  end if;
  if p_proposal_id is null then
    raise exception 'Automatic proposal ID is required.';
  end if;
  if p_role is null or p_role not in ('source', 'destination') then
    raise exception 'Automatic proposal member role is invalid.';
  end if;
  if p_offset is null or p_offset < 0 then
    raise exception 'Automatic proposal member offset must be zero or greater.';
  end if;
  if p_limit is null or p_limit < 1 or p_limit > 50 then
    raise exception 'Automatic proposal member limit must be between 1 and 50.';
  end if;
  if p_actor is null or btrim(p_actor) = '' then
    raise exception 'Automatic proposal member actor is required.';
  end if;

  select
    run.trigger,
    run.actor,
    run.finished_at,
    proposal.rule_key,
    proposal.rule_version,
    proposal.grouping_key,
    proposal.summary_snapshot
  into v_run
  from public.financial_reconciliation_automatic_runs run
  join public.financial_reconciliation_automatic_proposals proposal
    on proposal.run_id = run.id
  where run.id = p_run_id
    and proposal.id = p_proposal_id;
  if not found then
    raise exception
      'Automatic grouped proposal was not found for the requested run.';
  end if;
  if v_run.trigger <> 'manual' then
    raise exception 'Automatic proposal members require a manual run.';
  end if;
  if (v_run.rule_key, v_run.rule_version) not in (
    ('cgd_bank_statement_fdm_credit_card_monthly_income', 1),
    ('cgd_bank_statement_fdm_credit_card_monthly_income', 2),
    ('fdm_bank_transfer_cgd_bank_statement_combination', 1),
    ('cgd_bank_statement_fdm_adyen_monthly_payments', 1)
  ) then
    raise exception 'Automation proposal does not use a grouped-member rule.';
  end if;
  if v_run.finished_at is null and v_run.actor <> p_actor then
    raise exception 'You do not have permission for this automation run.';
  end if;

  select
    count(*) filter (where membership.role = 'source')::integer,
    coalesce(sum(membership.amount) filter (
      where membership.role = 'source'
    ), 0::numeric),
    count(*) filter (where membership.role = 'destination')::integer,
    coalesce(sum(membership.amount) filter (
      where membership.role = 'destination'
    ), 0::numeric),
    count(*) filter (where membership.role = p_role)::integer
  into
    v_source_count,
    v_source_total,
    v_destination_count,
    v_destination_total,
    v_total
  from public.financial_reconciliation_automatic_proposal_memberships membership
  where membership.proposal_id = p_proposal_id;

  select jsonb_agg(jsonb_build_object(
    'role', page.role,
    'sourceType', page.source_type,
    'sourceId', page.source_id,
    'ordinal', page.ordinal,
    'sourceDate', page.source_date,
    'amount', page.amount,
    'description', page.description,
    'account', page.account,
    'rowSnapshot', page.row_snapshot
  ) order by page.ordinal)
  into v_members
  from (
    select membership.*
    from public.financial_reconciliation_automatic_proposal_memberships membership
    where membership.proposal_id = p_proposal_id
      and membership.role = p_role
    order by membership.ordinal
    offset p_offset
    limit p_limit
  ) page;

  return jsonb_build_object(
    'runId', p_run_id,
    'proposalId', p_proposal_id,
    'ruleKey', v_run.rule_key,
    'ruleVersion', v_run.rule_version,
    'groupingKey', v_run.grouping_key,
    'summarySnapshot', v_run.summary_snapshot,
    'sourceCount', v_source_count,
    'sourceTotal', v_source_total,
    'destinationCount', v_destination_count,
    'destinationTotal', v_destination_total,
    'role', p_role,
    'offset', p_offset,
    'limit', p_limit,
    'totalCount', v_total,
    'members', coalesce(v_members, '[]'::jsonb)
  );
end
$$;

revoke all on function public.get_financial_reconciliation_automatic_manual_rules()
  from public, anon, authenticated, service_role;
revoke all on function public.get_financial_reconciliation_automation_settings()
  from public, anon, authenticated, service_role;
revoke all on function public.get_financial_reconciliation_automatic_run(uuid)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_automatic_progress_or_run(uuid)
  from public, anon, authenticated, service_role;
revoke all on function public.create_financial_reconciliation_automatic_analysis(text[],text,text,uuid)
  from public, anon, authenticated, service_role;
revoke all on function public.get_financial_reconciliation_automatic_proposal_members(uuid,uuid,text,integer,integer,text)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_refresh_automatic_batch(uuid)
  from public, anon, authenticated, service_role;
revoke all on function public.claim_financial_reconciliation_automatic_schedule(timestamptz,text)
  from public, anon, authenticated, service_role;

grant execute on function public.get_financial_reconciliation_automatic_manual_rules()
  to service_role;
grant execute on function public.get_financial_reconciliation_automation_settings()
  to service_role;
grant execute on function public.get_financial_reconciliation_automatic_run(uuid)
  to service_role;
grant execute on function public.create_financial_reconciliation_automatic_analysis(text[],text,text,uuid)
  to service_role;
grant execute on function public.get_financial_reconciliation_automatic_proposal_members(uuid,uuid,text,integer,integer,text)
  to service_role;
grant execute on function public.claim_financial_reconciliation_automatic_schedule(timestamptz,text)
  to service_role;

notify pgrst, 'reload schema';
