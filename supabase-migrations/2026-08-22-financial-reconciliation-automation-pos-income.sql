-- Install the Card Payments - POS - Income catalog and immutable membership
-- foundation. Analysis, execution, and paging functions are appended later.

do $migration$
declare
  v_definition jsonb := jsonb_build_object(
    'matchingMode', 'monthly_aggregate',
    'sourceDescriptionPattern', '%POS VENDAS%',
    'destinationAccount', 'Credit Card',
    'calendarGrouping', 'closed_month',
    'fixedMaxDifferenceDays', 31,
    'eligibilityFloor', '2026-01-01',
    'requiresNonNullAmount', true
  );
  v_logic text := 'Every unlocked CGD Bank Statement POS VENDAS record is reconciled against every unlocked FDM Credit Card record in the same closed calendar month; the difference is Bank Statement total minus FDM Accounts total.';
begin
  insert into public.financial_reconciliation_automatic_rule_definitions (
    rule_key, version, display_name, base_source_type,
    destination_source_types, logic_description, definition
  ) values (
    'cgd_bank_statement_fdm_credit_card_monthly_income',
    1,
    'Card Payments - POS - Income',
    'import_cgd_extrato_ordem',
    '["import_fdm_accounts"]'::jsonb,
    v_logic,
    v_definition
  )
  on conflict (rule_key, version) do nothing;

  if (select count(*)
      from public.financial_reconciliation_automatic_rule_definitions definition
      where definition.rule_key =
          'cgd_bank_statement_fdm_credit_card_monthly_income'
        and definition.version = 1) <> 1
    or not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_definitions definition
      where definition.rule_key =
          'cgd_bank_statement_fdm_credit_card_monthly_income'
        and definition.version = 1
        and definition.display_name = 'Card Payments - POS - Income'
        and definition.base_source_type = 'import_cgd_extrato_ordem'
        and definition.destination_source_types =
          '["import_fdm_accounts"]'::jsonb
        and definition.logic_description = v_logic
        and definition.definition = v_definition
    ) then
    raise exception 'Installed POS income automatic reconciliation definition differs from the approved immutable v1 definition.';
  end if;
end
$migration$;

do $migration$
declare
  v_next_priority integer;
begin
  lock table public.financial_reconciliation_automatic_rule_configs
    in share row exclusive mode;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs config
    where config.rule_key =
      'cgd_bank_statement_fdm_credit_card_monthly_income'
  ) then
    select coalesce(max(config.priority), 0) + 1
    into v_next_priority
    from public.financial_reconciliation_automatic_rule_configs config;

    insert into public.financial_reconciliation_automatic_rule_configs (
      rule_key, rule_version, enabled, allow_manual_execution,
      include_in_scheduled_batch, difference_allowed,
      max_difference_days, priority
    ) values (
      'cgd_bank_statement_fdm_credit_card_monthly_income',
      1, false, false, false, 7500.00, 31, v_next_priority
    );
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs config
    where config.rule_key =
        'cgd_bank_statement_fdm_credit_card_monthly_income'
      and config.rule_version = 1
      and config.max_difference_days = 31
  ) then
    raise exception 'Installed POS income configuration requires rule version 1 and fixed 31-day display property.';
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
        'financial_reconciliation_rule_configs_pos_income_days_check'
  ) then
    alter table public.financial_reconciliation_automatic_rule_configs
      add constraint financial_reconciliation_rule_configs_pos_income_days_check
      check (
        rule_key <> 'cgd_bank_statement_fdm_credit_card_monthly_income'
        or max_difference_days = 31
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
  create temporary table pos_income_days_constraint_expected (
    rule_key text,
    max_difference_days integer,
    constraint pos_income_days_constraint_expected_check check (
      rule_key <> 'cgd_bank_statement_fdm_credit_card_monthly_income'
      or max_difference_days = 31
    )
  ) on commit drop;

  select constraint_row.contype,
         pg_get_constraintdef(constraint_row.oid, true)
  into strict v_constraint_type, v_expected_definition
  from pg_constraint constraint_row
  where constraint_row.conrelid =
      'pos_income_days_constraint_expected'::regclass
    and constraint_row.conname =
      'pos_income_days_constraint_expected_check';

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
      'financial_reconciliation_rule_configs_pos_income_days_check';

  drop table pos_income_days_constraint_expected;

  if v_constraint_type is distinct from 'c'
    or v_installed_definition is distinct from v_expected_definition then
    raise exception 'Installed POS income fixed-days constraint differs from the required definition.';
  end if;
end
$migration$;

alter table public.financial_reconciliation_automatic_rule_configs
  validate constraint financial_reconciliation_rule_configs_pos_income_days_check;

alter table public.financial_reconciliation_automatic_proposals
  add column if not exists grouping_key text,
  add column if not exists summary_snapshot jsonb not null default '{}'::jsonb;

do $migration$
begin
  if not exists (
      select 1
      from information_schema.columns column_row
      where column_row.table_schema = 'public'
        and column_row.table_name =
          'financial_reconciliation_automatic_proposals'
        and column_row.column_name = 'grouping_key'
        and column_row.data_type = 'text'
        and column_row.is_nullable = 'YES'
    ) or not exists (
      select 1
      from information_schema.columns column_row
      where column_row.table_schema = 'public'
        and column_row.table_name =
          'financial_reconciliation_automatic_proposals'
        and column_row.column_name = 'summary_snapshot'
        and column_row.data_type = 'jsonb'
        and column_row.is_nullable = 'NO'
        and column_row.column_default = '''{}''::jsonb'
  ) then
    raise exception 'Installed POS income proposal columns differ from the required schema.';
  end if;

  if not exists (
    select 1
    from pg_constraint constraint_row
    where constraint_row.conrelid =
        'public.financial_reconciliation_automatic_proposals'::regclass
      and constraint_row.conname =
        'financial_reconciliation_proposals_summary_snapshot_check'
  ) then
    alter table public.financial_reconciliation_automatic_proposals
      add constraint financial_reconciliation_proposals_summary_snapshot_check
      check (jsonb_typeof(summary_snapshot) = 'object') not valid;
  end if;
end
$migration$;

do $migration$
declare
  v_constraint_type "char";
  v_installed_definition text;
  v_expected_definition text;
begin
  create temporary table pos_income_summary_constraint_expected (
    summary_snapshot jsonb,
    constraint pos_income_summary_constraint_expected_check
      check (jsonb_typeof(summary_snapshot) = 'object')
  ) on commit drop;

  select constraint_row.contype,
         pg_get_constraintdef(constraint_row.oid, true)
  into strict v_constraint_type, v_expected_definition
  from pg_constraint constraint_row
  where constraint_row.conrelid =
      'pos_income_summary_constraint_expected'::regclass
    and constraint_row.conname =
      'pos_income_summary_constraint_expected_check';

  select constraint_row.contype,
         regexp_replace(
           pg_get_constraintdef(constraint_row.oid, true),
           '\s+NOT VALID$', ''
         )
  into v_constraint_type, v_installed_definition
  from pg_constraint constraint_row
  where constraint_row.conrelid =
      'public.financial_reconciliation_automatic_proposals'::regclass
    and constraint_row.conname =
      'financial_reconciliation_proposals_summary_snapshot_check';

  drop table pos_income_summary_constraint_expected;

  if v_constraint_type is distinct from 'c'
    or v_installed_definition is distinct from v_expected_definition then
    raise exception 'Installed POS income proposal summary constraint differs from the required definition.';
  end if;
end
$migration$;

alter table public.financial_reconciliation_automatic_proposals
  validate constraint financial_reconciliation_proposals_summary_snapshot_check;

create or replace function public.prevent_financial_reconciliation_monthly_snapshot_change()
returns trigger
language plpgsql
set search_path = public, pg_temp
as $$
begin
  if new.grouping_key is distinct from old.grouping_key
    or new.summary_snapshot is distinct from old.summary_snapshot then
    raise exception 'Automatic proposal monthly audit snapshot is immutable.';
  end if;
  return new;
end
$$;

revoke all on function public.prevent_financial_reconciliation_monthly_snapshot_change()
  from public, anon, authenticated, service_role;

do $migration$
declare
  v_grouping_key_attnum smallint;
  v_summary_snapshot_attnum smallint;
begin
  select attribute.attnum
  into strict v_grouping_key_attnum
  from pg_attribute attribute
  where attribute.attrelid =
      'public.financial_reconciliation_automatic_proposals'::regclass
    and attribute.attname = 'grouping_key'
    and not attribute.attisdropped;

  select attribute.attnum
  into strict v_summary_snapshot_attnum
  from pg_attribute attribute
  where attribute.attrelid =
      'public.financial_reconciliation_automatic_proposals'::regclass
    and attribute.attname = 'summary_snapshot'
    and not attribute.attisdropped;

  if not exists (
    select 1
    from pg_trigger trigger_row
    where trigger_row.tgrelid =
        'public.financial_reconciliation_automatic_proposals'::regclass
      and trigger_row.tgname =
        'financial_reconciliation_automatic_monthly_snapshot_immutable'
  ) then
    create trigger financial_reconciliation_automatic_monthly_snapshot_immutable
    before update of grouping_key, summary_snapshot
    on public.financial_reconciliation_automatic_proposals
    for each row
    execute function public.prevent_financial_reconciliation_monthly_snapshot_change();
  end if;

  if not exists (
    select 1
    from pg_trigger trigger_row
    where trigger_row.tgrelid =
        'public.financial_reconciliation_automatic_proposals'::regclass
      and trigger_row.tgname =
        'financial_reconciliation_automatic_monthly_snapshot_immutable'
      and not trigger_row.tgisinternal
      and trigger_row.tgenabled = 'O'
      and trigger_row.tgfoid =
        'public.prevent_financial_reconciliation_monthly_snapshot_change()'::regprocedure
      and trigger_row.tgattr = format(
        '%s %s', v_grouping_key_attnum, v_summary_snapshot_attnum
      )::int2vector
      and trigger_row.tgtype = (1 | 2 | 16)
      and trigger_row.tgnargs = 0
      and octet_length(trigger_row.tgargs) = 0
      and trigger_row.tgqual is null
      and trigger_row.tgconstraint = 0
      and trigger_row.tgconstrrelid = 0
      and trigger_row.tgconstrindid = 0
      and not trigger_row.tgdeferrable
      and not trigger_row.tginitdeferred
      and trigger_row.tgoldtable is null
      and trigger_row.tgnewtable is null
  ) then
    raise exception 'Installed POS income proposal monthly-snapshot trigger differs from the required definition.';
  end if;
end
$migration$;

create table if not exists public.financial_reconciliation_automatic_proposal_memberships (
  proposal_id uuid not null,
  role text not null,
  source_type text not null,
  source_id uuid not null,
  ordinal integer not null,
  source_date date not null,
  amount numeric(14,2) not null,
  description text not null default '',
  account text not null default '',
  row_snapshot jsonb not null,
  created_at timestamptz not null default now(),
  constraint fr_auto_proposal_memberships_proposal_fkey
    foreign key (proposal_id)
    references public.financial_reconciliation_automatic_proposals(id)
    on delete cascade,
  constraint fr_auto_proposal_memberships_role_check
    check (role in ('source','destination')),
  constraint fr_auto_proposal_memberships_source_type_check
    check (source_type in ('import_cgd_extrato_ordem','import_fdm_accounts')),
  constraint fr_auto_proposal_memberships_ordinal_check
    check (ordinal > 0),
  constraint fr_auto_proposal_memberships_snapshot_check
    check (jsonb_typeof(row_snapshot) = 'object'),
  constraint fr_auto_proposal_memberships_pkey
    primary key (proposal_id, role, source_type, source_id),
  constraint fr_auto_proposal_memberships_role_ordinal_key
    unique (proposal_id, role, ordinal),
  constraint fr_auto_proposal_memberships_source_key
    unique (proposal_id, source_type, source_id)
);

do $migration$
declare
  v_column_count integer;
  v_constraint record;
  v_actual_type "char";
  v_actual_definition text;
  v_expected_type "char";
  v_expected_definition text;
begin
  select count(*)
  into v_column_count
  from information_schema.columns column_row
  where column_row.table_schema = 'public'
    and column_row.table_name =
      'financial_reconciliation_automatic_proposal_memberships';

  if v_column_count <> 11
    or not exists (
      select 1 from information_schema.columns
      where table_schema = 'public'
        and table_name =
          'financial_reconciliation_automatic_proposal_memberships'
        and column_name = 'proposal_id' and data_type = 'uuid'
        and is_nullable = 'NO'
    ) or not exists (
      select 1 from information_schema.columns
      where table_schema = 'public'
        and table_name =
          'financial_reconciliation_automatic_proposal_memberships'
        and column_name = 'role' and data_type = 'text'
        and is_nullable = 'NO'
    ) or not exists (
      select 1 from information_schema.columns
      where table_schema = 'public'
        and table_name =
          'financial_reconciliation_automatic_proposal_memberships'
        and column_name = 'source_type' and data_type = 'text'
        and is_nullable = 'NO'
    ) or not exists (
      select 1 from information_schema.columns
      where table_schema = 'public'
        and table_name =
          'financial_reconciliation_automatic_proposal_memberships'
        and column_name = 'source_id' and data_type = 'uuid'
        and is_nullable = 'NO'
    ) or not exists (
      select 1 from information_schema.columns
      where table_schema = 'public'
        and table_name =
          'financial_reconciliation_automatic_proposal_memberships'
        and column_name = 'ordinal' and data_type = 'integer'
        and is_nullable = 'NO'
    ) or not exists (
      select 1 from information_schema.columns
      where table_schema = 'public'
        and table_name =
          'financial_reconciliation_automatic_proposal_memberships'
        and column_name = 'source_date' and data_type = 'date'
        and is_nullable = 'NO'
    ) or not exists (
      select 1 from information_schema.columns
      where table_schema = 'public'
        and table_name =
          'financial_reconciliation_automatic_proposal_memberships'
        and column_name = 'amount' and data_type = 'numeric'
        and numeric_precision = 14 and numeric_scale = 2
        and is_nullable = 'NO'
    ) or not exists (
      select 1 from information_schema.columns
      where table_schema = 'public'
        and table_name =
          'financial_reconciliation_automatic_proposal_memberships'
        and column_name = 'description' and data_type = 'text'
        and is_nullable = 'NO'
        and column_default = $default$''::text$default$
    ) or not exists (
      select 1 from information_schema.columns
      where table_schema = 'public'
        and table_name =
          'financial_reconciliation_automatic_proposal_memberships'
        and column_name = 'account' and data_type = 'text'
        and is_nullable = 'NO'
        and column_default = $default$''::text$default$
    ) or not exists (
      select 1 from information_schema.columns
      where table_schema = 'public'
        and table_name =
          'financial_reconciliation_automatic_proposal_memberships'
        and column_name = 'row_snapshot' and data_type = 'jsonb'
        and is_nullable = 'NO' and column_default is null
    ) or not exists (
      select 1 from information_schema.columns
      where table_schema = 'public'
        and table_name =
          'financial_reconciliation_automatic_proposal_memberships'
        and column_name = 'created_at'
        and data_type = 'timestamp with time zone'
        and is_nullable = 'NO' and column_default = 'now()'
  ) then
    raise exception 'Installed POS income membership columns differ from the required contract.';
  end if;

  create temporary table pos_income_memberships_expected (
    proposal_id uuid not null,
    role text not null,
    source_type text not null,
    source_id uuid not null,
    ordinal integer not null,
    source_date date not null,
    amount numeric(14,2) not null,
    description text not null default '',
    account text not null default '',
    row_snapshot jsonb not null,
    created_at timestamptz not null default now(),
    constraint pos_income_memberships_expected_proposal_fkey
      foreign key (proposal_id)
      references public.financial_reconciliation_automatic_proposals(id)
      on delete cascade,
    constraint pos_income_memberships_expected_role_check
      check (role in ('source','destination')),
    constraint pos_income_memberships_expected_source_type_check
      check (source_type in ('import_cgd_extrato_ordem','import_fdm_accounts')),
    constraint pos_income_memberships_expected_ordinal_check
      check (ordinal > 0),
    constraint pos_income_memberships_expected_snapshot_check
      check (jsonb_typeof(row_snapshot) = 'object'),
    constraint pos_income_memberships_expected_pkey
      primary key (proposal_id, role, source_type, source_id),
    constraint pos_income_memberships_expected_role_ordinal_key
      unique (proposal_id, role, ordinal),
    constraint pos_income_memberships_expected_source_key
      unique (proposal_id, source_type, source_id)
  ) on commit drop;

  for v_constraint in
    select * from (values
      ('fr_auto_proposal_memberships_proposal_fkey',
       'pos_income_memberships_expected_proposal_fkey'),
      ('fr_auto_proposal_memberships_role_check',
       'pos_income_memberships_expected_role_check'),
      ('fr_auto_proposal_memberships_source_type_check',
       'pos_income_memberships_expected_source_type_check'),
      ('fr_auto_proposal_memberships_ordinal_check',
       'pos_income_memberships_expected_ordinal_check'),
      ('fr_auto_proposal_memberships_snapshot_check',
       'pos_income_memberships_expected_snapshot_check'),
      ('fr_auto_proposal_memberships_pkey',
       'pos_income_memberships_expected_pkey'),
      ('fr_auto_proposal_memberships_role_ordinal_key',
       'pos_income_memberships_expected_role_ordinal_key'),
      ('fr_auto_proposal_memberships_source_key',
       'pos_income_memberships_expected_source_key')
    ) expected(actual_name, expected_name)
  loop
    select constraint_row.contype,
           pg_get_constraintdef(constraint_row.oid, true)
    into v_actual_type, v_actual_definition
    from pg_constraint constraint_row
    where constraint_row.conrelid =
        'public.financial_reconciliation_automatic_proposal_memberships'::regclass
      and constraint_row.conname = v_constraint.actual_name;

    select constraint_row.contype,
           pg_get_constraintdef(constraint_row.oid, true)
    into strict v_expected_type, v_expected_definition
    from pg_constraint constraint_row
    where constraint_row.conrelid =
        'pos_income_memberships_expected'::regclass
      and constraint_row.conname = v_constraint.expected_name;

    if v_actual_type is distinct from v_expected_type
      or v_actual_definition is distinct from v_expected_definition then
      raise exception 'Installed POS income membership constraint % differs from the required definition.',
        v_constraint.actual_name;
    end if;
  end loop;

  drop table pos_income_memberships_expected;
end
$migration$;

create index if not exists financial_reconciliation_automatic_memberships_role_ordinal_idx
  on public.financial_reconciliation_automatic_proposal_memberships
  (proposal_id, role, ordinal);

create index if not exists import_cgd_extrato_ordem_pos_income_lock_idx
  on public.import_cgd_extrato_ordem (data, id)
  where descritivo ilike '%POS VENDAS%';

create index if not exists import_fdm_accounts_credit_card_lock_idx
  on public.import_fdm_accounts (event_date, id)
  where account = 'Credit Card';

do $migration$
declare
  v_index record;
  v_actual record;
  v_expected record;
begin
  create temporary table pos_income_membership_index_expected (
    proposal_id uuid,
    role text,
    ordinal integer
  ) on commit drop;
  create index pos_income_membership_index_expected_idx
    on pos_income_membership_index_expected (proposal_id, role, ordinal);

  create temporary table pos_income_bank_index_expected (
    data date,
    id uuid,
    descritivo text
  ) on commit drop;
  create index pos_income_bank_index_expected_idx
    on pos_income_bank_index_expected (data, id)
    where descritivo ilike '%POS VENDAS%';

  create temporary table pos_income_fdm_index_expected (
    event_date date,
    id uuid,
    account text
  ) on commit drop;
  create index pos_income_fdm_index_expected_idx
    on pos_income_fdm_index_expected (event_date, id)
    where account = 'Credit Card';

  for v_index in
    select * from (values
      ('financial_reconciliation_automatic_memberships_role_ordinal_idx',
       'public.financial_reconciliation_automatic_proposal_memberships',
       'pos_income_membership_index_expected_idx'),
      ('import_cgd_extrato_ordem_pos_income_lock_idx',
       'public.import_cgd_extrato_ordem',
       'pos_income_bank_index_expected_idx'),
      ('import_fdm_accounts_credit_card_lock_idx',
       'public.import_fdm_accounts',
       'pos_income_fdm_index_expected_idx')
    ) expected(index_name, table_name, expected_index_name)
  loop
    select
      index_row.indrelid as table_oid,
      index_row.indisunique,
      index_row.indisprimary,
      index_row.indisexclusion,
      index_row.indimmediate,
      index_row.indisclustered,
      index_row.indisreplident,
      index_row.indisvalid,
      index_row.indisready,
      index_row.indislive,
      index_row.indnkeyatts,
      index_row.indnatts,
      regexp_replace(
        pg_get_indexdef(index_row.indexrelid),
        '^CREATE (UNIQUE )?INDEX [^ ]+ ON (ONLY )?[^ ]+',
        'CREATE \1INDEX ON \2<table>'
      ) as normalized_definition
    into v_actual
    from pg_index index_row
    where index_row.indexrelid =
      to_regclass('public.' || v_index.index_name);

    select
      index_row.indisunique,
      index_row.indisprimary,
      index_row.indisexclusion,
      index_row.indimmediate,
      index_row.indisclustered,
      index_row.indisreplident,
      index_row.indisvalid,
      index_row.indisready,
      index_row.indislive,
      index_row.indnkeyatts,
      index_row.indnatts,
      regexp_replace(
        pg_get_indexdef(index_row.indexrelid),
        '^CREATE (UNIQUE )?INDEX [^ ]+ ON (ONLY )?[^ ]+',
        'CREATE \1INDEX ON \2<table>'
      ) as normalized_definition
    into strict v_expected
    from pg_index index_row
    where index_row.indexrelid =
      to_regclass(v_index.expected_index_name);

    if v_actual.table_oid is distinct from to_regclass(v_index.table_name)
      or v_actual.indisunique is distinct from v_expected.indisunique
      or v_actual.indisprimary is distinct from v_expected.indisprimary
      or v_actual.indisexclusion is distinct from v_expected.indisexclusion
      or v_actual.indimmediate is distinct from v_expected.indimmediate
      or v_actual.indisclustered is distinct from v_expected.indisclustered
      or v_actual.indisreplident is distinct from v_expected.indisreplident
      or v_actual.indisvalid is distinct from v_expected.indisvalid
      or v_actual.indisready is distinct from v_expected.indisready
      or v_actual.indislive is distinct from v_expected.indislive
      or v_actual.indnkeyatts is distinct from v_expected.indnkeyatts
      or v_actual.indnatts is distinct from v_expected.indnatts
      or v_actual.normalized_definition is distinct from
        v_expected.normalized_definition then
      raise exception 'Installed POS income index % differs from the required definition.',
        v_index.index_name;
    end if;
  end loop;

  drop table pos_income_membership_index_expected;
  drop table pos_income_bank_index_expected;
  drop table pos_income_fdm_index_expected;
end
$migration$;

create or replace function public.prevent_financial_reconciliation_automatic_membership_change()
returns trigger
language plpgsql
set search_path = public, pg_temp
as $$
begin
  raise exception 'Automatic proposal memberships are immutable.';
end
$$;

revoke all on function public.prevent_financial_reconciliation_automatic_membership_change()
  from public, anon, authenticated, service_role;

do $migration$
begin
  if not exists (
    select 1
    from pg_trigger trigger_row
    where trigger_row.tgrelid =
        'public.financial_reconciliation_automatic_proposal_memberships'::regclass
      and trigger_row.tgname =
        'financial_reconciliation_automatic_membership_immutable'
  ) then
    create trigger financial_reconciliation_automatic_membership_immutable
    before update
    on public.financial_reconciliation_automatic_proposal_memberships
    for each row
    execute function public.prevent_financial_reconciliation_automatic_membership_change();
  end if;

  if not exists (
    select 1
    from pg_trigger trigger_row
    where trigger_row.tgrelid =
        'public.financial_reconciliation_automatic_proposal_memberships'::regclass
      and trigger_row.tgname =
        'financial_reconciliation_automatic_membership_immutable'
      and not trigger_row.tgisinternal
      and trigger_row.tgenabled = 'O'
      and trigger_row.tgfoid =
        'public.prevent_financial_reconciliation_automatic_membership_change()'::regprocedure
      and (trigger_row.tgtype & 1) = 1
      and (trigger_row.tgtype & 2) = 2
      and (trigger_row.tgtype & 16) = 16
      and (trigger_row.tgtype & (4 | 8 | 32)) = 0
  ) then
    raise exception 'Installed POS income membership immutability trigger differs from the required definition.';
  end if;
end
$migration$;

alter table public.financial_reconciliation_automatic_proposal_memberships
  enable row level security;

revoke all on table public.financial_reconciliation_automatic_proposal_memberships
  from public, anon, authenticated, service_role;

insert into public.financial_reconciliation_source_rules (
  base_source_type, matching_source_type, operator
) values (
  'import_cgd_extrato_ordem', 'import_fdm_accounts', '-'
)
on conflict (base_source_type, matching_source_type) do nothing;

do $migration$
begin
  if not exists (
    select 1
    from public.financial_reconciliation_source_rules source_rule
    where source_rule.base_source_type = 'import_cgd_extrato_ordem'
      and source_rule.matching_source_type = 'import_fdm_accounts'
      and source_rule.operator = '-'
  ) then
    raise exception 'The managed POS income source rule must remain enabled with operator -.';
  end if;
end
$migration$;

create or replace function public.replace_financial_reconciliation_source_rules(p_rules jsonb)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
begin
  if p_rules is null or jsonb_typeof(p_rules) <> 'array' then
    raise exception 'Reconciliation rules must be an array.';
  end if;
  if exists (
    select 1 from jsonb_array_elements(p_rules) rule
    where jsonb_typeof(rule) <> 'object'
  ) then
    raise exception 'Each reconciliation rule must be an object.';
  end if;
  if exists (
    select 1 from jsonb_array_elements(p_rules) rule
    where coalesce(rule->>'base_source_type', '') not in (
      'financial_documents', 'import_fdm_accounts',
      'import_cgd_cartao_credito', 'import_cgd_extrato_ordem'
    ) or coalesce(rule->>'matching_source_type', '') not in (
      'financial_documents', 'import_fdm_accounts',
      'import_cgd_cartao_credito', 'import_cgd_extrato_ordem'
    )
  ) then
    raise exception 'Rule source type is invalid.';
  end if;
  if exists (
    select 1 from jsonb_array_elements(p_rules) rule
    where rule->>'base_source_type' = rule->>'matching_source_type'
  ) then
    raise exception 'Rule sources must be different.';
  end if;
  if exists (
    select 1 from jsonb_array_elements(p_rules) rule
    where coalesce(rule->>'operator', '') not in ('+', '-')
  ) then
    raise exception 'Rule operator must be ''+'' or ''-''.';
  end if;
  if exists (
    select 1
    from jsonb_to_recordset(p_rules) as rule(
      base_source_type text, matching_source_type text, operator text
    )
    group by rule.base_source_type, rule.matching_source_type
    having count(*) > 1
  ) then
    raise exception 'Duplicate reconciliation rule.';
  end if;
  if (
    select count(*)
    from jsonb_to_recordset(p_rules) as rule(
      base_source_type text, matching_source_type text, operator text
    )
    where rule.base_source_type = 'financial_documents'
      and rule.matching_source_type = 'import_cgd_cartao_credito'
      and rule.operator = '+'
  ) <> 1 then
    raise exception 'The managed Credit Card source rule must remain enabled with operator +.';
  end if;
  if (
    select count(*)
    from jsonb_to_recordset(p_rules) as rule(
      base_source_type text, matching_source_type text, operator text
    )
    where rule.base_source_type = 'financial_documents'
      and rule.matching_source_type = 'import_cgd_extrato_ordem'
      and rule.operator = '+'
  ) <> 1 then
    raise exception 'The managed Bank Statement source rule must remain enabled with operator +.';
  end if;
  if (
    select count(*)
    from jsonb_to_recordset(p_rules) as rule(
      base_source_type text, matching_source_type text, operator text
    )
    where rule.base_source_type = 'import_cgd_extrato_ordem'
      and rule.matching_source_type = 'import_fdm_accounts'
      and rule.operator = '-'
  ) <> 1 then
    raise exception 'The managed POS income source rule must remain enabled with operator -.';
  end if;

  lock table public.financial_reconciliation_source_rules
    in share row exclusive mode;
  delete from public.financial_reconciliation_source_rules
  where base_source_type in (
    'financial_documents', 'import_fdm_accounts',
    'import_cgd_cartao_credito', 'import_cgd_extrato_ordem'
  );

  insert into public.financial_reconciliation_source_rules (
    base_source_type, matching_source_type, operator
  )
  select rule.base_source_type, rule.matching_source_type, rule.operator
  from jsonb_to_recordset(p_rules) as rule(
    base_source_type text, matching_source_type text, operator text
  );

  return coalesce((
    select jsonb_agg(jsonb_build_object(
      'base_source_type', base_source_type,
      'matching_source_type', matching_source_type,
      'operator', operator
    ) order by base_source_type, matching_source_type)
    from public.financial_reconciliation_source_rules
  ), '[]'::jsonb);
end
$$;

revoke all on function public.replace_financial_reconciliation_source_rules(jsonb)
  from public, anon, authenticated;
grant execute on function public.replace_financial_reconciliation_source_rules(jsonb)
  to service_role;

create or replace function public.financial_reconciliation_automatic_rule_contract(
  p_rule_key text,
  p_rule_version integer
)
returns jsonb
language sql
immutable
security definer set search_path = public, pg_temp
as $$
  select case
    when p_rule_key = 'financial_documents_cgd_bank_statement' and p_rule_version = 2 then
      jsonb_build_object('payment','Banco','destinationSourceType','import_cgd_extrato_ordem',
        'descriptionThreshold',0.60,'supplierThreshold',0.70,'maxDestinationRecords',4,'maxCandidates',12)
    when p_rule_key = 'financial_documents_cgd_credit_card' and p_rule_version = 1 then
      jsonb_build_object('payment','Visa','destinationSourceType','import_cgd_cartao_credito',
        'descriptionThreshold',0.55,'supplierThreshold',0.60,'maxDestinationRecords',4,'maxCandidates',12)
    when p_rule_key = 'financial_documents_cgd_bank_statement_amount_only' and p_rule_version = 1 then
      jsonb_build_object('payment','Banco','destinationSourceType','import_cgd_extrato_ordem',
        'matchingMode','amount_only_one_to_one','maxDestinationRecords',1,'maxCandidates',12,
        'fixedDifferenceAllowed',0)
    when p_rule_key = 'financial_documents_cgd_credit_card_amount_only' and p_rule_version = 1 then
      jsonb_build_object('payment','Visa','destinationSourceType','import_cgd_cartao_credito',
        'matchingMode','amount_only_one_to_one','maxDestinationRecords',1,'maxCandidates',12,
        'fixedDifferenceAllowed',0)
    when p_rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income' and p_rule_version = 1 then
      jsonb_build_object(
        'destinationSourceType','import_fdm_accounts',
        'matchingMode','monthly_aggregate',
        'sourceDescriptionPattern','%POS VENDAS%',
        'destinationAccount','Credit Card',
        'calendarGrouping','closed_month',
        'eligibilityFloor','2026-01-01',
        'fixedMaxDifferenceDays',31,
        'requiresNonNullAmount',true
      )
    else null
  end
$$;

create or replace function public.financial_reconciliation_automatic_monthly_income_count()
returns bigint
language sql
stable
security definer set search_path = public, pg_temp
as $$
  with source_rows as materialized (
    select date_trunc('month', bank.data)::date as calendar_month
    from public.import_cgd_extrato_ordem bank
    where bank.data >= date '2026-01-01'
      and bank.data < date_trunc('month', current_date)::date
      and bank.montante is not null
      and bank.descritivo ilike '%POS VENDAS%'
      and not exists (
        select 1
        from public.financial_reconciliation_items locked
        where locked.source_type = 'import_cgd_extrato_ordem'
          and locked.source_id = bank.id
      )
  ), destination_rows as materialized (
    select date_trunc('month', fdm.event_date)::date as calendar_month
    from public.import_fdm_accounts fdm
    where fdm.event_date >= date '2026-01-01'
      and fdm.event_date < date_trunc('month', current_date)::date
      and fdm.amount is not null
      and fdm.account = 'Credit Card'
      and not exists (
        select 1
        from public.financial_reconciliation_items locked
        where locked.source_type = 'import_fdm_accounts'
          and locked.source_id = fdm.id
      )
  ), source_months as (
    select source_rows.calendar_month
    from source_rows
    group by source_rows.calendar_month
  ), destination_months as (
    select destination_rows.calendar_month
    from destination_rows
    group by destination_rows.calendar_month
  )
  select count(*)
  from source_months source
  join destination_months destination using (calendar_month)
$$;

create or replace function public.financial_reconciliation_automatic_monthly_income_page(
  p_after_month date,
  p_limit integer
)
returns table (
  calendar_month date,
  source_count integer,
  source_total numeric(14,2),
  destination_count integer,
  destination_total numeric(14,2),
  calculated_difference numeric(14,2),
  technical_base_source_id uuid,
  technical_base_source_date date
)
language plpgsql
stable
security definer set search_path = public, pg_temp
as $$
begin
  if p_limit is null or p_limit not between 1 and 25 then
    raise exception 'Automatic monthly analysis page size must be between 1 and 25.';
  end if;

  return query
  with source_rows as materialized (
    select bank.id, bank.data, bank.montante, bank.descritivo,
           date_trunc('month', bank.data)::date as calendar_month
    from public.import_cgd_extrato_ordem bank
    where bank.data >= date '2026-01-01'
      and bank.data < date_trunc('month', current_date)::date
      and bank.montante is not null
      and bank.descritivo ilike '%POS VENDAS%'
      and not exists (
        select 1
        from public.financial_reconciliation_items locked
        where locked.source_type = 'import_cgd_extrato_ordem'
          and locked.source_id = bank.id
      )
  ), destination_rows as materialized (
    select fdm.id, fdm.event_date, fdm.amount, fdm.description, fdm.account,
           date_trunc('month', fdm.event_date)::date as calendar_month
    from public.import_fdm_accounts fdm
    where fdm.event_date >= date '2026-01-01'
      and fdm.event_date < date_trunc('month', current_date)::date
      and fdm.amount is not null
      and fdm.account = 'Credit Card'
      and not exists (
        select 1
        from public.financial_reconciliation_items locked
        where locked.source_type = 'import_fdm_accounts'
          and locked.source_id = fdm.id
      )
  ), source_months as (
    select
      source_rows.calendar_month,
      count(*)::integer as source_count,
      sum(source_rows.montante)::numeric(14,2) as source_total,
      (array_agg(source_rows.id order by source_rows.data, source_rows.id))[1]
        as technical_base_source_id,
      min(source_rows.data) as technical_base_source_date
    from source_rows
    group by source_rows.calendar_month
  ), destination_months as (
    select
      destination_rows.calendar_month,
      count(*)::integer as destination_count,
      sum(destination_rows.amount)::numeric(14,2) as destination_total
    from destination_rows
    group by destination_rows.calendar_month
  )
  select
    source.calendar_month,
    source.source_count,
    source.source_total,
    destination.destination_count,
    destination.destination_total,
    (source.source_total - destination.destination_total)::numeric(14,2)
      as calculated_difference,
    source.technical_base_source_id,
    source.technical_base_source_date
  from source_months source
  join destination_months destination using (calendar_month)
  where source.calendar_month > coalesce(p_after_month, date '0001-01-01')
  order by source.calendar_month
  limit p_limit;
end
$$;

create or replace function public.financial_reconciliation_continue_automatic_monthly_income(
  p_run_id uuid,
  p_rule jsonb
)
returns jsonb
language plpgsql
set search_path = public, pg_temp
as $$
declare
  v_month record;
  v_page_count integer := 0;
  v_total bigint;
  v_last_month date;
  v_last_id uuid;
  v_difference_allowed numeric(14,2);
  v_operator text;
  v_source_ids uuid[];
  v_destination_ids uuid[];
  v_inserted_source_ids uuid[];
  v_inserted_destination_ids uuid[];
  v_inserted_source_total numeric(14,2);
  v_inserted_destination_total numeric(14,2);
  v_inserted_base_id uuid;
  v_inserted_base_date date;
  v_base_snapshot jsonb;
  v_summary_snapshot jsonb;
  v_signature text;
  v_status text;
  v_reason text;
  v_proposal_id uuid;
  v_inserted boolean;
  v_inserted_count integer;
begin
  if p_rule->>'ruleKey' is distinct from
      'cgd_bank_statement_fdm_credit_card_monthly_income'
    or p_rule->>'ruleVersion' is distinct from '1'
    or p_rule->>'destinationSourceType' is distinct from 'import_fdm_accounts'
    or p_rule->>'operator' is distinct from '-'
    or p_rule->>'maxDifferenceDays' is distinct from '31'
    or coalesce(p_rule->>'differenceAllowed', '') !~ '^[0-9]+(\.[0-9]+)?$'
    or p_rule->'definition' is distinct from jsonb_build_object(
      'matchingMode', 'monthly_aggregate',
      'sourceDescriptionPattern', '%POS VENDAS%',
      'destinationAccount', 'Credit Card',
      'calendarGrouping', 'closed_month',
      'fixedMaxDifferenceDays', 31,
      'eligibilityFloor', '2026-01-01',
      'requiresNonNullAmount', true
    ) then
    raise exception 'Automatic monthly rule snapshot contract is invalid.';
  end if;

  v_difference_allowed := (p_rule->>'differenceAllowed')::numeric(14,2);
  v_operator := p_rule->>'operator';

  select public.financial_reconciliation_automatic_monthly_income_count()
  into v_total;
  update public.financial_reconciliation_automatic_runs
  set analysis_total = greatest(analysis_total, analysis_processed, v_total),
      updated_at = now()
  where id = p_run_id;

  for v_month in
    select *
    from public.financial_reconciliation_automatic_monthly_income_page(
      (
        select run.analysis_cursor_date
        from public.financial_reconciliation_automatic_runs run
        where run.id = p_run_id
      ),
      25
    )
  loop
    v_page_count := v_page_count + 1;
    v_last_month := v_month.calendar_month;
    v_last_id := v_month.technical_base_source_id;

    begin
      select array_agg(bank.id order by bank.data, bank.id)
      into v_source_ids
      from public.import_cgd_extrato_ordem bank
      where bank.data >= v_month.calendar_month
        and bank.data < v_month.calendar_month + interval '1 month'
        and bank.data >= date '2026-01-01'
        and bank.data < date_trunc('month', current_date)::date
        and bank.montante is not null
        and bank.descritivo ilike '%POS VENDAS%'
        and not exists (
          select 1
          from public.financial_reconciliation_items locked
          where locked.source_type = 'import_cgd_extrato_ordem'
            and locked.source_id = bank.id
        );

      select array_agg(fdm.id order by fdm.event_date, fdm.id)
      into v_destination_ids
      from public.import_fdm_accounts fdm
      where fdm.event_date >= v_month.calendar_month
        and fdm.event_date < v_month.calendar_month + interval '1 month'
        and fdm.event_date >= date '2026-01-01'
        and fdm.event_date < date_trunc('month', current_date)::date
        and fdm.amount is not null
        and fdm.account = 'Credit Card'
        and not exists (
          select 1
          from public.financial_reconciliation_items locked
          where locked.source_type = 'import_fdm_accounts'
            and locked.source_id = fdm.id
        );

      if coalesce(cardinality(v_source_ids), 0) <> v_month.source_count
        or coalesce(cardinality(v_destination_ids), 0) <>
          v_month.destination_count then
        raise exception 'Automatic monthly membership changed during analysis.';
      end if;

      select jsonb_build_object(
        'sourceType', 'import_cgd_extrato_ordem',
        'sourceId', bank.id,
        'sourceDate', bank.data,
        'amount', bank.montante,
        'description', bank.descritivo
      )
      into strict v_base_snapshot
      from public.import_cgd_extrato_ordem bank
      where bank.id = v_month.technical_base_source_id
        and bank.id = any(v_source_ids)
        and bank.montante is not null;

      v_status := case
        when abs(v_month.calculated_difference) <= v_difference_allowed
          then 'proposed'
        else 'ambiguous'
      end;
      v_reason := case
        when v_status = 'ambiguous' then 'monthly_difference_exceeded'
        else ''
      end;
      v_signature := public.financial_reconciliation_extension_sha256(
        jsonb_build_object(
          'ruleKey', 'cgd_bank_statement_fdm_credit_card_monthly_income',
          'ruleVersion', 1,
          'calendarMonth', v_month.calendar_month,
          'sourceIds', to_jsonb(v_source_ids),
          'destinationIds', to_jsonb(v_destination_ids),
          'sourceTotal', v_month.source_total,
          'destinationTotal', v_month.destination_total,
          'calculatedDifference', v_month.calculated_difference,
          'differenceAllowed', v_difference_allowed,
          'operator', v_operator
        )::text
      );
      v_summary_snapshot := jsonb_build_object(
        'ruleKey', 'cgd_bank_statement_fdm_credit_card_monthly_income',
        'ruleVersion', 1,
        'calendarMonth', v_month.calendar_month,
        'sourceDescriptionPattern', '%POS VENDAS%',
        'destinationAccount', 'Credit Card',
        'operator', v_operator,
        'differenceAllowed', v_difference_allowed,
        'maxDifferenceDays', 31,
        'sourceCount', v_month.source_count,
        'sourceTotal', v_month.source_total,
        'destinationCount', v_month.destination_count,
        'destinationTotal', v_month.destination_total,
        'calculatedDifference', v_month.calculated_difference,
        'technicalBaseSourceId', v_month.technical_base_source_id,
        'technicalBaseSourceDate', v_month.technical_base_source_date,
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
        p_run_id, 'cgd_bank_statement_fdm_credit_card_monthly_income', 1,
        'import_cgd_extrato_ordem', v_month.technical_base_source_id,
        v_month.technical_base_source_date, v_base_snapshot,
        v_month.calculated_difference, v_difference_allowed,
        v_status, v_reason, v_signature,
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
            'cgd_bank_statement_fdm_credit_card_monthly_income'
          and proposal.rule_version = 1
          and proposal.base_source_type = 'import_cgd_extrato_ordem'
          and proposal.base_source_id = v_month.technical_base_source_id
          and proposal.signature = v_signature;
      else
        insert into public.financial_reconciliation_automatic_proposal_memberships (
          proposal_id, role, source_type, source_id, ordinal, source_date,
          amount, description, account, row_snapshot
        )
        select
          v_proposal_id,
          'source',
          'import_cgd_extrato_ordem',
          source_member.id,
          source_member.ordinal,
          source_member.data,
          source_member.montante,
          source_member.descritivo,
          '',
          jsonb_build_object(
            'sourceType', 'import_cgd_extrato_ordem',
            'sourceId', source_member.id,
            'sourceDate', source_member.data,
            'amount', source_member.montante,
            'description', source_member.descritivo
          )
        from (
          select bank.id, bank.data, bank.montante, bank.descritivo,
                 row_number() over (order by bank.data, bank.id)::integer
                   as ordinal
          from public.import_cgd_extrato_ordem bank
          where bank.id = any(v_source_ids)
            and bank.montante is not null
        ) source_member
        order by source_member.ordinal;
        get diagnostics v_inserted_count = row_count;
        if v_inserted_count <> v_month.source_count then
          raise exception 'Automatic monthly source membership insert was incomplete.';
        end if;

        insert into public.financial_reconciliation_automatic_proposal_memberships (
          proposal_id, role, source_type, source_id, ordinal, source_date,
          amount, description, account, row_snapshot
        )
        select
          v_proposal_id,
          'destination',
          'import_fdm_accounts',
          destination_member.id,
          destination_member.ordinal,
          destination_member.event_date,
          destination_member.amount,
          destination_member.description,
          destination_member.account,
          jsonb_build_object(
            'sourceType', 'import_fdm_accounts',
            'sourceId', destination_member.id,
            'sourceDate', destination_member.event_date,
            'amount', destination_member.amount,
            'description', destination_member.description,
            'account', destination_member.account
          )
        from (
          select fdm.id, fdm.event_date, fdm.amount, fdm.description,
                 fdm.account,
                 row_number() over (order by fdm.event_date, fdm.id)::integer
                   as ordinal
          from public.import_fdm_accounts fdm
          where fdm.id = any(v_destination_ids)
            and fdm.amount is not null
        ) destination_member
        order by destination_member.ordinal;
        get diagnostics v_inserted_count = row_count;
        if v_inserted_count <> v_month.destination_count then
          raise exception 'Automatic monthly destination membership insert was incomplete.';
        end if;

        select
          array_agg(membership.source_id order by membership.ordinal),
          sum(membership.amount)::numeric(14,2),
          (array_agg(membership.source_id order by membership.ordinal))[1],
          (array_agg(membership.source_date order by membership.ordinal))[1]
        into
          v_inserted_source_ids,
          v_inserted_source_total,
          v_inserted_base_id,
          v_inserted_base_date
        from public.financial_reconciliation_automatic_proposal_memberships membership
        where membership.proposal_id = v_proposal_id
          and membership.role = 'source';

        select
          array_agg(membership.source_id order by membership.ordinal),
          sum(membership.amount)::numeric(14,2)
        into v_inserted_destination_ids, v_inserted_destination_total
        from public.financial_reconciliation_automatic_proposal_memberships membership
        where membership.proposal_id = v_proposal_id
          and membership.role = 'destination';

        if v_inserted_source_ids is distinct from v_source_ids
          or v_inserted_destination_ids is distinct from v_destination_ids
          or v_inserted_source_total is distinct from v_month.source_total
          or v_inserted_destination_total is distinct from
            v_month.destination_total
          or v_inserted_base_id is distinct from
            v_month.technical_base_source_id
          or v_inserted_base_date is distinct from
            v_month.technical_base_source_date then
          raise exception 'Automatic monthly proposal summary and memberships diverged.';
        end if;
      end if;
    exception when others then
      raise;
    end;
  end loop;

  if v_page_count > 0 then
    update public.financial_reconciliation_automatic_runs
    set analysis_cursor_date = v_last_month,
        analysis_cursor_id = v_last_id,
        analysis_processed = greatest(
          analysis_processed,
          analysis_processed + v_page_count
        ),
        analysis_total = greatest(
          analysis_total,
          analysis_processed + v_page_count
        ),
        updated_at = now(),
        analysis_error_code = null,
        analysis_error_at = null
    where id = p_run_id
      and (
        analysis_cursor_date is null
        or v_last_month > analysis_cursor_date
      );
  end if;

  if v_page_count < 25 then
    return public.financial_reconciliation_finalize_automatic_analysis(p_run_id);
  end if;
  return public.financial_reconciliation_automatic_progress_or_run(p_run_id);
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
  v_contract jsonb;
  v_base record;
  v_combination record;
  v_combination_count integer;
  v_page_count integer := 0;
  v_total bigint;
  v_last_date date;
  v_last_id uuid;
  v_destination_source_type text;
  v_max_candidates integer;
  v_max_destination_records integer;
begin
  if nullif(trim(coalesce(p_actor, '')), '') is null then
    raise exception 'Actor is required.';
  end if;

  select * into v_run
  from public.financial_reconciliation_automatic_runs
  where id = p_run_id
  for update;
  if not found then raise exception 'Automatic analysis run was not found.'; end if;
  if v_run.actor <> p_actor then raise exception 'Automatic analysis run belongs to another actor.'; end if;
  if v_run.analysis_completed_at is not null then
    return public.get_financial_reconciliation_automatic_run(p_run_id);
  end if;
  if v_run.status <> 'analyzing' then
    raise exception 'Automatic analysis run is not resumable.';
  end if;

  begin
    if jsonb_array_length(v_run.definition_config_snapshot) <> 1 then
      raise exception 'Resumable automatic analysis requires exactly one snapshotted rule.';
    end if;

    select value into strict v_rule
    from jsonb_array_elements(v_run.definition_config_snapshot) value;

    v_contract := public.financial_reconciliation_automatic_rule_contract(
      v_rule->>'ruleKey',
      (v_rule->>'ruleVersion')::integer
    );
    if v_contract is null then
      raise exception 'Automatic reconciliation rule is unsupported.';
    end if;

    if v_rule->>'ruleKey' =
        'cgd_bank_statement_fdm_credit_card_monthly_income'
      and (v_rule->>'ruleVersion')::integer = 1 then
      return public.financial_reconciliation_continue_automatic_monthly_income(
        v_run.id,
        v_rule
      );
    end if;

    v_destination_source_type := coalesce(
      nullif(v_rule->>'destinationSourceType', ''),
      v_contract->>'destinationSourceType'
    );
    v_max_candidates := (v_contract->>'maxCandidates')::integer;
    v_max_destination_records := (v_contract->>'maxDestinationRecords')::integer;
    if v_destination_source_type is null
      or v_destination_source_type <> v_contract->>'destinationSourceType'
      or coalesce(v_rule->>'operator', '') not in ('+', '-')
      or coalesce(v_max_candidates, 0) < 1
      or coalesce(v_max_destination_records, 0) < 1 then
      raise exception 'Automatic rule snapshot contract is invalid.';
    end if;

    if v_run.analysis_total = 0 then
      select public.financial_reconciliation_automatic_base_count(
        v_rule->>'ruleKey',
        (v_rule->>'ruleVersion')::integer
      ) into v_total;
      update public.financial_reconciliation_automatic_runs
      set analysis_total = v_total, updated_at = now()
      where id = v_run.id;
      v_run.analysis_total := v_total;
    end if;

    for v_base in
      select *
      from public.financial_reconciliation_automatic_candidate_page(
        v_rule->>'ruleKey',
        (v_rule->>'ruleVersion')::integer,
        (v_rule->>'differenceAllowed')::numeric,
        (v_rule->>'maxDifferenceDays')::integer,
        v_run.analysis_cursor_date,
        v_run.analysis_cursor_id,
        25
      )
    loop
      v_page_count := v_page_count + 1;
      v_last_date := v_base.base_source_date;
      v_last_id := v_base.base_source_id;

      if v_base.candidate_count > v_max_candidates then
        insert into public.financial_reconciliation_automatic_proposals (
          run_id, rule_key, rule_version, base_source_type,
          base_source_id, base_source_date, base_snapshot, candidate_groups,
          allowed_difference, status, reason, signature
        ) values (
          v_run.id, v_rule->>'ruleKey', (v_rule->>'ruleVersion')::integer,
          'financial_documents', v_base.base_source_id, v_base.base_source_date,
          v_base.base_snapshot, v_base.candidates,
          (v_rule->>'differenceAllowed')::numeric,
          'ambiguous', 'candidate_limit',
          public.financial_reconciliation_extension_sha256(
            'candidate_limit:' || v_base.base_source_id::text
          )
        ) on conflict do nothing;
      else
        select count(*) into v_combination_count
        from public.financial_reconciliation_automatic_build_combinations(
          v_base.base_snapshot,
          v_base.candidates,
          jsonb_build_object(v_destination_source_type, v_rule->>'operator'),
          (v_rule->>'differenceAllowed')::numeric,
          v_max_destination_records
        );

        if v_combination_count = 1 then
          select * into strict v_combination
          from public.financial_reconciliation_automatic_build_combinations(
            v_base.base_snapshot,
            v_base.candidates,
            jsonb_build_object(v_destination_source_type, v_rule->>'operator'),
            (v_rule->>'differenceAllowed')::numeric,
            v_max_destination_records
          );
          insert into public.financial_reconciliation_automatic_proposals (
            run_id, rule_key, rule_version, base_source_type,
            base_source_id, base_source_date, base_snapshot, items, evidence,
            candidate_groups, calculated_difference, allowed_difference,
            status, signature
          ) values (
            v_run.id, v_rule->>'ruleKey', (v_rule->>'ruleVersion')::integer,
            'financial_documents', v_base.base_source_id, v_base.base_source_date,
            v_base.base_snapshot, v_combination.items,
            (select coalesce(jsonb_agg(value->'evidence'), '[]'::jsonb)
             from jsonb_array_elements(v_combination.items) value),
            jsonb_build_array(v_combination.items),
            v_combination.calculated_difference,
            (v_rule->>'differenceAllowed')::numeric,
            'proposed', v_combination.signature
          ) on conflict do nothing;
        elsif v_combination_count > 1 then
          insert into public.financial_reconciliation_automatic_proposals (
            run_id, rule_key, rule_version, base_source_type,
            base_source_id, base_source_date, base_snapshot, candidate_groups,
            allowed_difference, status, reason, signature
          ) values (
            v_run.id, v_rule->>'ruleKey', (v_rule->>'ruleVersion')::integer,
            'financial_documents', v_base.base_source_id, v_base.base_source_date,
            v_base.base_snapshot,
            (select coalesce(jsonb_agg(items order by signature), '[]'::jsonb)
             from public.financial_reconciliation_automatic_build_combinations(
               v_base.base_snapshot,
               v_base.candidates,
               jsonb_build_object(v_destination_source_type, v_rule->>'operator'),
               (v_rule->>'differenceAllowed')::numeric,
               v_max_destination_records
             )),
            (v_rule->>'differenceAllowed')::numeric,
            'ambiguous', 'multiple_combinations',
            public.financial_reconciliation_extension_sha256(
              'multiple:' || v_base.base_source_id::text
            )
          ) on conflict do nothing;
        else
          insert into public.financial_reconciliation_automatic_proposals (
            run_id, rule_key, rule_version, base_source_type,
            base_source_id, base_source_date, base_snapshot, candidate_groups,
            allowed_difference, status, reason, signature
          ) values (
            v_run.id, v_rule->>'ruleKey', (v_rule->>'ruleVersion')::integer,
            'financial_documents', v_base.base_source_id, v_base.base_source_date,
            v_base.base_snapshot, v_base.candidates,
            (v_rule->>'differenceAllowed')::numeric,
            'skipped', 'no_qualifying_combination',
            public.financial_reconciliation_extension_sha256(
              'skipped:' || v_base.base_source_id::text
            )
          ) on conflict do nothing;
        end if;
      end if;
    end loop;

    if v_page_count > 0 then
      update public.financial_reconciliation_automatic_runs
      set analysis_cursor_date = v_last_date,
          analysis_cursor_id = v_last_id,
          analysis_processed = analysis_processed + v_page_count,
          analysis_total = greatest(analysis_total, analysis_processed + v_page_count),
          updated_at = now(),
          analysis_error_code = null,
          analysis_error_at = null
      where id = v_run.id;
    end if;

    if v_page_count < 25 then
      return public.financial_reconciliation_finalize_automatic_analysis(v_run.id);
    end if;
    return public.financial_reconciliation_automatic_progress_or_run(v_run.id);
  exception when others then
    update public.financial_reconciliation_automatic_runs
    set status = 'failed',
        error_summary = 'Automatic analysis could not be completed.',
        analysis_error_code = 'analysis_continuation_failed',
        analysis_error_at = now(),
        finished_at = coalesce(finished_at, now()),
        updated_at = now()
    where id = p_run_id;
    return public.get_financial_reconciliation_automatic_run(p_run_id);
  end;
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
  v_total bigint;
begin
  if nullif(trim(coalesce(p_actor, '')), '') is null then raise exception 'Actor is required.'; end if;
  if p_client_request_id is null then raise exception 'Client request ID is required.'; end if;
  if p_mode is null or p_rule_keys is null then
    raise exception 'Manual automatic analysis requires exactly one selected rule.';
  end if;
  if p_mode <> 'manual_rule' or cardinality(p_rule_keys) <> 1 then
    raise exception 'Manual automatic analysis requires exactly one selected rule.';
  end if;

  perform pg_advisory_xact_lock(
    hashtextextended('financial_reconciliation_automatic_manual:' || p_actor, 0)
  );

  select run.id, run.client_request_id into v_current_run_id, v_current_request_id
  from public.financial_reconciliation_automatic_runs run
  where run.actor = p_actor and run.trigger = 'manual' and run.finished_at is null
  for update;
  if v_current_run_id is not null then
    if v_current_request_id = p_client_request_id then
      return public.continue_financial_reconciliation_automatic_analysis(v_current_run_id, p_actor);
    end if;
    raise exception 'Automatic analysis conflict: an unfinished manual run already exists for this actor.';
  end if;

  select run.id into v_run_id
  from public.financial_reconciliation_automatic_runs run
  where run.actor = p_actor and run.client_request_id = p_client_request_id;
  if v_run_id is not null then
    return public.continue_financial_reconciliation_automatic_analysis(v_run_id, p_actor);
  end if;

  lock table public.financial_reconciliation_source_rules in share row exclusive mode;
  lock table public.financial_reconciliation_automatic_rule_configs in share row exclusive mode;
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
    and config.max_difference_days between 0 and 90;

  if jsonb_array_length(v_snapshot) <> 1 then
    raise exception 'Automatic rule is not enabled for manual analysis.';
  end if;

  if v_snapshot->0->>'ruleKey' =
      'cgd_bank_statement_fdm_credit_card_monthly_income'
    and (v_snapshot->0->>'ruleVersion')::integer = 1 then
    select public.financial_reconciliation_automatic_monthly_income_count()
    into v_total;
  else
    select public.financial_reconciliation_automatic_base_count(
      v_snapshot->0->>'ruleKey',
      (v_snapshot->0->>'ruleVersion')::integer
    ) into v_total;
  end if;

  insert into public.financial_reconciliation_automatic_runs (
    trigger, scope, actor, client_request_id, definition_config_snapshot,
    analysis_processed, analysis_total
  ) values (
    'manual', 'rule', p_actor, p_client_request_id, v_snapshot, 0, v_total
  ) returning id into v_run_id;

  return public.continue_financial_reconciliation_automatic_analysis(v_run_id, p_actor);
end
$$;

create or replace function public.financial_reconciliation_finalize_automatic_analysis(p_run_id uuid)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
begin
  with display_source_usage as (
    select
      item->>'sourceType' as source_type,
      item->>'sourceId' as source_id,
      count(distinct proposal.base_source_id) as base_count
    from public.financial_reconciliation_automatic_proposals proposal
    join lateral (
      select item.value as item
      from jsonb_array_elements(proposal.items) item(value)
      union all
      select item.value as item
      from jsonb_array_elements(proposal.candidate_groups) candidate_group(value)
      join lateral jsonb_array_elements(
        case
          when jsonb_typeof(candidate_group.value) = 'array' then candidate_group.value
          else jsonb_build_array(candidate_group.value)
        end
      ) item(value) on true
    ) source_item on true
    where proposal.run_id = p_run_id and proposal.status in ('proposed', 'ambiguous')
    group by item->>'sourceType', item->>'sourceId'
  ), display_overlapping as (
    select distinct proposal.id
    from public.financial_reconciliation_automatic_proposals proposal
    join lateral (
      select item.value as item
      from jsonb_array_elements(proposal.items) item(value)
      union all
      select item.value as item
      from jsonb_array_elements(proposal.candidate_groups) candidate_group(value)
      join lateral jsonb_array_elements(
        case
          when jsonb_typeof(candidate_group.value) = 'array' then candidate_group.value
          else jsonb_build_array(candidate_group.value)
        end
      ) item(value) on true
    ) source_item on true
    join display_source_usage usage
      on usage.source_type = item->>'sourceType'
     and usage.source_id = item->>'sourceId'
    where proposal.run_id = p_run_id
      and proposal.status in ('proposed', 'ambiguous')
      and usage.base_count > 1
  ), amount_only_memberships as (
    select
      proposal.id as proposal_id,
      proposal.base_source_id,
      'import_cgd_extrato_ordem'::text as source_type,
      bank.id as source_id
    from public.financial_reconciliation_automatic_proposals proposal
    join public.financial_reconciliation_automatic_runs run
      on run.id = proposal.run_id
    cross join lateral jsonb_array_elements(run.definition_config_snapshot) snapshot(rule)
    join public.import_cgd_extrato_ordem bank
      on bank.montante = -(proposal.base_snapshot->>'amount')::numeric
     and round(bank.montante * 100)::bigint =
         -round((proposal.base_snapshot->>'amount')::numeric * 100)::bigint
     and bank.data between
         proposal.base_source_date - (snapshot.rule->>'maxDifferenceDays')::integer
         and proposal.base_source_date + (snapshot.rule->>'maxDifferenceDays')::integer
    where proposal.run_id = p_run_id
      and proposal.rule_key = 'financial_documents_cgd_bank_statement_amount_only'
      and proposal.rule_version = 1
      and proposal.status in ('proposed', 'ambiguous')
      and proposal.allowed_difference = 0
      and snapshot.rule->>'ruleKey' = proposal.rule_key
      and (snapshot.rule->>'ruleVersion')::integer = proposal.rule_version
      and (snapshot.rule->>'differenceAllowed')::numeric = 0
      and (snapshot.rule->>'maxDifferenceDays')::integer between 0 and 90
      and bank.data is not null
      and bank.data >= date '2026-01-01'
      and bank.montante is not null
      and not exists (
        select 1
        from public.financial_reconciliation_items item
        where item.source_type = 'import_cgd_extrato_ordem'
          and item.source_id = bank.id
      )
    union all
    select
      proposal.id as proposal_id,
      proposal.base_source_id,
      'import_cgd_cartao_credito'::text as source_type,
      card.id as source_id
    from public.financial_reconciliation_automatic_proposals proposal
    join public.financial_reconciliation_automatic_runs run
      on run.id = proposal.run_id
    cross join lateral jsonb_array_elements(run.definition_config_snapshot) snapshot(rule)
    join public.import_cgd_cartao_credito card
      on card.valor = -(proposal.base_snapshot->>'amount')::numeric
     and round(card.valor * 100)::bigint =
         -round((proposal.base_snapshot->>'amount')::numeric * 100)::bigint
     and card.data between
         proposal.base_source_date - (snapshot.rule->>'maxDifferenceDays')::integer
         and proposal.base_source_date + (snapshot.rule->>'maxDifferenceDays')::integer
    where proposal.run_id = p_run_id
      and proposal.rule_key = 'financial_documents_cgd_credit_card_amount_only'
      and proposal.rule_version = 1
      and proposal.status in ('proposed', 'ambiguous')
      and proposal.allowed_difference = 0
      and snapshot.rule->>'ruleKey' = proposal.rule_key
      and (snapshot.rule->>'ruleVersion')::integer = proposal.rule_version
      and (snapshot.rule->>'differenceAllowed')::numeric = 0
      and (snapshot.rule->>'maxDifferenceDays')::integer between 0 and 90
      and card.data is not null
      and card.data >= date '2026-01-01'
      and card.valor is not null
      and not exists (
        select 1
        from public.financial_reconciliation_items item
        where item.source_type = 'import_cgd_cartao_credito'
          and item.source_id = card.id
      )
  ), amount_only_source_usage as (
    select
      membership.source_type,
      membership.source_id,
      count(distinct membership.base_source_id) as base_count
    from amount_only_memberships membership
    group by membership.source_type, membership.source_id
  ), amount_only_overlapping as (
    select distinct membership.proposal_id as id
    from amount_only_memberships membership
    join amount_only_source_usage usage
      on usage.source_type = membership.source_type
     and usage.source_id = membership.source_id
    where usage.base_count > 1
  ), overlapping as (
    select id from display_overlapping
    union
    select id from amount_only_overlapping
  )
  update public.financial_reconciliation_automatic_proposals proposal
  set status = 'ambiguous', reason = 'cross_base_overlap', updated_at = now()
  where proposal.id in (select id from overlapping);

  update public.financial_reconciliation_automatic_runs run
  set status = case when exists (
        select 1
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = p_run_id and proposal.status = 'proposed'
      ) then 'ready' else 'completed' end,
      finished_at = case when exists (
        select 1
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = p_run_id and proposal.status = 'proposed'
      ) then null else now() end,
      analysis_completed_at = now(),
      updated_at = now(),
      analysis_error_code = null,
      analysis_error_at = null,
      counts = (
        select jsonb_build_object(
          'bases', count(distinct proposal.base_source_id),
          'proposed', count(*) filter (where proposal.status = 'proposed'),
          'ambiguous', count(*) filter (where proposal.status = 'ambiguous'),
          'skipped', count(*) filter (where proposal.status = 'skipped')
        )
        from public.financial_reconciliation_automatic_proposals proposal
        where proposal.run_id = p_run_id
      )
  where run.id = p_run_id and run.analysis_completed_at is null;

  return public.get_financial_reconciliation_automatic_run(p_run_id);
end
$$;

create or replace function public.get_financial_reconciliation_automatic_run(p_run_id uuid)
returns jsonb
language plpgsql
stable
security definer set search_path = public, pg_temp
as $$
declare
  v_run public.financial_reconciliation_automatic_runs%rowtype;
  v_proposals jsonb;
begin
  select * into v_run
  from public.financial_reconciliation_automatic_runs
  where id = p_run_id;
  if not found then raise exception 'Automatic analysis run was not found.'; end if;

  select coalesce(jsonb_agg(jsonb_build_object(
    'id', proposal.id,
    'runId', proposal.run_id,
    'ruleKey', proposal.rule_key,
    'ruleVersion', proposal.rule_version,
    'baseSourceType', proposal.base_source_type,
    'baseSourceId', proposal.base_source_id,
    'baseSourceDate', proposal.base_source_date,
    'baseSnapshot', proposal.base_snapshot,
    'items', case
      when proposal.rule_key =
          'cgd_bank_statement_fdm_credit_card_monthly_income'
        and proposal.rule_version = 1 then '[]'::jsonb
      else proposal.items
    end,
    'evidence', proposal.evidence,
    'candidateGroups', case
      when proposal.rule_key =
          'cgd_bank_statement_fdm_credit_card_monthly_income'
        and proposal.rule_version = 1 then '[]'::jsonb
      else proposal.candidate_groups
    end,
    'groupingKey', proposal.grouping_key,
    'summarySnapshot', proposal.summary_snapshot,
    'calculatedDifference', proposal.calculated_difference,
    'allowedDifference', proposal.allowed_difference,
    'status', proposal.status,
    'reason', proposal.reason,
    'signature', proposal.signature,
    'reconciliationId', proposal.reconciliation_id,
    'createdAt', proposal.created_at,
    'updatedAt', proposal.updated_at
  ) order by proposal.base_source_date, proposal.base_source_id, proposal.signature), '[]'::jsonb)
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
    'analysisComplete', v_run.analysis_completed_at is not null,
    'analysisCompletedAt', v_run.analysis_completed_at,
    'startedAt', v_run.started_at,
    'finishedAt', v_run.finished_at,
    'proposals', v_proposals
  );
end
$$;

create or replace function public.financial_reconciliation_automatic_progress_or_run(p_run_id uuid)
returns jsonb
language plpgsql
stable
security definer set search_path = public, pg_temp
as $$
declare
  v_run public.financial_reconciliation_automatic_runs%rowtype;
begin
  select * into v_run
  from public.financial_reconciliation_automatic_runs
  where id = p_run_id;
  if not found then raise exception 'Automatic analysis run was not found.'; end if;
  if v_run.analysis_completed_at is not null then
    return public.get_financial_reconciliation_automatic_run(p_run_id);
  end if;

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
    'analysisComplete', false,
    'analysisCompletedAt', null,
    'startedAt', v_run.started_at,
    'finishedAt', v_run.finished_at,
    'proposals', '[]'::jsonb
  );
end
$$;

-- Preserve the exact four-rule executor installed by the prior migrations.
-- On reapply the private copy already exists, so the monthly dispatcher remains
-- stable without dynamic SQL or replacing the saved implementation.
do $migration$
begin
  if to_regprocedure(
      'public.financial_reconciliation_execute_prior_proposal(uuid,text)'
    ) is null then
    alter function public.execute_financial_reconciliation_automatic_proposal(uuid,text)
      rename to financial_reconciliation_execute_prior_proposal;
  end if;
end
$migration$;

create or replace function public.financial_reconciliation_execute_monthly_income_proposal(
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
  v_contract jsonb;
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
  v_stored_count bigint;
  v_locked_source_count bigint;
  v_locked_destination_count bigint;
  v_membership_mismatch boolean;
  v_stored_source_count bigint;
  v_stored_source_total numeric;
  v_stored_destination_count bigint;
  v_stored_destination_total numeric;
  v_live_source_count bigint;
  v_live_source_total numeric;
  v_live_destination_count bigint;
  v_live_destination_total numeric;
  v_live_difference numeric;
  v_technical_base_source_id uuid;
  v_technical_base_source_date date;
  v_technical_base_snapshot jsonb;
  v_source_ids uuid[];
  v_destination_ids uuid[];
  v_current_signature text;
  v_matching_source_types jsonb;
  v_matching_source_rules jsonb;
  v_membership_snapshots jsonb;
  v_reconciliation_id uuid;
  v_actual_item_count bigint;
  v_actual_difference numeric(14,2);
  v_comment text;
  v_failure_message text;
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
    'ambiguous', 'skipped', 'deselected', 'failed'
  ) then
    raise exception 'Automation proposal with status % cannot be executed.',
      v_proposal.status;
  end if;
  if v_proposal.status <> 'proposed' then
    raise exception 'Automation proposal is already being executed.';
  end if;
  if v_run.finished_at is not null then
    raise exception 'Automation proposal belongs to a finished run.';
  end if;
  if v_run.analysis_completed_at is null or v_run.status = 'analyzing' then
    raise exception 'Automatic analysis must finish before proposals can be executed.';
  end if;
  if v_proposal.rule_key <>
      'cgd_bank_statement_fdm_credit_card_monthly_income'
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

  -- Lock the current definition/config and required directional source rule
  -- before interpreting the immutable run snapshot.
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

  select source_rule.operator
  into v_current_operator
  from public.financial_reconciliation_source_rules source_rule
  where source_rule.base_source_type = 'import_cgd_extrato_ordem'
    and source_rule.matching_source_type = 'import_fdm_accounts'
  for share of source_rule;

  if v_current_definition is null
    or v_current_operator is null then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_snapshot_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_snapshot_changed'
    );
  end if;

  -- Validate every textual snapshot field and numeric bound before casting.
  if jsonb_typeof(v_run.definition_config_snapshot) <> 'array'
    or jsonb_array_length(v_run.definition_config_snapshot) <> 1
    or jsonb_typeof(v_run.definition_config_snapshot->0) <> 'object' then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_snapshot_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_snapshot_changed'
    );
  end if;
  v_rule_snapshot := v_run.definition_config_snapshot->0;
  if v_rule_snapshot->>'ruleKey' is distinct from v_proposal.rule_key
    or coalesce(v_rule_snapshot->>'ruleVersion', '') !~ '^[0-9]+$'
    or length(v_rule_snapshot->>'ruleVersion') > 10
    or v_rule_snapshot->>'displayName' is distinct from
      'Card Payments - POS - Income'
    or v_rule_snapshot->>'destinationSourceType' is distinct from
      'import_fdm_accounts'
    or v_rule_snapshot->>'operator' is distinct from '-'
    or coalesce(v_rule_snapshot->>'differenceAllowed', '') !~
      '^[0-9]{1,12}(\.[0-9]{1,2})?$'
    or coalesce(v_rule_snapshot->>'maxDifferenceDays', '') !~ '^[0-9]+$'
    or length(v_rule_snapshot->>'maxDifferenceDays') > 10
    or coalesce(v_rule_snapshot->>'priority', '') !~ '^[0-9]+$'
    or length(v_rule_snapshot->>'priority') > 10
    or v_rule_snapshot->'definition' is distinct from jsonb_build_object(
      'matchingMode', 'monthly_aggregate',
      'sourceDescriptionPattern', '%POS VENDAS%',
      'destinationAccount', 'Credit Card',
      'calendarGrouping', 'closed_month',
      'fixedMaxDifferenceDays', 31,
      'eligibilityFloor', '2026-01-01',
      'requiresNonNullAmount', true
    ) then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_snapshot_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_snapshot_changed'
    );
  end if;
  if (v_rule_snapshot->>'ruleVersion')::bigint not between 1 and 2147483647
    or (v_rule_snapshot->>'maxDifferenceDays')::bigint not between 0 and 90
    or (v_rule_snapshot->>'priority')::bigint not between 1 and 2147483647 then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_snapshot_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_snapshot_changed'
    );
  end if;

  v_snapshot_rule_version := (v_rule_snapshot->>'ruleVersion')::integer;
  v_snapshot_difference_allowed :=
    (v_rule_snapshot->>'differenceAllowed')::numeric(14,2);
  v_snapshot_max_difference_days :=
    (v_rule_snapshot->>'maxDifferenceDays')::integer;
  v_snapshot_priority := (v_rule_snapshot->>'priority')::integer;

  v_contract := public.financial_reconciliation_automatic_rule_contract(
    v_proposal.rule_key, v_proposal.rule_version
  );
  if v_contract is distinct from jsonb_build_object(
      'destinationSourceType', 'import_fdm_accounts',
      'matchingMode', 'monthly_aggregate',
      'sourceDescriptionPattern', '%POS VENDAS%',
      'destinationAccount', 'Credit Card',
      'calendarGrouping', 'closed_month',
      'eligibilityFloor', '2026-01-01',
      'fixedMaxDifferenceDays', 31,
      'requiresNonNullAmount', true
    )
    or v_snapshot_rule_version <> v_proposal.rule_version
    or v_snapshot_max_difference_days <> 31
    or v_current_rule_version is distinct from v_snapshot_rule_version
    or not v_current_enabled
    or (v_run.trigger = 'manual' and not v_current_allow_manual)
    or (v_run.trigger = 'scheduled' and not v_current_include_scheduled)
    or v_current_definition is distinct from v_rule_snapshot->'definition'
    or v_current_display_name is distinct from
      v_rule_snapshot->>'displayName'
    or v_current_base_source_type is distinct from
      'import_cgd_extrato_ordem'
    or v_current_destination_source_types is distinct from
      '["import_fdm_accounts"]'::jsonb
    or v_current_max_difference_days is distinct from
      v_snapshot_max_difference_days
    or v_current_priority is distinct from v_snapshot_priority then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_snapshot_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_snapshot_changed'
    );
  end if;
  if v_current_operator is distinct from '-' then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'operator_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'operator_changed'
    );
  end if;
  if v_current_difference_allowed is distinct from
      v_snapshot_difference_allowed
    or v_proposal.allowed_difference is distinct from
      v_snapshot_difference_allowed then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'tolerance_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'tolerance_changed'
    );
  end if;

  if jsonb_typeof(v_proposal.summary_snapshot) <> 'object'
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
    or v_proposal.summary_snapshot->>'sourceDescriptionPattern' is distinct from
      '%POS VENDAS%'
    or v_proposal.summary_snapshot->>'destinationAccount' is distinct from
      'Credit Card'
    or v_proposal.summary_snapshot->>'operator' is distinct from '-'
    or v_proposal.summary_snapshot->>'maxDifferenceDays' is distinct from '31'
    or coalesce(v_proposal.summary_snapshot->>'sourceCount', '') !~ '^[0-9]+$'
    or length(v_proposal.summary_snapshot->>'sourceCount') > 10
    or coalesce(v_proposal.summary_snapshot->>'destinationCount', '') !~
      '^[0-9]+$'
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
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'source_snapshot_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'source_snapshot_changed'
    );
  end if;
  v_calendar_month :=
    (v_proposal.summary_snapshot->>'calendarMonth')::date;

  -- Global monthly-execution table-lock order: Bank source, FDM source, then
  -- reconciliation item locks. SHARE ROW EXCLUSIVE conflicts with every
  -- external INSERT/UPDATE/DELETE writer (ROW EXCLUSIVE) and with another
  -- monthly executor, while this transaction may still write its own item
  -- locks. These transaction-scoped locks are held through lifecycle completion.
  lock table
    public.import_cgd_extrato_ordem,
    public.import_fdm_accounts,
    public.financial_reconciliation_items
  in share row exclusive mode;

  -- Only after the predicate-level writer barrier is held, lock immutable
  -- membership rows and live members in deterministic source/date/ID order,
  -- then recompute and compare the entire eligible month before any write.
  perform membership.source_id
  from public.financial_reconciliation_automatic_proposal_memberships membership
  where membership.proposal_id = v_proposal.id
  order by membership.source_type, membership.source_date, membership.source_id
  for share;
  get diagnostics v_stored_count = row_count;

  perform bank.id
  from public.import_cgd_extrato_ordem bank
  where bank.data >= v_calendar_month
    and bank.data < v_calendar_month + interval '1 month'
    and bank.data >= date '2026-01-01'
    and bank.data < date_trunc('month', current_date)::date
    and bank.montante is not null
    and bank.descritivo ilike '%POS VENDAS%'
    and not exists (
      select 1
      from public.financial_reconciliation_items locked
      where locked.source_type = 'import_cgd_extrato_ordem'
        and locked.source_id = bank.id
    )
  order by bank.data, bank.id
  for update of bank;
  get diagnostics v_locked_source_count = row_count;

  perform fdm.id
  from public.import_fdm_accounts fdm
  where fdm.event_date >= v_calendar_month
    and fdm.event_date < v_calendar_month + interval '1 month'
    and fdm.event_date >= date '2026-01-01'
    and fdm.event_date < date_trunc('month', current_date)::date
    and fdm.amount is not null
    and fdm.account = 'Credit Card'
    and not exists (
      select 1
      from public.financial_reconciliation_items locked
      where locked.source_type = 'import_fdm_accounts'
        and locked.source_id = fdm.id
    )
  order by fdm.event_date, fdm.id
  for update of fdm;
  get diagnostics v_locked_destination_count = row_count;

  select
    count(*) filter (where membership.role = 'source'),
    sum(membership.amount) filter (where membership.role = 'source'),
    count(*) filter (where membership.role = 'destination'),
    sum(membership.amount) filter (where membership.role = 'destination'),
    (array_agg(membership.source_id order by membership.source_date,
               membership.source_id)
       filter (where membership.role = 'source'))[1],
    (array_agg(membership.source_date order by membership.source_date,
               membership.source_id)
       filter (where membership.role = 'source'))[1],
    (array_agg(membership.row_snapshot order by membership.source_date,
               membership.source_id)
       filter (where membership.role = 'source'))[1],
    array_agg(membership.source_id order by membership.ordinal)
      filter (where membership.role = 'source'),
    array_agg(membership.source_id order by membership.ordinal)
      filter (where membership.role = 'destination'),
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
    ) order by membership.source_type, membership.source_date,
               membership.source_id)
  into
    v_stored_source_count,
    v_stored_source_total,
    v_stored_destination_count,
    v_stored_destination_total,
    v_technical_base_source_id,
    v_technical_base_source_date,
    v_technical_base_snapshot,
    v_source_ids,
    v_destination_ids,
    v_membership_snapshots
  from public.financial_reconciliation_automatic_proposal_memberships membership
  where membership.proposal_id = v_proposal.id;

  with stored as (
    select
      membership.role,
      membership.source_type,
      membership.source_id,
      membership.source_date,
      membership.amount
    from public.financial_reconciliation_automatic_proposal_memberships membership
    where membership.proposal_id = v_proposal.id
  ), live as (
    select
      'source'::text as role,
      'import_cgd_extrato_ordem'::text as source_type,
      bank.id as source_id,
      bank.data as source_date,
      bank.montante as amount
    from public.import_cgd_extrato_ordem bank
    where bank.data >= v_calendar_month
      and bank.data < v_calendar_month + interval '1 month'
      and bank.data >= date '2026-01-01'
      and bank.data < date_trunc('month', current_date)::date
      and bank.montante is not null
      and bank.descritivo ilike '%POS VENDAS%'
      and not exists (
        select 1
        from public.financial_reconciliation_items locked
        where locked.source_type = 'import_cgd_extrato_ordem'
          and locked.source_id = bank.id
      )
    union all
    select
      'destination'::text,
      'import_fdm_accounts'::text,
      fdm.id,
      fdm.event_date,
      fdm.amount
    from public.import_fdm_accounts fdm
    where fdm.event_date >= v_calendar_month
      and fdm.event_date < v_calendar_month + interval '1 month'
      and fdm.event_date >= date '2026-01-01'
      and fdm.event_date < date_trunc('month', current_date)::date
      and fdm.amount is not null
      and fdm.account = 'Credit Card'
      and not exists (
        select 1
        from public.financial_reconciliation_items locked
        where locked.source_type = 'import_fdm_accounts'
          and locked.source_id = fdm.id
      )
  ), lost as (
    select * from stored
    except
    select * from live
  ), gained as (
    select * from live
    except
    select * from stored
  )
  select exists (select 1 from lost) or exists (select 1 from gained)
  into v_membership_mismatch;

  select
    count(*) filter (where live.role = 'source'),
    sum(live.amount) filter (where live.role = 'source'),
    count(*) filter (where live.role = 'destination'),
    sum(live.amount) filter (where live.role = 'destination')
  into
    v_live_source_count,
    v_live_source_total,
    v_live_destination_count,
    v_live_destination_total
  from (
    select 'source'::text as role, bank.montante as amount
    from public.import_cgd_extrato_ordem bank
    where bank.data >= v_calendar_month
      and bank.data < v_calendar_month + interval '1 month'
      and bank.data >= date '2026-01-01'
      and bank.data < date_trunc('month', current_date)::date
      and bank.montante is not null
      and bank.descritivo ilike '%POS VENDAS%'
      and not exists (
        select 1 from public.financial_reconciliation_items locked
        where locked.source_type = 'import_cgd_extrato_ordem'
          and locked.source_id = bank.id
      )
    union all
    select 'destination'::text, fdm.amount
    from public.import_fdm_accounts fdm
    where fdm.event_date >= v_calendar_month
      and fdm.event_date < v_calendar_month + interval '1 month'
      and fdm.event_date >= date '2026-01-01'
      and fdm.event_date < date_trunc('month', current_date)::date
      and fdm.amount is not null
      and fdm.account = 'Credit Card'
      and not exists (
        select 1 from public.financial_reconciliation_items locked
        where locked.source_type = 'import_fdm_accounts'
          and locked.source_id = fdm.id
      )
  ) live;
  v_live_difference := round(
    v_live_source_total - v_live_destination_total, 2
  );

  v_current_signature := public.financial_reconciliation_extension_sha256(
    jsonb_build_object(
      'ruleKey', v_proposal.rule_key,
      'ruleVersion', v_proposal.rule_version,
      'calendarMonth', v_calendar_month,
      'sourceIds', to_jsonb(v_source_ids),
      'destinationIds', to_jsonb(v_destination_ids),
      'sourceTotal', v_stored_source_total,
      'destinationTotal', v_stored_destination_total,
      'calculatedDifference',
        (v_stored_source_total - v_stored_destination_total)::numeric(14,2),
      'differenceAllowed', v_snapshot_difference_allowed,
      'operator', '-'
    )::text
  );

  if v_stored_count <> v_stored_source_count + v_stored_destination_count
    or v_locked_source_count <> v_stored_source_count
    or v_locked_destination_count <> v_stored_destination_count
    or v_stored_source_count < 1
    or v_stored_destination_count < 1
    or exists (
      select 1
      from (
        select
          membership.*,
          row_number() over (
            partition by membership.role
            order by membership.source_date, membership.source_id
          )::integer as expected_ordinal
        from public.financial_reconciliation_automatic_proposal_memberships membership
        where membership.proposal_id = v_proposal.id
      ) membership
      where membership.ordinal <> membership.expected_ordinal
        or (
          membership.role = 'source'
          and (
            membership.source_type <> 'import_cgd_extrato_ordem'
            or membership.account <> ''
            or membership.row_snapshot is distinct from jsonb_build_object(
              'sourceType', membership.source_type,
              'sourceId', membership.source_id,
              'sourceDate', membership.source_date,
              'amount', membership.amount,
              'description', membership.description
            )
          )
        )
        or (
          membership.role = 'destination'
          and (
            membership.source_type <> 'import_fdm_accounts'
            or membership.row_snapshot is distinct from jsonb_build_object(
              'sourceType', membership.source_type,
              'sourceId', membership.source_id,
              'sourceDate', membership.source_date,
              'amount', membership.amount,
              'description', membership.description,
              'account', membership.account
            )
          )
        )
    )
    or v_membership_mismatch
    or v_live_source_count is distinct from v_stored_source_count
    or v_live_source_total is distinct from v_stored_source_total
    or v_live_destination_count is distinct from v_stored_destination_count
    or v_live_destination_total is distinct from v_stored_destination_total
    or (v_proposal.summary_snapshot->>'sourceCount')::bigint <>
      v_stored_source_count
    or (v_proposal.summary_snapshot->>'sourceTotal')::numeric(14,2) is distinct from
      v_stored_source_total
    or (v_proposal.summary_snapshot->>'destinationCount')::bigint <>
      v_stored_destination_count
    or (v_proposal.summary_snapshot->>'destinationTotal')::numeric(14,2)
      is distinct from v_stored_destination_total
    or (v_proposal.summary_snapshot->>'calculatedDifference')::numeric(14,2)
      is distinct from v_live_difference
    or (v_proposal.summary_snapshot->>'differenceAllowed')::numeric(14,2)
      is distinct from v_snapshot_difference_allowed
    or v_proposal.calculated_difference is distinct from v_live_difference
    or v_proposal.base_source_id is distinct from
      v_technical_base_source_id
    or v_proposal.base_source_date is distinct from
      v_technical_base_source_date
    or v_proposal.base_snapshot is distinct from v_technical_base_snapshot
    or v_proposal.summary_snapshot->>'technicalBaseSourceId' is distinct from
      v_technical_base_source_id::text
    or v_proposal.summary_snapshot->>'technicalBaseSourceDate'
      is distinct from v_technical_base_source_date::text
    or v_proposal.signature is distinct from v_current_signature
    or abs(v_live_difference) > v_snapshot_difference_allowed then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'source_snapshot_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'source_snapshot_changed'
    );
  end if;

  v_matching_source_types := '["import_fdm_accounts"]'::jsonb;
  v_matching_source_rules := jsonb_build_array(jsonb_build_object(
    'sourceType', 'import_fdm_accounts', 'operator', '-'
  ));

  begin
    update public.financial_reconciliation_automatic_proposals
    set status = 'executing', reason = '', error = '', error_detail = '',
        updated_at = now()
    where id = v_proposal.id;

    insert into public.financial_reconciliations (
      status, base_source_type, matching_source_types,
      matching_source_rules, created_by
    ) values (
      'started', 'import_cgd_extrato_ordem', v_matching_source_types,
      v_matching_source_rules, p_actor
    ) returning id into v_reconciliation_id;

    insert into public.financial_reconciliation_items (
      reconciliation_id, source_type, source_id, amount_snapshot, created_by
    )
    select v_reconciliation_id, membership.source_type,
           membership.source_id, membership.amount, p_actor
    from public.financial_reconciliation_automatic_proposal_memberships membership
    where membership.proposal_id = v_proposal.id
      and membership.role = 'source'
      and membership.source_id = v_technical_base_source_id;

    update public.financial_reconciliations
    set difference_amount =
          (v_technical_base_snapshot->>'amount')::numeric(14,2),
        updated_at = timezone('utc', now())
    where id = v_reconciliation_id;

    insert into public.financial_reconciliation_audit (
      reconciliation_id, action, actor, difference_amount, metadata
    ) values (
      v_reconciliation_id, 'start', p_actor,
      (v_technical_base_snapshot->>'amount')::numeric(14,2),
      jsonb_build_object(
        'sourceType', 'import_cgd_extrato_ordem',
        'sourceId', v_technical_base_source_id,
        'differenceAmount',
          (v_technical_base_snapshot->>'amount')::numeric(14,2),
        'matchingSourceRules', v_matching_source_rules
      )
    );

    insert into public.financial_reconciliation_items (
      reconciliation_id, source_type, source_id, amount_snapshot, created_by
    )
    select v_reconciliation_id, membership.source_type,
           membership.source_id, membership.amount, p_actor
    from public.financial_reconciliation_automatic_proposal_memberships membership
    where membership.proposal_id = v_proposal.id
      and not (
        membership.role = 'source'
        and membership.source_id = v_technical_base_source_id
      )
    order by membership.source_type, membership.source_date,
             membership.source_id;

    with ordered as (
      select
        membership.*,
        row_number() over (
          order by membership.source_type, membership.source_date,
                   membership.source_id
        ) as sequence,
        sum(case
          when membership.source_type = 'import_cgd_extrato_ordem'
            then membership.amount
          else -membership.amount
        end) over (
          order by membership.source_type, membership.source_date,
                   membership.source_id
          rows between unbounded preceding and current row
        )::numeric(14,2) as running_difference
      from public.financial_reconciliation_automatic_proposal_memberships membership
      where membership.proposal_id = v_proposal.id
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
      'import_cgd_extrato_ordem', v_matching_source_rules,
      v_reconciliation_id
    ) into v_actual_difference;
    update public.financial_reconciliations
    set difference_amount = v_actual_difference,
        updated_at = timezone('utc', now())
    where id = v_reconciliation_id;

    if v_actual_item_count <> v_stored_count
      or v_actual_difference is distinct from v_live_difference
      or exists (
        select membership.source_type, membership.source_id, membership.amount
        from public.financial_reconciliation_automatic_proposal_memberships membership
        where membership.proposal_id = v_proposal.id
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
        where membership.proposal_id = v_proposal.id
      ) then
      raise exception 'Automatic monthly reconciliation lifecycle snapshots changed after revalidation.';
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
      v_comment := null;
      perform public.financial_reconciliation_action(
        'complete', p_actor, v_reconciliation_id, null, null, null
      );
    else
      v_comment := 'Automatic monthly reconciliation for '
        || v_proposal.grouping_key
        || ': Bank Statement total '
        || to_char(v_stored_source_total, 'FM999999999990.00')
        || ' EUR; FDM Credit Card total '
        || to_char(v_stored_destination_total, 'FM999999999990.00')
        || ' EUR; difference '
        || to_char(v_actual_difference, 'FM999999999990.00')
        || ' EUR within allowed '
        || to_char(v_snapshot_difference_allowed, 'FM999999999990.00')
        || ' EUR; run ' || v_run.id::text
        || '; proposal ' || v_proposal.id::text || '.';
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
          'displayName', v_rule_snapshot->>'displayName',
          'definition', v_rule_snapshot->'definition'
        ),
        'configSnapshot', jsonb_build_object(
          'differenceAllowed', v_snapshot_difference_allowed,
          'maxDifferenceDays', v_snapshot_max_difference_days,
          'priority', v_snapshot_priority
        ),
        'operatorSnapshot', jsonb_build_object(
          'import_fdm_accounts', '-'
        ),
        'summarySnapshot', v_proposal.summary_snapshot,
        'membershipSnapshots', v_membership_snapshots,
        'proposalSignature', v_proposal.signature,
        'trigger', v_run.trigger,
        'runId', v_run.id,
        'proposalId', v_proposal.id,
        'tolerance', v_snapshot_difference_allowed,
        'calculatedDifference', v_actual_difference
      )
    );

    if not exists (
      select 1
      from public.financial_reconciliations reconciliation
      where reconciliation.id = v_reconciliation_id
        and reconciliation.status = 'complete'
        and reconciliation.completion_type = case
          when v_actual_difference = 0 then 'normal' else 'forced'
        end
        and reconciliation.difference_amount = v_actual_difference
        and reconciliation.forced_completion_comment is not distinct from
          v_comment
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
      raise exception 'Automatic monthly reconciliation lifecycle snapshots changed after revalidation.';
    end if;

    update public.financial_reconciliation_automatic_proposals
    set status = 'completed',
        reconciliation_id = v_reconciliation_id,
        completed_at = now(),
        reason = '',
        error = '',
        error_detail = '',
        updated_at = now()
    where id = v_proposal.id;
  exception
    when unique_violation then
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
    when others then
      get stacked diagnostics
        v_failure_message = message_text;
      if v_failure_message in (
        'Automatic monthly reconciliation lifecycle snapshots changed after revalidation.',
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
          error_detail = '',
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
  v_rule_key text;
  v_rule_version integer;
begin
  if p_proposal_id is null then
    raise exception 'Automation proposal ID is required.';
  end if;
  if nullif(trim(coalesce(p_actor, '')), '') is null then
    raise exception 'Actor is required.';
  end if;

  select proposal.rule_key, proposal.rule_version
  into v_rule_key, v_rule_version
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.id = p_proposal_id;
  if not found then
    raise exception 'Automation proposal was not found.';
  end if;

  if v_rule_key = 'cgd_bank_statement_fdm_credit_card_monthly_income'
    and v_rule_version = 1 then
    return public.financial_reconciliation_execute_monthly_income_proposal(
      p_proposal_id, p_actor
    );
  elsif (v_rule_key, v_rule_version) in (
    ('financial_documents_cgd_bank_statement', 2),
    ('financial_documents_cgd_credit_card', 1),
    ('financial_documents_cgd_bank_statement_amount_only', 1),
    ('financial_documents_cgd_credit_card_amount_only', 1)
  ) then
    return public.financial_reconciliation_execute_prior_proposal(
      p_proposal_id, p_actor
    );
  end if;

  raise exception 'Automation proposal rule/version is not executable.';
end
$$;

revoke all on function public.financial_reconciliation_automatic_rule_contract(text,integer)
  from public, anon, authenticated;
revoke all on function public.financial_reconciliation_automatic_monthly_income_count()
  from public, anon, authenticated;
revoke all on function public.financial_reconciliation_automatic_monthly_income_page(date,integer)
  from public, anon, authenticated;
revoke all on function public.financial_reconciliation_continue_automatic_monthly_income(uuid,jsonb)
  from public, anon, authenticated, service_role;
revoke all on function public.continue_financial_reconciliation_automatic_analysis(uuid,text)
  from public, anon, authenticated;
revoke all on function public.create_financial_reconciliation_automatic_analysis(text[],text,text,uuid)
  from public, anon, authenticated;
revoke all on function public.financial_reconciliation_finalize_automatic_analysis(uuid)
  from public, anon, authenticated;
revoke all on function public.financial_reconciliation_automatic_progress_or_run(uuid)
  from public, anon, authenticated;
revoke all on function public.get_financial_reconciliation_automatic_run(uuid)
  from public, anon, authenticated;
revoke all on function public.financial_reconciliation_execute_prior_proposal(uuid,text)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_execute_monthly_income_proposal(uuid,text)
  from public, anon, authenticated, service_role;
revoke all on function public.execute_financial_reconciliation_automatic_proposal(uuid,text)
  from public, anon, authenticated, service_role;

grant execute on function public.financial_reconciliation_automatic_rule_contract(text,integer)
  to service_role;
grant execute on function public.financial_reconciliation_automatic_monthly_income_count()
  to service_role;
grant execute on function public.financial_reconciliation_automatic_monthly_income_page(date,integer)
  to service_role;
grant execute on function public.continue_financial_reconciliation_automatic_analysis(uuid,text)
  to service_role;
grant execute on function public.create_financial_reconciliation_automatic_analysis(text[],text,text,uuid)
  to service_role;
grant execute on function public.financial_reconciliation_finalize_automatic_analysis(uuid)
  to service_role;
grant execute on function public.financial_reconciliation_automatic_progress_or_run(uuid)
  to service_role;
grant execute on function public.get_financial_reconciliation_automatic_run(uuid)
  to service_role;
grant execute on function public.execute_financial_reconciliation_automatic_proposal(uuid,text)
  to service_role;

notify pgrst, 'reload schema';
