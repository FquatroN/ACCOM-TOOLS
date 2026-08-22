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
    'eligibilityFloor', '2026-01-01'
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

notify pgrst, 'reload schema';
