create extension if not exists pgcrypto;
create extension if not exists unaccent;
create extension if not exists pg_trgm;

create table if not exists public.financial_reconciliation_automatic_rule_definitions (
  rule_key text not null,
  version integer not null check (version > 0),
  display_name text not null,
  base_source_type text not null,
  destination_source_types jsonb not null check (jsonb_typeof(destination_source_types) = 'array'),
  logic_description text not null,
  definition jsonb not null check (jsonb_typeof(definition) = 'object'),
  created_at timestamptz not null default now(),
  primary key (rule_key, version)
);

create table if not exists public.financial_reconciliation_automatic_rule_configs (
  rule_key text primary key,
  rule_version integer not null,
  enabled boolean not null default false,
  allow_manual_execution boolean not null default false,
  include_in_scheduled_batch boolean not null default false,
  difference_allowed numeric(14,2) not null default 0 check (difference_allowed >= 0),
  max_difference_days integer not null default 7 check (max_difference_days between 0 and 365),
  priority integer not null check (priority > 0),
  updated_by text not null default '',
  updated_at timestamptz not null default now(),
  foreign key (rule_key, rule_version) references public.financial_reconciliation_automatic_rule_definitions(rule_key, version),
  constraint financial_reconciliation_automatic_rule_configs_priority_key
    unique (priority) deferrable initially deferred
);

do $$
begin
  if exists (
    select 1 from pg_constraint
    where conrelid = 'public.financial_reconciliation_automatic_rule_configs'::regclass
      and conname = 'financial_reconciliation_automatic_rule_configs_priority_key'
      and not condeferrable
  ) then
    alter table public.financial_reconciliation_automatic_rule_configs
      drop constraint financial_reconciliation_automatic_rule_configs_priority_key;
    alter table public.financial_reconciliation_automatic_rule_configs
      add constraint financial_reconciliation_automatic_rule_configs_priority_key
      unique (priority) deferrable initially deferred;
  elsif not exists (
    select 1 from pg_constraint
    where conrelid = 'public.financial_reconciliation_automatic_rule_configs'::regclass
      and conname = 'financial_reconciliation_automatic_rule_configs_priority_key'
  ) then
    alter table public.financial_reconciliation_automatic_rule_configs
      add constraint financial_reconciliation_automatic_rule_configs_priority_key
      unique (priority) deferrable initially deferred;
  end if;
end $$;

create table if not exists public.financial_reconciliation_automatic_schedule (
  id boolean primary key default true check (id),
  enabled boolean not null default false,
  time_of_day time without time zone not null default '02:00',
  time_zone text not null default 'Europe/Lisbon' check (time_zone = 'Europe/Lisbon'),
  updated_by text not null default '',
  updated_at timestamptz not null default now()
);

create table if not exists public.financial_reconciliation_automatic_runs (
  id uuid primary key default gen_random_uuid(),
  trigger text not null check (trigger in ('manual','scheduled')),
  scope text not null check (scope in ('rule','batch')),
  status text not null default 'analyzing' check (status in ('analyzing','ready','running','completed','partial','failed')),
  actor text not null,
  client_request_id uuid null,
  scheduled_slot text null check (scheduled_slot is null or scheduled_slot ~ '^\d{4}-\d{2}-\d{2}$'),
  definition_config_snapshot jsonb not null default '[]'::jsonb check (jsonb_typeof(definition_config_snapshot) = 'array'),
  counts jsonb not null default '{}'::jsonb check (jsonb_typeof(counts) = 'object'),
  error_summary text not null default '',
  analysis_completed_at timestamptz null,
  started_at timestamptz not null default now(),
  finished_at timestamptz null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  check (
    (trigger = 'manual' and client_request_id is not null and scheduled_slot is null)
    or
    (trigger = 'scheduled' and scope = 'batch' and client_request_id is null and scheduled_slot is not null)
  ),
  unique (actor, client_request_id)
);

create unique index if not exists financial_reconciliation_automatic_runs_scheduled_slot_uidx
  on public.financial_reconciliation_automatic_runs (scheduled_slot)
  where scheduled_slot is not null;

create table if not exists public.financial_reconciliation_automatic_proposals (
  id uuid primary key default gen_random_uuid(),
  run_id uuid not null references public.financial_reconciliation_automatic_runs(id) on delete cascade,
  rule_key text not null,
  rule_version integer not null,
  base_source_type text not null,
  base_source_id uuid not null,
  base_source_date date not null,
  items jsonb not null default '[]'::jsonb check (jsonb_typeof(items) = 'array'),
  evidence jsonb not null default '[]'::jsonb check (jsonb_typeof(evidence) = 'array'),
  candidate_groups jsonb not null default '[]'::jsonb check (jsonb_typeof(candidate_groups) = 'array'),
  calculated_difference numeric(14,2) not null default 0,
  allowed_difference numeric(14,2) not null check (allowed_difference >= 0),
  status text not null default 'proposed' check (status in ('proposed','ambiguous','deselected','executing','completed','stale','failed')),
  reason text not null default '',
  signature text not null,
  reconciliation_id uuid null references public.financial_reconciliations(id),
  error text not null default '',
  error_detail text not null default '',
  completed_at timestamptz null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  foreign key (rule_key, rule_version) references public.financial_reconciliation_automatic_rule_definitions(rule_key, version),
  unique (run_id, rule_key, base_source_type, base_source_id, signature)
);

alter table public.financial_reconciliations
  add column if not exists origin text not null default 'user' check (origin in ('user','automatic')),
  add column if not exists automatic_trigger text null check (automatic_trigger in ('manual','scheduled')),
  add column if not exists automatic_rule_key text null,
  add column if not exists automatic_rule_version integer null,
  add column if not exists automatic_run_id uuid null references public.financial_reconciliation_automatic_runs(id),
  add column if not exists automatic_proposal_id uuid null references public.financial_reconciliation_automatic_proposals(id);

do $$
begin
  if not exists (
    select 1 from pg_constraint
    where conrelid = 'public.financial_reconciliations'::regclass
      and conname = 'financial_reconciliations_automatic_provenance_check'
  ) then
    alter table public.financial_reconciliations
      add constraint financial_reconciliations_automatic_provenance_check check (
        (
          origin = 'user'
          and automatic_trigger is null
          and automatic_rule_key is null
          and automatic_rule_version is null
          and automatic_run_id is null
          and automatic_proposal_id is null
        )
        or
        (
          origin = 'automatic'
          and automatic_trigger is not null
          and automatic_rule_key is not null
          and automatic_rule_version is not null
          and automatic_run_id is not null
          and automatic_proposal_id is not null
        )
      ) not valid;
  end if;
end $$;

alter table public.financial_reconciliations
  validate constraint financial_reconciliations_automatic_provenance_check;

insert into public.financial_reconciliation_automatic_rule_definitions (
  rule_key, version, display_name, base_source_type, destination_source_types,
  logic_description, definition
) values (
  'financial_documents_cgd_bank_statement',
  1,
  'Financial Documents to CGD Bank Statement',
  'financial_documents',
  '["import_cgd_extrato_ordem"]'::jsonb,
  'A bank candidate must match at least one of three OR identity branches: compact document-number containment, document-description similarity, or supplier-to-bank-description word similarity. A base record is executable only when exactly one complete destination combination is valid; multiple combinations are reported as ambiguous and are never selected automatically.',
  $$
  {
    "baseSourceType": "financial_documents",
    "destinationSourceTypes": ["import_cgd_extrato_ordem"],
    "identityBranches": {
      "document_number": {"algorithm": "compact_containment"},
      "description_similarity": {"algorithm": "similarity"},
      "supplier_similarity": {"algorithm": "word_similarity"}
    },
    "documentNumberMinimumCompactLength": 4,
    "descriptionSimilarityThreshold": 0.60,
    "supplierWordSimilarityThreshold": 0.70,
    "maxDestinationRecords": 4,
    "maxIdentityCandidatesPerBase": 12
  }
  $$::jsonb
)
on conflict (rule_key, version) do update set
  display_name = excluded.display_name,
  base_source_type = excluded.base_source_type,
  destination_source_types = excluded.destination_source_types,
  logic_description = excluded.logic_description,
  definition = excluded.definition;

insert into public.financial_reconciliation_automatic_rule_configs (
  rule_key, rule_version, enabled, allow_manual_execution,
  include_in_scheduled_batch, difference_allowed, max_difference_days, priority
) values (
  'financial_documents_cgd_bank_statement', 1, false, false, false, 0.00, 7, 1
)
on conflict (rule_key) do nothing;

insert into public.financial_reconciliation_automatic_schedule (
  id, enabled, time_of_day, time_zone
) values (
  true, false, '02:00', 'Europe/Lisbon'
)
on conflict (id) do nothing;

alter table public.financial_reconciliation_automatic_rule_definitions enable row level security;
alter table public.financial_reconciliation_automatic_rule_configs enable row level security;
alter table public.financial_reconciliation_automatic_schedule enable row level security;
alter table public.financial_reconciliation_automatic_runs enable row level security;
alter table public.financial_reconciliation_automatic_proposals enable row level security;

revoke all on table public.financial_reconciliation_automatic_rule_definitions from public, anon, authenticated, service_role;
revoke all on table public.financial_reconciliation_automatic_rule_configs from public, anon, authenticated, service_role;
revoke all on table public.financial_reconciliation_automatic_schedule from public, anon, authenticated, service_role;
revoke all on table public.financial_reconciliation_automatic_runs from public, anon, authenticated, service_role;
revoke all on table public.financial_reconciliation_automatic_proposals from public, anon, authenticated, service_role;

create or replace function public.get_financial_reconciliation_automation_settings()
returns jsonb language plpgsql stable security definer set search_path = public, pg_temp as $$
declare
  v_schedule jsonb;
  v_rules jsonb;
  v_last_scheduled_run jsonb;
begin
  select jsonb_build_object(
    'enabled', s.enabled,
    'timeOfDay', to_char(s.time_of_day, 'HH24:MI'),
    'timeZone', s.time_zone,
    'updatedBy', s.updated_by,
    'updatedAt', s.updated_at
  )
  into v_schedule
  from public.financial_reconciliation_automatic_schedule s
  where s.id = true;

  select coalesce(jsonb_agg(jsonb_build_object(
    'ruleKey', d.rule_key,
    'ruleVersion', c.rule_version,
    'displayName', d.display_name,
    'baseSourceType', d.base_source_type,
    'destinationSourceTypes', d.destination_source_types,
    'logicDescription', d.logic_description,
    'definition', d.definition,
    'enabled', c.enabled,
    'allowManualExecution', c.allow_manual_execution,
    'includeInScheduledBatch', c.include_in_scheduled_batch,
    'differenceAllowed', c.difference_allowed,
    'maxDifferenceDays', c.max_difference_days,
    'priority', c.priority,
    'updatedBy', c.updated_by,
    'updatedAt', c.updated_at
  ) order by c.priority, d.rule_key), '[]'::jsonb)
  into v_rules
  from public.financial_reconciliation_automatic_rule_configs c
  join public.financial_reconciliation_automatic_rule_definitions d
    on d.rule_key = c.rule_key and d.version = c.rule_version;

  select jsonb_build_object(
    'id', r.id,
    'trigger', r.trigger,
    'scope', r.scope,
    'status', r.status,
    'scheduledSlot', r.scheduled_slot,
    'counts', r.counts,
    'analysisCompletedAt', r.analysis_completed_at,
    'startedAt', r.started_at,
    'finishedAt', r.finished_at
  )
  into v_last_scheduled_run
  from public.financial_reconciliation_automatic_runs r
  where r.trigger = 'scheduled'
  order by r.started_at desc, r.id desc
  limit 1;

  return jsonb_build_object(
    'schedule', v_schedule,
    'rules', v_rules,
    'lastScheduledRun', v_last_scheduled_run
  );
end $$;

create or replace function public.replace_financial_reconciliation_automation_settings(p_schedule jsonb, p_rules jsonb, p_actor text)
returns jsonb language plpgsql security definer set search_path = public, pg_temp as $$
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
    or coalesce(p_schedule->>'time_of_day', '') !~ '^(?:[01][0-9]|2[0-3]):[0-5][0-9]$'
    or jsonb_typeof(p_schedule->'time_zone') <> 'string'
    or p_schedule->>'time_zone' <> 'Europe/Lisbon' then
    raise exception 'Automatic schedule values are invalid.';
  end if;

  if p_rules is null or jsonb_typeof(p_rules) <> 'array'
    or exists (select 1 from jsonb_array_elements(p_rules) rule where jsonb_typeof(rule) <> 'object') then
    raise exception 'Automatic rules payload must be an array of objects.';
  end if;

  if exists (
    select 1
    from jsonb_array_elements(p_rules) rule
    where (select count(*) from jsonb_object_keys(rule)) <> 8
       or not (rule ?& array[
         'rule_key','rule_version','enabled','allow_manual_execution',
         'include_in_scheduled_batch','difference_allowed','max_difference_days','priority'
       ])
       or exists (
         select 1 from jsonb_object_keys(rule) key
         where key not in (
           'rule_key','rule_version','enabled','allow_manual_execution',
           'include_in_scheduled_batch','difference_allowed','max_difference_days','priority'
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
       or coalesce(rule->>'difference_allowed', '') !~ '^(0|[0-9]+)(\.[0-9]{1,2})?$'
       or (rule->>'difference_allowed')::numeric not between 0 and 999999999999.99
       or jsonb_typeof(rule->'max_difference_days') <> 'number'
       or coalesce(rule->>'max_difference_days', '') !~ '^[0-9]+$'
       or (rule->>'max_difference_days')::numeric not between 0 and 365
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
    where not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_definitions d
      where d.rule_key = rule->>'rule_key'
    )
  ) then
    raise exception 'Automatic rule is invalid.';
  end if;

  if exists (
    select 1
    from jsonb_array_elements(p_rules) rule
    where not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_definitions d
      where d.rule_key = rule->>'rule_key'
        and d.version = (rule->>'rule_version')::integer
    )
  ) then
    raise exception 'Automatic rule version is invalid.';
  end if;

  if jsonb_array_length(p_rules) <> (
      select count(*) from public.financial_reconciliation_automatic_rule_configs
    )
    or exists (
      select 1
      from public.financial_reconciliation_automatic_rule_configs c
      where not exists (
        select 1 from jsonb_array_elements(p_rules) rule
        where rule->>'rule_key' = c.rule_key
      )
    ) then
    raise exception 'Automation settings require every managed rule exactly once.';
  end if;

  lock table public.financial_reconciliation_source_rules in share row exclusive mode;
  lock table public.financial_reconciliation_automatic_rule_configs in share row exclusive mode;
  lock table public.financial_reconciliation_automatic_schedule in share row exclusive mode;

  if exists (
    select 1
    from jsonb_array_elements(p_rules) rule
    join public.financial_reconciliation_automatic_rule_definitions d
      on d.rule_key = rule->>'rule_key'
     and d.version = (rule->>'rule_version')::integer
    cross join lateral jsonb_array_elements_text(d.destination_source_types) destination(source_type)
    where (rule->>'enabled')::boolean
      and not exists (
        select 1
        from public.financial_reconciliation_source_rules source_rule
        where source_rule.base_source_type = d.base_source_type
          and source_rule.matching_source_type = destination.source_type
      )
  ) then
    raise exception 'No directional source rule exists for an enabled automatic rule.';
  end if;

  update public.financial_reconciliation_automatic_rule_configs c
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
  where c.rule_key = input.rule_key;

  update public.financial_reconciliation_automatic_schedule
  set enabled = (p_schedule->>'enabled')::boolean,
      time_of_day = (p_schedule->>'time_of_day')::time,
      time_zone = p_schedule->>'time_zone',
      updated_by = trim(p_actor),
      updated_at = now()
  where id = true;

  return public.get_financial_reconciliation_automation_settings();
end $$;

revoke all on function public.get_financial_reconciliation_automation_settings() from public, anon, authenticated;
revoke all on function public.replace_financial_reconciliation_automation_settings(jsonb,jsonb,text) from public, anon, authenticated;
grant execute on function public.get_financial_reconciliation_automation_settings() to service_role;
grant execute on function public.replace_financial_reconciliation_automation_settings(jsonb,jsonb,text) to service_role;

notify pgrst, 'reload schema';
