do $migration$
declare
  v_pgcrypto_schema text;
  v_unaccent_schema text;
  v_pg_trgm_schema text;
begin
  select n.nspname into v_pgcrypto_schema
  from pg_catalog.pg_extension e
  join pg_catalog.pg_namespace n on n.oid = e.extnamespace
  where e.extname = 'pgcrypto';

  select n.nspname into v_unaccent_schema
  from pg_catalog.pg_extension e
  join pg_catalog.pg_namespace n on n.oid = e.extnamespace
  where e.extname = 'unaccent';

  select n.nspname into v_pg_trgm_schema
  from pg_catalog.pg_extension e
  join pg_catalog.pg_namespace n on n.oid = e.extnamespace
  where e.extname = 'pg_trgm';

  if v_pgcrypto_schema is null or v_unaccent_schema is null or v_pg_trgm_schema is null then
    raise exception 'Required reconciliation extensions pgcrypto, unaccent, and pg_trgm must be installed.';
  end if;

  execute format($definition$
    create or replace function public.financial_reconciliation_extension_unaccent(p_value text)
    returns text
    language sql
    immutable strict
    set search_path = pg_catalog, pg_temp
    as $function$
      select %I.unaccent(p_value)
    $function$
  $definition$, v_unaccent_schema);

  execute format($definition$
    create or replace function public.financial_reconciliation_extension_similarity(p_left text, p_right text)
    returns real
    language sql
    immutable strict
    set search_path = pg_catalog, pg_temp
    as $function$
      select %I.similarity(p_left, p_right)
    $function$
  $definition$, v_pg_trgm_schema);

  execute format($definition$
    create or replace function public.financial_reconciliation_extension_word_similarity(p_left text, p_right text)
    returns real
    language sql
    immutable strict
    set search_path = pg_catalog, pg_temp
    as $function$
      select %I.word_similarity(p_left, p_right)
    $function$
  $definition$, v_pg_trgm_schema);

  execute format($definition$
    create or replace function public.financial_reconciliation_extension_sha256(p_value text)
    returns text
    language sql
    immutable strict
    set search_path = pg_catalog, pg_temp
    as $function$
      select pg_catalog.encode(%I.digest(p_value, 'sha256'::text), 'hex'::text)
    $function$
  $definition$, v_pgcrypto_schema);
end
$migration$;

create or replace function public.financial_reconciliation_match_normalize(p_value text)
returns text
language sql
stable strict
as $$
  with tokens as (
    select token, ordinal
    from regexp_split_to_table(
      btrim(regexp_replace(public.financial_reconciliation_extension_unaccent(lower(p_value)), '[^[:alnum:]]+', ' ', 'g')),
      '[[:space:]]+'
    ) with ordinality as values(token, ordinal)
  )
  select coalesce(string_agg(token, ' ' order by ordinal), '')
  from tokens
  where token ~ '^[0-9]+$' or (token ~ '.*[[:alpha:]].*' and char_length(token) >= 3)
$$;

create or replace function public.financial_reconciliation_match_compact(p_value text)
returns text
language sql
stable strict
as $$
  select regexp_replace(public.financial_reconciliation_extension_unaccent(lower(p_value)), '[^[:alnum:]]', '', 'g')
$$;

create or replace function public.financial_reconciliation_automatic_build_combinations(
  p_base jsonb,
  p_candidates jsonb,
  p_operators jsonb,
  p_tolerance numeric,
  p_max_group_size integer
)
returns table (
  items jsonb,
  calculated_difference numeric,
  signature text
)
language sql
stable
security definer set search_path = public, pg_temp
as $$
  with recursive candidates as (
    select
      row_number() over (order by value->>'sourceType', value->>'sourceId')::integer as candidate_order,
      value->>'sourceType' as source_type,
      value->>'sourceId' as source_id,
      (value->>'sourceDate')::date as source_date,
      round((value->>'amount')::numeric * 100)::bigint as amount_cents,
      case p_operators->>(value->>'sourceType') when '+' then 1 when '-' then -1 else 0 end as multiplier,
      value as item
    from jsonb_array_elements(coalesce(p_candidates, '[]'::jsonb)) value
    where p_operators ? (value->>'sourceType')
  ),
  base as (
    select round((p_base->>'amount')::numeric * 100)::bigint as base_amount_cents,
           round(coalesce(p_tolerance, 0) * 100)::bigint as tolerance_cents
  ),
  subsets as (
    select
      array[c.candidate_order]::integer[] as candidate_orders,
      c.candidate_order as last_order,
      1 as group_size,
      c.amount_cents * c.multiplier as destination_cents,
      jsonb_build_array(c.item) as items,
      jsonb_build_array(jsonb_build_object(
        'sourceType', c.source_type,
        'sourceId', c.source_id,
        'amountCents', c.amount_cents,
        'sourceDate', c.source_date
      )) as signature_items
    from candidates c
    where c.multiplier <> 0
    union all
    select
      s.candidate_orders || c.candidate_order,
      c.candidate_order,
      s.group_size + 1,
      s.destination_cents + (c.amount_cents * c.multiplier),
      s.items || jsonb_build_array(c.item),
      s.signature_items || jsonb_build_array(jsonb_build_object(
        'sourceType', c.source_type,
        'sourceId', c.source_id,
        'amountCents', c.amount_cents,
        'sourceDate', c.source_date
      ))
    from subsets s
    join candidates c on c.candidate_order > s.last_order and c.multiplier <> 0
    where s.group_size < greatest(1, p_max_group_size)
  ),
  qualifying as (
    select s.*, b.base_amount_cents + s.destination_cents as calculated_difference_cents, b.tolerance_cents
    from subsets s cross join base b
  )
  select
    items,
    calculated_difference_cents::numeric / 100,
    public.financial_reconciliation_extension_sha256(signature_items::text) as signature
  from qualifying
  where abs(calculated_difference_cents) <= tolerance_cents
  order by signature
$$;

create or replace function public.financial_reconciliation_automatic_rule_candidates(
  p_rule_key text,
  p_rule_version integer,
  p_difference_allowed numeric,
  p_max_difference_days integer
)
returns table (
  base_source_id uuid,
  base_source_date date,
  base_snapshot jsonb,
  candidates jsonb,
  candidate_count integer
)
language sql
stable
security definer set search_path = public, pg_temp
as $$
  with bases as (
    select d.id, d.document_date, d.doc_number, d.description, d.supplier_name, d.amount,
           public.financial_reconciliation_match_compact(d.doc_number) as compact_document_number,
           public.financial_reconciliation_match_normalize(d.description) as normalized_document_description,
           public.financial_reconciliation_match_normalize(d.supplier_name) as normalized_supplier_name
    from public.financial_documents d
    where p_rule_key = 'financial_documents_cgd_bank_statement'
      and p_rule_version = 1
      and d.fat = 'S'
      and d.document_date >= date '2026-01-01'
      and not exists (
        select 1 from public.financial_reconciliation_items i
        where i.source_type = 'financial_documents' and i.source_id = d.id
      )
  ),
  qualified as (
    select
      d.id as base_id,
      d.document_date as base_date,
      jsonb_build_object('sourceType', 'financial_documents', 'sourceId', d.id, 'sourceDate', d.document_date,
        'amount', d.amount, 'docNumber', d.doc_number, 'description', d.description, 'supplierName', d.supplier_name) as base_snapshot,
      b.id as source_id,
      b.data as source_date,
      b.montante as amount,
      b.descritivo as description,
      public.financial_reconciliation_match_normalize(b.descritivo) as normalized_bank_description,
      d.compact_document_number,
      d.normalized_document_description,
      d.normalized_supplier_name
    from bases d
    left join public.import_cgd_extrato_ordem b
      on b.data between d.document_date - p_max_difference_days and d.document_date + p_max_difference_days
     and b.data >= date '2026-01-01'
     and b.montante is not null
     and not exists (
       select 1 from public.financial_reconciliation_items i
       where i.source_type = 'import_cgd_extrato_ordem' and i.source_id = b.id
     )
  ),
  scored as (
    select q.*,
      coalesce(char_length(q.compact_document_number) >= 4
        and q.source_id is not null
        and position(q.compact_document_number in public.financial_reconciliation_match_compact(q.description)) > 0, false) as document_number_matched,
      case when nullif(q.normalized_document_description, '') is null or nullif(q.normalized_bank_description, '') is null then 0::real
        else public.financial_reconciliation_extension_similarity(normalized_document_description, normalized_bank_description) end as description_score,
      case when nullif(q.normalized_supplier_name, '') is null or nullif(q.normalized_bank_description, '') is null then 0::real
        else public.financial_reconciliation_extension_word_similarity(normalized_supplier_name, normalized_bank_description) end as supplier_score
    from qualified q
  ),
  identity_candidates as (
    select *,
      document_number_matched
      or description_score >= 0.60
      or supplier_score >= 0.70 as identity_matched
    from scored
  ),
  grouped as (
    select
      base_id, base_date, base_snapshot,
      coalesce(jsonb_agg(jsonb_build_object(
        'sourceType', 'import_cgd_extrato_ordem', 'sourceId', source_id, 'sourceDate', source_date,
        'amount', amount, 'description', description,
        'evidence', jsonb_build_object(
          'documentNumber', jsonb_build_object('matched', document_number_matched, 'normalized', compact_document_number),
          'description', jsonb_build_object('matched', description_score >= 0.60, 'score', description_score, 'threshold', 0.60),
          'supplier', jsonb_build_object('matched', supplier_score >= 0.70, 'score', supplier_score, 'threshold', 0.70)
        )
      ) order by source_date, source_id) filter (where identity_matched), '[]'::jsonb) as candidates,
      count(*) filter (where identity_matched)::integer as candidate_count
    from identity_candidates
    group by base_id, base_date, base_snapshot
  )
  select base_id, base_date, base_snapshot, candidates, candidate_count
  from grouped
  order by base_date, base_id
$$;

create or replace function public.get_financial_reconciliation_automatic_run(p_run_id uuid)
returns jsonb
language plpgsql
stable
security definer set search_path = public, pg_temp
as $$
declare v_run public.financial_reconciliation_automatic_runs%rowtype; v_proposals jsonb;
begin
  select * into v_run from public.financial_reconciliation_automatic_runs where id = p_run_id;
  if not found then raise exception 'Automatic analysis run was not found.'; end if;
  select coalesce(jsonb_agg(jsonb_build_object(
    'id', p.id, 'runId', p.run_id, 'ruleKey', p.rule_key, 'ruleVersion', p.rule_version,
    'baseSourceType', p.base_source_type, 'baseSourceId', p.base_source_id, 'baseSourceDate', p.base_source_date,
    'baseSnapshot', p.base_snapshot,
    'items', p.items, 'evidence', p.evidence, 'candidateGroups', p.candidate_groups,
    'calculatedDifference', p.calculated_difference, 'allowedDifference', p.allowed_difference,
    'status', p.status, 'reason', p.reason, 'signature', p.signature, 'reconciliationId', p.reconciliation_id,
    'createdAt', p.created_at, 'updatedAt', p.updated_at
  ) order by p.base_source_date, p.base_source_id, p.signature), '[]'::jsonb)
  into v_proposals from public.financial_reconciliation_automatic_proposals p where p.run_id = v_run.id;
  return jsonb_build_object(
    'runId', v_run.id, 'trigger', v_run.trigger, 'scope', v_run.scope, 'status', v_run.status,
    'actor', v_run.actor, 'clientRequestId', v_run.client_request_id, 'scheduledSlot', v_run.scheduled_slot,
    'definitions', v_run.definition_config_snapshot, 'counts', v_run.counts,
    'analysisCompletedAt', v_run.analysis_completed_at, 'startedAt', v_run.started_at, 'finishedAt', v_run.finished_at,
    'proposals', v_proposals
  );
end $$;

create or replace function public.populate_financial_reconciliation_automatic_run(p_run_id uuid)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
declare
  v_run public.financial_reconciliation_automatic_runs%rowtype;
  v_rule jsonb; v_base record; v_combination record; v_operator text;
  v_total integer := 0; v_combination_count integer := 0; v_proposed integer := 0; v_ambiguous integer := 0;
begin
  select * into v_run from public.financial_reconciliation_automatic_runs where id = p_run_id for update;
  if not found then raise exception 'Automatic analysis run was not found.'; end if;
  if v_run.analysis_completed_at is not null then return public.get_financial_reconciliation_automatic_run(p_run_id); end if;

  for v_rule in select value from jsonb_array_elements(v_run.definition_config_snapshot) value order by (value->>'priority')::integer, value->>'ruleKey' loop
    v_operator := v_rule->>'operator';
    if v_operator not in ('+', '-') then raise exception 'Automatic rule has no directional source operator.'; end if;
    for v_base in
      select * from public.financial_reconciliation_automatic_rule_candidates(
        v_rule->>'ruleKey', (v_rule->>'ruleVersion')::integer,
        (v_rule->>'differenceAllowed')::numeric, (v_rule->>'maxDifferenceDays')::integer
      )
    loop
      v_total := v_total + 1;
      if v_base.candidate_count > 12 then
        insert into public.financial_reconciliation_automatic_proposals (
          run_id, rule_key, rule_version, base_source_type, base_source_id, base_source_date,
          base_snapshot, candidate_groups, allowed_difference, status, reason, signature
        ) values (
          v_run.id, v_rule->>'ruleKey', (v_rule->>'ruleVersion')::integer, 'financial_documents',
          v_base.base_source_id, v_base.base_source_date, v_base.base_snapshot, v_base.candidates,
          (v_rule->>'differenceAllowed')::numeric, 'ambiguous', 'candidate_limit',
          public.financial_reconciliation_extension_sha256('candidate_limit:' || v_base.base_source_id::text)
        ) on conflict do nothing;
        v_ambiguous := v_ambiguous + 1;
      else
        select count(*) into strict v_combination_count from public.financial_reconciliation_automatic_build_combinations(
          v_base.base_snapshot, v_base.candidates,
          jsonb_build_object('import_cgd_extrato_ordem', v_operator),
          (v_rule->>'differenceAllowed')::numeric, 4
        );
        if v_combination_count = 1 then
          select * into strict v_combination from public.financial_reconciliation_automatic_build_combinations(
            v_base.base_snapshot, v_base.candidates, jsonb_build_object('import_cgd_extrato_ordem', v_operator),
            (v_rule->>'differenceAllowed')::numeric, 4
          );
          insert into public.financial_reconciliation_automatic_proposals (
            run_id, rule_key, rule_version, base_source_type, base_source_id, base_source_date,
            base_snapshot, items, evidence, candidate_groups, calculated_difference, allowed_difference, status, signature
          ) values (
            v_run.id, v_rule->>'ruleKey', (v_rule->>'ruleVersion')::integer, 'financial_documents',
            v_base.base_source_id, v_base.base_source_date, v_base.base_snapshot, v_combination.items,
            (select coalesce(jsonb_agg(value->'evidence'), '[]'::jsonb) from jsonb_array_elements(v_combination.items) value),
            jsonb_build_array(v_combination.items), v_combination.calculated_difference,
            (v_rule->>'differenceAllowed')::numeric, 'proposed', v_combination.signature
          ) on conflict do nothing;
          v_proposed := v_proposed + 1;
        elsif v_combination_count > 1 then
          insert into public.financial_reconciliation_automatic_proposals (
            run_id, rule_key, rule_version, base_source_type, base_source_id, base_source_date,
            base_snapshot, candidate_groups, allowed_difference, status, reason, signature
          ) values (
            v_run.id, v_rule->>'ruleKey', (v_rule->>'ruleVersion')::integer, 'financial_documents',
            v_base.base_source_id, v_base.base_source_date, v_base.base_snapshot,
            (select coalesce(jsonb_agg(items order by signature), '[]'::jsonb) from public.financial_reconciliation_automatic_build_combinations(
              v_base.base_snapshot, v_base.candidates, jsonb_build_object('import_cgd_extrato_ordem', v_operator),
              (v_rule->>'differenceAllowed')::numeric, 4
            )), (v_rule->>'differenceAllowed')::numeric, 'ambiguous', 'multiple_combinations',
            public.financial_reconciliation_extension_sha256('multiple:' || v_base.base_source_id::text)
          ) on conflict do nothing;
          v_ambiguous := v_ambiguous + 1;
        else
          insert into public.financial_reconciliation_automatic_proposals (
            run_id, rule_key, rule_version, base_source_type, base_source_id, base_source_date,
            base_snapshot, candidate_groups, allowed_difference, status, reason, signature
          ) values (
            v_run.id, v_rule->>'ruleKey', (v_rule->>'ruleVersion')::integer, 'financial_documents',
            v_base.base_source_id, v_base.base_source_date, v_base.base_snapshot, v_base.candidates,
            (v_rule->>'differenceAllowed')::numeric, 'skipped', 'no_qualifying_combination',
            public.financial_reconciliation_extension_sha256('skipped:' || v_base.base_source_id::text)
          ) on conflict do nothing;
        end if;
      end if;
    end loop;
  end loop;

  with source_usage as (
    select item->>'sourceType' as source_type, item->>'sourceId' as source_id,
           count(distinct p.base_source_id) as base_count
    from public.financial_reconciliation_automatic_proposals p
    join lateral (
      select item.value as item from jsonb_array_elements(p.items) as item(value)
      union all
      select item.value as item
      from jsonb_array_elements(p.candidate_groups) as candidate_group(value)
      join lateral jsonb_array_elements(
        case when jsonb_typeof(candidate_group.value) = 'array' then candidate_group.value else jsonb_build_array(candidate_group.value) end
      ) as item(value) on true
    ) source_item on true
    where p.run_id = v_run.id and p.status in ('proposed', 'ambiguous')
    group by item->>'sourceType', item->>'sourceId'
  ), overlapping as (
    select distinct p.id
    from public.financial_reconciliation_automatic_proposals p
    join lateral (
      select item.value as item from jsonb_array_elements(p.items) as item(value)
      union all
      select item.value as item
      from jsonb_array_elements(p.candidate_groups) as candidate_group(value)
      join lateral jsonb_array_elements(
        case when jsonb_typeof(candidate_group.value) = 'array' then candidate_group.value else jsonb_build_array(candidate_group.value) end
      ) as item(value) on true
    ) source_item on true
    join source_usage u on u.source_type = item->>'sourceType' and u.source_id = item->>'sourceId'
    where p.run_id = v_run.id and p.status in ('proposed', 'ambiguous') and u.base_count > 1
  )
  update public.financial_reconciliation_automatic_proposals p
  set status = 'ambiguous', reason = 'cross_base_overlap', updated_at = now()
  where p.id in (select id from overlapping);

  update public.financial_reconciliation_automatic_runs
  set status = 'ready', analysis_completed_at = now(), updated_at = now(),
      counts = (select jsonb_build_object(
        'bases', count(distinct base_source_id),
        'proposed', count(*) filter (where status = 'proposed'),
        'ambiguous', count(*) filter (where status = 'ambiguous'),
        'skipped', count(*) filter (where status = 'skipped')
      ) from public.financial_reconciliation_automatic_proposals where run_id = v_run.id)
  where id = v_run.id;
  return public.get_financial_reconciliation_automatic_run(v_run.id);
end $$;

create or replace function public.create_financial_reconciliation_automatic_analysis(
  p_rule_keys text[], p_mode text, p_actor text, p_client_request_id uuid
)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
declare v_snapshot jsonb; v_run_id uuid; v_scope text;
begin
  if nullif(trim(coalesce(p_actor, '')), '') is null then raise exception 'Actor is required.'; end if;
  if p_client_request_id is null then raise exception 'Client request ID is required.'; end if;
  if p_mode = 'manual_rule' then v_scope := 'rule';
  elsif p_mode = 'manual_batch' then v_scope := 'batch';
  else raise exception 'Automatic analysis mode is invalid.'; end if;
  if p_rule_keys is null or cardinality(p_rule_keys) = 0 or cardinality(p_rule_keys) <> cardinality(array(select distinct key from unnest(p_rule_keys) key)) then
    raise exception 'Automatic rule selection is invalid.';
  end if;
  lock table public.financial_reconciliation_source_rules in share row exclusive mode;
  lock table public.financial_reconciliation_automatic_rule_configs in share row exclusive mode;
  select coalesce(jsonb_agg(jsonb_build_object(
    'ruleKey', c.rule_key, 'ruleVersion', c.rule_version, 'displayName', d.display_name, 'priority', c.priority,
    'differenceAllowed', c.difference_allowed, 'maxDifferenceDays', c.max_difference_days,
    'definition', d.definition, 'operator', sr.operator
  ) order by c.priority, c.rule_key), '[]'::jsonb)
  into v_snapshot
  from public.financial_reconciliation_automatic_rule_configs c
  join public.financial_reconciliation_automatic_rule_definitions d on d.rule_key = c.rule_key and d.version = c.rule_version
  join public.financial_reconciliation_source_rules sr on sr.base_source_type = d.base_source_type
    and sr.matching_source_type = 'import_cgd_extrato_ordem'
  where c.rule_key = any(p_rule_keys) and c.enabled
    and ((p_mode = 'manual_rule' and c.allow_manual_execution) or (p_mode = 'manual_batch' and c.include_in_scheduled_batch));
  if jsonb_array_length(v_snapshot) <> cardinality(p_rule_keys) then raise exception 'Automatic rule is not enabled for requested analysis mode.'; end if;
  insert into public.financial_reconciliation_automatic_runs (trigger, scope, actor, client_request_id, definition_config_snapshot)
  values ('manual', v_scope, p_actor, p_client_request_id, v_snapshot)
  on conflict (actor, client_request_id) do nothing returning id into v_run_id;
  if v_run_id is null then select id into strict v_run_id from public.financial_reconciliation_automatic_runs where actor = p_actor and client_request_id = p_client_request_id; end if;
  return public.populate_financial_reconciliation_automatic_run(v_run_id);
end $$;

create or replace function public.claim_financial_reconciliation_automatic_schedule(p_now timestamptz, p_actor text)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
declare v_schedule public.financial_reconciliation_automatic_schedule%rowtype; v_local timestamp; v_slot text; v_snapshot jsonb; v_run_id uuid;
begin
  if nullif(trim(coalesce(p_actor, '')), '') is null then raise exception 'Actor is required.'; end if;
  select * into strict v_schedule from public.financial_reconciliation_automatic_schedule where id = true for update;
  v_local := p_now at time zone 'Europe/Lisbon'; v_slot := to_char(v_local::date, 'YYYY-MM-DD');
  if not v_schedule.enabled then return jsonb_build_object('claimed', false, 'reason', 'schedule_disabled'); end if;
  select id into v_run_id from public.financial_reconciliation_automatic_runs
  where trigger = 'scheduled' and finished_at is null
  order by scheduled_slot, started_at for update;
  if found then return jsonb_build_object('claimed', true, 'resumed', true, 'run', public.get_financial_reconciliation_automatic_run(v_run_id)); end if;
  if v_local::time < v_schedule.time_of_day then return jsonb_build_object('claimed', false, 'reason', 'before_scheduled_time'); end if;
  select coalesce(jsonb_agg(jsonb_build_object(
    'ruleKey', c.rule_key, 'ruleVersion', c.rule_version, 'displayName', d.display_name, 'priority', c.priority,
    'differenceAllowed', c.difference_allowed, 'maxDifferenceDays', c.max_difference_days,
    'definition', d.definition, 'operator', sr.operator
  ) order by c.priority, c.rule_key), '[]'::jsonb)
  into v_snapshot
  from public.financial_reconciliation_automatic_rule_configs c
  join public.financial_reconciliation_automatic_rule_definitions d on d.rule_key = c.rule_key and d.version = c.rule_version
  join public.financial_reconciliation_source_rules sr on sr.base_source_type = d.base_source_type and sr.matching_source_type = 'import_cgd_extrato_ordem'
  where c.enabled and c.include_in_scheduled_batch;
  if jsonb_array_length(v_snapshot) = 0 then return jsonb_build_object('claimed', false, 'reason', 'no_enabled_rules'); end if;
  insert into public.financial_reconciliation_automatic_runs (trigger, scope, actor, scheduled_slot, definition_config_snapshot)
  values ('scheduled', 'batch', p_actor, v_slot, v_snapshot)
  on conflict do nothing returning id into v_run_id;
  if v_run_id is null then
    select id into strict v_run_id from public.financial_reconciliation_automatic_runs where scheduled_slot = v_slot;
    return jsonb_build_object('claimed', true, 'resumed', true, 'run', public.get_financial_reconciliation_automatic_run(v_run_id));
  end if;
  return jsonb_build_object('claimed', true, 'resumed', false, 'run', public.get_financial_reconciliation_automatic_run(v_run_id));
end $$;

revoke all on function public.financial_reconciliation_extension_unaccent(text) from public, anon, authenticated;
revoke all on function public.financial_reconciliation_extension_similarity(text,text) from public, anon, authenticated;
revoke all on function public.financial_reconciliation_extension_word_similarity(text,text) from public, anon, authenticated;
revoke all on function public.financial_reconciliation_extension_sha256(text) from public, anon, authenticated;
revoke all on function public.financial_reconciliation_match_normalize(text) from public, anon, authenticated;
revoke all on function public.financial_reconciliation_match_compact(text) from public, anon, authenticated;
revoke all on function public.financial_reconciliation_automatic_build_combinations(jsonb,jsonb,jsonb,numeric,integer) from public, anon, authenticated;
revoke all on function public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer) from public, anon, authenticated;
revoke all on function public.populate_financial_reconciliation_automatic_run(uuid) from public, anon, authenticated;
revoke all on function public.create_financial_reconciliation_automatic_analysis(text[],text,text,uuid) from public, anon, authenticated;
revoke all on function public.claim_financial_reconciliation_automatic_schedule(timestamptz,text) from public, anon, authenticated;
revoke all on function public.get_financial_reconciliation_automatic_run(uuid) from public, anon, authenticated;
grant execute on function public.financial_reconciliation_extension_unaccent(text) to service_role;
grant execute on function public.financial_reconciliation_extension_similarity(text,text) to service_role;
grant execute on function public.financial_reconciliation_extension_word_similarity(text,text) to service_role;
grant execute on function public.financial_reconciliation_extension_sha256(text) to service_role;
grant execute on function public.financial_reconciliation_match_normalize(text) to service_role;
grant execute on function public.financial_reconciliation_match_compact(text) to service_role;
grant execute on function public.financial_reconciliation_automatic_build_combinations(jsonb,jsonb,jsonb,numeric,integer) to service_role;
grant execute on function public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer) to service_role;
grant execute on function public.populate_financial_reconciliation_automatic_run(uuid) to service_role;
grant execute on function public.create_financial_reconciliation_automatic_analysis(text[],text,text,uuid) to service_role;
grant execute on function public.claim_financial_reconciliation_automatic_schedule(timestamptz,text) to service_role;
grant execute on function public.get_financial_reconciliation_automatic_run(uuid) to service_role;
notify pgrst, 'reload schema';
