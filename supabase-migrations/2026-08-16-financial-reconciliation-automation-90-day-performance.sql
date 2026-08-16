-- Automatic reconciliation: indexed CGD search and resumable analysis pages.
-- Apply after 2026-08-16-financial-reconciliation-automation-banco-v2.sql.

do $migration$
declare v_had_cursor boolean;
begin
  select exists (
    select 1 from pg_attribute
    where attrelid = 'public.financial_reconciliation_automatic_runs'::regclass
      and attname = 'analysis_cursor_date' and not attisdropped
  ) into v_had_cursor;

  alter table public.financial_reconciliation_automatic_runs
    add column if not exists analysis_cursor_date date,
    add column if not exists analysis_cursor_id uuid,
    add column if not exists analysis_processed integer not null default 0,
    add column if not exists analysis_total integer not null default 0,
    add column if not exists analysis_error_code text,
    add column if not exists analysis_error_at timestamptz;

  if not v_had_cursor then
    execute $sql$
      update public.financial_reconciliation_automatic_runs
      set status = 'failed',
          error_summary = 'Analysis must be restarted after the 90-day performance upgrade.',
          analysis_error_code = 'analysis_upgrade_restart_required',
          analysis_error_at = now(), finished_at = coalesce(finished_at, now()), updated_at = now()
      where status = 'analyzing' and analysis_completed_at is null
    $sql$;
  end if;
end
$migration$;

update public.financial_reconciliation_automatic_rule_configs
set max_difference_days = least(max_difference_days, 90),
    updated_at = now()
where max_difference_days > 90;

do $migration$
declare v_constraint record;
begin
  for v_constraint in
    select conname
    from pg_constraint
    where conrelid = 'public.financial_reconciliation_automatic_rule_configs'::regclass
      and contype = 'c'
      and pg_get_constraintdef(oid) ilike '%max_difference_days%'
  loop
    execute format(
      'alter table public.financial_reconciliation_automatic_rule_configs drop constraint %I',
      v_constraint.conname
    );
  end loop;
  alter table public.financial_reconciliation_automatic_rule_configs
    add constraint financial_reconciliation_automatic_rule_configs_max_days_check
    check (max_difference_days between 0 and 90);
end
$migration$;

create table if not exists public.financial_reconciliation_cgd_match_search (
  source_id uuid primary key references public.import_cgd_extrato_ordem(id) on delete cascade,
  source_date date not null,
  amount numeric,
  description text,
  normalized_description text not null,
  compact_description text not null,
  updated_at timestamptz not null default now()
);

create index if not exists financial_reconciliation_cgd_match_search_date_id_idx
  on public.financial_reconciliation_cgd_match_search(source_date, source_id);

do $migration$
declare v_trgm_schema text;
begin
  select n.nspname into strict v_trgm_schema
  from pg_extension e join pg_namespace n on n.oid = e.extnamespace
  where e.extname = 'pg_trgm';

  execute format(
    'create index if not exists financial_reconciliation_cgd_match_search_normalized_trgm_idx
       on public.financial_reconciliation_cgd_match_search
       using gin (normalized_description %I.gin_trgm_ops)',
    v_trgm_schema
  );
  execute format(
    'create index if not exists financial_reconciliation_cgd_match_search_compact_trgm_idx
       on public.financial_reconciliation_cgd_match_search
       using gin (compact_description %I.gin_trgm_ops)',
    v_trgm_schema
  );
end
$migration$;

create or replace function public.financial_reconciliation_sync_cgd_match_search()
returns trigger
language plpgsql
security definer set search_path = public, pg_temp
as $$
begin
  if tg_op = 'DELETE' then
    delete from public.financial_reconciliation_cgd_match_search where source_id = old.id;
    return old;
  end if;

  if tg_op = 'UPDATE' and old.id is distinct from new.id then
    delete from public.financial_reconciliation_cgd_match_search where source_id = old.id;
  end if;

  insert into public.financial_reconciliation_cgd_match_search (
    source_id, source_date, amount, description,
    normalized_description, compact_description, updated_at
  ) values (
    new.id, new.data, new.montante, new.descritivo,
    coalesce(public.financial_reconciliation_match_normalize(new.descritivo), ''),
    coalesce(public.financial_reconciliation_match_compact(new.descritivo), ''),
    now()
  )
  on conflict (source_id) do update set
    source_date = excluded.source_date,
    amount = excluded.amount,
    description = excluded.description,
    normalized_description = excluded.normalized_description,
    compact_description = excluded.compact_description,
    updated_at = excluded.updated_at;
  return new;
end
$$;

drop trigger if exists financial_reconciliation_sync_cgd_match_search_trigger
  on public.import_cgd_extrato_ordem;
create trigger financial_reconciliation_sync_cgd_match_search_trigger
after insert or update or delete on public.import_cgd_extrato_ordem
for each row execute function public.financial_reconciliation_sync_cgd_match_search();

insert into public.financial_reconciliation_cgd_match_search (
  source_id, source_date, amount, description,
  normalized_description, compact_description, updated_at
)
select
  bank.id, bank.data, bank.montante, bank.descritivo,
  coalesce(public.financial_reconciliation_match_normalize(bank.descritivo), ''),
  coalesce(public.financial_reconciliation_match_compact(bank.descritivo), ''),
  now()
from public.import_cgd_extrato_ordem bank
on conflict (source_id) do update set
  source_date = excluded.source_date,
  amount = excluded.amount,
  description = excluded.description,
  normalized_description = excluded.normalized_description,
  compact_description = excluded.compact_description,
  updated_at = excluded.updated_at;

do $migration$
declare v_trgm_schema text; v_sql text;
begin
  select n.nspname into strict v_trgm_schema
  from pg_extension e join pg_namespace n on n.oid = e.extnamespace
  where e.extname = 'pg_trgm';

  v_sql := format($function$
create or replace function public.financial_reconciliation_automatic_candidates_for_base_ids(
  p_rule_key text,
  p_rule_version integer,
  p_difference_allowed numeric,
  p_max_difference_days integer,
  p_base_ids uuid[]
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
security definer set search_path = public, %1$I, pg_temp
set pg_trgm.similarity_threshold = 0.60
set pg_trgm.word_similarity_threshold = 0.70
as $body$
  with bases as materialized (
    select
      d.id, d.document_date, d.doc_number, d.description, d.supplier_name, d.amount,
      public.financial_reconciliation_match_compact(d.doc_number) as compact_document_number,
      public.financial_reconciliation_match_normalize(d.description) as normalized_document_description,
      public.financial_reconciliation_match_normalize(d.supplier_name) as normalized_supplier_name
    from public.financial_documents d
    where p_rule_key = 'financial_documents_cgd_bank_statement'
      and p_rule_version = 2
      and p_max_difference_days between 0 and 90
      and d.id = any(coalesce(p_base_ids, array[]::uuid[]))
      and d.fat = 'S'
      and d.payment = 'Banco'
      and d.document_date >= date '2026-01-01'
      and not exists (
        select 1 from public.financial_reconciliation_items i
        where i.source_type = 'financial_documents' and i.source_id = d.id
      )
  ), qualified as materialized (
    select
      d.id as base_id,
      d.document_date as base_date,
      jsonb_build_object(
        'sourceType', 'financial_documents', 'sourceId', d.id,
        'sourceDate', d.document_date, 'amount', d.amount,
        'docNumber', d.doc_number, 'description', d.description,
        'supplierName', d.supplier_name
      ) as base_snapshot,
      search.source_id, search.source_date, search.amount, search.description,
      search.normalized_description,
      d.compact_document_number,
      d.normalized_document_description,
      d.normalized_supplier_name
    from bases d
    left join lateral (
      select s.*
      from public.financial_reconciliation_cgd_match_search s
      where s.source_date between d.document_date - p_max_difference_days
                              and d.document_date + p_max_difference_days
        and s.source_date >= date '2026-01-01'
        and s.amount is not null
        and not exists (
          select 1 from public.financial_reconciliation_items i
          where i.source_type = 'import_cgd_extrato_ordem' and i.source_id = s.source_id
        )
        and (
          (char_length(d.compact_document_number) >= 4
            and s.compact_description like '%%' || d.compact_document_number || '%%')
          or s.normalized_description OPERATOR(%1$I.%%) d.normalized_document_description
          or d.normalized_supplier_name OPERATOR(%1$I.<%%) s.normalized_description
        )
      order by s.source_date, s.source_id
    ) search on true
  ), scored as materialized (
    select
      q.*,
      coalesce(
        char_length(q.compact_document_number) >= 4
          and q.source_id is not null
          and position(q.compact_document_number in public.financial_reconciliation_match_compact(q.description)) > 0,
        false
      ) as document_number_matched,
      case
        when nullif(q.normalized_document_description, '') is null
          or nullif(q.normalized_description, '') is null then 0::real
        else public.financial_reconciliation_extension_similarity(
          q.normalized_document_description, q.normalized_description
        )
      end as description_score,
      case
        when nullif(q.normalized_supplier_name, '') is null
          or nullif(q.normalized_description, '') is null then 0::real
        else public.financial_reconciliation_extension_word_similarity(
          q.normalized_supplier_name, q.normalized_description
        )
      end as supplier_score
    from qualified q
  ), grouped as (
    select
      base_id, base_date, base_snapshot,
      coalesce(jsonb_agg(jsonb_build_object(
        'sourceType', 'import_cgd_extrato_ordem',
        'sourceId', source_id, 'sourceDate', source_date,
        'amount', amount, 'description', description,
        'evidence', jsonb_build_object(
          'documentNumber', jsonb_build_object(
            'matched', document_number_matched, 'normalized', compact_document_number
          ),
          'description', jsonb_build_object(
            'matched', description_score >= 0.60,
            'score', description_score, 'threshold', 0.60
          ),
          'supplier', jsonb_build_object(
            'matched', supplier_score >= 0.70,
            'score', supplier_score, 'threshold', 0.70
          )
        )
      ) order by source_date, source_id) filter (
        where source_id is not null and (
          document_number_matched or description_score >= 0.60 or supplier_score >= 0.70
        )
      ), '[]'::jsonb) as candidates,
      count(*) filter (
        where source_id is not null and (
          document_number_matched or description_score >= 0.60 or supplier_score >= 0.70
        )
      )::integer as candidate_count
    from scored
    group by base_id, base_date, base_snapshot
  )
  select base_id, base_date, base_snapshot, candidates, candidate_count
  from grouped order by base_date, base_id
$body$;
$function$, v_trgm_schema);
  execute v_sql;
end
$migration$;

create or replace function public.financial_reconciliation_automatic_candidate_page(
  p_rule_key text,
  p_rule_version integer,
  p_difference_allowed numeric,
  p_max_difference_days integer,
  p_after_date date,
  p_after_id uuid,
  p_page_size integer default 25
)
returns table (
  base_source_id uuid,
  base_source_date date,
  base_snapshot jsonb,
  candidates jsonb,
  candidate_count integer
)
language plpgsql
stable
security definer set search_path = public, pg_temp
as $$
declare v_base_ids uuid[];
begin
  if p_rule_key <> 'financial_documents_cgd_bank_statement' or p_rule_version <> 2 then
    raise exception 'Automatic reconciliation rule is unsupported.';
  end if;
  if p_max_difference_days not between 0 and 90 then
    raise exception 'Max difference in days must be between 0 and 90.';
  end if;
  if p_page_size not between 1 and 25 then
    raise exception 'Automatic analysis page size must be between 1 and 25.';
  end if;

  select coalesce(array_agg(page.id order by page.document_date, page.id), array[]::uuid[])
  into v_base_ids
  from (
    select d.id, d.document_date
    from public.financial_documents d
    where d.fat = 'S'
      and d.payment = 'Banco'
      and d.document_date >= date '2026-01-01'
      and (p_after_date is null or (d.document_date, d.id) > (p_after_date, p_after_id))
      and not exists (
        select 1 from public.financial_reconciliation_items i
        where i.source_type = 'financial_documents' and i.source_id = d.id
      )
    order by d.document_date, d.id
    limit p_page_size
  ) page;

  return query
  select * from public.financial_reconciliation_automatic_candidates_for_base_ids(
    p_rule_key, p_rule_version, p_difference_allowed, p_max_difference_days, v_base_ids
  );
end
$$;

create or replace function public.financial_reconciliation_automatic_single_base_candidates(
  p_rule_key text,
  p_rule_version integer,
  p_difference_allowed numeric,
  p_max_difference_days integer,
  p_base_source_id uuid
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
  select * from public.financial_reconciliation_automatic_candidates_for_base_ids(
    p_rule_key, p_rule_version, p_difference_allowed, p_max_difference_days,
    array[p_base_source_id]
  )
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
  select * from public.financial_reconciliation_automatic_candidates_for_base_ids(
    p_rule_key, p_rule_version, p_difference_allowed, p_max_difference_days,
    array(
      select d.id from public.financial_documents d
      where d.fat = 'S' and d.payment = 'Banco'
        and d.document_date >= date '2026-01-01'
        and not exists (
          select 1 from public.financial_reconciliation_items i
          where i.source_type = 'financial_documents' and i.source_id = d.id
        )
      order by d.document_date, d.id
    )
  )
$$;

create or replace function public.get_financial_reconciliation_automatic_run(p_run_id uuid)
returns jsonb
language plpgsql
stable
security definer set search_path = public, pg_temp
as $$
declare v_run public.financial_reconciliation_automatic_runs%rowtype; v_proposals jsonb;
begin
  select * into v_run
  from public.financial_reconciliation_automatic_runs
  where id = p_run_id;
  if not found then raise exception 'Automatic analysis run was not found.'; end if;

  select coalesce(jsonb_agg(jsonb_build_object(
    'id', p.id, 'runId', p.run_id, 'ruleKey', p.rule_key,
    'ruleVersion', p.rule_version, 'baseSourceType', p.base_source_type,
    'baseSourceId', p.base_source_id, 'baseSourceDate', p.base_source_date,
    'baseSnapshot', p.base_snapshot, 'items', p.items,
    'evidence', p.evidence, 'candidateGroups', p.candidate_groups,
    'calculatedDifference', p.calculated_difference,
    'allowedDifference', p.allowed_difference, 'status', p.status,
    'reason', p.reason, 'signature', p.signature,
    'reconciliationId', p.reconciliation_id,
    'createdAt', p.created_at, 'updatedAt', p.updated_at
  ) order by p.base_source_date, p.base_source_id, p.signature), '[]'::jsonb)
  into v_proposals
  from public.financial_reconciliation_automatic_proposals p
  where p.run_id = v_run.id;

  return jsonb_build_object(
    'runId', v_run.id, 'trigger', v_run.trigger, 'scope', v_run.scope,
    'status', v_run.status, 'actor', v_run.actor,
    'clientRequestId', v_run.client_request_id,
    'scheduledSlot', v_run.scheduled_slot,
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
    'startedAt', v_run.started_at, 'finishedAt', v_run.finished_at,
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
declare v_run public.financial_reconciliation_automatic_runs%rowtype;
begin
  select * into v_run
  from public.financial_reconciliation_automatic_runs
  where id = p_run_id;
  if not found then raise exception 'Automatic analysis run was not found.'; end if;
  if v_run.analysis_completed_at is not null then
    return public.get_financial_reconciliation_automatic_run(p_run_id);
  end if;
  return jsonb_build_object(
    'runId', v_run.id, 'trigger', v_run.trigger, 'scope', v_run.scope,
    'status', v_run.status, 'actor', v_run.actor,
    'clientRequestId', v_run.client_request_id,
    'scheduledSlot', v_run.scheduled_slot,
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
    'startedAt', v_run.started_at, 'finishedAt', v_run.finished_at,
    'proposals', '[]'::jsonb
  );
end
$$;

create or replace function public.financial_reconciliation_finalize_automatic_analysis(p_run_id uuid)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
begin
  with source_usage as (
    select item->>'sourceType' as source_type, item->>'sourceId' as source_id,
           count(distinct p.base_source_id) as base_count
    from public.financial_reconciliation_automatic_proposals p
    join lateral (
      select item.value as item from jsonb_array_elements(p.items) item(value)
      union all
      select item.value as item
      from jsonb_array_elements(p.candidate_groups) candidate_group(value)
      join lateral jsonb_array_elements(
        case when jsonb_typeof(candidate_group.value) = 'array'
          then candidate_group.value else jsonb_build_array(candidate_group.value) end
      ) item(value) on true
    ) source_item on true
    where p.run_id = p_run_id and p.status in ('proposed', 'ambiguous')
    group by item->>'sourceType', item->>'sourceId'
  ), overlapping as (
    select distinct p.id
    from public.financial_reconciliation_automatic_proposals p
    join lateral (
      select item.value as item from jsonb_array_elements(p.items) item(value)
      union all
      select item.value as item
      from jsonb_array_elements(p.candidate_groups) candidate_group(value)
      join lateral jsonb_array_elements(
        case when jsonb_typeof(candidate_group.value) = 'array'
          then candidate_group.value else jsonb_build_array(candidate_group.value) end
      ) item(value) on true
    ) source_item on true
    join source_usage usage
      on usage.source_type = item->>'sourceType'
     and usage.source_id = item->>'sourceId'
    where p.run_id = p_run_id
      and p.status in ('proposed', 'ambiguous')
      and usage.base_count > 1
  )
  update public.financial_reconciliation_automatic_proposals proposal
  set status = 'ambiguous', reason = 'cross_base_overlap', updated_at = now()
  where proposal.id in (select id from overlapping);

  update public.financial_reconciliation_automatic_runs run
  set status = 'ready', analysis_completed_at = now(), updated_at = now(),
      analysis_error_code = null, analysis_error_at = null,
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
  v_base record;
  v_combination record;
  v_combination_count integer;
  v_page_count integer := 0;
  v_last_date date;
  v_last_id uuid;
  v_operator text;
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

  select value into strict v_rule
  from jsonb_array_elements(v_run.definition_config_snapshot) value
  order by (value->>'priority')::integer, value->>'ruleKey'
  limit 1;
  v_operator := v_rule->>'operator';
  if v_operator not in ('+', '-') then
    raise exception 'Automatic rule has no directional source operator.';
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

    if v_base.candidate_count > 12 then
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
        v_base.base_snapshot, v_base.candidates,
        jsonb_build_object('import_cgd_extrato_ordem', v_operator),
        (v_rule->>'differenceAllowed')::numeric, 4
      );

      if v_combination_count = 1 then
        select * into strict v_combination
        from public.financial_reconciliation_automatic_build_combinations(
          v_base.base_snapshot, v_base.candidates,
          jsonb_build_object('import_cgd_extrato_ordem', v_operator),
          (v_rule->>'differenceAllowed')::numeric, 4
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
               v_base.base_snapshot, v_base.candidates,
               jsonb_build_object('import_cgd_extrato_ordem', v_operator),
               (v_rule->>'differenceAllowed')::numeric, 4
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
        updated_at = now(), analysis_error_code = null, analysis_error_at = null
    where id = v_run.id;
  end if;

  if v_page_count < 25 then
    return public.financial_reconciliation_finalize_automatic_analysis(v_run.id);
  end if;
  return public.financial_reconciliation_automatic_progress_or_run(v_run.id);
exception when others then
  update public.financial_reconciliation_automatic_runs
  set analysis_error_code = sqlstate, analysis_error_at = now(), updated_at = now()
  where id = p_run_id;
  raise;
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
declare v_snapshot jsonb; v_run_id uuid; v_scope text; v_total integer;
begin
  if nullif(trim(coalesce(p_actor, '')), '') is null then raise exception 'Actor is required.'; end if;
  if p_client_request_id is null then raise exception 'Client request ID is required.'; end if;
  if p_mode = 'manual_rule' then v_scope := 'rule';
  elsif p_mode = 'manual_batch' then v_scope := 'batch';
  else raise exception 'Automatic analysis mode is invalid.'; end if;
  if p_rule_keys is null or cardinality(p_rule_keys) = 0
    or cardinality(p_rule_keys) <> cardinality(array(select distinct key from unnest(p_rule_keys) key)) then
    raise exception 'Automatic rule selection is invalid.';
  end if;

  lock table public.financial_reconciliation_source_rules in share row exclusive mode;
  lock table public.financial_reconciliation_automatic_rule_configs in share row exclusive mode;
  select coalesce(jsonb_agg(jsonb_build_object(
    'ruleKey', config.rule_key, 'ruleVersion', config.rule_version,
    'displayName', definition.display_name, 'priority', config.priority,
    'differenceAllowed', config.difference_allowed,
    'maxDifferenceDays', config.max_difference_days,
    'definition', definition.definition, 'operator', source_rule.operator
  ) order by config.priority, config.rule_key), '[]'::jsonb)
  into v_snapshot
  from public.financial_reconciliation_automatic_rule_configs config
  join public.financial_reconciliation_automatic_rule_definitions definition
    on definition.rule_key = config.rule_key and definition.version = config.rule_version
  join public.financial_reconciliation_source_rules source_rule
    on source_rule.base_source_type = definition.base_source_type
   and source_rule.matching_source_type = 'import_cgd_extrato_ordem'
  where config.rule_key = any(p_rule_keys) and config.enabled
    and config.max_difference_days between 0 and 90
    and ((p_mode = 'manual_rule' and config.allow_manual_execution)
      or (p_mode = 'manual_batch' and config.include_in_scheduled_batch));
  if jsonb_array_length(v_snapshot) <> cardinality(p_rule_keys) then
    raise exception 'Automatic rule is not enabled for requested analysis mode.';
  end if;
  if jsonb_array_length(v_snapshot) <> 1 then
    raise exception 'Resumable automatic analysis currently requires one selected rule.';
  end if;

  select count(*) into v_total
  from public.financial_documents document
  where document.fat = 'S' and document.payment = 'Banco'
    and document.document_date >= date '2026-01-01'
    and not exists (
      select 1 from public.financial_reconciliation_items item
      where item.source_type = 'financial_documents' and item.source_id = document.id
    );

  insert into public.financial_reconciliation_automatic_runs (
    trigger, scope, actor, client_request_id, definition_config_snapshot,
    analysis_processed, analysis_total
  ) values ('manual', v_scope, p_actor, p_client_request_id, v_snapshot, 0, v_total)
  on conflict (actor, client_request_id) do nothing returning id into v_run_id;
  if v_run_id is null then
    select id into strict v_run_id
    from public.financial_reconciliation_automatic_runs
    where actor = p_actor and client_request_id = p_client_request_id;
  end if;
  return public.continue_financial_reconciliation_automatic_analysis(v_run_id, p_actor);
end
$$;

create or replace function public.get_financial_reconciliation_automatic_active_run(p_actor text)
returns jsonb
language plpgsql
stable
security definer set search_path = public, pg_temp
as $$
declare v_run_id uuid;
begin
  if nullif(trim(coalesce(p_actor, '')), '') is null then raise exception 'Actor is required.'; end if;
  select id into v_run_id
  from public.financial_reconciliation_automatic_runs
  where actor = p_actor and trigger = 'manual'
    and status = 'analyzing' and analysis_completed_at is null
  order by started_at desc, id desc
  limit 1;
  if v_run_id is null then return null; end if;
  return public.financial_reconciliation_automatic_progress_or_run(v_run_id);
end
$$;

create or replace function public.continue_financial_reconciliation_automatic_oldest_analysis(p_worker text)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
declare v_run_id uuid; v_actor text; v_total integer;
begin
  if p_worker <> 'system:reconciliation' then
    raise exception 'Automatic reconciliation worker identity is invalid.';
  end if;
  select id, actor into v_run_id, v_actor
  from public.financial_reconciliation_automatic_runs
  where status = 'analyzing' and analysis_completed_at is null
  order by started_at, id
  for update skip locked
  limit 1;
  if v_run_id is null then return jsonb_build_object('continued', false); end if;

  select count(*) into v_total
  from public.financial_documents document
  where document.fat = 'S' and document.payment = 'Banco'
    and document.document_date >= date '2026-01-01'
    and not exists (
      select 1 from public.financial_reconciliation_items item
      where item.source_type = 'financial_documents' and item.source_id = document.id
    );
  update public.financial_reconciliation_automatic_runs
  set analysis_total = greatest(analysis_total, v_total), updated_at = now()
  where id = v_run_id;

  perform public.continue_financial_reconciliation_automatic_analysis(v_run_id, v_actor);
  return jsonb_build_object(
    'continued', true,
    'run', public.financial_reconciliation_automatic_progress_or_run(v_run_id)
  );
end
$$;

do $migration$
declare
  v_definition text;
  v_old_call text := $old$from public.financial_reconciliation_automatic_rule_candidates(
    v_proposal.rule_key,
    v_proposal.rule_version,
    (v_rule_snapshot->>'differenceAllowed')::numeric,
    (v_rule_snapshot->>'maxDifferenceDays')::integer
  ) candidates$old$;
  v_new_call text := $new$from public.financial_reconciliation_automatic_single_base_candidates(
    v_proposal.rule_key,
    v_proposal.rule_version,
    (v_rule_snapshot->>'differenceAllowed')::numeric,
    (v_rule_snapshot->>'maxDifferenceDays')::integer,
    v_proposal.base_source_id
  ) candidates$new$;
  v_old_guard text := $old$  if v_run.finished_at is not null then
    raise exception 'Automation proposal belongs to a finished run.';
  end if;$old$;
  v_new_guard text := $new$  if v_run.finished_at is not null then
    raise exception 'Automation proposal belongs to a finished run.';
  end if;
  if v_run.analysis_completed_at is null or v_run.status = 'analyzing' then
    raise exception 'Automatic analysis must finish before proposals can be executed.';
  end if;$new$;
begin
  select pg_get_functiondef(
    'public.execute_financial_reconciliation_automatic_proposal(uuid,text)'::regprocedure
  ) into strict v_definition;
  if strpos(v_definition, v_old_call) > 0 then
    v_definition := replace(v_definition, v_old_call, v_new_call);
  elsif strpos(v_definition, v_new_call) = 0 then
    raise exception 'Unexpected automatic proposal candidate revalidation definition.';
  end if;
  if strpos(v_definition, v_old_guard) > 0 then
    v_definition := replace(v_definition, v_old_guard, v_new_guard);
  elsif strpos(v_definition, 'Automatic analysis must finish before proposals can be executed.') = 0 then
    raise exception 'Unexpected automatic proposal run-state guard definition.';
  end if;
  execute v_definition;
end
$migration$;

alter table public.financial_reconciliation_cgd_match_search enable row level security;
revoke all on table public.financial_reconciliation_cgd_match_search
  from public, anon, authenticated, service_role;
grant select on table public.financial_reconciliation_cgd_match_search to service_role;

revoke all on function public.financial_reconciliation_sync_cgd_match_search() from public, anon, authenticated;
revoke all on function public.financial_reconciliation_automatic_candidates_for_base_ids(text,integer,numeric,integer,uuid[]) from public, anon, authenticated;
revoke all on function public.financial_reconciliation_automatic_candidate_page(text,integer,numeric,integer,date,uuid,integer) from public, anon, authenticated;
revoke all on function public.financial_reconciliation_automatic_single_base_candidates(text,integer,numeric,integer,uuid) from public, anon, authenticated;
revoke all on function public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer) from public, anon, authenticated;
revoke all on function public.financial_reconciliation_automatic_progress_or_run(uuid) from public, anon, authenticated;
revoke all on function public.financial_reconciliation_finalize_automatic_analysis(uuid) from public, anon, authenticated;
revoke all on function public.continue_financial_reconciliation_automatic_analysis(uuid,text) from public, anon, authenticated;
revoke all on function public.create_financial_reconciliation_automatic_analysis(text[],text,text,uuid) from public, anon, authenticated;
revoke all on function public.get_financial_reconciliation_automatic_active_run(text) from public, anon, authenticated;
revoke all on function public.continue_financial_reconciliation_automatic_oldest_analysis(text) from public, anon, authenticated;
revoke all on function public.get_financial_reconciliation_automatic_run(uuid) from public, anon, authenticated;
revoke all on function public.execute_financial_reconciliation_automatic_proposal(uuid,text) from public, anon, authenticated;

grant execute on function public.financial_reconciliation_automatic_candidate_page(text,integer,numeric,integer,date,uuid,integer) to service_role;
grant execute on function public.financial_reconciliation_automatic_single_base_candidates(text,integer,numeric,integer,uuid) to service_role;
grant execute on function public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer) to service_role;
grant execute on function public.continue_financial_reconciliation_automatic_analysis(uuid,text) to service_role;
grant execute on function public.create_financial_reconciliation_automatic_analysis(text[],text,text,uuid) to service_role;
grant execute on function public.get_financial_reconciliation_automatic_active_run(text) to service_role;
grant execute on function public.continue_financial_reconciliation_automatic_oldest_analysis(text) to service_role;
grant execute on function public.get_financial_reconciliation_automatic_run(uuid) to service_role;
grant execute on function public.execute_financial_reconciliation_automatic_proposal(uuid,text) to service_role;

notify pgrst, 'reload schema';
