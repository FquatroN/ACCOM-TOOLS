insert into public.financial_reconciliation_automatic_rule_definitions (
  rule_key, version, display_name, base_source_type,
  destination_source_types, logic_description, definition
) values (
  'financial_documents_cgd_credit_card', 1,
  'Financial Documents to CGD Credit Card', 'financial_documents',
  '["import_cgd_cartao_credito"]'::jsonb,
  'Payment must equal exactly Visa. Each credit-card candidate must satisfy invoice containment, description similarity, or supplier word similarity. Exactly one one-to-four-record amount combination is executable.',
  '{
    "baseEligibility":{"payment":{"operator":"exact_text_equal","value":"Visa","caseSensitive":true,"trim":false}},
    "identityBranches":{"document_number":{"algorithm":"symmetric_compact_containment"},"description_similarity":{"algorithm":"similarity"},"supplier_similarity":{"algorithm":"word_similarity"}},
    "documentNumberMinimumCompactLength":4,
    "descriptionSimilarityThreshold":0.55,
    "supplierWordSimilarityThreshold":0.60,
    "maxDestinationRecords":4,
    "maxIdentityCandidatesPerBase":12
  }'::jsonb
) on conflict (rule_key, version) do nothing;

do $migration$
begin
  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_definitions definition
    where definition.rule_key = 'financial_documents_cgd_credit_card'
      and definition.version = 1
      and definition.display_name = 'Financial Documents to CGD Credit Card'
      and definition.base_source_type = 'financial_documents'
      and definition.destination_source_types = '["import_cgd_cartao_credito"]'::jsonb
      and definition.logic_description = 'Payment must equal exactly Visa. Each credit-card candidate must satisfy invoice containment, description similarity, or supplier word similarity. Exactly one one-to-four-record amount combination is executable.'
      and definition.definition = '{
        "baseEligibility":{"payment":{"operator":"exact_text_equal","value":"Visa","caseSensitive":true,"trim":false}},
        "identityBranches":{"document_number":{"algorithm":"symmetric_compact_containment"},"description_similarity":{"algorithm":"similarity"},"supplier_similarity":{"algorithm":"word_similarity"}},
        "documentNumberMinimumCompactLength":4,
        "descriptionSimilarityThreshold":0.55,
        "supplierWordSimilarityThreshold":0.60,
        "maxDestinationRecords":4,
        "maxIdentityCandidatesPerBase":12
      }'::jsonb
  ) then
    raise exception 'Managed automatic reconciliation credit-card rule differs from the expected immutable definition.';
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_source_rules source_rule
    where source_rule.base_source_type = 'financial_documents'
      and source_rule.matching_source_type = 'import_cgd_cartao_credito'
      and source_rule.operator = '+'
  ) then
    raise exception 'Managed automatic reconciliation credit-card source rule must use operator +.';
  end if;
end
$migration$;

create or replace function public.replace_financial_reconciliation_source_rules(p_rules jsonb)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
begin
  if jsonb_typeof(p_rules) <> 'array' then
    raise exception 'Reconciliation rules must be an array.';
  end if;
  if p_rules is null then
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

  lock table public.financial_reconciliation_source_rules in share row exclusive mode;
  delete from public.financial_reconciliation_source_rules;
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

do $migration$
declare
  v_had_credit_card_config boolean;
begin
  select exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs config
    where config.rule_key = 'financial_documents_cgd_credit_card'
  ) into v_had_credit_card_config;

  if not v_had_credit_card_config then
    set constraints financial_reconciliation_automatic_rule_configs_priority_key deferred;

    update public.financial_reconciliation_automatic_rule_configs
    set priority = 1,
        updated_at = now()
    where rule_key = 'financial_documents_cgd_bank_statement';

    insert into public.financial_reconciliation_automatic_rule_configs (
      rule_key, rule_version, enabled, allow_manual_execution,
      include_in_scheduled_batch, difference_allowed, max_difference_days, priority
    ) values (
      'financial_documents_cgd_credit_card', 1, false, false, false, 0.00, 10, 2
    ) on conflict (rule_key) do nothing;

    if not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_configs config
      where config.rule_key = 'financial_documents_cgd_bank_statement'
        and config.priority = 1
    ) then
      raise exception 'Managed Banco automatic reconciliation priority could not be normalized to 1.';
    end if;
  end if;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs config
    where config.rule_key = 'financial_documents_cgd_credit_card'
      and config.rule_version = 1
  ) then
    raise exception 'Managed credit-card automatic reconciliation config must reference version 1.';
  end if;
end
$migration$;

create table if not exists public.financial_reconciliation_cgd_credit_card_match_search (
  source_id uuid primary key references public.import_cgd_cartao_credito(id) on update cascade on delete cascade,
  source_date date not null,
  amount numeric,
  description text,
  normalized_description text not null,
  compact_description text not null,
  updated_at timestamptz not null default now()
);

do $migration$
declare
  v_fk record;
begin
  select constraint_row.oid, constraint_row.conname, constraint_row.confupdtype
  into strict v_fk
  from pg_constraint constraint_row
  where constraint_row.conrelid = 'public.financial_reconciliation_cgd_credit_card_match_search'::regclass
    and constraint_row.confrelid = 'public.import_cgd_cartao_credito'::regclass
    and constraint_row.contype = 'f';
  if v_fk.confupdtype <> 'c' then
    execute format(
      'alter table public.financial_reconciliation_cgd_credit_card_match_search drop constraint %I',
      v_fk.conname
    );
    alter table public.financial_reconciliation_cgd_credit_card_match_search
      add constraint financial_reconciliation_cgd_credit_card_match_search_source_id_fkey
      foreign key (source_id) references public.import_cgd_cartao_credito(id)
      on update cascade on delete cascade;
  end if;
end
$migration$;

create index if not exists financial_reconciliation_cgd_credit_card_match_search_date_id_idx
  on public.financial_reconciliation_cgd_credit_card_match_search (source_date, source_id);

do $migration$
declare
  v_trgm_schema text;
begin
  select n.nspname into strict v_trgm_schema
  from pg_extension e join pg_namespace n on n.oid = e.extnamespace
  where e.extname = 'pg_trgm';

  execute format(
    'create index if not exists financial_reconciliation_cgd_credit_card_match_search_normalized_trgm_idx
       on public.financial_reconciliation_cgd_credit_card_match_search
       using gin (normalized_description %I.gin_trgm_ops)',
    v_trgm_schema
  );
  execute format(
    'create index if not exists financial_reconciliation_cgd_credit_card_match_search_compact_trgm_idx
       on public.financial_reconciliation_cgd_credit_card_match_search
       using gin (compact_description %I.gin_trgm_ops)',
    v_trgm_schema
  );
end
$migration$;

create or replace function public.financial_reconciliation_sync_cgd_credit_card_match_search()
returns trigger
language plpgsql
security definer set search_path = public, pg_temp
as $$
begin
  if tg_op = 'DELETE' then
    delete from public.financial_reconciliation_cgd_credit_card_match_search where source_id = old.id;
    return old;
  end if;

  if tg_op = 'UPDATE' and old.id is distinct from new.id then
    delete from public.financial_reconciliation_cgd_credit_card_match_search where source_id = old.id;
  end if;

  if new.data is null then
    delete from public.financial_reconciliation_cgd_credit_card_match_search where source_id = new.id;
    return new;
  end if;

  insert into public.financial_reconciliation_cgd_credit_card_match_search (
    source_id, source_date, amount, description,
    normalized_description, compact_description, updated_at
  ) values (
    new.id, new.data, new.valor, new.descricao,
    coalesce(public.financial_reconciliation_match_normalize(new.descricao), ''),
    coalesce(public.financial_reconciliation_match_compact(new.descricao), ''),
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

drop trigger if exists financial_reconciliation_sync_cgd_credit_card_match_search_trigger
  on public.import_cgd_cartao_credito;
create trigger financial_reconciliation_sync_cgd_credit_card_match_search_trigger
after insert or update or delete on public.import_cgd_cartao_credito
for each row execute function public.financial_reconciliation_sync_cgd_credit_card_match_search();

insert into public.financial_reconciliation_cgd_credit_card_match_search (
  source_id, source_date, amount, description,
  normalized_description, compact_description, updated_at
)
select
  card.id, card.data, card.valor, card.descricao,
  coalesce(public.financial_reconciliation_match_normalize(card.descricao), ''),
  coalesce(public.financial_reconciliation_match_compact(card.descricao), ''),
  now()
from public.import_cgd_cartao_credito card
where card.data is not null
on conflict (source_id) do update set
  source_date = excluded.source_date,
  amount = excluded.amount,
  description = excluded.description,
  normalized_description = excluded.normalized_description,
  compact_description = excluded.compact_description,
  updated_at = excluded.updated_at;

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
    else null
  end
$$;

do $migration$
declare v_trgm_schema text; v_sql text;
begin
  select n.nspname into strict v_trgm_schema
  from pg_extension e join pg_namespace n on n.oid = e.extnamespace
  where e.extname = 'pg_trgm';

  v_sql := format($function$
create or replace function public.financial_reconciliation_automatic_bank_candidates_for_base_ids(
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

do $migration$
declare v_trgm_schema text; v_sql text;
begin
  select n.nspname into strict v_trgm_schema
  from pg_extension e join pg_namespace n on n.oid = e.extnamespace
  where e.extname = 'pg_trgm';

  v_sql := format($function$
create or replace function public.financial_reconciliation_automatic_credit_card_candidates_for_base_ids(
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
set pg_trgm.similarity_threshold = 0.55
set pg_trgm.word_similarity_threshold = 0.60
as $body$
  with bases as materialized (
    select
      d.id, d.document_date, d.doc_number, d.description, d.supplier_name, d.amount,
      public.financial_reconciliation_match_compact(d.doc_number) as compact_document_number,
      public.financial_reconciliation_match_normalize(d.description) as normalized_document_description,
      public.financial_reconciliation_match_normalize(d.supplier_name) as normalized_supplier_name
    from public.financial_documents d
    where p_rule_key = 'financial_documents_cgd_credit_card'
      and p_rule_version = 1
      and p_max_difference_days between 0 and 90
      and d.id = any(coalesce(p_base_ids, array[]::uuid[]))
      and d.fat = 'S'
      and d.payment = 'Visa'
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
      search.normalized_description, search.compact_description,
      d.compact_document_number,
      d.normalized_document_description,
      d.normalized_supplier_name
    from bases d
    left join lateral (
      select s.*
      from public.financial_reconciliation_cgd_credit_card_match_search s
      where s.source_date between d.document_date - p_max_difference_days
                              and d.document_date + p_max_difference_days
        and s.source_date >= date '2026-01-01'
        and s.amount is not null
        and not exists (
          select 1 from public.financial_reconciliation_items i
          where i.source_type = 'import_cgd_cartao_credito' and i.source_id = s.source_id
        )
        and (
          (char_length(d.compact_document_number) >= 4
            and nullif(s.compact_description, '') is not null
            and (
              s.compact_description like '%%' || d.compact_document_number || '%%'
              or d.compact_document_number like '%%' || s.compact_description || '%%'
            ))
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
          and nullif(q.compact_description, '') is not null
          and (
            position(q.compact_document_number in q.compact_description) > 0
            or position(q.compact_description in q.compact_document_number) > 0
          ),
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
        'sourceType', 'import_cgd_cartao_credito',
        'sourceId', source_id, 'sourceDate', source_date,
        'amount', amount, 'description', description,
        'evidence', jsonb_build_object(
          'documentNumber', jsonb_build_object(
            'matched', document_number_matched, 'normalized', compact_document_number
          ),
          'description', jsonb_build_object(
            'matched', description_score >= 0.55,
            'score', description_score, 'threshold', 0.55
          ),
          'supplier', jsonb_build_object(
            'matched', supplier_score >= 0.60,
            'score', supplier_score, 'threshold', 0.60
          )
        )
      ) order by source_date, source_id) filter (
        where source_id is not null and (
          document_number_matched or description_score >= 0.55 or supplier_score >= 0.60
        )
      ), '[]'::jsonb) as candidates,
      count(*) filter (
        where source_id is not null and (
          document_number_matched or description_score >= 0.55 or supplier_score >= 0.60
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
language plpgsql
stable
security definer set search_path = public, pg_temp
as $$
begin
  if public.financial_reconciliation_automatic_rule_contract(p_rule_key, p_rule_version) is null then
    raise exception 'Automatic reconciliation rule is unsupported.';
  end if;
  if p_max_difference_days not between 0 and 90 then
    raise exception 'Max difference in days must be between 0 and 90.';
  end if;

  if p_rule_key = 'financial_documents_cgd_bank_statement' and p_rule_version = 2 then
    return query
    select *
    from public.financial_reconciliation_automatic_bank_candidates_for_base_ids(
      p_rule_key, p_rule_version, p_difference_allowed, p_max_difference_days, p_base_ids
    );
  elsif p_rule_key = 'financial_documents_cgd_credit_card' and p_rule_version = 1 then
    return query
    select *
    from public.financial_reconciliation_automatic_credit_card_candidates_for_base_ids(
      p_rule_key, p_rule_version, p_difference_allowed, p_max_difference_days, p_base_ids
    );
  end if;
end
$$;

create or replace function public.financial_reconciliation_automatic_base_page(
  p_rule_key text,
  p_rule_version integer,
  p_after_date date,
  p_after_id uuid,
  p_page_size integer
)
returns table (
  id uuid,
  document_date date
)
language plpgsql
stable
security definer set search_path = public, pg_temp
as $$
declare
  v_contract jsonb;
begin
  v_contract := public.financial_reconciliation_automatic_rule_contract(p_rule_key, p_rule_version);
  if v_contract is null then
    raise exception 'Automatic reconciliation rule is unsupported.';
  end if;
  if p_page_size not between 1 and 25 then
    raise exception 'Automatic analysis page size must be between 1 and 25.';
  end if;

  return query
  select document.id, document.document_date
  from public.financial_documents document
  where document.fat = 'S'
    and document.payment = v_contract->>'payment'
    and document.document_date >= date '2026-01-01'
    and (p_after_date is null or (document.document_date, document.id) > (p_after_date, p_after_id))
    and not exists (
      select 1 from public.financial_reconciliation_items item
      where item.source_type = 'financial_documents' and item.source_id = document.id
    )
  order by document.document_date, document.id
  limit p_page_size;
end
$$;

create or replace function public.financial_reconciliation_automatic_base_count(
  p_rule_key text,
  p_rule_version integer
)
returns bigint
language plpgsql
stable
security definer set search_path = public, pg_temp
as $$
declare
  v_contract jsonb;
  v_count bigint;
begin
  v_contract := public.financial_reconciliation_automatic_rule_contract(p_rule_key, p_rule_version);
  if v_contract is null then
    raise exception 'Automatic reconciliation rule is unsupported.';
  end if;

  select count(*) into v_count
  from public.financial_documents document
  where document.fat = 'S'
    and document.payment = v_contract->>'payment'
    and document.document_date >= date '2026-01-01'
    and not exists (
      select 1 from public.financial_reconciliation_items item
      where item.source_type = 'financial_documents' and item.source_id = document.id
    );
  return v_count;
end
$$;

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
begin
  if p_max_difference_days not between 0 and 90 then
    raise exception 'Max difference in days must be between 0 and 90.';
  end if;

  return query
  with page as materialized (
    select base.id, base.document_date
    from public.financial_reconciliation_automatic_base_page(
      p_rule_key, p_rule_version, p_after_date, p_after_id, p_page_size
    ) base
  )
  select * from public.financial_reconciliation_automatic_candidates_for_base_ids(
    p_rule_key,
    p_rule_version,
    p_difference_allowed,
    p_max_difference_days,
    array(select page.id from page order by page.document_date, page.id)
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
language plpgsql
stable
security definer set search_path = public, pg_temp
as $$
declare
  v_contract jsonb;
begin
  v_contract := public.financial_reconciliation_automatic_rule_contract(p_rule_key, p_rule_version);
  if v_contract is null then
    raise exception 'Automatic reconciliation rule is unsupported.';
  end if;

  return query
  select *
  from public.financial_reconciliation_automatic_candidates_for_base_ids(
    p_rule_key, p_rule_version, p_difference_allowed, p_max_difference_days,
    array(
      select document.id
      from public.financial_documents document
      where document.fat = 'S'
        and document.payment = v_contract->>'payment'
        and document.document_date >= date '2026-01-01'
        and not exists (
          select 1 from public.financial_reconciliation_items item
          where item.source_type = 'financial_documents' and item.source_id = document.id
        )
      order by document.document_date, document.id
    )
  );
end
$$;

with ranked_open_manual_runs as (
  select
    run.id,
    row_number() over (
      partition by actor
      order by started_at desc, created_at desc, id desc
    ) as actor_run_rank
  from public.financial_reconciliation_automatic_runs run
  where run.trigger = 'manual' and run.finished_at is null
)
update public.financial_reconciliation_automatic_runs run
set status = 'failed',
    error_summary = 'A newer unfinished manual automatic analysis superseded this run.',
    analysis_error_code = 'superseded_open_manual_run',
    analysis_error_at = now(),
    finished_at = now(),
    updated_at = now()
from ranked_open_manual_runs ranked
where run.id = ranked.id and ranked.actor_run_rank > 1;

create unique index if not exists financial_reconciliation_automatic_runs_open_manual_actor_uidx
  on public.financial_reconciliation_automatic_runs (actor)
  where trigger = 'manual' and finished_at is null;

create or replace function public.financial_reconciliation_finalize_automatic_analysis(p_run_id uuid)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
begin
  with source_usage as (
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
  ), overlapping as (
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
    join source_usage usage
      on usage.source_type = item->>'sourceType'
     and usage.source_id = item->>'sourceId'
    where proposal.run_id = p_run_id
      and proposal.status in ('proposed', 'ambiguous')
      and usage.base_count > 1
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

  select public.financial_reconciliation_automatic_base_count(
    v_snapshot->0->>'ruleKey',
    (v_snapshot->0->>'ruleVersion')::integer
  ) into v_total;

  insert into public.financial_reconciliation_automatic_runs (
    trigger, scope, actor, client_request_id, definition_config_snapshot,
    analysis_processed, analysis_total
  ) values (
    'manual', 'rule', p_actor, p_client_request_id, v_snapshot, 0, v_total
  ) returning id into v_run_id;

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
  select run.id into v_run_id
  from public.financial_reconciliation_automatic_runs run
  where run.actor = p_actor
    and run.trigger = 'manual'
    and run.finished_at is null
  order by run.started_at desc, run.id desc
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
declare
  v_run_id uuid;
  v_actor text;
  v_snapshot jsonb;
  v_rule jsonb;
  v_total bigint;
begin
  if p_worker is distinct from 'system:reconciliation' then
    raise exception 'Automatic reconciliation worker identity is invalid.';
  end if;

  select run.id, run.actor, run.definition_config_snapshot
  into v_run_id, v_actor, v_snapshot
  from public.financial_reconciliation_automatic_runs run
  where run.status = 'analyzing' and run.analysis_completed_at is null
  order by run.started_at, run.id
  for update skip locked
  limit 1;
  if v_run_id is null then return jsonb_build_object('continued', false); end if;

  if jsonb_typeof(v_snapshot) = 'array' then
    if jsonb_array_length(v_snapshot) = 1 then
      v_rule := v_snapshot->0;
      begin
        if public.financial_reconciliation_automatic_rule_contract(
          v_rule->>'ruleKey',
          (v_rule->>'ruleVersion')::integer
        ) is not null then
          select public.financial_reconciliation_automatic_base_count(
            v_rule->>'ruleKey',
            (v_rule->>'ruleVersion')::integer
          ) into v_total;
          update public.financial_reconciliation_automatic_runs
          set analysis_total = greatest(analysis_total, v_total), updated_at = now()
          where id = v_run_id;
        end if;
      exception when others then
        null;
      end;
    end if;
  end if;

  perform public.continue_financial_reconciliation_automatic_analysis(v_run_id, v_actor);
  return jsonb_build_object(
    'continued', true,
    'run', public.financial_reconciliation_automatic_progress_or_run(v_run_id)
  );
end
$$;

create or replace function public.financial_reconciliation_automatic_lock_destination_items(
  p_source_type text,
  p_items jsonb
)
returns integer
language plpgsql
security definer set search_path = public, pg_temp
as $$
declare
  v_count integer;
begin
  if p_source_type = 'import_cgd_extrato_ordem' then
    perform bank.id
    from jsonb_array_elements(p_items) item(value)
    join public.import_cgd_extrato_ordem bank
      on bank.id = (item.value->>'sourceId')::uuid
    order by bank.data, bank.id
    for update of bank;
  elsif p_source_type = 'import_cgd_cartao_credito' then
    perform card.id
    from jsonb_array_elements(p_items) item(value)
    join public.import_cgd_cartao_credito card
      on card.id = (item.value->>'sourceId')::uuid
    order by card.data, card.id
    for update of card;
  else
    raise exception 'Automatic reconciliation destination source is unsupported.';
  end if;
  get diagnostics v_count = row_count;
  return v_count;
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
  v_run public.financial_reconciliation_automatic_runs%rowtype;
  v_proposal public.financial_reconciliation_automatic_proposals%rowtype;
  v_contract jsonb;
  v_rule_snapshot jsonb;
  v_destination_source_type text;
  v_max_candidates integer;
  v_max_destination_records integer;
  v_current_definition jsonb;
  v_current_display_name text;
  v_current_base_source_type text;
  v_current_destination_source_types jsonb;
  v_current_rule_version integer;
  v_current_operator text;
  v_locked_destination_count integer;
  v_distinct_destination_count integer;
  v_base record;
  v_combination record;
  v_combination_count integer;
  v_current_evidence jsonb;
  v_action_result jsonb;
  v_item record;
  v_reconciliation_id uuid;
  v_expected_item_count integer;
  v_actual_item_count integer;
  v_expected_matching_source_rules jsonb;
  v_actual_matching_source_rules jsonb;
  v_actual_difference numeric;
  v_completion_action text;
  v_comment text;
  v_trigger_label text;
  v_failure_message text;
  v_failure_detail text;
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
  from public.financial_reconciliation_automatic_runs
  where id = v_run_id
  for update;

  select * into strict v_proposal
  from public.financial_reconciliation_automatic_proposals
  where id = p_proposal_id
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
  if v_proposal.status in ('ambiguous', 'skipped', 'deselected', 'failed') then
    raise exception 'Automation proposal with status % cannot be executed.', v_proposal.status;
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

  v_contract := public.financial_reconciliation_automatic_rule_contract(
    v_proposal.rule_key,
    v_proposal.rule_version
  );
  if v_contract is null then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_version_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_version_changed'
    );
  end if;
  v_destination_source_type := v_contract->>'destinationSourceType';
  v_max_candidates := (v_contract->>'maxCandidates')::integer;
  v_max_destination_records := (v_contract->>'maxDestinationRecords')::integer;

  if v_proposal.base_source_type <> 'financial_documents'
    or jsonb_typeof(v_run.definition_config_snapshot) <> 'array'
    or jsonb_array_length(v_run.definition_config_snapshot) <> 1 then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_snapshot_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_snapshot_changed'
    );
  end if;
  v_rule_snapshot := v_run.definition_config_snapshot->0;
  if jsonb_typeof(v_rule_snapshot) <> 'object'
    or v_rule_snapshot->>'ruleKey' is distinct from v_proposal.rule_key
    or coalesce(v_rule_snapshot->>'ruleVersion', '') !~ '^[0-9]+$' then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_snapshot_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_snapshot_changed'
    );
  end if;
  if (v_rule_snapshot->>'ruleVersion')::integer <> v_proposal.rule_version
    or v_rule_snapshot->>'destinationSourceType' is distinct from v_destination_source_type
    or nullif(v_rule_snapshot->>'displayName', '') is null
    or jsonb_typeof(v_rule_snapshot->'definition') is distinct from 'object'
    or coalesce(v_rule_snapshot->>'operator', '') not in ('+', '-')
    or coalesce(v_rule_snapshot->>'differenceAllowed', '') !~ '^[0-9]+(\.[0-9]+)?$'
    or coalesce(v_rule_snapshot->>'maxDifferenceDays', '') !~ '^[0-9]+$'
    or coalesce(v_rule_snapshot->>'priority', '') !~ '^[0-9]+$'
    or coalesce(v_destination_source_type, '') = ''
    or coalesce(v_max_candidates, 0) < 1
    or coalesce(v_max_destination_records, 0) < 1 then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_snapshot_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_snapshot_changed'
    );
  end if;

  select
    definition.definition,
    definition.display_name,
    definition.base_source_type,
    definition.destination_source_types,
    config.rule_version,
    source_rule.operator
  into
    v_current_definition,
    v_current_display_name,
    v_current_base_source_type,
    v_current_destination_source_types,
    v_current_rule_version,
    v_current_operator
  from public.financial_reconciliation_automatic_rule_definitions definition
  join public.financial_reconciliation_automatic_rule_configs config
    on config.rule_key = definition.rule_key
  join public.financial_reconciliation_source_rules source_rule
    on source_rule.base_source_type = definition.base_source_type
   and source_rule.matching_source_type = v_destination_source_type
  where definition.rule_key = v_proposal.rule_key
    and definition.version = v_proposal.rule_version
  for share of definition, config, source_rule;

  if not found
    or v_current_rule_version is distinct from v_proposal.rule_version
    or v_current_definition is distinct from v_rule_snapshot->'definition'
    or v_current_display_name is distinct from v_rule_snapshot->>'displayName'
    or v_current_base_source_type is distinct from v_proposal.base_source_type
    or v_current_destination_source_types is distinct from jsonb_build_array(v_destination_source_type) then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_snapshot_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_snapshot_changed'
    );
  end if;
  if v_current_operator is distinct from v_rule_snapshot->>'operator' then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'operator_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'operator_changed'
    );
  end if;
  if v_proposal.allowed_difference is distinct from
      (v_rule_snapshot->>'differenceAllowed')::numeric then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'tolerance_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'tolerance_changed'
    );
  end if;

  perform document.id
  from public.financial_documents document
  where document.id = v_proposal.base_source_id
  for update;
  if not found then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'source_snapshot_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'source_snapshot_changed'
    );
  end if;
  if jsonb_array_length(v_proposal.items) < 1
    or jsonb_array_length(v_proposal.items) > v_max_destination_records
    or exists (
      select 1
      from jsonb_array_elements(v_proposal.items) item(value)
      where value->>'sourceType' is distinct from v_destination_source_type
        or coalesce(value->>'sourceId', '') !~
          '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
        or coalesce(value->>'sourceDate', '') !~ '^\d{4}-\d{2}-\d{2}$'
    ) then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'source_snapshot_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'source_snapshot_changed'
    );
  end if;
  select count(distinct item.value->>'sourceId')
  into v_distinct_destination_count
  from jsonb_array_elements(v_proposal.items) item(value);
  if v_distinct_destination_count <> jsonb_array_length(v_proposal.items) then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'source_snapshot_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'source_snapshot_changed'
    );
  end if;

  v_locked_destination_count := public.financial_reconciliation_automatic_lock_destination_items(
    v_destination_source_type,
    v_proposal.items
  );
  if v_locked_destination_count <> jsonb_array_length(v_proposal.items) then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'source_snapshot_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'source_snapshot_changed'
    );
  end if;

  select * into v_base
  from public.financial_reconciliation_automatic_single_base_candidates(
    v_proposal.rule_key,
    v_proposal.rule_version,
    (v_rule_snapshot->>'differenceAllowed')::numeric,
    (v_rule_snapshot->>'maxDifferenceDays')::integer,
    v_proposal.base_source_id
  ) candidates;
  if not found
    or v_base.base_source_date is distinct from v_proposal.base_source_date
    or v_base.base_snapshot is distinct from v_proposal.base_snapshot
    or v_base.candidate_count > v_max_candidates then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'source_snapshot_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'source_snapshot_changed'
    );
  end if;

  select count(*) into v_combination_count
  from public.financial_reconciliation_automatic_build_combinations(
    v_base.base_snapshot,
    v_base.candidates,
    jsonb_build_object(v_destination_source_type, v_rule_snapshot->>'operator'),
    (v_rule_snapshot->>'differenceAllowed')::numeric,
    v_max_destination_records
  );
  if v_combination_count <> 1 then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'combination_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'combination_changed'
    );
  end if;
  select * into strict v_combination
  from public.financial_reconciliation_automatic_build_combinations(
    v_base.base_snapshot,
    v_base.candidates,
    jsonb_build_object(v_destination_source_type, v_rule_snapshot->>'operator'),
    (v_rule_snapshot->>'differenceAllowed')::numeric,
    v_max_destination_records
  );
  select coalesce(jsonb_agg(item.value->'evidence' order by item.ordinality), '[]'::jsonb)
  into v_current_evidence
  from jsonb_array_elements(v_combination.items) with ordinality item(value, ordinality);

  if v_combination.signature is distinct from v_proposal.signature
    or v_combination.items is distinct from v_proposal.items
    or v_current_evidence is distinct from v_proposal.evidence
    or v_combination.calculated_difference is distinct from v_proposal.calculated_difference then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'proposal_evidence_changed', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'proposal_evidence_changed'
    );
  end if;

  begin
    update public.financial_reconciliation_automatic_proposals
    set status = 'executing', reason = '', error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;

    v_action_result := public.financial_reconciliation_action(
      'start', p_actor, null,
      v_proposal.base_source_type, v_proposal.base_source_id, null
    );
    v_reconciliation_id := (v_action_result#>>'{reconciliation,id}')::uuid;
    if v_reconciliation_id is null then
      raise exception 'Automatic reconciliation start returned no reconciliation.';
    end if;

    for v_item in
      select value
      from jsonb_array_elements(v_proposal.items) item(value)
      order by value->>'sourceType', (value->>'sourceDate')::date, value->>'sourceId'
    loop
      perform public.financial_reconciliation_action(
        'add_item', p_actor, v_reconciliation_id,
        v_item.value->>'sourceType', (v_item.value->>'sourceId')::uuid, null
      );
    end loop;

    v_expected_item_count := 1 + jsonb_array_length(v_proposal.items);
    select count(*) into v_actual_item_count
    from public.financial_reconciliation_items item
    where item.reconciliation_id = v_reconciliation_id;
    if v_actual_item_count <> v_expected_item_count
      or exists (
        select 1
        from public.financial_reconciliation_items locked_item
        where locked_item.reconciliation_id = v_reconciliation_id
          and (
            (
              locked_item.source_type = v_proposal.base_source_type
              and (
                locked_item.source_id <> v_proposal.base_source_id
                or locked_item.amount_snapshot is distinct from
                  (v_base.base_snapshot->>'amount')::numeric
              )
            )
            or
            (
              locked_item.source_type <> v_proposal.base_source_type
              and not exists (
                select 1
                from jsonb_array_elements(v_proposal.items) proposal_item(value)
                where proposal_item.value->>'sourceType' = locked_item.source_type
                  and (proposal_item.value->>'sourceId')::uuid = locked_item.source_id
                  and (proposal_item.value->>'amount')::numeric = locked_item.amount_snapshot
              )
            )
          )
      ) then
      raise exception 'Automatic reconciliation lifecycle snapshots changed after revalidation.';
    end if;

    v_expected_matching_source_rules := jsonb_build_object(
      'sourceType', v_destination_source_type,
      'operator', v_rule_snapshot->>'operator'
    );
    select matching_rule.value, reconciliation.difference_amount
    into v_actual_matching_source_rules, v_actual_difference
    from public.financial_reconciliations reconciliation
    join lateral jsonb_array_elements(reconciliation.matching_source_rules) matching_rule(value)
      on matching_rule.value->>'sourceType' = v_destination_source_type
    where reconciliation.id = v_reconciliation_id;
    if not found
      or v_actual_matching_source_rules is distinct from v_expected_matching_source_rules
      or v_actual_difference is distinct from v_proposal.calculated_difference
      or abs(v_actual_difference) > v_proposal.allowed_difference then
      raise exception 'Automatic reconciliation lifecycle snapshots changed after revalidation.';
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
      v_completion_action := 'complete';
      v_comment := null;
    else
      v_completion_action := 'force_complete';
      v_trigger_label := case v_run.trigger when 'manual' then 'Manual' else 'Scheduled' end;
      v_comment := 'Automatically completed by rule '
        || v_rule_snapshot->>'displayName'
        || ' v' || v_proposal.rule_version::text || '; difference '
        || chr(8364) || to_char(v_actual_difference, 'FM999999999990.00')
        || ' within allowed tolerance ' || chr(8364)
        || to_char(v_proposal.allowed_difference, 'FM999999999990.00')
        || '; trigger ' || v_trigger_label || '; batch ' || v_run.id::text || '.';
    end if;

    if v_completion_action = 'complete' then
      perform public.financial_reconciliation_action(
        'complete', p_actor, v_reconciliation_id, null, null, null
      );
    else
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
          'definition', v_rule_snapshot->'definition'
        ),
        'configSnapshot', jsonb_build_object(
          'differenceAllowed', (v_rule_snapshot->>'differenceAllowed')::numeric,
          'maxDifferenceDays', (v_rule_snapshot->>'maxDifferenceDays')::integer,
          'priority', (v_rule_snapshot->>'priority')::integer
        ),
        'operatorSnapshot', jsonb_build_object(
          v_destination_source_type, v_rule_snapshot->>'operator'
        ),
        'baseSnapshot', v_proposal.base_snapshot,
        'destinationSnapshots', v_proposal.items,
        'identityEvidence', v_proposal.evidence,
        'proposalSignature', v_proposal.signature,
        'trigger', v_run.trigger,
        'runId', v_run.id,
        'proposalId', v_proposal.id,
        'tolerance', v_proposal.allowed_difference,
        'calculatedDifference', v_actual_difference
      )
    );

    update public.financial_reconciliation_automatic_proposals
    set status = 'completed',
        reconciliation_id = v_reconciliation_id,
        completed_at = now(),
        reason = '',
        error = '',
        error_detail = '',
        updated_at = now()
    where id = v_proposal.id;
  exception when others then
    get stacked diagnostics
      v_failure_message = message_text,
      v_failure_detail = pg_exception_detail;
    if v_failure_message in (
      'Automatic reconciliation lifecycle snapshots changed after revalidation.',
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
        error_detail = left(concat_ws(' ', v_failure_message, nullif(v_failure_detail, '')), 2000),
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

create table if not exists public.financial_reconciliation_automatic_batches (
  id uuid primary key default gen_random_uuid(),
  scheduled_slot text not null check (scheduled_slot ~ '^\d{4}-\d{2}-\d{2}$'),
  actor text not null,
  status text not null check (status in ('pending','running','completed','partial','failed')),
  rule_snapshot jsonb not null check (jsonb_typeof(rule_snapshot) = 'array'),
  counts jsonb not null default '{}'::jsonb check (jsonb_typeof(counts) = 'object'),
  started_at timestamptz not null default now(),
  finished_at timestamptz,
  updated_at timestamptz not null default now(),
  unique (scheduled_slot)
);

alter table public.financial_reconciliation_automatic_runs
  add column if not exists batch_id uuid,
  add column if not exists batch_rule_key text,
  add column if not exists batch_rule_position integer,
  add column if not exists batch_rule_count integer;

do $migration$
begin
  if not exists (
    select 1
    from pg_constraint constraint_row
    where constraint_row.conrelid = 'public.financial_reconciliation_automatic_runs'::regclass
      and constraint_row.conname = 'financial_reconciliation_automatic_runs_batch_id_fkey'
  ) then
    alter table public.financial_reconciliation_automatic_runs
      add constraint financial_reconciliation_automatic_runs_batch_id_fkey
      foreign key (batch_id)
      references public.financial_reconciliation_automatic_batches(id);
  end if;
  if not exists (
    select 1
    from pg_constraint constraint_row
    where constraint_row.conrelid = 'public.financial_reconciliation_automatic_runs'::regclass
      and constraint_row.conname = 'financial_reconciliation_automatic_runs_batch_rule_position_check'
  ) then
    alter table public.financial_reconciliation_automatic_runs
      add constraint financial_reconciliation_automatic_runs_batch_rule_position_check
      check (batch_rule_position is null or batch_rule_position > 0);
  end if;
  if not exists (
    select 1
    from pg_constraint constraint_row
    where constraint_row.conrelid = 'public.financial_reconciliation_automatic_runs'::regclass
      and constraint_row.conname = 'financial_reconciliation_automatic_runs_batch_rule_count_check'
  ) then
    alter table public.financial_reconciliation_automatic_runs
      add constraint financial_reconciliation_automatic_runs_batch_rule_count_check
      check (batch_rule_count is null or batch_rule_count > 0);
  end if;
end
$migration$;

update public.financial_reconciliation_automatic_runs run
set status = 'failed',
    error_summary = 'Analysis must be restarted after the 90-day performance upgrade.',
    analysis_error_code = 'analysis_upgrade_restart_required',
    analysis_error_at = now(),
    finished_at = now(),
    updated_at = now()
where run.trigger = 'scheduled'
  and run.batch_id is null
  and run.finished_at is null;

insert into public.financial_reconciliation_automatic_batches (
  scheduled_slot, actor, status, rule_snapshot, counts,
  started_at, finished_at, updated_at
)
select
  run.scheduled_slot,
  run.actor,
  case
    when run.status = 'completed' then 'completed'
    when run.status = 'partial' then 'partial'
    else 'failed'
  end,
  run.definition_config_snapshot,
  run.counts,
  run.started_at,
  coalesce(run.finished_at, now()),
  run.updated_at
from public.financial_reconciliation_automatic_runs run
where run.trigger = 'scheduled'
  and run.scheduled_slot is not null
  and run.batch_id is null
on conflict (scheduled_slot) do nothing;

update public.financial_reconciliation_automatic_runs run
set batch_id = batch.id
from public.financial_reconciliation_automatic_batches batch
where run.trigger = 'scheduled'
  and run.scheduled_slot = batch.scheduled_slot
  and run.batch_id is null;

do $migration$
declare
  v_constraint record;
begin
  for v_constraint in
    select constraint_row.conname
    from pg_constraint constraint_row
    where constraint_row.conrelid = 'public.financial_reconciliation_automatic_runs'::regclass
      and constraint_row.contype = 'c'
      and pg_get_constraintdef(constraint_row.oid) ilike '%client_request_id%'
      and pg_get_constraintdef(constraint_row.oid) ilike '%scheduled_slot%'
      and pg_get_constraintdef(constraint_row.oid) ilike '%trigger%'
  loop
    execute format(
      'alter table public.financial_reconciliation_automatic_runs drop constraint %I',
      v_constraint.conname
    );
  end loop;

  alter table public.financial_reconciliation_automatic_runs
    add constraint financial_reconciliation_automatic_runs_origin_check check (
      (
        trigger = 'manual'
        and client_request_id is not null
        and scheduled_slot is null
        and batch_id is null
        and batch_rule_key is null
        and batch_rule_position is null
        and batch_rule_count is null
      )
      or
      (
        trigger = 'scheduled'
        and scope = 'batch'
        and client_request_id is null
        and scheduled_slot is not null
        and batch_id is not null
        and batch_rule_key is null
        and batch_rule_position is null
        and batch_rule_count is null
      )
      or
      (
        trigger = 'scheduled'
        and scope = 'rule'
        and client_request_id is null
        and scheduled_slot is not null
        and batch_id is not null
        and batch_rule_key is not null
        and batch_rule_position is not null
        and batch_rule_count is not null
        and batch_rule_position <= batch_rule_count
        and jsonb_array_length(definition_config_snapshot) = 1
        and definition_config_snapshot->0->>'ruleKey' = batch_rule_key
      )
    ) not valid;
end
$migration$;

alter table public.financial_reconciliation_automatic_runs
  validate constraint financial_reconciliation_automatic_runs_origin_check;

drop index if exists public.financial_reconciliation_automatic_runs_scheduled_slot_uidx;
create unique index if not exists financial_reconciliation_automatic_runs_legacy_scheduled_slot_uidx
  on public.financial_reconciliation_automatic_runs (scheduled_slot)
  where scheduled_slot is not null and batch_id is null;
create unique index if not exists financial_reconciliation_automatic_runs_batch_position_uidx
  on public.financial_reconciliation_automatic_runs (batch_id, batch_rule_position)
  where batch_id is not null and batch_rule_position is not null;
create unique index if not exists financial_reconciliation_automatic_runs_batch_rule_uidx
  on public.financial_reconciliation_automatic_runs (batch_id, batch_rule_key)
  where batch_id is not null and batch_rule_key is not null;

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
    'items', proposal.items,
    'evidence', proposal.evidence,
    'candidateGroups', proposal.candidate_groups,
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

create or replace function public.financial_reconciliation_refresh_automatic_batch(p_batch_id uuid)
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
      'startedAt', v_batch.started_at,
      'finishedAt', v_batch.finished_at,
      'updatedAt', v_batch.updated_at
    );
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
      'bases', coalesce(sum(coalesce((run.counts->>'bases')::bigint, 0)), 0),
      'proposed', coalesce(sum(coalesce((run.counts->>'proposed')::bigint, 0)), 0),
      'ambiguous', coalesce(sum(coalesce((run.counts->>'ambiguous')::bigint, 0)), 0),
      'skipped', coalesce(sum(coalesce((run.counts->>'skipped')::bigint, 0)), 0),
      'deselected', coalesce(sum(coalesce((run.counts->>'deselected')::bigint, 0)), 0),
      'completed', coalesce(sum(coalesce((run.counts->>'completed')::bigint, 0)), 0),
      'stale', coalesce(sum(coalesce((run.counts->>'stale')::bigint, 0)), 0),
      'failed', coalesce(sum(coalesce((run.counts->>'failed')::bigint, 0)), 0)
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

  v_status := case
    when v_child_count = 0 then 'pending'
    when v_child_count < v_rule_count or v_unfinished_children > 0 then 'running'
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
    'startedAt', v_batch.started_at,
    'finishedAt', v_batch.finished_at,
    'updatedAt', v_batch.updated_at
  );
end
$$;

create or replace function public.financial_reconciliation_refresh_automatic_batch_from_run()
returns trigger
language plpgsql
security definer set search_path = public, pg_temp
as $$
begin
  if new.batch_id is not null and new.scope = 'rule' then
    perform public.financial_reconciliation_refresh_automatic_batch(new.batch_id);
  end if;
  return new;
end
$$;

drop trigger if exists financial_reconciliation_refresh_automatic_batch_trigger
  on public.financial_reconciliation_automatic_runs;
create trigger financial_reconciliation_refresh_automatic_batch_trigger
after insert or update of status, finished_at, counts, batch_id
on public.financial_reconciliation_automatic_runs
for each row execute function public.financial_reconciliation_refresh_automatic_batch_from_run();

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
  v_contract jsonb;
  v_run_id uuid;
begin
  if p_now is null then raise exception 'Schedule time is required.'; end if;
  if nullif(trim(coalesce(p_actor, '')), '') is null then raise exception 'Actor is required.'; end if;

  lock table public.financial_reconciliation_source_rules in share row exclusive mode;
  lock table public.financial_reconciliation_automatic_rule_configs in share row exclusive mode;
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
      return jsonb_build_object('claimed', false, 'reason', 'before_scheduled_time');
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
        return jsonb_build_object('claimed', false, 'reason', 'no_enabled_rules');
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
      cross join lateral (
        select public.financial_reconciliation_automatic_rule_contract(
          config.rule_key, config.rule_version
        ) as value
      ) contract
      join public.financial_reconciliation_source_rules source_rule
        on source_rule.base_source_type = definition.base_source_type
       and source_rule.matching_source_type = destination.source_type
      where config.enabled
        and config.include_in_scheduled_batch
        and config.max_difference_days between 0 and 90
        and jsonb_array_length(definition.destination_source_types) = 1
        and contract.value is not null
        and contract.value->>'destinationSourceType' = destination.source_type
        and source_rule.operator in ('+', '-');

      if jsonb_array_length(v_snapshot) <> v_enabled_rule_count then
        return jsonb_build_object('claimed', false, 'reason', 'unsupported_rule_set');
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
      'run', public.financial_reconciliation_automatic_progress_or_run(v_run_id)
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

  if jsonb_typeof(v_selected_rule) <> 'object'
    or coalesce(v_selected_rule->>'ruleVersion', '') !~ '^[0-9]+$'
    or coalesce(v_selected_rule->>'priority', '') !~ '^[0-9]+$'
    or coalesce(v_selected_rule->>'differenceAllowed', '') !~ '^[0-9]+(\.[0-9]+)?$'
    or coalesce(v_selected_rule->>'maxDifferenceDays', '') !~ '^[0-9]+$'
    or jsonb_typeof(v_selected_rule->'definition') is distinct from 'object'
    or coalesce(v_selected_rule->>'operator', '') not in ('+', '-') then
    raise exception 'Automatic scheduled batch snapshot is invalid.';
  end if;
  v_selected_rule_key := v_selected_rule->>'ruleKey';
  v_selected_rule_version := (v_selected_rule->>'ruleVersion')::integer;
  v_contract := public.financial_reconciliation_automatic_rule_contract(
    v_selected_rule_key, v_selected_rule_version
  );
  if v_contract is null
    or v_selected_rule->>'destinationSourceType' is distinct from
      v_contract->>'destinationSourceType' then
    raise exception 'Automatic scheduled batch snapshot is invalid.';
  end if;

  insert into public.financial_reconciliation_automatic_runs (
    trigger, scope, actor, scheduled_slot, definition_config_snapshot,
    analysis_processed, analysis_total,
    batch_id, batch_rule_key, batch_rule_position, batch_rule_count
  ) values (
    'scheduled', 'rule', v_batch.actor, v_batch.scheduled_slot,
    jsonb_build_array(v_selected_rule), 0,
    public.financial_reconciliation_automatic_base_count(
      v_selected_rule_key, v_selected_rule_version
    ),
    v_batch.id, v_selected_rule_key, v_selected_position, v_batch_rule_count
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
      'run', public.financial_reconciliation_automatic_progress_or_run(v_run_id)
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

create or replace function public.get_financial_reconciliation_automation_settings()
returns jsonb
language plpgsql
stable
security definer set search_path = public, pg_temp
as $$
declare
  v_schedule jsonb;
  v_rules jsonb;
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
    'updatedBy', config.updated_by,
    'updatedAt', config.updated_at
  ) order by config.priority, definition.rule_key), '[]'::jsonb)
  into v_rules
  from public.financial_reconciliation_automatic_rule_configs config
  join public.financial_reconciliation_automatic_rule_definitions definition
    on definition.rule_key = config.rule_key
   and definition.version = config.rule_version;

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

alter table public.financial_reconciliation_automatic_batches enable row level security;
revoke all on table public.financial_reconciliation_automatic_batches
  from public, anon, authenticated, service_role;

alter table public.financial_reconciliation_cgd_credit_card_match_search enable row level security;
revoke all on table public.financial_reconciliation_cgd_credit_card_match_search
  from public, anon, authenticated, service_role;
grant select on table public.financial_reconciliation_cgd_credit_card_match_search to service_role;

revoke all on function public.financial_reconciliation_sync_cgd_credit_card_match_search()
  from public, anon, authenticated;

revoke all on function public.financial_reconciliation_automatic_rule_contract(text,integer)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_automatic_bank_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_automatic_credit_card_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_automatic_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_automatic_base_page(text,integer,date,uuid,integer)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_automatic_base_count(text,integer)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_automatic_candidate_page(text,integer,numeric,integer,date,uuid,integer)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_automatic_single_base_candidates(text,integer,numeric,integer,uuid)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_finalize_automatic_analysis(uuid)
  from public, anon, authenticated, service_role;
revoke all on function public.continue_financial_reconciliation_automatic_analysis(uuid,text)
  from public, anon, authenticated, service_role;
revoke all on function public.create_financial_reconciliation_automatic_analysis(text[],text,text,uuid)
  from public, anon, authenticated, service_role;
revoke all on function public.get_financial_reconciliation_automatic_active_run(text)
  from public, anon, authenticated, service_role;
revoke all on function public.continue_financial_reconciliation_automatic_oldest_analysis(text)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_automatic_lock_destination_items(text,jsonb)
  from public, anon, authenticated, service_role;
revoke all on function public.execute_financial_reconciliation_automatic_proposal(uuid,text)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_refresh_automatic_batch(uuid)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_refresh_automatic_batch_from_run()
  from public, anon, authenticated, service_role;
revoke all on function public.claim_financial_reconciliation_automatic_schedule(timestamptz,text)
  from public, anon, authenticated, service_role;
revoke all on function public.get_financial_reconciliation_automatic_run(uuid)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_automatic_progress_or_run(uuid)
  from public, anon, authenticated, service_role;
revoke all on function public.get_financial_reconciliation_automation_settings()
  from public, anon, authenticated, service_role;
revoke all on function public.replace_financial_reconciliation_source_rules(jsonb)
  from public, anon, authenticated, service_role;

grant execute on function public.financial_reconciliation_automatic_rule_contract(text,integer)
  to service_role;
grant execute on function public.financial_reconciliation_automatic_bank_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  to service_role;
grant execute on function public.financial_reconciliation_automatic_credit_card_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  to service_role;
grant execute on function public.financial_reconciliation_automatic_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  to service_role;
grant execute on function public.financial_reconciliation_automatic_base_page(text,integer,date,uuid,integer)
  to service_role;
grant execute on function public.financial_reconciliation_automatic_base_count(text,integer)
  to service_role;
grant execute on function public.financial_reconciliation_automatic_candidate_page(text,integer,numeric,integer,date,uuid,integer)
  to service_role;
grant execute on function public.financial_reconciliation_automatic_single_base_candidates(text,integer,numeric,integer,uuid)
  to service_role;
grant execute on function public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer)
  to service_role;
grant execute on function public.continue_financial_reconciliation_automatic_analysis(uuid,text)
  to service_role;
grant execute on function public.create_financial_reconciliation_automatic_analysis(text[],text,text,uuid)
  to service_role;
grant execute on function public.get_financial_reconciliation_automatic_active_run(text)
  to service_role;
grant execute on function public.continue_financial_reconciliation_automatic_oldest_analysis(text)
  to service_role;
grant execute on function public.execute_financial_reconciliation_automatic_proposal(uuid,text)
  to service_role;
grant execute on function public.financial_reconciliation_refresh_automatic_batch(uuid)
  to service_role;
grant execute on function public.claim_financial_reconciliation_automatic_schedule(timestamptz,text)
  to service_role;
grant execute on function public.get_financial_reconciliation_automatic_run(uuid)
  to service_role;
grant execute on function public.financial_reconciliation_automatic_progress_or_run(uuid)
  to service_role;
grant execute on function public.get_financial_reconciliation_automation_settings()
  to service_role;
grant execute on function public.replace_financial_reconciliation_source_rules(jsonb)
  to service_role;

notify pgrst, 'reload schema';
