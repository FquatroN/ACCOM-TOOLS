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

notify pgrst, 'reload schema';
