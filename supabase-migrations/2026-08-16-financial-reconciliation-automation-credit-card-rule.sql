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

notify pgrst, 'reload schema';
