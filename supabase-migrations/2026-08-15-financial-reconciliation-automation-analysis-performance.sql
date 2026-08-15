do $migration$
begin
  if not exists (
    select 1
    from pg_catalog.pg_index i
    join pg_catalog.pg_class t on t.oid = i.indrelid
    join pg_catalog.pg_namespace n on n.oid = t.relnamespace
    join pg_catalog.pg_class idx on idx.oid = i.indexrelid
    join pg_catalog.pg_am am on am.oid = idx.relam
    join pg_catalog.pg_attribute a on a.attrelid = t.oid and a.attnum = i.indkey[0]
    where n.nspname = 'public'
      and t.relname = 'financial_documents'
      and a.attname = 'document_date'
      and am.amname = 'btree'
      and i.indisvalid
      and i.indisready
      and i.indpred is null
      and i.indexprs is null
  ) then
    execute 'create index financial_documents_automatic_reconciliation_date_idx on public.financial_documents (document_date)';
  end if;

  if not exists (
    select 1
    from pg_catalog.pg_index i
    join pg_catalog.pg_class t on t.oid = i.indrelid
    join pg_catalog.pg_namespace n on n.oid = t.relnamespace
    join pg_catalog.pg_class idx on idx.oid = i.indexrelid
    join pg_catalog.pg_am am on am.oid = idx.relam
    join pg_catalog.pg_attribute a on a.attrelid = t.oid and a.attnum = i.indkey[0]
    where n.nspname = 'public'
      and t.relname = 'import_cgd_extrato_ordem'
      and a.attname = 'data'
      and am.amname = 'btree'
      and i.indisvalid
      and i.indisready
      and i.indpred is null
      and i.indexprs is null
  ) then
    execute 'create index import_cgd_extrato_ordem_automatic_reconciliation_date_idx on public.import_cgd_extrato_ordem (data)';
  end if;
end
$migration$;

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
  with bases as materialized (
    select
      d.id,
      d.document_date,
      d.doc_number,
      d.description,
      d.supplier_name,
      d.amount,
      public.financial_reconciliation_match_compact(d.doc_number) as compact_document_number,
      public.financial_reconciliation_match_normalize(d.description) as normalized_document_description,
      public.financial_reconciliation_match_normalize(d.supplier_name) as normalized_supplier_name
    from public.financial_documents d
    where p_rule_key = 'financial_documents_cgd_bank_statement'
      and p_rule_version = 1
      and d.fat = 'S'
      and d.document_date >= date '2026-01-01'
      and not exists (
        select 1
        from public.financial_reconciliation_items i
        where i.source_type = 'financial_documents'
          and i.source_id = d.id
      )
  ),
  bank_rows as materialized (
    select
      b.id,
      b.data,
      b.montante,
      b.descritivo,
      public.financial_reconciliation_match_normalize(b.descritivo) as normalized_bank_description
    from public.import_cgd_extrato_ordem b
    where b.data >= date '2026-01-01'
      and b.montante is not null
      and not exists (
        select 1
        from public.financial_reconciliation_items i
        where i.source_type = 'import_cgd_extrato_ordem'
          and i.source_id = b.id
      )
  ),
  qualified as materialized (
    select
      d.id as base_id,
      d.document_date as base_date,
      jsonb_build_object(
        'sourceType', 'financial_documents',
        'sourceId', d.id,
        'sourceDate', d.document_date,
        'amount', d.amount,
        'docNumber', d.doc_number,
        'description', d.description,
        'supplierName', d.supplier_name
      ) as base_snapshot,
      b.id as source_id,
      b.data as source_date,
      b.montante as amount,
      b.descritivo as description,
      b.normalized_bank_description,
      d.compact_document_number,
      d.normalized_document_description,
      d.normalized_supplier_name
    from bases d
    left join bank_rows b
      on b.data between d.document_date - p_max_difference_days and d.document_date + p_max_difference_days
  ),
  scored as materialized (
    select
      q.*,
      coalesce(
        char_length(q.compact_document_number) >= 4
          and q.source_id is not null
          and position(
            q.compact_document_number in public.financial_reconciliation_match_compact(q.description)
          ) > 0,
        false
      ) as document_number_matched,
      case
        when nullif(q.normalized_document_description, '') is null
          or nullif(q.normalized_bank_description, '') is null then 0::real
        else public.financial_reconciliation_extension_similarity(
          q.normalized_document_description,
          q.normalized_bank_description
        )
      end as description_score,
      case
        when nullif(q.normalized_supplier_name, '') is null
          or nullif(q.normalized_bank_description, '') is null then 0::real
        else public.financial_reconciliation_extension_word_similarity(
          q.normalized_supplier_name,
          q.normalized_bank_description
        )
      end as supplier_score
    from qualified q
  ),
  identity_candidates as materialized (
    select
      *,
      document_number_matched
        or description_score >= 0.60
        or supplier_score >= 0.70 as identity_matched
    from scored
  ),
  grouped as (
    select
      base_id,
      base_date,
      base_snapshot,
      coalesce(
        jsonb_agg(
          jsonb_build_object(
            'sourceType', 'import_cgd_extrato_ordem',
            'sourceId', source_id,
            'sourceDate', source_date,
            'amount', amount,
            'description', description,
            'evidence', jsonb_build_object(
              'documentNumber', jsonb_build_object(
                'matched', document_number_matched,
                'normalized', compact_document_number
              ),
              'description', jsonb_build_object(
                'matched', description_score >= 0.60,
                'score', description_score,
                'threshold', 0.60
              ),
              'supplier', jsonb_build_object(
                'matched', supplier_score >= 0.70,
                'score', supplier_score,
                'threshold', 0.70
              )
            )
          ) order by source_date, source_id
        ) filter (where identity_matched),
        '[]'::jsonb
      ) as candidates,
      count(*) filter (where identity_matched)::integer as candidate_count
    from identity_candidates
    group by base_id, base_date, base_snapshot
  )
  select
    base_id,
    base_date,
    base_snapshot,
    candidates,
    candidate_count
  from grouped
  order by base_date, base_id
$$;

revoke all on function public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer) from public, anon, authenticated;
grant execute on function public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer) to service_role;

notify pgrst, 'reload schema';
