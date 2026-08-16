do $migration$
declare
  v_definition jsonb := $json$
  {
    "baseSourceType": "financial_documents",
    "destinationSourceTypes": ["import_cgd_extrato_ordem"],
    "baseEligibility": {
      "payment": {
        "operator": "exact_text_equal",
        "value": "Banco",
        "caseSensitive": true,
        "trim": false
      }
    },
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
  $json$::jsonb;
  v_logic text := 'Payment must equal exactly Banco. A bank candidate must match at least one of three OR identity branches: compact document-number containment, document-description similarity, or supplier-to-bank-description word similarity. A base record is executable only when exactly one complete destination combination is valid; multiple combinations are reported as ambiguous and are never selected automatically.';
begin
  insert into public.financial_reconciliation_automatic_rule_definitions (
    rule_key, version, display_name, base_source_type,
    destination_source_types, logic_description, definition
  ) values (
    'financial_documents_cgd_bank_statement',
    2,
    'Financial Documents to CGD Bank Statement',
    'financial_documents',
    '["import_cgd_extrato_ordem"]'::jsonb,
    v_logic,
    v_definition
  ) on conflict (rule_key, version) do nothing;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_definitions definition
    where definition.rule_key = 'financial_documents_cgd_bank_statement'
      and definition.version = 2
      and definition.display_name = 'Financial Documents to CGD Bank Statement'
      and definition.base_source_type = 'financial_documents'
      and definition.destination_source_types = '["import_cgd_extrato_ordem"]'::jsonb
      and definition.logic_description = v_logic
      and definition.definition = v_definition
  ) then
    raise exception 'Managed automatic reconciliation rule version 2 differs from the expected immutable definition.';
  end if;

  update public.financial_reconciliation_automatic_rule_configs
  set rule_version = 2,
      updated_at = now()
  where rule_key = 'financial_documents_cgd_bank_statement'
    and rule_version = 1;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs
    where rule_key = 'financial_documents_cgd_bank_statement'
      and rule_version = 2
  ) then
    raise exception 'Managed automatic reconciliation configuration could not be moved to version 2.';
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
      and p_rule_version = 2
      and d.fat = 'S'
      and d.payment = 'Banco'
      and d.document_date >= date '2026-01-01'
      and not exists (
        select 1
        from public.financial_reconciliation_items i
        where i.source_type = 'financial_documents'
          and i.source_id = d.id
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
    left join lateral (
      select
        bank.id,
        bank.data,
        bank.montante,
        bank.descritivo,
        public.financial_reconciliation_match_normalize(bank.descritivo) as normalized_bank_description
      from public.import_cgd_extrato_ordem bank
      where bank.data between d.document_date - p_max_difference_days and d.document_date + p_max_difference_days
        and bank.data >= date '2026-01-01'
        and bank.montante is not null
        and not exists (
          select 1
          from public.financial_reconciliation_items i
          where i.source_type = 'import_cgd_extrato_ordem'
            and i.source_id = bank.id
        )
    ) b on true
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

revoke all on function public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer)
  from public, anon, authenticated;
grant execute on function public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer)
  to service_role;

do $migration$
declare
  v_definition text;
  v_old_version text := 'or v_proposal.rule_version <> 1';
  v_new_version text := 'or v_proposal.rule_version <> 2';
  v_old_comment text := 'v_comment := ''Automatically completed by rule Financial Documents to CGD Bank Statement v1; difference ''';
  v_new_comment text := 'v_comment := ''Automatically completed by rule Financial Documents to CGD Bank Statement v'' || v_proposal.rule_version::text || ''; difference ''';
begin
  select pg_get_functiondef(
    'public.execute_financial_reconciliation_automatic_proposal(uuid,text)'::regprocedure
  ) into strict v_definition;

  if strpos(v_definition, v_old_version) > 0 then
    v_definition := replace(v_definition, v_old_version, v_new_version);
  elsif strpos(v_definition, v_new_version) = 0 then
    raise exception 'Unexpected automatic proposal execution version guard.';
  end if;

  if strpos(v_definition, v_old_comment) > 0 then
    v_definition := replace(v_definition, v_old_comment, v_new_comment);
  elsif strpos(v_definition, v_new_comment) = 0 then
    raise exception 'Unexpected automatic proposal execution comment definition.';
  end if;

  execute v_definition;
end
$migration$;

revoke all on function public.execute_financial_reconciliation_automatic_proposal(uuid,text)
  from public, anon, authenticated;
grant execute on function public.execute_financial_reconciliation_automatic_proposal(uuid,text)
  to service_role;

notify pgrst, 'reload schema';
