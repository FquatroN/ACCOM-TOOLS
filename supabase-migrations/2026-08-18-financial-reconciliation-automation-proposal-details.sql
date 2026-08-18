create or replace function public.financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(
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
security definer set search_path = public, pg_temp
as $$
  with bases as materialized (
    select
      document.id,
      document.document_date,
      document.amount,
      document.doc_number,
      document.description,
      document.supplier_name,
      document.supplier_nif,
      round(document.amount * 100)::bigint as base_amount_cents
    from public.financial_documents document
    where p_rule_key = 'financial_documents_cgd_bank_statement_amount_only'
      and p_rule_version = 1
      and p_difference_allowed = 0
      and p_max_difference_days between 0 and 90
      and document.id = any(coalesce(p_base_ids, array[]::uuid[]))
      and document.fat = 'S'
      and document.payment = 'Banco'
      and document.document_date is not null
      and document.document_date >= date '2026-01-01'
      and document.amount is not null
      and not exists (
        select 1
        from public.financial_reconciliation_items item
        where item.source_type = 'financial_documents'
          and item.source_id = document.id
      )
  ), qualified as materialized (
    select
      base.id as base_id,
      base.document_date as base_date,
      base.base_amount_cents,
      jsonb_build_object(
        'sourceType', 'financial_documents',
        'sourceId', base.id,
        'sourceDate', base.document_date,
        'amount', base.amount,
        'docNumber', base.doc_number,
        'description', base.description,
        'supplierName', base.supplier_name,
        'supplierNif', base.supplier_nif
      ) as base_snapshot,
      destination.source_id,
      destination.source_date,
      destination.amount,
      destination.description,
      destination.destination_amount_cents,
      destination.candidate_ordinal,
      destination.total_candidate_count
    from bases base
    left join lateral (
      select
        bank.id as source_id,
        bank.data as source_date,
        bank.montante as amount,
        bank.descritivo as description,
        round(bank.montante * 100)::bigint as destination_amount_cents,
        row_number() over (order by bank.data, bank.id) as candidate_ordinal,
        (count(*) over ())::integer as total_candidate_count
      from public.import_cgd_extrato_ordem bank
      where bank.montante = -base.amount
        and round(bank.montante * 100)::bigint = -base.base_amount_cents
        and bank.data between base.document_date - p_max_difference_days
                          and base.document_date + p_max_difference_days
        and bank.data >= date '2026-01-01'
        and bank.data is not null
        and bank.montante is not null
        and not exists (
          select 1
          from public.financial_reconciliation_items item
          where item.source_type = 'import_cgd_extrato_ordem'
            and item.source_id = bank.id
        )
      order by bank.data, bank.id
      limit 13
    ) destination on true
  ), grouped as (
    select
      base_id,
      base_date,
      base_snapshot,
      coalesce(jsonb_agg(jsonb_build_object(
        'sourceType', 'import_cgd_extrato_ordem',
        'sourceId', source_id,
        'sourceDate', source_date,
        'amount', amount,
        'description', description,
        'evidence', jsonb_build_object(
          'amount', jsonb_build_object(
            'baseAmountCents', base_amount_cents,
            'destinationAmountCents', destination_amount_cents,
            'signedDifferenceCents', base_amount_cents + destination_amount_cents,
            'matched', base_amount_cents + destination_amount_cents = 0
          ),
          'date', jsonb_build_object(
            'distanceDays', abs(source_date - base_date),
            'maxDifferenceDays', p_max_difference_days,
            'matched', source_date between base_date - p_max_difference_days
                                      and base_date + p_max_difference_days
          )
        )
      ) order by source_date, source_id) filter (
        where source_id is not null and candidate_ordinal <= 12
      ), '[]'::jsonb) as candidates,
      coalesce(max(total_candidate_count), 0)::integer as candidate_count
    from qualified
    group by base_id, base_date, base_snapshot
  )
  select base_id, base_date, base_snapshot, candidates, candidate_count
  from grouped
  order by base_date, base_id
$$;

create or replace function public.financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(
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
security definer set search_path = public, pg_temp
as $$
  with bases as materialized (
    select
      document.id,
      document.document_date,
      document.amount,
      document.doc_number,
      document.description,
      document.supplier_name,
      document.supplier_nif,
      round(document.amount * 100)::bigint as base_amount_cents
    from public.financial_documents document
    where p_rule_key = 'financial_documents_cgd_credit_card_amount_only'
      and p_rule_version = 1
      and p_difference_allowed = 0
      and p_max_difference_days between 0 and 90
      and document.id = any(coalesce(p_base_ids, array[]::uuid[]))
      and document.fat = 'S'
      and document.payment = 'Visa'
      and document.document_date is not null
      and document.document_date >= date '2026-01-01'
      and document.amount is not null
      and not exists (
        select 1
        from public.financial_reconciliation_items item
        where item.source_type = 'financial_documents'
          and item.source_id = document.id
      )
  ), qualified as materialized (
    select
      base.id as base_id,
      base.document_date as base_date,
      base.base_amount_cents,
      jsonb_build_object(
        'sourceType', 'financial_documents',
        'sourceId', base.id,
        'sourceDate', base.document_date,
        'amount', base.amount,
        'docNumber', base.doc_number,
        'description', base.description,
        'supplierName', base.supplier_name,
        'supplierNif', base.supplier_nif
      ) as base_snapshot,
      destination.source_id,
      destination.source_date,
      destination.amount,
      destination.description,
      destination.destination_amount_cents,
      destination.candidate_ordinal,
      destination.total_candidate_count
    from bases base
    left join lateral (
      select
        card.id as source_id,
        card.data as source_date,
        card.valor as amount,
        card.descricao as description,
        round(card.valor * 100)::bigint as destination_amount_cents,
        row_number() over (order by card.data, card.id) as candidate_ordinal,
        (count(*) over ())::integer as total_candidate_count
      from public.import_cgd_cartao_credito card
      where card.valor = -base.amount
        and round(card.valor * 100)::bigint = -base.base_amount_cents
        and card.data between base.document_date - p_max_difference_days
                          and base.document_date + p_max_difference_days
        and card.data >= date '2026-01-01'
        and card.data is not null
        and card.valor is not null
        and not exists (
          select 1
          from public.financial_reconciliation_items item
          where item.source_type = 'import_cgd_cartao_credito'
            and item.source_id = card.id
        )
      order by card.data, card.id
      limit 13
    ) destination on true
  ), grouped as (
    select
      base_id,
      base_date,
      base_snapshot,
      coalesce(jsonb_agg(jsonb_build_object(
        'sourceType', 'import_cgd_cartao_credito',
        'sourceId', source_id,
        'sourceDate', source_date,
        'amount', amount,
        'description', description,
        'evidence', jsonb_build_object(
          'amount', jsonb_build_object(
            'baseAmountCents', base_amount_cents,
            'destinationAmountCents', destination_amount_cents,
            'signedDifferenceCents', base_amount_cents + destination_amount_cents,
            'matched', base_amount_cents + destination_amount_cents = 0
          ),
          'date', jsonb_build_object(
            'distanceDays', abs(source_date - base_date),
            'maxDifferenceDays', p_max_difference_days,
            'matched', source_date between base_date - p_max_difference_days
                                      and base_date + p_max_difference_days
          )
        )
      ) order by source_date, source_id) filter (
        where source_id is not null and candidate_ordinal <= 12
      ), '[]'::jsonb) as candidates,
      coalesce(max(total_candidate_count), 0)::integer as candidate_count
    from qualified
    group by base_id, base_date, base_snapshot
  )
  select base_id, base_date, base_snapshot, candidates, candidate_count
  from grouped
  order by base_date, base_id
$$;

do $migration$
begin
  alter table public.financial_reconciliation_automatic_proposals
    disable trigger financial_reconciliation_automatic_proposal_snapshot_immutable;

  begin
    with eligible as materialized (
  select
    proposal.id,
    proposal.base_snapshot,
    proposal.items,
    proposal.candidate_groups
  from public.financial_reconciliation_automatic_proposals proposal
  where proposal.rule_key in (
      'financial_documents_cgd_bank_statement_amount_only',
      'financial_documents_cgd_credit_card_amount_only'
    )
    and proposal.rule_version = 1
    and proposal.status <> 'completed'
    and proposal.reconciliation_id is null
    and proposal.completed_at is null
    and exists (
      select 1
      from public.financial_reconciliation_automatic_runs run
      where run.id = proposal.run_id
        and run.finished_at is null
    )
), enriched_bases as (
  select
    eligible.id,
    coalesce((
      select jsonb_build_object(
        'docNumber', document.doc_number,
        'description', document.description,
        'supplierName', document.supplier_name,
        'supplierNif', document.supplier_nif
      ) || eligible.base_snapshot
      from public.financial_documents document
      where eligible.base_snapshot->>'sourceType' = 'financial_documents'
        and document.id::text = eligible.base_snapshot->>'sourceId'
    ), eligible.base_snapshot) as base_snapshot
  from eligible
), enriched_items as (
  select
    eligible.id,
    coalesce((
      select jsonb_agg(
        case item.value->>'sourceType'
          when 'import_cgd_extrato_ordem' then coalesce((
            select jsonb_build_object('description', bank.descritivo) || item.value
            from public.import_cgd_extrato_ordem bank
            where bank.id::text = item.value->>'sourceId'
          ), item.value)
          when 'import_cgd_cartao_credito' then coalesce((
            select jsonb_build_object('description', card.descricao) || item.value
            from public.import_cgd_cartao_credito card
            where card.id::text = item.value->>'sourceId'
          ), item.value)
          else item.value
        end
        order by item.ordinality
      )
      from jsonb_array_elements(eligible.items) with ordinality item(value, ordinality)
    ), '[]'::jsonb) as items
  from eligible
), enriched_candidate_groups as (
  select
    eligible.id,
    coalesce((
      select jsonb_agg(
        case
          when jsonb_typeof(candidate_group.value) = 'array' then coalesce((
            select jsonb_agg(
              case nested_item.value->>'sourceType'
                when 'import_cgd_extrato_ordem' then coalesce((
                  select jsonb_build_object('description', bank.descritivo) || nested_item.value
                  from public.import_cgd_extrato_ordem bank
                  where bank.id::text = nested_item.value->>'sourceId'
                ), nested_item.value)
                when 'import_cgd_cartao_credito' then coalesce((
                  select jsonb_build_object('description', card.descricao) || nested_item.value
                  from public.import_cgd_cartao_credito card
                  where card.id::text = nested_item.value->>'sourceId'
                ), nested_item.value)
                else nested_item.value
              end
              order by nested_item.ordinality
            )
            from jsonb_array_elements(candidate_group.value)
              with ordinality nested_item(value, ordinality)
          ), '[]'::jsonb)
          else case candidate_group.value->>'sourceType'
            when 'import_cgd_extrato_ordem' then coalesce((
              select jsonb_build_object('description', bank.descritivo) || candidate_group.value
              from public.import_cgd_extrato_ordem bank
              where bank.id::text = candidate_group.value->>'sourceId'
            ), candidate_group.value)
            when 'import_cgd_cartao_credito' then coalesce((
              select jsonb_build_object('description', card.descricao) || candidate_group.value
              from public.import_cgd_cartao_credito card
              where card.id::text = candidate_group.value->>'sourceId'
            ), candidate_group.value)
            else candidate_group.value
          end
        end
        order by candidate_group.ordinality
      )
      from jsonb_array_elements(eligible.candidate_groups)
        with ordinality candidate_group(value, ordinality)
    ), '[]'::jsonb) as candidate_groups
  from eligible
), enriched as (
  select
    eligible.id,
    enriched_bases.base_snapshot,
    enriched_items.items,
    enriched_candidate_groups.candidate_groups
  from eligible
  join enriched_bases on enriched_bases.id = eligible.id
  join enriched_items on enriched_items.id = eligible.id
  join enriched_candidate_groups on enriched_candidate_groups.id = eligible.id
)
    update public.financial_reconciliation_automatic_proposals proposal
    set base_snapshot = enriched.base_snapshot,
        items = enriched.items,
        candidate_groups = enriched.candidate_groups,
        updated_at = now()
    from enriched
    where proposal.id = enriched.id
      and (
        proposal.base_snapshot is distinct from enriched.base_snapshot
        or proposal.items is distinct from enriched.items
        or proposal.candidate_groups is distinct from enriched.candidate_groups
      );
  exception when others then
    alter table public.financial_reconciliation_automatic_proposals
      enable trigger financial_reconciliation_automatic_proposal_snapshot_immutable;
    raise;
  end;

  alter table public.financial_reconciliation_automatic_proposals
    enable trigger financial_reconciliation_automatic_proposal_snapshot_immutable;
end
$migration$;

revoke all on function public.financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  from public, anon, authenticated;
revoke all on function public.financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  from public, anon, authenticated;
grant execute on function public.financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  to service_role;
grant execute on function public.financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  to service_role;

notify pgrst, 'reload schema';
