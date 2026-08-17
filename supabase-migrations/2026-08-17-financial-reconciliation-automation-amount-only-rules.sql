do $migration$
declare
  v_bank_definition jsonb := $json$
  {
    "baseSourceType":"financial_documents",
    "destinationSourceTypes":["import_cgd_extrato_ordem"],
    "baseEligibility":{"payment":{"operator":"exact_text_equal","value":"Banco","caseSensitive":true,"trim":false}},
    "matchingMode":"amount_only_one_to_one",
    "fixedDifferenceAllowed":0,
    "maxDifferenceDays":{"minimum":0,"maximum":90,"default":1},
    "maxDestinationRecords":1
  }
  $json$::jsonb;
  v_card_definition jsonb := $json$
  {
    "baseSourceType":"financial_documents",
    "destinationSourceTypes":["import_cgd_cartao_credito"],
    "baseEligibility":{"payment":{"operator":"exact_text_equal","value":"Visa","caseSensitive":true,"trim":false}},
    "matchingMode":"amount_only_one_to_one",
    "fixedDifferenceAllowed":0,
    "maxDifferenceDays":{"minimum":0,"maximum":90,"default":1},
    "maxDestinationRecords":1
  }
  $json$::jsonb;
begin
  insert into public.financial_reconciliation_automatic_rule_definitions (
    rule_key, version, display_name, base_source_type,
    destination_source_types, logic_description, definition
  ) values
  (
    'financial_documents_cgd_bank_statement_amount_only',
    1,
    'Financial Documents to CGD Bank Account – AMOUNT ONLY',
    'financial_documents',
    '["import_cgd_extrato_ordem"]'::jsonb,
    'Payment must equal exactly Banco. Exactly one CGD Bank Account destination record must make the signed amounts sum to zero within the inclusive configured date window; identity fields and similarity are not used.',
    v_bank_definition
  ),
  (
    'financial_documents_cgd_credit_card_amount_only',
    1,
    'Financial Documents to CGD Credit Card – AMOUNT ONLY',
    'financial_documents',
    '["import_cgd_cartao_credito"]'::jsonb,
    'Payment must equal exactly Visa. Exactly one CGD Credit Card destination record must make the signed amounts sum to zero within the inclusive configured date window; identity fields and similarity are not used.',
    v_card_definition
  )
  on conflict (rule_key, version) do nothing;

  if not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_definitions definition
      where definition.rule_key = 'financial_documents_cgd_bank_statement_amount_only'
        and definition.version = 1
        and definition.display_name = 'Financial Documents to CGD Bank Account – AMOUNT ONLY'
        and definition.base_source_type = 'financial_documents'
        and definition.destination_source_types = '["import_cgd_extrato_ordem"]'::jsonb
        and definition.logic_description = 'Payment must equal exactly Banco. Exactly one CGD Bank Account destination record must make the signed amounts sum to zero within the inclusive configured date window; identity fields and similarity are not used.'
        and definition.definition = v_bank_definition
    ) or not exists (
      select 1
      from public.financial_reconciliation_automatic_rule_definitions definition
      where definition.rule_key = 'financial_documents_cgd_credit_card_amount_only'
        and definition.version = 1
        and definition.display_name = 'Financial Documents to CGD Credit Card – AMOUNT ONLY'
        and definition.base_source_type = 'financial_documents'
        and definition.destination_source_types = '["import_cgd_cartao_credito"]'::jsonb
        and definition.logic_description = 'Payment must equal exactly Visa. Exactly one CGD Credit Card destination record must make the signed amounts sum to zero within the inclusive configured date window; identity fields and similarity are not used.'
        and definition.definition = v_card_definition
  ) then
    raise exception 'Managed amount-only automatic reconciliation definitions differ from the expected immutable definitions.';
  end if;

  if not exists (
      select 1
      from public.financial_reconciliation_source_rules source_rule
      where source_rule.base_source_type = 'financial_documents'
        and source_rule.matching_source_type = 'import_cgd_extrato_ordem'
        and source_rule.operator = '+'
    ) or not exists (
      select 1
      from public.financial_reconciliation_source_rules source_rule
      where source_rule.base_source_type = 'financial_documents'
        and source_rule.matching_source_type = 'import_cgd_cartao_credito'
        and source_rule.operator = '+'
  ) then
    raise exception 'Managed amount-only automatic reconciliation source rules must use operator +.';
  end if;
end
$migration$;

do $migration$
declare
  v_next_priority integer;
  v_has_bank boolean;
  v_has_card boolean;
begin
  lock table public.financial_reconciliation_automatic_rule_configs
    in share row exclusive mode;
  set constraints financial_reconciliation_automatic_rule_configs_priority_key deferred;

  if exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs config
    where (config.rule_key, config.rule_version) not in (
      ('financial_documents_cgd_bank_statement', 2),
      ('financial_documents_cgd_credit_card', 1),
      ('financial_documents_cgd_bank_statement_amount_only', 1),
      ('financial_documents_cgd_credit_card_amount_only', 1)
    )
  ) then
    raise exception 'Installed automatic reconciliation configuration is not in the managed rule/version allowlist.';
  end if;

  select exists (
    select 1 from public.financial_reconciliation_automatic_rule_configs
    where rule_key = 'financial_documents_cgd_bank_statement_amount_only'
  ) into v_has_bank;
  select exists (
    select 1 from public.financial_reconciliation_automatic_rule_configs
    where rule_key = 'financial_documents_cgd_credit_card_amount_only'
  ) into v_has_card;

  if v_has_card and not v_has_bank then
    raise exception 'Managed amount-only automatic reconciliation configurations are partially installed.';
  end if;

  if not v_has_bank then
    select coalesce(max(priority), 0) + 1
    into v_next_priority
    from public.financial_reconciliation_automatic_rule_configs;

    insert into public.financial_reconciliation_automatic_rule_configs (
      rule_key, rule_version, enabled, allow_manual_execution,
      include_in_scheduled_batch, difference_allowed, max_difference_days, priority
    ) values (
      'financial_documents_cgd_bank_statement_amount_only',
      1, false, false, false, 0.00, 1, v_next_priority
    ) on conflict (rule_key) do nothing;
  end if;

  if not v_has_card then
    select coalesce(max(priority), 0) + 1
    into v_next_priority
    from public.financial_reconciliation_automatic_rule_configs;

    insert into public.financial_reconciliation_automatic_rule_configs (
      rule_key, rule_version, enabled, allow_manual_execution,
      include_in_scheduled_batch, difference_allowed, max_difference_days, priority
    ) values (
      'financial_documents_cgd_credit_card_amount_only',
      1, false, false, false, 0.00, 1, v_next_priority
    ) on conflict (rule_key) do nothing;
  end if;

  if exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs config
    where config.rule_key in (
      'financial_documents_cgd_bank_statement_amount_only',
      'financial_documents_cgd_credit_card_amount_only'
    )
      and (config.rule_version <> 1 or config.difference_allowed <> 0)
  ) then
    raise exception 'Managed amount-only automatic reconciliation configurations require version 1 and zero difference allowed.';
  end if;
end
$migration$;

do $migration$
begin
  if not exists (
    select 1
    from pg_constraint constraint_row
    where constraint_row.conrelid = 'public.financial_reconciliation_automatic_rule_configs'::regclass
      and constraint_row.conname = 'financial_reconciliation_automatic_rule_configs_amount_only_zero_check'
  ) then
    alter table public.financial_reconciliation_automatic_rule_configs
      add constraint financial_reconciliation_automatic_rule_configs_amount_only_zero_check
      check (
        rule_key not in (
          'financial_documents_cgd_bank_statement_amount_only',
          'financial_documents_cgd_credit_card_amount_only'
        )
        or difference_allowed = 0
      ) not valid;
  end if;
end
$migration$;

alter table public.financial_reconciliation_automatic_rule_configs
  validate constraint financial_reconciliation_automatic_rule_configs_amount_only_zero_check;

create index if not exists import_cgd_extrato_ordem_reconciliation_amount_date_id_idx
  on public.import_cgd_extrato_ordem (montante, data, id)
  where montante is not null and data is not null;

create index if not exists import_cgd_cartao_credito_reconciliation_amount_date_id_idx
  on public.import_cgd_cartao_credito (valor, data, id)
  where valor is not null and data is not null;

do $migration$
begin
  if not exists (
      select 1
      from pg_index index_row
      where index_row.indexrelid = 'public.import_cgd_extrato_ordem_reconciliation_amount_date_id_idx'::regclass
        and index_row.indrelid = 'public.import_cgd_extrato_ordem'::regclass
        and index_row.indnkeyatts = 3
        and pg_get_indexdef(index_row.indexrelid, 1, true) = 'montante'
        and pg_get_indexdef(index_row.indexrelid, 2, true) = 'data'
        and pg_get_indexdef(index_row.indexrelid, 3, true) = 'id'
        and position('montante IS NOT NULL' in pg_get_expr(index_row.indpred, index_row.indrelid)) > 0
        and position('data IS NOT NULL' in pg_get_expr(index_row.indpred, index_row.indrelid)) > 0
    ) or not exists (
      select 1
      from pg_index index_row
      where index_row.indexrelid = 'public.import_cgd_cartao_credito_reconciliation_amount_date_id_idx'::regclass
        and index_row.indrelid = 'public.import_cgd_cartao_credito'::regclass
        and index_row.indnkeyatts = 3
        and pg_get_indexdef(index_row.indexrelid, 1, true) = 'valor'
        and pg_get_indexdef(index_row.indexrelid, 2, true) = 'data'
        and pg_get_indexdef(index_row.indexrelid, 3, true) = 'id'
        and position('valor IS NOT NULL' in pg_get_expr(index_row.indpred, index_row.indrelid)) > 0
        and position('data IS NOT NULL' in pg_get_expr(index_row.indpred, index_row.indrelid)) > 0
  ) then
    raise exception 'Managed amount-only destination lookup index differs from the expected definition.';
  end if;
end
$migration$;

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
    when p_rule_key = 'financial_documents_cgd_bank_statement_amount_only' and p_rule_version = 1 then
      jsonb_build_object('payment','Banco','destinationSourceType','import_cgd_extrato_ordem',
        'matchingMode','amount_only_one_to_one','maxDestinationRecords',1,'maxCandidates',12,
        'fixedDifferenceAllowed',0)
    when p_rule_key = 'financial_documents_cgd_credit_card_amount_only' and p_rule_version = 1 then
      jsonb_build_object('payment','Visa','destinationSourceType','import_cgd_cartao_credito',
        'matchingMode','amount_only_one_to_one','maxDestinationRecords',1,'maxCandidates',12,
        'fixedDifferenceAllowed',0)
    else null
  end
$$;

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
        'amount', base.amount
      ) as base_snapshot,
      destination.source_id,
      destination.source_date,
      destination.amount,
      destination.destination_amount_cents,
      destination.candidate_ordinal,
      destination.total_candidate_count
    from bases base
    left join lateral (
      select
        bank.id as source_id,
        bank.data as source_date,
        bank.montante as amount,
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
        'amount', base.amount
      ) as base_snapshot,
      destination.source_id,
      destination.source_date,
      destination.amount,
      destination.destination_amount_cents,
      destination.candidate_ordinal,
      destination.total_candidate_count
    from bases base
    left join lateral (
      select
        card.id as source_id,
        card.data as source_date,
        card.valor as amount,
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
  elsif p_rule_key = 'financial_documents_cgd_bank_statement_amount_only' and p_rule_version = 1 then
    return query
    select *
    from public.financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(
      p_rule_key, p_rule_version, p_difference_allowed, p_max_difference_days, p_base_ids
    );
  elsif p_rule_key = 'financial_documents_cgd_credit_card_amount_only' and p_rule_version = 1 then
    return query
    select *
    from public.financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(
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
    and document.document_date is not null
    and document.document_date >= date '2026-01-01'
    and (
      (p_rule_key, p_rule_version) not in (
        ('financial_documents_cgd_bank_statement_amount_only', 1),
        ('financial_documents_cgd_credit_card_amount_only', 1)
      )
      or document.amount is not null
    )
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
    and document.document_date is not null
    and document.document_date >= date '2026-01-01'
    and (
      (p_rule_key, p_rule_version) not in (
        ('financial_documents_cgd_bank_statement_amount_only', 1),
        ('financial_documents_cgd_credit_card_amount_only', 1)
      )
      or document.amount is not null
    )
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

create or replace function public.financial_reconciliation_finalize_automatic_analysis(p_run_id uuid)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
begin
  with display_source_usage as (
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
  ), display_overlapping as (
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
    join display_source_usage usage
      on usage.source_type = item->>'sourceType'
     and usage.source_id = item->>'sourceId'
    where proposal.run_id = p_run_id
      and proposal.status in ('proposed', 'ambiguous')
      and usage.base_count > 1
  ), amount_only_memberships as (
    select
      proposal.id as proposal_id,
      proposal.base_source_id,
      'import_cgd_extrato_ordem'::text as source_type,
      bank.id as source_id
    from public.financial_reconciliation_automatic_proposals proposal
    join public.financial_reconciliation_automatic_runs run
      on run.id = proposal.run_id
    cross join lateral jsonb_array_elements(run.definition_config_snapshot) snapshot(rule)
    join public.import_cgd_extrato_ordem bank
      on bank.montante = -(proposal.base_snapshot->>'amount')::numeric
     and round(bank.montante * 100)::bigint =
         -round((proposal.base_snapshot->>'amount')::numeric * 100)::bigint
     and bank.data between
         proposal.base_source_date - (snapshot.rule->>'maxDifferenceDays')::integer
         and proposal.base_source_date + (snapshot.rule->>'maxDifferenceDays')::integer
    where proposal.run_id = p_run_id
      and proposal.rule_key = 'financial_documents_cgd_bank_statement_amount_only'
      and proposal.rule_version = 1
      and proposal.status in ('proposed', 'ambiguous')
      and proposal.allowed_difference = 0
      and snapshot.rule->>'ruleKey' = proposal.rule_key
      and (snapshot.rule->>'ruleVersion')::integer = proposal.rule_version
      and (snapshot.rule->>'differenceAllowed')::numeric = 0
      and (snapshot.rule->>'maxDifferenceDays')::integer between 0 and 90
      and bank.data is not null
      and bank.data >= date '2026-01-01'
      and bank.montante is not null
      and not exists (
        select 1
        from public.financial_reconciliation_items item
        where item.source_type = 'import_cgd_extrato_ordem'
          and item.source_id = bank.id
      )
    union all
    select
      proposal.id as proposal_id,
      proposal.base_source_id,
      'import_cgd_cartao_credito'::text as source_type,
      card.id as source_id
    from public.financial_reconciliation_automatic_proposals proposal
    join public.financial_reconciliation_automatic_runs run
      on run.id = proposal.run_id
    cross join lateral jsonb_array_elements(run.definition_config_snapshot) snapshot(rule)
    join public.import_cgd_cartao_credito card
      on card.valor = -(proposal.base_snapshot->>'amount')::numeric
     and round(card.valor * 100)::bigint =
         -round((proposal.base_snapshot->>'amount')::numeric * 100)::bigint
     and card.data between
         proposal.base_source_date - (snapshot.rule->>'maxDifferenceDays')::integer
         and proposal.base_source_date + (snapshot.rule->>'maxDifferenceDays')::integer
    where proposal.run_id = p_run_id
      and proposal.rule_key = 'financial_documents_cgd_credit_card_amount_only'
      and proposal.rule_version = 1
      and proposal.status in ('proposed', 'ambiguous')
      and proposal.allowed_difference = 0
      and snapshot.rule->>'ruleKey' = proposal.rule_key
      and (snapshot.rule->>'ruleVersion')::integer = proposal.rule_version
      and (snapshot.rule->>'differenceAllowed')::numeric = 0
      and (snapshot.rule->>'maxDifferenceDays')::integer between 0 and 90
      and card.data is not null
      and card.data >= date '2026-01-01'
      and card.valor is not null
      and not exists (
        select 1
        from public.financial_reconciliation_items item
        where item.source_type = 'import_cgd_cartao_credito'
          and item.source_id = card.id
      )
  ), amount_only_source_usage as (
    select
      membership.source_type,
      membership.source_id,
      count(distinct membership.base_source_id) as base_count
    from amount_only_memberships membership
    group by membership.source_type, membership.source_id
  ), amount_only_overlapping as (
    select distinct membership.proposal_id as id
    from amount_only_memberships membership
    join amount_only_source_usage usage
      on usage.source_type = membership.source_type
     and usage.source_id = membership.source_id
    where usage.base_count > 1
  ), overlapping as (
    select id from display_overlapping
    union
    select id from amount_only_overlapping
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

create or replace function public.replace_financial_reconciliation_automation_settings(
  p_schedule jsonb,
  p_rules jsonb,
  p_actor text
)
returns jsonb
language plpgsql
security definer set search_path = public, pg_temp
as $$
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
    or exists (
      select 1 from jsonb_array_elements(p_rules) rule
      where jsonb_typeof(rule) <> 'object'
    ) then
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
       or (rule->>'max_difference_days')::numeric not between 0 and 90
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
    where rule->>'rule_key' not in (
      'financial_documents_cgd_bank_statement',
      'financial_documents_cgd_credit_card',
      'financial_documents_cgd_bank_statement_amount_only',
      'financial_documents_cgd_credit_card_amount_only'
    )
  ) then
    raise exception 'Automatic rule is invalid.';
  end if;

  if exists (
    select 1
    from jsonb_array_elements(p_rules) rule
    where (rule->>'rule_key', (rule->>'rule_version')::integer) not in (
      ('financial_documents_cgd_bank_statement', 2),
      ('financial_documents_cgd_credit_card', 1),
      ('financial_documents_cgd_bank_statement_amount_only', 1),
      ('financial_documents_cgd_credit_card_amount_only', 1)
    )
  ) then
    raise exception 'Automatic rule version is invalid.';
  end if;

  if exists (
    select 1
    from jsonb_array_elements(p_rules) rule
    where rule->>'rule_key' in (
      'financial_documents_cgd_bank_statement_amount_only',
      'financial_documents_cgd_credit_card_amount_only'
    )
      and (rule->>'difference_allowed')::numeric <> 0
  ) then
    raise exception 'Amount-only automatic rules require zero difference allowed.';
  end if;

  lock table public.financial_reconciliation_source_rules in share row exclusive mode;
  lock table public.financial_reconciliation_automatic_rule_configs in share row exclusive mode;
  lock table public.financial_reconciliation_automatic_schedule in share row exclusive mode;
  set constraints financial_reconciliation_automatic_rule_configs_priority_key deferred;

  if jsonb_array_length(p_rules) <> (
      select count(*) from public.financial_reconciliation_automatic_rule_configs
    )
    or exists (
      select 1
      from public.financial_reconciliation_automatic_rule_configs config
      where not exists (
        select 1
        from jsonb_array_elements(p_rules) rule
        where rule->>'rule_key' = config.rule_key
      )
    ) then
    raise exception 'Automation settings require every managed rule exactly once.';
  end if;

  if exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs config
    where (config.rule_key, config.rule_version) not in (
      ('financial_documents_cgd_bank_statement', 2),
      ('financial_documents_cgd_credit_card', 1),
      ('financial_documents_cgd_bank_statement_amount_only', 1),
      ('financial_documents_cgd_credit_card_amount_only', 1)
    )
  ) then
    raise exception 'Installed automatic reconciliation configuration is not in the managed rule/version allowlist.';
  end if;

  if exists (
    select 1
    from jsonb_array_elements(p_rules) rule
    join public.financial_reconciliation_automatic_rule_configs config
      on config.rule_key = rule->>'rule_key'
    where config.rule_version <> (rule->>'rule_version')::integer
  ) then
    raise exception 'Submitted automatic rule version does not match managed configuration.';
  end if;

  if exists (
    select 1
    from jsonb_array_elements(p_rules) rule
    join public.financial_reconciliation_automatic_rule_configs config
      on config.rule_key = rule->>'rule_key'
    join public.financial_reconciliation_automatic_rule_definitions definition
      on definition.rule_key = config.rule_key
     and definition.version = config.rule_version
    cross join lateral jsonb_array_elements_text(
      definition.destination_source_types
    ) destination(source_type)
    where (rule->>'enabled')::boolean
      and not exists (
        select 1
        from public.financial_reconciliation_source_rules source_rule
        where source_rule.base_source_type = definition.base_source_type
          and source_rule.matching_source_type = destination.source_type
          and source_rule.operator = '+'
      )
  ) then
    raise exception 'No fixed + directional source rule exists for an enabled automatic rule.';
  end if;

  update public.financial_reconciliation_automatic_rule_configs config
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
  where config.rule_key = input.rule_key;

  update public.financial_reconciliation_automatic_schedule
  set enabled = (p_schedule->>'enabled')::boolean,
      time_of_day = (p_schedule->>'time_of_day')::time,
      time_zone = p_schedule->>'time_zone',
      updated_by = trim(p_actor),
      updated_at = now()
  where id = true;

  return public.get_financial_reconciliation_automation_settings();
end
$$;

alter table public.financial_reconciliation_automatic_rule_definitions enable row level security;
alter table public.financial_reconciliation_automatic_rule_configs enable row level security;

revoke all on table public.financial_reconciliation_automatic_rule_definitions
  from public, anon, authenticated, service_role;
revoke all on table public.financial_reconciliation_automatic_rule_configs
  from public, anon, authenticated, service_role;

revoke all on function public.get_financial_reconciliation_automation_settings()
  from public, anon, authenticated, service_role;
revoke all on function public.replace_financial_reconciliation_automation_settings(jsonb,jsonb,text)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_finalize_automatic_analysis(uuid)
  from public, anon, authenticated, service_role;

grant execute on function public.get_financial_reconciliation_automation_settings()
  to service_role;
grant execute on function public.replace_financial_reconciliation_automation_settings(jsonb,jsonb,text)
  to service_role;
grant execute on function public.financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  to service_role;
grant execute on function public.financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  to service_role;

notify pgrst, 'reload schema';
