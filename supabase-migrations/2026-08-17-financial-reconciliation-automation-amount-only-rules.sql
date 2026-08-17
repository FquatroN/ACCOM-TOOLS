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

do $migration$
begin
  if to_regprocedure(
      'public.financial_reconciliation_execute_identity_proposal(uuid,text)'
    ) is null then
    alter function public.execute_financial_reconciliation_automatic_proposal(uuid,text)
      rename to financial_reconciliation_execute_identity_proposal;
  end if;
end
$migration$;

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
  v_current_definition jsonb;
  v_current_display_name text;
  v_current_base_source_type text;
  v_current_destination_source_types jsonb;
  v_current_rule_version integer;
  v_current_difference_allowed numeric;
  v_current_max_difference_days integer;
  v_current_priority integer;
  v_current_operator text;
  v_locked_destination_count integer;
  v_base record;
  v_combination record;
  v_combination_count integer;
  v_current_evidence jsonb;
  v_action_result jsonb;
  v_reconciliation_id uuid;
  v_actual_item_count integer;
  v_expected_matching_source_rule jsonb;
  v_actual_matching_source_rule jsonb;
  v_actual_difference numeric;
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

  if (v_proposal.rule_key, v_proposal.rule_version) not in (
    ('financial_documents_cgd_bank_statement_amount_only', 1),
    ('financial_documents_cgd_credit_card_amount_only', 1)
  ) then
    return public.financial_reconciliation_execute_identity_proposal(
      p_proposal_id, p_actor
    );
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
  if v_contract is null
    or v_contract->>'matchingMode' is distinct from 'amount_only_one_to_one'
    or coalesce(v_contract->>'fixedDifferenceAllowed', '') !~ '^-?[0-9]+(\.[0-9]+)?$'
    or coalesce(v_contract->>'maxDestinationRecords', '') !~ '^[0-9]+$' then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_version_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_version_changed'
    );
  end if;
  if (v_contract->>'fixedDifferenceAllowed')::numeric <> 0
    or (v_contract->>'maxDestinationRecords')::integer <> 1 then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_version_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_version_changed'
    );
  end if;
  v_destination_source_type := v_contract->>'destinationSourceType';

  if v_proposal.base_source_type <> 'financial_documents'
    or jsonb_typeof(v_run.definition_config_snapshot) <> 'array'
    or jsonb_array_length(v_run.definition_config_snapshot) <> 1 then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_snapshot_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_snapshot_changed'
    );
  end if;
  v_rule_snapshot := v_run.definition_config_snapshot->0;
  if jsonb_typeof(v_rule_snapshot) <> 'object'
    or v_rule_snapshot->>'ruleKey' is distinct from v_proposal.rule_key
    or coalesce(v_rule_snapshot->>'ruleVersion', '') !~ '^[0-9]+$'
    or v_rule_snapshot->>'destinationSourceType' is distinct from v_destination_source_type
    or nullif(v_rule_snapshot->>'displayName', '') is null
    or jsonb_typeof(v_rule_snapshot->'definition') is distinct from 'object'
    or coalesce(v_rule_snapshot->>'maxDifferenceDays', '') !~ '^[0-9]+$'
    or coalesce(v_rule_snapshot->>'priority', '') !~ '^[0-9]+$'
    or coalesce(v_destination_source_type, '') = '' then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_snapshot_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_snapshot_changed'
    );
  end if;
  if (v_rule_snapshot->>'ruleVersion')::integer <> v_proposal.rule_version
    or (v_rule_snapshot->>'maxDifferenceDays')::integer not between 0 and 90 then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_snapshot_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_snapshot_changed'
    );
  end if;
  if coalesce(v_rule_snapshot->>'differenceAllowed', '') !~ '^[0-9]+(\.[0-9]+)?$'
  then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'tolerance_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'tolerance_changed'
    );
  end if;
  if (v_rule_snapshot->>'differenceAllowed')::numeric <> 0
    or v_proposal.allowed_difference is distinct from 0::numeric then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'tolerance_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'tolerance_changed'
    );
  end if;
  if v_rule_snapshot->>'operator' is distinct from '+' then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'operator_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'operator_changed'
    );
  end if;

  select
    definition.definition,
    definition.display_name,
    definition.base_source_type,
    definition.destination_source_types,
    config.rule_version,
    config.difference_allowed,
    config.max_difference_days,
    config.priority,
    source_rule.operator
  into
    v_current_definition,
    v_current_display_name,
    v_current_base_source_type,
    v_current_destination_source_types,
    v_current_rule_version,
    v_current_difference_allowed,
    v_current_max_difference_days,
    v_current_priority,
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
    or v_current_destination_source_types is distinct from jsonb_build_array(v_destination_source_type)
    or v_current_difference_allowed is distinct from 0::numeric
    or v_current_max_difference_days is distinct from
      (v_rule_snapshot->>'maxDifferenceDays')::integer
    or v_current_priority is distinct from (v_rule_snapshot->>'priority')::integer then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'rule_snapshot_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'rule_snapshot_changed'
    );
  end if;
  if v_current_operator is distinct from '+' then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'operator_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'operator_changed'
    );
  end if;

  perform document.id
  from public.financial_documents document
  where document.id = v_proposal.base_source_id
  for update;
  if not found then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'source_snapshot_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
    where id = v_proposal.id;
    return jsonb_build_object(
      'proposalId', v_proposal.id, 'runId', v_run.id,
      'status', 'stale', 'reason', 'source_snapshot_changed'
    );
  end if;

  if jsonb_typeof(v_proposal.items) <> 'array'
    or jsonb_array_length(v_proposal.items) <> 1
    or exists (
      select 1
      from jsonb_array_elements(v_proposal.items) item(value)
      where item.value->>'sourceType' is distinct from v_destination_source_type
        or coalesce(item.value->>'sourceId', '') !~
          '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
        or coalesce(item.value->>'sourceDate', '') !~ '^\d{4}-\d{2}-\d{2}$'
        or coalesce(item.value->>'amount', '') !~ '^-?[0-9]+(\.[0-9]+)?$'
    ) then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'source_snapshot_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
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
  if v_locked_destination_count <> 1 then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'source_snapshot_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
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
    0,
    (v_rule_snapshot->>'maxDifferenceDays')::integer,
    v_proposal.base_source_id
  ) candidates;
  if not found
    or v_base.base_source_date is distinct from v_proposal.base_source_date
    or v_base.base_snapshot is distinct from v_proposal.base_snapshot
    or v_base.candidate_count <> 1
    or jsonb_typeof(v_base.candidates) <> 'array'
    or jsonb_array_length(v_base.candidates) <> 1 then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'source_snapshot_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
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
    jsonb_build_object(v_destination_source_type, '+'),
    0,
    1
  );
  if v_combination_count <> 1 then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'combination_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
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
    jsonb_build_object(v_destination_source_type, '+'),
    0,
    1
  );
  select coalesce(jsonb_agg(item.value->'evidence' order by item.ordinality), '[]'::jsonb)
  into v_current_evidence
  from jsonb_array_elements(v_combination.items) with ordinality item(value, ordinality);

  if v_combination.signature is distinct from v_proposal.signature
    or v_combination.items is distinct from v_proposal.items
    or v_current_evidence is distinct from v_proposal.evidence
    or v_combination.calculated_difference is distinct from 0::numeric
    or v_proposal.calculated_difference is distinct from 0::numeric then
    update public.financial_reconciliation_automatic_proposals
    set status = 'stale', reason = 'proposal_evidence_changed',
        reconciliation_id = null, completed_at = null,
        error = '', error_detail = '', updated_at = now()
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

    perform public.financial_reconciliation_action(
      'add_item', p_actor, v_reconciliation_id,
      v_destination_source_type,
      (v_proposal.items#>>'{0,sourceId}')::uuid,
      null
    );

    select count(*) into v_actual_item_count
    from public.financial_reconciliation_items item
    where item.reconciliation_id = v_reconciliation_id;
    if v_actual_item_count <> 2
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
              locked_item.source_type = v_destination_source_type
              and (
                locked_item.source_id <> (v_proposal.items#>>'{0,sourceId}')::uuid
                or locked_item.amount_snapshot is distinct from
                  (v_proposal.items#>>'{0,amount}')::numeric
              )
            )
            or locked_item.source_type not in (
              v_proposal.base_source_type, v_destination_source_type
            )
          )
      ) then
      raise exception 'Automatic reconciliation lifecycle snapshots changed after revalidation.';
    end if;

    v_expected_matching_source_rule := jsonb_build_object(
      'sourceType', v_destination_source_type,
      'operator', '+'
    );
    select matching_rule.value, reconciliation.difference_amount
    into v_actual_matching_source_rule, v_actual_difference
    from public.financial_reconciliations reconciliation
    join lateral jsonb_array_elements(reconciliation.matching_source_rules) matching_rule(value)
      on matching_rule.value->>'sourceType' = v_destination_source_type
    where reconciliation.id = v_reconciliation_id;
    if not found
      or v_actual_matching_source_rule is distinct from v_expected_matching_source_rule
      or v_actual_difference is distinct from 0::numeric then
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

    perform public.financial_reconciliation_action(
      'complete', p_actor, v_reconciliation_id, null, null, null
    );

    insert into public.financial_reconciliation_audit (
      reconciliation_id, action, actor, comment, difference_amount, metadata
    ) values (
      v_reconciliation_id,
      'automatic_complete',
      p_actor,
      null,
      0,
      jsonb_build_object(
        'ruleSnapshot', jsonb_build_object(
          'ruleKey', v_proposal.rule_key,
          'ruleVersion', v_proposal.rule_version,
          'definition', v_rule_snapshot->'definition'
        ),
        'configSnapshot', jsonb_build_object(
          'differenceAllowed', 0,
          'maxDifferenceDays', (v_rule_snapshot->>'maxDifferenceDays')::integer,
          'priority', (v_rule_snapshot->>'priority')::integer
        ),
        'operatorSnapshot', jsonb_build_object(v_destination_source_type, '+'),
        'baseSnapshot', v_proposal.base_snapshot,
        'destinationSnapshots', v_proposal.items,
        'identityEvidence', v_proposal.evidence,
        'proposalSignature', v_proposal.signature,
        'trigger', v_run.trigger,
        'runId', v_run.id,
        'proposalId', v_proposal.id,
        'tolerance', 0,
        'calculatedDifference', 0
      )
    );

    if not exists (
      select 1
      from public.financial_reconciliations reconciliation
      where reconciliation.id = v_reconciliation_id
        and reconciliation.status = 'complete'
        and reconciliation.completion_type = 'normal'
        and reconciliation.difference_amount = 0
        and reconciliation.origin = 'automatic'
        and reconciliation.automatic_trigger = v_run.trigger
        and reconciliation.automatic_rule_key = v_proposal.rule_key
        and reconciliation.automatic_rule_version = v_proposal.rule_version
        and reconciliation.automatic_run_id = v_run.id
        and reconciliation.automatic_proposal_id = v_proposal.id
        and reconciliation.matching_source_rules = jsonb_build_array(
          v_expected_matching_source_rule
        )
    ) or (select count(*)
          from public.financial_reconciliation_items item
          where item.reconciliation_id = v_reconciliation_id) <> 2
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
              locked_item.source_type = v_destination_source_type
              and (
                locked_item.source_id <> (v_proposal.items#>>'{0,sourceId}')::uuid
                or locked_item.amount_snapshot is distinct from
                  (v_proposal.items#>>'{0,amount}')::numeric
              )
            )
            or locked_item.source_type not in (
              v_proposal.base_source_type, v_destination_source_type
            )
          )
      ) then
      raise exception 'Automatic reconciliation lifecycle snapshots changed after revalidation.';
    end if;

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
revoke all on function public.financial_reconciliation_automatic_lock_destination_items(text,jsonb)
  from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_execute_identity_proposal(uuid,text)
  from public, anon, authenticated, service_role;
revoke all on function public.execute_financial_reconciliation_automatic_proposal(uuid,text)
  from public, anon, authenticated, service_role;

grant execute on function public.get_financial_reconciliation_automation_settings()
  to service_role;
grant execute on function public.replace_financial_reconciliation_automation_settings(jsonb,jsonb,text)
  to service_role;
grant execute on function public.financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  to service_role;
grant execute on function public.financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  to service_role;
grant execute on function public.execute_financial_reconciliation_automatic_proposal(uuid,text)
  to service_role;

notify pgrst, 'reload schema';
