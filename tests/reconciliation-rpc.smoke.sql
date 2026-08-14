-- Run against a disposable Supabase project after applying the migration.
-- The migration is reapplied inside this transaction to exercise its historical
-- snapshot backfill against the explicit legacy fixture below.
begin;
insert into financial_reconciliations (
  id, status, base_source_type, matching_source_types, matching_source_rules, created_by
) values (
  gen_random_uuid(), 'started', 'financial_documents', '["import_cgd_extrato_ordem"]'::jsonb, '[]'::jsonb,
  'smoke-legacy-backfill-' || txid_current()
);
\ir ../supabase-migrations/2026-08-11-financial-reconciliation-source-rules.sql
\ir ../supabase-migrations/2026-08-12-financial-reconciliation-oldest-first-candidates.sql
\ir ../supabase-migrations/2026-08-12-financial-reconciliation-oldest-first-candidates.sql
\ir ../supabase-migrations/2026-08-12-financial-reconciliation-history-source-summary.sql
\ir ../supabase-migrations/2026-08-12-financial-reconciliation-history-source-summary.sql
\ir ../supabase-migrations/2026-08-14-financial-reconciliation-automation-schema.sql
\ir ../supabase-migrations/2026-08-14-financial-reconciliation-automation-analysis.sql
\ir ../supabase-migrations/2026-08-14-financial-reconciliation-automation-execution.sql
do $$ declare doc_id uuid:=gen_random_uuid(); bank_id uuid:=gen_random_uuid(); fdm_bank_id uuid:=gen_random_uuid(); card_id uuid:=gen_random_uuid(); fdm_id uuid:=gen_random_uuid(); old_doc_id uuid := '00000000-0000-0000-0000-000000000101'; same_date_low_id uuid := '00000000-0000-0000-0000-000000000102'; same_date_high_id uuid := '00000000-0000-0000-0000-000000000103'; new_doc_id uuid := '00000000-0000-0000-0000-000000000104'; candidate_ids uuid[]; r jsonb; rid uuid; fdm_rid uuid; v_rules jsonb; v_before jsonb; v_invalid jsonb; v_rejected boolean; history_rid uuid := gen_random_uuid(); history_row jsonb; history_source_ids text[]; history_card_item_id uuid := gen_random_uuid();
begin
  if not coalesce((
    select c.relrowsecurity
    from pg_class c
    join pg_namespace n on n.oid = c.relnamespace
    where n.nspname = 'public' and c.relname = 'financial_reconciliation_source_rules'
  ), false) then
    raise exception 'Financial reconciliation source rules must have RLS enabled';
  end if;
  if exists (
    select 1
    from (values ('anon'), ('authenticated')) roles(role_name)
    cross join (values ('select'), ('insert'), ('update'), ('delete'), ('truncate'), ('references'), ('trigger')) privileges(privilege_name)
    where has_table_privilege(roles.role_name::name, 'public.financial_reconciliation_source_rules'::text, privileges.privilege_name)
  ) then
    raise exception 'Anon and authenticated roles must not have source-rule table privileges';
  end if;
  if has_function_privilege('anon'::name, 'public.replace_financial_reconciliation_source_rules(jsonb)'::regprocedure, 'execute')
     or has_function_privilege('authenticated'::name, 'public.replace_financial_reconciliation_source_rules(jsonb)'::regprocedure, 'execute')
     or not has_function_privilege('service_role'::name, 'public.replace_financial_reconciliation_source_rules(jsonb)'::regprocedure, 'execute') then
    raise exception 'Source-rule replacement RPC permissions are incorrect';
  end if;
  if not exists (
    select 1
    from financial_reconciliations historical
    where historical.created_by = 'smoke-legacy-backfill-' || txid_current()
      and historical.matching_source_rules @> '[{"sourceType":"import_cgd_extrato_ordem","operator":"+"}]'::jsonb
  ) then
    raise exception 'Legacy blank snapshot fixture was not backfilled with its legacy plus operator';
  end if;
  if exists (
    select 1
    from financial_reconciliations historical
    where jsonb_array_length(historical.matching_source_types) > 0
      and historical.matching_source_rules = '[]'::jsonb
  ) then
    raise exception 'Historical reconciliation with matching sources was not backfilled with a rule snapshot';
  end if;
  if exists (
    select 1
    from financial_reconciliations historical
    where historical.base_source_type = 'financial_documents'
      and historical.matching_source_types @> '["import_cgd_extrato_ordem"]'::jsonb
      and not (historical.matching_source_rules @> '[{"sourceType":"import_cgd_extrato_ordem","operator":"+"}]'::jsonb)
  ) then
    raise exception 'Historical Financial Documents bank snapshots must preserve the legacy plus operator';
  end if;
  if not exists (select 1 from financial_reconciliation_source_rules where base_source_type='financial_documents' and matching_source_type='import_cgd_extrato_ordem' and operator='+')
     or not exists (select 1 from financial_reconciliation_source_rules where base_source_type='import_cgd_extrato_ordem' and matching_source_type='financial_documents' and operator='+') then
    raise exception 'Independent reverse-direction rules were not seeded';
  end if;
  insert into financial_reconciliations(
    id,status,base_source_type,matching_source_types,matching_source_rules,created_by
  ) values (
    history_rid,'started','financial_documents',
    '["import_cgd_extrato_ordem","import_cgd_cartao_credito"]'::jsonb,
    '[{"sourceType":"import_cgd_extrato_ordem","operator":"+"},{"sourceType":"import_cgd_cartao_credito","operator":"-"}]'::jsonb,
    'smoke-history-summary'
  );

  insert into financial_reconciliation_items(
    reconciliation_id,source_type,source_id,amount_snapshot,created_by
  ) values
    (history_rid,'financial_documents',gen_random_uuid(),200,'smoke-history-summary'),
    (history_rid,'financial_documents',gen_random_uuid(),250,'smoke-history-summary'),
    (history_rid,'import_cgd_extrato_ordem',gen_random_uuid(),-200,'smoke-history-summary'),
    (history_rid,'import_cgd_extrato_ordem',gen_random_uuid(),-250,'smoke-history-summary'),
    (history_rid,'import_cgd_cartao_credito',history_card_item_id,25,'smoke-history-summary');

  r := get_financial_reconciliation_workspace(
    history_rid,'financial_documents','{}'::jsonb,1,50
  );
  select history_record
  into history_row
  from jsonb_array_elements(r->'history') history_records(history_record)
  where history_record->>'id'=history_rid::text;

  select array_agg(source_entry->>'sourceType' order by position)
  into history_source_ids
  from jsonb_array_elements(history_row->'sourceSummary')
    with ordinality source_entries(source_entry,position);

  if history_source_ids is distinct from array[
    'financial_documents','import_cgd_extrato_ordem','import_cgd_cartao_credito'
  ] then
    raise exception 'History sources were not ordered base-first then by saved matching source order: %', history_source_ids;
  end if;

  if history_row->'sourceSummary'->0 is distinct from
     '{"sourceType":"financial_documents","recordCount":2,"amountTotal":450}'::jsonb
     or history_row->'sourceSummary'->1 is distinct from
     '{"sourceType":"import_cgd_extrato_ordem","recordCount":2,"amountTotal":-450}'::jsonb
     or history_row->'sourceSummary'->2 is distinct from
     '{"sourceType":"import_cgd_cartao_credito","recordCount":1,"amountTotal":25}'::jsonb then
    raise exception 'History source summary did not preserve raw counts and amount snapshots: %', history_row->'sourceSummary';
  end if;

  delete from financial_reconciliation_items
  where reconciliation_id=history_rid
    and source_type='import_cgd_cartao_credito'
    and source_id=history_card_item_id;

  r := get_financial_reconciliation_workspace(
    history_rid,'financial_documents','{}'::jsonb,1,50
  );
  select history_record
  into history_row
  from jsonb_array_elements(r->'history') history_records(history_record)
  where history_record->>'id'=history_rid::text;

  if jsonb_array_length(history_row->'sourceSummary') <> 2
     or exists (
       select 1
       from jsonb_array_elements(history_row->'sourceSummary') entry
       where entry->>'sourceType'='import_cgd_cartao_credito'
     ) then
    raise exception 'Removed or unused history source was not omitted: %', history_row->'sourceSummary';
  end if;
  select coalesce(jsonb_agg(jsonb_build_object('base_source_type', base_source_type, 'matching_source_type', matching_source_type, 'operator', operator) order by base_source_type, matching_source_type), '[]'::jsonb)
    into v_before
    from financial_reconciliation_source_rules;
  for v_invalid in
    select rules
    from (values
      ('[{"base_source_type":"unknown","matching_source_type":"financial_documents","operator":"+"}]'::jsonb),
      ('[{"base_source_type":"financial_documents","matching_source_type":"financial_documents","operator":"+"}]'::jsonb),
      ('[{"base_source_type":"financial_documents","matching_source_type":"import_cgd_extrato_ordem","operator":"*"}]'::jsonb),
      ('[{"base_source_type":"financial_documents","matching_source_type":"import_cgd_extrato_ordem","operator":"+"},{"base_source_type":"financial_documents","matching_source_type":"import_cgd_extrato_ordem","operator":"-"}]'::jsonb)
    ) invalid_rules(rules)
  loop
    v_rejected := false;
    begin
      perform public.replace_financial_reconciliation_source_rules(v_invalid);
    exception when others then
      v_rejected := true;
    end;
    if not v_rejected then
      raise exception 'Invalid source-rule replacement payload was accepted';
    end if;
  end loop;
  if (select coalesce(jsonb_agg(jsonb_build_object('base_source_type', base_source_type, 'matching_source_type', matching_source_type, 'operator', operator) order by base_source_type, matching_source_type), '[]'::jsonb) from financial_reconciliation_source_rules) is distinct from v_before then
    raise exception 'Invalid source-rule replacement mutated persisted rules';
  end if;
  insert into financial_documents(id,document_date,amount,fat,created_by,description,supplier_name) values(doc_id,'2026-01-02',100,'S','smoke','smoke document','Smoke Supplier');
  insert into import_cgd_extrato_ordem(id,import_batch,row_key,data,montante) values(bank_id,'smoke','recon-smoke-bank-'||bank_id,'2026-01-02',-100);
  insert into import_cgd_extrato_ordem(id,import_batch,row_key,data,montante) values(fdm_bank_id,'smoke','recon-smoke-fdm-bank-'||fdm_bank_id,'2026-01-02',-30);
  insert into import_cgd_cartao_credito(id,import_batch,row_key,data,debito) values(card_id,'smoke','recon-smoke-card-'||card_id,'2026-01-02',30);
  insert into import_fdm_accounts(id,import_batch,account,date_time_raw,event_date,category,amount,invoice_flag) values(fdm_id,'smoke','smoke','2026-01-02','2026-01-02','Compras',-30,true);
  insert into financial_documents(id,document_date,amount,fat,created_by,description) values
    (old_doc_id,'2026-02-01',1,'S','smoke','oldest ordering fixture'),
    (same_date_high_id,'2026-02-02',1,'S','smoke','same date high id ordering fixture'),
    (same_date_low_id,'2026-02-02',1,'S','smoke','same date low id ordering fixture'),
    (new_doc_id,'2026-02-03',1,'S','smoke','newest ordering fixture');
  r := get_financial_reconciliation_workspace(
    null,
    'financial_documents',
    '{"dateFrom":"2026-02-01","dateTo":"2026-02-03","amountMin":"1","amountMax":"1","description":"ordering fixture"}'::jsonb,
    1,
    3
  );
  select array_agg((candidate->>'id')::uuid order by ordinal)
    into candidate_ids
    from jsonb_array_elements(r->'candidates') with ordinality candidates(candidate, ordinal);
  if candidate_ids is distinct from array[old_doc_id, same_date_low_id, same_date_high_id] then
    raise exception 'Candidates were not paginated after oldest-first deterministic ordering: %', candidate_ids;
  end if;
  r:=financial_reconciliation_action('start','smoke',null,'financial_documents',doc_id,null); rid:=(r->'reconciliation'->>'id')::uuid;
  if r->'reconciliation'->'matching_source_rules' @> '[{"sourceType":"import_cgd_extrato_ordem","operator":"+"}]'::jsonb is not true then raise exception 'Start did not snapshot directional rules'; end if;
  select coalesce(jsonb_agg(jsonb_build_object(
    'base_source_type', base_source_type,
    'matching_source_type', matching_source_type,
    'operator', case when base_source_type='financial_documents' and matching_source_type='import_cgd_extrato_ordem' then '-' else operator end
  ) order by base_source_type, matching_source_type), '[]'::jsonb)
    into v_rules
    from financial_reconciliation_source_rules;
  perform public.replace_financial_reconciliation_source_rules(v_rules);
  if not exists (select 1 from financial_reconciliation_source_rules where base_source_type='import_cgd_extrato_ordem' and matching_source_type='financial_documents' and operator='+') then
    raise exception 'Changing one direction changed the independently configured reverse rule';
  end if;
  perform financial_reconciliation_action('add_item','smoke',rid,'import_cgd_extrato_ordem',bank_id,null);
  if (select difference_amount from financial_reconciliations where id=rid) <> 0 then raise exception 'Expected zero difference'; end if;
  r := financial_reconciliation_action('complete','smoke',rid,null,null,null);
  if not exists (
      select 1 from public.financial_reconciliations manual_reconciliation
      where manual_reconciliation.id = rid
        and manual_reconciliation.origin = 'user'
        and manual_reconciliation.automatic_trigger is null
        and manual_reconciliation.automatic_rule_key is null
        and manual_reconciliation.automatic_rule_version is null
        and manual_reconciliation.automatic_run_id is null
        and manual_reconciliation.automatic_proposal_id is null
    )
    or not (r->'reconciliation' ?& array[
      'origin','automaticTrigger','automaticRuleKey','automaticRuleVersion','automaticRunId'
    ])
    or r#>>'{reconciliation,origin}' <> 'user'
    or r#>'{reconciliation,automaticTrigger}' <> 'null'::jsonb
    or r#>'{reconciliation,automaticRuleKey}' <> 'null'::jsonb
    or r#>'{reconciliation,automaticRuleVersion}' <> 'null'::jsonb
    or r#>'{reconciliation,automaticRunId}' <> 'null'::jsonb then
    raise exception 'Manual lifecycle did not preserve user origin and null automation provenance';
  end if;
  perform financial_reconciliation_action('reopen','smoke',rid,null,null,null);
  perform financial_reconciliation_action('remove_item','smoke',rid,'import_cgd_extrato_ordem',bank_id,null);
  if exists(select 1 from financial_reconciliation_items where source_type='import_cgd_extrato_ordem' and source_id=bank_id) then raise exception 'Removed record is still locked'; end if;
  r:=get_financial_reconciliation_workspace(rid,'import_cgd_extrato_ordem','{}'::jsonb,1,50);
  if not exists (
    select 1
    from jsonb_array_elements(r->'items') item
    where (item->>'source_id')::uuid = doc_id
      and item->>'source_date' = '2026-01-02'
      and item->>'description' = 'smoke document'
      and item->>'supplier' = 'Smoke Supplier'
  ) then
    raise exception 'Workspace item details were not returned';
  end if;
  if not exists (select 1 from jsonb_array_elements(r->'candidates') candidate where (candidate->>'id')::uuid=bank_id) then raise exception 'Removed record did not reappear as a candidate'; end if;
  begin
    perform financial_reconciliation_action('add_item','smoke',rid,'financial_documents',doc_id,null);
    raise exception 'Expected duplicate membership failure';
  exception when others then if sqlerrm <> 'This record is already reconciled.' then raise; end if; end;
  r:=financial_reconciliation_action('start','smoke',null,'import_fdm_accounts',fdm_id,null); fdm_rid:=(r->'reconciliation'->>'id')::uuid;
  if r->'reconciliation'->'matching_source_rules' @> '[{"sourceType":"import_cgd_extrato_ordem","operator":"-"}]'::jsonb is not true then raise exception 'FDM start did not snapshot its minus operator'; end if;
  update financial_reconciliation_source_rules set operator='+' where base_source_type='import_fdm_accounts' and matching_source_type='import_cgd_extrato_ordem';
  perform financial_reconciliation_action('add_item','smoke',fdm_rid,'import_cgd_extrato_ordem',fdm_bank_id,null);
  if (select difference_amount from financial_reconciliations where id=fdm_rid) <> 0 then raise exception 'Captured minus snapshot did not calculate through SQL'; end if;
  begin
    perform financial_reconciliation_action('add_item','smoke',fdm_rid,'import_cgd_cartao_credito',card_id,null);
    raise exception 'Expected source absent from snapshot failure';
  exception when others then if sqlerrm <> 'Item source type is not allowed for this reconciliation.' then raise; end if; end;
end $$;
rollback;
