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
do $$ declare doc_id uuid:=gen_random_uuid(); bank_id uuid:=gen_random_uuid(); card_id uuid:=gen_random_uuid(); fdm_id uuid:=gen_random_uuid(); r jsonb; rid uuid; fdm_rid uuid;
begin
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
  insert into financial_documents(id,document_date,amount,fat,created_by,description,supplier_name) values(doc_id,'2026-01-02',100,'S','smoke','smoke document','Smoke Supplier');
  insert into import_cgd_extrato_ordem(id,import_batch,row_key,data,montante) values(bank_id,'smoke','recon-smoke-bank-'||bank_id,'2026-01-02',-100);
  insert into import_cgd_cartao_credito(id,import_batch,row_key,data,debito) values(card_id,'smoke','recon-smoke-card-'||card_id,'2026-01-02',30);
  insert into import_fdm_accounts(id,import_batch,account,date_time_raw,event_date,category,amount,invoice_flag) values(fdm_id,'smoke','smoke','2026-01-02','2026-01-02','Compras',-30,true);
  r:=financial_reconciliation_action('start','smoke',null,'financial_documents',doc_id,null); rid:=(r->'reconciliation'->>'id')::uuid;
  if r->'reconciliation'->'matching_source_rules' @> '[{"sourceType":"import_cgd_extrato_ordem","operator":"+"}]'::jsonb is not true then raise exception 'Start did not snapshot directional rules'; end if;
  update financial_reconciliation_source_rules set operator='-' where base_source_type='financial_documents' and matching_source_type='import_cgd_extrato_ordem';
  perform financial_reconciliation_action('add_item','smoke',rid,'import_cgd_extrato_ordem',bank_id,null);
  if (select difference_amount from financial_reconciliations where id=rid) <> 0 then raise exception 'Expected zero difference'; end if;
  perform financial_reconciliation_action('complete','smoke',rid,null,null,null);
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
  begin
    perform financial_reconciliation_action('add_item','smoke',fdm_rid,'import_cgd_cartao_credito',card_id,null);
    raise exception 'Expected source absent from snapshot failure';
  exception when others then if sqlerrm <> 'Item source type is not allowed for this reconciliation.' then raise; end if; end;
end $$;
rollback;
