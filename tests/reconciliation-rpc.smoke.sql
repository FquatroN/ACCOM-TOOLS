-- Run against a disposable Supabase project after applying the migration.
begin;
do $$ declare doc_id uuid:=gen_random_uuid(); bank_id uuid:=gen_random_uuid(); card_id uuid:=gen_random_uuid(); fdm_id uuid:=gen_random_uuid(); r jsonb; rid uuid;
begin
  insert into financial_documents(id,document_date,amount,fat,created_by,description,supplier_name) values(doc_id,'2026-01-02',100,'S','smoke','smoke document','Smoke Supplier');
  insert into import_cgd_extrato_ordem(id,import_batch,row_key,data,montante) values(bank_id,'smoke','recon-smoke-bank-'||bank_id,'2026-01-02',-40);
  insert into import_cgd_cartao_credito(id,import_batch,row_key,data,debito) values(card_id,'smoke','recon-smoke-card-'||card_id,'2026-01-02',30);
  insert into import_fdm_accounts(id,import_batch,account,date_time_raw,event_date,category,amount,invoice_flag) values(fdm_id,'smoke','smoke','2026-01-02','2026-01-02','Compras',-30,true);
  r:=financial_reconciliation_action('start','smoke',null,'financial_documents',array['import_cgd_extrato_ordem','import_cgd_cartao_credito','import_fdm_accounts'],'financial_documents',doc_id,null); rid:=(r->'reconciliation'->>'id')::uuid;
  perform financial_reconciliation_action('add_item','smoke',rid,null,null,'import_cgd_extrato_ordem',bank_id,null);
  perform financial_reconciliation_action('add_item','smoke',rid,null,null,'import_cgd_cartao_credito',card_id,null);
  perform financial_reconciliation_action('add_item','smoke',rid,null,null,'import_fdm_accounts',fdm_id,null);
  if (select difference_amount from financial_reconciliations where id=rid) <> 0 then raise exception 'Expected zero difference'; end if;
  perform financial_reconciliation_action('complete','smoke',rid,null,null,null,null,null);
  perform financial_reconciliation_action('reopen','smoke',rid,null,null,null,null,null);
  perform financial_reconciliation_action('remove_item','smoke',rid,null,null,'import_cgd_cartao_credito',card_id,null);
  if exists(select 1 from financial_reconciliation_items where source_type='import_cgd_cartao_credito' and source_id=card_id) then raise exception 'Removed record is still locked'; end if;
  r:=get_financial_reconciliation_workspace(rid,'import_cgd_cartao_credito',array['import_cgd_extrato_ordem','import_cgd_cartao_credito','import_fdm_accounts']);
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
  if not exists (select 1 from jsonb_array_elements(r->'candidates') candidate where (candidate->>'id')::uuid=card_id) then raise exception 'Removed record did not reappear as a candidate'; end if;
  begin
    perform financial_reconciliation_action('add_item','smoke',rid,null,null,'financial_documents',doc_id,null);
    raise exception 'Expected duplicate membership failure';
  exception when others then if sqlerrm <> 'This record is already reconciled.' then raise; end if; end;
end $$;
rollback;
