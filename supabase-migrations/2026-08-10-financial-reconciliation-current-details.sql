create or replace function public.get_financial_reconciliation_workspace(
  p_reconciliation_id uuid, p_source_type text, p_matching_source_types text[], p_filters jsonb default '{}'::jsonb,
  p_page integer default 1, p_page_size integer default 50
) returns jsonb language plpgsql security definer set search_path = public as $$
declare v_rec public.financial_reconciliations%rowtype; v_offset int; v_candidates jsonb; v_count int; v_started_count int; v_complete_count int;
begin
  if p_source_type not in ('financial_documents','import_fdm_accounts','import_cgd_cartao_credito','import_cgd_extrato_ordem') then raise exception 'Source type is invalid.'; end if;
  if p_page < 1 or p_page_size not between 1 and 100 then raise exception 'Page and page size are invalid.'; end if;
  if p_reconciliation_id is not null then
    select * into v_rec from financial_reconciliations where id=p_reconciliation_id and deleted_at is null;
    if not found then raise exception 'Reconciliation not found.'; end if;
    if p_source_type <> v_rec.base_source_type and not (v_rec.matching_source_types ? p_source_type) then raise exception 'Source type is not allowed for this reconciliation.'; end if;
    if p_matching_source_types is null or to_jsonb(p_matching_source_types) <> v_rec.matching_source_types then raise exception 'Matching source types do not match this reconciliation.'; end if;
  else
    -- Before a group exists, p_source_type is the requested base type.  Validate
    -- the complete mode here as strictly as the start action does.
    if p_matching_source_types is null or cardinality(p_matching_source_types)=0 or cardinality(p_matching_source_types)>3
      or exists(select 1 from unnest(p_matching_source_types) x where x not in ('financial_documents','import_fdm_accounts','import_cgd_cartao_credito','import_cgd_extrato_ordem') or x=p_source_type)
      or cardinality(p_matching_source_types) <> cardinality(array(select distinct x from unnest(p_matching_source_types) x)) then raise exception 'Matching source types are invalid.'; end if;
    if p_source_type <> 'financial_documents' and cardinality(p_matching_source_types) <> 1 then raise exception 'Non-document bases require exactly one matching source type.'; end if;
    if (p_source_type='financial_documents' and exists(select 1 from unnest(p_matching_source_types) x where x not in ('import_fdm_accounts','import_cgd_cartao_credito','import_cgd_extrato_ordem')))
      or (p_source_type='import_fdm_accounts' and p_matching_source_types <> array['import_cgd_extrato_ordem'])
      or (p_source_type='import_cgd_cartao_credito' and p_matching_source_types[1] not in ('financial_documents','import_cgd_extrato_ordem'))
      or (p_source_type='import_cgd_extrato_ordem' and p_matching_source_types[1] not in ('financial_documents','import_fdm_accounts','import_cgd_cartao_credito')) then raise exception 'Matching source type is not allowed for selected base.'; end if;
  end if;
  v_offset := (p_page - 1) * p_page_size;
  with candidate_rows as (
    select s.* from (
      select d.id, d.amount, d.document_date as source_date, d.description, d.supplier_name as supplier, d.cc as account, d.category, d.payment, d.fat as document_fat, d.supplier_nif, ''::text as reservation_id from financial_documents d where p_source_type='financial_documents' and d.document_date >= date '2026-01-01' and d.fat='S'
      union all select f.id,f.amount,f.event_date,f.description,f.guest,f.account,f.category,f.invoice,'','',f.reservation_id from import_fdm_accounts f where p_source_type='import_fdm_accounts' and f.event_date >= date '2026-01-01' and (coalesce(f.invoice_flag,false) or f.category='Compras') and (coalesce(v_rec.base_source_type,'') <> 'financial_documents' or f.category='Compras')
      union all select c.id,c.valor,c.data,c.descricao,'','','','','','','' from import_cgd_cartao_credito c where p_source_type='import_cgd_cartao_credito' and c.data >= date '2026-01-01'
      union all select b.id,b.montante,b.data,b.descritivo,'','','','','','','' from import_cgd_extrato_ordem b where p_source_type='import_cgd_extrato_ordem' and b.data >= date '2026-01-01'
    ) s where not exists (select 1 from financial_reconciliation_items i join financial_reconciliations r on r.id=i.reconciliation_id where i.source_type=p_source_type and i.source_id=s.id and r.deleted_at is null)
      and (nullif(p_filters->>'dateFrom','') is null or s.source_date >= (p_filters->>'dateFrom')::date)
      and (nullif(p_filters->>'dateTo','') is null or s.source_date <= (p_filters->>'dateTo')::date)
      and (nullif(p_filters->>'amountMin','') is null or s.amount >= (p_filters->>'amountMin')::numeric)
      and (nullif(p_filters->>'amountMax','') is null or s.amount <= (p_filters->>'amountMax')::numeric)
      and (nullif(p_filters->>'description','') is null or s.description ilike '%' || p_filters->>'description' || '%')
      and (nullif(p_filters->>'supplier','') is null or s.supplier ilike '%' || p_filters->>'supplier' || '%')
      and (nullif(p_filters->>'payment','') is null or s.payment = p_filters->>'payment')
      and (nullif(p_filters->>'account','') is null or s.account = p_filters->>'account')
      and (nullif(p_filters->>'category','') is null or s.category = p_filters->>'category')
  ) select count(*) into v_count from candidate_rows;
  select count(*) into v_started_count from financial_reconciliation_items i join financial_reconciliations r on r.id=i.reconciliation_id where i.source_type=p_source_type and r.deleted_at is null and r.status='started';
  select count(*) into v_complete_count from financial_reconciliation_items i join financial_reconciliations r on r.id=i.reconciliation_id where i.source_type=p_source_type and r.deleted_at is null and r.status='complete';
  with candidate_rows as (
    select s.* from (
      select d.id, d.amount, d.document_date as source_date, d.description, d.supplier_name as supplier, d.cc as account, d.category, d.payment, d.fat as document_fat, d.supplier_nif, ''::text as reservation_id from financial_documents d where p_source_type='financial_documents' and d.document_date >= date '2026-01-01' and d.fat='S'
      union all select f.id,f.amount,f.event_date,f.description,f.guest,f.account,f.category,f.invoice,'','',f.reservation_id from import_fdm_accounts f where p_source_type='import_fdm_accounts' and f.event_date >= date '2026-01-01' and (coalesce(f.invoice_flag,false) or f.category='Compras') and (coalesce(v_rec.base_source_type,'') <> 'financial_documents' or f.category='Compras')
      union all select c.id,c.valor,c.data,c.descricao,'','','','','','','' from import_cgd_cartao_credito c where p_source_type='import_cgd_cartao_credito' and c.data >= date '2026-01-01'
      union all select b.id,b.montante,b.data,b.descritivo,'','','','','','','' from import_cgd_extrato_ordem b where p_source_type='import_cgd_extrato_ordem' and b.data >= date '2026-01-01'
    ) s where not exists (select 1 from financial_reconciliation_items i join financial_reconciliations r on r.id=i.reconciliation_id where i.source_type=p_source_type and i.source_id=s.id and r.deleted_at is null)
      and (nullif(p_filters->>'dateFrom','') is null or s.source_date >= (p_filters->>'dateFrom')::date)
      and (nullif(p_filters->>'dateTo','') is null or s.source_date <= (p_filters->>'dateTo')::date)
      and (nullif(p_filters->>'amountMin','') is null or s.amount >= (p_filters->>'amountMin')::numeric)
      and (nullif(p_filters->>'amountMax','') is null or s.amount <= (p_filters->>'amountMax')::numeric)
      and (nullif(p_filters->>'description','') is null or s.description ilike '%' || p_filters->>'description' || '%')
      and (nullif(p_filters->>'supplier','') is null or s.supplier ilike '%' || p_filters->>'supplier' || '%')
      and (nullif(p_filters->>'payment','') is null or s.payment = p_filters->>'payment')
      and (nullif(p_filters->>'account','') is null or s.account = p_filters->>'account')
      and (nullif(p_filters->>'category','') is null or s.category = p_filters->>'category')
  ) select coalesce(jsonb_agg(to_jsonb(x) order by x.source_date desc), '[]'::jsonb) into v_candidates from (select * from candidate_rows order by source_date desc offset v_offset limit p_page_size) x;
  return jsonb_build_object('candidates',v_candidates,'totalCount',v_count,'counts',jsonb_build_object('notStarted',v_count,'started',v_started_count,'complete',v_complete_count),'page',p_page,'pageSize',p_page_size,'sourceConfig',jsonb_build_object('sourceType',p_source_type,'dateColumn','source_date','amountColumn','amount','columns',case p_source_type when 'financial_documents' then jsonb_build_array('document_fat','supplier_nif','supplier','payment') when 'import_fdm_accounts' then jsonb_build_array('reservation_id','account','category') else jsonb_build_array('description') end,'filterFields',case p_source_type when 'financial_documents' then jsonb_build_array('dateFrom','dateTo','amountMin','amountMax','description','supplier','payment','account','category') when 'import_fdm_accounts' then jsonb_build_array('dateFrom','dateTo','amountMin','amountMax','description','account','category') when 'import_cgd_cartao_credito' then jsonb_build_array('dateFrom','dateTo','amountMin','amountMax','description') else jsonb_build_array('dateFrom','dateTo','amountMin','amountMax','description') end),'reconciliation',case when p_reconciliation_id is null then null else to_jsonb(v_rec) end,'items',coalesce((select jsonb_agg(to_jsonb(i) || jsonb_build_object('source_date',s.source_date,'description',s.description,'supplier',nullif(s.supplier,'')) order by i.created_at) from financial_reconciliation_items i left join lateral financial_reconciliation_source(i.source_type,i.source_id) s on true where i.reconciliation_id=p_reconciliation_id),'[]'::jsonb),'audit',coalesce((select jsonb_agg(to_jsonb(a) order by a.created_at,a.id) from financial_reconciliation_audit a where a.reconciliation_id=p_reconciliation_id),'[]'::jsonb),'history',coalesce((select jsonb_agg(to_jsonb(h) order by h.created_at desc) from (select * from financial_reconciliations where deleted_at is null order by created_at desc limit 100) h),'[]'::jsonb));
end $$;
