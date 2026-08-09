-- Financial reconciliation workflow.  Writes are intentionally routed through
-- financial_reconciliation_action so row locks and audit entries are atomic.
create extension if not exists pgcrypto;

create table if not exists public.financial_reconciliations (
  id uuid primary key default gen_random_uuid(),
  status text not null check (status in ('started', 'complete')),
  base_source_type text not null check (base_source_type in ('financial_documents', 'import_fdm_accounts', 'import_cgd_cartao_credito', 'import_cgd_extrato_ordem')),
  matching_source_types jsonb not null default '[]'::jsonb,
  completion_type text check (completion_type in ('normal', 'forced')),
  difference_amount numeric(14,2) not null default 0,
  forced_completion_comment text,
  created_by text not null,
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now()),
  completed_by text, completed_at timestamptz,
  deleted_by text, deleted_at timestamptz
);

create table if not exists public.financial_reconciliation_items (
  id uuid primary key default gen_random_uuid(),
  reconciliation_id uuid not null references public.financial_reconciliations(id),
  source_type text not null check (source_type in ('financial_documents', 'import_fdm_accounts', 'import_cgd_cartao_credito', 'import_cgd_extrato_ordem')),
  source_id uuid not null,
  amount_snapshot numeric(14,2) not null,
  created_by text not null,
  created_at timestamptz not null default timezone('utc', now()),
  unique (source_type, source_id)
);

create table if not exists public.financial_reconciliation_audit (
  id uuid primary key default gen_random_uuid(),
  reconciliation_id uuid not null references public.financial_reconciliations(id),
  action text not null check (action in ('start', 'add_item', 'remove_item', 'complete', 'force_complete', 'reopen', 'delete')),
  actor text not null,
  comment text,
  difference_amount numeric(14,2) not null default 0,
  metadata jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default timezone('utc', now())
);

create index if not exists financial_reconciliations_active_created_idx on public.financial_reconciliations (created_at desc) where deleted_at is null;
create index if not exists financial_reconciliation_items_reconciliation_idx on public.financial_reconciliation_items (reconciliation_id, created_at);
create index if not exists financial_reconciliation_items_lock_idx on public.financial_reconciliation_items (source_type, source_id);
create index if not exists financial_reconciliation_audit_chronology_idx on public.financial_reconciliation_audit (reconciliation_id, created_at, id);

alter table public.financial_reconciliations enable row level security;
alter table public.financial_reconciliation_items enable row level security;
alter table public.financial_reconciliation_audit enable row level security;
drop policy if exists "financial reconciliations authenticated select" on public.financial_reconciliations;
create policy "financial reconciliations authenticated select" on public.financial_reconciliations for select to authenticated using (true);
drop policy if exists "financial reconciliation items authenticated select" on public.financial_reconciliation_items;
create policy "financial reconciliation items authenticated select" on public.financial_reconciliation_items for select to authenticated using (true);
drop policy if exists "financial reconciliation audit authenticated select" on public.financial_reconciliation_audit;
create policy "financial reconciliation audit authenticated select" on public.financial_reconciliation_audit for select to authenticated using (true);
grant select on public.financial_reconciliations, public.financial_reconciliation_items, public.financial_reconciliation_audit to authenticated;

create or replace function public.financial_reconciliation_source(p_source_type text, p_source_id uuid)
returns table (source_id uuid, amount numeric, source_date date, description text, supplier text, account text, category text, payment text, eligible boolean)
language plpgsql stable security definer set search_path = public as $$
begin
  if p_source_type = 'financial_documents' then
    return query select d.id, d.amount, d.document_date, d.description, d.supplier_name, d.cc, d.category, d.payment,
      d.document_date >= date '2026-01-01' and d.fat = 'S' from financial_documents d where d.id = p_source_id;
  elsif p_source_type = 'import_fdm_accounts' then
    return query select f.id, f.amount, f.event_date, f.description, f.guest, f.account, f.category, f.invoice,
      f.event_date >= date '2026-01-01' and (coalesce(f.invoice_flag, false) or f.category = 'Compras') from import_fdm_accounts f where f.id = p_source_id;
  elsif p_source_type = 'import_cgd_cartao_credito' then
    return query select c.id, c.valor, c.data, c.descricao, ''::text, ''::text, ''::text, ''::text,
      c.data >= date '2026-01-01' from import_cgd_cartao_credito c where c.id = p_source_id;
  elsif p_source_type = 'import_cgd_extrato_ordem' then
    return query select b.id, b.montante, b.data, b.descritivo, ''::text, ''::text, ''::text, ''::text,
      b.data >= date '2026-01-01' from import_cgd_extrato_ordem b where b.id = p_source_id;
  else raise exception 'Source type is invalid.'; end if;
end $$;

create or replace function public.financial_reconciliation_difference(p_base text, p_matching jsonb, p_reconciliation_id uuid)
returns numeric language plpgsql stable security definer set search_path = public as $$
declare d numeric := 0; f numeric := 0; c numeric := 0; b numeric := 0; m text;
begin
  select coalesce(sum(amount_snapshot) filter (where source_type = 'financial_documents'),0), coalesce(sum(amount_snapshot) filter (where source_type = 'import_fdm_accounts'),0), coalesce(sum(amount_snapshot) filter (where source_type = 'import_cgd_cartao_credito'),0), coalesce(sum(amount_snapshot) filter (where source_type = 'import_cgd_extrato_ordem'),0) into d,f,c,b from financial_reconciliation_items where reconciliation_id = p_reconciliation_id;
  if p_base = 'financial_documents' then return round(d + f + c + b, 2); end if;
  m := p_matching ->> 0;
  if p_base = 'import_fdm_accounts' or (p_base = 'import_cgd_extrato_ordem' and m = 'import_fdm_accounts') then return round(f - b, 2); end if;
  if p_base = 'import_cgd_cartao_credito' then return round(c + case when m = 'financial_documents' then d else b end, 2); end if;
  return round(b + case when m = 'financial_documents' then d else c end, 2);
end $$;

create or replace function public.get_financial_reconciliation_workspace(
  p_reconciliation_id uuid, p_source_type text, p_matching_source_types text[], p_filters jsonb default '{}'::jsonb,
  p_page integer default 1, p_page_size integer default 50
) returns jsonb language plpgsql security definer set search_path = public as $$
declare v_rec public.financial_reconciliations%rowtype; v_offset int; v_candidates jsonb; v_count int;
begin
  if p_source_type not in ('financial_documents','import_fdm_accounts','import_cgd_cartao_credito','import_cgd_extrato_ordem') then raise exception 'Source type is invalid.'; end if;
  if p_page < 1 or p_page_size not between 1 and 100 then raise exception 'Page and page size are invalid.'; end if;
  if p_reconciliation_id is not null then
    select * into v_rec from financial_reconciliations where id=p_reconciliation_id and deleted_at is null;
    if not found then raise exception 'Reconciliation not found.'; end if;
    if p_source_type <> v_rec.base_source_type and not (v_rec.matching_source_types ? p_source_type) then raise exception 'Source type is not allowed for this reconciliation.'; end if;
    if p_matching_source_types is null or to_jsonb(p_matching_source_types) <> v_rec.matching_source_types then raise exception 'Matching source types do not match this reconciliation.'; end if;
  elsif p_matching_source_types is null or cardinality(p_matching_source_types) = 0 then raise exception 'Matching source types are required.'; end if;
  v_offset := (p_page - 1) * p_page_size;
  with candidate_rows as (
    select s.* from (
      select d.id, d.amount, d.document_date as source_date, d.description, d.supplier_name as supplier, d.cc as account, d.category, d.payment, d.fat as document_fat, d.supplier_nif, ''::text as reservation_id from financial_documents d where p_source_type='financial_documents' and d.document_date >= date '2026-01-01' and d.fat='S'
      union all select f.id,f.amount,f.event_date,f.description,f.guest,f.account,f.category,f.invoice,'','',f.reservation_id from import_fdm_accounts f where p_source_type='import_fdm_accounts' and f.event_date >= date '2026-01-01' and (coalesce(f.invoice_flag,false) or f.category='Compras')
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
  with candidate_rows as (
    select s.* from (
      select d.id, d.amount, d.document_date as source_date, d.description, d.supplier_name as supplier, d.cc as account, d.category, d.payment, d.fat as document_fat, d.supplier_nif, ''::text as reservation_id from financial_documents d where p_source_type='financial_documents' and d.document_date >= date '2026-01-01' and d.fat='S'
      union all select f.id,f.amount,f.event_date,f.description,f.guest,f.account,f.category,f.invoice,'','',f.reservation_id from import_fdm_accounts f where p_source_type='import_fdm_accounts' and f.event_date >= date '2026-01-01' and (coalesce(f.invoice_flag,false) or f.category='Compras')
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
  return jsonb_build_object('candidates',v_candidates,'totalCount',v_count,'page',p_page,'pageSize',p_page_size,'sourceConfig',jsonb_build_object('sourceType',p_source_type,'dateColumn',case p_source_type when 'financial_documents' then 'document_date' when 'import_fdm_accounts' then 'event_date' else 'data' end,'amountColumn',case p_source_type when 'financial_documents' then 'amount' when 'import_fdm_accounts' then 'amount' when 'import_cgd_cartao_credito' then 'valor' else 'montante' end,'columns',case p_source_type when 'financial_documents' then jsonb_build_array('document_fat','supplier_nif','supplier','payment') when 'import_fdm_accounts' then jsonb_build_array('reservation_id','account','category') else jsonb_build_array('description') end),'filterFields',jsonb_build_array('dateFrom','dateTo','amountMin','amountMax','description','supplier','payment','account','category'),'reconciliation',case when p_reconciliation_id is null then null else to_jsonb(v_rec) end,'items',coalesce((select jsonb_agg(to_jsonb(i) order by i.created_at) from financial_reconciliation_items i where i.reconciliation_id=p_reconciliation_id),'[]'::jsonb),'audit',coalesce((select jsonb_agg(to_jsonb(a) order by a.created_at,a.id) from financial_reconciliation_audit a where a.reconciliation_id=p_reconciliation_id),'[]'::jsonb),'history',coalesce((select jsonb_agg(to_jsonb(h) order by h.created_at desc) from (select * from financial_reconciliations where deleted_at is null order by created_at desc limit 100) h),'[]'::jsonb));
end $$;

create or replace function public.financial_reconciliation_action(
  p_action text, p_actor text, p_reconciliation_id uuid default null, p_base_source_type text default null,
  p_matching_source_types text[] default null, p_source_type text default null, p_source_id uuid default null, p_comment text default null
) returns jsonb language plpgsql security definer set search_path = public as $$
declare r public.financial_reconciliations%rowtype; s record; v_id uuid; v_matching jsonb; v_diff numeric;
begin
  if p_action not in ('start','add_item','remove_item','complete','force_complete','reopen','delete') then raise exception 'Reconciliation action is invalid.'; end if;
  if nullif(trim(coalesce(p_actor,'')),'') is null then raise exception 'Actor is required.'; end if;
  if p_action='start' then
    if p_base_source_type not in ('financial_documents','import_fdm_accounts','import_cgd_cartao_credito','import_cgd_extrato_ordem') or p_source_type <> p_base_source_type then raise exception 'Start source type must match the base source type.'; end if;
    if p_matching_source_types is null or cardinality(p_matching_source_types)=0 or cardinality(p_matching_source_types)>3 or exists(select 1 from unnest(p_matching_source_types) x where x not in ('financial_documents','import_fdm_accounts','import_cgd_cartao_credito','import_cgd_extrato_ordem') or x=p_base_source_type) or cardinality(p_matching_source_types) <> cardinality(array(select distinct x from unnest(p_matching_source_types) x)) then raise exception 'Matching source types are invalid.'; end if;
    if p_base_source_type <> 'financial_documents' and cardinality(p_matching_source_types) <> 1 then raise exception 'Non-document bases require exactly one matching source type.'; end if;
    if (p_base_source_type='financial_documents' and exists(select 1 from unnest(p_matching_source_types) x where x not in ('import_fdm_accounts','import_cgd_cartao_credito','import_cgd_extrato_ordem'))) or (p_base_source_type='import_fdm_accounts' and p_matching_source_types <> array['import_cgd_extrato_ordem']) or (p_base_source_type='import_cgd_cartao_credito' and p_matching_source_types[1] not in ('financial_documents','import_cgd_extrato_ordem')) or (p_base_source_type='import_cgd_extrato_ordem' and p_matching_source_types[1] not in ('financial_documents','import_fdm_accounts','import_cgd_cartao_credito')) then raise exception 'Matching source type is not allowed for selected base.'; end if;
    select * into s from financial_reconciliation_source(p_source_type,p_source_id); if not found or not s.eligible then raise exception 'Source record is not eligible for reconciliation.'; end if;
    v_matching:=to_jsonb(p_matching_source_types); insert into financial_reconciliations(status,base_source_type,matching_source_types,created_by) values ('started',p_base_source_type,v_matching,p_actor) returning id into v_id;
    begin insert into financial_reconciliation_items(reconciliation_id,source_type,source_id,amount_snapshot,created_by) values(v_id,p_source_type,p_source_id,s.amount,p_actor); exception when unique_violation then raise exception 'This record is already reconciled.'; end;
    v_diff:=financial_reconciliation_difference(p_base_source_type,v_matching,v_id);
    update financial_reconciliations set difference_amount=v_diff,updated_at=timezone('utc',now()) where id=v_id;
    insert into financial_reconciliation_audit(reconciliation_id,action,actor,difference_amount,metadata) values(v_id,'start',p_actor,v_diff,jsonb_build_object('sourceType',p_source_type,'sourceId',p_source_id,'differenceAmount',v_diff));
  else
    select * into r from financial_reconciliations where id=p_reconciliation_id and deleted_at is null for update; if not found then raise exception 'Reconciliation not found.'; end if; v_id:=r.id;
    if p_action in ('add_item','remove_item','complete','force_complete') and r.status <> 'started' then raise exception 'Only started reconciliations can be edited or completed.'; end if;
    if p_action='reopen' and r.status <> 'complete' then raise exception 'Only complete reconciliations can be reopened.'; end if;
    if p_action='add_item' then
      if p_source_type <> r.base_source_type and not (r.matching_source_types ? p_source_type) then raise exception 'Item source type is not allowed for the selected reconciliation mode.'; end if;
      select * into s from financial_reconciliation_source(p_source_type,p_source_id); if not found or not s.eligible then raise exception 'Source record is not eligible for reconciliation.'; end if;
      begin insert into financial_reconciliation_items(reconciliation_id,source_type,source_id,amount_snapshot,created_by) values(v_id,p_source_type,p_source_id,s.amount,p_actor); exception when unique_violation then raise exception 'This record is already reconciled.'; end;
    elsif p_action='remove_item' then delete from financial_reconciliation_items where reconciliation_id=v_id and source_type=p_source_type and source_id=p_source_id; if not found then raise exception 'Reconciliation item not found.'; end if;
    elsif p_action in ('complete','force_complete') then
      v_diff:=financial_reconciliation_difference(r.base_source_type,r.matching_source_types,v_id);
      if p_action='complete' and v_diff <> 0 then raise exception 'Complete requires a zero difference.'; end if;
      if p_action='force_complete' and (v_diff=0 or nullif(trim(coalesce(p_comment,'')),'') is null) then raise exception 'Force complete requires a non-zero difference and a comment.'; end if;
      update financial_reconciliations set status='complete', completion_type=case when p_action='complete' then 'normal' else 'forced' end, difference_amount=v_diff, forced_completion_comment=case when p_action='force_complete' then trim(p_comment) else null end, completed_by=p_actor, completed_at=timezone('utc',now()), updated_at=timezone('utc',now()) where id=v_id;
    elsif p_action='reopen' then update financial_reconciliations set status='started',completion_type=null,forced_completion_comment=null,completed_by=null,completed_at=null,updated_at=timezone('utc',now()) where id=v_id;
    elsif p_action='delete' then
      v_diff:=financial_reconciliation_difference(r.base_source_type,r.matching_source_types,v_id);
      update financial_reconciliations set deleted_by=p_actor,deleted_at=timezone('utc',now()),updated_at=timezone('utc',now()) where id=v_id;
      delete from financial_reconciliation_items where reconciliation_id=v_id;
    end if;
    if p_action <> 'delete' then
      v_diff:=financial_reconciliation_difference(r.base_source_type,r.matching_source_types,v_id);
      update financial_reconciliations set difference_amount=v_diff,updated_at=timezone('utc',now()) where id=v_id;
    end if;
    insert into financial_reconciliation_audit(reconciliation_id,action,actor,comment,difference_amount,metadata) values(v_id,p_action,p_actor,nullif(trim(coalesce(p_comment,'')),''),v_diff,jsonb_build_object('sourceType',p_source_type,'sourceId',p_source_id,'differenceAmount',v_diff));
  end if;
  if p_action = 'delete' then return jsonb_build_object('reconciliationId',v_id,'deleted',true,'differenceAmount',v_diff); end if;
  select * into r from financial_reconciliations where id=v_id;
  return get_financial_reconciliation_workspace(v_id, coalesce(p_source_type,r.base_source_type), array(select jsonb_array_elements_text(r.matching_source_types)));
end $$;

grant execute on function public.get_financial_reconciliation_workspace(uuid,text,text[],jsonb,integer,integer) to authenticated, service_role;
revoke all on function public.financial_reconciliation_action(text,text,uuid,text,text[],text,uuid,text) from public, anon, authenticated;
grant execute on function public.financial_reconciliation_action(text,text,uuid,text,text[],text,uuid,text) to service_role;
