-- Persisted, directional reconciliation rules. Started reconciliations snapshot
-- these rows so later configuration changes cannot alter an in-progress group.
create table if not exists public.financial_reconciliation_source_rules (
  base_source_type text not null check (base_source_type in ('financial_documents','import_fdm_accounts','import_cgd_cartao_credito','import_cgd_extrato_ordem')),
  matching_source_type text not null check (matching_source_type in ('financial_documents','import_fdm_accounts','import_cgd_cartao_credito','import_cgd_extrato_ordem')),
  operator text not null check (operator in ('+','-')),
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now()),
  primary key (base_source_type, matching_source_type),
  check (base_source_type <> matching_source_type)
);

alter table public.financial_reconciliations
  add column if not exists matching_source_rules jsonb not null default '[]'::jsonb;

insert into public.financial_reconciliation_source_rules (base_source_type, matching_source_type, operator) values
  ('financial_documents', 'import_fdm_accounts', '+'),
  ('financial_documents', 'import_cgd_cartao_credito', '+'),
  ('financial_documents', 'import_cgd_extrato_ordem', '+'),
  ('import_fdm_accounts', 'import_cgd_extrato_ordem', '-'),
  ('import_cgd_cartao_credito', 'financial_documents', '+'),
  ('import_cgd_cartao_credito', 'import_cgd_extrato_ordem', '+'),
  ('import_cgd_extrato_ordem', 'financial_documents', '+'),
  ('import_cgd_extrato_ordem', 'import_fdm_accounts', '-'),
  ('import_cgd_extrato_ordem', 'import_cgd_cartao_credito', '+')
on conflict (base_source_type, matching_source_type) do nothing;

update public.financial_reconciliations r
set matching_source_rules = coalesce((
  select jsonb_agg(jsonb_build_object('sourceType', old_rule.source_type, 'operator', legacy_rule.operator) order by old_rule.ordinality)
  from jsonb_array_elements_text(r.matching_source_types) with ordinality old_rule(source_type, ordinality)
  cross join lateral (
    select case
      when r.base_source_type = 'financial_documents' and old_rule.source_type in ('import_fdm_accounts','import_cgd_cartao_credito','import_cgd_extrato_ordem') then '+'
      when r.base_source_type = 'import_fdm_accounts' and old_rule.source_type = 'import_cgd_extrato_ordem' then '-'
      when r.base_source_type = 'import_cgd_cartao_credito' and old_rule.source_type in ('financial_documents','import_cgd_extrato_ordem') then '+'
      when r.base_source_type = 'import_cgd_extrato_ordem' and old_rule.source_type = 'import_fdm_accounts' then '-'
      when r.base_source_type = 'import_cgd_extrato_ordem' and old_rule.source_type in ('financial_documents','import_cgd_cartao_credito') then '+'
    end as operator
  ) legacy_rule
  where legacy_rule.operator is not null
), '[]'::jsonb)
where r.matching_source_rules = '[]'::jsonb
  and jsonb_array_length(r.matching_source_types) > 0;

create or replace function public.financial_reconciliation_difference(p_base text, p_rules jsonb, p_reconciliation_id uuid)
returns numeric language plpgsql stable security definer set search_path = public as $$
begin
  if p_base not in ('financial_documents','import_fdm_accounts','import_cgd_cartao_credito','import_cgd_extrato_ordem') then
    raise exception 'Source type is invalid.';
  end if;
  if exists (
    select 1
    from public.financial_reconciliation_items i
    where i.reconciliation_id = p_reconciliation_id
      and i.source_type <> p_base
      and not exists (
        select 1 from jsonb_array_elements(coalesce(p_rules, '[]'::jsonb)) rule
        where rule->>'sourceType' = i.source_type and rule->>'operator' in ('+', '-')
      )
  ) then
    raise exception 'Item source type is not allowed for this reconciliation.';
  end if;
  return round(coalesce((
    select sum(case
      when i.source_type = p_base then i.amount_snapshot
      when rule.value->>'operator' = '+' then i.amount_snapshot
      else -i.amount_snapshot
    end)
    from public.financial_reconciliation_items i
    left join lateral (
      select value from jsonb_array_elements(coalesce(p_rules, '[]'::jsonb)) value
      where value->>'sourceType' = i.source_type
      limit 1
    ) rule on true
    where i.reconciliation_id = p_reconciliation_id
  ), 0), 2);
end $$;

create or replace function public.get_financial_reconciliation_workspace(
  p_reconciliation_id uuid, p_source_type text, p_filters jsonb default '{}'::jsonb,
  p_page integer default 1, p_page_size integer default 50
) returns jsonb language plpgsql security definer set search_path = public as $$
declare v_rec public.financial_reconciliations%rowtype; v_rules jsonb; v_offset int; v_candidates jsonb; v_count int; v_started_count int; v_complete_count int;
begin
  if p_source_type not in ('financial_documents','import_fdm_accounts','import_cgd_cartao_credito','import_cgd_extrato_ordem') then raise exception 'Source type is invalid.'; end if;
  if p_page < 1 or p_page_size not between 1 and 100 then raise exception 'Page and page size are invalid.'; end if;
  if p_reconciliation_id is not null then
    select * into v_rec from public.financial_reconciliations where id=p_reconciliation_id and deleted_at is null;
    if not found then raise exception 'Reconciliation not found.'; end if;
    v_rules := v_rec.matching_source_rules;
    if p_source_type <> v_rec.base_source_type and not exists (select 1 from jsonb_array_elements(v_rules) rule where rule->>'sourceType'=p_source_type) then
      raise exception 'Source type is not allowed for this reconciliation.';
    end if;
  else
    select coalesce(jsonb_agg(jsonb_build_object('sourceType', matching_source_type, 'operator', operator) order by matching_source_type), '[]'::jsonb)
      into v_rules from public.financial_reconciliation_source_rules where base_source_type=p_source_type;
  end if;
  v_offset := (p_page - 1) * p_page_size;
  with candidate_rows as (
    select s.* from (
      select d.id,d.amount,d.document_date as source_date,d.description,d.supplier_name as supplier,d.cc as account,d.category,d.payment,d.fat as document_fat,d.supplier_nif,''::text as reservation_id from public.financial_documents d where p_source_type='financial_documents' and d.document_date >= date '2026-01-01' and d.fat='S'
      union all select f.id,f.amount,f.event_date,f.description,f.guest,f.account,f.category,f.invoice,''::text,''::text,f.reservation_id from public.import_fdm_accounts f where p_source_type='import_fdm_accounts' and f.event_date >= date '2026-01-01' and (coalesce(f.invoice_flag,false) or f.category='Compras') and (coalesce(v_rec.base_source_type,'') <> 'financial_documents' or f.category='Compras')
      union all select c.id,c.valor,c.data,c.descricao,''::text,''::text,''::text,''::text,''::text,''::text,''::text from public.import_cgd_cartao_credito c where p_source_type='import_cgd_cartao_credito' and c.data >= date '2026-01-01'
      union all select b.id,b.montante,b.data,b.descritivo,''::text,''::text,''::text,''::text,''::text,''::text,''::text from public.import_cgd_extrato_ordem b where p_source_type='import_cgd_extrato_ordem' and b.data >= date '2026-01-01'
    ) s where not exists (select 1 from public.financial_reconciliation_items i join public.financial_reconciliations r on r.id=i.reconciliation_id where i.source_type=p_source_type and i.source_id=s.id and r.deleted_at is null)
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
  select count(*) into v_started_count from public.financial_reconciliation_items i join public.financial_reconciliations r on r.id=i.reconciliation_id where i.source_type=p_source_type and r.deleted_at is null and r.status='started';
  select count(*) into v_complete_count from public.financial_reconciliation_items i join public.financial_reconciliations r on r.id=i.reconciliation_id where i.source_type=p_source_type and r.deleted_at is null and r.status='complete';
  with candidate_rows as (
    select s.* from (
      select d.id,d.amount,d.document_date as source_date,d.description,d.supplier_name as supplier,d.cc as account,d.category,d.payment,d.fat as document_fat,d.supplier_nif,''::text as reservation_id from public.financial_documents d where p_source_type='financial_documents' and d.document_date >= date '2026-01-01' and d.fat='S'
      union all select f.id,f.amount,f.event_date,f.description,f.guest,f.account,f.category,f.invoice,''::text,''::text,f.reservation_id from public.import_fdm_accounts f where p_source_type='import_fdm_accounts' and f.event_date >= date '2026-01-01' and (coalesce(f.invoice_flag,false) or f.category='Compras') and (coalesce(v_rec.base_source_type,'') <> 'financial_documents' or f.category='Compras')
      union all select c.id,c.valor,c.data,c.descricao,''::text,''::text,''::text,''::text,''::text,''::text,''::text from public.import_cgd_cartao_credito c where p_source_type='import_cgd_cartao_credito' and c.data >= date '2026-01-01'
      union all select b.id,b.montante,b.data,b.descritivo,''::text,''::text,''::text,''::text,''::text,''::text,''::text from public.import_cgd_extrato_ordem b where p_source_type='import_cgd_extrato_ordem' and b.data >= date '2026-01-01'
    ) s where not exists (select 1 from public.financial_reconciliation_items i join public.financial_reconciliations r on r.id=i.reconciliation_id where i.source_type=p_source_type and i.source_id=s.id and r.deleted_at is null)
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
  return jsonb_build_object('candidates',v_candidates,'totalCount',v_count,'counts',jsonb_build_object('notStarted',v_count,'started',v_started_count,'complete',v_complete_count),'page',p_page,'pageSize',p_page_size,'sourceConfig',jsonb_build_object('sourceType',p_source_type,'dateColumn','source_date','amountColumn','amount','columns',case p_source_type when 'financial_documents' then jsonb_build_array('document_fat','supplier_nif','supplier','payment') when 'import_fdm_accounts' then jsonb_build_array('reservation_id','account','category') else jsonb_build_array('description') end,'filterFields',case p_source_type when 'financial_documents' then jsonb_build_array('dateFrom','dateTo','amountMin','amountMax','description','supplier','payment','account','category') when 'import_fdm_accounts' then jsonb_build_array('dateFrom','dateTo','amountMin','amountMax','description','account','category') else jsonb_build_array('dateFrom','dateTo','amountMin','amountMax','description') end),'rules',case when p_reconciliation_id is null then v_rules else '[]'::jsonb end,'reconciliation',case when p_reconciliation_id is null then null else to_jsonb(v_rec) || jsonb_build_object('matchingSourceRules',v_rules) end,'matchingSourceRules',case when p_reconciliation_id is null then '[]'::jsonb else v_rules end,'items',coalesce((select jsonb_agg(to_jsonb(i) || jsonb_build_object('source_date',s.source_date,'description',s.description,'supplier',nullif(s.supplier,'')) order by i.created_at) from public.financial_reconciliation_items i left join lateral public.financial_reconciliation_source(i.source_type,i.source_id) s on true where i.reconciliation_id=p_reconciliation_id),'[]'::jsonb),'audit',coalesce((select jsonb_agg(to_jsonb(a) order by a.created_at,a.id) from public.financial_reconciliation_audit a where a.reconciliation_id=p_reconciliation_id),'[]'::jsonb),'history',coalesce((select jsonb_agg(to_jsonb(h) order by h.created_at desc) from (select * from public.financial_reconciliations where deleted_at is null order by created_at desc limit 100) h),'[]'::jsonb));
end $$;

create or replace function public.financial_reconciliation_action(
  p_action text, p_actor text, p_reconciliation_id uuid default null, p_source_type text default null,
  p_source_id uuid default null, p_comment text default null
) returns jsonb language plpgsql security definer set search_path = public as $$
declare r public.financial_reconciliations%rowtype; s record; v_id uuid; v_rules jsonb; v_matching jsonb; v_diff numeric;
begin
  if p_action not in ('start','add_item','remove_item','complete','force_complete','reopen','delete') then raise exception 'Reconciliation action is invalid.'; end if;
  if nullif(trim(coalesce(p_actor,'')),'') is null then raise exception 'Actor is required.'; end if;
  if p_action='start' then
    if p_source_type not in ('financial_documents','import_fdm_accounts','import_cgd_cartao_credito','import_cgd_extrato_ordem') then raise exception 'Source type is invalid.'; end if;
    select coalesce(jsonb_agg(jsonb_build_object('sourceType', matching_source_type, 'operator', operator) order by matching_source_type), '[]'::jsonb)
      into v_rules from public.financial_reconciliation_source_rules where base_source_type=p_source_type;
    if v_rules = '[]'::jsonb then raise exception 'No reconciliation rules are configured for this source.'; end if;
    select * into s from public.financial_reconciliation_source(p_source_type,p_source_id); if not found or not s.eligible then raise exception 'Source record is not eligible for reconciliation.'; end if;
    v_matching := coalesce((select jsonb_agg(rule->>'sourceType' order by rule->>'sourceType') from jsonb_array_elements(v_rules) rule), '[]'::jsonb);
    insert into public.financial_reconciliations(status,base_source_type,matching_source_types,matching_source_rules,created_by) values ('started',p_source_type,v_matching,v_rules,p_actor) returning id into v_id;
    begin insert into public.financial_reconciliation_items(reconciliation_id,source_type,source_id,amount_snapshot,created_by) values(v_id,p_source_type,p_source_id,s.amount,p_actor); exception when unique_violation then raise exception 'This record is already reconciled.'; end;
    v_diff:=public.financial_reconciliation_difference(p_source_type,v_rules,v_id);
    update public.financial_reconciliations set difference_amount=v_diff,updated_at=timezone('utc',now()) where id=v_id;
    insert into public.financial_reconciliation_audit(reconciliation_id,action,actor,difference_amount,metadata) values(v_id,'start',p_actor,v_diff,jsonb_build_object('sourceType',p_source_type,'sourceId',p_source_id,'differenceAmount',v_diff,'matchingSourceRules',v_rules));
  else
    select * into r from public.financial_reconciliations where id=p_reconciliation_id and deleted_at is null for update; if not found then raise exception 'Reconciliation not found.'; end if; v_id:=r.id;
    if p_action in ('add_item','remove_item','complete','force_complete') and r.status <> 'started' then raise exception 'Only started reconciliations can be edited or completed.'; end if;
    if p_action='reopen' and r.status <> 'complete' then raise exception 'Only complete reconciliations can be reopened.'; end if;
    if p_action='add_item' then
      if p_source_type <> r.base_source_type and not exists (select 1 from jsonb_array_elements(r.matching_source_rules) rule where rule->>'sourceType'=p_source_type) then raise exception 'Item source type is not allowed for this reconciliation.'; end if;
      select * into s from public.financial_reconciliation_source(p_source_type,p_source_id); if not found or not s.eligible then raise exception 'Source record is not eligible for reconciliation.'; end if;
      if r.base_source_type='financial_documents' and p_source_type='import_fdm_accounts' and s.category <> 'Compras' then raise exception 'Financial Documents-led reconciliations require FDM category Compras.'; end if;
      begin insert into public.financial_reconciliation_items(reconciliation_id,source_type,source_id,amount_snapshot,created_by) values(v_id,p_source_type,p_source_id,s.amount,p_actor); exception when unique_violation then raise exception 'This record is already reconciled.'; end;
    elsif p_action='remove_item' then delete from public.financial_reconciliation_items where reconciliation_id=v_id and source_type=p_source_type and source_id=p_source_id; if not found then raise exception 'Reconciliation item not found.'; end if;
    elsif p_action in ('complete','force_complete') then
      v_diff:=public.financial_reconciliation_difference(r.base_source_type,r.matching_source_rules,v_id);
      if p_action='complete' and v_diff <> 0 then raise exception 'Complete requires a zero difference.'; end if;
      if p_action='force_complete' and (v_diff=0 or nullif(trim(coalesce(p_comment,'')),'') is null) then raise exception 'Force complete requires a non-zero difference and a comment.'; end if;
      update public.financial_reconciliations set status='complete',completion_type=case when p_action='complete' then 'normal' else 'forced' end,difference_amount=v_diff,forced_completion_comment=case when p_action='force_complete' then trim(p_comment) else null end,completed_by=p_actor,completed_at=timezone('utc',now()),updated_at=timezone('utc',now()) where id=v_id;
    elsif p_action='reopen' then update public.financial_reconciliations set status='started',completion_type=null,forced_completion_comment=null,completed_by=null,completed_at=null,updated_at=timezone('utc',now()) where id=v_id;
    elsif p_action='delete' then
      v_diff:=public.financial_reconciliation_difference(r.base_source_type,r.matching_source_rules,v_id);
      update public.financial_reconciliations set deleted_by=p_actor,deleted_at=timezone('utc',now()),updated_at=timezone('utc',now()) where id=v_id;
      delete from public.financial_reconciliation_items where reconciliation_id=v_id;
    end if;
    if p_action <> 'delete' then
      v_diff:=public.financial_reconciliation_difference(r.base_source_type,r.matching_source_rules,v_id);
      update public.financial_reconciliations set difference_amount=v_diff,updated_at=timezone('utc',now()) where id=v_id;
    end if;
    insert into public.financial_reconciliation_audit(reconciliation_id,action,actor,comment,difference_amount,metadata) values(v_id,p_action,p_actor,nullif(trim(coalesce(p_comment,'')),''),v_diff,jsonb_build_object('sourceType',p_source_type,'sourceId',p_source_id,'differenceAmount',v_diff));
  end if;
  if p_action = 'delete' then return jsonb_build_object('reconciliationId',v_id,'deleted',true,'differenceAmount',v_diff); end if;
  select * into r from public.financial_reconciliations where id=v_id;
  return public.get_financial_reconciliation_workspace(v_id,coalesce(p_source_type,r.base_source_type),'{}'::jsonb,1,50);
end $$;

revoke all on function public.get_financial_reconciliation_workspace(uuid,text,text[],jsonb,integer,integer) from public, anon, authenticated, service_role;
revoke all on function public.financial_reconciliation_action(text,text,uuid,text,text[],text,uuid,text) from public, anon, authenticated, service_role;
revoke all on function public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer) from public, anon, authenticated;
revoke all on function public.financial_reconciliation_action(text,text,uuid,text,uuid,text) from public, anon, authenticated;
grant execute on function public.get_financial_reconciliation_workspace(uuid,text,jsonb,integer,integer) to service_role;
grant execute on function public.financial_reconciliation_action(text,text,uuid,text,uuid,text) to service_role;
