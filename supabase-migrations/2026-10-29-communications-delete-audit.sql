create table if not exists public.communications_audit_events (
  id uuid primary key default gen_random_uuid(),
  action text not null check (action = 'deleted'),
  communication_id uuid not null,
  actor_user_id uuid not null,
  actor_email text not null check (btrim(actor_email) <> ''),
  occurred_at timestamptz not null default timezone('utc', now()),
  record_snapshot jsonb not null
);

create index if not exists communications_audit_events_occurred_at_idx
  on public.communications_audit_events (occurred_at desc);

alter table public.communications_audit_events enable row level security;

revoke all on table public.communications_audit_events from public, anon, authenticated;

create or replace function public.delete_communication_with_audit(
  p_communication_id uuid,
  p_actor_user_id uuid,
  p_actor_email text
)
returns setof public.communications
language plpgsql
security definer
set search_path = public
as $$
declare
  v_communication public.communications%rowtype;
begin
  if p_communication_id is null or p_actor_user_id is null or btrim(coalesce(p_actor_email, '')) = '' then
    raise exception 'A communication ID and verified actor details are required.';
  end if;

  select * into v_communication
  from public.communications
  where id = p_communication_id
  for update;

  if not found then
    return;
  end if;

  insert into public.communications_audit_events (
    action,
    communication_id,
    actor_user_id,
    actor_email,
    record_snapshot
  ) values (
    'deleted',
    v_communication.id,
    p_actor_user_id,
    lower(btrim(p_actor_email)),
    to_jsonb(v_communication)
  );

  return query
  delete from public.communications
  where id = v_communication.id
  returning *;
end;
$$;

revoke all on function public.delete_communication_with_audit(uuid, uuid, text) from public, anon, authenticated;
grant execute on function public.delete_communication_with_audit(uuid, uuid, text) to service_role;
