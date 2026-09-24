create table if not exists public.financial_document_supplier_email_schedules (
  id uuid primary key default gen_random_uuid(),
  supplier_nif text not null,
  supplier_name text not null,
  recipients jsonb not null default '[]'::jsonb,
  day_of_month integer not null check (day_of_month between 1 and 31),
  subject text not null check (btrim(subject) <> ''),
  enabled boolean not null default true,
  created_by text not null default '',
  updated_by text not null default '',
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now()),
  unique (supplier_nif)
);

create table if not exists public.financial_document_supplier_email_runs (
  id uuid primary key default gen_random_uuid(),
  schedule_id uuid not null references public.financial_document_supplier_email_schedules(id) on delete cascade,
  period_start date not null,
  period_end date not null,
  attempt_number integer not null default 1 check (attempt_number > 0),
  trigger_type text not null default 'cron' check (trigger_type in ('cron', 'retry')),
  status text not null default 'running' check (status in ('running', 'sent', 'failed', 'skipped', 'uncertain')),
  recipients_snapshot jsonb not null default '[]'::jsonb,
  subject_snapshot text not null default '',
  document_snapshot jsonb not null default '[]'::jsonb,
  attachment_count integer not null default 0,
  attachment_raw_bytes bigint not null default 0,
  resend_message_id text not null default '',
  resend_idempotency_key text not null default '',
  error_code text not null default '',
  error_message text not null default '',
  retryable boolean not null default false,
  retry_of_run_id uuid references public.financial_document_supplier_email_runs(id),
  started_at timestamptz not null default timezone('utc', now()),
  completed_at timestamptz,
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now()),
  unique (schedule_id, period_start, attempt_number),
  check (period_end >= period_start)
);

create index if not exists financial_document_supplier_email_runs_schedule_period_idx
  on public.financial_document_supplier_email_runs (schedule_id, period_start desc, attempt_number desc);

create unique index if not exists financial_document_supplier_email_active_delivery_idx
  on public.financial_document_supplier_email_runs (schedule_id, period_start)
  where status in ('running', 'sent', 'skipped', 'uncertain');

drop trigger if exists financial_document_supplier_email_schedules_set_updated_at on public.financial_document_supplier_email_schedules;
create trigger financial_document_supplier_email_schedules_set_updated_at
before update on public.financial_document_supplier_email_schedules
for each row execute procedure public.set_updated_at();

drop trigger if exists financial_document_supplier_email_runs_set_updated_at on public.financial_document_supplier_email_runs;
create trigger financial_document_supplier_email_runs_set_updated_at
before update on public.financial_document_supplier_email_runs
for each row execute procedure public.set_updated_at();

alter table public.financial_document_supplier_email_schedules enable row level security;
alter table public.financial_document_supplier_email_runs enable row level security;

create or replace function public.claim_financial_document_supplier_email_run(
  p_schedule_id uuid,
  p_period_start date,
  p_period_end date,
  p_trigger_type text default 'cron'
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_schedule public.financial_document_supplier_email_schedules%rowtype;
  v_existing public.financial_document_supplier_email_runs%rowtype;
  v_attempt integer;
  v_run public.financial_document_supplier_email_runs%rowtype;
begin
  select * into v_schedule
  from public.financial_document_supplier_email_schedules
  where id = p_schedule_id
  for update;
  if not found then
    raise exception 'Supplier email schedule not found.';
  end if;

  select * into v_existing
  from public.financial_document_supplier_email_runs
  where schedule_id = p_schedule_id
    and period_start = p_period_start
    and status in ('running', 'sent', 'skipped', 'uncertain')
  order by attempt_number desc
  limit 1;
  if found then
    return jsonb_build_object('claimed', false, 'reason', 'already_processed', 'run', to_jsonb(v_existing));
  end if;

  select coalesce(max(attempt_number), 0) + 1 into v_attempt
  from public.financial_document_supplier_email_runs
  where schedule_id = p_schedule_id and period_start = p_period_start;

  insert into public.financial_document_supplier_email_runs (
    schedule_id, period_start, period_end, attempt_number, trigger_type, recipients_snapshot,
    subject_snapshot, resend_idempotency_key
  ) values (
    p_schedule_id, p_period_start, p_period_end, v_attempt,
    case when p_trigger_type = 'retry' then 'retry' else 'cron' end,
    v_schedule.recipients, v_schedule.subject,
    format('financial-document-supplier-email/%s/%s', p_schedule_id, to_char(p_period_start, 'YYYY-MM'))
  ) returning * into v_run;

  return jsonb_build_object('claimed', true, 'reason', 'claimed', 'run', to_jsonb(v_run));
end;
$$;

create or replace function public.complete_financial_document_supplier_email_run(
  p_run_id uuid,
  p_status text,
  p_document_snapshot jsonb default '[]'::jsonb,
  p_attachment_count integer default 0,
  p_attachment_raw_bytes bigint default 0,
  p_resend_message_id text default '',
  p_error_code text default '',
  p_error_message text default '',
  p_retryable boolean default false
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_run public.financial_document_supplier_email_runs%rowtype;
begin
  if p_status not in ('sent', 'failed', 'skipped', 'uncertain') then
    raise exception 'Invalid supplier email run status.';
  end if;
  update public.financial_document_supplier_email_runs
  set status = p_status,
      document_snapshot = coalesce(p_document_snapshot, '[]'::jsonb),
      attachment_count = greatest(coalesce(p_attachment_count, 0), 0),
      attachment_raw_bytes = greatest(coalesce(p_attachment_raw_bytes, 0), 0),
      resend_message_id = coalesce(p_resend_message_id, ''),
      error_code = coalesce(p_error_code, ''),
      error_message = coalesce(p_error_message, ''),
      retryable = case when p_status = 'failed' then coalesce(p_retryable, false) else false end,
      completed_at = timezone('utc', now())
  where id = p_run_id and status = 'running'
  returning * into v_run;
  if not found then
    raise exception 'Supplier email run is not running.';
  end if;
  return to_jsonb(v_run);
end;
$$;

create or replace function public.retry_financial_document_supplier_email_run(p_run_id uuid)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_previous public.financial_document_supplier_email_runs%rowtype;
begin
  select * into v_previous
  from public.financial_document_supplier_email_runs
  where id = p_run_id
  for update;
  if not found then
    raise exception 'Supplier email run not found.';
  end if;
  if v_previous.status <> 'failed' or not v_previous.retryable then
    raise exception 'This supplier email run cannot be retried.';
  end if;
  return public.claim_financial_document_supplier_email_run(
    v_previous.schedule_id, v_previous.period_start, v_previous.period_end, 'retry'
  );
end;
$$;

revoke all on public.financial_document_supplier_email_schedules from anon, authenticated;
revoke all on public.financial_document_supplier_email_runs from anon, authenticated;
revoke all on function public.claim_financial_document_supplier_email_run(uuid, date, date, text) from public;
revoke all on function public.complete_financial_document_supplier_email_run(uuid, text, jsonb, integer, bigint, text, text, text, boolean) from public;
revoke all on function public.retry_financial_document_supplier_email_run(uuid) from public;
grant execute on function public.claim_financial_document_supplier_email_run(uuid, date, date, text) to service_role;
grant execute on function public.complete_financial_document_supplier_email_run(uuid, text, jsonb, integer, bigint, text, text, text, boolean) to service_role;
grant execute on function public.retry_financial_document_supplier_email_run(uuid) to service_role;
