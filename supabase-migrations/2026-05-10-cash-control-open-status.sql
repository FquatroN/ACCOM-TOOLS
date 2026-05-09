alter table if exists public.cash_control_records
  add column if not exists status text;

update public.cash_control_records
set status = 'C'
where status is null
   or btrim(status) = ''
   or upper(status) not in ('O', 'C');

alter table public.cash_control_records
  alter column status set default 'C';

alter table public.cash_control_records
  alter column status set not null;

do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conname = 'cash_control_records_status_check'
  ) then
    alter table public.cash_control_records
      add constraint cash_control_records_status_check
      check (status in ('O', 'C'));
  end if;
end $$;
