alter table if exists public.guest_records
  alter column check_in drop not null,
  alter column check_out drop not null;
