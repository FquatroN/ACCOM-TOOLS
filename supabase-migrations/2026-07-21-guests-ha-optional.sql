-- Allow guest records without an HA value. Existing H/A values are preserved.
alter table if exists public.guest_records
  alter column ha drop not null,
  alter column ha drop default;

alter table if exists public.guest_records
  drop constraint if exists guest_records_ha_check;

alter table if exists public.guest_records
  add constraint guest_records_ha_check check (ha is null or ha in ('H', 'A'));
