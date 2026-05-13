grant usage on schema public to anon;
grant usage on schema public to authenticated;
grant usage on schema public to service_role;

do $$
declare
  table_name text;
  app_tables text[] := array[
    'communications',
    'lost_found',
    'app_settings',
    'app_profiles',
    'user_profile_assignments',
    'group_proposals',
    'services',
    'shopping_orders',
    'bakery_orders',
    'laundry_records',
    'hours_register_records',
    'properties',
    'review_import_runs',
    'review_import_staging',
    'reviews',
    'cash_control_records',
    'guest_records',
    'guests_blacklist',
    'guest_api_calls'
  ];
begin
  foreach table_name in array app_tables loop
    if to_regclass(format('public.%I', table_name)) is null then
      continue;
    end if;
    execute format('grant select, insert, update, delete on table public.%I to anon', table_name);
    execute format('grant select, insert, update, delete on table public.%I to authenticated', table_name);
    execute format('grant select, insert, update, delete on table public.%I to service_role', table_name);
  end loop;
end $$;
