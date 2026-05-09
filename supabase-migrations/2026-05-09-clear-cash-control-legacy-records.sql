update public.app_settings
set
  payload = jsonb_set(coalesce(payload, '{}'::jsonb), '{records}', '[]'::jsonb, true),
  updated_at = now()
where setting_key = 'cash_control';
