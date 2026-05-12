alter table if exists public.guest_api_calls
  add column if not exists response_headers jsonb not null default '{}'::jsonb;
