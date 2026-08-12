-- The source-rules migration changed the action RPC signature. Remove the
-- previous eight-parameter overload so PostgREST can resolve Start/Add calls.
drop function if exists public.financial_reconciliation_action(text,text,uuid,text,text[],text,uuid,text);
notify pgrst, 'reload schema';
