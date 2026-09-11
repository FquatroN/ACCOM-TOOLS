-- Consolidate legacy Travelocity reviews under the Expedia source.
-- The unique review-source index will safely reject this migration if a
-- Travelocity row conflicts with an existing Expedia review ID.
update public.reviews
set source = 'expedia'
where lower(btrim(source)) = 'travelocity';
