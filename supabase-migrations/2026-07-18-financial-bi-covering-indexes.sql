create index if not exists financial_documents_bi_cover_idx
  on public.financial_documents (document_date, cc, category)
  include (amount);

create index if not exists import_fdm_occupancy_kpi_bi_cover_idx
  on public.import_fdm_occupancy_kpi (mes, property)
  include (charge);

create index if not exists import_fdm_sales_bi_cover_idx
  on public.import_fdm_sales (sale_date, reservation_id, sale_item)
  include (total);

create index if not exists import_fdm_bookings_bi_cover_idx
  on public.import_fdm_bookings (booking_number, created_at desc)
  include (room_type);

analyze public.financial_documents;
analyze public.import_fdm_occupancy_kpi;
analyze public.import_fdm_sales;
analyze public.import_fdm_bookings;
