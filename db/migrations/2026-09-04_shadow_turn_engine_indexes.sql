-- Support occupancy-bounded shadow-turn aggregation without blocking writes.

CREATE INDEX CONCURRENTLY IF NOT EXISTS appfolio_work_orders_unit_created_idx
  ON appfolio_work_orders(unit_id, created_at)
  WHERE unit_id IS NOT NULL AND created_at IS NOT NULL;

CREATE INDEX CONCURRENTLY IF NOT EXISTS appfolio_tenant_directory_unit_moveout_idx
  ON appfolio_tenant_directory(unit_id, move_out_date)
  WHERE unit_id IS NOT NULL;

CREATE INDEX CONCURRENTLY IF NOT EXISTS appfolio_tenant_directory_unit_movein_status_idx
  ON appfolio_tenant_directory(unit_id, move_in_date, lower(status))
  WHERE unit_id IS NOT NULL;

CREATE INDEX CONCURRENTLY IF NOT EXISTS appfolio_unit_turn_details_unit_moveout_idx
  ON appfolio_unit_turn_details(unit_id, move_out_date)
  WHERE unit_id IS NOT NULL AND move_out_date IS NOT NULL;

CREATE INDEX CONCURRENTLY IF NOT EXISTS appfolio_unit_inspections_unit_date_idx
  ON appfolio_unit_inspections(unit_id, last_inspection_date)
  WHERE unit_id IS NOT NULL AND last_inspection_date IS NOT NULL;

CREATE INDEX CONCURRENTLY IF NOT EXISTS appfolio_unit_vacancies_unit_updated_idx
  ON appfolio_unit_vacancies(unit_id, last_updated_at DESC, cached_at DESC)
  WHERE unit_id IS NOT NULL;
