-- Repair additive schema drift for the Postgres sync control plane and the
-- normalized AppFolio v2 report populations. Safe to re-run.

ALTER TABLE IF EXISTS sync_job_runs
  ADD COLUMN IF NOT EXISTS execution_start_cursor TEXT;

ALTER TABLE IF EXISTS appfolio_unit_inspections
  ADD COLUMN IF NOT EXISTS property_id TEXT,
  ADD COLUMN IF NOT EXISTS property_name TEXT,
  ADD COLUMN IF NOT EXISTS unit_id TEXT,
  ADD COLUMN IF NOT EXISTS unit_name TEXT,
  ADD COLUMN IF NOT EXISTS last_inspection_date TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS tenant_name TEXT,
  ADD COLUMN IF NOT EXISTS tenant_primary_phone_number TEXT,
  ADD COLUMN IF NOT EXISTS move_in_date TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS move_out_date TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS rentable TEXT,
  ADD COLUMN IF NOT EXISTS occupancy_id TEXT,
  ADD COLUMN IF NOT EXISTS unit_tags TEXT,
  ADD COLUMN IF NOT EXISTS raw_json JSONB NOT NULL DEFAULT '{}',
  ADD COLUMN IF NOT EXISTS last_updated_at TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS cached_at TIMESTAMPTZ NOT NULL DEFAULT NOW();

ALTER TABLE IF EXISTS appfolio_tenant_directory
  ADD COLUMN IF NOT EXISTS property_id TEXT,
  ADD COLUMN IF NOT EXISTS property_name TEXT,
  ADD COLUMN IF NOT EXISTS unit_id TEXT,
  ADD COLUMN IF NOT EXISTS unit_name TEXT,
  ADD COLUMN IF NOT EXISTS tenant_name TEXT,
  ADD COLUMN IF NOT EXISTS status TEXT,
  ADD COLUMN IF NOT EXISTS tenant_type TEXT,
  ADD COLUMN IF NOT EXISTS phone_numbers TEXT,
  ADD COLUMN IF NOT EXISTS emails TEXT,
  ADD COLUMN IF NOT EXISTS move_in_date TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS lease_to TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS move_out_date TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS rent TEXT,
  ADD COLUMN IF NOT EXISTS tenant_tags TEXT,
  ADD COLUMN IF NOT EXISTS tenant_agent TEXT,
  ADD COLUMN IF NOT EXISTS tenant_visibility TEXT,
  ADD COLUMN IF NOT EXISTS occupancy_id TEXT,
  ADD COLUMN IF NOT EXISTS unit_tags TEXT,
  ADD COLUMN IF NOT EXISTS raw_json JSONB NOT NULL DEFAULT '{}',
  ADD COLUMN IF NOT EXISTS last_updated_at TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS cached_at TIMESTAMPTZ NOT NULL DEFAULT NOW();

ALTER TABLE IF EXISTS appfolio_unit_turn_details
  ADD COLUMN IF NOT EXISTS property_id TEXT,
  ADD COLUMN IF NOT EXISTS property_name TEXT,
  ADD COLUMN IF NOT EXISTS unit_id TEXT,
  ADD COLUMN IF NOT EXISTS unit_name TEXT,
  ADD COLUMN IF NOT EXISTS notes TEXT,
  ADD COLUMN IF NOT EXISTS reference_user TEXT,
  ADD COLUMN IF NOT EXISTS move_out_date TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS turn_end_date TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS expected_move_in_date TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS target_days_to_complete INTEGER,
  ADD COLUMN IF NOT EXISTS total_days_to_complete INTEGER,
  ADD COLUMN IF NOT EXISTS labor_from_work_orders TEXT,
  ADD COLUMN IF NOT EXISTS purchase_orders_from_work_orders TEXT,
  ADD COLUMN IF NOT EXISTS billables_from_work_orders TEXT,
  ADD COLUMN IF NOT EXISTS inventory_from_work_orders TEXT,
  ADD COLUMN IF NOT EXISTS total_billed TEXT,
  ADD COLUMN IF NOT EXISTS unit_turn_status TEXT,
  ADD COLUMN IF NOT EXISTS property_visibility TEXT,
  ADD COLUMN IF NOT EXISTS raw_json JSONB NOT NULL DEFAULT '{}',
  ADD COLUMN IF NOT EXISTS last_updated_at TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS cached_at TIMESTAMPTZ NOT NULL DEFAULT NOW();

ALTER TABLE IF EXISTS appfolio_unit_vacancies
  ADD COLUMN IF NOT EXISTS property_id TEXT,
  ADD COLUMN IF NOT EXISTS property_name TEXT,
  ADD COLUMN IF NOT EXISTS unit_id TEXT,
  ADD COLUMN IF NOT EXISTS unit_name TEXT,
  ADD COLUMN IF NOT EXISTS vacant_from TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS available_on TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS market_rent TEXT,
  ADD COLUMN IF NOT EXISTS bedrooms TEXT,
  ADD COLUMN IF NOT EXISTS bathrooms TEXT,
  ADD COLUMN IF NOT EXISTS days_vacant INTEGER,
  ADD COLUMN IF NOT EXISTS status TEXT,
  ADD COLUMN IF NOT EXISTS property_visibility TEXT,
  ADD COLUMN IF NOT EXISTS raw_json JSONB NOT NULL DEFAULT '{}',
  ADD COLUMN IF NOT EXISTS last_updated_at TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS cached_at TIMESTAMPTZ NOT NULL DEFAULT NOW();

CREATE INDEX IF NOT EXISTS sync_job_runs_endpoint_status_idx
  ON sync_job_runs(endpoint_key, status);
CREATE INDEX IF NOT EXISTS appfolio_unit_inspections_property_idx
  ON appfolio_unit_inspections(property_id);
CREATE INDEX IF NOT EXISTS appfolio_unit_inspections_unit_idx
  ON appfolio_unit_inspections(unit_id);
CREATE INDEX IF NOT EXISTS appfolio_tenant_directory_property_idx
  ON appfolio_tenant_directory(property_id);
CREATE INDEX IF NOT EXISTS appfolio_tenant_directory_unit_idx
  ON appfolio_tenant_directory(unit_id);
CREATE INDEX IF NOT EXISTS appfolio_unit_turn_details_property_idx
  ON appfolio_unit_turn_details(property_id);
CREATE INDEX IF NOT EXISTS appfolio_unit_turn_details_unit_idx
  ON appfolio_unit_turn_details(unit_id);
CREATE INDEX IF NOT EXISTS appfolio_unit_vacancies_property_idx
  ON appfolio_unit_vacancies(property_id);
CREATE INDEX IF NOT EXISTS appfolio_unit_vacancies_unit_idx
  ON appfolio_unit_vacancies(unit_id);