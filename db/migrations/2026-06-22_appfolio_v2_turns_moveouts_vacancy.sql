-- Migration: 2026-06-22_appfolio_v2_turns_moveouts_vacancy
-- Adds normalized v2 report fact tables used by local Postgres read endpoints.
-- Safe to re-run.

CREATE TABLE IF NOT EXISTS appfolio_unit_inspections (
  inspection_id                 TEXT PRIMARY KEY,
  property_id                   TEXT,
  property_name                 TEXT,
  unit_id                       TEXT,
  unit_name                     TEXT,
  last_inspection_date          TIMESTAMPTZ,
  tenant_name                   TEXT,
  tenant_primary_phone_number   TEXT,
  move_in_date                  TIMESTAMPTZ,
  move_out_date                 TIMESTAMPTZ,
  rentable                      TEXT,
  occupancy_id                  TEXT,
  unit_tags                     TEXT,
  raw_json                      JSONB NOT NULL DEFAULT '{}',
  last_updated_at               TIMESTAMPTZ,
  cached_at                     TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS appfolio_unit_inspections_property_idx ON appfolio_unit_inspections(property_id);
CREATE INDEX IF NOT EXISTS appfolio_unit_inspections_unit_idx ON appfolio_unit_inspections(unit_id);
CREATE INDEX IF NOT EXISTS appfolio_unit_inspections_inspection_idx ON appfolio_unit_inspections(last_inspection_date);

CREATE TABLE IF NOT EXISTS appfolio_tenant_directory (
  record_id           TEXT PRIMARY KEY,
  property_id         TEXT,
  property_name       TEXT,
  unit_id             TEXT,
  unit_name           TEXT,
  tenant_name         TEXT,
  status              TEXT,
  tenant_type         TEXT,
  phone_numbers       TEXT,
  emails              TEXT,
  move_in_date        TIMESTAMPTZ,
  lease_to            TIMESTAMPTZ,
  move_out_date       TIMESTAMPTZ,
  rent                TEXT,
  tenant_tags         TEXT,
  tenant_agent        TEXT,
  tenant_visibility   TEXT,
  occupancy_id        TEXT,
  unit_tags           TEXT,
  raw_json            JSONB NOT NULL DEFAULT '{}',
  last_updated_at     TIMESTAMPTZ,
  cached_at           TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS appfolio_tenant_directory_property_idx ON appfolio_tenant_directory(property_id);
CREATE INDEX IF NOT EXISTS appfolio_tenant_directory_unit_idx ON appfolio_tenant_directory(unit_id);
CREATE INDEX IF NOT EXISTS appfolio_tenant_directory_move_out_idx ON appfolio_tenant_directory(move_out_date);

CREATE TABLE IF NOT EXISTS appfolio_unit_turn_details (
  turn_id                             TEXT PRIMARY KEY,
  property_id                         TEXT,
  property_name                       TEXT,
  unit_id                             TEXT,
  unit_name                           TEXT,
  notes                               TEXT,
  reference_user                      TEXT,
  move_out_date                       TIMESTAMPTZ,
  turn_end_date                       TIMESTAMPTZ,
  expected_move_in_date               TIMESTAMPTZ,
  target_days_to_complete             INTEGER,
  total_days_to_complete              INTEGER,
  labor_from_work_orders              TEXT,
  purchase_orders_from_work_orders    TEXT,
  billables_from_work_orders          TEXT,
  inventory_from_work_orders          TEXT,
  total_billed                        TEXT,
  unit_turn_status                    TEXT,
  property_visibility                 TEXT,
  raw_json                            JSONB NOT NULL DEFAULT '{}',
  last_updated_at                     TIMESTAMPTZ,
  cached_at                           TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS appfolio_unit_turn_details_property_idx ON appfolio_unit_turn_details(property_id);
CREATE INDEX IF NOT EXISTS appfolio_unit_turn_details_unit_idx ON appfolio_unit_turn_details(unit_id);
CREATE INDEX IF NOT EXISTS appfolio_unit_turn_details_move_out_idx ON appfolio_unit_turn_details(move_out_date);

CREATE TABLE IF NOT EXISTS appfolio_unit_vacancies (
  record_id            TEXT PRIMARY KEY,
  property_id          TEXT,
  property_name        TEXT,
  unit_id              TEXT,
  unit_name            TEXT,
  vacant_from          TIMESTAMPTZ,
  available_on         TIMESTAMPTZ,
  market_rent          TEXT,
  bedrooms             TEXT,
  bathrooms            TEXT,
  days_vacant          INTEGER,
  status               TEXT,
  property_visibility  TEXT,
  raw_json             JSONB NOT NULL DEFAULT '{}',
  last_updated_at      TIMESTAMPTZ,
  cached_at            TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS appfolio_unit_vacancies_property_idx ON appfolio_unit_vacancies(property_id);
CREATE INDEX IF NOT EXISTS appfolio_unit_vacancies_unit_idx ON appfolio_unit_vacancies(unit_id);
CREATE INDEX IF NOT EXISTS appfolio_unit_vacancies_vacant_from_idx ON appfolio_unit_vacancies(vacant_from);
