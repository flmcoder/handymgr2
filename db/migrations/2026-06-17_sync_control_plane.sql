-- Migration: 2026-06-17_sync_control_plane
-- Creates the AppFolio sync control plane tables:
--   sync_job_runs, sync_job_cursors, appfolio_raw_responses,
--   appfolio_request_log, appfolio_rate_counters, appfolio_patch_queue
-- Also ensures the core AppFolio domain tables exist for the migration scope.
-- Safe to re-run (all statements use IF NOT EXISTS / ON CONFLICT DO NOTHING).

-- ── Core domain tables ───────────────────────────────────────────────────────

CREATE TABLE IF NOT EXISTS appfolio_properties (
  id                TEXT PRIMARY KEY,
  name              TEXT NOT NULL,
  property_group_id TEXT,
  street            TEXT,
  city              TEXT,
  state             TEXT,
  zip               TEXT,
  raw_json          JSONB NOT NULL DEFAULT '{}',
  last_updated_at   TIMESTAMPTZ,
  cached_at         TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS appfolio_properties_group_idx ON appfolio_properties(property_group_id);
CREATE INDEX IF NOT EXISTS appfolio_properties_name_idx  ON appfolio_properties(name);

CREATE TABLE IF NOT EXISTS appfolio_units (
  unit_id         TEXT PRIMARY KEY,
  property_id     TEXT NOT NULL,
  name            TEXT,
  unit_number     TEXT,
  status          TEXT,
  bedrooms        INTEGER,
  bathrooms       REAL,
  square_feet     INTEGER,
  market_rent     REAL,
  raw_json        JSONB NOT NULL DEFAULT '{}',
  last_updated_at TIMESTAMPTZ,
  cached_at       TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS appfolio_units_property_idx ON appfolio_units(property_id);
CREATE INDEX IF NOT EXISTS appfolio_units_status_idx   ON appfolio_units(status);

CREATE TABLE IF NOT EXISTS appfolio_work_orders (
  id                   TEXT PRIMARY KEY,
  wo_number            TEXT,
  property_id          TEXT,
  unit_id              TEXT,
  property_group_id    TEXT,
  description          TEXT,
  category             TEXT,
  priority             TEXT,
  status               TEXT,
  assigned_user_id     TEXT,
  assigned_user_name   TEXT,
  vendor_id            TEXT,
  vendor_name          TEXT,
  estimated_amount     REAL,
  total_cost           REAL,
  created_at           TIMESTAMPTZ,
  updated_at           TIMESTAMPTZ,
  raw_json             JSONB NOT NULL DEFAULT '{}'
);
CREATE INDEX IF NOT EXISTS appfolio_work_orders_number_idx   ON appfolio_work_orders(wo_number);
CREATE INDEX IF NOT EXISTS appfolio_work_orders_status_idx   ON appfolio_work_orders(status);
CREATE INDEX IF NOT EXISTS appfolio_work_orders_property_idx ON appfolio_work_orders(property_id);
CREATE INDEX IF NOT EXISTS appfolio_work_orders_unit_idx     ON appfolio_work_orders(unit_id);

CREATE TABLE IF NOT EXISTS appfolio_estimates (
  estimate_id      TEXT PRIMARY KEY,
  work_order_id    TEXT,
  work_order_number TEXT,
  current_status   TEXT NOT NULL,
  property_group_id TEXT,
  source           TEXT,
  status_history   JSONB NOT NULL DEFAULT '[]',
  raw_data         JSONB NOT NULL DEFAULT '{}',
  created_at       TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at       TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS appfolio_estimates_status_idx ON appfolio_estimates(current_status);
CREATE INDEX IF NOT EXISTS appfolio_estimates_group_idx  ON appfolio_estimates(property_group_id);
CREATE INDEX IF NOT EXISTS appfolio_estimates_wo_idx     ON appfolio_estimates(work_order_id);

CREATE TABLE IF NOT EXISTS unit_turn_tracker (
  tracking_uuid    TEXT PRIMARY KEY,
  tracking_code    TEXT,
  turn_key         TEXT NOT NULL,
  unit_turn_id     TEXT,
  unit_id          TEXT,
  property_id      TEXT,
  unit_name        TEXT,
  property_name    TEXT,
  status           TEXT NOT NULL DEFAULT 'open',
  confidence_score REAL,
  confidence_label TEXT,
  source_flags     JSONB NOT NULL DEFAULT '{}',
  metadata         JSONB NOT NULL DEFAULT '{}',
  closed_at        TIMESTAMPTZ,
  created_at       TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at       TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE UNIQUE INDEX IF NOT EXISTS unit_turn_tracker_turn_key_unique ON unit_turn_tracker(turn_key);
CREATE INDEX IF NOT EXISTS unit_turn_tracker_status_idx   ON unit_turn_tracker(status);
CREATE INDEX IF NOT EXISTS unit_turn_tracker_unit_idx     ON unit_turn_tracker(unit_id);
CREATE INDEX IF NOT EXISTS unit_turn_tracker_property_idx ON unit_turn_tracker(property_id);

CREATE TABLE IF NOT EXISTS unit_turn_milestones (
  id             SERIAL PRIMARY KEY,
  tracking_uuid  TEXT NOT NULL,
  milestone_key  TEXT NOT NULL,
  milestone_date TIMESTAMPTZ,
  source         TEXT,
  notes          TEXT,
  created_at     TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS unit_turn_milestones_tracking_idx ON unit_turn_milestones(tracking_uuid);

CREATE TABLE IF NOT EXISTS unit_turn_work_orders (
  id             SERIAL PRIMARY KEY,
  tracking_uuid  TEXT NOT NULL,
  wo_id          TEXT NOT NULL,
  wo_db_uuid     TEXT,
  source         TEXT DEFAULT 'manual',
  status         TEXT,
  removed        BOOLEAN DEFAULT FALSE,
  created_at     TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS unit_turn_work_orders_tracking_idx ON unit_turn_work_orders(tracking_uuid);
CREATE INDEX IF NOT EXISTS unit_turn_work_orders_wo_idx       ON unit_turn_work_orders(wo_id);

-- ── Sync Control Plane ───────────────────────────────────────────────────────

CREATE TABLE IF NOT EXISTS sync_job_runs (
  id                     SERIAL PRIMARY KEY,
  run_id                 TEXT NOT NULL,
  endpoint_key           TEXT NOT NULL,
  api_version            TEXT NOT NULL,
  trigger_type           TEXT NOT NULL,
  status                 TEXT NOT NULL DEFAULT 'running',
  filters_fingerprint    TEXT,
  pages_completed        INTEGER NOT NULL DEFAULT 0,
  rows_upserted          INTEGER NOT NULL DEFAULT 0,
  rows_skipped           INTEGER NOT NULL DEFAULT 0,
  last_error             TEXT,
  execution_start_cursor TEXT,
  started_at             TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  completed_at           TIMESTAMPTZ
);
CREATE INDEX IF NOT EXISTS sync_job_runs_run_id_idx          ON sync_job_runs(run_id);
CREATE INDEX IF NOT EXISTS sync_job_runs_endpoint_status_idx ON sync_job_runs(endpoint_key, status);

CREATE TABLE IF NOT EXISTS sync_job_cursors (
  id                SERIAL PRIMARY KEY,
  run_id            TEXT NOT NULL,
  endpoint_key      TEXT NOT NULL,
  page_index        INTEGER NOT NULL DEFAULT 0,
  cursor_in         TEXT,
  cursor_out        TEXT,
  cursor_expires_at TIMESTAMPTZ,
  fetched_at        TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  record_count      INTEGER DEFAULT 0,
  retries_used      INTEGER DEFAULT 0,
  is_terminal       BOOLEAN DEFAULT FALSE
);
CREATE INDEX IF NOT EXISTS sync_job_cursors_run_page_idx ON sync_job_cursors(run_id, page_index);
CREATE INDEX IF NOT EXISTS sync_job_cursors_endpoint_idx ON sync_job_cursors(endpoint_key);

CREATE TABLE IF NOT EXISTS appfolio_raw_responses (
  id            SERIAL PRIMARY KEY,
  run_id        TEXT NOT NULL,
  endpoint_key  TEXT NOT NULL,
  page_index    INTEGER NOT NULL DEFAULT 0,
  cursor_in     TEXT,
  cursor_out    TEXT,
  status_code   INTEGER,
  record_count  INTEGER DEFAULT 0,
  response_json JSONB NOT NULL DEFAULT '{}',
  fetched_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS appfolio_raw_responses_run_idx       ON appfolio_raw_responses(run_id);
CREATE INDEX IF NOT EXISTS appfolio_raw_responses_endpoint_idx  ON appfolio_raw_responses(endpoint_key);
CREATE INDEX IF NOT EXISTS appfolio_raw_responses_fetched_at_idx ON appfolio_raw_responses(fetched_at);

CREATE TABLE IF NOT EXISTS appfolio_request_log (
  id                   SERIAL PRIMARY KEY,
  run_id               TEXT,
  endpoint_key         TEXT NOT NULL,
  api_version          TEXT,
  method               TEXT DEFAULT 'GET',
  status_code          INTEGER,
  latency_ms           INTEGER,
  retry_after_seconds  INTEGER,
  attempt_number       INTEGER DEFAULT 1,
  error_text           TEXT,
  cursor_snapshot      TEXT,
  requested_at         TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS appfolio_request_log_endpoint_time_idx ON appfolio_request_log(endpoint_key, requested_at);
CREATE INDEX IF NOT EXISTS appfolio_request_log_run_idx           ON appfolio_request_log(run_id);

CREATE TABLE IF NOT EXISTS appfolio_rate_counters (
  api_version    TEXT NOT NULL,
  endpoint_key   TEXT NOT NULL,
  window_type    TEXT NOT NULL,
  window_start   TEXT NOT NULL,
  request_count  INTEGER NOT NULL DEFAULT 0,
  status_429_count INTEGER NOT NULL DEFAULT 0,
  PRIMARY KEY (api_version, endpoint_key, window_type, window_start)
);
CREATE INDEX IF NOT EXISTS appfolio_rate_counters_window_idx ON appfolio_rate_counters(window_type, window_start);

CREATE TABLE IF NOT EXISTS appfolio_patch_queue (
  id            SERIAL PRIMARY KEY,
  resource_id   TEXT NOT NULL,
  resource_type TEXT NOT NULL,
  method        TEXT DEFAULT 'PATCH',
  endpoint_path TEXT NOT NULL,
  payload_json  JSONB NOT NULL DEFAULT '{}',
  status        TEXT NOT NULL DEFAULT 'pending',
  priority      INTEGER DEFAULT 100,
  attempt_count INTEGER DEFAULT 0,
  last_error    TEXT,
  locked_at     TIMESTAMPTZ,
  lock_owner    TEXT,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS appfolio_patch_queue_resource_pending_idx ON appfolio_patch_queue(resource_id, status);
CREATE INDEX IF NOT EXISTS appfolio_patch_queue_status_priority_idx  ON appfolio_patch_queue(status, priority, id);

-- Seed safe rate-limit defaults (idempotent).
CREATE TABLE IF NOT EXISTS appfolio_rate_limits (
  key               TEXT PRIMARY KEY,
  api_version       TEXT NOT NULL,
  endpoint_key      TEXT NOT NULL,
  max_per_second    INTEGER NOT NULL DEFAULT 8,
  max_per_minute    INTEGER NOT NULL DEFAULT 256,
  max_per_hour      INTEGER NOT NULL DEFAULT 4096,
  cooldown_until    TIMESTAMPTZ,
  updated_at        TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
INSERT INTO appfolio_rate_limits (key, api_version, endpoint_key, max_per_second, max_per_minute, max_per_hour)
VALUES
  ('v0:global', 'v0', 'global', 8, 256, 4096),
  ('v2:global', 'v2', 'global', 8, 256, 4096)
ON CONFLICT (key) DO UPDATE SET
  max_per_second = EXCLUDED.max_per_second,
  max_per_minute = EXCLUDED.max_per_minute,
  max_per_hour   = EXCLUDED.max_per_hour,
  updated_at     = NOW();
