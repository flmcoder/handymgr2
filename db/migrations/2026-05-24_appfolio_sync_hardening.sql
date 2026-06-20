PRAGMA foreign_keys = ON;

-- Global and per-endpoint controls for AppFolio requests.
CREATE TABLE IF NOT EXISTS appfolio_rate_limits (
  key TEXT PRIMARY KEY,
  api_version TEXT NOT NULL,
  endpoint_key TEXT NOT NULL,
  max_per_second INTEGER NOT NULL DEFAULT 8,
  max_per_minute INTEGER NOT NULL DEFAULT 250,
  max_per_hour INTEGER NOT NULL DEFAULT 4000,
  cooldown_until TEXT,
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_appfolio_rate_limits_version_endpoint
  ON appfolio_rate_limits(api_version, endpoint_key);

-- Outbound queue so UI never calls AppFolio directly.
CREATE TABLE IF NOT EXISTS appfolio_sync_queue (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  api_version TEXT NOT NULL,
  endpoint_key TEXT NOT NULL,
  request_key TEXT NOT NULL,
  method TEXT NOT NULL DEFAULT 'GET',
  payload_json TEXT,
  priority INTEGER NOT NULL DEFAULT 100,
  status TEXT NOT NULL DEFAULT 'queued',
  attempt_count INTEGER NOT NULL DEFAULT 0,
  next_attempt_at TEXT NOT NULL DEFAULT (datetime('now')),
  locked_at TEXT,
  lock_owner TEXT,
  last_error TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE UNIQUE INDEX IF NOT EXISTS uq_appfolio_sync_queue_request_key
  ON appfolio_sync_queue(request_key);

CREATE INDEX IF NOT EXISTS idx_appfolio_sync_queue_claim
  ON appfolio_sync_queue(status, next_attempt_at, priority, id);

CREATE INDEX IF NOT EXISTS idx_appfolio_sync_queue_endpoint
  ON appfolio_sync_queue(api_version, endpoint_key, status, next_attempt_at);

-- Rolling per-second/per-minute/per-hour counters for throttle checks.
CREATE TABLE IF NOT EXISTS appfolio_rate_counters (
  api_version TEXT NOT NULL,
  endpoint_key TEXT NOT NULL,
  window_type TEXT NOT NULL,
  window_start TEXT NOT NULL,
  request_count INTEGER NOT NULL DEFAULT 0,
  status_429_count INTEGER NOT NULL DEFAULT 0,
  PRIMARY KEY (api_version, endpoint_key, window_type, window_start)
);

CREATE INDEX IF NOT EXISTS idx_appfolio_rate_counters_window
  ON appfolio_rate_counters(window_type, window_start);

-- Cache keyed by endpoint + request fingerprint, safe for stale-while-revalidate.
CREATE TABLE IF NOT EXISTS appfolio_response_cache (
  cache_key TEXT PRIMARY KEY,
  api_version TEXT NOT NULL,
  endpoint_key TEXT NOT NULL,
  request_fingerprint TEXT NOT NULL,
  response_json TEXT NOT NULL,
  etag TEXT,
  status_code INTEGER,
  fetched_at TEXT NOT NULL DEFAULT (datetime('now')),
  expires_at TEXT,
  stale_until TEXT,
  source TEXT NOT NULL DEFAULT 'sync_worker'
);

CREATE UNIQUE INDEX IF NOT EXISTS uq_appfolio_response_cache_lookup
  ON appfolio_response_cache(api_version, endpoint_key, request_fingerprint);

CREATE INDEX IF NOT EXISTS idx_appfolio_response_cache_expiry
  ON appfolio_response_cache(expires_at, stale_until);

-- Bridge from v2 rows (number-based) to v0 UUID identities.
CREATE TABLE IF NOT EXISTS appfolio_work_order_bridge (
  work_order_uuid TEXT PRIMARY KEY,
  work_order_number TEXT NOT NULL,
  property_id TEXT,
  property_uuid TEXT,
  status TEXT,
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE UNIQUE INDEX IF NOT EXISTS uq_appfolio_work_order_bridge_number
  ON appfolio_work_order_bridge(work_order_number);

CREATE INDEX IF NOT EXISTS idx_appfolio_work_order_bridge_property
  ON appfolio_work_order_bridge(property_id, property_uuid);

-- Request log for observability and 429 forensics.
CREATE TABLE IF NOT EXISTS appfolio_request_log (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  api_version TEXT NOT NULL,
  endpoint_key TEXT NOT NULL,
  request_key TEXT,
  status_code INTEGER,
  latency_ms INTEGER,
  retry_after_seconds INTEGER,
  error_text TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_appfolio_request_log_time
  ON appfolio_request_log(created_at);

CREATE INDEX IF NOT EXISTS idx_appfolio_request_log_endpoint
  ON appfolio_request_log(api_version, endpoint_key, created_at);

-- Sync cursors for resumable pagination and timeout-safe jobs.
CREATE TABLE IF NOT EXISTS appfolio_sync_state (
  key TEXT PRIMARY KEY,
  api_version TEXT NOT NULL,
  endpoint_key TEXT NOT NULL,
  next_page_url TEXT,
  last_cursor TEXT,
  last_success_at TEXT,
  last_error TEXT,
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE UNIQUE INDEX IF NOT EXISTS uq_appfolio_sync_state_endpoint
  ON appfolio_sync_state(api_version, endpoint_key);

-- Seed baseline limits (safe defaults below hard caps).
INSERT INTO appfolio_rate_limits (key, api_version, endpoint_key, max_per_second, max_per_minute, max_per_hour)
VALUES
  ('v0:global', 'v0', 'global', 8, 250, 4000),
  ('v2:global', 'v2', 'global', 8, 250, 4000)
ON CONFLICT(key) DO UPDATE SET
  max_per_second = excluded.max_per_second,
  max_per_minute = excluded.max_per_minute,
  max_per_hour = excluded.max_per_hour,
  updated_at = (datetime('now'));

-- Optional helper view for dashboarding throttle pressure.
CREATE VIEW IF NOT EXISTS appfolio_rate_health AS
SELECT
  api_version,
  endpoint_key,
  SUM(CASE WHEN window_type = 'second' THEN request_count ELSE 0 END) AS reqs_second_windows,
  SUM(CASE WHEN window_type = 'minute' THEN request_count ELSE 0 END) AS reqs_minute_windows,
  SUM(CASE WHEN window_type = 'hour' THEN request_count ELSE 0 END) AS reqs_hour_windows,
  SUM(status_429_count) AS total_429s
FROM appfolio_rate_counters
GROUP BY api_version, endpoint_key;