CREATE TABLE IF NOT EXISTS appfolio_v2_report_cache (
  cache_key TEXT PRIMARY KEY,
  report_name TEXT NOT NULL,
  request_json JSONB NOT NULL DEFAULT '{}'::jsonb,
  active_generation TEXT,
  status TEXT NOT NULL DEFAULT 'hydrating',
  row_count INTEGER NOT NULL DEFAULT 0,
  hydrated_rows INTEGER NOT NULL DEFAULT 0,
  fetched_at TIMESTAMPTZ,
  expires_at TIMESTAMPTZ,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  last_error TEXT
);

CREATE TABLE IF NOT EXISTS appfolio_v2_report_cache_rows (
  cache_key TEXT NOT NULL REFERENCES appfolio_v2_report_cache(cache_key) ON DELETE CASCADE,
  generation TEXT NOT NULL,
  row_index INTEGER NOT NULL,
  row_json JSONB NOT NULL,
  PRIMARY KEY (cache_key, generation, row_index)
);

CREATE INDEX IF NOT EXISTS appfolio_v2_report_cache_report_idx
  ON appfolio_v2_report_cache(report_name, updated_at DESC);

CREATE INDEX IF NOT EXISTS appfolio_v2_report_cache_rows_generation_idx
  ON appfolio_v2_report_cache_rows(cache_key, generation, row_index);