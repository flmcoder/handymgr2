CREATE TABLE IF NOT EXISTS app_debug_events (
  id BIGSERIAL PRIMARY KEY,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  client_ts TIMESTAMPTZ,
  session_id TEXT,
  user_name TEXT,
  login_email TEXT,
  role TEXT,
  property_group_id TEXT,
  level TEXT NOT NULL DEFAULT 'info',
  event_type TEXT NOT NULL DEFAULT 'runtime',
  source TEXT NOT NULL DEFAULT 'frontend',
  action TEXT,
  message TEXT NOT NULL,
  details JSONB NOT NULL DEFAULT '{}'::jsonb
);

CREATE INDEX IF NOT EXISTS app_debug_events_created_idx
  ON app_debug_events (created_at DESC);

CREATE INDEX IF NOT EXISTS app_debug_events_session_idx
  ON app_debug_events (session_id, created_at DESC);
