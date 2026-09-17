CREATE TABLE IF NOT EXISTS user_last_access (
  user_id TEXT PRIMARY KEY,
  last_accessed_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  session_info JSONB NOT NULL DEFAULT '{}'::jsonb,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS user_last_access_last_accessed_idx
  ON user_last_access (last_accessed_at DESC);
