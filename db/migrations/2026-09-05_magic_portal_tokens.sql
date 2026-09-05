CREATE TABLE IF NOT EXISTS magic_tokens (
  token TEXT PRIMARY KEY,
  short_code TEXT NOT NULL UNIQUE,
  wo_id TEXT NOT NULL,
  wo_number TEXT,
  tech_id TEXT NOT NULL,
  tech_name TEXT NOT NULL,
  tech_phone TEXT,
  tenant_name TEXT,
  tenant_phone TEXT,
  property_address TEXT,
  expires_at TIMESTAMPTZ NOT NULL,
  used BOOLEAN NOT NULL DEFAULT FALSE,
  used_at TIMESTAMPTZ,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  metadata JSONB NOT NULL DEFAULT '{}'::jsonb
);

CREATE INDEX IF NOT EXISTS magic_tokens_wo_id_idx
  ON magic_tokens(wo_id);

CREATE INDEX IF NOT EXISTS magic_tokens_expires_at_idx
  ON magic_tokens(expires_at);