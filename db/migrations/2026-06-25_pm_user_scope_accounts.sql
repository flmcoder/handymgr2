-- Migration: 2026-06-25_pm_user_scope_accounts
-- Purpose: normalize PM OTP login scope mapping so one PM identity can be authorized for
-- multiple property-group scopes while preserving legacy pm_proxy_users rows.

CREATE TABLE IF NOT EXISTS pm_proxy_users (
  user_uuid TEXT PRIMARY KEY,
  email TEXT NOT NULL,
  full_name TEXT,
  phone TEXT,
  property_group_uuid TEXT,
  roles JSONB NOT NULL DEFAULT '[]'::jsonb,
  active BOOLEAN DEFAULT TRUE,
  raw_json JSONB NOT NULL DEFAULT '{}'::jsonb,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- Legacy schema enforced unique email across all scopes. Drop it so the same PM identifier
-- can be represented across distinct scoped accounts.
DROP INDEX IF EXISTS pm_proxy_users_email_unique;

CREATE INDEX IF NOT EXISTS pm_proxy_users_email_idx ON pm_proxy_users(lower(email));
CREATE INDEX IF NOT EXISTS pm_proxy_users_group_idx ON pm_proxy_users(property_group_uuid);

CREATE TABLE IF NOT EXISTS pm_proxy_user_scopes (
  id INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  user_uuid TEXT NOT NULL,
  property_group_uuid TEXT NOT NULL,
  is_primary BOOLEAN DEFAULT FALSE,
  active BOOLEAN DEFAULT TRUE,
  source TEXT DEFAULT 'legacy',
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  UNIQUE (user_uuid, property_group_uuid)
);

CREATE INDEX IF NOT EXISTS pm_proxy_user_scopes_user_idx
  ON pm_proxy_user_scopes(user_uuid);

CREATE INDEX IF NOT EXISTS pm_proxy_user_scopes_group_idx
  ON pm_proxy_user_scopes(property_group_uuid);

-- Backfill scopes from legacy single-scope column.
INSERT INTO pm_proxy_user_scopes (user_uuid, property_group_uuid, is_primary, active, source)
SELECT user_uuid, property_group_uuid, TRUE, COALESCE(active, TRUE), 'legacy_backfill'
FROM pm_proxy_users
WHERE COALESCE(property_group_uuid, '') <> ''
ON CONFLICT (user_uuid, property_group_uuid) DO UPDATE
SET
  active = EXCLUDED.active,
  updated_at = NOW();
