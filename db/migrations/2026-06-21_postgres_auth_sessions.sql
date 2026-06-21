CREATE TABLE IF NOT EXISTS trusted_devices (
  device_token TEXT PRIMARY KEY,
  user_name TEXT,
  role TEXT DEFAULT 'full',
  login_email TEXT,
  property_group_uuid TEXT,
  phone TEXT,
  revoked BOOLEAN DEFAULT FALSE,
  last_seen_at TIMESTAMPTZ,
  expires_at TIMESTAMPTZ,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS device_otps (
  id INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  email TEXT NOT NULL,
  code TEXT NOT NULL,
  used BOOLEAN DEFAULT FALSE,
  expires_at TIMESTAMPTZ NOT NULL,
  user_name TEXT,
  role_hint TEXT,
  property_group_uuid TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  used_at TIMESTAMPTZ
);

CREATE INDEX IF NOT EXISTS device_otps_email_idx ON device_otps(email);

CREATE TABLE IF NOT EXISTS pm_proxy_users (
  user_uuid TEXT PRIMARY KEY,
  email TEXT NOT NULL,
  full_name TEXT,
  phone TEXT,
  property_group_uuid TEXT NOT NULL,
  roles JSONB NOT NULL DEFAULT '[]'::jsonb,
  active BOOLEAN DEFAULT TRUE,
  raw_json JSONB NOT NULL DEFAULT '{}'::jsonb,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE UNIQUE INDEX IF NOT EXISTS pm_proxy_users_email_unique ON pm_proxy_users(email);
CREATE INDEX IF NOT EXISTS pm_proxy_users_group_idx ON pm_proxy_users(property_group_uuid);

CREATE TABLE IF NOT EXISTS proxy_config (
  key TEXT PRIMARY KEY,
  value TEXT,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);