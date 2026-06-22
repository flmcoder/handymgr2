-- Migration: 2026-06-22_appfolio_property_groups_normalized
-- Adds normalized AppFolio property groups table used by local API reads.
-- Safe to re-run.

CREATE TABLE IF NOT EXISTS appfolio_property_groups (
  id TEXT PRIMARY KEY,
  uuid TEXT,
  name TEXT NOT NULL,
  type TEXT,
  property_ids JSONB NOT NULL DEFAULT '[]'::jsonb,
  raw_json JSONB NOT NULL DEFAULT '{}'::jsonb,
  last_updated_at TIMESTAMPTZ,
  cached_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS appfolio_property_groups_uuid_idx
  ON appfolio_property_groups(uuid);

CREATE INDEX IF NOT EXISTS appfolio_property_groups_name_idx
  ON appfolio_property_groups(name);

CREATE INDEX IF NOT EXISTS appfolio_property_groups_updated_idx
  ON appfolio_property_groups(last_updated_at);
