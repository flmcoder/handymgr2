-- Migration: 2026-07-12_appfolio_webhook_ingest
-- Adds webhook event_id/raw_payload columns and keeps legacy columns backfilled.

CREATE TABLE IF NOT EXISTS webhook_events (
  id                BIGSERIAL PRIMARY KEY,
  event_id          TEXT,
  topic             TEXT NOT NULL,
  event_type        TEXT,
  resource_type     TEXT,
  resource_id       TEXT,
  signature         TEXT,
  raw_payload       JSONB NOT NULL DEFAULT '{}'::jsonb,
  received_at       TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  processed_at      TIMESTAMPTZ,
  processing_status TEXT DEFAULT 'pending',
  event_uuid        TEXT,
  payload_json      JSONB
);

ALTER TABLE webhook_events ADD COLUMN IF NOT EXISTS event_id TEXT;
ALTER TABLE webhook_events ADD COLUMN IF NOT EXISTS raw_payload JSONB NOT NULL DEFAULT '{}'::jsonb;
ALTER TABLE webhook_events ADD COLUMN IF NOT EXISTS event_uuid TEXT;
ALTER TABLE webhook_events ADD COLUMN IF NOT EXISTS payload_json JSONB;
ALTER TABLE webhook_events ADD COLUMN IF NOT EXISTS processing_status TEXT DEFAULT 'pending';

UPDATE webhook_events
SET event_id = COALESCE(event_id, event_uuid)
WHERE event_id IS NULL AND event_uuid IS NOT NULL;

UPDATE webhook_events
SET raw_payload = COALESCE(raw_payload, payload_json, '{}'::jsonb)
WHERE raw_payload IS NULL OR raw_payload = '{}'::jsonb;

CREATE INDEX IF NOT EXISTS webhook_events_topic_idx
  ON webhook_events(topic);

CREATE INDEX IF NOT EXISTS webhook_events_resource_idx
  ON webhook_events(resource_type, resource_id);

CREATE INDEX IF NOT EXISTS webhook_events_status_idx
  ON webhook_events(processing_status);

CREATE INDEX IF NOT EXISTS webhook_events_event_id_idx
  ON webhook_events(event_id);

CREATE UNIQUE INDEX IF NOT EXISTS webhook_events_event_id_unique
  ON webhook_events(event_id)
  WHERE event_id IS NOT NULL AND event_id <> '';