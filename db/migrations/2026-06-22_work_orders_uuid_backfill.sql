-- Migration: 2026-06-22_work_orders_uuid_backfill
-- Adds explicit UUID column for AppFolio v0 work orders and backfills from known fields.
-- Safe to re-run.

ALTER TABLE appfolio_work_orders
  ADD COLUMN IF NOT EXISTS work_order_uuid TEXT;

CREATE INDEX IF NOT EXISTS appfolio_work_orders_uuid_idx
  ON appfolio_work_orders(work_order_uuid);

-- Backfill from raw JSON UUID-ish fields where possible.
UPDATE appfolio_work_orders
SET work_order_uuid = COALESCE(
  NULLIF(raw_json->>'work_order_uuid', ''),
  NULLIF(raw_json->>'v0_uuid', ''),
  NULLIF(raw_json->>'UUID', ''),
  NULLIF(raw_json->>'uuid', ''),
  CASE
    WHEN id ~* '^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$' THEN id
    ELSE NULL
  END
)
WHERE COALESCE(work_order_uuid, '') = '';
