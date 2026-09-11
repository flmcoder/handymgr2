-- Migration: 2026-09-10_work_orders_created_at_backfill
-- Backfills appfolio_work_orders.created_at from the raw AppFolio payload for
-- rows that were synced before the CreatedAt field was mapped into the column.
-- Safe to re-run; only touches rows whose created_at is still NULL. Every raw
-- cast is guarded by a leading-date regex so a malformed value cannot error.

UPDATE appfolio_work_orders
SET created_at = CASE
  WHEN (raw_json->>'CreatedAt') ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}'
    THEN (raw_json->>'CreatedAt')::timestamptz
  WHEN (raw_json->>'created_at') ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}'
    THEN (raw_json->>'created_at')::timestamptz
  WHEN (raw_json->>'CreatedDate') ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}'
    THEN (raw_json->>'CreatedDate')::timestamptz
  WHEN (raw_json->>'created_date') ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}'
    THEN (raw_json->>'created_date')::timestamptz
  ELSE created_at
END
WHERE created_at IS NULL
  AND raw_json IS NOT NULL;