-- Repair appfolio_work_orders schema drift (Postgres error 42703:
-- column "vendor_trade" of relation "appfolio_work_orders" does not exist).
-- 2026-01-16_add_missing_v0_fields.sql was never applied on some databases, so
-- the v0:work_orders upsert failed with 0 rows written. Every statement is
-- additive and idempotent; safe to re-run.

ALTER TABLE IF EXISTS appfolio_work_orders
  ADD COLUMN IF NOT EXISTS work_order_uuid TEXT,
  ADD COLUMN IF NOT EXISTS vendor_trade TEXT,
  ADD COLUMN IF NOT EXISTS occupancy_id TEXT,
  ADD COLUMN IF NOT EXISTS completed_on TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS work_completed_on TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS canceled_on TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS scheduled_start TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS scheduled_end TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS wo_type TEXT,
  ADD COLUMN IF NOT EXISTS work_order_issue TEXT,
  ADD COLUMN IF NOT EXISTS requesting_tenant_id TEXT;

CREATE INDEX IF NOT EXISTS appfolio_work_orders_vendor_trade_idx ON appfolio_work_orders(vendor_trade);
CREATE INDEX IF NOT EXISTS appfolio_work_orders_wo_type_idx ON appfolio_work_orders(wo_type);
