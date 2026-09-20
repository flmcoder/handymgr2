-- Add missing v0 API fields to work orders and bills tables
-- Migration date: 2026-01-16

-- Work Orders: add 10 new columns for v0 API fields
ALTER TABLE appfolio_work_orders
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

-- Bills: add 4 new columns for v0 API fields
ALTER TABLE appfolio_bills
  ADD COLUMN IF NOT EXISTS work_order_id TEXT,
  ADD COLUMN IF NOT EXISTS cash_account_id TEXT,
  ADD COLUMN IF NOT EXISTS posting_date TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS account_number TEXT;

-- Add indexes for commonly queried new fields
CREATE INDEX IF NOT EXISTS idx_work_orders_wo_type ON appfolio_work_orders(wo_type);
CREATE INDEX IF NOT EXISTS idx_work_orders_vendor_trade ON appfolio_work_orders(vendor_trade);
CREATE INDEX IF NOT EXISTS idx_bills_work_order_id ON appfolio_bills(work_order_id);
