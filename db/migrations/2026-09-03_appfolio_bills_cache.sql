CREATE TABLE IF NOT EXISTS appfolio_bills (
  id TEXT PRIMARY KEY,
  bill_number TEXT,
  vendor_id TEXT,
  vendor_name TEXT,
  property_id TEXT,
  property_name TEXT,
  unit_id TEXT,
  status TEXT,
  status_label TEXT,
  bill_total_amount REAL,
  invoice_date TIMESTAMPTZ,
  due_date TIMESTAMPTZ,
  paid_at TIMESTAMPTZ,
  work_order_id TEXT,
  cash_account_id TEXT,
  posting_date TIMESTAMPTZ,
  account_number TEXT,
  raw_json JSONB NOT NULL DEFAULT '{}'::jsonb,
  cached_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ
);

ALTER TABLE appfolio_bills
  ADD COLUMN IF NOT EXISTS work_order_id TEXT,
  ADD COLUMN IF NOT EXISTS cash_account_id TEXT,
  ADD COLUMN IF NOT EXISTS posting_date TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS account_number TEXT;

CREATE INDEX IF NOT EXISTS appfolio_bills_vendor_idx ON appfolio_bills(vendor_id);
CREATE INDEX IF NOT EXISTS appfolio_bills_property_idx ON appfolio_bills(property_id);
CREATE INDEX IF NOT EXISTS appfolio_bills_status_idx ON appfolio_bills(status);
CREATE INDEX IF NOT EXISTS appfolio_bills_updated_idx ON appfolio_bills(updated_at DESC);
CREATE INDEX IF NOT EXISTS appfolio_bills_work_order_idx ON appfolio_bills(work_order_id);