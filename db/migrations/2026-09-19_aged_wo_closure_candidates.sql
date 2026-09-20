CREATE TABLE IF NOT EXISTS aged_wo_closure_candidates (
  id TEXT PRIMARY KEY,
  work_order_id TEXT NOT NULL,
  wo_number TEXT,
  work_order_uuid TEXT,
  bill_id TEXT NOT NULL,
  bill_number TEXT,
  vendor_id TEXT,
  vendor_name TEXT,
  property_id TEXT,
  property_name TEXT,
  unit_id TEXT,
  wo_status TEXT,
  wo_total_cost REAL,
  bill_total_amount REAL,
  amount_delta REAL,
  match_score REAL,
  match_reason TEXT,
  status TEXT NOT NULL DEFAULT 'pending_review',
  pipeline_run_id TEXT,
  reviewed_by TEXT,
  reviewed_at TIMESTAMPTZ,
  review_notes TEXT,
  closure_result JSONB,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS aged_wo_closure_candidates_status_idx
  ON aged_wo_closure_candidates(status);

CREATE INDEX IF NOT EXISTS aged_wo_closure_candidates_wo_idx
  ON aged_wo_closure_candidates(work_order_id);

CREATE INDEX IF NOT EXISTS aged_wo_closure_candidates_bill_idx
  ON aged_wo_closure_candidates(bill_id);

CREATE INDEX IF NOT EXISTS aged_wo_closure_candidates_vendor_idx
  ON aged_wo_closure_candidates(vendor_id);

CREATE INDEX IF NOT EXISTS aged_wo_closure_candidates_run_idx
  ON aged_wo_closure_candidates(pipeline_run_id);

CREATE INDEX IF NOT EXISTS aged_wo_closure_candidates_property_idx
  ON aged_wo_closure_candidates(property_id);
