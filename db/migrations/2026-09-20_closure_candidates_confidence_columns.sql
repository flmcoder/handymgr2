-- Phase 1: Multi-variable confidence matching columns
-- Adds date windows, property address, match flags, and confidence scoring
-- to aged_wo_closure_candidates for the upgraded WO Closure Assistant.

ALTER TABLE aged_wo_closure_candidates
  ADD COLUMN IF NOT EXISTS wo_created_at TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS wo_completed_on TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS bill_invoice_date TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS bill_paid_at TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS bill_service_from TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS bill_service_to TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS property_address TEXT,
  ADD COLUMN IF NOT EXISTS match_flags JSONB NOT NULL DEFAULT '{}'::jsonb,
  ADD COLUMN IF NOT EXISTS confidence_score INTEGER,
  ADD COLUMN IF NOT EXISTS confidence_tier TEXT;

ALTER TABLE aged_wo_closure_candidates
  ADD CONSTRAINT chk_confidence_tier
  CHECK (confidence_tier IN ('very_high', 'high', 'medium', 'low', 'none'));

CREATE INDEX IF NOT EXISTS aged_wo_closure_candidates_confidence_tier_idx
  ON aged_wo_closure_candidates (confidence_tier);

CREATE INDEX IF NOT EXISTS aged_wo_closure_candidates_confidence_score_idx
  ON aged_wo_closure_candidates (confidence_score DESC NULLS LAST);
