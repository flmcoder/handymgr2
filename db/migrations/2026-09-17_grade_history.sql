-- Stage 4: Grade History System
CREATE TABLE IF NOT EXISTS tech_grade_history (
  id BIGSERIAL PRIMARY KEY,
  tech_id TEXT NOT NULL,
  wo_id TEXT,
  wo_number TEXT,
  event_type TEXT NOT NULL,
  score_delta REAL NOT NULL DEFAULT 0,
  reason TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS tech_grade_history_tech_id_idx ON tech_grade_history(tech_id);
CREATE INDEX IF NOT EXISTS tech_grade_history_created_at_idx ON tech_grade_history(created_at DESC);
CREATE INDEX IF NOT EXISTS tech_grade_history_wo_id_idx ON tech_grade_history(wo_id);

ALTER TABLE tech_grades ADD COLUMN IF NOT EXISTS manual_score_override BOOLEAN DEFAULT FALSE;

-- Stage 7: Data Persistence Verification
ALTER TABLE tech_grades ADD COLUMN IF NOT EXISTS last_synced_at TIMESTAMPTZ;

-- Magic Link Open Tracking
ALTER TABLE magic_tokens ADD COLUMN IF NOT EXISTS opened BOOLEAN DEFAULT FALSE;
ALTER TABLE magic_tokens ADD COLUMN IF NOT EXISTS opened_at TIMESTAMPTZ;
ALTER TABLE magic_tokens ADD COLUMN IF NOT EXISTS open_count INTEGER DEFAULT 0;
CREATE INDEX IF NOT EXISTS magic_tokens_tech_id_idx ON magic_tokens(tech_id);
