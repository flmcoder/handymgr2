-- Additive portal session model. Existing links retain legacy single-use behavior.
ALTER TABLE magic_tokens
  ADD COLUMN IF NOT EXISTS session_id UUID DEFAULT gen_random_uuid(),
  ADD COLUMN IF NOT EXISTS revoked_at TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS session_mode TEXT NOT NULL DEFAULT 'legacy_single_use';

UPDATE magic_tokens
SET session_id = gen_random_uuid()
WHERE session_id IS NULL;

ALTER TABLE magic_tokens
  ALTER COLUMN session_id SET NOT NULL;

CREATE UNIQUE INDEX IF NOT EXISTS magic_tokens_session_id_idx
  ON magic_tokens(session_id);

CREATE INDEX IF NOT EXISTS magic_tokens_revoked_expires_idx
  ON magic_tokens(revoked_at, expires_at);

CREATE TABLE IF NOT EXISTS portal_submissions (
  id BIGSERIAL PRIMARY KEY,
  session_id UUID NOT NULL,
  wo_id TEXT NOT NULL,
  session_mode TEXT NOT NULL,
  action TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT '',
  note_text TEXT NOT NULL DEFAULT '',
  idempotency_key TEXT,
  submitted_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  outcome JSONB NOT NULL DEFAULT '{}'::jsonb
);

CREATE INDEX IF NOT EXISTS portal_submissions_session_idx
  ON portal_submissions(session_id, submitted_at DESC);

CREATE INDEX IF NOT EXISTS portal_submissions_work_order_idx
  ON portal_submissions(wo_id, submitted_at DESC);

CREATE UNIQUE INDEX IF NOT EXISTS portal_submissions_idempotency_idx
  ON portal_submissions(session_id, idempotency_key)
  WHERE idempotency_key IS NOT NULL AND idempotency_key <> '';