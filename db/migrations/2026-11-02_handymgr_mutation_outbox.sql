-- Durable outbox for optimistic frontend mutations. The background worker drains
-- pending rows at AppFolio's permitted rate; the frontend never waits on AppFolio.
CREATE TABLE IF NOT EXISTS handymgr_mutation_outbox (
  id BIGSERIAL PRIMARY KEY,
  idempotency_key TEXT NOT NULL,
  entity_type TEXT NOT NULL,
  entity_id TEXT NOT NULL,
  action TEXT NOT NULL,
  payload JSONB NOT NULL DEFAULT '{}'::jsonb,
  status TEXT NOT NULL DEFAULT 'sync_pending',
  attempts INTEGER NOT NULL DEFAULT 0,
  next_attempt_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  last_error TEXT NOT NULL DEFAULT '',
  appfolio_response JSONB NOT NULL DEFAULT '{}'::jsonb,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  synced_at TIMESTAMPTZ,
  CONSTRAINT handymgr_mutation_outbox_status_chk
    CHECK (status IN ('sync_pending', 'in_flight', 'synced', 'failed'))
);

CREATE UNIQUE INDEX IF NOT EXISTS handymgr_mutation_outbox_idempotency_idx
  ON handymgr_mutation_outbox(idempotency_key);

CREATE INDEX IF NOT EXISTS handymgr_mutation_outbox_drain_idx
  ON handymgr_mutation_outbox(status, next_attempt_at);
