import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { mock, test } from 'node:test';
import {
  resolveTrackerIdentity,
  summarizeTurnTrackerFailures,
  writeTurnTrackerWithDiagnostics,
} from './sync/turnTrackerUpsertPolicy.ts';

const incoming = { trackingUuid: 'new-uuid', turnKey: 'native:T1' };

test('inserts when neither identity exists', () => {
  assert.deepEqual(resolveTrackerIdentity(incoming, null, null), {
    trackingUuid: 'new-uuid', existing: null, rekey: false, conflictingRow: null,
  });
});

test('turn_key match wins even when the incoming tracking_uuid differs (normalization changed the uuid)', () => {
  const row = { trackingUuid: 'old-uuid', turnKey: 'native:T1' };
  const result = resolveTrackerIdentity(incoming, row, null);
  assert.equal(result.trackingUuid, 'old-uuid');
  assert.equal(result.existing, row);
  assert.equal(result.rekey, false);
});

test('legacy row found only by tracking_uuid is re-keyed to the normalized turn_key', () => {
  const legacy = { trackingUuid: 'new-uuid', turnKey: 'unit:p:u:2026-01-01' };
  const result = resolveTrackerIdentity(incoming, null, legacy);
  assert.equal(result.rekey, true);
  assert.equal(result.trackingUuid, 'new-uuid');
  assert.equal(resolveTrackerIdentity(incoming, null, { trackingUuid: 'new-uuid', turnKey: 'native:T1' }).rekey, false);
});

test('two different rows matching the identities resolve to the turn_key row and report the orphan', () => {
  const byKey = { trackingUuid: 'a', turnKey: 'native:T1' };
  const byUuid = { trackingUuid: 'new-uuid', turnKey: 'unit:old' };
  const result = resolveTrackerIdentity(incoming, byKey, byUuid);
  assert.equal(result.trackingUuid, 'a');
  assert.equal(result.rekey, false);
  assert.equal(result.conflictingRow, byUuid);
  assert.equal(resolveTrackerIdentity(incoming, byKey, byKey).conflictingRow, null);
});

test('logs SQLSTATE, constraint and identifying payload for a failed write, then rethrows', async () => {
  const log = mock.method(console, 'error', () => {});
  try {
    const driver = Object.assign(new Error('duplicate key'), {
      code: '23505', constraint: 'unit_turn_tracker_turn_key_unique', detail: 'Key (turn_key)=(native:T1) already exists.',
    });
    await assert.rejects(writeTurnTrackerWithDiagnostics(
      { turn_key: 'native:T1', tracking_uuid: 'u', unit_id: 'unit', property_id: 'prop', status: 'open', metadata: { secret: 'x' } },
      async () => { throw new Error('Failed query', { cause: driver }); },
    ), /Failed query/);
    const [label, details] = log.mock.calls[0].arguments as [string, Record<string, unknown>];
    assert.match(label, /unit_turn_tracker:upsert/);
    assert.equal(details.code, '23505');
    assert.equal(details.constraint, 'unit_turn_tracker_turn_key_unique');
    assert.equal(details.turn_key, 'native:T1');
    assert.equal(details.transient, false);
    assert.equal('metadata' in details, false);
  } finally {
    log.mock.restore();
  }
});

test('returns the write result without logging on success', async () => {
  const log = mock.method(console, 'error', () => {});
  try {
    assert.equal(await writeTurnTrackerWithDiagnostics({ turn_key: 'k' }, async () => 42), 42);
    assert.equal(log.mock.callCount(), 0);
  } finally {
    log.mock.restore();
  }
});

test('batch summary: data errors are skipped, transient and total failures are fatal', () => {
  const dataErr = Object.assign(new Error('dup'), { code: '23505' });
  const netErr = Object.assign(new Error('reset'), { code: 'ECONNRESET' });
  assert.deepEqual(summarizeTurnTrackerFailures(3, 2, [{ turnKey: 'a', error: dataErr }]), { skipped: 1, fatal: null });
  assert.deepEqual(summarizeTurnTrackerFailures(2, 0, [{ turnKey: 'a', error: dataErr }, { turnKey: 'b', error: dataErr }]), { skipped: 2, fatal: dataErr });
  assert.equal(summarizeTurnTrackerFailures(3, 2, [{ turnKey: 'a', error: netErr }]).fatal, netErr);
  assert.deepEqual(summarizeTurnTrackerFailures(0, 0, []), { skipped: 0, fatal: null });
});

test('repository upserts tracker by both identities, updates children in place, and db.ts uses resilient pool options', async () => {
  const repo = await readFile(new URL('./sync/repositories.ts', import.meta.url), 'utf8');
  const start = repo.indexOf('async function upsertTurnTrackerRecord');
  const body = repo.slice(start, repo.indexOf('// ── ', start) === -1 ? undefined : repo.length);
  assert.match(body, /resolveTrackerIdentity\(/);
  assert.match(body, /database\.transaction/);
  assert.match(body, /writeTurnTrackerWithDiagnostics\(/);
  assert.match(body, /\.update\(unitTurnMilestones\)/);
  assert.match(body, /\.update\(unitTurnWorkOrders\)/);
  assert.doesNotMatch(body, /onConflictDoUpdate\(\{\s*target: unitTurnTracker\.trackingUuid/);

  const dbSource = await readFile(new URL('./db.ts', import.meta.url), 'utf8');
  assert.match(dbSource, /\.\.\.buildPostgresResilienceOptions\(\)/);
  const runner = await readFile(new URL('./sync/syncRunner.ts', import.meta.url), 'utf8');
  assert.match(runner, /withDbRetry\(\(\) => def\.upsert\(rows\)/);
});
