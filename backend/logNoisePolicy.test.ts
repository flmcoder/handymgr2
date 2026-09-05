import assert from 'node:assert/strict';
import test from 'node:test';
import {
  formatSignInLine,
  formatSyncSummaryLine,
  shouldLogSyncSummary,
} from './logNoisePolicy.ts';

test('shouldLogSyncSummary skips zero-upsert completed ticks', () => {
  assert.equal(shouldLogSyncSummary({ status: 'completed', rowsUpserted: 0, rowsSkipped: 12, pagesCompleted: 3 }), false);
});

test('shouldLogSyncSummary logs when rows were upserted', () => {
  assert.equal(shouldLogSyncSummary({ status: 'completed', rowsUpserted: 5 }), true);
});

test('shouldLogSyncSummary logs failures even with zero upserts', () => {
  assert.equal(shouldLogSyncSummary({ status: 'failed', rowsUpserted: 0 }), true);
});

test('shouldLogSyncSummary handles missing summary and missing fields', () => {
  assert.equal(shouldLogSyncSummary(undefined), false);
  assert.equal(shouldLogSyncSummary(null), false);
  assert.equal(shouldLogSyncSummary({}), false);
});

test('formatSyncSummaryLine includes upserted, skipped, and pages', () => {
  assert.equal(
    formatSyncSummaryLine('v0:work_orders', { status: 'completed', rowsUpserted: 42, rowsSkipped: 7, pagesCompleted: 2 }),
    '[SYNC] v0:work_orders | status=completed | upserted=42 | skipped=7 | pages=2',
  );
});

test('formatSyncSummaryLine defaults missing numeric fields to zero', () => {
  assert.equal(
    formatSyncSummaryLine('v0:bills', { status: 'failed' }),
    '[SYNC] v0:bills | status=failed | upserted=0 | skipped=0 | pages=0',
  );
});

test('formatSignInLine renders a password sign-in', () => {
  assert.equal(
    formatSignInLine({ method: 'password', userName: 'alice', role: 'manager' }),
    '[auth] sign-in | method=password | user="alice" | role=manager',
  );
});

test('formatSignInLine renders an otp sign-in with email and scope', () => {
  assert.equal(
    formatSignInLine({ method: 'otp', userName: 'Bob PM', role: 'pm_readonly', email: 'bob@example.com', scopeUuid: 'abc-123' }),
    '[auth] sign-in | method=otp | user="Bob PM" | email="bob@example.com" | role=pm_readonly | scope=abc-123',
  );
});

test('formatSignInLine marks signed-fallback tokens', () => {
  assert.equal(
    formatSignInLine({ method: 'password', userName: 'alice', role: 'full', tokenKind: 'signed-fallback' }),
    '[auth] sign-in | method=password | user="alice" | role=full | token=signed-fallback',
  );
});

test('formatSignInLine omits scope for non-otp methods and strips stray quotes', () => {
  const line = formatSignInLine({ method: 'device_setup', userName: 'dev"ice', role: 'full', scopeUuid: 'ignored' });
  assert.equal(line, '[auth] sign-in | method=device_setup | user="device" | role=full');
});
