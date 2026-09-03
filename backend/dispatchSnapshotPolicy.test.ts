import assert from 'node:assert/strict';
import test from 'node:test';
import { shouldRefreshDispatchSnapshot } from './dispatchSnapshotPolicy.ts';

test('shouldRefreshDispatchSnapshot refreshes cron and manual snapshots by default', () => {
  assert.equal(shouldRefreshDispatchSnapshot({}), true);
  assert.equal(shouldRefreshDispatchSnapshot({ source: 'cron' }), true);
});

test('shouldRefreshDispatchSnapshot skips refresh for persisted UI reads', () => {
  assert.equal(shouldRefreshDispatchSnapshot({ read_only: '1' }), false);
  assert.equal(shouldRefreshDispatchSnapshot({ readOnly: 'true' }), false);
});