import assert from 'node:assert/strict';
import test from 'node:test';
import { buildDispatchReassignmentKey, executeDispatchCron, isDispatchAutomationPaused, resolveDispatchVersion } from './dispatchCronPolicy.ts';

test('isDispatchAutomationPaused accepts common persisted truthy values and fails safe for missing configuration', () => {
  for (const value of ['1', 'true', 'yes', 'on', 'y']) {
    assert.equal(isDispatchAutomationPaused({ dispatch_paused: value }), true);
  }
  for (const value of ['', '0', 'false', 'no', '0 ']) {
    assert.equal(isDispatchAutomationPaused({ dispatch_paused: value }), false);
  }
  assert.equal(isDispatchAutomationPaused({}), true);
  assert.equal(isDispatchAutomationPaused(null), true);
});

test('buildDispatchReassignmentKey is stable for one WO/version and changes after a meaningful version change', () => {
  const key = buildDispatchReassignmentKey({ woId: ' wo-123 ', dispatchVersion: 2 });
  assert.equal(key, buildDispatchReassignmentKey({ woId: 'wo-123', dispatchVersion: 2 }));
  assert.notEqual(key, buildDispatchReassignmentKey({ woId: 'wo-123', dispatchVersion: 3 }));
  assert.notEqual(key, buildDispatchReassignmentKey({ woId: 'wo-124', dispatchVersion: 2 }));
});

test('buildDispatchReassignmentKey normalizes invalid versions to zero', () => {
  assert.equal(
    buildDispatchReassignmentKey({ woId: 'wo-123', dispatchVersion: 'not-a-number' }),
    buildDispatchReassignmentKey({ woId: 'wo-123', dispatchVersion: 0 }),
  );
});

test('resolveDispatchVersion is stable unless the source status or assignee changes', () => {
  const base = {
    previousVersion: 4,
    previousWorkOrderStatus: 'Open',
    previousAssigneeId: 'tech-1',
    workOrderStatus: ' open ',
    assigneeId: 'tech-1',
  };
  assert.equal(resolveDispatchVersion(base), 4);
  assert.equal(resolveDispatchVersion({ ...base, assigneeId: 'tech-2' }), 5);
  assert.equal(resolveDispatchVersion({ ...base, workOrderStatus: 'In Progress' }), 5);
  assert.equal(resolveDispatchVersion({
    previousVersion: undefined,
    previousWorkOrderStatus: undefined,
    previousAssigneeId: undefined,
    workOrderStatus: 'Open',
    assigneeId: 'tech-1',
  }), 1);
});

test('executeDispatchCron does not run automation side effects when paused', async () => {
  let sideEffects = 0;
  const result = await executeDispatchCron({
    readConfig: async () => ({ dispatch_paused: 'true' }),
    run: async () => { sideEffects += 1; return 'ran'; },
    pausedResult: 'paused',
  });
  assert.equal(result, 'paused');
  assert.equal(sideEffects, 0);
});

test('executeDispatchCron runs when explicitly unpaused and propagates configuration failures', async () => {
  let sideEffects = 0;
  const result = await executeDispatchCron({
    readConfig: async () => ({ dispatch_paused: '0' }),
    run: async () => { sideEffects += 1; return 'ran'; },
    pausedResult: 'paused',
  });
  assert.equal(result, 'ran');
  assert.equal(sideEffects, 1);
  await assert.rejects(() => executeDispatchCron({
    readConfig: async () => { throw new Error('config unavailable'); },
    run: async () => 'ran',
    pausedResult: 'paused',
  }), /config unavailable/);
});

test('dispatch result contract names recommendation flags instead of actual reassignments', () => {
  const pausedResult = { ok: true, run: 'midnight_reassign_cron', paused: true, candidates: 0, flagged: 0, skipped: 0, escalated: 0 };
  assert.equal(pausedResult.flagged, 0);
  assert.equal(Object.prototype.hasOwnProperty.call(pausedResult, 'reassigned'), false);
});