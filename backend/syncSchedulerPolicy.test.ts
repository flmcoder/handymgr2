import assert from 'node:assert/strict';
import test from 'node:test';
import { buildRequestedSyncEndpoints, runSequentially } from './syncSchedulerPolicy.ts';

test('buildRequestedSyncEndpoints appends required bill and v2 syncs', () => {
  assert.deepEqual(
    buildRequestedSyncEndpoints(['v0:work_orders'], {
      billsEnabled: true,
      v2Enabled: true,
      requiredV2Endpoints: ['v2:unit_vacancy'],
    }),
    ['v0:work_orders', 'v0:bills', 'v2:unit_vacancy'],
  );
});

test('buildRequestedSyncEndpoints honors explicit bill and v2 opt-outs', () => {
  assert.deepEqual(
    buildRequestedSyncEndpoints(['v0:work_orders', 'v0:bills', 'v2:unit_vacancy'], {
      billsEnabled: false,
      v2Enabled: false,
      requiredV2Endpoints: ['v2:unit_vacancy'],
    }),
    ['v0:work_orders'],
  );
});

test('buildRequestedSyncEndpoints removes duplicate required endpoints', () => {
  assert.deepEqual(
    buildRequestedSyncEndpoints(['v0:bills', 'v2:unit_vacancy'], {
      billsEnabled: true,
      v2Enabled: true,
      requiredV2Endpoints: ['v2:unit_vacancy'],
    }),
    ['v0:bills', 'v2:unit_vacancy'],
  );
});

test('runSequentially preserves order and limits concurrency to one', async () => {
  const completed: string[] = [];
  let active = 0;
  let maxActive = 0;

  await runSequentially(['v2:first', 'v2:second', 'v2:third'], async (endpoint) => {
    active += 1;
    maxActive = Math.max(maxActive, active);
    await new Promise<void>((resolve) => setImmediate(resolve));
    completed.push(endpoint);
    active -= 1;
  });

  assert.equal(maxActive, 1);
  assert.deepEqual(completed, ['v2:first', 'v2:second', 'v2:third']);
});

test('runSequentially handles an empty endpoint list', async () => {
  let calls = 0;
  await runSequentially([], async () => { calls += 1; });
  assert.equal(calls, 0);
});
