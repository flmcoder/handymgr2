import assert from 'node:assert/strict';
import test from 'node:test';
import { buildRequestedSyncEndpoints } from './syncSchedulerPolicy.ts';

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