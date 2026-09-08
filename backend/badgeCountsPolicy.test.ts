import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildBadgeCountsPayload,
  toBadgeCount,
  OPEN_WORK_ORDER_STATUS_FILTER,
} from './badgeCountsPolicy.ts';

test('toBadgeCount returns integers for valid numbers', () => {
  assert.equal(toBadgeCount(42), 42);
  assert.equal(toBadgeCount('17'), 17);
  assert.equal(toBadgeCount(9.9), 9);
});

test('toBadgeCount clamps invalid and negative values to zero', () => {
  assert.equal(toBadgeCount(undefined), 0);
  assert.equal(toBadgeCount(null), 0);
  assert.equal(toBadgeCount(NaN), 0);
  assert.equal(toBadgeCount(-5), 0);
  assert.equal(toBadgeCount('abc'), 0);
});

test('buildBadgeCountsPayload produces ok payload with clamped counts', () => {
  const p = buildBadgeCountsPayload({ workOrders: 120, turns: 8, inspections: 3, propertyGroupId: 'g1' });
  assert.equal(p.ok, true);
  assert.equal(p.work_orders, 120);
  assert.equal(p.turns, 8);
  assert.equal(p.inspections, 3);
  assert.equal(p.property_group_id, 'g1');
  assert.equal(p.source, 'postgres_local');
});

test('buildBadgeCountsPayload tolerates missing counts', () => {
  const p = buildBadgeCountsPayload({});
  assert.deepEqual([p.work_orders, p.turns, p.inspections], [0, 0, 0]);
  assert.equal(p.property_group_id, '');
});

test('OPEN_WORK_ORDER_STATUS_FILTER excludes terminal statuses', () => {
  assert.match(OPEN_WORK_ORDER_STATUS_FILTER, /not like '%completed%'/);
  assert.match(OPEN_WORK_ORDER_STATUS_FILTER, /not like '%cancel%'/);
  assert.match(OPEN_WORK_ORDER_STATUS_FILTER, /not like '%no need to bill%'/);
});
