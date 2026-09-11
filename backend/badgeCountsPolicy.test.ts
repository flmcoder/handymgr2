import assert from 'node:assert/strict';
import test from 'node:test';
import {
  ACTIVE_INSPECTION_RESIDENT_FILTER,
  ACTIVE_TURN_STATUS_FILTER,
  buildBadgeCountsPayload,
  PROPERTY_IDENTITY_MATCH_FILTER,
  toBadgeCount,
  OPEN_WORK_ORDER_STATUS_FILTER,
  WORK_ORDER_CREATED_AT_EXPR,
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

test('buildBadgeCountsPayload produces authoritative scoped dashboard metrics', () => {
  const p = buildBadgeCountsPayload({
    workOrders: 120,
    urgentWorkOrders: 9,
    workOrdersAge0To7: 40,
    workOrdersAge8To30: 50,
    workOrdersAge31To60: 20,
    workOrdersAge61Plus: 10,
    workOrdersAgeUnknown: 4,
    turns: 8,
    upcomingTurns: 2,
    inspections: 3,
    propertyGroupId: 'g1',
  });
  assert.equal(p.ok, true);
  assert.equal(p.work_orders, 120);
  assert.equal(p.urgent_work_orders, 9);
  assert.deepEqual(p.work_order_aging, { age_0_7: 40, age_8_30: 50, age_31_60: 20, age_61_plus: 10, age_unknown: 4 });
  assert.equal(p.turns, 8);
  assert.equal(p.upcoming_turns, 2);
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

test('WORK_ORDER_CREATED_AT_EXPR falls back to the raw AppFolio payload when the column is null', () => {
  assert.match(WORK_ORDER_CREATED_AT_EXPR, /coalesce\(/);
  assert.match(WORK_ORDER_CREATED_AT_EXPR, /created_at/);
  assert.match(WORK_ORDER_CREATED_AT_EXPR, /raw_json ->> 'CreatedAt'/);
  assert.match(WORK_ORDER_CREATED_AT_EXPR, /::timestamptz/);
  assert.match(WORK_ORDER_CREATED_AT_EXPR, /updated_at/);
  // The raw-value casts stay guarded by a leading-date regex, so a malformed
  // payload string can never raise an invalid-syntax cast for the aggregate.
  const guardedCasts = (WORK_ORDER_CREATED_AT_EXPR.match(/~ '\^\[0-9\]\{4\}/g) || []).length;
  assert.equal(guardedCasts, 3);
  assert.equal((WORK_ORDER_CREATED_AT_EXPR.match(/::timestamptz/g) || []).length, 3);
  // Never includes an unguarded table alias (the badge aggregate has none).
  assert.doesNotMatch(WORK_ORDER_CREATED_AT_EXPR, /wo\.created_at/);
});

test('ACTIVE_TURN_STATUS_FILTER excludes completed and closed turns within the active window', () => {
  assert.match(ACTIVE_TURN_STATUS_FILTER, /not like '%completed%'/);
  assert.match(ACTIVE_TURN_STATUS_FILTER, /not like '%closed%'/);
  assert.match(ACTIVE_TURN_STATUS_FILTER, /updated_at/);
  assert.match(ACTIVE_TURN_STATUS_FILTER, /interval '90 days'/);
});

test('ACTIVE_INSPECTION_RESIDENT_FILTER includes current residents only', () => {
  assert.match(ACTIVE_INSPECTION_RESIDENT_FILTER, /lower\(coalesce\(occ\.status, ''\)\) = 'current'/);
  assert.doesNotMatch(ACTIVE_INSPECTION_RESIDENT_FILTER, /'past'/);
  assert.match(ACTIVE_INSPECTION_RESIDENT_FILTER, /occ\.move_out_date is null or occ\.move_out_date >= current_date/);
  assert.match(ACTIVE_INSPECTION_RESIDENT_FILTER, /occ\.property_id/);
  assert.match(ACTIVE_INSPECTION_RESIDENT_FILTER, /financially responsible/);
  assert.match(ACTIVE_INSPECTION_RESIDENT_FILTER, /occ\.unit_id[\s\S]*or[\s\S]*occ\.occupancy_id/);
  assert.match(ACTIVE_INSPECTION_RESIDENT_FILTER, /occ\.occupancy_id/);
  assert.match(ACTIVE_INSPECTION_RESIDENT_FILTER, /occ\.lease_to is null or occ\.lease_to >= current_date/);
  assert.doesNotMatch(ACTIVE_INSPECTION_RESIDENT_FILTER, /last_inspection_date/);
});

test('PROPERTY_IDENTITY_MATCH_FILTER bridges canonical and raw AppFolio property identifiers', () => {
  assert.match(PROPERTY_IDENTITY_MATCH_FILTER, /p\.id = source_property_id/);
  assert.match(PROPERTY_IDENTITY_MATCH_FILTER, /raw_json/);
  assert.match(PROPERTY_IDENTITY_MATCH_FILTER, /PropertyId/);
});
