import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildActiveWorkOrdersUrl,
  buildWorkOrderPagination,
  resolveWorkOrderHistoryDays,
} from './workOrderQueryPolicy.ts';

test('buildActiveWorkOrdersUrl requests all active rows without a date window', () => {
  const url = buildActiveWorkOrdersUrl('https://handymgr.example/', '', 5_000);

  assert.equal(url, 'https://handymgr.example/api/local/work_orders?limit=5000');
  assert.equal(url.includes('days='), false);
});

test('buildActiveWorkOrdersUrl encodes the trusted property-group scope', () => {
  assert.equal(
    buildActiveWorkOrdersUrl('', ' group/a & b ', 100),
    '/api/local/work_orders?limit=100&property_group_id=group%2Fa+%26+b',
  );
});

test('buildActiveWorkOrdersUrl includes a bounded page offset without losing scope', () => {
  assert.equal(
    buildActiveWorkOrdersUrl('', 'group-1', 100, 100),
    '/api/local/work_orders?limit=100&offset=100&property_group_id=group-1',
  );
  assert.equal(
    buildActiveWorkOrdersUrl('', '', 100, -1),
    '/api/local/work_orders?limit=100',
  );
});

test('buildActiveWorkOrdersUrl bounds invalid and excessive limits', () => {
  assert.equal(buildActiveWorkOrdersUrl('', '', 'invalid'), '/api/local/work_orders?limit=5000');
  assert.equal(buildActiveWorkOrdersUrl('', '', 0), '/api/local/work_orders?limit=5000');
  assert.equal(buildActiveWorkOrdersUrl('', '', -4), '/api/local/work_orders?limit=1');
  assert.equal(buildActiveWorkOrdersUrl('', '', 50_000), '/api/local/work_orders?limit=20000');
});

test('resolveWorkOrderHistoryDays treats an omitted window as all stored history', () => {
  assert.equal(resolveWorkOrderHistoryDays(undefined), null);
  assert.equal(resolveWorkOrderHistoryDays(null), null);
  assert.equal(resolveWorkOrderHistoryDays(''), null);
  assert.equal(resolveWorkOrderHistoryDays('  '), null);
});

test('resolveWorkOrderHistoryDays bounds supplied and malformed windows', () => {
  assert.equal(resolveWorkOrderHistoryDays(365), 365);
  assert.equal(resolveWorkOrderHistoryDays(0), 1);
  assert.equal(resolveWorkOrderHistoryDays(9_999), 3_650);
  assert.equal(resolveWorkOrderHistoryDays('invalid'), 3_650);
});

test('buildWorkOrderPagination identifies first, middle, and final pages', () => {
  assert.deepEqual(buildWorkOrderPagination(238, 100, 0, 100), {
    total: 238,
    limit: 100,
    offset: 0,
    has_next: true,
    has_previous: false,
  });
  assert.deepEqual(buildWorkOrderPagination(238, 100, 100, 100), {
    total: 238,
    limit: 100,
    offset: 100,
    has_next: true,
    has_previous: true,
  });
  assert.deepEqual(buildWorkOrderPagination(238, 100, 200, 38), {
    total: 238,
    limit: 100,
    offset: 200,
    has_next: false,
    has_previous: true,
  });
});

test('buildWorkOrderPagination clamps invalid values and handles exact page boundaries', () => {
  assert.deepEqual(buildWorkOrderPagination('bad', 100, -1, 'bad'), {
    total: 0,
    limit: 100,
    offset: 0,
    has_next: false,
    has_previous: false,
  });
  assert.equal(buildWorkOrderPagination(200, 100, 100, 100).has_next, false);
});