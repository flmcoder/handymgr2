import assert from 'node:assert/strict';
import test from 'node:test';
import { buildActiveWorkOrdersUrl, resolveWorkOrderHistoryDays } from './workOrderQueryPolicy.ts';

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