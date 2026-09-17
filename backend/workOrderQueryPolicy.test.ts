import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import {
  buildActiveWorkOrdersUrl,
  buildWorkOrderPagination,
  resolveExactWorkOrderReference,
  resolveWorkOrderLookupSearch,
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

test('global work-order lookup requires three non-whitespace characters', () => {
  assert.equal(resolveWorkOrderLookupSearch(''), '');
  assert.equal(resolveWorkOrderLookupSearch(' 12 '), '');
  assert.equal(resolveWorkOrderLookupSearch(' WO-554 '), 'WO-554');
});

test('global work-order lookup bounds oversized search terms', () => {
  assert.equal(resolveWorkOrderLookupSearch('x'.repeat(500)).length, 180);
});

test('cross-group lookup accepts only explicit work-order references', () => {
  assert.equal(resolveExactWorkOrderReference('#WO-554'), 'WO-554');
  assert.equal(resolveExactWorkOrderReference('554'), '554');
  assert.equal(resolveExactWorkOrderReference('plumbing'), '');
  assert.equal(resolveExactWorkOrderReference('property 554'), '');
});

test('work-order endpoint keeps default scope and redacts global lookup rows', async () => {
  const source = await readFile(new URL('./server.ts', import.meta.url), 'utf8');
  const routeStart = source.indexOf("app.get('/api/local/work_orders'");
  const routeEnd = source.indexOf("app.get('/api/local/db_search'", routeStart);
  const route = source.slice(routeStart, routeEnd);

  assert.match(route, /resolveWorkOrderLookupSearch\(req\.query\.search\)/);
  assert.match(route, /resolveExactWorkOrderReference\(searchQuery\)/);
  assert.match(route, /buildWorkOrderPropertyGroupScopeSql\('wo', '\$6'\)/);
  assert.match(route, /property_group_id\s*=\s*ANY\(\$\{scopeIds\}::text\[\]\)/i);
  assert.match(route, /PropertyGroupIds'\s*\?\|\s*\$2::text\[\]/);
  assert.match(route, /case when can_open then id else '' end as id/i);
  assert.match(route, /case when can_open then raw_json else '\{\}'::jsonb end as raw_json/i);
  assert.match(route, /global_search_result/);
  assert.doesNotMatch(route, /appfolio_work_orders\.property_name/);
});

test('PM work-order detail endpoints enforce scope after global lookup', async () => {
  const source = await readFile(new URL('./server.ts', import.meta.url), 'utf8');
  assert.match(source, /async function resolveAccessibleWorkOrder[\s\S]*PropertyGroupIds'[\s\S]*limit 1/);
  assert.match(source, /when wo\.id = \$1 then 0[\s\S]*when wo\.wo_number = \$1 then 1/);
  assert.match(source, /nullif\(wo\.raw_json->>'work_order_uuid', ''\)/);
  assert.equal((source.match(/await resolveAccessibleWorkOrder\(req, workOrderRef\)/g) || []).length, 3);
});

test('frontend sends debounced lookup terms and blocks opening restricted rows', async () => {
  const source = await readFile(new URL('../src/app.js', import.meta.url), 'utf8');
  assert.match(source, /api\/local\/work_orders\?limit=100&offset=0&search=' \+ encodeURIComponent\(searchQuery\)/);
  assert.match(source, /await fetchWorkOrderLookup\(searchTerm\)/);
  assert.match(source, /WORK_ORDER_LOOKUP_RESULTS/);
  assert.match(source, /if \(wo\.canOpen === false\)[\s\S]*Lookup only/);
  assert.match(source, /propertyGroupName:\s*wo\.propertyGroupName/);
});

test('session transitions invalidate scoped requests and clear lookup state', async () => {
  const source = await readFile(new URL('../src/app.js', import.meta.url), 'utf8');
  assert.match(source, /function resetInMemoryDataForSessionTransition\(\)\s*\{\s*(?:stopDataSourceFreshnessMonitor\(\);\s*)?_scopeRequestGeneration\+\+;\s*_workOrdersRequestGeneration\+\+;\s*clearWorkOrderLookup\(\)/);
  assert.match(source, /function forceProxySessionExpiryLockout[\s\S]{0,900}resetInMemoryDataForSessionTransition\(\)/);
  assert.match(source, /_workOrdersRequestGeneration\+\+;/);
  assert.match(source, /window\.WORK_ORDERS = \[\]/);
  assert.match(source, /window\.AppDB\.properties\.clear\(\)/);
  assert.match(source, /if \(woSearch\) woSearch\.value = ''/);
});
