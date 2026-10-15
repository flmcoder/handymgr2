import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import {
  buildActiveWorkOrdersUrl,
  buildLocalWorkOrdersCountQuery,
  buildLocalWorkOrdersQuery,
  buildWorkOrderStatusScopeSql,
  buildWorkOrderPagination,
  resolveExactWorkOrderReference,
  resolveWorkOrderLookupSearch,
  resolveWorkOrderStatusScope,
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

test('work-order lookup status scope distinguishes active and inactive searches', () => {
  assert.equal(resolveWorkOrderStatusScope('inactive'), 'inactive');
  assert.equal(resolveWorkOrderStatusScope('completed'), 'active');
  assert.equal(resolveWorkOrderStatusScope(undefined), 'active');
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
  assert.match(route, /buildLocalWorkOrdersQuery\(/);
  assert.match(route, /buildLocalWorkOrdersCountQuery\(/);
  assert.doesNotMatch(route, /queryClient`/);
  assert.match(route, /PropertyGroupIds'\s*\?\|\s*\$2::text\[\]/);
  assert.match(route, /case when can_open then id else '' end as id/i);
  assert.match(route, /case when can_open then raw_json else '\{\}'::jsonb end as raw_json/i);
  assert.match(route, /global_search_result/);
  assert.doesNotMatch(route, /appfolio_work_orders\.property_name/);
  assert.match(route, /resolveWorkOrderStatusScope\(req\.query\.status_scope\)/);
  assert.match(route, /and \$\{statusScopeSql\}/);
});

test('work-order frontend searches the server across active and inactive status scopes', async () => {
  const source = await readFile(new URL('../src/app.js', import.meta.url), 'utf8');
  const lookupStart = source.indexOf('async function fetchWorkOrderLookup(');
  const lookupEnd = source.indexOf('async function fetchInactiveWorkOrders(', lookupStart);
  const lookup = source.slice(lookupStart, lookupEnd);
  const searchBindingStart = source.indexOf("$('#woSearch').addEventListener('input'");
  const searchBindingEnd = source.indexOf("if ($('#woPriorityFilter'))", searchBindingStart);
  const searchBinding = source.slice(searchBindingStart, searchBindingEnd);

  assert.match(lookup, /status_scope=' \+ encodeURIComponent\(currentWOTab === 'inactive' \? 'inactive' : 'active'\)/);
  assert.match(searchBinding, /await fetchWorkOrderLookup\(searchTerm\)/);
  assert.doesNotMatch(searchBinding, /currentWOTab === 'inactive'[\s\S]{0,100}clearWorkOrderLookup\(\)/);
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

test('work-order row query and count query share the same WHERE clause and params', () => {
  for (const statusScope of ['active', 'inactive'] as const) {
    for (const scopeIds of [[], ['g1', 'g2']]) {
      const opts = { scopeIds, statusScope, statusFilter: 'Assigned' };
      const rows = buildLocalWorkOrdersQuery({ ...opts, limit: 100, offset: 200 });
      const count = buildLocalWorkOrdersCountQuery(opts);
      const rowWhere = rows.sql.split(/\bwhere\b/).slice(1).join('where').split('order by')[0].trim();
      const countWhere = count.sql.split(/\bwhere\b/).slice(1).join('where').trim();
      assert.equal(rowWhere, countWhere);
      assert.deepEqual(rows.params.slice(0, -2), count.params);
    }
  }
});

test('work-order row query binds values and never binds SQL fragments as parameters', () => {
  const { sql, params } = buildLocalWorkOrdersQuery({
    scopeIds: ['g1'], statusScope: 'inactive', statusFilter: '', limit: 50, offset: 0,
  });
  assert.deepEqual(params, [['g1'], 50, 0]);
  assert.match(sql, /limit \$2 offset \$3/);
  assert.match(sql, /like '%completed%'/);
  assert.ok(params.every((p) => typeof p !== 'string' || !/coalesce|like/i.test(p)));
});

test('unscoped work-order query has no group clause and status filter is optional', () => {
  const { sql, params } = buildLocalWorkOrdersQuery({
    scopeIds: [], statusScope: 'active', statusFilter: '', limit: 100, offset: 0,
  });
  assert.deepEqual(params, [100, 0]);
  assert.doesNotMatch(sql, /PropertyGroupIds/);
  assert.doesNotMatch(sql, /coalesce\(wo\.status, ''\) = /);
  assert.match(sql, /limit \$1 offset \$2/);
});

test('work-order row query supports the legacy schema without work_order_uuid', () => {
  const { sql } = buildLocalWorkOrdersQuery({
    scopeIds: [], statusScope: 'active', statusFilter: '', limit: 1, offset: 0, includeWorkOrderUuid: false,
  });
  assert.match(sql, /null::text as work_order_uuid/);
  assert.doesNotMatch(sql, /wo\.work_order_uuid/);
});

test('active status scope matches the dashboard open filter and qualifies the column', () => {
  const active = buildWorkOrderStatusScopeSql('active', 'wo.status');
  assert.match(active, /lower\(wo\.status\)/);
  assert.doesNotMatch(active, /[^.]\bstatus\b/);
  assert.match(buildWorkOrderStatusScopeSql('inactive'), /like '%cancel%'/);
});

test('work-order query applies status filter and scope params in order', () => {
  const { params, sql } = buildLocalWorkOrdersQuery({
    scopeIds: ['g1'], statusScope: 'active', statusFilter: 'Assigned', limit: 10, offset: 5,
  });
  assert.deepEqual(params, [['g1'], 'Assigned', 10, 5]);
  assert.match(sql, /= \$2::text/);
  assert.match(sql, /limit \$3 offset \$4/);
});

test('frontend and backend both treat All Groups sentinels as unscoped', async () => {
  const app = await readFile(new URL('../src/app.js', import.meta.url), 'utf8');
  const fnStart = app.indexOf('function normalizeGroupSelectionValue(');
  const fn = app.slice(fnStart, app.indexOf('function getEffectiveGroupId', fnStart));
  assert.match(fn, /'all groups'/);
  assert.match(fn, /'__all__'/);

  const server = await readFile(new URL('./server.ts', import.meta.url), 'utf8');
  assert.match(server, /function parsePropertyGroupIds[\s\S]{0,200}normalizeRequestedScopeIds\(/);
});
