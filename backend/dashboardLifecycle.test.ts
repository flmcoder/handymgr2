import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

import { shouldInitializeLegacyDashboard } from '../src/dashboardLifecycle.ts';

test('initializes when a legacy dashboard chart host exists', () => {
  assert.equal(shouldInitializeLegacyDashboard(['dashOccupancyDonut']), true);
});

test('does not initialize for charts owned by app.js', () => {
  assert.equal(
    shouldInitializeLegacyDashboard(['dashPmLoadChart', 'dashWoTypeChart', 'dashUrgencyChart']),
    false,
  );
});

test('does not initialize when no chart hosts exist', () => {
  assert.equal(shouldInitializeLegacyDashboard([]), false);
});

test('ignores malformed host identifiers', () => {
  assert.equal(shouldInitializeLegacyDashboard([null, '', 42] as unknown as string[]), false);
});

test('dashboard stats route uses the scope-exact work-order aggregate', async () => {
  const source = await readFile(new URL('./server.ts', import.meta.url), 'utf8');
  const routeStart = source.indexOf("app.get('/api/local/dashboard_stats'");
  const routeEnd = source.indexOf("app.get('/api/local/db_search'", routeStart);
  const route = source.slice(routeStart, routeEnd);

  assert.match(route, /getPropertyGroupFilters\(req\)/);
  assert.match(route, /buildWorkOrdersAnalyticsQuery\(scopeIds\)/);
  assert.match(route, /status_counts/);
  assert.match(route, /postgres_local_dashboard_stats/);
});
