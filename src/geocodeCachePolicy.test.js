import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

import {
  createGeocodeCoordinateCache,
  createConcurrencyLimiter,
  mapWithConcurrency,
  normalizeGeocodeAddress,
} from './geocodeCachePolicy.js';

test('geocode address keys normalize case and whitespace', () => {
  assert.equal(normalizeGeocodeAddress('  51 N   Euclid Ave  '), '51 n euclid ave');
});

test('coordinate cache persists valid coordinates and reloads them', () => {
  const values = new Map();
  const storage = {
    getItem: (key) => values.get(key) || null,
    setItem: (key, value) => values.set(key, value),
  };
  const cache = createGeocodeCoordinateCache(() => storage, 'geo-test');

  assert.equal(cache.set('51 N Euclid Ave', { lat: 32.2, lon: -110.9 }), true);
  cache.flush();

  const reloaded = createGeocodeCoordinateCache(() => storage, 'geo-test');
  assert.deepEqual(reloaded.get(' 51 n  euclid ave '), { lat: 32.2, lon: -110.9 });
});

test('coordinate cache rejects invalid entries, handles unavailable storage, and evicts oldest entries', () => {
  const cache = createGeocodeCoordinateCache(() => { throw new Error('storage blocked'); }, 'geo-test', 2);

  assert.equal(cache.set('', { lat: 32, lon: -110 }), false);
  assert.equal(cache.set('invalid', { lat: 91, lon: 0 }), false);
  assert.equal(cache.set('one', { lat: 1, lon: 1 }), true);
  assert.equal(cache.set('two', { lat: 2, lon: 2 }), true);
  assert.equal(cache.set('three', { lat: 3, lon: 3 }), true);
  assert.equal(cache.get('one'), null);
  assert.deepEqual(cache.get('three'), { lat: 3, lon: 3 });
  assert.doesNotThrow(() => cache.flush());
});

test('mapWithConcurrency limits workers, preserves result order, and handles empty input', async () => {
  let active = 0;
  let maximumActive = 0;
  const results = await mapWithConcurrency([1, 2, 3, 4, 5], 3, async (value) => {
    active++;
    maximumActive = Math.max(maximumActive, active);
    await new Promise((resolve) => setTimeout(resolve, 2));
    active--;
    return value * 2;
  });

  assert.equal(maximumActive, 3);
  assert.deepEqual(results, [2, 4, 6, 8, 10]);
  assert.deepEqual(await mapWithConcurrency([], 3, async (value) => value), []);
});

test('mapWithConcurrency rejects mapper failures', async () => {
  await assert.rejects(
    mapWithConcurrency([1, 2], 1, async (value) => {
      if (value === 2) throw new Error('mapper failed');
      return value;
    }),
    /mapper failed/,
  );
});

test('shared geocode concurrency limiter caps overlapping map runs', async () => {
  const limiter = createConcurrencyLimiter(3);
  let active = 0;
  let maximumActive = 0;
  const runTask = async (value) => limiter.run(async () => {
    active++;
    maximumActive = Math.max(maximumActive, active);
    await new Promise((resolve) => setTimeout(resolve, 2));
    active--;
    return value;
  });

  const results = await Promise.all([
    mapWithConcurrency([1, 2, 3, 4], 3, runTask),
    mapWithConcurrency([5, 6, 7, 8], 3, runTask),
  ]);

  assert.equal(maximumActive, 3);
  assert.deepEqual(results.flat(), [1, 2, 3, 4, 5, 6, 7, 8]);
});

test('dashboard never derives the work-order navigation badge from the current page', async () => {
  const source = await readFile(new URL('./app.js', import.meta.url), 'utf8');
  const dashboardStart = source.indexOf('function renderDashboardKPIs()');
  const dashboardEnd = source.indexOf('// ---------------------------------------------------------------------------\n// Geo-chart panel', dashboardStart);
  const dashboard = source.slice(dashboardStart, dashboardEnd);
  const cacheBranch = source.slice(source.indexOf("if (fresh) {", source.indexOf('async function initApp()')), source.indexOf('// STEP 2: Pre-flight', source.indexOf('async function initApp()')));

  assert.match(dashboard, /\$\('#woBadge'\)\.textContent\s*=\s*getNavBadgeDisplay\('work_orders'\)/);
  assert.doesNotMatch(dashboard, /#woBadge[\s\S]{0,180}(?:openWOs\.length|WORK_ORDERS\.length)/);
  assert.match(cacheBranch, /fetchNavBadgeTotals\(true\)/);
});

test('dashboard defers address geocoding until critical startup reads settle and bounds concurrency to three', async () => {
  const source = await readFile(new URL('./app.js', import.meta.url), 'utf8');
  const geoStart = source.indexOf('function renderDashboardGeoCharts(force)');
  const geoEnd = source.indexOf('// Portfolio Health Sunburst', geoStart);
  const geo = source.slice(geoStart, geoEnd);

  assert.match(geo, /_initialCriticalDataReady\.then\(async function\(\)/);
  assert.match(geo, /mapWithConcurrency\(Array\.from\(addressGroups\.values\(\)\),\s*3,/);
  assert.match(geo, /_geocodeCoordinateCache\.get\(group\.address\)/);
  assert.match(geo, /_geocodeCoordinateCache\.set\(group\.address, resolved\)/);
});