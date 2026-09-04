import assert from 'node:assert/strict';
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