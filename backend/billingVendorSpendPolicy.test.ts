import assert from 'node:assert/strict';
import test from 'node:test';
import { classifyBillingVendorBucket, parseBillingSpendTimeframe } from './billingVendorSpendPolicy.ts';

test('classifyBillingVendorBucket treats any name containing Fort Lowell as In-House', () => {
  assert.equal(classifyBillingVendorBucket('Fort Lowell Maintenance'), 'in_house');
  assert.equal(classifyBillingVendorBucket('Fort Lowell Realty'), 'in_house');
  assert.equal(classifyBillingVendorBucket('Fort Lowell Maintenance LLC'), 'in_house');
  assert.equal(classifyBillingVendorBucket(' fort lowell realty group '), 'in_house');
});

test('classifyBillingVendorBucket treats unrelated vendors as 3rd-Party', () => {
  assert.equal(classifyBillingVendorBucket('Acme Plumbing'), 'third_party');
  assert.equal(classifyBillingVendorBucket(''), 'third_party');
  assert.equal(classifyBillingVendorBucket(null), 'third_party');
});

test('parseBillingSpendTimeframe defaults to a 180-day lookback', () => {
  const now = new Date('2026-09-04T00:00:00Z');
  const result = parseBillingSpendTimeframe({}, now);
  assert.equal(result.days, 180);
  assert.equal(result.sinceIso, new Date('2026-03-08T00:00:00Z').toISOString());
});

test('parseBillingSpendTimeframe honors an explicit bounded days window', () => {
  const now = new Date('2026-09-04T00:00:00Z');
  const result = parseBillingSpendTimeframe({ days: '30' }, now);
  assert.equal(result.days, 30);
  assert.equal(result.sinceIso, new Date('2026-08-05T00:00:00Z').toISOString());
});

test('parseBillingSpendTimeframe caps an excessive days request', () => {
  const now = new Date('2026-09-04T00:00:00Z');
  const result = parseBillingSpendTimeframe({ days: '999999' }, now, 180, 3650);
  assert.equal(result.days, 3650);
});

test('parseBillingSpendTimeframe prefers an explicit updated_from date over days', () => {
  const now = new Date('2026-09-04T00:00:00Z');
  const result = parseBillingSpendTimeframe({ days: '30', updated_from: '2026-06-01T00:00:00Z' }, now);
  assert.equal(result.sinceIso, new Date('2026-06-01T00:00:00Z').toISOString());
  assert.equal(result.days, 95);
});
