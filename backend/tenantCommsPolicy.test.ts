import assert from 'node:assert/strict';
import test from 'node:test';
import { parseTenantCommsQuery } from './tenantCommsPolicy.ts';

test('tenant communications defaults to a bounded first page', () => {
  assert.deepEqual(parseTenantCommsQuery({}), { limit: 50, offset: 0, page: 1 });
});

test('tenant communications caps limits and calculates page offsets', () => {
  assert.deepEqual(parseTenantCommsQuery({ limit: '5000', page: '3' }), {
    limit: 100,
    offset: 200,
    page: 3,
  });
});

test('tenant communications accepts an explicit nonnegative offset', () => {
  assert.deepEqual(parseTenantCommsQuery({ limit: '60', offset: '120' }), {
    limit: 60,
    offset: 120,
    page: 3,
  });
});