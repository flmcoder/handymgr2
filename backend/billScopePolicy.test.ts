import assert from 'node:assert/strict';
import test from 'node:test';
import {
  filterBillsForPropertyScope,
  resolveBillScope,
} from './billScopePolicy.ts';

test('resolveBillScope enforces the PM session group over a requested group', () => {
  assert.deepEqual(
    resolveBillScope('pm_readonly', 'session-group', 'other-group'),
    { allowed: true, propertyGroupId: 'session-group' },
  );
});

test('resolveBillScope rejects a PM session without an assigned group', () => {
  assert.deepEqual(
    resolveBillScope('pm_readonly', '', 'requested-group'),
    { allowed: false, propertyGroupId: '' },
  );
});

test('resolveBillScope allows managers to request one group or all groups', () => {
  assert.deepEqual(
    resolveBillScope('manager', '', 'requested-group'),
    { allowed: true, propertyGroupId: 'requested-group' },
  );
  assert.deepEqual(
    resolveBillScope('manager', '', ''),
    { allowed: true, propertyGroupId: '' },
  );
});

test('filterBillsForPropertyScope excludes missing, unmapped, and mixed-group properties', () => {
  const bills = [
    { Id: 'direct', PropertyId: 'property-a' },
    { Id: 'line-item', LineItems: [{ PropertyId: 'property-b' }] },
    { Id: 'missing' },
    { Id: 'unmapped', PropertyId: 'property-z' },
    {
      Id: 'mixed',
      LineItems: [{ PropertyId: 'property-a' }, { PropertyId: 'property-z' }],
    },
  ];

  assert.deepEqual(
    filterBillsForPropertyScope(bills, new Set(['property-a', 'property-b'])).map((bill) => bill.Id),
    ['direct', 'line-item'],
  );
});

test('filterBillsForPropertyScope returns no rows when group membership is empty', () => {
  assert.deepEqual(
    filterBillsForPropertyScope([{ Id: 'bill', PropertyId: 'property-a' }], new Set()),
    [],
  );
});