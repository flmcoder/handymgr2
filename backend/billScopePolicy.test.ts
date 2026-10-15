import assert from 'node:assert/strict';
import test from 'node:test';
import {
  filterBillsForPropertyScope,
  isGlobalScopeToken,
  normalizeRequestedScopeIds,
  resolveBillScope,
} from './billScopePolicy.ts';

test('resolveBillScope enforces the PM session group over a requested group', () => {
  assert.deepEqual(
    resolveBillScope('pm_readonly', 'session-group', 'other-group'),
    { allowed: true, propertyGroupId: 'session-group', propertyGroupIds: ['session-group'] },
  );
});

test('resolveBillScope rejects a PM session without an assigned group', () => {
  assert.deepEqual(
    resolveBillScope('pm_readonly', '', 'requested-group'),
    { allowed: false, propertyGroupId: '', propertyGroupIds: [] },
  );
});

test('resolveBillScope allows managers to request one group or all groups', () => {
  assert.deepEqual(
    resolveBillScope('manager', '', 'requested-group'),
    { allowed: true, propertyGroupId: 'requested-group', propertyGroupIds: ['requested-group'] },
  );
  assert.deepEqual(
    resolveBillScope('manager', '', ''),
    { allowed: true, propertyGroupId: '', propertyGroupIds: [] },
  );
});

test('resolveBillScope carries a multi-group union for PM sessions', () => {
  assert.deepEqual(
    resolveBillScope('pm_readonly', ['group-a', 'group-b', 'group-a'], 'group-a'),
    { allowed: true, propertyGroupId: 'group-a', propertyGroupIds: ['group-a', 'group-b'] },
  );
  assert.deepEqual(
    resolveBillScope('manager', '', 'group-a,group-b'),
    { allowed: true, propertyGroupId: 'group-a', propertyGroupIds: ['group-a', 'group-b'] },
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
test('global scope tokens resolve to an unrestricted requested scope', () => {
  for (const token of ['__all__', 'ALL', '*', 'All Groups', ' all properties ', 'all-groups', 'all_groups']) {
    assert.equal(isGlobalScopeToken(token), true, token);
    assert.deepEqual(normalizeRequestedScopeIds([token]), [], token);
  }
  assert.deepEqual(normalizeRequestedScopeIds(['group-a', '__all__']), []);
});

test('real group ids and empty input are preserved by scope normalization', () => {
  assert.equal(isGlobalScopeToken(''), false);
  assert.equal(isGlobalScopeToken(undefined), false);
  assert.equal(isGlobalScopeToken('group-a'), false);
  assert.deepEqual(normalizeRequestedScopeIds([]), []);
  assert.deepEqual(normalizeRequestedScopeIds(['group-a', 'group-b']), ['group-a', 'group-b']);
});

test('PM sessions ignore a global request token and keep every assigned group', () => {
  const requested = normalizeRequestedScopeIds(['__all__']);
  assert.deepEqual(
    resolveBillScope('pm_readonly', ['g1', 'g2'], requested),
    { allowed: true, propertyGroupId: 'g1', propertyGroupIds: ['g1', 'g2'] },
  );
  assert.deepEqual(
    resolveBillScope('manager', [], requested),
    { allowed: true, propertyGroupId: '', propertyGroupIds: [] },
  );
});
