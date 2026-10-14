import assert from 'node:assert/strict';
import test from 'node:test';
import { groupClosureCandidates } from './closureAssistPolicy.js';

test('groups multiple bill matches under one work order without dropping candidate ids', () => {
  const candidates = [
    { id: 'wo-1:bill-1', work_order_id: 'wo-1', bill_id: 'bill-1' },
    { id: 'wo-1:bill-2', work_order_id: 'wo-1', bill_id: 'bill-2' },
    { id: 'wo-2:bill-3', work_order_id: 'wo-2', bill_id: 'bill-3' },
  ];

  const groups = groupClosureCandidates(candidates);

  assert.equal(groups.length, 2);
  assert.deepEqual(groups[0].candidates.map((candidate) => candidate.id), ['wo-1:bill-1', 'wo-1:bill-2']);
  assert.deepEqual(groups[1].candidates.map((candidate) => candidate.id), ['wo-2:bill-3']);
});

test('keeps candidates without a work-order id separate using their candidate id', () => {
  const groups = groupClosureCandidates([
    { id: 'candidate-a', bill_id: 'bill-1' },
    { id: 'candidate-b', bill_id: 'bill-2' },
  ]);

  assert.equal(groups.length, 2);
});

test('handles empty and malformed candidate lists', () => {
  assert.deepEqual(groupClosureCandidates([]), []);
  assert.deepEqual(groupClosureCandidates(null), []);
  assert.deepEqual(groupClosureCandidates([null, {}, { id: 'candidate-c' }]).map((group) => group.candidates.length), [0, 1, 1]);
});

test('does not let a tier-filtered bill group split the underlying work-order group', () => {
  const all = groupClosureCandidates([
    { id: 'wo-1:bill-1', work_order_id: 'wo-1', confidence_tier: 'very_high' },
    { id: 'wo-1:bill-2', work_order_id: 'wo-1', confidence_tier: 'low' },
  ]);
  const filtered = groupClosureCandidates(all[0].candidates.filter((candidate) => candidate.confidence_tier === 'very_high'));
  const byKey = new Map(all.map((group) => [group.key, group]));

  assert.equal(byKey.get(filtered[0].key).candidates.length, 2);
});
