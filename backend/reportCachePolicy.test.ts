import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildReportCacheKey,
  chunkReportRows,
  extractReportRows,
  shouldHydrateReport,
} from './reportCachePolicy.ts';

test('buildReportCacheKey is stable across object key order and differs by scope', () => {
  const first = buildReportCacheKey('bill_detail', {
    occurred_on_from: '2026-01-01',
    properties: { property_groups_ids: ['group-a'] },
  });
  const reordered = buildReportCacheKey('bill_detail', {
    properties: { property_groups_ids: ['group-a'] },
    occurred_on_from: '2026-01-01',
  });
  const otherScope = buildReportCacheKey('bill_detail', {
    properties: { property_groups_ids: ['group-b'] },
    occurred_on_from: '2026-01-01',
  });

  assert.equal(first, reordered);
  assert.notEqual(first, otherScope);
});

test('extractReportRows supports AppFolio array, results, and data payloads', () => {
  assert.deepEqual(extractReportRows([{ id: 1 }]), [{ id: 1 }]);
  assert.deepEqual(extractReportRows({ results: [{ id: 2 }] }), [{ id: 2 }]);
  assert.deepEqual(extractReportRows({ data: [{ id: 3 }] }), [{ id: 3 }]);
  assert.deepEqual(extractReportRows({ error: 'invalid' }), []);
});

test('shouldHydrateReport includes the configured threshold boundary', () => {
  assert.equal(shouldHydrateReport(4_999, 5_000), false);
  assert.equal(shouldHydrateReport(5_000, 5_000), true);
  assert.equal(shouldHydrateReport(Number.NaN, 5_000), false);
});

test('chunkReportRows preserves all rows in bounded batches', () => {
  const rows = Array.from({ length: 1_201 }, (_, index) => index);
  const chunks = chunkReportRows(rows, 500);

  assert.deepEqual(chunks.map((chunk) => chunk.length), [500, 500, 201]);
  assert.deepEqual(chunks.flat(), rows);
});