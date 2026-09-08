import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildTableSearchQuery,
  clampSearchLimit,
  resolveSearchableTable,
  SEARCHABLE_TABLES,
} from './dbSearchPolicy.ts';

test('resolveSearchableTable returns a table for a known key', () => {
  const t = resolveSearchableTable('work_orders');
  assert.ok(t);
  assert.equal(t.table, 'appfolio_work_orders');
});

test('resolveSearchableTable rejects unknown keys (prevents SQL injection into arbitrary tables)', () => {
  assert.equal(resolveSearchableTable('users; drop table users;--'), null);
  assert.equal(resolveSearchableTable(''), null);
  assert.equal(resolveSearchableTable(undefined), null);
  assert.equal(resolveSearchableTable('pg_catalog.pg_authid'), null);
});

test('clampSearchLimit bounds the row cap', () => {
  assert.equal(clampSearchLimit(50), 50);
  assert.equal(clampSearchLimit(0), 1);
  assert.equal(clampSearchLimit(5000), 200);
  assert.equal(clampSearchLimit('not-a-number'), 50);
  assert.equal(clampSearchLimit(undefined), 50);
});

test('buildTableSearchQuery parameterizes the search term (no string interpolation of user input)', () => {
  const q = buildTableSearchQuery('work_orders', "leak'; drop table x;--", 25);
  assert.ok(q);
  // Term must appear only as a bound parameter, never concatenated into SQL.
  assert.ok(q.sql.indexOf("leak'") === -1);
  assert.ok(q.params.indexOf("%leak'; drop table x;--%") !== -1);
  assert.match(q.sql, /from appfolio_work_orders/);
  assert.match(q.sql, /limit \$\d+/);
});

test('buildTableSearchQuery without a term emits no WHERE clause', () => {
  const q = buildTableSearchQuery('properties', '', 10);
  assert.ok(q);
  assert.ok(!/where/i.test(q.sql));
  assert.deepEqual(q.params, [10]);
});

test('buildTableSearchQuery returns null for non-allow-listed table', () => {
  assert.equal(buildTableSearchQuery('device_otps', 'x', 10), null);
});

test('every allow-listed table has at least one searchable column', () => {
  SEARCHABLE_TABLES.forEach((t) => {
    assert.ok(t.columns.length > 0, t.key + ' columns');
    assert.ok(t.searchColumns.length > 0, t.key + ' searchColumns');
  });
});
