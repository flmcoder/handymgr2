import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import { splitSqlStatements } from './sqlMigrationPolicy.ts';

test('splitSqlStatements drops comments and splits on statement terminators', () => {
  assert.deepEqual(
    splitSqlStatements('-- header\nALTER TABLE a ADD COLUMN b TEXT;\n\n-- note\nCREATE INDEX i ON a(b);\n'),
    ['ALTER TABLE a ADD COLUMN b TEXT', 'CREATE INDEX i ON a(b)'],
  );
});

test('splitSqlStatements handles empty input and a trailing statement without newline', () => {
  assert.deepEqual(splitSqlStatements(''), []);
  assert.deepEqual(splitSqlStatements(undefined as unknown as string), []);
  assert.deepEqual(splitSqlStatements('SELECT 1;'), ['SELECT 1']);
});

test('work order v0 repair migration adds vendor_trade idempotently', async () => {
  const sql = await readFile(new URL('../db/migrations/2026-10-15_work_orders_v0_columns_repair.sql', import.meta.url), 'utf8');
  const statements = splitSqlStatements(sql);
  assert.ok(statements.length >= 2);
  assert.match(sql, /ADD COLUMN IF NOT EXISTS vendor_trade TEXT/);
  assert.ok(statements.every((s) => /IF (NOT )?EXISTS/i.test(s)));
});

test('every drizzle work-order column is covered by a migration and startup applies the repair', async () => {
  const schema = await readFile(new URL('./schema.ts', import.meta.url), 'utf8');
  const tableSource = schema.slice(schema.indexOf("'appfolio_work_orders'"), schema.indexOf('export const appfolioUsers'));
  const columns = [...tableSource.matchAll(/:\s*(?:text|real|timestamp|jsonb)\('([a-z_]+)'/g)].map((m) => m[1]);
  assert.ok(columns.includes('vendor_trade'));

  const repair = await readFile(new URL('../db/migrations/2026-10-15_work_orders_v0_columns_repair.sql', import.meta.url), 'utf8');
  const original = await readFile(new URL('../db/migrations/2026-01-16_add_missing_v0_fields.sql', import.meta.url), 'utf8');
  const covered = ['id', 'raw_json', 'created_at', 'updated_at', 'wo_number', 'property_id', 'unit_id', 'property_group_id',
    'description', 'category', 'priority', 'status', 'assigned_user_id', 'assigned_user_name', 'vendor_id', 'vendor_name',
    'estimated_amount', 'total_cost'];
  for (const column of columns) {
    assert.ok(covered.includes(column) || repair.includes(column), `${column} missing from repair migration`);
    assert.ok(!original.includes(column) || repair.includes(column) || covered.includes(column), column);
  }

  const server = await readFile(new URL('./server.ts', import.meta.url), 'utf8');
  assert.match(server, /await applyWorkOrderV0ColumnsRepair\(\);[\s\S]*startRecurringSyncScheduler\(\)/);
});
