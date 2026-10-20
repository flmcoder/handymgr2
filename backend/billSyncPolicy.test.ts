import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';
import { getTableConfig } from 'drizzle-orm/pg-core';
import { appfolioBills } from './schema.ts';
import { normalizeBillSyncRow } from './sync/billSyncPolicy.ts';

test('ensureBillsTable creates or repairs every column declared by appfolioBills', () => {
  const serverSource = readFileSync(new URL('./server.ts', import.meta.url), 'utf8');
  const ensureBillsTable = serverSource.match(
    /async function ensureBillsTable\(\): Promise<void> \{([\s\S]*?)\n\}/,
  )?.[1];

  assert.ok(ensureBillsTable, 'ensureBillsTable implementation should exist');

  const createTableSql = ensureBillsTable.match(
    /CREATE TABLE IF NOT EXISTS appfolio_bills \(([\s\S]*?)\n\s*\)/,
  )?.[1];
  assert.ok(createTableSql, 'ensureBillsTable should create appfolio_bills');

  const coveredColumns = new Set<string>();
  for (const line of createTableSql.split('\n')) {
    const columnName = line.trim().match(/^([a-z_][a-z0-9_]*)\s+/i)?.[1];
    if (columnName) coveredColumns.add(columnName);
  }
  for (const match of ensureBillsTable.matchAll(
    /ALTER TABLE appfolio_bills ADD COLUMN IF NOT EXISTS ([a-z_][a-z0-9_]*)/gi,
  )) {
    coveredColumns.add(match[1]);
  }

  const declaredColumns = getTableConfig(appfolioBills).columns.map((column) => column.name);
  assert.deepEqual(
    declaredColumns.filter((columnName) => !coveredColumns.has(columnName)),
    [],
    'every appfolioBills schema column must be present in the initializer DDL or repair statements',
  );
});

test('normalizeBillSyncRow maps a direct-property AppFolio bill', () => {
  const normalized = normalizeBillSyncRow({
    Id: 'bill-1',
    BillNumber: 'INV-42',
    VendorId: 'vendor-1',
    VendorName: 'Acme Repair',
    PropertyId: 'property-1',
    Status: 'Unpaid',
    BillTotalAmount: '$1,234.50',
    InvoiceDate: '2026-09-01',
  });

  assert.equal(normalized?.id, 'bill-1');
  assert.equal(normalized?.propertyId, 'property-1');
  assert.equal(normalized?.billTotalAmount, 1234.5);
  assert.equal(normalized?.invoiceDate?.toISOString(), '2026-09-01T00:00:00.000Z');
});

test('normalizeBillSyncRow accepts one line-item property when the bill has no direct property', () => {
  const normalized = normalizeBillSyncRow({
    Id: 'bill-2',
    LineItems: [{ PropertyId: 'property-2' }, { PropertyId: 'property-2' }],
  });

  assert.equal(normalized?.propertyId, 'property-2');
});

test('normalizeBillSyncRow leaves mixed-property bills unscoped', () => {
  const normalized = normalizeBillSyncRow({
    Id: 'bill-3',
    LineItems: [{ PropertyId: 'property-1' }, { PropertyId: 'property-2' }],
  });

  assert.equal(normalized?.propertyId, null);
});

test('normalizeBillSyncRow rejects malformed rows without an ID', () => {
  assert.equal(normalizeBillSyncRow({ PropertyId: 'property-1' }), null);
  assert.equal(normalizeBillSyncRow(null), null);
});

test('normalizeBillSyncRow clears malformed optional amounts and dates', () => {
  const normalized = normalizeBillSyncRow({
    Id: 'bill-4',
    BillTotalAmount: 'not-an-amount',
    InvoiceDate: 'not-a-date',
    DueDate: '',
  });

  assert.equal(normalized?.billTotalAmount, null);
  assert.equal(normalized?.invoiceDate, null);
  assert.equal(normalized?.dueDate, null);
});