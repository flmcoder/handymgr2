import assert from 'node:assert/strict';
import test from 'node:test';
import { normalizeBillSyncRow } from './sync/billSyncPolicy.ts';

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