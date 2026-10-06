import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import { tenantTransactionDisplayRow } from './occupancyReportPolicy.js';

test('tenant transaction model accepts the documented property_name_address and unit_name fields', () => {
  assert.deepEqual(tenantTransactionDisplayRow({
    occupancy_name: 'Synthetic Occupancy',
    property_name_address: '123 Example St, Tucson, AZ',
    unit_name: 'Unit 4',
    rent_charges: '1000.00',
    cash_payments: '950.00',
    ending_balance: '50.00',
  }), {
    tenant: 'Synthetic Occupancy',
    propertyUnit: '123 Example St, Tucson, AZ / Unit 4',
    charges: '1000.00',
    payments: '950.00',
    balance: '50.00',
    lastActivity: '',
  });
});

test('tenant transaction model preserves legacy aliases and blank-safe fallbacks', () => {
  assert.deepEqual(tenantTransactionDisplayRow({
    tenant_name: 'Synthetic Tenant',
    property: 'Legacy Property',
    unit: 'Legacy Unit',
    total_charges: 20,
    total_payments: 5,
    balance: 15,
    last_activity_date: '2026-10-01',
  }), {
    tenant: 'Synthetic Tenant',
    propertyUnit: 'Legacy Property / Legacy Unit',
    charges: 20,
    payments: 5,
    balance: 15,
    lastActivity: '2026-10-01',
  });

  assert.equal(tenantTransactionDisplayRow({ property_name: '  ', unit_name: null }).propertyUnit, '');
});

test('tenant transaction renderer delegates to the tested display model', async () => {
  const source = await readFile(new URL('./app.js', import.meta.url), 'utf8');
  const reportMap = source.slice(source.indexOf("'tenant-transactions':"), source.indexOf("'tenant-directory':"));
  assert.match(reportMap, /tenantTransactionDisplayRow\(r\)/);
});