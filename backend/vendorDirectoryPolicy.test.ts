import assert from 'node:assert/strict';
import test from 'node:test';
import {
  classifyVendorSpend,
  evaluateVendorCompliance,
  parseVendorDirectoryQuery,
} from './vendorDirectoryPolicy.ts';

test('parseVendorDirectoryQuery defaults to a 50-row first page', () => {
  assert.deepEqual(parseVendorDirectoryQuery({}), {
    limit: 50,
    offset: 0,
    page: 1,
    search: '',
    tradeCategory: '',
    propertyId: '',
    sort: 'name',
  });
});

test('parseVendorDirectoryQuery caps page size and derives a safe offset', () => {
  assert.deepEqual(parseVendorDirectoryQuery({ limit: '5000', page: '3' }), {
    limit: 100,
    offset: 200,
    page: 3,
    search: '',
    tradeCategory: '',
    propertyId: '',
    sort: 'name',
  });
});

test('parseVendorDirectoryQuery normalizes filters and rejects unsupported sorting', () => {
  assert.deepEqual(parseVendorDirectoryQuery({
    limit: '75',
    offset: '150',
    search: '  plumbing  ',
    trade_category: '  Plumbing ',
    property_id: ' property-1 ',
    sort: 'drop table vendors',
  }), {
    limit: 75,
    offset: 150,
    page: 3,
    search: 'plumbing',
    tradeCategory: 'Plumbing',
    propertyId: 'property-1',
    sort: 'name',
  });
});

test('parseVendorDirectoryQuery supports trade and recent-work sorting', () => {
  assert.equal(parseVendorDirectoryQuery({ sort: 'trade' }).sort, 'trade');
  assert.equal(parseVendorDirectoryQuery({ sort: 'recent' }).sort, 'recent');
});

test('classifyVendorSpend isolates the two fixed in-house baselines', () => {
  assert.equal(classifyVendorSpend('Fort Lowell Maintenance'), 'in_house_maintenance');
  assert.equal(classifyVendorSpend(' fort lowell realty '), 'in_house_realty');
  assert.equal(classifyVendorSpend('Fort Lowell Maintenance LLC'), 'third_party');
  assert.equal(classifyVendorSpend('Acme Plumbing'), 'third_party');
});

test('vendor is compliant only when both policies are present and unexpired', () => {
  assert.deepEqual(
    evaluateVendorCompliance(
      { liability_ins_expires: '2027-01-31', workers_comp_expires: '2026-12-01' },
      new Date('2026-09-04'),
    ),
    { compliant: true, missing: [], expired: [] },
  );
});

test('missing or unparsable policy dates are non-compliant', () => {
  assert.deepEqual(
    evaluateVendorCompliance(
      { liability_ins_expires: '', workers_comp_expires: 'not-a-date' },
      new Date('2026-09-04'),
    ),
    {
      compliant: false,
      missing: ['liability_ins_expires', 'workers_comp_expires'],
      expired: [],
    },
  );
});

test('a lapsed policy date is reported as expired and non-compliant', () => {
  assert.deepEqual(
    evaluateVendorCompliance(
      { liability_ins_expires: '2026-08-01', workers_comp_expires: '2027-03-15' },
      new Date('2026-09-04'),
    ),
    { compliant: false, missing: [], expired: ['liability_ins_expires'] },
  );
});