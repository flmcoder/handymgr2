import assert from 'node:assert/strict';
import test from 'node:test';

import { normalizeTenantDirectoryRow, tenantDirectoryPropertyMatchSql } from './tenantDirectoryPolicy.ts';

test('normalizes the requested Tenant Directory property and unit aliases', () => {
  assert.deepEqual(normalizeTenantDirectoryRow({
    property: 'AppFolio Property Name',
    unit: 'Unit 12',
    property_id: 418,
    tenant: 'Tenant Name',
    occupancy_id: 992,
    move_in: '2025-01-01',
  }), {
    propertyId: '418',
    propertyName: 'AppFolio Property Name',
    unitId: null,
    unitName: 'Unit 12',
    tenantName: 'Tenant Name',
    occupancyId: '992',
    moveIn: '2025-01-01',
    moveOut: undefined,
  });
});

test('prefers explicit IDs and normalized names over display aliases', () => {
  assert.deepEqual(normalizeTenantDirectoryRow({
    property_id: 'canonical-property',
    property_name: 'Normalized property',
    property: 'Display property',
    unit_id: 'canonical-unit',
    unit_name: 'Normalized unit',
    unit: 'Display unit',
  }), {
    propertyId: 'canonical-property',
    propertyName: 'Normalized property',
    unitId: 'canonical-unit',
    unitName: 'Normalized unit',
    tenantName: null,
    occupancyId: null,
    moveIn: undefined,
    moveOut: undefined,
  });
});

test('does not fabricate unit IDs from display labels and ignores blank aliases', () => {
  assert.deepEqual(normalizeTenantDirectoryRow({ property: ' ', unit: ' Apt 3 ' }), {
    propertyId: null,
    propertyName: null,
    unitId: null,
    unitName: 'Apt 3',
    tenantName: null,
    occupancyId: null,
    moveIn: undefined,
    moveOut: undefined,
  });
});

test('property identity join matches UUID, raw AppFolio ID, and Link-derived ID', () => {
  const sql = tenantDirectoryPropertyMatchSql();
  assert.match(sql, /p\.id = t\.property_id/);
  assert.match(sql, /raw_json ->> 'PropertyId'/);
  assert.match(sql, /regexp_replace[\s\S]*raw_json ->> 'Link'/);
});