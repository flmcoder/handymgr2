import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import { inspectionMissingIndicators, parseInspectionPage } from './inspectionPolicy.ts';

test('inspection pages are bounded and offset-aware', () => {
  assert.deepEqual(parseInspectionPage({ limit: '5000', page: '3' }), { limit: 100, offset: 200, page: 3 });
});

test('active lease without a post-move-in inspection is explicitly flagged', () => {
  assert.deepEqual(inspectionMissingIndicators({
    move_in_date: '2026-01-01',
    move_out_date: '2026-12-31',
    last_inspection_date: '2025-12-01',
    move_out_inspection_date: null,
  }, new Date('2026-09-04')), {
    missing_move_in_inspection: true,
    missing_move_out_inspection: true,
  });
});

test('future move-in and completed move-out inspection are not flagged', () => {
  assert.deepEqual(inspectionMissingIndicators({
    move_in_date: '2026-10-01',
    move_out_date: '2026-12-31',
    last_inspection_date: '2026-10-02',
    move_out_inspection_date: '2026-12-31',
  }, new Date('2026-09-04')), {
    missing_move_in_inspection: false,
    missing_move_out_inspection: false,
  });
});

test('inspection grid uses active-only residency and returns association fields', async () => {
  const source = await readFile(new URL('./server.ts', import.meta.url), 'utf8');

  assert.match(source, /activeOnly[\s\S]{0,5000}ACTIVE_INSPECTION_RESIDENT_FILTER/);
  assert.match(source, /lease_to/);
  assert.match(source, /tenant_type/);
  assert.match(source, /has_valid_unit_association/);
  assert.match(source, /raw_json\s*->>\s*'Link'\s*=\s*'https:\/\/flraz\.appfolio\.com\/properties\/'\s*\|\|\s*occ\.property_id/);
  assert.match(source, /PropertyGroupIds'\s*@>/);
});