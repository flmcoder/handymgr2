import assert from 'node:assert/strict';
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