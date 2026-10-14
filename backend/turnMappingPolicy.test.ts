import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

import { canonicalTurnKey, mergeTurnSources, normalizeUnitTurnDetail, turnOverlapKey } from './turnMappingPolicy.ts';

test('normalizes the documented unit_turn_detail snake-case source fields', () => {
  assert.deepEqual(normalizeUnitTurnDetail({
    property_id: 'prop-1',
    property: 'Sample Property',
    unit_id: 'unit-7',
    unit: 'Unit 7',
    unit_turn_id: 'turn-88',
    move_out_date: '2026-08-01T00:00:00Z',
    turn_end_date: '2026-08-09T00:00:00Z',
    expected_move_in_date: '2026-08-15T00:00:00Z',
    unit_turn_status: 'In Progress',
  }), {
    turnId: 'turn-88',
    propertyId: 'prop-1',
    propertyName: 'Sample Property',
    unitId: 'unit-7',
    unitName: 'Unit 7',
    moveOutDate: '2026-08-01T00:00:00Z',
    turnEndDate: '2026-08-09T00:00:00Z',
    expectedMoveInDate: '2026-08-15T00:00:00Z',
    status: 'In Progress',
    rawJson: {
      property_id: 'prop-1',
      property: 'Sample Property',
      unit_id: 'unit-7',
      unit: 'Unit 7',
      unit_turn_id: 'turn-88',
      move_out_date: '2026-08-01T00:00:00Z',
      turn_end_date: '2026-08-09T00:00:00Z',
      expected_move_in_date: '2026-08-15T00:00:00Z',
      unit_turn_status: 'In Progress',
    },
  });
});

test('turn upsert uses the normalized source-field mapper and stores its original row', async () => {
  const source = await readFile(new URL('./sync/repositories.ts', import.meta.url), 'utf8');
  const start = source.indexOf('export async function upsertUnitTurnDetails(');
  const end = source.indexOf('// ── Unit Vacancies', start);
  const upsert = source.slice(start, end);
  assert.match(upsert, /normalizeUnitTurnDetail\(row\)/);
  assert.match(upsert, /rawJson:\s*row/);
});

test('frontend turn pipeline uses date-normalized shared-source merge identity', async () => {
  const source = await readFile(new URL('../src/app.js', import.meta.url), 'utf8');
  const pipelineStart = source.indexOf('function buildTurnPipeline()');
  const pipelineEnd = source.indexOf('async function fetchPropertyStats()', pipelineStart);
  const pipeline = source.slice(pipelineStart, pipelineEnd);
  assert.match(pipeline, /moveOutDateKey/);
  assert.match(pipeline, /sourceKeys/);
  assert.match(pipeline, /workOrders/);
});

test('normalizes title-case and alternate IDs/dates without dropping mapping fields', () => {
  const row = normalizeUnitTurnDetail({
    PropertyId: 'prop-2',
    PropertyName: 'Another Property',
    UnitId: 'unit-3',
    UnitName: 'Apt 3',
    UnitTurnId: 'turn-99',
    MoveOutDate: '2026-07-02',
    TurnEndDate: '2026-07-10',
    ExpectedMoveInDate: '2026-07-20',
    Status: 'Ready',
  });
  assert.equal(row.turnId, 'turn-99');
  assert.equal(row.propertyId, 'prop-2');
  assert.equal(row.propertyName, 'Another Property');
  assert.equal(row.unitId, 'unit-3');
  assert.equal(row.unitName, 'Apt 3');
  assert.equal(row.moveOutDate, '2026-07-02');
  assert.equal(row.turnEndDate, '2026-07-10');
  assert.equal(row.expectedMoveInDate, '2026-07-20');
  assert.equal(row.status, 'Ready');
});

test('falls back to stable property/unit/move-out identity when AppFolio turn id is absent', () => {
  const row = normalizeUnitTurnDetail({ property_id: 'p', unit_id: 'u', move_out_date: '2026-07-02T15:00:00Z' });
  assert.equal(row.turnId, 'p|u|2026-07-02|');
});

test('same-unit move-out notice and native turn share an overlap key', () => {
  const native = { propertyId: 'p1', unitId: 'u7', unitTurnId: 'native-123', moveOutDate: '2026-08-01T00:00:00Z' };
  const moveOut = { propertyId: 'p1', unitId: 'u7', moveOutDate: '2026-08-01T12:00:00Z' };
  assert.equal(turnOverlapKey(native), turnOverlapKey(moveOut));
  assert.notEqual(canonicalTurnKey(native), canonicalTurnKey(moveOut));
});

test('open make-ready work order attaches to overlapping native turn without creating another turn', () => {
  const native = { propertyId: 'p1', unitId: 'u7', unitTurnId: 'native-123', moveOutDate: '2026-08-01T00:00:00Z' };
  const moveOutNotice = { propertyId: 'p1', unitId: 'u7', moveOutDate: '2026-08-01T00:00:00Z' };
  const openMakeReady = { propertyId: 'p1', unitId: 'u7', moveOutDate: '2026-08-01T00:00:00Z', woNumber: '1042', status: 'Assigned' };

  const merged = mergeTurnSources([
    { source: 'native', ...native, propertyName: 'Sample Property', unitName: 'Unit 7' },
    { source: 'moveout', ...moveOutNotice },
  ]);
  const turn = merged.find((entry) => entry.overlapKey === turnOverlapKey(native));
  assert.equal(merged.length, 1);
  assert.equal(turn?.sources.length, 2);
  assert.equal(turn?.propertyName, 'Sample Property');
  assert.equal(turn?.unitName, 'Unit 7');

  const attached = turn?.workOrders.concat(openMakeReady);
  assert.equal(attached?.length, 1);
  assert.equal(attached?.[0].woNumber, '1042');
});