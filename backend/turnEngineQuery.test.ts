import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

import { TURN_ENGINE_SQL } from './turnEngineQuery.ts';

test('turn engine unions native and occupancy-derived shadow turns', () => {
  assert.match(TURN_ENGINE_SQL, /native_turns\s+as/i);
  assert.match(TURN_ENGINE_SQL, /shadow_turns\s+as/i);
  assert.match(TURN_ENGINE_SQL, /union all/i);
  assert.match(TURN_ENGINE_SQL, /appfolio_tenant_directory/i);
});

test('work-order sweep is bounded by move-out and next move-in', () => {
  assert.match(TURN_ENGINE_SQL, /wo\.unit_id\s*=\s*tb\.unit_id/i);
  assert.match(TURN_ENGINE_SQL, /wo\.created_at\s*>?=\s*tb\.move_out_date/i);
  assert.match(TURN_ENGINE_SQL, /coalesce\(tb\.next_move_in_date,\s*current_date/i);
});

test('turn engine emits all eleven milestone labels in order', () => {
  const labels = [
    'Move-Out Recorded',
    'Move-Out Inspection',
    'Locks Rekeyed',
    'Estimates Approved',
    'Maintenance / Repair',
    'Paint',
    'Flooring / Carpet',
    'Cleaning / Housekeeping',
    'Appliances',
    'Rent Ready',
    'Marketing Active',
  ];

  let previousIndex = -1;
  for (const label of labels) {
    const index = TURN_ENGINE_SQL.indexOf(label);
    assert.ok(index > previousIndex, `${label} must appear in milestone order`);
    previousIndex = index;
  }
});

test('turn engine includes compliance, financial, and strict completion fields', () => {
  assert.match(TURN_ENGINE_SQL, /in_house_cost/i);
  assert.match(TURN_ENGINE_SQL, /third_party_cost/i);
  assert.match(TURN_ENGINE_SQL, /rogue_wos_detected/i);
  assert.match(TURN_ENGINE_SQL, /is_native_turn/i);
  assert.match(TURN_ENGINE_SQL, /all_work_orders_completed/i);
  assert.match(TURN_ENGINE_SQL, /has_current_resident/i);
  assert.match(TURN_ENGINE_SQL, /strict_completed/i);
});

test('turn engine property scope uses Link + PropertyGroupIds array', () => {
  assert.match(TURN_ENGINE_SQL, /p_scope\.raw_json->>'Link'/i);
  assert.match(TURN_ENGINE_SQL, /p_scope\.raw_json->'PropertyGroupIds' \?\|/i);
});

test('turn completion is vacancy-driven, never gated on a current resident', () => {
  // Regression: gating strict_completed on has_current_resident froze the
  // board at 0 Completed (live turns are vacant). Completion must rest on
  // work evidence / AppFolio turn end / detected move-in instead.
  assert.doesNotMatch(
    TURN_ENGINE_SQL,
    /has_current_resident[\s\S]{0,200}as strict_completed/i,
  );
  assert.match(TURN_ENGINE_SQL, /all_work_orders_completed\)\s*\n\s*or te\.turn_end_date is not null/i);
  assert.match(TURN_ENGINE_SQL, /te\.next_move_in_date is not null and te\.next_move_in_date <= now\(\)/i);
});

test('turn engine exposes occupancy triage for the priority view', () => {
  assert.match(TURN_ENGINE_SQL, /occupied-active/);
  assert.match(TURN_ENGINE_SQL, /vacant-rented/);
  assert.match(TURN_ENGINE_SQL, /vacant-unrented/);
  assert.match(TURN_ENGINE_SQL, /as occupancy_state/i);
  assert.match(TURN_ENGINE_SQL, /as ready_to_close/i);
  assert.match(TURN_ENGINE_SQL, /as is_priority/i);
});

test('turn scope matches badge identity (singular group id or PropertyGroupIds array)', () => {
  assert.match(TURN_ENGINE_SQL, /p_scope\.property_group_id = ANY\(\$3::text\[\]\)/i);
  assert.match(TURN_ENGINE_SQL, /p_scope\.raw_json->'PropertyGroupIds' \?\|\s*\(\$3::text\[\]\)/);
});

test('Express registers authenticated unit-turn tracker synchronization', async () => {
  const source = await readFile(new URL('./server.ts', import.meta.url), 'utf8');

  assert.match(source, /unit_turns_sync:\s*async/);
  assert.match(source, /unit_turns_sync:[\s\S]{0,300}requireProxySession/);
  assert.match(source, /unit_turns_sync:[\s\S]{0,700}upsertTurnTracker/);
});