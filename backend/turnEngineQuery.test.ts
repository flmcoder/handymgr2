import assert from 'node:assert/strict';
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