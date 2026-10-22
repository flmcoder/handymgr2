import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

import { buildTurnEngineCountQuery, TURN_ENGINE_SQL } from './turnEngineQuery.ts';
import { tenantDirectoryPropertyMatchSql } from './sync/tenantDirectoryPolicy.ts';

test('turn engine unions native and occupancy-derived shadow turns', () => {
  assert.match(TURN_ENGINE_SQL, /native_turns\s+as/i);
  assert.match(TURN_ENGINE_SQL, /shadow_turns\s+as/i);
  assert.match(TURN_ENGINE_SQL, /union all/i);
  assert.match(TURN_ENGINE_SQL, /appfolio_tenant_directory/i);
});

test('native turn and same-unit same-date move-out notice produce one canonical turn identity', () => {
  assert.match(TURN_ENGINE_SQL, /'unit:'\s*\|\|\s*coalesce\(d\.property_id, ''\)\s*\|\|\s*':'\s*\|\|\s*d\.unit_id\s*\|\|\s*':'\s*\|\|\s*to_char\(d\.move_out_date at time zone 'UTC', 'YYYY-MM-DD'\)/i);
  assert.match(TURN_ENGINE_SQL, /n\.move_out_date::date\s*=\s*o\.move_out_date::date/i);
});

test('native and shadow turns carry canonical property-group UUIDs through the response', async () => {
  const source = await readFile(new URL('./server.ts', import.meta.url), 'utf8');
  const routeStart = source.indexOf("app.get('/api/local/v2/turns'");
  const routeEnd = source.indexOf("app.get('/api/local/turns'", routeStart);
  const route = source.slice(routeStart, routeEnd);

  assert.match(tenantDirectoryPropertyMatchSql('p', 'd.property_id'), /raw_json ->> 'PropertyId'/);
  assert.match(TURN_ENGINE_SQL, /left join appfolio_properties p on \(\s*p\.id = d\.property_id/i);
  assert.match(TURN_ENGINE_SQL, /left join appfolio_properties p on \(\s*p\.id = td\.property_id/i);
  assert.match(TURN_ENGINE_SQL, /left join appfolio_property_groups pg on pg\.uuid = p\.property_group_id or pg\.id = p\.property_group_id/i);
  assert.equal((TURN_ENGINE_SQL.match(/coalesce\(nullif\(pg\.uuid, ''\), nullif\(p\.property_group_id, ''\)\) as property_group_uuid/gi) || []).length, 2);
  assert.match(TURN_ENGINE_SQL, /fr\.property_group_uuid/i);
  assert.match(route, /property_group_uuid: row\.property_group_uuid \? String\(row\.property_group_uuid\) : null/);
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

test('turn engine property scope is badge-identical (singular id or PropertyGroupIds array)', () => {
  assert.match(TURN_ENGINE_SQL, /p_scope\.property_group_id = ANY/i);
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

test('badge counts aggregate Turn Engine metrics in PostgreSQL instead of transferring full rows', async () => {
  const source = await readFile(new URL('./server.ts', import.meta.url), 'utf8');
  const routeStart = source.indexOf("app.get('/api/local/badge_counts'");
  const routeEnd = source.indexOf("app.get('/api/local/analytics/work-orders'", routeStart);
  const route = source.slice(routeStart, routeEnd);

  const countSql = buildTurnEngineCountQuery();
  assert.match(route, /buildTurnEngineCountQuery\(\)/);
  assert.match(countSql, /count\(\*\) filter \(where lower\(coalesce\(turn_rows\.status, ''\)\) <> 'completed'\)::int as turns/i);
  assert.match(countSql, /from \([\s\S]*\) as turn_rows/i);
  assert.doesNotMatch(countSql, /\nlimit \$\d+/i);
  assert.match(countSql, /\$1::int[\s\S]*\$2::text\[\][\s\S]*\$3::text/i);
  assert.match(route, /turnMetrics\.turns/);
  assert.doesNotMatch(route, /const activeTurnRows\s*=\s*\(turnRows as any\[\]\)\.filter/);
});

test('Express registers authenticated unit-turn tracker synchronization', async () => {
  const source = await readFile(new URL('./server.ts', import.meta.url), 'utf8');

  assert.match(source, /unit_turns_sync:\s*async/);
  assert.match(source, /unit_turns_sync:[\s\S]{0,300}requireProxySession/);
  assert.match(source, /unit_turns_sync:[\s\S]{0,700}upsertTurnTracker/);
});