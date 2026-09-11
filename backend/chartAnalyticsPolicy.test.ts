import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildInspectionsAnalyticsQuery,
  buildWorkOrdersAnalyticsQuery,
  toCountBuckets,
} from './chartAnalyticsPolicy.ts';

test('work-orders analytics uses the badge open filter and aging buckets', () => {
  const { sql, params } = buildWorkOrdersAnalyticsQuery(null);
  assert.equal(params.length, 0);
  assert.match(sql, /not like '%completed%'/);
  assert.match(sql, /by_status/);
  assert.match(sql, /by_type/);
  assert.match(sql, /by_owner/);
  assert.match(sql, /by_priority/);
  assert.match(sql, /avg_age_by_owner/);
  assert.match(sql, /by_property/);
  assert.match(sql, /by_status_owner/);
  assert.match(sql, /age_0_7/);
  assert.match(sql, /age_61_plus/);
});

test('work-orders analytics scopes by property group with one bind param', () => {
  const scope = 'bee73529-9eca-11ee-8b51-02167481f3bc';
  const { sql, params } = buildWorkOrdersAnalyticsQuery(scope);
  assert.deepEqual(params, [scope]);
  assert.match(sql, /wo\.property_group_id = \$1/);
});

test('inspections analytics encodes the missing move-in definition', () => {
  const { sql, params } = buildInspectionsAnalyticsQuery(null);
  assert.equal(params.length, 0);
  assert.match(sql, /ACTIVE_INSPECTION_RESIDENT_FILTER|current.*financially responsible/s);
  assert.match(sql, /last_inspection_date is null or i\.last_inspection_date::date < occ\.move_in_date/);
  assert.match(sql, /total_missing/);
  assert.match(sql, /by_property/);
});

test('inspections analytics scopes through the numeric link and group array', () => {
  const scope = 'bee73529-9eca-11ee-8b51-02167481f3bc';
  const { sql, params } = buildInspectionsAnalyticsQuery(scope);
  assert.deepEqual(params, [scope]);
  assert.match(sql, /raw_json->>'Link'/);
  assert.match(sql, /PropertyGroupIds'\s*@>\s*jsonb_build_array\(\$1::text\)/);
});

test('toCountBuckets sorts descending and caps at topN', () => {
  assert.deepEqual(toCountBuckets({ b: 1, a: 5, c: 3 }, 2), [
    { label: 'a', value: 5 },
    { label: 'c', value: 3 },
  ]);
  assert.deepEqual(toCountBuckets(null), []);
});
