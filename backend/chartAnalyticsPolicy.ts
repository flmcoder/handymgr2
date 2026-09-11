/**
 * Server-side aggregates backing the dashboard charts. These reuse the exact
 * predicates from badge_counts and the grids (open-work-order filter, active
 * resident filter, missing-inspection definition, numeric-Link property
 * resolution, PropertyGroupIds scope) so charts, badges and lists reconcile
 * by construction instead of each client re-aggregating truncated row pages.
 */
import {
  ACTIVE_INSPECTION_RESIDENT_FILTER,
  OPEN_WORK_ORDER_STATUS_FILTER,
  WORK_ORDER_CREATED_AT_EXPR,
} from './badgeCountsPolicy.ts';

export const WO_OWNER_EXPR = `coalesce(nullif(wo.assigned_user_name, ''), nullif(wo.vendor_name, ''), nullif(wo.raw_json->>'assigned_user_name', ''), nullif(wo.raw_json->>'vendor_name', ''), 'Unassigned')`;

export const WO_TYPE_EXPR = `coalesce(nullif(wo.category, ''), nullif(wo.raw_json->>'category', ''), nullif(wo.raw_json->>'type', ''), 'Unspecified')`;

export const WO_PRIORITY_EXPR = `coalesce(nullif(wo.priority, ''), 'Unspecified')`;

export const WO_STATUS_EXPR = `coalesce(nullif(wo.status, ''), 'Unknown')`;

export const WO_PROPERTY_EXPR = `coalesce(nullif(p.name, ''), nullif(wo.raw_json->>'property_name', ''), 'Unknown')`;

const WO_BASE_WHERE = (scoped: boolean): string => scoped
  ? `${OPEN_WORK_ORDER_STATUS_FILTER} and wo.property_group_id = ANY($1::text[])`
  : OPEN_WORK_ORDER_STATUS_FILTER;

const WO_AGE = WORK_ORDER_CREATED_AT_EXPR;

export function buildWorkOrdersAnalyticsQuery(scopeIds: string[]): { sql: string; params: string[][] } {
  const scoped = scopeIds.length > 0;
  const where = WO_BASE_WHERE(scoped);
  const params: string[][] = scoped ? [scopeIds] : [];
  const sql = `
    select
      (select count(*)::int from appfolio_work_orders wo where ${where}) as total,
      (select count(*)::int from appfolio_work_orders wo where ${where} and lower(coalesce(wo.priority, '')) in ('urgent', 'emergency', 'critical')) as urgent,
      (select count(*)::int from appfolio_work_orders wo where ${where} and ${WO_AGE} is not null and current_date - (${WO_AGE})::date between 0 and 7) as age_0_7,
      (select count(*)::int from appfolio_work_orders wo where ${where} and ${WO_AGE} is not null and current_date - (${WO_AGE})::date between 8 and 30) as age_8_30,
      (select count(*)::int from appfolio_work_orders wo where ${where} and ${WO_AGE} is not null and current_date - (${WO_AGE})::date between 31 and 60) as age_31_60,
      (select count(*)::int from appfolio_work_orders wo where ${where} and ${WO_AGE} is not null and current_date - (${WO_AGE})::date >= 61) as age_61_plus,
      (select count(*)::int from appfolio_work_orders wo where ${where} and ${WO_AGE} is null) as age_unknown,
      (select coalesce(jsonb_object_agg(t.label, t.cnt), '{}'::jsonb) from (select ${WO_STATUS_EXPR} as label, count(*)::int as cnt from appfolio_work_orders wo where ${where} group by 1 order by 2 desc limit 24) t) as by_status,
      (select coalesce(jsonb_object_agg(t.label, t.cnt), '{}'::jsonb) from (select ${WO_TYPE_EXPR} as label, count(*)::int as cnt from appfolio_work_orders wo where ${where} group by 1 order by 2 desc limit 24) t) as by_type,
      (select coalesce(jsonb_object_agg(t.label, t.cnt), '{}'::jsonb) from (select ${WO_OWNER_EXPR} as label, count(*)::int as cnt from appfolio_work_orders wo where ${where} group by 1 order by 2 desc limit 24) t) as by_owner,
      (select coalesce(jsonb_object_agg(t.label, t.cnt), '{}'::jsonb) from (select ${WO_PRIORITY_EXPR} as label, count(*)::int as cnt from appfolio_work_orders wo where ${where} group by 1 order by 2 desc limit 12) t) as by_priority,
      (select coalesce(jsonb_object_agg(t.label, t.avg_age), '{}'::jsonb) from (select ${WO_OWNER_EXPR} as label, round(avg(current_date - (${WO_AGE})::date))::int as avg_age from appfolio_work_orders wo where ${where} and ${WO_AGE} is not null group by 1 order by 2 desc limit 12) t) as avg_age_by_owner,
      (select coalesce(jsonb_object_agg(t.label, t.cnt), '{}'::jsonb) from (select ${WO_PROPERTY_EXPR} as label, count(*)::int as cnt from appfolio_work_orders wo left join appfolio_properties p on p.id = wo.property_id where ${where} group by 1 order by 2 desc limit 24) t) as by_property,
      (select coalesce(jsonb_agg(t), '[]'::jsonb) from (select ${WO_STATUS_EXPR} as status, ${WO_OWNER_EXPR} as owner, count(*)::int as value from appfolio_work_orders wo where ${where} group by 1, 2 order by 3 desc limit 60) t) as by_status_owner
  `;
  return { sql, params };
}

const OCC_SCOPE_JOIN = (scope: boolean): string => scope
  ? `join appfolio_properties p on p.raw_json->>'Link' = 'https://flraz.appfolio.com/properties/' || occ.property_id
     and p.raw_json->'PropertyGroupIds' ?| ($1::text[])`
  : '';

const INSP_LATERAL = `
  left join lateral (
    select i0.last_inspection_date
    from appfolio_unit_inspections i0
    where (
      (coalesce(occ.occupancy_id, '') <> '' and i0.occupancy_id = occ.occupancy_id)
      or (coalesce(occ.occupancy_id, '') = '' and i0.unit_id = occ.unit_id)
    )
    order by coalesce(i0.last_inspection_date, i0.last_updated_at, i0.cached_at) desc nulls last
    limit 1
  ) i on true
`;

const TURN_LINK_JOIN = `
  left join (
    select distinct property_id, unit_id
    from unit_turn_tracker
    where coalesce(lower(status), '') not like '%completed%'
      and coalesce(lower(status), '') not like '%closed%'
  ) tl on tl.property_id = occ.property_id and tl.unit_id = occ.unit_id
`;

const INSP_BASE = (scope: boolean): string => `
  select
    coalesce(occ.property_id, '') as property_id,
    coalesce(occ.property_name, '') as property_name,
    coalesce(occ.raw_json->>'property_address', occ.raw_json->>'property', '') as property_address,
    i.last_inspection_date,
    occ.move_in_date,
    (i.last_inspection_date is null or i.last_inspection_date::date < occ.move_in_date) as missing,
    case
      when occ.move_in_date is not null
           and occ.move_in_date <= current_date
           and (i.last_inspection_date is null or i.last_inspection_date::date < occ.move_in_date)
        then occ.move_in_date::date
      else i.last_inspection_date::date
    end as anchor_date,
    tl.property_id is not null as turn_linked
  from appfolio_tenant_directory occ
  ${OCC_SCOPE_JOIN(scope)}
  ${INSP_LATERAL}
  ${TURN_LINK_JOIN}
  where ${ACTIVE_INSPECTION_RESIDENT_FILTER}
`;

export function buildInspectionsAnalyticsQuery(
  scopeIds: string[],
  overdueDays = 365,
  dueSoonDays = 270,
): { sql: string; params: string[][] } {
  const od = Math.max(1, Math.floor(Number(overdueDays) || 365));
  const ds = Math.max(1, Math.floor(Number(dueSoonDays) || 270));
  const scoped = scopeIds.length > 0;
  const params: string[][] = scoped ? [scopeIds] : [];
  const daysSince = `greatest(coalesce((current_date - b.anchor_date), 0), 0)`;
  const isOverdue = `(b.missing or b.anchor_date is null or (current_date - b.anchor_date) > ${od})`;
  const sql = `
    with base as (${INSP_BASE(scoped)})
    select
      (select count(*)::int from base b) as total_active,
      (select count(*)::int from base b where b.missing) as total_missing,
      (select count(*)::int from base b where ${isOverdue}) as overdue,
      (select count(*)::int from base b where not ${isOverdue} and (current_date - b.anchor_date) > ${ds}) as due_soon,
      (select count(*)::int from base b where ${daysSince} between 0 and 30) as age_0_30,
      (select count(*)::int from base b where ${daysSince} between 31 and 90) as age_31_90,
      (select count(*)::int from base b where ${daysSince} between 91 and 180) as age_91_180,
      (select count(*)::int from base b where ${daysSince} >= 181) as age_181_plus,
      (select count(*)::int from base b where b.turn_linked) as linked,
      (select count(*)::int from base b where not b.turn_linked) as not_linked,
      (select coalesce(jsonb_agg(t), '[]'::jsonb) from (select b.property_id, max(b.property_name) as property_name, max(b.property_address) as property_address, count(*) filter (where b.missing)::int as missing from base b group by 1 having count(*) filter (where b.missing) > 0 order by 4 desc) t) as by_property
  `;
  return { sql, params };
}

export type CountBucket = { label: string; value: number };

/** Shape a jsonb {label: count} map into top-N descending buckets. */
export function toCountBuckets(value: unknown, topN = 8): CountBucket[] {
  const obj = (value && typeof value === 'object' ? value : {}) as Record<string, unknown>;
  return Object.entries(obj)
    .map(([label, raw]) => ({ label: String(label), value: Number(raw) || 0 }))
    .sort((a, b) => b.value - a.value)
    .slice(0, Math.max(1, topN));
}
