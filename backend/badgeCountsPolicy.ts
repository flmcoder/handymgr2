/**
 * Builds the SQL fragments + response shape for the lightweight nav-badge
 * count endpoint. Kept pure/string-based so it can be unit-tested without a DB.
 */

export const OPEN_WORK_ORDER_STATUS_FILTER = `(
  coalesce(lower(status), '') not like '%completed%'
  and coalesce(lower(status), '') not like '%cancel%'
  and coalesce(lower(status), '') not like '%no need to bill%'
)`;

/**
 * The effective created timestamp of a work order. Prefers the normalized
 * `created_at` column but falls back to the raw AppFolio payload keys
 * (`CreatedAt`, snake-case variants) and finally `updated_at`, so rows synced
 * before `created_at` was populated still report accurate aging. Unqualified on
 * purpose — the nav-badge aggregate has no table alias. Safe to extend; each
 * raw fallback is guarded by a leading-date regex so a malformed value can
 * never break the whole aggregate.
 */
export const WORK_ORDER_CREATED_AT_EXPR = `coalesce(
  created_at,
  case
    when (raw_json ->> 'CreatedAt') ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}'
      then (raw_json ->> 'CreatedAt')::timestamptz
    else null
  end,
  case
    when (raw_json ->> 'created_at') ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}'
      then (raw_json ->> 'created_at')::timestamptz
    else null
  end,
  case
    when (raw_json ->> 'created_date') ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}'
      then (raw_json ->> 'created_date')::timestamptz
    else null
  end,
  updated_at
)`;

export const ACTIVE_TURN_STATUS_FILTER = `(
  coalesce(lower(t.status), '') not like '%completed%'
  and coalesce(lower(t.status), '') not like '%closed%'
  and coalesce(t.updated_at, t.created_at, now()) >= now() - interval '90 days'
)`;

export const ACTIVE_INSPECTION_RESIDENT_FILTER = `(
  lower(coalesce(occ.status, '')) = 'current'
  and coalesce(occ.tenant_name, '') <> ''
  and lower(coalesce(occ.tenant_type, '')) = 'financially responsible'
  and coalesce(occ.property_id, '') <> ''
  and coalesce(occ.occupancy_id, '') <> ''
  and (coalesce(occ.unit_id, '') <> '' or coalesce(occ.occupancy_id, '') <> '')
  and occ.move_in_date is not null
  and occ.move_in_date <= current_date
  and (occ.lease_to is null or occ.lease_to >= current_date)
  and (occ.move_out_date is null or occ.move_out_date >= current_date)
)`;

export const PROPERTY_IDENTITY_MATCH_FILTER = `(
  p.id = source_property_id
  or coalesce(
    p.raw_json ->> 'PropertyId',
    p.raw_json ->> 'property_id',
    p.raw_json ->> 'Id',
    p.raw_json ->> 'id'
  ) = source_property_id
)`;

export function propertyIdentityMatch(sourceExpression: string, propertyAlias = 'p'): string {
  return PROPERTY_IDENTITY_MATCH_FILTER
    .replaceAll('p.', `${propertyAlias}.`)
    .replaceAll('source_property_id', sourceExpression);
}

/** Normalize a badge count to a safe non-negative integer. */
export function toBadgeCount(value: unknown): number {
  const n = Number(value);
  if (!Number.isFinite(n) || n < 0) return 0;
  return Math.floor(n);
}

export type BadgeCounts = {
  work_orders: number;
  urgent_work_orders: number;
  work_order_aging: {
    age_0_7: number;
    age_8_30: number;
    age_31_60: number;
    age_61_plus: number;
    age_unknown: number;
  };
  turns: number;
  upcoming_turns: number;
  inspections: number;
};

/**
 * Shape the JSON payload returned to the frontend. All counts are clamped to
 * non-negative integers so a bad COUNT row never renders a negative/NaN badge.
 */
export function buildBadgeCountsPayload(input: {
  workOrders?: unknown;
  urgentWorkOrders?: unknown;
  workOrdersAge0To7?: unknown;
  workOrdersAge8To30?: unknown;
  workOrdersAge31To60?: unknown;
  workOrdersAge61Plus?: unknown;
  workOrdersAgeUnknown?: unknown;
  turns?: unknown;
  upcomingTurns?: unknown;
  inspections?: unknown;
  propertyGroupId?: unknown;
}): BadgeCounts & { ok: true; property_group_id: string; source: string } {
  return {
    ok: true,
    work_orders: toBadgeCount(input.workOrders),
    urgent_work_orders: toBadgeCount(input.urgentWorkOrders),
    work_order_aging: {
      age_0_7: toBadgeCount(input.workOrdersAge0To7),
      age_8_30: toBadgeCount(input.workOrdersAge8To30),
      age_31_60: toBadgeCount(input.workOrdersAge31To60),
      age_61_plus: toBadgeCount(input.workOrdersAge61Plus),
      age_unknown: toBadgeCount(input.workOrdersAgeUnknown),
    },
    turns: toBadgeCount(input.turns),
    upcoming_turns: toBadgeCount(input.upcomingTurns),
    inspections: toBadgeCount(input.inspections),
    property_group_id: String(input.propertyGroupId || ''),
    source: 'postgres_local',
  };
}
