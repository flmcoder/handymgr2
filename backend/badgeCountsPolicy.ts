/**
 * Builds the SQL fragments + response shape for the lightweight nav-badge
 * count endpoint. Kept pure/string-based so it can be unit-tested without a DB.
 */

export const OPEN_WORK_ORDER_STATUS_FILTER = `(
  coalesce(lower(status), '') not like '%completed%'
  and coalesce(lower(status), '') not like '%cancel%'
  and coalesce(lower(status), '') not like '%no need to bill%'
)`;

/** Normalize a badge count to a safe non-negative integer. */
export function toBadgeCount(value: unknown): number {
  const n = Number(value);
  if (!Number.isFinite(n) || n < 0) return 0;
  return Math.floor(n);
}

export type BadgeCounts = {
  work_orders: number;
  turns: number;
  inspections: number;
};

/**
 * Shape the JSON payload returned to the frontend. All counts are clamped to
 * non-negative integers so a bad COUNT row never renders a negative/NaN badge.
 */
export function buildBadgeCountsPayload(input: {
  workOrders?: unknown;
  turns?: unknown;
  inspections?: unknown;
  propertyGroupId?: unknown;
}): BadgeCounts & { ok: true; property_group_id: string; source: string } {
  return {
    ok: true,
    work_orders: toBadgeCount(input.workOrders),
    turns: toBadgeCount(input.turns),
    inspections: toBadgeCount(input.inspections),
    property_group_id: String(input.propertyGroupId || ''),
    source: 'postgres_local',
  };
}
