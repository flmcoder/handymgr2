import { OPEN_WORK_ORDER_STATUS_FILTER, buildWorkOrderPropertyGroupScopeSql } from './badgeCountsPolicy.ts';

export function buildActiveWorkOrdersUrl(
  baseUrl: unknown,
  propertyGroupUuid: unknown = '',
  requestedLimit: unknown = 5_000,
  requestedOffset: unknown = 0,
): string {
  const base = String(baseUrl || '').trim().replace(/\/+$/, '');
  const parsedLimit = Number.parseInt(String(requestedLimit || ''), 10);
  const limit = Number.isFinite(parsedLimit) ? Math.max(1, Math.min(20_000, parsedLimit)) : 5_000;
  const parsedOffset = Number.parseInt(String(requestedOffset || ''), 10);
  const offset = Number.isFinite(parsedOffset) ? Math.max(0, Math.min(500_000, parsedOffset)) : 0;
  const params = new URLSearchParams({ limit: String(limit) });
  if (offset > 0) params.set('offset', String(offset));
  const scope = String(propertyGroupUuid || '').trim();
  if (scope) params.set('property_group_id', scope);
  return `${base}/api/local/work_orders?${params.toString()}`;
}

export function buildWorkOrderPagination(
  totalValue: unknown,
  limitValue: unknown,
  offsetValue: unknown,
  pageCountValue: unknown,
): { total: number; limit: number; offset: number; has_next: boolean; has_previous: boolean } {
  const total = Math.max(0, Number.parseInt(String(totalValue || ''), 10) || 0);
  const limit = Math.max(1, Number.parseInt(String(limitValue || ''), 10) || 100);
  const offset = Math.max(0, Number.parseInt(String(offsetValue || ''), 10) || 0);
  const pageCount = Math.max(0, Number.parseInt(String(pageCountValue || ''), 10) || 0);
  return {
    total,
    limit,
    offset,
    has_next: offset + pageCount < total,
    has_previous: offset > 0,
  };
}

export function resolveWorkOrderLookupSearch(value: unknown, minimumLength = 3, maximumLength = 180): string {
  const search = String(value ?? '').trim().slice(0, Math.max(1, maximumLength));
  return search.length >= Math.max(1, minimumLength) ? search : '';
}

export function resolveWorkOrderStatusScope(value: unknown): 'active' | 'inactive' {
  return String(value || '').trim().toLowerCase() === 'inactive' ? 'inactive' : 'active';
}

/**
 * Cross-portfolio PM lookup is intentionally narrow: only a work-order
 * number/reference may escape the normal property-group scope. Free-text
 * searches remain scoped and cannot enumerate another portfolio.
 */
export function resolveExactWorkOrderReference(value: unknown): string {
  const raw = String(value ?? '').trim().replace(/^#/, '');
  if (!raw) return '';
  if (/^(?:wo[-_\s]*)?\d+$/i.test(raw)) return raw;
  return '';
}

export function resolveWorkOrderHistoryDays(value: unknown, maximumDays = 3_650): number | null {
  if (value === undefined || value === null || String(value).trim() === '') return null;
  const parsed = Number.parseInt(String(value), 10);
  if (!Number.isFinite(parsed)) return maximumDays;
  return Math.max(1, Math.min(maximumDays, parsed));
}

const INACTIVE_STATUS_SQL = `(coalesce(lower(status), '') like '%completed%' or coalesce(lower(status), '') like '%cancel%' or coalesce(lower(status), '') like '%no need to bill%')`;

/** SQL predicate for the active/inactive work-order tabs; `status` is qualified with `statusColumn`. */
export function buildWorkOrderStatusScopeSql(
  statusScope: 'active' | 'inactive',
  statusColumn = 'status',
): string {
  const base = statusScope === 'inactive' ? INACTIVE_STATUS_SQL : OPEN_WORK_ORDER_STATUS_FILTER;
  return base.replace(/\bstatus\b/g, statusColumn);
}

export type LocalWorkOrdersQueryOptions = {
  scopeIds: readonly string[];
  statusScope: 'active' | 'inactive';
  statusFilter: string;
  limit: number;
  offset: number;
  includeWorkOrderUuid?: boolean;
};

export type BuiltSqlQuery = { sql: string; params: unknown[] };

function firstRawValue(raw: Record<string, unknown>, keys: string[]): unknown {
  for (const key of keys) {
    const value = raw[key];
    if (value !== undefined && value !== null && String(value).trim() !== '') return value;
  }
  return null;
}

function toIsoDate(value: unknown): string {
  if (value === undefined || value === null || value === '') return '';
  const date = value instanceof Date ? value : new Date(String(value));
  return Number.isNaN(date.getTime()) ? '' : date.toISOString();
}

/** Preserve stored work-order lifecycle timestamps, using raw AppFolio dates only when a column is null. */
export function normalizeWorkOrderDateFields(row: Record<string, any>): {
  completed_on: string;
  work_completed_on: string;
  canceled_on: string;
} {
  const raw = row.raw_json && typeof row.raw_json === 'object' ? row.raw_json as Record<string, unknown> : {};
  const value = (column: string, rawKeys: string[]): string => toIsoDate(
    row[column] !== undefined && row[column] !== null ? row[column] : firstRawValue(raw, rawKeys),
  );
  return {
    completed_on: value('completed_on', ['completed_on', 'CompletedOn', 'CompletedDate', 'completed_date']),
    work_completed_on: value('work_completed_on', ['work_completed_on', 'WorkCompletedOn', 'WorkCompletedDate', 'work_completed_date']),
    canceled_on: value('canceled_on', ['canceled_on', 'CanceledOn', 'CancelledOn', 'CanceledDate', 'canceled_date']),
  };
}

function buildLocalWorkOrdersWhere(
  options: Pick<LocalWorkOrdersQueryOptions, 'scopeIds' | 'statusScope' | 'statusFilter'>,
): { where: string; params: unknown[] } {
  const clauses: string[] = [];
  const params: unknown[] = [];
  if (options.scopeIds.length) {
    params.push([...options.scopeIds]);
    clauses.push(buildWorkOrderPropertyGroupScopeSql('wo', `$${params.length}`));
  }
  clauses.push(buildWorkOrderStatusScopeSql(options.statusScope, 'wo.status'));
  if (options.statusFilter) {
    params.push(options.statusFilter);
    clauses.push(`coalesce(wo.status, '') = $${params.length}::text`);
  }
  return { where: clauses.join('\n and '), params };
}

/**
 * Row-fetch query for /api/local/work_orders. Built from the same WHERE as the
 * count query so the grid and the total can never disagree. Every value is a
 * bound parameter; SQL fragments are concatenated, never bound.
 */
export function buildLocalWorkOrdersQuery(options: LocalWorkOrdersQueryOptions): BuiltSqlQuery {
  const { where, params } = buildLocalWorkOrdersWhere(options);
  params.push(options.limit, options.offset);
  const limitParam = `$${params.length - 1}`;
  const offsetParam = `$${params.length}`;
  const uuidColumn = options.includeWorkOrderUuid === false ? 'null::text as work_order_uuid' : 'wo.work_order_uuid';
  return {
    sql: `select wo.id, ${uuidColumn}, wo.wo_number, wo.property_id, wo.unit_id, wo.property_group_id, wo.description,
            wo.category, wo.priority, wo.status, wo.assigned_user_id, wo.assigned_user_name,
            wo.vendor_id, wo.vendor_name, wo.estimated_amount, wo.total_cost,
            wo.completed_on, wo.work_completed_on, wo.canceled_on,
            wo.created_at, wo.updated_at, wo.raw_json
          from appfolio_work_orders wo
          where ${where}
          order by coalesce(wo.updated_at, wo.created_at) desc nulls last, wo.id
          limit ${limitParam} offset ${offsetParam}`,
    params,
  };
}

export function buildLocalWorkOrdersCountQuery(
  options: Pick<LocalWorkOrdersQueryOptions, 'scopeIds' | 'statusScope' | 'statusFilter'>,
): BuiltSqlQuery {
  const { where, params } = buildLocalWorkOrdersWhere(options);
  return {
    sql: `select count(*)::int as total from appfolio_work_orders wo where ${where}`,
    params,
  };
}
