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

export function resolveWorkOrderHistoryDays(value: unknown, maximumDays = 3_650): number | null {
  if (value === undefined || value === null || String(value).trim() === '') return null;
  const parsed = Number.parseInt(String(value), 10);
  if (!Number.isFinite(parsed)) return maximumDays;
  return Math.max(1, Math.min(maximumDays, parsed));
}