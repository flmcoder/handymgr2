export function buildActiveWorkOrdersUrl(
  baseUrl: unknown,
  propertyGroupUuid: unknown = '',
  requestedLimit: unknown = 5_000,
): string {
  const base = String(baseUrl || '').trim().replace(/\/+$/, '');
  const parsedLimit = Number.parseInt(String(requestedLimit || ''), 10);
  const limit = Number.isFinite(parsedLimit) ? Math.max(1, Math.min(20_000, parsedLimit)) : 5_000;
  const params = new URLSearchParams({ limit: String(limit) });
  const scope = String(propertyGroupUuid || '').trim();
  if (scope) params.set('property_group_id', scope);
  return `${base}/api/local/work_orders?${params.toString()}`;
}

export function resolveWorkOrderHistoryDays(value: unknown, maximumDays = 3_650): number | null {
  if (value === undefined || value === null || String(value).trim() === '') return null;
  const parsed = Number.parseInt(String(value), 10);
  if (!Number.isFinite(parsed)) return maximumDays;
  return Math.max(1, Math.min(maximumDays, parsed));
}