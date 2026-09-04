export type TenantCommsQuery = {
  limit: number;
  offset: number;
  page: number;
};

function parsePositiveInteger(value: unknown, fallback: number): number {
  const parsed = Number.parseInt(String(value ?? ''), 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

export function parseTenantCommsQuery(query: Record<string, unknown>): TenantCommsQuery {
  const limit = Math.min(100, parsePositiveInteger(query.limit, 50));
  const requestedPage = parsePositiveInteger(query.page, 1);
  const requestedOffset = Number.parseInt(String(query.offset ?? ''), 10);
  const offset = Number.isFinite(requestedOffset) && requestedOffset >= 0
    ? Math.min(250_000, requestedOffset)
    : Math.min(250_000, (requestedPage - 1) * limit);

  return { limit, offset, page: Math.floor(offset / limit) + 1 };
}