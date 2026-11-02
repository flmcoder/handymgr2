import { createHash } from 'node:crypto';

export const DEFAULT_REPORT_CACHE_ROW_THRESHOLD = 5_000;
export const DEFAULT_REPORT_CACHE_BATCH_SIZE = 500;
export const DEFAULT_REPORT_CACHE_TTL_MS = 30 * 60_000;

// Recursively drops null, undefined, empty arrays, and empty objects, then sorts keys,
// so semantically empty fields hash the same whether a client sent them or not.
function normalizeForKey(value: unknown): unknown {
  if (value === null || value === undefined) return undefined;
  if (Array.isArray(value)) {
    const items = value.map(normalizeForKey).filter((item) => item !== undefined);
    return items.length ? items : undefined;
  }
  if (typeof value !== 'object') return value;

  const normalized: Record<string, unknown> = {};
  for (const key of Object.keys(value as Record<string, unknown>).sort()) {
    const child = normalizeForKey((value as Record<string, unknown>)[key]);
    if (child !== undefined) normalized[key] = child;
  }
  return Object.keys(normalized).length ? normalized : undefined;
}

export function buildReportCacheKey(reportName: string, payload: Record<string, unknown>): string {
  const normalized = normalizeForKey(payload) ?? {};
  return createHash('sha256')
    .update(JSON.stringify({ reportName, payload: normalized }))
    .digest('hex');
}

export function extractReportRows(payload: unknown): unknown[] {
  if (Array.isArray(payload)) return payload;
  if (!payload || typeof payload !== 'object') return [];
  const record = payload as Record<string, unknown>;
  if (Array.isArray(record.results)) return record.results;
  if (Array.isArray(record.data)) return record.data;
  return [];
}

export function shouldHydrateReport(rowCount: number, threshold = DEFAULT_REPORT_CACHE_ROW_THRESHOLD): boolean {
  return Number.isFinite(rowCount) && rowCount >= Math.max(1, threshold);
}

export function chunkReportRows<T>(rows: T[], batchSize = DEFAULT_REPORT_CACHE_BATCH_SIZE): T[][] {
  const size = Math.max(1, Math.floor(batchSize));
  const chunks: T[][] = [];
  for (let index = 0; index < rows.length; index += size) {
    chunks.push(rows.slice(index, index + size));
  }
  return chunks;
}