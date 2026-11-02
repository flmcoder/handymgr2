import { createHash } from 'node:crypto';

export const DEFAULT_REPORT_CACHE_ROW_THRESHOLD = 5_000;
export const DEFAULT_REPORT_CACHE_BATCH_SIZE = 500;
export const DEFAULT_REPORT_CACHE_TTL_MS = 30 * 60_000;

function stableValue(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(stableValue);
  if (!value || typeof value !== 'object') return value;

  return Object.keys(value as Record<string, unknown>)
    .sort()
    .reduce<Record<string, unknown>>((result, key) => {
      result[key] = stableValue((value as Record<string, unknown>)[key]);
      return result;
    }, {});
}

// Empty scope containers mean "all", so they are dropped before hashing.
const EMPTY_SCOPE_KEYS = ['properties_ids', 'property_groups_ids', 'portfolios_ids', 'owners_ids'] as const;

function withoutEmptyScopes(payload: Record<string, unknown>): Record<string, unknown> {
  const properties = payload.properties;
  if (!properties || typeof properties !== 'object' || Array.isArray(properties)) return payload;
  const scopes = properties as Record<string, unknown>;
  const kept: Record<string, unknown> = {};
  for (const key of Object.keys(scopes)) {
    const value = scopes[key];
    const isEmptyScope = (EMPTY_SCOPE_KEYS as readonly string[]).includes(key) && Array.isArray(value) && value.length === 0;
    if (!isEmptyScope) kept[key] = value;
  }
  const { properties: _ignored, ...rest } = payload;
  return Object.keys(kept).length ? { ...rest, properties: kept } : rest;
}

export function buildReportCacheKey(reportName: string, payload: Record<string, unknown>): string {
  return createHash('sha256')
    .update(JSON.stringify({ reportName, payload: stableValue(withoutEmptyScopes(payload)) }))
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