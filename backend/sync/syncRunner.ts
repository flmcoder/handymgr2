/**
 * AppFolio sync runner.
 *
 * Executes a paginated AppFolio fetch for a given endpoint, archiving every
 * raw response page before mapping to domain tables.
 *
 * Rules enforced:
 *  1. Cursor follows provider's exact next_page_path (v0) / next_page_url (v2).
 *  2. v2 cursors are checked for expiry before each follow; expired → abort + restart signal.
 *  3. Every raw response page is written to appfolio_raw_responses before any upsert.
 *  4. Rate-limit token is acquired before each request via acquireToken().
 *  5. Retry-After is honored as authoritative on 429.
 *  6. executionStartCursor (job start timestamp) is used for incremental LastUpdatedAtFrom.
 *
 * Usage:
 *   const summary = await runSync({ endpointKey: 'v0:units', triggerType: 'nightly' });
 */

import { db } from '../db.ts';
import { appfolioRawResponses } from '../schema.ts';
import { afFetch, isCursorExpired, AfFetchError } from './fetchWorker.ts';
import { afBaseUrl } from './afCredentials.ts';
import {
  startRun, persistPageCursor, finalizeRun, incrementRunCounters,
  getLastSuccessfulCursor,
} from './runStore.ts';
import {
  upsertPropertyGroups,
  upsertProperties, upsertUnits, upsertWorkOrders,
  upsertEstimates, upsertTurnTracker,
  upsertMaintenanceTechUsers,
  upsertUnitInspections, upsertTenantDirectory, upsertUnitTurnDetails, upsertUnitVacancies,
} from './repositories.ts';

// ── Endpoint Definitions ─────────────────────────────────────────────────────

interface EndpointDef {
  apiVersion: 'v0' | 'v2';
  buildFirstRequest?: (opts: { baseUrl: string; incrementalFrom: string | null; lookbackDays: number; forceLookback: boolean }) => { url: string; method?: 'GET' | 'POST'; body?: any };
  buildFirstUrl?: (opts: { baseUrl: string; incrementalFrom: string | null; lookbackDays: number; forceLookback: boolean }) => string;
  /** Extract the rows array from the response payload. */
  extractRows: (data: any) => any[];
  upsert: (rows: any[]) => Promise<{ upserted: number; skipped: number }>;
}

function toIsoDate(value: Date | string | null | undefined): string {
  if (!value) return '';
  const date = value instanceof Date ? value : new Date(String(value));
  if (Number.isNaN(date.getTime())) return '';
  return date.toISOString().slice(0, 10);
}

function isDispatchMaintenanceRole(value: unknown): boolean {
  const role = String(value ?? '').trim().toLowerCase().replace(/[_-]+/g, ' ');
  if (!role) return false;
  if (role === 'maintenance tech' || role === 'maintenance technician') return true;
  if (role === 'technician' || role === 'tech') return true;
  if (role.includes('maintenance') && role.includes('tech')) return true;
  if (role.includes('maintenance') && role.includes('supervisor')) return true;
  if (role.includes('service') && role.includes('tech')) return true;
  if (role.includes('make ready') || role.includes('turnover')) return true;
  if (role.includes('handyman')) return true;
  return false;
}

function buildV2ReportRequest(report: string, body: Record<string, any>): { url: string; method: 'POST'; body: any } {
  return {
    url: `${afBaseUrl('v2')}/api/v2/reports/${report}.json`,
    method: 'POST',
    body,
  };
}

const ENDPOINTS: Record<string, EndpointDef> = {
  'v0:units': {
    apiVersion: 'v0',
    buildFirstUrl: ({ baseUrl, incrementalFrom, lookbackDays, forceLookback }) => {
      const clampedLookback = Math.max(1, Math.min(3650, Number(lookbackDays || 180)));
      const from = incrementalFrom
        ? (forceLookback
          ? new Date(Date.now() - clampedLookback * 86400_000).toISOString().slice(0, 19) + 'Z'
          : new Date(incrementalFrom).toISOString().slice(0, 19) + 'Z')
        : new Date(Date.now() - clampedLookback * 86400_000).toISOString().slice(0, 19) + 'Z';
      return `${baseUrl}/api/v0/units?filters%5BLastUpdatedAtFrom%5D=${encodeURIComponent(from)}&page%5Bsize%5D=100`;
    },
    extractRows: (data) => data?.data ?? data?.results ?? [],
    upsert: upsertUnits,
  },

  'v0:properties': {
    apiVersion: 'v0',
    buildFirstUrl: ({ baseUrl, incrementalFrom, lookbackDays, forceLookback }) => {
      const clampedLookback = Math.max(1, Math.min(3650, Number(lookbackDays || 180)));
      const from = incrementalFrom
        ? (forceLookback
          ? new Date(Date.now() - clampedLookback * 86400_000).toISOString().slice(0, 19) + 'Z'
          : new Date(incrementalFrom).toISOString().slice(0, 19) + 'Z')
        : new Date(Date.now() - clampedLookback * 86400_000).toISOString().slice(0, 19) + 'Z';
      return `${baseUrl}/api/v0/properties?filters%5BLastUpdatedAtFrom%5D=${encodeURIComponent(from)}&page%5Bsize%5D=100`;
    },
    extractRows: (data) => data?.data ?? data?.results ?? [],
    upsert: upsertProperties,
  },

  'v0:property_groups': {
    apiVersion: 'v0',
    buildFirstUrl: ({ baseUrl, incrementalFrom, lookbackDays, forceLookback }) => {
      const clampedLookback = Math.max(1, Math.min(3650, Number(lookbackDays || 180)));
      const from = incrementalFrom
        ? (forceLookback
          ? new Date(Date.now() - clampedLookback * 86400_000).toISOString().slice(0, 19) + 'Z'
          : new Date(incrementalFrom).toISOString().slice(0, 19) + 'Z')
        : new Date(Date.now() - clampedLookback * 86400_000).toISOString().slice(0, 19) + 'Z';
      return `${baseUrl}/api/v0/property_groups?filters%5BLastUpdatedAtFrom%5D=${encodeURIComponent(from)}&page%5Bsize%5D=100`;
    },
    extractRows: (data) => data?.data ?? data?.results ?? [],
    upsert: upsertPropertyGroups,
  },

  'v0:users': {
    apiVersion: 'v0',
    buildFirstUrl: ({ baseUrl, incrementalFrom, lookbackDays, forceLookback }) => {
      const clampedLookback = Math.max(1, Math.min(3650, Number(lookbackDays || 180)));
      const from = incrementalFrom
        ? (forceLookback
          ? new Date(Date.now() - clampedLookback * 86400_000).toISOString().slice(0, 19) + 'Z'
          : new Date(incrementalFrom).toISOString().slice(0, 19) + 'Z')
        : new Date(Date.now() - clampedLookback * 86400_000).toISOString().slice(0, 19) + 'Z';
      return `${baseUrl}/api/v0/users?filters%5BLastUpdatedAtFrom%5D=${encodeURIComponent(from)}&page%5Bsize%5D=100`;
    },
    extractRows: (data) => {
      const rows = data?.data ?? data?.results ?? [];
      if (!Array.isArray(rows)) return [];
      return rows.filter((row: any) => {
        const role = row?.UserRole ?? row?.user_role ?? row?.Role ?? row?.role;
        return isDispatchMaintenanceRole(role);
      });
    },
    upsert: upsertMaintenanceTechUsers,
  },

  'v0:work_orders': {
    apiVersion: 'v0',
    buildFirstUrl: ({ baseUrl, incrementalFrom, lookbackDays, forceLookback }) => {
      const clampedLookback = Math.max(1, Math.min(3650, Number(lookbackDays || 180)));
      const from = incrementalFrom
        ? (forceLookback
          ? new Date(Date.now() - clampedLookback * 86400_000).toISOString().slice(0, 19) + 'Z'
          : new Date(incrementalFrom).toISOString().slice(0, 19) + 'Z')
        : new Date(Date.now() - clampedLookback * 86400_000).toISOString().slice(0, 19) + 'Z';
      return `${baseUrl}/api/v0/work_orders?filters%5BLastUpdatedAtFrom%5D=${encodeURIComponent(from)}&page%5Bsize%5D=100`;
    },
    extractRows: (data) => data?.data ?? data?.results ?? [],
    upsert: upsertWorkOrders,
  },

  'v2:unit_inspection': {
    apiVersion: 'v2',
    buildFirstRequest: ({ incrementalFrom, lookbackDays, forceLookback }) => buildV2ReportRequest('unit_inspection', {
      unit_visibility: 'active',
      last_inspection_on_from: toIsoDate(
        forceLookback
          ? new Date(Date.now() - Math.max(1, Math.min(3650, Number(lookbackDays || 180))) * 86400_000)
          : (incrementalFrom || new Date()),
      ),
      include_blank_inspection_date: '1',
      columns: [
        'property', 'property_name', 'property_id', 'property_address', 'property_street', 'property_street2',
        'property_city', 'property_state', 'property_zip', 'unit_name', 'last_inspection_date', 'tenant_name',
        'tenant_primary_phone_number', 'move_in_date', 'move_out_date', 'unit_id', 'occupancy_id', 'rentable', 'unit_tags',
      ],
    }),
    extractRows: (data) => Array.isArray(data) ? data : (data?.results ?? data?.data ?? []),
    upsert: upsertUnitInspections,
  },

  'v2:tenant_directory': {
    apiVersion: 'v2',
    buildFirstRequest: ({ incrementalFrom, lookbackDays, forceLookback }) => buildV2ReportRequest('tenant_directory', {
      tenant_visibility: 'active',
      tenant_statuses: ['0', '4'],
      tenant_types: 'all',
      property_visibility: 'active',
      last_updated_at_from: toIsoDate(
        forceLookback
          ? new Date(Date.now() - Math.max(1, Math.min(3650, Number(lookbackDays || 180))) * 86400_000)
          : (incrementalFrom || new Date(Date.now() - 365 * 86400_000)),
      ),
      columns: [
        'property', 'property_name', 'property_id', 'property_address', 'unit', 'tenant', 'status', 'tenant_type',
        'phone_numbers', 'emails', 'move_in', 'lease_to', 'rent', 'tenant_tags', 'tenant_agent', 'tenant_visibility',
        'move_out', 'unit_tags', 'occupancy_id',
      ],
    }),
    extractRows: (data) => Array.isArray(data) ? data : (data?.results ?? data?.data ?? []),
    upsert: upsertTenantDirectory,
  },

  'v2:unit_turn_detail': {
    apiVersion: 'v2',
    buildFirstRequest: ({ incrementalFrom, lookbackDays, forceLookback }) => buildV2ReportRequest('unit_turn_detail', {
      property_visibility: 'active',
      move_out_date_from: toIsoDate(
        forceLookback
          ? new Date(Date.now() - Math.max(1, Math.min(3650, Number(lookbackDays || 180))) * 86400_000)
          : (incrementalFrom || new Date(Date.now() - 365 * 86400_000)),
      ),
      move_out_date_to: toIsoDate(new Date(Date.now() + 365 * 86400_000)),
      unit_turn_status: 'All',
      columns: [
        'property', 'unit', 'unit_turn_id', 'notes', 'reference_user', 'move_out_date', 'turn_end_date',
        'expected_move_in_date', 'target_days_to_complete', 'total_days_to_complete', 'labor_from_work_orders',
        'purchase_orders_from_work_orders', 'billables_from_work_orders', 'inventory_from_work_orders', 'total_billed',
        'property_id', 'unit_id',
      ],
    }),
    extractRows: (data) => Array.isArray(data) ? data : (data?.results ?? data?.data ?? []),
    upsert: upsertUnitTurnDetails,
  },

  'v2:unit_vacancy': {
    apiVersion: 'v2',
    buildFirstRequest: () => buildV2ReportRequest('unit_vacancy', {
      unit_visibility: 'active',
      level_of_detail: 'detail_view',
      bedrooms: 'any',
      bathrooms: 'any',
      columns: [
        'property', 'property_name', 'property_id', 'unit', 'unit_id', 'unit_tags',
        'bed_and_bath', 'unit_status', 'rent_ready', 'days_vacant', 'last_rent',
        'last_move_out', 'available_on', 'computed_market_rent',
      ],
    }),
    extractRows: (data) => Array.isArray(data) ? data : (data?.results ?? data?.data ?? []),
    upsert: upsertUnitVacancies,
  },
};

// ── Sync Runner ───────────────────────────────────────────────────────────────

export interface SyncRunSummary {
  runId: string;
  endpointKey: string;
  status: 'completed' | 'failed' | 'cursor_expired';
  pagesCompleted: number;
  rowsUpserted: number;
  rowsSkipped: number;
  error?: string;
}

export async function runSync(opts: {
  endpointKey: string;
  triggerType?: string;
  filtersFingerprint?: string;
  /** Hard cap on pages per run chunk (Phase 6 backfill safety). 0 = unlimited. */
  maxPages?: number;
  /** Optional historical backfill window, in days, used by compatible endpoints. */
  lookbackDays?: number;
  /** Force compatible endpoints to use lookbackDays instead of incremental cursor. */
  forceLookback?: boolean;
}): Promise<SyncRunSummary> {
  const {
    endpointKey,
    triggerType = 'manual',
    filtersFingerprint,
    maxPages = 0,
    lookbackDays = 180,
    forceLookback = false,
  } = opts;

  const def = ENDPOINTS[endpointKey];
  if (!def) throw new Error(`Unknown endpointKey: ${endpointKey}`);

  const baseUrl = afBaseUrl(def.apiVersion);
  const incrementalFrom = await getLastSuccessfulCursor(endpointKey);

  const ctx = await startRun({
    endpointKey,
    apiVersion: def.apiVersion,
    triggerType,
    filtersFingerprint,
  });

  const firstRequest = def.buildFirstRequest
    ? def.buildFirstRequest({ baseUrl, incrementalFrom, lookbackDays, forceLookback })
    : { url: def.buildFirstUrl?.({ baseUrl, incrementalFrom, lookbackDays, forceLookback }) || '' };
  let url: string | null = firstRequest.url || null;
  let method: 'GET' | 'POST' = firstRequest.method || 'GET';
  let body: any = firstRequest.body;
  let pageIndex = 0;
  let totalUpserted = 0;
  let totalSkipped = 0;

  console.log(`[syncRunner] START runId=${ctx.runId} endpoint=${endpointKey} triggerType=${triggerType} incrementalFrom=${incrementalFrom ?? 'none'}`);

  try {
    while (url) {
      // Phase 6 safety: respect per-chunk page cap.
      if (maxPages > 0 && pageIndex >= maxPages) {
        console.log(`[syncRunner] maxPages=${maxPages} reached at page ${pageIndex}; pausing run.`);
        await finalizeRun(ctx.runId, 'paused', {
          rowsUpserted: totalUpserted, rowsSkipped: totalSkipped, pagesCompleted: pageIndex,
        });
        return { runId: ctx.runId, endpointKey, status: 'completed', pagesCompleted: pageIndex, rowsUpserted: totalUpserted, rowsSkipped: totalSkipped };
      }

      // Fetch page.
      const result = await afFetch(url, {
        endpointKey,
        apiVersion: def.apiVersion,
        runId: ctx.runId,
        attempt: pageIndex + 1,
        method,
        body: body ? JSON.stringify(body) : undefined,
      });

      // Archive raw response BEFORE any mapping.
      await db.insert(appfolioRawResponses).values({
        runId: ctx.runId,
        endpointKey,
        pageIndex,
        cursorIn: url,
        cursorOut: result.cursorOut,
        statusCode: result.status,
        recordCount: def.extractRows(result.data).length,
        responseJson: result.data,
        fetchedAt: new Date(),
      });

      // Persist cursor checkpoint.
      await persistPageCursor({
        runId: ctx.runId,
        endpointKey,
        pageIndex,
        cursorIn: url,
        cursorOut: result.cursorOut,
        cursorExpiresAt: result.cursorExpiresAt,
        recordCount: def.extractRows(result.data).length,
      });

      // Map and upsert rows into domain table.
      const rows = def.extractRows(result.data);
      const { upserted, skipped } = await def.upsert(rows);
      totalUpserted += upserted;
      totalSkipped += skipped;

      await incrementRunCounters(ctx.runId, {
        pagesCompleted: 1, rowsUpserted: upserted, rowsSkipped: skipped,
      });

      pageIndex++;

      // Advance cursor — v2 expiry check before following.
      if (result.cursorOut) {
        if (def.apiVersion === 'v2' && isCursorExpired(result.cursorExpiresAt)) {
          console.warn(`[syncRunner] v2 cursor expired at page ${pageIndex}; signalling restart.`);
          await finalizeRun(ctx.runId, 'failed', {
            lastError: 'v2_cursor_expired',
            pagesCompleted: pageIndex, rowsUpserted: totalUpserted, rowsSkipped: totalSkipped,
          });
          return { runId: ctx.runId, endpointKey, status: 'cursor_expired', pagesCompleted: pageIndex, rowsUpserted: totalUpserted, rowsSkipped: totalSkipped, error: 'v2_cursor_expired' };
        }
        url = def.apiVersion === 'v2' ? result.cursorOut : `${baseUrl}${result.cursorOut}`;
        method = def.apiVersion === 'v2' ? 'POST' : 'GET';
        body = undefined;
      } else {
        url = null; // Terminal — no more pages.
      }
    }

    await finalizeRun(ctx.runId, 'completed', {
      pagesCompleted: pageIndex, rowsUpserted: totalUpserted, rowsSkipped: totalSkipped,
    });

    console.log(`[syncRunner] DONE runId=${ctx.runId} pages=${pageIndex} upserted=${totalUpserted} skipped=${totalSkipped}`);

    return { runId: ctx.runId, endpointKey, status: 'completed', pagesCompleted: pageIndex, rowsUpserted: totalUpserted, rowsSkipped: totalSkipped };

  } catch (err) {
    const errorText = err instanceof AfFetchError
      ? `${err.kind}: ${err.message}`
      : String((err as any)?.message ?? err);

    console.error(`[syncRunner] FAILED runId=${ctx.runId} endpoint=${endpointKey}`, errorText);

    await finalizeRun(ctx.runId, 'failed', {
      lastError: errorText, pagesCompleted: pageIndex, rowsUpserted: totalUpserted, rowsSkipped: totalSkipped,
    });

    return { runId: ctx.runId, endpointKey, status: 'failed', pagesCompleted: pageIndex, rowsUpserted: totalUpserted, rowsSkipped: totalSkipped, error: errorText };
  }
}
