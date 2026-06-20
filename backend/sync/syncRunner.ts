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
  upsertProperties, upsertUnits, upsertWorkOrders,
  upsertEstimates, upsertTurnTracker,
} from './repositories.ts';

// ── Endpoint Definitions ─────────────────────────────────────────────────────

interface EndpointDef {
  apiVersion: 'v0' | 'v2';
  buildFirstUrl: (opts: { baseUrl: string; incrementalFrom: string | null }) => string;
  /** Extract the rows array from the response payload. */
  extractRows: (data: any) => any[];
  upsert: (rows: any[]) => Promise<{ upserted: number; skipped: number }>;
}

const ENDPOINTS: Record<string, EndpointDef> = {
  'v0:units': {
    apiVersion: 'v0',
    buildFirstUrl: ({ baseUrl, incrementalFrom }) => {
      const from = incrementalFrom
        ? new Date(incrementalFrom).toISOString().slice(0, 19) + 'Z'
        : new Date(Date.now() - 5 * 365 * 86400_000).toISOString().slice(0, 19) + 'Z';
      return `${baseUrl}/api/v0/units?filters%5BLastUpdatedAtFrom%5D=${encodeURIComponent(from)}&page%5Bsize%5D=100`;
    },
    extractRows: (data) => data?.data ?? data?.results ?? [],
    upsert: upsertUnits,
  },

  'v0:properties': {
    apiVersion: 'v0',
    buildFirstUrl: ({ baseUrl, incrementalFrom }) => {
      const from = incrementalFrom
        ? new Date(incrementalFrom).toISOString().slice(0, 19) + 'Z'
        : new Date(Date.now() - 5 * 365 * 86400_000).toISOString().slice(0, 19) + 'Z';
      return `${baseUrl}/api/v0/properties?filters%5BLastUpdatedAtFrom%5D=${encodeURIComponent(from)}&page%5Bsize%5D=100`;
    },
    extractRows: (data) => data?.data ?? data?.results ?? [],
    upsert: upsertProperties,
  },

  'v0:work_orders': {
    apiVersion: 'v0',
    buildFirstUrl: ({ baseUrl, incrementalFrom }) => {
      const from = incrementalFrom
        ? new Date(incrementalFrom).toISOString().slice(0, 19) + 'Z'
        : new Date(Date.now() - 180 * 86400_000).toISOString().slice(0, 19) + 'Z';
      return `${baseUrl}/api/v0/work_orders?filters%5BLastUpdatedAtFrom%5D=${encodeURIComponent(from)}&page%5Bsize%5D=100`;
    },
    extractRows: (data) => data?.data ?? data?.results ?? [],
    upsert: upsertWorkOrders,
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
}): Promise<SyncRunSummary> {
  const { endpointKey, triggerType = 'manual', filtersFingerprint, maxPages = 0 } = opts;

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

  let url: string | null = def.buildFirstUrl({ baseUrl, incrementalFrom });
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
