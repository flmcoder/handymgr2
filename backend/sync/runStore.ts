/**
 * Sync run and cursor store.
 *
 * Provides typed helpers to:
 *  - Create a new sync job run row
 *  - Persist a page cursor after each successful fetch
 *  - Mark a run completed or failed
 *  - Load the latest live cursor for an endpoint (for incremental restarts)
 *
 * The execution_start_cursor is set once at job start and is the value
 * that should be used as the `LastUpdatedAtFrom` filter for the NEXT
 * incremental run — NOT the timestamp of the last record received.
 */

import { randomUUID } from 'node:crypto';
import { eq, and, desc } from 'drizzle-orm';
import { db } from '../db.ts';
import { syncJobRuns, syncJobCursors } from '../schema.ts';

export interface RunContext {
  runId: string;
  endpointKey: string;
  apiVersion: string;
  triggerType: string;
  executionStartCursor: string; // ISO timestamp of job start — used for next incremental filter
  filtersFingerprint: string | null;
}

/** Start a new sync run. Returns the run context needed by downstream workers. */
export async function startRun(opts: {
  endpointKey: string;
  apiVersion: string;
  triggerType: string;
  filtersFingerprint?: string;
}): Promise<RunContext> {
  const runId = randomUUID();
  const executionStartCursor = new Date().toISOString();

  await db.insert(syncJobRuns).values({
    runId,
    endpointKey: opts.endpointKey,
    apiVersion: opts.apiVersion,
    triggerType: opts.triggerType,
    status: 'running',
    filtersFingerprint: opts.filtersFingerprint ?? null,
    executionStartCursor,
  });

  return {
    runId,
    endpointKey: opts.endpointKey,
    apiVersion: opts.apiVersion,
    triggerType: opts.triggerType,
    executionStartCursor,
    filtersFingerprint: opts.filtersFingerprint ?? null,
  };
}

/** Checkpoint a successfully-fetched page's cursor. */
export async function persistPageCursor(opts: {
  runId: string;
  endpointKey: string;
  pageIndex: number;
  cursorIn: string | null;
  cursorOut: string | null;
  cursorExpiresAt: Date | null;
  recordCount: number;
  retriesUsed?: number;
}): Promise<void> {
  await db.insert(syncJobCursors).values({
    runId: opts.runId,
    endpointKey: opts.endpointKey,
    pageIndex: opts.pageIndex,
    cursorIn: opts.cursorIn,
    cursorOut: opts.cursorOut,
    cursorExpiresAt: opts.cursorExpiresAt,
    recordCount: opts.recordCount,
    retriesUsed: opts.retriesUsed ?? 0,
    isTerminal: opts.cursorOut === null,
  });
}

/** Update run totals and mark as completed or failed. */
export async function finalizeRun(
  runId: string,
  outcome: 'completed' | 'failed' | 'paused',
  totals?: { rowsUpserted?: number; rowsSkipped?: number; pagesCompleted?: number; lastError?: string },
): Promise<void> {
  await db
    .update(syncJobRuns)
    .set({
      status: outcome,
      completedAt: new Date(),
      rowsUpserted: totals?.rowsUpserted ?? 0,
      rowsSkipped: totals?.rowsSkipped ?? 0,
      pagesCompleted: totals?.pagesCompleted ?? 0,
      lastError: totals?.lastError ?? null,
    })
    .where(eq(syncJobRuns.runId, runId));
}

/** Increment run counters in-place after each page. */
export async function incrementRunCounters(
  runId: string,
  delta: { rowsUpserted?: number; rowsSkipped?: number; pagesCompleted?: number },
): Promise<void> {
  // Drizzle doesn't support increment natively; use raw sql increment.
  const { sql } = await import('drizzle-orm');
  await db
    .update(syncJobRuns)
    .set({
      rowsUpserted: sql`${syncJobRuns.rowsUpserted} + ${delta.rowsUpserted ?? 0}`,
      rowsSkipped:  sql`${syncJobRuns.rowsSkipped}  + ${delta.rowsSkipped  ?? 0}`,
      pagesCompleted: sql`${syncJobRuns.pagesCompleted} + ${delta.pagesCompleted ?? 0}`,
    })
    .where(eq(syncJobRuns.runId, runId));
}

/**
 * Returns the executionStartCursor from the most recent successfully completed
 * run for the given endpoint.  Used to populate LastUpdatedAtFrom on incremental runs.
 * Returns null if no prior completed run exists (triggers full backfill).
 */
export async function getLastSuccessfulCursor(endpointKey: string): Promise<string | null> {
  const rows = await db
    .select({ executionStartCursor: syncJobRuns.executionStartCursor })
    .from(syncJobRuns)
    .where(and(
      eq(syncJobRuns.endpointKey, endpointKey),
      eq(syncJobRuns.status, 'completed'),
    ))
    .orderBy(desc(syncJobRuns.startedAt))
    .limit(1);

  return rows[0]?.executionStartCursor ?? null;
}
