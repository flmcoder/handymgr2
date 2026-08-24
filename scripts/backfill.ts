/**
 * Initial AppFolio backfill into Postgres.
 *
 * Uses the same bootstrap + sync logic as the backend runtime:
 *   - load env from the repo root
 *   - initialize the Postgres connection via backend/db.ts
 *   - trigger a full historical sync for the core AppFolio entities
 *
 * Usage:
 *   npm run backfill:appfolio
 *   # or
 *   npx tsx scripts/backfill.ts
 */

import path from 'node:path';
import { fileURLToPath } from 'node:url';

const scriptPath = fileURLToPath(import.meta.url);
const repoRoot = path.resolve(path.dirname(scriptPath), '..');

const { config: loadDotenv } = await import('dotenv');
loadDotenv({ path: path.join(repoRoot, '.env') });
loadDotenv({ path: path.join(repoRoot, '.env.local'), override: true });

const { pingDatabase, closeDatabase } = await import('../backend/db.ts');
const { runSync } = await import('../backend/sync/syncRunner.ts');

const ENDPOINTS = [
  { key: 'v0:properties', label: 'properties' },
  { key: 'v0:units', label: 'units' },
  { key: 'v2:tenant_directory', label: 'tenants' },
  { key: 'v0:work_orders', label: 'work orders' },
] as const;

async function runBackfill(): Promise<void> {
  console.log('[backfill] Checking Postgres connectivity...');
  const dbOk = await pingDatabase();
  if (!dbOk) {
    throw new Error('Postgres is not reachable. Confirm SQL_SE/DB_NAME or the local DB env vars are set before running the sync.');
  }

  console.log('[backfill] Starting AppFolio full backfill for properties, units, tenants, and work orders...');

  const summaries = [] as Array<{ endpointKey: string; status: string; rowsUpserted: number; rowsSkipped: number; pagesCompleted: number; error?: string }>;

  for (const { key, label } of ENDPOINTS) {
    console.log(`[backfill] ${label.toUpperCase()} -> ${key}`);

    const result = await runSync({
      endpointKey: key,
      triggerType: 'backfill',
      forceLookback: true,
      lookbackDays: 3650,
    });

    console.log(`[backfill] ${label} summary`, {
      status: result.status,
      pagesCompleted: result.pagesCompleted,
      rowsUpserted: result.rowsUpserted,
      rowsSkipped: result.rowsSkipped,
      error: result.error ?? null,
    });

    summaries.push({
      endpointKey: key,
      status: result.status,
      rowsUpserted: result.rowsUpserted,
      rowsSkipped: result.rowsSkipped,
      pagesCompleted: result.pagesCompleted,
      error: result.error,
    });

    if (result.status !== 'completed') {
      throw new Error(`${key} failed during backfill: ${result.error ?? 'unknown error'}`);
    }
  }

  console.log('[backfill] Complete. All AppFolio backfill runs finished successfully.');
  console.log('[backfill] Summary', summaries);
}

try {
  await runBackfill();
} catch (error) {
  const message = error instanceof Error ? error.message : String(error);
  console.error('[backfill] FAIL:', message);
  process.exitCode = 1;
} finally {
  await closeDatabase();
}
