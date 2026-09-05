/**
 * Pure helpers that decide which operational log lines are worth emitting.
 * Extracted so the noise-reduction rules are unit-testable without booting
 * the Express server or a database connection.
 */

export type SyncSummaryLike = {
  status?: string;
  rowsUpserted?: number;
  rowsSkipped?: number;
  pagesCompleted?: number;
};

/**
 * Decide whether a completed sync-scheduler endpoint run should be logged.
 * Returns false for the noisy "nothing changed" ticks, true when rows were
 * upserted or the run failed.
 */
export function shouldLogSyncSummary(summary: SyncSummaryLike | null | undefined): boolean {
  if (!summary) return false;
  const upserted = Number(summary.rowsUpserted ?? 0);
  return upserted > 0 || String(summary.status || '') === 'failed';
}

/** Single-line sync summary including the upserted row count. */
export function formatSyncSummaryLine(endpointKey: string, summary: SyncSummaryLike): string {
  return `[SYNC] ${endpointKey} | status=${String(summary.status || 'unknown')} | upserted=${Number(summary.rowsUpserted ?? 0)} | skipped=${Number(summary.rowsSkipped ?? 0)} | pages=${Number(summary.pagesCompleted ?? 0)}`;
}

export type SignInInfo = {
  method: 'device_setup' | 'otp' | 'password';
  userName: string;
  role: string;
  email?: string;
  scopeUuid?: string;
  tokenKind?: 'db' | 'signed-fallback';
};

/** Single-line sign-in audit record. */
export function formatSignInLine(info: SignInInfo): string {
  const parts = [
    `[auth] sign-in`,
    `method=${info.method}`,
    `user="${String(info.userName || '').replace(/"/g, '')}"`,
  ];
  if (info.email) parts.push(`email="${String(info.email).replace(/"/g, '')}"`);
  parts.push(`role=${info.role}`);
  if (info.method === 'otp') parts.push(`scope=${info.scopeUuid || 'none'}`);
  if (info.tokenKind && info.tokenKind !== 'db') parts.push(`token=${info.tokenKind}`);
  return parts.join(' | ');
}
