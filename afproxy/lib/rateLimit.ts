// ============================================================================
// lib/rateLimit.ts — Sliding-window in-memory rate limiter.
//
// Designed for Val Town / single-instance Deno deployments.  No external
// store is needed — the module-level Map persists for the lifetime of the
// isolate (typically hours between cold starts).
//
// Algorithm: sliding-window log.  Each entry is the timestamp of one
// accepted request.  When the window slides past old entries they are
// trimmed.  A request is rejected (not recorded) when the trimmed count
// already matches the limit, which ensures the count never exceeds `limit`.
//
// Keying conventions (enforced by callers, not this module):
//   otp_id:<identifier>   — per-identifier OTP request rate
//   otp_global            — global OTP throughput cap (all identifiers)
//   setup:<ip>            — device-setup pin attempts by source IP
//   wh:<ip>               — inbound webhook delivery by source IP
//   portal:<ip>           — magic-portal action throughput by source IP
//
// Thread safety: single-threaded Deno runtime; no concurrency hazard.
// ============================================================================

const _store = new Map<string, number[]>();
let _lastCleanup = 0;

const CLEANUP_INTERVAL_MS = 5 * 60_000; // prune dead keys every 5 minutes
const MAX_WINDOW_MS = 15 * 60_000 + 1_000; // slightly > longest window in use

function trimWindow(
  timestamps: number[],
  windowMs: number,
  now: number,
): number[] {
  const cutoff = now - windowMs;
  let i = 0;
  while (i < timestamps.length && timestamps[i] <= cutoff) i++;
  return i > 0 ? timestamps.slice(i) : timestamps;
}

function maybePurge(now: number): void {
  if (now - _lastCleanup < CLEANUP_INTERVAL_MS) return;
  _lastCleanup = now;
  const cutoff = now - MAX_WINDOW_MS;
  for (const [key, ts] of _store) {
    if (!ts.length || ts[ts.length - 1] <= cutoff) _store.delete(key);
  }
}

/**
 * Check and (if allowed) record one request for `key`.
 *
 * Returns `{ allowed: true, count, limit }` when the request is within the
 * sliding window budget, or `{ allowed: false, count, limit, retryAfterMs }`
 * when the budget is exhausted.  A rejected request is NOT recorded, so the
 * window does not advance on abuse attempts.
 */
export function checkRateLimit(
  key: string,
  limit: number,
  windowMs: number,
): { allowed: boolean; count: number; limit: number; retryAfterMs?: number } {
  const now = Date.now();
  maybePurge(now);
  const trimmed = trimWindow(_store.get(key) ?? [], windowMs, now);
  if (trimmed.length >= limit) {
    // Do NOT push — prevent timestamps from growing past the limit.
    _store.set(key, trimmed);
    const retryAfterMs = windowMs - (now - trimmed[0]) + 1;
    return { allowed: false, count: trimmed.length, limit, retryAfterMs };
  }
  trimmed.push(now);
  _store.set(key, trimmed);
  return { allowed: true, count: trimmed.length, limit };
}
