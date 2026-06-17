/**
 * AppFolio-safe rate limiter.
 *
 * Enforces three sliding windows simultaneously:
 *   - 8 requests / second
 *   - 256 requests / minute
 *   - 4096 requests / hour
 *
 * All limits are configurable via env vars:
 *   AF_MAX_PER_SECOND, AF_MAX_PER_MINUTE, AF_MAX_PER_HOUR
 *
 * When a 429 is received, the caller must invoke setRetryAfter(seconds) which
 * hard-blocks all requests for the exact duration specified by AppFolio.
 * This takes priority over all window-based throttling.
 */

const MAX_PER_SECOND = Number(process.env.AF_MAX_PER_SECOND ?? 8);
const MAX_PER_MINUTE = Number(process.env.AF_MAX_PER_MINUTE ?? 256);
const MAX_PER_HOUR   = Number(process.env.AF_MAX_PER_HOUR   ?? 4096);

// In-memory sliding window counters — sufficient for single-instance Render deployment.
// A Postgres-backed version can be layered in for horizontal scaling.
const windows = {
  second: { count: 0, windowStart: 0, max: MAX_PER_SECOND, ms: 1_000 },
  minute: { count: 0, windowStart: 0, max: MAX_PER_MINUTE, ms: 60_000 },
  hour:   { count: 0, windowStart: 0, max: MAX_PER_HOUR,   ms: 3_600_000 },
};

let retryAfterUntil = 0; // epoch ms — blocks ALL requests until this time

export function setRetryAfter(seconds: number): void {
  const until = Date.now() + seconds * 1_000;
  if (until > retryAfterUntil) {
    retryAfterUntil = until;
    console.warn(`[rateLimiter] Retry-After set: blocking for ${seconds}s until ${new Date(until).toISOString()}`);
  }
}

export function getRetryAfterRemaining(): number {
  return Math.max(0, retryAfterUntil - Date.now());
}

/** Advance the sliding window counter and return true if budget remains. */
function windowAllows(window: typeof windows.second): boolean {
  const now = Date.now();
  if (now - window.windowStart >= window.ms) {
    window.windowStart = now;
    window.count = 0;
  }
  return window.count < window.max;
}

/** Milliseconds until the most-constrained window resets. */
function msUntilWindowResets(): number {
  let minWait = 0;
  for (const w of Object.values(windows)) {
    const now = Date.now();
    if (w.count >= w.max) {
      const reset = w.windowStart + w.ms - now;
      minWait = Math.max(minWait, reset);
    }
  }
  return Math.max(0, minWait);
}

/**
 * Acquire a rate-limit token.  Waits until budget is available,
 * respecting both Retry-After blocks and sliding window limits.
 * Returns the time waited in ms (useful for logging).
 */
export async function acquireToken(): Promise<number> {
  const start = Date.now();
  let waited = 0;

  while (true) {
    // Retry-After from a 429 takes absolute priority.
    const retryBlock = getRetryAfterRemaining();
    if (retryBlock > 0) {
      await sleep(Math.min(retryBlock, 10_000));
      waited += Math.min(retryBlock, 10_000);
      continue;
    }

    // Check all sliding windows.
    if (
      windowAllows(windows.second) &&
      windowAllows(windows.minute) &&
      windowAllows(windows.hour)
    ) {
      windows.second.count++;
      windows.minute.count++;
      windows.hour.count++;
      return Date.now() - start;
    }

    const wait = Math.max(50, msUntilWindowResets());
    await sleep(Math.min(wait, 2_000));
    waited += wait;
  }
}

/** Snapshot current counters for logging / telemetry. */
export function getRateLimiterStats(): object {
  return {
    retryAfterRemainingMs: getRetryAfterRemaining(),
    second: { count: windows.second.count, max: windows.second.max },
    minute: { count: windows.minute.count, max: windows.minute.max },
    hour:   { count: windows.hour.count,   max: windows.hour.max },
  };
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
