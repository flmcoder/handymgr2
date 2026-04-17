// ============================================================================
// lib/fetchUtils.ts — fetchWithTimeout (full v8.9 retry logic),
//                     delay helper, daysAgo + today date utilities.
//
// Shared by every handler that talks to AppFolio or RingCentral.
// AppFolio rate limits: 8 req/s · 256 req/min · 4096 req/hr
// ============================================================================

import { FETCH_TIMEOUT_MS, PAGE_DELAY_MS } from "../config.ts";

// Re-export PAGE_DELAY_MS so handlers only need one import for pacing
export { PAGE_DELAY_MS };

// ── Rate-limiter: three independent per-bucket throttles ──────────────────────
// AppFolio API limits: 8 req/s · 256 req/min · 4096 req/hr
//
// Three buckets isolate different traffic classes so a slow report init never
// stalls time-sensitive v0 DB API calls:
//   v0Bucket     — Database API v0, all general calls   (6/s · 200/min · 3500/hr)
//   v2InitBucket — Reports API v2 report-init POSTs     (2/s · 30/min)
//   v2PageBucket — Reports API v2 pagination GETs       (5/s · 150/min)
//
// A 429 from ANY bucket applies a Retry-After cooldown to ALL three buckets
// because AppFolio rate limits are per API key, not per endpoint class.
const ONE_SEC_MS = 1000;
const ONE_MIN_MS = 60_000;
const ONE_HOUR_MS = 60 * ONE_MIN_MS;
export const MAX_RETRY_AFTER_MS = 60 * ONE_MIN_MS;

export interface RateBucketCfg {
  reqPerSec: number;
  reqPerMin: number;
  reqPerHour?: number;
  label?: string;
}

export class RateBucket {
  private readonly _timestamps: number[] = [];
  private _queue: Promise<void> = Promise.resolve();
  private _cooldownUntil = 0;

  constructor(private readonly cfg: RateBucketCfg) {}

  setCooldown(until: number): void {
    if (until > this._cooldownUntil) this._cooldownUntil = until;
  }

  private prune(nowMs: number): void {
    const cutoff = this.cfg.reqPerHour
      ? nowMs - ONE_HOUR_MS
      : nowMs - ONE_MIN_MS;
    while (this._timestamps.length && this._timestamps[0] < cutoff) {
      this._timestamps.shift();
    }
  }

  async acquire(): Promise<void> {
    let release!: () => void;
    const next = new Promise<void>((resolve) => { release = resolve; });
    const previous = this._queue;
    this._queue = next;

    await previous;
    try {
      while (true) {
        const now = Date.now();
        this.prune(now);

        const n = this._timestamps.length;
        const secCount = this._timestamps.filter((ts) => (now - ts) < ONE_SEC_MS).length;
        const minCount = this._timestamps.filter((ts) => (now - ts) < ONE_MIN_MS).length;

        const secWait = secCount >= this.cfg.reqPerSec
          ? Math.max(1, ONE_SEC_MS - (now - this._timestamps[n - secCount]))
          : 0;
        const minWait = minCount >= this.cfg.reqPerMin
          ? Math.max(1, ONE_MIN_MS - (now - this._timestamps[n - minCount]))
          : 0;
        const hourWait = (this.cfg.reqPerHour && n >= this.cfg.reqPerHour)
          ? Math.max(1, ONE_HOUR_MS - (now - this._timestamps[0]))
          : 0;
        const coolWait = this._cooldownUntil > now
          ? (this._cooldownUntil - now)
          : 0;

        const waitMs = Math.max(secWait, minWait, hourWait, coolWait);
        if (waitMs <= 0) {
          this._timestamps.push(now);
          return;
        }
        await delay(waitMs);
      }
    } finally {
      release();
    }
  }
}

export const v0Bucket = new RateBucket({ reqPerSec: 6, reqPerMin: 200, reqPerHour: 3500, label: "v0" });
export const v2InitBucket = new RateBucket({ reqPerSec: 2, reqPerMin: 30, label: "v2-init" });
export const v2PageBucket = new RateBucket({ reqPerSec: 5, reqPerMin: 150, label: "v2-page" });

/** Set Retry-After cooldown on every bucket — limit is per API key, not per class. */
function setAllBucketsCooldown(resumeAt: number): void {
  v0Bucket.setCooldown(resumeAt);
  v2InitBucket.setCooldown(resumeAt);
  v2PageBucket.setCooldown(resumeAt);
}

function parseRetryAfterMs(rawHeader: string | null): number {
  const raw = String(rawHeader || "").trim();
  if (!raw) return 2000;

  const sec = Number(raw);
  if (Number.isFinite(sec) && sec >= 0) {
    return Math.min(Math.floor(sec * 1000), MAX_RETRY_AFTER_MS);
  }

  const asDate = Date.parse(raw);
  if (!isNaN(asDate)) {
    return Math.min(Math.max(0, asDate - Date.now()), MAX_RETRY_AFTER_MS);
  }

  return 2000;
}

// ── Tiny delay helper (used for rate-limit pacing between paginated requests)
export function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// ── fetchWithTimeout ──────────────────────────────────────────────────────────
// Wraps fetch() with:
//   • AbortController timeout (default 30 s — bumped from 15 s for large reports)
//   • Auto-retry on 429 Rate Limit  — honours Retry-After header, capped at 60 s
//   • Auto-retry on 503 Unavailable — exponential back-off (2 s, 4 s, 8 s)
//   • Auto-retry on 533 AF Maintenance — fires nightly 9 PM–4 AM PST
//   • Never retries 422 — callers must change the payload before retrying
//   • Auto-retry on AbortError timeout — 1 s pause then re-attempt
export async function fetchWithTimeout(
  url: string,
  opts: RequestInit = {},
  timeoutMs: number = FETCH_TIMEOUT_MS,
  bucket: RateBucket = v0Bucket,
): Promise<Response> {
  const MAX_RETRIES = 6;

  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    await bucket.acquire();

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);

    try {
      const resp = await fetch(url, { ...opts, signal: controller.signal });
      clearTimeout(timer);

      // ── 422 Unprocessable Entity ───────────────────────────────────────
      // Do not retry semantic validation failures without modifying payload.
      if (resp.status === 422) {
        return resp;
      }

      // ── 429 Rate Limited ─────────────────────────────────────────────────
      if (resp.status === 429 && attempt < MAX_RETRIES) {
        const waitMs = parseRetryAfterMs(resp.headers.get("Retry-After"));
        const resumeAt = Date.now() + waitMs;
        setAllBucketsCooldown(resumeAt);
        console.log(
          `Rate limited (429) — waiting ${
            Math.ceil(waitMs / 1000)
          }s before retry ${attempt + 1}/${MAX_RETRIES}`,
        );
        await delay(waitMs);
        continue;
      }

      // ── 503 Service Unavailable ──────────────────────────────────────────
      if (resp.status === 503 && attempt < MAX_RETRIES) {
        const waitMs = Math.pow(2, attempt + 1) * 1000; // 2 s, 4 s, 8 s
        console.log(
          `Service unavailable (503) — waiting ${waitMs / 1000}s before retry ${
            attempt + 1
          }/${MAX_RETRIES}`,
        );
        await delay(waitMs);
        continue;
      }

      // ── 533 AppFolio DB Maintenance ──────────────────────────────────────
      // Typically fires between 9 PM – 4 AM PST. Back off progressively.
      if (resp.status === 533 && attempt < MAX_RETRIES) {
        const waitMs = Math.min(30_000, Math.pow(2, attempt + 2) * 1000); // 4 s, 8 s, 16 s
        console.log(
          `DB maintenance (533) — waiting ${waitMs / 1000}s before retry ${
            attempt + 1
          }/${MAX_RETRIES}`,
        );
        await delay(waitMs);
        continue;
      }

      return resp;
    } catch (err) {
      clearTimeout(timer);

      // ── AbortError = timeout ─────────────────────────────────────────────
      if (attempt < MAX_RETRIES && (err as Error).name === "AbortError") {
        console.log(
          `Timeout on attempt ${attempt + 1} for ${
            url.substring(0, 80)
          }… retrying`,
        );
        await delay(1000);
        continue;
      }

      throw err;
    }
  }

  throw new Error(
    `fetchWithTimeout: max retries exceeded for ${url.substring(0, 80)}`,
  );
}

// ── Date Helpers ──────────────────────────────────────────────────────────────
// Used by every handler that builds AppFolio date-range filters.

/** Returns an ISO date string for N days ago — e.g. "2026-01-15" */
export function daysAgo(n: number): string {
  const d = new Date();
  d.setDate(d.getDate() - n);
  return d.toISOString().split("T")[0];
}

/** Returns today's date as an ISO date string — e.g. "2026-03-21" */
export function today(): string {
  return new Date().toISOString().split("T")[0];
}

// ── corsJson helper ───────────────────────────────────────────────────────────
// Convenience wrapper used by main.ts and any handler returning raw JSON.
import { CORS_HEADERS } from "../config.ts";

export function corsJson(data: unknown, status = 200): Response {
  return new Response(JSON.stringify(data), {
    status,
    headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
  });
}

// ── snapDays — cache-key cardinality control ──────────────────────────────────
// Snaps a raw user-supplied ?days= value to the nearest canonical bucket for
// the given endpoint domain.  Prevents unbounded api_cache key growth from
// callers supplying arbitrary integers.
//
// Bucket lists are intentionally small and cover the full practical range
// for each domain.  Values outside the list snap to the nearest endpoint.
// If no bucket list exists for a domain the value is clamped to [1, 730].
const _DAYS_BUCKETS: Readonly<Record<string, readonly number[]>> = {
  work_orders: [30, 60, 90, 180, 365],
  work_orders_history: [30, 60, 90, 180, 365, 730],
  turn_work_orders: [30, 60, 90, 180],
  users: [7, 14, 30, 60, 90, 180, 365, 540, 730],
  turns: [30, 60, 90, 180],
  unit_turns: [30, 60, 90, 180, 365, 540],
  bills: [30, 90, 180, 365, 730],
  inspections: [60, 90, 180, 365],
  labor: [1, 7, 14, 30, 90],
  upcoming_moveouts: [14, 30, 60, 90],
};

/**
 * Snap `raw` days to the nearest canonical bucket for `domain`.
 * Falls back to clamping between 1 and 730 when the domain has no bucket list.
 */
export function snapDays(raw: number, domain: string): number {
  const clamped = Math.max(1, Math.round(raw));
  const buckets = _DAYS_BUCKETS[domain];
  if (!buckets || buckets.length === 0) return Math.min(clamped, 730);
  let nearest = buckets[buckets.length - 1];
  let minDist = Infinity;
  for (const b of buckets) {
    const d = Math.abs(b - clamped);
    if (d < minDist) {
      minDist = d;
      nearest = b;
    }
  }
  return nearest;
}