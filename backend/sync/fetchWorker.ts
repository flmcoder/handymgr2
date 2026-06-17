/**
 * AppFolio-safe fetch worker.
 *
 * Wraps every outbound AppFolio HTTP call with:
 *  - Rate-gate token acquisition before dispatch
 *  - Retry-After header extraction and enforcement on 429
 *  - Capped exponential backoff with jitter on 5xx / network errors
 *  - Per-call logging to appfolio_request_log
 *  - Error classification: retryable vs terminal
 *
 * v2 cursor expiry policy:
 *   Callers fetching v2 report pages must call isCursorExpired(expiresAt)
 *   before following next_page_url. If expired, abort and restart the chain.
 *
 * Usage:
 *   const result = await afFetch('/api/v0/units?...', { headers, runId, endpointKey });
 */

import { acquireToken, setRetryAfter } from './rateLimiter.ts';
import { db } from '../db.ts';
import { appfolioRequestLog } from '../schema.ts';
import { afHeaders } from './afCredentials.ts';

export interface AfFetchOptions {
  method?: string;
  body?: string;
  runId?: string;
  endpointKey: string;
  apiVersion?: string;
  /** Current attempt number (1-based), used for logging. */
  attempt?: number;
  /** Timeout in ms. Default 45 000. */
  timeoutMs?: number;
}

export interface AfFetchResult {
  ok: boolean;
  status: number;
  data: any;
  cursorOut: string | null;  // next_page_path (v0) or next_page_url (v2)
  cursorExpiresAt: Date | null;
  latencyMs: number;
  retryAfterSeconds: number | null;
  errorText: string | null;
}

/** Error types — callers decide whether to retry based on this. */
export type AfErrorKind =
  | 'retryable_5xx'
  | 'retryable_network'
  | 'rate_limited_429'
  | 'auth_error'
  | 'client_error_terminal'
  | 'schema_error';

export class AfFetchError extends Error {
  constructor(
    public readonly kind: AfErrorKind,
    message: string,
    public readonly status?: number,
  ) {
    super(message);
    this.name = 'AfFetchError';
  }
}

const DEFAULT_TIMEOUT_MS = 45_000;
const MAX_RETRIES        = 4;
const BACKOFF_BASE_MS    = 1_000;
const BACKOFF_MAX_MS     = 30_000;
// v2 page URLs expire after 30 min; reject if < 90s remain to leave headroom.
const V2_CURSOR_EXPIRY_HEADROOM_MS = 90_000;

/**
 * Returns true if a v2 next_page_url should be considered expired.
 * Always call this before following a v2 cursor.
 */
export function isCursorExpired(expiresAt: Date | null): boolean {
  if (!expiresAt) return false;
  return expiresAt.getTime() - Date.now() < V2_CURSOR_EXPIRY_HEADROOM_MS;
}

/**
 * Core rate-gated fetch with retry/backoff.
 * baseUrl must be the full path (e.g. 'https://api.appfolio.com/api/v0/units?...')
 */
export async function afFetch(
  url: string,
  opts: AfFetchOptions,
): Promise<AfFetchResult> {
  const {
    method = 'GET',
    body,
    runId,
    endpointKey,
    apiVersion = url.includes('/api/v2/') ? 'v2' : 'v0',
    attempt = 1,
    timeoutMs = DEFAULT_TIMEOUT_MS,
  } = opts;

  let lastError: AfFetchError | null = null;

  for (let tryNum = 1; tryNum <= MAX_RETRIES; tryNum++) {
    // Acquire rate-limit token — waits if windows are full or 429 block is active.
    await acquireToken();

    const started = Date.now();
    let status = 0;
    let retryAfterSeconds: number | null = null;

    try {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), timeoutMs);

      let res: Response;
      try {
        res = await fetch(url, {
          method,
          headers: afHeaders(apiVersion),
          body: body ?? undefined,
          signal: controller.signal,
        });
      } finally {
        clearTimeout(timer);
      }

      const latencyMs = Date.now() - started;
      status = res.status;
      retryAfterSeconds = res.headers.has('Retry-After')
        ? parseInt(res.headers.get('Retry-After')!, 10)
        : null;

      // Log every call regardless of outcome.
      await logRequest({
        runId, endpointKey, apiVersion, method, status,
        latencyMs, retryAfterSeconds, attemptNumber: tryNum,
        cursorSnapshot: url.split('?')[1]?.slice(0, 200) ?? null,
      });

      // 429 — honor Retry-After exactly.
      if (status === 429) {
        const waitSec = retryAfterSeconds ?? 60;
        setRetryAfter(waitSec);
        lastError = new AfFetchError('rate_limited_429', `429 rate limited; Retry-After=${waitSec}s`, 429);
        if (tryNum < MAX_RETRIES) continue;
        throw lastError;
      }

      // Auth errors — terminal, do not retry.
      if (status === 401 || status === 403) {
        throw new AfFetchError('auth_error', `Auth error HTTP ${status}`, status);
      }

      // Other 4xx — terminal.
      if (status >= 400 && status < 500) {
        throw new AfFetchError('client_error_terminal', `Client error HTTP ${status}`, status);
      }

      // 5xx — retryable with backoff.
      if (status >= 500) {
        lastError = new AfFetchError('retryable_5xx', `Server error HTTP ${status}`, status);
        if (tryNum < MAX_RETRIES) {
          await jitterBackoff(tryNum);
          continue;
        }
        throw lastError;
      }

      // Success — parse body.
      let data: any;
      try {
        data = await res.json();
      } catch (parseErr) {
        throw new AfFetchError('schema_error', `JSON parse failed: ${String(parseErr)}`);
      }

      // Extract cursor. v0 uses next_page_path; v2 uses next_page_url.
      const cursorOut: string | null =
        data?.next_page_url ?? data?.next_page_path ?? null;

      // v2 cursors expire ~30 min after generation. Approximate expiry from now.
      const cursorExpiresAt: Date | null =
        cursorOut && apiVersion === 'v2'
          ? new Date(Date.now() + 28 * 60 * 1_000) // 28 min conservative
          : null;

      return {
        ok: true,
        status,
        data,
        cursorOut,
        cursorExpiresAt,
        latencyMs,
        retryAfterSeconds,
        errorText: null,
      };

    } catch (err) {
      const latencyMs = Date.now() - started;

      if (err instanceof AfFetchError) throw err;

      // Network / timeout errors — retryable.
      const isAbort = (err as any)?.name === 'AbortError';
      const msg = isAbort
        ? `Request timed out after ${timeoutMs}ms`
        : String((err as any)?.message ?? err);

      lastError = new AfFetchError('retryable_network', msg);

      await logRequest({
        runId, endpointKey, apiVersion, method, status,
        latencyMs, retryAfterSeconds, attemptNumber: tryNum, errorText: msg,
        cursorSnapshot: url.split('?')[1]?.slice(0, 200) ?? null,
      });

      if (tryNum < MAX_RETRIES) {
        await jitterBackoff(tryNum);
        continue;
      }
      throw lastError;
    }
  }

  throw lastError ?? new AfFetchError('retryable_network', 'Exhausted retries');
}

// ── Helpers ──────────────────────────────────────────────────────────────────

async function jitterBackoff(attempt: number): Promise<void> {
  const base = Math.min(BACKOFF_MAX_MS, BACKOFF_BASE_MS * Math.pow(2, attempt - 1));
  const jitter = Math.floor(Math.random() * 1_000);
  await new Promise((r) => setTimeout(r, base + jitter));
}

async function logRequest(entry: {
  runId?: string;
  endpointKey: string;
  apiVersion?: string;
  method?: string;
  status?: number;
  latencyMs?: number;
  retryAfterSeconds?: number | null;
  attemptNumber?: number;
  errorText?: string;
  cursorSnapshot?: string | null;
}): Promise<void> {
  try {
    await db.insert(appfolioRequestLog).values({
      runId: entry.runId ?? null,
      endpointKey: entry.endpointKey,
      apiVersion: entry.apiVersion ?? null,
      method: entry.method ?? 'GET',
      statusCode: entry.status ?? null,
      latencyMs: entry.latencyMs ?? null,
      retryAfterSeconds: entry.retryAfterSeconds ?? null,
      attemptNumber: entry.attemptNumber ?? 1,
      errorText: entry.errorText ?? null,
      cursorSnapshot: entry.cursorSnapshot ?? null,
    });
  } catch {
    // Log failures must never propagate to the sync worker.
  }
}
