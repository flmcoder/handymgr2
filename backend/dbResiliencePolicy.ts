/**
 * Resilience helpers for a database reached through an SSH/localtonet tunnel.
 * Kept free of driver imports so retry classification and backoff are unit-testable.
 */

const TRANSIENT_NODE_CODES = new Set([
  'CONNECT_TIMEOUT',
  'CONNECTION_CLOSED',
  'CONNECTION_DESTROYED',
  'CONNECTION_ENDED',
  'ECONNRESET',
  'ECONNREFUSED',
  'ECONNABORTED',
  'ETIMEDOUT',
  'EPIPE',
  'EHOSTUNREACH',
  'ENETUNREACH',
  'EAI_AGAIN',
]);

// SQLSTATEs that are safe to retry: server restarts/shutdowns, connection
// exhaustion, serialization failures and deadlocks.
const TRANSIENT_SQLSTATES = new Set(['57P01', '57P02', '57P03', '53300', '40001', '40P01']);

function asRecord(value: unknown): Record<string, unknown> | null {
  return value !== null && typeof value === 'object' ? value as Record<string, unknown> : null;
}

/** True when the error (or any wrapped cause) is a connection-level or otherwise retryable failure. */
export function isTransientDbError(error: unknown): boolean {
  const seen = new Set<unknown>();
  let current = asRecord(error);
  while (current && !seen.has(current)) {
    seen.add(current);
    const code = String(current.code ?? '');
    if (TRANSIENT_NODE_CODES.has(code) || TRANSIENT_SQLSTATES.has(code) || code.startsWith('08')) return true;
    const message = String(current.message ?? '');
    if (/terminating connection|connection terminated|connection closed|server closed the connection|socket hang up/i.test(message)) {
      return true;
    }
    current = asRecord(current.cause);
  }
  return false;
}

export type BackoffOptions = {
  baseMs?: number;
  maxMs?: number;
  /** Fraction of the delay added as random jitter (0 disables). */
  jitter?: number;
  random?: () => number;
};

/** Exponential backoff in milliseconds for a 1-based attempt number, capped at `maxMs`. */
export function computeDbBackoffMs(attempt: number, options: BackoffOptions = {}): number {
  const baseMs = Math.max(1, options.baseMs ?? 500);
  const maxMs = Math.max(baseMs, options.maxMs ?? 15_000);
  const safeAttempt = Math.max(1, Math.floor(Number.isFinite(attempt) ? attempt : 1));
  const delay = Math.min(baseMs * 2 ** (safeAttempt - 1), maxMs);
  const jitter = Math.max(0, options.jitter ?? 0);
  if (!jitter) return delay;
  const random = options.random ?? Math.random;
  return Math.min(maxMs, Math.round(delay + delay * jitter * random()));
}

export type DbRetryOptions = BackoffOptions & {
  /** Total attempts including the first. */
  attempts?: number;
  label?: string;
  sleep?: (ms: number) => Promise<void>;
  isRetryable?: (error: unknown) => boolean;
  onRetry?: (info: { attempt: number; delayMs: number; error: unknown; label: string }) => void;
};

/**
 * Runs `operation`, retrying transient connection errors with exponential
 * backoff. Non-transient errors (constraint violations, bad SQL) are thrown
 * immediately, as is the last transient error once attempts are exhausted.
 * `operation` must be idempotent.
 */
export async function withDbRetry<T>(operation: () => Promise<T>, options: DbRetryOptions = {}): Promise<T> {
  const attempts = Math.max(1, Math.floor(options.attempts ?? 4));
  const isRetryable = options.isRetryable ?? isTransientDbError;
  const sleep = options.sleep ?? ((ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms)));
  const label = options.label ?? 'db';

  for (let attempt = 1; ; attempt++) {
    try {
      return await operation();
    } catch (error) {
      if (attempt >= attempts || !isRetryable(error)) throw error;
      const delayMs = computeDbBackoffMs(attempt, { jitter: 0.25, ...options });
      options.onRetry?.({ attempt, delayMs, error, label });
      await sleep(delayMs);
    }
  }
}

function envNumber(env: Record<string, string | undefined>, name: string, fallback: number, min: number): number {
  const parsed = Number(env[name]);
  return Number.isFinite(parsed) && String(env[name] ?? '').trim() !== '' ? Math.max(min, parsed) : fallback;
}

export type PostgresResilienceOptions = {
  max: number;
  idle_timeout: number;
  connect_timeout: number;
  max_lifetime: number;
  keep_alive: number;
  backoff: (retries: number) => number;
};

/**
 * Connection-pool settings tuned for a tunnel that periodically drops idle
 * sockets. All durations are seconds (postgres.js units).
 *  - keep_alive: TCP keep-alive probes so idle tunnel sockets are not silently reaped.
 *  - idle_timeout < tunnel idle limit: pool recycles sockets before the tunnel kills them.
 *  - max_lifetime: periodic recycling so half-open sockets cannot linger.
 *  - backoff: exponential reconnect delay (seconds), capped.
 */
export function buildPostgresResilienceOptions(
  env: Record<string, string | undefined> = process.env,
): PostgresResilienceOptions {
  const backoffBaseMs = envNumber(env, 'DB_BACKOFF_BASE_MS', 500, 50);
  const backoffMaxMs = envNumber(env, 'DB_BACKOFF_MAX_MS', 15_000, backoffBaseMs);
  return {
    max: Math.floor(envNumber(env, 'DB_POOL_MAX', 10, 1)),
    idle_timeout: envNumber(env, 'DB_IDLE_TIMEOUT', 20, 0),
    connect_timeout: envNumber(env, 'DB_CONNECT_TIMEOUT', 15, 1),
    max_lifetime: envNumber(env, 'DB_MAX_LIFETIME', 1_800, 60),
    keep_alive: envNumber(env, 'DB_KEEP_ALIVE', 30, 1),
    backoff: (retries) => computeDbBackoffMs(retries + 1, { baseMs: backoffBaseMs, maxMs: backoffMaxMs, jitter: 0.25 }) / 1000,
  };
}
