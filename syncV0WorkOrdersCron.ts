// syncV0WorkOrdersCron.ts — Triggers the v0 Database API → Turso sync
// as a background task on the proxy. Returns immediately (202 Accepted)
// so the cron val doesn't time out on Val Town's 1-minute free-tier limit.
declare const Deno: {
  env: {
    get(name: string): string | undefined;
  };
};

const PRIMARY_PROXY_URL = (Deno.env.get("HANDYMGR_PROXY_URL") || "").trim();
const INTERNAL_SYNC_TOKEN = (Deno.env.get("HANDYMGR_INTERNAL_TOKEN") || "").trim();
const MS_PER_DAY = 86_400_000;
const DEFAULT_LOOKBACK_DAYS = Math.max(
  1,
  Math.min(365, parseInt(Deno.env.get("HANDYMGR_SYNC_LOOKBACK_DAYS") || "180", 10)),
);
const REQUEST_TIMEOUT_MS = Math.max(
  5_000,
  Math.min(60_000, parseInt(Deno.env.get("HANDYMGR_SYNC_TIMEOUT_MS") || "30000", 10)),
);
const MAX_RETRIES = Math.max(1, Math.min(5, parseInt(Deno.env.get("HANDYMGR_SYNC_RETRIES") || "3", 10)));

if (!PRIMARY_PROXY_URL) {
  throw new Error("HANDYMGR_PROXY_URL env var is required for sync val");
}

if (!INTERNAL_SYNC_TOKEN) {
  throw new Error("HANDYMGR_INTERNAL_TOKEN env var is required for sync val");
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function fetchWithTimeout(url: string, init: RequestInit): Promise<Response> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
  try {
    return await fetch(url, { ...init, signal: controller.signal });
  } finally {
    clearTimeout(timeout);
  }
}

async function triggerSync(): Promise<any> {
  // Use v0 sync endpoint (preferred — paginates DB API directly to Turso)
  const target = new URL(PRIMARY_PROXY_URL);
  target.searchParams.set("action", "sync-v0-work-orders");
  const fromDate = new Date(Date.now() - (DEFAULT_LOOKBACK_DAYS * MS_PER_DAY))
    .toISOString()
    .slice(0, 10);
  target.searchParams.set("from_date", fromDate);

  let attempt = 0;
  let lastErr: unknown = undefined;

  while (attempt < MAX_RETRIES) {
    attempt++;
    try {
      const res = await fetchWithTimeout(target.toString(), {
        method: "POST",
        headers: {
          "x-cron-secret": INTERNAL_SYNC_TOKEN,
          "Content-Type": "application/json",
        },
      });

      const text = await res.text();
      const data = text ? JSON.parse(text) : {};

      // 202 Accepted = background task started, that's a success
      if (res.status === 202 || res.ok) {
        return {
          ok: true,
          accepted: res.status === 202,
          attempt,
          ...data,
        };
      }

      // 429 = back off and retry
      if (res.status === 429) {
        const ra = parseInt(res.headers.get("Retry-After") || "10", 10);
        console.warn(`[sync] 429 — waiting ${ra}s before retry ${attempt}`);
        await delay(ra * 1000);
        continue;
      }

      const err = new Error(
        `[sync-v0-work-orders] ${res.status} ${res.statusText}: ${text.slice(0, 200)}`,
      );
      throw err;
    } catch (err) {
      lastErr = err;
      const backoff = Math.min(30_000, 2 ** attempt * 500);
      console.warn(`sync-v0-work-orders attempt ${attempt} failed`, err);
      if (attempt >= MAX_RETRIES) break;
      await delay(backoff);
    }
  }

  throw lastErr instanceof Error
    ? lastErr
    : new Error(String(lastErr || "sync failed"));
}

export default async function syncV0WorkOrdersCron(): Promise<any> {
  const startedAt = new Date().toISOString();
  const result = await triggerSync();
  const finishedAt = new Date().toISOString();

  console.log(
    "[sync-v0-work-orders]",
    JSON.stringify({ startedAt, finishedAt, ...result }),
  );

  return result;
}
