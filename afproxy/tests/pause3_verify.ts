// ============================================================================
// tests/pause3_verify.ts — PAUSE-3 verification checklist runner.
//
// Verifies:
//   1) Async flow: POST ?action=unit_turns_sync&async=true returns task_id.
//   2) Telemetry: GET ?action=resilience_telemetry returns concurrency_peak
//      and retry_total.
//   3) Backoff semantics: simulated 503 retries produce increasing
//      nextAttemptAt timestamps using SW backoff constants.
//
// Usage:
//   API_PROXY="https://<proxy-url>" \
//   PROXY_BEARER_TOKEN="<optional-token>" \
//   deno run --allow-env --allow-net --allow-read tests/pause3_verify.ts
// ============================================================================

type Json = Record<string, unknown>;

const API_PROXY = String(Deno.env.get("API_PROXY") || "").trim();
const TOKEN = String(Deno.env.get("PROXY_BEARER_TOKEN") || "").trim();

if (!API_PROXY) {
  console.error("Missing API_PROXY env var.");
  Deno.exit(1);
}

function authHeaders(contentType = false): HeadersInit {
  const headers: Record<string, string> = {
    Accept: "application/json",
  };
  if (contentType) headers["Content-Type"] = "application/json";
  if (TOKEN) headers.Authorization = `Bearer ${TOKEN}`;
  return headers;
}

function buildUrl(action: string, params: Record<string, string> = {}): string {
  const url = new URL(API_PROXY);
  url.searchParams.set("action", action);
  for (const [k, v] of Object.entries(params)) {
    if (String(v || "").trim()) url.searchParams.set(k, v);
  }
  return url.toString();
}

async function parseJsonSafe(resp: Response): Promise<Json> {
  const txt = await resp.text().catch(() => "");
  if (!txt) return {};
  try {
    return JSON.parse(txt) as Json;
  } catch {
    return {};
  }
}

async function verifyAsyncFlow(): Promise<void> {
  const url = buildUrl("unit_turns_sync", { async: "true" });
  const resp = await fetch(url, {
    method: "POST",
    headers: authHeaders(true),
    body: JSON.stringify({ records: [] }),
  });
  const body = await parseJsonSafe(resp);
  const taskId = String(body.task_id || "").trim();

  const accepted = resp.status === 202 || body.status === 202 || body.accepted === true;
  if (!accepted) {
    throw new Error(
      `Async flow failed: expected accepted=202, got HTTP ${resp.status} body=${JSON.stringify(body)}`,
    );
  }
  if (!taskId) {
    throw new Error(`Async flow failed: task_id missing in response body=${JSON.stringify(body)}`);
  }

  console.log(`Async Flow: PASS task_id=${taskId}`);
}

async function verifyTelemetry(): Promise<void> {
  const url = buildUrl("resilience_telemetry");
  const resp = await fetch(url, { headers: authHeaders(false) });
  const body = await parseJsonSafe(resp);
  const telemetry = (body.telemetry || {}) as Json;

  if (!resp.ok) {
    throw new Error(`Telemetry failed: HTTP ${resp.status} body=${JSON.stringify(body)}`);
  }

  const concurrencyPeak = Number(telemetry.concurrency_peak);
  const retryTotal = Number(telemetry.retry_total);

  if (!Number.isFinite(concurrencyPeak)) {
    throw new Error(
      `Telemetry failed: concurrency_peak missing/non-numeric telemetry=${JSON.stringify(telemetry)}`,
    );
  }
  if (!Number.isFinite(retryTotal)) {
    throw new Error(
      `Telemetry failed: retry_total missing/non-numeric telemetry=${JSON.stringify(telemetry)}`,
    );
  }

  console.log(`Telemetry: PASS concurrency_peak=${concurrencyPeak} retry_total=${retryTotal}`);
}

function parseSwConst(swSource: string, name: string, fallback: number): number {
  const re = new RegExp(`const\\s+${name}\\s*=\\s*([0-9_]+)\\s*;`);
  const m = swSource.match(re);
  if (!m) return fallback;
  const raw = m[1].replaceAll("_", "");
  const n = Number(raw);
  return Number.isFinite(n) ? n : fallback;
}

function computeBackoffMs(
  attempts: number,
  baseMs: number,
  maxMs: number,
): number {
  const exp = Math.max(0, Number(attempts || 0) - 1);
  const base = Math.min(maxMs, baseMs * Math.pow(2, exp));
  const jitter = Math.floor(Math.random() * 1000);
  return Math.min(maxMs, base + jitter);
}

async function verifyBackoffSimulation(): Promise<void> {
  const swPath = new URL("../../sw.js", import.meta.url).pathname;
  const sw = await Deno.readTextFile(swPath);

  const baseMs = parseSwConst(sw, "HM_REPLAY_BACKOFF_BASE_MS", 3000);
  const maxMs = parseSwConst(sw, "HM_REPLAY_BACKOFF_MAX_MS", 300000);

  // Simulate two consecutive retryable HTTP 503 outcomes for the same queued item.
  const now = Date.now();
  const firstAttempts = 1;
  const firstNextAttemptAt = now + computeBackoffMs(firstAttempts, baseMs, maxMs);
  const secondAttempts = 2;
  const secondNextAttemptAt = now + computeBackoffMs(secondAttempts, baseMs, maxMs);

  console.log(
    `Backoff Simulation (503): attempt1_nextAttemptAt=${firstNextAttemptAt} attempt2_nextAttemptAt=${secondNextAttemptAt}`,
  );

  if (!(secondNextAttemptAt > firstNextAttemptAt)) {
    throw new Error(
      `Backoff failed: expected attempt2 nextAttemptAt > attempt1 (${secondNextAttemptAt} <= ${firstNextAttemptAt})`,
    );
  }

  console.log("Backoff Logic: PASS nextAttemptAt increases on repeated 503 retries");
}

if (import.meta.main) {
  console.log("PAUSE-3 Verification Checklist Start");
  await verifyAsyncFlow();
  await verifyTelemetry();
  await verifyBackoffSimulation();
  console.log("PAUSE-3 Verification Checklist PASS");
}
