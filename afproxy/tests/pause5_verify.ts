// ============================================================================
// tests/pause5_verify.ts — PAUSE-5 verification checklist runner.
//
// Verifies:
//   1) Link action writes association and reports cache invalidation.
//   2) Unlink action marks association removed and reports cache invalidation.
//   3) Optional terminal-turn lock returns 409 for link/unlink attempts.
//
// Usage:
//   API_PROXY="https://<proxy-url>" \
//   PROXY_BEARER_TOKEN="<optional-token>" \
//   TURN_KEY="<turn-key>" \
//   WO_ID="<work-order-number>" \
//   WO_DB_UUID="<optional-uuid>" \
//   TERMINAL_TURN_KEY="<optional-terminal-turn-key>" \
//   deno run --allow-env --allow-net tests/pause5_verify.ts
// ============================================================================

type Json = Record<string, unknown>;

const API_PROXY = String(Deno.env.get("API_PROXY") || "").trim();
const TOKEN = String(Deno.env.get("PROXY_BEARER_TOKEN") || "").trim();
const TURN_KEY = String(Deno.env.get("TURN_KEY") || "").trim();
const WO_ID = String(Deno.env.get("WO_ID") || "").trim();
const WO_DB_UUID = String(Deno.env.get("WO_DB_UUID") || "").trim();
const TERMINAL_TURN_KEY = String(Deno.env.get("TERMINAL_TURN_KEY") || "").trim();

if (!API_PROXY) {
  console.error("Missing API_PROXY env var.");
  Deno.exit(1);
}
if (!TURN_KEY || !WO_ID) {
  console.error("Missing TURN_KEY or WO_ID env vars.");
  Deno.exit(1);
}

function authHeaders(contentType = false): HeadersInit {
  const headers: Record<string, string> = { Accept: "application/json" };
  if (contentType) headers["Content-Type"] = "application/json";
  if (TOKEN) headers.Authorization = `Bearer ${TOKEN}`;
  return headers;
}

function buildUrl(action: string): string {
  const url = new URL(API_PROXY);
  url.searchParams.set("action", action);
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

async function postAction(action: string, payload: Json): Promise<{ resp: Response; body: Json }> {
  const resp = await fetch(buildUrl(action), {
    method: "POST",
    headers: authHeaders(true),
    body: JSON.stringify(payload),
  });
  const body = await parseJsonSafe(resp);
  return { resp, body };
}

function assertCacheSignal(body: Json, context: string): void {
  if (!("cache_invalidated" in body)) {
    throw new Error(`${context}: missing cache_invalidated signal body=${JSON.stringify(body)}`);
  }
}

async function verifyLinkUnlinkRoundTrip(): Promise<void> {
  const basePayload: Json = {
    turn_key: TURN_KEY,
    wo_id: WO_ID,
  };
  if (WO_DB_UUID) basePayload.wo_db_uuid = WO_DB_UUID;

  const linkRes = await postAction("unit_turn_wo_link", basePayload);
  if (!linkRes.resp.ok || linkRes.body.ok === false) {
    throw new Error(
      `Link failed: HTTP ${linkRes.resp.status} body=${JSON.stringify(linkRes.body)}`,
    );
  }
  assertCacheSignal(linkRes.body, "Link");
  console.log(
    `Link: PASS cache_invalidated=${String(linkRes.body.cache_invalidated)} tracking_uuid=${String(linkRes.body.tracking_uuid || "")}`,
  );

  const unlinkRes = await postAction("unit_turn_wo_unlink", basePayload);
  if (!unlinkRes.resp.ok || unlinkRes.body.ok === false) {
    throw new Error(
      `Unlink failed: HTTP ${unlinkRes.resp.status} body=${JSON.stringify(unlinkRes.body)}`,
    );
  }
  assertCacheSignal(unlinkRes.body, "Unlink");
  console.log(
    `Unlink: PASS cache_invalidated=${String(unlinkRes.body.cache_invalidated)} tracking_uuid=${String(unlinkRes.body.tracking_uuid || "")}`,
  );
}

async function verifyTerminalLockIfProvided(): Promise<void> {
  if (!TERMINAL_TURN_KEY) {
    console.log("Terminal Lock: SKIP (set TERMINAL_TURN_KEY to enforce 409 check)");
    return;
  }

  const payload: Json = { turn_key: TERMINAL_TURN_KEY, wo_id: WO_ID };
  if (WO_DB_UUID) payload.wo_db_uuid = WO_DB_UUID;

  const linkRes = await postAction("unit_turn_wo_link", payload);
  if (linkRes.resp.status !== 409 || String(linkRes.body.error || "") !== "unit_turn_locked") {
    throw new Error(
      `Terminal lock (link) failed: expected 409 unit_turn_locked, got HTTP ${linkRes.resp.status} body=${JSON.stringify(linkRes.body)}`,
    );
  }

  const unlinkRes = await postAction("unit_turn_wo_unlink", payload);
  if (unlinkRes.resp.status !== 409 || String(unlinkRes.body.error || "") !== "unit_turn_locked") {
    throw new Error(
      `Terminal lock (unlink) failed: expected 409 unit_turn_locked, got HTTP ${unlinkRes.resp.status} body=${JSON.stringify(unlinkRes.body)}`,
    );
  }

  console.log("Terminal Lock: PASS link/unlink correctly blocked with 409");
}

if (import.meta.main) {
  console.log("PAUSE-5 Verification Checklist Start");
  await verifyLinkUnlinkRoundTrip();
  await verifyTerminalLockIfProvided();
  console.log("PAUSE-5 Verification Checklist PASS");
}
