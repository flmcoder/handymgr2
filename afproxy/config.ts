// ============================================================================
// config.ts — Fort Lowell Realty HandyManager Proxy v9.2.2
// by Aaron Dunn
// ============================================================================

// Single source of truth for operator-visible runtime/app version.
export const PROXY_APP_VERSION = "v9.2.2";

// ── Turso ────────────────────────────────────────────────────────────────────
export const TURSO_DATABASE_URL = Deno.env.get("TURSO_DATABASE_URL") || "";
export const TURSO_AUTH_TOKEN = Deno.env.get("TURSO_AUTH_TOKEN") || "";
// Internal alias kept for compatibility with existing storage helpers
export const TURSO_URL = TURSO_DATABASE_URL;

// ── AppFolio Domains ────────────────────────────────────────────────────────
// AF_DB: universal domain for DB API v0, same for every AppFolio tenant.
// AF_REPORTS: tenant-specific subdomain — configure via AF_VHOST env var.
//   Example: AF_VHOST=acme → uses https://acme.appfolio.com for Reports API v2.
//   Defaults to "flraz" (Fort Lowell Realty) when the env var is absent.
export const AF_VHOST = Deno.env.get("AF_VHOST") || "flraz";
export const AF_REPORTS = `https://${AF_VHOST}.appfolio.com`; // Reports API v2
export const AF_DB = "https://api.appfolio.com"; // Database API v0

// ── AppFolio Credentials ─────────────────────────────────────────────────────
export const V2_CLIENT_ID = Deno.env.get("AF_V2_CLIENT_ID") || "";
export const V2_CLIENT_SECRET = Deno.env.get("AF_V2_CLIENT_SECRET") || "";
export const V0_AUTH_TOKEN = Deno.env.get("AF_V0_AUTH") || "";
export const DEV = Deno.env.get("AF_DEVELOPER_ID") || "";
// Plaintext creds are available for write-oriented flows if needed
export const AF_V0_CLIENT_ID = Deno.env.get("AF_V0_CLIENT_ID_store") || "";
export const AF_V0_CLIENT_SECRET = Deno.env.get("AF_V0_CLIENT_SECRET_store") ||
  "";

// ── Pre-computed Auth Strings (once at cold-start, not per request) ───────────
// v2: standard Basic auth — btoa(client_id:client_secret) at module load
export const AUTH_V2 = "Basic " + btoa(V2_CLIENT_ID + ":" + V2_CLIENT_SECRET);
// v0: AF_V0_AUTH is a pre-encoded Base64 token from AppFolio — prepend "Basic "
//     Falls back to AUTH_V2 if AF_V0_AUTH is unset
function normalizeBasicToken(token: string): string {
  return String(token || "").replace(/^\s*Basic\s+/i, "").trim();
}
export const AUTH_V0 = normalizeBasicToken(V0_AUTH_TOKEN)
  ? "Basic " + normalizeBasicToken(V0_AUTH_TOKEN)
  : AUTH_V2;

// ── Auth Header Factories ────────────────────────────────────────────────────
// Public helper names are kept stable because multiple handlers depend on them.
// Header key: "X-AppFolio-Developer-ID" — exact casing required by AF API

// Used by: all /api/v2/reports/ calls on AF_REPORTS domain.
// Reports API v2 and Database API v0 intentionally use different auth values.
export function reportsHeaders(): Record<string, string> {
  return {
    "Authorization": AUTH_V2,
    "X-AppFolio-Developer-ID": DEV,
    "Accept": "application/json",
    "Content-Type": "application/json",
  };
}

// Used by: all /api/v0/ calls on AF_DB domain.
export function dbHeaders(): Record<string, string> {
  return {
    "Authorization": AUTH_V0,
    "X-AppFolio-Developer-ID": DEV,
    "Accept": "application/json",
    "Content-Type": "application/json",
  };
}

// ── System / Admin ────────────────────────────────────────────────────────────
export const PROXY_ADMIN_KEY = Deno.env.get("PROXY_ADMIN_KEY") || "";

// ── Feature-specific Env Vars (add these to Val Town project settings) ──────
export const RC_CLIENT_ID = Deno.env.get("RC_CLIENT_ID") || "";
export const RC_CLIENT_SECRET = Deno.env.get("RC_CLIENT_SECRET") || "";
export const RC_JWT = Deno.env.get("RC_JWT") || "";
export const RC_FROM_NUMBER = Deno.env.get("RC_FROM_NUMBER") || "";
export const MAGIC_LINK_SECRET = Deno.env.get("MAGIC_LINK_SECRET") || "";
// CRITICAL: must be set in environment. Empty fallback forces runtime error in signMagicToken/verifyMagicToken.
export const ADMIN_NOTIFY_EMAIL = Deno.env.get("ADMIN_EMAIL") || "";
export const PROXY_BASE_URL = Deno.env.get("PROXY_BASE_URL") || "";
export const CRON_SECRET = Deno.env.get("CRON_SECRET") || "";

// ── Timing ────────────────────────────────────────────────────────────────────
// 30s timeout — bumped from 15s to handle large paginated reports
export const FETCH_TIMEOUT_MS = 30000;
// ~150ms between paginated requests — safe under AppFolio's 8 req/s limit
export const PAGE_DELAY_MS = 150;

// ── CORS ──────────────────────────────────────────────────────────────────────
// Val Town stops injecting default CORS headers the moment we set any CORS
// header ourselves — this object must be exhaustive including DELETE and PATCH
export const CORS_HEADERS: Record<string, string> = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, PUT, PATCH, DELETE, OPTIONS",
  "Access-Control-Allow-Headers":
    "Content-Type, Authorization, X-Admin-Key, X-JWS-Signature",
  "Access-Control-Max-Age": "86400",
};

// ── FLR Property Group UUIDs ────────────────────────────────────────────────
export const FLR_GROUPS: Record<string, string> = {
  Phoenix: "efe085ca-229e-11ef-bfba-069ca18f5865",
  Tucson: "a3db4460-22b3-11ef-bfba-069ca18f5865",
};

// ── Cache TTLs (minutes) ────────────────────────────────────────────────────
export const CACHE_TTL: Record<string, number> = {
  work_orders: 15,
  turns: 30,
  turn_work_orders: 15,
  vendors: 120,
  inspections: 60,
  property_groups: 360,
  properties: 360,
  property_map: 360,
  recent_tasks: 10,
  upcoming_moveouts: 30,
  labor: 10,
  bills: 30,
  wo_comparison: 60, // expensive multi-fetch — longer TTL
};

// ── Webhook → Cache Invalidation Map ────────────────────────────────────────
export const WEBHOOK_CACHE_MAP: Record<string, string[]> = {
  work_order: ["work_orders", "turn_work_orders", "recent_tasks", "labor"],
  unit_turn: ["turns", "turn_work_orders"],
  vendor: ["vendors"],
  inspection: ["inspections"],
  task: ["recent_tasks"],
  property: ["properties", "property_groups", "property_map"],
  tenant: ["upcoming_moveouts"],
  lease: ["upcoming_moveouts"],
  bill: ["bills"],
};

// ── Storage Budget ───────────────────────────────────────────────────────────
// Turso free tier = 500 MB. Val Town free plan = 10 MB total (7 MB usable).
export const STORAGE_BUDGET_BYTES = TURSO_URL ? 500_000_000 : 7 * 1024 * 1024;
export const WEBHOOK_MAX_EVENTS = 500;

// Keep webhook history longer than 30 days by default; env can override.
const WEBHOOK_MAX_DAYS_ENV = Number(Deno.env.get("WEBHOOK_MAX_DAYS") || "");
export const WEBHOOK_MAX_DAYS = Number.isFinite(WEBHOOK_MAX_DAYS_ENV) &&
    WEBHOOK_MAX_DAYS_ENV > 30
  ? Math.floor(WEBHOOK_MAX_DAYS_ENV)
  : 90;

// ── Chunked cache row limit ──────────────────────────────────────────────────
// Turso: 4 MB per row (well within 5 GB free tier, fewer chunks = faster)
// Val Town: 750 KB per row (leaves room under 1 MB SQL limit)
export const CHUNK_LIMIT = TURSO_URL ? 4_000_000 : 750_000;