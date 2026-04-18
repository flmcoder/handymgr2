// ============================================================================
// main.ts — HandyManager Proxy v9.2.2
// Fort Lowell Realty & Property Management
// Val Town HTTP Entry Point — Router ONLY.
//
// This file contains zero business logic. Every action is delegated to a
// handler module. Adding a new endpoint = one import + one switch case.
//
// AppFolio API constraints enforced across all handlers:
//   • AssignedUsers must reference a user with the Maintenance Tech role —
//     any other role returns 422 "User not found"
//   • Simultaneous PATCH requests to the same resource will cause the second
//     request to fail — all cron writes are sequential
//   • Note body field key is "Body" (capital B)
//   • All requests and responses use Content-Type: application/json
//   • 503 (unavailable) and 533 (AF maintenance) are retried automatically
//   • Semantic errors return 422 — ensure parameters contain valid values
//   • All date-time values must follow ISO 8601 YYYY-MM-DDTHH:mm:ssZ
//
// Val Town Cron Vals (create separately — see bottom of this file):
//   Pass A — Noon Warning:      schedule 0 19 * * *  (12:00 PM AZ / 19:00 UTC)
//   Pass B — Midnight Reassign: schedule 0 7  * * *  (12:00 AM AZ / 07:00 UTC)
//   Arizona does not observe DST — UTC offset is always -7 year-round.
// ============================================================================

import { ensureTables, rowsAsObjects, sqlite } from "./db.ts";
import {
  AF_DB,
  AF_REPORTS,
  CORS_HEADERS,
  CRON_SECRET,
  DEV,
  dbHeaders,
  MAGIC_LINK_SECRET,
  PROXY_APP_VERSION,
  PROXY_ADMIN_KEY,
  reportsHeaders,
  V2_CLIENT_ID,
  V2_CLIENT_SECRET,
} from "./config.ts";
import { resolveShortLink } from "./lib/auth.ts";
import { checkRateLimit } from "./lib/rateLimit.ts";

// ── Core data handlers ───────────────────────────────────────────────────────
import {
  handleCompletedWorkOrdersHistory,
  handleLabor,
  handleRecentTasks,
  handleTurnWorkOrders,
  handleWoBilledAmount,
  handleWoDetail,
  handleWoNoteCreate,
  handleWoNotes,
  handleWorkOrders,
} from "./handlers/workOrders.ts";
import { handleVendorOverride, handleVendors } from "./handlers/vendors.ts";
import { handleInspections } from "./handlers/inspections.ts";
import { handleBills, handleBillsRoute } from "./handlers/bills.ts";
import {
  handleProperties,
  handlePropertyGroups,
  handlePropertyMap,
  handleUpcomingMoveouts,
} from "./handlers/properties.ts";
import {
  handleAdminSyncRoute,
  handleMigrateV8,
  handleWebhookCleanupEndpoint,
  handleWebhookEvents,
  handleWebhookFeed,
  handleWebhookLive, // ← ADD THIS
  handleWebhookMigrate,
  handleWebhookPost,
  handleWebhookResolve,
  handleWebhookStats,
} from "./handlers/webhook.ts";
import {
  handleClosedTurns,
  handleTurnRecords,
  handleTurnRecordStage,
  handleTurns,
  handleTurnsIncremental,
  handleUnitTurns,
  handleUnitTurnsHistory,
  handleUnitTurnsSync,
  handleUnitTurnWorkOrderLink,
  handleUnitTurnWorkOrderUnlink,
} from "./handlers/turns.ts";
import { handleWoComparisonReport } from "./handlers/woComparison.ts";
import {
  handleCacheInvalidate,
  handleCacheStats,
  handleDebugSqlite,
  handleForceRefresh,
  handleGenericReport,
  handlePassthrough,
  handlePing,
  handleSqlExecute,
  handleSqlQuery,
  handleStorageCleanup,
} from "./handlers/passthrough.ts";

// ── v9.2.2 Reassignment engine handlers ──────────────────────────────────────
import {
  handleMidnightReassignCron,
  handleNoonWarningCron,
} from "./handlers/reassignment.ts";
import { handleMagicPortal } from "./handlers/magicPortal.ts";
import {
  handleAddMonitoredWO,
  handleGenerateMagicLink,
  handlePortalNoContact,
  handlePortalNote,
  handlePortalReassignRequest,
  handlePortalReschedule,
  handlePortalSchedule,
  handlePortalValidate,
  handleRemoveMonitoredWO,
  handleSendMagicLinkTestSMS,
  handleSendTenantSMS,
  handleTenantCommsLog,
} from "./handlers/tenantComms.ts";
import {
  handleDispatchSeedReassignmentTest,
  handleDispatchSyncAssignees,
  handleTechRoster,
} from "./handlers/techRoster.ts";
import { handleUsers } from "./handlers/users.ts";
import { handleReassignmentQueue } from "./handlers/queue.ts";
import {
  getTrustedDeviceSession,
  handleDeviceOtpRequest,
  handleDeviceOtpVerify,
  handleDeviceSetup,
  handlePmProxyUserDelete,
  handlePmProxyUserUpsert,
  handleTrustedDeviceList,
  handleTrustedDeviceRevoke,
  handleVerifyRole,
  touchDeviceSession,
} from "./handlers/deviceAuth.ts";

const PROXY_VERSION = PROXY_APP_VERSION;


function jsonResp(data: unknown, status = 200): Response {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "Content-Type": "application/json", ...CORS_HEADERS },
  });
}

// ============================================================================
// CRON SECRET GUARD
// Applied to any action in CRON_ACTIONS before the main switch runs.
// Checks both ?secret= URL param and x-cron-secret header.
// ============================================================================
const CRON_ACTIONS = new Set([
  "noon_warning_cron",
  "midnight_reassign_cron",
]);

const FRONTEND_PROXY_SECRET = Deno.env.get("FRONTEND_PROXY_SECRET") || "";
// Emergency override: set to false to allow all frontend requests without bearer auth.
// Keep enabled for normal operation.
const ENFORCE_FRONTEND_AUTH = true;

function extractFrontendToken(headers: Headers): string {
  const auth = headers.get("authorization") || "";
  const bearer = auth.toLowerCase().startsWith("bearer ")
    ? auth.slice(7).trim()
    : "";
  const alt = (headers.get("x-proxy-token") || "").trim();
  return (bearer || alt || "").trim();
}

async function isTrustedDeviceToken(token: string): Promise<boolean> {
  if (!token) return false;
  try {
    return !!(await getTrustedDeviceSession(token));
  } catch {
    return false;
  }
}

async function isFrontendAuthorized(headers: Headers): Promise<boolean> {
  const token = extractFrontendToken(headers);
  if (!token) return false;
  if (FRONTEND_PROXY_SECRET && token === FRONTEND_PROXY_SECRET) return true;
  return await isTrustedDeviceToken(token);
}

function isCronAuthorized(
  params: Record<string, string>,
  headers: Headers,
): boolean {
  // CRITICAL: CRON_SECRET must be set and match. Don't allow open access.
  if (!CRON_SECRET) return false;
  const supplied = params.secret ||
    headers.get("x-cron-secret") ||
    "";
  return supplied === CRON_SECRET;
}

function safeParseJSON<T>(raw: string | null | undefined, fallback: T): T {
  if (!raw) return fallback;
  try {
    return JSON.parse(raw) as T;
  } catch {
    return fallback;
  }
}

type ProxySettingRow = {
  key: string;
  value: string;
  updated_at: string;
};

function isLikelyMissingSettingsTableError(err: unknown): boolean {
  const msg = String((err as any)?.message || err || "").toLowerCase();
  return msg.includes("no such table") ||
    msg.includes("does not exist") ||
    (msg.includes("sql_input_error") && msg.includes("table"));
}

function isLikelyMissingUpdatedAtError(err: unknown): boolean {
  const msg = String((err as any)?.message || err || "").toLowerCase();
  return msg.includes("no such column") && msg.includes("updated_at");
}

async function readSettingsTable(tableName: "proxy_config" | "app_settings"): Promise<ProxySettingRow[]> {
  const mapRows = (result: any) =>
    rowsAsObjects(result).map((row: any) => ({
      key: String(row.key ?? ""),
      value: String(row.value ?? ""),
      updated_at: String(row.updated_at ?? ""),
    }));

  try {
    const result = await sqlite.execute({
      sql: `SELECT key, value, COALESCE(updated_at, '') AS updated_at
            FROM ${tableName}
            ORDER BY key`,
    });
    return mapRows(result);
  } catch (err: any) {
    if (isLikelyMissingSettingsTableError(err)) return [];
    if (!isLikelyMissingUpdatedAtError(err)) throw err;
  }

  try {
    const result = await sqlite.execute({
      sql: `SELECT key, value, '' AS updated_at
            FROM ${tableName}
            ORDER BY key`,
    });
    return mapRows(result);
  } catch (err: any) {
    if (isLikelyMissingSettingsTableError(err)) return [];
    const msg = String(err?.message || err || "").toLowerCase();
    if (msg.includes("no such column") && (msg.includes("key") || msg.includes("value"))) {
      return [];
    }
    throw err;
  }
}

async function upsertSettingTable(
  tableName: "proxy_config" | "app_settings",
  key: string,
  value: string,
): Promise<void> {
  const tableInfo = rowsAsObjects(
    await sqlite.execute(`PRAGMA table_info(${tableName})`),
  );
  if (!tableInfo.length) return;

  const hasUpdatedAt = tableInfo.some((col: any) => String(col.name || "") === "updated_at");
  if (hasUpdatedAt) {
    await sqlite.execute({
      sql: `INSERT INTO ${tableName} (key, value, updated_at)
            VALUES (?, ?, datetime('now'))
            ON CONFLICT(key) DO UPDATE SET
              value = excluded.value,
              updated_at = datetime('now')`,
      args: [key, value],
    });
    return;
  }

  await sqlite.execute({
    sql: `INSERT INTO ${tableName} (key, value)
          VALUES (?, ?)
          ON CONFLICT(key) DO UPDATE SET
            value = excluded.value`,
    args: [key, value],
  });
}

// ============================================================================
// MAIN EXPORT — Val Town HTTP handler entry point
// ============================================================================
export default async function handler(req: Request): Promise<Response> {
  // ── CORS preflight — always respond immediately, no DB needed ─────────────
  if (req.method === "OPTIONS") {
    return new Response(null, { status: 204, headers: CORS_HEADERS });
  }

  let action = "";
  try {
    // ── Ensure all SQLite/Turso tables exist (no-op after first cold start) ─
    await ensureTables();

    const url = new URL(req.url);
    const rawPath = (url.pathname || "").replace(/\/+$/, "");
    const path = rawPath.toLowerCase();
    action = url.searchParams.get("action") || "";
    if (!action && req.method === "POST") {
      const adminPathToAction: Record<string, string> = {
        "/admin/sync/groups": "admin_sync_groups",
        "/admin/sync/properties": "admin_sync_properties",
        "/admin/sync/vendors": "admin_sync_vendors",
        "/admin/sync/work-orders": "admin_sync_work_orders",
        "/admin/sync/billing": "admin_sync_billing",
        "/admin/rebuild-cache": "admin_rebuild_cache",
        "/admin/invalidate-cache": "admin_invalidate_cache",
        "/admin/resolve-context": "admin_resolve_context",
        "/admin/resolve-groups": "admin_resolve_groups",
      };
      action = adminPathToAction[path] || "";
    }
    const params: Record<string, string> = {};
    url.searchParams.forEach((v, k) => {
      params[k] = v;
    });

    // Source IP — read from trustworthy proxy headers (Val Town / Cloudflare)
    const requestIp =
      (req.headers.get("x-forwarded-for") || "").split(",")[0].trim() ||
      req.headers.get("cf-connecting-ip") ||
      req.headers.get("x-real-ip") ||
      "unknown";

    // ── Public short-link redirect (must stay before env/auth guards) ───────
    // This endpoint must remain public because SMS recipients are unauthenticated.
    if (req.method === "GET" && rawPath.startsWith("/s/")) {
      const shortLinkMatch = rawPath.match(/^\/s\/([A-Za-z0-9_-]{4,64})$/);
      if (!shortLinkMatch) {
        return jsonResp({ ok: false, error: "No code provided" }, 400);
      }

      const resolved = await resolveShortLink(shortLinkMatch[1]);
      if (!resolved.ok || !resolved.fullUrl) {
        const status = resolved.reason === "expired" ? 410 : 404;
        const err = resolved.reason === "expired"
          ? "Link expired"
          : "Link not found";
        return jsonResp(
          { ok: false, error: err, code: shortLinkMatch[1] },
          status,
        );
      }

      return Response.redirect(resolved.fullUrl, 302);
    }

    // ── Validate required AppFolio env vars before any API call ─────────────
    // Ensures the formatting of the Client ID and Client Secret passed into the
    // authorization header are valid.
    if (!V2_CLIENT_ID || !V2_CLIENT_SECRET || !DEV) {
      return jsonResp(
        {
          ok: false,
          error:
            "Missing required environment variables: AF_V2_CLIENT_ID, AF_V2_CLIENT_SECRET, AF_DEVELOPER_ID",
          hint:
            "Set these in Val Town Project Settings -> Environment Variables",
        },
        500,
      );
    }
    if (!MAGIC_LINK_SECRET) {
      return jsonResp(
        {
          ok: false,
          error:
            "Missing required environment variable: MAGIC_LINK_SECRET (used to sign magic link tokens)",
          hint:
            "Set MAGIC_LINK_SECRET in Val Town Project Settings -> Environment Variables",
        },
        500,
      );
    }

    const isWebhookIngress = req.method === "POST" && !action;
    const isPortalAction = action === "portal";
    const PORTAL_API_ACTIONS = new Set([
      "portal_validate",
      "portal_schedule",
      "portal_reschedule",
      "portal_note",
      "portal_no_contact",
      "portal_reassign_request",
      "send_tenant_sms",
    ]);
    const isPortalApiAction = PORTAL_API_ACTIONS.has(action);
    const isShortLinkIngress = req.method === "GET" &&
      rawPath.startsWith("/s/");
    const isCronAction = CRON_ACTIONS.has(action);
    const isDeviceSetup = action === "device_setup";
    const isDeviceOtpRequest = action === "device_otp_request";
    const isDeviceOtpVerify = action === "device_otp_verify";
    const isVerifyRole = action === "verify_role";
    const isPublicHealth = action === "ping";
    const isPublicInfo =
      req.method === "GET" && (!action || action === "version");
    const headerAdminKey = (req.headers.get("x-admin-key") || "").trim();
    const headerAdminToken = (req.headers.get("x-admin-token") || "").trim();
    const adminSecret = (Deno.env.get("ADMIN_SECRET") || "").trim();
    const isHeaderAdminAuthorized =
      (!!PROXY_ADMIN_KEY && headerAdminKey === PROXY_ADMIN_KEY) ||
      (!!adminSecret && headerAdminToken === adminSecret);
    const frontendToken = extractFrontendToken(req.headers);
    let frontendSession: any = null;
    if (frontendToken) {
      try {
        frontendSession = await getTrustedDeviceSession(frontendToken);
      } catch (e: any) {
        // Session enrichment should never take down unrelated proxy actions.
        console.warn(
          "frontend session lookup failed; continuing without session:",
          String(e?.message || e).substring(0, 220),
        );
        frontendSession = null;
      }
    }

    // Frontend authentication wall (prevents open relay abuse).
    // Exemptions:
    //   - AppFolio webhook ingress POST (no action)
    //   - Magic portal HTML route
    //   - Magic portal API actions (portal_validate, portal_schedule, etc) — use magic token, not device token
    //   - Cron actions (separately guarded by CRON secret)
    //   - Public info routes (root help + version)
    //   - Requests carrying valid admin key in x-admin-key header
    if (
      ENFORCE_FRONTEND_AUTH && !isWebhookIngress && !isPortalAction &&
      !isPortalApiAction && !isShortLinkIngress && !isCronAction &&
      !isDeviceSetup && !isDeviceOtpRequest && !isDeviceOtpVerify &&
      !isVerifyRole && !isPublicHealth && !isPublicInfo &&
      !isHeaderAdminAuthorized
    ) {
      if (!(await isFrontendAuthorized(req.headers))) {
        return jsonResp({
          ok: false,
          error: "Unauthorized — missing or invalid frontend bearer token",
        }, 401);
      }
    }

    const isWriteMethod = req.method !== "GET" && req.method !== "HEAD" &&
      req.method !== "OPTIONS";
    if (
      frontendSession && frontendSession.role === "pm_readonly" && isWriteMethod
    ) {
      // PM OTP mode is intentionally view-only at API layer.
      if (
        action !== "device_otp_request" && action !== "device_otp_verify" &&
        action !== "device_setup"
      ) {
        return jsonResp({
          ok: false,
          error:
            "Read-only session: write operations are disabled for PM OTP access",
        }, 403);
      }
    }

    // ── PM scope enforcement — inject property group constraint into params ───
    // When a session has a non-full role and a bound property_group_uuid, all
    // scoped data-read actions are constrained to that group.  Admin sessions
    // (role === 'full') and unauthenticated requests pass through unfiltered.
    // The group_uuid is injected into params here so every handler that already
    // accepts group_uuid/property_group_id automatically enforces it.
    const sessionScopeUuid = (
      frontendSession &&
      frontendSession.property_group_uuid &&
      frontendSession.role !== "full"
    ) ? String(frontendSession.property_group_uuid) : null;

    const SCOPED_DATA_ACTIONS = new Set([
      "work_orders",
      "work_orders_completed_history",
      "completed_work_orders_history",
      "bills",
      "bills_stats",
      "bills_history",
      "bill_detail",
      "bills_sync",
      "bills_list",
      "bills_by_vendor",
      "bills_by_property",
      "bills_by_wo",
      "bills_by_wo_number",
      "bills_by_invoice",
      "bills_due_range",
      "turns",
      "unit_turns",
      "unit_turns_history",
      "closed_turns",
      "upcoming_moveouts",
      "turn_work_orders",
      "labor",
      "recent_tasks",
      "wo_comparison_report",
      "inspections",
      "property_groups",
    ]);

    if (sessionScopeUuid && SCOPED_DATA_ACTIONS.has(action)) {
      // Always override — scoped PMs cannot bypass their group boundary.
      params.group_uuid = sessionScopeUuid;
      params.property_group_id = sessionScopeUuid;
    }

    // Optional path-style endpoint for frontend polling.
    if (
      (req.method === "GET" || req.method === "HEAD") &&
      path.endsWith("/webapp/feed")
    ) {
      return jsonResp(await handleWebhookFeed(params));
    }

    // ── HTML-returning routes bypass corsJson wrapper ─────────────────────────
    // These must be checked BEFORE the main switch to return raw Responses.

    // Magic link tech portal — returns text/html
    if (action === "portal") {
      return await handleMagicPortal(params);
    }

    // ── Cron endpoint auth guard ──────────────────────────────────────────────
    if (CRON_ACTIONS.has(action)) {
      if (!isCronAuthorized(params, req.headers)) {
        return jsonResp({
          ok: false,
          error: "Unauthorized — invalid or missing cron secret",
        }, 401);
      }
    }

    // ── AppFolio inbound webhook (POST with no ?action param) ─────────────────
    // AppFolio delivers all webhook events as HTTPS POST to the registered URL.
    // All outgoing webhooks are signed — raw body is stored verbatim .
    if (req.method === "POST" && !action) {
      // Rate-limit by source IP: 300 deliveries per minute is well above any
      // legitimate AppFolio burst but blocks runaway replay abuse.
      const whRl = checkRateLimit(`wh:${requestIp}`, 300, 60_000);
      if (!whRl.allowed) {
        return jsonResp(
          { ok: false, error: "Webhook rate limit exceeded — retry later" },
          429,
        );
      }
      const webhookResult = await handleWebhookPost(req);
      return jsonResp(webhookResult, webhookResult?.status || 200);
    }

    // ── Non-GET passthrough (PATCH/PUT/DELETE/POST) ────────────────────────
    // Frontend write calls use action=passthrough with the original HTTP verb.
    // Route any mutating verb directly so request body and method are preserved.
    if (
      action === "passthrough" && req.method !== "GET" && req.method !== "HEAD"
    ) {
      const result = await handlePassthrough(params, req);
      if (result instanceof Response) return result;
      return jsonResp(result);
    }

    // ── POST actions ──────────────────────────────────────────────────────────
    if (req.method === "POST" && action) {
      // ── Per-action rate limits (checked before switch dispatch) ────────────
      // device_setup: prevent setup-PIN brute force from unauthenticated callers.
      if (action === "device_setup") {
        const setupRl = checkRateLimit(`setup:${requestIp}`, 5, 10 * 60_000);
        if (!setupRl.allowed) {
          return jsonResp(
            {
              ok: false,
              error:
                "Too many setup attempts — please wait before retrying",
              retry_after_ms: setupRl.retryAfterMs,
            },
            429,
          );
        }
      }
      // portal_* actions: prevent replayed token or cross-token enumeration
      // from the same source IP.  60 per minute is well above normal portal use.
      if (action.startsWith("portal_")) {
        const portalRl = checkRateLimit(`portal:${requestIp}`, 60, 60_000);
        if (!portalRl.allowed) {
          return jsonResp(
            {
              ok: false,
              error: "Too many portal requests — try again later",
              retry_after_ms: portalRl.retryAfterMs,
            },
            429,
          );
        }
      }

      let result: any;

      switch (action) {
        // Turn tracker writes
        case "turn_records":
          result = await handleTurnRecords(params, req);
          break;
        case "turn_record_stage":
          result = await handleTurnRecordStage(req);
          break;
        case "wo_note_create":
          result = await handleWoNoteCreate(req);
          break;
        case "bill_attachment_upload":
          result = await handleBills(params, req, action);
          break;
        case "closed_turns":
          result = await handleClosedTurns(params, req);
          break;
        case "unit_turns_sync":
          result = await handleUnitTurnsSync(req);
          break;
        case "unit_turn_wo_link":
          result = await handleUnitTurnWorkOrderLink(req);
          break;
        case "unit_turn_wo_unlink":
          result = await handleUnitTurnWorkOrderUnlink(req);
          break;

        // Generic report POST (filters in JSON body)
        case "report":
          result = await handleGenericReport(params, req);
          break;

        // Raw API proxy (POST passthrough)
        case "passthrough":
          result = await handlePassthrough(params, req);
          if (result instanceof Response) return result;
          break;

        // Legacy webhook POST alias
        case "webhook":
          result = await handleWebhookPost(req);
          if (result?.status && result.status !== 200) {
            return jsonResp(result, result.status);
          }
          break;
        case "webhook_events":
          // Allow POST webhook delivery via action=webhook_events as a hardened alias.
          result = await handleWebhookPost(req);
          if (result?.status && result.status !== 200) {
            return jsonResp(result, result.status);
          }
          break;

        case "device_setup":
          result = await handleDeviceSetup(req);
          if (result?.status && result.status !== 200) {
            return jsonResp(result, result.status);
          }
          break;

        case "device_otp_request":
          result = await handleDeviceOtpRequest(req);
          if (result?.status && result.status !== 200) {
            return jsonResp(result, result.status);
          }
          break;

        case "device_otp_verify":
          result = await handleDeviceOtpVerify(req);
          if (result?.status && result.status !== 200) {
            return jsonResp(result, result.status);
          }
          break;

        case "verify_role":
          result = await handleVerifyRole(req);
          if (result?.status && result.status !== 200) {
            return jsonResp(result, result.status);
          }
          break;

        case "trusted_device_revoke":
          result = await handleTrustedDeviceRevoke(params, req);
          break;
        case "pm_proxy_user_upsert":
          result = await handlePmProxyUserUpsert(req);
          break;
        case "pm_proxy_user_delete":
          result = await handlePmProxyUserDelete(req);
          break;

        // Webhook resolve
        case "webhook_resolve":
          result = await handleWebhookResolve(params);
          break;

        // SQL admin (read + write) — key and query must be in JSON body, never URL params
        case "sql_query":
          result = await handleSqlQuery(params, req);
          break;
        case "sql_execute":
          result = await handleSqlExecute(params, req);
          break;

        // Proxy config settings (OTP policy etc.) — admin key required in body
        case "settings_set": {
          let sbody: any = {};
          try {
            sbody = await req.json();
          } catch {
            sbody = {};
          }
          const skey = String(sbody.key || "").trim();
          const sval = String(sbody.value ?? "").trim();
          const adminKey =
            String(sbody.admin_key || sbody.key_auth || "").trim() ||
            (req.headers.get("x-admin-key") || "");
          if (!PROXY_ADMIN_KEY || adminKey !== PROXY_ADMIN_KEY) {
            result = { ok: false, error: "Unauthorized: invalid admin key" };
            break;
          }
          const ALLOWED_SETTINGS = new Set([
            "otp_enabled",
            "otp_allowed_domain",
            "otp_require_pm_membership",
            "otp_ttl_minutes",
            "brand_name",
            "brand_logo_url",
          ]);
          if (!skey || !ALLOWED_SETTINGS.has(skey)) {
            result = {
              ok: false,
              error: `Unknown or disallowed setting key: '${skey}'`,
            };
            break;
          }

          await upsertSettingTable("proxy_config", skey, sval);
          try {
            await upsertSettingTable("app_settings", skey, sval);
          } catch (mirrorErr: any) {
            console.warn(
              `[settings_set] app_settings mirror failed for ${skey}: ${String(mirrorErr?.message || mirrorErr)}`,
            );
          }

          result = { ok: true, key: skey, value: sval, source: "proxy_config" };
          break;
        }

        // v9.2.2 — tech roster upsert
        case "tech_roster":
          result = await handleTechRoster(params, req);
          break;

        // vendor category / compliant override persist
        case "vendor_override":
          result = await handleVendorOverride(params, req);
          break;

        case "routing_monitor": {
          const _rmMod = await import("./handlers/routingMonitor.ts");
          const _rmFn = _rmMod.handleRoutingMonitor ?? null;
          if (typeof _rmFn !== "function") {
            console.error(
              "[routing_monitor POST] export check failed. Found:",
              Object.keys(_rmMod),
            );
            return jsonResp({
              ok: false,
              error: "handleRoutingMonitor export not found",
              exports: Object.keys(_rmMod),
            }, 500);
          }
          result = await _rmFn(params, req);
          break;
        }

        // v9.2.2 — tenant SMS via magic link portal
        case "send_tenant_sms":
          result = await handleSendTenantSMS(req);
          break;
        case "portal_validate":
          result = await handlePortalValidate(req);
          break;
        case "portal_schedule":
          result = await handlePortalSchedule(req);
          break;
        case "portal_reschedule":
          result = await handlePortalReschedule(req);
          break;
        case "portal_note":
          result = await handlePortalNote(req);
          break;
        case "portal_no_contact":
          result = await handlePortalNoContact(req);
          break;
        case "portal_reassign_request":
          result = await handlePortalReassignRequest(req);
          break;
        case "generate_magic_link":
          result = await handleGenerateMagicLink(req);
          break;

        // v9.2.2 — monitored work order queue management
        case "add_monitored_work_order":
          result = await handleAddMonitoredWO(req);
          break;
        case "remove_monitored_work_order":
          result = await handleRemoveMonitoredWO(req);
          break;

        // v9.2.2 — dispatch test message / magic link verification
        case "send_magic_link_test_sms":
          result = await handleSendMagicLinkTestSMS(req);
          break;

        // v9.2.2 — cron triggers (also accept POST from cron vals)
        case "noon_warning_cron":
          result = await handleNoonWarningCron(params);
          break;
        case "midnight_reassign_cron":
          result = await handleMidnightReassignCron(params);
          break;

        // v9.2.2 — dispatch roster sync aliases
        case "dispatch_sync_assignees":
        case "sync_assignee_roster":
        case "reassignment_sync_assignees":
          result = await handleDispatchSyncAssignees(params, req);
          break;

        // v9.2.2 — dispatch test queue seed aliases
        case "dispatch_seed_reassignment_test":
        case "seed_reassignment_queue_test":
        case "reassignment_queue_seed":
          result = await handleDispatchSeedReassignmentTest(params, req);
          break;

        case "admin_sync_groups":
        case "admin_sync_properties":
        case "admin_sync_vendors":
        case "admin_sync_work_orders":
        case "admin_sync_billing":
        case "admin_rebuild_cache":
        case "admin_invalidate_cache":
        case "admin_resolve_context":
        case "admin_resolve_groups":
          result = await handleAdminSyncRoute(action, params, req);
          break;

        default:
          result = { ok: false, error: `Unknown POST action: "${action}"` };
      }

      if (result instanceof Response) return result;

      return jsonResp(result);
    }

    // ── GET / HEAD actions ────────────────────────────────────────────────────
    if (req.method === "GET" || req.method === "HEAD") {
      if (action === "wo_detail" || action === "wo_notes") {
        const woRef = String(params.uuid || params.wo_id || params.id || "").trim() || "(missing)";
        const woReadRl = checkRateLimit(
          `wo-read:${action}:${requestIp}:${woRef}`,
          20,
          60_000,
        );
        if (!woReadRl.allowed) {
          return jsonResp(
            {
              ok: false,
              error: "Too many work-order detail/note requests — try again later",
              retry_after_ms: woReadRl.retryAfterMs,
            },
            429,
          );
        }
      }

      let result: any;

      switch (action) {
        // ── Health ────────────────────────────────────────────────────────────
        case "ping":
          result = await handlePing();
          break;
        case "version":
          result = { ok: true, version: PROXY_VERSION, proxy: PROXY_VERSION };
          break;
        case "session_info":
          // Slide the rolling 30-day expiry on every authenticated heartbeat.
          if (frontendSession && frontendToken) {
            void touchDeviceSession(frontendToken);
          }
          result = {
            ok: true,
            authenticated: !!frontendSession,
            session: frontendSession || null,
          };
          break;

        case "webhook_live":
          result = await handleWebhookLive(params);
          break;
        case "webapp_feed":
          result = await handleWebhookFeed(params);
          break;
        // ── Work order data ───────────────────────────────────────────────────
        // Delegates to handleWorkOrders (Reports API v2 + api_cache).
        // Notes fetched on-demand via ?action=wo_notes&wo_id=<uuid>.
        case "work_orders":
          result = await handleWorkOrders(params);
          break;
        case "wo_notes":
          result = await handleWoNotes(params);
          break;
        case "wo_detail":
          result = await handleWoDetail(params);
          break;
        case "wo_billed_amount":
          result = await handleWoBilledAmount(params);
          break;
        case "turn_work_orders":
          result = await handleTurnWorkOrders(params);
          break;
        case "work_orders_completed_history":
          result = await handleCompletedWorkOrdersHistory(params);
          break;
        case "completed_work_orders_history":
          result = await handleCompletedWorkOrdersHistory(params);
          break;
        case "labor":
          result = await handleLabor(params);
          break;
        case "recent_tasks":
          result = await handleRecentTasks(params);
          break;
        case "routing_monitor": {
          const _rmMod = await import("./handlers/routingMonitor.ts");
          const _rmFn = _rmMod.handleRoutingMonitor ?? null;
          if (typeof _rmFn !== "function") {
            console.error(
              "[routing_monitor GET] export check failed. Found:",
              Object.keys(_rmMod),
            );
            return jsonResp({
              ok: false,
              error: "handleRoutingMonitor export not found",
              exports: Object.keys(_rmMod),
            }, 500);
          }
          result = await _rmFn(params, req);
          break;
        }

        // ── Property data ─────────────────────────────────────────────────────
        case "properties":
          result = await handleProperties(params);
          break;
        case "property_groups":
          result = await handlePropertyGroups(params);
          break;
        case "property_map":
          result = await handlePropertyMap(params);
          break;
        case "upcoming_moveouts":
          result = await handleUpcomingMoveouts(params);
          break;

        // ── Turn pipeline ─────────────────────────────────────────────────────
        case "turns":
          result = await handleTurns(params);
          break;
        case "unit_turns":
          result = await handleUnitTurns(params);
          break;
        case "turn_records":
          result = await handleTurnRecords(params);
          break;
        case "turns_incremental":
          result = await handleTurnsIncremental(params);
          break;
        case "closed_turns":
          result = await handleClosedTurns(params);
          break;
        case "unit_turns_history":
          result = await handleUnitTurnsHistory(params);
          break;

        // ── Inspections ───────────────────────────────────────────────────────
        case "inspections":
          result = await handleInspections(params);
          break;

        // ── Vendors ───────────────────────────────────────────────────────────
        case "vendors":
          result = await handleVendors(params);
          break;
        case "vendor_override":
          result = await handleVendorOverride(params, req);
          break;

        // ── Bills ─────────────────────────────────────────────────────────────
        case "bills":
        case "bills_stats":
        case "bills_sync":
        case "bill_detail":
        case "bills_history":
        case "bill_attachments":
          result = await handleBills(params, req, action);
          break;
        case "bills_by_vendor":
        case "bills_by_property":
        case "bills_by_wo":
        case "bills_by_wo_number":
        case "bills_by_invoice":
        case "bills_due_range":
        case "bills_list":
          {
            const billsParams = new URLSearchParams(url.searchParams);
            if (params.sessionScopeUuid && !billsParams.get("group_id")) {
              billsParams.set("group_id", params.sessionScopeUuid);
            }
            result = await handleBillsRoute(req, sqlite, billsParams);
          }
          if (result instanceof Response) return result;
          break;
        // ── WO Comparison Report ──────────────────────────────────────────────
        // Returns JSON normally; returns raw CSV Response when ?format=csv.
        // Staggering requests prevents congestion on this multi-fetch report .
        case "wo_comparison_report":
          result = await handleWoComparisonReport(params, req);
          if (result instanceof Response) return result; // CSV download
          break;

        // ── Webhook endpoints ─────────────────────────────────────────────────
        // AppFolio signs all outgoing webhooks — raw bodies stored verbatim .
        case "webhook_events":
          result = await handleWebhookEvents(params);
          break;
        case "webhook_stats":
          result = await handleWebhookStats(params);
          break;
        case "webhook_resolve":
          result = await handleWebhookResolve(params);
          break;
        case "webhook_migrate":
          result = await handleWebhookMigrate();
          break;
        case "webhook_cleanup":
          result = await handleWebhookCleanupEndpoint();
          break;

        // ── Generic report (GET form) ─────────────────────────────────────────
        case "report":
          result = await handleGenericReport(params, req);
          break;

        // ── Cache management ──────────────────────────────────────────────────
        case "cache_stats":
          result = await handleCacheStats();
          break;
        case "cache_invalidate":
          result = await handleCacheInvalidate(params);
          break;
        case "force_refresh":
          result = await handleForceRefresh(params);
          break;
        case "storage_cleanup":
          result = await handleStorageCleanup();
          break;

        // ── Diagnostics ───────────────────────────────────────────────────────
        case "debug_sqlite":
          result = await handleDebugSqlite();
          break;

        // ── Migration (one-time) ──────────────────────────────────────────────
        case "migrate_v8":
          result = await handleMigrateV8();
          break;

        // ── Passthrough / raw proxy ───────────────────────────────────────────
        // Routes /api/v0/... to AF_DB with dbHeaders()
        // Routes all other paths to AF_REPORTS with reportsHeaders()
        // If the request URL is not correct, check the path is valid .
        case "passthrough":
          result = await handlePassthrough(params, req);
          if (result instanceof Response) return result;
          break;

        // ── v9.2.2 Cron triggers ──────────────────────────────────────────────
        // These are also callable via GET for manual testing.
        // Auth checked above via isCronAuthorized().
        case "noon_warning_cron":
          result = await handleNoonWarningCron(params);
          break;
        case "midnight_reassign_cron":
          result = await handleMidnightReassignCron(params);
          break;

        // ── v9.2.2 Dispatch Control data ──────────────────────────────────────
        // AssignedUsers must reference a Maintenance Tech role user .
        case "reassignment_queue":
          result = await handleReassignmentQueue(params);
          break;
        case "trusted_devices":
          result = await handleTrustedDeviceList(params, req);
          break;
        case "pm_proxy_users":
          const { handlePmProxyUsers } = await import("./handlers/pmProxyUsers.ts");
          result = await handlePmProxyUsers();
          break;
        case "settings_get": {
          const adminKey = params.key || req.headers.get("x-admin-key") || "";
          if (!PROXY_ADMIN_KEY || adminKey !== PROXY_ADMIN_KEY) {
            result = { ok: false, error: "Unauthorized: invalid admin key" };
            break;
          }
          const limit = Math.max(
            1,
            Math.min(500, parseInt(String(params.limit || "200"), 10) || 200),
          );
          const offset = Math.max(
            0,
            parseInt(String(params.offset || "0"), 10) || 0,
          );
          try {
            let proxyRows: ProxySettingRow[] = [];
            let appRows: ProxySettingRow[] = [];
            const warnings: string[] = [];

            try {
              proxyRows = await readSettingsTable("proxy_config");
            } catch (proxyErr: any) {
              warnings.push(
                `proxy_config read failed: ${String(proxyErr?.message || proxyErr)}`,
              );
            }

            try {
              appRows = await readSettingsTable("app_settings");
            } catch (appErr: any) {
              warnings.push(
                `app_settings read failed: ${String(appErr?.message || appErr)}`,
              );
            }

            const mergedMap = new Map<string, ProxySettingRow>();
            for (const row of appRows) mergedMap.set(row.key, row);
            for (const row of proxyRows) mergedMap.set(row.key, row);

            const safeRows = Array.from(mergedMap.values()).sort((a, b) =>
              a.key.localeCompare(b.key)
            );
            const source = proxyRows.length && appRows.length
              ? "merged"
              : proxyRows.length
              ? "proxy_config"
              : appRows.length
              ? "app_settings"
              : "empty";
          const total = safeRows.length;
          const pagedRows = safeRows.slice(offset, offset + limit);
          const settingsMap: Record<string, string> = {};
          for (const r of pagedRows) {
            if (r?.key != null) {
              settingsMap[String(r.key)] = String(r.value ?? "");
            }
          }

          result = {
            ok: true,
            source,
            settings: pagedRows,
            settings_map: settingsMap,
            count: pagedRows.length,
            total,
            limit,
            offset,
            warnings,
          };
          } catch (dbErr: any) {
            result = {
              ok: false,
              error: "db_error",
              message: `Database error in settings_get: ${String(dbErr?.message || dbErr)}`,
            };
          }
          break;
        }
        case "tech_roster":
          result = await handleTechRoster(params);
          break;
        case "tenant_comms_log":
          result = await handleTenantCommsLog(params);
          break;
        case "users":
          result = await handleUsers(params);
          break;

        // v9.2.2 — dispatch sync/seed aliases (GET for manual testing)
        case "dispatch_sync_assignees":
        case "sync_assignee_roster":
        case "reassignment_sync_assignees":
          result = await handleDispatchSyncAssignees(params);
          break;
        case "dispatch_seed_reassignment_test":
        case "seed_reassignment_queue_test":
        case "reassignment_queue_seed":
          result = await handleDispatchSeedReassignmentTest(params);
          break;

        // ── Default: ?path= raw proxy compatibility mode or help text ─────────
        default:
          // Compatibility raw proxy pattern: ?path=/api/v0/...
          if (params.path) {
            result = await handlePassthrough(params, req);
            if (result instanceof Response) return result;
            break;
          }

          // No action, no path — return full endpoint directory
          result = {
            ok: true,
            service: "HandyManager Proxy",
            version: PROXY_VERSION,
            timestamp: new Date().toISOString(),
            database: (await import("./config.ts")).TURSO_URL
              ? "turso"
              : "valtown_sqlite",
            domains: {
              db_api: "https://api.appfolio.com",
              reports_api: "https://flraz.appfolio.com",
            },
            endpoints: {
              health: [
                "GET  ?action=ping",
                "GET  ?action=cache_stats",
                "GET  ?action=debug_sqlite",
                "GET  ?action=storage_cleanup",
              ],
              work_orders: [
                "GET  ?action=work_orders&days=180",
                "GET  ?action=wo_notes&wo_id=UUID",
                "GET  ?action=wo_detail&wo_id=UUID",
                "GET  ?action=wo_billed_amount&wo_number=12345",
                "GET  ?action=turn_work_orders&days=90",
                "GET  ?action=work_orders_completed_history&days=365",
                "GET  ?action=completed_work_orders_history&days=365  (alias)",
                "GET  ?action=labor&days=1&statuses=8",
                "GET  ?action=recent_tasks",
                "GET  ?action=wo_comparison_report&from_date=2026-01-01&to_date=2026-03-21",
                "GET  ?action=wo_comparison_report&format=csv",
                "GET  ?action=wo_comparison_report&group=phoenix&format=csv",
                "GET  ?action=wo_comparison_report&group=tucson&format=csv",
                "GET  ?action=wo_comparison_report&force=1",
              ],
              vendors_properties: [
                "GET  ?action=vendors",
                "GET  ?action=routing_monitor&op=capabilities",
                "GET  ?action=routing_monitor&op=pm_map",
                "GET  ?action=routing_monitor&op=events&status=pending&days=30",
                "GET  ?action=routing_monitor&op=pm_stats&days=30",
                "POST ?action=routing_monitor&op=scan                body: {events:[...]}",
                "POST ?action=routing_monitor&op=review              body: {id, review_status, review_notes?}",
                "POST ?action=routing_monitor&op=capability_upsert   body: {id?, trade, keywords[], active}",
                "POST ?action=routing_monitor&op=capability_delete   body: {id}",
                "POST ?action=routing_monitor&op=pm_map_upsert       body: {group_name, pm_name}",
                "POST ?action=routing_monitor&op=pm_map_bulk         body: {entries:[{group_name,pm_name}]}",
                "GET  ?action=properties",
                "GET  ?action=property_groups",
                "GET  ?action=property_map",
                "GET  ?action=upcoming_moveouts&days=60",
              ],
              turns_inspections: [
                "GET  ?action=turns&days=90",
                "GET  ?action=unit_turns&days=180&limit=50&offset=0",
                "GET  ?action=unit_turns_history&days=540&limit=300",
                "GET  ?action=turns_incremental&since=2026-04-01T00:00:00Z&limit=800",
                "GET  ?action=turn_records",
                "GET  ?action=closed_turns",
                "POST ?action=turn_records      body: {id, ...fields}",
                "POST ?action=turn_record_stage body: {id, stage, data}",
                "POST ?action=unit_turns_sync      body: {records:[...]}",
                "POST ?action=unit_turn_wo_link    body: {turn_key, wo_id, ...}",
                "POST ?action=unit_turn_wo_unlink  body: {turn_key, wo_id}",
                "GET  ?action=inspections&days=180",
                "GET  ?action=bills&days=180",
                "GET  ?action=bills_stats",
                "GET  ?action=bills_history&days=365",
                "GET  ?action=bill_detail&bill_id=UUID",
                "GET  ?action=bill_attachments&bill_id=UUID",
                "POST ?action=bill_attachment_upload&bill_id=UUID",
              ],
              webhooks: [
                "POST (no action)               ← AppFolio webhook delivery",
                "POST ?action=webhook           ← legacy alias",
                "GET  ?action=webhook_events&limit=200&type=work_order",
                "GET  ?action=webhook_live&since_id=123&limit=20",
                "GET  ?action=webapp_feed&limit=50&since=2026-03-30T00:00:00Z",
                "GET  ?action=webhook_stats",
                "GET  ?action=webhook_resolve&resource_type=work_order&resource_id=UUID",
                "GET  ?action=webhook_migrate",
                "GET  ?action=webhook_cleanup",
              ],
              trusted_devices: [
                "POST ?action=device_setup       body: { pin, user_name }",
                "GET  ?action=trusted_devices&key=PROXY_ADMIN_KEY&limit=100&offset=0",
                "POST ?action=trusted_device_revoke body: { token, key }",
                "GET  ?action=pm_proxy_users&key=PROXY_ADMIN_KEY&limit=100&offset=0",
                "POST ?action=pm_proxy_user_upsert body: { key, email, property_group_uuid, ... }",
                "POST ?action=pm_proxy_user_delete body: { key, user_uuid|email }",
                "POST ?action=verify_role body: { password }",
              ],
              cache: [
                "GET  ?action=cache_invalidate&type=work_orders",
                "GET  ?action=force_refresh&type=ENTITY",
              ],
              reports: [
                "POST ?action=report&name=REPORT_NAME  body: {filters}",
                "GET  ?action=report&name=REPORT_NAME",
              ],
              passthrough: [
                "GET  ?action=passthrough&path=/api/v0/work_orders?...",
                "POST ?action=passthrough&path=/api/v2/reports/NAME.json",
                "GET  ?path=/api/v0/...  (compat mode — no action param)",
              ],
              admin_sql: [
                "POST ?action=sql_query   body: { key, query }  (SELECT/PRAGMA/EXPLAIN/WITH only)",
                "POST ?action=sql_execute body: { key, sql, args? }  (writes — key never in URL)",
                "GET  ?action=settings_get&key=PROXY_ADMIN_KEY&limit=200&offset=0",
                "POST ?action=settings_set body: { admin_key, key, value }",
              ],
              migration: [
                "GET  ?action=migrate_v8",
              ],
              v9_dispatch: [
                "GET  ?action=reassignment_queue&limit=50",
                "GET  ?action=reassignment_queue&wo_id=UUID",
                "GET  ?action=tech_roster",
                "POST ?action=tech_roster       body: {tech_id, tech_name, tech_phone, tier?, geo_zone?, active?}",
                "GET  ?action=tenant_comms_log&limit=100",
                "GET  ?action=tenant_comms_log&wo_id=UUID",
                "GET  ?action=users&role=maintenance_tech&since=2020-01-01T00:00:00Z",
                "GET  ?action=portal&token=TOKEN              ← returns HTML",
                "POST ?action=portal_validate               body: {token}",
                "POST ?action=portal_schedule               body: {token, scheduled_date, scheduled_window}",
                "POST ?action=portal_reschedule             body: {token, scheduled_date, scheduled_window}",
                "POST ?action=portal_note                   body: {token, note_text}",
                "POST ?action=portal_no_contact             body: {token, attempts, reason?, details?}",
                "POST ?action=portal_reassign_request       body: {token, reason, details?}",
                "POST ?action=send_tenant_sms                 body: {token, template: enroute|schedule|today}",
                "POST ?action=send_magic_link_test_sms        body: {phone, tech_name?, tech_id?}",
                "POST ?action=generate_magic_link             body: {wo_id, tech_id, tech_phone, ...}",
                "POST ?action=add_monitored_work_order        body: {wo_id}",
                "POST ?action=remove_monitored_work_order     body: {wo_id}",
                "POST ?action=dispatch_sync_assignees         body: {tier1_group_uuid, tier2_group_uuid, branch?}",
                "POST ?action=sync_assignee_roster            (alias)",
                "POST ?action=reassignment_sync_assignees     (alias)",
                "POST ?action=dispatch_seed_reassignment_test body: {limit?, branch?, tier1_group_uuid?, tier2_group_uuid?}",
                "POST ?action=seed_reassignment_queue_test    (alias)",
                "POST ?action=reassignment_queue_seed         (alias)",
                "GET  ?action=noon_warning_cron&secret=KEY",
                "GET  ?action=midnight_reassign_cron&secret=KEY",
                "POST ?action=noon_warning_cron               ← cron val POST",
                "POST ?action=midnight_reassign_cron          ← cron val POST",
              ],
            },
          };

          if (action) {
            result = { ok: false, error: `Unknown action: "${action}"` };
          }
      }

      return jsonResp(result);
    }

    // ── Unsupported method ────────────────────────────────────────────────────
    return jsonResp(
      { ok: false, error: `Method "${req.method}" not allowed` },
      405,
    );
  } catch (err: any) {
    // Top-level error boundary — logs to Val Town console + returns safe JSON.
    // 503 and 533 from AppFolio are retried automatically in fetchWithTimeout .
    // If a semantic error (422) bubbles here, check parameter validity .
    console.error(
      `[HandyManager v9.2.2] Unhandled error on action "${action}":`,
      err,
    );
    return jsonResp(
      {
        ok: false,
        error: "internal_error",
        message: String(err?.message || err),
        action: action || "(none)",
      },
      500,
    );
  }
}