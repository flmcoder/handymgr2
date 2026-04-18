// ============================================================================
// handlers/passthrough.ts — Direct API proxy, cache management, diagnostics,
//                           SQL admin, and health check endpoints.
//
// Exports:
//   handlePassthrough      — raw proxy to AF DB API v0 or Reports API v2
//   handlePing             — health check against both AF API domains
//   handleCacheStats       — storage usage + per-entity cache summary
//   handleCacheInvalidate  — manual cache eviction by entity type
//   handleForceRefresh     — invalidate + immediately re-fetch one entity
//   handleStorageCleanup   — purge expired cache + trim webhook overflow
//   handleDebugSqlite      — raw PRAGMA dump for troubleshooting
//   handleSqlQuery         — admin SELECT endpoint (guarded by PROXY_ADMIN_KEY)
//   handleSqlExecute       — admin write endpoint  (guarded by PROXY_ADMIN_KEY)
//
// Passthrough routing logic:
//   Paths starting with /api/v0  → AF_DB      (api.appfolio.com) + dbHeaders()
//   All other paths              → AF_REPORTS (flraz.appfolio.com) + reportsHeaders()
//
// Ensure the formatting of the Client ID and Client Secret passed into your
// authorization header are valid . If an incorrect parameter value
// was passed, the returned error codes should be descriptive enough to assist
// in identifying the problem .
//
// The correct user roles must be enabled in AppFolio Property Manager for
// certain API requests — for example, assigning a work order to a user
// requires that user to have the Maintenance Tech role, or the API will
// return a 422 .
//
// Staggering the rate at which requests are issued will prevent network
// congestion and errors . All paginated passthrough calls inherit
// the 150 ms PAGE_DELAY_MS pacing from fetchWithTimeout's retry loop.
// ============================================================================

import {
  cacheInvalidate,
  getCacheSizeBytes,
  getWebhookSizeBytes,
  rowsAsObjects,
  sqlite,
  webhookCleanup,
} from "../db.ts";

import {
  AF_DB,
  AF_REPORTS,
  AF_VHOST,
  CORS_HEADERS,
  dbHeaders,
  PROXY_APP_VERSION,
  PROXY_ADMIN_KEY,
  reportsHeaders,
  STORAGE_BUDGET_BYTES,
  TURSO_URL,
} from "../config.ts";

import { fetchWithTimeout } from "../lib/fetchUtils.ts";
import { handleWorkOrders } from "./workOrders.ts";
import { handleVendors } from "./vendors.ts";
import { handleInspections } from "./inspections.ts";
import { handleBills } from "./bills.ts";
import {
  handleProperties,
  handlePropertyGroups,
  handlePropertyMap,
  handleUpcomingMoveouts,
} from "./properties.ts";
import { handleTurns } from "./turns.ts";
import {
  handleLabor,
  handleRecentTasks,
  handleTurnWorkOrders,
} from "./workOrders.ts";
import { handleWoComparisonReport } from "./woComparison.ts";

const passthroughMutationLocks = new Map<string, Promise<void>>();

async function withPassthroughMutationLock<T>(
  lockKey: string,
  fn: () => Promise<T>,
): Promise<T> {
  while (passthroughMutationLocks.has(lockKey)) {
    await passthroughMutationLocks.get(lockKey);
  }

  let release!: () => void;
  const gate = new Promise<void>((resolve) => {
    release = resolve;
  });
  passthroughMutationLocks.set(lockKey, gate);

  try {
    return await fn();
  } finally {
    passthroughMutationLocks.delete(lockKey);
    release();
  }
}

// ── handlePassthrough ─────────────────────────────────────────────────────────
// Routes a request directly to the AppFolio API and streams the response back.
// Used by the HandyManager frontend for ad-hoc API exploration and by the
// HandyManager cockpit's "Raw API" panel.
//
// ?path=/api/v0/...   → routed to AF_DB with dbHeaders()
// ?path=/api/v2/...   → routed to AF_REPORTS with reportsHeaders()
// ?path=https://...   → absolute URL — domain stripped, path re-routed as above
//
// Returns a raw Response (not JSON) so the caller in main.ts must check
// `if (result instanceof Response) return result` before corsJson().
//
// Ensure the Client ID and Client Secret in the authorization header are
// correctly formatted . Is the URL correct before proxying .
export async function handlePassthrough(
  params: Record<string, string>,
  req: Request,
): Promise<any> {
  const targetPath = params.path;
  if (!targetPath) return { ok: false, error: "Missing ?path= parameter" };

  // Strip domain from absolute URLs — re-route based on path prefix
  let cleanPath = targetPath;
  if (targetPath.startsWith("http")) {
    try {
      const u = new URL(targetPath);
      cleanPath = u.pathname + u.search;
    } catch { /* use as-is */ }
  }

  // Route to the correct domain AND credentials based on API version prefix
  const isV0 = cleanPath.startsWith("/api/v0");
  const baseDomain = isV0 ? AF_DB : AF_REPORTS;
  const targetUrl = `${baseDomain}${cleanPath}`;

  const hdrs: Record<string, string> = isV0
    ? { ...dbHeaders() }
    : { ...reportsHeaders() };

  // Forward Content-Type if the original request set one
  const ct = req.headers.get("content-type");
  if (ct) hdrs["Content-Type"] = ct;

  const bodyText = req.method !== "GET" && req.method !== "HEAD"
    ? await req.text()
    : undefined;
  const doFetch = () =>
    fetchWithTimeout(targetUrl, {
      method: req.method,
      headers: hdrs,
      body: bodyText,
    });

  const isMutation = req.method !== "GET" && req.method !== "HEAD";
  const lockKey = `${req.method.toUpperCase()}:${cleanPath.split("?")[0]}`;
  const resp = isMutation
    ? await withPassthroughMutationLock(lockKey, doFetch)
    : await doFetch();

  const body = await resp.text();

  // Return raw Response — preserves status code and Content-Type from AF
  return new Response(body, {
    status: resp.status,
    statusText: resp.statusText,
    headers: {
      "Content-Type": resp.headers.get("Content-Type") || "application/json",
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Headers": "*",
    },
  });
}

// ── handlePing ────────────────────────────────────────────────────────────────
// Health check — fires one request against each AF API domain sequentially.
// Returns ok=true if EITHER domain is reachable (not both required).
// Used by the HandyManager header status indicator.
//
// The Webhook URL must be secured by HTTPS and the domain must be owned by
// your company . The ping confirms the proxy can reach AppFolio —
// it does not validate the webhook endpoint separately.
export async function handlePing(): Promise<any> {
  const t0 = Date.now();

  const dbResp = await Promise.allSettled([
    fetch(
      `${AF_DB}/api/v0/properties?filters%5BLastUpdatedAtFrom%5D=2025-01-01T00%3A00%3A00Z&page%5Bsize%5D=1`,
      { headers: dbHeaders() },
    ),
  ]).then((r) => r[0]);
  const rptResp = await Promise.allSettled([
    fetch(
      `${AF_REPORTS}/api/v2/reports/property_directory.json`,
      {
        method: "POST",
        headers: { ...reportsHeaders(), "Content-Type": "application/json" },
        body: JSON.stringify({
          property_visibility: "active",
          paginate_results: true,
          per_page: 1,
        }),
      },
    ),
  ]).then((r) => r[0]);

  const elapsed = Date.now() - t0;
  const dbOk = dbResp.status === "fulfilled" && dbResp.value.ok;
  const rptOk = rptResp.status === "fulfilled" && rptResp.value.ok;

  let dbDetail = "";
  let rptDetail = "";
  try {
    if (dbResp.status === "fulfilled") {
      dbDetail = (await dbResp.value.text()).substring(0, 200);
    }
  } catch {}
  try {
    if (rptResp.status === "fulfilled") {
      rptDetail = (await rptResp.value.text()).substring(0, 100);
    }
  } catch {}

  let brandName = "Fort Lowell Realty";
  let brandLogoUrl = "";
  let appVersion = PROXY_APP_VERSION;
  try {
    const cfgRows = rowsAsObjects(
      await sqlite.execute(
        `SELECT key, value FROM proxy_config WHERE key IN ('brand_name', 'brand_logo_url', 'app_version')`,
      ),
    );
    cfgRows.forEach((row: any) => {
      if (row.key === "brand_name") brandName = String(row.value || brandName);
      if (row.key === "brand_logo_url") brandLogoUrl = String(row.value || "");
      if (row.key === "app_version") {
        appVersion = String(row.value || appVersion);
      }
    });
  } catch {
    // Non-fatal — branding is optional.
  }

  // Lightweight schema health checks to surface deployment drift quickly.
  let schemaHealth: any = {
    ok: true,
    checked_tables: [
      "trusted_devices",
      "device_otps",
      "pm_proxy_users",
      "proxy_config",
      "webhook_events",
    ],
    missing_tables: [],
  };
  try {
    const tableRows = rowsAsObjects(
      await sqlite.execute(
        `SELECT name FROM sqlite_master WHERE type='table'`,
      ),
    );
    const existing = new Set(tableRows.map((r: any) => String(r.name || "")));
    const missing = schemaHealth.checked_tables.filter(function (t: string) {
      return !existing.has(t);
    });
    schemaHealth.missing_tables = missing;
    schemaHealth.ok = missing.length === 0;
  } catch (e: any) {
    schemaHealth.ok = false;
    schemaHealth.error = String(e?.message || e);
  }

  return {
    ok: dbOk || rptOk,
    db_api: {
      ok: dbOk,
      domain: AF_DB,
      status: dbResp.status === "fulfilled" ? dbResp.value.status : 0,
      detail: dbOk ? "Connected" : dbDetail,
    },
    reports_api: {
      ok: rptOk,
      domain: AF_REPORTS,
      status: rptResp.status === "fulfilled" ? rptResp.value.status : 0,
      detail: rptOk ? "Connected" : rptDetail,
    },
    latency_ms: elapsed,
    timestamp: new Date().toISOString(),
    version: appVersion,
    proxy: appVersion,
    database: TURSO_URL ? "turso" : "valtown_sqlite",
    schema: schemaHealth,
    vhost: AF_VHOST,
    brand: {
      name: brandName,
      logo_url: brandLogoUrl,
    },
  };
}

// ── handleCacheStats ──────────────────────────────────────────────────────────
// Returns a full snapshot of the Turso/SQLite storage state:
//   • Per-entity cache summary (entries, last cached, total records)
//   • Webhook event counts (total, pending, max id)
//   • Turn record count
//   • v9.0 reassignment engine table counts
//   • Storage usage vs budget breakdown
export async function handleCacheStats(): Promise<any> {
  const [
    cacheResult,
    webhookCount,
    turnCount,
    reassignCount,
    techCount,
    tokenCount,
    auditCount,
    commsCount,
  ] = await Promise.all([
    sqlite.execute(
      `SELECT entity_type,
              COUNT(*)          AS entries,
              MAX(cached_at)    AS last_cached,
              SUM(record_count) AS total_records
       FROM api_cache
       GROUP BY entity_type`,
    ),
    sqlite.execute(
      `SELECT COUNT(*) AS total,
              SUM(CASE WHEN processed = 0 THEN 1 ELSE 0 END) AS pending_unprocessed,
              SUM(CASE WHEN processed = 0 AND length(raw_body) > 2 THEN 1 ELSE 0 END) AS pending,
              SUM(CASE WHEN processed = 0 AND length(raw_body) <= 2 THEN 1 ELSE 0 END) AS pending_empty,
              MAX(id) AS max_id
       FROM webhook_events`,
    ),
    sqlite.execute(`SELECT COUNT(*) AS total FROM turn_records`),
    sqlite.execute(`SELECT COUNT(*) AS total FROM reassignment_queue`),
    sqlite.execute(`SELECT COUNT(*) AS total FROM tech_grades`),
    sqlite.execute(`SELECT COUNT(*) AS total FROM magic_link_tokens`),
    sqlite.execute(`SELECT COUNT(*) AS total FROM wo_audit_log`),
    sqlite.execute(`SELECT COUNT(*) AS total FROM tenant_comms_log`),
  ]);

  const whRows = rowsAsObjects(webhookCount);
  const cacheBytes = await getCacheSizeBytes();
  const webhookBytes = await getWebhookSizeBytes();
  const totalBytes = cacheBytes + webhookBytes;

  return {
    ok: true,
    version: PROXY_APP_VERSION,
    database: TURSO_URL ? "turso" : "valtown_sqlite",

    // v8.9 core tables
    cache: rowsAsObjects(cacheResult),
    webhooks: whRows[0] || { total: 0, pending: 0, max_id: 0 },
    turn_records: rowsAsObjects(turnCount)[0]?.total || 0,

    // v9.0 engine tables
    v9: {
      reassignment_queue: rowsAsObjects(reassignCount)[0]?.total || 0,
      tech_grades: rowsAsObjects(techCount)[0]?.total || 0,
      magic_link_tokens: rowsAsObjects(tokenCount)[0]?.total || 0,
      wo_audit_log: rowsAsObjects(auditCount)[0]?.total || 0,
      tenant_comms_log: rowsAsObjects(commsCount)[0]?.total || 0,
    },

    // Storage budget breakdown
    storage: {
      cache_bytes: cacheBytes,
      webhook_bytes: webhookBytes,
      total_bytes: totalBytes,
      total_mb: (totalBytes / 1024 / 1024).toFixed(2),
      budget_mb: (STORAGE_BUDGET_BYTES / 1024 / 1024).toFixed(1),
      usage_pct: ((totalBytes / STORAGE_BUDGET_BYTES) * 100).toFixed(1),
    },
  };
}

// ── handleCacheInvalidate ─────────────────────────────────────────────────────
// Manually evict one entity type from api_cache, or wipe the entire cache
// if no ?type= param is provided.
// Used by the HandyManager cache management panel.
export async function handleCacheInvalidate(
  params: Record<string, string>,
): Promise<any> {
  const entityType = params.type || "";

  if (!entityType) {
    await sqlite.execute(`DELETE FROM api_cache`);
    return { ok: true, message: "All caches cleared" };
  }

  await cacheInvalidate(entityType);
  return { ok: true, message: `Cache cleared for entity type: ${entityType}` };
}

// ── handleForceRefresh ────────────────────────────────────────────────────────
// Invalidate a specific entity type's cache and immediately re-fetch from
// AppFolio so the next request is served warm.
// Supports all v8.9 and v9.0 entity types.
// If an incorrect parameter value was passed, AppFolio error codes are
// descriptive enough to identify the problem .
export async function handleForceRefresh(
  params: Record<string, string>,
): Promise<any> {
  const entityType = params.type || "";
  if (!entityType) return { ok: false, error: "Missing ?type= parameter" };

  await cacheInvalidate(entityType);

  switch (entityType) {
    case "work_orders":
      return handleWorkOrders(params);
    case "turns":
      return handleTurns(params);
    case "turn_work_orders":
      return handleTurnWorkOrders(params);
    case "vendors":
      return handleVendors(params);
    case "inspections":
      return handleInspections(params);
    case "properties":
      return handleProperties(params);
    case "property_groups":
      return handlePropertyGroups(params);
    case "property_map":
      return handlePropertyMap(params);
    case "recent_tasks":
      return handleRecentTasks(params);
    case "upcoming_moveouts":
      return handleUpcomingMoveouts(params);
    case "labor":
      return handleLabor(params);
    case "bills":
      return handleBills(params);
    case "wo_comparison":
      return handleWoComparisonReport(params);
    default:
      return { ok: false, error: `Unknown entity type: "${entityType}"` };
  }
}

// ── handleStorageCleanup ──────────────────────────────────────────────────────
// Manual trigger for storage housekeeping:
//   1. Purge all expired api_cache rows
//   2. Delete webhook events older than WEBHOOK_MAX_DAYS
//   3. Trim webhook_events to WEBHOOK_MAX_EVENTS (keep newest)
// Reports bytes freed and webhook rows deleted.
export async function handleStorageCleanup(): Promise<any> {
  const before = await getCacheSizeBytes();

  // 1. Purge expired cache entries
  await sqlite.execute(
    `DELETE FROM api_cache WHERE datetime(expires_at) < datetime('now')`,
  );

  // 2 + 3. Webhook housekeeping (trim by age then by count)
  const wh = await webhookCleanup();

  const after = await getCacheSizeBytes();

  return {
    ok: true,
    freed_bytes: before - after,
    freed_mb: ((before - after) / 1024 / 1024).toFixed(2),
    webhook_deleted: {
      old_events: wh.deleted_old,
      overflow_events: wh.deleted_overflow,
    },
    cache_bytes_now: after,
  };
}

// ── handleDebugSqlite ─────────────────────────────────────────────────────────
// Diagnostic endpoint — dumps raw PRAGMA and sqlite_master data so the
// developer can verify the actual row format and schema without needing
// direct database access.
//
// Includes PRAGMA table_info for all v8.9 and v9.0 tables.
// Useful when diagnosing rowsAsObjects() issues between Val Town global.ts
// (array rows) and Turso (object rows).
export async function handleDebugSqlite(): Promise<any> {
  const diag: any = { ok: true, timestamp: new Date().toISOString() };

  // All tables in the DB
  try {
    const tables = await sqlite.execute(
      `SELECT name, sql FROM sqlite_master WHERE type='table' ORDER BY name`,
    );
    diag.tables_obj = rowsAsObjects(tables);
  } catch (e: any) {
    diag.tables_error = e.message;
  }

  // All indexes
  try {
    const indexes = await sqlite.execute(
      `SELECT name, tbl_name FROM sqlite_master WHERE type='index' ORDER BY tbl_name, name`,
    );
    diag.indexes_obj = rowsAsObjects(indexes);
  } catch (e: any) {
    diag.indexes_error = e.message;
  }

  // PRAGMA table_info for every table we manage
  const ALL_TABLES = [
    // v8.9 core
    "api_cache",
    "turn_records",
    "webhook_events",
    "trusted_devices",
    // v9.0 engine
    "reassignment_queue",
    "magic_link_tokens",
    "tenant_comms_log",
    "wo_audit_log",
    "tech_grades",
    "blast_events",
    "tier2_claims",
    "proxy_config",
  ];

  for (const tbl of ALL_TABLES) {
    try {
      const info = await sqlite.execute(`PRAGMA table_info(${tbl})`);
      diag[`pragma_${tbl}`] = rowsAsObjects(info);
    } catch (e: any) {
      diag[`pragma_${tbl}_error`] = e.message;
    }
  }

  // Spot-check one row from key tables
  for (
    const tbl of [
      "webhook_events",
      "api_cache",
      "tech_grades",
      "reassignment_queue",
    ]
  ) {
    try {
      const sample = await sqlite.execute(`SELECT * FROM ${tbl} LIMIT 1`);
      diag[`sample_${tbl}`] = rowsAsObjects(sample);
    } catch (e: any) {
      diag[`sample_${tbl}_error`] = e.message;
    }
  }

  diag._tablesReady = true;
  diag._database = TURSO_URL ? "turso" : "valtown_sqlite";

  return diag;
}

// ── handleSqlQuery ────────────────────────────────────────────────────────────
// Admin endpoint — executes a read-only SQL statement and returns rows.
// Guarded by PROXY_ADMIN_KEY. Must be called via POST with a JSON body
// containing { key, query } — the admin key is NEVER accepted in URL params
// to prevent it appearing in server logs, browser history, or CDN traces.
//
// Allowed statements: SELECT, PRAGMA, EXPLAIN, WITH
// Blocked: everything else (use handleSqlExecute for writes)
export async function handleSqlQuery(
  _params: Record<string, string>,
  req: Request,
): Promise<any> {
  if (!PROXY_ADMIN_KEY) {
    return {
      ok: false,
      error:
        "sql_query is disabled — set PROXY_ADMIN_KEY in Val Town env vars to enable",
    };
  }

  // Key and query MUST come from the JSON body — never from URL params.
  let key = "";
  let query = "";
  try {
    const ct = req.headers.get("content-type") || "";
    if (ct.includes("application/json")) {
      const body = await req.json();
      key = body.key || "";
      query = body.query || body.sql || "";
    }
  } catch { /* keep empty — validation below will catch missing fields */ }

  if (!key || key !== PROXY_ADMIN_KEY) {
    return {
      ok: false,
      error: "Unauthorized: admin key required in request body",
    };
  }
  if (!query) {
    return { ok: false, error: "Missing query or sql in request body" };
  }

  // Allowlist — read-only statements only
  const normalized = query.trim().replace(/\s+/g, " ").toUpperCase();
  const allowed = normalized.startsWith("SELECT") ||
    normalized.startsWith("PRAGMA") ||
    normalized.startsWith("EXPLAIN") ||
    normalized.startsWith("WITH");

  if (!allowed) {
    return {
      ok: false,
      error:
        "sql_query only allows SELECT / PRAGMA / EXPLAIN / WITH. Use sql_execute for writes.",
    };
  }

  try {
    const result = await sqlite.execute(query);
    const columns = result.columns || [];
    const rows = rowsAsObjects(result);
    return { ok: true, rows, columns, count: rows.length };
  } catch (err: any) {
    return { ok: false, error: err.message };
  }
}

// ── handleSqlExecute ──────────────────────────────────────────────────────────
// Admin endpoint — executes a write SQL statement (INSERT, UPDATE, DELETE, etc.)
// Guarded by PROXY_ADMIN_KEY. Must be called via POST with a JSON body
// containing { key, sql, args? } — the admin key is NEVER accepted in URL
// params to prevent it appearing in server logs or CDN traces.
//
// SELECT and EXPLAIN are blocked here — use handleSqlQuery for reads.
//
// Supports parameterised queries:
//   { "key": "...", "sql": "UPDATE tech_grades SET tier=? WHERE tech_id=?", "args": [2, "uuid"] }
export async function handleSqlExecute(
  _params: Record<string, string>,
  req: Request,
): Promise<any> {
  if (!PROXY_ADMIN_KEY) {
    return {
      ok: false,
      error:
        "sql_execute is disabled — set PROXY_ADMIN_KEY in Val Town env vars to enable",
    };
  }

  // Key, sql, and args MUST come from the JSON body — never from URL params.
  let key = "";
  let sqlStmt = "";
  let sqlArgs: any[] = [];

  try {
    const ct = req.headers.get("content-type") || "";
    if (ct.includes("application/json")) {
      const body = await req.json();
      key = body.key || "";
      sqlStmt = body.sql || body.query || "";
      sqlArgs = body.args || body.params || [];
    }
  } catch { /* keep empty — validation below will catch missing fields */ }

  if (!key || key !== PROXY_ADMIN_KEY) {
    return {
      ok: false,
      error: "Unauthorized: admin key required in request body",
    };
  }
  if (!sqlStmt) {
    return { ok: false, error: "Missing sql or query in request body" };
  }

  // Block SELECT via this endpoint — keeps the separation clean
  const normalized = sqlStmt.trim().toUpperCase();
  if (normalized.startsWith("SELECT") || normalized.startsWith("EXPLAIN")) {
    return {
      ok: false,
      error: "Use ?action=sql_query for SELECT / EXPLAIN statements",
    };
  }

  try {
    const result = await sqlite.execute({ sql: sqlStmt, args: sqlArgs });
    return {
      ok: true,
      rowsAffected: result.rowsAffected || 0,
      lastInsertRowid: result.lastInsertRowid != null
        ? Number(result.lastInsertRowid)
        : null,
    };
  } catch (err: any) {
    return { ok: false, error: err.message };
  }
}

// ── handleGenericReport ───────────────────────────────────────────────────────
// Pass-any-report-name endpoint — used by the HandyManager "Raw Report" panel
// and by any one-off report pulls that don't have a dedicated handler.
//
// GET:  ?action=report&name=work_order_labor_summary
// POST: same action, JSON body = filter object
//
// If an incorrect parameter value was passed to AppFolio, the returned error
// codes should be descriptive enough to identify the problem .
export async function handleGenericReport(
  params: Record<string, string>,
  req: Request,
): Promise<any> {
  const reportName = params.name || "";
  if (!reportName) {
    return { ok: false, error: "Missing ?name= parameter (report name)" };
  }

  let filters: Record<string, any> = {};
  if (req.method === "POST") {
    try {
      filters = await req.json();
    } catch { /* empty filters */ }
  }

  const { fetchReport } = await import("../lib/appfolio.ts");
  const rows = await fetchReport(reportName, filters);
  return { ok: true, count: rows.length, results: rows };
}