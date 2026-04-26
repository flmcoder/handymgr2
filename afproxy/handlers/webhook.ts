// ============================================================================
// handlers/webhook.ts — AppFolio webhook ingestion + query endpoints.
//
// AppFolio delivers all webhook events as HTTPS POST requests to the
// registered Webhook URL . The server must be behind the HTTPS URL
// configured in the AppFolio Admin Settings webhook card .
//
// AppFolio signs all outgoing webhooks to enable recipients to verify the
// authenticity and integrity of received notifications . Payloads
// arrive in one of two formats:
//
//   JWS (compact):  header.payload.signature  (3 dot-separated base64url parts)
//   Plain JSON:     standard JSON body
//
// This handler attempts JWS decode first, falls back to plain JSON, and
// stores the raw body verbatim regardless of parse outcome so the signature
// can always be verified later downstream .
//
// After storing the event, relevant api_cache entity types are invalidated
// so the next HandyManager data request fetches fresh data from AppFolio.
//
// Webhook URL setup:
//   1. Admin Settings → Webhook Card → configure HTTPS URL
//   2. Subscribe to desired topics in Topic Subscriptions
//   3. Use Send Test Event to verify the URL is reachable
//   4. Check Webhook Logs in AppFolio to confirm delivery
//
// Exports:
//   handleWebhookPost     — ingest incoming AF webhook POST
//   handleWebhookEvents   — query stored webhook events with filters
//   handleWebhookStats    — aggregate counts + breakdown by type/day
//   handleWebhookResolve  — resolve resource_type + resource_id to live record
//   handleWebhookMigrate  — blob → SQLite migration (v7 → v8 one-time)
//   handleMigrateV8       — mark old empty webhook_events as processed
// ============================================================================

import {
  cacheInvalidate,
  rowsAsObjects,
  sqlite,
  webhookCleanup,
} from "../db.ts";
import {
  CORS_HEADERS,
  AF_DB,
  AF_REPORTS,
  dbHeaders,
  WEBHOOK_CACHE_MAP,
  WEBHOOK_MAX_DAYS,
} from "../config.ts";
import {
  invalidateCacheForGroup,
  processWebhookSync,
  rebuildGroupResolutionCache,
  resolveGroupsForProperty,
  resolveRoutingContext,
  resolveWebhookResource,
  syncBillingMap,
  syncPropertyGroups,
  syncPropertyMap,
  syncVendorMap,
  syncWorkOrderMap,
} from "../lib/appfolio.ts";
import { fetchWithTimeout } from "../lib/fetchUtils.ts";
import { blob } from "https://esm.town/v/std/blob";
import {
  base64url,
  compactVerify,
  createRemoteJWKSet,
  decodeProtectedHeader,
  errors as joseErrors,
} from "npm:jose@5.9.6";

const APPFOLIO_JWKS_URL = Deno.env.get("APPFOLIO_JWKS_URL") ||
  "https://api.appfolio.com/.well-known/jwks.json";
const APPFOLIO_JWKS = createRemoteJWKSet(new URL(APPFOLIO_JWKS_URL));
const ADMIN_SECRET = (Deno.env.get("ADMIN_SECRET") || "").trim();
const WEBHOOK_LIVE_CACHE_TTL = 8000;

let webhookLiveCache:
  | {
    data: any;
    timestamp: number;
    sinceId: number;
  }
  | null = null;

function adminJson(
  body: unknown,
  status = 200,
  extraHeaders: Record<string, string> = {},
): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      "Content-Type": "application/json",
      ...CORS_HEADERS,
      ...extraHeaders,
    },
  });
}

function isRateLimitedSyncError(err: unknown): boolean {
  return !!(
    err && typeof err === "object" &&
    (err as Record<string, unknown>).rateLimited === true
  );
}

function isSqlWriteBlockedError(err: unknown): boolean {
  const msg = String((err as any)?.message || err || "").toLowerCase();
  return msg.includes("write operations are forbidden") ||
    msg.includes("operation was blocked") ||
    (msg.includes("blocked") && msg.includes("sql write"));
}

function requireAdminAuth(req: Request): Response | null {
  const adminToken = (req.headers.get("X-Admin-Token") || "").trim();
  if (ADMIN_SECRET && adminToken === ADMIN_SECRET) return null;

  const legacyAdminKey = (req.headers.get("x-admin-key") || "").trim();
  if (legacyAdminKey) {
    const proxyAdminKey = (Deno.env.get("PROXY_ADMIN_KEY") || "").trim();
    if (proxyAdminKey && legacyAdminKey === proxyAdminKey) return null;
  }

  return adminJson({ ok: false, error: "Unauthorized" }, 401);
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function normalizeResourceType(raw: string | null | undefined): string {
  let t = String(raw || "").trim().toLowerCase();
  if (!t) return "";
  if (t.includes(".")) t = t.split(".")[0];
  t = t.replace(/[^a-z0-9_]/g, "");
  if (t.endsWith("ies")) t = t.slice(0, -3) + "y";
  else if (t.endsWith("s")) t = t.slice(0, -1);
  return t;
}

async function fetchDbPath(path: string): Promise<{
  ok: boolean;
  status: number;
  payload?: any;
  detail?: string;
}> {
  let resp = await fetchWithTimeout(`${AF_REPORTS}${path}`, {
    headers: dbHeaders(),
  });
  if (resp.status === 401 || resp.status === 403 || resp.status === 404) {
    resp = await fetchWithTimeout(`${AF_DB}${path}`, { headers: dbHeaders() });
  }
  if (!resp.ok) {
    const detail = await resp.text().catch(() => "");
    return {
      ok: false,
      status: resp.status,
      detail: detail.substring(0, 400),
    };
  }
  const payload = await resp.json().catch(() => ({}));
  return { ok: true, status: resp.status, payload };
}

function firstRecord(payload: any): any {
  return (payload?.data && !Array.isArray(payload.data)
    ? payload.data
    : null) ||
    (payload?.data && Array.isArray(payload.data) ? payload.data[0] : null) ||
    payload?.results?.[0] ||
    (Array.isArray(payload) ? payload[0] : null) ||
    null;
}

function extractLatestNoteTimestamp(notes: any[]): string {
  if (!Array.isArray(notes) || notes.length === 0) return "";
  let latestIso = "";
  let latestMs = 0;
  for (const n of notes) {
    const ts = String(n?.CreatedAt || n?.created_at || n?.createdAt || "");
    if (!ts) continue;
    const ms = new Date(ts).getTime();
    if (isNaN(ms)) continue;
    if (ms > latestMs) {
      latestMs = ms;
      latestIso = new Date(ms).toISOString();
    }
  }
  return latestIso;
}

function extractTotalPages(payload: any): number {
  const metaCandidates = [
    payload?.meta?.total_pages,
    payload?.metadata?.total_pages,
    payload?.pagination?.total_pages,
    payload?.pages,
    payload?.total_pages,
  ];
  for (const c of metaCandidates) {
    const n = Number(c);
    if (!isNaN(n) && n > 0) return Math.floor(n);
  }
  return 1;
}

async function fetchLatestWorkOrderNoteAt(
  workOrderId: string,
): Promise<string> {
  const encodedId = encodeURIComponent(workOrderId);

  // Preferred path: descending sort should return newest first.
  const sortedPath =
    `/api/v0/work_orders/${encodedId}/notes?sort=-CreatedAt&page[size]=1`;
  const sortedResp = await fetchDbPath(sortedPath);
  if (sortedResp.ok) {
    const sortedNotes = Array.isArray(sortedResp.payload?.data)
      ? sortedResp.payload.data
      : (Array.isArray(sortedResp.payload?.results)
        ? sortedResp.payload.results
        : (Array.isArray(sortedResp.payload) ? sortedResp.payload : []));
    const newestFromSorted = extractLatestNoteTimestamp(sortedNotes);
    if (newestFromSorted) return newestFromSorted;
  }

  // Fallback for endpoints that ignore/reject sort.
  // If oldest-first pagination is used, newest is on the final page.
  const basePath = `/api/v0/work_orders/${encodedId}/notes?page[size]=50`;
  const firstPageResp = await fetchDbPath(basePath);
  if (!firstPageResp.ok) return "";

  const firstNotes = Array.isArray(firstPageResp.payload?.data)
    ? firstPageResp.payload.data
    : (Array.isArray(firstPageResp.payload?.results)
      ? firstPageResp.payload.results
      : (Array.isArray(firstPageResp.payload) ? firstPageResp.payload : []));

  const totalPages = extractTotalPages(firstPageResp.payload);
  if (totalPages <= 1) {
    return extractLatestNoteTimestamp(firstNotes);
  }

  const lastPageResp = await fetchDbPath(
    `${basePath}&page[number]=${totalPages}`,
  );
  if (!lastPageResp.ok) {
    return extractLatestNoteTimestamp(firstNotes);
  }

  const lastNotes = Array.isArray(lastPageResp.payload?.data)
    ? lastPageResp.payload.data
    : (Array.isArray(lastPageResp.payload?.results)
      ? lastPageResp.payload.results
      : (Array.isArray(lastPageResp.payload) ? lastPageResp.payload : []));
  return extractLatestNoteTimestamp(lastNotes) ||
    extractLatestNoteTimestamp(firstNotes);
}

async function fetchWorkOrderById(resourceId: string): Promise<{
  ok: boolean;
  record?: any;
  status?: number;
  detail?: string;
}> {
  const byId = await fetchDbPath(
    `/api/v0/work_orders/${encodeURIComponent(resourceId)}`,
  );
  if (byId.ok) {
    const rec = firstRecord(byId.payload);
    if (rec) return { ok: true, record: rec, status: byId.status };
  }

  // Fallback to filter lookup for compatibility with UUID-only resources.
  const byFilter = await fetchDbPath(
    `/api/v0/work_orders?filters[Id]=${
      encodeURIComponent(resourceId)
    }&page[size]=1`,
  );
  if (!byFilter.ok) {
    return {
      ok: false,
      status: byFilter.status,
      detail: byFilter.detail,
    };
  }
  const rec = firstRecord(byFilter.payload);
  if (!rec) {
    return { ok: false, status: 404, detail: "work_order_not_found" };
  }
  return { ok: true, record: rec, status: byFilter.status };
}

function extractWorkOrderState(record: any): {
  id: string;
  status_code: number | null;
  status_text: string;
  assigned_user_id: string;
  assigned_user_name: string;
  last_activity_at: string;
  last_note_at: string;
  wo_number: string;
  property_address: string;
  raw_snapshot: string;
} {
  const assignedUsers = Array.isArray(record?.AssignedUsers)
    ? record.AssignedUsers
    : (Array.isArray(record?.assigned_users) ? record.assigned_users : []);
  const firstAssigned = assignedUsers[0] || {};

  const rawStatusCode = record?.StatusCode ?? record?.status_code ??
    record?.StatusId ?? record?.status_id ?? null;
  const parsedStatusCode = rawStatusCode === null || rawStatusCode === undefined
    ? null
    : (Number(rawStatusCode) || null);

  return {
    id: String(record?.Id || record?.id || ""),
    status_code: parsedStatusCode,
    status_text: String(
      record?.Status || record?.status || record?.WorkOrderStatus ||
        record?.work_order_status || "",
    ),
    assigned_user_id: String(
      record?.AssignedUserId || record?.assigned_user_id || firstAssigned?.Id ||
        firstAssigned?.id || "",
    ),
    assigned_user_name: String(
      record?.AssignedUserName || record?.assigned_user_name ||
        firstAssigned?.Name || firstAssigned?.name || "",
    ),
    last_activity_at: String(
      record?.LastUpdatedAt || record?.last_updated_at || record?.UpdatedAt ||
        record?.updated_at || record?.CreatedAt || record?.created_at ||
        new Date().toISOString(),
    ),
    last_note_at: "",
    wo_number: String(
      record?.WorkOrderNumber || record?.work_order_number || record?.Number ||
        record?.number || "",
    ),
    property_address: String(
      record?.PropertyAddress || record?.property_address ||
        record?.Address || record?.address || "",
    ),
    raw_snapshot: JSON.stringify(record || {}),
  };
}

async function upsertWorkOrderState(
  state: ReturnType<typeof extractWorkOrderState>,
  eventType: string | null,
  resourceType: string,
): Promise<void> {
  await sqlite.execute({
    sql: `INSERT INTO wo_states (
            id, status_code, status_text, assigned_user_id, assigned_user_name,
            last_activity_at, last_note_at, event_type, resource_type, wo_number,
            property_address, raw_snapshot, fetched_at, updated_at
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'), datetime('now'))
          ON CONFLICT(id) DO UPDATE SET
            status_code = excluded.status_code,
            status_text = excluded.status_text,
            assigned_user_id = excluded.assigned_user_id,
            assigned_user_name = excluded.assigned_user_name,
            last_activity_at = excluded.last_activity_at,
            last_note_at = excluded.last_note_at,
            event_type = excluded.event_type,
            resource_type = excluded.resource_type,
            wo_number = excluded.wo_number,
            property_address = excluded.property_address,
            raw_snapshot = excluded.raw_snapshot,
            fetched_at = datetime('now'),
            updated_at = datetime('now')`,
    args: [
      state.id,
      state.status_code,
      state.status_text,
      state.assigned_user_id,
      state.assigned_user_name,
      state.last_activity_at,
      state.last_note_at,
      eventType,
      resourceType,
      state.wo_number,
      state.property_address,
      state.raw_snapshot,
    ],
  });
}

async function getPreviousWorkOrderState(workOrderId: string): Promise<
  {
    id: string;
    assigned_user_id: string;
    status_code: number | null;
  } | null
> {
  const result = await sqlite.execute({
    sql: `SELECT id, assigned_user_id, status_code
          FROM wo_states
          WHERE id = ?
          LIMIT 1`,
    args: [workOrderId],
  });
  const row = rowsAsObjects(result)[0];
  if (!row) return null;
  return {
    id: String(row.id || ""),
    assigned_user_id: String(row.assigned_user_id || ""),
    status_code: row.status_code === null || row.status_code === undefined
      ? null
      : (Number(row.status_code) || null),
  };
}

async function insertWorkOrderStateChange(
  workOrderId: string,
  changeType: "reassigned" | "status_changed" | "note_added",
  oldValue: string,
  newValue: string,
): Promise<void> {
  await sqlite.execute({
    sql: `INSERT INTO wo_state_changes
            (work_order_id, change_type, old_value, new_value, created_at)
          VALUES (?, ?, ?, ?, datetime('now'))`,
    args: [workOrderId, changeType, oldValue, newValue],
  });
}

// ============================================================================
// JWS / JSON PARSING
// ============================================================================
// ── interpretWebhookEvent ─────────────────────────────────────────────────────
// Converts a raw decoded AppFolio webhook payload into a structured,
// human-readable event description.
//
// AppFolio signs all outgoing webhooks . After JWS decode the payload
// contains: topic, event/event_type, resource_type, resource_id, and optional
// data fields depending on the topic.
//
// Webhook event logs include a record of every webhook attempt and the payload
// and which topic the notification was sent to .
// ── handleWebhookLive ─────────────────────────────────────────────────────────
// Optimised polling endpoint for the frontend live feed.
// Returns only events newer than since_id, fully decoded.
// Webhooks make it easier to adhere to rate limits by removing the need
// to periodically poll for changes .
//
// The frontend polls this every 8 seconds. Each response includes the
// highest event ID seen so the next poll can pass since_id correctly.
//
// GET ?action=webhook_live&since_id=LAST_SEEN_ID&limit=15
export async function handleWebhookLive(
  params: Record<string, string>,
): Promise<any> {
  const sinceId = parseInt(params.since_id || "0", 10);
  const limit = parseInt(params.limit || "15", 10);

  // Check cache first
  if (webhookLiveCache && 
      (Date.now() - webhookLiveCache.timestamp) < WEBHOOK_LIVE_CACHE_TTL &&
      webhookLiveCache.sinceId === sinceId) {
    return webhookLiveCache.data;
  }

  try {
    const rs = await sqlite.execute({
      sql: `SELECT id, received_at, raw_body, event_type,
                    resource_type, resource_id
             FROM   webhook_events
             WHERE  id > ?
               AND  length(raw_body) > 2
             ORDER  BY id DESC
             LIMIT  ?`,
      args: [sinceId, limit],
    });

    const rows = rowsAsObjects(rs);
    const events = rows.map((r: any) => {
      // Re-parse the raw body to extract all available fields
      const parsed = parseIncomingWebhook(r.raw_body || "");

      // Run the full interpreter
      const interpreted = interpretWebhookEvent(
        parsed.payload,
        parsed.eventType || r.event_type || null,
        parsed.resourceType || r.resource_type || null,
        parsed.resourceId || r.resource_id || null,
      );

      return {
        id: r.id,
        received_at: r.received_at,
        event_id: parsed.eventId,
        ...interpreted,
      };
    });

    const maxId = rows.length > 0
      ? Math.max(...rows.map((r: any) => Number(r.id)))
      : sinceId;

    const result = {
      ok: true,
      events,
      count: events.length,
      max_id: maxId,
      has_new: events.length > 0,
    };

    // Cache the result
    webhookLiveCache = {
      data: result,
      timestamp: Date.now(),
      sinceId
    };

    return result;
  } catch (e: any) {
    return { ok: false, error: e.message, events: [], max_id: sinceId };
  }
}
export function interpretWebhookEvent(
  payload: any,
  eventType: string | null,
  resourceType: string | null,
  resourceId: string | null,
): {
  title: string; // Short human-readable title e.g. "Work Order Updated"
  description: string; // Detail e.g. "WO #4521 at 123 Main St → Completed"
  category: string; // "work_order" | "unit_turn" | "vendor" | etc.
  action: string; // "created" | "updated" | "completed" | "cancelled" etc.
  severity: string; // "info" | "success" | "warning" | "danger"
  icon: string; // Emoji icon for the toast
  resource_id: string | null;
  invalidates: string[]; // Cache keys to invalidate in the frontend
  raw: any;
} {
  const p = payload || {};
  const rType = (resourceType || p.resource_type || "").toLowerCase().replace(
    /\s+/g,
    "_",
  );

  // ── Normalise action ────────────────────────────────────────────────────────
  // AppFolio topics can come as "work_order.updated" or separate topic+event fields
  // Webhook topics refer to a specific event or category of events
  const rawTopic = String(p.topic || "").toLowerCase();
  const rawEvent = String(p.event || p.event_type || "").toLowerCase();

  let action = rawEvent;
  if (!action && rawTopic.includes(".")) {
    action = rawTopic.split(".").slice(1).join(".");
  }
  if (!action) action = "updated";

  // ── Pull useful data fields out of the payload ──────────────────────────────
  const data = p.data || p.payload || p.resource || {};
  const woNumber = data.work_order_number || data.number ||
    p.work_order_number || "";
  const address = data.property_address || data.address || p.property_address ||
    "";
  const status = data.status || p.status || "";
  const unitName = data.unit || p.unit || "";
  const tenantName = data.tenant_name || data.tenant || p.tenant_name || "";
  const vendorName = data.vendor_name || data.name || p.vendor_name || "";
  const propName = data.property_name || data.name || p.property_name || "";
  const category = data.category || p.category || "";
  const assignedTo = data.assigned_to || p.assigned_to || "";
  const priority = data.priority || p.priority || "";

  // ── Event map: resource_type + action → human readable ──────────────────────
  const EVENT_MAP: Record<
    string,
    Record<string, {
      title: string;
      icon: string;
      severity: string;
      desc: (d: any) => string;
    }>
  > = {
    work_order: {
      created: {
        title: "New Work Order",
        icon: "🔧",
        severity: "info",
        desc: (d) =>
          [
            woNumber ? `WO #${woNumber}` : null,
            address ? `at ${address}` : null,
            category ? `· ${category}` : null,
            priority ? `· ⚡ ${priority}` : null,
            assignedTo ? `→ Assigned to ${assignedTo}` : null,
          ].filter(Boolean).join(" "),
      },
      updated: {
        title: "Work Order Updated",
        icon: "📝",
        severity: "info",
        desc: (d) =>
          [
            woNumber ? `WO #${woNumber}` : null,
            address ? `at ${address}` : null,
            status ? `→ ${status}` : null,
          ].filter(Boolean).join(" "),
      },
      completed: {
        title: "Work Order Completed",
        icon: "✅",
        severity: "success",
        desc: (d) =>
          [
            woNumber ? `WO #${woNumber}` : null,
            address ? `at ${address}` : null,
          ].filter(Boolean).join(" "),
      },
      work_completed: {
        title: "Work Completed",
        icon: "✅",
        severity: "success",
        desc: (d) =>
          [
            woNumber ? `WO #${woNumber}` : null,
            address ? `at ${address}` : null,
          ].filter(Boolean).join(" "),
      },
      cancelled: {
        title: "Work Order Cancelled",
        icon: "🚫",
        severity: "warning",
        desc: (d) =>
          [
            woNumber ? `WO #${woNumber}` : null,
            address ? `at ${address}` : null,
          ].filter(Boolean).join(" "),
      },
      canceled: {
        title: "Work Order Cancelled",
        icon: "🚫",
        severity: "warning",
        desc: (d) =>
          [
            woNumber ? `WO #${woNumber}` : null,
            address ? `at ${address}` : null,
          ].filter(Boolean).join(" "),
      },
      assigned: {
        title: "Work Order Assigned",
        icon: "👷",
        severity: "info",
        desc: (d) =>
          [
            woNumber ? `WO #${woNumber}` : null,
            address ? `at ${address}` : null,
            assignedTo ? `→ Assigned to ${assignedTo}` : null,
          ].filter(Boolean).join(" "),
      },
      note_added: {
        title: "Note Added to Work Order",
        icon: "💬",
        severity: "info",
        desc: (d) =>
          [
            woNumber ? `WO #${woNumber}` : null,
            address ? `at ${address}` : null,
          ].filter(Boolean).join(" "),
      },
      status_changed: {
        title: "Work Order Status Changed",
        icon: "🔄",
        severity: "info",
        desc: (d) =>
          [
            woNumber ? `WO #${woNumber}` : null,
            status ? `→ ${status}` : null,
            address ? `at ${address}` : null,
          ].filter(Boolean).join(" "),
      },
    },

    unit_turn: {
      created: {
        title: "Unit Turn Started",
        icon: "🏠",
        severity: "info",
        desc: (d) =>
          [
            unitName ? `Unit ${unitName}` : null,
            propName ? `at ${propName}` : null,
          ].filter(Boolean).join(" "),
      },
      updated: {
        title: "Unit Turn Updated",
        icon: "🔄",
        severity: "info",
        desc: (d) =>
          [
            unitName ? `Unit ${unitName}` : null,
            propName ? `at ${propName}` : null,
            status ? `→ ${status}` : null,
          ].filter(Boolean).join(" "),
      },
      completed: {
        title: "Unit Turn Completed",
        icon: "🎉",
        severity: "success",
        desc: (d) =>
          [
            unitName ? `Unit ${unitName}` : null,
            propName ? `at ${propName}` : null,
          ].filter(Boolean).join(" "),
      },
    },

    tenant: {
      created: {
        title: "New Tenant Added",
        icon: "👤",
        severity: "info",
        desc: (d) =>
          [
            tenantName ? tenantName : null,
            propName ? `at ${propName}` : null,
            unitName ? `Unit ${unitName}` : null,
          ].filter(Boolean).join(" "),
      },
      updated: {
        title: "Tenant Updated",
        icon: "✏️",
        severity: "info",
        desc: (d) =>
          [
            tenantName ? tenantName : null,
            propName ? `at ${propName}` : null,
          ].filter(Boolean).join(" "),
      },
      move_in: {
        title: "Tenant Move-In",
        icon: "🔑",
        severity: "success",
        desc: (d) =>
          [
            tenantName ? tenantName : null,
            unitName ? `into Unit ${unitName}` : null,
            propName ? `at ${propName}` : null,
          ].filter(Boolean).join(" "),
      },
      move_out: {
        title: "Tenant Move-Out",
        icon: "📦",
        severity: "warning",
        desc: (d) =>
          [
            tenantName ? tenantName : null,
            unitName ? `from Unit ${unitName}` : null,
            propName ? `at ${propName}` : null,
          ].filter(Boolean).join(" "),
      },
      notice_given: {
        title: "Notice to Vacate",
        icon: "📋",
        severity: "warning",
        desc: (d) =>
          [
            tenantName ? tenantName : null,
            propName ? `at ${propName}` : null,
            unitName ? `Unit ${unitName}` : null,
          ].filter(Boolean).join(" "),
      },
    },

    lease: {
      created: {
        title: "New Lease Created",
        icon: "📄",
        severity: "info",
        desc: (d) =>
          [
            tenantName ? tenantName : null,
            propName ? `at ${propName}` : null,
          ].filter(Boolean).join(" "),
      },
      updated: {
        title: "Lease Updated",
        icon: "📄",
        severity: "info",
        desc: (d) =>
          [
            tenantName ? tenantName : null,
            propName ? `at ${propName}` : null,
          ].filter(Boolean).join(" "),
      },
      renewed: {
        title: "Lease Renewed",
        icon: "✅",
        severity: "success",
        desc: (d) =>
          [
            tenantName ? tenantName : null,
            propName ? `at ${propName}` : null,
          ].filter(Boolean).join(" "),
      },
    },

    vendor: {
      created: {
        title: "New Vendor Added",
        icon: "🏢",
        severity: "info",
        desc: (d) => vendorName ? vendorName : "New vendor added",
      },
      updated: {
        title: "Vendor Updated",
        icon: "✏️",
        severity: "info",
        desc: (d) => vendorName ? vendorName : "Vendor record updated",
      },
    },

    inspection: {
      created: {
        title: "Inspection Scheduled",
        icon: "🔍",
        severity: "info",
        desc: (d) =>
          [
            propName ? `at ${propName}` : null,
            unitName ? `Unit ${unitName}` : null,
          ].filter(Boolean).join(" "),
      },
      updated: {
        title: "Inspection Updated",
        icon: "🔍",
        severity: "info",
        desc: (d) =>
          [
            propName ? `at ${propName}` : null,
            status ? `→ ${status}` : null,
          ].filter(Boolean).join(" "),
      },
      completed: {
        title: "Inspection Completed",
        icon: "✅",
        severity: "success",
        desc: (d) =>
          [
            propName ? `at ${propName}` : null,
            unitName ? `Unit ${unitName}` : null,
          ].filter(Boolean).join(" "),
      },
    },

    property: {
      created: {
        title: "New Property Added",
        icon: "🏘️",
        severity: "info",
        desc: (d) => propName || "New property added",
      },
      updated: {
        title: "Property Updated",
        icon: "🏘️",
        severity: "info",
        desc: (d) => propName || "Property record updated",
      },
    },

    bill: {
      created: {
        title: "New Bill Created",
        icon: "💰",
        severity: "info",
        desc: (d) =>
          [
            vendorName ? `from ${vendorName}` : null,
            propName ? `at ${propName}` : null,
          ].filter(Boolean).join(" "),
      },
      updated: {
        title: "Bill Updated",
        icon: "💰",
        severity: "info",
        desc: (d) =>
          [
            vendorName ? `from ${vendorName}` : null,
            status ? `→ ${status}` : null,
          ].filter(Boolean).join(" "),
      },
      approved: {
        title: "Bill Approved",
        icon: "✅",
        severity: "success",
        desc: (d) =>
          [
            vendorName ? `from ${vendorName}` : null,
            propName ? `at ${propName}` : null,
          ].filter(Boolean).join(" "),
      },
    },
  };

  // ── Cache keys the frontend should invalidate on this event ─────────────────
  const INVALIDATION_MAP: Record<string, string[]> = {
    work_order: ["work_orders", "turn_work_orders", "recent_tasks", "labor"],
    unit_turn: ["turns", "turn_work_orders"],
    vendor: ["vendors"],
    inspection: ["inspections"],
    tenant: ["upcoming_moveouts"],
    lease: ["upcoming_moveouts"],
    property: ["properties", "property_groups", "property_map"],
    bill: ["bills"],
  };

  // ── Look up in map ───────────────────────────────────────────────────────────
  const typeMap = EVENT_MAP[rType];
  const eventDef = typeMap?.[action] || typeMap?.["updated"];

  if (eventDef) {
    return {
      title: eventDef.title,
      description: eventDef.desc(data) || `${rType} ${action}`,
      category: rType,
      action,
      severity: eventDef.severity,
      icon: eventDef.icon,
      resource_id: resourceId,
      invalidates: INVALIDATION_MAP[rType] || [],
      raw: p,
    };
  }

  // ── Fallback: unknown topic ──────────────────────────────────────────────────
  const fallbackTitle = [rType, action]
    .filter(Boolean)
    .map((s) =>
      s.replace(/_/g, " ").replace(/\b\w/g, (c: string) => c.toUpperCase())
    )
    .join(" — ");

  return {
    title: fallbackTitle || "AppFolio Event",
    description: eventType || "Webhook event received",
    category: rType || "unknown",
    action: action || "unknown",
    severity: "info",
    icon: "📡",
    resource_id: resourceId,
    invalidates: [],
    raw: p,
  };
}

function eventVerbFromEventType(eventType: string | null): string {
  const evt = String(eventType || "").toLowerCase();
  if (!evt) return "updated";
  if (evt.includes("schedule")) return "scheduled";
  if (evt.includes("assign")) return "reassigned";
  if (evt.includes("complete")) return "completed";
  if (evt.includes("create") || evt.includes("open")) return "created";
  if (evt.includes("cancel") || evt.includes("delete")) return "removed";
  if (evt.includes("status")) return "status updated";
  return "updated";
}

function composeWorkOrderDescriptionFromState(
  state: {
    id: string;
    wo_number: string;
    property_address: string;
    status_text: string;
  },
  eventType: string | null,
): string {
  const woNumber = String(state.wo_number || "").trim();
  const address = String(state.property_address || "").trim();
  const status = String(state.status_text || "").trim();
  const verb = eventVerbFromEventType(eventType);

  const head = woNumber
    ? `WO #${woNumber} ${verb}`
    : `Work order ${state.id.slice(0, 8)} ${verb}`;

  const atPart = address ? ` @ ${address}` : "";
  const statusPart = status ? ` -> ${status}` : "";
  return `${head}${atPart}${statusPart}`;
}

function composeVendorDescription(
  vendorName: string,
  eventType: string | null,
): string {
  const name = String(vendorName || "").trim() || "Vendor";
  const verb = eventVerbFromEventType(eventType);
  return `${name} ${verb}`;
}

function findVendorNameInCachedDirectory(
  cached: any,
  resourceId: string,
): string | null {
  const rows = Array.isArray(cached) ? cached : [];
  const target = String(resourceId || "").trim();
  if (!target) return null;

  for (const row of rows) {
    const rowId = String(
      row?.id || row?.Id || row?.vendor_id || row?.VendorId || "",
    ).trim();
    if (!rowId || rowId !== target) continue;

    const name = String(
      row?.name || row?.Name || row?.vendor_name || row?.VendorName || "",
    ).trim();
    if (name) return name;
  }

  return null;
}

async function resolveHumanDescriptionFromCache(
  resourceType: string,
  resourceId: string | null,
  eventType: string | null,
): Promise<string | null> {
  if (!resourceId) return null;

  if (resourceType === "work_order") {
    try {
      const result = await sqlite.execute({
        sql: `SELECT id, wo_number, property_address, status_text
              FROM wo_states
              WHERE id = ?
              LIMIT 1`,
        args: [resourceId],
      });
      const rows = rowsAsObjects(result);
      if (rows.length > 0) {
        return composeWorkOrderDescriptionFromState(rows[0] as any, eventType);
      }
    } catch (_) {
      // Continue to live resolve fallback.
    }

    try {
      const found = await resolveWebhookResource(
        "work_order",
        String(resourceId),
      );
      if (found.ok && found.record) {
        const state = extractWorkOrderState(found.record);
        if (state.id) {
          return composeWorkOrderDescriptionFromState(state, eventType);
        }
      }
    } catch (_) {
      // Continue to generic fallback.
    }
  }

  if (resourceType === "vendor") {
    // First try the cached vendor directory payload in api_cache.
    try {
      const cached = await sqlite.execute({
        sql: `SELECT data
              FROM api_cache
              WHERE cache_key = 'vendors'
              ORDER BY cached_at DESC
              LIMIT 1`,
      });
      const row = rowsAsObjects(cached)[0];
      if (row?.data) {
        const payload = JSON.parse(String(row.data || "[]"));
        const nameFromCache = findVendorNameInCachedDirectory(
          payload,
          String(resourceId || ""),
        );
        if (nameFromCache) {
          return composeVendorDescription(nameFromCache, eventType);
        }
      }
    } catch (_) {
      // Continue to live resolve fallback.
    }

    // Fallback to live AppFolio resolve for vendor UUIDs.
    try {
      const found = await resolveWebhookResource("vendor", String(resourceId));
      if (found.ok && found.record) {
        const name = String(
          found.record?.Name || found.record?.name || found.record?.Title || "",
        ).trim();
        if (name) {
          return composeVendorDescription(name, eventType);
        }
      }
    } catch (_) {
      return null;
    }
  }

  // Generic fallback for UUID-heavy events where payload fields are minimal.
  try {
    const found = await resolveWebhookResource(
      resourceType,
      String(resourceId),
    );
    if (found.ok && found.record) {
      const rec = found.record || {};
      const status = String(
        rec.Status || rec.status || rec.WorkOrderStatus ||
          rec.work_order_status ||
          rec.UnitTurnStatus || rec.unit_turn_status || rec.State ||
          rec.state || "",
      ).trim();
      const name = String(
        rec.Name || rec.name || rec.Title || rec.title || rec.PropertyName ||
          rec.property_name || rec.Unit || rec.unit || rec.Number ||
          rec.number ||
          rec.WorkOrderNumber || rec.work_order_number || "",
      ).trim();
      const verb = eventVerbFromEventType(eventType);
      const typeLabel = String(resourceType || "resource").replace(/_/g, " ");
      if (name || status) {
        return `${typeLabel} ${name || String(resourceId).slice(0, 8)} ${verb}${
          status ? ` -> ${status}` : ""
        }`;
      }
    }
  } catch (_) {
    return null;
  }

  return null;
}
// ── decodeBase64Url ───────────────────────────────────────────────────────────
// Converts a URL-safe base64 string (no padding) to a decoded UTF-8 string.
// Used to extract the payload segment of a compact JWS token.
// AppFolio signs all outgoing webhooks — the payload is the middle segment
// of the three-part dot-separated compact serialisation .
function decodeBase64Url(input: string): string {
  const b64 = input.replace(/-/g, "+").replace(/_/g, "/");
  const padded = b64 + "=".repeat((4 - (b64.length % 4)) % 4);
  return atob(padded);
}

// ── normalizeWebhookFields ────────────────────────────────────────────────────
// Extracts the canonical event_type, resource_type, resource_id, and event_id
// from a decoded webhook payload object.
// AppFolio webhooks may use either a flat structure (event, resource_type,
// resource_id) or a topic + event_type nested structure depending on the
// subscription topic .
function normalizeWebhookFields(payload: any): {
  payload: any;
  eventType: string | null;
  resourceType: string | null;
  resourceId: string | null;
  eventId: string | null;
} {
  const p = payload && typeof payload === "object" ? payload : {};
  const body = (p.data && typeof p.data === "object" && p.data) ||
    (p.payload && typeof p.payload === "object" && p.payload) ||
    (p.resource && typeof p.resource === "object" && p.resource) ||
    {};
  const nestedResource =
    (body.resource && typeof body.resource === "object" && body.resource) ||
    (body.entity && typeof body.entity === "object" && body.entity) ||
    {};
  const topic = String(p.topic || "").trim().toLowerCase();
  const evt = String(p.event_type || body.event_type || "").trim()
    .toLowerCase();
  const event = String(p.event || body.event || "").trim().toLowerCase();

  const eventType = event ||
    (topic && evt ? `${topic}.${evt}` : (evt || null));

  const resourceType = (p.resource_type && String(p.resource_type).trim()) ||
    (body.resource_type && String(body.resource_type).trim()) ||
    (body.type && String(body.type).trim()) ||
    (nestedResource.type && String(nestedResource.type).trim()) ||
    (topic && topic.split(".")[0]) ||
    (event && event.split(".")[0]) ||
    null;

  const resourceId = (p.resource_id && String(p.resource_id).trim()) ||
    (body.resource_id && String(body.resource_id).trim()) ||
    (body.ResourceId && String(body.ResourceId).trim()) ||
    (body.work_order_id && String(body.work_order_id).trim()) ||
    (body.vendor_id && String(body.vendor_id).trim()) ||
    (body.unit_turn_id && String(body.unit_turn_id).trim()) ||
    (nestedResource.id && String(nestedResource.id).trim()) ||
    (p.resource && p.resource.id && String(p.resource.id).trim()) ||
    (body.id && String(body.id).trim()) ||
    (p.id && String(p.id).trim()) ||
    null;

  const eventId = (p.event_id && String(p.event_id).trim()) ||
    (body.event_id && String(body.event_id).trim()) ||
    (p.idempotency_key && String(p.idempotency_key).trim()) ||
    null;

  return { payload: p, eventType, resourceType, resourceId, eventId };
}

// ── parseIncomingWebhook ──────────────────────────────────────────────────────
// Attempts JWS decode first (3-part dot-separated), falls back to plain JSON.
// Returns normalised fields regardless of parse mode.
// The raw body is always stored verbatim so the HMAC signature embedded in
// the JWS can be verified independently .
function parseIncomingWebhook(rawBody: string): {
  payload: any;
  eventType: string | null;
  resourceType: string | null;
  resourceId: string | null;
  eventId: string | null;
  parseMode: string;
} {
  const raw = String(rawBody || "").trim();

  if (!raw) {
    return {
      payload: {},
      eventType: null,
      resourceType: null,
      resourceId: null,
      eventId: null,
      parseMode: "empty",
    };
  }

  // ── Try JWS compact serialisation: header.payload.signature ──────────────
  // AppFolio signs all outgoing webhooks  using this format.
  try {
    const parts = raw.split(".");
    if (parts.length === 3) {
      const payloadJson = decodeBase64Url(parts[1]);
      const payload = JSON.parse(payloadJson);
      return { ...normalizeWebhookFields(payload), parseMode: "jws" };
    }
  } catch { /* fall through to plain JSON */ }

  // ── Try plain JSON ────────────────────────────────────────────────────────
  try {
    const payload = JSON.parse(raw);
    return { ...normalizeWebhookFields(payload), parseMode: "json" };
  } catch {
    return {
      payload: {},
      eventType: null,
      resourceType: null,
      resourceId: null,
      eventId: null,
      parseMode: "unknown",
    };
  }
}

// ── verifyDetachedJws ────────────────────────────────────────────────────────
// AppFolio sends detached JWS signatures via X-JWS-Signature using the form
// protected..signature (payload omitted). We reconstruct compact JWS by
// base64url-encoding the raw request body as the middle segment and verifying
// against AppFolio's JWKS endpoint.
async function verifyDetachedJws(
  req: Request,
  rawBody: string,
): Promise<{ ok: true; alg: string; kid: string | null }> {
  const signatureHeader = req.headers.get("x-jws-signature") ||
    req.headers.get("X-JWS-Signature") ||
    "";

  if (!signatureHeader) {
    throw new Error("missing_x_jws_signature");
  }

  const parts = signatureHeader.split(".");
  if (parts.length !== 3 || parts[1] !== "") {
    throw new Error("invalid_detached_jws_format");
  }

  const [encodedProtected, _emptyPayload, encodedSignature] = parts;
  if (!encodedProtected || !encodedSignature) {
    throw new Error("invalid_detached_jws_segments");
  }

  const encodedPayload = base64url.encode(new TextEncoder().encode(rawBody));
  const compactJws =
    `${encodedProtected}.${encodedPayload}.${encodedSignature}`;

  const header = decodeProtectedHeader(compactJws);
  const alg = String(header.alg || "");
  if (!alg || (alg !== "PS256" && alg !== "RSASSA_PSS_SHA_256")) {
    throw new Error("unexpected_jws_alg");
  }

  await compactVerify(compactJws, APPFOLIO_JWKS, {
    algorithms: ["PS256"],
  });

  return { ok: true, alg, kid: header.kid ? String(header.kid) : null };
}

// ============================================================================
// INGEST HANDLER
// ============================================================================

// ── handleWebhookPost ─────────────────────────────────────────────────────────
// Receives an HTTPS POST from AppFolio, stores the raw body in webhook_events,
// and invalidates relevant api_cache entries.
//
// AppFolio delivers webhooks via HTTPS POST to the registered URL .
// All outgoing webhooks are signed by AppFolio for authenticity .
// Use Send Test Event in AppFolio Admin Settings to verify connectivity .
// The Webhook Logs page shows delivery status and timestamps .
// Keep this handler lightweight so it can return 200 within a few seconds and
// avoid AppFolio retry storms.
//
// If req.method is POST with no ?action param, main.ts routes here directly.
// The legacy ?action=webhook POST is also routed here for v7 compat.
export async function handleWebhookPost(req: Request): Promise<any> {
  const rawBody = await req.text();

  // Enforce AppFolio detached JWS verification before any processing/storage.
  try {
    await verifyDetachedJws(req, rawBody);
  } catch (e: any) {
    const reason = (e && e.message)
      ? String(e.message)
      : "signature_verify_failed";
    const status = e instanceof joseErrors.JWSSignatureVerificationFailed
      ? 401
      : (reason === "missing_x_jws_signature" ||
          reason.startsWith("invalid_") || reason.startsWith("unexpected_"))
      ? 401
      : 401;

    return {
      ok: false,
      status,
      error: "Webhook verification failed",
      reason,
    };
  }

  const parsed = parseIncomingWebhook(rawBody);

  const { resourceType, resourceId, eventType } = parsed;
  const canonicalResourceType = normalizeResourceType(resourceType);
  const resolvedResourceType = canonicalResourceType || "unknown";
  const interpreted = interpretWebhookEvent(
    parsed.payload,
    eventType,
    resolvedResourceType,
    resourceId,
  );

  let humanDescription = String(
    interpreted.description || interpreted.title || "Webhook event received",
  );
  const cachedDescription = await resolveHumanDescriptionFromCache(
    resolvedResourceType,
    resourceId,
    eventType,
  );
  if (cachedDescription) {
    humanDescription = cachedDescription;
  }

  let fetchBack: any = {
    attempted: false,
    ok: false,
    reason: "not_applicable",
  };

  // Store raw body verbatim — preserves JWS signature for downstream verification
  let inserted: any;
  let storageBlocked = false;
  try {
    inserted = await sqlite.execute({
      sql: `INSERT INTO webhook_events
             (raw_body, resource_type, resource_id, event_type, human_description)
           VALUES (?, ?, ?, ?, ?)`,
      args: [
        rawBody,
        resolvedResourceType || null,
        resourceId,
        eventType,
        humanDescription,
      ],
    });
  } catch (e: any) {
    const msg = String(e?.message || "");
    if (isSqlWriteBlockedError(e)) {
      storageBlocked = true;
    } else if (/human_description|no such column|SQL_INPUT_ERROR/i.test(msg)) {
      try {
        inserted = await sqlite.execute({
          sql: `INSERT INTO webhook_events
                 (raw_body, resource_type, resource_id, event_type)
               VALUES (?, ?, ?, ?)`,
          args: [
            rawBody,
            resolvedResourceType || null,
            resourceId,
            eventType,
          ],
        });
      } catch (fallbackErr: any) {
        if (isSqlWriteBlockedError(fallbackErr)) {
          storageBlocked = true;
        } else {
          throw fallbackErr;
        }
      }
    } else {
      throw e;
    }
  }

  if (storageBlocked) {
    return {
      ok: true,
      status: 202,
      accepted: true,
      stored: false,
      warning: "database_write_blocked",
      hint: "Upgrade Val Town/Turso plan or reduce writes to restore durable webhook logging",
      event_type: eventType,
      resource_type: resolvedResourceType,
      resource_id: resourceId,
      parse_mode: parsed.parseMode,
      fetch_back: { attempted: false, ok: false, reason: "storage_blocked" },
    };
  }

  const insertedEventId = Number((inserted as any)?.lastInsertRowid || 0);

  // Process local routing matrix updates asynchronously so inbound webhook ACK
  // is not blocked by enrichment or downstream sync work.
  if (insertedEventId > 0) {
    Promise.resolve()
      .then(() =>
        processWebhookSync(
          sqlite,
          eventType || "",
          parsed.payload || {},
          insertedEventId,
        )
      )
      .catch((err: unknown) => {
        console.error(
          `[webhook] processWebhookSync failed for event ${insertedEventId}:`,
          err,
        );
      });
  }

  // Fetch-back for work_orders with a short delay to avoid eventual-consistency races.
  if (resolvedResourceType === "work_order" && resourceId) {
    fetchBack = { attempted: true, ok: false, reason: "resolve_failed" };
    try {
      await sleep(1500);
      const woResolved = await fetchWorkOrderById(String(resourceId));
      const latestNoteAt = await fetchLatestWorkOrderNoteAt(String(resourceId));
      if (woResolved.ok && woResolved.record) {
        const state = extractWorkOrderState(woResolved.record);
        state.last_note_at = latestNoteAt || "";
        if (state.id) {
          const previous = await getPreviousWorkOrderState(state.id);
          const changes: Array<
            {
              type: "reassigned" | "status_changed" | "note_added";
              oldValue: string;
              newValue: string;
            }
          > = [];

          if (previous) {
            const prevAssigned = String(previous.assigned_user_id || "");
            const nextAssigned = String(state.assigned_user_id || "");
            if (prevAssigned !== nextAssigned) {
              changes.push({
                type: "reassigned",
                oldValue: prevAssigned || "(unassigned)",
                newValue: nextAssigned || "(unassigned)",
              });
            }

            const prevStatus = previous.status_code === null
              ? ""
              : String(previous.status_code);
            const nextStatus = state.status_code === null
              ? ""
              : String(state.status_code);
            if (prevStatus !== nextStatus) {
              changes.push({
                type: "status_changed",
                oldValue: prevStatus || "(none)",
                newValue: nextStatus || "(none)",
              });
            }
          }

          // Record note_added change when event signals a note was added.
          // Also handles cases where last_note_at advances regardless of event type.
          const isNoteEvent = (eventType || "").toLowerCase().includes("note");
          if (isNoteEvent && latestNoteAt) {
            changes.push({
              type: "note_added",
              oldValue: "",
              newValue: latestNoteAt,
            });
          }

          for (const c of changes) {
            await insertWorkOrderStateChange(
              state.id,
              c.type,
              c.oldValue,
              c.newValue,
            );
          }

          await upsertWorkOrderState(state, eventType, resolvedResourceType);

          const enrichedDescription = composeWorkOrderDescriptionFromState(
            state,
            eventType,
          );
          if (insertedEventId > 0 && enrichedDescription) {
            try {
              await sqlite.execute({
                sql: `UPDATE webhook_events
                    SET human_description = ?
                    WHERE id = ?`,
                args: [enrichedDescription, insertedEventId],
              });
            } catch (e: any) {
              const msg = String(e?.message || "");
              if (
                !/human_description|no such column|SQL_INPUT_ERROR/i.test(msg)
              ) {
                throw e;
              }
            }
            humanDescription = enrichedDescription;
          }

          const ghostUpdateDropped = !!previous && changes.length === 0;
          fetchBack = {
            attempted: true,
            ok: true,
            id: state.id,
            status_code: state.status_code,
            assigned_user_id: state.assigned_user_id,
            last_activity_at: state.last_activity_at,
            last_note_at: state.last_note_at || null,
            changes_logged: changes.length,
            ghost_update: ghostUpdateDropped,
            human_description: humanDescription,
          };
          if (ghostUpdateDropped) {
            fetchBack.reason = "ghost_update_no_relevant_change";
          }
        } else {
          fetchBack = {
            attempted: true,
            ok: false,
            reason: "resolved_record_missing_id",
          };
        }
      } else {
        fetchBack = {
          attempted: true,
          ok: false,
          reason: woResolved.detail || "work_order_fetch_failed",
        };
      }
    } catch (e: any) {
      fetchBack = {
        attempted: true,
        ok: false,
        reason: String(e?.message || "fetch_back_error"),
      };
    }
  }

  // Invalidate api_cache entries for the affected resource type
  if (resolvedResourceType && WEBHOOK_CACHE_MAP[resolvedResourceType]) {
    try {
      for (const entityType of WEBHOOK_CACHE_MAP[resolvedResourceType]) {
        await cacheInvalidate(entityType);
      }
    } catch (e: any) {
      if (!isSqlWriteBlockedError(e)) throw e;
      console.warn(
        "[webhook] cache invalidation skipped: SQL writes blocked by plan/quota",
      );
    }
  }

  // Invalidate the short-TTL per-WO notes cache when a note event fires so
  // the frontend sees fresh notes on the next fetch rather than waiting 5 min.
  if (
    resolvedResourceType === "work_order" &&
    resourceId &&
    (eventType || "").toLowerCase().includes("note")
  ) {
    try {
      await sqlite.execute({
        sql: `DELETE FROM api_cache WHERE cache_key = ?`,
        args: [`wo_notes_${resourceId}`],
      });
    } catch (_) { /* non-fatal */ }
  }

  return {
    ok: true,
    event_type: eventType,
    resource_type: resolvedResourceType,
    resource_id: resourceId,
    parse_mode: parsed.parseMode,
    fetch_back: fetchBack,
  };
}

// ============================================================================
// QUERY ENDPOINTS
// ============================================================================

// ── handleWebhookEvents ───────────────────────────────────────────────────────
// Returns stored webhook events with optional filters and pagination.
// Each returned event includes the raw body, parsed fields, and a resolved
// title for display in the HandyManager webhook feed.
//
// Supported query params:
//   limit       number  default 200
//   offset      number  default 0
//   since_id    number  return only events with id > since_id (live polling)
//   search      string  substring match on event_type, resource_type, resource_id
//   type        string  substring match on event_type only
//   source      string  "appfolio" | "has_data" | "empty"
//   from        string  ISO timestamp — received_at >=
//   to          string  ISO timestamp — received_at <=
export async function handleWebhookEvents(
  params: Record<string, string>,
): Promise<any> {
  const limit = parseInt(params.limit || "200", 10);
  const offset = parseInt(params.offset || "0", 10);
  const sinceId = parseInt(params.since_id || "0", 10);
  const search = params.search || "";
  const typeFilter = params.type || "";
  const sourceFilter = params.source || "";
  const fromDate = params.from || "";
  const toDate = params.to || "";

  const conditions: string[] = [];
  const args: any[] = [];

  if (sinceId > 0) {
    conditions.push("id > ?");
    args.push(sinceId);
  }
  if (search) {
    conditions.push(
      "(event_type LIKE ? OR resource_type LIKE ? OR resource_id LIKE ?)",
    );
    const s = `%${search}%`;
    args.push(s, s, s);
  }
  if (typeFilter) {
    conditions.push("event_type LIKE ?");
    args.push(`%${typeFilter}%`);
  }
  if (sourceFilter === "appfolio" || sourceFilter === "has_data") {
    conditions.push("length(raw_body) > 2");
  } else if (sourceFilter === "empty") {
    conditions.push("length(raw_body) <= 2");
  }
  if (fromDate) {
    conditions.push("received_at >= ?");
    args.push(fromDate);
  }
  if (toDate) {
    conditions.push("received_at <= ?");
    args.push(toDate);
  }

  const whereClause = conditions.length > 0
    ? "WHERE " + conditions.join(" AND ")
    : "";

  let result: any;
  let total: any;
  try {
    result = await sqlite.execute({
      sql: `SELECT id,
                  received_at AS ts,
                  event_type  AS type,
                  resource_type,
                  resource_id,
                  human_description,
                  raw_body,
                  CASE WHEN length(raw_body) > 2
                    THEN 'has_data'
                    ELSE 'empty'
                  END AS body_status
           FROM webhook_events
           ${whereClause}
           ORDER BY id DESC
           LIMIT ? OFFSET ?`,
      args: [...args, limit, offset],
    });
  } catch (e: any) {
    const msg = String(e?.message || "");
    if (!/human_description|no such column|SQL_INPUT_ERROR/i.test(msg)) {
      return {
        ok: false,
        error: "db_error",
        message: `Could not read webhook_events: ${msg}`,
      };
    }
    try {
      result = await sqlite.execute({
        sql: `SELECT id,
                    received_at AS ts,
                    event_type  AS type,
                    resource_type,
                    resource_id,
                    '' AS human_description,
                    raw_body,
                    CASE WHEN length(raw_body) > 2
                      THEN 'has_data'
                      ELSE 'empty'
                    END AS body_status
             FROM webhook_events
             ${whereClause}
             ORDER BY id DESC
             LIMIT ? OFFSET ?`,
        args: [...args, limit, offset],
      });
    } catch (fallbackErr: any) {
      return {
        ok: false,
        error: "db_error",
        message: `Could not read webhook_events: ${
          String(fallbackErr?.message || fallbackErr)
        }`,
      };
    }
  }

  try {
    total = await sqlite.execute({
      sql: `SELECT COUNT(*) AS cnt FROM webhook_events ${whereClause}`,
      args: args,
    });
  } catch (countErr: any) {
    return {
      ok: false,
      error: "db_error",
      message: `Could not count webhook_events: ${
        String(countErr?.message || countErr)
      }`,
    };
  }

  const eventRows = (result && Array.isArray(result.rows))
    ? rowsAsObjects(result)
    : [];
  const totalRows = (total && Array.isArray(total.rows))
    ? rowsAsObjects(total)
    : [];

  return {
    ok: true,
    count: eventRows.length,
    total: totalRows[0]?.cnt || 0,
    events: eventRows.map((r) => {
      const parsed = parseIncomingWebhook(r.raw_body || "");
      const rawObj: any = parsed.payload || {};
      const src = r.body_status === "has_data" ? "appfolio" : "empty";
      const resolvedType = parsed.eventType || r.type || null;
      const resolvedResourceType = parsed.resourceType || r.resource_type ||
        null;
      const resolvedResourceId = parsed.resourceId || r.resource_id || null;
      const interpreted = interpretWebhookEvent(
        rawObj,
        resolvedType,
        resolvedResourceType,
        resolvedResourceId,
      );
      const title = interpreted.title || "Webhook Event";
      const description = String(
        r.human_description || interpreted.description || title,
      );

      return {
        id: r.id,
        ts: r.ts,
        type: resolvedType,
        event_type: rawObj.event_type || resolvedType || null,
        topic: rawObj.topic || null,
        event_id: parsed.eventId || rawObj.event_id || null,
        resource_type: resolvedResourceType,
        resource_id: resolvedResourceId,
        body_status: r.body_status,
        title,
        human_description: String(r.human_description || ""),
        description,
        category: interpreted.category,
        action: interpreted.action,
        severity: interpreted.severity,
        icon: interpreted.icon,
        body: r.raw_body || "",
        priority: "normal",
        source: src,
        raw: rawObj,
      };
    }),
  };
}

// ── handleWebhookStats ────────────────────────────────────────────────────────
// Returns aggregate counts and breakdowns for the HandyManager webhook panel.
// Includes totals by event type, source (appfolio vs empty), and by day
// for the configured retention window.
export async function handleWebhookStats(
  _params: Record<string, string>,
): Promise<any> {
  const [totalResult, byType, bySource, byDay] = await Promise.all([
    sqlite.execute(
      `SELECT COUNT(*) AS total,
              SUM(CASE WHEN processed     = 0    THEN 1 ELSE 0 END) AS pending_unprocessed,
              SUM(CASE WHEN processed = 0 AND length(raw_body) > 2 THEN 1 ELSE 0 END) AS pending,
              SUM(CASE WHEN processed = 0 AND length(raw_body) <= 2 THEN 1 ELSE 0 END) AS pending_empty,
              SUM(CASE WHEN length(raw_body) > 2 THEN 1 ELSE 0 END) AS has_data,
              SUM(CASE WHEN length(raw_body) <= 2 THEN 1 ELSE 0 END) AS empty_body
       FROM webhook_events`,
    ),
    sqlite.execute(
      `SELECT COALESCE(event_type, 'unknown') AS type,
              COUNT(*) AS count
       FROM webhook_events
       GROUP BY event_type
       ORDER BY count DESC`,
    ),
    sqlite.execute(
      `SELECT CASE WHEN length(raw_body) > 2
                THEN 'appfolio'
                ELSE 'empty'
              END AS source,
              COUNT(*) AS count
       FROM webhook_events
       GROUP BY source`,
    ),
    sqlite.execute({
      sql: `SELECT date(received_at) AS day,
                   COUNT(*) AS count
            FROM webhook_events
            WHERE received_at >= datetime('now', ?)
            GROUP BY day
            ORDER BY day DESC`,
      args: [`-${WEBHOOK_MAX_DAYS} days`],
    }),
  ]);

  const stats = rowsAsObjects(totalResult)[0] || {};

  return {
    ok: true,
    total: stats.total || 0,
    pending: stats.pending || 0,
    pending_unprocessed: stats.pending_unprocessed || 0,
    pending_empty: stats.pending_empty || 0,
    has_data: stats.has_data || 0,
    empty_body: stats.empty_body || 0,
    by_type: rowsAsObjects(byType),
    by_source: rowsAsObjects(bySource),
    by_day: rowsAsObjects(byDay),
  };
}

// ── handleWebhookFeed ────────────────────────────────────────────────────────
// Frontend polling endpoint: recent resolved work-order state snapshots.
// Query params:
//   limit   number default 50 (max 200)
//   since   ISO timestamp; returns rows updated at/after this value
export async function handleWebhookFeed(
  params: Record<string, string>,
): Promise<any> {
  const limit = Math.max(1, Math.min(parseInt(params.limit || "50", 10), 200));
  const since = String(params.since || "").trim();

  const where: string[] = [];
  const args: any[] = [];
  if (since) {
    where.push("updated_at >= ?");
    args.push(since);
  }
  const whereClause = where.length ? `WHERE ${where.join(" AND ")}` : "";

  const result = await sqlite.execute({
    sql: `SELECT id, status_code, status_text, assigned_user_id,
                 assigned_user_name, last_activity_at, last_note_at, event_type,
                 resource_type, wo_number, property_address,
                 raw_snapshot, fetched_at, updated_at
          FROM wo_states
          ${whereClause}
          ORDER BY updated_at DESC
          LIMIT ?`,
    args: [...args, limit],
  });

  const rows = rowsAsObjects(result);
  return {
    ok: true,
    count: rows.length,
    states: rows.map((r: any) => {
      let snapshot: any = {};
      try {
        snapshot = JSON.parse(String(r.raw_snapshot || "{}"));
      } catch {
        snapshot = {};
      }
      return {
        id: r.id,
        status_code: r.status_code,
        status_text: r.status_text,
        assigned_user_id: r.assigned_user_id,
        assigned_user_name: r.assigned_user_name,
        last_activity_at: r.last_activity_at,
        last_note_at: r.last_note_at,
        event_type: r.event_type,
        resource_type: r.resource_type,
        wo_number: r.wo_number,
        property_address: r.property_address,
        fetched_at: r.fetched_at,
        updated_at: r.updated_at,
        snapshot,
      };
    }),
  };
}

// ── handleWebhookResolve ──────────────────────────────────────────────────────
// Resolves a webhook resource_type + resource_id to a live AppFolio record.
// AppFolio webhook payloads contain opaque UUID resource IDs .
// Correct user roles must be enabled for certain API requests .
// If a user does not have the correct permissions, the request will
// result in a 422 "User not found" error .
//
// Query params:
//   resource_type  or  type   e.g. "work_order"
//   resource_id    or  id     AppFolio UUID string
export async function handleWebhookResolve(
  params: Record<string, string>,
): Promise<any> {
  const rawType = (params.resource_type || params.type || "").trim();
  const rawId = (params.resource_id || params.id || "").trim();

  if (!rawType || !rawId) {
    return { ok: false, error: "Missing resource_type/type or resource_id/id" };
  }

  // Normalise type: lowercase, strip plurals, strip special chars
  const normalizeType = (raw: string): string => {
    let t = String(raw || "").trim().toLowerCase();
    if (!t) return "";
    if (t.includes(".")) t = t.split(".")[0];
    t = t.replace(/[^a-z0-9_]/g, "");
    if (t.endsWith("ies")) t = t.slice(0, -3) + "y";
    else if (t.endsWith("s")) t = t.slice(0, -1);
    return t;
  };

  const singular = normalizeType(rawType);
  const found = await resolveWebhookResource(singular, rawId);

  if (!found.ok || !found.record) {
    return {
      ok: false,
      error: `resolve failed: HTTP ${found.status || 404}`,
      status: found.status || 404,
      domain: found.domain || "unknown",
      resource_type: singular,
      resource_id: rawId,
      detail: found.detail || "Not found via ID or filters[Id] lookup",
    };
  }

  // Build a human-readable summary from common field names
  const toSummary = (record: any) => ({
    title: record.Title || record.Name || record.Unit ||
      record.PropertyName || record.Description || record.Id || rawId,
    status: record.Status || record.WorkOrderStatus ||
      record.UnitTurnStatus || record.State || "",
    reference: record.Number || record.Reference ||
      record.WorkOrderNumber || record.Unit || "",
  });

  return {
    ok: true,
    resource_type: singular,
    resource_id: rawId,
    domain: found.domain,
    summary: toSummary(found.record),
    record: found.record,
  };
}

// ============================================================================
// MIGRATION / MAINTENANCE
// ============================================================================

// ── handleMigrateV8 ───────────────────────────────────────────────────────────
// Marks legacy empty webhook events (raw_body length <= 2, e.g. "{}" or "")
// as processed so they no longer appear as pending in cache_stats.
// Run once after upgrading from v7 to v8.
export async function handleMigrateV8(): Promise<any> {
  const result = await sqlite.execute(
    `UPDATE webhook_events
     SET    processed = 1, processed_at = datetime('now')
     WHERE  length(raw_body) <= 2
       AND  processed = 0`,
  );
  return {
    ok: true,
    message: `Marked ${
      (result as any).rowsAffected || 0
    } empty events as processed`,
  };
}

// ── handleWebhookMigrate ──────────────────────────────────────────────────────
// One-time blob → SQLite migration for v7 installs that stored webhook events
// in Val Town's blob storage rather than SQLite. Safe to run multiple times —
// the blob is cleared after a successful migration so subsequent runs are no-ops.
export async function handleWebhookMigrate(): Promise<any> {
  let migrated = 0;
  try {
    const stored = await blob.getJSON("webhook_events");
    if (Array.isArray(stored) && stored.length > 0) {
      for (const evt of stored) {
        await sqlite.execute({
          sql: `INSERT INTO webhook_events
                   (received_at, raw_body, resource_type, resource_id, event_type)
                 VALUES (?, ?, ?, ?, ?)`,
          args: [
            evt.ts || new Date().toISOString(),
            JSON.stringify(evt),
            evt.resource_type || null,
            evt.resource_id || null,
            evt.type || null,
          ],
        });
        migrated++;
      }
      // Clear blob after successful migration
      await blob.setJSON("webhook_events", []);
    }
  } catch { /* no blob data to migrate — silent no-op */ }

  return {
    ok: true,
    migrated,
    message: migrated > 0
      ? `Migrated ${migrated} events from blob storage to SQLite`
      : "No blob data to migrate",
  };
}

// ── handleWebhookCleanup ──────────────────────────────────────────────────────
// Manually trigger the webhook housekeeping routine (also runs on cold start).
// Trims events older than WEBHOOK_MAX_DAYS and caps total at WEBHOOK_MAX_EVENTS.
export async function handleWebhookCleanupEndpoint(): Promise<any> {
  const result = await webhookCleanup();
  return {
    ok: true,
    ...result,
    message:
      `Deleted ${result.deleted_old} old events, ${result.deleted_overflow} overflow events`,
  };
}

export async function handleAdminSyncRoute(
  action: string,
  params: Record<string, string>,
  req: Request,
): Promise<Response> {
  const authFailed = requireAdminAuth(req);
  if (authFailed) return authFailed;

  const groupId = String(params.group_id || "").trim() || undefined;

  try {
    switch (action) {
      case "admin_sync_groups": {
        const count = await syncPropertyGroups(sqlite);
        return adminJson({ ok: true, message: "Property groups synced", count });
      }
      case "admin_sync_properties": {
        const count = await syncPropertyMap(sqlite, groupId);
        return adminJson({
          ok: true,
          message: groupId
            ? `Properties synced for group ${groupId}`
            : "Properties synced",
          count,
          group_id: groupId || null,
        });
      }
      case "admin_sync_vendors": {
        const count = await syncVendorMap(sqlite);
        return adminJson({ ok: true, message: "Vendors synced", count });
      }
      case "admin_sync_work_orders": {
        const count = await syncWorkOrderMap(sqlite, groupId);
        return adminJson({
          ok: true,
          message: groupId
            ? `Work orders synced for group ${groupId}`
            : "Work orders synced",
          count,
          group_id: groupId || null,
        });
      }
      case "admin_sync_billing": {
        const count = await syncBillingMap(sqlite, groupId);
        return adminJson({
          ok: true,
          message: groupId
            ? `Billing synced for group ${groupId}`
            : "Billing synced",
          count,
          group_id: groupId || null,
        });
      }
      case "admin_rebuild_cache": {
        const count = await rebuildGroupResolutionCache(sqlite, groupId);
        return adminJson({
          ok: true,
          message: groupId
            ? `Resolution cache rebuilt for group ${groupId}`
            : "Resolution cache rebuilt",
          count,
          group_id: groupId || null,
        });
      }
      case "admin_invalidate_cache": {
        if (!groupId) {
          return adminJson(
            { ok: false, error: "Missing required query param: group_id" },
            400,
          );
        }
        await invalidateCacheForGroup(sqlite, groupId);
        return adminJson({
          ok: true,
          message: `Resolution cache invalidated for group ${groupId}`,
          group_id: groupId,
        });
      }
      case "admin_resolve_context": {
        const propertyMapId = String(params.property_map_id || "").trim();
        if (!propertyMapId) {
          return adminJson(
            { ok: false, error: "Missing required query param: property_map_id" },
            400,
          );
        }
        const context = await resolveRoutingContext(sqlite, propertyMapId);
        return adminJson({ ok: true, property_map_id: propertyMapId, context });
      }
      case "admin_resolve_groups": {
        const propertyMapId = String(params.property_map_id || "").trim();
        if (!propertyMapId) {
          return adminJson(
            { ok: false, error: "Missing required query param: property_map_id" },
            400,
          );
        }
        const groups = await resolveGroupsForProperty(sqlite, propertyMapId);
        return adminJson({ ok: true, property_map_id: propertyMapId, groups });
      }
      default:
        return adminJson({ ok: false, error: `Unknown admin action: ${action}` }, 404);
    }
  } catch (e: unknown) {
    if (isRateLimitedSyncError(e)) {
      return adminJson(
        { ok: false, error: "Rate limited by AppFolio. Retry shortly." },
        429,
        { "Retry-After": "60" },
      );
    }
    const message = e instanceof Error ? e.message : String(e || "Unknown error");
    return adminJson({ ok: false, error: message }, 500);
  }
}