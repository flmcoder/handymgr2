// ============================================================================
// handlers/reassignment.ts — v9.0 Automated Reassignment Engine.
//
// This is the core intelligence layer of HandyManager. It runs on a
// scheduled cadence via two Val Town Cron vals:
//
//   Pass A — Noon Warning (12:00 PM AZ / 19:00 UTC)
//     Warns technicians 12 hours before their work order will be pulled.
//     Sends a RingCentral SMS with a magic link to contact the tenant.
//     Writes a system note to the AppFolio work order .
//
//   Pass B — Midnight Reassign (12:00 AM AZ / 07:00 UTC)
//     Pulls stale work orders (no activity in 48+ hours) and reassigns them
//     to the next available Tier 1 tech by load-weighted selection.
//     If no Tier 1 tech is available, fires a Tier 2 simultaneous SMS blast.
//     Writes a system note to the AppFolio work order on every action.
//     Notifies both the outgoing and incoming technician via RingCentral SMS.
//
// APPFOLIO WRITE RULES enforced by this engine:
//   • All work order PATCHes use AssignedUsers: [{ Id: "uuid" }].
//     The UUID must belong to a user with the Maintenance Tech role in
//     AppFolio Property Manager — any other role returns 422 .
//   • PATCHes are issued sequentially with 200 ms between each work order.
//     Concurrent PATCHes to the same work order will cause the second to
//     fail. Staggering requests prevents network congestion and errors .
//   • The system note Body field key is "Body" (capital B) — confirmed by AF docs.
//   • All list endpoints that need filtering use filters[LastUpdatedAtFrom]
//     or the request will return 400 Bad Request.
//
// :stop-auto: command:
//   Any note on a work order containing ":stop-auto:" (case-insensitive)
//   permanently exempts that WO from both Pass A and Pass B until the note
//   is removed and the reassignment_queue row is cleared by an admin.
//   A confirmation note is written to AppFolio and the admin is emailed .
//
// Grace period:
//   If the midnight cron detects note activity within the 48-hour window
//   AND the grace_used flag is 0, one grace extension is granted. The WO
//   is skipped for that cycle and grace_used is set to 1. No second grace.
// ============================================================================

import { cacheInvalidate, rowsAsObjects, sqlite } from "../db.ts";
import {
  ADMIN_NOTIFY_EMAIL,
  AF_DB,
  AF_REPORTS,
  CRON_SECRET,
  dbHeaders,
} from "../config.ts";
import { patchWorkOrder, postWoNote } from "../lib/appfolio.ts";
import { sendBulkSMS, sendSMS } from "../lib/ringcentral.ts";
import { generateMagicLink, isDispatchShortLink } from "../lib/auth.ts";
import { auditLog } from "../lib/audit.ts";
import { delay, fetchWithTimeout } from "../lib/fetchUtils.ts";
import { handleWoNotes } from "./workOrders.ts";
import { getTenantContact } from "./properties.ts";

const STALE_STATE_MS = 4 * 3600 * 1000;
const MAX_STALE_REFRESH_PER_RUN = 25;
const STALE_REFRESH_DELAY_MS = 500;
const TECH_ROLE_CACHE_TTL_MS = 15 * 60 * 1000;

type TechRoleCheckResult = {
  ok: boolean;
  role: string;
  reason: string;
};

const techRoleCheckCache = new Map<string, {
  checkedAt: number;
  result: TechRoleCheckResult;
}>();

function toMs(ts: string | null | undefined): number {
  const ms = new Date(String(ts || "")).getTime();
  return isNaN(ms) ? 0 : ms;
}

function normalizeStatusCode(record: any): number | null {
  const raw = record?.StatusCode ?? record?.status_code ??
    record?.StatusId ?? record?.status_id ?? null;
  if (raw === null || raw === undefined || raw === "") return null;
  const n = Number(raw);
  return isNaN(n) ? null : n;
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
  const candidates = [
    payload?.meta?.total_pages,
    payload?.metadata?.total_pages,
    payload?.pagination?.total_pages,
    payload?.pages,
    payload?.total_pages,
  ];
  for (const c of candidates) {
    const n = Number(c);
    if (!isNaN(n) && n > 0) return Math.floor(n);
  }
  return 1;
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
  return (payload?.data && !Array.isArray(payload.data) ? payload.data : null) ||
    (payload?.data && Array.isArray(payload.data) ? payload.data[0] : null) ||
    payload?.results?.[0] ||
    (Array.isArray(payload) ? payload[0] : null) ||
    null;
}

function normalizeRoleText(rawRole: unknown): string {
  return String(rawRole || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function isMaintenanceTechRole(rawRole: unknown): boolean {
  const role = normalizeRoleText(rawRole);
  if (!role) return false;
  return role === "maintenance tech" ||
    role === "maintenance technician" ||
    (role.includes("maintenance") && role.includes("tech"));
}

function extractUserRole(record: any): string {
  return String(
    record?.UserRole ?? record?.user_role ??
      record?.Role ?? record?.role ??
      record?.Title ?? record?.title ??
      "",
  ).trim();
}

async function fetchUserById(userId: string): Promise<{
  ok: boolean;
  record?: any;
  detail?: string;
}> {
  const encodedId = encodeURIComponent(userId);
  const byId = await fetchDbPath(`/api/v0/users/${encodedId}`);
  if (byId.ok) {
    const rec = firstRecord(byId.payload);
    if (rec) return { ok: true, record: rec };
  }

  const byFilter = await fetchDbPath(
    `/api/v0/users?filters[Id]=${encodedId}&page[size]=1`,
  );
  if (!byFilter.ok) {
    return { ok: false, detail: byFilter.detail || "user_fetch_failed" };
  }

  const rec = firstRecord(byFilter.payload);
  if (!rec) return { ok: false, detail: "user_not_found" };
  return { ok: true, record: rec };
}

async function validateMaintenanceTechRole(
  userId: string,
): Promise<TechRoleCheckResult> {
  const key = String(userId || "").trim();
  if (!key) {
    return { ok: false, role: "", reason: "missing_user_id" };
  }

  const now = Date.now();
  const cached = techRoleCheckCache.get(key);
  if (cached && (now - cached.checkedAt) < TECH_ROLE_CACHE_TTL_MS) {
    return cached.result;
  }

  const fetched = await fetchUserById(key);
  if (!fetched.ok || !fetched.record) {
    const result = {
      ok: false,
      role: "",
      reason: fetched.detail || "user_lookup_failed",
    };
    techRoleCheckCache.set(key, { checkedAt: now, result });
    return result;
  }

  const role = extractUserRole(fetched.record);
  const valid = isMaintenanceTechRole(role);
  const result = {
    ok: valid,
    role,
    reason: valid ? "ok" : "role_not_maintenance_tech",
  };
  techRoleCheckCache.set(key, { checkedAt: now, result });
  return result;
}

async function fetchWorkOrderById(woId: string): Promise<{
  ok: boolean;
  record?: any;
  detail?: string;
}> {
  const byId = await fetchDbPath(`/api/v0/work_orders/${encodeURIComponent(woId)}`);
  if (byId.ok) {
    const rec = firstRecord(byId.payload);
    if (rec) return { ok: true, record: rec };
  }

  const byFilter = await fetchDbPath(
    `/api/v0/work_orders?filters[Id]=${encodeURIComponent(woId)}&page[size]=1`,
  );
  if (!byFilter.ok) return { ok: false, detail: byFilter.detail || "wo_fetch_failed" };
  const rec = firstRecord(byFilter.payload);
  if (!rec) return { ok: false, detail: "wo_not_found" };
  return { ok: true, record: rec };
}

async function fetchLatestWorkOrderNoteAt(woId: string): Promise<string> {
  const encoded = encodeURIComponent(woId);

  const sorted = await fetchDbPath(
    `/api/v0/work_orders/${encoded}/notes?sort=-CreatedAt&page[size]=1`,
  );
  if (sorted.ok) {
    const notes = Array.isArray(sorted.payload?.data)
      ? sorted.payload.data
      : (Array.isArray(sorted.payload?.results)
        ? sorted.payload.results
        : (Array.isArray(sorted.payload) ? sorted.payload : []));
    const newest = extractLatestNoteTimestamp(notes);
    if (newest) return newest;
  }

  const basePath = `/api/v0/work_orders/${encoded}/notes?page[size]=50`;
  const first = await fetchDbPath(basePath);
  if (!first.ok) return "";

  const firstNotes = Array.isArray(first.payload?.data)
    ? first.payload.data
    : (Array.isArray(first.payload?.results)
      ? first.payload.results
      : (Array.isArray(first.payload) ? first.payload : []));

  const totalPages = extractTotalPages(first.payload);
  if (totalPages <= 1) return extractLatestNoteTimestamp(firstNotes);

  const last = await fetchDbPath(`${basePath}&page[number]=${totalPages}`);
  if (!last.ok) return extractLatestNoteTimestamp(firstNotes);
  const lastNotes = Array.isArray(last.payload?.data)
    ? last.payload.data
    : (Array.isArray(last.payload?.results)
      ? last.payload.results
      : (Array.isArray(last.payload) ? last.payload : []));

  return extractLatestNoteTimestamp(lastNotes) || extractLatestNoteTimestamp(firstNotes);
}

async function getAssignedWoStates(): Promise<any[]> {
  const result = await sqlite.execute({
    sql: `SELECT id, status_code, status_text, assigned_user_id, assigned_user_name,
                 last_activity_at, last_note_at, fetched_at, raw_snapshot
          FROM wo_states
          WHERE status_code = 9
          ORDER BY fetched_at ASC`,
  });
  return rowsAsObjects(result);
}

function parseWoSnapshot(row: any): any {
  try {
    return JSON.parse(String(row?.raw_snapshot || "{}"));
  } catch {
    return {};
  }
}

function woMetaFromStateRow(row: any): {
  woId: string;
  woNumber: string;
  techName: string;
  techId: string;
  address: string;
  unitRef: string;
  status: string;
  priority: string;
  category: string;
} {
  const snap = parseWoSnapshot(row);
  const assignedUsers = Array.isArray(snap?.AssignedUsers)
    ? snap.AssignedUsers
    : (Array.isArray(snap?.assigned_users) ? snap.assigned_users : []);
  const firstAssigned = assignedUsers[0] || {};

  const woId = String(row?.id || snap?.Id || snap?.id || "");
  const woNumber = String(
    snap?.WorkOrderNumber || snap?.work_order_number || snap?.Number ||
      snap?.number || woId,
  );
  const techName = String(
    row?.assigned_user_name || snap?.AssignedUserName || snap?.assigned_user_name ||
      firstAssigned?.Name || firstAssigned?.name || snap?.assigned_to ||
      snap?.AssignedTo || "Unknown Tech",
  );
  const techId = String(
    row?.assigned_user_id || snap?.AssignedUserId || snap?.assigned_user_id ||
      firstAssigned?.Id || firstAssigned?.id || snap?.assigned_to_id ||
      snap?.AssignedToId || techName,
  );
  const address = String(
    snap?.PropertyAddress || snap?.property_address || snap?.unit_address ||
      snap?.property || "",
  );
  const unitRef = String(
    snap?.UnitName || snap?.unit_name || snap?.unit || snap?.Unit || "",
  );
  const status = String(row?.status_text || snap?.Status || snap?.work_order_status || "");
  const priority = String(snap?.Priority || snap?.priority || "Normal");
  const category = String(snap?.Category || snap?.category || "");

  return {
    woId,
    woNumber,
    techName,
    techId,
    address,
    unitRef,
    status,
    priority,
    category,
  };
}

function isStateStale(fetchedAt: string, nowMs: number): boolean {
  const fetchedMs = toMs(fetchedAt);
  if (!fetchedMs) return true;
  return (nowMs - fetchedMs) > STALE_STATE_MS;
}

function isPastThresholdBoth(
  lastActivityAt: string,
  lastNoteAt: string,
  nowMs: number,
  thresholdMs: number,
): boolean {
  const actMs = toMs(lastActivityAt);
  const noteMs = toMs(lastNoteAt);
  const activityOld = !!actMs && (nowMs - actMs) >= thresholdMs;
  const noteOld = !noteMs || (nowMs - noteMs) >= thresholdMs;
  return activityOld && noteOld;
}

async function refreshWoStateFromApi(woId: string): Promise<{
  ok: boolean;
  row?: any;
  reason?: string;
}> {
  const [woRes, latestNoteAt] = await Promise.all([
    fetchWorkOrderById(woId),
    fetchLatestWorkOrderNoteAt(woId),
  ]);
  if (!woRes.ok || !woRes.record) {
    return { ok: false, reason: woRes.detail || "wo_refresh_failed" };
  }

  const rec = woRes.record;
  const statusCode = normalizeStatusCode(rec);
  const statusText = String(rec?.Status || rec?.status || rec?.WorkOrderStatus || rec?.work_order_status || "");
  const assignedUsers = Array.isArray(rec?.AssignedUsers)
    ? rec.AssignedUsers
    : (Array.isArray(rec?.assigned_users) ? rec.assigned_users : []);
  const firstAssigned = assignedUsers[0] || {};
  const assignedUserId = String(
    rec?.AssignedUserId || rec?.assigned_user_id || firstAssigned?.Id ||
      firstAssigned?.id || "",
  );
  const assignedUserName = String(
    rec?.AssignedUserName || rec?.assigned_user_name || firstAssigned?.Name ||
      firstAssigned?.name || "",
  );
  const lastActivityAt = String(
    rec?.LastUpdatedAt || rec?.last_updated_at || rec?.UpdatedAt ||
      rec?.updated_at || rec?.CreatedAt || rec?.created_at ||
      new Date().toISOString(),
  );
  const woNumber = String(rec?.WorkOrderNumber || rec?.work_order_number || rec?.Number || rec?.number || woId);
  const propertyAddress = String(rec?.PropertyAddress || rec?.property_address || rec?.Address || rec?.address || "");

  await sqlite.execute({
    sql: `INSERT INTO wo_states (
            id, status_code, status_text, assigned_user_id, assigned_user_name,
            last_activity_at, last_note_at, event_type, resource_type, wo_number,
            property_address, raw_snapshot, fetched_at, updated_at
          ) VALUES (?, ?, ?, ?, ?, ?, ?, 'manual_refresh', 'work_order', ?, ?, ?, datetime('now'), datetime('now'))
          ON CONFLICT(id) DO UPDATE SET
            status_code = excluded.status_code,
            status_text = excluded.status_text,
            assigned_user_id = excluded.assigned_user_id,
            assigned_user_name = excluded.assigned_user_name,
            last_activity_at = excluded.last_activity_at,
            last_note_at = excluded.last_note_at,
            event_type = 'manual_refresh',
            resource_type = 'work_order',
            wo_number = excluded.wo_number,
            property_address = excluded.property_address,
            raw_snapshot = excluded.raw_snapshot,
            fetched_at = datetime('now'),
            updated_at = datetime('now')`,
    args: [
      String(rec?.Id || rec?.id || woId),
      statusCode,
      statusText,
      assignedUserId,
      assignedUserName,
      lastActivityAt,
      latestNoteAt || "",
      woNumber,
      propertyAddress,
      JSON.stringify(rec || {}),
    ],
  });

  const row = rowsAsObjects(await sqlite.execute({
    sql: `SELECT id, status_code, status_text, assigned_user_id, assigned_user_name,
                 last_activity_at, last_note_at, fetched_at, raw_snapshot
          FROM wo_states WHERE id = ? LIMIT 1`,
    args: [String(rec?.Id || rec?.id || woId)],
  }))[0];

  if (!row) return { ok: false, reason: "refresh_upsert_missing" };
  return { ok: true, row };
}

// ── Shared open-status list ───────────────────────────────────────────────────
// Work orders in any of these statuses are candidates for warning/reassignment.
// Terminal statuses (Completed, Work Completed, Canceled) are excluded.
const OPEN_STATUSES = [
  "assigned",
  "scheduled",
  "waiting",
  "in progress",
  "new",
  "estimated",
  "estimate requested",
];

function isOpenStatus(status: string): boolean {
  const s = status.toLowerCase();
  return OPEN_STATUSES.some((o) => s.includes(o));
}

async function getProxyConfigValue(key: string): Promise<string> {
  const row = rowsAsObjects(await sqlite.execute({
    sql: `SELECT value FROM proxy_config WHERE key = ? LIMIT 1`,
    args: [key],
  }))[0];
  return String(row?.value || "");
}

async function getDispatchPaused(): Promise<boolean> {
  const value = await getProxyConfigValue("dispatch_paused");
  return value === "1" || value.toLowerCase() === "true";
}

async function getHiddenAssigneeIds(): Promise<string[]> {
  const raw = await getProxyConfigValue("dispatch_hidden_assignees");
  try {
    const map = JSON.parse(raw || "{}");
    if (!map || typeof map !== "object") return [];
    return Object.keys(map).filter((k) => !!map[k]);
  } catch {
    return [];
  }
}

// ============================================================================
// STOP COMMAND — :stop-auto:
// ============================================================================

// ── scanForStopCommand ────────────────────────────────────────────────────────
// Fetches the notes for a single work order and searches every note body
// for the ":stop-auto:" command string (case-insensitive).
//
// Returns { exempt: true, noted_by, noted_at, note_id } on match.
// Returns { exempt: false } on no match or any fetch error.
//
// This is the FIRST check run in both Pass A and Pass B — if exempt is true,
// the work order is skipped immediately and applyStopExemption() is called
// once to record the exemption, write an AF note, and email the admin.
//
// The Send Test Event webhook feature can be used to verify connectivity
// before relying on note data in production .
export async function scanForStopCommand(
  woId: string,
): Promise<{
  exempt: boolean;
  noted_by?: string;
  noted_at?: string;
  note_id?: string;
}> {
  try {
    const nr = await handleWoNotes({ wo_id: woId });
    if (!nr.ok || !Array.isArray(nr.results)) return { exempt: false };

    for (const note of nr.results) {
      const body = String(
        note.Body ||
          note.body ||
          note.Content ||
          note.content || "",
      ).toLowerCase().trim();

      if (body.includes(":stop-auto:")) {
        return {
          exempt: true,
          noted_by: note.CreatedByName || note.AuthorName || note.CreatedBy ||
            "Unknown",
          noted_at: note.CreatedAt || note.created_at ||
            new Date().toISOString(),
          note_id: String(note.Id || note.id || ""),
        };
      }
    }
    return { exempt: false };
  } catch {
    // Safe default — never let a scan error block the reassignment engine
    return { exempt: false };
  }
}

// ── applyStopExemption ────────────────────────────────────────────────────────
// Idempotent: only executes the full exemption sequence once per WO.
// Subsequent calls for the same wo_id are silently skipped.
//
// Sequence:
//   1. Upsert reassignment_queue with auto_exempt=1
//   2. POST confirmation note to AppFolio work order 
//   3. Append to wo_audit_log
//   4. Email admin via Val Town std/email (non-fatal)
export async function applyStopExemption(
  woId: string,
  woNumber: string,
  address: string,
  techName: string,
  status: string,
  noted_by: string,
  noted_at: string,
  note_id: string,
): Promise<void> {
  // Idempotency check
  const check = rowsAsObjects(
    await sqlite.execute({
      sql: `SELECT auto_exempt FROM reassignment_queue WHERE wo_id = ?`,
      args: [woId],
    }),
  )[0];
  if (check?.auto_exempt === 1) return;

  // 1. Upsert queue record
  await sqlite.execute({
    sql: `INSERT INTO reassignment_queue
             (wo_id, wo_number, property_address, assigned_tech_name, wo_status,
              auto_exempt, auto_exempt_at, auto_exempt_by, auto_exempt_note_id)
           VALUES (?, ?, ?, ?, ?, 1, ?, ?, ?)
           ON CONFLICT(wo_id) DO UPDATE SET
             auto_exempt         = 1,
             auto_exempt_at      = excluded.auto_exempt_at,
             auto_exempt_by      = excluded.auto_exempt_by,
             auto_exempt_note_id = excluded.auto_exempt_note_id`,
    args: [
      woId,
      woNumber,
      address,
      techName,
      status,
      noted_at,
      noted_by,
      note_id,
    ],
  });

  // 2. AppFolio system note — Body key is "Body" (capital B)
  // The Webhook Logs page will confirm note delivery 
  await postWoNote(
    woId,
    `[SYSTEM — HandyManager Auto-Exempt Activated | ${
      new Date().toISOString()
    }]\n` +
      `:stop-auto: command detected in notes by ${noted_by}.\n` +
      `Automated reassignment and warning system is DISABLED for this work order.\n` +
      `To re-enable: remove the :stop-auto: note and clear the queue entry via the Dispatch Control panel.`,
  );

  // 3. Audit log
  await auditLog(woId, "auto_exempt_activated", {
    noted_by,
    noted_at,
    note_id,
    address,
    status,
    tech: techName,
  });

  // 4. Admin email (non-fatal — email failure must never block the cron)
  if (ADMIN_NOTIFY_EMAIL) {
    try {
      const { email } = await import("https://esm.town/v/std/email");
      await email({
        to: ADMIN_NOTIFY_EMAIL,
        subject: `🔕 Auto-Reassign Disabled — WO #${woNumber} at ${address}`,
        text: `:stop-auto: detected on WO #${woNumber}\n\n` +
          `Address:   ${address}\n` +
          `Noted by:  ${noted_by}\n` +
          `Noted at:  ${noted_at}\n` +
          `Status:    ${status}\n` +
          `Tech:      ${techName}\n\n` +
          `This work order is now EXEMPT from automated reassignment.\n` +
          `To re-enable: remove the :stop-auto: note and clear the exemption via the Dispatch Control panel.`,
      });
    } catch { /* non-fatal */ }
  }
}

// ============================================================================
// TECH SELECTION & GRADING
// ============================================================================

// ── getNextAvailableTech ──────────────────────────────────────────────────────
// Selects the Tier 1 tech with the lowest weighted work load.
// Load score = active_wo_count × load_weight
// Excludes the current assignee (excludeTechId) to prevent re-assigning
// to the same person who just got the WO pulled.
//
// In order to successfully carry out certain requests using the API, the
// correct user roles must be enabled in AppFolio Property Manager .
// Verify every tech in tech_grades has the Maintenance Tech role in AppFolio
// before enabling the midnight cron in production.
export async function getNextAvailableTech(
  excludeTechId: string,
): Promise<{ tech_id: string; tech_name: string; tech_phone: string } | null> {
  const hidden = await getHiddenAssigneeIds();
  const hiddenClause = hidden.length
    ? ` AND tech_id NOT IN (${hidden.map(() => "?").join(",")})`
    : "";
  const candidates = rowsAsObjects(
    await sqlite.execute({
      sql: `SELECT tech_id, tech_name, tech_phone
             FROM   tech_grades
             WHERE  tier   = 1
               AND  active = 1
               AND  tech_id  != ?
               ${hiddenClause}
               AND  tech_phone IS NOT NULL
               AND  tech_phone != ''
             ORDER BY (active_wo_count * COALESCE(load_weight, 1.0)) ASC
             LIMIT 25`,
      args: [excludeTechId, ...hidden],
    }),
  );

  for (const row of candidates) {
    const roleCheck = await validateMaintenanceTechRole(String(row.tech_id || ""));
    if (!roleCheck.ok && roleCheck.reason === "role_not_maintenance_tech") continue;
    return {
      tech_id: row.tech_id,
      tech_name: row.tech_name,
      tech_phone: row.tech_phone,
    };
  }

  return null;
}

// ── updateTechGrades ──────────────────────────────────────────────────────────
// Called immediately after a successful auto-reassignment.
// Old tech: decrement active_wo_count, increment total_auto_reassigned,
//           increase load_weight (max 2.0) so they receive fewer future WOs.
// New tech: increment active_wo_count and total_assigned.
export async function updateTechGrades(
  oldTechId: string,
  oldTechName: string,
  newTechId: string,
  newTechName: string,
): Promise<void> {
  // Penalise old tech
  await sqlite.execute({
    sql: `INSERT INTO tech_grades
             (tech_id, tech_name, active_wo_count, total_auto_reassigned, load_weight, updated_at)
           VALUES (?, ?, 0, 1, 1.1, datetime('now'))
           ON CONFLICT(tech_id) DO UPDATE SET
             active_wo_count       = MAX(0, active_wo_count - 1),
             total_auto_reassigned = total_auto_reassigned + 1,
             load_weight           = MIN(2.0, load_weight + 0.1),
             updated_at            = datetime('now')`,
    args: [oldTechId, oldTechName],
  });

  // Reward new tech
  await sqlite.execute({
    sql: `INSERT INTO tech_grades
             (tech_id, tech_name, active_wo_count, total_assigned, updated_at)
           VALUES (?, ?, 1, 1, datetime('now'))
           ON CONFLICT(tech_id) DO UPDATE SET
             active_wo_count  = active_wo_count + 1,
             total_assigned   = total_assigned + 1,
             last_assigned_at = datetime('now'),
             updated_at       = datetime('now')`,
    args: [newTechId, newTechName],
  });
}

// ── recalculateTechGrades ─────────────────────────────────────────────────────
// Full grade recalculation run at the END of every midnight cron cycle.
// Computes performance_score and load_weight for all active Tier 1 techs,
// then calculates target_share_pct proportionally from relative scores.
//
// Scoring formula (4 weighted components):
//   Speed score    (40%) — based on avg_completion_hours vs 48-hr target
//   Go-back score  (30%) — forgives ≤2%, penalises heavily above that
//   Reassign score (20%) — rewards clean record, penalises repeated pulls
//   Response rate  (10%) — ratio of on-time warning responses
//
// load_weight range: 0.5 (best performer) → 2.0 (worst performer)
// A higher load_weight means the tech receives proportionally fewer new WOs.
export async function recalculateTechGrades(): Promise<void> {
  const techs = rowsAsObjects(
    await sqlite.execute(
      `SELECT * FROM tech_grades WHERE active = 1 AND tier = 1`,
    ),
  );

  for (const tech of techs) {
    const total = Number(tech.total_assigned) || 1; // prevent div/0
    const completed = Number(tech.total_completed) || 0;
    const goBacks = Number(tech.total_go_backs) || 0;
    const reassigned = Number(tech.total_auto_reassigned) || 0;
    const warned = Number(tech.total_warnings_recv) || 1;
    const onTime = Number(tech.total_on_time_resp) || 0;
    const avgHours = Number(tech.avg_completion_hours) || 0;

    const goBackPct = (goBacks / Math.max(completed, 1)) * 100;
    const reassignPct = (reassigned / total) * 100;
    const responseRt = (onTime / warned) * 100;

    // ── Speed score ───────────────────────────────────────────────────────────
    let speedScore: number;
    if (avgHours <= 48) speedScore = 100;
    else if (avgHours <= 72) speedScore = 75;
    else if (avgHours <= 96) speedScore = 50;
    else speedScore = Math.max(0, 100 - ((avgHours - 48) / 48) * 50);

    // ── Go-back score — forgives ≤2%, exponential penalty above ──────────────
    let goBackScore: number;
    if (goBackPct <= 2) goBackScore = 100;
    else if (goBackPct <= 5) goBackScore = 70;
    else if (goBackPct <= 10) goBackScore = 40;
    else goBackScore = Math.max(0, 40 - (goBackPct - 10) * 4);

    // ── Reassign score ────────────────────────────────────────────────────────
    let reassignScore: number;
    if (reassignPct === 0) reassignScore = 100;
    else if (reassignPct <= 5) reassignScore = 85;
    else if (reassignPct <= 15) reassignScore = 60;
    else reassignScore = Math.max(0, 60 - (reassignPct - 15) * 3);

    // ── Composite score ───────────────────────────────────────────────────────
    const score = (speedScore * 0.40) +
      (goBackScore * 0.30) +
      (reassignScore * 0.20) +
      (responseRt * 0.10);

    // load_weight inversely proportional to score
    // Score 100 → 1.0 (no penalty) · Score 0 → 2.0 (doubled load cost)
    const loadWeight = Math.max(0.5, Math.min(2.0, 1 + ((100 - score) / 100)));

    await sqlite.execute({
      sql: `UPDATE tech_grades
             SET  go_back_pct        = ?,
                  reassign_pct       = ?,
                  response_rate_pct  = ?,
                  performance_score  = ?,
                  load_weight        = ?,
                  score_updated_at   = datetime('now'),
                  updated_at         = datetime('now')
             WHERE tech_id = ?`,
      args: [
        goBackPct,
        reassignPct,
        responseRt,
        score,
        loadWeight,
        tech.tech_id,
      ],
    });
  }

  // Recalculate target_share_pct proportionally from relative scores
  const updated = rowsAsObjects(
    await sqlite.execute(
      `SELECT tech_id, performance_score FROM tech_grades WHERE active = 1 AND tier = 1`,
    ),
  );
  const totalScore = updated.reduce(
    (s: number, t: any) => s + Number(t.performance_score),
    0,
  );
  for (const t of updated) {
    const share = totalScore > 0
      ? (Number(t.performance_score) / totalScore) * 100
      : 33.3;
    await sqlite.execute({
      sql: `UPDATE tech_grades SET target_share_pct = ? WHERE tech_id = ?`,
      args: [share, t.tech_id],
    });
  }
}

// ============================================================================
// TIER 2 BLAST
// ============================================================================

// ── executeTier2Blast ─────────────────────────────────────────────────────────
// Fires a simultaneous SMS to all active Tier 2 techs when no Tier 1 tech
// is available to receive a reassigned work order.
//
// Flow:
//   1. Query tech_grades for all Tier 2 techs ordered by active_wo_count ASC
//   2. Create a blast_events record with a 24-hour claim window
//   3. Send SMS to each Tier 2 tech via RingCentral (300 ms between each)
//      Staggering requests prevents network congestion and errors 
//   4. Insert a tier2_claims row per tech to track reply status
//   5. Write audit log entry
//
// First Tier 2 tech to reply "Y" claims the job.
// The inbound SMS webhook handler (future scope) resolves claims.
export async function executeTier2Blast(
  woId: string,
  woNumber: string,
  address: string,
  category: string,
  priority: string,
  fromTech: string,
): Promise<{ ok: boolean; blast_id?: number; techs_notified?: number }> {
  const hidden = await getHiddenAssigneeIds();
  const hiddenClause = hidden.length
    ? ` AND tech_id NOT IN (${hidden.map(() => "?").join(",")})`
    : "";
  const tier2 = rowsAsObjects(
    await sqlite.execute({
      sql: `SELECT * FROM tech_grades
       WHERE  tier       = 2
         AND  active     = 1
         ${hiddenClause}
         AND  tech_phone IS NOT NULL
         AND  tech_phone != ''
       ORDER BY active_wo_count ASC`,
      args: hidden,
    }),
  );
  if (tier2.length === 0) return { ok: false };

  const expiresAt = new Date(Date.now() + 24 * 3600 * 1000).toISOString();

  // Create blast event
  const blastRes = await sqlite.execute({
    sql: `INSERT INTO blast_events
             (wo_id, wo_number, property_addr, category, priority, expires_at)
           VALUES (?, ?, ?, ?, ?, ?)
           RETURNING id`,
    args: [woId, woNumber, address, category, priority, expiresAt],
  });
  const blastId = rowsAsObjects(blastRes)[0]?.id;
  if (!blastId) return { ok: false };

  const blastMessage = `🔧 FLR DISPATCH — Job Available!\n` +
    `📍 ${address}\n` +
    `🔧 ${category} | ⚡ ${priority}\n\n` +
    `Previously with ${fromTech}.\n` +
    `First to reply Y claims this job.\n` +
    `⏳ Offer expires in 24 hours.`;

  let notified = 0;
  for (const tech of tier2) {
    const smsRes = await sendSMS(tech.tech_phone, blastMessage);

    await sqlite.execute({
      sql: `INSERT INTO tier2_claims
               (blast_id, wo_id, tech_id, tech_name, tech_phone, sms_sent_at)
             VALUES (?, ?, ?, ?, ?, datetime('now'))`,
      args: [blastId, woId, tech.tech_id, tech.tech_name, tech.tech_phone],
    });

    if (smsRes.ok) notified++;
    await delay(300); // Stagger RC SMS calls 
  }

  await auditLog(woId, "tier2_blast_sent", {
    blast_id: blastId,
    techs_notified: notified,
    address,
    category,
    priority,
    from_tech: fromTech,
  });

  return { ok: true, blast_id: blastId, techs_notified: notified };
}

// ============================================================================
// PASS A — NOON WARNING CRON (12:00 PM AZ / 19:00 UTC)
// ============================================================================

// ── handleNoonWarningCron ─────────────────────────────────────────────────────
// Triggered by Cron Val at 0 19 * * * (UTC).
// Finds all open work orders that have been in an open status for ≥ 36 hours
// without any note activity, and sends a warning SMS to the assigned tech.
//
// For each candidate work order:
//   0. :stop-auto: gate — if exempt, record and skip
//   1. Skip if already warned or exempt in reassignment_queue
//   2. Check for recent notes (informational — impacts warning message text)
//   3. Look up tech phone from tech_grades roster
//   4. Look up tenant phone via getTenantContact()
//   5. Generate 24-hr magic link for tech → tenant SMS portal
//   6. Send warning SMS via RingCentral (non-fatal if RC not configured)
//   7. POST system note to AppFolio work order 
//   8. Upsert reassignment_queue with warning_sent = 1
//   9. Append to wo_audit_log
//
// Staggering the rate at which requests are issued will prevent network
// congestion and errors . Each WO is processed with a 200 ms delay.
export async function handleNoonWarningCron(
  params: Record<string, string>,
): Promise<any> {
  // Cron secret guard
  if (CRON_SECRET && (params.secret || "") !== CRON_SECRET) {
    return { ok: false, error: "Unauthorized" };
  }

  if (await getDispatchPaused()) {
    return {
      ok: true,
      paused: true,
      run: "noon_warning_cron",
      candidates: 0,
      warned: 0,
      skipped: 0,
    };
  }

  const WARN_THRESHOLD_MS = 36 * 3600 * 1000; // 36 hours
  const warned: string[] = [];
  const skipped: { wo_id: string; reason: string }[] = [];
  let staleRefreshed = 0;

  const now = Date.now();
  const assignedRows = await getAssignedWoStates();
  let candidateCount = 0;

  for (const baseRow of assignedRows) {
    const woId = String(baseRow.id || "");
    if (!woId) continue;

    let row = baseRow;
    if (isStateStale(String(row.fetched_at || ""), now)) {
      if (staleRefreshed >= MAX_STALE_REFRESH_PER_RUN) {
        skipped.push({ wo_id: woId, reason: "stale state refresh cap reached" });
        continue;
      }
      const refreshed = await refreshWoStateFromApi(woId);
      staleRefreshed++;
      await delay(STALE_REFRESH_DELAY_MS);
      if (!refreshed.ok || !refreshed.row) {
        skipped.push({ wo_id: woId, reason: `stale refresh failed: ${refreshed.reason || "unknown"}` });
        continue;
      }
      row = refreshed.row;
    }

    const warnViolated = isPastThresholdBoth(
      String(row.last_activity_at || ""),
      String(row.last_note_at || ""),
      now,
      WARN_THRESHOLD_MS,
    );
    if (!warnViolated) continue;

    candidateCount++;

    const meta = woMetaFromStateRow(row);
    const woNumber = meta.woNumber;
    const techName = meta.techName;
    const techId = meta.techId;
    const address = meta.address;
    const unitRef = meta.unitRef;
    const status = meta.status;
    const priority = meta.priority;
    const category = meta.category;

    // ── 0. :stop-auto: gate ───────────────────────────────────────────────────
    const stopCheck = await scanForStopCommand(woId);
    if (stopCheck.exempt) {
      await applyStopExemption(
        woId,
        woNumber,
        address,
        techName,
        status,
        stopCheck.noted_by!,
        stopCheck.noted_at!,
        stopCheck.note_id!,
      );
      skipped.push({ wo_id: woId, reason: ":stop-auto: exempt" });
      await delay(150);
      continue;
    }

    // ── 1. Skip if already warned or exempt ───────────────────────────────────
    const qRow = rowsAsObjects(
      await sqlite.execute({
        sql:
          `SELECT warning_sent, auto_exempt FROM reassignment_queue WHERE wo_id = ?`,
        args: [woId],
      }),
    )[0];
    if (qRow && (qRow.warning_sent === 1 || qRow.auto_exempt === 1)) {
      skipped.push({ wo_id: woId, reason: "already warned or exempt" });
      continue;
    }

    // ── 2. Check for recent note activity ─────────────────────────────────────
    const hasRecent = false;

    // ── 3. Look up tech phone from roster ─────────────────────────────────────
    const techRow = rowsAsObjects(
      await sqlite.execute({
        sql: `SELECT tech_phone FROM tech_grades WHERE tech_id = ?`,
        args: [techId],
      }),
    )[0];
    const techPhone = techRow?.tech_phone || "";

    // ── 4. Look up tenant contact ─────────────────────────────────────────────
    const tenant = await getTenantContact(address, unitRef);

    // ── 5. Generate magic link ────────────────────────────────────────────────
    const magicLink = await generateMagicLink(
      woId,
      techId,
      techName,
      tenant.phone,
      tenant.name,
      address,
    );
    if (!isDispatchShortLink(magicLink)) {
      skipped.push({ wo_id: woId, reason: "dispatch short link unavailable" });
      console.log(`No dispatch short link available for warning WO ${woNumber}`);
      continue;
    }

    // ── 6. Build and send warning SMS ─────────────────────────────────────────
    const graceNote = hasRecent
      ? "\n✅ Note on file — 1 grace extension active this cycle."
      : "";

    const warnMsg = `⚠️ FLR DISPATCH WARNING — WO #${woNumber}\n` +
      `📍 ${address}\n` +
      `🔧 ${category} | ⚡ ${priority}` +
      `${graceNote}\n\n` +
      `No updates detected in 36 hours.\n` +
      `You have ~12 hours to take one of these actions:\n\n` +
      `  ✅ Add a note in AppFolio\n` +
      `  ✅ Change status to Scheduled\n` +
      `  ✅ Log any status update\n\n` +
      `Failure to act = AUTO-REASSIGNMENT at midnight.\n\n` +
      `📲 Contact tenant now:\n${magicLink}`;

    let smsSent = false;
    if (techPhone) {
      const smsRes = await sendSMS(techPhone, warnMsg);
      smsSent = smsRes.ok;
      if (!smsRes.ok) {
        console.log(`Warning SMS failed for WO ${woNumber}: ${smsRes.error}`);
      }
    }

    // ── 7. AppFolio system note ───────────────────────────────────────────────
    // The Send Test Event feature confirms note delivery 
    await postWoNote(
      woId,
      `[SYSTEM — HandyManager Pre-Reassignment Warning | ${
        new Date().toISOString()
      }]\n` +
        `12-hour warning issued to ${techName}.\n` +
        `No note or status activity detected in 36+ hours.\n` +
        `Work order will be AUTO-REASSIGNED at midnight if no qualifying update is recorded.\n` +
        `Current status: ${status}.` +
        (hasRecent
          ? "\nNote: Recent note activity detected — grace period may apply."
          : ""),
    );

    // ── 8. Upsert reassignment_queue ──────────────────────────────────────────
    await sqlite.execute({
      sql: `INSERT INTO reassignment_queue
               (wo_id, wo_number, property_address, assigned_tech_id, assigned_tech_name,
                wo_status, wo_priority, wo_category, warning_sent, warning_sent_at, warning_channel)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1, datetime('now'), 'sms')
             ON CONFLICT(wo_id) DO UPDATE SET
               warning_sent       = 1,
               warning_sent_at    = datetime('now'),
               warning_channel    = 'sms',
               assigned_tech_id   = excluded.assigned_tech_id,
               assigned_tech_name = excluded.assigned_tech_name,
               wo_status          = excluded.wo_status`,
      args: [
        woId,
        woNumber,
        address,
        techId,
        techName,
        status,
        priority,
        category,
      ],
    });

    // ── 9. Audit log ──────────────────────────────────────────────────────────
    await auditLog(woId, "reassignment_warning_sent", {
      tech: techName,
      phone_found: !!techPhone,
      sms_sent: smsSent,
      has_recent_note: hasRecent,
      magic_link: magicLink,
      tenant_name: tenant.name || "",
      tenant_phone: tenant.phone || "",
      address,
      status,
      priority,
    });

    warned.push(woId);
    await delay(200); // Stagger per-WO processing 
  }

  return {
    ok: true,
    run: "noon_warning",
    timestamp: new Date().toISOString(),
    candidates: candidateCount,
    warned: warned.length,
    skipped: skipped.length,
    details: { warned, skipped, stale_refreshed: staleRefreshed },
  };
}

// ============================================================================
// PASS B — MIDNIGHT REASSIGN CRON (12:00 AM AZ / 07:00 UTC)
// ============================================================================

// ── handleMidnightReassignCron ────────────────────────────────────────────────
// Triggered by Cron Val at 0 7 * * * (UTC).
// Finds all open work orders that have been in an open status for ≥ 48 hours
// without activity, and reassigns them via AppFolio API PATCH.
//
// For each candidate work order:
//   0. Re-scan for :stop-auto: (may have been added since noon warning)
//   1. Check grace_used + note activity gate (one-time grace extension)
//   2. Select next available Tier 1 tech via load-weighted query
//   3. If no Tier 1 available → fire Tier 2 blast, skip direct reassign
//   4. PATCH AppFolio WO AssignedUsers → new tech UUID 
//   5. POST system note to AppFolio WO documenting the reassignment 
//   6. SMS old tech: WO pulled notice
//   7. Generate new magic link for incoming tech
//   8. SMS new tech: assignment details + magic link
//   9. Update reassignment_queue (count++, new tech, reset warning flags)
//  10. Call updateTechGrades() for both techs
//  11. If reassignment_count >= 2: flag escalated, email admin
//  12. Audit log
//
// After all WOs processed: run recalculateTechGrades() and invalidate cache.
//
// CRITICAL: AssignedUsers must reference a user with the Maintenance Tech
// role in AppFolio Property Manager, or the PATCH returns 422 .
// All WOs are processed sequentially with 200 ms between each to prevent
// network congestion and simultaneous PATCH failures .
export async function handleMidnightReassignCron(
  params: Record<string, string>,
): Promise<any> {
  // Cron secret guard
  if (CRON_SECRET && (params.secret || "") !== CRON_SECRET) {
    return { ok: false, error: "Unauthorized" };
  }

  if (await getDispatchPaused()) {
    return {
      ok: true,
      paused: true,
      run: "midnight_reassign_cron",
      candidates: 0,
      reassigned: 0,
      skipped: 0,
      escalated: 0,
    };
  }

  const REASSIGN_THRESHOLD_MS = 48 * 3600 * 1000; // 48 hours
  const reassigned: string[] = [];
  const skipped: { wo_id: string; reason: string }[] = [];
  const escalated: string[] = [];
  let staleRefreshed = 0;

  const now = Date.now();
  const assignedRows = await getAssignedWoStates();
  let candidateCount = 0;

  for (const baseRow of assignedRows) {
    const woId = String(baseRow.id || "");
    if (!woId) continue;

    let row = baseRow;
    if (isStateStale(String(row.fetched_at || ""), now)) {
      if (staleRefreshed >= MAX_STALE_REFRESH_PER_RUN) {
        skipped.push({ wo_id: woId, reason: "stale state refresh cap reached" });
        continue;
      }
      const refreshed = await refreshWoStateFromApi(woId);
      staleRefreshed++;
      await delay(STALE_REFRESH_DELAY_MS);
      if (!refreshed.ok || !refreshed.row) {
        skipped.push({ wo_id: woId, reason: `stale refresh failed: ${refreshed.reason || "unknown"}` });
        continue;
      }
      row = refreshed.row;
    }

    const meta = woMetaFromStateRow(row);
    const woNumber = meta.woNumber;
    const techName = meta.techName;
    const techId = meta.techId;
    const address = meta.address;
    const unitRef = meta.unitRef;
    const status = meta.status;
    const priority = meta.priority;
    const category = meta.category;

    // ── 0. :stop-auto: gate (re-scan — command may have been added since noon) ─
    const stopCheck = await scanForStopCommand(woId);
    if (stopCheck.exempt) {
      await applyStopExemption(
        woId,
        woNumber,
        address,
        techName,
        status,
        stopCheck.noted_by!,
        stopCheck.noted_at!,
        stopCheck.note_id!,
      );
      skipped.push({ wo_id: woId, reason: ":stop-auto: exempt" });
      await delay(150);
      continue;
    }

    // ── 1. Fetch queue state ──────────────────────────────────────────────────
    const qRow = rowsAsObjects(
      await sqlite.execute({
        sql: `SELECT grace_used, reassignment_count, auto_exempt
               FROM   reassignment_queue
               WHERE  wo_id = ?`,
        args: [woId],
      }),
    )[0] || {};

    const graceUsed = Number(qRow.grace_used || 0);
    const reassignCount = Number(qRow.reassignment_count || 0);

    const graceViolated = isPastThresholdBoth(
      String(row.last_activity_at || ""),
      String(row.last_note_at || ""),
      now,
      REASSIGN_THRESHOLD_MS,
    );

    // ── 1a. One-time note grace gate ──────────────────────────────────────────
    if (!graceViolated) {
      const noteMs = toMs(String(row.last_note_at || ""));
      const hasRecent = !!noteMs && (now - noteMs) < REASSIGN_THRESHOLD_MS;

      if (!graceUsed && hasRecent) {
        await sqlite.execute({
          sql: `INSERT INTO reassignment_queue
                   (wo_id, wo_number, assigned_tech_name, grace_used, warning_sent)
                 VALUES (?, ?, ?, 1, 0)
                 ON CONFLICT(wo_id) DO UPDATE SET
                   grace_used      = 1,
                   warning_sent    = 0,
                   warning_sent_at = NULL`,
          args: [woId, woNumber, techName],
        });

        // System note confirming grace grant 
        await postWoNote(
          woId,
          `[SYSTEM — HandyManager Grace Period Granted | ${
            new Date().toISOString()
          }]\n` +
            `Note activity detected within 48 hours. Auto-reassignment clock reset for this cycle.\n` +
            `⚠️ Grace periods are ONE-TIME ONLY. Continued inactivity will trigger reassignment in the next cycle.`,
        );

        await auditLog(woId, "grace_period_granted", {
          tech: techName,
          address,
          status,
        });
        skipped.push({ wo_id: woId, reason: "one-time grace period granted" });
        await delay(150);
        continue;
      }

      skipped.push({ wo_id: woId, reason: "activity still within 48h window" });
      continue;
    }

    candidateCount++;

    // ── 2. Select next available Tier 1 tech ──────────────────────────────────
    const newTech = await getNextAvailableTech(techId);

    // ── 3. No Tier 1 available → Tier 2 blast ────────────────────────────────
    if (!newTech) {
      const blastRes = await executeTier2Blast(
        woId,
        woNumber,
        address,
        category,
        priority,
        techName,
      );
      escalated.push(woId);

      await auditLog(woId, "escalation_tier2_blast", {
        from_tech: techName,
        blast_ok: blastRes.ok,
        techs_notified: blastRes.techs_notified,
        address,
      });

      if (!blastRes.ok && ADMIN_NOTIFY_EMAIL) {
        try {
          const { email } = await import("https://esm.town/v/std/email");
          await email({
            to: ADMIN_NOTIFY_EMAIL,
            subject:
              `🚨 WO #${woNumber} — Cannot Auto-Reassign (No Techs Available)`,
            text:
              `WO #${woNumber} at ${address} requires reassignment but no Tier 1 OR Tier 2 techs\n` +
              `are available in the tech_grades roster.\n\n` +
              `Current tech:  ${techName}\n` +
              `Status:        ${status}\n\n` +
              `Add techs via POST ?action=tech_roster in HandyManager.`,
          });
        } catch { /* non-fatal */ }
      }

      skipped.push({
        wo_id: woId,
        reason: "no Tier 1 available — Tier 2 blast fired",
      });
      await delay(150);
      continue;
    }

    // ── 4. PATCH AppFolio WO ──────────────────────────────────────────────────
    // AssignedUsers is the confirmed AF DB API v0 field for work order assignment.
    // The tech_id UUID MUST belong to a user with the Maintenance Tech role
    // in AppFolio Property Manager — otherwise 422 is returned .
    // ⚠️ If 422 is received in logs, verify the UUID via:
    //   GET ?action=passthrough&path=/api/v0/users
    //   and confirm UserRole === "Maintenance Tech" for this user.
    const roleCheck = await validateMaintenanceTechRole(newTech.tech_id);
    if (!roleCheck.ok) {
      await auditLog(woId, "reassignment_blocked_invalid_assignee_role", {
        candidate_tech_id: newTech.tech_id,
        candidate_tech_name: newTech.tech_name,
        candidate_role: roleCheck.role || "unknown",
        reason: roleCheck.reason,
        from_tech: techName,
        address,
        status,
      });

      skipped.push({
        wo_id: woId,
        reason: `candidate assignee role invalid (${roleCheck.role || "unknown"})`,
      });
      await delay(150);
      continue;
    }

    const patchRes = await patchWorkOrder(woId, {
      AssignedUsers: [{ Id: newTech.tech_id }],
    });

    const newCount = reassignCount + 1;

    // ── 5. AppFolio system note ───────────────────────────────────────────────
    // Writes a timestamped note to the work order itself so the reassignment
    // is permanently visible in AppFolio. Webhook Logs confirm delivery .
    await postWoNote(
      woId,
      `[SYSTEM — HandyManager Auto-Reassignment #${newCount} | ${
        new Date().toISOString()
      }]\n` +
        `Work order reassigned: ${techName} → ${newTech.tech_name}.\n` +
        `Reason: No status updates or note activity detected in 48+ hours.\n` +
        `AppFolio PATCH: ${
          patchRes.ok
            ? "✅ Success"
            : `⚠️ Failed (HTTP ${patchRes.status}) — manual update required in AppFolio. Error: ${
              patchRes.error?.substring(0, 100)
            }`
        }.`,
    );

    // ── 6. SMS old tech: pulled notice ────────────────────────────────────────
    const oldTechPhone = rowsAsObjects(
      await sqlite.execute({
        sql: `SELECT tech_phone FROM tech_grades WHERE tech_id = ?`,
        args: [techId],
      }),
    )[0]?.tech_phone || "";

    if (oldTechPhone) {
      await sendSMS(
        oldTechPhone,
        `📋 FLR DISPATCH — WO #${woNumber}\n` +
          `📍 ${address}\n\n` +
          `This work order has been auto-reassigned due to 48-hour inactivity policy.\n` +
          `No further action required from you.`,
      );
    }

    // ── 7. Generate new magic link for incoming tech ──────────────────────────
    const tenant = await getTenantContact(address, unitRef);
    const newMagicLink = await generateMagicLink(
      woId,
      newTech.tech_id,
      newTech.tech_name,
      tenant.phone,
      tenant.name,
      address,
    );
    const newTechMessage = !isDispatchShortLink(newMagicLink)
      ? `🔧 FLR DISPATCH — WO #${woNumber} Assigned to You\n` +
        `📍 ${address}\n` +
        `🔧 ${category} | ⚡ ${priority}\n\n` +
        `Previously assigned to ${techName} (auto-reassigned due to inactivity).\n\n` +
        `Dispatch short link is not available yet. Open the Dispatch dashboard to regenerate the tenant contact link.`
      : `🔧 FLR DISPATCH — WO #${woNumber} Assigned to You\n` +
        `📍 ${address}\n` +
        `🔧 ${category} | ⚡ ${priority}\n\n` +
        `Previously assigned to ${techName} (auto-reassigned due to inactivity).\n\n` +
        `📲 Contact tenant:\n${newMagicLink}`;

    // ── 8. SMS new tech: assignment + magic link ──────────────────────────────
    await sendSMS(
      newTech.tech_phone,
      newTechMessage,
    );

    // ── 9. Update reassignment_queue ──────────────────────────────────────────
    await sqlite.execute({
      sql: `INSERT INTO reassignment_queue
               (wo_id, wo_number, property_address, assigned_tech_id, assigned_tech_name,
                wo_status, reassignment_count, last_reassigned_at,
                warning_sent, warning_sent_at, grace_used)
             VALUES (?, ?, ?, ?, ?, ?, 1, datetime('now'), 0, NULL, 0)
             ON CONFLICT(wo_id) DO UPDATE SET
               assigned_tech_id   = excluded.assigned_tech_id,
               assigned_tech_name = excluded.assigned_tech_name,
               wo_status          = excluded.wo_status,
               reassignment_count = reassignment_count + 1,
               last_reassigned_at = datetime('now'),
               warning_sent       = 0,
               warning_sent_at    = NULL,
               grace_used         = 0`,
      args: [
        woId,
        woNumber,
        address,
        newTech.tech_id,
        newTech.tech_name,
        status,
      ],
    });

    // ── 10. Update tech grades ────────────────────────────────────────────────
    await updateTechGrades(
      techId,
      techName,
      newTech.tech_id,
      newTech.tech_name,
    );

    // ── 11. Escalate if repeatedly reassigned ─────────────────────────────────
    if (newCount >= 2) {
      escalated.push(woId);

      await sqlite.execute({
        sql: `UPDATE reassignment_queue SET escalated = 1 WHERE wo_id = ?`,
        args: [woId],
      });

      if (ADMIN_NOTIFY_EMAIL) {
        try {
          const { email } = await import("https://esm.town/v/std/email");
          await email({
            to: ADMIN_NOTIFY_EMAIL,
            subject:
              `🚨 WO #${woNumber} — Reassigned ${newCount}x (Needs Attention)`,
            text:
              `WO #${woNumber} at ${address} has been auto-reassigned ${newCount} times.\n\n` +
              `Now assigned to:  ${newTech.tech_name}\n` +
              `Status:           ${status}\n` +
              `Priority:         ${priority}\n` +
              `Category:         ${category}\n\n` +
              `This work order requires manual admin attention.\n` +
              `View the full audit trail via HandyManager: ?action=reassignment_queue`,
          });
        } catch { /* non-fatal */ }
      }
    }

    // ── 12. Audit log ─────────────────────────────────────────────────────────
    await auditLog(woId, "auto_reassigned", {
      from: techName,
      to: newTech.tech_name,
      reassignment_count: newCount,
      af_patch_ok: patchRes.ok,
      af_patch_status: patchRes.status,
      address,
      status,
      priority,
    });

    reassigned.push(woId);
    await delay(200); // Sequential processing — prevents concurrent PATCH failures 
  }

  // ── Post-cycle: recalculate grades + invalidate WO cache ──────────────────
  if (reassigned.length > 0) {
    await recalculateTechGrades();
    await cacheInvalidate("work_orders");
    await cacheInvalidate("turn_work_orders");
  }

  return {
    ok: true,
    run: "midnight_reassign",
    timestamp: new Date().toISOString(),
    candidates: candidateCount,
    reassigned: reassigned.length,
    skipped: skipped.length,
    escalated: escalated.length,
    details: { reassigned, skipped, escalated, stale_refreshed: staleRefreshed },
  };
}