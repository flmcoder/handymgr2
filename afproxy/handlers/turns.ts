// ============================================================================
// handlers/turns.ts — Unit turn pipeline + local turn record store.
//
// Exports:
//   handleTurns           — unit_turn_detail report (Reports API v2, cached 30 min)
//   handleTurnRecords     — local SQLite turn record GET + POST
//   handleTurnRecordStage — local SQLite turn stage PATCH (POST only)
//
// Turn completeness rule (core HandyManager domain logic):
//   A turn is ONLY "Complete" if it has an explicit turnEnd date from AppFolio
//   OR if all associated work orders are in a terminal status
//   (Completed, Work Completed, Canceled).
//   This cannot be derived from the unit_turn_detail report alone — the
//   turn_records table stores the enriched local state to bridge this gap.
//
// The unit turn work order category field requires a valid UnitId or
// OccupancyId if provided .
// The next page of paginated results is found in the next_page_path field
// of the response body .
// ============================================================================

import { cacheGet, cacheSet, rowsAsObjects, sqlite } from "../db.ts";
import { propertyInScope, resolveGroupPropertyIds } from "../lib/groupUtils.ts";
import { fetchDbApi, fetchReport } from "../lib/appfolio.ts";
import { daysAgo, snapDays } from "../lib/fetchUtils.ts";

const turnsIncrementalLocks = new Set<string>();
let lastTurnsIncrementalResult: any = {
  ok: true,
  skipped: true,
  reason: "cold_start",
  since: "",
  synced_at: "",
  count: 0,
  results: [],
};
let lastTurnsIncrementalAt = 0;
const TURNS_INCREMENTAL_COALESCE_MS = 15_000;

function isValidUUID(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i
    .test(String(value || "").trim());
}

function isIsoDateTime(value: string): boolean {
  return /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z$/.test(
    String(value || "").trim(),
  );
}

function trackingCodeFromRow(row: Record<string, any>): string {
  const short = String(row.tracking_uuid || "").slice(0, 8).toUpperCase();
  const dt = String(row.move_out_date || "").slice(0, 10).replace(/-/g, "") ||
    new Date().toISOString().slice(0, 10).replace(/-/g, "");
  const prop =
    String(row.property_id || "P").replace(/[^A-Za-z0-9]/g, "").slice(0, 8) ||
    "P";
  const unit =
    String(row.unit_id || "U").replace(/[^A-Za-z0-9]/g, "").slice(0, 8) || "U";
  return `TURN-${dt}-${prop}-${unit}-${short || "NEW"}`;
}

function safeParseJSON<T>(raw: any, fallback: T): T {
  if (raw === null || raw === undefined) return fallback;
  if (typeof raw === "object") return raw as T;
  const str = String(raw || "").trim();
  if (!str) return fallback;
  try {
    return JSON.parse(str) as T;
  } catch {
    return fallback;
  }
}

function resolvedTurnStatus(row: Record<string, any>): string {
  const terminalStatuses = new Set([
    "completed",
    "work completed",
    "canceled",
    "cancelled",
  ]);
  const current = String(row.status || row.turn_status || "").trim();
  const metadata = safeParseJSON<Record<string, any>>(row.metadata, {});
  const linked = Array.isArray(row.linked_work_orders)
    ? row.linked_work_orders
    : [];
  const allWosTerminal = linked.length > 0
    ? linked.every((w: any) =>
      terminalStatuses.has(String(w?.status || "").toLowerCase())
    )
    : !!metadata.all_wos_terminal;

  const hasTurnEnd =
    !!(row.closed_at || metadata.turn_end || metadata.turnEnd || row.turn_end);
  if (hasTurnEnd || allWosTerminal) return "Complete";
  return current || "In Progress";
}

function resolvePropertyGroupId(row: Record<string, any>): string {
  if (row.property_group_id) return String(row.property_group_id);
  const metadata = safeParseJSON<Record<string, any>>(row.metadata, {});
  return String(
    metadata.property_group_id ||
      metadata.propertyGroupId ||
      metadata.group_id ||
      metadata.groupId ||
      "",
  );
}

function normalizeTurnStatusFilter(value: string): string {
  const normalized = String(value || "").trim().toLowerCase();
  if (!normalized || normalized === "all") return "all";
  if (["complete", "completed", "closed"].includes(normalized)) {
    return "Complete";
  }
  if (["in progress", "in_progress", "active", "open"].includes(normalized)) {
    return "In Progress";
  }
  return String(value || "").trim();
}

async function loadTrackerBundle(days: number): Promise<any[]> {
  const rows = rowsAsObjects(
    await sqlite.execute({
      sql: `SELECT *
          FROM unit_turn_tracker
          WHERE coalesce(move_out_date, updated_at) >= datetime('now', ?)
             OR status IN ('active', 'on_radar')
          ORDER BY updated_at DESC`,
      args: [`-${days} days`],
    }),
  );

  if (rows.length === 0) return [];

  const byUuid: Record<string, any> = {};
  rows.forEach((r) => {
    byUuid[String(r.tracking_uuid)] = {
      ...r,
      tracking_code: r.tracking_code || trackingCodeFromRow(r),
      milestones: {},
      linked_work_orders: [],
    };
  });

  let milestones: Record<string, any>[] = [];
  try {
    milestones = rowsAsObjects(
      await sqlite.execute(
        `SELECT tracking_uuid, milestone_key, milestone_date, source, notes, updated_at
       FROM unit_turn_milestones
       ORDER BY updated_at DESC`,
      ),
    );
  } catch (err: any) {
    console.warn("[unit_turns milestones]", String(err?.message || err));
  }
  milestones.forEach((m) => {
    const key = String(m.tracking_uuid || "");
    if (!byUuid[key]) return;
    byUuid[key].milestones[m.milestone_key] = {
      date: m.milestone_date,
      source: m.source || "",
      notes: m.notes || "",
      updated_at: m.updated_at,
    };
  });

  let linked: Record<string, any>[] = [];
  try {
    linked = rowsAsObjects(
      await sqlite.execute(
        `SELECT tracking_uuid, wo_id, wo_db_uuid, source, status, created_at, removed, updated_at
       FROM unit_turn_work_orders
       WHERE removed = 0
       ORDER BY updated_at DESC`,
      ),
    );
  } catch (err: any) {
    console.warn("[unit_turns linked_wos]", String(err?.message || err));
  }
  linked.forEach((w) => {
    const key = String(w.tracking_uuid || "");
    if (!byUuid[key]) return;
    byUuid[key].linked_work_orders.push(w);
  });

  return Object.values(byUuid);
}

// ── handleTurns ───────────────────────────────────────────────────────────────
// Fetches unit turn detail records via Reports API v2.
// Default: last 90 days, status "In Progress".
// The move_out_date_from filter controls the lookback window.
// Cache key includes both days and status so multiple filter combinations
// are stored independently.
export async function handleTurns(
  params: Record<string, string>,
): Promise<any> {
  const days = snapDays(parseInt(params.days || "90", 10), "turns");
  const status = params.status || "In Progress";
  const cacheKey = `turns_${days}_${status.replace(/\s+/g, "_")}`;

  const cached = await cacheGet(cacheKey, "turns");
  let results: any[];
  let fromCache = false;
  let cachedAt = "";

  if (cached) {
    results = cached.data;
    fromCache = true;
    cachedAt = String(cached.cached_at || "");
  } else {
    results = await fetchReport("unit_turn_detail", {
      move_out_date_from: daysAgo(days),
      unit_turn_status: status,
      property_visibility: "active",
    });
    await cacheSet(cacheKey, "turns", results, results.length);
  }

  // Server-side PM scope enforcement.
  const allowedIds = await resolveGroupPropertyIds(params);
  if (allowedIds) {
    results = results.filter((r: any) =>
      propertyInScope(String(r.property_id || r.PropertyId || ""), allowedIds)
    );
  }

  const resp: Record<string, any> = { ok: true, results, count: results.length, from_cache: fromCache };
  if (cachedAt) resp.cached_at = cachedAt;
  return resp;
}

// ── handleTurnsIncremental ──────────────────────────────────────────────────
// Incremental turn refresh from DB API v0 using LastUpdatedAtFrom filter.
// Returns a shape compatible with the frontend turns mapper.
export async function handleTurnsIncremental(
  params: Record<string, string>,
): Promise<any> {
  const rawSince = String(params.since || "").trim();
  const since = isIsoDateTime(rawSince)
    ? rawSince
    : new Date(Date.now() - 6 * 60 * 60 * 1000).toISOString();
  const limit = Math.max(
    1,
    Math.min(parseInt(params.limit || "500", 10), 2000),
  );
  const lockKey = "turns_incremental";

  const now = Date.now();
  if (now - lastTurnsIncrementalAt < TURNS_INCREMENTAL_COALESCE_MS) {
    return {
      ...lastTurnsIncrementalResult,
      ok: true,
      coalesced: true,
      skipped: true,
      reason: "coalesced_recent_sync",
    };
  }

  if (turnsIncrementalLocks.has(lockKey)) {
    return {
      ...lastTurnsIncrementalResult,
      ok: true,
      skipped: true,
      reason: "sync_in_progress",
      coalesced: true,
    };
  }

  turnsIncrementalLocks.add(lockKey);
  try {
    const path = `/api/v0/unit_turns?filters[LastUpdatedAtFrom]=${
      encodeURIComponent(since)
    }&page[size]=100`;
    const rows = await fetchDbApi(path, limit);

    const results = rows
      .filter((t: Record<string, any>) => {
        const unitId = String(t.UnitId || t.unit_id || t.unit_uuid || "")
          .trim();
        return !unitId || isValidUUID(unitId);
      })
      .map((t: Record<string, any>) => ({
        unit_turn_id: t.Id || t.id || t.UnitTurnId || t.unit_turn_id || "",
        unit_turn_uuid: t.Id || t.id || t.unit_turn_uuid || "",
        unit: t.UnitAddress || t.unit || t.unit_name || t.Unit || "",
        unit_name: t.unit_name || t.Unit || t.UnitAddress || "",
        property: t.PropertyAddress || t.property || t.property_name ||
          t.PropertyName || "",
        property_name: t.property_name || t.PropertyName || t.PropertyAddress ||
          "",
        property_id: t.PropertyId || t.property_id || t.property_uuid || "",
        unit_id: t.UnitId || t.unit_id || t.unit_uuid || "",
        move_out_date: t.MoveOutDate || t.move_out_date || "",
        expected_move_in_date: t.ExpectedMoveInDate ||
          t.expected_move_in_date || "",
        turn_end_date: t.TurnEnd || t.turn_end || t.turn_end_date || "",
        unit_turn_status: t.Status || t.unit_turn_status || t.status || "",
        site_manager: t.SiteManagerName || t.site_manager ||
          t.property_site_manager || "",
        bedrooms: t.Bedrooms ?? t.bedrooms ?? null,
        bathrooms: t.Bathrooms ?? t.bathrooms ?? null,
        dogs_allowed: t.DogsAllowed ?? t.dogs_allowed ?? "",
        insurance_exp: t.InsuranceExpiresAt ?? t.insurance_exp ?? "",
        LastUpdatedAt: t.LastUpdatedAt || t.last_updated_at || t.UpdatedAt ||
          t.updated_at || "",
      }));

    const payload = {
      ok: true,
      since,
      synced_at: new Date().toISOString(),
      count: results.length,
      results,
    };
    lastTurnsIncrementalResult = payload;
    lastTurnsIncrementalAt = Date.now();
    return payload;
  } finally {
    turnsIncrementalLocks.delete(lockKey);
  }
}

// ── handleUnitTurns ───────────────────────────────────────────────────────────
// Fetches live unit turn records via DB API v0 /api/v0/unit_turns/.
// Provides richer status, deposit tracking, and move-in scheduling data
// not available in the Reports API unit_turn_detail report.
// Cache TTL: 30 minutes (same cadence as Reports turns).
export async function handleUnitTurns(
  params: Record<string, string>,
): Promise<any> {
  const days = snapDays(parseInt(params.days || "180", 10), "unit_turns");
  const pgFilter = String(params.property_group_id || "").trim();
  const statusFilter = normalizeTurnStatusFilter(String(params.status || ""));
  const limit = Math.max(1, Math.min(parseInt(params.limit || "50", 10), 200));
  const offset = Math.max(parseInt(params.offset || "0", 10), 0);
  const forceSync = String(params.sync || "").toLowerCase() === "true";
  const cacheKey = `unit_turns_sql_${days}`;

  if (forceSync) {
    // Best-effort refresh trigger; non-fatal if unavailable.
    try {
      await handleTurnsIncremental({
        since: new Date(Date.now() - 6 * 60 * 60 * 1000).toISOString(),
        limit: "2000",
      });
    } catch (err: any) {
      console.warn("[unit_turns sync]", String(err?.message || err));
    }
  }

  const cached = await cacheGet(cacheKey, "unit_turns");
  let allResults: any[];
  let fromCache = false;
  let cachedAt = "";

  if (cached && Array.isArray(cached.data)) {
    allResults = cached.data;
    fromCache = true;
    cachedAt = String(cached.cached_at || "");
  } else {
    allResults = await loadTrackerBundle(days);
    await cacheSet(cacheKey, "unit_turns", allResults, allResults.length);
  }

  const filtered = allResults
    .map((row: any) => ({ ...row, turn_status: resolvedTurnStatus(row) }))
    .filter((row: any) => {
      if (statusFilter && statusFilter !== "all") {
        if (String(row.turn_status || "").toLowerCase() !== statusFilter.toLowerCase()) return false;
      }
      if (pgFilter && pgFilter !== "all") {
        if (resolvePropertyGroupId(row) !== pgFilter) return false;
      }
      return true;
    });

  const total = filtered.length;
  const paged = filtered.slice(offset, offset + limit);

  return {
    ok: true,
    mode: "sql_tracker",
    total,
    limit,
    offset,
    count: paged.length,
    results: paged,
    from_cache: fromCache,
    cached_at: cachedAt || undefined,
  };
}

// Batch upsert inferred + confirmed turn tracker records from frontend pipeline.
export async function handleUnitTurnsSync(req: Request): Promise<any> {
  let body: any = {};
  try {
    body = await req.json();
  } catch {
    return { ok: false, error: "Invalid JSON body" };
  }

  const records = Array.isArray(body.records) ? body.records : [];
  if (records.length === 0) return { ok: true, upserted: 0 };

  let upserted = 0;
  for (const rec of records) {
    const turnKey = String(rec.turn_key || rec.turnKey || "").trim();
    if (!turnKey) continue;

    const existing = rowsAsObjects(
      await sqlite.execute({
        sql:
          `SELECT tracking_uuid, tracking_code FROM unit_turn_tracker WHERE turn_key = ? LIMIT 1`,
        args: [turnKey],
      }),
    );

    const trackingUuid =
      (existing[0] && String(existing[0].tracking_uuid || "")) ||
      String(rec.tracking_uuid || rec.trackingUuid || crypto.randomUUID());

    const rowForCode = {
      tracking_uuid: trackingUuid,
      move_out_date: rec.move_out_date || rec.moveOutDate || "",
      property_id: rec.property_id || rec.propertyId || "",
      unit_id: rec.unit_id || rec.unitId || "",
    };
    const trackingCode =
      (existing[0] && String(existing[0].tracking_code || "")) ||
      String(
        rec.tracking_code || rec.trackingCode ||
          trackingCodeFromRow(rowForCode),
      );

    const status = String(rec.status || "on_radar");
    const closedAtInput = rec.closed_at || rec.closedAt || null;
    const normalizedClosedAt = (status === "closed" || status === "completed")
      ? (closedAtInput || new Date().toISOString())
      : null;

    await sqlite.execute({
      sql: `INSERT INTO unit_turn_tracker (
              tracking_uuid, tracking_code, turn_key, unit_turn_id,
              unit_id, property_id, unit_name, property_name,
              move_out_date, move_in_date, inspection_date, first_wo_date,
              estimate_requested_date, estimate_received_date,
              status, confidence_score, confidence_label,
              site_manager, source_flags, metadata, closed_at, updated_at
            ) VALUES (
              ?, ?, ?, ?,
              ?, ?, ?, ?,
              ?, ?, ?, ?,
              ?, ?,
              ?, ?, ?,
              ?, ?, ?, ?, datetime('now')
            )
            ON CONFLICT(turn_key) DO UPDATE SET
              unit_turn_id            = excluded.unit_turn_id,
              unit_id                 = excluded.unit_id,
              property_id             = excluded.property_id,
              unit_name               = excluded.unit_name,
              property_name           = excluded.property_name,
              move_out_date           = excluded.move_out_date,
              move_in_date            = excluded.move_in_date,
              inspection_date         = excluded.inspection_date,
              first_wo_date           = excluded.first_wo_date,
              estimate_requested_date = excluded.estimate_requested_date,
              estimate_received_date  = excluded.estimate_received_date,
              status                  = excluded.status,
              confidence_score        = excluded.confidence_score,
              confidence_label        = excluded.confidence_label,
              site_manager            = excluded.site_manager,
              source_flags            = excluded.source_flags,
              metadata                = excluded.metadata,
              closed_at               = excluded.closed_at,
              updated_at              = datetime('now')`,
      args: [
        trackingUuid,
        trackingCode,
        turnKey,
        rec.unit_turn_id || rec.unitTurnId || "",
        rec.unit_id || rec.unitId || "",
        rec.property_id || rec.propertyId || "",
        rec.unit_name || rec.unitName || "",
        rec.property_name || rec.propertyName || "",
        rec.move_out_date || rec.moveOutDate || "",
        rec.move_in_date || rec.moveInDate || "",
        rec.inspection_date || rec.inspectionDate || "",
        rec.first_wo_date || rec.firstWoDate || "",
        rec.estimate_requested_date || rec.estimateRequestedDate || "",
        rec.estimate_received_date || rec.estimateReceivedDate || "",
        status,
        parseInt(
          String(rec.confidence_score || rec.confidenceScore || 0),
          10,
        ) || 0,
        rec.confidence_label || rec.confidenceLabel || "low",
        rec.site_manager || rec.siteManager || "",
        JSON.stringify(rec.source_flags || rec.sourceFlags || {}),
        JSON.stringify(rec.metadata || {}),
        normalizedClosedAt,
      ],
    });

    const milestones = Array.isArray(rec.milestones)
      ? rec.milestones
      : Object.entries(rec.milestones || {}).map(([k, v]: [string, any]) => ({
        key: k,
        date: v && typeof v === "object" ? v.date : v,
        source: v && typeof v === "object" ? v.source : "derived",
        notes: v && typeof v === "object" ? v.notes : "",
      }));
    for (const m of milestones) {
      const mKey = String(m.key || m.milestone_key || "").trim();
      if (!mKey) continue;
      await sqlite.execute({
        sql:
          `INSERT INTO unit_turn_milestones (tracking_uuid, milestone_key, milestone_date, source, notes, updated_at)
              VALUES (?, ?, ?, ?, ?, datetime('now'))
              ON CONFLICT(tracking_uuid, milestone_key) DO UPDATE SET
                milestone_date = excluded.milestone_date,
                source         = excluded.source,
                notes          = excluded.notes,
                updated_at     = datetime('now')`,
        args: [
          trackingUuid,
          mKey,
          m.date || m.milestone_date || null,
          m.source || "derived",
          m.notes || "",
        ],
      });
    }

    const linked = Array.isArray(rec.work_orders)
      ? rec.work_orders
      : Array.isArray(rec.linked_work_orders)
      ? rec.linked_work_orders
      : [];

    if (rec.replace_work_orders) {
      await sqlite.execute({
        sql:
          `UPDATE unit_turn_work_orders SET removed = 1, updated_at = datetime('now') WHERE tracking_uuid = ?`,
        args: [trackingUuid],
      });
    }

    for (const w of linked) {
      const woId = String(w.wo_id || w.id || "").trim();
      if (!woId) continue;
      await sqlite.execute({
        sql:
          `INSERT INTO unit_turn_work_orders (tracking_uuid, wo_id, wo_db_uuid, source, status, created_at, removed, updated_at)
              VALUES (?, ?, ?, ?, ?, ?, 0, datetime('now'))
              ON CONFLICT(tracking_uuid, wo_id) DO UPDATE SET
                wo_db_uuid = excluded.wo_db_uuid,
                source     = excluded.source,
                status     = excluded.status,
                created_at = excluded.created_at,
                removed    = 0,
                updated_at = datetime('now')`,
        args: [
          trackingUuid,
          woId,
          w.wo_db_uuid || w.dbApiId || "",
          w.source || "inferred",
          w.status || "",
          w.created_at || w.created || null,
        ],
      });
    }

    upserted++;
  }

  const refreshedBundle = await loadTrackerBundle(180);
  await cacheSet(
    `unit_turns_sql_180`,
    "unit_turns",
    refreshedBundle,
    refreshedBundle.length,
  );
  return { ok: true, upserted };
}

export async function handleUnitTurnWorkOrderLink(req: Request): Promise<any> {
  let body: any = {};
  try {
    body = await req.json();
  } catch {
    return { ok: false, error: "Invalid JSON body" };
  }

  const turnKey = String(body.turn_key || body.turnKey || "").trim();
  const woId = String(body.wo_id || body.woId || "").trim();
  if (!turnKey || !woId) {
    return { ok: false, error: "turn_key and wo_id are required" };
  }

  const existing = rowsAsObjects(
    await sqlite.execute({
      sql:
        `SELECT tracking_uuid FROM unit_turn_tracker WHERE turn_key = ? LIMIT 1`,
      args: [turnKey],
    }),
  );
  let trackingUuid = existing.length > 0
    ? String(existing[0].tracking_uuid || "")
    : crypto.randomUUID();
  if (existing.length === 0) {
    await sqlite.execute({
      sql:
        `INSERT INTO unit_turn_tracker (tracking_uuid, tracking_code, turn_key, status, updated_at)
            VALUES (?, ?, ?, 'on_radar', datetime('now'))`,
      args: [
        trackingUuid,
        trackingCodeFromRow({ tracking_uuid: trackingUuid }),
        turnKey,
      ],
    });
  }

  await sqlite.execute({
    sql:
      `INSERT INTO unit_turn_work_orders (tracking_uuid, wo_id, wo_db_uuid, source, status, created_at, removed, updated_at)
          VALUES (?, ?, ?, ?, ?, ?, 0, datetime('now'))
          ON CONFLICT(tracking_uuid, wo_id) DO UPDATE SET
            wo_db_uuid = excluded.wo_db_uuid,
            source     = excluded.source,
            status     = excluded.status,
            removed    = 0,
            updated_at = datetime('now')`,
    args: [
      trackingUuid,
      woId,
      body.wo_db_uuid || body.dbApiId || "",
      body.source || "manual",
      body.status || "",
      body.created_at || body.created || null,
    ],
  });

  return { ok: true, tracking_uuid: trackingUuid, wo_id: woId };
}

export async function handleUnitTurnWorkOrderUnlink(
  req: Request,
): Promise<any> {
  let body: any = {};
  try {
    body = await req.json();
  } catch {
    return { ok: false, error: "Invalid JSON body" };
  }

  const turnKey = String(body.turn_key || body.turnKey || "").trim();
  const woId = String(body.wo_id || body.woId || "").trim();
  if (!turnKey || !woId) {
    return { ok: false, error: "turn_key and wo_id are required" };
  }

  const existing = rowsAsObjects(
    await sqlite.execute({
      sql:
        `SELECT tracking_uuid FROM unit_turn_tracker WHERE turn_key = ? LIMIT 1`,
      args: [turnKey],
    }),
  );
  if (existing.length === 0) return { ok: false, error: "turn_key not found" };
  const trackingUuid = String(existing[0].tracking_uuid || "");

  await sqlite.execute({
    sql: `UPDATE unit_turn_work_orders
          SET removed = 1, updated_at = datetime('now')
          WHERE tracking_uuid = ? AND wo_id = ?`,
    args: [trackingUuid, woId],
  });

  return { ok: true, tracking_uuid: trackingUuid, wo_id: woId };
}

// ── handleTurnRecords ─────────────────────────────────────────────────────────
// Local SQLite turn record store — read (GET) and write (POST).
//
// GET:  Returns all locally stored turn records ordered by last update.
//       These records are enriched with local stage/status data that AppFolio
//       does not expose natively (e.g. cleaning stage, key handoff, etc.)
//
// POST: Upserts a turn record by id / unit_turn_id.
//       Body: { id | unit_turn_id, ...any additional fields }
//       The full body is stored as JSON — no schema enforcement at this layer.
//
// Completeness check note:
//   When evaluating whether a turn is complete, the caller must verify that
//   either turnEnd is set (from AppFolio) OR all associated work orders have
//   reached a terminal status. The work order category field on unit turns
//   requires a valid UnitId or OccupancyId .
export async function handleTurnRecords(
  params: Record<string, string>,
  req?: Request,
): Promise<any> {
  // ── POST: upsert a turn record ─────────────────────────────────────────────
  if (req && req.method === "POST") {
    try {
      const body = await req.json();
      const id = body.id || body.unit_turn_id;
      if (!id) {
        return { ok: false, error: "Missing id or unit_turn_id in body" };
      }

      await sqlite.execute({
        sql:
          `INSERT OR REPLACE INTO turn_records (unit_turn_id, data, updated_at)
               VALUES (?, ?, datetime('now'))`,
        args: [id, JSON.stringify(body)],
      });

      return { ok: true, message: `Turn record saved for ${id}` };
    } catch (err: any) {
      return { ok: false, error: err.message };
    }
  }

  // ── GET: return all turn records ───────────────────────────────────────────
  try {
    const result = await sqlite.execute(
      `SELECT unit_turn_id, data, updated_at
       FROM turn_records
       ORDER BY updated_at DESC`,
    );
    const rows = rowsAsObjects(result);
    const records = rows.map((r) => {
      try {
        return JSON.parse(r.data);
      } catch {
        return { unit_turn_id: r.unit_turn_id, raw: r.data };
      }
    });

    return { ok: true, records, count: records.length };
  } catch (err: any) {
    return { ok: false, error: err.message };
  }
}

// ── handleTurnRecordStage ─────────────────────────────────────────────────────
// POST-only endpoint to update a specific stage within an existing turn record.
//
// Body:
//   {
//     id:    string,    // unit_turn_id — required
//     stage: string,    // e.g. "cleaning", "paint", "key_handoff"
//     data:  object     // merged into record.stages[stage]
//   }
//
// If the turn record does not exist yet, a skeleton record is created.
// Each stage object is deep-merged with the existing stage data and stamped
// with an updatedAt timestamp.
//
// This endpoint exists because the unit turn work order category field on
// AppFolio requires a valid UnitId or OccupancyId , and not all
// stage transitions map to discrete AppFolio work order statuses. Local
// stage tracking supplements AppFolio's native turn pipeline.
export async function handleTurnRecordStage(req: Request): Promise<any> {
  try {
    const body = await req.json();
    const id = body.id;
    const stage = body.stage;
    const stageData = body.data;

    if (!id) return { ok: false, error: "Missing id (unit_turn_id)" };
    if (!stage) return { ok: false, error: "Missing stage name" };

    // Load existing record or start fresh
    const existing = await sqlite.execute({
      sql: `SELECT data FROM turn_records WHERE unit_turn_id = ?`,
      args: [id],
    });
    const existingRows = rowsAsObjects(existing);

    let record: any = {};
    if (existingRows.length > 0) {
      try {
        record = JSON.parse(existingRows[0].data);
      } catch { /* start fresh if parse fails */ }
    }

    // Deep merge stage data
    if (!record.stages) record.stages = {};
    record.stages[stage] = {
      ...record.stages[stage],
      ...stageData,
      updatedAt: new Date().toISOString(),
    };

    // Stamp top-level metadata
    record.unit_turn_id = id;
    record.updatedAt = new Date().toISOString();

    await sqlite.execute({
      sql: `INSERT OR REPLACE INTO turn_records (unit_turn_id, data, updated_at)
             VALUES (?, ?, datetime('now'))`,
      args: [id, JSON.stringify(record)],
    });

    return {
      ok: true,
      message: `Stage '${stage}' saved for turn ${id}`,
      stages: Object.keys(record.stages),
    };
  } catch (err: any) {
    return { ok: false, error: err.message };
  }
}

// ── handleClosedTurns ─────────────────────────────────────────────────────────
// GET  ?action=closed_turns  → returns all closed turn IDs
// POST ?action=closed_turns  body: { turn_id }  → marks a turn as closed
export async function handleClosedTurns(
  _params: Record<string, string>,
  req?: Request,
): Promise<any> {
  if (req?.method === "POST") {
    let body: any = {};
    try {
      body = await req.json();
    } catch {
      return { ok: false, error: "Invalid JSON body" };
    }
    const { turn_id } = body;
    if (!turn_id) return { ok: false, error: "turn_id is required" };
    try {
      await sqlite.execute({
        sql: `INSERT INTO closed_turns (
                turn_id, closed_at, close_reason, close_source, closed_by,
                property_id, property_name, unit_id, unit_name, move_out_date, move_in_date
              ) VALUES (
                ?, datetime('now'), ?, ?, ?,
                ?, ?, ?, ?, ?, ?
              )
              ON CONFLICT(turn_id) DO UPDATE SET
                closed_at = datetime('now'),
                close_reason = excluded.close_reason,
                close_source = excluded.close_source,
                closed_by = excluded.closed_by,
                property_id = excluded.property_id,
                property_name = excluded.property_name,
                unit_id = excluded.unit_id,
                unit_name = excluded.unit_name,
                move_out_date = excluded.move_out_date,
                move_in_date = excluded.move_in_date`,
        args: [
          String(turn_id),
          String(body.close_reason || body.reason || "manual_close"),
          String(body.close_source || "manual"),
          String(body.closed_by || "ui"),
          String(body.property_id || ""),
          String(body.property_name || ""),
          String(body.unit_id || ""),
          String(body.unit_name || ""),
          String(body.move_out_date || ""),
          String(body.move_in_date || ""),
        ],
      });
      return { ok: true, turn_id };
    } catch (e: any) {
      return { ok: false, error: e.message };
    }
  }

  // GET — return all closed turn IDs
  try {
    const rows = rowsAsObjects(
      await sqlite.execute(
        `SELECT turn_id, closed_at FROM closed_turns ORDER BY closed_at DESC`,
      ),
    );
    return { ok: true, results: rows, count: rows.length };
  } catch (e: any) {
    return { ok: false, error: e.message };
  }
}

export async function handleUnitTurnsHistory(
  params: Record<string, string>,
): Promise<any> {
  try {
    const days = parseInt(params.days || "365", 10);
    const limit = Math.max(
      1,
      Math.min(500, parseInt(params.limit || "200", 10)),
    );
    const rows = rowsAsObjects(
      await sqlite.execute({
        sql: `SELECT
              t.tracking_uuid,
              t.tracking_code,
              t.turn_key,
              t.unit_turn_id,
              t.unit_id,
              t.property_id,
              t.unit_name,
              t.property_name,
              t.move_out_date,
              t.move_in_date,
              t.status,
              t.confidence_score,
              t.confidence_label,
              t.site_manager,
              t.closed_at,
              c.closed_at AS manually_closed_at,
              c.close_source,
              c.close_reason,
              c.closed_by
            FROM unit_turn_tracker t
            LEFT JOIN closed_turns c ON c.turn_id = t.turn_key
            WHERE (
              t.status IN ('closed', 'completed')
              OR c.turn_id IS NOT NULL
            )
              AND coalesce(t.closed_at, c.closed_at, t.move_in_date, t.updated_at) >= datetime('now', ?)
            ORDER BY coalesce(t.closed_at, c.closed_at, t.move_in_date, t.updated_at) DESC
            LIMIT ?`,
        args: [`-${days} days`, limit],
      }),
    );

    const results = rows.map((r: Record<string, any>) => {
      const closeSource = r.close_source ||
        ((r.status === "closed" && r.manually_closed_at)
          ? "manual"
          : ((r.move_in_date &&
              new Date(String(r.move_in_date)).getTime() <= Date.now())
            ? "system_move_in"
            : "system"));
      const closeReason = r.close_reason ||
        (closeSource === "system_move_in"
          ? "move_in_detected"
          : (r.status === "completed" ? "completed" : "closed"));
      return {
        tracking_uuid: r.tracking_uuid,
        tracking_code: r.tracking_code,
        turn_key: r.turn_key,
        unit_turn_id: r.unit_turn_id,
        unit_id: r.unit_id,
        property_id: r.property_id,
        unit_name: r.unit_name,
        property_name: r.property_name,
        move_out_date: r.move_out_date,
        move_in_date: r.move_in_date,
        status: r.status,
        confidence_score: r.confidence_score,
        confidence_label: r.confidence_label,
        site_manager: r.site_manager,
        closed_at: r.closed_at || r.manually_closed_at || r.move_in_date || "",
        close_source: closeSource,
        close_reason: closeReason,
        closed_by: r.closed_by || "system",
      };
    });

    // Server-side PM scope enforcement.
    const allowedIds = await resolveGroupPropertyIds(params);
    const scoped = allowedIds
      ? results.filter((r: any) => propertyInScope(String(r.property_id || ""), allowedIds))
      : results;

    return { ok: true, results: scoped, count: scoped.length };
  } catch (e: any) {
    return { ok: false, error: e.message };
  }
}