// ============================================================================
// handlers/workOrders.ts — Work order data handlers.
//
// Contains all handlers that read or resolve work order records:
//   handleWorkOrders      — bulk WO fetch via Reports API v2 (cached)
//   handleWoNotes         — single WO notes via DB API v0 (no cache / short TTL)
//   handleTurnWorkOrders  — DB API v0, client-side Unit Turn filter (cached)
//   handleLabor           — work_order_labor_summary report (cached)
//   handleRecentTasks     — DB API v0 /tasks, capped at 50 (cached)
//
// AssignedUsers on any PATCH must reference a valid AppFolio Maintenance Tech
// role user or the request returns 422 "User not found" .
// Simultaneous PATCHes to the same WO will cause the second to fail .
// All write operations live in lib/appfolio.ts — this file is read-only.
// ============================================================================
// ============================================================================
// handlers/workOrders.ts — Work order data handlers.
// ============================================================================

import { cacheGet, cacheSet, rowsAsObjects, sqlite, upsertWorkOrderRows } from "../db.ts";
import { AF_DB, AF_REPORTS, dbHeaders } from "../config.ts";
import { fetchDbApi, fetchReport } from "../lib/appfolio.ts";
import { daysAgo, fetchWithTimeout, snapDays, today } from "../lib/fetchUtils.ts";

function isUuid(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i
    .test(String(value || "").trim());
}

async function resolveWoUuidFromReference(
  woRef: string,
): Promise<string> {
  const ref = String(woRef || "").trim();
  if (!ref) return "";
  if (isUuid(ref)) return ref;

  // 1) Check turn tracker linkage table first (fast local lookup).
  try {
    const linkRows = rowsAsObjects(
      await sqlite.execute({
        sql: `SELECT wo_db_uuid
            FROM unit_turn_work_orders
            WHERE wo_id = ? AND wo_db_uuid IS NOT NULL AND wo_db_uuid <> ''
            ORDER BY updated_at DESC
            LIMIT 1`,
        args: [ref],
      }),
    );
    const linked = String(linkRows[0]?.wo_db_uuid || "").trim();
    if (isUuid(linked)) return linked;
  } catch {
    // Non-fatal; continue to next resolver.
  }

  // 2) Check webhook state cache (WO snapshots keyed by UUID with wo_number).
  try {
    const stateRows = rowsAsObjects(
      await sqlite.execute({
        sql: `SELECT id
            FROM wo_states
            WHERE wo_number = ?
            ORDER BY updated_at DESC
            LIMIT 1`,
        args: [ref],
      }),
    );
    const cachedId = String(stateRows[0]?.id || "").trim();
    if (isUuid(cachedId)) return cachedId;
  } catch {
    // Non-fatal; continue to network resolver.
  }

  // 3) Last resort: ask Reports API for the work order row and extract UUID.
  try {
    const rows = await fetchReport("work_order", {
      work_order_numbers: [ref],
      paginate_results: true,
      per_page: 1,
      status_date: "0",
    });
    if (Array.isArray(rows) && rows.length > 0) {
      const candidate = String(
        rows[0]?.work_order_id || rows[0]?.Id || rows[0]?.id || "",
      ).trim();
      if (isUuid(candidate)) return candidate;
    }
  } catch {
    // Non-fatal; caller will return a validation error.
  }

  return "";
}

export async function handleWorkOrders(
  params: Record<string, string>,
): Promise<any> {
  const days = snapDays(parseInt(params.days || "180", 10), "work_orders");
  const cacheKey = `work_orders_${days}`;

  const cached = await cacheGet(cacheKey, "work_orders");
  if (cached) {
    return {
      ok: true,
      results: cached.data,
      count: cached.record_count,
      cached_at: cached.cached_at,
      from_cache: true,
    };
  }

  const results = await fetchReport("work_order", {
    work_order_statuses: ["0", "1", "2", "9", "3", "6", "8", "12"],
    status_date_range_from: daysAgo(days),
    status_date: "0",
  });

  await cacheSet(cacheKey, "work_orders", results, results.length);
  upsertWorkOrderRows(results).catch(() => {});
  return { ok: true, results, count: results.length, from_cache: false };
}

export async function handleCompletedWorkOrdersHistory(
  params: Record<string, string>,
): Promise<any> {
  const requestedFrom = String(params.from_date || "").slice(0, 10);
  const requestedTo = String(params.to_date || "").slice(0, 10);
  const hasExplicitWindow = !!requestedFrom;
  const days = snapDays(
    parseInt(params.days || "30", 10),
    "work_orders_history",
  );
  const effectiveFrom = hasExplicitWindow ? requestedFrom : daysAgo(days);
  const effectiveTo = requestedTo || today();
  const cacheKey = hasExplicitWindow
    ? `work_orders_completed_history_${effectiveFrom}_${effectiveTo}`
    : `work_orders_completed_history_${days}`;

  const cached = await cacheGet(cacheKey, "work_orders_history");
  if (cached) {
    return {
      ok: true,
      results: cached.data,
      count: cached.record_count,
      cached_at: cached.cached_at,
      from_cache: true,
    };
  }

  const results = await fetchReport("work_order", {
    work_order_statuses: ["4", "5", "7"],
    status_date_range_from: effectiveFrom,
    status_date_range_to: effectiveTo,
    status_date: "0",
  });

  await cacheSet(cacheKey, "work_orders_history", results, results.length);
  return { ok: true, results, count: results.length, from_cache: false };
}

export async function handleWoNotes(
  params: Record<string, string>,
): Promise<any> {
  const woRef = params.wo_id || params.uuid || params.id || "";
  const woId = await resolveWoUuidFromReference(String(woRef || ""));
  if (!woId) {
    return {
      ok: false,
      results: [],
      error: "Missing valid work order UUID",
    };
  }

  const cacheKey = `wo_notes_${woId}`;
  const NOTES_TTL = 5;
  const notesExpiresAt = new Date(Date.now() + NOTES_TTL * 60_000)
    .toISOString();

  try {
    const result = await sqlite.execute({
      sql: `SELECT data, cached_at FROM api_cache
             WHERE cache_key = ?
               AND datetime(cached_at, '+5 minutes') > datetime('now')`,
      args: [cacheKey],
    });
    const rows = rowsAsObjects(result);
    if (rows.length > 0) {
      return { ok: true, results: JSON.parse(rows[0].data), from_cache: true };
    }
  } catch { /* cache miss */ }

  try {
    const notesPath = `/api/v0/work_orders/${encodeURIComponent(woId)}/notes`;

    let notesResp = await fetchWithTimeout(`${AF_DB}${notesPath}`, {
      headers: dbHeaders(),
    });

    let notesDomain = AF_DB;
    if (notesResp.status === 401) {
      notesDomain = AF_REPORTS;
      notesResp = await fetchWithTimeout(`${AF_REPORTS}${notesPath}`, {
        headers: dbHeaders(),
      });
    }

    if (notesResp.status === 404) {
      const legacyPath =
        `/api/v0/work_order_notes?filters[WorkOrderId]=${encodeURIComponent(woId)}&page[size]=200`;
      notesResp = await fetchWithTimeout(`${AF_DB}${legacyPath}`, {
        headers: dbHeaders(),
      });
      notesDomain = AF_DB;
      if (notesResp.status === 401) {
        notesDomain = AF_REPORTS;
        notesResp = await fetchWithTimeout(`${AF_REPORTS}${legacyPath}`, {
          headers: dbHeaders(),
        });
      }
    }

    if (!notesResp.ok) {
      const errBody = await notesResp.text().catch(() => "");
      return {
        ok: false,
        error: `wo_notes fetch failed: HTTP ${notesResp.status}`,
        status: notesResp.status,
        domain: notesDomain,
        detail: errBody.substring(0, 500),
      };
    }

    const notesData = await notesResp.json();
      // AppFolio v0 list endpoints return { items: [...] } — check items first.
      const notes = notesData.items ||
        notesData.Items ||
        notesData.results ||
        notesData.Results ||
        notesData.data ||
        notesData.Data ||
        (Array.isArray(notesData) ? notesData : []);

    try {
      await sqlite.execute({
        sql: `INSERT OR REPLACE INTO api_cache
                 (cache_key, entity_type, data, cached_at, expires_at, record_count)
               VALUES (?, 'wo_notes', ?, datetime('now'), ?, ?)`,
        args: [cacheKey, JSON.stringify(notes), notesExpiresAt, notes.length],
      });
    } catch { /* non-fatal */ }

    return {
      ok: true,
      results: notes,
      count: notes.length,
      domain: notesDomain,
    };
  } catch (err: any) {
    return { ok: false, error: err.message, results: [] };
  }
}

export async function handleWoDetail(
  params: Record<string, string>,
): Promise<any> {
  const woRef = String(params.uuid || params.wo_id || params.id || "").trim();
  if (!woRef) {
    return { ok: false, status: 400, error: "Missing uuid/wo_id parameter" };
  }

  const woUuid = await resolveWoUuidFromReference(woRef);
  if (!woUuid) {
    return {
      ok: false,
      status: 404,
      error: "Could not resolve work order UUID",
      wo_ref: woRef,
    };
  }

  const cacheKey = `wo_detail_${woUuid}`;
  try {
    const cacheRes = await sqlite.execute({
      sql: `SELECT data, cached_at FROM api_cache
            WHERE cache_key = ?
              AND datetime(cached_at, '+5 minutes') > datetime('now')`,
      args: [cacheKey],
    });
    const rows = rowsAsObjects(cacheRes);
    if (rows.length > 0) {
      return {
        ok: true,
        uuid: woUuid,
        result: JSON.parse(String(rows[0].data || "{}")),
        from_cache: true,
      };
    }
  } catch {
    // Non-fatal cache miss.
  }

  const path = `/api/v0/work_orders/${encodeURIComponent(woUuid)}`;
  let resp = await fetchWithTimeout(`${AF_DB}${path}`, {
    headers: dbHeaders(),
  });
  let sourceDomain = AF_DB;
  if (resp.status === 401 || resp.status === 403 || resp.status === 404) {
    sourceDomain = AF_REPORTS;
    resp = await fetchWithTimeout(`${AF_REPORTS}${path}`, {
      headers: dbHeaders(),
    });
  }

  if (!resp.ok) {
    const errBody = await resp.text().catch(() => "");
    return {
      ok: false,
      status: resp.status,
      error: `wo_detail fetch failed: HTTP ${resp.status}`,
      domain: sourceDomain,
      detail: errBody.substring(0, 500),
      uuid: woUuid,
    };
  }

  const payload = await resp.json().catch(() => ({}));
  const record =
    (payload?.data && !Array.isArray(payload.data) ? payload.data : null) ||
    payload?.results?.[0] ||
    payload ||
    null;

  try {
    await sqlite.execute({
      sql: `INSERT OR REPLACE INTO api_cache
            (cache_key, entity_type, data, cached_at, expires_at, record_count)
            VALUES (?, 'wo_detail', ?, datetime('now'), datetime('now', '+5 minutes'), 1)`,
      args: [cacheKey, JSON.stringify(record || {})],
    });
  } catch {
    // Non-fatal cache write miss.
  }

  return {
    ok: true,
    uuid: woUuid,
    result: record,
    from_domain: sourceDomain,
  };
}

export async function handleWoNoteCreate(req: Request): Promise<any> {
  let body: any = {};
  try {
    body = await req.json();
  } catch {
    body = {};
  }

  const woRef = String(body.uuid || body.wo_id || body.id || "").trim();
  const noteBody = String(body.body_text || body.body || "").trim();
  if (!woRef || !noteBody) {
    return {
      ok: false,
      status: 400,
      error: "uuid and body_text are required",
    };
  }

  const woUuid = await resolveWoUuidFromReference(woRef);
  if (!woUuid) {
    return {
      ok: false,
      status: 404,
      error: "Could not resolve work order UUID",
      wo_ref: woRef,
    };
  }

  const path = `/api/v0/work_orders/${encodeURIComponent(woUuid)}/notes`;
  const postInit = {
    method: "POST",
    headers: {
      ...dbHeaders(),
      "Content-Type": "application/json",
      "Accept": "application/json",
    },
    body: JSON.stringify({
      Body: noteBody,
    }),
  };

  let resp = await fetchWithTimeout(`${AF_DB}${path}`, postInit);
  let sourceDomain = AF_DB;
  if (resp.status === 401 || resp.status === 403 || resp.status === 404) {
    sourceDomain = AF_REPORTS;
    resp = await fetchWithTimeout(`${AF_REPORTS}${path}`, postInit);
  }

  if (!resp.ok) {
    const errText = await resp.text().catch(() => "");
    return {
      ok: false,
      status: resp.status,
      error: "appfolio_error",
      message: errText || `HTTP ${resp.status}`,
      domain: sourceDomain,
      uuid: woUuid,
    };
  }

  try {
    await sqlite.execute({
      sql: `DELETE FROM api_cache WHERE cache_key = ?`,
      args: [`wo_notes_${woUuid}`],
    });
  } catch {
    // Non-fatal cache bust best effort.
  }

  return {
    ok: true,
    uuid: woUuid,
  };
}

export async function handleTurnWorkOrders(
  params: Record<string, string>,
): Promise<any> {
  const days = snapDays(parseInt(params.days || "90", 10), "turn_work_orders");
  const cacheKey = `turn_work_orders_${days}`;

  const cached = await cacheGet(cacheKey, "turn_work_orders");
  if (cached) {
    return {
      ok: true,
      results: cached.data,
      count: cached.record_count,
      cached_at: cached.cached_at,
      from_cache: true,
    };
  }

  const fromDate = new Date();
  fromDate.setDate(fromDate.getDate() - days);

  const path = `/api/v0/work_orders?filters[LastUpdatedAtFrom]=${
    encodeURIComponent(fromDate.toISOString())
  }&page[size]=200`;

  const rows = await fetchDbApi(path, 500);
  const turnWOs = rows.filter((wo: any) =>
    String(wo?.Type || wo?.type || "").trim().toLowerCase() === "unit turn"
  );

  await cacheSet(cacheKey, "turn_work_orders", turnWOs, turnWOs.length);
  return {
    ok: true,
    results: turnWOs,
    count: turnWOs.length,
    from_cache: false,
  };
}

export async function handleLabor(
  params: Record<string, string>,
): Promise<any> {
  const days = parseInt(params.days || "1", 10);
  const statusCodes = params.statuses || "8";
  const cacheKey = `labor_${days}_${statusCodes}`;

  const cached = await cacheGet(cacheKey, "labor");
  if (cached) {
    return {
      ok: true,
      results: cached.data,
      count: cached.record_count,
      cached_at: cached.cached_at,
      from_cache: true,
    };
  }

  const results = await fetchReport("work_order_labor_summary", {
    work_order_statuses: statusCodes.split(","),
    labor_performed_from: daysAgo(days),
    labor_performed_to: today(),
  });

  await cacheSet(cacheKey, "labor", results, results.length);
  return { ok: true, results, count: results.length, from_cache: false };
}

export async function handleRecentTasks(
  _params: Record<string, string>,
): Promise<any> {
  const cacheKey = "recent_tasks";

  const cached = await cacheGet(cacheKey, "recent_tasks");
  if (cached) {
    return {
      ok: true,
      results: cached.data,
      count: cached.record_count,
      cached_at: cached.cached_at,
      from_cache: true,
    };
  }

  try {
    const results = await fetchDbApi("/api/v0/tasks?page[size]=50", 50);
    await cacheSet(cacheKey, "recent_tasks", results, results.length);
    return { ok: true, results, count: results.length, from_cache: false };
  } catch (err: any) {
    console.log(`recent_tasks fetch failed: ${err.message?.substring(0, 120)}`);
    return {
      ok: true,
      results: [],
      count: 0,
      from_cache: false,
      warning: "Tasks API unavailable for this account",
    };
  }
}

function parseMoneyLike(value: any): number {
  if (value === null || value === undefined) return 0;
  const n = Number(String(value).replace(/[^0-9.\-]/g, ""));
  return Number.isFinite(n) ? n : 0;
}

function sumBilledFromRows(rows: any[]): number {
  const keys = [
    "line_item_total",
    "line_item_amount",
    "billable_amount",
    "amount_billed",
    "total",
    "amount",
  ];
  let total = 0;
  for (const row of rows || []) {
    for (const key of keys) {
      if (
        row && row[key] !== undefined && row[key] !== null &&
        String(row[key]).trim() !== ""
      ) {
        total += parseMoneyLike(row[key]);
        break;
      }
    }
  }
  return Number(total.toFixed(2));
}

export async function handleWoBilledAmount(
  params: Record<string, string>,
): Promise<any> {
  const woNumber = String(params.wo_number || "").trim();
  if (!woNumber) {
    return { ok: false, status: 400, error: "Missing wo_number" };
  }

  const cacheKey = `wo_billed_${woNumber}`;
  const cached = await cacheGet(cacheKey, "bills");
  if (cached && cached.data && typeof cached.data.total_billed === "number") {
    return {
      ok: true,
      wo_number: woNumber,
      total_billed: cached.data.total_billed,
      rows_used: cached.data.rows_used || 0,
      source_report: cached.data.source_report || "cached",
      from_cache: true,
    };
  }

  // Try line-item style report first, then fallback to work_order report fields.
  const reportAttempts: Array<{ name: string; filters: Record<string, any> }> =
    [
      {
        name: "work_order_labor_summary",
        filters: {
          paginate_results: true,
          per_page: 200,
          work_order_numbers: [woNumber],
        },
      },
      {
        name: "work_order",
        filters: {
          paginate_results: true,
          per_page: 50,
          work_order_numbers: [woNumber],
          status_date: "0",
        },
      },
    ];

  for (const attempt of reportAttempts) {
    try {
      const rows = await fetchReport(attempt.name, attempt.filters);
      const total = sumBilledFromRows(rows || []);
      if (rows.length > 0) {
        const payload = {
          total_billed: total,
          rows_used: rows.length,
          source_report: attempt.name,
        };
        await cacheSet(cacheKey, "bills", payload, 1);
        return {
          ok: true,
          wo_number: woNumber,
          total_billed: total,
          rows_used: rows.length,
          source_report: attempt.name,
          from_cache: false,
        };
      }
    } catch {
      // Try next report variant.
    }
  }

  return {
    ok: true,
    wo_number: woNumber,
    total_billed: 0,
    rows_used: 0,
    source_report: "none",
    warning: "No billed line-item rows found",
    from_cache: false,
  };
}