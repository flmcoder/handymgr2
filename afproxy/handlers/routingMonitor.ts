import { rowsAsObjects, sqlite } from "../db.ts";

async function ensureRoutingTables(): Promise<void> {
  await sqlite.execute(`CREATE TABLE IF NOT EXISTS routing_capabilities (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    trade TEXT NOT NULL,
    keywords_json TEXT NOT NULL DEFAULT '[]',
    active INTEGER DEFAULT 1,
    created_at TEXT DEFAULT (datetime('now')),
    updated_at TEXT DEFAULT (datetime('now'))
  )`);
  await sqlite.execute(`CREATE TABLE IF NOT EXISTS routing_pm_map (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    group_name TEXT NOT NULL,
    pm_name TEXT NOT NULL,
    created_at TEXT DEFAULT (datetime('now')),
    updated_at TEXT DEFAULT (datetime('now'))
  )`);
  await sqlite.execute(`CREATE TABLE IF NOT EXISTS routing_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    wo_id TEXT,
    wo_number TEXT,
    event_type TEXT,
    trade TEXT,
    review_status TEXT DEFAULT 'pending',
    review_notes TEXT,
    source_json TEXT,
    created_at TEXT DEFAULT (datetime('now')),
    updated_at TEXT DEFAULT (datetime('now'))
  )`);
}

function num(value: string | undefined, fallback: number): number {
  const v = Number(value || fallback);
  return Number.isFinite(v) ? v : fallback;
}

export async function handleRoutingMonitor(
  params: Record<string, string>,
  req: Request,
): Promise<any> {
  await ensureRoutingTables();
  const op = String(params.op || "").trim().toLowerCase();

  if (req.method === "GET") {
    if (op === "capabilities") {
      const res = await sqlite.execute(
        `SELECT id, trade, keywords_json, active, created_at, updated_at
           FROM routing_capabilities
          ORDER BY trade ASC, id DESC`,
      );
      const rows = rowsAsObjects(res).map((r: any) => ({
        id: Number(r.id || 0),
        trade: String(r.trade || ""),
        keywords: (() => {
          try {
            const parsed = JSON.parse(String(r.keywords_json || "[]"));
            return Array.isArray(parsed) ? parsed : [];
          } catch {
            return [];
          }
        })(),
        active: Number(r.active || 0),
        created_at: String(r.created_at || ""),
        updated_at: String(r.updated_at || ""),
      }));
      return { ok: true, results: rows, count: rows.length };
    }

    if (op === "pm_map") {
      const res = await sqlite.execute(
        `SELECT id, group_name, pm_name, created_at, updated_at
           FROM routing_pm_map
          ORDER BY group_name ASC, id DESC`,
      );
      const rows = rowsAsObjects(res);
      return { ok: true, results: rows, count: rows.length };
    }

    if (op === "events") {
      const status = String(params.status || "pending").trim().toLowerCase();
      const limit = Math.max(1, Math.min(500, num(params.limit, 200)));
      const res = await sqlite.execute({
        sql: `SELECT id, wo_id, wo_number, event_type, trade, review_status, review_notes, source_json, created_at, updated_at
                FROM routing_events
               WHERE (? = '' OR review_status = ?)
               ORDER BY id DESC
               LIMIT ?`,
        args: [status === "all" ? "" : status, status === "all" ? "" : status, limit],
      });
      const rows = rowsAsObjects(res).map((r: any) => ({
        ...r,
        source: (() => {
          try {
            return JSON.parse(String(r.source_json || "{}"));
          } catch {
            return {};
          }
        })(),
      }));
      return { ok: true, results: rows, count: rows.length };
    }

    if (op === "pm_stats") {
      const res = await sqlite.execute(
        `SELECT COALESCE(review_status, 'pending') AS review_status, COUNT(1) AS cnt
           FROM routing_events
          GROUP BY COALESCE(review_status, 'pending')`,
      );
      const rows = rowsAsObjects(res);
      return { ok: true, stats: rows };
    }

    return { ok: false, status: 400, error: "Unknown routing monitor op" };
  }

  let body: any = {};
  try {
    body = await req.json();
  } catch {
    body = {};
  }

  if (op === "capability_upsert") {
    const id = Number(body.id || 0);
    const trade = String(body.trade || "").trim();
    const keywords = Array.isArray(body.keywords)
      ? body.keywords.map((k: any) => String(k || "").trim()).filter(Boolean)
      : [];
    const active = Number(body.active || 0) === 0 ? 0 : 1;
    if (!trade) return { ok: false, status: 400, error: "trade is required" };

    if (id > 0) {
      await sqlite.execute({
        sql: `UPDATE routing_capabilities
                 SET trade = ?, keywords_json = ?, active = ?, updated_at = datetime('now')
               WHERE id = ?`,
        args: [trade, JSON.stringify(keywords), active, id],
      });
      return { ok: true, id, trade, keywords, active };
    }

    await sqlite.execute({
      sql: `INSERT INTO routing_capabilities (trade, keywords_json, active, created_at, updated_at)
            VALUES (?, ?, ?, datetime('now'), datetime('now'))`,
      args: [trade, JSON.stringify(keywords), active],
    });
    return { ok: true, trade, keywords, active };
  }

  if (op === "capability_delete") {
    const id = Number(body.id || 0);
    if (!id) return { ok: false, status: 400, error: "id is required" };
    const res: any = await sqlite.execute({
      sql: `DELETE FROM routing_capabilities WHERE id = ?`,
      args: [id],
    });
    return { ok: true, deleted: Number(res?.rowsAffected || 0) > 0 };
  }

  if (op === "pm_map_upsert") {
    const groupName = String(body.group_name || "").trim();
    const pmName = String(body.pm_name || "").trim();
    if (!groupName || !pmName) {
      return { ok: false, status: 400, error: "group_name and pm_name are required" };
    }
    await sqlite.execute({
      sql: `INSERT INTO routing_pm_map (group_name, pm_name, created_at, updated_at)
            VALUES (?, ?, datetime('now'), datetime('now'))`,
      args: [groupName, pmName],
    });
    return { ok: true, group_name: groupName, pm_name: pmName };
  }

  if (op === "pm_map_bulk") {
    const entries = Array.isArray(body.entries) ? body.entries : [];
    let inserted = 0;
    for (const entry of entries) {
      const groupName = String(entry?.group_name || "").trim();
      const pmName = String(entry?.pm_name || "").trim();
      if (!groupName || !pmName) continue;
      await sqlite.execute({
        sql: `INSERT INTO routing_pm_map (group_name, pm_name, created_at, updated_at)
              VALUES (?, ?, datetime('now'), datetime('now'))`,
        args: [groupName, pmName],
      });
      inserted += 1;
    }
    return { ok: true, inserted };
  }

  if (op === "review") {
    const id = Number(body.id || 0);
    const status = String(body.review_status || "pending").trim().toLowerCase();
    const notes = String(body.review_notes || "").trim();
    if (!id) return { ok: false, status: 400, error: "id is required" };
    await sqlite.execute({
      sql: `UPDATE routing_events
               SET review_status = ?, review_notes = ?, updated_at = datetime('now')
             WHERE id = ?`,
      args: [status, notes, id],
    });
    return { ok: true, id, review_status: status };
  }

  if (op === "scan") {
    const events = Array.isArray(body.events) ? body.events : [];
    let inserted = 0;
    for (const event of events) {
      const woId = String(event?.wo_id || event?.work_order_id || "").trim();
      const woNumber = String(event?.wo_number || event?.work_order_number || "").trim();
      const eventType = String(event?.event_type || event?.type || "routing_event").trim();
      const trade = String(event?.trade || event?.category || "").trim();
      await sqlite.execute({
        sql: `INSERT INTO routing_events
              (wo_id, wo_number, event_type, trade, review_status, source_json, created_at, updated_at)
              VALUES (?, ?, ?, ?, 'pending', ?, datetime('now'), datetime('now'))`,
        args: [woId, woNumber, eventType, trade, JSON.stringify(event || {})],
      });
      inserted += 1;
    }
    return { ok: true, inserted };
  }

  return { ok: false, status: 400, error: "Unknown routing monitor op" };
}
