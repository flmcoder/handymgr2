import { rowsAsObjects, sqlite } from "../db.ts";

async function ensureQueueTables(): Promise<void> {
  await sqlite.execute(`CREATE TABLE IF NOT EXISTS reassignment_queue (
    wo_id TEXT PRIMARY KEY,
    wo_number TEXT,
    property_id TEXT,
    property_group_uuid TEXT,
    property_addr TEXT,
    category TEXT,
    priority TEXT,
    status TEXT,
    assigned_to TEXT,
    auto_exempt INTEGER DEFAULT 0,
    auto_exempt_at TEXT,
    auto_exempt_by TEXT,
    updated_at TEXT DEFAULT (datetime('now')),
    created_at TEXT DEFAULT (datetime('now'))
  )`);

  await sqlite.execute(`CREATE TABLE IF NOT EXISTS tech_grades (
    tech_id TEXT PRIMARY KEY,
    tech_name TEXT,
    tier INTEGER DEFAULT 1,
    grade REAL DEFAULT 0,
    jobs_completed INTEGER DEFAULT 0,
    no_contact_count INTEGER DEFAULT 0,
    active INTEGER DEFAULT 1,
    updated_at TEXT DEFAULT (datetime('now')),
    created_at TEXT DEFAULT (datetime('now'))
  )`);

  await sqlite.execute(`CREATE TABLE IF NOT EXISTS reassignment_audit (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    wo_id TEXT,
    event_type TEXT,
    event_message TEXT,
    payload_json TEXT,
    created_at TEXT DEFAULT (datetime('now'))
  )`);
}

function toInt(value: string | undefined, fallback: number): number {
  const parsed = Number(value || fallback);
  return Number.isFinite(parsed) ? parsed : fallback;
}

export async function handleReassignmentQueue(
  params: Record<string, string>,
): Promise<any> {
  await ensureQueueTables();

  if (String(params.sync_assignees || "") === "1") {
    return {
      ok: true,
      synced: 0,
      message: "sync_assignees accepted; no-op in lightweight fallback handler",
    };
  }

  const limit = Math.max(1, Math.min(500, toInt(params.limit, 100)));
  const woId = String(params.wo_id || "").trim();

  const queueRes = woId
    ? await sqlite.execute({
      sql: `SELECT * FROM reassignment_queue WHERE wo_id = ? ORDER BY updated_at DESC LIMIT ?`,
      args: [woId, limit],
    })
    : await sqlite.execute({
      sql: `SELECT * FROM reassignment_queue ORDER BY updated_at DESC LIMIT ?`,
      args: [limit],
    });

  const techRes = await sqlite.execute({
    sql: `SELECT * FROM tech_grades ORDER BY active DESC, grade DESC, tech_name ASC LIMIT 500`,
    args: [],
  });

  const auditRes = await sqlite.execute({
    sql: `SELECT id, wo_id, event_type, event_message, created_at
            FROM reassignment_audit
           ORDER BY id DESC
           LIMIT 300`,
    args: [],
  });

  const queue = rowsAsObjects(queueRes);
  const techs = rowsAsObjects(techRes).map((t: any) => ({
    tech_id: String(t.tech_id || ""),
    tech_name: String(t.tech_name || ""),
    tier: Number(t.tier || 1),
    active: Number(t.active ?? 1),
    grade: Number(t.grade || 0),
    jobs_completed: Number(t.jobs_completed || 0),
    no_contact_count: Number(t.no_contact_count || 0),
    updated_at: String(t.updated_at || ""),
  }));
  const audit = rowsAsObjects(auditRes);

  return {
    ok: true,
    queue,
    tech_roster: techs,
    audit,
    blasts: [],
    tier2_claims: [],
    monitored_work_orders: [],
    stats: {
      total_queue: queue.length,
      total_techs: techs.length,
      pending: queue.filter((q: any) => String(q.status || "pending") === "pending").length,
      exempt: queue.filter((q: any) => Number(q.auto_exempt || 0) === 1).length,
    },
  };
}
