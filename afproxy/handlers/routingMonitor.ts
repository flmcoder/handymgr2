// ============================================================================
// handlers/routingMonitor.ts — PM routing / vendor leakage monitor
//
// Tracks work orders routed to external vendors when work appears in-scope
// for in-house teams. Supports:
// - capability registry (trade + keywords + active toggle)
// - property-group to PM mapping
// - flagged event upsert + review workflow
// - PM leaderboard stats
// ============================================================================

import { rowsAsObjects, sqlite, ensureTables } from "../db.ts";

const DEFAULT_CAPABILITIES: Array<{ trade: string; keywords: string[] }> = [
  {
    trade: "HVAC",
    keywords: [
      "hvac",
      "ac",
      "air",
      "thermostat",
      "filter",
      "condenser",
      "compressor",
      "heat",
    ],
  },
  {
    trade: "Plumbing",
    keywords: [
      "plumb",
      "toilet",
      "faucet",
      "sink",
      "drain",
      "pipe",
      "leak",
      "water heater",
      "shower",
    ],
  },
  {
    trade: "Electrical",
    keywords: [
      "electrical",
      "outlet",
      "switch",
      "breaker",
      "circuit",
      "wiring",
      "light fixture",
    ],
  },
  {
    trade: "Appliances",
    keywords: [
      "appliance",
      "refrigerator",
      "dishwasher",
      "washer",
      "dryer",
      "oven",
      "stove",
      "microwave",
    ],
  },
  {
    trade: "Painting",
    keywords: ["paint", "painting", "touch up", "touch-up", "repaint"],
  },
  {
    trade: "Landscaping",
    keywords: [
      "landscape",
      "lawn",
      "grass",
      "sprinkler",
      "irrigation",
      "tree",
      "yard",
    ],
  },
  {
    trade: "Drywall",
    keywords: [
      "drywall",
      "sheetrock",
      "patch",
      "texture",
      "plaster",
      "wall repair",
    ],
  },
  {
    trade: "Locksmith",
    keywords: ["lock", "key", "rekey", "deadbolt", "lockout", "smart lock"],
  },
  {
    trade: "Cleaning",
    keywords: [
      "clean",
      "cleaning",
      "janitorial",
      "make ready",
      "turnover",
      "sanitize",
    ],
  },
  {
    trade: "General",
    keywords: [
      "handyman",
      "maintenance",
      "minor repair",
      "caulk",
      "hinge",
      "door handle",
    ],
  },
];

function asInt(v: unknown, fallback = 0): number {
  const n = Number(v);
  return Number.isFinite(n) ? Math.trunc(n) : fallback;
}

function normalizeConfidence(v: unknown): string {
  const s = String(v || "").trim().toLowerCase();
  if (s === "high" || s === "medium" || s === "low") return s;
  return "medium";
}

function normalizeReviewStatus(v: unknown): string {
  const s = String(v || "").trim().toLowerCase();
  if (
    s === "pending" ||
    s === "approved_external" ||
    s === "reassign_inhouse" ||
    s === "dismissed"
  ) return s;
  return "pending";
}

// routing_capabilities, routing_events, and routing_pm_group_map tables are created by db.ts ensureTables()

async function seedDefaultCapabilitiesIfEmpty(): Promise<void> {
  await ensureTables();
  const countRes = await sqlite.execute(
    "SELECT COUNT(*) AS c FROM routing_capabilities",
  );
  const countRows = rowsAsObjects(countRes);
  const existing = asInt(countRows[0]?.c, 0);
  if (existing > 0) return;

  for (const cap of DEFAULT_CAPABILITIES) {
    await sqlite.execute({
      sql:
        `INSERT INTO routing_capabilities (trade, keywords, active, created_at, updated_at)
            VALUES (?, ?, 1, datetime('now'), datetime('now'))`,
      args: [cap.trade, JSON.stringify(cap.keywords)],
    });
  }
}

async function getCapabilities(): Promise<any[]> {
  await seedDefaultCapabilitiesIfEmpty();
  const res = await sqlite.execute(
    `SELECT id, trade, keywords, active, created_at, updated_at
     FROM routing_capabilities
     ORDER BY trade ASC`,
  );
  return rowsAsObjects(res);
}

async function getPmMap(): Promise<any[]> {
  await ensureTables();
  const res = await sqlite.execute(
    `SELECT group_name, pm_name, updated_at
     FROM routing_pm_group_map
     ORDER BY group_name ASC`,
  );
  return rowsAsObjects(res);
}

async function getEvents(params: Record<string, string>): Promise<any> {
  await ensureTables();
  const status = String(params.status || "pending").trim().toLowerCase();
  const pm = String(params.pm || "").trim();
  const days = Math.max(1, Math.min(365, asInt(params.days, 30)));
  const limit = Math.max(1, Math.min(1000, asInt(params.limit, 300)));

  const where: string[] = ["datetime(detected_at) >= datetime('now', ?)"];
  const args: any[] = [`-${days} days`];

  if (status && status !== "all") {
    where.push("review_status = ?");
    args.push(status);
  }
  if (pm && pm !== "all") {
    where.push("pm_name = ?");
    args.push(pm);
  }

  args.push(limit);

  const res = await sqlite.execute({
    sql: `SELECT
            id, wo_uuid, wo_number, property_id, property_name, unit_name,
            property_group, pm_name, vendor_id, vendor_name, vendor_category,
            wo_status, wo_priority, wo_created_at, description,
            matched_trade, confidence, review_status, review_notes,
            reviewed_at, detected_at, source
          FROM routing_events
          WHERE ${where.join(" AND ")}
          ORDER BY datetime(detected_at) DESC
          LIMIT ?`,
    args,
  });

  return {
    ok: true,
    results: rowsAsObjects(res),
  };
}

async function getPmStats(params: Record<string, string>): Promise<any> {
  const days = Math.max(1, Math.min(365, asInt(params.days, 30)));

  const res = await sqlite.execute({
    sql: `SELECT
            pm_name,
            COUNT(*) AS total_flagged,
            SUM(CASE WHEN review_status = 'pending' THEN 1 ELSE 0 END) AS pending,
            SUM(CASE WHEN review_status = 'approved_external' THEN 1 ELSE 0 END) AS approved_external,
            SUM(CASE WHEN review_status = 'reassign_inhouse' THEN 1 ELSE 0 END) AS reassigned,
            SUM(CASE WHEN confidence = 'high' THEN 1 ELSE 0 END) AS high_confidence,
            MAX(detected_at) AS last_flagged
          FROM routing_events
          WHERE datetime(detected_at) >= datetime('now', ?)
          GROUP BY pm_name
          ORDER BY total_flagged DESC, pm_name ASC`,
    args: [`-${days} days`],
  });

  return {
    ok: true,
    results: rowsAsObjects(res),
    days,
  };
}

async function upsertEventsFromScan(body: any): Promise<any> {
  await ensureTables();
  const events = Array.isArray(body?.events) ? body.events : [];
  let upserted = 0;

  for (const ev of events) {
    const woUuid = String(ev.wo_uuid || ev.work_order_uuid || "").trim();
    if (!woUuid) continue;

    await sqlite.execute({
      sql: `INSERT INTO routing_events (
              wo_uuid, wo_number, property_id, property_name, unit_name,
              property_group, pm_name,
              vendor_id, vendor_name, vendor_category,
              wo_status, wo_priority, wo_created_at, description,
              matched_trade, confidence,
              review_status, review_notes, reviewed_at,
              detected_at, source
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                     COALESCE((SELECT review_status FROM routing_events WHERE wo_uuid = ?), 'pending'),
                     COALESCE((SELECT review_notes FROM routing_events WHERE wo_uuid = ?), NULL),
                     COALESCE((SELECT reviewed_at FROM routing_events WHERE wo_uuid = ?), NULL),
                     datetime('now'), ?)
            ON CONFLICT(wo_uuid) DO UPDATE SET
              wo_number = excluded.wo_number,
              property_id = excluded.property_id,
              property_name = excluded.property_name,
              unit_name = excluded.unit_name,
              property_group = excluded.property_group,
              pm_name = excluded.pm_name,
              vendor_id = excluded.vendor_id,
              vendor_name = excluded.vendor_name,
              vendor_category = excluded.vendor_category,
              wo_status = excluded.wo_status,
              wo_priority = excluded.wo_priority,
              wo_created_at = excluded.wo_created_at,
              description = excluded.description,
              matched_trade = excluded.matched_trade,
              confidence = excluded.confidence,
              detected_at = datetime('now'),
              source = excluded.source`,
      args: [
        woUuid,
        String(ev.wo_number || ""),
        String(ev.property_id || ""),
        String(ev.property_name || ""),
        String(ev.unit_name || ""),
        String(ev.property_group || ""),
        String(ev.pm_name || "Unmapped PM"),
        String(ev.vendor_id || ""),
        String(ev.vendor_name || ""),
        String(ev.vendor_category || ""),
        String(ev.wo_status || ""),
        String(ev.wo_priority || ""),
        String(ev.wo_created_at || ""),
        String(ev.description || ""),
        String(ev.matched_trade || ""),
        normalizeConfidence(ev.confidence),
        woUuid,
        woUuid,
        woUuid,
        String(ev.source || "client_scan"),
      ],
    });
    upserted++;
  }

  return { ok: true, upserted, received: events.length };
}

async function upsertCapability(body: any): Promise<any> {
  await ensureTables();
  const id = body?.id;
  const trade = String(body?.trade || "").trim();
  const active = asInt(body?.active, 1) ? 1 : 0;
  const keywords = Array.isArray(body?.keywords)
    ? body.keywords.map((k: unknown) => String(k).trim()).filter(Boolean)
    : [];

  if (!trade) return { ok: false, error: "trade is required" };

  if (id) {
    await sqlite.execute({
      sql: `UPDATE routing_capabilities
            SET trade = ?, keywords = ?, active = ?, updated_at = datetime('now')
            WHERE id = ?`,
      args: [trade, JSON.stringify(keywords), active, asInt(id)],
    });
  } else {
    await sqlite.execute({
      sql:
        `INSERT INTO routing_capabilities (trade, keywords, active, created_at, updated_at)
            VALUES (?, ?, ?, datetime('now'), datetime('now'))
            ON CONFLICT(trade) DO UPDATE SET
              keywords = excluded.keywords,
              active = excluded.active,
              updated_at = datetime('now')`,
      args: [trade, JSON.stringify(keywords), active],
    });
  }

  return { ok: true };
}

async function deleteCapability(body: any): Promise<any> {
  await ensureTables();
  const id = asInt(body?.id, 0);
  if (!id) return { ok: false, error: "id is required" };
  await sqlite.execute({
    sql: `DELETE FROM routing_capabilities WHERE id = ?`,
    args: [id],
  });
  return { ok: true };
}

async function upsertPmMap(body: any): Promise<any> {
  await ensureTables();
  const groupName = String(body?.group_name || "").trim();
  const pmName = String(body?.pm_name || "").trim();
  if (!groupName || !pmName) {
    return { ok: false, error: "group_name and pm_name are required" };
  }

  await sqlite.execute({
    sql: `INSERT INTO routing_pm_group_map (group_name, pm_name, updated_at)
          VALUES (?, ?, datetime('now'))
          ON CONFLICT(group_name) DO UPDATE SET
            pm_name = excluded.pm_name,
            updated_at = datetime('now')`,
    args: [groupName, pmName],
  });
  return { ok: true };
}

async function bulkUpsertPmMap(body: any): Promise<any> {
  await ensureTables();
  const entries = Array.isArray(body?.entries) ? body.entries : [];
  let count = 0;

  for (const row of entries) {
    const groupName = String(row?.group_name || "").trim();
    const pmName = String(row?.pm_name || "").trim();
    if (!groupName || !pmName) continue;

    await sqlite.execute({
      sql: `INSERT INTO routing_pm_group_map (group_name, pm_name, updated_at)
            VALUES (?, ?, datetime('now'))
            ON CONFLICT(group_name) DO UPDATE SET
              pm_name = excluded.pm_name,
              updated_at = datetime('now')`,
      args: [groupName, pmName],
    });
    count++;
  }

  return { ok: true, upserted: count };
}

async function updateReview(body: any): Promise<any> {
  await ensureTables();
  const id = asInt(body?.id, 0);
  if (!id) return { ok: false, error: "id is required" };

  await sqlite.execute({
    sql: `UPDATE routing_events
          SET review_status = ?,
              review_notes = ?,
              reviewed_at = datetime('now')
          WHERE id = ?`,
    args: [
      normalizeReviewStatus(body?.review_status),
      body?.review_notes ? String(body.review_notes) : null,
      id,
    ],
  });

  return { ok: true };
}

export async function handleRoutingMonitor(
  params: Record<string, string>,
  req?: Request,
): Promise<any> {
  const opFromParams = String(params.op || "").trim().toLowerCase();

  if (req?.method === "POST") {
    let body: any = {};
    try {
      body = await req.json();
    } catch {
      return { ok: false, error: "Invalid JSON body" };
    }

    const op = String(opFromParams || body?.op || "").trim().toLowerCase();

    // CRITICAL: routing events can be destructive; require admin auth for any POST
    if (!op || op === "scan") {
      const { PROXY_ADMIN_KEY } = await import("../config.ts");
      const bodyKey = String(body?.key || "").trim();
      if (!PROXY_ADMIN_KEY || !bodyKey || bodyKey !== PROXY_ADMIN_KEY) {
        return {
          ok: false,
          error: "Unauthorized — missing or invalid PROXY_ADMIN_KEY",
        };
      }
    }

    switch (op) {
      case "scan":
        return await upsertEventsFromScan(body);
      case "review":
        return await updateReview(body);
      case "capability_upsert":
        return await upsertCapability(body);
      case "capability_delete":
        return await deleteCapability(body);
      case "pm_map_upsert":
        return await upsertPmMap(body);
      case "pm_map_bulk":
        return await bulkUpsertPmMap(body);
      default:
        return { ok: false, error: `Unknown POST op: ${op || "(empty)"}` };
    }
  }

  const op = opFromParams;

  switch (op) {
    case "events":
      return await getEvents(params);
    case "pm_stats":
      return await getPmStats(params);
    case "capabilities":
      return { ok: true, results: await getCapabilities() };
    case "pm_map":
      return { ok: true, results: await getPmMap() };
    default:
      return { ok: false, error: `Unknown GET op: ${op || "(empty)"}` };
  }
}