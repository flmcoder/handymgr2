import { rowsAsObjects, sqlite } from "../db.ts";

function parseJson(raw: unknown, fallback: any) {
  if (raw === null || raw === undefined || raw === "") return fallback;
  if (typeof raw === "object") return raw;
  try {
    return JSON.parse(String(raw));
  } catch {
    return fallback;
  }
}

function asString(value: unknown): string {
  return String(value ?? "").trim();
}

function toPositiveInt(value: unknown, fallback: number): number {
  const parsed = parseInt(String(value ?? ""), 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

function makeTrackingUuid(): string {
  if (globalThis.crypto && typeof globalThis.crypto.randomUUID === "function") {
    return globalThis.crypto.randomUUID();
  }
  return `turn-${Date.now()}-${Math.random().toString(16).slice(2, 10)}`;
}

async function getTrackingUuidForTurnKey(turnKey: string): Promise<string | null> {
  const key = asString(turnKey);
  if (!key) return null;
  const res = await sqlite.execute({
    sql: `SELECT tracking_uuid FROM unit_turn_tracker WHERE turn_key = ? LIMIT 1`,
    args: [key],
  });
  const row = rowsAsObjects(res)[0] || null;
  return row ? asString(row.tracking_uuid) : null;
}

async function loadTrackerRows(whereSql: string, args: unknown[] = []): Promise<any[]> {
  const result = await sqlite.execute({
    sql: `SELECT
            tracking_uuid,
            tracking_code,
            turn_key,
            unit_turn_id,
            unit_id,
            property_id,
            unit_name,
            property_name,
            move_out_date,
            move_in_date,
            inspection_date,
            first_wo_date,
            estimate_requested_date,
            estimate_received_date,
            status,
            confidence_score,
            confidence_label,
            site_manager,
            source_flags,
            metadata,
            closed_at,
            created_at,
            updated_at
          FROM unit_turn_tracker
          ${whereSql}
          ORDER BY datetime(COALESCE(updated_at, created_at, '1970-01-01')) DESC`,
    args,
  });
  const rows = rowsAsObjects(result);
  if (!rows.length) return [];

  const trackingIds = rows.map((row) => asString(row.tracking_uuid)).filter(Boolean);
  const milestonesByTracking: Record<string, Record<string, any>> = {};
  const workOrdersByTracking: Record<string, any[]> = {};

  if (trackingIds.length) {
    const milestonePlaceholders = trackingIds.map(() => "?").join(",");
    const milestoneRes = await sqlite.execute({
      sql: `SELECT tracking_uuid, milestone_key, milestone_date, source, notes
            FROM unit_turn_milestones
            WHERE tracking_uuid IN (${milestonePlaceholders})`,
      args: trackingIds,
    });
    for (const row of rowsAsObjects(milestoneRes)) {
      const trackingUuid = asString(row.tracking_uuid);
      if (!trackingUuid) continue;
      if (!milestonesByTracking[trackingUuid]) milestonesByTracking[trackingUuid] = {};
      milestonesByTracking[trackingUuid][asString(row.milestone_key)] = {
        date: asString(row.milestone_date),
        source: asString(row.source),
        notes: asString(row.notes),
      };
    }

    const workOrderRes = await sqlite.execute({
      sql: `SELECT tracking_uuid, wo_id, wo_db_uuid, source, status, created_at, removed
            FROM unit_turn_work_orders
            WHERE tracking_uuid IN (${milestonePlaceholders}) AND COALESCE(removed, 0) = 0
            ORDER BY datetime(COALESCE(created_at, '1970-01-01')) ASC`,
      args: trackingIds,
    });
    for (const row of rowsAsObjects(workOrderRes)) {
      const trackingUuid = asString(row.tracking_uuid);
      if (!trackingUuid) continue;
      if (!workOrdersByTracking[trackingUuid]) workOrdersByTracking[trackingUuid] = [];
      workOrdersByTracking[trackingUuid].push({
        wo_id: asString(row.wo_id),
        wo_db_uuid: asString(row.wo_db_uuid),
        source: asString(row.source) || "manual",
        status: asString(row.status),
        created_at: asString(row.created_at),
      });
    }
  }

  return rows.map((row) => {
    const trackingUuid = asString(row.tracking_uuid);
    return {
      ...row,
      milestones: milestonesByTracking[trackingUuid] || {},
      linked_work_orders: workOrdersByTracking[trackingUuid] || [],
      source_flags: parseJson(row.source_flags, {}),
      metadata: parseJson(row.metadata, {}),
    };
  });
}

export async function handleTurns(params: Record<string, string>): Promise<any> {
  const results = await loadTrackerRows(`WHERE closed_at IS NULL OR closed_at = ''`);
  return { ok: true, results, count: results.length, source: "unit_turn_tracker" };
}

export async function handleUnitTurns(params: Record<string, string>): Promise<any> {
  const days = toPositiveInt(params.days, 90);
  const since = new Date(Date.now() - (days * 86400_000)).toISOString();
  const results = await loadTrackerRows(
    `WHERE (closed_at IS NULL OR closed_at = '') AND datetime(COALESCE(updated_at, created_at, '1970-01-01')) >= datetime(?)`,
    [since],
  );
  return { ok: true, results, count: results.length, source: "unit_turn_tracker" };
}

export async function handleTurnsIncremental(params: Record<string, string>): Promise<any> {
  const since = asString(params.since);
  if (!since) return await handleTurns(params);
  const results = await loadTrackerRows(
    `WHERE datetime(COALESCE(updated_at, created_at, '1970-01-01')) >= datetime(?)`,
    [since],
  );
  return { ok: true, results, count: results.length, source: "unit_turn_tracker" };
}

export async function handleUnitTurnsHistory(params: Record<string, string>): Promise<any> {
  const limit = Math.min(500, toPositiveInt(params.limit, 300));
  const result = await sqlite.execute({
    sql: `SELECT turn_id, closed_at, close_reason, close_source, closed_by,
                 property_id, property_name, unit_id, unit_name, move_out_date, move_in_date
          FROM closed_turns
          ORDER BY datetime(COALESCE(closed_at, '1970-01-01')) DESC
          LIMIT ?`,
    args: [limit],
  });
  return { ok: true, results: rowsAsObjects(result), count: Number(result.rowsAffected || 0) || rowsAsObjects(result).length };
}

export async function handleClosedTurns(params: Record<string, string>, req?: Request): Promise<any> {
  if (req && req.method !== "GET") {
    let body: any = {};
    try { body = await req.json(); } catch { body = {}; }
    const turnId = asString(body.turn_id || body.turnId || params.turn_id || params.turnId);
    if (!turnId) return { ok: false, status: 400, error: "Missing turn_id" };

    await sqlite.execute({
      sql: `INSERT INTO closed_turns (
              turn_id, closed_at, close_reason, close_source, closed_by,
              property_id, property_name, unit_id, unit_name, move_out_date, move_in_date
            ) VALUES (?, datetime('now'), ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
        turnId,
        asString(body.close_reason),
        asString(body.close_source),
        asString(body.closed_by),
        asString(body.property_id),
        asString(body.property_name),
        asString(body.unit_id),
        asString(body.unit_name),
        asString(body.move_out_date),
        asString(body.move_in_date),
      ],
    });

    await sqlite.execute({
      sql: `UPDATE unit_turn_tracker
            SET status = 'closed', closed_at = datetime('now'), updated_at = datetime('now')
            WHERE turn_key = ? OR tracking_uuid = ?`,
      args: [turnId, turnId],
    });

    return { ok: true, turn_id: turnId, closed_at: new Date().toISOString() };
  }

  const result = await sqlite.execute({
    sql: `SELECT turn_id, closed_at, close_reason, close_source, closed_by,
                 property_id, property_name, unit_id, unit_name, move_out_date, move_in_date
          FROM closed_turns
          ORDER BY datetime(COALESCE(closed_at, '1970-01-01')) DESC`,
  });
  return { ok: true, results: rowsAsObjects(result) };
}

export async function handleTurnRecords(params: Record<string, string>, req?: Request): Promise<any> {
  if (req && req.method !== "GET") {
    let body: any = {};
    try { body = await req.json(); } catch { body = {}; }
    const unitTurnId = asString(body.unit_turn_id || body.unitTurnId || body.id || body.turn_id);
    if (!unitTurnId) return { ok: false, status: 400, error: "Missing unit_turn_id" };
    await sqlite.execute({
      sql: `INSERT INTO turn_records (unit_turn_id, data, updated_at)
            VALUES (?, ?, datetime('now'))
            ON CONFLICT(unit_turn_id) DO UPDATE SET data = excluded.data, updated_at = datetime('now')`,
      args: [unitTurnId, JSON.stringify(body)],
    });
    return { ok: true, unit_turn_id: unitTurnId };
  }

  const result = await sqlite.execute({
    sql: `SELECT unit_turn_id, data, updated_at
          FROM turn_records
          ORDER BY datetime(COALESCE(updated_at, '1970-01-01')) DESC`,
  });
  const records = rowsAsObjects(result).map((row) => {
    const parsed = parseJson(row.data, {});
    return { unit_turn_id: asString(row.unit_turn_id), updated_at: asString(row.updated_at), ...parsed };
  });
  return { ok: true, records };
}

export async function handleTurnRecordStage(req: Request): Promise<any> {
  let body: any = {};
  try { body = await req.json(); } catch { body = {}; }
  const id = asString(body.id || body.unit_turn_id || body.turn_id);
  const stage = asString(body.stage);
  if (!id || !stage) return { ok: false, status: 400, error: "Missing id or stage" };

  const existing = await sqlite.execute({
    sql: `SELECT data FROM turn_records WHERE unit_turn_id = ? LIMIT 1`,
    args: [id],
  });
  const current = parseJson(rowsAsObjects(existing)[0]?.data, {});
  if (!current.stages || typeof current.stages !== "object") current.stages = {};
  current.stages[stage] = body.data || {};
  current.unit_turn_id = id;

  await sqlite.execute({
    sql: `INSERT INTO turn_records (unit_turn_id, data, updated_at)
          VALUES (?, ?, datetime('now'))
          ON CONFLICT(unit_turn_id) DO UPDATE SET data = excluded.data, updated_at = datetime('now')`,
    args: [id, JSON.stringify(current)],
  });
  return { ok: true, unit_turn_id: id, stage };
}

export async function handleUnitTurnsSync(req: Request): Promise<any> {
  let body: any = {};
  try { body = await req.json(); } catch { body = {}; }
  const records = Array.isArray(body.records) ? body.records : [];
  let upserted = 0;

  for (const record of records) {
    const turnKey = asString(record.turn_key);
    if (!turnKey) continue;
    const trackingUuid = asString(record.tracking_uuid) || await getTrackingUuidForTurnKey(turnKey) || makeTrackingUuid();

    await sqlite.execute({
      sql: `INSERT INTO unit_turn_tracker (
              tracking_uuid, tracking_code, turn_key, unit_turn_id, unit_id, property_id,
              unit_name, property_name, move_out_date, move_in_date, inspection_date,
              first_wo_date, estimate_requested_date, estimate_received_date, status,
              confidence_score, confidence_label, site_manager, source_flags, metadata,
              closed_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
            ON CONFLICT(turn_key) DO UPDATE SET
              tracking_code = excluded.tracking_code,
              unit_turn_id = excluded.unit_turn_id,
              unit_id = excluded.unit_id,
              property_id = excluded.property_id,
              unit_name = excluded.unit_name,
              property_name = excluded.property_name,
              move_out_date = excluded.move_out_date,
              move_in_date = excluded.move_in_date,
              inspection_date = excluded.inspection_date,
              first_wo_date = excluded.first_wo_date,
              estimate_requested_date = excluded.estimate_requested_date,
              estimate_received_date = excluded.estimate_received_date,
              status = excluded.status,
              confidence_score = excluded.confidence_score,
              confidence_label = excluded.confidence_label,
              site_manager = excluded.site_manager,
              source_flags = excluded.source_flags,
              metadata = excluded.metadata,
              closed_at = excluded.closed_at,
              updated_at = datetime('now')`,
      args: [
        trackingUuid,
        asString(record.tracking_code),
        turnKey,
        asString(record.unit_turn_id),
        asString(record.unit_id),
        asString(record.property_id),
        asString(record.unit_name),
        asString(record.property_name),
        asString(record.move_out_date),
        asString(record.move_in_date),
        asString(record.inspection_date),
        asString(record.first_wo_date),
        asString(record.estimate_requested_date),
        asString(record.estimate_received_date),
        asString(record.status) || "on_radar",
        toPositiveInt(record.confidence_score, 0),
        asString(record.confidence_label) || "low",
        asString(record.site_manager),
        JSON.stringify(record.source_flags || {}),
        JSON.stringify(record.metadata || {}),
        asString(record.closed_at),
      ],
    });

    const milestones = record.milestones && typeof record.milestones === "object"
      ? record.milestones
      : {};
    for (const milestoneKey of Object.keys(milestones)) {
      const milestone = milestones[milestoneKey] || {};
      await sqlite.execute({
        sql: `INSERT INTO unit_turn_milestones (tracking_uuid, milestone_key, milestone_date, source, notes, updated_at)
              VALUES (?, ?, ?, ?, ?, datetime('now'))
              ON CONFLICT(tracking_uuid, milestone_key) DO UPDATE SET
                milestone_date = excluded.milestone_date,
                source = excluded.source,
                notes = excluded.notes,
                updated_at = datetime('now')`,
        args: [
          trackingUuid,
          milestoneKey,
          asString(milestone.date),
          asString(milestone.source),
          asString(milestone.notes),
        ],
      });
    }

    if (record.replace_work_orders) {
      await sqlite.execute({
        sql: `UPDATE unit_turn_work_orders SET removed = 1, updated_at = datetime('now') WHERE tracking_uuid = ?`,
        args: [trackingUuid],
      });
    }

    const workOrders = Array.isArray(record.work_orders) ? record.work_orders : [];
    for (const workOrder of workOrders) {
      const woId = asString(workOrder.wo_id || workOrder.id);
      if (!woId) continue;
      await sqlite.execute({
        sql: `INSERT INTO unit_turn_work_orders (
                tracking_uuid, wo_id, wo_db_uuid, source, status, created_at, removed, updated_at
              ) VALUES (?, ?, ?, ?, ?, ?, 0, datetime('now'))
              ON CONFLICT(tracking_uuid, wo_id) DO UPDATE SET
                wo_db_uuid = excluded.wo_db_uuid,
                source = excluded.source,
                status = excluded.status,
                created_at = excluded.created_at,
                removed = 0,
                updated_at = datetime('now')`,
        args: [
          trackingUuid,
          woId,
          asString(workOrder.wo_db_uuid),
          asString(workOrder.source) || "manual",
          asString(workOrder.status),
          asString(workOrder.created_at),
        ],
      });
    }

    if (["closed", "completed"].includes(asString(record.status).toLowerCase()) || asString(record.closed_at)) {
      const metadata = record.metadata || {};
      await sqlite.execute({
        sql: `INSERT INTO closed_turns (
                turn_id, closed_at, close_reason, close_source, closed_by,
                property_id, property_name, unit_id, unit_name, move_out_date, move_in_date
              ) VALUES (?, datetime('now'), ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
          turnKey,
          asString(metadata.close_reason),
          asString(metadata.close_source),
          "sync",
          asString(record.property_id),
          asString(record.property_name),
          asString(record.unit_id),
          asString(record.unit_name),
          asString(record.move_out_date),
          asString(record.move_in_date),
        ],
      });
    }

    upserted += 1;
  }

  return { ok: true, upserted, count: upserted };
}

export async function handleUnitTurnWorkOrderLink(req: Request): Promise<any> {
  let body: any = {};
  try { body = await req.json(); } catch { body = {}; }
  const turnKey = asString(body.turn_key);
  const woId = asString(body.wo_id || body.id);
  if (!turnKey || !woId) return { ok: false, status: 400, error: "Missing turn_key or wo_id" };
  const trackingUuid = await getTrackingUuidForTurnKey(turnKey) || makeTrackingUuid();
  await sqlite.execute({
    sql: `INSERT INTO unit_turn_tracker (tracking_uuid, turn_key, status, updated_at)
          VALUES (?, ?, 'on_radar', datetime('now'))
          ON CONFLICT(turn_key) DO UPDATE SET updated_at = datetime('now')`,
    args: [trackingUuid, turnKey],
  });
  await sqlite.execute({
    sql: `INSERT INTO unit_turn_work_orders (tracking_uuid, wo_id, wo_db_uuid, source, status, created_at, removed, updated_at)
          VALUES (?, ?, ?, ?, ?, ?, 0, datetime('now'))
          ON CONFLICT(tracking_uuid, wo_id) DO UPDATE SET
            wo_db_uuid = excluded.wo_db_uuid,
            source = excluded.source,
            status = excluded.status,
            created_at = excluded.created_at,
            removed = 0,
            updated_at = datetime('now')`,
    args: [trackingUuid, woId, asString(body.wo_db_uuid), asString(body.source) || 'manual', asString(body.status), asString(body.created_at)],
  });
  return { ok: true, turn_key: turnKey, wo_id: woId };
}

export async function handleUnitTurnWorkOrderUnlink(req: Request): Promise<any> {
  let body: any = {};
  try { body = await req.json(); } catch { body = {}; }
  const turnKey = asString(body.turn_key);
  const woId = asString(body.wo_id || body.id);
  if (!turnKey || !woId) return { ok: false, status: 400, error: "Missing turn_key or wo_id" };
  const trackingUuid = await getTrackingUuidForTurnKey(turnKey);
  if (!trackingUuid) return { ok: false, status: 404, error: "Turn not found" };
  await sqlite.execute({
    sql: `UPDATE unit_turn_work_orders SET removed = 1, updated_at = datetime('now') WHERE tracking_uuid = ? AND wo_id = ?`,
    args: [trackingUuid, woId],
  });
  return { ok: true, turn_key: turnKey, wo_id: woId };
}