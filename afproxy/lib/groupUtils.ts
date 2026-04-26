// ============================================================================
// lib/groupUtils.ts — Shared property-group resolution helpers.
// Used by bills, workOrders, turns, and any handler that must filter results
// to a specific property group.
// ============================================================================

import { cacheGet, rowsAsObjects, sqlite } from "../db.ts";

/**
 * Resolves the Set of AF property UUIDs that belong to the given property group.
 * Accepts params with `group_uuid` / `group_id` (UUID) or `group_name` (display name).
 * Returns null if no group filter is specified or the group cannot be resolved.
 */
export async function resolveGroupPropertyIds(
  params: Record<string, string>,
): Promise<Set<string> | null> {
  const groupUuid = String(params.group_uuid || params.group_id || "").trim();
  const groupName = String(params.group_name || params.group || "").trim()
    .toLowerCase();
  if (!groupUuid && !groupName) return null;

  const cached = await cacheGet("property_groups", "property_groups");
  const rows = Array.isArray(cached?.data) ? cached!.data : [];
  if (!rows.length) return null;

  const target = rows.find((g: any) => {
    const gid = String(
      g.Id || g.id || g.uuid || g.GroupUuid || g.group_uuid ||
        g.property_group_uuid || "",
    ).trim();
    const gname = String(g.Name || g.name || "").trim().toLowerCase();
    if (groupUuid && gid && gid === groupUuid) return true;
    if (groupName && gname && gname === groupName) return true;
    return false;
  });

  if (!target) return null;

  const set = new Set<string>();
  const addValue = (v: any) => {
    const id = String(v || "").trim();
    if (id) set.add(id);
  };

  const props = target.Properties || target.properties || [];
  if (Array.isArray(props)) {
    props.forEach((p: any) => addValue(p?.Id || p?.id || p?.PropertyId || p));
  }

  const propIds = target.PropertyIds || target.property_ids || [];
  if (Array.isArray(propIds)) propIds.forEach(addValue);

  if (set.size === 0) return null;

  // Normalize group members to AF property UUIDs where possible.
  // Some group payloads carry property_map_id values while bills carry property_id/PropertyId.
  const rawIds = Array.from(set);
  const placeholders = rawIds.map(() => "?").join(",");
  const canonical = new Set<string>(rawIds);

  const addCanonical = (v: unknown) => {
    const id = String(v || "").trim();
    if (id) canonical.add(id);
  };

  try {
    const refRows = rowsAsObjects(await sqlite.execute({
      sql:
        `SELECT property_id, property_map_id
         FROM property_reference
         WHERE property_map_id IN (${placeholders})
            OR property_id IN (${placeholders})`,
      args: [...rawIds, ...rawIds],
    }));
    refRows.forEach((row) => addCanonical(row.property_id || row.property_map_id));
  } catch {
    // Non-fatal: fallback to cached group IDs.
  }

  try {
    const pmRows = rowsAsObjects(await sqlite.execute({
      sql:
        `SELECT property_id, id
         FROM property_map
         WHERE id IN (${placeholders})
            OR property_id IN (${placeholders})`,
      args: [...rawIds, ...rawIds],
    }));
    pmRows.forEach((row) => addCanonical(row.property_id || row.id));
  } catch {
    // Non-fatal: fallback to ids already collected.
  }

  try {
    const args: string[] = [];
    let sql =
      `SELECT property_id, property_map_id
       FROM group_resolution_cache
       WHERE (property_map_id IN (${placeholders}) OR property_id IN (${placeholders}))`;
    args.push(...rawIds, ...rawIds);
    if (groupUuid) {
      sql += " AND property_group_id = ?";
      args.push(groupUuid);
    }
    const grcRows = rowsAsObjects(await sqlite.execute({ sql, args }));
    grcRows.forEach((row) => addCanonical(row.property_id || row.property_map_id));
  } catch {
    // Non-fatal: fallback to ids already collected.
  }

  return canonical.size > 0 ? canonical : null;
}

/**
 * Returns true if the given property ID is in the allowed set.
 * When allowedIds is null (no group filter), every record passes.
 */
export function propertyInScope(
  propertyId: string | null | undefined,
  allowedIds: Set<string> | null,
): boolean {
  if (!allowedIds) return true;
  if (!propertyId) return false;
  return allowedIds.has(propertyId.trim());
}

/**
 * Resolves a UnitId to its parent PropertyId by querying the units cache table.
 * Returns null when the unit is not found in the local cache.
 *
 * Intended for use when a work order / bill / turn record carries a UnitId but
 * no PropertyId — the resolved PropertyId can then be tested with propertyInScope.
 */
export async function unitToPropertyId(
  unitId: string | null | undefined,
): Promise<string | null> {
  if (!unitId) return null;
  const id = unitId.trim();
  if (!id) return null;
  try {
    const res = await sqlite.execute({
      sql: `SELECT property_id FROM units WHERE unit_id = ? LIMIT 1`,
      args: [id],
    });
    const rows = rowsAsObjects(res);
    const propId = rows[0]?.property_id;
    return propId ? String(propId) : null;
  } catch {
    return null;
  }
}
