// ============================================================================
// lib/groupUtils.ts — Shared property-group resolution helpers.
// Used by bills, workOrders, turns, and any handler that must filter results
// to a specific property group.
// ============================================================================

import { cacheGet } from "../db.ts";

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

  return set.size > 0 ? set : null;
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
