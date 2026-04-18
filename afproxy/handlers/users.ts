// ============================================================================
// handlers/users.ts - AppFolio users endpoint proxy.
//
// GET ?action=users&role=maintenance_tech&since=2020-01-01T00:00:00Z
//
// Returns normalized AppFolio user records. Default behavior filters to
// maintenance-tech users for dispatch/roster workflows.
// ============================================================================

import { cacheGet, cacheSet } from "../db.ts";
import { fetchDbApi } from "../lib/appfolio.ts";
import { daysAgo, snapDays } from "../lib/fetchUtils.ts";

function normalizeRole(value: unknown): string {
  return String(value || "").trim().toLowerCase();
}

function roleMatches(userRole: string, requestedRole: string): boolean {
  if (!requestedRole || requestedRole === "all") return true;

  const role = normalizeRole(userRole);
  if (!role) return false;

  if (requestedRole === "maintenance_tech") {
    return role.includes("maintenance") && role.includes("tech");
  }

  return role === requestedRole;
}

export async function handleUsers(
  params: Record<string, string>,
): Promise<any> {
  const requestedSince = String(
    params.since || params.last_updated_from || "2020-01-01T00:00:00Z",
  ).trim();
  const parsed = Date.parse(requestedSince);
  const ageMs = Number.isFinite(parsed) ? Math.max(0, Date.now() - parsed) :
    365 * 24 * 60 * 60 * 1000;
  const requestedDays = Math.max(1, Math.round(ageMs / (24 * 60 * 60 * 1000)));
  const snappedDays = snapDays(requestedDays, "users");
  const since = `${daysAgo(snappedDays)}T00:00:00Z`;
  const role = normalizeRole(params.role || "maintenance_tech");
  const includeInactive = String(params.include_inactive || "0") === "1";

  const cacheKey = `users_${since}_${role}_${
    includeInactive ? "all" : "active"
  }`;
  const cached = await cacheGet(cacheKey, "users");
  if (cached) {
    return {
      ok: true,
      count: cached.record_count,
      results: cached.data,
      from_cache: true,
      since,
      requested_since: requestedSince,
      snapped_days: snappedDays,
      role,
    };
  }

  try {
    const path = `/api/v0/users?filters[LastUpdatedAtFrom]=${
      encodeURIComponent(since)
    }&page[size]=200`;
    const rows = await fetchDbApi(path, 2000);

    const filtered = (rows || []).filter((u: any) => {
      const roleOk = roleMatches(u?.UserRole || u?.user_role || "", role);
      if (!roleOk) return false;
      if (includeInactive) return true;
      const activeFlag = String(u?.Active ?? u?.active ?? "").toLowerCase();
      return activeFlag !== "false" && activeFlag !== "0";
    }).map((u: any) => ({
      id: String(u?.Id || u?.id || "").trim(),
      firstName: String(u?.FirstName || u?.first_name || "").trim(),
      lastName: String(u?.LastName || u?.last_name || "").trim(),
      name: [
        u?.FirstName || u?.first_name || "",
        u?.LastName || u?.last_name || "",
      ]
        .map((x: unknown) => String(x || "").trim())
        .filter(Boolean)
        .join(" "),
      email: String(u?.Email || u?.email || "").trim(),
      userRole: String(u?.UserRole || u?.user_role || "").trim(),
      active: u?.Active ?? u?.active ?? true,
      lastUpdatedAt: String(u?.LastUpdatedAt || u?.last_updated_at || "")
        .trim(),
    })).filter((u: any) => !!u.id);

    await cacheSet(cacheKey, "users", filtered, filtered.length);

    return {
      ok: true,
      count: filtered.length,
      results: filtered,
      from_cache: false,
      since,
      requested_since: requestedSince,
      snapped_days: snappedDays,
      role,
    };
  } catch (err: any) {
    return {
      ok: false,
      error: err?.message || "Unable to fetch users",
      results: [],
      since,
      requested_since: requestedSince,
      snapped_days: snappedDays,
      role,
    };
  }
}