// ============================================================================
// handlers/properties.ts — Property, property group, and tenant data.
//
// Exports:
//   handleProperties       — property_directory report (cached 360 min)
//   handlePropertyGroups   — DB API v0 /property_groups (cached 360 min)
//   handlePropertyMap      — DB API v0 UUID → {name, address} map (cached 360 min)
//   handleUpcomingMoveouts — tenant_directory with move_out filtering (cached 30 min)
//   getTenantContact       — v9.0 helper: cross-reference phone + name for a unit
//
// DB API v0 endpoints require filters[LastUpdatedAtFrom] or filters[Id] —
// omitting a filter on list endpoints returns a 400 Bad Request .
// AppFolio property data includes Address1, SiteManager, and City fields .
// ============================================================================

import { cacheGet, cacheSet, rowsAsObjects, sqlite, upsertPropertyRows } from "../db.ts";
import { AF_DB, AF_REPORTS, dbHeaders } from "../config.ts";
import { fetchDbApi, fetchReport } from "../lib/appfolio.ts";
import { fetchWithTimeout } from "../lib/fetchUtils.ts";

async function readStaleCache(cacheKey: string): Promise<
  {
    data: any;
    cached_at: string;
    record_count: number;
  } | null
> {
  try {
    const result = await sqlite.execute({
      sql: `SELECT data, cached_at, record_count
            FROM api_cache
            WHERE cache_key = ?
            ORDER BY cached_at DESC
            LIMIT 1`,
      args: [cacheKey],
    });
    const row = rowsAsObjects(result)[0] as any;
    if (!row || !row.data) return null;

    const meta = JSON.parse(String(row.data || "null"));
    if (meta && typeof meta._chunks === "number") {
      let fullJson = "";
      for (let i = 0; i < Number(meta._chunks); i++) {
        const chunkResult = await sqlite.execute({
          sql: `SELECT data FROM api_cache WHERE cache_key = ?`,
          args: [`${cacheKey}::${i}`],
        });
        const chunkRow = rowsAsObjects(chunkResult)[0] as any;
        if (!chunkRow || chunkRow.data === undefined) return null;
        fullJson += String(chunkRow.data || "");
      }
      return {
        data: JSON.parse(fullJson || "[]"),
        cached_at: String(row.cached_at || ""),
        record_count: Number(row.record_count || 0),
      };
    }

    return {
      data: meta,
      cached_at: String(row.cached_at || ""),
      record_count: Number(row.record_count || 0),
    };
  } catch {
    return null;
  }
}

// ── handleProperties ──────────────────────────────────────────────────────────
// Active properties via property_directory report (Reports API v2).
// Cached 360 minutes — properties rarely change mid-day.
export async function handleProperties(
  _params: Record<string, string>,
): Promise<any> {
  const cacheKey = "properties";

  const cached = await cacheGet(cacheKey, "properties");
  if (cached) {
    return {
      ok: true,
      results: cached.data,
      count: cached.record_count,
      cached_at: cached.cached_at,
      from_cache: true,
    };
  }

  let results: any[] = [];
  try {
    results = await fetchReport("property_directory", {
      property_visibility: "active",
    });
  } catch (fetchErr: any) {
    const stale = await readStaleCache(cacheKey);
    if (stale && Array.isArray(stale.data)) {
      return {
        ok: true,
        results: stale.data,
        count: stale.record_count || stale.data.length,
        cached_at: stale.cached_at,
        from_cache: true,
        stale_cache: true,
        warning: `properties fetch error: ${fetchErr.message}`,
      };
    }
    return {
      ok: false,
      error: `properties fetch error: ${fetchErr.message}`,
    };
  }

  await cacheSet(cacheKey, "properties", results, results.length);
  return { ok: true, results, count: results.length, from_cache: false };
}

// ── handlePropertyGroups ──────────────────────────────────────────────────────
// Fetches property groups via DB API v0.
// Required filter: LastUpdatedAtFrom (AF DB v0 list endpoints require a filter).
// Domain fallback: if api.appfolio.com returns 401, retries on the tenant domain.
// Paginates up to 20 pages (2000 groups maximum).
export async function handlePropertyGroups(
  params: Record<string, string>,
): Promise<any> {
  const cacheKey = "property_groups";

  const cached = await cacheGet(cacheKey, "property_groups");
  if (cached) {
    return {
      ok: true,
      results: cached.data,
      count: cached.record_count,
      cached_at: cached.cached_at,
      from_cache: true,
    };
  }

  const lastFrom = params["filters[LastUpdatedAtFrom]"] ||
    "2024-01-01T00:00:00Z";
  // IMPORTANT (2026-04-02): Keep this exact query shape for AppFolio DB API v0.
  // Known-good command:
  // curl -X GET "https://api.appfolio.com/api/v0/property_groups?filters%5BLastUpdatedAtFrom%5D=2024-01-01T00:00:00Z&page%5Bsize%5D=50"
  // Do not change encoded bracket keys (%5B / %5D) or default page size (50)
  // unless the integration is re-validated end-to-end, because this has caused
  // property group rendering regressions in production.
  const pgSize = "50";
  const pgPath = `/api/v0/property_groups?filters%5BLastUpdatedAtFrom%5D=${
    encodeURIComponent(lastFrom)
  }&page%5Bsize%5D=${pgSize}`;

  let pgResp: Response;
  let domainUsed = AF_DB;
  try {
    pgResp = await fetchWithTimeout(`${AF_DB}${pgPath}`, {
      headers: dbHeaders(),
    });
    if ([401, 403, 404, 422].includes(pgResp.status)) {
      domainUsed = AF_REPORTS;
      pgResp = await fetchWithTimeout(`${AF_REPORTS}${pgPath}`, {
        headers: dbHeaders(),
      });
    }
  } catch (fetchErr: any) {
    // AbortError (timeout) or network failure — serve stale cache if available
    const stale = await readStaleCache(cacheKey);
    if (stale && Array.isArray(stale.data)) {
      return {
        ok: true,
        results: stale.data,
        count: stale.record_count || stale.data.length,
        cached_at: stale.cached_at,
        from_cache: true,
        stale_cache: true,
        warning: `property_groups fetch error: ${fetchErr.message}`,
      };
    }
    return {
      ok: false,
      error: `property_groups fetch error: ${fetchErr.message}`,
      domain: domainUsed,
    };
  }

  if (!pgResp.ok) {
    const errBody = await pgResp.text().catch(() => "");
    const stale = await readStaleCache(cacheKey);
    if (stale && Array.isArray(stale.data)) {
      return {
        ok: true,
        results: stale.data,
        count: stale.record_count || stale.data.length,
        cached_at: stale.cached_at,
        from_cache: true,
        stale_cache: true,
        warning: `property_groups upstream failed: HTTP ${pgResp.status}`,
      };
    }
    return {
      ok: false,
      error: `property_groups fetch failed: HTTP ${pgResp.status}`,
      status: pgResp.status,
      domain: domainUsed,
      detail: errBody.substring(0, 500),
    };
  }

  const pgData = await pgResp.json();
  let allGroups = pgData.results || pgData.data ||
    (Array.isArray(pgData) ? pgData : []);

  // Paginate up to 20 pages (matches v8.9 — 2000 groups max).
  let nextPage: string | null = pgData.next_page_path || null;
  let page = 1;
  while (nextPage && page < 20) {
    const fullUrl = nextPage.startsWith("http")
      ? nextPage
      : `${domainUsed}${nextPage}`;
    const nr = await fetchWithTimeout(fullUrl, { headers: dbHeaders() });
    if (!nr.ok) break;
    const nd = await nr.json();
    const nResults = nd.results || nd.data || [];
    allGroups = allGroups.concat(nResults);
    nextPage = nd.next_page_path || null;
    if (nResults.length === 0) break;
    page++;
  }

  if (allGroups.length > 0) {
    await cacheSet(cacheKey, "property_groups", allGroups, allGroups.length);
  }
  return {
    ok: true,
    results: allGroups,
    count: allGroups.length,
    domain: domainUsed,
    from_cache: false,
  };
}

// ── handlePropertyMap ─────────────────────────────────────────────────────────
// Builds a UUID → { name, address } map from DB API v0 /properties.
// Used by the HandyManager cockpit to resolve property UUIDs from webhook
// payloads to human-readable names and addresses .
// Cached 360 minutes. Paginates up to 20 pages (cap 2000 properties).
export async function handlePropertyMap(
  _params: Record<string, string>,
): Promise<any> {
  const cacheKey = "property_map";

  const cached = await cacheGet(cacheKey, "property_map");
  if (cached) {
    return {
      ok: true,
      property_uuid_map: cached.data,
      cached_at: cached.cached_at,
      from_cache: true,
    };
  }

  const pmPath =
    "/api/v0/properties?filters%5BLastUpdatedAtFrom%5D=2024-01-01T00%3A00%3A00Z&page%5Bsize%5D=200";

  let pmResp: Response;
  let pmDomain = AF_DB;
  try {
    pmResp = await fetchWithTimeout(`${AF_DB}${pmPath}`, {
      headers: dbHeaders(),
    });
    if ([401, 403, 404, 422].includes(pmResp.status)) {
      pmDomain = AF_REPORTS;
      pmResp = await fetchWithTimeout(`${AF_REPORTS}${pmPath}`, {
        headers: dbHeaders(),
      });
    }
  } catch (fetchErr: any) {
    const stale = await readStaleCache(cacheKey);
    if (stale && stale.data && typeof stale.data === "object") {
      const fallbackCount = stale.record_count ||
        Object.keys(stale.data).length;
      return {
        ok: true,
        property_uuid_map: stale.data,
        count: fallbackCount,
        cached_at: stale.cached_at,
        from_cache: true,
        stale_cache: true,
        warning: `property_map fetch error: ${fetchErr.message}`,
      };
    }
    return {
      ok: false,
      error: `property_map fetch error: ${fetchErr.message}`,
      domain: pmDomain,
    };
  }

  if (!pmResp.ok) {
    const errBody = await pmResp.text().catch(() => "");
    const stale = await readStaleCache(cacheKey);
    if (stale && stale.data && typeof stale.data === "object") {
      const fallbackCount = stale.record_count ||
        Object.keys(stale.data).length;
      return {
        ok: true,
        property_uuid_map: stale.data,
        count: fallbackCount,
        cached_at: stale.cached_at,
        from_cache: true,
        stale_cache: true,
        warning: `property_map upstream failed: HTTP ${pmResp.status}`,
      };
    }
    return {
      ok: false,
      error: `property_map fetch failed: HTTP ${pmResp.status}`,
      domain: pmDomain,
      detail: errBody.substring(0, 500),
    };
  }

  const pmData = await pmResp.json();
  let dbProps = pmData.results || pmData.data ||
    (Array.isArray(pmData) ? pmData : []);

  // Paginate — cap at 2000 properties
  let pmNext: string | null = pmData.next_page_path || null;
  let pmPage = 1;
  while (pmNext && dbProps.length < 2000 && pmPage < 20) {
    const fu = pmNext.startsWith("http") ? pmNext : `${pmDomain}${pmNext}`;
    const nr = await fetchWithTimeout(fu, { headers: dbHeaders() });
    if (!nr.ok) break;
    const nd = await nr.json();
    const nR = nd.results || nd.data || [];
    dbProps = dbProps.concat(nR);
    pmNext = nd.next_page_path || null;
    if (nR.length === 0) break;
    pmPage++;
  }

  // Build UUID → property metadata map.
  // Includes extra fields used by frontend turn-board and WO detail enrichment.
  const propMap: Record<string, {
    name: string;
    address: string;
    site_manager_name?: string;
    group_ids?: string[];
    maintenance_notes?: string;
  }> = {};
  dbProps.forEach((p: any) => {
    if (p.Id) {
      let siteManagerName = "";
      const sm = p.SiteManager || p.site_manager || null;
      if (sm && typeof sm === "object") {
        siteManagerName = [sm.FirstName, sm.LastName].filter(Boolean).join(" ")
          .trim();
      } else if (typeof sm === "string") {
        siteManagerName = sm.trim();
      }

      let groupIds: string[] = [];
      const rawGroupIds = p.PropertyGroupIds || p.property_group_ids || [];
      if (Array.isArray(rawGroupIds)) {
        groupIds = rawGroupIds.map((v: any) => String(v || "").trim()).filter(
          Boolean,
        );
      } else if (typeof rawGroupIds === "string" && rawGroupIds.trim()) {
        try {
          const parsed = JSON.parse(rawGroupIds);
          if (Array.isArray(parsed)) {
            groupIds = parsed.map((v: any) => String(v || "").trim()).filter(
              Boolean,
            );
          }
        } catch {
          // Non-fatal: keep empty groupIds when value is not parseable.
        }
      }
      // Fallback: singular PropertyGroupId (AppFolio v0 returns this form)
      if (groupIds.length === 0) {
        const singular = String(p.PropertyGroupId || p.property_group_id || "").trim();
        if (singular) groupIds = [singular];
      }

      propMap[p.Id] = {
        name: p.Name || p.PropertyName || "",
        address: p.Address1 || p.StreetAddress || "",
        site_manager_name: siteManagerName || "",
        group_ids: groupIds,
        maintenance_notes: p.MaintenanceNotes || p.maintenance_notes || "",
      };
    }
  });

  await cacheSet(
    cacheKey,
    "property_map",
    propMap,
    Object.keys(propMap).length,
  );
  upsertPropertyRows(dbProps).catch(() => {});
  return {
    ok: true,
    property_uuid_map: propMap,
    count: Object.keys(propMap).length,
    domain: pmDomain,
    from_cache: false,
  };
}

// ── handleUpcomingMoveouts ────────────────────────────────────────────────────
// Fetches tenant directory with move_out date filtering.
// Returns tenants with move-out dates within the next `days` days,
// plus a 14-day lookback window for recently moved-out tenants.
// Includes phone_numbers column — used by getTenantContact() below.
export async function handleUpcomingMoveouts(
  params: Record<string, string>,
): Promise<any> {
  const days = parseInt(params.days || "60", 10);
  const cacheKey = `upcoming_moveouts_${days}`;

  const cached = await cacheGet(cacheKey, "upcoming_moveouts");
  if (cached) {
    return {
      ok: true,
      results: cached.data,
      count: cached.record_count,
      cached_at: cached.cached_at,
      from_cache: true,
    };
  }

  const rows = await fetchReport("tenant_directory", {
    tenant_visibility: "active",
    tenant_statuses: ["0", "4"], // 0 = Current, 4 = Notice
    property_visibility: "active",
    columns: [
      "property",
      "property_name",
      "property_id",
      "unit",
      "unit_id",
      "tenant",
      "status",
      "move_out",
      "move_in",
      "phone_numbers",
      "emails",
      "occupancy_id",
      "rent",
    ],
  });

  // Post-fetch date filter — keep tenants in the lookback + forward window
  const now = new Date();
  const cutoffFuture = new Date();
  cutoffFuture.setDate(cutoffFuture.getDate() + days);
  const cutoffPast = new Date();
  cutoffPast.setDate(cutoffPast.getDate() - 14); // include recent move-outs

  const filtered = rows.filter((r: any) => {
    if (!r.move_out) return false;
    const moDate = new Date(r.move_out);
    return moDate >= cutoffPast && moDate <= cutoffFuture;
  });

  await cacheSet(cacheKey, "upcoming_moveouts", filtered, filtered.length);
  return {
    ok: true,
    results: filtered,
    count: filtered.length,
    from_cache: false,
  };
}

// ── getTenantContact ──────────────────────────────────────────────────────────
// v9.0 helper — resolves tenant phone + name for a property+unit combination.
//
// Used by:
//   • handleNoonWarningCron   — populates magic link tenant_phone
//   • handleMidnightReassignCron — populates new tech's magic link
//   • executeTier2Blast        — stores tenant info in blast_events for later
//
// Cross-references the upcoming_moveouts cache (120-day window) using
// fuzzy substring matching on property name and unit string.
// Returns { phone: "", name: "" } on any failure — callers handle gracefully.
// The phone_numbers field from the tenant_directory report is the authoritative
// source for tenant contact numbers.
export async function getTenantContact(
  propertyRef: string,
  unitRef: string,
): Promise<{ phone: string; name: string }> {
  try {
    const mo = await handleUpcomingMoveouts({ days: "120" });
    if (!mo.ok || !Array.isArray(mo.results)) return { phone: "", name: "" };

    const pLow = propertyRef.toLowerCase();
    const uLow = unitRef.toLowerCase();

    const match = mo.results.find((t: any) => {
      const tp = String(t.property_name || t.property || "").toLowerCase();
      const tu = String(t.unit || "").toLowerCase();
      return tp.includes(pLow) && (uLow ? tu.includes(uLow) : true);
    });

    if (!match) return { phone: "", name: "" };

    // phone_numbers may be a comma-separated string or an array
    const phones = match.phone_numbers || "";
    const raw = Array.isArray(phones)
      ? String(phones[0] || "")
      : String(phones).split(",")[0].trim();

    return {
      phone: raw,
      name: String(match.tenant || match.tenant_name || ""),
    };
  } catch {
    return { phone: "", name: "" };
  }
}