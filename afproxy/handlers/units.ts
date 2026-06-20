// ============================================================================
// handlers/units.ts — AppFolio /api/v0/units cache handler.
// @ts-nocheck - Deno code, type checking disabled
//
// Fetches all units via the Database API v0 with full pagination.
// Uses LastUpdatedAtFrom = 5 years ago for the required filter.
// Cache TTL: 6 hours. Returns flat array of unit records from cache.
//
// Units are the critical link between UnitId (on work orders / turns / bills)
// and PropertyId (used for property group scoping and portfolio filters).
//
// Timeout: 90 000 ms — paginated endpoint (same bucket as bills/properties).
// ============================================================================

import { rowsAsObjects, sqlite, upsertUnits } from "../db.ts";
import { AF_DB, dbHeaders } from "../config.ts";
import { fetchWithTimeout } from "../lib/fetchUtils.ts";

const UNITS_CACHE_TTL_MS = 6 * 60 * 60 * 1000; // 6 hours
const UNITS_FETCH_TIMEOUT = 45_000;
const UNITS_PAGE_SIZE = 100;

// ── handleUnits ───────────────────────────────────────────────────────────────
// action=units
// Returns all cached unit records. Re-fetches from AppFolio when the cache is
// stale or empty.
export async function handleUnits(
  params: Record<string, string>,
): Promise<any> {
  const forceRefresh = String(params.refresh || "").toLowerCase() === "true";

  // Check cache staleness
  if (!forceRefresh) {
    try {
      const meta = await sqlite.execute(
        `SELECT MAX(cached_at) AS latest FROM units`,
      );
      const rows = rowsAsObjects(meta);
      const latest = rows[0]?.latest as string | null;
      if (latest) {
        const ageMs = Date.now() - new Date(latest).getTime();
        if (ageMs < UNITS_CACHE_TTL_MS) {
          return await _readFromCache();
        }
      }
    } catch {
      // Cache read failed — fall through to live fetch
    }
  }

  return await _fetchAndCache();
}

// ── handleUnitLookup ──────────────────────────────────────────────────────────
// action=unit_lookup  ?unit_id=<uuid>
// Returns the single cached unit record for the given UnitId.
export async function handleUnitLookup(
  params: Record<string, string>,
): Promise<any> {
  const unitId = String(params.unit_id || "").trim();
  if (!unitId) return { ok: false, error: "Missing unit_id parameter" };

  try {
    const res = await sqlite.execute({
      sql: `SELECT * FROM units WHERE unit_id = ? LIMIT 1`,
      args: [unitId],
    });
    const rows = rowsAsObjects(res);
    if (rows.length === 0) {
      return { ok: false, error: "Unit not found in cache", unit_id: unitId };
    }
    return { ok: true, unit: rows[0] };
  } catch (err: unknown) {
    return {
      ok: false,
      error: `unit_lookup error: ${err instanceof Error ? err.message : String(err)}`,
    };
  }
}

// ── _readFromCache ────────────────────────────────────────────────────────────
async function _readFromCache(): Promise<any> {
  try {
    const res = await sqlite.execute(
      `SELECT * FROM units ORDER BY property_id, name`,
    );
    const rows = rowsAsObjects(res);
    return {
      ok: true,
      results: rows,
      count: rows.length,
      from_cache: true,
    };
  } catch (err: unknown) {
    return {
      ok: false,
      error: `units cache read error: ${err instanceof Error ? err.message : String(err)}`,
    };
  }
}

// ── _fetchAndCache ────────────────────────────────────────────────────────────
async function _fetchAndCache(): Promise<any> {
  // LastUpdatedAtFrom = 5 years ago — required filter per API docs
  const fromDate = new Date();
  fromDate.setFullYear(fromDate.getFullYear() - 5);
  const fromISO = fromDate.toISOString().slice(0, 19) + "Z";

  const allUnits: any[] = [];
  let nextPath: string | null =
    `/api/v0/units?filters%5BLastUpdatedAtFrom%5D=${
      encodeURIComponent(fromISO)
    }&page%5Bsize%5D=${UNITS_PAGE_SIZE}`;

  try {
    while (nextPath) {
      const resp = await fetchWithTimeout(`${AF_DB}${nextPath}`, {
        headers: dbHeaders(),
      }, UNITS_FETCH_TIMEOUT);

      if (resp.status === 429) {
        const retryAfter = parseInt(
          resp.headers.get("Retry-After") || "60",
          10,
        );
        await new Promise((r) => setTimeout(r, retryAfter * 1000));
        // Retry once after back-off
        const retry = await fetchWithTimeout(`${AF_DB}${nextPath}`, {
          headers: dbHeaders(),
        }, UNITS_FETCH_TIMEOUT);
        if (!retry.ok) {
          return {
            ok: false,
            error: `AppFolio units fetch failed after 429: HTTP ${retry.status}`,
          };
        }
        const retryData = await retry.json();
        const retryRows: any[] = retryData.data || retryData.results || [];
        allUnits.push(...retryRows);
        nextPath = retryData.next_page_path || null;
        continue;
      }

      if (!resp.ok) {
        return {
          ok: false,
          error: `AppFolio units fetch failed: HTTP ${resp.status}`,
        };
      }

      const data = await resp.json();
      const pageRows: any[] = data.data || data.results || [];
      allUnits.push(...pageRows);
      nextPath = data.next_page_path || null;
    }

    if (allUnits.length > 0) {
      await upsertUnits(allUnits);
    }

    return {
      ok: true,
      results: allUnits,
      count: allUnits.length,
      from_cache: false,
    };
  } catch (err: unknown) {
    return {
      ok: false,
      error: `units fetch error: ${err instanceof Error ? err.message : String(err)}`,
    };
  }
}
