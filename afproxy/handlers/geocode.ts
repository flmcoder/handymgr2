import { rowsAsObjects, sqlite } from "../db.ts";

let _geocodeCacheReady = false;

function normalizeAddress(value: string): string {
  return String(value || "").trim().replace(/\s+/g, " ").toLowerCase();
}

async function ensureGeocodeCache(): Promise<void> {
  if (_geocodeCacheReady) return;
  await sqlite.execute(`CREATE TABLE IF NOT EXISTS geocode_cache (
    address_key   TEXT PRIMARY KEY,
    query_address TEXT NOT NULL,
    lat           REAL,
    lon           REAL,
    display_name  TEXT,
    cached_at     TEXT NOT NULL DEFAULT (datetime('now'))
  )`);
  _geocodeCacheReady = true;
}

export async function handleGeocode(params: Record<string, string>): Promise<any> {
  await ensureGeocodeCache();
  const address = String(params.address || params.q || "").trim();
  if (!address) {
    return { ok: false, status: 400, error: "Missing address" };
  }

  const addressKey = normalizeAddress(address);
  try {
    const cached = await sqlite.execute({
      sql: `SELECT lat, lon, display_name FROM geocode_cache WHERE address_key = ? LIMIT 1`,
      args: [addressKey],
    });
    const cachedRow = rowsAsObjects(cached)[0] || null;
    if (cachedRow && cachedRow.lat != null && cachedRow.lon != null) {
      return {
        ok: true,
        lat: Number(cachedRow.lat),
        lon: Number(cachedRow.lon),
        display_name: String(cachedRow.display_name || address),
        cached: true,
      };
    }
  } catch (_) {
    // Cache miss path should remain available.
  }

  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), 10_000);
  try {
    const url = `https://nominatim.openstreetmap.org/search?format=jsonv2&limit=1&q=${encodeURIComponent(address)}`;
    const res = await fetch(url, {
      headers: {
        "Accept": "application/json",
        "User-Agent": "HandyManager/9.7.8 geocode",
      },
      signal: ctrl.signal,
    });
    if (!res.ok) {
      return { ok: false, status: res.status, error: `Geocode failed: HTTP ${res.status}` };
    }
    const rows = await res.json().catch(() => []);
    const first = Array.isArray(rows) ? rows[0] : null;
    const lat = Number(first?.lat);
    const lon = Number(first?.lon);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) {
      return { ok: false, status: 404, error: "No geocode result found" };
    }
    const displayName = String(first?.display_name || address).trim();
    try {
      await sqlite.execute({
        sql: `INSERT INTO geocode_cache (address_key, query_address, lat, lon, display_name, cached_at)
              VALUES (?, ?, ?, ?, ?, datetime('now'))
              ON CONFLICT(address_key) DO UPDATE SET
                query_address = excluded.query_address,
                lat = excluded.lat,
                lon = excluded.lon,
                display_name = excluded.display_name,
                cached_at = datetime('now')`,
        args: [addressKey, address, lat, lon, displayName],
      });
    } catch (_) {}
    return { ok: true, lat, lon, display_name: displayName, cached: false };
  } catch (err: any) {
    return {
      ok: false,
      status: 502,
      error: `Geocode request failed: ${String(err?.message || err || "unknown error")}`,
    };
  } finally {
    clearTimeout(timer);
  }
}
