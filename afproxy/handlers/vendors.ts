// ============================================================================
// handlers/vendors.ts — Vendor directory (Reports API v2, cached 120 min).
//
// Fetches active vendors via the vendor_directory report.
// Used by: HandyManager Vendors tab, Tier 2 tech pool management.
//
// Note: Tier 2 "deep bench" techs may be modelled as Vendors in AppFolio
// rather than Users. If so, their assignment flows through WO VendorId
// (not AssignedUsers). The tech_grades table is the authoritative roster
// for the reassignment engine regardless of AppFolio model .
// ============================================================================

import { cacheGet, cacheSet, rowsAsObjects, sqlite, upsertVendorRows } from "../db.ts";
import { fetchReport } from "../lib/appfolio.ts";

export async function handleVendors(
  _params: Record<string, string>,
): Promise<any> {
  const cacheKey = "vendors";

  const cached = await cacheGet(cacheKey, "vendors");
  if (cached) {
    return {
      ok: true,
      results: cached.data,
      count: cached.record_count,
      cached_at: cached.cached_at,
      from_cache: true,
    };
  }

  const results = await fetchReport("vendor_directory", {
    vendor_visibility: "active",
  });

  await cacheSet(cacheKey, "vendors", results, results.length);
  upsertVendorRows(results).catch(() => {});
  return { ok: true, results, count: results.length, from_cache: false };
}

// ── handleVendorOverride ──────────────────────────────────────────────────────
// GET  ?action=vendor_override  → returns all overrides
// POST ?action=vendor_override  body: { vendor_id, category?, compliant? }
//   compliant: 1 = manually compliant, 0 = non-compliant, null = clear override
export async function handleVendorOverride(
  _params: Record<string, string>,
  req?: Request,
): Promise<any> {
  if (req?.method === "POST") {
    let body: any = {};
    try {
      body = await req.json();
    } catch {
      return { ok: false, error: "Invalid JSON body" };
    }

    const { vendor_id } = body;
    const hasCategory = Object.prototype.hasOwnProperty.call(body, "category");
    const hasTradeCategory = Object.prototype.hasOwnProperty.call(
      body,
      "trade_category",
    );
    const hasCompliant = Object.prototype.hasOwnProperty.call(
      body,
      "compliant",
    );
    const category = hasCategory
      ? (body.category === null || body.category === undefined
        ? null
        : String(body.category))
      : null;
    const tradeCategory = hasTradeCategory
      ? (body.trade_category === null || body.trade_category === undefined
        ? null
        : String(body.trade_category))
      : null;
    const compliant = hasCompliant
      ? (body.compliant !== null && body.compliant !== undefined
        ? Number(body.compliant)
        : null)
      : null;
    if (!vendor_id) return { ok: false, error: "vendor_id is required" };

    try {
      await sqlite.execute({
        sql:
          `INSERT INTO vendor_overrides (vendor_id, category, trade_category, compliant, updated_at)
              VALUES (?, ?, ?, ?, datetime('now'))
              ON CONFLICT(vendor_id) DO UPDATE SET
                category       = CASE WHEN ? = 1 THEN excluded.category ELSE vendor_overrides.category END,
                trade_category = CASE WHEN ? = 1 THEN excluded.trade_category ELSE vendor_overrides.trade_category END,
                compliant      = CASE WHEN ? = 1 THEN excluded.compliant ELSE vendor_overrides.compliant END,
                updated_at = datetime('now')`,
        args: [
          String(vendor_id),
          category,
          tradeCategory,
          compliant,
          hasCategory ? 1 : 0,
          hasTradeCategory ? 1 : 0,
          hasCompliant ? 1 : 0,
        ],
      });
      return { ok: true, vendor_id };
    } catch (e: any) {
      return { ok: false, error: e.message };
    }
  }

  // GET — return all overrides
  try {
    const rows = rowsAsObjects(
      await sqlite.execute(
        `SELECT vendor_id, category, trade_category, compliant, updated_at FROM vendor_overrides`,
      ),
    );
    return { ok: true, results: rows, count: rows.length };
  } catch (e: any) {
    return { ok: false, error: e.message };
  }
}