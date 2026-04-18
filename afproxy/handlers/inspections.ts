// ============================================================================
// handlers/inspections.ts — Unit inspections (Reports API v2, cached 60 min).
//
// Fetches unit inspection records from the unit_inspection report.
// Accepts either a from_date param (ISO date string) or a days param
// that is converted to a from_date by the daysAgo() helper.
// include_blank_inspection_date: "1" ensures units with no recorded
// inspection date are included in the results (matches v8.9).
// ============================================================================

import { cacheGet, cacheSet } from "../db.ts";
import { fetchReport } from "../lib/appfolio.ts";
import { daysAgo, snapDays } from "../lib/fetchUtils.ts";

function normalizeText(value: unknown): string {
  return String(value || "").trim().toLowerCase();
}

function makeInspectionKey(
  propertyId: unknown,
  unitId: unknown,
  propertyName: unknown,
  unitName: unknown,
): string {
  const pid = normalizeText(propertyId);
  const uid = normalizeText(unitId);
  if (pid && uid) return `${pid}|${uid}`;

  const pName = normalizeText(propertyName);
  const uName = normalizeText(unitName);
  if (pName && uName) return `${pName}|${uName}`;

  return "";
}

// Hard floor — this company did not use AppFolio before 2021.
// Any row with a recorded inspection date before this year is a
// data-migration artifact and must be excluded unconditionally.
const AF_EPOCH = new Date("2021-01-01T00:00:00Z");

function isWithinLookback(
  row: Record<string, unknown>,
  lookbackDays: number,
): boolean {
  const raw = String(row.last_inspection_date || "").trim();
  if (!raw) return true; // no recorded date = never inspected = include (overdue)

  const inspectedAt = new Date(raw);
  if (Number.isNaN(inspectedAt.getTime())) return true;

  // Reject pre-AppFolio data migration artifacts
  if (inspectedAt < AF_EPOCH) return false;

  const ageMs = Date.now() - inspectedAt.getTime();
  return ageMs <= lookbackDays * 24 * 60 * 60 * 1000;
}

function looksInactive(value: unknown): boolean {
  const v = String(value || "").trim().toLowerCase();
  if (!v) return false;
  return [
    "inactive",
    "in-active",
    "archived",
    "off market",
    "off_market",
    "not rentable",
    "not_rentable",
    "unavailable",
    "false",
    "0",
    "no",
  ].includes(v);
}

function isActiveInspectionRow(row: Record<string, unknown>): boolean {
  if (!row || typeof row !== "object") return false;

  if (looksInactive(row.property_status)) return false;
  if (looksInactive(row.unit_status)) return false;
  if (looksInactive(row.status)) return false;

  const rentable = String(row.rentable || "").trim().toLowerCase();
  if (
    ["false", "0", "no", "n", "not rentable", "inactive"].includes(rentable)
  ) {
    return false;
  }

  return true;
}

export async function handleInspections(
  params: Record<string, string>,
): Promise<any> {
  const lookbackDays = snapDays(parseInt(params.days || "180", 10), "inspections");
  const fromDate = params.from_date || daysAgo(lookbackDays);
  // Default active_only=1 — exclude inactive/off-market units
  const activeOnly = String(params.active_only || "1") === "1";
  const cacheKey = `inspections_${fromDate}_${activeOnly ? "active" : "all"}`;

  const cached = await cacheGet(cacheKey, "inspections");
  if (cached) {
    return {
      ok: true,
      results: cached.data,
      count: cached.record_count,
      cached_at: cached.cached_at,
      from_cache: true,
    };
  }

  const rawResults = await fetchReport("unit_inspection", {
    last_inspection_on_from: fromDate,
    include_blank_inspection_date: "1",
  });

  let activeOccupancyKeys: Set<string> | null = null;
  if (activeOnly) {
    const tenants = await fetchReport("tenant_directory", {
      tenant_visibility: "active",
      tenant_statuses: ["0", "4"],
      property_visibility: "active",
      columns: [
        "property_id",
        "property_name",
        "property",
        "unit_id",
        "unit",
        "move_out",
        "status",
      ],
    });

    const now = Date.now();
    activeOccupancyKeys = new Set(
      (tenants || [])
        .filter((row: Record<string, unknown>) => {
          const rawMoveOut = String(row.move_out || "").trim();
          if (!rawMoveOut) return true;
          const moveOutAt = new Date(rawMoveOut).getTime();
          return Number.isNaN(moveOutAt) || moveOutAt >= now;
        })
        .map((row: Record<string, unknown>) =>
          makeInspectionKey(
            row.property_id,
            row.unit_id,
            row.property_name || row.property,
            row.unit,
          )
        )
        .filter(Boolean),
    );
  }

  const results = (rawResults || []).filter((r: Record<string, unknown>) => {
    if (!isWithinLookback(r, lookbackDays)) return false;
    if (!activeOnly) return true;
    if (!isActiveInspectionRow(r || {})) return false;

    const rawMoveOut = String(r.move_out_date || r.move_out || "").trim();
    if (rawMoveOut) {
      const moveOutAt = new Date(rawMoveOut).getTime();
      if (!Number.isNaN(moveOutAt) && moveOutAt < Date.now()) return false;
    }

    const key = makeInspectionKey(
      r.property_id,
      r.unit_id,
      r.property_name || r.property,
      r.unit_name || r.unit,
    );
    if (!key) return false;
    return !!activeOccupancyKeys?.has(key);
  });

  await cacheSet(cacheKey, "inspections", results, results.length);
  return { ok: true, results, count: results.length, from_cache: false };
}