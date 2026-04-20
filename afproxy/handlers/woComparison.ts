// ============================================================================
// handlers/woComparison.ts — Work Order Comparison Report.
//
// Produces a merged, enriched work order dataset spanning both the Phoenix
// and Tucson property groups. Supports JSON and CSV download formats.
//
// Data pipeline:
//   Phase 1A: Fetch Phoenix WOs  (work_order report, filtered by group UUID)
//   Phase 1B: Fetch Tucson WOs   (work_order report, filtered by group UUID)
//   Phase 1C: Fetch property staff assignments → build site manager lookup map
//   Phase 2:  Client-side date guard (handles edge cases where AF API ignores
//             status_date_range_to on some account configurations)
//   Phase 3:  Merge, enrich (assignment type, site manager), sort, cache
//
// CSV output: streamed directly as a file download Response — bypasses
//             corsJson() in main.ts (caller must check instanceof Response).
//
// Staggering the rate at which requests are issued prevents network
// congestion and errors . A 250 ms delay is inserted between
// each Reports API v2 call in this handler.
//
// If an incorrect parameter value is passed, the returned AppFolio error
// codes are descriptive enough to identify the problem .
// The next page of results is returned in the next_page_path field .
// All list endpoints return 401 if credentials are invalid .
// ============================================================================

import { cacheGet, cacheSet } from "../db.ts";
import { CORS_HEADERS, FLR_GROUPS } from "../config.ts";
import { fetchReport } from "../lib/appfolio.ts";
import { delay, today } from "../lib/fetchUtils.ts";

// ── csvEscape ─────────────────────────────────────────────────────────────────
// RFC 4180-compliant CSV cell escaping.
// Wraps values containing commas, quotes, or newlines in double-quotes,
// and escapes internal double-quotes by doubling them.
function csvEscape(val: any): string {
  const s = String(val ?? "");
  if (
    s.includes(",") ||
    s.includes('"') ||
    s.includes("\n") ||
    s.includes("\r")
  ) {
    return `"${s.replace(/"/g, '""')}"`;
  }
  return s;
}

// ── Column definition ─────────────────────────────────────────────────────────
// Shared between buildWoComparisonCsv and the merged record map step.
// Order here determines column order in the CSV download.
const WO_COMPARISON_COLS: { key: string; label: string }[] = [
  { key: "property_group", label: "Property Group" },
  { key: "property_name", label: "Property" },
  { key: "site_manager", label: "Site Manager" },
  { key: "unit_name", label: "Unit" },
  { key: "work_order_number", label: "WO #" },
  { key: "status", label: "Status" },
  { key: "category", label: "Category" },
  { key: "work_order_issue", label: "Issue" },
  { key: "assignment_type", label: "Assignment Type" },
  { key: "assignee_name", label: "Assignee / Vendor" },
  { key: "vendor_trade", label: "Vendor Trade" },
  { key: "requesting_tenant", label: "Requesting Tenant" },
  { key: "created_at", label: "Created Date" },
  { key: "completed_at", label: "Completed Date" },
  { key: "follow_up_on", label: "Follow Up" },
  { key: "status_notes", label: "Status Notes" },
  { key: "unit_turn_category", label: "Turn Category" },
  { key: "maintenance_limit", label: "Maint. Limit" },
  { key: "description", label: "Description" },
];

// ── buildWoComparisonCsv ──────────────────────────────────────────────────────
// Converts the merged record array into a downloadable CSV Response.
// Content-Disposition header triggers a browser file download.
// CORS headers are included so the HandyManager frontend can receive the blob.
function buildWoComparisonCsv(
  records: any[],
  fromDate: string,
  toDate: string,
): Response {
  const header = WO_COMPARISON_COLS.map((c) => csvEscape(c.label)).join(",");
  const rows = records.map((r) =>
    WO_COMPARISON_COLS.map((c) => csvEscape(r[c.key])).join(",")
  );
  const csv = [header, ...rows].join("\r\n");
  const fname = `FLR_WO_Comparison_${fromDate}_to_${toDate}.csv`;

  return new Response(csv, {
    status: 200,
    headers: {
      ...CORS_HEADERS,
      "Content-Type": "text/csv; charset=utf-8",
      "Content-Disposition": `attachment; filename="${fname}"`,
    },
  });
}

// ── handleWoComparisonReport ──────────────────────────────────────────────────
// Main handler — fetch, enrich, merge, sort, and return.
//
// Query params:
//   from_date  ISO date string — default "2026-01-01"
//   to_date    ISO date string — default today()
//   group      "phoenix" | "tucson" | "all" — default "all"
//   format     "json" | "csv"               — default "json"
//   force      "1" | "true"                 — bypass cache
//
// Returns: JSON object normally, or a raw CSV Response when format=csv.
// The caller in main.ts must check `if (result instanceof Response) return result`.
export async function handleWoComparisonReport(
  params: Record<string, string>,
  _req?: Request,
): Promise<any> {
  const fromDate = params.from_date || new Date(Date.now() - 90 * 86400_000).toISOString().slice(0, 10);
  const toDate = params.to_date || today();
  const groupFilter = (params.group || "all").toLowerCase();
  const format = (params.format || "json").toLowerCase();
  const force = params.force === "1" || params.force === "true";

  const cacheKey = `wo_comparison_${fromDate}_${toDate}_${groupFilter}`;

  // ── Cache hit ───────────────────────────────────────────────────────────────
  if (!force) {
    const cached = await cacheGet(cacheKey, "wo_comparison");
    if (cached) {
      if (format === "csv") {
        return buildWoComparisonCsv(cached.data, fromDate, toDate);
      }
      return {
        ok: true,
        results: cached.data,
        count: cached.record_count,
        from_cache: true,
        cached_at: cached.cached_at,
        from_date: fromDate,
        to_date: toDate,
      };
    }
  }

  // ── Phase 1A & 1B: Fetch WOs per group ────────────────────────────────────
  // status_date: "0"          → filter the date range against WO Created Date
  // Omitting work_order_statuses → AppFolio returns ALL statuses
  //   (open + completed + canceled) so the comparison report is comprehensive.
  // Stagger requests 250 ms apart to prevent congestion .
  const woBase = {
    property_visibility: "all",
    status_date: "0",
    status_date_range_from: fromDate,
    status_date_range_to: toDate,
  };

  let phoenixWOs: any[] = [];
  let tucsonWOs: any[] = [];

  if (groupFilter === "all" || groupFilter === "phoenix") {
    phoenixWOs = await fetchReport("work_order", {
      ...woBase,
      properties: { property_groups_ids: [FLR_GROUPS.Phoenix] },
    });
    // Tag every record with its group at fetch time — avoids a costly post-join
    phoenixWOs = phoenixWOs.map((wo: any) => ({ ...wo, _group: "Phoenix" }));
    await delay(250); // rate-limit buffer between v2 calls 
  }

  if (groupFilter === "all" || groupFilter === "tucson") {
    tucsonWOs = await fetchReport("work_order", {
      ...woBase,
      properties: { property_groups_ids: [FLR_GROUPS.Tucson] },
    });
    tucsonWOs = tucsonWOs.map((wo: any) => ({ ...wo, _group: "Tucson" }));
    await delay(250);
  }

  // ── Phase 1C: Property Staff Assignments → site manager lookup map ─────────
  // Fetches the property_staff_assignments report to map property_id → manager.
  // Roles containing "manager" or "portfolio" are considered site managers.
  // First match per property wins (avoids duplicate manager entries).
  // The next page of paginated results is in the next_page_path field .
  await delay(250);
  const staffRows = await fetchReport("property_staff_assignments", {
    property_visibility: "active",
    properties: {
      property_groups_ids: [FLR_GROUPS.Phoenix, FLR_GROUPS.Tucson],
    },
  });

  const PM_KEYWORDS = ["manager", "portfolio"];
  const siteManagerMap: Record<string, string> = {};

  for (const row of staffRows) {
    const role = String(row.staff_role || "").toLowerCase();
    const isManager = PM_KEYWORDS.some((kw) => role.includes(kw));
    if (isManager && row.property_id != null) {
      const pid = String(row.property_id);
      if (!siteManagerMap[pid]) {
        siteManagerMap[pid] = String(row.staff_name || "—");
      }
    }
  }

  // ── Phase 2: Merge + client-side date guard ────────────────────────────────
  // The AppFolio Reports API v2 may not always honour status_date_range_to on
  // certain account configurations. A client-side guard ensures correctness.
  // If an incorrect parameter is passed, AppFolio error codes are descriptive
  // enough to identify the problem .
  const allWOs = [...phoenixWOs, ...tucsonWOs];
  const fromMs = new Date(fromDate).getTime();
  const toMs = new Date(`${toDate}T23:59:59Z`).getTime();

  const inRange = allWOs.filter((wo: any) => {
    const raw = wo.created_at || wo.created_date || "";
    if (!raw) return true; // include if no date — don't silently discard
    const ms = new Date(raw).getTime();
    return ms >= fromMs && ms <= toMs;
  });

  // ── Phase 3: Enrich, merge, sort ──────────────────────────────────────────
  const merged = inRange.map((wo: any) => {
    const hasVendor = !!(wo.vendor_name && String(wo.vendor_name).trim());
    const hasAssignee = !!(wo.assigned_to && String(wo.assigned_to).trim());
    const assignmentType = hasVendor
      ? "3rd Party Vendor"
      : hasAssignee
      ? "In-House Tech"
      : "Unassigned";

    const pid = String(wo.property_id ?? "");

    return {
      property_group: wo._group ?? "—",
      property_name: wo.property_name ?? "—",
      property_id: pid,
      site_manager: siteManagerMap[pid] || "Unassigned",
      unit_name: wo.unit_name ?? "—",
      work_order_number: wo.work_order_number ?? wo.number ?? "—",
      status: wo.status ?? "—",
      category: wo.category ?? "—",
      work_order_issue: wo.work_order_issue ?? wo.issue ?? "—",
      description: String(wo.description ?? "")
        .replace(/[\r\n]+/g, " ").trim() || "—",
      assignment_type: assignmentType,
      assignee_name: String(
        (hasVendor ? wo.vendor_name : wo.assigned_to) ?? "—",
      ),
      vendor_trade: wo.vendor_trade ?? "—",
      requesting_tenant: wo.requesting_tenant ?? "—",
      created_at: wo.created_at ?? "—",
      completed_at: wo.completed_at ?? "—",
      follow_up_on: wo.follow_up_on ?? "—",
      status_notes: String(wo.status_notes ?? "")
        .replace(/[\r\n]+/g, " ").trim() || "—",
      unit_turn_category: wo.unit_turn_category ?? "—",
      maintenance_limit: wo.maintenance_limit ?? "—",
    };
  });

  // Sort: Phoenix before Tucson → property name → status
  merged.sort((a: any, b: any) => {
    const g = a.property_group.localeCompare(b.property_group);
    if (g !== 0) return g;
    const p = a.property_name.localeCompare(b.property_name);
    if (p !== 0) return p;
    return a.status.localeCompare(b.status);
  });

  // ── Cache + return ─────────────────────────────────────────────────────────
  await cacheSet(cacheKey, "wo_comparison", merged, merged.length);

  if (format === "csv") {
    return buildWoComparisonCsv(merged, fromDate, toDate);
  }

  return {
    ok: true,
    results: merged,
    count: merged.length,
    from_cache: false,
    from_date: fromDate,
    to_date: toDate,
    phoenix_count: phoenixWOs.length,
    tucson_count: tucsonWOs.length,
    staff_rows_fetched: staffRows.length,
    site_managers_mapped: Object.keys(siteManagerMap).length,
    filtered_out: allWOs.length - inRange.length,
  };
}