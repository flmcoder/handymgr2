// ============================================================================
// handlers/estimates.ts — Flatten estimate data from cached work orders.
//
// Estimates are embedded in work-order payloads (v0 `Estimates` array) or in
// v2 report fields like `current_estimate_approval_status`. This handler reads
// the Turso cache table `work_orders_cache`, extracts both patterns, and
// returns a single flat list for the frontend Estimates table.
// ============================================================================

import { rowsAsObjects, sqlite } from "../db.ts";

type EstimateRow = {
  estimate_id: string;
  work_order_number: string;
  work_order_id: string;
  property_unit_address: string;
  vendor_name: string;
  estimate_amount: number | null;
  approval_status: string;
  source: "v0_estimates" | "v2_current_status";
  property_group_id: string;
  created_at: string;
  raw_data: Record<string, any>;
  status_history?: Array<{ status: string; changed_at: string }>;
};

const ALLOWED_STATUSES = new Set(["pending", "approved", "rejected"]);
let _estimatesTableReady = false;

function normalizeStatus(value: unknown): string {
  const raw = String(value || "").trim().toLowerCase();
  if (!raw) return "Pending";
  if (raw.includes("approve")) return "Approved";
  if (raw.includes("reject") || raw.includes("declin")) return "Rejected";
  if (raw.includes("pending") || raw.includes("request")) return "Pending";
  // Unknown statuses still represent an estimate lifecycle state; keep them visible.
  return "Pending";
}

function parseHistory(raw: unknown): Array<{ status: string; changed_at: string }> {
  if (!raw) return [];
  try {
    const parsed = typeof raw === "string" ? JSON.parse(raw) : raw;
    if (!Array.isArray(parsed)) return [];
    return parsed.map((entry) => ({
      status: String(entry?.status || "").trim(),
      changed_at: String(entry?.changed_at || "").trim(),
    })).filter((entry) => entry.status);
  } catch {
    return [];
  }
}

async function ensureEstimatesTable(): Promise<void> {
  if (_estimatesTableReady) return;
  await sqlite.execute(`CREATE TABLE IF NOT EXISTS estimates (
    estimate_id        TEXT PRIMARY KEY,
    work_order_id      TEXT,
    work_order_number  TEXT,
    current_status     TEXT NOT NULL,
    property_group_id  TEXT,
    source             TEXT,
    status_history_json TEXT NOT NULL DEFAULT '[]',
    raw_data           TEXT NOT NULL DEFAULT '{}',
    created_at         TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at         TEXT NOT NULL DEFAULT (datetime('now'))
  )`);
  try {
    await sqlite.execute(`CREATE INDEX IF NOT EXISTS idx_estimates_status ON estimates(current_status)`);
  } catch (_) {}
  try {
    await sqlite.execute(`CREATE INDEX IF NOT EXISTS idx_estimates_group ON estimates(property_group_id)`);
  } catch (_) {}
  try {
    await sqlite.execute(`CREATE INDEX IF NOT EXISTS idx_estimates_wo ON estimates(work_order_id)`);
  } catch (_) {}
  _estimatesTableReady = true;
}

function toNumber(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  const parsed = Number(raw.replace(/[^0-9.\-]/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
}

function pickFirst(obj: Record<string, any>, keys: string[]): any {
  for (const k of keys) {
    if (obj[k] !== undefined && obj[k] !== null && String(obj[k]).trim() !== "") {
      return obj[k];
    }
  }
  return "";
}

function buildAddress(row: Record<string, any>, raw: Record<string, any>): string {
  const propertyName = String(
    pickFirst(row, ["property_name", "property", "propertyName"]) ||
      pickFirst(raw, ["PropertyName", "property_name", "property"]) || "",
  ).trim();
  const unit = String(
    pickFirst(row, ["unit_number", "unit_name", "unit", "unitNumber"]) ||
      pickFirst(raw, ["UnitName", "unit_name", "unit"]) || "",
  ).trim();
  const street = String(
    pickFirst(row, ["address"]) ||
      [
        raw.PropertyStreet,
        raw.property_street,
        raw.PropertyCity,
        raw.property_city,
        raw.PropertyState,
        raw.property_state,
        raw.PropertyZip,
        raw.property_zip,
      ].filter(Boolean).join(" ") ||
      "",
  ).trim();

  const base = [propertyName, unit ? `Unit ${unit}` : ""].filter(Boolean).join(" • ");
  if (!street) return base || "—";
  return [base, street].filter(Boolean).join(" • ");
}

function flattenFromWorkOrderRow(row: Record<string, any>): EstimateRow[] {
  const out: EstimateRow[] = [];
  const raw = (() => {
    try {
      return JSON.parse(String(row.raw_json || "{}"));
    } catch {
      return {};
    }
  })();

  const workOrderNumber = String(
    pickFirst(row, ["wo_number", "work_order_number"]) ||
      pickFirst(raw, ["WorkOrderNumber", "work_order_number", "service_request_number"]) ||
      "",
  ).trim();
  const workOrderId = String(
    pickFirst(row, ["id"]) ||
      pickFirst(raw, ["Id", "work_order_id", "uuid"]) ||
      "",
  ).trim();
  const vendorName = String(
    pickFirst(row, ["assigned_vendor"]) ||
      pickFirst(raw, ["VendorName", "vendor", "vendor_name"]) ||
      "",
  ).trim();
  const propertyGroupId = String(pickFirst(row, ["property_group_id"]) || "").trim();
  const resolvedPropertyGroupId = String(
    propertyGroupId ||
      pickFirst(row, ["property_group_uuid", "uuid_prop_group", "group_uuid"]) ||
      pickFirst(raw, [
        "property_group_id",
        "property_group_uuid",
        "uuid_prop_group",
        "group_uuid",
      ]) ||
      "",
  ).trim();
  const createdAt = String(
    pickFirst(row, ["updated_at", "created_date"]) ||
      pickFirst(raw, ["UpdatedAt", "CreatedAt", "created_at"]) ||
      "",
  ).trim();

  const address = buildAddress(row, raw);

  const estimateArray = Array.isArray(raw.Estimates)
    ? raw.Estimates
    : (Array.isArray(raw.estimates) ? raw.estimates : []);

  if (estimateArray.length) {
    for (let index = 0; index < estimateArray.length; index += 1) {
      const item = estimateArray[index];
      const itemObj = item && typeof item === "object" ? item as Record<string, any> : {};
      const status = normalizeStatus(
        pickFirst(itemObj, [
          "approval_status",
          "ApprovalStatus",
          "status",
          "estimate_approval_status",
        ]) || pickFirst(raw, ["current_estimate_approval_status"]),
      );
      const estimateId = String(
        pickFirst(itemObj, ["estimate_id", "EstimateId", "id", "Id", "uuid"]) ||
          `${workOrderId || workOrderNumber || "estimate"}:v0:${index}`,
      ).trim();
      out.push({
        estimate_id: estimateId,
        work_order_number: workOrderNumber,
        work_order_id: workOrderId,
        property_unit_address: address,
        vendor_name: String(
          pickFirst(itemObj, ["vendor_name", "VendorName", "vendor"]) || vendorName,
        ).trim(),
        estimate_amount: toNumber(
          pickFirst(itemObj, ["amount", "Amount", "estimated_amount", "EstimateAmount"]) ||
            pickFirst(row, ["estimated_amount"]),
        ),
        approval_status: status,
        source: "v0_estimates",
        property_group_id: resolvedPropertyGroupId,
        created_at: createdAt,
        raw_data: {
          work_order_id: workOrderId,
          work_order_number: workOrderNumber,
          source: "v0_estimates",
          estimate: itemObj,
        },
      });
    }
    return out;
  }

  const v2StatusRaw = pickFirst(raw, [
    "current_estimate_approval_status",
    "CurrentEstimateApprovalStatus",
    "estimate_approval_status",
  ]);
  const fallbackStatusRaw = pickFirst(row, ["status"]);
  const normalizedV2 = normalizeStatus(v2StatusRaw || fallbackStatusRaw);

  if (String(v2StatusRaw || fallbackStatusRaw || "").trim()) {
    out.push({
      estimate_id: `${workOrderId || workOrderNumber || "estimate"}:v2_current_status`,
      work_order_number: workOrderNumber,
      work_order_id: workOrderId,
      property_unit_address: address,
      vendor_name: vendorName,
      estimate_amount: toNumber(pickFirst(row, ["estimated_amount"]) || pickFirst(raw, ["EstimatedAmount", "estimated_amount"])),
      approval_status: normalizedV2,
      source: "v2_current_status",
      property_group_id: resolvedPropertyGroupId,
      created_at: createdAt,
      raw_data: {
        work_order_id: workOrderId,
        work_order_number: workOrderNumber,
        source: "v2_current_status",
        current_estimate_approval_status: v2StatusRaw || fallbackStatusRaw,
        estimated_amount: pickFirst(row, ["estimated_amount"]) || pickFirst(raw, ["EstimatedAmount", "estimated_amount"]),
      },
    });
  }

  return out;
}

async function syncEstimateHistory(rows: EstimateRow[]): Promise<Map<string, Array<{ status: string; changed_at: string }>>> {
  await ensureEstimatesTable();
  const historyById = new Map<string, Array<{ status: string; changed_at: string }>>();
  const nowIso = new Date().toISOString();

  for (const row of rows) {
    const estimateId = String(row.estimate_id || "").trim();
    if (!estimateId) continue;

    const existingResult = await sqlite.execute({
      sql: `SELECT current_status, status_history_json FROM estimates WHERE estimate_id = ? LIMIT 1`,
      args: [estimateId],
    });
    const existing = rowsAsObjects(existingResult)[0] || null;
    const history = parseHistory(existing?.status_history_json);
    if (!history.length || String(existing?.current_status || "").trim() !== row.approval_status) {
      history.push({ status: row.approval_status, changed_at: nowIso });
    }

    historyById.set(estimateId, history);
    await sqlite.execute({
      sql: `INSERT INTO estimates (
              estimate_id,
              work_order_id,
              work_order_number,
              current_status,
              property_group_id,
              source,
              status_history_json,
              raw_data,
              updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
            ON CONFLICT(estimate_id) DO UPDATE SET
              work_order_id = excluded.work_order_id,
              work_order_number = excluded.work_order_number,
              current_status = excluded.current_status,
              property_group_id = excluded.property_group_id,
              source = excluded.source,
              status_history_json = excluded.status_history_json,
              raw_data = excluded.raw_data,
              updated_at = datetime('now')`,
      args: [
        estimateId,
        row.work_order_id,
        row.work_order_number,
        row.approval_status,
        row.property_group_id,
        row.source,
        JSON.stringify(history),
        JSON.stringify(row.raw_data || {}),
      ],
    });
  }

  return historyById;
}

export async function handleEstimates(params: Record<string, string>): Promise<any> {
  const statusesRaw = String(params.statuses || "").trim();
  const wanted = new Set(
    statusesRaw
      ? statusesRaw.split(",").map((s) => s.trim().toLowerCase()).filter(Boolean)
      : Array.from(ALLOWED_STATUSES),
  );

  const selectSql = `SELECT
              id,
              wo_number,
              property_name,
              unit_number,
              address,
              assigned_vendor,
              estimated_amount,
              property_group_id,
              created_date,
              updated_at,
              raw_json,
              status
            FROM work_orders_cache`;

  const orderSql = `
            ORDER BY datetime(COALESCE(updated_at, created_date, '1970-01-01T00:00:00Z')) DESC`;

  try {
    await ensureEstimatesTable();
    let result: any;
    try {
      result = await sqlite.execute({
        // Guard JSON operations with json_valid so malformed rows do not crash the endpoint.
        sql: `${selectSql}
            WHERE
              LOWER(COALESCE(status, '')) IN ('estimate requested', 'estimated')
              OR (
                raw_json IS NOT NULL
                AND json_valid(raw_json) = 1
                AND (
                  COALESCE(json_array_length(json_extract(raw_json, '$.Estimates')), 0) > 0
                  OR COALESCE(json_array_length(json_extract(raw_json, '$.estimates')), 0) > 0
                  OR LOWER(COALESCE(json_extract(raw_json, '$.current_estimate_approval_status'), '')) IN ('pending', 'approved', 'rejected')
                )
              )
            ${orderSql}`,
      });
    } catch {
      // Fallback for environments where JSON SQL functions are unavailable or unstable.
      result = await sqlite.execute({
        sql: `${selectSql}
              WHERE LOWER(COALESCE(status, '')) IN ('estimate requested', 'estimated')
              ${orderSql}`,
      });
    }

    const rows = rowsAsObjects(result);
    const flattened = rows.flatMap(flattenFromWorkOrderRow);
    const historyById = await syncEstimateHistory(flattened);
    const filtered = flattened
      .map((row) => ({
        ...row,
        status_history: historyById.get(row.estimate_id) || [],
      }))
      .filter((r) => wanted.has(String(r.approval_status || "").trim().toLowerCase()));

    return {
      ok: true,
      results: filtered,
      count: filtered.length,
      statuses: Array.from(wanted),
      source_table: "work_orders_cache",
      history_table: "estimates",
    };
  } catch (err: any) {
    return {
      ok: false,
      status: 500,
      error: `estimates query failed: ${String(err?.message || err || "unknown error")}`,
    };
  }
}
