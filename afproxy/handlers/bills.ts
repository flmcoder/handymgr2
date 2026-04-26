// ============================================================================
// handlers/bills.ts — Billing/AP list + stats + detail + history.
//
// AppFolio DB API field note:
//   status is ApprovalStatus (not Status)
// ============================================================================

import { cacheGet, cacheSet, rowsAsObjects, sqlite, upsertBillingRows } from "../db.ts";
import { resolveGroupPropertyIds } from "../lib/groupUtils.ts";
import { AF_DB, AF_REPORTS, CORS_HEADERS, dbHeaders } from "../config.ts";
import { fetchDbApi, fetchReport } from "../lib/appfolio.ts";
import { fetchWithTimeout, snapDays } from "../lib/fetchUtils.ts";

type BillRecord = Record<string, any>;
type SqlClient = {
  execute: (stmt: string | { sql: string; args?: any[] }) => Promise<any>;
};

const BILL_STATUS_MAP: Record<string, string> = {
  pending: "pending_approval",
  pendingapproval: "pending_approval",
  pending_approval: "pending_approval",
  awaiting_approval: "pending_approval",
  needs_approval: "pending_approval",
  requires_approval: "pending_approval",
  unapproved: "pending_approval",
  approved: "approved",
  approved_for_payment: "approved",
  paid: "paid",
  partially_paid: "paid",
  partial_payment: "paid",
  void: "void",
  voided: "void",
  cancelled: "void",
  canceled: "void",
  rejected: "rejected",
};

function normalizeStatus(raw: unknown): string {
  const s = String(raw || "").trim().toLowerCase().replace(/\s+/g, "_");
  if (!s) return "unknown";
  return BILL_STATUS_MAP[s] || s;
}

function statusLabel(status: string): string {
  if (status === "pending_approval") return "Pending Approval";
  if (status === "approved") return "Approved";
  if (status === "paid") return "Paid";
  if (status === "void") return "Void";
  if (status === "rejected") return "Rejected";
  return "Unknown";
}

function billDateString(row: BillRecord): string {
  return String(
    row.InvoiceDate || row.invoice_date || row.BillDate || row.bill_date ||
      row.PostingDate || row.posting_date || row.LastUpdatedAt ||
      row.last_updated_at || "",
  ).slice(0, 10);
}

// Populate property_reference table for PM login filtering
// This creates a denormalized view mapping properties to groups and vendors
async function populatePropertyReference(
  bills: BillRecord[],
): Promise<void> {
  const now = Date.now();
  const seen = new Set<string>();

  for (const bill of bills) {
    const propertyId = String(
      bill.property_id || bill.PropertyId || bill.property_id_str || "",
    ).trim();
    if (!propertyId || seen.has(propertyId)) continue;
    seen.add(propertyId);

    const vendorId = String(bill.vendor_id || bill.VendorId || "").trim();
    const vendorName = String(bill.vendor_name || bill.VendorName || vendorId || "").trim();
    const propertyName = String(bill.property_name || bill.PropertyName || propertyId || "").trim();
    const propertyGroup = String(bill.property_group || bill.PropertyGroup || "").trim();
    const propertyManager = String(bill.property_manager || bill.PropertyManager || "").trim();

    try {
      await sqlite.execute({
        sql: `INSERT OR REPLACE INTO property_reference
              (property_map_id, property_id, property_name, property_group_uuid, 
               property_group_name, vendor_uuid, vendor_name, pm_name, updated_at)
              VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        args: [
          propertyId,                    // property_map_id (use property_id as map id)
          propertyId,                    // property_id
          propertyName,                  // property_name
          propertyGroup,                 // property_group_uuid
          propertyGroup,                 // property_group_name
          vendorId || null,              // vendor_uuid
          vendorName || null,            // vendor_name
          propertyManager || null,       // pm_name
          now,                           // updated_at
        ],
      });
    } catch {
      // Non-fatal; property reference enrichment is optional
    }
  }
}

function mapBill(
  row: BillRecord,
  vendorMap: Record<string, string>,
  propertyMap: Record<string, string>,
  propertyGroupMap: Record<string, string>,
  pmMap: Record<string, string>,
): BillRecord {
  const lineItems = Array.isArray(row.LineItems)
    ? row.LineItems
    : (Array.isArray(row.line_items)
      ? row.line_items
      : (() => {
        const rawJson = row.line_items_json;
        if (!rawJson) return [];
        try {
          const parsed = JSON.parse(String(rawJson));
          return Array.isArray(parsed) ? parsed : [];
        } catch {
          return [];
        }
      })());
  const firstLine = lineItems[0] || {};
  const payeeUuid = String(
    row.PayeeUuid || row.payee_uuid || row.PayeeId || row.payee_id || "",
  );
  const vendorUuid = String(
    row.VendorUuid || row.vendor_uuid || row.VendorId || row.vendor_id ||
      payeeUuid || "",
  );
  const vendorId = String(
    row.VendorId || row.vendor_id || payeeUuid || vendorUuid || "",
  );
  const propertyId = String(
    row.PropertyId || row.property_id || row.PropertyUuid || row.property_uuid ||
      firstLine.PropertyId || firstLine.property_id || firstLine.PropertyUuid ||
      firstLine.property_uuid || "",
  );
  const status = normalizeStatus(
    row.bill_status || row.BillStatus || row.ApprovalStatus || row.approval_status ||
      row.Status || row.status,
  );
  const amountRaw = row.TotalAmount || row.total_amount || row.Amount ||
    row.amount || "0";
  const amount = Number.parseFloat(String(amountRaw).replace(/[^0-9.-]/g, "")) ||
    0;
  const propertyGroup = String(
    row.property_group_name || row.PropertyGroupName || row.property_group ||
      row.PropertyGroup || propertyGroupMap[propertyId] || "",
  );
  const propertyGroupId = String(
    row.property_group_id || row.property_group_uuid || row.PropertyGroupId ||
      row.PropertyGroupUuid || "",
  );
  const propertyManager = String(
    row.property_manager || row.PropertyManager ||
      (propertyGroup ? (pmMap[propertyGroup] || "") : ""),
  );
  const billId = String(
    row.BillId || row.bill_id || row.Id || row.id || "",
  ).trim();
  const billNumber = String(
    row.bill_number || row.BillNumber || row.InvoiceNumber || row.invoice_number ||
      row.Reference || row.reference || billId,
  ).trim();

  return {
    id: billId,
    bill_number: billNumber,
    status,
    status_label: statusLabel(status),
    vendor_id: vendorId,
    vendor_uuid: vendorUuid || vendorId,
    payee_uuid: payeeUuid || vendorUuid || vendorId,
    vendor_name: String(
      row.VendorName || row.vendor_name || row.PayeeName || row.payee_name ||
        vendorMap[vendorId] || vendorId || "",
    ),
    property_id: propertyId,
    property_name: String(
      row.PropertyName || row.property_name || propertyMap[propertyId] ||
        propertyId || "",
    ),
    property_group_id: propertyGroupId,
    property_group_uuid: propertyGroupId,
    property_group_name: propertyGroup,
    property_group: propertyGroup,
    pm_name: propertyManager,
    property_manager: propertyManager,
    unit_id: String(row.UnitId || row.unit_id || firstLine.UnitId || firstLine.unit_id || ""),
    work_order_id: String(row.WorkOrderId || row.work_order_id || ""),
    invoice_date: String(row.InvoiceDate || row.invoice_date || "").slice(0, 10),
    due_date: String(row.DueDate || row.due_date || "").slice(0, 10),
    posting_date: String(row.PostingDate || row.posting_date || "").slice(0, 10),
    amount,
    total_amount: amount,
    reference: String(row.Reference || row.reference || ""),
      remarks: String(row.Remarks || row.remarks || row.CheckMemo || ""),
      // LineItems — present in both list and detail responses on v0.
      // Mapped here so the frontend detail modal can access them without raw fallback.
      line_items: lineItems,
      raw: row,
  };
}

async function loadPropertyGroupMap(): Promise<Record<string, string>> {
  const out: Record<string, string> = {};
  const cached = await cacheGet("property_groups", "property_groups");
  const rows = Array.isArray(cached?.data) ? cached!.data : [];
  for (const row of rows) {
    const groupName = String(row.Name || row.name || row.group_name || "").trim();
    if (!groupName) continue;
    const props = Array.isArray(row.Properties)
      ? row.Properties
      : (Array.isArray(row.properties) ? row.properties : []);
    for (const p of props) {
      const pid = String(p?.Id || p?.id || p?.PropertyId || p || "").trim();
      if (!pid) continue;
      if (!out[pid]) out[pid] = groupName;
    }
    const propIds = Array.isArray(row.PropertyIds)
      ? row.PropertyIds
      : (Array.isArray(row.property_ids) ? row.property_ids : []);
    for (const p of propIds) {
      const pid = String(p || "").trim();
      if (!pid) continue;
      if (!out[pid]) out[pid] = groupName;
    }
  }
  return out;
}

async function loadRoutingPmMap(): Promise<Record<string, string>> {
  const out: Record<string, string> = {};
  try {
    const res = await sqlite.execute(
      `SELECT group_name, pm_name FROM routing_pm_group_map ORDER BY group_name ASC`,
    );
    const rows = Array.isArray(res?.rows) ? res.rows : [];
    for (const row of rows as any[]) {
      const groupName = String(row.group_name || row.GroupName || "").trim();
      const pmName = String(row.pm_name || row.PmName || "").trim();
      if (groupName && pmName) out[groupName] = pmName;
    }
  } catch {
    // Non-fatal; PM map enrichment is optional.
  }
  return out;
}

function filterBills(
  rows: BillRecord[],
  params: Record<string, string>,
  propertyIds: Set<string> | null,
): BillRecord[] {
  const wantedStatus = normalizeStatus(params.status || "all");
  const vendorId = String(params.vendor_id || "").trim();
  const dateFrom = String(params.date_from || "").slice(0, 10);
  const dateTo = String(params.date_to || "").slice(0, 10);

  return rows.filter((row) => {
    // AppFolio v0 may embed PropertyId only in LineItems — check there too.
    const lineItems = Array.isArray(row.LineItems)
      ? row.LineItems
      : (Array.isArray(row.line_items)
        ? row.line_items
        : (() => {
          const rawJson = row.line_items_json;
          if (!rawJson) return [];
          try { const p = JSON.parse(String(rawJson)); return Array.isArray(p) ? p : []; } catch { return []; }
        })());
    const firstLine = lineItems[0] || {};
    const propId = String(
      row.PropertyId || row.property_id || row.PropertyUuid || row.property_uuid ||
        firstLine.PropertyId || firstLine.property_id ||
        firstLine.PropertyUuid || firstLine.property_uuid || "",
    );
    if (propertyIds && propertyIds.size > 0 && !propertyIds.has(propId)) {
      return false;
    }

    const st = normalizeStatus(
      row.ApprovalStatus || row.approval_status || row.Status || row.status,
    );
    if (wantedStatus !== "all" && st !== wantedStatus) return false;

    const vId = String(
      row.VendorId || row.vendor_id || row.VendorUuid || row.vendor_uuid ||
        row.PayeeUuid || row.payee_uuid || row.PayeeId || row.payee_id || "",
    );
    if (vendorId && vId !== vendorId) return false;

    const d = billDateString(row);
    if (dateFrom && d && d < dateFrom) return false;
    if (dateTo && d && d > dateTo) return false;
    return true;
  }).sort((a, b) => {
    const da = billDateString(a);
    const db = billDateString(b);
    return db.localeCompare(da);
  });
}

async function loadVendorMap(): Promise<Record<string, string>> {
  const out: Record<string, string> = {};
  const cached = await cacheGet("vendors", "vendors");
  const rows = Array.isArray(cached?.data) ? cached!.data : [];
  for (const row of rows) {
    const id = String(row.vendor_id || row.id || row.VendorId || "");
    if (!id) continue;
    const name = String(
      row.company_name || row.name ||
        [row.first_name, row.last_name].filter(Boolean).join(" ") || id,
    );
    out[id] = name;
  }
  return out;
}

async function loadPropertyMap(): Promise<Record<string, string>> {
  const out: Record<string, string> = {};

  const propMapCached = await cacheGet("property_map", "property_map");
  if (propMapCached?.data && typeof propMapCached.data === "object") {
    for (const [id, value] of Object.entries(propMapCached.data as Record<string, any>)) {
      if (value && typeof value === "object") {
        out[id] = String((value as Record<string, any>).name || id);
      }
    }
  }

  if (Object.keys(out).length > 0) return out;

  const propsCached = await cacheGet("properties", "properties");
  const rows = Array.isArray(propsCached?.data) ? propsCached!.data : [];
  for (const row of rows) {
    const id = String(row.property_id || row.id || row.PropertyId || "");
    if (!id) continue;
    out[id] = String(row.property_name || row.name || row.property || id);
  }

  return out;
}

async function fetchRawBills(
  params: Record<string, string>,
  force = false,
): Promise<{ rows: BillRecord[]; fromCache: boolean; cachedAt?: string; sourceApi: "v2" | "v0" }> {
  const days = snapDays(parseInt(params.days || "365", 10) || 365, "bills");
  const max = Math.max(50, Math.min(10000, parseInt(params.max || "1500", 10) || 1500));
  const groupUuid = String(params.group_uuid || "").trim();
  const preferV2 = String(params.prefer_v2 || "true").toLowerCase() !== "false";
  const cacheKey = groupUuid
    ? `bills_${preferV2 ? "v2" : "v0"}_grp_${groupUuid}_${days}_${max}`
    : `bills_${preferV2 ? "v2" : "v0"}_${days}_${max}`;

  if (!force) {
    const cached = await cacheGet(cacheKey, "bills");
    if (cached) {
      const data = Array.isArray(cached.data) ? cached.data : [];
      return {
        rows: data,
        fromCache: true,
        cachedAt: cached.cached_at,
        sourceApi: preferV2 ? "v2" : "v0",
      };
    }
  }

  if (preferV2) {
    try {
      const fromDate = new Date();
      fromDate.setDate(fromDate.getDate() - days);
      const toDate = new Date();
      const reportFilters: Record<string, any> = {
        occurred_on_from: fromDate.toISOString().slice(0, 10),
        occurred_on_to: toDate.toISOString().slice(0, 10),
        date_type: "Bill Date",
        paginate_results: true,
        limit: max,
      };
      const v2Rows = await fetchReport("bill_detail", reportFilters);
      const rows = Array.isArray(v2Rows) ? v2Rows : [];
      await cacheSet(cacheKey, "bills", rows, rows.length);
      upsertBillingRows(rows).catch(() => {});
      return { rows, fromCache: false, sourceApi: "v2" };
    } catch {
      // Fall through to v0 fallback below.
    }
  }

  const fromDate = new Date();
  fromDate.setDate(fromDate.getDate() - days);

  let path = `/api/v0/bills?filters[LastUpdatedAtFrom]=${
    encodeURIComponent(fromDate.toISOString())
  }&page%5Bsize%5D=200`;

  if (groupUuid) {
    path += `&property_group_id=${encodeURIComponent(groupUuid)}`;
  }

  const rows = await fetchDbApi(path, max);
  await cacheSet(cacheKey, "bills", rows, rows.length);
  upsertBillingRows(rows).catch(() => {});
  return { rows, fromCache: false, sourceApi: "v0" };
}

function paginate(rows: BillRecord[], params: Record<string, string>) {
  const page = Math.max(1, parseInt(params.page || "1", 10) || 1);
  const perPage = Math.max(
    1,
    Math.min(200, parseInt(params.per_page || "50", 10) || 50),
  );
  const total = rows.length;
  const totalPages = Math.max(1, Math.ceil(total / perPage));
  const safePage = Math.min(page, totalPages);
  const start = (safePage - 1) * perPage;
  const end = start + perPage;
  return {
    page: safePage,
    perPage,
    total,
    totalPages,
    rows: rows.slice(start, end),
  };
}

async function handleBillsList(params: Record<string, string>): Promise<any> {
  const force = String(params.force_refresh || "").toLowerCase() === "true";
  const raw = await fetchRawBills(params, force);
  const propertyIds = await resolveGroupPropertyIds(params);
  const filtered = filterBills(raw.rows, params, propertyIds);
  const page = paginate(filtered, params);

  const [vendorMap, propertyMap, propertyGroupMap, pmMap] = await Promise.all([
    loadVendorMap(),
    loadPropertyMap(),
    loadPropertyGroupMap(),
    loadRoutingPmMap(),
  ]);

  const results = page.rows.map((r) => mapBill(r, vendorMap, propertyMap, propertyGroupMap, pmMap));

  // Populate property reference table asynchronously (non-blocking)
  populatePropertyReference(results).catch(() => {});

  return {
    ok: true,
    results,
    count: page.rows.length,
    total: page.total,
    page: page.page,
    per_page: page.perPage,
    total_pages: page.totalPages,
    from_cache: raw.fromCache,
    cached_at: raw.cachedAt,
    source_api: raw.sourceApi,
  };
}

async function handleBillsStats(params: Record<string, string>): Promise<any> {
  const raw = await fetchRawBills(params, false);
  const propertyIds = await resolveGroupPropertyIds(params);
  const filtered = filterBills(raw.rows, params, propertyIds);

  function sumAmount(rows: BillRecord[]): number {
    return rows.reduce((s, r) => {
      const n = Number.parseFloat(
        String(r.TotalAmount || r.total_amount || r.Amount || r.amount || "0")
          .replace(/[^0-9.-]/g, ""),
      ) || 0;
      return s + n;
    }, 0);
  }

  const pending = filtered.filter((r) =>
    normalizeStatus(r.ApprovalStatus || r.approval_status || r.Status || r.status) === "pending_approval"
  );
  const approved = filtered.filter((r) =>
    normalizeStatus(r.ApprovalStatus || r.approval_status || r.Status || r.status) === "approved"
  );
  const paid = filtered.filter((r) =>
    normalizeStatus(r.ApprovalStatus || r.approval_status || r.Status || r.status) === "paid"
  );

  const vendorsWithOpen = new Set<string>();
  [...pending, ...approved].forEach((r) => {
    const vid = String(
      r.VendorId || r.vendor_id || r.VendorUuid || r.vendor_uuid ||
        r.PayeeUuid || r.payee_uuid || "",
    ).trim();
    if (vid) vendorsWithOpen.add(vid);
  });

  return {
    ok: true,
    pending_approval_count: pending.length,
    pending_approval_amount: sumAmount(pending),
    approved_not_paid_count: approved.length,
    total_outstanding: sumAmount(approved),
    paid_this_period_count: paid.length,
    paid_this_period_amount: sumAmount(paid),
    vendors_with_open_bills: vendorsWithOpen.size,
    source_api: raw.sourceApi,
    from_cache: raw.fromCache,
    cached_at: raw.cachedAt,
  };
}

async function handleBillDetail(params: Record<string, string>): Promise<any> {
  const billId = String(params.bill_id || params.id || "").trim();
  if (!billId) return { ok: false, error: "Missing bill_id" };

  const raw = await fetchRawBills({ ...params, max: "5000" }, false);
  const matchesBillId = (row: BillRecord, target: string): boolean => {
    const t = String(target || "").trim();
    if (!t) return false;
    const candidates = [
      row.BillId,
      row.bill_id,
      row.Id,
      row.id,
      row.Reference,
      row.reference,
      row.bill_number,
      row.invoice_number,
    ];
    return candidates.some((value) => String(value || "").trim() === t);
  };

  let found: BillRecord | undefined = raw.rows.find((r) =>
    matchesBillId(r, billId)
  );

  if (!found) {
    const matches = await fetchDbApi(
      `/api/v0/bills?filters[Id]=${encodeURIComponent(billId)}&page%5Bsize%5D=1`,
      1,
    );
    found = matches[0];
  }

  if (!found) return { ok: false, error: "Bill not found", bill_id: billId };

  const lookupBillId = String(
    found.BillId || found.bill_id || found.Id || found.id || billId,
  ).trim() || billId;

  try {
    const detailPath = `/api/v0/bills/${encodeURIComponent(lookupBillId)}`;
    let detailResp = await fetchWithTimeout(`${AF_DB}${detailPath}`, {
      headers: dbHeaders(),
    });
    if (detailResp.status === 401 || detailResp.status === 403) {
      detailResp = await fetchWithTimeout(`${AF_REPORTS}${detailPath}`, {
        headers: dbHeaders(),
      });
    }
    if (detailResp.ok) {
      const detailData = await detailResp.json().catch(() => null);
      if (detailData && typeof detailData === "object") {
        found = { ...found, ...detailData };
      }
    }
  } catch {
    // non-fatal
  }

  const [vendorMap, propertyMap, propertyGroupMap, pmMap] = await Promise.all([
    loadVendorMap(),
    loadPropertyMap(),
    loadPropertyGroupMap(),
    loadRoutingPmMap(),
  ]);

  const result = mapBill(found!, vendorMap, propertyMap, propertyGroupMap, pmMap);

  // Populate property reference table asynchronously (non-blocking)
  populatePropertyReference([result]).catch(() => {});

  return {
    ok: true,
    result,
  };
}

async function handleBillsHistory(params: Record<string, string>): Promise<any> {
  const dateFrom = String(params.date_from || "").slice(0, 10);
  const dateTo = String(params.date_to || "").slice(0, 10);
  if (!dateFrom || !dateTo) {
    return { ok: false, error: "date_from and date_to are required" };
  }

  const fromMs = Date.parse(`${dateFrom}T00:00:00Z`);
  const toMs = Date.parse(`${dateTo}T23:59:59Z`);
  if (isNaN(fromMs) || isNaN(toMs) || toMs < fromMs) {
    return { ok: false, error: "Invalid date range: date_from/date_to" };
  }

  const rangeDays = Math.max(1, Math.ceil((toMs - fromMs) / 86_400_000) + 1);
  const fetchDays = Math.max(30, Math.min(3650, rangeDays + 14));
  const propertyIds = await resolveGroupPropertyIds(params);
  const preferV2 = String(params.prefer_v2 || "true").toLowerCase() !== "false";

  let rawRows: BillRecord[] = [];
  let fromCache = false;
  let cachedAt: string | undefined;
  let sourceApi: "v2" | "v0" = "v0";

  if (preferV2) {
    try {
      const reportFilters: Record<string, any> = {
        occurred_on_from: dateFrom,
        occurred_on_to: dateTo,
        date_type: "Bill Date",
        paginate_results: true,
        limit: 5000,
      };
      if (propertyIds && propertyIds.size > 0) {
        reportFilters.properties = {
          property_visibility: "both",
          properties_ids: Array.from(propertyIds),
        };
      }
      const v2Rows = await fetchReport("bill_detail", reportFilters);
      rawRows = (Array.isArray(v2Rows) ? v2Rows : []).map((row: BillRecord) => ({
        Id: row.bill_id || row.id || "",
        BillId: row.bill_id || row.id || "",
        BillNumber: row.bill_number || row.invoice_number || row.reference || "",
        Reference: row.bill_number || row.invoice_number || row.reference || "",
        VendorId: row.vendor_id || row.vendor_uuid || row.payee_id || row.payee_uuid || "",
        VendorName: row.vendor_name || row.payee_name || "",
        PropertyId: row.property_id || "",
        PropertyName: row.property_name || "",
        PropertyGroupName: row.property_group_name || row.property_group || "",
        PropertyGroupUuid: row.property_group_uuid || row.property_group_id || "",
        PropertyManager: row.pm_name || row.property_manager || "",
        InvoiceDate: row.bill_date || row.invoice_date || "",
        DueDate: row.due_date || "",
        PostingDate: row.payment_date || row.bill_date || "",
        ApprovalStatus: row.bill_status || row.status || "",
        TotalAmount: row.total_amount || row.amount || "0",
        WorkOrderId: row.work_order_id || "",
        UnitId: row.unit_id || "",
        line_items: row.line_items || [],
        raw: row,
      }));
      sourceApi = "v2";
    } catch {
      // Fall through to v0 path below.
    }
  }

  if (!rawRows.length) {
    const raw = await fetchRawBills({ ...params, days: String(fetchDays) }, false);
    rawRows = raw.rows;
    fromCache = raw.fromCache;
    cachedAt = raw.cachedAt;
    sourceApi = "v0";
  }

  const filtered = filterBills(rawRows, params, propertyIds);
  const inRange = filtered.filter((r) => {
    const d = billDateString(r);
    return !!d && d >= dateFrom && d <= dateTo;
  });

  const page = Math.max(1, parseInt(String(params.page || "1"), 10) || 1);
  const perPage = Math.max(
    1,
    Math.min(200, parseInt(String(params.per_page || "100"), 10) || 100),
  );
  const total = inRange.length;
  const totalPages = Math.max(1, Math.ceil(total / perPage));
  const safePage = Math.min(page, totalPages);
  const start = (safePage - 1) * perPage;
  const paged = inRange.slice(start, start + perPage);

  const [vendorMap, propertyMap, propertyGroupMap, pmMap] = await Promise.all([
    loadVendorMap(),
    loadPropertyMap(),
    loadPropertyGroupMap(),
    loadRoutingPmMap(),
  ]);

  const results = paged.map((r) => mapBill(r, vendorMap, propertyMap, propertyGroupMap, pmMap));

  // Populate property reference table asynchronously (non-blocking)
  populatePropertyReference(results).catch(() => {});

  return {
    ok: true,
    results,
    count: results.length,
    total,
    page: safePage,
    per_page: perPage,
    total_pages: totalPages,
    truncated: total > perPage,
    from_cache: fromCache,
    cached_at: cachedAt,
    source_api: sourceApi,
  };
}

async function handleBillsSync(params: Record<string, string>): Promise<any> {
  const raw = await fetchRawBills(params, true);
  return {
    ok: true,
    synced: raw.rows.length,
    from_cache: raw.fromCache,
    cached_at: raw.cachedAt,
  };
}

async function handleBillAttachmentsList(
  params: Record<string, string>,
): Promise<any> {
  const billId = String(params.bill_id || "").trim();
  if (!billId) return { ok: false, error: "bill_id is required" };

  const cacheKey = `bill_attachments_${billId}`;
  const force = String(params.force || params.force_refresh || "").toLowerCase() ===
    "true";
  if (!force) {
    const cached = await cacheGet(cacheKey, "bills");
    if (cached && Array.isArray(cached.data)) {
      return {
        ok: true,
        attachments: cached.data,
        count: cached.data.length,
        from_cache: true,
        cached_at: cached.cached_at,
      };
    }
  }

  const path = `/api/v0/bills/${encodeURIComponent(billId)}/attachments`;
  let resp = await fetchWithTimeout(`${AF_DB}${path}`, { headers: dbHeaders() });
  if ([401, 403, 404, 422].includes(resp.status)) {
    resp = await fetchWithTimeout(`${AF_REPORTS}${path}`, {
      headers: dbHeaders(),
    });
  }

  if (!resp.ok) {
    const detail = await resp.text().catch(() => "");
    return {
      ok: false,
      error: `attachments fetch failed: HTTP ${resp.status}`,
      status: resp.status,
      detail: detail.substring(0, 400),
    };
  }

  const data = await resp.json().catch(() => []);
  const attachments = Array.isArray(data)
    ? data
    : (Array.isArray(data?.Results)
      ? data.Results
      : (Array.isArray(data?.results)
        ? data.results
        : (Array.isArray(data?.data) ? data.data : [])));

  await cacheSet(cacheKey, "bills", attachments, attachments.length);
  return {
    ok: true,
    attachments,
    count: attachments.length,
    from_cache: false,
  };
}

async function handleBillAttachmentUpload(
  params: Record<string, string>,
  req?: Request,
): Promise<any> {
  const billId = String(params.bill_id || "").trim();
  if (!billId) return { ok: false, error: "bill_id is required" };
  if (!req) return { ok: false, error: "Request body required" };

  const contentType = req.headers.get("Content-Type") ||
    "application/octet-stream";
  const bodyBuffer = await req.arrayBuffer();
  if (!bodyBuffer || bodyBuffer.byteLength === 0) {
    return { ok: false, error: "Empty upload body" };
  }

  const path = `/api/v0/bills/${encodeURIComponent(billId)}/attachments`;
  const headers = {
    ...dbHeaders(),
    "Content-Type": contentType,
  };

  let resp = await fetchWithTimeout(`${AF_DB}${path}`, {
    method: "POST",
    headers,
    body: bodyBuffer,
  });
  if ([401, 403, 404, 422].includes(resp.status)) {
    resp = await fetchWithTimeout(`${AF_REPORTS}${path}`, {
      method: "POST",
      headers,
      body: bodyBuffer,
    });
  }

  const text = await resp.text().catch(() => "");

  // Bust only this bill's attachment cache.
  const cacheKey = `bill_attachments_${billId}`;
  try {
    await sqlite.execute({
      sql: `DELETE FROM api_cache WHERE cache_key = ? OR cache_key LIKE ?`,
      args: [cacheKey, `${cacheKey}::%`],
    });
  } catch {
    // non-fatal
  }

  let parsed: any = null;
  try {
    parsed = text ? JSON.parse(text) : null;
  } catch {
    parsed = null;
  }

  return {
    ok: resp.ok,
    status: resp.status,
    detail: parsed ?? text.substring(0, 800),
    bill_id: billId,
  };
}

export async function handleBills(
  params: Record<string, string>,
  req?: Request,
  action = "bills",
): Promise<any> {
  if (action === "bills_list_v2") return await handleBillsList(params);
  if (action === "bills_stats") return await handleBillsStats(params);
  if (action === "bill_detail") return await handleBillDetail(params);
  if (action === "bills_history") return await handleBillsHistory(params);
  if (action === "bills_sync") return await handleBillsSync(params);
  if (action === "bill_attachments") return await handleBillAttachmentsList(params);
  if (action === "bill_attachment_upload") {
    return await handleBillAttachmentUpload(params, req);
  }
  return await handleBillsList(params);
}

function billsJson(
  body: unknown,
  status = 200,
): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      "Content-Type": "application/json",
      ...CORS_HEADERS,
    },
  });
}

function parseLimitOffset(params: URLSearchParams): { limit: number; offset: number } {
  const rawLimit = Number(params.get("limit") || "50");
  const rawOffset = Number(params.get("offset") || "0");
  const limit = Number.isFinite(rawLimit)
    ? Math.max(1, Math.min(200, Math.floor(rawLimit)))
    : 50;
  const offset = Number.isFinite(rawOffset)
    ? Math.max(0, Math.floor(rawOffset))
    : 0;
  return { limit, offset };
}

function asCount(result: any): number {
  const rows = rowsAsObjects(result);
  if (!rows.length) return 0;
  return Number(rows[0].total || 0) || 0;
}

function normalizeLikeText(value: string | null): string | null {
  const v = String(value || "").trim();
  return v ? v : null;
}

export async function handleBillsRoute(
  _req: Request,
  _db: SqlClient,
  params: URLSearchParams,
): Promise<Response> {
  const action = String(params.get("action") || "").trim();
  const { limit, offset } = parseLimitOffset(params);
  const groupId = normalizeLikeText(params.get("group_id"));
  const status = normalizeLikeText(params.get("status"));
  const flatParams: Record<string, string> = {};
  for (const [k, v] of params.entries()) flatParams[k] = v;
  if (groupId) flatParams.group_id = groupId;
  if (status) flatParams.status = status;

  const actionHandlers = {
    bills_by_vendor: true,
    bills_by_property: true,
    bills_by_unit: true,
    bills_by_wo: true,
    bills_by_wo_number: true,
    bills_by_invoice: true,
    bills_due_range: true,
    bills_list: true,
  } as const;
  if (!actionHandlers[action as keyof typeof actionHandlers]) {
    return billsJson({ ok: false, error: `Unsupported bills route action: ${action}` }, 404);
  }

  if (action === "bills_list" && !groupId) {
    return billsJson({ ok: false, error: "Missing required query param: group_id" }, 400);
  }

  const vendorId = normalizeLikeText(params.get("vendor_id"));
  const propertyId = normalizeLikeText(params.get("property_id"));
  const unitId = normalizeLikeText(params.get("unit_id"));
  const woId = normalizeLikeText(params.get("wo_id"));
  const woNumber = normalizeLikeText(params.get("wo_number"));
  const invoiceNumber = normalizeLikeText(params.get("invoice_number"));
  const dueFrom = normalizeLikeText(params.get("due_from"));
  const dueTo = normalizeLikeText(params.get("due_to"));

  if (action === "bills_by_vendor" && !vendorId) {
    return billsJson({ ok: false, error: "Missing required query param: vendor_id" }, 400);
  }
  if (action === "bills_by_property" && !propertyId) {
    return billsJson({ ok: false, error: "Missing required query param: property_id" }, 400);
  }
  if (action === "bills_by_unit" && !unitId) {
    return billsJson({ ok: false, error: "Missing required query param: unit_id" }, 400);
  }
  if (action === "bills_by_wo" && !woId) {
    return billsJson({ ok: false, error: "Missing required query param: wo_id" }, 400);
  }
  if (action === "bills_by_wo_number" && !woNumber) {
    return billsJson({ ok: false, error: "Missing required query param: wo_number" }, 400);
  }
  if (action === "bills_by_invoice" && !invoiceNumber) {
    return billsJson({ ok: false, error: "Missing required query param: invoice_number" }, 400);
  }
  if (action === "bills_due_range" && (!dueFrom || !dueTo)) {
    return billsJson({ ok: false, error: "Missing required query params: due_from and due_to" }, 400);
  }

  const rowPropertyId = (row: BillRecord): string => {
    const lineItems = Array.isArray(row.LineItems)
      ? row.LineItems
      : (Array.isArray(row.line_items)
        ? row.line_items
        : (() => {
          const rawJson = row.line_items_json;
          if (!rawJson) return [];
          try {
            const parsed = JSON.parse(String(rawJson));
            return Array.isArray(parsed) ? parsed : [];
          } catch {
            return [];
          }
        })());
    const firstLine = lineItems[0] || {};
    return String(
      row.PropertyId || row.property_id || row.PropertyUuid || row.property_uuid ||
        firstLine.PropertyId || firstLine.property_id || firstLine.PropertyUuid ||
        firstLine.property_uuid || "",
    ).trim();
  };

  const rowUnitId = (row: BillRecord): string => {
    const lineItems = Array.isArray(row.LineItems)
      ? row.LineItems
      : (Array.isArray(row.line_items)
        ? row.line_items
        : []);
    const firstLine = lineItems[0] || {};
    return String(row.UnitId || row.unit_id || firstLine.UnitId || firstLine.unit_id || "").trim();
  };

  const rowVendorId = (row: BillRecord): string =>
    String(
      row.VendorId || row.vendor_id || row.VendorUuid || row.vendor_uuid ||
        row.PayeeUuid || row.payee_uuid || row.PayeeId || row.payee_id || "",
    ).trim();

  const rowWoId = (row: BillRecord): string =>
    String(row.WorkOrderId || row.work_order_id || "").trim();

  const rowWoNumber = (row: BillRecord): string =>
    String(row.WorkOrderNumber || row.work_order_number || row.work_order_num || "").trim();

  const rowInvoice = (row: BillRecord): string =>
    String(row.InvoiceNumber || row.invoice_number || row.Reference || row.reference || row.InvoiceDate || row.invoice_date || "").trim();

  const rowDueDate = (row: BillRecord): string =>
    String(row.DueDate || row.due_date || "").slice(0, 10);

  try {
    const raw = await fetchRawBills(flatParams, false);
    const propertyIds = await resolveGroupPropertyIds(flatParams);
    let filtered = filterBills(raw.rows, flatParams, propertyIds);

    if (action === "bills_by_vendor" && vendorId) {
      filtered = filtered.filter((row) => rowVendorId(row) === vendorId);
    } else if (action === "bills_by_property" && propertyId) {
      filtered = filtered.filter((row) => rowPropertyId(row) === propertyId);
    } else if (action === "bills_by_unit" && unitId) {
      filtered = filtered.filter((row) => rowUnitId(row) === unitId);
    } else if (action === "bills_by_wo" && woId) {
      filtered = filtered.filter((row) => rowWoId(row) === woId);
    } else if (action === "bills_by_wo_number" && woNumber) {
      filtered = filtered.filter((row) => rowWoNumber(row) === woNumber);
    } else if (action === "bills_by_invoice" && invoiceNumber) {
      filtered = filtered.filter((row) => rowInvoice(row) === invoiceNumber);
    } else if (action === "bills_due_range" && dueFrom && dueTo) {
      filtered = filtered.filter((row) => {
        const due = rowDueDate(row);
        if (!due) return false;
        return due >= dueFrom && due <= dueTo;
      });
    }

    const [vendorMap, propertyMap, propertyGroupMap, pmMap] = await Promise.all([
      loadVendorMap(),
      loadPropertyMap(),
      loadPropertyGroupMap(),
      loadRoutingPmMap(),
    ]);

    const total = filtered.length;
    const sliced = filtered.slice(offset, offset + limit);
    const data = sliced.map((row) => mapBill(row, vendorMap, propertyMap, propertyGroupMap, pmMap));

    return billsJson({
      ok: true,
      data,
      total,
      limit,
      offset,
    });
  } catch (e: any) {
    return billsJson(
      {
        ok: false,
        error: String(e?.message || "Billing query failed"),
      },
      500,
    );
  }
}