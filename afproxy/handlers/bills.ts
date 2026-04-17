// ============================================================================
// handlers/bills.ts — Billing/AP list + stats + detail + history.
//
// AppFolio DB API field note:
//   status is ApprovalStatus (not Status)
// ============================================================================

import { cacheGet, cacheSet, rowsAsObjects, sqlite, upsertBillingRows } from "../db.ts";
import { AF_DB, AF_REPORTS, CORS_HEADERS, dbHeaders } from "../config.ts";
import { fetchDbApi } from "../lib/appfolio.ts";
import { fetchWithTimeout, snapDays } from "../lib/fetchUtils.ts";

type BillRecord = Record<string, any>;
type SqlClient = {
  execute: (stmt: string | { sql: string; args?: any[] }) => Promise<any>;
};

const BILL_STATUS_MAP: Record<string, string> = {
  pending: "pending_approval",
  pendingapproval: "pending_approval",
  pending_approval: "pending_approval",
  approved: "approved",
  paid: "paid",
  void: "void",
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

function mapBill(
  row: BillRecord,
  vendorMap: Record<string, string>,
  propertyMap: Record<string, string>,
  propertyGroupMap: Record<string, string>,
  pmMap: Record<string, string>,
): BillRecord {
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
      "",
  );
  const status = normalizeStatus(
    row.ApprovalStatus || row.approval_status || row.Status || row.status,
  );
  const amountRaw = row.TotalAmount || row.total_amount || row.Amount ||
    row.amount || "0";
  const amount = Number.parseFloat(String(amountRaw).replace(/[^0-9.-]/g, "")) ||
    0;
  const propertyGroup = String(
    row.property_group || row.PropertyGroup || propertyGroupMap[propertyId] ||
      "",
  );
  const propertyManager = String(
    row.property_manager || row.PropertyManager ||
      (propertyGroup ? (pmMap[propertyGroup] || "") : ""),
  );

  return {
    id: String(row.Id || row.id || row.BillId || row.bill_id || ""),
    bill_number: String(row.Reference || row.reference || row.Id || row.id || ""),
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
    property_group: propertyGroup,
    property_manager: propertyManager,
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
      line_items: Array.isArray(row.LineItems)
        ? row.LineItems
        : (Array.isArray(row.line_items) ? row.line_items : []),
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
    const propId = String(
      row.PropertyId || row.property_id || row.PropertyUuid || row.property_uuid ||
        "",
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

async function resolveGroupPropertyIds(
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

  return set;
}

async function fetchRawBills(
  params: Record<string, string>,
  force = false,
): Promise<{ rows: BillRecord[]; fromCache: boolean; cachedAt?: string }> {
  const days = snapDays(parseInt(params.days || "365", 10) || 365, "bills");
  const max = Math.max(50, Math.min(10000, parseInt(params.max || "1500", 10) || 1500));
  const groupUuid = String(params.group_uuid || "").trim();
  const cacheKey = groupUuid
    ? `bills_grp_${groupUuid}_${days}_${max}`
    : `bills_${days}_${max}`;

  if (!force) {
    const cached = await cacheGet(cacheKey, "bills");
    if (cached) {
      const data = Array.isArray(cached.data) ? cached.data : [];
      return { rows: data, fromCache: true, cachedAt: cached.cached_at };
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
  return { rows, fromCache: false };
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

  return {
    ok: true,
    results: page.rows.map((r) => mapBill(r, vendorMap, propertyMap, propertyGroupMap, pmMap)),
    count: page.rows.length,
    total: page.total,
    page: page.page,
    per_page: page.perPage,
    total_pages: page.totalPages,
    from_cache: raw.fromCache,
    cached_at: raw.cachedAt,
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
    from_cache: raw.fromCache,
    cached_at: raw.cachedAt,
  };
}

async function handleBillDetail(params: Record<string, string>): Promise<any> {
  const billId = String(params.bill_id || params.id || "").trim();
  if (!billId) return { ok: false, error: "Missing bill_id" };

  const raw = await fetchRawBills({ ...params, max: "5000" }, false);
  let found: BillRecord | undefined = raw.rows.find((r) =>
    String(r.Id || r.id || "").trim() === billId
  );

  if (!found) {
    const matches = await fetchDbApi(
      `/api/v0/bills?filters[Id]=${encodeURIComponent(billId)}&page%5Bsize%5D=1`,
      1,
    );
    found = matches[0];
  }

  if (!found) return { ok: false, error: "Bill not found", bill_id: billId };

  try {
    const detailPath = `/api/v0/bills/${encodeURIComponent(billId)}`;
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

  return {
    ok: true,
    result: mapBill(found!, vendorMap, propertyMap, propertyGroupMap, pmMap),
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

  const raw = await fetchRawBills({ ...params, days: String(fetchDays) }, false);
  const propertyIds = await resolveGroupPropertyIds(params);
  const filtered = filterBills(raw.rows, params, propertyIds);
  const inRange = filtered.filter((r) => {
    const d = billDateString(r);
    return !!d && d >= dateFrom && d <= dateTo;
  });

  const [vendorMap, propertyMap, propertyGroupMap, pmMap] = await Promise.all([
    loadVendorMap(),
    loadPropertyMap(),
    loadPropertyGroupMap(),
    loadRoutingPmMap(),
  ]);

  return {
    ok: true,
    results: inRange.slice(0, 2000).map((r) => mapBill(r, vendorMap, propertyMap, propertyGroupMap, pmMap)),
    count: inRange.length,
    from_cache: raw.fromCache,
    cached_at: raw.cachedAt,
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
  db: SqlClient,
  params: URLSearchParams,
): Promise<Response> {
  const action = String(params.get("action") || "").trim();
  const { limit, offset } = parseLimitOffset(params);
  const groupId = normalizeLikeText(params.get("group_id"));
  const status = normalizeLikeText(params.get("status"));

  let dataSql = "";
  let countSql = "";
  let dataArgs: any[] = [];
  let countArgs: any[] = [];

  if (action === "bills_by_vendor") {
    const vendorId = normalizeLikeText(params.get("vendor_id"));
    if (!vendorId) {
      return billsJson({ ok: false, error: "Missing required query param: vendor_id" }, 400);
    }
    if (groupId) {
      dataSql = "SELECT bm.* FROM billing_map bm INNER JOIN group_resolution_cache grc ON grc.property_map_id = bm.property_map_id WHERE grc.property_group_id = ? AND bm.vendor_id = ? ORDER BY bm.due_date DESC LIMIT ? OFFSET ?";
      countSql = "SELECT COUNT(*) AS total FROM billing_map bm INNER JOIN group_resolution_cache grc ON grc.property_map_id = bm.property_map_id WHERE grc.property_group_id = ? AND bm.vendor_id = ?";
      dataArgs = [groupId, vendorId, limit, offset];
      countArgs = [groupId, vendorId];
    } else {
      dataSql = "SELECT bm.* FROM billing_map bm WHERE bm.vendor_id = ? ORDER BY bm.due_date DESC LIMIT ? OFFSET ?";
      countSql = "SELECT COUNT(*) AS total FROM billing_map bm WHERE bm.vendor_id = ?";
      dataArgs = [vendorId, limit, offset];
      countArgs = [vendorId];
    }
  } else if (action === "bills_by_property") {
    const propertyId = normalizeLikeText(params.get("property_id"));
    if (!propertyId) {
      return billsJson({ ok: false, error: "Missing required query param: property_id" }, 400);
    }
    if (groupId) {
      dataSql = "SELECT bm.* FROM billing_map bm INNER JOIN group_resolution_cache grc ON grc.property_map_id = bm.property_map_id WHERE grc.property_group_id = ? AND bm.property_map_id = ? ORDER BY bm.due_date DESC LIMIT ? OFFSET ?";
      countSql = "SELECT COUNT(*) AS total FROM billing_map bm INNER JOIN group_resolution_cache grc ON grc.property_map_id = bm.property_map_id WHERE grc.property_group_id = ? AND bm.property_map_id = ?";
      dataArgs = [groupId, propertyId, limit, offset];
      countArgs = [groupId, propertyId];
    } else {
      dataSql = "SELECT bm.* FROM billing_map bm WHERE bm.property_map_id = ? ORDER BY bm.due_date DESC LIMIT ? OFFSET ?";
      countSql = "SELECT COUNT(*) AS total FROM billing_map bm WHERE bm.property_map_id = ?";
      dataArgs = [propertyId, limit, offset];
      countArgs = [propertyId];
    }
  } else if (action === "bills_by_wo") {
    const woId = normalizeLikeText(params.get("wo_id"));
    if (!woId) {
      return billsJson({ ok: false, error: "Missing required query param: wo_id" }, 400);
    }
    dataSql = "SELECT bm.* FROM billing_map bm WHERE bm.work_order_id = ? ORDER BY bm.due_date DESC LIMIT ? OFFSET ?";
    countSql = "SELECT COUNT(*) AS total FROM billing_map bm WHERE bm.work_order_id = ?";
    dataArgs = [woId, limit, offset];
    countArgs = [woId];
  } else if (action === "bills_by_wo_number") {
    const woNumber = normalizeLikeText(params.get("wo_number"));
    if (!woNumber) {
      return billsJson({ ok: false, error: "Missing required query param: wo_number" }, 400);
    }
    dataSql = "SELECT bm.* FROM billing_map bm WHERE bm.work_order_number = ? ORDER BY bm.due_date DESC LIMIT ? OFFSET ?";
    countSql = "SELECT COUNT(*) AS total FROM billing_map bm WHERE bm.work_order_number = ?";
    dataArgs = [woNumber, limit, offset];
    countArgs = [woNumber];
  } else if (action === "bills_by_invoice") {
    const invoiceNumber = normalizeLikeText(params.get("invoice_number"));
    if (!invoiceNumber) {
      return billsJson({ ok: false, error: "Missing required query param: invoice_number" }, 400);
    }
    // billing_map stores invoice_date; this route maps invoice_number input to that field.
    dataSql = "SELECT bm.* FROM billing_map bm WHERE bm.invoice_date = ? ORDER BY bm.due_date DESC LIMIT ? OFFSET ?";
    countSql = "SELECT COUNT(*) AS total FROM billing_map bm WHERE bm.invoice_date = ?";
    dataArgs = [invoiceNumber, limit, offset];
    countArgs = [invoiceNumber];
  } else if (action === "bills_due_range") {
    const dueFrom = normalizeLikeText(params.get("due_from"));
    const dueTo = normalizeLikeText(params.get("due_to"));
    if (!dueFrom || !dueTo) {
      return billsJson({ ok: false, error: "Missing required query params: due_from and due_to" }, 400);
    }
    if (groupId) {
      dataSql = "SELECT bm.* FROM billing_map bm INNER JOIN group_resolution_cache grc ON grc.property_map_id = bm.property_map_id WHERE grc.property_group_id = ? AND bm.due_date >= ? AND bm.due_date <= ? ORDER BY bm.due_date DESC LIMIT ? OFFSET ?";
      countSql = "SELECT COUNT(*) AS total FROM billing_map bm INNER JOIN group_resolution_cache grc ON grc.property_map_id = bm.property_map_id WHERE grc.property_group_id = ? AND bm.due_date >= ? AND bm.due_date <= ?";
      dataArgs = [groupId, dueFrom, dueTo, limit, offset];
      countArgs = [groupId, dueFrom, dueTo];
    } else {
      dataSql = "SELECT bm.* FROM billing_map bm WHERE bm.due_date >= ? AND bm.due_date <= ? ORDER BY bm.due_date DESC LIMIT ? OFFSET ?";
      countSql = "SELECT COUNT(*) AS total FROM billing_map bm WHERE bm.due_date >= ? AND bm.due_date <= ?";
      dataArgs = [dueFrom, dueTo, limit, offset];
      countArgs = [dueFrom, dueTo];
    }
  } else if (action === "bills_list") {
    if (!groupId) {
      return billsJson({ ok: false, error: "Missing required query param: group_id" }, 400);
    }
    dataSql = "SELECT bm.* FROM billing_map bm INNER JOIN group_resolution_cache grc ON grc.property_map_id = bm.property_map_id WHERE grc.property_group_id = ? AND (? IS NULL OR lower(bm.line_items_json) LIKE '%' || lower(?) || '%') ORDER BY bm.due_date DESC LIMIT ? OFFSET ?";
    countSql = "SELECT COUNT(*) AS total FROM billing_map bm INNER JOIN group_resolution_cache grc ON grc.property_map_id = bm.property_map_id WHERE grc.property_group_id = ? AND (? IS NULL OR lower(bm.line_items_json) LIKE '%' || lower(?) || '%')";
    dataArgs = [groupId, status, status, limit, offset];
    countArgs = [groupId, status, status];
  } else {
    return billsJson({ ok: false, error: `Unsupported bills route action: ${action}` }, 404);
  }

  try {
    const dataRes = await db.execute({ sql: dataSql, args: dataArgs });
    const totalRes = await db.execute({ sql: countSql, args: countArgs });

    return billsJson({
      ok: true,
      data: rowsAsObjects(dataRes),
      total: asCount(totalRes),
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