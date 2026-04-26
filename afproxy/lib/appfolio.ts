// ============================================================================
// lib/appfolio.ts — AppFolio API wrappers (read + write).
//
// TWO API surfaces:
//   Reports API v2  →  https://flraz.appfolio.com/api/v2/reports/{report}.json
//                       Auth: reportsHeaders() — btoa(clientId:secret)
//   Database API v0 →  https://api.appfolio.com/api/v0/{endpoint}
//                       Auth: dbHeaders() — pre-encoded Base64 token
//
// WRITE RULES:
//   • postWoNote and patchWorkOrder must be called SEQUENTIALLY.
//     Concurrent PATCHes to the same resource will cause the second to fail .
//   • AssignedUsers must reference a user with the Maintenance Tech role or
//     the request returns 422 "User not found" .
//   • Note body field key is "Body" (capital B).
//   • HTTP Basic Auth is used for all requests .
//   • 422 responses are semantic failures — do not retry without changing payload.
//   • Status strings are exact and case-sensitive (e.g. Scheduled, Waiting,
//     Work Completed).
//
// REPORT RULES:
//   • Reports API v2 filters belong in the JSON POST body, not query params.
//   • Paginated responses return rows in `results`; non-paginated reports may
//     return a raw array.
//   • When using paginate_results=false, always send a bounded date range to
//     avoid timeouts on large reports.
//
// VALUE FORMATS:
//   • Dates: YYYY-MM-DD
//   • DateTimes: YYYY-MM-DDTHH:mm:ssZ (UTC)
//   • Amounts: strings with exactly two decimals, e.g. "150.00"
// ============================================================================

import {
  AF_DB,
  AF_REPORTS,
  dbHeaders,
  FETCH_TIMEOUT_MS,
  PAGE_DELAY_MS,
  reportsHeaders,
} from "../config.ts";
import { delay, fetchWithTimeout, v2InitBucket, v2PageBucket } from "./fetchUtils.ts";

export const APPFOLIO_WORK_ORDER_STATUSES = [
  "Scheduled",
  "Waiting",
  "Work Completed",
] as const;

export function isValidAppfolioDate(value: string): boolean {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(value || "").trim());
}

export function isValidAppfolioDateTimeUtc(value: string): boolean {
  return /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/.test(
    String(value || "").trim(),
  );
}

export function isValidAppfolioAmount(value: string): boolean {
  return /^-?\d+\.\d{2}$/.test(String(value || "").trim());
}

export function isValidAppfolioWorkOrderStatus(value: string): boolean {
  return APPFOLIO_WORK_ORDER_STATUSES.includes(
    value as (typeof APPFOLIO_WORK_ORDER_STATUSES)[number],
  );
}

// ── Compliance assertions ─────────────────────────────────────────────────────
// Fail-fast guards that catch call-site misuse before traffic reaches AppFolio.

/**
 * Throws if paginate_results=false is set without bounded date filters.
 * Unbounded full-table scans time out on large datasets and burn rate limit.
 */
function assertReportFiltersCompliant(
  reportName: string,
  filters: Record<string, any>,
): void {
  const noPage = filters.paginate_results === false ||
    String(filters.paginate_results || "").toLowerCase() === "false";
  if (!noPage) return;

  const hasFrom = filters.from_date || filters.start_date ||
    filters.posted_on_gte || filters.start_on_gte;
  const hasTo = filters.to_date || filters.end_date ||
    filters.posted_on_lte || filters.end_on_lte;

  if (!hasFrom || !hasTo) {
    throw new Error(
      `fetchReport(${reportName}): paginate_results=false requires both ` +
        `from_date and to_date — unbounded scans risk timeout and rate-limit exhaustion.`,
    );
  }
}

const _HISTORY_ENDPOINTS = ["bills_history", "work_orders_completed_history"];

/**
 * Throws if a v0 history endpoint path is missing a from_date query param.
 * History endpoints return unbounded result sets without date bounds.
 */
function assertHistoryPathBounded(path: string): void {
  const lower = path.toLowerCase();
  if (
    _HISTORY_ENDPOINTS.some((ep) => lower.includes(ep)) &&
    !/[?&]from_date=/.test(path)
  ) {
    throw new Error(
      `API compliance: history endpoint requires from_date — ` +
        path.substring(0, 120),
    );
  }
}

// ── fetchReport ───────────────────────────────────────────────────────────────
// Reports API v2 — POST to initiate, GET to paginate.
// A valid URL looks like https://vhost.appfolio.com/api/v2/reports/{endpoint} .
// Filters must stay in the POST body. Paginated responses use `results`, while
// non-paginated responses may be returned as a raw array by AppFolio.
export async function fetchReport(
  reportName: string,
  filters: Record<string, any> = {},
): Promise<any[]> {
  assertReportFiltersCompliant(reportName, filters);
  const reportUrl = `${AF_REPORTS}/api/v2/reports/${reportName}.json`;

  const initResp = await fetchWithTimeout(reportUrl, {
    method: "POST",
    headers: { ...reportsHeaders(), "Content-Type": "application/json" },
    body: JSON.stringify(filters),
  }, FETCH_TIMEOUT_MS, v2InitBucket);

  if (!initResp.ok) {
    const errText = await initResp.text().catch(() => "");
    throw new Error(
      `Report ${reportName} failed: ${initResp.status} ${initResp.statusText} — ${
        errText.substring(0, 300)
      }`,
    );
  }

  let data = await initResp.json();
  let allRows: any[] = [];

  if (data.results || data.Results) {
    allRows = allRows.concat(data.results || data.Results);
  }
  else if (Array.isArray(data.data)) {
    allRows = allRows.concat(data.data);
  }
  else if (Array.isArray(data)) {
    allRows = allRows.concat(data);
  }

  let nextUrl: string | null = data.next_page_url || data.next_page_path ||
    null;
  let page = 1;

  while (nextUrl && page < 50) {
    if (page > 1) await delay(PAGE_DELAY_MS);
    const fullUrl = nextUrl.startsWith("http")
      ? nextUrl
      : `${AF_REPORTS}${nextUrl}`;
    const pageResp = await fetchWithTimeout(fullUrl, {
      headers: reportsHeaders(),
    }, FETCH_TIMEOUT_MS, v2PageBucket);
    if (!pageResp.ok) break;
    data = await pageResp.json();
    if (data.results || data.Results) {
      allRows = allRows.concat(data.results || data.Results);
    }
    else if (Array.isArray(data.data)) allRows = allRows.concat(data.data);
    else break;
    nextUrl = data.next_page_url || data.next_page_path || null;
    page++;
  }

  return allRows;
}

// ── fetchDbApi ────────────────────────────────────────────────────────────────
// Database API v0 — REST pagination via next_page_path.
// Base URL for all requests: https://api.appfolio.com/api/v0/{endpoint} .
export async function fetchDbApi(
  path: string,
  maxRecords: number = 9999,
): Promise<any[]> {
  assertHistoryPathBounded(path);
  let allResults: any[] = [];
  let nextPath: string | null = path;
  let pageNum = 0;

  while (nextPath && allResults.length < maxRecords) {
    if (pageNum > 0) await delay(PAGE_DELAY_MS);
    const fullUrl = nextPath.startsWith("http")
      ? nextPath
      : `${AF_DB}${nextPath}`;
    const resp = await fetchWithTimeout(fullUrl, { headers: dbHeaders() });
    if (!resp.ok) {
      const errText = await resp.text().catch(() => "");
      throw new Error(
        `DB API failed: ${resp.status} — ${errText.substring(0, 200)}`,
      );
    }
    const data = await resp.json();
    const results = data.results || data.Results || data.data ||
      (Array.isArray(data) ? data : []);
    allResults = allResults.concat(results);
    nextPath = data.next_page_path || null;
    if (results.length === 0) break;
    pageNum++;
  }

  return allResults.slice(0, maxRecords);
}

// ── postWoNote ────────────────────────────────────────────────────────────────
// POST a system note to a work order.
// Body field key is "Body" (capital B) — required by AppFolio note schema.
// 401 fallback: retries on tenant domain (flraz.appfolio.com) if api.appfolio.com rejects.
export async function postWoNote(
  woId: string,
  noteText: string,
): Promise<boolean> {
  const path = `/api/v0/work_orders/${encodeURIComponent(woId)}/notes`;
  try {
    let resp = await fetchWithTimeout(`${AF_DB}${path}`, {
      method: "POST",
      headers: { ...dbHeaders(), "Content-Type": "application/json" },
      body: JSON.stringify({ Body: noteText }),
    });
    if (resp.status === 401) {
      resp = await fetchWithTimeout(`${AF_REPORTS}${path}`, {
        method: "POST",
        headers: { ...dbHeaders(), "Content-Type": "application/json" },
        body: JSON.stringify({ Body: noteText }),
      });
    }
    if (!resp.ok) {
      const errText = await resp.text().catch(() => "");
      console.log(
        `postWoNote failed for ${woId}: HTTP ${resp.status} — ${
          errText.substring(0, 150)
        }`,
      );
    }
    return resp.ok;
  } catch (e: any) {
    console.log(`postWoNote error for ${woId}: ${e.message}`);
    return false;
  }
}

function normalizeAttachmentList(data: any): any[] {
  if (Array.isArray(data)) return data;
  if (Array.isArray(data?.items)) return data.items;
  if (Array.isArray(data?.Results)) return data.Results;
  if (Array.isArray(data?.results)) return data.results;
  if (Array.isArray(data?.data)) return data.data;
  return [];
}

export async function fetchWorkOrderAttachments(
  woId: string,
): Promise<{
  ok: boolean;
  status?: number;
  attachments?: any[];
  detail?: string;
}> {
  const path = `/api/v0/work_orders/${encodeURIComponent(woId)}/attachments`;
  let resp = await fetchWithTimeout(`${AF_DB}${path}`, { headers: dbHeaders() });
  if ([401, 403, 404, 422].includes(resp.status)) {
    resp = await fetchWithTimeout(`${AF_REPORTS}${path}`, {
      headers: dbHeaders(),
    });
  }

  if (resp.status === 429) {
    const ra = parseInt(resp.headers.get("Retry-After") || "2", 10);
    await new Promise((r) => setTimeout(r, ra * 1000));
    resp = await fetchWithTimeout(`${AF_DB}${path}`, { headers: dbHeaders() });
  }

  if (!resp.ok) {
    const detail = await resp.text().catch(() => "");
    return {
      ok: false,
      status: resp.status,
      detail: detail.substring(0, 400),
    };
  }

  const data = await resp.json().catch(() => []);
  return {
    ok: true,
    status: resp.status,
    attachments: normalizeAttachmentList(data),
  };
}

export async function uploadWorkOrderAttachment(
  woId: string,
  contentType: string,
  bodyBuffer: ArrayBuffer,
): Promise<{
  ok: boolean;
  status: number;
  detail: any;
}> {
  const path = `/api/v0/work_orders/${encodeURIComponent(woId)}/attachments`;
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

  if (resp.status === 429) {
    const ra = parseInt(resp.headers.get("Retry-After") || "2", 10);
    await new Promise((r) => setTimeout(r, ra * 1000));
    resp = await fetchWithTimeout(`${AF_DB}${path}`, {
      method: "POST",
      headers,
      body: bodyBuffer,
    });
  }

  const text = await resp.text().catch(() => "");
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
  };
}

// ── patchWorkOrder ────────────────────────────────────────────────────────────
// PATCH a work order via DB API v0.
//
// CRITICAL:
//   AssignedUsers must reference a user with the Maintenance Tech role.
//   Any other role returns 422 "User not found" .
//   Never fire concurrent PATCHes to the same WO — the second will fail .
//   All cron loops call this sequentially with 200 ms between each WO.
export async function patchWorkOrder(
  woId: string,
  patchBody: Record<string, any>,
): Promise<{ ok: boolean; status?: number; error?: string }> {
  const path = `/api/v0/work_orders/${encodeURIComponent(woId)}`;
  try {
    let resp = await fetchWithTimeout(`${AF_DB}${path}`, {
      method: "PATCH",
      headers: { ...dbHeaders(), "Content-Type": "application/json" },
      body: JSON.stringify(patchBody),
    });
    if (resp.status === 401) {
      resp = await fetchWithTimeout(`${AF_REPORTS}${path}`, {
        method: "PATCH",
        headers: { ...dbHeaders(), "Content-Type": "application/json" },
        body: JSON.stringify(patchBody),
      });
    }
    if (!resp.ok) {
      const detail = await resp.text().catch(() => "");
      console.log(
        `patchWorkOrder failed for ${woId}: HTTP ${resp.status} — ${
          detail.substring(0, 200)
        }`,
      );
      return {
        ok: false,
        status: resp.status,
        error: detail.substring(0, 200),
      };
    }
    return { ok: true, status: resp.status };
  } catch (e: any) {
    console.log(`patchWorkOrder error for ${woId}: ${e.message}`);
    return { ok: false, error: e.message };
  }
}

// ── resolveWebhookResource ────────────────────────────────────────────────────
// Resolves a webhook resource_type + resource_id to a live AppFolio record.
// Tries direct /{id} lookup first, falls back to ?filters[Id]= query.
// Ensure the formatting of the Client ID and Secret in the auth header are
// valid before calling this .
export async function resolveWebhookResource(
  resourceType: string,
  resourceId: string,
): Promise<{
  ok: boolean;
  record?: any;
  domain?: string;
  status?: number;
  detail?: string;
}> {
  const endpointMap: Record<string, string> = {
    work_order: "work_orders",
    unit_turn: "unit_turns",
    inspection: "inspections",
    vendor: "vendors",
    property: "properties",
    tenant: "tenants",
    lease: "leases",
    bill: "bills",
    task: "tasks",
  };

  const collection = endpointMap[resourceType] || `${resourceType}s`;
  const byIdPath = `/api/v0/${collection}/${encodeURIComponent(resourceId)}`;
  const byFilterPath = `/api/v0/${collection}?filters[Id]=${
    encodeURIComponent(resourceId)
  }&page[size]=1`;

  const tryFetch = async (path: string): Promise<{
    ok: boolean;
    record?: any;
    domain?: string;
    status?: number;
    detail?: string;
  }> => {
    let domain = AF_REPORTS;
    let r = await fetchWithTimeout(`${AF_REPORTS}${path}`, {
      headers: dbHeaders(),
    });
    if (r.status === 401 || r.status === 403 || r.status === 404) {
      domain = AF_DB;
      r = await fetchWithTimeout(`${AF_DB}${path}`, { headers: dbHeaders() });
    }
    if (!r.ok) {
      const detail = await r.text().catch(() => "");
      return {
        ok: false,
        domain,
        status: r.status,
        detail: detail.substring(0, 500),
      };
    }
    const payload = await r.json().catch(() => ({}));
    const record =
      (payload?.data && !Array.isArray(payload.data) ? payload.data : null) ||
      (payload?.data && Array.isArray(payload.data) ? payload.data[0] : null) ||
      payload?.results?.[0] ||
      payload?.Results?.[0] ||
      (Array.isArray(payload) ? payload[0] : null) ||
      payload ||
      null;
    return { ok: !!record, record, domain, status: r.status };
  };

  let found = await tryFetch(byIdPath);
  if (!found.ok) found = await tryFetch(byFilterPath);

  return found;
}

type SqlStmt = { sql: string; args?: any[] };
type SqlClient = {
  execute: (stmt: string | SqlStmt) => Promise<any>;
  batch?: (stmts: SqlStmt[]) => Promise<any>;
};

type RoutingContextRow = {
  propertyGroupId: string;
  propertyGroupName: string;
  pmId: string | null;
  pmEmail: string | null;
};

function isRateLimitedError(err: any): boolean {
  return !!(err && typeof err === "object" && err.rateLimited === true);
}

function asArray<T = any>(payload: any): T[] {
  if (Array.isArray(payload?.Results)) return payload.Results;
  if (Array.isArray(payload?.results)) return payload.results;
  if (Array.isArray(payload?.data)) return payload.data;
  if (Array.isArray(payload)) return payload;
  return [];
}

function firstNonEmpty(...vals: any[]): string {
  for (const v of vals) {
    const s = String(v ?? "").trim();
    if (s) return s;
  }
  return "";
}

function workOrderIsTerminal(status: string): boolean {
  const s = String(status || "").trim();
  return s === "Completed" || s === "Work Completed" || s === "Canceled";
}

function isTurnCategory(category: string): number {
  const c = String(category || "").trim().toLowerCase();
  return ["make ready", "turn", "unit turn", "turnover"].includes(c)
    ? 1
    : 0;
}

async function executeBatchSequential(
  db: SqlClient,
  statements: SqlStmt[],
): Promise<void> {
  if (statements.length === 0) return;
  if (typeof db.batch === "function") {
    await db.batch(statements);
    return;
  }
  for (const stmt of statements) {
    await db.execute(stmt);
  }
}

async function fetchDbPathStrict(path: string): Promise<any> {
  const url = path.startsWith("http") ? path : `${AF_DB}${path}`;
  const resp = await fetchWithTimeout(url, { headers: dbHeaders() });
  if (resp.status === 429 || resp.status === 533) throw { rateLimited: true };
  if (!resp.ok) {
    const text = await resp.text().catch(() => "");
    throw {
      status: resp.status,
      message: `DB API ${resp.status}: ${text.substring(0, 200)}`,
    };
  }
  return await resp.json().catch(() => ({}));
}

async function paginateDbPath(path: string): Promise<any[]> {
  const rows: any[] = [];
  let nextPath: string | null = path;
  while (nextPath) {
    const payload = await fetchDbPathStrict(nextPath);
    const chunk = asArray(payload);
    rows.push(...chunk);
    const np = String(payload?.next_page_path || "").trim();
    nextPath = np || null;
    if (chunk.length === 0) break;
    if (nextPath) await delay(PAGE_DELAY_MS);
  }
  return rows;
}

async function querySingleValue(
  db: SqlClient,
  sql: string,
  args: any[] = [],
): Promise<any> {
  const res = await db.execute({ sql, args });
  const row = Array.isArray(res?.rows) && res.rows.length > 0 ? res.rows[0] : null;
  if (!row) return null;
  if (Array.isArray(row)) return row[0] ?? null;
  const cols = Array.isArray(res?.columns) ? res.columns : [];
  return cols.length ? row[cols[0]] ?? null : null;
}

// PropertyIds from AppFolio is an authoritative replacement list.
// An empty array clears all memberships. We never diff or merge.
// Delete all existing members for this group, then insert fresh.
export async function syncPropertyGroups(db: SqlClient): Promise<void> {
  try {
    const groups = await paginateDbPath(
      "/api/v0/property_groups?filters[LastUpdatedAtFrom]=2024-01-01T00:00:00Z&page[size]=100",
    );
    const now = Date.now();

    for (const g of groups) {
      const groupId = firstNonEmpty(g?.Id, g?.id);
      if (!groupId) continue;
      const name = firstNonEmpty(g?.Name, g?.name, groupId);
      const ids = Array.isArray(g?.PropertyIds)
        ? g.PropertyIds.map((v: any) => String(v || "").trim()).filter(Boolean)
        : [];

      await db.execute({
        sql:
          `INSERT OR REPLACE INTO property_group_map (id, name, property_ids_json, cached_at, last_membership_sync)
           VALUES (?, ?, ?, ?, ?)`,
        args: [groupId, name, JSON.stringify(ids), now, now],
      });

      const members = ids.map((pid: string) => ({
        sql:
          `INSERT OR IGNORE INTO property_group_members (property_group_id, property_map_id) VALUES (?, ?)`,
        args: [groupId, pid],
      }));

      const chunks: SqlStmt[][] = [];
      chunks.push([
        {
          sql: `DELETE FROM property_group_members WHERE property_group_id = ?`,
          args: [groupId],
        },
      ]);
      for (let i = 0; i < members.length; i += 50) {
        chunks.push(members.slice(i, i + 50));
      }

      for (const chunk of chunks) {
        await executeBatchSequential(db, chunk);
      }
    }
  } catch (e: any) {
    if (isRateLimitedError(e)) throw e;
    throw new Error("[syncPropertyGroups] " + (e as Error).message);
  }
}

export async function syncPropertyMap(
  db: SqlClient,
  propertyGroupId?: string,
): Promise<void> {
  try {
    const suffix = propertyGroupId
      ? `&property_group_id=${encodeURIComponent(propertyGroupId)}`
      : "";
    const properties = await paginateDbPath(
      `/api/v0/properties?filters[LastUpdatedAtFrom]=2024-01-01T00:00:00Z&page[size]=100${suffix}`,
    );
    const now = Date.now();

    for (const prop of properties) {
      const propertyId = firstNonEmpty(prop?.Id, prop?.id);
      if (!propertyId) continue;

      const unitIds = Array.isArray(prop?.UnitIds)
        ? prop.UnitIds.map((v: any) => String(v || "").trim()).filter(Boolean)
        : [];
      const inlineUnits = Array.isArray(prop?.Units) ? prop.Units : [];
      const multifamily = unitIds.length > 0 || inlineUnits.length > 0;

      let units: any[] = [];
      if (multifamily) {
        try {
          units = await paginateDbPath(
            `/api/v0/units?property_id=${encodeURIComponent(propertyId)}&page[size]=100`,
          );
        } catch (err: any) {
          if (err?.status === 403 || err?.status === 422) {
            console.error(`syncPropertyMap skip property ${propertyId}: HTTP ${err.status}`);
            continue;
          }
          throw err;
        }
      }

      const statements: SqlStmt[] = [];
      if (units.length > 0) {
        for (const unit of units) {
          const unitId = firstNonEmpty(unit?.Id, unit?.id);
          if (!unitId) continue;
          statements.push({
            sql:
              `INSERT OR REPLACE INTO property_map
               (id, property_id, unit_id, is_unit, property_group_id, property_name, unit_name, address, city, state, zip, cached_at, last_sync_at)
               VALUES (?, ?, ?, 1, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
            args: [
              unitId,
              propertyId,
              unitId,
              firstNonEmpty(prop?.PropertyGroupId, prop?.property_group_id, propertyGroupId || ""),
              firstNonEmpty(prop?.Name, prop?.name),
              firstNonEmpty(unit?.Name, unit?.name, unit?.UnitName, unit?.unit_name),
              firstNonEmpty(unit?.Address, unit?.address, prop?.Address, prop?.address),
              firstNonEmpty(unit?.City, unit?.city, prop?.City, prop?.city),
              firstNonEmpty(unit?.State, unit?.state, prop?.State, prop?.state, "AZ"),
              firstNonEmpty(unit?.Zip, unit?.zip, prop?.Zip, prop?.zip),
              now,
              now,
            ],
          });
        }
      } else {
        statements.push({
          sql:
            `INSERT OR REPLACE INTO property_map
             (id, property_id, unit_id, is_unit, property_group_id, property_name, unit_name, address, city, state, zip, cached_at, last_sync_at)
             VALUES (?, ?, NULL, 0, ?, ?, NULL, ?, ?, ?, ?, ?, ?)`,
          args: [
            propertyId,
            propertyId,
            firstNonEmpty(prop?.PropertyGroupId, prop?.property_group_id, propertyGroupId || ""),
            firstNonEmpty(prop?.Name, prop?.name),
            firstNonEmpty(prop?.Address, prop?.address),
            firstNonEmpty(prop?.City, prop?.city),
            firstNonEmpty(prop?.State, prop?.state, "AZ"),
            firstNonEmpty(prop?.Zip, prop?.zip),
            now,
            now,
          ],
        });
      }

      for (let i = 0; i < statements.length; i += 50) {
        await executeBatchSequential(db, statements.slice(i, i + 50));
      }
    }
  } catch (e: any) {
    if (isRateLimitedError(e)) throw e;
    throw new Error("[syncPropertyMap] " + (e as Error).message);
  }
}

export async function syncVendorMap(db: SqlClient): Promise<void> {
  try {
    const vendors = await paginateDbPath(
      "/api/v0/vendors?filters[LastUpdatedAtFrom]=2024-01-01T00:00:00Z&page[size]=100",
    );
    const now = Date.now();
    const statements: SqlStmt[] = vendors.map((v: any) => ({
      sql:
        `INSERT OR REPLACE INTO vendor_map
         (id, name, company_name, email, phone, license_number, insurance_expiry, cached_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
      args: [
        firstNonEmpty(v?.Id, v?.id),
        firstNonEmpty(v?.Name, v?.name, v?.CompanyName, v?.company_name),
        firstNonEmpty(v?.CompanyName, v?.company_name),
        firstNonEmpty(v?.Email, v?.email),
        firstNonEmpty(v?.Phone, v?.phone),
        firstNonEmpty(v?.LicenseNumber, v?.license_number),
        firstNonEmpty(v?.InsuranceExpiry, v?.insurance_expiry),
        now,
      ],
    })).filter((s: SqlStmt) => String(s.args?.[0] || "").trim().length > 0);

    for (let i = 0; i < statements.length; i += 50) {
      await executeBatchSequential(db, statements.slice(i, i + 50));
    }
  } catch (e: any) {
    if (isRateLimitedError(e)) throw e;
    throw new Error("[syncVendorMap] " + (e as Error).message);
  }
}

export async function syncWorkOrderMap(
  db: SqlClient,
  propertyGroupId?: string,
): Promise<void> {
  try {
    const suffix = propertyGroupId
      ? `&property_group_id=${encodeURIComponent(propertyGroupId)}`
      : "";
    const workOrders = await paginateDbPath(
      `/api/v0/work_orders?filters[LastUpdatedAtFrom]=2024-01-01T00:00:00Z&page[size]=100${suffix}`,
    );
    const now = Date.now();

    for (const wo of workOrders) {
      const id = firstNonEmpty(wo?.Id, wo?.id);
      if (!id) continue;
      const propertyId = firstNonEmpty(wo?.PropertyId, wo?.property_id);
      const unitId = firstNonEmpty(wo?.UnitId, wo?.unit_id);
      const mapId = firstNonEmpty(unitId, propertyId);
      if (!mapId) continue;

      const vendorId = firstNonEmpty(wo?.VendorId, wo?.vendor_id);
      const status = firstNonEmpty(wo?.Status, wo?.status);
      const priority = firstNonEmpty(wo?.Priority, wo?.priority);
      const category = firstNonEmpty(wo?.Category, wo?.category);
      const isTurnWo = isTurnCategory(category);
      let turnTrackingUuid: string | null = null;

      if (isTurnWo === 1) {
        const q = await db.execute({
          sql:
            `SELECT tracking_uuid
             FROM unit_turn_tracker
             WHERE closed_at IS NULL
               AND (unit_id = ? OR property_id = ?)
             ORDER BY created_at DESC
             LIMIT 1`,
          args: [unitId || "", propertyId || ""],
        });
        const row = Array.isArray(q?.rows) && q.rows.length > 0 ? q.rows[0] : null;
        if (row) {
          turnTrackingUuid = Array.isArray(row)
            ? String(row[0] || "")
            : String(row.tracking_uuid || "");
          if (!turnTrackingUuid) turnTrackingUuid = null;
        }
      }

      await db.execute({
        sql:
          `INSERT OR REPLACE INTO work_order_map
           (id, work_order_number, property_map_id, vendor_id, occupancy_id, status, priority, category, description,
            is_turn_wo, turn_tracking_uuid, assigned_users_json, created_date, completed_date, last_updated_at, cached_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        args: [
          id,
          firstNonEmpty(wo?.WorkOrderNumber, wo?.work_order_number, id),
          mapId,
          vendorId || null,
          firstNonEmpty(wo?.OccupancyId, wo?.occupancy_id) || null,
          status,
          priority,
          category,
          firstNonEmpty(wo?.Description, wo?.description),
          isTurnWo,
          turnTrackingUuid,
          JSON.stringify(Array.isArray(wo?.AssignedUsers) ? wo.AssignedUsers : []),
          firstNonEmpty(wo?.CreatedDate, wo?.created_date),
          firstNonEmpty(wo?.CompletedDate, wo?.completed_date),
          firstNonEmpty(wo?.LastUpdatedAt, wo?.last_updated_at),
          now,
        ],
      });

      if (!workOrderIsTerminal(status)) {
        const createdDate = firstNonEmpty(wo?.CreatedDate, wo?.created_date);
        const createdMs = Date.parse(createdDate);
        const daysOpen = Number.isFinite(createdMs)
          ? Math.floor((Date.now() - createdMs) / 86400000)
          : null;
        const propertyName = String(
          await querySingleValue(db, `SELECT property_name FROM property_map WHERE id = ? LIMIT 1`, [mapId]) ||
            "",
        );
        const unitName = String(
          await querySingleValue(db, `SELECT unit_name FROM property_map WHERE id = ? LIMIT 1`, [mapId]) || "",
        );
        const address = String(
          await querySingleValue(db, `SELECT address FROM property_map WHERE id = ? LIMIT 1`, [mapId]) || "",
        );
        const propertyGroup = String(
          await querySingleValue(db, `SELECT property_group_id FROM property_map WHERE id = ? LIMIT 1`, [mapId]) || "",
        );
        const vendorName = vendorId
          ? String(await querySingleValue(db, `SELECT name FROM vendor_map WHERE id = ? LIMIT 1`, [vendorId]) || "")
          : "";

        await db.execute({
          sql:
            `INSERT OR REPLACE INTO open_work_orders_view
             (id, work_order_number, property_map_id, property_name, unit_name, address, property_group_id,
              vendor_id, vendor_name, status, priority, category, description, is_turn_wo, turn_tracking_uuid,
              occupancy_id, assigned_users_json, created_date, days_open, cached_at)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
          args: [
            id,
            firstNonEmpty(wo?.WorkOrderNumber, wo?.work_order_number, id),
            mapId,
            propertyName,
            unitName,
            address,
            propertyGroup,
            vendorId || null,
            vendorName,
            status,
            priority,
            category,
            firstNonEmpty(wo?.Description, wo?.description),
            isTurnWo,
            turnTrackingUuid,
            firstNonEmpty(wo?.OccupancyId, wo?.occupancy_id) || null,
            JSON.stringify(Array.isArray(wo?.AssignedUsers) ? wo.AssignedUsers : []),
            createdDate,
            daysOpen,
            now,
          ],
        });
      } else {
        await db.execute({
          sql: `DELETE FROM open_work_orders_view WHERE id = ?`,
          args: [id],
        });
      }
    }
  } catch (e: any) {
    if (isRateLimitedError(e)) throw e;
    throw new Error("[syncWorkOrderMap] " + (e as Error).message);
  }
}

export async function syncBillingMap(
  db: SqlClient,
  propertyGroupId?: string,
): Promise<void> {
  try {
    const suffix = propertyGroupId
      ? `&property_group_id=${encodeURIComponent(propertyGroupId)}`
      : "";
    const bills = await paginateDbPath(
      `/api/v0/bills?filters[LastUpdatedAtFrom]=2024-01-01T00:00:00Z&page[size]=100${suffix}`,
    );
    const now = Date.now();

    for (const bill of bills) {
      const id = firstNonEmpty(bill?.Id, bill?.id);
      if (!id) continue;
      const propertyId = firstNonEmpty(bill?.PropertyId, bill?.property_id);
      const unitId = firstNonEmpty(bill?.UnitId, bill?.unit_id);
      const mapId = firstNonEmpty(unitId, propertyId);
      if (!mapId) continue;

      const exists = await querySingleValue(
        db,
        `SELECT id FROM property_map WHERE id = ? LIMIT 1`,
        [mapId],
      );
      if (!exists) {
        console.error(`syncBillingMap skip bill ${id}: property_map missing ${mapId}`);
        continue;
      }

      await db.execute({
        sql:
          `INSERT OR REPLACE INTO billing_map
           (id, vendor_id, property_map_id, work_order_id, work_order_number, invoice_date, due_date, total_amount,
            check_memo, management_company_payee, line_items_json, last_updated_at, cached_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        args: [
          id,
          firstNonEmpty(bill?.VendorId, bill?.vendor_id) || null,
          mapId,
          firstNonEmpty(bill?.WorkOrderId, bill?.work_order_id) || null,
          firstNonEmpty(bill?.WorkOrderNumber, bill?.work_order_number) || null,
          firstNonEmpty(bill?.InvoiceDate, bill?.invoice_date),
          firstNonEmpty(bill?.DueDate, bill?.due_date),
          Number(bill?.TotalAmount ?? bill?.total_amount ?? 0) || 0,
          firstNonEmpty(bill?.CheckMemo, bill?.check_memo),
          bill?.ManagementCompanyAsPayee ? 1 : 0,
          JSON.stringify(Array.isArray(bill?.LineItems) ? bill.LineItems : []),
          firstNonEmpty(bill?.LastUpdatedAt, bill?.last_updated_at),
          now,
        ],
      });
    }
  } catch (e: any) {
    if (isRateLimitedError(e)) throw e;
    throw new Error("[syncBillingMap] " + (e as Error).message);
  }
}

export async function rebuildGroupResolutionCache(
  db: SqlClient,
  propertyGroupId?: string,
): Promise<void> {
  try {
    if (propertyGroupId) {
      await db.execute({
        sql: `DELETE FROM group_resolution_cache WHERE property_group_id = ?`,
        args: [propertyGroupId],
      });
    } else {
      await db.execute(`DELETE FROM group_resolution_cache`);
    }

    await db.execute({
      sql:
        `INSERT OR REPLACE INTO group_resolution_cache
           (property_map_id, property_id, unit_id, is_unit, property_group_id, property_group_name, pm_id, pm_email, resolved_at)
         SELECT
           pm.id,
           pm.property_id,
           pm.unit_id,
           pm.is_unit,
           pgm.property_group_id,
           pg.name,
           pma.pm_id,
           pma.pm_email,
           unixepoch()
         FROM property_map pm
         JOIN property_group_members pgm ON pm.id = pgm.property_map_id
         JOIN property_group_map pg ON pgm.property_group_id = pg.id
         LEFT JOIN pm_assignments pma
           ON pma.property_group_id = pgm.property_group_id
          AND pma.is_active = 1
         WHERE (? IS NULL OR pgm.property_group_id = ?)`,
      args: [propertyGroupId || null, propertyGroupId || null],
    });
  } catch (e: any) {
    throw new Error("[rebuildGroupResolutionCache] " + (e as Error).message);
  }
}

export async function processWebhookSync(
  db: SqlClient,
  eventType: string,
  payload: unknown,
  eventId: number,
): Promise<void> {
  try {
    const evt = String(eventType || "").trim();
    const p = (payload && typeof payload === "object") ? (payload as Record<string, any>) : {};
    const data = (p.data && typeof p.data === "object") ? p.data : p;
    const now = Date.now();

    if (["property_group.created", "property_group.updated"].includes(evt)) {
      const id = firstNonEmpty(data?.Id, data?.id);
      const name = firstNonEmpty(data?.Name, data?.name, id);
      const propertyIds = Array.isArray(data?.PropertyIds)
        ? data.PropertyIds.map((v: any) => String(v || "").trim()).filter(Boolean)
        : [];
      await db.execute({
        sql:
          `INSERT OR REPLACE INTO property_group_map (id, name, property_ids_json, cached_at, last_membership_sync)
           VALUES (?, ?, ?, ?, ?)`,
        args: [id, name, JSON.stringify(propertyIds), now, now],
      });
      const stmts: SqlStmt[] = [{
        sql: `DELETE FROM property_group_members WHERE property_group_id = ?`,
        args: [id],
      }].concat(propertyIds.map((pid: string) => ({
        sql: `INSERT OR IGNORE INTO property_group_members (property_group_id, property_map_id) VALUES (?, ?)`,
        args: [id, pid],
      })));
      for (let i = 0; i < stmts.length; i += 50) {
        await executeBatchSequential(db, stmts.slice(i, i + 50));
      }
      await rebuildGroupResolutionCache(db, id);
    } else if (evt === "property_group.deleted") {
      const id = firstNonEmpty(data?.Id, data?.id);
      await db.execute({ sql: `DELETE FROM property_group_map WHERE id = ?`, args: [id] });
      await db.execute({ sql: `DELETE FROM property_group_members WHERE property_group_id = ?`, args: [id] });
      await db.execute({ sql: `DELETE FROM group_resolution_cache WHERE property_group_id = ?`, args: [id] });
    } else if (["work_order.created", "work_order.updated"].includes(evt)) {
      const id = firstNonEmpty(data?.Id, data?.id);
      const propertyId = firstNonEmpty(data?.PropertyId, data?.property_id);
      const unitId = firstNonEmpty(data?.UnitId, data?.unit_id);
      const mapId = firstNonEmpty(unitId, propertyId);
      const status = firstNonEmpty(data?.Status, data?.status);
      const category = firstNonEmpty(data?.Category, data?.category);
      const priority = firstNonEmpty(data?.Priority, data?.priority);
      const vendorId = firstNonEmpty(data?.VendorId, data?.vendor_id);
      const isTurnWo = isTurnCategory(category);

      let turnTrackingUuid: string | null = null;
      if (isTurnWo === 1) {
        const q = await db.execute({
          sql:
            `SELECT tracking_uuid
             FROM unit_turn_tracker
             WHERE closed_at IS NULL
               AND (unit_id = ? OR property_id = ?)
             ORDER BY created_at DESC
             LIMIT 1`,
          args: [unitId || "", propertyId || ""],
        });
        const r = Array.isArray(q?.rows) && q.rows.length > 0 ? q.rows[0] : null;
        if (r) {
          turnTrackingUuid = Array.isArray(r)
            ? String(r[0] || "")
            : String(r.tracking_uuid || "");
          if (!turnTrackingUuid) turnTrackingUuid = null;
        }
      }

      await db.execute({
        sql:
          `INSERT OR REPLACE INTO work_order_map
           (id, work_order_number, property_map_id, vendor_id, occupancy_id, status, priority, category, description,
            is_turn_wo, turn_tracking_uuid, assigned_users_json, created_date, completed_date, last_updated_at, cached_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        args: [
          id,
          firstNonEmpty(data?.WorkOrderNumber, data?.work_order_number, id),
          mapId,
          vendorId || null,
          firstNonEmpty(data?.OccupancyId, data?.occupancy_id) || null,
          status,
          priority,
          category,
          firstNonEmpty(data?.Description, data?.description),
          isTurnWo,
          turnTrackingUuid,
          JSON.stringify(Array.isArray(data?.AssignedUsers) ? data.AssignedUsers : []),
          firstNonEmpty(data?.CreatedDate, data?.created_date),
          firstNonEmpty(data?.CompletedDate, data?.completed_date),
          firstNonEmpty(data?.LastUpdatedAt, data?.last_updated_at),
          now,
        ],
      });

      if (workOrderIsTerminal(status)) {
        await db.execute({
          sql: `DELETE FROM open_work_orders_view WHERE id = ?`,
          args: [id],
        });
      } else {
        const propertyName = String(
          await querySingleValue(db, `SELECT property_name FROM property_map WHERE id = ? LIMIT 1`, [mapId]) || "",
        );
        const unitName = String(
          await querySingleValue(db, `SELECT unit_name FROM property_map WHERE id = ? LIMIT 1`, [mapId]) || "",
        );
        const address = String(
          await querySingleValue(db, `SELECT address FROM property_map WHERE id = ? LIMIT 1`, [mapId]) || "",
        );
        const propertyGroup = String(
          await querySingleValue(db, `SELECT property_group_id FROM property_map WHERE id = ? LIMIT 1`, [mapId]) || "",
        );
        const vendorName = vendorId
          ? String(await querySingleValue(db, `SELECT name FROM vendor_map WHERE id = ? LIMIT 1`, [vendorId]) || "")
          : "";
        const createdDate = firstNonEmpty(data?.CreatedDate, data?.created_date);
        const createdMs = Date.parse(createdDate);
        const daysOpen = Number.isFinite(createdMs)
          ? Math.floor((Date.now() - createdMs) / 86400000)
          : null;

        await db.execute({
          sql:
            `INSERT OR REPLACE INTO open_work_orders_view
             (id, work_order_number, property_map_id, property_name, unit_name, address, property_group_id,
              vendor_id, vendor_name, status, priority, category, description, is_turn_wo, turn_tracking_uuid,
              occupancy_id, assigned_users_json, created_date, days_open, cached_at)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
          args: [
            id,
            firstNonEmpty(data?.WorkOrderNumber, data?.work_order_number, id),
            mapId,
            propertyName,
            unitName,
            address,
            propertyGroup,
            vendorId || null,
            vendorName,
            status,
            priority,
            category,
            firstNonEmpty(data?.Description, data?.description),
            isTurnWo,
            turnTrackingUuid,
            firstNonEmpty(data?.OccupancyId, data?.occupancy_id) || null,
            JSON.stringify(Array.isArray(data?.AssignedUsers) ? data.AssignedUsers : []),
            createdDate,
            daysOpen,
            now,
          ],
        });
      }
    } else if (evt === "work_order.deleted") {
      const id = firstNonEmpty(data?.Id, data?.id);
      await db.execute({ sql: `DELETE FROM work_order_map WHERE id = ?`, args: [id] });
      await db.execute({ sql: `DELETE FROM open_work_orders_view WHERE id = ?`, args: [id] });
    } else if (["bill.created", "bill.updated"].includes(evt)) {
      const id = firstNonEmpty(data?.Id, data?.id);
      const propertyId = firstNonEmpty(data?.PropertyId, data?.property_id);
      const unitId = firstNonEmpty(data?.UnitId, data?.unit_id);
      const mapId = firstNonEmpty(unitId, propertyId);
      await db.execute({
        sql:
          `INSERT OR REPLACE INTO billing_map
           (id, vendor_id, property_map_id, work_order_id, work_order_number, invoice_date, due_date, total_amount,
            check_memo, management_company_payee, line_items_json, last_updated_at, cached_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        args: [
          id,
          firstNonEmpty(data?.VendorId, data?.vendor_id) || null,
          mapId,
          firstNonEmpty(data?.WorkOrderId, data?.work_order_id) || null,
          firstNonEmpty(data?.WorkOrderNumber, data?.work_order_number) || null,
          firstNonEmpty(data?.InvoiceDate, data?.invoice_date),
          firstNonEmpty(data?.DueDate, data?.due_date),
          Number(data?.TotalAmount ?? data?.total_amount ?? 0) || 0,
          firstNonEmpty(data?.CheckMemo, data?.check_memo),
          data?.ManagementCompanyAsPayee ? 1 : 0,
          JSON.stringify(Array.isArray(data?.LineItems) ? data.LineItems : []),
          firstNonEmpty(data?.LastUpdatedAt, data?.last_updated_at),
          now,
        ],
      });
    } else if (evt === "bill.deleted") {
      const id = firstNonEmpty(data?.Id, data?.id);
      await db.execute({ sql: `DELETE FROM billing_map WHERE id = ?`, args: [id] });
    } else if (["vendor.created", "vendor.updated"].includes(evt)) {
      const id = firstNonEmpty(data?.Id, data?.id);
      await db.execute({
        sql:
          `INSERT OR REPLACE INTO vendor_map
           (id, name, company_name, email, phone, license_number, insurance_expiry, cached_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
        args: [
          id,
          firstNonEmpty(data?.Name, data?.name, data?.CompanyName, data?.company_name),
          firstNonEmpty(data?.CompanyName, data?.company_name),
          firstNonEmpty(data?.Email, data?.email),
          firstNonEmpty(data?.Phone, data?.phone),
          firstNonEmpty(data?.LicenseNumber, data?.license_number),
          firstNonEmpty(data?.InsuranceExpiry, data?.insurance_expiry),
          now,
        ],
      });
    } else if (evt === "vendor.deleted") {
      const id = firstNonEmpty(data?.Id, data?.id);
      await db.execute({ sql: `DELETE FROM vendor_map WHERE id = ?`, args: [id] });
    }

    await db.execute({
      sql: `UPDATE webhook_events SET processed = 1, processed_at = datetime('now') WHERE id = ?`,
      args: [eventId],
    });
  } catch (e: any) {
    if (isRateLimitedError(e)) throw e;
    throw new Error("[processWebhookSync] " + (e as Error).message);
  }
}

// This function is the foundation for all future routing logic.
// When a webhook fires with a PropertyId or UnitId, callers resolve
// the PM and group here before dispatching notifications or updating
// dashboard state — with no AppFolio API dependency.
export async function resolveRoutingContext(
  db: SqlClient,
  propertyMapId: string,
): Promise<RoutingContextRow[]> {
  try {
    const fast = await db.execute({
      sql:
        `SELECT property_group_id, property_group_name, pm_id, pm_email
         FROM group_resolution_cache
         WHERE property_map_id = ?`,
      args: [propertyMapId],
    });
    if (Array.isArray(fast?.rows) && fast.rows.length > 0) {
      return fast.rows.map((r: any) => ({
        propertyGroupId: Array.isArray(r) ? String(r[0] || "") : String(r.property_group_id || ""),
        propertyGroupName: Array.isArray(r) ? String(r[1] || "") : String(r.property_group_name || ""),
        pmId: Array.isArray(r) ? (r[2] ? String(r[2]) : null) : (r.pm_id ? String(r.pm_id) : null),
        pmEmail: Array.isArray(r) ? (r[3] ? String(r[3]) : null) : (r.pm_email ? String(r.pm_email) : null),
      }));
    }

    const slow = await db.execute({
      sql:
        `SELECT pgm.property_group_id, pg.name, pma.pm_id, pma.pm_email
         FROM property_group_members pgm
         JOIN property_group_map pg ON pgm.property_group_id = pg.id
         LEFT JOIN pm_assignments pma ON pma.property_group_id = pgm.property_group_id AND pma.is_active = 1
         WHERE pgm.property_map_id = ?`,
      args: [propertyMapId],
    });
    return (Array.isArray(slow?.rows) ? slow.rows : []).map((r: any) => ({
      propertyGroupId: Array.isArray(r) ? String(r[0] || "") : String(r.property_group_id || ""),
      propertyGroupName: Array.isArray(r) ? String(r[1] || "") : String(r.name || ""),
      pmId: Array.isArray(r) ? (r[2] ? String(r[2]) : null) : (r.pm_id ? String(r.pm_id) : null),
      pmEmail: Array.isArray(r) ? (r[3] ? String(r[3]) : null) : (r.pm_email ? String(r.pm_email) : null),
    }));
  } catch (e: any) {
    throw new Error("[resolveRoutingContext] " + (e as Error).message);
  }
}

export async function resolveGroupsForProperty(
  db: SqlClient,
  propertyMapId: string,
): Promise<string[]> {
  try {
    const fast = await db.execute({
      sql:
        `SELECT property_group_id FROM group_resolution_cache WHERE property_map_id = ?`,
      args: [propertyMapId],
    });
    const fastRows = Array.isArray(fast?.rows) ? fast.rows : [];
    if (fastRows.length > 0) {
      return fastRows.map((r: any) => Array.isArray(r) ? String(r[0] || "") : String(r.property_group_id || "")).filter(Boolean);
    }
    const fallback = await db.execute({
      sql: `SELECT property_group_id FROM property_group_members WHERE property_map_id = ?`,
      args: [propertyMapId],
    });
    return (Array.isArray(fallback?.rows) ? fallback.rows : [])
      .map((r: any) => Array.isArray(r) ? String(r[0] || "") : String(r.property_group_id || ""))
      .filter(Boolean);
  } catch (e: any) {
    throw new Error("[resolveGroupsForProperty] " + (e as Error).message);
  }
}

export async function invalidateCacheForGroup(
  db: SqlClient,
  propertyGroupId: string,
): Promise<void> {
  try {
    await db.execute({
      sql: `DELETE FROM open_work_orders_view WHERE property_group_id = ?`,
      args: [propertyGroupId],
    });
    await db.execute({
      sql: `DELETE FROM group_resolution_cache WHERE property_group_id = ?`,
      args: [propertyGroupId],
    });
    await db.execute({
      sql:
        `UPDATE billing_map SET cached_at = 0
         WHERE property_map_id IN (
           SELECT property_map_id FROM property_group_members WHERE property_group_id = ?
         )`,
      args: [propertyGroupId],
    });
    await db.execute({
      sql:
        `UPDATE work_order_map SET cached_at = 0
         WHERE property_map_id IN (
           SELECT property_map_id FROM property_group_members WHERE property_group_id = ?
         )`,
      args: [propertyGroupId],
    });
  } catch (e: any) {
    throw new Error("[invalidateCacheForGroup] " + (e as Error).message);
  }
}