import 'dotenv/config';
import express, { type Request, type Response, type NextFunction } from 'express';
import cors from 'cors';
import path from 'node:path';
import { pingDatabase, queryClient } from './db';

function ensureDenoCompat(): void {
  const globalAny = globalThis as any;

  if (!globalAny.Deno) {
    globalAny.Deno = {};
  }

  if (!globalAny.Deno.env) {
    globalAny.Deno.env = {
      get(name: string): string | undefined {
        return process.env[name];
      },
    };
  }

  if (typeof globalAny.btoa !== 'function') {
    globalAny.btoa = (value: string) => Buffer.from(value, 'binary').toString('base64');
  }

  if (typeof globalAny.atob !== 'function') {
    globalAny.atob = (value: string) => Buffer.from(value, 'base64').toString('binary');
  }
}

ensureDenoCompat();

// @ts-ignore - afproxy uses Deno globals; will be replaced with Postgres reads
const [unitsHandlers, turnsHandlers, estimatesHandlers, queueHandlers, deviceAuthHandlers] = await Promise.all([
  // @ts-ignore
  import('../afproxy/handlers/units.ts'),
  // @ts-ignore
  import('../afproxy/handlers/turns.ts'),
  // @ts-ignore
  import('../afproxy/handlers/estimates.ts'),
  // @ts-ignore
  import('../afproxy/handlers/queue.ts'),
  // @ts-ignore
  import('../afproxy/handlers/deviceAuth.ts'),
]);

type Params = Record<string, string>;
type Mode = 'params' | 'request' | 'params+request';

function toParams(req: Request): Params {
  const out: Params = {};

  Object.entries(req.query || {}).forEach(([key, value]) => {
    if (Array.isArray(value)) {
      out[key] = String(value[0] ?? '');
    } else if (value !== undefined && value !== null) {
      out[key] = String(value);
    }
  });

  Object.entries(req.params || {}).forEach(([key, value]) => {
    out[key] = String(value ?? '');
  });

  return out;
}

function toRequestLike(req: Request): Request {
  const headers = new Headers();
  Object.entries(req.headers || {}).forEach(([key, value]) => {
    if (Array.isArray(value)) {
      headers.set(key, value.join(','));
    } else if (typeof value === 'string') {
      headers.set(key, value);
    }
  });

  const requestLike = {
    method: req.method,
    url: req.originalUrl,
    headers,
    json: async () => req.body ?? {},
    text: async () => (typeof req.body === 'string' ? req.body : JSON.stringify(req.body ?? {})),
  };

  return requestLike as unknown as Request;
}

function logTunnelError(error: unknown, route: string): void {
  const message = String((error as any)?.message || error || 'Unknown error');
  const code = String((error as any)?.code || '');
  console.error(`[server:error] route=${route} code=${code} message=${message}`);

  if (/ECONNREFUSED|ETIMEDOUT|ECONNRESET|socket|connect|pgbouncer|tunnel|connection/i.test(`${code} ${message}`)) {
    console.error('[server:tunnel] Possible PgBouncer/Localtonet tunnel failure detected.', {
      route,
      code,
      message,
    });
  }
}

function wrapDenoHandler(handler: any, mode: Mode = 'params') {
  return async (req: Request, res: Response) => {
    try {
      const params = toParams(req);
      const requestLike = toRequestLike(req);

      let result: any;
      if (mode === 'request') {
        result = await handler(requestLike);
      } else if (mode === 'params+request') {
        result = await handler(params, requestLike);
      } else {
        result = await handler(params);
      }

      const explicitStatus = Number(result?.status);
      const status = Number.isFinite(explicitStatus) && explicitStatus >= 100 && explicitStatus <= 599
        ? explicitStatus
        : (result?.ok === false ? 400 : 200);

      res.status(status).json(result ?? { ok: true });
    } catch (error) {
      logTunnelError(error, req.path);
      res.status(500).json({
        ok: false,
        error: String((error as any)?.message || error || 'Internal server error'),
      });
    }
  };
}

function parseDays(value: unknown, fallback: number, max = 3650): number {
  const parsed = Number.parseInt(String(value ?? ''), 10);
  if (!Number.isFinite(parsed) || parsed <= 0) return fallback;
  return Math.max(1, Math.min(max, parsed));
}

function parseLimit(value: unknown, fallback: number, max = 5000): number {
  const parsed = Number.parseInt(String(value ?? ''), 10);
  if (!Number.isFinite(parsed) || parsed <= 0) return fallback;
  return Math.max(1, Math.min(max, parsed));
}

function parsePropertyGroupId(value: unknown): string {
  const raw = String(value ?? '').trim();
  if (!raw) return '';
  // Accept single values or comma-delimited lists; local endpoints currently scope to one group.
  const first = raw.split(',').map((v) => v.trim()).find(Boolean);
  return String(first || '');
}

function getPropertyGroupFilter(req: Request): string {
  return parsePropertyGroupId(
    req.query.property_group_id
      ?? req.query.property_group_uuid
      ?? req.query.group_id
      ?? req.query.propertyGroupId,
  );
}

function asIso(value: unknown): string {
  if (!value) return '';
  const d = value instanceof Date ? value : new Date(String(value));
  return Number.isNaN(d.getTime()) ? '' : d.toISOString();
}

function pickRaw(obj: any, keys: string[]): any {
  for (const key of keys) {
    const val = obj?.[key];
    if (val !== undefined && val !== null && String(val).trim() !== '') {
      return val;
    }
  }
  return '';
}

function extractSiteManager(raw: any): string {
  const siteManager = pickRaw(raw, ['site_manager', 'siteManager', 'SiteManager']);
  if (typeof siteManager === 'object' && siteManager !== null) {
    const firstName = String(siteManager?.FirstName || siteManager?.first_name || '').trim();
    const lastName = String(siteManager?.LastName || siteManager?.last_name || '').trim();
    return [firstName, lastName].filter(Boolean).join(' ');
  }
  return String(siteManager || '');
}

function normalizePropertyRow(row: any): Record<string, any> {
  const raw = (row?.raw_json && typeof row.raw_json === 'object') ? row.raw_json : {};
  const siteManager = extractSiteManager(raw);
  const visibility = String(pickRaw(raw, ['visibility', 'Visibility']) || '');
  const managementEndDate = String(pickRaw(raw, ['management_end_date', 'managementEndDate', 'ManagementEndDate']) || '');
  const managementEndReason = String(pickRaw(raw, ['management_end_reason', 'managementEndReason', 'ManagementEndReason']) || '');
  
  const normalized: Record<string, any> = {
    ...raw,
    id: String(row?.id || pickRaw(raw, ['property_id', 'id', 'Id']) || ''),
    name: String(row?.name || pickRaw(raw, ['property_name', 'name', 'Name']) || ''),
    address: String(row?.street || pickRaw(raw, ['property_street', 'street', 'Street']) || ''),
    city: String(row?.city || pickRaw(raw, ['property_city', 'city', 'City']) || ''),
    state: String(row?.state || pickRaw(raw, ['property_state', 'state', 'State']) || ''),
    zip: String(row?.zip || pickRaw(raw, ['property_zip', 'zip', 'Zip']) || ''),
    property_type: String(pickRaw(raw, ['property_type', 'type', 'Type']) || ''),
    portfolio_id: String(pickRaw(raw, ['portfolio_id', 'portfolioId', 'PortfolioId']) || ''),
    portfolio: String(pickRaw(raw, ['portfolio', 'portfolio_name', 'portfolioName', 'PortfolioName', 'group_name', 'property_group']) || ''),
    portfolio_name: String(pickRaw(raw, ['portfolio_name', 'portfolioName', 'PortfolioName']) || ''),
    property_group: String(pickRaw(raw, ['property_group', 'group_name', 'group', 'Group']) || ''),
    property_group_id: String(row?.property_group_id || pickRaw(raw, ['property_group_id', 'PropertyGroupId', 'property_group_uuid', 'PropertyGroupUuid']) || ''),
    group: String(pickRaw(raw, ['group', 'Group']) || ''),
    group_name: String(pickRaw(raw, ['group_name', 'groupName', 'GroupName']) || ''),
    maintenance_limit: String(pickRaw(raw, ['maintenance_limit', 'maintenanceLimit', 'MaintenanceLimit']) || ''),
    maintenance_notes: String(pickRaw(raw, ['maintenance_notes', 'maintenanceNotes', 'MaintenanceNotes']) || ''),
    site_manager: siteManager,
    visibility,
    management_end_date: managementEndDate,
    management_end_reason: managementEndReason,
    units: pickRaw(raw, ['units', 'Units']),
    sqft: pickRaw(raw, ['sqft', 'Sqft', 'SquareFeet']),
    market_rent: pickRaw(raw, ['market_rent', 'marketRent', 'MarketRent']),
    owners: pickRaw(raw, ['owners', 'Owners']),
    link: '',
    _source: 'postgres_local',
  };
  return normalized;
}

function isManagedPropertyState(prop: Record<string, any>): boolean {
  const visibility = String(prop.visibility || '').trim().toLowerCase();
  if (visibility && visibility !== 'active') return false;

  const endDate = prop.management_end_date ? new Date(prop.management_end_date) : null;
  if (endDate && !Number.isNaN(endDate.getTime())) {
    const now = new Date();
    now.setHours(0, 0, 0, 0);
    endDate.setHours(0, 0, 0, 0);
    if (endDate <= now) return false;
  }

  return true;
}

function normalizeWorkOrderRow(row: any): Record<string, any> {
  const raw = (row?.raw_json && typeof row.raw_json === 'object') ? row.raw_json : {};
  const createdAt = asIso(row?.created_at) || String(pickRaw(raw, ['CreatedAt', 'created_at', 'created_date']) || '');
  const updatedAt = asIso(row?.updated_at) || String(pickRaw(raw, ['LastUpdatedAt', 'last_updated_at', 'updated_at']) || '');
  const status = String(row?.status || pickRaw(raw, ['Status', 'status']) || '');
  const normalized: Record<string, any> = {
    ...raw,
    db_api_id: String(row?.id || pickRaw(raw, ['db_api_id', 'dbApiId', 'UUID', 'Id']) || ''),
    work_order_id: String(row?.id || pickRaw(raw, ['work_order_id', 'WorkOrderId', 'Id']) || ''),
    work_order_number: String(row?.wo_number || pickRaw(raw, ['work_order_number', 'WorkOrderNumber', 'Number']) || ''),
    property_id: String(row?.property_id || pickRaw(raw, ['property_id', 'PropertyId']) || ''),
    unit_id: String(row?.unit_id || pickRaw(raw, ['unit_id', 'UnitId']) || ''),
    property_group_id: String(row?.property_group_id || pickRaw(raw, ['property_group_id', 'PropertyGroupId']) || ''),
    status,
    priority: String(row?.priority || pickRaw(raw, ['priority', 'Priority']) || ''),
    vendor_id: String(row?.vendor_id || pickRaw(raw, ['vendor_id', 'VendorId']) || ''),
    vendor_name: String(row?.vendor_name || pickRaw(raw, ['vendor_name', 'VendorName']) || ''),
    vendor: String(row?.vendor_name || pickRaw(raw, ['vendor_name', 'VendorName']) || ''),
    description: String(row?.description || pickRaw(raw, ['description', 'Description', 'subject', 'Subject']) || ''),
    created_at: createdAt,
    completed_on: updatedAt,
    work_completed_on: updatedAt,
    estimated_amount: row?.estimated_amount,
    total_cost: row?.total_cost,
    _source: 'postgres_local',
  };
  return normalized;
}

function normalizeUnitRow(row: any): Record<string, any> {
  const raw = (row?.raw_json && typeof row.raw_json === 'object') ? row.raw_json : {};
  const normalized: Record<string, any> = {
    ...raw,
    unit_id: String(row?.unit_id || pickRaw(raw, ['unit_id', 'UnitId', 'Id']) || ''),
    property_id: String(row?.property_id || pickRaw(raw, ['property_id', 'PropertyId']) || ''),
    name: String(row?.name || pickRaw(raw, ['name', 'Name', 'unit_name', 'UnitName']) || ''),
    unit_number: String(row?.unit_number || pickRaw(raw, ['unit_number', 'UnitNumber', 'Number']) || ''),
    status: String(row?.status || pickRaw(raw, ['status', 'Status']) || 'Active'),
    bedrooms: row?.bedrooms || pickRaw(raw, ['bedrooms', 'Bedrooms']) || 0,
    bathrooms: row?.bathrooms || pickRaw(raw, ['bathrooms', 'Bathrooms']) || 0,
    square_feet: row?.square_feet || pickRaw(raw, ['square_feet', 'squareFeet', 'SquareFeet']) || 0,
    market_rent: row?.market_rent || pickRaw(raw, ['market_rent', 'marketRent', 'MarketRent']) || 0,
    type: String(pickRaw(raw, ['type', 'Type', 'unit_type', 'UnitType']) || ''),
    lease_status: String(pickRaw(raw, ['lease_status', 'leaseStatus', 'LeaseStatus']) || ''),
    tenant_name: String(pickRaw(raw, ['tenant_name', 'tenantName', 'TenantName', 'tenant', 'Tenant']) || ''),
    lease_start: String(pickRaw(raw, ['lease_start', 'leaseStart', 'LeaseStart']) || ''),
    lease_end: String(pickRaw(raw, ['lease_end', 'leaseEnd', 'LeaseEnd']) || ''),
    _source: 'postgres_local',
  };
  return normalized;
}
const app = express();
const DIST_DIR = path.resolve(process.cwd(), 'dist');

function getLegacyAction(req: Request): string {
  const fromQuery = String(req.query?.action ?? '').trim();
  if (fromQuery) return fromQuery;
  const body = (req.body && typeof req.body === 'object') ? req.body as Record<string, unknown> : null;
  const fromBody = String(body?.action ?? '').trim();
  return fromBody;
}

async function respondSessionInfo(req: Request, res: Response): Promise<void> {
  const auth = req.headers.authorization || '';
  const token = auth.startsWith('Bearer ') ? auth.slice(7).trim() : '';
  if (!token) {
    res.status(401).json({ ok: false, authenticated: false, error: 'Missing bearer token' });
    return;
  }

  const session = await deviceAuthHandlers.getTrustedDeviceSession(token);
  if (!session) {
    res.status(401).json({ ok: false, authenticated: false, error: 'Invalid session' });
    return;
  }

  res.json({ ok: true, authenticated: true, session });
}

app.use(express.json({ limit: '10mb' }));

const allowedOrigins = new Set([
  'https://handymgr.app',
  'https://handymgr2.onrender.com',
  'http://localhost:5173',
]);

[
  process.env.RENDER_EXTERNAL_URL,
  process.env.APP_ORIGIN,
  process.env.FRONTEND_ORIGIN,
].forEach((origin) => {
  const value = String(origin || '').trim();
  if (value) allowedOrigins.add(value);
});

app.use(cors({
  origin(origin, callback) {
    if (!origin || allowedOrigins.has(origin)) {
      callback(null, true);
      return;
    }
    callback(new Error(`Origin not allowed by CORS: ${origin}`));
  },
  methods: ['GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With'],
  credentials: true,
}));

const legacyActionRoutes = {
  units: wrapDenoHandler(unitsHandlers.handleUnits, 'params'),
  unit_lookup: wrapDenoHandler(unitsHandlers.handleUnitLookup, 'params'),
  turns: wrapDenoHandler(turnsHandlers.handleTurns, 'params'),
  unit_turns: wrapDenoHandler(turnsHandlers.handleUnitTurns, 'params'),
  turns_incremental: wrapDenoHandler(turnsHandlers.handleTurnsIncremental, 'params'),
  unit_turns_history: wrapDenoHandler(turnsHandlers.handleUnitTurnsHistory, 'params'),
  estimates: wrapDenoHandler(estimatesHandlers.handleEstimates, 'params'),
  queue: wrapDenoHandler(queueHandlers.handleReassignmentQueue, 'params'),
  reassignment_queue: wrapDenoHandler(queueHandlers.handleReassignmentQueue, 'params'),
  device_setup: wrapDenoHandler(deviceAuthHandlers.handleDeviceSetup, 'request'),
  device_otp_request: wrapDenoHandler(deviceAuthHandlers.handleDeviceOtpRequest, 'request'),
  device_otp_verify: wrapDenoHandler(deviceAuthHandlers.handleDeviceOtpVerify, 'request'),
  verify_role: wrapDenoHandler(deviceAuthHandlers.handleVerifyRole, 'request'),
} as const;

app.all(['/', '/api', '/api/'], async (req: Request, res: Response, next: NextFunction) => {
  const action = getLegacyAction(req);
  if (!action) {
    next();
    return;
  }

  if (action === 'session_info') {
    try {
      await respondSessionInfo(req, res);
    } catch (error) {
      logTunnelError(error, '/api?action=session_info');
      res.status(500).json({ ok: false, error: 'Session validation failed' });
    }
    return;
  }

  const routeHandler = (legacyActionRoutes as Record<string, any>)[action];
  if (!routeHandler) {
    next();
    return;
  }

  return routeHandler(req, res);
});

app.get('/health', async (_req: Request, res: Response) => {
  const dbOk = await pingDatabase();
  res.status(dbOk ? 200 : 503).json({ ok: dbOk, service: 'handymgr-backend', database: dbOk ? 'up' : 'down' });
});

app.get('/api/local/work_orders', async (req: Request, res: Response) => {
  try {
    const days = parseDays(req.query.days, 180);
    const limit = parseLimit(req.query.limit, 2500);
    const propertyGroupId = getPropertyGroupFilter(req);
    const rows = propertyGroupId
      ? await queryClient`
        select id, wo_number, property_id, unit_id, property_group_id, description,
               category, priority, status, assigned_user_id, assigned_user_name,
               vendor_id, vendor_name, estimated_amount, total_cost,
               created_at, updated_at, raw_json
        from appfolio_work_orders
        where coalesce(updated_at, created_at) >= now() - (${days}::int * interval '1 day')
          and property_group_id = ${propertyGroupId}
          and (
            coalesce(lower(status), '') not like '%completed%'
            and coalesce(lower(status), '') not like '%cancel%'
            and coalesce(lower(status), '') not like '%no need to bill%'
          )
        order by coalesce(updated_at, created_at) desc
        limit ${limit}
      `
      : await queryClient`
        select id, wo_number, property_id, unit_id, property_group_id, description,
               category, priority, status, assigned_user_id, assigned_user_name,
               vendor_id, vendor_name, estimated_amount, total_cost,
               created_at, updated_at, raw_json
        from appfolio_work_orders
        where coalesce(updated_at, created_at) >= now() - (${days}::int * interval '1 day')
          and (
            coalesce(lower(status), '') not like '%completed%'
            and coalesce(lower(status), '') not like '%cancel%'
            and coalesce(lower(status), '') not like '%no need to bill%'
          )
        order by coalesce(updated_at, created_at) desc
        limit ${limit}
      `;

    const results = (rows as any[]).map(normalizeWorkOrderRow);
    res.json({ ok: true, results, count: results.length, source: 'postgres_local' });
  } catch (error) {
    logTunnelError(error, '/api/local/work_orders');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local work orders query failed') });
  }
});

app.get('/api/local/work_orders/inactive', async (req: Request, res: Response) => {
  try {
    const days = parseDays(req.query.days, 3650, 3650);
    const limit = parseLimit(req.query.limit, 5000);
    const propertyGroupId = getPropertyGroupFilter(req);
    const rows = propertyGroupId
      ? await queryClient`
        select id, wo_number, property_id, unit_id, property_group_id, description,
               category, priority, status, assigned_user_id, assigned_user_name,
               vendor_id, vendor_name, estimated_amount, total_cost,
               created_at, updated_at, raw_json
        from appfolio_work_orders
        where coalesce(updated_at, created_at) >= now() - (${days}::int * interval '1 day')
          and property_group_id = ${propertyGroupId}
          and (
            coalesce(lower(status), '') like '%completed%'
            or coalesce(lower(status), '') like '%cancel%'
            or coalesce(lower(status), '') like '%no need to bill%'
          )
        order by coalesce(updated_at, created_at) desc
        limit ${limit}
      `
      : await queryClient`
        select id, wo_number, property_id, unit_id, property_group_id, description,
               category, priority, status, assigned_user_id, assigned_user_name,
               vendor_id, vendor_name, estimated_amount, total_cost,
               created_at, updated_at, raw_json
        from appfolio_work_orders
        where coalesce(updated_at, created_at) >= now() - (${days}::int * interval '1 day')
          and (
            coalesce(lower(status), '') like '%completed%'
            or coalesce(lower(status), '') like '%cancel%'
            or coalesce(lower(status), '') like '%no need to bill%'
          )
        order by coalesce(updated_at, created_at) desc
        limit ${limit}
      `;

    const results = (rows as any[]).map(normalizeWorkOrderRow);
    res.json({ ok: true, results, count: results.length, source: 'postgres_local', status: 'inactive' });
  } catch (error) {
    logTunnelError(error, '/api/local/work_orders/inactive');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local inactive work orders query failed') });
  }
});

app.get('/api/local/properties', async (req: Request, res: Response) => {
  try {
    const limit = parseLimit(req.query.limit, 5000);
    const propertyGroupId = getPropertyGroupFilter(req);
    const rows = propertyGroupId
      ? await queryClient`
        select id, name, property_group_id, street, city, state, zip, raw_json
        from appfolio_properties
        where property_group_id = ${propertyGroupId}
        order by name asc
        limit ${limit}
      `
      : await queryClient`
        select id, name, property_group_id, street, city, state, zip, raw_json
        from appfolio_properties
        order by name asc
        limit ${limit}
      `;

    const results = (rows as any[])
      .map(normalizePropertyRow)
      .filter(isManagedPropertyState);

    res.json({ ok: true, results, count: results.length, source: 'postgres_local' });
  } catch (error) {
    logTunnelError(error, '/api/local/properties');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local properties query failed') });
  }
});

app.get('/api/local/property_groups', async (req: Request, res: Response) => {
  try {
    const limit = parseLimit(req.query.limit, 1000);
    const includeInactive = String(req.query.include_inactive || '').toLowerCase() === 'true';
    const propertyGroupId = getPropertyGroupFilter(req);

    const rows = propertyGroupId
      ? await queryClient`
        select id, name, property_group_id, street, city, state, zip, raw_json
        from appfolio_properties
        where property_group_id = ${propertyGroupId}
      `
      : await queryClient`
        select id, name, property_group_id, street, city, state, zip, raw_json
        from appfolio_properties
      `;

    let properties = (rows as any[]).map(normalizePropertyRow);
    if (!includeInactive) {
      properties = properties.filter(isManagedPropertyState);
    }

    const groupsById = new Map<string, { id: string; name: string; propertyIds: string[] }>();

    for (const prop of properties) {
      const groupId = String(prop.property_group_id || '').trim();
      if (!groupId) continue;

      const existing = groupsById.get(groupId);
      const groupName = String(
        prop.property_group || prop.group_name || prop.portfolio || prop.portfolio_name || groupId,
      ).trim() || groupId;
      const propertyId = String(prop.id || '').trim();

      if (!existing) {
        groupsById.set(groupId, {
          id: groupId,
          name: groupName,
          propertyIds: propertyId ? [propertyId] : [],
        });
      } else if (propertyId && !existing.propertyIds.includes(propertyId)) {
        existing.propertyIds.push(propertyId);
      }
    }

    const results = Array.from(groupsById.values())
      .sort((a, b) => a.name.localeCompare(b.name))
      .slice(0, limit)
      .map((group) => ({
        id: group.id,
        Id: group.id,
        name: group.name,
        Name: group.name,
        type: 'property_group',
        Type: 'property_group',
        property_ids: group.propertyIds,
        PropertyIds: group.propertyIds,
        property_count: group.propertyIds.length,
        _source: 'postgres_local',
      }));

    res.json({ ok: true, results, count: results.length, source: 'postgres_local' });
  } catch (error) {
    logTunnelError(error, '/api/local/property_groups');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local property groups query failed') });
  }
});

app.get('/api/local/units', async (req: Request, res: Response) => {
  try {
    const limit = parseLimit(req.query.limit, 5000);
    const propertyId = req.query.property_id ? String(req.query.property_id) : undefined;
    const propertyGroupId = getPropertyGroupFilter(req);

    if (propertyId && propertyGroupId) {
      const rows = await queryClient`
        select u.unit_id, u.property_id, u.name, u.unit_number, u.status, u.bedrooms, u.bathrooms,
               u.square_feet, u.market_rent, u.raw_json
        from appfolio_units u
        inner join appfolio_properties p on p.id = u.property_id
        where u.property_id = ${propertyId}
          and p.property_group_id = ${propertyGroupId}
        order by u.unit_number asc
        limit ${limit}
      `;
      const results = (rows as any[]).map(normalizeUnitRow);
      res.json({ ok: true, results, count: results.length, source: 'postgres_local' });
    } else if (propertyId) {
      const rows = await queryClient`
        select unit_id, property_id, name, unit_number, status, bedrooms, bathrooms,
               square_feet, market_rent, raw_json
        from appfolio_units
        where property_id = ${propertyId}
        order by unit_number asc
        limit ${limit}
      `;
      const results = (rows as any[]).map(normalizeUnitRow);
      res.json({ ok: true, results, count: results.length, source: 'postgres_local' });
    } else if (propertyGroupId) {
      const rows = await queryClient`
        select u.unit_id, u.property_id, u.name, u.unit_number, u.status, u.bedrooms, u.bathrooms,
               u.square_feet, u.market_rent, u.raw_json
        from appfolio_units u
        inner join appfolio_properties p on p.id = u.property_id
        where p.property_group_id = ${propertyGroupId}
        order by u.unit_number asc
        limit ${limit}
      `;
      const results = (rows as any[]).map(normalizeUnitRow);
      res.json({ ok: true, results, count: results.length, source: 'postgres_local' });
    } else {
      const rows = await queryClient`
        select unit_id, property_id, name, unit_number, status, bedrooms, bathrooms,
               square_feet, market_rent, raw_json
        from appfolio_units
        order by unit_number asc
        limit ${limit}
      `;
      const results = (rows as any[]).map(normalizeUnitRow);
      res.json({ ok: true, results, count: results.length, source: 'postgres_local' });
    }
  } catch (error) {
    logTunnelError(error, '/api/local/units');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local units query failed') });
  }
});
app.get('/api/local/turns', async (req: Request, res: Response) => {
  try {
    const days = parseDays(req.query.days, 90);
    const limit = parseLimit(req.query.limit, 3000);
    const statusFilter = String(req.query.status || '').trim().toLowerCase();

    const rows = await queryClient`
      select
        t.tracking_uuid,
        t.tracking_code,
        t.turn_key,
        t.unit_turn_id,
        t.unit_id,
        t.property_id,
        t.unit_name,
        t.property_name,
        t.status,
        t.confidence_score,
        t.confidence_label,
        t.source_flags,
        t.metadata,
        t.closed_at,
        t.created_at,
        t.updated_at,
        coalesce((
          select jsonb_object_agg(m.milestone_key, jsonb_build_object(
            'date', m.milestone_date,
            'source', m.source,
            'notes', m.notes
          ))
          from unit_turn_milestones m
          where m.tracking_uuid = t.tracking_uuid
        ), '{}'::jsonb) as milestones,
        coalesce((
          select jsonb_agg(jsonb_build_object(
            'wo_id', w.wo_id,
            'wo_db_uuid', w.wo_db_uuid,
            'source', w.source,
            'status', w.status,
            'created_at', w.created_at
          ) order by w.created_at asc)
          from unit_turn_work_orders w
          where w.tracking_uuid = t.tracking_uuid and coalesce(w.removed, false) = false
        ), '[]'::jsonb) as linked_work_orders
      from unit_turn_tracker t
      where coalesce(t.updated_at, t.created_at, now()) >= now() - (${days}::int * interval '1 day')
      order by coalesce(t.updated_at, t.created_at) desc
      limit ${limit}
    `;

    const filtered = (rows as any[]).filter((row) => {
      if (!statusFilter) return true;
      const status = String(row?.status || '').toLowerCase();
      return status.includes(statusFilter);
    });

    const results = filtered.map((row) => ({
      tracking_uuid: row.tracking_uuid,
      tracking_code: row.tracking_code,
      turn_key: row.turn_key,
      unit_turn_id: row.unit_turn_id,
      unit_id: row.unit_id,
      property_id: row.property_id,
      unit_name: row.unit_name,
      property_name: row.property_name,
      status: row.status,
      confidence_score: row.confidence_score,
      confidence_label: row.confidence_label,
      source_flags: row.source_flags || {},
      metadata: row.metadata || {},
      closed_at: asIso(row.closed_at),
      created_at: asIso(row.created_at),
      updated_at: asIso(row.updated_at),
      milestones: row.milestones || {},
      linked_work_orders: Array.isArray(row.linked_work_orders) ? row.linked_work_orders : [],
      _source: 'postgres_local',
    }));

    res.json({ ok: true, results, count: results.length, source: 'postgres_local' });
  } catch (error) {
    logTunnelError(error, '/api/local/turns');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local turns query failed') });
  }
});

app.get('/api/local/turn_work_orders', async (req: Request, res: Response) => {
  try {
    const days = parseDays(req.query.days, 90);
    const limit = parseLimit(req.query.limit, 3000);
    const rows = await queryClient`
      select
        tw.wo_id,
        tw.wo_db_uuid,
        tw.source,
        tw.status as tracker_status,
        tw.created_at as linked_at,
        wo.id as work_order_id,
        wo.wo_number,
        wo.unit_id,
        wo.unit_id as "UnitId",
        wo.unit_id as unit_id,
        wo.unit_id as unit_uuid,
        wo.unit_id as "unit_id",
        wo.property_id,
        wo.status,
        wo.priority,
        wo.category,
        wo.description,
        wo.vendor_id,
        wo.vendor_name,
        wo.created_at,
        wo.updated_at,
        wo.raw_json
      from unit_turn_work_orders tw
      left join appfolio_work_orders wo
        on wo.id = coalesce(tw.wo_db_uuid, tw.wo_id)
           or wo.wo_number = tw.wo_id
      where coalesce(tw.removed, false) = false
        and coalesce(tw.created_at, now()) >= now() - (${days}::int * interval '1 day')
      order by coalesce(wo.updated_at, tw.created_at) desc
      limit ${limit}
    `;

    const results = (rows as any[]).map((row) => {
      const raw = (row?.raw_json && typeof row.raw_json === 'object') ? row.raw_json : {};
      return {
        ...raw,
        Id: row.work_order_id || row.wo_db_uuid || row.wo_id || pickRaw(raw, ['Id', 'id']) || '',
        id: row.work_order_id || row.wo_db_uuid || row.wo_id || pickRaw(raw, ['Id', 'id']) || '',
        WorkOrderNumber: row.wo_number || pickRaw(raw, ['WorkOrderNumber', 'work_order_number']) || row.wo_id || '',
        work_order_number: row.wo_number || pickRaw(raw, ['work_order_number', 'WorkOrderNumber']) || row.wo_id || '',
        UnitTurnId: pickRaw(raw, ['UnitTurnId', 'unit_turn_id']) || '',
        unit_turn_id: pickRaw(raw, ['unit_turn_id', 'UnitTurnId']) || '',
        UnitId: row.unit_id || pickRaw(raw, ['UnitId', 'unit_id']) || '',
        unit_id: row.unit_id || pickRaw(raw, ['unit_id', 'UnitId']) || '',
        PropertyId: row.property_id || pickRaw(raw, ['PropertyId', 'property_id']) || '',
        property_id: row.property_id || pickRaw(raw, ['property_id', 'PropertyId']) || '',
        Status: row.status || row.tracker_status || pickRaw(raw, ['Status', 'status']) || '',
        status: row.status || row.tracker_status || pickRaw(raw, ['status', 'Status']) || '',
        Priority: row.priority || pickRaw(raw, ['Priority', 'priority']) || '',
        priority: row.priority || pickRaw(raw, ['priority', 'Priority']) || '',
        VendorId: row.vendor_id || pickRaw(raw, ['VendorId', 'vendor_id']) || '',
        vendor_id: row.vendor_id || pickRaw(raw, ['vendor_id', 'VendorId']) || '',
        VendorName: row.vendor_name || pickRaw(raw, ['VendorName', 'vendor_name']) || '',
        vendor_name: row.vendor_name || pickRaw(raw, ['vendor_name', 'VendorName']) || '',
        JobDescription: row.description || pickRaw(raw, ['JobDescription', 'Description', 'description']) || '',
        description: row.description || pickRaw(raw, ['description', 'Description', 'JobDescription']) || '',
        LastUpdatedAt: asIso(row.updated_at) || asIso(row.linked_at),
        last_updated_at: asIso(row.updated_at) || asIso(row.linked_at),
        _source: 'postgres_local',
      };
    });

    res.json({ ok: true, results, count: results.length, source: 'postgres_local' });
  } catch (error) {
    logTunnelError(error, '/api/local/turn_work_orders');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local turn work orders query failed') });
  }
});

// Serve built frontend assets when running monolith mode on Render.
app.use(express.static(DIST_DIR, {
  index: 'index.html',
}));

// Units
app.post('/api/units', wrapDenoHandler(unitsHandlers.handleUnits, 'params'));
app.post('/api/unit_lookup', wrapDenoHandler(unitsHandlers.handleUnitLookup, 'params'));

// Turns
app.post('/api/turns', wrapDenoHandler(turnsHandlers.handleTurns, 'params'));
app.post('/api/unit_turns', wrapDenoHandler(turnsHandlers.handleUnitTurns, 'params'));
app.post('/api/turns_incremental', wrapDenoHandler(turnsHandlers.handleTurnsIncremental, 'params'));
app.post('/api/unit_turns_history', wrapDenoHandler(turnsHandlers.handleUnitTurnsHistory, 'params'));

// Estimates
app.post('/api/estimates', wrapDenoHandler(estimatesHandlers.handleEstimates, 'params'));

// Queue
app.post('/api/queue', wrapDenoHandler(queueHandlers.handleReassignmentQueue, 'params'));
app.post('/api/reassignment_queue', wrapDenoHandler(queueHandlers.handleReassignmentQueue, 'params'));

// Device Auth
app.post('/api/device/setup', wrapDenoHandler(deviceAuthHandlers.handleDeviceSetup, 'request'));
app.post('/api/device/otp/request', wrapDenoHandler(deviceAuthHandlers.handleDeviceOtpRequest, 'request'));
app.post('/api/device/otp/verify', wrapDenoHandler(deviceAuthHandlers.handleDeviceOtpVerify, 'request'));
app.post('/api/device/verify-role', wrapDenoHandler(deviceAuthHandlers.handleVerifyRole, 'request'));

// Backward-compatible aliases
app.post('/api/device_setup', wrapDenoHandler(deviceAuthHandlers.handleDeviceSetup, 'request'));
app.post('/api/device_otp_request', wrapDenoHandler(deviceAuthHandlers.handleDeviceOtpRequest, 'request'));
app.post('/api/device_otp_verify', wrapDenoHandler(deviceAuthHandlers.handleDeviceOtpVerify, 'request'));
app.post('/api/verify_role', wrapDenoHandler(deviceAuthHandlers.handleVerifyRole, 'request'));

// ── Sync admin routes ────────────────────────────────────────────────────────
// Trigger a sync run. Protected by INTERNAL_SYNC_TOKEN.
app.post('/api/admin/sync', async (req: Request, res: Response) => {
  try {
    const expectedToken = String(process.env.INTERNAL_SYNC_TOKEN || '').trim();
    const provided = String(req.headers['x-sync-token'] || req.headers['x-cron-secret'] || '').trim();
    if (!expectedToken || provided !== expectedToken) {
      res.status(401).json({ ok: false, error: 'Unauthorized' });
      return;
    }

    const endpointKey = String(req.body?.endpoint ?? 'v0:units').trim();
    const triggerType = String(req.body?.triggerType ?? 'manual').trim();
    const maxPages    = Number(req.body?.maxPages ?? 0);

    // Respond 202 immediately; sync runs async.
    res.status(202).json({ ok: true, accepted: true, endpointKey, triggerType });

    const { runSync } = await import('./sync/syncRunner.ts');
    runSync({ endpointKey, triggerType, maxPages }).catch((err: unknown) => {
      console.error('[server:sync] uncaught sync error', String((err as any)?.message ?? err));
    });
  } catch (error) {
    logTunnelError(error, '/api/admin/sync');
    res.status(500).json({ ok: false, error: String((error as any)?.message ?? 'sync trigger failed') });
  }
});

app.get('/api/session_info', async (req: Request, res: Response) => {
  try {
    await respondSessionInfo(req, res);
  } catch (error) {
    logTunnelError(error, '/api/session_info');
    res.status(500).json({ ok: false, error: 'Session validation failed' });
  }
});

// SPA fallback: non-API routes should return the built frontend entrypoint.
app.use((req: Request, res: Response, next: NextFunction) => {
  if (req.path.startsWith('/api') || req.path === '/health') {
    return next();
  }

  // Only serve SPA shell for browser navigations, not for missing static assets.
  const acceptsHtml = String(req.headers.accept || '').includes('text/html');
  const hasFileExtension = /\.[a-zA-Z0-9]+$/.test(req.path);
  if (!acceptsHtml || hasFileExtension) {
    return next();
  }

  return res.sendFile(path.join(DIST_DIR, 'index.html'));
});

app.use((_req: Request, res: Response) => {
  res.status(404).json({ ok: false, error: 'Route not found' });
});

app.use((error: Error, _req: Request, res: Response, _next: NextFunction) => {
  logTunnelError(error, 'express-middleware');
  res.status(500).json({ ok: false, error: error.message || 'Unhandled server error' });
});

const PORT = Number(process.env.PORT || 3000);
const HOST = '0.0.0.0';

app.listen(PORT, HOST, () => {
  console.log(`[server] Express backend listening on ${HOST}:${PORT}`);
  console.log('[server] Runtime command: npx tsx backend/server.ts');
});
