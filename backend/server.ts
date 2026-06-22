import 'dotenv/config';
import express, { type Request, type Response, type NextFunction } from 'express';
import cors from 'cors';
import path from 'node:path';
import { pingDatabase, queryClient } from './db';
import * as deviceAuthHandlers from './deviceAuth';

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
const [unitsHandlers, turnsHandlers, estimatesHandlers, queueHandlers] = await Promise.all([
  // @ts-ignore
  import('../afproxy/handlers/units.ts'),
  // @ts-ignore
  import('../afproxy/handlers/turns.ts'),
  // @ts-ignore
  import('../afproxy/handlers/estimates.ts'),
  // @ts-ignore
  import('../afproxy/handlers/queue.ts'),
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

function mapTurnStagesFromMilestones(milestones: any): Record<string, any> {
  const out: Record<string, any> = {};
  const keys = ['moveout', 'inspection', 'wo_created', 'est_requested', 'est_received', 'movein'];
  for (const key of keys) {
    const node = milestones?.[key];
    const date = asIso(node?.date);
    if (!date) continue;
    out[key] = {
      done: true,
      date,
      notes: String(node?.notes || ''),
    };
  }
  return out;
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

const GROUP_NAME_BY_UUID: Record<string, string> = {
  'c8f40f01-9e94-11ee-8b51-02167481f3bc': 'Ana Consuelo Properties',
  '9985f145-3cb0-11f0-bfba-069ca18f5865': 'Andrea Robidoux Properties',
  '06fdebec-9e9b-11ee-8b51-02167481f3bc': 'Chris Meehan Properties',
  '3405e65e-9e9c-11ee-8b51-02167481f3bc': 'Jessmar Romea Properties',
  'b368729d-9eca-11ee-8b51-02167481f3bc': 'Jennifer Hazlett Properties',
  '44b79f5e-9e9c-11ee-8b51-02167481f3bc': 'Mary Rees Properties',
  '121e7ca4-9eca-11ee-8b51-02167481f3bc': 'Veronica Garcia Properties',
  '1036e611-9e9c-11ee-8b51-02167481f3bc': 'Nita Lauer Properties',
  '9a434f3b-a04d-11ee-8b51-02167481f3bc': 'Jacquelina Brantley Properties',
  '66a16517-9eca-11ee-8b51-02167481f3bc': 'Angela Hogan Properties',
  '930c330b-60ce-11f0-bfba-069ca18f5865': 'Michelle Kovach Properties',
  '7d4a69d6-9eca-11ee-8b51-02167481f3bc': 'Deborah Lago Properties',
  'bee73529-9eca-11ee-8b51-02167481f3bc': 'Jordan Hammerschmidt Properties',
  'a5774de7-a04d-11ee-8b51-02167481f3bc': 'Michelle Cunningham Properties',
  'f922348b-ea67-11f0-bfba-069ca18f5865': 'Jamie Monty Properties',
  '61a5b6d1-251b-11f0-bfba-069ca18f5865': 'Cari Rascon Properties',
  '7f65b11f-7c52-11f0-bfba-069ca18f5865': 'Sara Anglin',
  'bb129607-e81e-11f0-bfba-069ca18f5865': 'MissionSprings',
  '114bcb4d-e81e-11f0-bfba-069ca18f5865': 'El Diablo',
  '0041c7dd-add1-11f0-bfba-069ca18f5865': 'Maggie Properties',
};

const GROUP_UUID_BY_MANAGER: Record<string, string> = {
  'ana consuelo': 'c8f40f01-9e94-11ee-8b51-02167481f3bc',
  'andrea robidoux': '9985f145-3cb0-11f0-bfba-069ca18f5865',
  'chris meehan': '06fdebec-9e9b-11ee-8b51-02167481f3bc',
  'jessmar romea': '3405e65e-9e9c-11ee-8b51-02167481f3bc',
  'jennifer hazlett': 'b368729d-9eca-11ee-8b51-02167481f3bc',
  'mary rees': '44b79f5e-9e9c-11ee-8b51-02167481f3bc',
  'veronica garcia': '121e7ca4-9eca-11ee-8b51-02167481f3bc',
  'nita lauer': '1036e611-9e9c-11ee-8b51-02167481f3bc',
  'jacquelina brantley': '9a434f3b-a04d-11ee-8b51-02167481f3bc',
  'angela hogan': '66a16517-9eca-11ee-8b51-02167481f3bc',
  'michelle kovach': '930c330b-60ce-11f0-bfba-069ca18f5865',
  'deborah lago': '7d4a69d6-9eca-11ee-8b51-02167481f3bc',
  'jordan hammerschmidt': 'bee73529-9eca-11ee-8b51-02167481f3bc',
  'michelle cunningham': 'a5774de7-a04d-11ee-8b51-02167481f3bc',
  'jamie monty': 'f922348b-ea67-11f0-bfba-069ca18f5865',
  'cari rascon': '61a5b6d1-251b-11f0-bfba-069ca18f5865',
  'sara anglin': '7f65b11f-7c52-11f0-bfba-069ca18f5865',
};

function resolveGroupFromRaw(raw: any, siteManager: string): { id: string; name: string } {
  const direct = String(pickRaw(raw, ['property_group_id', 'PropertyGroupId', 'property_group_uuid', 'PropertyGroupUuid']) || '').trim();
  if (direct) {
    const nm = GROUP_NAME_BY_UUID[direct] || String(pickRaw(raw, ['name_of_property_group', 'NameOfPropertyGroup', 'property_group_name', 'PropertyGroupName']) || direct);
    return { id: direct, name: nm };
  }

  const arr = Array.isArray(raw?.PropertyGroupIds) ? raw.PropertyGroupIds : [];
  for (const candidate of arr) {
    const id = String(candidate || '').trim();
    if (!id) continue;
    if (GROUP_NAME_BY_UUID[id]) return { id, name: GROUP_NAME_BY_UUID[id] };
  }

  const mgr = String(siteManager || '').trim().toLowerCase();
  if (mgr && GROUP_UUID_BY_MANAGER[mgr]) {
    const id = GROUP_UUID_BY_MANAGER[mgr];
    return { id, name: GROUP_NAME_BY_UUID[id] || id };
  }

  return { id: '', name: '' };
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
  const resolvedGroup = resolveGroupFromRaw(raw, siteManager);
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
    property_group: String(pickRaw(raw, ['property_group', 'group_name', 'group', 'Group']) || resolvedGroup.name || ''),
    property_group_id: String(row?.property_group_id || resolvedGroup.id || pickRaw(raw, ['property_group_id', 'PropertyGroupId', 'property_group_uuid', 'PropertyGroupUuid']) || ''),
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
  const workOrderUuid = String(
    row?.work_order_uuid || pickRaw(raw, ['work_order_uuid', 'v0_uuid', 'UUID', 'uuid']) || '',
  );
  const normalized: Record<string, any> = {
    ...raw,
    db_api_id: String(row?.id || pickRaw(raw, ['db_api_id', 'dbApiId', 'UUID', 'Id']) || ''),
    work_order_id: workOrderUuid || String(row?.id || pickRaw(raw, ['work_order_id', 'WorkOrderId', 'Id']) || ''),
    work_order_uuid: workOrderUuid,
    uuid: workOrderUuid,
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
    const days = parseDays(req.query.days, 3650, 3650);
    const limit = parseLimit(req.query.limit, 10000, 20000);
    const propertyGroupId = getPropertyGroupFilter(req);
    const rows = propertyGroupId
      ? await queryClient`
        select id, work_order_uuid, wo_number, property_id, unit_id, property_group_id, description,
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
        select id, work_order_uuid, wo_number, property_id, unit_id, property_group_id, description,
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
    const limit = parseLimit(req.query.limit, 10000, 25000);
    const propertyGroupId = getPropertyGroupFilter(req);
    const rows = propertyGroupId
      ? await queryClient`
        select id, work_order_uuid, wo_number, property_id, unit_id, property_group_id, description,
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
        select id, work_order_uuid, wo_number, property_id, unit_id, property_group_id, description,
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
        d.notes as turn_notes,
        d.reference_user as reference_user,
        d.move_out_date as detail_move_out_date,
        d.turn_end_date as turn_end_date,
        d.expected_move_in_date as expected_move_in_date,
        d.target_days_to_complete as target_days_to_complete,
        d.total_days_to_complete as total_days_to_complete,
        d.labor_from_work_orders as labor_from_work_orders,
        d.purchase_orders_from_work_orders as purchase_orders_from_work_orders,
        d.billables_from_work_orders as billables_from_work_orders,
        d.inventory_from_work_orders as inventory_from_work_orders,
        d.total_billed as total_billed,
        d.unit_turn_status as unit_turn_status,
        ui.last_inspection_date as inspection_date,
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
      left join appfolio_unit_turn_details d
        on d.turn_id = coalesce(t.unit_turn_id, t.turn_key)
      left join appfolio_unit_inspections ui
        on ui.unit_id = t.unit_id and ui.property_id = t.property_id
      where coalesce(t.updated_at, t.created_at, now()) >= now() - (${days}::int * interval '1 day')
      order by coalesce(t.updated_at, t.created_at) desc
      limit ${limit}
    `;

    let filtered = (rows as any[]).filter((row) => {
      if (!statusFilter) return true;
      const status = String(row?.status || '').toLowerCase();
      return status.includes(statusFilter);
    });

    if (filtered.length === 0) {
      const detailRows = await queryClient`
        select
          d.turn_id,
          d.property_id,
          d.property_name,
          d.unit_id,
          d.unit_name,
          d.notes,
          d.reference_user,
          d.move_out_date,
          d.turn_end_date,
          d.expected_move_in_date,
          d.target_days_to_complete,
          d.total_days_to_complete,
          d.labor_from_work_orders,
          d.purchase_orders_from_work_orders,
          d.billables_from_work_orders,
          d.inventory_from_work_orders,
          d.total_billed,
          d.unit_turn_status,
          d.cached_at,
          d.last_updated_at
        from appfolio_unit_turn_details d
        where coalesce(d.last_updated_at, d.cached_at, now()) >= now() - (${days}::int * interval '1 day')
        order by coalesce(d.move_out_date, d.cached_at, d.last_updated_at) desc
        limit ${limit}
      `;

      filtered = (detailRows as any[]).map((row) => ({
        tracking_uuid: String(row.turn_id || ''),
        tracking_code: null,
        turn_key: String(row.turn_id || ''),
        unit_turn_id: String(row.turn_id || ''),
        unit_id: String(row.unit_id || ''),
        property_id: String(row.property_id || ''),
        unit_name: String(row.unit_name || ''),
        property_name: String(row.property_name || ''),
        status: String(row.unit_turn_status || 'In Progress'),
        confidence_score: null,
        confidence_label: null,
        source_flags: {},
        metadata: { source: 'unit_turn_detail' },
        closed_at: row.turn_end_date,
        created_at: row.cached_at,
        updated_at: row.last_updated_at || row.cached_at,
        detail_move_out_date: row.move_out_date,
        expected_move_in_date: row.expected_move_in_date,
        turn_end_date: row.turn_end_date,
        target_days_to_complete: row.target_days_to_complete,
        total_days_to_complete: row.total_days_to_complete,
        labor_from_work_orders: row.labor_from_work_orders,
        purchase_orders_from_work_orders: row.purchase_orders_from_work_orders,
        billables_from_work_orders: row.billables_from_work_orders,
        inventory_from_work_orders: row.inventory_from_work_orders,
        total_billed: row.total_billed,
        unit_turn_status: row.unit_turn_status,
        reference_user: row.reference_user,
        milestones: {},
        linked_work_orders: [],
      })).filter((row) => {
        if (!statusFilter) return true;
        return String(row.status || '').toLowerCase().includes(statusFilter);
      });
    }

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
      move_out_date: asIso(row.detail_move_out_date) || asIso(row.metadata?.move_out_date),
      move_in_date: asIso(row.expected_move_in_date) || asIso(row.metadata?.move_in_date),
      inspection_date: asIso(row.inspection_date) || asIso(row.metadata?.inspection_date),
      expected_move_in_date: asIso(row.expected_move_in_date),
      turn_end_date: asIso(row.turn_end_date),
      target_days_to_complete: Number(row.target_days_to_complete || 0) || 0,
      total_days_to_complete: Number(row.total_days_to_complete || 0) || 0,
      labor_from_work_orders: String(row.labor_from_work_orders || ''),
      purchase_orders_from_work_orders: String(row.purchase_orders_from_work_orders || ''),
      billables_from_work_orders: String(row.billables_from_work_orders || ''),
      inventory_from_work_orders: String(row.inventory_from_work_orders || ''),
      total_billed: String(row.total_billed || ''),
      unit_turn_status: String(row.unit_turn_status || ''),
      reference_user: String(row.reference_user || ''),
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

app.get('/api/local/turn_records', async (req: Request, res: Response) => {
  try {
    const days = parseDays(req.query.days, 540, 3650);
    const limit = parseLimit(req.query.limit, 500, 5000);
    const propertyGroupId = getPropertyGroupFilter(req);

    const rows = propertyGroupId
      ? await queryClient`
        select
          t.turn_key,
          t.unit_turn_id,
          t.tracking_uuid,
          t.updated_at,
          coalesce((
            select jsonb_object_agg(m.milestone_key, jsonb_build_object(
              'date', m.milestone_date,
              'source', m.source,
              'notes', m.notes
            ))
            from unit_turn_milestones m
            where m.tracking_uuid = t.tracking_uuid
          ), '{}'::jsonb) as milestones
        from unit_turn_tracker t
        inner join appfolio_properties p on p.id = t.property_id
        where coalesce(t.updated_at, t.created_at, now()) >= now() - (${days}::int * interval '1 day')
          and p.property_group_id = ${propertyGroupId}
        order by coalesce(t.updated_at, t.created_at) desc
        limit ${limit}
      `
      : await queryClient`
        select
          t.turn_key,
          t.unit_turn_id,
          t.tracking_uuid,
          t.updated_at,
          coalesce((
            select jsonb_object_agg(m.milestone_key, jsonb_build_object(
              'date', m.milestone_date,
              'source', m.source,
              'notes', m.notes
            ))
            from unit_turn_milestones m
            where m.tracking_uuid = t.tracking_uuid
          ), '{}'::jsonb) as milestones
        from unit_turn_tracker t
        where coalesce(t.updated_at, t.created_at, now()) >= now() - (${days}::int * interval '1 day')
        order by coalesce(t.updated_at, t.created_at) desc
        limit ${limit}
      `;

    const records = (rows as any[]).map((row) => ({
      id: String(row.turn_key || row.unit_turn_id || row.tracking_uuid || '').trim(),
      stages: mapTurnStagesFromMilestones(row.milestones || {}),
      updated_at: asIso(row.updated_at),
      _source: 'postgres_local',
    })).filter((r) => r.id);

    res.json({ ok: true, records, count: records.length, source: 'postgres_local' });
  } catch (error) {
    logTunnelError(error, '/api/local/turn_records');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local turn records query failed') });
  }
});

app.get('/api/local/closed_turns', async (req: Request, res: Response) => {
  try {
    const days = parseDays(req.query.days, 540, 3650);
    const limit = parseLimit(req.query.limit, 500, 5000);
    const propertyGroupId = getPropertyGroupFilter(req);

    const rows = propertyGroupId
      ? await queryClient`
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
          t.closed_at,
          t.metadata,
          t.updated_at,
          (select m.milestone_date from unit_turn_milestones m
            where m.tracking_uuid = t.tracking_uuid and m.milestone_key = 'moveout'
            order by coalesce(m.milestone_date, m.created_at) desc nulls last
            limit 1) as move_out_date,
          (select m.milestone_date from unit_turn_milestones m
            where m.tracking_uuid = t.tracking_uuid and m.milestone_key = 'movein'
            order by coalesce(m.milestone_date, m.created_at) desc nulls last
            limit 1) as move_in_date
        from unit_turn_tracker t
        inner join appfolio_properties p on p.id = t.property_id
        where p.property_group_id = ${propertyGroupId}
          and (t.closed_at is not null or lower(coalesce(t.status, '')) in ('closed', 'completed'))
          and coalesce(t.closed_at, t.updated_at, t.created_at, now()) >= now() - (${days}::int * interval '1 day')
        order by coalesce(t.closed_at, t.updated_at, t.created_at) desc
        limit ${limit}
      `
      : await queryClient`
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
          t.closed_at,
          t.metadata,
          t.updated_at,
          (select m.milestone_date from unit_turn_milestones m
            where m.tracking_uuid = t.tracking_uuid and m.milestone_key = 'moveout'
            order by coalesce(m.milestone_date, m.created_at) desc nulls last
            limit 1) as move_out_date,
          (select m.milestone_date from unit_turn_milestones m
            where m.tracking_uuid = t.tracking_uuid and m.milestone_key = 'movein'
            order by coalesce(m.milestone_date, m.created_at) desc nulls last
            limit 1) as move_in_date
        from unit_turn_tracker t
        where (t.closed_at is not null or lower(coalesce(t.status, '')) in ('closed', 'completed'))
          and coalesce(t.closed_at, t.updated_at, t.created_at, now()) >= now() - (${days}::int * interval '1 day')
        order by coalesce(t.closed_at, t.updated_at, t.created_at) desc
        limit ${limit}
      `;

    const results = (rows as any[]).map((row) => {
      const meta = (row?.metadata && typeof row.metadata === 'object') ? row.metadata : {};
      const turnId = String(row.turn_key || row.unit_turn_id || row.tracking_uuid || '').trim();
      return {
        turn_id: turnId,
        turn_key: String(row.turn_key || '').trim(),
        tracking_uuid: String(row.tracking_uuid || '').trim(),
        unit_turn_id: String(row.unit_turn_id || '').trim(),
        tracking_code: String(row.tracking_code || '').trim(),
        property_id: String(row.property_id || '').trim(),
        property_name: String(row.property_name || '').trim(),
        unit_id: String(row.unit_id || '').trim(),
        unit_name: String(row.unit_name || '').trim(),
        move_out_date: asIso(row.move_out_date),
        move_in_date: asIso(row.move_in_date),
        site_manager: String(meta.site_manager || ''),
        close_source: String(meta.close_source || ''),
        close_reason: String(meta.close_reason || ''),
        status: String(row.status || ''),
        closed_at: asIso(row.closed_at) || asIso(row.updated_at),
        _source: 'postgres_local',
      };
    }).filter((r) => r.turn_id);

    res.json({ ok: true, results, count: results.length, source: 'postgres_local' });
  } catch (error) {
    logTunnelError(error, '/api/local/closed_turns');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local closed turns query failed') });
  }
});

app.get('/api/local/turns_history', async (req: Request, res: Response) => {
  try {
    const days = parseDays(req.query.days, 540, 3650);
    const limit = parseLimit(req.query.limit, 300, 5000);
    const propertyGroupId = getPropertyGroupFilter(req);

    const rows = propertyGroupId
      ? await queryClient`
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
          t.closed_at,
          t.metadata,
          t.updated_at,
          (select m.milestone_date from unit_turn_milestones m
            where m.tracking_uuid = t.tracking_uuid and m.milestone_key = 'moveout'
            order by coalesce(m.milestone_date, m.created_at) desc nulls last
            limit 1) as move_out_date,
          (select m.milestone_date from unit_turn_milestones m
            where m.tracking_uuid = t.tracking_uuid and m.milestone_key = 'movein'
            order by coalesce(m.milestone_date, m.created_at) desc nulls last
            limit 1) as move_in_date
        from unit_turn_tracker t
        inner join appfolio_properties p on p.id = t.property_id
        where p.property_group_id = ${propertyGroupId}
          and (t.closed_at is not null or lower(coalesce(t.status, '')) in ('closed', 'completed'))
          and coalesce(t.closed_at, t.updated_at, t.created_at, now()) >= now() - (${days}::int * interval '1 day')
        order by coalesce(t.closed_at, t.updated_at, t.created_at) desc
        limit ${limit}
      `
      : await queryClient`
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
          t.closed_at,
          t.metadata,
          t.updated_at,
          (select m.milestone_date from unit_turn_milestones m
            where m.tracking_uuid = t.tracking_uuid and m.milestone_key = 'moveout'
            order by coalesce(m.milestone_date, m.created_at) desc nulls last
            limit 1) as move_out_date,
          (select m.milestone_date from unit_turn_milestones m
            where m.tracking_uuid = t.tracking_uuid and m.milestone_key = 'movein'
            order by coalesce(m.milestone_date, m.created_at) desc nulls last
            limit 1) as move_in_date
        from unit_turn_tracker t
        where (t.closed_at is not null or lower(coalesce(t.status, '')) in ('closed', 'completed'))
          and coalesce(t.closed_at, t.updated_at, t.created_at, now()) >= now() - (${days}::int * interval '1 day')
        order by coalesce(t.closed_at, t.updated_at, t.created_at) desc
        limit ${limit}
      `;

    const results = (rows as any[]).map((row) => {
      const meta = (row?.metadata && typeof row.metadata === 'object') ? row.metadata : {};
      return {
        closed_at: asIso(row.closed_at) || asIso(row.updated_at),
        close_source: String(meta.close_source || ''),
        close_reason: String(meta.close_reason || ''),
        property_name: String(row.property_name || ''),
        unit_name: String(row.unit_name || ''),
        move_out_date: asIso(row.move_out_date),
        move_in_date: asIso(row.move_in_date),
        site_manager: String(meta.site_manager || ''),
        tracking_code: String(row.tracking_code || ''),
        turn_key: String(row.turn_key || ''),
        tracking_uuid: String(row.tracking_uuid || ''),
        unit_turn_id: String(row.unit_turn_id || ''),
        _source: 'postgres_local',
      };
    });

    res.json({ ok: true, results, count: results.length, source: 'postgres_local' });
  } catch (error) {
    logTunnelError(error, '/api/local/turns_history');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local turns history query failed') });
  }
});

app.get('/api/local/estimates', async (req: Request, res: Response) => {
  try {
    const limit = parseLimit(req.query.limit, 2500, 10000);
    const propertyGroupId = getPropertyGroupFilter(req);
    const rows = propertyGroupId
      ? await queryClient`
        select
          e.estimate_id,
          e.work_order_id,
          e.work_order_number,
          e.current_status,
          e.property_group_id,
          e.raw_data,
          wo.vendor_name,
          p.name as property_name,
          u.name as unit_name,
          e.updated_at
        from appfolio_estimates e
        left join appfolio_work_orders wo on wo.id = e.work_order_id
        left join appfolio_properties p on p.id = wo.property_id
        left join appfolio_units u on u.unit_id = wo.unit_id
        where coalesce(e.property_group_id, wo.property_group_id) = ${propertyGroupId}
        order by coalesce(e.updated_at, now()) desc
        limit ${limit}
      `
      : await queryClient`
        select
          e.estimate_id,
          e.work_order_id,
          e.work_order_number,
          e.current_status,
          e.property_group_id,
          e.raw_data,
          wo.vendor_name,
          p.name as property_name,
          u.name as unit_name,
          e.updated_at
        from appfolio_estimates e
        left join appfolio_work_orders wo on wo.id = e.work_order_id
        left join appfolio_properties p on p.id = wo.property_id
        left join appfolio_units u on u.unit_id = wo.unit_id
        order by coalesce(e.updated_at, now()) desc
        limit ${limit}
      `;

    const results = (rows as any[]).map((row) => {
      const raw = (row?.raw_data && typeof row.raw_data === 'object') ? row.raw_data : {};
      const propertyName = String(row?.property_name || pickRaw(raw, ['property_name', 'PropertyName']) || '').trim();
      const unitName = String(row?.unit_name || pickRaw(raw, ['unit_name', 'UnitName']) || '').trim();
      return {
        estimate_id: String(row.estimate_id || ''),
        work_order_id: String(row.work_order_id || ''),
        work_order_number: String(row.work_order_number || pickRaw(raw, ['work_order_number', 'WorkOrderNumber']) || ''),
        property_unit_address: [propertyName, unitName].filter(Boolean).join(' · '),
        vendor_name: String(row.vendor_name || pickRaw(raw, ['vendor_name', 'VendorName']) || ''),
        estimate_amount: pickRaw(raw, ['estimate_amount', 'EstimateAmount', 'amount', 'Amount']),
        approval_status: String(row.current_status || pickRaw(raw, ['approval_status', 'ApprovalStatus', 'current_status']) || 'Pending'),
        property_group_id: String(row.property_group_id || pickRaw(raw, ['property_group_id', 'property_group_uuid', 'PropertyGroupId']) || ''),
        updated_at: asIso(row.updated_at),
        _source: 'postgres_local',
      };
    });

    res.json({ ok: true, results, count: results.length, source: 'postgres_local' });
  } catch (error) {
    logTunnelError(error, '/api/local/estimates');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local estimates query failed') });
  }
});

app.get('/api/local/work_orders_completed_history', async (req: Request, res: Response) => {
  try {
    const days = parseDays(req.query.days, 365, 3650);
    const limit = parseLimit(req.query.limit, 10000, 25000);
    const propertyGroupId = getPropertyGroupFilter(req);
    const rows = propertyGroupId
      ? await queryClient`
        select id, work_order_uuid, wo_number, property_id, unit_id, property_group_id, description,
               category, priority, status, assigned_user_id, assigned_user_name,
               vendor_id, vendor_name, estimated_amount, total_cost,
               created_at, updated_at, raw_json
        from appfolio_work_orders
        where coalesce(updated_at, created_at, now()) >= now() - (${days}::int * interval '1 day')
          and property_group_id = ${propertyGroupId}
          and (
            coalesce(lower(status), '') like '%completed%'
            or coalesce(lower(status), '') like '%no need to bill%'
            or coalesce(lower(status), '') like '%cancel%'
          )
        order by coalesce(updated_at, created_at) desc
        limit ${limit}
      `
      : await queryClient`
        select id, work_order_uuid, wo_number, property_id, unit_id, property_group_id, description,
               category, priority, status, assigned_user_id, assigned_user_name,
               vendor_id, vendor_name, estimated_amount, total_cost,
               created_at, updated_at, raw_json
        from appfolio_work_orders
        where coalesce(updated_at, created_at, now()) >= now() - (${days}::int * interval '1 day')
          and (
            coalesce(lower(status), '') like '%completed%'
            or coalesce(lower(status), '') like '%no need to bill%'
            or coalesce(lower(status), '') like '%cancel%'
          )
        order by coalesce(updated_at, created_at) desc
        limit ${limit}
      `;

    const results = (rows as any[]).map(normalizeWorkOrderRow);
    res.json({ ok: true, results, count: results.length, source: 'postgres_local' });
  } catch (error) {
    logTunnelError(error, '/api/local/work_orders_completed_history');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local completed work orders query failed') });
  }
});

app.get('/api/local/inspections', async (req: Request, res: Response) => {
  try {
    const limit = parseLimit(req.query.limit, 6000, 15000);
    const propertyGroupId = getPropertyGroupFilter(req);
    let rows = propertyGroupId
      ? await queryClient`
        select
          i.inspection_id,
          i.property_id,
          coalesce(i.property_name, p.name) as property_name,
          i.unit_id,
          coalesce(i.unit_name, u.name, '') as unit_name,
          i.last_inspection_date,
          i.tenant_name,
          i.tenant_primary_phone_number,
          i.move_in_date,
          i.move_out_date,
          i.rentable,
          i.unit_tags
        from appfolio_unit_inspections i
        left join appfolio_properties p on p.id = i.property_id
        left join appfolio_units u on u.unit_id = i.unit_id
        where p.property_group_id = ${propertyGroupId}
        order by coalesce(i.last_inspection_date, i.cached_at) desc, coalesce(i.property_name, p.name) asc, coalesce(i.unit_name, u.name) asc
        limit ${limit}
      `
      : await queryClient`
        select
          i.inspection_id,
          i.property_id,
          coalesce(i.property_name, p.name) as property_name,
          i.unit_id,
          coalesce(i.unit_name, u.name, '') as unit_name,
          i.last_inspection_date,
          i.tenant_name,
          i.tenant_primary_phone_number,
          i.move_in_date,
          i.move_out_date,
          i.rentable,
          i.unit_tags
        from appfolio_unit_inspections i
        left join appfolio_properties p on p.id = i.property_id
        left join appfolio_units u on u.unit_id = i.unit_id
        order by coalesce(i.last_inspection_date, i.cached_at) desc, coalesce(i.property_name, p.name) asc, coalesce(i.unit_name, u.name) asc
        limit ${limit}
      `;

    if ((rows as any[]).length === 0) {
      rows = propertyGroupId
        ? await queryClient`
          select
            u.unit_id,
            u.property_id,
            p.name as property_name,
            coalesce(nullif(u.raw_json->>'unit_name',''), nullif(u.raw_json->>'UnitName',''), u.name, '') as unit_name,
            coalesce(nullif(u.raw_json->>'last_inspection_date',''), nullif(u.raw_json->>'LastInspectionDate','')) as last_inspection_date,
            coalesce(nullif(u.raw_json->>'tenant_name',''), nullif(u.raw_json->>'TenantName','')) as tenant_name,
            coalesce(nullif(u.raw_json->>'tenant_primary_phone_number',''), nullif(u.raw_json->>'TenantPrimaryPhoneNumber','')) as tenant_primary_phone_number,
            coalesce(nullif(u.raw_json->>'move_in_date',''), nullif(u.raw_json->>'MoveInDate','')) as move_in_date,
            coalesce(nullif(u.raw_json->>'move_out_date',''), nullif(u.raw_json->>'MoveOutDate','')) as move_out_date,
            coalesce(nullif(u.raw_json->>'rentable',''), nullif(u.raw_json->>'Rentable','')) as rentable,
            coalesce(u.raw_json->>'unit_tags', u.raw_json->>'UnitTags', '') as unit_tags
          from appfolio_units u
          inner join appfolio_properties p on p.id = u.property_id
          where p.property_group_id = ${propertyGroupId}
          order by p.name asc, u.name asc
          limit ${limit}
        `
        : await queryClient`
          select
            u.unit_id,
            u.property_id,
            p.name as property_name,
            coalesce(nullif(u.raw_json->>'unit_name',''), nullif(u.raw_json->>'UnitName',''), u.name, '') as unit_name,
            coalesce(nullif(u.raw_json->>'last_inspection_date',''), nullif(u.raw_json->>'LastInspectionDate','')) as last_inspection_date,
            coalesce(nullif(u.raw_json->>'tenant_name',''), nullif(u.raw_json->>'TenantName','')) as tenant_name,
            coalesce(nullif(u.raw_json->>'tenant_primary_phone_number',''), nullif(u.raw_json->>'TenantPrimaryPhoneNumber','')) as tenant_primary_phone_number,
            coalesce(nullif(u.raw_json->>'move_in_date',''), nullif(u.raw_json->>'MoveInDate','')) as move_in_date,
            coalesce(nullif(u.raw_json->>'move_out_date',''), nullif(u.raw_json->>'MoveOutDate','')) as move_out_date,
            coalesce(nullif(u.raw_json->>'rentable',''), nullif(u.raw_json->>'Rentable','')) as rentable,
            coalesce(u.raw_json->>'unit_tags', u.raw_json->>'UnitTags', '') as unit_tags
          from appfolio_units u
          inner join appfolio_properties p on p.id = u.property_id
          order by p.name asc, u.name asc
          limit ${limit}
        `;

    }

    const results = (rows as any[]).map((row) => ({
      property_name: String(row.property_name || ''),
      property_id: String(row.property_id || ''),
      unit_name: String(row.unit_name || ''),
      unit_id: String(row.unit_id || ''),
      last_inspection_date: String(row.last_inspection_date || ''),
      tenant_name: String(row.tenant_name || ''),
      tenant_primary_phone_number: String(row.tenant_primary_phone_number || ''),
      move_in_date: String(row.move_in_date || ''),
      move_out_date: String(row.move_out_date || ''),
      rentable: String(row.rentable || ''),
      unit_tags: row.unit_tags || '',
      _source: 'postgres_local',
    }));

    res.json({ ok: true, results, count: results.length, source: 'postgres_local' });
  } catch (error) {
    logTunnelError(error, '/api/local/inspections');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local inspections query failed') });
  }
});

app.get('/api/local/upcoming_moveouts', async (req: Request, res: Response) => {
  try {
    const days = parseDays(req.query.days, 60, 3650);
    const limit = parseLimit(req.query.limit, 2500, 10000);
    const propertyGroupId = getPropertyGroupFilter(req);
    let rows = propertyGroupId
      ? await queryClient`
        select
          t.record_id,
          t.property_id,
          coalesce(t.property_name, p.name) as property_name,
          t.unit_id,
          coalesce(t.unit_name, u.name, '') as unit_name,
          t.tenant_name,
          t.move_out_date,
          t.move_in_date,
          t.phone_numbers,
          t.emails,
          t.rent,
          t.occupancy_id
        from appfolio_tenant_directory t
        left join appfolio_properties p on p.id = t.property_id
        left join appfolio_units u on u.unit_id = t.unit_id
        where p.property_group_id = ${propertyGroupId}
        order by coalesce(t.move_out_date, t.cached_at) asc, coalesce(t.property_name, p.name) asc, coalesce(t.unit_name, u.name) asc
        limit ${limit}
      `
      : await queryClient`
        select
          t.record_id,
          t.property_id,
          coalesce(t.property_name, p.name) as property_name,
          t.unit_id,
          coalesce(t.unit_name, u.name, '') as unit_name,
          t.tenant_name,
          t.move_out_date,
          t.move_in_date,
          t.phone_numbers,
          t.emails,
          t.rent,
          t.occupancy_id
        from appfolio_tenant_directory t
        left join appfolio_properties p on p.id = t.property_id
        left join appfolio_units u on u.unit_id = t.unit_id
        order by coalesce(t.move_out_date, t.cached_at) asc, coalesce(t.property_name, p.name) asc, coalesce(t.unit_name, u.name) asc
        limit ${limit}
      `;

    if ((rows as any[]).length === 0) {
      rows = propertyGroupId
        ? await queryClient`
          select
            u.unit_id,
            u.property_id,
            p.name as property_name,
            coalesce(nullif(u.raw_json->>'unit_name',''), nullif(u.raw_json->>'UnitName',''), u.name, '') as unit_name,
            coalesce(nullif(u.raw_json->>'tenant_name',''), nullif(u.raw_json->>'TenantName','')) as tenant_name,
            coalesce(nullif(u.raw_json->>'move_out_date',''), nullif(u.raw_json->>'MoveOutDate','')) as move_out_date,
            coalesce(nullif(u.raw_json->>'move_in_date',''), nullif(u.raw_json->>'MoveInDate','')) as move_in_date,
            coalesce(nullif(u.raw_json->>'tenant_primary_phone_number',''), nullif(u.raw_json->>'TenantPrimaryPhoneNumber','')) as phone_numbers,
            coalesce(nullif(u.raw_json->>'tenant_email',''), nullif(u.raw_json->>'TenantEmail','')) as emails,
            coalesce(nullif(u.raw_json->>'rent',''), nullif(u.raw_json->>'Rent',''), nullif(u.raw_json->>'market_rent','')) as rent,
            coalesce(nullif(u.raw_json->>'occupancy_id',''), nullif(u.raw_json->>'OccupancyId','')) as occupancy_id
          from appfolio_units u
          inner join appfolio_properties p on p.id = u.property_id
          where p.property_group_id = ${propertyGroupId}
          order by p.name asc, u.name asc
          limit ${limit}
        `
        : await queryClient`
          select
            u.unit_id,
            u.property_id,
            p.name as property_name,
            coalesce(nullif(u.raw_json->>'unit_name',''), nullif(u.raw_json->>'UnitName',''), u.name, '') as unit_name,
            coalesce(nullif(u.raw_json->>'tenant_name',''), nullif(u.raw_json->>'TenantName','')) as tenant_name,
            coalesce(nullif(u.raw_json->>'move_out_date',''), nullif(u.raw_json->>'MoveOutDate','')) as move_out_date,
            coalesce(nullif(u.raw_json->>'move_in_date',''), nullif(u.raw_json->>'MoveInDate','')) as move_in_date,
            coalesce(nullif(u.raw_json->>'tenant_primary_phone_number',''), nullif(u.raw_json->>'TenantPrimaryPhoneNumber','')) as phone_numbers,
            coalesce(nullif(u.raw_json->>'tenant_email',''), nullif(u.raw_json->>'TenantEmail','')) as emails,
            coalesce(nullif(u.raw_json->>'rent',''), nullif(u.raw_json->>'Rent',''), nullif(u.raw_json->>'market_rent','')) as rent,
            coalesce(nullif(u.raw_json->>'occupancy_id',''), nullif(u.raw_json->>'OccupancyId','')) as occupancy_id
          from appfolio_units u
          inner join appfolio_properties p on p.id = u.property_id
          order by p.name asc, u.name asc
          limit ${limit}
        `;

    }

    const now = Date.now();
    const maxMs = now + (days * 86400000);
    const results = (rows as any[]).map((row) => ({
      property_name: String(row.property_name || ''),
      property_id: String(row.property_id || ''),
      unit_name: String(row.unit_name || ''),
      unit_id: String(row.unit_id || ''),
      occupancy_name: String(row.tenant_name || ''),
      move_out_date: String(row.move_out_date || ''),
      move_in_date: String(row.move_in_date || ''),
      phone_numbers: String(row.phone_numbers || ''),
      emails: String(row.emails || ''),
      rent: String(row.rent || ''),
      occupancy_id: String(row.occupancy_id || ''),
      _source: 'postgres_local',
    })).filter((row) => {
      if (!row.move_out_date) return false;
      const d = new Date(row.move_out_date);
      if (Number.isNaN(d.getTime())) return true;
      const ms = d.getTime();
      return ms >= now && ms <= maxMs;
    });

    res.json({ ok: true, results, count: results.length, source: 'postgres_local' });
  } catch (error) {
    logTunnelError(error, '/api/local/upcoming_moveouts');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local upcoming move-outs query failed') });
  }
});

app.get('/api/local/vacancies', async (req: Request, res: Response) => {
  try {
    const limit = parseLimit(req.query.limit, 5000, 15000);
    const propertyGroupId = getPropertyGroupFilter(req);
    let rows = propertyGroupId
      ? await queryClient`
        select
          v.record_id,
          v.property_id,
          coalesce(v.property_name, p.name) as property_name,
          v.unit_id,
          coalesce(v.unit_name, u.name, '') as unit_name,
          v.vacant_from,
          v.market_rent,
          v.bedrooms,
          v.days_vacant,
          v.status,
          p.property_group_id,
          p.raw_json
        from appfolio_unit_vacancies v
        left join appfolio_properties p on p.id = v.property_id
        left join appfolio_units u on u.unit_id = v.unit_id
        where p.property_group_id = ${propertyGroupId}
        order by coalesce(v.vacant_from, v.cached_at) asc, coalesce(v.property_name, p.name) asc, coalesce(v.unit_name, u.name) asc
        limit ${limit}
      `
      : await queryClient`
        select
          v.record_id,
          v.property_id,
          coalesce(v.property_name, p.name) as property_name,
          v.unit_id,
          coalesce(v.unit_name, u.name, '') as unit_name,
          v.vacant_from,
          v.market_rent,
          v.bedrooms,
          v.days_vacant,
          v.status,
          p.property_group_id,
          p.raw_json
        from appfolio_unit_vacancies v
        left join appfolio_properties p on p.id = v.property_id
        left join appfolio_units u on u.unit_id = v.unit_id
        order by coalesce(v.vacant_from, v.cached_at) asc, coalesce(v.property_name, p.name) asc, coalesce(v.unit_name, u.name) asc
        limit ${limit}
      `;

    if ((rows as any[]).length === 0) {
      rows = propertyGroupId
        ? await queryClient`
          select
            u.unit_id,
            u.property_id,
            p.name as property_name,
            coalesce(nullif(u.raw_json->>'unit_name',''), nullif(u.raw_json->>'UnitName',''), u.name, '') as unit_name,
            coalesce(nullif(u.raw_json->>'move_out_date',''), nullif(u.raw_json->>'MoveOutDate','')) as vacant_from,
            coalesce(nullif(u.raw_json->>'market_rent',''), nullif(u.raw_json->>'MarketRent',''), nullif(u.raw_json->>'rent','')) as market_rent,
            coalesce(nullif(u.raw_json->>'bedrooms',''), nullif(u.raw_json->>'Bedrooms','')) as bedrooms,
            p.property_group_id,
            p.raw_json
          from appfolio_units u
          inner join appfolio_properties p on p.id = u.property_id
          where p.property_group_id = ${propertyGroupId}
            and lower(coalesce(u.status, '')) like '%vacant%'
          order by p.name asc, u.name asc
          limit ${limit}
        `
        : await queryClient`
          select
            u.unit_id,
            u.property_id,
            p.name as property_name,
            coalesce(nullif(u.raw_json->>'unit_name',''), nullif(u.raw_json->>'UnitName',''), u.name, '') as unit_name,
            coalesce(nullif(u.raw_json->>'move_out_date',''), nullif(u.raw_json->>'MoveOutDate','')) as vacant_from,
            coalesce(nullif(u.raw_json->>'market_rent',''), nullif(u.raw_json->>'MarketRent',''), nullif(u.raw_json->>'rent','')) as market_rent,
            coalesce(nullif(u.raw_json->>'bedrooms',''), nullif(u.raw_json->>'Bedrooms','')) as bedrooms,
            p.property_group_id,
            p.raw_json
          from appfolio_units u
          inner join appfolio_properties p on p.id = u.property_id
          where lower(coalesce(u.status, '')) like '%vacant%'
          order by p.name asc, u.name asc
          limit ${limit}
        `;
    }

    const results = (rows as any[]).map((row) => {
      const raw = (row?.raw_json && typeof row.raw_json === 'object') ? row.raw_json : {};
      const groupName = String(pickRaw(raw, ['property_group', 'group_name', 'property_group_name', 'NameOfPropertyGroup']) || '');
      const vacantFrom = String(row.vacant_from || '');
      let daysVacant: number | null = null;
      if (vacantFrom) {
        const d = new Date(vacantFrom);
        if (!Number.isNaN(d.getTime())) {
          daysVacant = Math.max(0, Math.floor((Date.now() - d.getTime()) / 86400000));
        }
      }
      return {
        property_id: String(row.property_id || ''),
        property_name: String(row.property_name || ''),
        unit: String(row.unit_name || ''),
        vacant_from: vacantFrom,
        days_vacant: daysVacant,
        market_rent: row.market_rent,
        bedrooms: row.bedrooms,
        property_group: groupName,
        _source: 'postgres_local',
      };
    });

    res.json({ ok: true, results, count: results.length, source: 'postgres_local' });
  } catch (error) {
    logTunnelError(error, '/api/local/vacancies');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local vacancies query failed') });
  }
});

app.get('/api/local/property_map', async (req: Request, res: Response) => {
  try {
    const limit = parseLimit(req.query.limit, 7000, 20000);
    const propertyGroupId = getPropertyGroupFilter(req);
    const rows = propertyGroupId
      ? await queryClient`
        select id, name, property_group_id, raw_json
        from appfolio_properties
        where property_group_id = ${propertyGroupId}
        order by name asc
        limit ${limit}
      `
      : await queryClient`
        select id, name, property_group_id, raw_json
        from appfolio_properties
        order by name asc
        limit ${limit}
      `;

    const property_uuid_map: Record<string, any> = {};
    for (const row of rows as any[]) {
      const raw = (row?.raw_json && typeof row.raw_json === 'object') ? row.raw_json : {};
      const siteManager = extractSiteManager(raw);
      const groupId = String(row?.property_group_id || pickRaw(raw, ['property_group_id', 'PropertyGroupId', 'property_group_uuid', 'PropertyGroupUuid']) || '').trim();
      property_uuid_map[String(row.id || '')] = {
        name: String(row.name || ''),
        site_manager_name: String(siteManager || ''),
        group_ids: groupId ? [groupId] : [],
        maintenance_notes: String(pickRaw(raw, ['maintenance_notes', 'maintenanceNotes', 'MaintenanceNotes']) || ''),
      };
    }

    res.json({ ok: true, property_uuid_map, count: Object.keys(property_uuid_map).length, source: 'postgres_local' });
  } catch (error) {
    logTunnelError(error, '/api/local/property_map');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local property map query failed') });
  }
});

app.get('/api/local/property_stats', async (req: Request, res: Response) => {
  try {
    const propertyGroupId = getPropertyGroupFilter(req);
    const rows = propertyGroupId
      ? await queryClient`
        select p.id,
          (
            select count(*)::int
            from appfolio_work_orders w
            where w.property_id = p.id
              and (
                coalesce(lower(w.status), '') not like '%completed%'
                and coalesce(lower(w.status), '') not like '%cancel%'
                and coalesce(lower(w.status), '') not like '%no need to bill%'
              )
          ) as open_work_orders,
          (
            select count(*)::int
            from appfolio_units u
            where u.property_id = p.id and lower(coalesce(u.status, '')) like '%vacant%'
          ) as vacancies,
          (
            select count(*)::int
            from appfolio_estimates e
            where e.property_group_id = p.property_group_id
          ) as estimates
        from appfolio_properties p
        where p.property_group_id = ${propertyGroupId}
      `
      : await queryClient`
        select p.id,
          (
            select count(*)::int
            from appfolio_work_orders w
            where w.property_id = p.id
              and (
                coalesce(lower(w.status), '') not like '%completed%'
                and coalesce(lower(w.status), '') not like '%cancel%'
                and coalesce(lower(w.status), '') not like '%no need to bill%'
              )
          ) as open_work_orders,
          (
            select count(*)::int
            from appfolio_units u
            where u.property_id = p.id and lower(coalesce(u.status, '')) like '%vacant%'
          ) as vacancies,
          (
            select count(*)::int
            from appfolio_estimates e
            where e.property_group_id = p.property_group_id
          ) as estimates
        from appfolio_properties p
      `;

    const by_property: Record<string, any> = {};
    for (const row of rows as any[]) {
      const id = String(row.id || '').trim();
      if (!id) continue;
      by_property[id] = {
        bills: Number(row.open_work_orders || 0),
        notes: Number(row.estimates || 0),
        listings: Number(row.vacancies || 0),
      };
    }

    res.json({ ok: true, by_property, count: Object.keys(by_property).length, source: 'postgres_local' });
  } catch (error) {
    logTunnelError(error, '/api/local/property_stats');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local property stats query failed') });
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

function startRecurringSyncScheduler(): void {
  const enabled = /^(1|true|yes|on)$/i.test(String(process.env.SYNC_SCHEDULER_ENABLED || '').trim());
  if (!enabled) return;

  const intervalMinutes = Math.max(5, Number(process.env.SYNC_SCHEDULER_INTERVAL_MINUTES || '30') || 30);
  const endpoints = String(process.env.SYNC_SCHEDULER_ENDPOINTS || 'v0:properties,v0:property_groups,v0:units,v0:work_orders,v2:tenant_directory,v2:unit_inspection,v2:unit_turn_detail,v2:unit_vacancy')
    .split(',')
    .map((value) => String(value || '').trim())
    .filter(Boolean);
  const maxPages = Math.max(0, Number(process.env.SYNC_SCHEDULER_MAX_PAGES || '0') || 0);
  const runOnBoot = !/^(0|false|no|off)$/i.test(String(process.env.SYNC_SCHEDULER_RUN_ON_BOOT || 'true').trim());
  const inFlight = new Set<string>();

  async function runEndpoint(endpointKey: string, triggerType: string): Promise<void> {
    if (!endpointKey || inFlight.has(endpointKey)) return;
    inFlight.add(endpointKey);
    try {
      const { runSync } = await import('./sync/syncRunner.ts');
      const summary = await runSync({ endpointKey, triggerType, maxPages });
      console.log('[server:sync-scheduler] completed', summary);
    } catch (error) {
      console.error('[server:sync-scheduler] failed', endpointKey, String((error as any)?.message || error));
    } finally {
      inFlight.delete(endpointKey);
    }
  }

  async function tick(triggerType: string): Promise<void> {
    for (const endpointKey of endpoints) {
      await runEndpoint(endpointKey, triggerType);
    }
  }

  if (runOnBoot) {
    void tick('startup');
  }

  setInterval(() => {
    void tick('scheduled');
  }, intervalMinutes * 60 * 1000);

  console.log('[server:sync-scheduler] enabled', { endpoints, intervalMinutes, maxPages, runOnBoot });
}

const PORT = Number(process.env.PORT || 3000);
const HOST = '0.0.0.0';

app.listen(PORT, HOST, () => {
  console.log(`[server] Express backend listening on ${HOST}:${PORT}`);
  console.log('[server] Runtime command: npx tsx backend/server.ts');
  startRecurringSyncScheduler();
});
