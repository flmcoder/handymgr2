import 'dotenv/config';
import express, { type Request, type Response, type NextFunction } from 'express';
import cors from 'cors';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { pingDatabase, queryClient } from './db';
import * as deviceAuthHandlers from './deviceAuth';
import { AF_DB_BASE, AF_REPORTS_BASE, afHeaders } from './sync/afCredentials';

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

const SYNC_ENDPOINT_ALLOWLIST = new Set<string>([
  'v0:properties',
  'v0:property_groups',
  'v0:units',
  'v0:work_orders',
  'v2:tenant_directory',
  'v2:unit_inspection',
  'v2:unit_turn_detail',
  'v2:unit_vacancy',
]);

const SYNC_ENDPOINT_BLOCKLIST = new Set<string>([
  'v2:work_orders',
]);

function sanitizeSyncEndpoints(
  requested: string[],
  fallback: string[],
  source: string,
): { accepted: string[]; rejected: string[] } {
  const dedupedRequested = Array.from(new Set(
    requested
      .map((value) => String(value || '').trim())
      .filter(Boolean),
  ));

  const rejected = dedupedRequested.filter((endpoint) => (
    !SYNC_ENDPOINT_ALLOWLIST.has(endpoint) || SYNC_ENDPOINT_BLOCKLIST.has(endpoint)
  ));

  let accepted = dedupedRequested.filter((endpoint) => (
    SYNC_ENDPOINT_ALLOWLIST.has(endpoint) && !SYNC_ENDPOINT_BLOCKLIST.has(endpoint)
  ));

  if (accepted.length === 0) {
    accepted = Array.from(new Set(
      fallback
        .map((value) => String(value || '').trim())
        .filter((endpoint) => SYNC_ENDPOINT_ALLOWLIST.has(endpoint) && !SYNC_ENDPOINT_BLOCKLIST.has(endpoint)),
    ));
  }

  if (rejected.length > 0) {
    console.warn(`[server:sync-guard] source=${source} rejected_endpoints=${rejected.join(',')}`);
  }

  return { accepted, rejected };
}

function parsePropertyGroupId(value: unknown): string {
  const raw = String(value ?? '').trim();
  if (!raw) return '';
  // Accept single values or comma-delimited lists; local endpoints currently scope to one group.
  const first = raw.split(',').map((v) => v.trim()).find(Boolean);
  return String(first || '');
}

function getRequestedPropertyGroupId(req: Request): string {
  return parsePropertyGroupId(
    req.query.property_group_id
      ?? req.query.property_group_uuid
      ?? req.query.group_id
      ?? req.query.propertyGroupId,
  );
}

function getBearerToken(req: Request): string {
  const auth = String(req.headers.authorization || '');
  return auth.startsWith('Bearer ') ? auth.slice(7).trim() : '';
}

type LocalScopeContext = {
  role: string;
  requestedGroupId: string;
  sessionGroupId: string;
  effectiveGroupId: string;
  scopeSource: 'query' | 'session';
  enforced: boolean;
};

async function pmScopeMiddleware(req: Request, res: Response, next: NextFunction): Promise<void> {
  try {
    const requestedGroupId = getRequestedPropertyGroupId(req);
    const token = getBearerToken(req);

    // Backward-compatible: if no bearer token is present, keep existing query behavior.
    if (!token) {
      (req as any).localScope = {
        role: 'anonymous',
        requestedGroupId,
        sessionGroupId: '',
        effectiveGroupId: requestedGroupId,
        scopeSource: 'query',
        enforced: false,
      } as LocalScopeContext;
      next();
      return;
    }

    const session = await deviceAuthHandlers.getTrustedDeviceSession(token);
    if (!session) {
      res.status(401).json({ ok: false, error: 'Invalid session' });
      return;
    }

    const role = String(session.role || '').toLowerCase();
    const sessionGroupId = normalizeScopeUuid(session.property_group_uuid || '');
    let effectiveGroupId = requestedGroupId;
    let scopeSource: 'query' | 'session' = 'query';
    let enforced = false;

    if (role === 'pm_readonly') {
      if (!sessionGroupId) {
        res.status(403).json({ ok: false, error: 'PM session missing scoped property group' });
        return;
      }
      effectiveGroupId = sessionGroupId;
      scopeSource = 'session';
      enforced = true;
    }

    (req as any).localScope = {
      role,
      requestedGroupId,
      sessionGroupId,
      effectiveGroupId,
      scopeSource,
      enforced,
    } as LocalScopeContext;

    next();
  } catch (error) {
    logTunnelError(error, 'pm-scope-middleware');
    res.status(500).json({ ok: false, error: 'Scope middleware failed' });
  }
}

function getPropertyGroupFilter(req: Request): string {
  const localScope = (req as any).localScope as LocalScopeContext | undefined;
  if (localScope && typeof localScope.effectiveGroupId === 'string') {
    return String(localScope.effectiveGroupId || '');
  }
  return getRequestedPropertyGroupId(req);
}

function asIso(value: unknown): string {
  if (!value) return '';
  const d = value instanceof Date ? value : new Date(String(value));
  return Number.isNaN(d.getTime()) ? '' : d.toISOString();
}

function looksLikeUuidLabel(value: unknown): boolean {
  const raw = String(value || '').trim();
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(raw);
}

function resolvePropertyGroupDisplayName(row: any, fallback = ''): string {
  const raw = (row?.raw_json && typeof row.raw_json === 'object') ? row.raw_json : {};
  const direct = String(row?.name || row?.group_name || '').trim();
  const candidates = [
    direct,
    pickRaw(raw, ['NameOfPropertyGroup', 'name_of_property_group', 'PropertyGroupName', 'property_group_name']),
    pickRaw(raw, ['property_group', 'PropertyGroup', 'group_name', 'GroupName']),
    pickRaw(raw, ['portfolio_name', 'PortfolioName', 'portfolio', 'Portfolio']),
    fallback,
    String(row?.uuid || row?.id || '').trim(),
  ].map((value) => String(value || '').trim()).filter(Boolean);

  const nonUuid = candidates.find((value) => !looksLikeUuidLabel(value));
  return nonUuid || candidates[0] || '';
}

const REPORT_FAST_MS = 6_000;
const REPORT_SLOW_MS = 15_000;
const REPORT_PREFER_LOCAL_MS = 30 * 60_000;
const REPORT_POLICY_TTL_MS = 5 * 60_000;

let reportLatencyTableEnsured = false;
const reportMonitorInFlight = new Set<string>();

async function ensureReportLatencyPolicyTable(): Promise<void> {
  if (reportLatencyTableEnsured) return;
  await queryClient.unsafe(`
    CREATE TABLE IF NOT EXISTS appfolio_report_latency_policy (
      dataset_key TEXT NOT NULL,
      property_group_id TEXT NOT NULL DEFAULT '',
      last_status TEXT NOT NULL DEFAULT 'unknown',
      last_latency_ms INTEGER,
      prefer_local_until TIMESTAMPTZ,
      last_checked_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      last_fast_at TIMESTAMPTZ,
      last_slow_at TIMESTAMPTZ,
      meta_json JSONB NOT NULL DEFAULT '{}'::jsonb,
      PRIMARY KEY (dataset_key, property_group_id)
    )
  `);
  reportLatencyTableEnsured = true;
}

async function readReportPolicy(datasetKey: string, propertyGroupId: string): Promise<any | null> {
  await ensureReportLatencyPolicyTable();
  const rows = await queryClient`
    select dataset_key, property_group_id, last_status, last_latency_ms,
           prefer_local_until, last_checked_at, last_fast_at, last_slow_at, meta_json
    from appfolio_report_latency_policy
    where dataset_key = ${datasetKey}
      and property_group_id = ${propertyGroupId}
    limit 1
  `;
  return (rows as any[])[0] || null;
}

async function upsertReportPolicy(
  datasetKey: string,
  propertyGroupId: string,
  status: 'fast' | 'warm' | 'slow' | 'error',
  latencyMs: number,
  meta: Record<string, unknown> = {},
  preferLocalUntil: Date | null = null,
): Promise<void> {
  await ensureReportLatencyPolicyTable();
  const nowIso = new Date().toISOString();
  const fastAt = status === 'fast' ? nowIso : null;
  const slowAt = status === 'slow' ? nowIso : null;
  await queryClient.unsafe(
    `INSERT INTO appfolio_report_latency_policy (
       dataset_key, property_group_id, last_status, last_latency_ms,
       prefer_local_until, last_checked_at, last_fast_at, last_slow_at, meta_json
     ) VALUES ($1, $2, $3, $4, $5, NOW(), $6, $7, $8::jsonb)
     ON CONFLICT (dataset_key, property_group_id) DO UPDATE SET
       last_status = EXCLUDED.last_status,
       last_latency_ms = EXCLUDED.last_latency_ms,
       prefer_local_until = EXCLUDED.prefer_local_until,
       last_checked_at = NOW(),
       last_fast_at = COALESCE(EXCLUDED.last_fast_at, appfolio_report_latency_policy.last_fast_at),
       last_slow_at = COALESCE(EXCLUDED.last_slow_at, appfolio_report_latency_policy.last_slow_at),
       meta_json = EXCLUDED.meta_json`,
    [
      datasetKey,
      propertyGroupId,
      status,
      Math.max(0, Math.round(Number(latencyMs || 0))),
      preferLocalUntil ? preferLocalUntil.toISOString() : null,
      fastAt,
      slowAt,
      JSON.stringify(meta || {}),
    ],
  );
}

async function fetchWithHardTimeout(url: string, init: RequestInit, timeoutMs: number): Promise<{ response?: Response; timedOut: boolean; elapsedMs: number; error?: unknown }> {
  const started = Date.now();
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(url, {
      ...init,
      signal: controller.signal,
    });
    return { response, timedOut: false, elapsedMs: Date.now() - started };
  } catch (error) {
    const err = error as any;
    return {
      timedOut: err?.name === 'AbortError',
      elapsedMs: Date.now() - started,
      error,
    };
  } finally {
    clearTimeout(timer);
  }
}

function mapLiveInspectionRows(rows: any[]): any[] {
  return (rows || []).map((row) => ({
    property_name: String(pickRaw(row, ['property_name', 'PropertyName', 'property', 'Property']) || ''),
    property_id: String(pickRaw(row, ['property_id', 'PropertyId']) || ''),
    unit_name: String(pickRaw(row, ['unit_name', 'UnitName', 'unit', 'Unit']) || ''),
    unit_id: String(pickRaw(row, ['unit_id', 'UnitId']) || ''),
    last_inspection_date: String(pickRaw(row, ['last_inspection_date', 'LastInspectionDate']) || ''),
    tenant_name: String(pickRaw(row, ['tenant_name', 'TenantName']) || ''),
    tenant_primary_phone_number: String(pickRaw(row, ['tenant_primary_phone_number', 'TenantPrimaryPhoneNumber']) || ''),
    move_in_date: String(pickRaw(row, ['move_in_date', 'MoveInDate']) || ''),
    move_out_date: String(pickRaw(row, ['move_out_date', 'MoveOutDate']) || ''),
    rentable: String(pickRaw(row, ['rentable', 'Rentable']) || ''),
    unit_tags: pickRaw(row, ['unit_tags', 'UnitTags']) || '',
    _source: 'appfolio_live',
  }));
}

async function fetchLiveInspectionRows(propertyGroupId: string, timeoutMs: number): Promise<{ ok: boolean; timedOut: boolean; latencyMs: number; status: number; rows: any[]; error: string }> {
  const body: Record<string, unknown> = {
    unit_visibility: 'active',
    include_blank_inspection_date: '1',
    columns: [
      'property', 'property_name', 'property_id', 'unit_name', 'unit_id',
      'last_inspection_date', 'tenant_name', 'tenant_primary_phone_number',
      'move_in_date', 'move_out_date', 'rentable', 'unit_tags',
    ],
  };
  if (propertyGroupId) body.property_groups_ids = [propertyGroupId];

  const url = `${AF_REPORTS_BASE}/api/v2/reports/unit_inspection.json`;
  const attempt = await fetchWithHardTimeout(url, {
    method: 'POST',
    headers: afHeaders('v2'),
    body: JSON.stringify(body),
  }, timeoutMs);

  if (!attempt.response) {
    return {
      ok: false,
      timedOut: !!attempt.timedOut,
      latencyMs: attempt.elapsedMs,
      status: 0,
      rows: [],
      error: String((attempt.error as any)?.message || attempt.error || 'live inspection fetch failed'),
    };
  }

  const response = attempt.response;
  const text = await response.text();
  let parsed: any = {};
  try {
    parsed = text ? JSON.parse(text) : {};
  } catch {
    parsed = {};
  }
  const rows = Array.isArray(parsed?.results)
    ? parsed.results
    : (Array.isArray(parsed?.data) ? parsed.data : []);

  return {
    ok: response.ok,
    timedOut: false,
    latencyMs: attempt.elapsedMs,
    status: response.status,
    rows,
    error: response.ok ? '' : String(parsed?.error || text || `HTTP ${response.status}`),
  };
}

async function monitorInspectionSlowPath(propertyGroupId: string): Promise<void> {
  const groupKey = String(propertyGroupId || '');
  const lockKey = `v2:unit_inspection:${groupKey || '__all__'}`;
  if (reportMonitorInFlight.has(lockKey)) return;
  reportMonitorInFlight.add(lockKey);

  try {
    const probe = await fetchLiveInspectionRows(groupKey, REPORT_SLOW_MS);
    if (probe.ok && probe.latencyMs <= REPORT_SLOW_MS) {
      await upsertReportPolicy('v2:unit_inspection', groupKey, probe.latencyMs <= REPORT_FAST_MS ? 'fast' : 'warm', probe.latencyMs, {
        mode: 'background_probe',
        status: probe.status,
        row_count: probe.rows.length,
      }, null);
      return;
    }

    const preferUntil = new Date(Date.now() + REPORT_PREFER_LOCAL_MS);
    await upsertReportPolicy('v2:unit_inspection', groupKey, probe.timedOut ? 'slow' : 'error', probe.latencyMs, {
      mode: 'background_probe',
      status: probe.status,
      timed_out: probe.timedOut,
      error: probe.error,
    }, preferUntil);

    const { runSync } = await import('./sync/syncRunner.ts');
    await runSync({
      endpointKey: 'v2:unit_inspection',
      triggerType: 'latency_slow_auto',
      maxPages: 1,
    });
  } catch (error) {
    console.error('[local:inspections] slow-path monitor failed', String((error as any)?.message || error));
  } finally {
    reportMonitorInFlight.delete(lockKey);
  }
}

async function monitorVendorSlowPath(propertyGroupId: string): Promise<void> {
  const groupKey = String(propertyGroupId || '');
  const lockKey = `v0:work_orders:vendors:${groupKey || '__all__'}`;
  if (reportMonitorInFlight.has(lockKey)) return;
  reportMonitorInFlight.add(lockKey);

  try {
    let url = `${AF_DB_BASE}/api/v0/work_orders?page%5Bsize%5D=1`;
    if (groupKey) {
      url += `&filters%5BPropertyGroupId%5D=${encodeURIComponent(groupKey)}`;
    }
    const probe = await fetchWithHardTimeout(url, {
      method: 'GET',
      headers: afHeaders('v0'),
    }, REPORT_SLOW_MS);

    const status = probe.response ? probe.response.status : 0;
    if (probe.response && probe.response.ok && probe.elapsedMs <= REPORT_SLOW_MS) {
      await upsertReportPolicy('v0:work_orders:vendors', groupKey, probe.elapsedMs <= REPORT_FAST_MS ? 'fast' : 'warm', probe.elapsedMs, {
        mode: 'background_probe',
        status,
      }, null);
      return;
    }

    const preferUntil = new Date(Date.now() + REPORT_PREFER_LOCAL_MS);
    await upsertReportPolicy('v0:work_orders:vendors', groupKey, probe.timedOut ? 'slow' : 'error', probe.elapsedMs, {
      mode: 'background_probe',
      status,
      timed_out: probe.timedOut,
      error: String((probe.error as any)?.message || probe.error || ''),
    }, preferUntil);

    const { runSync } = await import('./sync/syncRunner.ts');
    await runSync({
      endpointKey: 'v0:work_orders',
      triggerType: 'latency_slow_auto',
      maxPages: 1,
    });
  } catch (error) {
    console.error('[local:vendors] slow-path monitor failed', String((error as any)?.message || error));
  } finally {
    reportMonitorInFlight.delete(lockKey);
  }
}

type JwkKey = {
  kid?: string;
  kty?: string;
  alg?: string;
  n?: string;
  e?: string;
  use?: string;
  key_ops?: string[];
};

let webhookTableEnsured = false;
let webhookJwksCache: { keys: JwkKey[]; fetchedAt: number } = { keys: [], fetchedAt: 0 };

const WEBHOOK_JWKS_URL = 'https://api.appfolio.com/.well-known/jwks.json';
const WEBHOOK_JWKS_TTL_MS = Math.max(60_000, Number(process.env.WEBHOOK_JWKS_TTL_MS || String(6 * 60 * 60 * 1000)) || (6 * 60 * 60 * 1000));
const WEBHOOK_VERIFY_SIGNATURE = !/^(0|false|no|off)$/i.test(String(process.env.WEBHOOK_VERIFY_SIGNATURE || 'true').trim());

function base64UrlToUint8Array(input: string): Uint8Array {
  const normalized = String(input || '').replace(/-/g, '+').replace(/_/g, '/');
  const padLen = normalized.length % 4;
  const padded = padLen === 0 ? normalized : normalized + '='.repeat(4 - padLen);
  return new Uint8Array(Buffer.from(padded, 'base64'));
}

function bufferToBase64Url(input: Buffer): string {
  return input.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

function singularizeTopic(topicRaw: string): string {
  const topic = String(topicRaw || '').trim().toLowerCase();
  if (!topic) return '';
  if (topic.endsWith('ies')) return topic.slice(0, -3) + 'y';
  if (topic.endsWith('s')) return topic.slice(0, -1);
  return topic;
}

function normalizeWebhookResourceType(rawResourceType: string, topic: string): string {
  const cleaned = String(rawResourceType || '').trim().toLowerCase().replace(/[^a-z0-9_.-]/g, '');
  if (cleaned.includes('.')) {
    return singularizeTopic(cleaned.split('.')[0] || '');
  }
  if (cleaned) return singularizeTopic(cleaned);
  return singularizeTopic(topic);
}

function parseWebhookEventType(rawEventType: string): string {
  const value = String(rawEventType || '').trim().toLowerCase();
  if (!value) return '';
  if (value.includes('.')) return String(value.split('.')[1] || '').trim();
  return value;
}

async function ensureWebhookEventsTable(): Promise<void> {
  if (webhookTableEnsured) return;
  await queryClient.unsafe(`
    CREATE TABLE IF NOT EXISTS webhook_events (
      id BIGSERIAL PRIMARY KEY,
      event_uuid TEXT,
      topic TEXT NOT NULL,
      event_type TEXT,
      resource_type TEXT,
      resource_id TEXT,
      signature TEXT,
      payload_json JSONB NOT NULL DEFAULT '{}'::jsonb,
      received_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      processed_at TIMESTAMPTZ,
      processing_status TEXT DEFAULT 'pending'
    )
  `);
  await queryClient.unsafe(`
    CREATE INDEX IF NOT EXISTS webhook_events_topic_idx
      ON webhook_events(topic)
  `);
  await queryClient.unsafe(`
    CREATE INDEX IF NOT EXISTS webhook_events_resource_idx
      ON webhook_events(resource_type, resource_id)
  `);
  await queryClient.unsafe(`
    CREATE INDEX IF NOT EXISTS webhook_events_status_idx
      ON webhook_events(processing_status)
  `);
  await queryClient.unsafe(`
    CREATE UNIQUE INDEX IF NOT EXISTS webhook_events_event_uuid_unique
      ON webhook_events(event_uuid)
      WHERE event_uuid IS NOT NULL AND event_uuid <> ''
  `);
  webhookTableEnsured = true;
}

async function fetchWebhookJwks(force = false): Promise<JwkKey[]> {
  const now = Date.now();
  if (!force && webhookJwksCache.keys.length > 0 && (now - webhookJwksCache.fetchedAt) < WEBHOOK_JWKS_TTL_MS) {
    return webhookJwksCache.keys;
  }

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 20_000);
  try {
    const response = await fetch(WEBHOOK_JWKS_URL, {
      method: 'GET',
      headers: { Accept: 'application/json' },
      signal: controller.signal,
    });
    if (!response.ok) {
      throw new Error(`JWKS fetch failed: HTTP ${response.status}`);
    }
    const payload: any = await response.json();
    const keys = Array.isArray(payload?.keys) ? payload.keys as JwkKey[] : [];
    if (!keys.length) throw new Error('JWKS returned no keys');
    webhookJwksCache = { keys, fetchedAt: now };
    return keys;
  } finally {
    clearTimeout(timeout);
  }
}

async function verifyAppfolioDetachedJws(rawBody: string, detachedSignatureHeader: string): Promise<{ ok: boolean; kid: string; alg: string }> {
  const detached = String(detachedSignatureHeader || '').trim();
  if (!detached) return { ok: false, kid: '', alg: '' };

  const parts = detached.split('..');
  if (parts.length !== 2) return { ok: false, kid: '', alg: '' };
  const protectedB64 = String(parts[0] || '').trim();
  const signatureB64 = String(parts[1] || '').trim();
  if (!protectedB64 || !signatureB64) return { ok: false, kid: '', alg: '' };

  let protectedHeader: any = {};
  try {
    protectedHeader = JSON.parse(Buffer.from(base64UrlToUint8Array(protectedB64)).toString('utf8'));
  } catch {
    return { ok: false, kid: '', alg: '' };
  }

  const kid = String(protectedHeader?.kid || '').trim();
  const alg = String(protectedHeader?.alg || '').trim();
  if (!kid || alg !== 'PS256') {
    return { ok: false, kid, alg };
  }

  const payloadB64 = bufferToBase64Url(Buffer.from(String(rawBody || ''), 'utf8'));
  const signingInput = new Uint8Array(Buffer.from(`${protectedB64}.${payloadB64}`, 'utf8'));
  const signature = base64UrlToUint8Array(signatureB64);

  let keys = await fetchWebhookJwks(false);
  let key = keys.find((entry) => String(entry?.kid || '').trim() === kid);
  if (!key) {
    keys = await fetchWebhookJwks(true);
    key = keys.find((entry) => String(entry?.kid || '').trim() === kid);
  }
  if (!key) return { ok: false, kid, alg };

  const cryptoKey = await crypto.subtle.importKey(
    'jwk',
    key as any,
    { name: 'RSA-PSS', hash: 'SHA-256' },
    false,
    ['verify'],
  );

  const verified = await crypto.subtle.verify(
    { name: 'RSA-PSS', saltLength: 32 },
    cryptoKey,
    signature as any,
    signingInput as any,
  );

  return { ok: !!verified, kid, alg };
}

function mapWebhookToSyncEndpoints(resourceTypeRaw: string, topicRaw: string): string[] {
  const resourceType = normalizeWebhookResourceType(resourceTypeRaw, topicRaw);
  switch (resourceType) {
    case 'work_order':
      return ['v0:work_orders'];
    case 'property':
      return ['v0:properties'];
    case 'property_group':
      return ['v0:property_groups'];
    case 'unit':
      return ['v0:units'];
    case 'tenant':
      return ['v2:tenant_directory'];
    case 'inspection':
    case 'unit_inspection':
      return ['v2:unit_inspection'];
    case 'unit_turn':
      return ['v2:unit_turn_detail'];
    case 'vacancy':
    case 'unit_vacancy':
      return ['v2:unit_vacancy'];
    default:
      return [];
  }
}

async function runWebhookDeltaSync(resourceTypeRaw: string, topicRaw: string): Promise<string[]> {
  const endpoints = mapWebhookToSyncEndpoints(resourceTypeRaw, topicRaw)
    .filter((endpoint) => SYNC_ENDPOINT_ALLOWLIST.has(endpoint) && !SYNC_ENDPOINT_BLOCKLIST.has(endpoint));
  if (!endpoints.length) return [];

  const { runSync } = await import('./sync/syncRunner.ts');
  for (const endpointKey of endpoints) {
    try {
      await runSync({ endpointKey, triggerType: 'webhook', maxPages: 1 });
    } catch (error) {
      console.error('[webhook:delta-sync] failed', endpointKey, String((error as any)?.message || error));
    }
  }
  return endpoints;
}

const WORK_ORDER_DETAIL_CACHE_TTL_MS = Math.max(
  60_000,
  Number(process.env.WO_DETAIL_CACHE_TTL_MS || String(15 * 60 * 1000)) || (15 * 60 * 1000),
);

let workOrderDetailCacheTableEnsured = false;

async function ensureWorkOrderDetailCacheTable(): Promise<void> {
  if (workOrderDetailCacheTableEnsured) return;
  await queryClient.unsafe(`
    CREATE TABLE IF NOT EXISTS appfolio_work_order_detail_cache (
      work_order_uuid TEXT NOT NULL,
      detail_type TEXT NOT NULL,
      payload JSONB NOT NULL DEFAULT '[]'::jsonb,
      payload_hash TEXT NOT NULL DEFAULT '',
      record_count INTEGER NOT NULL DEFAULT 0,
      last_fetched_at TIMESTAMPTZ,
      last_checked_at TIMESTAMPTZ,
      fetch_count INTEGER NOT NULL DEFAULT 0,
      PRIMARY KEY (work_order_uuid, detail_type)
    )
  `);
  await queryClient.unsafe(`
    CREATE INDEX IF NOT EXISTS appfolio_work_order_detail_cache_fetched_idx
      ON appfolio_work_order_detail_cache(last_fetched_at)
  `);
  workOrderDetailCacheTableEnsured = true;
}

function isUuidLike(value: unknown): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(String(value || '').trim());
}

function normalizeDetailList(payload: any): any[] {
  if (!payload) return [];
  if (Array.isArray(payload)) return payload;
  if (Array.isArray(payload.results)) return payload.results;
  if (Array.isArray(payload.Results)) return payload.Results;
  if (Array.isArray(payload.items)) return payload.items;
  if (Array.isArray(payload.data)) return payload.data;
  if (Array.isArray(payload.attachments)) return payload.attachments;
  return [];
}

function detailPayloadHash(payload: any[]): string {
  return createHash('sha256').update(JSON.stringify(payload || [])).digest('hex').slice(0, 16);
}

function detailItemKey(item: any): string {
  const candidates = [
    item?.id,
    item?.Id,
    item?.uuid,
    item?.UUID,
    item?.work_order_note_id,
    item?.WorkOrderNoteId,
    item?.work_order_attachment_id,
    item?.WorkOrderAttachmentId,
    item?.created_at,
    item?.CreatedAt,
    item?.updated_at,
    item?.UpdatedAt,
    item?.file_name,
    item?.FileName,
    item?.name,
    item?.Name,
    item?.body,
    item?.Body,
  ];
  for (const value of candidates) {
    const normalized = String(value || '').trim();
    if (normalized) return normalized;
  }
  return createHash('sha256').update(JSON.stringify(item || {})).digest('hex').slice(0, 16);
}

function computeDetailDelta(previousPayload: any[], nextPayload: any[]): { changed: boolean; added: number; removed: number; previous_count: number; current_count: number } {
  const prevKeys = new Set((previousPayload || []).map((item) => detailItemKey(item)));
  const nextKeys = new Set((nextPayload || []).map((item) => detailItemKey(item)));
  let added = 0;
  let removed = 0;
  for (const key of nextKeys.values()) {
    if (!prevKeys.has(key)) added++;
  }
  for (const key of prevKeys.values()) {
    if (!nextKeys.has(key)) removed++;
  }
  return {
    changed: added > 0 || removed > 0,
    added,
    removed,
    previous_count: previousPayload.length,
    current_count: nextPayload.length,
  };
}

async function resolveWorkOrderUuid(refRaw: string): Promise<string> {
  const ref = String(refRaw || '').trim();
  if (!ref) return '';
  if (isUuidLike(ref)) return ref;

  const rows = await queryClient`
    select work_order_uuid, id, wo_number, raw_json
    from appfolio_work_orders
    where id = ${ref} or wo_number = ${ref}
    order by updated_at desc nulls last
    limit 1
  `;
  const row = (rows as any[])[0] || {};
  const resolved = String(
    row?.work_order_uuid || row?.id || row?.raw_json?.work_order_uuid || row?.raw_json?.v0_uuid || row?.raw_json?.UUID || row?.raw_json?.uuid || '',
  ).trim();
  return isUuidLike(resolved) ? resolved : '';
}

async function fetchWorkOrderDetailFromAppFolio(workOrderUuid: string, detailType: 'notes' | 'attachments'): Promise<any[]> {
  const url = `${AF_DB_BASE}/api/v0/work_orders/${encodeURIComponent(workOrderUuid)}/${detailType}`;
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 30_000);
  try {
    const response = await fetch(url, {
      method: 'GET',
      headers: afHeaders('v0'),
      signal: controller.signal,
    });
    const rawText = await response.text();
    if (!response.ok) {
      const preview = String(rawText || '').replace(/\s+/g, ' ').slice(0, 220);
      throw new Error(`AppFolio ${detailType} fetch failed: HTTP ${response.status} ${preview}`);
    }

    let parsed: any = {};
    try {
      parsed = rawText ? JSON.parse(rawText) : {};
    } catch {
      parsed = {};
    }
    return normalizeDetailList(parsed);
  } finally {
    clearTimeout(timeout);
  }
}

async function getWorkOrderDetailCached(args: { workOrderRef: string; detailType: 'notes' | 'attachments'; forceRefresh: boolean }): Promise<{ workOrderUuid: string; results: any[]; source: 'live' | 'cached' | 'stale_cache'; stale: boolean; fetched_at: string; delta: { changed: boolean; added: number; removed: number; previous_count: number; current_count: number } }> {
  await ensureWorkOrderDetailCacheTable();

  const resolvedUuid = await resolveWorkOrderUuid(args.workOrderRef);
  if (!resolvedUuid) {
    throw new Error('Unable to resolve AppFolio work order UUID for detail fetch');
  }

  const cachedRows = await queryClient`
    select payload, payload_hash, record_count, last_fetched_at, last_checked_at, fetch_count
    from appfolio_work_order_detail_cache
    where work_order_uuid = ${resolvedUuid}
      and detail_type = ${args.detailType}
    limit 1
  `;

  const cached = (cachedRows as any[])[0] || null;
  const now = Date.now();
  const cachedPayload = normalizeDetailList(cached?.payload || []);
  const lastFetchedAt = cached?.last_fetched_at ? new Date(String(cached.last_fetched_at)).getTime() : 0;
  const isFresh = !!lastFetchedAt && (now - lastFetchedAt) <= WORK_ORDER_DETAIL_CACHE_TTL_MS;

  if (cached && isFresh && !args.forceRefresh) {
    return {
      workOrderUuid: resolvedUuid,
      results: cachedPayload,
      source: 'cached',
      stale: false,
      fetched_at: asIso(cached.last_fetched_at) || new Date(lastFetchedAt).toISOString(),
      delta: {
        changed: false,
        added: 0,
        removed: 0,
        previous_count: cachedPayload.length,
        current_count: cachedPayload.length,
      },
    };
  }

  try {
    const freshPayload = await fetchWorkOrderDetailFromAppFolio(resolvedUuid, args.detailType);
    const freshHash = detailPayloadHash(freshPayload);
    const previousHash = String(cached?.payload_hash || '');
    const changed = freshHash !== previousHash;
    const delta = computeDetailDelta(cachedPayload, freshPayload);

    if (!cached) {
      await queryClient`
        insert into appfolio_work_order_detail_cache (
          work_order_uuid, detail_type, payload, payload_hash, record_count,
          last_fetched_at, last_checked_at, fetch_count
        )
        values (
          ${resolvedUuid}, ${args.detailType}, ${JSON.stringify(freshPayload)}::jsonb, ${freshHash}, ${freshPayload.length},
          now(), now(), 1
        )
      `;
    } else if (changed) {
      await queryClient`
        update appfolio_work_order_detail_cache
        set payload = ${JSON.stringify(freshPayload)}::jsonb,
            payload_hash = ${freshHash},
            record_count = ${freshPayload.length},
            last_fetched_at = now(),
            last_checked_at = now(),
            fetch_count = coalesce(fetch_count, 0) + 1
        where work_order_uuid = ${resolvedUuid}
          and detail_type = ${args.detailType}
      `;
    } else {
      await queryClient`
        update appfolio_work_order_detail_cache
        set last_checked_at = now(),
            last_fetched_at = now(),
            fetch_count = coalesce(fetch_count, 0) + 1
        where work_order_uuid = ${resolvedUuid}
          and detail_type = ${args.detailType}
      `;
    }

    return {
      workOrderUuid: resolvedUuid,
      results: freshPayload,
      source: 'live',
      stale: false,
      fetched_at: new Date().toISOString(),
      delta: {
        ...delta,
        changed,
      },
    };
  } catch (error) {
    if (cached) {
      return {
        workOrderUuid: resolvedUuid,
        results: cachedPayload,
        source: 'stale_cache',
        stale: true,
        fetched_at: asIso(cached.last_fetched_at),
        delta: {
          changed: false,
          added: 0,
          removed: 0,
          previous_count: cachedPayload.length,
          current_count: cachedPayload.length,
        },
      };
    }
    throw error;
  }
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

// AppFolio webhook ingress uses detached JWS verification on the exact raw payload.
// Keep this route above express.json middleware to avoid body mutation.
app.post('/api/webhooks/appfolio', express.text({ type: '*/*', limit: '2mb' }), async (req: Request, res: Response) => {
  try {
    await ensureWebhookEventsTable();
    const rawBody = typeof req.body === 'string' ? req.body : String(req.body || '');
    const signatureHeader = String(req.header('x-jws-signature') || req.header('X-JWS-Signature') || '').trim();

    const verification = WEBHOOK_VERIFY_SIGNATURE
      ? await verifyAppfolioDetachedJws(rawBody, signatureHeader)
      : { ok: true, kid: '', alg: '' };

    if (WEBHOOK_VERIFY_SIGNATURE && !verification.ok) {
      await queryClient`
        insert into webhook_events (
          event_uuid, topic, event_type, resource_type, resource_id,
          signature, payload_json, processed_at, processing_status
        )
        values (
          null, 'unverified', '', '', '',
          ${signatureHeader}, ${JSON.stringify({ raw: rawBody, verification })}::jsonb, now(), 'invalid_signature'
        )
      `;
      res.status(401).json({ ok: false, error: 'Invalid webhook signature' });
      return;
    }

    let payload: any = {};
    try {
      payload = rawBody ? JSON.parse(rawBody) : {};
    } catch {
      res.status(400).json({ ok: false, error: 'Invalid JSON payload' });
      return;
    }

    const topic = String(payload?.topic || '').trim().toLowerCase();
    const eventType = parseWebhookEventType(String(payload?.event_type || payload?.type || ''));
    const resourceType = normalizeWebhookResourceType(String(payload?.resource_type || ''), topic);
    const resourceId = String(payload?.resource_id || '').trim();
    const eventUuid = String(payload?.event_id || '').trim();

    if (eventUuid) {
      const existing = await queryClient`
        select id from webhook_events where event_uuid = ${eventUuid} limit 1
      `;
      if ((existing as any[]).length > 0) {
        res.status(200).json({ ok: true, duplicate: true, event_uuid: eventUuid });
        return;
      }
    }

    const insertRows = await queryClient`
      insert into webhook_events (
        event_uuid, topic, event_type, resource_type, resource_id,
        signature, payload_json, processing_status
      )
      values (
        ${eventUuid || null}, ${topic || 'unknown'}, ${eventType || ''}, ${resourceType || ''}, ${resourceId || ''},
        ${signatureHeader}, ${JSON.stringify(payload)}::jsonb, 'received'
      )
      returning id
    `;
    const webhookRowId = Number((insertRows as any[])[0]?.id || 0) || 0;

    const syncedEndpoints = await runWebhookDeltaSync(resourceType, topic);
    await queryClient`
      update webhook_events
      set processed_at = now(),
          processing_status = ${syncedEndpoints.length ? 'processed' : 'stored'}
      where id = ${webhookRowId}
    `;

    res.status(200).json({
      ok: true,
      verified: verification.ok,
      event_uuid: eventUuid || null,
      topic,
      event_type: eventType,
      resource_type: resourceType,
      resource_id: resourceId || null,
      synced_endpoints: syncedEndpoints,
    });
  } catch (error) {
    logTunnelError(error, '/api/webhooks/appfolio');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Webhook ingest failed') });
  }
});

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

async function requireProxySession(req: Request, res: Response): Promise<any | null> {
  const token = getBearerToken(req);
  if (!token) {
    res.status(401).json({ ok: false, error: 'Missing bearer token' });
    return null;
  }

  const session = await deviceAuthHandlers.getTrustedDeviceSession(token);
  if (!session) {
    res.status(401).json({ ok: false, error: 'Invalid session' });
    return null;
  }

  return session;
}

async function proxyAppFolioAction(req: Request, res: Response, apiPath: string): Promise<void> {
  const safePath = String(apiPath || '').trim();
  if (!safePath.startsWith('/api/v0/') && !safePath.startsWith('/api/v2/')) {
    res.status(400).json({ ok: false, error: `Unsupported passthrough path: ${safePath}` });
    return;
  }

  const targetBase = safePath.startsWith('/api/v0/') ? AF_DB_BASE : AF_REPORTS_BASE;
  const targetUrl = `${targetBase}${safePath}`;
  const method = String(req.method || 'GET').toUpperCase();
  const headers: Record<string, string> = {
    ...afHeaders(safePath.startsWith('/api/v0/') ? 'v0' : 'v2'),
  };

  let body: string | undefined;
  if (method !== 'GET' && method !== 'HEAD') {
    headers['Content-Type'] = String(req.headers['content-type'] || 'application/json');
    body = typeof req.body === 'string' ? req.body : JSON.stringify(req.body ?? {});
  }

  const upstream = await fetch(targetUrl, {
    method,
    headers,
    body,
  });

  const text = await upstream.text();
  res.status(upstream.status);
  res.setHeader('content-type', upstream.headers.get('content-type') || 'application/json; charset=utf-8');
  res.send(text);
}

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

  if (action === 'passthrough') {
    try {
      const session = await requireProxySession(req, res);
      if (!session) return;
      const path = String(req.query.path || '').trim();
      await proxyAppFolioAction(req, res, path);
    } catch (error) {
      logTunnelError(error, '/api?action=passthrough');
      res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Passthrough failed') });
    }
    return;
  }

  if (action === 'report') {
    try {
      const session = await requireProxySession(req, res);
      if (!session) return;
      const reportName = String(req.query.name || '').trim().replace(/[^a-z0-9_]/gi, '');
      if (!reportName) {
        res.status(400).json({ ok: false, error: 'Missing report name' });
        return;
      }
      await proxyAppFolioAction(req, res, `/api/v2/reports/${reportName}.json`);
    } catch (error) {
      logTunnelError(error, '/api?action=report');
      res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Report proxy failed') });
    }
    return;
  }

  if (action === 'admin_rate_limits') {
    const isClear = /^(1|true|yes)$/i.test(String((req.body as any)?.clear_all || req.query.clear_all || '').trim());
    const ipAddress = String((req.body as any)?.ip_address || req.query.ip_address || '').trim();
    res.json({
      ok: true,
      data: [],
      message: isClear
        ? 'Rate limits cleared (compatibility shim)'
        : (ipAddress ? `Rate limit cleared for ${ipAddress} (compatibility shim)` : 'Rate limit list unavailable in local backend shim'),
      compatibility_shim: true,
    });
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

app.get('/api/local/webhook_events', async (req: Request, res: Response) => {
  try {
    await ensureWebhookEventsTable();
    const limit = parseLimit(req.query.limit, 100, 1000);
    const rows = await queryClient`
      select id, event_uuid, topic, event_type, resource_type, resource_id,
             received_at, processed_at, processing_status
      from webhook_events
      order by received_at desc
      limit ${limit}
    `;
    const results = (rows as any[]).map((row) => ({
      id: Number(row?.id || 0) || 0,
      event_uuid: String(row?.event_uuid || ''),
      topic: String(row?.topic || ''),
      event_type: String(row?.event_type || ''),
      resource_type: String(row?.resource_type || ''),
      resource_id: String(row?.resource_id || ''),
      received_at: asIso(row?.received_at),
      processed_at: asIso(row?.processed_at),
      processing_status: String(row?.processing_status || ''),
    }));
    res.json({ ok: true, results, count: results.length });
  } catch (error) {
    logTunnelError(error, '/api/local/webhook_events');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Webhook events query failed') });
  }
});

app.get('/api/local/ping', async (_req: Request, res: Response) => {
  try {
    const dbOk = await pingDatabase();
    const status = dbOk ? 200 : 503;
    res.status(status).json({
      ok: dbOk,
      status,
      latency_ms: 0,
      database: dbOk ? 'up' : 'down',
      db_api: { ok: dbOk, status },
      reports_api: { ok: true, status: 200 },
      schema: { ok: true, missing_tables: [] },
      version: String(process.env.APP_VERSION || process.env.RENDER_GIT_COMMIT || 'server-local'),
    });
  } catch (error) {
    logTunnelError(error, '/api/local/ping');
    res.status(500).json({
      ok: false,
      status: 500,
      error: String((error as any)?.message || error || 'Ping failed'),
      db_api: { ok: false, status: 500 },
      reports_api: { ok: false, status: 500 },
      schema: { ok: false, missing_tables: [] },
    });
  }
});

let vendorOverrideTableEnsured = false;

async function ensureVendorOverrideTable(): Promise<void> {
  if (vendorOverrideTableEnsured) return;
  await queryClient.unsafe(`
    CREATE TABLE IF NOT EXISTS vendor_overrides (
      vendor_id TEXT PRIMARY KEY,
      category TEXT,
      trade_category TEXT,
      compliant BOOLEAN,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    )
  `);
  vendorOverrideTableEnsured = true;
}

app.get('/api/local/recent_tasks', async (req: Request, res: Response) => {
  try {
    const session = await requireProxySession(req, res);
    if (!session) return;

    const upstream = await fetch(`${AF_DB_BASE}/api/v0/tasks?page[size]=50`, {
      headers: afHeaders('v0'),
    });
    const text = await upstream.text();
    let parsed: any = {};
    try { parsed = text ? JSON.parse(text) : {}; } catch { parsed = {}; }
    const results = Array.isArray(parsed?.data)
      ? parsed.data
      : (Array.isArray(parsed?.results) ? parsed.results : []);

    if (!upstream.ok) {
      res.status(upstream.status).json({ ok: false, error: String(parsed?.error || text || `HTTP ${upstream.status}`) });
      return;
    }

    res.json({ ok: true, results, count: results.length, source: 'appfolio_live' });
  } catch (error) {
    logTunnelError(error, '/api/local/recent_tasks');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Recent tasks query failed') });
  }
});

app.get('/api/local/vendor_overrides', async (req: Request, res: Response) => {
  try {
    const session = await requireProxySession(req, res);
    if (!session) return;
    await ensureVendorOverrideTable();
    const rows = await queryClient.unsafe(`SELECT vendor_id, category, trade_category, compliant, updated_at FROM vendor_overrides ORDER BY vendor_id`);
    res.json({ ok: true, results: rows, count: (rows as any[]).length, source: 'postgres_local' });
  } catch (error) {
    logTunnelError(error, '/api/local/vendor_overrides');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Vendor overrides query failed') });
  }
});

app.post('/api/local/vendor_overrides/upsert', async (req: Request, res: Response) => {
  try {
    const session = await requireAdminSession(req, res);
    if (!session) return;
    await ensureVendorOverrideTable();

    const vendorId = String(req.body?.vendor_id || '').trim();
    if (!vendorId) {
      res.status(400).json({ ok: false, error: 'vendor_id is required' });
      return;
    }

    const category = req.body?.category == null ? null : String(req.body.category || '').trim() || null;
    const tradeCategory = req.body?.trade_category == null ? null : String(req.body.trade_category || '').trim() || null;
    const compliant = req.body?.compliant === null || req.body?.compliant === undefined || req.body?.compliant === ''
      ? null
      : !!req.body?.compliant;

    await queryClient.unsafe(
      `INSERT INTO vendor_overrides (vendor_id, category, trade_category, compliant, updated_at)
       VALUES ($1, $2, $3, $4, NOW())
       ON CONFLICT (vendor_id) DO UPDATE SET
         category = EXCLUDED.category,
         trade_category = EXCLUDED.trade_category,
         compliant = EXCLUDED.compliant,
         updated_at = NOW()`,
      [vendorId, category, tradeCategory, compliant],
    );

    res.json({ ok: true, vendor_id: vendorId, category, trade_category: tradeCategory, compliant });
  } catch (error) {
    logTunnelError(error, '/api/local/vendor_overrides/upsert');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Vendor override upsert failed') });
  }
});

app.get('/api/local/system_health', async (_req: Request, res: Response) => {
  try {
    const dbOk = await pingDatabase();
    const checks: Array<{ key: string; label: string; status: 'green' | 'yellow' | 'red'; detail: string }> = [
      {
        key: 'database',
        label: 'Postgres Connectivity',
        status: dbOk ? 'green' : 'red',
        detail: dbOk ? 'Postgres reachable' : 'Postgres unavailable',
      },
      {
        key: 'local_api',
        label: 'Local API Routes',
        status: 'green',
        detail: 'Using local /api/local routes for operational reads',
      },
      {
        key: 'work_orders_uuid',
        label: 'Work Orders UUID Path',
        status: 'green',
        detail: 'Work orders sourced from local v0-backed table with UUID support',
      },
    ];

    const summary = {
      green_count: checks.filter((c) => c.status === 'green').length,
      yellow_count: checks.filter((c) => c.status === 'yellow').length,
      red_count: checks.filter((c) => c.status === 'red').length,
    };
    const overall = summary.red_count > 0 ? 'red' : (summary.yellow_count > 0 ? 'yellow' : 'green');

    res.json({
      ok: true,
      status: overall,
      generated_at: new Date().toISOString(),
      summary,
      checks,
      debug: {
        debug_query: [
          'select count(*) from appfolio_work_orders;',
          'select count(*) from appfolio_properties;',
          'select count(*) from appfolio_unit_inspections;',
        ],
        payload: {
          db_ok: dbOk,
          scheduler_endpoints: syncSchedulerState.endpoints,
          scheduler_rejected_endpoints: syncSchedulerState.rejectedEndpoints,
        },
      },
    });
  } catch (error) {
    logTunnelError(error, '/api/local/system_health');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'System health failed') });
  }
});

// Enforce PM scoped property-group filtering on all local data endpoints.
app.use('/api/local', pmScopeMiddleware);

app.get('/api/local/work_orders', async (req: Request, res: Response) => {
  try {
    const days = parseDays(req.query.days, 3650, 3650);
    const limit = parseLimit(req.query.limit, 10000, 20000);
    const propertyGroupId = getPropertyGroupFilter(req);
    let rows: any[] = [];
    try {
      rows = propertyGroupId
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
    } catch (error) {
      const message = String((error as any)?.message || error || '');
      const code = String((error as any)?.code || '');
      if (!(code === '42703' && /work_order_uuid/i.test(message))) {
        throw error;
      }

      rows = propertyGroupId
        ? await queryClient`
          select id, null::text as work_order_uuid, wo_number, property_id, unit_id, property_group_id, description,
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
          select id, null::text as work_order_uuid, wo_number, property_id, unit_id, property_group_id, description,
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
    }

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
    let rows: any[] = [];
    try {
      rows = propertyGroupId
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
    } catch (error) {
      const message = String((error as any)?.message || error || '');
      const code = String((error as any)?.code || '');
      if (!(code === '42703' && /work_order_uuid/i.test(message))) {
        throw error;
      }

      rows = propertyGroupId
        ? await queryClient`
          select id, null::text as work_order_uuid, wo_number, property_id, unit_id, property_group_id, description,
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
          select id, null::text as work_order_uuid, wo_number, property_id, unit_id, property_group_id, description,
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
    }

    const results = (rows as any[]).map(normalizeWorkOrderRow);
    res.json({ ok: true, results, count: results.length, source: 'postgres_local', status: 'inactive' });
  } catch (error) {
    logTunnelError(error, '/api/local/work_orders/inactive');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local inactive work orders query failed') });
  }
});

app.get('/api/local/work_orders/:workOrderRef/notes', async (req: Request, res: Response) => {
  try {
    const workOrderRef = String(req.params.workOrderRef || '').trim();
    const forceRefresh = /^(1|true|yes|on)$/i.test(String(req.query.force_refresh || '').trim());
    if (!workOrderRef) {
      res.status(400).json({ ok: false, error: 'Missing work order reference' });
      return;
    }

    const payload = await getWorkOrderDetailCached({
      workOrderRef,
      detailType: 'notes',
      forceRefresh,
    });

    res.json({
      ok: true,
      work_order_uuid: payload.workOrderUuid,
      results: payload.results,
      source: payload.source,
      stale: payload.stale,
      fetched_at: payload.fetched_at,
      delta: payload.delta,
    });
  } catch (error) {
    logTunnelError(error, '/api/local/work_orders/:workOrderRef/notes');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local notes query failed') });
  }
});

app.get('/api/local/work_orders/:workOrderRef/attachments', async (req: Request, res: Response) => {
  try {
    const workOrderRef = String(req.params.workOrderRef || '').trim();
    const forceRefresh = /^(1|true|yes|on)$/i.test(String(req.query.force_refresh || '').trim());
    if (!workOrderRef) {
      res.status(400).json({ ok: false, error: 'Missing work order reference' });
      return;
    }

    const payload = await getWorkOrderDetailCached({
      workOrderRef,
      detailType: 'attachments',
      forceRefresh,
    });

    res.json({
      ok: true,
      work_order_uuid: payload.workOrderUuid,
      results: payload.results,
      source: payload.source,
      stale: payload.stale,
      fetched_at: payload.fetched_at,
      delta: payload.delta,
    });
  } catch (error) {
    logTunnelError(error, '/api/local/work_orders/:workOrderRef/attachments');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local attachments query failed') });
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

app.get('/api/local/vendors', async (req: Request, res: Response) => {
  try {
    const limit = parseLimit(req.query.limit, 2500, 10000);
    const propertyGroupId = getPropertyGroupFilter(req);
    const vendorPolicy = await readReportPolicy('v0:work_orders:vendors', String(propertyGroupId || ''));
    const vendorPolicyCheckedAt = Date.parse(String(vendorPolicy?.last_checked_at || ''));
    if (!vendorPolicy || !Number.isFinite(vendorPolicyCheckedAt) || (Date.now() - vendorPolicyCheckedAt) > REPORT_POLICY_TTL_MS) {
      void monitorVendorSlowPath(String(propertyGroupId || ''));
    }

    const rows = propertyGroupId
      ? await queryClient`
        select
          coalesce(nullif(vendor_id, ''), nullif(raw_json->>'vendor_id', ''), nullif(raw_json->>'VendorId', '')) as vendor_id,
          coalesce(
            nullif(vendor_name, ''),
            nullif(raw_json->>'vendor_name', ''),
            nullif(raw_json->>'VendorName', ''),
            nullif(assigned_user_name, ''),
            nullif(raw_json->>'assigned_user_name', ''),
            nullif(raw_json->>'AssignedUserName', ''),
            nullif(vendor_id, ''),
            nullif(raw_json->>'vendor_id', ''),
            nullif(raw_json->>'VendorId', '')
          ) as vendor_name,
          count(*)::int as open_work_order_count,
          max(coalesce(updated_at, created_at, now())) as last_seen_at
        from appfolio_work_orders
        where property_group_id = ${propertyGroupId}
        group by 1, 2
        having coalesce(
          nullif(vendor_name, ''),
          nullif(raw_json->>'vendor_name', ''),
          nullif(raw_json->>'VendorName', ''),
          nullif(assigned_user_name, ''),
          nullif(raw_json->>'assigned_user_name', ''),
          nullif(raw_json->>'AssignedUserName', ''),
          nullif(vendor_id, ''),
          nullif(raw_json->>'vendor_id', ''),
          nullif(raw_json->>'VendorId', '')
        ) is not null
        order by vendor_name asc
        limit ${limit}
      `
      : await queryClient`
        select
          coalesce(nullif(vendor_id, ''), nullif(raw_json->>'vendor_id', ''), nullif(raw_json->>'VendorId', '')) as vendor_id,
          coalesce(
            nullif(vendor_name, ''),
            nullif(raw_json->>'vendor_name', ''),
            nullif(raw_json->>'VendorName', ''),
            nullif(assigned_user_name, ''),
            nullif(raw_json->>'assigned_user_name', ''),
            nullif(raw_json->>'AssignedUserName', ''),
            nullif(vendor_id, ''),
            nullif(raw_json->>'vendor_id', ''),
            nullif(raw_json->>'VendorId', '')
          ) as vendor_name,
          count(*)::int as open_work_order_count,
          max(coalesce(updated_at, created_at, now())) as last_seen_at
        from appfolio_work_orders
        group by 1, 2
        having coalesce(
          nullif(vendor_name, ''),
          nullif(raw_json->>'vendor_name', ''),
          nullif(raw_json->>'VendorName', ''),
          nullif(assigned_user_name, ''),
          nullif(raw_json->>'assigned_user_name', ''),
          nullif(raw_json->>'AssignedUserName', ''),
          nullif(vendor_id, ''),
          nullif(raw_json->>'vendor_id', ''),
          nullif(raw_json->>'VendorId', '')
        ) is not null
        order by vendor_name asc
        limit ${limit}
      `;

    const results = (rows as any[])
      .map((row) => ({
        vendor_id: String(row?.vendor_id || '').trim(),
        company_name: String(row?.vendor_name || '').trim(),
        vendor_name: String(row?.vendor_name || '').trim(),
        open_work_order_count: Number(row?.open_work_order_count || 0),
        last_seen_at: asIso(row?.last_seen_at),
        _source: 'postgres_local',
      }))
      .filter((row) => !!row.company_name || !!row.vendor_id);

    res.json({
      ok: true,
      results,
      count: results.length,
      source: 'postgres_local',
      latency_policy: {
        dataset_key: 'v0:work_orders:vendors',
        property_group_id: String(propertyGroupId || ''),
        last_status: String(vendorPolicy?.last_status || ''),
        last_latency_ms: Number(vendorPolicy?.last_latency_ms || 0) || 0,
        prefer_local_until: asIso(vendorPolicy?.prefer_local_until),
      },
    });
  } catch (error) {
    logTunnelError(error, '/api/local/vendors');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Local vendors query failed') });
  }
});

app.get('/api/local/property_group_directory', async (req: Request, res: Response) => {
  try {
    const limit = parseLimit(req.query.limit, 1000, 5000);
    let results: Array<{ property_group_uuid: string; property_group_name: string }> = [];
    try {
      const rows = await queryClient`
        select coalesce(uuid, id) as group_uuid, coalesce(name, id) as group_name, raw_json
        from appfolio_property_groups
        order by coalesce(name, id) asc
        limit ${limit}
      `;

      results = (rows as any[])
        .map((row) => ({
          property_group_uuid: String(row?.group_uuid || '').trim(),
          property_group_name: resolvePropertyGroupDisplayName(row),
        }))
        .filter((row) => !!row.property_group_uuid);
    } catch (tableError) {
      const message = String((tableError as any)?.message || tableError || '');
      const code = String((tableError as any)?.code || '');
      if (!(code === '42P01' && /appfolio_property_groups/i.test(message))) {
        throw tableError;
      }

      const propertyRows = await queryClient`
        select property_group_id, raw_json
        from appfolio_properties
        where coalesce(property_group_id, '') <> ''
      `;

      const byGroup = new Map<string, string>();
      for (const row of propertyRows as any[]) {
        const groupId = String(row?.property_group_id || '').trim();
        if (!groupId || byGroup.has(groupId)) continue;
        const raw = (row?.raw_json && typeof row.raw_json === 'object') ? row.raw_json : {};
        const nameHint = String(
          raw?.property_group || raw?.group_name || raw?.portfolio || raw?.portfolio_name || groupId,
        ).trim() || groupId;
        byGroup.set(groupId, nameHint);
      }

      results = Array.from(byGroup.entries())
        .map(([property_group_uuid, property_group_name]) => ({ property_group_uuid, property_group_name }))
        .sort((a, b) => a.property_group_name.localeCompare(b.property_group_name))
        .slice(0, limit);
    }

    res.json({ ok: true, results, count: results.length, source: 'postgres_local' });
  } catch (error) {
    logTunnelError(error, '/api/local/property_group_directory');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Property group directory failed') });
  }
});

app.get('/api/local/property_groups', async (req: Request, res: Response) => {
  try {
    const limit = parseLimit(req.query.limit, 1000);
    const includeInactive = String(req.query.include_inactive || '').toLowerCase() === 'true';
    const propertyGroupId = getPropertyGroupFilter(req);

    let groupsRows: any[] = [];
    try {
      groupsRows = propertyGroupId
        ? await queryClient`
          select id, uuid, name, type, property_ids, raw_json, last_updated_at, cached_at
          from appfolio_property_groups
          where coalesce(uuid, id) = ${propertyGroupId}
             or id = ${propertyGroupId}
          order by name asc
          limit ${limit}
        `
        : await queryClient`
          select id, uuid, name, type, property_ids, raw_json, last_updated_at, cached_at
          from appfolio_property_groups
          order by name asc
          limit ${limit}
        `;
    } catch (error) {
      const message = String((error as any)?.message || error || '');
      const code = String((error as any)?.code || '');
      if (!(code === '42P01' && /appfolio_property_groups/i.test(message))) {
        throw error;
      }
      groupsRows = [];
    }

    if (groupsRows.length > 0) {
      const groupIdToPropertyIds = new Map<string, string[]>();
      const groupIdToNameHints = new Map<string, string>();
      const allProperties = await queryClient`
        select id, property_group_id, raw_json
        from appfolio_properties
      `;
      for (const row of allProperties as any[]) {
        const gid = String(row?.property_group_id || '').trim();
        const pid = String(row?.id || '').trim();
        if (!gid || !pid) continue;
        const arr = groupIdToPropertyIds.get(gid) || [];
        if (!arr.includes(pid)) arr.push(pid);
        groupIdToPropertyIds.set(gid, arr);

        const raw = row?.raw_json || {};
        const nameHint = String(
          raw?.property_group || raw?.group_name || raw?.portfolio || raw?.portfolio_name || '',
        ).trim();
        if (nameHint && !groupIdToNameHints.get(gid)) {
          groupIdToNameHints.set(gid, nameHint);
        }
      }

      const results = (groupsRows as any[])
        .map((group) => {
          const rawIds = Array.isArray(group?.property_ids) ? group.property_ids : [];
          const explicitIds = rawIds.map((value: any) => {
            if (typeof value === 'string') return value.trim();
            if (value && typeof value === 'object') return String(value.Id || value.id || value.PropertyId || value.property_id || '').trim();
            return String(value || '').trim();
          }).filter(Boolean);

          const groupId = String(group?.id || '').trim();
          const groupUuid = String(group?.uuid || '').trim();
          const keyCandidates = [groupUuid, groupId].filter(Boolean);
          const inferredIds = keyCandidates.flatMap((key) => groupIdToPropertyIds.get(key) || []);
          const mergedPropertyIds = Array.from(new Set([...explicitIds, ...inferredIds]));
          const inferredName = groupIdToNameHints.get(groupId) || groupIdToNameHints.get(groupUuid) || '';
          const name = resolvePropertyGroupDisplayName(group, inferredName) || groupId;
          return {
            id: String(groupUuid || groupId || ''),
            Id: String(groupUuid || groupId || ''),
            group_id: groupId,
            group_uuid: groupUuid,
            name,
            Name: name,
            type: String(group?.type || 'property_group'),
            Type: String(group?.type || 'property_group'),
            property_ids: mergedPropertyIds,
            PropertyIds: mergedPropertyIds,
            property_count: mergedPropertyIds.length,
            last_updated_at: asIso(group?.last_updated_at) || asIso(group?.cached_at),
            _source: 'postgres_local_groups_table',
          };
        })
        .sort((a, b) => String(a.name || '').localeCompare(String(b.name || '')))
        .slice(0, limit);

      res.json({ ok: true, results, count: results.length, source: 'postgres_local_groups_table' });
      return;
    }

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
        group_id: group.id,
        group_uuid: group.id,
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

app.post('/api/local/property_groups/sync', async (req: Request, res: Response) => {
  try {
    const auth = String(req.headers.authorization || '');
    const token = auth.startsWith('Bearer ') ? auth.slice(7).trim() : '';
    if (!token) {
      res.status(401).json({ ok: false, error: 'Missing bearer token' });
      return;
    }

    const session = await deviceAuthHandlers.getTrustedDeviceSession(token);
    if (!session) {
      res.status(401).json({ ok: false, error: 'Invalid session' });
      return;
    }

    const maxPages = Math.max(0, Number(req.body?.maxPages ?? 0) || 0);
    const triggerType = 'manual_ui';
    const { runSync } = await import('./sync/syncRunner.ts');

    const groupsSummary = await runSync({ endpointKey: 'v0:property_groups', triggerType, maxPages });
    const propertiesSummary = await runSync({ endpointKey: 'v0:properties', triggerType, maxPages });

    res.json({
      ok: true,
      synced: true,
      endpointKeys: ['v0:property_groups', 'v0:properties'],
      summaries: {
        property_groups: groupsSummary,
        properties: propertiesSummary,
      },
    });
  } catch (error) {
    logTunnelError(error, '/api/local/property_groups/sync');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Property group sync failed') });
  }
});

app.post('/api/local/bootstrap_sync', async (req: Request, res: Response) => {
  try {
    const auth = String(req.headers.authorization || '');
    const token = auth.startsWith('Bearer ') ? auth.slice(7).trim() : '';
    if (!token) {
      res.status(401).json({ ok: false, error: 'Missing bearer token' });
      return;
    }

    const session = await deviceAuthHandlers.getTrustedDeviceSession(token);
    if (!session) {
      res.status(401).json({ ok: false, error: 'Invalid session' });
      return;
    }

    const maxPages = Math.max(0, Number(req.body?.maxPages ?? 0) || 0);
    const lookbackDays = Math.max(1, Math.min(3650, Number(req.body?.lookback_days ?? 180) || 180));
    const forceLookback = String(req.body?.force_lookback ?? 'false').toLowerCase() === 'true'
      || req.body?.force_lookback === true;
    const triggerType = 'manual_ui';
    const defaultEndpoints = [
      'v0:properties',
      'v0:property_groups',
      'v0:units',
      'v0:work_orders',
      'v2:tenant_directory',
      'v2:unit_inspection',
      'v2:unit_turn_detail',
      'v2:unit_vacancy',
    ];

    const requestedEndpoints = Array.isArray(req.body?.endpoints)
      ? req.body.endpoints.map((value: unknown) => String(value || '').trim()).filter(Boolean)
      : defaultEndpoints;
    const endpointSelection = sanitizeSyncEndpoints(requestedEndpoints, defaultEndpoints, 'bootstrap_sync');
    const finalEndpoints = endpointSelection.accepted;

    const { runSync } = await import('./sync/syncRunner.ts');
    const summaries: Record<string, unknown> = {};
    for (const endpointKey of finalEndpoints) {
      const useLookbackControls = endpointKey === 'v0:work_orders';
      summaries[endpointKey] = await runSync({
        endpointKey,
        triggerType,
        maxPages,
        lookbackDays: useLookbackControls ? lookbackDays : 180,
        forceLookback: useLookbackControls ? forceLookback : false,
      });
    }

    res.json({
      ok: true,
      synced: true,
      endpointKeys: finalEndpoints,
      rejected_endpoint_keys: endpointSelection.rejected,
      summaries,
    });
  } catch (error) {
    logTunnelError(error, '/api/local/bootstrap_sync');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Bootstrap sync failed') });
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
  const days = parseDays(req.query.days, 90);
  const limit = parseLimit(req.query.limit, 3000);
  const statusFilter = String(req.query.status || '').trim().toLowerCase();

  try {
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
    const message = String((error as any)?.message || error || 'Local turns query failed');
    const code = String((error as any)?.code || '');

    if (code === '42P01' && /appfolio_unit_turn_details/i.test(message)) {
      try {
        const fallbackRows = await queryClient`
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
            t.updated_at
          from unit_turn_tracker t
          where coalesce(t.updated_at, t.created_at, now()) >= now() - (${days}::int * interval '1 day')
          order by coalesce(t.updated_at, t.created_at) desc
          limit ${limit}
        `;

        const filtered = (fallbackRows as any[]).filter((row) => {
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
          move_out_date: asIso(row.metadata?.move_out_date),
          move_in_date: asIso(row.metadata?.move_in_date),
          inspection_date: asIso(row.metadata?.inspection_date),
          expected_move_in_date: asIso(row.metadata?.move_in_date),
          turn_end_date: asIso(row.closed_at),
          target_days_to_complete: 0,
          total_days_to_complete: 0,
          labor_from_work_orders: '',
          purchase_orders_from_work_orders: '',
          billables_from_work_orders: '',
          inventory_from_work_orders: '',
          total_billed: '',
          unit_turn_status: String(row.status || ''),
          reference_user: String(row.metadata?.reference_user || ''),
          milestones: {},
          linked_work_orders: [],
          _source: 'postgres_local_fallback',
        }));

        res.json({ ok: true, results, count: results.length, source: 'postgres_local_fallback' });
        return;
      } catch (fallbackError) {
        logTunnelError(fallbackError, '/api/local/turns:fallback');
      }
    }

    logTunnelError(error, '/api/local/turns');
    res.status(500).json({ ok: false, error: message });
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
    let rows: any[] = [];
    try {
      rows = propertyGroupId
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
    } catch (error) {
      const message = String((error as any)?.message || error || '');
      const code = String((error as any)?.code || '');
      if (!(code === '42703' && /work_order_uuid/i.test(message))) {
        throw error;
      }

      rows = propertyGroupId
        ? await queryClient`
          select id, null::text as work_order_uuid, wo_number, property_id, unit_id, property_group_id, description,
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
          select id, null::text as work_order_uuid, wo_number, property_id, unit_id, property_group_id, description,
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
    }

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
    const activeOnly = /^(1|true|yes|on)$/i.test(String(req.query.active_only || '1').trim());
    const groupKey = String(propertyGroupId || '');
    const inspectionPolicy = await readReportPolicy('v2:unit_inspection', groupKey);
    const preferLocalUntilMs = Date.parse(String(inspectionPolicy?.prefer_local_until || ''));
    const lastCheckedAtMs = Date.parse(String(inspectionPolicy?.last_checked_at || ''));
    const preferLocal = Number.isFinite(preferLocalUntilMs) && preferLocalUntilMs > Date.now();

    if (!preferLocal) {
      const fastLive = await fetchLiveInspectionRows(groupKey, REPORT_FAST_MS);
      if (fastLive.ok && fastLive.latencyMs <= REPORT_FAST_MS) {
        await upsertReportPolicy('v2:unit_inspection', groupKey, 'fast', fastLive.latencyMs, {
          mode: 'fast_live',
          status: fastLive.status,
          row_count: fastLive.rows.length,
        }, null);

        let liveResults = mapLiveInspectionRows(fastLive.rows);
        if (activeOnly) {
          const todayIso = new Date().toISOString().slice(0, 10);
          liveResults = liveResults.filter((row) => {
            const moveIn = String(row.move_in_date || '').slice(0, 10);
            if (moveIn && moveIn > todayIso) return false;
            const moveOut = String(row.move_out_date || '').slice(0, 10);
            if (moveOut && moveOut < todayIso) return false;
            return true;
          });
        }

        res.json({
          ok: true,
          results: liveResults,
          count: liveResults.length,
          source: 'appfolio_live_fast',
          latency_ms: fastLive.latencyMs,
          property_group_id: groupKey,
          latency_policy: {
            dataset_key: 'v2:unit_inspection',
            status: 'fast',
            threshold_fast_ms: REPORT_FAST_MS,
            threshold_slow_ms: REPORT_SLOW_MS,
          },
        });
        return;
      }

      if (fastLive.ok) {
        await upsertReportPolicy('v2:unit_inspection', groupKey, 'warm', fastLive.latencyMs, {
          mode: 'fast_gate',
          status: fastLive.status,
        }, null);
      }

      void monitorInspectionSlowPath(groupKey);
    } else if (!Number.isFinite(lastCheckedAtMs) || (Date.now() - lastCheckedAtMs) > REPORT_POLICY_TTL_MS) {
      void monitorInspectionSlowPath(groupKey);
    }

    let rows: any[] = [];
    try {
      rows = propertyGroupId
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
    } catch (error) {
      const message = String((error as any)?.message || error || '');
      const code = String((error as any)?.code || '');
      if (!(code === '42P01' && /appfolio_unit_inspections/i.test(message))) {
        throw error;
      }
      rows = [];
    }

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

    let results = (rows as any[]).map((row) => ({
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

    if (activeOnly) {
      const todayIso = new Date().toISOString().slice(0, 10);
      results = results.filter((row) => {
        const moveIn = String(row.move_in_date || '').slice(0, 10);
        if (moveIn && moveIn > todayIso) return false;

        const moveOut = String(row.move_out_date || '').slice(0, 10);
        if (moveOut && moveOut < todayIso) return false;

        return true;
      });
    }

    const currentPolicy = await readReportPolicy('v2:unit_inspection', groupKey);
    res.json({
      ok: true,
      results,
      count: results.length,
      source: 'postgres_local',
      property_group_id: groupKey,
      latency_policy: {
        dataset_key: 'v2:unit_inspection',
        last_status: String(currentPolicy?.last_status || ''),
        last_latency_ms: Number(currentPolicy?.last_latency_ms || 0) || 0,
        prefer_local_until: asIso(currentPolicy?.prefer_local_until),
        threshold_fast_ms: REPORT_FAST_MS,
        threshold_slow_ms: REPORT_SLOW_MS,
      },
    });
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

function normalizeScopeUuid(value: unknown): string {
  const raw = String(value || '').trim().toLowerCase();
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(raw)
    ? raw
    : '';
}

async function requireAdminSession(req: Request, res: Response): Promise<any | null> {
  const auth = String(req.headers.authorization || '');
  const token = auth.startsWith('Bearer ') ? auth.slice(7).trim() : '';
  if (!token) {
    res.status(401).json({ ok: false, error: 'Missing bearer token' });
    return null;
  }

  const session = await deviceAuthHandlers.getTrustedDeviceSession(token);
  if (!session) {
    res.status(401).json({ ok: false, error: 'Invalid session' });
    return null;
  }

  const role = String(session.role || '').toLowerCase();
  if (role !== 'full' && role !== 'manager' && role !== 'advanced::manager' && role !== 'advanced_manager') {
    res.status(403).json({ ok: false, error: 'Admin access required' });
    return null;
  }

  return session;
}

async function ensurePmAccountTables(): Promise<void> {
  if (pmAccountTablesEnsured) return;
  await queryClient.unsafe(`
    CREATE TABLE IF NOT EXISTS pm_proxy_users (
      user_uuid TEXT PRIMARY KEY,
      email TEXT NOT NULL,
      full_name TEXT,
      phone TEXT,
      property_group_uuid TEXT,
      roles JSONB DEFAULT '["pm_readonly"]'::jsonb,
      active BOOLEAN DEFAULT TRUE,
      raw_json JSONB DEFAULT '{}'::jsonb,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
  `);

  await queryClient.unsafe(`
    CREATE TABLE IF NOT EXISTS pm_proxy_user_scopes (
      id BIGSERIAL PRIMARY KEY,
      user_uuid TEXT NOT NULL,
      property_group_uuid TEXT NOT NULL,
      is_primary BOOLEAN DEFAULT FALSE,
      active BOOLEAN DEFAULT TRUE,
      source TEXT,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      UNIQUE (user_uuid, property_group_uuid)
    );
  `);

  await queryClient.unsafe(`
    CREATE INDEX IF NOT EXISTS pm_proxy_users_email_idx ON pm_proxy_users(lower(email));
    CREATE INDEX IF NOT EXISTS pm_proxy_users_group_idx ON pm_proxy_users(property_group_uuid);
    CREATE INDEX IF NOT EXISTS pm_proxy_user_scopes_user_idx ON pm_proxy_user_scopes(user_uuid);
    CREATE INDEX IF NOT EXISTS pm_proxy_user_scopes_group_idx ON pm_proxy_user_scopes(property_group_uuid);
  `);
  pmAccountTablesEnsured = true;
}

async function ensureOtpSettingsTable(): Promise<void> {
  if (otpSettingsTableEnsured) return;
  await queryClient.unsafe(`
    CREATE TABLE IF NOT EXISTS proxy_config (
      key TEXT PRIMARY KEY,
      value TEXT,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
  `);
  otpSettingsTableEnsured = true;
}

let pmAccountTablesEnsured = false;
let otpSettingsTableEnsured = false;

app.get('/api/local/pm_users', async (req: Request, res: Response) => {
  try {
    const session = await requireAdminSession(req, res);
    if (!session) return;
    await ensurePmAccountTables();

    const rows = await queryClient.unsafe(`
      SELECT
        u.user_uuid,
        u.email,
        u.full_name,
        u.phone,
        COALESCE(u.property_group_uuid, '') AS property_group_uuid,
        CASE WHEN COALESCE(u.active, TRUE) THEN 1 ELSE 0 END AS active,
        COALESCE(
          string_agg(
            DISTINCT s.property_group_uuid,
            ',' ORDER BY s.property_group_uuid
          ) FILTER (WHERE COALESCE(s.active, TRUE) = TRUE),
          ''
        ) AS scope_uuids
      FROM pm_proxy_users u
      LEFT JOIN pm_proxy_user_scopes s ON s.user_uuid = u.user_uuid
      GROUP BY u.user_uuid, u.email, u.full_name, u.phone, u.property_group_uuid, u.active
      ORDER BY lower(u.email)
    `);

    res.json({ ok: true, users: rows, count: (rows as any[]).length });
  } catch (error) {
    logTunnelError(error, '/api/local/pm_users');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Failed to load PM users') });
  }
});

app.post('/api/local/pm_users/upsert', async (req: Request, res: Response) => {
  try {
    const session = await requireAdminSession(req, res);
    if (!session) return;
    await ensurePmAccountTables();

    const email = String(req.body?.email || '').trim().toLowerCase();
    const fullName = String(req.body?.full_name || '').trim();
    const phone = String(req.body?.phone || '').trim();
    const active = req.body?.active !== false;
    const primaryScope = normalizeScopeUuid(req.body?.property_group_uuid || '');
    const scopeInput = Array.isArray(req.body?.scope_uuids) ? req.body.scope_uuids : [];
    const scopeSet = new Set<string>(scopeInput.map((value: unknown) => normalizeScopeUuid(value)).filter(Boolean));
    if (primaryScope) scopeSet.add(primaryScope);
    const scopes = Array.from(scopeSet.values());

    if (!email) {
      res.status(400).json({ ok: false, error: 'Email is required' });
      return;
    }
    if (!fullName) {
      res.status(400).json({ ok: false, error: 'Full name is required' });
      return;
    }
    if (!phone) {
      res.status(400).json({ ok: false, error: 'Phone is required' });
      return;
    }
    if (!primaryScope) {
      res.status(400).json({ ok: false, error: 'Primary property group UUID is required' });
      return;
    }
    if (!scopes.length) {
      res.status(400).json({ ok: false, error: 'At least one valid property group scope is required' });
      return;
    }

    const userUuidRaw = String(req.body?.user_uuid || '').trim();
    const userUuid = userUuidRaw || (`pm-${email.replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 48)}`);

    await queryClient.unsafe(
      `INSERT INTO pm_proxy_users (user_uuid, email, full_name, phone, property_group_uuid, roles, active, raw_json, updated_at)
       VALUES ($1, $2, $3, $4, $5, '["pm_readonly"]'::jsonb, $6, '{}'::jsonb, NOW())
       ON CONFLICT (user_uuid) DO UPDATE SET
         email = EXCLUDED.email,
         full_name = EXCLUDED.full_name,
         phone = EXCLUDED.phone,
         property_group_uuid = EXCLUDED.property_group_uuid,
         active = EXCLUDED.active,
         updated_at = NOW()`,
      [userUuid, email, fullName, phone, primaryScope, active],
    );

    await queryClient.unsafe(
      `UPDATE pm_proxy_user_scopes SET active = FALSE, updated_at = NOW() WHERE user_uuid = $1`,
      [userUuid],
    );

    for (const scopeUuid of scopes) {
      await queryClient.unsafe(
        `INSERT INTO pm_proxy_user_scopes (user_uuid, property_group_uuid, is_primary, active, source, updated_at)
         VALUES ($1, $2, $3, TRUE, 'local_ui', NOW())
         ON CONFLICT (user_uuid, property_group_uuid) DO UPDATE SET
           is_primary = EXCLUDED.is_primary,
           active = TRUE,
           source = EXCLUDED.source,
           updated_at = NOW()`,
        [userUuid, scopeUuid, scopeUuid === primaryScope],
      );
    }

    res.json({
      ok: true,
      user: {
        user_uuid: userUuid,
        email,
        full_name: fullName,
        phone,
        property_group_uuid: primaryScope,
        scope_uuids: scopes,
        active: active ? 1 : 0,
      },
    });
  } catch (error) {
    logTunnelError(error, '/api/local/pm_users/upsert');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Failed to save PM user') });
  }
});

app.post('/api/local/pm_users/toggle', async (req: Request, res: Response) => {
  try {
    const session = await requireAdminSession(req, res);
    if (!session) return;
    await ensurePmAccountTables();

    const userUuid = String(req.body?.user_uuid || '').trim();
    const active = req.body?.active !== false;
    if (!userUuid) {
      res.status(400).json({ ok: false, error: 'user_uuid is required' });
      return;
    }

    await queryClient.unsafe(
      `UPDATE pm_proxy_users SET active = $2, updated_at = NOW() WHERE user_uuid = $1`,
      [userUuid, active],
    );
    await queryClient.unsafe(
      `UPDATE pm_proxy_user_scopes SET active = $2, updated_at = NOW() WHERE user_uuid = $1`,
      [userUuid, active],
    );

    res.json({ ok: true, user_uuid: userUuid, active: active ? 1 : 0 });
  } catch (error) {
    logTunnelError(error, '/api/local/pm_users/toggle');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Failed to toggle PM user') });
  }
});

app.get('/api/local/otp_settings', async (req: Request, res: Response) => {
  try {
    const session = await requireAdminSession(req, res);
    if (!session) return;
    await ensureOtpSettingsTable();

    const rows = await queryClient.unsafe(
      `SELECT key, value FROM proxy_config WHERE key IN ('otp_enabled','otp_allowed_domain','otp_require_pm_membership','otp_ttl_minutes')`,
    );
    const map: Record<string, string> = {};
    (rows as any[]).forEach((row) => {
      map[String(row?.key || '')] = String(row?.value || '');
    });

    const defaultDomain = String(process.env.OTP_ALLOWED_DOMAIN || 'flraz.com').replace(/^@/, '').toLowerCase();
    const defaultTtl = Math.max(3, Number(process.env.DEVICE_OTP_TTL_MINUTES || '10') || 10);

    res.json({
      ok: true,
      settings: {
        otp_enabled: (map.otp_enabled ?? '1') !== '0',
        otp_require_pm_membership: (map.otp_require_pm_membership ?? '1') !== '0',
        otp_allowed_domain: String(map.otp_allowed_domain || defaultDomain),
        otp_ttl_minutes: Math.max(3, Number(map.otp_ttl_minutes || defaultTtl) || defaultTtl),
      },
    });
  } catch (error) {
    logTunnelError(error, '/api/local/otp_settings');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Failed to load OTP settings') });
  }
});

app.post('/api/local/otp_settings', async (req: Request, res: Response) => {
  try {
    const session = await requireAdminSession(req, res);
    if (!session) return;
    await ensureOtpSettingsTable();

    const enabled = req.body?.otp_enabled !== false;
    const requireMembership = req.body?.otp_require_pm_membership !== false;
    const allowedDomain = String(req.body?.otp_allowed_domain || '').trim().replace(/^@/, '').toLowerCase();
    const ttlMinutes = Math.max(3, Number(req.body?.otp_ttl_minutes || 10) || 10);

    const entries: Array<[string, string]> = [
      ['otp_enabled', enabled ? '1' : '0'],
      ['otp_require_pm_membership', requireMembership ? '1' : '0'],
      ['otp_allowed_domain', allowedDomain],
      ['otp_ttl_minutes', String(ttlMinutes)],
    ];

    for (const [key, value] of entries) {
      await queryClient.unsafe(
        `INSERT INTO proxy_config (key, value, updated_at)
         VALUES ($1, $2, NOW())
         ON CONFLICT (key) DO UPDATE SET
           value = EXCLUDED.value,
           updated_at = NOW()`,
        [key, value],
      );
    }

    res.json({
      ok: true,
      settings: {
        otp_enabled: enabled,
        otp_require_pm_membership: requireMembership,
        otp_allowed_domain: allowedDomain,
        otp_ttl_minutes: ttlMinutes,
      },
    });
  } catch (error) {
    logTunnelError(error, '/api/local/otp_settings');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Failed to save OTP settings') });
  }
});

app.get('/api/local/proxy_config', async (req: Request, res: Response) => {
  try {
    const session = await requireAdminSession(req, res);
    if (!session) return;
    await ensureOtpSettingsTable();

    const rows = await queryClient.unsafe(
      `SELECT key, value FROM proxy_config ORDER BY key`,
    );
    res.json({ ok: true, rows: rows as any[] });
  } catch (error) {
    logTunnelError(error, '/api/local/proxy_config');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Failed to load proxy config') });
  }
});

app.post('/api/local/proxy_config/upsert', async (req: Request, res: Response) => {
  try {
    const session = await requireAdminSession(req, res);
    if (!session) return;
    await ensureOtpSettingsTable();

    const key = String(req.body?.key || '').trim();
    if (!key) {
      res.status(400).json({ ok: false, error: 'Missing config key' });
      return;
    }
    const value = String(req.body?.value ?? '');

    await queryClient.unsafe(
      `INSERT INTO proxy_config (key, value, updated_at)
       VALUES ($1, $2, NOW())
       ON CONFLICT (key) DO UPDATE SET
         value = EXCLUDED.value,
         updated_at = NOW()`,
      [key, value],
    );

    res.json({ ok: true, key, value });
  } catch (error) {
    logTunnelError(error, '/api/local/proxy_config/upsert');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Failed to save proxy config') });
  }
});

app.post('/api/local/reassignment_queue/clear_exempt', async (req: Request, res: Response) => {
  try {
    const session = await requireAdminSession(req, res);
    if (!session) return;

    const woId = String(req.body?.wo_id || '').trim();
    if (!woId) {
      res.status(400).json({ ok: false, error: 'Missing wo_id' });
      return;
    }

    await queryClient.unsafe(
      `UPDATE reassignment_queue
         SET auto_exempt = 0,
             auto_exempt_at = NULL,
             auto_exempt_by = NULL
       WHERE wo_id = $1`,
      [woId],
    );

    res.json({ ok: true, wo_id: woId });
  } catch (error) {
    logTunnelError(error, '/api/local/reassignment_queue/clear_exempt');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Failed to clear exemption') });
  }
});

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

    const endpointSelection = sanitizeSyncEndpoints(
      [String(req.body?.endpoint ?? 'v0:units').trim()],
      ['v0:units'],
      'admin_sync',
    );
    const endpointKey = endpointSelection.accepted[0] || '';
    if (!endpointKey) {
      res.status(400).json({
        ok: false,
        error: 'Requested endpoint is not allowed',
        rejected_endpoint_keys: endpointSelection.rejected,
      });
      return;
    }
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

const syncSchedulerState: {
  enabled: boolean;
  intervalMinutes: number;
  endpoints: string[];
  maxPages: number;
  runOnBoot: boolean;
  startedAt: string;
  lastTickStartedAt: string;
  lastTickFinishedAt: string;
  lastTickTriggerType: string;
  lastError: string;
  inFlightEndpoints: string[];
  endpointSummaries: Record<string, any>;
  rejectedEndpoints: string[];
} = {
  enabled: false,
  intervalMinutes: 0,
  endpoints: [],
  maxPages: 0,
  runOnBoot: false,
  startedAt: '',
  lastTickStartedAt: '',
  lastTickFinishedAt: '',
  lastTickTriggerType: '',
  lastError: '',
  inFlightEndpoints: [],
  endpointSummaries: {},
  rejectedEndpoints: [],
};

function getSchedulerEndpointGroups(endpoints: string[]): { v0: string[]; v2: string[]; other: string[]; v2Paused: boolean } {
  const groups = {
    v0: [] as string[],
    v2: [] as string[],
    other: [] as string[],
    v2Paused: true,
  };
  for (const endpoint of endpoints) {
    if (endpoint.startsWith('v0:')) groups.v0.push(endpoint);
    else if (endpoint.startsWith('v2:')) groups.v2.push(endpoint);
    else groups.other.push(endpoint);
  }
  groups.v2Paused = groups.v2.length === 0;
  return groups;
}

app.get('/api/local/sync/status', async (_req: Request, res: Response) => {
  try {
    const endpointGroups = getSchedulerEndpointGroups(syncSchedulerState.endpoints);
    const payload: Record<string, any> = {
      ok: true,
      scheduler: {
        enabled: syncSchedulerState.enabled,
        interval_minutes: syncSchedulerState.intervalMinutes,
        endpoints: syncSchedulerState.endpoints,
        max_pages: syncSchedulerState.maxPages,
        run_on_boot: syncSchedulerState.runOnBoot,
        started_at: syncSchedulerState.startedAt || null,
        last_tick_started_at: syncSchedulerState.lastTickStartedAt || null,
        last_tick_finished_at: syncSchedulerState.lastTickFinishedAt || null,
        last_tick_trigger_type: syncSchedulerState.lastTickTriggerType || null,
        in_flight_endpoints: syncSchedulerState.inFlightEndpoints,
        last_error: syncSchedulerState.lastError || null,
        endpoint_summaries: syncSchedulerState.endpointSummaries,
        rejected_endpoints: syncSchedulerState.rejectedEndpoints,
        endpoint_groups: {
          v0: endpointGroups.v0,
          v2: endpointGroups.v2,
          other: endpointGroups.other,
        },
        v2_paused: endpointGroups.v2Paused,
      },
      runs: [] as any[],
      cursors: {
        last_successful_work_orders_execution_start_cursor: null as string | null,
      },
    };

    try {
      const rows = await queryClient`
        select run_id, endpoint_key, trigger_type, status, started_at, completed_at,
               pages_completed, rows_upserted, rows_skipped, last_error, execution_start_cursor
        from sync_job_runs
        order by started_at desc
        limit 25
      `;

      payload.runs = (rows as any[]).map((row) => ({
        run_id: String(row?.run_id || ''),
        endpoint_key: String(row?.endpoint_key || ''),
        trigger_type: String(row?.trigger_type || ''),
        status: String(row?.status || ''),
        started_at: asIso(row?.started_at) || asIso(row?.created_at),
        completed_at: asIso(row?.completed_at),
        pages_completed: Number(row?.pages_completed || 0),
        rows_upserted: Number(row?.rows_upserted || 0),
        rows_skipped: Number(row?.rows_skipped || 0),
        last_error: String(row?.last_error || ''),
        execution_start_cursor: asIso(row?.execution_start_cursor) || String(row?.execution_start_cursor || ''),
      }));

      const lastWoRun = payload.runs.find((row: any) =>
        row.endpoint_key === 'v0:work_orders' && row.status === 'completed' && row.execution_start_cursor,
      );
      payload.cursors.last_successful_work_orders_execution_start_cursor = lastWoRun
        ? String(lastWoRun.execution_start_cursor || '')
        : null;
    } catch (tableErr) {
      const msg = String((tableErr as any)?.message || tableErr || '');
      if (!/sync_job_runs/i.test(msg)) throw tableErr;
      payload.ok = false;
      payload.error = 'sync_job_runs table not available yet';
    }

    res.json(payload);
  } catch (error) {
    logTunnelError(error, '/api/local/sync/status');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'Sync status failed') });
  }
});

app.get('/api/local/sync/v2_probe', async (req: Request, res: Response) => {
  try {
    const session = await requireAdminSession(req, res);
    if (!session) return;

    const timeoutMs = Math.max(5_000, Math.min(60_000, Number(req.query.timeout_ms || 20_000) || 20_000));
    const endpointBodies: Array<{ report: string; body: Record<string, unknown> }> = [
      {
        report: 'tenant_directory',
        body: {
          tenant_visibility: 'active',
          property_visibility: 'active',
          tenant_statuses: ['0', '4'],
          tenant_types: 'all',
          columns: ['property', 'property_id', 'unit', 'tenant', 'status'],
        },
      },
      {
        report: 'unit_inspection',
        body: {
          unit_visibility: 'active',
          include_blank_inspection_date: '1',
          columns: ['property', 'property_id', 'unit_name', 'last_inspection_date', 'unit_id'],
        },
      },
      {
        report: 'unit_turn_detail',
        body: {
          property_visibility: 'active',
          unit_turn_status: 'All',
          columns: ['property', 'property_id', 'unit', 'unit_turn_id', 'move_out_date'],
        },
      },
      {
        report: 'unit_vacancy',
        body: {
          property_visibility: 'active',
          columns: ['property', 'property_id', 'unit', 'unit_id', 'vacant_from', 'status'],
        },
      },
    ];

    const results: any[] = [];
    for (const endpoint of endpointBodies) {
      const startedAt = Date.now();
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), timeoutMs);
      const url = `${AF_REPORTS_BASE}/api/v2/reports/${endpoint.report}.json`;
      try {
        const response = await fetch(url, {
          method: 'POST',
          headers: afHeaders('v2'),
          body: JSON.stringify(endpoint.body),
          signal: controller.signal,
        });
        const rawBody = await response.text();
        const preview = String(rawBody || '').replace(/\s+/g, ' ').slice(0, 240);
        results.push({
          report: endpoint.report,
          ok: response.ok,
          status: response.status,
          latency_ms: Date.now() - startedAt,
          response_preview: preview,
        });
      } catch (error) {
        const err = error as any;
        const causeCode = String(err?.cause?.code || err?.code || '');
        const reason = err?.name === 'AbortError'
          ? `timeout after ${timeoutMs}ms`
          : String(err?.message || err || 'fetch failed');
        results.push({
          report: endpoint.report,
          ok: false,
          status: 0,
          latency_ms: Date.now() - startedAt,
          error: causeCode ? `${reason} (${causeCode})` : reason,
        });
      } finally {
        clearTimeout(timer);
      }
    }

    const failures = results.filter((row) => !row.ok);
    res.json({
      ok: failures.length === 0,
      reports_base: AF_REPORTS_BASE,
      timeout_ms: timeoutMs,
      failures: failures.length,
      results,
    });
  } catch (error) {
    logTunnelError(error, '/api/local/sync/v2_probe');
    res.status(500).json({ ok: false, error: String((error as any)?.message || error || 'V2 probe failed') });
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
  const enabled = !/^(0|false|no|off)$/i.test(String(process.env.SYNC_SCHEDULER_ENABLED || 'true').trim());
  syncSchedulerState.enabled = enabled;
  if (!enabled) return;

  const intervalMinutes = Math.max(5, Number(process.env.SYNC_SCHEDULER_INTERVAL_MINUTES || '30') || 30);
  const defaultEndpoints = [
    'v0:properties',
    'v0:property_groups',
    'v0:units',
    'v0:work_orders',
    'v2:tenant_directory',
    'v2:unit_inspection',
    'v2:unit_turn_detail',
    'v2:unit_vacancy',
  ];
  const configuredEndpoints = String(process.env.SYNC_SCHEDULER_ENDPOINTS || defaultEndpoints.join(','))
    .split(',')
    .map((value) => String(value || '').trim())
    .filter(Boolean);
  const endpointSelection = sanitizeSyncEndpoints(configuredEndpoints, defaultEndpoints, 'scheduler_env');
  const endpoints = endpointSelection.accepted;
  const maxPages = Math.max(0, Number(process.env.SYNC_SCHEDULER_MAX_PAGES || '0') || 0);
  const runOnBoot = !/^(0|false|no|off)$/i.test(String(process.env.SYNC_SCHEDULER_RUN_ON_BOOT || 'true').trim());
  const inFlight = new Set<string>();
  syncSchedulerState.intervalMinutes = intervalMinutes;
  syncSchedulerState.endpoints = endpoints.slice();
  syncSchedulerState.rejectedEndpoints = endpointSelection.rejected.slice();
  syncSchedulerState.maxPages = maxPages;
  syncSchedulerState.runOnBoot = runOnBoot;
  syncSchedulerState.startedAt = new Date().toISOString();

  async function runEndpoint(endpointKey: string, triggerType: string): Promise<void> {
    if (!endpointKey || inFlight.has(endpointKey)) return;
    inFlight.add(endpointKey);
    syncSchedulerState.inFlightEndpoints = Array.from(inFlight.values());
    try {
      const { runSync } = await import('./sync/syncRunner.ts');
      const summary = await runSync({ endpointKey, triggerType, maxPages });
      syncSchedulerState.endpointSummaries[endpointKey] = {
        ...summary,
        updated_at: new Date().toISOString(),
      };
      console.log('[server:sync-scheduler] completed', summary);
    } catch (error) {
      syncSchedulerState.lastError = `${endpointKey}: ${String((error as any)?.message || error)}`;
      console.error('[server:sync-scheduler] failed', endpointKey, String((error as any)?.message || error));
    } finally {
      inFlight.delete(endpointKey);
      syncSchedulerState.inFlightEndpoints = Array.from(inFlight.values());
    }
  }

  async function tick(triggerType: string): Promise<void> {
    syncSchedulerState.lastTickTriggerType = triggerType;
    syncSchedulerState.lastTickStartedAt = new Date().toISOString();
    for (const endpointKey of endpoints) {
      await runEndpoint(endpointKey, triggerType);
    }
    syncSchedulerState.lastTickFinishedAt = new Date().toISOString();
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
