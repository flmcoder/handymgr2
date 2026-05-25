import process from 'node:process';
import { createClient } from '@libsql/client';

const proxyUrl = (process.env.PROXY_URL || '').trim();
const proxyBearer = (process.env.PROXY_BEARER_TOKEN || '').trim();
const proxyAdminKey = (process.env.PROXY_ADMIN_KEY || '').trim();
const mode = proxyUrl ? 'proxy' : 'turso';

let db = null;
if (mode === 'turso') {
  if (!process.env.TURSO_DATABASE_URL || !process.env.TURSO_AUTH_TOKEN) {
    console.error('Missing TURSO_DATABASE_URL or TURSO_AUTH_TOKEN.');
    console.error('Or set PROXY_URL (and PROXY_BEARER_TOKEN) to run in live proxy mode.');
    process.exit(1);
  }
  db = createClient({
    url: process.env.TURSO_DATABASE_URL,
    authToken: process.env.TURSO_AUTH_TOKEN,
  });
}

function parseDbIdentity(url) {
  try {
    const parsed = new URL(url || '');
    return {
      protocol: parsed.protocol || '',
      host: parsed.host || '',
      pathname: parsed.pathname || '',
    };
  } catch {
    return {
      protocol: '',
      host: '',
      pathname: '',
    };
  }
}

async function scalar(sql, args = []) {
  const res = await db.execute({ sql, args });
  const row = res.rows?.[0] || {};
  const firstKey = Object.keys(row)[0];
  return firstKey ? row[firstKey] : null;
}

async function rows(sql, args = []) {
  const res = await db.execute({ sql, args });
  return res.rows || [];
}

function countFromCacheStats(cacheStats, entityType) {
  const list = Array.isArray(cacheStats?.cache) ? cacheStats.cache : [];
  const entry = list.find((x) => String(x.entity_type || '').toLowerCase() === String(entityType || '').toLowerCase());
  return entry ? Number(entry.total_records || 0) : null;
}

async function proxyGet(action) {
  const sep = proxyUrl.includes('?') ? '&' : '?';
  const url = `${proxyUrl}${sep}action=${encodeURIComponent(action)}`;
  const headers = { Accept: 'application/json' };
  if (proxyBearer) headers.Authorization = `Bearer ${proxyBearer}`;
  const res = await fetch(url, { headers });
  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`Proxy GET ${action} failed: HTTP ${res.status}${body ? ` - ${body.slice(0, 180)}` : ''}`);
  }
  return await res.json();
}

async function proxyGetSafe(action) {
  try {
    return await proxyGet(action);
  } catch (err) {
    return {
      ok: false,
      error: String(err && (err.message || err) || `failed action=${action}`),
      action,
    };
  }
}

async function proxySqlScalar(query) {
  if (!proxyAdminKey) return null;
  const sep = proxyUrl.includes('?') ? '&' : '?';
  const url = `${proxyUrl}${sep}action=sql_query`;
  const headers = {
    Accept: 'application/json',
    'Content-Type': 'application/json',
  };
  if (proxyBearer) headers.Authorization = `Bearer ${proxyBearer}`;
  const res = await fetch(url, {
    method: 'POST',
    headers,
    body: JSON.stringify({ key: proxyAdminKey, query }),
  });
  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`Proxy sql_query failed: HTTP ${res.status}${body ? ` - ${body.slice(0, 180)}` : ''}`);
  }
  const data = await res.json();
  const row = (data && Array.isArray(data.rows) && data.rows[0]) || {};
  const firstKey = Object.keys(row)[0];
  return firstKey ? row[firstKey] : null;
}

const checks = {};

checks.mode = mode;

if (mode === 'turso') {
  checks.target_database = parseDbIdentity(process.env.TURSO_DATABASE_URL || '');

  checks.work_order_map_count = Number(await scalar('SELECT COUNT(*) AS c FROM work_order_map'));
  checks.work_orders_cache_count = Number(await scalar('SELECT COUNT(*) AS c FROM work_orders_cache'));
  checks.bridge_count = Number(await scalar('SELECT COUNT(*) AS c FROM appfolio_work_order_bridge'));

  checks.v0_uuid_shape_ok = Number(await scalar(
    "SELECT COUNT(*) AS c FROM work_order_map WHERE id GLOB '????????-????-????-????-????????????'"
  ));

  checks.v0_missing_work_order_number = Number(await scalar(
    "SELECT COUNT(*) AS c FROM work_order_map WHERE TRIM(COALESCE(work_order_number, '')) = ''"
  ));

  checks.v0_to_v2_number_join_matches = Number(await scalar(
    `SELECT COUNT(*) AS c
     FROM work_order_map wom
     JOIN work_orders_cache woc
       ON woc.wo_number = wom.work_order_number`
  ));

  checks.v2_without_v0_number_match = Number(await scalar(
    `SELECT COUNT(*) AS c
     FROM work_orders_cache woc
     LEFT JOIN work_order_map wom
       ON wom.work_order_number = woc.wo_number
     WHERE wom.id IS NULL`
  ));

  checks.bridge_without_v0_uuid = Number(await scalar(
    `SELECT COUNT(*) AS c
     FROM appfolio_work_order_bridge b
     LEFT JOIN work_order_map wom
       ON wom.id = b.work_order_uuid
     WHERE wom.id IS NULL`
  ));

  checks.cache_rows = Number(await scalar('SELECT COUNT(*) AS c FROM appfolio_response_cache'));
  checks.queue_rows = Number(await scalar('SELECT COUNT(*) AS c FROM appfolio_sync_queue'));
  checks.queue_backlog = Number(await scalar(
    "SELECT COUNT(*) AS c FROM appfolio_sync_queue WHERE status IN ('queued','retry')"
  ));

  checks.request_429_last_24h = Number(await scalar(
    "SELECT COUNT(*) AS c FROM appfolio_request_log WHERE status_code = 429 AND created_at >= datetime('now','-1 day')"
  ));

  checks.request_5xx_last_24h = Number(await scalar(
    "SELECT COUNT(*) AS c FROM appfolio_request_log WHERE status_code >= 500 AND created_at >= datetime('now','-1 day')"
  ));

  checks.top_error_tags_last_24h = await rows(
    `SELECT tag, COUNT(*) AS c
     FROM hm_logs
     WHERE level IN ('error','warn')
       AND server_received_at >= datetime('now','-1 day')
     GROUP BY tag
     ORDER BY c DESC
     LIMIT 10`
  );

  checks.sample_unmatched_v2 = await rows(
    `SELECT woc.id, woc.wo_number, woc.property_name, woc.status, woc.updated_at
     FROM work_orders_cache woc
     LEFT JOIN work_order_map wom
       ON wom.work_order_number = woc.wo_number
     WHERE wom.id IS NULL
     ORDER BY woc.updated_at DESC
     LIMIT 15`
  );

  checks.join_query_plan = await rows(
    `EXPLAIN QUERY PLAN
     SELECT wom.id, woc.wo_number
     FROM work_order_map wom
     JOIN work_orders_cache woc
       ON woc.wo_number = wom.work_order_number
     WHERE wom.status = 'Open'
     LIMIT 50`
  );
} else {
  checks.target_proxy = parseDbIdentity(proxyUrl);
  checks.target_proxy.url = proxyUrl;
  checks.proxy_admin_key_present = !!proxyAdminKey;

  const ping = await proxyGetSafe('ping');
  const cacheStats = await proxyGetSafe('cache_stats');
  const workOrdersProbe = await proxyGetSafe('work_orders&days=30');
  const propertyMapProbe = await proxyGetSafe('property_map');
  checks.proxy_ping = {
    ok: !!ping.ok,
    version: ping.version || ping.proxy || '',
    database: ping.database || '',
  };

  checks.cache_stats = {
    ok: !!cacheStats.ok,
    cache_entities: Array.isArray(cacheStats.cache) ? cacheStats.cache.length : 0,
    work_orders_records: countFromCacheStats(cacheStats, 'work_orders'),
    property_map_records: countFromCacheStats(cacheStats, 'property_map'),
    cache_stats_error: cacheStats.ok ? '' : String(cacheStats.error || ''),
  };

  checks.proxy_probes = {
    work_orders_results: Array.isArray(workOrdersProbe.results) ? workOrdersProbe.results.length : null,
    property_map_results: Array.isArray(propertyMapProbe.results) ? propertyMapProbe.results.length : null,
  };

  // Exact metrics if admin key is available; otherwise fallback to cache_stats estimates.
  if (proxyAdminKey) {
    checks.work_orders_cache_count = Number(await proxySqlScalar('SELECT COUNT(*) AS c FROM work_orders_cache'));
    checks.work_order_map_count = Number(await proxySqlScalar('SELECT COUNT(*) AS c FROM work_order_map'));
    checks.queue_backlog = Number(await proxySqlScalar("SELECT COUNT(*) AS c FROM appfolio_sync_queue WHERE status IN ('queued','retry')"));
    checks.request_429_last_24h = Number(await proxySqlScalar("SELECT COUNT(*) AS c FROM appfolio_request_log WHERE status_code = 429 AND created_at >= datetime('now','-1 day')"));
  } else {
    const woFromCache = countFromCacheStats(cacheStats, 'work_orders');
    const pmFromCache = countFromCacheStats(cacheStats, 'property_map');
    const woFallback = Array.isArray(workOrdersProbe.results) ? workOrdersProbe.results.length : 0;
    const pmFallback = Array.isArray(propertyMapProbe.results) ? propertyMapProbe.results.length : 0;
    checks.work_orders_cache_count = Number(woFromCache != null ? woFromCache : woFallback || 0);
    checks.work_order_map_count = Number(pmFromCache != null ? pmFromCache : pmFallback || 0);
    checks.queue_backlog = null;
    checks.request_429_last_24h = null;
  }
}

const workOrdersCacheOk = checks.work_orders_cache_count > 0;
const workOrderMapOk = checks.work_order_map_count > 0;
const request429Ok = checks.request_429_last_24h == null ? null : checks.request_429_last_24h <= 5;
const queueBacklogOk = checks.queue_backlog == null ? null : checks.queue_backlog <= 5;

checks.health_summary = {
  pass: workOrdersCacheOk && workOrderMapOk && (request429Ok == null || request429Ok) && (queueBacklogOk == null || queueBacklogOk),
  checks: {
    work_orders_cache_count_gt_0: {
      pass: workOrdersCacheOk,
      value: checks.work_orders_cache_count,
      threshold: '> 0',
    },
    work_order_map_count_gt_0: {
      pass: workOrderMapOk,
      value: checks.work_order_map_count,
      threshold: '> 0',
    },
    request_429_last_24h_low_or_zero: {
      pass: request429Ok,
      value: checks.request_429_last_24h,
      threshold: '<= 5 (ideal 0)'
    },
    queue_backlog_near_zero: {
      pass: queueBacklogOk,
      value: checks.queue_backlog,
      threshold: '<= 5 (ideal 0)'
    },
  },
};

checks.diagnostics = {
  likely_db_target_mismatch:
    checks.work_orders_cache_count === 0 &&
    checks.work_order_map_count === 0 &&
    (checks.request_429_last_24h == null || checks.request_429_last_24h === 0) &&
    (checks.queue_backlog == null || checks.queue_backlog === 0),
  hint: 'If both core counts are 0 while proxy UI shows data, this shell is likely pointed at a different Turso database URL/token than the live proxy.',
  note: mode === 'proxy' && !proxyAdminKey
    ? 'PROXY_ADMIN_KEY not set: request_429_last_24h and queue_backlog are unavailable in proxy mode and may be null.'
    : '',
};

console.log(JSON.stringify(checks, null, 2));