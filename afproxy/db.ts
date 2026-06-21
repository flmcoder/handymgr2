// @ts-nocheck - Deno code, type checking disabled
import {
  upsertUnitsToDb,
} from "./dbUpserts.ts";

type ExecuteStatement =
  | string
  | {
    sql: string;
    args?: unknown[];
  };

type SqlResult = {
  rows?: unknown[];
  columns?: string[];
  rowsAffected?: number;
};

type SqlLikeClient = {
  execute: (stmt: ExecuteStatement) => Promise<SqlResult | unknown>;
};

function normalizeLibsqlUrl(raw: string): string {
  const value = String(raw || "").trim();
  if (!value) return "";
  // In Render/Node we set DATABASE_URL to PostgreSQL. Do not treat it as libsql.
  if (/^postgres(ql)?:\/\//i.test(value)) return "";
  return value;
}

const DB_URL = normalizeLibsqlUrl(
  Deno.env.get("SQLITE_URL") ||
  Deno.env.get("TURSO_DATABASE_URL") ||
  Deno.env.get("DATABASE_URL") ||
  "",
);
const DB_TOKEN =
  Deno.env.get("SQLITE_AUTH_TOKEN") ||
  Deno.env.get("TURSO_AUTH_TOKEN") ||
  Deno.env.get("DATABASE_AUTH_TOKEN") ||
  "";

function getEmbeddedSqliteClient(): SqlLikeClient {
  const candidate = (globalThis as any).sqlite;
  if (!candidate || typeof candidate.execute !== "function") {
    throw new Error(
      "SQLite client unavailable. Set SQLITE_URL/TURSO_DATABASE_URL (+ auth token) or provide global sqlite client.",
    );
  }
  return candidate as SqlLikeClient;
}

async function getRemoteSqlClient(): Promise<SqlLikeClient> {
  const isNodeRuntime = typeof process !== "undefined" && !!(process as any)?.versions?.node;
  const mod: any = isNodeRuntime
    ? await import("@libsql/client")
    : await import("npm:@libsql/client");
  const createClient = (mod as any).createClient;
  if (typeof createClient !== "function") {
    throw new Error("Failed to load @libsql/client createClient");
  }
  const client = createClient({
    url: DB_URL,
    authToken: DB_TOKEN || undefined,
  });
  return {
    execute: async (stmt: ExecuteStatement) => {
      if (typeof stmt === "string") {
        return await client.execute(stmt);
      }
      return await client.execute({ sql: stmt.sql, args: stmt.args || [] });
    },
  };
}

let _clientPromise: Promise<SqlLikeClient> | null = null;

async function getSqlClient(): Promise<SqlLikeClient> {
  if (_clientPromise) return await _clientPromise;
  _clientPromise = (async () => {
    if (DB_URL) {
      return await getRemoteSqlClient();
    }
    return getEmbeddedSqliteClient();
  })();
  return await _clientPromise;
}

export const sqlite: SqlLikeClient = {
  execute: async (stmt: ExecuteStatement) => {
    const client = await getSqlClient();
    return await client.execute(stmt);
  },
};

// Auth DB currently aliases to same secure DB pool. Kept as separate export for
// compatibility with existing handlers.
export const sqliteAuth: SqlLikeClient = {
  execute: async (stmt: ExecuteStatement) => {
    const client = await getSqlClient();
    return await client.execute(stmt);
  },
};

export function rowsAsObjects(result: any): any[] {
  if (!result) return [];
  if (Array.isArray(result.rows) && Array.isArray(result.columns)) {
    return result.rows.map((row: any) => {
      if (Array.isArray(row)) {
        const obj: Record<string, unknown> = {};
        for (let i = 0; i < result.columns.length; i++) {
          obj[String(result.columns[i])] = row[i];
        }
        return obj;
      }
      if (row && typeof row === "object") {
        return row;
      }
      return { value: row };
    });
  }
  if (Array.isArray(result.rows)) {
    return result.rows as any[];
  }
  if (Array.isArray(result)) {
    return result as any[];
  }
  return [];
}

let _tablesEnsured = false;

function bootstrapStatements(): string[] {
  return [
    `CREATE TABLE IF NOT EXISTS api_cache (
      cache_key TEXT PRIMARY KEY,
      data TEXT,
      cached_at INTEGER,
      ttl_ms INTEGER,
      etag TEXT,
      last_modified TEXT
    )`,
    `CREATE INDEX IF NOT EXISTS idx_api_cache_cached_at ON api_cache(cached_at)`,

    `CREATE TABLE IF NOT EXISTS proxy_config (
      key TEXT PRIMARY KEY,
      value TEXT,
      updated_at TEXT DEFAULT (datetime('now'))
    )`,
    `CREATE TABLE IF NOT EXISTS app_settings (
      key TEXT PRIMARY KEY,
      value TEXT,
      updated_at TEXT DEFAULT (datetime('now'))
    )`,

    `CREATE TABLE IF NOT EXISTS magic_link_tokens (
      token TEXT PRIMARY KEY,
      wo_id TEXT,
      phone TEXT,
      tech_id TEXT,
      tenant_phone TEXT,
      expires_at TEXT,
      used INTEGER DEFAULT 0,
      used_at TEXT,
      created_at TEXT DEFAULT (datetime('now'))
    )`,
    `CREATE INDEX IF NOT EXISTS idx_magic_link_tokens_wo_id ON magic_link_tokens(wo_id)`,

    `CREATE TABLE IF NOT EXISTS reassignment_queue (
      wo_id TEXT PRIMARY KEY,
      wo_number TEXT,
      property_id TEXT,
      property_group_uuid TEXT,
      property_addr TEXT,
      category TEXT,
      priority TEXT,
      status TEXT,
      assigned_to TEXT,
      auto_exempt INTEGER DEFAULT 0,
      auto_exempt_at TEXT,
      auto_exempt_by TEXT,
      updated_at TEXT DEFAULT (datetime('now')),
      created_at TEXT DEFAULT (datetime('now'))
    )`,
    `CREATE TABLE IF NOT EXISTS tech_grades (
      tech_id TEXT PRIMARY KEY,
      tech_name TEXT,
      tier INTEGER DEFAULT 1,
      grade REAL DEFAULT 0,
      jobs_completed INTEGER DEFAULT 0,
      no_contact_count INTEGER DEFAULT 0,
      updated_at TEXT DEFAULT (datetime('now')),
      created_at TEXT DEFAULT (datetime('now'))
    )`,

    `CREATE TABLE IF NOT EXISTS routing_capabilities (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      trade TEXT NOT NULL,
      keywords_json TEXT NOT NULL DEFAULT '[]',
      active INTEGER DEFAULT 1,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now'))
    )`,
    `CREATE INDEX IF NOT EXISTS idx_routing_capabilities_trade ON routing_capabilities(trade)`,
    `CREATE TABLE IF NOT EXISTS routing_pm_map (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      group_name TEXT NOT NULL,
      pm_name TEXT NOT NULL,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now'))
    )`,
    `CREATE INDEX IF NOT EXISTS idx_routing_pm_map_group ON routing_pm_map(group_name)`,
    `CREATE TABLE IF NOT EXISTS routing_events (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      wo_id TEXT,
      wo_number TEXT,
      event_type TEXT,
      trade TEXT,
      review_status TEXT DEFAULT 'pending',
      review_notes TEXT,
      source_json TEXT,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now'))
    )`,
    `CREATE INDEX IF NOT EXISTS idx_routing_events_status ON routing_events(review_status)`,
    `CREATE INDEX IF NOT EXISTS idx_routing_events_wo_id ON routing_events(wo_id)`,

    `CREATE TABLE IF NOT EXISTS trusted_devices (
      device_token TEXT PRIMARY KEY,
      user_name TEXT,
      role TEXT DEFAULT 'full',
      login_email TEXT,
      property_group_uuid TEXT,
      phone TEXT,
      last_seen_at TEXT,
      expires_at TEXT,
      revoked INTEGER DEFAULT 0,
      auth_source TEXT,
      created_at TEXT DEFAULT (datetime('now'))
    )`,
    `CREATE INDEX IF NOT EXISTS idx_trusted_devices_revoked ON trusted_devices(revoked)`,

    `CREATE TABLE IF NOT EXISTS pm_proxy_users (
      id TEXT,
      user_uuid TEXT PRIMARY KEY,
      email TEXT UNIQUE NOT NULL,
      full_name TEXT,
      phone TEXT,
      property_group_uuid TEXT NOT NULL,
      roles TEXT NOT NULL DEFAULT '[]',
      is_active INTEGER DEFAULT 1,
      raw_json TEXT NOT NULL DEFAULT '{}',
      active INTEGER DEFAULT 1,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now'))
    )`,

    `CREATE TABLE IF NOT EXISTS device_otps (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      email TEXT NOT NULL,
      code TEXT NOT NULL,
      used INTEGER DEFAULT 0,
      expires_at TEXT NOT NULL,
      user_name TEXT,
      created_at TEXT DEFAULT (datetime('now')),
      used_at TEXT,
      role_hint TEXT,
      property_group_uuid TEXT
    )`,
    `CREATE INDEX IF NOT EXISTS idx_device_otps_email ON device_otps(email)`,

    `CREATE TABLE IF NOT EXISTS pm_proxy_login_audit (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      user_uuid TEXT,
      email TEXT,
      role TEXT,
      property_group_uuid TEXT,
      device_token TEXT,
      created_at TEXT DEFAULT (datetime('now'))
    )`,

    `CREATE TABLE IF NOT EXISTS pm_notifications (
      uuid TEXT PRIMARY KEY,
      message TEXT NOT NULL,
      scope_group_uuid TEXT,
      created_by_role TEXT,
      created_by_user TEXT,
      created_at TEXT DEFAULT (datetime('now'))
    )`,
    `CREATE TABLE IF NOT EXISTS pm_notification_reads (
      notification_uuid TEXT NOT NULL,
      device_token TEXT NOT NULL,
      read_at TEXT DEFAULT (datetime('now')),
      PRIMARY KEY (notification_uuid, device_token)
    )`,

    `CREATE TABLE IF NOT EXISTS units (
      unit_id TEXT PRIMARY KEY,
      property_id TEXT,
      name TEXT,
      status TEXT,
      bedrooms REAL,
      bathrooms TEXT,
      address1 TEXT,
      city TEXT,
      state TEXT,
      zip TEXT,
      leasing_type TEXT,
      rent_ready INTEGER,
      hidden_at TEXT,
      last_updated_at TEXT,
      cached_at TEXT
    )`,
    `CREATE INDEX IF NOT EXISTS idx_units_property_id ON units(property_id)`
  ];
}

export async function ensureTables(): Promise<void> {
  if (_tablesEnsured) return;
  for (const sql of bootstrapStatements()) {
    try {
      await sqlite.execute(sql);
    } catch {
      // Continue best-effort bootstrap so one table does not block the app.
    }
  }
  _tablesEnsured = true;
}

export async function upsertUnits(rows: any[]): Promise<void> {
  await ensureTables();
  await upsertUnitsToDb(sqlite as any, rows || []);
}
