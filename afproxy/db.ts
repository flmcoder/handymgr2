// ============================================================================
// db.ts — Turso/SQLite client, dual rowsAsObjects, full schema init,
//         cache layer (cacheGet/Set/Invalidate) + storage helpers.
//         Legacy-compatible core tables are preserved; dispatch and portal
//         tables are appended below.
// ============================================================================

import { createClient } from "npm:@libsql/client@0.14.0/web";
import {
  CACHE_TTL,
  CHUNK_LIMIT,
  PROXY_APP_VERSION,
  STORAGE_BUDGET_BYTES,
  TURSO_AUTH_TOKEN,
  TURSO_DATABASE_URL,
  TURSO_URL,
  WEBHOOK_MAX_DAYS,
  WEBHOOK_MAX_EVENTS,
} from "./config.ts";

// ── Client Initialisation ─────────────────────────────────────────────────────
// Prefer Turso if both vars are set, else fall back to Val Town's
// account-scoped global.ts SQLite.
// IMPORTANT: use global.ts (account-scoped) NOT main.ts (val-scoped).
// global.ts returns rows as arrays; Turso returns rows as objects.
// A different val-specific DB lives behind main.ts — never use it here.
//
// Two clients are exported:
//   sqlite     — general cache client; falls back to ValTown sqlite on Turso
//                transport errors (preserves read availability for non-auth ops)
//   sqliteAuth — auth-only (trusted_devices) client; Turso-strict, never falls
//                back to ValTown so session tokens never split across backends.
//                If Turso is not configured it uses the same ValTown sqlite as
//                sqlite (no alternative), but the split-brain fallback is gone.
let sqlite: any;
let sqliteAuth: any;
if (TURSO_DATABASE_URL && TURSO_AUTH_TOKEN) {
  const tursoClient = createClient({
    url: TURSO_DATABASE_URL,
    authToken: TURSO_AUTH_TOKEN,
  });

  let fallbackSqlite: any = null;
  let fallbackActive = false;

  function isTursoTransportError(err: any): boolean {
    const msg = String(err?.message || err || "").toLowerCase();
    return msg.includes("invalidcontenttype") ||
      msg.includes("error reading a body from connection") ||
      msg.includes("client error (sendrequest)") ||
      msg.includes("connection error");
  }

  async function getFallbackSqlite(): Promise<any> {
    if (fallbackSqlite) return fallbackSqlite;
    const vtMod = await import("https://esm.town/v/std/sqlite/global.ts");
    fallbackSqlite = vtMod.sqlite;
    return fallbackSqlite;
  }

  sqlite = {
    execute: async (stmt: any) => {
      if (fallbackActive) {
        const fb = await getFallbackSqlite();
        return await fb.execute(stmt);
      }
      try {
        return await tursoClient.execute(stmt);
      } catch (err: any) {
        if (!isTursoTransportError(err)) throw err;
        console.warn(
          "Turso transport error detected, switching DB client to Val Town sqlite fallback:",
          String(err?.message || err).substring(0, 220),
        );
        fallbackActive = true;
        const fb = await getFallbackSqlite();
        return await fb.execute(stmt);
      }
    },
  };

  // Auth-only client: direct Turso access, no fallback.
  // A transport error here returns 500 rather than routing session tokens
  // to a different backend (which would make valid tokens appear expired).
  sqliteAuth = {
    execute: (stmt: any) => tursoClient.execute(stmt),
  };
} else {
  const vtMod = await import("https://esm.town/v/std/sqlite/global.ts");
  sqlite = vtMod.sqlite;
  // No Turso configured — ValTown sqlite is the only option for both clients.
  sqliteAuth = vtMod.sqlite;
}
export { sqlite, sqliteAuth };

// ── rowsAsObjects ─────────────────────────────────────────────────────────────
// Dual-format helper — Val Town global.ts returns arrays, Turso returns objects.
// Both handled transparently so every handler uses the same interface.
export function rowsAsObjects(result: any): Record<string, any>[] {
  if (!result.rows || result.rows.length === 0) return [];
  const firstRow = result.rows[0];

  // Turso / object rows — already keyed, normalise via columns if present
  if (firstRow && typeof firstRow === "object" && !Array.isArray(firstRow)) {
    if (result.columns) {
      return result.rows.map((row: any) => {
        const obj: Record<string, any> = {};
        for (const col of result.columns) obj[col] = row[col];
        return obj;
      });
    }
    return result.rows as Record<string, any>[];
  }

  // Val Town array rows — convert using positional columns
  if (!result.columns) return [];
  return result.rows.map((row: any) => {
    const obj: Record<string, any> = {};
    for (let i = 0; i < result.columns.length; i++) {
      obj[result.columns[i]] = row[i];
    }
    return obj;
  });
}

// ── ensureTables ──────────────────────────────────────────────────────────────
// Creates/migrates all tables to match the current schema version.
// Uses a version sentinel in proxy_config to track schema state.
const SCHEMA_VERSION = 10;
let _tablesReady = false;

export async function ensureTables(): Promise<void> {
  if (_tablesReady) return; // only runs once per cold start

  try {
    // ── Core compatibility tables ───────────────────────────────────────────

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS api_cache (
      cache_key    TEXT PRIMARY KEY,
      entity_type  TEXT NOT NULL,
      data         TEXT NOT NULL,
      cached_at    TEXT NOT NULL DEFAULT (datetime('now')),
      expires_at   TEXT,
      record_count INTEGER DEFAULT 0
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS turn_records (
      unit_turn_id TEXT PRIMARY KEY,
      data         TEXT NOT NULL,
      updated_at   TEXT NOT NULL DEFAULT (datetime('now'))
    )`);

    // ── webhook_events migration path ───────────────────────────────────────
    // If the table exists from v7 with a different schema, rename, recreate,
    // then copy what we can. ALTER TABLE ADD COLUMN with NOT NULL DEFAULT fails
    // on older libsql, so full recreate is safer.
    const V8_WH_COLS = new Set([
      "id",
      "received_at",
      "raw_body",
      "resource_type",
      "resource_id",
      "event_type",
      "processed",
      "processed_at",
    ]);

    let needsRecreate = false;
    try {
      const info = await sqlite.execute(`PRAGMA table_info(webhook_events)`);
      const existingCols = new Set(
        rowsAsObjects(info).map((r: any) => r.name as string),
      );
      if (existingCols.size > 0) {
        for (const col of V8_WH_COLS) {
          if (!existingCols.has(col)) {
            console.log(
              `webhook_events missing column '${col}' — will recreate table`,
            );
            needsRecreate = true;
            break;
          }
        }
      }
    } catch (_) { /* PRAGMA failed — table doesn't exist yet, create below */ }

    if (needsRecreate) {
      try {
        await sqlite.execute(
          `ALTER TABLE webhook_events RENAME TO webhook_events_v7_bak`,
        );
        console.log("Renamed webhook_events → webhook_events_v7_bak");
      } catch (_) {
        try {
          await sqlite.execute(`DROP TABLE IF EXISTS webhook_events`);
        } catch (_2) {}
      }
    }

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS webhook_events (
      id            INTEGER PRIMARY KEY AUTOINCREMENT,
      received_at   TEXT    NOT NULL DEFAULT (datetime('now')),
      raw_body      TEXT    NOT NULL,
      resource_type TEXT,
      resource_id   TEXT,
      event_type    TEXT,
      human_description TEXT,
      processed     INTEGER NOT NULL DEFAULT 0,
      processed_at  TEXT
    )`);

    // Backward-compatible add for installs created before human_description existed.
    try {
      await sqlite.execute(
        `ALTER TABLE webhook_events ADD COLUMN human_description TEXT`,
      );
    } catch (_) {}

    // Resolved work-order state cache populated by webhook fetch-back flow.
    await sqlite.execute(`CREATE TABLE IF NOT EXISTS wo_states (
      id                TEXT PRIMARY KEY,
      status_code       INTEGER,
      status_text       TEXT,
      assigned_user_id  TEXT,
      assigned_user_name TEXT,
      last_activity_at  TEXT,
      last_note_at      TEXT,
      event_type        TEXT,
      resource_type     TEXT,
      wo_number         TEXT,
      property_address  TEXT,
      raw_snapshot      TEXT NOT NULL,
      fetched_at        TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at        TEXT NOT NULL DEFAULT (datetime('now'))
    )`);

    // Backward-compatible column add for existing installs upgraded from older schema.
    try {
      await sqlite.execute(
        `ALTER TABLE wo_states ADD COLUMN last_note_at TEXT`,
      );
    } catch (_) {}

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS wo_state_changes (
      id            INTEGER PRIMARY KEY AUTOINCREMENT,
      work_order_id TEXT NOT NULL,
      change_type   TEXT NOT NULL,
      old_value     TEXT,
      new_value     TEXT,
      created_at    TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY(work_order_id) REFERENCES wo_states(id)
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS trusted_devices (
      device_token TEXT PRIMARY KEY,
      user_name    TEXT,
      created_at   TEXT NOT NULL DEFAULT (datetime('now'))
    )`);
    try {
      await sqlite.execute(
        `ALTER TABLE trusted_devices ADD COLUMN role TEXT DEFAULT 'full'`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE trusted_devices ADD COLUMN login_email TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE trusted_devices ADD COLUMN property_group_uuid TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(`ALTER TABLE trusted_devices ADD COLUMN phone TEXT`);
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE trusted_devices ADD COLUMN last_seen_at TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE trusted_devices ADD COLUMN expires_at TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE trusted_devices ADD COLUMN revoked INTEGER DEFAULT 0`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE trusted_devices ADD COLUMN auth_source TEXT`,
      );
    } catch (_) {}

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS device_otps (
      id         INTEGER PRIMARY KEY AUTOINCREMENT,
      email      TEXT NOT NULL,
      code       TEXT NOT NULL,
      used       INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      expires_at TEXT NOT NULL,
      used_at    TEXT,
      user_name  TEXT
    )`);
    try {
      await sqlite.execute(`ALTER TABLE device_otps ADD COLUMN role_hint TEXT`);
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE device_otps ADD COLUMN property_group_uuid TEXT`,
      );
    } catch (_) {}

    await sqlite.execute(
      `CREATE INDEX IF NOT EXISTS idx_device_otps_email_created ON device_otps(email, created_at DESC)`,
    );

    if (needsRecreate) {
      try {
        const bakInfo = await sqlite.execute(
          `PRAGMA table_info(webhook_events_v7_bak)`,
        );
        const bakCols = new Set(
          rowsAsObjects(bakInfo).map((r: any) => r.name as string),
        );
        if (bakCols.has("raw_body")) {
          const tsCol = bakCols.has("ts") ? "ts" : "datetime('now')";
          await sqlite.execute(
            `INSERT INTO webhook_events (id, received_at, raw_body)
             SELECT id, ${tsCol}, raw_body FROM webhook_events_v7_bak`,
          );
          console.log("Migrated data from webhook_events_v7_bak");
        }
        await sqlite.execute(`DROP TABLE IF EXISTS webhook_events_v7_bak`);
        console.log("Dropped webhook_events_v7_bak");
      } catch (migErr: any) {
        console.log(
          `webhook_events migration partial: ${
            migErr.message?.substring(0, 120)
          }`,
        );
        try {
          await sqlite.execute(`DROP TABLE IF EXISTS webhook_events_v7_bak`);
        } catch (_) {}
      }
    }

    // ── Core indexes ────────────────────────────────────────────────────────
    for (
      const idx of [
        `CREATE INDEX IF NOT EXISTS idx_wh_resource   ON webhook_events(resource_type, resource_id)`,
        `CREATE INDEX IF NOT EXISTS idx_wh_received   ON webhook_events(received_at DESC)`,
        `CREATE INDEX IF NOT EXISTS idx_ws_activity   ON wo_states(status_code, last_activity_at)`,
        `CREATE INDEX IF NOT EXISTS idx_ws_updated    ON wo_states(updated_at DESC)`,
        `CREATE INDEX IF NOT EXISTS idx_ws_assignee   ON wo_states(assigned_user_id)`,
        `CREATE INDEX IF NOT EXISTS idx_wsc_wo        ON wo_state_changes(work_order_id, created_at DESC)`,
        `CREATE INDEX IF NOT EXISTS idx_wsc_type      ON wo_state_changes(change_type, created_at DESC)`,
        `CREATE INDEX IF NOT EXISTS idx_td_created    ON trusted_devices(created_at DESC)`,
        `CREATE INDEX IF NOT EXISTS idx_cache_type    ON api_cache(entity_type)`,
        `CREATE INDEX IF NOT EXISTS idx_cache_expires ON api_cache(expires_at)`,
      ]
    ) {
      try {
        await sqlite.execute(idx);
      } catch (_) {}
    }

    // ── Dispatch / portal tables ────────────────────────────────────────────
    // All use CREATE TABLE IF NOT EXISTS — safe to run on top of existing DBs.

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS reassignment_queue (
      id                   INTEGER PRIMARY KEY AUTOINCREMENT,
      wo_id                TEXT NOT NULL UNIQUE,
      wo_number            TEXT,
      property_address     TEXT,
      assigned_tech_id     TEXT,
      assigned_tech_name   TEXT,
      wo_status            TEXT,
      wo_priority          TEXT,
      wo_category          TEXT,
      first_seen_at        TEXT NOT NULL DEFAULT (datetime('now')),
      warning_sent         INTEGER DEFAULT 0,
      warning_sent_at      TEXT,
      warning_channel      TEXT,
      auto_exempt          INTEGER DEFAULT 0,
      auto_exempt_at       TEXT,
      auto_exempt_by       TEXT,
      auto_exempt_note_id  TEXT,
      grace_used           INTEGER DEFAULT 0,
      reassignment_count   INTEGER DEFAULT 0,
      last_reassigned_at   TEXT,
      escalated            INTEGER DEFAULT 0
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS magic_link_tokens (
      id               INTEGER PRIMARY KEY AUTOINCREMENT,
      token            TEXT NOT NULL UNIQUE,
      wo_id            TEXT NOT NULL,
      tech_id          TEXT NOT NULL,
      tenant_phone     TEXT NOT NULL,
      tenant_name      TEXT,
      property_address TEXT,
      tech_name        TEXT,
      created_at       TEXT NOT NULL DEFAULT (datetime('now')),
      expires_at       TEXT NOT NULL,
      used             INTEGER DEFAULT 0,
      used_at          TEXT,
      used_template    TEXT
    )`);

    try {
      await sqlite.execute(
        `ALTER TABLE magic_link_tokens ADD COLUMN short_code TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE magic_link_tokens ADD COLUMN lang_pref TEXT DEFAULT 'en'`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE magic_link_tokens ADD COLUMN scheduled_date TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE magic_link_tokens ADD COLUMN scheduled_window TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE magic_link_tokens ADD COLUMN stop_auto INTEGER DEFAULT 0`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE magic_link_tokens ADD COLUMN exempt_until TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE magic_link_tokens ADD COLUMN last_action TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE magic_link_tokens ADD COLUMN last_action_at TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE magic_link_tokens ADD COLUMN portal_opened INTEGER DEFAULT 0`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE magic_link_tokens ADD COLUMN portal_opened_at TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE magic_link_tokens ADD COLUMN meta_json TEXT`,
      );
    } catch (_) {}

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS short_links (
      code         TEXT PRIMARY KEY,
      full_url     TEXT NOT NULL,
      token        TEXT,
      clicks       INTEGER DEFAULT 0,
      created_at   TEXT NOT NULL DEFAULT (datetime('now')),
      expires_at   TEXT,
      last_hit_at  TEXT
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS portal_events (
      id           INTEGER PRIMARY KEY AUTOINCREMENT,
      token        TEXT NOT NULL,
      wo_id        TEXT,
      tech_id      TEXT,
      action       TEXT NOT NULL,
      payload      TEXT,
      ip           TEXT,
      user_agent   TEXT,
      created_at   TEXT NOT NULL DEFAULT (datetime('now'))
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS tenant_comms_log (
      id               INTEGER PRIMARY KEY AUTOINCREMENT,
      wo_id            TEXT NOT NULL,
      tech_id          TEXT NOT NULL,
      tech_name        TEXT,
      tenant_phone     TEXT NOT NULL,
      template_used    TEXT NOT NULL,
      message_body     TEXT,
      sent_at          TEXT NOT NULL DEFAULT (datetime('now')),
      rc_message_id    TEXT,
      status           TEXT DEFAULT 'sent',
      appfolio_noted   INTEGER DEFAULT 0
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS monitored_work_orders (
      wo_id       TEXT PRIMARY KEY,
      created_at  TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at  TEXT NOT NULL DEFAULT (datetime('now'))
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS wo_audit_log (
      id           INTEGER PRIMARY KEY AUTOINCREMENT,
      wo_id        TEXT NOT NULL,
      event_type   TEXT NOT NULL,
      event_data   TEXT,
      actor        TEXT DEFAULT 'system',
      created_at   TEXT NOT NULL DEFAULT (datetime('now'))
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS tech_grades (
      tech_id                TEXT PRIMARY KEY,
      tech_name              TEXT NOT NULL,
      tech_phone             TEXT,
      tier                   INTEGER DEFAULT 1,
      active                 INTEGER DEFAULT 1,
      geo_zone               TEXT DEFAULT 'central',
      total_assigned         INTEGER DEFAULT 0,
      total_completed        INTEGER DEFAULT 0,
      total_go_backs         INTEGER DEFAULT 0,
      total_auto_reassigned  INTEGER DEFAULT 0,
      total_warnings_recv    INTEGER DEFAULT 0,
      total_on_time_resp     INTEGER DEFAULT 0,
      avg_completion_hours   REAL    DEFAULT 0,
      go_back_pct            REAL    DEFAULT 0,
      reassign_pct           REAL    DEFAULT 0,
      response_rate_pct      REAL    DEFAULT 0,
      performance_score      REAL    DEFAULT 100,
      active_wo_count        INTEGER DEFAULT 0,
      load_weight            REAL    DEFAULT 1.0,
      target_share_pct       REAL    DEFAULT 33.3,
      last_assigned_at       TEXT,
      score_updated_at       TEXT,
      updated_at             TEXT NOT NULL DEFAULT (datetime('now'))
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS blast_events (
      id            INTEGER PRIMARY KEY AUTOINCREMENT,
      wo_id         TEXT NOT NULL,
      wo_number     TEXT,
      property_addr TEXT,
      category      TEXT,
      priority      TEXT,
      blasted_at    TEXT NOT NULL DEFAULT (datetime('now')),
      expires_at    TEXT NOT NULL,
      claimed_by    TEXT,
      claimed_at    TEXT,
      status        TEXT DEFAULT 'open'
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS tier2_claims (
      id             INTEGER PRIMARY KEY AUTOINCREMENT,
      blast_id       INTEGER NOT NULL,
      wo_id          TEXT NOT NULL,
      tech_id        TEXT NOT NULL,
      tech_name      TEXT,
      tech_phone     TEXT,
      sms_sent_at    TEXT,
      reply_received TEXT,
      reply_at       TEXT,
      claim_status   TEXT DEFAULT 'pending'
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS proxy_config (
      key        TEXT PRIMARY KEY,
      value      TEXT NOT NULL,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS vendor_overrides (
      vendor_id  TEXT PRIMARY KEY,
      category   TEXT,
      trade_category TEXT,
      compliant  INTEGER,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS pm_proxy_users (
      id                   TEXT,
      user_uuid            TEXT PRIMARY KEY,
      email                TEXT NOT NULL UNIQUE,
      full_name            TEXT NOT NULL DEFAULT '',
      phone                TEXT NOT NULL DEFAULT '',
      property_group_uuid  TEXT,
      roles                TEXT NOT NULL DEFAULT '[]',
      is_active            INTEGER NOT NULL DEFAULT 1,
      active               INTEGER NOT NULL DEFAULT 1,
      raw_json             TEXT NOT NULL DEFAULT '{}',
      created_at           TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at           TEXT NOT NULL DEFAULT (datetime('now'))
    )`);

    // Backward/forward-compatible columns for mixed deployments.
    try {
      await sqlite.execute(`ALTER TABLE pm_proxy_users ADD COLUMN id TEXT`);
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE pm_proxy_users ADD COLUMN user_uuid TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE pm_proxy_users ADD COLUMN roles TEXT NOT NULL DEFAULT '[]'`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE pm_proxy_users ADD COLUMN is_active INTEGER NOT NULL DEFAULT 1`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE pm_proxy_users ADD COLUMN raw_json TEXT NOT NULL DEFAULT '{}'`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE pm_proxy_users ADD COLUMN property_group_uuid TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE pm_proxy_users ADD COLUMN active INTEGER NOT NULL DEFAULT 1`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE pm_proxy_users ADD COLUMN created_at TEXT NOT NULL DEFAULT (datetime('now'))`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE pm_proxy_users ADD COLUMN updated_at TEXT NOT NULL DEFAULT (datetime('now'))`,
      );
    } catch (_) {}

    // Normalize identity and activity fields across legacy/new schemas.
    try {
      await sqlite.execute(`UPDATE pm_proxy_users
                            SET id = COALESCE(NULLIF(id, ''), NULLIF(user_uuid, ''), lower(email))
                            WHERE id IS NULL OR id = ''`);
    } catch (_) {}
    try {
      await sqlite.execute(`UPDATE pm_proxy_users
                            SET user_uuid = COALESCE(NULLIF(user_uuid, ''), NULLIF(id, ''), lower(email))
                            WHERE user_uuid IS NULL OR user_uuid = ''`);
    } catch (_) {}
    try {
      await sqlite.execute(`UPDATE pm_proxy_users
                            SET active = CASE
                              WHEN active IS NULL THEN COALESCE(is_active, 1)
                              ELSE active
                            END`);
    } catch (_) {}
    try {
      await sqlite.execute(`UPDATE pm_proxy_users
                            SET is_active = CASE
                              WHEN is_active IS NULL THEN COALESCE(active, 1)
                              ELSE is_active
                            END`);
    } catch (_) {}

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS wo_notes_cache (
      wo_uuid     TEXT PRIMARY KEY,
      notes_json  TEXT NOT NULL DEFAULT '[]',
      fetched_at  TEXT NOT NULL DEFAULT (datetime('now'))
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS app_settings (
      key        TEXT PRIMARY KEY,
      value      TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS work_orders_cache (
      id                TEXT    PRIMARY KEY,
      wo_number         TEXT,
      property_id       TEXT,
      property_uuid     TEXT,
      property_name     TEXT,
      unit_id           TEXT,
      unit_number       TEXT,
      address           TEXT,
      description       TEXT,
      status            TEXT,
      priority          TEXT,
      wo_type           TEXT,
      assigned_vendor   TEXT,
      vendor_id         TEXT,
      created_date      TEXT,
      completed_date    TEXT,
      estimated_amount  REAL,
      actual_amount     REAL,
      property_group_id TEXT,
      assigned_to       TEXT,
      is_flagged        INTEGER NOT NULL DEFAULT 0,
      flag_notes        TEXT,
      raw_json          TEXT NOT NULL DEFAULT '{}',
      updated_at        TEXT NOT NULL DEFAULT (datetime('now'))
    )`);
    try {
      await sqlite.execute(
        `ALTER TABLE work_orders_cache ADD COLUMN is_flagged INTEGER NOT NULL DEFAULT 0`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE work_orders_cache ADD COLUMN flag_notes TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE work_orders_cache ADD COLUMN raw_json TEXT NOT NULL DEFAULT '{}'`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE work_orders_cache ADD COLUMN assigned_to TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE work_orders_cache ADD COLUMN property_uuid TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE work_orders_cache ADD COLUMN unit_number TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE work_orders_cache ADD COLUMN wo_type TEXT`,
      );
    } catch (_) {}

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS pm_proxy_login_audit (
      id                  INTEGER PRIMARY KEY AUTOINCREMENT,
      user_uuid           TEXT,
      email               TEXT,
      role                TEXT,
      property_group_uuid TEXT,
      device_token        TEXT,
      created_at          TEXT NOT NULL DEFAULT (datetime('now'))
    )`);

    try {
      await sqlite.execute(
        `ALTER TABLE vendor_overrides ADD COLUMN trade_category TEXT`,
      );
    } catch (_) {}

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS routing_capabilities (
      id         INTEGER PRIMARY KEY AUTOINCREMENT,
      trade      TEXT NOT NULL UNIQUE,
      keywords   TEXT NOT NULL DEFAULT '[]',
      active     INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS routing_pm_group_map (
      group_name TEXT PRIMARY KEY,
      pm_name    TEXT NOT NULL,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS routing_events (
      id             INTEGER PRIMARY KEY AUTOINCREMENT,
      wo_uuid        TEXT NOT NULL UNIQUE,
      wo_number      TEXT,
      property_id    TEXT,
      property_name  TEXT,
      unit_name      TEXT,
      property_group TEXT,
      pm_name        TEXT,
      vendor_id      TEXT,
      vendor_name    TEXT,
      vendor_category TEXT,
      wo_status      TEXT,
      wo_priority    TEXT,
      wo_created_at  TEXT,
      description    TEXT,
      matched_trade  TEXT,
      confidence     TEXT NOT NULL DEFAULT 'medium',
      review_status  TEXT NOT NULL DEFAULT 'pending',
      review_notes   TEXT,
      reviewed_at    TEXT,
      detected_at    TEXT NOT NULL DEFAULT (datetime('now')),
      source         TEXT NOT NULL DEFAULT 'client_scan'
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS closed_turns (
      turn_id    TEXT PRIMARY KEY,
      closed_at  TEXT NOT NULL DEFAULT (datetime('now')),
      close_reason TEXT,
      close_source TEXT,
      closed_by TEXT,
      property_id TEXT,
      property_name TEXT,
      unit_id TEXT,
      unit_name TEXT,
      move_out_date TEXT,
      move_in_date TEXT
    )`);

    // Backward-compatible column adds for older installs.
    try {
      await sqlite.execute(
        `ALTER TABLE closed_turns ADD COLUMN close_reason TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE closed_turns ADD COLUMN close_source TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE closed_turns ADD COLUMN closed_by TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE closed_turns ADD COLUMN property_id TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE closed_turns ADD COLUMN property_name TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(`ALTER TABLE closed_turns ADD COLUMN unit_id TEXT`);
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE closed_turns ADD COLUMN unit_name TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE closed_turns ADD COLUMN move_out_date TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE closed_turns ADD COLUMN move_in_date TEXT`,
      );
    } catch (_) {}

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS unit_turn_tracker (
      tracking_uuid            TEXT PRIMARY KEY,
      tracking_code            TEXT,
      turn_key                 TEXT NOT NULL UNIQUE,
      unit_turn_id             TEXT,
      unit_id                  TEXT,
      property_id              TEXT,
      unit_name                TEXT,
      property_name            TEXT,
      move_out_date            TEXT,
      move_in_date             TEXT,
      inspection_date          TEXT,
      first_wo_date            TEXT,
      estimate_requested_date  TEXT,
      estimate_received_date   TEXT,
      status                   TEXT DEFAULT 'on_radar',
      confidence_score         INTEGER DEFAULT 0,
      confidence_label         TEXT DEFAULT 'low',
      site_manager             TEXT,
      source_flags             TEXT,
      metadata                 TEXT,
      closed_at                TEXT,
      created_at               TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at               TEXT NOT NULL DEFAULT (datetime('now'))
    )`);
    try {
      await sqlite.execute(
        `ALTER TABLE unit_turn_tracker ADD COLUMN estimate_requested_date TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE unit_turn_tracker ADD COLUMN estimate_received_date TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE unit_turn_tracker ADD COLUMN confidence_score INTEGER DEFAULT 0`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE unit_turn_tracker ADD COLUMN confidence_label TEXT DEFAULT 'low'`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE unit_turn_tracker ADD COLUMN site_manager TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE unit_turn_tracker ADD COLUMN source_flags TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE unit_turn_tracker ADD COLUMN closed_at TEXT`,
      );
    } catch (_) {}
    try {
      await sqlite.execute(
        `ALTER TABLE unit_turn_tracker ADD COLUMN tracking_code TEXT`,
      );
    } catch (_) {}

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS unit_turn_milestones (
      id             INTEGER PRIMARY KEY AUTOINCREMENT,
      tracking_uuid  TEXT NOT NULL,
      milestone_key  TEXT NOT NULL,
      milestone_date TEXT,
      source         TEXT,
      notes          TEXT,
      updated_at     TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(tracking_uuid, milestone_key)
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS unit_turn_work_orders (
      id             INTEGER PRIMARY KEY AUTOINCREMENT,
      tracking_uuid  TEXT NOT NULL,
      wo_id          TEXT NOT NULL,
      wo_db_uuid     TEXT,
      source         TEXT DEFAULT 'inferred',
      status         TEXT,
      created_at     TEXT,
      removed        INTEGER DEFAULT 0,
      updated_at     TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(tracking_uuid, wo_id)
    )`);

    // ── Offline routing / resolution matrix ────────────────────────────────
    // Added to support group-scoped routing and fully local PM resolution.
    await sqlite.execute(`PRAGMA foreign_keys = ON`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS property_group_map (
      id                    TEXT PRIMARY KEY,
      name                  TEXT NOT NULL,
      property_ids_json     TEXT,
      cached_at             INTEGER NOT NULL,
      last_membership_sync  INTEGER
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS property_map (
      id                 TEXT PRIMARY KEY,
      property_id        TEXT NOT NULL,
      unit_id            TEXT,
      is_unit            INTEGER NOT NULL DEFAULT 0,
      property_group_id  TEXT,
      property_name      TEXT,
      unit_name          TEXT,
      address            TEXT,
      city               TEXT,
      state              TEXT DEFAULT 'AZ',
      zip                TEXT,
      cached_at          INTEGER NOT NULL,
      last_sync_at       INTEGER,
      FOREIGN KEY(property_group_id) REFERENCES property_group_map(id)
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS vendor_map (
      id                TEXT PRIMARY KEY,
      name              TEXT NOT NULL,
      company_name      TEXT,
      email             TEXT,
      phone             TEXT,
      license_number    TEXT,
      insurance_expiry  TEXT,
      cached_at         INTEGER NOT NULL
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS pm_assignments (
      pm_id              TEXT NOT NULL,
      pm_email           TEXT NOT NULL,
      pm_display_name    TEXT,
      property_group_id  TEXT NOT NULL,
      is_active          INTEGER NOT NULL DEFAULT 1,
      assigned_at        INTEGER NOT NULL,
      last_login_at      INTEGER,
      PRIMARY KEY (pm_id, property_group_id),
      FOREIGN KEY(property_group_id) REFERENCES property_group_map(id)
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS property_group_members (
      property_group_id  TEXT NOT NULL,
      property_map_id    TEXT NOT NULL,
      PRIMARY KEY (property_group_id, property_map_id),
      FOREIGN KEY(property_group_id) REFERENCES property_group_map(id),
      FOREIGN KEY(property_map_id) REFERENCES property_map(id)
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS work_order_map (
      id                   TEXT PRIMARY KEY,
      work_order_number    TEXT UNIQUE NOT NULL,
      property_map_id      TEXT NOT NULL,
      property_id          TEXT,
      unit_id              TEXT,
      vendor_id            TEXT,
      occupancy_id         TEXT,
      status               TEXT,
      priority             TEXT,
      category             TEXT,
      description          TEXT,
      is_turn_wo           INTEGER DEFAULT 0,
      turn_tracking_uuid   TEXT,
      assigned_users_json  TEXT,
      created_date         TEXT,
      completed_date       TEXT,
      last_updated_at      TEXT,
      cached_at            INTEGER NOT NULL,
      FOREIGN KEY(property_map_id) REFERENCES property_map(id),
      FOREIGN KEY(vendor_id) REFERENCES vendor_map(id),
      FOREIGN KEY(turn_tracking_uuid) REFERENCES unit_turn_tracker(tracking_uuid)
    )`);

    // Additive migrations for existing work_order_map installs.
    try { await sqlite.execute(`ALTER TABLE work_order_map ADD COLUMN property_id TEXT`); } catch (_) {}
    try { await sqlite.execute(`ALTER TABLE work_order_map ADD COLUMN unit_id TEXT`); } catch (_) {}

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS billing_map (
      id                       TEXT PRIMARY KEY,
      vendor_id                TEXT,
      property_map_id          TEXT,
      property_id              TEXT,
      work_order_id            TEXT,
      work_order_number        TEXT,
      invoice_date             TEXT,
      due_date                 TEXT,
      total_amount             REAL,
      check_memo               TEXT,
      approval_status          TEXT,
      bill_number              TEXT,
      management_company_payee INTEGER DEFAULT 0,
      line_items_json          TEXT,
      last_updated_at          TEXT,
      cached_at                INTEGER NOT NULL,
      FOREIGN KEY(vendor_id) REFERENCES vendor_map(id),
      FOREIGN KEY(property_map_id) REFERENCES property_map(id),
      FOREIGN KEY(work_order_id) REFERENCES work_order_map(id)
    )`);

    // Additive migrations for existing billing_map installs.
    try { await sqlite.execute(`ALTER TABLE billing_map ADD COLUMN property_id TEXT`); } catch (_) {}
    try { await sqlite.execute(`ALTER TABLE billing_map ADD COLUMN approval_status TEXT`); } catch (_) {}
    try { await sqlite.execute(`ALTER TABLE billing_map ADD COLUMN bill_number TEXT`); } catch (_) {}

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS bill_line_items (
      id             INTEGER PRIMARY KEY AUTOINCREMENT,
      bill_id        TEXT NOT NULL,
      unit_id        TEXT,
      gl_account_id  TEXT,
      amount         REAL,
      description    TEXT,
      line_item_type TEXT,
      quantity       REAL,
      unit_price     REAL,
      cached_at      TEXT NOT NULL DEFAULT (datetime('now'))
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS open_work_orders_view (
      id                   TEXT PRIMARY KEY,
      work_order_number    TEXT NOT NULL,
      property_map_id      TEXT NOT NULL,
      property_name        TEXT,
      unit_name            TEXT,
      address              TEXT,
      property_group_id    TEXT,
      vendor_id            TEXT,
      vendor_name          TEXT,
      status               TEXT,
      priority             TEXT,
      category             TEXT,
      description          TEXT,
      is_turn_wo           INTEGER DEFAULT 0,
      turn_tracking_uuid   TEXT,
      occupancy_id         TEXT,
      assigned_users_json  TEXT,
      created_date         TEXT,
      days_open            INTEGER,
      cached_at            INTEGER NOT NULL,
      FOREIGN KEY(property_map_id) REFERENCES property_map(id),
      FOREIGN KEY(vendor_id) REFERENCES vendor_map(id),
      FOREIGN KEY(property_group_id) REFERENCES property_group_map(id),
      FOREIGN KEY(turn_tracking_uuid) REFERENCES unit_turn_tracker(tracking_uuid)
    )`);

    await sqlite.execute(`CREATE TABLE IF NOT EXISTS group_resolution_cache (
      property_map_id      TEXT NOT NULL,
      property_id          TEXT NOT NULL,
      unit_id              TEXT,
      is_unit              INTEGER DEFAULT 0,
      property_group_id    TEXT NOT NULL,
      property_group_name  TEXT,
      pm_id                TEXT,
      pm_email             TEXT,
      resolved_at          INTEGER NOT NULL,
      PRIMARY KEY (property_map_id, property_group_id),
      FOREIGN KEY(property_map_id) REFERENCES property_map(id),
      FOREIGN KEY(property_group_id) REFERENCES property_group_map(id)
    )`);

    // ── Dispatch / portal indexes ───────────────────────────────────────────
    for (
      const idx of [
        `CREATE INDEX IF NOT EXISTS idx_rq_wo    ON reassignment_queue(wo_id)`,
        `CREATE INDEX IF NOT EXISTS idx_rq_state ON reassignment_queue(warning_sent, auto_exempt)`,
        `CREATE INDEX IF NOT EXISTS idx_ml_tok   ON magic_link_tokens(token)`,
        `CREATE INDEX IF NOT EXISTS idx_ml_wo    ON magic_link_tokens(wo_id)`,
        `CREATE INDEX IF NOT EXISTS idx_ml_short ON magic_link_tokens(short_code)`,
        `CREATE INDEX IF NOT EXISTS idx_ml_sched ON magic_link_tokens(stop_auto, exempt_until)`,
        `CREATE INDEX IF NOT EXISTS idx_sl_token ON short_links(token)`,
        `CREATE INDEX IF NOT EXISTS idx_sl_exp   ON short_links(expires_at)`,
        `CREATE INDEX IF NOT EXISTS idx_pe_tok   ON portal_events(token, created_at DESC)`,
        `CREATE INDEX IF NOT EXISTS idx_pe_wo    ON portal_events(wo_id, created_at DESC)`,
        `CREATE INDEX IF NOT EXISTS idx_mon_wo   ON monitored_work_orders(wo_id)`,
        `CREATE INDEX IF NOT EXISTS idx_mon_date ON monitored_work_orders(created_at DESC)`,
        `CREATE INDEX IF NOT EXISTS idx_tcl_wo   ON tenant_comms_log(wo_id)`,
        `CREATE INDEX IF NOT EXISTS idx_wal_wo   ON wo_audit_log(wo_id)`,
        `CREATE INDEX IF NOT EXISTS idx_wal_evt  ON wo_audit_log(event_type, created_at DESC)`,
        `CREATE INDEX IF NOT EXISTS idx_tg_tier  ON tech_grades(tier, active)`,
        `CREATE INDEX IF NOT EXISTS idx_rte_detected ON routing_events(detected_at DESC)`,
        `CREATE INDEX IF NOT EXISTS idx_rte_pm ON routing_events(pm_name, detected_at DESC)`,
        `CREATE INDEX IF NOT EXISTS idx_rte_status ON routing_events(review_status, detected_at DESC)`,
        `CREATE INDEX IF NOT EXISTS idx_rte_group ON routing_events(property_group)`,
        `CREATE INDEX IF NOT EXISTS idx_rpm_group ON routing_pm_group_map(group_name)`,
        `CREATE INDEX IF NOT EXISTS idx_closed_turns_closed_at ON closed_turns(closed_at DESC)`,
        `CREATE UNIQUE INDEX IF NOT EXISTS uq_pm_proxy_users_id ON pm_proxy_users(id)`,
        `CREATE INDEX IF NOT EXISTS idx_pm_proxy_users_email ON pm_proxy_users(email, active)`,
        `CREATE INDEX IF NOT EXISTS idx_pm_proxy_users_group ON pm_proxy_users(property_group_uuid, active)`,
        `CREATE INDEX IF NOT EXISTS idx_wo_notes_cache_fetched ON wo_notes_cache(fetched_at DESC)`,
        `CREATE INDEX IF NOT EXISTS idx_woc_status ON work_orders_cache(status)`,
        `CREATE INDEX IF NOT EXISTS idx_woc_priority ON work_orders_cache(priority)`,
        `CREATE INDEX IF NOT EXISTS idx_woc_pg ON work_orders_cache(property_group_id)`,
        `CREATE INDEX IF NOT EXISTS idx_woc_created ON work_orders_cache(created_date DESC)`,
        `CREATE INDEX IF NOT EXISTS idx_woc_updated ON work_orders_cache(updated_at DESC)`,
        `CREATE INDEX IF NOT EXISTS idx_woc_vendor ON work_orders_cache(vendor_id)`,
        `CREATE INDEX IF NOT EXISTS idx_pm_proxy_login_audit_user ON pm_proxy_login_audit(user_uuid, created_at DESC)`,
        `CREATE INDEX IF NOT EXISTS idx_utt_turn_key ON unit_turn_tracker(turn_key)`,
        `CREATE INDEX IF NOT EXISTS idx_utt_status ON unit_turn_tracker(status, updated_at DESC)`,
        `CREATE INDEX IF NOT EXISTS idx_utt_unit_prop ON unit_turn_tracker(unit_id, property_id)`,
        `CREATE INDEX IF NOT EXISTS idx_utm_uuid ON unit_turn_milestones(tracking_uuid, milestone_key)`,
        `CREATE INDEX IF NOT EXISTS idx_utwo_uuid ON unit_turn_work_orders(tracking_uuid, removed)`,
        `CREATE INDEX IF NOT EXISTS idx_utwo_wo ON unit_turn_work_orders(wo_id, removed)`,
        `CREATE INDEX IF NOT EXISTS idx_pma_group_id ON pm_assignments(property_group_id)`,
        `CREATE INDEX IF NOT EXISTS idx_pma_email ON pm_assignments(pm_email)`,
        `CREATE INDEX IF NOT EXISTS idx_pma_active ON pm_assignments(is_active)`,
        `CREATE INDEX IF NOT EXISTS idx_pgmap_cached ON property_group_map(cached_at)`,
        `CREATE INDEX IF NOT EXISTS idx_pgm_map_id ON property_group_members(property_map_id)`,
        `CREATE INDEX IF NOT EXISTS idx_pgm_group_id ON property_group_members(property_group_id)`,
        `CREATE INDEX IF NOT EXISTS idx_pm_property_id ON property_map(property_id)`,
        `CREATE INDEX IF NOT EXISTS idx_pm_unit_id ON property_map(unit_id)`,
        `CREATE INDEX IF NOT EXISTS idx_pm_is_unit ON property_map(is_unit)`,
        `CREATE INDEX IF NOT EXISTS idx_pm_group_id ON property_map(property_group_id)`,
        `CREATE INDEX IF NOT EXISTS idx_pm_zip ON property_map(zip)`,
        `CREATE INDEX IF NOT EXISTS idx_vm_name ON vendor_map(name)`,
        `CREATE INDEX IF NOT EXISTS idx_wom_number ON work_order_map(work_order_number)`,
        `CREATE INDEX IF NOT EXISTS idx_wom_property ON work_order_map(property_map_id)`,
        `CREATE INDEX IF NOT EXISTS idx_wom_property_id ON work_order_map(property_id)`,
        `CREATE INDEX IF NOT EXISTS idx_wom_unit_id ON work_order_map(unit_id)`,
        `CREATE INDEX IF NOT EXISTS idx_wom_vendor ON work_order_map(vendor_id)`,
        `CREATE INDEX IF NOT EXISTS idx_wom_status ON work_order_map(status)`,
        `CREATE INDEX IF NOT EXISTS idx_wom_priority ON work_order_map(priority)`,
        `CREATE INDEX IF NOT EXISTS idx_wom_is_turn ON work_order_map(is_turn_wo)`,
        `CREATE INDEX IF NOT EXISTS idx_wom_turn_uuid ON work_order_map(turn_tracking_uuid)`,
        `CREATE INDEX IF NOT EXISTS idx_wom_occupancy ON work_order_map(occupancy_id)`,
        `CREATE INDEX IF NOT EXISTS idx_bm_vendor ON billing_map(vendor_id)`,
        `CREATE INDEX IF NOT EXISTS idx_bm_property ON billing_map(property_map_id)`,
        `CREATE INDEX IF NOT EXISTS idx_bm_property_id ON billing_map(property_id)`,
        `CREATE INDEX IF NOT EXISTS idx_bm_wo_id ON billing_map(work_order_id)`,
        `CREATE INDEX IF NOT EXISTS idx_bm_wo_number ON billing_map(work_order_number)`,
        `CREATE INDEX IF NOT EXISTS idx_bm_invoice ON billing_map(invoice_date)`,
        `CREATE INDEX IF NOT EXISTS idx_bm_due ON billing_map(due_date)`,
        `CREATE INDEX IF NOT EXISTS idx_bm_approval ON billing_map(approval_status)`,
        `CREATE INDEX IF NOT EXISTS idx_bli_bill_id ON bill_line_items(bill_id)`,
        `CREATE INDEX IF NOT EXISTS idx_bli_unit ON bill_line_items(unit_id)`,
        `CREATE INDEX IF NOT EXISTS idx_bli_gl ON bill_line_items(gl_account_id)`,
        `CREATE INDEX IF NOT EXISTS idx_owov_group ON open_work_orders_view(property_group_id)`,
        `CREATE INDEX IF NOT EXISTS idx_owov_vendor ON open_work_orders_view(vendor_id)`,
        `CREATE INDEX IF NOT EXISTS idx_owov_is_turn ON open_work_orders_view(is_turn_wo)`,
        `CREATE INDEX IF NOT EXISTS idx_owov_status ON open_work_orders_view(status)`,
        `CREATE INDEX IF NOT EXISTS idx_owov_priority ON open_work_orders_view(priority)`,
        `CREATE INDEX IF NOT EXISTS idx_owov_turn_uuid ON open_work_orders_view(turn_tracking_uuid)`,
        `CREATE INDEX IF NOT EXISTS idx_grc_property_id ON group_resolution_cache(property_id)`,
        `CREATE INDEX IF NOT EXISTS idx_grc_group ON group_resolution_cache(property_group_id)`,
        `CREATE INDEX IF NOT EXISTS idx_grc_pm_id ON group_resolution_cache(pm_id)`,
        `CREATE INDEX IF NOT EXISTS idx_grc_pm_email ON group_resolution_cache(pm_email)`,
      ]
    ) {
      try {
        await sqlite.execute(idx);
      } catch (_) {}
    }

    // ── Seed proxy_config defaults (INSERT OR IGNORE — fully idempotent) ─────
    for (
      const [key, value] of [
        ["warn_threshold_hours", "36"],
        ["reassign_threshold_hours", "48"],
        ["go_back_tolerance_pct", "2"],
        ["tier2_claim_window_hours", "24"],
        ["grace_period_enabled", "1"],
        ["max_reassigns_before_escalate", "2"],
        ["dispatch_paused", "0"],
        ["dispatch_tier1_group_uuid", "efe085ca-229e-11ef-bfba-069ca18f5865"],
        ["dispatch_tier2_group_uuid", "a3db4460-22b3-11ef-bfba-069ca18f5865"],
        ["dispatch_active_branch", "all"],
        ["dispatch_hidden_assignees", "{}"],
        ["dispatch_auto_sync_assignees", "1"],
        ["dispatch_auto_sync_cooldown_sec", "120"],
        ["brand_name", "Fort Lowell Realty"],
        ["brand_logo_url", ""],
        ["portal_brand_name", "Fort Lowell Realty Tech Dispatch"],
        [
          "portal_brand_logo_url",
          "https://pfst.cf2.poecdn.net/base/image/57c851c04753092259d83d0a1aa34e2fd889c7218b50a338e6100dbf21ae922c?w=733&h=982",
        ],
        ["app_version", PROXY_APP_VERSION],
        [
          "tier1_open_statuses",
          "assigned,scheduled,waiting,in progress,new,estimated",
        ],
        // OTP / auth policy (editable via Advanced Manager Settings panel)
        ["otp_enabled", "1"],
        ["otp_allowed_domain", ""],
        ["otp_require_pm_membership", "1"],
        ["otp_ttl_minutes", "10"],
      ]
    ) {
      try {
        await sqlite.execute({
          sql: `INSERT OR IGNORE INTO proxy_config (key, value) VALUES (?, ?)`,
          args: [key, value],
        });
      } catch (_) {}
    }

    // Cold-start housekeeping
    try {
      await webhookCleanup();
    } catch (_) {}

    _tablesReady = true;
    console.log("ensureTables: all tables ready ✓");
  } catch (err: any) {
    // Do NOT mark ready — allow retry on next request so tables get created
    // once the DB recovers. CREATE TABLE IF NOT EXISTS is idempotent and fast.
    const message = String(
      err?.message || err || "unknown ensureTables failure",
    );
    console.log(`ensureTables ERROR: ${message.substring(0, 200)}`);
    throw new Error(`ensureTables failed: ${message}`);
  }
}

// ── Cache Helpers ───────────────────────────────────────────────────────────

export async function cacheGet(cacheKey: string, entityType: string) {
  try {
    const ttl = CACHE_TTL[entityType] || 15;
    const result = await sqlite.execute({
      sql: `SELECT data, cached_at, record_count
             FROM api_cache
             WHERE cache_key = ?
               AND datetime(cached_at, '+' || ? || ' minutes') > datetime('now')`,
      args: [cacheKey, ttl],
    });
    const rows = rowsAsObjects(result);
    if (rows.length === 0) return null;
    const row = rows[0];
    const meta = JSON.parse(row.data);

    // Chunked entry — meta row stores { _chunks: N }
    if (meta && typeof meta._chunks === "number") {
      let fullJson = "";
      for (let i = 0; i < meta._chunks; i++) {
        const chunkResult = await sqlite.execute({
          sql: `SELECT data FROM api_cache WHERE cache_key = ?`,
          args: [`${cacheKey}::${i}`],
        });
        const chunkRows = rowsAsObjects(chunkResult);
        if (chunkRows.length === 0) {
          console.log(`Cache GET: chunk ${i} missing for ${cacheKey}`);
          return null; // incomplete — treat as miss
        }
        fullJson += chunkRows[0].data;
      }
      return {
        data: JSON.parse(fullJson),
        cached_at: row.cached_at,
        record_count: row.record_count,
        from_cache: true,
      };
    }

    // Non-chunked (small payload) — return directly
    return {
      data: meta,
      cached_at: row.cached_at,
      record_count: row.record_count,
      from_cache: true,
    };
  } catch (err: any) {
    console.log(
      `Cache GET error for ${entityType}: ${err.message?.substring(0, 120)}`,
    );
    return null;
  }
}

export async function cacheSet(
  cacheKey: string,
  entityType: string,
  data: unknown,
  recordCount = 0,
) {
  try {
    const jsonStr = JSON.stringify(data);
    const sizeKB = (jsonStr.length / 1024).toFixed(0);
    const ttl = CACHE_TTL[entityType] || 15;
    const expiresAt = new Date(Date.now() + ttl * 60_000).toISOString();

    await evictCacheIfNeeded(jsonStr.length);

    // Delete stale chunks + meta row before writing
    await sqlite.execute({
      sql: `DELETE FROM api_cache WHERE cache_key LIKE ?`,
      args: [`${cacheKey}::%`],
    });
    await sqlite.execute({
      sql: `DELETE FROM api_cache WHERE cache_key = ?`,
      args: [cacheKey],
    });

    if (jsonStr.length <= CHUNK_LIMIT) {
      await sqlite.execute({
        sql: `INSERT INTO api_cache
                 (cache_key, entity_type, data, cached_at, expires_at, record_count)
               VALUES (?, ?, ?, datetime('now'), ?, ?)`,
        args: [cacheKey, entityType, jsonStr, expiresAt, recordCount],
      });
      console.log(
        `Cache SET: ${entityType} (${sizeKB} KB, ${recordCount} records, TTL ${ttl}m)`,
      );
    } else {
      // Split JSON into CHUNK_LIMIT-sized pieces
      const numChunks = Math.ceil(jsonStr.length / CHUNK_LIMIT);
      for (let i = 0; i < numChunks; i++) {
        const chunk = jsonStr.substring(i * CHUNK_LIMIT, (i + 1) * CHUNK_LIMIT);
        await sqlite.execute({
          sql: `INSERT INTO api_cache
                   (cache_key, entity_type, data, cached_at, expires_at, record_count)
                 VALUES (?, ?, ?, datetime('now'), ?, ?)`,
          args: [`${cacheKey}::${i}`, entityType, chunk, expiresAt, 0],
        });
      }
      // Meta row — tells cacheGet how many chunks to reassemble
      await sqlite.execute({
        sql: `INSERT INTO api_cache
                 (cache_key, entity_type, data, cached_at, expires_at, record_count)
               VALUES (?, ?, ?, datetime('now'), ?, ?)`,
        args: [
          cacheKey,
          entityType,
          JSON.stringify({ _chunks: numChunks }),
          expiresAt,
          recordCount,
        ],
      });
      console.log(
        `Cache SET (chunked): ${entityType} (${sizeKB} KB → ${numChunks} chunks, ${recordCount} records, TTL ${ttl}m)`,
      );
    }
  } catch (err: any) {
    // Non-fatal — API response still returned even if cache write fails
    console.log(
      `Cache WRITE FAILED for ${entityType}: ${err.message?.substring(0, 150)}`,
    );
  }
}

export async function cacheInvalidate(entityType: string): Promise<void> {
  await sqlite.execute({
    sql: `DELETE FROM api_cache WHERE entity_type = ?`,
    args: [entityType],
  });
}

// ── Storage Budget Helpers ────────────────────────────────────────────────────

export async function getCacheSizeBytes(): Promise<number> {
  try {
    const res = await sqlite.execute(
      `SELECT COALESCE(SUM(LENGTH(data)), 0) AS total FROM api_cache`,
    );
    return Number(rowsAsObjects(res)[0]?.total || 0);
  } catch {
    return 0;
  }
}

export async function getWebhookSizeBytes(): Promise<number> {
  try {
    const res = await sqlite.execute(
      `SELECT COALESCE(SUM(LENGTH(raw_body)), 0) AS total FROM webhook_events`,
    );
    return Number(rowsAsObjects(res)[0]?.total || 0);
  } catch {
    return 0;
  }
}

export async function evictCacheIfNeeded(incomingBytes: number): Promise<void> {
  // 1. Purge expired entries first (free space without losing live data)
  try {
    await sqlite.execute(
      `DELETE FROM api_cache WHERE datetime(expires_at) < datetime('now')`,
    );
  } catch {}

  let used = await getCacheSizeBytes();
  if (used + incomingBytes <= STORAGE_BUDGET_BYTES) return;

  // 2. Evict oldest entity types one at a time until space is available
  try {
    const entityTypes = await sqlite.execute(
      `SELECT entity_type, MIN(cached_at) AS oldest
       FROM api_cache
       GROUP BY entity_type
       ORDER BY oldest ASC`,
    );
    for (const et of rowsAsObjects(entityTypes)) {
      if (used + incomingBytes <= STORAGE_BUDGET_BYTES) break;
      await sqlite.execute({
        sql: `DELETE FROM api_cache WHERE entity_type = ?`,
        args: [et.entity_type],
      });
      console.log(`Evicted cache: ${et.entity_type} to free space`);
      used = await getCacheSizeBytes();
    }
  } catch (e: any) {
    console.log(`evictCacheIfNeeded error: ${e.message?.substring(0, 100)}`);
  }
}

export async function webhookCleanup(): Promise<{
  deleted_old: number;
  deleted_overflow: number;
}> {
  let deletedOld = 0;
  let deletedOverflow = 0;

  try {
    const cutoff = new Date(Date.now() - WEBHOOK_MAX_DAYS * 86400_000)
      .toISOString();
    const r1 = await sqlite.execute({
      sql: `DELETE FROM webhook_events WHERE received_at < ?`,
      args: [cutoff],
    });
    deletedOld = r1.rowsAffected || 0;

    const countRes = await sqlite.execute(
      `SELECT COUNT(*) AS n FROM webhook_events`,
    );
    const count = Number(rowsAsObjects(countRes)[0]?.n || 0);
    if (count > WEBHOOK_MAX_EVENTS) {
      const excess = count - WEBHOOK_MAX_EVENTS;
      await sqlite.execute({
        sql: `DELETE FROM webhook_events WHERE id IN (
                 SELECT id FROM webhook_events ORDER BY id ASC LIMIT ?
               )`,
        args: [excess],
      });
      deletedOverflow = excess;
    }
  } catch (e: any) {
    console.log(`webhookCleanup error: ${e.message?.substring(0, 100)}`);
  }

  return { deleted_old: deletedOld, deleted_overflow: deletedOverflow };
}

// ── Relational upsert helpers ─────────────────────────────────────────────────
// Side-write helpers called from handlers after a fresh API fetch.
// Accept raw AppFolio rows (v0 PascalCase or v2 snake_case), normalise, and
// batch-upsert into the relational map tables using multi-row INSERT statements
// (one round-trip per 50 rows).  All errors are caught — callers are never
// affected.  Tables are populated lazily: each entity family must be upserted
// before dependents (properties → vendors → work_orders → bills) for FK
// consistency, but partial failures are silently retried on the next sync.

const _UPSERT_BATCH = 50;

function _s(v: unknown, max = 500): string | null {
  const s = String(v ?? "").trim();
  return s ? s.substring(0, max) : null;
}

function _r(v: unknown): number | null {
  const n = parseFloat(String(v ?? "").replace(/[^0-9.-]/g, ""));
  return isFinite(n) ? n : null;
}

// Upsert raw property rows from DB API v0 /properties.
// Sets property_map.id = AF property UUID so work_order_map.property_map_id
// can use the same value as a FK reference.
export async function upsertPropertyRows(rows: any[]): Promise<void> {
  const now = Date.now();
  for (let i = 0; i < rows.length; i += _UPSERT_BATCH) {
    const chunk = rows.slice(i, i + _UPSERT_BATCH);
    const ph: string[] = [];
    const vals: unknown[] = [];
    for (const row of chunk) {
      const id = _s(row.Id || row.id || row.property_id);
      if (!id) continue;
      const pgIds: string[] = Array.isArray(row.PropertyGroupIds || row.property_group_ids)
        ? (row.PropertyGroupIds || row.property_group_ids)
            .map((v: unknown) => String(v ?? "").trim())
            .filter(Boolean)
        : [];
      ph.push("(?,?,?,?,?,?,?,?,?,?)");
      vals.push(
        id,                                                         // id = AF property UUID
        id,                                                         // property_id
        null,                                                       // unit_id
        0,                                                          // is_unit
        pgIds[0] ?? null,                                           // property_group_id
        _s(row.Name || row.PropertyName || row.property_name),     // property_name
        null,                                                       // unit_name
        _s(row.Address1 || row.StreetAddress || row.address),      // address
        _s(row.City || row.city),                                   // city
        now,                                                        // cached_at
      );
    }
    if (ph.length === 0) continue;
    try {
      await sqlite.execute({
        sql: `INSERT OR REPLACE INTO property_map
              (id, property_id, unit_id, is_unit, property_group_id,
               property_name, unit_name, address, city, cached_at)
              VALUES ${ph.join(",")}`,
        args: vals,
      });
    } catch (_) {}
  }
}

// Upsert raw vendor rows from Reports API v2 vendor_directory.
export async function upsertVendorRows(rows: any[]): Promise<void> {
  const now = Date.now();
  for (let i = 0; i < rows.length; i += _UPSERT_BATCH) {
    const chunk = rows.slice(i, i + _UPSERT_BATCH);
    const ph: string[] = [];
    const vals: unknown[] = [];
    for (const row of chunk) {
      const id = _s(row.vendor_id || row.VendorId || row.Id || row.id);
      if (!id) continue;
      const name =
        _s(
          row.company_name ||
            row.CompanyName ||
            [row.first_name || row.FirstName, row.last_name || row.LastName]
              .filter(Boolean)
              .join(" ") ||
            row.name ||
            row.Name,
        ) ?? id;
      ph.push("(?,?,?,?,?,?,?,?)");
      vals.push(
        id,
        name,
        _s(row.company_name || row.CompanyName),
        _s(row.email || row.Email),
        _s(row.phone || row.Phone || row.phone_numbers),
        _s(row.license_number || row.LicenseNumber),
        _s(row.insurance_expiry || row.InsuranceExpiry),
        now,
      );
    }
    if (ph.length === 0) continue;
    try {
      await sqlite.execute({
        sql: `INSERT OR REPLACE INTO vendor_map
              (id, name, company_name, email, phone, license_number, insurance_expiry, cached_at)
              VALUES ${ph.join(",")}`,
        args: vals,
      });
    } catch (_) {}
  }
}

// Upsert raw work order rows from Reports API v2 work_order report.
// property_map_id is set to the AF property UUID (consistent with
// upsertPropertyRows which keys property_map.id by AF UUID).
// If property_map hasn't been populated yet the INSERT fails the FK check
// and is silently skipped; it succeeds on the next sync after properties load.
export async function upsertWorkOrderRows(rows: any[]): Promise<void> {
  const now = Date.now();
  for (let i = 0; i < rows.length; i += _UPSERT_BATCH) {
    const chunk = rows.slice(i, i + _UPSERT_BATCH);
    const ph: string[] = [];
    const vals: unknown[] = [];
    for (const row of chunk) {
      const id = _s(row.work_order_id || row.WorkOrderId || row.Id || row.id);
      if (!id) continue;
      const woNumber = _s(row.work_order_number || row.WorkOrderNumber || row.Number);
      if (!woNumber) continue;
      const propertyId = _s(row.property_id || row.PropertyId) ?? "";
      // property_map_id = AF property UUID (property_map.id keyed by AF UUID)
      const propertyMapId = propertyId || id;
      const assignedRaw = row.assigned_users ?? row.AssignedUsers ?? [];
      ph.push("(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
      vals.push(
        id,
        woNumber,
        propertyMapId,                                                        // property_map_id (FK)
        propertyId || null,                                                   // property_id (denorm)
        _s(row.unit_id || row.UnitId),                                        // unit_id
        _s(row.vendor_id || row.VendorId),                                    // vendor_id
        _s(row.status || row.Status),
        _s(row.priority || row.Priority),
        _s(row.category || row.Category || row.work_order_type || row.WorkOrderType),
        _s(row.description || row.Description || row.subject || row.Subject, 500),
        typeof assignedRaw === "string" ? assignedRaw : JSON.stringify(assignedRaw),
        _s(row.created_date || row.CreatedDate || row.created_at || row.CreatedAt)?.slice(0, 10) ?? null,
        _s(row.completed_date || row.CompletedDate || row.closed_date)?.slice(0, 10) ?? null,
        _s(row.last_updated_at || row.LastUpdatedAt)?.slice(0, 24) ?? null,
        now,
      );
    }
    if (ph.length === 0) continue;
    try {
      await sqlite.execute({
        sql: `INSERT OR REPLACE INTO work_order_map
              (id, work_order_number, property_map_id, property_id, unit_id, vendor_id,
               status, priority, category, description, assigned_users_json,
               created_date, completed_date, last_updated_at, cached_at)
              VALUES ${ph.join(",")}`,
        args: vals,
      });
    } catch (_) {}
  }
}

// Upsert raw bill rows from DB API v0 /bills.
// Also expands LineItems into bill_line_items (delete-then-insert per bill).
// property_map_id is kept null to avoid FK violations when property_map
// hasn't been populated yet; property_id (denormalized) is always written.
export async function upsertBillingRows(rows: any[]): Promise<void> {
  const now = Date.now();

  // First pass: upsert billing_map rows
  for (let i = 0; i < rows.length; i += _UPSERT_BATCH) {
    const chunk = rows.slice(i, i + _UPSERT_BATCH);
    const ph: string[] = [];
    const vals: unknown[] = [];
    for (const row of chunk) {
      const id = _s(row.Id || row.id || row.BillId || row.bill_id);
      if (!id) continue;
      const propertyId = _s(row.PropertyId || row.property_id);
      const lineItems = Array.isArray(row.LineItems)
        ? row.LineItems
        : (Array.isArray(row.line_items) ? row.line_items : []);
      ph.push("(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
      vals.push(
        id,
        _s(row.VendorId || row.vendor_id || row.PayeeId || row.payee_id),     // vendor_id
        null,                                                                   // property_map_id (null avoids FK miss)
        propertyId,                                                             // property_id (denorm)
        _s(row.WorkOrderId || row.work_order_id),                              // work_order_id
        _s(row.WorkOrderNumber || row.work_order_number),                      // work_order_number
        _s(row.InvoiceDate || row.invoice_date)?.slice(0, 10) ?? null,
        _s(row.DueDate || row.due_date)?.slice(0, 10) ?? null,
        _r(row.TotalAmount || row.total_amount || row.Amount || row.amount),
        _s(row.CheckMemo || row.check_memo || row.Remarks || row.remarks, 500),
        _s(row.ApprovalStatus || row.approval_status || row.Status || row.status), // approval_status
        _s(row.Reference || row.reference || id),                              // bill_number
        lineItems.length > 0 ? JSON.stringify(lineItems) : null,               // line_items_json
        _s(row.LastUpdatedAt || row.last_updated_at)?.slice(0, 24) ?? null,
        now,
      );
    }
    if (ph.length === 0) continue;
    try {
      await sqlite.execute({
        sql: `INSERT OR REPLACE INTO billing_map
              (id, vendor_id, property_map_id, property_id, work_order_id, work_order_number,
               invoice_date, due_date, total_amount, check_memo, approval_status, bill_number,
               line_items_json, last_updated_at, cached_at)
              VALUES ${ph.join(",")}`,
        args: vals,
      });
    } catch (_) {}
  }

  // Second pass: expand line items into bill_line_items
  for (const row of rows) {
    const billId = _s(row.Id || row.id || row.BillId || row.bill_id);
    if (!billId) continue;
    const lineItems = Array.isArray(row.LineItems)
      ? row.LineItems
      : (Array.isArray(row.line_items) ? row.line_items : []);
    if (lineItems.length === 0) continue;
    try {
      await sqlite.execute({ sql: `DELETE FROM bill_line_items WHERE bill_id = ?`, args: [billId] });
      for (let i = 0; i < lineItems.length; i += _UPSERT_BATCH) {
        const chunk = lineItems.slice(i, i + _UPSERT_BATCH);
        const ph: string[] = [];
        const vals: unknown[] = [];
        for (const li of chunk) {
          ph.push("(?,?,?,?,?,?,?,?)");
          vals.push(
            billId,
            _s(li.UnitId || li.unit_id || li.UnitUuid || li.unit_uuid),
            _s(li.GlAccountId || li.gl_account_id || li.GlAccount || li.gl_account),
            _r(li.Amount || li.amount),
            _s(li.Description || li.description, 500),
            _s(li.LineItemType || li.line_item_type || li.Type || li.type),
            _r(li.Quantity || li.quantity),
            _r(li.UnitPrice || li.unit_price || li.Rate || li.rate),
          );
        }
        await sqlite.execute({
          sql: `INSERT INTO bill_line_items
                (bill_id, unit_id, gl_account_id, amount, description,
                 line_item_type, quantity, unit_price)
                VALUES ${ph.join(",")}`,
          args: vals,
        });
      }
    } catch (_) {}
  }
}