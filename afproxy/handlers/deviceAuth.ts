import { rowsAsObjects, sqlite, sqliteAuth } from "../db.ts";
import { PROXY_ADMIN_KEY } from "../config.ts";
import { checkRateLimit } from "../lib/rateLimit.ts";
import { sendSMS } from "../lib/ringcentral.ts";

let _trustedDevicesTableReady = false;
let _pmProxyUsersTableReady = false;
let _deviceOtpsTableReady = false;

// UUID generator (fallback if crypto.randomUUID is unavailable)
function generateUuid(): string {
  if (globalThis.crypto && globalThis.crypto.randomUUID) {
    return globalThis.crypto.randomUUID();
  }
  // Fallback: generate a v4-like UUID using random values
  const bytes = new Uint8Array(16);
  if (globalThis.crypto && globalThis.crypto.getRandomValues) {
    globalThis.crypto.getRandomValues(bytes);
  } else {
    // Last resort: throw — Deno always has crypto so this is unreachable
    throw new Error("Crypto API unavailable — cannot generate secure UUIDs");
  }
  bytes[6] = (bytes[6] & 0x0f) | 0x40; // version 4
  bytes[8] = (bytes[8] & 0x3f) | 0x80; // variant
  const hex = Array.from(bytes).map((b) => b.toString(16).padStart(2, "0"))
    .join("");
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${
    hex.slice(16, 20)
  }-${hex.slice(20)}`;
}

async function ensureTrustedDevicesTable(): Promise<void> {
  if (_trustedDevicesTableReady) return;
  await sqliteAuth.execute(`CREATE TABLE IF NOT EXISTS trusted_devices (
    device_token TEXT PRIMARY KEY,
    user_name    TEXT,
    created_at   TEXT NOT NULL DEFAULT (datetime('now'))
  )`);
  try {
    await sqliteAuth.execute(
      `ALTER TABLE trusted_devices ADD COLUMN role TEXT DEFAULT 'full'`,
    );
  } catch (_) {}
  try {
    await sqliteAuth.execute(
      `ALTER TABLE trusted_devices ADD COLUMN login_email TEXT`,
    );
  } catch (_) {}
  try {
    await sqliteAuth.execute(
      `ALTER TABLE trusted_devices ADD COLUMN property_group_uuid TEXT`,
    );
  } catch (_) {}
  try {
    await sqliteAuth.execute(
      `ALTER TABLE trusted_devices ADD COLUMN phone TEXT`,
    );
  } catch (_) {}
  try {
    await sqliteAuth.execute(
      `ALTER TABLE trusted_devices ADD COLUMN last_seen_at TEXT`,
    );
  } catch (_) {}
  try {
    await sqliteAuth.execute(
      `ALTER TABLE trusted_devices ADD COLUMN expires_at TEXT`,
    );
  } catch (_) {}
  try {
    await sqliteAuth.execute(
      `ALTER TABLE trusted_devices ADD COLUMN revoked INTEGER DEFAULT 0`,
    );
  } catch (_) {}
  try {
    await sqliteAuth.execute(
      `ALTER TABLE trusted_devices ADD COLUMN auth_source TEXT`,
    );
  } catch (_) {}
  _trustedDevicesTableReady = true;
}

async function ensurePmProxyUsersTable(): Promise<void> {
  if (_pmProxyUsersTableReady) return;
  await sqlite.execute(`CREATE TABLE IF NOT EXISTS pm_proxy_users (
    user_uuid           TEXT PRIMARY KEY,
    email               TEXT UNIQUE NOT NULL,
    full_name           TEXT,
    phone               TEXT,
    property_group_uuid TEXT NOT NULL,
    active              INTEGER DEFAULT 1,
    created_at          TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at          TEXT NOT NULL DEFAULT (datetime('now'))
  )`);
  try {
    await sqlite.execute(
      `ALTER TABLE pm_proxy_users ADD COLUMN created_at TEXT DEFAULT (datetime('now'))`,
    );
  } catch (_) {}
  try {
    await sqlite.execute(
      `ALTER TABLE pm_proxy_users ADD COLUMN updated_at TEXT DEFAULT (datetime('now'))`,
    );
  } catch (_) {}
  _pmProxyUsersTableReady = true;
}

async function ensureDeviceOtpsTable(): Promise<void> {
  if (_deviceOtpsTableReady) return;
  await sqlite.execute(`CREATE TABLE IF NOT EXISTS device_otps (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    email               TEXT NOT NULL,
    code                TEXT NOT NULL,
    used                INTEGER DEFAULT 0,
    expires_at          TEXT NOT NULL,
    user_name           TEXT,
    created_at          TEXT NOT NULL DEFAULT (datetime('now')),
    used_at             TEXT,
    role_hint           TEXT,
    property_group_uuid TEXT
  )`);
  try {
    await sqlite.execute(`ALTER TABLE device_otps ADD COLUMN user_name TEXT`);
  } catch (_) {}
  try {
    await sqlite.execute(`ALTER TABLE device_otps ADD COLUMN created_at TEXT`);
  } catch (_) {}
  try {
    await sqlite.execute(`ALTER TABLE device_otps ADD COLUMN used_at TEXT`);
  } catch (_) {}
  try {
    await sqlite.execute(`ALTER TABLE device_otps ADD COLUMN role_hint TEXT`);
  } catch (_) {}
  try {
    await sqlite.execute(
      `ALTER TABLE device_otps ADD COLUMN property_group_uuid TEXT`,
    );
  } catch (_) {}
  _deviceOtpsTableReady = true;
}

const DEVICE_SETUP_PIN = Deno.env.get("DEVICE_SETUP_PIN") || "";
// Env-var fallbacks — these are overridden by proxy_config DB values when set.
const OTP_ALLOWED_DOMAIN_ENV =
  (Deno.env.get("OTP_ALLOWED_DOMAIN") || "flraz.com")
    .replace(/^@/, "")
    .toLowerCase();
const OTP_TTL_MINUTES_ENV = Math.max(
  3,
  Number(Deno.env.get("DEVICE_OTP_TTL_MINUTES") || "10") || 10,
);

/** Read a single row from proxy_config, returning the string value or a default. */
async function getProxyConfig(key: string, fallback: string): Promise<string> {
  try {
    const res = await sqlite.execute({
      sql: `SELECT value FROM proxy_config WHERE key = ? LIMIT 1`,
      args: [key],
    });
    const rows = rowsAsObjects(res);
    if (rows.length && rows[0].value !== undefined && rows[0].value !== null) {
      return String(rows[0].value);
    }
  } catch (_) {}
  return fallback;
}

/** Read all OTP policy settings in one query. */
async function getOtpPolicy(): Promise<{
  enabled: boolean;
  allowedDomain: string;
  requireMembership: boolean;
  ttlMinutes: number;
}> {
  try {
    const res = await sqlite.execute({
      sql: `SELECT key, value FROM proxy_config
            WHERE key IN ('otp_enabled','otp_allowed_domain','otp_require_pm_membership','otp_ttl_minutes')`,
    });
    const rows = rowsAsObjects(res);
    const map: Record<string, string> = {};
    for (const r of rows) map[String(r.key)] = String(r.value);
    const rawDomain = (map["otp_allowed_domain"] || "").replace(/^@/, "")
      .toLowerCase();
    return {
      enabled: (map["otp_enabled"] ?? "1") !== "0",
      allowedDomain: rawDomain || OTP_ALLOWED_DOMAIN_ENV,
      requireMembership: (map["otp_require_pm_membership"] ?? "1") !== "0",
      ttlMinutes: Math.max(
        3,
        Number(map["otp_ttl_minutes"]) || OTP_TTL_MINUTES_ENV,
      ),
    };
  } catch (_) {
    return {
      enabled: true,
      allowedDomain: OTP_ALLOWED_DOMAIN_ENV,
      requireMembership: true,
      ttlMinutes: OTP_TTL_MINUTES_ENV,
    };
  }
}

async function sendOtpSms(
  toPhone: string,
  code: string,
  ttlMinutes: number,
): Promise<void> {
  const message =
    `HandyManager verification code: ${code}. Expires in ${ttlMinutes} minutes. ` +
    `If you did not request this code, ignore this message.`;
  const sms = await sendSMS(toPhone, message);
  if (!sms.ok) {
    throw new Error(sms.error || "RingCentral SMS send failed");
  }
}

function normalizeOrgEmail(rawEmail: string, allowedDomain: string): string {
  const email = String(rawEmail || "").trim().toLowerCase();
  if (!email || email.indexOf("@") === -1) return "";
  const parts = email.split("@");
  if (parts.length !== 2 || !parts[0] || !parts[1]) return "";
  if (allowedDomain && parts[1] !== allowedDomain) return "";
  return email;
}

function deriveOtpIdentityEmail(
  pmUser: PmProxyUser | null,
  requestedIdentifier: string,
  allowedDomain: string,
): string {
  const direct = normalizeOrgEmail(
    pmUser?.email || requestedIdentifier,
    allowedDomain,
  );
  if (direct) return direct;
  if (!pmUser) return "";
  const domain = String(allowedDomain || "").trim().toLowerCase();
  if (!domain) return "";
  const userKey = String(pmUser.user_uuid || pmUser.email || "").toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
  if (!userKey) return "";
  return `pm-${userKey}@${domain}`;
}

function generateOtpCode(): string {
  const arr = new Uint32Array(1);
  crypto.getRandomValues(arr);
  const value = arr[0] % 1000000;
  return value.toString().padStart(6, "0");
}

function normalizeEmail(rawEmail: string): string {
  return String(rawEmail || "").trim().toLowerCase();
}

function normalizePhone(rawPhone: string): string {
  const val = String(rawPhone || "").trim();
  if (!val) return "";
  const digits = val.replace(/\D+/g, "");
  if (digits.length === 11 && digits.startsWith("1")) return "+" + digits;
  if (digits.length === 10) return "+1" + digits;
  if (val.startsWith("+") && digits.length >= 10 && digits.length <= 15) {
    return "+" + digits;
  }
  return "";
}

type PmProxyUser = {
  user_uuid: string;
  email: string;
  full_name: string;
  phone: string;
  property_group_uuid: string;
  active: number;
};

async function getPmProxyUserByEmail(
  email: string,
): Promise<PmProxyUser | null> {
  const normalized = normalizeEmail(email);
  if (!normalized) return null;
  try {
    const result = await sqlite.execute({
      sql: `SELECT user_uuid, email, full_name, phone, property_group_uuid, active
            FROM pm_proxy_users
            WHERE lower(email) = lower(?) AND active = 1
            LIMIT 1`,
      args: [normalized],
    });
    const rows = rowsAsObjects(result || { rows: [], columns: [] });
    return rows.length ? (rows[0] as PmProxyUser) : null;
  } catch (_) {
    return null;
  }
}

async function getPmUserAccountsByEmail(
  email: string,
): Promise<PmProxyUser | null> {
  const normalized = normalizeEmail(email);
  if (!normalized) return null;
  try {
    const result = await sqlite.execute({
      sql: `SELECT user_uuid, email, full_name, phone, property_group_uuid, active
            FROM pm_user_accounts
            WHERE lower(email) = lower(?) AND coalesce(active,1) = 1
            LIMIT 1`,
      args: [normalized],
    });
    const rows = rowsAsObjects(result || { rows: [], columns: [] });
    return rows.length ? (rows[0] as PmProxyUser) : null;
  } catch (_) {
    return null;
  }
}

async function getPmProxyUserByPhone(
  phone: string,
): Promise<PmProxyUser | null> {
  const normalizedPhone = normalizePhone(phone);
  if (!normalizedPhone) return null;
  try {
    const result = await sqlite.execute({
      sql: `SELECT user_uuid, email, full_name, phone, property_group_uuid, active
            FROM pm_proxy_users
            WHERE active = 1`,
    });
    const rows = rowsAsObjects(result || { rows: [], columns: [] }) as PmProxyUser[];
    for (const row of rows) {
      if (normalizePhone(String(row.phone || "")) === normalizedPhone) return row;
    }
  } catch (_) {
    return null;
  }
  return null;
}

async function getPmUserAccountsByPhone(
  phone: string,
): Promise<PmProxyUser | null> {
  const normalizedPhone = normalizePhone(phone);
  if (!normalizedPhone) return null;
  try {
    const result = await sqlite.execute({
      sql: `SELECT user_uuid, email, full_name, phone, property_group_uuid, active
            FROM pm_user_accounts
            WHERE coalesce(active,1) = 1`,
    });
    const rows = rowsAsObjects(result || { rows: [], columns: [] }) as PmProxyUser[];
    for (const row of rows) {
      if (normalizePhone(String(row.phone || "")) === normalizedPhone) return row;
    }
  } catch (_) {
    return null;
  }
  return null;
}

async function resolvePmProxyUser(
  identifier: string,
): Promise<PmProxyUser | null> {
  const byEmail = await getPmProxyUserByEmail(identifier) ||
    await getPmUserAccountsByEmail(identifier);
  if (byEmail) return byEmail;
  return await getPmProxyUserByPhone(identifier) ||
    await getPmUserAccountsByPhone(identifier);
}

async function insertDeviceOtp(
  email: string,
  code: string,
  ttlMinutes: number,
  userName: string,
  scopeUuid: string,
): Promise<void> {
  try {
    await sqlite.execute({
      sql:
        `INSERT INTO device_otps (email, code, used, expires_at, user_name, role_hint, property_group_uuid)
          VALUES (?, ?, 0, datetime('now', ?), ?, ?, ?)`,
      args: [
        email,
        code,
        `+${ttlMinutes} minutes`,
        userName,
        "pm_readonly",
        scopeUuid,
      ],
    });
    return;
  } catch (err: any) {
    const msg = String(err?.message || err || "").toLowerCase();
    if (
      msg.includes("no such column") ||
      msg.includes("has no column named")
    ) {
      await sqlite.execute({
        sql:
          `INSERT INTO device_otps (email, code, used, expires_at, user_name)
            VALUES (?, ?, 0, datetime('now', ?), ?)`,
        args: [email, code, `+${ttlMinutes} minutes`, userName],
      });
      return;
    }
    throw err;
  }
}

function getBodyField(body: any, ...keys: string[]): string {
  if (!body || typeof body !== "object") return "";
  for (const k of keys) {
    const v = body[k];
    if (v !== undefined && v !== null && String(v).trim()) {
      return String(v).trim();
    }
  }
  return "";
}

export async function handleDeviceSetup(req: Request): Promise<any> {
  await ensureTrustedDevicesTable();
  if (!DEVICE_SETUP_PIN) {
    return {
      ok: false,
      status: 500,
      error: "Trusted device setup disabled — missing DEVICE_SETUP_PIN env var",
    };
  }

  let body: any = {};
  try {
    body = await req.json();
  } catch {
    body = {};
  }

  const pin = getBodyField(body, "pin", "setup_pin", "setupPin");
  if (!pin || pin !== DEVICE_SETUP_PIN) {
    return { ok: false, status: 401, error: "Invalid setup pin" };
  }

  const userName = getBodyField(body, "user_name", "userName", "user") ||
    "trusted-device";
  const token = generateUuid();

  await sqliteAuth.execute({
    sql: `INSERT INTO trusted_devices
            (device_token, user_name, role, last_seen_at, expires_at, created_at)
          VALUES (?, ?, 'full', datetime('now'), datetime('now', '+30 days'), datetime('now'))`,
    args: [token, userName],
  });

  return {
    ok: true,
    token,
    user_name: userName,
    created_at: new Date().toISOString(),
  };
}

export async function handleDeviceOtpRequest(req: Request): Promise<any> {
  try {
    await ensurePmProxyUsersTable();
    await ensureDeviceOtpsTable();
    const policy = await getOtpPolicy();
    if (!policy.enabled) {
      return {
        ok: false,
        status: 403,
        error: "OTP login is currently disabled by administrator.",
      };
    }

    let body: any = {};
    try {
      body = await req.json();
    } catch {
      body = {};
    }

    const requestedIdentifier = getBodyField(
      body,
      "identifier",
      "email",
      "phone",
    );

    // Rate limiting: prevent OTP flood against a single identifier (email/phone)
    // and cap overall OTP throughput to block mass account enumeration.
    if (requestedIdentifier) {
      const idKey = `otp_id:${requestedIdentifier.substring(0, 200)}`;
      const idRl = checkRateLimit(idKey, 3, 15 * 60_000); // 3 per 15 minutes per identifier
      if (!idRl.allowed) {
        return {
          ok: false,
          status: 429,
          error: "Too many OTP requests for this account — try again in 15 minutes.",
          retry_after_ms: idRl.retryAfterMs,
        };
      }
    }
    const globalRl = checkRateLimit("otp_global", 30, 60_000); // 30 per minute globally
    if (!globalRl.allowed) {
      return {
        ok: false,
        status: 429,
        error: "OTP service is temporarily busy — try again shortly.",
        retry_after_ms: globalRl.retryAfterMs,
      };
    }

    const otpAccessErr =
      "This PM email/phone is not enabled for OTP access. Please contact administrator.";

    const pmUser = await resolvePmProxyUser(requestedIdentifier);
    if (policy.requireMembership && !pmUser) {
      return { ok: false, status: 403, error: otpAccessErr };
    }

    const email = deriveOtpIdentityEmail(
      pmUser,
      requestedIdentifier,
      policy.allowedDomain,
    );
    const smsPhone = normalizePhone(pmUser?.phone || "");
    if (!email || !smsPhone) {
      return { ok: false, status: 403, error: otpAccessErr };
    }

    const userName = getBodyField(body, "user_name", "userName", "user") ||
      "trusted-device";
    const code = generateOtpCode();

    await insertDeviceOtp(
      email,
      code,
      policy.ttlMinutes,
      userName,
      String(pmUser?.property_group_uuid || ""),
    );

    try {
      await sendOtpSms(smsPhone, code, policy.ttlMinutes);
    } catch (err: any) {
      return {
        ok: false,
        status: 502,
        error: `Failed to send OTP SMS: ${String(err?.message || err)}`,
      };
    }

    return {
      ok: true,
      email,
      phone: smsPhone,
      expires_in_minutes: policy.ttlMinutes,
    };
  } catch (err: any) {
    console.error("[DEVICE_OTP_REQUEST_ERROR]", err?.message || String(err), err);
    return {
      ok: false,
      status: 500,
      error: "OTP request failed internally",
      message: String(err?.message || err),
    };
  }
}

export async function handleDeviceOtpVerify(req: Request): Promise<any> {
  await ensureTrustedDevicesTable();
  await ensurePmProxyUsersTable();
  const policy = await getOtpPolicy();
  if (!policy.enabled) {
    return {
      ok: false,
      status: 403,
      error: "OTP login is currently disabled by administrator.",
    };
  }

  let body: any = {};
  try {
    body = await req.json();
  } catch {
    body = {};
  }

  const requestedIdentifier = getBodyField(
    body,
    "identifier",
    "email",
    "phone",
  );
  const otpAccessErr =
    "This PM email/phone is not enabled for OTP access. Please contact administrator.";

  const pmUserCheck = await resolvePmProxyUser(requestedIdentifier);
  if (policy.requireMembership && !pmUserCheck) {
    return { ok: false, status: 403, error: otpAccessErr };
  }

  const email = deriveOtpIdentityEmail(
    pmUserCheck,
    requestedIdentifier,
    policy.allowedDomain,
  );
  const code = getBodyField(body, "code", "otp");
  if (!email || !code) {
    return { ok: false, status: 403, error: otpAccessErr };
  }

  const result = await sqlite.execute({
    sql: `SELECT id, code, expires_at, used
          FROM device_otps
          WHERE email = ?
          ORDER BY id DESC
          LIMIT 1`,
    args: [email],
  });

  const rows = rowsAsObjects(result);
  if (!rows.length) {
    return {
      ok: false,
      status: 401,
      error: "No OTP request found for this email",
    };
  }

  const row = rows[0] as any;
  if (Number(row.used || 0) === 1) {
    return { ok: false, status: 401, error: "OTP code already used" };
  }
  if (new Date(String(row.expires_at || "")).getTime() < Date.now()) {
    return { ok: false, status: 401, error: "OTP code expired" };
  }
  if (String(row.code || "") !== String(code).trim()) {
    return { ok: false, status: 401, error: "Invalid OTP code" };
  }

  await sqlite.execute({
    sql:
      `UPDATE device_otps SET used = 1, used_at = datetime('now') WHERE id = ?`,
    args: [row.id],
  });

  const pmUser = pmUserCheck || await resolvePmProxyUser(requestedIdentifier);
  const userName = getBodyField(body, "user_name", "userName", "user") ||
    (pmUser?.full_name || email);
  const token = generateUuid();
  const role = "pm_readonly";
  const scopeUuid = String(pmUser?.property_group_uuid || "");
  const phone = String(pmUser?.phone || "");

  await sqliteAuth.execute({
    sql: `INSERT INTO trusted_devices
            (device_token, user_name, role, login_email, property_group_uuid, phone,
             last_seen_at, expires_at, created_at)
          VALUES (?, ?, ?, ?, ?, ?, datetime('now'), datetime('now', '+30 days'), datetime('now'))`,
    args: [token, userName, role, email, scopeUuid, phone],
  });

  await sqlite.execute({
    sql: `INSERT INTO pm_proxy_login_audit
            (user_uuid, email, role, property_group_uuid, device_token, created_at)
          VALUES (?, ?, ?, ?, ?, datetime('now'))`,
    args: [pmUser?.user_uuid || "", email, role, scopeUuid, token],
  });

  return {
    ok: true,
    token,
    user_name: userName,
    email,
    role,
    property_group_uuid: scopeUuid,
    phone,
    created_at: new Date().toISOString(),
  };
}

export async function getTrustedDeviceSession(
  token: string,
): Promise<any | null> {
  if (!token) return null;
  await ensureTrustedDevicesTable();
  const result = await sqliteAuth.execute({
    sql: `SELECT device_token, user_name, role, login_email, property_group_uuid,
                 phone, created_at, last_seen_at, expires_at
          FROM trusted_devices
          WHERE device_token = ?
            AND revoked = 0
            AND (expires_at IS NULL OR expires_at > datetime('now'))
          LIMIT 1`,
    args: [token],
  });
  const rows = rowsAsObjects(result);
  if (!rows.length) return null;
  const row: any = rows[0];
  return {
    device_token: row.device_token,
    user_name: row.user_name || "",
    role: row.role || "full",
    login_email: row.login_email || "",
    property_group_uuid: row.property_group_uuid || "",
    phone: row.phone || "",
    created_at: row.created_at || "",
    last_seen_at: row.last_seen_at || "",
    expires_at: row.expires_at || "",
  };
}

// Slide the 30-day rolling expiry and touch last_seen_at.
// Called on every authenticated session_info request.
export async function touchDeviceSession(token: string): Promise<void> {
  if (!token) return;
  try {
    await sqliteAuth.execute({
      sql: `UPDATE trusted_devices
            SET last_seen_at = datetime('now'),
                expires_at   = datetime('now', '+30 days')
            WHERE device_token = ? AND revoked = 0`,
      args: [token],
    });
  } catch (e: any) {
    console.warn(
      "[touchDeviceSession] failed:",
      String(e?.message || e).substring(0, 120),
    );
  }
}

export async function handlePmProxyUsersList(
  params: Record<string, string>,
  req: Request,
): Promise<any> {
  try {
    await ensurePmProxyUsersTable();
    if (!PROXY_ADMIN_KEY) {
      return {
        ok: false,
        error: "pm user admin disabled — set PROXY_ADMIN_KEY",
      };
    }
    const key = params.key || req.headers.get("x-admin-key") || "";
    if (key !== PROXY_ADMIN_KEY) {
      return { ok: false, error: "Unauthorized: invalid admin key" };
    }
    const limit = Math.max(
      1,
      Math.min(500, parseInt(String(params.limit || "100"), 10) || 100),
    );
    const offset = Math.max(
      0,
      parseInt(String(params.offset || "0"), 10) || 0,
    );
    const totalResult = await sqlite.execute({
      sql: `SELECT COUNT(*) AS total FROM pm_proxy_users`,
    });
    const total = Number(rowsAsObjects(totalResult)[0]?.total || 0);
    const result = await sqlite.execute({
      sql:
        `SELECT user_uuid, email, full_name, phone, property_group_uuid, active, created_at, updated_at
            FROM pm_proxy_users
            ORDER BY active DESC, updated_at DESC
            LIMIT ? OFFSET ?`,
      args: [limit, offset],
    });
    if (!result || typeof result !== "object") {
      return { ok: true, results: [], count: 0 };
    }
    let rows: any[] = [];
    try {
      rows = rowsAsObjects(result) || [];
    } catch {
      // Guard against malformed DB payloads in fresh/edge fallback states.
      rows = [];
    }
    if (!Array.isArray(rows)) rows = [];
    return {
      ok: true,
      results: rows,
      count: rows.length,
      total,
      limit,
      offset,
    };
  } catch (err: any) {
    const errMsg = err?.message || String(err) || "Unknown error";
    console.error("[PM_PROXY_USERS_ERROR]", errMsg, err);
    return {
      ok: false,
      status: 500,
      error: `Failed to fetch PM users: ${errMsg}`,
    };
  }
}

export async function handlePmProxyUserUpsert(req: Request): Promise<any> {
  try {
    await ensurePmProxyUsersTable();
    if (!PROXY_ADMIN_KEY) {
      return {
        ok: false,
        error: "pm user admin disabled — set PROXY_ADMIN_KEY",
      };
    }
    let body: any = {};
    try {
      body = await req.json();
    } catch {
      body = {};
    }
    const key = getBodyField(body, "key", "admin_key");
    if (key !== PROXY_ADMIN_KEY) {
      return { ok: false, error: "Unauthorized: invalid admin key" };
    }
    const email = normalizeEmail(getBodyField(body, "email"));
    const propertyGroupUuid = getBodyField(
      body,
      "property_group_uuid",
      "scope_uuid",
    );
    if (!email || !propertyGroupUuid) {
      return {
        ok: false,
        status: 400,
        error: "Missing email or property_group_uuid",
      };
    }
    const userUuid = getBodyField(body, "user_uuid") || generateUuid();
    const fullName = getBodyField(body, "full_name", "name");
    const phone = getBodyField(body, "phone");
    const active = String(getBodyField(body, "active") || "1") === "0" ? 0 : 1;

    await sqlite.execute({
      sql: `INSERT INTO pm_proxy_users
              (user_uuid, email, full_name, phone, property_group_uuid, active, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, datetime('now'), datetime('now'))
            ON CONFLICT(email) DO UPDATE SET
              full_name=excluded.full_name,
              phone=excluded.phone,
              property_group_uuid=excluded.property_group_uuid,
              active=excluded.active,
              updated_at=datetime('now')`,
      args: [userUuid, email, fullName, phone, propertyGroupUuid, active],
    });

    return {
      ok: true,
      user_uuid: userUuid,
      email,
      full_name: fullName,
      phone,
      property_group_uuid: propertyGroupUuid,
      active,
    };
  } catch (err: any) {
    const errMsg = err?.message || String(err) || "Unknown error";
    console.error("[PM_PROXY_UPSERT_ERROR]", errMsg, err);
    return {
      ok: false,
      status: 500,
      error: `Failed to save PM user: ${errMsg}`,
    };
  }
}

export async function handlePmProxyUserDelete(req: Request): Promise<any> {
  try {
    await ensurePmProxyUsersTable();
    if (!PROXY_ADMIN_KEY) {
      return {
        ok: false,
        error: "pm user admin disabled — set PROXY_ADMIN_KEY",
      };
    }
    let body: any = {};
    try {
      body = await req.json();
    } catch {
      body = {};
    }
    const key = getBodyField(body, "key", "admin_key");
    if (key !== PROXY_ADMIN_KEY) {
      return { ok: false, error: "Unauthorized: invalid admin key" };
    }
    const userUuid = getBodyField(body, "user_uuid");
    const email = normalizeEmail(getBodyField(body, "email"));
    if (!userUuid && !email) {
      return { ok: false, status: 400, error: "Missing user_uuid or email" };
    }
    const sql = userUuid
      ? `DELETE FROM pm_proxy_users WHERE user_uuid = ?`
      : `DELETE FROM pm_proxy_users WHERE lower(email) = lower(?)`;
    const arg = userUuid || email;
    const result = await sqlite.execute({ sql, args: [arg] });
    return {
      ok: true,
      deleted: (result.rowsAffected || 0) > 0,
      rowsAffected: result.rowsAffected || 0,
    };
  } catch (err: any) {
    const errMsg = err?.message || String(err) || "Unknown error";
    console.error("[PM_PROXY_DELETE_ERROR]", errMsg, err);
    return {
      ok: false,
      status: 500,
      error: `Failed to delete PM user: ${errMsg}`,
    };
  }
}

export async function handleTrustedDeviceList(
  params: Record<string, string>,
  req: Request,
): Promise<any> {
  await ensureTrustedDevicesTable();
  if (!PROXY_ADMIN_KEY) {
    return {
      ok: false,
      error: "trusted device admin is disabled — set PROXY_ADMIN_KEY",
    };
  }

  const key = params.key || req.headers.get("x-admin-key") || "";
  if (key !== PROXY_ADMIN_KEY) {
    return { ok: false, error: "Unauthorized: invalid admin key" };
  }

  const limit = Math.max(
    1,
    Math.min(500, parseInt(String(params.limit || "100"), 10) || 100),
  );
  const offset = Math.max(
    0,
    parseInt(String(params.offset || "0"), 10) || 0,
  );

  const totalResult = await sqliteAuth.execute({
    sql: `SELECT COUNT(*) AS total FROM trusted_devices WHERE revoked = 0`,
  });
  const total = Number(rowsAsObjects(totalResult)[0]?.total || 0);

  const result = await sqliteAuth.execute({
    sql: `SELECT device_token, user_name, role, login_email, created_at, last_seen_at, expires_at
          FROM trusted_devices
          WHERE revoked = 0
          ORDER BY created_at DESC
          LIMIT ? OFFSET ?`,
    args: [limit, offset],
  });

  const devices = rowsAsObjects(result).map((d: any) => ({
    device_token: d.device_token,
    token_preview: String(d.device_token || "").slice(0, 8) + "...",
    user_name: d.user_name,
    role: d.role || "full",
    login_email: d.login_email || "",
    created_at: d.created_at,
    last_seen_at: d.last_seen_at || "",
    expires_at: d.expires_at || "",
  }));

  return {
    ok: true,
    results: devices,
    count: devices.length,
    total,
    limit,
    offset,
  };
}

export async function handleTrustedDeviceRevoke(
  params: Record<string, string>,
  req: Request,
): Promise<any> {
  await ensureTrustedDevicesTable();

  if (!PROXY_ADMIN_KEY) {
    return {
      ok: false,
      error: "trusted device admin is disabled — set PROXY_ADMIN_KEY",
    };
  }

  let body: any = {};
  try {
    body = await req.json();
  } catch {
    body = {};
  }

  const key = params.key || body.key || req.headers.get("x-admin-key") || "";
  if (key !== PROXY_ADMIN_KEY) {
    return { ok: false, error: "Unauthorized: invalid admin key" };
  }

  const token = getBodyField(body, "token", "device_token", "deviceToken") ||
    getBodyField(params, "token", "device_token", "deviceToken");

  if (!token) {
    return { ok: false, error: "Missing token/device_token" };
  }

  const result = await sqliteAuth.execute({
    sql: `UPDATE trusted_devices SET revoked = 1 WHERE device_token = ?`,
    args: [token],
  });

  return {
    ok: true,
    revoked: (result.rowsAffected || 0) > 0,
    rowsAffected: result.rowsAffected || 0,
  };
}

/**
 * Verify a GUI role password against Val.town env vars:
 *   GUI_ADMIN   → role "full"   (admin, all tabs)
 *   GUI_GM      → role "manager" (GM view, no dispatch/dbadmin)
 *   GUI_VENDORS → role "vendors" (vendor-only view)
 *
 * This endpoint is intentionally PUBLIC (no device token required) so that
 * a first-time login can verify the password and receive the access role.
 * The response only reveals the role name — never which env var matched.
 */
export async function handleVerifyRole(req: Request): Promise<any> {
  let body: any = {};
  try {
    body = await req.json();
  } catch {
    body = {};
  }

  const password = getBodyField(body, "password", "pass");
  if (!password) {
    return { ok: false, status: 400, error: "Password is required" };
  }

  const guiAdmin = Deno.env.get("GUI_ADMIN") || "";
  const guiGm = Deno.env.get("GUI_GM") || "";
  const guiVendors = Deno.env.get("GUI_VENDORS") || "";

  if (!guiAdmin && !guiGm && !guiVendors) {
    return {
      ok: false,
      status: 500,
      error:
        "GUI role passwords are not configured on the server. Set GUI_ADMIN, GUI_GM, and/or GUI_VENDORS in Val.town environment variables.",
    };
  }

  let matchedRole: string | null = null;
  if (guiAdmin && password === guiAdmin) matchedRole = "full";
  else if (guiGm && password === guiGm) matchedRole = "manager";
  else if (guiVendors && password === guiVendors) matchedRole = "vendors";

  if (!matchedRole) return { ok: false, status: 401, error: "Invalid password" };

  // Mint a trusted device token so the frontend can authenticate bearer requests.
  try {
    await ensureTrustedDevicesTable();
    const token = generateUuid();
    const userName = getBodyField(body, "user_name", "userName", "user") || "password-session";
    await sqliteAuth.execute({
      sql: `INSERT INTO trusted_devices
              (device_token, user_name, role, last_seen_at, expires_at, created_at)
            VALUES (?, ?, ?, datetime('now'), datetime('now', '+30 days'), datetime('now'))`,
      args: [token, userName, matchedRole],
    });
    return { ok: true, role: matchedRole, token };
  } catch (e: any) {
    // Do not return a tokenless success: frontend auth requires bearer token.
    console.warn("[handleVerifyRole] token mint failed:", String(e?.message || e));
    return {
      ok: false,
      status: 505,
      error:
        "Password verified but session token could not be created. Check proxy database write access and trusted_devices schema.",
    };
  }
}