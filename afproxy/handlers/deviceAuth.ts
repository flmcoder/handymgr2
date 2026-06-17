// ============================================================================
// handlers/deviceAuth.ts — PM OTP Login + Trusted Device Session Management
//
// Handles:
//   device_otp_request  — look up PM user by email OR phone → send OTP via SMS
//   device_otp_verify   — verify 6-digit OTP code → issue trusted device token
//   device_setup        — register device with shared secret PIN
//   verify_role         — return session role/group for an existing device token
//   trusted_device_list/revoke — admin management of sessions
//   pm_proxy_user_upsert/delete — admin CRUD for PM login accounts
//
// Lookup behaviour (FIXED v9.7.7C):
//   • Accepts either a saved email address or a saved phone number as identifier
//   • Phone numbers are normalised to last-10-digits before comparison so that
//     "(505) 399-2823", "505-399-2823", "+15053992823" all resolve to the same
//     stored account regardless of the format saved by the admin
//   • OTP is always sent to the phone on file for the found account
//   • "not enabled for OTP" error fires ONLY when no active account matches
// ============================================================================

import { rowsAsObjects, sqlite, sqliteAuth } from "../db.ts";
import { PROXY_ADMIN_KEY } from "../config.ts";
import { checkRateLimit } from "../lib/rateLimit.ts";
import { sendSMS } from "../lib/ringcentral.ts";

// ── Stateless HMAC-signed session tokens ─────────────────────────────────────
// Used as a fallback when the DB is unavailable for writes (e.g. Val Town free
// plan blocks SQL writes).  Format: "v1.<b64url-payload>.<b64url-sig>"
// where payload = base64url(JSON({r:role,u:userName,iat:epochSeconds})).
// Signed with HMAC-SHA256 keyed on FRONTEND_PROXY_SECRET.  Never stored in DB.

const _SESSION_SIGN_KEY_RAW = Deno.env.get("FRONTEND_PROXY_SECRET") || "";

function _b64url(buf: ArrayBuffer): string {
  return btoa(String.fromCharCode(...new Uint8Array(buf)))
    .replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
function _b64urlDecode(s: string): Uint8Array {
  const padded = s.replace(/-/g, "+").replace(/_/g, "/") +
    "==".slice(0, (4 - (s.length % 4)) % 4);
  const bin = atob(padded);
  return Uint8Array.from(bin, (c) => c.charCodeAt(0));
}

async function _getSignKey(): Promise<CryptoKey | null> {
  if (!_SESSION_SIGN_KEY_RAW) return null;
  try {
    return await crypto.subtle.importKey(
      "raw",
      new TextEncoder().encode(_SESSION_SIGN_KEY_RAW),
      { name: "HMAC", hash: "SHA-256" },
      false,
      ["sign", "verify"],
    );
  } catch {
    return null;
  }
}

async function mintSignedToken(
  role: string,
  userName: string,
): Promise<string | null> {
  const key = await _getSignKey();
  if (!key) return null;
  const payload = _b64url(
    new TextEncoder().encode(
      JSON.stringify({ r: role, u: userName, iat: Math.floor(Date.now() / 1000) }),
    ).buffer,
  );
  const sig = _b64url(
    await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(payload)),
  );
  return `v1.${payload}.${sig}`;
}

async function verifySignedToken(
  token: string,
): Promise<{ role: string; userName: string; iat: number } | null> {
  if (!token || !token.startsWith("v1.")) return null;
  const parts = token.split(".");
  if (parts.length !== 3) return null;
  const [, payload64, sig64] = parts;
  const key = await _getSignKey();
  if (!key) return null;
  try {
    const valid = await crypto.subtle.verify(
      "HMAC",
      key,
      _b64urlDecode(sig64).buffer as ArrayBuffer,
      new TextEncoder().encode(payload64),
    );
    if (!valid) return null;
    const data = JSON.parse(new TextDecoder().decode(_b64urlDecode(payload64)));
    if (!data || !data.r || typeof data.iat !== "number") return null;
    // Tokens expire after 30 days (2592000 seconds).
    if (Math.floor(Date.now() / 1000) - data.iat > 2592000) return null;
    return { role: String(data.r), userName: String(data.u || ""), iat: data.iat };
  } catch {
    return null;
  }
}

let _trustedDevicesTableReady = false;
let _pmProxyUsersTableReady = false;
let _deviceOtpsTableReady = false;
let _proxyConfigTableReady = false;
let _pmProxyLoginAuditTableReady = false;
let _trustedDevicesDbFallbackWarned = false;
const RECENT_TRUSTED_SESSION_TTL_MS = 10 * 60 * 1000;
const _recentTrustedSessions = new Map<string, any>();

function rememberRecentTrustedSession(session: {
  device_token: string;
  user_name: string;
  role?: string;
  login_email?: string;
  property_group_uuid?: string;
  phone?: string;
  created_at?: string;
  last_seen_at?: string;
  expires_at?: string;
}): void {
  const token = String(session.device_token || '').trim();
  if (!token) return;
  _recentTrustedSessions.set(token, {
    device_token: token,
    user_name: String(session.user_name || ''),
    role: String(session.role || 'full') || 'full',
    login_email: String(session.login_email || ''),
    property_group_uuid: String(session.property_group_uuid || ''),
    phone: String(session.phone || ''),
    created_at: String(session.created_at || new Date().toISOString()),
    last_seen_at: String(session.last_seen_at || new Date().toISOString()),
    expires_at: String(session.expires_at || new Date(Date.now() + (30 * 24 * 60 * 60 * 1000)).toISOString()),
    cached_at_ms: Date.now(),
  });
}

function readRecentTrustedSession(token: string): any | null {
  const entry = _recentTrustedSessions.get(String(token || '').trim());
  if (!entry) return null;
  if ((Date.now() - Number(entry.cached_at_ms || 0)) > RECENT_TRUSTED_SESSION_TTL_MS) {
    _recentTrustedSessions.delete(String(token || '').trim());
    return null;
  }
  return {
    device_token: entry.device_token,
    user_name: entry.user_name || '',
    role: entry.role || 'full',
    login_email: entry.login_email || '',
    property_group_uuid: entry.property_group_uuid || '',
    phone: entry.phone || '',
    created_at: entry.created_at || '',
    last_seen_at: entry.last_seen_at || '',
    expires_at: entry.expires_at || '',
  };
}

function forgetRecentTrustedSession(token: string): void {
  _recentTrustedSessions.delete(String(token || '').trim());
}

function logTrustedDevicesDbFallback(err: unknown): void {
  if (_trustedDevicesDbFallbackWarned) return;
  _trustedDevicesDbFallbackWarned = true;
  console.warn(
    "[trusted_devices] sqliteAuth unavailable, falling back to sqlite:",
    String((err as any)?.message || err || "unknown error").substring(0, 180),
  );
}

async function executeTrustedDevicesSql(
  sql: string,
  args: any[] = [],
): Promise<any> {
  try {
    return await sqliteAuth.execute({ sql, args });
  } catch (err) {
    logTrustedDevicesDbFallback(err);
    return await sqlite.execute({ sql, args });
  }
}

async function executeTrustedDevicesSqlOnBoth(
  sql: string,
  args: any[] = [],
): Promise<void> {
  try {
    await sqliteAuth.execute({ sql, args });
  } catch (err) {
    logTrustedDevicesDbFallback(err);
    try {
      await sqlite.execute({ sql, args });
    } catch (_) {}
    return;
  }
  try {
    await sqlite.execute({ sql, args });
  } catch (_) {}
}

function isMissingColumnError(err: unknown): boolean {
  const msg = String((err as any)?.message || err || "").toLowerCase();
  return msg.includes("no such column") || msg.includes("has no column named");
}

async function insertTrustedDeviceSession(args: {
  token: string;
  userName: string;
  role?: string;
  loginEmail?: string;
  propertyGroupUuid?: string;
  phone?: string;
}): Promise<void> {
  const token = String(args.token || "").trim();
  const userName = String(args.userName || "").trim() || "trusted-device";
  const role = String(args.role || "full").trim() || "full";
  const loginEmail = String(args.loginEmail || "").trim();
  const propertyGroupUuid = String(args.propertyGroupUuid || "").trim();
  const phone = String(args.phone || "").trim();

  try {
    await executeTrustedDevicesSqlOnBoth(
      `INSERT INTO trusted_devices
         (device_token, user_name, role, login_email, property_group_uuid, phone,
          last_seen_at, expires_at, created_at)
       VALUES (?, ?, ?, ?, ?, ?, datetime('now'), datetime('now', '+30 days'), datetime('now'))`,
      [token, userName, role, loginEmail, propertyGroupUuid, phone],
    );
    rememberRecentTrustedSession({
      device_token: token,
      user_name: userName,
      role,
      login_email: loginEmail,
      property_group_uuid: propertyGroupUuid,
      phone,
    });
    return;
  } catch (err: any) {
    if (!isMissingColumnError(err)) throw err;
  }

  try {
    await executeTrustedDevicesSqlOnBoth(
      `INSERT INTO trusted_devices
         (device_token, user_name, role, last_seen_at, expires_at, created_at)
       VALUES (?, ?, ?, datetime('now'), datetime('now', '+30 days'), datetime('now'))`,
      [token, userName, role],
    );
    rememberRecentTrustedSession({
      device_token: token,
      user_name: userName,
      role,
      login_email: loginEmail,
      property_group_uuid: propertyGroupUuid,
      phone,
    });
    return;
  } catch (err: any) {
    if (!isMissingColumnError(err)) throw err;
  }

  await executeTrustedDevicesSqlOnBoth(
    `INSERT INTO trusted_devices (device_token, user_name, created_at)
     VALUES (?, ?, datetime('now'))`,
    [token, userName],
  );
  rememberRecentTrustedSession({
    device_token: token,
    user_name: userName,
    role,
    login_email: loginEmail,
    property_group_uuid: propertyGroupUuid,
    phone,
  });
}

function generateUuid(): string {
  if (globalThis.crypto && globalThis.crypto.randomUUID) {
    return globalThis.crypto.randomUUID();
  }
  const bytes = new Uint8Array(16);
  if (globalThis.crypto && globalThis.crypto.getRandomValues) {
    globalThis.crypto.getRandomValues(bytes);
  } else {
    throw new Error("Crypto API unavailable — cannot generate secure UUIDs");
  }
  bytes[6] = (bytes[6] & 0x0f) | 0x40; // version 4
  bytes[8] = (bytes[8] & 0x3f) | 0x80; // variant
  const hex = Array.from(bytes).map((b) => b.toString(16).padStart(2, "0")).join("");
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
}

async function ensureTrustedDevicesTable(): Promise<void> {
  if (_trustedDevicesTableReady) return;
  await executeTrustedDevicesSqlOnBoth(
    `CREATE TABLE IF NOT EXISTS trusted_devices (
       device_token TEXT PRIMARY KEY,
       user_name    TEXT,
       created_at   TEXT NOT NULL DEFAULT (datetime('now'))
     )`,
  );
  const cols = [
    "role TEXT DEFAULT 'full'",
    "login_email TEXT",
    "property_group_uuid TEXT",
    "phone TEXT",
    "last_seen_at TEXT",
    "expires_at TEXT",
    "revoked INTEGER DEFAULT 0",
    "auth_source TEXT"
  ];
  for (const col of cols) {
    try { await executeTrustedDevicesSqlOnBoth(`ALTER TABLE trusted_devices ADD COLUMN ${col}`); } catch (_) {}
  }
  _trustedDevicesTableReady = true;
}

async function ensurePmProxyUsersTable(): Promise<void> {
  if (_pmProxyUsersTableReady) return;
  await sqlite.execute(`CREATE TABLE IF NOT EXISTS pm_proxy_users (
    id                  TEXT,
    user_uuid           TEXT PRIMARY KEY,
    email               TEXT UNIQUE NOT NULL,
    full_name           TEXT,
    phone               TEXT,
    property_group_uuid TEXT NOT NULL,
    roles               TEXT NOT NULL DEFAULT '[]',
    is_active           INTEGER DEFAULT 1,
    raw_json            TEXT NOT NULL DEFAULT '{}',
    active              INTEGER DEFAULT 1,
    created_at          TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at          TEXT NOT NULL DEFAULT (datetime('now'))
  )`);
  
  const cols = [
    "id TEXT",
    "roles TEXT NOT NULL DEFAULT '[]'",
    "is_active INTEGER DEFAULT 1",
    "raw_json TEXT NOT NULL DEFAULT '{}'",
    "created_at TEXT DEFAULT (datetime('now'))",
    "updated_at TEXT DEFAULT (datetime('now'))"
  ];
  for (const col of cols) {
      try { await sqlite.execute(`ALTER TABLE pm_proxy_users ADD COLUMN ${col}`); } catch (_) {}
  }
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
  const cols = ["user_name TEXT", "created_at TEXT", "used_at TEXT", "role_hint TEXT", "property_group_uuid TEXT"];
  for (const col of cols) {
    try { await sqlite.execute(`ALTER TABLE device_otps ADD COLUMN ${col}`); } catch (_) {}
  }
  _deviceOtpsTableReady = true;
}

async function ensureProxyConfigTable(): Promise<void> {
  if (_proxyConfigTableReady) return;
  await sqlite.execute(`CREATE TABLE IF NOT EXISTS proxy_config (
    key        TEXT PRIMARY KEY,
    value      TEXT,
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
  )`);
  try {
    await sqlite.execute(
      `CREATE INDEX IF NOT EXISTS idx_proxy_config_key ON proxy_config(key)`,
    );
  } catch (_) {}
  _proxyConfigTableReady = true;
}

async function ensurePmProxyLoginAuditTable(): Promise<void> {
  if (_pmProxyLoginAuditTableReady) return;
  await sqlite.execute(`CREATE TABLE IF NOT EXISTS pm_proxy_login_audit (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    user_uuid           TEXT,
    email               TEXT,
    role                TEXT,
    property_group_uuid TEXT,
    device_token        TEXT,
    created_at          TEXT NOT NULL DEFAULT (datetime('now'))
  )`);
  _pmProxyLoginAuditTableReady = true;
}

const DEVICE_SETUP_PIN = Deno.env.get("DEVICE_SETUP_PIN") || "";
const OTP_ALLOWED_DOMAIN_ENV = (Deno.env.get("OTP_ALLOWED_DOMAIN") || "flraz.com").replace(/^@/, "").toLowerCase();
const OTP_TTL_MINUTES_ENV = Math.max(3, Number(Deno.env.get("DEVICE_OTP_TTL_MINUTES") || "10") || 10);

async function getProxyConfig(key: string, fallback: string): Promise<string> {
  await ensureProxyConfigTable();
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

async function getOtpPolicy(): Promise<{
  enabled: boolean;
  allowedDomain: string;
  requireMembership: boolean;
  ttlMinutes: number;
}> {
  await ensureProxyConfigTable();
  try {
    const res = await sqlite.execute({
      sql: `SELECT key, value FROM proxy_config
            WHERE key IN ('otp_enabled','otp_allowed_domain','otp_require_pm_membership','otp_ttl_minutes')`,
    });
    const rows = rowsAsObjects(res);
    const map: Record<string, string> = {};
    for (const r of rows) map[String(r.key)] = String(r.value);
    const rawDomain = (map["otp_allowed_domain"] || "").replace(/^@/, "").toLowerCase();
    return {
      enabled: (map["otp_enabled"] ?? "1") !== "0",
      allowedDomain: rawDomain || OTP_ALLOWED_DOMAIN_ENV,
      requireMembership: (map["otp_require_pm_membership"] ?? "1") !== "0",
      ttlMinutes: Math.max(3, Number(map["otp_ttl_minutes"]) || OTP_TTL_MINUTES_ENV),
    };
  } catch (_) {
    return {
      enabled: true,
      allowedDomain: OTP_ALLOWED_DOMAIN_ENV,
      // Keep OTP reachable during bootstrap/drift when PM rows are not yet synced.
      requireMembership: false,
      ttlMinutes: OTP_TTL_MINUTES_ENV,
    };
  }
}

async function sendOtpSms(toPhone: string, code: string, ttlMinutes: number): Promise<void> {
  const message = `HandyManager verification code: ${code}. Expires in ${ttlMinutes} minutes. If you did not request this code, ignore this message.`;
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
  const direct = normalizeOrgEmail(pmUser?.email || requestedIdentifier, allowedDomain);
  if (direct) return direct;
  const domain = String(allowedDomain || "").trim().toLowerCase();
  if (!domain) return "";
  const phoneKey = phoneMatchDigits(requestedIdentifier || "");
  if (!pmUser && phoneKey) {
    return `pm-phone-${phoneKey}@${domain}`;
  }
  if (!pmUser) return "";
  const userKey = String(pmUser.user_uuid || pmUser.email || "").toLowerCase()
    .replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
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

/** Strip everything except digits and return the last 10. (v9.7.7C) */
function last10Digits(raw: string): string {
  const digits = String(raw || "").replace(/\D/g, "");
  return digits.slice(-10);
}

/** Remove extension-like suffixes (x123, ext 123) before digit parsing. */
function stripPhoneExtension(raw: string): string {
  return String(raw || "").replace(/(?:ext\.?|extension|x|#)\s*\d+\s*$/i, "").trim();
}

/** Normalize phone into a stable 10-digit key for account matching. */
function phoneMatchDigits(rawPhone: string): string {
  const withoutExt = stripPhoneExtension(rawPhone);
  const digits = withoutExt.replace(/\D/g, "");
  if (digits.length === 10) return digits;
  if (digits.length === 11 && digits.startsWith("1")) return digits.slice(1);
  return last10Digits(digits);
}

/**
 * Build match keys for flexible lookup across country-code and local formats.
 * Keys include full digits, optional NANP local 10, and last-10 fallback.
 */
function phoneMatchKeys(rawPhone: string): string[] {
  const withoutExt = stripPhoneExtension(rawPhone);
  const digits = withoutExt.replace(/\D/g, "");
  if (!digits) return [];

  const keys: string[] = [digits];
  if (digits.length === 11 && digits.startsWith("1")) {
    keys.push(digits.slice(1));
  }
  if (digits.length >= 10) {
    keys.push(last10Digits(digits));
  }

  return Array.from(new Set(keys.filter(Boolean)));
}

/** Normalise a phone string to E.164 (+1XXXXXXXXXX) or return empty string. (v9.7.7C) */
function normalizePhone(rawPhone: string): string {
  const digits = stripPhoneExtension(rawPhone).replace(/\D/g, "");
  if (digits.length === 10) return `+1${digits}`;
  if (digits.length === 11 && digits.startsWith("1")) return `+${digits}`;
  if (digits.length > 11) return `+${digits}`;
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

function parseJsonObject(raw: unknown): Record<string, any> {
  if (!raw) return {};
  if (typeof raw === "object") return raw as Record<string, any>;
  try {
    const parsed = JSON.parse(String(raw));
    return parsed && typeof parsed === "object" ? parsed as Record<string, any> : {};
  } catch {
    return {};
  }
}

function firstNonEmpty(...values: unknown[]): string {
  for (const value of values) {
    const normalized = String(value ?? "").trim();
    if (normalized) return normalized;
  }
  return "";
}

function collectPrimitiveStrings(raw: unknown, sink: string[]): void {
  if (raw === null || raw === undefined) return;
  if (typeof raw === "string" || typeof raw === "number") {
    const normalized = String(raw).trim();
    if (normalized) sink.push(normalized);
    return;
  }
  if (Array.isArray(raw)) {
    for (const entry of raw) collectPrimitiveStrings(entry, sink);
    return;
  }
  if (typeof raw === "object") {
    const obj = raw as Record<string, unknown>;
    for (const key of Object.keys(obj)) {
      collectPrimitiveStrings(obj[key], sink);
    }
  }
}

function collectPhoneLikeStrings(raw: unknown, sink: string[]): void {
  if (raw === null || raw === undefined) return;
  if (typeof raw === "string" || typeof raw === "number") {
    const normalized = String(raw).trim();
    if (normalized) sink.push(normalized);
    return;
  }
  if (Array.isArray(raw)) {
    for (const entry of raw) collectPhoneLikeStrings(entry, sink);
    return;
  }
  if (typeof raw === "object") {
    const obj = raw as Record<string, unknown>;
    const directCandidates = [
      obj.number,
      obj.phone,
      obj.phone_number,
      obj.phoneNumber,
      obj.mobile,
      obj.mobile_phone,
      obj.mobilePhone,
      obj.cell,
      obj.cell_phone,
      obj.cellPhone,
      obj.value,
      obj.e164,
      obj.e_164,
    ];
    for (const candidate of directCandidates) {
      if (typeof candidate === "string" || typeof candidate === "number") {
        const normalized = String(candidate).trim();
        if (normalized) sink.push(normalized);
      }
    }
    for (const key of Object.keys(obj)) {
      if (!/(phone|mobile|cell|contact)/i.test(key)) continue;
      collectPhoneLikeStrings(obj[key], sink);
    }
  }
}

function possiblePhoneValues(row: Record<string, any>): string[] {
  const rawJson = parseJsonObject(row.raw_json || row.rawJson || row.profile_json || row.profileJson);
  const nestedContact = parseJsonObject(rawJson.contact || rawJson.contact_info || rawJson.contactInfo);
  const directValues = [
    row.phone,
    row.phone_number,
    row.phoneNumber,
    row.mobile_phone,
    row.mobilePhone,
    row.mobile,
    row.cell,
    row.cell_phone,
    row.cellPhone,
    row.phone_numbers,
    row.phoneNumbers,
    rawJson.phone,
    rawJson.phone_number,
    rawJson.phoneNumber,
    rawJson.mobile_phone,
    rawJson.mobilePhone,
    rawJson.mobile,
    rawJson.cell,
    rawJson.cell_phone,
    rawJson.cellPhone,
    rawJson.phone_numbers,
    rawJson.phoneNumbers,
    nestedContact.phone,
    nestedContact.phone_number,
    nestedContact.phoneNumber,
    nestedContact.mobile,
    nestedContact.mobile_phone,
    nestedContact.mobilePhone,
    nestedContact.cell,
    nestedContact.cell_phone,
    nestedContact.cellPhone,
  ];

  const values: string[] = [];
  for (const entry of directValues) {
    collectPhoneLikeStrings(entry, values);
  }
  collectPhoneLikeStrings(rawJson, values);
  collectPhoneLikeStrings(nestedContact, values);

  return Array.from(new Set(values.map((v) => String(v || "").trim()).filter(Boolean)));
}

function possibleGroupValues(row: Record<string, any>): string[] {
  const rawJson = parseJsonObject(row.raw_json || row.rawJson || row.profile_json || row.profileJson);
  const nestedScope = parseJsonObject(rawJson.scope || rawJson.group || rawJson.property_group || rawJson.propertyGroup);
  const values: string[] = [];
  const directCandidates = [
    row.property_group_uuid,
    row.propertyGroupUuid,
    row.property_group_id,
    row.propertyGroupId,
    row.group_uuid,
    row.groupUuid,
    row.group_id,
    row.groupId,
    row.scope_uuid,
    row.scopeUuid,
    row.scope_group_uuid,
    row.scopeGroupUuid,
    rawJson.property_group_uuid,
    rawJson.propertyGroupUuid,
    rawJson.property_group_id,
    rawJson.propertyGroupId,
    rawJson.group_uuid,
    rawJson.groupUuid,
    rawJson.group_id,
    rawJson.groupId,
    rawJson.scope_uuid,
    rawJson.scopeUuid,
    rawJson.scope_group_uuid,
    rawJson.scopeGroupUuid,
    nestedScope.property_group_uuid,
    nestedScope.propertyGroupUuid,
    nestedScope.property_group_id,
    nestedScope.propertyGroupId,
    nestedScope.group_uuid,
    nestedScope.groupUuid,
    nestedScope.group_id,
    nestedScope.groupId,
    nestedScope.uuid,
    nestedScope.id,
  ];

  for (const entry of directCandidates) {
    collectPrimitiveStrings(entry, values);
  }
  return Array.from(new Set(values.map((v) => String(v || "").trim()).filter(Boolean)));
}

function toPmProxyUserRow(row: Record<string, any>): PmProxyUser {
  const rawJson = parseJsonObject(row.raw_json || row.rawJson || row.profile_json || row.profileJson);
  const phoneValues = possiblePhoneValues(row);
  const groupValues = possibleGroupValues(row);
  return {
    user_uuid: firstNonEmpty(row.user_uuid, row.userUuid, row.id, rawJson.user_uuid, rawJson.userUuid),
    email: firstNonEmpty(row.email, row.login_email, row.loginEmail, rawJson.email),
    full_name: firstNonEmpty(row.full_name, row.fullName, row.name, rawJson.full_name, rawJson.fullName, rawJson.name),
    phone: firstNonEmpty(...phoneValues),
    property_group_uuid: firstNonEmpty(...groupValues),
    active: Number(row.active ?? row.is_active ?? rawJson.active ?? 1) === 0 ? 0 : 1,
  };
}

async function getPmProxyUserByEmail(email: string): Promise<PmProxyUser | null> {
  const normalized = normalizeEmail(email);
  if (!normalized) return null;
  try {
    const result = await sqlite.execute({
      sql: `SELECT *
            FROM pm_proxy_users
            WHERE lower(email) = lower(?) AND active = 1
            LIMIT 1`,
      args: [normalized],
    });
    const rows = rowsAsObjects(result || { rows: [], columns: [] }) as Record<string, any>[];
    return rows.length ? toPmProxyUserRow(rows[0]) : null;
  } catch (_) { return null; }
}

async function getPmUserAccountsByEmail(email: string): Promise<PmProxyUser | null> {
  const normalized = normalizeEmail(email);
  if (!normalized) return null;
  try {
    const result = await sqlite.execute({
      sql: `SELECT *
            FROM pm_user_accounts
            WHERE lower(email) = lower(?) AND coalesce(active,1) = 1
            LIMIT 1`,
      args: [normalized],
    });
    const rows = rowsAsObjects(result || { rows: [], columns: [] }) as Record<string, any>[];
    return rows.length ? toPmProxyUserRow(rows[0]) : null;
  } catch (_) { return null; }
}

async function getPmProxyUserByPhone(phone: string): Promise<PmProxyUser | null> {
  const inputKeys = phoneMatchKeys(phone);
  if (!inputKeys.length) return null;
  try {
    const result = await sqlite.execute({
      sql: `SELECT *
            FROM pm_proxy_users
            WHERE active = 1`,
    });
    const rows = rowsAsObjects(result || { rows: [], columns: [] }) as Record<string, any>[];
    for (const rawRow of rows) {
      const row = toPmProxyUserRow(rawRow);
      const candidates = possiblePhoneValues(rawRow);
      const storedKeys = Array.from(new Set(candidates.flatMap((c) => phoneMatchKeys(c))));
      if (storedKeys.some((k) => inputKeys.includes(k))) {
        return row;
      }
    }
  } catch (_) { return null; }
  return null;
}

async function getPmUserAccountsByPhone(phone: string): Promise<PmProxyUser | null> {
  const inputKeys = phoneMatchKeys(phone);
  if (!inputKeys.length) return null;
  try {
    const result = await sqlite.execute({
      sql: `SELECT *
            FROM pm_user_accounts
            WHERE coalesce(active,1) = 1`,
    });
    const rows = rowsAsObjects(result || { rows: [], columns: [] }) as Record<string, any>[];
    for (const rawRow of rows) {
      const row = toPmProxyUserRow(rawRow);
      const candidates = possiblePhoneValues(rawRow);
      const storedKeys = Array.from(new Set(candidates.flatMap((c) => phoneMatchKeys(c))));
      if (storedKeys.some((k) => inputKeys.includes(k))) {
        return row;
      }
    }
  } catch (_) { return null; }
  return null;
}

async function resolvePmProxyUser(identifier: string): Promise<PmProxyUser | null> {
  const byEmail = await getPmProxyUserByEmail(identifier) || await getPmUserAccountsByEmail(identifier);
  if (byEmail) return byEmail;
  return await getPmProxyUserByPhone(identifier) || await getPmUserAccountsByPhone(identifier);
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
      sql: `INSERT INTO device_otps (email, code, used, expires_at, user_name, role_hint, property_group_uuid)
            VALUES (?, ?, 0, datetime('now', ?), ?, ?, ?)`,
      args: [email, code, `+${ttlMinutes} minutes`, userName, "pm_readonly", scopeUuid],
    });
    return;
  } catch (err: any) {
    const msg = String(err?.message || err || "").toLowerCase();
    if (msg.includes("no such column") || msg.includes("has no column named")) {
      await sqlite.execute({
        sql: `INSERT INTO device_otps (email, code, used, expires_at, user_name)
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
  try { body = await req.json(); } catch { body = {}; }

  const pin = getBodyField(body, "pin", "setup_pin", "setupPin");
  if (!pin || pin !== DEVICE_SETUP_PIN) {
    return { ok: false, status: 401, error: "Invalid setup pin" };
  }

  const userName = getBodyField(body, "user_name", "userName", "user") || "trusted-device";
  const token = generateUuid();

  await insertTrustedDeviceSession({ token, userName, role: "full" });

  return { ok: true, token, user_name: userName, created_at: new Date().toISOString() };
}

export async function handleDeviceOtpRequest(req: Request): Promise<any> {
  try {
    await ensurePmProxyUsersTable();
    await ensureDeviceOtpsTable();
    await ensureProxyConfigTable();
    const policy = await getOtpPolicy();
    if (!policy.enabled) {
      return { ok: false, status: 403, error: "OTP login is currently disabled by administrator." };
    }

    let body: any = {};
    try { body = await req.json(); } catch { body = {}; }

    const requestedIdentifier = getBodyField(body, "identifier", "email", "phone");

    // Rate limiting
    if (requestedIdentifier) {
      const idKey = `otp_id:${requestedIdentifier.substring(0, 200)}`;
      const idRl = checkRateLimit(idKey, 3, 15 * 60_000); 
      if (!idRl.allowed) {
        return {
          ok: false, status: 429,
          error: "Too many OTP requests for this account — try again in 15 minutes.",
          retry_after_ms: idRl.retryAfterMs,
        };
      }
    }
    const globalRl = checkRateLimit("otp_global", 30, 60_000); 
    if (!globalRl.allowed) {
      return {
        ok: false, status: 429,
        error: "OTP service is temporarily busy — try again shortly.",
        retry_after_ms: globalRl.retryAfterMs,
      };
    }

    const otpAccessErr = "No active PM account found for that identifier. Please verify the identifier you registered with, or contact your administrator.";

    const pmUser = await resolvePmProxyUser(requestedIdentifier);
    if (policy.requireMembership && !pmUser) {
      return { ok: false, status: 403, error: otpAccessErr };
    }

    const email = deriveOtpIdentityEmail(pmUser, requestedIdentifier, policy.allowedDomain);
    const fallbackPhone = normalizePhone(requestedIdentifier);
    const smsPhone = normalizePhone(pmUser?.phone || "") ||
      (!policy.requireMembership ? fallbackPhone : "");
    
    if (!email || !smsPhone) {
      return { ok: false, status: 403, error: "No phone number on file for this PM account. Please ask your administrator to add a phone number before using OTP login." };
    }

    const userName = getBodyField(body, "user_name", "userName", "user") || "trusted-device";
    const code = generateOtpCode();

    await insertDeviceOtp(email, code, policy.ttlMinutes, userName, String(pmUser?.property_group_uuid || ""));

    try {
      await sendOtpSms(smsPhone, code, policy.ttlMinutes);
    } catch (err: any) {
      return { ok: false, status: 502, error: `Failed to send OTP SMS: ${String(err?.message || err)}` };
    }

    const phoneHintOut = smsPhone.length > 7
      ? smsPhone.slice(0, 5) + "***" + smsPhone.slice(-2)
      : "***";

    return {
      ok: true,
      message: pmUser
        ? `OTP sent to phone on file for ${pmUser?.full_name || "PM User"}.`
        : "OTP sent to the provided phone number.",
      phone_hint: phoneHintOut,
      expires_in_minutes: policy.ttlMinutes,
    };
  } catch (err: any) {
    console.error("[DEVICE_OTP_REQUEST_ERROR]", err?.message || String(err), err);
    return {
      ok: false, status: 500, error: "OTP request failed internally", message: String(err?.message || err)
    };
  }
}

export async function handleDeviceOtpVerify(req: Request): Promise<any> {
  await ensureTrustedDevicesTable();
  await ensurePmProxyUsersTable();
  await ensureDeviceOtpsTable();
  await ensureProxyConfigTable();
  await ensurePmProxyLoginAuditTable();
  const policy = await getOtpPolicy();
  if (!policy.enabled) {
    return { ok: false, status: 403, error: "OTP login is currently disabled by administrator." };
  }

  let body: any = {};
  try { body = await req.json(); } catch { body = {}; }

  const requestedIdentifier = getBodyField(body, "identifier", "email", "phone");
  const otpAccessErr = "No active PM account found for that identifier.";

  const pmUserCheck = await resolvePmProxyUser(requestedIdentifier);
  if (policy.requireMembership && !pmUserCheck) {
    return { ok: false, status: 403, error: otpAccessErr };
  }

  const email = deriveOtpIdentityEmail(pmUserCheck, requestedIdentifier, policy.allowedDomain);
  const code = getBodyField(body, "code", "otp");
  if (!email || !code) {
    return { ok: false, status: 403, error: otpAccessErr };
  }

  const result = await sqlite.execute({
    sql: `SELECT id, code, expires_at, used FROM device_otps WHERE email = ? ORDER BY id DESC LIMIT 1`,
    args: [email],
  });

  const rows = rowsAsObjects(result);
  if (!rows.length) return { ok: false, status: 401, error: "No OTP request found for this email" };

  const row = rows[0] as any;
  if (Number(row.used || 0) === 1) return { ok: false, status: 401, error: "OTP code already used" };
  if (new Date(String(row.expires_at || "")).getTime() < Date.now()) return { ok: false, status: 401, error: "OTP code expired" };
  if (String(row.code || "") !== String(code).trim()) return { ok: false, status: 401, error: "Invalid OTP code" };

  await sqlite.execute({
    sql: `UPDATE device_otps SET used = 1, used_at = datetime('now') WHERE id = ?`,
    args: [row.id],
  });

  const pmUser = pmUserCheck || await resolvePmProxyUser(requestedIdentifier);
  const userName = getBodyField(body, "user_name", "userName", "user") || (pmUser?.full_name || email);
  const token = generateUuid();
  const role = "pm_readonly";
  const scopeUuid = String(pmUser?.property_group_uuid || "");
  const phone = String(pmUser?.phone || "");

  await insertTrustedDeviceSession({ token, userName, role, loginEmail: email, propertyGroupUuid: scopeUuid, phone });

  try {
    await sqlite.execute({
      sql: `INSERT INTO pm_proxy_login_audit (user_uuid, email, role, property_group_uuid, device_token, created_at) VALUES (?, ?, ?, ?, ?, datetime('now'))`,
      args: [pmUser?.user_uuid || "", email, role, scopeUuid, token],
    });
  } catch (_) {}

  return { ok: true, token, user_name: userName, email, role, property_group_uuid: scopeUuid, phone, created_at: new Date().toISOString() };
}

export async function getTrustedDeviceSession(token: string): Promise<any | null> {
  if (!token) return null;

  if (token.startsWith("v1.")) {
    const decoded = await verifySignedToken(token);
    if (decoded) {
      return {
        device_token: token,
        user_name: decoded.userName || "password-session",
        role: decoded.role,
        login_email: "",
        property_group_uuid: "",
        phone: "",
        created_at: new Date(decoded.iat * 1000).toISOString(),
        last_seen_at: new Date().toISOString(),
        expires_at: new Date((decoded.iat + 2592000) * 1000).toISOString(),
      };
    }
    return null;
  }

  const recent = readRecentTrustedSession(token);
  if (recent) return recent;
  await ensureTrustedDevicesTable();

  async function selectSessionModern(exec: (q: string, a?: any[]) => Promise<any>) {
    const result = await exec(
      `SELECT device_token, user_name, role, login_email, property_group_uuid, phone, created_at, last_seen_at, expires_at
         FROM trusted_devices
        WHERE device_token = ? AND revoked = 0 AND (expires_at IS NULL OR expires_at > datetime('now')) LIMIT 1`,
      [token],
    );
    return rowsAsObjects(result);
  }

  async function selectSessionLegacy(exec: (q: string, a?: any[]) => Promise<any>) {
    const legacy = await exec(`SELECT device_token, user_name, created_at FROM trusted_devices WHERE device_token = ? LIMIT 1`, [token]);
    return rowsAsObjects(legacy);
  }

  const authExec = (sql: string, args: any[] = []) => sqliteAuth.execute({ sql, args });
  const fallbackExec = (sql: string, args: any[] = []) => sqlite.execute({ sql, args });

  try {
    const rows = await selectSessionModern(authExec);
    if (!rows.length) {
      const fallbackRows = await selectSessionModern(fallbackExec).catch(() => []);
      if (!fallbackRows.length) return null;
      const row: any = fallbackRows[0];
      rememberRecentTrustedSession({
        device_token: row.device_token, user_name: row.user_name || '', role: row.role || 'full', login_email: row.login_email || '', property_group_uuid: row.property_group_uuid || '', phone: row.phone || '', created_at: row.created_at || '', last_seen_at: row.last_seen_at || '', expires_at: row.expires_at || ''
      });
      return {
        device_token: row.device_token, user_name: row.user_name || "", role: row.role || "full", login_email: row.login_email || "", property_group_uuid: row.property_group_uuid || "", phone: row.phone || "", created_at: row.created_at || "", last_seen_at: row.last_seen_at || "", expires_at: row.expires_at || ""
      };
    }
    const row: any = rows[0];
    rememberRecentTrustedSession({
      device_token: row.device_token, user_name: row.user_name || '', role: row.role || 'full', login_email: row.login_email || '', property_group_uuid: row.property_group_uuid || '', phone: row.phone || '', created_at: row.created_at || '', last_seen_at: row.last_seen_at || '', expires_at: row.expires_at || ''
    });
    return {
      device_token: row.device_token, user_name: row.user_name || "", role: row.role || "full", login_email: row.login_email || "", property_group_uuid: row.property_group_uuid || "", phone: row.phone || "", created_at: row.created_at || "", last_seen_at: row.last_seen_at || "", expires_at: row.expires_at || ""
    };
  } catch (err: any) {
    if (!isMissingColumnError(err)) {
      const fallbackRows = await selectSessionModern(fallbackExec).catch(() => []);
      if (fallbackRows.length) {
        const row: any = fallbackRows[0];
        rememberRecentTrustedSession({
          device_token: row.device_token, user_name: row.user_name || '', role: row.role || 'full', login_email: row.login_email || '', property_group_uuid: row.property_group_uuid || '', phone: row.phone || '', created_at: row.created_at || '', last_seen_at: row.last_seen_at || '', expires_at: row.expires_at || ''
        });
        return {
          device_token: row.device_token, user_name: row.user_name || "", role: row.role || "full", login_email: row.login_email || "", property_group_uuid: row.property_group_uuid || "", phone: row.phone || "", created_at: row.created_at || "", last_seen_at: row.last_seen_at || "", expires_at: row.expires_at || ""
        };
      }
      throw err;
    }
  }

  let rows = await selectSessionLegacy(authExec).catch(() => []);
  if (!rows.length) rows = await selectSessionLegacy(fallbackExec).catch(() => []);
  if (!rows.length) return null;
  const row: any = rows[0];
  rememberRecentTrustedSession({ device_token: row.device_token, user_name: row.user_name || '', role: 'full', created_at: row.created_at || '' });
  return { device_token: row.device_token, user_name: row.user_name || "", role: "full", login_email: "", property_group_uuid: "", phone: "", created_at: row.created_at || "", last_seen_at: "", expires_at: "" };
}

export async function touchDeviceSession(token: string): Promise<void> {
  if (!token) return;
  try {
    await executeTrustedDevicesSqlOnBoth(
      `UPDATE trusted_devices SET last_seen_at = datetime('now'), expires_at = datetime('now', '+30 days') WHERE device_token = ? AND revoked = 0`,
      [token],
    );
    const recent = readRecentTrustedSession(token);
    if (recent) {
      rememberRecentTrustedSession({
        device_token: token, user_name: recent.user_name || '', role: recent.role || 'full', login_email: recent.login_email || '', property_group_uuid: recent.property_group_uuid || '', phone: recent.phone || '', created_at: recent.created_at || '',
      });
    }
  } catch (e: any) {
    if (isMissingColumnError(e)) return;
    console.warn("[touchDeviceSession] failed:", String(e?.message || e).substring(0, 120));
  }
}

export async function handlePmProxyUsersList(params: Record<string, string>, req: Request): Promise<any> {
  try {
    await ensurePmProxyUsersTable();
    if (!PROXY_ADMIN_KEY) return { ok: false, error: "pm user admin disabled — set PROXY_ADMIN_KEY" };
    const key = params.key || req.headers.get("x-admin-key") || "";
    if (key !== PROXY_ADMIN_KEY) return { ok: false, error: "Unauthorized: invalid admin key" };
    const limit = Math.max(1, Math.min(500, parseInt(String(params.limit || "100"), 10) || 100));
    const offset = Math.max(0, parseInt(String(params.offset || "0"), 10) || 0);
    const totalResult = await sqlite.execute({ sql: `SELECT COUNT(*) AS total FROM pm_proxy_users` });
    const total = Number(rowsAsObjects(totalResult)[0]?.total || 0);
    const result = await sqlite.execute({
      sql: `SELECT user_uuid, email, full_name, phone, property_group_uuid, active, created_at, updated_at FROM pm_proxy_users ORDER BY active DESC, updated_at DESC LIMIT ? OFFSET ?`,
      args: [limit, offset],
    });
    if (!result || typeof result !== "object") return { ok: true, results: [], count: 0 };
    let rows: any[] = [];
    try { rows = rowsAsObjects(result) || []; } catch { rows = []; }
    if (!Array.isArray(rows)) rows = [];
    return { ok: true, results: rows, count: rows.length, total, limit, offset };
  } catch (err: any) {
    const errMsg = err?.message || String(err) || "Unknown error";
    console.error("[PM_PROXY_USERS_ERROR]", errMsg, err);
    return { ok: false, status: 500, error: `Failed to fetch PM users: ${errMsg}` };
  }
}

export async function handlePmProxyUserUpsert(req: Request): Promise<any> {
  try {
    await ensurePmProxyUsersTable();
    if (!PROXY_ADMIN_KEY) return { ok: false, error: "pm user admin disabled — set PROXY_ADMIN_KEY" };
    let body: any = {};
    try { body = await req.json(); } catch { body = {}; }
    const key = getBodyField(body, "key", "admin_key");
    if (key !== PROXY_ADMIN_KEY) return { ok: false, error: "Unauthorized: invalid admin key" };
    const email = normalizeEmail(getBodyField(body, "email"));
    const propertyGroupUuid = getBodyField(body, "property_group_uuid", "scope_uuid");
    if (!email || !propertyGroupUuid) return { ok: false, status: 400, error: "Missing email or property_group_uuid" };
    const userUuid = getBodyField(body, "user_uuid") || generateUuid();
    const fullName = getBodyField(body, "full_name", "name");
    const phone = getBodyField(body, "phone");
    const active = String(getBodyField(body, "active") || "1") === "0" ? 0 : 1;

    await sqlite.execute({
      sql: `INSERT INTO pm_proxy_users (id, user_uuid, email, full_name, phone, property_group_uuid, roles, is_active, active, raw_json, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'), datetime('now'))
            ON CONFLICT(email) DO UPDATE SET id=COALESCE(NULLIF(pm_proxy_users.id, ''), excluded.id), user_uuid=COALESCE(NULLIF(pm_proxy_users.user_uuid, ''), excluded.user_uuid), full_name=excluded.full_name, phone=excluded.phone, property_group_uuid=excluded.property_group_uuid, roles=COALESCE(pm_proxy_users.roles, excluded.roles), is_active=excluded.is_active, active=excluded.active, raw_json=COALESCE(pm_proxy_users.raw_json, excluded.raw_json), updated_at=datetime('now')`,
      args: [userUuid, userUuid, email, fullName, phone, propertyGroupUuid, "[]", active, active, "{}"],
    });

    const confirmResult = await sqlite.execute({
      sql: `SELECT user_uuid FROM pm_proxy_users WHERE lower(email) = lower(?) LIMIT 1`, args: [email],
    });
    const confirmedUuid = String(rowsAsObjects(confirmResult || { rows: [], columns: [] })[0]?.user_uuid || userUuid);
    return { ok: true, user_uuid: confirmedUuid, email, full_name: fullName, phone, property_group_uuid: propertyGroupUuid, active };
  } catch (err: any) {
    const errMsg = err?.message || String(err) || "Unknown error";
    console.error("[PM_PROXY_UPSERT_ERROR]", errMsg, err);
    return { ok: false, status: 500, error: `Failed to save PM user: ${errMsg}` };
  }
}

export async function handlePmProxyUserDelete(req: Request): Promise<any> {
  try {
    await ensurePmProxyUsersTable();
    if (!PROXY_ADMIN_KEY) return { ok: false, error: "pm user admin disabled — set PROXY_ADMIN_KEY" };
    let body: any = {};
    try { body = await req.json(); } catch { body = {}; }
    const key = getBodyField(body, "key", "admin_key");
    if (key !== PROXY_ADMIN_KEY) return { ok: false, error: "Unauthorized: invalid admin key" };
    const userUuid = getBodyField(body, "user_uuid");
    const email = normalizeEmail(getBodyField(body, "email"));
    if (!userUuid && !email) return { ok: false, status: 400, error: "Missing user_uuid or email" };
    const sql = userUuid ? `DELETE FROM pm_proxy_users WHERE user_uuid = ?` : `DELETE FROM pm_proxy_users WHERE lower(email) = lower(?)`;
    const arg = userUuid || email;
    const result = await sqlite.execute({ sql, args: [arg] });
    return { ok: true, deleted: (result.rowsAffected || 0) > 0, rowsAffected: result.rowsAffected || 0 };
  } catch (err: any) {
    const errMsg = err?.message || String(err) || "Unknown error";
    console.error("[PM_PROXY_DELETE_ERROR]", errMsg, err);
    return { ok: false, status: 500, error: `Failed to delete PM user: ${errMsg}` };
  }
}

export async function handleTrustedDeviceList(params: Record<string, string>, req: Request): Promise<any> {
  await ensureTrustedDevicesTable();
  if (!PROXY_ADMIN_KEY) return { ok: false, error: "trusted device admin is disabled — set PROXY_ADMIN_KEY" };
  const key = params.key || req.headers.get("x-admin-key") || "";
  if (key !== PROXY_ADMIN_KEY) return { ok: false, error: "Unauthorized: invalid admin key" };
  const limit = Math.max(1, Math.min(500, parseInt(String(params.limit || "100"), 10) || 100));
  const offset = Math.max(0, parseInt(String(params.offset || "0"), 10) || 0);
  const totalResult = await executeTrustedDevicesSql(`SELECT COUNT(*) AS total FROM trusted_devices WHERE revoked = 0`);
  const total = Number(rowsAsObjects(totalResult)[0]?.total || 0);
  const result = await executeTrustedDevicesSql(`SELECT device_token, user_name, role, login_email, created_at, last_seen_at, expires_at FROM trusted_devices WHERE revoked = 0 ORDER BY created_at DESC LIMIT ? OFFSET ?`, [limit, offset]);
  const devices = rowsAsObjects(result).map((d: any) => ({ device_token: d.device_token, token_preview: String(d.device_token || "").slice(0, 8) + "...", user_name: d.user_name, role: d.role || "full", login_email: d.login_email || "", created_at: d.created_at, last_seen_at: d.last_seen_at || "", expires_at: d.expires_at || "" }));
  return { ok: true, results: devices, count: devices.length, total, limit, offset };
}

export async function handleTrustedDeviceRevoke(params: Record<string, string>, req: Request): Promise<any> {
  await ensureTrustedDevicesTable();
  if (!PROXY_ADMIN_KEY) return { ok: false, error: "trusted device admin is disabled — set PROXY_ADMIN_KEY" };
  let body: any = {};
  try { body = await req.json(); } catch { body = {}; }
  const key = params.key || body.key || req.headers.get("x-admin-key") || "";
  if (key !== PROXY_ADMIN_KEY) return { ok: false, error: "Unauthorized: invalid admin key" };
  const token = getBodyField(body, "token", "device_token", "deviceToken") || getBodyField(params, "token", "device_token", "deviceToken");
  if (!token) return { ok: false, error: "Missing token/device_token" };
  const result = await executeTrustedDevicesSql(`UPDATE trusted_devices SET revoked = 1 WHERE device_token = ?`, [token]);
  return { ok: true, revoked: (result.rowsAffected || 0) > 0, rowsAffected: result.rowsAffected || 0 };
}

export async function handleVerifyRole(req: Request): Promise<any> {
  let body: any = {};
  try { body = await req.json(); } catch { body = {}; }
  const password = getBodyField(body, "password", "pass");
  if (!password) return { ok: false, status: 400, error: "Password is required" };
  const guiAdmin = Deno.env.get("GUI_ADMIN") || "";
  const guiGm = Deno.env.get("GUI_GM") || "";
  const guiVendors = Deno.env.get("GUI_VENDORS") || "";
  if (!guiAdmin && !guiGm && !guiVendors) return { ok: false, status: 500, error: "GUI role passwords are not configured on the server. Set GUI_ADMIN, GUI_GM, and/or GUI_VENDORS in Val.town environment variables." };
  let matchedRole: string | null = null;
  if (guiAdmin && password === guiAdmin) matchedRole = "full";
  else if (guiGm && password === guiGm) matchedRole = "manager";
  else if (guiVendors && password === guiVendors) matchedRole = "vendors";
  if (!matchedRole) return { ok: false, status: 401, error: "Invalid password" };
  const userName = getBodyField(body, "user_name", "userName", "user") || "password-session";
  try {
    await ensureTrustedDevicesTable();
    const token = generateUuid();
    await insertTrustedDeviceSession({ token, userName, role: matchedRole });
    return { ok: true, role: matchedRole, token };
  } catch (e: any) {
    console.warn("[handleVerifyRole] DB write failed, attempting signed token fallback:", String(e?.message || e).substring(0, 160));
    const signedToken = await mintSignedToken(matchedRole, userName);
    if (signedToken) return { ok: true, role: matchedRole, token: signedToken };
    return { ok: false, status: 503, error: "Password verified but session token could not be created. Check proxy database write access or set FRONTEND_PROXY_SECRET to enable stateless auth." };
  }
}