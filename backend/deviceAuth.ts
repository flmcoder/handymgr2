import { queryClient } from './db';

type PmProxyUser = {
  user_uuid: string;
  email: string;
  full_name: string;
  phone: string;
  property_group_uuid: string;
  scopes: string[];
  primary_scope_uuid: string;
  active: boolean;
};

type ScopeOption = {
  property_group_uuid: string;
  property_group_name: string;
};

const DEVICE_SETUP_PIN = String(process.env.DEVICE_SETUP_PIN || '').trim();
const OTP_ALLOWED_DOMAIN_ENV = String(process.env.OTP_ALLOWED_DOMAIN || 'flraz.com').replace(/^@/, '').toLowerCase();
const OTP_TTL_MINUTES_ENV = Math.max(3, Number(process.env.DEVICE_OTP_TTL_MINUTES || '10') || 10);
const GUI_ADMIN = String(process.env.GUI_ADMIN || '').trim();
const GUI_GM = String(process.env.GUI_GM || '').trim();
const GUI_VENDORS = String(process.env.GUI_VENDORS || '').trim();
const RC_SERVER_URL = String(process.env.RC_SERVER_URL || '').trim();
const RC_TOKEN = String(process.env.RC_ACCESS_TOKEN || '').trim();
const RC_FROM = String(process.env.RC_FROM_NUMBER || '').trim();
const SESSION_SIGN_KEY_RAW = String(process.env.FRONTEND_PROXY_SECRET || '').trim();

const recentTrustedSessions = new Map<string, any>();
const memoryOtpStore = new Map<string, { code: string; used: boolean; expiresAtMs: number; userName: string; scopeUuid: string }>();
let authTablesReady = false;

function getBodyField(body: any, ...keys: string[]): string {
  if (!body || typeof body !== 'object') return '';
  for (const key of keys) {
    const value = body[key];
    if (value !== undefined && value !== null && String(value).trim()) {
      return String(value).trim();
    }
  }
  return '';
}

function normalizeEmail(rawEmail: string): string {
  return String(rawEmail || '').trim().toLowerCase();
}

function normalizeScopeUuid(rawScope: string): string {
  const value = String(rawScope || '').trim().toLowerCase();
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value)
    ? value
    : '';
}

function stripPhoneExtension(raw: string): string {
  return String(raw || '').replace(/(?:ext\.?|extension|x|#)\s*\d+\s*$/i, '').trim();
}

function last10Digits(raw: string): string {
  const digits = String(raw || '').replace(/\D/g, '');
  return digits.slice(-10);
}

function phoneMatchKeys(rawPhone: string): string[] {
  const digits = stripPhoneExtension(rawPhone).replace(/\D/g, '');
  if (!digits) return [];
  const keys = [digits];
  if (digits.length === 11 && digits.startsWith('1')) keys.push(digits.slice(1));
  if (digits.length >= 10) keys.push(last10Digits(digits));
  return Array.from(new Set(keys.filter(Boolean)));
}

function normalizePhone(rawPhone: string): string {
  const digits = stripPhoneExtension(rawPhone).replace(/\D/g, '');
  if (digits.length === 10) return `+1${digits}`;
  if (digits.length === 11 && digits.startsWith('1')) return `+${digits}`;
  if (digits.length > 11) return `+${digits}`;
  return '';
}

function normalizeOrgEmail(rawEmail: string, allowedDomain: string): string {
  const email = normalizeEmail(rawEmail);
  if (!email || email.indexOf('@') === -1) return '';
  const parts = email.split('@');
  if (parts.length !== 2) return '';
  if (allowedDomain && parts[1] !== allowedDomain) return '';
  return email;
}

function deriveOtpIdentityEmail(pmUser: PmProxyUser | null, requestedIdentifier: string, allowedDomain: string): string {
  const direct = normalizeOrgEmail(pmUser?.email || requestedIdentifier, allowedDomain);
  if (direct) return direct;
  if (!allowedDomain) return '';
  const phoneKey = last10Digits(requestedIdentifier || '');
  if (!pmUser && phoneKey) return `pm-phone-${phoneKey}@${allowedDomain}`;
  const userKey = String(pmUser?.user_uuid || pmUser?.email || '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
  if (!userKey) return '';
  return `pm-${userKey}@${allowedDomain}`;
}

function generateOtpCode(): string {
  const arr = new Uint32Array(1);
  crypto.getRandomValues(arr);
  return String(arr[0] % 1000000).padStart(6, '0');
}

function generateUuid(): string {
  return crypto.randomUUID();
}

function rememberRecentTrustedSession(session: any): void {
  const token = String(session.device_token || '').trim();
  if (!token) return;
  recentTrustedSessions.set(token, {
    ...session,
    cached_at_ms: Date.now(),
  });
}

function readRecentTrustedSession(token: string): any | null {
  const entry = recentTrustedSessions.get(String(token || '').trim());
  if (!entry) return null;
  if ((Date.now() - Number(entry.cached_at_ms || 0)) > 10 * 60 * 1000) {
    recentTrustedSessions.delete(String(token || '').trim());
    return null;
  }
  return entry;
}

function otpMemoryKey(email: string, scopeUuid: string): string {
  return `${normalizeEmail(email)}::${normalizeScopeUuid(scopeUuid)}`;
}

function rememberMemoryOtp(email: string, code: string, ttlMinutes: number, userName: string, scopeUuid: string): void {
  const key = otpMemoryKey(email, scopeUuid);
  if (!key) return;
  memoryOtpStore.set(key, {
    code,
    used: false,
    expiresAtMs: Date.now() + Math.max(1, ttlMinutes) * 60_000,
    userName,
    scopeUuid,
  });
}

function readMemoryOtp(email: string, scopeUuid: string) {
  const key = otpMemoryKey(email, scopeUuid);
  const entry = memoryOtpStore.get(key);
  if (!entry) return null;
  if (entry.expiresAtMs < Date.now()) {
    memoryOtpStore.delete(key);
    return null;
  }
  return entry;
}

async function getSignKey(): Promise<CryptoKey | null> {
  if (!SESSION_SIGN_KEY_RAW) return null;
  try {
    return await crypto.subtle.importKey('raw', new TextEncoder().encode(SESSION_SIGN_KEY_RAW), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign', 'verify']);
  } catch {
    return null;
  }
}

function b64url(buf: ArrayBuffer): string {
  return Buffer.from(buf).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

function b64urlDecode(input: string): Uint8Array {
  const padded = input.replace(/-/g, '+').replace(/_/g, '/') + '=='.slice(0, (4 - (input.length % 4)) % 4);
  return new Uint8Array(Buffer.from(padded, 'base64'));
}

async function mintSignedToken(role: string, userName: string): Promise<string | null> {
  const key = await getSignKey();
  if (!key) return null;
  const payload = b64url(new TextEncoder().encode(JSON.stringify({ r: role, u: userName, iat: Math.floor(Date.now() / 1000) })).buffer);
  const sig = b64url(await crypto.subtle.sign('HMAC', key, new TextEncoder().encode(payload)));
  return `v1.${payload}.${sig}`;
}

async function verifySignedToken(token: string): Promise<{ role: string; userName: string; iat: number } | null> {
  if (!token || !token.startsWith('v1.')) return null;
  const parts = token.split('.');
  if (parts.length !== 3) return null;
  const [, payload64, sig64] = parts;
  const key = await getSignKey();
  if (!key) return null;
  try {
    const sigBytes = new Uint8Array(b64urlDecode(sig64));
    const valid = await crypto.subtle.verify('HMAC', key, sigBytes, new TextEncoder().encode(payload64));
    if (!valid) return null;
    const decoded = JSON.parse(Buffer.from(b64urlDecode(payload64)).toString('utf8'));
    if (!decoded || typeof decoded.iat !== 'number') return null;
    if ((Math.floor(Date.now() / 1000) - decoded.iat) > 2592000) return null;
    return { role: String(decoded.r || ''), userName: String(decoded.u || ''), iat: decoded.iat };
  } catch {
    return null;
  }
}

async function ensureAuthTables(): Promise<void> {
  if (authTablesReady) return;
  await queryClient.unsafe(`
    CREATE TABLE IF NOT EXISTS trusted_devices (
      device_token TEXT PRIMARY KEY,
      user_name TEXT,
      role TEXT DEFAULT 'full',
      login_email TEXT,
      property_group_uuid TEXT,
      phone TEXT,
      revoked BOOLEAN DEFAULT FALSE,
      last_seen_at TIMESTAMPTZ,
      expires_at TIMESTAMPTZ,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
    CREATE TABLE IF NOT EXISTS device_otps (
      id INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
      email TEXT NOT NULL,
      code TEXT NOT NULL,
      used BOOLEAN DEFAULT FALSE,
      expires_at TIMESTAMPTZ NOT NULL,
      user_name TEXT,
      role_hint TEXT,
      property_group_uuid TEXT,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      used_at TIMESTAMPTZ
    );
    CREATE INDEX IF NOT EXISTS device_otps_email_idx ON device_otps(email);
    CREATE TABLE IF NOT EXISTS pm_proxy_users (
      user_uuid TEXT PRIMARY KEY,
      email TEXT NOT NULL,
      full_name TEXT,
      phone TEXT,
      property_group_uuid TEXT,
      roles JSONB NOT NULL DEFAULT '[]'::jsonb,
      active BOOLEAN DEFAULT TRUE,
      raw_json JSONB NOT NULL DEFAULT '{}'::jsonb,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
    DROP INDEX IF EXISTS pm_proxy_users_email_unique;
    CREATE INDEX IF NOT EXISTS pm_proxy_users_email_idx ON pm_proxy_users(lower(email));
    CREATE INDEX IF NOT EXISTS pm_proxy_users_group_idx ON pm_proxy_users(property_group_uuid);
    CREATE TABLE IF NOT EXISTS pm_proxy_user_scopes (
      id INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
      user_uuid TEXT NOT NULL,
      property_group_uuid TEXT NOT NULL,
      is_primary BOOLEAN DEFAULT FALSE,
      active BOOLEAN DEFAULT TRUE,
      source TEXT DEFAULT 'runtime_bootstrap',
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      UNIQUE (user_uuid, property_group_uuid)
    );
    CREATE INDEX IF NOT EXISTS pm_proxy_user_scopes_user_idx ON pm_proxy_user_scopes(user_uuid);
    CREATE INDEX IF NOT EXISTS pm_proxy_user_scopes_group_idx ON pm_proxy_user_scopes(property_group_uuid);
    INSERT INTO pm_proxy_user_scopes (user_uuid, property_group_uuid, is_primary, active, source)
    SELECT user_uuid, property_group_uuid, TRUE, COALESCE(active, TRUE), 'legacy_backfill'
    FROM pm_proxy_users
    WHERE COALESCE(property_group_uuid, '') <> ''
    ON CONFLICT (user_uuid, property_group_uuid) DO UPDATE
    SET active = EXCLUDED.active, updated_at = NOW();
    CREATE TABLE IF NOT EXISTS proxy_config (
      key TEXT PRIMARY KEY,
      value TEXT,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
  `);
  authTablesReady = true;
}

async function getOtpPolicySafe(): Promise<{ enabled: boolean; allowedDomain: string; requireMembership: boolean; ttlMinutes: number }> {
  await ensureAuthTables();
  try {
    const rows = await queryClient.unsafe(`SELECT key, value FROM proxy_config WHERE key IN ('otp_enabled','otp_allowed_domain','otp_require_pm_membership','otp_ttl_minutes')`);
    const map: Record<string, string> = {};
    for (const row of rows as any[]) {
      map[String(row.key || '')] = String(row.value || '');
    }
    const rawDomain = String(map.otp_allowed_domain || '').replace(/^@/, '').toLowerCase();
    return {
      enabled: (map.otp_enabled ?? '1') !== '0',
      allowedDomain: rawDomain || OTP_ALLOWED_DOMAIN_ENV,
      requireMembership: (map.otp_require_pm_membership ?? '1') !== '0',
      ttlMinutes: Math.max(3, Number(map.otp_ttl_minutes) || OTP_TTL_MINUTES_ENV),
    };
  } catch {
    return {
      enabled: true,
      allowedDomain: OTP_ALLOWED_DOMAIN_ENV,
      requireMembership: true,
      ttlMinutes: OTP_TTL_MINUTES_ENV,
    };
  }
}

async function sendOtpSms(toPhone: string, code: string, ttlMinutes: number): Promise<void> {
  const toNumber = normalizePhone(toPhone);
  if (!toNumber) throw new Error('Missing recipient phone');
  if (!RC_SERVER_URL || !RC_TOKEN || !RC_FROM) {
    console.warn('[ringcentral] env vars missing; SMS send skipped');
    return;
  }
  const resp = await fetch(`${RC_SERVER_URL}/restapi/v1.0/account/~/extension/~/sms`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${RC_TOKEN}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      from: { phoneNumber: RC_FROM },
      to: [{ phoneNumber: toNumber }],
      text: `HandyManager verification code: ${code}. Expires in ${ttlMinutes} minutes. If you did not request this code, ignore this message.`,
    }),
  });
  if (!resp.ok) {
    const body = await resp.text().catch(() => '');
    throw new Error(`RingCentral HTTP ${resp.status}: ${body.slice(0, 180)}`);
  }
}

async function buildScopeOptions(scopeUuids: string[]): Promise<ScopeOption[]> {
  const uuids = Array.from(new Set((scopeUuids || []).map((v) => normalizeScopeUuid(v)).filter(Boolean)));
  if (!uuids.length) return [];
  const labelMap: Record<string, string> = {};
  try {
    const rows = await queryClient.unsafe(
      `SELECT uuid, name FROM appfolio_property_groups WHERE uuid = ANY($1::text[])`,
      [uuids],
    );
    for (const row of rows as any[]) {
      const uuid = normalizeScopeUuid(String(row?.uuid || ''));
      if (!uuid) continue;
      labelMap[uuid] = String(row?.name || uuid);
    }
  } catch {
    // Optional lookup table may not be hydrated yet.
  }

  return uuids.map((uuid) => ({
    property_group_uuid: uuid,
    property_group_name: labelMap[uuid] || uuid,
  }));
}

async function resolvePmProxyUsers(identifier: string): Promise<PmProxyUser[]> {
  await ensureAuthTables();
  const email = normalizeEmail(identifier);
  let rows: any[] = [];

  if (email && email.includes('@')) {
    rows = await queryClient.unsafe(
      `SELECT user_uuid, email, full_name, phone, property_group_uuid, active
       FROM pm_proxy_users
       WHERE lower(email) = lower($1) AND active = true`,
      [email],
    );
  } else {
    rows = await queryClient.unsafe(`SELECT user_uuid, email, full_name, phone, property_group_uuid, active FROM pm_proxy_users WHERE active = true`);
    const inputKeys = phoneMatchKeys(identifier);
    if (!inputKeys.length) rows = [];
    else rows = rows.filter((row) => phoneMatchKeys(String(row?.phone || '')).some((key) => inputKeys.includes(key)));
  }

  if (!rows.length) return [];

  const userUuids = Array.from(new Set(rows.map((row) => String(row?.user_uuid || '').trim()).filter(Boolean)));
  const scopesByUser = new Map<string, string[]>();

  try {
    const scopeRows = await queryClient.unsafe(
      `SELECT user_uuid, property_group_uuid
       FROM pm_proxy_user_scopes
       WHERE active = true AND user_uuid = ANY($1::text[])`,
      [userUuids],
    );
    for (const scopeRow of scopeRows as any[]) {
      const userUuid = String(scopeRow?.user_uuid || '').trim();
      const scopeUuid = normalizeScopeUuid(String(scopeRow?.property_group_uuid || ''));
      if (!userUuid || !scopeUuid) continue;
      if (!scopesByUser.has(userUuid)) scopesByUser.set(userUuid, []);
      scopesByUser.get(userUuid)!.push(scopeUuid);
    }
  } catch {
    // Scope table might not exist in older runtimes; fallback to legacy column only.
  }

  return rows.map((row: any) => {
    const userUuid = String(row?.user_uuid || '').trim();
    const legacyScope = normalizeScopeUuid(String(row?.property_group_uuid || ''));
    const joinedScopes = scopesByUser.get(userUuid) || [];
    const scopes = Array.from(new Set([legacyScope, ...joinedScopes].filter(Boolean)));
    const primaryScope = scopes.includes(legacyScope) ? legacyScope : (scopes[0] || '');
    return {
      user_uuid: userUuid,
      email: String(row?.email || ''),
      full_name: String(row?.full_name || ''),
      phone: String(row?.phone || ''),
      property_group_uuid: legacyScope,
      scopes,
      primary_scope_uuid: primaryScope,
      active: Boolean(row?.active),
    } as PmProxyUser;
  });
}

async function selectPmScopedAccount(users: PmProxyUser[], requestedScopeUuidRaw: string): Promise<{ user: PmProxyUser | null; scopeUuid: string; scopeOptions?: ScopeOption[]; error?: string }> {
  if (!users.length) return { user: null, scopeUuid: '' };

  const requestedScopeUuid = normalizeScopeUuid(requestedScopeUuidRaw);
  const allScopeUuids = users.flatMap((user) => user.scopes || []).filter(Boolean);

  if (requestedScopeUuid) {
    const scoped = users.find((user) => (user.scopes || []).includes(requestedScopeUuid));
    if (!scoped) {
      return {
        user: null,
        scopeUuid: '',
        scopeOptions: await buildScopeOptions(allScopeUuids),
        error: 'Requested property group scope is not assigned to this PM account.',
      };
    }
    return { user: scoped, scopeUuid: requestedScopeUuid };
  }

  if (users.length === 1) {
    const user = users[0];
    if (user.primary_scope_uuid) return { user, scopeUuid: user.primary_scope_uuid };
    if ((user.scopes || []).length === 1) return { user, scopeUuid: user.scopes[0] };
  }

  const distinctScopes = Array.from(new Set(allScopeUuids));
  if (distinctScopes.length === 1) {
    const onlyScope = distinctScopes[0];
    const user = users.find((candidate) => (candidate.scopes || []).includes(onlyScope)) || users[0];
    return { user, scopeUuid: onlyScope };
  }

  return {
    user: null,
    scopeUuid: '',
    scopeOptions: await buildScopeOptions(distinctScopes),
    error: 'Multiple property-group scopes are assigned to this PM account. Provide property_group_uuid to continue OTP login.',
  };
}

async function insertTrustedDeviceSession(args: { token: string; userName: string; role?: string; loginEmail?: string; propertyGroupUuid?: string; phone?: string }): Promise<void> {
  await ensureAuthTables();
  await queryClient.unsafe(
    `INSERT INTO trusted_devices (device_token, user_name, role, login_email, property_group_uuid, phone, last_seen_at, expires_at, created_at)
     VALUES ($1, $2, $3, $4, $5, $6, NOW(), NOW() + INTERVAL '30 days', NOW())
     ON CONFLICT (device_token) DO UPDATE SET
       user_name = EXCLUDED.user_name,
       role = EXCLUDED.role,
       login_email = EXCLUDED.login_email,
       property_group_uuid = EXCLUDED.property_group_uuid,
       phone = EXCLUDED.phone,
       revoked = FALSE,
       last_seen_at = NOW(),
       expires_at = NOW() + INTERVAL '30 days'`,
    [args.token, args.userName, args.role || 'full', args.loginEmail || '', args.propertyGroupUuid || '', args.phone || ''],
  );
}

export async function handleDeviceSetup(req: Request): Promise<any> {
  await ensureAuthTables();
  if (!DEVICE_SETUP_PIN) return { ok: false, status: 500, error: 'Trusted device setup disabled — missing DEVICE_SETUP_PIN env var' };
  let body: any = {};
  try { body = await req.json(); } catch { body = {}; }
  const pin = getBodyField(body, 'pin', 'setup_pin', 'setupPin');
  if (!pin || pin !== DEVICE_SETUP_PIN) return { ok: false, status: 401, error: 'Invalid setup pin' };
  const userName = getBodyField(body, 'user_name', 'userName', 'user') || 'trusted-device';
  const token = generateUuid();
  await insertTrustedDeviceSession({ token, userName, role: 'full' });
  return { ok: true, token, user_name: userName, created_at: new Date().toISOString() };
}

export async function handleDeviceOtpRequest(req: Request): Promise<any> {
  try {
    await ensureAuthTables();
    const policy = await getOtpPolicySafe();
    if (!policy.enabled) return { ok: false, status: 403, error: 'OTP login is currently disabled by administrator.' };
    let body: any = {};
    try { body = await req.json(); } catch { body = {}; }
    const requestedIdentifier = getBodyField(body, 'identifier', 'email', 'phone');
    const requestedScopeUuid = getBodyField(body, 'property_group_uuid', 'scope_uuid', 'propertyGroupUuid');
    const pmUsers = await resolvePmProxyUsers(requestedIdentifier);
    if (policy.requireMembership && !pmUsers.length) {
      return { ok: false, status: 403, error: 'No active PM account found for that identifier. Please verify the identifier you registered with, or contact your administrator.' };
    }
    const selection = await selectPmScopedAccount(pmUsers, requestedScopeUuid);
    const pmUser = selection.user;
    const scopeUuid = selection.scopeUuid;
    if (policy.requireMembership && !pmUser) {
      return {
        ok: false,
        status: 409,
        error: selection.error || 'No PM scope is available for this account.',
        scope_options: selection.scopeOptions || [],
      };
    }
    const email = deriveOtpIdentityEmail(pmUser, requestedIdentifier, policy.allowedDomain);
    const fallbackPhone = normalizePhone(requestedIdentifier);
    const smsPhone = normalizePhone(pmUser?.phone || '') || (!policy.requireMembership ? fallbackPhone : '');
    if (!email || !smsPhone) {
      return { ok: false, status: 403, error: 'No phone number on file for this PM account. Please ask your administrator to add a phone number before using OTP login.' };
    }
    const userName = getBodyField(body, 'user_name', 'userName', 'user') || 'trusted-device';
    const code = generateOtpCode();
    try {
      await queryClient.unsafe(
        `INSERT INTO device_otps (email, code, used, expires_at, user_name, role_hint, property_group_uuid)
         VALUES ($1, $2, FALSE, NOW() + ($3 * INTERVAL '1 minute'), $4, 'pm_readonly', $5)`,
        [email, code, policy.ttlMinutes, userName, String(scopeUuid || pmUser?.primary_scope_uuid || '')],
      );
    } catch {
      rememberMemoryOtp(email, code, policy.ttlMinutes, userName, String(scopeUuid || pmUser?.primary_scope_uuid || ''));
    }
    await sendOtpSms(smsPhone, code, policy.ttlMinutes);
    return {
      ok: true,
      message: pmUser ? `OTP sent to phone on file for ${pmUser.full_name || 'PM User'}.` : 'OTP sent to the provided phone number.',
      phone_hint: smsPhone.length > 7 ? smsPhone.slice(0, 5) + '***' + smsPhone.slice(-2) : '***',
      property_group_uuid: String(scopeUuid || pmUser?.primary_scope_uuid || ''),
      expires_in_minutes: policy.ttlMinutes,
    };
  } catch (err: any) {
    console.error('[DEVICE_OTP_REQUEST_ERROR]', err?.message || String(err), err);
    return { ok: false, status: 500, error: 'OTP request failed internally', message: String(err?.message || err) };
  }
}

export async function handleDeviceOtpVerify(req: Request): Promise<any> {
  await ensureAuthTables();
  const policy = await getOtpPolicySafe();
  if (!policy.enabled) return { ok: false, status: 403, error: 'OTP login is currently disabled by administrator.' };
  let body: any = {};
  try { body = await req.json(); } catch { body = {}; }
  const requestedIdentifier = getBodyField(body, 'identifier', 'email', 'phone');
  const requestedScopeUuid = getBodyField(body, 'property_group_uuid', 'scope_uuid', 'propertyGroupUuid');
  const pmUsers = await resolvePmProxyUsers(requestedIdentifier);
  const selection = await selectPmScopedAccount(pmUsers, requestedScopeUuid);
  const pmUser = selection.user;
  const requestedScope = selection.scopeUuid;
  if (policy.requireMembership && !pmUsers.length) return { ok: false, status: 403, error: 'No active PM account found for that identifier.' };
  if (policy.requireMembership && !pmUser) {
    return {
      ok: false,
      status: 409,
      error: selection.error || 'No PM scope is available for this account.',
      scope_options: selection.scopeOptions || [],
    };
  }
  const email = deriveOtpIdentityEmail(pmUser, requestedIdentifier, policy.allowedDomain);
  const code = getBodyField(body, 'code', 'otp');
  if (!email || !code) return { ok: false, status: 403, error: 'No active PM account found for that identifier.' };

  let userNameFromOtp = '';
  let scopeUuidFromOtp = '';
  const rows = await queryClient.unsafe(
    `SELECT id, code, expires_at, used, user_name, property_group_uuid
     FROM device_otps
     WHERE email = $1
       AND ($2 = '' OR coalesce(property_group_uuid, '') = $2)
     ORDER BY id DESC LIMIT 1`,
    [email, requestedScope],
  ).catch(() => [] as any[]);
  if ((rows as any[]).length) {
    const row: any = rows[0];
    if (row.used) return { ok: false, status: 401, error: 'OTP code already used' };
    if (new Date(String(row.expires_at || '')).getTime() < Date.now()) return { ok: false, status: 401, error: 'OTP code expired' };
    if (String(row.code || '') !== String(code).trim()) return { ok: false, status: 401, error: 'Invalid OTP code' };
    userNameFromOtp = String(row.user_name || '');
    scopeUuidFromOtp = String(row.property_group_uuid || '');
    await queryClient.unsafe(`UPDATE device_otps SET used = TRUE, used_at = NOW() WHERE id = $1`, [row.id]).catch(() => {});
  } else {
    const memoryRow = readMemoryOtp(email, requestedScope);
    if (!memoryRow) return { ok: false, status: 401, error: 'No OTP request found for this email' };
    if (memoryRow.used) return { ok: false, status: 401, error: 'OTP code already used' };
    if (memoryRow.expiresAtMs < Date.now()) return { ok: false, status: 401, error: 'OTP code expired' };
    if (String(memoryRow.code || '') !== String(code).trim()) return { ok: false, status: 401, error: 'Invalid OTP code' };
    memoryRow.used = true;
    userNameFromOtp = memoryRow.userName;
    scopeUuidFromOtp = memoryRow.scopeUuid;
    memoryOtpStore.set(otpMemoryKey(email, scopeUuidFromOtp), memoryRow);
  }

  const userName = getBodyField(body, 'user_name', 'userName', 'user') || pmUser?.full_name || userNameFromOtp || email;
  const role = 'pm_readonly';
  const scopeUuid = String(requestedScope || pmUser?.primary_scope_uuid || scopeUuidFromOtp || '');
  const phone = String(pmUser?.phone || '');
  let token = generateUuid();
  try {
    await insertTrustedDeviceSession({ token, userName, role, loginEmail: email, propertyGroupUuid: scopeUuid, phone });
  } catch (err) {
    const fallback = await mintSignedToken(role, userName);
    if (!fallback) {
      return { ok: false, status: 503, error: 'OTP verified but session token could not be created. Configure FRONTEND_PROXY_SECRET or restore database access.' };
    }
    token = fallback;
  }
  return { ok: true, token, user_name: userName, email, role, property_group_uuid: scopeUuid, phone, created_at: new Date().toISOString() };
}

export async function getTrustedDeviceSession(token: string): Promise<any | null> {
  if (!token) return null;
  if (token.startsWith('v1.')) {
    const decoded = await verifySignedToken(token);
    if (!decoded) return null;
    return {
      device_token: token,
      user_name: decoded.userName || 'password-session',
      role: decoded.role,
      login_email: '',
      property_group_uuid: '',
      phone: '',
      created_at: new Date(decoded.iat * 1000).toISOString(),
      last_seen_at: new Date().toISOString(),
      expires_at: new Date((decoded.iat + 2592000) * 1000).toISOString(),
    };
  }
  const recent = readRecentTrustedSession(token);
  if (recent) return recent;
  await ensureAuthTables();
  const rows = await queryClient.unsafe(
    `SELECT device_token, user_name, role, login_email, property_group_uuid, phone, created_at, last_seen_at, expires_at
     FROM trusted_devices
     WHERE device_token = $1 AND coalesce(revoked, false) = false AND (expires_at IS NULL OR expires_at > NOW())
     LIMIT 1`,
    [token],
  );
  if (!(rows as any[]).length) return null;
  const row: any = (rows as any[])[0];
  await queryClient.unsafe(`UPDATE trusted_devices SET last_seen_at = NOW(), expires_at = NOW() + INTERVAL '30 days' WHERE device_token = $1`, [token]).catch(() => {});
  const session = {
    device_token: String(row.device_token || ''),
    user_name: String(row.user_name || ''),
    role: String(row.role || 'full') || 'full',
    login_email: String(row.login_email || ''),
    property_group_uuid: String(row.property_group_uuid || ''),
    phone: String(row.phone || ''),
    created_at: row.created_at ? new Date(row.created_at).toISOString() : '',
    last_seen_at: new Date().toISOString(),
    expires_at: row.expires_at ? new Date(row.expires_at).toISOString() : '',
  };
  rememberRecentTrustedSession(session);
  return session;
}

export async function handleVerifyRole(req: Request): Promise<any> {
  let body: any = {};
  try { body = await req.json(); } catch { body = {}; }
  const password = getBodyField(body, 'password', 'pass');
  if (!password) return { ok: false, status: 400, error: 'Password is required' };
  if (!GUI_ADMIN && !GUI_GM && !GUI_VENDORS) {
    return { ok: false, status: 500, error: 'GUI role passwords are not configured on the server. Set GUI_ADMIN, GUI_GM, and/or GUI_VENDORS in environment variables.' };
  }
  let matchedRole: string | null = null;
  if (GUI_ADMIN && password === GUI_ADMIN) matchedRole = 'full';
  else if (GUI_GM && password === GUI_GM) matchedRole = 'manager';
  else if (GUI_VENDORS && password === GUI_VENDORS) matchedRole = 'vendors';
  if (!matchedRole) return { ok: false, status: 401, error: 'Invalid password' };
  const userName = getBodyField(body, 'user_name', 'userName', 'user') || 'password-session';
  try {
    const token = generateUuid();
    await insertTrustedDeviceSession({ token, userName, role: matchedRole });
    return { ok: true, role: matchedRole, token };
  } catch {
    const signedToken = await mintSignedToken(matchedRole, userName);
    if (signedToken) return { ok: true, role: matchedRole, token: signedToken };
    return { ok: false, status: 503, error: 'Password verified but session token could not be created. Check database write access or set FRONTEND_PROXY_SECRET.' };
  }
}