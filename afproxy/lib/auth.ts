// ============================================================================
// lib/auth.ts — Magic link token signing, verification, and generation.
//
// Magic links are HMAC-SHA-256 signed URL tokens that give a technician
// one-tap access to a mobile portal for contacting a tenant via SMS.
//
// Token structure:  base64url(JSON payload) + "." + base64url(HMAC sig)
// Signing key:      MAGIC_LINK_SECRET env var (stored in Val Town settings)
// Expiry:           24 hours from generation (embedded in payload.expires_at)
// Single-use:       Enforced by magic_link_tokens.used flag in Turso
//
// The Web Crypto API (crypto.subtle) is available natively in Deno/Val Town —
// no external JWT library is needed or imported.
// ============================================================================

import { MAGIC_LINK_SECRET, PROXY_BASE_URL } from "../config.ts";
import { rowsAsObjects, sqlite } from "../db.ts";

const SHORT_LINK_CHARS =
  "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789";
const SHORT_LINK_ATTEMPTS = 5;

// ── Base64url encode / decode helpers ─────────────────────────────────────────
function _b64url(buf: ArrayBufferLike): string {
  let bin = "";
  new Uint8Array(buf).forEach((b) => (bin += String.fromCharCode(b)));
  return btoa(bin).replace(/\+/g, "-").replace(/\//g, "_").replace(/=/g, "");
}

function _fromb64url(s: string): Uint8Array {
  const pad = s.replace(/-/g, "+").replace(/_/g, "/");
  const padded = pad + "=".repeat((4 - (pad.length % 4)) % 4);
  return Uint8Array.from(atob(padded), (c) => c.charCodeAt(0));
}

function _normalizeProxyBaseUrl(): string {
  return String(PROXY_BASE_URL || "").replace(/\/+$/, "");
}

function _isAllowedPublicBaseUrl(base: string): boolean {
  if (!base) return false;
  try {
    const u = new URL(base);
    if (u.protocol !== "https:") return false;
    const host = String(u.hostname || "").toLowerCase();
    // Runtime invariant: links must resolve from a real HTTPS host.
    // Domain ownership policy is deployment-specific and should be enforced
    // via PROXY_BASE_URL configuration, not hard-blocked here.
    return !!host;
  } catch {
    return false;
  }
}

function _shortCode(length = 6): string {
  const bytes = crypto.getRandomValues(new Uint8Array(length));
  let out = "";
  for (let i = 0; i < bytes.length; i++) {
    out += SHORT_LINK_CHARS[bytes[i] % SHORT_LINK_CHARS.length];
  }
  return out;
}

export function buildShortLinkUrl(code: string): string {
  const base = _normalizeProxyBaseUrl();
  return base ? `${base}/s/${encodeURIComponent(code)}` : "";
}

export function isDispatchShortLink(
  url: string,
  code?: string | null,
): boolean {
  const base = _normalizeProxyBaseUrl();
  const normalizedUrl = String(url || "").trim().replace(/\/+$/, "");
  if (!normalizedUrl || !base || !_isAllowedPublicBaseUrl(base)) return false;

  if (code) {
    return normalizedUrl === buildShortLinkUrl(String(code || "").trim());
  }

  return normalizedUrl.startsWith(`${base}/s/`);
}

export async function createShortLink(
  fullUrl: string,
  token: string,
  expiresAt?: string,
): Promise<{ code: string; shortUrl: string } | null> {
  const base = _normalizeProxyBaseUrl();
  if (!base || !fullUrl || !token) return null;

  for (let attempt = 0; attempt < SHORT_LINK_ATTEMPTS; attempt++) {
    const code = _shortCode(6);
    try {
      await sqlite.execute({
        sql: `INSERT INTO short_links (code, full_url, token, expires_at)
               VALUES (?, ?, ?, ?)`,
        args: [code, fullUrl, token, expiresAt || null],
      });
      return { code, shortUrl: buildShortLinkUrl(code) };
    } catch {
      // Collision or unavailable table — retry collisions, fall through on final failure.
    }
  }

  // Fallback for partial migrations: store the short code on magic_link_tokens.
  // resolveShortLink() already knows how to expand via this column.
  for (let attempt = 0; attempt < SHORT_LINK_ATTEMPTS; attempt++) {
    const code = _shortCode(6);
    try {
      await sqlite.execute({
        sql: `UPDATE magic_link_tokens
                 SET short_code = ?
               WHERE token = ?`,
        args: [code, token],
      });
      return { code, shortUrl: buildShortLinkUrl(code) };
    } catch {
      // Continue trying alternate codes.
    }
  }

  return null;
}

export async function resolveShortLink(
  code: string,
): Promise<{ ok: boolean; fullUrl?: string; token?: string; reason?: string }> {
  try {
    const codeNorm = String(code || "").trim();
    if (!codeNorm) return { ok: false, reason: "not_found" };

    const rows = rowsAsObjects(
      await sqlite.execute({
        sql: `SELECT code, full_url, token, expires_at
             FROM short_links
            WHERE lower(code) = lower(?)
            LIMIT 1`,
        args: [codeNorm],
      }),
    );
    const row = rows[0] as any;
    if (!row) {
      // Fallback path for partial migrations or stale short_links rows.
      const tokenRows = rowsAsObjects(
        await sqlite.execute({
          sql: `SELECT token, expires_at
               FROM magic_link_tokens
              WHERE lower(short_code) = lower(?)
              LIMIT 1`,
          args: [codeNorm],
        }),
      );
      const tok = tokenRows[0] as any;
      if (!tok || !tok.token) return { ok: false, reason: "not_found" };
      if (tok.expires_at && new Date(String(tok.expires_at)) < new Date()) {
        return { ok: false, reason: "expired" };
      }
      const base = _normalizeProxyBaseUrl();
      const fullUrl = base
        ? `${base}?action=portal&token=${encodeURIComponent(String(tok.token))}`
        : "";
      return {
        ok: !!fullUrl,
        fullUrl,
        token: String(tok.token || ""),
        reason: fullUrl ? undefined : "lookup_failed",
      };
    }

    if (row.expires_at && new Date(String(row.expires_at)) < new Date()) {
      return { ok: false, reason: "expired" };
    }

    await sqlite.execute({
      sql: `UPDATE short_links
               SET clicks = COALESCE(clicks, 0) + 1,
                   last_hit_at = datetime('now')
                  WHERE lower(code) = lower(?)`,
      args: [codeNorm],
    }).catch(() => {});

    return {
      ok: true,
      fullUrl: String(row.full_url || ""),
      token: String(row.token || ""),
    };
  } catch {
    return { ok: false, reason: "lookup_failed" };
  }
}

export type MagicLinkSession = {
  token: string;
  longUrl: string;
  url: string;
  expiresAt: string;
  shortCode: string | null;
};

// ── signMagicToken ─────────────────────────────────────────────────────────────
// Serialise payload as JSON, HMAC-SHA-256 sign it, return base64url(data).base64url(sig).
export async function signMagicToken(
  payload: Record<string, any>,
): Promise<string> {
  const enc = new TextEncoder().encode(JSON.stringify(payload));
  const key = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(MAGIC_LINK_SECRET),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"],
  );
  const sig = await crypto.subtle.sign("HMAC", key, enc);
  return `${_b64url(enc.buffer)}.${_b64url(sig)}`;
}

// ── verifyMagicToken ───────────────────────────────────────────────────────────
// Verify the HMAC signature and check the expiry timestamp.
// Returns the decoded payload object, or null if invalid/expired/tampered.
export async function verifyMagicToken(
  token: string,
): Promise<Record<string, any> | null> {
  try {
    const [dataB64, sigB64] = token.split(".");
    if (!dataB64 || !sigB64) return null;

    const dataBytes = new Uint8Array(_fromb64url(dataB64));
    const sigBytes = new Uint8Array(_fromb64url(sigB64));

    const key = await crypto.subtle.importKey(
      "raw",
      new TextEncoder().encode(MAGIC_LINK_SECRET),
      { name: "HMAC", hash: "SHA-256" },
      false,
      ["verify"],
    );

    const valid = await crypto.subtle.verify("HMAC", key, sigBytes, dataBytes);
    if (!valid) return null;

    const payload = JSON.parse(new TextDecoder().decode(dataBytes));

    // Reject expired tokens
    if (payload.expires_at && new Date(payload.expires_at) < new Date()) {
      return null;
    }

    return payload;
  } catch {
    return null; // any parse/decode error = invalid token
  }
}

// ── generateMagicLink ─────────────────────────────────────────────────────────
// Build a signed 24-hour magic link for a technician's portal session.
// Stores the token in magic_link_tokens (Turso) so single-use can be enforced.
// Returns the full URL or a placeholder string if PROXY_BASE_URL is not set.
export async function createMagicLinkSession(
  woId: string,
  techId: string,
  techName: string,
  tenantPhone: string,
  tenantName: string,
  propertyAddress: string,
  extraPayload: Record<string, any> = {},
): Promise<MagicLinkSession> {
  const expiresAt = new Date(Date.now() + 24 * 3600 * 1000).toISOString();
  const langPref = String(extraPayload.lang_pref || "en").trim() || "en";

  const payload: Record<string, any> = {
    wo_id: woId,
    tech_id: techId,
    tech_name: techName,
    tenant_phone: tenantPhone,
    tenant_name: tenantName,
    property_address: propertyAddress,
    expires_at: expiresAt,
    ...extraPayload,
  };

  const token = await signMagicToken(payload);
  const base = _normalizeProxyBaseUrl();
  const longUrl = base ? `${base}?action=portal&token=${token}` : "";

  // Persist token so we can enforce single-use and lookup context on POST
  try {
    await sqlite.execute({
      sql: `INSERT OR REPLACE INTO magic_link_tokens
               (token, wo_id, tech_id, tech_name,
                tenant_phone, tenant_name, property_address, expires_at)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
      args: [
        token,
        woId,
        techId,
        techName,
        tenantPhone,
        tenantName,
        propertyAddress,
        expiresAt,
      ],
    });

    // Best-effort metadata backfill for partially migrated installs.
    await sqlite.execute({
      sql: `UPDATE magic_link_tokens
               SET lang_pref = ?,
                   meta_json = ?,
                   last_action = 'generated',
                   last_action_at = datetime('now')
             WHERE token = ?`,
      args: [langPref, JSON.stringify(extraPayload || {}), token],
    }).catch(() => {});
  } catch (err: unknown) {
    // CRITICAL: if DB write fails, the single-use enforcement is broken.
    // Fail loudly so the caller knows the token is unsafe to return.
    const errMsg = err instanceof Error ? err.message : String(err);
    console.error(`createMagicLinkSession DB write FAILED for WO ${woId}:`, errMsg.substring(0, 220));
    throw new Error(`Failed to persist magic link token — DB write error.`);
  }

  if (!base || !_isAllowedPublicBaseUrl(base)) {
    return {
      token,
      longUrl:
        "[Set PROXY_BASE_URL to an HTTPS company-owned domain to enable magic links]",
      url:
        "[Set PROXY_BASE_URL to an HTTPS company-owned domain to enable magic links]",
      expiresAt,
      shortCode: null,
    };
  }

  let shortCode: string | null = null;
  let url = "";
  const shortLink = await createShortLink(longUrl, token, expiresAt);
  if (shortLink && isDispatchShortLink(shortLink.shortUrl, shortLink.code)) {
    shortCode = shortLink.code;
    url = shortLink.shortUrl;
    try {
      await sqlite.execute({
        sql: `UPDATE magic_link_tokens SET short_code = ? WHERE token = ?`,
        args: [shortCode, token],
      });
    } catch {
      // Non-fatal — returning the short link is still safe.
    }
  } else {
    console.log(
      `createMagicLinkSession: short link unavailable for WO ${woId}`,
    );
  }

  return { token, longUrl, url, expiresAt, shortCode };
}

export async function generateMagicLink(
  woId: string,
  techId: string,
  techName: string,
  tenantPhone: string,
  tenantName: string,
  propertyAddress: string,
  extraPayload: Record<string, any> = {},
): Promise<string> {
  const session = await createMagicLinkSession(
    woId,
    techId,
    techName,
    tenantPhone,
    tenantName,
    propertyAddress,
    extraPayload,
  );
  return session.url;
}

// ── lookupMagicToken ──────────────────────────────────────────────────────────
// Fetch the stored token row from Turso — used to check single-use status.
// Returns null if the token is not found in the DB.
export async function lookupMagicToken(
  token: string,
): Promise<
  {
    token: string;
    used: number;
    used_template: string | null;
    wo_id: string;
    tech_id: string;
    tech_name: string;
    tenant_phone: string;
    tenant_name: string;
    property_address?: string;
    short_code?: string | null;
    lang_pref?: string | null;
    scheduled_date?: string | null;
    scheduled_window?: string | null;
    stop_auto?: number;
    exempt_until?: string | null;
    last_action?: string | null;
    last_action_at?: string | null;
    portal_opened?: number;
    portal_opened_at?: string | null;
    meta_json?: string | null;
  } | null
> {
  try {
    const result = await sqlite.execute({
      sql: `SELECT token, used, used_template, wo_id, tech_id,
                    tech_name, tenant_phone, tenant_name, property_address,
                    short_code, lang_pref, scheduled_date, scheduled_window,
                    stop_auto, exempt_until, last_action, last_action_at,
                    portal_opened, portal_opened_at, meta_json
             FROM magic_link_tokens
             WHERE token = ?`,
      args: [token],
    });
    const rows = rowsAsObjects(result);
    return rows.length > 0 ? (rows[0] as any) : null;
  } catch {
    return null;
  }
}

// ── markTokenUsed ─────────────────────────────────────────────────────────────
// Mark a magic link token as consumed after a tenant SMS is sent.
// template: human-readable label e.g. "I'm On My Way"
export async function markTokenUsed(
  token: string,
  template: string,
): Promise<void> {
  try {
    await sqlite.execute({
      sql: `UPDATE magic_link_tokens
             SET used = 1,
                 used_at = datetime('now'),
                 used_template = ?,
                 last_action = 'tenant_sms',
                 last_action_at = datetime('now')
             WHERE token = ?`,
      args: [template, token],
    });
  } catch (_) { /* non-fatal */ }
}