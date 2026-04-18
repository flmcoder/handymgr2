// ============================================================================
// lib/ringcentral.ts — RingCentral JWT auth + SMS dispatcher.
//
// Auth flow:
//   1. POST to /restapi/oauth/token with grant_type=jwt-bearer + assertion=JWT
//   2. Receive a short-lived access_token (expires_in ~7199 s)
//   3. Use Bearer {access_token} on every SMS POST
//
// A new access token is fetched per sendSMS call. The token lifetime (2 hrs)
// is long enough that for low-volume cron use this is acceptable. If SMS
// volume increases significantly, add an in-memory token cache here with
// a 110-minute TTL.
//
// The RC_JWT env var is a personal JWT credential created in the RingCentral
// Developer Console, restricted to this application's Client ID only.
// Password grant type is no longer supported by RingCentral — JWT is required.
// ============================================================================

import {
  RC_CLIENT_ID,
  RC_CLIENT_SECRET,
  RC_FROM_NUMBER,
  RC_JWT,
} from "../config.ts";
import { fetchWithTimeout } from "./fetchUtils.ts";

// Module-level token cache to avoid OAuth exchange on every SMS
let _rcTokenCache: string | null = null;
let _rcTokenExpiresAt = 0;

// ── getRCAccessToken ──────────────────────────────────────────────────────────
// Exchange the stored JWT credential for a short-lived OAuth access token.
// Returns cached token if still valid (110-min TTL < 120-min token lifetime).
// Returns null (non-throwing) on any failure so callers can handle gracefully.
export async function getRCAccessToken(): Promise<string | null> {
  if (!RC_CLIENT_ID || !RC_JWT) {
    console.log("RC auth skipped: RC_CLIENT_ID or RC_JWT env var not set");
    return null;
  }

  // Check cache first — if token is still valid, return it
  if (
    _rcTokenCache && _rcTokenExpiresAt > Date.now()
  ) {
    return _rcTokenCache;
  }

  try {
    const resp = await fetchWithTimeout(
      "https://platform.ringcentral.com/restapi/oauth/token",
      {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
          // Authorization = "Basic " + base64(clientId:clientSecret)
          "Authorization": `Basic ${
            btoa(`${RC_CLIENT_ID}:${RC_CLIENT_SECRET}`)
          }`,
        },
        body: new URLSearchParams({
          grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
          assertion: RC_JWT,
        }).toString(),
      },
      12000, // 12 s timeout — auth endpoint is typically fast
    );

    if (!resp.ok) {
      const errText = await resp.text().catch(() => "");
      console.log(
        `RC token exchange failed: HTTP ${resp.status} — ${
          errText.substring(0, 200)
        }`,
      );
      return null;
    }

    const d = await resp.json();
    const token = d.access_token || null;
    if (token) {
      // Cache token for 110 minutes (66,000 ms). Token lifetime is 120 min; we use 110 min to be safe.
      _rcTokenCache = token;
      _rcTokenExpiresAt = Date.now() + 110 * 60 * 1000;
    }
    return token;
  } catch (e: any) {
    console.log(`RC auth error: ${e.message}`);
    return null;
  }
}

// ── sendSMS ───────────────────────────────────────────────────────────────────
// Send a single SMS via the RingCentral REST API.
//
// @param toNumber  Recipient in E.164 format — e.g. "+15205551234"
//                  Val Town stores this in tech_grades.tech_phone
// @param message   Plain text body (no HTML). Keep under 160 chars for
//                  single-segment delivery; RC handles splitting automatically.
//
// Returns: { ok: true, message_id } on success
//          { ok: false, error: "..." } on any failure — never throws
export async function sendSMS(
  toNumber: string,
  message: string,
): Promise<{ ok: boolean; message_id?: string; error?: string }> {
  // Guard: RC not configured
  if (!RC_FROM_NUMBER || !RC_CLIENT_ID) {
    return {
      ok: false,
      error:
        "RingCentral not configured — set RC_CLIENT_ID, RC_CLIENT_SECRET, RC_JWT, RC_FROM_NUMBER env vars",
    };
  }

  // Guard: empty/invalid recipient
  if (!toNumber || !toNumber.startsWith("+")) {
    return {
      ok: false,
      error:
        `Invalid toNumber "${toNumber}" — must be E.164 format e.g. +15205551234`,
    };
  }

  const token = await getRCAccessToken();
  if (!token) {
    return {
      ok: false,
      error: "RC token exchange failed — verify RC_JWT env var",
    };
  }

  try {
    const resp = await fetchWithTimeout(
      "https://platform.ringcentral.com/restapi/v1.0/account/~/extension/~/sms",
      {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          from: { phoneNumber: RC_FROM_NUMBER },
          to: [{ phoneNumber: toNumber }],
          text: message,
        }),
      },
      12000,
    );

    if (!resp.ok) {
      const errText = await resp.text().catch(() => "");
      return {
        ok: false,
        error: `RC SMS HTTP ${resp.status}: ${errText.substring(0, 200)}`,
      };
    }

    const d = await resp.json();
    return {
      ok: true,
      message_id: String(d.id || d.messageId || ""),
    };
  } catch (e: any) {
    return { ok: false, error: e.message };
  }
}

// ── sendBulkSMS ───────────────────────────────────────────────────────────────
// Send the same message to a list of recipients sequentially with a 300 ms
// gap between each to avoid bursting RC's per-second limit.
// Used by: executeTier2Blast() in handlers/reassignment.ts
//
// Returns an array of per-recipient results in the same order as recipients.
export async function sendBulkSMS(
  recipients: { phone: string; name: string }[],
  message: string,
): Promise<
  {
    phone: string;
    name: string;
    ok: boolean;
    message_id?: string;
    error?: string;
  }[]
> {
  const { delay } = await import("./fetchUtils.ts");
  const results: {
    phone: string;
    name: string;
    ok: boolean;
    message_id?: string;
    error?: string;
  }[] = [];

  for (const r of recipients) {
    const res = await sendSMS(r.phone, message);
    results.push({ ...r, ...res });
    await delay(300); // 300 ms between RC calls
  }

  return results;
}