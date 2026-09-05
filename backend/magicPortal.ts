import { createHmac, randomBytes } from 'node:crypto';

type Environment = Record<string, string | undefined>;

type FetchLike = typeof fetch;

type SqlConnection = {
  unsafe: (sql: string, params?: unknown[]) => Promise<any[]>;
  release: () => void | Promise<void>;
};

type SqlPool = {
  unsafe?: (sql: string, params?: unknown[]) => Promise<any[]>;
  reserve: () => Promise<SqlConnection>;
};

export type MagicPortalInput = {
  woId: string;
  woNumber: string;
  techId: string;
  techName: string;
  techPhone: string;
  tenantName: string;
  tenantPhone: string;
  propertyAddress: string;
};

export type PortalSubmission = {
  token: string;
  status: string;
  noteText: string;
  action?: string;
};

const ALLOWED_STATUSES = new Set(['Scheduled', 'Waiting', 'Work Completed']);

function normalizePhone(rawPhone: string): string {
  const digits = String(rawPhone || '').replace(/\D/g, '');
  if (digits.length === 10) return `+1${digits}`;
  if (digits.length === 11 && digits.startsWith('1')) return `+${digits}`;
  if (digits.length > 11 && digits.length <= 15) return `+${digits}`;
  return '';
}

function resolveRingCentralServerUrl(env: Environment): string {
  const value = String(env.RC_SERVER_URL || env.RINGCENTRAL_SERVER_URL || 'https://platform.ringcentral.com').trim().replace(/\/+$/, '');
  let parsed: URL;
  try {
    parsed = new URL(value);
  } catch {
    throw new Error('RingCentral server URL is invalid');
  }
  const hostname = parsed.hostname.toLowerCase();
  if (parsed.protocol !== 'https:' || (hostname !== 'ringcentral.com' && !hostname.endsWith('.ringcentral.com'))) {
    throw new Error('RingCentral server URL must use HTTPS on a ringcentral.com host');
  }
  return value;
}

function isPublicHttpsUrl(rawValue: string): boolean {
  try {
    const url = new URL(rawValue);
    const host = url.hostname.toLowerCase();
    return url.protocol === 'https:'
      && !!host
      && host !== 'localhost'
      && host !== '0.0.0.0'
      && host !== '127.0.0.1'
      && host !== 'undefined'
      && host !== '::1';
  } catch {
    return false;
  }
}

export function resolveMagicPortalBaseUrl(env: Environment = process.env): string {
  const candidates = [
    env.MAGIC_PORTAL_BASE_URL,
    env.BASE_URL,
    env.RENDER_EXTERNAL_URL,
    env.APP_ORIGIN,
    env.PROXY_BASE_URL,
    env.HOST,
  ];

  for (const candidate of candidates) {
    const value = String(candidate || '').trim().replace(/\/+$/, '');
    if (isPublicHttpsUrl(value)) return value;
  }

  throw new Error(
    'Magic Portal requires a public HTTPS base URL in MAGIC_PORTAL_BASE_URL, BASE_URL, RENDER_EXTERNAL_URL, APP_ORIGIN, PROXY_BASE_URL, or HOST',
  );
}

export function buildMagicPortalLink(baseUrl: string, shortCode: string): string {
  const normalizedBase = String(baseUrl || '').trim().replace(/\/+$/, '');
  if (!isPublicHttpsUrl(normalizedBase)) {
    throw new Error('Magic Portal base URL must be a public HTTPS URL');
  }
  const normalizedCode = String(shortCode || '').trim();
  if (!/^[A-Za-z0-9_-]{4,64}$/.test(normalizedCode)) {
    throw new Error('Magic Portal short code is invalid');
  }
  return `${normalizedBase}/s/${encodeURIComponent(normalizedCode)}`;
}

export function buildMagicPortalSmsMessage(woNumber: string, magicLink: string): string {
  const workOrderLabel = String(woNumber || '').trim() || 'assigned work order';
  return `Fort Lowell Realty dispatch link for WO #${workOrderLabel}: ${magicLink}`;
}

async function resolveRingCentralAccessToken(env: Environment, fetchImpl: FetchLike): Promise<string> {
  const staticToken = String(env.RC_ACCESS_TOKEN || env.RINGCENTRAL_ACCESS_TOKEN || '').trim();
  if (staticToken) return staticToken;

  const jwt = String(env.RC_JWT || env.RINGCENTRAL_JWT || '').trim();
  const clientId = String(env.RC_CLIENT_ID || env.RINGCENTRAL_CLIENT_ID || '').trim();
  const clientSecret = String(env.RC_CLIENT_SECRET || env.RINGCENTRAL_CLIENT_SECRET || '').trim();
  if (!jwt || !clientId || !clientSecret) {
    throw new Error('RingCentral authentication is not configured for Magic Portal dispatch');
  }

  const serverUrl = resolveRingCentralServerUrl(env);
  const form = new URLSearchParams({
    grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
    assertion: jwt,
  });
  const response = await fetchImpl(`${serverUrl}/restapi/oauth/token`, {
    method: 'POST',
    headers: {
      Authorization: `Basic ${Buffer.from(`${clientId}:${clientSecret}`).toString('base64')}`,
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: form.toString(),
  });
  const payload = await response.json().catch(() => ({})) as Record<string, unknown>;
  if (!response.ok || !payload.access_token) {
    throw new Error(`RingCentral token exchange failed: HTTP ${response.status}`);
  }
  return String(payload.access_token);
}

export async function sendMagicPortalSms(
  toPhone: string,
  message: string,
  env: Environment = process.env,
  fetchImpl: FetchLike = fetch,
): Promise<{ messageId: string }> {
  const recipient = normalizePhone(toPhone);
  const fromPhone = normalizePhone(String(env.RC_FROM_NUMBER || env.RINGCENTRAL_FROM_NUMBER || ''));
  if (!recipient) throw new Error('Magic Portal recipient phone is invalid');
  if (!fromPhone) throw new Error('RingCentral sender number is not configured');

  const accessToken = await resolveRingCentralAccessToken(env, fetchImpl);
  const serverUrl = resolveRingCentralServerUrl(env);
  const response = await fetchImpl(`${serverUrl}/restapi/v1.0/account/~/extension/~/sms`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      from: { phoneNumber: fromPhone },
      to: [{ phoneNumber: recipient }],
      text: message,
    }),
  });
  const responseText = await response.text();
  let payload: Record<string, unknown> = {};
  try {
    payload = responseText ? JSON.parse(responseText) as Record<string, unknown> : {};
  } catch {
    payload = {};
  }
  if (!response.ok) {
    throw new Error(`RingCentral HTTP ${response.status}: ${responseText.slice(0, 180)}`);
  }
  return { messageId: String(payload.id || payload.messageId || '') };
}

export function normalizePortalPayload(body: unknown): PortalSubmission {
  const payload = body && typeof body === 'object' ? body as Record<string, unknown> : {};
  return {
    token: String(payload.token || '').trim(),
    status: String(payload.status || '').trim(),
    noteText: String(payload.note_text ?? payload.noteText ?? '').trim().slice(0, 1200),
    action: String(payload.action || 'status_update').trim() || 'status_update',
  };
}

export async function ensureMagicPortalTables(db: Pick<SqlPool, 'unsafe'>): Promise<void> {
  if (!db.unsafe) throw new Error('Database client does not support direct SQL');
  await db.unsafe(`
    CREATE TABLE IF NOT EXISTS magic_tokens (
      token TEXT PRIMARY KEY,
      short_code TEXT NOT NULL UNIQUE,
      wo_id TEXT NOT NULL,
      wo_number TEXT,
      tech_id TEXT NOT NULL,
      tech_name TEXT NOT NULL,
      tech_phone TEXT,
      tenant_name TEXT,
      tenant_phone TEXT,
      property_address TEXT,
      expires_at TIMESTAMPTZ NOT NULL,
      used BOOLEAN NOT NULL DEFAULT FALSE,
      used_at TIMESTAMPTZ,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      metadata JSONB NOT NULL DEFAULT '{}'::jsonb
    )
  `);
  await db.unsafe('CREATE INDEX IF NOT EXISTS magic_tokens_wo_id_idx ON magic_tokens(wo_id)');
  await db.unsafe('CREATE INDEX IF NOT EXISTS magic_tokens_expires_at_idx ON magic_tokens(expires_at)');
}

function signPortalPayload(payload: Record<string, string>, secret: string): string {
  const encodedPayload = Buffer.from(JSON.stringify(payload)).toString('base64url');
  const signature = createHmac('sha256', secret).update(encodedPayload).digest('base64url');
  return `${encodedPayload}.${signature}`;
}

export async function createMagicPortalSession(
  db: Pick<SqlPool, 'unsafe'>,
  input: MagicPortalInput,
  env: Environment = process.env,
): Promise<{ token: string; shortCode: string; magicLink: string; expiresAt: string }> {
  if (!input.woId || !input.techId || !input.techName || !input.tenantPhone || !input.propertyAddress) {
    throw new Error('Missing required Magic Portal work order, technician, resident, or property fields');
  }
  const secret = String(env.MAGIC_LINK_SECRET || '').trim();
  if (!secret) throw new Error('MAGIC_LINK_SECRET is not configured');
  if (!db.unsafe) throw new Error('Database client does not support direct SQL');

  const baseUrl = resolveMagicPortalBaseUrl(env);
  const expiresAt = new Date(Date.now() + (24 * 60 * 60 * 1000)).toISOString();
  const nonce = randomBytes(18).toString('base64url');
  const shortCode = randomBytes(16).toString('base64url');
  const token = signPortalPayload({
    wo_id: input.woId,
    tech_id: input.techId,
    expires_at: expiresAt,
    nonce,
  }, secret);

  await db.unsafe(
    `INSERT INTO magic_tokens (
       token, short_code, wo_id, wo_number, tech_id, tech_name, tech_phone,
       tenant_name, tenant_phone, property_address, expires_at, metadata
     ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11::timestamptz, $12::jsonb)`,
    [
      token,
      shortCode,
      input.woId,
      input.woNumber,
      input.techId,
      input.techName,
      input.techPhone,
      input.tenantName,
      input.tenantPhone,
      input.propertyAddress,
      expiresAt,
      JSON.stringify(input),
    ],
  );

  return {
    token,
    shortCode,
    magicLink: buildMagicPortalLink(baseUrl, shortCode),
    expiresAt,
  };
}

export async function findMagicPortalToken(
  db: Pick<SqlPool, 'unsafe'>,
  lookup: { token?: string; shortCode?: string },
): Promise<Record<string, any> | null> {
  if (!db.unsafe) throw new Error('Database client does not support direct SQL');
  const token = String(lookup.token || '').trim();
  const shortCode = String(lookup.shortCode || '').trim();
  if (!token && !shortCode) return null;
  const rows = await db.unsafe(
    `SELECT token, short_code, wo_id, wo_number, tech_id, tech_name, tech_phone,
            tenant_name, tenant_phone, property_address, expires_at, used, used_at, metadata
       FROM magic_tokens
      WHERE ($1::text <> '' AND token = $1)
        OR ($2::text <> '' AND short_code = $2)
      LIMIT 1`,
    [token, shortCode],
  );
  return rows[0] || null;
}

function escapeHtml(value: unknown): string {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

export function renderMagicPortalHtml(tokenRow: Record<string, any>): string {
  const token = JSON.stringify(String(tokenRow.token || '')).replace(/</g, '\\u003c');
  const expired = new Date(tokenRow.expires_at).getTime() <= Date.now();
  const unavailable = tokenRow.used === true || expired;
  const unavailableMessage = tokenRow.used === true
    ? 'This link has already been submitted.'
    : 'This link has expired. Contact dispatch for a new link.';
  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>HandyManager Work Order Portal</title>
<style>
:root{color-scheme:dark;--bg:#111827;--panel:#1f2937;--line:#374151;--text:#f9fafb;--muted:#9ca3af;--accent:#22c55e;--danger:#ef4444}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font:16px/1.45 ui-sans-serif,system-ui,sans-serif}.shell{max-width:620px;margin:auto;padding:22px}.brand{font-size:12px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:var(--accent)}h1{font-size:25px;margin:7px 0 4px}.meta{color:var(--muted);margin-bottom:18px}.card{background:var(--panel);border:1px solid var(--line);border-radius:8px;padding:18px}.row{margin-bottom:15px}.row label{display:block;font-size:12px;font-weight:700;color:var(--muted);margin-bottom:6px}select,textarea{width:100%;border:1px solid var(--line);border-radius:6px;background:#111827;color:var(--text);padding:12px;font:inherit}textarea{min-height:120px;resize:vertical}button{width:100%;border:0;border-radius:6px;background:var(--accent);color:#052e16;padding:13px;font-weight:800;cursor:pointer}button:disabled{opacity:.55;cursor:not-allowed}.status{margin-top:12px;font-size:14px}.error{color:#fca5a5}.success{color:#86efac}
</style></head><body><main class="shell"><div class="brand">Fort Lowell Realty</div><h1>Work Order #${escapeHtml(tokenRow.wo_number || tokenRow.wo_id)}</h1><div class="meta">${escapeHtml(tokenRow.property_address)} · ${escapeHtml(tokenRow.tech_name)}</div><section class="card">${unavailable
    ? `<p class="error">${escapeHtml(unavailableMessage)}</p>`
    : `<form id="portalForm"><div class="row"><label for="status">Work order status</label><select id="status" name="status" required><option value="">Select status</option><option>Scheduled</option><option>Waiting</option><option>Work Completed</option></select></div><div class="row"><label for="note">Completion or exception note</label><textarea id="note" name="note_text" maxlength="1200"></textarea></div><button id="submit" type="submit">Submit update</button><div id="result" class="status" role="status"></div></form>`}</section></main>${unavailable ? '' : `<script>const token=${token};document.getElementById('portalForm').addEventListener('submit',async(event)=>{event.preventDefault();const button=document.getElementById('submit');const result=document.getElementById('result');button.disabled=true;result.className='status';result.textContent='Submitting…';try{const response=await fetch('/api/magic-portal/submit',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({token,status:document.getElementById('status').value,note_text:document.getElementById('note').value})});const data=await response.json();if(!response.ok||!data.ok)throw new Error(data.error||'Submission failed');result.className='status success';result.textContent='Update received successfully.';}catch(error){button.disabled=false;result.className='status error';result.textContent=error.message||'Submission failed';}});</script>`}</body></html>`;
}

export async function consumeMagicTokenTransaction(
  pool: Pick<SqlPool, 'reserve'>,
  submission: PortalSubmission,
): Promise<{ workOrderId: string; status: string }> {
  const token = String(submission.token || '').trim();
  const status = String(submission.status || '').trim();
  const noteText = String(submission.noteText || '').trim().slice(0, 1200);
  const action = String(submission.action || 'status_update').trim() || 'status_update';
  if (!token) throw new Error('Magic Portal token is required');
  if (!status && action !== 'note') throw new Error('Magic Portal status is required');
  if (status && !ALLOWED_STATUSES.has(status)) throw new Error('Magic Portal status is invalid');

  const connection = await pool.reserve();
  try {
    await connection.unsafe('BEGIN');
    const tokenRows = await connection.unsafe(
      `SELECT token, wo_id, used, expires_at
         FROM magic_tokens
        WHERE token = $1
        FOR UPDATE`,
      [token],
    );
    const tokenRow = tokenRows[0];
    if (!tokenRow) throw new Error('Magic Portal token was not found');
    if (tokenRow.used === true) throw new Error('Magic Portal token has already been used');
    if (new Date(tokenRow.expires_at).getTime() <= Date.now()) throw new Error('Magic Portal token has expired');

    const workOrderRows = await connection.unsafe(
        `UPDATE appfolio_work_orders
          SET status = COALESCE(NULLIF($2::text, ''), status),
              raw_json = COALESCE(raw_json, '{}'::jsonb) || jsonb_build_object(
                'MagicPortalLastSubmission', jsonb_build_object(
                  'action', $4::text,
                  'status', $2::text,
                  'note', $3::text,
                  'submitted_at', NOW()
                )
              ),
              updated_at = NOW()
        WHERE id = $1 OR work_order_uuid = $1
        RETURNING id`,
      [String(tokenRow.wo_id), status, noteText, action],
    );
    if (!workOrderRows[0]) throw new Error('Magic Portal work order was not found');

    const consumedRows = await connection.unsafe(
      `UPDATE magic_tokens
          SET used = TRUE,
              used_at = NOW(),
              metadata = metadata || jsonb_build_object(
                'submission_action', $2::text,
                'submission_status', $3::text
              )
        WHERE token = $1 AND used = FALSE
        RETURNING token`,
      [token, action, status],
    );
    if (!consumedRows[0]) throw new Error('Magic Portal token could not be consumed');

    await connection.unsafe('COMMIT');
    return { workOrderId: String(workOrderRows[0].id), status };
  } catch (error) {
    try {
      await connection.unsafe('ROLLBACK');
    } catch (rollbackError) {
      console.error('[magic-portal] transaction rollback failed', rollbackError);
    }
    throw error;
  } finally {
    await connection.release();
  }
}

export const MAGIC_PORTAL_TRANSACTION_SQL = `
BEGIN;
SELECT token, wo_id, used, expires_at FROM magic_tokens WHERE token = $1 FOR UPDATE;
UPDATE appfolio_work_orders SET status = $2, updated_at = NOW() WHERE id = $1 OR work_order_uuid = $1;
UPDATE magic_tokens SET used = TRUE, used_at = NOW() WHERE token = $1 AND used = FALSE;
COMMIT;
`;