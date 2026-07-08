/**
 * AppFolio credential resolver.
 * Builds the Authorization and Accept headers for v0 (Database API)
 * and v2 (Reports API) calls from environment variables.
 *
 * Required env vars:
 *   AF_DEVELOPER_ID   — AppFolio developer ID header
 *   AF_CLIENT_ID      — AppFolio API client ID
 *   AF_CLIENT_SECRET  — AppFolio API client secret
 *   AF_SUBDOMAIN      — Your AppFolio subdomain (e.g. 'flraz')
 *
 * Legacy aliases accepted for compatibility:
 *   AF_V0_CLIENT_ID, AF_V0_CLIENT_SECRET, AF_VHOST, GUI_MANAGER_ID, GUI_MANAGER_PW
 *
 * Optional:
 *   AF_DB_BASE        — Override v0 base URL (default: https://api.appfolio.com)
 *   AF_REPORTS_BASE   — Override v2 base URL (default: https://<subdomain>.appfolio.com)
 */

const DEVELOPER_ID  = String(process.env.AF_DEVELOPER_ID || process.env.DEV || '').trim();
const CLIENT_ID     = String(
  process.env.AF_CLIENT_ID ||
  process.env.AF_V0_CLIENT_ID ||
  process.env.AF_V0_CLIENT_ID_store ||
  process.env.GUI_MANAGER_ID ||
  ''
).trim();
const CLIENT_SECRET = String(
  process.env.AF_CLIENT_SECRET ||
  process.env.AF_V0_CLIENT_SECRET ||
  process.env.AF_V0_CLIENT_SECRET_store ||
  process.env.GUI_MANAGER_PW ||
  ''
).trim();
const REPORTS_CLIENT_ID = String(
  process.env.AF_REPORTS_CLIENT_ID ||
  process.env.AF_V2_CLIENT_ID ||
  process.env.AF_CLIENT_ID ||
  process.env.AF_V0_CLIENT_ID ||
  process.env.AF_V0_CLIENT_ID_store ||
  process.env.GUI_MANAGER_ID ||
  ''
).trim();
const REPORTS_CLIENT_SECRET = String(
  process.env.AF_REPORTS_CLIENT_SECRET ||
  process.env.AF_V2_CLIENT_SECRET ||
  process.env.AF_CLIENT_SECRET ||
  process.env.AF_V0_CLIENT_SECRET ||
  process.env.AF_V0_CLIENT_SECRET_store ||
  process.env.GUI_MANAGER_PW ||
  ''
).trim();
const SUBDOMAIN     = String(process.env.AF_SUBDOMAIN || process.env.AF_VHOST || process.env.APPFOLIO_SUBDOMAIN || 'flraz').trim();

function normalizeAppfolioBaseUrl(rawValue: string, fallbackValue: string): string {
  const raw = String(rawValue || '').trim() || fallbackValue;
  const withScheme = /^[a-z]+:\/\//i.test(raw) ? raw : `https://${raw}`;

  try {
    const parsed = new URL(withScheme);
    let host = String(parsed.hostname || '').trim().toLowerCase();

    // Guard against duplicated domains from env drift such as
    // flraz.appfolio.com.appfolio.com, which breaks TLS validation.
    if (host.endsWith('.appfolio.com.appfolio.com')) {
      host = host.replace(/\.appfolio\.com\.appfolio\.com$/i, '.appfolio.com');
    }

    parsed.hostname = host;
    parsed.pathname = '';
    parsed.search = '';
    parsed.hash = '';
    return parsed.toString().replace(/\/$/, '');
  } catch {
    return fallbackValue;
  }
}

export const AF_DB_BASE = normalizeAppfolioBaseUrl(
  String(process.env.AF_DB_BASE || 'https://api.appfolio.com').trim(),
  'https://api.appfolio.com',
);

export const AF_REPORTS_BASE = normalizeAppfolioBaseUrl(
  String(process.env.AF_REPORTS_BASE || `https://${SUBDOMAIN}.appfolio.com`).trim(),
  `https://${SUBDOMAIN}.appfolio.com`,
);

let _cachedBasicV0: string | null = null;
let _cachedBasicV2: string | null = null;

function basicToken(apiVersion: string = 'v0'): string {
  const useReportsCreds = apiVersion === 'v2';
  const clientId = useReportsCreds ? REPORTS_CLIENT_ID : CLIENT_ID;
  const clientSecret = useReportsCreds ? REPORTS_CLIENT_SECRET : CLIENT_SECRET;
  const cached = useReportsCreds ? _cachedBasicV2 : _cachedBasicV0;

  if (!cached) {
    if (!clientId || !clientSecret) {
      throw new Error(
        '[afCredentials] AppFolio client credentials are required. ' +
        'Set AF_CLIENT_ID/AF_CLIENT_SECRET or the v0 aliases in Render before running sync jobs.',
      );
    }
    const encoded = Buffer.from(`${clientId}:${clientSecret}`).toString('base64');
    if (useReportsCreds) _cachedBasicV2 = encoded;
    else _cachedBasicV0 = encoded;
  }

  return useReportsCreds ? String(_cachedBasicV2 || '') : String(_cachedBasicV0 || '');
}

export function afHeaders(apiVersion: string = 'v0'): Record<string, string> {
  const headers: Record<string, string> = {
    Authorization: `Basic ${basicToken(apiVersion)}`,
    Accept: 'application/json',
    'Content-Type': 'application/json',
    'User-Agent': 'handymgr2-sync/1.0',
  };

  if (apiVersion !== 'v2' && DEVELOPER_ID) {
    headers['X-AppFolio-Developer-ID'] = DEVELOPER_ID;
  }

  return headers;
}

export function afReportCredentialMode(): 'dedicated_reports_creds' | 'shared_creds' {
  return process.env.AF_REPORTS_CLIENT_ID || process.env.AF_REPORTS_CLIENT_SECRET
    ? 'dedicated_reports_creds'
    : 'shared_creds';
}

export function afBaseUrl(apiVersion: string = 'v0'): string {
  return apiVersion === 'v2' ? AF_REPORTS_BASE : AF_DB_BASE;
}
