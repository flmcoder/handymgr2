/**
 * AppFolio credential resolver.
 * Builds the Authorization and Accept headers for v0 (Database API)
 * and v2 (Reports API) calls from environment variables.
 *
 * Required env vars:
 *   AF_CLIENT_ID      — AppFolio API client ID
 *   AF_CLIENT_SECRET  — AppFolio API client secret
 *   AF_SUBDOMAIN      — Your AppFolio subdomain (e.g. 'flraz')
 *
 * Optional:
 *   AF_DB_BASE        — Override v0 base URL (default: https://api.appfolio.com)
 *   AF_REPORTS_BASE   — Override v2 base URL (default: https://<subdomain>.appfolio.com)
 */

const CLIENT_ID     = String(process.env.AF_CLIENT_ID     || process.env.GUI_MANAGER_ID  || '').trim();
const CLIENT_SECRET = String(process.env.AF_CLIENT_SECRET || process.env.GUI_MANAGER_PW  || '').trim();
const SUBDOMAIN     = String(process.env.AF_SUBDOMAIN     || process.env.APPFOLIO_SUBDOMAIN || 'flraz').trim();

export const AF_DB_BASE      = String(process.env.AF_DB_BASE      || 'https://api.appfolio.com').trim();
export const AF_REPORTS_BASE = String(process.env.AF_REPORTS_BASE || `https://${SUBDOMAIN}.appfolio.com`).trim();

let _cachedBasic: string | null = null;

function basicToken(): string {
  if (!_cachedBasic) {
    if (!CLIENT_ID || !CLIENT_SECRET) {
      throw new Error(
        '[afCredentials] AF_CLIENT_ID and AF_CLIENT_SECRET are required. ' +
        'Set these env vars in Render before running sync jobs.',
      );
    }
    _cachedBasic = Buffer.from(`${CLIENT_ID}:${CLIENT_SECRET}`).toString('base64');
  }
  return _cachedBasic;
}

export function afHeaders(apiVersion: string = 'v0'): Record<string, string> {
  return {
    Authorization: `Basic ${basicToken()}`,
    Accept: 'application/json',
    'Content-Type': 'application/json',
    'User-Agent': 'handymgr2-sync/1.0',
  };
}

export function afBaseUrl(apiVersion: string = 'v0'): string {
  return apiVersion === 'v2' ? AF_REPORTS_BASE : AF_DB_BASE;
}
