const SCANNER_PREFIXES = ['/.env', '/.git', '/.aws', '/.kube', '/wp-admin', '/wp-login', '/wp-json'];

/** Identify common internet scanner probes; API paths are deliberately exempt. */
export function isKnownScannerProbePath(pathname: unknown): boolean {
  const path = String(pathname || '').toLowerCase();
  if (!path.startsWith('/') || path === '/api' || path.startsWith('/api/')) return false;
  if (SCANNER_PREFIXES.some((prefix) => path.startsWith(prefix))) return true;
  if (path === '/xmlrpc.php') return true;
  return /(?:^|\/)[^/]+\.php(?:\/|$)/i.test(path);
}