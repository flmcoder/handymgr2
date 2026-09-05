export interface StaticHeaderResponse {
  setHeader(name: string, value: string): void;
}

const STRICT_NO_CACHE = 'no-cache, no-store, must-revalidate';

export function setStaticCacheHeaders(response: StaticHeaderResponse, filePath: string): void {
  const fileName = filePath.split(/[\\/]/).pop()?.toLowerCase() || '';
  const isManifest = fileName === 'manifest.json'
    || fileName.endsWith('.manifest.json')
    || fileName.endsWith('.webmanifest');

  if (fileName !== 'index.html' && fileName !== 'sw.js' && !isManifest) return;

  response.setHeader('Cache-Control', STRICT_NO_CACHE);
  response.setHeader('Pragma', 'no-cache');
  response.setHeader('Expires', '0');
}