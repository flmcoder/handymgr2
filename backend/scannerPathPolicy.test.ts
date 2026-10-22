import assert from 'node:assert/strict';
import test from 'node:test';

import { readFile } from 'node:fs/promises';
import { isKnownScannerProbePath } from './scannerPathPolicy.ts';

test('scanner path policy rejects known probes and PHP paths case-insensitively', () => {
  for (const path of ['/.env', '/.env.production', '/.git/config', '/.aws/credentials', '/.kube/config', '/wp-admin', '/wp-login.php', '/wp-json/wp/v2', '/xmlrpc.php', '/legacy/test.PHP']) {
    assert.equal(isKnownScannerProbePath(path), true, `${path} should be rejected`);
  }
});

test('scanner path policy allows normal pages, assets, health, and all API routes', () => {
  for (const path of ['/', '/dashboard', '/assets/app.js', '/health', '/api', '/api/local/work_orders', '/api/.env', '/api/legacy.php']) {
    assert.equal(isKnownScannerProbePath(path), false, `${path} should be allowed`);
  }
});

test('scanner middleware is registered before application routes and static/SPA handling', async () => {
  const source = await readFile(new URL('./server.ts', import.meta.url), 'utf8');
  const guardIndex = source.indexOf('isKnownScannerProbePath(req.path)');
  const firstRouteIndex = source.indexOf("app.post('/api/webhooks/appfolio'");
  const staticIndex = source.indexOf('app.use(express.static(DIST_DIR');
  const spaIndex = source.indexOf('// SPA fallback: non-API routes');

  assert.ok(guardIndex >= 0);
  assert.ok(guardIndex < firstRouteIndex);
  assert.ok(guardIndex < staticIndex);
  assert.ok(guardIndex < spaIndex);
});