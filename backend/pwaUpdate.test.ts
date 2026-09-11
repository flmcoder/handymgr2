import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

import { setStaticCacheHeaders } from './staticCache.ts';

const strictNoCache = 'no-cache, no-store, must-revalidate';

test('static cache headers disable caching for app entry points and manifests', () => {
  for (const filePath of [
    '/srv/dist/index.html',
    '/srv/dist/sw.js',
    '/srv/dist/manifest.json',
    '/srv/dist/manifest.webmanifest',
  ]) {
    const headers = new Map<string, string>();
    setStaticCacheHeaders({
      setHeader(name: string, value: string) {
        headers.set(name, value);
      },
    }, filePath);

    assert.equal(headers.get('Cache-Control'), strictNoCache);
    assert.equal(headers.get('Pragma'), 'no-cache');
    assert.equal(headers.get('Expires'), '0');
  }
});

test('static cache headers leave hashed assets available for long-lived caching', () => {
  const headers = new Map<string, string>();
  setStaticCacheHeaders({
    setHeader(name: string, value: string) {
      headers.set(name, value);
    },
  }, '/srv/dist/assets/index-a1b2c3.js');

  assert.equal(headers.size, 0);
});

test('Express static serving uses the cache-control header hook', async () => {
  const source = await readFile(new URL('./server.ts', import.meta.url), 'utf8');
  assert.match(source, /express\.static\(DIST_DIR,[\s\S]*setHeaders:\s*setStaticCacheHeaders/);
});

test('Vite emits content-hashed production assets', async () => {
  const source = await readFile(new URL('../vite.config.ts', import.meta.url), 'utf8');

  assert.match(source, /entryFileNames:\s*`assets\/\[name\]-\[hash\]\.js`/);
  assert.match(source, /assetFileNames:\s*`assets\/\[name\]-\[hash\]\.\[ext\]`/);
});

test('frontend polls for waiting service worker updates and routes them through the blocking version modal', async () => {
  const source = await readFile(new URL('../src/app.js', import.meta.url), 'utf8');

  assert.match(source, /addEventListener\(['"]updatefound['"]/);
  assert.match(source, /then\(function\(registration\)[\s\S]*registration\.update\(\)/);
  assert.match(source, /document\.visibilityState\s*===\s*['"]visible['"]/);
  // A newly installed worker must defer to the real version check, not show its own UI.
  assert.match(source, /installingWorker\.state\s*===\s*['"]installed['"][\s\S]{0,200}void checkForRequiredAppUpdate\(\)/);
  // The only update surface is the blocking modal — no separate in-page corner prompt.
  assert.doesNotMatch(source, /app-update-prompt/);
  assert.doesNotMatch(source, /postMessage\(\{\s*type:\s*['"]SKIP_WAITING['"]\s*\}\)/);
  assert.match(source, /New version available/);
  assert.match(source, /Refresh now/);
});

test('service worker waits for refresh approval and handles SKIP_WAITING', async () => {
  const source = await readFile(new URL('../public/sw.js', import.meta.url), 'utf8');

  assert.doesNotMatch(source, /addEventListener\(['"]install['"][\s\S]{0,400}skipWaiting\(\)/);
  assert.match(source, /event\.data\.type\s*===\s*['"]SKIP_WAITING['"]/);
  assert.match(source, /self\.skipWaiting\(\)/);
});

test('version update modal only fires for a genuine newer server version and cannot be dismissed', async () => {
  const source = await readFile(new URL('../src/app.js', import.meta.url), 'utf8');

  // Every trigger point (ping pre-flight, health-check poll) is gated on the real
  // version comparison — no unconditional/silent location.reload() version path.
  assert.doesNotMatch(source, /hm_version_mismatch/);
  assert.match(source, /shouldForceVersionReload\(SERVER_VERSION,\s*APP_VERSION\)/);
  assert.match(source, /shouldForceVersionReload\(currentServerVersion,\s*APP_VERSION\)/);
  // Overlay clicks and Escape must not dismiss the modal — only the refresh button can.
  assert.match(source, /overlay\.addEventListener\(['"]click['"],\s*function\(e\)\s*\{\s*e\.stopPropagation\(\);\s*\}\)/);
  assert.match(source, /if\s*\(e\.key === ['"]Escape['"]\)\s*\{\s*e\.preventDefault\(\);\s*e\.stopPropagation\(\);\s*\}/);
});

test('visible tabs poll the health endpoint and check immediately when they regain focus', async () => {
  const source = await readFile(new URL('../src/app.js', import.meta.url), 'utf8');

  assert.match(source, /VERSION_MISMATCH_POLL_MS\s*=\s*60\s*\*\s*1000/);
  assert.match(source, /setInterval\(function\(\)\s*\{[\s\S]{0,180}checkForRequiredAppUpdate\(\)/);
  assert.match(source, /document\.addEventListener\(['"]visibilitychange['"][\s\S]{0,240}checkForRequiredAppUpdate\(\)/);
});

test('frontend startup does not invoke an undefined history renderer', async () => {
  const source = await readFile(new URL('../src/app.js', import.meta.url), 'utf8');

  assert.doesNotMatch(source, /^\s*renderHistory\(\);\s*$/m);
});