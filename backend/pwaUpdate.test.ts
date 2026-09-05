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

test('frontend prompts for waiting updates and polls when a tab becomes visible', async () => {
  const source = await readFile(new URL('../src/app.js', import.meta.url), 'utf8');

  assert.match(source, /addEventListener\(['"]updatefound['"]/);
  assert.match(source, /then\(function\(registration\)[\s\S]*registration\.update\(\)/);
  assert.match(source, /document\.visibilityState\s*===\s*['"]visible['"]/);
  assert.match(source, /postMessage\(\{\s*type:\s*['"]SKIP_WAITING['"]\s*\}\)/);
  assert.match(source, /window\.location\.reload\(\)/);
  assert.match(source, /Update Available/);
  assert.match(source, /Please Refresh/);
});

test('service worker waits for refresh approval and handles SKIP_WAITING', async () => {
  const source = await readFile(new URL('../public/sw.js', import.meta.url), 'utf8');

  assert.doesNotMatch(source, /addEventListener\(['"]install['"][\s\S]{0,400}skipWaiting\(\)/);
  assert.match(source, /event\.data\.type\s*===\s*['"]SKIP_WAITING['"]/);
  assert.match(source, /self\.skipWaiting\(\)/);
});