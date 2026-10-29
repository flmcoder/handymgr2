import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

// app.js is a single large browser script, so these checks read its source text.
const source = await readFile(new URL('./app.js', import.meta.url), 'utf8');

function functionBody(name, nextMarker) {
  const start = source.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `${name} must exist`);
  const end = source.indexOf(nextMarker, start + 1);
  assert.notEqual(end, -1, `${name} end marker must exist`);
  return source.slice(start, end);
}

test('rate limiter tracks batch progress', () => {
  assert.match(source, /batchTotal: 0,\s*\n\s*batchDone: 0,/);
});

test('enqueue starts a new batch only when the queue is idle', () => {
  assert.match(source, /if \(rateLimiter\.queue\.length === 0 && rateLimiter\.inFlight === 0\) \{\s*rateLimiter\.batchTotal = 0;\s*rateLimiter\.batchDone = 0;/);
  assert.match(source, /rateLimiter\.batchTotal\+\+;/);
});

test('each settled request advances batchDone on both success and failure', () => {
  const matches = source.match(/rateLimiter\.inFlight--;\s*\n\s*rateLimiter\.batchDone\+\+;/g) || [];
  assert.equal(matches.length, 2);
});

test('header pill shows batch progress instead of per-request slots', () => {
  const body = functionBody('updateRateBadge', '\n// Operational Alerts Engine');
  assert.match(body, /'Loading ' \+ rateLimiter\.batchDone \+ '\/' \+ rateLimiter\.batchTotal/);
  assert.match(body, /'Loaded ' \+ rateLimiter\.batchDone \+ '\/' \+ rateLimiter\.batchTotal/);
  assert.doesNotMatch(body, /inFlight\) \+ '\/'/);
});
