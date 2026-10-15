import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildPostgresResilienceOptions,
  computeDbBackoffMs,
  isTransientDbError,
  withDbRetry,
} from './dbResiliencePolicy.ts';

const err = (code: string, message = 'x') => Object.assign(new Error(message), { code });

test('classifies tunnel and server-restart failures as transient, including wrapped causes', () => {
  for (const code of ['CONNECT_TIMEOUT', 'CONNECTION_CLOSED', 'CONNECTION_ENDED', 'ECONNRESET', 'ETIMEDOUT', 'EPIPE', 'ECONNREFUSED', '57P01', '08006', '53300', '40P01']) {
    assert.equal(isTransientDbError(err(code)), true, code);
  }
  assert.equal(isTransientDbError(new Error('Failed query', { cause: err('ECONNRESET') })), true);
  assert.equal(isTransientDbError(new Error('terminating connection due to administrator command')), true);
});

test('does not retry data, constraint, or schema errors', () => {
  for (const code of ['23505', '23503', '42703', '22P02', '42P01']) {
    assert.equal(isTransientDbError(err(code)), false, code);
  }
  assert.equal(isTransientDbError(null), false);
  assert.equal(isTransientDbError('boom'), false);
  assert.equal(isTransientDbError(new Error('duplicate key')), false);
  const circular: Record<string, unknown> = { code: '23505' };
  circular.cause = circular;
  assert.equal(isTransientDbError(circular), false);
});

test('backoff doubles per attempt, is capped, and tolerates bad input', () => {
  const opts = { baseMs: 500, maxMs: 4_000 };
  assert.deepEqual([1, 2, 3, 4, 5, 9].map((n) => computeDbBackoffMs(n, opts)), [500, 1_000, 2_000, 4_000, 4_000, 4_000]);
  assert.equal(computeDbBackoffMs(0, opts), 500);
  assert.equal(computeDbBackoffMs(Number.NaN, opts), 500);
  assert.equal(computeDbBackoffMs(2, { ...opts, jitter: 0.5, random: () => 1 }), 1_500);
  assert.equal(computeDbBackoffMs(4, { ...opts, jitter: 0.5, random: () => 1 }), 4_000);
});

test('withDbRetry retries transient failures with growing delays then succeeds', async () => {
  const delays: number[] = [];
  let calls = 0;
  const result = await withDbRetry(async () => {
    if (++calls < 3) throw err('ECONNRESET');
    return 'ok';
  }, { baseMs: 100, maxMs: 1_000, jitter: 0, sleep: async (ms) => { delays.push(ms); } });
  assert.equal(result, 'ok');
  assert.equal(calls, 3);
  assert.deepEqual(delays, [100, 200]);
});

test('withDbRetry throws immediately for non-transient errors and after exhausting attempts', async () => {
  let calls = 0;
  await assert.rejects(withDbRetry(async () => { calls++; throw err('23505'); }, { sleep: async () => {} }), /x/);
  assert.equal(calls, 1);

  calls = 0;
  const retries: number[] = [];
  await assert.rejects(withDbRetry(async () => { calls++; throw err('CONNECT_TIMEOUT'); }, {
    attempts: 3, sleep: async () => {}, onRetry: ({ attempt }) => retries.push(attempt),
  }));
  assert.equal(calls, 3);
  assert.deepEqual(retries, [1, 2]);

  calls = 0;
  await assert.rejects(withDbRetry(async () => { calls++; throw err('ECONNRESET'); }, { attempts: 0, sleep: async () => {} }));
  assert.equal(calls, 1);
});

test('pool options enable keep-alive, recycling, and exponential reconnect backoff', () => {
  const options = buildPostgresResilienceOptions({});
  assert.equal(options.keep_alive, 30);
  assert.equal(options.max_lifetime, 1_800);
  assert.equal(options.connect_timeout, 15);
  assert.equal(options.idle_timeout, 20);
  assert.equal(options.max, 10);
  const first = options.backoff(0);
  const later = options.backoff(10);
  assert.ok(first >= 0.5 && first <= 0.7, String(first));
  assert.ok(later <= 15 && later > first);
});

test('pool options honour env overrides and ignore invalid values', () => {
  const options = buildPostgresResilienceOptions({
    DB_POOL_MAX: '4', DB_KEEP_ALIVE: '10', DB_CONNECT_TIMEOUT: 'abc', DB_MAX_LIFETIME: '5', DB_IDLE_TIMEOUT: '0',
  });
  assert.equal(options.max, 4);
  assert.equal(options.keep_alive, 10);
  assert.equal(options.connect_timeout, 15);
  assert.equal(options.max_lifetime, 60);
  assert.equal(options.idle_timeout, 0);
});
