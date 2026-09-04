import assert from 'node:assert/strict';
import test from 'node:test';

import { isClientAbortError } from './requestErrorPolicy.ts';

test('recognizes Express client disconnect errors', () => {
  assert.equal(isClientAbortError({ code: 'ECONNABORTED', message: 'request aborted' }), true);
  assert.equal(isClientAbortError({ code: 'ECONNRESET', message: 'socket hang up' }), true);
});

test('recognizes abort messages without a transport code', () => {
  assert.equal(isClientAbortError(new Error('request aborted')), true);
});

test('does not hide application or database failures', () => {
  assert.equal(isClientAbortError(new Error('query failed')), false);
  assert.equal(isClientAbortError(null), false);
});