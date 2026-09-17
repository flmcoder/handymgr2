import assert from 'node:assert/strict';
import test from 'node:test';
import { normalizeDebugEvent, resolveDebugRetentionDays } from './debugEventsPolicy.ts';

test('normalizeDebugEvent clips fields and redacts credentials', () => {
  const event = normalizeDebugEvent({
    level: 'error',
    event_type: 'api_error',
    source: 'fetch',
    message: 'Request failed',
    details: { authorization: 'Bearer secret', nested: { password: 'hidden', code: 401 } },
  });

  assert.equal(event?.level, 'error');
  assert.equal(event?.details.authorization, '[redacted]');
  assert.deepEqual(event?.details.nested, { password: '[redacted]', code: 401 });
});

test('normalizeDebugEvent rejects empty messages and normalizes levels', () => {
  assert.equal(normalizeDebugEvent({ message: '' }), null);
  assert.equal(normalizeDebugEvent({ level: 'unknown', message: 'hello' })?.level, 'info');
});

test('resolveDebugRetentionDays bounds configured retention', () => {
  assert.equal(resolveDebugRetentionDays('90'), 90);
  assert.equal(resolveDebugRetentionDays('9999'), 365);
  assert.equal(resolveDebugRetentionDays('0'), 1);
  assert.equal(resolveDebugRetentionDays('bad'), 30);
});
