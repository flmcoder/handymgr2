import assert from 'node:assert/strict';
import test from 'node:test';
import { enforceScopedSession } from './scopedSessionGuard.ts';

test('enforceScopedSession stops when authentication fails', async () => {
  let scopeApplied = false;
  const allowed = await enforceScopedSession({}, {}, {
    requireSession: async () => null,
    applyScope: () => {
      scopeApplied = true;
      return true;
    },
  });

  assert.equal(allowed, false);
  assert.equal(scopeApplied, false);
});

test('enforceScopedSession stops when trusted scope is rejected', async () => {
  const session = { role: 'pm_readonly' };
  const allowed = await enforceScopedSession({}, {}, {
    requireSession: async () => session,
    applyScope: (_request, _response, receivedSession) => {
      assert.equal(receivedSession, session);
      return false;
    },
  });

  assert.equal(allowed, false);
});

test('enforceScopedSession allows an authenticated scoped request', async () => {
  const allowed = await enforceScopedSession({}, {}, {
    requireSession: async () => ({ role: 'manager' }),
    applyScope: () => true,
  });

  assert.equal(allowed, true);
});