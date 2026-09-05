import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';

import {
  buildMagicPortalLink,
  buildMagicPortalSmsMessage,
  consumeMagicTokenTransaction,
  normalizePortalPayload,
  resolveMagicPortalBaseUrl,
  sendMagicPortalSms,
} from './magicPortal.ts';

test('resolveMagicPortalBaseUrl prefers an explicit public Render URL', () => {
  assert.equal(resolveMagicPortalBaseUrl({
    BASE_URL: '',
    RENDER_EXTERNAL_URL: 'https://handymgr2.onrender.com/',
    HOST: '0.0.0.0',
  }), 'https://handymgr2.onrender.com');
});

test('resolveMagicPortalBaseUrl rejects localhost and undefined hosts', () => {
  assert.throws(() => resolveMagicPortalBaseUrl({ BASE_URL: 'http://localhost:3000' }), /public HTTPS/);
  assert.throws(() => resolveMagicPortalBaseUrl({ HOST: 'undefined' }), /public HTTPS/);
});

test('resolveMagicPortalBaseUrl uses the legacy proxy URL only as a final public fallback', () => {
  assert.equal(resolveMagicPortalBaseUrl({
    RENDER_EXTERNAL_URL: '',
    APP_ORIGIN: '',
    PROXY_BASE_URL: 'https://handymgr.app/',
  }), 'https://handymgr.app');
});

test('buildMagicPortalLink creates an encoded Express short URL', () => {
  assert.equal(
    buildMagicPortalLink('https://handymgr.app/', 'abc_123'),
    'https://handymgr.app/s/abc_123',
  );
});

test('buildMagicPortalSmsMessage includes the work order and live short link', () => {
  assert.equal(
    buildMagicPortalSmsMessage('1042', 'https://handymgr.app/s/abc_123'),
    'Fort Lowell Realty dispatch link for WO #1042: https://handymgr.app/s/abc_123',
  );
});

test('sendMagicPortalSms posts the expected RingCentral payload', async () => {
  const calls: Array<{ url: string; init?: RequestInit }> = [];
  const fetchStub = async (url: string | URL | Request, init?: RequestInit) => {
    calls.push({ url: String(url), init });
    return new Response(JSON.stringify({ id: 'rc-message-1' }), { status: 200 });
  };

  const result = await sendMagicPortalSms(
    '(520) 555-0100',
    'Portal link',
    {
      RC_ACCESS_TOKEN: 'test-token',
      RC_FROM_NUMBER: '+15205550199',
      RC_SERVER_URL: 'https://platform.ringcentral.com',
    },
    fetchStub,
  );

  assert.equal(result.messageId, 'rc-message-1');
  assert.equal(calls.length, 1);
  assert.deepEqual(JSON.parse(String(calls[0].init?.body)), {
    from: { phoneNumber: '+15205550199' },
    to: [{ phoneNumber: '+15205550100' }],
    text: 'Portal link',
  });
});

test('sendMagicPortalSms exposes a failed RingCentral hand-off', async () => {
  const fetchStub = async () => new Response('upstream unavailable', { status: 503 });
  await assert.rejects(
    sendMagicPortalSms(
      '+15205550100',
      'Portal link',
      { RC_ACCESS_TOKEN: 'test-token', RC_FROM_NUMBER: '+15205550199' },
      fetchStub,
    ),
    /RingCentral HTTP 503/,
  );
});

test('sendMagicPortalSms refuses to send credentials to a non-RingCentral host', async () => {
  await assert.rejects(
    sendMagicPortalSms(
      '+15205550100',
      'Portal link',
      {
        RC_ACCESS_TOKEN: 'test-token',
        RC_FROM_NUMBER: '+15205550199',
        RC_SERVER_URL: 'https://attacker.example',
      },
    ),
    /ringcentral\.com host/,
  );
});

test('normalizePortalPayload accepts JSON and urlencoded body values', () => {
  assert.deepEqual(normalizePortalPayload({ token: ' token ', status: 'Waiting', note_text: 42 }), {
    token: 'token',
    status: 'Waiting',
    noteText: '42',
    action: 'status_update',
  });
});

test('consumeMagicTokenTransaction commits work-order update before token use', async () => {
  const statements: string[] = [];
  const client = {
    unsafe: async (sql: string) => {
      statements.push(sql.trim());
      if (/select .* from magic_tokens/is.test(sql)) {
        return [{ token: 'valid', wo_id: 'wo-1', used: false, expires_at: new Date(Date.now() + 60_000) }];
      }
      return [{ id: 'wo-1' }];
    },
    release: () => { statements.push('RELEASE'); },
  };

  const result = await consumeMagicTokenTransaction(
    { reserve: async () => client },
    { token: 'valid', status: 'Waiting', noteText: 'Parts ordered' },
  );

  assert.equal(result.workOrderId, 'wo-1');
  assert.match(statements[0], /^BEGIN/i);
  assert.match(statements[1], /FOR UPDATE/i);
  assert.match(statements[2], /UPDATE appfolio_work_orders/i);
  assert.match(statements[3], /UPDATE magic_tokens/i);
  assert.match(statements[4], /^COMMIT/i);
  assert.equal(statements.at(-1), 'RELEASE');
});

test('consumeMagicTokenTransaction rolls back and leaves token unused when work-order update fails', async () => {
  const statements: string[] = [];
  const client = {
    unsafe: async (sql: string) => {
      statements.push(sql.trim());
      if (/select .* from magic_tokens/is.test(sql)) {
        return [{ token: 'valid', wo_id: 'wo-1', used: false, expires_at: new Date(Date.now() + 60_000) }];
      }
      if (/UPDATE appfolio_work_orders/i.test(sql)) throw new Error('work order write failed');
      return [];
    },
    release: () => { statements.push('RELEASE'); },
  };

  await assert.rejects(
    consumeMagicTokenTransaction(
      { reserve: async () => client },
      { token: 'valid', status: 'Waiting', noteText: '' },
    ),
    /work order write failed/,
  );
  assert.ok(statements.some((sql) => /^ROLLBACK/i.test(sql)));
  assert.ok(!statements.some((sql) => /UPDATE magic_tokens/i.test(sql)));
  assert.equal(statements.at(-1), 'RELEASE');
});

test('consumeMagicTokenTransaction rejects an already-used token', async () => {
  const statements: string[] = [];
  const client = {
    unsafe: async (sql: string) => {
      statements.push(sql.trim());
      if (/select .* from magic_tokens/is.test(sql)) {
        return [{ token: 'used', wo_id: 'wo-1', used: true, expires_at: new Date(Date.now() + 60_000) }];
      }
      return [];
    },
    release: () => { statements.push('RELEASE'); },
  };

  await assert.rejects(
    consumeMagicTokenTransaction(
      { reserve: async () => client },
      { token: 'used', status: 'Waiting', noteText: '' },
    ),
    /already been used/,
  );
  assert.ok(statements.some((sql) => /^ROLLBACK/i.test(sql)));
});

test('consumeMagicTokenTransaction does not consume a token for an empty status update', async () => {
  let reserved = false;
  await assert.rejects(
    consumeMagicTokenTransaction(
      { reserve: async () => { reserved = true; throw new Error('must not reserve'); } },
      { token: 'valid', status: '', noteText: '', action: 'status_update' },
    ),
    /status is required/,
  );
  assert.equal(reserved, false);
});

test('Express wires the Magic Portal receiver with JSON and urlencoded middleware', () => {
  const serverSource = readFileSync(new URL('./server.ts', import.meta.url), 'utf8');
  assert.match(
    serverSource,
    /app\.post\('\/api\/magic-portal\/submit', magicPortalJson, magicPortalForm/,
  );
  assert.match(serverSource, /res\.status\(200\)\.json\(\{\s*ok: true,\s*received: true/);
  assert.match(serverSource, /portal_status.*portal_schedule.*portal_reschedule.*portal_note.*portal_reassign_request/s);
});
