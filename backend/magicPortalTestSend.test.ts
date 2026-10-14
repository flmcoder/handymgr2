import assert from 'node:assert/strict';
import type { Request, Response } from 'express';
import test from 'node:test';

import { createMagicPortalTestSendHandler } from './magicPortalTestSend.ts';

function makeRequest(body: Record<string, unknown>): Request {
  return { body } as Request;
}

function makeResponse(): { response: Response; statusCode: number; body: Record<string, unknown> } {
  const result = { statusCode: 200, body: {} as Record<string, unknown> };
  const response = {
    status(code: number) {
      result.statusCode = code;
      return response;
    },
    json(payload: Record<string, unknown>) {
      result.body = payload;
      return response;
    },
  } as unknown as Response;
  return { response, ...result } as unknown as { response: Response; statusCode: number; body: Record<string, unknown> };
}

function captureResponse(): { response: Response; read: () => { statusCode: number; body: Record<string, unknown> } } {
  let statusCode = 200;
  let body: Record<string, unknown> = {};
  const response = {
    status(code: number) {
      statusCode = code;
      return response;
    },
    json(payload: Record<string, unknown>) {
      body = payload;
      return response;
    },
  } as unknown as Response;
  return { response, read: () => ({ statusCode, body }) };
}

test('test-send creates a synthetic session and sends SMS only after full-admin auth', async () => {
  const calls: { created?: Record<string, unknown>; sms?: { phone: string; message: string } } = {};
  const handler = createMagicPortalTestSendHandler({
    requireFullAdminSession: async () => ({ role: 'full' }),
    createSession: async (input) => {
      calls.created = input as unknown as Record<string, unknown>;
      return { magicLink: 'https://handymgr.app/s/test-code-1234', shortCode: 'test-code-1234', expiresAt: '2030-01-01T00:00:00.000Z' };
    },
    sendSms: async (phone, message) => {
      calls.sms = { phone, message };
      return { messageId: 'stub-message-id' };
    },
    buildMessage: (woNumber, link) => `WO ${woNumber}: ${link}`,
  });
  const { response, read } = captureResponse();

  await handler(makeRequest({ phone: '+15551234567' }), response);

  const result = read();
  assert.equal(result.statusCode, 200);
  assert.equal(result.body.ok, true);
  assert.equal(result.body.test, true);
  assert.equal(result.body.magic_link, 'https://handymgr.app/s/test-code-1234');
  assert.equal(result.body.rc_message_id, 'stub-message-id');
  assert.equal(calls.created?.sessionMode, 'test');
  assert.match(String(calls.created?.woId), /^TEST-TEST-/);
  assert.equal(calls.created?.tenantName, 'John Doe (Test Resident)');
  assert.equal(calls.created?.tenantPhone, '+15551232314');
  assert.equal(calls.sms?.phone, '+15551234567');
  assert.match(String(calls.sms?.message), /https:\/\/handymgr\.app\/s\/test-code-1234/);
});

test('test-send stops on failed full-admin authorization and never creates a session or sends SMS', async () => {
  let created = false;
  let sent = false;
  const handler = createMagicPortalTestSendHandler({
    requireFullAdminSession: async (_req, res) => {
      res.status(403).json({ ok: false, error: 'Full administrator session required' });
      return null;
    },
    createSession: async () => {
      created = true;
      throw new Error('must not create');
    },
    sendSms: async () => {
      sent = true;
      return { messageId: '' };
    },
    buildMessage: () => '',
  });
  const { response, read } = captureResponse();

  await handler(makeRequest({ phone: '+15551234567' }), response);

  assert.equal(read().statusCode, 403);
  assert.equal(created, false);
  assert.equal(sent, false);
});

test('test-send rejects invalid phone before session creation or SMS handoff', async () => {
  let created = false;
  let sent = false;
  const handler = createMagicPortalTestSendHandler({
    requireFullAdminSession: async () => ({ role: 'full' }),
    createSession: async () => {
      created = true;
      throw new Error('must not create');
    },
    sendSms: async () => {
      sent = true;
      return { messageId: '' };
    },
    buildMessage: () => '',
  });
  const { response, read } = captureResponse();

  await handler(makeRequest({ phone: '5551234567' }), response);

  assert.equal(read().statusCode, 400);
  assert.equal(created, false);
  assert.equal(sent, false);
});

test('test-send returns provider failure without reporting success', async () => {
  const handler = createMagicPortalTestSendHandler({
    requireFullAdminSession: async () => ({ role: 'full' }),
    createSession: async () => ({ magicLink: 'https://handymgr.app/s/test-code-1234', shortCode: 'test-code-1234', expiresAt: '2030-01-01T00:00:00.000Z' }),
    sendSms: async () => { throw new Error('RingCentral HTTP 503'); },
    buildMessage: () => 'stub message',
  });
  const { response, read } = captureResponse();

  await handler(makeRequest({ phone: '+15551234567' }), response);

  assert.equal(read().statusCode, 502);
  assert.equal(read().body.ok, false);
  assert.equal(read().body.test, true);
  assert.match(String(read().body.error), /503/);
});