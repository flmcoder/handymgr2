import assert from 'node:assert/strict';
import test from 'node:test';
import requestValidator from '../middleware/requestValidator.js';

function makeResponse() {
  return {
    statusCode: null,
    ended: false,
    status(code) {
      this.statusCode = code;
      return this;
    },
    end() {
      this.ended = true;
      return this;
    },
  };
}

test('traffic guard rejects known automated scraper and scanner user agents', () => {
  for (const userAgent of [
    'Go-http-client/1.1',
    'python-requests/2.32.0',
    'curl/8.5.0',
    'Wget/1.21.4',
    'Nikto/2.5.0',
    'sqlmap/1.8',
  ]) {
    const res = makeResponse();
    let nextCalled = false;

    requestValidator({ method: 'GET', path: '/api/health', headers: { 'user-agent': userAgent } }, res, () => {
      nextCalled = true;
    });

    assert.equal(res.statusCode, 403, `${userAgent} should be blocked`);
    assert.equal(res.ended, true);
    assert.equal(nextCalled, false);
  }
});

test('traffic guard rejects POST to the root path with 405', () => {
  const res = makeResponse();
  let nextCalled = false;

  requestValidator({ method: 'POST', path: '/', headers: { 'user-agent': 'Mozilla/5.0' } }, res, () => {
    nextCalled = true;
  });

  assert.equal(res.statusCode, 405);
  assert.equal(res.ended, true);
  assert.equal(nextCalled, false);
});

test('traffic guard allows normal browser and API requests through', () => {
  for (const req of [
    { method: 'GET', path: '/', headers: { 'user-agent': 'Mozilla/5.0' } },
    { method: 'POST', path: '/api/work-orders', headers: { 'user-agent': 'Mozilla/5.0' } },
    { method: 'GET', path: '/api/work-orders', headers: {}, get: () => 'Mozilla/5.0' },
    { method: 'GET', path: '/api/health', headers: {} },
  ]) {
    const res = makeResponse();
    let nextCalled = false;

    requestValidator(req, res, () => {
      nextCalled = true;
    });

    assert.equal(nextCalled, true);
    assert.equal(res.statusCode, null);
    assert.equal(res.ended, false);
  }
});