import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import requestValidator, {
  getRootPostAction,
  getRootActionImplementationPolicy,
  isAllowedRootPostAction,
  isRootPostAllowedActionHandled,
  rejectUnrecognizedRootPost,
  ROOT_ACTIONS_WITH_INLINE_HANDLERS,
  ROOT_API_ACTIONS,
} from '../middleware/requestValidator.js';

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

test('rejects POST to root without a recognized action', () => {
  const res = makeResponse();
  let nextCalled = false;

  rejectUnrecognizedRootPost({ method: 'POST', path: '/', headers: {} }, res, () => {
    nextCalled = true;
  });

  assert.equal(res.statusCode, 405);
  assert.equal(res.ended, true);
  assert.equal(nextCalled, false);
});

test('unsupported method at root passes through even when action is unknown', () => {
  const res = makeResponse();
  let nextCalled = false;
  requestValidator({ method: 'PATCH', path: '/', query: { action: 'unknown' } }, res, () => {
    nextCalled = true;
  });
  assert.equal(nextCalled, true);
  assert.equal(res.statusCode, null);
});

test('non-root POSTs are unaffected by root action allowlisting', () => {
  const res = makeResponse();
  let nextCalled = false;
  requestValidator({ method: 'POST', path: '/api/local/debug/events', query: {}, body: {} }, res, () => {
    nextCalled = true;
  });
  assert.equal(nextCalled, true);
  assert.equal(res.statusCode, null);
});

test('object-valued action is not accepted as a root action', () => {
  const res = makeResponse();
  let nextCalled = false;

  requestValidator({
    method: 'POST',
    path: '/',
    query: { action: '' },
    body: { action: { value: 'verify_role' } },
  }, res, () => {
    nextCalled = true;
  });

  assert.equal(getRootPostAction({ query: { action: '' }, body: { action: { value: 'verify_role' } } }), '[object Object]');
  assert.equal(res.statusCode, 405);
  assert.equal(res.ended, true);
  assert.equal(nextCalled, false);
});

test('query and body action resolution is consistent', () => {
  assert.equal(getRootPostAction({ query: { action: '' }, body: { action: { value: 'verify_role' } } }), '[object Object]');
  assert.equal(getRootPostAction({ query: { action: ['verify_role', 'bills'] }, body: {} }), 'verify_role,bills');
  assert.equal(getRootPostAction({ query: { action: [] }, body: { action: 'verify_role' } }), 'verify_role');
});

test('query action has priority over body action', () => {
  assert.equal(getRootPostAction({ query: { action: 'verify_role' }, body: { action: 'bills' } }), 'verify_role');
});

test('empty query action falls back to body action', () => {
  assert.equal(getRootPostAction({ query: { action: '' }, body: { action: 'verify_role' } }), 'verify_role');
});

test('missing query and body actions resolve empty', () => {
  assert.equal(getRootPostAction({ query: {}, body: null }), '');
});

test('root POST validation only applies at exactly /', () => {
  for (const req of [
    { method: 'GET', path: '/', query: {} },
    { method: 'POST', path: '/api', query: {} },
    { method: 'POST', path: '/api/', query: {} },
    { method: 'PATCH', path: '/', query: {} },
  ]) {
    const res = makeResponse();
    let nextCalled = false;

    requestValidator(req, res, () => {
      nextCalled = true;
    });

    assert.equal(nextCalled, true, `${req.method} ${req.path} should pass through`);
    assert.equal(res.statusCode, null);
    assert.equal(res.ended, false);
  }
});

test('allows login and OTP actions through root validation', () => {
  for (const action of ['verify_role', 'device_setup', 'device_otp_request', 'device_otp_verify']) {
    const res = makeResponse();
    let nextCalled = false;

    requestValidator({ method: 'POST', path: '/', query: { action } }, res, () => {
      nextCalled = true;
    });

    assert.equal(nextCalled, true, `${action} should pass through`);
    assert.equal(res.statusCode, null);
    assert.equal(res.ended, false);
  }
});

test('login and OTP request action strings remain recognized root actions', () => {
  for (const action of ['verify_role', 'device_setup', 'device_otp_request', 'device_otp_verify']) {
    assert.equal(isAllowedRootPostAction(action), true, `${action} is required by login/authentication`);
  }
});

test('login and OTP root actions retain their legacy handler entries', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  const routesStart = serverSource.indexOf('const legacyActionRoutes = {');
  const routesEnd = serverSource.indexOf('\n} as const;', routesStart);
  assert.ok(routesStart >= 0 && routesEnd > routesStart);
  const routeMap = serverSource.slice(routesStart, routesEnd);

  for (const action of ['verify_role', 'device_setup', 'device_otp_request', 'device_otp_verify']) {
    assert.match(routeMap, new RegExp(`^\\s{2}${action}:`, 'm'));
  }
});

test('login verifier route delegates to the canonical auth handler', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  assert.match(serverSource, /app\.post\('\/api\/verify_role', wrapDenoHandler\(deviceAuthHandlers\.handleVerifyRole/);
});

test('device auth endpoints remain registered', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  for (const path of ['/api/device/setup', '/api/device/otp/request', '/api/device/otp/verify', '/api/device/verify-role']) {
    assert.ok(serverSource.includes(`app.post('${path}'`), `${path} should remain registered`);
  }
});

test('rejects an allowed root action unless the backend has an inline or mapped handler', () => {
  assert.equal(isRootPostAllowedActionHandled('unknown-action', []), false);
  assert.equal(isRootPostAllowedActionHandled('unknown-action', ['unknown-action']), false);
  assert.equal(isRootPostAllowedActionHandled('units', []), false);
  assert.equal(isRootPostAllowedActionHandled('units', ['units']), true);
  assert.equal(isRootPostAllowedActionHandled('session_info', []), true);
});

test('returns through to the backend for a legitimate root auth request when the user-agent resembles curl', () => {
  const res = makeResponse();
  let nextCalled = false;

  requestValidator({
    method: 'POST',
    path: '/',
    query: { action: 'verify_role' },
    headers: { 'user-agent': 'curl/8.5.0' },
  }, res, () => {
    nextCalled = true;
  });

  assert.equal(nextCalled, true);
  assert.equal(res.statusCode, null);
  assert.equal(res.ended, false);
});

test('common automation user-agents are not blocked on API requests', () => {
  for (const userAgent of ['Go-http-client/1.1', 'python-requests/2.32', 'curl/8.5', 'Wget/1.21', 'Nikto/2.5', 'sqlmap/1.8']) {
    const res = makeResponse();
    let nextCalled = false;
    requestValidator({ method: 'GET', path: '/api/health', headers: { 'user-agent': userAgent } }, res, () => {
      nextCalled = true;
    });
    assert.equal(nextCalled, true, `${userAgent} should not be blocked based on User-Agent alone`);
    assert.equal(res.statusCode, null);
  }
});

test('root API routes remain registered for auth after adding the POST guard', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  for (const path of ['/api/verify_role', '/api/device_setup', '/api/device_otp_request', '/api/device_otp_verify']) {
    assert.ok(serverSource.includes(`app.post('${path}'`), `${path} must remain a registered auth route`);
  }
});

test('root dispatcher uses the same query/body parser as the validator', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  assert.match(serverSource, /function getLegacyAction\(req: Request\): string \{\s*return getRootPostAction\(req\);\s*\}/);
});

test('middleware ordering keeps parser and CORS ahead of validation and routes', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  const jsonIndex = serverSource.indexOf('app.use(express.json({ limit: \'10mb\' }));');
  const corsIndex = serverSource.indexOf('app.use(cors({');
  const validatorIndex = serverSource.indexOf('app.use(rejectUnrecognizedRootPost);');
  const dispatcherIndex = serverSource.indexOf("app.all(['/', '/api', '/api/']");
  const spaIndex = serverSource.indexOf('app.use(express.static(DIST_DIR');

  assert.ok(jsonIndex >= 0 && corsIndex > jsonIndex);
  assert.ok(validatorIndex > corsIndex);
  assert.ok(dispatcherIndex > validatorIndex);
  assert.ok(spaIndex > dispatcherIndex);
});

test('body action parser is registered before validator, root dispatcher, and API routes', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  const jsonIndex = serverSource.indexOf('app.use(express.json({ limit: \'10mb\' }));');
  const validatorIndex = serverSource.indexOf('app.use(rejectUnrecognizedRootPost);');
  const dispatcherIndex = serverSource.indexOf("app.all(['/', '/api', '/api/']");
  const firstApiRouteIndex = serverSource.indexOf("app.post('/api/local/vendor_overrides/upsert'");
  assert.ok(jsonIndex >= 0 && validatorIndex > jsonIndex);
  assert.ok(dispatcherIndex > validatorIndex);
  assert.ok(firstApiRouteIndex > validatorIndex);
});

test('root dispatcher validates only root POST requests', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  const dispatcherStart = serverSource.indexOf("app.all(['/', '/api', '/api/']");
  const dispatcherEnd = serverSource.indexOf('\n});', dispatcherStart);
  assert.ok(dispatcherStart >= 0 && dispatcherEnd > dispatcherStart);
  const dispatcher = serverSource.slice(dispatcherStart, dispatcherEnd);
  assert.match(dispatcher, /req\.method === 'POST' && req\.path === '\/'/);
  assert.doesNotMatch(dispatcher, /req\.method !== 'POST'\s*\)\s*\{\s*res\.status\(405\)/);
});

test('root 405 response follows CORS middleware', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  const corsIndex = serverSource.indexOf('app.use(cors({');
  const guardIndex = serverSource.indexOf("res.status(405).json({ ok: false, error: 'POST to root requires a recognized action' })");
  assert.ok(corsIndex >= 0 && guardIndex > corsIndex);
});

test('root POST validation tests exact path matching, not API prefixes', () => {
  const res = makeResponse();
  let nextCalled = false;
  requestValidator({ method: 'POST', path: '/api/', query: {} }, res, () => {
    nextCalled = true;
  });
  assert.equal(nextCalled, true);
  assert.equal(res.statusCode, null);
});

test('production and preview app origins remain in the CORS allowlist', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  const originsStart = serverSource.indexOf('const allowedOrigins = new Set([');
  const corsStart = serverSource.indexOf('app.use(cors({', originsStart);
  const origins = serverSource.slice(originsStart, corsStart);
  for (const origin of ['https://handymgr.app', 'https://handymgr2.onrender.com']) {
    assert.ok(origins.includes(origin), `${origin} should remain CORS-allowed`);
  }
});

test('CORS allows auth headers and OPTIONS', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  const corsStart = serverSource.indexOf('app.use(cors({');
  const corsEnd = serverSource.indexOf('\n}));', corsStart);
  assert.ok(corsStart >= 0 && corsEnd > corsStart);
  const corsBlock = serverSource.slice(corsStart, corsEnd);
  assert.match(corsBlock, /methods:\s*\[[^\]]*'OPTIONS'/s);
  assert.match(corsBlock, /allowedHeaders:\s*\[[^\]]*'Content-Type'[^\]]*'Authorization'/s);
});

test('Express root dispatcher keeps all and only POST / subject to the root action policy', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  const dispatcherStart = serverSource.indexOf("app.all(['/', '/api', '/api/']");
  const dispatcherEnd = serverSource.indexOf('\n});', dispatcherStart);
  const dispatcher = serverSource.slice(dispatcherStart, dispatcherEnd);
  assert.ok(dispatcherStart >= 0 && dispatcherEnd > dispatcherStart);
  assert.match(dispatcher, /req\.method === 'POST' && req\.path === '\/' && !isAllowedRootPostAction\(action\)/);
  assert.match(dispatcher, /req\.method === 'POST' && req\.path === '\/' && !isRootPostAllowedActionHandled\(action, Object\.keys\(legacyActionRoutes\)\)/);
});

test('dispatcher invokes the shared allowlist-and-handler policy', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  assert.match(serverSource, /isRootPostAllowedActionHandled\(action, Object\.keys\(legacyActionRoutes\)\)/);
});

test('implementation policy returns a classification for each root action state', () => {
  assert.equal(getRootActionImplementationPolicy('not-a-valid-action'), 'rejected');
  assert.equal(getRootActionImplementationPolicy('session_info'), 'inline');
  assert.equal(getRootActionImplementationPolicy('units'), 'unhandled');
  assert.equal(getRootActionImplementationPolicy('units', ['units']), 'route');
});

test('inline action exceptions are backed by dispatcher branches and remain allowlisted', async () => {
  const source = await readFile(new URL('../middleware/requestValidator.js', import.meta.url), 'utf8');
  const start = source.indexOf('const ROOT_ACTIONS_WITH_INLINE_HANDLERS = Object.freeze([');
  const end = source.indexOf('\n]);', start);
  assert.ok(start >= 0 && end > start);
  const actions = Array.from(source.slice(start, end).matchAll(/'([A-Za-z][A-Za-z0-9_]*)'/g), (match) => match[1]);
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  const dispatcherStart = serverSource.indexOf("app.all(['/', '/api', '/api/']");
  const dispatcherEnd = serverSource.indexOf('\n});', dispatcherStart);
  const dispatcher = serverSource.slice(dispatcherStart, dispatcherEnd);

  for (const action of actions) {
    assert.ok(dispatcher.includes(`action === '${action}'`), `${action} must have an explicit dispatcher branch`);
    assert.equal(isAllowedRootPostAction(action), true, `${action} must be allowlisted`);
  }
  for (const action of ROOT_ACTIONS_WITH_INLINE_HANDLERS) {
    assert.ok(actions.includes(action), `${action} is marked inline but has no branch`);
  }
});

test('inline policy actions are explicitly allowed by the master list', async () => {
  const source = await readFile(new URL('../middleware/requestValidator.js', import.meta.url), 'utf8');
  const inlineStart = source.indexOf('const ROOT_ACTIONS_WITH_INLINE_HANDLERS = Object.freeze([');
  const inlineEnd = source.indexOf('\n]);', inlineStart);
  assert.ok(inlineStart >= 0 && inlineEnd > inlineStart);
  const actions = Array.from(source.slice(inlineStart, inlineEnd).matchAll(/'([A-Za-z][A-Za-z0-9_]*)'/g), (match) => match[1]);
  for (const action of actions) assert.ok(ROOT_API_ACTIONS.includes(action), `${action} is missing from master policy`);
});

test('inline-dispatched auth actions still have their compatibility route handlers', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  const routeStart = serverSource.indexOf('const legacyActionRoutes = {');
  const routeEnd = serverSource.indexOf('\n} as const;', routeStart);
  const routeMap = serverSource.slice(routeStart, routeEnd);
  for (const action of ['verify_role', 'device_setup', 'device_otp_request', 'device_otp_verify']) {
    assert.match(routeMap, new RegExp(`^\\s{2}${action}:`, 'm'));
  }
});

test('legacy route-map actions remain in the root action allowlist', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  const routesStart = serverSource.indexOf('const legacyActionRoutes = {');
  const routesEnd = serverSource.indexOf('\n} as const;', routesStart);
  assert.ok(routesStart >= 0 && routesEnd > routesStart);

  const routeMap = serverSource.slice(routesStart, routesEnd);
  for (const match of routeMap.matchAll(/^\s{2}([A-Za-z][A-Za-z0-9_]*):/gm)) {
    assert.equal(isAllowedRootPostAction(match[1]), true, `route '${match[1]}' must be allowed at root`);
    assert.equal(getRootActionImplementationPolicy(match[1], [match[1]]), 'route', `route '${match[1]}' must be classified as routed`);
  }
});

test('each auth legacy handler uses its matching device auth function', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  const routesStart = serverSource.indexOf('const legacyActionRoutes = {');
  const routesEnd = serverSource.indexOf('\n} as const;', routesStart);
  assert.ok(routesStart >= 0 && routesEnd > routesStart);
  const routeMap = serverSource.slice(routesStart, routesEnd);
  const handlers = {
    verify_role: 'handleVerifyRole',
    device_setup: 'handleDeviceSetup',
    device_otp_request: 'handleDeviceOtpRequest',
    device_otp_verify: 'handleDeviceOtpVerify',
  };
  for (const [action, handler] of Object.entries(handlers)) {
    assert.match(routeMap, new RegExp(`${action}: wrapDenoHandler\\(deviceAuthHandlers\\.${handler}`));
  }
});

test('root action allowlist is immutable', () => {
  assert.equal(Object.isFrozen(ROOT_API_ACTIONS), true);
  assert.throws(() => ROOT_API_ACTIONS.push('unexpected_action'), TypeError);
});

test('root action allowlist contains no duplicate compatibility actions', () => {
  assert.equal(new Set(ROOT_API_ACTIONS).size, ROOT_API_ACTIONS.length);
});

test('root POST action matching trims whitespace', () => {
  assert.equal(isAllowedRootPostAction('  verify_role  '), true);
  assert.equal(isAllowedRootPostAction(' verify_role '), true);
  assert.equal(isAllowedRootPostAction('unknown-action'), false);
});

test('root POST action matching is case-sensitive', () => {
  assert.equal(isAllowedRootPostAction('VERIFY_ROLE'), false);
  assert.equal(isAllowedRootPostAction('verify_role'), true);
});

test('nullish/missing action values do not become valid routes', () => {
  for (const action of [undefined, null, '', '   ']) {
    assert.equal(isAllowedRootPostAction(action), false);
  }
});

test('middleware has no broad User-Agent denylist', async () => {
  const middlewareSource = await readFile(new URL('../middleware/requestValidator.js', import.meta.url), 'utf8');
  assert.doesNotMatch(middlewareSource, /go-http-client\|python-requests\|curl\|wget\|nikto\|sqlmap/i);
});

test('resolves query actions before body actions and supports body actions', () => {
  assert.equal(getRootPostAction({ query: { action: 'verify_role' }, body: { action: 'bills' } }), 'verify_role');
  assert.equal(getRootPostAction({ query: { action: ['verify_role', 'bills'] }, body: {} }), 'verify_role,bills');
  assert.equal(getRootPostAction({ query: { action: [] }, body: { action: 'verify_role' } }), 'verify_role');
  assert.equal(getRootPostAction({ query: { action: { value: 'verify_role' } }, body: {} }), '[object Object]');
  assert.equal(getRootPostAction({ query: { action: '' }, body: { action: 'verify_role' } }), 'verify_role');
  assert.equal(getRootPostAction({ query: {}, body: { action: 'verify_role' } }), 'verify_role');
  assert.equal(getRootPostAction({ query: {}, body: { action: 123 } }), '123');
  assert.equal(getRootPostAction({ query: {}, body: { action: false } }), 'false');
  assert.equal(getRootPostAction({ query: {}, body: null }), '');
  assert.equal(getRootPostAction({ query: {}, body: { action: { name: 'verify_role' } } }), '[object Object]');
});

test('does not allow a body action to override an unknown query action', () => {
  const res = makeResponse();
  let nextCalled = false;

  requestValidator({
    method: 'POST',
    path: '/',
    query: { action: 'unknown-action' },
    body: { action: 'verify_role' },
  }, res, () => {
    nextCalled = true;
  });

  assert.equal(res.statusCode, 405);
  assert.equal(res.ended, true);
  assert.equal(nextCalled, false);
});

test('blocks root POST with a body action not present in the action allowlist', () => {
  const res = makeResponse();
  let nextCalled = false;

  requestValidator({ method: 'POST', path: '/', query: {}, body: { action: 'unknown-action' } }, res, () => {
    nextCalled = true;
  });

  assert.equal(res.statusCode, 405);
  assert.equal(res.ended, true);
  assert.equal(nextCalled, false);
});

test('allows a root POST when the query action is empty and the body action is allowed', () => {
  const res = makeResponse();
  let nextCalled = false;

  requestValidator({
    method: 'POST',
    path: '/',
    query: { action: '' },
    body: { action: 'verify_role' },
  }, res, () => {
    nextCalled = true;
  });

  assert.equal(nextCalled, true);
  assert.equal(res.statusCode, null);
  assert.equal(res.ended, false);
});

test('missing root action handler returns JSON 404 before SPA fallback', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  const dispatcherIndex = serverSource.indexOf("app.all(['/', '/api', '/api/']");
  const handlerIndex = serverSource.indexOf('const routeHandler = (legacyActionRoutes as Record<string, any>)[action]', dispatcherIndex);
  const missingRoute404Index = serverSource.indexOf("res.status(404).json({ ok: false, error: `Unknown action '", handlerIndex);
  const staticIndex = serverSource.indexOf('app.use(express.static(DIST_DIR');
  assert.ok(dispatcherIndex >= 0 && handlerIndex > dispatcherIndex);
  assert.ok(missingRoute404Index > handlerIndex && staticIndex > missingRoute404Index);
});

test('backend legacy dispatcher covers recognized root actions', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  const dispatcherStart = serverSource.indexOf("app.all(['/', '/api', '/api/']");
  const dispatcherEnd = serverSource.indexOf('\n});', dispatcherStart);
  assert.ok(dispatcherStart >= 0);
  assert.ok(dispatcherEnd > dispatcherStart);
  assert.match(serverSource.slice(dispatcherStart, dispatcherEnd), /getLegacyAction\(req\)[\s\S]*isAllowedRootPostAction\(action\)/);
  const dispatcher = serverSource.slice(dispatcherStart, dispatcherEnd);
  const actions = new Set();

  for (const match of dispatcher.matchAll(/action\s*===\s*['"]([A-Za-z][A-Za-z0-9_]*)['"]/g)) {
    actions.add(match[1]);
  }
  for (const match of dispatcher.matchAll(/\[(['"])([A-Za-z][A-Za-z0-9_]*)\1((?:\s*,\s*['"][A-Za-z][A-Za-z0-9_]*['"])*)\]\.includes\(action\)/g)) {
    actions.add(match[2]);
    for (const item of match[3].matchAll(/['"]([A-Za-z][A-Za-z0-9_]*)['"]/g)) actions.add(item[1]);
  }
  for (const match of dispatcher.matchAll(/action\s*===\s*['"]([A-Za-z][A-Za-z0-9_]*)['"](?:\s*\|\|\s*action\s*===\s*['"][A-Za-z][A-Za-z0-9_]*['"])+/g)) {
    for (const item of match[0].matchAll(/action\s*===\s*['"]([A-Za-z][A-Za-z0-9_]*)['"]/g)) actions.add(item[1]);
  }

  const routesStart = serverSource.indexOf('const legacyActionRoutes = {', dispatcherEnd);
  const routesEnd = serverSource.indexOf('\n} as const;', routesStart);
  assert.ok(routesStart > dispatcherEnd);
  assert.ok(routesEnd > routesStart);
  const routeMap = serverSource.slice(routesStart, routesEnd);

  assert.ok(actions.size > 0, 'dispatcher actions should be discovered');
  for (const action of ROOT_API_ACTIONS) {
    assert.ok(actions.has(action) || routeMap.includes(`${action}:`), `${action} must have a backend dispatcher branch or handler`);
    assert.equal(isAllowedRootPostAction(action), true, `${action} must remain allowed`);
  }
  for (const action of actions) {
    assert.equal(isAllowedRootPostAction(action), true, `backend root action '${action}' must remain allowlisted`);
  }
  for (const action of ['verify_role', 'device_setup', 'device_otp_request', 'device_otp_verify']) {
    assert.ok(actions.has(action), `${action} must be handled by the dispatcher`);
    assert.match(routeMap, new RegExp(`^\\s{2}${action}: wrapDenoHandler\\(deviceAuthHandlers\\.`, 'm'), `${action} must have a mapped auth handler`);
  }
});

test('backend parses bodies and applies CORS before root POST validation and the SPA fallback', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  const jsonParserIndex = serverSource.indexOf('app.use(express.json({ limit: \'10mb\' }));');
  const corsDefinitionIndex = serverSource.indexOf('const allowedOrigins = new Set([');
  const corsIndex = serverSource.indexOf('app.use(cors({');
  const validatorIndex = serverSource.indexOf('app.use(rejectUnrecognizedRootPost);');
  const dispatcherIndex = serverSource.indexOf("app.all(['/', '/api', '/api/']");
  const staticIndex = serverSource.indexOf('app.use(express.static(DIST_DIR');

  assert.ok(jsonParserIndex >= 0 && corsDefinitionIndex > jsonParserIndex);
  assert.ok(corsIndex > corsDefinitionIndex);
  assert.ok(validatorIndex > corsIndex);
  assert.ok(dispatcherIndex > validatorIndex);
  assert.ok(staticIndex > dispatcherIndex);
});

test('rejected root POST responses are JSON and occur after CORS', async () => {
  const source = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  const corsIndex = source.indexOf('app.use(cors({');
  const rejectionIndex = source.indexOf("res.status(405).json({ ok: false, error: 'POST to root requires a recognized action' })");
  assert.ok(corsIndex >= 0 && rejectionIndex > corsIndex);
  assert.match(source.slice(rejectionIndex - 120, rejectionIndex), /if \(req\.method === 'POST' && req\.path === '\/'/);
});

test('CORS configuration includes the frontend origin used for login', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  const corsConfigStart = serverSource.indexOf('const allowedOrigins = new Set([');
  const corsIndex = serverSource.indexOf('app.use(cors({', corsConfigStart);
  assert.ok(corsConfigStart >= 0 && corsIndex > corsConfigStart);
  assert.match(serverSource.slice(corsConfigStart, corsIndex), /https:\/\/handymgr\.app/);
});

test('preflight requests pass root POST validation', () => {
  const res = makeResponse();
  let nextCalled = false;

  requestValidator({
    method: 'OPTIONS',
    path: '/',
    query: {},
    headers: { origin: 'https://handymgr.app' },
  }, res, () => {
    nextCalled = true;
  });

  assert.equal(nextCalled, true);
  assert.equal(res.statusCode, null);
  assert.equal(res.ended, false);
});

test('root dispatcher applies the shared 405 rule before legacy route lookup', async () => {
  const serverSource = await readFile(new URL('../backend/server.ts', import.meta.url), 'utf8');
  const dispatcherStart = serverSource.indexOf("app.all(['/', '/api', '/api/']");
  const guardIndex = serverSource.indexOf('!isAllowedRootPostAction(action)', dispatcherStart);
  const routeLookupIndex = serverSource.indexOf('(legacyActionRoutes as Record<string, any>)[action]', dispatcherStart);
  const dispatcherEnd = serverSource.indexOf('\n});', dispatcherStart);
  const json405Index = serverSource.indexOf("res.status(405).json({ ok: false, error: 'POST to root requires a recognized action' })", dispatcherStart);

  assert.ok(dispatcherStart >= 0);
  assert.ok(guardIndex > dispatcherStart);
  assert.ok(dispatcherEnd > guardIndex);
  assert.ok(json405Index > guardIndex && json405Index < routeLookupIndex);
  assert.ok(routeLookupIndex > guardIndex);
});

test('allows ordinary browser/API traffic and non-root POST paths', () => {
  for (const req of [
    { method: 'GET', path: '/', headers: { 'user-agent': 'Mozilla/5.0' } },
    { method: 'POST', path: '/api/work-orders', headers: { 'user-agent': 'Mozilla/5.0' } },
    { method: 'GET', path: '/health', headers: {} },
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

test('known scanner probe paths are rejected without requiring a user-agent match', async () => {
  const { isKnownScannerProbePath } = await import('../backend/scannerPathPolicy.ts');
  for (const path of ['/.env', '/.git/config', '/wp-admin', '/xmlrpc.php']) {
    assert.equal(isKnownScannerProbePath(path), true, `${path} should remain blocked by path policy`);
  }
  for (const path of ['/api/verify_role', '/api/local/work_orders', '/health']) {
    assert.equal(isKnownScannerProbePath(path), false, `${path} should remain available`);
  }
});

test('allows query-based actions on supported non-root API paths', () => {
  for (const path of ['/api', '/api/', '/api/local/work_orders']) {
    const res = makeResponse();
    let nextCalled = false;

    requestValidator({
      method: 'POST',
      path,
      query: { action: 'unknown-action' },
      headers: {},
    }, res, () => {
      nextCalled = true;
    });

    assert.equal(nextCalled, true, `${path} should pass through root-only validation`);
    assert.equal(res.statusCode, null);
    assert.equal(res.ended, false);
  }
});

test('passes through API requests with common automated client user-agents', () => {
  for (const userAgent of ['Go-http-client/1.1', 'python-requests/2.32', 'curl/8.5', 'Wget/1.21', 'Nikto/2.5', 'sqlmap/1.8']) {
    const res = makeResponse();
    let nextCalled = false;

    requestValidator({ method: 'GET', path: '/api/health', headers: { 'user-agent': userAgent } }, res, () => {
      nextCalled = true;
    });

    assert.equal(nextCalled, true, `${userAgent} is not blocked solely by its user-agent string`);
    assert.equal(res.statusCode, null);
    assert.equal(res.ended, false);
  }
});