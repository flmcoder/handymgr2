import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

import { tenantDirectoryPropertyMatchSql } from './sync/tenantDirectoryPolicy.ts';

test('Tenant Directory and upcoming move-outs use AppFolio numeric property identity joins', async () => {
  const source = await readFile(new URL('./server.ts', import.meta.url), 'utf8');
  const tenantRoute = source.slice(source.indexOf("app.get('/api/local/tenant_directory'"), source.indexOf("app.get('/api/local/upcoming_moveouts'"));
  const moveoutRoute = source.slice(source.indexOf("app.get('/api/local/upcoming_moveouts'"), source.indexOf("app.get('/api/local/" , source.indexOf("app.get('/api/local/upcoming_moveouts'")+20));

  assert.match(tenantRoute, /tenantDirectoryPropertyMatchSql\('p', 't\.property_id'\)/);
  assert.match(moveoutRoute, /tenantDirectoryPropertyMatchSql\('p', 't\.property_id'\)/);
  assert.match(tenantDirectoryPropertyMatchSql(), /raw_json ->> 'PropertyId'/);
  assert.match(tenantDirectoryPropertyMatchSql(), /raw_json ->> 'Link'/);
});

test('Tenant Directory sync requests a unit identity as well as display aliases', async () => {
  const source = await readFile(new URL('./sync/syncRunner.ts', import.meta.url), 'utf8');
  const request = source.slice(source.indexOf("'v2:tenant_directory'"), source.indexOf("'v2:unit_turn_detail'"));
  assert.match(request, /'unit_id'/);
  assert.match(request, /'unit'/);
  assert.match(request, /'property_name'/);
});