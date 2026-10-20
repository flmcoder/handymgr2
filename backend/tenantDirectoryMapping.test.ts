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

test('upcoming move-outs and turn work orders project nullable canonical property-group UUIDs', async () => {
  const source = await readFile(new URL('./server.ts', import.meta.url), 'utf8');
  const moveoutStart = source.indexOf("app.get('/api/local/upcoming_moveouts'");
  const moveoutEnd = source.indexOf("app.get('/api/local/vacancies'", moveoutStart);
  const moveoutRoute = source.slice(moveoutStart, moveoutEnd);
  const turnWoStart = source.indexOf("app.get('/api/local/turn_work_orders'");
  const turnWoEnd = source.indexOf("app.get('/api/local/turn_records'", turnWoStart);
  const turnWoRoute = source.slice(turnWoStart, turnWoEnd);

  assert.equal((moveoutRoute.match(/as property_group_uuid/g) || []).length, 4);
  assert.equal((moveoutRoute.match(/left join appfolio_property_groups pg/g) || []).length, 4);
  assert.match(moveoutRoute, /left join appfolio_properties p on \$\{propertyJoin\}/);
  assert.equal((moveoutRoute.match(/tenantDirectoryPropertyMatchSql\('p', 'u\.property_id'\)/g) || []).length, 2);
  assert.match(moveoutRoute, /property_group_uuid: row\.property_group_uuid \? String\(row\.property_group_uuid\) : null/);

  assert.match(turnWoRoute, /tenantDirectoryPropertyMatchSql\('p', 'wo\.property_id'\)/);
  assert.match(turnWoRoute, /left join appfolio_property_groups pg on pg\.uuid = p\.property_group_id or pg\.id = p\.property_group_id/);
  assert.match(turnWoRoute, /as property_group_uuid/);
  assert.match(turnWoRoute, /property_group_uuid: row\.property_group_uuid \? String\(row\.property_group_uuid\) : null/);
});

test('Tenant Directory sync requests a unit identity as well as display aliases', async () => {
  const source = await readFile(new URL('./sync/syncRunner.ts', import.meta.url), 'utf8');
  const request = source.slice(source.indexOf("'v2:tenant_directory'"), source.indexOf("'v2:unit_turn_detail'"));
  assert.match(request, /'unit_id'/);
  assert.match(request, /'unit'/);
  assert.match(request, /'property_name'/);
});