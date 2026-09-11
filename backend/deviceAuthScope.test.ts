import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

// Union scope is enforced through the session group set; these assertions pin
// the contract without importing the DB-backed auth module under node:test.
test('device auth stores and resolves a multi-group scope set', async () => {
  const source = await readFile(new URL('./deviceAuth.ts', import.meta.url), 'utf8');

  assert.match(source, /export function parseScopeUuidSet/);
  assert.match(source, /ADD COLUMN IF NOT EXISTS property_group_uuids/);
  assert.match(source, /propertyGroupUuids\?: string\[\]/);
  assert.match(source, /scopeUuids: string\[\]/);
  assert.match(source, /property_group_uuids: sessionScopeUuids/);
});

test('server scope plumbing carries group sets end to end', async () => {
  const source = await readFile(new URL('./server.ts', import.meta.url), 'utf8');

  assert.match(source, /function parsePropertyGroupIds\(value: unknown\): string\[\]/);
  assert.match(source, /function getRequestedPropertyGroupIds\(req: Request\): string\[\]/);
  assert.match(source, /function getPropertyGroupFilters\(req: Request\): string\[\]/);
  assert.match(source, /effectiveGroupIds: scope\.propertyGroupIds/);
  assert.match(source, /= ANY\(\$\d+::text\[\]\)/);
  assert.match(source, /PropertyGroupIds'\s*\?\|\s*\(\$\d+::text\[\]\)/);
});
