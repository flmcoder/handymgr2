import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { mock, test } from 'node:test';
import { postgresErrorDetails, writeWorkOrderWithDiagnostics } from './sync/workOrderUpsertPolicy.ts';

test('single representative AppFolio work-order insert preserves values and nulls blank identifiers', async () => {
  const payload = {
    id: 'f66ef506-c9f7-11ef-bfba-069ca18f5865',
    workOrderUuid: 'f66ef506-c9f7-11ef-bfba-069ca18f5865',
    woNumber: '22055-1',
    status: 'Completed',
    propertyId: 'f4dfb838-9899-11ee-8ab6-12de4bf481cd',
    unitId: '', propertyGroupId: '  ', assignedUserId: '',
    vendorId: '\t', occupancyId: '', requestingTenantId: '',
    rawJson: { AssignedUsers: [], Status: 'Completed' },
  };
  let writes = 0;
  await writeWorkOrderWithDiagnostics(payload, async (values) => {
    writes++;
    assert.equal(values.id, payload.id);
    assert.equal(values.workOrderUuid, payload.workOrderUuid);
    assert.equal(values.propertyId, payload.propertyId);
    for (const field of ['unitId', 'propertyGroupId', 'assignedUserId', 'vendorId', 'occupancyId', 'requestingTenantId']) {
      assert.equal(values[field], null);
    }
    assert.equal(values.status, 'Completed');
    assert.deepEqual(values.rawJson, payload.rawJson);
  });
  assert.equal(writes, 1);
  assert.equal(payload.unitId, '');
});

test('blank UUID and id become null but nonblank legacy text identifiers are preserved', async () => {
  await writeWorkOrderWithDiagnostics({ id: '', workOrderUuid: ' ', propertyId: '6545', totalCost: 0 }, async (values) => {
    assert.deepEqual(values, { id: null, workOrderUuid: null, propertyId: '6545', totalCost: 0 });
  });
});

test('unwraps SQLSTATE, detail and constraint from a nested Drizzle cause', () => {
  const driver = Object.assign(new Error('invalid input syntax for type uuid'), {
    code: '22P02', detail: 'Invalid UUID input', constraint: 'work_order_uuid_check', column: 'work_order_uuid',
  });
  assert.deepEqual(postgresErrorDetails(new Error('Failed query', { cause: driver })), {
    code: '22P02', detail: 'Invalid UUID input', constraint: 'work_order_uuid_check', column: 'work_order_uuid',
  });
  assert.deepEqual(postgresErrorDetails(null), {});
  assert.deepEqual(postgresErrorDetails('connection failed'), {});
  const circular: Record<string, unknown> = { code: 'CONNECT_TIMEOUT' };
  circular.cause = circular;
  assert.deepEqual(postgresErrorDetails(circular), { code: 'CONNECT_TIMEOUT' });
});

test('logs the failed normalized payload and missing-column metadata, redacts credentials, and rethrows', async () => {
  const log = mock.method(console, 'error', () => {});
  try {
    const driver = Object.assign(new Error('missing column'), {
      code: '42703', detail: 'Column is absent', constraint: '', column: 'work_order_uuid',
    });
    const failure = new Error('Failed query', { cause: driver });
    await assert.rejects(writeWorkOrderWithDiagnostics({
      id: 'wo-1', workOrderUuid: '', createdAt: new Date('2026-10-15T00:00:00Z'),
      rawJson: { Status: 'Completed', access_token: 'private', VendorPortalLink: 'private-link', nested: [{ api_key: 'private' }] },
    }, async () => { throw failure; }), (error) => error === failure);
    assert.equal(log.mock.callCount(), 1);
    assert.deepEqual(log.mock.calls[0].arguments, [
      '[work_orders:upsert] Database write failed',
      {
        code: '42703', detail: 'Column is absent', constraint: '', column: 'work_order_uuid',
        payload: {
          id: 'wo-1', workOrderUuid: null, createdAt: '2026-10-15T00:00:00.000Z',
          rawJson: { Status: 'Completed', access_token: '[redacted]', VendorPortalLink: '[redacted]', nested: [{ api_key: '[redacted]' }] },
        },
      },
    ]);
  } finally {
    log.mock.restore();
  }
});

test('work-order repository routes inserts through diagnostic wrapper with sanitized payload logging', async () => {
  const source = await readFile(new URL('./sync/repositories.ts', import.meta.url), 'utf8');
  const start = source.indexOf('export async function upsertWorkOrders(');
  const end = source.indexOf('// ── Unit Inspections', start);
  const block = source.slice(start, end);
  assert.match(block, /writeWorkOrderWithDiagnostics\(/);
  assert.match(block, /\.values\(sanitizedValues\)/);
});

test('sync runner logs database error metadata for failed endpoint runs', async () => {
  const source = await readFile(new URL('./sync/syncRunner.ts', import.meta.url), 'utf8');
  assert.match(source, /postgresErrorDetails\(err\)/);
  assert.match(source, /code:\s*dbError\.code/);
  assert.match(source, /detail:\s*dbError\.detail/);
  assert.match(source, /constraint:\s*dbError\.constraint/);
});
