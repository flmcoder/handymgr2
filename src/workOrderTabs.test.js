import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

// app.js is a single large browser script, so these checks read its source
// text and assert on the functions that own the Open / Completed tab behavior.
const source = await readFile(new URL('./app.js', import.meta.url), 'utf8');

function functionBody(name, nextMarker) {
  const start = source.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `${name} must exist`);
  const end = source.indexOf(nextMarker, start + 1);
  assert.notEqual(end, -1, `${name} end marker must exist`);
  return source.slice(start, end);
}

test('completed work-order history has its own data-source state', () => {
  assert.match(source, /work_orders_inactive: \{ status: 'idle'/);
  assert.match(source, /work_orders_inactive: 'Completed work orders'/);
});

test('fetchInactiveWorkOrders reports loading, success, and failure', () => {
  const body = functionBody('fetchInactiveWorkOrders', '\nfunction normalizeEstimateStatus(');
  assert.match(body, /setDataSourceState\('work_orders_inactive', 'loading'/);
  assert.match(body, /setDataSourceState\('work_orders_inactive', 'ok'/);
  assert.match(body, /setDataSourceState\('work_orders_inactive', 'error'/);
});

test('switching work-order tabs clears a stale status chip', () => {
  const body = functionBody('setWOSubtab', '\nfunction setBillingSubtab(');
  assert.match(body, /if \(currentWOFilter !== 'all'\)/);
  assert.match(body, /currentWOFilter = 'all';/);
});

test('completed tab re-renders even when history load fails', () => {
  const body = functionBody('setWOSubtab', '\nfunction setBillingSubtab(');
  assert.match(body, /fetchInactiveWorkOrders\(\)\.then\(function\(\) \{\s*if \(currentWOSubtab === 'completed'\) renderWorkOrders\(\);/);
});

test('renderWorkOrders error gate uses the active tab source', () => {
  const body = functionBody('renderWorkOrders', '\nfunction ');
  assert.match(body, /workOrderSourceRows = sourceIsInactive \? WORK_ORDERS_INACTIVE : WORK_ORDERS/);
  assert.match(body, /getDataSourceState\(sourceIsInactive \? 'work_orders_inactive' : 'work_orders'\)/);
  assert.match(body, /if \(workOrderSourceRows\.length === 0 &&/);
});

test('retry reloads the source of the current tab', () => {
  const helper = functionBody('retryCurrentWorkOrderSource', '\nfunction renderWorkOrders(');
  assert.match(helper, /currentWOTab === 'inactive' \? fetchInactiveWorkOrders\(\) : fetchWorkOrders\(\{ offset: 0 \}\)/);
  const renderBody = functionBody('renderWorkOrders', '\nfunction ');
  assert.doesNotMatch(renderBody, /fetchWorkOrders\(\{ offset: 0 \}\)\.then/);
});
