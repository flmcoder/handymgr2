import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

import { isRecordInPropertyGroupScope } from './propertyScopePolicy.js';

test('strict property-group scope requires a matching record UUID when active', () => {
  assert.equal(isRecordInPropertyGroupScope({ property_group_uuid: 'Group-A' }, ['group-a']), true);
  assert.equal(isRecordInPropertyGroupScope({ property_group_uuid: 'group-b' }, ['group-a']), false);
  assert.equal(isRecordInPropertyGroupScope({ propertyId: 'numeric-42', property: 'Known property' }, ['group-a']), false);
  assert.equal(isRecordInPropertyGroupScope(null, ['group-a']), false);
});

test('work-order lifecycle status canonicalizer handles AppFolio underscore values', async () => {
  const source = await readFile(new URL('./app.js', import.meta.url), 'utf8');
  const start = source.indexOf('function normalizeWorkOrderDisplayStatus(status)');
  const end = source.indexOf('function normalizeLocalWorkOrder(r)', start);
  const normalizer = source.slice(start, end);
  const aliases = Object.fromEntries(Array.from(normalizer.matchAll(/'([^']+)': '([^']+)'/g), (match) => [match[1], match[2]]));
  const canonicalize = (status) => {
    const normalized = String(status || '').trim().toLowerCase().replace(/[_-]+/g, ' ').replace(/\s+/g, ' ');
    return aliases[normalized] || 'Other';
  };

  assert.equal(canonicalize('in_progress'), 'Assigned');
  assert.equal(canonicalize('estimate_requested'), 'Estimate Requested');
  assert.equal(canonicalize('work_done'), 'Work Done');
  assert.equal(canonicalize('ready_to_bill'), 'Ready to Bill');
});

test('strict property-group scope supports multi-group assignments and unrestricted scope', () => {
  assert.equal(isRecordInPropertyGroupScope({ property_group_uuid: 'group-b' }, ['group-a', 'group-b']), true);
  assert.equal(isRecordInPropertyGroupScope({ property_group_uuid: 'group-b' }, []), true);
  assert.equal(isRecordInPropertyGroupScope({}, []), true);
});

test('Turn Board sources preserve UUIDs and all effective-scope calls pass records', async () => {
  const source = await readFile(new URL('./app.js', import.meta.url), 'utf8');
  const turnStart = source.indexOf('function normalizeTurnRecord(t)');
  const turnEnd = source.indexOf('// Turns: Proxy', turnStart);
  const turnNormalizer = source.slice(turnStart, turnEnd);
  const moveoutStart = source.indexOf('async function fetchUpcomingMoveouts()');
  const moveoutEnd = source.indexOf('// Turn Work Orders:', moveoutStart);
  const moveoutFetcher = source.slice(moveoutStart, moveoutEnd);
  const turnWoStart = source.indexOf('async function fetchTurnWorkOrders()');
  const turnWoEnd = source.indexOf('// Unit Turns (SQL tracker)', turnWoStart);
  const turnWoFetcher = source.slice(turnWoStart, turnWoEnd);
  const pipelineStart = source.indexOf('function buildTurnPipeline()');
  const pipelineEnd = source.indexOf('async function fetchPropertyStats()', pipelineStart);
  const pipeline = source.slice(pipelineStart, pipelineEnd);

  assert.match(turnNormalizer, /property_group_uuid:/);
  assert.match(moveoutFetcher, /property_group_uuid:/);
  assert.match(turnWoFetcher, /property_group_uuid:/);
  assert.match(pipeline, /property_group_uuid:/);
  assert.match(source, /isRecordInPropertyGroupScope\(record, scopeIds\)/);
  assert.doesNotMatch(source, /isInEffectivePropertyScope\((?:p|turn|mo)\.propertyId,\s*(?:p|turn|mo)\.property\)/);
});

test('work-order statuses canonicalize underscored lifecycle values and properties expose fetch state', async () => {
  const source = await readFile(new URL('./app.js', import.meta.url), 'utf8');
  const statusStart = source.indexOf('function normalizeWorkOrderDisplayStatus(status)');
  const statusEnd = source.indexOf('function normalizeLocalWorkOrder(r)', statusStart);
  const statusNormalizer = source.slice(statusStart, statusEnd);
  const propertiesStart = source.indexOf('async function fetchProperties()');
  const propertiesEnd = source.indexOf('function normalizeTurnRecord(t)', propertiesStart);
  const propertiesFetcher = source.slice(propertiesStart, propertiesEnd);

  assert.match(statusNormalizer, /'in progress': 'Assigned'/);
  assert.match(statusNormalizer, /'estimate requested': 'Estimate Requested'/);
  assert.match(statusNormalizer, /'work done': 'Work Done'/);
  assert.match(statusNormalizer, /'ready to bill': 'Ready to Bill'/);
  assert.match(propertiesFetcher, /setDataSourceState\('properties', 'loading'/);
  assert.match(propertiesFetcher, /setDataSourceState\('properties', 'ok'/);
  assert.match(propertiesFetcher, /setDataSourceState\('properties', 'error'/);
});

test('empty failed Work Orders board offers an actionable retry control', async () => {
  const source = await readFile(new URL('./app.js', import.meta.url), 'utf8');
  const start = source.indexOf('function renderWorkOrders()');
  const end = source.indexOf('function bindWorkOrderPaginationControls()', start);
  const renderer = source.slice(start, end);

  assert.match(renderer, /Unable to load work orders:/);
  assert.match(renderer, /id="workOrdersRetryButton"/);
  assert.match(renderer, /fetchWorkOrders\(\{ offset: 0 \}\)\.then\(function\(\) \{ renderWorkOrders\(\); \}\)/);
});

test('Dispatch poll degrades its indicator and records warnings after consecutive failures', async () => {
  const source = await readFile(new URL('./app.js', import.meta.url), 'utf8');
  const start = source.indexOf('_pollLive: async function()');
  const end = source.indexOf('maybeAutoSyncAssignees:', start);
  const poll = source.slice(start, end);

  assert.match(poll, /_livePollFailures>=2/);
  assert.match(poll, /var\(--warning\)/);
  assert.match(poll, /debugLogEvent\('warning',\s*'dispatch_live_poll_degraded'/);
});

test('caught renderer exceptions are reported through console.error', async () => {
  const source = await readFile(new URL('./app.js', import.meta.url), 'utf8');
  const start = source.indexOf('function renderAll()');
  const end = source.indexOf('/* =================================================================\n   INIT', start);
  const renderAll = source.slice(start, end);

  assert.match(renderAll, /catch \(e\) \{ console\.error\('renderAll sub-render failed:', e\); \}/);
  assert.match(source, /catch \(e\) \{\s*console\.error\('renderTurnPipelineUI failed:', e\);\s*\}/);
  assert.match(source, /catch \(e\) \{ console\.error\('renderAttentionPanel failed:', e\); \}/);
});