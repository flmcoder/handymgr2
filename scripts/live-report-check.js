/* HandyManager live report check — paste into DevTools console on handymgr.app
 * while logged in as manager with a property group selected in the global filter.
 * Uses the app's existing session (no password needed). Read-only GETs only.
 */
(async () => {
  const API = 'https://handymgr2.onrender.com';
  const token = localStorage.getItem('hm_auth_token') || localStorage.getItem('hm_device_token') || localStorage.getItem('hm_proxy_token') || localStorage.getItem('hm_access_token') || '';
  const group = (window.HandyMgr && window.HandyMgr.getCurrentPropertyGroup && window.HandyMgr.getCurrentPropertyGroup()) || (document.querySelector('#groupFilter, #globalGroupFilter') || {}).value || 'ALL';
  console.log('Group:', group, '| token:', token ? 'present (' + token.length + ' chars)' : 'MISSING — log in first');
  if (!token) return { ok: false, error: 'No session token. Log in as manager first.' };
  const H = { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token };
  const get = async (p) => {
    const r = await fetch(API + p, { headers: H });
    const t = await r.text();
    let j; try { j = JSON.parse(t); } catch { j = { _raw: t.slice(0, 300) }; }
    return { status: r.status, ...j };
  };
  const days = (d) => d ? Math.floor((Date.now() - new Date(d).getTime()) / 864e5) : null;
  const out = { group, checks: [] };
  const push = (name, res, fn) => { try { out.checks.push({ report: name, status: res.status, ...fn(res) }); } catch (e) { out.checks.push({ report: name, status: res.status, error: String(e.message || e) }); } };

  push('badge_counts', await get('/api/local/badge_counts'), (r) => ({ work_orders: r.work_orders, turns: r.turns, inspections: r.inspections, blank: !r.ok || (r.work_orders == null && r.turns == null) ? 'BLANK/FAIL' : 'ok' }));
  push('work_orders', await get('/api/local/work_orders?limit=100'), (r) => {
    const rows = r.results || []; const over180 = rows.filter(w => days(w.created_at) > 180).length;
    return { count: rows.length, open_180d_plus: over180, blank_fields: rows.filter(w => !w.description && !w.property_id).length, note: 'closure assistant exempt from lease filter' };
  });
  push('inspections', await get('/api/local/inspections?limit=500'), (r) => {
    const rows = r.results || []; const t = new Date().toISOString().slice(0, 10);
    const bad = rows.filter(x => (x.tenant_status && String(x.tenant_status).toLowerCase() !== 'current') || (x.move_in_date && x.move_in_date.slice(0, 10) > t) || (x.move_out_date && x.move_out_date.slice(0, 10) < t));
    return { count: rows.length, non_current_lease_rows: bad.length, blank_tenant: rows.filter(x => !x.tenant_name).length, sample_bad: bad.slice(0, 3) };
  });
  push('turns', await get('/api/local/v2/turns?limit=200'), (r) => {
    const rows = r.results || r.turns || []; const old = rows.filter(x => days(x.move_out_date) > 180).length;
    return { count: rows.length, turns_6mo_plus: old, without_resident: rows.filter(x => x.has_current_resident === false).length, sample_old: rows.filter(x => days(x.move_out_date) > 180).slice(0, 3).map(x => ({ unit: x.unit_name, property: x.property_name, move_out: x.move_out_date })) };
  });
  push('vacancies', await get('/api/local/vacancies'), (r) => {
    const rows = r.results || []; const over90 = rows.filter(v => Number(v.days_vacant ?? v.vacancy_days ?? v.days ?? 0) >= 90);
    return { count: rows.length, vacant_90d_plus: over90.length, sample: over90.slice(0, 5).map(v => ({ p: v.property_name || v.property, u: v.unit || v.unit_name, d: v.days_vacant ?? v.vacancy_days })) };
  });
  push('tenant_directory', await get('/api/local/tenant_directory?limit=100'), (r) => ({ count: (r.results || []).length }));
  push('upcoming_moveouts', await get('/api/local/upcoming_moveouts'), (r) => ({ count: (r.results || []).length }));
  push('bills', await get('/api/local/bills?limit=100'), (r) => {
    const rows = r.results || r.bills || [];
    const pend = rows.filter(b => /pending|not.*approv/i.test(String(b.approval_status || b.status || b.approval || ''))).length;
    return { count: rows.length, pending_or_hold: pend };
  });
  for (const rep of ['guest_cards', 'showings', 'renewal_summary', 'unit_vacancy']) {
    push('v2:' + rep, await get('/api/local/property_performance'), (r) => ({}));
    out.checks.pop(); // property_performance probed once; per-report v2 bodies need POST — check via UI tables below
  }
  // In-memory UI state (no fetch): guest cards / showings / renewals neglect + dashboard alerts
  try {
    const gc = (window.GUEST_CARDS || []).length, sh = (window.SHOWINGS || []).length;
    const wo = (window.WORK_ORDERS || []).filter(w => days(w.createdAt || w.created_at) > 180).length;
    out.checks.push({ report: 'ui-state', guest_cards_loaded: gc, showings_loaded: sh, work_orders_180d_in_memory: wo, renewals_loaded: (window.RENEWALS_ROWS || []).length });
  } catch (e) { out.checks.push({ report: 'ui-state', error: String(e.message || e) }); }
  console.table(out.checks);
  console.log('FULL:', JSON.stringify(out, null, 1));
  return out;
})();
