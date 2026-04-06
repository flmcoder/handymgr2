/* ==============================================================
   MAINTENANCE COCKPIT — AppFolio Live Integration Dashboard
   AES-256-GCM Credential Vault + Rate-Limited API Client
   ============================================================== */

// ---- Dark/Light Mode ----
var _manualTheme = null; // null = follow system, 'dark' or 'light' = manual override
// Light theme by default — manual toggle only (no system detection)
function toggleTheme() {
  var isDark = document.documentElement.classList.contains('dark');
  if (isDark) {
    document.documentElement.classList.remove('dark');
    _manualTheme = 'light';
  } else {
    document.documentElement.classList.add('dark');
    _manualTheme = 'dark';
  }
  updateThemeIcon();
}
function updateThemeIcon() {
  var btn = document.querySelector('#themeToggle');
  if (!btn) return;
  var isDark = document.documentElement.classList.contains('dark');
  btn.innerHTML = isDark ? '<i class="fas fa-sun"></i>' : '<i class="fas fa-moon"></i>';
  btn.title = isDark ? 'Switch to light mode' : 'Switch to dark mode';
}

// ---- Helpers ----
function $(sel) { return document.querySelector(sel); }
function $$(sel) { return document.querySelectorAll(sel); }
function formatDate(d) {
  if (!d) return '—';
  if (typeof d === 'string') { d = new Date(d); }
  if (isNaN(d.getTime())) return '—';
  var m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return m[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
}
function daysBetween(a, b) {
  var da = typeof a === 'string' ? new Date(a) : a;
  var db = typeof b === 'string' ? new Date(b) : b;
  return Math.round(Math.abs(db - da) / 86400000);
}
function currency(n) { return '$' + Number(n || 0).toLocaleString('en-US', { minimumFractionDigits: 0 }); }
function showToast(msg, durationOrOpts) {
  var t = $('#toast');
  var iconEl = $('#toastIcon');
  var msgEl = $('#toastMsg');
  if (!t || !msgEl) return;

  var opts = {};
  if (typeof durationOrOpts === 'number') {
    opts.duration = durationOrOpts;
  } else if (durationOrOpts && typeof durationOrOpts === 'object') {
    opts = durationOrOpts;
  }

  var kind = String(opts.kind || 'info').toLowerCase();
  var kindMap = {
    success: { border: 'var(--success)', icon: 'fa-circle-check' },
    warning: { border: 'var(--warning)', icon: 'fa-triangle-exclamation' },
    danger: { border: 'var(--danger)', icon: 'fa-circle-exclamation' },
    info: { border: 'var(--accent)', icon: 'fa-circle-info' }
  };
  var meta = kindMap[kind] || kindMap.info;

  msgEl.textContent = msg;
  t.style.borderColor = meta.border;
  t.style.display = 'block';

  if (iconEl) {
    var iconClass = opts.iconClass || meta.icon;
    iconEl.innerHTML = '<i class="fas ' + iconClass + '"></i>';
    iconEl.style.color = meta.border;
  }

  clearTimeout(t._tid);
  t._tid = setTimeout(function() {
    t.style.display = 'none';
  }, opts.duration || 3500);
}
function closeModal(id) { document.getElementById(id).classList.remove('show'); }
function openModal(id) { document.getElementById(id).classList.add('show'); }
function escapeHtml(s) { var d = document.createElement('div'); d.textContent = s || ''; return d.innerHTML; }
function loadingHtml(msg) { return '<div class="loading-overlay"><i class="fas fa-circle-notch"></i><p>' + escapeHtml(msg) + '</p></div>'; }
function emptyHtml(icon, msg) { return '<div class="empty-state"><i class="fas ' + icon + '"></i><p>' + escapeHtml(msg) + '</p></div>'; }
function appfolioUrl(type, id) {
  if (!id) return '';
  var base = API_VHOST ? 'https://' + API_VHOST + '.appfolio.com' : '';
  if (!base) return '';
  switch (type) {
    case 'work_order':
      var woNum = String(id);
      // If a UUID was passed (long string), try to resolve to WO number
      if (woNum.length > 20) {
        var match = WORK_ORDERS.find(function(w) { return w.uuid === woNum; });
        if (match && match.id) woNum = String(match.id);
      }
      // Strip hyphen suffix: "12345-1" → "12345"
      woNum = woNum.replace(/-\d+$/, '');
      return base + '/maintenance/service_requests/' + encodeURIComponent(woNum) + '/';
    case 'vendor': return base + '/vendor_details?vendor_id=' + encodeURIComponent(id);
    case 'property': return base + '/properties/' + encodeURIComponent(id) + '/';
    case 'unit_turn': return base + '/maintenance/unit_turns/' + encodeURIComponent(id) + '/';
    case 'inspection_property':
    case 'inspection':
      var propertyId = (typeof id === 'object' && id)
        ? (id.propertyId || id.id || '')
        : String(id || '');
      propertyId = String(propertyId || '').replace(/^p_/i, '');
      return propertyId ? (base + '/maintenance/inspections?filters%5Bproperty_ids_list%5D=p_' + encodeURIComponent(propertyId)) : '';
    case 'tenant': return base + '/tenant_details?occupancy_id=' + encodeURIComponent(id);
    default: return base + '/' + type + '/' + encodeURIComponent(id);
  }
}

// ---- Shared group filter logic ----
// Uses pre-computed UUID-based lookup maps built by resolvePropertyGroupNames.
// Maps are: _nameToGroups (name→groups), _idToGroups (id→groups), _uuidToGroups (uuid→groups).
var _isInGroupMissLogCount = 0; // throttle miss logs
function isInPropertyGroup(propertyId, propertyName, groupName) {
  if (!groupName) return true; // no filter = show all

  // 1. Fast lookup by property name (covers both Reports API and DB API names)
  if (propertyName) {
    var groups = _nameToGroups[String(propertyName).trim().toLowerCase()];
    if (groups && groups.indexOf(groupName) !== -1) return true;
  }

  // 2. Fast lookup by Reports API property_id
  if (propertyId) {
    var idGroups = _idToGroups[String(propertyId)];
    if (idGroups && idGroups.indexOf(groupName) !== -1) return true;
  }

  // 3. Portfolio fallback — check if property's portfolio matches the group name
  if (propertyName || propertyId) {
    var prop = PROPERTIES.find(function(p) {
      if (propertyName && p.name && p.name.trim().toLowerCase() === String(propertyName).trim().toLowerCase()) return true;
      if (propertyId && String(p.id) === String(propertyId)) return true;
      return false;
    });
    if (prop && prop.portfolio) {
      if (prop.portfolio.trim() === groupName.trim() || prop.portfolio.trim().toLowerCase() === groupName.trim().toLowerCase()) return true;
    }
  }

  // Throttled diagnostic: log first 5 misses per filter change to help debug
  if (_isInGroupMissLogCount < 5) {
    _isInGroupMissLogCount++;
    console.log('[PG] isInPropertyGroup MISS: id=' + propertyId + ' name="' + propertyName + '" group="' + groupName + '"');
  }

  return false;
}

// Helper: add a group name to a lookup map entry (creates array if needed)
function _addToGroupMap(map, key, groupName) {
  if (!key) return;
  var k = String(key);
  if (!map[k]) map[k] = [];
  if (map[k].indexOf(groupName) === -1) map[k].push(groupName);
}

// Populates the GLOBAL group filter dropdown (replaces per-tab dropdowns)
function populateGroupFilters() {
  var el = document.getElementById('globalGroupFilter');
  if (!el) return;
  var current = el.value || currentPropertyGroup; // preserve current selection
  // Clear existing options except the first (All Properties)
  while (el.options.length > 1) el.remove(1);

  var addedCount = 0;
  if (PROPERTY_GROUPS.length > 0) {
    // Sort groups alphabetically for better UX
    var sorted = PROPERTY_GROUPS.slice().sort(function(a, b) {
      return (a.name || '').localeCompare(b.name || '');
    });
    sorted.forEach(function(g) {
      if (!g.name) return;
      var opt = document.createElement('option');
      opt.value = g.name;
      // Show resolved property count for feedback
      var pCount = (g.resolvedNames && g.resolvedNames.length) || (g.properties && g.properties.length) || 0;
      opt.textContent = g.name + (pCount > 0 ? ' (' + pCount + ')' : '');
      el.appendChild(opt);
      addedCount++;
    });
    console.log('[PG] populateGroupFilters: ' + addedCount + ' groups from PROPERTY_GROUPS');
  }

  // Fallback: always add portfolio-based groups that aren't already present
  var existingNames = {};
  for (var i = 0; i < el.options.length; i++) {
    existingNames[el.options[i].value.toLowerCase()] = true;
  }
  var grps = {};
  PROPERTIES.forEach(function(p) {
    var candidates = [
      p.portfolio,
      p.portfolioName,
      p.propertyGroup,
      p.group,
      p.groupName,
    ];
    candidates.forEach(function(raw) {
      var pf = String(raw || '').trim();
      if (pf && !existingNames[pf.toLowerCase()]) grps[pf] = true;
    });
  });
  // Secondary fallback: any groups already discovered from UUID/name maps.
  Object.keys(_nameToGroups || {}).forEach(function(k) {
    var groups = _nameToGroups[k] || [];
    groups.forEach(function(g) {
      var gn = String(g || '').trim();
      if (gn && !existingNames[gn.toLowerCase()]) grps[gn] = true;
    });
  });
  var pfKeys = Object.keys(grps).sort();
  if (pfKeys.length > 0) {
    pfKeys.forEach(function(g) {
      var opt = document.createElement('option');
      opt.value = g; opt.textContent = g + ' (portfolio)';
      el.appendChild(opt);
      addedCount++;
    });
    console.log('[PG] populateGroupFilters: added ' + pfKeys.length + ' portfolio fallback groups');
  }

  if (addedCount === 0) {
    console.log('[PG] populateGroupFilters: NO groups available — PROPERTY_GROUPS=' +
      PROPERTY_GROUPS.length + ', PROPERTIES=' + PROPERTIES.length);
  }

  // Restore previous selection if still valid
  if (current) {
    for (var j = 0; j < el.options.length; j++) {
      if (el.options[j].value === current) { el.value = current; break; }
    }
  }
  // Update the active indicator badge
  updateGlobalGroupIndicator();
}

function updateGlobalGroupIndicator() {
  var activeEl = document.getElementById('globalGroupActive');
  var nameEl = document.getElementById('globalGroupName');
  if (!activeEl || !nameEl) return;
  if (currentPropertyGroup) {
    nameEl.textContent = currentPropertyGroup;
    activeEl.style.display = 'inline-flex';
  } else {
    activeEl.style.display = 'none';
  }
}

// ---- Generic Item Detail Card ----
function showItemDetail(title, fields, afLink) {
  var modal = document.getElementById('itemDetailModal');
  if (!modal) return;
  document.getElementById('itemDetailTitle').textContent = title;
  var html = '';
  fields.forEach(function(f) {
    if (f.section) {
      html += '<div class="detail-section-title" style="margin-top:' + (html ? '14px' : '0') + '"><i class="fas ' + (f.icon || 'fa-info-circle') + '"></i> ' + escapeHtml(f.section) + '</div>';
      return;
    }
    html += '<div class="detail-row" style="margin-bottom:8px"><div class="detail-row-label">' + escapeHtml(f.label) + '</div><div class="detail-row-value">' + (f.html || escapeHtml(String(f.value || '\u2014'))) + '</div></div>';
  });
  document.getElementById('itemDetailBody').innerHTML = html;
  var linkBtn = document.getElementById('itemDetailLink');
  if (linkBtn) {
    if (afLink) {
      linkBtn.href = afLink;
      linkBtn.style.display = '';
    } else {
      linkBtn.style.display = 'none';
    }
  }
  openModal('itemDetailModal');
}
function skeletonRows(n) {
  var h = '';
  for (var i = 0; i < n; i++) {
    h += '<div class="skeleton-row">';
    h += '<div class="skeleton-block" style="width:60px"></div>';
    h += '<div class="skeleton-block" style="width:120px"></div>';
    h += '<div class="skeleton-block" style="width:80px"></div>';
    h += '<div class="skeleton-block" style="flex:1"></div>';
    h += '</div>';
  }
  return h;
}
function parseWebhookTs(dateStr) {
  var s = String(dateStr || '').trim();
  if (!s) return null;
  // SQLite datetime('now') often comes back as "YYYY-MM-DD HH:MM:SS" (UTC, no zone).
  // Treat it explicitly as UTC to avoid negative "ago" values in local timezones.
  if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(s)) s = s.replace(' ', 'T') + 'Z';
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function timeAgo(dateStr) {
  var d = parseWebhookTs(dateStr);
  if (!d) return '\u2014';
  var diff = Math.floor((Date.now() - d.getTime()) / 1000);
  if (diff < 0) diff = 0;
  if (diff < 60) return diff + 's ago';
  if (diff < 3600) return Math.floor(diff / 60) + 'm ago';
  if (diff < 86400) return Math.floor(diff / 3600) + 'h ago';
  return Math.floor(diff / 86400) + 'd ago';
}

/* =================================================================
   IndexedDB CACHE — Persists data across sessions
   (localStorage/sessionStorage unavailable in Poe iframe)
   ================================================================= */
var CACHE_DB_NAME = 'maint_cockpit_cache';
var CACHE_DB_VERSION = 3;
var CACHE_STORE = 'api_data';
var CACHE_TTL_MS = 15 * 60 * 1000; // 15 minutes — matches vault timeout
var _cacheDb = null;

function openCacheDB() {
  return new Promise(function(resolve, reject) {
    if (_cacheDb) { resolve(_cacheDb); return; }
    try {
      var req = indexedDB.open(CACHE_DB_NAME, CACHE_DB_VERSION);
      req.onupgradeneeded = function(e) {
        var db = e.target.result;
        if (!db.objectStoreNames.contains(CACHE_STORE)) {
          db.createObjectStore(CACHE_STORE, { keyPath: 'key' });
        }
        if (!db.objectStoreNames.contains('wo_flags')) {
          db.createObjectStore('wo_flags', { keyPath: 'woId' });
        }
        if (!db.objectStoreNames.contains('vendor_overrides')) {
          db.createObjectStore('vendor_overrides', { keyPath: 'vendorId' });
        }
      };
      req.onsuccess = function(e) { _cacheDb = e.target.result; resolve(_cacheDb); };
      req.onerror = function() { reject(new Error('IndexedDB unavailable')); };
    } catch (e) { reject(e); }
  });
}

function cacheGet(key) {
  return openCacheDB().then(function(db) {
    return new Promise(function(resolve) {
      var tx = db.transaction(CACHE_STORE, 'readonly');
      var store = tx.objectStore(CACHE_STORE);
      var req = store.get(key);
      req.onsuccess = function() { resolve(req.result || null); };
      req.onerror = function() { resolve(null); };
    });
  }).catch(function() { return null; });
}

function cacheSet(key, data) {
  return openCacheDB().then(function(db) {
    return new Promise(function(resolve) {
      var tx = db.transaction(CACHE_STORE, 'readwrite');
      tx.objectStore(CACHE_STORE).put({ key: key, data: data, timestamp: Date.now() });
      tx.oncomplete = function() { resolve(); };
      tx.onerror = function() { resolve(); };
    });
  }).catch(function() { /* ignore */ });
}

function isCacheFresh(entry) {
  return entry && entry.timestamp && (Date.now() - entry.timestamp) < CACHE_TTL_MS;
}

function cacheAgeStr(entry) {
  if (!entry || !entry.timestamp) return 'never';
  var diff = Math.floor((Date.now() - entry.timestamp) / 1000);
  if (diff < 10) return 'just now';
  if (diff < 60) return diff + 's ago';
  if (diff < 3600) return Math.floor(diff / 60) + 'm ago';
  return Math.floor(diff / 3600) + 'h ago';
}

async function saveAllToCache() {
  try {
    await Promise.all([
      cacheSet('work_orders', WORK_ORDERS),
      cacheSet('vendors', VENDORS),
      cacheSet('properties', PROPERTIES),
      cacheSet('turns', TURNS),
      cacheSet('inspections', INSPECTIONS),
      cacheSet('webhooks', WEBHOOK_EVENTS)
    ]);
    updateCacheBadge('live', Date.now());
    console.log('Cache saved: WO=' + WORK_ORDERS.length + ' V=' + VENDORS.length + ' P=' + PROPERTIES.length + ' T=' + TURNS.length + ' I=' + INSPECTIONS.length);
  } catch (e) {
    console.log('Cache save failed: ' + (e.message || e));
  }
}

// Export all data as a downloadable JSON file — reads from MEMORY (not IndexedDB)
function exportCacheToJSON() {
  try {
    var counts = {
      work_orders: WORK_ORDERS.length,
      vendors: VENDORS.length,
      properties: PROPERTIES.length,
      turns: TURNS.length,
      inspections: INSPECTIONS.length
    };
    var total = counts.work_orders + counts.vendors + counts.properties + counts.turns + counts.inspections;
    if (total === 0) {
      showToast('Nothing to export \u2014 load data from API first');
      return;
    }
    var exportData = {
      _meta: {
        exported: new Date().toISOString(),
        version: 2,
        dataWindow: DATA_WINDOW_DAYS + ' days',
        counts: counts
      },
      work_orders: WORK_ORDERS,
      vendors: VENDORS,
      properties: PROPERTIES,
      turns: TURNS,
      inspections: INSPECTIONS
    };
    var json = JSON.stringify(exportData);
    var blob = new Blob([json], { type: 'application/json' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = 'maint-cockpit-' + new Date().toISOString().split('T')[0] + '.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function() { URL.revokeObjectURL(url); }, 5000);
    var sizeKB = Math.round(json.length / 1024);
    showToast('Exported ' + total + ' records (' + sizeKB + ' KB) \u2014 WO:' + counts.work_orders + ' V:' + counts.vendors + ' P:' + counts.properties + ' T:' + counts.turns + ' I:' + counts.inspections);
  } catch (e) {
    showToast('Export failed: ' + (e.message || e));
  }
}

// Import data from a JSON file — writes to MEMORY + IndexedDB
async function importCacheFromJSON(file) {
  if (!file) return;
  try {
    var text = await new Promise(function(resolve, reject) {
      var reader = new FileReader();
      reader.onload = function(ev) { resolve(ev.target.result); };
      reader.onerror = function() { reject(new Error('File read error')); };
      reader.readAsText(file);
    });
    var data = JSON.parse(text);
    // Support v2 format (arrays) and v1 format ({data:[], timestamp})
    function extractArr(key) {
      var val = data[key];
      if (!val) return [];
      if (Array.isArray(val)) return val;
      if (val.data && Array.isArray(val.data)) return val.data;
      return [];
    }
    WORK_ORDERS = extractArr('work_orders');
    VENDORS = extractArr('vendors');
    PROPERTIES = extractArr('properties');
    TURNS = extractArr('turns');
    INSPECTIONS = extractArr('inspections');
    var total = WORK_ORDERS.length + VENDORS.length + PROPERTIES.length + TURNS.length + INSPECTIONS.length;
    // Persist to IndexedDB for future sessions
    await saveAllToCache();
    updateCacheBadge('cached', Date.now(), false);
    renderAll();
    showToast('Imported ' + total + ' records — WO:' + WORK_ORDERS.length + ' V:' + VENDORS.length + ' P:' + PROPERTIES.length + ' T:' + TURNS.length + ' I:' + INSPECTIONS.length);
  } catch (e) {
    showToast('Import failed: ' + (e.message || e));
  }
}

function updateCacheBadge(state, timestamp, isStale) {
  var badge = $('#cacheBadge');
  var badgeText = $('#cacheBadgeText');
  var tsEl = $('#syncTimestamp');
  if (!badge || !badgeText) return;

  badge.className = 'cache-badge';
  if (state === 'live') {
    badge.classList.add('live');
    badgeText.textContent = 'LIVE';
    if (tsEl) tsEl.textContent = 'Synced ' + (timestamp ? cacheAgeStr({ timestamp: timestamp }) : 'now');
  } else if (state === 'cached') {
    badge.classList.add(isStale ? 'stale' : 'cached');
    badgeText.textContent = isStale ? 'STALE' : 'CACHED';
    if (tsEl) tsEl.textContent = 'From ' + (timestamp ? cacheAgeStr({ timestamp: timestamp }) : 'cache');
  } else if (state === 'loading') {
    badge.classList.add('cached');
    badgeText.textContent = 'LOADING';
    if (tsEl) tsEl.textContent = '';
  } else {
    badge.classList.add('offline');
    badgeText.textContent = 'OFFLINE';
    if (tsEl) tsEl.textContent = 'No data';
  }
}

async function loadFlags() {
  try {
    var db = await openCacheDB();
    return new Promise(function(resolve) {
      var tx = db.transaction('wo_flags', 'readonly');
      var store = tx.objectStore('wo_flags');
      var req = store.getAll();
      req.onsuccess = function() {
        WO_FLAGS = {};
        (req.result || []).forEach(function(f) { WO_FLAGS[f.woId] = f; });
        resolve();
      };
      req.onerror = function() { resolve(); };
    });
  } catch (e) { /* IndexedDB unavailable */ }
}
async function saveFlag(woId, note) {
  WO_FLAGS[woId] = { woId: woId, note: note || '', ts: Date.now() };
  try {
    var db = await openCacheDB();
    var tx = db.transaction('wo_flags', 'readwrite');
    tx.objectStore('wo_flags').put(WO_FLAGS[woId]);
  } catch (e) { /* best-effort */ }
}
async function removeFlag(woId) {
  delete WO_FLAGS[woId];
  try {
    var db = await openCacheDB();
    var tx = db.transaction('wo_flags', 'readwrite');
    tx.objectStore('wo_flags').delete(woId);
  } catch (e) { /* best-effort */ }
}
async function toggleFlag(woId) {
  if (WO_FLAGS[woId]) { await removeFlag(woId); } else { await saveFlag(woId, ''); }
}
function isWOFlagged(woId) { return !!WO_FLAGS[woId]; }

/* =================================================================
   VENDOR OVERRIDES — IndexedDB-persisted compliance + category
   Stores: { vendorId, compliant (bool|null), category (string) }
   null compliant = use API value; true/false = manual override
   ================================================================= */
var VENDOR_OVERRIDES = {};
var VENDOR_CATEGORIES = ['Employee', 'In-House Tech', 'Vendor', 'Subcontractor', 'Utilities', 'HOA', 'Insurance', 'Uncategorized'];
var VENDOR_TRADE_CATEGORIES = ['General', 'HVAC', 'Plumbing', 'Electrical', 'Appliance', 'Flooring', 'Painting', 'Landscaping', 'Cleaning', 'Roofing', 'Pest Control', 'Pool/Spa', 'Locksmith', 'Security', 'Other'];

function normalizeVendorTradeCategory(value) {
  var raw = String(value || '').trim().toLowerCase();
  if (!raw) return 'General';
  if (raw.indexOf('hvac') !== -1 || raw.indexOf('air') !== -1 || raw.indexOf('ac') === 0) return 'HVAC';
  if (raw.indexOf('plumb') !== -1 || raw.indexOf('drain') !== -1 || raw.indexOf('sewer') !== -1) return 'Plumbing';
  if (raw.indexOf('elect') !== -1 || raw.indexOf('wiring') !== -1) return 'Electrical';
  if (raw.indexOf('appliance') !== -1) return 'Appliance';
  if (raw.indexOf('floor') !== -1 || raw.indexOf('tile') !== -1 || raw.indexOf('carpet') !== -1) return 'Flooring';
  if (raw.indexOf('paint') !== -1 || raw.indexOf('drywall') !== -1) return 'Painting';
  if (raw.indexOf('landscape') !== -1 || raw.indexOf('tree') !== -1 || raw.indexOf('irrig') !== -1) return 'Landscaping';
  if (raw.indexOf('clean') !== -1 || raw.indexOf('janitor') !== -1 || raw.indexOf('maid') !== -1) return 'Cleaning';
  if (raw.indexOf('roof') !== -1) return 'Roofing';
  if (raw.indexOf('pest') !== -1 || raw.indexOf('termite') !== -1) return 'Pest Control';
  if (raw.indexOf('pool') !== -1 || raw.indexOf('spa') !== -1) return 'Pool/Spa';
  if (raw.indexOf('lock') !== -1 || raw.indexOf('key') !== -1) return 'Locksmith';
  if (raw.indexOf('security') !== -1 || raw.indexOf('alarm') !== -1 || raw.indexOf('camera') !== -1) return 'Security';
  if (raw.indexOf('general') !== -1 || raw.indexOf('handyman') !== -1 || raw.indexOf('maintenance') !== -1) return 'General';
  return 'Other';
}

function inferVendorTradeCategory(vendor) {
  var trades = String((vendor && vendor.trades) || '').trim();
  if (trades) {
    var primary = trades.split(',')[0].split('/')[0];
    return normalizeVendorTradeCategory(primary);
  }
  return 'General';
}

async function loadVendorOverrides() {
  // 1. Load from IndexedDB first (fast, local cache)
  try {
    var db = await openCacheDB();
    await new Promise(function(resolve) {
      var tx = db.transaction('vendor_overrides', 'readonly');
      var store = tx.objectStore('vendor_overrides');
      var req = store.getAll();
      req.onsuccess = function() {
        VENDOR_OVERRIDES = {};
        (req.result || []).forEach(function(v) { VENDOR_OVERRIDES[v.vendorId] = v; });
        resolve();
      };
      req.onerror = function() { resolve(); };
    });
  } catch (e) { /* IndexedDB unavailable */ }

  // 2. Layer in SQL overrides (authoritative — survives device change)
  if (API_PROXY) {
    try {
      var sqlData = await proxyAction('vendor_override');
      if (sqlData && sqlData.ok && Array.isArray(sqlData.results)) {
        sqlData.results.forEach(function(row) {
          var vid = row.vendor_id;
          var existing = VENDOR_OVERRIDES[vid] || { vendorId: vid };
          if (row.category !== null && row.category !== undefined) existing.category = row.category;
          if (row.trade_category !== null && row.trade_category !== undefined) existing.tradeCategory = row.trade_category;
          if (row.compliant !== null && row.compliant !== undefined) existing.compliant = !!row.compliant;
          VENDOR_OVERRIDES[vid] = existing;
        });
      }
    } catch (e) { console.warn('vendor_override SQL load failed (non-fatal):', e.message || e); }
  }
}

async function saveVendorOverride(vendorId, overrides) {
  var existing = VENDOR_OVERRIDES[vendorId] || { vendorId: vendorId };
  VENDOR_OVERRIDES[vendorId] = Object.assign(existing, overrides, { ts: Date.now() });
  // Write to IndexedDB (local cache)
  try {
    var db = await openCacheDB();
    var tx = db.transaction('vendor_overrides', 'readwrite');
    tx.objectStore('vendor_overrides').put(VENDOR_OVERRIDES[vendorId]);
  } catch (e) { /* best-effort */ }
  // Write to SQL (authoritative, cross-device)
  if (API_PROXY) {
    var payload = { vendor_id: vendorId };
    if (overrides.category !== undefined) payload.category = overrides.category;
    if (overrides.tradeCategory !== undefined) payload.trade_category = overrides.tradeCategory;
    if (overrides.compliant !== undefined) payload.compliant = overrides.compliant === true ? 1 : (overrides.compliant === false ? 0 : null);
    try { await proxyPost('vendor_override', payload); } catch (e) { /* best-effort */ }
  }
}

function getVendorOverride(vendorId) {
  return VENDOR_OVERRIDES[vendorId] || null;
}

function getVendorCategory(vendorId) {
  var ov = VENDOR_OVERRIDES[vendorId];
  return (ov && ov.category) ? ov.category : '';
}

function getVendorTradeCategory(vendor) {
  if (!vendor) return 'General';
  var ov = VENDOR_OVERRIDES[vendor.id];
  if (ov && ov.tradeCategory) return normalizeVendorTradeCategory(ov.tradeCategory);
  return inferVendorTradeCategory(vendor);
}

function isVendorManuallyCompliant(vendorId) {
  var ov = VENDOR_OVERRIDES[vendorId];
  if (!ov || ov.compliant === null || ov.compliant === undefined) return null; // no override
  return ov.compliant;
}

/* =================================================================
   CREDENTIAL VAULT — AES-256-GCM + PBKDF2
   ================================================================= */
var VAULT_BLOBS = [
  { // Vault profile A
    s: 'oakh6uQFKiJWj95xXH/hJg==',
    i: 'pbU5H33tdyjncIza',
    t: '6jhJXkqaMFtyrXueMQbzhw==',
    c: '4EuCDYImFsxsjPQg65D/hYVji16JZjiEb7IeOf55tUlU9SbBIVNFhrtWS/MDDs2bthUGQl1xQwg6Ds9fX3dW2psUvyMM8FeD62BV7oq9r4ItZHk7Yz/29AuROg/MECEbZhRyzRGRt21c5PNJ1oFih9aR/QmmXKkRo8wvm99Yn+ODsFyCHC15EsFIOzmA288qwA=='
  },
  { // Vault profile B
    s: 'TteV8E8jlLfolw5t39vUiA==',
    i: 'YJYvbSiwahNcaEFW',
    t: 'W4kEMK0/WO8kRJPII6RJ3g==',
    c: '8kD52jM1FFDqORA4amjV5cJgRp6LICgYBqgeP9m7o+IX8XAkxwfa4pFxxU+6Y06xNkeqlL2LZbvYhv4chkydPONGxnKMvtuPurDg43L2QPAf1decHbgWvkcPCDuHzh/mOHH26pdAQjVvrt+RkGqawcpEjI//HsXA/NnUwsHYzU6gL3l1Q+HnZlzu8a+wU8Cnaw=='
  },
  { // Vault profile C (restricted view)
    s: 'h7M5Px4JaZmyG6VElz8WYA==',
    i: 'kSKuoNSunr9A63nU',
    t: '1b14DlE9WJbzFBH+dhQBAA==',
    c: 'rdnAgJHdQrv87rhu1vANkr9ZFIvE8iEFMR1CppCyfx6/evvk31d5CGOg2kt1jXlJqc1xGEWprBc='
  }
];

var API_CREDS = null;
var API_VHOST = null;
var API_PROXY = '';
var _accessRole = 'full'; // 'full' = all tabs, 'manager' = app without dispatch/db admin, 'vendors' = vendor-only restricted access

function decodeSecretLabel(b64) {
  try { return atob(b64); } catch (e) { return ''; }
}
var ROLE_TOKEN_VENDOR = decodeSecretLabel('aGFuZHk6OnZlbmRvcnM=');
var ROLE_TOKEN_MANAGER = decodeSecretLabel('aGFuZHk6Om1hbmFnZXI=');
var ROLE_TOKEN_ADMIN = decodeSecretLabel('YWR2YW5jZWQ6Om1hbmFnZXI=');

function normalizeAccessRole(role) {
  var value = String(role || '').trim().toLowerCase();
  if (value === 'vendors' || value === 'manager' || value === 'full') return value;
  return 'full';
}

function persistAccessRole(role) {
  try { localStorage.setItem('hm_access_role', normalizeAccessRole(role)); } catch (e) { /* */ }
}

function getStoredAccessRole() {
  try { return normalizeAccessRole(localStorage.getItem('hm_access_role') || 'full'); } catch (e) { return 'full'; }
}

function b64ToU8(b64) {
  var bin = atob(b64);
  var u8 = new Uint8Array(bin.length);
  for (var i = 0; i < bin.length; i++) { u8[i] = bin.charCodeAt(i); }
  return u8;
}

async function decryptVaultBlob(blob, passphrase) {
  var enc = new TextEncoder();
  var salt = b64ToU8(blob.s);
  var iv = b64ToU8(blob.i);
  var tag = b64ToU8(blob.t);
  var ciphertext = b64ToU8(blob.c);
  var combined = new Uint8Array(ciphertext.length + tag.length);
  combined.set(ciphertext);
  combined.set(tag, ciphertext.length);
  var keyMaterial = await crypto.subtle.importKey('raw', enc.encode(passphrase), 'PBKDF2', false, ['deriveKey']);
  var aesKey = await crypto.subtle.deriveKey(
    { name: 'PBKDF2', salt: salt, iterations: 150000, hash: 'SHA-256' },
    keyMaterial, { name: 'AES-GCM', length: 256 }, false, ['decrypt']
  );
  var decrypted = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: iv }, aesKey, combined.buffer);
  return JSON.parse(new TextDecoder().decode(decrypted));
}

async function decryptVault(passphrase) {
  // Allow rotating manager/full-admin aliases without re-encrypting vault blobs.
  var normalizedPassphrase = passphrase;
  var requestedRole = 'full';
  if (passphrase === ROLE_TOKEN_ADMIN) {
    normalizedPassphrase = ROLE_TOKEN_MANAGER;
  } else if (passphrase === ROLE_TOKEN_MANAGER) {
    requestedRole = 'manager';
  } else if (passphrase === ROLE_TOKEN_VENDOR) {
    requestedRole = 'vendors';
  }
  // Try each vault blob — supports multiple passphrases
  for (var i = 0; i < VAULT_BLOBS.length; i++) {
    try {
      var result = await decryptVaultBlob(VAULT_BLOBS[i], normalizedPassphrase);
      _accessRole = requestedRole;
      return result;
    } catch (e) { /* try next blob */ }
  }
  throw new Error('Decryption failed for all vault blobs');
}

function isTabAllowedForRole(tabName) {
  if (_accessRole === 'vendors') return tabName === 'vendors';
  if (_accessRole === 'manager') {
    // Manager role — allowed tabs: workorders, turnboard, vendors, inspections, errors
    var allowedTabs = ['dashboard', 'workorders', 'turnboard', 'vendors', 'inspections', 'errors'];
    return allowedTabs.indexOf(tabName) !== -1;
  }
  return true;
}

function forceActiveTab(tabName) {
  $$('.nav-tab').forEach(function(t) {
    t.classList.toggle('active', t.getAttribute('data-tab') === tabName);
  });
  $$('.section').forEach(function(s) {
    s.classList.toggle('active', s.id === 'sec-' + tabName);
  });
}

function wipeCredentials() {
  if (API_CREDS) {
    if (API_CREDS.a) { API_CREDS.a = '0'.repeat(API_CREDS.a.length); }
    if (API_CREDS.d) { API_CREDS.d = '0'.repeat(API_CREDS.d.length); }
  }
  API_CREDS = null;
  API_VHOST = null;
  API_PROXY = '';
  _accessRole = 'full';
}

// ── Auto-sync: selective background refresh every 30 min ───────────────────
var AUTO_SYNC_INTERVAL_MS = 30 * 60 * 1000;

function startAutoSync() {
  stopAutoSync();
  _autoSyncTimer = setInterval(async function() {
    if (!API_CREDS || !API_VHOST) return;
    try {
      await fetchWorkOrders();
      await fetchTurns();
      await fetchInspections();
      renderWorkOrders();
      renderTurnBoard();
      renderInspections($('#inspSearch') ? $('#inspSearch').value : '');
      renderDashboardKPIs();
      var now = new Date();
      var timeStr = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      showToast('Auto-synced at ' + timeStr);
    } catch (err) {
      console.warn('Auto-sync failed:', err.message || err);
    }
  }, AUTO_SYNC_INTERVAL_MS);
}

function stopAutoSync() {
  if (_autoSyncTimer) { clearInterval(_autoSyncTimer); _autoSyncTimer = null; }
}

function lockVault() {
  wipeCredentials();
  appInitialized = false;
  WORK_ORDERS = []; VENDORS = []; PROPERTIES = []; PROPERTY_GROUPS = []; TURNS = []; INSPECTIONS = []; RECENT_TASKS = []; WEBHOOK_EVENTS = []; TURN_RECORDS = []; TURN_PIPE_DATA = []; UNIT_TURNS_DB = []; API_ERRORS = [];
  _nameToGroups = {}; _idToGroups = {}; _uuidToGroups = {};
  detailCacheClear();
  _vendorsLazyLoaded = false; _inspLazyLoaded = false;
  if (_webhookPollTimer) { clearInterval(_webhookPollTimer); _webhookPollTimer = null; }
  stopAutoSync();
  // Note: IndexedDB cache is NOT cleared on lock — data persists for next unlock
  updateCacheBadge('offline');
  $('#appShell').classList.remove('unlocked');
  $('#vaultScreen').style.display = 'flex';
  $('#vaultPassphrase').value = '';
  $('#vaultError').classList.remove('show');
  $('#corsBanner').classList.remove('show');
  // Restore all role-based tab visibility on lock.
  $$('.nav-tab').forEach(function(t) { t.style.display = ''; });
  $$('.section').forEach(function(s) { s.style.display = ''; });
  document.body.classList.remove('role-vendors');
  $('#vaultPassphrase').focus();
}

// Apply access role restrictions — hides tabs/sections not allowed for the role
// Called after unlock + before initApp
function applyAccessRole() {
  if (_accessRole === 'vendors') {
    document.body.classList.add('role-vendors');
    $$('.nav-tab').forEach(function(t) {
      var tabName = t.getAttribute('data-tab');
      t.style.display = isTabAllowedForRole(tabName) ? '' : 'none';
    });
    $$('.section').forEach(function(s) {
      var sectionName = s.id.indexOf('sec-') === 0 ? s.id.substring(4) : '';
      s.style.display = sectionName && !isTabAllowedForRole(sectionName) ? 'none' : '';
    });
    forceActiveTab('vendors');
    // Hide global filter bar in vendor-only mode (no property group filtering needed)
    var gfBar = document.getElementById('globalFilterBar');
    if (gfBar) gfBar.style.display = 'none';
  } else {
    document.body.classList.remove('role-vendors');
    var activeTab = document.querySelector('.nav-tab.active');
    var activeTabName = activeTab ? activeTab.getAttribute('data-tab') : 'dashboard';
    $$('.nav-tab').forEach(function(t) {
      var tabName = t.getAttribute('data-tab');
      var allowed = isTabAllowedForRole(tabName);
      t.style.display = allowed ? '' : 'none';
      if (!allowed) t.classList.remove('active');
    });
    $$('.section').forEach(function(s) {
      var sectionName = s.id.indexOf('sec-') === 0 ? s.id.substring(4) : '';
      var allowed = !sectionName || isTabAllowedForRole(sectionName);
      s.style.display = allowed ? '' : 'none';
      if (!allowed) s.classList.remove('active');
    });
    if (!isTabAllowedForRole(activeTabName)) {
      forceActiveTab('dashboard');
    }
    var gfBar2 = document.getElementById('globalFilterBar');
    if (gfBar2) gfBar2.style.display = '';
  }
  persistAccessRole(_accessRole);
}

function getAuthHeader() { return API_CREDS ? API_CREDS.a : null; }
function getDevId() { return API_CREDS ? API_CREDS.d : null; }
function getDirectBaseUrl() { return API_VHOST ? 'https://' + API_VHOST + '.appfolio.com' : null; }
function getProxyAccessToken() {
  if (API_CREDS && API_CREDS.p) return API_CREDS.p;
  try {
    var deviceToken = localStorage.getItem('hm_device_token') || '';
    if (deviceToken) return deviceToken;
  } catch (e) { /* */ }
  try { return localStorage.getItem('hm_proxy_token') || ''; } catch (e) { return ''; }
}

async function setupTrustedDevice(setupPin, userName) {
  if (!API_PROXY) throw new Error('No proxy configured');
  var sep = API_PROXY.indexOf('?') !== -1 ? '&' : '?';
  var url = API_PROXY + sep + 'action=device_setup';
  var payload = {
    pin: setupPin,
    user_name: userName || ('hm-' + (navigator && navigator.platform ? navigator.platform : 'device'))
  };
  var res = await fetchWithTimeout(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
    body: JSON.stringify(payload)
  }, 30000);
  var data = {};
  try { data = await res.json(); } catch (e) { /* */ }
  if (!res.ok || !data.ok || !data.token) {
    throw new Error(data.error || ('Device setup failed (HTTP ' + res.status + ')'));
  }
  return data.token;
}

function normalizeOtpEmail(raw) {
  var email = String(raw || '').trim().toLowerCase();
  if (!email) return '';
  var m = email.match(/^[a-z0-9._%+\-]+@flraz\.com$/i);
  return m ? email : '';
}

async function requestDeviceOtp(email, userName) {
  if (!API_PROXY) throw new Error('No proxy configured');
  var sep = API_PROXY.indexOf('?') !== -1 ? '&' : '?';
  var url = API_PROXY + sep + 'action=device_otp_request';
  var payload = {
    email: email,
    user_name: userName || ('hm-' + (navigator && navigator.platform ? navigator.platform : 'device'))
  };
  var res = await fetchWithTimeout(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
    body: JSON.stringify(payload)
  }, 30000);
  var data = {};
  try { data = await res.json(); } catch (e) { /* */ }
  if (!res.ok || !data.ok) {
    throw new Error(data.error || ('OTP request failed (HTTP ' + res.status + ')'));
  }
  return data;
}

async function verifyDeviceOtp(email, code, userName) {
  if (!API_PROXY) throw new Error('No proxy configured');
  var sep = API_PROXY.indexOf('?') !== -1 ? '&' : '?';
  var url = API_PROXY + sep + 'action=device_otp_verify';
  var payload = {
    email: email,
    code: code,
    user_name: userName || ('hm-' + (navigator && navigator.platform ? navigator.platform : 'device'))
  };
  var res = await fetchWithTimeout(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
    body: JSON.stringify(payload)
  }, 30000);
  var data = {};
  try { data = await res.json(); } catch (e) { /* */ }
  if (!res.ok || !data.ok || !data.token) {
    throw new Error(data.error || ('OTP verify failed (HTTP ' + res.status + ')'));
  }
  return data.token;
}

// ---- Timeout helper ----
// Wraps a promise with an AbortController-based timeout (defaults 45s)
function fetchWithTimeout(url, opts, timeoutMs) {
  timeoutMs = timeoutMs || 45000;
  var controller = new AbortController();
  var timer = setTimeout(function() { controller.abort(); }, timeoutMs);
  var fetchOpts = Object.assign({}, opts || {}, { signal: controller.signal });
  return fetch(url, fetchOpts).finally(function() { clearTimeout(timer); });
}

// ---- Proxy action endpoint caller ----
// Makes ONE request to proxy like ?action=work_orders&days=180
// Proxy does all pagination server-side and returns complete dataset
// Includes 45-second timeout — never hangs forever
async function proxyAction(action, params) {
  if (!API_PROXY) throw new Error('No proxy configured');
  var sep = API_PROXY.indexOf('?') !== -1 ? '&' : '?';
  var url = API_PROXY + sep + 'action=' + encodeURIComponent(action);
  if (params) {
    Object.keys(params).forEach(function(k) {
      url += '&' + encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
    });
  }
  var maxRetries = 2; // retry once for transient 502/503/network errors
  var timeoutMs = action === 'turn_work_orders' ? 70000 : 45000;
  var token = getProxyAccessToken();
  var reqHeaders = { 'Accept': 'application/json' };
  if (token) reqHeaders['Authorization'] = 'Bearer ' + token;
  for (var attempt = 0; attempt <= maxRetries; attempt++) {
    var res;
    try {
      res = await fetchWithTimeout(url, { headers: reqHeaders }, timeoutMs);
    } catch (abortErr) {
      if (abortErr.name === 'AbortError') {
        var tmsg = 'Proxy action=' + action + ' timed out after ' + Math.round(timeoutMs / 1000) + 's';
        logApiError(0, tmsg, 'queued');
        throw new Error(tmsg);
      }
      // Network error (CORS, DNS, connection refused) — retry once after backoff
      if (attempt < maxRetries) {
        var netWait = Math.pow(2, attempt + 1) * 1000; // 2s, 4s
        logApiError(0, 'Proxy action=' + action + ' network error (attempt ' + (attempt + 1) + '/' + (maxRetries + 1) + ') — retrying in ' + (netWait / 1000) + 's', 'retry');
        await sleep(netWait);
        continue;
      }
      throw abortErr;
    }
    // Retryable server errors: 502, 503, 504
    if ((res.status === 502 || res.status === 503 || res.status === 504) && attempt < maxRetries) {
      var backoff = Math.pow(2, attempt + 1) * 1000; // 2s, 4s
      logApiError(res.status, 'Proxy action=' + action + ' HTTP ' + res.status + ' (attempt ' + (attempt + 1) + '/' + (maxRetries + 1) + ') — retrying in ' + (backoff / 1000) + 's', 'retry');
      await sleep(backoff);
      continue;
    }
    if (!res.ok) {
      if (action === 'wo_notes' && res.status === 404) {
        return { ok: true, results: [], warning: 'wo_notes not found' };
      }
      var errBody = '';
      try { errBody = await res.text(); } catch (e) { /* empty */ }
      var errMsg = 'Proxy action=' + action + ' failed: HTTP ' + res.status;
      if (errBody) {
        try { var ej = JSON.parse(errBody); if (ej.error) errMsg += ' \u2014 ' + ej.error; } catch (e) { errMsg += ' \u2014 ' + errBody.substring(0, 200); }
      }
      logApiError(res.status, errMsg, 'queued');
      throw new Error(errMsg);
    }
    var data = await res.json();
    if (data && data.ok === false) {
      var dataStatus = parseInt(data.status || '0', 10) || 0;
      var isRetryable = dataStatus === 0 || dataStatus === 429 || dataStatus === 502 || dataStatus === 503 || dataStatus === 504 || dataStatus === 533;
      if (action === 'wo_notes' && dataStatus === 404) {
        return { ok: true, results: [], warning: data.error || 'wo_notes not found' };
      }
      if (attempt < maxRetries && isRetryable) {
        var retryWait = Math.pow(2, attempt + 1) * 1000;
        logApiError(dataStatus || 502, 'Proxy action=' + action + ': ' + (data.error || 'Unknown') + ' — retrying in ' + (retryWait / 1000) + 's', 'retry');
        await sleep(retryWait);
        continue;
      }
      var msg = 'Proxy action=' + action + ': ' + (data.error || 'Unknown error');
      logApiError(dataStatus || 502, msg, dataStatus === 404 ? 'resolved' : 'queued');
      throw new Error(msg);
    }
    return data;
  }
}

// POST helper for sensitive admin actions where secrets must NOT appear in the
// URL (preventing exposure in server logs, browser history, CDN traces).
// Sends { action } in the query string only; all payload (including key/secret)
// travels as a JSON body. Attaches the same bearer token as proxyAction.
async function proxyPost(action, bodyObj, extraHeaders) {
  if (!API_PROXY) throw new Error('No proxy configured');
  var sep = API_PROXY.indexOf('?') !== -1 ? '&' : '?';
  var url = API_PROXY + sep + 'action=' + encodeURIComponent(action);
  var token = getProxyAccessToken();
  var headers = Object.assign({ 'Content-Type': 'application/json', 'Accept': 'application/json' }, extraHeaders || {});
  if (token) headers['Authorization'] = 'Bearer ' + token;
  var res = await fetch(url, {
    method: 'POST',
    headers: headers,
    body: JSON.stringify(bodyObj || {})
  });
  if (!res.ok) {
    var errBody = '';
    try { errBody = await res.text(); } catch (e) { /* empty */ }
    throw new Error('Proxy POST action=' + action + ' failed: HTTP ' + res.status + (errBody ? ' \u2014 ' + errBody.substring(0, 200) : ''));
  }
  return res.json();
}

// Resolve a path to a fetchable URL.
// When a proxy is active, the proxy has the domain + credentials hardcoded
// server-side, so we only send the API path (e.g. /api/v0/properties).
// Used for raw pass-through calls (PATCH work orders, POST notes, etc.)
function resolveUrl(path, method) {
  method = (method || 'GET').toUpperCase();
  if (API_PROXY) {
    // Server-side proxy: send only the API path, not the full URL
    var apiPath = path;
    if (path.indexOf('http') === 0) {
      // Pagination next_page_path may come back as absolute or relative — extract path+query
      try {
        var u = new URL(path);
        apiPath = u.pathname + u.search;
      } catch (e) { /* use as-is */ }
    }
    var sep = API_PROXY.indexOf('?') !== -1 ? '&' : '?';
    if (method !== 'GET' && method !== 'HEAD') {
      return API_PROXY + sep + 'action=passthrough&path=' + encodeURIComponent(apiPath);
    }
    return API_PROXY + sep + 'path=' + encodeURIComponent(apiPath);
  }
  // Direct connection (no proxy) — build full URL
  if (path.indexOf('http') === 0) return path;
  var direct = getDirectBaseUrl();
  if (!direct) return path;
  return direct + path;
}

/* =================================================================
   AppFolio API Client — Rate-limited with retry logic
   ================================================================= */
var rateLimiter = {
  queue: [],
  inFlight: 0,
  maxPerSec: 4,
  windowStart: 0,
  windowCount: 0,
  processing: false,

  enqueue: function(fn) {
    return new Promise(function(resolve, reject) {
      rateLimiter.queue.push({ fn: fn, resolve: resolve, reject: reject });
      rateLimiter.process();
    });
  },

  process: function() {
    if (rateLimiter.processing) return;
    rateLimiter.processing = true;

    (function tick() {
      if (rateLimiter.queue.length === 0) {
        rateLimiter.processing = false;
        return;
      }

      var now = Date.now();
      if (now - rateLimiter.windowStart >= 1000) {
        rateLimiter.windowStart = now;
        rateLimiter.windowCount = 0;
      }

      if (rateLimiter.windowCount >= rateLimiter.maxPerSec) {
        var wait = 1000 - (now - rateLimiter.windowStart) + 10;
        setTimeout(tick, wait);
        return;
      }

      var item = rateLimiter.queue.shift();
      rateLimiter.windowCount++;
      rateLimiter.inFlight++;
      updateRateBadge();

      item.fn().then(function(r) {
        rateLimiter.inFlight--;
        updateRateBadge();
        item.resolve(r);
        tick();
      }).catch(function(e) {
        rateLimiter.inFlight--;
        updateRateBadge();
        item.reject(e);
        tick();
      });
    })();
  }
};

function updateRateBadge() {
  var el = $('#rateBadge');
  if (el) { el.textContent = (rateLimiter.maxPerSec - rateLimiter.inFlight) + '/' + rateLimiter.maxPerSec + ' req/s'; }
}

// Core fetch wrapper with auth, retries, and error logging
async function apiFetch(path, options) {
  if (!API_PROXY) {
    // Direct mode needs vault credentials
    var auth = getAuthHeader();
    var devId = getDevId();
    if (!getDirectBaseUrl() || !auth || !devId) { throw new Error('Vault locked or missing config'); }
  }

  var hdrs = {};
  var opts = Object.assign({ method: 'GET' }, options || {});
  if (API_PROXY) {
    // Server-side proxy has credentials hardcoded — no auth headers needed
    hdrs['Content-Type'] = 'application/json';
    var proxyToken = getProxyAccessToken();
    if (proxyToken) hdrs['Authorization'] = 'Bearer ' + proxyToken;
  } else {
    hdrs['Authorization'] = auth;
    hdrs['X-AppFolio-Developer-ID'] = devId;
    hdrs['Content-Type'] = 'application/json';
  }
  opts.headers = Object.assign({}, hdrs, opts.headers || {});
  var url = resolveUrl(path, opts.method);

  // If POST with form params (supports array values for Reports API filters)
  if (opts.formParams) {
    opts.headers['Content-Type'] = 'application/x-www-form-urlencoded';
    var fparams = new URLSearchParams();
    Object.keys(opts.formParams).forEach(function(k) {
      var val = opts.formParams[k];
      if (Array.isArray(val)) {
        val.forEach(function(v) { fparams.append(k + '[]', v); });
      } else {
        fparams.append(k, val);
      }
    });
    opts.body = fparams.toString();
    delete opts.formParams;
  }

  var maxRetries = 3;
  for (var attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      var res = await rateLimiter.enqueue(function() { return fetchWithTimeout(url, opts, 30000); });

      // --- Retryable errors (429 rate limit, 533 DB unavailable) ---
      if (res.status === 429) {
        var retryAfter = parseInt(res.headers.get('Retry-After') || '5', 10);
        logApiError(429, 'Rate limit exceeded — Retry-After: ' + retryAfter + 's', 'retry');
        if (attempt < maxRetries) { await sleep(retryAfter * 1000); continue; }
        throw new Error('429: Rate limited after retries');
      }
      if (res.status === 533) {
        var backoff = Math.pow(2, attempt + 2) * 1000; // 4s, 8s, 16s, 32s
        logApiError(533, 'Database busy — waiting ' + (backoff / 1000) + 's before retry (' + (attempt + 1) + '/' + maxRetries + ')', 'retry');
        if (attempt < maxRetries) { await sleep(backoff); continue; }
        throw new Error('533: Database unavailable after ' + maxRetries + ' retries');
      }

      // --- Non-retryable errors (throw immediately, no retry) ---
      if (res.status === 401) {
        logApiError(401, 'Unauthorized — proxy may have wrong credentials. Verify the deployed proxy has the correct Client ID, Client Secret, and Developer ID.', 'resolved');
        showCorsError('401 Unauthorized — AppFolio rejected the credentials. Verify the deployed proxy has the correct Client ID, Client Secret, and Developer ID configured server-side.');
        throw new Error('401: Unauthorized — check proxy credentials');
      }
      if (res.status === 404) {
        logApiError(404, 'Not Found — endpoint does not exist or credentials are invalid. URL: ' + path, 'resolved');
        showCorsError('404 Not Found — The endpoint may not exist or credentials may be rejected. Verify the deployed proxy URL and server-side credentials.');
        throw new Error('404: Not Found — ' + path);
      }
      if (res.status === 422) {
        var body422 = await res.text();
        logApiError(422, 'Semantic error: ' + body422.substring(0, 200), 'resolved');
        throw new Error('422: ' + body422);
      }
      if (res.status === 400) {
        var body400 = await res.text();
        logApiError(400, 'Bad request: ' + body400.substring(0, 200), 'resolved');
        throw new Error('400: ' + body400);
      }
      if (res.status === 403) {
        var body403 = await res.text();
        logApiError(403, 'Forbidden: ' + body403.substring(0, 200), 'resolved');
        throw new Error('403: Forbidden — ' + body403);
      }
      if (res.status === 526) {
        logApiError(526, 'Invalid SSL — check subdomain is correct.', 'resolved');
        throw new Error('526: Invalid SSL — verify subdomain');
      }
      if (!res.ok) {
        var bodyErr = '';
        try { bodyErr = await res.text(); } catch (ignored) { /* empty */ }
        logApiError(res.status, 'HTTP ' + res.status + ': ' + (bodyErr || res.statusText).substring(0, 200), 'queued');
        throw new Error('HTTP ' + res.status + ': ' + (bodyErr || res.statusText));
      }

      return await res.json();
    } catch (err) {
      // Timeout errors — abort immediately, no retry
      if (err.name === 'AbortError') {
        logApiError(0, 'Request timed out (30s): ' + path, 'queued');
        throw new Error('Request timed out: ' + path);
      }
      // Network errors (CORS, CSP, DNS, SSL, connection refused)
      if (err.name === 'TypeError') {
        var netMsg = err.message || 'Network request failed';
        var isCsp = netMsg.indexOf('Content Security Policy') !== -1 || netMsg.indexOf('Refused to connect') !== -1;
        var isSsl = netMsg.indexOf('SSL') !== -1 || netMsg.indexOf('ERR_SSL') !== -1 || netMsg.indexOf('ERR_CERT') !== -1;
        if (isCsp) {
          logApiError(0, 'CSP BLOCKED: ' + netMsg + '. Click "Allow additional resources" popup at top of page.', 'queued');
        } else if (isSsl) {
          logApiError(0, 'SSL ERROR: ' + netMsg + '. Your worker SSL cert may not be ready yet (wait 2-3 min) or the URL is wrong.', 'queued');
        } else {
          logApiError(0, 'Network error: ' + netMsg + ' — URL: ' + path, 'queued');
        }
        showCorsError(netMsg);
        throw err;
      }
      // Non-retryable errors already thrown above; only retryable reach here
      if (attempt === maxRetries) throw err;
    }
  }
}

function sleep(ms) { return new Promise(function(r) { setTimeout(r, ms); }); }

// Paginated fetch for Database API v0 (follows next_page_path)
// Uses small page sizes to avoid 533 "Database unavailable" errors
var FETCH_PAGE_SIZE = 50; // Database API page size (conservative — AppFolio default 1000 triggers 533)

async function fetchAllPages(path, maxRecords) {
  var allResults = [];
  // Ensure page[size] is set on the initial request
  var currentPath = path;
  if (currentPath.indexOf('page[size]') === -1) {
    var joiner = currentPath.indexOf('?') !== -1 ? '&' : '?';
    currentPath = currentPath + joiner + 'page[size]=' + FETCH_PAGE_SIZE;
  }
  var pageCount = 0;
  while (currentPath && pageCount < 200) {
    var data = await apiFetch(currentPath);
    var pageItems = 0;
    // API v0 wraps list results in { data: [...], next_page_path: "..." }
    if (data && Array.isArray(data.data)) {
      allResults = allResults.concat(data.data);
      pageItems = data.data.length;
    } else if (data && data.results) {
      // Reports API fallback
      allResults = allResults.concat(data.results);
      pageItems = data.results.length;
    } else if (Array.isArray(data)) {
      allResults = allResults.concat(data);
      pageItems = data.length;
    }
    // Enforce max records limit (e.g. 100 WOs to reduce load time)
    if (maxRecords && allResults.length >= maxRecords) {
      allResults = allResults.slice(0, maxRecords);
      setApiStatus('loading', 'Reached ' + maxRecords + ' record limit \u2014 done');
      break;
    }
    // next_page_path is a relative path like /api/v0/work_orders?page[number]=2
    currentPath = (data && data.next_page_path) ? data.next_page_path : null;
    pageCount++;
    // Throttle between pages to reduce DB pressure
    if (currentPath) { await sleep(300); }
    // Update status with progress
    setApiStatus('loading', 'Loaded ' + allResults.length + ' records (page ' + pageCount + ')\u2026');
  }
  return allResults;
}

// Reports API v2 fetch (POST initial, GET pagination) — 5000 rows/page
// Pagination next_page_url is NOT rate-limited and valid for 30 minutes
async function fetchReport(reportName, filters) {
  var path = '/api/v2/reports/' + reportName + '.json';
  var allRows = [];
  var formParams = {};
  if (filters) {
    Object.keys(filters).forEach(function(k) {
      formParams['filters[' + k + ']'] = filters[k];
    });
  }

  var nextUrl = null;
  var pageCount = 0;
  while (pageCount < 50) {
    var data;
    if (nextUrl) {
      // Pagination pages are GET requests (no filters needed, cached server-side)
      data = await apiFetch(nextUrl);
    } else {
      // Initial request is POST with filters
      data = await apiFetch(path, { method: 'POST', formParams: formParams });
    }
    if (data && data.results) {
      allRows = allRows.concat(data.results);
    } else if (data && Array.isArray(data.data)) {
      allRows = allRows.concat(data.data);
    } else if (Array.isArray(data)) {
      // paginate_results=false returns a raw array
      allRows = allRows.concat(data);
      break;
    }
    setApiStatus('loading', reportName + ': ' + allRows.length + ' rows (page ' + (pageCount + 1) + ')');
    nextUrl = (data && (data.next_page_url || data.next_page_path)) ? (data.next_page_url || data.next_page_path) : null;
    if (!nextUrl) break;
    pageCount++;
    // Pagination pages are NOT rate-limited, but brief pause to be polite
    await sleep(100);
  }
  return allRows;
}

function showCorsError(detail) {
  var banner = $('#corsBanner');
  var detailEl = $('#corsDetail');
  var msgEl = $('#corsMsg');
  var detailStr = detail || 'Unknown failure';
  var isAuthStripped = detailStr.indexOf('401') !== -1 || detailStr.indexOf('404') !== -1 || detailStr.indexOf('strips auth') !== -1;
  var isCsp = detailStr.indexOf('Content Security Policy') !== -1 || detailStr.indexOf('Refused to connect') !== -1;
  var isSsl = detailStr.indexOf('SSL') !== -1 || detailStr.indexOf('ERR_SSL') !== -1 || detailStr.indexOf('ERR_CERT') !== -1;
  if (isCsp) {
    if (msgEl) { msgEl.innerHTML = '<strong style="color:var(--warning)">CSP Blocked.</strong> The Poe iframe blocks connections to external domains by default. Click the <strong style="color:var(--accent)">"Allow additional resources"</strong> popup at the top of the page. The page will reload — then re-enter your passphrase, subdomain, and proxy URL to connect again.'; }
    setApiStatus('error', 'CSP Blocked — Click Allow');
  } else if (isSsl) {
    if (msgEl) { msgEl.innerHTML = '<strong style="color:var(--warning)">SSL Certificate Error.</strong> Your proxy\'s SSL certificate may not be provisioned yet. New Val Town endpoints can take <strong>1-2 minutes</strong> for SSL to activate. Wait and try again, or verify the proxy URL is correct.'; }
    setApiStatus('error', 'SSL Error — Wait & Retry');
  } else if (isAuthStripped) {
    if (msgEl) { msgEl.innerHTML = '<strong style="color:var(--danger)">AppFolio rejected credentials.</strong> Verify your deployed proxy has the correct AppFolio Client ID, Client Secret, and Developer ID configured server-side. Visit the proxy URL in a browser and confirm the JSON response shows <code>"service": "HandyManager Proxy"</code> and a current <code>"version"</code> value.'; }
    setApiStatus('error', 'Auth Stripped by Proxy');
  } else if (API_PROXY) {
    if (msgEl) { msgEl.textContent = 'The proxy (' + API_PROXY.substring(0, 50) + ') failed. It may be down, rate-limited, or CSP-blocked.'; }
    setApiStatus('error', 'Proxy Error');
  } else {
    if (msgEl) { msgEl.textContent = 'Direct browser requests blocked (no CORS headers). You need a proxy — lock vault and configure one.'; }
    setApiStatus('error', 'CORS Blocked');
  }
  if (detailEl) { detailEl.textContent = detailStr; }
  banner.classList.add('show');
}

function setApiStatus(state, text) {
  var el = $('#apiStatus');
  var textEl = $('#apiStatusText');
  el.className = 'topbar-status ' + state;
  textEl.textContent = text;
}

var APP_VERSION = 'v9.1a';
var APP_VERSION_UPDATED = 'April 2026';
var SERVER_VERSION = '';

function applyVersionBadge(serverVersion) {
  var badge = document.getElementById('versionBadge');
  if (!badge) return;
  var txt = APP_VERSION + ' (updated ' + APP_VERSION_UPDATED + ')';
  if (serverVersion && String(serverVersion).trim() && String(serverVersion).trim() !== APP_VERSION) {
    txt += ' \u2022 server ' + String(serverVersion).trim();
  }
  txt += ' \u2022 Highlights (next 36h)';
  badge.textContent = txt;
}

function enforceServerVersionGuard(pingData) {
  var serverVersion = String((pingData && (pingData.version || pingData.proxy)) || '').trim();
  if (!serverVersion) {
    applyVersionBadge('');
    return false;
  }

  SERVER_VERSION = serverVersion;
  applyVersionBadge(serverVersion);
  if (serverVersion === APP_VERSION) return false;

  var seenKey = 'hm_force_refresh_seen_' + serverVersion;
  try {
    if (sessionStorage.getItem(seenKey) !== '1') {
      sessionStorage.setItem(seenKey, '1');
      alert('A newer HandyManager build (' + serverVersion + ') is available. The app will refresh now.');
      window.location.reload();
      return true;
    }
  } catch (_) {
    alert('A newer HandyManager build (' + serverVersion + ') is available. Please refresh now.');
    window.location.reload();
    return true;
  }

  setApiStatus('error', 'Update required \u2014 refresh the app');
  showToast('New build detected (' + serverVersion + '). Please hard refresh.', {
    kind: 'warning', iconClass: 'fa-triangle-exclamation', duration: 7000,
  });
  return true;
}

/* =================================================================
   API Error Log
   ================================================================= */
var API_ERRORS = [];

function logApiError(code, msg, action) {
  var now = new Date();
  var ts = now.getHours().toString().padStart(2, '0') + ':' + now.getMinutes().toString().padStart(2, '0') + ':' + now.getSeconds().toString().padStart(2, '0');
  API_ERRORS.unshift({ code: code, ts: ts, msg: msg, action: action });
  if (API_ERRORS.length > 100) API_ERRORS.length = 100;
  renderErrorLog();
}

/* =================================================================
   DATA STORES — populated from live API
   ================================================================= */
var WORK_ORDERS = [];
var VENDORS = [];
var PROPERTIES = [];
var PROPERTY_GROUPS = [];
var TURNS = [];
var UPCOMING_MOVEOUTS = []; // from tenant_directory — tenants on notice
var TURN_WORK_ORDERS = []; // from DB API — unit turn WOs with real-time status
var UNIT_TURNS_DB = [];    // SQL-backed unit turn tracker records
var UNIT_TURN_TRACKER_BY_KEY = {};
var _turnTrackerSyncInFlight = false;
var _lastTurnTrackerSyncHash = '';
var BILLS = []; // from DB API — AP bills for WO close-assist
var RECENT_TASKS = [];
var WEBHOOK_EVENTS = [];
var _webhookPollTimer = null;
var _autoSyncTimer = null;      // 30-min background selective refresh
var _vendorsLazyLoaded = false; // lazy-load flag — vendors fetched on tab click
var _inspLazyLoaded = false;   // lazy-load flag — inspections fetched on tab click
var _whLazyLoaded = false;     // lazy-load flag — webhook data loaded on tab click
var _groupFilterDirty = {};    // tabs that need re-render after group filter change
var _proxyVersion = 'v7';     // detected on ping; default is legacy-safe if version is omitted
// UUID-based property group lookup maps (built by resolvePropertyGroupNames)
var _nameToGroups = {};        // property name (lowercase) → [group names]
var _idToGroups = {};          // Reports API property_id (string) → [group names]
var _uuidToGroups = {};        // DB API property UUID → [group names]
var appInitialized = false;
var WO_FLAGS = {};
var WO_DETAIL_CACHE = {};
var WO_DETAIL_CACHE_KEYS = []; // LRU order tracker
var WO_DETAIL_CACHE_MAX = 50;  // max cached entries
var CURRENT_WO_MODAL = null;
var PAYROLL_WEEK_OFFSET = 0;
var currentPropertyGroup = '';
var currentTurnFilter = 'open';
var currentWOCloseAssistAge = 14;
var _billsLoading = false;
var _billsLoadedAt = 0;
var _vendorRenderLimit = 0;
var currentVendorInitial = '';
var _vendorRenderKey = '';
var _vendorsNeedRender = false;

/* =================================================================
   CONFIG — Consolidated thresholds (edit here, not scattered in code)
   ================================================================= */
var CONFIG = {
  // Turn pipeline
  TURN_TARGET_DAYS: 30,         // default target completion days
  TURN_STALLED_DAYS: 7,         // days before a turn is flagged stalled
  TURN_WARNING_DAYS: 14,        // elapsed days before amber warning
  DEPOSIT_BUSINESS_DAYS: 21,    // AZ deposit deadline: 21 business days
  DEPOSIT_HOLIDAYS: [           // company holidays (YYYY-MM-DD) — update annually
    '2026-01-01','2026-01-19','2026-02-16','2026-05-25','2026-07-03',
    '2026-09-07','2026-11-26','2026-11-27','2026-12-24','2026-12-25'
  ],
  // Inspections
  INSPECTION_OVERDUE_DAYS: 365,
  INSPECTION_DUE_SOON_DAYS: 270,
  // Vendor compliance
  VENDOR_EXPIRY_ALERT_DAYS: 60,
  VENDOR_GRID_INITIAL_LIMIT_DESKTOP: 140,
  VENDOR_GRID_INITIAL_LIMIT_MOBILE: 40,
  VENDOR_GRID_LOAD_MORE_STEP: 80,
  VENDOR_SELECT_INITIAL_LIMIT_DESKTOP: 250,
  VENDOR_SELECT_INITIAL_LIMIT_MOBILE: 60,
  VENDOR_SELECT_SEARCH_LIMIT: 260,
  VENDOR_SELECT_MIN_SEARCH_CHARS: 2,
  // Move-outs
  MOVEOUT_WINDOW_DAYS: 60,
  // UI
  TOAST_DURATION_MS: 3500,
  DEBOUNCE_MS: 300,
  // WO aging buckets (dashboard)
  WO_AGING_BUCKETS: [
    { label: '0–7d',  max: 7,   cls: 'fresh' },
    { label: '8–30d', max: 30,  cls: 'fresh' },
    { label: '31–60d', max: 60, cls: 'moderate' },
    { label: '60d+',  max: Infinity, cls: 'old' }
  ]
};

/* =================================================================
   UTILITY — debounce, timeline, helpers
   ================================================================= */
function debounce(fn, delay) {
  var timer;
  return function() {
    var ctx = this, args = arguments;
    clearTimeout(timer);
    timer = setTimeout(function() { fn.apply(ctx, args); }, delay);
  };
}

function proxyMajorVersion(versionTag) {
  var m = String(versionTag || '').match(/^v?(\d+)/i);
  return m ? parseInt(m[1], 10) : 7;
}

function supportsServerCacheOps() {
  return proxyMajorVersion(_proxyVersion) >= 8;
}

function isConstrainedDevice() {
  var narrow = window.matchMedia && window.matchMedia('(max-width: 1024px)').matches;
  var touch = window.matchMedia && window.matchMedia('(pointer: coarse)').matches;
  return !!(narrow || touch);
}

// 21-business-day deposit deadline countdown from move-out date
// Returns { deadline, businessDaysLeft, calendarDaysLeft, overdue (bool), pct }
function calculateDepositDeadline(moveOutDateStr) {
  if (!moveOutDateStr) return null;
  var start = new Date(moveOutDateStr);
  if (isNaN(start.getTime())) return null;
  var holidays = CONFIG.DEPOSIT_HOLIDAYS;
  var current = new Date(start);
  var added = 0;
  while (added < CONFIG.DEPOSIT_BUSINESS_DAYS) {
    current.setDate(current.getDate() + 1);
    var dow = current.getDay();
    if (dow === 0 || dow === 6) continue;
    var iso = current.toISOString().slice(0, 10);
    if (holidays.indexOf(iso) !== -1) continue;
    added++;
  }
  var today = new Date();
  today.setHours(0,0,0,0);
  var deadline = new Date(current);
  deadline.setHours(0,0,0,0);
  var calendarDaysLeft = Math.ceil((deadline - today) / 86400000);
  var businessDaysLeft = 0;
  if (calendarDaysLeft > 0) {
    var checkDate = new Date(today);
    for (var di = 0; di < calendarDaysLeft; di++) {
      checkDate.setDate(checkDate.getDate() + 1);
      var d = checkDate.getDay();
      if (d !== 0 && d !== 6) {
        var isoCheck = checkDate.toISOString().slice(0, 10);
        if (holidays.indexOf(isoCheck) === -1) businessDaysLeft++;
      }
    }
  }
  return {
    deadline: deadline,
    businessDaysLeft: Math.max(0, businessDaysLeft),
    calendarDaysLeft: calendarDaysLeft,
    overdue: calendarDaysLeft < 0,
    breached: calendarDaysLeft < 0, // compat alias
    pct: Math.max(0, Math.min(100, ((CONFIG.DEPOSIT_BUSINESS_DAYS - businessDaysLeft) / CONFIG.DEPOSIT_BUSINESS_DAYS) * 100))
  };
}

// Build vendor compliance lookup for cross-tab warnings
function buildVendorComplianceMap() {
  var map = {};
  var today = new Date();
  VENDORS.forEach(function(v) {
    if (!v.name) return;
    var key = v.name.toLowerCase();
    var insDate = v.insurance ? new Date(v.insurance) : null;
    if (insDate && insDate < today) {
      map[key] = 'expired';
    } else if (insDate && daysBetween(today, insDate) <= CONFIG.VENDOR_EXPIRY_ALERT_DAYS) {
      map[key] = 'expiring';
    }
  });
  return map;
}

function isClosedTurnWorkOrderStatus(status) {
  var normalized = String(status || '').trim().toLowerCase();
  return normalized === 'completed' || normalized === 'work completed' || normalized === 'canceled' || normalized === 'cancelled';
}

function isTurnWorkDoneStatus(status) {
  var normalized = String(status || '').trim().toLowerCase();
  return normalized === 'work done' || normalized === 'ready to bill' || isClosedTurnWorkOrderStatus(normalized);
}

// Aggregate turn completion: check ALL Unit Turn WOs for a unit
function isTurnFullyComplete(matchingWOs) {
  if (!matchingWOs || matchingWOs.length === 0) return false;
  return matchingWOs.every(function(wo) {
    return isClosedTurnWorkOrderStatus(wo.status);
  });
}

/* =================================================================
   LRU Detail Cache helpers — bounded at WO_DETAIL_CACHE_MAX entries
   ================================================================= */
function detailCacheGet(key) {
  if (!(key in WO_DETAIL_CACHE)) return undefined;
  // Move to end (most recent)
  var idx = WO_DETAIL_CACHE_KEYS.indexOf(key);
  if (idx > -1) WO_DETAIL_CACHE_KEYS.splice(idx, 1);
  WO_DETAIL_CACHE_KEYS.push(key);
  return WO_DETAIL_CACHE[key];
}
function detailCacheSet(key, value) {
  if (key in WO_DETAIL_CACHE) {
    var idx = WO_DETAIL_CACHE_KEYS.indexOf(key);
    if (idx > -1) WO_DETAIL_CACHE_KEYS.splice(idx, 1);
  }
  WO_DETAIL_CACHE_KEYS.push(key);
  WO_DETAIL_CACHE[key] = value;
  // Evict oldest if over limit
  while (WO_DETAIL_CACHE_KEYS.length > WO_DETAIL_CACHE_MAX) {
    var evict = WO_DETAIL_CACHE_KEYS.shift();
    delete WO_DETAIL_CACHE[evict];
  }
}
function detailCacheClear() {
  WO_DETAIL_CACHE = {};
  WO_DETAIL_CACHE_KEYS = [];
}

var APP_CONFIG_KEY = 'handymgr_config_v1'; // persisted via IndexedDB cache store

function _encodeConfigPayload(payload) {
  try {
    return btoa(unescape(encodeURIComponent(JSON.stringify(payload))));
  } catch (e) {
    return '';
  }
}

function _decodeConfigPayload(raw) {
  try {
    return JSON.parse(decodeURIComponent(escape(atob(raw))));
  } catch (e) {
    return null;
  }
}

function getVaultConfigFromInputs() {
  return {
    vhost: sanitizeVhost(($('#vaultVhost') && $('#vaultVhost').value) || ''),
    proxy: sanitizeProxy(($('#vaultProxy') && $('#vaultProxy').value) || ''),
    updatedAt: new Date().toISOString()
  };
}

function applyVaultConfigToInputs(cfg) {
  if (!cfg) return;
  if ($('#vaultVhost') && cfg.vhost) {
    $('#vaultVhost').value = sanitizeVhost(cfg.vhost);
    $('#vhostPreview').textContent = $('#vaultVhost').value || 'yourco';
  }
  if ($('#vaultProxy') && cfg.proxy) {
    $('#vaultProxy').value = sanitizeProxy(cfg.proxy);
  }
}

async function saveVaultConfig(cfg) {
  if (!cfg || !cfg.vhost) return;
  await cacheSet(APP_CONFIG_KEY, cfg);
  var ts = $('#vaultConfigSavedAt');
  if (ts) ts.textContent = 'Saved ' + timeAgo(cfg.updatedAt || new Date().toISOString());
}

async function loadVaultConfig() {
  var saved = await cacheGet(APP_CONFIG_KEY);
  if (saved && saved.data) return saved.data;
  return null;
}

function getConfigFromUrl() {
  try {
    var qs = new URLSearchParams(window.location.search);
    var encoded = qs.get('cfg');
    if (!encoded) return null;
    return _decodeConfigPayload(encoded);
  } catch (e) {
    return null;
  }
}

function buildConfigShareUrl() {
  var cfg = getVaultConfigFromInputs();
  if (!cfg.vhost) return '';
  var encoded = _encodeConfigPayload(cfg);
  if (!encoded) return '';
  var u = new URL(window.location.href);
  u.searchParams.set('cfg', encoded);
  return u.toString();
}

async function initVaultConfigUI() {
  var rememberEl = $('#vaultRememberConfig');
  var copyBtn = $('#vaultCopyConfigLink');
  var vhostEl = $('#vaultVhost');
  var proxyEl = $('#vaultProxy');

  if (!vhostEl || !proxyEl) return;

  var cfgFromUrl = getConfigFromUrl();
  if (cfgFromUrl) {
    applyVaultConfigToInputs(cfgFromUrl);
    if (!rememberEl || rememberEl.checked) {
      await saveVaultConfig(cfgFromUrl);
    }
  } else {
    var savedCfg = await loadVaultConfig();
    if (savedCfg) applyVaultConfigToInputs(savedCfg);
  }

  function handleConfigInput() {
    if (rememberEl && rememberEl.checked) {
      saveVaultConfig(getVaultConfigFromInputs());
    }
  }

  vhostEl.addEventListener('change', handleConfigInput);
  proxyEl.addEventListener('change', handleConfigInput);
  if (rememberEl) {
    rememberEl.addEventListener('change', function() {
      if (rememberEl.checked) {
        saveVaultConfig(getVaultConfigFromInputs());
      }
    });
  }
  if (copyBtn) {
    copyBtn.addEventListener('click', async function() {
      var shareUrl = buildConfigShareUrl();
      if (!shareUrl) {
        showToast('Enter subdomain first to generate a sync link');
        return;
      }
      try {
        await navigator.clipboard.writeText(shareUrl);
        showToast('Config link copied — open on another device to sync settings');
      } catch (e) {
        showToast('Could not copy link — clipboard access denied');
      }
    });
  }
}

/* =================================================================
   VAULT UI
   ================================================================= */
// Sanitize vhost input — extract just the subdomain portion
function sanitizeVhost(raw) {
  var val = raw.trim().toLowerCase();
  // Strip protocol
  val = val.replace(/^https?:\/\//, '');
  // Strip trailing slashes/paths
  val = val.replace(/\/.*$/, '');
  // Strip .appfolio.com suffix (user may paste full domain)
  val = val.replace(/\.appfolio\.com$/, '');
  // Only allow valid subdomain chars
  val = val.replace(/[^a-z0-9\-]/g, '');
  return val;
}

$('#vaultVhost').addEventListener('input', function() {
  var val = sanitizeVhost(this.value);
  this.value = val;
  $('#vhostPreview').textContent = val || 'yourco';
});

$('#vaultToggleVis').addEventListener('click', function() {
  var inp = $('#vaultPassphrase');
  var isPass = inp.type === 'password';
  inp.type = isPass ? 'text' : 'password';
  this.querySelector('i').className = isPass ? 'fas fa-eye-slash' : 'fas fa-eye';
});

$('#vaultPassphrase').addEventListener('keydown', function(e) {
  if (e.key === 'Enter') { $('#vaultUnlockBtn').click(); }
});
$('#vaultVhost').addEventListener('keydown', function(e) {
  if (e.key === 'Enter') { $('#vaultUnlockBtn').click(); }
});

// Advanced panel toggle
$('#advancedToggle').addEventListener('click', function() {
  this.classList.toggle('open');
  $('#advancedPanel').classList.toggle('show');
});

// Proxy preset buttons
$$('.vault-proxy-preset').forEach(function(btn) {
  btn.addEventListener('click', function() {
    $$('.vault-proxy-preset').forEach(function(b) { b.classList.remove('active'); });
    btn.classList.add('active');
    $('#vaultProxy').value = btn.getAttribute('data-proxy');
  });
});

if ($('#btnSendOtp')) {
  $('#btnSendOtp').addEventListener('click', async function() {
    var emailRaw = $('#vaultOtpEmail') ? $('#vaultOtpEmail').value : '';
    var email = normalizeOtpEmail(emailRaw);
    if (!email) {
      $('#vaultError').textContent = 'Enter a valid @flraz.com email to receive OTP.';
      $('#vaultError').classList.add('show');
      return;
    }
    API_PROXY = sanitizeProxy($('#vaultProxy').value || '');
    if (!API_PROXY) {
      $('#vaultError').textContent = 'Proxy URL is required before requesting OTP.';
      $('#vaultError').classList.add('show');
      return;
    }
    var btn = this;
    btn.disabled = true;
    btn.textContent = 'Sending...';
    $('#vaultError').classList.remove('show');
    try {
      await requestDeviceOtp(email, 'dispatcher');
      showToast('OTP sent to ' + email, { kind: 'success' });
    } catch (err) {
      $('#vaultError').textContent = err.message || String(err);
      $('#vaultError').classList.add('show');
    } finally {
      btn.disabled = false;
      btn.textContent = 'Send OTP';
    }
  });
}

if ($('#btnVerifyOtp')) {
  $('#btnVerifyOtp').addEventListener('click', async function() {
    var email = normalizeOtpEmail($('#vaultOtpEmail') ? $('#vaultOtpEmail').value : '');
    var code = String(($('#vaultOtpCode') && $('#vaultOtpCode').value) || '').trim();
    if (!email) {
      $('#vaultError').textContent = 'Enter a valid @flraz.com email first.';
      $('#vaultError').classList.add('show');
      return;
    }
    if (!/^\d{6}$/.test(code)) {
      $('#vaultError').textContent = 'Enter the 6-digit OTP code.';
      $('#vaultError').classList.add('show');
      return;
    }
    API_PROXY = sanitizeProxy($('#vaultProxy').value || '');
    if (!API_PROXY) {
      $('#vaultError').textContent = 'Proxy URL is required before verifying OTP.';
      $('#vaultError').classList.add('show');
      return;
    }
    var btn = this;
    btn.disabled = true;
    btn.textContent = 'Verifying...';
    $('#vaultError').classList.remove('show');
    try {
      var token = await verifyDeviceOtp(email, code, 'dispatcher');
      try { localStorage.setItem('hm_device_token', token); } catch (e) { /* */ }
      try { localStorage.setItem('hm_proxy_token', token); } catch (e2) { /* */ }
      if ($('#vaultPassphrase')) $('#vaultPassphrase').value = '';
      showToast('Device verified — you can now connect without setup PIN.', { kind: 'success' });
    } catch (err) {
      $('#vaultError').textContent = err.message || String(err);
      $('#vaultError').classList.add('show');
    } finally {
      btn.disabled = false;
      btn.textContent = 'Verify Device';
    }
  });
}

// Sanitize proxy URL — ensure https:// prefix, trim whitespace
function sanitizeProxy(raw) {
  var val = (raw || '').trim();
  if (!val) return '';
  // Auto-add https:// if user forgot
  if (val && !val.match(/^https?:\/\//i)) {
    val = 'https://' + val;
  }
  // Remove trailing slash for consistency
  val = val.replace(/\/+$/, '');
  return val;
}

initVaultConfigUI();

$('#vaultUnlockBtn').addEventListener('click', async function() {
  var pass = $('#vaultPassphrase').value;
  // Re-sanitize at unlock time in case user pasted a full URL
  var rawVhost = $('#vaultVhost').value;
  var vhost = sanitizeVhost(rawVhost);
  $('#vaultVhost').value = vhost;
  $('#vhostPreview').textContent = vhost || 'yourco';
  // Sanitize proxy URL
  var proxyUrl = sanitizeProxy($('#vaultProxy').value);
  $('#vaultProxy').value = proxyUrl;
  var existingDeviceToken = '';
  try { existingDeviceToken = localStorage.getItem('hm_device_token') || ''; } catch (e) { /* */ }

  if (!vhost) {
    $('#vaultError').textContent = 'AppFolio subdomain is required. Enter just the subdomain (e.g. "flraz"), not the full URL.';
    $('#vaultError').classList.add('show');
    return;
  }
  if (!pass && !existingDeviceToken) {
    $('#vaultError').textContent = 'Use Setup PIN or verify this device with @flraz.com OTP before connecting.';
    $('#vaultError').classList.add('show');
    return;
  }
  if (proxyUrl) {
    console.log('Proxy URL: ' + proxyUrl);
  }

  var btn = this;
  btn.disabled = true;
  btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Decrypting\u2026';
  $('#vaultError').classList.remove('show');

  try {
    API_VHOST = vhost;
    API_PROXY = proxyUrl;

    if (API_PROXY && existingDeviceToken && !pass) {
      API_CREDS = { p: existingDeviceToken };
      _accessRole = getStoredAccessRole();
      try { localStorage.setItem('hm_proxy_token', existingDeviceToken); } catch (e) { /* */ }
    } else if (API_PROXY) {
      var proxyToken = '';
      try {
        proxyToken = await setupTrustedDevice(pass, 'dispatcher');
        try { localStorage.setItem('hm_device_token', proxyToken); } catch (e) { /* */ }
      } catch (_setupErr) {
        // Fallback: treat pass field as static bearer token when setup pin flow fails.
        proxyToken = String(pass || '').trim();
      }
      API_CREDS = { p: proxyToken };
      _accessRole = pass === ROLE_TOKEN_VENDOR ? 'vendors' : (pass === ROLE_TOKEN_MANAGER ? 'manager' : 'full');
      persistAccessRole(_accessRole);
      try { localStorage.setItem('hm_proxy_token', proxyToken); } catch (e) { /* */ }

      // Validate auth with lightweight ping. cache_stats can fail on schema drift.
      await proxyAction('ping');
    } else {
      API_CREDS = await decryptVault(pass);
      persistAccessRole(_accessRole);
    }

    if (!$('#vaultRememberConfig') || $('#vaultRememberConfig').checked) {
      await saveVaultConfig(getVaultConfigFromInputs());
    }
    $('#vaultPassphrase').value = '';
    $('#vaultScreen').style.display = 'none';
    $('#appShell').classList.add('unlocked');
    applyAccessRole();
    await initApp();
    applyAccessRole();
    startAutoSync();
    var proxyInfo = API_PROXY ? ' via proxy' : ' (direct)';
    if (_accessRole === 'vendors') {
      showToast('Vendor access \u2014 connecting to ' + vhost + '.appfolio.com' + proxyInfo);
    } else {
      showToast('Connected \u2014 ' + vhost + '.appfolio.com' + proxyInfo);
      maybeShowWhatsNew();
    }
  } catch (err) {
    var errMsg = (err && (err.message || String(err))) || '';
    var schemaErr = /no such table|no such column|SQLITE_UNKNOWN|SQL_INPUT_ERROR/i.test(errMsg);
    wipeCredentials();
    if (API_PROXY) {
      $('#vaultError').textContent = schemaErr
        ? 'Proxy connected, but database schema is out of date (missing tables/columns). Deploy latest proxy migrations and retry.'
        : 'Proxy auth failed \u2014 enter Setup PIN or static bearer token.';
    } else {
      $('#vaultError').textContent = 'Decryption failed \u2014 incorrect passphrase or corrupted vault.';
    }
    $('#vaultError').classList.add('show');
    $('#vaultPassphrase').value = '';
    $('#vaultPassphrase').focus();
  } finally {
    btn.disabled = false;
    btn.innerHTML = '<i class="fas fa-sign-in-alt"></i> Connect';
  }
});

$('#lockBtn').addEventListener('click', function() { openModal('lockModal'); });
$('#lockModalClose').addEventListener('click', function() { closeModal('lockModal'); });
$('#lockCancelBtn').addEventListener('click', function() { closeModal('lockModal'); });
$('#lockConfirmBtn').addEventListener('click', function() {
  closeModal('lockModal');
  lockVault();
  showToast('Vault locked \u2014 credentials wiped from memory');
});

$('#woModalClose').addEventListener('click', function() { closeModal('woModal'); });
$('#woModalCloseBtn').addEventListener('click', function() { closeModal('woModal'); });
$('#newWOModalClose').addEventListener('click', function() { closeModal('newWOModal'); });
$('#newWOCancelBtn').addEventListener('click', function() { closeModal('newWOModal'); });

/* =================================================================
   COMMUNICATION TEMPLATES (local — no API endpoint for these)
   ================================================================= */
var TEMPLATES = [
  { title: 'Tenant \u2014 Work Completed', trigger: 'Status \u2192 Work Completed', icon: 'fa-check-circle', body: 'Maintenance Alert: The technician has completed the <span class="var">{{ description }}</span> at <span class="var">{{ unit_name }}</span>. If you have further issues, please reply to this thread.' },
  { title: 'Owner \u2014 Urgent Repair', trigger: 'Priority = Urgent', icon: 'fa-exclamation-circle', body: 'Management Notification: An urgent repair was required at <span class="var">{{ property_name }}</span> for <span class="var">{{ description }}</span>. The total cost was <span class="var">{{ total_cost }}</span>. Photos are available in your portal.' },
  { title: 'Vendor \u2014 Dispatch', trigger: 'Status \u2192 Assigned', icon: 'fa-paper-plane', body: 'New Work Order Dispatch: You have been assigned <span class="var">{{ work_order_id }}</span> at <span class="var">{{ address }}</span>, <span class="var">{{ unit_name }}</span>. Scheduled start: <span class="var">{{ scheduled_start }}</span>.' },
  { title: 'Tenant \u2014 Scheduled Visit', trigger: 'Status \u2192 Scheduled', icon: 'fa-calendar-check', body: 'Hi <span class="var">{{ first_name }}</span>, a maintenance visit has been scheduled for <span class="var">{{ scheduled_date }}</span>. The technician will address: <span class="var">{{ description }}</span>.' },
  { title: 'Owner \u2014 Monthly Summary', trigger: 'Report: 1st of month', icon: 'fa-chart-bar', body: 'Monthly Maintenance Summary for <span class="var">{{ property_name }}</span>: <span class="var">{{ wo_count }}</span> work orders completed, total spend <span class="var">{{ total_spend }}</span>. <span class="var">{{ open_count }}</span> orders remain open.' },
  { title: 'Tenant \u2014 Estimate Approval', trigger: 'Estimate created', icon: 'fa-file-invoice-dollar', body: 'A repair estimate of <span class="var">{{ estimate_amount }}</span> has been prepared for <span class="var">{{ description }}</span> at <span class="var">{{ unit_name }}</span>. Reply "APPROVE" to proceed.' }
];

/* =================================================================
   SMART FILTERS — 180-day window, open-only WOs, small chunks
   ================================================================= */
var DATA_WINDOW_DAYS = 180;
var PROPERTY_GROUPS_LAST_UPDATED_FROM = '2024-01-01T00:00:00Z';
var PROPERTY_GROUPS_PAGE_SIZE = 100;

function dateNDaysAgo(n) {
  var d = new Date();
  d.setDate(d.getDate() - n);
  return d.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

// Open WO status codes for Reports API (excludes 4=Completed, 5=Canceled, 7=CompletedNoNeedToBill)
// Defined in fetchWorkOrders() as OPEN_WO_STATUS_CODES

function buildPropertyGroupsQueryParams() {
  return {
    'filters[LastUpdatedAtFrom]': PROPERTY_GROUPS_LAST_UPDATED_FROM,
    'page[size]': PROPERTY_GROUPS_PAGE_SIZE
  };
}

function buildPropertyGroupsProxyCurl(proxyUrl) {
  if (!proxyUrl) return '';
  var qp = new URLSearchParams();
  qp.append('action', 'property_groups');
  qp.append('filters[LastUpdatedAtFrom]', PROPERTY_GROUPS_LAST_UPDATED_FROM);
  qp.append('page[size]', String(PROPERTY_GROUPS_PAGE_SIZE));
  return 'curl -X GET "' + proxyUrl + '?' + qp.toString() + '" --compressed';
}

/* =================================================================
   PROGRESS DOCK — non-blocking corner indicator
   ================================================================= */
var _progSteps = [];

function showProgress(title, steps) {
  _progSteps = steps.map(function(s) { return { label: s, state: 'pending' }; });
  var dock = $('#progressDock');
  $('#progTitle').textContent = title;
  $('#progStatus').textContent = 'Starting\u2026';
  $('#progBar').style.width = '0%';
  var stepsHtml = '';
  _progSteps.forEach(function() { stepsHtml += '<div class="progress-step"></div>'; });
  $('#progSteps').innerHTML = stepsHtml;
  dock.classList.remove('hidden');
}

function updateProgress(stepIndex, state, statusText) {
  if (stepIndex >= 0 && stepIndex < _progSteps.length) {
    _progSteps[stepIndex].state = state;
  }
  var doneCount = _progSteps.filter(function(s) { return s.state === 'done'; }).length;
  var pct = _progSteps.length > 0 ? Math.round((doneCount / _progSteps.length) * 100) : 0;
  $('#progBar').style.width = pct + '%';
  if (statusText) { $('#progStatus').textContent = statusText; }
  // Update step dots
  var dots = $$('#progSteps .progress-step');
  _progSteps.forEach(function(s, i) {
    if (dots[i]) { dots[i].className = 'progress-step ' + s.state; }
  });
}

function hideProgress() {
  setTimeout(function() { $('#progressDock').classList.add('hidden'); }, 2000);
}

/* =================================================================
   DATA FETCHING — Smart filtered API calls (180-day window)
   ================================================================= */

// Work Orders: Proxy ?action=work_orders — server-side pagination, one request
// Open statuses: 0=New, 1=EstReq, 2=Estimated, 9=Assigned, 3=Scheduled,
//   6=Waiting, 8=WorkDone, 12=ReadyToBill
// Excludes: 4=Completed, 5=Canceled, 7=CompletedNoNeedToBill
async function fetchWorkOrders() {
  try {
    setApiStatus('loading', 'Loading work orders (server-side)\u2026');
    var data = await proxyAction('work_orders', { days: DATA_WINDOW_DAYS });
    var results = data.results || [];
    WORK_ORDERS = results.map(function(r) {
      return {
        id: r.work_order_number || r.WorkOrderNumber || r.service_request_number || '',
        uuid: r.work_order_id || r.Id || '',
        propertyId: r.property_id || r.PropertyId || '',
        propertyName: r.property_name || r.property || r.PropertyName || '',
        propertyAddress: ((r.property_street || '') + ' ' + (r.property_city || '') + ' ' + (r.property_state || '') + ' ' + (r.property_zip || '')).trim(),
        unitId: r.unit_id || r.UnitId || '',
        unit: r.unit_name || r.UnitName || r.unit_id || '',
        priority: r.priority || r.Priority || 'Normal',
        status: r.status || r.Status || 'New',
        description: r.job_description || r.JobDescription || r.service_request_description || r.Description || '',
        vendorName: r.vendor || r.VendorName || '',
        vendorId: r.vendor_id || r.VendorId || '',
        vendorTrade: r.vendor_trade || r.VendorTrade || '',
        created: r.created_at || r.CreatedAt || '',
        updated: r.completed_on || r.CompletedOn || r.created_at || r.CreatedAt || '',
        completedOn: r.completed_on || r.CompletedOn || '',
        workCompletedOn: r.work_completed_on || r.WorkCompletedOn || '',
        scheduledStart: r.scheduled_start || r.ScheduledStart || '',
        scheduledEnd: r.scheduled_end || r.ScheduledEnd || '',
        type: r.work_order_type || r.Type || '',
        amount: r.amount || r.Amount || '',
        tenant: r.primary_tenant || r.PrimaryTenant || '',
        tenantEmail: r.primary_tenant_email || r.PrimaryTenantEmail || '',
        tenantPhone: r.primary_tenant_phone_number || r.PrimaryTenantPhoneNumber || '',
        createdBy: r.created_by || r.CreatedBy || '',
        assignedUser: r.assigned_user || r.AssignedUser || r.AssignedTo || '',
        statusNotes: r.status_notes || r.StatusNotes || '',
        maintenanceLimit: r.maintenance_limit || r.MaintenanceLimit || '',
        link: r.Link || r.link || ''
      };
    });
    var cacheNote = (data.from_cache && data.cached_at) ? ' (cached)' : '';
    setApiStatus('loading', 'Work orders: ' + WORK_ORDERS.length + ' loaded' + cacheNote);
    return true;
  } catch (err) {
    WORK_ORDERS = [];
    return false;
  }
}

// Bills: Proxy ?action=bills — DB API v0
// Used by WO close-assist to infer likely completion.
async function fetchBills(days) {
  try {
    var lookback = parseInt(days || 365, 10) || 365;
    var data = await proxyAction('bills', { days: String(lookback), max: '1500' });
    var results = data.results || data.data || [];
    BILLS = results.map(function(b) {
      return {
        id: b.Id || b.id || b.BillId || '',
        vendorId: b.VendorId || b.vendor_id || b.PayeeId || b.payee_id || b.PayeeUuid || b.payee_uuid || '',
        vendorName: b.VendorName || b.vendor_name || b.PayeeName || b.payee_name || b.Name || b.name || '',
        propertyId: b.PropertyId || b.property_id || b.PropertyUuid || b.property_uuid || '',
        propertyName: b.PropertyName || b.property_name || b.Property || b.property || '',
        amount: b.Amount || b.amount || b.Total || b.total || '',
        date: b.BillDate || b.bill_date || b.InvoiceDate || b.invoice_date || b.PaidOn || b.paid_on || b.CreatedAt || b.created_at || b.LastUpdatedAt || b.last_updated_at || '',
        raw: b
      };
    });
    _billsLoadedAt = Date.now();
    return true;
  } catch (err) {
    console.log('fetchBills error: ' + (err.message || err));
    BILLS = [];
    return false;
  }
}

// Vendors: Proxy ?action=vendors — server-side pagination, one request
async function fetchVendors() {
  try {
    setApiStatus('loading', 'Loading vendors (server-side)\u2026');
    var data = await proxyAction('vendors');
    var results = data.results || [];
    VENDORS = results.map(function(v) {
      var displayName = v.company_name || ((v.first_name || '') + ' ' + (v.last_name || '')).trim() || v.name || '';
      return {
        id: v.vendor_id || '',
        name: displayName,
        companyName: v.company_name || '',
        firstName: v.first_name || '',
        lastName: v.last_name || '',
        isCompany: !!v.company_name,
        compliant: false,
        compliantStatus: 'Unknown',
        insurance: v.liability_ins_expires || '',
        autoInsurance: v.auto_ins_expires || '',
        workersComp: v.workers_comp_expires || '',
        phone: v.phone_numbers || '',
        email: v.email || '',
        address: ((v.street || '') + ' ' + (v.city || '') + ' ' + (v.state || '') + ' ' + (v.zip || '')).trim(),
        trades: v.vendor_trades || '',
        vendorType: v.vendor_type || '',
        doNotUse: v.do_not_use_for_work_order || false,
        tags: v.tags || '',
        link: ''
      };
    });
    return true;
  } catch (err) {
    VENDORS = [];
    return false;
  }
}

// Properties: Proxy ?action=properties — server-side pagination, one request
async function fetchProperties() {
  try {
    setApiStatus('loading', 'Loading properties (server-side)…');
    var data = await proxyAction('properties');
    var results = data.results || data.data || [];
    PROPERTIES = results.map(function(p) {
      return {
        id: p.property_id || p.id || '',
        name: p.property_name || p.property || p.name || '',
        address: ((p.property_street || p.street || '') + (p.property_street2 ? ' ' + p.property_street2 : '')).trim(),
        city: p.property_city || p.city || '',
        state: p.property_state || p.state || '',
        zip: p.property_zip || p.zip || '',
        propertyType: p.property_type || p.type || '',
        portfolioId: p.portfolio_id || p.portfolioId || '',
        portfolio: p.portfolio || p.portfolio_name || p.portfolioName || p.group_name || p.property_group || '',
        portfolioName: p.portfolio_name || p.portfolioName || '',
        propertyGroup: p.property_group || p.group_name || p.group || '',
        group: p.group || '',
        groupName: p.group_name || '',
        maintenanceLimit: p.maintenance_limit || p.maintenanceLimit || '',
        maintenanceNotes: p.maintenance_notes || p.maintenanceNotes || '',
        siteManager: p.site_manager || p.siteManager || '',
        units: p.units || '',
        sqft: p.sqft || '',
        marketRent: p.market_rent || p.marketRent || '',
        owners: p.owners || '',
        link: ''
      };
    });
    return true;
  } catch (err) {
    PROPERTIES = [];
    return false;
  }
}

// Turns: Proxy ?action=turns — merged In Progress + Completed for richer detail context
async function fetchTurns() {
  try {
    setApiStatus('loading', 'Loading turns (In Progress + Completed)…');
    var activeData = await proxyAction('turns', { days: 120, status: 'In Progress' });
    var activeRows = activeData.results || activeData.data || [];
    var completedRows = [];
    try {
      var completedData = await proxyAction('turns', { days: 365, status: 'Completed' });
      completedRows = completedData.results || completedData.data || [];
    } catch (e) {
      console.log('Completed turns fetch skipped: ' + (e.message || e));
    }

    var mergedById = {};
    activeRows.concat(completedRows).forEach(function(r) {
      var id = r.unit_turn_id || r.unit_turn_uuid || r.unitTurnId || r.id || (r.unit || '') + '|' + (r.property || '') + '|' + (r.move_out_date || r.move_out || '');
      mergedById[String(id)] = r;
    });
    var results = Object.keys(mergedById).map(function(k) { return mergedById[k]; });

    TURNS = results.map(function(t) {
      var moveOut = t.move_out_date || t.move_out || t.moveOutDate || t.moveOut || '';
      var turnEnd = t.turn_end_date || t.turn_end || t.turnEndDate || t.completed_at || '';
      var daysToComplete = parseInt(t.total_days_to_complete || t.totalDays || 0, 10) || 0;
      if (!daysToComplete && moveOut && turnEnd) {
        var d1 = new Date(moveOut), d2 = new Date(turnEnd);
        if (!isNaN(d1) && !isNaN(d2)) { daysToComplete = Math.round((d2 - d1) / 86400000); }
      }
      return {
        unitTurnId: t.unit_turn_id || t.unit_turn_uuid || t.unitTurnId || t.id || '',
        unit: t.unit || t.unit_name || t.unit_number || t.unitName || '',
        property: t.property || t.property_name || t.propertyName || '',
        propertyId: t.property_id || t.property_uuid || t.propertyId || 0,
        unitId: t.unit_id || t.unit_uuid || t.unitId || 0,
        notes: t.notes || '',
        referenceUser: t.reference_user || t.referenceUser || t.created_by || '',
        moveOut: moveOut,
        turnEnd: turnEnd,
        expectedMoveIn: t.expected_move_in_date || t.expected_move_in || t.expectedMoveIn || '',
        targetDays: parseInt(t.target_days_to_complete || t.targetDays || 0, 10) || 0,
        totalDays: daysToComplete,
        laborCost: t.labor_from_work_orders || t.labor_cost || '$0.00',
        purchaseOrders: t.purchase_orders_from_work_orders || t.purchase_orders || '$0.00',
        billables: t.billables_from_work_orders || t.billables || '$0.00',
        inventory: t.inventory_from_work_orders || t.inventory || '$0.00',
        totalBilled: t.total_billed || t.totalBilled || '$0.00',
        status: t.unit_turn_status || t.status || '',
        siteManager: t.site_manager || t.property_site_manager || '',
        maintenanceLimit: t.maintenance_limit || t.property_maintenance_limit || '',
        propertyNotes: t.property_notes || t.maintenance_notes || '',
        isRegisteredUnitTurn: !!(t.registered_unit_turn || t.is_registered_unit_turn || t.registered_turn || t.registered_unit_turn_id),
        registeredUnitTurnId: t.registered_unit_turn_id || t.registered_turn_id || '',
        registrationSource: t.registered_unit_turn_source || t.registration_source || ''
      };
    });
    return true;
  } catch (err) {
    TURNS = [];
    return false;
  }
}

// Inspections: Proxy ?action=inspections — server-side pagination, one request
var INSPECTIONS = [];
var INSPECTION_LOOKBACK_DAYS = 180;

function getCurrentYearStartDate(nowRef) {
  var now = nowRef || new Date();
  return new Date(now.getFullYear(), 0, 1);
}

function isUuidString(value) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(String(value || '').trim());
}

function isActiveInspectionProperty(row) {
  if (!PROPERTIES || PROPERTIES.length === 0) return true;
  var propId = String(row.propertyId || '').trim();
  var propName = String(row.propertyName || '').trim().toLowerCase();
  return PROPERTIES.some(function(p) {
    return (propId && String(p.id || '').trim() === propId) ||
      (propName && String(p.name || '').trim().toLowerCase() === propName);
  });
}

var INSPECTION_AF_EPOCH = new Date('2021-01-01T00:00:00Z');

function toValidDateOrNull(value) {
  if (!value) return null;
  var d = new Date(value);
  if (isNaN(d.getTime())) return null;
  return d;
}

function getInspectionCompliance(row, nowRef) {
  var now = nowRef || new Date();
  var lastDate = toValidDateOrNull(row && row.lastInspection);
  var moveInDate = toValidDateOrNull(row && row.moveIn);
  var hasPastMoveIn = !!(moveInDate && moveInDate <= now);
  var hasMoveInInspection = !!(hasPastMoveIn && lastDate && lastDate >= moveInDate);
  var missingMoveInInspection = !!(hasPastMoveIn && !hasMoveInInspection);

  var anchorDate = hasPastMoveIn
    ? (hasMoveInInspection ? lastDate : moveInDate)
    : lastDate;

  var daysSince = anchorDate ? daysBetween(anchorDate, now) : 999;
  var overdue = missingMoveInInspection || !anchorDate || daysSince > CONFIG.INSPECTION_OVERDUE_DAYS;
  var dueSoon = !overdue && daysSince > CONFIG.INSPECTION_DUE_SOON_DAYS;

  return {
    anchorDate: anchorDate,
    lastDate: lastDate,
    moveInDate: moveInDate,
    hasPastMoveIn: hasPastMoveIn,
    hasMoveInInspection: hasMoveInInspection,
    missingMoveInInspection: missingMoveInInspection,
    daysSince: daysSince,
    overdue: overdue,
    dueSoon: dueSoon,
    current: !overdue && !dueSoon
  };
}

function isInspectionWithinWindow(row) {
  var now = new Date();
  var state = getInspectionCompliance(row, now);
  var anchor = state.anchorDate;
  if (!anchor) return true;
  if (anchor < INSPECTION_AF_EPOCH) return false;
  var yearStart = getCurrentYearStartDate(now);
  if (anchor < yearStart) return false;
  var daysOld = daysBetween(anchor, now);
  return daysOld <= INSPECTION_LOOKBACK_DAYS;
}

async function fetchInspections() {
  try {
    var now = new Date();
    var yearStart = getCurrentYearStartDate(now);
    INSPECTION_LOOKBACK_DAYS = Math.max(1, daysBetween(yearStart, now) + 1);
    setApiStatus('loading', 'Loading inspections (active properties, ' + INSPECTION_LOOKBACK_DAYS + 'd window)\u2026');
    var data = await proxyAction('inspections', { days: String(INSPECTION_LOOKBACK_DAYS), active_only: '1' });
    var results = data.results || [];
    INSPECTIONS = results.map(function(r) {
      return {
        propertyName: r.property_name || r.property || '',
        propertyId: r.property_id || 0,
        unit: r.unit_name || '',
        unitId: r.unit_id || 0,
        lastInspection: r.last_inspection_date || '',
        tenant: r.tenant_name || '',
        tenantPhone: r.tenant_primary_phone_number || '',
        moveIn: r.move_in_date || '',
        moveOut: r.move_out_date || '',
        rentable: r.rentable || '',
        tags: r.unit_tags || ''
      };
    }).filter(function(row) {
      return isActiveInspectionProperty(row) && isInspectionWithinWindow(row);
    });
    return true;
  } catch (err) {
    INSPECTIONS = [];
    return false;
  }
}

// Property Groups: Proxy ?action=property_groups — DB API v0
// DB API returns { data: [ { Id, Name, PropertyIds, Type, LastUpdatedAt } ] }
// UUID→Name resolution via separate ?action=property_map (non-blocking)
async function fetchPropertyGroups() {
  var previousGroups = Array.isArray(PROPERTY_GROUPS) ? PROPERTY_GROUPS.slice() : [];
  try {
    setApiStatus('loading', 'Loading property groups\u2026');
    var pgParams = buildPropertyGroupsQueryParams();
    var data;
    try {
      data = await proxyAction('property_groups', pgParams);
    } catch (actionErr) {
      // Fallback for proxies that do not forward DB API filters for this action.
      var pgPath = '/api/v0/property_groups?filters[LastUpdatedAtFrom]=' +
        encodeURIComponent(PROPERTY_GROUPS_LAST_UPDATED_FROM) +
        '&page[size]=' + encodeURIComponent(PROPERTY_GROUPS_PAGE_SIZE);
      console.log('[PG] property_groups action failed; using passthrough fallback: ' + (actionErr.message || actionErr));
      data = await apiFetch(pgPath);
    }
    console.log('[PG] property_groups response keys: ' + Object.keys(data).join(', '));

    var results = data.results || data.data || [];
    if (results.length > 0) {
      var rawPids = results[0].PropertyIds || results[0].Properties || results[0].properties || results[0].property_ids || [];
      console.log('[PG] Sample group=' + (results[0].Name || results[0].name || 'n/a') +
        ', PropertyIds type=' + (Array.isArray(rawPids) ? 'array(' + rawPids.length + ')' : typeof rawPids));
    }

    PROPERTY_GROUPS = results.map(function(g) {
      // Normalize PropertyIds: may be plain UUID strings OR objects like {Id:"uuid"}
      var rawProps = g.PropertyIds || g.Properties || g.properties || g.property_ids || [];
      var normalizedProps = (Array.isArray(rawProps) ? rawProps : []).map(function(pid) {
        if (typeof pid === 'string') return pid.trim();
        if (pid && typeof pid === 'object') return String(pid.Id || pid.id || pid.PropertyId || pid.property_id || '').trim();
        return String(pid || '').trim();
      }).filter(Boolean);

      return {
        id: g.Id || g.id || '',
        name: (g.Name || g.name || '').trim(),
        properties: normalizedProps,
        propertyNames: [],
        resolvedNames: []
      };
    });

    console.log('[PG] Parsed groups: ' + PROPERTY_GROUPS.length + ', with PropertyIds: ' +
      PROPERTY_GROUPS.filter(function(g) { return g.properties.length > 0; }).length);

    // Fetch UUID→Name map — await so resolution completes before filters are used
    await resolvePropertyGroupNames();

    return true;
  } catch (err) {
    console.log('[PG] fetchPropertyGroups FATAL error: ' + (err.message || err));
    // Keep last-known-good groups on transient failures.
    if (previousGroups.length > 0) {
      PROPERTY_GROUPS = previousGroups;
      console.log('[PG] preserving previous property groups: ' + previousGroups.length);
      return true;
    }

    // Final fallback: synthesize group names from loaded property metadata.
    var synthetic = {};
    (PROPERTIES || []).forEach(function(p) {
      var candidates = [
        p.portfolio,
        p.portfolioName,
        p.propertyGroup,
        p.group,
        p.groupName,
      ];
      candidates.forEach(function(raw) {
        var name = String(raw || '').trim();
        if (name) synthetic[name] = true;
      });
    });

    var names = Object.keys(synthetic).sort();
    if (names.length > 0) {
      PROPERTY_GROUPS = names.map(function(name, idx) {
        return {
          id: 'synthetic-' + idx,
          name: name,
          properties: [],
          propertyNames: [],
          resolvedNames: [],
        };
      });
      console.log('[PG] synthesized property groups from properties: ' + PROPERTY_GROUPS.length);
      return true;
    }

    PROPERTY_GROUPS = [];
    return false;
  }
}

// Resolve UUID→Name mapping for property groups (called after groups load)
// This is a separate call so property_groups never times out.
// If the UUID map fails, falls back to portfolio-based matching from PROPERTIES.
// Stores diagnostic info on _pgDiag for the Reload Groups shift-click report.
async function resolvePropertyGroupNames() {
  // Reset lookup maps
  _nameToGroups = {};
  _idToGroups = {};
  _uuidToGroups = {};
  // Diagnostic collector — shift-click "Reload Groups" shows this
  var diag = { uuidMapSize: 0, dbNameCount: 0, uuidHits: 0, uuidMisses: 0,
    nameMatches: 0, idMatches: 0, portfolioMatches: 0, errors: [] };

  var uuidMapOk = false;

  // Step 1: UUID resolution via property_map
  // Builds: UUID → group names, DB API name → group names
  try {
    var mapData;
    try {
      mapData = await proxyAction('property_map');
    } catch (mapErr) {
      // Fallback: build UUID→name map from DB API properties via passthrough.
      var propsPath = '/api/v0/properties?filters[LastUpdatedAtFrom]=' +
        encodeURIComponent(PROPERTY_GROUPS_LAST_UPDATED_FROM) +
        '&page[size]=' + encodeURIComponent(PROPERTY_GROUPS_PAGE_SIZE);
      console.log('[PG] property_map action failed; building map from properties fallback: ' + (mapErr.message || mapErr));
      var dbProps = await fetchAllPages(propsPath);
      var uuidMapFallback = {};
      dbProps.forEach(function(p) {
        var pid = p.Id || p.id || '';
        var pname = p.Name || p.name || p.PropertyName || p.property_name || '';
        if (pid && pname) uuidMapFallback[String(pid)] = { name: String(pname) };
      });
      mapData = { property_uuid_map: uuidMapFallback };
    }
    console.log('[PG] property_map response keys: ' + Object.keys(mapData).join(', '));
    var uuidMap = mapData.property_uuid_map || {};
    diag.uuidMapSize = Object.keys(uuidMap).length;
    console.log('[PG] UUID map entries: ' + diag.uuidMapSize);
    if (diag.uuidMapSize > 0) {
      var sampleKey = Object.keys(uuidMap)[0];
      console.log('[PG] UUID map sample: ' + sampleKey + ' → ' + JSON.stringify(uuidMap[sampleKey]));
    }

    // Also build a reverse map: DB API name (lowercase, trimmed) → [UUIDs]
    var dbNameToUuids = {};
    Object.keys(uuidMap).forEach(function(uuid) {
      var entry = uuidMap[uuid];
      var name = '';
      if (entry) {
        name = typeof entry === 'string' ? entry : (entry.name || entry.Name || '');
      }
      name = name.trim();
      if (name) {
        var key = name.toLowerCase();
        if (!dbNameToUuids[key]) dbNameToUuids[key] = [];
        dbNameToUuids[key].push(uuid);
      }
    });
    diag.dbNameCount = Object.keys(dbNameToUuids).length;
    console.log('[PG] Reverse name→UUID map: ' + diag.dbNameCount + ' distinct names');

    PROPERTY_GROUPS.forEach(function(g) {
      if (!g.name) return;
      var props = Array.isArray(g.properties) ? g.properties : [];
      var resolvedNames = [];

      props.forEach(function(uuid) {
        // Map UUID → group name
        _addToGroupMap(_uuidToGroups, uuid, g.name);

        // Resolve UUID → property name via the map
        var mapped = uuidMap[uuid];
        var mName = '';
        if (mapped) {
          mName = (typeof mapped === 'string' ? mapped : (mapped.name || mapped.Name || '')).trim();
        }
        if (mName) {
          diag.uuidHits++;
          if (resolvedNames.indexOf(mName) === -1) resolvedNames.push(mName);
          // Index by DB API name (lowercase, trimmed)
          _addToGroupMap(_nameToGroups, mName.toLowerCase(), g.name);
        } else {
          diag.uuidMisses++;
        }
      });

      g.resolvedNames = resolvedNames;
      g.propertyNames = resolvedNames;
    });

    uuidMapOk = diag.uuidMapSize > 0;
    console.log('[PG] UUID hits: ' + diag.uuidHits + ', misses: ' + diag.uuidMisses);

    // Step 1b: Bridge DB API UUIDs to Reports API property_id via name matching.
    // For each PROPERTIES entry, find its UUID(s) by matching name, then index by id.
    if (uuidMapOk) {
      PROPERTIES.forEach(function(p) {
        if (!p.name) return;
        var pNameLower = (p.name || '').trim().toLowerCase();
        // Check if this Reports API property name matches any DB API property name
        var matchedUuids = dbNameToUuids[pNameLower];
        if (matchedUuids && matchedUuids.length > 0) {
          diag.nameMatches++;
          // Store the first UUID on the property for reference
          p._dbUuid = matchedUuids[0];
          // For each matched UUID, copy that UUID's group memberships to the property_id
          matchedUuids.forEach(function(uuid) {
            var uGroups = _uuidToGroups[uuid];
            if (uGroups) {
              uGroups.forEach(function(gn) {
                _addToGroupMap(_idToGroups, String(p.id), gn);
                // Also ensure the Reports API name is indexed
                _addToGroupMap(_nameToGroups, pNameLower, gn);
              });
            }
          });
        }
      });
      diag.idMatches = Object.keys(_idToGroups).length;
    }

    console.log('[PG] Step 1 done — nameMap: ' + Object.keys(_nameToGroups).length +
      ', idMap: ' + Object.keys(_idToGroups).length +
      ', nameMatches(props↔dbNames): ' + diag.nameMatches + '/' + PROPERTIES.length);
  } catch (err) {
    var errMsg = 'UUID map failed: ' + (err.message || err);
    diag.errors.push(errMsg);
    console.log('[PG] ' + errMsg + ' (will try portfolio fallback)');
  }

  // Step 2: Portfolio fallback — supplements UUID mapping.
  // If a group's name matches a PROPERTIES portfolio, index those properties too.
  try {
    PROPERTY_GROUPS.forEach(function(g) {
      if (!g.name) return;

      // If UUID resolution worked for this group, also check resolved names for portfolios
      if (g.resolvedNames && g.resolvedNames.length > 0) {
        g.resolvedNames.forEach(function(rn) {
          var prop = PROPERTIES.find(function(p) {
            return (p.name || '').trim().toLowerCase() === String(rn).trim().toLowerCase();
          });
          if (prop && prop.id) {
            _addToGroupMap(_idToGroups, String(prop.id), g.name);
            _addToGroupMap(_nameToGroups, (prop.name || '').trim().toLowerCase(), g.name);
          }
        });
      }

      // If UUID resolution didn't populate this group, try portfolio name match
      if (!uuidMapOk || !g.resolvedNames || g.resolvedNames.length === 0) {
        PROPERTIES.forEach(function(p) {
          var portfolio = (p.portfolio || '').trim();
          var gName = g.name.trim();
          if (portfolio && (portfolio === gName || portfolio.toLowerCase() === gName.toLowerCase())) {
            diag.portfolioMatches++;
            _addToGroupMap(_nameToGroups, (p.name || '').trim().toLowerCase(), g.name);
            _addToGroupMap(_idToGroups, String(p.id), g.name);
          }
        });
      }
    });

    console.log('[PG] Step 2 done — final nameMap: ' + Object.keys(_nameToGroups).length +
      ', idMap: ' + Object.keys(_idToGroups).length +
      ', portfolioMatches: ' + diag.portfolioMatches);
  } catch (err) {
    var pfErr = 'Portfolio fallback error: ' + (err.message || err);
    diag.errors.push(pfErr);
    console.log('[PG] ' + pfErr);
  }

  // Store diagnostics for the shift-click report
  window._pgDiag = diag;
}

// Recent Tasks: Proxy ?action=recent_tasks — server-side pagination via DB API v0
async function fetchRecentTasks() {
  try {
    setApiStatus('loading', 'Loading recent tasks\u2026');
    var data = await proxyAction('recent_tasks');
    var results = data.results || data.data || [];
    RECENT_TASKS = results.map(function(t) {
      return {
        id: t.Id || t.id || '',
        taskType: t.TaskType || t.task_type || '',
        subject: t.Subject || t.subject || '',
        body: t.Body || t.body || '',
        status: t.Status || t.status || '',
        assignee: t.AssignedTo || t.assigned_to || '',
        dueDate: t.DueDate || t.due_date || '',
        completedDate: t.CompletedDate || t.completed_date || '',
        createdAt: t.CreatedAt || t.created_at || '',
        updatedAt: t.UpdatedAt || t.updated_at || '',
        priority: t.Priority || t.priority || '',
        propertyName: t.PropertyName || t.property_name || '',
        unitName: t.UnitName || t.unit_name || '',
        linkedResourceType: t.LinkedResourceType || t.linked_resource_type || '',
        linkedResourceId: t.LinkedResourceId || t.linked_resource_id || ''
      };
    });
    // Sort by most recent
    RECENT_TASKS.sort(function(a, b) {
      return new Date(b.updatedAt || b.createdAt || 0) - new Date(a.updatedAt || a.createdAt || 0);
    });
    return true;
  } catch (err) {
    console.log('fetchRecentTasks error: ' + (err.message || err));
    RECENT_TASKS = [];
    return false;
  }
}

// Upcoming Move-Outs: Proxy ?action=upcoming_moveouts — tenant directory
// Returns tenants on notice or current with move_out dates in the window
async function fetchUpcomingMoveouts() {
  try {
    setApiStatus('loading', 'Loading upcoming move-outs\u2026');
    var data = await proxyAction('upcoming_moveouts', { days: 60 });
    var results = data.results || data.data || [];
    UPCOMING_MOVEOUTS = results.map(function(t) {
      return {
        property: t.property || t.property_name || '',
        propertyId: t.property_id || 0,
        unit: t.unit || '',
        unitId: t.unit_id || 0,
        tenant: t.tenant || '',
        status: t.status || '',
        moveOut: t.move_out || '',
        moveIn: t.move_in || '',
        phone: t.phone_numbers || '',
        email: t.emails || '',
        rent: t.rent || '',
        occupancyId: t.occupancy_id || ''
      };
    });
    return true;
  } catch (err) {
    console.log('fetchUpcomingMoveouts error: ' + (err.message || err));
    UPCOMING_MOVEOUTS = [];
    return false;
  }
}

// Turn Work Orders: Proxy ?action=turn_work_orders — DB API v0
// Real-time WO status for unit turn WOs (more current than Reports API)
async function fetchTurnWorkOrders() {
  try {
    setApiStatus('loading', 'Loading turn work orders\u2026');
    var data = await proxyAction('turn_work_orders', { days: 90 });
    var results = data.results || data.data || [];
    TURN_WORK_ORDERS = results.map(function(wo) {
      return {
        id: wo.Id || wo.id || '',
        unitId: wo.UnitId || wo.unit_id || '',
        propertyId: wo.PropertyId || wo.property_id || '',
        status: wo.Status || wo.status || '',
        type: wo.Type || wo.type || '',
        priority: wo.Priority || wo.priority || '',
        description: wo.JobDescription || wo.Description || wo.description || '',
        vendorId: wo.VendorId || wo.vendor_id || '',
        vendorTrade: wo.VendorTrade || '',
        createdAt: wo.CreatedAt || wo.created_at || '',
        lastUpdated: wo.LastUpdatedAt || wo.last_updated_at || '',
        workCompletedOn: wo.WorkCompletedOn || wo.work_completed_on || '',
        scheduledStart: wo.ScheduledStart || wo.scheduled_start || '',
        assignedUsers: wo.AssignedUsers || [],
        woNumber: wo.WorkOrderNumber || wo.work_order_number || '',
        link: wo.Link || ''
      };
    });
    return true;
  } catch (err) {
    console.log('fetchTurnWorkOrders error: ' + (err.message || err));
    TURN_WORK_ORDERS = [];
    return false;
  }
}

// Unit Turns (SQL tracker): Proxy ?action=unit_turns
// Stores inferred + manual turn tracking records synced from the pipeline.

async function fetchUnitTurnsDB() {
  try {
    setApiStatus('loading', 'Loading unit turn tracker…');
    var data = await proxyAction('unit_turns', { days: '180' });
    var results = data.results || data.data || [];
    UNIT_TURN_TRACKER_BY_KEY = {};
    UNIT_TURNS_DB = results.map(function(t) {
      var linked = Array.isArray(t.linked_work_orders) ? t.linked_work_orders : [];
      var normalized = {
        id: t.tracking_uuid || t.Id || t.id || '',
        trackingUuid: t.tracking_uuid || '',
        trackingCode: t.tracking_code || '',
        turnKey: t.turn_key || '',
        unitTurnId: t.unit_turn_id || t.UnitTurnId || '',
        unitId: t.unit_id || t.UnitId || t.unitId || '',
        propertyId: t.property_id || t.PropertyId || t.propertyId || '',
        unit: t.unit_name || t.UnitName || t.unit || '',
        property: t.property_name || t.PropertyName || t.property || '',
        moveOut: t.move_out_date || t.MoveOutDate || t.moveOut || '',
        moveIn: t.move_in_date || t.MoveInDate || t.moveIn || '',
        inspectionDate: t.inspection_date || '',
        firstWoDate: t.first_wo_date || '',
        estimateRequestedDate: t.estimate_requested_date || '',
        estimateReceivedDate: t.estimate_received_date || '',
        expectedMoveIn: t.move_in_date || t.expected_move_in_date || t.ExpectedMoveInDate || t.expectedMoveIn || '',
        status: t.status || t.Status || '',
        confidenceScore: parseInt(t.confidence_score || t.confidenceScore || 0, 10) || 0,
        confidenceLabel: t.confidence_label || t.confidenceLabel || 'low',
        depositStatus: t.deposit_status || t.DepositStatus || '',
        depositAmount: t.deposit_amount || t.DepositAmount || '',
        depositReturnDeadline: t.deposit_return_deadline || t.DepositReturnDeadline || '',
        totalBilled: t.total_billed || t.TotalBilled || '',
        laborCost: t.labor_cost || t.LaborCost || '',
        siteManager: t.site_manager || t.SiteManager || '',
        targetDays: parseInt(t.target_days_to_complete || t.TargetDaysToComplete || 0, 10) || 0,
        turnEnd: t.turn_end_date || t.TurnEndDate || t.turnEnd || '',
        link: t.link || t.Link || '',
        milestones: t.milestones || {},
        linkedWorkOrders: linked.map(function(w) {
          return {
            id: w.wo_id || w.id || '',
            dbApiId: w.wo_db_uuid || '',
            source: w.source || 'manual',
            status: w.status || '',
            created: w.created_at || ''
          };
        })
      };
      if (normalized.turnKey) UNIT_TURN_TRACKER_BY_KEY[normalized.turnKey] = normalized;
      return normalized;
    });

    // Merge tracker data into TURNS array by unit_turn_id or unit/property key
    if (UNIT_TURNS_DB.length > 0 && TURNS.length > 0) {
      var dbById = {};
      UNIT_TURNS_DB.forEach(function(u) {
        if (u.id) dbById[String(u.id)] = u;
        if (u.unitTurnId) dbById[String(u.unitTurnId)] = u;
      });
      var dbByUnitProp = {};
      UNIT_TURNS_DB.forEach(function(u) {
        var k = String(u.unitId || '').toLowerCase() + '|' + String(u.propertyId || '').toLowerCase();
        if (k !== '|') dbByUnitProp[k] = u;
      });
      TURNS.forEach(function(turn) {
        var dbMatch = (turn.unitTurnId && dbById[String(turn.unitTurnId)]) ||
          dbByUnitProp[String(turn.unitId || '').toLowerCase() + '|' + String(turn.propertyId || '').toLowerCase()];
        if (!dbMatch) return;
        // Augment with live fields if not already set
        if (!turn.depositStatus) turn.depositStatus = dbMatch.depositStatus;
        if (!turn.depositAmount) turn.depositAmount = dbMatch.depositAmount;
        if (!turn.depositReturnDeadline) turn.depositReturnDeadline = dbMatch.depositReturnDeadline;
        if (!turn.expectedMoveIn && dbMatch.expectedMoveIn) turn.expectedMoveIn = dbMatch.expectedMoveIn;
        if (!turn.moveIn && dbMatch.moveIn) turn.moveIn = dbMatch.moveIn;
        if (!turn.turnEnd && dbMatch.turnEnd) turn.turnEnd = dbMatch.turnEnd;
        if (dbMatch.link) turn.dbLink = dbMatch.link;
      });
    }
    return true;
  } catch (err) {
    console.log('fetchUnitTurnsDB error: ' + (err.message || err));
    UNIT_TURNS_DB = [];
    return false;
  }
}

/* =================================================================
   WEBHOOK EVENTS — HTTP POST relay for Make.com / Zapier / etc.
   ================================================================= */
async function pollWebhookEvents() {
  try {
    var data = await proxyAction('webhook_events');
    if (data && Array.isArray(data.events)) {
      // Merge new events (dedup by AppFolio event_id when present)
      var existing = {};
      WEBHOOK_EVENTS.forEach(function(e) { existing[webhookEventDedupeKey(e)] = true; });
      data.events.forEach(function(e) {
        var key = webhookEventDedupeKey(e);
        if (!existing[key]) {
          var evtBody = e.body || '';
          if (evtBody.length > 500) evtBody = evtBody.substring(0, 500) + '\u2026';
          var meta = extractWebhookMeta(e);
          // Decode event type to friendly label
          var friendlyTitle = decodeWebhookTitle(e, meta);
          WEBHOOK_EVENTS.push({
            id: e.id || 0,
            ts: e.ts || e.timestamp || new Date().toISOString(),
            type: meta.eventType || e.type || e.event_type || 'webhook',
            event_label: e.event_label || friendlyTitle,
            resource_type: meta.resourceType || '',
            resource_id: meta.resourceId || '',
            resource_name: meta.resourceName || '',
            title: friendlyTitle,
            body: evtBody,
            priority: e.priority || 'normal',
            source: e.source || 'appfolio'
          });
          existing[key] = true;
        }
      });
      // Sort newest first (timezone-safe parser)
      WEBHOOK_EVENTS.sort(function(a, b) {
        var bt = parseWebhookTs(b.ts);
        var at = parseWebhookTs(a.ts);
        return (bt ? bt.getTime() : 0) - (at ? at.getTime() : 0);
      });
      // Keep last 200 after a final dedupe pass
      WEBHOOK_EVENTS = dedupeWebhookEvents(WEBHOOK_EVENTS);
      if (WEBHOOK_EVENTS.length > 200) WEBHOOK_EVENTS = WEBHOOK_EVENTS.slice(0, 200);
    }
    return true;
  } catch (err) {
    console.log('Webhook poll error: ' + (err.message || err));
    return false;
  }
}

// Decode webhook event to human-readable title using proxy data + client-side resolution
var WEBHOOK_LABELS = {
  'work_order.created': 'Work Order Created',
  'work_order.updated': 'Work Order Updated',
  'work_order.deleted': 'Work Order Deleted',
  'unit_turn.created': 'Unit Turn Started',
  'unit_turn.updated': 'Unit Turn Updated',
  'unit.created': 'Unit Created',
  'unit.updated': 'Unit Updated',
  'property.created': 'Property Added',
  'property.updated': 'Property Updated',
  'inspection.created': 'Inspection Scheduled',
  'inspection.updated': 'Inspection Updated',
  'task.created': 'Task Created',
  'task.updated': 'Task Updated',
  'tenant.created': 'Tenant Added',
  'tenant.updated': 'Tenant Updated',
  'vendor.created': 'Vendor Added',
  'vendor.updated': 'Vendor Updated',
  'lease.created': 'Lease Created',
  'lease.updated': 'Lease Updated'
};

var WEBHOOK_EVENT_VERBS = {
  create: 'Created',
  created: 'Created',
  update: 'Updated',
  updated: 'Updated',
  delete: 'Deleted',
  deleted: 'Deleted'
};

// resource_type/resource_id -> resolved AppFolio record (via proxy action=webhook_resolve)
var WEBHOOK_RESOURCE_CACHE = {};
var WEBHOOK_RESOURCE_INFLIGHT = {};

function webhookResolveKey(resourceType, resourceId) {
  var rt = String(resourceType || '').toLowerCase().trim();
  var rid = String(resourceId || '').trim();
  if (!rt || !rid) return '';
  return rt + ':' + rid;
}

async function resolveWebhookResource(resourceType, resourceId) {
  var key = webhookResolveKey(resourceType, resourceId);
  if (!key) return null;
  var normalizedType = normalizeWebhookResourceType(resourceType, '');
  if (normalizedType === 'work_order' && !isUuidString(resourceId)) return null;
  if (WEBHOOK_RESOURCE_CACHE[key]) return WEBHOOK_RESOURCE_CACHE[key];
  if (WEBHOOK_RESOURCE_INFLIGHT[key]) return WEBHOOK_RESOURCE_INFLIGHT[key];

  WEBHOOK_RESOURCE_INFLIGHT[key] = (async function() {
    try {
      var res = await proxyAction('webhook_resolve', {
        resource_type: resourceType,
        resource_id: resourceId
      });
      if (res && res.ok) {
        WEBHOOK_RESOURCE_CACHE[key] = {
          ts: Date.now(),
          summary: res.summary || {},
          record: res.record || {},
          endpoint: res.endpoint || '',
          domain: res.domain || ''
        };
      }
    } catch (e) {
      // Keep quiet; unresolved IDs should never break webhook rendering.
    } finally {
      delete WEBHOOK_RESOURCE_INFLIGHT[key];
    }
    return WEBHOOK_RESOURCE_CACHE[key] || null;
  })();

  return WEBHOOK_RESOURCE_INFLIGHT[key];
}

function getWebhookChangeSummary(payload) {
  var p = payload || {};
  if (Array.isArray(p.changed_fields) && p.changed_fields.length > 0) {
    return 'Fields changed: ' + p.changed_fields.join(', ');
  }
  if (Array.isArray(p.changed_attributes) && p.changed_attributes.length > 0) {
    return 'Fields changed: ' + p.changed_attributes.join(', ');
  }
  if (p.changes && typeof p.changes === 'object') {
    var keys = Object.keys(p.changes);
    if (keys.length > 0) return 'Fields changed: ' + keys.join(', ');
  }
  if (p.delta && typeof p.delta === 'object') {
    var dkeys = Object.keys(p.delta);
    if (dkeys.length > 0) return 'Fields changed: ' + dkeys.join(', ');
  }
  return 'Field-level change details were not provided by this webhook payload.';
}

function parseWebhookJson(raw) {
  if (!raw) return null;
  if (typeof raw === 'object') return raw;
  if (typeof raw !== 'string') return null;
  var s = raw.trim();
  if (!s) return null;
  try { return JSON.parse(s); } catch (e1) { /* ignore */ }
  try {
    var cleaned = s.replace(/^"|"$/g, '').replace(/\\"/g, '"').replace(/""/g, '"');
    return JSON.parse(cleaned);
  } catch (e2) { /* ignore */ }
  return null;
}

function singularizeTopic(topic) {
  var t = String(topic || '').trim().toLowerCase();
  if (!t) return '';
  if (t.slice(-3) === 'ies') return t.slice(0, -3) + 'y';
  if (t.slice(-1) === 's') return t.slice(0, -1);
  return t;
}

function normalizeWebhookResourceType(rawType, topic) {
  var rt = String(rawType || '').trim().toLowerCase();
  if (rt && rt.indexOf('.') !== -1) rt = rt.split('.')[0];
  rt = rt.replace(/[^a-z0-9_]/g, '');
  if (!rt) rt = singularizeTopic(topic || '');
  return singularizeTopic(rt);
}

function titleCaseWords(str) {
  return String(str || '').replace(/_/g, ' ').replace(/\b\w/g, function(c) { return c.toUpperCase(); });
}

function resolveWebhookResourceName(resourceType, resourceId) {
  if (!resourceId) return '';
  var rid = String(resourceId);
  var rtype = normalizeWebhookResourceType(resourceType, '');
  var ckey = webhookResolveKey(rtype, rid);
  var cached = ckey ? WEBHOOK_RESOURCE_CACHE[ckey] : null;
  if (cached && cached.summary) {
    var s = cached.summary;
    var base = s.title || s.reference || '';
    if (base && s.reference && base !== s.reference) base += ' #' + s.reference;
    if (base) return base;
  }

  if (rtype === 'work_order') {
    var wo = WORK_ORDERS.find(function(w) { return String(w.uuid) === rid || String(w.id) === rid; });
    if (wo) return 'WO#' + (wo.id || rid.slice(0, 8)) + (wo.propertyName ? ' - ' + wo.propertyName : '');
  }
  if (rtype === 'property') {
    var prop = PROPERTIES.find(function(p) { return String(p._dbUuid) === rid || String(p.id) === rid; });
    if (prop) return prop.name || '';
  }
  if (rtype === 'vendor') {
    var vendor = VENDORS.find(function(v) { return String(v.id) === rid; });
    if (vendor) return vendor.name || '';
  }
  if (rtype === 'unit_turn') {
    var turn = TURNS.find(function(t) { return String(t.unitTurnId) === rid || String(t.unitId) === rid; });
    if (turn) return (turn.unit || rid.slice(0, 8)) + (turn.property ? ' - ' + turn.property : '');
  }
  if (rtype === 'property_group') {
    var grp = PROPERTY_GROUPS.find(function(g) { return String(g.id) === rid; });
    if (grp) return grp.name || '';
  }
  return '';
}

function getWebhookReadableContext(meta, e) {
  meta = meta || extractWebhookMeta(e || {});
  var payload = (meta && meta.payload) || {};
  var parts = [];
  var number = payload.work_order_number || payload.number || payload.reference_number || '';
  var name = payload.title || payload.name || payload.inspection_name || payload.subject || '';
  var propertyName = payload.property_name || payload.property || payload.property_address || payload.address || '';
  var unit = payload.unit_name || payload.unit || '';
  var person = payload.tenant_name || payload.tenant || payload.vendor_name || payload.assigned_to || '';
  var status = payload.status || payload.work_order_status || '';

  if (number) parts.push('#' + number);
  if (name && (!number || String(name).indexOf(String(number)) === -1)) parts.push(name);
  if (propertyName) parts.push(propertyName);
  if (unit) parts.push(unit);
  if (person) parts.push(person);
  if (status) parts.push('Status: ' + titleCaseWords(status));

  return parts.filter(Boolean).join(' • ').substring(0, 140);
}

function getWebhookPreviewText(e, meta) {
  meta = meta || extractWebhookMeta(e || {});
  var view = interpretWebhookEvent(e || {}, meta);
  return (view && view.description) || getWebhookReadableContext(meta, e) || '';
}

function extractWebhookMeta(e) {
  var payload = parseWebhookJson(e.payload) || parseWebhookJson(e.raw) || parseWebhookJson(e.body) || {};
  var topic = String(e.topic || payload.topic || '').trim().toLowerCase();
  var eventType = String(e.event_type || payload.event_type || e.type || '').trim().toLowerCase();
  var resourceType = normalizeWebhookResourceType(e.resource_type || payload.resource_type || '', topic);
  var resourceId = String(e.resource_id || payload.resource_id || '').trim();
  var resourceName = String(e.resource_name || payload.resource_name || '').trim();
  if (!resourceName && resourceId) resourceName = resolveWebhookResourceName(resourceType, resourceId);
  return {
    payload: payload,
    topic: topic,
    eventType: eventType,
    resourceType: resourceType,
    resourceId: resourceId,
    resourceName: resourceName
  };
}

var WEBHOOK_INVALIDATION_MAP = {
  work_order: ['work_orders', 'turn_work_orders', 'recent_tasks', 'labor'],
  unit_turn: ['turns', 'turn_work_orders'],
  vendor: ['vendors'],
  inspection: ['inspections'],
  tenant: ['upcoming_moveouts'],
  lease: ['upcoming_moveouts'],
  property: ['properties', 'property_groups', 'property_map'],
  bill: ['bills']
};

function normalizeWebhookAction(meta, e) {
  var raw = String((meta && meta.eventType) || e.event_type || e.type || '').toLowerCase().trim();
  if (!raw) return 'updated';
  if (raw.indexOf('.') !== -1) raw = raw.split('.').slice(1).join('.');
  if (raw === 'work_completed') return 'completed';
  if (raw === 'status_changed') return 'updated';
  return raw;
}

function interpretWebhookEvent(e, meta) {
  meta = meta || extractWebhookMeta(e || {});
  var payload = meta.payload || {};
  var action = normalizeWebhookAction(meta, e || {});
  var category = String(meta.resourceType || singularizeTopic(meta.topic || '') || 'unknown').toLowerCase();
  var status = String(payload.status || payload.work_order_status || '').toLowerCase();

  var severity = 'info';
  if (/(cancel|failed|declin|error)/.test(action) || /(cancel|failed|declin|error)/.test(status)) severity = 'warning';
  if (/(complete|closed|approved|renewed|move_in)/.test(action) || /(complete|closed|approved)/.test(status)) severity = 'success';

  var iconClass = 'fa-circle-info';
  if (category === 'work_order') iconClass = 'fa-wrench';
  else if (category === 'unit_turn') iconClass = 'fa-arrows-rotate';
  else if (category === 'inspection') iconClass = 'fa-clipboard-check';
  else if (category === 'tenant') iconClass = 'fa-user';
  else if (category === 'lease') iconClass = 'fa-file-contract';
  else if (category === 'vendor') iconClass = 'fa-hard-hat';
  else if (category === 'property') iconClass = 'fa-building';

  var title = decodeWebhookTitle(e, meta);
  var description = '';
  var woNum = payload.work_order_number || payload.number || '';
  var addr = payload.property_address || payload.address || '';
  var who = payload.assigned_to || payload.tenant_name || payload.vendor_name || '';
  if (category === 'work_order') {
    description = [woNum ? 'WO #' + woNum : '', addr ? 'at ' + addr : '', status ? '\u2192 ' + status : '', who ? 'for ' + who : ''].filter(Boolean).join(' ');
  } else {
    description = [payload.name || '', payload.property_name || '', payload.unit || ''].filter(Boolean).join(' / ');
  }

  return {
    title: title || 'Webhook Event',
    description: description || e.body || '',
    category: category,
    action: action,
    severity: severity,
    iconClass: iconClass,
    invalidates: WEBHOOK_INVALIDATION_MAP[category] || []
  };
}

function shouldToastWebhookEvent(view) {
  if (!view) return false;
  if (view.severity === 'success' || view.severity === 'warning') return true;
  if (view.category === 'work_order' && /(created|assigned|updated|completed|canceled|cancelled)/.test(view.action)) return true;
  return false;
}

function dispatchWebhookInvalidations(keys, sourceEventCount) {
  if (!keys || !keys.length) return;
  window.dispatchEvent(new CustomEvent('handymgr:webhook-invalidate', {
    detail: {
      keys: keys,
      eventCount: sourceEventCount || 0,
      at: new Date().toISOString()
    }
  }));
}

// Build an AppFolio-style human-readable title from a resolved record.
// Returns null when insufficient fields are present (caller falls back to generic label).
function buildAppFolioStyleTitle(resourceType, eventType, record) {
  if (!record || !resourceType) return null;
  var r = record;
  var rtype = String(resourceType || '').toLowerCase().replace(/[^a-z_]/g, '');
  var etype = String(eventType || '').toLowerCase();
  var isCreate = etype === 'create' || etype === 'created';
  var isUpdate = etype === 'update' || etype === 'updated';
  var isDelete = etype === 'delete' || etype === 'deleted';

  if (rtype === 'work_order') {
    var woTitle = r.Title || r.Description || '';
    var num     = r.Number ? ' #' + r.Number : '';
    var prop    = r.PropertyName || '';
    var unit    = r.UnitName || r.Unit || '';
    var loc     = [prop, unit].filter(Boolean).join(' / ');
    var submitter = r.SubmittedByName || r.CreatedByName || r.CreatedBy || r.ReportedBy || '';
    var rawCat  = r.Category || r.ServiceType || '';
    var category = rawCat.replace(/_/g, ' ').replace(/\b\w/g, function(c) { return c.toUpperCase(); });
    if (isCreate) {
      var objLabel = category || 'Work Order';
      var parts = [];
      if (submitter) parts.push('from ' + submitter);
      if (loc)       parts.push(loc);
      return 'New ' + objLabel + (parts.length ? ' ' + parts.join(' / ') : (woTitle ? ' - ' + woTitle : '')) + num;
    }
    if (isUpdate) return 'Work Order Updated' + (woTitle ? ' - ' + woTitle : (loc ? ' - ' + loc : '')) + num;
    if (isDelete) return 'Work Order Removed' + (woTitle ? ' - ' + woTitle : '') + num;
    return (woTitle || 'Work Order') + num + (loc ? ' \u2014 ' + loc : '');
  }

  if (rtype === 'vendor') {
    var vName = r.Name || r.CompanyName || '';
    if (!vName) return null;
    if (isCreate) return 'Vendor Added \u2014 ' + vName;
    if (isUpdate) return 'Vendor Updated \u2014 ' + vName;
    if (isDelete) return 'Vendor Removed \u2014 ' + vName;
    return vName;
  }

  if (rtype === 'tenant') {
    var tName = r.FullName || r.Name || [r.FirstName, r.LastName].filter(Boolean).join(' ') || '';
    var tUnit = r.UnitName || r.Unit || '';
    var tProp = r.PropertyName || '';
    var tLoc  = [tProp, tUnit].filter(Boolean).join(' / ');
    if (!tName && !tLoc) return null;
    if (isCreate) return 'Tenant Added' + (tName ? ' \u2014 ' + tName : '') + (tLoc ? ' | ' + tLoc : '');
    if (isUpdate) return 'Tenant Updated' + (tName ? ' \u2014 ' + tName : '');
    return tName || 'Tenant';
  }

  if (rtype === 'unit_turn') {
    var utUnit  = r.Unit || r.UnitName || '';
    var utProp  = r.PropertyName || '';
    var utStage = r.Stage || '';
    var utLoc   = [utProp, utUnit].filter(Boolean).join(' / ');
    if (!utLoc) return null;
    if (isCreate) return 'Turn Started \u2014 ' + utLoc;
    if (isUpdate) return 'Turn Updated' + (utStage ? ' \u2192 ' + utStage : '') + ' \u2014 ' + utLoc;
    return 'Unit Turn \u2014 ' + utLoc;
  }

  if (rtype === 'lease') {
    var lTenants = r.TenantNames || r.TenantName || r.Name || '';
    if (Array.isArray(lTenants)) lTenants = lTenants.join(', ');
    var lUnit = r.UnitName || r.Unit || '';
    var lProp = r.PropertyName || '';
    var lLoc  = [lProp, lUnit].filter(Boolean).join(' / ');
    if (isCreate) return 'Lease Created' + (lTenants ? ' \u2014 ' + lTenants : '') + (lLoc ? ' / ' + lLoc : '');
    if (isUpdate) return 'Lease Updated'  + (lTenants ? ' \u2014 ' + lTenants : '');
    return 'Lease' + (lTenants ? ' \u2014 ' + lTenants : '');
  }

  if (rtype === 'inspection') {
    var iName = r.Name || r.InspectionName || r.Title || '';
    var iProp = r.PropertyName || '';
    if (isCreate) return 'Inspection Scheduled' + (iName ? ' \u2014 ' + iName : '') + (iProp ? ' / ' + iProp : '');
    if (isUpdate) return 'Inspection Updated'   + (iName ? ' \u2014 ' + iName : '');
    return 'Inspection' + (iName ? ' \u2014 ' + iName : '');
  }

  if (rtype === 'property') {
    var pName = r.Name || r.PropertyName || '';
    if (!pName) return null;
    if (isCreate) return 'Property Added \u2014 ' + pName;
    if (isUpdate) return 'Property Updated \u2014 ' + pName;
    return pName;
  }

  if (rtype === 'task') {
    var tskTitle = r.Title || r.Name || r.Subject || '';
    var tskWho   = r.AssignedToName || r.AssignedTo || '';
    if (isCreate) return 'Task Created' + (tskTitle ? ' \u2014 ' + tskTitle : '') + (tskWho ? ' for ' + tskWho : '');
    if (isUpdate) return 'Task Updated'  + (tskTitle ? ' \u2014 ' + tskTitle : '');
    return tskTitle || null;
  }

  return null;
}

function decodeWebhookTitle(e, meta) {
  meta = meta || extractWebhookMeta(e);
  var hasResolvedName = !!(meta && meta.resourceName);
  var looksLikeFallbackProxyTitle = false;
  if (e && e.title) {
    // Example fallback: "work_orders.update -> work_orders/uuid"
    looksLikeFallbackProxyTitle = /(?:\u2192|->)\s*[a-z_]+\/[0-9a-f-]{8,}/i.test(String(e.title));
  }
  // Preserve explicit non-fallback labels unless we now have a resolved resource name.
  if (!hasResolvedName && e.event_label && e.event_label !== 'webhook' && e.event_label !== e.type) return e.event_label;
  if (!hasResolvedName && e.title && e.title !== 'Webhook Event' && !looksLikeFallbackProxyTitle) return e.title;

  // When a resolved record is cached, build an AppFolio-style rich title.
  var rckey = (meta.resourceType && meta.resourceId) ? webhookResolveKey(meta.resourceType, meta.resourceId) : '';
  var rcCached = rckey ? WEBHOOK_RESOURCE_CACHE[rckey] : null;
  if (rcCached && rcCached.record) {
    var richTitle = buildAppFolioStyleTitle(meta.resourceType, meta.eventType, rcCached.record);
    if (richTitle) return richTitle;
  }

  var combinedKey = '';
  if (meta.resourceType && meta.eventType) combinedKey = meta.resourceType + '.' + meta.eventType;

  // Decode from event type
  var label = WEBHOOK_LABELS[combinedKey] || WEBHOOK_LABELS[e.type] || WEBHOOK_LABELS[e.event_type] || '';
  if (!label && meta.resourceType && meta.eventType) {
    label = titleCaseWords(meta.resourceType) + ' ' + (WEBHOOK_EVENT_VERBS[meta.eventType] || titleCaseWords(meta.eventType));
  }
  if (!label && meta.topic && meta.eventType) {
    label = titleCaseWords(singularizeTopic(meta.topic)) + ' ' + (WEBHOOK_EVENT_VERBS[meta.eventType] || titleCaseWords(meta.eventType));
  }
  if (!label && e.type) {
    label = String(e.type).replace(/_/g, ' ').replace(/\./g, ' ').replace(/\b\w/g, function(c) { return c.toUpperCase(); });
  }

  // Try to resolve resource name
  var resName = meta.resourceName || '';
  if (!resName && meta.resourceId) {
    resName = meta.resourceId.substring(0, 8) + '\u2026';
  }
  return (label || 'Webhook Event') + (resName ? ' - ' + resName : '');
}

function webhookEventDedupeKey(e) {
  var payload = parseWebhookJson(e && (e.payload || e.raw || e.body)) || {};
  var eid = String((e && e.event_id) || payload.event_id || '').trim();
  if (eid) return 'event_id:' + eid;
  var rt = String((e && e.resource_type) || payload.resource_type || payload.topic || '').trim().toLowerCase();
  var rid = String((e && e.resource_id) || payload.resource_id || '').trim();
  var et = String((e && e.type) || (e && e.event_type) || payload.event_type || '').trim().toLowerCase();
  var ts = String((e && e.ts) || payload.event_timestamp || payload.message_sent_at || '').trim();
  return rt + '|' + rid + '|' + et + '|' + ts;
}

function dedupeWebhookEvents(list) {
  var out = [];
  var seen = {};
  (list || []).forEach(function(e) {
    var k = webhookEventDedupeKey(e);
    if (seen[k]) return;
    seen[k] = true;
    out.push(e);
  });
  return out;
}

function renderWebhookEventList() {
  var el = $('#webhookEventList');
  var countEl = $('#webhookEventCount');
  if (countEl) countEl.textContent = WEBHOOK_EVENTS.length;
  if (!el) return;
  if (WEBHOOK_EVENTS.length === 0) {
    el.innerHTML = 'No events yet \u2014 POST to the webhook URL to push events.';
    return;
  }
  var html = '';
  WEBHOOK_EVENTS.slice(0, 25).forEach(function(e, idx) {
    var isPri = e.priority === 'urgent' || e.priority === 'high';
    var iconClass = 'fa-plug';
    var iconColor = 'var(--purple)';
    var rtype = e.resource_type || '';
    if (rtype === 'work_order') { iconClass = 'fa-wrench'; iconColor = 'var(--accent)'; }
    else if (rtype === 'unit_turn') { iconClass = 'fa-exchange-alt'; iconColor = 'var(--info,#60a5fa)'; }
    else if (rtype === 'inspection') { iconClass = 'fa-clipboard-check'; iconColor = 'var(--success)'; }
    else if (rtype === 'tenant') { iconClass = 'fa-user'; iconColor = 'var(--warning)'; }
    else if (rtype === 'vendor') { iconClass = 'fa-hard-hat'; iconColor = 'var(--success)'; }
    else if (rtype === 'property') { iconClass = 'fa-building'; iconColor = 'var(--accent)'; }
    var detailId = 'whPanelDetail_' + (e.id || idx);
    var meta = extractWebhookMeta(e);
    var previewText = getWebhookPreviewText(e, meta);
    var rckey = (meta.resourceType && meta.resourceId) ? webhookResolveKey(meta.resourceType, meta.resourceId) : '';
    var resolved = rckey ? WEBHOOK_RESOURCE_CACHE[rckey] : null;

    // Build detail content
    var richTitle = (resolved && resolved.record) ? buildAppFolioStyleTitle(meta.resourceType, meta.eventType, resolved.record) : null;
    var objectLine = richTitle || (resolved && resolved.summary && resolved.summary.title) || '';
    if (!objectLine && meta.resourceId) objectLine = (meta.resourceType || '') + ' / ' + meta.resourceId;
    var statusLine = (resolved && resolved.summary && resolved.summary.status) || '';
    var refLine    = (resolved && resolved.summary && resolved.summary.reference) ? '#' + resolved.summary.reference : '';
    var changes    = getWebhookChangeSummary(meta.payload || {});
    var rec        = (resolved && resolved.record) || {};
    var ctxParts   = [];
    if (rec.PropertyName) ctxParts.push(rec.PropertyName);
    if (rec.UnitName || rec.Unit) ctxParts.push(rec.UnitName || rec.Unit);
    if (rec.AssignedToName || rec.VendorName) ctxParts.push(rec.AssignedToName || rec.VendorName);

    html += '<div style="border-bottom:1px solid var(--border)">';
    // Header row — clickable to expand
    html += '<div class="wh-panel-item" data-whpanel="' + detailId + '" style="padding:5px 2px;display:flex;gap:6px;align-items:flex-start;cursor:pointer;user-select:none">';
    html += '<i class="fas ' + iconClass + '" style="color:' + iconColor + ';margin-top:2px;font-size:11px;width:14px;text-align:center;flex-shrink:0"></i>';
    html += '<div style="flex:1;min-width:0">';
    html += '<span style="color:var(--text-muted);font-size:10px">' + escapeHtml(e.ts ? timeAgo(e.ts) : '\u2014') + '</span> ';
    if (isPri) html += '<span style="color:var(--danger);font-weight:600">\u26a0 </span>';
    html += '<strong style="color:var(--text-primary);font-size:11px">' + escapeHtml(e.title || decodeWebhookTitle(e, meta)) + '</strong>';
    if (previewText) html += '<div style="color:var(--text-secondary);margin-top:1px;font-size:10px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + escapeHtml(previewText) + '</div>';
    html += '</div>';
    html += '<i class="fas fa-chevron-down" style="font-size:9px;color:var(--text-muted);margin-top:3px;flex-shrink:0;transition:transform .15s" id="' + detailId + '_chev"></i>';
    html += '</div>';

    // Collapsible detail panel
    html += '<div id="' + detailId + '" style="display:none;padding:6px 8px 8px 20px;font-size:10px;font-family:var(--font-mono);color:var(--text-secondary);background:var(--bg-input);line-height:1.7">';
    if (objectLine) html += '<div><span style="color:var(--text-muted)">Object:</span> ' + escapeHtml(objectLine + (refLine ? ' ' + refLine : '')) + (statusLine ? ' <span style="color:var(--text-muted)">(' + escapeHtml(statusLine) + ')</span>' : '') + '</div>';
    if (ctxParts.length) html += '<div><span style="color:var(--text-muted)">Where:</span> ' + escapeHtml(ctxParts.join(' \u00b7 ')) + '</div>';
    html += '<div><span style="color:var(--text-muted)">Change:</span> ' + escapeHtml(changes) + '</div>';
    if (e.body)    html += '<div style="margin-top:2px"><span style="color:var(--text-muted)">Body:</span> ' + escapeHtml(e.body) + '</div>';
    if (e.raw) {
      var rawStr = typeof e.raw === 'string' ? e.raw : JSON.stringify(e.raw, null, 2);
      html += '<pre style="margin:4px 0 0;white-space:pre-wrap;word-break:break-all;max-height:100px;overflow-y:auto;background:var(--bg-tertiary);padding:4px;border-radius:3px;font-size:9px">' + escapeHtml(rawStr) + '</pre>';
    }
    html += '</div>';
    html += '</div>';
  });
  if (WEBHOOK_EVENTS.length > 25) {
    html += '<div style="padding:4px 0;color:var(--text-muted);text-align:center;font-size:10px">\u2026 and ' + (WEBHOOK_EVENTS.length - 25) + ' more</div>';
  }
  el.innerHTML = html;

  // Wire up expand/collapse via delegation on the list container
  el.querySelectorAll('.wh-panel-item').forEach(function(row) {
    row.addEventListener('click', function() {
      var detailId = row.getAttribute('data-whpanel');
      var detail = document.getElementById(detailId);
      var chev   = document.getElementById(detailId + '_chev');
      if (!detail) return;
      var open = detail.style.display !== 'none';
      detail.style.display = open ? 'none' : 'block';
      if (chev) chev.style.transform = open ? '' : 'rotate(180deg)';
    });
  });
}

/* =================================================================
   WEBHOOK DATA REVIEW — SQLite-backed data table with filters
   ================================================================= */
var _whPage = 0;
var _whPageSize = 50;
var _whTotal = 0;
var _whFilters = { search: '', type: '', source: '', from: '', to: '' };

async function loadWebhookData() {
  var body = $('#whDataBody');
  if (!body) return;
  body.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:20px;color:var(--text-muted)"><i class="fas fa-circle-notch fa-spin"></i> Loading webhook data\u2026</td></tr>';
  try {
    var params = {
      limit: _whPageSize,
      offset: _whPage * _whPageSize
    };
    if (_whFilters.search) params.search = _whFilters.search;
    if (_whFilters.type) params.type = _whFilters.type;
    if (_whFilters.source) params.source = _whFilters.source;
    if (_whFilters.from) params.from = _whFilters.from + 'T00:00:00Z';
    if (_whFilters.to) params.to = _whFilters.to + 'T23:59:59Z';
    var data = await proxyAction('webhook_events', params);
    var events = (data.events || []).map(function(evt) {
      var meta = extractWebhookMeta(evt);
      var title = decodeWebhookTitle(evt, meta);
      return Object.assign({}, evt, {
        type: meta.eventType || evt.type || evt.event_type || 'webhook',
        resource_type: meta.resourceType || evt.resource_type || '',
        resource_id: meta.resourceId || evt.resource_id || '',
        resource_name: meta.resourceName || evt.resource_name || '',
        title: title,
        event_label: evt.event_label || title
      });
    });
    events = dedupeWebhookEvents(events);
    if (_whFilters.source) {
      events = events.filter(function(evt) {
        return String(evt.source || '').toLowerCase() === String(_whFilters.source).toLowerCase();
      });
    }
    _whTotal = _whFilters.source ? events.length : (data.total || events.length);
    // Update badge
    var badge = $('#whBadge');
    if (badge) badge.textContent = _whTotal || '\u2014';
    renderWebhookDataTable(events);
    updateWhPagination();
    // Populate filter dropdowns from data (first load)
    populateWhFilterDropdowns(events);

    // Progressive enrichment: resolve opaque resource IDs into human-readable objects.
    // Cap per page to avoid flooding AppFolio rate limits.
    (async function enrichVisibleWebhookEvents() {
      var unresolved = [];
      events.forEach(function(evt) {
        var rt = String(evt.resource_type || '').trim();
        var rid = String(evt.resource_id || '').trim();
        var key = webhookResolveKey(rt, rid);
        if (!key) return;
        if (WEBHOOK_RESOURCE_CACHE[key] || WEBHOOK_RESOURCE_INFLIGHT[key]) return;
        unresolved.push({ rt: rt, rid: rid });
      });

      if (unresolved.length === 0) return;
      var maxResolve = Math.min(20, unresolved.length);
      var changed = false;
      for (var i = 0; i < maxResolve; i++) {
        var item = unresolved[i];
        var got = await resolveWebhookResource(item.rt, item.rid);
        if (got) changed = true;
        // Keep under AppFolio 8 req/s limit with small jittered delay.
        await sleep(140 + Math.floor(Math.random() * 40));
      }
      if (!changed) return;

      // Recompute titles with resolved names and re-render current page.
      events = events.map(function(evt) {
        var meta = extractWebhookMeta(evt);
        var title = decodeWebhookTitle(evt, meta);
        return Object.assign({}, evt, {
          resource_name: meta.resourceName || evt.resource_name || '',
          title: title,
          event_label: evt.event_label || title
        });
      });
      renderWebhookDataTable(events);
    })();
  } catch (err) {
    body.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:20px;color:var(--danger)">Failed to load: ' + escapeHtml(err.message || String(err)) + '</td></tr>';
  }
}

function renderWebhookDataTable(events) {
  var body = $('#whDataBody');
  if (!body) return;
  if (events.length === 0) {
    body.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:20px;color:var(--text-muted)">No webhook events found</td></tr>';
    return;
  }
  var html = '';
  events.forEach(function(e, idx) {
    var meta = extractWebhookMeta(e);
    var ckey = webhookResolveKey(meta.resourceType || e.resource_type, meta.resourceId || e.resource_id);
    var resolved = ckey ? WEBHOOK_RESOURCE_CACHE[ckey] : null;
    var changeSummary = getWebhookChangeSummary(meta.payload || {});
    var isPri = e.priority === 'urgent' || e.priority === 'high';
    var priClass = isPri ? 'color:var(--danger);font-weight:600' : 'color:var(--text-secondary)';
    var rowNum = _whPage * _whPageSize + idx + 1;
    html += '<tr class="wh-data-row" data-whid="' + (e.id || idx) + '" style="cursor:pointer">';
    html += '<td style="font-family:var(--font-mono);font-size:11px;color:var(--text-muted)">' + rowNum + '</td>';
    html += '<td style="font-family:var(--font-mono);font-size:11px;white-space:nowrap">' + escapeHtml(e.ts ? timeAgo(e.ts) : '\u2014') + '</td>';
    html += '<td><span class="tag wh-type-' + escapeHtml(String(e.type || 'webhook').replace(/[^a-z0-9_-]/gi, '')) + '">' + escapeHtml(e.type || 'webhook') + '</span></td>';
    html += '<td style="max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + escapeHtml(e.title || '\u2014');
    // v8 enriched data: show resource_type/resource_id if available
    if (e.resource_type) {
      html += ' <span style="font-size:10px;color:var(--text-muted)">(' + escapeHtml(e.resource_type);
      if (e.resource_id) html += ':' + escapeHtml(String(e.resource_id).substring(0, 8));
      html += ')</span>';
    }
    if (e.body_status === 'empty') {
      html += ' <span style="font-size:9px;color:var(--warning)" title="Webhook body was empty (pre-v8 bug)">\u26a0 empty</span>';
    }
    html += '</td>';
    html += '<td style="' + priClass + '">' + escapeHtml(e.priority || 'normal') + '</td>';
    html += '<td style="font-size:11px;color:var(--text-muted)">' + escapeHtml(e.source || '\u2014') + '</td>';
    html += '<td><button class="action-btn" style="padding:2px 6px;font-size:10px" data-whexpand="' + (e.id || idx) + '" title="View raw"><i class="fas fa-eye"></i></button></td>';
    html += '</tr>';
    // Expandable detail row (hidden by default)
    html += '<tr class="wh-detail-row hidden" id="whDetail_' + (e.id || idx) + '">';
    html += '<td colspan="7" style="padding:10px 14px;background:var(--bg-input);border-bottom:2px solid var(--accent)">';
    html += '<div style="font-size:11px;font-family:var(--font-mono);line-height:1.5">';
    html += '<strong>Object:</strong> ';
    if (resolved && resolved.record) {
      var richObj = buildAppFolioStyleTitle(meta.resourceType, meta.eventType, resolved.record);
      if (richObj) {
        html += escapeHtml(richObj);
      } else if (resolved.summary) {
        html += escapeHtml((resolved.summary.title || 'Resolved') + (resolved.summary.reference ? ' #' + resolved.summary.reference : ''));
      }
      if (resolved.summary && resolved.summary.status) {
        html += ' <span style="color:var(--text-muted)">(' + escapeHtml(resolved.summary.status) + ')</span>';
      }
      // Contextual details line: property · unit · assigned vendor
      var rec = resolved.record;
      var ctxParts = [];
      if (rec.PropertyName) ctxParts.push(rec.PropertyName);
      if (rec.UnitName || rec.Unit) ctxParts.push(rec.UnitName || rec.Unit);
      if (rec.AssignedToName || rec.VendorName) ctxParts.push('Assigned: ' + (rec.AssignedToName || rec.VendorName));
      if (rec.Priority) ctxParts.push('Priority: ' + rec.Priority);
      if (ctxParts.length > 0) {
        html += '<br><span style="font-size:10px;color:var(--text-muted);padding-left:8px">' + escapeHtml(ctxParts.join(' \u00b7 ')) + '</span>';
      }
    } else if (resolved && resolved.summary) {
      html += escapeHtml((resolved.summary.title || 'Resolved') + (resolved.summary.reference ? ' #' + resolved.summary.reference : ''));
      if (resolved.summary.status) html += ' <span style="color:var(--text-muted)">(' + escapeHtml(resolved.summary.status) + ')</span>';
    } else if ((meta.resourceType || e.resource_type) && (meta.resourceId || e.resource_id)) {
      html += escapeHtml(String(meta.resourceType || e.resource_type)) + ' / ' + escapeHtml(String(meta.resourceId || e.resource_id));
    } else {
      html += '(unknown)';
    }
    html += '<br>';
    html += '<strong>Change:</strong> ' + escapeHtml(changeSummary) + '<br>';
    html += '<strong>Body:</strong> ' + escapeHtml(e.body || '(empty)') + '<br>';
    if (e.raw) {
      var rawStr = typeof e.raw === 'string' ? e.raw : JSON.stringify(e.raw);
      html += '<strong>Raw JSON:</strong><pre style="margin:4px 0;padding:6px;background:var(--bg-tertiary);border-radius:4px;overflow-x:auto;max-height:120px;font-size:10px">' + escapeHtml(rawStr) + '</pre>';
    }
    html += '</div></td></tr>';
  });
  body.innerHTML = html;
  // Event listeners handled by delegation in wireUpUI() — no re-attachment needed
}

function updateWhPagination() {
  var info = $('#whPaginationInfo');
  var prevBtn = $('#btnWhPrev');
  var nextBtn = $('#btnWhNext');
  var start = _whPage * _whPageSize + 1;
  var end = Math.min(start + _whPageSize - 1, _whTotal);
  if (_whTotal === 0) { start = 0; end = 0; }
  if (info) info.textContent = start + '\u2013' + end + ' of ' + _whTotal + ' events';
  if (prevBtn) prevBtn.disabled = (_whPage === 0);
  if (nextBtn) nextBtn.disabled = (end >= _whTotal);
}

function populateWhFilterDropdowns(events) {
  var typeSelect = $('#whTypeFilter');
  var sourceSelect = $('#whSourceFilter');
  if (!typeSelect || !sourceSelect) return;
  // Only populate if still at defaults (one option)
  if (typeSelect.options.length <= 1) {
    var types = {};
    var sources = {};
    events.forEach(function(e) {
      if (e.type) types[e.type] = true;
      if (e.source) sources[e.source] = true;
    });
    Object.keys(types).sort().forEach(function(t) {
      var opt = document.createElement('option');
      opt.value = t;
      opt.textContent = t;
      typeSelect.appendChild(opt);
    });
    Object.keys(sources).sort().forEach(function(s) {
      var opt = document.createElement('option');
      opt.value = s;
      opt.textContent = s;
      sourceSelect.appendChild(opt);
    });
  }
}

async function loadWebhookStats() {
  var panel = $('#whStatsPanel');
  var content = $('#whStatsContent');
  if (!panel || !content) return;
  panel.classList.remove('hidden');
  content.innerHTML = '<i class="fas fa-circle-notch fa-spin"></i> Loading stats\u2026';
  try {
    var data = await proxyAction('webhook_stats');
    var html = '<div style="display:flex;gap:20px;flex-wrap:wrap">';
    html += '<div><strong style="font-size:18px;color:var(--accent)">' + (data.total || 0) + '</strong><div style="color:var(--text-muted)">Total Events</div></div>';
    // v8: show has_data vs empty_body counts
    if (data.has_data !== undefined) {
      html += '<div><strong style="font-size:18px;color:var(--success)">' + (data.has_data || 0) + '</strong><div style="color:var(--text-muted)">With Data</div></div>';
      html += '<div><strong style="font-size:18px;color:var(--warning)">' + (data.empty_body || 0) + '</strong><div style="color:var(--text-muted)">Empty (pre-v8)</div></div>';
    }
    // By type
    if (data.by_type && data.by_type.length > 0) {
      html += '<div><strong>By Type:</strong><div style="margin-top:4px">';
      data.by_type.forEach(function(t) {
        html += '<span class="tag" style="margin:2px">' + escapeHtml(t.type) + ' <strong>' + t.count + '</strong></span> ';
      });
      html += '</div></div>';
    }
    // By source
    if (data.by_source && data.by_source.length > 0) {
      html += '<div><strong>By Source:</strong><div style="margin-top:4px">';
      data.by_source.forEach(function(s) {
        html += '<span class="tag" style="margin:2px">' + escapeHtml(s.source) + ' <strong>' + s.count + '</strong></span> ';
      });
      html += '</div></div>';
    }
    html += '</div>';
    // By day (mini chart using bar divs)
    if (data.by_day && data.by_day.length > 0) {
      var maxCount = Math.max.apply(null, data.by_day.map(function(d) { return d.count; }));
      html += '<div style="margin-top:10px"><strong>Last 30 Days:</strong>';
      html += '<div style="display:flex;align-items:flex-end;gap:2px;height:60px;margin-top:6px">';
      data.by_day.slice().reverse().forEach(function(d) {
        var pct = maxCount > 0 ? Math.max(4, (d.count / maxCount) * 100) : 4;
        html += '<div style="flex:1;background:var(--accent);border-radius:2px 2px 0 0;height:' + pct + '%;min-width:4px;opacity:0.8" title="' + escapeHtml(d.day) + ': ' + d.count + '"></div>';
      });
      html += '</div></div>';
    }
    // v8+: append server-side cache stats
    if (supportsServerCacheOps()) {
      html += '<div style="margin-top:12px;padding-top:10px;border-top:1px solid var(--border)">';
      html += '<strong style="color:var(--accent)"><i class="fas fa-database"></i> Server Cache (' + _proxyVersion + ')</strong>';
      try {
        var cStats = await proxyAction('cache_stats');
        if (cStats.cache && cStats.cache.length > 0) {
          html += '<table style="width:100%;margin-top:6px;font-size:12px;border-collapse:collapse">';
          html += '<tr style="color:var(--text-muted);text-align:left"><th style="padding:3px 6px">Entity</th><th style="padding:3px 6px">Entries</th><th style="padding:3px 6px">Records</th><th style="padding:3px 6px">Last Cached</th></tr>';
          cStats.cache.forEach(function(c) {
            html += '<tr style="border-top:1px solid var(--border)">';
            html += '<td style="padding:3px 6px;font-family:var(--font-mono)">' + escapeHtml(c.entity_type) + '</td>';
            html += '<td style="padding:3px 6px">' + (c.entries || 0) + '</td>';
            html += '<td style="padding:3px 6px">' + (c.total_records || 0) + '</td>';
            html += '<td style="padding:3px 6px;color:var(--text-muted)">' + (c.last_cached ? timeAgo(c.last_cached) : '\u2014') + '</td>';
            html += '</tr>';
          });
          html += '</table>';
        } else {
          html += '<div style="margin-top:4px;color:var(--text-muted)">No cached data yet</div>';
        }
        if (cStats.webhooks) {
          html += '<div style="margin-top:6px;font-size:11px;color:var(--text-muted)">';
          html += 'Webhooks: ' + (cStats.webhooks.total || 0) + ' total, ' + (cStats.webhooks.pending || 0) + ' pending';
          if (cStats.turn_records !== undefined) html += ' | Turn records: ' + cStats.turn_records;
          html += '</div>';
        }
      } catch (csErr) {
        html += '<div style="margin-top:4px;color:var(--warning)">Cache stats unavailable: ' + escapeHtml(csErr.message || String(csErr)) + '</div>';
      }
      html += '</div>';
    }
    content.innerHTML = html;
  } catch (err) {
    content.innerHTML = '<span style="color:var(--danger)">Failed: ' + escapeHtml(err.message || String(err)) + '</span>';
  }
}

async function migrateWebhookBlob() {
  var btn = $('#btnWhMigrate');
  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Migrating\u2026'; }
  try {
    var data = await proxyAction('webhook_migrate');
    showToast('Migrated ' + (data.migrated || 0) + ' events from blob to SQLite');
    loadWebhookData();
  } catch (err) {
    showToast('Migration failed: ' + (err.message || err));
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = '<i class="fas fa-database"></i> Migrate Blob'; }
  }
}

/* =================================================================
   WEBHOOK ORGANIZER ENGINE — processWebhookEvent()
   Transforms passive webhook log into active data processor:
   • Matches WO IDs in event body via regex
   • Auto-updates WO status / priority in WORK_ORDERS
   • Auto-flags urgent items
   • Advances turn pipeline stages in TURN_PIPE_DATA
   ================================================================= */
function processWebhookEvent(evt) {
  if (!evt || !evt.body) return;
  var body = String(evt.body);
  var title = String(evt.title || '');
  var combined = title + ' ' + body;
  var changed = false;

  // 1. Extract WO number(s) from event body — supports numbers like 12345 and 12345-1.
  var woMatches = combined.match(/(?:WO[#\-:\s]?|work.?order[#\-:\s]?)(\d{3,}(?:-\d+)?)/gi) || [];
  var woMap = {};
  woMatches.forEach(function(m) {
    var digits = m.match(/(\d{3,}(?:-\d+)?)/);
    if (!digits) return;
    var full = digits[1];
    woMap[full] = true;
    woMap[full.replace(/-\d+$/, '')] = true;
  });
  var woNumbers = Object.keys(woMap).filter(Boolean);

  // Also try UUID patterns for DB API links
  var uuidMatches = combined.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi) || [];

  // 2. Detect status keywords in event
  var statusKeywords = {
    'completed': 'Completed',
    'complete': 'Completed',
    'closed': 'Completed',
    'in progress': 'In Progress',
    'in-progress': 'In Progress',
    'started': 'In Progress',
    'assigned': 'Assigned',
    'approved': 'Approved',
    'on hold': 'On Hold',
    'canceled': 'Canceled',
    'cancelled': 'Canceled',
    'estimate received': 'Estimate Received',
    'estimate requested': 'Estimate Requested',
    'bid received': 'Estimate Received',
    'bid requested': 'Estimate Requested',
    'new': 'New'
  };

  var detectedStatus = null;
  var combinedLower = combined.toLowerCase();
  Object.keys(statusKeywords).forEach(function(kw) {
    if (combinedLower.indexOf(kw) !== -1) {
      detectedStatus = statusKeywords[kw];
    }
  });

  // 3. Detect priority escalation
  var isUrgent = /\b(urgent|emergency|critical|asap|immediate)\b/i.test(combined);
  if (isUrgent && evt.priority !== 'urgent') {
    evt.priority = 'urgent';
  }

  // 4. Match and update WORK_ORDERS
  woNumbers.forEach(function(woNum) {
    var wo = WORK_ORDERS.find(function(w) {
      var workOrderId = String(w.id || '');
      return workOrderId === woNum || workOrderId.replace(/-\d+$/, '') === woNum;
    });
    if (!wo) return;

    // Update status if detected
    if (detectedStatus && wo.status !== detectedStatus) {
      wo.status = detectedStatus;
      changed = true;
    }

    // Flag urgent WOs
    if (isUrgent) {
      wo.priority = 'Urgent';
      if (!WO_FLAGS[wo.id]) WO_FLAGS[wo.id] = {};
      WO_FLAGS[wo.id].urgent = true;
      WO_FLAGS[wo.id].urgentTs = evt.ts || new Date().toISOString();
      changed = true;
    }

    // Append webhook note to statusNotes
    var note = '[Webhook ' + (evt.ts ? timeAgo(evt.ts) : 'now') + '] ' + title;
    if (wo.statusNotes) {
      wo.statusNotes = note + ' | ' + wo.statusNotes;
    } else {
      wo.statusNotes = note;
    }
  });

  // 5. Match UUIDs against TURN_WORK_ORDERS and advance pipeline
  uuidMatches.forEach(function(uuid) {
    var tWo = TURN_WORK_ORDERS.find(function(tw) { return tw.id === uuid; });
    if (!tWo) return;

    // If status detected, update DB-API-sourced WO too
    if (detectedStatus && tWo.status !== detectedStatus) {
      tWo.status = detectedStatus;
      changed = true;
    }
  });

  // 6. Advance TURN_PIPE_DATA stages based on status changes
  if (detectedStatus && TURN_PIPE_DATA.length > 0) {
    var stageMap = {
      'New': 'wo_created',
      'Estimate Requested': 'est_requested',
      'Estimate Received': 'est_received',
      'Approved': 'assigned',
      'Assigned': 'assigned',
      'In Progress': 'assigned',
      'Completed': 'work_done'
    };
    var targetStage = stageMap[detectedStatus];
    if (targetStage) {
      woNumbers.forEach(function(woNum) {
        TURN_PIPE_DATA.forEach(function(p) {
          // Match by WO numbers in the pipeline entry
          var pipeWos = p.workOrders || [];
          var hasWo = pipeWos.some(function(pw) { return String(pw.id || pw.woNumber) === woNum; });
          if (hasWo) {
            var targetIdx = PIPE_STAGES.findIndex(function(ps) { return ps.key === targetStage; });
            if (targetIdx > -1 && targetIdx > p.currentStageIdx) {
              p.currentStageIdx = targetIdx;
              p.currentStage = PIPE_STAGES[targetIdx].key;
              changed = true;
            }
          }
        });
      });
    }
  }

  return changed;
}

var _lastWebhookMaxId = 0; // track highest seen event ID to skip redundant processing
function setupWebhookAutoPoll(intervalSec) {
  if (_webhookPollTimer) { clearInterval(_webhookPollTimer); _webhookPollTimer = null; }
  if (intervalSec > 0) {
    _webhookPollTimer = setInterval(function() {
      // Wrap in async IIFE to avoid blocking the interval handler
      (async function() {
        var prevCount = WEBHOOK_EVENTS.length;
        var ok = await pollWebhookEvents();
        if (!ok) return;
        // Check if any genuinely new events arrived (by max ID)
        var maxId = 0;
        WEBHOOK_EVENTS.forEach(function(e) { if (e.id && e.id > maxId) maxId = e.id; });
        if (maxId <= _lastWebhookMaxId) return; // No new events — skip all re-renders
        _lastWebhookMaxId = maxId;
        // Process only NEW events through the organizer engine
        var newEvents = WEBHOOK_EVENTS.slice(0, WEBHOOK_EVENTS.length - prevCount);
        // Skip events with empty body/raw (common with placeholder webhooks)
        var meaningfulEvents = newEvents.filter(function(evt) {
          return (evt.body && String(evt.body).trim().length > 0) || (evt.type && String(evt.type).trim().length > 0);
        });
        var interpreted = meaningfulEvents.map(function(evt) {
          return interpretWebhookEvent(evt, extractWebhookMeta(evt));
        });
        var anyChanged = false;
        meaningfulEvents.forEach(function(evt) {
          var didChange = processWebhookEvent(evt);
          if (didChange) anyChanged = true;
        });
        var invalidateMap = {};
        interpreted.forEach(function(v) {
          (v.invalidates || []).forEach(function(k) { invalidateMap[k] = true; });
        });
        var invalidateKeys = Object.keys(invalidateMap);
        // Defer renders to next animation frame to avoid blocking
        requestAnimationFrame(function() {
          renderWebhookEventList();
          renderActivityFeed();
          if (anyChanged) {
            renderWorkOrders();
            renderTurnBoard();
            renderDashboardKPIs();
          }
          if (invalidateKeys.length > 0) {
            dispatchWebhookInvalidations(invalidateKeys, meaningfulEvents.length);
          }
          var toastables = interpreted.filter(shouldToastWebhookEvent);
          if (toastables.length > 0) {
            var top = toastables[0];
            var suffix = toastables.length > 1 ? ' +' + (toastables.length - 1) + ' more' : '';
            showToast(top.title + suffix, {
              kind: top.severity,
              iconClass: top.iconClass,
              duration: 4200
            });
          } else if (anyChanged && meaningfulEvents.length > 0) {
            showToast('Webhook updated ' + meaningfulEvents.length + ' event(s)', { kind: 'info', iconClass: 'fa-bolt' });
          }
        });
      })();
    }, intervalSec * 1000);
  }
}

/* =================================================================
   DETAIL FETCHERS — On-demand for WO detail panel
   ================================================================= */
async function fetchPropertyDetail(propId) {
  if (!propId) return null;
  var cached = detailCacheGet('prop_' + propId);
  if (cached) return cached;
  // Reports API returns numeric IDs (e.g. 6393), but DB API v0 requires UUIDs.
  // Only attempt fetch if propId looks like a UUID; otherwise return from PROPERTIES.
  var isUuid = /^[0-9a-f]{8}-[0-9a-f]{4}-/i.test(String(propId));
  if (!isUuid) {
    // Try to find in already-loaded PROPERTIES instead of hitting the API
    var localProp = PROPERTIES.find(function(p) {
      return String(p.id) === String(propId);
    });
    if (localProp) { detailCacheSet('prop_' + propId, localProp); return localProp; }
    return null; // Don't call DB API with a numeric ID — it'll 404
  }
  try {
    var data = await apiFetch('/api/v0/properties/' + propId);
    detailCacheSet('prop_' + propId, data);
    return data;
  } catch (e) { return null; }
}

// Resolve a WO's DB API UUID — needed because Reports API returns numeric IDs
// while DB API v0 endpoints require UUIDs
function resolveWODbUuid(wo) {
  if (!wo) return '';
  // 1. Check if wo already has a DB API link (contains UUID in path)
  if (wo.link) {
    var linkMatch = String(wo.link).match(/work_orders\/([0-9a-f\-]{36})/i);
    if (linkMatch) return linkMatch[1];
  }
  // 2. Look up matching DB API work order by WO number
  if (wo.id) {
    var dbWo = TURN_WORK_ORDERS.find(function(tw) {
      return String(tw.woNumber) === String(wo.id);
    });
    if (dbWo && dbWo.id) return dbWo.id;
  }
  // 3. Fallback: use wo.uuid (may be numeric — DB API may reject it)
  return wo.uuid || '';
}

async function fetchWONotes(woIdOrUuid) {
  if (!woIdOrUuid) return [];
  var woRef = String(woIdOrUuid || '').trim();
  if (!isUuidString(woRef)) {
    var dbWo = TURN_WORK_ORDERS.find(function(tw) {
      return String(tw.woNumber || '') === woRef && isUuidString(tw.id);
    });
    woRef = dbWo ? String(dbWo.id) : '';
  }
  if (!woRef || !isUuidString(woRef)) return [];
  var notesCached = detailCacheGet('notes_' + woRef);
  if (typeof notesCached !== 'undefined') return notesCached;
  try {
    // Use dedicated proxy action with v0 credentials (like property_groups)
    var data = await proxyAction('wo_notes', { wo_id: woRef });
    var notes = (data && data.results) ? data.results :
                (data && data.data) ? data.data :
                (Array.isArray(data) ? data : []);
    detailCacheSet('notes_' + woRef, notes);
    return notes;
  } catch (e) { return []; }
}

function renderWONotesList(notes) {
  var nl = document.getElementById('detailNotesList');
  if (!nl) return;
  if (!notes || notes.length === 0) {
    nl.innerHTML = '<div style="text-align:center;padding:10px;color:var(--text-muted);font-size:12px">No notes</div>';
    return;
  }
  var nh = '';
  notes.forEach(function(n) {
    var createdBy = n.CreatedBy || n.created_by || n.Author || n.author || n.UserName || '—';
    var createdAt = n.CreatedAt || n.created_at || n.UpdatedAt || n.updated_at || '';
    var body = n.Body || n.body || n.Content || n.content || n.Note || n.note || n.Message || n.message || '';
    nh += '<div class="note-item"><div class="note-item-header"><span>' + escapeHtml(createdBy) + '</span><span>' + formatDate(createdAt) + '</span></div>';
    nh += '<div class="note-item-body">' + escapeHtml(body) + '</div></div>';
  });
  nl.innerHTML = nh;
}

function refreshCurrentWONotes(forceUuid) {
  var modal = document.getElementById('woModal');
  if (!modal || !modal.classList.contains('show') || !CURRENT_WO_MODAL) return;
  var targetUuid = forceUuid || CURRENT_WO_MODAL.woDbUuid;
  if (!targetUuid) return;
  delete WO_DETAIL_CACHE['notes_' + targetUuid];
  fetchWONotes(targetUuid).then(function(notes) {
    if (!CURRENT_WO_MODAL || CURRENT_WO_MODAL.woDbUuid !== targetUuid) return;
    renderWONotesList(notes);
  });
}

/* =================================================================
   RENDER FUNCTIONS — All use live data
   ================================================================= */

function getUpcomingMoveOuts() {
  var today = new Date();
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() + 60);
  var results = [];
  var seen = {};

  // Primary: use UPCOMING_MOVEOUTS from tenant directory (most accurate)
  UPCOMING_MOVEOUTS.forEach(function(mo) {
    if (!mo.moveOut) return;
    var moDate = new Date(mo.moveOut);
    if (isNaN(moDate.getTime())) return;
    if (moDate >= today && moDate <= cutoff) {
      var key = (mo.property || '') + '|' + (mo.unit || '');
      if (seen[key]) return;
      seen[key] = true;
      var daysLeft = Math.round((moDate - today) / 86400000);
      results.push({ property: mo.property, unit: mo.unit, tenant: mo.tenant, moveOut: mo.moveOut, daysLeft: daysLeft, phone: mo.phone || '', rent: mo.rent || '' });
    }
  });

  // Fallback: also check inspections for move-outs not in UPCOMING_MOVEOUTS
  INSPECTIONS.forEach(function(r) {
    if (!r.moveOut) return;
    var moDate = new Date(r.moveOut);
    if (isNaN(moDate.getTime())) return;
    if (moDate >= today && moDate <= cutoff) {
      var key = (r.propertyName || '') + '|' + (r.unit || '');
      if (seen[key]) return;
      seen[key] = true;
      var daysLeft = Math.round((moDate - today) / 86400000);
      results.push({ property: r.propertyName, unit: r.unit, tenant: r.tenant, moveOut: r.moveOut, daysLeft: daysLeft });
    }
  });

  results.sort(function(a, b) { return a.daysLeft - b.daysLeft; });
  return results;
}

function renderMoveOuts() {
  var body = $('#moveOutBody');
  if (!body) return;
  var moves = getUpcomingMoveOuts();
  if (moves.length === 0) {
    body.innerHTML = '<tr><td colspan="5" style="text-align:center;color:var(--text-muted);padding:16px;font-size:12px">No upcoming move-outs in the next 60 days</td></tr>';
    return;
  }
  var html = '';
  moves.forEach(function(m) {
    var urgClass = m.daysLeft <= 14 ? 'moveout-urgent' : m.daysLeft <= 30 ? 'moveout-soon' : 'moveout-normal';
    html += '<tr>';
    html += '<td>' + escapeHtml(m.property) + '</td>';
    html += '<td>' + escapeHtml(m.unit) + '</td>';
    html += '<td>' + escapeHtml(m.tenant || '\u2014') + '</td>';
    html += '<td style="font-family:var(--font-mono)">' + formatDate(m.moveOut) + '</td>';
    html += '<td><span class="tag ' + urgClass + '">' + m.daysLeft + 'd</span></td>';
    html += '</tr>';
  });
  body.innerHTML = html;
}

function getDashboardTurnPageSize() {
  if (document.body && document.body.classList.contains('tv-mode')) return 6;
  if (window.matchMedia && window.matchMedia('(max-width: 700px)').matches) return 1;
  if (window.matchMedia && window.matchMedia('(max-width: 1100px)').matches) return 2;
  return 4;
}

function ensureNavigationChromeVisible() {
  var topbar = document.querySelector('.topbar');
  var navTabs = document.getElementById('navTabs') || document.querySelector('.nav-tabs');
  if (topbar) topbar.style.display = '';
  if (navTabs) navTabs.style.display = '';
}

function syncTvModeScope() {
  if (!document.body) return;
  var dashSection = document.getElementById('sec-dashboard');
  var onDashboard = !!(dashSection && dashSection.classList.contains('active'));
  document.body.classList.toggle('tv-mode-dashboard', document.body.classList.contains('tv-mode') && onDashboard);
}

function applyTvMode(enabled) {
  if (!document.body) return;
  if (enabled) {
    forceActiveTab('dashboard');
  }
  document.body.classList.toggle('tv-mode', !!enabled);
  syncTvModeScope();
  if (!enabled) ensureNavigationChromeVisible();
  try { localStorage.setItem('hm_tv_mode', enabled ? '1' : '0'); } catch (e) { /* */ }
  var btn = $('#dashTvMode');
  if (btn) {
    btn.innerHTML = enabled
      ? '<i class="fas fa-tv"></i> Exit TV Mode'
      : '<i class="fas fa-tv"></i> TV Mode';
  }
  DASH_TURN_PAGE = 0;
  renderTurnDashboardStrip();
}

function getDashboardTurnEntries() {
  return TURN_PIPE_DATA.filter(function(p) {
    if (p.isClosed || p.isCompleted) return false;
    if (!isInPropertyGroup(p.propertyId, p.property, currentPropertyGroup)) return false;
    var pmName = String(p.siteManager || 'Unassigned PM').trim() || 'Unassigned PM';
    if (DASH_TURN_PM_FILTER && pmName !== DASH_TURN_PM_FILTER) return false;
    return true;
  }).sort(function(a, b) {
    if (a.isStalled !== b.isStalled) return a.isStalled ? -1 : 1;
    if (a.isUpcoming !== b.isUpcoming) return a.isUpcoming ? 1 : -1;
    return b.elapsed - a.elapsed;
  });
}

function syncDashboardTurnPmFilter(entries) {
  var sel = $('#dashTurnPmFilter');
  if (!sel) return;
  var counts = {};
  entries.forEach(function(p) {
    var pmName = String(p.siteManager || 'Unassigned PM').trim() || 'Unassigned PM';
    counts[pmName] = (counts[pmName] || 0) + 1;
  });
  var current = DASH_TURN_PM_FILTER;
  sel.innerHTML = '<option value="">All PMs</option>';
  Object.keys(counts).sort(function(a, b) { return a.localeCompare(b); }).forEach(function(name) {
    var opt = document.createElement('option');
    opt.value = name;
    opt.textContent = name + ' (' + counts[name] + ')';
    sel.appendChild(opt);
  });
  if (current && !counts[current]) DASH_TURN_PM_FILTER = '';
  sel.value = DASH_TURN_PM_FILTER;
}

function resetDashboardTurnRotator(totalPages) {
  if (DASH_TURN_ROTATOR) {
    clearInterval(DASH_TURN_ROTATOR);
    DASH_TURN_ROTATOR = null;
  }
  var dashSection = document.getElementById('sec-dashboard');
  if (!dashSection || !dashSection.classList.contains('active') || totalPages <= 1) return;
  DASH_TURN_ROTATOR = setInterval(function() {
    DASH_TURN_PAGE = (DASH_TURN_PAGE + 1) % totalPages;
    renderTurnDashboardStrip();
  }, DASH_TURN_ROTATE_MS);
}

function openTurnBoardDetail(turnId) {
  if (!turnId) return;
  OPEN_TURN_DETAIL_ID = turnId;
  currentTurnPipeFilter = 'all';
  if ($('#turnPipeFilter')) $('#turnPipeFilter').value = 'all';
  $$('.nav-tab').forEach(function(t) { t.classList.remove('active'); });
  var turnTab = document.querySelector('[data-tab="turnboard"]');
  if (turnTab) turnTab.classList.add('active');
  $$('.section').forEach(function(s) { s.classList.remove('active'); });
  var turnSection = document.getElementById('sec-turnboard');
  if (turnSection) turnSection.classList.add('active');
  renderTurnBoard();
  requestAnimationFrame(function() {
    var card = Array.prototype.find.call(document.querySelectorAll('#turnPipeline .pipe-card'), function(node) {
      return node.getAttribute('data-pipeid') === turnId;
    });
    if (card) card.scrollIntoView({ behavior: 'smooth', block: 'center' });
  });
}

function renderTurnDashboardStrip() {
  var strip = $('#dashTurnStrip');
  var summary = $('#dashTurnSummary');
  var pageLabel = $('#dashTurnPageLabel');
  var prevBtn = $('#dashTurnPrev');
  var nextBtn = $('#dashTurnNext');
  if (!strip || !summary || !pageLabel || !prevBtn || !nextBtn) return;

  var baseEntries = TURN_PIPE_DATA.filter(function(p) {
    return !p.isClosed && !p.isCompleted && isInPropertyGroup(p.propertyId, p.property, currentPropertyGroup);
  });
  syncDashboardTurnPmFilter(baseEntries);

  var entries = getDashboardTurnEntries();
  var pageSize = getDashboardTurnPageSize();
  var totalPages = Math.max(1, Math.ceil(entries.length / pageSize));
  if (DASH_TURN_PAGE >= totalPages) DASH_TURN_PAGE = 0;

  var activeOpenWOs = entries.reduce(function(sum, p) {
    return sum + p.matchingWOs.filter(function(wo) { return !isClosedTurnWorkOrderStatus(wo.status); }).length;
  }, 0);
  var stalledCount = entries.filter(function(p) { return p.isStalled; }).length;
  var avgElapsed = entries.length ? Math.round(entries.reduce(function(sum, p) { return sum + Math.max(0, p.elapsed); }, 0) / entries.length) : 0;
  summary.textContent = entries.length + ' active turns • ' + activeOpenWOs + ' open WOs • avg ' + avgElapsed + 'd elapsed • ' + stalledCount + ' stalled • ' + (DASH_TURN_PM_FILTER || 'All PMs');

  prevBtn.disabled = totalPages <= 1;
  nextBtn.disabled = totalPages <= 1;
  pageLabel.textContent = totalPages <= 1 ? '1 / 1' : ((DASH_TURN_PAGE + 1) + ' / ' + totalPages);

  if (entries.length === 0) {
    strip.innerHTML = '<div class="turn-dash-empty"><i class="fas fa-layer-group"></i><div>No active turns for this filter.</div></div>';
    resetDashboardTurnRotator(1);
    return;
  }

  var pageStart = DASH_TURN_PAGE * pageSize;
  var pageItems = entries.slice(pageStart, pageStart + pageSize);
  strip.innerHTML = pageItems.map(function(p) {
    var completedStages = PIPE_STAGES.filter(function(ps) { return p.stages[ps.key] && p.stages[ps.key].done; }).length;
    var progressPct = Math.max(12, Math.round((completedStages / PIPE_STAGES.length) * 100));
    var nextStage = p.currentStageIdx < PIPE_STAGES.length - 1 ? PIPE_STAGES[p.currentStageIdx + 1] : null;
    var pmName = escapeHtml(String(p.siteManager || 'Unassigned PM').trim() || 'Unassigned PM');
    var openWOs = p.matchingWOs.filter(function(wo) { return !isClosedTurnWorkOrderStatus(wo.status); }).length;
    var closedWOs = p.matchingWOs.length - openWOs;
    var statusTone = p.isStalled ? 'danger' : p.isOnRadar ? 'info' : p.isUpcoming ? 'cool' : 'ok';
    var statusLabel = p.isUpcoming ? ('Move-out ' + formatDate(p.moveOut)) : p.isOnRadar ? 'Awaiting inspection' : nextStage ? ('Next: ' + nextStage.title) : 'In progress';
    return '<button class="turn-dash-card ' + statusTone + '" data-turndash-open="' + escapeHtml(p.id) + '">' +
      '<div class="turn-dash-card-head"><div><div class="turn-dash-unit">' + escapeHtml(p.unit || 'Unit') + '</div><div class="turn-dash-property">' + escapeHtml(p.property || '') + '</div></div><span class="turn-dash-pm">' + pmName + '</span></div>' +
      '<div class="turn-dash-status">' + escapeHtml(statusLabel) + '</div>' +
      '<div class="turn-dash-progress"><div class="turn-dash-progress-fill" style="width:' + progressPct + '%"></div></div>' +
      '<div class="turn-dash-metrics">' +
        '<span class="turn-dash-metric"><strong>' + Math.max(0, p.elapsed) + 'd</strong> elapsed</span>' +
        '<span class="turn-dash-metric"><strong>' + closedWOs + '/' + p.matchingWOs.length + '</strong> closed WOs</span>' +
        '<span class="turn-dash-metric"><strong>' + completedStages + '/' + PIPE_STAGES.length + '</strong> stages</span>' +
      '</div>' +
      (openWOs > 0 ? '<div class="turn-dash-alert"><i class="fas fa-exclamation-circle"></i> ' + openWOs + ' open work order' + (openWOs === 1 ? '' : 's') + '</div>' : '<div class="turn-dash-alert ok"><i class="fas fa-check-circle"></i> No open work orders</div>') +
    '</button>';
  }).join('');

  resetDashboardTurnRotator(totalPages);
}

function getPayrollWeek(offset) {
  var now = new Date();
  var day = now.getDay(); // 0=Sun
  // Find most recent Friday
  var fridayOffset = (day + 2) % 7; // days since last Friday
  var endDate = new Date(now);
  endDate.setDate(endDate.getDate() - fridayOffset);
  endDate.setHours(23, 59, 59, 999);
  // Apply week offset
  endDate.setDate(endDate.getDate() + (offset * 7));
  var startDate = new Date(endDate);
  startDate.setDate(startDate.getDate() - 6);
  startDate.setHours(0, 0, 0, 0);
  return { start: startDate, end: endDate };
}

// Payroll uses shared currentPropertyGroup (global group filter across all tabs)

function renderPayroll() {
  var period = getPayrollWeek(PAYROLL_WEEK_OFFSET);
  var rangeEl = $('#payrollRange');
  if (rangeEl) rangeEl.textContent = formatDate(period.start) + ' \u2014 ' + formatDate(period.end);

  // Sync global group filter dropdown
  var pgSel = $('#globalGroupFilter');
  if (pgSel && pgSel.value !== currentPropertyGroup) pgSel.value = currentPropertyGroup;

  var workDone = WORK_ORDERS.filter(function(wo) {
    if (wo.status !== 'Work Done' && wo.status !== 'Ready to Bill') return false;
    var cd = wo.workCompletedOn ? new Date(wo.workCompletedOn) : (wo.completedOn ? new Date(wo.completedOn) : null);
    if (!cd) return false;
    if (cd < period.start || cd > period.end) return false;
    if (!isInPropertyGroup(wo.propertyId, wo.propertyName, currentPropertyGroup)) return false;
    return true;
  });

  var totalAmt = workDone.reduce(function(s, wo) { return s + (parseFloat(wo.amount) || 0); }, 0);
  var vendorSet = {};
  var propSet = {};
  workDone.forEach(function(wo) {
    if (wo.vendorName) vendorSet[wo.vendorName] = true;
    if (wo.propertyName) propSet[wo.propertyName] = true;
  });

  var countEl = $('#payrollCount');
  if (countEl) countEl.textContent = workDone.length;
  var countSub = $('#payrollCountSub');
  if (countSub) countSub.textContent = 'orders completed this period';
  var totalEl = $('#payrollTotal');
  if (totalEl) totalEl.textContent = currency(totalAmt);
  var totalSub = $('#payrollTotalSub');
  if (totalSub) totalSub.textContent = totalAmt > 0 ? 'total labor this period' : 'no amounts recorded';
  var vendEl = $('#payrollVendors');
  if (vendEl) vendEl.textContent = Object.keys(vendorSet).length;
  var vendSub = $('#payrollVendorsSub');
  if (vendSub) vendSub.textContent = 'unique vendors this period';
  var propEl = $('#payrollProps');
  if (propEl) propEl.textContent = Object.keys(propSet).length;
  var propSub = $('#payrollPropsSub');
  if (propSub) propSub.textContent = 'properties with work done';

  var body = $('#payrollBody');
  if (!body) return;
  if (workDone.length === 0) {
    body.innerHTML = '<tr><td colspan="8" style="text-align:center;color:var(--text-muted);padding:20px;font-size:12px">No completed work orders in this pay period</td></tr>';
    return;
  }
  var html = '';
  workDone.forEach(function(wo) {
    var flagged = isWOFlagged(wo.id);
    var afUrl = appfolioUrl('work_order', wo.id || wo.uuid);
    html += '<tr class="payroll-row" data-woid="' + escapeHtml(String(wo.id)) + '" data-wouuid="' + escapeHtml(String(wo.uuid)) + '" style="cursor:pointer;' + (flagged ? 'background:var(--warning-dim)' : '') + '">';
    html += '<td style="font-family:var(--font-mono);color:var(--accent)">#' + escapeHtml(String(wo.id)) + (afUrl ? ' <i class="fas fa-external-link-alt" style="font-size:9px;opacity:0.5"></i>' : '') + '</td>';
    html += '<td>' + escapeHtml(wo.propertyName) + '</td>';
    html += '<td>' + escapeHtml(wo.unit) + '</td>';
    html += '<td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + escapeHtml(wo.description) + '</td>';
    html += '<td>' + escapeHtml(wo.vendorName || '\u2014') + '</td>';
    html += '<td style="font-family:var(--font-mono)">' + formatDate(wo.workCompletedOn || wo.completedOn) + '</td>';
    html += '<td style="font-family:var(--font-mono)">' + (wo.amount ? currency(parseFloat(wo.amount)) : '\u2014') + '</td>';
    html += '<td><button class="flag-toggle-btn' + (flagged ? ' active' : '') + '" data-flagwo="' + escapeHtml(String(wo.id)) + '"><i class="fas fa-flag"></i></button></td>';
    html += '</tr>';
  });
  body.innerHTML = html;
  // Event listeners handled by delegation in wireUpUI() — no re-attachment needed
}

function renderDashboardKPIs() {
  // Sync global group filter dropdown
  var gfSel = $('#globalGroupFilter');
  if (gfSel && gfSel.value !== currentPropertyGroup) gfSel.value = currentPropertyGroup;
  updateGlobalGroupIndicator();

  var openWOs = WORK_ORDERS.filter(function(w) {
    if (w.status === 'Completed' || w.status === 'Canceled') return false;
    if (!isInPropertyGroup(w.propertyId, w.propertyName, currentPropertyGroup)) return false;
    return true;
  });
  var urgentWOs = WORK_ORDERS.filter(function(w) {
    if (!((w.priority === 'Urgent' || w.priority === 'Emergency') && w.status !== 'Completed' && w.status !== 'Canceled')) return false;
    if (!isInPropertyGroup(w.propertyId, w.propertyName, currentPropertyGroup)) return false;
    return true;
  });
  var activeTurns = TURNS.filter(function(t) {
    if (t.turnEnd) return false;
    if (!isInPropertyGroup(t.propertyId, t.property, currentPropertyGroup)) return false;
    return true;
  });

  var moveOuts = getUpcomingMoveOuts();
  var flaggedCount = Object.keys(WO_FLAGS).length;

  // WO aging buckets
  var agingCounts = CONFIG.WO_AGING_BUCKETS.map(function() { return 0; });
  var today = new Date();
  openWOs.forEach(function(wo) {
    if (!wo.created) return;
    var age = daysBetween(new Date(wo.created), today);
    for (var bi = 0; bi < CONFIG.WO_AGING_BUCKETS.length; bi++) {
      if (age <= CONFIG.WO_AGING_BUCKETS[bi].max) { agingCounts[bi]++; break; }
    }
  });
  var agingHtml = '<div class="aging-badges">';
  CONFIG.WO_AGING_BUCKETS.forEach(function(b, bi) {
    if (agingCounts[bi] > 0) agingHtml += '<span class="aging-badge ' + b.cls + '">' + b.label + ': ' + agingCounts[bi] + '</span>';
  });
  agingHtml += '</div>';

  // Unassigned urgent count
  var unassignedUrgent = urgentWOs.filter(function(w) { return !w.vendorName && !w.vendor; });

  $('#kpiOpen').textContent = openWOs.length;
  var openSubEl = $('#kpiOpenSub');
  if (openSubEl) openSubEl.innerHTML = WORK_ORDERS.length + ' loaded (' + DATA_WINDOW_DAYS + 'd)' + agingHtml;
  $('#kpiUrgent').textContent = urgentWOs.length;
  var urgSubText = urgentWOs.length > 0 ? urgentWOs.length + ' require attention' : 'No urgent items';
  if (unassignedUrgent.length > 0) urgSubText += ' \u2022 ' + unassignedUrgent.length + ' unassigned';
  $('#kpiUrgentSub').textContent = urgSubText;
  $('#kpiTurns').textContent = activeTurns.length;
  $('#kpiTurnsSub').textContent = TURNS.length + ' total turns';
  $('#kpiMoveOuts').textContent = moveOuts.length;
  $('#kpiMoveOutsSub').textContent = moveOuts.length > 0 ? moveOuts[0].daysLeft + 'd until next' : 'None in ' + CONFIG.MOVEOUT_WINDOW_DAYS + ' days';
  $('#kpiFlagged').textContent = flaggedCount;
  $('#kpiFlaggedSub').textContent = flaggedCount > 0 ? flaggedCount + ' items flagged' : 'No flagged items';

  $('#woBadge').textContent = openWOs.length || '0';
  $('#turnBadge').textContent = activeTurns.length || '0';

  // Active filter indicator
  var fiEl = document.getElementById('filterIndicator');
  if (fiEl) {
    if (currentPropertyGroup) {
      // Count properties in this group
      var propsInGroup = PROPERTIES.filter(function(p) {
        return isInPropertyGroup(p.id, p.name, currentPropertyGroup);
      }).length;
      fiEl.style.display = '';
      fiEl.className = 'filter-indicator';
      fiEl.innerHTML = '<i class="fas fa-filter"></i> Filtering: <strong>' + escapeHtml(currentPropertyGroup) +
        '</strong> (' + propsInGroup + ' properties) &mdash; ' + openWOs.length + ' WOs, ' + activeTurns.length + ' turns' +
        ' <button class="fi-clear" id="fiClearBtn">Clear Filter</button>';
      var clearBtn = document.getElementById('fiClearBtn');
      if (clearBtn) {
        clearBtn.onclick = function() {
          currentPropertyGroup = '';
          var ggf = document.getElementById('globalGroupFilter');
          if (ggf) ggf.value = '';
          updateGlobalGroupIndicator();
          renderAll();
        };
      }
    } else {
      fiEl.style.display = 'none';
    }
  }

  renderTurnDashboardStrip();
}

function renderActivityFeed() {
  var tbody = $('#activityBody');
  var filter = currentActivityFilter || 'all';

  // Build comprehensive activity feed from all data sources
  var activities = [];

  // === v0 API Tasks — real-time activity (highest priority) ===
  RECENT_TASKS.slice(0, 40).forEach(function(task) {
    var dateStr = task.updatedAt || task.createdAt || '';
    var isComplete = task.status === 'Completed' || task.status === 'Done';
    var taskIcon = 'fa-tasks';
    var taskColor = 'var(--accent)';
    var taskTag = '<span class="tag new">' + escapeHtml(task.status || 'Task') + '</span>';
    if (isComplete) { taskTag = '<span class="tag completed">Done</span>'; taskColor = 'var(--success)'; taskIcon = 'fa-check-circle'; }
    else if (task.status === 'In Progress' || task.status === 'Active') { taskTag = '<span class="tag assigned">In Progress</span>'; taskColor = 'var(--info)'; }
    else if (task.status === 'Overdue') { taskTag = '<span class="tag urgent">Overdue</span>'; taskColor = 'var(--danger)'; taskIcon = 'fa-exclamation-circle'; }
    var isUrgent = task.priority === 'Urgent' || task.priority === 'High' || task.status === 'Overdue';
    if (isUrgent) { taskIcon = 'fa-exclamation-circle'; taskColor = 'var(--danger)'; }

    var entityStr = escapeHtml(task.subject || task.taskType || 'Task');
    if (task.linkedResourceId) { entityStr = '#' + escapeHtml(String(task.linkedResourceId)); }
    var detailParts = [];
    if (task.propertyName) detailParts.push(escapeHtml(task.propertyName));
    if (task.unitName) detailParts.push(escapeHtml(task.unitName));
    if (task.subject && task.linkedResourceId) detailParts.push(escapeHtml(task.subject.substring(0, 80)));
    if (detailParts.length === 0 && task.body) detailParts.push(escapeHtml(task.body.substring(0, 80)));

    activities.push({
      sortDate: new Date(dateStr || 0).getTime(),
      time: dateStr ? timeAgo(dateStr) : '\u2014',
      type: 'work_order',
      urgent: isUrgent,
      icon: taskIcon,
      iconColor: taskColor,
      event: taskTag,
      entity: entityStr,
      detail: detailParts.join(' / '),
      extra: (task.assignee ? '<span style="color:var(--purple)"><i class="fas fa-user" style="font-size:9px"></i> ' + escapeHtml(task.assignee) + '</span>' : '') +
             (task.dueDate ? ' <span style="color:var(--text-muted)"><i class="fas fa-calendar" style="font-size:9px"></i> Due: ' + formatDate(task.dueDate) + '</span>' : '')
    });
  });

  // === Webhook events — decoded from AppFolio webhook relay ===
  WEBHOOK_EVENTS.slice(0, 20).forEach(function(wh) {
    var isPri = wh.priority === 'urgent' || wh.priority === 'high';
    // Pick icon based on resource type
    var whIcon = 'fa-plug'; var whColor = 'var(--purple)';
    if (wh.resource_type === 'work_order') { whIcon = 'fa-wrench'; whColor = 'var(--accent)'; }
    else if (wh.resource_type === 'unit_turn') { whIcon = 'fa-exchange-alt'; whColor = 'var(--info,#60a5fa)'; }
    else if (wh.resource_type === 'inspection') { whIcon = 'fa-clipboard-check'; whColor = 'var(--success)'; }
    else if (wh.resource_type === 'tenant') { whIcon = 'fa-user'; whColor = 'var(--warning)'; }
    else if (wh.resource_type === 'vendor') { whIcon = 'fa-hard-hat'; whColor = 'var(--success)'; }
    var eventLabel = WEBHOOK_LABELS[wh.type] || wh.event_label || wh.type || 'webhook';
    activities.push({
      sortDate: new Date(wh.ts || 0).getTime(),
      time: wh.ts ? timeAgo(wh.ts) : '\u2014',
      type: 'work_order',
      urgent: isPri,
      icon: isPri ? 'fa-exclamation-circle' : whIcon,
      iconColor: isPri ? 'var(--danger)' : whColor,
      event: '<span class="tag ' + (isPri ? 'urgent' : 'new') + '">' + escapeHtml(eventLabel) + '</span>',
      entity: escapeHtml(wh.resource_name || wh.source || 'appfolio'),
      detail: escapeHtml(wh.title || ''),
      extra: '<span style="color:var(--purple)"><i class="fas fa-satellite-dish" style="font-size:9px"></i> live</span>'
    });
  });

  // === Work Order events — fill in if tasks are sparse ===
  var sortedWOs = WORK_ORDERS.slice().sort(function(a, b) {
    return new Date(b.created || 0) - new Date(a.created || 0);
  });
  sortedWOs.slice(0, 20).forEach(function(wo) {
    var isUrgent = wo.priority === 'Urgent' || wo.priority === 'Emergency';
    activities.push({
      sortDate: new Date(wo.created || 0).getTime(),
      time: timeAgo(wo.created),
      type: 'work_order',
      urgent: isUrgent,
      icon: isUrgent ? 'fa-exclamation-circle' : 'fa-wrench',
      iconColor: isUrgent ? 'var(--danger)' : 'var(--accent)',
      event: '<span class="tag ' + String(wo.status).toLowerCase().replace(/\s+/g, '-') + '">' + escapeHtml(wo.status) + '</span>',
      entity: '#' + String(wo.id),
      detail: escapeHtml(wo.propertyName || '') + (wo.unit ? ' / ' + escapeHtml(wo.unit) : '') + ' \u2014 ' + escapeHtml((wo.description || '').substring(0, 80)),
      extra: (wo.vendorName ? '<span style="color:var(--info)"><i class="fas fa-hard-hat" style="font-size:9px"></i> ' + escapeHtml(wo.vendorName) + '</span>' : '') +
             (wo.assignedUser ? ' <span style="color:var(--purple)"><i class="fas fa-user" style="font-size:9px"></i> ' + escapeHtml(wo.assignedUser) + '</span>' : '') +
             (wo.tenant ? ' <span style="color:var(--text-muted)"><i class="fas fa-user-friends" style="font-size:9px"></i> ' + escapeHtml(wo.tenant) + '</span>' : '')
    });
  });

  // Turn events
  TURNS.slice(0, 10).forEach(function(t) {
    var isActive = !t.turnEnd;
    activities.push({
      sortDate: new Date(t.moveOut || 0).getTime(),
      time: timeAgo(t.moveOut || new Date().toISOString()),
      type: 'turn',
      urgent: false,
      icon: isActive ? 'fa-exchange-alt' : 'fa-check-circle',
      iconColor: isActive ? 'var(--warning)' : 'var(--success)',
      event: isActive ? '<span class="tag waiting">In Progress</span>' : '<span class="tag completed">Completed</span>',
      entity: escapeHtml(t.unit),
      detail: escapeHtml(t.property) + ' \u2014 ' + (isActive ? (t.moveOut ? daysBetween(t.moveOut, new Date()) + 'd since move-out' : 'Active turn') : t.totalDays + ' days total'),
      extra: t.totalBilled && t.totalBilled !== '$0.00' ? '<span style="color:var(--danger)">Billed: ' + escapeHtml(t.totalBilled) + '</span>' : ''
    });
  });

  // Inspection events — overdue only
  var today = new Date();
  INSPECTIONS.filter(function(r) {
    var lastDate = r.lastInspection ? new Date(r.lastInspection) : null;
    return !lastDate || daysBetween(lastDate, today) > 365;
  }).slice(0, 8).forEach(function(r) {
    activities.push({
      sortDate: r.lastInspection ? new Date(r.lastInspection).getTime() : 0,
      time: r.lastInspection ? timeAgo(r.lastInspection) : 'Never',
      type: 'inspection',
      urgent: true,
      icon: 'fa-clipboard-check',
      iconColor: 'var(--danger)',
      event: '<span class="tag non-compliant">Overdue</span>',
      entity: escapeHtml(r.unit || ''),
      detail: escapeHtml(r.propertyName) + (r.tenant ? ' \u2014 ' + escapeHtml(r.tenant) : ''),
      extra: ''
    });
  });

  // Sort by date descending
  activities.sort(function(a, b) { return b.sortDate - a.sortDate; });

  // Apply filter
  if (filter !== 'all') {
    if (filter === 'urgent') {
      activities = activities.filter(function(a) { return a.urgent; });
    } else {
      activities = activities.filter(function(a) { return a.type === filter; });
    }
  }

  activities = activities.slice(0, 30);

  if (activities.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4">' + emptyHtml('fa-inbox', 'No activity to show') + '</td></tr>';
    return;
  }

  var html = '';
  activities.forEach(function(a) {
    html += '<tr>';
    html += '<td style="font-family:var(--font-mono);font-size:11px;color:var(--text-muted);white-space:nowrap;"><i class="fas ' + a.icon + '" style="color:' + a.iconColor + ';margin-right:4px;font-size:10px"></i>' + a.time + '</td>';
    html += '<td>' + a.event + '</td>';
    html += '<td style="font-family:var(--font-mono);font-size:12px;color:var(--accent)">' + a.entity + '</td>';
    html += '<td style="font-size:11px"><div style="color:var(--text-secondary)">' + a.detail + '</div>';
    if (a.extra) { html += '<div style="margin-top:2px;font-size:10px">' + a.extra + '</div>'; }
    html += '</td></tr>';
  });
  tbody.innerHTML = html;
}

var currentWOFilter = 'all';
var currentWOPriority = '';
var currentWOType = '';
var currentWOProperty = '';
var currentActivityFilter = 'all';

// Kanban status columns and their display labels
var KANBAN_STATUSES = [
  { key: 'New', label: 'New' },
  { key: 'Estimate Requested', label: 'Est. Req.' },
  { key: 'Estimated', label: 'Estimated' },
  { key: 'Assigned', label: 'Assigned' },
  { key: 'Scheduled', label: 'Scheduled' },
  { key: 'Waiting', label: 'Waiting' },
  { key: 'Work Done', label: 'Work Done' },
  { key: 'Ready to Bill', label: 'Ready to Bill' }
];

function getFilteredWOs() {
  var search = $('#woSearch') ? $('#woSearch').value : '';
  return WORK_ORDERS.filter(function(wo) {
    // Status filter (from filter buttons or kanban column click)
    if (currentWOFilter && currentWOFilter !== 'all' && wo.status !== currentWOFilter) return false;
    // Priority dropdown
    if (currentWOPriority && wo.priority !== currentWOPriority) return false;
    // Type dropdown
    if (currentWOType && wo.type !== currentWOType) return false;
    // Property dropdown
    if (currentWOProperty && wo.propertyName !== currentWOProperty) return false;
    // Property group filter — shared helper
    if (!isInPropertyGroup(wo.propertyId, wo.propertyName, currentPropertyGroup)) return false;
    // Flagged filter
    if (currentWOFilter === 'flagged' && !isWOFlagged(wo.id)) return false;
    // Search
    if (search) {
      var s = search.toLowerCase();
      var haystack = [String(wo.id), String(wo.description || ''), String(wo.propertyName || ''), String(wo.vendorName || ''), String(wo.unit || ''), String(wo.tenant || ''), String(wo.assignedUser || '')].join(' ').toLowerCase();
      return haystack.indexOf(s) !== -1;
    }
    return true;
  });
}

function renderWorkOrders() {
  var board = $('#kanbanBoard');
  if (!board) return;
  var filtered = getFilteredWOs();

  // Populate property dropdown with unique properties from current WOs
  var propSel = $('#woPropertyFilter');
  if (propSel && propSel.options.length <= 1) {
    var props = {};
    WORK_ORDERS.forEach(function(wo) { if (wo.propertyName) props[wo.propertyName] = true; });
    Object.keys(props).sort().forEach(function(p) {
      var opt = document.createElement('option');
      opt.value = p; opt.textContent = p;
      propSel.appendChild(opt);
    });
  }

  // Sync global group filter dropdown
  var grpSel = $('#globalGroupFilter');
  if (grpSel && grpSel.value !== currentPropertyGroup) grpSel.value = currentPropertyGroup;
  updateGlobalGroupIndicator();

  // Vendor compliance map for cross-tab warnings
  var vendorCompliance = buildVendorComplianceMap();

  if (WORK_ORDERS.length === 0) {
    board.innerHTML = '<div style="padding:40px;text-align:center;color:var(--text-muted);width:100%"><i class="fas fa-inbox" style="font-size:36px;display:block;margin-bottom:12px;color:var(--border)"></i>No work orders loaded. Connect to API or import a cache file.</div>';
    return;
  }

  // Group WOs by status
  var groups = {};
  filtered.forEach(function(wo) {
    var st = wo.status || 'New';
    if (!groups[st]) groups[st] = [];
    groups[st].push(wo);
  });

  var html = '';
  KANBAN_STATUSES.forEach(function(col) {
    var wos = groups[col.key] || [];
    var selected = currentWOFilter === col.key ? ' selected' : '';
    html += '<div class="kanban-col">';
    html += '<div class="kanban-col-head' + selected + '" data-status="' + escapeHtml(col.key) + '">';
    html += '<span class="kanban-col-title">' + escapeHtml(col.label) + '</span>';
    html += '<span class="kanban-col-count">' + wos.length + '</span>';
    html += '</div>';
    html += '<div class="kanban-col-body">';
    if (wos.length === 0) {
      html += '<div style="padding:12px;text-align:center;color:var(--text-muted);font-size:10px;font-family:var(--font-mono)">empty</div>';
    }
    wos.forEach(function(wo) {
      var pc = String(wo.priority || 'normal').toLowerCase();
      var flagged = isWOFlagged(wo.id);
      html += '<div class="kanban-card ' + pc + (flagged ? ' flagged-card' : '') + '" data-woid="' + escapeHtml(String(wo.id)) + '">';
      html += '<div class="kc-top"><span class="kc-id">#' + escapeHtml(String(wo.id)) + (flagged ? ' <i class="fas fa-flag kc-flag"></i>' : '') + '</span><span class="kc-priority"><span class="tag ' + pc + '">' + escapeHtml(wo.priority) + '</span></span></div>';
      html += '<div class="kc-desc">' + escapeHtml(wo.description || 'No description') + '</div>';
      html += '<div class="kc-meta">';
      if (wo.propertyName) html += '<span><i class="fas fa-building"></i> ' + escapeHtml(wo.propertyName) + '</span>';
      if (wo.unit) html += '<span><i class="fas fa-door-open"></i> ' + escapeHtml(wo.unit) + '</span>';
      if (wo.vendorName) {
        html += '<span><i class="fas fa-hard-hat"></i> ' + escapeHtml(wo.vendorName) + '</span>';
        var vKey = (wo.vendorName || '').toLowerCase();
        if (vendorCompliance[vKey]) {
          html += '<div class="wo-vendor-warn"><i class="fas fa-exclamation-triangle"></i> Insurance ' + vendorCompliance[vKey] + '</div>';
        }
      }
      if (wo.created) html += '<span><i class="fas fa-clock"></i> ' + timeAgo(wo.created) + '</span>';
      if (wo.tenant) html += '<span><i class="fas fa-user"></i> ' + escapeHtml(wo.tenant) + '</span>';
      html += '</div></div>';
    });
    html += '</div></div>';
  });

  // Also show any WOs with statuses not in KANBAN_STATUSES
  var otherWos = filtered.filter(function(wo) { return !KANBAN_STATUSES.some(function(s) { return s.key === wo.status; }); });
  if (otherWos.length > 0) {
    html += '<div class="kanban-col"><div class="kanban-col-head"><span class="kanban-col-title">Other</span><span class="kanban-col-count">' + otherWos.length + '</span></div><div class="kanban-col-body">';
    otherWos.forEach(function(wo) {
      var pc = String(wo.priority || 'normal').toLowerCase();
      html += '<div class="kanban-card ' + pc + '" data-woid="' + escapeHtml(String(wo.id)) + '"><div class="kc-top"><span class="kc-id">#' + escapeHtml(String(wo.id)) + '</span><span class="kc-priority"><span class="tag ' + pc + '">' + escapeHtml(wo.priority) + '</span></span></div>';
      html += '<div class="kc-desc">' + escapeHtml(wo.description || 'No description') + '</div>';
      html += '<div class="kc-meta"><span>' + escapeHtml(wo.status) + '</span>';
      if (wo.propertyName) html += '<span><i class="fas fa-building"></i> ' + escapeHtml(wo.propertyName) + '</span>';
      html += '</div></div>';
    });
    html += '</div></div>';
  }

  board.innerHTML = html;
  $('#woBadge').textContent = filtered.length || '0';
  renderWOCloseAssist();
  renderWOFollowupQueue();
  // Event listeners handled by delegation in wireUpUI() — no re-attachment needed
}

function normalizeLooseKey(v) {
  return String(v || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function getWOCreatedDate(wo) {
  var candidates = [wo.created, wo.scheduledStart, wo.updated];
  for (var i = 0; i < candidates.length; i++) {
    var d = new Date(candidates[i]);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function woAgeBucket(days) {
  if (days >= 180) return '180d+';
  if (days >= 60) return '60d+';
  if (days >= 30) return '30d+';
  if (days >= 14) return '14d+';
  return '<14d';
}

function findBillMatchesForWO(wo, createdDate) {
  if (!BILLS || BILLS.length === 0) return [];
  var woVendorId = normalizeLooseKey(wo.vendorId);
  var woVendorName = normalizeLooseKey(wo.vendorName);
  var woPropId = normalizeLooseKey(wo.propertyId);
  var woPropName = normalizeLooseKey(wo.propertyName);
  var start = createdDate ? new Date(createdDate.getTime() - (2 * 86400000)) : null;
  var end = new Date();

  return BILLS.filter(function(b) {
    var bVendorId = normalizeLooseKey(b.vendorId);
    var bVendorName = normalizeLooseKey(b.vendorName);
    var bPropId = normalizeLooseKey(b.propertyId);
    var bPropName = normalizeLooseKey(b.propertyName);

    var vendorMatch = (woVendorId && bVendorId && woVendorId === bVendorId) ||
      (woVendorName && bVendorName && woVendorName === bVendorName);
    if (!vendorMatch) return false;

    var propMatch = (woPropId && bPropId && woPropId === bPropId) ||
      (woPropName && bPropName && woPropName === bPropName);
    if (!propMatch) return false;

    if (!start) return true;
    var billDate = new Date(b.date || '');
    if (isNaN(billDate.getTime())) return true;
    return billDate >= start && billDate <= end;
  });
}

function computeWOCloseAssistRows() {
  var closedStatuses = { 'completed': true, 'canceled': true, 'work completed': true };
  var rows = [];

  WORK_ORDERS.forEach(function(wo) {
    var status = normalizeLooseKey(wo.status);
    if (closedStatuses[status]) return;
    if (!isInPropertyGroup(wo.propertyId, wo.propertyName, currentPropertyGroup)) return;

    var createdDate = getWOCreatedDate(wo);
    if (!createdDate) return;
    var ageDays = daysBetween(createdDate, new Date());
    if (ageDays < currentWOCloseAssistAge) return;

    var hasVendor = !!(normalizeLooseKey(wo.vendorId) || normalizeLooseKey(wo.vendorName));
    var hasAssignee = !!normalizeLooseKey(wo.assignedUser);
    var billMatches = hasVendor ? findBillMatchesForWO(wo, createdDate) : [];

    var confidence = 'Low';
    var suggestion = 'Review and update current status';
    if (hasVendor && billMatches.length > 0) {
      confidence = 'High';
      suggestion = 'Likely completed — verify and close';
    } else if (hasVendor && ageDays >= 60) {
      confidence = 'Medium';
      suggestion = 'Vendor assigned but no AP match — follow up with vendor';
    } else if (hasAssignee && ageDays >= 30) {
      confidence = 'Medium';
      suggestion = 'Assigned internally — request completion update';
    } else if (!hasVendor && !hasAssignee) {
      confidence = 'Low';
      suggestion = 'Unassigned — triage owner and next step required';
    }

    rows.push({
      wo: wo,
      ageDays: ageDays,
      bucket: woAgeBucket(ageDays),
      billMatches: billMatches,
      confidence: confidence,
      suggestion: suggestion
    });
  });

  rows.sort(function(a, b) {
    var confRank = { High: 0, Medium: 1, Low: 2 };
    if (confRank[a.confidence] !== confRank[b.confidence]) return confRank[a.confidence] - confRank[b.confidence];
    return b.ageDays - a.ageDays;
  });
  return rows;
}

function renderWOCloseAssist() {
  var body = $('#woCloseAssistBody');
  var summary = $('#woCloseAssistSummary');
  if (!body || !summary) return;

  var rows = computeWOCloseAssistRows();
  var hi = rows.filter(function(r) { return r.confidence === 'High'; }).length;
  var med = rows.filter(function(r) { return r.confidence === 'Medium'; }).length;
  var apLabel = BILLS.length > 0 ? 'AP bills: ' + BILLS.length : 'AP bills: none loaded — click Refresh AP for evidence';
  summary.textContent = rows.length + ' candidate(s) • High: ' + hi + ' • Medium: ' + med + ' • ' + apLabel;

  if (rows.length === 0) {
    var emptyMsg;
    if (WORK_ORDERS.length === 0) {
      emptyMsg = 'Work orders not loaded — refresh to load data.';
    } else if (BILLS.length === 0) {
      emptyMsg = 'No aged open WOs match current filter. ℹ️ AP bills not loaded — click ‘Refresh AP’ above for vendor-match evidence.';
    } else {
      emptyMsg = 'No candidates for current age/filter window.';
    }
    body.innerHTML = '<tr><td colspan="8" style="text-align:center;color:var(--text-muted)">' + emptyMsg + '</td></tr>';
    return;
  }

  var html = '';
  rows.slice(0, 80).forEach(function(r) {
    var wo = r.wo;
    var confColor = r.confidence === 'High' ? 'var(--danger)' : (r.confidence === 'Medium' ? 'var(--warning)' : 'var(--text-muted)');
    var billEvidence = r.billMatches.length > 0 ? (r.billMatches.length + ' AP match') : '—';
    html += '<tr>';
    html += '<td><button class="action-btn" data-woid="' + escapeHtml(String(wo.id)) + '" style="padding:2px 8px">#' + escapeHtml(String(wo.id)) + '</button></td>';
    html += '<td>' + escapeHtml(wo.propertyName || '—') + '</td>';
    html += '<td>' + escapeHtml(wo.unit || '—') + '</td>';
    html += '<td>' + r.ageDays + 'd <span style="color:var(--text-muted)">(' + r.bucket + ')</span></td>';
    html += '<td>' + escapeHtml(wo.assignedUser || wo.vendorName || '—') + '</td>';
    html += '<td>' + escapeHtml(billEvidence) + '</td>';
    html += '<td><span style="font-weight:700;color:' + confColor + '">' + r.confidence + '</span></td>';
    html += '<td>' + escapeHtml(r.suggestion) + '</td>';
    html += '</tr>';
  });
  body.innerHTML = html;

  Array.prototype.forEach.call(body.querySelectorAll('button[data-woid]'), function(btn) {
    btn.addEventListener('click', function() {
      showWODetail(btn.getAttribute('data-woid'));
    });
  });
}

function renderWOFollowupQueue() {
  var body = $('#woFollowupBody');
  var summary = $('#woFollowupSummary');
  if (!body || !summary) return;

  var rows = WORK_ORDERS.filter(function(wo) {
    return isWOFlagged(wo.id) && isInPropertyGroup(wo.propertyId, wo.propertyName, currentPropertyGroup);
  }).map(function(wo) {
    var flag = WO_FLAGS[wo.id] || {};
    return { wo: wo, flag: flag };
  });

  rows.sort(function(a, b) {
    return (b.flag.ts || 0) - (a.flag.ts || 0);
  });

  summary.textContent = rows.length > 0
    ? rows.length + ' flagged WO(s) awaiting follow-up'
    : 'No WOs flagged for follow-up — use the flag button on any WO to add it here';

  if (rows.length === 0) {
    body.innerHTML = '<tr><td colspan="8" style="text-align:center;color:var(--text-muted)">No WOs flagged for follow-up — use the flag button on any WO to add it here</td></tr>';
    return;
  }

  var html = '';
  rows.forEach(function(r) {
    var wo = r.wo;
    var note = (r.flag.note || '');
    var flaggedAt = r.flag.ts ? new Date(r.flag.ts).toLocaleDateString() : '—';
    html += '<tr>';
    html += '<td><button class="action-btn" data-woid="' + escapeHtml(String(wo.id)) + '" style="padding:2px 8px">#' + escapeHtml(String(wo.id)) + '</button></td>';
    html += '<td>' + escapeHtml(wo.propertyName || '—') + '</td>';
    html += '<td>' + escapeHtml(wo.unit || '—') + '</td>';
    html += '<td>' + escapeHtml((wo.description || '').substring(0, 90) || '—') + '</td>';
    html += '<td>' + escapeHtml(wo.status || '—') + '</td>';
    html += '<td>' + escapeHtml(flaggedAt) + '</td>';
    html += '<td><input class="form-input" data-follow-note="' + escapeHtml(String(wo.id)) + '" value="' + escapeHtml(note) + '" placeholder="Add follow-up note" style="min-width:180px"></td>';
    html += '<td style="white-space:nowrap">';
    html += '<button class="action-btn" data-saveflag="' + escapeHtml(String(wo.id)) + '"><i class="fas fa-save"></i></button> ';
    html += '<button class="action-btn" data-clearflag="' + escapeHtml(String(wo.id)) + '"><i class="fas fa-flag-checkered"></i></button>';
    html += '</td>';
    html += '</tr>';
  });
  body.innerHTML = html;

  Array.prototype.forEach.call(body.querySelectorAll('button[data-woid]'), function(btn) {
    btn.addEventListener('click', function() {
      showWODetail(btn.getAttribute('data-woid'));
    });
  });

  Array.prototype.forEach.call(body.querySelectorAll('button[data-saveflag]'), function(btn) {
    btn.addEventListener('click', async function() {
      var woId = btn.getAttribute('data-saveflag');
      var input = body.querySelector('input[data-follow-note="' + woId + '"]');
      var noteVal = input ? input.value.trim() : '';
      await saveFlag(woId, noteVal);
      renderWorkOrders();
      renderDashboardKPIs();
      showToast('Follow-up note saved for WO #' + woId);
    });
  });

  Array.prototype.forEach.call(body.querySelectorAll('button[data-clearflag]'), function(btn) {
    btn.addEventListener('click', async function() {
      var woId = btn.getAttribute('data-clearflag');
      await removeFlag(woId);
      renderWorkOrders();
      renderDashboardKPIs();
      showToast('Follow-up cleared for WO #' + woId);
    });
  });
}

function showWODetail(id) {
  var wo = WORK_ORDERS.find(function(w) { return String(w.id) === String(id); });
  if (!wo) return;
  var woDbUuid = resolveWODbUuid(wo);
  CURRENT_WO_MODAL = { woId: String(wo.id), woDbUuid: woDbUuid || '' };
  var woAfUrl = appfolioUrl('work_order', wo.id || wo.uuid);
  $('#woModalTitle').innerHTML = '#' + escapeHtml(String(wo.id)) + ' \u2014 ' + escapeHtml(wo.propertyName) + ' ' + escapeHtml(wo.unit) + (woAfUrl ? ' <a href="' + escapeHtml(woAfUrl) + '" target="_blank" rel="noopener noreferrer" style="font-size:12px;color:var(--accent);margin-left:8px;text-decoration:none" title="View in AppFolio"><i class="fas fa-external-link-alt"></i></a>' : '');

  // Flag button state
  var flagBtn = $('#woModalFlag');
  if (flagBtn) {
    flagBtn.className = 'flag-toggle-btn' + (isWOFlagged(wo.id) ? ' active' : '');
    flagBtn.onclick = async function() {
      await toggleFlag(wo.id);
      flagBtn.className = 'flag-toggle-btn' + (isWOFlagged(wo.id) ? ' active' : '');
      renderWorkOrders();
      renderDashboardKPIs();
    };
  }

  var STATUSES = ['New','Estimate Requested','Estimated','Assigned','Scheduled','Waiting','Work Done','Ready to Bill','Work Completed','Completed','Canceled'];
  var PRIORITIES = ['Urgent','Normal','Low'];

  var html = '';
  // -- WO Info Section --
  html += '<div class="detail-section"><div class="detail-section-title"><i class="fas fa-info-circle"></i> Work Order Info</div>';
  html += '<div class="detail-grid">';
  html += '<div class="detail-row"><div class="detail-row-label">Status</div><select class="form-select" id="detailStatus">';
  STATUSES.forEach(function(s) { html += '<option' + (s === wo.status ? ' selected' : '') + '>' + s + '</option>'; });
  html += '</select></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Priority</div><select class="form-select" id="detailPriority">';
  PRIORITIES.forEach(function(p) { html += '<option' + (p === wo.priority ? ' selected' : '') + '>' + p + '</option>'; });
  html += '</select></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Type</div><div class="detail-row-value">' + escapeHtml(wo.type || '\u2014') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Amount</div><div class="detail-row-value">' + escapeHtml(wo.amount || '\u2014') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Created</div><div class="detail-row-value">' + formatDate(wo.created) + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Scheduled</div><div class="detail-row-value">' + formatDate(wo.scheduledStart) + '</div></div>';
  html += '</div>';
  html += '<div class="form-group" style="margin-top:10px"><label class="form-label">Description</label><textarea class="form-textarea" id="detailDesc">' + escapeHtml(wo.description) + '</textarea></div>';
  html += '</div>';

  // -- Property & Unit Section --
  html += '<div class="detail-section"><div class="detail-section-title"><i class="fas fa-building"></i> Property &amp; Unit</div>';
  html += '<div class="detail-grid">';
  html += '<div class="detail-row"><div class="detail-row-label">Property</div><div class="detail-row-value">' + escapeHtml(wo.propertyName) + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Unit</div><div class="detail-row-value">' + escapeHtml(wo.unit || '\u2014') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Address</div><div class="detail-row-value">' + escapeHtml(wo.propertyAddress || '\u2014') + '</div></div>';
  // Site manager placeholder — will be filled async
  html += '<div class="detail-row"><div class="detail-row-label">Site Manager</div><div class="detail-row-value" id="detailSiteMgr"><i class="fas fa-spinner fa-spin" style="font-size:10px"></i></div></div>';
  html += '</div></div>';

  // -- Tenant Section --
  html += '<div class="detail-section"><div class="detail-section-title"><i class="fas fa-user"></i> Tenant</div>';
  html += '<div class="detail-grid">';
  html += '<div class="detail-row"><div class="detail-row-label">Name</div><div class="detail-row-value">' + escapeHtml(wo.tenant || '\u2014') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Email</div><div class="detail-row-value">' + escapeHtml(wo.tenantEmail || '\u2014') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Phone</div><div class="detail-row-value">' + escapeHtml(wo.tenantPhone || '\u2014') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Assigned To</div><div class="detail-row-value">' + escapeHtml(wo.assignedUser || '\u2014') + '</div></div>';
  html += '</div></div>';

  // -- Vendor Section --
  html += '<div class="detail-section"><div class="detail-section-title"><i class="fas fa-hard-hat"></i> Vendor</div>';
  html += '<div class="detail-grid">';
  html += '<div class="detail-row"><div class="detail-row-label">Vendor</div><div class="detail-row-value">' + escapeHtml(wo.vendorName || 'Unassigned') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Trade</div><div class="detail-row-value">' + escapeHtml(wo.vendorTrade || '\u2014') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Created By</div><div class="detail-row-value">' + escapeHtml(wo.createdBy || '\u2014') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Maint. Limit</div><div class="detail-row-value">' + escapeHtml(wo.maintenanceLimit || '\u2014') + '</div></div>';
  html += '</div></div>';

  // -- Notes Section (async load) --
  html += '<div class="detail-section"><div class="detail-section-title"><i class="fas fa-sticky-note"></i> Notes</div>';
  html += '<div class="note-list" id="detailNotesList"><div style="text-align:center;padding:10px;color:var(--text-muted)"><i class="fas fa-spinner fa-spin"></i> Loading notes\u2026</div></div></div>';

  // -- Add Note --
  html += '<div class="detail-section"><div class="detail-section-title"><i class="fas fa-plus-circle"></i> Add Note</div>';
  html += '<textarea class="form-textarea" placeholder="Type a note\u2026" id="detailNote"></textarea></div>';

  $('#woModalBody').innerHTML = html;

  // Async: fetch property detail for site manager
  if (wo.propertyId) {
    var prop = PROPERTIES.find(function(p) { return p.id === wo.propertyId || String(p.id) === String(wo.propertyId); });
    if (prop && prop.siteManager) {
      var smEl = document.getElementById('detailSiteMgr');
      if (smEl) smEl.textContent = prop.siteManager;
    } else {
      fetchPropertyDetail(wo.propertyId).then(function(data) {
        var smEl = document.getElementById('detailSiteMgr');
        if (smEl) smEl.textContent = (data && (data.site_manager || data.SiteManager)) || '\u2014';
      });
    }
  } else {
    var smEl = document.getElementById('detailSiteMgr');
    if (smEl) smEl.textContent = '\u2014';
  }

  // Async: fetch notes (use resolved DB API UUID for /api/v0/ endpoint)
  fetchWONotes(woDbUuid).then(function(notes) {
    renderWONotesList(notes);
  });

  $('#woModalSave').onclick = async function() {
    var newStatus = $('#detailStatus').value;
    var newPriority = $('#detailPriority').value;
    var note = ($('#detailNote') && $('#detailNote').value) ? $('#detailNote').value.trim() : '';

    try {
      if (newStatus !== wo.status || newPriority !== wo.priority) {
        if (!woDbUuid) { showToast('Cannot update: no DB API UUID for this WO'); return; }
        await apiFetch('/api/v0/work_orders/' + woDbUuid, {
          method: 'PATCH',
          body: JSON.stringify({ Status: newStatus, Priority: newPriority })
        });
        wo.status = newStatus;
        wo.priority = newPriority;
      }
      if (note) {
        if (!woDbUuid) { showToast('Cannot add note: no DB API UUID for this WO'); return; }
        try {
          await apiFetch('/api/v0/work_orders/' + woDbUuid + '/notes', {
            method: 'POST',
            body: JSON.stringify({ Body: note })
          });
          refreshCurrentWONotes(woDbUuid);
        } catch (noteErr) {
          showToast('Note failed: ' + noteErr.message);
        }
      }
      renderWorkOrders();
      renderDashboardKPIs();
      closeModal('woModal');
      showToast('Updated #' + wo.id + ' successfully');
      await saveAllToCache();
    } catch (err) {
      showToast('Update failed: ' + err.message);
    }
  };
  openModal('woModal');
}

/* =================================================================
   TURN PIPELINE — Stage Tracking Engine
   Correlates: Turns (v2) + Work Orders + Inspections + Webhook Events
   Stages: MO → INS → WO → REQ → EST → ASN → DONE
   ================================================================= */
var TURN_RECORDS = []; // persisted stage overrides from proxy blob
var TURN_PIPE_DATA = []; // computed pipeline entries
var OPEN_TURN_DETAIL_ID = '';
var DASH_TURN_PM_FILTER = '';
var DASH_TURN_PAGE = 0;
var DASH_TURN_ROTATOR = null;
var DASH_TURN_ROTATE_MS = 8000;
var currentTurnPipeFilter = 'active';
var currentTurnPipeGroup = '';
var _inspSortCol = 'daysSince'; // default sort column
var _inspSortDir = 'desc';      // 'asc' or 'desc'

// Stage definitions — Hybrid Turn Pipeline phases
// Upcoming = pre-turn (tenant gave notice), then MO → INS → WO → REQ → EST → ASN → DONE
var PIPE_STAGES = [
  { key: 'upcoming',    label: 'UPC',  icon: 'fa-calendar-alt', title: 'Upcoming' },
  { key: 'moveout',     label: 'MO',   icon: 'fa-sign-out-alt', title: 'Move-Out' },
  { key: 'inspection',  label: 'INS',  icon: 'fa-clipboard-check', title: 'Inspection' },
  { key: 'wo_created',  label: 'WO',   icon: 'fa-wrench', title: 'WO Created' },
  { key: 'est_requested', label: 'REQ', icon: 'fa-file-invoice', title: 'Bidding' },
  { key: 'est_received', label: 'EST', icon: 'fa-file-invoice-dollar', title: 'Estimated' },
  { key: 'assigned',    label: 'ASN',  icon: 'fa-user-check', title: 'Approved' },
  { key: 'work_done',   label: 'DONE', icon: 'fa-check-circle', title: 'Work Done' }
];

// Fetch persisted turn records from proxy
async function fetchTurnRecords() {
  try {
    var data = await proxyAction('turn_records');
    if (data && Array.isArray(data.records)) TURN_RECORDS = data.records;
  } catch (err) {
    console.log('Turn records fetch error: ' + (err.message || err));
  }
}

// Save a turn record stage to proxy
async function saveTurnRecordStage(turnId, stage, stageData) {
  try {
    var sep = API_PROXY.indexOf('?') !== -1 ? '&' : '?';
    var url = API_PROXY + sep + 'action=turn_record_stage';
    await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id: turnId, stage: stage, data: stageData })
    });
  } catch (err) {
    console.log('Save stage error: ' + (err.message || err));
  }
}

// Save full turn record to proxy
async function saveTurnRecord(record) {
  try {
    var sep = API_PROXY.indexOf('?') !== -1 ? '&' : '?';
    var url = API_PROXY + sep + 'action=turn_records';
    await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(record)
    });
  } catch (err) {
    console.log('Save turn record error: ' + (err.message || err));
  }
}

// ---- Auto-correlation engine ----
// Builds a unified pipeline entry for each turn by matching WOs + inspections
// ---- Hybrid Turn Pipeline Builder ----
// Combines: unit_turn_detail report (TURNS) + upcoming moveouts (UPCOMING_MOVEOUTS)
//           + work orders (WORK_ORDERS + TURN_WORK_ORDERS) + inspections + webhook events
//           + persisted state from TURN_RECORDS blob
function buildTurnPipeline() {
  TURN_PIPE_DATA = [];
  var today = new Date();
  var seenKeys = {};

  function isTurnWorkOrderCandidate(woLike, moveOutDate, source) {
    var type = String((woLike && (woLike.type || woLike.Type)) || '').toLowerCase();
    var desc = String((woLike && (woLike.description || woLike.Description || woLike.JobDescription)) || '').toLowerCase();
    var looksTurnType = type.indexOf('unit_turn') !== -1 || /(^|\b)turn(\b|$)/.test(type);
    var looksTurnText = /(unit\s*turn|make[-\s]?ready|ready[-\s]?to[-\s]?rent|move[-\s]?out|vacan|turnover|rehab|renov)/.test(desc);

    // turn_work_orders is already narrowed server-side; keep unless obviously out-of-window
    if (source === 'db') {
      if (moveOutDate && woLike.createdAt) {
        var dbDate = new Date(woLike.createdAt);
        if (!isNaN(dbDate.getTime())) {
          var dbDiff = (dbDate - moveOutDate) / 86400000;
          if (dbDiff < -45 || dbDiff > 300) return false;
        }
      }
      return true;
    }

    if (!(looksTurnType || looksTurnText)) return false;
    if (moveOutDate && woLike.created) {
      var woDate = new Date(woLike.created);
      if (!isNaN(woDate.getTime())) {
        var diff = (woDate - moveOutDate) / 86400000;
        if (diff < -45 || diff > 300) return false;
      }
    }
    return true;
  }

  // ---- Pre-build lookup indexes for O(1) matching (avoids O(n*m) iteration) ----
  // Index WORK_ORDERS by composite key: "unit_lower|prop_lower"
  var _woByUnitProp = {};
  WORK_ORDERS.forEach(function(wo) {
    if (!wo.unit || !wo.propertyName) return;
    var key = String(wo.unit).toLowerCase() + '|' + String(wo.propertyName).toLowerCase();
    if (!_woByUnitProp[key]) _woByUnitProp[key] = [];
    _woByUnitProp[key].push(wo);
  });
  // Index TURN_WORK_ORDERS by unitId + propertyId
  var _turnWoByUnit = {};
  var _turnWoByProp = {};
  var _turnWoByNum = {};
  TURN_WORK_ORDERS.forEach(function(wo) {
    if (wo.unitId) { if (!_turnWoByUnit[wo.unitId]) _turnWoByUnit[wo.unitId] = []; _turnWoByUnit[wo.unitId].push(wo); }
    if (wo.propertyId) { if (!_turnWoByProp[wo.propertyId]) _turnWoByProp[wo.propertyId] = []; _turnWoByProp[wo.propertyId].push(wo); }
    if (wo.woNumber) _turnWoByNum[wo.woNumber] = wo;
  });

  // Helper: create a composite key for deduplication
  function makeKey(propId, unitId, moveOut) {
    var k = String(propId || '') + '-' + String(unitId || '') + '-' + (moveOut || '');
    return k === '--' ? null : k;
  }

  // Helper: find matching WOs for a unit+property pair using pre-built indexes
  function findMatchingWOs(unit, property, propId, unitId, moveOutDate) {
    var wos = [];
    // From Reports API work orders — O(1) lookup by composite key
    var lookupKey = (unit && property) ? String(unit).toLowerCase() + '|' + String(property).toLowerCase() : null;
    var reportsWOs = lookupKey ? (_woByUnitProp[lookupKey] || []) : [];
    reportsWOs.forEach(function(wo) {
      if (!isTurnWorkOrderCandidate(wo, moveOutDate, 'reports')) return;
      if (moveOutDate && wo.created) {
        var daysDiff = (new Date(wo.created) - moveOutDate) / 86400000;
        if (daysDiff < -30) return;
      }
      wos.push({ source: 'reports', id: wo.id, status: wo.status, description: wo.description || '',
        created: wo.created, vendor: wo.vendor || '', unit: wo.unit, property: wo.propertyName, priority: wo.priority });
    });
    // From DB API turn work orders — O(1) lookup by unit/property ID
    var dbCandidates = [];
    if (unitId && _turnWoByUnit[unitId]) dbCandidates = dbCandidates.concat(_turnWoByUnit[unitId]);
    if (propId && _turnWoByProp[propId]) {
      _turnWoByProp[propId].forEach(function(tw) {
        if (!dbCandidates.some(function(c) { return c.id === tw.id; })) dbCandidates.push(tw);
      });
    }
    // Also check WO number matches
    wos.forEach(function(w) {
      if (w.id && _turnWoByNum[String(w.id)]) {
        var tw = _turnWoByNum[String(w.id)];
        if (!dbCandidates.some(function(c) { return c.id === tw.id; })) dbCandidates.push(tw);
      }
    });
    dbCandidates.forEach(function(wo) {
      if (!isTurnWorkOrderCandidate(wo, moveOutDate, 'db')) return;
      var dupe = wos.find(function(w) { return String(w.id) === String(wo.id) || String(w.id) === String(wo.woNumber); });
      if (dupe) {
        dupe.status = wo.status;
        dupe.dbApiId = wo.id;
        return;
      }
      wos.push({ source: 'db_api', id: wo.woNumber || wo.id, dbApiId: wo.id, status: wo.status,
        description: wo.description || '', created: wo.createdAt, vendor: wo.vendorTrade || '', priority: wo.priority });
    });
    return wos;
  }

  // Helper: find matching inspection
  function findMatchingInsp(unit, property) {
    return INSPECTIONS.find(function(insp) {
      return insp.unit && unit && String(insp.unit).toLowerCase() === String(unit).toLowerCase() &&
             insp.propertyName && property && String(insp.propertyName).toLowerCase() === String(property).toLowerCase();
    });
  }

  // Helper: derive stages from available data
  function deriveStages(moveOut, matchingWOs, matchingInsp, isUpcoming) {
    var stages = {};
    var moveOutDate = moveOut ? new Date(moveOut) : null;

    // Stage 0: Upcoming — tenant gave notice, move-out in the future
    stages.upcoming = { done: true, date: moveOut || null };

    // Stage 1: Move-Out — has the tenant actually moved out?
    var movedOut = moveOutDate ? moveOutDate <= today : false;
    stages.moveout = { done: movedOut, date: movedOut ? moveOut : null };

    // Stage 2: Inspection — check if inspection happened after move-out
    var inspDone = false;
    var inspDate = null;
    if (matchingInsp && matchingInsp.lastInspection) {
      inspDate = matchingInsp.lastInspection;
      if (moveOutDate) {
        inspDone = new Date(inspDate) >= moveOutDate;
      } else {
        inspDone = !!inspDate;
      }
    }
    stages.inspection = { done: inspDone, date: inspDate };

    // Stage 3-7: Derive from WO statuses
    var hasWO = matchingWOs.length > 0;
    var woStatuses = matchingWOs.map(function(w) { return w.status; });
    var woCreatedDate = hasWO ? matchingWOs[0].created : null;

    var hasEstReq = woStatuses.some(function(s) { return s === 'Estimate Requested'; });
    var hasEstimated = woStatuses.some(function(s) { return s === 'Estimated'; });
    var hasAssigned = woStatuses.some(function(s) { return s === 'Assigned' || s === 'Scheduled'; });
    var hasAnyWorkDone = woStatuses.some(function(s) { return isTurnWorkDoneStatus(s); });
    // ALL WOs must be in terminal status for work_done to be truly "done"
    var allWorkDone = hasWO && woStatuses.every(function(s) { return isClosedTurnWorkOrderStatus(s); });
    var doneCount = woStatuses.filter(function(s) { return isClosedTurnWorkOrderStatus(s); }).length;

    // Progressive — later stages imply earlier ones
    stages.wo_created = { done: hasWO, date: woCreatedDate, woIds: matchingWOs.map(function(w) { return w.id; }) };
    stages.est_requested = { done: hasEstReq || hasEstimated || hasAssigned || hasAnyWorkDone, date: null };
    stages.est_received = { done: hasEstimated || hasAssigned || hasAnyWorkDone, date: null, vendors: [] };
    stages.assigned = { done: hasAssigned || hasAnyWorkDone, date: null };
    // work_done requires ALL WOs complete, not just one
    stages.work_done = { done: allWorkDone, date: null, doneCount: doneCount, totalCount: matchingWOs.length };

    return stages;
  }

  // Helper: build a pipeline entry
  function addEntry(key, unit, property, propId, unitId, moveOut, turnData, moveoutTenant) {
    if (!key || seenKeys[key]) return;
    seenKeys[key] = true;

    // Exclude units with a confirmed move-in (turn has reset/closed).
    var actualMoveIn = (turnData && (turnData.moveIn || turnData.actual_move_in || turnData.actualMoveIn)) || '';
    if (!actualMoveIn && UNIT_TURNS_DB.length > 0) {
      var dbMatch = UNIT_TURNS_DB.find(function(u) {
        return (unitId && String(u.unitId) === String(unitId)) ||
               (unit && u.unit && String(u.unit).toLowerCase() === String(unit).toLowerCase() &&
                property && u.property && String(u.property).toLowerCase() === String(property).toLowerCase());
      });
      if (dbMatch) actualMoveIn = dbMatch.moveIn || dbMatch.actual_move_in || dbMatch.actualMoveIn || '';
    }
    if (actualMoveIn) {
      var miDate = new Date(actualMoveIn);
      if (!isNaN(miDate.getTime()) && miDate <= today) return;
    }

    var moveOutDate = moveOut ? new Date(moveOut) : null;
    var isUpcoming = moveOutDate ? moveOutDate > today : false;
    var turnYearStart = getCurrentYearStartDate(today);
    if (moveOutDate && !isNaN(moveOutDate.getTime()) && moveOutDate < turnYearStart) return;
    var matchingWOs = findMatchingWOs(unit, property, propId, unitId, moveOutDate);
    var matchingInsp = findMatchingInsp(unit, property);
    var stages = deriveStages(moveOut, matchingWOs, matchingInsp, isUpcoming);

    // Drop stale, unconfirmed radar rows with no turn signals.
    if (!turnData && moveOutDate && !isUpcoming) {
      var ageDays = daysBetween(moveOutDate, today);
      var hasSignal = (matchingWOs.length > 0) || (stages.inspection && stages.inspection.done);
      if (ageDays > 180 && !hasSignal) return;
      if (ageDays > 365) return;
    }
    var propMeta = PROPERTIES.find(function(p) {
      return String(p.id) === String(propId) ||
             ((p.name || '').toLowerCase() === (property || '').toLowerCase());
    }) || null;
    var relatedTurns = TURNS.filter(function(t) {
      if (!turnData) return false;
      var sameTurn = String(t.unitTurnId || '') === String(turnData.unitTurnId || '');
      var sameUnit = String(t.unitId || '') === String(unitId || '') ||
                     (((t.unit || '').toLowerCase() === (unit || '').toLowerCase()) &&
                      ((t.property || '').toLowerCase() === (property || '').toLowerCase()));
      return !sameTurn && sameUnit;
    });

    // If this is from unit_turn_detail and has a turnEnd, mark work_done
    if (turnData && turnData.turnEnd) {
      stages.work_done.done = true;
      stages.work_done.date = turnData.turnEnd;
    }

    // Merge persisted overrides from TURN_RECORDS
    var savedRec = TURN_RECORDS.find(function(r) { return r.id === key; });
    if (savedRec && savedRec.stages) {
      PIPE_STAGES.forEach(function(ps) {
        var saved = savedRec.stages[ps.key];
        if (saved) {
          if (!stages[ps.key]) stages[ps.key] = {};
          if (saved.done) stages[ps.key].done = true;
          if (saved.date && !stages[ps.key].date) stages[ps.key].date = saved.date;
          if (saved.notes) stages[ps.key].notes = saved.notes;
          if (saved.vendors) stages[ps.key].vendors = saved.vendors;
        }
      });
    }

    // Find webhook events matching this turn
    var webhookMatches = WEBHOOK_EVENTS.filter(function(wh) {
      var t = (wh.title || '').toLowerCase();
      var b = (wh.body || '').toLowerCase();
      var uLow = (unit || '').toLowerCase();
      var pLow = (property || '').toLowerCase();
      return uLow && pLow && (t.indexOf(uLow) !== -1 || b.indexOf(uLow) !== -1) && (t.indexOf(pLow) !== -1 || b.indexOf(pLow) !== -1);
    });

    // Elapsed days since move-out (or days until move-out for upcoming)
    var elapsed = 0;
    if (moveOutDate) {
      if (isUpcoming) {
        elapsed = -daysBetween(today, moveOutDate); // negative = days until
      } else {
        elapsed = daysBetween(moveOutDate, today);
      }
    }
    var target = (turnData && turnData.targetDays) || CONFIG.TURN_TARGET_DAYS;
    var isStalled = !isUpcoming && elapsed > CONFIG.TURN_STALLED_DAYS && currentStageIdx >= 1 && currentStageIdx < PIPE_STAGES.length - 1;
    // Aggregate completion: ALL linked WOs must be in terminal status
    var allWOsDone = isTurnFullyComplete(matchingWOs);
    // A turn is CONFIRMED when AppFolio is tracking it (comes from unit_turn_detail report)
    // OR when a post-move-out inspection exists for that unit (the confirmation gate).
    var isConfirmed = (turnData !== null) || stages.inspection.done;
    var isOnRadar = !isConfirmed; // watching but not yet a confirmed active turn
    // Completed requires confirmation — unconfirmed upcoming never counts as done
    var isCompleted = isConfirmed && (
      (turnData && !!turnData.turnEnd) ||
      (allWOsDone && matchingWOs.length > 0) ||
      (matchingWOs.length === 0 && stages.work_done && stages.work_done.done));
    // Safety: if WOs exist but not all done, ensure not marked complete
    if (matchingWOs.length > 0 && !allWOsDone) {
      stages.work_done.done = false;
      isCompleted = false;
    }
    // Compute current stage index after work-order completion safety checks.
    var currentStageIdx = -1;
    PIPE_STAGES.forEach(function(ps, i) {
      if (stages[ps.key] && stages[ps.key].done) currentStageIdx = i;
    });
    // Look up DB API unit turn data for this entry
    var dbTurnMatch = null;
    if (UNIT_TURNS_DB.length > 0) {
      dbTurnMatch = UNIT_TURNS_DB.find(function(u) {
        return (unitId && String(u.unitId) === String(unitId) && propId && String(u.propertyId) === String(propId)) ||
               (unit && property && u.unit && u.property &&
                String(u.unit).toLowerCase() === String(unit).toLowerCase() &&
                String(u.property).toLowerCase() === String(property).toLowerCase());
      }) || null;
    }

    // Deposit deadline countdown (override with DB API date if available)
    var depositMoveOut = (dbTurnMatch && dbTurnMatch.moveOut) || moveOut;
    var sla = (!isUpcoming && depositMoveOut) ? calculateDepositDeadline(depositMoveOut) : null;

    // Parse cost
    var costNum = 0;
    var totalBilledStr = (dbTurnMatch && dbTurnMatch.totalBilled) || (turnData && turnData.totalBilled) || '$0';
    costNum = parseFloat(String(totalBilledStr).replace(/[^0-9.\-]/g, '')) || 0;

    TURN_PIPE_DATA.push({
      id: key,
      turn: turnData || null,
      unit: unit,
      property: property,
      propertyId: propId,
      unitId: unitId,
      moveOut: moveOut,
      tenant: moveoutTenant || '',
      isUpcoming: isUpcoming,
      isConfirmed: isConfirmed,
      isOnRadar: isOnRadar,
      stages: stages,
      currentStageIdx: currentStageIdx,
      matchingWOs: matchingWOs,
      matchingInsp: matchingInsp,
      webhookEvents: webhookMatches,
      elapsed: elapsed,
      target: target,
      isStalled: isStalled && !isCompleted,
      isCompleted: isCompleted,
      allWOsDone: allWOsDone,
      siteManager: (turnData && turnData.siteManager) || (propMeta && propMeta.siteManager) || '',
      maintenanceLimit: (turnData && turnData.maintenanceLimit) || (propMeta && propMeta.maintenanceLimit) || '',
      propertyNotes: (turnData && turnData.propertyNotes) || (propMeta && (propMeta.maintenanceNotes || propMeta.propertyNotes)) || '',
      unitTurnStatus: (dbTurnMatch && dbTurnMatch.status) || (turnData && turnData.status) || '',
      isRegisteredUnitTurn: !!(turnData && turnData.isRegisteredUnitTurn),
      registeredUnitTurnId: (turnData && turnData.registeredUnitTurnId) || '',
      relatedTurns: relatedTurns,
      sla: sla,
      costNum: costNum,
      totalBilled: totalBilledStr,
      depositStatus: (dbTurnMatch && dbTurnMatch.depositStatus) || '',
      depositAmount: (dbTurnMatch && dbTurnMatch.depositAmount) || '',
      depositReturnDeadline: (dbTurnMatch && dbTurnMatch.depositReturnDeadline) || '',
      expectedMoveIn: (dbTurnMatch && dbTurnMatch.expectedMoveIn) || (turnData && turnData.expectedMoveIn) || '',
      dbLink: (turnData && turnData.dbLink) || (dbTurnMatch && dbTurnMatch.link) || '',
      savedRecord: savedRec || null
    });
  }

  // PASS 1: Add all turns from unit_turn_detail report
  TURNS.forEach(function(turn) {
    var key = makeKey(turn.propertyId, turn.unitId, turn.moveOut) ||
              turn.unitTurnId || (turn.unit + '|' + turn.property);
    addEntry(key, turn.unit, turn.property, turn.propertyId, turn.unitId, turn.moveOut, turn, '');
  });

  // PASS 2: Add upcoming move-outs not already in pipeline (the "Upcoming" phase)
  UPCOMING_MOVEOUTS.forEach(function(mo) {
    var key = makeKey(mo.propertyId, mo.unitId, mo.moveOut);
    if (!key || seenKeys[key]) return;
    addEntry(key, mo.unit, mo.property, mo.propertyId, mo.unitId, mo.moveOut, null, mo.tenant);
  });

  // Sort: confirmed active (most elapsed) → on-radar (upcoming first, then oldest-awaiting)
  // → completed (most recently done first)
  TURN_PIPE_DATA.sort(function(a, b) {
    var aGroup = a.isCompleted ? 2 : a.isConfirmed ? 0 : 1; // 0=active, 1=radar, 2=done
    var bGroup = b.isCompleted ? 2 : b.isConfirmed ? 0 : 1;
    if (aGroup !== bGroup) return aGroup - bGroup;
    if (aGroup === 0) return b.elapsed - a.elapsed; // active: most elapsed first
    if (aGroup === 1) { // on-radar: upcoming (soonest MO first) then past-awaiting-insp
      if (a.isUpcoming !== b.isUpcoming) return a.isUpcoming ? -1 : 1;
      if (a.isUpcoming && b.isUpcoming) return a.elapsed - b.elapsed; // days-until asc
      return b.elapsed - a.elapsed; // days-since-moveout desc
    }
    return b.elapsed - a.elapsed; // completed: most recent first
  });
}

function renderTurnBoard() {
  try {
    buildTurnPipeline();
  } catch (e) {
    console.log('buildTurnPipeline error: ' + (e.message || e));
    TURN_PIPE_DATA = [];
  }
  try {
    renderTurnPipelineUI();
  } catch (e) {
    console.log('renderTurnPipelineUI error: ' + (e.message || e));
  }
  try {
    renderTurnKPIs();
  } catch (e) {
    console.log('renderTurnKPIs error: ' + (e.message || e));
  }
}

function renderTurnKPIs() {
  // Separate confirmed-active turns from on-radar (unconfirmed) entries
  var confirmed = TURN_PIPE_DATA.filter(function(p) { return p.isConfirmed && !p.isCompleted; });
  var onRadar   = TURN_PIPE_DATA.filter(function(p) { return p.isOnRadar; });
  var upcoming  = TURN_PIPE_DATA.filter(function(p) { return p.isUpcoming; }); // subset of onRadar
  var active    = confirmed; // alias — same set for downstream stats
  var awaitEst  = confirmed.filter(function(p) {
    return p.stages.wo_created && p.stages.wo_created.done && p.stages.est_received && !p.stages.est_received.done;
  });
  var totalBilled = 0;
  confirmed.forEach(function(p) { totalBilled += p.costNum; });
  var avgDays = 0;
  if (confirmed.length > 0) {
    var totalDays = 0;
    confirmed.forEach(function(p) { totalDays += Math.abs(p.elapsed); });
    avgDays = Math.round(totalDays / confirmed.length);
  }

  var e = function(id, v) { var el = document.getElementById(id); if (el) el.textContent = v; };
  e('kpiActiveTurns', confirmed.length);
  e('kpiActiveTurnsSub', confirmed.length + ' confirmed' + (upcoming.length > 0 ? ', ' + upcoming.length + ' upcoming' : ''));
  e('kpiOnRadar', onRadar.length);
  var radarUpcoming = onRadar.filter(function(p) { return p.isUpcoming; }).length;
  var radarPast     = onRadar.filter(function(p) { return !p.isUpcoming; }).length;
  e('kpiOnRadarSub', (radarUpcoming > 0 ? radarUpcoming + ' upcoming' : '') + (radarUpcoming > 0 && radarPast > 0 ? ', ' : '') + (radarPast > 0 ? radarPast + ' awaiting inspection' : '') || 'no possible turns');
  e('kpiAvgTurnDays', avgDays > 0 ? avgDays + 'd' : '\u2014');
  e('kpiAvgTurnSub', avgDays > 0 ? 'avg days elapsed' : 'no active turns');
  e('kpiAwaitEst', awaitEst.length);
  e('kpiAwaitEstSub', awaitEst.length + ' turns pending vendor bids');
  e('kpiTurnBilled', currency(totalBilled));
  e('kpiTurnBilledSub', 'active turns combined');

  var tb = $('#turnBadge');
  if (tb) tb.textContent = active.length + upcoming.length;

  // Populate property group dropdown
  var groupSel = $('#turnPipeGroup');
  if (groupSel && groupSel.options.length <= 1) {
    var props = {};
    TURN_PIPE_DATA.forEach(function(p) { if (p.property) props[p.property] = true; });
    Object.keys(props).sort().forEach(function(name) {
      var opt = document.createElement('option');
      opt.value = name;
      opt.textContent = name;
      groupSel.appendChild(opt);
    });
  }
}

function renderTurnPipelineUI() {
  var container = $('#turnPipeline');
  if (!container) return;

  if (TURNS.length === 0 && UPCOMING_MOVEOUTS.length === 0) {
    container.innerHTML = emptyHtml('fa-exchange-alt', 'No turn or move-out data loaded yet. Data may still be syncing — check progress above, or try Refresh.');
    return;
  }

  var filter = currentTurnPipeFilter;
  var group = currentTurnPipeGroup;
  var search = ($('#turnPipeSearch') ? $('#turnPipeSearch').value : '').toLowerCase();

  // Sync global group filter dropdown
  var turnGfSel = $('#globalGroupFilter');
  if (turnGfSel && turnGfSel.value !== currentPropertyGroup) turnGfSel.value = currentPropertyGroup;

  var filtered = TURN_PIPE_DATA.filter(function(p) {
    // 'active' = confirmed turns that aren't yet complete
    if (filter === 'active' && (!p.isConfirmed || p.isCompleted)) return false;
    // 'on_radar' = unconfirmed (possible) turns
    if (filter === 'on_radar' && !p.isOnRadar) return false;
    if (filter === 'completed' && !p.isCompleted) return false;
    if (filter === 'stalled' && (!p.isStalled || p.isCompleted)) return false;
    if (filter === 'upcoming' && !p.isUpcoming) return false;
    if (group && p.property !== group) return false;
    // Property group filter (global)
    if (!isInPropertyGroup(p.propertyId, p.property, currentPropertyGroup)) return false;
    if (search) {
      var hay = (p.unit + ' ' + p.property + ' ' + (p.tenant || '')).toLowerCase();
      if (hay.indexOf(search) === -1) return false;
    }
    return true;
  });

  if (filtered.length === 0) {
    container.innerHTML = '<div style="padding:20px;text-align:center;color:var(--text-muted);font-size:13px">' +
      '<i class="fas fa-filter" style="font-size:20px;display:block;margin-bottom:8px"></i>No turns match current filter</div>';
    return;
  }

  var html = '';
  // Track distinct groups to decide whether to render section dividers
  var groupCounts = { active: 0, on_radar: 0, completed: 0 };
  filtered.forEach(function(p) {
    var g = p.isCompleted ? 'completed' : p.isConfirmed ? 'active' : 'on_radar';
    groupCounts[g]++;
  });
  var multiGroup = (groupCounts.active > 0 ? 1 : 0) + (groupCounts.on_radar > 0 ? 1 : 0) + (groupCounts.completed > 0 ? 1 : 0) > 1;
  var prevGroup = null;
  filtered.forEach(function(p, idx) {
    if (multiGroup) {
      var thisGroup = p.isCompleted ? 'completed' : p.isConfirmed ? 'active' : 'on_radar';
      if (thisGroup !== prevGroup) {
        var secLabels = {
          active:    '<i class="fas fa-exchange-alt" style="margin-right:6px"></i>Active Turns \u2014 ' + groupCounts.active,
          on_radar:  '<i class="fas fa-satellite-dish" style="margin-right:6px"></i>On Radar \u2014 ' + groupCounts.on_radar + ' possible',
          completed: '<i class="fas fa-check-circle" style="margin-right:6px"></i>Completed \u2014 ' + groupCounts.completed
        };
        var secCls = thisGroup === 'active' ? 'active-section' : thisGroup === 'on_radar' ? 'on-radar-section' : 'completed-section';
        html += '<div class="pipe-section-head ' + secCls + '">' + secLabels[thisGroup] + '</div>';
        prevGroup = thisGroup;
      }
    }
    var slaBreach = p.sla && p.sla.breached && !p.isCompleted;
    var cardClass = p.isCompleted ? '' : p.isUpcoming ? 'upcoming' : p.isOnRadar ? 'on-radar' : slaBreach ? 'sla-breach' : p.isStalled ? 'stalled' : p.elapsed < CONFIG.TURN_WARNING_DAYS ? 'on-track' : 'waiting';
    html += '<div class="pipe-card ' + cardClass + '" data-pipeidx="' + idx + '" data-pipeid="' + escapeHtml(p.id) + '">';

    // Left: unit info
    html += '<div class="pipe-card-unit"><div class="pipe-card-unit-name">' + escapeHtml(p.unit || 'Unit') + '</div>';
    html += '<div class="pipe-card-prop">' + escapeHtml(p.property) + '</div>';
    if (p.tenant) html += '<div class="pipe-card-prop" style="font-size:10px;color:var(--accent)">' + escapeHtml(p.tenant) + '</div>';
    html += '</div>';

    // Center: stage dots
    html += '<div class="pipe-card-stages">';
    PIPE_STAGES.forEach(function(ps, si) {
      var stage = p.stages[ps.key] || {};
      var dotClass = '';
      var isGate     = p.isOnRadar && ps.key === 'inspection' && !stage.done;
      var isPostGate = p.isOnRadar && si > 2 && !(p.stages.inspection && p.stages.inspection.done);
      if (stage.done) {
        dotClass = 'done';
      } else if (isGate) {
        dotClass = 'gate';
      } else if (!p.isOnRadar && si === p.currentStageIdx + 1) {
        dotClass = p.isStalled ? 'warn' : 'active';
      } else if (isPostGate) {
        dotClass = 'waiting-gate';
      }
      html += '<div class="pipe-dot ' + dotClass + '" title="' + escapeHtml(ps.title) + (stage.date ? ' — ' + formatDate(stage.date) : '') + '">';
      html += '<i class="fas ' + ps.icon + '"></i></div>';
    });
    html += '</div>';

    // Right: status
    html += '<div class="pipe-card-status">';
    if (p.isCompleted) {
      var completeDays = (p.turn && p.turn.totalDays) || p.elapsed;
      html += '<span class="pipe-card-elapsed" style="color:var(--success)">' + completeDays + 'd</span>';
      html += '<span class="pipe-card-cost">' + escapeHtml(p.totalBilled) + ' &bull; Complete</span>';
    } else if (p.isUpcoming) {
      var daysUntil = Math.abs(p.elapsed);
      html += '<span class="pipe-card-elapsed" style="color:var(--info,#60a5fa)">' + daysUntil + 'd</span>';
      html += '<span class="pipe-card-cost">Move-out: ' + formatDate(p.moveOut) + '</span>';
    } else if (p.isOnRadar) {
      html += '<span class="radar-badge" style="margin-bottom:3px">ON RADAR</span>';
      html += '<span class="pipe-card-elapsed" style="color:#60a5fa">' + p.elapsed + 'd</span>';
      html += '<span class="pipe-card-cost" style="color:var(--text-muted)">Awaiting inspection</span>';
    } else {
      var eColor = p.elapsed > p.target ? 'var(--danger)' : p.elapsed > CONFIG.TURN_WARNING_DAYS ? 'var(--warning)' : 'var(--text-primary)';
      html += '<span class="pipe-card-elapsed" style="color:' + eColor + '">' + p.elapsed + 'd</span>';
      var nextStage = p.currentStageIdx < PIPE_STAGES.length - 1 ? PIPE_STAGES[p.currentStageIdx + 1] : null;
      html += '<span class="pipe-card-cost">' + escapeHtml(p.totalBilled);
      if (nextStage) html += ' &bull; Next: ' + nextStage.label;
      // WO progress: show X/Y done when there are multiple WOs
      if (p.stages.work_done && p.stages.work_done.totalCount > 0) {
        html += ' <span style="color:var(--text-muted);font-size:10px">(' + p.stages.work_done.doneCount + '/' + p.stages.work_done.totalCount + ' WOs)</span>';
      }
      html += '</span>';
      // Deposit deadline progress bar
      if (p.sla) {
        var slaColor = p.sla.businessDaysLeft <= 2 ? 'red' : p.sla.businessDaysLeft <= 6 ? 'yellow' : 'green';
        var slaLabel = p.sla.overdue ? 'Past deposit deadline' : p.sla.businessDaysLeft + ' biz days remaining';
        html += '<div class="sla-bar" title="Deposit deadline — ' + slaLabel + '"><div class="sla-bar-fill ' + slaColor + '" style="width:' + p.sla.pct + '%"></div></div>';
      }
    }
    html += '</div>';

    html += '</div>'; // end pipe-card

    // Detail panel (hidden by default)
    html += '<div class="pipe-detail" id="pipeDetail_' + idx + '" data-pipeid="' + escapeHtml(p.id) + '">';
    html += '<div class="pipe-detail-grid">';

    // Left column: stage timeline
    html += '<div>';
    html += '<div class="detail-section-title"><i class="fas fa-stream"></i> Stage Timeline</div>';
    html += '<ul class="pipe-timeline">';
    PIPE_STAGES.forEach(function(ps) {
      var stage = p.stages[ps.key] || {};
      var dotCls = stage.done ? 'done' : 'pending';
      html += '<li><div class="pipe-tl-dot ' + dotCls + '"><i class="fas ' + (stage.done ? 'fa-check' : 'fa-circle') + '"></i></div>';
      html += '<div><div class="pipe-tl-label">' + escapeHtml(ps.title) + '</div>';
      if (stage.date) html += '<div class="pipe-tl-date">' + formatDate(stage.date) + '</div>';
      if (stage.notes) html += '<div class="pipe-tl-note">' + escapeHtml(stage.notes) + '</div>';
      if (ps.key === 'wo_created' && stage.woIds && stage.woIds.length > 0) {
        html += '<div class="pipe-tl-note">WOs: ' + stage.woIds.map(function(id) { return '#' + id; }).join(', ') + '</div>';
      }
      if (ps.key === 'est_received' && stage.vendors && stage.vendors.length > 0) {
        html += '<div class="pipe-tl-note">Vendors: ' + stage.vendors.map(function(v) { return escapeHtml(v); }).join(', ') + '</div>';
      }
      html += '</div></li>';
    });
    html += '</ul></div>';

    // Right column: associated data
    html += '<div>';

    // Matched Work Orders
    var woSecTitle = p.isOnRadar
      ? 'Potential Work Orders (' + p.matchingWOs.length + ') \u2014 unconfirmed'
      : 'Linked Work Orders (' + p.matchingWOs.length + ')';
    html += '<div class="detail-section-title"><i class="fas fa-wrench"></i> ' + woSecTitle + '</div>';
    if (p.isOnRadar) {
      html += '<div style="font-size:11px;color:#60a5fa;padding:5px 8px;background:rgba(96,165,250,0.09);border-radius:4px;margin-bottom:6px"><i class="fas fa-info-circle" style="margin-right:4px"></i>Once a move-out inspection is recorded, this turn is confirmed and these WOs will be formally linked.</div>';
    }
    if (p.matchingWOs.length > 0) {
      html += '<div class="pipe-wo-list">';
      p.matchingWOs.forEach(function(wo) {
        // Prefer DB API Link field (direct AppFolio URL), fall back to constructed URL
        var dbWo = wo.dbApiId ? TURN_WORK_ORDERS.find(function(tw) { return tw.id === wo.dbApiId; }) : null;
        var woLink = (dbWo && dbWo.link) ? dbWo.link : appfolioUrl('work_order', wo.id);
        html += '<div class="pipe-wo-item"><div><span class="pipe-wo-id">#' + wo.id + '</span> <span class="tag ' + String(wo.status).toLowerCase().replace(/\s+/g, '-') + '">' + escapeHtml(wo.status) + '</span>';
        if (woLink) html += ' <a href="' + escapeHtml(woLink) + '" target="_blank" rel="noopener noreferrer" style="font-size:9px;color:var(--accent);text-decoration:none" title="View WO in AppFolio" onclick="event.stopPropagation()"><i class="fas fa-external-link-alt"></i></a>';
        html += '</div>';
        html += '<div style="font-size:11px;color:var(--text-secondary)">' + escapeHtml((wo.description || '').substring(0, 60)) + '</div></div>';
      });
      html += '</div>';
    } else {
      html += '<div style="font-size:12px;color:var(--text-muted);padding:8px 0">No linked work orders found</div>';
    }

    // Turn details
    html += '<div class="detail-section-title" style="margin-top:12px"><i class="fas fa-info-circle"></i> Turn Details</div>';
    html += '<div class="detail-grid">';
    html += '<div class="detail-row"><div class="detail-row-label">Move-Out</div><div class="detail-row-value">' + (p.moveOut ? formatDate(p.moveOut) : '\u2014') + '</div></div>';
    html += '<div class="detail-row"><div class="detail-row-label">Expected Move-In</div><div class="detail-row-value">' + (p.expectedMoveIn ? formatDate(p.expectedMoveIn) : '\u2014') + '</div></div>';
    if (p.unitTurnStatus) html += '<div class="detail-row"><div class="detail-row-label">Turn Status</div><div class="detail-row-value">' + escapeHtml(p.unitTurnStatus) + '</div></div>';
    if (p.depositStatus) html += '<div class="detail-row"><div class="detail-row-label">Deposit Status</div><div class="detail-row-value">' + escapeHtml(p.depositStatus) + '</div></div>';
    if (p.depositReturnDeadline) html += '<div class="detail-row"><div class="detail-row-label">Deposit Return Date</div><div class="detail-row-value">' + formatDate(p.depositReturnDeadline) + '</div></div>';
    html += '<div class="detail-row"><div class="detail-row-label">Site Manager</div><div class="detail-row-value">' + escapeHtml(p.siteManager || '\u2014') + '</div></div>';
    html += '<div class="detail-row"><div class="detail-row-label">Maintenance Limit</div><div class="detail-row-value">' + escapeHtml(p.maintenanceLimit || '\u2014') + '</div></div>';
    html += '<div class="detail-row"><div class="detail-row-label">Property Notes</div><div class="detail-row-value">' + escapeHtml(p.propertyNotes || '\u2014') + '</div></div>';
    if (p.isOnRadar) {
      html += '<div class="detail-row"><div class="detail-row-label">Confirmation</div><div class="detail-row-value" style="color:#60a5fa"><i class="fas fa-satellite-dish" style="margin-right:4px"></i>On Radar \u2014 awaiting inspection</div></div>';
    }
    if (p.tenant) html += '<div class="detail-row"><div class="detail-row-label">Tenant</div><div class="detail-row-value">' + escapeHtml(p.tenant) + '</div></div>';
    if (p.turn) {
      html += '<div class="detail-row"><div class="detail-row-label">Target Days</div><div class="detail-row-value">' + (p.target || '\u2014') + '</div></div>';
      html += '<div class="detail-row"><div class="detail-row-label">Registered Unit Turn</div><div class="detail-row-value">' + (p.isRegisteredUnitTurn ? 'Yes' : 'No') + '</div></div>';
      if (p.registeredUnitTurnId) {
        html += '<div class="detail-row"><div class="detail-row-label">Registered Turn ID</div><div class="detail-row-value">' + escapeHtml(p.registeredUnitTurnId) + '</div></div>';
      }
      // Deposit deadline
      if (p.sla) {
        var slaStyle = p.sla.overdue ? 'color:var(--danger);font-weight:700' : p.sla.businessDaysLeft <= 3 ? 'color:var(--warning);font-weight:700' : '';
        var slaText = p.sla.overdue ? 'Past deadline (' + Math.abs(p.sla.calendarDaysLeft) + 'd overdue)' : p.sla.businessDaysLeft + ' biz days remaining (due ' + formatDate(p.sla.deadline.toISOString()) + ')';
        html += '<div class="detail-row"><div class="detail-row-label">Deposit Deadline (21 biz days)</div><div class="detail-row-value" style="' + slaStyle + '">' + slaText + '</div></div>';
      }
      // WO completion aggregate
      if (p.matchingWOs.length > 0) {
        var doneCount = p.matchingWOs.filter(function(w) { return isClosedTurnWorkOrderStatus(w.status); }).length;
        var woCompStyle = doneCount === p.matchingWOs.length ? 'color:var(--success)' : 'color:var(--warning)';
        html += '<div class="detail-row"><div class="detail-row-label">WO Completion</div><div class="detail-row-value" style="' + woCompStyle + '">' + doneCount + '/' + p.matchingWOs.length + ' complete</div></div>';
      }
      html += '<div class="detail-row"><div class="detail-row-label">Total Billed</div><div class="detail-row-value">' + escapeHtml(p.totalBilled) + '</div></div>';
      html += '<div class="detail-row"><div class="detail-row-label">Labor</div><div class="detail-row-value">' + escapeHtml(p.turn.laborCost || '$0') + '</div></div>';
      html += '<div class="detail-row"><div class="detail-row-label">Reference</div><div class="detail-row-value">' + escapeHtml(p.turn.referenceUser || '\u2014') + '</div></div>';
    } else {
      html += '<div class="detail-row"><div class="detail-row-label">Status</div><div class="detail-row-value" style="color:var(--info,#60a5fa)">Upcoming — awaiting move-out</div></div>';
      html += '<div class="detail-row"><div class="detail-row-label">Days Until</div><div class="detail-row-value">' + Math.abs(p.elapsed) + ' days</div></div>';
    }
    html += '</div>';

    if (p.isRegisteredUnitTurn && p.relatedTurns && p.relatedTurns.length > 0) {
      html += '<div class="detail-section-title" style="margin-top:12px"><i class="fas fa-link"></i> Other Unit Turns (' + p.relatedTurns.length + ')</div>';
      p.relatedTurns.forEach(function(rt) {
        html += '<div class="pipe-wo-item">';
        html += '<div><span class="pipe-wo-id">' + escapeHtml(rt.unitTurnId || 'turn') + '</span> <span class="tag ' + (rt.turnEnd ? 'completed' : 'waiting') + '">' + escapeHtml(rt.status || (rt.turnEnd ? 'Completed' : 'In Progress')) + '</span></div>';
        html += '<div style="font-size:11px;color:var(--text-secondary)">Move-out: ' + (rt.moveOut ? formatDate(rt.moveOut) : '\u2014') + ' • Move-in: ' + (rt.expectedMoveIn ? formatDate(rt.expectedMoveIn) : '\u2014') + ' • Billed: ' + escapeHtml(rt.totalBilled || '$0') + '</div>';
        html += '</div>';
      });
    }

    // Webhook events
    if (p.webhookEvents.length > 0) {
      html += '<div class="detail-section-title" style="margin-top:12px"><i class="fas fa-plug"></i> Webhook Events (' + p.webhookEvents.length + ')</div>';
      p.webhookEvents.slice(0, 5).forEach(function(wh) {
        html += '<div style="font-size:11px;padding:4px 0;border-bottom:1px solid var(--border)">';
        html += '<span style="color:var(--text-muted)">' + timeAgo(wh.ts) + '</span> ';
        html += '<strong>' + escapeHtml(wh.title) + '</strong>';
        if (wh.body) html += ' — ' + escapeHtml(wh.body.substring(0, 80));
        html += '</div>';
      });
    }

    html += '</div>'; // end right col
    html += '</div>'; // end detail-grid

    // Actions
    html += '<div class="pipe-actions">';
    var nextIdx = p.currentStageIdx + 1;
    if (nextIdx < PIPE_STAGES.length && !p.isCompleted) {
      html += '<button class="action-btn primary" data-advance="' + escapeHtml(p.id) + '" data-stage="' + PIPE_STAGES[nextIdx].key + '"><i class="fas fa-arrow-right"></i> Confirm ' + PIPE_STAGES[nextIdx].title + '</button>';
    }
    // AppFolio deep link — prefer DB API direct link, then unitTurnId, then property
    var turnAfUrl = p.dbLink || appfolioUrl('unit_turn', p.registeredUnitTurnId || (p.turn && p.turn.unitTurnId) || '') || appfolioUrl('property', p.propertyId);
    if (turnAfUrl) {
      html += '<a class="action-btn" href="' + escapeHtml(turnAfUrl) + '" target="_blank" rel="noopener noreferrer" style="text-decoration:none" onclick="event.stopPropagation()"><i class="fas fa-external-link-alt"></i> View in AppFolio</a>';
    }
    html += '<button class="action-btn" data-close-detail="' + escapeHtml(p.id) + '"><i class="fas fa-times"></i> Close</button>';
    html += '</div>';

    html += '</div>'; // end pipe-detail
  });

  container.innerHTML = html;
  if (OPEN_TURN_DETAIL_ID) {
    var openCard = Array.prototype.find.call(container.querySelectorAll('.pipe-card'), function(card) {
      return card.getAttribute('data-pipeid') === OPEN_TURN_DETAIL_ID;
    });
    if (openCard) {
      var openIdx = openCard.getAttribute('data-pipeidx');
      var openDetail = document.getElementById('pipeDetail_' + openIdx);
      if (openDetail) openDetail.classList.add('show');
    } else {
      OPEN_TURN_DETAIL_ID = '';
    }
  }
  // Event listeners handled by delegation in wireUpUI() — no re-attachment needed
}

// Manual stage advancement
async function confirmTurnStage(turnId, stageKey) {
  var stageData = { done: true, date: new Date().toISOString(), manual: true };

  // Update local record first
  var rec = TURN_RECORDS.find(function(r) { return r.id === turnId; });
  if (!rec) {
    rec = { id: turnId, stages: {} };
    TURN_RECORDS.push(rec);
    // Cap TURN_RECORDS at 500 to prevent unbounded growth
    if (TURN_RECORDS.length > 500) TURN_RECORDS = TURN_RECORDS.slice(-500);
    // Create the record on the server first so stage update doesn't 404
    await saveTurnRecord(rec);
  }
  if (!rec.stages) rec.stages = {};
  rec.stages[stageKey] = stageData;

  // Persist stage to proxy
  try {
    await saveTurnRecordStage(turnId, stageKey, stageData);
  } catch (err) {
    // If stage update fails (e.g. 404), save full record as fallback
    await saveTurnRecord(rec);
  }

  // Re-render
  renderTurnBoard();
  var stageLabel = PIPE_STAGES.find(function(s) { return s.key === stageKey; });
  showToast('Stage confirmed: ' + (stageLabel ? stageLabel.title : stageKey));
}

/* =================================================================
   INSPECTIONS — Enhanced with KPIs + Turn-linking
   ================================================================= */
function renderInspections(search) {
  var body = $('#inspBody');
  if (!body) return;

  var statusFilter = $('#inspStatusFilter') ? $('#inspStatusFilter').value : 'all';
  var today = new Date();

  // Filter out pre-AppFolio era artifacts (before 2021) from stale cache
  var validInspections = INSPECTIONS.filter(function(r) {
    var state = getInspectionCompliance(r, today);
    return !state.anchorDate || state.anchorDate >= INSPECTION_AF_EPOCH;
  });

  // Classify each inspection
  var classified = validInspections.map(function(r) {
    var state = getInspectionCompliance(r, today);
    // Check if linked to an active turn
    var linkedTurn = TURN_PIPE_DATA.find(function(tp) {
      return !tp.isCompleted &&
        tp.unit && r.unit && String(tp.unit).toLowerCase() === String(r.unit).toLowerCase() &&
        tp.property && r.propertyName && String(tp.property).toLowerCase() === String(r.propertyName).toLowerCase();
    });
    return {
      r: r,
      daysSince: state.daysSince,
      overdue: state.overdue,
      dueSoon: state.dueSoon,
      current: state.current,
      missingMoveInInspection: state.missingMoveInInspection,
      linkedTurn: linkedTurn || null,
      status: state.overdue ? 'overdue' : state.dueSoon ? 'due_soon' : 'current'
    };
  });

  // KPI counts
  var overdueCount = classified.filter(function(c) { return c.overdue; }).length;
  var dueSoonCount = classified.filter(function(c) { return c.dueSoon; }).length;
  var currentCount = classified.filter(function(c) { return c.current; }).length;
  var turnLinkedCount = classified.filter(function(c) { return c.linkedTurn; }).length;

  var e = function(id, v) { var el = document.getElementById(id); if (el) el.textContent = v; };
  e('kpiInspOverdue', overdueCount);
  e('kpiInspDueSoon', dueSoonCount);
  e('kpiInspCurrent', currentCount);
  e('kpiInspTurnLinked', turnLinkedCount);

  // Sync global group filter dropdown
  var inspGrpSel = $('#globalGroupFilter');
  if (inspGrpSel && inspGrpSel.value !== currentPropertyGroup) inspGrpSel.value = currentPropertyGroup;

  // Filter
  var filtered = classified.filter(function(c) {
    if (statusFilter === 'overdue' && !c.overdue) return false;
    if (statusFilter === 'due_soon' && !c.dueSoon) return false;
    if (statusFilter === 'current' && !c.current) return false;
    if (statusFilter === 'turn_linked' && !c.linkedTurn) return false;
    // Property group filter
    if (!isInPropertyGroup(c.r.propertyId, c.r.propertyName, currentPropertyGroup)) return false;
    if (search) {
      var s = search.toLowerCase();
      return (c.r.propertyName || '').toLowerCase().indexOf(s) !== -1
        || (c.r.unit || '').toLowerCase().indexOf(s) !== -1
        || (c.r.tenant || '').toLowerCase().indexOf(s) !== -1;
    }
    return true;
  });

  // Sort filtered results
  var sortCol = _inspSortCol;
  var sortDir = _inspSortDir === 'asc' ? 1 : -1;
  filtered.sort(function(a, b) {
    var va, vb;
    if (sortCol === 'property') { va = (a.r.propertyName || '').toLowerCase(); vb = (b.r.propertyName || '').toLowerCase(); }
    else if (sortCol === 'unit') { va = (a.r.unit || '').toLowerCase(); vb = (b.r.unit || '').toLowerCase(); }
    else if (sortCol === 'daysSince') { va = a.daysSince; vb = b.daysSince; }
    else if (sortCol === 'status') { va = a.overdue ? 0 : a.dueSoon ? 1 : 2; vb = b.overdue ? 0 : b.dueSoon ? 1 : 2; }
    else { va = 0; vb = 0; }
    if (va < vb) return -1 * sortDir;
    if (va > vb) return 1 * sortDir;
    return 0;
  });

  // Update sort indicators in table header
  document.querySelectorAll('[data-inspsort]').forEach(function(th) {
    th.classList.remove('asc', 'desc');
    if (th.getAttribute('data-inspsort') === sortCol) th.classList.add(_inspSortDir);
  });

  if (filtered.length === 0) {
    body.innerHTML = '<tr><td colspan="8">' + emptyHtml('fa-clipboard-check', INSPECTIONS.length === 0 ? 'No inspection data. Try refreshing.' : 'No inspections match filter') + '</td></tr>';
    return;
  }

  var html = '';
  filtered.forEach(function(c, idx) {
    var r = c.r;
    var statusTag = c.missingMoveInInspection
      ? '<span class="tag non-compliant" title="Move-in exists but no inspection on/after move-in">Missing move-in inspection</span>'
      : c.overdue
      ? '<span class="tag non-compliant">Overdue</span>'
      : c.dueSoon
        ? '<span class="tag" style="background:var(--warning-dim);color:var(--warning)">Due soon</span>'
        : '<span class="tag compliant">Current</span>';
    var turnTag = c.linkedTurn
      ? '<span class="tag assigned" title="Linked to active turn"><i class="fas fa-exchange-alt" style="font-size:8px"></i> ' + escapeHtml(c.linkedTurn.unit) + '</span>'
      : '<span style="color:var(--text-muted)">\u2014</span>';
    html += '<tr class="insp-row" data-inspidx="' + idx + '" style="cursor:pointer;' + (c.overdue ? 'background:var(--danger-dim)' : '') + '">';
    html += '<td>' + escapeHtml(r.propertyName) + ' <i class="fas fa-external-link-alt" style="font-size:8px;opacity:0.4"></i></td>';
    html += '<td>' + escapeHtml(r.unit) + '</td>';
    html += '<td style="font-family:var(--font-mono)">' + (r.lastInspection ? formatDate(r.lastInspection) + ' <span style="color:var(--text-muted);font-size:10px">(' + c.daysSince + 'd ago)</span>' : '<span style="color:var(--danger)">Never</span>') + '</td>';
    html += '<td>' + escapeHtml(r.tenant || '\u2014') + '</td>';
    html += '<td style="font-family:var(--font-mono)">' + (r.moveIn ? formatDate(r.moveIn) : '\u2014') + '</td>';
    html += '<td style="font-family:var(--font-mono)">' + (r.moveOut ? formatDate(r.moveOut) : '\u2014') + '</td>';
    html += '<td>' + statusTag + '</td>';
    html += '<td>' + turnTag + '</td>';
    html += '</tr>';
  });
  body.innerHTML = html;
  // Event listeners handled by delegation in wireUpUI() — no re-attachment needed

  // Store filtered data for delegation handler to access
  body._filteredData = filtered;

  var ib = $('#inspBadge');
  if (ib) ib.textContent = overdueCount;
}

function vendorCatClass(cat) {
  if (!cat) return 'cat-uncategorized';
  return 'cat-' + cat.toLowerCase().replace(/[\s]+/g, '-');
}

function getVendorInitial(name) {
  var first = String(name || '').trim().charAt(0).toUpperCase();
  return /[A-Z]/.test(first) ? first : '#';
}

function resolveVendorCompliance(v) {
  // Manual override takes priority over API value
  var manual = isVendorManuallyCompliant(v.id);
  if (manual !== null) return { compliant: manual, isManual: true };
  return { compliant: v.compliant, isManual: false };
}

function renderVendors(search) {
  var container = $('#vendorGrid');
  if (!container) return;

  var sec = $('#sec-vendors');
  if (sec && !sec.classList.contains('active')) {
    _vendorsNeedRender = true;
    return;
  }
  _vendorsNeedRender = false;

  // Sync global group filter dropdown
  var vendGrpSel = $('#globalGroupFilter');
  if (vendGrpSel && vendGrpSel.value !== currentPropertyGroup) vendGrpSel.value = currentPropertyGroup;

  // Read filter states
  var catFilter = $('#vendorCategoryFilter') ? $('#vendorCategoryFilter').value : '';
  var tradeFilter = $('#vendorTradeFilter') ? $('#vendorTradeFilter').value : '';
  var sortMode = $('#vendorSortMode') ? $('#vendorSortMode').value : 'name';
  var compFilter = $('#vendorComplianceFilter') ? $('#vendorComplianceFilter').value : '';

  var tradeCounts = {};
  VENDORS.forEach(function(v) {
    var tc = getVendorTradeCategory(v);
    tradeCounts[tc] = (tradeCounts[tc] || 0) + 1;
  });
  var tradeSel = $('#vendorTradeFilter');
  if (tradeSel) {
    var currentTradeFilter = tradeFilter;
    tradeSel.innerHTML = '<option value="">All Trade Categories</option>';
    VENDOR_TRADE_CATEGORIES.slice().sort(function(a, b) { return a.localeCompare(b); }).forEach(function(tc) {
      if (!tradeCounts[tc]) return;
      var opt = document.createElement('option');
      opt.value = tc;
      opt.textContent = tc + ' (' + tradeCounts[tc] + ')';
      tradeSel.appendChild(opt);
    });
    if (currentTradeFilter && !tradeCounts[currentTradeFilter]) currentTradeFilter = '';
    tradeSel.value = currentTradeFilter;
    tradeFilter = currentTradeFilter;
  }

  // Note: Vendors don't have a direct property association for group filtering,
  // but we can filter by checking which vendors have WOs in the group
  var vendorsInGroup = null;
  if (currentPropertyGroup) {
    vendorsInGroup = {};
    WORK_ORDERS.forEach(function(wo) {
      if (isInPropertyGroup(wo.propertyId, wo.propertyName, currentPropertyGroup) && wo.vendorName) {
        vendorsInGroup[wo.vendorName.toLowerCase()] = true;
      }
    });
  }

  var searchText = (search || '').trim();
  var searchLower = searchText.toLowerCase();
  var baseFiltered = VENDORS.filter(function(v) {
    if (vendorsInGroup && !vendorsInGroup[(v.name || '').toLowerCase()]) return false;

    // Category filter
    if (catFilter) {
      var vCat = getVendorCategory(v.id) || 'Uncategorized';
      if (vCat !== catFilter) return false;
    }

    // Compliance filter
    if (compFilter) {
      var res = resolveVendorCompliance(v);
      if (compFilter === 'compliant' && !res.compliant) return false;
      if (compFilter === 'non-compliant' && res.compliant) return false;
      if (compFilter === 'manual-compliant' && !(res.compliant && res.isManual)) return false;
    }

    if (tradeFilter) {
      var vTradeCat = getVendorTradeCategory(v);
      if (vTradeCat !== tradeFilter) return false;
    }

    if (!searchText) return true;
    var s = searchLower;
    var vCatStr = getVendorCategory(v.id) || '';
    return (v.name || '').toLowerCase().indexOf(s) !== -1
      || (v.email || '').toLowerCase().indexOf(s) !== -1
      || (v.trades || '').toLowerCase().indexOf(s) !== -1
      || vCatStr.toLowerCase().indexOf(s) !== -1;
  });

  baseFiltered.sort(function(a, b) {
    var aName = String(a.name || '').toLowerCase();
    var bName = String(b.name || '').toLowerCase();
    if (sortMode === 'trade') {
      var aTrade = getVendorTradeCategory(a);
      var bTrade = getVendorTradeCategory(b);
      if (aTrade !== bTrade) return aTrade.localeCompare(bTrade);
    }
    return aName.localeCompare(bName);
  });

  var availableInitials = {};
  baseFiltered.forEach(function(v) { availableInitials[getVendorInitial(v.name)] = true; });

  var filtered = baseFiltered.filter(function(v) {
    return !currentVendorInitial || getVendorInitial(v.name) === currentVendorInitial;
  });

  if (filtered.length === 0) {
    container.innerHTML = emptyHtml('fa-hard-hat', VENDORS.length === 0 ? 'No vendors loaded' : 'No vendors match filters');
    return;
  }

  var renderKey = [currentPropertyGroup || '', catFilter || '', tradeFilter || '', sortMode || 'name', compFilter || '', searchLower, currentVendorInitial || ''].join('|');
  if (_vendorRenderKey !== renderKey) {
    _vendorRenderKey = renderKey;
    _vendorRenderLimit = isConstrainedDevice() ? CONFIG.VENDOR_GRID_INITIAL_LIMIT_MOBILE : CONFIG.VENDOR_GRID_INITIAL_LIMIT_DESKTOP;
    if (searchText.length >= CONFIG.VENDOR_SELECT_MIN_SEARCH_CHARS) {
      _vendorRenderLimit = Math.max(_vendorRenderLimit, CONFIG.VENDOR_GRID_LOAD_MORE_STEP * 2);
    }
  }

  var visible = filtered.slice(0, _vendorRenderLimit);
  var constrained = isConstrainedDevice();
  container.classList.toggle('stagger', !(constrained || filtered.length > 120));
  var html = '<div class="vendor-render-meta">Showing ' + visible.length + ' of ' + filtered.length + ' vendors' + (searchText ? ' for "' + escapeHtml(searchText) + '"' : '') + (currentVendorInitial ? ' • initial ' + escapeHtml(currentVendorInitial) : '') + '.</div>';
  html += '<div class="vendor-directory-shell">';
  html += '<div class="vendor-directory-rail">';
  html += '<button class="vendor-rail-btn' + (!currentVendorInitial ? ' active' : '') + '" data-vendor-initial="">All</button>';
  'ABCDEFGHIJKLMNOPQRSTUVWXYZ#'.split('').forEach(function(letter) {
    var enabled = !!availableInitials[letter];
    html += '<button class="vendor-rail-btn' + (currentVendorInitial === letter ? ' active' : '') + '" data-vendor-initial="' + letter + '"' + (enabled ? '' : ' disabled') + '>' + letter + '</button>';
  });
  html += '</div>';
  html += '<div class="vendor-cards">';
  var today = new Date();
  visible.forEach(function(v) {
    var ed = v.insurance ? new Date(v.insurance) : null;
    var exp = ed ? ed < today : false;
    var due = ed ? daysBetween(today, ed) : 999;
    var wrn = !exp && due <= 60;
    var cRes = resolveVendorCompliance(v);
    var cc = '';
    if (cRes.compliant && cRes.isManual) { cc = 'manual-compliant'; }
    else if (exp) { cc = 'expired'; }
    else if (wrn) { cc = 'warn'; }
    var vCat = getVendorCategory(v.id);
    var catBadgeCls = vendorCatClass(vCat);
    var afUrl = appfolioUrl('vendor', v.id);

    html += '<div class="vendor-card vendor-card-compact ' + cc + '" data-vendorid="' + escapeHtml(String(v.id)) + '" data-vendor-initial="' + getVendorInitial(v.name) + '" style="cursor:pointer">';
    html += '<div class="vendor-card-head">';
    html += '<div class="vendor-name">' + escapeHtml(v.name) + (afUrl ? ' <a href="' + escapeHtml(afUrl) + '" target="_blank" rel="noopener noreferrer" style="font-size:10px;color:var(--accent);text-decoration:none" title="View in AppFolio" onclick="event.stopPropagation()"><i class="fas fa-external-link-alt"></i></a>' : '') + '</div>';
    html += '<div class="vendor-id"><span><i class="fas fa-fingerprint"></i> ' + escapeHtml(String(v.id)) + '</span>';
    html += '<span class="vendor-category-badge ' + catBadgeCls + '">' + escapeHtml(vCat || 'Uncategorized') + '</span>';
    html += '</div>';
    html += '</div>';
    if (v.trades) html += '<div class="vendor-trades">' + escapeHtml(v.trades) + '</div>';
    html += '<div class="vendor-row vendor-row-compact" style="align-items:center"><span class="vendor-row-label">Category</span><select class="vendor-cat-select" data-vid="' + escapeHtml(String(v.id)) + '" onclick="event.stopPropagation()">';
    VENDOR_CATEGORIES.forEach(function(c) {
      html += '<option value="' + escapeHtml(c) + '"' + ((vCat || 'Uncategorized') === c ? ' selected' : '') + '>' + escapeHtml(c) + '</option>';
    });
    html += '</select></div>';
    var vTradeCat = getVendorTradeCategory(v);
    html += '<div class="vendor-row vendor-row-compact" style="align-items:center"><span class="vendor-row-label">Trade Cat.</span><select class="vendor-trade-cat-select" data-vid="' + escapeHtml(String(v.id)) + '" onclick="event.stopPropagation()">';
    VENDOR_TRADE_CATEGORIES.forEach(function(tc) {
      html += '<option value="' + escapeHtml(tc) + '"' + (vTradeCat === tc ? ' selected' : '') + '>' + escapeHtml(tc) + '</option>';
    });
    html += '</select></div>';
    var compLabel = cRes.compliant ? 'Compliant' : (v.compliantStatus || 'Non-Compliant');
    var toggleCls = 'vendor-compliance-toggle' + (cRes.compliant ? ' is-compliant' : '') + (cRes.isManual ? ' is-manual' : '');
    html += '<div class="vendor-row vendor-row-compact" style="align-items:center"><span class="vendor-row-label">Compliance</span>';
    html += '<button class="' + toggleCls + '" data-vid="' + escapeHtml(String(v.id)) + '" onclick="event.stopPropagation()" title="Click to toggle compliance manually">';
    html += '<i class="fas ' + (cRes.compliant ? 'fa-check-circle' : 'fa-times-circle') + '"></i> ' + escapeHtml(compLabel);
    if (cRes.isManual) html += ' <span style="font-size:9px;opacity:0.7">(manual)</span>';
    html += '</button></div>';
    if (v.phone || v.email) {
      html += '<div class="vendor-row vendor-row-compact"><span class="vendor-row-label">Contact</span><span class="vendor-row-value vendor-contact-compact">' + escapeHtml(v.phone || v.email || '\u2014') + (v.phone && v.email ? ' • ' + escapeHtml(v.email) : '') + '</span></div>';
    }
    if (v.insurance) {
      html += '<div class="vendor-row vendor-row-compact"><span class="vendor-row-label">Insurance</span><span class="vendor-row-value" style="font-family:var(--font-mono);color:' + (exp ? 'var(--danger)' : wrn ? 'var(--warning)' : 'var(--text-secondary)') + '">' + escapeHtml(v.insurance) + (exp ? ' (EXPIRED)' : wrn ? ' (' + due + 'd)' : '') + '</span></div>';
    }
    html += '</div>';
  });
  if (filtered.length > visible.length) {
    var remaining = filtered.length - visible.length;
    html += '<div class="vendor-load-more-wrap"><button class="vendor-load-more-btn" data-vendor-load-more="1">Load ' + Math.min(CONFIG.VENDOR_GRID_LOAD_MORE_STEP, remaining) + ' More (' + remaining + ' remaining)</button></div>';
  }
  html += '</div></div>';
  container.innerHTML = html;
  // Event listeners handled by delegation in wireUpUI() — no re-attachment needed
}

function renderNewWOVendorOptions(filterText) {
  var vendSelect = $('#nwoVendor');
  if (!vendSelect) return;

  var hint = $('#nwoVendorHint');
  var term = (filterText || '').trim().toLowerCase();
  var minChars = CONFIG.VENDOR_SELECT_MIN_SEARCH_CHARS;
  var selectedId = String(vendSelect.value || '');
  var constrained = isConstrainedDevice();

  var matches = VENDORS.filter(function(v) {
    if (!term) return true;
    return (v.name || '').toLowerCase().indexOf(term) !== -1
      || (v.email || '').toLowerCase().indexOf(term) !== -1
      || (v.trades || '').toLowerCase().indexOf(term) !== -1;
  });

  var limit;
  if (term && term.length >= minChars) {
    limit = CONFIG.VENDOR_SELECT_SEARCH_LIMIT;
  } else {
    limit = constrained ? CONFIG.VENDOR_SELECT_INITIAL_LIMIT_MOBILE : CONFIG.VENDOR_SELECT_INITIAL_LIMIT_DESKTOP;
  }

  var visible = matches.slice(0, limit);
  var opts = ['<option value="">— Select Vendor —</option>'];
  visible.forEach(function(v) {
    opts.push('<option value="' + escapeHtml(String(v.id)) + '">' + escapeHtml(v.name) + '</option>');
  });

  if (selectedId && !visible.some(function(v) { return String(v.id) === selectedId; })) {
    var selectedVendor = VENDORS.find(function(v) { return String(v.id) === selectedId; });
    if (selectedVendor) {
      opts.splice(1, 0, '<option value="' + escapeHtml(String(selectedVendor.id)) + '">' + escapeHtml(selectedVendor.name) + ' (selected)</option>');
    }
  }

  vendSelect.innerHTML = opts.join('');
  if (selectedId) vendSelect.value = selectedId;

  if (!hint) return;
  if (matches.length === 0) {
    hint.textContent = term ? 'No vendor matches this search.' : 'No vendors loaded yet.';
    return;
  }

  if (!term && VENDORS.length > limit) {
    hint.textContent = 'Showing first ' + visible.length + ' of ' + VENDORS.length + ' vendors. Type at least ' + minChars + ' characters to narrow results.';
    return;
  }

  if (matches.length > visible.length) {
    hint.textContent = 'Showing first ' + visible.length + ' of ' + matches.length + ' matches. Refine search to narrow.';
    return;
  }

  hint.textContent = 'Showing ' + visible.length + ' vendor' + (visible.length === 1 ? '' : 's') + '.';
}

/* renderReconciliation — Removed (billing stripped to lighten payload) */

function renderTemplates() {
  var container = $('#templateGrid');
  var html = '';
  TEMPLATES.forEach(function(t) {
    html += '<div class="template-card"><div class="template-card-title"><i class="fas ' + t.icon + '" style="color:var(--accent)"></i> ' + escapeHtml(t.title) + '</div>';
    html += '<div class="template-card-desc"><i class="fas fa-bolt" style="font-size:10px"></i> Trigger: ' + escapeHtml(t.trigger) + '</div>';
    html += '<div class="template-preview">' + t.body + '</div>';
    html += '<div style="margin-top:10px;display:flex;gap:6px">';
    html += '<button class="action-btn" style="flex:1" data-tcopy><i class="fas fa-copy"></i> Copy</button>';
    html += '<button class="action-btn" style="flex:1" data-tedit><i class="fas fa-edit"></i> Edit</button>';
    html += '</div></div>';
  });
  container.innerHTML = html;
  // Event listeners handled by delegation in wireUpUI() — no re-attachment needed
}

function renderErrorLog() {
  var container = $('#errorLog');
  if (API_ERRORS.length === 0) {
    container.innerHTML = '<div style="padding:20px;text-align:center;color:var(--text-muted);font-size:12px"><i class="fas fa-check-circle" style="color:var(--success);margin-right:6px"></i> No API errors recorded this session</div>';
    return;
  }
  var html = '';
  API_ERRORS.forEach(function(e) {
    var codeLabel = e.code === 0 ? 'CORS' : String(e.code);
    html += '<div class="error-row"><span class="error-code c' + e.code + '">' + codeLabel + '</span>';
    html += '<span class="error-ts">' + escapeHtml(e.ts) + '</span><span class="error-msg">' + escapeHtml(e.msg) + '</span>';
    html += '<span class="error-action ' + e.action + '">' + (e.action === 'retry' ? 'RETRY' : e.action === 'resolved' ? 'RESOLVED' : 'QUEUED') + '</span></div>';
  });
  container.innerHTML = html;
}

function populateDropdowns() {
  // Properties dropdown for New WO modal (build string first, assign once)
  var propSelect = $('#nwoProperty');
  if (propSelect) {
    var propOpts = ['<option value="">— Select Property —</option>'];
    PROPERTIES.forEach(function(p) {
      propOpts.push('<option value="' + escapeHtml(String(p.id)) + '">' + escapeHtml(p.name) + (p.address ? ' \u2014 ' + escapeHtml(p.address) : '') + '</option>');
    });
    // Also add properties extracted from work orders if API didn't return properties
    if (PROPERTIES.length === 0) {
      var seen = {};
      WORK_ORDERS.forEach(function(wo) {
        if (wo.propertyName && !seen[wo.propertyName]) {
          seen[wo.propertyName] = true;
          propOpts.push('<option value="' + escapeHtml(wo.propertyName) + '">' + escapeHtml(wo.propertyName) + '</option>');
        }
      });
    }
    propSelect.innerHTML = propOpts.join('');
  }

  // Vendors dropdown for New WO modal (build string first, assign once)
  var vendSelect = $('#nwoVendor');
  if (vendSelect) {
    renderNewWOVendorOptions($('#nwoVendorSearch') ? $('#nwoVendorSearch').value : '');
  }

  // Populate ALL group filter dropdowns across every tab
  populateGroupFilters();
}

/* =================================================================
   RENDER ALL — convenience wrapper
   ================================================================= */
function renderAll() {
  // Each render is wrapped in try/catch so one crash doesn't kill the others
  var fns = [
    function() { renderWorkOrders(); },
    function() { renderVendors($('#vendorSearch') ? $('#vendorSearch').value : ''); },
    function() { renderTurnBoard(); },
    function() { renderInspections($('#inspSearch') ? $('#inspSearch').value : ''); },
    function() { renderPayroll(); },
    function() { renderMoveOuts(); },
    function() { renderDashboardKPIs(); },
    function() { renderActivityFeed(); },
    function() { populateDropdowns(); }
  ];
  fns.forEach(function(fn) {
    try { fn(); } catch (e) { console.log('renderAll sub-error: ' + (e.message || e)); }
  });
}

/* =================================================================
   INIT — Cache-first with WO-based pre-flight
   ================================================================= */
var _uiWired = false;

// What's New popup — shown once per 36 hours after vault unlock
function maybeShowWhatsNew() {
  var WHATS_NEW_KEY = 'hnmgr_whats_new_v92';
  if (localStorage.getItem(WHATS_NEW_KEY + '_skip') === '1') return;
  var lastShown = parseInt(localStorage.getItem(WHATS_NEW_KEY) || '0', 10);
  if (Date.now() - lastShown < 36 * 60 * 60 * 1000) return;
  localStorage.setItem(WHATS_NEW_KEY, String(Date.now()));
  setTimeout(function() { openModal('whatsNewModal'); }, 800);
}
function dismissWhatsNew() {
  if ($('#whatsNewDontShow') && $('#whatsNewDontShow').checked) {
    localStorage.setItem('hnmgr_whats_new_v92_skip', '1');
  }
  closeModal('whatsNewModal');
}

function wireUpUI() {
  if (_uiWired) return;
  _uiWired = true;

  /* =================================================================
     EVENT DELEGATION — Single listeners on parent containers
     Replaces per-render querySelectorAll+addEventListener patterns.
     Set up ONCE here; never re-attached on re-render.
     ================================================================= */

  // Kanban board — card clicks + column header clicks
  (function() {
    var board = $('#kanbanBoard');
    if (board) board.addEventListener('click', function(e) {
      var card = e.target.closest('.kanban-card');
      if (card) { showWODetail(card.getAttribute('data-woid')); return; }
      var head = e.target.closest('.kanban-col-head');
      if (head) {
        var status = head.getAttribute('data-status');
        if (currentWOFilter === status) { currentWOFilter = 'all'; }
        else { currentWOFilter = status; }
        $$('[data-filter]').forEach(function(b) {
          b.classList.toggle('active', b.getAttribute('data-filter') === currentWOFilter);
        });
        renderWorkOrders();
      }
    });
  })();

  // Payroll table — row clicks + flag toggles
  (function() {
    var payBody = $('#payrollBody');
    if (payBody) payBody.addEventListener('click', function(e) {
      var flagBtn = e.target.closest('[data-flagwo]');
      if (flagBtn) {
        e.stopPropagation();
        var wid = flagBtn.getAttribute('data-flagwo');
        toggleFlag(wid).then(function() { renderPayroll(); renderWorkOrders(); });
        return;
      }
      var row = e.target.closest('.payroll-row');
      if (row) {
        var woid = row.getAttribute('data-woid');
        var wo = WORK_ORDERS.find(function(w) { return String(w.id) === woid; });
        if (!wo) return;
        showItemDetail('Payroll \u2014 WO #' + wo.id, [
          { section: 'Work Order', icon: 'fa-wrench' },
          { label: 'WO Number', value: '#' + wo.id },
          { label: 'Property', value: wo.propertyName },
          { label: 'Unit', value: wo.unit },
          { label: 'Description', value: wo.description },
          { label: 'Vendor', value: wo.vendorName || '\u2014' },
          { label: 'Status', value: wo.status },
          { label: 'Priority', value: wo.priority },
          { section: 'Payroll', icon: 'fa-money-check-alt' },
          { label: 'Completed', value: formatDate(wo.workCompletedOn || wo.completedOn) },
          { label: 'Amount', value: wo.amount ? currency(parseFloat(wo.amount)) : '\u2014' },
          { label: 'Tenant', value: wo.tenant || '\u2014' },
          { label: 'Assigned To', value: wo.assignedUser || '\u2014' }
        ], appfolioUrl('work_order', wo.id || wo.uuid));
      }
    });
  })();

  // Turn pipeline — card clicks + stage advance buttons
  (function() {
    var pipeline = $('#turnPipeline');
    if (pipeline) pipeline.addEventListener('click', function(e) {
      var closeDetailBtn = e.target.closest('[data-close-detail]');
      if (closeDetailBtn) {
        e.stopPropagation();
        OPEN_TURN_DETAIL_ID = '';
        var detailPanel = closeDetailBtn.closest('.pipe-detail');
        if (detailPanel) detailPanel.classList.remove('show');
        return;
      }
      var advBtn = e.target.closest('[data-advance]');
      if (advBtn) {
        e.stopPropagation();
        confirmTurnStage(advBtn.getAttribute('data-advance'), advBtn.getAttribute('data-stage'));
        return;
      }
      var card = e.target.closest('.pipe-card');
      if (card) {
        var idx = card.getAttribute('data-pipeidx');
        var detail = document.getElementById('pipeDetail_' + idx);
        if (detail) {
          var isOpen = detail.classList.contains('show');
          pipeline.querySelectorAll('.pipe-detail').forEach(function(d) { d.classList.remove('show'); });
          OPEN_TURN_DETAIL_ID = isOpen ? '' : card.getAttribute('data-pipeid');
          if (!isOpen) detail.classList.add('show');
        }
      }
    });
  })();

  // Inspection table — row clicks
  (function() {
    var inspBody = $('#inspBody');
    if (inspBody) inspBody.addEventListener('click', function(e) {
      var row = e.target.closest('.insp-row');
      if (!row) return;
      var idx = parseInt(row.getAttribute('data-inspidx'), 10);
      var filtered = inspBody._filteredData;
      if (!filtered || !filtered[idx]) return;
      var c = filtered[idx];
      var r = c.r;
      showItemDetail('Inspection \u2014 ' + r.propertyName + ' ' + r.unit, [
        { section: 'Inspection Details', icon: 'fa-clipboard-check' },
        { label: 'Property', value: r.propertyName },
        { label: 'Unit', value: r.unit },
        { label: 'Last Inspection', value: r.lastInspection ? formatDate(r.lastInspection) + ' (' + c.daysSince + ' days ago)' : 'Never' },
        { label: 'Status', value: c.overdue ? 'OVERDUE' : c.dueSoon ? 'Due Soon' : 'Current' },
        { section: 'Tenant', icon: 'fa-user' },
        { label: 'Tenant', value: r.tenant || '\u2014' },
        { label: 'Move-In', value: r.moveIn ? formatDate(r.moveIn) : '\u2014' },
        { label: 'Move-Out', value: r.moveOut ? formatDate(r.moveOut) : '\u2014' },
        { label: 'Tags', value: r.tags || '\u2014' }
      ], appfolioUrl('inspection_property', r.propertyId));
    });
  })();

  // Inspection table — sortable header clicks
  document.querySelectorAll('[data-inspsort]').forEach(function(th) {
    th.addEventListener('click', function() {
      var col = this.getAttribute('data-inspsort');
      if (_inspSortCol === col) {
        _inspSortDir = _inspSortDir === 'asc' ? 'desc' : 'asc';
      } else {
        _inspSortCol = col;
        _inspSortDir = col === 'daysSince' ? 'desc' : 'asc';
      }
      renderInspections($('#inspSearch') ? $('#inspSearch').value : '');
    });
  });

  // Vendor grid — card clicks, compliance toggles, category selects
  (function() {
    var vendGrid = $('#vendorGrid');
    if (!vendGrid) return;
    vendGrid.addEventListener('click', function(e) {
      var initialBtn = e.target.closest('[data-vendor-initial]');
      if (initialBtn) {
        e.stopPropagation();
        currentVendorInitial = initialBtn.getAttribute('data-vendor-initial') || '';
        renderVendors($('#vendorSearch') ? $('#vendorSearch').value : '');
        return;
      }

      var moreBtn = e.target.closest('[data-vendor-load-more]');
      if (moreBtn) {
        e.stopPropagation();
        _vendorRenderLimit += CONFIG.VENDOR_GRID_LOAD_MORE_STEP;
        renderVendors($('#vendorSearch') ? $('#vendorSearch').value : '');
        return;
      }

      // Compliance toggle
      var compBtn = e.target.closest('.vendor-compliance-toggle');
      if (compBtn) {
        e.stopPropagation();
        var vid = compBtn.getAttribute('data-vid');
        var current = isVendorManuallyCompliant(vid);
        var newVal;
        if (current === null) { newVal = true; }
        else if (current === true) { newVal = false; }
        else { newVal = null; }
        saveVendorOverride(vid, { compliant: newVal }).then(function() {
          renderVendors($('#vendorSearch') ? $('#vendorSearch').value : '');
          var label = newVal === true ? 'Marked compliant (manual)' : newVal === false ? 'Marked non-compliant (manual)' : 'Reset to API value';
          showToast(label);
        });
        return;
      }
      // Vendor card click (detail modal)
      var card = e.target.closest('.vendor-card');
      if (card && !e.target.closest('.vendor-cat-select') && !e.target.closest('.vendor-trade-cat-select')) {
        var vid2 = card.getAttribute('data-vendorid');
        var v = VENDORS.find(function(vn) { return String(vn.id) === vid2; });
        if (!v) return;
        var cRes = resolveVendorCompliance(v);
        var vCat = getVendorCategory(v.id) || 'Uncategorized';
        var compText = cRes.compliant ? 'Compliant' : (v.compliantStatus || 'Non-Compliant');
        if (cRes.isManual) compText += ' (manual override)';
        showItemDetail('Vendor \u2014 ' + v.name, [
          { section: 'Vendor Info', icon: 'fa-hard-hat' },
          { label: 'Name', value: v.name },
          { label: 'ID', value: String(v.id) },
          { label: 'Category', value: vCat },
          { label: 'Trade Category', value: getVendorTradeCategory(v) },
          { label: 'Type (API)', value: v.vendorType || '\u2014' },
          { label: 'Trades', value: v.trades || '\u2014' },
          { section: 'Contact', icon: 'fa-phone' },
          { label: 'Phone', value: v.phone || '\u2014' },
          { label: 'Email', value: v.email || '\u2014' },
          { label: 'Address', value: v.address || '\u2014' },
          { section: 'Compliance', icon: 'fa-shield-alt' },
          { label: 'Status', value: compText },
          { label: 'Liability Ins. Exp.', value: v.insurance || '\u2014' },
          { label: 'Auto Ins. Exp.', value: v.autoInsurance || '\u2014' },
          { label: 'Workers Comp Exp.', value: v.workersComp || '\u2014' },
          { label: 'Do Not Use', value: v.doNotUse ? 'YES' : 'No' }
        ], appfolioUrl('vendor', v.id));
      }
    });
    vendGrid.addEventListener('change', function(e) {
      var sel = e.target.closest('.vendor-cat-select');
      if (sel) {
        e.stopPropagation();
        var vid = sel.getAttribute('data-vid');
        saveVendorOverride(vid, { category: sel.value }).then(function() {
          renderVendors($('#vendorSearch') ? $('#vendorSearch').value : '');
          showToast('Category \u2192 ' + sel.value);
        });
        return;
      }

      var tradeSel = e.target.closest('.vendor-trade-cat-select');
      if (tradeSel) {
        e.stopPropagation();
        var tvid = tradeSel.getAttribute('data-vid');
        saveVendorOverride(tvid, { tradeCategory: tradeSel.value }).then(function() {
          renderVendors($('#vendorSearch') ? $('#vendorSearch').value : '');
          showToast('Trade category \u2192 ' + tradeSel.value);
        });
      }
    });
  })();

  // Webhook data table — expand/collapse buttons
  (function() {
    var whBody = $('#whDataBody');
    if (whBody) whBody.addEventListener('click', function(e) {
      var btn = e.target.closest('[data-whexpand]');
      if (btn) {
        e.stopPropagation();
        var whId = btn.getAttribute('data-whexpand');
        var detailRow = document.getElementById('whDetail_' + whId);
        if (detailRow) detailRow.classList.toggle('hidden');
      }
    });
  })();

  // Template grid — copy + edit buttons
  (function() {
    var tGrid = $('#templateGrid');
    if (tGrid) tGrid.addEventListener('click', function(e) {
      if (e.target.closest('[data-tcopy]')) { showToast('Template copied to clipboard'); }
      if (e.target.closest('[data-tedit]')) { showToast('Edit mode \u2014 modify template variables'); }
    });
  })();

  // Navigation tabs (with lazy loading for Vendors & Inspections)
  $$('.nav-tab').forEach(function(tab) {
    tab.addEventListener('click', async function() {
      var tabName = tab.getAttribute('data-tab');
      if (!isTabAllowedForRole(tabName)) return;
      $$('.nav-tab').forEach(function(t) { t.classList.remove('active'); });
      tab.classList.add('active');
      $$('.section').forEach(function(s) { s.classList.remove('active'); });
      var sec = document.getElementById('sec-' + tabName);
      if (sec) sec.classList.add('active');
      if (document.body && document.body.classList.contains('tv-mode') && tabName !== 'dashboard') {
        applyTvMode(false);
      }
      syncTvModeScope();

      // DOM cleanup — free heavy inner HTML of deactivated tabs to reduce memory
      // (they'll re-render when next activated)
      var heavyTabs = {
        workorders: '#kanbanBoard',
        vendors: '#vendorGrid',
        inspections: '#inspBody',
        webhooks: '#whDataBody'
      };
      Object.keys(heavyTabs).forEach(function(tName) {
        if (tName !== tabName) {
          var el = document.querySelector(heavyTabs[tName]);
          if (el && el.children.length > 200) {
            el.innerHTML = '<div style="padding:20px;text-align:center;color:var(--text-muted)"><i class="fas fa-redo"></i> Tab paused \u2014 click to reload</div>';
            el._needsRerender = true;
          }
        }
      });
      // Re-render if returning to a tab that was cleaned (DOM cleanup above)
      var activeHeavy = heavyTabs[tabName];
      if (activeHeavy) {
        var el = document.querySelector(activeHeavy);
        if (el && el._needsRerender) {
          el._needsRerender = false;
          if (tabName === 'workorders') renderWorkOrders();
          else if (tabName === 'vendors') renderVendors($('#vendorSearch') ? $('#vendorSearch').value : '');
          else if (tabName === 'inspections') renderInspections($('#inspSearch') ? $('#inspSearch').value : '');
          else if (tabName === 'webhooks') loadWebhookData();
        }
      }
      // Re-render if tab was dirtied by a group filter change while inactive
      if (_groupFilterDirty[tabName]) {
        _groupFilterDirty[tabName] = false;
        if (typeof _tabRenderMap !== 'undefined' && _tabRenderMap[tabName]) {
          try { _tabRenderMap[tabName](); } catch (e) { /* safe */ }
        }
      }

      // Auto-load AP bills when switching to Work Orders tab (enables WO Close Assist)
      if (tabName === 'workorders' && (!BILLS || BILLS.length === 0)) {
        fetchBills(365).then(function() { renderWOCloseAssist(); }).catch(function() {});
      }

      // Lazy-load Vendors on first tab click
      if (tabName === 'vendors' && !_vendorsLazyLoaded && VENDORS.length === 0) {
        _vendorsLazyLoaded = true;
        var vendGrid = $('#vendorGrid');
        if (vendGrid) vendGrid.innerHTML = loadingHtml('Loading vendors\u2026');
        try {
          var ok = await fetchVendors();
          if (ok) {
            renderVendors($('#vendorSearch') ? $('#vendorSearch').value : '');
            populateDropdowns();
            await saveAllToCache();
            showToast('Vendors loaded \u2014 ' + VENDORS.length);
          }
        } catch (e) { showToast('Vendor load failed: ' + (e.message || e)); }
      }

      // Lazy-load Inspections on first tab click
      if (tabName === 'inspections' && !_inspLazyLoaded && INSPECTIONS.length === 0) {
        _inspLazyLoaded = true;
        var inspList = $('#inspList');
        if (inspList) inspList.innerHTML = loadingHtml('Loading inspections\u2026');
        try {
          var inspOk = await fetchInspections();
          if (inspOk) {
            renderInspections($('#inspSearch') ? $('#inspSearch').value : '');
            renderActivityFeed();
            await saveAllToCache();
            showToast('Inspections loaded \u2014 ' + INSPECTIONS.length);
          }
        } catch (e) { showToast('Inspection load failed: ' + (e.message || e)); }
      }

      // Lazy-load Webhook Data on first tab click
      if (tabName === 'webhooks' && !_whLazyLoaded) {
        _whLazyLoaded = true;
        loadWebhookData();
      }
    });
  });

  // WO status filter buttons
  $$('[data-filter]').forEach(function(btn) {
    btn.addEventListener('click', function() {
      $$('[data-filter]').forEach(function(b) { b.classList.remove('active'); });
      btn.classList.add('active');
      currentWOFilter = btn.getAttribute('data-filter');
      renderWorkOrders();
    });
  });

  // WO dropdown filters
  $('#woSearch').addEventListener('input', debounce(function() { renderWorkOrders(); }, CONFIG.DEBOUNCE_MS));
  if ($('#woPriorityFilter')) {
    $('#woPriorityFilter').addEventListener('change', function() { currentWOPriority = this.value; renderWorkOrders(); });
  }
  if ($('#woTypeFilter')) {
    $('#woTypeFilter').addEventListener('change', function() { currentWOType = this.value; renderWorkOrders(); });
  }
  if ($('#woPropertyFilter')) {
    $('#woPropertyFilter').addEventListener('change', function() { currentWOProperty = this.value; renderWorkOrders(); });
  }
  if ($('#woCloseAge')) {
    $('#woCloseAge').addEventListener('change', function() {
      currentWOCloseAssistAge = parseInt(this.value || '14', 10) || 14;
      renderWOCloseAssist();
    });
  }
  if ($('#btnRefreshCloseAssist')) {
    $('#btnRefreshCloseAssist').addEventListener('click', async function() {
      if (_billsLoading) return;
      _billsLoading = true;
      this.disabled = true;
      this.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Loading AP\u2026';
      try {
        var ok = await fetchBills(365);
        renderWOCloseAssist();
        showToast(ok ? ('AP loaded — ' + BILLS.length + ' bills') : 'Could not load AP bills', ok ? { kind: 'success' } : { kind: 'warning' });
      } finally {
        _billsLoading = false;
        this.disabled = false;
        this.innerHTML = '<i class="fas fa-receipt"></i> Refresh AP';
      }
    });
  }
  // WO group filter wired below with global sync

  // Payroll navigation
  if ($('#payrollPrev')) {
    $('#payrollPrev').addEventListener('click', function() { PAYROLL_WEEK_OFFSET--; renderPayroll(); });
  }
  if ($('#payrollNext')) {
    $('#payrollNext').addEventListener('click', function() { PAYROLL_WEEK_OFFSET++; renderPayroll(); });
  }

  // Turn pipeline controls
  if ($('#turnPipeFilter')) {
    $('#turnPipeFilter').addEventListener('change', function() {
      currentTurnPipeFilter = this.value;
      renderTurnPipelineUI();
    });
  }
  if ($('#turnPipeGroup')) {
    $('#turnPipeGroup').addEventListener('change', function() {
      currentTurnPipeGroup = this.value;
      renderTurnPipelineUI();
    });
  }
  if ($('#turnPipeSearch')) {
    $('#turnPipeSearch').addEventListener('input', debounce(function() {
      renderTurnPipelineUI();
    }, CONFIG.DEBOUNCE_MS));
  }
  if ($('#dashTurnPmFilter')) {
    $('#dashTurnPmFilter').addEventListener('change', function() {
      DASH_TURN_PM_FILTER = this.value;
      DASH_TURN_PAGE = 0;
      renderTurnDashboardStrip();
    });
  }
  if ($('#dashTurnPrev')) {
    $('#dashTurnPrev').addEventListener('click', function() {
      var totalPages = Math.max(1, Math.ceil(getDashboardTurnEntries().length / getDashboardTurnPageSize()));
      DASH_TURN_PAGE = (DASH_TURN_PAGE - 1 + totalPages) % totalPages;
      renderTurnDashboardStrip();
    });
  }
  if ($('#dashTurnNext')) {
    $('#dashTurnNext').addEventListener('click', function() {
      var totalPages = Math.max(1, Math.ceil(getDashboardTurnEntries().length / getDashboardTurnPageSize()));
      DASH_TURN_PAGE = (DASH_TURN_PAGE + 1) % totalPages;
      renderTurnDashboardStrip();
    });
  }
  if ($('#dashTvMode')) {
    $('#dashTvMode').addEventListener('click', function() {
      var isEnabled = document.body && document.body.classList.contains('tv-mode');
      applyTvMode(!isEnabled);
    });
    try {
      if (localStorage.getItem('hm_tv_mode') === '1') applyTvMode(true);
    } catch (e) { /* */ }
  }
  document.addEventListener('keydown', function(e) {
    if (e.key !== 'Escape' && e.key !== 'Esc') return;
    if (document.body && document.body.classList.contains('tv-mode')) {
      applyTvMode(false);
      showToast('TV mode disabled — full navigation restored');
      return;
    }
    ensureNavigationChromeVisible();
  });
  if ($('#dashTurnStrip')) {
    $('#dashTurnStrip').addEventListener('click', function(e) {
      var card = e.target.closest('[data-turndash-open]');
      if (!card) return;
      openTurnBoardDetail(card.getAttribute('data-turndash-open'));
    });
  }

  // Inspection status filter
  if ($('#inspStatusFilter')) {
    $('#inspStatusFilter').addEventListener('change', function() {
      renderInspections($('#inspSearch') ? $('#inspSearch').value : '');
    });
  }

  // Clickable KPI cards
  $$('.kpi-clickable[data-kpi]').forEach(function(card) {
    card.addEventListener('click', function() {
      var kpi = this.getAttribute('data-kpi');
      if (kpi === 'open' || kpi === 'urgent' || kpi === 'flagged') {
        // Switch to WO tab and filter
        $$('.nav-tab').forEach(function(t) { t.classList.remove('active'); });
        var woTab = document.querySelector('[data-tab="workorders"]');
        if (woTab) woTab.classList.add('active');
        $$('.section').forEach(function(s) { s.classList.remove('active'); });
        var woSec = document.getElementById('sec-workorders');
        if (woSec) woSec.classList.add('active');
        if (kpi === 'urgent') {
          currentWOPriority = 'Urgent';
          if ($('#woPriorityFilter')) $('#woPriorityFilter').value = 'Urgent';
        } else if (kpi === 'flagged') {
          currentWOFilter = 'flagged';
          $$('[data-filter]').forEach(function(b) {
            b.classList.toggle('active', b.getAttribute('data-filter') === 'flagged');
          });
        }
        renderWorkOrders();
      } else if (kpi === 'turns') {
        $$('.nav-tab').forEach(function(t) { t.classList.remove('active'); });
        var tTab = document.querySelector('[data-tab="turnboard"]');
        if (tTab) tTab.classList.add('active');
        $$('.section').forEach(function(s) { s.classList.remove('active'); });
        var tSec = document.getElementById('sec-turnboard');
        if (tSec) tSec.classList.add('active');
      } else if (kpi === 'moveouts') {
        // Scroll to move-out section on dashboard
        var moSec = document.getElementById('moveOutSection');
        if (moSec) moSec.scrollIntoView({ behavior: 'smooth' });
      }
    });
  });

  // Activity feed filters
  $$('[data-actfilter]').forEach(function(btn) {
    btn.addEventListener('click', function() {
      $$('[data-actfilter]').forEach(function(b) { b.classList.remove('active'); });
      btn.classList.add('active');
      currentActivityFilter = btn.getAttribute('data-actfilter');
      renderActivityFeed();
    });
  });

  // Item Detail modal close
  if ($('#itemDetailClose')) {
    $('#itemDetailClose').addEventListener('click', function() { closeModal('itemDetailModal'); });
  }
  if ($('#itemDetailCloseBtn')) {
    $('#itemDetailCloseBtn').addEventListener('click', function() { closeModal('itemDetailModal'); });
  }

  // What's New modal wiring
  if ($('#whatsNewClose')) $('#whatsNewClose').addEventListener('click', dismissWhatsNew);
  if ($('#whatsNewDismiss')) $('#whatsNewDismiss').addEventListener('click', dismissWhatsNew);

  /* btnLoadBills removed — billing stripped */
  $('#vendorSearch').addEventListener('input', debounce(function() { renderVendors(this.value); }, isConstrainedDevice() ? 450 : CONFIG.DEBOUNCE_MS));
  if ($('#vendorCategoryFilter')) {
    $('#vendorCategoryFilter').addEventListener('change', function() { renderVendors($('#vendorSearch') ? $('#vendorSearch').value : ''); });
  }
  if ($('#vendorComplianceFilter')) {
    $('#vendorComplianceFilter').addEventListener('change', function() { renderVendors($('#vendorSearch') ? $('#vendorSearch').value : ''); });
  }
  if ($('#vendorTradeFilter')) {
    $('#vendorTradeFilter').addEventListener('change', function() { renderVendors($('#vendorSearch') ? $('#vendorSearch').value : ''); });
  }
  if ($('#vendorSortMode')) {
    $('#vendorSortMode').addEventListener('change', function() { renderVendors($('#vendorSearch') ? $('#vendorSearch').value : ''); });
  }
  document.addEventListener('visibilitychange', function() {
    if (!document.hidden && _vendorsNeedRender) {
      renderVendors($('#vendorSearch') ? $('#vendorSearch').value : '');
    }
  });
  if ($('#nwoVendorSearch')) {
    $('#nwoVendorSearch').addEventListener('input', debounce(function() {
      renderNewWOVendorOptions(this.value);
    }, CONFIG.DEBOUNCE_MS));
  }
  $('#btnNewWO').addEventListener('click', function() {
    renderNewWOVendorOptions($('#nwoVendorSearch') ? $('#nwoVendorSearch').value : '');
    openModal('newWOModal');
  });
  $('#btnNewTemplate').addEventListener('click', function() { showToast('Template editor \u2014 define trigger, variables, and body'); });
  $('#btnRefreshTurns').addEventListener('click', function() { sectionRefresh('turns', this); });
  $('#btnClearErrors').addEventListener('click', function() {
    API_ERRORS = API_ERRORS.filter(function(e) { return e.action !== 'resolved'; });
    renderErrorLog();
    showToast('Cleared resolved errors');
  });

  // Refresh button — force full reload from API
  $('#refreshBtn').addEventListener('click', function() { refreshData(); });

  // Progress dock close
  $('#progClose').addEventListener('click', function() { $('#progressDock').classList.add('hidden'); });

  // Inspections search
  $('#inspSearch').addEventListener('input', debounce(function() { renderInspections(this.value); }, CONFIG.DEBOUNCE_MS));
  $('#btnRefreshInsp').addEventListener('click', function() { sectionRefresh('inspections', this); });

  // Per-section refresh buttons
  $('#btnRefreshDash').addEventListener('click', function() { sectionRefresh('dashboard', this); });
  $('#btnRefreshWO').addEventListener('click', function() { sectionRefresh('workorders', this); });
  $('#btnRefreshVendors').addEventListener('click', function() { sectionRefresh('vendors', this); });

  // Theme toggle
  $('#themeToggle').addEventListener('click', function() { toggleTheme(); });
  updateThemeIcon(); // sync icon with initial state

  // Cache export / import
  $('#btnExportCache').addEventListener('click', function() { exportCacheToJSON(); });
  $('#btnImportCache').addEventListener('click', function() { $('#cacheFileInput').click(); });
  $('#cacheFileInput').addEventListener('change', function() {
    if (this.files && this.files[0]) {
      importCacheFromJSON(this.files[0]);
      this.value = ''; // reset so same file can be re-imported
    }
  });

  // (Load Groups button moved to global filter bar — wired below)

  // Render vendors when user navigates back to that tab if a refresh was deferred.
  document.body.addEventListener('click', function(e) {
    var tab = e.target.closest('.nav-tab[data-tab="vendors"]');
    if (tab && _vendorsNeedRender) {
      setTimeout(function() {
        renderVendors($('#vendorSearch') ? $('#vendorSearch').value : '');
      }, 0);
    }
  });

  // ---- Global property group filter (single dropdown in sticky header) ----
  // Map tab names to their render functions (called on-demand, not all at once)
  var _tabRenderMap = {
    dashboard: function() { renderDashboardKPIs(); },
    workorders: function() { renderWorkOrders(); },
    payroll: function() { renderPayroll(); },
    turnboard: function() { renderTurnPipelineUI(); },
    inspections: function() { renderInspections($('#inspSearch') ? $('#inspSearch').value : ''); },
    vendors: function() { renderVendors($('#vendorSearch') ? $('#vendorSearch').value : ''); }
  };

  // Optimized: only re-render the active tab + KPIs, defer via requestAnimationFrame
  // Other tabs marked dirty and re-rendered lazily when switched to
  function applyGroupFilterChange() {
    var activeTab = document.querySelector('.nav-tab.active');
    var tabName = activeTab ? activeTab.getAttribute('data-tab') : 'dashboard';
    // Mark all renderable tabs as dirty
    Object.keys(_tabRenderMap).forEach(function(t) { _groupFilterDirty[t] = true; });
    // Immediately clear + render the active tab (deferred one frame to unblock the event)
    _groupFilterDirty[tabName] = false;
    requestAnimationFrame(function() {
      try {
        if (_tabRenderMap[tabName]) _tabRenderMap[tabName]();
      } catch (e) { /* safe */ }
      // Dashboard KPIs always refresh regardless of active tab
      if (tabName !== 'dashboard') {
        try { renderDashboardKPIs(); } catch (e) { /* safe */ }
      }
    });
    updateGlobalGroupIndicator();
  }

  function clearPropertyGroupFilters() {
    currentPropertyGroup = '';
    _isInGroupMissLogCount = 0;
    if (globalGrpEl) globalGrpEl.value = '';
    currentTurnPipeGroup = '';
    var turnGroupEl = $('#turnPipeGroup');
    if (turnGroupEl) turnGroupEl.value = '';
    applyGroupFilterChange();
  }

  // Wire up the single global group filter dropdown
  var globalGrpEl = $('#globalGroupFilter');
  if (globalGrpEl) {
    globalGrpEl.addEventListener('change', function() {
      currentPropertyGroup = this.value;
      _isInGroupMissLogCount = 0; // reset miss log throttle on filter change
      applyGroupFilterChange();
    });
  }

  // Wire up the global group clear button
  var globalClearBtn = $('#globalGroupClear');
  if (globalClearBtn) {
    globalClearBtn.addEventListener('click', function() {
      clearPropertyGroupFilters();
    });
  }

  // Wire up the "Reload Groups" button in the global filter bar
  // Shift+Click shows detailed diagnostics toast for debugging
  var btnGlobalLoadGroups = $('#btnGlobalLoadGroups');
  if (btnGlobalLoadGroups) {
    btnGlobalLoadGroups.addEventListener('click', async function(evt) {
      var wantDiag = evt && evt.shiftKey;
      var wantCopyCommand = evt && (evt.altKey || evt.ctrlKey || evt.metaKey);
      btnGlobalLoadGroups.disabled = true;
      btnGlobalLoadGroups.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Loading\u2026';
      try {
        var ok = await fetchPropertyGroups();
        if (ok) {
          populateGroupFilters();
          if (wantCopyCommand && navigator.clipboard) {
            var cmd = buildPropertyGroupsProxyCurl(API_PROXY);
            if (cmd) {
              await navigator.clipboard.writeText(cmd);
              showToast('Copied property-groups proxy curl command');
            }
          }
          if (wantDiag && window._pgDiag) {
            var d = window._pgDiag;
            var lines = [
              'Groups: ' + PROPERTY_GROUPS.length,
              'UUID map: ' + d.uuidMapSize + ' entries',
              'DB names: ' + d.dbNameCount,
              'UUID hits/misses: ' + d.uuidHits + '/' + d.uuidMisses,
              'Name bridges: ' + d.nameMatches + '/' + PROPERTIES.length,
              'ID map: ' + d.idMatches + ' props',
              'Portfolio matches: ' + d.portfolioMatches,
              'nameMap: ' + Object.keys(_nameToGroups).length,
              'idMap: ' + Object.keys(_idToGroups).length
            ];
            if (d.errors.length > 0) lines.push('ERRORS: ' + d.errors.join('; '));
            showToast(lines.join(' | '), 12000);
          } else {
            showToast('Property groups refreshed \u2014 ' + PROPERTY_GROUPS.length + ' groups, ' +
              Object.keys(_nameToGroups).length + ' name mappings, ' +
              Object.keys(_idToGroups).length + ' ID mappings');
          }
        } else {
          showToast('Failed to load property groups \u2014 check console for [PG] logs');
        }
      } catch (e) {
        showToast('Error: ' + (e.message || e));
      } finally {
        btnGlobalLoadGroups.disabled = false;
        btnGlobalLoadGroups.innerHTML = '<i class="fas fa-sync-alt"></i> Reload Groups';
      }
    });
  }

  // CORS banner toggle (collapsible)
  if ($('#corsBannerToggle')) {
    $('#corsBannerToggle').addEventListener('click', function() {
      var body = $('#corsBannerBody');
      var icon = this.querySelector('.cors-toggle-icon');
      if (body.style.display === 'none') {
        body.style.display = 'block';
        if (icon) icon.classList.add('open');
      } else {
        body.style.display = 'none';
        if (icon) icon.classList.remove('open');
      }
    });
  }

  // Webhook modal open/close
  if ($('#btnWebhookConfig')) {
    $('#btnWebhookConfig').addEventListener('click', function() {
      var urlEl = $('#webhookUrl');
      if (urlEl) urlEl.value = API_PROXY + '?action=webhook';
      openModal('webhookModal');
      renderWebhookEventList();
    });
  }
  if ($('#webhookModalClose')) {
    $('#webhookModalClose').addEventListener('click', function() { closeModal('webhookModal'); });
  }
  if ($('#webhookModalCloseBtn')) {
    $('#webhookModalCloseBtn').addEventListener('click', function() { closeModal('webhookModal'); });
  }
  if ($('#btnCopyWebhook')) {
    $('#btnCopyWebhook').addEventListener('click', function() {
      var url = $('#webhookUrl').value;
      if (navigator.clipboard) {
        navigator.clipboard.writeText(url).then(function() { showToast('Webhook URL copied'); });
      } else {
        showToast('Clipboard not available');
      }
    });
  }
  if ($('#btnWebhookPoll')) {
    $('#btnWebhookPoll').addEventListener('click', async function() {
      var btn = this;
      btn.disabled = true;
      btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Polling\u2026';
      try {
        await pollWebhookEvents();
        renderWebhookEventList();
        renderActivityFeed();
        showToast('Polled ' + WEBHOOK_EVENTS.length + ' webhook events');
      } catch (err) {
        showToast('Webhook poll failed: ' + (err.message || err));
      } finally {
        btn.disabled = false;
        btn.innerHTML = '<i class="fas fa-sync-alt"></i> Poll Now';
      }
    });
  }
  if ($('#webhookPollInterval')) {
    $('#webhookPollInterval').addEventListener('change', function() {
      setupWebhookAutoPoll(parseInt(this.value, 10) || 0);
    });
  }

  // Webhook Data Review — event listeners
  var _whSearchTimer = null;
  if ($('#whSearch')) {
    $('#whSearch').addEventListener('input', function() {
      clearTimeout(_whSearchTimer);
      _whSearchTimer = setTimeout(function() {
        _whFilters.search = $('#whSearch').value.trim();
        _whPage = 0;
        loadWebhookData();
      }, 400);
    });
  }
  if ($('#whTypeFilter')) {
    $('#whTypeFilter').addEventListener('change', function() {
      _whFilters.type = this.value;
      _whPage = 0;
      loadWebhookData();
    });
  }
  if ($('#whSourceFilter')) {
    $('#whSourceFilter').addEventListener('change', function() {
      _whFilters.source = this.value;
      _whPage = 0;
      loadWebhookData();
    });
  }
  if ($('#whDateFrom')) {
    $('#whDateFrom').addEventListener('change', function() {
      _whFilters.from = this.value;
      _whPage = 0;
      loadWebhookData();
    });
  }
  if ($('#whDateTo')) {
    $('#whDateTo').addEventListener('change', function() {
      _whFilters.to = this.value;
      _whPage = 0;
      loadWebhookData();
    });
  }
  if ($('#btnWhClear')) {
    $('#btnWhClear').addEventListener('click', function() {
      _whFilters = { search: '', type: '', source: '', from: '', to: '' };
      _whPage = 0;
      if ($('#whSearch')) $('#whSearch').value = '';
      if ($('#whTypeFilter')) $('#whTypeFilter').value = '';
      if ($('#whSourceFilter')) $('#whSourceFilter').value = '';
      if ($('#whDateFrom')) $('#whDateFrom').value = '';
      if ($('#whDateTo')) $('#whDateTo').value = '';
      loadWebhookData();
      showToast('Webhook filters cleared');
    });
  }
  if ($('#btnWhRefresh')) {
    $('#btnWhRefresh').addEventListener('click', function() { loadWebhookData(); });
  }
  if ($('#btnWhStats')) {
    $('#btnWhStats').addEventListener('click', function() {
      var panel = $('#whStatsPanel');
      if (panel && !panel.classList.contains('hidden')) {
        panel.classList.add('hidden');
      } else {
        loadWebhookStats();
      }
    });
  }
  if ($('#btnWhMigrate')) {
    $('#btnWhMigrate').addEventListener('click', function() { migrateWebhookBlob(); });
  }
  if ($('#btnWhPrev')) {
    $('#btnWhPrev').addEventListener('click', function() {
      if (_whPage > 0) { _whPage--; loadWebhookData(); }
    });
  }
  if ($('#btnWhNext')) {
    $('#btnWhNext').addEventListener('click', function() {
      if ((_whPage + 1) * _whPageSize < _whTotal) { _whPage++; loadWebhookData(); }
    });
  }

  // Create WO handler
  $('#btnCreateWO').addEventListener('click', async function() {
    var desc = $('#nwoDesc').value.trim();
    if (!desc) { showToast('Please enter a description for the work order.'); return; }

    var btn = this;
    btn.disabled = true;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Creating\u2026';

    try {
      var payload = {
        JobDescription: desc,
        Priority: $('#nwoPriority').value
      };
      var propVal = $('#nwoProperty').value;
      if (propVal) { payload.PropertyId = propVal; }
      var unitVal = $('#nwoUnit').value.trim();
      if (unitVal) { payload.UnitId = unitVal; }
      var vendVal = $('#nwoVendor').value;
      if (vendVal) { payload.VendorId = vendVal; }

      await apiFetch('/api/v0/work_orders', {
        method: 'POST',
        body: JSON.stringify(payload)
      });

      closeModal('newWOModal');
      showToast('Work order created \u2014 refreshing list\u2026');
      await fetchWorkOrders();
      renderWorkOrders();
      renderDashboardKPIs();
      await saveAllToCache();
    } catch (err) {
      showToast('Create failed: ' + err.message);
    } finally {
      btn.disabled = false;
      btn.innerHTML = '<i class="fas fa-plus"></i> Create';
    }
  });
}

// Per-section refresh — reload just one dataset
async function sectionRefresh(section, btn) {
  if (btn.disabled) return;
  btn.disabled = true;
  btn.classList.add('spinning');
  try {
    if (section === 'workorders' || section === 'dashboard') {
      showToast('Refreshing open work orders\u2026');
      await fetchWorkOrders();
      if (!_billsLoadedAt || (Date.now() - _billsLoadedAt) > (30 * 60 * 1000)) {
        await fetchBills(365);
      }
      renderWorkOrders();
    }
    if (section === 'vendors' || section === 'dashboard') {
      showToast('Refreshing vendors\u2026');
      await fetchVendors();
      renderVendors($('#vendorSearch') ? $('#vendorSearch').value : '');
    }
    if (section === 'turns' || section === 'dashboard') {
      showToast('Refreshing turn data\u2026');
      await fetchTurns();
      renderTurnBoard();
    }
    if (section === 'inspections' || section === 'dashboard') {
      showToast('Refreshing inspections\u2026');
      await fetchInspections();
      renderInspections($('#inspSearch') ? $('#inspSearch').value : '');
    }
    renderDashboardKPIs();
    renderActivityFeed();
    populateDropdowns();
    await saveAllToCache();
    showToast('Section refreshed');
  } catch (err) {
    showToast('Refresh failed: ' + (err.message || err));
  } finally {
    btn.disabled = false;
    btn.classList.remove('spinning');
  }
}

async function initApp() {
  if (appInitialized) return;
  appInitialized = true;
  applyVersionBadge('');

  // ================================================================
  // VENDOR-ONLY MODE: streamlined init — only load vendor data
  // ================================================================
  if (_accessRole === 'vendors') {
    setApiStatus('loading', 'Initializing vendor view\u2026');
    updateCacheBadge('loading');
    await loadVendorOverrides();

    if ($('#vendorGrid')) $('#vendorGrid').innerHTML = loadingHtml('Loading vendors\u2026');
    wireUpUI();

    // Try cached vendors first
    try {
      var vCached = await cacheGet('vendors');
      if (vCached && Array.isArray(vCached.data) && vCached.data.length > 0) {
        VENDORS = vCached.data;
        renderVendors($('#vendorSearch') ? $('#vendorSearch').value : '');
        showToast('Loaded ' + VENDORS.length + ' vendors from cache');
      }
    } catch (e) { /* no cache */ }

    // Skip ping — just fetch vendors directly (vendors only need Reports API)
    try {
      setApiStatus('loading', 'Loading vendors\u2026');
      _vendorsLazyLoaded = true;
      var vOk = await fetchVendors();
      // Retry once if first attempt fails (Val Town cold start / transient timeout)
      if (!vOk) {
        setApiStatus('loading', 'Retrying vendors\u2026');
        await sleep(2000);
        vOk = await fetchVendors();
      }
      if (vOk) {
        renderVendors($('#vendorSearch') ? $('#vendorSearch').value : '');
        setApiStatus('', 'Vendor Access [' + APP_VERSION + '] \u2014 ' + VENDORS.length + ' vendors');
        $('#apiStatus').className = 'topbar-status';
        showToast('Vendors loaded \u2014 ' + VENDORS.length);
        await saveAllToCache();
      } else {
        setApiStatus('error', 'Vendor fetch failed \u2014 check proxy');
        if (VENDORS.length === 0) showToast('Could not load vendors from proxy');
      }
    } catch (e) {
      setApiStatus('error', 'Connection failed');
      if (VENDORS.length === 0) showToast('Connection failed \u2014 ' + (e.message || e));
    }
    updateCacheBadge(VENDORS.length > 0 ? 'cached' : 'offline');
    return;
  }

  // ================================================================
  // FULL ACCESS MODE: normal init flow
  // ================================================================
  setApiStatus('loading', 'Initializing\u2026');
  updateCacheBadge('loading');

  // Load flags + vendor overrides from IndexedDB
  await loadFlags();
  await loadVendorOverrides();

  // Show skeleton loading states
  if ($('#kanbanBoard')) $('#kanbanBoard').innerHTML = loadingHtml('Checking cache\u2026');
  if ($('#vendorGrid')) $('#vendorGrid').innerHTML = loadingHtml('Checking cache\u2026');
  if ($('#turnPipeline')) $('#turnPipeline').innerHTML = loadingHtml('Checking cache\u2026');
  if ($('#inspBody')) $('#inspBody').innerHTML = '<tr><td colspan="8">' + loadingHtml('Checking cache\u2026') + '</td></tr>';
  renderTemplates();
  renderErrorLog();
  wireUpUI();

  // ================================================================
  // STEP 1: Check IndexedDB cache for fresh work order data
  // ================================================================
  var cacheLoaded = false;
  try {
    var cachedWO = await cacheGet('work_orders');
    if (cachedWO && Array.isArray(cachedWO.data) && cachedWO.data.length > 0) {
      var fresh = isCacheFresh(cachedWO);
      var cachedVendors = await cacheGet('vendors');
      var cachedProps = await cacheGet('properties');
      var cachedTurns = await cacheGet('turns');
      var cachedInsp = await cacheGet('inspections');

      WORK_ORDERS = cachedWO.data;
      VENDORS = (cachedVendors && cachedVendors.data) ? cachedVendors.data : [];
      PROPERTIES = (cachedProps && cachedProps.data) ? cachedProps.data : [];
      TURNS = (cachedTurns && cachedTurns.data) ? cachedTurns.data : [];
      INSPECTIONS = (cachedInsp && cachedInsp.data) ? cachedInsp.data : [];

      cacheLoaded = true;
      updateCacheBadge('cached', cachedWO.timestamp, !fresh);
      setApiStatus('', (fresh ? 'Cached' : 'Stale cache') + ' \u2014 ' + WORK_ORDERS.length + ' WOs');
      if (!fresh) { setApiStatus('loading', 'Stale cache loaded \u2014 refreshing\u2026'); }
      showToast((fresh ? 'Loaded from cache' : 'Loaded stale cache') + ' \u2014 ' + WORK_ORDERS.length + ' work orders, ' + VENDORS.length + ' vendors');

      // Render immediately with cached data
      renderAll();

      if (fresh) {
        // Cache is fresh — no need to fetch now. Schedule background refresh.
        setApiStatus('', 'Connected (cached)');
        $('#apiStatus').className = 'topbar-status';
        return;
      }
      // Cache is stale — fall through to fetch fresh data
    }
  } catch (cacheErr) {
    console.log('Cache read failed: ' + (cacheErr.message || cacheErr));
  }

  // ================================================================
  // STEP 2: Pre-flight = lightweight API test (Database API, 5 records)
  //         Tests connectivity before launching Reports API bulk load
  // ================================================================
  setApiStatus('loading', 'Connecting to AppFolio\u2026');
  if (!cacheLoaded) {
    if ($('#kanbanBoard')) $('#kanbanBoard').innerHTML = loadingHtml('Testing API connection\u2026');
  }

  // ================================================================
  // STEP 2: Pre-flight = Proxy ?action=ping
  //         Tests proxy connectivity + AppFolio auth in one shot
  // ================================================================
  try {
    setApiStatus('loading', 'Pinging proxy\u2026');
    var pingData = await proxyAction('ping');

    // Detect proxy version when the proxy includes a version field
    _proxyVersion = pingData.version || 'v7';
    if (enforceServerVersionGuard(pingData)) return;

    // Check APIs individually — be lenient: continue loading if Reports API is reachable
    var rptOk = pingData.reports_api ? pingData.reports_api.ok : pingData.ok;
    var dbOk = pingData.db_api ? pingData.db_api.ok : true;

    if (!pingData.ok && !rptOk) {
      // BOTH APIs failed — block loading
      var dbSt = (pingData.db_api && pingData.db_api.status) || pingData.status || 0;
      var rptSt = (pingData.reports_api && pingData.reports_api.status) || 0;
      var detail = 'DB:' + dbSt + ' Reports:' + rptSt;
      logApiError(dbSt, 'Pre-flight: Both APIs unreachable \u2014 ' + detail, 'resolved');
      showCorsError('Pre-flight ping failed (' + detail + '). Verify proxy has correct credentials and both domains are accessible.');
      setApiStatus('error', 'Auth Failed (' + detail + ')');
      if (!cacheLoaded) {
        await loadStaleCache();
        renderAll();
      }
      updateCacheBadge(cacheLoaded ? 'cached' : 'offline', null, true);
      return;
    }
    // At least Reports API is reachable — continue loading even if DB API is down
    var pingMsg = 'Proxy OK';
    if (pingData.latency_ms) pingMsg += ' (' + pingData.latency_ms + 'ms)';
    if (_proxyVersion !== 'v7') pingMsg += ' [' + _proxyVersion + ']';
    if (!dbOk) {
      pingMsg += ' (DB API down, Reports OK)';
      logApiError(0, 'DB API v0 unreachable — some features may be limited', 'resolved');
    }
    pingMsg += ' \u2014 loading data\u2026';
    setApiStatus('loading', pingMsg);
    console.log('Proxy version: ' + _proxyVersion + ', DB: ' + (dbOk ? 'ok' : 'down') + ', Reports: ' + (rptOk ? 'ok' : 'down'));
  } catch (preErr) {
    var peMsg = preErr.message || 'Connection failed';
    var isCsp = peMsg.indexOf('Content Security Policy') !== -1 || peMsg.indexOf('CSP') !== -1 || peMsg.indexOf('Refused to connect') !== -1 || preErr.name === 'TypeError';
    var isPingServerError = /action=ping failed: HTTP 5\d\d/i.test(peMsg) || /action=ping: .*5\d\d/i.test(peMsg);
    if (isCsp) {
      logApiError(0, 'Pre-flight BLOCKED: ' + peMsg + '. Click "Allow additional resources" popup, then re-enter credentials.', 'queued');
      showCorsError(peMsg);
      setApiStatus('error', 'CSP Blocked \u2014 Click Allow');
      if (!cacheLoaded) {
        await loadStaleCache();
        renderAll();
      }
      updateCacheBadge(cacheLoaded ? 'cached' : 'offline', null, true);
      return;
    } else if (isPingServerError) {
      logApiError(500, 'Pre-flight ping failed but continuing: ' + peMsg, 'retry');
      setApiStatus('loading', 'Ping failed — continuing with direct dataset fetch…');
      showToast('Proxy ping failed (500). Continuing data fetch anyway…', { kind: 'warning', iconClass: 'fa-triangle-exclamation', duration: 4500 });
    } else {
      logApiError(0, 'Pre-flight failed: ' + peMsg, 'queued');
      setApiStatus('loading', 'Pre-flight warning — trying full data fetch…');
    }
  }

  // ================================================================
  // STEP 3: Full data fetch via proxy action endpoints
  //         Each action does server-side pagination — ONE request per dataset
  // ================================================================
  await fetchAllLive();
}

// ---- Step-level timeout wrapper ----
// Wraps a fetch function with a timeout so no single step can block forever
function withStepTimeout(fn, timeoutMs) {
  timeoutMs = timeoutMs || 60000; // default 60s per step
  return new Promise(function(resolve) {
    var done = false;
    var timer = setTimeout(function() {
      if (!done) { done = true; resolve(false); }
    }, timeoutMs);
    fn().then(function(result) {
      if (!done) { done = true; clearTimeout(timer); resolve(result); }
    }).catch(function() {
      if (!done) { done = true; clearTimeout(timer); resolve(false); }
    });
  });
}

// Fetch all data via proxy action endpoints — ONE request per dataset
// Each ?action= call does server-side pagination and returns complete results
// Every step has a 60-second timeout — NOTHING hangs forever
async function fetchAllLive() {
  var anySuccess = false;
  updateCacheBadge('loading');
  // Vendors & Inspections lazy-loaded on tab click — removed from initial sync
  var steps = ['Work Orders', 'Properties', 'Turns', 'Move-Outs', 'Turn WOs', 'Groups', 'Tasks', 'Turn Tracker'];
  showProgress('Syncing AppFolio (' + DATA_WINDOW_DAYS + 'd)', steps);

  try {
    // Step 0: Work Orders (proxy action — Reports API)
    updateProgress(0, 'active', 'Fetching work orders\u2026');
    var woOk = await withStepTimeout(fetchWorkOrders, 60000);
    updateProgress(0, woOk ? 'done' : 'error', woOk ? WORK_ORDERS.length + ' open work orders' : 'Work orders failed');
    if (woOk) { renderWorkOrders(); renderDashboardKPIs(); renderActivityFeed(); }

    // Step 1: Properties (proxy action — Reports API)
    updateProgress(1, 'active', 'Fetching properties\u2026');
    var propOk = await withStepTimeout(fetchProperties, 60000);
    updateProgress(1, propOk ? 'done' : 'error', propOk ? PROPERTIES.length + ' properties' : 'Properties failed');
    if (propOk) { populateDropdowns(); renderWorkOrders(); }

    // Step 2: Turns — In Progress only, 60-day window (proxy action — Reports API)
    // Short timeout (20s) — turns are supplementary; pipeline works from WOs + move-outs too
    updateProgress(2, 'active', 'Fetching in-progress turns\u2026');
    var turnOk = await withStepTimeout(function() { return fetchTurns(); }, 20000);
    updateProgress(2, turnOk ? 'done' : 'error', turnOk ? TURNS.length + ' turns' : 'Turns skipped (timeout)');
    if (turnOk) { renderTurnBoard(); renderActivityFeed(); }

    // Step 2b: Unit Turns (DB API) — live deposit / scheduling data to augment Reports API turns
    fetchUnitTurnsDB().then(function(ok) { if (ok && turnOk) renderTurnBoard(); }).catch(function(){});

    // Step 3: Upcoming Move-Outs — tenant directory, Notice tenants (proxy action — Reports API)
    updateProgress(3, 'active', 'Fetching upcoming move-outs\u2026');
    var moOk = await withStepTimeout(fetchUpcomingMoveouts, 45000);
    updateProgress(3, moOk ? 'done' : 'error', moOk ? UPCOMING_MOVEOUTS.length + ' upcoming' : 'Move-outs skipped');
    if (moOk) { renderTurnBoard(); renderDashboardKPIs(); }

    // Step 4: Turn Work Orders — DB API v0, Unit Turn type only (real-time status)
    updateProgress(4, 'active', 'Fetching turn work orders\u2026');
    var twoOk = await withStepTimeout(fetchTurnWorkOrders, 20000);
    updateProgress(4, twoOk ? 'done' : 'error', twoOk ? TURN_WORK_ORDERS.length + ' turn WOs' : 'Turn WOs skipped');
    if (twoOk) { renderTurnBoard(); }

    // Step 5: Property Groups (proxy action — DB API v0)
    updateProgress(5, 'active', 'Fetching property groups\u2026');
    var grpOk = await withStepTimeout(fetchPropertyGroups, 45000);
    var grpMsg = grpOk
      ? PROPERTY_GROUPS.length + ' groups, ' + Object.keys(_idToGroups).length + ' ID maps'
      : 'Groups skipped';
    updateProgress(5, grpOk ? 'done' : 'error', grpMsg);
    if (grpOk) { populateDropdowns(); renderWorkOrders(); }

    // Step 6: Recent Tasks (proxy action — DB API v0)
    updateProgress(6, 'active', 'Fetching recent tasks\u2026');
    var taskOk = await withStepTimeout(fetchRecentTasks, 45000);
    updateProgress(6, taskOk ? 'done' : 'error', taskOk ? RECENT_TASKS.length + ' tasks' : 'Tasks skipped');
    if (taskOk) { renderActivityFeed(); }

    // Step 7: Turn Tracker records (proxy blob — persisted stage overrides)
    updateProgress(7, 'active', 'Loading turn tracker\u2026');
    var trkOk = await withStepTimeout(function() {
      return fetchTurnRecords().then(function() { return true; });
    }, 30000);
    updateProgress(7, trkOk ? 'done' : 'error', trkOk ? TURN_RECORDS.length + ' tracked' : 'Tracker skipped');

    // Final re-render: turns with all available correlated data
    renderTurnBoard();

    if (BILLS.length === 0) {
      // Best-effort AP load for WO close-assist; non-blocking and non-fatal.
      try { await withStepTimeout(function() { return fetchBills(365); }, 30000); } catch (e) { /* ignore */ }
    }

    anySuccess = woOk || propOk || turnOk;
  } catch (e) {
    // individual errors already logged
  }

  // Final full render to ensure everything is consistent
  renderAll();
  renderWOCloseAssist();

  if (anySuccess) {
    var summary = 'WO:' + WORK_ORDERS.length + ' P:' + PROPERTIES.length + ' T:' + TURNS.length + ' (V/I lazy)';
    var versionTag = _proxyVersion !== 'v7' ? ' [' + _proxyVersion + ']' : '';
    setApiStatus('', 'Connected' + versionTag + ' \u2014 ' + summary);
    $('#apiStatus').className = 'topbar-status';
    await saveAllToCache();
    updateProgress(-1, '', 'Sync complete \u2014 ' + summary);
    hideProgress();
    // Start webhook auto-poll (default 60s)
    var pollSel = $('#webhookPollInterval');
    var pollInterval = pollSel ? (parseInt(pollSel.value, 10) || 0) : 60;
    if (pollInterval > 0) setupWebhookAutoPoll(pollInterval);
  } else if (API_ERRORS.length > 0) {
    setApiStatus('error', 'API Errors \u2014 Check Log');
    updateCacheBadge('offline');
    updateProgress(-1, '', 'Sync failed');
    hideProgress();
  } else {
    setApiStatus('error', 'No Data Loaded');
    updateCacheBadge('offline');
    hideProgress();
  }
}

// Load stale cache as fallback when API is unreachable
async function loadStaleCache() {
  try {
    var cachedWO = await cacheGet('work_orders');
    if (cachedWO && Array.isArray(cachedWO.data) && cachedWO.data.length > 0) {
      var cachedVendors = await cacheGet('vendors');
      var cachedProps = await cacheGet('properties');
      var cachedTurns = await cacheGet('turns');
      var cachedInsp = await cacheGet('inspections');
      WORK_ORDERS = cachedWO.data;
      VENDORS = (cachedVendors && cachedVendors.data) ? cachedVendors.data : [];
      PROPERTIES = (cachedProps && cachedProps.data) ? cachedProps.data : [];
      TURNS = (cachedTurns && cachedTurns.data) ? cachedTurns.data : [];
      INSPECTIONS = (cachedInsp && cachedInsp.data) ? cachedInsp.data : [];
      updateCacheBadge('cached', cachedWO.timestamp, true);
      showToast('API unavailable \u2014 loaded stale cache (' + WORK_ORDERS.length + ' WOs, ' + cacheAgeStr(cachedWO) + ')');
    }
  } catch (e) {
    console.log('Stale cache load failed: ' + (e.message || e));
  }
}

// Manual refresh — force re-fetch everything from API
// On v8 proxy: first invalidates server-side cache so fresh data is fetched
async function refreshData() {
  var btn = $('#refreshBtn');
  if (btn.disabled) return;
  btn.disabled = true;
  btn.classList.add('spinning');
  btn.innerHTML = '<i class="fas fa-sync-alt"></i> Syncing\u2026';

  try {
    // v8+: invalidate all server-side caches before re-fetching
    if (supportsServerCacheOps()) {
      try {
        await proxyAction('cache_invalidate');
        console.log('v8+: Server-side cache cleared before refresh');
      } catch (e) {
        console.log('v8+: cache_invalidate failed (non-fatal): ' + (e.message || e));
      }
    }
    await fetchAllLive();
    showToast('Data refreshed \u2014 ' + WORK_ORDERS.length + ' work orders loaded');
  } catch (err) {
    showToast('Refresh failed: ' + (err.message || err));
  } finally {
    btn.disabled = false;
    btn.classList.remove('spinning');
    btn.innerHTML = '<i class="fas fa-sync-alt"></i> Refresh';
  }
}


/* =================================================================
   ATTENTION REQUIRED — Dashboard overview panel
   Renders stalled turns, overdue inspections, and vendor alerts
   into the #attentionSection on the dashboard.
   Called by renderAll() and renderDashboardKPIs().
   ================================================================= */
function renderAttentionPanel() {
  var attnToday = new Date();

  // Stalled turns
  var stalledEl = document.getElementById('attentionStalledBody');
  if (stalledEl) {
    var stalled = TURN_PIPE_DATA.filter(function(p) { return p.isStalled && !p.isCompleted; });
    // Also show turns past deposit deadline
    var depositOverdue = TURN_PIPE_DATA.filter(function(p) { return p.isConfirmed && p.sla && p.sla.overdue && !p.isCompleted; });
    if (stalled.length === 0 && depositOverdue.length === 0) {
      stalledEl.innerHTML = '<div class="attn-empty"><i class="fas fa-check-circle" style="color:var(--success)"></i> No stalled turns</div>';
    } else {
      var stalledHtml = '';
      if (depositOverdue.length > 0) {
        stalledHtml += '<div class="attn-count" style="color:var(--danger)">' + depositOverdue.length + ' turn' + (depositOverdue.length > 1 ? 's' : '') + ' past deposit deadline</div>';
        depositOverdue.slice(0, 3).forEach(function(p) {
          stalledHtml += '<div class="attn-item"><span class="attn-label">' + escapeHtml(p.unit) + ' — ' + escapeHtml(p.property) + '</span><span class="attn-value" style="color:var(--danger)">' + Math.abs(p.sla.calendarDaysLeft) + 'd over</span></div>';
        });
      }
      if (stalled.length > 0) {
        stalledHtml += '<div style="margin-top:' + (depositOverdue.length > 0 ? '6' : '0') + 'px;font-weight:600;font-size:11px;color:var(--warning)">' + stalled.length + ' stalled (' + CONFIG.TURN_STALLED_DAYS + 'd+ no progress)</div>';
        stalled.slice(0, 3).forEach(function(p) {
          stalledHtml += '<div class="attn-item"><span class="attn-label">' + escapeHtml(p.unit) + ' — ' + escapeHtml(p.property) + '</span><span class="attn-value" style="color:var(--warning)">' + p.elapsed + 'd</span></div>';
        });
        if (stalled.length > 3) stalledHtml += '<div style="text-align:center;color:var(--text-muted);font-size:10px;margin-top:4px">+' + (stalled.length - 3) + ' more</div>';
      }
      stalledEl.innerHTML = stalledHtml;
    }
  }

  // Overdue inspections + compliance %
  var overdueEl = document.getElementById('attentionOverdueBody');
  if (overdueEl) {
    var overdue = INSPECTIONS.filter(function(r) {
      var state = getInspectionCompliance(r, attnToday);
      return state.overdue;
    });
    var totalInsp = INSPECTIONS.length;
    var compliant = totalInsp - overdue.length;
    var compPct = totalInsp > 0 ? Math.round((compliant / totalInsp) * 100) : 100;
    if (overdue.length === 0) {
      overdueEl.innerHTML = '<div class="attn-empty"><i class="fas fa-check-circle" style="color:var(--success)"></i> All inspections current (' + compPct + '% compliant)</div>';
    } else {
      var overdueHtml = '<div class="attn-count">' + overdue.length + ' overdue <span style="font-size:10px;font-weight:400;color:var(--text-muted)">(' + compPct + '% compliant)</span></div>';
      overdue.slice(0, 5).forEach(function(r) {
        var state = getInspectionCompliance(r, attnToday);
        var ago = state.missingMoveInInspection
          ? 'Missing move-in inspection'
          : (state.anchorDate ? state.daysSince + 'd ago' : 'Never');
        overdueHtml += '<div class="attn-item"><span class="attn-label">' + escapeHtml(r.unit || '') + ' — ' + escapeHtml(r.propertyName || '') + '</span><span class="attn-value" style="color:var(--danger)">' + ago + '</span></div>';
      });
      if (overdue.length > 5) overdueHtml += '<div style="text-align:center;color:var(--text-muted);font-size:10px;margin-top:4px">+' + (overdue.length - 5) + ' more</div>';
      overdueEl.innerHTML = overdueHtml;
    }
  }

  // Vendor alerts (expired insurance)
  var vendorEl = document.getElementById('attentionVendorsBody');
  if (vendorEl) {
    var vAlerts = VENDORS.filter(function(v) {
      var ed = v.insurance ? new Date(v.insurance) : null;
      return ed && ed < attnToday;
    });
    var expiringSoon = VENDORS.filter(function(v) {
      var ed = v.insurance ? new Date(v.insurance) : null;
      if (!ed || ed < attnToday) return false;
      return daysBetween(attnToday, ed) <= CONFIG.VENDOR_EXPIRY_ALERT_DAYS;
    });
    if (vAlerts.length === 0 && expiringSoon.length === 0) {
      vendorEl.innerHTML = '<div class="attn-empty"><i class="fas fa-check-circle" style="color:var(--success)"></i> All vendors compliant</div>';
    } else {
      var vendorHtml = '';
      if (vAlerts.length > 0) {
        vendorHtml += '<div class="attn-count" style="color:var(--danger)">' + vAlerts.length + ' expired</div>';
        vAlerts.slice(0, 3).forEach(function(v) {
          vendorHtml += '<div class="attn-item"><span class="attn-label">' + escapeHtml(v.name) + '</span><span class="attn-value" style="color:var(--danger)">Expired</span></div>';
        });
        if (vAlerts.length > 3) vendorHtml += '<div style="text-align:center;color:var(--text-muted);font-size:10px;margin-top:2px">+' + (vAlerts.length - 3) + ' more</div>';
      }
      if (expiringSoon.length > 0) {
        vendorHtml += '<div style="margin-top:6px;font-weight:600;font-size:11px;color:var(--warning)">' + expiringSoon.length + ' expiring within ' + CONFIG.VENDOR_EXPIRY_ALERT_DAYS + ' days</div>';
        expiringSoon.slice(0, 3).forEach(function(v) {
          var ed = new Date(v.insurance);
          var daysLeft = daysBetween(attnToday, ed);
          vendorHtml += '<div class="attn-item"><span class="attn-label">' + escapeHtml(v.name) + '</span><span class="attn-value" style="color:var(--warning)">' + daysLeft + 'd</span></div>';
        });
      }
      vendorEl.innerHTML = vendorHtml;
    }
  }

  // Urgent WOs — unassigned + high priority
  var urgentEl = document.getElementById('attentionUrgentBody');
  if (urgentEl) {
    var urgentOpen = WORK_ORDERS.filter(function(w) {
      return (w.priority === 'Urgent' || w.priority === 'Emergency') && w.status !== 'Completed' && w.status !== 'Canceled';
    });
    var unassigned = urgentOpen.filter(function(w) { return !w.vendorName && !w.vendor; });
    if (urgentOpen.length === 0) {
      urgentEl.innerHTML = '<div class="attn-empty"><i class="fas fa-check-circle" style="color:var(--success)"></i> No urgent work orders</div>';
    } else {
      var urgentHtml = '<div class="attn-count">' + urgentOpen.length + ' urgent</div>';
      if (unassigned.length > 0) {
        urgentHtml += '<div style="font-weight:600;font-size:11px;color:var(--danger);margin-bottom:4px">' + unassigned.length + ' unassigned!</div>';
        unassigned.slice(0, 3).forEach(function(w) {
          urgentHtml += '<div class="attn-item"><span class="attn-label">#' + escapeHtml(String(w.id)) + ' — ' + escapeHtml((w.description || '').substring(0, 40)) + '</span><span class="attn-value" style="color:var(--danger)">No vendor</span></div>';
        });
      } else {
        urgentOpen.slice(0, 3).forEach(function(w) {
          urgentHtml += '<div class="attn-item"><span class="attn-label">#' + escapeHtml(String(w.id)) + ' — ' + escapeHtml((w.description || '').substring(0, 40)) + '</span><span class="attn-value">' + escapeHtml(w.vendorName || w.vendor || '') + '</span></div>';
        });
      }
      if (urgentOpen.length > 3) urgentHtml += '<div style="text-align:center;color:var(--text-muted);font-size:10px;margin-top:4px">+' + (urgentOpen.length - 3) + ' more</div>';
      urgentEl.innerHTML = urgentHtml;
    }
  }
}

// Patch renderAll to include attention panel
var _origRenderAll = renderAll;
renderAll = function() {
  _origRenderAll();
  try { renderAttentionPanel(); } catch (e) { console.log('renderAttentionPanel error: ' + (e.message || e)); }
};

// Patch renderDashboardKPIs to include attention panel
var _origRenderDashKPIs = renderDashboardKPIs;
renderDashboardKPIs = function() {
  _origRenderDashKPIs();
  try { renderAttentionPanel(); } catch (e) { /* */ }
};

// ═══════════════════════════════════════════════════════
// DATABASE ADMIN GUI — sql_query / sql_execute via proxy
// ═══════════════════════════════════════════════════════
(function initDbAdmin() {
  var editor = document.getElementById('dbEditor');
  var runBtn = document.getElementById('dbRunBtn');
  var execBtn = document.getElementById('dbExecBtn');
  var clearBtn = document.getElementById('dbClearBtn');
  var resultsBody = document.getElementById('dbResultsBody');
  var resultsMeta = document.getElementById('dbResultsMeta');
  var connStatus = document.getElementById('dbConnStatus');
  var csvBtn = document.getElementById('dbExportCsv');
  var jsonBtn = document.getElementById('dbExportJson');
  var histList = document.getElementById('dbHistoryList');
  var confirmOverlay = document.getElementById('dbConfirmOverlay');
  var confirmSQL = document.getElementById('dbConfirmSQL');
  var confirmCancel = document.getElementById('dbConfirmCancel');
  var confirmExec = document.getElementById('dbConfirmExec');
  if (!editor || !runBtn) return;

  var _dbHistory = [];
  var _dbLastRows = [];
  var _dbLastCols = [];
  var _pendingSQL = null;
  var keyInput = document.getElementById('dbAdminKey');

  // Load persisted admin key
  if (keyInput) {
    keyInput.value = localStorage.getItem('hm_proxy_admin_key') || '';
    keyInput.addEventListener('change', function() {
      localStorage.setItem('hm_proxy_admin_key', keyInput.value.trim());
    });
    keyInput.addEventListener('blur', function() {
      localStorage.setItem('hm_proxy_admin_key', keyInput.value.trim());
    });
  }

  function getAdminKey() {
    return (keyInput ? keyInput.value.trim() : '') || localStorage.getItem('hm_proxy_admin_key') || '';
  }

  function setStatus(cls, text) {
    connStatus.className = 'dbadmin-status ' + cls;
    connStatus.innerHTML = '<i class="fas fa-circle" style="font-size:6px"></i> ' + text;
  }

  function addHistory(sql) {
    _dbHistory = _dbHistory.filter(function(h) { return h !== sql; });
    _dbHistory.unshift(sql);
    if (_dbHistory.length > 50) _dbHistory.length = 50;
    renderHistory();
  }

  function renderHistory() {
    if (!histList) return;
    histList.innerHTML = _dbHistory.map(function(sql) {
      return '<div class="dbadmin-history-item" title="' + escapeHtml(sql) + '">' + escapeHtml(sql.length > 80 ? sql.substring(0, 80) + '\u2026' : sql) + '</div>';
    }).join('') || '<div style="padding:8px;color:var(--text-muted);font-size:11px">No history yet</div>';
  }

  // History click → populate editor
  if (histList) {
    histList.addEventListener('click', function(ev) {
      var item = ev.target.closest('.dbadmin-history-item');
      if (item) { editor.value = item.title; editor.focus(); }
    });
  }

  // Shortcut buttons
  var shortcuts = document.getElementById('dbShortcuts');
  if (shortcuts) {
    shortcuts.addEventListener('click', function(ev) {
      var btn = ev.target.closest('[data-sql]');
      if (btn) { editor.value = btn.dataset.sql; editor.focus(); }
    });
  }

  function runQuery() {
    var sql = editor.value.trim();
    if (!sql) return;
    setStatus('ok', 'Running\u2026');
    var t0 = Date.now();
    var key = getAdminKey();
    proxyPost('sql_query', { query: sql, key: key }).then(function(data) {
      var ms = Date.now() - t0;
      setStatus('ok', 'OK');
      addHistory(sql);
      if (data.rows && data.columns) {
        _dbLastRows = data.rows;
        _dbLastCols = data.columns;
        renderDbTable(data.rows, data.columns);
        resultsMeta.textContent = data.rows.length + ' row' + (data.rows.length !== 1 ? 's' : '') + ' \u00b7 ' + ms + 'ms';
        csvBtn.style.display = data.rows.length > 0 ? '' : 'none';
        jsonBtn.style.display = data.rows.length > 0 ? '' : 'none';
      } else {
        resultsBody.innerHTML = '<div class="dbadmin-msg">' + escapeHtml(JSON.stringify(data, null, 2).substring(0, 500)) + '</div>';
        resultsMeta.textContent = ms + 'ms';
        csvBtn.style.display = 'none';
        jsonBtn.style.display = 'none';
      }
    }).catch(function(err) {
      setStatus('err', 'Error');
      resultsBody.innerHTML = '<div class="dbadmin-msg" style="color:var(--danger)"><i class="fas fa-exclamation-circle"></i> ' + escapeHtml(err.message) + '</div>';
      resultsMeta.textContent = '';
      csvBtn.style.display = 'none';
      jsonBtn.style.display = 'none';
    });
  }

  function runExecute(sql) {
    sql = sql || editor.value.trim();
    if (!sql) return;
    setStatus('ok', 'Executing\u2026');
    var t0 = Date.now();
    var key = getAdminKey();
    proxyPost('sql_execute', { query: sql, key: key }).then(function(data) {
      var ms = Date.now() - t0;
      setStatus('ok', 'OK');
      addHistory(sql);
      var affected = data.rowsAffected != null ? data.rowsAffected : '?';
      resultsBody.innerHTML = '<div class="dbadmin-msg" style="color:var(--success)"><i class="fas fa-check-circle"></i> Executed successfully \u2014 ' + affected + ' row(s) affected' + (data.lastInsertRowid ? ' (last rowid: ' + data.lastInsertRowid + ')' : '') + '</div>';
      resultsMeta.textContent = ms + 'ms';
      csvBtn.style.display = 'none';
      jsonBtn.style.display = 'none';
    }).catch(function(err) {
      setStatus('err', 'Error');
      resultsBody.innerHTML = '<div class="dbadmin-msg" style="color:var(--danger)"><i class="fas fa-exclamation-circle"></i> ' + escapeHtml(err.message) + '</div>';
    });
  }

  function renderDbTable(rows, cols) {
    if (!rows.length) { resultsBody.innerHTML = '<div class="dbadmin-msg">No rows returned</div>'; return; }
    var thead = '<thead><tr>' + cols.map(function(c) { return '<th>' + escapeHtml(c) + '</th>'; }).join('') + '</tr></thead>';
    var tbody = '<tbody>' + rows.map(function(row) {
      return '<tr>' + cols.map(function(c) {
        var v = typeof row === 'object' && !Array.isArray(row) ? row[c] : row[cols.indexOf(c)];
        return (v === null || v === undefined) ? '<td class="db-null">NULL</td>' : '<td title="' + escapeHtml(String(v)) + '">' + escapeHtml(String(v)) + '</td>';
      }).join('') + '</tr>';
    }).join('') + '</tbody>';
    resultsBody.innerHTML = '<table>' + thead + tbody + '</table>';
  }

  function checkDestructive() {
    var sql = editor.value.trim();
    if (!sql) return;
    var up = sql.toUpperCase();
    var destructive = ['DROP', 'DELETE', 'TRUNCATE', 'UPDATE', 'ALTER', 'INSERT', 'CREATE'];
    var hit = destructive.find(function(kw) { return new RegExp('(^|;\\s*)' + kw + '\\b').test(up); });
    if (hit) {
      _pendingSQL = sql;
      confirmSQL.textContent = sql.length > 200 ? sql.substring(0, 200) + '\u2026' : sql;
      confirmOverlay.style.display = '';
    } else {
      runExecute(sql);
    }
  }

  // Event listeners
  runBtn.addEventListener('click', runQuery);
  execBtn.addEventListener('click', checkDestructive);
  clearBtn.addEventListener('click', function() { editor.value = ''; resultsBody.innerHTML = '<div class="dbadmin-msg">Run a query to see results here</div>'; resultsMeta.textContent = ''; csvBtn.style.display = 'none'; jsonBtn.style.display = 'none'; });
  editor.addEventListener('keydown', function(ev) {
    if ((ev.metaKey || ev.ctrlKey) && ev.key === 'Enter') { ev.preventDefault(); runQuery(); }
  });

  // Confirm dialog
  if (confirmCancel) confirmCancel.addEventListener('click', function() { confirmOverlay.style.display = 'none'; _pendingSQL = null; });
  if (confirmExec) confirmExec.addEventListener('click', function() { confirmOverlay.style.display = 'none'; if (_pendingSQL) runExecute(_pendingSQL); _pendingSQL = null; });

  // Export
  if (csvBtn) csvBtn.addEventListener('click', function() {
    if (!_dbLastRows.length) return;
    var esc2 = function(v) { var s = v == null ? '' : String(v); return (s.indexOf(',') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1) ? '"' + s.replace(/"/g, '""') + '"' : s; };
    var lines = [_dbLastCols.map(esc2).join(',')];
    _dbLastRows.forEach(function(row) {
      lines.push(_dbLastCols.map(function(c) { return esc2(typeof row === 'object' && !Array.isArray(row) ? row[c] : row[_dbLastCols.indexOf(c)]); }).join(','));
    });
    var blob = new Blob([lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a'); a.href = url; a.download = 'db_query_' + new Date().toISOString().slice(0, 10) + '.csv';
    document.body.appendChild(a); a.click(); document.body.removeChild(a); setTimeout(function() { URL.revokeObjectURL(url); }, 5000);
  });
  if (jsonBtn) jsonBtn.addEventListener('click', function() {
    if (!_dbLastRows.length) return;
    var blob = new Blob([JSON.stringify(_dbLastRows, null, 2)], { type: 'application/json;charset=utf-8;' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a'); a.href = url; a.download = 'db_query_' + new Date().toISOString().slice(0, 10) + '.json';
    document.body.appendChild(a); a.click(); document.body.removeChild(a); setTimeout(function() { URL.revokeObjectURL(url); }, 5000);
  });

  renderHistory();
})();

// ═══════════════════════════════════════════════════════
// LIVE WEBHOOK FEED DRAWER — polls webhook_events, decodes
// ═══════════════════════════════════════════════════════
(function initWebhookLiveDrawer() {
  var drawer = document.getElementById('live-feed-drawer');
  var items = document.getElementById('live-feed-items');
  var status = document.getElementById('live-feed-status');
  var badge = document.getElementById('feed-badge');
  var fabBadge = document.getElementById('fab-badge');
  if (!drawer || !items || !status || !badge || !fabBadge) return;

  var _events = [];
  var _seenCount = 0;
  var _pollTimer = null;

  var WH_LABELS = {
    'work_order.created': 'Work Order Created',
    'work_order.updated': 'Work Order Updated',
    'work_order.status_changed': 'WO Status Changed',
    'work_order.completed': 'Work Order Completed',
    'work_order.canceled': 'Work Order Canceled',
    'unit_turn.created': 'Turn Created',
    'unit_turn.updated': 'Turn Updated',
    'unit_turn.completed': 'Turn Completed',
    'inspection.created': 'Inspection Created',
    'inspection.completed': 'Inspection Completed',
    'tenant.move_out': 'Tenant Move-Out',
    'tenant.move_in': 'Tenant Move-In',
    'lease.created': 'Lease Created',
    'lease.updated': 'Lease Updated'
  };

  var WH_ICONS = {
    work_order: { icon: 'fa-wrench', color: 'var(--accent)' },
    unit_turn: { icon: 'fa-exchange-alt', color: 'var(--purple)' },
    inspection: { icon: 'fa-clipboard-check', color: 'var(--info)' },
    tenant: { icon: 'fa-user', color: 'var(--success)' },
    lease: { icon: 'fa-file-contract', color: 'var(--warning)' }
  };

  function isOpen() {
    return drawer.classList.contains('open');
  }

  function clearUnseen() {
    _seenCount = _events.length;
    updateBadge();
  }

  function updateBadge() {
    var unseen = _events.length - _seenCount;
    if (unseen < 0) unseen = 0;
    var text = unseen > 99 ? '99+' : String(unseen);
    badge.textContent = text;
    fabBadge.textContent = text;
    badge.style.display = unseen > 0 ? '' : 'none';
    fabBadge.style.display = unseen > 0 ? '' : 'none';
  }

  function decodeEvent(e) {
    var meta = extractWebhookMeta(e);
    var evtType = meta.eventType || e.event_type || e.type || '';
    var resType = meta.resourceType || e.resource_type || '';
    var view = interpretWebhookEvent(e, meta);
    var iconMeta = WH_ICONS[resType] || { icon: view.iconClass || 'fa-circle-dot', color: 'var(--text-muted)' };
    return {
      label: view.title || decodeWebhookTitle(e, meta) || WH_LABELS[evtType] || e.event_label || e.title || evtType || 'Webhook Event',
      body: view.description || meta.resourceName || e.resource_name || e.body || '',
      icon: iconMeta.icon,
      color: iconMeta.color,
      ts: e.ts || e.timestamp || e.received_at || ''
    };
  }

  function render() {
    if (!_events.length) {
      items.innerHTML = '<p class="feed-empty">No events yet. Webhook activity will appear here.</p>';
      return;
    }
    items.innerHTML = _events.slice(0, 100).map(function(e) {
      var d = decodeEvent(e);
      return '<div class="feed-item">' +
        '<div class="feed-item-header">' +
          '<span style="color:' + d.color + ';width:16px;text-align:center"><i class="fas ' + d.icon + '"></i></span>' +
          '<span class="feed-item-title">' + escapeHtml(d.label) + '</span>' +
          '<span class="feed-item-time">' + (d.ts ? timeAgo(d.ts) : '') + '</span>' +
        '</div>' +
        (d.body ? '<div class="feed-item-body">' + escapeHtml(String(d.body).substring(0, 220)) + '</div>' : '') +
      '</div>';
    }).join('');
  }

  function mergeEvents(events) {
    if (!Array.isArray(events) || events.length === 0) return;
    var existing = {};
    _events.forEach(function(e) {
      existing[(e.id || '') + '|' + (e.ts || e.timestamp || e.received_at || '') + '|' + (e.title || e.event_label || '')] = true;
    });
    var added = 0;
    events.forEach(function(e) {
      var key = (e.id || '') + '|' + (e.ts || e.timestamp || e.received_at || '') + '|' + (e.title || e.event_label || '');
      if (!existing[key]) {
        _events.push(e);
        existing[key] = true;
        added++;
      }
    });
    if (added > 0) {
      _events.sort(function(a, b) {
        return new Date(b.ts || b.timestamp || b.received_at || 0) - new Date(a.ts || a.timestamp || a.received_at || 0);
      });
      if (_events.length > 500) _events.length = 500;
      if (isOpen()) clearUnseen();
      updateBadge();
      render();
    }
  }

  function setStatus(text) {
    status.innerHTML = '<span class="feed-pulse"></span> ' + escapeHtml(text);
  }

  function pollLiveEvents() {
    proxyAction('webhook_events', { limit: 100 }).then(function(data) {
      mergeEvents(data && Array.isArray(data.events) ? data.events : []);
      setStatus('Listening for AppFolio events…');
    }).catch(function() {
      setStatus('Live feed temporarily unavailable. Retrying…');
    });
  }

  window.WebhookLive = {
    clearUnseen: clearUnseen,
    merge: mergeEvents
  };

  if (typeof pollWebhookEvents === 'function') {
    var _origPollWH = pollWebhookEvents;
    pollWebhookEvents = function() {
      var result = _origPollWH.apply(this, arguments);
      setTimeout(function() {
        if (typeof WEBHOOK_EVENTS !== 'undefined' && WEBHOOK_EVENTS.length > 0) {
          mergeEvents(WEBHOOK_EVENTS);
        }
      }, 200);
      return result;
    };
  }

  _pollTimer = setInterval(pollLiveEvents, 15000);
  setTimeout(pollLiveEvents, 2000);
  render();
})();

// ═══════════════════════════════════════════════════════
// NAV TAB WIRING FOR DBADMIN
// ═══════════════════════════════════════════════════════
(function wireDbAdminTab() {
  var dbTab = document.querySelector('[data-tab="dbadmin"]');
  if (dbTab) {
    dbTab.addEventListener('click', function() {
      if (!isTabAllowedForRole('dbadmin')) return;
      // Navigate to sec-dbadmin like other tabs
      document.querySelectorAll('.section').forEach(function(s) { s.classList.remove('active'); });
      document.querySelectorAll('.nav-tab').forEach(function(t) { t.classList.remove('active'); });
      var sec = document.getElementById('sec-dbadmin');
      if (sec) sec.classList.add('active');
      dbTab.classList.add('active');
    });
  }
})();

/* ================================================================
   HANDYMANAGER v9.0 — PART 3: DISPATCH CONTROL MODULE
   All proxy calls use existing proxyAction() + API_PROXY.
   Tech UUIDs must hold "Maintenance Tech" role in AppFolio or
   every automated PATCH returns 422 "User not found".
   Concurrent PATCHes to the same WO: first succeeds, second fails.
   The midnight cron serialises writes with 200ms delays.
   ================================================================ */

// ── Module-level state ──────────────────────────────────────────
var DISPATCH = {
  queue: [], techs: [], audit: [], blasts: [], claims: [], comms: [], stats: {}, monitored: [],
  initialized: false, activePanel: 'grades',
  cronSecret: '', queueFilter: '', queueStatus: '',
  auditFilter: '', auditWoFilter: '',
  paused: false,
  activeBranch: 'all',
  hiddenAssignees: {},
  autoSyncAssignees: true,
  autoSyncCooldownSec: 120,
  _lastAssigneeSyncAt: 0,
  tier1GroupUuid: '',
  tier2GroupUuid: '',
  _pollTimer: null, _lastAuditMax: 0, POLL_MS: 12000,
};

var DISPATCH_BRANCHES = {
  phoenix: { key: 'phoenix', label: 'Phoenix', uuid: 'efe085ca-229e-11ef-bfba-069ca18f5865' },
  tucson: { key: 'tucson', label: 'Tucson', uuid: 'a3db4460-22b3-11ef-bfba-069ca18f5865' }
};

// ── Config field definitions ────────────────────────────────────
var DISPATCH_CONFIG_FIELDS = [
  { key:'warn_threshold_hours',        label:'Warning Threshold',       hint:'Hours inactive before 36-hr warning SMS',       type:'number', min:12,  max:47,  default:36, unit:'hours' },
  { key:'reassign_threshold_hours',    label:'Reassignment Threshold',  hint:'Hours inactive before midnight auto-reassign',  type:'number', min:24,  max:120, default:48, unit:'hours' },
  { key:'go_back_tolerance_pct',       label:'Go-Back Forgiveness',     hint:'Go-back % at or below this gets no penalty',    type:'number', min:0,   max:10,  default:2,  unit:'%', step:0.5 },
  { key:'tier2_claim_window_hours',    label:'Tier 2 Claim Window',     hint:'Hours for Tier 2 to reply Y and claim',         type:'number', min:1,   max:48,  default:24, unit:'hours' },
  { key:'grace_period_enabled',        label:'Grace Period',            hint:'Allow one note-activity grace per WO',          type:'toggle', default:1 },
  { key:'max_reassigns_before_escalate',label:'Escalation Trigger',    hint:'Reassignments before admin alert fires',        type:'number', min:1,   max:5,   default:2,  unit:'reassigns' },
  { key:'dispatch_auto_sync_assignees',label:'Auto Sync Assignees',    hint:'Automatically sync assignee roster when AppFolio assignment events arrive', type:'toggle', default:1 },
  { key:'dispatch_auto_sync_cooldown_sec',label:'Auto Sync Cooldown',  hint:'Minimum seconds between automatic assignee sync runs', type:'number', min:30, max:1800, default:120, unit:'sec' },
];

// ── Audit event metadata ────────────────────────────────────────
var AUDIT_EVENT_META = {
  auto_reassigned:              { icon:'fa-random',          cls:'danger',  label:'Auto-Reassigned'       },
  reassignment_warning_sent:    { icon:'fa-bell',            cls:'warning', label:'Warning Sent'          },
  grace_period_granted:         { icon:'fa-hand-paper',      cls:'info',    label:'Grace Period Granted'  },
  auto_exempt_activated:        { icon:'fa-bell-slash',      cls:'muted',   label:'Exempt Activated'      },
  tier2_blast_sent:             { icon:'fa-satellite-dish',  cls:'danger',  label:'Tier 2 Blast Sent'     },
  escalation_tier2_blast:       { icon:'fa-exclamation',     cls:'danger',  label:'Escalated → Tier 2'   },
  escalation_no_tech_available: { icon:'fa-user-times',      cls:'danger',  label:'No Tech Available'     },
  tenant_sms_sent:              { icon:'fa-comment-dots',    cls:'success', label:'Tenant SMS Sent'       },
  tech_roster_updated:          { icon:'fa-user-edit',       cls:'info',    label:'Roster Updated'        },
};

// ── Helpers ─────────────────────────────────────────────────────
function getDispatchCronSecret() {
  return localStorage.getItem('hm_cron_secret') || DISPATCH.cronSecret || '';
}
function getDispatchAdminKey() {
  var el = document.getElementById('dbAdminKey');
  return (el ? el.value.trim() : '') || localStorage.getItem('hm_proxy_admin_key') || '';
}
function _dispatchSqlSafe(v) {
  return String(v || '').replace(/'/g, "''");
}
function getBranchGroupNameByUuid(uuid) {
  var id = String(uuid || '').trim();
  if (!id) return '';
  var g = (PROPERTY_GROUPS || []).find(function(x) { return String(x.id || '') === id; });
  return g && g.name ? String(g.name) : '';
}
function getDispatchWoBranch(wo) {
  if (!wo) return 'unknown';
  var name = wo.propertyName || wo.property_name || '';
  var pid = wo.propertyId || wo.property_id || '';
  var phx = getBranchGroupNameByUuid(DISPATCH_BRANCHES.phoenix.uuid);
  var tuc = getBranchGroupNameByUuid(DISPATCH_BRANCHES.tucson.uuid);
  if (phx && isInPropertyGroup(pid, name, phx)) return 'phoenix';
  if (tuc && isInPropertyGroup(pid, name, tuc)) return 'tucson';
  return 'unknown';
}
function getDispatchQueueBranch(row) {
  if (!row) return 'unknown';
  var woId = String(row.wo_id || '').trim();
  var woNum = String(row.wo_number || '').trim();
  var match = null;
  if (woId) {
    match = (WORK_ORDERS || []).find(function(w) {
      return String(w.uuid || '') === woId;
    });
  }
  if (!match && woNum) {
    match = (WORK_ORDERS || []).find(function(w) {
      return String(w.id || '') === woNum;
    });
  }
  return getDispatchWoBranch(match);
}
function normalizeTechBranch(tech) {
  var z = String((tech && tech.geo_zone) || '').toLowerCase();
  if (z.indexOf('tucson') !== -1) return 'tucson';
  if (z.indexOf('phoenix') !== -1) return 'phoenix';
  if (z === 'tucson' || z === 'phoenix') return z;
  return 'unknown';
}
function isTechHidden(techId) {
  return !!DISPATCH.hiddenAssignees[String(techId || '')];
}
function parseHiddenAssigneeMap(raw) {
  if (!raw) return {};
  try {
    var obj = JSON.parse(raw);
    return obj && typeof obj === 'object' ? obj : {};
  } catch (_) {
    return {};
  }
}
function isMonitoredWorkOrder(woId) {
  var id = String(woId || '').trim();
  if (!id) return false;
  return (DISPATCH.monitored || []).some(function(m) {
    return String((m && m.wo_id) || '').trim() === id;
  });
}
async function saveDispatchConfigKey(key, value) {
  var adminKey = getDispatchAdminKey();
  if (!adminKey) return false;
  var q = "INSERT OR REPLACE INTO proxy_config (key,value,updated_at) VALUES ('" +
    _dispatchSqlSafe(key) + "','" + _dispatchSqlSafe(value) + "',datetime('now'))";
  var r = await proxyPost('sql_execute', { key: adminKey, query: q });
  return !!(r && (r.ok || r.rowsAffected >= 0));
}
function extractAssigneeCandidatesFromWorkOrders(rows) {
  var out = {};
  (rows || []).forEach(function(r) {
    var branch = getDispatchWoBranch({
      propertyName: r.property_name || r.property || r.PropertyName || '',
      propertyId: r.property_id || r.PropertyId || ''
    });
    if (branch === 'unknown') return;
    var addCandidate = function(id, name) {
      var cid = String(id || '').trim();
      var cname = String(name || '').trim();
      if (!cid || !cname) return;
      if (!out[cid]) {
        out[cid] = { tech_id: cid, tech_name: cname, branch: branch, count: 0 };
      }
      out[cid].count += 1;
    };
    addCandidate(
      r.assigned_user_id || r.AssignedUserId || r.assigned_to_id || r.AssignedToId || r.assigned_user_uuid || r.AssignedUserUUID,
      r.assigned_user || r.AssignedUser || r.assigned_to || r.AssignedTo || r.assigned_user_name || r.AssignedUserName
    );
    var arr = r.AssignedUsers || r.assigned_users || [];
    if (Array.isArray(arr)) {
      arr.forEach(function(u) {
        if (!u || typeof u !== 'object') return;
        addCandidate(u.Id || u.id || u.UserId || u.user_id, u.Name || u.name || u.FullName || u.full_name);
      });
    }
  });
  return Object.keys(out).map(function(k) { return out[k]; }).sort(function(a, b) { return b.count - a.count; });
}

// ── v9 Toast ────────────────────────────────────────────────────
function v9Toast(title, desc, severity, durationMs) {
  var container = document.getElementById('v9-toast-container');
  if (!container) return;
  severity   = severity   || 'info';
  durationMs = durationMs || 4500;
  var ICONS  = { success:'fa-check-circle', warning:'fa-exclamation-triangle', danger:'fa-times-circle', info:'fa-info-circle' };
  var COLORS = { success:'var(--success)', warning:'var(--warning)', danger:'var(--danger)', info:'var(--accent)' };
  var t = document.createElement('div');
  t.style.cssText = 'display:flex;align-items:flex-start;gap:10px;padding:10px 14px;border-radius:10px;background:var(--bg-card);border:1px solid var(--border);box-shadow:var(--shadow);pointer-events:all;opacity:0;transform:translateX(16px);transition:opacity .2s,transform .2s;min-width:220px;border-left:3px solid ' + (COLORS[severity] || 'var(--accent)');
  t.innerHTML =
    '<span style="color:' + (COLORS[severity] || 'var(--accent)') + ';font-size:1rem;flex-shrink:0;margin-top:1px"><i class="fas ' + (ICONS[severity] || 'fa-info-circle') + '"></i></span>' +
    '<div style="flex:1;min-width:0">' +
      '<div style="font-weight:700;font-size:.84rem;line-height:1.3">' + escapeHtml(title) + '</div>' +
      (desc ? '<div style="font-size:.75rem;color:var(--text-muted);margin-top:2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + escapeHtml(String(desc).substring(0,110)) + '</div>' : '') +
    '</div>' +
    '<button onclick="this.parentNode.remove()" style="background:none;border:none;color:var(--text-muted);cursor:pointer;font-size:.85rem;padding:0 2px">✕</button>';
  container.appendChild(t);
  requestAnimationFrame(function(){ t.style.opacity='1'; t.style.transform='translateX(0)'; });
  setTimeout(function(){ t.style.opacity='0'; t.style.transform='translateX(8px)'; setTimeout(function(){ t.remove(); },250); }, durationMs);
}

// ── Stats bar ───────────────────────────────────────────────────
function updateDispatchStats(d) {
  if (!d) return;
  var s = d.stats || {};
  var techs = (d.tech_roster || []).filter(function(t) {
    if (isTechHidden(t.tech_id)) return false;
    var branch = normalizeTechBranch(t);
    return DISPATCH.activeBranch === 'all' || branch === DISPATCH.activeBranch;
  });
  var blasts = (d.blasts || []).filter(function(b) {
    if (DISPATCH.activeBranch === 'all') return true;
    var qb = getDispatchQueueBranch({ wo_id: b.wo_id, wo_number: b.wo_number });
    return qb === DISPATCH.activeBranch;
  });
  var tier1Active = techs.filter(function(t){ return t.tier===1 && t.active; }).length;
  var tier2Active = techs.filter(function(t){ return t.tier===2 && t.active; }).length;
  var openBlasts  = blasts.filter(function(b){ return b.status==='open'; }).length;
  var map = {
    'dstat-queue': s.total||0, 'dstat-warned': s.warned_total||0,
    'dstat-escalated': s.escalated||0, 'dstat-exempt': s.exempt||0,
    'dstat-tier1': tier1Active, 'dstat-tier2': tier2Active,
    'dstat-blasts': openBlasts, 'dstat-comms': d.comms_count||0,
  };
  Object.keys(map).forEach(function(id){
    var chip = document.getElementById(id);
    if (chip) chip.querySelector('.stat-chip__value').textContent = map[id];
  });
  var badge = document.getElementById('dispatchBadge');
  if (badge) {
    var n = (s.escalated||0) + openBlasts;
    badge.textContent = n; badge.style.display = n>0?'':'none';
  }
}

// ── Proxy POST helper ───────────────────────────────────────────
function resolveDispatchProxyBaseUrl() {
  var base = sanitizeProxy(API_PROXY || '');
  if (base) {
    API_PROXY = base;
    try { localStorage.setItem('hm_proxy_url', base); } catch (e) { /* */ }
    return base;
  }

  var proxyInput = $('#vaultProxy');
  if (proxyInput && proxyInput.value) {
    base = sanitizeProxy(proxyInput.value || '');
    if (base) {
      API_PROXY = base;
      try { localStorage.setItem('hm_proxy_url', base); } catch (e2) { /* */ }
      return base;
    }
  }

  try {
    base = sanitizeProxy(localStorage.getItem('hm_proxy_url') || '');
    if (base) {
      API_PROXY = base;
      return base;
    }
  } catch (e3) { /* */ }

  return '';
}

function dispatchPost(action, body) {
  var base = resolveDispatchProxyBaseUrl();
  if (!base) {
    return Promise.reject(new Error('Dispatch system: proxy base URL not configured. Open Vault and set Proxy URL.'));
  }
  var sep = base.indexOf('?')!==-1?'&':'?';
  var url = base + sep + 'action=' + encodeURIComponent(action);
  return fetch(url, {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(body)
  }).then(function(r){
    if (!r.ok) {
      return r.text().then(function(t) {
        throw new Error('Dispatch ' + action + ' failed: HTTP ' + r.status + (t ? ' — ' + t.substring(0, 200) : ''));
      });
    }
    return r.json();
  });
}

// ══════════════════════════════════════════════════════════════════
// RENDER FUNCTIONS
// ══════════════════════════════════════════════════════════════════

function renderDispatchGrades(techs) {
  var tbody = document.getElementById('gradeTableBody');
  if (!tbody) return;
  var active = (techs||[]).filter(function(t){
    if (!t.active) return false;
    if (isTechHidden(t.tech_id)) return false;
    var branch = normalizeTechBranch(t);
    if (DISPATCH.activeBranch !== 'all' && branch !== DISPATCH.activeBranch) return false;
    return true;
  })
    .sort(function(a,b){ if(a.tier!==b.tier) return a.tier-b.tier; return (Number(b.performance_score)||0)-(Number(a.performance_score)||0); });
  if (active.length===0) {
    tbody.innerHTML='<tr><td colspan="10"><div class="dispatch-empty"><i class="fas fa-hard-hat"></i>No active techs. Add techs in the Roster tab.</div></td></tr>';
    return;
  }
  var html='';
  active.forEach(function(t,i){
    var score=Number(t.performance_score||100).toFixed(1);
    var share=Number(t.target_share_pct||0).toFixed(1);
    var sn=Number(score);
    var goBack=Number(t.go_back_pct||0).toFixed(1);
    var reassign=Number(t.reassign_pct||0).toFixed(1);
    var scoreColor=sn>=80?'var(--success)':sn>=60?'var(--warning)':'var(--danger)';
    var scoreBg=sn>=80?'var(--success-dim)':sn>=60?'var(--warning-dim)':'var(--danger-dim)';
    var tierColor=t.tier===1?'var(--accent)':'var(--purple)';
    var tierBg=t.tier===1?'var(--accent-dim)':'var(--purple-dim)';
    var statusLabel=sn>=80?'🟢 Strong':sn>=60?'🟡 Monitor':'🔴 At Risk';
    var sharePct=Math.min(100,Number(share));
    html+='<tr>'+
      '<td style="font-family:var(--font-mono);font-size:.78rem;color:var(--text-muted)">'+(i+1)+'</td>'+
      '<td><strong>'+escapeHtml(t.tech_name)+'</strong>'+(t.tier?'<span style="margin-left:6px;background:'+tierBg+';color:'+tierColor+';padding:1px 6px;border-radius:4px;font-size:.68rem;font-weight:700">T'+t.tier+'</span>':'')+(t.tech_phone?'<div style="font-family:var(--font-mono);font-size:.68rem;color:var(--text-muted)">'+escapeHtml(t.tech_phone)+'</div>':'')+'</td>'+
      '<td><span style="background:'+scoreBg+';color:'+scoreColor+';font-family:var(--font-mono);font-weight:800;font-size:.85rem;padding:2px 8px;border-radius:6px">'+score+'</span></td>'+
      '<td><div style="position:relative;width:80px;height:16px;background:var(--bg-input);border-radius:4px;overflow:hidden"><div style="height:100%;width:'+sharePct+'%;background:var(--accent);opacity:.7"></div><span style="position:absolute;inset:0;display:flex;align-items:center;justify-content:center;font-family:var(--font-mono);font-size:.65rem;font-weight:700">'+share+'%</span></div></td>'+
      '<td style="font-family:var(--font-mono);font-weight:600">'+(t.active_wo_count||0)+'</td>'+
      '<td style="font-family:var(--font-mono);color:'+(Number(goBack)>5?'var(--danger)':Number(goBack)>2?'var(--warning)':'inherit')+'">'+goBack+'%</td>'+
      '<td style="font-family:var(--font-mono);color:'+(Number(reassign)>15?'var(--danger)':Number(reassign)>5?'var(--warning)':'inherit')+'">'+reassign+'%</td>'+
      '<td style="font-size:.78rem;color:var(--text-muted)">'+escapeHtml(normalizeTechBranch(t)==='unknown'?(t.geo_zone||'—'):normalizeTechBranch(t))+'</td>'+
      '<td style="font-size:.8rem">'+statusLabel+'</td>'+
      '<td><button class="btn-dispatch-secondary btn-xs-dispatch" onclick="DispatchRoster.openEdit(\''+escapeHtml(t.tech_id)+'\')"><i class="fas fa-pencil-alt"></i></button></td>'+
      '</tr>';
  });
  tbody.innerHTML=html;
}

function renderDispatchQueue(queue) {
  var tbody=document.getElementById('queueTableBody');
  if (!tbody) return;
  var search=(DISPATCH.queueFilter||'').toLowerCase();
  var status=(DISPATCH.queueStatus||'');
  var STATUS_ICONS={'exempt':'🔕','Escalated':'🚨','warned':'⚠️','grace':'🔵','reassigned':'🔄','monitoring':'🟢'};
  var filtered=(queue||[]).filter(function(r){
    var rowBranch = getDispatchQueueBranch(r);
    if (DISPATCH.activeBranch !== 'all' && rowBranch !== DISPATCH.activeBranch) return false;
    if (isTechHidden(r.assigned_tech_id)) return false;
    if(search){var hay=((r.property_address||'')+' '+(r.assigned_tech_name||'')).toLowerCase();if(hay.indexOf(search)===-1)return false;}
    if(status==='warned'&&!r.warning_sent)return false;
    if(status==='exempt'&&!r.auto_exempt)return false;
    if(status==='escalated'&&!r.escalated)return false;
    if(status==='grace'&&!r.grace_used)return false;
    if(status==='reassigned'&&!r.reassignment_count)return false;
    if(status==='monitoring'&&!isMonitoredWorkOrder(r.wo_id))return false;
    return true;
  });
  if(filtered.length===0){
    tbody.innerHTML='<tr><td colspan="8"><div class="dispatch-empty"><i class="fas fa-check-circle" style="color:var(--success)"></i>'+(search||status?'No queue entries match your filter.':'Queue is empty — no stale work orders.')+'</div></td></tr>';
    return;
  }
  var html='';
  filtered.forEach(function(r){
    var rowBranch = getDispatchQueueBranch(r);
    var monitored = isMonitoredWorkOrder(r.wo_id);
    var icon=monitored?'🟢':r.auto_exempt?'🔕':r.escalated?'🚨':r.warning_sent?'⚠️':r.grace_used?'🔵':'⚪';
    var woId=escapeHtml(String(r.wo_id||''));
    var woNum=escapeHtml(String(r.wo_number||r.wo_id||'').substring(0,20));
    var addr=escapeHtml(String(r.property_address||'—').substring(0,40));
    var tech=escapeHtml(String(r.assigned_tech_name||'—'));
    var branchLabel = rowBranch === 'unknown' ? 'Unknown' : (rowBranch === 'phoenix' ? 'Phoenix' : 'Tucson');
    var firstSeen=r.first_seen_at?timeAgo(r.first_seen_at):'—';
    var lastAct=r.last_reassigned_at?timeAgo(r.last_reassigned_at):(r.warning_sent_at?timeAgo(r.warning_sent_at):'—');
    var exemptBtn=!r.auto_exempt
      ?'<button class="btn-dispatch-secondary btn-xs-dispatch" onclick="DispatchQueue.markExempt(\''+woId+'\',\''+woNum+'\')">🔕 Exempt</button>'
      :'<button class="btn-dispatch-danger btn-xs-dispatch" onclick="DispatchQueue.clearExempt(\''+woId+'\')">🔓 Clear</button>';
    var monitorBtn=monitored
      ?'<button class="btn-dispatch-primary btn-xs-dispatch" onclick="DispatchQueue.unmonitor(\''+woId+'\')">🟢 Monitored</button>'
      :'<button class="btn-dispatch-secondary btn-xs-dispatch" onclick="DispatchQueue.monitor(\''+woId+'\')">⭐ Monitor</button>';
    html+='<tr style="'+(r.escalated?'background:rgba(209,59,59,.04)':'')+(r.auto_exempt?';opacity:.55':'')+'">'+
      '<td style="text-align:center;font-size:1rem">'+icon+'</td>'+
      '<td style="font-family:var(--font-mono);font-size:.8rem;color:var(--accent)">#'+woNum+'</td>'+
      '<td title="'+escapeHtml(r.property_address||'')+'">'+addr+'</td>'+
      '<td>'+tech+'<div style="font-size:.66rem;color:var(--text-muted);font-family:var(--font-mono)">'+branchLabel+'</div></td>'+
      '<td>'+(Number(r.reassignment_count)>0?'<span style="background:var(--danger-dim);color:var(--danger);padding:1px 6px;border-radius:4px;font-family:var(--font-mono);font-size:.72rem;font-weight:700">'+r.reassignment_count+'×</span>':'<span style="color:var(--text-muted)">—</span>')+'</td>'+
      '<td style="font-family:var(--font-mono);font-size:.72rem;color:var(--text-muted)">'+firstSeen+'</td>'+
      '<td style="font-family:var(--font-mono);font-size:.72rem;color:var(--text-muted)">'+lastAct+'</td>'+
      '<td style="display:flex;gap:4px;flex-wrap:wrap">'+monitorBtn+exemptBtn+
        '<button class="btn-dispatch-secondary btn-xs-dispatch" onclick="DispatchQueue.viewAudit(\''+woId+'\')">📋 Audit</button>'+
      '</td>'+
      '</tr>';
  });
  tbody.innerHTML=html;
}

function renderDispatchRoster(techs) {
  var tbody=document.getElementById('rosterTableBody');
  if (!tbody) return;
  if(!techs||techs.length===0){
    tbody.innerHTML='<tr><td colspan="9"><div class="dispatch-empty"><i class="fas fa-hard-hat"></i>No techs yet. Click <strong>+ Add Tech</strong>.</div></td></tr>';
    return;
  }
  var html='';
  (techs || []).forEach(function(t){
    var branch = normalizeTechBranch(t);
    if (DISPATCH.activeBranch !== 'all' && branch !== DISPATCH.activeBranch) return;
    var sn=Number(t.performance_score||100);
    var scoreColor=sn>=80?'var(--success)':sn>=60?'var(--warning)':'var(--danger)';
    var scoreBg=sn>=80?'var(--success-dim)':sn>=60?'var(--warning-dim)':'var(--danger-dim)';
    var hidden = isTechHidden(t.tech_id);
    var rowOpacity = (!t.active || hidden) ? '.5' : '1';
    var tid=escapeHtml(t.tech_id||'');
    var tname=escapeHtml(t.tech_name||'');
    html+='<tr style="opacity:'+rowOpacity+'">'+
      '<td><span style="font-size:.7rem;font-family:var(--font-mono);font-weight:700;padding:2px 6px;border-radius:4px;background:'+(t.tier===1?'var(--accent-dim)':'var(--purple-dim)')+';color:'+(t.tier===1?'var(--accent)':'var(--purple)')+'">Tier '+t.tier+'</span></td>'+
      '<td><strong>'+tname+'</strong><div style="font-family:var(--font-mono);font-size:.66rem;color:var(--text-muted)">'+escapeHtml(String(t.tech_id||'').substring(0,20))+'…</div></td>'+
      '<td style="font-family:var(--font-mono);font-size:.78rem">'+(t.tech_phone?escapeHtml(t.tech_phone):'<span style="color:var(--danger)">⚠️ No phone</span>')+'</td>'+
      '<td style="font-size:.78rem;color:var(--text-muted)">'+escapeHtml(branch==='unknown'?(t.geo_zone||'—'):branch)+'</td>'+
      '<td><span style="background:'+scoreBg+';color:'+scoreColor+';font-family:var(--font-mono);font-weight:800;font-size:.82rem;padding:2px 8px;border-radius:6px">'+sn.toFixed(1)+'</span></td>'+
      '<td style="font-family:var(--font-mono);font-weight:600">'+(t.active_wo_count||0)+'</td>'+
      '<td style="font-family:var(--font-mono);font-size:.78rem">'+Number(t.target_share_pct||0).toFixed(1)+'%</td>'+
      '<td><span style="color:'+(t.active?'var(--success)':'var(--text-muted)')+'">●</span> '+(t.active?'Active':'Inactive')+(hidden?'<div style="font-size:.66rem;color:var(--warning);font-family:var(--font-mono)">Hidden</div>':'')+'</td>'+
      '<td style="display:flex;gap:4px;flex-wrap:wrap">'+
        '<button class="btn-dispatch-secondary btn-xs-dispatch" onclick="DispatchRoster.openEdit(\''+tid+'\')"><i class="fas fa-pencil-alt"></i> Edit</button>'+
        '<button class="btn-dispatch-'+(t.active?'danger':'secondary')+' btn-xs-dispatch" onclick="DispatchRoster.toggleActive(\''+tid+'\',\''+tname+'\','+(t.active?1:0)+')">'+
          (t.active?'Deactivate':'Reactivate')+'</button>'+
        '<button class="btn-dispatch-'+(hidden?'secondary':'warning')+' btn-xs-dispatch" onclick="DispatchRoster.toggleHidden(\''+tid+'\',\''+tname+'\','+(hidden?1:0)+')">'+(hidden?'Unhide':'Hide')+'</button>'+
      '</td>'+
      '</tr>';
  });
  if (!html) {
    tbody.innerHTML='<tr><td colspan="9"><div class="dispatch-empty"><i class="fas fa-filter"></i>No roster entries for the selected branch.</div></td></tr>';
    return;
  }
  tbody.innerHTML=html;
}

function renderDispatchConfig(configRows) {
  var el=document.getElementById('configForm');
  if (!el) return;
  var configMap={};
  (configRows||[]).forEach(function(r){ configMap[r.key]=r.value; });
  var pausedVal = Number(configMap.dispatch_paused || 0) === 1;
  DISPATCH.paused = pausedVal;
  var savedTier1 = localStorage.getItem('hm_dispatch_tier1_group_uuid') || configMap.dispatch_tier1_group_uuid || DISPATCH.tier1GroupUuid || DISPATCH_BRANCHES.phoenix.uuid;
  var savedTier2 = localStorage.getItem('hm_dispatch_tier2_group_uuid') || configMap.dispatch_tier2_group_uuid || DISPATCH.tier2GroupUuid || DISPATCH_BRANCHES.tucson.uuid;
  DISPATCH.tier1GroupUuid = savedTier1;
  DISPATCH.tier2GroupUuid = savedTier2;
  var savedBranch = localStorage.getItem('hm_dispatch_active_branch') || configMap.dispatch_active_branch || DISPATCH.activeBranch || 'all';
  DISPATCH.activeBranch = ['all','phoenix','tucson'].indexOf(savedBranch) !== -1 ? savedBranch : 'all';
  var hiddenRaw = localStorage.getItem('hm_dispatch_hidden_assignees') || configMap.dispatch_hidden_assignees || '{}';
  DISPATCH.hiddenAssignees = parseHiddenAssigneeMap(hiddenRaw);
  var savedSecret=localStorage.getItem('hm_cron_secret')||'';
  var html=
    '<div style="display:flex;justify-content:space-between;align-items:center;padding:10px 0;border-bottom:1px solid var(--border);gap:12px;flex-wrap:wrap">'+
    '<div><label style="font-weight:700;font-size:.85rem;display:block">Automation Kill Switch</label>'+
    '<span style="font-size:.72rem;color:var(--text-muted)">Pause all reassignment automation immediately from GUI.</span></div>'+
    '<div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap">'+
      '<span id="dispatchPauseState" style="font-family:var(--font-mono);font-size:.72rem;padding:3px 8px;border-radius:999px;background:'+(pausedVal?'var(--danger-dim)':'var(--success-dim)')+';color:'+(pausedVal?'var(--danger)':'var(--success)')+'">'+(pausedVal?'PAUSED':'RUNNING')+'</span>'+
      '<button class="'+(pausedVal?'btn-dispatch-primary':'btn-dispatch-warning')+' btn-xs-dispatch" onclick="DispatchConfig.togglePause()">'+(pausedVal?'<i class="fas fa-play"></i> Resume':'<i class="fas fa-pause"></i> Pause')+'</button>'+
    '</div></div>'+
    '<div style="display:flex;justify-content:space-between;align-items:flex-start;padding:10px 0;border-bottom:1px solid var(--border);gap:12px;flex-wrap:wrap">'+
    '<div><label style="font-weight:700;font-size:.85rem;display:block">Load Assignees From AppFolio</label>'+
    '<span style="font-size:.72rem;color:var(--text-muted)">Click to import current AppFolio users plus recent assignment activity into the dispatch roster using the two property-group UUIDs below.</span></div>'+
    '<div style="display:flex;flex-direction:column;gap:6px;min-width:300px">'+
      '<input type="text" id="cfgTier1GroupUuid" value="'+escapeHtml(savedTier1)+'" placeholder="Tier 1 Property Group UUID" style="font-family:var(--font-mono);font-size:.74rem;padding:5px 9px;border-radius:7px;border:1px solid var(--border);background:var(--bg-input);color:var(--text-primary)">'+
      '<input type="text" id="cfgTier2GroupUuid" value="'+escapeHtml(savedTier2)+'" placeholder="Tier 2 Property Group UUID" style="font-family:var(--font-mono);font-size:.74rem;padding:5px 9px;border-radius:7px;border:1px solid var(--border);background:var(--bg-input);color:var(--text-primary)">'+
      '<div style="display:flex;gap:6px;flex-wrap:wrap">'+
        '<button class="btn-dispatch-secondary btn-xs-dispatch" onclick="DispatchConfig.saveTierGroupUuids()"><i class="fas fa-save"></i> Save UUIDs</button>'+
        '<button class="btn-dispatch-primary btn-xs-dispatch" onclick="DispatchConfig.syncAssignees()"><i class="fas fa-users-cog"></i> Load From AppFolio</button>'+
        '<button class="btn-dispatch-warning btn-xs-dispatch" onclick="DispatchConfig.populateTestQueue()"><i class="fas fa-vial"></i> Populate Test Queue</button>'+
      '</div>'+
      '<span style="font-size:.7rem;color:var(--text-muted)">Tip: paste your two target property-group UUIDs, save them, then click Load From AppFolio.</span>'+
    '</div></div>'+
    '<div style="display:flex;justify-content:space-between;align-items:center;padding:10px 0;border-bottom:1px solid var(--border);gap:12px;flex-wrap:wrap">'+
    '<div><label style="font-weight:700;font-size:.85rem;display:block">Dispatch Branch Scope</label>'+
    '<span style="font-size:.72rem;color:var(--text-muted)">Filters roster and queue so Phoenix and Tucson stay fully separated.</span></div>'+
    '<div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap">'+
      '<select id="cfgDispatchActiveBranch" style="font-family:var(--font-mono);font-size:.74rem;padding:5px 9px;border-radius:7px;border:1px solid var(--border);background:var(--bg-input);color:var(--text-primary)">'+
        '<option value="all" '+(DISPATCH.activeBranch==='all'?'selected':'')+'>All Branches</option>'+
        '<option value="phoenix" '+(DISPATCH.activeBranch==='phoenix'?'selected':'')+'>Phoenix</option>'+
        '<option value="tucson" '+(DISPATCH.activeBranch==='tucson'?'selected':'')+'>Tucson</option>'+
      '</select>'+
      '<button class="btn-dispatch-secondary btn-xs-dispatch" onclick="DispatchConfig.saveActiveBranch()"><i class="fas fa-save"></i> Save</button>'+
    '</div></div>'+
    '<div style="display:flex;justify-content:space-between;align-items:center;padding:10px 0;border-bottom:1px solid var(--border);gap:12px;flex-wrap:wrap">'+
    '<div><label style="font-weight:700;font-size:.85rem;display:block">Hidden Assignees</label>'+
    '<span style="font-size:.72rem;color:var(--text-muted)">Keep AppFolio users available but hide them from active dispatch controls.</span></div>'+
    '<div style="font-family:var(--font-mono);font-size:.76rem;color:var(--text-muted)">'+Object.keys(DISPATCH.hiddenAssignees).length+' hidden</div>'+
    '</div>'+
    '<div style="display:flex;justify-content:space-between;align-items:center;padding:10px 0;border-bottom:1px solid var(--border);gap:12px;flex-wrap:wrap">'+
    '<div><label style="font-weight:700;font-size:.85rem;display:block">Cron Secret</label>'+
    '<span style="font-size:.72rem;color:var(--text-muted)">CRON_SECRET env var — required to trigger manual cron passes</span></div>'+
    '<div style="display:flex;gap:6px;align-items:center">'+
      '<input type="password" id="cfgCronSecret" value="'+escapeHtml(savedSecret)+'" placeholder="Enter CRON_SECRET…" style="font-family:var(--font-mono);font-size:.78rem;padding:5px 9px;border-radius:7px;border:1px solid var(--border);background:var(--bg-input);color:var(--text-primary);width:200px">'+
      '<button class="btn-dispatch-secondary btn-xs-dispatch" onclick="DispatchConfig.saveCronSecret()"><i class="fas fa-save"></i> Save</button>'+
    '</div></div>';
  DISPATCH_CONFIG_FIELDS.forEach(function(f){
    var current=configMap[f.key]!==undefined?configMap[f.key]:f.default;
    var inputHtml=f.type==='toggle'
      ?'<label style="position:relative;display:inline-block;width:42px;height:22px;flex-shrink:0">'+
         '<input type="checkbox" id="cfg-'+f.key+'" '+(Number(current)?'checked':'')+' style="opacity:0;width:0;height:0">'+
         '<span style="position:absolute;inset:0;background:'+(Number(current)?'var(--accent)':'var(--border)')+';border-radius:22px;cursor:pointer;transition:.2s;display:block" onclick="this.style.background=this.previousElementSibling.checked?\'var(--border)00\':\'var(--accent)\'" id="cfgSlider-'+f.key+'"></span>'+
         '</label>'
      :'<div style="display:flex;align-items:center;gap:6px">'+
         '<input type="number" id="cfg-'+f.key+'" value="'+escapeHtml(String(current))+'" min="'+f.min+'" max="'+f.max+'" step="'+(f.step||1)+'" style="font-family:var(--font-mono);font-size:.82rem;padding:5px 9px;border-radius:7px;border:1px solid var(--border);background:var(--bg-input);color:var(--text-primary);width:78px;text-align:center">'+
         '<span style="font-size:.75rem;color:var(--text-muted)">'+escapeHtml(f.unit||'')+'</span>'+
         '</div>';
    html+='<div style="display:flex;justify-content:space-between;align-items:center;padding:10px 0;border-bottom:1px solid var(--border);gap:12px;flex-wrap:wrap">'+
      '<div><label style="font-weight:700;font-size:.85rem;display:block" for="cfg-'+f.key+'">'+escapeHtml(f.label)+'</label>'+
      '<span style="font-size:.72rem;color:var(--text-muted)">'+escapeHtml(f.hint)+'</span></div>'+
      '<div style="display:flex;align-items:center;gap:8px">'+
        inputHtml+
        '<button class="btn-dispatch-primary btn-xs-dispatch" onclick="DispatchConfig.saveField(\''+f.key+'\')">Save</button>'+
      '</div></div>';
  });
  el.innerHTML=html;
  // Wire toggle slider visual state properly
  DISPATCH_CONFIG_FIELDS.filter(function(f){return f.type==='toggle';}).forEach(function(f){
    var cb=document.getElementById('cfg-'+f.key);
    var sl=document.getElementById('cfgSlider-'+f.key);
    if(cb&&sl){
      sl.style.background=cb.checked?'var(--accent)':'var(--border)';
      cb.addEventListener('change',function(){ sl.style.background=cb.checked?'var(--accent)':'var(--border)'; });
    }
  });
  var topBranch = document.getElementById('dispatchBranchSelect');
  if (topBranch) topBranch.value = DISPATCH.activeBranch;
}

function renderDispatchAudit(auditRows) {
  var el=document.getElementById('auditFeed');
  if (!el) return;
  var evtFilter=DISPATCH.auditFilter||'';
  var woFilter=DISPATCH.auditWoFilter||'';
  var filtered=(auditRows||[]).filter(function(a){
    if(evtFilter&&a.event_type!==evtFilter)return false;
    if(woFilter&&String(a.wo_id||'').indexOf(woFilter)===-1)return false;
    return true;
  });
  if(filtered.length===0){
    el.innerHTML='<div class="dispatch-empty"><i class="fas fa-stream"></i>'+(evtFilter||woFilter?'No audit events match your filter.':'No audit events recorded yet.')+'</div>';
    return;
  }
  var ICON_COLORS={danger:'var(--danger)',warning:'var(--warning)',success:'var(--success)',info:'var(--info)',muted:'var(--text-muted)'};
  var html='';
  filtered.slice(0,100).forEach(function(a){
    var meta=AUDIT_EVENT_META[a.event_type]||{icon:'fa-circle-dot',cls:'muted',label:a.event_type||'Event'};
    var evtData={};
    try{evtData=JSON.parse(a.event_data||'{}');}catch(e){}
    var parts=[];
    if(evtData.from&&evtData.to)parts.push(escapeHtml(evtData.from)+' → '+escapeHtml(evtData.to));
    if(evtData.tech)parts.push(escapeHtml(evtData.tech));
    if(evtData.address)parts.push(escapeHtml(String(evtData.address).substring(0,50)));
    if(evtData.reassignment_count)parts.push('Reassign #'+evtData.reassignment_count);
    if(evtData.techs_notified)parts.push(evtData.techs_notified+' techs notified');
    var iconBg={'danger':'var(--danger-dim)','warning':'var(--warning-dim)','success':'var(--success-dim)','info':'var(--info-dim)','muted':'var(--bg-input)'}[meta.cls]||'var(--bg-input)';
    html+='<div style="display:flex;align-items:flex-start;gap:10px;padding:10px 0;border-bottom:1px solid var(--border);font-size:.82rem">'+
      '<div style="width:28px;height:28px;border-radius:50%;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:.72rem;background:var(--bg-input);color:'+(ICON_COLORS[meta.cls]||'var(--text-muted)')+'"><i class="fas '+meta.icon+'"></i></div>'+
      '<div style="flex:1;min-width:0">'+
        '<div style="font-weight:600;line-height:1.4">'+escapeHtml(meta.label)+'</div>'+
        (parts.length?'<div style="font-size:.72rem;color:var(--text-muted);margin-top:2px;font-family:var(--font-mono)">'+parts.join(' · ')+'</div>':'')+
        '<div style="font-size:.72rem;color:var(--text-muted);margin-top:2px;font-family:var(--font-mono)">WO: <strong style="color:var(--accent)">'+escapeHtml(String(a.wo_id||'').substring(0,24))+'</strong></div>'+
      '</div>'+
      '<div style="font-family:var(--font-mono);font-size:.7rem;color:var(--text-muted);white-space:nowrap;flex-shrink:0">'+(a.created_at?timeAgo(a.created_at):'—')+'</div>'+
      '</div>';
  });
  if(auditRows&&auditRows.length>100)html+='<div style="text-align:center;padding:.75rem;font-size:.75rem;color:var(--text-muted)">Showing 100 of '+auditRows.length+' events</div>';
  el.innerHTML=html;
}

function renderDispatchBlasts(blasts, claims) {
  var el=document.getElementById('blastsFeed');
  if (!el) return;
  var openBlasts=(blasts||[]).filter(function(b){return b.status==='open';});
  var badge=document.getElementById('blastsBadge');
  if(badge){badge.textContent=openBlasts.length;badge.style.display=openBlasts.length>0?'':'none';}
  if(!blasts||blasts.length===0){
    el.innerHTML='<div class="dispatch-empty"><i class="fas fa-satellite-dish"></i>No Tier 2 blast events. Blasts fire when no Tier 1 tech is available.</div>';
    return;
  }
  var now=Date.now();
  var html='';
  blasts.slice(0,30).forEach(function(b){
    var expiresMs=b.expires_at?new Date(b.expires_at).getTime():0;
    var remainMs=expiresMs-now;
    var remainHrs=remainMs>0?(remainMs/3600000).toFixed(1):0;
    var timerText=remainMs<=0?'Expired':remainHrs+'h remaining';
    var timerColor=remainMs<=0?'var(--text-muted)':Number(remainHrs)<4?'var(--danger)':'var(--warning)';
    var statusColor={open:'var(--danger)',claimed:'var(--success)',expired:'var(--text-muted)'}[b.status]||'var(--text-muted)';
    var blastClaims=(claims||[]).filter(function(c){return String(c.blast_id)===String(b.id);});
    var claimsHtml='';
    if(blastClaims.length>0){
      claimsHtml='<div style="margin-top:8px;display:flex;flex-direction:column;gap:4px">';
      blastClaims.forEach(function(c){
        var won=c.claim_status==='won';
        claimsHtml+='<div style="display:flex;justify-content:space-between;align-items:center;font-size:.76rem;padding:4px 8px;border-radius:5px;background:'+(won?'var(--success-dim)':'var(--bg-input)')+'">'+
          '<span><i class="fas fa-hard-hat" style="margin-right:4px"></i>'+escapeHtml(c.tech_name||'')+'</span>'+
          '<span style="font-family:var(--font-mono);font-size:.7rem;color:var(--text-muted)">'+(c.sms_sent_at?timeAgo(c.sms_sent_at):'—')+'</span>'+
          '<span style="font-size:.68rem;font-weight:700;color:'+(won?'var(--success)':'var(--warning)')+'">'+escapeHtml(c.claim_status||'pending')+'</span>'+
          '</div>';
      });
      claimsHtml+='</div>';
    }
    html+='<div style="background:var(--bg-card);border:1px solid var(--border);border-left:3px solid '+statusColor+';border-radius:10px;padding:12px;margin-bottom:8px">'+
      '<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:6px">'+
        '<div>'+
          '<div style="font-weight:700;font-size:.9rem">'+escapeHtml(b.property_addr||'—')+'</div>'+
          '<div style="font-size:.76rem;color:var(--text-muted);display:flex;gap:10px;flex-wrap:wrap;margin-top:2px">'+
            '<span><i class="fas fa-wrench" style="margin-right:3px"></i>'+escapeHtml(b.category||'—')+'</span>'+
            '<span>⚡ '+escapeHtml(b.priority||'—')+'</span>'+
            '<span><i class="fas fa-clock" style="margin-right:3px"></i>'+(b.blasted_at?timeAgo(b.blasted_at):'—')+'</span>'+
          '</div>'+
        '</div>'+
        '<div style="text-align:right">'+
          '<div style="font-family:var(--font-mono);font-weight:700;font-size:.76rem;color:'+timerColor+'">'+timerText+'</div>'+
          '<span style="font-size:.7rem;font-weight:700;color:'+statusColor+'">'+escapeHtml(b.status||'')+'</span>'+
        '</div>'+
      '</div>'+
      (blastClaims.length>0?'<div style="font-size:.7rem;color:var(--text-muted);font-family:var(--font-mono)">'+blastClaims.length+' techs notified</div>'+claimsHtml:'')+
      '</div>';
  });
  el.innerHTML=html;
}

function renderDispatchComms(comms) {
  var el=document.getElementById('commsFeed');
  if (!el) return;
  if(!comms||comms.length===0){
    el.innerHTML='<div class="dispatch-empty"><i class="fas fa-comment-slash"></i>No tenant communications logged yet.</div>';
    return;
  }
  var ICONS={"I'm On My Way":'🚗',"Let's Schedule a Visit":'📅','Arriving Today':'📬'};
  var html='';
  comms.slice(0,80).forEach(function(c){
    var icon=ICONS[c.template_used]||'💬';
    html+='<div style="display:flex;align-items:flex-start;gap:10px;padding:10px 0;border-bottom:1px solid var(--border);font-size:.82rem">'+
      '<div style="font-size:1.1rem;margin-top:1px">'+icon+'</div>'+
      '<div style="flex:1;min-width:0">'+
        '<div style="font-weight:600">'+escapeHtml(c.tech_name||'Unknown')+
          (c.template_used?'<span style="margin-left:6px;font-size:.7rem;font-family:var(--font-mono);padding:1px 5px;border-radius:4px;background:var(--accent-dim);color:var(--accent)">'+escapeHtml(c.template_used)+'</span>':'')+
        '</div>'+
        '<div style="font-size:.72rem;color:var(--text-muted);font-family:var(--font-mono);margin-top:2px">'+
          'WO: <strong style="color:var(--accent)">'+escapeHtml(String(c.wo_id||'').substring(0,20))+'</strong>'+
          ' · To: '+escapeHtml(c.tenant_phone||'—')+
          (c.rc_message_id?' · RC: '+escapeHtml(String(c.rc_message_id).substring(0,12)):'')+
        '</div>'+
      '</div>'+
      '<div style="font-family:var(--font-mono);font-size:.7rem;color:var(--text-muted);white-space:nowrap">'+(c.sent_at?timeAgo(c.sent_at):'—')+'</div>'+
      '</div>';
  });
  el.innerHTML=html;
}

// ══════════════════════════════════════════════════════════════════
// CONTROLLER OBJECTS
// ══════════════════════════════════════════════════════════════════

var DispatchQueue = {
  monitor: async function(woId) {
    try {
      var res = await dispatchPost('add_monitored_work_order', { wo_id: String(woId || '').trim() });
      if (!res || !res.ok) {
        v9Toast('Monitor failed', (res && res.error) || 'Unknown error', 'danger');
        return;
      }
      v9Toast('Monitoring enabled', 'WO ' + woId + ' added to monitored queue', 'success');
      DispatchControl.refresh();
    } catch (e) {
      v9Toast('Monitor failed', e.message || String(e), 'danger');
    }
  },
  unmonitor: async function(woId) {
    try {
      var res = await dispatchPost('remove_monitored_work_order', { wo_id: String(woId || '').trim() });
      if (!res || !res.ok) {
        v9Toast('Unmonitor failed', (res && res.error) || 'Unknown error', 'danger');
        return;
      }
      v9Toast('Monitoring removed', 'WO ' + woId + ' removed from monitored queue', 'success');
      DispatchControl.refresh();
    } catch (e) {
      v9Toast('Unmonitor failed', e.message || String(e), 'danger');
    }
  },
  markExempt: async function(woId, woNum) {
    if (!confirm('Mark WO #'+woNum+' as exempt?\n\nThis writes a system note to AppFolio and stops automated reassignment for this WO.')) return;
    try {
      var sep=API_PROXY.indexOf('?')!==-1?'&':'?';
      await fetch(API_PROXY+sep+'action=passthrough&path='+encodeURIComponent('/api/v0/work_orders/'+woId+'/notes'), {
        method:'POST',headers:{'Content-Type':'application/json'},
        body:JSON.stringify({Body:':stop-auto: Manual admin exemption applied via HandyManager Dispatch Control.'})
      });
      v9Toast('Exemption applied','WO '+woNum+' is now exempt','success');
      DispatchControl.refresh();
    } catch(e) { v9Toast('Exemption failed',e.message,'danger'); }
  },
  clearExempt: async function(woId) {
    if (!confirm('Clear exemption for WO '+woId+'?\n\nIt will re-enter automation on the next cron run.')) return;
    try {
      var key=getDispatchAdminKey();
      if(!key){v9Toast('Admin key required','Set PROXY_ADMIN_KEY in Database tab','warning');return;}
      await proxyPost('sql_execute',{key:key,query:"UPDATE reassignment_queue SET auto_exempt=0,auto_exempt_at=NULL,auto_exempt_by=NULL WHERE wo_id='"+woId.replace(/'/g,'')+"'"});
      v9Toast('Exemption cleared','WO '+woId+' will re-enter automation','success');
      DispatchControl.refresh();
    } catch(e) { v9Toast('Clear failed',e.message,'danger'); }
  },
  viewAudit: function(woId) {
    DISPATCH.auditWoFilter=woId;
    var inp=document.getElementById('auditWoFilter');
    if(inp)inp.value=woId;
    document.querySelectorAll('#dispatchSubnav .subnav-btn').forEach(function(b){b.classList.toggle('active',b.getAttribute('data-dpanel')==='audit');});
    document.querySelectorAll('#sec-dispatch .dispatch-subpanel').forEach(function(p){p.classList.toggle('active',p.id==='dispatch-panel-audit');});
    DISPATCH.activePanel='audit';
    renderDispatchAudit(DISPATCH.audit);
  }
};

var DispatchRoster = {
  _editing: false,
  _open: function() { document.getElementById('techRosterModal').classList.add('show'); },
  _close: function() { document.getElementById('techRosterModal').classList.remove('show'); },
  openAdd: function() {
    this._editing=false;
    document.getElementById('techRosterModalTitle').textContent='Add Tech to Roster';
    ['rosterTechId','rosterTechName','rosterTechPhone'].forEach(function(id){document.getElementById(id).value='';});
    document.getElementById('rosterTechId').disabled=false;
    document.getElementById('rosterTechIdHidden').value='';
    document.getElementById('rosterTechTier').value='1';
    document.getElementById('rosterTechZone').value=DISPATCH.activeBranch==='all'?'phoenix':DISPATCH.activeBranch;
    document.getElementById('rosterTechHidden').value='0';
    document.getElementById('rosterTechActive').value='1';
    this._open();
  },
  openEdit: function(techId) {
    var t=DISPATCH.techs.find(function(x){return x.tech_id===techId;});
    if(!t)return;
    this._editing=true;
    document.getElementById('techRosterModalTitle').textContent='Edit Tech';
    document.getElementById('rosterTechId').value=t.tech_id; document.getElementById('rosterTechId').disabled=true;
    document.getElementById('rosterTechIdHidden').value=t.tech_id;
    document.getElementById('rosterTechName').value=t.tech_name||'';
    document.getElementById('rosterTechPhone').value=t.tech_phone||'';
    document.getElementById('rosterTechTier').value=String(t.tier||1);
    document.getElementById('rosterTechZone').value=normalizeTechBranch(t)==='unknown'?'phoenix':normalizeTechBranch(t);
    document.getElementById('rosterTechHidden').value=isTechHidden(t.tech_id)?'1':'0';
    document.getElementById('rosterTechActive').value=String(t.active!==undefined?t.active:1);
    this._open();
  },
  save: async function() {
    var id=document.getElementById('rosterTechIdHidden').value||document.getElementById('rosterTechId').value;
    var name=document.getElementById('rosterTechName').value.trim();
    var phone=document.getElementById('rosterTechPhone').value.trim();
    var tier=Number(document.getElementById('rosterTechTier').value);
    var zone=document.getElementById('rosterTechZone').value;
    var hidden=Number(document.getElementById('rosterTechHidden').value||'0')===1;
    var active=Number(document.getElementById('rosterTechActive').value);
    if(!id||!name){v9Toast('Validation error','UUID and Name are required','warning');return;}
    if(phone&&!phone.startsWith('+')){v9Toast('Invalid phone','Must start with + (E.164): +15205551234','warning');return;}
    if(!this._editing&&id.length<30){if(!confirm('The UUID "'+id+'" looks short. Is this a valid AppFolio user UUID?\n\nThe user must have the Maintenance Tech role enabled or reassignment PATCHes will return 422.'))return;}
    try {
      var r=await dispatchPost('tech_roster',{tech_id:id,tech_name:name,tech_phone:phone,tier:tier,geo_zone:zone,active:active});
      if(r.ok){
        if (hidden) {
          DISPATCH.hiddenAssignees[id] = true;
        } else {
          delete DISPATCH.hiddenAssignees[id];
        }
        await DispatchConfig.persistHiddenAssignees();
        v9Toast('Tech saved',name+' (Tier '+tier+')','success');
        this._close();
        DispatchControl.refresh();
      }
      else v9Toast('Save failed',r.error||'Unknown error','danger');
    } catch(e){v9Toast('Save failed',e.message,'danger');}
  },
  toggleActive: async function(techId,techName,currentlyActive) {
    if(!confirm((currentlyActive?'Deactivate':'Reactivate')+' '+techName+'?'))return;
    try {
      var r=await dispatchPost('tech_roster',{tech_id:techId,tech_name:techName,active:currentlyActive?0:1});
      if(r.ok){v9Toast(techName+(currentlyActive?' deactivated':' reactivated'),'',currentlyActive?'warning':'success');DispatchControl.refresh();}
      else v9Toast('Action failed',r.error,'danger');
    }catch(e){v9Toast('Action failed',e.message,'danger');}
  },
  toggleHidden: async function(techId, techName, currentlyHidden) {
    if (!confirm((currentlyHidden ? 'Unhide ' : 'Hide ') + techName + ' from active dispatch roster?')) return;
    if (currentlyHidden) {
      delete DISPATCH.hiddenAssignees[String(techId)];
    } else {
      DISPATCH.hiddenAssignees[String(techId)] = true;
    }
    await DispatchConfig.persistHiddenAssignees();
    v9Toast(currentlyHidden ? 'Assignee visible' : 'Assignee hidden', techName, currentlyHidden ? 'success' : 'warning');
    DispatchControl._renderActivePanel();
  }
};

var DispatchComms = {
  sendMagicLinkTest: async function() {
    var phoneEl = document.getElementById('dispatchTestPhone');
    var btn = document.getElementById('btnDispatchSendTestSms');
    var resultEl = document.getElementById('dispatchTestResult');
    var phone = phoneEl ? String(phoneEl.value || '').trim() : '';
    if (!/^\+\d{10,15}$/.test(phone)) {
      if (resultEl) resultEl.textContent = 'Enter a valid E.164 phone number such as +15551234567.';
      v9Toast('Invalid phone number', 'Use E.164 format like +15551234567', 'warning');
      return;
    }

    var originalLabel = btn ? btn.innerHTML : '';
    if (btn) {
      btn.disabled = true;
      btn.innerHTML = '<i class="fas fa-circle-notch fa-spin"></i> Sending…';
    }
    if (resultEl) resultEl.textContent = 'Sending test magic-link SMS to ' + phone + '…';

    try {
      var response = await dispatchPost('send_magic_link_test_sms', {
        phone: phone,
        tech_name: 'Dispatch Test',
        tech_id: 'dispatch-test'
      });
      if (!response || !response.ok) {
        throw new Error((response && response.error) || 'Test SMS send failed');
      }
      if (resultEl) {
        var linkHtml = response.magic_link
          ? ' <a href="' + escapeHtml(response.magic_link) + '" target="_blank" rel="noopener noreferrer">Open portal</a>'
          : '';
        resultEl.innerHTML = 'Sent test message to ' + escapeHtml(phone) + '.' + linkHtml;
      }
      v9Toast('Test link sent', phone, 'success');
      DispatchControl.refresh();
    } catch (e) {
      if (resultEl) resultEl.textContent = 'Send failed: ' + (e.message || e);
      v9Toast('Test link failed', e.message || String(e), 'danger');
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.innerHTML = originalLabel;
      }
    }
  }
};

var DispatchConfig = {
  _getTierGroupInputs: function() {
    var t1=(document.getElementById('cfgTier1GroupUuid')||{}).value||'';
    var t2=(document.getElementById('cfgTier2GroupUuid')||{}).value||'';
    return { tier1:t1.trim(), tier2:t2.trim() };
  },
  applyPauseUi: function() {
    var paused=!!DISPATCH.paused;
    var btnTop=document.getElementById('btnDispatchPauseToggle');
    if(btnTop){
      btnTop.className=paused?'btn-dispatch-primary':'btn-dispatch-secondary';
      btnTop.innerHTML=paused?'<i class="fas fa-play"></i> Resume Automation':'<i class="fas fa-pause"></i> Pause Automation';
    }
    var warnBtn=document.getElementById('btnRunWarnCron');
    var reasBtn=document.getElementById('btnRunReassignCron');
    var warnManual=document.getElementById('btnManualWarn');
    var reasManual=document.getElementById('btnManualReassign');
    [warnBtn,reasBtn,warnManual,reasManual].forEach(function(b){ if(b) b.disabled=paused; });
  },
  saveTierGroupUuids: function() {
    var g=this._getTierGroupInputs();
    DISPATCH.tier1GroupUuid=g.tier1;
    DISPATCH.tier2GroupUuid=g.tier2;
    localStorage.setItem('hm_dispatch_tier1_group_uuid',g.tier1);
    localStorage.setItem('hm_dispatch_tier2_group_uuid',g.tier2);
    saveDispatchConfigKey('dispatch_tier1_group_uuid', g.tier1).catch(function(){});
    saveDispatchConfigKey('dispatch_tier2_group_uuid', g.tier2).catch(function(){});
    v9Toast('Group UUIDs saved','Tier 1/2 UUID targets stored','success');
  },
  persistHiddenAssignees: async function() {
    var payload = JSON.stringify(DISPATCH.hiddenAssignees || {});
    localStorage.setItem('hm_dispatch_hidden_assignees', payload);
    try { await saveDispatchConfigKey('dispatch_hidden_assignees', payload); } catch (_) {}
  },
  saveActiveBranch: async function() {
    var sel = document.getElementById('cfgDispatchActiveBranch');
    var val = sel ? sel.value : 'all';
    if (['all','phoenix','tucson'].indexOf(val) === -1) val = 'all';
    DISPATCH.activeBranch = val;
    localStorage.setItem('hm_dispatch_active_branch', val);
    try { await saveDispatchConfigKey('dispatch_active_branch', val); } catch (_) {}
    var topSel = document.getElementById('dispatchBranchSelect');
    if (topSel) topSel.value = val;
    DispatchControl._renderActivePanel();
    v9Toast('Branch scope updated', val === 'all' ? 'Showing all branches' : ('Showing ' + val + ' only'), 'success');
  },
  togglePause: async function() {
    var adminKey=getDispatchAdminKey();
    if(!adminKey){v9Toast('Admin key required','Set PROXY_ADMIN_KEY in Database tab first','warning');return;}
    var next = DISPATCH.paused ? '0' : '1';
    try {
      var q="INSERT OR REPLACE INTO proxy_config (key,value,updated_at) VALUES ('dispatch_paused','"+next+"',datetime('now'))";
      var r=await proxyPost('sql_execute',{key:adminKey,query:q});
      if(!(r.ok||r.rowsAffected>=0)){v9Toast('Pause toggle failed',r.error||'Unknown error','danger');return;}
      DISPATCH.paused = next==='1';
      this.applyPauseUi();
      renderDispatchConfig([{key:'dispatch_paused',value:next}].concat(DISPATCH._lastConfigRows||[]));
      v9Toast(DISPATCH.paused?'Automation paused':'Automation resumed','Dispatch reassignment controls updated',DISPATCH.paused?'warning':'success');
    } catch(e){ v9Toast('Pause toggle failed',e.message,'danger'); }
  },
  syncAssignees: async function() {
    var g=this._getTierGroupInputs();
    if(!g.tier1 || !g.tier2){v9Toast('Missing UUIDs','Enter both Tier 1 and Tier 2 property group UUIDs first','warning');return;}
    this.saveTierGroupUuids();
    var actions=['dispatch_sync_assignees','sync_assignee_roster','reassignment_sync_assignees'];
    var payload={tier1_group_uuid:g.tier1,tier2_group_uuid:g.tier2,source:'assignee_role'};
    var lastErr='';
    for(var i=0;i<actions.length;i++){
      try{
        var resp=await dispatchPost(actions[i],payload);
        if(resp&&resp.ok){
          v9Toast('Assignees synced',String(resp.synced||resp.count||0)+' techs imported from AppFolio roles','success');
          DispatchControl.refresh();
          return;
        }
        lastErr=(resp&&resp.error)?resp.error:'not supported';
      }catch(e){ lastErr=e.message||String(e); }
    }

    // Fallback: build assignee roster from AppFolio work_orders payload in browser.
    try {
      if (!PROPERTY_GROUPS || PROPERTY_GROUPS.length === 0) {
        try { await fetchPropertyGroups(); } catch (_) {}
      }
      var wo = await proxyAction('work_orders', { days: '90' });
      var candidates = extractAssigneeCandidatesFromWorkOrders((wo && wo.results) || []);
      if (candidates.length) {
        var selected = candidates.filter(function(c) {
          if (DISPATCH.activeBranch === 'all') return true;
          return c.branch === DISPATCH.activeBranch;
        });
        var count = 0;
        for (var ci = 0; ci < selected.length; ci++) {
          var c = selected[ci];
          if (isTechHidden(c.tech_id)) continue;
          var up = await dispatchPost('tech_roster', {
            tech_id: c.tech_id,
            tech_name: c.tech_name,
            tier: c.branch === 'phoenix' ? 1 : 2,
            geo_zone: c.branch,
            active: 1
          });
          if (up && up.ok) count += 1;
        }
        if (count > 0) {
          v9Toast('Assignees synced', count + ' assignees imported from AppFolio work orders', 'success');
          DispatchControl.refresh();
          return;
        }
      }
    } catch (fallbackErr) {
      lastErr = fallbackErr.message || String(fallbackErr);
    }

    try {
      var fallback = await proxyAction('reassignment_queue', {
        limit:'50',
        sync_assignees:'1',
        tier1_group_uuid:g.tier1,
        tier2_group_uuid:g.tier2
      });
      if(fallback && fallback.ok){
        v9Toast('Assignees synced','Sync completed through reassignment_queue fallback','success');
        DispatchControl.refresh();
        return;
      }
    } catch(ignoreErr) {}
    v9Toast('Sync unavailable','Proxy needs a sync_assignees action. Last error: '+lastErr,'danger',6500);
  },
  populateTestQueue: async function() {
    var g=this._getTierGroupInputs();
    if(!g.tier1 || !g.tier2){v9Toast('Missing UUIDs','Enter both Tier 1 and Tier 2 property group UUIDs first','warning');return;}
    var actions=['dispatch_seed_reassignment_test','seed_reassignment_queue_test','reassignment_queue_seed'];
    var payload={tier1_group_uuid:g.tier1,tier2_group_uuid:g.tier2,mode:'test',limit:50};
    if (DISPATCH.activeBranch !== 'all') payload.branch = DISPATCH.activeBranch;
    var lastErr='';
    for(var i=0;i<actions.length;i++){
      try{
        var resp=await dispatchPost(actions[i],payload);
        if(resp&&resp.ok){
          v9Toast('Test queue populated',String(resp.inserted||resp.seeded||resp.count||0)+' WOs added for reassignment testing','success');
          DispatchControl.refresh();
          return;
        }
        lastErr=(resp&&resp.error)?resp.error:'not supported';
      }catch(e){ lastErr=e.message||String(e); }
    }
    try {
      var fallback = await proxyAction('reassignment_queue', {
        limit:'100',
        force_refresh:'1',
        seed_test:'1',
        tier1_group_uuid:g.tier1,
        tier2_group_uuid:g.tier2
      });
      if(fallback && fallback.ok){
        v9Toast('Test queue refresh complete','Fallback queue refresh executed','success');
        DispatchControl.refresh();
        return;
      }
    } catch(ignoreErr) {}
    v9Toast('Populate unavailable','Proxy needs a queue seed action. Last error: '+lastErr,'danger',6500);
  },
  saveCronSecret: function() {
    var val=(document.getElementById('cfgCronSecret')||{}).value||'';
    localStorage.setItem('hm_cron_secret',val.trim());
    DISPATCH.cronSecret=val.trim();
    v9Toast('Cron secret saved','Stored locally in browser','success');
  },
  saveField: async function(key) {
    var f=DISPATCH_CONFIG_FIELDS.find(function(x){return x.key===key;});
    if(!f) return;
    var adminKey=getDispatchAdminKey();
    if(!adminKey){v9Toast('Admin key required','Set PROXY_ADMIN_KEY in Database tab first','warning');return;}
    var inp=document.getElementById('cfg-'+key);
    var value=f.type==='toggle'?(inp.checked?'1':'0'):(inp?inp.value:'');
    try {
      var r=await proxyPost('sql_execute',{key:adminKey,query:"INSERT OR REPLACE INTO proxy_config (key,value,updated_at) VALUES ('"+key.replace(/'/g,'')+"','"+String(value).replace(/'/g,'')+"',datetime('now'))"});
      if(r.ok||r.rowsAffected>=0)v9Toast(f.label+' saved',value+' '+(f.unit||''),'success');
      else v9Toast('Save failed',r.error||'Check PROXY_ADMIN_KEY','danger');
    }catch(e){v9Toast('Save failed',e.message,'danger');}
  },
  runCron: async function(action,btnEl) {
    if(DISPATCH.paused){v9Toast('Automation paused','Resume automation before running cron passes','warning');return;}
    var secret=getDispatchCronSecret();
    if(!secret){v9Toast('Cron secret required','Enter CRON_SECRET in Config panel and Save first','warning');return;}
    var label=btnEl.innerHTML;
    btnEl.disabled=true; btnEl.innerHTML='<i class="fas fa-circle-notch fa-spin"></i> Running…';
    var resultEl=document.getElementById('configCronResult');
    if(resultEl)resultEl.style.display='none';
    try {
      var data=await proxyPost(action,{},{'x-cron-secret':secret});
      var isOk=data.ok;
      if(resultEl){
        resultEl.className='cron-result '+(isOk?'cron-result--ok':'cron-result--err');
        resultEl.style.display='';
        var isWarn=action.indexOf('warn')!==-1;
        resultEl.innerHTML=isOk
          ?'✅ <strong>'+escapeHtml(data.run||action)+'</strong> completed '+new Date().toLocaleTimeString()+
            '<br>Candidates: '+(data.candidates||0)+' · '+(isWarn?'Warned: '+(data.warned||0):'Reassigned: '+(data.reassigned||0))+
            ' · Skipped: '+(data.skipped||0)+(data.escalated>0?' · <span style="color:var(--danger)">🚨 Escalated: '+data.escalated+'</span>':'')
          :'❌ Error: '+escapeHtml(data.error||'Unknown error');
      }
      if(isOk)v9Toast(isWarn?'Warning pass complete':'Reassign pass complete',(isWarn?data.warned:data.reassigned)+' WOs processed','success');
      else v9Toast('Cron failed',data.error,'danger');
      DispatchControl.refresh();
    }catch(e){
      if(resultEl){resultEl.className='cron-result cron-result--err';resultEl.style.display='';resultEl.innerHTML='❌ Network error: '+escapeHtml(e.message);}
      v9Toast('Cron error',e.message,'danger');
    }finally{btnEl.disabled=false;btnEl.innerHTML=label;}
  }
};

// ══════════════════════════════════════════════════════════════════
// DISPATCH CONTROL — Master controller
// ══════════════════════════════════════════════════════════════════
var DispatchControl = {
  refresh: async function() {
    var dot=document.getElementById('dispatchLiveDot');
    var lbl=document.getElementById('dispatchLiveLabel');
    if(dot)dot.style.background='var(--warning)';
    if(lbl)lbl.textContent='Refreshing…';
    try {
      var results=await Promise.all([
        proxyAction('reassignment_queue',{limit:'100'}),
        proxyAction('tenant_comms_log',{limit:'60'})
      ]);
      var queueData=results[0], commsData=results[1];
      if(queueData.ok){
        DISPATCH.queue =queueData.queue        ||[];
        DISPATCH.techs =queueData.tech_roster  ||[];
        DISPATCH.audit =queueData.audit        ||[];
        DISPATCH.blasts=queueData.blasts       ||[];
        DISPATCH.claims=queueData.tier2_claims ||[];
        DISPATCH.stats =queueData.stats        ||{};
        DISPATCH.monitored = queueData.monitored_work_orders || [];
        updateDispatchStats(queueData);
      }
      if(commsData.ok)DISPATCH.comms=commsData.results||[];
      var adminKey=getDispatchAdminKey();
      if(adminKey){
        try{
          var cfg=await proxyPost('sql_query',{key:adminKey,query:'SELECT key,value FROM proxy_config ORDER BY key'});
          if(cfg.ok){
            DISPATCH._lastConfigRows=cfg.rows||[];
            renderDispatchConfig(cfg.rows||[]);
            DispatchConfig.applyPauseUi();
          }
        }
        catch(e){}
      }
      var branchSel = document.getElementById('dispatchBranchSelect');
      if (branchSel) branchSel.value = DISPATCH.activeBranch || 'all';
      DispatchConfig.applyPauseUi();
      this._renderActivePanel();
      if(dot)dot.style.background='var(--success)';
      if(lbl)lbl.textContent='Live · '+new Date().toLocaleTimeString();
    }catch(e){
      if(dot)dot.style.background='var(--danger)';
      if(lbl)lbl.textContent='Error: '+(e.message||'refresh failed');
      v9Toast('Dispatch refresh failed',e.message,'danger');
    }
  },
  _renderActivePanel: function() {
    var p=DISPATCH.activePanel;
    if(p==='grades') renderDispatchGrades(DISPATCH.techs);
    if(p==='queue')  renderDispatchQueue(DISPATCH.queue);
    if(p==='roster') renderDispatchRoster(DISPATCH.techs);
    if(p==='audit')  renderDispatchAudit(DISPATCH.audit);
    if(p==='blasts') renderDispatchBlasts(DISPATCH.blasts,DISPATCH.claims);
    if(p==='comms')  renderDispatchComms(DISPATCH.comms);
  },
  init: async function() {
    if(DISPATCH.initialized)return;
    DISPATCH.initialized=true;
    // Sub-tab nav
    document.querySelectorAll('.subnav-btn[data-dpanel]').forEach(function(btn){
      btn.addEventListener('click',function(){
        document.querySelectorAll('.subnav-btn[data-dpanel]').forEach(function(b){b.classList.remove('active');});
        document.querySelectorAll('.dispatch-subpanel').forEach(function(p){p.classList.remove('active');});
        btn.classList.add('active');
        var panel=document.getElementById('dispatch-panel-'+btn.getAttribute('data-dpanel'));
        if(panel)panel.classList.add('active');
        DISPATCH.activePanel=btn.getAttribute('data-dpanel');
        DispatchControl._renderActivePanel();
      });
    });
    // Top-bar buttons
    var r=document.getElementById('btnDispatchRefresh');
    if(r)r.addEventListener('click',function(){DispatchControl.refresh();});
    var bsel=document.getElementById('dispatchBranchSelect');
    if(bsel)bsel.addEventListener('change',function(){
      DISPATCH.activeBranch=this.value||'all';
      localStorage.setItem('hm_dispatch_active_branch',DISPATCH.activeBranch);
      DispatchConfig.saveActiveBranch();
      DispatchControl._renderActivePanel();
    });
    var pz=document.getElementById('btnDispatchPauseToggle');
    if(pz)pz.addEventListener('click',function(){DispatchConfig.togglePause();});
    var bw=document.getElementById('btnRunWarnCron');
    if(bw)bw.addEventListener('click',function(){DispatchConfig.runCron('noon_warning_cron',bw);});
    var br=document.getElementById('btnRunReassignCron');
    if(br)br.addEventListener('click',function(){DispatchConfig.runCron('midnight_reassign_cron',br);});
    var mw=document.getElementById('btnManualWarn');
    if(mw)mw.addEventListener('click',function(){DispatchConfig.runCron('noon_warning_cron',mw);});
    var mr=document.getElementById('btnManualReassign');
    if(mr)mr.addEventListener('click',function(){DispatchConfig.runCron('midnight_reassign_cron',mr);});
    var testBtn=document.getElementById('btnDispatchSendTestSms');
    if(testBtn)testBtn.addEventListener('click',function(){DispatchComms.sendMagicLinkTest();});
    var testPhone=document.getElementById('dispatchTestPhone');
    if(testPhone)testPhone.addEventListener('keydown',function(e){if(e.key==='Enter'){e.preventDefault();DispatchComms.sendMagicLinkTest();}});
    // Queue filters
    var qs=document.getElementById('queueSearch');
    if(qs)qs.addEventListener('input',function(){DISPATCH.queueFilter=this.value;renderDispatchQueue(DISPATCH.queue);});
    var qf=document.getElementById('queueStatusFilter');
    if(qf)qf.addEventListener('change',function(){DISPATCH.queueStatus=this.value;renderDispatchQueue(DISPATCH.queue);});
    // Audit filters
    var ae=document.getElementById('auditEventFilter');
    if(ae)ae.addEventListener('change',function(){DISPATCH.auditFilter=this.value;renderDispatchAudit(DISPATCH.audit);});
    var aw=document.getElementById('auditWoFilter');
    if(aw)aw.addEventListener('input',function(){DISPATCH.auditWoFilter=this.value.trim();renderDispatchAudit(DISPATCH.audit);});
    // Roster modal
    var btnAdd=document.getElementById('btnAddTech');
    if(btnAdd)btnAdd.addEventListener('click',function(){DispatchRoster.openAdd();});
    var btnClose=document.getElementById('techRosterModalClose');
    var btnCancel=document.getElementById('techRosterModalCancel');
    var btnSave=document.getElementById('techRosterModalSave');
    if(btnClose)btnClose.addEventListener('click',function(){DispatchRoster._close();});
    if(btnCancel)btnCancel.addEventListener('click',function(){DispatchRoster._close();});
    if(btnSave)btnSave.addEventListener('click',function(){DispatchRoster.save();});
    var modal=document.getElementById('techRosterModal');
    if(modal)modal.addEventListener('click',function(e){if(e.target===modal)DispatchRoster._close();});
    // Cron secret live-save
    document.addEventListener('input',function(e){
      if(e.target&&e.target.id==='cfgCronSecret'){
        localStorage.setItem('hm_cron_secret',e.target.value.trim());
        DISPATCH.cronSecret=e.target.value.trim();
      }
    });
    // Initial render
    renderDispatchConfig([]);
    DispatchConfig.applyPauseUi();
    await this.refresh();
    // Live poll (only when dispatch section is active)
    DISPATCH._pollTimer=setInterval(function(){
      var sec=document.getElementById('sec-dispatch');
      if(sec&&sec.classList.contains('active'))DispatchControl._pollLive();
    }, DISPATCH.POLL_MS);
  },
  _pollLive: async function() {
    try {
      var data=await proxyAction('reassignment_queue',{limit:'100'});
      if(!data.ok)return;
      DISPATCH.queue=data.queue||[]; DISPATCH.techs=data.tech_roster||[];
      DISPATCH.audit=data.audit||[]; DISPATCH.blasts=data.blasts||[];
      DISPATCH.claims=data.tier2_claims||[]; DISPATCH.stats=data.stats||{};
      DISPATCH.monitored = data.monitored_work_orders || [];
      updateDispatchStats(data);
      // Toast new important audit events
      var newAudit=(data.audit||[]).filter(function(a){return a.id&&a.id>DISPATCH._lastAuditMax;});
      if(newAudit.length>0){
        DISPATCH._lastAuditMax=Math.max.apply(null,newAudit.map(function(a){return a.id||0;}));
        var importantEvts=['auto_reassigned','auto_exempt_activated','tier2_blast_sent','escalation_tier2_blast'];
        newAudit.forEach(function(a){
          if(importantEvts.indexOf(a.event_type)===-1)return;
          var meta=AUDIT_EVENT_META[a.event_type];
          if(meta)v9Toast(meta.label,'WO: '+String(a.wo_id||'').substring(0,20),meta.cls);
        });
      }
      DispatchControl._renderActivePanel();
      var lbl=document.getElementById('dispatchLiveLabel');
      if(lbl)lbl.textContent='Live · '+new Date().toLocaleTimeString();
    }catch(e){ /* silent */ }
  },
  maybeAutoSyncAssignees: async function(sourceLabel) {
    if (!DISPATCH.autoSyncAssignees) return;
    var now = Date.now();
    var cooldownMs = Math.max(30, Number(DISPATCH.autoSyncCooldownSec || 120)) * 1000;
    if (now - Number(DISPATCH._lastAssigneeSyncAt || 0) < cooldownMs) return;

    var t1 = DISPATCH.tier1GroupUuid || localStorage.getItem('hm_dispatch_tier1_group_uuid') || '';
    var t2 = DISPATCH.tier2GroupUuid || localStorage.getItem('hm_dispatch_tier2_group_uuid') || '';
    if (!t1 || !t2) return;

    DISPATCH._lastAssigneeSyncAt = now;
    try {
      var resp = await dispatchPost('dispatch_sync_assignees', {
        tier1_group_uuid: t1,
        tier2_group_uuid: t2,
        source: sourceLabel || 'webhook_live'
      });
      if (resp && resp.ok && DISPATCH.initialized) {
        await this._pollLive();
      }
    } catch (_) {
      // Non-fatal; live webhook loop should never break on sync errors.
    }
  }
};

// Wire dispatch nav tab (overrides scaffold wiring)
(function wireDispatchTab(){
  var tab=document.querySelector('.nav-tab[data-tab="dispatch"]');
  if(!tab)return;
  // Remove any prior click listener by cloning
  var fresh=tab.cloneNode(true);
  tab.parentNode.replaceChild(fresh,tab);
  fresh.addEventListener('click',function(){
    if (!isTabAllowedForRole('dispatch')) return;
    document.querySelectorAll('.nav-tab').forEach(function(t){t.classList.remove('active');});
    document.querySelectorAll('.section').forEach(function(s){s.classList.remove('active');});
    fresh.classList.add('active');
    var sec=document.getElementById('sec-dispatch');
    if(sec)sec.classList.add('active');
    DispatchControl.init();
  });
})();

// Alt+D keyboard shortcut
document.addEventListener('keydown',function(e){
  if(e.altKey&&(e.key==='d'||e.key==='D')){
    if (!isTabAllowedForRole('dispatch')) return;
    var tab=document.querySelector('.nav-tab[data-tab="dispatch"]');
    if(tab)tab.click();
  }
});

/* ================================================================
   HANDYMANAGER v9.0 — PART 4: LIVE WEBHOOK EVENT ENGINE

   Decodes every webhook event into human-readable labels,
   fires v9Toast for important events (completions/escalations),
   auto-refreshes visible panels via handymgr:webhook-invalidate,
   and updates the existing drawer with decoded titles.
   ================================================================ */

var LIVE_TOAST_EVENTS = [
  'work_order.completed','work_order.work_completed','work_order.canceled',
  'work_order.cancelled','work_order.assigned','unit_turn.completed',
  'tenant.move_out','tenant.move_in','tenant.notice_given',
  'lease.renewed','bill.approved',
  'auto_reassigned','auto_exempt_activated','tier2_blast_sent','escalation_tier2_blast',
];

// ── Description builders ────────────────────────────────────────
function _woBuildDesc(p){
  if(!p)return'';
  var parts=[];
  if(!p._omit_number && (p.work_order_number||p.number))parts.push('WO #'+(p.work_order_number||p.number));
  if(p.property_address||p.property_name)parts.push(p.property_address||p.property_name);
  if(p.category)parts.push(p.category);
  if(p.priority)parts.push('⚡ '+p.priority);
  return parts.join(' · ').substring(0,120);
}
function _turnBuildDesc(p){if(!p)return'';var pa=[];if(p.unit)pa.push('Unit '+p.unit);if(p.property_name)pa.push(p.property_name);if(p.stage)pa.push('→ '+p.stage);return pa.join(' · ').substring(0,100);}
function _tenantBuildDesc(p){if(!p)return'';var name=p.full_name||p.name||((p.first_name||'')+' '+(p.last_name||'')).trim();var pa=[];if(name)pa.push(name);if(p.property_name)pa.push(p.property_name);if(p.unit)pa.push('Unit '+p.unit);return pa.join(' · ').substring(0,100);}
function _leaseBuildDesc(p){if(!p)return'';var t=p.tenant_names||p.tenant_name||'';if(Array.isArray(t))t=t.join(', ');var pa=[];if(t)pa.push(t);if(p.property_name)pa.push(p.property_name);if(p.unit)pa.push('Unit '+p.unit);return pa.join(' · ').substring(0,100);}

// ── Event interpreter map ───────────────────────────────────────
var V9_EVENT_DEFS = {
  work_order: {
    created:       function(p){var num=p&&((p.work_order_number||p.number||'')+'');var descPayload=Object.assign({},p||{},{_omit_number:true});return{title:num?('New Work Order '+num):'New Work Order',desc:_woBuildDesc(descPayload),severity:'info',icon:'fa-wrench',color:'var(--accent)'  };},
    updated:       function(p){return{title:'Work Order Updated',    desc:_woBuildDesc(p),severity:'info',    icon:'fa-pencil-alt',          color:'var(--accent)'  };},
    completed:     function(p){return{title:'✅ WO Completed',       desc:_woBuildDesc(p),severity:'success', icon:'fa-check-circle',        color:'var(--success)' };},
    work_completed:function(p){return{title:'✅ Work Completed',     desc:_woBuildDesc(p),severity:'success', icon:'fa-check-circle',        color:'var(--success)' };},
    canceled:      function(p){return{title:'WO Cancelled',          desc:_woBuildDesc(p),severity:'warning', icon:'fa-ban',                 color:'var(--warning)' };},
    cancelled:     function(p){return{title:'WO Cancelled',          desc:_woBuildDesc(p),severity:'warning', icon:'fa-ban',                 color:'var(--warning)' };},
    assigned:      function(p){return{title:'👷 WO Assigned',        desc:_woBuildDesc(p)+(p.assigned_to?' → '+p.assigned_to:''),severity:'info',icon:'fa-user-check',color:'var(--info)' };},
    note_added:    function(p){return{title:'💬 Note Added',         desc:_woBuildDesc(p),severity:'info',    icon:'fa-sticky-note',         color:'var(--purple)'  };},
    status_changed:function(p){return{title:'WO Status → '+(p.status||'Updated'),desc:_woBuildDesc(p),severity:'info',icon:'fa-exchange-alt',color:'var(--accent)'  };},
  },
  unit_turn: {
    created:   function(p){return{title:'🏠 Turn Started',  desc:_turnBuildDesc(p),severity:'info',    icon:'fa-arrows-rotate', color:'var(--purple)' };},
    updated:   function(p){return{title:'Turn Updated',     desc:_turnBuildDesc(p),severity:'info',    icon:'fa-arrows-rotate', color:'var(--purple)' };},
    completed: function(p){return{title:'🎉 Turn Completed',desc:_turnBuildDesc(p),severity:'success', icon:'fa-flag-checkered',color:'var(--success)'};},
  },
  tenant: {
    created:      function(p){return{title:'👤 New Tenant',       desc:_tenantBuildDesc(p),severity:'info',   icon:'fa-user-plus',    color:'var(--success)'};},
    updated:      function(p){return{title:'Tenant Updated',       desc:_tenantBuildDesc(p),severity:'info',   icon:'fa-user-edit',    color:'var(--info)'   };},
    move_in:      function(p){return{title:'🔑 Tenant Move-In',   desc:_tenantBuildDesc(p),severity:'success',icon:'fa-sign-in-alt',  color:'var(--success)'};},
    move_out:     function(p){return{title:'📦 Tenant Move-Out',  desc:_tenantBuildDesc(p),severity:'warning',icon:'fa-sign-out-alt', color:'var(--warning)'};},
    notice_given: function(p){return{title:'📋 Notice to Vacate', desc:_tenantBuildDesc(p),severity:'warning',icon:'fa-file-alt',     color:'var(--warning)'};},
  },
  lease: {
    created: function(p){return{title:'📄 Lease Created',desc:_leaseBuildDesc(p),severity:'info',   icon:'fa-file-contract',color:'var(--accent)' };},
    updated: function(p){return{title:'Lease Updated',   desc:_leaseBuildDesc(p),severity:'info',   icon:'fa-file-contract',color:'var(--accent)' };},
    renewed: function(p){return{title:'✅ Lease Renewed',desc:_leaseBuildDesc(p),severity:'success',icon:'fa-redo',         color:'var(--success)'};},
  },
  vendor: {
    created: function(p){return{title:'🏢 Vendor Added',  desc:p.name||'',severity:'info',icon:'fa-hard-hat',color:'var(--purple)'};},
    updated: function(p){return{title:'Vendor Updated',   desc:p.name||'',severity:'info',icon:'fa-hard-hat',color:'var(--purple)'};},
  },
  inspection: {
    created:   function(p){return{title:'🔍 Inspection Scheduled',desc:(p.property_name||''),severity:'info',   icon:'fa-clipboard-check',color:'var(--info)'   };},
    updated:   function(p){return{title:'Inspection Updated',      desc:(p.property_name||''),severity:'info',   icon:'fa-clipboard-check',color:'var(--info)'   };},
    completed: function(p){return{title:'✅ Inspection Complete',  desc:(p.property_name||''),severity:'success',icon:'fa-clipboard-check',color:'var(--success)'};},
  },
  bill: {
    created:  function(p){return{title:'💰 Bill Created', desc:(p.vendor_name||'')+(p.property_name?' · '+p.property_name:''),severity:'info',   icon:'fa-file-invoice-dollar',color:'var(--warning)'};},
    approved: function(p){return{title:'✅ Bill Approved',desc:(p.vendor_name||'')+(p.property_name?' · '+p.property_name:''),severity:'success',icon:'fa-check',              color:'var(--success)'};},
    updated:  function(p){return{title:'Bill Updated',    desc:(p.vendor_name||''),severity:'info',   icon:'fa-file-invoice-dollar',color:'var(--warning)'};},
  },
};

var _INVAL_MAP={
  work_order:['work_orders','turn_work_orders','recent_tasks'],
  unit_turn:['turns','turn_work_orders'],vendor:['vendors'],
  inspection:['inspections'],tenant:['upcoming_moveouts'],
  lease:['upcoming_moveouts'],property:['properties'],bill:['bills'],
};

// ── Core decode ─────────────────────────────────────────────────
function decodeWebhookEventV9(evt) {
  if(!evt)return null;
  var meta=extractWebhookMeta(evt);
  var rawType = String(meta.resourceType || evt.resource_type || evt.resourceType || '').toLowerCase();
  var rType=rawType.replace(/[^a-z_]/g,'');
  var typeAliases={workorder:'work_order',work_orders:'work_order',unitturn:'unit_turn',unit_turns:'unit_turn'};
  rType=typeAliases[rType]||rType;

  var action=String(meta.eventType||evt.event_type||evt.type||'').toLowerCase();
  if(action.indexOf('.')!==-1)action=action.split('.').slice(1).join('.');
  action=action.replace(/[^a-z_]/g,'_');
  var aliases={work_completed:'completed',status_changed:'status_changed',note_added:'note_added',move_in:'move_in',move_out:'move_out',notice_given:'notice_given',work_done:'completed'};
  action=aliases[action]||action||'updated';
  var handlers=V9_EVENT_DEFS[rType];
  var view=null;
  if(handlers){try{view=(handlers[action]||handlers['updated']||null);if(view)view=view(meta.payload||{});}catch(e){view=null;}}
  if(!view){
    var fallbackTitle=typeof decodeWebhookTitle==='function'?decodeWebhookTitle(evt,meta):'Webhook Event';
    view={title:fallbackTitle||'Webhook Event',desc:meta.resourceName||evt.resource_name||'',severity:'info',icon:'fa-plug',color:'var(--purple)'};
  }
  var eventKey=rType+'.'+action;
  return {
    id:evt.id, ts:evt.ts||evt.timestamp||evt.received_at||'',
    raw:evt, category:rType, action:action,
    resource_id:meta.resourceId, resource_type:rType,
    title:view.title, desc:view.desc||'',
    severity:view.severity||'info', icon:view.icon||'fa-circle-dot', color:view.color||'var(--text-muted)',
    invalidates:_INVAL_MAP[rType]||[],
    shouldToast:LIVE_TOAST_EVENTS.indexOf(eventKey)!==-1,
  };
}

// ── Live event engine (replaces basic WebhookLive poll) ─────────
(function initLiveEventEngineV9(){
  var _lastId=0, _events=[], _unseenCount=0, _pollTimer=null, _isPolling=false;
  var POLL_MS=8000;

  function _getEl(id){return document.getElementById(id);}

  function _updateBadge(){
    var n=Math.max(0,_unseenCount); var t=n>99?'99+':String(n);
    var b=_getEl('feed-badge'), fb=_getEl('fab-badge');
    if(b){b.textContent=t;b.style.display=n>0?'':'none';}
    if(fb){fb.textContent=t;fb.style.display=n>0?'':'none';}
  }
  function _clearUnseen(){_unseenCount=0;_updateBadge();}

  async function _seedLastId(){
    try{var d=await proxyAction('webhook_events',{limit:1});if(d.ok&&d.events&&d.events.length)_lastId=d.events[0].id||0;}catch(e){}
  }

  async function _poll(){
    if(_isPolling||!API_PROXY)return;
    _isPolling=true;
    try{
      var data=await proxyAction('webhook_live',{since_id:String(_lastId),limit:'20'});
      if(!data.ok||!data.has_new)return;
      var evts=data.events||[];
      if(!evts.length)return;
      _lastId=data.max_id||_lastId;
      var decoded=evts.map(decodeWebhookEventV9).filter(Boolean);
      if(!decoded.length)return;
      decoded.forEach(function(v){_events.unshift(v);});
      if(_events.length>200)_events.length=200;
      if(decoded.some(function(v){return v.category==='work_order';})) refreshCurrentWONotes();
      // Toasts for important events
      decoded.forEach(function(v){if(v.shouldToast)v9Toast(v.title,v.desc,v.severity);});
      // Panel auto-refresh
      var keys={};
      decoded.forEach(function(v){(v.invalidates||[]).forEach(function(k){keys[k]=true;});});
      var keyArr=Object.keys(keys);
      if(keyArr.length){
        window.dispatchEvent(new CustomEvent('handymgr:webhook-invalidate',{detail:{keys:keyArr,eventCount:decoded.length,at:new Date().toISOString()}}));
        requestAnimationFrame(function(){
          var SECMAP={'sec-workorders':['work_orders','turn_work_orders','recent_tasks'],'sec-turnboard':['turns','turn_work_orders'],'sec-dashboard':['work_orders','turns','upcoming_moveouts'],'sec-inspections':['inspections'],'sec-vendors':['vendors']};
          Object.keys(SECMAP).forEach(function(secId){
            var sec=document.getElementById(secId);
            if(!sec||!sec.classList.contains('active'))return;
            if(!keyArr.some(function(k){return SECMAP[secId].indexOf(k)!==-1;}))return;
            if(secId==='sec-workorders'){try{renderWorkOrders();}catch(e){}}
            else if(secId==='sec-turnboard'){try{renderTurnBoard();}catch(e){}}
            else if(secId==='sec-dashboard'){try{renderDashboardKPIs();}catch(e){} try{renderActivityFeed();}catch(e){}}
            else if(secId==='sec-inspections'){try{renderInspections('');}catch(e){}}
            else if(secId==='sec-vendors'){try{renderVendors('');}catch(e){}}
          });
          var disSec=document.getElementById('sec-dispatch');
          if(disSec&&disSec.classList.contains('active')&&DISPATCH.initialized)DispatchControl._pollLive();
        });
      }
      // Badge
      var drawer=_getEl('live-feed-drawer');
      var drawerOpen=drawer&&drawer.classList.contains('open');
      if(!drawerOpen){_unseenCount+=decoded.length;_updateBadge();}
      _renderFeed();
      var st=_getEl('live-feed-status');
      if(st)st.innerHTML='<span class="feed-pulse"></span> Live · '+new Date().toLocaleTimeString();
    }catch(e){
      var st2=_getEl('live-feed-status');
      if(st2)st2.innerHTML='<span class="feed-pulse"></span> Reconnecting…';
    }finally{_isPolling=false;}
  }

  function _renderFeed(){
    var feedEl=_getEl('live-feed-items');
    if(!feedEl)return;
    if(!_events.length){feedEl.innerHTML='<p class="feed-empty">No events yet. Webhook activity will appear here.</p>';return;}
    var html='';
    _events.slice(0,80).forEach(function(v){
      var severityBorderColor={success:'var(--success)',warning:'var(--warning)',danger:'var(--danger)',info:'var(--accent)'}[v.severity]||'transparent';
      html+='<div class="feed-item" style="border-left:3px solid '+severityBorderColor+'">'+
        '<div class="feed-item-header">'+
          '<span style="color:'+escapeHtml(v.color)+';font-size:.9rem;flex-shrink:0;width:16px;text-align:center"><i class="fas '+escapeHtml(v.icon)+'"></i></span>'+
          '<span class="feed-item-title">'+escapeHtml(v.title)+'</span>'+
          '<span class="feed-item-time">'+(v.ts?timeAgo(v.ts):'')+'</span>'+
        '</div>'+
        (v.desc?'<div class="feed-item-body">'+escapeHtml(String(v.desc).substring(0,160))+'</div>':'')+
          (v.resource_id?'<div style="margin-top:4px"><button class="feed-resolve-btn" data-resource-type="'+escapeHtml(v.resource_type || '')+'" data-resource-id="'+escapeHtml(v.resource_id)+'"><i class="fas fa-search" style="margin-right:3px"></i>Resolve</button></div>':'')+
        '</div>';
    });
    feedEl.innerHTML=html;
      Array.prototype.forEach.call(feedEl.querySelectorAll('.feed-resolve-btn'), function(btn){
        btn.onclick=function(){
          window.LiveEngineV9.resolve(btn.getAttribute('data-resource-type') || '', btn.getAttribute('data-resource-id') || '');
        };
      });
  }

  // Public API
  window.LiveEngineV9={
    resolve:function(rType,rId){
      v9Toast('Resolving…',rType+' / '+String(rId).substring(0,14),'info',2500);
      proxyAction('webhook_resolve',{resource_type:rType,resource_id:rId}).then(function(data){
        if(data.ok&&data.record){
          var s=data.summary||{};
          showItemDetail((s.title||rType)+(s.reference?' #'+s.reference:''),
            [{section:'Resolved Record',icon:'fa-link'},{label:'Title',value:s.title||'—'},{label:'Status',value:s.status||'—'},{label:'UUID',value:rId||'—'}],null);
        }else{v9Toast('Could not resolve',data.error||'Record not found','warning');}
      }).catch(function(e){v9Toast('Resolve failed',e.message,'danger');});
    },
    clearUnseen:_clearUnseen,
    getEvents:function(){return _events;}
  };

  // Supersede old WebhookLive.clearUnseen so FAB/drawer buttons keep working
  if(window.WebhookLive)window.WebhookLive.clearUnseen=_clearUnseen;

  // Boot
  function _boot(){
    _seedLastId().then(function(){
      _pollTimer=setInterval(function(){if(!document.hidden)_poll();},POLL_MS);
      setTimeout(_poll,1600);
    });
    // Patch existing pollWebhookEvents to backfill decoded labels
    if(typeof pollWebhookEvents==='function'&&!pollWebhookEvents._v9patched){
      var _orig=pollWebhookEvents;
      pollWebhookEvents=async function(){
        var res=await _orig.apply(this,arguments);
        WEBHOOK_EVENTS.forEach(function(evt){
          if(!evt.event_label||evt.event_label===evt.type){
            try{var v=decodeWebhookEventV9(evt);if(v&&v.title){evt.event_label=v.title;evt.severity=v.severity;evt._icon=v.icon;evt._color=v.color;}}catch(e){}
          }
        });
        return res;
      };
      pollWebhookEvents._v9patched=true;
    }
    // Respond to invalidation events → refresh activity feed + dashboard KPIs
    window.addEventListener('handymgr:webhook-invalidate',function(e){
      requestAnimationFrame(function(){
        try{renderActivityFeed();}catch(err){}
        var woKeys=['work_orders','turn_work_orders','turns','upcoming_moveouts'];
        var detail=(e&&e.detail&&e.detail.keys)?e.detail.keys:[];
        if(detail.some(function(k){return woKeys.indexOf(k)!==-1;})){try{renderDashboardKPIs();}catch(err){}}
      });
    });
    document.addEventListener('visibilitychange',function(){if(!document.hidden&&_pollTimer)_poll();});
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',_boot);
  else _boot();
})();
