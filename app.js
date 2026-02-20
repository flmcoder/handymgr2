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
function showToast(msg) { var t = $('#toast'); $('#toastMsg').textContent = msg; t.style.display = 'block'; clearTimeout(t._tid); t._tid = setTimeout(function() { t.style.display = 'none'; }, 3500); }
function closeModal(id) { document.getElementById(id).classList.remove('show'); }
function openModal(id) { document.getElementById(id).classList.add('show'); }
function escapeHtml(s) { var d = document.createElement('div'); d.textContent = s || ''; return d.innerHTML; }
function loadingHtml(msg) { return '<div class="loading-overlay"><i class="fas fa-circle-notch"></i><p>' + escapeHtml(msg) + '</p></div>'; }
function emptyHtml(icon, msg) { return '<div class="empty-state"><i class="fas ' + icon + '"></i><p>' + escapeHtml(msg) + '</p></div>'; }

// ---- AppFolio deep link builder ----
// Builds URLs to view resources directly in AppFolio
function appfolioUrl(type, id) {
  if (!id) return '';
  var base = API_VHOST ? 'https://' + API_VHOST + '.appfolio.com' : '';
  if (!base) return '';
  switch (type) {
    case 'work_order': return base + '/tasks/work_orders/' + encodeURIComponent(id);
    case 'vendor': return base + '/vendor_details?vendor_id=' + encodeURIComponent(id);
    case 'property': return base + '/property_details?property_id=' + encodeURIComponent(id);
    case 'bill': return base + '/bills/' + encodeURIComponent(id);
    case 'unit_turn': return base + '/tasks/unit_turns/' + encodeURIComponent(id);
    case 'inspection': return base + '/tasks/unit_inspections/' + encodeURIComponent(id);
    case 'tenant': return base + '/tenant_details?occupancy_id=' + encodeURIComponent(id);
    default: return base + '/' + type + '/' + encodeURIComponent(id);
  }
}

// ---- Shared group filter logic ----
// Checks if a property (by ID or name) belongs to the given group
function isInPropertyGroup(propertyId, propertyName, groupName) {
  if (!groupName) return true; // no filter = show all
  // Check property.group (set by fetchPropertyGroups)
  var prop = PROPERTIES.find(function(p) {
    return p.name === propertyName || String(p.id) === String(propertyId);
  });
  if (prop && prop.group) return prop.group === groupName;
  if (prop && prop.portfolio) return prop.portfolio === groupName;
  // Direct check against PROPERTY_GROUPS
  return PROPERTY_GROUPS.some(function(g) {
    return g.name === groupName && Array.isArray(g.properties) && g.properties.some(function(pid) {
      return String(pid) === String(propertyId);
    });
  });
}

// Populates ALL group filter dropdowns across every tab
function populateGroupFilters() {
  var selectors = [
    '#woGroupFilter', '#payrollGroupFilter', '#dashGroupFilter',
    '#inspGroupFilter', '#vendorGroupFilter', '#reconGroupFilter',
    '#turnGroupFilter'
  ];
  selectors.forEach(function(sel) {
    var el = document.querySelector(sel);
    if (!el) return;
    var current = el.value; // preserve current selection
    // Clear existing options except the first (All Groups)
    while (el.options.length > 1) el.remove(1);
    if (PROPERTY_GROUPS.length > 0) {
      PROPERTY_GROUPS.forEach(function(g) {
        if (!g.name) return;
        var opt = document.createElement('option');
        opt.value = g.name; opt.textContent = g.name;
        el.appendChild(opt);
      });
    } else {
      // Fallback: use portfolio field from PROPERTIES
      var grps = {};
      PROPERTIES.forEach(function(p) { if (p.portfolio) grps[p.portfolio] = true; });
      Object.keys(grps).sort().forEach(function(g) {
        var opt = document.createElement('option');
        opt.value = g; opt.textContent = g;
        el.appendChild(opt);
      });
    }
    // Restore previous selection if still valid
    if (current) {
      for (var i = 0; i < el.options.length; i++) {
        if (el.options[i].value === current) { el.value = current; break; }
      }
    }
  });
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
function timeAgo(dateStr) {
  var d = new Date(dateStr);
  var diff = Math.floor((Date.now() - d.getTime()) / 1000);
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
var CACHE_DB_VERSION = 2;
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
      var store = tx.objectStore(CACHE_STORE);
      store.put({ key: key, data: data, timestamp: Date.now() });
      tx.oncomplete = function() { resolve(true); };
      tx.onerror = function() { resolve(false); };
    });
  }).catch(function() { return false; });
}

function cacheClearAll() {
  return openCacheDB().then(function(db) {
    return new Promise(function(resolve) {
      var tx = db.transaction(CACHE_STORE, 'readwrite');
      tx.objectStore(CACHE_STORE).clear();
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
      cacheSet('bills', BILLS),
      cacheSet('turns', TURNS),
      cacheSet('inspections', INSPECTIONS)
    ]);
    updateCacheBadge('live', Date.now());
    console.log('Cache saved: WO=' + WORK_ORDERS.length + ' V=' + VENDORS.length + ' P=' + PROPERTIES.length + ' B=' + BILLS.length + ' T=' + TURNS.length + ' I=' + INSPECTIONS.length);
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
      bills: BILLS.length,
      turns: TURNS.length,
      inspections: INSPECTIONS.length
    };
    var total = counts.work_orders + counts.vendors + counts.properties + counts.bills + counts.turns + counts.inspections;
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
      bills: BILLS,
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
    showToast('Exported ' + total + ' records (' + sizeKB + ' KB) \u2014 WO:' + counts.work_orders + ' V:' + counts.vendors + ' P:' + counts.properties + ' B:' + counts.bills + ' T:' + counts.turns + ' I:' + counts.inspections);
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
    BILLS = extractArr('bills');
    TURNS = extractArr('turns');
    INSPECTIONS = extractArr('inspections');
    var total = WORK_ORDERS.length + VENDORS.length + PROPERTIES.length + BILLS.length + TURNS.length + INSPECTIONS.length;
    // Persist to IndexedDB for future sessions
    await saveAllToCache();
    updateCacheBadge('cached', Date.now(), false);
    renderAll();
    showToast('Imported ' + total + ' records \u2014 WO:' + WORK_ORDERS.length + ' V:' + VENDORS.length + ' P:' + PROPERTIES.length + ' B:' + BILLS.length + ' T:' + TURNS.length + ' I:' + INSPECTIONS.length);
  } catch (e) {
    showToast('Import failed: ' + (e.message || e));
  }
}

function updateCacheBadge(state, timestamp, isStale) {
  var badge = $('#cacheBadge');
  var badgeText = $('#cacheBadgeText');
  var tsEl = $('#syncTimestamp');
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

/* =================================================================
   WO FLAGS — IndexedDB-persisted follow-up markers
   ================================================================= */
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
   CREDENTIAL VAULT — AES-256-GCM + PBKDF2
   ================================================================= */
var VAULT_BLOBS = [
  { // Passphrase: maint::cockpit
    s: 'oakh6uQFKiJWj95xXH/hJg==',
    i: 'pbU5H33tdyjncIza',
    t: '6jhJXkqaMFtyrXueMQbzhw==',
    c: '4EuCDYImFsxsjPQg65D/hYVji16JZjiEb7IeOf55tUlU9SbBIVNFhrtWS/MDDs2bthUGQl1xQwg6Ds9fX3dW2psUvyMM8FeD62BV7oq9r4ItZHk7Yz/29AuROg/MECEbZhRyzRGRt21c5PNJ1oFih9aR/QmmXKkRo8wvm99Yn+ODsFyCHC15EsFIOzmA288qwA=='
  },
  { // Passphrase: handy::manager
    s: 'TteV8E8jlLfolw5t39vUiA==',
    i: 'YJYvbSiwahNcaEFW',
    t: 'W4kEMK0/WO8kRJPII6RJ3g==',
    c: '8kD52jM1FFDqORA4amjV5cJgRp6LICgYBqgeP9m7o+IX8XAkxwfa4pFxxU+6Y06xNkeqlL2LZbvYhv4chkydPONGxnKMvtuPurDg43L2QPAf1decHbgWvkcPCDuHzh/mOHH26pdAQjVvrt+RkGqawcpEjI//HsXA/NnUwsHYzU6gL3l1Q+HnZlzu8a+wU8Cnaw=='
  }
];

var API_CREDS = null;
var API_VHOST = null;
var API_PROXY = '';
var VAULT_TIMEOUT_MS = 60 * 60 * 1000; // 60 minutes
var vaultTimeoutId = null;
var vaultCountdownId = null;

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
  // Try each vault blob — supports multiple passphrases
  for (var i = 0; i < VAULT_BLOBS.length; i++) {
    try {
      return await decryptVaultBlob(VAULT_BLOBS[i], passphrase);
    } catch (e) { /* try next blob */ }
  }
  throw new Error('Decryption failed for all vault blobs');
}

function wipeCredentials() {
  if (API_CREDS) {
    if (API_CREDS.a) { API_CREDS.a = '0'.repeat(API_CREDS.a.length); }
    if (API_CREDS.d) { API_CREDS.d = '0'.repeat(API_CREDS.d.length); }
  }
  API_CREDS = null;
  API_VHOST = null;
  API_PROXY = '';
}

function lockVault() {
  wipeCredentials();
  clearTimeout(vaultTimeoutId);
  clearInterval(vaultCountdownId);
  vaultTimeoutId = null;
  vaultCountdownId = null;
  appInitialized = false;
  WORK_ORDERS = []; VENDORS = []; BILLS = []; PROPERTIES = []; PROPERTY_GROUPS = []; TURNS = []; INSPECTIONS = []; RECENT_TASKS = []; WEBHOOK_EVENTS = []; TURN_RECORDS = []; TURN_PIPE_DATA = []; API_ERRORS = [];
  if (_webhookPollTimer) { clearInterval(_webhookPollTimer); _webhookPollTimer = null; }
  // Note: IndexedDB cache is NOT cleared on lock — data persists for next unlock
  updateCacheBadge('offline');
  $('#appShell').classList.remove('unlocked');
  $('#vaultScreen').style.display = 'flex';
  $('#vaultPassphrase').value = '';
  $('#vaultError').classList.remove('show');
  $('#corsBanner').classList.remove('show');
  $('#vaultPassphrase').focus();
}

function resetInactivityTimer() {
  if (!API_CREDS) return;
  clearTimeout(vaultTimeoutId);
  vaultTimeoutId = setTimeout(async function() {
    try { await saveAllToCache(); } catch(e) { /* best-effort */ }
    lockVault();
    showToast('Vault auto-locked after 60 minutes \u2014 data saved to cache');
  }, VAULT_TIMEOUT_MS);
  startCountdown();
}

function startCountdown() {
  clearInterval(vaultCountdownId);
  var targetTime = Date.now() + VAULT_TIMEOUT_MS;
  vaultCountdownId = setInterval(function() {
    var remaining = Math.max(0, targetTime - Date.now());
    if (remaining <= 0) { clearInterval(vaultCountdownId); return; }
    var mins = Math.floor(remaining / 60000);
    var secs = Math.floor((remaining % 60000) / 1000);
    var el = $('#vaultCountdown');
    if (el) { el.textContent = mins + ':' + (secs < 10 ? '0' : '') + secs; }
  }, 1000);
}

function getAuthHeader() { return API_CREDS ? API_CREDS.a : null; }
function getDevId() { return API_CREDS ? API_CREDS.d : null; }
function getDirectBaseUrl() { return API_VHOST ? 'https://' + API_VHOST + '.appfolio.com' : null; }

// ---- Timeout helper ----
// Wraps a promise with an AbortController-based timeout (defaults 45s)
function fetchWithTimeout(url, opts, timeoutMs) {
  timeoutMs = timeoutMs || 45000;
  var controller = new AbortController();
  var timer = setTimeout(function() { controller.abort(); }, timeoutMs);
  var fetchOpts = Object.assign({}, opts || {}, { signal: controller.signal });
  return fetch(url, fetchOpts).finally(function() { clearTimeout(timer); });
}

// ---- Proxy v6 action endpoint caller ----
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
  var res;
  try {
    res = await fetchWithTimeout(url, { headers: { 'Accept': 'application/json' } }, 45000);
  } catch (abortErr) {
    if (abortErr.name === 'AbortError') {
      var tmsg = 'Proxy action=' + action + ' timed out after 45s';
      logApiError(0, tmsg, 'queued');
      throw new Error(tmsg);
    }
    throw abortErr;
  }
  if (!res.ok) {
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
    var msg = 'Proxy action=' + action + ': ' + (data.error || 'Unknown error');
    logApiError(502, msg, 'queued');
    throw new Error(msg);
  }
  return data;
}

// Resolve a path to a fetchable URL.
// When a proxy is active, the proxy has the domain + credentials hardcoded
// server-side, so we only send the API path (e.g. /api/v0/properties).
// Used for raw pass-through calls (PATCH work orders, POST notes, etc.)
function resolveUrl(path) {
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

  var url = resolveUrl(path);
  var hdrs = {};
  if (API_PROXY) {
    // Server-side proxy has credentials hardcoded — no auth headers needed
    hdrs['Content-Type'] = 'application/json';
  } else {
    hdrs['Authorization'] = auth;
    hdrs['X-AppFolio-Developer-ID'] = devId;
    hdrs['Content-Type'] = 'application/json';
  }
  var opts = Object.assign({ method: 'GET', headers: hdrs }, options || {});

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
        logApiError(401, 'Unauthorized — proxy may have wrong credentials. Verify the v6 proxy has correct Client ID, Secret, and Developer ID.', 'resolved');
        showCorsError('401 Unauthorized — AppFolio rejected the credentials. Verify your v6 proxy has the correct Client ID, Client Secret, and Developer ID hardcoded.');
        throw new Error('401: Unauthorized — check proxy credentials');
      }
      if (res.status === 404) {
        logApiError(404, 'Not Found — endpoint does not exist or credentials are invalid. URL: ' + path, 'resolved');
        showCorsError('404 Not Found — The endpoint may not exist or credentials may be rejected. Verify your v6 proxy configuration.');
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
    if (msgEl) { msgEl.innerHTML = '<strong style="color:var(--danger)">AppFolio rejected credentials.</strong> Make sure you deployed the <strong>v6 proxy code</strong> with correct Client ID, Secret, and Developer ID hardcoded inside. Visit your proxy URL in a browser \u2014 you should see a JSON response with <code>"proxy": "v6"</code>.'; }
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
var BILLS = [];
var PROPERTIES = [];
var PROPERTY_GROUPS = [];
var TURNS = [];
var UPCOMING_MOVEOUTS = []; // from tenant_directory — tenants on notice
var TURN_WORK_ORDERS = []; // from DB API — unit turn WOs with real-time status
var RECENT_TASKS = [];
var WEBHOOK_EVENTS = [];
var _webhookPollTimer = null;
var appInitialized = false;
var WO_FLAGS = {};
var WO_DETAIL_CACHE = {};
var PAYROLL_WEEK_OFFSET = 0;
var currentPropertyGroup = '';
var currentTurnFilter = 'open';

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

  if (!pass) {
    $('#vaultError').textContent = 'Passphrase cannot be empty.';
    $('#vaultError').classList.add('show');
    return;
  }
  if (!vhost) {
    $('#vaultError').textContent = 'AppFolio subdomain is required. Enter just the subdomain (e.g. "flraz"), not the full URL.';
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
    API_CREDS = await decryptVault(pass);
    API_VHOST = vhost;
    API_PROXY = proxyUrl;
    $('#vaultPassphrase').value = '';
    $('#vaultScreen').style.display = 'none';
    $('#appShell').classList.add('unlocked');
    resetInactivityTimer();
    initApp();
    var proxyInfo = API_PROXY ? ' via proxy' : ' (direct)';
    showToast('Vault unlocked \u2014 connecting to ' + vhost + '.appfolio.com' + proxyInfo);
  } catch (err) {
    wipeCredentials();
    $('#vaultError').textContent = 'Decryption failed \u2014 incorrect passphrase or corrupted vault.';
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

['click', 'keydown', 'mousemove', 'touchstart', 'scroll'].forEach(function(evt) {
  document.addEventListener(evt, function() { if (API_CREDS) { resetInactivityTimer(); } }, { passive: true });
});

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

function dateNDaysAgo(n) {
  var d = new Date();
  d.setDate(d.getDate() - n);
  return d.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

// Open WO status codes for Reports API (excludes 4=Completed, 5=Canceled, 7=CompletedNoNeedToBill)
// Defined in fetchWorkOrders() as OPEN_WO_STATUS_CODES

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

// Work Orders: Proxy v6 ?action=work_orders — server-side pagination, one request
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
        id: r.work_order_number || r.service_request_number || '',
        uuid: r.work_order_id || '',
        propertyId: r.property_id || '',
        propertyName: r.property_name || r.property || '',
        propertyAddress: ((r.property_street || '') + ' ' + (r.property_city || '') + ' ' + (r.property_state || '') + ' ' + (r.property_zip || '')).trim(),
        unitId: r.unit_id || '',
        unit: r.unit_name || r.unit_id || '',
        priority: r.priority || 'Normal',
        status: r.status || 'New',
        description: r.job_description || r.service_request_description || '',
        vendorName: r.vendor || '',
        vendorId: r.vendor_id || '',
        vendorTrade: r.vendor_trade || '',
        created: r.created_at || '',
        updated: r.completed_on || r.created_at || '',
        completedOn: r.completed_on || '',
        workCompletedOn: r.work_completed_on || '',
        scheduledStart: r.scheduled_start || '',
        scheduledEnd: r.scheduled_end || '',
        type: r.work_order_type || '',
        amount: r.amount || '',
        tenant: r.primary_tenant || '',
        tenantEmail: r.primary_tenant_email || '',
        tenantPhone: r.primary_tenant_phone_number || '',
        createdBy: r.created_by || '',
        assignedUser: r.assigned_user || '',
        statusNotes: r.status_notes || '',
        maintenanceLimit: r.maintenance_limit || '',
        link: ''
      };
    });
    setApiStatus('loading', 'Work orders: ' + WORK_ORDERS.length + ' loaded');
    return true;
  } catch (err) {
    WORK_ORDERS = [];
    return false;
  }
}

// Vendors: Proxy v6 ?action=vendors — server-side pagination, one request
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

// Bills: Proxy v6 ?action=bills — server-side pagination, one request
var MAX_BILL_RECORDS = 200;
async function fetchBills() {
  try {
    setApiStatus('loading', 'Loading bills (server-side)\u2026');
    var data = await proxyAction('bills', { days: DATA_WINDOW_DAYS, max: MAX_BILL_RECORDS });
    var results = data.results || [];
    BILLS = results.map(function(b) {
      var vendName = '';
      if (b.VendorId) {
        var vend = VENDORS.find(function(v) { return v.id === b.VendorId; });
        if (vend) { vendName = vend.name; }
      }
      var propName = '';
      var propAddr = '';
      var propId = b.PropertyId || '';
      if (!propId && b.LineItems && b.LineItems.length > 0 && b.LineItems[0].PropertyId) {
        propId = b.LineItems[0].PropertyId;
      }
      if (propId) {
        var prop = PROPERTIES.find(function(p) { return p.id === propId; });
        if (prop) { propName = prop.name; propAddr = prop.address; }
      }
      return {
        id: b.Id || '',
        vendorName: vendName,
        vendorId: b.VendorId || '',
        propertyName: propName,
        propertyId: propId,
        propertyAddress: propAddr,
        amount: parseFloat(b.TotalAmount || '0') || 0,
        approvalStatus: b.ApprovalStatus || '',
        reference: b.Reference || '',
        description: b.Description || b.CheckMemo || '',
        dueDate: b.DueDate || '',
        invoiceDate: b.InvoiceDate || '',
        workOrderId: b.WorkOrderId || null,
        lineItems: b.LineItems || []
      };
    });
    return true;
  } catch (err) {
    BILLS = [];
    return false;
  }
}

// Properties: Proxy v6 ?action=properties — server-side pagination, one request
async function fetchProperties() {
  try {
    setApiStatus('loading', 'Loading properties (server-side)\u2026');
    var data = await proxyAction('properties');
    var results = data.results || [];
    PROPERTIES = results.map(function(p) {
      return {
        id: p.property_id || '',
        name: p.property_name || p.property || '',
        address: ((p.property_street || '') + (p.property_street2 ? ' ' + p.property_street2 : '')).trim(),
        city: p.property_city || '',
        state: p.property_state || '',
        zip: p.property_zip || '',
        propertyType: p.property_type || '',
        portfolioId: p.portfolio_id || '',
        portfolio: p.portfolio || '',
        maintenanceLimit: p.maintenance_limit || '',
        maintenanceNotes: p.maintenance_notes || '',
        siteManager: p.site_manager || '',
        units: p.units || '',
        sqft: p.sqft || '',
        marketRent: p.market_rent || '',
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

// Turns: Proxy v6 ?action=turns — 60-day window, In Progress only by default
async function fetchTurns() {
  try {
    setApiStatus('loading', 'Loading turns (In Progress, 60d)\u2026');
    var data = await proxyAction('turns', { days: 60, status: 'In Progress' });
    var results = data.results || [];
    TURNS = results.map(function(t) {
      var moveOut = t.move_out_date || '';
      var turnEnd = t.turn_end_date || '';
      var daysToComplete = parseInt(t.total_days_to_complete || 0, 10) || 0;
      // If no total_days_to_complete and we have dates, calculate
      if (!daysToComplete && moveOut && turnEnd) {
        var d1 = new Date(moveOut), d2 = new Date(turnEnd);
        if (!isNaN(d1) && !isNaN(d2)) { daysToComplete = Math.round((d2 - d1) / 86400000); }
      }
      return {
        unitTurnId: t.unit_turn_id || '',
        unit: t.unit || '',
        property: t.property || '',
        propertyId: t.property_id || 0,
        unitId: t.unit_id || 0,
        notes: t.notes || '',
        referenceUser: t.reference_user || '',
        moveOut: moveOut,
        turnEnd: turnEnd,
        expectedMoveIn: t.expected_move_in_date || '',
        targetDays: parseInt(t.target_days_to_complete || 0, 10) || 0,
        totalDays: daysToComplete,
        laborCost: t.labor_from_work_orders || '$0.00',
        purchaseOrders: t.purchase_orders_from_work_orders || '$0.00',
        billables: t.billables_from_work_orders || '$0.00',
        inventory: t.inventory_from_work_orders || '$0.00',
        totalBilled: t.total_billed || '$0.00'
      };
    });
    return true;
  } catch (err) {
    TURNS = [];
    return false;
  }
}

// Inspections: Proxy v6 ?action=inspections — server-side pagination, one request
var INSPECTIONS = [];
async function fetchInspections() {
  try {
    setApiStatus('loading', 'Loading inspections (server-side)\u2026');
    var data = await proxyAction('inspections', { days: DATA_WINDOW_DAYS });
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
    });
    return true;
  } catch (err) {
    INSPECTIONS = [];
    return false;
  }
}

// Property Groups: Proxy v6 ?action=property_groups — DB API v0
// DB API returns { data: [ { Id, Name, PropertyIds, Type, LastUpdatedAt } ] }
async function fetchPropertyGroups() {
  try {
    setApiStatus('loading', 'Loading property groups\u2026');
    var data = await proxyAction('property_groups');
    var results = data.results || data.data || [];
    PROPERTY_GROUPS = results.map(function(g) {
      return {
        id: g.Id || g.id || '',
        name: g.Name || g.name || '',
        properties: g.PropertyIds || g.Properties || g.properties || g.property_ids || []
      };
    });
    // Build a lookup: property ID -> group name
    PROPERTY_GROUPS.forEach(function(g) {
      if (Array.isArray(g.properties)) {
        g.properties.forEach(function(pid) {
          // Tag the matching PROPERTIES entry with its group
          // DB API uses UUIDs, Reports API uses numeric IDs — match both
          var prop = PROPERTIES.find(function(p) {
            return String(p.id) === String(pid) || String(p.uuid) === String(pid);
          });
          if (prop) { prop.group = g.name; }
        });
      }
    });
    return true;
  } catch (err) {
    console.log('fetchPropertyGroups error: ' + (err.message || err));
    PROPERTY_GROUPS = [];
    return false;
  }
}

// Recent Tasks: Proxy v6 ?action=recent_tasks — server-side pagination via DB API v0
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

// Upcoming Move-Outs: Proxy v6 ?action=upcoming_moveouts — tenant directory
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

// Turn Work Orders: Proxy v6 ?action=turn_work_orders — DB API v0
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

/* =================================================================
   WEBHOOK EVENTS — HTTP POST relay for Make.com / Zapier / etc.
   ================================================================= */
async function pollWebhookEvents() {
  try {
    var data = await proxyAction('webhook_events');
    if (data && Array.isArray(data.events)) {
      // Merge new events (dedup by timestamp + title)
      var existing = {};
      WEBHOOK_EVENTS.forEach(function(e) { existing[e.ts + '|' + e.title] = true; });
      data.events.forEach(function(e) {
        var key = (e.ts || e.timestamp || '') + '|' + (e.title || '');
        if (!existing[key]) {
          WEBHOOK_EVENTS.push({
            ts: e.ts || e.timestamp || new Date().toISOString(),
            type: e.type || 'webhook',
            title: e.title || 'Webhook Event',
            body: e.body || '',
            priority: e.priority || 'normal',
            source: e.source || 'webhook'
          });
          existing[key] = true;
        }
      });
      // Sort newest first
      WEBHOOK_EVENTS.sort(function(a, b) { return new Date(b.ts) - new Date(a.ts); });
      // Keep last 200
      if (WEBHOOK_EVENTS.length > 200) WEBHOOK_EVENTS = WEBHOOK_EVENTS.slice(0, 200);
    }
    return true;
  } catch (err) {
    console.log('Webhook poll error: ' + (err.message || err));
    return false;
  }
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
  WEBHOOK_EVENTS.slice(0, 20).forEach(function(e) {
    var isPri = e.priority === 'urgent' || e.priority === 'high';
    html += '<div style="padding:4px 0;border-bottom:1px solid var(--border)">';
    html += '<span style="color:var(--text-muted)">' + escapeHtml(e.ts ? timeAgo(e.ts) : '\u2014') + '</span> ';
    if (isPri) html += '<span style="color:var(--danger);font-weight:600">\u26a0 </span>';
    html += '<strong style="color:var(--text-primary)">' + escapeHtml(e.title) + '</strong>';
    if (e.body) html += '<div style="color:var(--text-secondary);margin-top:2px">' + escapeHtml(e.body.substring(0, 120)) + '</div>';
    html += '</div>';
  });
  if (WEBHOOK_EVENTS.length > 20) {
    html += '<div style="padding:4px 0;color:var(--text-muted);text-align:center">\u2026 and ' + (WEBHOOK_EVENTS.length - 20) + ' more</div>';
  }
  el.innerHTML = html;
}

function setupWebhookAutoPoll(intervalSec) {
  if (_webhookPollTimer) { clearInterval(_webhookPollTimer); _webhookPollTimer = null; }
  if (intervalSec > 0) {
    _webhookPollTimer = setInterval(async function() {
      var ok = await pollWebhookEvents();
      if (ok && WEBHOOK_EVENTS.length > 0) {
        renderWebhookEventList();
        renderActivityFeed();
      }
    }, intervalSec * 1000);
  }
}

/* =================================================================
   DETAIL FETCHERS — On-demand for WO detail panel
   ================================================================= */
async function fetchPropertyDetail(propId) {
  if (!propId) return null;
  if (WO_DETAIL_CACHE['prop_' + propId]) return WO_DETAIL_CACHE['prop_' + propId];
  try {
    var data = await apiFetch('/api/v0/properties/' + propId);
    WO_DETAIL_CACHE['prop_' + propId] = data;
    return data;
  } catch (e) { return null; }
}

async function fetchWONotes(uuid) {
  if (!uuid) return [];
  if (WO_DETAIL_CACHE['notes_' + uuid]) return WO_DETAIL_CACHE['notes_' + uuid];
  try {
    var data = await apiFetch('/api/v0/work_orders/' + uuid + '/notes');
    var notes = (data && data.data) ? data.data : (Array.isArray(data) ? data : []);
    WO_DETAIL_CACHE['notes_' + uuid] = notes;
    return notes;
  } catch (e) { return []; }
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

  // Sync payroll group dropdown with global filter
  var pgSel = $('#payrollGroupFilter');
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
    var afUrl = appfolioUrl('work_order', wo.uuid);
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

  // Click row to show detail card
  body.querySelectorAll('.payroll-row').forEach(function(row) {
    row.addEventListener('click', function(e) {
      if (e.target.closest('[data-flagwo]')) return; // don't open detail on flag click
      var woid = this.getAttribute('data-woid');
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
      ], appfolioUrl('work_order', wo.uuid));
    });
  });

  body.querySelectorAll('[data-flagwo]').forEach(function(btn) {
    btn.addEventListener('click', async function(e) {
      e.stopPropagation();
      var wid = this.getAttribute('data-flagwo');
      await toggleFlag(wid);
      renderPayroll();
      renderWorkOrders();
    });
  });
}

function renderDashboardKPIs() {
  // Sync dashboard group filter
  var dashGrpSel = $('#dashGroupFilter');
  if (dashGrpSel && dashGrpSel.value !== currentPropertyGroup) dashGrpSel.value = currentPropertyGroup;

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

  $('#kpiOpen').textContent = openWOs.length;
  $('#kpiOpenSub').textContent = WORK_ORDERS.length + ' loaded (open, ' + DATA_WINDOW_DAYS + 'd)';
  $('#kpiUrgent').textContent = urgentWOs.length;
  $('#kpiUrgentSub').textContent = urgentWOs.length > 0 ? urgentWOs.length + ' require immediate attention' : 'No urgent items';
  $('#kpiTurns').textContent = activeTurns.length;
  $('#kpiTurnsSub').textContent = TURNS.length + ' total turns';
  $('#kpiMoveOuts').textContent = moveOuts.length;
  $('#kpiMoveOutsSub').textContent = moveOuts.length > 0 ? moveOuts[0].daysLeft + 'd until next' : 'None in 60 days';
  $('#kpiFlagged').textContent = flaggedCount;
  $('#kpiFlaggedSub').textContent = flaggedCount > 0 ? flaggedCount + ' items flagged' : 'No flagged items';

  $('#woBadge').textContent = openWOs.length || '0';
  $('#turnBadge').textContent = activeTurns.length || '0';
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

  // === Webhook events — from HTTP POST relay ===
  WEBHOOK_EVENTS.slice(0, 20).forEach(function(wh) {
    var isPri = wh.priority === 'urgent' || wh.priority === 'high';
    activities.push({
      sortDate: new Date(wh.ts || 0).getTime(),
      time: wh.ts ? timeAgo(wh.ts) : '\u2014',
      type: 'work_order',
      urgent: isPri,
      icon: isPri ? 'fa-exclamation-circle' : 'fa-plug',
      iconColor: isPri ? 'var(--danger)' : 'var(--purple)',
      event: '<span class="tag ' + (isPri ? 'urgent' : 'new') + '">' + escapeHtml(wh.type || 'Webhook') + '</span>',
      entity: escapeHtml(wh.source || 'webhook'),
      detail: escapeHtml(wh.title || '') + (wh.body ? ' \u2014 ' + escapeHtml(wh.body.substring(0, 80)) : ''),
      extra: '<span style="color:var(--purple)"><i class="fas fa-plug" style="font-size:9px"></i> via webhook</span>'
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

  // Group dropdown populated by populateGroupFilters() — just sync value
  var grpSel = $('#woGroupFilter');
  if (grpSel && grpSel.value !== currentPropertyGroup) grpSel.value = currentPropertyGroup;

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
      if (wo.vendorName) html += '<span><i class="fas fa-hard-hat"></i> ' + escapeHtml(wo.vendorName) + '</span>';
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

  // Wire up card clicks → detail modal
  board.querySelectorAll('.kanban-card').forEach(function(card) {
    card.addEventListener('click', function() {
      showWODetail(this.getAttribute('data-woid'));
    });
  });

  // Wire up column header clicks → filter by status
  board.querySelectorAll('.kanban-col-head').forEach(function(head) {
    head.addEventListener('click', function() {
      var status = this.getAttribute('data-status');
      if (currentWOFilter === status) {
        currentWOFilter = 'all'; // toggle off
      } else {
        currentWOFilter = status;
      }
      // Update filter buttons
      $$('[data-filter]').forEach(function(b) {
        b.classList.toggle('active', b.getAttribute('data-filter') === currentWOFilter);
      });
      renderWorkOrders();
    });
  });
}

function showWODetail(id) {
  var wo = WORK_ORDERS.find(function(w) { return String(w.id) === String(id); });
  if (!wo) return;
  var woAfUrl = appfolioUrl('work_order', wo.uuid);
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

  // Async: fetch notes
  fetchWONotes(wo.uuid).then(function(notes) {
    var nl = document.getElementById('detailNotesList');
    if (!nl) return;
    if (!notes || notes.length === 0) {
      nl.innerHTML = '<div style="text-align:center;padding:10px;color:var(--text-muted);font-size:12px">No notes</div>';
      return;
    }
    var nh = '';
    notes.forEach(function(n) {
      nh += '<div class="note-item"><div class="note-item-header"><span>' + escapeHtml(n.CreatedBy || n.created_by || '') + '</span><span>' + formatDate(n.CreatedAt || n.created_at || '') + '</span></div>';
      nh += '<div class="note-item-body">' + escapeHtml(n.Content || n.content || '') + '</div></div>';
    });
    nl.innerHTML = nh;
  });

  $('#woModalSave').onclick = async function() {
    var newStatus = $('#detailStatus').value;
    var newPriority = $('#detailPriority').value;
    var note = ($('#detailNote') && $('#detailNote').value) ? $('#detailNote').value.trim() : '';

    try {
      if (newStatus !== wo.status || newPriority !== wo.priority) {
        await apiFetch('/api/v0/work_orders/' + wo.uuid, {
          method: 'PATCH',
          body: JSON.stringify({ Status: newStatus, Priority: newPriority })
        });
        wo.status = newStatus;
        wo.priority = newPriority;
      }
      if (note) {
        try {
          await apiFetch('/api/v0/work_orders/' + wo.uuid + '/notes', {
            method: 'POST',
            body: JSON.stringify({ Content: note })
          });
          // Clear notes cache so next open re-fetches
          delete WO_DETAIL_CACHE['notes_' + wo.uuid];
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
var currentTurnPipeFilter = 'active';
var currentTurnPipeGroup = '';

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

  // Helper: create a composite key for deduplication
  function makeKey(propId, unitId, moveOut) {
    var k = String(propId || '') + '-' + String(unitId || '') + '-' + (moveOut || '');
    return k === '--' ? null : k;
  }

  // Helper: find matching WOs for a unit+property pair (from both Reports + DB API data)
  function findMatchingWOs(unit, property, propId, unitId, moveOutDate) {
    var wos = [];
    // From Reports API work orders (WORK_ORDERS)
    WORK_ORDERS.forEach(function(wo) {
      var unitMatch = wo.unit && unit && String(wo.unit).toLowerCase() === String(unit).toLowerCase();
      var propMatch = wo.propertyName && property && String(wo.propertyName).toLowerCase() === String(property).toLowerCase();
      if (!unitMatch || !propMatch) return;
      if (moveOutDate && wo.created) {
        var daysDiff = (new Date(wo.created) - moveOutDate) / 86400000;
        if (daysDiff < -30) return;
      }
      wos.push({ source: 'reports', id: wo.id, status: wo.status, description: wo.description || '',
        created: wo.created, vendor: wo.vendor || '', unit: wo.unit, property: wo.propertyName, priority: wo.priority });
    });
    // From DB API turn work orders (TURN_WORK_ORDERS) — more current status
    TURN_WORK_ORDERS.forEach(function(wo) {
      var idMatch = (unitId && wo.unitId && String(wo.unitId) === String(unitId)) ||
                    (propId && wo.propertyId && String(wo.propertyId) === String(propId));
      if (!idMatch) return;
      // Skip if already found from Reports
      var dupe = wos.find(function(w) { return String(w.id) === String(wo.id) || String(w.id) === String(wo.woNumber); });
      if (dupe) {
        // DB API has more current status — update it
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
    var hasWorkDone = woStatuses.some(function(s) {
      return s === 'Work Done' || s === 'Work Completed' || s === 'Ready to Bill' || s === 'Completed';
    });

    // Progressive — later stages imply earlier ones
    stages.wo_created = { done: hasWO, date: woCreatedDate, woIds: matchingWOs.map(function(w) { return w.id; }) };
    stages.est_requested = { done: hasEstReq || hasEstimated || hasAssigned || hasWorkDone, date: null };
    stages.est_received = { done: hasEstimated || hasAssigned || hasWorkDone, date: null, vendors: [] };
    stages.assigned = { done: hasAssigned || hasWorkDone, date: null };
    stages.work_done = { done: hasWorkDone, date: null };

    return stages;
  }

  // Helper: build a pipeline entry
  function addEntry(key, unit, property, propId, unitId, moveOut, turnData, moveoutTenant) {
    if (!key || seenKeys[key]) return;
    seenKeys[key] = true;

    var moveOutDate = moveOut ? new Date(moveOut) : null;
    var isUpcoming = moveOutDate ? moveOutDate > today : false;
    var matchingWOs = findMatchingWOs(unit, property, propId, unitId, moveOutDate);
    var matchingInsp = findMatchingInsp(unit, property);
    var stages = deriveStages(moveOut, matchingWOs, matchingInsp, isUpcoming);

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

    // Compute current stage index (highest completed)
    var currentStageIdx = -1;
    PIPE_STAGES.forEach(function(ps, i) {
      if (stages[ps.key] && stages[ps.key].done) currentStageIdx = i;
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
    var target = (turnData && turnData.targetDays) || 30;
    var isStalled = !isUpcoming && elapsed > 7 && currentStageIdx >= 1 && currentStageIdx < PIPE_STAGES.length - 1;
    var isCompleted = (turnData && !!turnData.turnEnd) || (stages.work_done && stages.work_done.done);

    // Parse cost
    var costNum = 0;
    if (turnData && turnData.totalBilled) {
      costNum = parseFloat(String(turnData.totalBilled).replace(/[^0-9.\-]/g, '')) || 0;
    }

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
      stages: stages,
      currentStageIdx: currentStageIdx,
      matchingWOs: matchingWOs,
      matchingInsp: matchingInsp,
      webhookEvents: webhookMatches,
      elapsed: elapsed,
      target: target,
      isStalled: isStalled && !isCompleted,
      isCompleted: isCompleted,
      costNum: costNum,
      totalBilled: (turnData && turnData.totalBilled) || '$0',
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

  // Sort: upcoming first (by days-until), then active (by elapsed desc), then completed
  TURN_PIPE_DATA.sort(function(a, b) {
    if (a.isCompleted !== b.isCompleted) return a.isCompleted ? 1 : -1;
    if (a.isUpcoming !== b.isUpcoming) return a.isUpcoming ? -1 : 1;
    if (a.isUpcoming && b.isUpcoming) return a.elapsed - b.elapsed; // days-until ascending
    return b.elapsed - a.elapsed; // elapsed descending for active
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
  var active = TURN_PIPE_DATA.filter(function(p) { return !p.isCompleted && !p.isUpcoming; });
  var upcoming = TURN_PIPE_DATA.filter(function(p) { return p.isUpcoming; });
  var awaitEst = active.filter(function(p) {
    return p.stages.wo_created && p.stages.wo_created.done && p.stages.est_received && !p.stages.est_received.done;
  });
  var totalBilled = 0;
  active.forEach(function(p) { totalBilled += p.costNum; });
  var avgDays = 0;
  if (active.length > 0) {
    var totalDays = 0;
    active.forEach(function(p) { totalDays += Math.abs(p.elapsed); });
    avgDays = Math.round(totalDays / active.length);
  }

  var e = function(id, v) { var el = document.getElementById(id); if (el) el.textContent = v; };
  e('kpiActiveTurns', active.length);
  e('kpiActiveTurnsSub', active.length + ' active' + (upcoming.length > 0 ? ', ' + upcoming.length + ' upcoming' : ''));
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

  // Sync group filter dropdown
  var turnGrpSel = $('#turnGroupFilter');
  if (turnGrpSel && turnGrpSel.value !== currentPropertyGroup) turnGrpSel.value = currentPropertyGroup;

  var filtered = TURN_PIPE_DATA.filter(function(p) {
    if (filter === 'active' && p.isCompleted) return false;
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
  filtered.forEach(function(p, idx) {
    var cardClass = p.isCompleted ? '' : p.isUpcoming ? 'upcoming' : p.isStalled ? 'stalled' : p.elapsed < 14 ? 'on-track' : 'waiting';
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
      if (stage.done) {
        dotClass = 'done';
      } else if (si === p.currentStageIdx + 1) {
        dotClass = p.isStalled ? 'warn' : 'active';
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
    } else {
      var eColor = p.elapsed > p.target ? 'var(--danger)' : p.elapsed > 14 ? 'var(--warning)' : 'var(--text-primary)';
      html += '<span class="pipe-card-elapsed" style="color:' + eColor + '">' + p.elapsed + 'd</span>';
      var nextStage = p.currentStageIdx < PIPE_STAGES.length - 1 ? PIPE_STAGES[p.currentStageIdx + 1] : null;
      html += '<span class="pipe-card-cost">' + escapeHtml(p.totalBilled);
      if (nextStage) html += ' &bull; Next: ' + nextStage.label;
      html += '</span>';
    }
    html += '</div>';

    html += '</div>'; // end pipe-card

    // Detail panel (hidden by default)
    html += '<div class="pipe-detail" id="pipeDetail_' + idx + '">';
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
    html += '<div class="detail-section-title"><i class="fas fa-wrench"></i> Linked Work Orders (' + p.matchingWOs.length + ')</div>';
    if (p.matchingWOs.length > 0) {
      html += '<div class="pipe-wo-list">';
      p.matchingWOs.forEach(function(wo) {
        var woLink = appfolioUrl('work_order', wo.uuid || wo.link);
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
    if (p.tenant) html += '<div class="detail-row"><div class="detail-row-label">Tenant</div><div class="detail-row-value">' + escapeHtml(p.tenant) + '</div></div>';
    if (p.turn) {
      html += '<div class="detail-row"><div class="detail-row-label">Expected Move-In</div><div class="detail-row-value">' + (p.turn.expectedMoveIn ? formatDate(p.turn.expectedMoveIn) : '\u2014') + '</div></div>';
      html += '<div class="detail-row"><div class="detail-row-label">Target Days</div><div class="detail-row-value">' + (p.target || '\u2014') + '</div></div>';
      html += '<div class="detail-row"><div class="detail-row-label">Total Billed</div><div class="detail-row-value">' + escapeHtml(p.totalBilled) + '</div></div>';
      html += '<div class="detail-row"><div class="detail-row-label">Labor</div><div class="detail-row-value">' + escapeHtml(p.turn.laborCost || '$0') + '</div></div>';
      html += '<div class="detail-row"><div class="detail-row-label">Reference</div><div class="detail-row-value">' + escapeHtml(p.turn.referenceUser || '\u2014') + '</div></div>';
    } else {
      html += '<div class="detail-row"><div class="detail-row-label">Status</div><div class="detail-row-value" style="color:var(--info,#60a5fa)">Upcoming — awaiting move-out</div></div>';
      html += '<div class="detail-row"><div class="detail-row-label">Days Until</div><div class="detail-row-value">' + Math.abs(p.elapsed) + ' days</div></div>';
    }
    html += '</div>';

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
    // AppFolio deep link — use turn ID if available, else property link
    var turnAfUrl = p.turn ? appfolioUrl('unit_turn', p.turn.turnId || p.turn.id) : appfolioUrl('property', p.propertyId);
    if (turnAfUrl) {
      html += '<a class="action-btn" href="' + escapeHtml(turnAfUrl) + '" target="_blank" rel="noopener noreferrer" style="text-decoration:none" onclick="event.stopPropagation()"><i class="fas fa-external-link-alt"></i> View in AppFolio</a>';
    }
    html += '<button class="action-btn" onclick="this.closest(\'.pipe-detail\').classList.remove(\'show\')"><i class="fas fa-times"></i> Close</button>';
    html += '</div>';

    html += '</div>'; // end pipe-detail
  });

  container.innerHTML = html;

  // Wire up card click → toggle detail
  container.querySelectorAll('.pipe-card').forEach(function(card) {
    card.addEventListener('click', function() {
      var idx = this.getAttribute('data-pipeidx');
      var detail = document.getElementById('pipeDetail_' + idx);
      if (detail) {
        var isOpen = detail.classList.contains('show');
        // Close all
        container.querySelectorAll('.pipe-detail').forEach(function(d) { d.classList.remove('show'); });
        if (!isOpen) detail.classList.add('show');
      }
    });
  });

  // Wire up "Confirm stage" buttons
  container.querySelectorAll('[data-advance]').forEach(function(btn) {
    btn.addEventListener('click', function(e) {
      e.stopPropagation();
      var turnId = this.getAttribute('data-advance');
      var stage = this.getAttribute('data-stage');
      confirmTurnStage(turnId, stage);
    });
  });
}

// Manual stage advancement
async function confirmTurnStage(turnId, stageKey) {
  var stageData = { done: true, date: new Date().toISOString(), manual: true };

  // Update local record first
  var rec = TURN_RECORDS.find(function(r) { return r.id === turnId; });
  if (!rec) {
    rec = { id: turnId, stages: {} };
    TURN_RECORDS.push(rec);
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

  // Classify each inspection
  var classified = INSPECTIONS.map(function(r) {
    var lastDate = r.lastInspection ? new Date(r.lastInspection) : null;
    var daysSince = lastDate ? daysBetween(lastDate, today) : 999;
    var overdue = !lastDate || daysSince > 365;
    var dueSoon = !overdue && daysSince > 270;
    // Check if linked to an active turn
    var linkedTurn = TURN_PIPE_DATA.find(function(tp) {
      return !tp.isCompleted &&
        tp.unit && r.unit && String(tp.unit).toLowerCase() === String(r.unit).toLowerCase() &&
        tp.property && r.propertyName && String(tp.property).toLowerCase() === String(r.propertyName).toLowerCase();
    });
    return {
      r: r,
      daysSince: daysSince,
      overdue: overdue,
      dueSoon: dueSoon,
      current: !overdue && !dueSoon,
      linkedTurn: linkedTurn || null,
      status: overdue ? 'overdue' : dueSoon ? 'due_soon' : 'current'
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

  // Sync group filter dropdown
  var inspGrpSel = $('#inspGroupFilter');
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

  if (filtered.length === 0) {
    body.innerHTML = '<tr><td colspan="8">' + emptyHtml('fa-clipboard-check', INSPECTIONS.length === 0 ? 'No inspection data. Try refreshing.' : 'No inspections match filter') + '</td></tr>';
    return;
  }

  var html = '';
  filtered.forEach(function(c, idx) {
    var r = c.r;
    var statusTag = c.overdue
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

  // Click inspection row to show detail
  body.querySelectorAll('.insp-row').forEach(function(row) {
    row.addEventListener('click', function() {
      var idx = parseInt(this.getAttribute('data-inspidx'), 10);
      var c = filtered[idx];
      if (!c) return;
      var r = c.r;
      var afLink = appfolioUrl('property', r.propertyId);
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
      ], afLink);
    });
  });

  var ib = $('#inspBadge');
  if (ib) ib.textContent = overdueCount;
}

function renderVendors(search) {
  var container = $('#vendorGrid');

  // Sync group filter dropdown
  var vendGrpSel = $('#vendorGroupFilter');
  if (vendGrpSel && vendGrpSel.value !== currentPropertyGroup) vendGrpSel.value = currentPropertyGroup;

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

  var filtered = VENDORS.filter(function(v) {
    if (vendorsInGroup && !vendorsInGroup[(v.name || '').toLowerCase()]) return false;
    if (!search) return true;
    var s = search.toLowerCase();
    return (v.name || '').toLowerCase().indexOf(s) !== -1 || (v.email || '').toLowerCase().indexOf(s) !== -1;
  });

  if (filtered.length === 0) {
    container.innerHTML = emptyHtml('fa-hard-hat', VENDORS.length === 0 ? 'No vendors loaded' : 'No vendors match your search');
    return;
  }

  var html = '';
  var today = new Date();
  filtered.forEach(function(v) {
    var ed = v.insurance ? new Date(v.insurance) : null;
    var exp = ed ? ed < today : false;
    var due = ed ? daysBetween(today, ed) : 999;
    var wrn = !exp && due <= 60;
    var cc = exp ? 'expired' : wrn ? 'warn' : '';
    var afUrl = appfolioUrl('vendor', v.id);
    html += '<div class="vendor-card ' + cc + '" data-vendorid="' + escapeHtml(String(v.id)) + '" style="cursor:pointer">';
    html += '<div class="vendor-name">' + escapeHtml(v.name) + (afUrl ? ' <a href="' + escapeHtml(afUrl) + '" target="_blank" rel="noopener noreferrer" style="font-size:10px;color:var(--accent);text-decoration:none" title="View in AppFolio" onclick="event.stopPropagation()"><i class="fas fa-external-link-alt"></i></a>' : '') + '</div>';
    html += '<div class="vendor-id"><i class="fas fa-fingerprint"></i> ID: ' + escapeHtml(String(v.id)) + '</div>';
    html += '<div class="vendor-row"><span>Compliance</span>' + (v.compliant ? '<span class="tag compliant">Compliant</span>' : '<span class="tag non-compliant">' + escapeHtml(v.compliantStatus) + '</span>') + '</div>';
    if (v.insurance) {
      html += '<div class="vendor-row"><span>Insurance Exp.</span><span style="font-family:var(--font-mono);font-size:12px;color:' + (exp ? 'var(--danger)' : wrn ? 'var(--warning)' : 'var(--text-secondary)') + '">' + escapeHtml(v.insurance) + (exp ? ' (EXPIRED)' : wrn ? ' (' + due + 'd)' : '') + '</span></div>';
    }
    if (v.phone) { html += '<div class="vendor-row"><span>Phone</span><span style="font-family:var(--font-mono);font-size:12px">' + escapeHtml(v.phone) + '</span></div>'; }
    if (v.email) { html += '<div class="vendor-row"><span>Email</span><span style="font-size:12px">' + escapeHtml(v.email) + '</span></div>'; }
    html += '</div>';
  });
  container.innerHTML = html;

  // Click vendor card to show detail
  container.querySelectorAll('.vendor-card').forEach(function(card) {
    card.addEventListener('click', function() {
      var vid = this.getAttribute('data-vendorid');
      var v = VENDORS.find(function(vn) { return String(vn.id) === vid; });
      if (!v) return;
      showItemDetail('Vendor \u2014 ' + v.name, [
        { section: 'Vendor Info', icon: 'fa-hard-hat' },
        { label: 'Name', value: v.name },
        { label: 'ID', value: String(v.id) },
        { label: 'Type', value: v.vendorType || '\u2014' },
        { label: 'Trades', value: v.trades || '\u2014' },
        { section: 'Contact', icon: 'fa-phone' },
        { label: 'Phone', value: v.phone || '\u2014' },
        { label: 'Email', value: v.email || '\u2014' },
        { label: 'Address', value: v.address || '\u2014' },
        { section: 'Compliance', icon: 'fa-shield-alt' },
        { label: 'Status', value: v.compliant ? 'Compliant' : (v.compliantStatus || 'Unknown') },
        { label: 'Liability Ins. Exp.', value: v.insurance || '\u2014' },
        { label: 'Auto Ins. Exp.', value: v.autoInsurance || '\u2014' },
        { label: 'Workers Comp Exp.', value: v.workersComp || '\u2014' },
        { label: 'Do Not Use', value: v.doNotUse ? 'YES' : 'No' }
      ], appfolioUrl('vendor', v.id));
    });
  });
}

function renderReconciliation() {
  var container = $('#reconList');

  // Sync group filter dropdown
  var reconGrpSel = $('#reconGroupFilter');
  if (reconGrpSel && reconGrpSel.value !== currentPropertyGroup) reconGrpSel.value = currentPropertyGroup;

  var unlinked = BILLS.filter(function(b) {
    if (b.workOrderId) return false;
    if (!isInPropertyGroup(b.propertyId, b.propertyName, currentPropertyGroup)) return false;
    return true;
  });

  if (unlinked.length === 0) {
    container.innerHTML = emptyHtml('fa-balance-scale', BILLS.length === 0 ? 'No bills loaded' : 'All bills are linked to work orders!');
    return;
  }

  var html = '';
  unlinked.forEach(function(bill) {
    // Match heuristic: vendor name + property match
    var matches = WORK_ORDERS.filter(function(wo) {
      var vendorMatch = bill.vendorName && wo.vendorName && String(bill.vendorName).toLowerCase() === String(wo.vendorName).toLowerCase();
      var propMatch = bill.propertyName && wo.propertyName && String(bill.propertyName).toLowerCase() === String(wo.propertyName).toLowerCase();
      return (vendorMatch || propMatch) && wo.status !== 'Completed' && wo.status !== 'Canceled';
    });
    var confidence = matches.length > 0 ? 'high' : 'low';
    var matchWO = matches.length > 0 ? matches[0] : null;
    if (matchWO && Math.abs(parseFloat(matchWO.amount || 0) - bill.amount) > 200) { confidence = 'medium'; }

    var billAfUrl = appfolioUrl('bill', bill.id);
    html += '<div class="recon-card"><div class="recon-card-header"><div style="font-family:var(--font-mono);font-weight:600;font-size:14px">' + escapeHtml(String(bill.id || bill.reference)) + (billAfUrl ? ' <a href="' + escapeHtml(billAfUrl) + '" target="_blank" rel="noopener noreferrer" style="font-size:10px;color:var(--accent);text-decoration:none" title="View bill in AppFolio"><i class="fas fa-external-link-alt"></i></a>' : '') + '</div><span class="recon-match ' + confidence + '">' + confidence + ' confidence</span></div>';
    html += '<div class="recon-detail"><div class="recon-side"><div class="recon-side-label"><i class="fas fa-file-invoice-dollar"></i> Bill</div>';
    html += '<div class="recon-field"><strong>Vendor:</strong> ' + escapeHtml(bill.vendorName) + '</div>';
    html += '<div class="recon-field"><strong>Property:</strong> ' + escapeHtml(bill.propertyName) + '</div>';
    html += '<div class="recon-field"><strong>Amount:</strong> ' + currency(bill.amount) + '</div>';
    html += '<div class="recon-field"><strong>Ref:</strong> ' + escapeHtml(bill.reference) + '</div>';
    html += '<div class="recon-field"><strong>Status:</strong> <span class="tag ' + (bill.approvalStatus === 'Approved' ? 'approved' : bill.approvalStatus === 'On Hold' ? 'paid' : 'pending-tag') + '">' + escapeHtml(bill.approvalStatus || 'N/A') + '</span></div></div>';
    html += '<div class="recon-arrow"><i class="fas fa-long-arrow-alt-right"></i></div>';
    html += '<div class="recon-side"><div class="recon-side-label"><i class="fas fa-wrench"></i> Work Order Match</div>';
    if (matchWO) {
      html += '<div class="recon-field"><strong>ID:</strong> <span style="color:var(--accent)">' + escapeHtml(String(matchWO.id)) + '</span></div>';
      html += '<div class="recon-field"><strong>Unit:</strong> ' + escapeHtml(matchWO.unit) + '</div>';
      html += '<div class="recon-field"><strong>Est. Cost:</strong> ' + currency(parseFloat(matchWO.amount || 0)) + '</div>';
      html += '<div class="recon-field"><strong>Status:</strong> <span class="tag ' + String(matchWO.status).toLowerCase().replace(/\s+/g, '-') + '">' + escapeHtml(matchWO.status) + '</span></div>';
    } else {
      html += '<div style="color:var(--text-muted);font-size:12px;padding:10px 0"><i class="fas fa-question-circle"></i> No matching work order found</div>';
    }
    html += '</div></div>';
    html += '<div class="recon-actions">';
    if (matchWO) {
      html += '<button class="action-btn" data-dismiss="' + escapeHtml(String(bill.id)) + '"><i class="fas fa-times"></i> Dismiss</button>';
      html += '<button class="action-btn primary" data-link="' + escapeHtml(String(bill.id)) + '|' + escapeHtml(String(matchWO.id)) + '"><i class="fas fa-link"></i> Link</button>';
    } else {
      html += '<button class="action-btn" data-flag="' + escapeHtml(String(bill.id)) + '"><i class="fas fa-flag"></i> Flag for Review</button>';
    }
    html += '</div></div>';
  });
  container.innerHTML = html;

  container.querySelectorAll('[data-dismiss]').forEach(function(btn) {
    btn.addEventListener('click', function() { showToast('Dismissed suggestion for ' + this.getAttribute('data-dismiss')); });
  });
  container.querySelectorAll('[data-link]').forEach(function(btn) {
    btn.addEventListener('click', function() {
      var parts = this.getAttribute('data-link').split('|');
      showToast('Linked ' + parts[0] + ' \u2192 ' + parts[1]);
    });
  });
  container.querySelectorAll('[data-flag]').forEach(function(btn) {
    btn.addEventListener('click', function() { showToast('Flagged ' + this.getAttribute('data-flag') + ' for manual review'); });
  });
}

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
  container.querySelectorAll('[data-tcopy]').forEach(function(btn) { btn.addEventListener('click', function() { showToast('Template copied to clipboard'); }); });
  container.querySelectorAll('[data-tedit]').forEach(function(btn) { btn.addEventListener('click', function() { showToast('Edit mode \u2014 modify template variables'); }); });
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
  // Properties dropdown for New WO modal
  var propSelect = $('#nwoProperty');
  if (propSelect) {
    propSelect.innerHTML = '<option value="">— Select Property —</option>';
    PROPERTIES.forEach(function(p) {
      propSelect.innerHTML += '<option value="' + escapeHtml(String(p.id)) + '">' + escapeHtml(p.name) + (p.address ? ' \u2014 ' + escapeHtml(p.address) : '') + '</option>';
    });
    // Also add properties extracted from work orders if API didn't return properties
    if (PROPERTIES.length === 0) {
      var seen = {};
      WORK_ORDERS.forEach(function(wo) {
        if (wo.propertyName && !seen[wo.propertyName]) {
          seen[wo.propertyName] = true;
          propSelect.innerHTML += '<option value="' + escapeHtml(wo.propertyName) + '">' + escapeHtml(wo.propertyName) + '</option>';
        }
      });
    }
  }

  // Vendors dropdown for New WO modal
  var vendSelect = $('#nwoVendor');
  if (vendSelect) {
    vendSelect.innerHTML = '<option value="">— Select Vendor —</option>';
    VENDORS.forEach(function(v) {
      vendSelect.innerHTML += '<option value="' + escapeHtml(String(v.id)) + '">' + escapeHtml(v.name) + '</option>';
    });
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
    function() { if (BILLS.length > 0) { renderReconciliation(); } },
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

function wireUpUI() {
  if (_uiWired) return;
  _uiWired = true;

  // Navigation tabs
  $$('.nav-tab').forEach(function(tab) {
    tab.addEventListener('click', function() {
      $$('.nav-tab').forEach(function(t) { t.classList.remove('active'); });
      tab.classList.add('active');
      $$('.section').forEach(function(s) { s.classList.remove('active'); });
      var sec = document.getElementById('sec-' + tab.getAttribute('data-tab'));
      if (sec) sec.classList.add('active');
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
  $('#woSearch').addEventListener('input', function() { renderWorkOrders(); });
  if ($('#woPriorityFilter')) {
    $('#woPriorityFilter').addEventListener('change', function() { currentWOPriority = this.value; renderWorkOrders(); });
  }
  if ($('#woTypeFilter')) {
    $('#woTypeFilter').addEventListener('change', function() { currentWOType = this.value; renderWorkOrders(); });
  }
  if ($('#woPropertyFilter')) {
    $('#woPropertyFilter').addEventListener('change', function() { currentWOProperty = this.value; renderWorkOrders(); });
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
    $('#turnPipeSearch').addEventListener('input', function() {
      renderTurnPipelineUI();
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

  // Manual bills load button
  if ($('#btnLoadBills')) {
    $('#btnLoadBills').addEventListener('click', async function() {
      var btn = this;
      btn.disabled = true;
      btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Loading Bills\u2026';
      try {
        await fetchBills();
        renderReconciliation();
        renderDashboardKPIs();
        await saveAllToCache();
        showToast('Bills loaded \u2014 ' + BILLS.length + ' bills');
      } catch (err) {
        showToast('Failed to load bills: ' + (err.message || err));
      } finally {
        btn.disabled = false;
        btn.innerHTML = '<i class="fas fa-download"></i> Load Bills';
      }
    });
  }
  $('#vendorSearch').addEventListener('input', function() { renderVendors(this.value); });
  $('#btnNewWO').addEventListener('click', function() { openModal('newWOModal'); });
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
  $('#inspSearch').addEventListener('input', function() { renderInspections(this.value); });
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

  // Load Groups button — fetch property groups on-demand via v0 API
  if ($('#btnLoadGroups')) {
    $('#btnLoadGroups').addEventListener('click', async function() {
      var btn = this;
      btn.disabled = true;
      btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Loading\u2026';
      try {
        await fetchPropertyGroups();
        populateGroupFilters();
        renderAll();
        showToast('Loaded ' + PROPERTY_GROUPS.length + ' property groups');
      } catch (err) {
        showToast('Failed to load groups: ' + (err.message || err));
      } finally {
        btn.disabled = false;
        btn.innerHTML = '<i class="fas fa-layer-group"></i> Load Groups';
      }
    });
  }

  // ---- Global property group filter ----
  // All group filter dropdowns share the same currentPropertyGroup variable
  // Change on one tab propagates to all tabs on next render
  var groupFilterIds = [
    '#payrollGroupFilter', '#dashGroupFilter', '#inspGroupFilter',
    '#vendorGroupFilter', '#reconGroupFilter', '#turnGroupFilter'
  ];
  groupFilterIds.forEach(function(sel) {
    var el = document.querySelector(sel);
    if (el) {
      el.addEventListener('change', function() {
        currentPropertyGroup = this.value;
        // Sync all group filter dropdowns
        groupFilterIds.concat(['#woGroupFilter']).forEach(function(s) {
          var otherEl = document.querySelector(s);
          if (otherEl && otherEl.value !== currentPropertyGroup) otherEl.value = currentPropertyGroup;
        });
        // Re-render all sections that use the group filter
        try { renderWorkOrders(); } catch (e) { /* */ }
        try { renderPayroll(); } catch (e) { /* */ }
        try { renderInspections($('#inspSearch') ? $('#inspSearch').value : ''); } catch (e) { /* */ }
        try { renderVendors($('#vendorSearch') ? $('#vendorSearch').value : ''); } catch (e) { /* */ }
        try { renderReconciliation(); } catch (e) { /* */ }
        try { renderTurnPipelineUI(); } catch (e) { /* */ }
        try { renderDashboardKPIs(); } catch (e) { /* */ }
      });
    }
  });

  // WO group filter also syncs globally
  if ($('#woGroupFilter')) {
    // Remove existing listener (was set above) and re-add with global sync
    var woGrp = $('#woGroupFilter');
    var newWoGrp = woGrp.cloneNode(true);
    woGrp.parentNode.replaceChild(newWoGrp, woGrp);
    newWoGrp.addEventListener('change', function() {
      currentPropertyGroup = this.value;
      groupFilterIds.forEach(function(s) {
        var otherEl = document.querySelector(s);
        if (otherEl && otherEl.value !== currentPropertyGroup) otherEl.value = currentPropertyGroup;
      });
      try { renderWorkOrders(); } catch (e) { /* */ }
      try { renderPayroll(); } catch (e) { /* */ }
      try { renderInspections($('#inspSearch') ? $('#inspSearch').value : ''); } catch (e) { /* */ }
      try { renderVendors($('#vendorSearch') ? $('#vendorSearch').value : ''); } catch (e) { /* */ }
      try { renderReconciliation(); } catch (e) { /* */ }
      try { renderTurnPipelineUI(); } catch (e) { /* */ }
      try { renderDashboardKPIs(); } catch (e) { /* */ }
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

  setApiStatus('loading', 'Initializing\u2026');
  updateCacheBadge('loading');

  // Load flags from IndexedDB
  await loadFlags();

  // Show skeleton loading states
  if ($('#kanbanBoard')) $('#kanbanBoard').innerHTML = loadingHtml('Checking cache\u2026');
  if ($('#vendorGrid')) $('#vendorGrid').innerHTML = loadingHtml('Checking cache\u2026');
  if ($('#turnPipeline')) $('#turnPipeline').innerHTML = loadingHtml('Checking cache\u2026');
  if ($('#inspBody')) $('#inspBody').innerHTML = '<tr><td colspan="8">' + loadingHtml('Checking cache\u2026') + '</td></tr>';
  if ($('#reconList')) $('#reconList').innerHTML = emptyHtml('fa-download', 'Bills load on demand \u2014 use the Load Bills button');
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
      var cachedBills = await cacheGet('bills');
      var cachedTurns = await cacheGet('turns');
      var cachedInsp = await cacheGet('inspections');

      WORK_ORDERS = cachedWO.data;
      VENDORS = (cachedVendors && cachedVendors.data) ? cachedVendors.data : [];
      PROPERTIES = (cachedProps && cachedProps.data) ? cachedProps.data : [];
      BILLS = (cachedBills && cachedBills.data) ? cachedBills.data : [];
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
  // STEP 2: Pre-flight = Proxy v6 ?action=ping
  //         Tests proxy connectivity + AppFolio auth in one shot
  // ================================================================
  try {
    setApiStatus('loading', 'Pinging proxy\u2026');
    var pingData = await proxyAction('ping');

    if (!pingData.ok) {
      var dbSt = (pingData.db_api && pingData.db_api.status) || pingData.status || 0;
      var rptSt = (pingData.reports_api && pingData.reports_api.status) || 0;
      var detail = 'DB:' + dbSt + ' Reports:' + rptSt;
      logApiError(dbSt, 'Pre-flight: Proxy ping failed \u2014 ' + detail, 'resolved');
      showCorsError('Pre-flight ping failed (' + detail + '). Verify proxy has correct credentials and both domains are accessible.');
      setApiStatus('error', 'Auth Failed (' + detail + ')');
      if (!cacheLoaded) {
        await loadStaleCache();
        renderAll();
      }
      updateCacheBadge(cacheLoaded ? 'cached' : 'offline', null, true);
      return;
    }
    setApiStatus('loading', 'Proxy OK (' + pingData.latency_ms + 'ms) \u2014 loading data\u2026');
  } catch (preErr) {
    var peMsg = preErr.message || 'Connection failed';
    var isCsp = peMsg.indexOf('Content Security Policy') !== -1 || peMsg.indexOf('CSP') !== -1 || peMsg.indexOf('Refused to connect') !== -1 || preErr.name === 'TypeError';
    if (isCsp) {
      logApiError(0, 'Pre-flight BLOCKED: ' + peMsg + '. Click "Allow additional resources" popup, then re-enter credentials.', 'queued');
      showCorsError(peMsg);
      setApiStatus('error', 'CSP Blocked \u2014 Click Allow');
    } else {
      logApiError(0, 'Pre-flight failed: ' + peMsg, 'queued');
      setApiStatus('error', 'Connection Failed');
    }
    if (!cacheLoaded) {
      await loadStaleCache();
      renderAll();
    }
    updateCacheBadge(cacheLoaded ? 'cached' : 'offline', null, true);
    return;
  }

  // ================================================================
  // STEP 3: Full data fetch via proxy v6 action endpoints
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

// Fetch all data via Proxy v6 action endpoints — ONE request per dataset
// Each ?action= call does server-side pagination and returns complete results
// Bills excluded from auto-load — user loads manually on Reconciliation tab
// Every step has a 60-second timeout — NOTHING hangs forever
async function fetchAllLive() {
  var anySuccess = false;
  updateCacheBadge('loading');
  var steps = ['Work Orders', 'Properties', 'Vendors', 'Turns', 'Move-Outs', 'Turn WOs', 'Inspections', 'Groups', 'Tasks', 'Turn Tracker'];
  showProgress('Syncing AppFolio (' + DATA_WINDOW_DAYS + 'd)', steps);

  try {
    // Step 0: Work Orders (proxy v6 — Reports API)
    updateProgress(0, 'active', 'Fetching work orders\u2026');
    var woOk = await withStepTimeout(fetchWorkOrders, 60000);
    updateProgress(0, woOk ? 'done' : 'error', woOk ? WORK_ORDERS.length + ' open work orders' : 'Work orders failed');
    if (woOk) { renderWorkOrders(); renderDashboardKPIs(); renderActivityFeed(); }

    // Step 1: Properties (proxy v6 — Reports API)
    updateProgress(1, 'active', 'Fetching properties\u2026');
    var propOk = await withStepTimeout(fetchProperties, 60000);
    updateProgress(1, propOk ? 'done' : 'error', propOk ? PROPERTIES.length + ' properties' : 'Properties failed');
    if (propOk) { populateDropdowns(); renderWorkOrders(); }

    // Step 2: Vendors (proxy v6 — Reports API)
    updateProgress(2, 'active', 'Fetching vendors\u2026');
    var vendOk = await withStepTimeout(fetchVendors, 60000);
    updateProgress(2, vendOk ? 'done' : 'error', vendOk ? VENDORS.length + ' vendors' : 'Vendors failed');
    if (vendOk) { renderVendors($('#vendorSearch') ? $('#vendorSearch').value : ''); populateDropdowns(); }

    // Step 3: Turns — In Progress only, 60-day window (proxy v6 — Reports API)
    // Short timeout (20s) — turns are supplementary; pipeline works from WOs + move-outs too
    updateProgress(3, 'active', 'Fetching in-progress turns\u2026');
    var turnOk = await withStepTimeout(function() { return fetchTurns(); }, 20000);
    updateProgress(3, turnOk ? 'done' : 'error', turnOk ? TURNS.length + ' turns' : 'Turns skipped (timeout)');
    if (turnOk) { renderTurnBoard(); renderActivityFeed(); }

    // Step 4: Upcoming Move-Outs — tenant directory, Notice tenants (proxy v6 — Reports API)
    updateProgress(4, 'active', 'Fetching upcoming move-outs\u2026');
    var moOk = await withStepTimeout(fetchUpcomingMoveouts, 45000);
    updateProgress(4, moOk ? 'done' : 'error', moOk ? UPCOMING_MOVEOUTS.length + ' upcoming' : 'Move-outs skipped');
    if (moOk) { renderTurnBoard(); renderDashboardKPIs(); }

    // Step 5: Turn Work Orders — DB API v0, Unit Turn type only (real-time status)
    updateProgress(5, 'active', 'Fetching turn work orders\u2026');
    var twoOk = await withStepTimeout(fetchTurnWorkOrders, 20000);
    updateProgress(5, twoOk ? 'done' : 'error', twoOk ? TURN_WORK_ORDERS.length + ' turn WOs' : 'Turn WOs skipped');
    if (twoOk) { renderTurnBoard(); }

    // Step 6: Inspections (proxy v6 — Reports API)
    updateProgress(6, 'active', 'Fetching inspections\u2026');
    var inspOk = await withStepTimeout(fetchInspections, 60000);
    updateProgress(6, inspOk ? 'done' : 'error', inspOk ? INSPECTIONS.length + ' inspections' : 'Inspections failed');
    if (inspOk) { renderInspections($('#inspSearch') ? $('#inspSearch').value : ''); renderActivityFeed(); }

    // Step 7: Property Groups (proxy v6 — DB API v0)
    updateProgress(7, 'active', 'Fetching property groups\u2026');
    var grpOk = await withStepTimeout(fetchPropertyGroups, 45000);
    updateProgress(7, grpOk ? 'done' : 'error', grpOk ? PROPERTY_GROUPS.length + ' groups' : 'Groups skipped');
    if (grpOk) { populateDropdowns(); renderWorkOrders(); }

    // Step 8: Recent Tasks (proxy v6 — DB API v0)
    updateProgress(8, 'active', 'Fetching recent tasks\u2026');
    var taskOk = await withStepTimeout(fetchRecentTasks, 45000);
    updateProgress(8, taskOk ? 'done' : 'error', taskOk ? RECENT_TASKS.length + ' tasks' : 'Tasks skipped');
    if (taskOk) { renderActivityFeed(); }

    // Step 9: Turn Tracker records (proxy blob — persisted stage overrides)
    updateProgress(9, 'active', 'Loading turn tracker\u2026');
    var trkOk = await withStepTimeout(function() {
      return fetchTurnRecords().then(function() { return true; });
    }, 30000);
    updateProgress(9, trkOk ? 'done' : 'error', trkOk ? TURN_RECORDS.length + ' tracked' : 'Tracker skipped');

    // Final re-render: turns + inspections with all available correlated data
    renderTurnBoard();
    renderInspections($('#inspSearch') ? $('#inspSearch').value : '');

    anySuccess = woOk || propOk || vendOk || turnOk || inspOk;
  } catch (e) {
    // individual errors already logged
  }

  // Final full render to ensure everything is consistent
  renderAll();

  if (anySuccess) {
    var summary = 'WO:' + WORK_ORDERS.length + ' V:' + VENDORS.length + ' P:' + PROPERTIES.length + ' T:' + TURNS.length + ' I:' + INSPECTIONS.length;
    setApiStatus('', 'Connected \u2014 ' + summary);
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
      var cachedBills = await cacheGet('bills');
      var cachedTurns = await cacheGet('turns');
      var cachedInsp = await cacheGet('inspections');
      WORK_ORDERS = cachedWO.data;
      VENDORS = (cachedVendors && cachedVendors.data) ? cachedVendors.data : [];
      PROPERTIES = (cachedProps && cachedProps.data) ? cachedProps.data : [];
      BILLS = (cachedBills && cachedBills.data) ? cachedBills.data : [];
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
async function refreshData() {
  var btn = $('#refreshBtn');
  if (btn.disabled) return;
  btn.disabled = true;
  btn.classList.add('spinning');
  btn.innerHTML = '<i class="fas fa-sync-alt"></i> Syncing\u2026';

  try {
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
