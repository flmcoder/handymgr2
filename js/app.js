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

var BRAND_NAME_DEFAULT = 'Fort Lowell Realty';
var BRAND_LOGO_DEFAULT = 'assets/logo.png';
var BRAND_LOGO_FALLBACK = 'https://pfst.cf2.poecdn.net/base/image/6ac452e679a06edc3e17d0dae13fac303de2fdbb970c22eb302651f44c558416?w=1996&h=938';
var PORTAL_BRAND_NAME_DEFAULT = 'Fort Lowell Realty Tech Dispatch';
var PORTAL_BRAND_LOGO_DEFAULT = 'https://pfst.cf2.poecdn.net/base/image/57c851c04753092259d83d0a1aa34e2fd889c7218b50a338e6100dbf21ae922c?w=733&h=982';
var APP_VERSION = 'v9.6.9';
var SERVER_VERSION = '';
var VERSION_MISMATCH_TIMER = null;

function compareVersions(v1, v2) {
  var p1 = String(v1).toLowerCase().replace(/^v/, '').split(/[\.\-]/);
  var p2 = String(v2).toLowerCase().replace(/^v/, '').split(/[\.\-]/);
  for (var i = 0; i < Math.max(p1.length, p2.length); i++) {
    var n1 = parseInt(p1[i] || 0, 10);
    var n2 = parseInt(p2[i] || 0, 10);
    if (n1 > n2) return 1;
    if (n1 < n2) return -1;
  }
  return 0;
}

function syncDisplayedAppVersion() {
  var bareVersion = String(APP_VERSION || '').replace(/^v/i, '');
  var buildTag = $('#buildBadgeTag');
  var vaultVersion = $('#vaultSessionVersion');
  if (buildTag) buildTag.textContent = bareVersion;
  if (vaultVersion) vaultVersion.textContent = 'Secured Session: ' + APP_VERSION;
}

syncDisplayedAppVersion();
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', syncDisplayedAppVersion);
}

function applyBrandConfig(brand) {
  if (!brand || typeof brand !== 'object') return;
  try {
    if (brand.logo_url !== undefined) {
      localStorage.setItem('hm_brand_logo', String(brand.logo_url || '').trim());
    }
    if (brand.name !== undefined) {
      localStorage.setItem('hm_brand_name', String(brand.name || '').trim());
    }
  } catch (_) {}
  applyBrandLogo();
}

function resolveBrandLogo() {
  var logoFromQuery = '';
  try {
    var params = new URLSearchParams(window.location.search || '');
    logoFromQuery = String(params.get('logo') || '').trim();
  } catch (_) {}

  var logoFromStorage = '';
  try {
    logoFromStorage = String(localStorage.getItem('hm_brand_logo') || '').trim();
  } catch (_) {}

  return logoFromQuery || logoFromStorage || BRAND_LOGO_DEFAULT;
}

function applyBrandLogo() {
  var src = resolveBrandLogo();
  ['vaultLogoImg', 'topbarLogoImg'].forEach(function(id) {
    var img = document.getElementById(id);
    if (!img) return;

    img.onerror = function() {
      if (img.dataset.fallbackApplied === '1') {
        img.style.display = 'none';
        return;
      }
      img.dataset.fallbackApplied = '1';
      img.src = BRAND_LOGO_FALLBACK;
    };

    img.src = src;
  });

  try {
    var storedName = String(localStorage.getItem('hm_brand_name') || '').trim();
    document.title = (storedName || BRAND_NAME_DEFAULT) + ' - AppFolio Utility';
  } catch (_) {}
}

applyBrandLogo();

function formatDate(d) {
  if (!d) return '—';
  if (typeof d === 'string') { d = new Date(d); }
  if (isNaN(d.getTime())) return '—';
  var m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return m[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
}
function formatNoteDateTime(d) {
  if (!d) return '—';
  if (typeof d === 'string') d = new Date(d);
  if (!d || isNaN(d.getTime())) return '—';
  var mm = String(d.getMonth() + 1).padStart(2, '0');
  var dd = String(d.getDate()).padStart(2, '0');
  var yyyy = d.getFullYear();
  var mins = String(d.getMinutes()).padStart(2, '0');
  var h24 = d.getHours();
  var ampm = h24 >= 12 ? 'PM' : 'AM';
  var hour = h24 % 12;
  if (hour === 0) hour = 12;
  return mm + '/' + dd + '/' + yyyy + ' ' + hour + ':' + mins + ' ' + ampm;
}
function isTerminalWOStatus(status) {
  var s = String(status || '').trim().toLowerCase();
  return s === 'completed' || s === 'work completed' || s === 'canceled' || s === 'cancelled';
}
function daysBetween(a, b) {
  var da = typeof a === 'string' ? new Date(a) : a;
  var db = typeof b === 'string' ? new Date(b) : b;
  return Math.round(Math.abs(db - da) / 86400000);
}
function amountToNumber(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  var raw = String(value == null ? '' : value).trim();
  if (!raw) return 0;
  var negative = false;
  if (/^\(.*\)$/.test(raw)) {
    negative = true;
    raw = raw.slice(1, -1);
  }
  var normalized = raw.replace(/[^0-9.\-]/g, '');
  var n = parseFloat(normalized);
  if (!isFinite(n)) return 0;
  return negative ? -Math.abs(n) : n;
}
function currency(n, digits) {
  var d = (typeof digits === 'number' && digits >= 0) ? digits : 0;
  return '$' + amountToNumber(n).toLocaleString('en-US', {
    minimumFractionDigits: d,
    maximumFractionDigits: d,
  });
}
function getToastStackRoot() {
  var root = $('#toastStack');
  if (root) return root;
  root = document.createElement('div');
  root.id = 'toastStack';
  root.style.position = 'fixed';
  root.style.right = '20px';
  root.style.bottom = '20px';
  root.style.zIndex = 'var(--z-layer-toast)';
  root.style.display = 'flex';
  root.style.flexDirection = 'column';
  root.style.gap = '8px';
  root.style.maxWidth = 'min(92vw, 420px)';
  root.style.pointerEvents = 'none';
  document.body.appendChild(root);
  return root;
}

function showToast(msg, durationOrOpts) {
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

  var root = getToastStackRoot();
  var toast = document.createElement('div');
  toast.style.pointerEvents = 'auto';
  toast.style.background = 'var(--bg-card)';
  toast.style.border = '1px solid ' + meta.border;
  toast.style.borderRadius = 'var(--radius)';
  toast.style.padding = '10px 14px';
  toast.style.fontSize = '13px';
  toast.style.color = 'var(--text-primary)';
  toast.style.boxShadow = 'var(--shadow)';
  toast.style.display = 'flex';
  toast.style.alignItems = 'flex-start';
  toast.style.gap = '8px';
  toast.style.animation = 'fadeIn 0.2s ease';
  toast.innerHTML =
    '<span style="margin-top:1px;color:' + meta.border + '"><i class="fas ' + (opts.iconClass || meta.icon) + '"></i></span>' +
    '<span style="line-height:1.35;word-break:break-word">' + escapeHtml(String(msg || '')) + '</span>';

  root.appendChild(toast);
  if (root.children.length > 6) {
    root.removeChild(root.children[0]);
  }

  setTimeout(function() {
    if (toast && toast.parentNode) toast.parentNode.removeChild(toast);
  }, opts.duration || 3500);
}

function registerOfflineServiceWorker() {
  if (!('serviceWorker' in navigator)) return;
  if (!window.isSecureContext && location.hostname !== 'localhost') return;

  navigator.serviceWorker.register('/sw.js')
    .then(function(reg) {
      if (reg && reg.waiting) {
        showToast('Offline mode ready', { kind: 'info', duration: 1800, iconClass: 'fa-wifi' });
      }
    })
    .catch(function(err) {
      console.warn('Service worker registration failed:', err && err.message ? err.message : err);
    });

  navigator.serviceWorker.addEventListener('message', function(evt) {
    var data = evt && evt.data ? evt.data : null;
    if (!data || typeof data !== 'object') return;

    if (data.type === 'HM_REQUEST_QUEUED') {
      var n = Number(data.remaining || 0);
      showToast('Offline: request queued' + (n > 0 ? ' (' + n + ' pending)' : ''), {
        kind: 'warning',
        duration: 2800,
        iconClass: 'fa-cloud-arrow-up'
      });
      return;
    }

    if (data.type === 'HM_QUEUE_REPLAY_RESULT') {
      var sent = Number(data.sent || 0);
      var remaining = Number(data.remaining || 0);
      if (sent > 0) {
        showToast('Synced ' + sent + ' queued request' + (sent === 1 ? '' : 's') + (remaining > 0 ? '; ' + remaining + ' still pending' : ''), {
          kind: 'success',
          duration: 3200,
          iconClass: 'fa-cloud-check'
        });
      }
    }
  });

  window.addEventListener('online', function() {
    if (!navigator.serviceWorker || !navigator.serviceWorker.controller) return;
    navigator.serviceWorker.controller.postMessage({ type: 'HM_REPLAY_QUEUE' });
  });
}

registerOfflineServiceWorker();

function closeModal(id) { document.getElementById(id).classList.remove('show'); }
function openModal(id) { document.getElementById(id).classList.add('show'); }
function escapeHtml(s) { var d = document.createElement('div'); d.textContent = s || ''; return d.innerHTML; }

// ── hmConfirm / hmPrompt ──────────────────────────────────────────────────────
// Promise-based modal replacements for native confirm() and prompt().
// Uses the existing .modal-overlay / .modal CSS classes.
function hmConfirm(msg, opts) {
  var o = opts || {};
  var okLabel = o.okLabel || 'Confirm';
  var okClass = o.danger ? 'background:var(--danger);color:#fff;border:1px solid var(--danger)' : 'background:var(--accent);color:#fff;border:1px solid var(--accent)';
  return new Promise(function(resolve) {
    var overlay = document.createElement('div');
    overlay.className = 'modal-overlay show';
    overlay.style.zIndex = 'var(--z-layer-modal-critical)';
    overlay.innerHTML =
      '<div class="modal" style="max-width:420px">' +
        '<div class="modal-head"><h3>' + escapeHtml(o.title || 'Confirm') + '</h3></div>' +
        '<div class="modal-body"><p style="margin:0;line-height:1.6;white-space:pre-wrap">' + escapeHtml(msg) + '</p></div>' +
        '<div class="modal-footer">' +
          '<button class="hm-modal-cancel" style="background:var(--bg-input);color:var(--text-secondary);border:1px solid var(--border);border-radius:var(--radius);padding:7px 16px;cursor:pointer;font-family:var(--font-mono);font-size:12px">Cancel</button>' +
          '<button class="hm-modal-ok" style="' + okClass + ';border-radius:var(--radius);padding:7px 16px;cursor:pointer;font-family:var(--font-mono);font-size:12px">' + escapeHtml(okLabel) + '</button>' +
        '</div>' +
      '</div>';
    document.body.appendChild(overlay);
    function done(v) { overlay.remove(); resolve(v); }
    overlay.querySelector('.hm-modal-ok').addEventListener('click', function() { done(true); });
    overlay.querySelector('.hm-modal-cancel').addEventListener('click', function() { done(false); });
    overlay.addEventListener('click', function(e) { if (e.target === overlay) done(false); });
    overlay.querySelector('.hm-modal-ok').focus();
  });
}

function hmPrompt(msg, defaultVal, opts) {
  var o = opts || {};
  return new Promise(function(resolve) {
    var overlay = document.createElement('div');
    overlay.className = 'modal-overlay show';
    overlay.style.zIndex = 'var(--z-layer-modal-critical)';
    overlay.innerHTML =
      '<div class="modal" style="max-width:420px">' +
        '<div class="modal-head"><h3>' + escapeHtml(o.title || 'Input') + '</h3></div>' +
        '<div class="modal-body">' +
          '<p style="margin:0 0 12px;line-height:1.6">' + escapeHtml(msg) + '</p>' +
          '<input class="hm-modal-input" type="text" style="width:100%;background:var(--bg-input);border:1px solid var(--border);border-radius:var(--radius);padding:8px 10px;color:var(--text-primary);font-family:var(--font-mono);font-size:12px;box-sizing:border-box">' +
        '</div>' +
        '<div class="modal-footer">' +
          '<button class="hm-modal-cancel" style="background:var(--bg-input);color:var(--text-secondary);border:1px solid var(--border);border-radius:var(--radius);padding:7px 16px;cursor:pointer;font-family:var(--font-mono);font-size:12px">Cancel</button>' +
          '<button class="hm-modal-ok" style="background:var(--accent);color:#fff;border:1px solid var(--accent);border-radius:var(--radius);padding:7px 16px;cursor:pointer;font-family:var(--font-mono);font-size:12px">OK</button>' +
        '</div>' +
      '</div>';
    document.body.appendChild(overlay);
    var inp = overlay.querySelector('.hm-modal-input');
    inp.value = defaultVal || '';
    setTimeout(function() { inp.focus(); inp.select(); }, 30);
    function done(v) { overlay.remove(); resolve(v); }
    overlay.querySelector('.hm-modal-ok').addEventListener('click', function() { done(inp.value); });
    overlay.querySelector('.hm-modal-cancel').addEventListener('click', function() { done(null); });
    inp.addEventListener('keydown', function(e) {
      if (e.key === 'Enter') done(inp.value);
      if (e.key === 'Escape') done(null);
    });
    overlay.addEventListener('click', function(e) { if (e.target === overlay) done(null); });
  });
}
function loadingHtml(msg) { return '<div class="loading-overlay"><i class="fas fa-circle-notch"></i><p>' + escapeHtml(msg) + '</p></div>'; }
function emptyHtml(icon, msg) { return '<div class="empty-state"><i class="fas ' + icon + '"></i><p>' + escapeHtml(msg) + '</p></div>'; }

// ---- AppFolio deep link builder ----
// Builds URLs to view resources directly in AppFolio
// Work orders: /maintenance/service_requests/{base_number}/
//   "12345-1" → "12345" (strip hyphen suffix)
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

function normalizePropertyLookupKey(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function groupsByPropertyName(propertyName) {
  var lower = String(propertyName || '').trim().toLowerCase();
  if (!lower) return [];
  var byExact = _nameToGroups[lower] || [];
  if (byExact.length) return byExact;
  var normalized = normalizePropertyLookupKey(lower);
  if (!normalized) return [];
  return _nameToGroups[normalized] || [];
}

function isInPropertyGroup(propertyId, propertyName, groupName) {
  var normalizedGroup = String(groupName || '').trim();
  if (!normalizedGroup) return true; // no filter = show all
  var normalizedGroupLower = normalizedGroup.toLowerCase();
  if (normalizedGroupLower === 'all properties') return true;
  if (normalizedGroup.charAt(0) === '*' && normalizedGroupLower.indexOf('all properties') !== -1) return true;

  groupName = normalizedGroup;

  // 1. Fast lookup by property name (covers both Reports API and DB API names)
  if (propertyName) {
    var groups = groupsByPropertyName(propertyName);
    if (groups && groups.indexOf(groupName) !== -1) return true;
  }

  // 2. Fast lookup by Reports API property_id
  if (propertyId) {
    var idGroups = _idToGroups[String(propertyId)];
    if (idGroups && idGroups.indexOf(groupName) !== -1) return true;
  }

  // 2.5. Lookup by DB API UUID (handles unit-turn and DB-sourced items with UUID propertyId)
  if (propertyId) {
    var uuidGroups = _uuidToGroups[String(propertyId).trim()];
    if (uuidGroups && uuidGroups.indexOf(groupName) !== -1) return true;
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

function _addNameToGroupMap(name, groupName) {
  var lower = String(name || '').trim().toLowerCase();
  if (!lower) return;
  _addToGroupMap(_nameToGroups, lower, groupName);
  var normalized = normalizePropertyLookupKey(lower);
  if (normalized && normalized !== lower) {
    _addToGroupMap(_nameToGroups, normalized, groupName);
  }
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
  if (forcedPropertyGroupUuid) {
    enforceScopedPropertyGroup();
  } else {
    el.disabled = false;
  }
  // Update the active indicator badge
  updateGlobalGroupIndicator();
  // Refresh PM user group dropdown if the dbadmin panel exposed its populate hook
  if (typeof window._repopulatePMGroupDropdown === 'function') window._repopulatePMGroupDropdown();
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

function setDashboardKpiSkeleton(active) {
  $$('#sec-dashboard .kpi-card').forEach(function(card) {
    card.classList.toggle('kpi-skeleton', !!active);
  });
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

function cacheDelete(key) {
  return openCacheDB().then(function(db) {
    return new Promise(function(resolve) {
      var tx = db.transaction(CACHE_STORE, 'readwrite');
      tx.objectStore(CACHE_STORE).delete(key);
      tx.oncomplete = function() { resolve(); };
      tx.onerror = function() { resolve(); };
    });
  }).catch(function() { /* ignore */ });
}

async function clearSessionScopedApiCache() {
  await Promise.all([
    cacheDelete('work_orders'),
    cacheDelete('vendors'),
    cacheDelete('properties'),
    cacheDelete('turns'),
    cacheDelete('inspections'),
    cacheDelete('webhooks')
  ]);
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

function _bufToBase64(buffer) {
  var bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  var bin = '';
  for (var i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return btoa(bin);
}

function _base64ToBuf(base64) {
  var bin = atob(String(base64 || ''));
  var out = new Uint8Array(bin.length);
  for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

async function _deriveCacheCryptoKey(passphrase, saltBytes) {
  var enc = new TextEncoder();
  var keyMaterial = await crypto.subtle.importKey(
    'raw',
    enc.encode(String(passphrase || '')),
    { name: 'PBKDF2' },
    false,
    ['deriveKey']
  );
  return crypto.subtle.deriveKey(
    {
      name: 'PBKDF2',
      salt: saltBytes,
      iterations: 120000,
      hash: 'SHA-256'
    },
    keyMaterial,
    { name: 'AES-GCM', length: 256 },
    false,
    ['encrypt', 'decrypt']
  );
}

async function _encryptCachePayload(plaintext, passphrase) {
  var salt = crypto.getRandomValues(new Uint8Array(16));
  var iv = crypto.getRandomValues(new Uint8Array(12));
  var key = await _deriveCacheCryptoKey(passphrase, salt);
  var enc = new TextEncoder();
  var cipher = await crypto.subtle.encrypt(
    { name: 'AES-GCM', iv: iv },
    key,
    enc.encode(String(plaintext || ''))
  );
  return {
    _meta: {
      format: 'hm-cache-encrypted-v1',
      exported: new Date().toISOString(),
      cipher: 'AES-GCM',
      kdf: 'PBKDF2-SHA256',
      iterations: 120000
    },
    salt: _bufToBase64(salt),
    iv: _bufToBase64(iv),
    ciphertext: _bufToBase64(cipher)
  };
}

async function _decryptCachePayload(payload, passphrase) {
  if (!payload || !payload.salt || !payload.iv || !payload.ciphertext) {
    throw new Error('Invalid encrypted cache file');
  }
  var salt = _base64ToBuf(payload.salt);
  var iv = _base64ToBuf(payload.iv);
  var cipher = _base64ToBuf(payload.ciphertext);
  var key = await _deriveCacheCryptoKey(passphrase, salt);
  var plain = await crypto.subtle.decrypt(
    { name: 'AES-GCM', iv: iv },
    key,
    cipher
  );
  return new TextDecoder().decode(plain);
}

function getCacheSecurityOptions() {
  var useEncEl = document.getElementById('settingsUseEncryption');
  var passEl = document.getElementById('settingsEncryptPass');
  return {
    encrypt: !!(useEncEl && useEncEl.checked),
    passphrase: passEl ? String(passEl.value || '').trim() : ''
  };
}

// Export all data as a downloadable cache file — reads from MEMORY (not IndexedDB)
async function exportCacheToJSON(options) {
  try {
    var opts = options || {};
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
    var useEncryption = !!opts.encrypt;
    var payloadText = json;
    var fileExt = 'json';
    if (useEncryption) {
      if (!opts.passphrase) {
        showToast('Enter an encryption passphrase in Settings before export', { kind: 'warning' });
        return;
      }
      var encryptedEnvelope = await _encryptCachePayload(json, opts.passphrase);
      payloadText = JSON.stringify(encryptedEnvelope);
      fileExt = 'hmc';
    }
    var blob = new Blob([payloadText], { type: 'application/json' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = 'maint-cockpit-' + new Date().toISOString().split('T')[0] + '.' + fileExt;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function() { URL.revokeObjectURL(url); }, 5000);
    var sizeKB = Math.round(payloadText.length / 1024);
    showToast(
      'Exported ' + total + ' records (' + sizeKB + ' KB)' + (useEncryption ? ' [encrypted]' : '') +
      '  WO:' + counts.work_orders + ' V:' + counts.vendors + ' P:' + counts.properties + ' T:' + counts.turns + ' I:' + counts.inspections,
      { kind: 'success' }
    );
  } catch (e) {
    showToast('Export failed: ' + (e.message || e));
  }
}

// Import data from a JSON file — writes to MEMORY + IndexedDB
async function importCacheFromJSON(file, options) {
  if (!file) return;
  try {
    var opts = options || {};
    var text = await new Promise(function(resolve, reject) {
      var reader = new FileReader();
      reader.onload = function(ev) { resolve(ev.target.result); };
      reader.onerror = function() { reject(new Error('File read error')); };
      reader.readAsText(file);
    });
    var parsed = JSON.parse(text);
    var data = parsed;
    var isEncrypted = !!(parsed && parsed._meta && parsed._meta.format === 'hm-cache-encrypted-v1');
    if (isEncrypted) {
      if (!opts.passphrase) {
        showToast('Encrypted file detected. Enter passphrase in Settings and retry import.', { kind: 'warning' });
        return;
      }
      try {
        var decrypted = await _decryptCachePayload(parsed, opts.passphrase);
        data = JSON.parse(decrypted);
      } catch (decErr) {
        showToast('Decrypt failed: invalid passphrase or file', { kind: 'danger' });
        return;
      }
    }
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
    showToast(
      'Imported ' + total + ' records' + (isEncrypted ? ' [decrypted]' : '') +
      '  WO:' + WORK_ORDERS.length + ' V:' + VENDORS.length + ' P:' + PROPERTIES.length + ' T:' + TURNS.length + ' I:' + INSPECTIONS.length,
      { kind: 'success' }
    );
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
var VENDOR_TRADE_CATEGORIES = ['General', 'HVAC', 'Plumbing', 'Electrical', 'Appliance', 'Flooring', 'Painting', 'Landscaping', 'Cleaning', 'Roofing', 'Pest Control', 'Pool/Spa', 'Locksmith', 'Security', 'Garage', 'Fences/Gates', 'Drywall', 'Gutter', 'Turnover', 'Other'];

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
  if (raw.indexOf('garage') !== -1 || raw.indexOf('door') !== -1 && raw.indexOf('garage') !== -1) return 'Garage';
  if (raw.indexOf('fence') !== -1 || raw.indexOf('gate') !== -1) return 'Fences/Gates';
  if (raw.indexOf('drywall') !== -1 || raw.indexOf('plaster') !== -1) return 'Drywall';
  if (raw.indexOf('gutter') !== -1) return 'Gutter';
  if (raw.indexOf('turnover') !== -1 || raw.indexOf('turn over') !== -1 || raw.indexOf('make ready') !== -1) return 'Turnover';
  if (raw.indexOf('general') !== -1 || raw.indexOf('handyman') !== -1 || raw.indexOf('handyperson') !== -1 || raw.indexOf('maintenance') !== -1) return 'General';
  return 'Other';
}

function inferVendorTradeCategory(vendor) {
  var v = vendor || {};
  var trades = String(v.trades || '').trim();
  if (trades) {
    var tokens = trades.split(/[\/,;|]/).map(function(t) { return String(t || '').trim(); }).filter(Boolean);
    for (var i = 0; i < tokens.length; i++) {
      var normalized = normalizeVendorTradeCategory(tokens[i]);
      if (normalized && normalized !== 'Other') return normalized;
    }
    var primary = trades.split(',')[0].split('/')[0];
    return normalizeVendorTradeCategory(primary);
  }

  var corpus = [
    v.name,
    v.email,
    v.tags,
    v.category,
    v.address,
    v.notes,
    v.description,
  ].map(function(s) { return String(s || '').trim(); }).filter(Boolean).join(' ');

  if (corpus) {
    var guess = normalizeVendorTradeCategory(corpus);
    if (guess && guess !== 'Other') return guess;
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
   AUTH GLOBALS — role verified server-side via Val.town proxy.
   GUI_ADMIN / GUI_GM / GUI_VENDORS env vars are checked on the proxy.
   ================================================================= */
var API_CREDS = null;
var API_VHOST = null;
var API_PROXY = '';
var _accessRole = 'full'; // 'full' | 'manager' | 'vendors' | 'pm_readonly'
var _pmScopeGroupUuid = '';
var _pmScopeEmail = '';
var DEFAULT_BILLS_LOOKBACK_DAYS = 90;
var BILLS_DEFAULT_LOOKBACK_DAYS = 7; // Used for the UI date filter default
var DEFAULT_COMPLETED_WO_LOOKBACK_DAYS = 30;

function normalizeAccessRole(role) {
  var value = String(role || '').trim().toLowerCase();
  if (value === 'vendors' || value === 'manager' || value === 'full' || value === 'pm_readonly') return value;
  return 'full';
}

function isReadOnlyAccessMode() {
  return _accessRole === 'pm_readonly';
}

function persistAccessRole(role) {
  try { localStorage.setItem('hm_access_role', normalizeAccessRole(role)); } catch (e) { /* */ }
}

function getStoredAccessRole() {
  try { return normalizeAccessRole(localStorage.getItem('hm_access_role') || 'full'); } catch (e) { return 'full'; }
}

function isTabAllowedForRole(tabName) {
  if (_accessRole === 'vendors') return tabName === 'vendors';
  if (_accessRole === 'pm_readonly') {
    var pmAllowedTabs = ['dashboard', 'workorders', 'billing', 'properties', 'turnboard', 'vendors', 'inspections', 'errors'];
    return pmAllowedTabs.indexOf(tabName) !== -1;
  }
  if (_accessRole === 'manager') {
    // Manager role allowed tabs: workorders, turnboard, vendors, inspections, errors
    var allowedTabs = ['dashboard', 'workorders', 'routing', 'billing', 'properties', 'turnboard', 'vendors', 'inspections', 'errors'];
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
  setMobileNavLabelFromActiveTab();
  closeMobileNav();
}

function isMobileViewport() {
  return window.matchMedia('(max-width: 768px)').matches;
}

function closeMobileNav() {
  document.body.classList.remove('mobile-nav-open');
  var toggleBtn = document.getElementById('mobileNavToggle');
  if (toggleBtn) toggleBtn.setAttribute('aria-expanded', 'false');
}

function toggleMobileNav() {
  if (!isMobileViewport()) return;
  var nextState = !document.body.classList.contains('mobile-nav-open');
  document.body.classList.toggle('mobile-nav-open', nextState);
  var toggleBtn = document.getElementById('mobileNavToggle');
  if (toggleBtn) toggleBtn.setAttribute('aria-expanded', nextState ? 'true' : 'false');
}

function setMobileNavLabelFromActiveTab() {
  var toggleBtn = document.getElementById('mobileNavToggle');
  if (!toggleBtn) return;
  var activeTab = document.querySelector('.nav-tab.active');
  var label = activeTab ? activeTab.textContent.replace(/\s+/g, ' ').trim() : 'Menu';
  toggleBtn.title = 'Menu: ' + label;
  toggleBtn.setAttribute('aria-label', 'Open navigation menu. Current tab: ' + label);
}

function syncMobileNavForViewport() {
  if (!isMobileViewport()) {
    closeMobileNav();
  }
  setMobileNavLabelFromActiveTab();
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
  _pmScopeGroupUuid = '';
  _pmScopeEmail = '';
  closeMobileNav();
}

// ── Auto-sync: selective background refresh every 30 min ───────────────────
var AUTO_SYNC_INTERVAL_MS = 30 * 60 * 1000; // 30 minutes
var _sessionExpiryHandled = false;
var _proxySessionWarmupUntil = 0;
var _proxySessionProbeInFlight = false;
var _proxySessionWarmupFailures = 0;
var _proxySessionConsecutive401 = 0;
var _proxySessionLastHealthyAt = 0;
var _proxySessionStartupGraceUntil = 0;

function markProxySessionHealthy() {
  _proxySessionLastHealthyAt = Date.now();
  _proxySessionConsecutive401 = 0;
  _proxySessionWarmupFailures = 0;
}

function beginProxySessionStartupGrace(ms) {
  var graceMs = Math.max(30000, Number(ms || 0) || 60000);
  _proxySessionStartupGraceUntil = Date.now() + graceMs;
  markProxySessionWarmup(Math.max(20000, graceMs));
}

function markProxySessionWarmup(ms) {
  var windowMs = Number(ms || 0) || 10000;
  _proxySessionWarmupUntil = Date.now() + Math.max(2000, windowMs);
}

function isProxySessionWarmupActive() {
  return Date.now() <= _proxySessionWarmupUntil;
}

function shouldProbeProxySessionBeforeLockout() {
  var now = Date.now();
  var inWarmup = now <= _proxySessionWarmupUntil;
  var inStartupGrace = now <= _proxySessionStartupGraceUntil;
  var recentlyHealthy = _proxySessionLastHealthyAt && ((now - _proxySessionLastHealthyAt) <= 120000);
  if (!inWarmup && !inStartupGrace && !recentlyHealthy) return false;
  if (_proxySessionProbeInFlight) return false;
  if (!API_PROXY) return false;
  return !!getProxyAccessToken();
}

async function probeProxySessionStillValid() {
  if (!API_PROXY) return false;
  var token = getProxyAccessToken();
  if (!token) return false;
  var sep = API_PROXY.indexOf('?') !== -1 ? '&' : '?';
  var url = API_PROXY + sep + 'action=session_info';
  var res = await fetchWithTimeout(url, {
    headers: {
      'Accept': 'application/json',
      'Authorization': 'Bearer ' + token
    }
  }, 12000);
  if (!res.ok) return false;
  var data = {};
  try { data = await res.json(); } catch (e) { return false; }
  return !!(data && data.ok && data.session);
}

function forceProxySessionExpiryLockout(contextLabel) {
  if (_sessionExpiryHandled) return;
  _sessionExpiryHandled = true;

  stopAutoSync();
  if (_webhookPollTimer) { clearInterval(_webhookPollTimer); _webhookPollTimer = null; }

  clearStoredProxySessionTokens();
  if (API_CREDS && API_CREDS.p) {
    API_CREDS.p = '0'.repeat(String(API_CREDS.p).length);
  }
  API_CREDS = null;
  appInitialized = false;
  _proxySessionStartupGraceUntil = 0;

  var vault = $('#vaultScreen');
  if (vault) vault.style.display = 'flex';
  var shell = $('#appShell');
  if (shell) shell.classList.remove('unlocked');
  setVaultPanel('main');
  setPmOtpStep('request');
  setVaultFeedback('Proxy session expired. Sign in again to continue.', '');
  if ($('#vaultPassphrase')) $('#vaultPassphrase').focus();

  var detail = contextLabel ? (' (' + contextLabel + ')') : '';
  showToast('Session expired. Please sign in again' + detail + '.', {
    kind: 'warning',
    iconClass: 'fa-triangle-exclamation',
    duration: 5000,
  });
}

function startAutoSync() {
  stopAutoSync(); // clear any existing
  _autoSyncTimer = setInterval(async function() {
    if (!API_CREDS || !API_VHOST) return; // not unlocked
    try {
      await fetchWorkOrders();
      await fetchTurns();
      await fetchUpcomingMoveouts();
      await fetchTurnWorkOrders();
      await fetchUnitTurnsDB();
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

function clearStoredProxySessionTokens() {
  try { localStorage.removeItem('hm_auth_token'); } catch (e0) { /* */ }
  try { localStorage.removeItem('hm_device_token'); } catch (e1) { /* */ }
  try { localStorage.removeItem('hm_proxy_token'); } catch (e2) { /* */ }
}

function handleProxySessionExpired(contextLabel) {
  if (_sessionExpiryHandled) return;
  _proxySessionConsecutive401++;
  var now = Date.now();

  if (shouldProbeProxySessionBeforeLockout()) {
    _proxySessionProbeInFlight = true;
    (async function() {
      try {
        var stillValid = await probeProxySessionStillValid();
        if (stillValid) {
          _proxySessionProbeInFlight = false;
          markProxySessionHealthy();
          markProxySessionWarmup(4000);
          return;
        }
      } catch (e) { /* fall through to lockout */ }
      _proxySessionProbeInFlight = false;
      _proxySessionWarmupFailures++;
      var inStartupGrace = now <= _proxySessionStartupGraceUntil;
      var recentlyHealthy = _proxySessionLastHealthyAt && ((now - _proxySessionLastHealthyAt) <= 120000);
      // Never force relock during the explicit post-login grace window.
      // Initial sync can trigger a burst of transient 401s before all
      // backend session state settles.
      if (inStartupGrace) {
        markProxySessionWarmup(7000);
        return;
      }
      // During immediate post-login startup, tolerate a few transient 401s
      // while session rows and token propagation settle.
      if ((isProxySessionWarmupActive() && _proxySessionWarmupFailures < 6) ||
          (recentlyHealthy && _proxySessionConsecutive401 < 4)) {
        markProxySessionWarmup(7000);
        return;
      }
      forceProxySessionExpiryLockout(contextLabel);
    })();
    return;
  }

  forceProxySessionExpiryLockout(contextLabel);
}

function lockVault() {
  wipeCredentials();
  _proxySessionWarmupUntil = 0;
  _proxySessionProbeInFlight = false;
  _proxySessionWarmupFailures = 0;
  _proxySessionConsecutive401 = 0;
  _proxySessionLastHealthyAt = 0;
  _proxySessionStartupGraceUntil = 0;
  appInitialized = false;
  WORK_ORDERS = []; VENDORS = []; PROPERTIES = []; PROPERTY_GROUPS = []; TURNS = []; INSPECTIONS = []; RECENT_TASKS = []; WEBHOOK_EVENTS = []; TURN_RECORDS = []; TURN_PIPE_DATA = []; UNIT_TURNS_DB = []; API_ERRORS = [];
  CLOSED_TURNS = new Set();
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
    // pm_readonly users are auto-scoped to their property group UUID — no manual filter needed
    if (gfBar2) gfBar2.style.display = _accessRole === 'pm_readonly' ? 'none' : '';
  }
  persistAccessRole(_accessRole);
}

function getAuthHeader() { return API_CREDS ? API_CREDS.a : null; }
function getDevId() { return API_CREDS ? API_CREDS.d : null; }
function getDirectBaseUrl() { return API_VHOST ? 'https://' + API_VHOST + '.appfolio.com' : null; }
function getProxyAccessToken() {
  if (API_CREDS && API_CREDS.p) return API_CREDS.p;
  try {
    var authToken = localStorage.getItem('hm_auth_token') || '';
    if (authToken) return authToken;
  } catch (e) { /* */ }
  try {
    var deviceToken = localStorage.getItem('hm_device_token') || '';
    if (deviceToken) return deviceToken;
  } catch (e) { /* */ }
  try { return localStorage.getItem('hm_proxy_token') || ''; } catch (e) { return ''; }
}

function isVaultUnlocked() {
  var vault = $('#vaultScreen');
  if (!vault) return true;
  return vault.style.display === 'none';
}

function isProxySessionReady() {
  if (!isVaultUnlocked()) return false;
  if (!API_PROXY) return false;
  if (API_CREDS) return true;
  return !!getProxyAccessToken();
}

function hasStoredDeviceSession() {
  try {
    return !!(localStorage.getItem('hm_auth_token') || localStorage.getItem('hm_device_token') || localStorage.getItem('hm_proxy_token') || '');
  } catch (e) {
    return false;
  }
}

function setVaultFeedback(message, state) {
  var el = $('#vaultError');
  if (!el) return;
  el.textContent = String(message || '');
  el.classList.remove('success', 'info');
  if (!message) {
    el.classList.remove('show');
    return;
  }
  if (state === 'success' || state === 'info') {
    el.classList.add(state);
  }
  el.classList.add('show');
}

function formatOtpIdentifierSummary(identifier) {
  var email = normalizeOtpEmail(identifier);
  if (email) return email;
  var phone = normalizeOtpPhone(identifier);
  if (!phone) return String(identifier || '').trim();
  var digits = phone.replace(/\D+/g, '');
  if (digits.length >= 10) {
    return '+1 (' + digits.slice(-10, -7) + ') ' + digits.slice(-7, -4) + '-' + digits.slice(-4);
  }
  return phone;
}

function setPmOtpStep(step, identifier) {
  var route = $('#vaultOtpRoute');
  var identifierWrap = $('#vaultOtpIdentifierWrap');
  var requestActions = $('#vaultOtpRequestActions');
  var summary = $('#vaultOtpSentSummary');
  var summaryValue = $('#vaultOtpSentValue');
  var editRow = $('#vaultOtpEditRow');
  var verifyRow = $('#vaultOtpVerifyRow');
  var identifierInput = $('#vaultOtpEmail');
  var codeInput = $('#vaultOtpCode');
  var verifyBtn = $('#btnVerifyOtp');
  var sent = step === 'verify';

  if (route) route.classList.toggle('compact', sent);
  if (identifierWrap) identifierWrap.classList.toggle('hidden', sent);
  if (requestActions) requestActions.classList.toggle('hidden', sent);
  if (summary) summary.classList.toggle('hidden', !sent);
  if (editRow) editRow.classList.toggle('hidden', !sent);
  if (verifyRow) verifyRow.classList.toggle('hidden', !sent);
  if (identifierInput) identifierInput.disabled = sent;
  if (summaryValue && sent) summaryValue.textContent = formatOtpIdentifierSummary(identifier || (identifierInput ? identifierInput.value : ''));
  if (verifyBtn) verifyBtn.textContent = 'Verify OTP';

  if (!sent && codeInput) codeInput.value = '';
  if (sent && codeInput) {
    setTimeout(function() { codeInput.focus(); codeInput.select(); }, 20);
  }
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
  var m = email.match(/^[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}$/i);
  return m ? email : '';
}

function normalizeOtpPhone(raw) {
  var val = String(raw || '').trim();
  if (!val) return '';
  var digits = val.replace(/\D+/g, '');
  if (digits.length === 11 && digits.charAt(0) === '1') return '+' + digits;
  if (digits.length === 10) return '+1' + digits;
  if (val.charAt(0) === '+' && digits.length >= 10 && digits.length <= 15) return '+' + digits;
  return '';
}

function normalizeOtpIdentifier(raw) {
  return normalizeOtpEmail(raw) || normalizeOtpPhone(raw);
}

async function requestDeviceOtp(identifier, userName) {
  if (!API_PROXY) throw new Error('No proxy configured');
  var sep = API_PROXY.indexOf('?') !== -1 ? '&' : '?';
  var url = API_PROXY + sep + 'action=device_otp_request';
  var email = normalizeOtpEmail(identifier);
  var phone = normalizeOtpPhone(identifier);
  var payload = {
    identifier: String(identifier || '').trim(),
    email: email || undefined,
    phone: phone || undefined,
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

async function verifyDeviceOtp(identifier, code, userName) {
  if (!API_PROXY) throw new Error('No proxy configured');
  var sep = API_PROXY.indexOf('?') !== -1 ? '&' : '?';
  var url = API_PROXY + sep + 'action=device_otp_verify';
  var email = normalizeOtpEmail(identifier);
  var phone = normalizeOtpPhone(identifier);
  var payload = {
    identifier: String(identifier || '').trim(),
    email: email || undefined,
    phone: phone || undefined,
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
  return data;
}

async function stabilizeProxySessionAfterLogin(timeoutMs) {
  if (!API_PROXY) return false;
  var token = getProxyAccessToken();
  if (!token) return false;

  var deadline = Date.now() + Math.max(10000, Number(timeoutMs || 30000));
  var attempt = 0;
  while (Date.now() < deadline) {
    attempt++;
    try {
      var sess = await proxyAction('session_info');
      if (sess && sess.ok && sess.authenticated && sess.session) return true;
    } catch (e) {
      // Keep retrying within the bounded window.
    }
    var waitMs = Math.min(250 + (attempt * 150), 1000);
    await sleep(waitMs);
  }
  return false;
}

// ---- Resilient timeout + retry helper ----
// Backward compatible with existing callers that pass a numeric timeout as arg #3.
async function fetchWithTimeout(url, opts, timeoutMsOrRetries, baseBackoffMs) {
  var options = Object.assign({}, opts || {});
  var timeoutMs = Number(options.timeout || 0) || 15000;
  var retries = 3;
  var backoffMs = Number(baseBackoffMs || 0) || 2000;

  if (typeof timeoutMsOrRetries === 'number') {
    if (timeoutMsOrRetries > 20) timeoutMs = timeoutMsOrRetries;
    else retries = Math.max(0, timeoutMsOrRetries);
  }

  delete options.timeout;

  var attempt = 0;
  while (true) {
    var controller = new AbortController();
    var timer = setTimeout(function() { controller.abort(); }, timeoutMs);
    try {
      var response = await fetch(url, Object.assign({}, options, { signal: controller.signal }));
      clearTimeout(timer);

      // Treat transient proxy/server failures as retryable, but return final response
      // when attempts are exhausted so existing callers can handle status codes.
      if (!response.ok && response.status >= 500 && attempt < retries) {
        var serverJitter = Math.floor(Math.random() * 500);
        var serverDelay = backoffMs + serverJitter;
        console.warn('[Network Fluctuated] Proxy/server HTTP ' + response.status + '. Retrying in ' + Math.round(serverDelay / 1000) + 's... (' + (retries - attempt) + ' attempts left)');
        await new Promise(function(resolve) { setTimeout(resolve, serverDelay); });
        backoffMs = backoffMs * 2;
        attempt++;
        continue;
      }

      return response;
    } catch (err) {
      clearTimeout(timer);
      var msg = String((err && err.message) || '').toLowerCase();
      var isNetworkError =
        (err && err.name === 'AbortError') ||
        msg.indexOf('fetch') !== -1 ||
        msg.indexOf('network') !== -1 ||
        msg.indexOf('name_not_resolved') !== -1 ||
        msg.indexOf('quic') !== -1;

      if (attempt < retries && isNetworkError) {
        var jitter = Math.floor(Math.random() * 500);
        var delay = backoffMs + jitter;
        console.warn('[Network Fluctuated] ' + ((err && err.message) || 'network error') + '. Retrying in ' + Math.round(delay / 1000) + 's... (' + (retries - attempt) + ' attempts left)');
        await new Promise(function(resolve) { setTimeout(resolve, delay); });
        backoffMs = backoffMs * 2;
        attempt++;
        continue;
      }
      throw err;
    }
  }
}

// ---- Proxy action endpoint caller ----
// Makes ONE request to proxy like ?action=work_orders&days=180
// Proxy does all pagination server-side and returns complete dataset
// Includes 45-second timeout — never hangs forever
async function proxyAction(action, params, options) {
  var opts = options || {};
  if (!API_PROXY) throw new Error('No proxy configured');
  var sep = API_PROXY.indexOf('?') !== -1 ? '&' : '?';
  var url = API_PROXY + sep + 'action=' + encodeURIComponent(action);
  if (params) {
    Object.keys(params).forEach(function(k) {
      url += '&' + encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
    });
  }
  var maxRetries = 2; // retry once for transient 502/503/network errors
  var timeoutByAction = {
    turn_work_orders: 120000,
    bills_history: 120000,
    properties: 90000,
    bills: 90000,
    turns: 90000,
    unit_turns: 90000,
    work_orders_completed_history: 90000
  };
  var baseTimeoutMs = timeoutByAction[action] || 60000;
  var token = getProxyAccessToken();
  var reqHeaders = { 'Accept': 'application/json' };
  if (token) reqHeaders['Authorization'] = 'Bearer ' + token;
  for (var attempt = 0; attempt <= maxRetries; attempt++) {
    var attemptTimeoutMs = Math.min(baseTimeoutMs + (attempt * 15000), 150000);
    var res;
    try {
      res = await fetchWithTimeout(url, { headers: reqHeaders }, attemptTimeoutMs);
    } catch (abortErr) {
      if (abortErr.name === 'AbortError') {
        if (attempt < maxRetries) {
          var timeoutWait = Math.pow(2, attempt + 1) * 1000;
          logApiError(0, 'Proxy action=' + action + ' timed out after ' + Math.round(attemptTimeoutMs / 1000) + 's (attempt ' + (attempt + 1) + '/' + (maxRetries + 1) + ') — retrying in ' + (timeoutWait / 1000) + 's', 'retry');
          await sleep(timeoutWait);
          continue;
        }
        var tmsg = 'Proxy action=' + action + ' timed out after ' + Math.round(attemptTimeoutMs / 1000) + 's';
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
      if (res.status === 401) {
        var errBodyLower = String(errBody || '').toLowerCase();
        var isProxySession401 = errBodyLower.indexOf('frontend bearer token') !== -1 ||
          errBodyLower.indexOf('proxy session') !== -1 ||
          errBodyLower.indexOf('read-only session') !== -1;
        if (isProxySession401) {
          if (!opts.suppressSessionExpiry) {
            handleProxySessionExpired('action=' + action);
          }
          var sessionMsg = 'Proxy action=' + action + ' failed: HTTP 401 — Proxy session missing or expired';
          logApiError(401, sessionMsg, 'resolved');
          throw new Error(sessionMsg);
        }
      }
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
        // If the proxy itself hit a 429 upstream, pause the shared queue
        if (dataStatus === 429) { rateLimiter.backpressure(retryWait); }
        logApiError(dataStatus || 502, 'Proxy action=' + action + ': ' + (data.error || 'Unknown') + ' — retrying in ' + (retryWait / 1000) + 's', 'retry');
        await sleep(retryWait);
        continue;
      }
      var msg = 'Proxy action=' + action + ': ' + (data.error || 'Unknown error');
      logApiError(dataStatus || 502, msg, dataStatus === 404 ? 'resolved' : 'queued');
      throw new Error(msg);
    }
    markProxySessionHealthy();
    return data;
  }
}

function buildProxyActionUrl(action, params) {
  if (!API_PROXY) throw new Error('No proxy configured');
  var sep = API_PROXY.indexOf('?') !== -1 ? '&' : '?';
  var url = API_PROXY + sep + 'action=' + encodeURIComponent(action);
  if (params) {
    Object.keys(params).forEach(function(k) {
      if (params[k] === undefined || params[k] === null) return;
      url += '&' + encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
    });
  }
  return url;
}

// POST helper for sensitive admin actions where secrets must NOT appear in the
// URL (preventing exposure in server logs, browser history, CDN traces).
// Sends { action } in the query string only; all payload (including key/secret)
// travels as a JSON body. Attaches the same bearer token as proxyAction.
async function proxyPost(action, bodyObj, extraHeaders) {
  var readOnlyAllowedWrites = {
    pm_notifications_ack: true
  };
  if (isReadOnlyAccessMode() && !readOnlyAllowedWrites[action]) {
    throw new Error('Read-only access mode: updates are disabled');
  }
  if (!API_PROXY) throw new Error('No proxy configured');
  var sep = API_PROXY.indexOf('?') !== -1 ? '&' : '?';
  var url = API_PROXY + sep + 'action=' + encodeURIComponent(action);
  var skipAuthActions = {
    verify_role: true,
    device_setup: true,
    device_otp_request: true,
    device_otp_verify: true
  };
  var token = skipAuthActions[action] ? '' : getProxyAccessToken();
  var headers = Object.assign({ 'Content-Type': 'application/json', 'Accept': 'application/json' }, extraHeaders || {});
  if (token) headers['Authorization'] = 'Bearer ' + token;
  var maxRetries = 2;
  for (var attempt = 0; attempt <= maxRetries; attempt++) {
    var res = await fetchWithTimeout(url, {
      method: 'POST',
      headers: headers,
      body: JSON.stringify(bodyObj || {})
    }, 30000);
    if (res.status === 429) {
      var retryAfter = parseInt(res.headers.get('Retry-After') || '5', 10);
      rateLimiter.backpressure(retryAfter * 1000);
      logApiError(429, 'POST action=' + action + ' rate limited — pausing ' + retryAfter + 's', 'retry');
      if (attempt < maxRetries) { await sleep(retryAfter * 1000); continue; }
      throw new Error('429: Rate limited on POST ' + action + ' after retries');
    }
    if ((res.status === 502 || res.status === 503 || res.status === 504) && attempt < maxRetries) {
      var backoff = Math.pow(2, attempt + 1) * 1000; // 2s, 4s
      logApiError(res.status, 'POST action=' + action + ' HTTP ' + res.status + ' — retrying in ' + (backoff / 1000) + 's', 'retry');
      await sleep(backoff);
      continue;
    }
    if (!res.ok) {
      var errBody = '';
      try { errBody = await res.text(); } catch (e) { /* empty */ }
      if (res.status === 401) {
        var errBodyLower = String(errBody || '').toLowerCase();
        if (errBodyLower.indexOf('frontend bearer token') !== -1 || errBodyLower.indexOf('proxy session') !== -1) {
          handleProxySessionExpired('post=' + action);
          throw new Error('Proxy POST action=' + action + ' failed: HTTP 401 — Proxy session missing or expired. Sign in again.');
        }
      }
      throw new Error('Proxy POST action=' + action + ' failed: HTTP ' + res.status + (errBody ? ' \u2014 ' + errBody.substring(0, 200) : ''));
    }
    var postData = await res.json();
    markProxySessionHealthy();
    return postData;
  }
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
  _defaultMaxPerSec: 4,
  windowStart: 0,
  windowCount: 0,
  processing: false,
  // Global pause gate: epoch-ms when queue may resume after a 429
  pauseUntil: 0,

  // Called whenever any request sees a 429. Freezes the queue and halves
  // dispatch rate for the pause window to reduce burst pressure.
  backpressure: function(ms) {
    var resumeAt = Date.now() + ms;
    if (resumeAt > rateLimiter.pauseUntil) {
      rateLimiter.pauseUntil = resumeAt;
    }
    // Halve rate during backpressure, floor at 1 req/s
    rateLimiter.maxPerSec = Math.max(1, Math.floor(rateLimiter._defaultMaxPerSec / 2));
    updateRateBadge();
    // Restore rate automatically after pause window clears
    setTimeout(function() {
      if (Date.now() >= rateLimiter.pauseUntil) {
        rateLimiter.maxPerSec = rateLimiter._defaultMaxPerSec;
        updateRateBadge();
      }
    }, ms + 150);
  },

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

      // Global 429 backpressure: hold the whole queue until the window clears
      if (now < rateLimiter.pauseUntil) {
        var pauseWait = rateLimiter.pauseUntil - now + 50;
        setTimeout(tick, pauseWait);
        return;
      }

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

    // ═══════════════════════════════════════════════════════
    // PM INBOX DRAWER — manager posting + PM acknowledgement
    // ═══════════════════════════════════════════════════════
    (function initPmInboxDrawer() {
      var drawer = document.getElementById('pm-inbox-drawer');
      var fab = document.getElementById('pm-inbox-fab');
      var items = document.getElementById('pm-inbox-items');
      var status = document.getElementById('pm-inbox-status');
      var badge = document.getElementById('pm-inbox-badge');
      var fabBadge = document.getElementById('pm-fab-badge');
      var composerWrap = document.getElementById('pmInboxComposerWrap');
      var composerText = document.getElementById('pmInboxComposerText');
      var composerScope = document.getElementById('pmInboxComposerScope');
      var composerSend = document.getElementById('pmInboxComposerSend');
      if (!drawer || !fab || !items || !status || !badge || !fabBadge) return;

      var _rows = [];
      var _pollTimer = null;
      var _pollMs = 25000;
      var _unread = 0;

      function isInboxRoleAllowed() {
        return _accessRole === 'manager' || _accessRole === 'pm_readonly' || _accessRole === 'full';
      }

      function isComposerAllowed() {
        return _accessRole === 'manager' || _accessRole === 'full';
      }

      function isOpen() {
        return drawer.classList.contains('open');
      }

      function setStatus(text) {
        status.innerHTML = '<span class="feed-pulse"></span> ' + escapeHtml(text || '');
      }

      function updateBadge() {
        var n = Math.max(0, Number(_unread || 0) || 0);
        var txt = n > 99 ? '99+' : String(n);
        badge.textContent = txt;
        fabBadge.textContent = txt;
        badge.style.display = n > 0 ? '' : 'none';
        fabBadge.style.display = n > 0 ? '' : 'none';
      }

      function renderRows() {
        if (!_rows.length) {
          items.innerHTML = '<p class="feed-empty">No inbox notices yet.</p>';
          return;
        }

        var html = '';
        _rows.forEach(function(row) {
          var by = String(row.created_by_user || row.created_by_role || 'Manager').trim();
          var ts = String(row.created_at || '').trim();
          var scope = String(row.scope_group_uuid || '').trim();
          var canAck = !row.read_by_me;
          html += '<div class="feed-item" style="border-left:3px solid ' + (row.read_by_me ? 'var(--success)' : 'var(--warning)') + '">';
          html += '<div class="feed-item-header">';
          html += '<span style="color:' + (row.read_by_me ? 'var(--success)' : 'var(--warning)') + ';font-size:.9rem;flex-shrink:0;width:16px;text-align:center"><i class="fas ' + (row.read_by_me ? 'fa-check-circle' : 'fa-bell') + '"></i></span>';
          html += '<span class="feed-item-title">' + escapeHtml(by) + (scope ? ' · Scoped' : ' · Global') + '</span>';
          html += '<span class="feed-item-time">' + (ts ? timeAgo(ts) : '') + '</span>';
          html += '</div>';
          html += '<div class="feed-item-body">' + escapeHtml(String(row.message || '').substring(0, 600)) + '</div>';
          if (canAck) {
            html += '<div style="margin-top:6px"><button class="action-btn" data-pm-ack="' + escapeHtml(String(row.uuid || '')) + '" style="font-size:10px;padding:2px 8px"><i class="fas fa-check" style="margin-right:4px"></i>Acknowledge</button></div>';
          } else {
            html += '<div style="margin-top:6px;font-size:10px;color:var(--success)"><i class="fas fa-check" style="margin-right:4px"></i>Read' + (row.read_at ? ' · ' + escapeHtml(timeAgo(row.read_at)) : '') + '</div>';
          }
          html += '</div>';
        });
        items.innerHTML = html;
      }

      function populateScopeOptions() {
        if (!composerScope) return;
        if (!isComposerAllowed()) return;
        var selected = String(composerScope.value || '').trim();
        var options = ['<option value="">All PMs (Global)</option>'];

        if (_accessRole === 'manager' && forcedPropertyGroupUuid) {
          var forcedLabel = resolveGroupNameFromUuid(forcedPropertyGroupUuid) || 'My Group';
          options = ['<option value="' + escapeHtml(forcedPropertyGroupUuid) + '">' + escapeHtml(forcedLabel) + ' (Manager Scope)</option>'];
          composerScope.innerHTML = options.join('');
          composerScope.value = forcedPropertyGroupUuid;
          composerScope.disabled = true;
          return;
        }

        composerScope.disabled = false;
        (PROPERTY_GROUPS || []).forEach(function(g) {
          var id = String(g.id || g.group_id || '').trim();
          var nm = String(g.name || g.group_name || '').trim();
          if (!id || !nm) return;
          options.push('<option value="' + escapeHtml(id) + '">' + escapeHtml(nm) + '</option>');
        });
        composerScope.innerHTML = options.join('');
        composerScope.value = selected || '';
      }

      async function refreshInbox() {
        if (!isInboxRoleAllowed()) return;
        if (!isProxySessionReady()) {
          setStatus('Awaiting sign-in…');
          return;
        }
        try {
          var data = await proxyAction('pm_notifications_inbox', { limit: '80' }, { suppressSessionExpiry: true });
          _rows = Array.isArray(data.notifications) ? data.notifications : [];
          _unread = Number(data.unread || 0) || 0;
          updateBadge();
          renderRows();
          setStatus(isOpen() ? 'Inbox synced' : ('Inbox synced · ' + _unread + ' unread'));
          populateScopeOptions();
        } catch (err) {
          setStatus('Inbox unavailable. Retrying…');
          console.debug('pm inbox poll skipped:', err && err.message ? err.message : err);
        }
      }

      async function submitPost() {
        if (!isComposerAllowed() || !composerText) return;
        var message = String(composerText.value || '').trim();
        if (!message) {
          showToast('Enter a message before posting', { kind: 'warning' });
          return;
        }
        if (message.length > 1200) {
          showToast('Message is too long (max 1200 chars)', { kind: 'warning' });
          return;
        }
        var scopeId = String((composerScope && composerScope.value) || '').trim();
        if (_accessRole === 'manager' && forcedPropertyGroupUuid) {
          scopeId = forcedPropertyGroupUuid;
        }
        try {
          if (composerSend) composerSend.disabled = true;
          await proxyPost('pm_notifications_post', {
            message: message,
            scope_group_uuid: scopeId
          });
          composerText.value = '';
          showToast('Inbox notice posted', { kind: 'success' });
          await refreshInbox();
        } catch (err) {
          showToast('Post failed: ' + (err.message || err), { kind: 'warning' });
        } finally {
          if (composerSend) composerSend.disabled = false;
        }
      }

      async function ackNotification(uuid) {
        if (!uuid) return;
        try {
          await proxyPost('pm_notifications_ack', { notification_uuid: uuid });
          await refreshInbox();
        } catch (err) {
          showToast('Acknowledge failed: ' + (err.message || err), { kind: 'warning' });
        }
      }

      function applyVisibility() {
        var show = isInboxRoleAllowed();
        fab.style.display = show ? '' : 'none';
        drawer.style.display = show ? '' : 'none';
        if (!show) return;
        if (composerWrap) composerWrap.style.display = isComposerAllowed() ? '' : 'none';
      }

      if (composerSend) {
        composerSend.addEventListener('click', function() { submitPost(); });
      }
      if (composerText) {
        composerText.addEventListener('keydown', function(e) {
          if (!(e.ctrlKey || e.metaKey) || e.key !== 'Enter') return;
          e.preventDefault();
          submitPost();
        });
      }

      if (items) {
        items.addEventListener('click', function(e) {
          var btn = e.target.closest('[data-pm-ack]');
          if (!btn) return;
          var id = String(btn.getAttribute('data-pm-ack') || '').trim();
          ackNotification(id);
        });
      }

      window.PmInbox = {
        refresh: refreshInbox,
        clearUnseen: function() { refreshInbox(); },
        applyVisibility: applyVisibility,
      };

      function boot() {
        applyVisibility();
        refreshInbox();
        if (_pollTimer) clearInterval(_pollTimer);
        _pollTimer = setInterval(function() {
          applyVisibility();
          if (!document.hidden) refreshInbox();
        }, _pollMs);
      }

      if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', boot);
      } else {
        boot();
      }
    })();
  }
};

function updateRateBadge() {
  var el = $('#rateBadge');
  if (!el) return;
  if (rateLimiter.pauseUntil > Date.now()) {
    var secsLeft = Math.ceil((rateLimiter.pauseUntil - Date.now()) / 1000);
    el.textContent = '\u23f8 paused ' + secsLeft + 's';
    el.style.color = 'var(--warning)';
  } else {
    el.textContent = (rateLimiter.maxPerSec - rateLimiter.inFlight) + '/' + rateLimiter.maxPerSec + ' req/s';
    el.style.color = '';
  }
}

// Core fetch wrapper with auth, retries, and error logging
async function apiFetch(path, options) {
  var methodCheck = ((options && options.method) || 'GET').toUpperCase();
  if (isReadOnlyAccessMode() && methodCheck !== 'GET' && methodCheck !== 'HEAD') {
    throw new Error('Read-only access mode: updates are disabled');
  }
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
        // Freeze the shared queue so all concurrent requests back off together
        rateLimiter.backpressure(retryAfter * 1000);
        logApiError(429, 'Rate limit exceeded — Retry-After: ' + retryAfter + 's (queue paused)', 'retry');
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

function setSectionBusy(sectionId, busy, message) {
  var section = document.getElementById(sectionId);
  if (!section) return;

  var overlay = section.querySelector('.section-loading-overlay');
  if (!busy) {
    section.classList.remove('section-busy');
    section.removeAttribute('aria-busy');
    if (overlay) overlay.remove();
    return;
  }

  section.classList.add('section-busy');
  section.setAttribute('aria-busy', 'true');
  if (!overlay) {
    overlay = document.createElement('div');
    overlay.className = 'section-loading-overlay';
    overlay.innerHTML =
      '<div class="section-loading-panel">' +
        '<i class="fas fa-spinner fa-spin" aria-hidden="true"></i>' +
        '<span class="section-loading-text"></span>' +
      '</div>';
    section.appendChild(overlay);
  }
  var textEl = overlay.querySelector('.section-loading-text');
  if (textEl) textEl.textContent = String(message || 'Loading…');
}

function applyProxySchemaHealth(pingData) {
  var el = $('#apiStatus');
  if (!el) return { suffix: '', hasIssue: false, detail: '' };
  var schema = pingData && pingData.schema;
  if (!schema || typeof schema !== 'object') {
    return { suffix: '', hasIssue: false, detail: '' };
  }
  var missing = Array.isArray(schema.missing_tables) ? schema.missing_tables : [];
  var hasIssue = schema.ok === false || missing.length > 0;
  var detail = hasIssue
    ? ('Schema issue: missing ' + (missing.length ? missing.join(', ') : 'unknown table(s)'))
    : 'Schema OK';
  var dbLabel = (pingData && pingData.database) ? (' [' + pingData.database + ']') : '';
  el.title = detail + dbLabel;
  return {
    suffix: hasIssue ? ' · schema issue' : ' · schema ok',
    hasIssue: hasIssue,
    detail: detail,
  };
}

/* =================================================================
   API Error Log
   ================================================================= */
var API_ERRORS = [];
var SYSTEM_HEALTH_STATE = {
  running: false,
  data: null,
  lastError: '',
};
var _systemHealthAutoRunDone = false;

function logApiError(code, msg, action) {
  var now = new Date();
  var ts = now.getHours().toString().padStart(2, '0') + ':' + now.getMinutes().toString().padStart(2, '0') + ':' + now.getSeconds().toString().padStart(2, '0');
  API_ERRORS.unshift({ code: code, ts: ts, msg: msg, action: action });
  if (API_ERRORS.length > 100) API_ERRORS.length = 100;
  renderErrorLog();
}

function canRunSystemHealthChecker() {
  return _accessRole === 'full' || _accessRole === 'manager';
}

function getSystemHealthStatusLabel(status) {
  var value = String(status || '').toLowerCase();
  if (value === 'green') return 'Verified Working';
  if (value === 'yellow') return 'Possible Issues';
  if (value === 'red') return 'Issue Confirmed';
  return 'Idle';
}

function getSystemHealthSummaryText() {
  if (SYSTEM_HEALTH_STATE.running) return 'Running checks...';
  if (SYSTEM_HEALTH_STATE.lastError) return 'Checker failed: ' + SYSTEM_HEALTH_STATE.lastError;
  if (!SYSTEM_HEALTH_STATE.data) return 'No checks run yet';
  var summary = SYSTEM_HEALTH_STATE.data.summary || {};
  var stamp = SYSTEM_HEALTH_STATE.data.generated_at
    ? (' @ ' + String(SYSTEM_HEALTH_STATE.data.generated_at).replace('T', ' ').replace('Z', ' UTC'))
    : '';
  return getSystemHealthStatusLabel(SYSTEM_HEALTH_STATE.data.status) +
    ' - red: ' + Number(summary.red_count || 0) +
    ', yellow: ' + Number(summary.yellow_count || 0) +
    ', green: ' + Number(summary.green_count || 0) + stamp;
}

function renderSystemHealthPanel() {
  var panel = $('#systemsHealthPanel');
  if (!panel) return;

  if (!canRunSystemHealthChecker()) {
    panel.style.display = 'none';
    return;
  }
  panel.style.display = '';

  var summaryEl = $('#systemHealthSummary');
  var gridEl = $('#systemHealthGrid');
  var sendBtn = $('#btnSendSystemHealthDebug');

  if (summaryEl) {
    var status = SYSTEM_HEALTH_STATE.running
      ? 'idle'
      : String((SYSTEM_HEALTH_STATE.data && SYSTEM_HEALTH_STATE.data.status) || '').toLowerCase();
    if (status !== 'green' && status !== 'yellow' && status !== 'red') status = 'idle';
    summaryEl.innerHTML =
      '<span class="systems-health-dot ' + status + '"></span>' +
      '<span>' + escapeHtml(getSystemHealthSummaryText()) + '</span>';
  }

  if (sendBtn) {
    sendBtn.style.display = SYSTEM_HEALTH_STATE.data ? '' : 'none';
  }

  if (!gridEl) return;
  if (!SYSTEM_HEALTH_STATE.data || !Array.isArray(SYSTEM_HEALTH_STATE.data.checks) || !SYSTEM_HEALTH_STATE.data.checks.length) {
    gridEl.innerHTML = '';
    return;
  }

  var html = '';
  SYSTEM_HEALTH_STATE.data.checks.forEach(function(check) {
    var status = String(check.status || 'yellow').toLowerCase();
    if (status !== 'green' && status !== 'yellow' && status !== 'red') status = 'yellow';
    html += '<div class="systems-health-card ' + status + '">';
    html += '<div class="systems-health-card-head">';
    html += '<div class="systems-health-card-title">' + escapeHtml(String(check.label || check.key || 'Check')) + '</div>';
    html += '<span class="systems-health-badge ' + status + '">' + escapeHtml(status) + '</span>';
    html += '</div>';
    html += '<div class="systems-health-card-detail">' + escapeHtml(String(check.detail || '')) + '</div>';
    html += '</div>';
  });
  gridEl.innerHTML = html;
}

async function runSystemHealthCheck() {
  if (!canRunSystemHealthChecker()) {
    showToast('System checker is restricted to manager/admin', {
      kind: 'warning',
      iconClass: 'fa-lock',
      duration: 2600
    });
    return;
  }
  if (SYSTEM_HEALTH_STATE.running) return;

  SYSTEM_HEALTH_STATE.running = true;
  SYSTEM_HEALTH_STATE.lastError = '';
  renderSystemHealthPanel();

  try {
    var data = await proxyAction('system_health');
    if (!data || !data.ok) {
      throw new Error((data && data.error) || 'system_health failed');
    }
    SYSTEM_HEALTH_STATE.data = data;
    var status = String(data.status || '').toLowerCase();
    if (status === 'green') {
      showToast('Systems checker: Verified Working', { kind: 'success', iconClass: 'fa-circle-check', duration: 2600 });
    } else if (status === 'yellow') {
      showToast('Systems checker: possible issues detected - please wait', { kind: 'warning', iconClass: 'fa-triangle-exclamation', duration: 3600 });
    } else {
      showToast('Systems checker: issue confirmed. Use Send Debug.', { kind: 'danger', iconClass: 'fa-circle-exclamation', duration: 4400 });
    }
  } catch (err) {
    SYSTEM_HEALTH_STATE.lastError = String((err && err.message) || err || 'unknown error');
    showToast('Systems checker failed: ' + SYSTEM_HEALTH_STATE.lastError, {
      kind: 'danger',
      iconClass: 'fa-circle-exclamation',
      duration: 4200
    });
  } finally {
    SYSTEM_HEALTH_STATE.running = false;
    renderSystemHealthPanel();
  }
}

function sendSystemHealthDebugEmail() {
  var report = SYSTEM_HEALTH_STATE.data;
  if (!report) {
    showToast('Run checker first to send debug payload', {
      kind: 'warning',
      iconClass: 'fa-envelope-open',
      duration: 2500
    });
    return;
  }

  var status = String(report.status || 'unknown').toUpperCase();
  var subject = '[HandyManager] System Health Debug - ' + status;
  var payload = report && report.debug && report.debug.payload
    ? report.debug.payload
    : report;
  var payloadText = '';
  try {
    payloadText = JSON.stringify(payload, null, 2);
  } catch (e) {
    payloadText = String(payload || 'payload_unavailable');
  }
  if (payloadText.length > 14000) {
    payloadText = payloadText.slice(0, 14000) + '\n... [truncated]';
  }

  var body = [
    'Aaron,',
    '',
    'Auto-generated HandyManager system health debug package.',
    '',
    'Overall status: ' + status,
    'Generated at: ' + String(report.generated_at || ''),
    '',
    'Suggested debug SQL:',
    Array.isArray(report && report.debug && report.debug.debug_query)
      ? report.debug.debug_query.join('\n')
      : 'n/a',
    '',
    'Payload:',
    payloadText,
  ].join('\n');

  var mailto = 'mailto:aaron@flraz.com?subject=' + encodeURIComponent(subject) + '&body=' + encodeURIComponent(body);
  window.location.href = mailto;
}

function maybeAutoRunSystemHealthCheck() {
  if (_systemHealthAutoRunDone) return;
  if (!canRunSystemHealthChecker()) return;
  _systemHealthAutoRunDone = true;
  setTimeout(function() {
    runSystemHealthCheck().catch(function() {});
  }, 300);
}

/* =================================================================
   DATA STORES — populated from live API
   ================================================================= */
var WORK_ORDERS = [];
var VENDORS = [];
var PROPERTIES = [];
var PROPERTY_GROUPS = [];
var TURNS = [];
var ROSTER = [];
var activeGroupId = '';
var UPCOMING_MOVEOUTS = []; // from tenant_directory — tenants on notice
var TURN_WORK_ORDERS = []; // from DB API — unit turn WOs with real-time status
var UNIT_TURNS_DB = [];    // SQL-backed unit turn tracker records
var UNIT_TURN_TRACKER_BY_KEY = {};
var _turnTrackerSyncInFlight = false;
var _lastTurnTrackerSyncHash = '';
var _turnAutoCloseSyncInFlight = false;
var _lastTurnAutoCloseHash = '';
var UNIT_TURN_HISTORY = [];
var BILLS = []; // from DB API — AP bills for WO close-assist
window._currentBillsCache = []; // currently visible bills for fast detail-card lookup
window._billingPageRows = [];
var _billsFetchInFlight = false;
var _lastBillSource = 'legacy';

// ── Billing cache for client-side filtering ──────────────────────────────────
var __CACHED_BILLS = []; // Full unfiltered bills array (loaded once, reused for all filters)
var __PROPERTY_TO_GROUP_MAP = {}; // property_id -> group_id lookup for instant filtering
var __CACHED_BILLS_LOADED_AT = 0; // Timestamp of when cache was populated
var __CACHED_BILLS_IN_FLIGHT = false; // Lock to prevent duplicate fetches
var _billingServerTotal = 0;
var _billingServerTotalPages = 1;
var _billingServerPage = 1;
var _billingRouteAction = 'bills_list';
var BILL_ROUTE_COLUMNS = [
  'id',
  'bill_number',
  'status',
  'status_label',
  'vendor_name',
  'property_name',
  'invoice_date',
  'due_date',
  'bill_total_amount'
].join(',');
var _billingRouteFilterValue = '';
var _billingDueFrom = '';
var _billingDueTo = '';
var _billingRouteStatus = '';
var RECENT_TASKS = [];
var WEBHOOK_EVENTS = [];
var _schemaHealthWarned = false;
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
window.AppDB = window.AppDB || {
  vendors: new Map(),
  properties: new Map(),
  propertyGroups: new Map(),
};
var appInitialized = false;
var WO_FLAGS = {};
var WO_DETAIL_CACHE = {};
var WO_DETAIL_CACHE_KEYS = []; // LRU order tracker
var WO_DETAIL_CACHE_MAX = 50;  // max cached entries
var CURRENT_WO_MODAL = null;
var PAYROLL_WEEK_OFFSET = 0;
var currentPropertyGroup = '';
var _propertiesLocalGroup = '';
var _propertiesRefreshInFlight = false;
var _propertiesPage = 0;
var _propertiesPageSize = 50;
var _propertiesVacancyOnly = false;
var _propertyStatsById = {};
window.filteredPropertyId = '';
window.filteredPropertyName = '';
window.filteredUnitId = '';
window.filteredUnitName = '';
var UNITS = [];
var _unitsByPropertyId = {};  // propertyId -> Unit[]
var forcedPropertyGroupUuid = '';
var forcedPropertyGroupName = '';
var currentTurnFilter = 'open';
var currentWOCloseAssistAge = 14;
var _billsLoading = false;
var _billsLoadedAt = 0;
var _vendorRenderLimit = 0;
var currentVendorInitial = '';
var _vendorRenderKey = '';
var _vendorsNeedRender = false;
var ROUTING_EVENTS = [];
var ROUTING_PM_STATS = [];
var ROUTING_CAPABILITIES = [];
var ROUTING_PM_MAP = {};
var _routingInitDone = false;
var _routingFiltersDelegated = false;
var _routingSearchDebounceTimer = 0;
var _routingLastScanAt = '';
var ROUTING_PAGE = 1;
var ROUTING_SETTINGS = {
  minConfidence: 'medium',
  highThreshold: 2,
  pageSize: 25
};

function loadRoutingSettings() {
  try {
    var raw = localStorage.getItem('hm_routing_settings');
    if (!raw) return;
    var parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') return;
    if (parsed.minConfidence) ROUTING_SETTINGS.minConfidence = String(parsed.minConfidence);
    var high = Number(parsed.highThreshold || 2);
    ROUTING_SETTINGS.highThreshold = isNaN(high) ? 2 : Math.max(2, high);
    var pageSize = Number(parsed.pageSize || 25);
    ROUTING_SETTINGS.pageSize = (pageSize === 50 || pageSize === 100) ? pageSize : 25;
  } catch (e) { /* */ }
}

function saveRoutingSettings() {
  try {
    localStorage.setItem('hm_routing_settings', JSON.stringify(ROUTING_SETTINGS));
  } catch (e) { /* */ }
}

loadRoutingSettings();

try { forcedPropertyGroupUuid = localStorage.getItem('hm_scope_group_uuid') || ''; } catch (e) { /* */ }
try { _pmScopeEmail = localStorage.getItem('hm_scope_email') || ''; } catch (e) { /* */ }

// Show returning-PM identity hint on the login screen before user connects
(function() {
  try {
    var storedEmail = _pmScopeEmail || '';
    var storedToken = '';
    try { storedToken = localStorage.getItem('hm_device_token') || ''; } catch (e) { /* */ }
    if (!storedEmail || !storedToken) return;
    // Derive a display name from email prefix: "alex.smith" → "Alex Smith"
    var prefix = storedEmail.split('@')[0] || storedEmail;
    var displayName = prefix.split(/[.\-_]/).map(function(p) {
      return p ? (p.charAt(0).toUpperCase() + p.slice(1).toLowerCase()) : '';
    }).filter(Boolean).join(' ') || storedEmail;
    // Populate the pre-existing hint pill
    var hint = document.getElementById('vaultUserHint');
    if (hint) {
      hint.textContent = '\u{1F464} ' + displayName + '  \u2022  ' + storedEmail;
      hint.style.display = '';
    }
    // Update subtitle for returning PM context
    var sub = document.getElementById('vaultSubtitle');
    if (sub) sub.textContent = 'Welcome back, ' + displayName + '. Tap Connect to resume.';
  } catch (e) { /* non-fatal */ }
})();

function resolveGroupNameFromUuid(groupUuid) {
  var gid = String(groupUuid || '').trim();
  if (!gid) return '';
  var g = (PROPERTY_GROUPS || []).find(function(x) {
    return String((x && x.id) || '').trim() === gid;
  });
  return g && g.name ? String(g.name) : '';
}

function resolveGroupUuidFromName(groupName) {
  var target = String(groupName || '').trim().toLowerCase();
  if (!target) return '';

  // If caller already passed a UUID, trust it directly.
  if (isUuidString(target)) return target;

  if (_accessRole === 'pm_readonly' && forcedPropertyGroupUuid) {
    return String(forcedPropertyGroupUuid).trim();
  }

  var g = (PROPERTY_GROUPS || []).find(function(x) {
    var gName = String((x && x.name) || '').trim().toLowerCase();
    var gId = String((x && x.id) || '').trim().toLowerCase();
    return gName === target || gId === target;
  });
  return g && g.id ? String(g.id).trim() : '';
}

function buildReferenceMaps(vendorsArray, propertiesArray) {
  var db = window.AppDB || { vendors: new Map(), properties: new Map(), propertyGroups: new Map() };
  if (!db.vendors || typeof db.vendors.set !== 'function') db.vendors = new Map();
  if (!db.properties || typeof db.properties.set !== 'function') db.properties = new Map();
  if (!db.propertyGroups || typeof db.propertyGroups.set !== 'function') db.propertyGroups = new Map();

  db.vendors.clear();
  db.properties.clear();
  db.propertyGroups.clear();

  (Array.isArray(vendorsArray) ? vendorsArray : []).forEach(function(v) {
    var keys = [
      v.id,
      v.Id,
      v.vendor_id,
      v.VendorId,
      v.vendorId,
      v.uuid,
      v.vendor_uuid,
      v.VendorUuid,
    ];
    keys.forEach(function(k) {
      var key = String(k || '').trim();
      if (!key) return;
      db.vendors.set(key, v);
    });
  });

  (Array.isArray(propertiesArray) ? propertiesArray : []).forEach(function(p) {
    var propIdKeys = [p.id, p.Id, p.property_id, p.PropertyId, p.propertyId, p.property_uuid, p.PropertyUuid, p.uuid];
    propIdKeys.forEach(function(k) {
      var key = String(k || '').trim();
      if (!key) return;
      db.properties.set(key, p);
    });

    var groupId = String(p.propertyGroupId || p.PropertyGroupId || p.property_group_id || '').trim();
    if (!groupId) return;
    if (!db.propertyGroups.has(groupId)) {
      db.propertyGroups.set(groupId, { id: groupId, properties: [] });
    }
    var normalizedPropId = String(p.id || p.property_id || p.propertyId || p.Id || '').trim();
    if (normalizedPropId) {
      db.propertyGroups.get(groupId).properties.push(normalizedPropId);
    }
  });

  window.AppDB = db;
  
  // Populate property-to-group map for instant bill filtering
  __PROPERTY_TO_GROUP_MAP = {};
  (Array.isArray(propertiesArray) ? propertiesArray : []).forEach(function(p) {
    var propId = String(p.id || p.Id || p.property_id || p.PropertyId || p.propertyId || '').trim();
    var groupId = String(p.propertyGroupId || p.PropertyGroupId || p.property_group_id || p.property_group_uuid || p.PropertyGroupUuid || '').trim();
    if (propId && groupId) {
      __PROPERTY_TO_GROUP_MAP[propId] = groupId;
    }
  });
  
  return db;
}

function normalizeBillApprovalStatus(value) {
  var raw = String(value || '').trim().toLowerCase().replace(/\s+/g, '_');
  if (!raw) return '';
  if (
    raw === 'pending' ||
    raw === 'pending_approval' ||
    raw === 'awaiting_approval' ||
    raw === 'needs_approval' ||
    raw === 'pending_review' ||
    raw === 'awaiting_review' ||
    raw === 'needs_review' ||
    raw === 'submitted_for_approval' ||
    raw === 'requires_approval' ||
    raw === 'unapproved'
  ) return 'pending_approval';
  if (raw === 'partially_paid' || raw === 'partial_payment' || raw === 'paid_partially') return 'paid';
  if (raw === 'approved_for_payment') return 'approved';
  if (raw === 'canceled' || raw === 'cancelled' || raw === 'voided') return 'void';
  return raw;
}

function getBillStatusKey(recordLike) {
  var b = (recordLike && typeof recordLike === 'object') ? recordLike : {};
  var raw = (b.raw && typeof b.raw === 'object') ? b.raw : b;
  var candidates = [
    b.status,
    b.statusLabel,
    b.approval_status,
    b.approvalStatus,
    b.bill_status,
    b.billStatus,
    raw.ApprovalStatus,
    raw.approval_status,
    raw.ApprovalState,
    raw.approval_state,
    raw.PendingApproval,
    raw.pending_approval,
    raw.Status,
    raw.status,
    raw.State,
    raw.state,
  ];
  for (var i = 0; i < candidates.length; i++) {
    var normalized = normalizeBillApprovalStatus(candidates[i]);
    if (normalized) return normalized;
  }
  return '';
}

function extractBillPropertyId(rawBill, normalizedLineItems) {
  var lineItems = Array.isArray(normalizedLineItems) ? normalizedLineItems : [];
  var first = lineItems[0] || {};
  return String(
    rawBill.property_id || rawBill.PropertyId || rawBill.propertyId || rawBill.property_uuid || rawBill.PropertyUuid ||
    first.PropertyId || first.property_id || first.propertyId || first.PropertyUuid || first.property_uuid || ''
  ).trim();
}

function extractBillUnitId(rawBill, normalizedLineItems) {
  var lineItems = Array.isArray(normalizedLineItems) ? normalizedLineItems : [];
  var first = lineItems[0] || {};
  return String(
    rawBill.unit_id || rawBill.UnitId || rawBill.unitId || rawBill.UnitUUID || rawBill.unit_uuid ||
    first.UnitId || first.unit_id || first.unitId || first.UnitUUID || first.unit_uuid || ''
  ).trim();
}

function getNormalizedBillDisplayNumber(billLike) {
  var b = billLike && typeof billLike === 'object' ? billLike : {};
  var raw = (b.raw && typeof b.raw === 'object') ? b.raw : b;
  var preferred = String(
    b.billNumber || b.bill_number || b.reference ||
    raw.InvoiceNumber || raw.invoice_number || raw.Reference || raw.reference || ''
  ).trim();

  if (preferred && !isUuidString(preferred)) return preferred;

  var fallbackRef = String(raw.Reference || raw.reference || '').trim();
  if (fallbackRef && !isUuidString(fallbackRef)) return fallbackRef;

  var fallbackInvoice = String(raw.InvoiceNumber || raw.invoice_number || '').trim();
  if (fallbackInvoice && !isUuidString(fallbackInvoice)) return fallbackInvoice;

  var id = String(b.id || raw.Id || raw.id || '').trim();
  if (!id) return '—';
  if (isUuidString(id)) return 'BILL-' + id.slice(0, 8).toUpperCase();
  return id;
}

function getNormalizedBillDetailId(billLike) {
  var b = billLike && typeof billLike === 'object' ? billLike : {};
  var raw = (b.raw && typeof b.raw === 'object') ? b.raw : b;
  var candidates = [
    b.id,
    b.bill_id,
    b.BillId,
    raw.BillId,
    raw.bill_id,
    raw.Id,
    raw.id,
  ];
  for (var i = 0; i < candidates.length; i++) {
    var value = String(candidates[i] || '').trim();
    if (value) return value;
  }
  return '';
}

function resolveBillAmountValue(billLike) {
  var b = billLike && typeof billLike === 'object' ? billLike : {};
  var raw = (b.raw && typeof b.raw === 'object') ? b.raw : b;
  var candidates = [
    b.amount,
    b.total_amount,
    b.totalAmount,
    b.net_amount,
    b.netAmount,
    b.amount_due,
    b.amountDue,
    raw.TotalAmount,
    raw.total_amount,
    raw.Amount,
    raw.amount,
    raw.NetAmount,
    raw.net_amount,
    raw.AmountDue,
    raw.amount_due,
    raw.Balance,
    raw.balance,
    raw.Unpaid,
    raw.unpaid,
    raw.Paid,
    raw.paid,
  ];

  for (var i = 0; i < candidates.length; i++) {
    var n = amountToNumber(candidates[i]);
    if (n !== 0) return n;
  }

  var lineItems = Array.isArray(b.lineItems)
    ? b.lineItems
    : (Array.isArray(raw.LineItems) ? raw.LineItems : (Array.isArray(raw.line_items) ? raw.line_items : []));
  if (lineItems.length) {
    var sum = lineItems.reduce(function(total, li) {
      return total + amountToNumber(
        (li && (li.Amount != null ? li.Amount : li.amount)) ||
        (li && (li.TotalAmount != null ? li.TotalAmount : li.total_amount)) ||
        0
      );
    }, 0);
    if (sum !== 0) return sum;
  }

  return 0;
}

function resolveVendorNameFromMaps(vendorId, fallback) {
  var key = String(vendorId || '').trim();
  if (!key || !window.AppDB || !window.AppDB.vendors) return String(fallback || '').trim();
  var vendor = window.AppDB.vendors.get(key);
  if (!vendor) return String(fallback || '').trim();
  return String(
    vendor.name || vendor.Name || vendor.company_name || vendor.CompanyName ||
    ((vendor.first_name || vendor.FirstName || '') + ' ' + (vendor.last_name || vendor.LastName || '')).trim() ||
    fallback || ''
  ).trim();
}

function resolvePropertyMetaFromMaps(propertyId, fallbackName, fallbackGroupId) {
  var key = String(propertyId || '').trim();
  var fallback = {
    id: key,
    name: String(fallbackName || '').trim(),
    groupId: String(fallbackGroupId || '').trim(),
    groupName: '',
    siteManager: '',
  };
  var prop = null;
  if (key && window.AppDB && window.AppDB.properties) {
    prop = window.AppDB.properties.get(key) || null;
  }

  // Fallback lookup by property name (handles cases where bills provide name but not id).
  if (!prop && fallback.name) {
    var fbLower = String(fallback.name || '').trim().toLowerCase();
    var fbNorm = normalizePropertyLookupKey(fallback.name || '');
    prop = (PROPERTIES || []).find(function(p) {
      var pName = String((p && p.name) || '').trim().toLowerCase();
      if (!pName) return false;
      if (pName === fbLower) return true;
      return normalizePropertyLookupKey(pName) === fbNorm;
    }) || null;
  }

  if (!prop) return fallback;

  var groupId = String(prop.propertyGroupId || prop.PropertyGroupId || prop.property_group_id || (Array.isArray(prop._groupIds) ? prop._groupIds[0] : '') || fallback.groupId || '').trim();
  var groupName = String(prop.propertyGroup || prop.groupName || prop.property_group || prop.portfolio || prop._propertyGroup || '').trim();
  var siteManager = String(prop.siteManager || prop.site_manager || prop.property_manager || prop.PropertyManager || '').trim();

  return {
    id: String(prop.id || prop.Id || prop.property_id || prop.PropertyId || key || '').trim(),
    name: String(prop.name || prop.Name || prop.property_name || prop.PropertyName || fallback.name || '').trim(),
    groupId: groupId,
    groupName: groupName,
    siteManager: siteManager,
  };
}

function enforceScopedPropertyGroup() {
  if (!forcedPropertyGroupUuid) return;
  var scopedName = resolveGroupNameFromUuid(forcedPropertyGroupUuid) || forcedPropertyGroupName;
  if (!scopedName) return;
  forcedPropertyGroupName = scopedName;
  currentPropertyGroup = scopedName;
  var sel = document.getElementById('globalGroupFilter');
  if (sel) {
    for (var i = 0; i < sel.options.length; i++) {
      if (sel.options[i].value === scopedName) {
        sel.value = scopedName;
        break;
      }
    }
    sel.disabled = true;
  }
  var clearBtn = document.getElementById('globalGroupClear');
  if (clearBtn) clearBtn.style.display = 'none';
}

function normalizeGroupSelectionValue(value) {
  var raw = String(value || '').trim();
  if (!raw) return '';
  var lower = raw.toLowerCase();
  if (lower === 'all properties') return '';
  // Treat AppFolio pseudo-global group entries as unscoped.
  if (lower.indexOf('all properties') !== -1 && (raw.charAt(0) === '*' || lower.indexOf('appfolio') !== -1)) return '';
  return raw;
}

// Returns the active group filter name, honouring PM scope enforcement.
// Use instead of reading currentPropertyGroup directly in scope-sensitive code.
function getEffectiveGroupId() {
  function _normalizeEffectiveGroup(value) {
    return normalizeGroupSelectionValue(value);
  }
  if (_accessRole === 'pm_readonly' && forcedPropertyGroupUuid) {
    return _normalizeEffectiveGroup(resolveGroupNameFromUuid(forcedPropertyGroupUuid) || forcedPropertyGroupName || currentPropertyGroup || '');
  }
  return _normalizeEffectiveGroup(currentPropertyGroup || '');
}

function getEffectiveGroupUuid(groupName) {
  var normalizedName = normalizeGroupSelectionValue(groupName || getEffectiveGroupId());
  if (!normalizedName) return '';
  if (_accessRole === 'pm_readonly' && forcedPropertyGroupUuid) {
    return String(forcedPropertyGroupUuid).trim();
  }
  return resolveGroupUuidFromName(normalizedName);
}

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
  ],
  WO_AGE_COLOR_DAYS: {
    yellow: 14,
    orange: 30,
    red: 60
  }
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

// Aggregate turn completion: check ALL Unit Turn WOs for a unit
function isClosedTurnWorkOrderStatus(status) {
  var normalized = String(status || '').trim().toLowerCase();
  return normalized === 'completed' ||
         normalized === 'work completed' ||
         normalized === 'work done' ||
         normalized === 'ready to bill' ||
         normalized === 'closed' ||
         normalized === 'resolved' ||
         normalized === 'canceled' ||
         normalized === 'cancelled';
}

function isTurnWorkDoneStatus(status) {
  var normalized = String(status || '').trim().toLowerCase();
  return normalized === 'work done' || normalized === 'ready to bill' || normalized === 'completed' || isClosedTurnWorkOrderStatus(normalized);
}

function isTurnWorkActiveStatus(status) {
  var normalized = String(status || '').trim().toLowerCase();
  if (!normalized) return false;
  if (isClosedTurnWorkOrderStatus(normalized)) return false;
  return normalized === 'new' ||
         normalized === 'open' ||
         normalized === 'assigned' ||
         normalized === 'scheduled' ||
         normalized === 'estimate requested' ||
         normalized === 'estimated' ||
         normalized === 'in progress' ||
         normalized === 'waiting parts' ||
         normalized === 'awaiting approval';
}

function isTerminalTurnStatus(status) {
  var normalized = String(status || '').trim().toLowerCase();
  return normalized === 'complete' || normalized === 'completed' ||
    normalized === 'closed' || normalized === 'canceled' || normalized === 'cancelled';
}

function isTurnActionLocked(turnKey) {
  var key = String(turnKey || '').trim();
  if (!key) return false;
  var pipelineTurn = (TURN_PIPE_DATA || []).find(function(p) { return String(p.id || '') === key; });
  if (!pipelineTurn) return false;
  if (pipelineTurn.isClosed || pipelineTurn.isCompleted) return true;
  return isTerminalTurnStatus(pipelineTurn.unitTurnStatus || (pipelineTurn.turn && pipelineTurn.turn.status) || '');
}

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

function resetInMemoryDataForSessionTransition() {
  WORK_ORDERS = []; VENDORS = []; PROPERTIES = []; PROPERTY_GROUPS = []; TURNS = []; INSPECTIONS = []; RECENT_TASKS = []; WEBHOOK_EVENTS = []; TURN_RECORDS = []; TURN_PIPE_DATA = []; UNIT_TURNS_DB = []; API_ERRORS = [];
  CLOSED_TURNS = new Set();
  window._currentBillsCache = [];
  window._billingPageRows = [];
  window._billingListCacheRows = [];
  _billingListCacheRows = [];
  _nameToGroups = {}; _idToGroups = {}; _uuidToGroups = {};
  detailCacheClear();
}

async function enforceSessionTypeTransitionReset(previousRole, nextRole) {
  var prev = normalizeAccessRole(previousRole || 'full');
  var next = normalizeAccessRole(nextRole || 'full');
  if (prev === next) return false;

  // Scope/session artifacts must not bleed between role types.
  resetInMemoryDataForSessionTransition();
  currentPropertyGroup = '';
  forcedPropertyGroupUuid = '';
  forcedPropertyGroupName = '';
  _pmScopeGroupUuid = '';
  _pmScopeEmail = '';
  try { localStorage.removeItem('hm_scope_group_uuid'); } catch (e1) { /* */ }
  try { localStorage.removeItem('hm_scope_email'); } catch (e2) { /* */ }
  await clearSessionScopedApiCache();
  updateCacheBadge('offline');
  showToast('Access session changed. Cached data cleared for a clean reload.', {
    kind: 'info',
    iconClass: 'fa-arrows-rotate',
    duration: 4000,
  });
  return true;
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
  if ($('#vhostPreviewPm')) $('#vhostPreviewPm').textContent = val || 'yourco';
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
if ($('#vaultOtpEmail')) {
  $('#vaultOtpEmail').addEventListener('keydown', function(e) {
    if (e.key === 'Enter') { $('#btnSendOtp').click(); }
  });
}
if ($('#vaultOtpCode')) {
  $('#vaultOtpCode').addEventListener('keydown', function(e) {
    if (e.key === 'Enter') { $('#btnVerifyOtp').click(); }
  });
}

// Advanced panel toggle
if ($('#advancedToggle')) {
  $('#advancedToggle').addEventListener('click', function() {
    this.classList.toggle('open');
    $('#advancedPanel').classList.toggle('show');
  });
}

function setVaultPanel(panel) {
  var main = $('#vaultMainPanel');
  var pm = $('#vaultPmPanel');
  if (!main || !pm) return;
  var showPm = panel === 'pm';
  main.classList.toggle('hidden', showPm);
  pm.classList.toggle('hidden', !showPm);
  pm.setAttribute('aria-hidden', showPm ? 'false' : 'true');
}

if ($('#pmLoginToggle')) {
  $('#pmLoginToggle').addEventListener('click', function() {
    setVaultPanel('pm');
    setPmOtpStep('request');
    if ($('#vaultOtpEmail')) $('#vaultOtpEmail').focus();
    setVaultFeedback('Enter your PM email or phone, then click Send OTP.', 'info');
  });
}

if ($('#pmLoginBackBtn')) {
  $('#pmLoginBackBtn').addEventListener('click', function() {
    setVaultPanel('main');
    setPmOtpStep('request');
    if ($('#vaultPassphrase')) $('#vaultPassphrase').focus();
    setVaultFeedback('');
  });
}

if ($('#btnEditOtpIdentifier')) {
  $('#btnEditOtpIdentifier').addEventListener('click', function() {
    setPmOtpStep('request');
    setVaultFeedback('Update the email or phone number, then send a new OTP.', 'info');
    if ($('#vaultOtpEmail')) $('#vaultOtpEmail').focus();
  });
}

// Proxy preset buttons
$$('.vault-proxy-preset').forEach(function(btn) {
  btn.addEventListener('click', function() {
    var presetProxy = btn.getAttribute('data-proxy');
    if (!presetProxy) return;
    $$('.vault-proxy-preset').forEach(function(b) { b.classList.remove('active'); });
    btn.classList.add('active');
    $('#vaultProxy').value = presetProxy;
  });
});

if ($('#btnSendOtp')) {
  $('#btnSendOtp').addEventListener('click', async function() {
    var identifierRaw = $('#vaultOtpEmail') ? $('#vaultOtpEmail').value : '';
    var identifier = normalizeOtpIdentifier(identifierRaw);
    if (!identifier) {
      setVaultFeedback('Enter a valid PM email or phone number to start PM login. US numbers work with or without +1.', '');
      return;
    }
    API_PROXY = sanitizeProxy($('#vaultProxy').value || '');
    if (API_PROXY) {
      try { localStorage.setItem('hm_proxy_url', API_PROXY); } catch (eSaveProxy) { /* */ }
    }
    if (!API_PROXY) {
      setVaultFeedback('Proxy URL is required before requesting OTP.', '');
      return;
    }
    var btn = this;
    btn.disabled = true;
    btn.textContent = 'Sending...';
    setVaultFeedback('');
    try {
      await requestDeviceOtp(identifier, 'dispatcher');
      showToast('OTP sent', { kind: 'success' });
      setPmOtpStep('verify', identifier);
      setVaultFeedback('OTP code sent successfully. Enter the 6-digit verification code.', 'success');
    } catch (err) {
      var msg = (err && (err.message || String(err))) || 'OTP request failed';
      setVaultFeedback(msg, '');
      showToast(msg, { kind: 'danger' });
    } finally {
      btn.disabled = false;
      btn.textContent = 'Send OTP';
    }
  });
}

if ($('#btnVerifyOtp')) {
  $('#btnVerifyOtp').addEventListener('click', async function() {
    var identifier = normalizeOtpIdentifier($('#vaultOtpEmail') ? $('#vaultOtpEmail').value : '');
    var code = String(($('#vaultOtpCode') && $('#vaultOtpCode').value) || '').trim();
    if (!identifier) {
      setPmOtpStep('request');
      setVaultFeedback('Enter your PM email or phone number first.', '');
      return;
    }
    if (!/^\d{6}$/.test(code)) {
      setVaultFeedback('Enter the 6-digit OTP code.', '');
      return;
    }
    API_PROXY = sanitizeProxy($('#vaultProxy').value || '');
    if (!API_PROXY) {
      setVaultFeedback('Proxy URL is required before verifying OTP.', '');
      return;
    }
    var rawVhost = $('#vaultVhost').value;
    var vhost = sanitizeVhost(rawVhost);
    $('#vaultVhost').value = vhost;
    $('#vhostPreview').textContent = vhost || 'yourco';
    $('#vhostPreviewPm').textContent = vhost || 'yourco';
    if (!vhost) {
      setVaultFeedback('AppFolio subdomain is required (e.g. "flraz").', '');
      return;
    }
    var btn = this;
    btn.disabled = true;
    btn.textContent = 'Verifying...';
    setVaultFeedback('');
    try {
      var verifyData = await verifyDeviceOtp(identifier, code, 'dispatcher');
      var token = verifyData.token;
      try { localStorage.setItem('hm_auth_token', token); } catch (eA) { /* */ }
      try { localStorage.setItem('hm_device_token', token); } catch (e) { /* */ }
      try { localStorage.setItem('hm_proxy_token', token); } catch (e2) { /* */ }
      if (verifyData.role) {
        _accessRole = normalizeAccessRole(verifyData.role);
        persistAccessRole(_accessRole);
      }
      if (verifyData.property_group_uuid) {
        forcedPropertyGroupUuid = String(verifyData.property_group_uuid);
        forcedPropertyGroupName = '';
        try { localStorage.setItem('hm_scope_group_uuid', forcedPropertyGroupUuid); } catch (e3) { /* */ }
      }
      if (verifyData.email) {
        _pmScopeEmail = String(verifyData.email || '');
        try { localStorage.setItem('hm_scope_email', _pmScopeEmail); } catch (e4) { /* */ }
      }
      if ($('#vaultPassphrase')) $('#vaultPassphrase').value = '';
      showToast('Device verified', { kind: 'success' });
      setVaultFeedback('OTP verified. Signing you in...', 'success');
      btn.textContent = 'Signing in...';
      try { localStorage.setItem('hm_proxy_url', API_PROXY); } catch (eProxySave) { /* */ }
      await unlockWithDeviceToken(token, vhost, API_PROXY);
    } catch (err) {
      var msg = (err && (err.message || String(err))) || 'OTP verification failed';
      setVaultFeedback(msg, '');
      showToast(msg, { kind: 'danger' });
    } finally {
      btn.disabled = false;
      btn.textContent = 'Verify OTP';
    }
  });
}

async function unlockWithDeviceToken(existingDeviceToken, vhost, proxyUrl) {
  var previousRole = getStoredAccessRole();
  _sessionExpiryHandled = false;
  beginProxySessionStartupGrace(60000);
  API_VHOST = vhost;
  API_PROXY = proxyUrl;
  API_CREDS = { p: existingDeviceToken };
  _accessRole = getStoredAccessRole();
  var stabilized = await stabilizeProxySessionAfterLogin(30000);
  if (!stabilized) {
    // Keep login non-blocking; startup calls can still settle during grace.
    markProxySessionWarmup(90000);
    showToast('Session is still warming up. Data may load gradually for a few moments.', {
      kind: 'warning',
      iconClass: 'fa-hourglass-half',
      duration: 5000,
    });
  }
  try {
    var sess = await proxyAction('session_info');
    if (sess && sess.ok && sess.session) {
      _accessRole = normalizeAccessRole(sess.session.role || _accessRole);
      if (sess.session.property_group_uuid) {
        forcedPropertyGroupUuid = String(sess.session.property_group_uuid);
        forcedPropertyGroupName = '';
        try { localStorage.setItem('hm_scope_group_uuid', forcedPropertyGroupUuid); } catch (scopeErr) { /* */ }
      }
      if (sess.session.login_email) {
        _pmScopeEmail = String(sess.session.login_email);
        try { localStorage.setItem('hm_scope_email', _pmScopeEmail); } catch (scopeEmailErr) { /* */ }
      }
      markProxySessionHealthy();
    }
  } catch (sessErr) { /* non-fatal — use stored role */ }
  await enforceSessionTypeTransitionReset(previousRole, _accessRole);
  persistAccessRole(_accessRole);
  try { localStorage.setItem('hm_auth_token', existingDeviceToken); } catch (eA) { /* */ }
  try { localStorage.setItem('hm_proxy_token', existingDeviceToken); } catch (e) { /* */ }
  if (!$('#vaultRememberConfig') || $('#vaultRememberConfig').checked) {
    await saveVaultConfig(getVaultConfigFromInputs());
  }
  $('#vaultPassphrase').value = '';
  $('#vaultScreen').style.display = 'none';
  $('#appShell').classList.add('unlocked');
  applyAccessRole();
  await initApp();
  maybeAutoRunSystemHealthCheck();
  if (_accessRole === 'pm_readonly') {
    try { forcedPropertyGroupUuid = localStorage.getItem('hm_scope_group_uuid') || forcedPropertyGroupUuid; } catch (eScope) { /* */ }
    enforceScopedPropertyGroup();
  }
  applyAccessRole();
  startAutoSync();
  showToast('Connected — ' + vhost + '.appfolio.com via verified device');
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
setPmOtpStep('request');
if ($('#vhostPreviewPm')) {
  var vVal = sanitizeVhost((($('#vaultVhost') && $('#vaultVhost').value) || ''));
  $('#vhostPreviewPm').textContent = vVal || 'yourco';
}
setVaultPanel('main');

// ── Session hydration on page load ─────────────────────────────────────────
// If the browser has a stored token + saved proxy config, silently re-validate
// the session against the proxy and bypass the vault screen entirely.
(async function _tryResumeSession() {
  var _token = '';
  try {
    _token = localStorage.getItem('hm_auth_token') ||
             localStorage.getItem('hm_device_token') ||
             localStorage.getItem('hm_proxy_token') || '';
  } catch (e) { return; }
  if (!_token) return;

  var _proxyUrl = '';
  try { _proxyUrl = sanitizeProxy(localStorage.getItem('hm_proxy_url') || ''); } catch (e) { /* */ }
  var _vhost = '';
  try {
    var _resumeCfg = await loadVaultConfig();
    if (_resumeCfg) {
      if (!_proxyUrl && _resumeCfg.proxy) _proxyUrl = sanitizeProxy(_resumeCfg.proxy);
      if (_resumeCfg.vhost) _vhost = sanitizeVhost(_resumeCfg.vhost);
    }
  } catch (e) { /* */ }
  if (!_proxyUrl) return;

  setVaultFeedback('Resuming session\u2026', 'info');
  try {
    var _sep = _proxyUrl.indexOf('?') !== -1 ? '&' : '?';
    var _siUrl = _proxyUrl + _sep + 'action=session_info';
    var _siRes = await fetchWithTimeout(_siUrl, {
      headers: { 'Accept': 'application/json', 'Authorization': 'Bearer ' + _token }
    }, 15000);
    if (!_siRes.ok) throw new Error('HTTP ' + _siRes.status);
    var _siData = await _siRes.json();
    if (!_siData || !_siData.ok || !_siData.authenticated || !_siData.session) {
      throw new Error('Session invalid');
    }
    // Valid — restore state and unlock vault
    var _prevRole = getStoredAccessRole();
    _sessionExpiryHandled = false;
    beginProxySessionStartupGrace(60000);
    API_PROXY = _proxyUrl;
    API_VHOST = _vhost;
    API_CREDS = { p: _token };
    _accessRole = normalizeAccessRole(_siData.session.role || _prevRole);
    if (_siData.session.property_group_uuid) {
      forcedPropertyGroupUuid = String(_siData.session.property_group_uuid);
      try { localStorage.setItem('hm_scope_group_uuid', forcedPropertyGroupUuid); } catch (e) { /* */ }
    }
    if (_siData.session.login_email) {
      _pmScopeEmail = String(_siData.session.login_email);
      try { localStorage.setItem('hm_scope_email', _pmScopeEmail); } catch (e) { /* */ }
    }
    markProxySessionHealthy();
    persistAccessRole(_accessRole);
    await enforceSessionTypeTransitionReset(_prevRole, _accessRole);
    setVaultFeedback('', '');
    $('#vaultScreen').style.display = 'none';
    $('#appShell').classList.add('unlocked');
    applyAccessRole();
    await initApp();
    maybeAutoRunSystemHealthCheck();
    if (_accessRole === 'pm_readonly') {
      try { forcedPropertyGroupUuid = localStorage.getItem('hm_scope_group_uuid') || forcedPropertyGroupUuid; } catch (e) { /* */ }
      enforceScopedPropertyGroup();
    }
    applyAccessRole();
    startAutoSync();
  } catch (err) {
    // Token expired or invalid — clean up and show vault
    API_CREDS = null;
    API_PROXY = '';
    clearStoredProxySessionTokens();
    setVaultFeedback('', '');
  }
})();

$('#vaultUnlockBtn').addEventListener('click', async function() {
  var previousRole = getStoredAccessRole();
  var pass = String(($('#vaultPassphrase') && $('#vaultPassphrase').value) || '').trim();
  var rawVhost = $('#vaultVhost').value;
  var vhost = sanitizeVhost(rawVhost);
  $('#vaultVhost').value = vhost;
  $('#vhostPreview').textContent = vhost || 'yourco';
  var proxyUrl = sanitizeProxy($('#vaultProxy').value);
  if (!proxyUrl) {
    try {
      proxyUrl = sanitizeProxy(localStorage.getItem('hm_proxy_url') || '');
      if (proxyUrl && $('#vaultProxy')) $('#vaultProxy').value = proxyUrl;
    } catch (eProxyLoad) { /* */ }
  }
  $('#vaultProxy').value = proxyUrl;
  if (!vhost) {
    $('#vaultError').textContent = 'AppFolio subdomain is required (e.g. "flraz").';
    $('#vaultError').classList.add('show');
    return;
  }
  if (!proxyUrl) {
    $('#vaultError').textContent = 'Proxy URL is required \u2014 enter your Val.town proxy URL.';
    $('#vaultError').classList.add('show');
    return;
  }
  if (!pass) {
    $('#vaultError').textContent = 'Enter your password. For PM access, use the separate PM login field and buttons below.';
    $('#vaultError').classList.add('show');
    return;
  }

  var btn = this;
  btn.disabled = true;
  btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Connecting\u2026';
  $('#vaultError').classList.remove('show');

  try {
    API_VHOST = vhost;
    API_PROXY = proxyUrl;
    try { localStorage.setItem('hm_proxy_url', API_PROXY); } catch (eProxy) { /* */ }

    var passAsIdentifier = normalizeOtpIdentifier(pass);

    if (passAsIdentifier) {
      throw new Error('PM login uses the separate PM login field and buttons below. Do not enter email or phone in the password box.');
    } else if (pass) {
      clearStoredProxySessionTokens();
      _sessionExpiryHandled = false;
      _proxySessionProbeInFlight = false;
      _proxySessionWarmupFailures = 0;
      _proxySessionConsecutive401 = 0;
      // Verify password against proxy env vars (GUI_ADMIN / GUI_GM / GUI_VENDORS)
      var roleResult = await proxyPost('verify_role', { password: pass });
      if (!roleResult || !roleResult.ok) {
        throw new Error((roleResult && roleResult.error) || 'Invalid password');
      }
      _accessRole = normalizeAccessRole(roleResult.role);
      var sessionToken = String(roleResult.token || '');
      if (!sessionToken) {
        throw new Error('Login succeeded but server did not mint a session token. Check proxy DB/write access and try again.');
      }
      _sessionExpiryHandled = false;
      beginProxySessionStartupGrace(60000);
      API_CREDS = { p: sessionToken };
      if (sessionToken) {
        try { localStorage.setItem('hm_auth_token', sessionToken); } catch (eTokA) { /* */ }
        try { localStorage.setItem('hm_device_token', sessionToken); } catch (eTok) { /* */ }
        try { localStorage.setItem('hm_proxy_token', sessionToken); } catch (eTok2) { /* */ }
      }
      var stabilized = await stabilizeProxySessionAfterLogin(30000);
      if (!stabilized) {
        markProxySessionWarmup(90000);
        showToast('Session is still warming up. Data may load gradually for a few moments.', {
          kind: 'warning',
          iconClass: 'fa-hourglass-half',
          duration: 5000,
        });
      }
      await enforceSessionTypeTransitionReset(previousRole, _accessRole);
      persistAccessRole(_accessRole);
    } else {
      throw new Error('Enter your password to connect.');
    }

    if (!$('#vaultRememberConfig') || $('#vaultRememberConfig').checked) {
      await saveVaultConfig(getVaultConfigFromInputs());
    }
    $('#vaultPassphrase').value = '';
    $('#vaultScreen').style.display = 'none';
    $('#appShell').classList.add('unlocked');
    applyAccessRole();
    await initApp();
    maybeAutoRunSystemHealthCheck();
    if (_accessRole === 'pm_readonly') {
      try { forcedPropertyGroupUuid = localStorage.getItem('hm_scope_group_uuid') || forcedPropertyGroupUuid; } catch (eScope) { /* */ }
      enforceScopedPropertyGroup();
    }
    applyAccessRole();
    startAutoSync();
    if (_accessRole === 'vendors') {
      showToast('Vendor access \u2014 connecting to ' + vhost + '.appfolio.com via proxy');
    } else {
      showToast('Connected \u2014 ' + vhost + '.appfolio.com via proxy');
    }
  } catch (err) {
    var errMsg = (err && (err.message || String(err))) || '';
    var schemaErr = /no such table|no such column|SQLITE_UNKNOWN|SQL_INPUT_ERROR/i.test(errMsg);
    wipeCredentials();
    $('#vaultError').textContent = schemaErr
      ? 'Proxy connected, but database schema is out of date. Deploy latest proxy migrations and retry.'
      : (errMsg || 'Login failed \u2014 check password and proxy URL.');
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

// Vendor detail modal — close buttons and cross-tab routing
if ($('#vendorDetailModalClose'))    $('#vendorDetailModalClose').addEventListener('click', function() { closeModal('vendorDetailModal'); });
if ($('#vendorDetailModalCloseBtn')) $('#vendorDetailModalCloseBtn').addEventListener('click', function() { closeModal('vendorDetailModal'); });
if ($('#vdmBtnOpenWOs')) {
  $('#vdmBtnOpenWOs').addEventListener('click', function() {
    var v = VENDORS.find(function(vn) { return String(vn.id) === _currentSelectedVendorId; });
    closeModal('vendorDetailModal');
    if (v) navigateToOpenWOsForVendor(v.name);
  });
}
if ($('#vdmBtnCompletedWOs')) {
  $('#vdmBtnCompletedWOs').addEventListener('click', function() {
    var v = VENDORS.find(function(vn) { return String(vn.id) === _currentSelectedVendorId; });
    closeModal('vendorDetailModal');
    if (v) navigateToCompletedWOsForVendor(v.name);
  });
}
if ($('#vdmBtnBills')) {
  $('#vdmBtnBills').addEventListener('click', function() {
    var v = VENDORS.find(function(vn) { return String(vn.id) === _currentSelectedVendorId; });
    closeModal('vendorDetailModal');
    if (v) navigateToBillsForVendor(v.id || v.name);
  });
}

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
        unitTurnId: r.unit_turn_id || r.UnitTurnId || '',
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

async function fetchCompletedWorkOrdersHistory(days) {
  var lookback = parseInt(days || DEFAULT_COMPLETED_WO_LOOKBACK_DAYS, 10) || DEFAULT_COMPLETED_WO_LOOKBACK_DAYS;
  lookback = Math.max(30, Math.min(3650, lookback));
  var data = await proxyAction('work_orders_completed_history', { days: String(lookback) });
  var results = data.results || [];
  completedWOHistoryRows = results.map(function(r) {
    return {
      id: r.work_order_number || r.WorkOrderNumber || r.service_request_number || '',
      uuid: r.work_order_id || r.Id || '',
      propertyId: r.property_id || r.PropertyId || '',
      propertyName: r.property_name || r.property || r.PropertyName || '',
      unit: r.unit_name || r.UnitName || r.unit_id || '',
      vendorName: r.vendor || r.VendorName || '',
      vendorId: r.vendor_id || r.VendorId || '',
      status: r.status || r.Status || '',
      completedOn: r.completed_on || r.CompletedOn || r.work_completed_on || r.WorkCompletedOn || r.status_date || r.StatusDate || '',
      amountBilled: r.amount_billed || r.AmountBilled || r.amount || r.Amount || r.total || r.Total || '',
      priority: r.priority || r.Priority || 'Normal',
      type: r.work_order_type || r.Type || '',
      description: r.job_description || r.JobDescription || r.service_request_description || r.Description || '',
      tenant: r.primary_tenant || r.PrimaryTenant || '',
      completedBy: r.completed_by || r.CompletedBy || '',
      paidBy: r.paid_by || r.PaidBy || ''
    };
  });
  completedWOHistoryPage = 0;
  return true;
}

// Bills: Proxy actions (bills, bills_stats, bills_history, bill_detail)
// Used by WO close-assist and Billing/AP tab.
function resolveBillGroupUuidFromRecord(record) {
  if (!record || typeof record !== 'object') return '';
  var nestedProperty = (record.property && typeof record.property === 'object') ? record.property : {};
  var candidates = [
    record.property_group_id,
    record.property_group_uuid,
    record.PropertyGroupId,
    record.PropertyGroupUuid,
    record.group_id,
    record.group_uuid,
    nestedProperty.property_group_id,
    nestedProperty.property_group_uuid,
    nestedProperty.PropertyGroupId,
    nestedProperty.PropertyGroupUuid,
  ];
  for (var i = 0; i < candidates.length; i++) {
    var v = String(candidates[i] || '').trim();
    if (v) return v;
  }
  return '';
}

// Resolve a human-readable group name for a mapped bill object.
// Checks propertyGroupId UUID -> group name map, then _uuidToGroups / _nameToGroups.
function resolveBillGroupName(b) {
  if (!b) return '';
  var direct = String(b.propertyGroup || '').trim();
  if (direct) return direct;

  var gid = String(b.propertyGroupId || resolveBillGroupUuidFromRecord(b.raw || b) || '').trim();
  if (gid) {
    var fromUuid = resolveGroupNameFromUuid(gid);
    if (fromUuid) return fromUuid;
  }

  var pid = String(b.propertyId || '').trim();
  if (pid) {
    var uGroups = _uuidToGroups[pid];
    if (uGroups && uGroups.length) return uGroups[0];
    var iGroups = _idToGroups[pid];
    if (iGroups && iGroups.length) return iGroups[0];
  }

  var pname = String(b.propertyName || '').trim().toLowerCase();
  if (pname && pname !== 'multiple/unknown') {
    var nGroups = _nameToGroups[pname];
    if (nGroups && nGroups.length) return nGroups[0];
    var normalizedName = normalizePropertyLookupKey(pname);
    if (normalizedName) {
      var nGroupsNormalized = _nameToGroups[normalizedName];
      if (nGroupsNormalized && nGroupsNormalized.length) return nGroupsNormalized[0];
    }
  }

  return '';
}

function resolveBillWorkOrderRef(record) {
  if (!record || typeof record !== 'object') return '';
  var raw = (record.raw && typeof record.raw === 'object') ? record.raw : record;
  var candidates = [
    record.workOrderNumber,
    record.work_order_number,
    raw.WorkOrderNumber,
    raw.work_order_number,
    raw.WorkOrder,
    raw.work_order,
    record.workOrderId,
    record.work_order_id,
    raw.WorkOrderId,
    raw.work_order_id,
    raw.work_order_uuid,
    raw.WorkOrderUuid,
  ];
  for (var i = 0; i < candidates.length; i++) {
    var v = String(candidates[i] || '').trim();
    if (v) return v;
  }
  return '';
}

function resolveLineItemUuidTarget(lineItem) {
  if (!lineItem || typeof lineItem !== 'object') return null;
  var pick = function(list) {
    for (var i = 0; i < list.length; i++) {
      var val = String(lineItem[list[i]] || '').trim();
      if (isUuidString(val)) return val;
    }
    return '';
  };

  var woUuid = pick(['WorkOrderId', 'work_order_id', 'work_order_uuid', 'WorkOrderUuid']);
  if (woUuid) return { resource: 'work_orders', label: 'Work Order', uuid: woUuid };

  var propertyUuid = pick(['PropertyId', 'property_id', 'property_uuid', 'PropertyUuid']);
  if (propertyUuid) return { resource: 'properties', label: 'Property', uuid: propertyUuid };

  var unitUuid = pick(['UnitId', 'unit_id', 'unit_uuid', 'UnitUUID', 'UnitUuid']);
  if (unitUuid) return { resource: 'units', label: 'Unit', uuid: unitUuid };

  var vendorUuid = pick(['VendorId', 'vendor_id', 'vendor_uuid', 'VendorUuid', 'PayeeUuid', 'payee_uuid']);
  if (vendorUuid) return { resource: 'vendors', label: 'Vendor', uuid: vendorUuid };

  return null;
}

async function showV0UuidDetailModal(target, lineItem) {
  if (!target || !target.resource || !target.uuid) return;
  var title = 'v0 UUID Detail — ' + target.label;
  var path = '/api/v0/' + target.resource + '/' + encodeURIComponent(String(target.uuid));
  showItemDetail(title, [
    { label: 'Resource', value: target.resource },
    { label: 'UUID', value: target.uuid },
    { label: 'Path', value: path },
    { section: 'Loading', icon: 'fa-spinner' },
    { label: 'Status', value: 'Fetching details…' }
  ], '');

  try {
    var payload = await apiFetch(path);
    var bodyEl = document.getElementById('itemDetailBody');
    if (!bodyEl) return;
    var pretty = '';
    try { pretty = JSON.stringify(payload || {}, null, 2); } catch (e) { pretty = String(payload || ''); }
    bodyEl.innerHTML =
      '<div class="detail-section">' +
        '<div class="detail-section-title"><i class="fas fa-code"></i> API v0 Response</div>' +
        '<div class="detail-grid">' +
          '<div class="detail-row"><div class="detail-row-label">Resource</div><div class="detail-row-value">' + escapeHtml(target.resource) + '</div></div>' +
          '<div class="detail-row"><div class="detail-row-label">UUID</div><div class="detail-row-value" style="font-family:var(--font-mono)">' + escapeHtml(target.uuid) + '</div></div>' +
          '<div class="detail-row"><div class="detail-row-label">Path</div><div class="detail-row-value" style="font-family:var(--font-mono)">' + escapeHtml(path) + '</div></div>' +
        '</div>' +
      '</div>' +
      '<pre style="max-height:360px;overflow:auto;background:var(--bg-secondary);border:1px solid var(--border);border-radius:8px;padding:10px;font-size:11px;line-height:1.45;font-family:var(--font-mono)">' + escapeHtml(pretty || '{}') + '</pre>';
  } catch (err) {
    showItemDetail(title, [
      { label: 'Resource', value: target.resource },
      { label: 'UUID', value: target.uuid },
      { label: 'Path', value: path },
      { section: 'Error', icon: 'fa-triangle-exclamation' },
      { label: 'Message', value: String((err && err.message) || err || 'Failed to fetch v0 details') }
    ], '');
  }
}

async function fetchBills(days, opts) {
  if (_billsFetchInFlight) {
    console.log('fetchBills wait: previous request still in flight');
    var waitMs = 0;
    while (_billsFetchInFlight && waitMs < 8000) {
      await new Promise(function(resolve) { setTimeout(resolve, 100); });
      waitMs += 100;
    }
    if (_billsFetchInFlight) {
      console.log('fetchBills busy timeout after ' + waitMs + 'ms');
      if (opts && opts.returnPayload) {
        return {
          ok: false,
          error: 'billing fetch busy, try again',
          rows: [],
          total: 0,
          page: 1,
          perPage: Math.max(1, Math.min(200, parseInt((opts.perPage || opts.limit || 50), 10) || 50)),
          totalPages: 1,
          fromCache: false,
          source: String(_lastBillSource || 'legacy')
        };
      }
      return false;
    }
  }
  _billsFetchInFlight = true;
  try {
    opts = opts || {};
    var lookback = parseInt(days || DEFAULT_BILLS_LOOKBACK_DAYS, 10) || DEFAULT_BILLS_LOOKBACK_DAYS;
    var params = {
      days: String(lookback),
      max: String(opts.max || 3000),
      columns: BILL_ROUTE_COLUMNS
    };
    var requestedPage = Math.max(1, parseInt(opts.page || 1, 10) || 1);
    var requestedPerPage = Math.max(1, Math.min(200, parseInt(opts.perPage || opts.limit || 50, 10) || 50));
    var assignGlobal = opts.assignGlobal !== false;
    if (opts.forceRefresh) params.force_refresh = 'true';
    var routeAction = String(opts.filterType || '').trim();
    var routeStatusFilter = String(opts.routeStatusFilter || opts.statusFilter || '').trim().toLowerCase();
    var filterValue = String(opts.filterValue || '').trim();
    var dueFrom = String(opts.dueFrom || '').trim();
    var dueTo = String(opts.dueTo || '').trim();
    var routeLimit = requestedPerPage;
    var routeOffset = Math.max(0, parseInt((opts.offset != null ? opts.offset : ((requestedPage - 1) * routeLimit)), 10) || 0);
    var totalRows = 0;
    var totalPages = 1;
    var currentPage = requestedPage;

    var allowedRouteActions = {
      bills_list: true,
      bills_by_vendor: true,
      bills_by_property: true,
      bills_by_wo: true,
      bills_by_wo_number: true,
      bills_by_invoice: true,
      bills_due_range: true
    };
    if (!allowedRouteActions[routeAction]) routeAction = '';

    var usedRouteLayer = false;
    var routeRows = [];
    var grpName = '';
    var groupUuid = '';

    if (opts.scoped !== false) {
      grpName = normalizeGroupSelectionValue(getEffectiveGroupId());
      groupUuid = getEffectiveGroupUuid(grpName);
      // Only send UUID scope to server-side filters; when UUID is missing,
      // rely on client-side property-group association maps to avoid empty legacy responses.
      if (groupUuid) params.group_uuid = groupUuid;
    }

    if (!routeAction) {
      params.page = String(requestedPage);
      params.per_page = String(requestedPerPage);
    }

    if (routeAction) {
      try {
        // bills_by_wo accepts only UUID; route non-UUID values to wo_number action.
        if (routeAction === 'bills_by_wo' && filterValue && !isUuidString(filterValue)) {
          routeAction = 'bills_by_wo_number';
        }

        var routeParams = {
          limit: String(routeLimit),
          offset: String(routeOffset),
          columns: BILL_ROUTE_COLUMNS
        };

        if (routeAction === 'bills_list') {
          if (!groupUuid) throw new Error('Missing group UUID for bills_list');
          routeParams.group_id = groupUuid;
        } else if (routeAction === 'bills_by_vendor') {
          routeParams.vendor_id = filterValue;
          if (groupUuid) routeParams.group_id = groupUuid;
        } else if (routeAction === 'bills_by_property') {
          routeParams.property_id = filterValue;
          if (groupUuid) routeParams.group_id = groupUuid;
        } else if (routeAction === 'bills_by_wo') {
          routeParams.wo_id = filterValue;
        } else if (routeAction === 'bills_by_wo_number') {
          routeParams.wo_number = filterValue;
        } else if (routeAction === 'bills_by_invoice') {
          routeParams.invoice_number = filterValue;
        } else if (routeAction === 'bills_due_range') {
          routeParams.due_from = dueFrom;
          routeParams.due_to = dueTo;
          if (groupUuid) routeParams.group_id = groupUuid;
        }

        Object.keys(routeParams).forEach(function(k) {
          if (routeParams[k] === '' || routeParams[k] == null) delete routeParams[k];
        });

        var routeDataByFilter = await proxyAction(routeAction, routeParams);
        routeRows = routeDataByFilter && (routeDataByFilter.data || routeDataByFilter.results)
          ? (routeDataByFilter.data || routeDataByFilter.results)
          : [];
        totalRows = Number(routeDataByFilter && routeDataByFilter.total || routeRows.length) || routeRows.length;
        totalPages = Math.max(1, Math.ceil(totalRows / routeLimit));
        currentPage = Math.max(1, Math.floor(routeOffset / routeLimit) + 1);
        usedRouteLayer = true;
      } catch (routeFilterErr) {
        console.log('fetchBills explicit route fallback: ' + (routeFilterErr.message || routeFilterErr));
      }
    }

    if (!usedRouteLayer && opts.scoped !== false && groupUuid) {
      // Prefer the new SQL route layer when we can provide a concrete group UUID.
      try {
        var routeData = await proxyAction('bills_list', {
          group_id: groupUuid,
          limit: String(routeLimit),
          offset: String(routeOffset),
          columns: BILL_ROUTE_COLUMNS
        });
        routeRows = routeData && (routeData.data || routeData.results)
          ? (routeData.data || routeData.results)
          : [];
        totalRows = Number(routeData && routeData.total || routeRows.length) || routeRows.length;
        totalPages = Math.max(1, Math.ceil(totalRows / routeLimit));
        currentPage = Math.max(1, Math.floor(routeOffset / routeLimit) + 1);
        usedRouteLayer = true;
      } catch (routeErrOuter) {
        console.log('fetchBills route-layer fallback: ' + (routeErrOuter.message || routeErrOuter));
      }
    }

    var results = [];
    if (usedRouteLayer) {
      results = Array.isArray(routeRows) ? routeRows : [];
      _lastBillSource = 'cached';
    } else {
      var data = await proxyAction('bills', params);
      results = data.results || data.data || [];
      totalRows = Number(data.total || data.count || results.length) || results.length;
      totalPages = Math.max(1, Number(data.total_pages || Math.ceil(totalRows / requestedPerPage) || 1));
      currentPage = Math.max(1, Number(data.page || requestedPage || 1));
      if (typeof data.from_cache === 'boolean') {
        _lastBillSource = data.from_cache ? 'cached' : 'live';
      } else {
        _lastBillSource = 'legacy';
      }
    }

    var mappedBills = results.map(function(b) {
      var raw = b.raw || b;
      var nestedProperty = (raw.property && typeof raw.property === 'object')
        ? raw.property
        : ((b.property && typeof b.property === 'object') ? b.property : {});
      var status = getBillStatusKey({ status: b.status, raw: raw });
      var amountNum = resolveBillAmountValue({
        raw: raw,
        lineItems: b.line_items || raw.LineItems || raw.line_items || [],
        amount: b.amount,
        total_amount: b.total_amount,
        net_amount: b.net_amount,
        amount_due: b.amount_due,
      });
      var lineItems = Array.isArray(b.line_items)
        ? b.line_items
        : (Array.isArray(raw.LineItems)
          ? raw.LineItems
          : (Array.isArray(raw.line_items)
            ? raw.line_items
            : (function() {
              if (!b.line_items_json && !raw.line_items_json) return [];
              try {
                var parsed = JSON.parse(b.line_items_json || raw.line_items_json || '[]');
                return Array.isArray(parsed) ? parsed : [];
              } catch (e) {
                return [];
              }
            })()));
      var propertyId = extractBillPropertyId(raw, lineItems);
      var unitId = extractBillUnitId(raw, lineItems);
      var propertyMeta = resolvePropertyMetaFromMaps(
        propertyId,
        b.property_name || raw.PropertyName || raw.property_name || raw.Property || raw.property || '',
        b.property_group_id || b.property_group_uuid || raw.property_group_id || raw.property_group_uuid || raw.PropertyGroupId || raw.PropertyGroupUuid || nestedProperty.property_group_id || nestedProperty.property_group_uuid || nestedProperty.PropertyGroupId || nestedProperty.PropertyGroupUuid || ''
      );
      var vendorId = b.vendor_id || raw.VendorId || raw.vendor_id || raw.PayeeId || raw.payee_id || raw.PayeeUuid || raw.payee_uuid || '';
      var vendorName = resolveVendorNameFromMaps(
        vendorId,
        b.vendor_name || raw.VendorName || raw.vendor_name || raw.PayeeName || raw.payee_name || raw.Name || raw.name || ''
      );
      return {
        id: b.id || b.Id || raw.Id || raw.id || raw.BillId || '',
        billNumber: b.bill_number || b.reference || raw.Reference || raw.reference || b.id || raw.Id || '',
        vendorId: vendorId,
        vendorUuid: b.vendor_uuid || raw.VendorUuid || raw.vendor_uuid || raw.VendorId || raw.vendor_id || raw.PayeeUuid || raw.payee_uuid || '',
        payeeUuid: b.payee_uuid || raw.PayeeUuid || raw.payee_uuid || raw.PayeeId || raw.payee_id || raw.VendorId || raw.vendor_id || '',
        vendorName: vendorName,
        propertyId: propertyMeta.id || propertyId,
        unitId: unitId,
        propertyName: propertyMeta.name || 'Multiple/Unknown',
        propertyGroup: propertyMeta.groupName || b.property_group || raw.property_group || raw.PropertyGroup || '',
        propertyGroupId: propertyMeta.groupId,
        propertyManager: propertyMeta.siteManager || b.property_manager || raw.property_manager || raw.PropertyManager || '',
        workOrderId: b.work_order_number || b.workOrderNumber || b.work_order_id || raw.WorkOrderNumber || raw.work_order_number || raw.WorkOrderId || raw.work_order_id || raw.WorkOrder || raw.work_order || raw.work_order_uuid || raw.WorkOrderUuid || '',
        amount: amountNum,
          date: b.invoice_date || raw.InvoiceDate || raw.invoice_date || b.due_date || raw.DueDate || raw.due_date || raw.BillDate || raw.bill_date || raw.PaidOn || raw.paid_on || raw.CreatedAt || raw.created_at || raw.LastUpdatedAt || raw.last_updated_at || '',
          lastUpdatedAt: b.last_updated_at || b.lastUpdatedAt || raw.LastUpdatedAt || raw.last_updated_at || raw.UpdatedAt || raw.updated_at || '',
          status: status,
          statusLabel: b.status_label || raw.ApprovalStatus || raw.approval_status || raw.Status || raw.status || (status ? status.replace(/_/g, ' ') : '—'),
          lineItems: lineItems,
          raw: raw
      };
    });

    var updatedFromFilter = String(opts.updatedFrom || '').slice(0, 10);
    var updatedToFilter = String(opts.updatedTo || '').slice(0, 10);
    if (updatedFromFilter || updatedToFilter) {
      mappedBills = mappedBills.filter(function(b) {
        var raw = b.raw || {};
        var updated = String(b.lastUpdatedAt || raw.LastUpdatedAt || raw.last_updated_at || b.date || '').slice(0, 10);
        if (!updated) return false;
        if (updatedFromFilter && updated < updatedFromFilter) return false;
        if (updatedToFilter && updated > updatedToFilter) return false;
        return true;
      });
    }

    if (routeStatusFilter) {
      mappedBills = mappedBills.filter(function(b) {
        return String(b.status || '').toLowerCase() === routeStatusFilter;
      });
    }

    // Strict client-side scope guard.
    if (opts.scoped !== false && (groupUuid || grpName)) {
      var groupUuidLower = String(groupUuid || '').trim().toLowerCase();
      var groupNameLower = String(grpName || '').trim().toLowerCase();
      var hasGroupMaps = !!(Object.keys(_nameToGroups || {}).length || Object.keys(_uuidToGroups || {}).length || Object.keys(_idToGroups || {}).length);
      mappedBills = mappedBills.filter(function(b) {
        var billGroupUuid = String(b.propertyGroupId || resolveBillGroupUuidFromRecord(b.raw || b) || '').trim().toLowerCase();
        if (groupUuidLower && billGroupUuid) return billGroupUuid === groupUuidLower;
        var billGroupName = String(resolveBillGroupName(b) || b.propertyGroup || b._propertyGroup || '').trim().toLowerCase();
        if (groupNameLower && billGroupName) return billGroupName === groupNameLower;
        if (hasGroupMaps) return isInPropertyGroup(b.propertyId, b.propertyName, grpName);
        // If maps are not ready yet and no explicit group fields are present, avoid false-empty results.
        return true;
      });
    }

    if (assignGlobal) {
      BILLS = mappedBills;
      _billsLoadedAt = Date.now();
    }

    if (opts.returnPayload) {
      return {
        ok: true,
        rows: mappedBills,
        total: totalRows,
        page: currentPage,
        perPage: routeLimit,
        totalPages: totalPages,
        fromCache: _lastBillSource === 'cached',
        source: _lastBillSource
      };
    }

    return true;
  } catch (err) {
    console.log('fetchBills error: ' + (err.message || err));
    _lastBillSource = 'legacy';
    if (assignGlobal) BILLS = [];
    if (opts && opts.returnPayload) {
      return {
        ok: false,
        error: String(err && (err.message || err) || 'fetch bills failed'),
        rows: [],
        total: 0,
        page: 1,
        perPage: requestedPerPage,
        totalPages: 1,
        fromCache: false,
        source: 'legacy'
      };
    }
    return false;
  } finally {
    _billsFetchInFlight = false;
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
        licenseExpires: v.license_expires || v.contractor_license_expires || v.business_license_expires || '',
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
    buildReferenceMaps(VENDORS, PROPERTIES);
    return true;
  } catch (err) {
    VENDORS = [];
    buildReferenceMaps(VENDORS, PROPERTIES);
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
      var siteManager = '';
      if (p.site_manager && typeof p.site_manager === 'object') {
        siteManager = [p.site_manager.FirstName, p.site_manager.LastName].filter(Boolean).join(' ').trim();
      }
      if (!siteManager && p.SiteManager && typeof p.SiteManager === 'object') {
        siteManager = [p.SiteManager.FirstName, p.SiteManager.LastName].filter(Boolean).join(' ').trim();
      }
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
        propertyGroupId: p.property_group_id || p.PropertyGroupId || p.property_group_uuid || p.PropertyGroupUuid || '',
        group: p.group || '',
        groupName: p.group_name || '',
        maintenanceLimit: p.maintenance_limit || p.maintenanceLimit || '',
        maintenanceNotes: p.maintenance_notes || p.maintenanceNotes || '',
        siteManager: siteManager || p.site_manager || p.siteManager || p.SiteManager || '',
        units: p.units || '',
        sqft: p.sqft || '',
        marketRent: p.market_rent || p.marketRent || '',
        owners: p.owners || '',
        link: ''
      };
    });
    buildReferenceMaps(VENDORS, PROPERTIES);
    return true;
  } catch (err) {
    PROPERTIES = [];
    buildReferenceMaps(VENDORS, PROPERTIES);
    return false;
  }
}

function normalizeTurnRecord(t) {
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
}

// Turns: Proxy ?action=turns — merged In Progress + Completed for richer detail context
async function fetchTurns() {
  try {
    setApiStatus('loading', 'Loading turns (In Progress + Completed)…');
    var activeData = await proxyAction('turns', { days: 60, status: 'In Progress' });
    var activeRows = activeData.results || activeData.data || [];
    var completedRows = [];
    try {
      var completedData = await proxyAction('turns', { days: 30, status: 'Completed' });
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

    TURNS = results.map(normalizeTurnRecord);
    DASH_TURN_LAST_SYNC_AT = new Date().toISOString();
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

// Hard floor — company did not use AppFolio before 2021.
// Pre-2021 last_inspection_date values are data-migration artifacts.
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

  // Move-in resets inspection compliance clock.
  // If no inspection occurred on/after move-in, keep it visible as a missing move-in inspection.
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
  if (!anchor) return true; // never inspected = include (overdue)
  if (anchor < INSPECTION_AF_EPOCH) return false; // pre-AppFolio artifact
  var yearStart = getCurrentYearStartDate(now);
  if (anchor < yearStart) return false; // current year only
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
        var smObj = p.SiteManager || p.site_manager || null;
        var smName = '';
        if (smObj && typeof smObj === 'object') {
          smName = [smObj.FirstName, smObj.LastName].filter(Boolean).join(' ').trim();
        } else if (typeof smObj === 'string') {
          smName = smObj.trim();
        }
        var gIds = [];
        var rawGroupIds = p.PropertyGroupIds || p.property_group_ids || [];
        if (Array.isArray(rawGroupIds)) {
          gIds = rawGroupIds.map(function(v) { return String(v || '').trim(); }).filter(Boolean);
        }
        if (pid && pname) uuidMapFallback[String(pid)] = {
          name: String(pname),
          site_manager_name: smName,
          group_ids: gIds,
          maintenance_notes: p.MaintenanceNotes || p.maintenance_notes || ''
        };
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

    function _normalizeGroupIds(entry) {
      if (!entry || typeof entry !== 'object') return [];
      var ids = entry.group_ids || entry.groupIds || [];
      if (!Array.isArray(ids)) return [];
      return ids.map(function(v) { return String(v || '').trim(); }).filter(Boolean);
    }

    function _propertyMetaFromMap(entry) {
      if (!entry || typeof entry !== 'object') return null;
      return {
        siteManager: String(entry.site_manager_name || entry.siteManager || '').trim(),
        maintenanceNotes: String(entry.maintenance_notes || entry.maintenanceNotes || '').trim(),
        groupIds: _normalizeGroupIds(entry)
      };
    }

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
          _addNameToGroupMap(mName, g.name);
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
        // Direct UUID match if Reports API id is already DB UUID-like.
        var directEntry = p.id ? uuidMap[String(p.id)] : null;
        var directMeta = _propertyMetaFromMap(directEntry);
        if (directMeta) {
          p._dbUuid = String(p.id || '');
          if (!String(p.siteManager || '').trim() && directMeta.siteManager) {
            p.siteManager = directMeta.siteManager;
          }
          if (!String(p.maintenanceNotes || '').trim() && directMeta.maintenanceNotes) {
            p.maintenanceNotes = directMeta.maintenanceNotes;
          }
          if ((!p._groupIds || !p._groupIds.length) && directMeta.groupIds.length) {
            p._groupIds = directMeta.groupIds.slice();
          }
        }

        if (!p.name) return;
        var pNameLower = (p.name || '').trim().toLowerCase();
        // Check if this Reports API property name matches any DB API property name
        var matchedUuids = dbNameToUuids[pNameLower];
        if (matchedUuids && matchedUuids.length > 0) {
          diag.nameMatches++;
          // Store the first UUID on the property for reference
          p._dbUuid = matchedUuids[0];

          // Pull richer metadata from DB map to close UI data gaps.
          var mappedEntry = uuidMap[matchedUuids[0]];
          var mappedMeta = _propertyMetaFromMap(mappedEntry);
          if (mappedMeta) {
            if (!String(p.siteManager || '').trim() && mappedMeta.siteManager) {
              p.siteManager = mappedMeta.siteManager;
            }
            if (!String(p.maintenanceNotes || '').trim() && mappedMeta.maintenanceNotes) {
              p.maintenanceNotes = mappedMeta.maintenanceNotes;
            }
            if ((!p._groupIds || !p._groupIds.length) && mappedMeta.groupIds.length) {
              p._groupIds = mappedMeta.groupIds.slice();
            }
          }

          // For each matched UUID, copy that UUID's group memberships to the property_id
          matchedUuids.forEach(function(uuid) {
            var uGroups = _uuidToGroups[uuid];
            if (uGroups) {
              uGroups.forEach(function(gn) {
                _addToGroupMap(_idToGroups, String(p.id), gn);
                // Also ensure the Reports API name is indexed
                _addNameToGroupMap(pNameLower, gn);
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
            _addNameToGroupMap(prop.name || '', g.name);
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
            _addNameToGroupMap(p.name || '', g.name);
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

  // Step 3: Enrich loaded WORK_ORDERS and TURNS with _propertyGroup field.
  // Runs after maps are built so each item carries its group for display and PM scope checks.
  try {
    WORK_ORDERS.forEach(function(wo) {
      var byName = groupsByPropertyName(wo.propertyName || '');
      var byId   = _idToGroups[String(wo.propertyId || '')] || [];
      var grps = byName.length ? byName : byId;
      wo._propertyGroup = grps[0] || '';
    });
    TURNS.forEach(function(t) {
      var byName = groupsByPropertyName(t.property || '');
      var byId   = _idToGroups[String(t.propertyId || '')] || [];
      var grps = byName.length ? byName : byId;
      t._propertyGroup = grps[0] || '';
    });
    // Enrich BILLS with _propertyGroup if already loaded
    (BILLS || []).forEach(function(b) {
      var byName = groupsByPropertyName(b.propertyName || '');
      var byId   = _idToGroups[String(b.propertyId || '')] || [];
      var grps = byName.length ? byName : byId;
      b._propertyGroup = grps[0] || '';
    });
    console.log('[PG] Step 3 done — enriched ' + WORK_ORDERS.length + ' WOs, ' + TURNS.length + ' turns with _propertyGroup');
  } catch (enrErr) {
    console.log('[PG] Enrichment error: ' + (enrErr.message || enrErr));
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
        unitTurnId: wo.UnitTurnId || wo.unit_turn_id || '',
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
    var data = await proxyAction('unit_turns', { days: '90' });
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
        if (!turn.linkedWorkOrders || turn.linkedWorkOrders.length === 0) {
          turn.linkedWorkOrders = dbMatch.linkedWorkOrders || [];
        }
        if (!turn.trackingUuid && dbMatch.trackingUuid) turn.trackingUuid = dbMatch.trackingUuid;
        if (!turn.trackingCode && dbMatch.trackingCode) turn.trackingCode = dbMatch.trackingCode;
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
  WEBHOOK EVENTS — AppFolio signed webhook ingest feed
  ================================================================= */
async function pollWebhookEvents() {
  if (!isProxySessionReady()) return false;
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
          var friendlyTitle = getWebhookDisplayTitle(e, meta);
          var friendlyDescription = getWebhookDisplayDescription(e, meta);
          WEBHOOK_EVENTS.push({
            id: e.id || 0,
            ts: e.ts || e.timestamp || new Date().toISOString(),
            type: meta.eventType || e.type || e.event_type || 'webhook',
            event_label: e.event_label || friendlyTitle,
            resource_type: meta.resourceType || '',
            resource_id: meta.resourceId || '',
            resource_name: e.resource_name || meta.resourceName || '',
            title: friendlyTitle,
            human_description: e.human_description || '',
            description: friendlyDescription,
            body: evtBody,
            priority: e.priority || 'normal',
            source: e.source || 'appfolio',
            raw: e.raw || null
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
    var vendor = VENDORS.find(function(v) {
      return String(v.id || '') === rid ||
        String(v.uuid || '') === rid ||
        String(v.vendor_uuid || '') === rid ||
        String(v.VendorUuid || '') === rid;
    });
    if (vendor) return vendor.name || '';
  }
  if (rtype === 'unit_turn') {
    var turn = TURNS.find(function(t) { return String(t.unitTurnId) === rid || String(t.unitId) === rid; });
    if (turn) return (turn.unit || rid.slice(0, 8)) + (turn.property ? ' - ' + turn.property : '');
  }
  if (rtype === 'inspection' || rtype === 'unit_inspection') {
    var insp = (typeof INSPECTIONS !== 'undefined' ? INSPECTIONS : []).find(function(i) {
      return String(i.id || i.uuid || i.inspectionId || '') === rid;
    });
    if (insp) return (insp.name || insp.inspectionName || 'Inspection') + (insp.propertyName ? ' - ' + insp.propertyName : '');
  }
  if (rtype === 'bill') {
    var bill = (typeof BILLS !== 'undefined' ? BILLS : []).find(function(b) {
      return String(b.id || '') === rid ||
        String(b.billId || '') === rid ||
        String((b.raw && (b.raw.Id || b.raw.id || b.raw.bill_id)) || '') === rid;
    });
    if (bill) return 'Bill' + (bill.billNumber ? ' #' + bill.billNumber : '') + (bill.vendorName ? ' - ' + bill.vendorName : '');
  }
  if (rtype === 'property_group') {
    var grp = PROPERTY_GROUPS.find(function(g) {
      return String(g.id || '') === rid ||
        String(g.uuid || '') === rid ||
        String(g.property_group_uuid || '') === rid ||
        String(g.GroupUuid || '') === rid ||
        String(g.Id || '') === rid;
    });
    if (grp) return grp.name || '';
  }
  // No local match and not yet resolved — caller should queue enrichment.
  return '\u2014\u00a0Unresolved';
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

function getWebhookDisplayTitle(e, meta) {
  meta = meta || extractWebhookMeta(e || {});
  var explicitTitle = String((e && (e.title || e.event_label)) || '').trim();
  return explicitTitle || decodeWebhookTitle(e || {}, meta) || 'Webhook Event';
}

function getWebhookDisplayDescription(e, meta) {
  meta = meta || extractWebhookMeta(e || {});
  var explicitDescription = String((e && (e.description || e.human_description)) || '').trim();
  return explicitDescription || getWebhookPreviewText(e || {}, meta) || '';
}

function extractWebhookMeta(e) {
  var payload = parseWebhookJson(e.payload) || parseWebhookJson(e.raw) || parseWebhookJson(e.body) || {};
  var topic = String(e.topic || payload.topic || '').trim().toLowerCase();
  var eventType = String(e.event_type || payload.event_type || e.type || '').trim().toLowerCase();
  var resourceType = normalizeWebhookResourceType(e.resource_type || payload.resource_type || '', topic);
  var resourceId = String(e.resource_id || payload.resource_id || '').trim();
  var resourceName = String(e.resource_name || payload.resource_name || '').trim();
  if (isUuidString(resourceName)) resourceName = '';
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

function webhookEventInCurrentGroup(e, meta) {
  if (!currentPropertyGroup) return true;
  meta = meta || extractWebhookMeta(e || {});
  var payload = (meta && meta.payload) || {};
  var propertyId = payload.property_id || payload.propertyId || payload.PropertyId || '';
  var propertyName = payload.property_name || payload.propertyName || payload.PropertyName || '';
  if (!propertyName && meta && meta.resourceType === 'property') {
    propertyName = meta.resourceName || '';
    if (!propertyId) propertyId = meta.resourceId || '';
  }
  if (!propertyId && !propertyName) return true;
  return isInPropertyGroup(propertyId, propertyName, currentPropertyGroup);
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
    var rawTitle = String(e.title || '');
    looksLikeFallbackProxyTitle = /(?:\u2192|->)\s*[a-z_]+\/[0-9a-f-]{8,}/i.test(rawTitle) ||
      /^[a-z_]+\s*\/\s*[0-9a-f-]{8,}$/i.test(rawTitle) ||
      /^[0-9a-f-]{16,}$/i.test(rawTitle);
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
  var resName = meta.resourceName || getWebhookReadableContext(meta, e) || '';
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
  var visibleEvents = WEBHOOK_EVENTS.filter(function(evt) {
    return webhookEventInCurrentGroup(evt);
  });
  if (countEl) countEl.textContent = visibleEvents.length;
  if (!el) return;
  if (visibleEvents.length === 0) {
    el.innerHTML = 'No events yet \u2014 waiting for signed AppFolio webhook deliveries.';
    return;
  }
  var html = '';
  visibleEvents.slice(0, 25).forEach(function(e, idx) {
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
    var displayTitle = getWebhookDisplayTitle(e, meta);
    var previewText = getWebhookDisplayDescription(e, meta);
    var rckey = (meta.resourceType && meta.resourceId) ? webhookResolveKey(meta.resourceType, meta.resourceId) : '';
    var resolved = rckey ? WEBHOOK_RESOURCE_CACHE[rckey] : null;

    // Build detail content
    var richTitle = (resolved && resolved.record) ? buildAppFolioStyleTitle(meta.resourceType, meta.eventType, resolved.record) : null;
    var objectLine = richTitle || (resolved && resolved.summary && resolved.summary.title) || '';
    var statusLine = (resolved && resolved.summary && resolved.summary.status) || '';
    var refLine    = (resolved && resolved.summary && resolved.summary.reference) ? '#' + resolved.summary.reference : '';
    var changes    = getWebhookChangeSummary(meta.payload || {});
    var rec        = (resolved && resolved.record) || {};
    var ctxParts   = [];
    if (rec.PropertyName) ctxParts.push(rec.PropertyName);
    if (rec.UnitName || rec.Unit) ctxParts.push(rec.UnitName || rec.Unit);
    if (rec.AssignedToName || rec.VendorName) ctxParts.push(rec.AssignedToName || rec.VendorName);
    if (ctxParts.length === 0) {
      var fallbackCtx = getWebhookReadableContext(meta, e);
      if (fallbackCtx) ctxParts.push(fallbackCtx);
    }

    html += '<div style="border-bottom:1px solid var(--border)">';
    // Header row — clickable to expand
    html += '<div class="wh-panel-item" data-whpanel="' + detailId + '" style="padding:5px 2px;display:flex;gap:6px;align-items:flex-start;cursor:pointer;user-select:none">';
    html += '<i class="fas ' + iconClass + '" style="color:' + iconColor + ';margin-top:2px;font-size:11px;width:14px;text-align:center;flex-shrink:0"></i>';
    html += '<div style="flex:1;min-width:0">';
    html += '<span style="color:var(--text-muted);font-size:10px">' + escapeHtml(e.ts ? timeAgo(e.ts) : '\u2014') + '</span> ';
    if (isPri) html += '<span style="color:var(--danger);font-weight:600">\u26a0 </span>';
    html += '<strong style="color:var(--text-primary);font-size:11px">' + escapeHtml(displayTitle) + '</strong>';
    if (previewText) html += '<div style="color:var(--text-secondary);margin-top:1px;font-size:10px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + escapeHtml(previewText) + '</div>';
    html += '</div>';
    html += '<i class="fas fa-chevron-down" style="font-size:9px;color:var(--text-muted);margin-top:3px;flex-shrink:0;transition:transform .15s" id="' + detailId + '_chev"></i>';
    html += '</div>';

    // Collapsible detail panel
    html += '<div id="' + detailId + '" style="display:none;padding:6px 8px 8px 20px;font-size:10px;font-family:var(--font-mono);color:var(--text-secondary);background:var(--bg-input);line-height:1.7">';
    if (objectLine) html += '<div><span style="color:var(--text-muted)">Object:</span> ' + escapeHtml(objectLine + (refLine ? ' ' + refLine : '')) + (statusLine ? ' <span style="color:var(--text-muted)">(' + escapeHtml(statusLine) + ')</span>' : '') + '</div>';
    if (previewText) html += '<div><span style="color:var(--text-muted)">Summary:</span> ' + escapeHtml(previewText) + '</div>';
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
  if (visibleEvents.length > 25) {
    html += '<div style="padding:4px 0;color:var(--text-muted);text-align:center;font-size:10px">\u2026 and ' + (visibleEvents.length - 25) + ' more</div>';
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
  if (!isProxySessionReady()) {
    body.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:20px;color:var(--text-muted)">Sign in to load webhook events</td></tr>';
    return;
  }
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
      var title = getWebhookDisplayTitle(evt, meta);
      var description = getWebhookDisplayDescription(evt, meta);
      return Object.assign({}, evt, {
        type: meta.eventType || evt.type || evt.event_type || 'webhook',
        resource_type: meta.resourceType || evt.resource_type || '',
        resource_id: meta.resourceId || evt.resource_id || '',
        resource_name: evt.resource_name || meta.resourceName || '',
        title: title,
        event_label: evt.event_label || title,
        description: description,
        human_description: evt.human_description || ''
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
        var title = getWebhookDisplayTitle(evt, meta);
        return Object.assign({}, evt, {
          resource_name: meta.resourceName || evt.resource_name || '',
          title: title,
          event_label: evt.event_label || title,
          description: getWebhookDisplayDescription(evt, meta)
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
    var displayTitle = getWebhookDisplayTitle(e, meta);
    var displayDescription = getWebhookDisplayDescription(e, meta);
    var isPri = e.priority === 'urgent' || e.priority === 'high';
    var priClass = isPri ? 'color:var(--danger);font-weight:600' : 'color:var(--text-secondary)';
    var rowNum = _whPage * _whPageSize + idx + 1;
    html += '<tr class="wh-data-row" data-whid="' + (e.id || idx) + '" style="cursor:pointer">';
    html += '<td style="font-family:var(--font-mono);font-size:11px;color:var(--text-muted)">' + rowNum + '</td>';
    html += '<td style="font-family:var(--font-mono);font-size:11px;white-space:nowrap">' + escapeHtml(e.ts ? timeAgo(e.ts) : '\u2014') + '</td>';
    html += '<td><span class="tag wh-type-' + escapeHtml(String(e.type || 'webhook').replace(/[^a-z0-9_-]/gi, '')) + '">' + escapeHtml(e.type || 'webhook') + '</span></td>';
    html += '<td style="max-width:300px">';
    html += '<div style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + escapeHtml(displayTitle || '\u2014') + '</div>';
    if (displayDescription) {
      html += '<div style="font-size:10px;color:var(--text-muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + escapeHtml(displayDescription) + '</div>';
    } else if (resolved && resolved.summary && resolved.summary.title) {
      html += '<div style="font-size:10px;color:var(--text-muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + escapeHtml(resolved.summary.title) + '</div>';
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
    var httpStatus = e.http_status || e.response_code || e.status_code || e.httpCode || 'N/A';
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
      html += escapeHtml(String(meta.resourceType || e.resource_type)) + ' / ⟳ resolving';
    } else {
      html += '(unknown)';
    }
    html += '<br>';
    html += '<strong>HTTP Status:</strong> ' + escapeHtml(String(httpStatus)) + '<br>';
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
  if (!isProxySessionReady()) {
    panel.classList.remove('hidden');
    content.innerHTML = 'Sign in to view webhook stats.';
    return;
  }
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
  if (isUuidString(wo.dbApiId || '')) return String(wo.dbApiId);
  // WORK_ORDERS items store the UUID in wo.uuid (r.work_order_id from Reports API)
  if (isUuidString(wo.uuid || '')) return String(wo.uuid);
  // 1. Check if wo already has a DB API link (contains UUID in path)
  if (wo.link) {
    var linkMatch = String(wo.link).match(/work_orders\/([0-9a-f\-]{36})/i);
    if (linkMatch) return linkMatch[1];
  }
  // 2. Look up matching DB API work order by WO number and turn context
  if (wo.id) {
    var requiredWoNumber = String(wo.id || '').trim();
    var requiredUnitTurnId = String(wo.unitTurnId || '').trim();
    var requiredUnitId = String(wo.unitId || '').trim();
    var dbWo = TURN_WORK_ORDERS.find(function(tw) {
      if (String(tw.woNumber || '').trim() !== requiredWoNumber) return false;
      if (requiredUnitTurnId && String(tw.unitTurnId || '').trim() !== requiredUnitTurnId) return false;
      if (requiredUnitId && String(tw.unitId || '').trim() !== requiredUnitId) return false;
      return isUuidString(tw.id || '');
    });
    if (dbWo && dbWo.id) return dbWo.id;
    var fallback = TURN_WORK_ORDERS.find(function(tw) {
      return String(tw.woNumber || '').trim() === requiredWoNumber && isUuidString(tw.id || '');
    });
    if (fallback && fallback.id) return fallback.id;
  }
  // 3. Fallback: use wo.uuid (may be numeric — DB API may reject it)
  return wo.uuid || '';
}

async function fetchWONotes(woIdOrUuid, woContext) {
  if (!woIdOrUuid) return [];
  var woRef = String(woIdOrUuid || '').trim();
  if (!isUuidString(woRef)) {
    var requiredUnitTurnId = String((woContext && woContext.unitTurnId) || '').trim();
    var requiredUnitId = String((woContext && woContext.unitId) || '').trim();
    var dbWo = TURN_WORK_ORDERS.find(function(tw) {
      if (String(tw.woNumber || '').trim() !== woRef) return false;
      if (requiredUnitTurnId && String(tw.unitTurnId || '').trim() !== requiredUnitTurnId) return false;
      if (requiredUnitId && String(tw.unitId || '').trim() !== requiredUnitId) return false;
      return isUuidString(tw.id || '');
    });
    if (dbWo && dbWo.id) {
      woRef = String(dbWo.id);
    } else {
      var fallback = TURN_WORK_ORDERS.find(function(tw) {
        return String(tw.woNumber || '').trim() === woRef && isUuidString(tw.id || '');
      });
      woRef = fallback ? String(fallback.id) : '';
    }
  }
  // Final fallback: ask proxy to resolve WO reference to DB API UUID.
  if ((!woRef || !isUuidString(woRef)) && API_PROXY) {
    try {
      var detailData = await proxyAction('wo_detail', { wo_id: String(woIdOrUuid || '') });
      var resolvedUuid = String((detailData && detailData.uuid) || (detailData && detailData.result && detailData.result.id) || '').trim();
      if (isUuidString(resolvedUuid)) woRef = resolvedUuid;
    } catch (e) {
      // Keep silent fallback behavior for modal; render function shows empty state.
    }
  }
  if (!woRef || !isUuidString(woRef)) return [];
  var notesCached = detailCacheGet('notes_' + woRef);
  if (typeof notesCached !== 'undefined') return notesCached;
  try {
    // Use dedicated proxy action with v0 credentials (like property_groups)
    var data = await proxyAction('wo_notes', { wo_id: woRef });
    var notes = (data && data.results) ? data.results :
                (data && data.Results) ? data.Results :
                (data && data.data) ? data.data :
                (data && data.Data) ? data.Data :
                (Array.isArray(data) ? data : []);
    detailCacheSet('notes_' + woRef, notes);
    return notes;
  } catch (e) { return []; }
}

function renderWONotesList(notes) {
  var nl = document.getElementById('detailNotesList');
  if (!nl) return;
  if (!notes || notes.length === 0) {
    nl.innerHTML = '<div style="text-align:center;padding:10px;color:var(--text-muted);font-size:12px">No notes yet.</div>';
    return;
  }
  var nh = '';
  notes.forEach(function(n) {
    var createdBy = n.CreatedBy || n.created_by || n.Author || n.author || n.UserName || '—';
    var createdAt = n.CreatedAt || n.created_at || n.UpdatedAt || n.updated_at || '';
    var body = n.Body || n.body || n.Content || n.content || n.Note || n.note || n.Message || n.message || '';
    nh += '<div class="note-item"><div class="note-item-header"><span>' + escapeHtml(createdBy) + '</span><span>' + escapeHtml(formatNoteDateTime(createdAt)) + '</span></div>';
    nh += '<div class="note-item-body">' + escapeHtml(body || '—') + '</div></div>';
  });
  nl.innerHTML = nh;
}

function normalizeWOAttachmentList(data) {
  if (!data) return [];
  if (Array.isArray(data)) return data;
  if (Array.isArray(data.attachments)) return data.attachments;
  if (Array.isArray(data.results)) return data.results;
  if (Array.isArray(data.Results)) return data.Results;
  if (Array.isArray(data.items)) return data.items;
  if (Array.isArray(data.data)) return data.data;
  return [];
}

function renderWOAttachmentsList(attachments, errorText) {
  var el = document.getElementById('detailAttachmentList');
  var metaEl = document.getElementById('detailAttachmentTypes');
  if (!el) return;
  if (errorText) {
    el.innerHTML = '<div style="color:var(--danger);font-size:11px;line-height:1.5">Failed to load attachments. ' + escapeHtml(errorText) + '</div>';
    if (metaEl) metaEl.textContent = 'debug: attachment request failed';
    return;
  }
  if (!attachments || !attachments.length) {
    el.innerHTML = '<div style="color:var(--text-muted);font-size:11px">No attachments on file.</div>';
    if (metaEl) metaEl.textContent = 'none found';
    return;
  }
  var contentTypes = [];
  var html = attachments.map(function(att) {
    var name = String(att.FileName || att.file_name || att.Name || att.name || att.Id || att.id || 'Attachment');
    var contentType = String(att.ContentType || att.content_type || att.MimeType || att.mime_type || 'unknown').trim();
    var createdAt = String(att.CreatedAt || att.created_at || att.UpdatedAt || att.updated_at || '');
    var url = String(att.DownloadUrl || att.download_url || att.Url || att.url || att.FileUrl || att.file_url || '').trim();
    if (contentType && contentTypes.indexOf(contentType) === -1) contentTypes.push(contentType);
    return '<div style="display:flex;justify-content:space-between;gap:10px;align-items:flex-start;padding:8px 0;border-bottom:1px solid rgba(148,163,184,.12)">' +
      '<div>' +
        '<div style="font-size:12px;font-weight:700">' + escapeHtml(name) + '</div>' +
        '<div style="font-size:11px;color:var(--text-muted)">' + escapeHtml(contentType || 'unknown') + (createdAt ? ' · ' + escapeHtml(formatNoteDateTime(createdAt)) : '') + '</div>' +
      '</div>' +
      (url ? '<a href="' + escapeHtml(url) + '" target="_blank" rel="noopener noreferrer" class="action-btn" style="padding:3px 8px;text-decoration:none">Open</a>' : '') +
    '</div>';
  }).join('');
  el.innerHTML = html;
  if (metaEl) metaEl.textContent = contentTypes.length ? contentTypes.join(', ') : 'unknown';
}

async function loadWOAttachments(woIdOrUuid) {
  var el = document.getElementById('detailAttachmentList');
  if (!el || !woIdOrUuid) return;
  el.innerHTML = '<div style="color:var(--text-muted);font-size:11px"><i class="fas fa-spinner fa-spin"></i> Loading attachments…</div>';
  try {
    var data = await proxyAction('wo_attachments', { wo_id: String(woIdOrUuid) }, { suppressSessionExpiry: true });
    if (!data || data.ok === false) {
      renderWOAttachmentsList([], String((data && (data.detail || data.error)) || 'No detail from proxy'));
      return;
    }
    renderWOAttachmentsList(normalizeWOAttachmentList(data));
  } catch (e) {
    renderWOAttachmentsList([], String((e && e.message) || e || 'Request failed'));
  }
}

function extractWODetailRecord(data) {
  if (!data) return null;
  if (data.result && typeof data.result === 'object') return data.result;
  if (data.work_order && typeof data.work_order === 'object') return data.work_order;
  if (Array.isArray(data.results) && data.results[0]) return data.results[0];
  if (Array.isArray(data.Results) && data.Results[0]) return data.Results[0];
  return typeof data === 'object' ? data : null;
}

function writeDetailText(id, value) {
  var el = document.getElementById(id);
  if (!el) return;
  el.textContent = String(value || '—');
}

async function loadWODetailExtras(wo, woDbUuid) {
  var ref = woDbUuid || (wo && (wo.uuid || wo.id));
  if (!ref) return;
  try {
    var data = await proxyAction('wo_detail', { wo_id: String(ref) });
    var detail = extractWODetailRecord(data);
    if (!detail) return;
    var createdAt = detail.CreatedAt || detail.created_at || detail.CreatedOn || detail.created_on || wo.created || wo.createdAt || '';
    var jobDescription = detail.JobDescription || detail.job_description || detail.Description || detail.description || wo.description || '—';
    var permission = detail.PermissionToEnter || detail.permission_to_enter || detail.AccessNotes || detail.access_notes || wo.permissionToEnter || '—';
    var vendorInstructions = detail.VendorInstructions || detail.vendor_instructions || detail.VendorNotes || detail.vendor_notes || '—';
    var estimates = detail.Estimates || detail.estimates || detail.EstimatedAmount || detail.estimated_amount || '';
    var tenantPhone = detail.PrimaryTenantPhoneNumber || detail.primary_tenant_phone_number || detail.TenantPhoneNumber || detail.tenant_phone_number || wo.tenantPhone || '—';
    var tenantEmail = detail.PrimaryTenantEmail || detail.primary_tenant_email || detail.TenantEmail || detail.tenant_email || wo.tenantEmail || '—';
    var assignedTo = detail.AssignedUsers || detail.assigned_users || detail.AssignedTo || detail.assigned_to || wo.assignedUser || '—';
    writeDetailText('detailCreatedAt', createdAt ? formatNoteDateTime(createdAt) : '—');
    writeDetailText('detailJobDescription', jobDescription || '—');
    writeDetailText('detailPermission', permission || '—');
    writeDetailText('detailVendorInstructions', vendorInstructions || '—');
    writeDetailText('detailEstimates', Array.isArray(estimates) ? estimates.join(', ') : (estimates || '—'));
    writeDetailText('detailTenantPhone', tenantPhone || '—');
    writeDetailText('detailTenantEmail', tenantEmail || '—');
    writeDetailText('detailAssignedTo', Array.isArray(assignedTo) ? assignedTo.join(', ') : assignedTo || '—');
  } catch (e) {
    writeDetailText('detailAttachmentTypes', 'debug: live detail unavailable');
  }
}

// bypassProxyCache=true: skips the 5-min Turso cache via ?path= passthrough,
// used after a successful note post so the new note appears immediately.
function refreshCurrentWONotes(forceUuid, bypassProxyCache) {
  var modal = document.getElementById('woModal');
  if (!modal || !modal.classList.contains('show') || !CURRENT_WO_MODAL) return;
  var targetUuid = forceUuid || CURRENT_WO_MODAL.woDbUuid;
  if (!targetUuid) return;
  delete WO_DETAIL_CACHE['notes_' + targetUuid];
  var fetchPromise;
  if (bypassProxyCache && API_PROXY) {
    // Use apiFetch which routes via ?path= compat mode, bypassing wo_notes Turso cache
    var notesApiPath = '/api/v0/work_orders/' + encodeURIComponent(targetUuid) + '/notes';
    fetchPromise = apiFetch(notesApiPath)
      .then(function(data) {
        var notes = (data && data.results) ? data.results
          : (data && data.Results) ? data.Results
          : (data && data.data) ? data.data
          : (data && data.Data) ? data.Data
          : (Array.isArray(data) ? data : []);
        detailCacheSet('notes_' + targetUuid, notes);
        return notes;
      })
      .catch(function() { return detailCacheGet('notes_' + targetUuid) || []; });
  } else {
    fetchPromise = fetchWONotes(targetUuid);
  }
  fetchPromise.then(function(notes) {
    if (!CURRENT_WO_MODAL || CURRENT_WO_MODAL.woDbUuid !== targetUuid) return;
    renderWONotesList(notes);
  });
}

async function postWONoteViaProxy(woDbUuid, noteText) {
  if (!woDbUuid) return { ok: false, status: 0, message: 'Missing work order UUID' };
  if (API_PROXY) {
    try {
      var postResp = await proxyPost('wo_note_create', {
        uuid: String(woDbUuid),
        body_text: String(noteText || '')
      });
      if (!postResp || postResp.ok === false) {
        return {
          ok: false,
          status: Number((postResp && postResp.status) || 500),
          message: String((postResp && (postResp.message || postResp.error)) || 'Request failed')
        };
      }
      return { ok: true };
    } catch (e) {
      return { ok: false, status: 500, message: String((e && e.message) || e || 'Request failed') };
    }
  }

  var path = '/api/v0/work_orders/' + encodeURIComponent(String(woDbUuid)) + '/notes';
  var headers = { 'Accept': 'application/json', 'Content-Type': 'application/json' };
  {
    var auth = getAuthHeader();
    var devId = getDevId();
    if (auth) headers['Authorization'] = auth;
    if (devId) headers['X-AppFolio-Developer-ID'] = devId;
  }

  var res = await fetchWithTimeout(resolveUrl(path, 'POST'), {
    method: 'POST',
    headers: headers,
    body: JSON.stringify({ Body: String(noteText || '') })
  }, 30000);

  if (res.ok) return { ok: true };

  var raw = '';
  try { raw = await res.text(); } catch (e) { /* empty */ }
  var msg = raw || res.statusText || 'Request failed';
  try {
    var parsed = JSON.parse(raw || '{}');
    msg = parsed.error || parsed.message || parsed.detail || parsed.title || msg;
  } catch (e) { /* keep raw body */ }

  return { ok: false, status: res.status, message: String(msg || 'Request failed') };
}

async function fetchWOBilledAmount(woNumber) {
  if (!woNumber) return { ok: true, total_billed: 0 };
  return await proxyAction('wo_billed_amount', { wo_number: String(woNumber) });
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
    if (!isInPropertyGroup(mo.propertyId, mo.property, currentPropertyGroup)) return;
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
    if (!isInPropertyGroup(r.propertyId, r.propertyName, currentPropertyGroup)) return;
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
  if (DASH_TURN_VIEW_MODE === 'compact') {
    if (document.body && document.body.classList.contains('tv-mode')) return 12;
    if (window.matchMedia && window.matchMedia('(max-width: 700px)').matches) return 8;
    return 10;
  }
  if (document.body && document.body.classList.contains('tv-mode')) return 6;
  if (window.matchMedia && window.matchMedia('(max-width: 700px)').matches) return 1;
  if (window.matchMedia && window.matchMedia('(max-width: 1100px)').matches) return 2;
  return 4;
}

function renderDashboardTurnSyncLabel() {
  var el = $('#dashTurnSyncAgo');
  if (!el) return;
  if (!DASH_TURN_LAST_SYNC_AT) {
    el.textContent = 'Never synced';
    return;
  }
  var ms = new Date(DASH_TURN_LAST_SYNC_AT).getTime();
  if (isNaN(ms)) {
    el.textContent = 'Synced';
    return;
  }
  var mins = Math.floor((Date.now() - ms) / 60000);
  if (mins <= 0) el.textContent = 'Just now';
  else el.textContent = mins + 'm ago';
}

function renderTurnDashboardTicker(entries) {
  var wrap = $('#dashTurnTicker');
  var track = $('#dashTurnTickerTrack');
  if (!wrap || !track) return;
  if (!entries || !entries.length) {
    wrap.style.display = 'none';
    track.innerHTML = '';
    return;
  }
  var top = entries.slice(0, Math.min(entries.length, 12));
  var doubled = top.concat(top);
  var html = doubled.map(function(p) {
    var tone = p.isStalled ? 'stalled' : (p.isUpcoming ? 'upcoming' : 'active');
    var woOpen = p.matchingWOs.filter(function(wo) { return !isClosedTurnWorkOrderStatus(wo.status); }).length;
    return '<span class="turn-dash-ticker-item">' +
      '<span class="turn-dash-ticker-dot ' + tone + '"></span>' +
      '<strong>' + escapeHtml(p.unit || 'Unit') + '</strong>' +
      '<span>' + escapeHtml(p.property || '') + '</span>' +
      '<span>' + Math.max(0, p.elapsed) + 'd</span>' +
      '<span>' + woOpen + ' open WO</span>' +
      '</span>';
  }).join('');
  track.innerHTML = html;
  track.classList.toggle('no-anim', top.length <= 3);
  wrap.style.display = '';
}

async function syncTurnDashboardIncremental(showMsg) {
  if (!API_PROXY) return false;
  var btn = $('#dashTurnSync');
  if (btn) { btn.disabled = true; btn.classList.add('spinning'); }
  try {
    var sinceIso = DASH_TURN_LAST_SYNC_AT || new Date(Date.now() - 12 * 60 * 60 * 1000).toISOString();
    var data = await proxyAction('turns_incremental', { since: sinceIso, limit: '800' });
    var rows = data.results || [];
    var merged = 0;
    if (rows.length) {
      var byId = {};
      TURNS.forEach(function(t) {
        byId[String(t.unitTurnId || '')] = t;
      });
      rows.forEach(function(raw) {
        var normalized = normalizeTurnRecord(raw);
        var key = String(normalized.unitTurnId || '');
        if (!key) return;
        byId[key] = Object.assign({}, byId[key] || {}, normalized);
        merged++;
      });
      TURNS = Object.keys(byId).map(function(k) { return byId[k]; });
      renderTurnBoard();
    }
    DASH_TURN_LAST_SYNC_AT = data.synced_at || new Date().toISOString();
    renderDashboardTurnSyncLabel();
    if (showMsg) {
      showToast('Turn sync complete — ' + merged + ' updated');
    }
    return true;
  } catch (err) {
    if (showMsg) showToast('Turn sync failed: ' + (err.message || err), 'warning');
    return false;
  } finally {
    if (btn) { btn.disabled = false; btn.classList.remove('spinning'); }
  }
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
  var isCompact = DASH_TURN_VIEW_MODE === 'compact';
  var pageSize = getDashboardTurnPageSize();
  var totalPages = Math.max(1, Math.ceil(entries.length / pageSize));
  if (DASH_TURN_PAGE >= totalPages) DASH_TURN_PAGE = 0;
  strip.classList.toggle('compact-mode', isCompact);
  renderTurnDashboardTicker(entries);
  renderDashboardTurnSyncLabel();

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
    if (isCompact) {
      var compactTone = p.isStalled ? 'stalled' : (p.isUpcoming ? 'upcoming' : 'active');
      return '<button class="turn-dash-row" data-turndash-open="' + escapeHtml(p.id) + '">' +
        '<span class="turn-dash-row-left">' +
          '<span class="turn-dash-row-title">' + escapeHtml(p.unit || 'Unit') + '</span>' +
          '<span class="turn-dash-row-sub">' + escapeHtml(p.property || '') + ' • ' + pmName + '</span>' +
        '</span>' +
        '<span class="turn-dash-row-right">' +
          '<span class="turn-dash-row-pill ' + compactTone + '">' + escapeHtml(statusLabel) + '</span>' +
          '<span class="turn-dash-row-pill">' + Math.max(0, p.elapsed) + 'd</span>' +
          '<span class="turn-dash-row-pill">' + openWOs + ' open</span>' +
        '</span>' +
      '</button>';
    }
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

// ─── Billing / AP Section ──────────────────────────────────────────────────
var _billsPage = 0;
var BILLS_PAGE_SIZE = 50;
var _billingLoadPromise = null;
var _billingQueuedOpts = null;
var _billingListCacheRows = [];
var _billingListCacheKey = '';

// ── Client-side bill filtering helpers ────────────────────────────────────
async function fetchAllBillsToCache() {
  if (__CACHED_BILLS.length > 0) return true; // Already cached
  if (__CACHED_BILLS_IN_FLIGHT) return false; // Fetch in progress
  
  __CACHED_BILLS_IN_FLIGHT = true;
  try {
    // Fetch all bills at once with high per_page limit to minimize API calls
    var payload = await fetchBills(DEFAULT_BILLS_LOOKBACK_DAYS, {
      scoped: false, // Bypass group scoping at fetch time (apply client-side)
      forceRefresh: false,
      filterType: '',
      filterValue: '',
      limit: 3000,
      perPage: 3000,
      page: 1,
      offset: 0,
      max: 3000,
      assignGlobal: false,
      returnPayload: true,
    });
    
    if (payload && payload.ok && Array.isArray(payload.rows)) {
      __CACHED_BILLS = payload.rows.slice();
      __CACHED_BILLS_LOADED_AT = Date.now();
      console.log('[Bill Cache] Loaded ' + __CACHED_BILLS.length + ' bills');
      return true;
    }
    return false;
  } catch (err) {
    console.warn('[Bill Cache] Fetch failed:', err.message || err);
    return false;
  } finally {
    __CACHED_BILLS_IN_FLIGHT = false;
  }
}

function applyLocalBillFilters(filterOpts) {
  filterOpts = filterOpts || {};
  if (__CACHED_BILLS.length === 0) return []; // No cache yet
  
  var groupToFilterBy = String(filterOpts.groupId || '').trim();
  var groupNameToFilterBy = String(filterOpts.groupName || '').trim();
  var statusToFilterBy = filterOpts.status || '';
  var searchText = (filterOpts.search || '').toLowerCase();
  var hasGroupMaps = !!(Object.keys(_nameToGroups || {}).length || Object.keys(_uuidToGroups || {}).length || Object.keys(_idToGroups || {}).length);
  
  return __CACHED_BILLS.filter(function(b) {
    // Filter by property group if specified
    if (groupToFilterBy || groupNameToFilterBy) {
      var effectiveGroupName = groupNameToFilterBy || groupToFilterBy;
      if (!billMatchesGroupScope(b, effectiveGroupName, groupToFilterBy, hasGroupMaps)) return false;
    }
    
    // Filter by status if specified
    if (statusToFilterBy && String(b.status || '').toLowerCase() !== statusToFilterBy.toLowerCase()) {
      return false;
    }
    
    // Filter by search text if specified
    if (searchText) {
      var hay = [
        b.vendorName,
        b.vendorId,
        b.payeeUuid,
        b.vendorUuid,
        b.propertyName,
        b.propertyId,
        b.unitId,
        b.propertyGroup,
        b.propertyManager,
        b.workOrderId,
        b.billNumber,
        b.id,
      ].join(' ').toLowerCase();
      if (hay.indexOf(searchText) === -1) return false;
    }
    
    return true;
  });
}

function getBillingDateRangeKey() {
  var from = String((document.getElementById('billUpdatedFrom') || {}).value || '').trim();
  var to = String((document.getElementById('billUpdatedTo') || {}).value || '').trim();
  return from + '|' + to;
}

function billMatchesGroupScope(b, grp, groupUuid, hasGroupMaps) {
  if (!grp) return true;
  var normalizedGroupUuid = String(groupUuid || '').trim().toLowerCase();
  var normalizedGroupName = String(grp || '').trim().toLowerCase();
  var billGroupUuid = String(b.propertyGroupId || resolveBillGroupUuidFromRecord(b.raw || b) || '').trim().toLowerCase();
  var billGroupName = String(resolveBillGroupName(b) || b.propertyGroup || b._propertyGroup || '').trim().toLowerCase();
  var raw = (b && b.raw && typeof b.raw === 'object') ? b.raw : b;
  var billPropertyId = String((b && b.propertyId) || raw.property_id || raw.PropertyId || raw.property_uuid || raw.PropertyUuid || '').trim();
  var billPropertyName = String((b && b.propertyName) || raw.property_name || raw.PropertyName || raw.property || raw.Property || '').trim();

  if (normalizedGroupUuid && billGroupUuid) return billGroupUuid === normalizedGroupUuid;
  if (normalizedGroupName && billGroupName) return billGroupName === normalizedGroupName;
  if (hasGroupMaps) return isInPropertyGroup(billPropertyId, billPropertyName, grp);
  // Strict scope: if a group filter is active and we cannot prove membership, exclude.
  return false;
}

function isBillingListRouteActive() {
  return String(_billingRouteAction || 'bills_list').trim() === 'bills_list' &&
    !_billingRouteFilterValue && !_billingDueFrom && !_billingDueTo;
}

function applyBillingListScopeFromCache(opts) {
  opts = opts || {};
  if (!Array.isArray(_billingListCacheRows) || _billingListCacheRows.length === 0) return false;

  var grp = normalizeGroupSelectionValue(getEffectiveGroupId());
  var groupUuid = getEffectiveGroupUuid(grp);
  var hasGroupMaps = !!(Object.keys(_nameToGroups || {}).length || Object.keys(_uuidToGroups || {}).length || Object.keys(_idToGroups || {}).length);
  var scopedRows = _billingListCacheRows.filter(function(b) {
    return billMatchesGroupScope(b, grp, groupUuid, hasGroupMaps);
  });

  if (opts.resetPage) _billsPage = 0;
  _billingServerTotal = scopedRows.length;
  _billingServerTotalPages = Math.max(1, Math.ceil(_billingServerTotal / BILLS_PAGE_SIZE));
  _billingServerPage = Math.max(1, Math.min(_billingServerTotalPages, (_billsPage + 1)));
  _billsPage = _billingServerPage - 1;

  var start = _billsPage * BILLS_PAGE_SIZE;
  window._billingPageRows = scopedRows.slice(start, start + BILLS_PAGE_SIZE);
  renderBillsSection();
  return true;
}

async function loadBillingPage(opts) {
  if (_billingLoadPromise) {
    var incoming = opts || {};
    // Ignore duplicate manual refresh taps while a request is active.
    // We still allow queued follow-ups for state-changing requests.
    if (incoming.forceRefresh && !incoming.resetPage && !incoming.forceHardLock) {
      return _billingLoadPromise;
    }
    var queued = _billingQueuedOpts || {};
    _billingQueuedOpts = {
      resetPage: !!(queued.resetPage || incoming.resetPage),
      forceRefresh: !!(queued.forceRefresh || incoming.forceRefresh),
      forceHardLock: !!(queued.forceHardLock || incoming.forceHardLock),
    };
    return _billingLoadPromise;
  }

  _billingLoadPromise = (async function() {
  opts = opts || {};
  if (opts.resetPage) _billsPage = 0;
  
  // Prefetch all bills to cache on first load
  if (__CACHED_BILLS.length === 0 && !__CACHED_BILLS_IN_FLIGHT) {
    await fetchAllBillsToCache();
  }

  var hardLock = !Array.isArray(window._billingPageRows) || window._billingPageRows.length === 0;
  if (opts.forceHardLock) hardLock = true;
  if (hardLock) setSectionBusy('sec-billing', true, 'Refreshing billing data…');

  var refreshBtn = $('#btnRefreshBills');
  var refreshText = refreshBtn ? refreshBtn.innerHTML : '';
  if (refreshBtn) {
    refreshBtn.disabled = true;
    refreshBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Loading…';
  }

  // Read Last Updated date range from UI (default 7 days)
  var billUpdatedFrom = String((document.getElementById('billUpdatedFrom') || {}).value || '').trim();
  var billUpdatedTo = String((document.getElementById('billUpdatedTo') || {}).value || '').trim();
  var lookbackDays = DEFAULT_BILLS_LOOKBACK_DAYS;
  if (billUpdatedFrom) {
    var fromMs = new Date(billUpdatedFrom).getTime();
    if (!isNaN(fromMs)) lookbackDays = Math.max(1, Math.ceil((Date.now() - fromMs) / 86400000));
  }

  var isListRoute = isBillingListRouteActive();

  if (isListRoute && __CACHED_BILLS.length > 0) {
    // Use cached bills with client-side group filtering
    var grp = normalizeGroupSelectionValue(getEffectiveGroupId());
    var groupUuid = getEffectiveGroupUuid(grp);
    var filteredBills = applyLocalBillFilters({ groupId: groupUuid, groupName: grp });
    
    _billingServerTotal = filteredBills.length;
    _billingServerTotalPages = Math.max(1, Math.ceil(_billingServerTotal / BILLS_PAGE_SIZE));
    _billingServerPage = 1;
    _billsPage = 0;
    
    var start = _billsPage * BILLS_PAGE_SIZE;
    window._billingPageRows = filteredBills.slice(start, start + BILLS_PAGE_SIZE);
    _lastBillSource = 'cached';
    renderBillsSection(filteredBills);
  } else if (isListRoute) {
    var cacheKey = getBillingDateRangeKey();
    var useCachedList = !opts.forceRefresh && _billingListCacheKey === cacheKey && Array.isArray(_billingListCacheRows) && _billingListCacheRows.length > 0;
    var listPayload = null;
    var listRows = [];

    if (useCachedList) {
      listRows = _billingListCacheRows.slice();
      _lastBillSource = 'cached';
    } else {
      listPayload = await fetchBills(lookbackDays, {
        scoped: false,
        forceRefresh: !!opts.forceRefresh,
        filterType: '',
        filterValue: '',
        dueFrom: '',
        dueTo: '',
        updatedFrom: billUpdatedFrom,
        updatedTo: billUpdatedTo,
        routeStatusFilter: '',
        limit: 3000,
        perPage: 3000,
        page: 1,
        offset: 0,
        max: 3000,
        assignGlobal: false,
        returnPayload: true,
      });

      if (!listPayload || !listPayload.ok) {
        if (hardLock) setSectionBusy('sec-billing', false);
        if (refreshBtn) {
          refreshBtn.disabled = false;
          refreshBtn.innerHTML = refreshText;
        }
        showToast('Billing refresh failed: ' + String(listPayload && listPayload.error || 'unknown error'), { kind: 'warning' });
        return;
      }

      listRows = Array.isArray(listPayload.rows) ? listPayload.rows : [];
      _billingListCacheRows = listRows.slice();
      _billingListCacheKey = cacheKey;
      _lastBillSource = String(listPayload.source || (listPayload.fromCache ? 'cached' : 'live') || 'legacy');
    }

    applyBillingListScopeFromCache();
  } else {
    var payload = await fetchBills(lookbackDays, {
      scoped: true,
      forceRefresh: !!opts.forceRefresh,
      filterType: _billingRouteAction,
      filterValue: _billingRouteFilterValue,
      dueFrom: _billingDueFrom,
      dueTo: _billingDueTo,
      updatedFrom: billUpdatedFrom,
      updatedTo: billUpdatedTo,
      routeStatusFilter: _billingRouteStatus,
      limit: BILLS_PAGE_SIZE,
      perPage: BILLS_PAGE_SIZE,
      page: _billsPage + 1,
      offset: _billsPage * BILLS_PAGE_SIZE,
      assignGlobal: false,
      returnPayload: true,
    });

    if (!payload || !payload.ok) {
      if (hardLock) setSectionBusy('sec-billing', false);
      if (refreshBtn) {
        refreshBtn.disabled = false;
        refreshBtn.innerHTML = refreshText;
      }
      showToast('Billing refresh failed: ' + String(payload && payload.error || 'unknown error'), { kind: 'warning' });
      return;
    }

    window._billingPageRows = Array.isArray(payload.rows) ? payload.rows : [];
    _billingServerTotal = Number(payload.total || window._billingPageRows.length) || 0;
    _billingServerTotalPages = Math.max(1, Number(payload.totalPages || 1) || 1);
    _billingServerPage = Math.max(1, Number(payload.page || (_billsPage + 1)) || 1);
    _billsPage = _billingServerPage - 1;
    _lastBillSource = String(payload.source || (payload.fromCache ? 'cached' : 'live') || 'legacy');
  }

  renderBillsSection();

  if (hardLock) setSectionBusy('sec-billing', false);
  if (refreshBtn) {
    refreshBtn.disabled = false;
    refreshBtn.innerHTML = refreshText;
  }
  })();

  try {
    return await _billingLoadPromise;
  } finally {
    _billingLoadPromise = null;
    if (_billingQueuedOpts) {
      var nextOpts = _billingQueuedOpts;
      _billingQueuedOpts = null;
      // Run one queued refresh after the active request settles.
      loadBillingPage(nextOpts);
    }
  }
}

function renderBillingPaginationControls(footer, rowsCount) {
  if (!footer) return;
  if (_billingServerTotal <= 0) {
    footer.style.display = 'none';
    return;
  }

  var start = (_billingServerPage - 1) * BILLS_PAGE_SIZE + 1;
  var end = Math.min(start + Math.max(0, rowsCount - 1), _billingServerTotal);
  var prevDisabled = _billingServerPage <= 1 ? ' disabled' : '';
  var nextDisabled = _billingServerPage >= _billingServerTotalPages ? ' disabled' : '';

  footer.style.display = '';
  footer.innerHTML =
    '<div style="display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap">' +
      '<span>Showing ' + start + '–' + end + ' of ' + _billingServerTotal + ' bills</span>' +
      '<span style="display:inline-flex;align-items:center;gap:8px">' +
        '<button class="action-btn" id="billPagePrev" style="padding:3px 8px"' + prevDisabled + '>Prev</button>' +
        '<span style="font-family:var(--font-mono);font-size:11px">Page ' + _billingServerPage + ' / ' + _billingServerTotalPages + '</span>' +
        '<button class="action-btn" id="billPageNext" style="padding:3px 8px"' + nextDisabled + '>Next</button>' +
        '<select id="billPageSize" style="font-family:var(--font-mono);font-size:11px;padding:4px 8px;border-radius:6px;border:1px solid var(--border);background:var(--bg-input);color:var(--text-primary)">' +
          '<option value="25"' + (BILLS_PAGE_SIZE === 25 ? ' selected' : '') + '>25</option>' +
          '<option value="50"' + (BILLS_PAGE_SIZE === 50 ? ' selected' : '') + '>50</option>' +
          '<option value="100"' + (BILLS_PAGE_SIZE === 100 ? ' selected' : '') + '>100</option>' +
        '</select>' +
      '</span>' +
    '</div>';

  var prevBtn = $('#billPagePrev');
  var nextBtn = $('#billPageNext');
  var sizeSel = $('#billPageSize');
  if (prevBtn) {
    prevBtn.addEventListener('click', function() {
      if (_billingServerPage <= 1) return;
      _billsPage = _billingServerPage - 2;
      loadBillingPage();
    });
  }
  if (nextBtn) {
    nextBtn.addEventListener('click', function() {
      if (_billingServerPage >= _billingServerTotalPages) return;
      _billsPage = _billingServerPage;
      loadBillingPage();
    });
  }
  if (sizeSel) {
    sizeSel.addEventListener('change', function() {
      var next = Math.max(1, Math.min(200, parseInt(String(sizeSel.value || BILLS_PAGE_SIZE), 10) || BILLS_PAGE_SIZE));
      if (next === BILLS_PAGE_SIZE) return;
      BILLS_PAGE_SIZE = next;
      _billsPage = 0;
      loadBillingPage({ resetPage: true });
    });
  }
}

function renderBillsSection(preFilteredData) {
  var tbody = $('#billBody');
  var footer = $('#billFooter');
  if (!tbody) return;

  var grp = normalizeGroupSelectionValue(getEffectiveGroupId());
  var groupUuid = grp ? (forcedPropertyGroupUuid || resolveGroupUuidFromName(grp)) : '';
  var hasGroupMaps = !!(Object.keys(_nameToGroups || {}).length || Object.keys(_uuidToGroups || {}).length || Object.keys(_idToGroups || {}).length);
  var activePropertyId = String(window.filteredPropertyId || '').trim();
  var activeUnitId = String(window.filteredUnitId || '').trim();
  var statusFilter = ($('#billStatusFilter') ? $('#billStatusFilter').value : '').toLowerCase();
  var search = ($('#billSearch') ? $('#billSearch').value : '').trim().toLowerCase();
  var now = Date.now();
  var thirtyDaysAgo = now - 30 * 86400000;

  renderBillingPropertyScopeChip(activePropertyId, String(window.filteredPropertyName || ''), activeUnitId, String(window.filteredUnitName || ''));

  // Use pre-filtered data if provided (from cache), otherwise use current page rows
  var sourceRows = Array.isArray(preFilteredData) && preFilteredData.length > 0
    ? preFilteredData
    : (Array.isArray(window._billingPageRows) && window._billingPageRows.length
      ? window._billingPageRows
      : (BILLS || []));

  // Local visual filters (search + property drilldown) on current server page.
  var filtered = sourceRows.filter(function(b) {
    if (!billMatchesGroupScope(b, grp, groupUuid, hasGroupMaps)) return false;
    if (statusFilter) {
      var st = getBillStatusKey(b);
      if (statusFilter === 'pending_approval' && st !== 'pending_approval') return false;
      if (statusFilter === 'approved' && st !== 'approved') return false;
      if (statusFilter === 'paid' && st !== 'paid') return false;
      if (statusFilter === 'void' && st !== 'void') return false;
    }
    if (search) {
      var hay = [
        b.vendorName,
        b.vendorId,
        b.payeeUuid,
        b.vendorUuid,
        b.propertyName,
        b.propertyId,
        b.unitId,
        b.propertyGroup,
        b.propertyManager,
        b.workOrderId,
        b.billNumber,
        b.id,
      ].join(' ').toLowerCase();
      if (hay.indexOf(search) === -1) return false;
    }
    if (activePropertyId) {
      if (String(b.propertyId || '').trim() !== activePropertyId) return false;
    }
    if (activeUnitId) {
      if (String(b.unitId || b.unit_id || '').trim() !== activeUnitId) return false;
    }
    return true;
  });

  // KPIs
  var pendingBills = filtered.filter(function(b) {
    var statusKey = getBillStatusKey(b);
    return statusKey === 'pending_approval';
  });
  var approvedBills = filtered.filter(function(b) { return getBillStatusKey(b) === 'approved'; });
  var paidBills = filtered.filter(function(b) {
    var st = getBillStatusKey(b);
    var d = new Date(b.date || '');
    return st === 'paid' && !isNaN(d.getTime()) && d.getTime() >= thirtyDaysAgo;
  });
  var uniqueVendors = new Set(filtered.map(function(b) { return b.vendorId || b.vendorName; }).filter(Boolean));
  var outstandingTotal = approvedBills.reduce(function(sum, b) {
    var n = amountToNumber(b.amount || 0);
    return sum + n;
  }, 0);
  var paidTotal = paidBills.reduce(function(sum, b) {
    var n = amountToNumber(b.amount || 0);
    return sum + n;
  }, 0);

  var setKpi = function(id, val) { var el = document.getElementById(id); if (el) el.textContent = val; };
  setKpi('billKpiPending', String(pendingBills.length));
  setKpi('billKpiPendingSub', pendingBills.length === 1 ? '1 bill' : pendingBills.length + ' bills');
  setKpi('billKpiTotal', '$' + outstandingTotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
  setKpi('billKpiTotalSub', 'approved, not yet paid');
  setKpi('billKpiPaid', '$' + paidTotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
  setKpi('billKpiPaidSub', paidBills.length + ' bills in last 30d');
  setKpi('billKpiVendors', String(uniqueVendors.size));
  setKpi('billKpiVendorsSub', 'distinct payees');

  // Update nav badge with pending count
  var badge = $('#billingBadge');
  if (badge) {
    badge.textContent = String(pendingBills.length);
    badge.style.display = pendingBills.length > 0 ? '' : 'none';
  }

  var sourceBadge = $('#billing-source-badge');
  if (sourceBadge) {
    var src = String(_lastBillSource || 'legacy').toLowerCase();
    if (src !== 'live' && src !== 'cached' && src !== 'legacy') src = 'legacy';
    sourceBadge.className = 'source-badge ' + src;
    sourceBadge.textContent = src === 'live'
      ? '● Live'
      : (src === 'cached' ? '○ Cached' : '⚠ Legacy');
  }

  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;color:var(--text-muted);padding:24px">' +
      (sourceRows.length === 0 ? 'No bills loaded — click Refresh to fetch from AppFolio' : 'No bills match current filter') + '</td></tr>';
    window._currentBillsCache = [];
    if (footer) footer.style.display = 'none';
    return;
  }

  window._currentBillsCache = filtered.slice();

  var rows = filtered.map(function(b) {
    var st = String(b.statusLabel || b.status || '—');
    var stKey = String(st).trim().toLowerCase().replace(/\s+/g, '-');
    var stLow = String(b.status || '').toLowerCase();
    var billNum = getNormalizedBillDisplayNumber(b);
    var detailId = getNormalizedBillDetailId(b);
    var amountVal = resolveBillAmountValue(b);
    var vendorText = b.vendorName || b.vendorId || '—';
    var payeeText = b.payeeUuid || b.vendorUuid || b.vendorId || '';
    if (b.vendorName && payeeText) vendorText += ' (' + payeeText + ')';
    var propertyText = String(b.propertyName || '—');
    var propertyIdText = String(b.propertyId || '').trim();
    var unitIdText = String(b.unitId || '').trim();
    var pmText = String(b.propertyManager || '').trim();
    var workOrderId = String(resolveBillWorkOrderRef(b) || '').trim();
    var woBadge = workOrderId
      ? ('<span title="Linked Work Order: ' + escapeHtml(workOrderId) + '"><i class="fas fa-wrench" style="color:var(--accent)"></i></span>')
      : '<span style="color:var(--text-muted)">—</span>';
    var groupName = resolveBillGroupName(b);
    var groupCell = groupName
      ? '<span class="tag" style="font-size:10px;padding:1px 6px">' + escapeHtml(groupName) + '</span>'
      : '<span style="color:var(--text-muted)">—</span>';
    var linkBits = [];
    if (propertyIdText) linkBits.push('PID ' + propertyIdText);
    if (unitIdText) linkBits.push('UNIT ' + unitIdText);
    if (groupName) linkBits.push('GROUP ' + groupName);
    if (pmText) linkBits.push('PM ' + pmText);
    var linkageHtml = linkBits.length
      ? ('<div style="margin-top:2px;font-size:10px;font-family:var(--font-mono);color:var(--text-muted);white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="' + escapeHtml(linkBits.join(' | ')) + '">' + escapeHtml(linkBits.join(' | ')) + '</div>')
      : '';
    return '<tr class="bill-row clickable-row" data-billdetail="' + escapeHtml(detailId) + '" tabindex="0" role="button" aria-label="Open bill ' + escapeHtml(billNum) + '">' +
      '<td style="font-family:var(--font-mono);font-size:11px;line-height:1.3">' +
        '<div style="font-weight:700;color:var(--accent)">' + escapeHtml(billNum) + '</div>' +
        '<div style="font-size:10px;color:var(--text-muted)">ID ' + escapeHtml(String(detailId || '').slice(0, 12) || '—') + '</div>' +
      '</td>' +
      '<td>' + escapeHtml(vendorText) + '</td>' +
      '<td><span style="display:inline-block;max-width:220px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="' + escapeHtml(propertyText) + '">' + escapeHtml(propertyText) + '</span>' + linkageHtml + '</td>' +
      '<td>' + groupCell + '</td>' +
      '<td style="font-family:var(--font-mono)">' + currency(amountVal, 2) + '</td>' +
      '<td>' + escapeHtml(b.date ? formatDate(b.date) : '—') + '</td>' +
      '<td><span class="status-badge status-' + escapeHtml(stKey) + ' status-' + escapeHtml(stLow.replace(/\s+/g, '-')) + '">' + escapeHtml(st) + '</span></td>' +
      '<td style="text-align:center">' + woBadge + '</td>' +
      '</tr>';
  });
  tbody.innerHTML = rows.join('');

  Array.prototype.forEach.call(tbody.querySelectorAll('tr[data-billdetail]'), function(row) {
    row.addEventListener('click', function() {
      var billId = row.getAttribute('data-billdetail');
      if (!billId) return;
      openBillKanbanCard(billId);
    });
    row.addEventListener('keydown', function(evt) {
      if (evt.key !== 'Enter' && evt.key !== ' ') return;
      evt.preventDefault();
      var billId = row.getAttribute('data-billdetail');
      if (!billId) return;
      openBillKanbanCard(billId);
    });
  });

  renderBillingPaginationControls(footer, filtered.length);
}

function openBillKanbanCard(billId) {
  var cache = Array.isArray(window._currentBillsCache) ? window._currentBillsCache : [];
  var target = String(billId || '').trim();
  var bill = cache.find(function(entry) {
    return getNormalizedBillDetailId(entry) === target;
  });
  if (!bill) {
    bill = cache.find(function(entry) {
      return String(getNormalizedBillDisplayNumber(entry) || '').trim() === target;
    });
  }
  if (!bill) {
    showToast('Bill not found in current view', { kind: 'warning' });
    return;
  }
  showBillDetailModal(getNormalizedBillDetailId(bill) || billId, bill);
}

function renderBillingPropertyScopeChip(propertyId, propertyName, unitId, unitName) {
  var chipHost = $('#billingPropertyScope');
  if (!chipHost) return;
  if (!propertyId) {
    chipHost.style.display = 'none';
    chipHost.innerHTML = '';
    return;
  }
  chipHost.style.display = '';
  if (unitId) {
    // Property + unit scope chip
    chipHost.innerHTML =
      '<span class="tag blue" style="display:inline-flex;align-items:center;gap:8px;">' +
        '<i class="fas fa-door-open"></i> ' +
        escapeHtml(propertyName || propertyId) + ' › ' + escapeHtml(unitName || unitId) +
        ' <button id="clearBillUnitScope" class="action-btn" style="padding:2px 8px;font-size:10px;">Clear Unit</button>' +
        ' <button id="clearBillPropertyScope" class="action-btn" style="padding:2px 8px;font-size:10px;">Clear All</button>' +
      '</span>';
    var clearUnitBtn = $('#clearBillUnitScope');
    if (clearUnitBtn) {
      clearUnitBtn.addEventListener('click', function() {
        window.filteredUnitId = '';
        window.filteredUnitName = '';
        renderBillsSection();
      });
    }
  } else {
    // Property-only scope chip
    chipHost.innerHTML =
      '<span class="tag blue" style="display:inline-flex;align-items:center;gap:8px;">' +
        '<i class="fas fa-building"></i> Property: ' + escapeHtml(propertyName || propertyId) +
        ' <button id="clearBillPropertyScope" class="action-btn" style="padding:2px 8px;font-size:10px;">Clear</button>' +
      '</span>';
  }
  var clearBtn = $('#clearBillPropertyScope');
  if (clearBtn) {
    clearBtn.addEventListener('click', function() {
      window.filteredPropertyId = '';
      window.filteredPropertyName = '';
      window.filteredUnitId = '';
      window.filteredUnitName = '';
      renderBillsSection();
    });
  }
}

function wireBillingFilters() {
  var filterType = document.getElementById('billing-filter-type');
  var filterInput = document.getElementById('billing-filter-input');
  var applyBtn = document.getElementById('billing-filter-apply');
  var resetBtn = document.getElementById('billing-filter-reset');
  var dateRange = document.getElementById('billing-date-range');
  var inputWrapper = document.getElementById('billing-input-wrapper');
  var statusWrap = document.getElementById('billing-status-wrapper');

  if (!filterType || !applyBtn) return;
  if (applyBtn.dataset.wired === '1') return;
  applyBtn.dataset.wired = '1';

  // Initialize Last Updated date range to 7-day default if not already set
  (function() {
    var fromEl = document.getElementById('billUpdatedFrom');
    var toEl = document.getElementById('billUpdatedTo');
    if (fromEl && !fromEl.value) {
      var d7 = new Date();
      d7.setDate(d7.getDate() - BILLS_DEFAULT_LOOKBACK_DAYS);
      fromEl.value = d7.toISOString().slice(0, 10);
    }
    if (toEl && !toEl.value) {
      toEl.value = new Date().toISOString().slice(0, 10);
    }
  })();

  var PLACEHOLDERS = {
    bills_by_vendor: 'Vendor ID (UUID)…',
    bills_by_property: 'Property ID (UUID)…',
    bills_by_wo: 'Work Order UUID or Number…',
    bills_by_wo_number: 'Work Order Number…',
    bills_by_invoice: 'Invoice Number…'
  };

  function toggleControls() {
    var v = String(filterType.value || '').trim();
    var isDueRange = v === 'bills_due_range';
    var isList = v === 'bills_list';

    if (dateRange) dateRange.style.display = isDueRange ? 'flex' : 'none';
    if (inputWrapper) inputWrapper.style.display = (isList || isDueRange) ? 'none' : 'flex';
    if (statusWrap) statusWrap.style.display = isList ? 'flex' : 'none';

    if (filterInput && PLACEHOLDERS[v]) filterInput.placeholder = PLACEHOLDERS[v];
  }

  async function applyBillingFilter() {
    var type = String(filterType.value || 'bills_list').trim();
    var value = filterInput ? String(filterInput.value || '').trim() : '';
    var status = String((document.getElementById('billing-status-filter') || {}).value || '').trim();
    var from = String((document.getElementById('billing-due-from') || {}).value || '').trim();
    var to = String((document.getElementById('billing-due-to') || {}).value || '').trim();

    _billingRouteAction = type || 'bills_list';
    _billingRouteFilterValue = value;
    _billingDueFrom = from;
    _billingDueTo = to;
    _billingRouteStatus = status;

    applyBtn.disabled = true;
    var originalText = applyBtn.innerHTML;
    applyBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Applying…';
    try {
      await loadBillingPage({ resetPage: true });
    } catch (err) {
      showToast('Billing filter failed: ' + (err.message || err), { kind: 'warning' });
    } finally {
      applyBtn.disabled = false;
      applyBtn.innerHTML = originalText;
    }
  }

  function resetBillingFilter() {
    if (filterInput) filterInput.value = '';
    var dueFromEl = document.getElementById('billing-due-from');
        var updFrom = document.getElementById('billUpdatedFrom');
        var updTo = document.getElementById('billUpdatedTo');
        if (updFrom) {
          var d7r = new Date(); d7r.setDate(d7r.getDate() - BILLS_DEFAULT_LOOKBACK_DAYS);
          updFrom.value = d7r.toISOString().slice(0, 10);
        }
        if (updTo) updTo.value = new Date().toISOString().slice(0, 10);
        var dueFromEl = document.getElementById('billing-due-from');
    var dueToEl = document.getElementById('billing-due-to');
    var statusEl = document.getElementById('billing-status-filter');
    if (dueFromEl) dueFromEl.value = '';
    if (dueToEl) dueToEl.value = '';
    if (statusEl) statusEl.value = '';
    window.filteredPropertyId = '';
    window.filteredPropertyName = '';
    window.filteredUnitId = '';
    window.filteredUnitName = '';
    window._billingPageRows = [];
    _billingServerTotal = 0;
    _billingServerTotalPages = 1;
    _billingServerPage = 1;
    _billingRouteAction = 'bills_list';
    _billingRouteFilterValue = '';
    _billingDueFrom = '';
    _billingDueTo = '';
    _billingRouteStatus = '';
    filterType.value = 'bills_list';
    _billsPage = 0;
    toggleControls();
    applyBillingFilter();
  }

  filterType.addEventListener('change', toggleControls);
  applyBtn.addEventListener('click', applyBillingFilter);
  if (resetBtn) resetBtn.addEventListener('click', resetBillingFilter);

    // Wire Last Updated date range to reload on change
    var updFromEl = document.getElementById('billUpdatedFrom');
    var updToEl = document.getElementById('billUpdatedTo');
    var dateChangeTimer = null;
    function scheduleReloadOnDateChange() {
      clearTimeout(dateChangeTimer);
      dateChangeTimer = setTimeout(function() {
        _billingListCacheRows = [];
        _billingListCacheKey = '';
        window._billingPageRows = [];
        _billingServerTotal = 0;
        _billingServerTotalPages = 1;
        _billingServerPage = 1;
        _billsPage = 0;
        loadBillingPage({ resetPage: true, forceRefresh: true });
      }, 600);
    }
    if (updFromEl) updFromEl.addEventListener('change', scheduleReloadOnDateChange);
    if (updToEl) updToEl.addEventListener('change', scheduleReloadOnDateChange);
  if (filterInput) {
    filterInput.addEventListener('keydown', function(e) {
      if (e.key === 'Enter') applyBillingFilter();
    });
  }

  document.addEventListener('groupFilterChanged', function() {
    window._currentBillsCache = [];
    _billsPage = 0;
    if (currentBillingSubtab === 'history') {
      BILL_HISTORY_PAGE = 0;
      runBillHistorySearch();
      return;
    }
    // Apply client-side group filter to cached bills instantly
    if (__CACHED_BILLS.length > 0) {
      var grp = normalizeGroupSelectionValue(getEffectiveGroupId());
      var groupUuid = getEffectiveGroupUuid(grp);
      var filteredBills = applyLocalBillFilters({ groupId: groupUuid, groupName: grp });
      _billingServerTotal = filteredBills.length;
      _billingServerTotalPages = Math.max(1, Math.ceil(_billingServerTotal / BILLS_PAGE_SIZE));
      _billingServerPage = 1;
      _billsPage = 0;
      var start = 0;
      window._billingPageRows = filteredBills.slice(start, start + BILLS_PAGE_SIZE);
      renderBillsSection(filteredBills);
      return;
    }
    // Fallback to legacy behavior if cache not yet loaded
    if (isBillingListRouteActive() && applyBillingListScopeFromCache({ resetPage: true })) {
      return;
    }
    loadBillingPage({ resetPage: true });
  });

  // Wire filter-type pill buttons
  var ftypeBtns = document.querySelectorAll('.billing-ftype-btn');
  Array.prototype.forEach.call(ftypeBtns, function(btn) {
    btn.addEventListener('click', function() {
      Array.prototype.forEach.call(ftypeBtns, function(b) { b.classList.remove('active'); });
      btn.classList.add('active');
      var ftype = String(btn.dataset.ftype || 'bills_list');
      if (filterType) { filterType.value = ftype; filterType.dispatchEvent(new Event('change')); }
    });
  });

  // Wire period preset buttons (sets hidden date inputs and triggers reload)
  var periodBtns = document.querySelectorAll('.billing-period-btn');
  var periodFrom2 = document.getElementById('billUpdatedFrom2');
  var periodTo2 = document.getElementById('billUpdatedTo2');
  function applyPeriodDays(days) {
    var toDate = new Date();
    var fromDate = new Date(toDate);
    fromDate.setDate(fromDate.getDate() - days);
    var fromStr = fromDate.toISOString().slice(0, 10);
    var toStr = toDate.toISOString().slice(0, 10);
    // Update visible custom inputs to match preset
    if (periodFrom2) periodFrom2.value = fromStr;
    if (periodTo2) periodTo2.value = toStr;
    // Update hidden inputs that fetchBills reads
    var hidFrom = document.getElementById('billUpdatedFrom');
    var hidTo = document.getElementById('billUpdatedTo');
    if (hidFrom) hidFrom.value = fromStr;
    if (hidTo) hidTo.value = toStr;
    _billingListCacheRows = [];
    _billingListCacheKey = '';
    window._billingPageRows = [];
    _billingServerTotal = 0;
    _billingServerTotalPages = 1;
    _billingServerPage = 1;
    _billsPage = 0;
    loadBillingPage({ resetPage: true, forceRefresh: true });
  }
  Array.prototype.forEach.call(periodBtns, function(btn) {
    btn.addEventListener('click', function() {
      Array.prototype.forEach.call(periodBtns, function(b) { b.classList.remove('active'); });
      btn.classList.add('active');
      applyPeriodDays(parseInt(String(btn.dataset.days || '30'), 10) || 30);
    });
  });
  // Wire custom date inputs (billUpdatedFrom2/billUpdatedTo2) to mirror to hidden inputs
  function onCustomDateChange() {
    Array.prototype.forEach.call(periodBtns, function(b) { b.classList.remove('active'); });
    var hidFrom = document.getElementById('billUpdatedFrom');
    var hidTo = document.getElementById('billUpdatedTo');
    if (hidFrom && periodFrom2) hidFrom.value = periodFrom2.value;
    if (hidTo && periodTo2) hidTo.value = periodTo2.value;
    clearTimeout(dateChangeTimer);
    dateChangeTimer = setTimeout(function() {
      _billingListCacheRows = [];
      _billingListCacheKey = '';
      window._billingPageRows = [];
      _billingServerTotal = 0;
      _billingServerTotalPages = 1;
      _billingServerPage = 1;
      _billsPage = 0;
      loadBillingPage({ resetPage: true, forceRefresh: true });
    }, 600);
  }
  if (periodFrom2) periodFrom2.addEventListener('change', onCustomDateChange);
  if (periodTo2) periodTo2.addEventListener('change', onCustomDateChange);

  // Seed 7d default on init
  applyPeriodDays(BILLS_DEFAULT_LOOKBACK_DAYS || 7);
  // Mark 7d button active
  Array.prototype.forEach.call(periodBtns, function(b) {
    if (parseInt(String(b.dataset.days || '0'), 10) === (BILLS_DEFAULT_LOOKBACK_DAYS || 7)) b.classList.add('active');
    else b.classList.remove('active');
  });

  toggleControls();
}

function buildBillLineItemsHtml(lineItems) {
  if (!Array.isArray(lineItems) || lineItems.length === 0) return '<span style="color:var(--text-muted)">No line items available</span>';
  var rows = lineItems.map(function(li) {
    var desc = li.Description || li.description || '—';
    var amt = li.Amount || li.amount || li.TotalAmount || li.total_amount || 0;
    var gl = li.GlAccountId || li.gl_account_id || li.GLAccountId || '—';
    var unit = li.UnitId || li.unit_id || '—';
    var target = resolveLineItemUuidTarget(li);
    var attrs = '';
    if (target) {
      attrs = ' class="clickable-row bill-lineitem-row" data-li-resource="' + escapeHtml(target.resource) + '" data-li-label="' + escapeHtml(target.label) + '" data-li-uuid="' + escapeHtml(target.uuid) + '" tabindex="0" role="button" aria-label="Open ' + escapeHtml(target.label) + ' UUID detail"';
    }
    return '<tr' + attrs + '>' +
      '<td style="padding:4px 6px;border-bottom:1px solid var(--border)">' + escapeHtml(String(desc)) + '</td>' +
      '<td style="padding:4px 6px;border-bottom:1px solid var(--border);text-align:right;font-family:var(--font-mono)">' + currency(amt, 2) + '</td>' +
      '<td style="padding:4px 6px;border-bottom:1px solid var(--border);font-family:var(--font-mono)">' + escapeHtml(String(gl)) + '</td>' +
      '<td style="padding:4px 6px;border-bottom:1px solid var(--border)">' + escapeHtml(String(unit)) + '</td>' +
      '</tr>';
  }).join('');
  return '<table style="width:100%;border-collapse:collapse;font-size:11px">' +
    '<thead><tr>' +
    '<th style="text-align:left;padding:4px 6px;border-bottom:1px solid var(--border)">Description</th>' +
    '<th style="text-align:right;padding:4px 6px;border-bottom:1px solid var(--border)">Amount</th>' +
    '<th style="text-align:left;padding:4px 6px;border-bottom:1px solid var(--border)">GL</th>' +
    '<th style="text-align:left;padding:4px 6px;border-bottom:1px solid var(--border)">Unit</th>' +
    '</tr></thead><tbody>' + rows + '</tbody></table>';
}

function buildBillPayloadHtml(payload) {
  var safe = payload && typeof payload === 'object' ? payload : {};
  var pretty = '';
  try { pretty = JSON.stringify(safe, null, 2); } catch (e) { pretty = String(safe || ''); }
  return '<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:8px">' +
    '<button class="action-btn" id="btnBillPayloadCopy" style="padding:3px 8px">Copy JSON</button>' +
    '<button class="action-btn" id="btnBillPayloadDownload" style="padding:3px 8px">Download JSON</button>' +
    '</div>' +
    '<pre id="billPayloadPre" style="max-height:280px;overflow:auto;background:var(--bg-secondary);border:1px solid var(--border);border-radius:8px;padding:10px;font-size:11px;line-height:1.45;font-family:var(--font-mono)">' +
    escapeHtml(pretty || '{}') +
    '</pre>';
}

async function loadBillAttachments(billId, containerEl) {
  if (!containerEl) return;
  containerEl.innerHTML = '<div style="color:var(--text-muted);font-size:11px">Loading attachments…</div>';
  try {
    var data = await proxyAction('bill_attachments', { bill_id: String(billId) });
    var list = data.attachments || data.results || [];
    if (!Array.isArray(list) || list.length === 0) {
      containerEl.innerHTML = '<div style="color:var(--text-muted);font-size:11px">No attachments on file.</div>';
      return;
    }

    containerEl.innerHTML = list.map(function(att) {
      var url = String(att.Url || att.url || att.DownloadUrl || att.download_url || '').trim();
      var name = String(att.FileName || att.file_name || att.Name || att.name || att.Id || 'Attachment');
      var ctype = String(att.ContentType || att.content_type || '').trim();
      return '<div style="display:flex;align-items:center;gap:6px;padding:4px 0;border-bottom:1px solid var(--border)">' +
        '<span style="font-size:12px">📎</span>' +
        (url
          ? ('<a href="' + escapeHtml(url) + '" target="_blank" rel="noopener" style="color:var(--accent);font-size:11px;text-decoration:none;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + escapeHtml(name) + '</a>')
          : ('<span style="flex:1;font-size:11px">' + escapeHtml(name) + '</span>')) +
        (ctype ? ('<span style="color:var(--text-muted);font-size:10px">' + escapeHtml(ctype) + '</span>') : '') +
        '</div>';
    }).join('');
  } catch (e) {
    containerEl.innerHTML = '<div style="color:var(--danger);font-size:11px">Failed to load attachments.</div>';
  }
}

async function uploadBillAttachment(billId, fileInput, statusEl, listEl) {
  if (!fileInput || !statusEl) return;
  var file = fileInput.files && fileInput.files[0];
  if (!file) {
    statusEl.style.color = 'var(--warning)';
    statusEl.textContent = 'No file selected.';
    return;
  }

  var allowed = ['application/pdf', 'image/png', 'image/jpeg', 'image/tiff'];
  var fileType = String(file.type || '').toLowerCase();
  var fileName = String(file.name || '').toLowerCase();
  var extOk = /\.(pdf|png|jpg|jpeg|tif|tiff)$/i.test(fileName);
  if (!(allowed.indexOf(fileType) !== -1 || extOk)) {
    statusEl.style.color = 'var(--danger)';
    statusEl.textContent = 'Unsupported file type. Use PDF, PNG, JPG, or TIFF.';
    return;
  }

  statusEl.style.color = 'var(--text-muted)';
  statusEl.textContent = 'Uploading…';

  try {
    var url = buildProxyActionUrl('bill_attachment_upload', { bill_id: String(billId) });
    var token = getProxyAccessToken();
    var headers = { 'Accept': 'application/json' };
    if (token) headers['Authorization'] = 'Bearer ' + token;

    var form = new FormData();
    form.append('File', file, file.name);

    var resp = await fetchWithTimeout(url, {
      method: 'POST',
      headers: headers,
      body: form
    }, 60000);

    var data = null;
    try { data = await resp.json(); } catch (eJson) { data = null; }

    if (resp.ok && data && data.ok) {
      statusEl.style.color = 'var(--success)';
      statusEl.textContent = 'Uploaded successfully.';
      fileInput.value = '';
      if (listEl) loadBillAttachments(billId, listEl);
    } else {
      var detail = data && data.detail
        ? (typeof data.detail === 'string' ? data.detail : JSON.stringify(data.detail))
        : ('HTTP ' + resp.status);
      statusEl.style.color = 'var(--danger)';
      statusEl.textContent = 'Upload failed: ' + detail.substring(0, 180);
    }
  } catch (e) {
    statusEl.style.color = 'var(--danger)';
    statusEl.textContent = 'Upload failed: ' + (e.message || e);
  }
}

async function showBillDetailModal(billId) {
  if (!billId) return;
  try {
    var b = arguments[1] || null; // optional pre-loaded bill object
    var detailFetchFailed = false;
    if (!b) {
      // Attempt cheap cache lookup before expensive API call
      var cachedRows = window._currentBillsCache || window._billingPageRows || [];
      b = cachedRows.find(function(r) { return getNormalizedBillDetailId(r) === String(billId || '').trim(); }) || null;
    }
    var hasRichDetail = !!(b && b.raw && (b.raw.LineItems || b.raw.line_items || b.raw.Remarks || b.raw.remarks || b.raw.CheckMemo));
    if (!b || !hasRichDetail) {
      try {
        var data = await proxyAction('bill_detail', { bill_id: String(billId) });
        var detailed = data.result || data.bill || null;
        if (detailed) b = detailed;
      } catch (e) {
        detailFetchFailed = true;
      }
    }
    if (!b) {
      showToast('Bill detail unavailable', { kind: 'warning' });
      return;
    }
    if (detailFetchFailed) {
      showToast('Loaded cached bill details only (live detail fetch failed)', { kind: 'warning' });
    }

    var raw = b.raw || {};
    var lineItems = b.line_items || b.lineItems || raw.LineItems || raw.line_items || [];
    var afLink = API_VHOST
      ? ('https://' + API_VHOST + '.appfolio.com/bills/' + encodeURIComponent(String(b.id || billId)))
      : '';

    // Resolve vendor info from local cache if missing
    var vendorIdResolved = String(b.vendor_id || b.vendorId || b.payee_uuid || b.payeeUuid || b.vendor_uuid || b.vendorUuid || '').trim();
    var vendorNameResolved = String(b.vendor_name || b.vendorName || '').trim() ||
      (vendorIdResolved ? resolveVendorNameFromMaps(vendorIdResolved, vendorIdResolved) : '—') || '—';

    showItemDetail('Bill — ' + String(b.bill_number || b.id || billId), [
      { section: 'Core', icon: 'fa-file-invoice-dollar' },
      { label: 'Bill ID', value: String(b.id || billId) },
      { label: 'Invoice #', value: String(b.bill_number || b.reference || '—') },
      { label: 'Status', value: String(b.status_label || b.status || '—') },
      { label: 'Amount', value: currency(b.amount || b.total_amount || 0, 2) },
      { label: 'Invoice Date', value: b.invoice_date ? formatDate(b.invoice_date) : '—' },
      { label: 'Due Date', value: b.due_date ? formatDate(b.due_date) : '—' },
      { label: 'Posting Date', value: b.posting_date ? formatDate(b.posting_date) : '—' },
      { section: 'Associations', icon: 'fa-link' },
      {
        label: 'Vendor / Payee',
        html: '<span>' + escapeHtml(vendorNameResolved) + (vendorIdResolved ? ' <span style="font-family:var(--font-mono);font-size:10px;color:var(--text-muted)">(' + escapeHtml(vendorIdResolved) + ')</span>' : '') + '</span>' +
          (vendorIdResolved
            ? ' <button class="action-btn" id="billGoToVendor" style="padding:2px 8px;font-size:10px;margin-left:8px"><i class="fas fa-external-link-alt"></i> Go to Vendor</button>'
            : '')
      },
      { label: 'Property', value: String((b.property_name || b.propertyName || '—') + ((b.property_id || b.propertyId) ? (' (' + (b.property_id || b.propertyId) + ')') : '')) },
      { label: 'Property Group', value: String(b.property_group || b.property_group_name || b.propertyGroup || b._propertyGroup || '—') },
      { label: 'Property Manager', value: String(b.pm_name || b.property_manager || raw.pm_name || raw.property_manager || raw.PropertyManager || '—') },
      { label: 'Work Order', value: (function() {
          var woRef = resolveBillWorkOrderRef(b);
          return woRef ? String(woRef) : '—';
        })() },
      { label: 'Reference', value: String(b.reference || '—') },
      { label: 'Remarks', value: String(b.remarks || '—') },
      { section: 'Line Items', icon: 'fa-list' },
      { label: 'Items', html: buildBillLineItemsHtml(lineItems) },
      { section: 'Attachments', icon: 'fa-paperclip' },
      {
        label: 'Files',
        html: '<div id="billAttachmentList" style="min-height:26px"></div>' +
          '<div style="margin-top:8px;display:flex;align-items:center;gap:8px;flex-wrap:wrap">' +
          '<input type="file" id="billAttachFile" accept=".pdf,.png,.jpg,.jpeg,.tif,.tiff" style="font-size:11px;max-width:220px" />' +
          '<button class="action-btn" id="btnUploadBillAttachment" style="padding:4px 8px">Upload</button>' +
          '<span id="billAttachStatus" style="font-size:11px;color:var(--text-muted)"></span>' +
          '</div>'
      },
      { section: 'Full Payload', icon: 'fa-code' },
      { label: 'Raw JSON', html: buildBillPayloadHtml(raw) }
    ], afLink);

    requestAnimationFrame(function() {
      var listEl = document.getElementById('billAttachmentList');
      var uploadBtn = document.getElementById('btnUploadBillAttachment');
      var fileInput = document.getElementById('billAttachFile');
      var statusEl = document.getElementById('billAttachStatus');
      var payloadCopyBtn = document.getElementById('btnBillPayloadCopy');
      var payloadDownloadBtn = document.getElementById('btnBillPayloadDownload');

      if (listEl) loadBillAttachments(billId, listEl);
        var goVendorBtn = document.getElementById('billGoToVendor');
        if (goVendorBtn) {
          goVendorBtn.addEventListener('click', function() {
            closeModal('itemDetailModal');
            // Try to open vendor modal directly if vendor is loaded
            var vendorMatch = VENDORS && VENDORS.find(function(v) { return String(v.id || '').trim() === vendorIdResolved; });
            if (vendorMatch) {
              var vendTab = document.querySelector('.nav-tab[data-tab="vendors"]');
              if (vendTab) vendTab.click();
              setTimeout(function() { openVendorModal(vendorIdResolved); }, 300);
            } else {
              // Fall back to navigating to billing-for-vendor cross-filter
              var vendTab2 = document.querySelector('.nav-tab[data-tab="vendors"]');
              if (vendTab2) vendTab2.click();
              setTimeout(function() {
                var vs = $('#vendorSearch');
                if (vs) { vs.value = vendorNameResolved || vendorIdResolved; vs.dispatchEvent(new Event('input')); }
              }, 300);
            }
          });
        }
      if (uploadBtn) {
        uploadBtn.addEventListener('click', function() {
          uploadBillAttachment(billId, fileInput, statusEl, listEl);
        });
      }
      Array.prototype.forEach.call(document.querySelectorAll('#itemDetailBody tr[data-li-uuid]'), function(row) {
        var openLineItemUuid = function(evt) {
          if (evt) evt.preventDefault();
          var uuid = String(row.getAttribute('data-li-uuid') || '').trim();
          var resource = String(row.getAttribute('data-li-resource') || '').trim();
          var label = String(row.getAttribute('data-li-label') || 'Record').trim();
          if (!uuid || !resource) return;
          showV0UuidDetailModal({ resource: resource, uuid: uuid, label: label }, null);
        };
        row.addEventListener('click', openLineItemUuid);
        row.addEventListener('keydown', function(evt) {
          if (evt.key !== 'Enter' && evt.key !== ' ') return;
          openLineItemUuid(evt);
        });
      });
      if (payloadCopyBtn && navigator && navigator.clipboard && navigator.clipboard.writeText) {
        payloadCopyBtn.addEventListener('click', async function() {
          try {
            await navigator.clipboard.writeText(JSON.stringify(raw || {}, null, 2));
            showToast('Bill payload copied', { kind: 'success' });
          } catch (e) {
            showToast('Copy failed', { kind: 'warning' });
          }
        });
      }
      if (payloadDownloadBtn) {
        payloadDownloadBtn.addEventListener('click', function() {
          try {
            var content = JSON.stringify(raw || {}, null, 2);
            var blob = new Blob([content], { type: 'application/json' });
            var url = URL.createObjectURL(blob);
            var a = document.createElement('a');
            a.href = url;
            a.download = 'bill-' + String(b.id || billId || 'payload') + '.json';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
          } catch (e) {
            showToast('Download failed', { kind: 'warning' });
          }
        });
      }
    });
  } catch (e) {
    showToast('Bill detail failed: ' + (e.message || e), { kind: 'warning' });
  }
}

var BILL_HISTORY_ROWS = [];
var BILL_HISTORY_PAGE = 0;
var BILL_HISTORY_PAGE_SIZE = 100;
var BILL_HISTORY_TOTAL = 0;
var BILL_HISTORY_TOTAL_PAGES = 1;

function renderBillHistoryPage(dateFrom, dateTo) {
  var body = document.getElementById('billHistBody');
  var footer = document.getElementById('billHistFooter');
  if (!body || !footer) return;

  if (!Array.isArray(BILL_HISTORY_ROWS) || BILL_HISTORY_ROWS.length === 0) {
    body.innerHTML = '<tr><td colspan="9" style="text-align:center;color:var(--text-muted);padding:20px">No historical bills found for selected criteria</td></tr>';
    footer.style.display = 'none';
    return;
  }

  var totalRows = Math.max(0, Number(BILL_HISTORY_TOTAL || BILL_HISTORY_ROWS.length) || 0);
  var totalPages = Math.max(1, Number(BILL_HISTORY_TOTAL_PAGES || 1) || 1);
  BILL_HISTORY_PAGE = Math.max(0, Math.min(BILL_HISTORY_PAGE, totalPages - 1));
  var start = BILL_HISTORY_PAGE * BILL_HISTORY_PAGE_SIZE;
  var end = Math.min(start + BILL_HISTORY_ROWS.length, totalRows);
  var pageRows = BILL_HISTORY_ROWS;

  var pendingHistory = pageRows.filter(function(r) {
    var statusKey = normalizeBillApprovalStatus(
      r.status || r.status_label || r.ApprovalStatus || r.approval_status || r.Status || ''
    );
    return statusKey === 'pending_approval';
  });
  var kpiPendingEl = document.getElementById('billKpiPending');
  var kpiPendingSubEl = document.getElementById('billKpiPendingSub');
  if (kpiPendingEl) kpiPendingEl.textContent = String(pendingHistory.length);
  if (kpiPendingSubEl) {
    kpiPendingSubEl.textContent = pendingHistory.length > 0
      ? (pendingHistory.length + ' pending in history result')
      : 'No pending approvals in history result';
  }

  body.innerHTML = pageRows.map(function(r) {
    var wo = String(resolveBillWorkOrderRef(r) || '').trim();
    var billId = getNormalizedBillDetailId(r);
    var vendorId = String(r.vendor_id || r.VendorId || r.payee_uuid || r.PayeeUuid || '').trim();
    var vendorName = String(r.vendor_name || r.VendorName || '').trim();
    if (!vendorName) vendorName = resolveVendorNameFromMaps(vendorId, vendorId || '—') || '—';

    var propertyId = String(r.property_id || r.PropertyId || '').trim();
    var propertyMeta = resolvePropertyMetaFromMaps(
      propertyId,
      String(r.property_name || r.PropertyName || '').trim(),
      String(r.property_group_id || r.property_group_uuid || r.PropertyGroupId || r.PropertyGroupUuid || '').trim(),
    );
    var propertyName = String(r.property_name || r.PropertyName || propertyMeta.name || propertyId || '—').trim() || '—';

    var grpResolved = String(
      r.property_group_name || r.property_group || r.PropertyGroup || propertyMeta.groupName || '',
    ).trim();
    if (!grpResolved) {
      var nGroups = _nameToGroups[(propertyName.toLowerCase())] || [];
      grpResolved = nGroups[0] || '';
    }

    var pmResolved = String(r.pm_name || r.property_manager || r.PropertyManager || '').trim();
    if (!pmResolved && grpResolved) {
      pmResolved = getRoutingPmForGroup(grpResolved);
    }

    var rowAttrs = '';
    var rowIndex = BILL_HISTORY_ROWS.indexOf(r);
    if (billId) {
      rowAttrs = ' class="clickable-row" data-billdetail="' + escapeHtml(billId) + '" data-billindex="' + escapeHtml(String(rowIndex)) + '" tabindex="0" role="button" aria-label="Open bill detail"';
    }

    return '<tr' + rowAttrs + '>' +
      '<td style="font-family:var(--font-mono);font-size:11px">' + (billId ? ('<button class="action-btn" data-billdetail="' + escapeHtml(billId) + '" style="padding:2px 8px">' + escapeHtml(getNormalizedBillDisplayNumber(r)) + '</button>') : escapeHtml(getNormalizedBillDisplayNumber(r))) + '</td>' +
      '<td>' + escapeHtml(vendorName) + '</td>' +
      '<td>' + escapeHtml(propertyName) + '</td>' +
      '<td>' + escapeHtml(grpResolved || '—') + '</td>' +
      '<td>' + escapeHtml(pmResolved || '—') + '</td>' +
      '<td style="font-family:var(--font-mono)">' + currency(r.amount || r.total_amount || r.TotalAmount || r.Paid || r.Unpaid || 0, 2) + '</td>' +
      '<td>' + escapeHtml((r.invoice_date || r.InvoiceDate || r.due_date || r.DueDate) ? formatDate(r.invoice_date || r.InvoiceDate || r.due_date || r.DueDate) : '—') + '</td>' +
      '<td>' + escapeHtml(String(r.status_label || r.status || r.ApprovalStatus || '—')) + '</td>' +
        '<td>' + (wo ? ('<button class="action-btn" data-wojump="' + escapeHtml(wo) + '" style="padding:2px 8px">#' + escapeHtml(wo) + '</button>') : '<span style="color:var(--text-muted)">—</span>') + '</td>' +
      '</tr>';
  }).join('');

  footer.style.display = '';
  footer.innerHTML =
    '<div style="display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap">' +
      '<span>History search returned ' + totalRows + ' bill(s) from ' + escapeHtml(dateFrom) + ' to ' + escapeHtml(dateTo) + '.</span>' +
      '<span style="display:inline-flex;align-items:center;gap:8px">' +
        '<button class="action-btn" id="billHistPrevPage" style="padding:3px 8px"' + (BILL_HISTORY_PAGE <= 0 ? ' disabled' : '') + '>Prev</button>' +
        '<span style="font-family:var(--font-mono);font-size:11px">Page ' + (BILL_HISTORY_PAGE + 1) + ' / ' + totalPages + '</span>' +
        '<button class="action-btn" id="billHistNextPage" style="padding:3px 8px"' + (BILL_HISTORY_PAGE >= (totalPages - 1) ? ' disabled' : '') + '>Next</button>' +
      '</span>' +
    '</div>';

  var prevBtn = document.getElementById('billHistPrevPage');
  var nextBtn = document.getElementById('billHistNextPage');
  if (prevBtn) {
    prevBtn.addEventListener('click', function() {
      if (BILL_HISTORY_PAGE <= 0) return;
      BILL_HISTORY_PAGE -= 1;
      runBillHistorySearch();
    });
  }
  if (nextBtn) {
    nextBtn.addEventListener('click', function() {
      if (BILL_HISTORY_PAGE >= (totalPages - 1)) return;
      BILL_HISTORY_PAGE += 1;
      runBillHistorySearch();
    });
  }

  Array.prototype.forEach.call(body.querySelectorAll('[data-wojump]'), function(btn) {
    btn.addEventListener('click', function(evt) {
      evt.stopPropagation();
      var woId = btn.getAttribute('data-wojump');
      if (!woId) return;
      showWODetail(woId);
    });
  });

  Array.prototype.forEach.call(body.querySelectorAll('button[data-billdetail]'), function(btn) {
    btn.addEventListener('click', function(evt) {
      evt.stopPropagation();
      var billId = btn.getAttribute('data-billdetail');
      if (!billId) return;
      var row = btn.closest('tr');
      var idx = Number(row && row.getAttribute('data-billindex'));
      var rowObj = (Number.isFinite(idx) && idx >= 0) ? BILL_HISTORY_ROWS[idx] : null;
      showBillDetailModal(billId, rowObj || null);
    });
  });

  Array.prototype.forEach.call(body.querySelectorAll('tr[data-billdetail]'), function(row) {
    row.addEventListener('click', function() {
      var billId = row.getAttribute('data-billdetail');
      if (!billId) return;
      var idx = Number(row.getAttribute('data-billindex'));
      var rowObj = (Number.isFinite(idx) && idx >= 0) ? BILL_HISTORY_ROWS[idx] : null;
      showBillDetailModal(billId, rowObj || null);
    });
    row.addEventListener('keydown', function(evt) {
      if (evt.key !== 'Enter' && evt.key !== ' ') return;
      evt.preventDefault();
      var billId = row.getAttribute('data-billdetail');
      if (!billId) return;
      var idx = Number(row.getAttribute('data-billindex'));
      var rowObj = (Number.isFinite(idx) && idx >= 0) ? BILL_HISTORY_ROWS[idx] : null;
      showBillDetailModal(billId, rowObj || null);
    });
  });
}

async function runBillHistorySearch() {
  var fromEl = document.getElementById('billHistFrom');
  var toEl = document.getElementById('billHistTo');
  var body = document.getElementById('billHistBody');
  var footer = document.getElementById('billHistFooter');
  if (!fromEl || !toEl || !body || !footer) return;

  function normalizeYmd(value) {
    var v = String(value || '').trim();
    if (!v) return '';
    if (/^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
    var d = new Date(v);
    if (isNaN(d.getTime())) return '';
    var y = d.getFullYear();
    var m = String(d.getMonth() + 1).padStart(2, '0');
    var day = String(d.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + day;
  }

  var dateFrom = normalizeYmd(fromEl.value);
  var dateTo = normalizeYmd(toEl.value);
  if (!dateFrom || !dateTo) {
    showToast('Select valid From and To dates for history search');
    return;
  }

  body.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:20px">' + loadingHtml('Searching history…') + '</td></tr>';
  footer.style.display = 'none';

  try {
    var rows = [];
    var usedDueRange = false;
    var historyTotal = 0;
    var historyPages = 1;
    var historyPage = BILL_HISTORY_PAGE + 1;
    var grpName = normalizeGroupSelectionValue(getEffectiveGroupId());
    var grpUuid = getEffectiveGroupUuid(grpName);

    if (grpUuid) {
      try {
        var dueOffset = BILL_HISTORY_PAGE * BILL_HISTORY_PAGE_SIZE;
        var dueRangeData = await proxyAction('bills_due_range', {
          group_id: String(grpUuid),
          due_from: dateFrom,
          due_to: dateTo,
          limit: String(BILL_HISTORY_PAGE_SIZE),
          offset: String(dueOffset),
          columns: BILL_ROUTE_COLUMNS
        });
        rows = Array.isArray(dueRangeData.data)
          ? dueRangeData.data
          : (Array.isArray(dueRangeData.results) ? dueRangeData.results : []);
        historyTotal = Number(dueRangeData.total || rows.length) || rows.length;
        var dueLimit = Math.max(1, Number(dueRangeData.limit || BILL_HISTORY_PAGE_SIZE) || BILL_HISTORY_PAGE_SIZE);
        historyPage = Math.max(1, Math.floor((Number(dueRangeData.offset || dueOffset) || 0) / dueLimit) + 1);
        historyPages = Math.max(1, Math.ceil(historyTotal / dueLimit));
        usedDueRange = true;
        _lastBillSource = 'cached';
      } catch (dueErr) {
        console.log('runBillHistorySearch due_range fallback: ' + (dueErr.message || dueErr));
      }
    }

    if (!usedDueRange) {
      var fromDt = new Date(dateFrom + 'T00:00:00');
      var toDt = new Date(dateTo + 'T23:59:59');
      var spanDays = 365;
      if (!isNaN(fromDt.getTime()) && !isNaN(toDt.getTime())) {
        spanDays = Math.max(30, Math.min(3650, Math.ceil((toDt - fromDt) / 86400000) + 30));
      }
      var params = {
        date_from: dateFrom,
        date_to: dateTo,
        page: String(BILL_HISTORY_PAGE + 1),
        per_page: String(BILL_HISTORY_PAGE_SIZE),
        days: String(spanDays),
        prefer_v2: 'true',
        max: '8000',
        columns: BILL_ROUTE_COLUMNS,
      };
      if (grpName) params.group_name = grpName;
      if (grpUuid) params.group_uuid = grpUuid;

      var st = ($('#billStatusFilter') ? $('#billStatusFilter').value : '').trim();
      if (st) params.status = st;

      var data = await proxyAction('bills_history', params);
      rows = Array.isArray(data.results) ? data.results : [];
      historyTotal = Number(data.total || rows.length) || rows.length;
      historyPages = Math.max(1, Number(data.total_pages || Math.ceil(historyTotal / BILL_HISTORY_PAGE_SIZE) || 1));
      historyPage = Math.max(1, Number(data.page || historyPage));
      if (typeof data.from_cache === 'boolean') {
        _lastBillSource = data.from_cache ? 'cached' : 'live';
      } else {
        _lastBillSource = 'legacy';
      }
    }

    if (rows.length && (grpName || grpUuid)) {
      var grpUuidLower = String(grpUuid || '').trim().toLowerCase();
      var grpNameLower = String(grpName || '').trim().toLowerCase();
      var hasGroupMaps = !!(Object.keys(_nameToGroups || {}).length || Object.keys(_uuidToGroups || {}).length || Object.keys(_idToGroups || {}).length);
      rows = rows.filter(function(r) {
        var rowGroupUuid = String(resolveBillGroupUuidFromRecord(r) || '').trim().toLowerCase();
        if (grpUuidLower && rowGroupUuid) return rowGroupUuid === grpUuidLower;

        var rowGroupName = String(
          r.property_group_name || r.property_group || r.PropertyGroup || resolveBillGroupName(r) || ''
        ).trim().toLowerCase();
        if (grpNameLower && rowGroupName) return rowGroupName === grpNameLower;

        var rowPropertyId = String(r.property_id || r.PropertyId || '').trim();
        var rowPropertyName = String(r.property_name || r.PropertyName || '').trim();
        if (hasGroupMaps && grpName) return isInPropertyGroup(rowPropertyId, rowPropertyName, grpName);
        return false;
      });
    }

    if (!rows.length) {
      BILL_HISTORY_ROWS = [];
      BILL_HISTORY_PAGE = 0;
      body.innerHTML = '<tr><td colspan="9" style="text-align:center;color:var(--text-muted);padding:20px">No historical bills found for selected criteria</td></tr>';
      footer.style.display = 'none';
      return;
    }

    BILL_HISTORY_ROWS = rows.slice();
    BILL_HISTORY_TOTAL = historyTotal;
    BILL_HISTORY_TOTAL_PAGES = historyPages;
    BILL_HISTORY_PAGE = Math.max(0, Math.min(historyPage - 1, historyPages - 1));
    renderBillHistoryPage(dateFrom, dateTo);
  } catch (e) {
    BILL_HISTORY_ROWS = [];
    BILL_HISTORY_PAGE = 0;
    body.innerHTML = '<tr><td colspan="9" style="text-align:center;color:var(--danger);padding:20px">History search failed: ' + escapeHtml(String(e.message || e)) + '</td></tr>';
    footer.style.display = 'none';
  }
}

function renderPayroll() {
  var period = getPayrollWeek(PAYROLL_WEEK_OFFSET);
  var rangeEl = $('#payrollRange');
  if (rangeEl) rangeEl.textContent = formatDate(period.start) + ' \u2014 ' + formatDate(period.end);

  // Sync global group filter dropdown
  var pgSel = $('#globalGroupFilter');
  if (pgSel && pgSel.value !== currentPropertyGroup) pgSel.value = currentPropertyGroup;

  var workDone = WORK_ORDERS.filter(function(wo) {
    if (!isTurnWorkDoneStatus(wo.status)) return false;
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
  if (propSub) propSub.textContent = 'properties with completed work';

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
  setDashboardKpiSkeleton(false);
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
  var pendingBillApprovals = (BILLS || []).filter(function(b) {
    if (!isInPropertyGroup(b.propertyId, b.propertyName, currentPropertyGroup)) return false;
    return getBillStatusKey(b) === 'pending_approval';
  });

  var completedTurns = TURNS.filter(function(t) {
    var endDate = t.turnEnd || t.moveIn;
    if (!endDate || !t.moveOut) return false;
    if (!isInPropertyGroup(t.propertyId, t.property, currentPropertyGroup)) return false;
    return true;
  });
  var turnDurations = completedTurns.map(function(t) {
    var start = new Date(t.moveOut);
    var end = new Date(t.turnEnd || t.moveIn);
    if (isNaN(start.getTime()) || isNaN(end.getTime())) return null;
    var days = daysBetween(start, end);
    return days >= 0 ? days : null;
  }).filter(function(v) { return v !== null; });

  var inspectionAges = INSPECTIONS.map(function(i) {
    if (!i.lastInspection) return null;
    if (!isInPropertyGroup(i.propertyId, i.propertyName, currentPropertyGroup)) return null;
    var d = new Date(i.lastInspection);
    if (isNaN(d.getTime())) return null;
    return Math.max(0, daysBetween(d, new Date()));
  }).filter(function(v) { return v !== null; });

  var woCompletionStatuses = {
    'work done': true,
    'work completed': true,
    'completed': true,
    'canceled': true,
    'cancelled': true
  };
  var woCompletionDurations = WORK_ORDERS.map(function(w) {
    if (!isInPropertyGroup(w.propertyId, w.propertyName, currentPropertyGroup)) return null;
    var statusKey = String(w.status || '').trim().toLowerCase();
    if (!woCompletionStatuses[statusKey]) return null;
    var start = getWOCreatedDate(w);
    var endCandidates = [
      w.workCompletedOn,
      w.completedOn,
      w.cancelledOn,
      w.canceledOn,
      w.cancelledAt,
      w.canceledAt,
      w.updated
    ];
    var end = null;
    for (var i = 0; i < endCandidates.length; i++) {
      var d = new Date(endCandidates[i]);
      if (!isNaN(d.getTime())) { end = d; break; }
    }
    if (!start || !end) return null;
    var delta = daysBetween(start, end);
    return delta >= 0 ? delta : null;
  }).filter(function(v) { return v !== null; });

  function averageDays(values) {
    if (!values || values.length === 0) return null;
    var sum = values.reduce(function(s, v) { return s + v; }, 0);
    return Math.round((sum / values.length) * 10) / 10;
  }

  var avgTurnCompletion = averageDays(turnDurations);
  var avgInspectionAge = averageDays(inspectionAges);
  var avgWOCompletion = averageDays(woCompletionDurations);

  var vacancyByGroup = {};
  function isVacantProperty(p) {
    var statusText = String(
      p.status || p.Status || p.occupancyStatus || p.OccupancyStatus ||
      p.rentalStatus || p.RentalStatus || p.unitStatus || p.UnitStatus || ''
    ).toLowerCase();
    if (p.isVacant === true || p.IsVacant === true || p.vacant === true) return true;
    return statusText.indexOf('vacant') !== -1 || statusText.indexOf('available') !== -1;
  }
  function propertyGroupName(p) {
    var direct = String(p.propertyGroup || p.PropertyGroup || p.group_name || '').trim();
    if (direct) return direct;
    var gid = String(p.propertyGroupId || p.PropertyGroupId || p.property_group_id || p.property_group_uuid || '').trim();
    if (!gid) return 'Unassigned';
    var grp = PROPERTY_GROUPS.find(function(g) {
      return String(g.id || g.Id || g.uuid || g.property_group_uuid || '').trim() === gid;
    });
    return String((grp && (grp.name || grp.Name)) || 'Unassigned').trim();
  }
  PROPERTIES.forEach(function(p) {
    if (!isInPropertyGroup(p.id, p.name, currentPropertyGroup)) return;
    if (!isVacantProperty(p)) return;
    var g = propertyGroupName(p) || 'Unassigned';
    vacancyByGroup[g] = (vacancyByGroup[g] || 0) + 1;
  });
  var vacancyGroups = Object.keys(vacancyByGroup);
  var vacancyTotal = vacancyGroups.reduce(function(sum, g) { return sum + vacancyByGroup[g]; }, 0);
  var topVacancyGroup = vacancyGroups.sort(function(a, b) {
    return (vacancyByGroup[b] || 0) - (vacancyByGroup[a] || 0);
  })[0] || '';

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
  if ($('#kpiPendingBills')) $('#kpiPendingBills').textContent = String(pendingBillApprovals.length);
  if ($('#kpiPendingBillsSub')) {
    $('#kpiPendingBillsSub').textContent = pendingBillApprovals.length > 0
      ? pendingBillApprovals.length + ' bill(s) awaiting approval'
      : (BILLS.length > 0 ? 'No pending approvals in scope' : 'AP bills not loaded yet');
  }

  var mgrTurn = $('#kpiMgrAvgTurnCompletion');
  var mgrTurnSub = $('#kpiMgrAvgTurnCompletionSub');
  if (mgrTurn) mgrTurn.textContent = avgTurnCompletion !== null ? avgTurnCompletion + 'd' : '—';
  if (mgrTurnSub) mgrTurnSub.textContent = turnDurations.length + ' completed turn(s) measured';

  var mgrInsp = $('#kpiMgrAvgInspectionAge');
  var mgrInspSub = $('#kpiMgrAvgInspectionAgeSub');
  if (mgrInsp) mgrInsp.textContent = avgInspectionAge !== null ? avgInspectionAge + 'd' : '—';
  if (mgrInspSub) mgrInspSub.textContent = inspectionAges.length + ' records with last inspection date';

  var mgrWo = $('#kpiMgrAvgWOCompletion');
  var mgrWoSub = $('#kpiMgrAvgWOCompletionSub');
  if (mgrWo) mgrWo.textContent = avgWOCompletion !== null ? avgWOCompletion + 'd' : '—';
  if (mgrWoSub) mgrWoSub.textContent = woCompletionDurations.length + ' closed WO(s): created -> terminal status';

  var mgrComp = $('#kpiMgrCompletedTurns');
  var mgrCompSub = $('#kpiMgrCompletedTurnsSub');
  if (mgrComp) mgrComp.textContent = completedTurns.length;
  if (mgrCompSub) mgrCompSub.textContent = completedTurns.length > 0 ? 'completed turnover records in scope' : 'No completed turnover records in scope';

  var vacancyKpi = $('#kpiVacancyGroups');
  var vacancySub = $('#kpiVacancyGroupsSub');
  if (vacancyKpi) vacancyKpi.textContent = String(vacancyTotal);
  if (vacancySub) {
    vacancySub.textContent = vacancyTotal > 0
      ? (topVacancyGroup + ': ' + vacancyByGroup[topVacancyGroup] + ' vacant')
      : 'No vacancies detected in scope';
  }
  var vacancyCard = document.getElementById('kpiVacancyCard') || (vacancyKpi ? vacancyKpi.closest('.kpi-card') : null);
  if (vacancyCard) {
    vacancyCard.classList.add('clickable-row');
    vacancyCard.setAttribute('role', 'button');
    vacancyCard.setAttribute('tabindex', '0');
    vacancyCard.setAttribute('aria-label', 'Open vacant properties');
    var openVacancies = function(evt) {
      if (evt && evt.type === 'keydown' && evt.key !== 'Enter' && evt.key !== ' ') return;
      if (evt && evt.type === 'keydown') evt.preventDefault();
      _propertiesVacancyOnly = true;
        _propertiesLocalGroup = '';
        if ($('#propertyGroupFilter')) $('#propertyGroupFilter').value = '';
        if ($('#propertySearch')) $('#propertySearch').value = '';
      _propertiesPage = 0;
      forceActiveTab('properties');
      setPropertiesSubtab('directory');
      renderPropertiesSection();
      showToast('Properties filtered to vacancies', { kind: 'info' });
    };
    vacancyCard.onclick = openVacancies;
    vacancyCard.onkeydown = openVacancies;
  }

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
    if (!isInPropertyGroup(task.propertyId, task.propertyName, currentPropertyGroup)) return;
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
    if (!webhookEventInCurrentGroup(wh)) return;
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
      detail: escapeHtml(wh.description || wh.title || wh.body || ''),
      extra: '<span style="color:var(--purple)"><i class="fas fa-satellite-dish" style="font-size:9px"></i> live</span>'
    });
  });

  // === Work Order events — fill in if tasks are sparse ===
  var sortedWOs = WORK_ORDERS.slice().sort(function(a, b) {
    return new Date(b.created || 0) - new Date(a.created || 0);
  });
  sortedWOs.slice(0, 20).forEach(function(wo) {
    if (!isInPropertyGroup(wo.propertyId, wo.propertyName, currentPropertyGroup)) return;
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
    if (!isInPropertyGroup(t.propertyId, t.property, currentPropertyGroup)) return;
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
    if (!isInPropertyGroup(r.propertyId, r.propertyName, currentPropertyGroup)) return false;
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
var currentWOVendor = '';
var currentWOProperty = '';
var currentWOAgeFilter = '';
var currentWOSort = 'oldest';
var currentWOView = 'board'; // 'board' | 'list'
var currentWOSubtab = 'active'; // active | completed | closure | followup
var currentBillingSubtab = 'queue'; // queue | history
var currentPropertiesSubtab = 'directory'; // directory | bulk
var currentActivityFilter = 'all';
var expandedWOColumn = '';
var kanbanBoardScrollState = { left: 0, top: 0 };
var woCloseAssist = { currentPage: 0, pageSize: 10 };
var showCompletedWOHistory = false;

function setWOView(viewType) {
  currentWOView = viewType === 'list' ? 'list' : 'board';
  var board = $('#kanbanBoard');
  if (board) board.classList.toggle('layout-list', currentWOView === 'list');
  var boardBtn = $('#btnWOViewBoard');
  var listBtn  = $('#btnWOViewList');
  if (boardBtn) boardBtn.classList.toggle('active', currentWOView === 'board');
  if (listBtn)  listBtn.classList.toggle('active',  currentWOView === 'list');
}

function setWOSubtab(tab) {
  var allowed = { active: true, completed: true, closure: true, followup: true };
  var target = allowed[tab] ? tab : 'active';
  currentWOSubtab = target;

  $$('[data-wo-subtab]').forEach(function(btn) {
    btn.classList.toggle('active', btn.dataset.woSubtab === target);
  });

  $$('.wo-subpanel').forEach(function(panel) {
    panel.classList.remove('active');
    panel.style.display = 'none';
  });
  var panel = $('#wo-subpanel-' + target);
  if (panel) {
    panel.classList.add('active');
    panel.style.display = '';
  }

  if (target === 'completed' && !showCompletedWOHistory) {
    showCompletedWOHistory = true;
    renderCompletedWOHistorySection();
  }
}

function setBillingSubtab(tab) {
  var allowed = { queue: true, history: true };
  var target = allowed[tab] ? tab : 'queue';
  currentBillingSubtab = target;

  $$('[data-billing-subtab]').forEach(function(btn) {
    btn.classList.toggle('active', btn.getAttribute('data-billing-subtab') === target);
  });

  ['queue', 'history'].forEach(function(name) {
    var panel = $('#billing-subpanel-' + name);
    if (!panel) return;
    var isActive = name === target;
    panel.classList.toggle('active', isActive);
    panel.style.display = isActive ? '' : 'none';
  });
}

function setPropertiesSubtab(tab) {
  var allowed = { directory: true, bulk: true };
  var target = allowed[tab] ? tab : 'directory';
  currentPropertiesSubtab = target;

  $$('[data-properties-subtab]').forEach(function(btn) {
    btn.classList.toggle('active', btn.getAttribute('data-properties-subtab') === target);
  });

  ['directory', 'bulk'].forEach(function(name) {
    var panel = $('#properties-subpanel-' + name);
    if (!panel) return;
    var isActive = name === target;
    panel.classList.toggle('active', isActive);
    panel.style.display = isActive ? '' : 'none';
  });

  if (target === 'directory') {
    renderPropertiesSection();
    return;
  }
  var scopeEl = $('#propertiesBulkScopeText');
  if (scopeEl) {
    var scope = getPropertiesScope();
    var grp = String(scope.effectiveGroup || 'All Groups');
    scopeEl.textContent = 'Current scope: ' + grp + '. Use "Open Bulk Note Composer" to apply the same note across scoped properties.';
  }
}

function getPropertiesScope(localValue) {
  var globalGroup = normalizeGroupSelectionValue(getEffectiveGroupId());
  var localGroup = normalizeGroupSelectionValue(localValue != null ? localValue : _propertiesLocalGroup);
  return {
    globalGroup: globalGroup,
    localGroup: localGroup,
    effectiveGroup: localGroup || globalGroup
  };
}

function propertyMatchesSelectedGroup(propertyRow, groupName) {
  var selected = normalizeGroupSelectionValue(groupName);
  if (!selected) return true;
  var row = propertyRow || {};
  var candidates = [
    row.propertyGroup,
    row.groupName,
    row.portfolio,
    row.property_group,
    row.PropertyGroup,
    row.PropertyGroupName,
    row.portfolio_name
  ].map(function(v) {
    return String(v || '').trim();
  }).filter(Boolean);

  if (!candidates.length) {
    return isInPropertyGroup(row.id || row.propertyId || '', row.name || '', selected);
  }

  var selectedLower = selected.toLowerCase();
  var directMatch = candidates.some(function(v) {
    return v.toLowerCase() === selectedLower;
  });
  if (directMatch) return true;

  // Some property rows carry portfolio/manager labels that differ from the
  // selected group display name. Fall back to canonical UUID/name map lookup.
  return isInPropertyGroup(row.id || row.propertyId || '', row.name || '', selected);
}
var completedWOHistoryRows = [];
var completedWOHistoryLoading = false;
var completedWOHistoryPage = 0;
var completedWOHistoryPageSize = 20;
var _lastWOScopedFilterGroup = '__init__';

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
    // Vendor dropdown
    if (currentWOVendor && String(wo.vendorName || '') !== currentWOVendor) return false;
    // Property dropdown
    if (currentWOProperty && wo.propertyName !== currentWOProperty) return false;
    // Age filter (minimum age in days)
    if (currentWOAgeFilter) {
      var createdDate = getWOCreatedDate(wo);
      if (!createdDate) return false;
      var woAgeDays = daysBetween(createdDate, new Date());
      if (woAgeDays < Number(currentWOAgeFilter)) return false;
    }
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

function rebuildWOScopedFilters() {
  var groupKey = currentPropertyGroup || '';
  var groupChanged = groupKey !== _lastWOScopedFilterGroup;
  _lastWOScopedFilterGroup = groupKey;

  if (groupChanged) {
    currentWOVendor = '';
    currentWOProperty = '';
    completedWOHistoryPage = 0;
  }

  var vendorSel = $('#woVendorFilter');
  if (vendorSel) {
    var vendors = {};
    WORK_ORDERS.forEach(function(wo) {
      if (!isInPropertyGroup(wo.propertyId, wo.propertyName, currentPropertyGroup)) return;
      var vendorName = String(wo.vendorName || '').trim();
      if (vendorName) vendors[vendorName] = true;
    });
    var vendorOptions = ['<option value="">All Vendors</option>'];
    Object.keys(vendors).sort().forEach(function(vendorName) {
      vendorOptions.push('<option value="' + escapeHtml(vendorName) + '">' + escapeHtml(vendorName) + '</option>');
    });
    vendorSel.innerHTML = vendorOptions.join('');
    if (currentWOVendor && vendors[currentWOVendor]) vendorSel.value = currentWOVendor;
    else { currentWOVendor = ''; vendorSel.value = ''; }
  }

  var propSel = $('#woPropertyFilter');
  if (propSel) {
    var props = {};
    PROPERTIES.forEach(function(p) {
      if (!isInPropertyGroup(p.id, p.name, currentPropertyGroup)) return;
      if (p.name) props[p.name] = true;
    });
    if (Object.keys(props).length === 0) {
      WORK_ORDERS.forEach(function(wo) {
        if (!isInPropertyGroup(wo.propertyId, wo.propertyName, currentPropertyGroup)) return;
        if (wo.propertyName) props[wo.propertyName] = true;
      });
    }
    var propOptions = ['<option value="">All Properties</option>'];
    Object.keys(props).sort().forEach(function(name) {
      propOptions.push('<option value="' + escapeHtml(name) + '">' + escapeHtml(name) + '</option>');
    });
    propSel.innerHTML = propOptions.join('');
    if (currentWOProperty && props[currentWOProperty]) propSel.value = currentWOProperty;
    else { currentWOProperty = ''; propSel.value = ''; }
  }
}

function getFilteredCompletedWOHistoryRows() {
  var search = $('#woSearch') ? $('#woSearch').value : '';
  return completedWOHistoryRows.filter(function(wo) {
    if (currentWOPriority && wo.priority !== currentWOPriority) return false;
    if (currentWOType && wo.type !== currentWOType) return false;
    if (currentWOVendor && String(wo.vendorName || '') !== currentWOVendor) return false;
    if (currentWOProperty && wo.propertyName !== currentWOProperty) return false;
    if (!isInPropertyGroup(wo.propertyId, wo.propertyName, currentPropertyGroup)) return false;
    if (search) {
      var s = search.toLowerCase();
      var haystack = [String(wo.id), String(wo.description || ''), String(wo.propertyName || ''), String(wo.vendorName || ''), String(wo.unit || ''), String(wo.tenant || '')].join(' ').toLowerCase();
      return haystack.indexOf(s) !== -1;
    }
    return true;
  });
}

function renderCompletedWOHistorySection() {
  var mount = $('#woCompletedHistoryMount');
  if (!mount) return;
  if (!showCompletedWOHistory) {
    mount.innerHTML = '';
    return;
  }
  if (completedWOHistoryLoading) {
    mount.innerHTML = '<div class="table-wrapper" style="margin-top:12px"><div class="table-header"><div class="table-title"><i class="fas fa-history" style="color:var(--info)"></i> Completed Work Order History</div></div><div class="wo-history-scroll" style="padding:24px;text-align:center;color:var(--text-muted)"><i class="fas fa-spinner fa-spin"></i> Loading completed history...</div></div>';
    return;
  }

  var rows = getFilteredCompletedWOHistoryRows();
  var totalPages = Math.max(1, Math.ceil(rows.length / completedWOHistoryPageSize));
  if (completedWOHistoryPage >= totalPages) completedWOHistoryPage = totalPages - 1;
  if (completedWOHistoryPage < 0) completedWOHistoryPage = 0;
  var start = completedWOHistoryPage * completedWOHistoryPageSize;
  var pageRows = rows.slice(start, start + completedWOHistoryPageSize);
  var bodyHtml = '';
  if (!rows.length) {
    bodyHtml = '<tr><td colspan="9" style="text-align:center;color:var(--text-muted)">No completed work orders match the current filters.</td></tr>';
  } else {
    pageRows.forEach(function(wo) {
      bodyHtml += '<tr>' +
        '<td>#' + escapeHtml(String(wo.id || '—')) + '</td>' +
        '<td>' + escapeHtml(String(wo.unit || '—')) + '</td>' +
        '<td>' + escapeHtml(String(wo.propertyName || '—')) + '</td>' +
        '<td>' + escapeHtml(String(wo.vendorName || '—')) + '</td>' +
        '<td>' + escapeHtml(String(wo.status || '—')) + '</td>' +
        '<td>' + escapeHtml(formatDate(wo.completedOn)) + '</td>' +
        '<td>' + escapeHtml(wo.amountBilled ? currency(wo.amountBilled) : '—') + '</td>' +
        '<td>' + escapeHtml(String(wo.type || '—')) + '</td>' +
        '<td style="max-width:220px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="' + escapeHtml(String(wo.description || '')) + '">' + escapeHtml(String(wo.description || '—')) + '</td>' +
      '</tr>';
    });
  }

  mount.innerHTML = '<div class="table-wrapper" style="margin-top:12px">' +
    '<div class="table-header"><div class="table-title"><i class="fas fa-history" style="color:var(--info)"></i> Completed Work Order History</div><div class="section-subtitle">Terminal statuses only • read-only</div></div>' +
    '<div class="wo-history-scroll" id="woHistoryScroll"><table class="data-table"><thead><tr><th>WO #</th><th>Unit</th><th>Property</th><th>Vendor</th><th>Status</th><th>Completed Date</th><th>Amount Billed</th><th>Type</th><th>Description</th></tr></thead><tbody>' + bodyHtml + '</tbody></table></div>' +
    '<div class="wo-history-pagination"><button class="action-btn" id="woHistoryPrev"' + (completedWOHistoryPage === 0 ? ' disabled' : '') + '>← Prev</button><span class="page-label">Page ' + (completedWOHistoryPage + 1) + ' of ' + totalPages + '</span><button class="action-btn" id="woHistoryNext"' + (completedWOHistoryPage >= totalPages - 1 ? ' disabled' : '') + '>Next →</button></div>' +
  '</div>';

  var scrollEl = $('#woHistoryScroll');
  var prevBtn = $('#woHistoryPrev');
  var nextBtn = $('#woHistoryNext');
  if (prevBtn) prevBtn.addEventListener('click', function() {
    if (completedWOHistoryPage <= 0) return;
    completedWOHistoryPage--;
    renderCompletedWOHistorySection();
    if (scrollEl) scrollEl.scrollTop = 0;
  });
  if (nextBtn) nextBtn.addEventListener('click', function() {
    if (completedWOHistoryPage >= totalPages - 1) return;
    completedWOHistoryPage++;
    renderCompletedWOHistorySection();
    if (scrollEl) scrollEl.scrollTop = 0;
  });
}

function getWOSortRank(wo) {
  var p = String(wo.priority || '').toLowerCase();
  if (p === 'urgent') return 0;
  if (p === 'normal') return 1;
  if (p === 'low') return 2;
  return 3;
}

function sortWorkOrders(rows) {
  return rows.slice().sort(function(a, b) {
    if (currentWOSort === 'priority') {
      var prDiff = getWOSortRank(a) - getWOSortRank(b);
      if (prDiff !== 0) return prDiff;
      var aTimeP = getWOCreatedDate(a);
      var bTimeP = getWOCreatedDate(b);
      var aMsP = aTimeP ? aTimeP.getTime() : 0;
      var bMsP = bTimeP ? bTimeP.getTime() : 0;
      return aMsP - bMsP;
    }
    var aTime = getWOCreatedDate(a);
    var bTime = getWOCreatedDate(b);
    var aMs = aTime ? aTime.getTime() : 0;
    var bMs = bTime ? bTime.getTime() : 0;
    return currentWOSort === 'newest' ? (bMs - aMs) : (aMs - bMs);
  });
}

function renderWorkOrders() {
  var board = $('#kanbanBoard');
  if (!board) return;
  rebuildWOScopedFilters();
  var filtered = sortWorkOrders(getFilteredWOs());

  // Sync global group filter dropdown
  var grpSel = $('#globalGroupFilter');
  if (grpSel && grpSel.value !== currentPropertyGroup) grpSel.value = currentPropertyGroup;
  updateGlobalGroupIndicator();

  // Vendor compliance map for cross-tab warnings
  var vendorCompliance = buildVendorComplianceMap();

  if (WORK_ORDERS.length === 0) {
    board.innerHTML = '<div style="padding:40px;text-align:center;color:var(--text-muted);width:100%"><i class="fas fa-inbox" style="font-size:36px;display:block;margin-bottom:12px;color:var(--border)"></i>No work orders loaded. Connect to API or import a cache file.</div>';
    renderCompletedWOHistorySection();
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
    var isExpanded = expandedWOColumn === col.key;
    html += '<div class="kanban-col' + (isExpanded ? ' column--expanded' : '') + '" data-column="' + escapeHtml(col.key) + '">';
    html += '<div class="kanban-col-head' + selected + '" data-status="' + escapeHtml(col.key) + '" data-column="' + escapeHtml(col.key) + '">';
    html += '<span class="kanban-col-title">' + escapeHtml(col.label) + '</span>';
    html += '<div style="display:flex;align-items:center;gap:8px"><span class="kanban-col-count">' + wos.length + '</span><span class="kanban-col-toggle">' + (isExpanded ? '▼' : '▶') + '</span></div>';
    html += '</div>';
    html += '<div class="kanban-col-body">';
    if (wos.length === 0) {
      html += '<div style="padding:12px;text-align:center;color:var(--text-muted);font-size:10px;font-family:var(--font-mono)">empty</div>';
    }
    wos.forEach(function(wo) {
      var pc = String(wo.priority || 'normal').toLowerCase();
      var flagged = isWOFlagged(wo.id);
      var ageMeta = getWOAgeMeta(wo);
      html += '<div class="kanban-card ' + pc + (ageMeta.cls ? (' ' + ageMeta.cls) : '') + (flagged ? ' flagged-card' : '') + '" data-woid="' + escapeHtml(String(wo.id)) + '">';
      html += '<div class="kc-top"><span class="kc-id">#' + escapeHtml(String(wo.id)) + (flagged ? ' <i class="fas fa-flag kc-flag"></i>' : '') + '</span><span class="kc-priority"><span class="tag ' + pc + '">' + escapeHtml(wo.priority) + '</span></span></div>';
      html += '<div class="kc-desc">' + escapeHtml(wo.description || 'No description') + '</div>';
      html += '<div class="kc-meta">';
      if (ageMeta.days !== null) html += '<span class="wo-age-pill ' + ageMeta.cls + '"><i class="fas fa-hourglass-half"></i> ' + escapeHtml(ageMeta.label) + '</span>';
      if (wo.propertyName) html += '<span><i class="fas fa-building"></i> ' + escapeHtml(wo.propertyName) + '</span>';
      if (!currentPropertyGroup && wo._propertyGroup) html += '<span class="kc-group-badge"><i class="fas fa-layer-group"></i> ' + escapeHtml(wo._propertyGroup) + '</span>';
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
    var otherExpanded = expandedWOColumn === 'Other';
    html += '<div class="kanban-col' + (otherExpanded ? ' column--expanded' : '') + '" data-column="Other"><div class="kanban-col-head" data-column="Other"><span class="kanban-col-title">Other</span><div style="display:flex;align-items:center;gap:8px"><span class="kanban-col-count">' + otherWos.length + '</span><span class="kanban-col-toggle">' + (otherExpanded ? '▼' : '▶') + '</span></div></div><div class="kanban-col-body">';
    otherWos.forEach(function(wo) {
      var pc = String(wo.priority || 'normal').toLowerCase();
      var ageMeta = getWOAgeMeta(wo);
      html += '<div class="kanban-card ' + pc + (ageMeta.cls ? (' ' + ageMeta.cls) : '') + '" data-woid="' + escapeHtml(String(wo.id)) + '"><div class="kc-top"><span class="kc-id">#' + escapeHtml(String(wo.id)) + '</span><span class="kc-priority"><span class="tag ' + pc + '">' + escapeHtml(wo.priority) + '</span></span></div>';
      html += '<div class="kc-desc">' + escapeHtml(wo.description || 'No description') + '</div>';
      html += '<div class="kc-meta"><span>' + escapeHtml(wo.status) + '</span>';
      if (ageMeta.days !== null) html += '<span class="wo-age-pill ' + ageMeta.cls + '"><i class="fas fa-hourglass-half"></i> ' + escapeHtml(ageMeta.label) + '</span>';
      if (wo.propertyName) html += '<span><i class="fas fa-building"></i> ' + escapeHtml(wo.propertyName) + '</span>';
      html += '</div></div>';
    });
    html += '</div></div>';
  }

  board.innerHTML = html;
  board.classList.toggle('has-expanded-column', !!expandedWOColumn);
  if (expandedWOColumn) {
    var expandedCol = board.querySelector('.kanban-col.column--expanded .kanban-col-body');
    if (expandedCol) expandedCol.scrollTop = 0;
  }
  $('#woBadge').textContent = filtered.length || '0';
  renderWOCloseAssist();
  renderWOFollowupQueue();
  renderCompletedWOHistorySection();
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

function normalizeWOAgeThresholds(raw) {
  var y = Math.max(1, Number(raw && raw.yellow) || 14);
  var o = Math.max(y + 1, Number(raw && raw.orange) || 30);
  var r = Math.max(o + 1, Number(raw && raw.red) || 60);
  return { yellow: y, orange: o, red: r };
}

function loadWOAgeThresholds() {
  try {
    var saved = localStorage.getItem('hm_wo_age_thresholds');
    if (!saved) return;
    var parsed = JSON.parse(saved);
    CONFIG.WO_AGE_COLOR_DAYS = normalizeWOAgeThresholds(parsed || {});
  } catch (_) {}
}

function syncWOAgeInputs() {
  var t = normalizeWOAgeThresholds(CONFIG.WO_AGE_COLOR_DAYS || {});
  var yEl = $('#woAgeYellow');
  var oEl = $('#woAgeOrange');
  var rEl = $('#woAgeRed');
  if (yEl) yEl.value = String(t.yellow);
  if (oEl) oEl.value = String(t.orange);
  if (rEl) rEl.value = String(t.red);
}

function saveWOAgeThresholdsFromInputs() {
  var yEl = $('#woAgeYellow');
  var oEl = $('#woAgeOrange');
  var rEl = $('#woAgeRed');
  var next = normalizeWOAgeThresholds({
    yellow: yEl ? yEl.value : 14,
    orange: oEl ? oEl.value : 30,
    red: rEl ? rEl.value : 60
  });
  CONFIG.WO_AGE_COLOR_DAYS = next;
  try { localStorage.setItem('hm_wo_age_thresholds', JSON.stringify(next)); } catch (_) {}
  syncWOAgeInputs();
}

function getWOAgeMeta(wo) {
  var createdDate = getWOCreatedDate(wo);
  if (!createdDate) return { days: null, cls: '', label: 'No date' };
  var days = Math.max(0, daysBetween(createdDate, new Date()));
  var t = normalizeWOAgeThresholds(CONFIG.WO_AGE_COLOR_DAYS || {});
  if (days >= t.red) return { days: days, cls: 'age-red', label: days + 'd old' };
  if (days >= t.orange) return { days: days, cls: 'age-orange', label: days + 'd old' };
  if (days >= t.yellow) return { days: days, cls: 'age-yellow', label: days + 'd old' };
  return { days: days, cls: '', label: days + 'd old' };
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
  var pagination = $('#woCloseAssistPagination');
  var container = $('#woCloseAssistContainer');
  if (!body || !summary || !pagination || !container) return;

  var rows = computeWOCloseAssistRows();
  var hi = rows.filter(function(r) { return r.confidence === 'High'; }).length;
  var med = rows.filter(function(r) { return r.confidence === 'Medium'; }).length;
  summary.textContent = rows.length + ' candidate(s) • High: ' + hi + ' • Medium: ' + med +
    (BILLS.length > 0 ? ' • AP bills: ' + BILLS.length : ' • AP bills: none loaded — click Refresh AP for evidence');

  if (rows.length === 0) {
    var noRowsMsg = (WORK_ORDERS.length === 0)
      ? 'Work orders not loaded — refresh to load data.'
      : (BILLS.length === 0)
        ? 'No aged open WOs match current filter. ℹ️ AP bills not loaded — click “Refresh AP” above for vendor-match evidence.'
        : 'No candidates for current age/filter window.';
    body.innerHTML = '<tr><td colspan="9" style="text-align:center;color:var(--text-muted)">' + noRowsMsg + '</td></tr>';
    pagination.innerHTML = '';
    woCloseAssist.currentPage = 0;
    return;
  }

  var totalPages = Math.max(1, Math.ceil(rows.length / woCloseAssist.pageSize));
  if (woCloseAssist.currentPage >= totalPages) woCloseAssist.currentPage = totalPages - 1;
  if (woCloseAssist.currentPage < 0) woCloseAssist.currentPage = 0;
  var start = woCloseAssist.currentPage * woCloseAssist.pageSize;
  var pageItems = rows.slice(start, start + woCloseAssist.pageSize);

  var html = '';
  pageItems.forEach(function(r) {
    var wo = r.wo;
    var confColor = r.confidence === 'High' ? 'var(--danger)' : (r.confidence === 'Medium' ? 'var(--warning)' : 'var(--text-muted)');
    var leadBillId = r.billMatches.length > 0 ? String(r.billMatches[0].id || '').trim() : '';
    var leadBillLabel = r.billMatches.length > 0
      ? String(r.billMatches[0].billNumber || r.billMatches[0].id || 'Bill')
      : '';
    var billEvidence = '—';
    if (r.billMatches.length > 0) {
      var total = r.billMatches.reduce(function(sum, b) {
        return sum + (parseFloat(String(b.amount || 0).replace(/[^0-9.\-]/g, '')) || 0);
      }, 0);
      var lead = r.billMatches.slice(0, 2).map(function(b) {
        var bid = String(b.id || '').trim();
        var label = String(b.billNumber || b.id || 'bill');
        if (!bid) return '<span class="tag" style="font-family:var(--font-mono)">' + escapeHtml(label) + '</span>';
        return '<button class="action-btn" data-billdetail="' + escapeHtml(bid) + '" style="padding:2px 6px;font-size:10px;font-family:var(--font-mono)" title="Open bill detail">' +
          escapeHtml(label) + '</button>';
      }).join(' ');
      billEvidence = '<div style="display:flex;flex-direction:column;gap:4px">' +
        '<div><span style="font-weight:600;color:var(--success)">' + r.billMatches.length + ' match' + (r.billMatches.length === 1 ? '' : 'es') + '</span> ' +
        '<span style="color:var(--text-muted)">(' + currency(total) + ')</span></div>' +
        '<div style="display:flex;gap:4px;flex-wrap:wrap">' + lead +
        (r.billMatches.length > 2 ? ('<span style="font-size:10px;color:var(--text-muted)">+' + (r.billMatches.length - 2) + ' more</span>') : '') +
        '</div></div>';
    }
    html += '<tr>';
    html += '<td><button class="action-btn" data-woid="' + escapeHtml(String(wo.id)) + '" style="padding:2px 8px">#' + escapeHtml(String(wo.id)) + '</button></td>';
    html += '<td>' + escapeHtml(wo.propertyName || '—') + '</td>';
    html += '<td>' + escapeHtml(wo.unit || '—') + '</td>';
    html += '<td>' + r.ageDays + 'd <span style="color:var(--text-muted)">(' + r.bucket + ')</span></td>';
    html += '<td>' + escapeHtml(wo.assignedUser || wo.vendorName || '—') + '</td>';
    html += '<td>' + billEvidence + '</td>';
    html += '<td><span style="font-weight:700;color:' + confColor + '">' + r.confidence + '</span></td>';
    html += '<td>' + escapeHtml(r.suggestion) + '</td>';
    html += '<td style="white-space:nowrap">' +
      '<button class="action-btn" data-woopen="' + escapeHtml(String(wo.id)) + '" style="padding:2px 8px;margin-right:4px" title="Open WO detail">WO</button>' +
      (leadBillId
        ? ('<button class="action-btn" data-billopen="' + escapeHtml(leadBillId) + '" style="padding:2px 8px" title="Open lead bill ' + escapeHtml(leadBillLabel) + '">Bill</button>')
        : '<button class="action-btn" disabled style="padding:2px 8px;opacity:0.5" title="No AP match">Bill</button>') +
      '</td>';
    html += '</tr>';
  });
  body.innerHTML = html;
  pagination.innerHTML = '<button class="action-btn" id="woCloseAssistPrev"' + (woCloseAssist.currentPage === 0 ? ' disabled' : '') + '>← Prev</button>' +
    '<span class="page-label">Page ' + (woCloseAssist.currentPage + 1) + ' of ' + totalPages + '</span>' +
    '<button class="action-btn" id="woCloseAssistNext"' + (woCloseAssist.currentPage >= totalPages - 1 ? ' disabled' : '') + '>Next →</button>';

  Array.prototype.forEach.call(body.querySelectorAll('button[data-woid]'), function(btn) {
    btn.addEventListener('click', function() {
      showWODetail(btn.getAttribute('data-woid'));
    });
  });
  Array.prototype.forEach.call(body.querySelectorAll('button[data-billdetail]'), function(btn) {
    btn.addEventListener('click', function() {
      var billId = btn.getAttribute('data-billdetail');
      if (!billId) return;
      showBillDetailModal(billId);
    });
  });
  Array.prototype.forEach.call(body.querySelectorAll('button[data-woopen]'), function(btn) {
    btn.addEventListener('click', function() {
      var woId = btn.getAttribute('data-woopen');
      if (!woId) return;
      showWODetail(woId);
    });
  });
  Array.prototype.forEach.call(body.querySelectorAll('button[data-billopen]'), function(btn) {
    btn.addEventListener('click', function() {
      var billId = btn.getAttribute('data-billopen');
      if (!billId) return;
      showBillDetailModal(billId);
    });
  });
  var prevBtn = document.getElementById('woCloseAssistPrev');
  var nextBtn = document.getElementById('woCloseAssistNext');
  if (prevBtn) prevBtn.addEventListener('click', function() {
    if (woCloseAssist.currentPage <= 0) return;
    woCloseAssist.currentPage--;
    renderWOCloseAssist();
    container.scrollTop = 0;
  });
  if (nextBtn) nextBtn.addEventListener('click', function() {
    if (woCloseAssist.currentPage >= totalPages - 1) return;
    woCloseAssist.currentPage++;
    renderWOCloseAssist();
    container.scrollTop = 0;
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
  var ref = String(id || '').trim().replace(/^#/, '');
  var refLower = ref.toLowerCase();
  var wo = WORK_ORDERS.find(function(w) {
    var idText = String(w.id || '').trim();
    var uuidText = String(w.uuid || '').trim();
    if (idText === ref || uuidText === ref) return true;
    if (idText.toLowerCase() === refLower || uuidText.toLowerCase() === refLower) return true;
    return false;
  });
  if (!wo) return;
  var woDbUuid = resolveWODbUuid(wo);
  var woRefForApi = woDbUuid || (isUuidString(wo.uuid || '') ? String(wo.uuid) : '') || String(wo.id || '');
  CURRENT_WO_MODAL = { woId: String(wo.id), woDbUuid: isUuidString(woRefForApi) ? woRefForApi : '' };
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
  if (isTerminalWOStatus(wo.status)) {
    html += '<div class="detail-row"><div class="detail-row-label">Amount Billed</div><div class="detail-row-value" id="detailBilledAmount"><i class="fas fa-spinner fa-spin" style="font-size:10px"></i> Loading billed amount...</div></div>';
  }
  html += '<div class="detail-row"><div class="detail-row-label">Created</div><div class="detail-row-value" id="detailCreatedAt">' + escapeHtml(wo.created ? formatNoteDateTime(wo.created) : '\u2014') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Scheduled</div><div class="detail-row-value">' + formatDate(wo.scheduledStart) + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Attachment Types</div><div class="detail-row-value" id="detailAttachmentTypes">Loading…</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Permission to Enter</div><div class="detail-row-value" id="detailPermission">' + escapeHtml(wo.permissionToEnter || '\u2014') + '</div></div>';
  html += '</div>';
  html += '<div class="form-group" style="margin-top:10px"><label class="form-label">Description</label><textarea class="form-textarea" id="detailDesc">' + escapeHtml(wo.description) + '</textarea></div>';
  html += '</div>';

  html += '<div class="detail-section"><div class="detail-section-title"><i class="fas fa-clipboard-list"></i> Scope &amp; Instructions</div>';
  html += '<div class="detail-grid">';
  html += '<div class="detail-row"><div class="detail-row-label">Job Description</div><div class="detail-row-value" id="detailJobDescription">' + escapeHtml(wo.description || '\u2014') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Vendor Instructions</div><div class="detail-row-value" id="detailVendorInstructions">—</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Estimates</div><div class="detail-row-value" id="detailEstimates">—</div></div>';
  html += '</div></div>';

  // -- Property & Unit Section --
  html += '<div class="detail-section"><div class="detail-section-title"><i class="fas fa-building"></i> Property &amp; Unit</div>';
  html += '<div class="detail-grid">';
  html += '<div class="detail-row"><div class="detail-row-label">Property</div><div class="detail-row-value">' + escapeHtml(wo.propertyName) + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Unit</div><div class="detail-row-value">' + escapeHtml(wo.unit || '\u2014') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Address</div><div class="detail-row-value">' + escapeHtml(wo.propertyAddress || '\u2014') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Site Manager</div><div class="detail-row-value" id="detailSiteMgr">—</div></div>';
  html += '</div></div>';

  // -- Tenant Section --
  html += '<div class="detail-section"><div class="detail-section-title"><i class="fas fa-user"></i> Tenant</div>';
  html += '<div class="detail-grid">';
  html += '<div class="detail-row"><div class="detail-row-label">Name</div><div class="detail-row-value">' + escapeHtml(wo.tenant || '\u2014') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Email</div><div class="detail-row-value" id="detailTenantEmail">' + escapeHtml(wo.tenantEmail || '\u2014') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Phone</div><div class="detail-row-value" id="detailTenantPhone">' + escapeHtml(wo.tenantPhone || '\u2014') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Assigned To</div><div class="detail-row-value" id="detailAssignedTo">' + escapeHtml(wo.assignedUser || '\u2014') + '</div></div>';
  html += '</div></div>';

  // -- Vendor Section --
  html += '<div class="detail-section"><div class="detail-section-title"><i class="fas fa-hard-hat"></i> Vendor</div>';
  html += '<div class="detail-grid">';
  html += '<div class="detail-row"><div class="detail-row-label">Vendor</div><div class="detail-row-value">' + escapeHtml(wo.vendorName || 'Unassigned') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Vendor ID</div><div class="detail-row-value">' + escapeHtml(wo.vendorId || '\u2014') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Trade</div><div class="detail-row-value">' + escapeHtml(wo.vendorTrade || '\u2014') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Created By</div><div class="detail-row-value">' + escapeHtml(wo.createdBy || '\u2014') + '</div></div>';
  html += '<div class="detail-row"><div class="detail-row-label">Maint. Limit</div><div class="detail-row-value">' + escapeHtml(wo.maintenanceLimit || '\u2014') + '</div></div>';
  html += '</div></div>';

  html += '<div class="detail-section"><div class="detail-section-title"><i class="fas fa-route"></i> Related Views</div>';
  html += '<div style="display:flex;gap:8px;flex-wrap:wrap">';
  if (wo.vendorName) html += '<button class="action-btn" id="detailOpenVendorWOs" style="padding:6px 10px">Vendor WOs</button>';
  if (wo.vendorId || wo.vendorName) html += '<button class="action-btn" id="detailOpenVendorBills" style="padding:6px 10px">Vendor Bills</button>';
  html += '<button class="action-btn" id="detailOpenWorkOrderBills" style="padding:6px 10px">Bills For This WO</button>';
  html += '</div></div>';

  html += '<div class="detail-section"><div class="detail-section-title"><i class="fas fa-paperclip"></i> Attachments</div>';
  html += '<div id="detailAttachmentList" style="min-height:26px"><div style="color:var(--text-muted);font-size:11px"><i class="fas fa-spinner fa-spin"></i> Loading attachments…</div></div></div>';

  // -- Notes Section (async load) --
  html += '<div class="detail-section"><div class="detail-section-title"><i class="fas fa-sticky-note"></i> Notes</div>';
  html += '<div class="note-list" id="detailNotesList"><div style="text-align:center;padding:10px;color:var(--text-muted)"><i class="fas fa-spinner fa-spin"></i> Loading notes...</div></div></div>';

  // -- Add Note --
  html += '<div class="detail-section"><div class="detail-section-title"><i class="fas fa-plus-circle"></i> Add Note</div>';
  html += '<textarea class="form-textarea" placeholder="Type a note\u2026" id="detailNote"></textarea>';
  html += '<div id="detailNoteError" style="display:none;margin-top:8px;font-size:12px;color:var(--error-text);background:var(--error-bg);border:1px solid var(--error-border);border-radius:8px;padding:8px 10px"></div></div>';

  $('#woModalBody').innerHTML = html;

  function setDetailNoteError(message) {
    var errEl = document.getElementById('detailNoteError');
    if (!errEl) return;
    if (!message) {
      errEl.style.display = 'none';
      errEl.textContent = '';
      return;
    }
    errEl.textContent = message;
    errEl.style.display = 'block';
  }

  // Cache-only site manager lookup (no live API call)
  if (wo.propertyId) {
    var prop = PROPERTIES.find(function(p) { return p.id === wo.propertyId || String(p.id) === String(wo.propertyId); });
    var smEl = document.getElementById('detailSiteMgr');
    if (smEl) smEl.textContent = (prop && prop.siteManager) ? prop.siteManager : '\u2014';
  } else {
    var smEl = document.getElementById('detailSiteMgr');
    if (smEl) smEl.textContent = '\u2014';
  }

  // Async: fetch notes (use resolved DB API UUID for /api/v0/ endpoint)
  if (isUuidString(woRefForApi)) delete WO_DETAIL_CACHE['notes_' + woRefForApi];
  fetchWONotes(woRefForApi, wo).then(function(notes) {
    renderWONotesList(notes);
  });
  loadWOAttachments(woRefForApi);
  loadWODetailExtras(wo, isUuidString(woRefForApi) ? woRefForApi : woDbUuid);

  var vendorWOBtn = document.getElementById('detailOpenVendorWOs');
  if (vendorWOBtn) vendorWOBtn.onclick = function() { navigateToOpenWOsForVendor(wo.vendorName || ''); closeModal('woModal'); };
  var vendorBillsBtn = document.getElementById('detailOpenVendorBills');
  if (vendorBillsBtn) vendorBillsBtn.onclick = function() { navigateToBillsForVendor(wo.vendorId || wo.vendorName || ''); closeModal('woModal'); };
  var woBillsBtn = document.getElementById('detailOpenWorkOrderBills');
  if (woBillsBtn) woBillsBtn.onclick = function() { navigateToBillsForWorkOrder(woRefForApi || wo.id || ''); closeModal('woModal'); };

  if (isTerminalWOStatus(wo.status)) {
    fetchWOBilledAmount(wo.id).then(function(data) {
      var billedEl = document.getElementById('detailBilledAmount');
      if (!billedEl) return;
      var billed = Number((data && data.total_billed) || 0);
      if (!isFinite(billed)) billed = 0;
      billedEl.textContent = '$' + billed.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }).catch(function(err) {
      var billedEl = document.getElementById('detailBilledAmount');
      if (!billedEl) return;
      billedEl.textContent = 'Unavailable';
      billedEl.title = (err && err.message) ? err.message : 'Unable to fetch billed amount';
    });
  }

  $('#woModalSave').onclick = async function() {
    if (isReadOnlyAccessMode()) {
      setDetailNoteError('Read-only access mode: updates are disabled.');
      return;
    }
    setDetailNoteError('');
    var newStatus = $('#detailStatus').value;
    var newPriority = $('#detailPriority').value;
    var noteInput = ($('#detailNote') && typeof $('#detailNote').value === 'string') ? $('#detailNote').value : '';
    var note = noteInput.trim();
    var statusChanged = newStatus !== wo.status || newPriority !== wo.priority;
    if (!statusChanged && noteInput.length > 0 && !note) {
      setDetailNoteError('Please enter a non-empty note.');
      return;
    }

    try {
      if (statusChanged) {
        if (!woDbUuid) { showToast('Cannot update: no DB API UUID for this WO'); return; }
        await apiFetch('/api/v0/work_orders/' + woDbUuid, {
          method: 'PATCH',
          body: JSON.stringify({ Status: newStatus, Priority: newPriority })
        });
        wo.status = newStatus;
        wo.priority = newPriority;
      }
      if (note) {
        if (!woDbUuid) {
          setDetailNoteError('Cannot add note: no DB API UUID for this work order.');
          return;
        }
        var noteResp = await postWONoteViaProxy(woDbUuid, note);
        if (!noteResp.ok) {
          setDetailNoteError('Add note failed (HTTP ' + noteResp.status + '): ' + noteResp.message);
          return;
        }
        if ($('#detailNote')) $('#detailNote').value = '';
        // bypassProxyCache=true so the new note appears immediately (skips 5-min Turso TTL)
        refreshCurrentWONotes(woDbUuid, true);
      }
      renderWorkOrders();
      renderDashboardKPIs();
      if (!note) closeModal('woModal');
      showToast(note ? 'Note added to #' + wo.id : 'Updated #' + wo.id + ' successfully');
      await saveAllToCache();
    } catch (err) {
      var msg = (err && err.message) ? err.message : 'Request failed';
      setDetailNoteError(msg);
    }
  };
  if (isReadOnlyAccessMode()) {
    var saveBtn = $('#woModalSave');
    if (saveBtn) {
      saveBtn.disabled = true;
      saveBtn.title = 'Read-only access';
      saveBtn.textContent = 'Read-Only';
    }
  } else {
    var saveBtn2 = $('#woModalSave');
    if (saveBtn2) {
      saveBtn2.disabled = false;
      saveBtn2.title = '';
      saveBtn2.textContent = 'Save Changes';
    }
  }
  openModal('woModal');
}

/* =================================================================
   TURN PIPELINE — Stage Tracking Engine
   Correlates: Turns (v2) + Work Orders + Inspections + Webhook Events
   Stages: MO → INS → WO → REQ → EST → ASN → DONE
   ================================================================= */
var TURN_RECORDS = []; // persisted stage overrides from proxy blob
var TURN_PIPE_DATA = []; // computed pipeline entries
var CLOSED_TURNS = new Set(); // turn IDs marked as closed (stored in SQL)
var OPEN_TURN_DETAIL_ID = '';
var DASH_TURN_PM_FILTER = (function(){ try { return localStorage.getItem('flr_turn_dash_filter') || ''; } catch(e){ return ''; } })();
var DASH_TURN_VIEW_MODE = 'cards';
var DASH_TURN_PAGE = 0;
var DASH_TURN_ROTATOR = null;
var DASH_TURN_ROTATE_MS = 8000;
var DASH_TURN_LAST_SYNC_AT = '';
var DASH_TURN_SYNC_TIMER = null;
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

// Fetch closed turns from proxy and populate CLOSED_TURNS set
async function loadClosedTurns() {
  if (!API_PROXY) return;
  try {
    var data = await proxyAction('closed_turns');
    if (data && data.ok && Array.isArray(data.results)) {
      CLOSED_TURNS = new Set(data.results.map(function(r) { return r.turn_id; }));
    }
  } catch (err) {
    console.warn('loadClosedTurns failed (non-fatal):', err.message || err);
  }
}

async function fetchUnitTurnHistory() {
  if (!API_PROXY) return;
  try {
    var data = await proxyAction('unit_turns_history', { days: '540', limit: '300' });
    if (data && data.ok && Array.isArray(data.results)) {
      UNIT_TURN_HISTORY = data.results;
    }
  } catch (err) {
    console.warn('fetchUnitTurnHistory failed (non-fatal):', err.message || err);
  }
}

function renderTurnHistoryPanel() {
  var body = document.getElementById('turnHistoryBody');
  if (!body) return;
  if (!UNIT_TURN_HISTORY || UNIT_TURN_HISTORY.length === 0) {
    body.innerHTML = '<div class="dbadmin-msg">No closed turn history yet</div>';
    return;
  }
  var cols = ['closed_at', 'close_source', 'close_reason', 'property_name', 'unit_name', 'move_out_date', 'move_in_date', 'site_manager', 'tracking_code'];
  var thead = '<thead><tr>' + cols.map(function(c) {
    return '<th>' + escapeHtml(c.replace(/_/g, ' '));
  }).join('') + '</tr></thead>';
  var tbody = '<tbody>' + UNIT_TURN_HISTORY.map(function(r) {
    return '<tr>' + cols.map(function(c) {
      var v = r[c];
      if (!v) return '<td class="db-null">—</td>';
      if (c.indexOf('date') !== -1 || c === 'closed_at') return '<td title="' + escapeHtml(String(v)) + '">' + escapeHtml(formatDate(v)) + '</td>';
      return '<td title="' + escapeHtml(String(v)) + '">' + escapeHtml(String(v)) + '</td>';
    }).join('') + '</tr>';
  }).join('') + '</tbody>';
  body.innerHTML = '<table>' + thead + tbody + '</table>';
}

// Close a turn — POST to proxy, then re-render
async function closeTurn(turnId) {
  if (!turnId) return;
  if (!await hmConfirm('Close this turn? It will be hidden from the active board.', { title: 'Close Turn', okLabel: 'Close Turn' })) return;
  try {
    var entry = TURN_PIPE_DATA.find(function(p) { return p.id === turnId; }) || null;
    var r = await proxyPost('closed_turns', {
      turn_id: turnId,
      close_reason: 'manual_close',
      close_source: 'manual',
      closed_by: 'ui',
      property_id: entry ? String(entry.propertyId || '') : '',
      property_name: entry ? String(entry.property || '') : '',
      unit_id: entry ? String(entry.unitId || '') : '',
      unit_name: entry ? String(entry.unit || '') : '',
      move_out_date: entry ? String(entry.moveOut || '') : '',
      move_in_date: entry ? String(entry.expectedMoveIn || '') : ''
    });
    if (r && r.ok) {
      CLOSED_TURNS.add(turnId);
      renderTurnBoard();
      fetchUnitTurnHistory().then(function() { renderTurnHistoryPanel(); }).catch(function() {});
      showToast('Turn closed and removed from active board');
    } else {
      showToast('Could not close turn: ' + (r && r.error ? r.error : 'Unknown error'));
    }
  } catch (e) {
    showToast('Close turn failed: ' + (e.message || e));
  }
}

// Save a turn record stage to proxy
async function saveTurnRecordStage(turnId, stage, stageData) {
  try {
    await proxyPost('turn_record_stage', { id: turnId, stage: stage, data: stageData });
  } catch (err) {
    console.log('Save stage error: ' + (err.message || err));
  }
}

// Save full turn record to proxy
async function saveTurnRecord(record) {
  try {
    await proxyPost('turn_records', record);
  } catch (err) {
    console.log('Save turn record error: ' + (err.message || err));
  }
}

function computeTurnConfidence(turnData, stages, matchingWOs, matchingInsp) {
  var score = 0;
  if (turnData) score += 45;
  if (stages && stages.moveout && stages.moveout.done) score += 15;
  if (matchingInsp && stages && stages.inspection && stages.inspection.done) score += 25;
  if (Array.isArray(matchingWOs) && matchingWOs.length > 0) score += Math.min(15, matchingWOs.length * 5);
  if (stages && stages.est_requested && stages.est_requested.done) score += 5;
  if (stages && stages.est_received && stages.est_received.done) score += 5;
  if (stages && stages.work_done && stages.work_done.done) score += 5;
  if (score > 100) score = 100;
  var label = score >= 80 ? 'high' : score >= 50 ? 'medium' : 'low';
  return { score: score, label: label };
}

function turnMilestonesFromPipeline(p) {
  return {
    moveout: { date: p.moveOut || '', source: p.turn ? 'reports' : 'inferred' },
    inspection: { date: p.stages && p.stages.inspection ? p.stages.inspection.date || '' : '', source: (p.stages && p.stages.inspection && p.stages.inspection.manual) ? 'manual' : 'inferred' },
    wo_created: { date: p.stages && p.stages.wo_created ? p.stages.wo_created.date || '' : '', source: 'inferred' },
    est_requested: { date: p.stages && p.stages.est_requested ? p.stages.est_requested.date || '' : '', source: 'inferred' },
    est_received: { date: p.stages && p.stages.est_received ? p.stages.est_received.date || '' : '', source: 'inferred' },
    movein: { date: p.expectedMoveIn || '', source: p.expectedMoveIn ? 'reports' : 'manual' }
  };
}

function buildTurnTrackerSyncPayload() {
  return TURN_PIPE_DATA.map(function(p) {
    var shouldAutoClose = !!(p.expectedMoveIn && (new Date(p.expectedMoveIn)).getTime() <= Date.now());
    var status = p.isClosed || shouldAutoClose
      ? 'closed'
      : p.isCompleted
      ? 'completed'
      : p.isConfirmed
      ? 'active'
      : 'on_radar';
    return {
      turn_key: p.id,
      unit_turn_id: (p.turn && p.turn.unitTurnId) || p.registeredUnitTurnId || '',
      unit_id: p.unitId || '',
      property_id: p.propertyId || '',
      unit_name: p.unit || '',
      property_name: p.property || '',
      move_out_date: p.moveOut || '',
      move_in_date: p.expectedMoveIn || '',
      inspection_date: (p.stages && p.stages.inspection && p.stages.inspection.date) || '',
      first_wo_date: (p.stages && p.stages.wo_created && p.stages.wo_created.date) || '',
      estimate_requested_date: (p.stages && p.stages.est_requested && p.stages.est_requested.date) || '',
      estimate_received_date: (p.stages && p.stages.est_received && p.stages.est_received.date) || '',
      status: status,
      closed_at: (status === 'closed' || status === 'completed') ? new Date().toISOString() : '',
      confidence_score: p.confidenceScore || 0,
      confidence_label: p.confidenceLabel || 'low',
      site_manager: p.siteManager || '',
      source_flags: {
        from_turn_report: !!p.turn,
        from_inspection: !!(p.matchingInsp && p.stages && p.stages.inspection && p.stages.inspection.done),
        from_work_orders: !!(p.matchingWOs && p.matchingWOs.length),
        is_on_radar: !!p.isOnRadar,
        auto_closed_by_move_in: shouldAutoClose
      },
      metadata: {
        close_reason: p.isClosed ? 'manual_close' : (shouldAutoClose ? 'move_in_detected' : ''),
        close_source: p.isClosed ? 'manual' : (shouldAutoClose ? 'system_move_in' : 'pipeline')
      },
      milestones: turnMilestonesFromPipeline(p),
      replace_work_orders: true,
      work_orders: (p.matchingWOs || []).map(function(w) {
        return {
          wo_id: String(w.id || ''),
          wo_db_uuid: w.dbApiId || '',
          source: w.source || 'inferred',
          status: w.status || '',
          created_at: w.created || ''
        };
      })
    };
  });
}

async function syncAutoClosedTurns(records) {
  if (!API_PROXY || _turnAutoCloseSyncInFlight || !Array.isArray(records) || records.length === 0) return;
  var hash = JSON.stringify(records.map(function(r) {
    return [r.turn_key, r.move_in_date, r.move_out_date, r.property_id, r.unit_id].join('|');
  }));
  if (hash === _lastTurnAutoCloseHash) return;
  _turnAutoCloseSyncInFlight = true;
  try {
    await proxyPost('unit_turns_sync', { records: records });
    _lastTurnAutoCloseHash = hash;
  } catch (err) {
    console.log('syncAutoClosedTurns error: ' + (err.message || err));
  } finally {
    _turnAutoCloseSyncInFlight = false;
  }
}

async function syncTurnTrackerFromPipeline() {
  if (!API_PROXY || _turnTrackerSyncInFlight || TURN_PIPE_DATA.length === 0) return;
  var payload = { records: buildTurnTrackerSyncPayload() };
  var hash = JSON.stringify(payload.records.map(function(r) {
    return [r.turn_key, r.status, r.confidence_score, r.work_orders.length, r.move_out_date, r.move_in_date].join('|');
  }));
  if (hash === _lastTurnTrackerSyncHash) return;
  _turnTrackerSyncInFlight = true;
  try {
    await proxyPost('unit_turns_sync', payload);
    _lastTurnTrackerSyncHash = hash;
  } catch (err) {
    console.log('syncTurnTrackerFromPipeline error: ' + (err.message || err));
  } finally {
    _turnTrackerSyncInFlight = false;
  }
}

async function linkTurnWorkOrder(turnKey, woId) {
  if (!turnKey || !woId) return;
  if (isTurnActionLocked(turnKey)) {
    throw new Error('Turn is terminal/closed; link actions are locked');
  }
  // WORK_ORDERS (Reports v2): id = WO number, uuid = UUID (work_order_id)
  // TURN_WORK_ORDERS (DB API v0): id = UUID, woNumber = WO number
  var woStr = String(woId).replace(/^\s*#/, '').trim();
  var wo = WORK_ORDERS.find(function(w) { return String(w.id) === woStr; }) ||
    TURN_WORK_ORDERS.find(function(w) { return String(w.woNumber || '') === woStr || String(w.id) === woStr; }) || null;
  // Prefer the Reports v2 UUID (wo.uuid); for DB API records wo.id is already the UUID
  var isDbApiMatch = !wo ? false : TURN_WORK_ORDERS.some(function(w) { return w === wo; });
  var wo_db_uuid = wo ? (isDbApiMatch ? String(wo.id || '') : String(wo.uuid || '')) : '';
  var payload = {
    turn_key: turnKey,
    wo_id: woStr,
    wo_db_uuid: wo_db_uuid,
    source: 'manual',
    status: wo && wo.status ? wo.status : '',
    created_at: wo && (wo.created || wo.createdAt) ? (wo.created || wo.createdAt) : ''
  };
  await proxyPost('unit_turn_wo_link', payload);
  await fetchUnitTurnsDB();
  renderTurnBoard();
}

async function unlinkTurnWorkOrder(turnKey, woId) {
  if (!turnKey || !woId) return;
  if (isTurnActionLocked(turnKey)) {
    throw new Error('Turn is terminal/closed; unlink actions are locked');
  }
  var woStr = String(woId).replace(/^\s*#/, '').trim();
  var wo = WORK_ORDERS.find(function(w) { return String(w.id) === woStr; }) ||
    TURN_WORK_ORDERS.find(function(w) { return String(w.woNumber || '') === woStr || String(w.id) === woStr; }) || null;
  var isDbApiMatch = !wo ? false : TURN_WORK_ORDERS.some(function(w) { return w === wo; });
  var woDbUuid = wo ? (isDbApiMatch ? String(wo.id || '') : String(wo.uuid || '')) : '';
  await proxyPost('unit_turn_wo_unlink', { turn_key: turnKey, wo_id: woStr, wo_db_uuid: woDbUuid });
  await fetchUnitTurnsDB();
  renderTurnBoard();
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
  var autoClosedCandidates = [];

  function isTurnWorkOrderCandidate(woLike, moveOutDate, source) {
    var createdLike = source === 'db'
      ? (woLike && (woLike.createdAt || woLike.lastUpdated))
      : (woLike && (woLike.created || woLike.updated));

    // Requirement: all WOs for the unit following move-out must be associated.
    if (moveOutDate && createdLike) {
      var woDate = new Date(createdLike);
      if (!isNaN(woDate.getTime())) {
        var diff = (woDate - moveOutDate) / 86400000;
        if (diff < -45) return false;
      }
    }
    return true;
  }

  // ---- Pre-build lookup index for O(1) matching ----
  // Strict matching is unit-level (unit_id + unit_turn_id), keyed by unit_id.
  var _turnWoByUnit = {};
  TURN_WORK_ORDERS.forEach(function(wo) {
    if (wo.unitId) { if (!_turnWoByUnit[wo.unitId]) _turnWoByUnit[wo.unitId] = []; _turnWoByUnit[wo.unitId].push(wo); }
  });

  // Helper: create a composite key for deduplication
  function makeKey(propId, unitId, moveOut) {
    var k = String(propId || '') + '-' + String(unitId || '') + '-' + (moveOut || '');
    return k === '--' ? null : k;
  }

  // Helper: find matching WOs using tolerant matching (unit/property/date + optional unit_turn_id)
  // This avoids false negatives caused by strict unit_turn_id-only gates and allows pre-moveout planning WOs.
  function findMatchingWOs(turnKey, unit, property, propId, unitId, moveOutDate, turnData) {
    var requiredUnitTurnId = String((turnData && turnData.unitTurnId) || '').trim();
    var requiredUnitId = String(unitId || '').trim();
    var requiredPropertyId = String(propId || '').trim();
    var requiredUnitName = String(unit || '').trim().toLowerCase();
    var requiredPropertyName = String(property || '').trim().toLowerCase();

    var winStart = moveOutDate ? new Date(moveOutDate.getTime() - 45 * 86400000) : null;

    function inWindow(dateLike) {
      if (!moveOutDate) return true;
      var d = new Date(dateLike || '');
      if (isNaN(d.getTime())) return false;
      return d >= winStart;
    }

    function candidateId(wo) {
      return String(wo.id || wo.woNumber || '').trim();
    }

    function isLikelyMatch(woUnitId, woUnitTurnId, woPropertyId, woUnitName, woPropertyName, woCreated) {
      var score = 0;
      if (requiredUnitId && woUnitId && String(woUnitId) === requiredUnitId) score += 4;
      if (requiredUnitTurnId && woUnitTurnId && String(woUnitTurnId) === requiredUnitTurnId) score += 5;
      if (requiredPropertyId && woPropertyId && String(woPropertyId) === requiredPropertyId) score += 2;
      if (requiredUnitName && woUnitName && String(woUnitName).toLowerCase() === requiredUnitName) score += 2;
      if (requiredPropertyName && woPropertyName && String(woPropertyName).toLowerCase() === requiredPropertyName) score += 1;
      if (inWindow(woCreated)) score += 2;

      // If there is an official unit_turn_id, still allow fallback via strong unit/property/date signals.
      if (requiredUnitTurnId) {
        return score >= 5;
      }
      return score >= 4;
    }

    var wos = [];
    function upsertWO(next) {
      var idVal = String(next.id || '').trim();
      if (!idVal) return;
      var dupe = wos.find(function(w) {
        return String(w.id) === idVal ||
          (next.dbApiId && String(w.dbApiId || '') === String(next.dbApiId)) ||
          (next.woUuid && String(w.woUuid || '') === String(next.woUuid));
      });
      if (!dupe) {
        wos.push(next);
        return;
      }
      // Merge with DB status/source when available.
      if (next.source === 'db_api') dupe.source = 'db_api';
      if (next.dbApiId) dupe.dbApiId = next.dbApiId;
      if (next.woUuid) dupe.woUuid = next.woUuid;
      if (next.status) dupe.status = next.status;
      if (!dupe.description && next.description) dupe.description = next.description;
      if (!dupe.created && next.created) dupe.created = next.created;
      if (!dupe.vendor && next.vendor) dupe.vendor = next.vendor;
      if (!dupe.priority && next.priority) dupe.priority = next.priority;
    }

    WORK_ORDERS.forEach(function(wo) {
      if (!isTurnWorkOrderCandidate(wo, moveOutDate, 'reports')) return;
      var woUnitId = String(wo.unitId || '').trim();
      var woTurnId = String(wo.unitTurnId || '').trim();
      var woPropId = String(wo.propertyId || '').trim();
      if (!isLikelyMatch(woUnitId, woTurnId, woPropId, wo.unit, wo.propertyName, wo.created)) return;
      upsertWO({
        source: 'reports',
        id: wo.id,
        woUuid: wo.uuid || '',
        status: wo.status,
        description: wo.description || '',
        created: wo.created,
        vendor: wo.vendorName || wo.vendor || '',
        unit: wo.unit,
        property: wo.propertyName,
        priority: wo.priority
      });
    });

    var dbCandidates = requiredUnitId ? (_turnWoByUnit[requiredUnitId] || []) : TURN_WORK_ORDERS;
    dbCandidates.forEach(function(wo) {
      if (!isTurnWorkOrderCandidate(wo, moveOutDate, 'db')) return;
      var woUnitId = String(wo.unitId || '').trim();
      var woTurnId = String(wo.unitTurnId || '').trim();
      var woPropId = String(wo.propertyId || '').trim();
      if (!isLikelyMatch(woUnitId, woTurnId, woPropId, wo.unit || '', wo.property || '', wo.createdAt || wo.lastUpdated)) return;
      upsertWO({
        source: 'db_api',
        id: wo.woNumber || wo.id,
        dbApiId: wo.id,
        status: wo.status,
        description: wo.description || '',
        created: wo.createdAt || wo.lastUpdated || '',
        vendor: wo.vendorTrade || '',
        priority: wo.priority
      });
    });

    var tracked = UNIT_TURN_TRACKER_BY_KEY[String(turnKey || '')];
    if (tracked && Array.isArray(tracked.linkedWorkOrders)) {
      tracked.linkedWorkOrders.forEach(function(w) {
        var idVal = String(w.id || '').trim();
        if (!idVal) return;
        upsertWO({
          source: 'tracker',
          id: idVal,
          dbApiId: w.dbApiId || '',
          status: w.status || '',
          created: w.created || '',
          description: '',
          vendor: '',
          priority: ''
        });
      });
    }

    wos.sort(function(a, b) {
      var ad = new Date(a.created || 0).getTime();
      var bd = new Date(b.created || 0).getTime();
      return (isNaN(ad) ? 0 : ad) - (isNaN(bd) ? 0 : bd);
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
    var woStatuses = matchingWOs.map(function(w) { return String(w.status || '').trim(); });
    var woCreatedDate = hasWO ? matchingWOs[0].created : null;

    var hasEstReq = woStatuses.some(function(s) { return String(s).toLowerCase() === 'estimate requested'; });
    var hasEstimated = woStatuses.some(function(s) { return String(s).toLowerCase() === 'estimated'; });
    var hasAssigned = woStatuses.some(function(s) {
      var n = String(s || '').toLowerCase();
      return n === 'assigned' || n === 'scheduled' || n === 'vendor assigned';
    });
    var hasActive = woStatuses.some(function(s) { return isTurnWorkActiveStatus(s); });
    var hasAnyWorkDone = woStatuses.some(function(s) { return isTurnWorkDoneStatus(s); });
    // ALL WOs must be in terminal status for work_done to be truly "done"
    var allWorkDone = hasWO && woStatuses.every(function(s) { return isClosedTurnWorkOrderStatus(s); });
    var doneCount = woStatuses.filter(function(s) { return isClosedTurnWorkOrderStatus(s); }).length;
    var progressPct = hasWO ? Math.round((doneCount / matchingWOs.length) * 100) : 0;

    // Progressive — later stages imply earlier ones
    var firstEstReq = null;
    var firstEstimated = null;
    var firstAssigned = null;
    matchingWOs.forEach(function(w) {
      var st = String(w.status || '').toLowerCase();
      if (!firstEstReq && (st === 'estimate requested' || st === 'estimated' || st === 'assigned' || st === 'scheduled' || st === 'vendor assigned' || isTurnWorkActiveStatus(st) || isTurnWorkDoneStatus(st))) firstEstReq = w.created || null;
      if (!firstEstimated && (st === 'estimated' || st === 'assigned' || st === 'scheduled' || st === 'vendor assigned' || isTurnWorkActiveStatus(st) || isTurnWorkDoneStatus(st))) firstEstimated = w.created || null;
      if (!firstAssigned && (st === 'assigned' || st === 'scheduled' || st === 'vendor assigned' || isTurnWorkActiveStatus(st) || isTurnWorkDoneStatus(st))) firstAssigned = w.created || null;
    });

    stages.wo_created = { done: hasWO, date: woCreatedDate, woIds: matchingWOs.map(function(w) { return w.id; }) };
    stages.est_requested = { done: hasEstReq || hasEstimated || hasAssigned || hasActive || hasAnyWorkDone, date: firstEstReq };
    stages.est_received = { done: hasEstimated || hasAssigned || hasActive || hasAnyWorkDone, date: firstEstimated, vendors: [] };
    stages.assigned = { done: hasAssigned || hasActive || hasAnyWorkDone, date: firstAssigned };
    // work_done requires ALL WOs complete, not just one
    stages.work_done = { done: allWorkDone, date: null, doneCount: doneCount, totalCount: matchingWOs.length, progressPct: progressPct };

    return stages;
  }

  // Helper: build a pipeline entry
  function addEntry(key, unit, property, propId, unitId, moveOut, turnData, moveoutTenant) {
    if (!key || seenKeys[key]) return;
    seenKeys[key] = true;

    // Exclude units that already have a confirmed move-in — the turn is over.
    // Check turnData, UNIT_TURNS_DB, and UPCOMING_MOVEOUTS for an actual_move_in date.
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
      if (!isNaN(miDate.getTime()) && miDate <= today) {
        autoClosedCandidates.push({
          turn_key: key,
          unit_turn_id: (turnData && turnData.unitTurnId) || '',
          unit_id: unitId || '',
          property_id: propId || '',
          unit_name: unit || '',
          property_name: property || '',
          move_out_date: moveOut || '',
          move_in_date: actualMoveIn,
          status: 'closed',
          closed_at: new Date().toISOString(),
          confidence_score: 100,
          confidence_label: 'high',
          site_manager: (turnData && turnData.siteManager) || '',
          source_flags: {
            auto_closed_by_move_in: true,
            from_turn_report: !!turnData
          },
          metadata: {
            close_reason: 'move_in_detected',
            close_source: 'system_move_in'
          },
          milestones: {
            moveout: { date: moveOut || '', source: turnData ? 'reports' : 'inferred' },
            movein: { date: actualMoveIn, source: 'reports' }
          },
          replace_work_orders: false,
          work_orders: []
        });
        return; // Unit has a confirmed past move-in — exclude from turn board
      }
    }

    var moveOutDate = moveOut ? new Date(moveOut) : null;
    var isUpcoming = moveOutDate ? moveOutDate > today : false;
    var turnYearStart = getCurrentYearStartDate(today);
    if (moveOutDate && !isNaN(moveOutDate.getTime()) && moveOutDate < turnYearStart) {
      return; // Exclude prior-year turns from active board and badge counts
    }
    var matchingWOs = findMatchingWOs(key, unit, property, propId, unitId, moveOutDate, turnData);
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
    var isStalled = !isUpcoming && elapsed > CONFIG.TURN_STALLED_DAYS && currentStageIdx >= 1 && currentStageIdx < PIPE_STAGES.length - 1;
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

    var confidence = computeTurnConfidence(turnData, stages, matchingWOs, matchingInsp);

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
      isClosed: CLOSED_TURNS.has(key), // manually closed from turn board
      allWOsDone: allWOsDone,
      confidenceScore: (dbTurnMatch && dbTurnMatch.confidenceScore) || confidence.score,
      confidenceLabel: (dbTurnMatch && dbTurnMatch.confidenceLabel) || confidence.label,
      trackingUuid: (dbTurnMatch && dbTurnMatch.trackingUuid) || '',
      trackingCode: (dbTurnMatch && dbTurnMatch.trackingCode) || '',
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

  if (autoClosedCandidates.length > 0) {
    syncAutoClosedTurns(autoClosedCandidates).catch(function() {});
  }
}

async function fetchPropertyStats() {
  try {
    var data = await proxyAction('property_stats');
    if (data && data.ok && data.by_property && typeof data.by_property === 'object') {
      _propertyStatsById = data.by_property;
      return true;
    }
  } catch (e) {
    console.log('fetchPropertyStats error:', e && (e.message || e));
  }
  _propertyStatsById = {};
  return false;
}

// Fetch all units from the proxy cache and build the _unitsByPropertyId lookup.
async function fetchUnits() {
  try {
    var data = await proxyAction('units');
    if (data && data.ok && Array.isArray(data.results)) {
      UNITS = data.results;
      _unitsByPropertyId = {};
      UNITS.forEach(function(u) {
        var pid = String(u.property_id || '').trim();
        if (!pid) return;
        if (!_unitsByPropertyId[pid]) _unitsByPropertyId[pid] = [];
        _unitsByPropertyId[pid].push(u);
      });
      return true;
    }
  } catch (e) {
    console.log('fetchUnits error:', e && (e.message || e));
  }
  UNITS = [];
  _unitsByPropertyId = {};
  return false;
}

function getPropertyStats(propId) {
  var pid = String(propId || '').trim();
  if (!pid || !_propertyStatsById || !_propertyStatsById[pid]) {
    return { bills: 0, notes: 0, listings: 0 };
  }
  var row = _propertyStatsById[pid] || {};
  return {
    bills: Number(row.bills || 0),
    notes: Number(row.notes || 0),
    listings: Number(row.listings || 0)
  };
}

function renderPropertiesRowsFromCache() {
  var tbody = $('#propertiesTableBody');
  if (!tbody) return;
  var searchEl = $('#propertySearch');
  var groupSelect = $('#propertyGroupFilter');
  var properties = (PROPERTIES || []).slice();

  var scope = getPropertiesScope();
  var globalGroup = scope.globalGroup;

  // Always enforce global scope first.
  if (globalGroup) {
    properties = properties.filter(function(p) {
      return propertyMatchesSelectedGroup(p, globalGroup);
    });
  }

  var groups = {};
  properties.forEach(function(p) {
    var g = String(p.propertyGroup || p.groupName || p.portfolio || '').trim();
    if (g) groups[g] = true;
  });

  var localGroup = scope.localGroup;
  var effectiveGroup = localGroup || globalGroup;

  if (groupSelect) {
    var existing = normalizeGroupSelectionValue(groupSelect.value || localGroup || '');
    var options = ['<option value="">All Groups</option>'];
    Object.keys(groups).sort().forEach(function(g) {
      options.push('<option value="' + escapeHtml(g) + '">' + escapeHtml(g) + '</option>');
    });
    groupSelect.innerHTML = options.join('');
    if (existing && !groups[existing]) existing = '';
    groupSelect.value = existing;
    _propertiesLocalGroup = normalizeGroupSelectionValue(groupSelect.value || '');
    localGroup = _propertiesLocalGroup;
    effectiveGroup = localGroup || globalGroup;
  }

  if (localGroup) {
    properties = properties.filter(function(p) {
      return propertyMatchesSelectedGroup(p, localGroup);
    });
  }

  var search = String(searchEl ? searchEl.value : '').trim().toLowerCase();
  function propertyIsVacantForFilter(p) {
    var statusText = String(p.status || p.occupancyStatus || p.unit_status || p.Status || '').toLowerCase();
    if (p.isVacant === true || p.IsVacant === true || p.vacant === true) return true;
    return statusText.indexOf('vacant') !== -1 || statusText.indexOf('available') !== -1;
  }

  if (_propertiesVacancyOnly) {
    properties = properties.filter(propertyIsVacantForFilter);
  }

  if (search) {
    properties = properties.filter(function(p) {
      var hay = [
        p.name,
        p.address,
        p.city,
        p.state,
        p.zip,
        p.propertyGroup,
        p.groupName,
        p.portfolio,
        p.siteManager
      ].join(' ').toLowerCase();
      return hay.indexOf(search) !== -1;
    });
  }

  properties.sort(function(a, b) {
    return String(a.name || '').localeCompare(String(b.name || ''));
  });

  var footer = $('#propertiesFooter');
  var meta = $('#propertiesPageMeta');
  var prevBtn = $('#propertiesPrevPage');
  var nextBtn = $('#propertiesNextPage');
  var pageSizeSel = $('#propertiesPageSize');
  if (pageSizeSel && String(pageSizeSel.value || '') !== String(_propertiesPageSize)) {
    pageSizeSel.value = String(_propertiesPageSize);
  }

  var total = properties.length;
  var size = Math.max(1, parseInt(_propertiesPageSize, 10) || 50);
  _propertiesPageSize = size;
  var pages = Math.max(1, Math.ceil(total / size));
  if (_propertiesPage >= pages) _propertiesPage = pages - 1;
  if (_propertiesPage < 0) _propertiesPage = 0;

  var start = _propertiesPage * size;
  var end = Math.min(start + size, total);
  var pageRows = properties.slice(start, end);

  if (footer) footer.style.display = 'flex';
  if (meta) {
    if (total === 0) {
      meta.textContent = 'Page 1 of 1 • 0 results' + (_propertiesVacancyOnly ? ' • vacancy-only' : '');
    } else {
      meta.textContent = 'Page ' + (_propertiesPage + 1) + ' of ' + pages + ' • Showing ' + (start + 1) + '-' + end + ' of ' + total + (_propertiesVacancyOnly ? ' • vacancy-only' : '');
    }
  }
  if (prevBtn) prevBtn.disabled = (_propertiesPage <= 0 || total === 0);
  if (nextBtn) nextBtn.disabled = (_propertiesPage >= pages - 1 || total === 0);

  if (!properties.length) {
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:20px;color:var(--text-muted);">' +
      (_propertiesVacancyOnly ? 'No vacant properties found for current scope' : 'No properties found') +
      '</td></tr>';
    return;
  }

  var html = pageRows.map(function(p) {
    var pid = String(p.id || '');
    var stats = getPropertyStats(pid);
    var cityState = [p.city || '', p.state || ''].filter(Boolean).join(', ');
    if (p.zip) cityState += (cityState ? ' ' : '') + p.zip;
    var groupName = String(p.propertyGroup || p.groupName || p.portfolio || '');
    var units = _unitsByPropertyId[pid] || [];
    var unitCount = units.length;
    var unitBadge = unitCount > 1
      ? '<span class="tag blue" title="' + unitCount + ' units">' + unitCount + ' units</span>'
      : (unitCount === 1 ? '<span class="tag normal">1 unit</span>' : '<span style="color:var(--text-muted)">—</span>');
    return '<tr>' +
      '<td>' + escapeHtml(String(p.name || '—')) + '</td>' +
      '<td>' + escapeHtml(String(p.address || '—')) + '</td>' +
      '<td>' + escapeHtml(cityState || '—') + '</td>' +
      '<td>' + escapeHtml(groupName || '—') + '</td>' +
      '<td>' + escapeHtml(String(p.siteManager || '—')) + '</td>' +
      '<td style="text-align:center">' + unitBadge + '</td>' +
      '<td style="text-align:center"><span class="tag normal">' + String(stats.bills) + '</span></td>' +
      '<td style="text-align:center"><span class="tag blue">' + String(stats.notes) + '</span></td>' +
      '<td style="text-align:center"><span class="tag completed">' + String(stats.listings) + '</span></td>' +
      '<td><button class="action-btn" data-property-id="' + escapeHtml(pid) + '" data-property-name="' + escapeHtml(String(p.name || '')) + '" onclick="showPropertyDetailModal(this)">Details</button></td>' +
    '</tr>';
  }).join('');

  tbody.innerHTML = html;
}

function renderPropertiesSection(opts) {
  opts = opts || {};
  var tbody = $('#propertiesTableBody');
  if (!tbody) return;

  var hasCache = Array.isArray(PROPERTIES) && PROPERTIES.length > 0;
  var forceRefresh = !!opts.forceRefresh;

  if (hasCache && !forceRefresh) {
    renderPropertiesRowsFromCache();
    return;
  }

  if (_propertiesRefreshInFlight) return;
  _propertiesRefreshInFlight = true;

  tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:20px;color:var(--text-muted);">Loading properties...</td></tr>';
  if (!hasCache) setSectionBusy('sec-properties', true, 'Refreshing properties…');

  Promise.all([
    fetchProperties(),
    fetchPropertyStats(),
    fetchUnits()
  ]).then(function() {
    renderPropertiesRowsFromCache();
  }).catch(function(err) {
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:20px;color:var(--danger);">Failed to load properties: ' + escapeHtml(String(err && (err.message || err) || 'Unknown error')) + '</td></tr>';
  }).finally(function() {
    _propertiesRefreshInFlight = false;
    setSectionBusy('sec-properties', false);
  });
}

function showPropertyDetailModal(btn) {
  var propertyId = btn ? String(btn.getAttribute('data-property-id') || '') : '';
  var propertyName = btn ? String(btn.getAttribute('data-property-name') || '') : '';
  if (!propertyId) return;

  var modalBody = $('#itemDetailBody');
  var modalTitle = $('#itemDetailTitle');
  var linkBtn = $('#itemDetailLink');
  if (!modalBody || !modalTitle) return;

  modalTitle.textContent = propertyName || 'Property Details';
  if (linkBtn) linkBtn.style.display = 'none';
  modalBody.innerHTML = loadingHtml('Loading property details…');
  openModal('itemDetailModal');

  Promise.all([
    proxyAction('property_notes', { property_id: propertyId }),
    proxyAction('property_listings', { property_id: propertyId, days: 365 })
  ]).then(function(results) {
    var notesRes = results[0] || { ok: false, results: [] };
    var listingsRes = results[1] || { ok: false, results: [] };
    var notes = (notesRes.ok && Array.isArray(notesRes.results)) ? notesRes.results : [];
    var listings = (listingsRes.ok && Array.isArray(listingsRes.results)) ? listingsRes.results : [];
    showPropertyModal(propertyId, propertyName, notes, listings);
  }).catch(function(err) {
    modalBody.innerHTML = '<div style="padding:14px;color:var(--danger)">Failed to load property details: ' + escapeHtml(String(err && (err.message || err) || 'Unknown error')) + '</div>';
  });
}

function showPropertyModal(propertyId, propertyName, notes, listings) {
  var modalBody = $('#itemDetailBody');
  var modalTitle = $('#itemDetailTitle');
  if (!modalBody || !modalTitle) return;
  modalTitle.textContent = propertyName || 'Property Details';
  var safePropertyId = String(propertyId || '').replace(/'/g, "\\'");
  var safePropertyName = String(propertyName || '').replace(/'/g, "\\'");

  var notesHtml = '<p style="color:var(--text-muted);margin:12px 0;">No notes for this property</p>';
  if (notes.length) {
    notesHtml = '<div style="display:flex;flex-direction:column;gap:8px;">' + notes.map(function(n) {
      var d = n.last_updated_at ? formatDate(n.last_updated_at) : '—';
      return '<div style="padding:8px;background:var(--bg-secondary);border-radius:6px;border-left:3px solid var(--accent);">' +
        '<div style="font-size:11px;color:var(--text-muted);">' + escapeHtml(d) + '</div>' +
        '<div style="font-size:12px;color:var(--text-primary);white-space:pre-wrap;">' + escapeHtml(String(n.body || '')) + '</div>' +
      '</div>';
    }).join('') + '</div>';
  }

  var listingsHtml = '<p style="color:var(--text-muted);margin:12px 0;">No active listings for this property</p>';
  if (listings.length) {
    listingsHtml = '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:10px;">' + listings.map(function(l) {
      var rent = l.advertised_rent || l.listed_rent || '—';
      var avail = l.available_on ? formatDate(l.available_on) : '—';
      return '<div style="padding:10px;border:1px solid var(--border);border-radius:6px;background:var(--bg-secondary);">' +
        '<div style="font-weight:600">Unit ' + escapeHtml(String(l.unit_id || '—')) + '</div>' +
        '<div style="font-size:12px;color:var(--text-muted)">' + escapeHtml(String(l.bedrooms || 0)) + ' bed / ' + escapeHtml(String(l.bathrooms || 0)) + ' bath</div>' +
        '<div style="font-family:var(--font-mono);margin-top:6px">$' + escapeHtml(String(typeof rent === 'number' ? rent.toLocaleString() : rent)) + '</div>' +
        '<div style="font-size:11px;color:var(--text-muted)">Available: ' + escapeHtml(avail) + '</div>' +
      '</div>';
    }).join('') + '</div>';
  }

  // Units tab — from the units cache
  var propertyUnits = _unitsByPropertyId[String(propertyId || '')] || [];
  var unitsHtml = '<p style="color:var(--text-muted);margin:12px 0;">No units cached for this property</p>';
  if (propertyUnits.length) {
    var STATUS_STYLE = {
      'Occupied': 'color:var(--success)',
      'Vacant': 'color:var(--danger)',
      'Vacant Rented': 'color:var(--warning)',
      'Notice': 'color:var(--warning)',
      'Notice Rented': 'color:var(--info)'
    };
    unitsHtml = '<div style="display:flex;flex-direction:column;gap:6px;max-height:360px;overflow-y:auto;">' +
      propertyUnits.map(function(u) {
        var uid = String(u.unit_id || '');
        var uname = String(u.name || u.unit_id || '—');
        var status = String(u.status || '');
        var statusStyle = STATUS_STYLE[status] || 'color:var(--text-muted)';
        var beds = u.bedrooms != null ? u.bedrooms + ' bd' : '';
        var baths = u.bathrooms ? u.bathrooms + ' ba' : '';
        var safeUid = uid.replace(/'/g, "\\'");
        var safeUname = uname.replace(/'/g, "\\'");
        return '<div style="display:flex;align-items:center;gap:10px;padding:8px;background:var(--bg-secondary);border-radius:6px;border-left:3px solid var(--border);">' +
          '<div style="flex:1;min-width:0;">' +
            '<div style="font-weight:600;font-size:12px;">' + escapeHtml(uname) + '</div>' +
            '<div style="font-size:11px;color:var(--text-muted);">' + escapeHtml([beds, baths].filter(Boolean).join(' / ') || '') + '</div>' +
          '</div>' +
          '<div style="font-size:11px;' + statusStyle + '">' + escapeHtml(status || '—') + '</div>' +
          (uid ? '<button class="action-btn" style="padding:2px 8px;font-size:10px;" onclick="filterBillsToUnit(\'' + safeUid + '\', \'' + safeUname + '\', \'' + safePropertyId + '\', \'' + safePropertyName + '\')"><i class="fas fa-filter"></i> Bills</button>' : '') +
        '</div>';
      }).join('') +
    '</div>';
  }

  modalBody.innerHTML =
    '<div style="display:flex;gap:8px;margin-bottom:10px;border-bottom:1px solid var(--border);padding-bottom:8px">' +
      '<button class="action-btn active modal-tab-btn" data-tab="overview" onclick="switchPropertyTab(this)">Overview</button>' +
      '<button class="action-btn modal-tab-btn" data-tab="units" onclick="switchPropertyTab(this)">Units (' + propertyUnits.length + ')</button>' +
      '<button class="action-btn modal-tab-btn" data-tab="notes" onclick="switchPropertyTab(this)">Notes (' + notes.length + ')</button>' +
      '<button class="action-btn modal-tab-btn" data-tab="listings" onclick="switchPropertyTab(this)">Listings (' + listings.length + ')</button>' +
    '</div>' +
    '<div id="prop-overview-content" class="property-tab-content">' +
      '<div class="detail-row"><div class="detail-row-label">Property</div><div class="detail-row-value">' + escapeHtml(propertyName || '—') + '</div></div>' +
      '<div style="margin-top:10px"><button class="action-btn primary" onclick="filterBillsToProperty(\'' + safePropertyId + '\', \'' + safePropertyName + '\')"><i class="fas fa-filter"></i> View Bills for This Property</button></div>' +
    '</div>' +
    '<div id="prop-units-content" class="property-tab-content" style="display:none">' + unitsHtml + '</div>' +
    '<div id="prop-notes-content" class="property-tab-content" style="display:none">' + notesHtml + '</div>' +
    '<div id="prop-listings-content" class="property-tab-content" style="display:none">' + listingsHtml + '</div>';
}

function switchPropertyTab(btn) {
  if (!btn) return;
  var tabName = String(btn.getAttribute('data-tab') || 'overview');
  $$('.modal-tab-btn').forEach(function(el) { el.classList.remove('active'); });
  btn.classList.add('active');
  $$('.property-tab-content').forEach(function(el) { el.style.display = 'none'; });
  var panel = $('#prop-' + tabName + '-content');
  if (panel) panel.style.display = '';
}

function filterBillsToProperty(propertyId, propertyName) {
  window.filteredPropertyId = String(propertyId || '');
  window.filteredPropertyName = String(propertyName || '');
  window.filteredUnitId = '';
  window.filteredUnitName = '';
  var search = $('#billSearch');
  if (search && propertyName) search.value = String(propertyName);
  forceActiveTab('billing');
  renderBillsSection();
}

function filterBillsToUnit(unitId, unitName, propertyId, propertyName) {
  window.filteredPropertyId = String(propertyId || '');
  window.filteredPropertyName = String(propertyName || '');
  window.filteredUnitId = String(unitId || '');
  window.filteredUnitName = String(unitName || '');
  var search = $('#billSearch');
  if (search && unitName) search.value = String(unitName);
  forceActiveTab('billing');
  renderBillsSection();
}

function openBulkNoteModal() {
  var groupEl = $('#propertyGroupFilter');
  var currentGroup = normalizeGroupSelectionValue(groupEl ? groupEl.value : _propertiesLocalGroup);
  var globalGroup = normalizeGroupSelectionValue(getEffectiveGroupId());

  // Determine the target properties based on current group filter
  var targetProps = (PROPERTIES || []).filter(function(p) {
    if (globalGroup && !propertyMatchesSelectedGroup(p, globalGroup)) return false;
    if (!currentGroup) return true;
    return propertyMatchesSelectedGroup(p, currentGroup);
  });

  var targetGroupEl = $('#bulkNoteTargetGroup');
  var countEl = $('#bulkNotePropertyCount');
  var textEl = $('#bulkNoteText');
  var statusEl = $('#bulkNoteStatus');
  var submitBtn = $('#bulkNoteSubmitBtn');

  if (targetGroupEl) targetGroupEl.textContent = currentGroup || globalGroup || 'All Groups';
  if (countEl) countEl.textContent = targetProps.length + ' propert' + (targetProps.length === 1 ? 'y' : 'ies') + ' will receive this note';
  if (textEl) textEl.value = '';
  if (statusEl) statusEl.textContent = '';
  if (submitBtn) submitBtn.disabled = false;

  openModal('bulkNoteModal');
}

function executeBulkNoteUpdate() {
  var textEl = $('#bulkNoteText');
  var statusEl = $('#bulkNoteStatus');
  var submitBtn = $('#bulkNoteSubmitBtn');
  var noteText = String(textEl ? textEl.value : '').trim();

  if (!noteText) {
    if (statusEl) { statusEl.style.color = 'var(--danger)'; statusEl.textContent = 'Note text is required.'; }
    return;
  }

  var groupEl = $('#propertyGroupFilter');
  var currentGroup = normalizeGroupSelectionValue(groupEl ? groupEl.value : _propertiesLocalGroup);
  var globalGroup = normalizeGroupSelectionValue(getEffectiveGroupId());
  var targetProps = (PROPERTIES || []).filter(function(p) {
    if (globalGroup && !propertyMatchesSelectedGroup(p, globalGroup)) return false;
    if (!currentGroup) return true;
    return propertyMatchesSelectedGroup(p, currentGroup);
  });

  if (!targetProps.length) {
    if (statusEl) { statusEl.style.color = 'var(--danger)'; statusEl.textContent = 'No properties found for current filter.'; }
    return;
  }

  var targetGroup = currentGroup || globalGroup || 'All Groups';
  var count = targetProps.length;

  hmConfirm(
    'Post this note to ' + count + ' propert' + (count === 1 ? 'y' : 'ies') + ' in "' + targetGroup + '"? This cannot be undone.',
    {
      confirmLabel: 'Post Notes',
      cancelLabel: 'Cancel',
      onConfirm: function() {
        var ids = targetProps.map(function(p) { return String(p.id || ''); }).filter(Boolean);
        if (!ids.length) {
          if (statusEl) { statusEl.style.color = 'var(--danger)'; statusEl.textContent = 'No valid property IDs found.'; }
          return;
        }

        if (submitBtn) submitBtn.disabled = true;
        if (statusEl) { statusEl.style.color = 'var(--text-muted)'; statusEl.textContent = 'Posting notes… (0/' + ids.length + ')'; }

        proxyPost('bulk_update_notes', { property_ids: ids, note_body: noteText })
          .then(function(res) {
            if (!res || !res.ok) {
              var errMsg = String(res && res.error || 'Unknown error');
              if (statusEl) { statusEl.style.color = 'var(--danger)'; statusEl.textContent = 'Failed: ' + errMsg; }
              if (submitBtn) submitBtn.disabled = false;
              return;
            }
            var succeeded = Number(res.success_count || 0);
            var failed = Number(res.failure_count || 0);
            if (statusEl) {
              statusEl.style.color = failed ? 'var(--warning)' : 'var(--success)';
              statusEl.textContent = 'Done — ' + succeeded + ' succeeded, ' + failed + ' failed.';
            }
            showToast('Notes posted: ' + succeeded + ' ok' + (failed ? ', ' + failed + ' failed' : ''), { type: failed ? 'warn' : 'success' });
            if (failed === 0) {
              setTimeout(function() { closeModal('bulkNoteModal'); }, 1200);
            } else if (submitBtn) {
              submitBtn.disabled = false;
            }
          })
          .catch(function(err) {
            var errMsg = String(err && (err.message || err) || 'Network error');
            if (statusEl) { statusEl.style.color = 'var(--danger)'; statusEl.textContent = 'Error: ' + errMsg; }
            if (submitBtn) submitBtn.disabled = false;
          });
      }
    }
  );
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
  syncTurnTrackerFromPipeline().catch(function() {});
}

function renderTurnKPIs() {
  var inScope = TURN_PIPE_DATA.filter(function(p) {
    return isInPropertyGroup(p.propertyId, p.property, currentPropertyGroup);
  });

  // Separate confirmed-active turns from on-radar (unconfirmed) entries
  var confirmed = inScope.filter(function(p) { return p.isConfirmed && !p.isCompleted; });
  var onRadar   = inScope.filter(function(p) { return p.isOnRadar; });
  var upcoming  = inScope.filter(function(p) { return p.isUpcoming; });
  var radarUpcomingCount = onRadar.filter(function(p) { return p.isUpcoming; }).length;
  var active    = confirmed; // alias — same set for downstream stats
  var awaitEst  = confirmed.filter(function(p) {
    return p.stages.wo_created && p.stages.wo_created.done && p.stages.est_received && !p.stages.est_received.done;
  });
  var avgDays = 0;
  if (confirmed.length > 0) {
    var totalDays = 0;
    confirmed.forEach(function(p) { totalDays += Math.abs(p.elapsed); });
    avgDays = Math.round(totalDays / confirmed.length);
  }

  var e = function(id, v) { var el = document.getElementById(id); if (el) el.textContent = v; };
  e('kpiActiveTurns', confirmed.length);
  e('kpiActiveTurnsSub', confirmed.length + ' confirmed' + (upcoming.length > 0 ? ' (' + upcoming.length + ' upcoming)' : ''));
  e('kpiOnRadar', onRadar.length);
  var radarUpcoming = radarUpcomingCount;
  var radarPast     = onRadar.filter(function(p) { return !p.isUpcoming; }).length;
  e('kpiOnRadarSub', (radarUpcoming > 0 ? radarUpcoming + ' upcoming' : '') + (radarUpcoming > 0 && radarPast > 0 ? ', ' : '') + (radarPast > 0 ? radarPast + ' awaiting inspection' : '') || 'no possible turns');
  e('kpiAvgTurnDays', avgDays > 0 ? avgDays + 'd' : '\u2014');
  e('kpiAvgTurnSub', avgDays > 0 ? 'avg days elapsed' : 'no active turns');
  e('kpiAwaitEst', awaitEst.length);
  e('kpiAwaitEstSub', awaitEst.length + ' turns pending vendor bids');
  e('kpiTurnBilled', inScope.length);
  e('kpiTurnBilledSub', 'active + on radar + completed');

  var tb = $('#turnBadge');
  if (tb) tb.textContent = active.length + onRadar.length;

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
    // Closed turns only show in 'closed' filter
    if (p.isClosed && filter !== 'closed') return false;
    if (filter === 'closed' && !p.isClosed) return false;
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
    var cardClass = p.isCompleted ? 'completed' : p.isUpcoming ? 'upcoming' : p.isOnRadar ? 'on-radar' : slaBreach ? 'sla-breach' : p.isStalled ? 'stalled' : p.elapsed < CONFIG.TURN_WARNING_DAYS ? 'on-track' : 'waiting';
    html += '<div class="pipe-card ' + cardClass + '" data-pipeidx="' + idx + '" data-pipeid="' + escapeHtml(p.id) + '">';

    // Left: unit info
    html += '<div class="pipe-card-unit"><div class="pipe-card-unit-name">' + escapeHtml(p.unit || 'Unit') + '</div>';
    html += '<div class="pipe-card-prop">' + escapeHtml(p.property) + '</div>';
    if (!currentPropertyGroup && p.turn && p.turn._propertyGroup) html += '<div class="pipe-card-prop pipe-card-group-badge"><i class="fas fa-layer-group" style="font-size:9px;margin-right:3px"></i>' + escapeHtml(p.turn._propertyGroup) + '</div>';
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
      html += '<span class="pipe-card-cost">Complete</span>';
    } else if (p.isUpcoming) {
      var daysUntil = Math.abs(p.elapsed);
      html += '<span class="pipe-card-elapsed" style="color:var(--info,#60a5fa)">' + daysUntil + 'd</span>';
      html += '<span class="pipe-card-cost">Move-out: ' + formatDate(p.moveOut) + '</span>';
    } else if (p.isOnRadar) {
      html += '<span class="radar-badge" style="margin-bottom:3px">ON RADAR</span>';
      html += '<span class="pipe-card-elapsed" style="color:var(--accent)">' + p.elapsed + 'd</span>';
      html += '<span class="pipe-card-cost" style="color:var(--text-muted)">Awaiting inspection</span>';
    } else {
      var eColor = p.elapsed > p.target ? 'var(--danger)' : p.elapsed > CONFIG.TURN_WARNING_DAYS ? 'var(--warning)' : 'var(--text-primary)';
      html += '<span class="pipe-card-elapsed" style="color:' + eColor + '">' + p.elapsed + 'd</span>';
      var nextStage = p.currentStageIdx < PIPE_STAGES.length - 1 ? PIPE_STAGES[p.currentStageIdx + 1] : null;
      html += '<span class="pipe-card-cost">';
      if (nextStage) html += 'Next: ' + nextStage.label;
      else html += 'In progress';
      // WO progress: show X/Y done when there are multiple WOs
      if (p.stages.work_done && p.stages.work_done.totalCount > 0) {
        html += ' <span style="color:var(--text-muted);font-size:10px">(' + p.stages.work_done.doneCount + '/' + p.stages.work_done.totalCount + ' WOs)</span>';
      }
      html += '</span>';
      // Deposit deadline progress bar
      if (p.sla) {
        var slaColor = p.sla.businessDaysLeft <= 2 ? 'red' : p.sla.businessDaysLeft <= 6 ? 'yellow' : 'green';
        var slaLabel = p.sla.overdue ? 'Past deposit deadline' : p.sla.businessDaysLeft + ' biz days remaining';
        html += '<div class="sla-bar" title="Deposit deadline \u2014 ' + slaLabel + '"><div class="sla-bar-fill ' + slaColor + '" style="width:' + p.sla.pct + '%"></div></div>';
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
      if (ps.key === 'work_done' && stage.totalCount > 0 && !stage.done) {
        html += '<div class="pipe-tl-note">WO Progress: ' + stage.doneCount + '/' + stage.totalCount + ' (' + (stage.progressPct || 0) + '%)</div>';
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
      html += '<div style="font-size:11px;color:var(--accent);padding:5px 8px;background:var(--accent-dim);border-radius:4px;margin-bottom:6px"><i class="fas fa-info-circle" style="margin-right:4px"></i>Once a move-out inspection is recorded, this turn is confirmed and these WOs will be formally linked.</div>';
    }
    var turnLocked = !!(p.isClosed || p.isCompleted || isTerminalTurnStatus(p.unitTurnStatus || (p.turn && p.turn.status) || ''));
    if (p.matchingWOs.length > 0) {
      html += '<div class="pipe-wo-list">';
      p.matchingWOs.forEach(function(wo) {
        // Prefer DB API Link field (direct AppFolio URL), fall back to constructed URL
        var dbWo = wo.dbApiId ? TURN_WORK_ORDERS.find(function(tw) { return tw.id === wo.dbApiId; }) : null;
        var woLink = (dbWo && dbWo.link) ? dbWo.link : appfolioUrl('work_order', wo.id);
        html += '<div class="pipe-wo-item"><div><span class="pipe-wo-id">#' + wo.id + '</span> <span class="tag ' + String(wo.status || 'Linked').toLowerCase().replace(/\s+/g, '-') + '">' + escapeHtml(wo.status || 'Linked') + '</span>';
        if (woLink) html += ' <a href="' + escapeHtml(woLink) + '" target="_blank" rel="noopener noreferrer" style="font-size:9px;color:var(--accent);text-decoration:none" title="View WO in AppFolio" onclick="event.stopPropagation()"><i class="fas fa-external-link-alt"></i></a>';
        html += ' <button class="action-btn" data-unlink-wo="' + escapeHtml(p.id) + '" data-woid="' + escapeHtml(String(wo.id)) + '" style="padding:1px 6px;font-size:10px" onclick="event.stopPropagation()"' + (turnLocked ? ' disabled title="Turn is terminal/closed"' : '') + '><i class="fas fa-unlink"></i></button>';
        html += '</div>';
        html += '<div style="font-size:11px;color:var(--text-secondary)">' + escapeHtml((wo.description || '').substring(0, 60)) + '</div></div>';
      });
      html += '</div>';
    } else {
      html += '<div style="font-size:12px;color:var(--text-muted);padding:8px 0">No linked work orders found</div>';
    }
    html += '<div style="display:flex;gap:6px;margin-top:8px">';
    html += '<input class="form-input" data-add-wo-input="' + escapeHtml(p.id) + '" placeholder="Add WO # (manual)" style="font-size:11px;max-width:170px" onclick="event.stopPropagation()"' + (turnLocked ? ' disabled title="Turn is terminal/closed"' : '') + '>';
    html += '<button class="action-btn" data-add-wo="' + escapeHtml(p.id) + '" style="font-size:11px;padding:4px 8px" onclick="event.stopPropagation()"' + (turnLocked ? ' disabled title="Turn is terminal/closed"' : '') + '><i class="fas fa-link"></i> Add WO</button>';
    html += '</div>';
    if (turnLocked) {
      html += '<div style="font-size:11px;color:var(--warning);margin-top:6px"><i class="fas fa-lock"></i> Turn is terminal/closed; work-order link actions are disabled to preserve audit trail.</div>';
    }

    // Turn details
    html += '<div class="detail-section-title" style="margin-top:12px"><i class="fas fa-info-circle"></i> Turn Details</div>';
    html += '<div class="detail-grid">';
    html += '<div class="detail-row"><div class="detail-row-label">Tracking</div><div class="detail-row-value">' + escapeHtml(p.trackingCode || p.id) + '</div></div>';
    html += '<div class="detail-row"><div class="detail-row-label">Confidence</div><div class="detail-row-value"><span class="tag">' + escapeHtml(String((p.confidenceLabel || 'low')).toUpperCase()) + '</span> ' + escapeHtml(String(p.confidenceScore || 0)) + '%</div></div>';
    html += '<div class="detail-row"><div class="detail-row-label">Move-Out</div><div class="detail-row-value">' + (p.moveOut ? formatDate(p.moveOut) : '\u2014') + '</div></div>';
    html += '<div class="detail-row"><div class="detail-row-label">Expected Move-In</div><div class="detail-row-value">' + (p.expectedMoveIn ? formatDate(p.expectedMoveIn) : '\u2014') + '</div></div>';
    if (p.unitTurnStatus) html += '<div class="detail-row"><div class="detail-row-label">Turn Status</div><div class="detail-row-value">' + escapeHtml(p.unitTurnStatus) + '</div></div>';
    if (p.depositStatus) html += '<div class="detail-row"><div class="detail-row-label">Deposit Status</div><div class="detail-row-value">' + escapeHtml(p.depositStatus) + '</div></div>';
    if (p.depositReturnDeadline) html += '<div class="detail-row"><div class="detail-row-label">Deposit Return Date</div><div class="detail-row-value">' + formatDate(p.depositReturnDeadline) + '</div></div>';
    html += '<div class="detail-row"><div class="detail-row-label">Site Manager</div><div class="detail-row-value">' + escapeHtml(p.siteManager || '\u2014') + '</div></div>';
    html += '<div class="detail-row"><div class="detail-row-label">Maintenance Limit</div><div class="detail-row-value">' + escapeHtml(p.maintenanceLimit || '\u2014') + '</div></div>';
    html += '<div class="detail-row"><div class="detail-row-label">Property Notes</div><div class="detail-row-value">' + escapeHtml(p.propertyNotes || '\u2014') + '</div></div>';
    if (p.isOnRadar) {
      html += '<div class="detail-row"><div class="detail-row-label">Confirmation</div><div class="detail-row-value" style="color:var(--accent)"><i class="fas fa-satellite-dish" style="margin-right:4px"></i>On Radar \u2014 awaiting inspection</div></div>';
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
        html += '<div style="font-size:11px;color:var(--text-secondary)">Move-out: ' + (rt.moveOut ? formatDate(rt.moveOut) : '\u2014') + ' • Move-in: ' + (rt.expectedMoveIn ? formatDate(rt.expectedMoveIn) : '\u2014') + '</div>';
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
    if (!p.isClosed) {
      html += '<button class="action-btn" style="color:var(--danger);border-color:rgba(239,68,68,0.4)" data-close-turn="' + escapeHtml(p.id) + '" onclick="event.stopPropagation()"><i class="fas fa-times-circle"></i> Close Turn</button>';
    }
    html += '<button class="action-btn" data-close-detail="' + escapeHtml(p.id) + '"><i class="fas fa-times"></i> Dismiss</button>';
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
   TURN BOARD — Kanban View & Detail Modal
   ================================================================= */
var currentTurnViewMode = 'list';

function toggleTurnView(mode) {
  currentTurnViewMode = mode;
  var listEl = $('#turnPipeline');
  var kanbanEl = $('#turnKanban');
  var listBtn = $('#btnTurnViewList');
  var kanbanBtn = $('#btnTurnViewKanban');
  var legendEl = $('#pipelineStageLegend');
    if (listEl)   listEl.style.display   = (mode === 'list')   ? '' : 'none';
    if (kanbanEl) kanbanEl.style.display = (mode === 'kanban') ? '' : 'none';
    if (legendEl) legendEl.style.display = (mode === 'list')   ? '' : 'none';
    if (listBtn)   listBtn.classList.toggle('active',   mode === 'list');
    if (kanbanBtn) kanbanBtn.classList.toggle('active', mode === 'kanban');
    if (mode === 'kanban') renderTurnKanban();
  }

  function getTurnKanbanStage(p) {
    if (p.isCompleted) return 'Completed';
    if (p.isUpcoming)  return 'Upcoming';
    if (p.isOnRadar)   return 'Move-Out';
    // Confirmed active: bucket by stage progression
    // PIPE_STAGES indices: 0=upcoming 1=moveout 2=inspection 3=wo_created 4=est_requested 5=est_received 6=assigned 7=work_done
    if (p.currentStageIdx >= 5) return 'Approved & Working';
    return 'Inspecting & Bidding';
  }

  function renderTurnKanban() {
    var container = $('#turnKanban');
    if (!container) return;

    var search = ($('#turnPipeSearch') ? $('#turnPipeSearch').value : '').toLowerCase().trim();
    var group = currentTurnPipeGroup;

    var entries = TURN_PIPE_DATA.filter(function(p) {
      if (p.isClosed) return false;
      if (!isInPropertyGroup(p.propertyId, p.property, currentPropertyGroup)) return false;
      if (group && p.property !== group) return false;
      if (search) {
        var hay = (p.unit + ' ' + p.property + ' ' + (p.tenant || '')).toLowerCase();
        if (hay.indexOf(search) === -1) return false;
      }
      return true;
    });

    var stageOrder = ['Upcoming', 'Move-Out', 'Inspecting & Bidding', 'Approved & Working', 'Completed'];
    var stageAccents = {
      'Upcoming':            'var(--info)',
      'Move-Out':            'var(--warning)',
      'Inspecting & Bidding':'var(--accent)',
      'Approved & Working':  'var(--purple)',
      'Completed':           'var(--success)'
    };

    var buckets = {};
    stageOrder.forEach(function(s) { buckets[s] = []; });
    entries.forEach(function(p) { buckets[getTurnKanbanStage(p)].push(p); });

    if (entries.length === 0) {
      container.innerHTML = '<div style="padding:28px;text-align:center;color:var(--text-muted);font-size:13px"><i class="fas fa-columns" style="font-size:22px;display:block;margin-bottom:10px"></i>No turns match current filter</div>';
      return;
    }

    var html = '';
    stageOrder.forEach(function(stageName) {
      var cards = buckets[stageName];
      var accent = stageAccents[stageName];
      html += '<div class="kanban-col">';
      html += '<div class="kanban-col-head" style="border-top:2px solid ' + accent + '">';
      html += '<span class="kanban-col-title">' + escapeHtml(stageName) + '</span>';
      html += '<span class="kanban-col-count" style="background:color-mix(in srgb,' + accent + ' 15%,transparent);color:' + accent + '">' + cards.length + '</span>';
      html += '</div>';
      html += '<div class="kanban-col-body">';

      if (cards.length === 0) {
        html += '<div style="padding:12px 8px;text-align:center;font-size:11px;color:var(--text-muted);font-style:italic">No units in this stage</div>';
      } else {
        cards.forEach(function(p) {
          var cardClass = 'turn-kcard';
          if      (p.isCompleted) cardClass += ' t-done';
          else if (p.isUpcoming)  cardClass += ' t-upcoming';
          else if (p.isOnRadar)   cardClass += ' t-radar';
          else if (p.isStalled)   cardClass += ' t-stalled';

          var woCount = p.matchingWOs ? p.matchingWOs.length : 0;
          var elapsed = p.elapsed || 0;
          var elapsedClass = (!p.isUpcoming && !p.isCompleted && elapsed > 14) ? 't-danger' : '';
          var elapsedLabel = p.isUpcoming
            ? (Math.abs(elapsed) + 'd until MO')
            : (p.isCompleted ? (elapsed + 'd total') : (elapsed + 'd elapsed'));

          html += '<div class="' + escapeHtml(cardClass) + '" data-pipeid="' + escapeHtml(p.id) + '">';
          html += '<div class="tkc-unit">' + escapeHtml(p.unit || '—') + '</div>';
          html += '<div class="tkc-prop"><i class="fas fa-building" style="font-size:9px;margin-right:3px"></i>' + escapeHtml(p.property || '—') + '</div>';
          html += '<div class="tkc-metrics">';
          html += '<span><i class="fas fa-wrench" style="margin-right:3px"></i>' + woCount + ' WO' + (woCount !== 1 ? 's' : '') + '</span>';
          html += '<span class="' + elapsedClass + '"><i class="fas fa-clock" style="margin-right:3px"></i>' + escapeHtml(elapsedLabel) + '</span>';
          html += '</div>';
          html += '</div>';
        });
      }

      html += '</div></div>';
    });

    container.innerHTML = html;
    // Click delegation is wired once at startup in wireUpUI — no per-render attachment needed.
  }

  function openTurnDetailModal(turnId) {
    var p = TURN_PIPE_DATA.find(function(x) { return x.id === turnId; });
    if (!p) return;

    var titleEl  = $('#tdmUnit');
    var propEl   = $('#tdmProperty');
    var stageEl  = $('#tdmStage');
    var moEl     = $('#tdmMoveOut');
    var miEl     = $('#tdmMoveIn');
    var elEl     = $('#tdmElapsed');
    var woBody   = $('#tdmWOBody');
    if (!titleEl || !woBody) return;

    titleEl.textContent = p.unit || '—';
    propEl.textContent  = p.property || '—';
    stageEl.textContent = getTurnKanbanStage(p);
    moEl.textContent    = p.moveOut        ? formatDate(p.moveOut)        : '—';
    miEl.textContent    = p.expectedMoveIn ? formatDate(p.expectedMoveIn) : (p.isCompleted ? 'Completed' : 'Pending');
    elEl.textContent    = p.isUpcoming
      ? (Math.abs(p.elapsed) + 'd until move-out')
      : (p.elapsed + 'd elapsed');

    if (!p.matchingWOs || p.matchingWOs.length === 0) {
      woBody.innerHTML = '<tr><td colspan="5" style="text-align:center;padding:14px;color:var(--text-muted)">No work orders linked to this turn</td></tr>';
    } else {
      var rows = '';
      p.matchingWOs.forEach(function(wo) {
        var statusCls = 'tag ' + String(wo.status || 'linked').toLowerCase().replace(/\s+/g, '-');
        var desc      = String(wo.description || '').substring(0, 55);
        var woLink    = appfolioUrl('work_order', wo.id);
        rows += '<tr>';
        rows += '<td><span style="font-family:var(--font-mono);font-size:11px;font-weight:700;color:var(--accent)">#' + escapeHtml(String(wo.id || '')) + '</span></td>';
        rows += '<td style="font-size:11px">' + escapeHtml(desc || '—') + '</td>';
        rows += '<td><span class="' + escapeHtml(statusCls) + '">' + escapeHtml(wo.status || '—') + '</span></td>';
        rows += '<td style="font-size:11px">' + escapeHtml(wo.vendor || '—') + '</td>';
        rows += '<td>' + (woLink ? '<a href="' + escapeHtml(woLink) + '" target="_blank" rel="noopener noreferrer" class="action-btn" style="font-size:10px;padding:2px 6px;text-decoration:none" onclick="event.stopPropagation()"><i class="fas fa-external-link-alt"></i></a>' : '') + '</td>';
        rows += '</tr>';
      });
      woBody.innerHTML = rows;
    }

    openModal('turnDetailModal');
  }

  /* =================================================================
     INSPECTIONS — Enhanced with KPIs + Turn-linking
   ================================================================= */
function renderInspections(search) {
  var body = $('#inspBody');
  if (!body) return;

  var statusFilter = $('#inspStatusFilter') ? $('#inspStatusFilter').value : 'all';
  var today = new Date();

  // Guard against stale IndexedDB cache containing pre-AppFolio artifacts
  var validInspections = INSPECTIONS.filter(function(r) {
    if (!isInPropertyGroup(r.propertyId, r.propertyName, currentPropertyGroup)) return false;
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

function normalizeVendorDate(value) {
  var raw = String(value || '').trim();
  if (!raw) return null;
  var d = new Date(raw);
  return isNaN(d.getTime()) ? null : d;
}

function summarizeVendorGovernance(v) {
  var now = new Date();
  var expiredDocs = [];

  var insuranceExp = normalizeVendorDate(v.insurance);
  if (insuranceExp && insuranceExp < now) expiredDocs.push('Liability Insurance');

  var autoInsuranceExp = normalizeVendorDate(v.autoInsurance);
  if (autoInsuranceExp && autoInsuranceExp < now) expiredDocs.push('Auto Insurance');

  var workersCompExp = normalizeVendorDate(v.workersComp);
  if (workersCompExp && workersCompExp < now) expiredDocs.push('Workers Comp');

  var licenseExp = normalizeVendorDate(v.licenseExpires || v.license || v.contractorLicenseExpires);
  if (licenseExp && licenseExp < now) expiredDocs.push('License');

  var manualDoNotUse = v.doNotUse === true || String(v.doNotUse || '').toLowerCase() === 'true';
  return {
    expiredDocs: expiredDocs,
    expired: expiredDocs.length > 0,
    doNotUse: manualDoNotUse || expiredDocs.length > 0,
  };
}

/* Vendor governance combines explicit do-not-use flags with expired compliance docs. */
var _currentSelectedVendorId = null;

function isVendorDoNotUse(v) {
  return summarizeVendorGovernance(v).doNotUse;
}

function openVendorModal(vendorId) {
  _currentSelectedVendorId = String(vendorId || '');
  var v = VENDORS.find(function(vn) { return String(vn.id) === _currentSelectedVendorId; });
  if (!v) return;

  var today = new Date();
  var governance = summarizeVendorGovernance(v);
  var insDate = normalizeVendorDate(v.insurance);
  var wcDate = normalizeVendorDate(v.workersComp);
  var insExpired = !!(insDate && insDate < today);
  var wcExpired  = !!(wcDate && wcDate < today);
  var isDNU = governance.doNotUse;

  var nameEl   = $('#vdmName');
  var tradeEl  = $('#vdmTrade');
  var badgeEl  = $('#vdmStatusBadge');
  var warnEl   = $('#vdmWarning');
  var docsEl   = $('#vdmExpiredDocs');
  var phoneEl  = $('#vdmPhone');
  var emailEl  = $('#vdmEmail');
  var insEl    = $('#vdmInsurance');
  var wcEl     = $('#vdmWorkersComp');
  var afLink   = $('#vdmAfLink');

  if (nameEl)  nameEl.textContent  = v.name || '—';
  if (tradeEl) tradeEl.textContent = getVendorTradeCategory(v);
  if (phoneEl) phoneEl.textContent = v.phone || '—';
  if (emailEl) emailEl.textContent = v.email || '—';

  if (insEl) {
    insEl.textContent = v.insurance || 'Not on file';
    insEl.style.color = insExpired ? 'var(--danger)' : (v.insurance && daysBetween(today, new Date(v.insurance)) <= 60 ? 'var(--warning)' : 'var(--text-secondary)');
  }
  if (wcEl) {
    wcEl.textContent = v.workersComp || 'Not on file';
    wcEl.style.color = wcExpired ? 'var(--danger)' : (v.workersComp && daysBetween(today, new Date(v.workersComp)) <= 60 ? 'var(--warning)' : 'var(--text-secondary)');
  }

  if (badgeEl) {
    badgeEl.textContent = isDNU ? 'DO NOT USE' : 'Active';
    badgeEl.style.background = isDNU ? 'rgba(211,47,47,.15)' : 'rgba(46,125,50,.15)';
    badgeEl.style.color = isDNU ? 'var(--danger)' : 'var(--success)';
    badgeEl.style.borderColor = isDNU ? 'var(--danger)' : 'var(--success)';
  }

  if (warnEl)  warnEl.style.display = isDNU ? '' : 'none';
  if (docsEl)  docsEl.textContent = governance.expired
    ? governance.expiredDocs.join(' & ') + ' expired'
    : 'Flagged as Do Not Use';

  var afUrl = appfolioUrl('vendor', v.id);
  if (afLink) {
    afLink.href = afUrl || '#';
    afLink.style.display = afUrl ? '' : 'none';
  }

  openModal('vendorDetailModal');
}

function navigateToBillsForVendor(vendorRef) {
  // Switch to billing tab
  var billingTab = document.querySelector('.nav-tab[data-tab="billing"]');
  if (billingTab) billingTab.click();
  var ref = String(vendorRef || '').trim();
  var byId = VENDORS.find(function(v) { return String(v.id || '').trim() === ref; });
  var byName = byId ? byId : VENDORS.find(function(v) { return String(v.name || '').trim().toLowerCase() === ref.toLowerCase(); });
  var vendorId = byName ? String(byName.id || '').trim() : ref;
  // Use the vendor route filter for server-side scoping
  var filterType = document.getElementById('billing-filter-type');
  var filterInput = document.getElementById('billing-filter-input');
  if (filterType) filterType.value = 'bills_by_vendor';
  if (filterInput) filterInput.value = vendorId || '';
  _billingRouteAction = 'bills_by_vendor';
  _billingRouteFilterValue = vendorId || '';
  _billsPage = 0;
  setTimeout(function() {
    loadBillingPage({ resetPage: true });
  }, 150);
}

function navigateToBillsForWorkOrder(workOrderRef) {
  var billingTab = document.querySelector('.nav-tab[data-tab="billing"]');
  if (billingTab) billingTab.click();
  var billSearch = $('#billSearch');
  if (billSearch) {
    billSearch.value = String(workOrderRef || '');
    _billsPage = 0;
    renderBillsSection();
  }
}

function navigateToOpenWOsForVendor(vendorName) {
  var woTab = document.querySelector('.nav-tab[data-tab="workorders"]');
  if (woTab) woTab.click();
  // Switch to active sub-panel
  var activeSubBtn = document.querySelector('[data-wo-subtab="active"]');
  if (activeSubBtn) activeSubBtn.click();
  // Apply vendor filter
  currentWOVendor = vendorName;
  var vendorSel = $('#woVendorFilter');
  if (vendorSel) vendorSel.value = vendorName;
  renderWorkOrders();
}

function navigateToCompletedWOsForVendor(vendorName) {
  var woTab = document.querySelector('.nav-tab[data-tab="workorders"]');
  if (woTab) woTab.click();
  // Switch to completed sub-panel
  var completedSubBtn = document.querySelector('[data-wo-subtab="completed"]');
  if (completedSubBtn) completedSubBtn.click();
  // Set vendor filter so renderCompletedWOHistorySection() respects it
  currentWOVendor = vendorName;
  var vendorSel = $('#woVendorFilter');
  if (vendorSel) vendorSel.value = vendorName;
  // Trigger search if history not yet loaded, otherwise re-render
  var fetchBtn = $('#btnFetchCompletedHistory');
  if (fetchBtn && !completedWOHistoryRows.length) {
    fetchBtn.click();
  } else {
    completedWOHistoryPage = 0;
    renderCompletedWOHistorySection();
  }
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

  var searchText = (search || '').trim();
  var searchLower = searchText.toLowerCase();
  var baseFiltered = VENDORS.filter(function(v) {
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
    var ed = normalizeVendorDate(v.insurance);
    var exp = ed ? ed < today : false;
    var due = ed ? daysBetween(today, ed) : 999;
    var wrn = !exp && due <= 60;
    var cRes = resolveVendorCompliance(v);
    var governance = summarizeVendorGovernance(v);
    var cc = '';
    if (cRes.compliant && cRes.isManual) { cc = 'manual-compliant'; }
    else if (governance.expired) { cc = 'expired'; }
    else if (wrn) { cc = 'warn'; }
    var vCat = getVendorCategory(v.id);
    var catBadgeCls = vendorCatClass(vCat);
    var afUrl = appfolioUrl('vendor', v.id);

    html += '<div class="vendor-card vendor-card-compact ' + cc + '" data-vendorid="' + escapeHtml(String(v.id)) + '" data-vendor-initial="' + getVendorInitial(v.name) + '" style="cursor:pointer">';
    html += '<div class="vendor-card-head">';
    html += '<div class="vendor-name">' + escapeHtml(v.name) + '</div>';
    html += '<div class="vendor-governance-badges">';
    if (governance.expired) {
      html += '<span class="vendor-governance-badge expired">Expired</span>';
    }
    if (governance.doNotUse) {
      html += '<span class="vendor-governance-badge blocked">Do not use</span>';
    }
    html += '</div>';
    html += '<div class="vendor-id"><span><i class="fas fa-fingerprint"></i> ' + escapeHtml(String(v.id)) + '</span>';
    html += '<span class="vendor-category-badge ' + catBadgeCls + '">' + escapeHtml(vCat || 'Uncategorized') + '</span>';
    html += '</div>';
    html += '</div>';
    if (v.trades) html += '<div class="vendor-trades">' + escapeHtml(v.trades) + '</div>';
    html += '<div class="vendor-row vendor-row-compact"><span class="vendor-row-label">Category</span><span class="vendor-row-value">' + escapeHtml(vCat || 'Uncategorized') + '</span></div>';
    var vTradeCat = getVendorTradeCategory(v);
    html += '<div class="vendor-row vendor-row-compact"><span class="vendor-row-label">Trade Cat.</span><span class="vendor-row-value">' + escapeHtml(vTradeCat || 'Uncategorized') + '</span></div>';
    var compLabel = cRes.compliant ? 'Compliant' : (v.compliantStatus || 'Non-Compliant');
    html += '<div class="vendor-row vendor-row-compact" style="align-items:center"><span class="vendor-row-label">Compliance</span>';
    html += '<span class="vendor-row-value" style="font-weight:600;color:' + (cRes.compliant ? 'var(--success)' : 'var(--danger)') + '"><i class="fas ' + (cRes.compliant ? 'fa-check-circle' : 'fa-times-circle') + '"></i> ' + escapeHtml(compLabel) + (cRes.isManual ? ' (manual)' : '') + '</span></div>';
    if (v.phone || v.email) {
      html += '<div class="vendor-row vendor-row-compact"><span class="vendor-row-label">Contact</span><span class="vendor-row-value vendor-contact-compact">' + escapeHtml(v.phone || v.email || '\u2014') + (v.phone && v.email ? ' \u2022 ' + escapeHtml(v.email) : '') + '</span></div>';
    }
    if (v.address) {
      html += '<div class="vendor-row vendor-row-compact"><span class="vendor-row-label">Address</span><span class="vendor-row-value" style="font-size:10px">' + escapeHtml(v.address) + '</span></div>';
    }
    if (v.insurance) {
      html += '<div class="vendor-row vendor-row-compact"><span class="vendor-row-label">Liability Ins.</span><span class="vendor-row-value" style="font-family:var(--font-mono);color:' + (exp ? 'var(--danger)' : wrn ? 'var(--warning)' : 'var(--text-secondary)') + '">' + escapeHtml(v.insurance) + (exp ? ' \u26a0\ufe0f EXPIRED' : wrn ? ' (' + due + 'd)' : '') + '</span></div>';
    }
    if (v.autoInsurance) {
      var aed = new Date(v.autoInsurance); var aexp = aed < today; var adue = daysBetween(today, aed);
      html += '<div class="vendor-row vendor-row-compact"><span class="vendor-row-label">Auto Ins.</span><span class="vendor-row-value" style="font-family:var(--font-mono);color:' + (aexp ? 'var(--danger)' : adue <= 60 ? 'var(--warning)' : 'var(--text-secondary)') + '">' + escapeHtml(v.autoInsurance) + (aexp ? ' \u26a0\ufe0f EXPIRED' : adue <= 60 ? ' (' + adue + 'd)' : '') + '</span></div>';
    }
    if (v.workersComp) {
      var wced = new Date(v.workersComp); var wcexp = wced < today; var wcdue = daysBetween(today, wced);
      html += '<div class="vendor-row vendor-row-compact"><span class="vendor-row-label">Workers Comp</span><span class="vendor-row-value" style="font-family:var(--font-mono);color:' + (wcexp ? 'var(--danger)' : wcdue <= 60 ? 'var(--warning)' : 'var(--text-secondary)') + '">' + escapeHtml(v.workersComp) + (wcexp ? ' \u26a0\ufe0f EXPIRED' : wcdue <= 60 ? ' (' + wcdue + 'd)' : '') + '</span></div>';
    }
    if (governance.doNotUse) {
      html += '<div class="vendor-row vendor-row-compact" style="color:var(--danger);font-weight:700"><span class="vendor-row-label" style="color:var(--danger)">⛔ Status</span><span class="vendor-row-value">DO NOT USE</span></div>';
    }
    html += '<div class="vendor-row vendor-row-compact" style="align-items:center;gap:6px">' +
      '<span class="vendor-row-label">Type</span>' +
      '<select class="vendor-cat-select" data-vid="' + escapeHtml(String(v.id)) + '">' +
        VENDOR_CATEGORIES.map(function(cat) {
          return '<option value="' + escapeHtml(cat) + '"' + ((vCat || 'Uncategorized') === cat ? ' selected' : '') + '>' + escapeHtml(cat) + '</option>';
        }).join('') +
      '</select>' +
    '</div>';
    html += '<div class="vendor-row vendor-row-compact" style="align-items:center;gap:6px">' +
      '<span class="vendor-row-label">Trade</span>' +
      '<select class="vendor-trade-cat-select" data-vid="' + escapeHtml(String(v.id)) + '">' +
        VENDOR_TRADE_CATEGORIES.map(function(tc) {
          return '<option value="' + escapeHtml(tc) + '"' + (vTradeCat === tc ? ' selected' : '') + '>' + escapeHtml(tc) + '</option>';
        }).join('') +
      '</select>' +
    '</div>';
    if (v.tags) {
      html += '<div class="vendor-row vendor-row-compact"><span class="vendor-row-label">Tags</span><span class="vendor-row-value" style="font-size:10px">' + escapeHtml(v.tags) + '</span></div>';
    }
    if (afUrl) {
      html += '<div class="vendor-row vendor-row-compact"><span class="vendor-row-label">AppFolio</span><span class="vendor-row-value" style="color:var(--text-muted)">Open from vendor details modal</span></div>';
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
    var governance = summarizeVendorGovernance(v);
    var disabledAttr = governance.doNotUse ? ' disabled' : '';
    var statusSuffix = governance.doNotUse
      ? ' — Do not use'
      : (governance.expired ? ' — Expired' : '');
    opts.push('<option value="' + escapeHtml(String(v.id)) + '"' + disabledAttr + '>' + escapeHtml(v.name + statusSuffix) + '</option>');
  });

  if (selectedId && !visible.some(function(v) { return String(v.id) === selectedId; })) {
    var selectedVendor = VENDORS.find(function(v) { return String(v.id) === selectedId; });
    if (selectedVendor) {
      var selectedGovernance = summarizeVendorGovernance(selectedVendor);
      var selectedDisabled = selectedGovernance.doNotUse ? ' disabled' : '';
      opts.splice(1, 0, '<option value="' + escapeHtml(String(selectedVendor.id)) + '"' + selectedDisabled + '>' + escapeHtml(selectedVendor.name) + ' (selected)</option>');
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

var ROUTING_DEFAULT_CAPABILITIES = [
  { trade: 'HVAC', keywords: ['hvac', 'ac', 'air', 'thermostat', 'filter', 'condenser', 'compressor', 'heat'] },
  { trade: 'Plumbing', keywords: ['plumb', 'toilet', 'faucet', 'sink', 'drain', 'pipe', 'leak', 'water heater', 'shower'] },
  { trade: 'Electrical', keywords: ['electrical', 'outlet', 'switch', 'breaker', 'circuit', 'wiring', 'fixture'] },
  { trade: 'Appliances', keywords: ['appliance', 'refrigerator', 'dishwasher', 'washer', 'dryer', 'oven', 'stove', 'microwave'] },
  { trade: 'Painting', keywords: ['paint', 'painting', 'touch up', 'touch-up', 'repaint'] },
  { trade: 'Landscaping', keywords: ['landscape', 'lawn', 'grass', 'sprinkler', 'irrigation', 'tree', 'yard'] },
  { trade: 'General', keywords: ['handyman', 'maintenance', 'minor repair', 'caulk', 'hinge', 'door handle'] }
];

function getRoutingPropertyGroup(wo) {
  if (!wo) return '';
  if (wo.propertyId && _idToGroups[String(wo.propertyId)] && _idToGroups[String(wo.propertyId)].length) {
    return _idToGroups[String(wo.propertyId)][0] || '';
  }
  if (wo.propertyName) {
    var byName = _nameToGroups[String(wo.propertyName).trim().toLowerCase()];
    if (byName && byName.length) return byName[0] || '';
  }
  var prop = PROPERTIES.find(function(p) {
    if (wo.propertyId && String(p.id) === String(wo.propertyId)) return true;
    if (wo.propertyName && p.name && p.name.trim().toLowerCase() === String(wo.propertyName).trim().toLowerCase()) return true;
    return false;
  });
  return (prop && (prop.portfolio || prop.groupName || prop.group || '')) || '';
}

function getRoutingPmForGroup(groupName) {
  if (!groupName) return 'Unmapped PM';
  return ROUTING_PM_MAP[groupName] || groupName;
}

function getRoutingVendorCategory(wo) {
  if (!wo) return '';
  var cat = wo.vendorId ? (getVendorCategory(wo.vendorId) || '') : '';
  if (!cat) {
    var byName = VENDORS.find(function(v) { return String(v.name || '').trim().toLowerCase() === String(wo.vendorName || '').trim().toLowerCase(); });
    if (byName) cat = getVendorCategory(byName.id) || '';
  }
  return cat || 'Vendor';
}

function getRoutingConfidenceRank(level) {
  var l = String(level || '').toLowerCase();
  if (l === 'high') return 3;
  if (l === 'medium') return 2;
  return 1;
}

function computeRoutingConfidence(matchCount) {
  var count = Number(matchCount || 0);
  var highThreshold = Math.max(2, Number(ROUTING_SETTINGS.highThreshold || 2));
  if (count >= highThreshold) return 'high';
  if (count >= 1) return 'medium';
  return 'low';
}

function detectRoutingEventsFromLoadedWorkOrders() {
  var activeCaps = ROUTING_CAPABILITIES.filter(function(c) { return Number(c.active) === 1; });
  if (!activeCaps.length) return [];

  var out = [];
  WORK_ORDERS.forEach(function(wo) {
    var st = String(wo.status || '').toLowerCase();
    if (st.indexOf('cancel') !== -1 || st.indexOf('complete') !== -1) return;
    if (!wo.vendorId && !wo.vendorName) return;

    var vCat = getRoutingVendorCategory(wo);
    if (vCat === 'Employee' || vCat === 'In-House Tech') return;

    var text = (String(wo.description || '') + ' ' + String(wo.type || '') + ' ' + String(wo.priority || '')).toLowerCase();
    var matches = [];
    activeCaps.forEach(function(cap) {
      var kws = [];
      if (Array.isArray(cap.keywords)) kws = cap.keywords;
      else if (typeof cap.keywords === 'string') {
        try { kws = JSON.parse(cap.keywords); } catch (e) { kws = []; }
      }
      var tradeHit = text.indexOf(String(cap.trade || '').toLowerCase()) !== -1;
      var keywordHit = kws.some(function(kw) { return text.indexOf(String(kw || '').toLowerCase()) !== -1; });
      if (tradeHit || keywordHit) matches.push(String(cap.trade || 'General'));
    });
    if (!matches.length) return;

    var confidence = computeRoutingConfidence(matches.length);
    if (getRoutingConfidenceRank(confidence) < getRoutingConfidenceRank(ROUTING_SETTINGS.minConfidence)) {
      return;
    }

    var groupName = getRoutingPropertyGroup(wo);
    var pmName = getRoutingPmForGroup(groupName);

    out.push({
      wo_uuid: String(wo.uuid || ''),
      wo_number: String(wo.id || ''),
      property_id: String(wo.propertyId || ''),
      property_name: String(wo.propertyName || ''),
      unit_name: String(wo.unit || ''),
      property_group: String(groupName || ''),
      pm_name: String(pmName || 'Unmapped PM'),
      vendor_id: String(wo.vendorId || ''),
      vendor_name: String(wo.vendorName || ''),
      vendor_category: String(vCat || 'Vendor'),
      wo_status: String(wo.status || ''),
      wo_priority: String(wo.priority || ''),
      wo_created_at: String(wo.created || ''),
      description: String(wo.description || ''),
      matched_trade: matches.join(', '),
      confidence: confidence,
      source: 'client_scan'
    });
  });

  return out;
}

async function loadRoutingCapabilities() {
  try {
    var data = await proxyAction('routing_monitor', { op: 'capabilities' });
    ROUTING_CAPABILITIES = data.results || [];

    if (!ROUTING_CAPABILITIES.length) {
      for (var i = 0; i < ROUTING_DEFAULT_CAPABILITIES.length; i++) {
        var cap = ROUTING_DEFAULT_CAPABILITIES[i];
        await proxyPost('routing_monitor', {
          op: 'capability_upsert',
          trade: cap.trade,
          keywords: cap.keywords,
          active: 1
        });
      }
      var seeded = await proxyAction('routing_monitor', { op: 'capabilities' });
      ROUTING_CAPABILITIES = seeded.results || [];
    }

    renderRoutingCapabilities();
  } catch (err) {
    ROUTING_CAPABILITIES = [];
    renderRoutingCapabilities();
    showToast('Routing monitor unavailable: ' + ((err && err.message) ? err.message : String(err)), { kind: 'warning', duration: 4500 });
  }
}

async function loadRoutingPmMap() {
  try {
    var data = await proxyAction('routing_monitor', { op: 'pm_map' });
    ROUTING_PM_MAP = {};
    (data.results || []).forEach(function(r) {
      ROUTING_PM_MAP[String(r.group_name)] = String(r.pm_name || '');
    });
  } catch (err) {
    ROUTING_PM_MAP = {};
    showToast('Routing PM map unavailable: ' + ((err && err.message) ? err.message : String(err)), { kind: 'warning', duration: 4500 });
  }
  renderRoutingPmMapEditor();
}

async function loadRoutingEventsAndStats() {
  var status = $('#routingStatusFilter') ? $('#routingStatusFilter').value : 'pending';
  var pm = $('#routingPmFilter') ? $('#routingPmFilter').value : 'all';
  var days = $('#routingDaysFilter') ? $('#routingDaysFilter').value : '30';

  try {
    var evData = await proxyAction('routing_monitor', {
      op: 'events',
      status: status,
      pm: pm,
      days: days,
      limit: '500'
    });
    ROUTING_EVENTS = evData.results || [];

    var stData = await proxyAction('routing_monitor', { op: 'pm_stats', days: days });
    ROUTING_PM_STATS = stData.results || [];
  } catch (err) {
    ROUTING_EVENTS = [];
    ROUTING_PM_STATS = [];
    showToast('Routing monitor unavailable: ' + ((err && err.message) ? err.message : String(err)), { kind: 'warning', duration: 4500 });
  }

  renderRoutingPmFilter();
  renderRoutingScorecard();
  renderRoutingEventsTable($('#routingSearch') ? $('#routingSearch').value : '');
  renderRoutingKpis();
}

function getRoutingTradeFilterNormalized() {
  var tradeFilter = ($('#routingTradeFilter') && $('#routingTradeFilter').value) || 'all';
  return String(tradeFilter || 'all').trim().toLowerCase();
}

function getRoutingEventTradeLabel(r) {
  var direct = String((r && r.matched_trade) || '').trim();
  if (direct) return direct;
  var inferred = String((r && r.inferred_trade) || '').trim();
  if (inferred) return inferred;
  var arr = (r && Array.isArray(r.inferred_trades)) ? r.inferred_trades : [];
  if (arr.length) return String(arr[0] || '').trim();
  return '';
}

function getFilteredRoutingEvents() {
  var tradeFilterNorm = getRoutingTradeFilterNormalized();
  return (ROUTING_EVENTS || []).filter(function(r) {
    if (tradeFilterNorm === 'all') return true;
    var eventTradeNorm = getRoutingEventTradeLabel(r).toLowerCase();
    return !!eventTradeNorm && eventTradeNorm === tradeFilterNorm;
  });
}

function renderRoutingPmFilter() {
  var sel = $('#routingPmFilter');
  if (!sel) return;
  var current = sel.value || 'all';
  var seen = {};
  var opts = ['<option value="all">All PMs</option>'];
  ROUTING_PM_STATS.forEach(function(r) {
    var pm = String(r.pm_name || '').trim();
    if (!pm || seen[pm]) return;
    seen[pm] = true;
    opts.push('<option value="' + escapeHtml(pm) + '">' + escapeHtml(pm) + '</option>');
  });
  sel.innerHTML = opts.join('');
  sel.value = seen[current] ? current : 'all';

  // Also populate trade filter from loaded events
  var tradeSel = $('#routingTradeFilter');
  if (tradeSel) {
    var currentTrade = tradeSel.value || 'all';
    var tradesSeen = {};
    var tradeOpts = ['<option value="all">All Work Types</option>'];
    ROUTING_EVENTS.forEach(function(e) {
      var t = String(e.matched_trade || e.inferred_trade || '').trim();
      if (!t && Array.isArray(e.inferred_trades) && e.inferred_trades.length) {
        t = String(e.inferred_trades[0] || '').trim();
      }
      if (!t || tradesSeen[t]) return;
      tradesSeen[t] = true;
      tradeOpts.push('<option value="' + escapeHtml(t) + '">' + escapeHtml(t) + '</option>');
    });
    tradeSel.innerHTML = tradeOpts.join('');
    tradeSel.value = tradesSeen[currentTrade] ? currentTrade : 'all';
  }
}

function renderRoutingKpis() {
  var scopedEvents = getFilteredRoutingEvents();
  var total = scopedEvents.length;
  var pending = scopedEvents.filter(function(e) { return String(e.review_status || 'pending') === 'pending'; }).length;
  var high = scopedEvents.filter(function(e) { return String(e.confidence || '') === 'high'; }).length;
  var reviewed = scopedEvents.filter(function(e) { return String(e.review_status || 'pending') !== 'pending'; }).length;

  if ($('#routingKpiFlagged')) $('#routingKpiFlagged').textContent = String(total);
  if ($('#routingKpiPending')) $('#routingKpiPending').textContent = String(pending);
  if ($('#routingKpiHigh')) $('#routingKpiHigh').textContent = String(high);
  if ($('#routingKpiReviewed')) $('#routingKpiReviewed').textContent = String(reviewed);
  if ($('#routingKpiLast')) $('#routingKpiLast').textContent = _routingLastScanAt || '-';

  var badge = $('#routingBadge');
  if (badge) {
    badge.textContent = String(pending);
    badge.style.display = pending > 0 ? '' : 'none';
  }
}

function renderRoutingScorecard() {
  var body = $('#routingPmBody');
  if (!body) return;
  var scopedEvents = getFilteredRoutingEvents();
  if (!scopedEvents.length) {
    body.innerHTML = '<tr><td colspan="7" style="color:var(--text-muted)">No PM routing data yet for current filters.</td></tr>';
    return;
  }

  var byPm = {};
  scopedEvents.forEach(function(e) {
    var pmName = String(e.pm_name || 'Unmapped PM').trim() || 'Unmapped PM';
    if (!byPm[pmName]) {
      byPm[pmName] = {
        pm_name: pmName,
        total_flagged: 0,
        high_confidence: 0,
        pending: 0,
        approved_external: 0,
        reassigned: 0,
        last_flagged: ''
      };
    }
    var row = byPm[pmName];
    row.total_flagged += 1;
    if (String(e.confidence || '') === 'high') row.high_confidence += 1;
    var review = String(e.review_status || 'pending');
    if (review === 'pending') row.pending += 1;
    if (review === 'approved_external') row.approved_external += 1;
    if (review === 'reassign_inhouse') row.reassigned += 1;
    var detected = String(e.detected_at || '').trim();
    if (detected && (!row.last_flagged || detected > row.last_flagged)) row.last_flagged = detected;
  });

  var scoreRows = Object.keys(byPm).map(function(k) { return byPm[k]; });
  scoreRows.sort(function(a, b) {
    if (Number(b.total_flagged || 0) !== Number(a.total_flagged || 0)) {
      return Number(b.total_flagged || 0) - Number(a.total_flagged || 0);
    }
    return String(a.pm_name || '').localeCompare(String(b.pm_name || ''));
  });

  var html = '';
  scoreRows.forEach(function(r) {
    var last = r.last_flagged ? formatDate(r.last_flagged) : '-';
    var total = Number(r.total_flagged || 0);
    var potentialBypass = Number(r.pending || 0) + Number(r.approved_external || 0);
    var riskLabel = 'Low';
    if (potentialBypass >= 5 || (total > 0 && (potentialBypass / total) >= 0.6)) riskLabel = 'High';
    else if (potentialBypass >= 2 || (total > 0 && (potentialBypass / total) >= 0.35)) riskLabel = 'Medium';
    html += '<tr>' +
      '<td><strong>' + escapeHtml(String(r.pm_name || 'Unmapped PM')) + '</strong></td>' +
      '<td>' + escapeHtml(String(r.total_flagged || '0')) + '</td>' +
      '<td>' + escapeHtml(String(r.high_confidence || '0')) + '</td>' +
      '<td>' + escapeHtml(String(r.pending || '0')) + '</td>' +
      '<td>' + escapeHtml(String(r.approved_external || '0')) + '</td>' +
      '<td>' + escapeHtml(String(r.reassigned || '0')) + '</td>' +
      '<td>' + escapeHtml(String(potentialBypass)) + ' <span class="status-badge status-' + riskLabel.toLowerCase() + '">' + escapeHtml(riskLabel) + '</span></td>' +
      '<td style="font-family:var(--font-mono);font-size:11px">' + escapeHtml(last) + '</td>' +
      '</tr>';
  });
  body.innerHTML = html;
}

function renderRoutingEventsTable(query) {
  var body = $('#routingEventsBody');
  if (!body) return;
  var q = String(query || '').trim().toLowerCase();
  var tradeFilterNorm = getRoutingTradeFilterNormalized();

  var rows = ROUTING_EVENTS.filter(function(r) {
    if (tradeFilterNorm !== 'all') {
      var eventTradeNorm = getRoutingEventTradeLabel(r).toLowerCase();
      if (!eventTradeNorm || eventTradeNorm !== tradeFilterNorm) return false;
    }
    if (!q) return true;
    var hay = [r.wo_number, r.property_name, r.property_group, r.pm_name, r.vendor_name, getRoutingEventTradeLabel(r), r.description].join(' ').toLowerCase();
    return hay.indexOf(q) !== -1;
  });

  if ($('#routingSummary')) {
    $('#routingSummary').textContent = 'Showing ' + rows.length + ' of ' + ROUTING_EVENTS.length + ' flagged events';
  }

  var pageSize = Number(ROUTING_SETTINGS.pageSize || 25);
  if (pageSize !== 25 && pageSize !== 50 && pageSize !== 100) pageSize = 25;
  var totalRows = rows.length;
  var totalPages = Math.max(1, Math.ceil(totalRows / pageSize));
  ROUTING_PAGE = Math.min(Math.max(1, ROUTING_PAGE), totalPages);
  var startIdx = (ROUTING_PAGE - 1) * pageSize;
  var endIdx = startIdx + pageSize;
  var pagedRows = rows.slice(startIdx, endIdx);

  var pageLabel = $('#routingEventsPageLabel');
  if (pageLabel) {
    pageLabel.textContent = 'Page ' + ROUTING_PAGE + ' of ' + totalPages + ' • ' + totalRows + ' result' + (totalRows === 1 ? '' : 's');
  }
  var prevBtn = $('#btnRoutingPrevPage');
  if (prevBtn) prevBtn.disabled = ROUTING_PAGE <= 1;
  var nextBtn = $('#btnRoutingNextPage');
  if (nextBtn) nextBtn.disabled = ROUTING_PAGE >= totalPages;

  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="12" style="color:var(--text-muted)">No flagged routing events match current filters.</td></tr>';
    return;
  }

  var html = '';
  pagedRows.forEach(function(r) {
    var conf = String(r.confidence || 'medium');
    var rev = String(r.review_status || 'pending');
    var flaggedAt = r.detected_at ? timeAgo(r.detected_at) : '-';
    var matchLabel = String(getRoutingEventTradeLabel(r) || '-').trim();
    var riskInsight = '';
    if (rev === 'approved_external') {
      riskInsight = 'External approved; FLM bid path likely bypassed.';
    } else if (rev === 'pending') {
      riskInsight = conf === 'high'
        ? 'High-risk bypass candidate. Verify FLM was offered bid/work first.'
        : 'Review needed: confirm FLM had first-look opportunity.';
    } else if (rev === 'reassign_inhouse') {
      riskInsight = 'Recovered in-house routing path.';
    } else {
      riskInsight = 'Reviewed and dismissed.';
    }
    html += '<tr data-routing-id="' + escapeHtml(String(r.id)) + '">' +
      '<td style="font-family:var(--font-mono)">' + escapeHtml(String(r.wo_number || '').replace(/^.*(\d{4,})$/, '$1')) + '</td>' +
      '<td>' + escapeHtml(String(r.property_name || '-')) + '</td>' +
      '<td>' + escapeHtml(String(r.property_group || '-')) + '</td>' +
      '<td><strong>' + escapeHtml(String(r.pm_name || 'Unmapped PM')) + '</strong></td>' +
      '<td>' + escapeHtml(String(r.vendor_name || '-')) + '</td>' +
      '<td>' + escapeHtml(matchLabel) + '</td>' +
      '<td><span class="routing-badge ' + conf + '">' + escapeHtml(conf) + '</span></td>' +
      '<td>' + escapeHtml(String(r.wo_status || '-')) + '</td>' +
      '<td style="font-family:var(--font-mono);font-size:11px">' + escapeHtml(flaggedAt) + '</td>' +
      '<td class="routing-review ' + rev + '">' + escapeHtml(rev.replace('_', ' ')) + '</td>' +
      '<td style="max-width:260px;white-space:normal;line-height:1.35">' + escapeHtml(riskInsight) + '</td>' +
      '<td>' +
      '<button class="filter-btn" data-routing-review="approved_external" data-routing-id="' + escapeHtml(String(r.id)) + '" style="padding:4px 6px">Approve</button> ' +
      '<button class="filter-btn" data-routing-review="reassign_inhouse" data-routing-id="' + escapeHtml(String(r.id)) + '" style="padding:4px 6px;color:var(--success)">Reassign</button> ' +
      '<button class="filter-btn" data-routing-review="dismissed" data-routing-id="' + escapeHtml(String(r.id)) + '" style="padding:4px 6px;color:var(--text-muted)">Dismiss</button>' +
      '</td>' +
      '</tr>';
  });
  body.innerHTML = html;
}

function renderRoutingCapabilities() {
  var body = $('#routingCapabilitiesBody');
  if (!body) return;
  if (!ROUTING_CAPABILITIES.length) {
    body.innerHTML = '<tr><td colspan="4" style="color:var(--text-muted)">No capability rules configured.</td></tr>';
    return;
  }
  var html = '';
  ROUTING_CAPABILITIES.forEach(function(c) {
    var kws = [];
    if (Array.isArray(c.keywords)) kws = c.keywords;
    else if (typeof c.keywords === 'string') {
      try { kws = JSON.parse(c.keywords); } catch (e) { kws = []; }
    }
    html += '<tr>' +
      '<td><strong>' + escapeHtml(String(c.trade || '')) + '</strong></td>' +
      '<td style="font-size:11px;color:var(--text-secondary)">' + escapeHtml(kws.join(', ')) + '</td>' +
      '<td><label style="display:inline-flex;align-items:center;gap:6px"><input type="checkbox" data-routing-cap-toggle="' + escapeHtml(String(c.id)) + '"' + (Number(c.active) === 1 ? ' checked' : '') + '> ' + (Number(c.active) === 1 ? 'On' : 'Off') + '</label></td>' +
      '<td><button class="filter-btn" data-routing-cap-edit="' + escapeHtml(String(c.id)) + '" style="padding:4px 6px">Edit</button> <button class="filter-btn" data-routing-cap-del="' + escapeHtml(String(c.id)) + '" style="padding:4px 6px;color:var(--danger)">Delete</button></td>' +
      '</tr>';
  });
  body.innerHTML = html;
}

function buildRoutingWorkTypesEditorHtml() {
  var rows = '';
  if (!ROUTING_CAPABILITIES.length) {
    rows = '<tr><td colspan="4" style="text-align:center;color:var(--text-muted);padding:12px">No work types configured.</td></tr>';
  } else {
    ROUTING_CAPABILITIES.forEach(function(c) {
      var kws = [];
      if (Array.isArray(c.keywords)) kws = c.keywords;
      else if (typeof c.keywords === 'string') {
        try { kws = JSON.parse(c.keywords); } catch (e) { kws = []; }
      }
      rows += '<tr>' +
        '<td><strong>' + escapeHtml(String(c.trade || '')) + '</strong></td>' +
        '<td style="font-size:11px;color:var(--text-secondary)">' + escapeHtml(kws.join(', ')) + '</td>' +
        '<td><label style="display:inline-flex;align-items:center;gap:6px"><input type="checkbox" data-routing-cap-modal-toggle="' + escapeHtml(String(c.id)) + '"' + (Number(c.active) === 1 ? ' checked' : '') + '> ' + (Number(c.active) === 1 ? 'On' : 'Off') + '</label></td>' +
        '<td><button class="filter-btn" data-routing-cap-modal-edit="' + escapeHtml(String(c.id)) + '" style="padding:4px 6px">Edit</button> <button class="filter-btn" data-routing-cap-modal-del="' + escapeHtml(String(c.id)) + '" style="padding:4px 6px;color:var(--danger)">Delete</button></td>' +
        '</tr>';
    });
  }

  return '' +
    '<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:10px">' +
      '<button class="action-btn" id="btnRoutingCapModalAdd"><i class="fas fa-plus"></i> Add Work Type</button>' +
      '<button class="action-btn" id="btnRoutingCapModalRefresh"><i class="fas fa-sync-alt"></i> Refresh</button>' +
      '<span id="routingCapSyncMsg" style="font-size:11px;color:var(--text-muted)">Changes save directly to routing SQL tables.</span>' +
    '</div>' +
    '<div style="border:1px solid var(--border);border-radius:10px;overflow:hidden">' +
      '<table class="data-table" style="margin:0">' +
        '<thead><tr><th>Work Type</th><th>Keywords</th><th>Enabled</th><th>Actions</th></tr></thead>' +
        '<tbody id="routingCapabilitiesModalBody">' + rows + '</tbody>' +
      '</table>' +
    '</div>';
}

function openRoutingWorkTypesModal() {
  showItemDetail('Routing Work Type Rules', [
    { section: 'Rule Editor', icon: 'fa-sliders-h' },
    { label: 'Work Types', html: buildRoutingWorkTypesEditorHtml() }
  ]);

  requestAnimationFrame(function() {
    var mount = document.getElementById('routingCapabilitiesModalBody');
    var addBtn = document.getElementById('btnRoutingCapModalAdd');
    var refreshBtn = document.getElementById('btnRoutingCapModalRefresh');

    if (refreshBtn) {
      refreshBtn.addEventListener('click', async function() {
        await loadRoutingCapabilities();
        openRoutingWorkTypesModal();
      });
    }

    if (addBtn) {
      addBtn.addEventListener('click', async function() {
        var trade = await hmPrompt('Work type name (example: Locksmith):', '', { title: 'Add Work Type' });
        if (!trade) return;
        var kwRaw = await hmPrompt('Keywords (comma separated):', '', { title: 'Add Work Type' });
        var keywords = String(kwRaw || '').split(',').map(function(k) { return k.trim(); }).filter(Boolean);
        await proxyPost('routing_monitor', { op: 'capability_upsert', trade: trade.trim(), keywords: keywords, active: 1 });
        await loadRoutingCapabilities();
        await loadRoutingEventsAndStats();
        showToast('Work type \u201c' + trade.trim() + '\u201d added', { kind: 'success' });
        openRoutingWorkTypesModal();
      });
    }

    if (!mount) return;
    mount.addEventListener('change', async function(e) {
      var chk = e.target.closest('[data-routing-cap-modal-toggle]');
      if (!chk) return;
      var capId = Number(chk.getAttribute('data-routing-cap-modal-toggle') || '0');
      var cap = ROUTING_CAPABILITIES.find(function(c) { return Number(c.id) === capId; });
      if (!cap) return;
      var parsedKeywords = [];
      if (typeof cap.keywords === 'string') {
        try { parsedKeywords = JSON.parse(cap.keywords || '[]'); } catch (er) { parsedKeywords = []; }
      } else {
        parsedKeywords = cap.keywords || [];
      }
      await proxyPost('routing_monitor', {
        op: 'capability_upsert',
        id: cap.id,
        trade: cap.trade,
        keywords: parsedKeywords,
        active: chk.checked ? 1 : 0
      });
      await loadRoutingCapabilities();
      await loadRoutingEventsAndStats();
      showToast('Work type ' + (chk.checked ? 'enabled' : 'disabled'), { kind: 'success' });
      openRoutingWorkTypesModal();
    });

    mount.addEventListener('click', async function(e) {
      var editBtn = e.target.closest('[data-routing-cap-modal-edit]');
      var delBtn = e.target.closest('[data-routing-cap-modal-del]');

      if (editBtn) {
        var cid = Number(editBtn.getAttribute('data-routing-cap-modal-edit') || '0');
        var cur = ROUTING_CAPABILITIES.find(function(c) { return Number(c.id) === cid; });
        if (!cur) return;
        var trade = await hmPrompt('Work type name:', String(cur.trade || ''), { title: 'Edit Work Type' });
        if (!trade) return;
        var kws = [];
        try { kws = typeof cur.keywords === 'string' ? JSON.parse(cur.keywords) : (cur.keywords || []); } catch (er) { kws = []; }
        var kwRaw = await hmPrompt('Keywords (comma separated):', kws.join(', '), { title: 'Edit Work Type' });
        if (kwRaw === null) return;
        var nextKws = String(kwRaw).split(',').map(function(k) { return k.trim(); }).filter(Boolean);
        await proxyPost('routing_monitor', { op: 'capability_upsert', id: cur.id, trade: trade.trim(), keywords: nextKws, active: Number(cur.active) === 1 ? 1 : 0 });
        await loadRoutingCapabilities();
        await loadRoutingEventsAndStats();
        showToast('\u201c' + trade.trim() + '\u201d updated', { kind: 'success' });
        openRoutingWorkTypesModal();
        return;
      }

      if (delBtn) {
        var did = Number(delBtn.getAttribute('data-routing-cap-modal-del') || '0');
        if (!did) return;
        if (!await hmConfirm('Delete this work type rule?', { title: 'Delete Rule', okLabel: 'Delete', danger: true })) return;
        await proxyPost('routing_monitor', { op: 'capability_delete', id: did });
        await loadRoutingCapabilities();
        await loadRoutingEventsAndStats();
        showToast('Work type deleted', { kind: 'success' });
        openRoutingWorkTypesModal();
      }
    });
  });
}

function renderRoutingPmMapEditor() {
  var body = $('#routingPmMapBody');
  if (!body) return;

  var groups = [];
  (PROPERTY_GROUPS || []).forEach(function(g) {
    if (!g || !g.name) return;
    groups.push(String(g.name));
  });
  Object.keys(ROUTING_PM_MAP).forEach(function(g) {
    if (groups.indexOf(g) === -1) groups.push(g);
  });
  groups.sort(function(a, b) { return a.localeCompare(b); });

  if (!groups.length) {
    body.innerHTML = '<tr><td colspan="2" style="color:var(--text-muted)">No property groups available yet. Load property groups first.</td></tr>';
    return;
  }

  var html = '';
  groups.forEach(function(group) {
    var pm = ROUTING_PM_MAP[group] || group;
    html += '<tr>' +
      '<td>' + escapeHtml(group) + '</td>' +
      '<td><input class="form-input routing-pm-map-input" data-group-name="' + escapeHtml(group) + '" value="' + escapeHtml(pm) + '" style="min-width:200px"></td>' +
      '</tr>';
  });
  body.innerHTML = html;
}

function guessPmNameForGroup(groupName) {
  var group = String(groupName || '').trim();
  if (!group) return '';

  var byManager = {};
  (PROPERTIES || []).forEach(function(p) {
    var pName = String(p && p.name || '').trim().toLowerCase();
    var gCandidates = [];
    if (p && p.id && _idToGroups[String(p.id)]) gCandidates = gCandidates.concat(_idToGroups[String(p.id)]);
    if (pName && _nameToGroups[pName]) gCandidates = gCandidates.concat(_nameToGroups[pName]);
    var direct = [p && p.portfolio, p && p.portfolioName, p && p.groupName, p && p.propertyGroup, p && p.group]
      .map(function(v) { return String(v || '').trim(); })
      .filter(Boolean);
    gCandidates = gCandidates.concat(direct);
    var inGroup = gCandidates.some(function(g) { return String(g || '').trim().toLowerCase() === group.toLowerCase(); });
    if (!inGroup) return;

    var siteManager = String((p && p.siteManager) || '').trim();
    if (!siteManager) return;
    byManager[siteManager] = (byManager[siteManager] || 0) + 1;
  });

  var best = '';
  var bestCount = 0;
  Object.keys(byManager).forEach(function(name) {
    if (byManager[name] > bestCount) {
      best = name;
      bestCount = byManager[name];
    }
  });
  return best;
}

function autoFillRoutingPmMap() {
  var updated = 0;
  $$('.routing-pm-map-input').forEach(function(input) {
    var groupName = String(input.getAttribute('data-group-name') || '').trim();
    if (!groupName) return;
    var current = String(input.value || '').trim();
    if (current && current.toLowerCase() !== groupName.toLowerCase()) return;
    var guess = guessPmNameForGroup(groupName);
    if (guess) {
      input.value = guess;
      updated++;
    }
  });
  if (updated > 0) showToast('Auto-filled ' + updated + ' PM mapping row' + (updated === 1 ? '' : 's'));
  else showToast('No PM names inferred yet. Load properties/work orders first.', 'warning');
}

async function saveRoutingPmMapFromUi() {
  var rows = [];
  $$('.routing-pm-map-input').forEach(function(input) {
    var groupName = input.getAttribute('data-group-name') || '';
    var pmName = String(input.value || '').trim();
    if (groupName && pmName) rows.push({ group_name: groupName, pm_name: pmName });
  });
  if (!rows.length) {
    showToast('No PM map rows to save');
    return;
  }
  await proxyPost('routing_monitor', { op: 'pm_map_bulk', entries: rows });
  await loadRoutingPmMap();
  await loadRoutingEventsAndStats();
  showToast('PM map saved (' + rows.length + ' groups)');
}

async function runRoutingScan() {
  var btn = $('#btnRoutingScan');
  if (btn) { btn.disabled = true; btn.classList.add('spinning'); }
  try {
    if (!WORK_ORDERS.length) {
      showToast('No work orders loaded yet. Refresh data first.');
      return;
    }
    if (!ROUTING_CAPABILITIES.length) await loadRoutingCapabilities();
    if (!Object.keys(ROUTING_PM_MAP).length) await loadRoutingPmMap();

    var events = detectRoutingEventsFromLoadedWorkOrders();
    var resp = await proxyPost('routing_monitor', { op: 'scan', events: events });
    _routingLastScanAt = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    await loadRoutingEventsAndStats();
    showToast('Routing scan complete — ' + events.length + ' matched, ' + (resp.upserted || 0) + ' saved');
  } catch (e) {
    showToast('Routing scan failed: ' + (e.message || e));
  } finally {
    if (btn) { btn.disabled = false; btn.classList.remove('spinning'); }
  }
}

function ensureRoutingFilterDelegation() {
  if (_routingFiltersDelegated) return;
  var section = $('#sec-routing');
  if (!section) return;

  section.addEventListener('change', function(e) {
    var t = e.target;
    if (!t || !t.id) return;

    if (t.id === 'routingStatusFilter' || t.id === 'routingPmFilter' || t.id === 'routingDaysFilter') {
      ROUTING_PAGE = 1;
      loadRoutingEventsAndStats();
      return;
    }

    if (t.id === 'routingTradeFilter') {
      ROUTING_PAGE = 1;
      renderRoutingEventsTable($('#routingSearch') ? $('#routingSearch').value : '');
      return;
    }

    if (t.id === 'routingPageSize') {
      var nextSize = Number(t.value || '25');
      ROUTING_SETTINGS.pageSize = (nextSize === 50 || nextSize === 100) ? nextSize : 25;
      ROUTING_PAGE = 1;
      saveRoutingSettings();
      renderRoutingEventsTable($('#routingSearch') ? $('#routingSearch').value : '');
      return;
    }
  });

  section.addEventListener('input', function(e) {
    var t = e.target;
    if (!t || t.id !== 'routingSearch') return;
    if (_routingSearchDebounceTimer) clearTimeout(_routingSearchDebounceTimer);
    _routingSearchDebounceTimer = setTimeout(function() {
      ROUTING_PAGE = 1;
      renderRoutingEventsTable(t.value || '');
    }, CONFIG.DEBOUNCE_MS);
  });

  section.addEventListener('click', function(e) {
    var prevBtn = e.target.closest('#btnRoutingPrevPage');
    if (prevBtn) {
      if (ROUTING_PAGE <= 1) return;
      ROUTING_PAGE -= 1;
      renderRoutingEventsTable($('#routingSearch') ? $('#routingSearch').value : '');
      return;
    }

    var nextBtn = e.target.closest('#btnRoutingNextPage');
    if (nextBtn) {
      ROUTING_PAGE += 1;
      renderRoutingEventsTable($('#routingSearch') ? $('#routingSearch').value : '');
    }
  });

  _routingFiltersDelegated = true;
}

async function initRoutingMonitor() {
  if (!_routingInitDone) {
    _routingInitDone = true;

    ensureRoutingFilterDelegation();

    if ($('#btnRoutingRefresh')) {
      $('#btnRoutingRefresh').addEventListener('click', function() { loadRoutingEventsAndStats(); });
    }
    if ($('#btnRoutingScan')) {
      $('#btnRoutingScan').addEventListener('click', function() { runRoutingScan(); });
    }
    if ($('#btnRoutingCapToggle')) {
      $('#btnRoutingCapToggle').addEventListener('click', function() {
        openRoutingWorkTypesModal();
      });
    }
    if ($('#btnRoutingMapToggle')) {
      $('#btnRoutingMapToggle').addEventListener('click', function() {
        var wrap = $('#routingPmMapWrap');
        if (!wrap) return;
        wrap.style.display = wrap.style.display === 'none' ? '' : 'none';
      });
    }
    if ($('#btnRoutingMapSave')) {
      $('#btnRoutingMapSave').addEventListener('click', function() { saveRoutingPmMapFromUi(); });
    }
    if ($('#btnRoutingMapAutofill')) {
      $('#btnRoutingMapAutofill').addEventListener('click', function() { autoFillRoutingPmMap(); });
    }
    if ($('#routingMinConfidence')) {
      $('#routingMinConfidence').value = ROUTING_SETTINGS.minConfidence;
      $('#routingMinConfidence').addEventListener('change', function() {
        ROUTING_SETTINGS.minConfidence = String(this.value || 'medium');
        saveRoutingSettings();
      });
    }
    if ($('#routingHighThreshold')) {
      $('#routingHighThreshold').value = String(ROUTING_SETTINGS.highThreshold);
      $('#routingHighThreshold').addEventListener('change', function() {
        var next = Number(this.value || '2');
        ROUTING_SETTINGS.highThreshold = isNaN(next) ? 2 : Math.max(2, next);
        saveRoutingSettings();
      });
    }
    if ($('#routingPageSize')) $('#routingPageSize').value = String(ROUTING_SETTINGS.pageSize || 25);
    if ($('#btnRoutingCapAdd')) {
      $('#btnRoutingCapAdd').addEventListener('click', async function() {
        var trade = await hmPrompt('Work type name (example: Locksmith):', '', { title: 'Add Work Type' });
        if (!trade) return;
        var kwRaw = await hmPrompt('Keywords (comma separated):', '', { title: 'Add Work Type' });
        var keywords = String(kwRaw || '').split(',').map(function(k) { return k.trim(); }).filter(Boolean);
        await proxyPost('routing_monitor', { op: 'capability_upsert', trade: trade.trim(), keywords: keywords, active: 1 });
        await loadRoutingCapabilities();
      });
    }
    var eventsBody = $('#routingEventsBody');
    if (eventsBody) {
      eventsBody.addEventListener('click', async function(e) {
        var btn = e.target.closest('[data-routing-review]');
        if (!btn) return;
        var id = Number(btn.getAttribute('data-routing-id') || '0');
        var review = btn.getAttribute('data-routing-review') || 'pending';
        if (!id) return;
        await proxyPost('routing_monitor', { op: 'review', id: id, review_status: review });
        await loadRoutingEventsAndStats();
      });
    }
    var capBody = $('#routingCapabilitiesBody');
    if (capBody) {
      capBody.addEventListener('change', async function(e) {
        var chk = e.target.closest('[data-routing-cap-toggle]');
        if (!chk) return;
        var capId = Number(chk.getAttribute('data-routing-cap-toggle') || '0');
        var cap = ROUTING_CAPABILITIES.find(function(c) { return Number(c.id) === capId; });
        if (!cap) return;
        var parsedKeywords = [];
        if (typeof cap.keywords === 'string') {
          try { parsedKeywords = JSON.parse(cap.keywords || '[]'); } catch (er) { parsedKeywords = []; }
        } else {
          parsedKeywords = cap.keywords || [];
        }
        await proxyPost('routing_monitor', {
          op: 'capability_upsert',
          id: cap.id,
          trade: cap.trade,
          keywords: parsedKeywords,
          active: chk.checked ? 1 : 0
        });
        await loadRoutingCapabilities();
      });
      capBody.addEventListener('click', async function(e) {
        var editBtn = e.target.closest('[data-routing-cap-edit]');
        var delBtn = e.target.closest('[data-routing-cap-del]');
        if (editBtn) {
          var cid = Number(editBtn.getAttribute('data-routing-cap-edit') || '0');
          var cur = ROUTING_CAPABILITIES.find(function(c) { return Number(c.id) === cid; });
          if (!cur) return;
          var trade = await hmPrompt('Work type name:', String(cur.trade || ''), { title: 'Edit Work Type' });
          if (!trade) return;
          var kws = [];
          try { kws = typeof cur.keywords === 'string' ? JSON.parse(cur.keywords) : (cur.keywords || []); } catch (er) { kws = []; }
          var kwRaw = await hmPrompt('Keywords (comma separated):', kws.join(', '), { title: 'Edit Work Type' });
          if (kwRaw === null) return;
          var nextKws = String(kwRaw).split(',').map(function(k) { return k.trim(); }).filter(Boolean);
          await proxyPost('routing_monitor', { op: 'capability_upsert', id: cur.id, trade: trade.trim(), keywords: nextKws, active: Number(cur.active) === 1 ? 1 : 0 });
          await loadRoutingCapabilities();
          return;
        }
        if (delBtn) {
          var did = Number(delBtn.getAttribute('data-routing-cap-del') || '0');
          if (!did) return;
          if (!await hmConfirm('Delete this work type rule?', { title: 'Delete Rule', okLabel: 'Delete', danger: true })) return;
          await proxyPost('routing_monitor', { op: 'capability_delete', id: did });
          await loadRoutingCapabilities();
        }
      });
    }
  }

  await loadRoutingCapabilities();
  await loadRoutingPmMap();
  await loadRoutingEventsAndStats();
}

/* renderReconciliation — Removed (billing stripped to lighten payload) */

function renderTemplates() {
  var container = $('#templateGrid');
  if (!TEMPLATES.length) {
    container.innerHTML = emptyHtml('fa-clipboard-list', 'No note templates configured yet');
    return;
  }
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
  renderSystemHealthPanel();
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
  renderSystemHealthPanel();
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
        var columnName = head.getAttribute('data-column') || head.getAttribute('data-status') || '';
        if (expandedWOColumn === columnName) {
          expandedWOColumn = '';
          renderWorkOrders();
          board.scrollLeft = kanbanBoardScrollState.left || 0;
          board.scrollTop = kanbanBoardScrollState.top || 0;
          return;
        }
        kanbanBoardScrollState.left = board.scrollLeft;
        kanbanBoardScrollState.top = board.scrollTop;
        expandedWOColumn = columnName;
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

  // Turn pipeline — card clicks + stage advance buttons + close turn
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
      var addWoBtn = e.target.closest('[data-add-wo]');
      if (addWoBtn) {
        e.stopPropagation();
        var turnKey = addWoBtn.getAttribute('data-add-wo');
        if (isTurnActionLocked(turnKey)) {
          showToast('Turn is terminal/closed; link actions are locked', { kind: 'warning' });
          return;
        }
        // Use closest detail panel to find the input, avoiding CSS selector issues
        // with special characters (colons, etc.) in the turn key value.
        var detailPanel = addWoBtn.closest('.pipe-detail');
        var input = detailPanel ? detailPanel.querySelector('input[data-add-wo-input]') : null;
        var woId = input ? String(input.value || '').trim() : '';
        if (!woId) {
          showToast('Enter a WO number first');
          return;
        }
        linkTurnWorkOrder(turnKey, woId)
          .then(function() { showToast('WO #' + woId + ' linked to turn'); })
          .catch(function(err) { showToast('Link failed: ' + (err.message || err)); });
        return;
      }
      var unlinkBtn = e.target.closest('[data-unlink-wo]');
      if (unlinkBtn) {
        e.stopPropagation();
        var tk = unlinkBtn.getAttribute('data-unlink-wo');
        if (isTurnActionLocked(tk)) {
          showToast('Turn is terminal/closed; unlink actions are locked', { kind: 'warning' });
          return;
        }
        var woid = unlinkBtn.getAttribute('data-woid');
        unlinkTurnWorkOrder(tk, woid)
          .then(function() { showToast('WO #' + woid + ' removed from turn'); })
          .catch(function(err) { showToast('Remove failed: ' + (err.message || err)); });
        return;
      }
      var advBtn = e.target.closest('[data-advance]');
      if (advBtn) {
        e.stopPropagation();
        confirmTurnStage(advBtn.getAttribute('data-advance'), advBtn.getAttribute('data-stage'));
        return;
      }
      var closeBtn = e.target.closest('[data-close-turn]');
      if (closeBtn) {
        e.stopPropagation();
        closeTurn(closeBtn.getAttribute('data-close-turn'));
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
      var initialBtn = e.target.closest('.vendor-rail-btn[data-vendor-initial]');
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
      // Vendor card click — open vendor detail modal
      var card = e.target.closest('.vendor-card');
      if (card && !e.target.closest('.vendor-cat-select') && !e.target.closest('.vendor-trade-cat-select')) {
        var vid2 = card.getAttribute('data-vendorid');
        openVendorModal(vid2);
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
  var mobileNavToggleBtn = document.getElementById('mobileNavToggle');
  if (mobileNavToggleBtn) {
    mobileNavToggleBtn.addEventListener('click', function() {
      toggleMobileNav();
    });
  }
  var mobileNavBackdrop = document.getElementById('mobileNavBackdrop');
  if (mobileNavBackdrop) {
    mobileNavBackdrop.addEventListener('click', function() {
      closeMobileNav();
    });
  }
  window.addEventListener('resize', function() {
    syncMobileNavForViewport();
  });
  document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') closeMobileNav();
  });
  syncMobileNavForViewport();

  $$('.nav-tab').forEach(function(tab) {
    tab.addEventListener('click', async function() {
      var tabName = tab.getAttribute('data-tab');
      if (!isTabAllowedForRole(tabName)) return;
      $$('.nav-tab').forEach(function(t) { t.classList.remove('active'); });
      tab.classList.add('active');
      closeMobileNav();
      setMobileNavLabelFromActiveTab();
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

      // Lazy-load AP bills when Work Orders tab is opened (needed for WO Close Assist)
      if (tabName === 'workorders' && (!BILLS || BILLS.length === 0)) {
        fetchBills(DEFAULT_BILLS_LOOKBACK_DAYS).then(function() { renderWOCloseAssist(); }).catch(function() {});
      }

      // Lazy-load AP Bills when Billing tab is opened
      if (tabName === 'billing') {
        if (!window._billingPageRows || window._billingPageRows.length === 0) {
          var billBody = document.getElementById('billBody');
          if (billBody) billBody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:24px">' + loadingHtml('Loading bills\u2026') + '</td></tr>';
          loadBillingPage({ resetPage: true, forceHardLock: true }).catch(function(e) {
            showToast('Bills load failed: ' + (e.message || e));
          });
        } else {
          renderBillsSection();
        }
      }

      if (tabName === 'properties') {
        renderPropertiesSection();
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

      // Initialize Routing Monitor on tab open
      if (tabName === 'routing') {
        try { await initRoutingMonitor(); } catch (e) { showToast('Routing monitor load failed: ' + (e.message || e)); }
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
  if ($('#woVendorFilter')) {
    $('#woVendorFilter').addEventListener('change', function() { currentWOVendor = this.value; completedWOHistoryPage = 0; renderWorkOrders(); });
  }
  if ($('#woPropertyFilter')) {
    $('#woPropertyFilter').addEventListener('change', function() { currentWOProperty = this.value; completedWOHistoryPage = 0; renderWorkOrders(); });
  }
  if ($('#woAgeFilter')) {
    $('#woAgeFilter').addEventListener('change', function() { currentWOAgeFilter = this.value; renderWorkOrders(); });
  }
  if ($('#woSort')) {
    $('#woSort').addEventListener('change', function() { currentWOSort = this.value || 'oldest'; renderWorkOrders(); });
  }
  if ($('#btnSaveWOAging')) {
    $('#btnSaveWOAging').addEventListener('click', function() {
      saveWOAgeThresholdsFromInputs();
      renderWorkOrders();
      showToast('Saved WO aging thresholds', { kind: 'success' });
    });
  }
  ['woAgeYellow', 'woAgeOrange', 'woAgeRed'].forEach(function(id) {
    var el = $('#' + id);
    if (!el) return;
    el.addEventListener('change', function() {
      saveWOAgeThresholdsFromInputs();
      renderWorkOrders();
    });
  });
  if ($('#woCloseAge')) {
    $('#woCloseAge').addEventListener('change', function() {
      currentWOCloseAssistAge = parseInt(this.value || '14', 10) || 14;
      woCloseAssist.currentPage = 0;
      renderWOCloseAssist();
    });
  }
  // WO sub-tabs
  $$('[data-wo-subtab]').forEach(function(btn) {
    btn.addEventListener('click', function() {
      setWOSubtab(this.dataset.woSubtab);
    });
  });
  setWOSubtab('active');

  // Billing sub-tabs
  $$('[data-billing-subtab]').forEach(function(btn) {
    btn.addEventListener('click', function() {
      setBillingSubtab(btn.getAttribute('data-billing-subtab'));
    });
  });
  setBillingSubtab(currentBillingSubtab || 'queue');

  // Properties sub-tabs
  $$('[data-properties-subtab]').forEach(function(btn) {
    btn.addEventListener('click', function() {
      setPropertiesSubtab(btn.getAttribute('data-properties-subtab'));
    });
  });
  setPropertiesSubtab(currentPropertiesSubtab || 'directory');

  // Collapsible panels
  $$('.collapsible-panel').forEach(function(panel) {
    var hdr = panel.querySelector('.table-header');
    if (!hdr) return;
    hdr.addEventListener('click', function(e) {
      if (e.target.closest('button,select,input,a')) return;
      panel.classList.toggle('collapsed');
      try { localStorage.setItem('wo-panel-' + panel.id, panel.classList.contains('collapsed') ? '1' : '0'); } catch(_) {}
    });
    // Restore state
    try {
      var saved = localStorage.getItem('wo-panel-' + panel.id);
      if (saved === '0') panel.classList.remove('collapsed');
    } catch(_) {}
  });

  if ($('#woCompletedHistoryToggle')) {
    $('#woCompletedHistoryToggle').addEventListener('change', async function() {
      showCompletedWOHistory = !!this.checked;
      completedWOHistoryPage = 0;
      renderCompletedWOHistorySection();
    });
  }
  if ($('#btnFetchCompletedHistory')) {
    $('#btnFetchCompletedHistory').addEventListener('click', async function() {
      var btn = this;
      var daysEl = $('#woCompletedHistoryDays');
      var days = parseInt(daysEl ? daysEl.value : String(DEFAULT_COMPLETED_WO_LOOKBACK_DAYS), 10) || DEFAULT_COMPLETED_WO_LOOKBACK_DAYS;
      days = Math.max(30, Math.min(3650, days));
      if (daysEl) daysEl.value = String(days);
      showCompletedWOHistory = true;
      if ($('#woCompletedHistoryToggle')) $('#woCompletedHistoryToggle').checked = true;
      completedWOHistoryPage = 0;
      completedWOHistoryLoading = true;
      btn.disabled = true;
      btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Searching…';
      renderCompletedWOHistorySection();
      try {
        await fetchCompletedWorkOrdersHistory(days);
      } catch (e) {
        completedWOHistoryRows = [];
        showToast('Completed history load failed: ' + ((e && e.message) || e), { kind: 'warning' });
      } finally {
        completedWOHistoryLoading = false;
        btn.disabled = false;
        btn.innerHTML = '<i class="fas fa-search"></i> Search Completed';
        renderCompletedWOHistorySection();
      }
    });
  }
  if ($('#btnRefreshCloseAssist')) {
    $('#btnRefreshCloseAssist').addEventListener('click', async function() {
      if (_billsLoading) return;
      _billsLoading = true;
      this.disabled = true;
      this.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Loading AP\u2026';
      try {
        var ok = await fetchBills(DEFAULT_BILLS_LOOKBACK_DAYS);
        woCloseAssist.currentPage = 0;
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
      try { localStorage.setItem('flr_turn_dash_filter', this.value); } catch(e){ /* */ }
      DASH_TURN_PAGE = 0;
      renderTurnDashboardStrip();
    });
  }
  if ($('#dashTurnViewMode')) {
    try {
      var savedView = localStorage.getItem('hm_dash_turn_view_mode');
      if (savedView === 'cards' || savedView === 'compact') DASH_TURN_VIEW_MODE = savedView;
    } catch (e) { /* */ }
    $('#dashTurnViewMode').value = DASH_TURN_VIEW_MODE;
    $('#dashTurnViewMode').addEventListener('change', function() {
      DASH_TURN_VIEW_MODE = this.value === 'compact' ? 'compact' : 'cards';
      DASH_TURN_PAGE = 0;
      try { localStorage.setItem('hm_dash_turn_view_mode', DASH_TURN_VIEW_MODE); } catch (e) { /* */ }
      renderTurnDashboardStrip();
    });
  }
  if ($('#dashTurnSync')) {
    $('#dashTurnSync').addEventListener('click', function() {
      syncTurnDashboardIncremental(true);
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
  renderDashboardTurnSyncLabel();
  if (DASH_TURN_SYNC_TIMER) {
    clearInterval(DASH_TURN_SYNC_TIMER);
    DASH_TURN_SYNC_TIMER = null;
  }
  DASH_TURN_SYNC_TIMER = setInterval(function() {
    if (document.hidden) return;
    var sec = document.getElementById('sec-dashboard');
    var isDashboardActive = !!(sec && sec.classList.contains('active'));
    var isTv = !!(document.body && document.body.classList.contains('tv-mode-dashboard'));
    if (isDashboardActive || isTv) syncTurnDashboardIncremental(false);
    renderDashboardTurnSyncLabel();
  }, 5 * 60 * 1000);

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

  // Bulk Note modal close
  if ($('#bulkNoteModalClose')) {
    $('#bulkNoteModalClose').addEventListener('click', function() { closeModal('bulkNoteModal'); });
  }
  if ($('#bulkNoteCloseBtn')) {
    $('#bulkNoteCloseBtn').addEventListener('click', function() { closeModal('bulkNoteModal'); });
  }

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

  // Turn Detail Modal close
  if ($('#turnDetailClose'))    $('#turnDetailClose').addEventListener('click', function() { closeModal('turnDetailModal'); });
  if ($('#turnDetailCloseBtn')) $('#turnDetailCloseBtn').addEventListener('click', function() { closeModal('turnDetailModal'); });

  // Turn Kanban — persistent click delegation (survives innerHTML replacement)
  if ($('#turnKanban')) {
    $('#turnKanban').addEventListener('click', function(ev) {
      var card = ev.target.closest('.turn-kcard[data-pipeid]');
      if (card) openTurnDetailModal(card.getAttribute('data-pipeid'));
    });
  }

  $('#btnClearErrors').addEventListener('click', function() {
    API_ERRORS = API_ERRORS.filter(function(e) { return e.action !== 'resolved'; });
    renderErrorLog();
    showToast('Cleared resolved errors');
  });
  if ($('#btnRunSystemHealth')) {
    $('#btnRunSystemHealth').addEventListener('click', function() {
      runSystemHealthCheck();
    });
  }
  if ($('#btnSendSystemHealthDebug')) {
    $('#btnSendSystemHealthDebug').addEventListener('click', function() {
      sendSystemHealthDebugEmail();
    });
  }

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

  // Billing section — Refresh, search, status filter
  var billRefreshBtn = $('#btnRefreshBills');
  if (billRefreshBtn) {
    billRefreshBtn.addEventListener('click', async function() {
      try {
        await loadBillingPage({ forceRefresh: true, forceHardLock: true });
        showToast('Billing refreshed', { kind: 'success' });
      } catch (e) {
        showToast('Bills refresh failed: ' + (e.message || e));
      }
    });
  }
  var billSearch = $('#billSearch');
  if (billSearch) billSearch.addEventListener('input', function() { _billsPage = 0; renderBillsSection(); });
  var billStatusFilter = $('#billStatusFilter');
  if (billStatusFilter) {
    billStatusFilter.addEventListener('change', function() {
      _billingRouteStatus = String(billStatusFilter.value || '').trim();
      _billsPage = 0;
      loadBillingPage({ resetPage: true });
    });
  }
  var billHistBtn = $('#btnBillHistSearch');
  if (billHistBtn) {
    billHistBtn.addEventListener('click', function() {
      BILL_HISTORY_PAGE = 0;
      runBillHistorySearch();
    });
  }
  wireBillingFilters();

  var propertySearch = $('#propertySearch');
  if (propertySearch) {
    propertySearch.addEventListener('input', debounce(function() {
      _propertiesVacancyOnly = false;
      _propertiesPage = 0;
      renderPropertiesSection();
    }, CONFIG.DEBOUNCE_MS));
  }
  var propertyGroupFilter = $('#propertyGroupFilter');
  if (propertyGroupFilter) {
    propertyGroupFilter.addEventListener('change', function() {
      _propertiesVacancyOnly = false;
      _propertiesLocalGroup = normalizeGroupSelectionValue(propertyGroupFilter.value || '');
      _propertiesPage = 0;
      renderPropertiesSection();
    });
  }
  var propertiesPageSize = $('#propertiesPageSize');
  if (propertiesPageSize) {
    propertiesPageSize.value = String(_propertiesPageSize);
    propertiesPageSize.addEventListener('change', function() {
      _propertiesPageSize = Math.max(1, parseInt(propertiesPageSize.value, 10) || 50);
      _propertiesPage = 0;
      renderPropertiesSection();
    });
  }
  var propertiesPrevPage = $('#propertiesPrevPage');
  if (propertiesPrevPage) {
    propertiesPrevPage.addEventListener('click', function() {
      if (_propertiesPage > 0) {
        _propertiesPage -= 1;
        renderPropertiesSection();
      }
    });
  }
  var propertiesNextPage = $('#propertiesNextPage');
  if (propertiesNextPage) {
    propertiesNextPage.addEventListener('click', function() {
      _propertiesPage += 1;
      renderPropertiesSection();
    });
  }
  var refreshPropertiesBtn = $('#refreshPropertiesBtn');
  if (refreshPropertiesBtn) {
    refreshPropertiesBtn.addEventListener('click', function() {
      renderPropertiesSection({ forceRefresh: true });
    });
  }

  // Theme toggle
  $('#themeToggle').addEventListener('click', function() { toggleTheme(); });
  updateThemeIcon(); // sync icon with initial state

  // Settings modal
  if ($('#appSettingsBtn')) {
    $('#appSettingsBtn').addEventListener('click', function() { openModal('appSettingsModal'); });
  }
  if ($('#appSettingsClose')) {
    $('#appSettingsClose').addEventListener('click', function() { closeModal('appSettingsModal'); });
  }
  if ($('#appSettingsCloseBtn')) {
    $('#appSettingsCloseBtn').addEventListener('click', function() { closeModal('appSettingsModal'); });
  }
  if ($('#btnSettingsExportCache')) {
    $('#btnSettingsExportCache').addEventListener('click', function() {
      var sec = getCacheSecurityOptions();
      exportCacheToJSON(sec);
    });
  }
  if ($('#btnSettingsImportCache')) {
    $('#btnSettingsImportCache').addEventListener('click', function() {
      $('#cacheFileInput').click();
    });
  }

  // Cache export / import
  if ($('#btnExportCache')) {
    $('#btnExportCache').addEventListener('click', function() { exportCacheToJSON(getCacheSecurityOptions()); });
  }
  if ($('#btnImportCache')) {
    $('#btnImportCache').addEventListener('click', function() { $('#cacheFileInput').click(); });
  }
  if ($('#cacheFileInput')) {
    $('#cacheFileInput').addEventListener('change', function() {
      if (this.files && this.files[0]) {
        importCacheFromJSON(this.files[0], getCacheSecurityOptions());
        this.value = ''; // reset so same file can be re-imported
      }
    });
  }

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
    routing: function() { loadRoutingEventsAndStats().catch(function() {}); loadRoutingCapabilities().catch(function() {}); },
    payroll: function() { renderPayroll(); },
    billing: function() { renderBillsSection(); },
    properties: function() { renderPropertiesSection(); },
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
    document.dispatchEvent(new CustomEvent('groupFilterChanged', {
      detail: {
        groupName: currentPropertyGroup || '',
        forcedGroupUuid: forcedPropertyGroupUuid || ''
      }
    }));
  }

  function clearPropertyGroupFilters() {
    if (forcedPropertyGroupUuid) {
      enforceScopedPropertyGroup();
      return;
    }
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
      if (forcedPropertyGroupUuid) {
        enforceScopedPropertyGroup();
        return;
      }
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
      if (urlEl) urlEl.value = API_PROXY;
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
        await fetchBills(DEFAULT_BILLS_LOOKBACK_DAYS);
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
      await fetchUpcomingMoveouts();
      await fetchTurnWorkOrders();
      await fetchUnitTurnsDB();
      loadClosedTurns().catch(function() {});
      fetchUnitTurnHistory().then(function() { renderTurnHistoryPanel(); }).catch(function() {});
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
        setApiStatus('', 'Vendor Access [v8] \u2014 ' + VENDORS.length + ' vendors');
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
  loadWOAgeThresholds();

  // Show skeleton loading states
  setDashboardKpiSkeleton(true);
  if ($('#kanbanBoard')) $('#kanbanBoard').innerHTML = loadingHtml('Checking cache\u2026');
  if ($('#vendorGrid')) $('#vendorGrid').innerHTML = loadingHtml('Checking cache\u2026');
  if ($('#turnPipeline')) $('#turnPipeline').innerHTML = loadingHtml('Checking cache\u2026');
  if ($('#inspBody')) $('#inspBody').innerHTML = '<tr><td colspan="8">' + loadingHtml('Checking cache\u2026') + '</td></tr>';
  renderTemplates();
  renderErrorLog();
  wireUpUI();
  syncWOAgeInputs();

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
    var schemaHealth = applyProxySchemaHealth(pingData);
    if (pingData && pingData.brand) applyBrandConfig(pingData.brand);

    // Cache server version and detect version mismatch (force refresh if needed)
    if (pingData && pingData.version) {
      SERVER_VERSION = String(pingData.version);
      localStorage.setItem('hm_server_version', SERVER_VERSION);
      if (SERVER_VERSION !== APP_VERSION && compareVersions(SERVER_VERSION, APP_VERSION) > 0) {
        localStorage.setItem('hm_version_mismatch', '1');
        setTimeout(function() { location.reload(); }, 2000);
      }
    }

    // Detect proxy version when the proxy includes a version field
    _proxyVersion = pingData.version || 'v7';

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
    if (schemaHealth && schemaHealth.suffix) pingMsg += schemaHealth.suffix;
    if (!dbOk) {
      pingMsg += ' (DB API down, Reports OK)';
      logApiError(0, 'DB API v0 unreachable — some features may be limited', 'resolved');
    }
    if (schemaHealth && schemaHealth.hasIssue) {
      logApiError(500, 'Proxy schema health warning: ' + schemaHealth.detail, 'queued');
      if (!_schemaHealthWarned) {
        _schemaHealthWarned = true;
        showToast('Proxy schema warning: ' + schemaHealth.detail, {
          kind: 'warning',
          iconClass: 'fa-triangle-exclamation',
          duration: 5500,
        });
      }
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
    // Load closed turns (non-blocking, best-effort)
    loadClosedTurns().catch(function() {});
    fetchUnitTurnHistory().then(function() { renderTurnHistoryPanel(); }).catch(function() {});

    // Final re-render: turns with all available correlated data
    renderTurnBoard();

    if (BILLS.length === 0) {
      // Best-effort AP load for WO close-assist; non-blocking and non-fatal.
      try { await withStepTimeout(function() { return fetchBills(DEFAULT_BILLS_LOOKBACK_DAYS); }, 30000); } catch (e) { /* ignore */ }
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
    var stalled = TURN_PIPE_DATA.filter(function(p) {
      if (!isInPropertyGroup(p.propertyId, p.property, currentPropertyGroup)) return false;
      return p.isStalled && !p.isCompleted;
    });
    // Also show turns past deposit deadline
    var depositOverdue = TURN_PIPE_DATA.filter(function(p) {
      if (!isInPropertyGroup(p.propertyId, p.property, currentPropertyGroup)) return false;
      return p.isConfirmed && p.sla && p.sla.overdue && !p.isCompleted;
    });
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
    var scopedInspections = INSPECTIONS.filter(function(r) {
      return isInPropertyGroup(r.propertyId, r.propertyName, currentPropertyGroup);
    });
    var overdue = scopedInspections.filter(function(r) {
      var state = getInspectionCompliance(r, attnToday);
      return state.overdue;
    });
    var totalInsp = scopedInspections.length;
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
    var vendorsInGroup = null;
    if (currentPropertyGroup) {
      vendorsInGroup = {};
      WORK_ORDERS.forEach(function(wo) {
        if (isInPropertyGroup(wo.propertyId, wo.propertyName, currentPropertyGroup) && wo.vendorName) {
          vendorsInGroup[String(wo.vendorName).toLowerCase()] = true;
        }
      });
    }
    var vAlerts = VENDORS.filter(function(v) {
      if (vendorsInGroup && !vendorsInGroup[String(v.name || '').toLowerCase()]) return false;
      var ed = v.insurance ? new Date(v.insurance) : null;
      return ed && ed < attnToday;
    });
    var expiringSoon = VENDORS.filter(function(v) {
      if (vendorsInGroup && !vendorsInGroup[String(v.name || '').toLowerCase()]) return false;
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
      if ((w.priority !== 'Urgent' && w.priority !== 'Emergency') || w.status === 'Completed' || w.status === 'Canceled') return false;
      return isInPropertyGroup(w.propertyId, w.propertyName, currentPropertyGroup);
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

  // Pending bill approvals
  var pendingBillsEl = document.getElementById('attentionBillsBody');
  if (pendingBillsEl) {
    var pendingBills = (BILLS || []).filter(function(b) {
      if (!isInPropertyGroup(b.propertyId, b.propertyName, currentPropertyGroup)) return false;
      return getBillStatusKey(b) === 'pending_approval';
    });

    if (!pendingBills.length) {
      pendingBillsEl.innerHTML = '<div class="attn-empty"><i class="fas fa-check-circle" style="color:var(--success)"></i> No pending bill approvals</div>';
    } else {
      var pbHtml = '<div class="attn-count">' + pendingBills.length + ' pending approval</div>';
      pendingBills.slice(0, 5).forEach(function(b) {
        pbHtml += '<div class="attn-item"><span class="attn-label">' +
          escapeHtml(String(b.propertyName || 'Unknown')) + '</span><span class="attn-value">' +
          escapeHtml(currency(b.amount || 0)) + '</span></div>';
      });
      if (pendingBills.length > 5) pbHtml += '<div style="text-align:center;color:var(--text-muted);font-size:10px;margin-top:4px">+' + (pendingBills.length - 5) + ' more</div>';
      pendingBillsEl.innerHTML = pbHtml;
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
  var turnHistoryRefreshBtn = document.getElementById('btnRefreshTurnHistory');
  var pmUsersBody = document.getElementById('pmUsersBody');
  var pmUserUuidEl = document.getElementById('pmUserUuid');
  var pmUserEmailEl = document.getElementById('pmUserEmail');
  var pmUserNameEl = document.getElementById('pmUserName');
  var pmUserPhoneEl = document.getElementById('pmUserPhone');
  var pmUserGroupUuidEl = document.getElementById('pmUserGroupUuid');
  var pmUserSaveBtn = document.getElementById('btnPmUserSave');
  var pmUserClearBtn = document.getElementById('btnPmUserClear');
  var pmUserRefreshBtn = document.getElementById('btnRefreshPmUsers');
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

    function looksLikeUuid(value) {
      return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(value || '').trim());
    }

    function shouldLinkWorkOrder(colName, value) {
      var col = String(colName || '').toLowerCase();
      var raw = String(value || '').trim();
      if (!raw) return false;
      if (col.indexOf('work_order') !== -1 || col === 'wo_id' || col === 'wo_uuid' || col === 'wo_number') return true;
      if (/^wo[-_ ]?\d+$/i.test(raw)) return true;
      if (/^\d{3,}$/i.test(raw) && (col === 'id' || col.indexOf('wo') !== -1)) return true;
      return false;
    }

    function renderDbCell(colName, value) {
      if (value === null || value === undefined) return '<td class="db-null">NULL</td>';
      var raw = String(value);
      var safe = escapeHtml(raw);
      var encoded = encodeURIComponent(raw);

      if (shouldLinkWorkOrder(colName, raw)) {
        return '<td title="' + safe + '"><button class="dbadmin-cell-link" data-db-link-kind="wo" data-db-link-ref="' + encoded + '">' + safe + '</button></td>';
      }

      if (looksLikeUuid(raw)) {
        return '<td title="' + safe + '"><button class="dbadmin-cell-link" data-db-link-kind="uuid" data-db-link-ref="' + encoded + '">' + safe + '</button></td>';
      }

      return '<td title="' + safe + '">' + safe + '</td>';
    }

    var thead = '<thead><tr>' + cols.map(function(c) { return '<th>' + escapeHtml(c) + '</th>'; }).join('') + '</tr></thead>';
    var tbody = '<tbody>' + rows.map(function(row) {
      return '<tr>' + cols.map(function(c) {
        var v = typeof row === 'object' && !Array.isArray(row) ? row[c] : row[cols.indexOf(c)];
        return renderDbCell(c, v);
      }).join('') + '</tr>';
    }).join('') + '</tbody>';
    resultsBody.innerHTML = '<table>' + thead + tbody + '</table>';
  }

  if (resultsBody) {
    resultsBody.addEventListener('click', function(ev) {
      var link = ev.target.closest('.dbadmin-cell-link');
      if (!link) return;
      var kind = String(link.getAttribute('data-db-link-kind') || '');
      var ref = decodeURIComponent(String(link.getAttribute('data-db-link-ref') || ''));
      if (!ref) return;

      if (kind === 'wo' && typeof showWODetail === 'function') {
        showWODetail(ref);
        return;
      }
      if (kind === 'uuid' && typeof showV0UuidDetailModal === 'function') {
        showV0UuidDetailModal(ref);
      }
    });
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

  function populatePMGroupDropdown(currentUuid) {
    currentUuid = currentUuid || '';
    var sel    = document.getElementById('pm-user-group-select');
    var hidden = document.getElementById('pmUserGroupUuid');
    if (!sel) return;
    var groups = (typeof PROPERTY_GROUPS !== 'undefined' ? PROPERTY_GROUPS : []);
    if (!Array.isArray(groups) || !groups.length) {
      sel.innerHTML = '<option value="">— No groups loaded — refresh data first —</option>';
      sel.style.borderColor = 'var(--warning, #f0a500)';
      if (hidden) hidden.value = '';
      return;
    }
    var sorted = groups.slice().sort(function(a, b) {
      return (a.name || a.Name || '').toLowerCase().localeCompare((b.name || b.Name || '').toLowerCase());
    });
    sel.innerHTML = '<option value="">— Select a Property Group —</option>';
    sorted.forEach(function(g) {
      var uuid = g.id || g.uuid || g.property_group_id || '';
      var name = g.name || g.Name || g.display_name || uuid;
      if (!uuid) return;
      var opt = document.createElement('option');
      opt.value = uuid;
      opt.textContent = name;
      if (uuid === currentUuid) opt.selected = true;
      sel.appendChild(opt);
    });
    sel.onchange = function() {
      if (hidden) hidden.value = sel.value;
      sel.style.borderColor = sel.value ? 'var(--accent, #5c9eff)' : '';
    };
    if (hidden) hidden.value = currentUuid;
    if (currentUuid) {
      sel.value = currentUuid;
      sel.style.borderColor = 'var(--accent, #5c9eff)';
    } else {
      sel.style.borderColor = '';
    }
  }

  function clearPmUserForm() {
    if (pmUserUuidEl) pmUserUuidEl.value = '';
    if (pmUserEmailEl) pmUserEmailEl.value = '';
    if (pmUserNameEl) pmUserNameEl.value = '';
    if (pmUserPhoneEl) pmUserPhoneEl.value = '';
    if (pmUserGroupUuidEl) pmUserGroupUuidEl.value = '';
    populatePMGroupDropdown('');
    if (pmUserSaveBtn) pmUserSaveBtn.innerHTML = '<i class="fas fa-save"></i> Save';
  }

  function renderPmUsers(users) {
    if (!pmUsersBody) return;
    if (!users || !users.length) {
      pmUsersBody.innerHTML = '<div class="dbadmin-msg">No PM users configured</div>';
      return;
    }
    var rows = users.map(function(u) {
      var uuid = escapeHtml(String(u.user_uuid || ''));
      var email = escapeHtml(String(u.email || ''));
      var fullName = escapeHtml(String(u.full_name || ''));
      var phone = escapeHtml(String(u.phone || ''));
      var groupUuid = escapeHtml(String(u.property_group_uuid || ''));
      var active = Number(u.active) === 1;
      return '<tr data-uuid="' + uuid + '" data-email="' + email + '" data-full-name="' + fullName + '" data-phone="' + phone + '" data-group-uuid="' + groupUuid + '">' +
        '<td style="font-family:var(--font-mono)">' + email + '</td>' +
        '<td>' + fullName + '</td>' +
        '<td style="font-family:var(--font-mono)">' + phone + '</td>' +
        '<td style="font-family:var(--font-mono)">' + groupUuid + '</td>' +
        '<td>' + (active ? '<span style="color:var(--success)">Active</span>' : '<span style="color:var(--danger)">Inactive</span>') + '</td>' +
        '<td style="text-align:right;white-space:nowrap">' +
          '<button class="dbadmin-btn pm-edit" style="padding:2px 8px;font-size:10px" data-uuid="' + uuid + '"><i class="fas fa-pen"></i> Edit</button> ' +
          '<button class="dbadmin-btn danger pm-delete" style="padding:2px 8px;font-size:10px" data-uuid="' + uuid + '"><i class="fas fa-trash"></i> Delete</button>' +
        '</td>' +
      '</tr>';
    }).join('');
    pmUsersBody.innerHTML = '<table><thead><tr><th>Email</th><th>Name</th><th>Phone</th><th>Property Group UUID</th><th>Status</th><th>Actions</th></tr></thead><tbody>' + rows + '</tbody></table>';
  }

  function loadPmUsers() {
    if (!pmUsersBody) return Promise.resolve();
    var adminKey = (document.getElementById('dbAdminKey') && document.getElementById('dbAdminKey').value || '').trim();
    // Try to retrieve from sessionStorage if not entered
    if (!adminKey) {
      try { adminKey = (sessionStorage.getItem('hm_pm_admin_key') || '').trim(); } catch (e) { /* */ }
    }
    pmUsersBody.innerHTML = '<div class="dbadmin-msg">Loading PM users…</div>';
    if (!API_PROXY) {
      try { var _sp = localStorage.getItem('hm_proxy_url'); if (_sp) API_PROXY = _sp; } catch (e) { /* */ }
    }
    if (!API_PROXY) {
      pmUsersBody.innerHTML = '<div class="dbadmin-msg" style="color:var(--warning)"><i class="fas fa-plug"></i> Proxy not configured. Connect via the vault first.</div>';
      return Promise.resolve();
    }
    if (!adminKey) {
      pmUsersBody.innerHTML = '<div class="dbadmin-msg" style="color:var(--warning)"><i class="fas fa-key"></i> Admin key required. Enter above to load PM users.</div>';
      return Promise.resolve();
    }
    // Store key in sessionStorage for future auto-load
    try { sessionStorage.setItem('hm_pm_admin_key', adminKey); } catch (e) { /* */ }
    return proxyAction('pm_proxy_users', { key: adminKey }).then(function(data) {
      renderPmUsers(data.results || []);
    }).catch(function(err) {
      pmUsersBody.innerHTML = '<div class="dbadmin-msg" style="color:var(--danger)"><i class="fas fa-exclamation-circle"></i> ' + escapeHtml(err.message || String(err)) + '</div>';
    });
  }

  function savePmUser() {
    var email = (pmUserEmailEl && pmUserEmailEl.value || '').trim().toLowerCase();
    var fullName = (pmUserNameEl && pmUserNameEl.value || '').trim();
    var phone = (pmUserPhoneEl && pmUserPhoneEl.value || '').trim();
    var groupUuid = (pmUserGroupUuidEl && pmUserGroupUuidEl.value || '').trim()
      || (document.getElementById('pm-user-group-select') && document.getElementById('pm-user-group-select').value || '').trim()
      || '';
    var userUuid = (pmUserUuidEl && pmUserUuidEl.value || '').trim();
    if (!email) {
      showToast('PM user email is required', 'error');
      return;
    }
    if (!fullName) {
      showToast('PM user full name is required', 'error');
      return;
    }
    if (!phone) {
      showToast('PM user phone is required', 'error');
      return;
    }
    if (!groupUuid) {
      showToast('Select a property group before saving.', 'error');
      return;
    }
    if (pmUserSaveBtn) pmUserSaveBtn.disabled = true;
    // Restore API_PROXY from localStorage if the session was reopened without full vault login
    if (!API_PROXY) {
      try { var _savedProxy = localStorage.getItem('hm_proxy_url'); if (_savedProxy) API_PROXY = _savedProxy; } catch (e) { /* */ }
    }
    if (!API_PROXY) {
      showToast('Proxy not configured. Connect to the proxy first (enter URL in the vault).', 'error');
      if (pmUserSaveBtn) pmUserSaveBtn.disabled = false;
      return;
    }
    proxyPost('pm_proxy_user_upsert', {
      key: getAdminKey(),
      user_uuid: userUuid || undefined,
      email: email,
      full_name: fullName,
      phone: phone,
      property_group_uuid: groupUuid,
      active: 1
    }).then(function(response) {
      if (!response || response.ok === false) {
        var errMsg = (response && (response.message || response.error)) || 'PM save failed';
        showToast(errMsg, 'error');
        return;
      }
      var msg = (userUuid ? 'PM user updated' : 'PM user created') + ' — UUID: ' + (response.user_uuid || '?').slice(0, 8);
      showToast(msg, { kind: 'success', duration: 5000 });
      clearPmUserForm();
      return loadPmUsers();
    }).catch(function(err) {
      showToast(err.message || String(err), 'error');
    }).finally(function() {
      if (pmUserSaveBtn) pmUserSaveBtn.disabled = false;
    });
  }

  async function deletePmUser(userUuid) {
    if (!userUuid) return;
    if (!await hmConfirm('Delete this PM login user?', { title: 'Delete PM User', okLabel: 'Delete', danger: true })) return;
    if (!API_PROXY) {
      try { var _savedProxy2 = localStorage.getItem('hm_proxy_url'); if (_savedProxy2) API_PROXY = _savedProxy2; } catch (e) { /* */ }
    }
    if (!API_PROXY) { showToast('Proxy not configured.', 'error'); return; }
    proxyPost('pm_proxy_user_delete', { key: getAdminKey(), user_uuid: userUuid }).then(function() {
      showToast('PM user deleted', 'success');
      if (pmUserUuidEl && pmUserUuidEl.value === userUuid) clearPmUserForm();
      return loadPmUsers();
    }).catch(function(err) {
      showToast(err.message || String(err), 'error');
    });
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

  if (turnHistoryRefreshBtn) {
    turnHistoryRefreshBtn.addEventListener('click', function() {
      fetchUnitTurnHistory().then(function() { renderTurnHistoryPanel(); }).catch(function() {});
    });
  }

  if (pmUserSaveBtn) pmUserSaveBtn.addEventListener('click', savePmUser);
  if (pmUserClearBtn) pmUserClearBtn.addEventListener('click', clearPmUserForm);
  if (pmUserRefreshBtn) pmUserRefreshBtn.addEventListener('click', function() { loadPmUsers(); });
  
  // Auto-load PM users on page load if proxy is available
  function autoLoadPmUsers() {
    if (!API_PROXY && !pmUsersBody) return;
    try { var _sp = localStorage.getItem('hm_proxy_url'); if (_sp) API_PROXY = _sp; } catch (e) { /* */ }
    if (API_PROXY) {
      loadPmUsers().catch(function() { /* silently ignore auto-load failures */ });
    }
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function() { setTimeout(autoLoadPmUsers, 500); });
  } else {
    setTimeout(autoLoadPmUsers, 500);
  }
  if (pmUsersBody) {
    pmUsersBody.addEventListener('click', function(ev) {
      var editBtn = ev.target.closest('.pm-edit');
      var deleteBtn = ev.target.closest('.pm-delete');
      if (editBtn) {
        var tr = editBtn.closest('tr');
        if (!tr) return;
        if (pmUserUuidEl) pmUserUuidEl.value = tr.getAttribute('data-uuid') || '';
        if (pmUserEmailEl) pmUserEmailEl.value = tr.getAttribute('data-email') || '';
        if (pmUserNameEl) pmUserNameEl.value = tr.getAttribute('data-full-name') || '';
        if (pmUserPhoneEl) pmUserPhoneEl.value = tr.getAttribute('data-phone') || '';
        if (pmUserGroupUuidEl) pmUserGroupUuidEl.value = tr.getAttribute('data-group-uuid') || '';
        populatePMGroupDropdown(tr.getAttribute('data-group-uuid') || '');
        if (pmUserSaveBtn) pmUserSaveBtn.innerHTML = '<i class="fas fa-save"></i> Update';
        if (pmUserEmailEl) pmUserEmailEl.focus();
        return;
      }
      if (deleteBtn) {
        deletePmUser(deleteBtn.getAttribute('data-uuid') || '');
      }
    });
  }

  fetchUnitTurnHistory().then(function() { renderTurnHistoryPanel(); }).catch(function() {});
  populatePMGroupDropdown('');
  // Expose a re-populate hook so populateGroupFilters() can call it once groups load
  window._repopulatePMGroupDropdown = function() { populatePMGroupDropdown(''); };
  // Also re-run when admin key is entered (groups may already be loaded by then)
  var dbKeyInput = document.getElementById('dbAdminKey');
  if (dbKeyInput) {
    dbKeyInput.addEventListener('change', function() {
      if (dbKeyInput.value.trim()) { populatePMGroupDropdown(''); loadPmUsers(); }
    });
  }
  loadPmUsers();

  // ── OTP Settings Panel ──────────────────────────────────────────────────
  (function initOtpSettings() {
    var enabledCb = document.getElementById('otpSettingEnabled');
    var enabledLabel = document.getElementById('otpSettingEnabledLabel');
    var membershipCb = document.getElementById('otpSettingRequireMembership');
    var membershipLabel = document.getElementById('otpSettingRequireMembershipLabel');
    var domainInput = document.getElementById('otpSettingDomain');
    var ttlInput = document.getElementById('otpSettingTtl');
    var saveBtn = document.getElementById('btnOtpSettingsSave');
    var saveStatus = document.getElementById('otpSettingsSaveStatus');
    var refreshBtn = document.getElementById('btnOtpSettingsRefresh');
    if (!enabledCb || !saveBtn) return;

    function updateCheckboxLabels() {
      if (enabledLabel) enabledLabel.textContent = enabledCb.checked ? 'Enabled' : 'Disabled';
      if (membershipLabel) membershipLabel.textContent = membershipCb && membershipCb.checked ? 'Required' : 'Not required';
    }
    if (enabledCb) enabledCb.addEventListener('change', updateCheckboxLabels);
    if (membershipCb) membershipCb.addEventListener('change', updateCheckboxLabels);

    function loadOtpSettings() {
      var keyInput = document.getElementById('dbAdminKey');
      var adminKey = (keyInput && keyInput.value || '').trim();
      if (!adminKey) {
        try {
          adminKey = String(localStorage.getItem('hm_proxy_admin_key') || '').trim() || String(sessionStorage.getItem('hm_pm_admin_key') || '').trim();
        } catch (e) { /* */ }
      }
      if (adminKey && keyInput && !String(keyInput.value || '').trim()) {
        keyInput.value = adminKey;
      }
      if (!adminKey) { if (saveStatus) { saveStatus.textContent = 'Enter Admin Key above to load settings'; saveStatus.style.color = 'var(--text-muted)'; } return; }
      proxyAction('settings_get', { key: adminKey }).then(function(data) {
        var settings = {};
        (data.settings || []).forEach(function(s) { settings[s.key] = s.value; });
        if (enabledCb) enabledCb.checked = (settings['otp_enabled'] ?? '1') !== '0';
        if (membershipCb) membershipCb.checked = (settings['otp_require_pm_membership'] ?? '1') !== '0';
        if (domainInput) domainInput.value = settings['otp_allowed_domain'] || '';
        if (ttlInput) ttlInput.value = settings['otp_ttl_minutes'] || '10';
        updateCheckboxLabels();
        if (saveStatus) { saveStatus.textContent = 'Settings loaded'; saveStatus.style.color = 'var(--success)'; setTimeout(function() { if (saveStatus) saveStatus.textContent = ''; }, 2000); }
      }).catch(function(err) {
        if (saveStatus) { saveStatus.textContent = err.message || String(err); saveStatus.style.color = 'var(--danger)'; }
      });
    }

    if (refreshBtn) refreshBtn.addEventListener('click', loadOtpSettings);

    saveBtn.addEventListener('click', function() {
      var adminKey = (document.getElementById('dbAdminKey') && document.getElementById('dbAdminKey').value || '').trim();
      if (!adminKey) { showToast('Enter Admin Key first', 'error'); return; }
      var saves = [
        { key: 'otp_enabled', value: enabledCb.checked ? '1' : '0' },
        { key: 'otp_require_pm_membership', value: (membershipCb && membershipCb.checked) ? '1' : '0' },
        { key: 'otp_allowed_domain', value: (domainInput && domainInput.value.trim().replace(/^@/, '')) || '' },
        { key: 'otp_ttl_minutes', value: String(Math.max(3, parseInt(ttlInput && ttlInput.value || '10', 10) || 10)) },
      ];
      saveBtn.disabled = true;
      if (saveStatus) { saveStatus.textContent = 'Saving…'; saveStatus.style.color = 'var(--text-muted)'; }
      var promises = saves.map(function(s) {
        return proxyPost('settings_set', { admin_key: adminKey, key: s.key, value: s.value });
      });
      Promise.all(promises).then(function() {
        showToast('OTP settings saved', 'success');
        if (saveStatus) { saveStatus.textContent = 'Saved ✓'; saveStatus.style.color = 'var(--success)'; setTimeout(function() { if (saveStatus) saveStatus.textContent = ''; }, 3000); }
      }).catch(function(err) {
        showToast(err.message || String(err), 'error');
        if (saveStatus) { saveStatus.textContent = err.message || 'Error'; saveStatus.style.color = 'var(--danger)'; }
      }).finally(function() { saveBtn.disabled = false; });
    });

    // Auto-load when admin key is entered
    var adminKeyInput = document.getElementById('dbAdminKey');
    if (adminKeyInput) {
      adminKeyInput.addEventListener('change', function() { if (adminKeyInput.value.trim()) loadOtpSettings(); });
    }

    loadOtpSettings();
  })();

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

  async function safePollWebhookEvents() {
    try {
      if (typeof pollWebhookEvents === 'function') {
        await pollWebhookEvents();
      }
    } catch (err) {
      console.debug('Webhook poll cycle skipped due to network/proxy disconnect.');
    }
  }

  async function safePollLiveEvents() {
    try {
      if (!isProxySessionReady()) {
        setStatus('Awaiting sign-in…');
        return;
      }
      var data = await proxyAction('webhook_events', { limit: 100 });
      mergeEvents(data && Array.isArray(data.events) ? data.events : []);
      setStatus('Listening for AppFolio events…');
    } catch (err) {
      setStatus('Live feed temporarily unavailable. Retrying…');
      console.debug('Live events poll cycle skipped due to network/proxy disconnect.');
    }
  }

  var POLL_INTERVAL_MS = 30000;
  async function _poll() {
    if (_pollTimer) {
      clearTimeout(_pollTimer);
      _pollTimer = null;
    }
    try {
      await Promise.allSettled([
        safePollWebhookEvents(),
        safePollLiveEvents()
      ]);
    } catch (criticalErr) {
      console.error('Critical error in webhook polling engine:', criticalErr);
    } finally {
      _pollTimer = setTimeout(_poll, POLL_INTERVAL_MS);
    }
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

  setTimeout(_poll, 2000);
  document.addEventListener('visibilitychange', function() {
    if (!document.hidden) {
      console.log('Tab active: forcing immediate webhook poll sync...');
      _poll();
    }
  });
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

/* =================================================================
   DISPATCH ENGINE — Implementation Notes
   Tech UUIDs must hold "Maintenance Tech" role in AppFolio or
   every automated PATCH returns 422 "User not found".
   Concurrent PATCHes to the same WO: first succeeds, second fails.
   The midnight cron serialises writes with 200ms delays.
   ================================================================ */

// ── Module-level state ──────────────────────────────────────────
var DISPATCH = {
  queue: [], techs: [], audit: [], blasts: [], claims: [], comms: [], stats: {},
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
  _pollTimer: null, _lastAuditMax: 0, POLL_MS: 30000,
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
  var headers = { 'Content-Type': 'application/json' };
  var token = getProxyAccessToken();
  if (token) {
    headers['Authorization'] = 'Bearer ' + token;
    headers['x-proxy-token'] = token;
  }
  return fetch(url, {
    method:'POST', headers: headers, body:JSON.stringify(body)
  }).then(function(r){
    if (!r.ok) {
      return r.text().then(function(t) {
        if (r.status === 401) {
          throw new Error('Dispatch ' + action + ' unauthorized: reconnect in Vault (missing or invalid proxy token).');
        }
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
    return true;
  });
  if(filtered.length===0){
    tbody.innerHTML='<tr><td colspan="8"><div class="dispatch-empty"><i class="fas fa-check-circle" style="color:var(--success)"></i>'+(search||status?'No queue entries match your filter.':'Queue is empty — no stale work orders.')+'</div></td></tr>';
    return;
  }
  var html='';
  filtered.forEach(function(r){
    var rowBranch = getDispatchQueueBranch(r);
    var icon=r.auto_exempt?'🔕':r.escalated?'🚨':r.warning_sent?'⚠️':r.grace_used?'🔵':'⚪';
    var woId=escapeHtml(String(r.wo_id||''));
    var woNum=escapeHtml(String(r.wo_number||r.wo_id||'').substring(0,20));
    var addr=escapeHtml(String(r.property_address||'—').substring(0,40));
    var tech=escapeHtml(String(r.assigned_tech_name||'—'));
    var branchLabel = rowBranch === 'unknown' ? 'Unknown' : (rowBranch === 'phoenix' ? 'Phoenix' : 'Tucson');
    var firstSeen=r.first_seen_at?timeAgo(r.first_seen_at):'—';
    var lastAct=r.last_reassigned_at?timeAgo(r.last_reassigned_at):(r.warning_sent_at?timeAgo(r.warning_sent_at):'—');
    var isMonitored = (DISPATCH.monitored||[]).some(function(m){return String(m.wo_id)===String(r.wo_id);});
    var exemptBtn=!r.auto_exempt
      ?'<button class="btn-dispatch-secondary btn-xs-dispatch" onclick="DispatchQueue.markExempt(\''+woId+'\',\''+woNum+'\')">🔕 Exempt</button>'
      :'<button class="btn-dispatch-danger btn-xs-dispatch" onclick="DispatchQueue.clearExempt(\''+woId+'\')">🔓 Clear</button>';
    var monBtn = isMonitored
      ? '<button class="btn-dispatch-danger btn-xs-dispatch" onclick="DispatchQueue.removeMonitored(\''+woId+'\')"><i class="fas fa-star"></i> Monitored</button>'
      : '<button class="btn-dispatch-secondary btn-xs-dispatch" onclick="DispatchQueue.addMonitored(\''+woId+'\',\''+woNum+'\')"><i class="fas fa-star-o"></i> Monitor</button>';
    html+='<tr style="'+(r.escalated?'background:rgba(209,59,59,.04)':'')+(r.auto_exempt?';opacity:.55':'')+'">'+
      '<td style="text-align:center;font-size:1rem">'+icon+'</td>'+
      '<td style="font-family:var(--font-mono);font-size:.8rem;color:var(--accent)">#'+woNum+'</td>'+
      '<td title="'+escapeHtml(r.property_address||'')+'">'+addr+'</td>'+
      '<td>'+tech+'<div style="font-size:.66rem;color:var(--text-muted);font-family:var(--font-mono)">'+branchLabel+'</div></td>'+
      '<td>'+(Number(r.reassignment_count)>0?'<span style="background:var(--danger-dim);color:var(--danger);padding:1px 6px;border-radius:4px;font-family:var(--font-mono);font-size:.72rem;font-weight:700">'+r.reassignment_count+'×</span>':'<span style="color:var(--text-muted)">—</span>')+'</td>'+
      '<td style="font-family:var(--font-mono);font-size:.72rem;color:var(--text-muted)">'+firstSeen+'</td>'+
      '<td style="font-family:var(--font-mono);font-size:.72rem;color:var(--text-muted)">'+lastAct+'</td>'+
      '<td style="display:flex;gap:4px;flex-wrap:wrap">'+exemptBtn+monBtn+
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
    var hidden = isTechHidden(t.tech_id);
    // Hidden assignees stay visible in roster so managers can edit/unhide them.
    if (DISPATCH.activeBranch !== 'all' && branch !== DISPATCH.activeBranch && !hidden) return;
    var sn=Number(t.performance_score||100);
    var scoreColor=sn>=80?'var(--success)':sn>=60?'var(--warning)':'var(--danger)';
    var scoreBg=sn>=80?'var(--success-dim)':sn>=60?'var(--warning-dim)':'var(--danger-dim)';
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
  var savedBrandName = localStorage.getItem('hm_brand_name') || configMap.brand_name || BRAND_NAME_DEFAULT;
  var savedBrandLogo = localStorage.getItem('hm_brand_logo') || configMap.brand_logo_url || '';
  var savedPortalBrandName = localStorage.getItem('hm_portal_brand_name') || configMap.portal_brand_name || PORTAL_BRAND_NAME_DEFAULT;
  var savedPortalBrandLogo = localStorage.getItem('hm_portal_brand_logo') || configMap.portal_brand_logo_url || PORTAL_BRAND_LOGO_DEFAULT;
  var savedPortalBrandTitle = localStorage.getItem('hm_portal_brand_title') || configMap.portal_brand_title || 'Dispatch Portal';
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
    '<div style="display:flex;justify-content:space-between;align-items:flex-start;padding:10px 0;border-bottom:1px solid var(--border);gap:12px;flex-wrap:wrap">'+
    '<div><label style="font-weight:700;font-size:.85rem;display:block">Web App Branding</label>'+
    '<span style="font-size:.72rem;color:var(--text-muted)">Set the logo URL used by the main HandyManager web app.</span></div>'+
    '<div style="display:flex;flex-direction:column;gap:6px;min-width:300px">'+
      '<input type="text" id="cfgBrandName" value="'+escapeHtml(savedBrandName)+'" placeholder="Brand name" style="font-family:var(--font-mono);font-size:.74rem;padding:5px 9px;border-radius:7px;border:1px solid var(--border);background:var(--bg-input);color:var(--text-primary)">'+
      '<input type="text" id="cfgBrandLogoUrl" value="'+escapeHtml(savedBrandLogo)+'" placeholder="https://.../logo.png" style="font-family:var(--font-mono);font-size:.74rem;padding:5px 9px;border-radius:7px;border:1px solid var(--border);background:var(--bg-input);color:var(--text-primary)">'+
      '<div style="display:flex;gap:6px;flex-wrap:wrap">'+
        '<button class="btn-dispatch-secondary btn-xs-dispatch" onclick="DispatchConfig.saveBranding()"><i class="fas fa-save"></i> Save Web App Branding</button>'+
        '<button class="btn-dispatch-warning btn-xs-dispatch" onclick="DispatchConfig.clearBranding()"><i class="fas fa-eraser"></i> Reset</button>'+
      '</div>'+
    '</div></div>'+
    '<div style="display:flex;justify-content:space-between;align-items:flex-start;padding:10px 0;border-bottom:1px solid var(--border);gap:12px;flex-wrap:wrap">'+
    '<div><label style="font-weight:700;font-size:.85rem;display:block">Portal Branding</label>'+
    '<span style="font-size:.72rem;color:var(--text-muted)">Set the logo, name, and title used only by the tech magic-link portal.</span></div>'+
    '<div style="display:flex;flex-direction:column;gap:6px;min-width:300px">'+
      '<input type="text" id="cfgPortalBrandTitle" value="'+escapeHtml(savedPortalBrandTitle)+'" placeholder="Portal title (e.g. Dispatch Portal)" style="font-family:var(--font-mono);font-size:.74rem;padding:5px 9px;border-radius:7px;border:1px solid var(--border);background:var(--bg-input);color:var(--text-primary)">'+
      '<input type="text" id="cfgPortalBrandName" value="'+escapeHtml(savedPortalBrandName)+'" placeholder="Portal brand name" style="font-family:var(--font-mono);font-size:.74rem;padding:5px 9px;border-radius:7px;border:1px solid var(--border);background:var(--bg-input);color:var(--text-primary)">'+
      '<input type="text" id="cfgPortalBrandLogoUrl" value="'+escapeHtml(savedPortalBrandLogo)+'" placeholder="https://.../portal-logo.png" style="font-family:var(--font-mono);font-size:.74rem;padding:5px 9px;border-radius:7px;border:1px solid var(--border);background:var(--bg-input);color:var(--text-primary)">'+
      '<div style="display:flex;gap:6px;flex-wrap:wrap">'+
        '<button class="btn-dispatch-secondary btn-xs-dispatch" onclick="DispatchConfig.savePortalBranding()"><i class="fas fa-save"></i> Save Portal Branding</button>'+
        '<button class="btn-dispatch-warning btn-xs-dispatch" onclick="DispatchConfig.resetPortalBranding()"><i class="fas fa-eraser"></i> Reset Portal Branding</button>'+
      '</div>'+
    '</div></div>'+
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
  var ICONS={"I'm On My Way":'🚗',"Let's Schedule a Visit":'📅','Arriving Today':'📬','Portal Opened':'👀','Magic Link Generated':'🔗','Scheduled Visit':'🗓️','Rescheduled Visit':'🔁','Portal Note':'📝','No Contact Reported':'📵','Reassign Requested':'🔄','Pre-Reassign Warning':'⚠️','Tier 2 Blast':'📣','Auto Reassigned':'♻️'};
  var html='';
  comms.slice(0,80).forEach(function(c){
    var icon=ICONS[c.template_used]||'💬';
    var metaParts=[];
    if (c.tenant_phone) metaParts.push('To: ' + escapeHtml(c.tenant_phone));
    if (c.rc_message_id) metaParts.push('RC: ' + escapeHtml(String(c.rc_message_id).substring(0,12)));
    if (c.magic_link) metaParts.push('<a href="' + escapeHtml(c.magic_link) + '" target="_blank" rel="noopener noreferrer">Open</a>');
    html+='<div style="display:flex;align-items:flex-start;gap:10px;padding:10px 0;border-bottom:1px solid var(--border);font-size:.82rem">'+
      '<div style="font-size:1.1rem;margin-top:1px">'+icon+'</div>'+
      '<div style="flex:1;min-width:0">'+
        '<div style="font-weight:600">'+escapeHtml(c.tech_name||'Unknown')+
          (c.template_used?'<span style="margin-left:6px;font-size:.7rem;font-family:var(--font-mono);padding:1px 5px;border-radius:4px;background:var(--accent-dim);color:var(--accent)">'+escapeHtml(c.template_used)+'</span>':'')+
        '</div>'+
        '<div style="font-size:.72rem;color:var(--text-muted);font-family:var(--font-mono);margin-top:2px">'+
          'WO: <strong style="color:var(--accent)">'+escapeHtml(String(c.wo_id||'').substring(0,20))+'</strong>'+
          (metaParts.length?' · '+metaParts.join(' · '):'')+
        '</div>'+
        (c.message_body?'<div style="margin-top:5px;color:var(--text-secondary);font-size:.75rem;line-height:1.45">'+escapeHtml(String(c.message_body).substring(0,180))+'</div>':'')+
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
  markExempt: async function(woId, woNum) {
    if (!await hmConfirm('Mark WO #'+woNum+' as exempt?\n\nThis writes a system note to AppFolio and stops automated reassignment for this WO.', { title: 'Mark Exempt', okLabel: 'Mark Exempt' })) return;
    try {
      await proxyPost('wo_note_create', {
        uuid: woId,
        body_text: ':stop-auto: Manual admin exemption applied via HandyManager Dispatch Control.'
      });
      v9Toast('Exemption applied','WO '+woNum+' is now exempt','success');
      DispatchControl.refresh();
    } catch(e) { v9Toast('Exemption failed',e.message,'danger'); }
  },
  clearExempt: async function(woId) {
    if (!await hmConfirm('Clear exemption for WO '+woId+'?\n\nIt will re-enter automation on the next cron run.', { title: 'Clear Exemption', okLabel: 'Clear Exemption' })) return;
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
  },
  addMonitored: async function(woId, woNum) {
    try {
      var resp = await proxyAction('add_monitored_work_order', { wo_id: woId });
      if (resp.ok) {
        if (!DISPATCH.monitored) DISPATCH.monitored = [];
        DISPATCH.monitored.push({ wo_id: woId, created_at: new Date().toISOString() });
        v9Toast('WO monitored', 'Now tracking WO #' + woNum, 'success');
        renderDispatchQueue(DISPATCH.queue);
      } else {
        v9Toast('Monitor failed', resp.error || 'Unknown error', 'danger');
      }
    } catch(e) { v9Toast('Monitor failed', e.message, 'danger'); }
  },
  removeMonitored: async function(woId) {
    try {
      var resp = await proxyAction('remove_monitored_work_order', { wo_id: woId });
      if (resp.ok) {
        DISPATCH.monitored = (DISPATCH.monitored || []).filter(function(m) { return String(m.wo_id) !== String(woId); });
        v9Toast('Monitoring stopped', 'Removed WO from monitoring', 'success');
        renderDispatchQueue(DISPATCH.queue);
      } else {
        v9Toast('Remove failed', resp.error || 'Unknown error', 'danger');
      }
    } catch(e) { v9Toast('Remove failed', e.message, 'danger'); }
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
    if(!this._editing&&id.length<30){if(!await hmConfirm('The UUID "'+id+'" looks short. Is this a valid AppFolio user UUID?\n\nThe user must have the Maintenance Tech role enabled or reassignment PATCHes will return 422.', { title: 'UUID Warning', okLabel: 'Continue Anyway' }))return;}
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
    if(!await hmConfirm((currentlyActive?'Deactivate':'Reactivate')+' '+techName+'?', { title: currentlyActive ? 'Deactivate Tech' : 'Reactivate Tech', okLabel: currentlyActive ? 'Deactivate' : 'Reactivate', danger: currentlyActive }))return;
    try {
      var r=await dispatchPost('tech_roster',{tech_id:techId,tech_name:techName,active:currentlyActive?0:1});
      if(r.ok){v9Toast(techName+(currentlyActive?' deactivated':' reactivated'),'',currentlyActive?'warning':'success');DispatchControl.refresh();}
      else v9Toast('Action failed',r.error,'danger');
    }catch(e){v9Toast('Action failed',e.message,'danger');}
  },
  toggleHidden: async function(techId, techName, currentlyHidden) {
    if (!await hmConfirm((currentlyHidden ? 'Unhide ' : 'Hide ') + techName + ' from active dispatch roster?', { title: currentlyHidden ? 'Unhide Assignee' : 'Hide Assignee', okLabel: currentlyHidden ? 'Unhide' : 'Hide' })) return;
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
  _readLivePayload: function() {
    return {
      wo_id: String((document.getElementById('dispatchLiveWoId')||{}).value || '').trim(),
      wo_number: String((document.getElementById('dispatchLiveWoNumber')||{}).value || '').trim(),
      tech_id: String((document.getElementById('dispatchLiveTechId')||{}).value || '').trim(),
      tech_name: String((document.getElementById('dispatchLiveTechName')||{}).value || '').trim(),
      tech_phone: String((document.getElementById('dispatchLiveTechPhone')||{}).value || '').trim(),
      tenant_name: String((document.getElementById('dispatchLiveTenantName')||{}).value || '').trim(),
      tenant_phone: String((document.getElementById('dispatchLiveTenantPhone')||{}).value || '').trim(),
      property_address: String((document.getElementById('dispatchLiveAddress')||{}).value || '').trim()
    };
  },
  generateLiveLink: async function(sendSms) {
    var resultEl = document.getElementById('dispatchLiveLinkResult');
    var btn = document.getElementById(sendSms ? 'btnDispatchSendLiveLink' : 'btnDispatchGenerateLiveLink');
    var payload = this._readLivePayload();
    if (!payload.wo_id || !payload.tech_id || !payload.tech_name || !payload.tenant_phone || !payload.property_address) {
      if (resultEl) resultEl.textContent = 'Required: WO UUID, Tech UUID, Tech name, Resident phone, and Property address.';
      v9Toast('Missing portal fields', 'Fill the required work order, tech, resident, and address fields first', 'warning');
      return;
    }
    if (sendSms && !/^\+\d{10,15}$/.test(payload.tech_phone)) {
      if (resultEl) resultEl.textContent = 'Generate + Send requires a valid tech phone in E.164 format.';
      v9Toast('Invalid tech phone', 'Use a format like +15551234567', 'warning');
      return;
    }

    var originalLabel = btn ? btn.innerHTML : '';
    if (btn) {
      btn.disabled = true;
      btn.innerHTML = '<i class="fas fa-circle-notch fa-spin"></i> Working…';
    }
    if (resultEl) resultEl.textContent = sendSms ? 'Generating and sending live magic link…' : 'Generating live magic link…';

    try {
      var response = await dispatchPost('generate_magic_link', Object.assign({}, payload, {
        send_sms: !!sendSms
      }));
      if (!response || !response.ok) throw new Error((response && response.error) || 'Magic link generation failed');
      if (!response.short_code || !response.magic_link || response.magic_link.indexOf('/s/') === -1) {
        throw new Error('Dispatch short link was not created. Raw portal URLs are blocked.');
      }
      if (resultEl) {
        var openLink = response.magic_link
          ? ' <a href="' + escapeHtml(response.magic_link) + '" target="_blank" rel="noopener noreferrer">Open portal</a>'
          : '';
        var shortLabel = response.short_code ? 'Short code: ' + escapeHtml(response.short_code) + '. ' : '';
        resultEl.innerHTML = shortLabel + (sendSms ? 'Sent live link to tech.' : 'Generated live link.') + openLink;
      }
      v9Toast(sendSms ? 'Live link sent' : 'Live link generated', payload.tech_name, 'success');
      DispatchControl.refresh();
    } catch (e) {
      if (resultEl) resultEl.textContent = 'Action failed: ' + (e.message || e);
      v9Toast('Magic link action failed', e.message || String(e), 'danger');
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.innerHTML = originalLabel;
      }
    }
  },
  sendMagicLinkTest: async function() {
    var phoneEl = document.getElementById('dispatchTestPhone');
    var btn = document.getElementById('btnDispatchSendTestSms');
    var resultEl = document.getElementById('dispatchTestResult');
    var phone = phoneEl ? String(phoneEl.value || '').trim() : '';
    var adminKey = getDispatchAdminKey();
    if (!/^\+\d{10,15}$/.test(phone)) {
      if (resultEl) resultEl.textContent = 'Enter a valid E.164 phone number such as +15551234567.';
      v9Toast('Invalid phone number', 'Use E.164 format like +15551234567', 'warning');
      return;
    }
    if (!adminKey) {
      if (resultEl) resultEl.textContent = 'Set PROXY_ADMIN_KEY in the Database tab before sending test links.';
      v9Toast('Admin key required', 'Set PROXY_ADMIN_KEY in Database tab first. Generate Link does not require this key.', 'warning');
      return;
    }

    var originalLabel = btn ? btn.innerHTML : '';
    if (btn) {
      btn.disabled = true;
      btn.innerHTML = '<i class="fas fa-circle-notch fa-spin"></i> Sending…';
    }
    if (resultEl) resultEl.textContent = 'Sending test magic-link SMS to ' + phone + '…';

    try {
      var response = await proxyPost('send_magic_link_test_sms', {
        key: adminKey,
        phone: phone,
        tech_name: 'Dispatch Test',
        tech_id: 'dispatch-test'
      });
      if (!response || !response.ok) {
        throw new Error((response && response.error) || 'Test SMS send failed');
      }
      if (!response.magic_link || response.magic_link.indexOf('/s/') === -1) {
        throw new Error('Dispatch short link was not created. Raw portal URLs are blocked.');
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
      var message = e && e.message ? e.message : String(e || 'Unknown error');
      if (/failed to fetch/i.test(message)) {
        message = 'Request could not reach the proxy. Check Proxy URL/session in Vault, then retry.';
      }
      if (resultEl) resultEl.textContent = 'Send failed: ' + message;
      v9Toast('Test link failed', message, 'danger');
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
  saveBranding: async function() {
    var nameEl = document.getElementById('cfgBrandName');
    var logoEl = document.getElementById('cfgBrandLogoUrl');
    var brandName = nameEl ? String(nameEl.value || '').trim() : '';
    var brandLogo = logoEl ? String(logoEl.value || '').trim() : '';
    try {
      if (brandName) {
        localStorage.setItem('hm_brand_name', brandName);
        await saveDispatchConfigKey('brand_name', brandName);
      }
      localStorage.setItem('hm_brand_logo', brandLogo);
      await saveDispatchConfigKey('brand_logo_url', brandLogo);
      applyBrandLogo();
      v9Toast('Web app branding saved', brandLogo ? 'Main webapp logo updated' : 'Brand name updated', 'success');
    } catch (e) {
      v9Toast('Branding save failed', e.message || String(e), 'danger');
    }
  },
  clearBranding: async function() {
    var nameEl = document.getElementById('cfgBrandName');
    var logoEl = document.getElementById('cfgBrandLogoUrl');
    if (nameEl) nameEl.value = BRAND_NAME_DEFAULT;
    if (logoEl) logoEl.value = '';
    try {
      localStorage.setItem('hm_brand_name', BRAND_NAME_DEFAULT);
      localStorage.removeItem('hm_brand_logo');
      await saveDispatchConfigKey('brand_name', BRAND_NAME_DEFAULT);
      await saveDispatchConfigKey('brand_logo_url', '');
      applyBrandLogo();
      v9Toast('Web app branding reset', 'Default main logo and brand name restored', 'success');
    } catch (e) {
      v9Toast('Branding reset failed', e.message || String(e), 'danger');
    }
  },
  savePortalBranding: async function() {
    var titleEl = document.getElementById('cfgPortalBrandTitle');
    var nameEl = document.getElementById('cfgPortalBrandName');
    var logoEl = document.getElementById('cfgPortalBrandLogoUrl');
    var brandTitle = titleEl ? String(titleEl.value || '').trim() : '';
    var brandName = nameEl ? String(nameEl.value || '').trim() : '';
    var brandLogo = logoEl ? String(logoEl.value || '').trim() : '';
    try {
      if (brandTitle) {
        localStorage.setItem('hm_portal_brand_title', brandTitle);
        await saveDispatchConfigKey('portal_brand_title', brandTitle);
      }
      if (brandName) {
        localStorage.setItem('hm_portal_brand_name', brandName);
        await saveDispatchConfigKey('portal_brand_name', brandName);
      }
      localStorage.setItem('hm_portal_brand_logo', brandLogo);
      await saveDispatchConfigKey('portal_brand_logo_url', brandLogo);
      v9Toast('Portal branding saved', 'Magic-link portal branding updated', 'success');
    } catch (e) {
      v9Toast('Portal branding save failed', e.message || String(e), 'danger');
    }
  },
  resetPortalBranding: async function() {
    var titleEl = document.getElementById('cfgPortalBrandTitle');
    var nameEl = document.getElementById('cfgPortalBrandName');
    var logoEl = document.getElementById('cfgPortalBrandLogoUrl');
    if (titleEl) titleEl.value = 'Dispatch Portal';
    if (nameEl) nameEl.value = PORTAL_BRAND_NAME_DEFAULT;
    if (logoEl) logoEl.value = PORTAL_BRAND_LOGO_DEFAULT;
    try {
      localStorage.setItem('hm_portal_brand_title', 'Dispatch Portal');
      localStorage.setItem('hm_portal_brand_name', PORTAL_BRAND_NAME_DEFAULT);
      localStorage.setItem('hm_portal_brand_logo', PORTAL_BRAND_LOGO_DEFAULT);
      await saveDispatchConfigKey('portal_brand_title', 'Dispatch Portal');
      await saveDispatchConfigKey('portal_brand_name', PORTAL_BRAND_NAME_DEFAULT);
      await saveDispatchConfigKey('portal_brand_logo_url', PORTAL_BRAND_LOGO_DEFAULT);
      v9Toast('Portal branding reset', 'Portal branding restored to configured defaults', 'success');
    } catch (e) {
      v9Toast('Portal branding reset failed', e.message || String(e), 'danger');
    }
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

    // Primary path: direct AppFolio users endpoint via proxy action=users.
    // This keeps roster sync working even when dispatch sync aliases are unavailable.
    try {
      if (!PROPERTY_GROUPS || PROPERTY_GROUPS.length === 0) {
        try { await fetchPropertyGroups(); } catch (_) {}
      }

      var branchByTechId = {};
      try {
        var woForBranch = await proxyAction('work_orders', { days: '90' });
        var branchCandidates = extractAssigneeCandidatesFromWorkOrders((woForBranch && woForBranch.results) || []);
        branchCandidates.forEach(function(c) {
          if (c && c.tech_id && !branchByTechId[c.tech_id]) branchByTechId[c.tech_id] = c.branch || 'phoenix';
        });
      } catch (_) {}

      var usersResp = await proxyAction('users', {
        role: 'maintenance_tech',
        since: '2020-01-01T00:00:00Z'
      });
      var users = (usersResp && usersResp.results) || [];
      if (Array.isArray(users) && users.length) {
        var countFromUsers = 0;
        for (var ui = 0; ui < users.length; ui++) {
          var u = users[ui] || {};
          var techId = String(u.id || u.Id || '').trim();
          if (!techId || isTechHidden(techId)) continue;

          var inferredBranch = branchByTechId[techId] || (DISPATCH.activeBranch === 'all' ? 'phoenix' : DISPATCH.activeBranch);
          if (DISPATCH.activeBranch !== 'all' && inferredBranch !== DISPATCH.activeBranch) continue;

          var techName = String(
            u.name ||
            [u.firstName || u.FirstName || '', u.lastName || u.LastName || ''].join(' ').trim() ||
            u.email || u.Email || techId
          ).trim();

          var upFromUsers = await dispatchPost('tech_roster', {
            tech_id: techId,
            tech_name: techName,
            tier: inferredBranch === 'phoenix' ? 1 : 2,
            geo_zone: inferredBranch,
            active: 1
          });
          if (upFromUsers && upFromUsers.ok) countFromUsers += 1;
        }

        if (countFromUsers > 0) {
          v9Toast('Assignees synced', countFromUsers + ' techs imported from AppFolio users', 'success');
          DispatchControl.refresh();
          return;
        }
      }
    } catch (usersErr) {
      lastErr = usersErr.message || String(usersErr);
    }

    for (var i = 0; i < actions.length; i++) {
      try {
        var resp = await dispatchPost(actions[i], payload);
        if (resp && resp.ok) {
          v9Toast('Assignees synced', String(resp.inserted || resp.count || 0) + ' techs synchronized', 'success');
          DispatchControl.refresh();
          return;
        }
        lastErr = (resp && resp.error) ? resp.error : 'not supported';
      } catch (e) {
        lastErr = e.message || String(e);
      }
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
        DISPATCH.monitored=queueData.monitored_work_orders||[];
        DISPATCH.stats =queueData.stats        ||{};
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
    var liveGenerateBtn=document.getElementById('btnDispatchGenerateLiveLink');
    if(liveGenerateBtn)liveGenerateBtn.addEventListener('click',function(){DispatchComms.generateLiveLink(false);});
    var liveSendBtn=document.getElementById('btnDispatchSendLiveLink');
    if(liveSendBtn)liveSendBtn.addEventListener('click',function(){DispatchComms.generateLiveLink(true);});
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
    var btnSyncRoster=document.getElementById('btnSyncRosterFromAf');
    if(btnSyncRoster)btnSyncRoster.addEventListener('click',function(){DispatchConfig.syncAssignees();});
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

/* =================================================================
  DISPATCH ENGINE — Implementation Notes
  Tech UUIDs must hold "Maintenance Tech" role in AppFolio or
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
    var fallbackDesc=meta.resourceName||evt.resource_name||'';
    // If resource is identified but not yet resolved, mark pending so enrichment pass can update it.
    if(!fallbackDesc&&meta.resourceId&&meta.resourceType)fallbackDesc='\u29d7 Resolving\u2026';
    view={title:fallbackTitle||'Webhook Event',desc:fallbackDesc,severity:'info',icon:'fa-plug',color:'var(--purple)'};
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
  var POLL_MS=20000;

  function _getEl(id){return document.getElementById(id);}

  function _updateBadge(){
    var n=Math.max(0,_unseenCount); var t=n>99?'99+':String(n);
    var b=_getEl('feed-badge'), fb=_getEl('fab-badge');
    if(b){b.textContent=t;b.style.display=n>0?'':'none';}
    if(fb){fb.textContent=t;fb.style.display=n>0?'':'none';}
  }
  function _clearUnseen(){_unseenCount=0;_updateBadge();}

  async function _seedLastId(){
    if (!isProxySessionReady()) return;
    try{var d=await proxyAction('webhook_events',{limit:1});if(d.ok&&d.events&&d.events.length)_lastId=d.events[0].id||0;}catch(e){}
  }

  async function _poll(){
    if(_isPolling||!API_PROXY)return;
    if(!isProxySessionReady()) return;
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
          var SECMAP={'sec-workorders':['work_orders','turn_work_orders','recent_tasks'],'sec-turnboard':['turns','turn_work_orders'],'sec-dashboard':['work_orders','turns','upcoming_moveouts','bills'],'sec-billing':['bills'],'sec-inspections':['inspections'],'sec-vendors':['vendors']};
          Object.keys(SECMAP).forEach(function(secId){
            var sec=document.getElementById(secId);
            if(!sec||!sec.classList.contains('active'))return;
            if(!keyArr.some(function(k){return SECMAP[secId].indexOf(k)!==-1;}))return;
            if(secId==='sec-workorders'){try{renderWorkOrders();}catch(e){}}
            else if(secId==='sec-turnboard'){try{renderTurnBoard();}catch(e){}}
            else if(secId==='sec-dashboard'){try{renderDashboardKPIs();}catch(e){} try{renderTurnDashboardStrip();}catch(e){} try{renderActivityFeed();}catch(e){}}
            else if(secId==='sec-billing'){try{renderBillsSection();}catch(e){}}
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

      // Progressive enrichment: resolve sparse-payload events asynchronously.
      // Cap at 10 per poll cycle to stay well under AppFolio rate limits.
      (async function _enrichLiveEvents(newEvts){
        var PENDING='\u29d7 Resolving\u2026';
        var toEnrich=newEvts.filter(function(v){
          if(!v.resource_id||!v.resource_type)return false;
          var k=webhookResolveKey(v.resource_type,v.resource_id);
          return k&&!WEBHOOK_RESOURCE_CACHE[k]&&!WEBHOOK_RESOURCE_INFLIGHT[k];
        });
        if(!toEnrich.length)return;
        var changed=false;
        for(var i=0;i<Math.min(10,toEnrich.length);i++){
          var item=toEnrich[i];
          var got=await resolveWebhookResource(item.resource_type,item.resource_id);
          if(got&&got.summary){
            var resolvedName=(got.summary.title||'')+(got.summary.reference&&got.summary.title?'':got.summary.reference?'#'+got.summary.reference:'');
            if(resolvedName){
              _events.forEach(function(ev){
                if(ev.resource_id===item.resource_id&&ev.resource_type===item.resource_type){
                  if(!ev.desc||ev.desc===PENDING){ev.desc=resolvedName;changed=true;}
                }
              });
            }
          }
          await sleep(140+Math.floor(Math.random()*40));
        }
        if(changed)_renderFeed();
      })(decoded);
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
      v9Toast('Resolving\u2026',rType+' / '+String(rId).substring(0,14),'info',2500);
      proxyAction('webhook_resolve',{resource_type:rType,resource_id:rId}).then(function(data){
        if(data.ok&&data.record){
          var s=data.summary||{};
          var resolvedName=(s.title||'')+(s.reference?(' #'+s.reference):'');
          // Update matching events in feed with resolved name, then re-render.
          var changed=false;
          var PENDING='\u29d7 Resolving\u2026';
          _events.forEach(function(ev){
            if(ev.resource_id===rId&&ev.resource_type===rType){
              if(!ev.desc||ev.desc===PENDING||ev.desc==='\u2014\u00a0Unresolved'){
                ev.desc=resolvedName||s.title||rType;changed=true;
              }
            }
          });
          if(changed)_renderFeed();
          showItemDetail((s.title||rType)+(s.reference?' #'+s.reference:''),
            [{section:'Resolved Record',icon:'fa-link'},{label:'Title',value:s.title||'\u2014'},{label:'Status',value:s.status||'\u2014'},{label:'UUID',value:rId||'\u2014'}],null);
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
    // Check for version mismatch before initializing
    if (localStorage.getItem('hm_version_mismatch') === '1') {
      localStorage.removeItem('hm_version_mismatch');
      location.reload();
      return;
    }

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
        var woKeys=['work_orders','turn_work_orders','turns','upcoming_moveouts','bills'];
        var detail=(e&&e.detail&&e.detail.keys)?e.detail.keys:[];
        if(detail.some(function(k){return woKeys.indexOf(k)!==-1;})){try{renderDashboardKPIs();}catch(err){}}
        if(detail.some(function(k){return k==='bills';})){try{renderBillsSection();}catch(err){}}
      });
    });
    document.addEventListener('visibilitychange',function(){if(!document.hidden&&_pollTimer)_poll();});
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',_boot);
  else _boot();
})();
