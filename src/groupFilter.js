export function createGroupFilterModule(deps) {
  var $ = deps.$;
  var documentRef = deps.documentRef;
  var requestAnimationFrameRef = deps.requestAnimationFrameRef;
  var tabRenderMap = deps.tabRenderMap || {};

  function getActiveTabName() {
    var activeTab = documentRef.querySelector('.nav-tab.active[data-tab], .tab-btn.active[data-tab], [data-tab].active');
    var tabName = activeTab ? activeTab.getAttribute('data-tab') : '';
    if (!tabName) {
      var activeSection = documentRef.querySelector('.section.active');
      if (activeSection && activeSection.id) {
        var sectionName = String(activeSection.id).replace(/^sec-/, '');
        if (tabRenderMap[sectionName]) tabName = sectionName;
      }
    }
    if (!tabName || !tabRenderMap[tabName]) return 'dashboard';
    return tabName;
  }

  function applyGroupFilterChange() {
    var tabName = getActiveTabName();
    Object.keys(tabRenderMap).forEach(function(name) {
      deps.groupFilterDirty[name] = true;
    });
    deps.groupFilterDirty[tabName] = false;

    requestAnimationFrameRef(function() {
      try {
        if (tabRenderMap[tabName]) tabRenderMap[tabName]();
      } catch (e) { /* safe */ }
      if (tabName !== 'dashboard') {
        try { deps.renderDashboardKPIs(); } catch (e) { /* safe */ }
      }
    });

    deps.updateGlobalGroupIndicator();
    deps.emitGroupFilterChanged({
      groupName: deps.getCurrentPropertyGroup() || '',
      forcedGroupUuid: deps.getForcedPropertyGroupUuid() || ''
    });
  }

  function clearPropertyGroupFilters() {
    if (deps.getForcedPropertyGroupUuid()) {
      deps.enforceScopedPropertyGroup();
      return;
    }

    deps.setCurrentPropertyGroup('');
    deps.resetGroupMissLogCount();

    var globalGroupEl = $('#globalGroupFilter');
    if (globalGroupEl) globalGroupEl.value = '';

    deps.setCurrentTurnPipeGroup('');
    var turnGroupEl = $('#turnPipeGroup');
    if (turnGroupEl) turnGroupEl.value = '';

    applyGroupFilterChange();
  }

  function bindGlobalGroupFilterControls() {
    var globalGroupEl = $('#globalGroupFilter');
    if (globalGroupEl) {
      globalGroupEl.addEventListener('change', function() {
        if (deps.getForcedPropertyGroupUuid()) {
          deps.enforceScopedPropertyGroup();
          return;
        }
        deps.setCurrentPropertyGroup(this.value);
        deps.resetGroupMissLogCount();
        applyGroupFilterChange();
      });
    }

    var globalClearBtn = $('#globalGroupClear');
    if (globalClearBtn) {
      globalClearBtn.addEventListener('click', function() {
        clearPropertyGroupFilters();
      });
    }
  }

  function bindSyncGroupsButton() {
    var btnGlobalLoadGroups = $('#btnGlobalLoadGroups');
    if (!btnGlobalLoadGroups) return;

    btnGlobalLoadGroups.addEventListener('click', async function(evt) {
      var wantDiag = evt && evt.shiftKey;
      var wantCopyCommand = evt && (evt.altKey || evt.ctrlKey || evt.metaKey);
      btnGlobalLoadGroups.disabled = true;
      btnGlobalLoadGroups.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Syncing...';
      try {
        await deps.syncPropertyGroupsToLocalDb();
        var ok = await deps.fetchPropertyGroups();
        if (ok) {
          deps.populateGroupFilters();
          if (wantCopyCommand && navigator.clipboard) {
            var cmd = deps.buildPropertyGroupsProxyCurl();
            if (cmd) {
              await navigator.clipboard.writeText(cmd);
              deps.showToast('Copied property-groups proxy curl command');
            }
          }
          if (wantDiag && deps.getPgDiagnostics()) {
            var diag = deps.getPgDiagnostics();
            var lines = [
              'Groups: ' + deps.getPropertyGroupsCount(),
              'UUID map: ' + diag.uuidMapSize + ' entries',
              'DB names: ' + diag.dbNameCount,
              'UUID hits/misses: ' + diag.uuidHits + '/' + diag.uuidMisses,
              'Name bridges: ' + diag.nameMatches + '/' + deps.getPropertiesCount(),
              'ID map: ' + diag.idMatches + ' props',
              'Portfolio matches: ' + diag.portfolioMatches,
              'nameMap: ' + deps.getNameMapCount(),
              'idMap: ' + deps.getIdMapCount()
            ];
            if (diag.errors.length > 0) lines.push('ERRORS: ' + diag.errors.join('; '));
            deps.showToast(lines.join(' | '), 12000);
          } else {
            deps.showToast('Property groups synced - ' + deps.getPropertyGroupsCount() + ' groups, ' +
              deps.getNameMapCount() + ' name mappings, ' +
              deps.getIdMapCount() + ' ID mappings');
          }
        } else {
          deps.showToast('Sync completed, but failed to load property groups - check console for [PG] logs');
        }
      } catch (e) {
        deps.showToast('Error: ' + (e.message || e));
      } finally {
        btnGlobalLoadGroups.disabled = false;
        btnGlobalLoadGroups.innerHTML = '<i class="fas fa-sync-alt"></i> Sync Groups';
      }
    });
  }

  return {
    getTabRenderMap: function() { return tabRenderMap; },
    applyGroupFilterChange: applyGroupFilterChange,
    clearPropertyGroupFilters: clearPropertyGroupFilters,
    bindGlobalGroupFilterControls: bindGlobalGroupFilterControls,
    bindSyncGroupsButton: bindSyncGroupsButton,
  };
}
