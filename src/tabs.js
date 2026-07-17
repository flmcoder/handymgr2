export function createTabsModule(deps) {
  var $$ = deps.$$;
  var getAccessRole = deps.getAccessRole;
  var setMobileNavLabelFromActiveTab = deps.setMobileNavLabelFromActiveTab;
  var closeMobileNav = deps.closeMobileNav;
  var syncSubtabDock = deps.syncSubtabDock;

  function isTabAllowedForRole(tabName) {
    var role = getAccessRole();
    if (role === 'vendors') return tabName === 'vendors';
    if (role === 'pm_readonly') {
      var pmAllowedTabs = ['dashboard', 'workorders', 'estimates', 'billing', 'occupancy', 'properties', 'turnboard', 'vendors', 'inspections', 'errors'];
      return pmAllowedTabs.indexOf(tabName) !== -1;
    }
    if (role === 'manager') {
      var gmAllowedTabs = ['dashboard', 'workorders', 'estimates', 'routing', 'billing', 'occupancy', 'properties', 'turnboard', 'vendors', 'inspections', 'managerreview', 'payroll', 'dbadmin', 'errors'];
      return gmAllowedTabs.indexOf(tabName) !== -1;
    }
    return true;
  }

  function setActiveMainSection(tabName) {
    var targetId = 'sec-' + String(tabName || 'dashboard');
    $$('.section').forEach(function(section) {
      var isActive = section.id === targetId;
      section.classList.toggle('active', isActive);
      if (isActive) {
        section.style.display = section.classList.contains('hm-neo-dashboard') ? 'flex' : 'block';
      } else {
        section.style.display = 'none';
      }
    });
  }

  function forceActiveTab(tabName) {
    $$('.nav-tab').forEach(function(t) {
      t.classList.toggle('active', t.getAttribute('data-tab') === tabName);
    });
    setActiveMainSection(tabName);
    setMobileNavLabelFromActiveTab();
    closeMobileNav();
    syncSubtabDock();
  }

  return {
    isTabAllowedForRole: isTabAllowedForRole,
    setActiveMainSection: setActiveMainSection,
    forceActiveTab: forceActiveTab,
  };
}
