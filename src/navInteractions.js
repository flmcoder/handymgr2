export function createNavInteractionsModule(deps) {
  var $$ = deps.$$;
  var documentRef = deps.documentRef;
  var setActiveMainSection = deps.setActiveMainSection;
  var resetAppScrollToTop = deps.resetAppScrollToTop;
  var renderWorkOrders = deps.renderWorkOrders;
  var openTurnBoardDetail = deps.openTurnBoardDetail;
  var getMoveOutSection = deps.getMoveOutSection;
  var setWorkOrderPriority = deps.setWorkOrderPriority;
  var setWorkOrderFilter = deps.setWorkOrderFilter;
  var setWorkOrderPriorityControl = deps.setWorkOrderPriorityControl;
  var setWorkOrderFilterControl = deps.setWorkOrderFilterControl;
  var renderDeferredVendors = deps.renderDeferredVendors;
  var shouldRenderDeferredVendors = deps.shouldRenderDeferredVendors;

  function activateTab(tabName) {
    $$('.nav-tab').forEach(function(tab) { tab.classList.remove('active'); });
    var tabEl = documentRef.querySelector('[data-tab="' + tabName + '"]');
    if (tabEl) tabEl.classList.add('active');
    setActiveMainSection(tabName);
    resetAppScrollToTop();
  }

  function bindClickableKpis() {
    $$('.kpi-clickable[data-kpi]').forEach(function(card) {
      card.addEventListener('click', function() {
        var kpi = this.getAttribute('data-kpi');
        if (kpi === 'open' || kpi === 'urgent' || kpi === 'flagged') {
          activateTab('workorders');
          if (kpi === 'urgent') {
            setWorkOrderPriority('Urgent');
            setWorkOrderPriorityControl('Urgent');
          } else if (kpi === 'flagged') {
            setWorkOrderFilter('flagged');
            setWorkOrderFilterControl('flagged');
          }
          renderWorkOrders();
          return;
        }

        if (kpi === 'turns') {
          activateTab('turnboard');
          return;
        }

        if (kpi === 'moveouts') {
          var moveOutSection = getMoveOutSection();
          if (moveOutSection) moveOutSection.scrollIntoView({ behavior: 'smooth' });
        }
      });
    });
  }

  function bindVendorDeferredRender() {
    if (!documentRef.body) return;
    documentRef.body.addEventListener('click', function(e) {
      var tab = e.target.closest('.nav-tab[data-tab="vendors"]');
      if (tab && shouldRenderDeferredVendors()) {
        setTimeout(function() {
          renderDeferredVendors();
        }, 0);
      }
    });
  }

  function bindDashboardTurnStripClick() {
    var strip = documentRef.getElementById('dashTurnStrip');
    if (!strip) return;
    strip.addEventListener('click', function(e) {
      var card = e.target.closest('[data-turndash-open]');
      if (!card) return;
      openTurnBoardDetail(card.getAttribute('data-turndash-open'));
    });
  }

  return {
    bindClickableKpis: bindClickableKpis,
    bindVendorDeferredRender: bindVendorDeferredRender,
    bindDashboardTurnStripClick: bindDashboardTurnStripClick,
  };
}
