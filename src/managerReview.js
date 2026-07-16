export function createManagerReviewModule(deps) {
  var $ = deps.$;
  var escapeHtml = deps.escapeHtml;
  var formatDate = deps.formatDate;
  var currency = deps.currency;
  var renewalStatusBadge = deps.renewalStatusBadge;
  var getEffectiveGroupUuid = deps.getEffectiveGroupUuid;
  var resolveGroupNameFromUuid = deps.resolveGroupNameFromUuid;
  var renderDashboardInsightChart = deps.renderDashboardInsightChart;
  var setSectionBusy = deps.setSectionBusy;
  var apiFetch = deps.apiFetch;
  var loadingHtml = deps.loadingHtml;

  var ticklerRows = [];
  var renewalRows = [];
  var ledgerRows = [];
  var estimateRows = [];
  var loading = false;
  var loadedKey = '';
  var lastError = '';
  var sourceMatch = ['flm estimates', 'flm phx estimate', 'flr', 'fort l prop mgmt', 'flm lowell maintenance'];

  function ensureManagerReviewDateRange() {
    var fromEl = $('#mgrReviewFrom');
    var toEl = $('#mgrReviewTo');
    if (!fromEl || !toEl) return;
    if (!toEl.value) toEl.value = dateInputValue(new Date());
    if (!fromEl.value) {
      var fromDate = new Date();
      fromDate.setDate(fromDate.getDate() - 90);
      fromEl.value = dateInputValue(fromDate);
    }
  }

  function dateInputValue(dateLike) {
    var date = dateLike instanceof Date ? new Date(dateLike.getTime()) : new Date(dateLike || Date.now());
    if (isNaN(date.getTime())) date = new Date();
    return date.toISOString().slice(0, 10);
  }

  function getManagerReviewScope() {
    ensureManagerReviewDateRange();
    var fromEl = $('#mgrReviewFrom');
    var toEl = $('#mgrReviewTo');
    var searchEl = $('#mgrReviewSearch');
    var fromDate = String(fromEl ? fromEl.value : '').trim();
    var toDate = String(toEl ? toEl.value : '').trim();
    var search = String(searchEl ? searchEl.value : '').trim().toLowerCase();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(fromDate) || !/^\d{4}-\d{2}-\d{2}$/.test(toDate)) {
      throw new Error('Manager review date range is required');
    }
    if (fromDate > toDate) {
      throw new Error('From date must be before To date');
    }
    var groupUuid = String(getEffectiveGroupUuid() || '').trim();
    return {
      fromDate: fromDate,
      toDate: toDate,
      fromMonth: fromDate.slice(0, 7),
      toMonth: toDate.slice(0, 7),
      search: search,
      groupUuid: groupUuid,
    };
  }

  function isManagerReviewSourceMatch(raw) {
    var text = String(raw || '').trim().toLowerCase();
    if (!text) return false;
    return sourceMatch.some(function(key) { return text.indexOf(key) !== -1; });
  }

  function identifyLeasingChargeType(row) {
    var combined = [
      row && row.account,
      row && row.account_name,
      row && row.description,
      row && row.notes,
      row && row.memo,
      row && row.comment,
      row && row.charge_description,
      row && row.gl_account,
    ].join(' ').toLowerCase();
    var hasRenewal = /renew/.test(combined);
    var hasMoveIn = /move[\s\-_]?in/.test(combined);
    if (hasRenewal && hasMoveIn) return 'Renewal + Move-In';
    if (hasRenewal) return 'Renewal';
    if (hasMoveIn) return 'Move-In';
    return '';
  }

  function normalizeManagerTicklerRows(rows) {
    return (rows || []).map(function(row) {
      return {
        occurred_on: row.occurred_on || row.occurred_date || row.date || row.tickler_date || '',
        event: row.event || row.tenant_event || row.tickler_event || '',
        property_name: row.property_name || row.property || '',
        property_id: row.property_id || '',
        tenant_name: row.tenant_name || row.tenant || row.name || '',
        lease_from: row.lease_from || row.lease_start || row.start_on || '',
        lease_to: row.lease_to || row.lease_end || row.end_on || '',
      };
    });
  }

  function normalizeManagerRenewalRows(rows) {
    return (rows || []).map(function(row) {
      return {
        status: row.status || '',
        property_name: row.property_name || row.property || '',
        property_id: row.property_id || '',
        unit_name: row.unit_name || row.unit || '',
        tenant_name: row.tenant_name || row.tenant || '',
        lease_start: row.lease_start || row.start_on || '',
        lease_end: row.lease_end || row.end_on || '',
        previous_rent: row.previous_rent || row.previous_amount || '',
        rent: row.rent || row.amount || '',
        turn: row.turn || row.turn_status || '',
      };
    });
  }

  function normalizeManagerLedgerRows(rows) {
    return (rows || []).map(function(row) {
      var chargeType = identifyLeasingChargeType(row);
      return {
        posted_on: row.posted_on || row.date || row.occurred_on || '',
        property_name: row.property_name || row.property || '',
        property_id: row.property_id || '',
        unit_name: row.unit_name || row.unit || '',
        tenant_name: row.tenant_name || row.tenant || '',
        account_name: row.account_name || row.account || row.gl_account || '',
        description: row.description || row.memo || row.notes || row.comment || '',
        amount: row.amount || row.net_amount || row.total || row.debit || row.credit || '',
        charge_type: chargeType,
      };
    }).filter(function(row) {
      return !!row.charge_type;
    });
  }

  function buildManagerEstimateRows(estimates, workOrders) {
    var woById = {};
    (workOrders || []).forEach(function(wo) {
      var id = String(wo.id || wo.work_order_uuid || wo.work_order_id || '').trim();
      if (!id) return;
      woById[id] = wo;
    });

    return (estimates || []).map(function(est) {
      var workOrderId = String(est.workOrderId || est.work_order_id || est.wo_id || '').trim();
      var wo = woById[workOrderId] || {};
      var source = String(est.source || est.vendor_name || wo.vendor_name || wo.vendor || '').trim();
      var status = String(est.currentStatus || est.current_status || '').trim();
      var statusLower = status.toLowerCase();
      var pipelineMatch = /approval|approve|pending|requested|awaiting/.test(statusLower);
      var sourceMatched = isManagerReviewSourceMatch(source);
      return {
        estimate_id: est.estimateId || est.estimate_id || '',
        work_order_number: est.workOrderNumber || est.work_order_number || wo.wo_number || '',
        source: source,
        status: status,
        property_group_id: est.propertyGroupId || est.property_group_id || wo.property_group_id || '',
        vendor_name: wo.vendor_name || wo.vendor || est.vendor_name || '',
        updated_at: est.updatedAt || est.updated_at || est.createdAt || est.created_at || '',
        _source_match: sourceMatched,
        _pipeline_match: pipelineMatch,
      };
    }).filter(function(row) {
      return row._source_match && row._pipeline_match;
    });
  }

  function getManagerReviewRequestKey() {
    var scope = getManagerReviewScope();
    return JSON.stringify({
      from: scope.fromDate,
      to: scope.toDate,
      groupUuid: scope.groupUuid || '',
      search: scope.search || '',
    });
  }

  function renderManagerReviewCharts() {
    var ticklerCounts = {};
    ticklerRows.forEach(function(row) {
      var key = String(row.event || 'Unknown').trim() || 'Unknown';
      ticklerCounts[key] = (ticklerCounts[key] || 0) + 1;
    });
    var ticklerChartRows = Object.keys(ticklerCounts).sort().map(function(key) {
      return { name: key, value: ticklerCounts[key] };
    });

    var renewalCounts = {};
    renewalRows.forEach(function(row) {
      var key = String(row.status || 'Unknown').trim() || 'Unknown';
      renewalCounts[key] = (renewalCounts[key] || 0) + 1;
    });
    var renewalChartRows = Object.keys(renewalCounts).sort().map(function(key) {
      return { name: key, value: renewalCounts[key] };
    });

    var estimateCounts = {};
    estimateRows.forEach(function(row) {
      var key = String(row.status || 'Unknown').trim() || 'Unknown';
      estimateCounts[key] = (estimateCounts[key] || 0) + 1;
    });
    var estimateChartRows = Object.keys(estimateCounts).sort().map(function(key) {
      return { name: key, value: estimateCounts[key] };
    });

    renderDashboardInsightChart('mgrTicklerChart', ticklerChartRows, 'Tickler Events');
    renderDashboardInsightChart('mgrRenewalChart', renewalChartRows, 'Renewal Status');
    renderDashboardInsightChart('mgrEstimateChart', estimateChartRows, 'Estimate Stage');

    var ticklerMeta = $('#mgrTicklerChartMeta');
    if (ticklerMeta) ticklerMeta.textContent = String(ticklerRows.length || 0) + ' events';
    var renewalMeta = $('#mgrRenewalChartMeta');
    if (renewalMeta) renewalMeta.textContent = String(renewalRows.length || 0) + ' rows';
    var estimateMeta = $('#mgrEstimateChartMeta');
    if (estimateMeta) estimateMeta.textContent = String(estimateRows.length || 0) + ' rows';
  }

  function renderManagerReviewSection(opts) {
    opts = opts || {};
    var ticklerBody = $('#mgrTicklerBody');
    var renewalBody = $('#mgrRenewalBody');
    var ledgerBody = $('#mgrLedgerBody');
    var estimateBody = $('#mgrEstimateBody');
    if (!ticklerBody || !renewalBody || !ledgerBody || !estimateBody) return;

    var scope;
    try {
      scope = getManagerReviewScope();
    } catch (err) {
      lastError = String(err && (err.message || err) || 'Invalid manager review filters');
      ticklerBody.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:20px;color:var(--danger);">' + escapeHtml(lastError) + '</td></tr>';
      renewalBody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:20px;color:var(--danger);">' + escapeHtml(lastError) + '</td></tr>';
      ledgerBody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:20px;color:var(--danger);">' + escapeHtml(lastError) + '</td></tr>';
      estimateBody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:20px;color:var(--danger);">' + escapeHtml(lastError) + '</td></tr>';
      return;
    }

    var requestKey = getManagerReviewRequestKey();
    if (loading) {
      ticklerBody.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:20px">' + loadingHtml('Loading manager review…') + '</td></tr>';
      renewalBody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:20px">' + loadingHtml('Loading manager review…') + '</td></tr>';
      ledgerBody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:20px">' + loadingHtml('Loading manager review…') + '</td></tr>';
      estimateBody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:20px">' + loadingHtml('Loading manager review…') + '</td></tr>';
      return;
    }

    if (!opts.forceRefresh && loadedKey === requestKey) {
      // Keep cached rows and just re-apply local filters.
    } else {
      loading = true;
      lastError = '';
      ticklerBody.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:20px">' + loadingHtml('Loading manager review…') + '</td></tr>';
      renewalBody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:20px">' + loadingHtml('Loading manager review…') + '</td></tr>';
      ledgerBody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:20px">' + loadingHtml('Loading manager review…') + '</td></tr>';
      estimateBody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:20px">' + loadingHtml('Loading manager review…') + '</td></tr>';
      setSectionBusy('sec-managerreview', true, 'Loading manager review…');

      var query = [
        'from=' + encodeURIComponent(scope.fromDate),
        'to=' + encodeURIComponent(scope.toDate),
      ];
      if (scope.groupUuid) query.push('property_group_uuid=' + encodeURIComponent(scope.groupUuid));
      apiFetch('/api/local/manager_review?' + query.join('&')).then(function(data) {
        var payload = data || {};
        ticklerRows = normalizeManagerTicklerRows(Array.isArray(payload.tickler_rows) ? payload.tickler_rows : []);
        renewalRows = normalizeManagerRenewalRows(Array.isArray(payload.renewal_rows) ? payload.renewal_rows : []);
        ledgerRows = normalizeManagerLedgerRows(Array.isArray(payload.ledger_rows) ? payload.ledger_rows : []);
        estimateRows = buildManagerEstimateRows(
          Array.isArray(payload.estimate_rows) ? payload.estimate_rows : [],
          Array.isArray(payload.work_order_rows) ? payload.work_order_rows : []
        );

        loadedKey = requestKey;
        lastError = '';
      }).catch(function(err) {
        ticklerRows = [];
        renewalRows = [];
        ledgerRows = [];
        estimateRows = [];
        lastError = 'Manager review load failed: ' + String(err && (err.message || err) || 'Unknown error');
      }).finally(function() {
        loading = false;
        setSectionBusy('sec-managerreview', false);
        renderManagerReviewSection({ forceRefresh: false });
      });
      return;
    }

    var filteredTickler = ticklerRows.slice();
    var filteredRenewals = renewalRows.slice();
    var filteredLedger = ledgerRows.slice();
    var filteredEstimates = estimateRows.slice();

    if (scope.search) {
      var match = function(row, keys) {
        var hay = keys.map(function(k) { return row && row[k] ? row[k] : ''; }).join(' ').toLowerCase();
        return hay.indexOf(scope.search) !== -1;
      };
      filteredTickler = filteredTickler.filter(function(row) { return match(row, ['event', 'property_name', 'tenant_name', 'lease_from', 'lease_to']); });
      filteredRenewals = filteredRenewals.filter(function(row) { return match(row, ['status', 'property_name', 'unit_name', 'tenant_name', 'lease_start', 'lease_end']); });
      filteredLedger = filteredLedger.filter(function(row) { return match(row, ['property_name', 'unit_name', 'tenant_name', 'account_name', 'description', 'charge_type']); });
      filteredEstimates = filteredEstimates.filter(function(row) { return match(row, ['estimate_id', 'work_order_number', 'source', 'status', 'vendor_name']); });
    }

    if (lastError) {
      var errorMsg = escapeHtml(lastError);
      ticklerBody.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:20px;color:var(--danger);">' + errorMsg + '</td></tr>';
      renewalBody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:20px;color:var(--danger);">' + errorMsg + '</td></tr>';
      ledgerBody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:20px;color:var(--danger);">' + errorMsg + '</td></tr>';
      estimateBody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:20px;color:var(--danger);">' + errorMsg + '</td></tr>';
    } else {
      ticklerBody.innerHTML = filteredTickler.length ? filteredTickler.map(function(row) {
        return '<tr>' +
          '<td>' + escapeHtml(formatDate(row.occurred_on)) + '</td>' +
          '<td>' + escapeHtml(String(row.event || '—')) + '</td>' +
          '<td>' + escapeHtml(String(row.property_name || '—')) + '</td>' +
          '<td>' + escapeHtml(String(row.tenant_name || '—')) + '</td>' +
          '<td>' + escapeHtml(formatDate(row.lease_from)) + '</td>' +
          '<td>' + escapeHtml(formatDate(row.lease_to)) + '</td>' +
        '</tr>';
      }).join('') : '<tr><td colspan="6" style="text-align:center;padding:20px;color:var(--text-muted);">No tenant tickler rows for selected filters.</td></tr>';

      renewalBody.innerHTML = filteredRenewals.length ? filteredRenewals.map(function(row) {
        return '<tr>' +
          '<td>' + renewalStatusBadge(row.status) + '</td>' +
          '<td>' + escapeHtml(String(row.property_name || '—')) + '</td>' +
          '<td>' + escapeHtml(String(row.unit_name || '—')) + '</td>' +
          '<td>' + escapeHtml(String(row.tenant_name || '—')) + '</td>' +
          '<td>' + escapeHtml(formatDate(row.lease_start)) + '</td>' +
          '<td>' + escapeHtml(formatDate(row.lease_end)) + '</td>' +
          '<td style="font-family:var(--font-mono)">' + escapeHtml(currency(row.previous_rent, 2)) + '</td>' +
          '<td style="font-family:var(--font-mono)">' + escapeHtml(currency(row.rent, 2)) + '</td>' +
          '<td>' + escapeHtml(String(row.turn || '—')) + '</td>' +
        '</tr>';
      }).join('') : '<tr><td colspan="9" style="text-align:center;padding:20px;color:var(--text-muted);">No renewal rows for selected filters.</td></tr>';

      ledgerBody.innerHTML = filteredLedger.length ? filteredLedger.map(function(row) {
        return '<tr>' +
          '<td>' + escapeHtml(formatDate(row.posted_on)) + '</td>' +
          '<td>' + escapeHtml(String(row.property_name || '—')) + '</td>' +
          '<td>' + escapeHtml(String(row.unit_name || '—')) + '</td>' +
          '<td>' + escapeHtml(String(row.tenant_name || '—')) + '</td>' +
          '<td>' + escapeHtml(String(row.account_name || '—')) + '</td>' +
          '<td>' + escapeHtml(String(row.description || '—')) + '</td>' +
          '<td style="font-family:var(--font-mono)">' + escapeHtml(currency(row.amount, 2)) + '</td>' +
          '<td><span class="tag normal">' + escapeHtml(String(row.charge_type || '—')) + '</span></td>' +
        '</tr>';
      }).join('') : '<tr><td colspan="8" style="text-align:center;padding:20px;color:var(--text-muted);">No leasing fee evidence rows detected for selected filters.</td></tr>';

      estimateBody.innerHTML = filteredEstimates.length ? filteredEstimates.map(function(row) {
        return '<tr>' +
          '<td>' + escapeHtml(String(row.estimate_id || '—')) + '</td>' +
          '<td>' + escapeHtml(String(row.work_order_number || '—')) + '</td>' +
          '<td>' + escapeHtml(String(row.source || '—')) + '</td>' +
          '<td>' + escapeHtml(String(row.status || '—')) + '</td>' +
          '<td>' + escapeHtml(String(resolveGroupNameFromUuid(row.property_group_id) || row.property_group_id || '—')) + '</td>' +
          '<td>' + escapeHtml(String(row.vendor_name || '—')) + '</td>' +
          '<td>' + escapeHtml(formatDate(row.updated_at)) + '</td>' +
        '</tr>';
      }).join('') : '<tr><td colspan="7" style="text-align:center;padding:20px;color:var(--text-muted);">No estimate approval pipeline rows for configured material sources.</td></tr>';
    }

    var scopeText = scope.groupUuid ? (resolveGroupNameFromUuid(scope.groupUuid) || scope.groupUuid) : 'All Properties';
    var ticklerSummary = $('#mgrTicklerSummary');
    if (ticklerSummary) ticklerSummary.textContent = filteredTickler.length + ' rows • ' + scopeText;
    var renewalSummary = $('#mgrRenewalSummary');
    if (renewalSummary) renewalSummary.textContent = filteredRenewals.length + ' rows • ' + scopeText;
    var ledgerSummary = $('#mgrLedgerSummary');
    if (ledgerSummary) ledgerSummary.textContent = filteredLedger.length + ' rows • renewal/move-in charges flagged';
    var estimateSummary = $('#mgrEstimateSummary');
    if (estimateSummary) estimateSummary.textContent = filteredEstimates.length + ' rows • configured source list';

    var kpiTickler = $('#kpiMgrTickler');
    if (kpiTickler) kpiTickler.textContent = String(filteredTickler.length || 0);
    var kpiRenewals = $('#kpiMgrRenewals');
    if (kpiRenewals) kpiRenewals.textContent = String(filteredRenewals.length || 0);
    var kpiLeasing = $('#kpiMgrLeasingFees');
    if (kpiLeasing) kpiLeasing.textContent = String(filteredLedger.length || 0);
    var kpiEstimate = $('#kpiMgrEstimatePipeline');
    if (kpiEstimate) kpiEstimate.textContent = String(filteredEstimates.length || 0);

    var kpiTicklerSub = $('#kpiMgrTicklerSub');
    if (kpiTicklerSub) kpiTicklerSub.textContent = scope.fromDate + ' → ' + scope.toDate;
    var kpiRenewalsSub = $('#kpiMgrRenewalsSub');
    if (kpiRenewalsSub) kpiRenewalsSub.textContent = scope.fromMonth + ' → ' + scope.toMonth;
    var kpiLeasingSub = $('#kpiMgrLeasingFeesSub');
    if (kpiLeasingSub) kpiLeasingSub.textContent = 'renewal/move-in evidence';
    var kpiEstimateSub = $('#kpiMgrEstimatePipelineSub');
    if (kpiEstimateSub) kpiEstimateSub.textContent = sourceMatch.join(', ');

    renderManagerReviewCharts();
  }

  function clearLoadedKey() {
    loadedKey = '';
  }

  return {
    renderManagerReviewSection: renderManagerReviewSection,
    clearLoadedKey: clearLoadedKey,
  };
}
