export function createAuthModule(deps) {
  var $ = deps.$;
  var fetchWithTimeout = deps.fetchWithTimeout;
  var getProxyUrl = deps.getProxyUrl;
  var getNavigatorPlatform = deps.getNavigatorPlatform;

  function normalizeOtpEmail(raw) {
    var email = String(raw || '').trim().toLowerCase();
    if (!email) return '';
    var m = email.match(/^[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}$/i);
    return m ? email : '';
  }

  function extractUsPhoneDigits(raw) {
    var digits = String(raw || '').replace(/\D+/g, '');
    if (!digits) return '';
    if (digits.length === 11 && digits.charAt(0) === '1') digits = digits.slice(1);
    else if (digits.length > 10) digits = digits.slice(-10);
    return digits;
  }

  function formatUsPhoneDisplay(raw) {
    var digits = extractUsPhoneDigits(raw);
    if (!digits) return '';
    if (digits.length <= 3) return '(' + digits;
    if (digits.length <= 6) return '(' + digits.slice(0, 3) + ') ' + digits.slice(3);
    return '(' + digits.slice(0, 3) + ') ' + digits.slice(3, 6) + '-' + digits.slice(6, 10);
  }

  function shouldTreatOtpInputAsEmail(raw) {
    var val = String(raw || '').trim();
    return val.indexOf('@') !== -1 || /[a-z]/i.test(val);
  }

  function normalizeOtpPhone(raw) {
    var digits = extractUsPhoneDigits(raw);
    if (digits.length === 10) return '+1' + digits;
    return '';
  }

  function normalizeOtpIdentifier(raw) {
    return normalizeOtpEmail(raw) || normalizeOtpPhone(raw);
  }

  function normalizeOtpScopeUuid(raw) {
    var value = String(raw || '').trim().toLowerCase();
    if (!value) return '';
    return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value) ? value : '';
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
    if (summaryValue && sent) {
      summaryValue.textContent = formatOtpIdentifierSummary(identifier || (identifierInput ? identifierInput.value : ''));
    }
    if (verifyBtn) verifyBtn.textContent = 'Verify OTP';

    if (!sent) {
      if (typeof deps.stopOtpCountdown === 'function') deps.stopOtpCountdown();
      if (codeInput) codeInput.value = '';
    } else if (codeInput) {
      setTimeout(function() { codeInput.focus(); codeInput.select(); }, 20);
    }
  }

  async function setupTrustedDevice(setupPin, userName) {
    var proxyUrl = getProxyUrl();
    if (!proxyUrl) throw new Error('No proxy configured');
    var sep = proxyUrl.indexOf('?') !== -1 ? '&' : '?';
    var url = proxyUrl + sep + 'action=device_setup';
    var payload = {
      pin: setupPin,
      user_name: userName || ('hm-' + getNavigatorPlatform())
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

  async function requestDeviceOtp(identifier, userName) {
    var proxyUrl = getProxyUrl();
    if (!proxyUrl) throw new Error('No proxy configured');
    var sep = proxyUrl.indexOf('?') !== -1 ? '&' : '?';
    var url = proxyUrl + sep + 'action=device_otp_request';
    var email = normalizeOtpEmail(identifier);
    var phone = normalizeOtpPhone(identifier);
    var payload = {
      identifier: String(identifier || '').trim(),
      email: email || undefined,
      phone: phone || undefined,
      user_name: userName || ('hm-' + getNavigatorPlatform())
    };
    var res = await fetchWithTimeout(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
      body: JSON.stringify(payload)
    }, 30000);
    var data = {};
    try { data = await res.json(); } catch (e) { /* */ }
    if (!res.ok || !data.ok) {
      var reqErr = new Error(data.error || data.message || ('OTP request failed (HTTP ' + res.status + ')'));
      reqErr.details = data;
      throw reqErr;
    }
    return data;
  }

  async function verifyDeviceOtp(identifier, code, userName) {
    var proxyUrl = getProxyUrl();
    if (!proxyUrl) throw new Error('No proxy configured');
    var sep = proxyUrl.indexOf('?') !== -1 ? '&' : '?';
    var url = proxyUrl + sep + 'action=device_otp_verify';
    var email = normalizeOtpEmail(identifier);
    var phone = normalizeOtpPhone(identifier);
    var payload = {
      identifier: String(identifier || '').trim(),
      email: email || undefined,
      phone: phone || undefined,
      code: code,
      user_name: userName || ('hm-' + getNavigatorPlatform())
    };
    var res = await fetchWithTimeout(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
      body: JSON.stringify(payload)
    }, 30000);
    var data = {};
    try { data = await res.json(); } catch (e) { /* */ }
    if (!res.ok || !data.ok || !data.token) {
      var verifyErr = new Error(data.error || data.message || ('OTP verify failed (HTTP ' + res.status + ')'));
      verifyErr.details = data;
      throw verifyErr;
    }
    return data;
  }

  return {
    normalizeOtpEmail: normalizeOtpEmail,
    extractUsPhoneDigits: extractUsPhoneDigits,
    normalizeOtpPhone: normalizeOtpPhone,
    formatOtpIdentifierSummary: formatOtpIdentifierSummary,
    formatUsPhoneDisplay: formatUsPhoneDisplay,
    shouldTreatOtpInputAsEmail: shouldTreatOtpInputAsEmail,
    normalizeOtpIdentifier: normalizeOtpIdentifier,
    normalizeOtpScopeUuid: normalizeOtpScopeUuid,
    setPmOtpStep: setPmOtpStep,
    setupTrustedDevice: setupTrustedDevice,
    requestDeviceOtp: requestDeviceOtp,
    verifyDeviceOtp: verifyDeviceOtp,
  };
}
