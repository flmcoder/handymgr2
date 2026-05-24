export function phoneDigitsOnly(value) {
  var digits = String(value || '').replace(/\D+/g, '');
  if (digits.length === 11 && digits.charAt(0) === '1') return digits.slice(1);
  return digits;
}

export function formatUsPhoneDashed(value) {
  var digits = phoneDigitsOnly(value).slice(0, 10);
  if (!digits) return '';
  if (digits.length <= 3) return digits;
  if (digits.length <= 6) return digits.slice(0, 3) + '-' + digits.slice(3);
  return digits.slice(0, 3) + '-' + digits.slice(3, 6) + '-' + digits.slice(6);
}

export function normalizeWorkPhone(value) {
  var digits = phoneDigitsOnly(value);
  if (digits.length === 10) return {
    e164: '+1' + digits,
    digits: digits,
    display: formatUsPhoneDashed(digits),
  };
  if (String(value || '').trim().charAt(0) === '+') {
    var allDigits = String(value || '').replace(/\D+/g, '');
    if (allDigits.length >= 10 && allDigits.length <= 15) {
      return {
        e164: '+' + allDigits,
        digits: allDigits,
        display: '+' + allDigits,
      };
    }
  }
  return { e164: '', digits: '', display: '' };
}

export function normalizeOtpEmail(value) {
  var email = String(value || '').trim().toLowerCase();
  if (!email) return '';
  return /^[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}$/i.test(email) ? email : '';
}

export function normalizeOtpIdentifier(value) {
  var email = normalizeOtpEmail(value);
  if (email) return email;
  return normalizeWorkPhone(value).e164;
}

export function formatOtpIdentifierSummary(value) {
  var email = normalizeOtpEmail(value);
  if (email) return email;
  var normalized = normalizeWorkPhone(value);
  if (!normalized.e164) return String(value || '').trim();
  return normalized.display || normalized.e164;
}
