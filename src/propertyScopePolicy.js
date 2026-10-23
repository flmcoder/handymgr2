export function isRecordInPropertyGroupScope(record, scopeValues) {
  const scopeIds = Array.from(scopeValues || [])
    .map((value) => String(value || '').trim().toLowerCase())
    .filter(Boolean);
  if (scopeIds.length === 0) return true;
  if (!record || typeof record !== 'object') return false;

  const recordGroupUuid = String(record.property_group_uuid || record.propertyGroupUuid || '')
    .trim()
    .toLowerCase();
  return !!recordGroupUuid && scopeIds.includes(recordGroupUuid);
}