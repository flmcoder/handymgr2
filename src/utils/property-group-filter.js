export function normalizeGroupSelectionValue(value) {
  var raw = String(value || '').trim();
  if (!raw) return '';
  var lower = raw.toLowerCase();
  if (lower === 'all properties') return '';
  if (lower.indexOf('all properties') !== -1 && (raw.charAt(0) === '*' || lower.indexOf('appfolio') !== -1)) return '';
  return raw;
}

export function resolvePropertyGroupSelection(config) {
  var currentGroup = normalizeGroupSelectionValue(config && config.currentGroup);
  var forcedUuid = String((config && config.forcedGroupUuid) || '').trim();
  var forcedName = String((config && config.forcedGroupName) || '').trim();
  var accessRole = String((config && config.accessRole) || '').trim();
  var resolveNameFromUuid = config && typeof config.resolveNameFromUuid === 'function'
    ? config.resolveNameFromUuid
    : function() { return ''; };
  var resolveUuidFromName = config && typeof config.resolveUuidFromName === 'function'
    ? config.resolveUuidFromName
    : function() { return ''; };

  if (accessRole === 'pm_readonly' && forcedUuid) {
    var scopedName = normalizeGroupSelectionValue(resolveNameFromUuid(forcedUuid) || forcedName || currentGroup);
    return {
      groupName: scopedName,
      groupUuid: forcedUuid,
      isScoped: true,
    };
  }
  return {
    groupName: currentGroup,
    groupUuid: normalizeGroupSelectionValue(currentGroup) ? String(resolveUuidFromName(currentGroup) || '').trim() : '',
    isScoped: false,
  };
}

export function readGroupCandidates(record) {
  var row = record && typeof record === 'object' ? record : {};
  var nestedProperty = row.property && typeof row.property === 'object' ? row.property : {};
  var names = [
    row.property_group,
    row.propertyGroup,
    row._propertyGroup,
    row.group_name,
    row.groupName,
    row.portfolio,
    nestedProperty.property_group,
    nestedProperty.group_name,
    nestedProperty.groupName,
  ].map(function(v) { return String(v || '').trim(); }).filter(Boolean);
  var ids = [
    row.property_group_id,
    row.property_group_uuid,
    row.propertyGroupId,
    row.PropertyGroupId,
    row.PropertyGroupUuid,
    row.group_id,
    row.group_uuid,
    nestedProperty.property_group_id,
    nestedProperty.property_group_uuid,
    nestedProperty.PropertyGroupId,
    nestedProperty.PropertyGroupUuid,
  ].map(function(v) { return String(v || '').trim(); }).filter(Boolean);
  return { names: names, ids: ids };
}

export function matchesPropertyGroupSelection(record, selection) {
  var normalizedName = normalizeGroupSelectionValue(selection && selection.groupName);
  var normalizedUuid = String((selection && selection.groupUuid) || '').trim();
  if (!normalizedName && !normalizedUuid) return true;
  var candidates = readGroupCandidates(record);
  if (normalizedUuid && candidates.ids.indexOf(normalizedUuid) !== -1) return true;
  if (!normalizedName) return false;
  var wanted = normalizedName.toLowerCase();
  for (var i = 0; i < candidates.names.length; i++) {
    if (String(candidates.names[i]).toLowerCase() === wanted) return true;
  }
  return false;
}
