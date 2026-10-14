export type TurnDetailRow = Record<string, unknown>;

function text(row: TurnDetailRow, ...keys: string[]): string {
  for (const key of keys) {
    const value = row[key];
    if (value !== null && value !== undefined && String(value).trim()) return String(value).trim();
  }
  return '';
}

function dateText(row: TurnDetailRow, ...keys: string[]): string {
  return text(row, ...keys);
}

export function normalizeUnitTurnDetail(row: TurnDetailRow) {
  const propertyId = text(row, 'property_id', 'PropertyId', 'property_uuid', 'PropertyUUID');
  const unitId = text(row, 'unit_id', 'UnitId', 'unit_uuid', 'UnitUUID');
  const moveOutDate = dateText(row, 'move_out_date', 'MoveOutDate', 'move_out', 'MoveOut');
  const turnEndDate = dateText(row, 'turn_end_date', 'TurnEndDate', 'turn_end', 'TurnEnd');
  const explicitTurnId = text(row, 'unit_turn_id', 'UnitTurnId', 'unit_turn_uuid', 'UnitTurnUUID');
  return {
    turnId: explicitTurnId || [propertyId, unitId, moveOutDate.slice(0, 10), turnEndDate.slice(0, 10)].join('|'),
    propertyId,
    propertyName: text(row, 'property_name', 'PropertyName', 'property', 'Property', 'property_name_address'),
    unitId,
    unitName: text(row, 'unit_name', 'UnitName', 'unit', 'Unit', 'unit_number', 'UnitNumber'),
    moveOutDate,
    turnEndDate,
    expectedMoveInDate: dateText(row, 'expected_move_in_date', 'ExpectedMoveInDate', 'expected_move_in', 'ExpectedMoveIn'),
    status: text(row, 'unit_turn_status', 'UnitTurnStatus', 'status', 'Status'),
    rawJson: row,
  };
}

export type TurnIdentity = {
  propertyId?: unknown;
  unitId?: unknown;
  moveOutDate?: unknown;
  unitTurnId?: unknown;
};

export function canonicalTurnKey(turn: TurnIdentity): string {
  const nativeId = String(turn.unitTurnId ?? '').trim();
  if (nativeId) return `native:${nativeId}`;
  const propertyId = String(turn.propertyId ?? '').trim();
  const unitId = String(turn.unitId ?? '').trim();
  const moveOutDate = String(turn.moveOutDate ?? '').slice(0, 10);
  if (!unitId || !moveOutDate) return '';
  return `unit:${propertyId}:${unitId}:${moveOutDate}`;
}

/** A source-independent overlap key used to merge native turns and notice-derived turns. */
export function turnOverlapKey(turn: TurnIdentity): string {
  const propertyId = String(turn.propertyId ?? '').trim();
  const unitId = String(turn.unitId ?? '').trim();
  const moveOutDate = String(turn.moveOutDate ?? '').slice(0, 10);
  if (!unitId || !moveOutDate) return '';
  return `${propertyId}|${unitId}|${moveOutDate}`;
}

export type TurnSourceRow = TurnIdentity & {
  source?: string;
  propertyName?: string;
  unitName?: string;
  workOrders?: Array<Record<string, unknown>>;
};

export type MergedTurnSource = {
  overlapKey: string;
  propertyId: string;
  unitId: string;
  moveOutDate: string;
  propertyName: string;
  unitName: string;
  sources: string[];
  workOrders: Array<Record<string, unknown>>;
};

/** Merge overlapping report/move-out sources by property, unit and date-only move-out. */
export function mergeTurnSources(rows: TurnSourceRow[]): MergedTurnSource[] {
  const merged = new Map<string, MergedTurnSource>();
  for (const row of rows) {
    const overlapKey = turnOverlapKey(row);
    if (!overlapKey) continue;
    let entry = merged.get(overlapKey);
    if (!entry) {
      entry = {
        overlapKey,
        propertyId: String(row.propertyId ?? '').trim(),
        unitId: String(row.unitId ?? '').trim(),
        moveOutDate: String(row.moveOutDate ?? '').slice(0, 10),
        propertyName: '',
        unitName: '',
        sources: [],
        workOrders: [],
      };
      merged.set(overlapKey, entry);
    }
    if (row.source && !entry.sources.includes(row.source)) entry.sources.push(row.source);
    if (!entry.propertyName && row.propertyName) entry.propertyName = String(row.propertyName).trim();
    if (!entry.unitName && row.unitName) entry.unitName = String(row.unitName).trim();
    for (const workOrder of row.workOrders || []) {
      const id = String(workOrder.woNumber || workOrder.wo_number || workOrder.id || workOrder.workOrderId || '').trim();
      if (!id) continue;
      if (!entry.workOrders.some((existing) => String(existing.woNumber || existing.wo_number || existing.id || existing.workOrderId || '').trim() === id)) {
        entry.workOrders.push(workOrder);
      }
    }
  }
  return Array.from(merged.values());
}