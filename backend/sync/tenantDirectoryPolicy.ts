export type NormalizedTenantDirectoryRow = {
  propertyId: string | null;
  propertyName: string | null;
  unitId: string | null;
  unitName: string | null;
  tenantName: string | null;
  occupancyId: string | null;
  moveIn: unknown;
  moveOut: unknown;
};

function firstText(row: Record<string, unknown>, keys: string[]): string | null {
  for (const key of keys) {
    const value = row[key];
    if (value === null || value === undefined) continue;
    const text = String(value).trim();
    if (text) return text;
  }
  return null;
}

/** Normalize documented AppFolio Tenant Directory aliases without inventing IDs. */
export function normalizeTenantDirectoryRow(row: Record<string, unknown>): NormalizedTenantDirectoryRow {
  return {
    propertyId: firstText(row, ['property_id', 'PropertyId']),
    propertyName: firstText(row, ['property_name', 'PropertyName', 'property', 'Property']),
    unitId: firstText(row, ['unit_id', 'UnitId']),
    unitName: firstText(row, ['unit_name', 'UnitName', 'unit', 'Unit']),
    tenantName: firstText(row, ['tenant', 'Tenant', 'tenant_name', 'TenantName']),
    occupancyId: firstText(row, ['occupancy_id', 'OccupancyId']),
    moveIn: row.move_in ?? row.MoveIn ?? row.move_in_date ?? row.MoveInDate,
    moveOut: row.move_out ?? row.MoveOut ?? row.move_out_date ?? row.MoveOutDate,
  };
}

/** Match a property by either its canonical database ID or AppFolio source ID. */
export function tenantDirectoryPropertyMatchSql(propertyAlias = 'p', sourceIdExpr = 't.property_id'): string {
  return `(
    ${propertyAlias}.id = ${sourceIdExpr}
    or coalesce(
      ${propertyAlias}.raw_json ->> 'PropertyId',
      ${propertyAlias}.raw_json ->> 'property_id',
      ${propertyAlias}.raw_json ->> 'Id',
      ${propertyAlias}.raw_json ->> 'id'
    ) = ${sourceIdExpr}
    or regexp_replace(coalesce(${propertyAlias}.raw_json ->> 'Link', ''), '^.*/properties/', '') = ${sourceIdExpr}
  )`;
}