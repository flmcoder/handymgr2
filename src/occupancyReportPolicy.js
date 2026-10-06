function firstValue(row, keys) {
  for (let i = 0; i < keys.length; i += 1) {
    const value = row?.[keys[i]];
    if (value !== null && value !== undefined && String(value).trim() !== '') return value;
  }
  return '';
}

export function tenantTransactionDisplayRow(row) {
  const property = firstValue(row, ['property_name', 'property_name_address', 'property']);
  const unit = firstValue(row, ['unit_name', 'unit']);
  return {
    tenant: firstValue(row, ['occupancy_name', 'tenant_name']),
    propertyUnit: property + (unit ? ` / ${unit}` : ''),
    charges: firstValue(row, ['rent_charges', 'total_charges']) || 0,
    payments: firstValue(row, ['cash_payments', 'total_payments']) || 0,
    balance: firstValue(row, ['ending_balance', 'balance']) || 0,
    lastActivity: firstValue(row, ['last_activity_date']),
  };
}