export type BillScopeDecision = {
  allowed: boolean;
  propertyGroupId: string;
};

export function resolveBillScope(
  role: unknown,
  sessionGroupId: unknown,
  requestedGroupId: unknown,
): BillScopeDecision {
  const normalizedRole = String(role || '').trim().toLowerCase();
  const normalizedSessionGroupId = String(sessionGroupId || '').trim();
  const normalizedRequestedGroupId = String(requestedGroupId || '').trim();

  if (normalizedRole === 'pm_readonly') {
    return {
      allowed: Boolean(normalizedSessionGroupId),
      propertyGroupId: normalizedSessionGroupId,
    };
  }

  return { allowed: true, propertyGroupId: normalizedRequestedGroupId };
}

function collectBillPropertyIds(bill: unknown): string[] {
  if (!bill || typeof bill !== 'object') return [];

  const row = bill as Record<string, any>;
  const raw = row.raw && typeof row.raw === 'object' ? row.raw : row;
  const nestedProperty = raw.property && typeof raw.property === 'object' ? raw.property : {};
  const directIds = [
    row.property_id,
    row.PropertyId,
    raw.property_id,
    raw.PropertyId,
    nestedProperty.id,
    nestedProperty.Id,
    nestedProperty.property_id,
    nestedProperty.PropertyId,
  ];
  const lineItems = [row.line_items, row.LineItems, raw.line_items, raw.LineItems]
    .find(Array.isArray) || [];

  const ids = directIds.concat(lineItems.flatMap((item: any) => [
    item?.property_id,
    item?.PropertyId,
    item?.property?.id,
    item?.property?.Id,
  ]));

  return Array.from(new Set(ids.map((value) => String(value || '').trim()).filter(Boolean)));
}

export function filterBillsForPropertyScope<T>(bills: T[], allowedPropertyIds: ReadonlySet<string>): T[] {
  if (!Array.isArray(bills) || allowedPropertyIds.size === 0) return [];

  return bills.filter((bill) => {
    const propertyIds = collectBillPropertyIds(bill);
    return propertyIds.length > 0 && propertyIds.every((propertyId) => allowedPropertyIds.has(propertyId));
  });
}