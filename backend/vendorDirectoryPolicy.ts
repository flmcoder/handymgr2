export type VendorDirectorySort = 'name' | 'trade' | 'recent';

export type VendorDirectoryQuery = {
  limit: number;
  offset: number;
  page: number;
  search: string;
  tradeCategory: string;
  propertyId: string;
  sort: VendorDirectorySort;
};

export type VendorSpendClassification =
  | 'in_house_maintenance'
  | 'in_house_realty'
  | 'third_party';

function parsePositiveInteger(value: unknown, fallback: number): number {
  const parsed = Number.parseInt(String(value ?? ''), 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

function cleanFilter(value: unknown, maxLength = 180): string {
  return String(value ?? '').trim().slice(0, maxLength);
}

export function parseVendorDirectoryQuery(query: Record<string, unknown>): VendorDirectoryQuery {
  const limit = Math.min(100, parsePositiveInteger(query.limit, 50));
  const requestedPage = parsePositiveInteger(query.page, 1);
  const requestedOffset = Number.parseInt(String(query.offset ?? ''), 10);
  const offset = Number.isFinite(requestedOffset) && requestedOffset >= 0
    ? Math.min(250_000, requestedOffset)
    : Math.min(250_000, (requestedPage - 1) * limit);
  const requestedSort = cleanFilter(query.sort, 20);
  const sort: VendorDirectorySort = requestedSort === 'trade' || requestedSort === 'recent'
    ? requestedSort
    : 'name';

  return {
    limit,
    offset,
    page: Math.floor(offset / limit) + 1,
    search: cleanFilter(query.search),
    tradeCategory: cleanFilter(query.trade_category),
    propertyId: cleanFilter(query.property_id),
    sort,
  };
}

export function classifyVendorSpend(vendorName: unknown): VendorSpendClassification {
  const normalizedName = String(vendorName ?? '').trim().toLowerCase();
  if (normalizedName === 'fort lowell maintenance') return 'in_house_maintenance';
  if (normalizedName === 'fort lowell realty') return 'in_house_realty';
  return 'third_party';
}

export type VendorComplianceResult = {
  compliant: boolean;
  missing: string[];
  expired: string[];
};

export function evaluateVendorCompliance(
  vendor: { liability_ins_expires?: unknown; workers_comp_expires?: unknown },
  today: Date = new Date(),
): VendorComplianceResult {
  const missing: string[] = [];
  const expired: string[] = [];
  const checks: Array<[string, unknown]> = [
    ['liability_ins_expires', vendor?.liability_ins_expires],
    ['workers_comp_expires', vendor?.workers_comp_expires],
  ];

  for (const [field, rawValue] of checks) {
    const value = String(rawValue ?? '').trim();
    if (!value) {
      missing.push(field);
      continue;
    }
    const parsed = new Date(value);
    if (Number.isNaN(parsed.getTime())) {
      missing.push(field);
      continue;
    }
    if (parsed < today) expired.push(field);
  }

  return { compliant: missing.length === 0 && expired.length === 0, missing, expired };
}