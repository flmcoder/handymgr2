export type BillingVendorBucket = 'in_house' | 'third_party';

// Task spec: any payee_name containing "Fort Lowell" is In-House (not an exact-name match).
export function classifyBillingVendorBucket(vendorName: unknown): BillingVendorBucket {
  const normalized = String(vendorName ?? '').trim().toLowerCase();
  return normalized.includes('fort lowell') ? 'in_house' : 'third_party';
}

export type BillingSpendTimeframe = {
  days: number;
  sinceIso: string;
};

export function parseBillingSpendTimeframe(
  query: Record<string, unknown>,
  now: Date = new Date(),
  fallbackDays = 180,
  maxDays = 3650,
): BillingSpendTimeframe {
  const explicit = String(query?.updated_from ?? query?.date_from ?? '').trim();
  if (explicit) {
    const parsed = new Date(explicit);
    if (!Number.isNaN(parsed.getTime())) {
      const days = Math.max(1, Math.round((now.getTime() - parsed.getTime()) / 86_400_000));
      return { days, sinceIso: parsed.toISOString() };
    }
  }

  const parsedDays = Number.parseInt(String(query?.days ?? ''), 10);
  const days = Number.isFinite(parsedDays) && parsedDays > 0
    ? Math.min(maxDays, parsedDays)
    : fallbackDays;
  return { days, sinceIso: new Date(now.getTime() - (days * 86_400_000)).toISOString() };
}
