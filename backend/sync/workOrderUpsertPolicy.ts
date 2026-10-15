const IDENTIFIER_FIELDS = [
  'id', 'workOrderUuid', 'propertyId', 'unitId', 'propertyGroupId',
  'assignedUserId', 'vendorId', 'occupancyId', 'requestingTenantId',
] as const;

function record(value: unknown): Record<string, unknown> | null {
  return value !== null && typeof value === 'object' ? value as Record<string, unknown> : null;
}

export function postgresErrorDetails(error: unknown): Record<string, unknown> {
  const details: Record<string, unknown> = {};
  const seen = new Set<unknown>();
  let current = record(error);
  while (current && !seen.has(current)) {
    seen.add(current);
    for (const key of ['code', 'detail', 'constraint', 'schema', 'table', 'column', 'dataType', 'routine']) {
      if (current[key] !== undefined) details[key] = current[key];
    }
    current = record(current.cause);
  }
  return details;
}

function diagnosticPayload(value: unknown): unknown {
  if (value instanceof Date) return value.toISOString();
  if (Array.isArray(value)) return value.map(diagnosticPayload);
  const object = record(value);
  if (!object) return value;
  return Object.fromEntries(Object.entries(object).map(([key, item]) => [
    key,
    /authorization|token|password|secret|api[_-]?key|cookie|vendorportallink/i.test(key)
      ? '[redacted]'
      : diagnosticPayload(item),
  ]));
}

export async function writeWorkOrderWithDiagnostics<T extends Record<string, unknown>>(
  payload: T,
  write: (values: T) => PromiseLike<unknown>,
): Promise<void> {
  const values = { ...payload };
  for (const field of IDENTIFIER_FIELDS) {
    if (typeof values[field] === 'string' && !values[field].trim()) {
      Object.assign(values, { [field]: null });
    }
  }
  try {
    await write(values);
  } catch (error) {
    // Drizzle wraps the driver's SQLSTATE and constraint metadata in cause.
    console.error('[work_orders:upsert] Database write failed', {
      ...postgresErrorDetails(error),
      payload: diagnosticPayload(values),
    });
    throw error;
  }
}
