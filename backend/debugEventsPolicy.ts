export type DebugEventLevel = 'info' | 'success' | 'warning' | 'error';

const LEVELS = new Set<DebugEventLevel>(['info', 'success', 'warning', 'error']);

export type NormalizedDebugEvent = {
  level: DebugEventLevel;
  event_type: string;
  source: string;
  action: string;
  message: string;
  details: Record<string, unknown>;
  client_ts: string;
};

function clip(value: unknown, max: number): string {
  return String(value ?? '').trim().slice(0, max);
}

function sanitizeDetails(value: unknown, depth = 0): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return {};
  if (depth > 3) return { value: '[depth limit]' };
  const blocked = /authorization|access_token|auth_token|device_token|password|secret|client_secret|api_key|cookie/i;
  const result: Record<string, unknown> = {};
  for (const [key, raw] of Object.entries(value as Record<string, unknown>)) {
    if (blocked.test(key)) {
      result[key] = '[redacted]';
    } else if (raw && typeof raw === 'object' && !Array.isArray(raw)) {
      result[key] = sanitizeDetails(raw, depth + 1);
    } else if (Array.isArray(raw)) {
      result[key] = raw.slice(0, 50).map((item) => (
        item && typeof item === 'object' ? sanitizeDetails(item, depth + 1) : String(item ?? '').slice(0, 500)
      ));
    } else {
      result[key] = typeof raw === 'string' ? raw.slice(0, 2000) : raw;
    }
  }
  return result;
}

export function normalizeDebugEvent(input: unknown): NormalizedDebugEvent | null {
  const row = input && typeof input === 'object' ? input as Record<string, unknown> : {};
  const levelRaw = clip(row.level || 'info', 20).toLowerCase() as DebugEventLevel;
  const level: DebugEventLevel = LEVELS.has(levelRaw) ? levelRaw : 'info';
  const message = clip(row.message ?? '', 4000);
  if (!message) return null;
  const clientTsRaw = clip(row.client_ts || '', 80);
  const clientTs = clientTsRaw && !Number.isNaN(new Date(clientTsRaw).getTime())
    ? new Date(clientTsRaw).toISOString()
    : new Date().toISOString();
  return {
    level,
    event_type: clip(row.event_type || 'runtime', 80) || 'runtime',
    source: clip(row.source || 'frontend', 120) || 'frontend',
    action: clip(row.action || '', 120),
    message,
    details: sanitizeDetails(row.details),
    client_ts: clientTs,
  };
}

export function resolveDebugRetentionDays(value: unknown, fallback = 30): number {
  const parsed = Number.parseInt(String(value ?? ''), 10);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.max(1, Math.min(365, parsed));
}
