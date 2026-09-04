export function isClientAbortError(error: unknown): boolean {
  if (!error || typeof error !== 'object') return false;

  const record = error as { code?: unknown; message?: unknown };
  const code = String(record.code || '').toUpperCase();
  const message = String(record.message || '').toLowerCase();

  return code === 'ECONNABORTED'
    || code === 'ECONNRESET'
    || message === 'request aborted'
    || message === 'socket hang up';
}