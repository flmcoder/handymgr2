export function shouldRefreshDispatchSnapshot(params: Record<string, unknown>): boolean {
  return !/^(1|true|yes|on)$/i.test(String(params.read_only || params.readOnly || '').trim());
}