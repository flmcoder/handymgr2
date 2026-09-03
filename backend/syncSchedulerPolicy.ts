export type RequiredSyncEndpointOptions = {
  billsEnabled: boolean;
  v2Enabled: boolean;
  requiredV2Endpoints: readonly string[];
};

export function buildRequestedSyncEndpoints(
  configuredEndpoints: readonly string[],
  options: RequiredSyncEndpointOptions,
): string[] {
  const configured = configuredEndpoints.filter((endpoint) => {
    if (!options.billsEnabled && endpoint === 'v0:bills') return false;
    if (!options.v2Enabled && endpoint.startsWith('v2:')) return false;
    return true;
  });
  const required = [
    ...(options.billsEnabled ? ['v0:bills'] : []),
    ...(options.v2Enabled ? options.requiredV2Endpoints : []),
  ];

  return Array.from(new Set([...configured, ...required]));
}