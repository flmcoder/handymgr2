export const LEGACY_DASHBOARD_CHART_IDS = [
  'dashOccupancyDonut',
  'dashMoveOutsBar',
  'dashPortfolioTreemap',
  'dashLeasingVelocity',
  'dashMainChart',
  'dashAgingChart',
] as const;

export function shouldInitializeLegacyDashboard(mountIds: readonly string[]): boolean {
  const availableIds = new Set(mountIds.filter((id): id is string => typeof id === 'string' && id.length > 0));
  return LEGACY_DASHBOARD_CHART_IDS.some((id) => availableIds.has(id));
}