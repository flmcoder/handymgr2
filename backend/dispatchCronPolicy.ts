export function isDispatchAutomationPaused(config: Record<string, unknown> | null | undefined): boolean {
  if (!config || !Object.prototype.hasOwnProperty.call(config, 'dispatch_paused')) return true;
  const value = String(config.dispatch_paused ?? '').trim().toLowerCase();
  return ['1', 'true', 'yes', 'on', 'y'].includes(value);
}

export async function executeDispatchCron<T>(input: {
  readConfig: () => Promise<Record<string, unknown>>;
  run: () => Promise<T>;
  pausedResult: T;
}): Promise<T> {
  const config = await input.readConfig();
  if (isDispatchAutomationPaused(config)) return input.pausedResult;
  return input.run();
}

export function buildDispatchReassignmentKey(input: {
  woId: unknown;
  dispatchVersion: unknown;
}): string {
  return JSON.stringify([
    String(input.woId ?? '').trim(),
    Math.max(0, Number(input.dispatchVersion) || 0),
  ]);
}

export function resolveDispatchVersion(input: {
  previousVersion: unknown;
  previousWorkOrderStatus: unknown;
  previousAssigneeId: unknown;
  workOrderStatus: unknown;
  assigneeId: unknown;
}): number {
  const previousVersion = Math.max(1, Number(input.previousVersion) || 1);
  const previousStatus = String(input.previousWorkOrderStatus ?? '').trim().toLowerCase();
  const nextStatus = String(input.workOrderStatus ?? '').trim().toLowerCase();
  const previousAssigneeId = String(input.previousAssigneeId ?? '').trim();
  const nextAssigneeId = String(input.assigneeId ?? '').trim();
  const statusChanged = !!previousStatus && !!nextStatus && previousStatus !== nextStatus;
  const assigneeChanged = !!previousAssigneeId && !!nextAssigneeId && previousAssigneeId !== nextAssigneeId;
  return previousVersion + (statusChanged || assigneeChanged ? 1 : 0);
}
