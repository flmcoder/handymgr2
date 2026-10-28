export const ROOT_API_ACTIONS = Object.freeze([
  'verify_role',
  'device_setup',
  'device_otp_request',
  'device_otp_verify',
  'session_info',
  'generate_magic_link',
  'handleGenerateMagicLink',
  'send_magic_link_test_sms',
  'portal_validate',
  'portal_status',
  'portal_schedule',
  'portal_reschedule',
  'portal_note',
  'portal_reassign_request',
  'passthrough',
  'report',
  'v2_report',
  'recent_tasks',
  'bills',
  'bills_list',
  'admin_rate_limits',
  'users',
  'tenant_comms_log',
  'tech_roster',
  'dispatch_sync_assignees',
  'dispatch_seed_reassignment_test',
  'seed_reassignment_queue_test',
  'reassignment_queue_seed',
  'reassignment_queue',
  'queue',
  'start_monitoring_work_order',
  'stop_monitoring_work_order',
  'noon_warning_cron',
  'midnight_reassign_cron',
  'manager_review',
  'units',
  'unit_lookup',
  'turns',
  'unit_turns',
  'turns_incremental',
  'unit_turns_history',
  'unit_turns_sync',
  'estimates',
]);

export const ROOT_ACTIONS_WITH_INLINE_HANDLERS = Object.freeze([
  'session_info',
  'generate_magic_link',
  'handleGenerateMagicLink',
  'send_magic_link_test_sms',
  'portal_validate',
  'portal_status',
  'portal_schedule',
  'portal_reschedule',
  'portal_note',
  'portal_reassign_request',
  'passthrough',
  'report',
  'v2_report',
  'recent_tasks',
  'bills',
  'bills_list',
  'admin_rate_limits',
  'users',
  'tenant_comms_log',
  'tech_roster',
  'dispatch_sync_assignees',
  'dispatch_seed_reassignment_test',
  'seed_reassignment_queue_test',
  'reassignment_queue_seed',
  'reassignment_queue',
  'queue',
  'start_monitoring_work_order',
  'stop_monitoring_work_order',
  'noon_warning_cron',
  'midnight_reassign_cron',
  'manager_review',
]);

/** Resolve a query-first action the same way the legacy backend dispatcher does. */
export function getRootPostAction(req) {
  const queryAction = String(req.query?.action ?? '').trim();
  if (queryAction) return queryAction;
  const bodyAction = req.body && typeof req.body === 'object' ? req.body.action : '';
  return String(bodyAction ?? '').trim();
}

/** Return whether a root POST action is on the legacy compatibility allowlist. */
export function isAllowedRootPostAction(action) {
  return ROOT_API_ACTIONS.includes(String(action ?? '').trim());
}

/** Return true when the action is handled by an explicit legacy dispatcher branch. */
export function isRootActionHandledInline(action) {
  return ROOT_ACTIONS_WITH_INLINE_HANDLERS.includes(String(action ?? '').trim());
}

export function getRootActionImplementationPolicy(action, routeActions = []) {
  const normalizedAction = String(action ?? '').trim();
  if (!isAllowedRootPostAction(normalizedAction)) return 'rejected';
  if (isRootActionHandledInline(normalizedAction)) return 'inline';
  if (routeActions.includes(normalizedAction)) return 'route';
  return 'unhandled';
}

/** Return whether the root action has an inline branch or a route-map handler. */
export function isRootPostAllowedActionHandled(action, routeActions = []) {
  const policy = getRootActionImplementationPolicy(action, routeActions);
  return policy === 'inline' || policy === 'route';
}

/** Reject POSTs to `/` unless they target an action in the legacy compatibility allowlist. */
export function rejectUnrecognizedRootPost(req, res, next) {
  const rootAction = getRootPostAction(req);
  if (req.method === 'POST' && req.path === '/' && !isAllowedRootPostAction(rootAction)) {
    res.status(405).end();
    return;
  }

  next();
}

export default rejectUnrecognizedRootPost;
