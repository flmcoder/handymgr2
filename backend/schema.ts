import {
  pgTable,
  text,
  integer,
  real,
  timestamp,
  boolean,
  jsonb,
  index,
  unique,
} from 'drizzle-orm/pg-core';

// AppFolio v0: Properties
export const appfolioProperties = pgTable(
  'appfolio_properties',
  {
    id: text('id').primaryKey(),
    name: text('name').notNull(),
    propertyGroupId: text('property_group_id'),
    street: text('street'),
    city: text('city'),
    state: text('state'),
    zip: text('zip'),
    rawJson: jsonb('raw_json').notNull().default({}),
    lastUpdatedAt: timestamp('last_updated_at', { withTimezone: true }),
    cachedAt: timestamp('cached_at', { withTimezone: true }).defaultNow().notNull(),
  },
  (table) => ({
    groupIdx: index('appfolio_properties_group_idx').on(table.propertyGroupId),
    nameIdx: index('appfolio_properties_name_idx').on(table.name),
  }),
);

// AppFolio v0: Property Groups
export const appfolioPropertyGroups = pgTable(
  'appfolio_property_groups',
  {
    id: text('id').primaryKey(),
    uuid: text('uuid'),
    name: text('name').notNull(),
    type: text('type'),
    propertyIds: jsonb('property_ids').notNull().default([]),
    rawJson: jsonb('raw_json').notNull().default({}),
    lastUpdatedAt: timestamp('last_updated_at', { withTimezone: true }),
    cachedAt: timestamp('cached_at', { withTimezone: true }).defaultNow().notNull(),
  },
  (table) => ({
    uuidIdx: index('appfolio_property_groups_uuid_idx').on(table.uuid),
    nameIdx: index('appfolio_property_groups_name_idx').on(table.name),
    updatedIdx: index('appfolio_property_groups_updated_idx').on(table.lastUpdatedAt),
  }),
);

// AppFolio v0: Units (includes unit status tracking)
export const appfolioUnits = pgTable(
  'appfolio_units',
  {
    unitId: text('unit_id').primaryKey(),
    propertyId: text('property_id').notNull(),
    name: text('name'),
    unitNumber: text('unit_number'),
    status: text('status'),
    bedrooms: integer('bedrooms'),
    bathrooms: real('bathrooms'),
    squareFeet: integer('square_feet'),
    marketRent: real('market_rent'),
    rawJson: jsonb('raw_json').notNull().default({}),
    lastUpdatedAt: timestamp('last_updated_at', { withTimezone: true }),
    cachedAt: timestamp('cached_at', { withTimezone: true }).defaultNow().notNull(),
  },
  (table) => ({
    propertyIdx: index('appfolio_units_property_idx').on(table.propertyId),
    statusIdx: index('appfolio_units_status_idx').on(table.status),
  }),
);

// AppFolio v0: Work Orders
export const appfolioWorkOrders = pgTable(
  'appfolio_work_orders',
  {
    id: text('id').primaryKey(),
    workOrderUuid: text('work_order_uuid'),
    woNumber: text('wo_number'),
    propertyId: text('property_id'),
    unitId: text('unit_id'),
    propertyGroupId: text('property_group_id'),
    description: text('description'),
    category: text('category'),
    priority: text('priority'),
    status: text('status'),
    assignedUserId: text('assigned_user_id'),
    assignedUserName: text('assigned_user_name'),
    vendorId: text('vendor_id'),
    vendorName: text('vendor_name'),
    estimatedAmount: real('estimated_amount'),
    totalCost: real('total_cost'),
    createdAt: timestamp('created_at', { withTimezone: true }),
    updatedAt: timestamp('updated_at', { withTimezone: true }),
    rawJson: jsonb('raw_json').notNull().default({}),
  },
  (table) => ({
    uuidIdx: index('appfolio_work_orders_uuid_idx').on(table.workOrderUuid),
    woNumberIdx: index('appfolio_work_orders_number_idx').on(table.woNumber),
    statusIdx: index('appfolio_work_orders_status_idx').on(table.status),
    propertyIdx: index('appfolio_work_orders_property_idx').on(table.propertyId),
    unitIdx: index('appfolio_work_orders_unit_idx').on(table.unitId),
  }),
);

// AppFolio v0: Users (Maintenance Tech baseline roster)
export const appfolioUsers = pgTable(
  'appfolio_users',
  {
    techId: text('tech_id').primaryKey(),
    techName: text('tech_name').notNull(),
    email: text('email'),
    userRole: text('user_role').notNull().default(''),
    appfolioActive: boolean('appfolio_active').default(true),
    rawJson: jsonb('raw_json').notNull().default({}),
    lastUpdatedAt: timestamp('last_updated_at', { withTimezone: true }),
    cachedAt: timestamp('cached_at', { withTimezone: true }).defaultNow().notNull(),
  },
  (table) => ({
    roleIdx: index('appfolio_users_role_idx').on(table.userRole),
    activeIdx: index('appfolio_users_active_idx').on(table.appfolioActive),
    nameIdx: index('appfolio_users_name_idx').on(table.techName),
  }),
);

// AppFolio v0 derived: Estimates
export const appfolioEstimates = pgTable(
  'appfolio_estimates',
  {
    estimateId: text('estimate_id').primaryKey(),
    workOrderId: text('work_order_id'),
    workOrderNumber: text('work_order_number'),
    currentStatus: text('current_status').notNull(),
    propertyGroupId: text('property_group_id'),
    source: text('source'),
    statusHistory: jsonb('status_history').notNull().default([]),
    rawData: jsonb('raw_data').notNull().default({}),
    createdAt: timestamp('created_at', { withTimezone: true }).defaultNow().notNull(),
    updatedAt: timestamp('updated_at', { withTimezone: true }).defaultNow().notNull(),
  },
  (table) => ({
    statusIdx: index('appfolio_estimates_status_idx').on(table.currentStatus),
    groupIdx: index('appfolio_estimates_group_idx').on(table.propertyGroupId),
    woIdx: index('appfolio_estimates_wo_idx').on(table.workOrderId),
  }),
);

export const appfolioUnitInspections = pgTable(
  'appfolio_unit_inspections',
  {
    inspectionId: text('inspection_id').primaryKey(),
    propertyId: text('property_id'),
    propertyName: text('property_name'),
    unitId: text('unit_id'),
    unitName: text('unit_name'),
    lastInspectionDate: timestamp('last_inspection_date', { withTimezone: true }),
    tenantName: text('tenant_name'),
    tenantPrimaryPhoneNumber: text('tenant_primary_phone_number'),
    moveInDate: timestamp('move_in_date', { withTimezone: true }),
    moveOutDate: timestamp('move_out_date', { withTimezone: true }),
    rentable: text('rentable'),
    occupancyId: text('occupancy_id'),
    unitTags: text('unit_tags'),
    rawJson: jsonb('raw_json').notNull().default({}),
    lastUpdatedAt: timestamp('last_updated_at', { withTimezone: true }),
    cachedAt: timestamp('cached_at', { withTimezone: true }).defaultNow().notNull(),
  },
  (table) => ({
    propertyIdx: index('appfolio_unit_inspections_property_idx').on(table.propertyId),
    unitIdx: index('appfolio_unit_inspections_unit_idx').on(table.unitId),
    inspectionIdx: index('appfolio_unit_inspections_inspection_idx').on(table.lastInspectionDate),
  }),
);

export const appfolioTenantDirectory = pgTable(
  'appfolio_tenant_directory',
  {
    recordId: text('record_id').primaryKey(),
    propertyId: text('property_id'),
    propertyName: text('property_name'),
    unitId: text('unit_id'),
    unitName: text('unit_name'),
    tenantName: text('tenant_name'),
    status: text('status'),
    tenantType: text('tenant_type'),
    phoneNumbers: text('phone_numbers'),
    emails: text('emails'),
    moveInDate: timestamp('move_in_date', { withTimezone: true }),
    leaseTo: timestamp('lease_to', { withTimezone: true }),
    moveOutDate: timestamp('move_out_date', { withTimezone: true }),
    rent: text('rent'),
    tenantTags: text('tenant_tags'),
    tenantAgent: text('tenant_agent'),
    tenantVisibility: text('tenant_visibility'),
    occupancyId: text('occupancy_id'),
    unitTags: text('unit_tags'),
    rawJson: jsonb('raw_json').notNull().default({}),
    lastUpdatedAt: timestamp('last_updated_at', { withTimezone: true }),
    cachedAt: timestamp('cached_at', { withTimezone: true }).defaultNow().notNull(),
  },
  (table) => ({
    propertyIdx: index('appfolio_tenant_directory_property_idx').on(table.propertyId),
    unitIdx: index('appfolio_tenant_directory_unit_idx').on(table.unitId),
    moveOutIdx: index('appfolio_tenant_directory_move_out_idx').on(table.moveOutDate),
  }),
);

export const appfolioUnitTurnDetails = pgTable(
  'appfolio_unit_turn_details',
  {
    turnId: text('turn_id').primaryKey(),
    propertyId: text('property_id'),
    propertyName: text('property_name'),
    unitId: text('unit_id'),
    unitName: text('unit_name'),
    notes: text('notes'),
    referenceUser: text('reference_user'),
    moveOutDate: timestamp('move_out_date', { withTimezone: true }),
    turnEndDate: timestamp('turn_end_date', { withTimezone: true }),
    expectedMoveInDate: timestamp('expected_move_in_date', { withTimezone: true }),
    targetDaysToComplete: integer('target_days_to_complete'),
    totalDaysToComplete: integer('total_days_to_complete'),
    laborFromWorkOrders: text('labor_from_work_orders'),
    purchaseOrdersFromWorkOrders: text('purchase_orders_from_work_orders'),
    billablesFromWorkOrders: text('billables_from_work_orders'),
    inventoryFromWorkOrders: text('inventory_from_work_orders'),
    totalBilled: text('total_billed'),
    unitTurnStatus: text('unit_turn_status'),
    propertyVisibility: text('property_visibility'),
    rawJson: jsonb('raw_json').notNull().default({}),
    lastUpdatedAt: timestamp('last_updated_at', { withTimezone: true }),
    cachedAt: timestamp('cached_at', { withTimezone: true }).defaultNow().notNull(),
  },
  (table) => ({
    propertyIdx: index('appfolio_unit_turn_details_property_idx').on(table.propertyId),
    unitIdx: index('appfolio_unit_turn_details_unit_idx').on(table.unitId),
    moveOutIdx: index('appfolio_unit_turn_details_move_out_idx').on(table.moveOutDate),
  }),
);

export const appfolioUnitVacancies = pgTable(
  'appfolio_unit_vacancies',
  {
    recordId: text('record_id').primaryKey(),
    propertyId: text('property_id'),
    propertyName: text('property_name'),
    unitId: text('unit_id'),
    unitName: text('unit_name'),
    vacantFrom: timestamp('vacant_from', { withTimezone: true }),
    availableOn: timestamp('available_on', { withTimezone: true }),
    marketRent: text('market_rent'),
    bedrooms: text('bedrooms'),
    bathrooms: text('bathrooms'),
    daysVacant: integer('days_vacant'),
    status: text('status'),
    propertyVisibility: text('property_visibility'),
    rawJson: jsonb('raw_json').notNull().default({}),
    lastUpdatedAt: timestamp('last_updated_at', { withTimezone: true }),
    cachedAt: timestamp('cached_at', { withTimezone: true }).defaultNow().notNull(),
  },
  (table) => ({
    propertyIdx: index('appfolio_unit_vacancies_property_idx').on(table.propertyId),
    unitIdx: index('appfolio_unit_vacancies_unit_idx').on(table.unitId),
    vacantFromIdx: index('appfolio_unit_vacancies_vacant_from_idx').on(table.vacantFrom),
  }),
);

// UUID-based unit turn tracking
export const unitTurnTracker = pgTable(
  'unit_turn_tracker',
  {
    trackingUuid: text('tracking_uuid').primaryKey(),
    trackingCode: text('tracking_code'),
    turnKey: text('turn_key').notNull(),
    unitTurnId: text('unit_turn_id'),
    unitId: text('unit_id'),
    propertyId: text('property_id'),
    unitName: text('unit_name'),
    propertyName: text('property_name'),
    status: text('status').notNull().default('open'),
    confidenceScore: real('confidence_score'),
    confidenceLabel: text('confidence_label'),
    sourceFlags: jsonb('source_flags').notNull().default({}),
    metadata: jsonb('metadata').notNull().default({}),
    closedAt: timestamp('closed_at', { withTimezone: true }),
    createdAt: timestamp('created_at', { withTimezone: true }).defaultNow().notNull(),
    updatedAt: timestamp('updated_at', { withTimezone: true }).defaultNow().notNull(),
  },
  (table) => ({
    turnKeyUnique: unique('unit_turn_tracker_turn_key_unique').on(table.turnKey),
    statusIdx: index('unit_turn_tracker_status_idx').on(table.status),
    unitIdx: index('unit_turn_tracker_unit_idx').on(table.unitId),
    propertyIdx: index('unit_turn_tracker_property_idx').on(table.propertyId),
  }),
);

export const unitTurnMilestones = pgTable(
  'unit_turn_milestones',
  {
    id: integer('id').primaryKey().generatedAlwaysAsIdentity(),
    trackingUuid: text('tracking_uuid').notNull(),
    milestoneKey: text('milestone_key').notNull(),
    milestoneDate: timestamp('milestone_date', { withTimezone: true }),
    source: text('source'),
    notes: text('notes'),
    createdAt: timestamp('created_at', { withTimezone: true }).defaultNow().notNull(),
  },
  (table) => ({
    trackingIdx: index('unit_turn_milestones_tracking_idx').on(table.trackingUuid),
  }),
);

export const unitTurnWorkOrders = pgTable(
  'unit_turn_work_orders',
  {
    id: integer('id').primaryKey().generatedAlwaysAsIdentity(),
    trackingUuid: text('tracking_uuid').notNull(),
    woId: text('wo_id').notNull(),
    woDbUuid: text('wo_db_uuid'),
    source: text('source').default('manual'),
    status: text('status'),
    removed: boolean('removed').default(false),
    createdAt: timestamp('created_at', { withTimezone: true }).defaultNow().notNull(),
  },
  (table) => ({
    trackingIdx: index('unit_turn_work_orders_tracking_idx').on(table.trackingUuid),
    woIdx: index('unit_turn_work_orders_wo_idx').on(table.woId),
  }),
);

// Reassignment / queue engine
export const reassignmentQueue = pgTable('reassignment_queue', {
  woId: text('wo_id').primaryKey(),
  woNumber: text('wo_number'),
  propertyGroupUuid: text('property_group_uuid'),
  propertyAddress: text('property_address'),
  category: text('category'),
  priority: text('priority'),
  assignedUserId: text('assigned_user_id'),
  assignedUserName: text('assigned_user_name'),
  status: text('status').default('pending'),
  score: real('score').default(0),
  autoExempt: boolean('auto_exempt').default(false),
  autoExemptAt: timestamp('auto_exempt_at', { withTimezone: true }),
  autoExemptBy: text('auto_exempt_by'),
  updatedAt: timestamp('updated_at', { withTimezone: true }).defaultNow().notNull(),
  createdAt: timestamp('created_at', { withTimezone: true }).defaultNow().notNull(),
});

export const techGrades = pgTable('tech_grades', {
  techId: text('tech_id').primaryKey(),
  techName: text('tech_name'),
  techEmail: text('tech_email'),
  tier: integer('tier').default(1),
  grade: real('grade').default(0),
  jobsCompleted: integer('jobs_completed').default(0),
  noContactCount: integer('no_contact_count').default(0),
  active: boolean('active').default(true),
  updatedAt: timestamp('updated_at', { withTimezone: true }).defaultNow().notNull(),
  createdAt: timestamp('created_at', { withTimezone: true }).defaultNow().notNull(),
});

// Device auth and PM access tables
export const trustedDevices = pgTable('trusted_devices', {
  deviceToken: text('device_token').primaryKey(),
  userName: text('user_name'),
  role: text('role').default('full'),
  loginEmail: text('login_email'),
  propertyGroupUuid: text('property_group_uuid'),
  phone: text('phone'),
  revoked: boolean('revoked').default(false),
  lastSeenAt: timestamp('last_seen_at', { withTimezone: true }),
  expiresAt: timestamp('expires_at', { withTimezone: true }),
  createdAt: timestamp('created_at', { withTimezone: true }).defaultNow().notNull(),
});

export const deviceOtps = pgTable(
  'device_otps',
  {
    id: integer('id').primaryKey().generatedAlwaysAsIdentity(),
    email: text('email').notNull(),
    code: text('code').notNull(),
    used: boolean('used').default(false),
    expiresAt: timestamp('expires_at', { withTimezone: true }).notNull(),
    userName: text('user_name'),
    roleHint: text('role_hint'),
    propertyGroupUuid: text('property_group_uuid'),
    createdAt: timestamp('created_at', { withTimezone: true }).defaultNow().notNull(),
    usedAt: timestamp('used_at', { withTimezone: true }),
  },
  (table) => ({
    emailIdx: index('device_otps_email_idx').on(table.email),
  }),
);

export const pmProxyUsers = pgTable(
  'pm_proxy_users',
  {
    userUuid: text('user_uuid').primaryKey(),
    email: text('email').notNull(),
    fullName: text('full_name'),
    phone: text('phone'),
    propertyGroupUuid: text('property_group_uuid'),
    roles: jsonb('roles').notNull().default([]),
    active: boolean('active').default(true),
    rawJson: jsonb('raw_json').notNull().default({}),
    createdAt: timestamp('created_at', { withTimezone: true }).defaultNow().notNull(),
    updatedAt: timestamp('updated_at', { withTimezone: true }).defaultNow().notNull(),
  },
  (table) => ({
    emailUnique: unique('pm_proxy_users_email_unique').on(table.email),
    propertyGroupIdx: index('pm_proxy_users_group_idx').on(table.propertyGroupUuid),
  }),
);

export const pmProxyUserScopes = pgTable(
  'pm_proxy_user_scopes',
  {
    id: integer('id').primaryKey().generatedAlwaysAsIdentity(),
    userUuid: text('user_uuid').notNull(),
    propertyGroupUuid: text('property_group_uuid').notNull(),
    isPrimary: boolean('is_primary').default(false),
    active: boolean('active').default(true),
    source: text('source').default('manual'),
    createdAt: timestamp('created_at', { withTimezone: true }).defaultNow().notNull(),
    updatedAt: timestamp('updated_at', { withTimezone: true }).defaultNow().notNull(),
  },
  (table) => ({
    userScopeUnique: unique('pm_proxy_user_scopes_user_scope_unique').on(table.userUuid, table.propertyGroupUuid),
    userIdx: index('pm_proxy_user_scopes_user_idx').on(table.userUuid),
    groupIdx: index('pm_proxy_user_scopes_group_idx').on(table.propertyGroupUuid),
  }),
);

// Webhook persistence
export const webhookEvents = pgTable(
  'webhook_events',
  {
    id: integer('id').primaryKey().generatedAlwaysAsIdentity(),
    eventId: text('event_id'),
    topic: text('topic').notNull(),
    eventType: text('event_type'),
    resourceType: text('resource_type'),
    resourceId: text('resource_id'),
    signature: text('signature'),
    rawPayload: jsonb('raw_payload').notNull().default({}),
    receivedAt: timestamp('received_at', { withTimezone: true }).defaultNow().notNull(),
    processedAt: timestamp('processed_at', { withTimezone: true }),
    processingStatus: text('processing_status').default('pending'),
  },
  (table) => ({
    topicIdx: index('webhook_events_topic_idx').on(table.topic),
    resourceIdx: index('webhook_events_resource_idx').on(table.resourceType, table.resourceId),
    statusIdx: index('webhook_events_status_idx').on(table.processingStatus),
    eventIdIdx: index('webhook_events_event_id_idx').on(table.eventId),
  }),
);

// Shared cache/config helpers used by migrated handlers
export const apiCache = pgTable('api_cache', {
  cacheKey: text('cache_key').primaryKey(),
  data: jsonb('data'),
  cachedAt: timestamp('cached_at', { withTimezone: true }).defaultNow(),
  ttlMs: integer('ttl_ms'),
  etag: text('etag'),
  lastModified: text('last_modified'),
  source: text('source'),
  expiresAt: timestamp('expires_at', { withTimezone: true }),
});

export const proxyConfig = pgTable('proxy_config', {
  key: text('key').primaryKey(),
  value: text('value'),
  updatedAt: timestamp('updated_at', { withTimezone: true }).defaultNow().notNull(),
});

export const appSettings = pgTable('app_settings', {
  key: text('key').primaryKey(),
  value: text('value'),
  updatedAt: timestamp('updated_at', { withTimezone: true }).defaultNow().notNull(),
});

// ── Sync Control Plane ────────────────────────────────────────────────────────

// One row per sync job execution (backfill run or scheduled sync).
export const syncJobRuns = pgTable(
  'sync_job_runs',
  {
    id: integer('id').primaryKey().generatedAlwaysAsIdentity(),
    runId: text('run_id').notNull(),           // UUID assigned at job start
    endpointKey: text('endpoint_key').notNull(), // e.g. 'v0:units', 'v2:work_orders'
    apiVersion: text('api_version').notNull(),   // 'v0' | 'v2'
    triggerType: text('trigger_type').notNull(), // 'backfill' | 'nightly' | 'incremental' | 'webhook' | 'manual'
    status: text('status').notNull().default('running'), // 'running' | 'completed' | 'failed' | 'paused'
    filtersFingerprint: text('filters_fingerprint'), // SHA-256 of sorted filter params
    pagesCompleted: integer('pages_completed').default(0).notNull(),
    rowsUpserted: integer('rows_upserted').default(0).notNull(),
    rowsSkipped: integer('rows_skipped').default(0).notNull(),
    lastError: text('last_error'),
    startedAt: timestamp('started_at', { withTimezone: true }).defaultNow().notNull(),
    completedAt: timestamp('completed_at', { withTimezone: true }),
    executionStartCursor: text('execution_start_cursor'), // timestamp recorded at job START (not last record) for incremental use
  },
  (table) => ({
    runIdIdx: index('sync_job_runs_run_id_idx').on(table.runId),
    endpointStatusIdx: index('sync_job_runs_endpoint_status_idx').on(table.endpointKey, table.status),
  }),
);

// Persists the exact cursor value between pages/runs so jobs are resumable.
export const syncJobCursors = pgTable(
  'sync_job_cursors',
  {
    id: integer('id').primaryKey().generatedAlwaysAsIdentity(),
    runId: text('run_id').notNull(),
    endpointKey: text('endpoint_key').notNull(),
    pageIndex: integer('page_index').notNull().default(0),
    cursorIn: text('cursor_in'),              // cursor used to fetch this page
    cursorOut: text('cursor_out'),            // next_page_path or next_page_url from response
    cursorExpiresAt: timestamp('cursor_expires_at', { withTimezone: true }), // v2 URLs expire after 30 min
    fetchedAt: timestamp('fetched_at', { withTimezone: true }).defaultNow().notNull(),
    recordCount: integer('record_count').default(0),
    retriesUsed: integer('retries_used').default(0),
    isTerminal: boolean('is_terminal').default(false), // true when cursor_out is null (last page)
  },
  (table) => ({
    runPageIdx: index('sync_job_cursors_run_page_idx').on(table.runId, table.pageIndex),
    endpointIdx: index('sync_job_cursors_endpoint_idx').on(table.endpointKey),
  }),
);

// Universal raw response archive — written before any mapping/upsert step.
export const appfolioRawResponses = pgTable(
  'appfolio_raw_responses',
  {
    id: integer('id').primaryKey().generatedAlwaysAsIdentity(),
    runId: text('run_id').notNull(),
    endpointKey: text('endpoint_key').notNull(),
    pageIndex: integer('page_index').notNull().default(0),
    cursorIn: text('cursor_in'),
    cursorOut: text('cursor_out'),
    statusCode: integer('status_code'),
    recordCount: integer('record_count').default(0),
    responseJson: jsonb('response_json').notNull().default({}),
    fetchedAt: timestamp('fetched_at', { withTimezone: true }).defaultNow().notNull(),
  },
  (table) => ({
    runIdx: index('appfolio_raw_responses_run_idx').on(table.runId),
    endpointIdx: index('appfolio_raw_responses_endpoint_idx').on(table.endpointKey),
    fetchedAtIdx: index('appfolio_raw_responses_fetched_at_idx').on(table.fetchedAt),
  }),
);

// Per-call request log for observability, 429 forensics, and rate budget tracking.
export const appfolioRequestLog = pgTable(
  'appfolio_request_log',
  {
    id: integer('id').primaryKey().generatedAlwaysAsIdentity(),
    runId: text('run_id'),
    endpointKey: text('endpoint_key').notNull(),
    apiVersion: text('api_version'),
    method: text('method').default('GET'),
    statusCode: integer('status_code'),
    latencyMs: integer('latency_ms'),
    retryAfterSeconds: integer('retry_after_seconds'),
    attemptNumber: integer('attempt_number').default(1),
    errorText: text('error_text'),
    cursorSnapshot: text('cursor_snapshot'),
    requestedAt: timestamp('requested_at', { withTimezone: true }).defaultNow().notNull(),
  },
  (table) => ({
    endpointTimeIdx: index('appfolio_request_log_endpoint_time_idx').on(table.endpointKey, table.requestedAt),
    runIdx: index('appfolio_request_log_run_idx').on(table.runId),
  }),
);

// Rate counters per window type — used by the token-bucket rate limiter.
export const appfolioRateCounters = pgTable(
  'appfolio_rate_counters',
  {
    apiVersion: text('api_version').notNull(),
    endpointKey: text('endpoint_key').notNull(),
    windowType: text('window_type').notNull(), // 'second' | 'minute' | 'hour'
    windowStart: text('window_start').notNull(), // ISO timestamp truncated to window
    requestCount: integer('request_count').notNull().default(0),
    status429Count: integer('status_429_count').notNull().default(0),
  },
  (table) => ({
    pk: unique('appfolio_rate_counters_pk').on(table.apiVersion, table.endpointKey, table.windowType, table.windowStart),
    windowIdx: index('appfolio_rate_counters_window_idx').on(table.windowType, table.windowStart),
  }),
);

// Serialized PATCH queue — ensures no two concurrent PATCHes to the same resource.
export const appfolioPatchQueue = pgTable(
  'appfolio_patch_queue',
  {
    id: integer('id').primaryKey().generatedAlwaysAsIdentity(),
    resourceId: text('resource_id').notNull(),   // AppFolio entity UUID / WO number
    resourceType: text('resource_type').notNull(), // 'work_order' | 'estimate' | 'turn'
    method: text('method').default('PATCH'),
    endpointPath: text('endpoint_path').notNull(),
    payloadJson: jsonb('payload_json').notNull().default({}),
    status: text('status').notNull().default('pending'), // 'pending' | 'in_flight' | 'done' | 'failed'
    priority: integer('priority').default(100),
    attemptCount: integer('attempt_count').default(0),
    lastError: text('last_error'),
    lockedAt: timestamp('locked_at', { withTimezone: true }),
    lockOwner: text('lock_owner'),
    createdAt: timestamp('created_at', { withTimezone: true }).defaultNow().notNull(),
    updatedAt: timestamp('updated_at', { withTimezone: true }).defaultNow().notNull(),
  },
  (table) => ({
    resourcePendingIdx: index('appfolio_patch_queue_resource_pending_idx').on(table.resourceId, table.status),
    statusPriorityIdx: index('appfolio_patch_queue_status_priority_idx').on(table.status, table.priority, table.id),
  }),
);
