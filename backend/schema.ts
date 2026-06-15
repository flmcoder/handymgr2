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
    woNumberIdx: index('appfolio_work_orders_number_idx').on(table.woNumber),
    statusIdx: index('appfolio_work_orders_status_idx').on(table.status),
    propertyIdx: index('appfolio_work_orders_property_idx').on(table.propertyId),
    unitIdx: index('appfolio_work_orders_unit_idx').on(table.unitId),
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
    propertyGroupUuid: text('property_group_uuid').notNull(),
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

// Webhook persistence
export const webhookEvents = pgTable(
  'webhook_events',
  {
    id: integer('id').primaryKey().generatedAlwaysAsIdentity(),
    eventUuid: text('event_uuid'),
    topic: text('topic').notNull(),
    eventType: text('event_type'),
    resourceType: text('resource_type'),
    resourceId: text('resource_id'),
    signature: text('signature'),
    payloadJson: jsonb('payload_json').notNull().default({}),
    receivedAt: timestamp('received_at', { withTimezone: true }).defaultNow().notNull(),
    processedAt: timestamp('processed_at', { withTimezone: true }),
    processingStatus: text('processing_status').default('pending'),
  },
  (table) => ({
    topicIdx: index('webhook_events_topic_idx').on(table.topic),
    resourceIdx: index('webhook_events_resource_idx').on(table.resourceType, table.resourceId),
    statusIdx: index('webhook_events_status_idx').on(table.processingStatus),
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
