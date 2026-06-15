// ============================================================================
// schema.ts — Drizzle ORM PostgreSQL Schema Definitions
// AppFolio API v0 Data Structures + Internal Tracking Tables
// ============================================================================

import { pgTable, text, integer, real, timestamp, boolean, jsonb, index, unique } from 'drizzle-orm/pg-core';

// ── AppFolio API v0 Core Resources ──────────────────────────────────────────

export const appfolioProperties = pgTable('appfolio_properties', {
  id: text('id').primaryKey(), // AppFolio Property UUID
  name: text('name').notNull(),
  propertyGroupId: text('property_group_id'),
  street: text('street'),
  city: text('city'),
  state: text('state'),
  zip: text('zip'),
  managementStatus: text('management_status'),
  rawJson: jsonb('raw_json'),
  cachedAt: timestamp('cached_at').defaultNow(),
  updatedAt: timestamp('updated_at'),
}, (table) => ({
  groupIdx: index('properties_group_idx').on(table.propertyGroupId),
  nameIdx: index('properties_name_idx').on(table.name),
}));

export const appfolioUnits = pgTable('appfolio_units', {
  unitId: text('unit_id').primaryKey(), // AppFolio Unit UUID
  propertyId: text('property_id').notNull(),
  name: text('name'),
  unitNumber: text('unit_number'),
  bedrooms: integer('bedrooms'),
  bathrooms: real('bathrooms'),
  squareFeet: integer('square_feet'),
  marketRent: real('market_rent'),
  status: text('status'), // Occupied, Vacant, etc.
  rawJson: jsonb('raw_json'),
  cachedAt: timestamp('cached_at').defaultNow(),
  updatedAt: timestamp('updated_at'),
}, (table) => ({
  propertyIdx: index('units_property_idx').on(table.propertyId),
  statusIdx: index('units_status_idx').on(table.status),
}));

export const appfolioWorkOrders = pgTable('appfolio_work_orders', {
  id: text('id').primaryKey(), // AppFolio Work Order UUID
  woNumber: text('wo_number'),
  propertyId: text('property_id'),
  unitId: text('unit_id'),
  propertyGroupId: text('property_group_id'),
  description: text('description'),
  category: text('category'),
  priority: text('priority'),
  status: text('status'),
  statusLabel: text('status_label'),
  assignedUserId: text('assigned_user_id'),
  assignedUserName: text('assigned_user_name'),
  vendorId: text('vendor_id'),
  vendorName: text('vendor_name'),
  estimatedAmount: real('estimated_amount'),
  totalCost: real('total_cost'),
  scheduledStart: timestamp('scheduled_start'),
  completedAt: timestamp('completed_at'),
  createdAt: timestamp('created_at'),
  rawJson: jsonb('raw_json').notNull().default('{}'),
  cachedAt: timestamp('cached_at').defaultNow(),
  updatedAt: timestamp('updated_at'),
}, (table) => ({
  woNumberIdx: index('wo_number_idx').on(table.woNumber),
  statusIdx: index('wo_status_idx').on(table.status),
  propertyIdx: index('wo_property_idx').on(table.propertyId),
  groupIdx: index('wo_group_idx').on(table.propertyGroupId),
  vendorIdx: index('wo_vendor_idx').on(table.vendorId),
}));

export const appfolioVendors = pgTable('appfolio_vendors', {
  id: text('id').primaryKey(), // AppFolio Vendor UUID
  name: text('name').notNull(),
  companyName: text('company_name'),
  email: text('email'),
  phone: text('phone'),
  licenseNumber: text('license_number'),
  insuranceExpiry: timestamp('insurance_expiry'),
  businessType: text('business_type'),
  tradeCategory: text('trade_category'),
  isCompliant: boolean('is_compliant'),
  rawJson: jsonb('raw_json'),
  cachedAt: timestamp('cached_at').defaultNow(),
  updatedAt: timestamp('updated_at'),
}, (table) => ({
  nameIdx: index('vendors_name_idx').on(table.name),
  tradeCategoryIdx: index('vendors_trade_idx').on(table.tradeCategory),
}));

export const appfolioBills = pgTable('appfolio_bills', {
  id: text('id').primaryKey(), // AppFolio Bill UUID
  billNumber: text('bill_number'),
  vendorId: text('vendor_id'),
  vendorName: text('vendor_name'),
  propertyId: text('property_id'),
  propertyName: text('property_name'),
  unitId: text('unit_id'),
  status: text('status'),
  statusLabel: text('status_label'),
  billTotalAmount: real('bill_total_amount'),
  invoiceDate: timestamp('invoice_date'),
  dueDate: timestamp('due_date'),
  paidAt: timestamp('paid_at'),
  rawJson: jsonb('raw_json'),
  cachedAt: timestamp('cached_at').defaultNow(),
  updatedAt: timestamp('updated_at'),
}, (table) => ({
  vendorIdx: index('bills_vendor_idx').on(table.vendorId),
  statusIdx: index('bills_status_idx').on(table.status),
  propertyIdx: index('bills_property_idx').on(table.propertyId),
}));

export const appfolioPropertyGroups = pgTable('appfolio_property_groups', {
  id: text('id').primaryKey(), // AppFolio Property Group UUID
  name: text('name').notNull(),
  type: text('type'),
  propertyIds: jsonb('property_ids').default('[]'),
  rawJson: jsonb('raw_json'),
  cachedAt: timestamp('cached_at').defaultNow(),
  updatedAt: timestamp('updated_at'),
}, (table) => ({
  nameIdx: index('property_groups_name_idx').on(table.name),
}));

// ── Internal Tracking Tables ────────────────────────────────────────────────

export const unitTurnTracker = pgTable('unit_turn_tracker', {
  trackingUuid: text('tracking_uuid').primaryKey(),
  trackingCode: text('tracking_code'),
  turnKey: text('turn_key').notNull(),
  unitTurnId: text('unit_turn_id'),
  unitId: text('unit_id'),
  propertyId: text('property_id'),
  unitName: text('unit_name'),
  propertyName: text('property_name'),
  moveOutDate: text('move_out_date'),
  moveInDate: text('move_in_date'),
  inspectionDate: text('inspection_date'),
  firstWoDate: text('first_wo_date'),
  estimateRequestedDate: text('estimate_requested_date'),
  estimateReceivedDate: text('estimate_received_date'),
  status: text('status').notNull().default('open'),
  confidenceScore: real('confidence_score'),
  confidenceLabel: text('confidence_label'),
  siteManager: text('site_manager'),
  sourceFlags: jsonb('source_flags').default('{}'),
  metadata: jsonb('metadata').default('{}'),
  closedAt: text('closed_at'),
  createdAt: text('created_at').notNull().default('datetime(\'now\')'),
  updatedAt: text('updated_at').notNull().default('datetime(\'now\')'),
}, (table) => ({
  turnKeyIdx: index('turn_tracker_key_idx').on(table.turnKey),
  statusIdx: index('turn_tracker_status_idx').on(table.status),
  unitIdx: index('turn_tracker_unit_idx').on(table.unitId),
  uniqueTurnKey: unique('turn_tracker_turn_key_unique').on(table.turnKey),
}));

export const unitTurnMilestones = pgTable('unit_turn_milestones', {
  id: integer('id').primaryKey().generatedAlwaysAsIdentity(),
  trackingUuid: text('tracking_uuid').notNull(),
  milestoneKey: text('milestone_key').notNull(),
  milestoneDate: text('milestone_date'),
  source: text('source'),
  notes: text('notes'),
  createdAt: text('created_at').notNull().default('datetime(\'now\')'),
}, (table) => ({
  trackingIdx: index('turn_milestones_tracking_idx').on(table.trackingUuid),
}));

export const unitTurnWorkOrders = pgTable('unit_turn_work_orders', {
  id: integer('id').primaryKey().generatedAlwaysAsIdentity(),
  trackingUuid: text('tracking_uuid').notNull(),
  woId: text('wo_id').notNull(),
  woDbUuid: text('wo_db_uuid'),
  source: text('source').default('manual'),
  status: text('status'),
  removed: integer('removed').default(0),
  createdAt: text('created_at').notNull().default('datetime(\'now\')'),
}, (table) => ({
  trackingIdx: index('turn_wo_tracking_idx').on(table.trackingUuid),
  woIdx: index('turn_wo_wo_idx').on(table.woId),
}));

export const closedTurns = pgTable('closed_turns', {
  turnId: text('turn_id').primaryKey(),
  closedAt: text('closed_at'),
  closeReason: text('close_reason'),
  closeSource: text('close_source'),
  closedBy: text('closed_by'),
  propertyId: text('property_id'),
  propertyName: text('property_name'),
  unitId: text('unit_id'),
  unitName: text('unit_name'),
  moveOutDate: text('move_out_date'),
  moveInDate: text('move_in_date'),
});

export const estimates = pgTable('estimates', {
  estimateId: text('estimate_id').primaryKey(),
  workOrderId: text('work_order_id'),
  workOrderNumber: text('work_order_number'),
  currentStatus: text('current_status').notNull(),
  propertyGroupId: text('property_group_id'),
  source: text('source'),
  statusHistoryJson: text('status_history_json').notNull().default('[]'),
  rawData: text('raw_data').notNull().default('{}'),
  createdAt: text('created_at').notNull().default('datetime(\'now\')'),
  updatedAt: text('updated_at').notNull().default('datetime(\'now\')'),
}, (table) => ({
  statusIdx: index('estimates_status_idx').on(table.currentStatus),
  groupIdx: index('estimates_group_idx').on(table.propertyGroupId),
  woIdx: index('estimates_wo_idx').on(table.workOrderId),
}));

// ── Auth & Session Tables ───────────────────────────────────────────────────

export const trustedDevices = pgTable('trusted_devices', {
  deviceToken: text('device_token').primaryKey(),
  userName: text('user_name').notNull(),
  role: text('role').notNull().default('full'),
  loginEmail: text('login_email'),
  propertyGroupUuid: text('property_group_uuid'),
  phone: text('phone'),
  lastSeenAt: text('last_seen_at'),
  expiresAt: text('expires_at'),
  createdAt: text('created_at').notNull().default('datetime(\'now\')'),
});

export const pmProxyUsers = pgTable('pm_proxy_users', {
  id: integer('id').primaryKey().generatedAlwaysAsIdentity(),
  email: text('email'),
  phone: text('phone'),
  propertyGroupUuid: text('property_group_uuid').notNull(),
  propertyGroupName: text('property_group_name'),
  fullName: text('full_name'),
  active: integer('active').default(1),
  createdAt: text('created_at').notNull().default('datetime(\'now\')'),
  updatedAt: text('updated_at').notNull().default('datetime(\'now\')'),
}, (table) => ({
  emailIdx: index('pm_users_email_idx').on(table.email),
  phoneIdx: index('pm_users_phone_idx').on(table.phone),
}));

export const deviceOtps = pgTable('device_otps', {
  identifier: text('identifier').primaryKey(),
  code: text('code').notNull(),
  phone: text('phone'),
  email: text('email'),
  expiresAt: text('expires_at').notNull(),
  verified: integer('verified').default(0),
  createdAt: text('created_at').notNull().default('datetime(\'now\')'),
});

// ── Reassignment / Dispatch Tables ──────────────────────────────────────────

export const reassignmentQueue = pgTable('reassignment_queue', {
  woId: text('wo_id').primaryKey(),
  woNumber: text('wo_number'),
  propertyGroupUuid: text('property_group_uuid'),
  propertyAddress: text('property_address'),
  category: text('category'),
  priority: text('priority'),
  assignedUserId: text('assigned_user_id'),
  assignedUserName: text('assigned_user_name'),
  lastAssignedAt: text('last_assigned_at'),
  autoExempt: integer('auto_exempt').default(0),
  autoExemptAt: text('auto_exempt_at'),
  autoExemptBy: text('auto_exempt_by'),
  status: text('status').default('pending'),
  score: real('score').default(0),
  createdAt: text('created_at').notNull().default('datetime(\'now\')'),
  updatedAt: text('updated_at').notNull().default('datetime(\'now\')'),
});

export const techGrades = pgTable('tech_grades', {
  techId: text('tech_id').primaryKey(),
  techName: text('tech_name'),
  grade: real('grade').default(0),
  jobsCount: integer('jobs_count').default(0),
  onTimeRate: real('on_time_rate').default(0),
  qualityScore: real('quality_score').default(0),
  lastScoredAt: text('last_scored_at'),
  createdAt: text('created_at').notNull().default('datetime(\'now\')'),
  updatedAt: text('updated_at').notNull().default('datetime(\'now\')'),
});

// ── Webhook Event Storage ───────────────────────────────────────────────────

export const webhookEvents = pgTable('webhook_events', {
  id: integer('id').primaryKey().generatedAlwaysAsIdentity(),
  eventUuid: text('event_uuid'),
  topic: text('topic').notNull(),
  resourceType: text('resource_type'),
  resourceId: text('resource_id'),
  action: text('action'),
  payloadJson: jsonb('payload_json').notNull().default('{}'),
  signature: text('signature'),
  receivedAt: text('received_at').notNull().default('datetime(\'now\')'),
  processedAt: text('processed_at'),
  processingStatus: text('processing_status').default('pending'),
}, (table) => ({
  topicIdx: index('webhook_events_topic_idx').on(table.topic),
  resourceIdx: index('webhook_events_resource_idx').on(table.resourceType, table.resourceId),
  statusIdx: index('webhook_events_status_idx').on(table.processingStatus),
}));

// ── Cache & Config Tables ───────────────────────────────────────────────────

export const apiCache = pgTable('api_cache', {
  cacheKey: text('cache_key').primaryKey(),
  data: text('data'),
  cachedAt: integer('cached_at'),
  ttlMs: integer('ttl_ms'),
  etag: text('etag'),
  lastModified: text('last_modified'),
  source: text('source'),
  expiresAt: text('expires_at'),
});

export const proxyConfig = pgTable('proxy_config', {
  key: text('key').primaryKey(),
  value: text('value'),
  updatedAt: text('updated_at').notNull().default('datetime(\'now\')'),
});

export const appSettings = pgTable('app_settings', {
  key: text('key').primaryKey(),
  value: text('value'),
  updatedAt: text('updated_at').notNull().default('datetime(\'now\')'),
});
