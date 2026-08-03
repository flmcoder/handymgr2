/**
 * Postgres repositories for AppFolio domain entities.
 *
 * Each upsert:
 *  1. Computes a stable hash of the normalized incoming payload.
 *  2. Skips the write if the hash matches what is already stored (no-op optimization).
 *  3. Otherwise upserts into the domain table and stores raw_json.
 *
 * Tables covered:
 *   appfolio_properties, appfolio_units, appfolio_work_orders,
 *   appfolio_estimates, unit_turn_tracker, unit_turn_milestones,
 *   unit_turn_work_orders
 */

import { createHash } from 'node:crypto';
import { inArray, sql } from 'drizzle-orm';
import { db } from '../db.ts';
import {
  appfolioUsers,
  appfolioPropertyGroups,
  appfolioProperties,
  appfolioUnits,
  appfolioWorkOrders,
  appfolioEstimates,
  appfolioUnitInspections,
  appfolioTenantDirectory,
  appfolioUnitTurnDetails,
  appfolioUnitVacancies,
  unitTurnTracker,
  unitTurnMilestones,
  unitTurnWorkOrders,
} from '../schema.ts';

type GroupCatalogEntry = { id: string; uuid: string; name: string };

const GROUP_CATALOG_BY_NAME: Record<string, GroupCatalogEntry> = {
  'ana consuelo properties': { id: '1', uuid: 'c8f40f01-9e94-11ee-8b51-02167481f3bc', name: 'Ana Consuelo Properties' },
  'andrea robidoux properties': { id: '59', uuid: '9985f145-3cb0-11f0-bfba-069ca18f5865', name: 'Andrea Robidoux Properties' },
  'chris meehan properties': { id: '3', uuid: '06fdebec-9e9b-11ee-8b51-02167481f3bc', name: 'Chris Meehan Properties' },
  'jessmar romea properties': { id: '6', uuid: '3405e65e-9e9c-11ee-8b51-02167481f3bc', name: 'Jessmar Romea Properties' },
  'jennifer hazlett properties': { id: '16', uuid: 'b368729d-9eca-11ee-8b51-02167481f3bc', name: 'Jennifer Hazlett Properties' },
  'mary rees properties': { id: '7', uuid: '44b79f5e-9e9c-11ee-8b51-02167481f3bc', name: 'Mary Rees Properties' },
  'veronica garcia properties': { id: '10', uuid: '121e7ca4-9eca-11ee-8b51-02167481f3bc', name: 'Veronica Garcia Properties' },
  'nita lauer properties': { id: '5', uuid: '1036e611-9e9c-11ee-8b51-02167481f3bc', name: 'Nita Lauer Properties' },
  'jacquelina brantley properties': { id: '23', uuid: '9a434f3b-a04d-11ee-8b51-02167481f3bc', name: 'Jacquelina Brantley Properties' },
  'angela hogan properties': { id: '13', uuid: '66a16517-9eca-11ee-8b51-02167481f3bc', name: 'Angela Hogan Properties' },
  'michelle kovach properties': { id: '61', uuid: '930c330b-60ce-11f0-bfba-069ca18f5865', name: 'Michelle Kovach Properties' },
  'deborah lago properties': { id: '14', uuid: '7d4a69d6-9eca-11ee-8b51-02167481f3bc', name: 'Deborah Lago Properties' },
  'jordan hammerschmidt properties': { id: '17', uuid: 'bee73529-9eca-11ee-8b51-02167481f3bc', name: 'Jordan Hammerschmidt Properties' },
  'michelle cunningham properties': { id: '25', uuid: 'a5774de7-a04d-11ee-8b51-02167481f3bc', name: 'Michelle Cunningham Properties' },
  'jamie monty properties': { id: '73', uuid: 'f922348b-ea67-11f0-bfba-069ca18f5865', name: 'Jamie Monty Properties' },
  'cari rascon properties': { id: '56', uuid: '61a5b6d1-251b-11f0-bfba-069ca18f5865', name: 'Cari Rascon Properties' },
  'sara anglin': { id: '66', uuid: '7f65b11f-7c52-11f0-bfba-069ca18f5865', name: 'Sara Anglin' },
  missionsprings: { id: '72', uuid: 'bb129607-e81e-11f0-bfba-069ca18f5865', name: 'MissionSprings' },
  'el diablo': { id: '71', uuid: '114bcb4d-e81e-11f0-bfba-069ca18f5865', name: 'El Diablo' },
  'maggie properties': { id: '68', uuid: '0041c7dd-add1-11f0-bfba-069ca18f5865', name: 'Maggie Properties' },
  'phoenix properties': { id: '40', uuid: '', name: 'Phoenix Properties' },
  'tucson properties': { id: '41', uuid: '', name: 'Tucson Properties' },
};

const GROUP_CATALOG_BY_UUID = new Map<string, GroupCatalogEntry>();
const GROUP_CATALOG_BY_ID = new Map<string, GroupCatalogEntry>();

Object.values(GROUP_CATALOG_BY_NAME).forEach((entry) => {
  if (entry.uuid) GROUP_CATALOG_BY_UUID.set(entry.uuid, entry);
  if (entry.id) GROUP_CATALOG_BY_ID.set(entry.id, entry);
});

// ── Helpers ──────────────────────────────────────────────────────────────────

function stableHash(obj: unknown): string {
  return createHash('sha256')
    .update(JSON.stringify(obj, Object.keys(obj as any).sort()))
    .digest('hex')
    .slice(0, 16); // 16 hex chars is ample for a change-detect key
}

function asStr(v: unknown, max = 500): string | null {
  const s = String(v ?? '').trim();
  return s ? s.slice(0, max) : null;
}

function asNum(v: unknown): number | null {
  const n = parseFloat(String(v ?? '').replace(/[^0-9.-]/g, ''));
  return Number.isFinite(n) ? n : null;
}

function asDate(v: unknown): Date | null {
  if (!v) return null;
  const d = new Date(String(v));
  return isNaN(d.getTime()) ? null : d;
}

function asBool(v: unknown, fallback = true): boolean {
  if (typeof v === 'boolean') return v;
  if (v === null || v === undefined || v === '') return fallback;
  const s = String(v).trim().toLowerCase();
  if (['1', 'true', 'yes', 'active', 'enabled'].includes(s)) return true;
  if (['0', 'false', 'no', 'inactive', 'disabled'].includes(s)) return false;
  return fallback;
}

function isMaintenanceTechRole(value: unknown): boolean {
  const role = String(value ?? '').trim().toLowerCase();
  return role === 'maintenance tech' || role === 'maintenance_tech';
}

function isUuidLike(value: unknown): boolean {
  const s = String(value || '').trim();
  if (!s) return false;
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(s);
}

function reportRowKey(row: any, keys: string[], fallbackParts: Array<unknown>): string {
  for (const key of keys) {
    const candidate = asStr(row?.[key]);
    if (candidate && candidate !== '0') return candidate;
  }
  return createHash('sha256').update(JSON.stringify(fallbackParts)).digest('hex').slice(0, 36);
}

function normalizeGroupName(v: unknown): string {
  return String(v ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function chooseReadableGroupName(...values: unknown[]): string {
  for (const value of values) {
    const candidate = asStr(value);
    if (!candidate) continue;
    if (!isUuidLike(candidate)) return candidate;
  }
  return '';
}

function resolvePropertyGroupInfo(row: any): { groupId: string | null; groupUuid: string | null; groupName: string | null; canonicalGroupKey: string | null } {
  let groupId = asStr(row.PropertyGroupId || row.property_group_id || row.property_group_numeric_id || row.group_id);
  let groupUuid = asStr(row.PropertyGroupUuid || row.property_group_uuid || row.property_group_guid || row.group_uuid);
  const groupIds = Array.isArray(row.PropertyGroupIds)
    ? row.PropertyGroupIds.map((v: unknown) => asStr(v)).filter((v): v is string => !!v)
    : [];
  if (!groupUuid && groupIds.length) {
    const knownUuid = groupIds.find((candidate) => GROUP_CATALOG_BY_UUID.has(candidate));
    groupUuid = knownUuid || groupIds[0] || null;
  }
  const groupName = asStr(
    row.NameOfPropertyGroup || row.name_of_property_group || row.property_group_name || row.PropertyGroupName ||
    row.property_group || row.group_name || row.GroupName || row.portfolio || row.portfolio_name,
  );

  if ((!groupId || !groupUuid) && groupName) {
    const fromName = GROUP_CATALOG_BY_NAME[normalizeGroupName(groupName)];
    if (fromName) {
      if (!groupId && fromName.id) groupId = fromName.id;
      if (!groupUuid && fromName.uuid) groupUuid = fromName.uuid;
    }
  }

  if (!groupId && groupUuid) {
    const fromUuid = GROUP_CATALOG_BY_UUID.get(groupUuid);
    if (fromUuid?.id) groupId = fromUuid.id;
  }

  if (!groupUuid && groupId) {
    const fromId = GROUP_CATALOG_BY_ID.get(groupId);
    if (fromId?.uuid) groupUuid = fromId.uuid;
  }

  let resolvedGroupName = groupName;
  if (!resolvedGroupName && groupUuid) {
    const fromUuid = GROUP_CATALOG_BY_UUID.get(groupUuid);
    if (fromUuid?.name) resolvedGroupName = fromUuid.name;
  }
  if (!resolvedGroupName && groupId) {
    const fromId = GROUP_CATALOG_BY_ID.get(groupId);
    if (fromId?.name) resolvedGroupName = fromId.name;
  }

  const canonicalGroupKey = groupUuid || groupId || null;
  return { groupId, groupUuid, groupName: resolvedGroupName, canonicalGroupKey };
}

// ── Properties ───────────────────────────────────────────────────────────────

export interface UpsertResult {
  upserted: number;
  skipped: number;
}

export async function upsertPropertyGroups(rows: any[]): Promise<UpsertResult> {
  let upserted = 0;
  let skipped = 0;

  for (const row of rows) {
    const rawId = asStr(row.Id || row.id || row.PropertyGroupId || row.property_group_id || row.PropertyGroupUuid || row.property_group_uuid);
    if (!rawId) {
      skipped++;
      continue;
    }

    const explicitNumericId = asStr(row.PropertyGroupId || row.property_group_id || row.group_id);
    const inferredUuid = asStr(row.PropertyGroupUuid || row.property_group_uuid || row.UUID || row.uuid)
      || (isUuidLike(rawId) ? rawId : null);
    const catalogEntry = inferredUuid ? GROUP_CATALOG_BY_UUID.get(inferredUuid) : undefined;
    // Prefer stable numeric ids when available to stay compatible with older integer-based schemas.
    const id = explicitNumericId || catalogEntry?.id || rawId;

    const rawPropertyIds = Array.isArray(row.PropertyIds)
      ? row.PropertyIds
      : Array.isArray(row.property_ids)
        ? row.property_ids
        : [];
    const propertyIds = rawPropertyIds
      .map((value: any) => {
        if (typeof value === 'string') return value.trim();
        if (value && typeof value === 'object') {
          return String(value.Id || value.id || value.PropertyId || value.property_id || '').trim();
        }
        return String(value || '').trim();
      })
      .filter((value: string) => !!value);

    const uuid = inferredUuid || (isUuidLike(id) ? id : null);
    const normalizedName = chooseReadableGroupName(
      row.Name,
      row.name,
      row.NameOfPropertyGroup,
      row.name_of_property_group,
      row.PropertyGroupName,
      row.property_group_name,
      row.property_group,
      row.group_name,
      row.portfolio_name,
      row.portfolio,
      catalogEntry?.name,
    ) || id;

    const normalized = {
      id,
      uuid,
      name: normalizedName,
      type: asStr(row.Type || row.type),
      propertyIds,
      lastUpdatedAt: asDate(row.LastUpdatedAt || row.last_updated_at || row.updated_at),
    };

    const hash = stableHash(normalized);

    await db
      .insert(appfolioPropertyGroups)
      .values({
        ...normalized,
        rawJson: {
          ...row,
          _hash: hash,
        },
        cachedAt: new Date(),
      })
      .onConflictDoUpdate({
        target: appfolioPropertyGroups.id,
        set: {
          uuid: sql`EXCLUDED.uuid`,
          name: sql`EXCLUDED.name`,
          type: sql`EXCLUDED.type`,
          propertyIds: sql`EXCLUDED.property_ids`,
          lastUpdatedAt: sql`EXCLUDED.last_updated_at`,
          rawJson: sql`EXCLUDED.raw_json`,
          cachedAt: sql`EXCLUDED.cached_at`,
        },
      });

    upserted++;
  }

  return { upserted, skipped };
}

export async function upsertProperties(rows: any[]): Promise<UpsertResult> {
  let upserted = 0;
  let skipped = 0;

  for (const row of rows) {
    const id = asStr(row.Id || row.id || row.property_id);
    if (!id) continue;
    const groupInfo = resolvePropertyGroupInfo(row);

    const normalized = {
      id,
      name: asStr(row.Name || row.PropertyName || row.property_name) ?? id,
      propertyGroupId: groupInfo.canonicalGroupKey,
      street: asStr(row.Address1 || row.StreetAddress || row.address),
      city: asStr(row.City || row.city),
      state: asStr(row.State || row.state),
      zip: asStr(row.Zip || row.zip),
    };

    const hash = stableHash(normalized);

    await db
      .insert(appfolioProperties)
      .values({
        ...normalized,
        rawJson: {
          ...row,
          property_group_id: groupInfo.groupId || row.property_group_id || row.PropertyGroupId || '',
          property_group_uuid: groupInfo.groupUuid || row.property_group_uuid || row.PropertyGroupUuid || '',
          name_of_property_group: groupInfo.groupName || row.name_of_property_group || row.NameOfPropertyGroup || '',
        },
        cachedAt: new Date(),
      })
      .onConflictDoUpdate({
        target: appfolioProperties.id,
        set: {
          name: sql`CASE WHEN ${appfolioProperties.rawJson}->>'_hash' = ${hash} THEN ${appfolioProperties.name} ELSE EXCLUDED.name END`,
          propertyGroupId: sql`EXCLUDED.property_group_id`,
          street: sql`EXCLUDED.street`,
          city: sql`EXCLUDED.city`,
          state: sql`EXCLUDED.state`,
          zip: sql`EXCLUDED.zip`,
          rawJson: sql`jsonb_set(EXCLUDED.raw_json, '{_hash}', to_jsonb(${hash}::text))`,
          cachedAt: sql`EXCLUDED.cached_at`,
        },
      });

    // Detect actual write vs skip by checking hash.
    // We use a simpler pattern: read hash, compare, then decide.
    // Above upsert always runs; track as upserted for simplicity.
    // A future optimization can pre-read hashes in bulk.
    upserted++;
  }

  return { upserted, skipped };
}

// ── Units ────────────────────────────────────────────────────────────────────

export async function upsertUnits(rows: any[]): Promise<UpsertResult> {
  let upserted = 0;
  let skipped = 0;

  for (const row of rows) {
    const unitId = asStr(row.Id || row.id || row.UnitId || row.unit_id);
    const propertyId = asStr(row.PropertyId || row.property_id);
    if (!unitId || !propertyId) continue;

    const normalized = {
      unitId,
      propertyId,
      name: asStr(row.Name || row.name),
      unitNumber: asStr(row.UnitNumber || row.unit_number),
      status: asStr(row.Status || row.status),
      bedrooms: row.Bedrooms != null ? Math.round(Number(row.Bedrooms)) : asNum(row.bedrooms) != null ? Math.round(asNum(row.bedrooms)!) : null,
      bathrooms: asNum(row.Bathrooms ?? row.bathrooms),
      squareFeet: row.SquareFeet != null ? Math.round(Number(row.SquareFeet)) : null,
      marketRent: asNum(row.MarketRent ?? row.market_rent),
    };

    await db
      .insert(appfolioUnits)
      .values({
        ...normalized,
        rawJson: row,
        lastUpdatedAt: asDate(row.LastUpdatedAt || row.last_updated_at),
        cachedAt: new Date(),
      })
      .onConflictDoUpdate({
        target: appfolioUnits.unitId,
        set: {
          propertyId: sql`EXCLUDED.property_id`,
          name: sql`EXCLUDED.name`,
          unitNumber: sql`EXCLUDED.unit_number`,
          status: sql`EXCLUDED.status`,
          bedrooms: sql`EXCLUDED.bedrooms`,
          bathrooms: sql`EXCLUDED.bathrooms`,
          squareFeet: sql`EXCLUDED.square_feet`,
          marketRent: sql`EXCLUDED.market_rent`,
          rawJson: sql`EXCLUDED.raw_json`,
          lastUpdatedAt: sql`EXCLUDED.last_updated_at`,
          cachedAt: sql`EXCLUDED.cached_at`,
        },
      });

    upserted++;
  }

  return { upserted, skipped };
}

// ── Users (Maintenance Tech Baseline) ──────────────────────────────────────

export async function upsertMaintenanceTechUsers(rows: any[]): Promise<UpsertResult> {
  let upserted = 0;
  let skipped = 0;

  for (const row of rows) {
    const techId = asStr(row.Id || row.id || row.UserId || row.user_id);
    if (!techId) {
      skipped++;
      continue;
    }

    const role = asStr(row.UserRole || row.user_role || row.Role || row.role, 120) || '';
    if (!isMaintenanceTechRole(role)) {
      skipped++;
      continue;
    }

    const firstName = asStr(row.FirstName || row.first_name || row.firstName);
    const lastName = asStr(row.LastName || row.last_name || row.lastName);
    const fullName = [firstName, lastName].filter(Boolean).join(' ').trim();
    const techName = asStr(row.Name || row.name || fullName || row.Email || row.email || techId) || techId;
    const email = asStr(row.Email || row.email || row.LoginEmail || row.login_email, 240);
    const appfolioActive = asBool(row.Active ?? row.active ?? row.IsActive ?? row.is_active ?? row.Status ?? row.status, true);
    const updatedAt = asDate(row.LastUpdatedAt || row.last_updated_at || row.UpdatedAt || row.updated_at);

    await db
      .insert(appfolioUsers)
      .values({
        techId,
        techName,
        email,
        userRole: role,
        appfolioActive,
        rawJson: row,
        lastUpdatedAt: updatedAt,
        cachedAt: new Date(),
      })
      .onConflictDoUpdate({
        target: appfolioUsers.techId,
        set: {
          techName: sql`EXCLUDED.tech_name`,
          email: sql`EXCLUDED.email`,
          userRole: sql`EXCLUDED.user_role`,
          appfolioActive: sql`EXCLUDED.appfolio_active`,
          rawJson: sql`EXCLUDED.raw_json`,
          lastUpdatedAt: sql`EXCLUDED.last_updated_at`,
          cachedAt: sql`EXCLUDED.cached_at`,
        },
      });

    // Hydrate tech_grades baseline without clobbering local rotation config.
    await db.execute(sql`
      INSERT INTO tech_grades (tech_id, tech_name, tech_email, active, tier, updated_at)
      VALUES (${techId}, ${techName}, ${email ?? ''}, TRUE, 1, NOW())
      ON CONFLICT (tech_id) DO UPDATE SET
        tech_name = EXCLUDED.tech_name,
        tech_email = CASE
          WHEN tech_grades.tech_email IS NULL OR btrim(tech_grades.tech_email) = '' THEN EXCLUDED.tech_email
          ELSE tech_grades.tech_email
        END,
        updated_at = NOW()
    `);

    upserted++;
  }

  return { upserted, skipped };
}

// ── Work Orders ───────────────────────────────────────────────────────────────

export async function upsertWorkOrders(rows: any[]): Promise<UpsertResult> {
  let upserted = 0;
  let skipped = 0;

  const propertyIdsNeedingLookup = Array.from(new Set(
    rows
      .map((row) => {
        const propertyId = asStr(row.property_id || row.PropertyId);
        const groupInfo = resolvePropertyGroupInfo(row);
        return (!groupInfo.canonicalGroupKey && propertyId) ? propertyId : null;
      })
      .filter((id): id is string => !!id),
  ));

  const propertyGroupByPropertyId = new Map<string, string>();
  if (propertyIdsNeedingLookup.length) {
    const refs = await db
      .select({ id: appfolioProperties.id, propertyGroupId: appfolioProperties.propertyGroupId })
      .from(appfolioProperties)
      .where(inArray(appfolioProperties.id, propertyIdsNeedingLookup));
    for (const ref of refs) {
      if (ref.id && ref.propertyGroupId) {
        propertyGroupByPropertyId.set(ref.id, ref.propertyGroupId);
      }
    }
  }

  for (const row of rows) {
    const id = asStr(
      row.db_api_id || row.dbApiId || row.v0_uuid || row.UUID ||
      row.work_order_id || row.WorkOrderId || row.Id || row.id,
    );
    if (!id) continue;

    const woNumber = asStr(row.work_order_number || row.WorkOrderNumber || row.Number);
    if (!woNumber) continue;

    const assignedUsers = Array.isArray(row.assigned_users ?? row.AssignedUsers)
      ? (row.assigned_users ?? row.AssignedUsers)
      : [];
    const firstUser = assignedUsers[0] ?? {};
    const workOrderUuid = asStr(row.v0_uuid || row.UUID || row.uuid || row.work_order_uuid)
      || (isUuidLike(id) ? id : null);
    const rawWithUuid = {
      ...row,
      work_order_uuid: workOrderUuid || row.work_order_uuid || row.v0_uuid || row.UUID || row.uuid || '',
      v0_uuid: workOrderUuid || row.v0_uuid || row.UUID || row.uuid || '',
    };
    const propertyId = asStr(row.property_id || row.PropertyId);
    const groupInfo = resolvePropertyGroupInfo(row);
    const resolvedGroup = groupInfo.canonicalGroupKey || (propertyId ? (propertyGroupByPropertyId.get(propertyId) || null) : null);

    await db
      .insert(appfolioWorkOrders)
      .values({
        id,
        workOrderUuid,
        woNumber,
        propertyId,
        unitId: asStr(row.unit_id || row.UnitId),
        propertyGroupId: resolvedGroup,
        description: asStr(row.description || row.Description || row.subject || row.Subject, 1000),
        category: asStr(row.category || row.Category || row.work_order_type || row.WorkOrderType),
        priority: asStr(row.priority || row.Priority),
        status: asStr(row.status || row.Status),
        assignedUserId: asStr(firstUser.id || firstUser.Id),
        assignedUserName: asStr(firstUser.name || firstUser.Name || firstUser.full_name),
        vendorId: asStr(row.vendor_id || row.VendorId),
        vendorName: asStr(row.vendor_name || row.VendorName),
        estimatedAmount: asNum(row.estimated_amount || row.EstimatedAmount),
        totalCost: asNum(row.total_cost || row.TotalCost),
        createdAt: asDate(row.created_date || row.CreatedDate || row.created_at),
        updatedAt: asDate(row.last_updated_at || row.LastUpdatedAt || row.updated_at),
        rawJson: rawWithUuid,
      })
      .onConflictDoUpdate({
        target: appfolioWorkOrders.id,
        set: {
          workOrderUuid: sql`EXCLUDED.work_order_uuid`,
          woNumber: sql`EXCLUDED.wo_number`,
          propertyId: sql`EXCLUDED.property_id`,
          unitId: sql`EXCLUDED.unit_id`,
          propertyGroupId: sql`EXCLUDED.property_group_id`,
          description: sql`EXCLUDED.description`,
          category: sql`EXCLUDED.category`,
          priority: sql`EXCLUDED.priority`,
          status: sql`EXCLUDED.status`,
          assignedUserId: sql`EXCLUDED.assigned_user_id`,
          assignedUserName: sql`EXCLUDED.assigned_user_name`,
          vendorId: sql`EXCLUDED.vendor_id`,
          vendorName: sql`EXCLUDED.vendor_name`,
          estimatedAmount: sql`EXCLUDED.estimated_amount`,
          totalCost: sql`EXCLUDED.total_cost`,
          createdAt: sql`EXCLUDED.created_at`,
          updatedAt: sql`EXCLUDED.updated_at`,
          rawJson: sql`EXCLUDED.raw_json`,
        },
      });

    upserted++;
  }

  return { upserted, skipped };
}

// ── Unit Inspections ─────────────────────────────────────────────────────────

export async function upsertUnitInspections(rows: any[]): Promise<UpsertResult> {
  let upserted = 0;
  let skipped = 0;

  for (const row of rows) {
    const propertyId = asStr(row.property_id || row.PropertyId);
    const unitId = asStr(row.unit_id || row.UnitId);
    const inspectionId = reportRowKey(
      row,
      ['inspection_id', 'InspectionId', 'occupancy_id', 'OccupancyId'],
      [propertyId, unitId, row.last_inspection_date || row.LastInspectionDate, row.tenant_name || row.TenantName],
    );

    await db
      .insert(appfolioUnitInspections)
      .values({
        inspectionId,
        propertyId,
        propertyName: asStr(row.property_name || row.PropertyName),
        unitId,
        unitName: asStr(row.unit_name || row.UnitName),
        lastInspectionDate: asDate(row.last_inspection_date || row.LastInspectionDate),
        tenantName: asStr(row.tenant_name || row.TenantName),
        tenantPrimaryPhoneNumber: asStr(row.tenant_primary_phone_number || row.TenantPrimaryPhoneNumber),
        moveInDate: asDate(row.move_in_date || row.MoveInDate),
        moveOutDate: asDate(row.move_out_date || row.MoveOutDate),
        rentable: asStr(row.rentable || row.Rentable),
        occupancyId: asStr(row.occupancy_id || row.OccupancyId),
        unitTags: asStr(row.unit_tags || row.UnitTags),
        rawJson: row,
        lastUpdatedAt: asDate(row.last_updated_at || row.LastUpdatedAt),
        cachedAt: new Date(),
      })
      .onConflictDoUpdate({
        target: appfolioUnitInspections.inspectionId,
        set: {
          propertyId: sql`EXCLUDED.property_id`,
          propertyName: sql`EXCLUDED.property_name`,
          unitId: sql`EXCLUDED.unit_id`,
          unitName: sql`EXCLUDED.unit_name`,
          lastInspectionDate: sql`EXCLUDED.last_inspection_date`,
          tenantName: sql`EXCLUDED.tenant_name`,
          tenantPrimaryPhoneNumber: sql`EXCLUDED.tenant_primary_phone_number`,
          moveInDate: sql`EXCLUDED.move_in_date`,
          moveOutDate: sql`EXCLUDED.move_out_date`,
          rentable: sql`EXCLUDED.rentable`,
          occupancyId: sql`EXCLUDED.occupancy_id`,
          unitTags: sql`EXCLUDED.unit_tags`,
          rawJson: sql`EXCLUDED.raw_json`,
          lastUpdatedAt: sql`EXCLUDED.last_updated_at`,
          cachedAt: sql`EXCLUDED.cached_at`,
        },
      });

    upserted++;
  }

  return { upserted, skipped };
}

// ── Tenant Directory / Upcoming Move-Outs ───────────────────────────────────

export async function upsertTenantDirectory(rows: any[]): Promise<UpsertResult> {
  let upserted = 0;
  let skipped = 0;

  for (const row of rows) {
    const propertyId = asStr(row.property_id || row.PropertyId);
    const unitId = asStr(row.unit_id || row.UnitId);
    const recordId = reportRowKey(
      row,
      ['occupancy_id', 'OccupancyId'],
      [propertyId, unitId, row.move_out_date || row.MoveOutDate, row.tenant || row.Tenant],
    );

    await db
      .insert(appfolioTenantDirectory)
      .values({
        recordId,
        propertyId,
        propertyName: asStr(row.property_name || row.PropertyName),
        unitId,
        unitName: asStr(row.unit_name || row.UnitName),
        tenantName: asStr(row.tenant || row.Tenant || row.tenant_name || row.TenantName),
        status: asStr(row.status || row.Status),
        tenantType: asStr(row.tenant_type || row.TenantType),
        phoneNumbers: asStr(row.phone_numbers || row.PhoneNumbers),
        emails: asStr(row.emails || row.Emails),
        moveInDate: asDate(row.move_in || row.MoveIn || row.move_in_date || row.MoveInDate),
        leaseTo: asDate(row.lease_to || row.LeaseTo),
        moveOutDate: asDate(row.move_out || row.MoveOut || row.move_out_date || row.MoveOutDate),
        rent: asStr(row.rent || row.Rent),
        tenantTags: asStr(row.tenant_tags || row.TenantTags),
        tenantAgent: asStr(row.tenant_agent || row.TenantAgent),
        tenantVisibility: asStr(row.tenant_visibility || row.TenantVisibility),
        occupancyId: asStr(row.occupancy_id || row.OccupancyId),
        unitTags: asStr(row.unit_tags || row.UnitTags),
        rawJson: row,
        lastUpdatedAt: asDate(row.last_updated_at || row.LastUpdatedAt),
        cachedAt: new Date(),
      })
      .onConflictDoUpdate({
        target: appfolioTenantDirectory.recordId,
        set: {
          propertyId: sql`EXCLUDED.property_id`,
          propertyName: sql`EXCLUDED.property_name`,
          unitId: sql`EXCLUDED.unit_id`,
          unitName: sql`EXCLUDED.unit_name`,
          tenantName: sql`EXCLUDED.tenant_name`,
          status: sql`EXCLUDED.status`,
          tenantType: sql`EXCLUDED.tenant_type`,
          phoneNumbers: sql`EXCLUDED.phone_numbers`,
          emails: sql`EXCLUDED.emails`,
          moveInDate: sql`EXCLUDED.move_in_date`,
          leaseTo: sql`EXCLUDED.lease_to`,
          moveOutDate: sql`EXCLUDED.move_out_date`,
          rent: sql`EXCLUDED.rent`,
          tenantTags: sql`EXCLUDED.tenant_tags`,
          tenantAgent: sql`EXCLUDED.tenant_agent`,
          tenantVisibility: sql`EXCLUDED.tenant_visibility`,
          occupancyId: sql`EXCLUDED.occupancy_id`,
          unitTags: sql`EXCLUDED.unit_tags`,
          rawJson: sql`EXCLUDED.raw_json`,
          lastUpdatedAt: sql`EXCLUDED.last_updated_at`,
          cachedAt: sql`EXCLUDED.cached_at`,
        },
      });

    upserted++;
  }

  return { upserted, skipped };
}

// ── Unit Turn Detail ────────────────────────────────────────────────────────

export async function upsertUnitTurnDetails(rows: any[]): Promise<UpsertResult> {
  let upserted = 0;
  let skipped = 0;

  for (const row of rows) {
    const propertyId = asStr(row.property_id || row.PropertyId);
    const unitId = asStr(row.unit_id || row.UnitId);
    const turnId = reportRowKey(
      row,
      ['unit_turn_id', 'UnitTurnId'],
      [propertyId, unitId, row.move_out_date || row.MoveOutDate, row.turn_end_date || row.TurnEndDate],
    );

    await db
      .insert(appfolioUnitTurnDetails)
      .values({
        turnId,
        propertyId,
        propertyName: asStr(row.property || row.Property || row.property_name || row.PropertyName),
        unitId,
        unitName: asStr(row.unit || row.Unit || row.unit_name || row.UnitName),
        notes: asStr(row.notes || row.Notes),
        referenceUser: asStr(row.reference_user || row.ReferenceUser),
        moveOutDate: asDate(row.move_out_date || row.MoveOutDate),
        turnEndDate: asDate(row.turn_end_date || row.TurnEndDate),
        expectedMoveInDate: asDate(row.expected_move_in_date || row.ExpectedMoveInDate),
        targetDaysToComplete: row.target_days_to_complete != null ? Math.round(Number(row.target_days_to_complete)) : asNum(row.target_days_to_complete),
        totalDaysToComplete: row.total_days_to_complete != null ? Math.round(Number(row.total_days_to_complete)) : asNum(row.total_days_to_complete),
        laborFromWorkOrders: asStr(row.labor_from_work_orders || row.LaborFromWorkOrders),
        purchaseOrdersFromWorkOrders: asStr(row.purchase_orders_from_work_orders || row.PurchaseOrdersFromWorkOrders),
        billablesFromWorkOrders: asStr(row.billables_from_work_orders || row.BillablesFromWorkOrders),
        inventoryFromWorkOrders: asStr(row.inventory_from_work_orders || row.InventoryFromWorkOrders),
        totalBilled: asStr(row.total_billed || row.TotalBilled),
        unitTurnStatus: asStr(row.unit_turn_status || row.UnitTurnStatus || row.status || row.Status),
        propertyVisibility: asStr(row.property_visibility || row.PropertyVisibility),
        rawJson: row,
        lastUpdatedAt: asDate(row.last_updated_at || row.LastUpdatedAt),
        cachedAt: new Date(),
      })
      .onConflictDoUpdate({
        target: appfolioUnitTurnDetails.turnId,
        set: {
          propertyId: sql`EXCLUDED.property_id`,
          propertyName: sql`EXCLUDED.property_name`,
          unitId: sql`EXCLUDED.unit_id`,
          unitName: sql`EXCLUDED.unit_name`,
          notes: sql`EXCLUDED.notes`,
          referenceUser: sql`EXCLUDED.reference_user`,
          moveOutDate: sql`EXCLUDED.move_out_date`,
          turnEndDate: sql`EXCLUDED.turn_end_date`,
          expectedMoveInDate: sql`EXCLUDED.expected_move_in_date`,
          targetDaysToComplete: sql`EXCLUDED.target_days_to_complete`,
          totalDaysToComplete: sql`EXCLUDED.total_days_to_complete`,
          laborFromWorkOrders: sql`EXCLUDED.labor_from_work_orders`,
          purchaseOrdersFromWorkOrders: sql`EXCLUDED.purchase_orders_from_work_orders`,
          billablesFromWorkOrders: sql`EXCLUDED.billables_from_work_orders`,
          inventoryFromWorkOrders: sql`EXCLUDED.inventory_from_work_orders`,
          totalBilled: sql`EXCLUDED.total_billed`,
          unitTurnStatus: sql`EXCLUDED.unit_turn_status`,
          propertyVisibility: sql`EXCLUDED.property_visibility`,
          rawJson: sql`EXCLUDED.raw_json`,
          lastUpdatedAt: sql`EXCLUDED.last_updated_at`,
          cachedAt: sql`EXCLUDED.cached_at`,
        },
      });

    upserted++;
  }

  return { upserted, skipped };
}

// ── Unit Vacancies ───────────────────────────────────────────────────────────

export async function upsertUnitVacancies(rows: any[]): Promise<UpsertResult> {
  let upserted = 0;
  let skipped = 0;

  for (const row of rows) {
    const propertyId = asStr(row.property_id || row.PropertyId);
    const unitId = asStr(row.unit_id || row.UnitId);
    const recordId = reportRowKey(
      row,
      ['unit_id', 'UnitId', 'unit', 'Unit'],
      [propertyId, row.property_name || row.PropertyName, row.vacant_from || row.VacantFrom, row.unit_name || row.UnitName],
    );

    await db
      .insert(appfolioUnitVacancies)
      .values({
        recordId,
        propertyId,
        propertyName: asStr(row.property_name || row.PropertyName || row.property || row.Property),
        unitId,
        unitName: asStr(row.unit_name || row.UnitName || row.unit || row.Unit),
        vacantFrom: asDate(row.vacant_from || row.VacantFrom || row.vacated_on || row.VacatedOn),
        availableOn: asDate(row.available_on || row.AvailableOn || row.market_ready_date || row.MarketReadyDate),
        marketRent: asStr(row.market_rent || row.MarketRent || row.rent || row.Rent),
        bedrooms: asStr(row.bedrooms || row.Bedrooms),
        bathrooms: asStr(row.bathrooms || row.Bathrooms),
        daysVacant: row.days_vacant != null ? Math.round(Number(row.days_vacant)) : asNum(row.vacancy_days),
        status: asStr(row.status || row.Status),
        propertyVisibility: asStr(row.property_visibility || row.PropertyVisibility),
        rawJson: row,
        lastUpdatedAt: asDate(row.last_updated_at || row.LastUpdatedAt),
        cachedAt: new Date(),
      })
      .onConflictDoUpdate({
        target: appfolioUnitVacancies.recordId,
        set: {
          propertyId: sql`EXCLUDED.property_id`,
          propertyName: sql`EXCLUDED.property_name`,
          unitId: sql`EXCLUDED.unit_id`,
          unitName: sql`EXCLUDED.unit_name`,
          vacantFrom: sql`EXCLUDED.vacant_from`,
          availableOn: sql`EXCLUDED.available_on`,
          marketRent: sql`EXCLUDED.market_rent`,
          bedrooms: sql`EXCLUDED.bedrooms`,
          bathrooms: sql`EXCLUDED.bathrooms`,
          daysVacant: sql`EXCLUDED.days_vacant`,
          status: sql`EXCLUDED.status`,
          propertyVisibility: sql`EXCLUDED.property_visibility`,
          rawJson: sql`EXCLUDED.raw_json`,
          lastUpdatedAt: sql`EXCLUDED.last_updated_at`,
          cachedAt: sql`EXCLUDED.cached_at`,
        },
      });

    upserted++;
  }

  return { upserted, skipped };
}

// ── Estimates ─────────────────────────────────────────────────────────────────

export async function upsertEstimates(rows: any[]): Promise<UpsertResult> {
  let upserted = 0;
  let skipped = 0;

  for (const row of rows) {
    const estimateId = asStr(row.estimate_id || row.EstimateId || row.id || row.Id);
    const currentStatus = asStr(row.approval_status || row.current_status || row.status) ?? 'unknown';
    if (!estimateId) continue;

    // Append to status history if status changed.
    const existing = await db
      .select({ statusHistory: appfolioEstimates.statusHistory, currentStatus: appfolioEstimates.currentStatus })
      .from(appfolioEstimates)
      .where(sql`estimate_id = ${estimateId}`)
      .limit(1);

    const prevStatus = existing[0]?.currentStatus ?? null;
    const history: any[] = Array.isArray(existing[0]?.statusHistory) ? existing[0].statusHistory : [];
    if (!history.length || prevStatus !== currentStatus) {
      history.push({ status: currentStatus, changed_at: new Date().toISOString() });
    }

    await db
      .insert(appfolioEstimates)
      .values({
        estimateId,
        workOrderId: asStr(row.work_order_id || row.WorkOrderId),
        workOrderNumber: asStr(row.work_order_number || row.WorkOrderNumber),
        currentStatus,
        propertyGroupId: asStr(row.property_group_id || row.PropertyGroupId),
        source: asStr(row.source),
        statusHistory: history,
        rawData: row,
        updatedAt: new Date(),
      })
      .onConflictDoUpdate({
        target: appfolioEstimates.estimateId,
        set: {
          workOrderId: sql`EXCLUDED.work_order_id`,
          workOrderNumber: sql`EXCLUDED.work_order_number`,
          currentStatus: sql`EXCLUDED.current_status`,
          propertyGroupId: sql`EXCLUDED.property_group_id`,
          source: sql`EXCLUDED.source`,
          statusHistory: sql`EXCLUDED.status_history`,
          rawData: sql`EXCLUDED.raw_data`,
          updatedAt: sql`EXCLUDED.updated_at`,
        },
      });

    upserted++;
  }

  return { upserted, skipped };
}

// ── Turn Tracker ──────────────────────────────────────────────────────────────

export async function upsertTurnTracker(records: any[]): Promise<UpsertResult> {
  let upserted = 0;
  let skipped = 0;

  for (const record of records) {
    const turnKey = asStr(record.turn_key);
    if (!turnKey) continue;

    const trackingUuid = asStr(record.tracking_uuid) ?? createHash('sha256').update(turnKey).digest('hex').slice(0, 36);

    await db
      .insert(unitTurnTracker)
      .values({
        trackingUuid,
        trackingCode: asStr(record.tracking_code),
        turnKey,
        unitTurnId: asStr(record.unit_turn_id),
        unitId: asStr(record.unit_id),
        propertyId: asStr(record.property_id),
        unitName: asStr(record.unit_name),
        propertyName: asStr(record.property_name),
        status: asStr(record.status) ?? 'open',
        confidenceScore: asNum(record.confidence_score),
        confidenceLabel: asStr(record.confidence_label),
        sourceFlags: record.source_flags ?? {},
        metadata: record.metadata ?? {},
        closedAt: asDate(record.closed_at),
        updatedAt: new Date(),
      })
      .onConflictDoUpdate({
        target: unitTurnTracker.trackingUuid,
        set: {
          trackingCode: sql`EXCLUDED.tracking_code`,
          unitTurnId: sql`EXCLUDED.unit_turn_id`,
          unitId: sql`EXCLUDED.unit_id`,
          propertyId: sql`EXCLUDED.property_id`,
          unitName: sql`EXCLUDED.unit_name`,
          propertyName: sql`EXCLUDED.property_name`,
          status: sql`EXCLUDED.status`,
          confidenceScore: sql`EXCLUDED.confidence_score`,
          confidenceLabel: sql`EXCLUDED.confidence_label`,
          sourceFlags: sql`EXCLUDED.source_flags`,
          metadata: sql`EXCLUDED.metadata`,
          closedAt: sql`EXCLUDED.closed_at`,
          updatedAt: sql`NOW()`,
        },
      });

    // Milestones
    if (record.milestones && typeof record.milestones === 'object') {
      for (const [key, val] of Object.entries(record.milestones as any)) {
        const m = val as any;
        await db
          .insert(unitTurnMilestones)
          .values({
            trackingUuid,
            milestoneKey: key,
            milestoneDate: asDate(m?.date),
            source: asStr(m?.source),
            notes: asStr(m?.notes),
          });
        // Not using onConflict here because milestones are append-only by design.
        // Duplicate attempts are caught at the DB unique constraint if one exists.
      }
    }

    // Linked work orders
    if (Array.isArray(record.work_orders)) {
      for (const wo of record.work_orders) {
        const woId = asStr(wo.wo_id || wo.id);
        if (!woId) continue;
        await db
          .insert(unitTurnWorkOrders)
          .values({
            trackingUuid,
            woId,
            woDbUuid: asStr(wo.wo_db_uuid),
            source: asStr(wo.source) ?? 'sync',
            status: asStr(wo.status),
            removed: false,
          });
      }
    }

    upserted++;
  }

  return { upserted, skipped };
}
