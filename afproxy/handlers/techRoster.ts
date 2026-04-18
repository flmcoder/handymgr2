// ============================================================================
// handlers/techRoster.ts — Technician roster CRUD.
//
// The tech_grades table is the authoritative source for the reassignment
// engine's technician roster. Every tech that should participate in
// automatic work order distribution must be added here via this endpoint.
//
// CRITICAL before enabling midnight cron in production:
//   Every tech_id in this roster must be a valid AppFolio user UUID with
//   the Maintenance Tech role enabled in AppFolio Property Manager.
//   Without this role the PATCH AssignedUsers call will return 422 .
//   Verify UUIDs via: GET ?action=passthrough&path=/api/v0/users
//
//   In order to successfully carry out certain requests using the API, the
//   correct user roles must be enabled in AppFolio Property Manager .
//
// GET  ?action=tech_roster
//   Returns all techs sorted by tier then by weighted work load.
//
// POST ?action=tech_roster
//   Body: { tech_id, tech_name, tech_phone, tier?, geo_zone?, active? }
//   Upserts the tech record. Use tier=2 for Deep Bench pool members.
//   geo_zone values: "central" | "north" | "east" | "west" (freeform string)
//
// DELETE-style deactivation:
//   POST with { tech_id, active: 0 } to remove a tech from the active pool
//   without deleting their historical grade data.
// ============================================================================

import { FLR_GROUPS } from "../config.ts";
import { rowsAsObjects, sqlite } from "../db.ts";
import { fetchDbApi } from "../lib/appfolio.ts";
import { handlePropertyGroups } from "./properties.ts";
import { handleWorkOrders } from "./workOrders.ts";

type BranchKey = "phoenix" | "tucson";

function normalizeBranch(value: string): BranchKey {
  const z = String(value || "").toLowerCase();
  return z.includes("tucson") ? "tucson" : "phoenix";
}

function safeJsonParse(raw: string): any {
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function normalizeUuid(v: unknown): string {
  return String(v || "").trim().toLowerCase();
}

function getAssigneeCandidates(rows: any[]): Array<{
  tech_id: string;
  tech_name: string;
  branch: BranchKey;
  count: number;
}> {
  const out: Record<string, {
    tech_id: string;
    tech_name: string;
    branch: BranchKey;
    count: number;
  }> = {};

  for (const r of rows || []) {
    const branch = normalizeBranch(r._branch || "phoenix");

    const push = (id: unknown, name: unknown) => {
      const techId = String(id || "").trim();
      const techName = String(name || "").trim();
      if (!techId || !techName) return;

      if (!out[techId]) {
        out[techId] = {
          tech_id: techId,
          tech_name: techName,
          branch,
          count: 0,
        };
      }
      out[techId].count += 1;
    };

    push(
      r.assigned_to_id || r.AssignedToId || r.assigned_user_id ||
        r.AssignedUserId || r.assigned_user_uuid || r.AssignedUserUUID,
      r.assigned_to || r.AssignedTo || r.assigned_user || r.AssignedUser ||
        r.assigned_user_name || r.AssignedUserName,
    );

    const assignedUsers = r.AssignedUsers || r.assigned_users || [];
    if (Array.isArray(assignedUsers)) {
      for (const u of assignedUsers) {
        if (!u || typeof u !== "object") continue;
        push(u.Id || u.id || u.UserId || u.user_id, u.Name || u.name);
      }
    }
  }

  return Object.values(out).sort((a, b) => b.count - a.count);
}

async function fetchAppfolioUsers(): Promise<any[]> {
  try {
    return await fetchDbApi(
      "/api/v0/users?filters[LastUpdatedAtFrom]=2024-01-01T00:00:00Z&page[size]=200",
      1000,
    );
  } catch (e: any) {
    console.log(`dispatch_sync_assignees users fetch failed: ${e.message}`);
    return [];
  }
}

function mergeUserRoster(
  userRows: any[],
  activityRows: Array<{
    tech_id: string;
    tech_name: string;
    branch: BranchKey;
    count: number;
  }>,
  selectedBranch: string,
): Array<{
  tech_id: string;
  tech_name: string;
  branch: BranchKey;
  count: number;
  email?: string;
  user_role?: string;
}> {
  const activityById = new Map(
    activityRows.map((row) => [String(row.tech_id).trim(), row]),
  );
  const out = new Map<string, {
    tech_id: string;
    tech_name: string;
    branch: BranchKey;
    count: number;
    email?: string;
    user_role?: string;
  }>();

  for (const user of userRows || []) {
    const techId = String(user.Id || user.id || "").trim();
    if (!techId) continue;

    const activity = activityById.get(techId);
    const role = String(user.UserRole || user.user_role || "").trim();
    const roleNorm = role.toLowerCase();
    const looksLikeMaintenance = roleNorm.includes("maintenance") ||
      roleNorm.includes("tech");

    if (!activity && !looksLikeMaintenance) continue;
    if (!activity && selectedBranch === "all") continue;

    const branch = activity?.branch || normalizeBranch(selectedBranch || "phoenix");
    const techName = [user.FirstName, user.LastName]
      .map((v) => String(v || "").trim())
      .filter(Boolean)
      .join(" ") || activity?.tech_name || String(user.Email || techId);

    out.set(techId, {
      tech_id: techId,
      tech_name: techName,
      branch,
      count: activity?.count || 0,
      email: String(user.Email || ""),
      user_role: role,
    });
  }

  for (const activity of activityRows) {
    if (selectedBranch !== "all" && activity.branch !== selectedBranch) continue;
    if (!out.has(activity.tech_id)) {
      out.set(activity.tech_id, {
        tech_id: activity.tech_id,
        tech_name: activity.tech_name,
        branch: activity.branch,
        count: activity.count,
      });
    }
  }

  return Array.from(out.values()).sort((a, b) => {
    if (b.count !== a.count) return b.count - a.count;
    return a.tech_name.localeCompare(b.tech_name);
  });
}

async function getGroupPropertyUuidMap(
  tier1GroupUuid: string,
  tier2GroupUuid: string,
): Promise<{ tier1: Set<string>; tier2: Set<string> }> {
  const groupsResp = await handlePropertyGroups({
    "filters[LastUpdatedAtFrom]": "2024-01-01T00:00:00Z",
    "page[size]": "200",
  });

  const all = Array.isArray(groupsResp?.results) ? groupsResp.results : [];
  const t1 = new Set<string>();
  const t2 = new Set<string>();
  const g1 = normalizeUuid(tier1GroupUuid || FLR_GROUPS.Phoenix);
  const g2 = normalizeUuid(tier2GroupUuid || FLR_GROUPS.Tucson);

  for (const g of all) {
    const gid = normalizeUuid(g?.Id || g?.id);
    if (!gid) continue;
    const raw = g?.PropertyIds || g?.properties || g?.Properties || [];
    const uuids = Array.isArray(raw)
      ? raw.map((x: any) => normalizeUuid(x?.Id || x?.id || x)).filter(Boolean)
      : [];

    if (gid === g1) uuids.forEach((u) => t1.add(u));
    if (gid === g2) uuids.forEach((u) => t2.add(u));
  }

  return { tier1: t1, tier2: t2 };
}

export async function handleDispatchSyncAssignees(
  params: Record<string, string>,
  req?: Request,
): Promise<any> {
  let body: any = {};
  if (req?.method === "POST") {
    body = safeJsonParse(await req.text()) || {};
  }

  const tier1GroupUuid = String(
    body.tier1_group_uuid || params.tier1_group_uuid || FLR_GROUPS.Phoenix,
  );
  const tier2GroupUuid = String(
    body.tier2_group_uuid || params.tier2_group_uuid || FLR_GROUPS.Tucson,
  );
  const selectedBranch = String(body.branch || params.branch || "all")
    .toLowerCase();

  const wo = await handleWorkOrders({ days: String(params.days || "90") });
  if (!wo.ok) return { ok: false, error: "Unable to fetch work orders" };

  const { tier1, tier2 } = await getGroupPropertyUuidMap(
    tier1GroupUuid,
    tier2GroupUuid,
  );

  const tagged = (wo.results || []).map((r: any) => {
    const pid = normalizeUuid(r.property_id || r.PropertyId || r.property_uuid);
    let branch: BranchKey = "phoenix";
    if (pid && tier2.has(pid)) branch = "tucson";
    else if (pid && tier1.has(pid)) branch = "phoenix";
    return { ...r, _branch: branch };
  }).filter((r: any) => {
    if (selectedBranch === "all") return true;
    return r._branch === selectedBranch;
  });

  const activityCandidates = getAssigneeCandidates(tagged);
  const appfolioUsers = await fetchAppfolioUsers();
  const assignees = mergeUserRoster(
    appfolioUsers,
    activityCandidates,
    selectedBranch,
  );
  let synced = 0;
  let tier1Count = 0;
  let tier2Count = 0;

  for (const a of assignees) {
    const tier = a.branch === "phoenix" ? 1 : 2;
    if (tier === 1) tier1Count++;
    else tier2Count++;

    await sqlite.execute({
      sql: `INSERT INTO tech_grades
             (tech_id, tech_name, tier, geo_zone, active, updated_at)
           VALUES (?, ?, ?, ?, 1, datetime('now'))
           ON CONFLICT(tech_id) DO UPDATE SET
             tech_name = excluded.tech_name,
             tier = excluded.tier,
             geo_zone = excluded.geo_zone,
             active = 1,
             updated_at = datetime('now')`,
      args: [a.tech_id, a.tech_name, tier, a.branch],
    });
    synced++;
  }

  return {
    ok: true,
    synced,
    tier1: tier1Count,
    tier2: tier2Count,
    techs: assignees.map((a) => ({
      tech_id: a.tech_id,
      tech_name: a.tech_name,
      tier: a.branch === "phoenix" ? 1 : 2,
      geo_zone: a.branch,
      active: 1,
      source_role: a.user_role || "",
      email: a.email || "",
    })),
    source: appfolioUsers.length > 0
      ? "appfolio_users_with_work_order_branch_map"
      : "work_orders_assignee_fallback",
    appfolio_user_count: appfolioUsers.length,
  };
}

export async function handleDispatchSeedReassignmentTest(
  params: Record<string, string>,
  req?: Request,
): Promise<any> {
  let body: any = {};
  if (req?.method === "POST") {
    body = safeJsonParse(await req.text()) || {};
  }

  const limit = Math.max(
    1,
    Math.min(200, parseInt(String(body.limit || params.limit || "50"), 10) || 50),
  );
  const branch = String(body.branch || params.branch || "all").toLowerCase();
  const wo = await handleWorkOrders({ days: String(params.days || "90") });
  if (!wo.ok) return { ok: false, error: "Unable to fetch work orders" };

  const { tier1, tier2 } = await getGroupPropertyUuidMap(
    String(body.tier1_group_uuid || params.tier1_group_uuid || FLR_GROUPS.Phoenix),
    String(body.tier2_group_uuid || params.tier2_group_uuid || FLR_GROUPS.Tucson),
  );

  let inserted = 0;
  for (const r of wo.results || []) {
    if (inserted >= limit) break;

    const woId = String(r.work_order_id || r.Id || r.id || "").trim();
    if (!woId) continue;

    const pid = normalizeUuid(r.property_id || r.PropertyId || r.property_uuid);
    const rowBranch: BranchKey = pid && tier2.has(pid) ? "tucson" : "phoenix";
    if (branch !== "all" && rowBranch !== branch) continue;

    const status = String(r.work_order_status || r.status || r.Status || "").toLowerCase();
    if (!["assigned", "scheduled", "waiting", "in progress", "new", "estimated", "estimate requested"].some((s) => status.includes(s))) {
      continue;
    }

    await sqlite.execute({
      sql: `INSERT INTO reassignment_queue
             (wo_id, wo_number, property_address, assigned_tech_id, assigned_tech_name,
              wo_status, wo_priority, wo_category, first_seen_at, warning_sent, grace_used,
              reassignment_count, escalated)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now','-49 hours'), 0, 0, 0, 0)
           ON CONFLICT(wo_id) DO UPDATE SET
             wo_number = excluded.wo_number,
             property_address = excluded.property_address,
             assigned_tech_id = excluded.assigned_tech_id,
             assigned_tech_name = excluded.assigned_tech_name,
             wo_status = excluded.wo_status,
             wo_priority = excluded.wo_priority,
             wo_category = excluded.wo_category`,
      args: [
        woId,
        String(r.work_order_number || r.Number || woId),
        String(r.property_address || r.property || r.property_name || ""),
        String(r.assigned_to_id || r.AssignedToId || ""),
        String(r.assigned_to || r.AssignedTo || "Unassigned"),
        String(r.work_order_status || r.status || r.Status || ""),
        String(r.priority || r.Priority || "Normal"),
        String(r.category || r.Category || ""),
      ],
    });
    inserted++;
  }

  return {
    ok: true,
    inserted,
    mode: "test",
    branch,
    limit,
  };
}

// ── handleTechRoster ──────────────────────────────────────────────────────────
export async function handleTechRoster(
  params: Record<string, string>,
  req?: Request,
): Promise<any> {
  // ── POST: upsert a tech ───────────────────────────────────────────────────
  if (req?.method === "POST") {
    let body: any = {};
    try {
      body = await req.json();
    } catch {
      return { ok: false, error: "Invalid JSON body" };
    }

    const {
      tech_id,
      tech_name,
      tech_phone = "",
      tier = 1,
      geo_zone = "phoenix",
      active = 1,
    } = body;

    if (!tech_id || !tech_name) {
      return {
        ok: false,
        error: "Missing required fields: tech_id, tech_name",
      };
    }

    // Phone number format advisory — RC requires E.164 e.g. "+15205551234"
    // In general, ensure the formatting of the authorization parameters passed
    // to AppFolio are valid — same principle applies to phone format .
    if (tech_phone && !String(tech_phone).startsWith("+")) {
      console.log(
        `techRoster: WARNING — tech_phone "${tech_phone}" does not appear to be E.164 format. ` +
          `RingCentral requires +1XXXXXXXXXX format for US numbers.`,
      );
    }

    try {
      await sqlite.execute({
        sql: `INSERT INTO tech_grades
                 (tech_id, tech_name, tech_phone, tier, geo_zone, active, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
               ON CONFLICT(tech_id) DO UPDATE SET
                 tech_name  = excluded.tech_name,
                 tech_phone = excluded.tech_phone,
                 tier       = excluded.tier,
                 geo_zone   = excluded.geo_zone,
                 active     = excluded.active,
                 updated_at = datetime('now')`,
        args: [
          String(tech_id),
          String(tech_name),
          String(tech_phone),
          Number(tier),
          normalizeBranch(String(geo_zone)),
          Number(active),
        ],
      });

      return {
        ok: true,
        message: `Tech "${tech_name}" saved to roster`,
        detail: {
          tech_id,
          tech_name,
          tech_phone: tech_phone || "(none)",
          tier: Number(tier),
          geo_zone: String(geo_zone),
          active: Number(active),
        },
      };
    } catch (e: any) {
      return { ok: false, error: e.message };
    }
  }

  // ── GET: return full roster ───────────────────────────────────────────────
  try {
    const rows = rowsAsObjects(
      await sqlite.execute(
        `SELECT
           tech_id,
           tech_name,
           tech_phone,
           tier,
           geo_zone,
           active,
           active_wo_count,
           total_assigned,
           total_completed,
           total_go_backs,
           total_auto_reassigned,
           performance_score,
           load_weight,
           target_share_pct,
           last_assigned_at,
           score_updated_at,
           updated_at
         FROM tech_grades
         ORDER BY
           tier             ASC,
           active           DESC,
           (active_wo_count * COALESCE(load_weight, 1.0)) ASC`,
      ),
    );

    // Compute a brief status label for each tech
    const annotated = rows.map((t: any) => ({
      ...t,
      status_label: !t.active
        ? "Inactive"
        : t.tier === 2
        ? "Tier 2 (Deep Bench)"
        : t.performance_score >= 80
        ? "🟢 Strong"
        : t.performance_score >= 60
        ? "🟡 Monitor"
        : "🔴 At Risk",
    }));

    return {
      ok: true,
      techs: annotated,
      count: annotated.length,
      summary: {
        tier1_active: rows.filter((t: any) => t.tier === 1 && t.active).length,
        tier2_active: rows.filter((t: any) => t.tier === 2 && t.active).length,
        inactive: rows.filter((t: any) => !t.active).length,
      },
    };
  } catch (e: any) {
    return { ok: false, error: e.message };
  }
}