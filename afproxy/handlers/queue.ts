// ============================================================================
// handlers/queue.ts — Reassignment queue viewer + audit trail + blast monitor.
//
// Exports:
//   handleReassignmentQueue — full queue state with stats, audit log, blasts
//
// GET ?action=reassignment_queue
//        &limit=50      (default 50 — controls queue + audit rows returned)
//        &wo_id=UUID    (optional — filter to one specific work order)
//
// Returns:
//   stats        — aggregate counts (warned, exempt, escalated, grace active)
//   queue        — reassignment_queue rows ordered by first_seen_at DESC
//   audit        — last 200 wo_audit_log entries ordered by created_at DESC
//   blasts       — last 20 blast_events ordered by blasted_at DESC
//   tier2_claims — last 50 tier2_claims for open blasts
//
// This endpoint powers the Dispatch Control tab in the HandyManager cockpit.
// It is read-only — all writes flow through the cron handlers and the
// tech roster endpoint.
//
// The Webhook Logs page in AppFolio provides complementary visibility into
// notification delivery . The See All Events link displays all
// webhook events available for debugging errant requests .
//
// In order to successfully carry out certain API requests, the correct user
// roles must be enabled in AppFolio Property Manager . The queue
// viewer does not make any AppFolio API calls — it reads only from Turso.
// ============================================================================

import { rowsAsObjects, sqlite } from "../db.ts";
import * as techRosterHandlers from "./techRoster.ts";

function normalizeBranch(value: string): "all" | "phoenix" | "tucson" {
  const v = String(value || "all").toLowerCase();
  if (v === "phoenix" || v === "tucson" || v === "all") return v;
  return "all";
}

function parseHiddenAssigneeMap(raw: string): Record<string, boolean> {
  try {
    const obj = JSON.parse(raw || "{}");
    if (obj && typeof obj === "object") return obj;
  } catch {
    // ignore malformed config
  }
  return {};
}

function isTechHidden(
  hiddenMap: Record<string, boolean>,
  techId: unknown,
): boolean {
  const tid = String(techId || "").trim();
  return !!(tid && hiddenMap[tid]);
}

function techBranch(v: any): "phoenix" | "tucson" | "unknown" {
  const z = String(v?.geo_zone || "").toLowerCase();
  if (z.includes("tucson")) return "tucson";
  if (z.includes("phoenix") || z === "central" || z === "") return "phoenix";
  return "unknown";
}

// ── handleReassignmentQueue ───────────────────────────────────────────────────
export async function handleReassignmentQueue(
  params: Record<string, string>,
): Promise<any> {
  const limit = parseInt(params.limit || "50", 10);
  const woId = params.wo_id || "";
  const branch = normalizeBranch(params.branch || "all");

  try {
    const isMissingSchemaError = (e: any): boolean => {
      const msg = String(e?.message || "");
      return /no such table|no such column|SQLITE_UNKNOWN|SQL_INPUT_ERROR/i
        .test(msg);
    };
    const executeOptional = async (
      sql: string,
      args: any[] = [],
    ): Promise<any[]> => {
      try {
        return rowsAsObjects(await sqlite.execute({ sql, args }));
      } catch (e: any) {
        if (isMissingSchemaError(e)) return [];
        throw e;
      }
    };

    if (
      params.sync_assignees === "1" &&
      typeof techRosterHandlers.handleDispatchSyncAssignees === "function"
    ) {
      await techRosterHandlers.handleDispatchSyncAssignees(params);
    }
    if (
      params.seed_test === "1" &&
      typeof techRosterHandlers.handleDispatchSeedReassignmentTest ===
        "function"
    ) {
      await techRosterHandlers.handleDispatchSeedReassignmentTest(params);
    }

    // ── Queue rows ────────────────────────────────────────────────────────────
    const queueSql = woId
      ? `SELECT * FROM reassignment_queue WHERE wo_id = ? ORDER BY first_seen_at DESC LIMIT ?`
      : `SELECT * FROM reassignment_queue ORDER BY first_seen_at DESC LIMIT ?`;
    const queueArgs: any[] = woId ? [woId, limit] : [limit];

    // ── Audit log rows ────────────────────────────────────────────────────────
    const auditSql = woId
      ? `SELECT * FROM wo_audit_log WHERE wo_id = ? ORDER BY created_at DESC LIMIT ?`
      : `SELECT * FROM wo_audit_log ORDER BY created_at DESC LIMIT ?`;
    const auditArgs: any[] = woId ? [woId, 200] : [200];

    // ── Run all queries in parallel for speed ─────────────────────────────────
    // Staggering the rate at which requests are issued prevents congestion
    // but parallel reads on the same DB connection are safe in Turso.
    const [
      queue,
      audit,
      blasts,
      claims,
      techRows,
      monitored,
    ] = await Promise.all([
      executeOptional(queueSql, queueArgs),
      executeOptional(auditSql, auditArgs),
      executeOptional(
        `SELECT * FROM blast_events ORDER BY blasted_at DESC LIMIT 20`,
      ),
      executeOptional(`SELECT tc.*, be.property_addr, be.category, be.priority
               FROM   tier2_claims tc
               JOIN   blast_events be ON be.id = tc.blast_id
               WHERE  be.status = 'open'
               ORDER  BY tc.sms_sent_at DESC
               LIMIT  50`),
      executeOptional(`SELECT tech_id, tech_name, tech_phone, tier, active,
                      active_wo_count, performance_score,
                      load_weight, target_share_pct, geo_zone
               FROM   tech_grades
               ORDER  BY tier ASC, active DESC,
                         (active_wo_count * COALESCE(load_weight,1.0)) ASC`),
      executeOptional(
        `SELECT wo_id, created_at FROM monitored_work_orders ORDER BY created_at DESC LIMIT 500`,
      ),
    ]);

    let techs = techRows;

    const cfg = rowsAsObjects(
      await sqlite.execute({
        sql:
          `SELECT key, value FROM proxy_config WHERE key IN ('dispatch_hidden_assignees')`,
        args: [],
      }),
    );
    const hiddenRaw = cfg.find((r: any) =>
      r.key === "dispatch_hidden_assignees"
    )
      ?.value || "{}";
    const hiddenMap = parseHiddenAssigneeMap(hiddenRaw);

    techs = techs.filter((t: any) => !isTechHidden(hiddenMap, t.tech_id));
    if (branch !== "all") {
      techs = techs.filter((t: any) => techBranch(t) === branch);
    }

    const allowedTechs = new Set(techs.map((t: any) => String(t.tech_id)));
    let filteredQueue = queue.filter((r: any) =>
      !isTechHidden(hiddenMap, r.assigned_tech_id)
    );
    if (branch !== "all") {
      filteredQueue = filteredQueue.filter((r: any) => {
        const tid = String(r.assigned_tech_id || "");
        if (!tid) return false;
        return allowedTechs.has(tid);
      });
    }

    // ── Aggregate stats ───────────────────────────────────────────────────────
    const stats = {
      total: filteredQueue.length,
      warning_pending: filteredQueue.filter((r: any) =>
        r.warning_sent && !r.reassignment_count
      ).length,
      warned_total: filteredQueue.filter((r: any) => r.warning_sent).length,
      exempt: filteredQueue.filter((r: any) => r.auto_exempt).length,
      grace_active: filteredQueue.filter((r: any) =>
        r.grace_used && !r.reassignment_count
      ).length,
      reassigned_once: filteredQueue.filter((r: any) =>
        Number(r.reassignment_count) === 1
      ).length,
      reassigned_multi: filteredQueue.filter((r: any) =>
        Number(r.reassignment_count) >= 2
      ).length,
      escalated: filteredQueue.filter((r: any) => r.escalated).length,
      total_reassignments: filteredQueue.reduce(
        (s: number, r: any) => s + (Number(r.reassignment_count) || 0),
        0,
      ),

      // Blast stats
      open_blasts: blasts.filter((b: any) => b.status === "open").length,
      claimed_blasts: blasts.filter((b: any) => b.status === "claimed").length,
      expired_blasts: blasts.filter((b: any) => b.status === "expired").length,

      // Tech roster summary
      tier1_active: techs.filter((t: any) => t.tier === 1 && t.active).length,
      tier2_active: techs.filter((t: any) => t.tier === 2 && t.active).length,
    };

    // ── Enrich queue rows with human-readable status labels ───────────────────
    const enrichedQueue = filteredQueue.map((r: any) => ({
      ...r,
      status_label: r.auto_exempt
        ? "🔕 Exempt (:stop-auto:)"
        : r.escalated
        ? "🚨 Escalated"
        : Number(r.reassignment_count) >= 2
        ? "🔴 Multi-Reassigned"
        : Number(r.reassignment_count) === 1
        ? "🟡 Reassigned Once"
        : r.grace_used
        ? "🔵 Grace Used"
        : r.warning_sent
        ? "⚠️ Warning Sent"
        : "🟢 Monitoring",
    }));

    // ── Enrich audit rows — parse event_data JSON for display ─────────────────
    const enrichedAudit = audit.map((a: any) => {
      let parsed: any = {};
      try {
        parsed = JSON.parse(a.event_data || "{}");
      } catch { /* keep empty */ }
      return { ...a, event_data_parsed: parsed };
    });

    // ── Build blast display with countdown ────────────────────────────────────
    const now = Date.now();
    const enrichedBlasts = blasts.map((b: any) => {
      const expiresMs = new Date(b.expires_at || 0).getTime();
      const remainMs = expiresMs - now;
      return {
        ...b,
        expires_in_hrs: remainMs > 0
          ? (remainMs / 3600_000).toFixed(1)
          : "expired",
        is_expired: remainMs <= 0,
      };
    });

    const result = {
      ok: true,
      stats,
      queue: enrichedQueue,
      audit: enrichedAudit,
      blasts: enrichedBlasts,
      tier2_claims: claims,
      tech_roster: techs,
      monitored_work_orders: monitored,
      meta: {
        queue_count: filteredQueue.length,
        audit_count: audit.length,
        blast_count: blasts.length,
        claims_count: claims.length,
        techs_count: techs.length,
        monitored_count: monitored.length,
        generated_at: new Date().toISOString(),
        branch,
      },
    };

    return result;
  } catch (e: any) {
    return { ok: false, error: e.message };
  }
}