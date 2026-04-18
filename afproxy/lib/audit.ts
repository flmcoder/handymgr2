// ============================================================================
// lib/audit.ts — Structured event audit logger.
//
// Every automated action the reassignment engine takes is recorded in
// wo_audit_log. Non-fatal: a write failure never blocks the calling handler.
// ============================================================================

import { sqlite } from "../db.ts";

/**
 * Append a structured event to wo_audit_log.
 *
 * @param woId      AppFolio work order ID (UUID string)
 * @param eventType Snake-case event label, e.g. "auto_reassigned"
 * @param eventData Arbitrary context object — stored as JSON text
 * @param actor     Who triggered the event: "system" | "admin" | "tech"
 */
export async function auditLog(
  woId: string,
  eventType: string,
  eventData: Record<string, any> = {},
  actor = "system",
): Promise<void> {
  try {
    await sqlite.execute({
      sql: `INSERT INTO wo_audit_log
               (wo_id, event_type, event_data, actor)
             VALUES (?, ?, ?, ?)`,
      args: [woId, eventType, JSON.stringify(eventData), actor],
    });
  } catch {
    // Intentionally non-fatal — audit failure must never block the main flow.
  }
}

// ── Canonical event_type values (kept here for reference) ────────────────────
//
//  auto_exempt_activated       :stop-auto: detected; WO removed from automation
//  reassignment_warning_sent   12-hour pre-reassignment warning SMS dispatched
//  grace_period_granted        Note activity found; one-time 48-hr clock reset
//  auto_reassigned             Midnight cron PATCHed WO to new tech
//  escalation_tier2_blast      No Tier 1 tech available; Tier 2 blast fired
//  tier2_blast_sent            SMS sent to all Tier 2 techs simultaneously
//  escalation_no_tech          No techs at all; admin notified manually
//  tenant_sms_sent             Tech used magic link portal to SMS tenant
//  tech_roster_updated         Tech added / modified in tech_grades table