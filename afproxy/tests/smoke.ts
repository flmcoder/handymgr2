// ============================================================================
// tests/smoke.ts — Action contract and schema regression checks.
//
// Verifies that:
//   1. All documented GET actions appear as `case "..."` entries in main.ts.
//   2. All documented POST actions appear in the POST switch block.
//   3. All expected DB tables appear in db.ts ensureTables().
//
// Run with:
//   deno test --allow-read tests/smoke.ts
//
// Or as a plain script (always exits 0/1):
//   deno run --allow-read tests/smoke.ts
//
// This file reads main.ts and db.ts as plain text so no env vars, DB
// connections, dependencies, or network access are needed.
// ============================================================================

// ── Helpers ───────────────────────────────────────────────────────────────────

let _mainTs = "";
let _dbTs = "";

async function loadSources(): Promise<void> {
  const base = new URL("../", import.meta.url).pathname;
  try {
    _mainTs = await Deno.readTextFile(base + "main.ts");
  } catch (e) {
    console.error(`Failed to load main.ts:`, e);
    throw e;
  }
  try {
    _dbTs = await Deno.readTextFile(base + "db.ts");
  } catch (e) {
    console.error(`Failed to load db.ts:`, e);
    throw e;
  }
}

function assertInSource(
  source: string,
  label: string,
  expected: string[],
  failures: string[],
): void {
  for (const item of expected) {
    // Check for quoted strings ("item", 'item') or bare word boundaries (\bitem\b)
    // This covers: case "action", case 'action', CREATE TABLE api_cache, etc.
    const quotePattern = source.includes(`"${item}"`) ||
      source.includes(`'${item}'`);
    const wordBoundary = new RegExp(`\\b${item}\\b`).test(source);
    if (!quotePattern && !wordBoundary) {
      failures.push(`${label}: missing "${item}"`);
    }
  }
}

// ── Contract lists ────────────────────────────────────────────────────────────

/** All GET action strings that main.ts MUST route. */
const REQUIRED_GET_ACTIONS: readonly string[] = [
  "ping",
  "version",
  "session_info",
  "work_orders",
  "wo_notes",
  "wo_detail",
  "wo_billed_amount",
  "turn_work_orders",
  "work_orders_completed_history",
  "completed_work_orders_history",
  "labor",
  "recent_tasks",
  "vendors",
  "inspections",
  "bills",
  "bills_stats",
  "bill_detail",
  "bills_history",
  "bill_attachments",
  "properties",
  "property_groups",
  "property_map",
  "upcoming_moveouts",
  "turns",
  "unit_turns",
  "turn_records",
  "closed_turns",
  "unit_turns_history",
  "wo_comparison_report",
  "webhook_events",
  "webhook_stats",
  "webhook_resolve",
  "webhook_migrate",
  "webhook_cleanup",
  "webhook_live",
  "report",
  "cache_stats",
  "cache_invalidate",
  "storage_cleanup",
  "debug_sqlite",
  "passthrough",
  "reassignment_queue",
  "trusted_devices",
  "pm_proxy_users",
  "settings_get",
  "tech_roster",
  "tenant_comms_log",
  "users",
  "noon_warning_cron",
  "midnight_reassign_cron",
  "dispatch_sync_assignees",
  "dispatch_seed_reassignment_test",
] as const;

/** All POST action strings that main.ts MUST route. */
const REQUIRED_POST_ACTIONS: readonly string[] = [
  "turn_records",
  "turn_record_stage",
  "wo_note_create",
  "unit_turns_sync",
  "unit_turn_wo_link",
  "unit_turn_wo_unlink",
  "report",
  "passthrough",
  "webhook",
  "device_setup",
  "device_otp_request",
  "device_otp_verify",
  "verify_role",
  "trusted_device_revoke",
  "pm_proxy_user_upsert",
  "pm_proxy_user_delete",
  "webhook_resolve",
  "sql_query",
  "sql_execute",
  "settings_set",
  "tech_roster",
  "vendor_override",
  "portal_validate",
  "portal_schedule",
  "portal_reschedule",
  "portal_note",
  "portal_no_contact",
  "portal_reassign_request",
  "generate_magic_link",
  "add_monitored_work_order",
  "remove_monitored_work_order",
  "send_tenant_sms",
  "send_magic_link_test_sms",
  "noon_warning_cron",
  "midnight_reassign_cron",
  "dispatch_sync_assignees",
  "dispatch_seed_reassignment_test",
] as const;

/** Core SQLite tables that ensureTables() MUST create. */
const REQUIRED_TABLES: readonly string[] = [
  "api_cache",
  "turn_records",
  "webhook_events",
  "wo_states",
  "wo_state_changes",
  "trusted_devices",
  "device_otps",
  "reassignment_queue",
  "magic_link_tokens",
  "short_links",
  "portal_events",
  "tenant_comms_log",
  "monitored_work_orders",
  "wo_audit_log",
  "tech_grades",
  "blast_events",
  "proxy_config",
  "vendor_overrides",
  "pm_proxy_users",
  "wo_notes_cache",
  "app_settings",
  "work_orders_cache",
  "pm_proxy_login_audit",
  "routing_events",
  "closed_turns",
  "unit_turn_tracker",
  "unit_turn_milestones",
  "unit_turn_work_orders",
] as const;

// ── Runner ────────────────────────────────────────────────────────────────────

async function runChecks(): Promise<{ passed: number; failed: string[] }> {
  const failures: string[] = [];
  let passed = 0;

  await loadSources();

  // 1. GET action contract
  const getFailures: string[] = [];
  assertInSource(_mainTs, "GET action", REQUIRED_GET_ACTIONS as string[], getFailures);
  passed += REQUIRED_GET_ACTIONS.length - getFailures.length;
  failures.push(...getFailures);

  // 2. POST action contract
  const postFailures: string[] = [];
  assertInSource(
    _mainTs,
    "POST action",
    REQUIRED_POST_ACTIONS as string[],
    postFailures,
  );
  passed += REQUIRED_POST_ACTIONS.length - postFailures.length;
  failures.push(...postFailures);

  // 3. DB table contract
  const tableFailures: string[] = [];
  assertInSource(_dbTs, "DB table", REQUIRED_TABLES as string[], tableFailures);
  passed += REQUIRED_TABLES.length - tableFailures.length;
  failures.push(...tableFailures);

  return { passed, failed: failures };
}

// ── Entry point ───────────────────────────────────────────────────────────────

if (import.meta.main) {
  const { passed, failed } = await runChecks();
  const total = passed + failed.length;
  console.log(`\nHandyManager Proxy — smoke checks`);
  console.log(`  passed : ${passed}/${total}`);
  if (failed.length) {
    console.log(`  FAILED : ${failed.length}`);
    for (const f of failed) console.log(`    ✗ ${f}`);
    Deno.exit(1);
  }
  console.log(`  All checks passed.`);
}

// ── Deno.test wrappers (used by deno test) ────────────────────────────────────

Deno.test("action contract: all required GET actions routed in main.ts", async () => {
  await loadSources();
  const failures: string[] = [];
  assertInSource(_mainTs, "GET action", REQUIRED_GET_ACTIONS as string[], failures);
  if (failures.length) throw new Error(failures.join("\n"));
});

Deno.test("action contract: all required POST actions routed in main.ts", async () => {
  await loadSources();
  const failures: string[] = [];
  assertInSource(
    _mainTs,
    "POST action",
    REQUIRED_POST_ACTIONS as string[],
    failures,
  );
  if (failures.length) throw new Error(failures.join("\n"));
});

Deno.test("schema contract: all required tables created in db.ts", async () => {
  await loadSources();
  const failures: string[] = [];
  assertInSource(_dbTs, "DB table", REQUIRED_TABLES as string[], failures);
  if (failures.length) throw new Error(failures.join("\n"));
});
