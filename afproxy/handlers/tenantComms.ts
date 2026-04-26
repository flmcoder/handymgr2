// ============================================================================
// handlers/tenantComms.ts — Tenant SMS dispatcher + communications log.
//
// Exports:
//   handleSendTenantSMS  — validates magic link token and sends one of three
//                          pre-filled SMS templates to the tenant via RC
//   handleTenantCommsLog — query stored comms log by WO or all recent entries
//
// This handler is the server-side target of the magic portal's fetch() call.
// It performs a second token validation (in addition to the portal's render-
// time check) before allowing any SMS to be sent — ensuring the token
// has not expired, been tampered with, or already been consumed.
//
// After a successful send:
//   1. Token is marked used in magic_link_tokens (single-use enforced)
//   2. Communication is logged to tenant_comms_log
//   3. A system note is written to the AppFolio work order
//   4. An entry is appended to wo_audit_log
//
// The note write uses "Body" (capital B) as confirmed by AppFolio docs.
// Ensure parameters contain valid values before note writes to avoid
// semantic errors .
// In order to carry out the note write, the correct user role must be
// enabled in AppFolio Property Manager .
// ============================================================================

import { rowsAsObjects, sqlite } from "../db.ts";
import { sendSMS } from "../lib/ringcentral.ts";
import {
  buildShortLinkUrl,
  createMagicLinkSession,
  generateMagicLink,
  isDispatchShortLink,
  lookupMagicToken,
  markTokenUsed,
  verifyMagicToken,
} from "../lib/auth.ts";
import {
  isValidAppfolioDate,
  isValidAppfolioDateTimeUtc,
  isValidAppfolioWorkOrderStatus,
  patchWorkOrder,
  postWoNote,
  uploadWorkOrderAttachment,
} from "../lib/appfolio.ts";
import { auditLog } from "../lib/audit.ts";
import { PROXY_ADMIN_KEY } from "../config.ts";

function parseJsonSafe(raw: string | null | undefined): Record<string, any> {
  try {
    return JSON.parse(String(raw || "{}"));
  } catch {
    return {};
  }
}

function stripHtml(input: string): string {
  return String(input || "").replace(/<[^>]*>/g, " ").replace(/\s+/g, " ")
    .trim();
}

function requestIp(req: Request): string {
  return req.headers.get("x-forwarded-for") ||
    req.headers.get("cf-connecting-ip") ||
    req.headers.get("x-real-ip") ||
    "";
}

function resolveRequiredDispatchShortLink(
  url: string,
  shortCode?: string | null,
): string {
  return isDispatchShortLink(url, shortCode) ? String(url || "").trim() : "";
}

function buildPortalContext(
  payload: Record<string, any>,
  tokenRow: any,
): Record<string, any> {
  const meta = parseJsonSafe(tokenRow?.meta_json);
  const shortCode = String(tokenRow?.short_code || "").trim();
  return {
    ...meta,
    ...payload,
    token: tokenRow?.token || "",
    wo_id: String(payload.wo_id || tokenRow?.wo_id || ""),
    wo_number: String(meta.wo_number || payload.wo_number || ""),
    property_address: String(
      payload.property_address || tokenRow?.property_address ||
        meta.property_address || "",
    ),
    tech_id: String(payload.tech_id || tokenRow?.tech_id || ""),
    tech_name: String(payload.tech_name || tokenRow?.tech_name || ""),
    tenant_phone: String(payload.tenant_phone || tokenRow?.tenant_phone || ""),
    tenant_name: String(payload.tenant_name || tokenRow?.tenant_name || ""),
    short_code: shortCode,
    short_url: shortCode ? buildShortLinkUrl(shortCode) : "",
    used: Number(tokenRow?.used || 0),
    used_template: String(tokenRow?.used_template || ""),
    lang_pref: String(
      tokenRow?.lang_pref || meta.lang_pref || payload.lang_pref || "en",
    ),
    scheduled_date: String(
      tokenRow?.scheduled_date || meta.scheduled_date || "",
    ),
    scheduled_window: String(
      tokenRow?.scheduled_window || meta.scheduled_window || "",
    ),
    stop_auto: Number(tokenRow?.stop_auto || 0),
    exempt_until: String(tokenRow?.exempt_until || meta.exempt_until || ""),
    last_action: String(tokenRow?.last_action || ""),
    last_action_at: String(tokenRow?.last_action_at || ""),
    portal_opened: Number(tokenRow?.portal_opened || 0),
    portal_opened_at: String(tokenRow?.portal_opened_at || ""),
  };
}

async function verifyPortalContext(token: string): Promise<
  | { ok: false; error: string }
  | {
    ok: true;
    payload: Record<string, any>;
    tokenRow: any;
    context: Record<string, any>;
  }
> {
  if (!token) return { ok: false, error: "Missing token" };

  const payload = await verifyMagicToken(token);
  if (!payload) {
    return {
      ok: false,
      error:
        "Magic link has expired or is invalid. Request a new dispatch notification.",
    };
  }

  const tokenRow = await lookupMagicToken(token);
  if (!tokenRow) {
    return {
      ok: false,
      error:
        "Magic link record not found. Request a new dispatch notification.",
    };
  }

  const context = buildPortalContext(payload, tokenRow);
  return { ok: true, payload, tokenRow, context };
}

async function logPortalEvent(
  req: Request,
  token: string,
  woId: string,
  techId: string,
  action: string,
  payload: Record<string, any> = {},
): Promise<void> {
  try {
    await sqlite.execute({
      sql:
        `INSERT INTO portal_events (token, wo_id, tech_id, action, payload, ip, user_agent)
             VALUES (?, ?, ?, ?, ?, ?, ?)`,
      args: [
        token,
        woId,
        techId,
        action,
        JSON.stringify(payload || {}),
        requestIp(req),
        req.headers.get("user-agent") || "",
      ],
    });
  } catch (e: any) {
    console.log(`portal_events write failed: ${e.message}`);
  }
}

async function readJsonBody(req: Request): Promise<any> {
  try {
    return await req.json();
  } catch {
    return null;
  }
}

function readPortalUploadToken(req: Request): string {
  try {
    const url = new URL(req.url);
    return String(
      url.searchParams.get("token") || req.headers.get("x-portal-token") || "",
    ).trim();
  } catch {
    return String(req.headers.get("x-portal-token") || "").trim();
  }
}

function readPortalUploadPhase(req: Request): string {
  try {
    const url = new URL(req.url);
    return String(url.searchParams.get("phase") || "general").trim().toLowerCase();
  } catch {
    return "general";
  }
}

export async function handlePortalValidate(req: Request): Promise<any> {
  const body = await readJsonBody(req);
  if (!body) {
    return { ok: false, error: "Invalid JSON body — expected { token }" };
  }

  const verified = await verifyPortalContext(String(body.token || ""));
  if (!verified.ok) return { ok: false, valid: false, error: verified.error };

  await sqlite.execute({
    sql: `UPDATE magic_link_tokens
             SET portal_opened = 1,
                 portal_opened_at = COALESCE(portal_opened_at, datetime('now')),
                 last_action = 'portal_opened',
                 last_action_at = datetime('now')
           WHERE token = ?`,
    args: [verified.tokenRow.token],
  }).catch((e: any) => {
    console.log(
      `magic_link_tokens portal_opened update failed: ${String(e?.message || e)}`,
    );
  });

  await logPortalEvent(
    req,
    verified.tokenRow.token,
    verified.context.wo_id,
    verified.context.tech_id,
    "portal_opened",
    { lang_pref: verified.context.lang_pref },
  );

  return {
    ok: true,
    valid: true,
    portal: verified.context,
  };
}

export async function handleGenerateMagicLink(req: Request): Promise<any> {
  const body = await readJsonBody(req);
  if (!body) {
    return {
      ok: false,
      error: "Invalid JSON body — expected magic link generation payload",
    };
  }

  const woId = String(body.wo_id || "").trim();
  const techId = String(body.tech_id || "").trim();
  const techName = String(body.tech_name || "").trim();
  const techPhone = String(body.tech_phone || "").trim();
  const tenantPhone = String(body.tenant_phone || "").trim();
  const tenantName = String(body.tenant_name || "Tenant").trim();
  const propertyAddress = String(body.property_address || "").trim();
  const sendToTech = body.send_sms === true;

  if (!woId || !techId || !techName || !tenantPhone || !propertyAddress) {
    return {
      ok: false,
      error:
        "Missing required fields: wo_id, tech_id, tech_name, tenant_phone, property_address",
    };
  }

  // Avoid blocking magic-link generation on live AppFolio lookups.
  // Token validation and downstream action checks remain server-side.

  const extraPayload = {
    wo_number: String(body.wo_number || "").trim(),
    wo_description: String(body.wo_description || "").trim(),
    wo_access_notes: String(body.wo_access_notes || "").trim(),
    vendor_notes: String(body.vendor_notes || "").trim(),
    tenant_email: String(body.tenant_email || "").trim(),
    pm_name: String(body.pm_name || "").trim(),
    pm_phone: String(body.pm_phone || "").trim(),
    pm_email: String(body.pm_email || "").trim(),
    tech_phone: techPhone,
    lang_pref: String(body.lang_pref || "en").trim() || "en",
  };

  const session = await createMagicLinkSession(
    woId,
    techId,
    techName,
    tenantPhone,
    tenantName,
    propertyAddress,
    extraPayload,
  );
  const shortMagicLink = resolveRequiredDispatchShortLink(
    session.url,
    session.shortCode,
  );
  if (!shortMagicLink) {
    return {
      ok: false,
      error:
        "Dispatch short link could not be created. No raw portal URL was sent.",
      token: session.token,
      expires_at: session.expiresAt,
    };
  }

  let smsSent = false;
  let smsResult: any = null;
  if (sendToTech) {
    if (!techPhone) {
      return {
        ok: false,
        error: "tech_phone is required when send_sms=true",
        magic_link: shortMagicLink,
        short_code: session.shortCode,
      };
    }

    const woLabel = extraPayload.wo_number || woId;
    const smsMessage =
      `Fort Lowell Realty dispatch link for WO #${woLabel}: ${shortMagicLink}`;
    smsResult = await sendSMS(techPhone, smsMessage);
    smsSent = !!smsResult.ok;
  }

  await logPortalEvent(
    req,
    session.token,
    woId,
    techId,
    "magic_link_generated",
    {
      short_url: shortMagicLink,
      short_code: session.shortCode,
      send_sms: sendToTech,
      sms_sent: smsSent,
    },
  );
  await auditLog(woId, "magic_link_generated", {
    tech: techName,
    short_code: session.shortCode,
    sms_sent: smsSent,
  });

  if (sendToTech && !smsSent) {
    return {
      ok: false,
      error: smsResult?.error || "RingCentral send failed",
      magic_link: shortMagicLink,
      token: session.token,
      short_code: session.shortCode,
      expires_at: session.expiresAt,
    };
  }

  return {
    ok: true,
    token: session.token,
    magic_link: shortMagicLink,
    long_link: session.longUrl,
    short_code: session.shortCode,
    expires_at: session.expiresAt,
    sms_sent: smsSent,
  };
}

async function upsertPortalExemption(
  woId: string,
  woNumber: string,
  propertyAddress: string,
  techId: string,
  techName: string,
  reason: string,
): Promise<void> {
  await sqlite.execute({
    sql: `INSERT INTO reassignment_queue
             (wo_id, wo_number, property_address, assigned_tech_id, assigned_tech_name,
              auto_exempt, auto_exempt_at, auto_exempt_by)
           VALUES (?, ?, ?, ?, ?, 1, datetime('now'), ?)
           ON CONFLICT(wo_id) DO UPDATE SET
             wo_number = excluded.wo_number,
             property_address = excluded.property_address,
             assigned_tech_id = excluded.assigned_tech_id,
             assigned_tech_name = excluded.assigned_tech_name,
             auto_exempt = 1,
             auto_exempt_at = datetime('now'),
             auto_exempt_by = excluded.auto_exempt_by`,
    args: [woId, woNumber, propertyAddress, techId, techName, reason],
  });
}

function buildScheduledExemptUntil(scheduledDate: string): string | null {
  if (!isValidAppfolioDate(scheduledDate)) return null;
  const value = `${scheduledDate}T23:59:59Z`;
  return isValidAppfolioDateTimeUtc(value) ? value : null;
}

async function portalScheduleAction(
  req: Request,
  actionName: "scheduled" | "rescheduled",
): Promise<any> {
  const body = await readJsonBody(req);
  if (!body) {
    return {
      ok: false,
      error: "Invalid JSON body — expected scheduling payload",
    };
  }

  const verified = await verifyPortalContext(String(body.token || ""));
  if (!verified.ok) return { ok: false, error: verified.error };

  const scheduledDate = String(body.scheduled_date || "").trim();
  const scheduledWindow = String(body.scheduled_window || "").trim();
  if (!scheduledDate || !scheduledWindow) {
    return {
      ok: false,
      error: "scheduled_date and scheduled_window are required",
    };
  }

  if (!isValidAppfolioDate(scheduledDate)) {
    return {
      ok: false,
      error: "scheduled_date must be formatted as YYYY-MM-DD",
    };
  }

  const exemptUntil = buildScheduledExemptUntil(scheduledDate);
  if (!exemptUntil) {
    return {
      ok: false,
      error:
        "Could not build a valid UTC exempt_until timestamp from scheduled_date",
    };
  }

  const status = "Scheduled";
  if (!isValidAppfolioWorkOrderStatus(status)) {
    return { ok: false, error: "Configured AppFolio status is invalid" };
  }

  const patchRes = await patchWorkOrder(verified.context.wo_id, {
    Status: status,
  });
  if (!patchRes.ok) {
    await logPortalEvent(
      req,
      verified.tokenRow.token,
      verified.context.wo_id,
      verified.context.tech_id,
      actionName,
      {
        scheduled_date: scheduledDate,
        scheduled_window: scheduledWindow,
        patch_ok: false,
        state_persisted: false,
        partial: true,
      },
    );
    await auditLog(verified.context.wo_id, `portal_${actionName}`, {
      tech: verified.context.tech_name,
      scheduled_date: scheduledDate,
      scheduled_window: scheduledWindow,
      patch_ok: false,
      state_persisted: false,
      partial: true,
    });
    return {
      ok: false,
      partial: true,
      error: "Failed to update AppFolio work order status; portal schedule was not persisted",
      scheduled_date: scheduledDate,
      scheduled_window: scheduledWindow,
      patch_ok: false,
      patch_status: patchRes.status,
      note_written: false,
      state_persisted: false,
    };
  }

  const noteBody =
    `[SYSTEM — HandyManager Portal ${
      actionName === "scheduled" ? "Schedule" : "Reschedule"
    } | ${new Date().toISOString()}]\n` +
    `Tech ${verified.context.tech_name} marked this work order as ${actionName}.\n` +
    `Scheduled date: ${scheduledDate}\n` +
    `Arrival window: ${scheduledWindow}`;
  const noteOk = await postWoNote(verified.context.wo_id, noteBody);

  await sqlite.execute({
    sql: `UPDATE magic_link_tokens
             SET scheduled_date = ?,
                 scheduled_window = ?,
                 stop_auto = 1,
                 exempt_until = ?,
                 lang_pref = ?,
                 last_action = ?,
                 last_action_at = datetime('now')
           WHERE token = ?`,
    args: [
      scheduledDate,
      scheduledWindow,
      exemptUntil,
      String(body.lang_pref || verified.context.lang_pref || "en"),
      actionName,
      verified.tokenRow.token,
    ],
  });

  await upsertPortalExemption(
    verified.context.wo_id,
    verified.context.wo_number || verified.context.wo_id,
    verified.context.property_address || "",
    verified.context.tech_id,
    verified.context.tech_name,
    actionName,
  );

  await logPortalEvent(
    req,
    verified.tokenRow.token,
    verified.context.wo_id,
    verified.context.tech_id,
    actionName,
    {
      scheduled_date: scheduledDate,
      scheduled_window: scheduledWindow,
      patch_ok: patchRes.ok,
      note_written: noteOk,
      state_persisted: true,
    },
  );
  await auditLog(verified.context.wo_id, `portal_${actionName}`, {
    tech: verified.context.tech_name,
    scheduled_date: scheduledDate,
    scheduled_window: scheduledWindow,
    patch_ok: patchRes.ok,
    note_written: noteOk,
    state_persisted: true,
  });

  return {
    ok: true,
    scheduled_date: scheduledDate,
    scheduled_window: scheduledWindow,
    patch_ok: patchRes.ok,
    patch_status: patchRes.status,
    note_written: noteOk,
    state_persisted: true,
  };
}

export async function handlePortalSchedule(req: Request): Promise<any> {
  return await portalScheduleAction(req, "scheduled");
}

export async function handlePortalReschedule(req: Request): Promise<any> {
  return await portalScheduleAction(req, "rescheduled");
}

export async function handlePortalNote(req: Request): Promise<any> {
  const body = await readJsonBody(req);
  if (!body) {
    return {
      ok: false,
      error: "Invalid JSON body — expected { token, note_text }",
    };
  }

  const verified = await verifyPortalContext(String(body.token || ""));
  if (!verified.ok) return { ok: false, error: verified.error };

  const noteText = stripHtml(String(body.note_text || "")).slice(0, 1000);
  if (!noteText) return { ok: false, error: "note_text is required" };

  const noteOk = await postWoNote(
    verified.context.wo_id,
    `[PORTAL NOTE | ${
      new Date().toISOString()
    }]\nTech ${verified.context.tech_name}: ${noteText}`,
  );

  await sqlite.execute({
    sql: `UPDATE magic_link_tokens
             SET last_action = 'note',
                 last_action_at = datetime('now')
           WHERE token = ?`,
    args: [verified.tokenRow.token],
  }).catch((e: any) => {
    console.log(
      `magic_link_tokens note update failed: ${String(e?.message || e)}`,
    );
  });

  await logPortalEvent(
    req,
    verified.tokenRow.token,
    verified.context.wo_id,
    verified.context.tech_id,
    "note",
    {
      note_text: noteText,
      note_written: noteOk,
    },
  );
  await auditLog(verified.context.wo_id, "portal_note_added", {
    tech: verified.context.tech_name,
    note_written: noteOk,
  });

  return { ok: noteOk, note_written: noteOk };
}

export async function handlePortalStatus(req: Request): Promise<any> {
  const body = await readJsonBody(req);
  if (!body) {
    return {
      ok: false,
      error: "Invalid JSON body — expected { token, status, note_text? }",
    };
  }

  const verified = await verifyPortalContext(String(body.token || ""));
  if (!verified.ok) return { ok: false, error: verified.error };

  const status = String(body.status || "").trim();
  const noteText = stripHtml(String(body.note_text || "")).slice(0, 1200);
  if (!status) return { ok: false, error: "status is required" };
  if (!isValidAppfolioWorkOrderStatus(status)) {
    return { ok: false, error: "status must be one of: Scheduled, Waiting, Work Completed" };
  }

  const patchRes = await patchWorkOrder(verified.context.wo_id, { Status: status });
  if (!patchRes.ok) {
    await logPortalEvent(
      req,
      verified.tokenRow.token,
      verified.context.wo_id,
      verified.context.tech_id,
      "status_failed",
      { status, patch_ok: false, patch_status: patchRes.status },
    );
    return {
      ok: false,
      error: "Failed to update AppFolio work order status",
      status_value: status,
      patch_status: patchRes.status,
    };
  }

  const noteLines = [
    `[PORTAL STATUS | ${new Date().toISOString()}]`,
    `Tech ${verified.context.tech_name} updated the work order status to ${status}.`,
  ];
  if (noteText) noteLines.push(`Details: ${noteText}`);
  const noteOk = await postWoNote(verified.context.wo_id, noteLines.join("\n"));

  await sqlite.execute({
    sql: `UPDATE magic_link_tokens
             SET last_action = ?,
                 last_action_at = datetime('now')
           WHERE token = ?`,
    args: [
      status === "Work Completed" ? "work_completed" : "status_update",
      verified.tokenRow.token,
    ],
  }).catch(() => {});

  await logPortalEvent(
    req,
    verified.tokenRow.token,
    verified.context.wo_id,
    verified.context.tech_id,
    "status",
    { status, note_written: noteOk },
  );
  await auditLog(verified.context.wo_id, "portal_status_updated", {
    tech: verified.context.tech_name,
    status,
    note_written: noteOk,
  });

  return {
    ok: true,
    status_value: status,
    patch_ok: patchRes.ok,
    patch_status: patchRes.status,
    note_written: noteOk,
  };
}

export async function handlePortalPhotoUpload(req: Request): Promise<any> {
  const token = readPortalUploadToken(req);
  const verified = await verifyPortalContext(token);
  if (!verified.ok) return { ok: false, error: verified.error };

  const bodyBuffer = await req.arrayBuffer();
  if (!bodyBuffer || bodyBuffer.byteLength === 0) {
    return { ok: false, error: "Empty upload body" };
  }

  const phase = readPortalUploadPhase(req);
  const contentType = req.headers.get("Content-Type") || "application/octet-stream";
  const uploadRes = await uploadWorkOrderAttachment(
    verified.context.wo_id,
    contentType,
    bodyBuffer,
  );
  if (!uploadRes.ok) {
    await logPortalEvent(
      req,
      verified.tokenRow.token,
      verified.context.wo_id,
      verified.context.tech_id,
      "photo_upload_failed",
      { phase, upload_status: uploadRes.status },
    );
    return {
      ok: false,
      error: "Failed to upload work-order attachment",
      upload_status: uploadRes.status,
      detail: uploadRes.detail,
    };
  }

  try {
    const cacheKey = `wo_attachments_${verified.context.wo_id}`;
    await sqlite.execute({
      sql: `DELETE FROM api_cache WHERE cache_key = ? OR cache_key LIKE ?`,
      args: [cacheKey, `${cacheKey}::%`],
    });
  } catch {
    // non-fatal cache bust
  }

  const phaseLabel = phase === "before"
    ? "before"
    : (phase === "after" ? "after" : "general");
  const noteOk = await postWoNote(
    verified.context.wo_id,
    `[PORTAL PHOTO | ${new Date().toISOString()}]\nTech ${verified.context.tech_name} uploaded a ${phaseLabel} photo through the dispatch portal.`,
  );

  await sqlite.execute({
    sql: `UPDATE magic_link_tokens
             SET last_action = 'photo_upload',
                 last_action_at = datetime('now')
           WHERE token = ?`,
    args: [verified.tokenRow.token],
  }).catch(() => {});

  await logPortalEvent(
    req,
    verified.tokenRow.token,
    verified.context.wo_id,
    verified.context.tech_id,
    "photo_upload",
    { phase: phaseLabel, note_written: noteOk },
  );
  await auditLog(verified.context.wo_id, "portal_photo_uploaded", {
    tech: verified.context.tech_name,
    phase: phaseLabel,
    note_written: noteOk,
  });

  return {
    ok: true,
    upload_status: uploadRes.status,
    phase: phaseLabel,
    note_written: noteOk,
    detail: uploadRes.detail,
  };
}

export async function handlePortalNoContact(req: Request): Promise<any> {
  const body = await readJsonBody(req);
  if (!body) {
    return {
      ok: false,
      error: "Invalid JSON body — expected no-contact payload",
    };
  }

  const verified = await verifyPortalContext(String(body.token || ""));
  if (!verified.ok) return { ok: false, error: verified.error };

  const attempts = Math.max(1, Number(body.attempts || 1));
  const details = stripHtml(String(body.details || "")).slice(0, 500);
  const pmPhone = String(verified.context.pm_phone || "").trim();
  let pmSmsOk = false;

  if (pmPhone) {
    const sms = await sendSMS(
      pmPhone,
      `No-contact alert: ${verified.context.tech_name} could not reach ${
        verified.context.tenant_name || "resident"
      } for WO #${
        verified.context.wo_number || verified.context.wo_id
      } at ${verified.context.property_address}. Attempts: ${attempts}.${
        details ? " Details: " + details : ""
      }`,
    );
    pmSmsOk = !!sms.ok;
  }

  const noteOk = await postWoNote(
    verified.context.wo_id,
    `[PORTAL NO-CONTACT | ${
      new Date().toISOString()
    }]\nTech ${verified.context.tech_name} reported resident no-contact after ${attempts} attempt(s).${
      details ? `\nDetails: ${details}` : ""
    }`,
  );

  await sqlite.execute({
    sql: `UPDATE magic_link_tokens
             SET last_action = 'no_contact',
                 last_action_at = datetime('now')
           WHERE token = ?`,
    args: [verified.tokenRow.token],
  }).catch((e: any) => {
    console.log(
      `magic_link_tokens no_contact update failed: ${String(e?.message || e)}`,
    );
  });

  await logPortalEvent(
    req,
    verified.tokenRow.token,
    verified.context.wo_id,
    verified.context.tech_id,
    "no_contact",
    {
      attempts,
      details,
      pm_sms_ok: pmSmsOk,
      note_written: noteOk,
    },
  );
  await auditLog(verified.context.wo_id, "portal_no_contact_reported", {
    tech: verified.context.tech_name,
    attempts,
    pm_sms_ok: pmSmsOk,
    note_written: noteOk,
  });

  return { ok: true, pm_sms_ok: pmSmsOk, note_written: noteOk };
}

export async function handlePortalReassignRequest(req: Request): Promise<any> {
  const body = await readJsonBody(req);
  if (!body) {
    return {
      ok: false,
      error: "Invalid JSON body — expected reassignment payload",
    };
  }

  const verified = await verifyPortalContext(String(body.token || ""));
  if (!verified.ok) return { ok: false, error: verified.error };

  const reason = stripHtml(String(body.reason || "")).slice(0, 120);
  const details = stripHtml(String(body.details || "")).slice(0, 500);
  if (!reason) return { ok: false, error: "reason is required" };

  await sqlite.execute({
    sql: `INSERT INTO reassignment_queue
             (wo_id, wo_number, property_address, assigned_tech_id, assigned_tech_name,
              escalated)
           VALUES (?, ?, ?, ?, ?, 1)
           ON CONFLICT(wo_id) DO UPDATE SET
             wo_number = excluded.wo_number,
             property_address = excluded.property_address,
             assigned_tech_id = excluded.assigned_tech_id,
             assigned_tech_name = excluded.assigned_tech_name,
             escalated = 1`,
    args: [
      verified.context.wo_id,
      verified.context.wo_number || verified.context.wo_id,
      verified.context.property_address || "",
      verified.context.tech_id,
      verified.context.tech_name,
    ],
  });

  const noteOk = await postWoNote(
    verified.context.wo_id,
    `[PORTAL REASSIGN REQUEST | ${
      new Date().toISOString()
    }]\nTech ${verified.context.tech_name} requested reassignment.\nReason: ${reason}${
      details ? `\nDetails: ${details}` : ""
    }`,
  );

  await sqlite.execute({
    sql: `UPDATE magic_link_tokens
             SET last_action = 'reassign_request',
                 last_action_at = datetime('now')
           WHERE token = ?`,
    args: [verified.tokenRow.token],
  }).catch((e: any) => {
    console.log(
      `magic_link_tokens reassign_request update failed: ${String(e?.message || e)}`,
    );
  });

  await logPortalEvent(
    req,
    verified.tokenRow.token,
    verified.context.wo_id,
    verified.context.tech_id,
    "reassign_request",
    {
      reason,
      details,
      note_written: noteOk,
    },
  );
  await auditLog(verified.context.wo_id, "portal_reassign_requested", {
    tech: verified.context.tech_name,
    reason,
    note_written: noteOk,
  });

  return { ok: true, note_written: noteOk };
}

// ── Message templates ─────────────────────────────────────────────────────────
// Three templates keyed by the short identifier sent from the portal JS.
// Each template auto-fills tech name and short address from the token payload.
// Designed to feel personal and professional — not automated or robotic.
//
// Template keys: "enroute" | "schedule" | "today"
// Human-readable labels stored in magic_link_tokens.used_template after send.

const TEMPLATE_LABELS: Record<string, string> = {
  enroute: "I'm On My Way",
  schedule: "Let's Schedule a Visit",
  today: "Arriving Today",
};

function buildTemplateMessage(
  template: string,
  techName: string,
  firstName: string,
  shortAddr: string,
): string | null {
  switch (template) {
    case "enroute":
      return (
        `Hi ${firstName}! This is ${techName} with Fort Lowell Realty. ` +
        `I'm currently on my way to your unit at ${shortAddr} and expect ` +
        `to arrive in about 15–20 minutes. ` +
        `Please let me know if you need to reschedule. Thank you!`
      );

    case "schedule":
      return (
        `Hi ${firstName}! This is ${techName} with Fort Lowell Realty. ` +
        `I'm reaching out to schedule the maintenance visit for your unit ` +
        `at ${shortAddr}. ` +
        `What days and times work best for you this week? ` +
        `I look forward to getting this taken care of for you!`
      );

    case "today":
      return (
        `Hi ${firstName}! Heads up from Fort Lowell Realty — ` +
        `${techName} will be stopping by ${shortAddr} today to address ` +
        `your maintenance request. ` +
        `Feel free to reply if you have any questions or concerns!`
      );

    default:
      return null;
  }
}

// ── handleSendTenantSMS ───────────────────────────────────────────────────────
// POST ?action=send_tenant_sms
// Body: { token: string, template: "enroute" | "schedule" | "today" }
//
// Called by the magic portal's embedded JavaScript via fetch().
// Validates the token server-side, builds the template message,
// sends via RingCentral, then burns the token and logs everything.
//
// Returns JSON (not HTML) — the portal JS parses this response.
export async function handleSendTenantSMS(req: Request): Promise<any> {
  // ── Parse request body ────────────────────────────────────────────────────
  let body: any = {};
  try {
    body = await req.json();
  } catch {
    return {
      ok: false,
      error: "Invalid JSON body — expected { token, template }",
    };
  }

  const token = String(body.token || "");
  const template = String(body.template || "");

  // ── Basic validation ──────────────────────────────────────────────────────
  if (!token || !template) {
    return { ok: false, error: "Missing token or template in request body" };
  }
  if (!TEMPLATE_LABELS[template]) {
    return {
      ok: false,
      error:
        `Invalid template key "${template}". Valid values: enroute | schedule | today`,
    };
  }

  // ── Server-side token verification (second check — first was at portal render)
  const payload = await verifyMagicToken(token);
  if (!payload) {
    return {
      ok: false,
      error:
        "Magic link has expired or is invalid. Request a new dispatch notification.",
    };
  }

  // ── Single-use enforcement ────────────────────────────────────────────────
  const tokenRow = await lookupMagicToken(token);
  if (Number(tokenRow?.used || 0) === 1) {
    const usedTemplate = String(
      tokenRow?.used_template || TEMPLATE_LABELS[template] || "template",
    );
    return {
      ok: true,
      already_sent: true,
      template: usedTemplate,
      message: `Tenant message was already sent (${usedTemplate}). You can still use schedule, note, and reassignment actions from this portal.`,
    };
  }

  // ── Extract context from verified payload ─────────────────────────────────
  const techName = String(payload.tech_name || "Your technician");
  const techId = String(payload.tech_id || techName);
  const tenantName = String(payload.tenant_name || "");
  const tenantPhone = String(payload.tenant_phone || "");
  const address = String(payload.property_address || "");
  const woId = String(payload.wo_id || "");
  const testMode = !!payload.test_mode;

  const firstName = tenantName.split(" ")[0] || "there";
  const shortAddr = address.split(",")[0].trim() || address;
  const label = TEMPLATE_LABELS[template];

  // ── Validate tenant phone ─────────────────────────────────────────────────
  if (!tenantPhone) {
    return {
      ok: false,
      error:
        "No tenant phone number found for this work order. Contact dispatch to obtain tenant contact details.",
    };
  }

  // ── Build message body ────────────────────────────────────────────────────
  const messageBody = buildTemplateMessage(
    template,
    techName,
    firstName,
    shortAddr,
  );
  if (!messageBody) {
    return {
      ok: false,
      error: `Could not build message for template "${template}"`,
    };
  }

  // ── Send SMS via RingCentral ──────────────────────────────────────────────
  const smsResult = await sendSMS(tenantPhone, messageBody);
  if (!smsResult.ok) {
    return {
      ok: false,
      error: `SMS delivery failed: ${smsResult.error}`,
    };
  }

  // ── Burn token (single-use mark) ──────────────────────────────────────────
  await markTokenUsed(token, label);

  // ── Log to tenant_comms_log ───────────────────────────────────────────────
  try {
    await sqlite.execute({
      sql: `INSERT INTO tenant_comms_log
               (wo_id, tech_id, tech_name, tenant_phone,
                template_used, message_body, rc_message_id, status, appfolio_noted)
             VALUES (?, ?, ?, ?, ?, ?, ?, 'sent', 0)`,
      args: [
        woId,
        techId,
        techName,
        tenantPhone,
        label,
        messageBody,
        smsResult.message_id || "",
      ],
    });
  } catch (e: any) {
    console.log(`tenant_comms_log write failed: ${e.message}`);
  }

  // ── AppFolio system note ──────────────────────────────────────────────────
  // Documents the tenant contact in AppFolio so all parties can see it.
  // The note write requires the system bot user to have the correct role
  // enabled in AppFolio Property Manager .
  // Ensure parameters contain valid values before this write .
  // Send Test Event can confirm note delivery is working .
  const noteSent = testMode ? false : await postWoNote(
    woId,
    `[SYSTEM — HandyManager Tenant Contact | ${new Date().toISOString()}]\n` +
      `Tech ${techName} sent a tenant notification via RingCentral SMS.\n` +
      `Template: "${label}"\n` +
      `Recipient: tenant contact on file.\n` +
      `Message sent via HandyManager Tech Dispatch Portal (magic link, single-use).`,
  );

  // Update appfolio_noted flag in comms log if note succeeded
  if (noteSent) {
    try {
      await sqlite.execute({
        sql: `UPDATE tenant_comms_log
               SET appfolio_noted = 1
               WHERE wo_id = ? AND rc_message_id = ?`,
        args: [woId, smsResult.message_id || ""],
      });
    } catch { /* non-fatal */ }
  }

  // ── Audit log ─────────────────────────────────────────────────────────────
  if (!testMode) {
    await auditLog(woId, "tenant_sms_sent", {
      tech: techName,
      template: label,
      rc_message_id: smsResult.message_id,
      note_written: noteSent,
    });
  }

  return {
    ok: true,
    message: testMode
      ? `✅ Test message sent to ${firstName}!`
      : `✅ Message sent to ${firstName}!`,
    template: label,
    rc_message_id: smsResult.message_id,
    test_mode: testMode,
  };
}

export async function handleSendMagicLinkTestSMS(req: Request): Promise<any> {
  let body: any = {};
  try {
    body = await req.json();
  } catch {
    return {
      ok: false,
      error: "Invalid JSON body — expected { phone, tech_name?, tech_id? }",
    };
  }

  const phone = String(body.phone || body.to || "").trim();
  const suppliedAdminKey = String(body.key || body.admin_key || "").trim();
  const techName = String(body.tech_name || "HandyManager Dispatch Test")
    .trim();
  const techId = String(body.tech_id || "dispatch-test").trim();
  const tenantName = String(body.tenant_name || "Dispatch Test").trim();
  const propertyAddress = String(
    body.property_address || "HandyManager Test Portal",
  ).trim();

  if (!phone) {
    return { ok: false, error: "Missing phone in request body" };
  }
  if (!PROXY_ADMIN_KEY || suppliedAdminKey !== PROXY_ADMIN_KEY) {
    return {
      ok: false,
      error: "Unauthorized — missing or invalid PROXY_ADMIN_KEY",
    };
  }
  if (!phone.startsWith("+")) {
    return {
      ok: false,
      error: `Invalid phone "${phone}" — use E.164 format like +15205551234`,
    };
  }

  const woId = `test:${Date.now()}`;
  const magicLink = await generateMagicLink(
    woId,
    techId,
    techName,
    phone,
    tenantName,
    propertyAddress,
    { test_mode: true },
  );

  if (!isDispatchShortLink(magicLink)) {
    return {
      ok: false,
      error: "Dispatch short link could not be created — SMS was not sent",
    };
  }

  const message = `HandyManager test portal for ${techName}: ${magicLink} ` +
    `Open the link and send a template message to verify the full RingCentral + magic-link flow.`;

  const smsResult = await sendSMS(phone, message);
  if (!smsResult.ok) {
    return {
      ok: false,
      error: smsResult.error || "RingCentral send failed",
      magic_link: magicLink,
    };
  }

  try {
    await sqlite.execute({
      sql: `INSERT INTO tenant_comms_log
               (wo_id, tech_id, tech_name, tenant_phone,
                template_used, message_body, rc_message_id, status, appfolio_noted)
             VALUES (?, ?, ?, ?, ?, ?, ?, 'sent', 0)`,
      args: [
        woId,
        techId,
        techName,
        phone,
        "Magic Link Test",
        message,
        smsResult.message_id || "",
      ],
    });
  } catch (e: any) {
    console.log(`magic_link_test log failed: ${e.message}`);
  }

  return {
    ok: true,
    message: `Test magic link sent to ${phone}`,
    rc_message_id: smsResult.message_id,
    magic_link: magicLink,
    wo_id: woId,
  };
}

// ── handleTenantCommsLog ──────────────────────────────────────────────────────
// GET ?action=tenant_comms_log
//        &wo_id=UUID        (optional — filters to one WO)
//        &limit=100         (optional — default 100)
//
// Returns the tenant communications log sorted newest first.
// Used by the HandyManager Dispatch Control tab to show all tenant
// contacts made via the magic link portal.
export async function handleTenantCommsLog(
  params: Record<string, string>,
): Promise<any> {
  const limit = parseInt(params.limit || "100", 10);
  const woId = params.wo_id || "";
  const includeWarnings = params.include_warnings !== "0";

  const sql = woId
    ? `SELECT * FROM tenant_comms_log WHERE wo_id = ? ORDER BY sent_at DESC LIMIT ?`
    : `SELECT * FROM tenant_comms_log ORDER BY sent_at DESC LIMIT ?`;
  const args: any[] = woId ? [woId, limit] : [limit];

  const parseJson = (raw: string): any => {
    try {
      return JSON.parse(raw || "{}");
    } catch {
      return {};
    }
  };

  try {
    const rows = rowsAsObjects(await sqlite.execute({ sql, args }));
    const smsRows = rows.map((r: any) => ({
      ...r,
      source_type: "tenant_sms",
      event_type: "tenant_sms_sent",
      event_label: r.template_used || "Tenant SMS",
    }));

    let warningRows: any[] = [];
    if (includeWarnings) {
      const warningSql = woId
        ? `SELECT wo_id, event_type, event_data, created_at
             FROM wo_audit_log
            WHERE wo_id = ?
              AND event_type IN ('reassignment_warning_sent', 'tier2_blast_sent', 'auto_reassigned')
            ORDER BY created_at DESC
            LIMIT ?`
        : `SELECT wo_id, event_type, event_data, created_at
             FROM wo_audit_log
            WHERE event_type IN ('reassignment_warning_sent', 'tier2_blast_sent', 'auto_reassigned')
            ORDER BY created_at DESC
            LIMIT ?`;
      const warningArgs: any[] = woId ? [woId, limit] : [limit];
      const audits = rowsAsObjects(
        await sqlite.execute({
          sql: warningSql,
          args: warningArgs,
        }),
      );

      warningRows = audits.map((a: any) => {
        const ed = parseJson(a.event_data || "{}");
        const label = a.event_type === "reassignment_warning_sent"
          ? "Pre-Reassign Warning"
          : a.event_type === "tier2_blast_sent"
          ? "Tier 2 Blast"
          : "Auto Reassigned";

        return {
          wo_id: a.wo_id,
          tech_id: "",
          tech_name: ed.tech || ed.from_tech || ed.to || "",
          tenant_phone: ed.tenant_phone || "",
          template_used: label,
          message_body: "",
          sent_at: a.created_at,
          rc_message_id: "",
          status: ed.sms_sent === false ? "failed" : "sent",
          appfolio_noted: 0,
          source_type: "dispatch_warning",
          event_type: a.event_type,
          event_label: label,
          magic_link: ed.magic_link || "",
          phone_found: !!ed.phone_found,
          tenant_name: ed.tenant_name || "",
          address: ed.address || "",
          priority: ed.priority || "",
        };
      });
    }

     const portalSql = woId
      ? `SELECT token, wo_id, tech_id, action, payload, created_at
        FROM portal_events WHERE wo_id = ? ORDER BY created_at DESC LIMIT ?`
      : `SELECT token, wo_id, tech_id, action, payload, created_at
        FROM portal_events ORDER BY created_at DESC LIMIT ?`;
    const portalArgs: any[] = woId ? [woId, limit] : [limit];
    const portalRows = rowsAsObjects(
      await sqlite.execute({
        sql: portalSql,
        args: portalArgs,
      }),
    ).map((r: any) => {
      const payload = parseJson(r.payload || "{}");
      const labelMap: Record<string, string> = {
        portal_opened: "Portal Opened",
        magic_link_generated: "Magic Link Generated",
        scheduled: "Scheduled Visit",
        rescheduled: "Rescheduled Visit",
        note: "Portal Note",
        no_contact: "No Contact Reported",
        reassign_request: "Reassign Requested",
      };
      return {
        wo_id: r.wo_id,
        tech_id: r.tech_id,
        tech_name: payload.tech_name || "",
        tenant_phone: payload.tenant_phone || "",
        template_used: labelMap[r.action] || r.action,
        message_body: payload.note_text || payload.details || "",
        sent_at: r.created_at,
        rc_message_id: "",
        status: "logged",
        appfolio_noted: Number(payload.note_written ? 1 : 0),
        source_type: "portal_event",
        event_type: r.action,
        event_label: labelMap[r.action] || r.action,
        magic_link: payload.short_url || payload.magic_link || "",
        tenant_name: payload.tenant_name || "",
        address: payload.property_address || "",
        priority: payload.priority || "",
        payload: payload,
      };
    });

    const merged = [...smsRows, ...warningRows, ...portalRows].sort(
      (a: any, b: any) => {
        const at = new Date(a.sent_at || 0).getTime();
        const bt = new Date(b.sent_at || 0).getTime();
        return bt - at;
      },
    ).slice(0, limit);

    return {
      ok: true,
      results: merged,
      count: merged.length,
      tenant_sms_count: smsRows.length,
      warning_count: warningRows.length,
      portal_event_count: portalRows.length,
    };
  } catch (e: any) {
    return { ok: false, error: e.message };
  }
}

// ── Add/remove monitored work orders ────────────────────────────────────────
export async function handleAddMonitoredWO(req: Request): Promise<any> {
  try {
    const body = await req.json();
    const woId = String(body?.wo_id || "").trim();
    if (!woId) return { ok: false, error: "wo_id required" };

    await sqlite.execute({
      sql:
        `INSERT OR IGNORE INTO monitored_work_orders (wo_id, created_at, updated_at) 
            VALUES (?, datetime('now'), datetime('now'))`,
      args: [woId],
    });

    await auditLog(woId, "dispatch_monitored_added", { wo_id: woId });
    return { ok: true, wo_id: woId, action: "added" };
  } catch (e: any) {
    return { ok: false, error: e.message };
  }
}

export async function handleRemoveMonitoredWO(req: Request): Promise<any> {
  try {
    const body = await req.json();
    const woId = String(body?.wo_id || "").trim();
    if (!woId) return { ok: false, error: "wo_id required" };

    await sqlite.execute({
      sql: `DELETE FROM monitored_work_orders WHERE wo_id = ?`,
      args: [woId],
    });

    await auditLog(woId, "dispatch_monitored_removed", { wo_id: woId });
    return { ok: true, wo_id: woId, action: "removed" };
  } catch (e: any) {
    return { ok: false, error: e.message };
  }
}