// ============================================================================
// handlers/magicPortal.ts — Mobile portal for tech-facing work-order actions.
// ============================================================================

import { CORS_HEADERS, PROXY_BASE_URL } from "../config.ts";
import { rowsAsObjects, sqlite } from "../db.ts";
import { verifyMagicToken } from "../lib/auth.ts";

type Branding = {
  name: string;
  logoUrl: string;
};

async function loadBranding(): Promise<Branding> {
  const fallbackName = "Fort Lowell Realty Tech Dispatch";
  let name = "";
  let logoUrl = "";
  try {
    const rows = rowsAsObjects(
      await sqlite.execute(
        `SELECT key, value FROM proxy_config WHERE key IN ('brand_name', 'brand_logo_url', 'portal_brand_name', 'portal_brand_logo_url')`,
      ),
    );
    rows.forEach((row: any) => {
      if (row.key === "brand_name" && row.value) {
        name = name || String(row.value);
      }
      if (row.key === "portal_brand_name" && row.value) {
        name = String(row.value);
      }
      if (row.key === "brand_logo_url" && row.value) {
        logoUrl = logoUrl || String(row.value);
      }
      if (row.key === "portal_brand_logo_url" && row.value) {
        logoUrl = String(row.value);
      }
    });
  } catch {
    // Non-fatal.
  }
  return { name: name || fallbackName, logoUrl };
}

function esc(value: unknown): string {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

// Safely stringify JSON for embedding in HTML <script> tags.
// Escapes </ sequences to prevent script injection.
function escapeJsonForScript(obj: unknown): string {
  return JSON.stringify(obj).replace(/\x3c\x2f/g, "\\x3c\\x2f");
}

function buildErrorHtml(message: string, branding: Branding): string {
  const logo = branding.logoUrl
    ? `<img src="${esc(branding.logoUrl)}" alt="${
      esc(branding.name)
    }" class="brand-logo" />`
    : `<div class="brand-mark">FLR</div>`;
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no">
<meta name="theme-color" content="#0f172a">
<title>${esc(branding.name)} Portal</title>
<style>
*{box-sizing:border-box}body{margin:0;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:radial-gradient(circle at top,#1e293b 0,#0f172a 52%,#020617 100%);color:#e2e8f0;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}.card{width:min(420px,100%);background:rgba(15,23,42,.88);border:1px solid rgba(148,163,184,.2);border-radius:24px;padding:28px;box-shadow:0 24px 64px rgba(2,6,23,.45);text-align:center}.brand-logo{max-width:180px;max-height:64px;object-fit:contain;margin:0 auto 14px;display:block}.brand-mark{width:60px;height:60px;border-radius:18px;background:linear-gradient(135deg,#38bdf8,#0ea5e9);display:flex;align-items:center;justify-content:center;margin:0 auto 14px;font-weight:800;letter-spacing:.08em;color:#082f49}.title{font-size:18px;font-weight:700;margin:0 0 6px}.sub{font-size:12px;color:#94a3b8;margin:0 0 18px}.err{padding:14px 16px;border-radius:16px;background:rgba(239,68,68,.12);border:1px solid rgba(239,68,68,.3);color:#fecaca;line-height:1.6;font-size:14px}
</style>
</head>
<body>
<div class="card">
${logo}
<p class="title">${esc(branding.name)}</p>
<p class="sub">Tech Dispatch Portal</p>
<div class="err">${esc(message)}</div>
</div>
</body>
</html>`;
}

function buildPortalHtml(ctx: {
  branding: Branding;
  token: string;
  proxyBaseUrl: string;
  initialContext: Record<string, any>;
}): string {
  const logo = ctx.branding.logoUrl
    ? `<img src="${esc(ctx.branding.logoUrl)}" alt="${
      esc(ctx.branding.name)
    }" class="brand-logo" />`
    : `<div class="brand-mark">FLR</div>`;

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no,viewport-fit=cover">
<meta name="theme-color" content="#0f172a">
<title>${esc(ctx.branding.name)} Portal</title>
<style>
:root{--bg:#020617;--bg2:#0f172a;--card:rgba(15,23,42,.86);--card2:rgba(30,41,59,.75);--line:rgba(148,163,184,.18);--text:#e2e8f0;--muted:#94a3b8;--accent:#38bdf8;--accent2:#0ea5e9;--ok:#22c55e;--warn:#f59e0b;--bad:#ef4444;--radius:22px;--shadow:0 20px 60px rgba(2,6,23,.45)}
*{box-sizing:border-box}html,body{margin:0;min-height:100%;background:linear-gradient(180deg,#0f172a 0,#020617 100%);color:var(--text);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif}body{padding:18px 14px 28px}.shell{width:min(720px,100%);margin:0 auto}.hero{background:linear-gradient(160deg,rgba(56,189,248,.18),rgba(15,23,42,.92) 34%,rgba(15,23,42,.96));border:1px solid rgba(56,189,248,.18);border-radius:28px;padding:18px 18px 16px;box-shadow:var(--shadow);margin-bottom:14px}.hero-top{display:flex;justify-content:space-between;gap:12px;align-items:flex-start}.brand-wrap{display:flex;gap:12px;align-items:center;min-width:0}.brand-logo{width:auto;max-width:152px;max-height:52px;object-fit:contain;display:block}.brand-mark{width:52px;height:52px;border-radius:16px;background:linear-gradient(135deg,var(--accent),var(--accent2));display:flex;align-items:center;justify-content:center;font-weight:900;color:#082f49;letter-spacing:.08em;flex-shrink:0}.brand-copy{min-width:0}.brand-name{font-size:17px;font-weight:800;line-height:1.2}.brand-sub{font-size:12px;color:var(--muted);margin-top:3px}.lang-toggle{display:flex;gap:6px}.lang-btn{border:1px solid var(--line);background:rgba(15,23,42,.62);color:var(--muted);padding:8px 10px;border-radius:999px;font-size:12px;font-weight:700;cursor:pointer}.lang-btn.active{color:#fff;background:rgba(56,189,248,.18);border-color:rgba(56,189,248,.5)}.hero-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:14px}.hero-stat{padding:12px;border-radius:18px;background:rgba(15,23,42,.55);border:1px solid var(--line)}.hero-label{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.08em;margin-bottom:4px}.hero-value{font-size:14px;font-weight:700;line-height:1.4}.stack{display:grid;gap:12px}.card{background:var(--card);border:1px solid var(--line);border-radius:24px;padding:16px;box-shadow:var(--shadow)}.card-title{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:12px}.card-title h2{margin:0;font-size:15px}.card-hint{font-size:11px;color:var(--muted)}.grid{display:grid;gap:10px}.grid.cols-2{grid-template-columns:1fr 1fr}.field label{display:block;font-size:11px;color:var(--muted);margin-bottom:6px;text-transform:uppercase;letter-spacing:.06em}.input,.select,.textarea{width:100%;border:1px solid var(--line);background:rgba(15,23,42,.62);color:var(--text);border-radius:16px;padding:12px 13px;font-size:15px}.textarea{min-height:108px;resize:vertical}.actions{display:flex;gap:8px;flex-wrap:wrap}.btn{border:none;border-radius:16px;padding:12px 14px;font-size:14px;font-weight:700;cursor:pointer}.btn.primary{background:linear-gradient(135deg,var(--accent),var(--accent2));color:#082f49}.btn.secondary{background:rgba(30,41,59,.9);color:var(--text);border:1px solid var(--line)}.btn.warning{background:rgba(245,158,11,.16);color:#fde68a;border:1px solid rgba(245,158,11,.35)}.btn.danger{background:rgba(239,68,68,.14);color:#fecaca;border:1px solid rgba(239,68,68,.35)}.btn.ghost{background:transparent;color:var(--muted);border:1px dashed var(--line)}.btn[disabled]{opacity:.55;cursor:not-allowed}.row{display:flex;justify-content:space-between;gap:8px;padding:7px 0;border-bottom:1px solid rgba(148,163,184,.1)}.row:last-child{border-bottom:none}.row-label{font-size:12px;color:var(--muted)}.row-value{font-size:13px;text-align:right;max-width:68%;word-break:break-word}.status{display:none;margin-bottom:12px;padding:12px 14px;border-radius:16px;font-size:13px;line-height:1.5}.status.show{display:block}.status.ok{background:rgba(34,197,94,.12);border:1px solid rgba(34,197,94,.28);color:#bbf7d0}.status.err{background:rgba(239,68,68,.12);border:1px solid rgba(239,68,68,.28);color:#fecaca}.status.info{background:rgba(56,189,248,.12);border:1px solid rgba(56,189,248,.28);color:#bae6fd}.pill{display:inline-flex;align-items:center;gap:6px;font-size:11px;font-weight:700;padding:5px 9px;border-radius:999px}.pill.ok{background:rgba(34,197,94,.14);color:#bbf7d0}.pill.warn{background:rgba(245,158,11,.16);color:#fde68a}.contact-links{display:grid;gap:8px}.contact-link{display:flex;justify-content:space-between;align-items:center;gap:8px;padding:12px 14px;border-radius:16px;border:1px solid var(--line);background:var(--card2);color:var(--text);text-decoration:none}.contact-link small{display:block;color:var(--muted);font-size:11px;margin-top:2px}.empty-note{font-size:12px;color:var(--muted);line-height:1.5}.footer{font-size:11px;color:var(--muted);text-align:center;margin-top:10px;line-height:1.6}@media (max-width:560px){.hero-grid,.grid.cols-2{grid-template-columns:1fr}.hero-top{flex-direction:column}.lang-toggle{align-self:flex-end;width:100%;justify-content:flex-end}.row{flex-direction:column}.row-value{text-align:left;max-width:100%}}
/* ── Accordion ─────────────────────────────────────────────────────────── */
.portal-accordion{border-top:1px solid var(--line);margin-top:4px}
.acc-item{border-bottom:1px solid var(--line)}
.acc-trigger{width:100%;display:flex;align-items:center;gap:12px;padding:16px;background:transparent;border:none;cursor:pointer;color:var(--text);font-size:15px;font-weight:600;text-align:left;-webkit-tap-highlight-color:transparent;transition:background .15s;border-radius:0}
.acc-trigger[aria-expanded="true"]{background:rgba(56,189,248,.06)}
.acc-trigger:active{background:rgba(148,163,184,.06)}
.acc-trigger:disabled{opacity:.5;cursor:not-allowed}
.acc-ico{font-size:18px;width:24px;text-align:center;flex-shrink:0}
.acc-lbl{flex:1}
.acc-chev{margin-left:auto;font-size:20px;line-height:1;color:var(--muted);transition:transform .2s ease}
.acc-body{background:rgba(2,6,23,.28)}
.acc-body-inner{padding:14px 16px 20px;display:flex;flex-direction:column}
.acc-fld{display:flex;flex-direction:column;gap:6px;margin-top:14px}
.acc-fld:first-child{margin-top:0}
.acc-fld label{font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.06em}
.acc-actions{margin-top:18px;display:flex;flex-direction:column;gap:10px}
.acc-actions .btn{width:100%}
.acc-hint{font-size:12px;color:var(--muted);line-height:1.4}
.char-count{font-size:11px;color:var(--muted);text-align:right}
.acc-contact{background:var(--card2);border:1px solid var(--line);border-radius:14px;padding:12px 14px}
.acc-contact+.acc-contact{margin-top:10px}
.acc-contact-role{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--muted)}
.acc-contact-name{font-size:15px;font-weight:700;color:var(--text);margin-top:4px}
.acc-contact-link{display:block;font-size:14px;color:#7dd3fc;text-decoration:none;font-weight:600;margin-top:4px}
.acc-contact-link:hover{text-decoration:underline}
</style>
</head>
<body>
<div class="shell">
  <div class="hero">
    <div class="hero-top">
      <div class="brand-wrap">
        ${logo}
        <div class="brand-copy">
          <div class="brand-name" id="brandName">${esc(ctx.branding.name)}</div>
          <div class="brand-sub" id="brandSub">Tech Dispatch Portal</div>
        </div>
      </div>
      <div class="lang-toggle">
        <button class="lang-btn" id="langEn" type="button">EN</button>
        <button class="lang-btn" id="langEs" type="button">ES</button>
      </div>
    </div>
    <div id="heroStatus" class="status"></div>
    <div class="hero-grid">
      <div class="hero-stat">
        <div class="hero-label" id="labelWo">Work Order</div>
        <div class="hero-value" id="woValue">—</div>
      </div>
      <div class="hero-stat">
        <div class="hero-label" id="labelTech">Technician</div>
        <div class="hero-value" id="techValue">—</div>
      </div>
      <div class="hero-stat">
        <div class="hero-label" id="labelAddress">Address</div>
        <div class="hero-value" id="addressValue">—</div>
      </div>
      <div class="hero-stat">
        <div class="hero-label" id="labelTenant">Resident</div>
        <div class="hero-value" id="tenantValue">—</div>
      </div>
    </div>
  </div>

  <div class="stack">
    <section class="card">
      <div class="card-title">
        <h2 id="overviewTitle">Overview</h2>
        <span class="pill ok" id="scheduleBadge" style="display:none"></span>
      </div>
      <div class="grid">
        <div class="row"><div class="row-label" id="descLabel">Description</div><div class="row-value" id="woDesc">—</div></div>
        <div class="row"><div class="row-label" id="accessLabel">Access</div><div class="row-value" id="accessNotes">—</div></div>
        <div class="row"><div class="row-label" id="vendorLabel">Vendor Notes</div><div class="row-value" id="vendorNotes">—</div></div>
      </div>
    </section>

    <section class="card">
      <div class="card-title">
        <h2 id="messagesTitle">Resident Messages</h2>
        <span class="card-hint" id="messagesHint">Single-use SMS actions</span>
      </div>
      <div class="actions">
        <button class="btn secondary portal-message-btn" data-template="enroute" type="button" id="btnEnroute">I’m On My Way</button>
        <button class="btn secondary portal-message-btn" data-template="schedule" type="button" id="btnScheduleText">Let’s Schedule a Visit</button>
        <button class="btn secondary portal-message-btn" data-template="today" type="button" id="btnToday">Arriving Today</button>
      </div>
      <div class="footer" id="messageFooter">Sending one marks tenant messaging complete. Schedule/notes/reassign actions remain available.</div>
    </section>

    <section class="card">
      <div class="card-title">
        <h2 id="actionsTitle">Portal Actions</h2>
        <span class="card-hint" id="actionsHint">Tap a section to expand</span>
      </div>
      <div class="portal-accordion" id="portal-accordion">

        <div class="acc-item">
          <button class="acc-trigger" data-target="acc-schedule" aria-expanded="false" type="button">
            <span class="acc-ico">📅</span>
            <span class="acc-lbl" id="accScheduleLabel">Schedule Work</span>
            <span class="acc-chev">›</span>
          </button>
          <div class="acc-body" id="acc-schedule" hidden>
            <div class="acc-body-inner">
              <div class="acc-fld">
                <label for="sched-date" id="scheduledDateLabel">Select Date</label>
                <input type="date" id="sched-date" class="input" min="" />
              </div>
              <div class="acc-fld">
                <label for="sched-window" id="scheduledWindowLabel">Arrival Window</label>
                <select id="sched-window" class="select">
                  <option value="">Select…</option>
                  <option value="Morning (8am–11am)">Morning (8am – 11am)</option>
                  <option value="Mid-Day (11am–1pm)">Mid-Day (11am – 1pm)</option>
                  <option value="Afternoon (1pm–4pm)">Afternoon (1pm – 4pm)</option>
                  <option value="Late Afternoon (4pm–6pm)">Late Afternoon (4pm – 6pm)</option>
                </select>
              </div>
              <div class="acc-actions">
                <button class="btn primary portal-action-btn" id="btn-schedule" type="button">Confirm Schedule</button>
                <button class="btn secondary portal-action-btn" id="btn-reschedule" type="button">Reschedule</button>
                <p class="acc-hint" id="scheduleHint">This will update work order status and notify dispatch.</p>
              </div>
            </div>
          </div>
        </div>

        <div class="acc-item">
          <button class="acc-trigger" data-target="acc-status" aria-expanded="false" type="button">
            <span class="acc-ico">🛠️</span>
            <span class="acc-lbl" id="accStatusLabel">Update Work Order</span>
            <span class="acc-chev">›</span>
          </button>
          <div class="acc-body" id="acc-status" hidden>
            <div class="acc-body-inner">
              <div class="acc-fld">
                <label for="status-select" id="statusLabel">Status</label>
                <select id="status-select" class="select">
                  <option value="">Select…</option>
                  <option value="Waiting">Waiting</option>
                  <option value="Work Completed">Work Completed</option>
                </select>
              </div>
              <div class="acc-fld">
                <label for="status-note" id="statusNoteLabel">Completion / Exception Note</label>
                <textarea id="status-note" class="textarea" maxlength="1200" placeholder="Summarize what was completed, what is blocked, or what dispatch should know."></textarea>
              </div>
              <div class="acc-actions">
                <button class="btn primary portal-action-btn" id="btn-status" type="button">Save Status Update</button>
                <p class="acc-hint" id="statusHint">Writes an AppFolio status update and a work-order note.</p>
              </div>
            </div>
          </div>
        </div>

        <div class="acc-item">
          <button class="acc-trigger" data-target="acc-note" aria-expanded="false" type="button">
            <span class="acc-ico">📝</span>
            <span class="acc-lbl" id="accNoteLabel">Leave a Note</span>
            <span class="acc-chev">›</span>
          </button>
          <div class="acc-body" id="acc-note" hidden>
            <div class="acc-body-inner">
              <div class="acc-fld">
                <label for="note-text" id="portalNoteLabel">Your Message</label>
                <textarea id="note-text" class="textarea" maxlength="1000" placeholder="e.g. Parts ordered, returning Thursday..."></textarea>
                <div class="char-count"><span id="note-chars">0</span>/1000</div>
              </div>
              <div class="acc-actions">
                <button class="btn primary portal-action-btn" id="btn-note" type="button">Submit Note</button>
              </div>
            </div>
          </div>
        </div>

        <div class="acc-item">
          <button class="acc-trigger" data-target="acc-contacts" aria-expanded="false" type="button">
            <span class="acc-ico">👤</span>
            <span class="acc-lbl" id="accContactsLabel">Contact Information</span>
            <span class="acc-chev">›</span>
          </button>
          <div class="acc-body" id="acc-contacts" hidden>
            <div class="acc-body-inner">
              <div class="acc-contact">
                <div class="acc-contact-role" id="contactResidentRole">Resident</div>
                <div class="acc-contact-name" id="ct-tenant-name">Not available</div>
                <a class="acc-contact-link" id="ct-tenant-phone" href="#">Not available</a>
              </div>
              <div class="acc-contact">
                <div class="acc-contact-role" id="contactPmRole">Property Manager</div>
                <div class="acc-contact-name" id="ct-pm-name">Fort Lowell Realty</div>
                <a class="acc-contact-link" id="ct-pm-phone" href="#">Not available</a>
              </div>
              <div class="acc-contact">
                <div class="acc-contact-role" id="contactTechRole">Technician</div>
                <div class="acc-contact-name" id="ct-tech-name">Assigned technician</div>
                <a class="acc-contact-link" id="ct-tech-phone" href="#">Not available</a>
              </div>
            </div>
          </div>
        </div>

        <div class="acc-item">
          <button class="acc-trigger" data-target="acc-photo" aria-expanded="false" type="button">
            <span class="acc-ico">📷</span>
            <span class="acc-lbl" id="accPhotoLabel">Attach Photo</span>
            <span class="acc-chev">›</span>
          </button>
          <div class="acc-body" id="acc-photo" hidden>
            <div class="acc-body-inner">
              <div class="acc-fld">
                <label for="photo-before-file" id="photoBeforeLabel">Before Photo</label>
                <input type="file" id="photo-before-file" accept="image/jpeg,image/png,image/webp" class="input" />
              </div>
              <div class="acc-fld">
                <label for="photo-after-file" id="photoAfterLabel">After Photo</label>
                <input type="file" id="photo-after-file" accept="image/jpeg,image/png,image/webp" class="input" />
              </div>
              <div class="acc-actions">
                <button class="btn primary portal-action-btn" id="btn-photo-before" type="button">Upload Before Photo</button>
                <button class="btn secondary portal-action-btn" id="btn-photo-after" type="button">Upload After Photo</button>
                <p class="acc-hint" id="uploadCaveat">JPEG/PNG/WebP multipart uploads are supported here. If AppFolio rejects a file, the exact error will display above.</p>
              </div>
            </div>
          </div>
        </div>

        <div class="acc-item">
          <button class="acc-trigger" data-target="acc-nocontact" aria-expanded="false" type="button">
            <span class="acc-ico">🚫</span>
            <span class="acc-lbl" id="accNoContactLabel">Can't Reach Resident</span>
            <span class="acc-chev">›</span>
          </button>
          <div class="acc-body" id="acc-nocontact" hidden>
            <div class="acc-body-inner">
              <div class="acc-fld">
                <label for="nocontact-attempts" id="noContactAttemptsLabel">How many times have you tried?</label>
                <select id="nocontact-attempts" class="select">
                  <option value="1">1 attempt</option>
                  <option value="2" selected>2 attempts</option>
                  <option value="3">3 attempts</option>
                  <option value="4">4+ attempts</option>
                </select>
              </div>
              <div class="acc-actions">
                <button class="btn danger portal-action-btn" id="btn-nocontact" type="button">Notify Dispatch</button>
                <p class="acc-hint" id="noContactHint">This will alert dispatch and add a work order note.</p>
              </div>
            </div>
          </div>
        </div>

        <div class="acc-item">
          <button class="acc-trigger" data-target="acc-reassign" aria-expanded="false" type="button">
            <span class="acc-ico">🔄</span>
            <span class="acc-lbl" id="accReassignLabel">Request Reassignment</span>
            <span class="acc-chev">›</span>
          </button>
          <div class="acc-body" id="acc-reassign" hidden>
            <div class="acc-body-inner">
              <div class="acc-fld">
                <label for="reassign-reason" id="reassignReasonLabel">Reason</label>
                <select id="reassign-reason" class="select">
                  <option value="">Select…</option>
                  <option value="Schedule conflict">Schedule conflict</option>
                  <option value="Wrong trade">Wrong trade</option>
                  <option value="Unavailable">Unavailable</option>
                  <option value="Other">Other</option>
                </select>
              </div>
              <div class="acc-fld">
                <label for="reassign-details" id="portalDetailsLabel">Details (optional)</label>
                <textarea id="reassign-details" class="textarea" maxlength="500" placeholder="Any additional context..."></textarea>
              </div>
              <div class="acc-actions">
                <button class="btn warning portal-action-btn" id="btn-reassign" type="button">Submit Request</button>
                <p class="acc-hint" id="reassignHint">Dispatch will be notified immediately and will follow up.</p>
              </div>
            </div>
          </div>
        </div>

      </div>
    </section>
  </div>

  <div class="footer" id="portalFooter">This mobile view uses your existing HandyManager token flow and additive portal APIs.</div>
</div>
<script>
const TOKEN = ${escapeJsonForScript(ctx.token)};
const CONFIGURED_PROXY = ${
    escapeJsonForScript(String(ctx.proxyBaseUrl || "").replace(/\/+$/, ""))
  };
const RUNTIME_PROXY = String(window.location.origin || '') + String(window.location.pathname || '');
const PROXY = (CONFIGURED_PROXY || RUNTIME_PROXY).replace(/\/+$/, '');
const INITIAL_CONTEXT = ${escapeJsonForScript(ctx.initialContext)};
const BRAND = ${escapeJsonForScript(ctx.branding)};
const I18N = {
  en: {
    sub: 'Tech Dispatch Portal', overview: 'Overview', workOrder: 'Work Order', tech: 'Technician', address: 'Address', tenant: 'Resident', description: 'Description', access: 'Access', vendor: 'Vendor Notes', messages: 'Resident Messages', messagesHint: 'Tenant SMS templates', enroute: 'I\'m On My Way', scheduleText: 'Let\'s Schedule a Visit', today: 'Arriving Today', messageFooter: 'Sending one marks tenant messaging complete. Schedule, status, notes, uploads, and reassignment remain available.', scheduleTitle: 'Schedule Work', scheduleHint: 'Updates the portal record and writes an AppFolio note', date: 'Date', window: 'Arrival Window', scheduleBtn: 'Schedule', rescheduleBtn: 'Reschedule', updateTitle: 'Update Work Order', updateHint: 'Status updates, completion notes, and exception reporting', statusLabel: 'Status', statusNoteLabel: 'Completion / Exception Note', statusSave: 'Save Status Update', statusHint: 'Writes an AppFolio status update and a work-order note.', addNote: 'Add Note', saveNote: 'Save Note', attempts: 'No-Contact Attempts', reassignReason: 'Reassignment Reason', details: 'Details', noContact: 'Resident Not Communicating', reassign: 'Request Reassignment', uploadCaveat: 'JPEG/PNG/WebP multipart uploads are supported here. If AppFolio rejects a file, the exact error will display above.', photoBefore: 'Before Photo', photoAfter: 'After Photo', uploadBefore: 'Upload Before Photo', uploadAfter: 'Upload After Photo', contacts: 'Contacts', contactsHint: 'Tap to call or text', footer: 'This mobile view uses your existing HandyManager token flow and additive portal APIs.', scheduledFor: 'Scheduled', noData: '—', pmPhone: 'Property Manager Phone', pmEmail: 'Property Manager Email', residentPhone: 'Resident Phone', residentEmail: 'Resident Email', techPhone: 'Technician Phone', open: 'Open', invalid: 'This portal link is invalid or expired.', success: 'Saved successfully.', network: 'Network error — please try again.'
  },
  es: {
    sub: 'Portal tecnico de despacho', overview: 'Resumen', workOrder: 'Orden de trabajo', tech: 'Tecnico', address: 'Direccion', tenant: 'Residente', description: 'Descripcion', access: 'Acceso', vendor: 'Notas del proveedor', messages: 'Mensajes al residente', messagesHint: 'Plantillas SMS al residente', enroute: 'Voy en camino', scheduleText: 'Programar visita', today: 'Llego hoy', messageFooter: 'Enviar uno marca los mensajes al residente como completados. Programar, estado, notas, fotos y reasignar siguen disponibles.', scheduleTitle: 'Programar trabajo', scheduleHint: 'Actualiza el portal y escribe una nota en AppFolio', date: 'Fecha', window: 'Ventana de llegada', scheduleBtn: 'Programar', rescheduleBtn: 'Reprogramar', updateTitle: 'Actualizar orden', updateHint: 'Cambios de estado, notas de finalizacion y excepciones', statusLabel: 'Estado', statusNoteLabel: 'Nota de finalizacion / excepcion', statusSave: 'Guardar estado', statusHint: 'Escribe un cambio de estado y una nota en la orden.', addNote: 'Agregar nota', saveNote: 'Guardar nota', attempts: 'Intentos sin contacto', reassignReason: 'Motivo de reasignacion', details: 'Detalles', noContact: 'Residente no responde', reassign: 'Solicitar reasignacion', uploadCaveat: 'Se admiten cargas multipart JPEG/PNG/WebP. Si AppFolio rechaza un archivo, el error exacto aparecera arriba.', photoBefore: 'Foto antes', photoAfter: 'Foto despues', uploadBefore: 'Subir foto antes', uploadAfter: 'Subir foto despues', contacts: 'Contactos', contactsHint: 'Toque para llamar o enviar texto', footer: 'Esta vista movil usa el flujo existente de tokens de HandyManager y las APIs aditivas del portal.', scheduledFor: 'Programado', noData: '—', pmPhone: 'Telefono del administrador', pmEmail: 'Correo del administrador', residentPhone: 'Telefono del residente', residentEmail: 'Correo del residente', techPhone: 'Telefono del tecnico', open: 'Abrir', invalid: 'Este enlace del portal no es valido o ya vencio.', success: 'Guardado correctamente.', network: 'Error de red — vuelva a intentar.'
  }
};
const state = { portal: Object.assign({}, INITIAL_CONTEXT || {}), lang: 'en', busy: false, autoOpenedSchedule: false };
function t(key){ var dict=I18N[state.lang]||I18N.en; return dict[key] || I18N.en[key] || key; }
function byId(id){ return document.getElementById(id); }
function setText(id, value){ var el = byId(id); if(el) el.textContent = String(value || ''); }
function setStatus(type, text){ var el=byId('heroStatus'); if(!el) return; if(!text){ el.className='status'; el.textContent=''; return; } el.className='status show '+type; el.textContent=text; }
function safe(v){ return String(v || '').trim(); }
function buildPortalActionUrl(action){ return PROXY + '?action=' + encodeURIComponent(action); }
function setBusy(next){ state.busy=!!next; Array.prototype.forEach.call(document.querySelectorAll('.portal-action-btn, .portal-message-btn'), function(btn){ btn.disabled=!!next; }); if(state.portal.used){ Array.prototype.forEach.call(document.querySelectorAll('.portal-message-btn'), function(btn){ btn.disabled = true; }); } }
function applyLangButtons(){ byId('langEn').classList.toggle('active', state.lang==='en'); byId('langEs').classList.toggle('active', state.lang==='es'); }
function openAccordion(targetId, doScroll){
  var triggers = document.querySelectorAll('.acc-trigger');
  var bodies = document.querySelectorAll('.acc-body');
  var target = byId(targetId);
  if(!target) return;
  Array.prototype.forEach.call(triggers, function(trigger){
    trigger.setAttribute('aria-expanded', 'false');
    var chev = trigger.querySelector('.acc-chev');
    if(chev) chev.style.transform = 'rotate(0deg)';
  });
  Array.prototype.forEach.call(bodies, function(body){ body.hidden = true; });
  target.hidden = false;
  var activeTrigger = document.querySelector('.acc-trigger[data-target="'+targetId+'"]');
  if(activeTrigger){
    activeTrigger.setAttribute('aria-expanded', 'true');
    var activeChev = activeTrigger.querySelector('.acc-chev');
    if(activeChev) activeChev.style.transform = 'rotate(90deg)';
  }
  if(doScroll){
    setTimeout(function(){ target.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); }, 60);
  }
}
function initAccordion(){
  var triggers = document.querySelectorAll('.acc-trigger');
  Array.prototype.forEach.call(triggers, function(trigger){
    trigger.addEventListener('click', function(){
      var targetId = trigger.getAttribute('data-target');
      var body = targetId ? byId(targetId) : null;
      if(!targetId || !body) return;
      var isOpen = !body.hidden;
      if(isOpen){
        trigger.setAttribute('aria-expanded', 'false');
        var chev = trigger.querySelector('.acc-chev');
        if(chev) chev.style.transform = 'rotate(0deg)';
        body.hidden = true;
        return;
      }
      openAccordion(targetId, true);
    });
  });

  var schedDate = byId('sched-date');
  if(schedDate){
    var today = new Date();
    var mm = String(today.getMonth()+1).padStart(2,'0');
    var dd = String(today.getDate()).padStart(2,'0');
    schedDate.min = today.getFullYear() + '-' + mm + '-' + dd;
  }

  var noteText = byId('note-text');
  var noteChars = byId('note-chars');
  if(noteText && noteChars){
    noteText.addEventListener('input', function(){ noteChars.textContent = String(noteText.value.length); });
    noteChars.textContent = String(noteText.value.length);
  }
}
function render(){
  var p=state.portal||{};
  document.documentElement.lang = state.lang;
  applyLangButtons();
  setText('brandName', BRAND.name || 'Fort Lowell Realty');
  setText('brandSub', t('sub'));
  setText('labelWo', t('workOrder'));
  setText('labelTech', t('tech'));
  setText('labelAddress', t('address'));
  setText('labelTenant', t('tenant'));
  setText('overviewTitle', t('overview'));
  setText('descLabel', t('description'));
  setText('accessLabel', t('access'));
  setText('vendorLabel', t('vendor'));
  setText('messagesTitle', t('messages'));
  setText('messagesHint', t('messagesHint'));
  setText('btnEnroute', t('enroute'));
  setText('btnScheduleText', t('scheduleText'));
  setText('btnToday', t('today'));
  setText('messageFooter', t('messageFooter'));
  setText('accStatusLabel', t('updateTitle'));
  setText('statusLabel', t('statusLabel'));
  setText('statusNoteLabel', t('statusNoteLabel'));
  setText('btn-status', t('statusSave'));
  setText('statusHint', t('statusHint'));
  setText('scheduledDateLabel', t('date'));
  setText('scheduledWindowLabel', t('window'));
  setText('portalNoteLabel', t('addNote'));
  setText('btn-note', t('saveNote'));
  setText('noContactAttemptsLabel', t('attempts'));
  setText('reassignReasonLabel', t('reassignReason'));
  setText('portalDetailsLabel', t('details'));
  setText('uploadCaveat', t('uploadCaveat'));
  setText('photoBeforeLabel', t('photoBefore'));
  setText('photoAfterLabel', t('photoAfter'));
  setText('btn-photo-before', t('uploadBefore'));
  setText('btn-photo-after', t('uploadAfter'));
  setText('portalFooter', t('footer'));
  setText('woValue', safe(p.wo_number) || safe(p.wo_id) || t('noData'));
  setText('techValue', safe(p.tech_name) || t('noData'));
  setText('addressValue', safe(p.property_address) || t('noData'));
  setText('tenantValue', safe(p.tenant_name) || t('noData'));
  setText('woDesc', safe(p.wo_description) || t('noData'));
  setText('accessNotes', safe(p.wo_access_notes) || t('noData'));
  setText('vendorNotes', safe(p.vendor_notes) || t('noData'));
  if (p.scheduled_date && byId('sched-date') && byId('sched-date').value !== p.scheduled_date) byId('sched-date').value = p.scheduled_date;
  if (p.scheduled_window && byId('sched-window')) byId('sched-window').value = p.scheduled_window;
  var badge = byId('scheduleBadge');
  if (p.scheduled_date) { badge.style.display='inline-flex'; badge.textContent = t('scheduledFor')+': '+p.scheduled_date+(p.scheduled_window?' · '+p.scheduled_window:''); }
  else { badge.style.display='none'; badge.textContent=''; }
  setText('ct-tenant-name', safe(p.tenant_name) || 'Not available');
  var tenantPhone = byId('ct-tenant-phone');
  if(tenantPhone){
    var phone = safe(p.tenant_phone);
    tenantPhone.textContent = phone || 'Not available';
    tenantPhone.href = phone ? ('tel:' + phone) : '#';
  }
  setText('ct-pm-name', safe(p.pm_name) || 'Fort Lowell Realty');
  var pmPhone = byId('ct-pm-phone');
  if(pmPhone){
    var pm = safe(p.pm_phone);
    pmPhone.textContent = pm || 'Not available';
    pmPhone.href = pm ? ('tel:' + pm) : '#';
  }
  setText('ct-tech-name', safe(p.tech_name) || 'Assigned technician');
  var techPhone = byId('ct-tech-phone');
  if(techPhone){
    var tech = safe(p.tech_phone);
    techPhone.textContent = tech || 'Not available';
    techPhone.href = tech ? ('tel:' + tech) : '#';
  }
  if (p.scheduled_date && !state.autoOpenedSchedule) {
    openAccordion('acc-schedule', false);
    state.autoOpenedSchedule = true;
  }
  if (!p.scheduled_date) state.autoOpenedSchedule = false;
  Array.prototype.forEach.call(document.querySelectorAll('.portal-message-btn'), function(btn){ btn.disabled = !!p.used || state.busy; });
}
async function api(action, body){
  var resp = await fetch(buildPortalActionUrl(action), { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(Object.assign({ token:TOKEN, lang_pref: state.lang }, body || {})) });
  var data = await resp.json().catch(function(){ return { ok:false, error:t('network') }; });
  if(!data.ok) throw new Error(data.error || t('network'));
  return data;
}
async function refreshPortal(){
  try{
    var data = await api('portal_validate', {});
    if(data && data.portal){ state.portal = Object.assign({}, state.portal, data.portal); render(); }
  }catch(e){ setStatus('err', e.message || t('invalid')); }
}
async function runAction(kind, fn){
  if(state.busy) return;
  setBusy(true); setStatus('info', kind);
  try{ await fn(); setStatus('ok', t('success')); await refreshPortal(); }
  catch(e){ setStatus('err', e.message || t('network')); }
  finally{ setBusy(false); render(); }
}
async function uploadPortalPhoto(phase, inputId){
  var input = byId(inputId);
  var file = input && input.files && input.files[0] ? input.files[0] : null;
  if(!file) throw new Error('Select a photo before uploading.');
  if(file.size > 10 * 1024 * 1024) throw new Error('Photo exceeds 10MB limit.');
  var resp = await fetch(buildPortalActionUrl('portal_photo_upload') + '&token=' + encodeURIComponent(TOKEN) + '&phase=' + encodeURIComponent(phase), {
    method:'POST',
    headers:{ 'Content-Type': file.type || 'application/octet-stream' },
    body: file
  });
  var data = await resp.json().catch(function(){ return { ok:false, error:t('network') }; });
  if(!data.ok) throw new Error(data.error || t('network'));
  if(input) input.value='';
  return data;
}
function bindClick(id, handler){ var el = byId(id); if(el) el.addEventListener('click', handler); }
bindClick('langEn', function(){ state.lang='en'; try{localStorage.setItem('hm_portal_lang','en');}catch(_){ } render(); });
bindClick('langEs', function(){ state.lang='es'; try{localStorage.setItem('hm_portal_lang','es');}catch(_){ } render(); });
Array.prototype.forEach.call(document.querySelectorAll('.portal-message-btn'), function(btn){ btn.addEventListener('click', function(){ var template = btn.getAttribute('data-template') || ''; runAction('Sending…', async function(){ var data = await api('send_tenant_sms', { template: template }); if(data && data.message) setStatus('ok', data.message); if(data && data.already_sent){ state.portal.used = true; state.portal.used_template = data.template || state.portal.used_template || template; return; } state.portal.used = true; state.portal.used_template = data.template || template; }); }); });
bindClick('btn-schedule', function(){ runAction('Saving schedule…', async function(){ await api('portal_schedule', { scheduled_date: byId('sched-date').value, scheduled_window: byId('sched-window').value }); }); });
bindClick('btn-reschedule', function(){ runAction('Saving reschedule…', async function(){ await api('portal_reschedule', { scheduled_date: byId('sched-date').value, scheduled_window: byId('sched-window').value }); }); });
bindClick('btn-status', function(){ runAction('Saving status…', async function(){ await api('portal_status', { status: byId('status-select').value, note_text: byId('status-note').value }); if(byId('status-note')) byId('status-note').value=''; if(byId('status-select')) byId('status-select').value=''; }); });
bindClick('btn-note', function(){ runAction('Saving note…', async function(){ await api('portal_note', { note_text: byId('note-text').value }); byId('note-text').value=''; var c = byId('note-chars'); if(c) c.textContent='0'; }); });
bindClick('btn-nocontact', function(){ runAction('Sending no-contact report…', async function(){ await api('portal_no_contact', { attempts: byId('nocontact-attempts').value, details: '' }); }); });
bindClick('btn-reassign', function(){ runAction('Submitting reassignment request…', async function(){ await api('portal_reassign_request', { reason: byId('reassign-reason').value, details: byId('reassign-details').value }); }); });
bindClick('btn-photo-before', function(){ runAction('Uploading before photo…', async function(){ await uploadPortalPhoto('before', 'photo-before-file'); }); });
bindClick('btn-photo-after', function(){ runAction('Uploading after photo…', async function(){ await uploadPortalPhoto('after', 'photo-after-file'); }); });
(function init(){
  try{ var saved = localStorage.getItem('hm_portal_lang'); if(saved==='es'||saved==='en') state.lang=saved; }catch(_){ }
  initAccordion();
  render();
  refreshPortal();
})();
</script>
</body>
</html>`;
}

export async function handleMagicPortal(
  params: Record<string, string>,
): Promise<Response> {
  const htmlHeaders = {
    "Content-Type": "text/html; charset=utf-8",
    ...CORS_HEADERS,
  };
  const branding = await loadBranding();
  const token = params.token || "";

  if (!token) {
    return new Response(
      buildErrorHtml(
        "No token provided. Request a new dispatch link from your coordinator.",
        branding,
      ),
      { status: 400, headers: htmlHeaders },
    );
  }

  const payload = await verifyMagicToken(token);
  if (!payload) {
    return new Response(
      buildErrorHtml(
        "This link has expired or is invalid. Request a new dispatch notification from your coordinator.",
        branding,
      ),
      { status: 401, headers: htmlHeaders },
    );
  }

  let tokenRow: any = {};
  try {
    tokenRow = rowsAsObjects(
      await sqlite.execute({
        sql:
          `SELECT used, used_template, scheduled_date, scheduled_window, lang_pref,
                   property_address, portal_opened, portal_opened_at, meta_json
              FROM magic_link_tokens
             WHERE token = ?
             LIMIT 1`,
        args: [token],
      }),
    )[0] || {};
  } catch {
    tokenRow = {};
  }

  let meta: Record<string, any> = {};
  try {
    meta = JSON.parse(String(tokenRow.meta_json || "{}"));
  } catch {
    meta = {};
  }

  const initialContext = {
    ...meta,
    ...payload,
    used: !!tokenRow.used,
    used_template: tokenRow.used_template || "",
    scheduled_date: tokenRow.scheduled_date || meta.scheduled_date || "",
    scheduled_window: tokenRow.scheduled_window || meta.scheduled_window || "",
    lang_pref: tokenRow.lang_pref || meta.lang_pref || payload.lang_pref ||
      "en",
    property_address: payload.property_address || tokenRow.property_address ||
      meta.property_address || "",
  };

  return new Response(
    buildPortalHtml({
      branding,
      token,
      proxyBaseUrl: PROXY_BASE_URL,
      initialContext,
    }),
    { status: 200, headers: htmlHeaders },
  );
}