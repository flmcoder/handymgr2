CHANGELOG

2026-05-04 (v9.7.3)

Bumped proxy runtime version to v9.7.3 and aligned config.ts + main.ts headers.

Bumped frontend runtime version to v9.7.3 so version mismatch checks and startup telemetry stay in sync.

Updated login panel/footer and topbar/sidebar build badges to display v9.7.3 immediately on first paint.

Bumped service worker static cache namespace to hm-static-v5 so deployed clients pick up the new app shell without stale badge text.

Included Turso reliability hardening and cross-tab KPI mini-chart expansion from this release cycle.

Fixed frontend version regression where the topbar build badge reported v9.2.2 while the app runtime was newer.

Bumped frontend app version to v9.2.7 and synchronized release notes references.

Hardened client polling reliability with exponential backoff + jitter in the shared fetch timeout wrapper.

Armored live webhook polling loop to prevent permanent stoppage after browser sleep/network transitions (ERR_NAME_NOT_RESOLVED, ERR_NETWORK_CHANGED, QUIC idle timeout scenarios).

Added wo_attachments and wo_attachment_upload actions for work-order file attachments (list + upload).

Added portal_status action — techs can update WO status (Scheduled / Waiting / Work Completed) from the portal.

Added portal_photo_upload action — before/after photo uploads from the portal, stored as WO attachments.

Portal UI: status update accordion, before/after photo file inputs replacing single disabled button.

Frontend WO detail modal: attachments list, scope/instructions section, permission-to-enter, vendor instructions, related bills/vendor cross-navigation buttons.

Added 429 Retry-After handling to attachment helpers.

Added wo_attachments to GET rate-limit guard alongside wo_detail and wo_notes.

Added units and unit_lookup actions: full paginated fetch of /api/v0/units (90 s timeout, LastUpdatedAtFrom = 5 yr) with Turso units cache table (unit_id PK, property_id, name, status, bedrooms, bathrooms, address, leasing_type, rent_ready, etc.). 6-hour TTL.

Added unitToPropertyId(unitId) helper in lib/groupUtils.ts for resolving UnitId → PropertyId via the units cache — foundation for correct property-group scoping when WO/bills carry only UnitId.

Added unit_id column to billing_map (additive migration) and updated upsertBillingRows to persist UnitId/unit_id from bill records.

Added bills_by_unit route to handlers/bills.ts — server-side pagination of bills filtered by unit_id, with optional group_id scope.

Added unit_id extraction to mapBill so the frontend BILLS array includes unit_id for client-side unit filtering.

Frontend (js/app.js): UNITS global, _unitsByPropertyId lookup, fetchUnits(). Properties section now shows unit count badge per property. Property detail modal adds Units tab (status, bed/bath, per-unit Bills filter button). Billing scope chip updated to show Property › Unit breadcrumb with Clear Unit / Clear All buttons. filterBillsToUnit() added. renderBillsSection and resetBillingFilter updated to respect window.filteredUnitId.

Added property_stats for per-property bill, note, and listing aggregates used by the frontend properties directory.

Added scoped listing retrieval via ?action=property_listings&property_id=... while keeping listings behavior compatible.

Documented and aligned the proxy release to v9.2.3 with the frontend Properties operations update.

Added send_magic_link_test_sms so RingCentral + magic-link delivery can be verified by sending a real test portal link to a supplied phone number.

Updated dispatch assignee sync to pull from AppFolio v0 /users and merge that with work-order assignment activity for branch mapping.

Documented webhook_live explicitly in the public endpoint list.

Standardized README endpoint descriptions around AppFolio Database API v0 and Reports API v2 terminology.
