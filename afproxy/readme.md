# HandyManager Proxy

Val Town / Deno proxy for the HandyManager frontend.

This service sits between the browser UI and AppFolio, adds authentication server-side, handles caching, exposes normalized action endpoints, and hosts the v9 dispatch / webhook tooling.

## Webapp Feature List

The HandyManager webapp currently includes the following major features.

### Access and session

- Vault-style unlock screen with passphrase entry
- AppFolio subdomain input and proxy URL input
- Auto-lock countdown and manual lock action
- Light/dark theme toggle
- Role-aware UI controls (including vendor-only restrictions)

### Data connectivity and reliability

- Proxy-backed API routing for AppFolio Database API v0 and Reports API v2
- Top-bar live status indicators (API status, rate, cache state)
- IndexedDB local caching for key datasets
- Cache export/import actions for offline recovery
- Retry/backoff and timeout handling in client fetch wrappers
- CORS/connection troubleshooting banner with guided fixes

### Global filtering and navigation

- Multi-tab navigation for dashboard and operations views
- Global property-group filter applied across tabs
- Group reload action and active-filter indicators

### Dashboard and operations overview

- KPI cards for open WOs, urgent WOs, turns, move-outs, and flagged items
- Attention panel for stalled turns, overdue inspections, and vendor alerts
- Upcoming move-outs table with urgency display
- Activity feed with category filters and webhook-enriched events

### Work orders

- Kanban board by status with priority/type/property filtering
- Search across work-order records
- Flag/unflag follow-up markers persisted locally
- Work-order detail modal (status, priority, notes, tenant, vendor, property)
- AppFolio deep-link out to the source work order
- Add-note support through proxy-backed v0 notes endpoint
- Live note refresh when new related webhook events arrive
- New work-order creation modal

### Payroll

- Friday-to-Friday payroll period navigation
- Payroll KPIs (count, totals, vendors, properties)
- Work-done table with click-through into work-order detail

### Turn board and inspections

- Turn pipeline stages (UPC, MO, INS, WO, REQ, EST, ASN, DONE)
- Turn filtering (active, on radar, upcoming, stalled, completed)
- Turn search and group scoping
- Detailed turn timeline modal with linked work orders/events
- Inspections dashboard with overdue/due-soon/current/turn-linked KPIs
- Active-property + rolling-window inspection filtering
- Sortable inspections table

### Vendors and templates

- Vendor directory with search, category filter, and compliance visibility
- Vendor override controls (category/compliance)
- Communication templates tab with preview/copy/edit actions

### Dispatch control

- Dispatch cockpit with live branch-scoped controls
- Queue, roster, grades, config, audit, blasts, and comms subpanels
- Manual warning/reassign cron triggers
- Pause/resume automation controls
- Assignee sync from AppFolio users and work-order activity
- Tech roster CRUD and hidden-assignee support
- Tenant communications log feed
- Test magic-link SMS send flow from UI

### Webhooks and live feed

- Webhook configuration modal with endpoint copy support
- Poll controls and event preview list
- Live event drawer with unseen badge and resolve actions
- Decoded webhook event labels, severity, and selective toasts
- Event-driven UI invalidation/refresh wiring across sections

### Database admin and diagnostics

- In-app SQL admin panel with shortcuts and result rendering
- API error log tab with retry/resolution visibility
- Proxy health/cache/debug endpoints surfaced in documentation

## Purpose

- Proxy AppFolio Database API v0 and Reports API v2 requests without exposing credentials to the browser
- Cache expensive datasets in Turso or Val Town SQLite
- Accept and store inbound AppFolio webhooks
- Support turn tracking, vendor/work-order/property data, and inspection reporting
- Provide dispatch control features such as reassignment queue, tech roster sync, magic links, and tenant SMS

## Runtime

- Entry point: [main.ts](/workspaces/handymgr2/afproxy/main.ts)
- Config and env vars: [config.ts](/workspaces/handymgr2/afproxy/config.ts)
- Database and cache layer: [db.ts](/workspaces/handymgr2/afproxy/db.ts)
- Shared utilities: [lib](/workspaces/handymgr2/afproxy/lib)
- Action handlers: [handlers](/workspaces/handymgr2/afproxy/handlers)

## Architecture

1. [main.ts](/workspaces/handymgr2/afproxy/main.ts) parses the request, enforces CORS and cron auth, then routes by `action`.
2. Handler modules fetch AppFolio data, read or write cache/database state, and return JSON-safe payloads.
3. [handlers/passthrough.ts](/workspaces/handymgr2/afproxy/handlers/passthrough.ts) handles raw API proxying and admin/cache endpoints.
4. [db.ts](/workspaces/handymgr2/afproxy/db.ts) stores cached API responses, webhook events, turn records, and v9 dispatch tables.

## Main Endpoint Groups

### Health and diagnostics

- `GET ?action=ping`
- `GET ?action=cache_stats`
- `GET ?action=debug_sqlite`
- `GET ?action=storage_cleanup`

### Core data

- `GET ?action=work_orders&days=180`
- `GET ?action=wo_notes&wo_id=UUID`
- `GET ?action=wo_detail&wo_id=UUID`
- `GET ?action=wo_billed_amount&wo_number=12345`
- `GET ?action=vendors`
- `GET ?action=properties`
- `GET ?action=property_groups`
- `GET ?action=property_map`
- `GET ?action=upcoming_moveouts&days=60`
- `GET ?action=turns&days=90`
- `GET ?action=turn_work_orders&days=90`
- `GET ?action=work_orders_completed_history&days=365`
- `GET ?action=completed_work_orders_history&days=365` (compat alias)
- `GET ?action=inspections&days=180`
- `GET ?action=bills&days=180`
- `GET ?action=bills_stats`
- `GET ?action=bills_history&days=365`
- `GET ?action=bill_detail&bill_id=UUID`
- `GET ?action=bill_attachments&bill_id=UUID`
- `GET ?action=recent_tasks`

### Turn tracker and reports

- `GET ?action=turn_records`
- `GET ?action=unit_turns&days=180&limit=50&offset=0`
- `GET ?action=unit_turns_history&days=540&limit=300`
- `GET ?action=closed_turns`
- `POST ?action=turn_records`
- `POST ?action=turn_record_stage`
- `POST ?action=unit_turns_sync`
- `POST ?action=unit_turn_wo_link`
- `POST ?action=unit_turn_wo_unlink`
- `GET ?action=wo_comparison_report`
- `POST ?action=report&name=REPORT_NAME`
- `GET ?action=report&name=REPORT_NAME`

### Webhooks

- `POST` with no `action` for AppFolio webhook delivery
- `POST ?action=webhook` as a legacy alias
- `GET ?action=webhook_events`
- `GET ?action=webhook_live`
- `GET ?action=webhook_stats`
- `GET ?action=webhook_resolve`
- `GET ?action=webhook_migrate`
- `GET ?action=webhook_cleanup`

### Raw passthrough and admin

- `GET ?action=passthrough&path=/api/v0/...`
- `POST ?action=passthrough&path=/api/v2/reports/...`
- `GET ?path=/api/v0/...` compatibility mode
- `POST ?action=sql_query` with body `{ key, query }`
- `POST ?action=sql_execute` with body `{ key, sql, args? }`

### Auth and admin lists

- `GET ?action=trusted_devices&key=PROXY_ADMIN_KEY&limit=100&offset=0`
- `GET ?action=pm_proxy_users&key=PROXY_ADMIN_KEY&limit=100&offset=0`
- `GET ?action=settings_get&key=PROXY_ADMIN_KEY&limit=200&offset=0`
- `POST ?action=pm_proxy_user_upsert`
- `POST ?action=pm_proxy_user_delete`
- `POST ?action=settings_set`
- `POST ?action=verify_role`

### Dispatch and communications

- `GET ?action=reassignment_queue`
- `GET/POST ?action=tech_roster`
- `GET ?action=tenant_comms_log`
- `GET ?action=portal&token=TOKEN`
- `POST ?action=portal_validate`
- `POST ?action=portal_schedule`
- `POST ?action=portal_reschedule`
- `POST ?action=portal_note`
- `POST ?action=portal_no_contact`
- `POST ?action=portal_reassign_request`
- `POST ?action=generate_magic_link`
- `POST ?action=add_monitored_work_order`
- `POST ?action=remove_monitored_work_order`
- `POST ?action=send_tenant_sms`
- `POST ?action=send_magic_link_test_sms` with body `{ key, phone, tech_name?, tech_id? }` (`key` = `PROXY_ADMIN_KEY`)
- `GET/POST ?action=noon_warning_cron`
- `GET/POST ?action=midnight_reassign_cron`
- `GET/POST ?action=dispatch_sync_assignees`
- `GET/POST ?action=dispatch_seed_reassignment_test`

## Changelog

### 2026-03-24

- Added `send_magic_link_test_sms` so RingCentral + magic-link delivery can be verified by sending a real test portal link to a supplied phone number.
- Updated dispatch assignee sync to pull from AppFolio `v0 /users` and merge that with work-order assignment activity for branch mapping.
- Documented `webhook_live` explicitly in the public endpoint list.
- Standardized README endpoint descriptions around AppFolio Database API `v0` and Reports API `v2` terminology.

## Environment Variables

Required for AppFolio access:

- `AF_V2_CLIENT_ID`
- `AF_V2_CLIENT_SECRET`
- `AF_DEVELOPER_ID`
- `MAGIC_LINK_SECRET`

Optional / feature-specific:

- `AF_V0_AUTH`
- `FRONTEND_PROXY_SECRET`
- `PROXY_ADMIN_KEY`
- `TURSO_DATABASE_URL`
- `TURSO_AUTH_TOKEN`
- `RC_CLIENT_ID`
- `RC_CLIENT_SECRET`
- `RC_JWT`
- `RC_FROM_NUMBER`
- `PROXY_BASE_URL`
- `CRON_SECRET`
- `ADMIN_EMAIL`

Notes:

- `send_magic_link_test_sms` requires `PROXY_ADMIN_KEY` in request body (`key`) to prevent non-admin SMS abuse.
- Cron endpoints deny access when `CRON_SECRET` is unset.
See [config.ts](/workspaces/handymgr2/afproxy/config.ts) for the exact names and defaults.

## Data Storage

The proxy stores these categories of data in SQLite or Turso:

- API response cache in `api_cache`
- Turn tracker records in `turn_records`
- Webhook event history in `webhook_events`
- Dispatch and communication tables such as `reassignment_queue`, `tech_grades`, `magic_link_tokens`, and `tenant_comms_log`

Schema creation and migrations are handled in [db.ts](/workspaces/handymgr2/afproxy/db.ts) by `ensureTables()`.

## Compatibility Notes

- `?path=/api/...` remains supported as a compatibility passthrough mode.
- `POST ?action=webhook` remains supported as a legacy webhook alias.
- Internal comments may still reference earlier generations of the proxy where behavior was preserved intentionally, but the active router and public contract are v9.

## Local Validation

Type-check the proxy with:

```sh
cd /workspaces/handymgr2/afproxy
deno check main.ts
```

## Key Files

- Router: [main.ts](/workspaces/handymgr2/afproxy/main.ts)
- Config: [config.ts](/workspaces/handymgr2/afproxy/config.ts)
- DB/cache layer: [db.ts](/workspaces/handymgr2/afproxy/db.ts)
- Passthrough and admin endpoints: [handlers/passthrough.ts](/workspaces/handymgr2/afproxy/handlers/passthrough.ts)
- Work order handlers: [handlers/workOrders.ts](/workspaces/handymgr2/afproxy/handlers/workOrders.ts)
- Property handlers: [handlers/properties.ts](/workspaces/handymgr2/afproxy/handlers/properties.ts)
- Turn handlers: [handlers/turns.ts](/workspaces/handymgr2/afproxy/handlers/turns.ts)
- Webhook handlers: [handlers/webhook.ts](/workspaces/handymgr2/afproxy/handlers/webhook.ts)
- Dispatch handlers: [handlers/reassignment.ts](/workspaces/handymgr2/afproxy/handlers/reassignment.ts), [handlers/queue.ts](/workspaces/handymgr2/afproxy/handlers/queue.ts), [handlers/techRoster.ts](/workspaces/handymgr2/afproxy/handlers/techRoster.ts), [handlers/tenantComms.ts](/workspaces/handymgr2/afproxy/handlers/tenantComms.ts)