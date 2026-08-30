# PostgreSQL Cutover Audit and Repair Plan

Date: 2026-08-30

## Required architecture

AppFolio is an upstream ingestion source, not a browser-time data source. Only scheduled sync, webhook, and explicit administrative ingestion workers may call AppFolio. Every user-facing read must query PostgreSQL and include source and freshness metadata.

```mermaid
flowchart LR
  AF[AppFolio v0 and v2] -->|scheduled sync and webhooks| INGEST[Server ingestion workers]
  INGEST --> RAW[(appfolio_raw_responses)]
  INGEST --> PG[(PostgreSQL normalized tables)]
  PG --> API[/api/local routes]
  API --> UI[Web application]
```

## Production findings

Production is available and PostgreSQL is reachable, but the application is degraded rather than fully healthy.

| Check                                                       | Result                                                          | Status                                     |
| ----------------------------------------------------------- | --------------------------------------------------------------- | ------------------------------------------ |
| `/health`                                                   | HTTP 200, database `up`, version `v9.8.5`                       | Pass                                       |
| `/api/local/system_health`                                  | Three green checks                                              | Incomplete health coverage                 |
| Work orders, properties, units, property groups             | PostgreSQL-backed and populated                                 | Pass                                       |
| Turns, vacancies, property map/stats, completed work orders | PostgreSQL-backed and populated                                 | Pass                                       |
| `/api/local/grid/inspections`                               | HTTP 500: `column i.tenant_status does not exist`               | Fail                                       |
| `/api/local/inspections`                                    | Returns `source: appfolio_live_fast`                            | Architecture violation                     |
| `/api/local/bills`                                          | Returns `source: appfolio_db_v0`                                | Architecture violation                     |
| `/api/local/recent_tasks`                                   | Calls AppFolio live and reports API unavailable                 | Fail and architecture violation            |
| `/api/local/manager_review`                                 | Calls three v2 reports live; general ledger rejects its columns | Partial failure and architecture violation |
| Estimates                                                   | PostgreSQL route returns zero rows                              | Not populated                              |
| Upcoming move-outs                                          | PostgreSQL route returns zero rows                              | Needs freshness/population verification    |
| Sync scheduler                                              | Degraded; v2 endpoints paused                                   | Fail                                       |
| SSH tunnel                                                  | Repeated restart failures while PostgreSQL remains reachable    | Misconfiguration/noise                     |

The scheduler currently runs only:

* `v0:properties`

* `v0:property_groups`

* `v0:units`

* `v0:work_orders`

It does not schedule `v0:users` or any v2 populations. The existing v2 runners for inspections, tenant directory, unit turns, vacancies, and estimates are therefore not maintaining their PostgreSQL tables.

The scheduler also reports a failed query for `sync_job_runs.execution_start_cursor`. The column exists in the migration source but is absent from, or incompatible with, the production schema. This is migration drift.

## SQLite and Turso disposition

The Render web service starts `backend/server.ts`; it does not import the `afproxy/` SQLite client. Authentication, PM users, trusted sessions, proxy configuration, queue operations, and operational reads in the active Express process are PostgreSQL-backed.

The remaining SQLite/Turso material is legacy:

* `afproxy/` contains the old Deno/libSQL implementation.

* `syncV0WorkOrdersCron.ts` targets the old proxy-to-Turso sync action.

* `scripts/turso-healthcheck.mjs` and the `db:health` package script test the old database.

* `@libsql/client` remains a production dependency only for the legacy tooling.

* Root `schema.ts` is a stale competing schema and contains SQLite-style `datetime('now')` defaults despite using PostgreSQL Drizzle types.

* Frontend messages and fallback names still refer to SQL/Turso/SQLite.

The legacy health test connected to its configured Turso target but found `work_order_map`, `work_orders_cache`, and `appfolio_work_order_bridge` all empty. It returned `pass: false` and `likely_db_target_mismatch: true`. That database must not be used as a fallback or migrated blindly into PostgreSQL.

## Browser-facing live AppFolio reads to remove

1. The inspection route races a live v2 report and serves it when fast.
2. Bills are fetched from AppFolio v0 inside the misleading `/api/local/bills` route.
3. Recent tasks are fetched live from AppFolio v0.
4. Manager Review fetches tenant tickler, renewal summary, and general ledger reports during each request.
5. Generic v0/v2 passthrough and report proxy routes remain callable by authenticated clients.
6. Work Order notes and attachments fetch AppFolio on cache miss, expiry, or force refresh.
7. Frontend compatibility fallbacks can leave `/api/local` and invoke legacy proxy actions.

AppFolio deep links and server-side mutations may remain. A mutation must enqueue or trigger reconciliation into PostgreSQL before the refreshed state is served.

## Repair plan

### Phase 0: Stabilize PostgreSQL

1. Add an idempotent repair migration for `sync_job_runs.execution_start_cursor` and all columns referenced by active queries.
2. Fix the inspection grid query to use `appfolio_tenant_directory.status`; do not reference `appfolio_unit_inspections.tenant_status` unless that column is deliberately added and populated.
3. Disable the SSH tunnel in Render when `SQL_SE` is the working direct connection. Otherwise repair the tunnel endpoint before relying on it. Never run both paths ambiguously.
4. Add a migration ledger table and startup schema check so a missing migration makes health red before traffic is accepted.

Acceptance gate: sync status is not degraded, inspection grid returns HTTP 200, and Render has no repeating schema or tunnel errors for one scheduler interval.

### Phase 1: Restore and backfill all existing sync populations

1. Enable scheduled `v0:users`.
2. Enable `v2:unit_inspection`, `v2:tenant_directory`, `v2:unit_turn_detail`, and `v2:unit_vacancy` after validating report payload columns. Estimates have an existing PostgreSQL table/upsert helper but no implemented AppFolio report definition, so they require a separate source contract before scheduling.
3. Run a bounded full backfill for each endpoint, then switch to incremental sync with a safety overlap window.
4. Reconcile source counts, distinct keys, null business keys, newest `cached_at`, and property-group coverage.

Acceptance gate: each population has a completed sync run, non-stale `cached_at`, zero unresolved upsert errors, and expected nonzero counts where AppFolio reports data.

### Phase 2: Persist currently live-only datasets

Create normalized tables plus raw JSON for:

* bills and bill attachments

* tasks

* tenant tickler events

* renewal summaries

* general ledger rows required by Manager Review

* Work Order notes and attachments

Add these datasets to the sync allowlist and scheduler. Preserve raw response pages before normalization. Use stable AppFolio IDs as conflict keys and retain deletion/tombstone state where supported.

Acceptance gate: every corresponding `/api/local` response reports `source: postgres_local`; blocking AppFolio egress does not break reads.

### Phase 3: Enforce server-only reads

1. Remove the fast-live branch from inspections.
2. Replace live bills, tasks, Manager Review, notes, and attachments reads with PostgreSQL repositories.
3. Remove empty-result fallback from local reads. Empty is a valid database result; stale/missing data must be represented by freshness metadata and alerts.
4. Restrict generic passthrough/report routes to administrative ingestion diagnostics, or remove them.
5. Return `source`, `synced_at`, `age_seconds`, `stale`, and `sync_run_id` from all population endpoints.

Acceptance gate: an automated route test rejects any browser-facing response whose source starts with `appfolio_live` or `appfolio_db_v0`.

### Phase 4: Retire SQLite/Turso

1. Confirm no deployed service, cron, environment variable, or external Val Town job still targets `afproxy` or the old sync action.
2. Archive or delete `afproxy/`, `syncV0WorkOrdersCron.ts`, `scripts/turso-healthcheck.mjs`, and the stale root schema after verifying no external consumer remains.
3. Remove `@libsql/client`, Turso environment variables, and SQLite-specific UI error handling.
4. Rename `db:health` to a PostgreSQL health/audit command that checks schema version, table counts, freshness, failed runs, and cursor validity.

Acceptance gate: repository search has no runtime libSQL/Turso imports, deployment has no Turso secrets, and the PostgreSQL audit command passes.

### Phase 5: Production verification

Automate these checks in CI and as a protected production smoke test:

1. Schema contract check against `information_schema`.
2. One read test per `/api/local` population, including PM property-group scoping.
3. Source assertion: PostgreSQL only.
4. Freshness thresholds per dataset.
5. Scheduler assertions: enabled endpoint set, recent successful run, no stuck `running` jobs, and monotonic cursors.
6. Data reconciliation against AppFolio performed only by an administrative audit worker.
7. AppFolio-egress-denied test proving the web application still reads all stored populations.

## Recommended implementation order

1. Production schema repair and inspection grid fix.
2. v2 scheduler enablement and controlled backfills.
3. Bills/tasks/report snapshot tables and workers.
4. Work Order note/attachment background ingestion.
5. Remove request-time live reads and frontend fallbacks.
6. Remove SQLite/Turso code, dependency, secrets, and cron jobs.
7. Expand health checks and add egress-denied acceptance testing.

Do not delete the legacy Turso target until external jobs have been inventoried and PostgreSQL reconciliation has passed. Its zero-row state means it is not useful as a production fallback, but deletion is irreversible and may affect an untracked cron.
