# HandyManager current-state onboarding guide

This document is the working map for the current version of the app as it exists in this repo. It is intended to help a new developer understand what is configured, what is live, what is still in migration/half-configured mode, and what likely needs to be completed before the app feels seamless across all sections.

## 1) What the app is

HandyManager is a property operations dashboard built around AppFolio data and internal operational workflows. The current version is a browser-first management app for:

- work order triage and routing
- turnover / turn-board tracking
- occupancy and unit/property views
- inspections and tenant data
- vendor workflows
- manager review and dispatch/admin tooling
- OTP-based PM login and trusted-device authentication

The app is not a clean single-framework app in the usual modern sense. It is a Vite web app with a custom DOM-driven UI shell and several JS modules that assemble the app experience. It also uses an Express/Node backend and a Postgres-backed data layer with AppFolio sync logic.

## 2) High-level architecture

Current repo layout is essentially:

- `index.html` — app shell, login vault, navigation, sections, and global UI chrome
- `src/app.js` — main orchestration file; loads modules, auth state, refresh policy, routing, and section boot logic
- `src/*.js` — feature modules such as auth, tabs, nav, group filter, manager review
- `src/dashboard.ts` — ECharts dashboard and KPI data rendering
- `backend/server.ts` — Express API entry point and route registration
- `backend/schema.ts` — Drizzle/Postgres schema for cached AppFolio data
- `backend/sync/*` — AppFolio sync workers and repository / rate-limit logic
- `db/migrations/*` — SQL migrations for the live data model
- `afproxy/*` — older/legacy handler layer being bridged into Express routes
- `scripts/*` and `db.ts` — DB connection, tunnel setup, and operational scripts

This means the app is best thought of as a hybrid system:

- browser UI renders the app shell and reads data
- backend routes expose functions around AppFolio + DB access
- Postgres stores normalized operational data the app uses repeatedly
- AppFolio remains the upstream source of truth for units, properties, work orders, tenant info, and turn data

## 3) How the front end is assembled

### Front-end entry point

`index.html` defines the actual app shell and starts at the login vault and secure session flow.

Important behaviors visible there:

- login screen (`#vaultScreen`) with standard manager login and PM OTP login
- top bar with version / app status / refresh / lock controls
- left navigation with tabs for dashboard, work orders, routing, occupancy, properties, vendors, etc.
- a global property-group filter bar
- main content sections for each app area

### UI composition pattern

The application uses modular JS rather than a React/Vue router. The main bootstrap pattern is in `src/app.js` and it does things like:

- initialize theme behavior
- register maintenance banner / migration state
- resolve proxy URL and auth tokens from localStorage/sessionStorage
- create feature modules for tabs, nav, auth, manager review, group filtering
- manage restore/resume logic for sessions
- bind UI actions for login, refresh, filtering, and section switching

This makes onboarding important: the app is a stateful DOM app, not a clean route-based app. The behavior is spread across many DOM hooks and storage keys.

## 4) How the app loads and refreshes data

### Browser-side data strategy

The front end is designed to read from a configured proxy endpoint and they persist tokens and proxy URLs in localStorage. In `src/app.js` and `src/dashboard.ts`, the app resolves a proxy URL and attaches auth headers before calling fetch.

Key patterns:

- `hm_proxy_url` and `hm_proxy_token` are used as primary app configuration values
- `hm_auth_token`, `hm_device_token`, and `hm_scope_group_uuid` also drive session and scope state
- dashboards and data-heavy views fetch from the proxy with the active auth token and selected property group

The app’s current runtime model is basically:

- user authenticates
- browser stores token and scope
- browser calls backend/proxy API endpoints
- backend/proxy resolves data from AppFolio cache or direct APIs
- browser renders data into cards, tables, charts, panels, and kanban queues

### Data refresh and “live” behavior

A few things point to a live operational dashboard rather than a static app:

- the top bar includes a refresh button, cache badge, sync timestamp, rate badge
- there is a refresh interval and force-refresh gating in `src/app.js`
- the dashboard has an ECharts-driven KPI view and “Operations Pulse” panels
- panels such as stale data and pager center show operational status rather than only raw entity lists

Important: this is a data-oriented management app that expects near-live state updates, not only a simple CRUD app.

## 5) Backend and system connections

### Express API layer

`backend/server.ts` is the current backend entry point. It wraps older Deno-style handlers under Express routes and exposes endpoints like:

- `/health`
- `/api/units`
- `/api/unit_lookup`
- `/api/turns`
- `/api/unit_turns`
- `/api/turns_incremental`
- `/api/unit_turns_history`
- `/api/closed_turns`
- `/api/turn_records`
- `/api/turn_record_stage`
- `/api/reassignment_queue`
- `/api/device_otp_request`
- `/api/device_otp_verify`
- `/api/device_setup`
- `/api/verify_role`
- `/api/session_info`
- `/api/trusted_devices`
- `/api/pm_proxy_user`

The backend also validates bearer tokens, applies CORS restrictions, and handles data access wrappers for instrumentation and auth.

### Database layer

The app moved from a legacy SQLite-style design to PostgreSQL via Drizzle and `postgres-js` in `db.ts`.

Key connection model:

- prefer `SQL_SE` direct connection string if present
- otherwise build from `DBI` / `PGHOST`, `DB_PORT`, `DB_NAME`, `DB_USER`, and `DB_PASSWORD`
- uses a PgBouncer/tunnel-friendly setup with low pool size and short idle timeout
- `testConnection()` is used to confirm DB health on startup

This is a strong sign that the app is meant to run in a Render-managed container with a tunnel or direct Postgres route.

### AppFolio as the upstream system

AppFolio is the core external dependency. It drives the operational dataset. The app caches normalized record types in Postgres tables such as:

- `appfolio_properties`
- `appfolio_property_groups`
- `appfolio_units`
- `appfolio_work_orders`
- `appfolio_users`
- `appfolio_estimates`
- `appfolio_unit_inspections`
- `appfolio_tenant_directory`
- `appfolio_unit_turn_details`
- `appfolio_unit_vacancies`

This matters because the UIs are not querying AppFolio live on every action; the app expects locally normalized data in Postgres, refreshed via sync jobs.

### Sync systems

The sync worker lives under `backend/sync/*` and is the key operational piece that brings AppFolio data into the app’s database.

It includes:

- `syncRunner.ts` — orchestrates endpoint fetches and upserts
- `fetchWorker.ts` — AppFolio fetch logic plus retry/rate limit behavior
- `rateLimiter.ts` — request pacing and limiter controls
- `repositories.ts` — maps raw AppFolio objects into database tables
- `runStore.ts` — tracking state and cursor progress

The pattern is straightforward:

1. fetch a page or report from AppFolio
2. persist raw responses for debugging / recovery
3. map row payloads into domain tables
4. upsert them into Postgres
5. move the cursor forward to the next page / incremental window

This is a mature data pipeline pattern. It is one of the biggest hints that the app is operationally designed around a synchronized cache rather than a pure-on-demand API architecture.

## 6) Auth, OTP, and PM access model

The app includes a separate trusted-device and PM OTP system, which is one of the more specialized parts of the system.

Relevant files:

- `src/auth.js` — OTP normalization, phone/email detection, login flow support
- `backend/deviceAuth.ts` — device auth and token/session handling
- `server.ts` — `/api/device_otp_request`, `/api/device_otp_verify`, `/api/device_setup`, `/api/verify_role`, `/api/session_info`

Important behaviors:

- OTP can be sent to a PM email or phone number
- session/token state is stored in browser storage and validated on backend
- device trust / session validation is separate from standard manager login
- role qualification is used to gate tabs and sections

The PM auth path is an important part of the app and is clearly considered a foundational part of the system rather than an afterthought.

## 7) Configuration and deployment state

The repo includes a specific environmental ownership contract in `CONFIG.md` and `docs/environment-contract.md`.

### Important config point

This is not a fully static app. It has a real deployment contract and environment split:

- `render.yaml` contains the runtime export snapshot
- `docs/environment-contract.md` describes what backend code actually reads
- `CONFIG.md` states this is the single source-of-truth doc about environment ownership

The contract shows the code currently expects the following major live dependencies:

- Postgres connection and SSH tunnel settings
- AppFolio API credentials and host selection
- RingCentral JWT / SMS support for OTP messaging
- OTP/session secret keys
- sync and webhook controls
- internal sync token protection

This means the app is configured to run in a deployed env with environment secrets managed outside git, but the repo itself documents what must be configured.

### Configured but not fully set up yet

`CONFIG.md` and `docs/environment-contract.md` strongly imply a partially configured runtime state. Several things are documented as required but not fully represented in the exported Render snapshot. In practice, this means:

- some runtime values are known and expected but not fully mirrored in the current export
- some legacy keys are still recognized so the app can degrade gracefully
- some parts of the app appear operationally active but not yet fully cleaned up to modern naming or config structure

This is not a broken app; it is a system in migration / hybrid configuration state, and the code explicitly marks the app as under construction in the maintenance banner (`src/app.js`).

## 8) Data dependencies and dependency chain

The practical dependency chain is:

1. browser UI loads from Vite + static assets
2. browser obtains or restores auth token, proxy URL, and property-scope state
3. browser calls backend or proxy endpoints
4. backend reads DB connection and auth/session settings
5. backend or sync layer fetches from AppFolio
6. AppFolio data is normalized and inserted into Postgres tables
7. UI reads the normalized data for dashboards, work orders, occupancy, vendor, and property views

This is a classic operational data pipeline: external system -> normalization -> runtime cache -> UI.

The main data dependencies are therefore:

- AppFolio (source of property, unit, tenant, work order, and turn data)
- Postgres (local warehouse / operational cache)
- RingCentral (OTP SMS / messaging)
- Render or hosting runtime env variables (config, secrets, origin rules)
- browser storage/session state (scope, tokens, UI state)

## 9) Mobile / phone formatting concerns

This app is clearly desktop-first, with some responsive treatment but not a full mobile-first operational design.

Evidence in CSS and UI layout:

- lots of fixed width and min-width rules in `src/css/app.css`
- many tables / cards / kanban columns rely on min-width values and horizontal scrolling
- the nav uses a slide-out drawer for mobile, which suggests a partial mobile experience rather than a complete redesign
- some sections still assume wide-screen viewing and higher-density desk use
- repeated use of widths like `min-width: 420px`, `min-width: 220px`, `max-width: 360px`, `overflow-x: auto`, and layout grids that collapse only at certain breakpoints

This means:

- the app is usable on phones in a limited way
- most operational flows still feel designed around a laptop or large screen
- the main risk is that any dense operations workflow (work orders, board management, dashboard charts, data tables) will feel cramped when squeezed into narrow screens

The current repo even contains explicit UX design docs for workflows and states that users are primarily desktop-heavy, which is consistent with the implementation.

## 10) Biggest app gaps and “not seamless yet” areas

Based on the code and the docs, the biggest areas to improve are:

### A. App consistency and state cohesion

The app is modular but not always unified. The code has multiple state sources:

- localStorage keys for app state
- session data for auth
- proxy URL and API config in browser state
- database sync state for AppFolio data

Risk: different sections can drift if not all loaded with the same auth state or property-group scope.

### B. Data freshness and cache health

The app is heavily dependent on sync jobs and cached tables. If sync fails or a page is stale, the UI can show outdated values even when the front end still looks healthy.

This is especially relevant because the dashboard, stale-data panels, and operational metrics are intentionally built around a freshness model.

### C. Mobile responsiveness

The system works better on desktop than on phones. Dense tables, fixed-width cards, and board views likely need a dedicated mobile workflow or progressive simplification.

### D. Operational section cohesion

Some sections are clearly more mature than others. The app shell includes many tabs, but not all likely have the same reliability or polished UX. The repo docs also imply work-order triage and board workflow improvements are still important.

### E. Config and deployment cleanup

There is a clear indication that environment config is being cleaned up and normalized, but not every service / env key is fully consistent. This is the sort of issue that can create drift, deployment surprises, and hard-to-trace outages.

## 11) Biggest improvements recommended

If the goal is to get the app sections working smoothly together, the most valuable improvements are:

1. Standardize state and auth rules across the entire app
   - normalize proxy/auth/session handling
   - ensure property-group scope is consistently applied everywhere
   - eliminate duplicate token storage patterns

2. Make data freshness explicit in the UI
   - add clearer last sync timestamps per section
   - show stale-data warnings and broken sync state visibly
   - build a single “backend health + sync health” status panel

3. Reduce wide-screen assumptions
   - mobile-first simplification for work orders, dashboard, and turn board
   - collapse dense tables into stacked cards or alternate mobile layouts
   - trim fixed-width controls in mobile view

4. Tighten app section parity
   - align filters / sort logic / labeling between dashboard, work orders, turn board, occupancy, and properties
   - ensure each section reads from the same normalized data model

5. Complete config cleanup and deployment contract enforcement
   - keep env docs updated when keys change
   - treat `CONFIG.md` and `docs/environment-contract.md` as active operational requirements
   - remove dead legacy keys only after confirming they are no longer used

6. Make sync and health monitoring first-class
   - each section should know whether it is showing live, cached, or stale data
   - add summary health checks for AppFolio sync, DB health, and OTP service health

## 12) Recommended next steps

For an upcoming implementation phase, these are the steps most likely to produce the biggest win:

- Step 1: unify auth and token state so all sections share one trusted source of truth
- Step 2: standardize property-group scoping across all views and sync operations
- Step 3: add shared data freshness indicators and sync status at the API and section level
- Step 4: rework the mobile experience for the highest-frequency flows: dashboard, work orders, turn board, and PM login
- Step 5: tighten env/config tracking and remove legacy drift from deployment settings
- Step 6: restore consistent semantics across sections (filters, naming, status states, counts, actions)
- Step 7: build a final end-to-end smoke test pass across the major user journeys

## 13) Bottom line

The current app is a functional, operationally rich property-management dashboard that is clearly in a migration/hybrid configuration state. It mostly works because it is built around a clear architecture:

- browser UI shell
- backend API adapters
- Postgres normalized cache
- AppFolio upstream sync
- specialized auth and PM login flows

The biggest remaining work is not writing a new app from scratch. It is tightening the edges:

- system consistency
- data freshness transparency
- mobile usability
- deployment/config cleanup
- section-to-section operational parity

That is the current state of the app as represented in this repo.
