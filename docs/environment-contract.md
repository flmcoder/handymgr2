# Environment Contract

This repository uses `render.yaml` as a partial deployment snapshot, but the backend reads more environment keys than the export currently lists. Treat this document as the contract between code and Render dashboard configuration.

## Current Render Export

`render.yaml` currently mirrors only the database tunnel/runtime keys:

- `DB_SSL`
- `DB_USER`
- `DB_NAME`
- `DBI` / `PGHOST`
- `DB_PORT` / `PGPORT`
- `DB_PASSWORD`
- `SSH_DB_TUNNEL_ENABLED`
- `SSH_DB_TUNNEL_HOST`
- `SSH_DB_TUNNEL_PORT`
- `SSH_DB_TUNNEL_USER`
- `SSH_DB_TUNNEL_LOCAL_PORT`
- `SSH_DB_TUNNEL_REMOTE_HOST`
- `SSH_DB_TUNNEL_REMOTE_PORT`
- `SSH_DB_TUNNEL_PRIVATE_KEY`
- `SSH_DB_TUNNEL_PRIVATE_KEY_B64`
- `SSH_DB_TUNNEL_IDENTITY_FILE`
- `SSH_DB_TUNNEL_READY_TIMEOUT_MS`

## Backend Keys Read by Code

### Build/runtime

- `PORT` - Render runtime port
- `package.json` version - application release label used by the frontend and backend
- `RENDER_GIT_COMMIT` - deployed commit shown in backend startup logs
- `GIT_COMMIT` - local/runtime commit fallback shown in backend startup logs
- `RENDER_EXTERNAL_URL` - allowed CORS origin
- `APP_ORIGIN` - allowed CORS origin
- `FRONTEND_ORIGIN` - allowed CORS origin

### Database

- `SQL_SE` - optional direct connection string
- `DBI` / `PGHOST` - tunnel host or DB host
- `DB_PORT` / `PGPORT` - tunnel or DB port
- `DB_NAME` / `PGDATABASE` - database name
- `DB_USER` / `PGUSER` - database user
- `DB_PASSWORD` / `PGPASSWORD` - database password
- `DB_SSL` - SSL mode
- `DB_POOL_MAX` - pool size
- `DB_IDLE_TIMEOUT` - pool idle timeout
- `DB_CONNECT_TIMEOUT` - connect timeout

### AppFolio

- `AF_DEVELOPER_ID`
- `AF_CLIENT_ID`
- `AF_CLIENT_SECRET`
- `AF_REPORTS_CLIENT_ID`
- `AF_REPORTS_CLIENT_SECRET`
- `AF_V0_CLIENT_ID`
- `AF_V0_CLIENT_SECRET`
- `AF_V2_CLIENT_ID`
- `AF_V2_CLIENT_SECRET`
- `AF_SUBDOMAIN` / `AF_VHOST` / `APPFOLIO_SUBDOMAIN`
- `AF_DB_BASE`
- `AF_REPORTS_BASE`
- `AF_MAX_PER_SECOND`
- `AF_MAX_PER_MINUTE`
- `AF_MAX_PER_HOUR`
- `AF_V2_PRE_FETCH_DELAY_MS`

### Device auth and OTP

- `DEVICE_SETUP_PIN`
- `OTP_ALLOWED_DOMAIN`
- `DEVICE_OTP_TTL_MINUTES`
- `FRONTEND_PROXY_SECRET`
- `GUI_ADMIN`
- `GUI_GM`
- `GUI_VENDORS`

### RingCentral / SMS

- `RC_SERVER_URL` / `RINGCENTRAL_SERVER_URL`
- `RC_ACCESS_TOKEN` / `RINGCENTRAL_ACCESS_TOKEN`
- `RC_JWT` / `RINGCENTRAL_JWT`
- `RC_CLIENT_ID` / `RINGCENTRAL_CLIENT_ID`
- `RC_CLIENT_SECRET` / `RINGCENTRAL_CLIENT_SECRET`
- `RC_FROM_NUMBER` / `RINGCENTRAL_FROM_NUMBER`

### Webhooks and sync control

- `WEBHOOK_VERIFY_SIGNATURE`
- `WEBHOOK_JWKS_TTL_MS`
- `INTERNAL_SYNC_TOKEN`
- `SYNC_SCHEDULER_ENABLED`
- `SYNC_SCHEDULER_INTERVAL_MINUTES`
- `SYNC_SCHEDULER_ENDPOINTS`
- `SYNC_SCHEDULER_V2_ENABLED` - defaults to true; appends all supported normalized v2 populations even when an older endpoint list is configured
- `SYNC_SCHEDULER_MAX_PAGES`
- `SYNC_SCHEDULER_RUN_ON_BOOT`
- `WO_DETAIL_CACHE_TTL_MS`

## Drift Notes

- The dashboard must contain the AppFolio, RingCentral, OTP, sync, and proxy keys above even though `render.yaml` does not currently export them.
- Secret values should stay in Render dashboard secrets, not in git.
- `DBI` is intentionally used by the codebase as the tunnel host key, so the dashboard key should stay aligned with that name unless the code is changed.
- Any new backend env var should be added here first, then mirrored in Render.

## Live Render Audit (2026-07-12)

Service audited: `handymgr2` Render dashboard Environment page.

### Keys observed in dashboard and aligned with backend runtime

- Database/tunnel: `DBI`, `DB_PORT`, `DB_NAME`, `DB_USER`, `DB_PASSWORD`, `DB_SSL`, `SSH_DB_TUNNEL_*`
- AppFolio: `AF_DEVELOPER_ID`, `AF_SUBDOMAIN`, `AF_VHOST`, `AF_V2_CLIENT_ID`, `AF_V2_CLIENT_SECRET`, `AF_V2_PRE_FETCH_DELAY_MS`
- Legacy AppFolio aliases currently in use by sync auth fallback: `AF_V0_CLIENT_ID_store`, `AF_V0_CLIENT_SECRET_store`
- OTP/session/sync: `DEVICE_SETUP_PIN`, `DEVICE_OTP_TTL_MINUTES`, `OTP_ALLOWED_DOMAIN`, `FRONTEND_PROXY_SECRET`, `INTERNAL_SYNC_TOKEN`, `SYNC_SCHEDULER_ENDPOINTS`
- RingCentral: `RC_CLIENT_ID`, `RC_CLIENT_SECRET`, `RC_JWT`, `RC_FROM_NUMBER`
- Role compatibility: `GUI_ADMIN`, `GUI_GM`, `GUI_VENDORS`

### Keys not observed in dashboard snapshot (but still read by backend with fallbacks/defaults)

- Build/CORS values: `PORT`, `RENDER_GIT_COMMIT`, `GIT_COMMIT`, `RENDER_EXTERNAL_URL`, `APP_ORIGIN`, `FRONTEND_ORIGIN`
- DB optional aliases/tuning: `SQL_SE`, `PGHOST`, `PGPORT`, `PGDATABASE`, `PGUSER`, `PGPASSWORD`, `DB_POOL_MAX`, `DB_IDLE_TIMEOUT`, `DB_CONNECT_TIMEOUT`
- AppFolio optional/legacy alternatives: `AF_CLIENT_ID`, `AF_CLIENT_SECRET`, `AF_REPORTS_CLIENT_ID`, `AF_REPORTS_CLIENT_SECRET`, `AF_V0_CLIENT_ID`, `AF_V0_CLIENT_SECRET`, `AF_DB_BASE`, `AF_REPORTS_BASE`, `AF_MAX_PER_SECOND`, `AF_MAX_PER_MINUTE`, `AF_MAX_PER_HOUR`, `APPFOLIO_SUBDOMAIN`, `DEV`, `GUI_MANAGER_ID`, `GUI_MANAGER_PW`
- RingCentral alternate naming/token mode: `RC_SERVER_URL`, `RC_ACCESS_TOKEN`, `RINGCENTRAL_*`
- Webhook/scheduler optional controls: `WEBHOOK_VERIFY_SIGNATURE`, `WEBHOOK_JWKS_TTL_MS`, `SYNC_SCHEDULER_ENABLED`, `SYNC_SCHEDULER_INTERVAL_MINUTES`, `SYNC_SCHEDULER_MAX_PAGES`, `SYNC_SCHEDULER_RUN_ON_BOOT`, `WO_DETAIL_CACHE_TTL_MS`

### Keys observed in dashboard that are not read by `backend/server.ts`

- `ADMIN_EMAIL`, `ADMIN_SECRET`, `CRON_SECRET`, `HANDYMGR_INTERNAL_TOKEN`, `MAGIC_LINK_SECRET`, `MY_APP_SECRET`, `PROXY_ADMIN_KEY`, `PROXY_BASE_URL`, `APPFOLIO_JWKS_URL`, `WEBHOOK_MAX_DAYS`, `DEVICE_OTP_REQUEST_LIMIT`, `DEVICE_OTP_REQUEST_WINDOW_SEC`, `DEVICE_OTP_GLOBAL_LIMIT`, `DEVICE_OTP_GLOBAL_WINDOW_SEC`, `OTP_DIRECTORY_*`, `TURSO_DATABASE_URL`, `TURSO_AUTH_TOKEN`

These appear to be legacy, frontend-admin, or `afproxy` service keys rather than active `backend/server.ts` runtime requirements.

### Recommended drift actions

1. Keep current live keys required by `backend/server.ts` and sync fallback paths (`AF_V0_*_store`) until code cleanup removes those branches.
2. Move non-backend keys to the correct service (or remove if retired) to reduce secret sprawl in `handymgr2`.
3. If webhook verification remains enabled (default), no immediate action is required for `APPFOLIO_JWKS_URL`; backend currently uses a fixed AppFolio JWKS URL constant.
