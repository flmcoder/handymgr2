# CONFIG

This file is the single source of truth for HandyManager environment configuration ownership.

## Ownership model

- Repository-managed:
  - `.env.example` (shape and defaults only, no secrets)
  - `render.yaml` service spec and exported key list
  - `docs/environment-contract.md` and this file
- Render Dashboard-managed:
  - All production secrets
  - Runtime values for integrations (AppFolio, RingCentral, OTP, sync controls)

## Environment groups

### 1) Node/Express backend (required in production)

| Key | Purpose | Secret | Owner |
| --- | --- | --- | --- |
| `PORT` | Render-assigned HTTP listen port | No | Render Dashboard |
| `APP_VERSION` | Preferred build/version label | No | Render Dashboard |
| `RENDER_GIT_COMMIT`, `GIT_COMMIT` | Version fallbacks | No | Render Runtime |
| `RENDER_EXTERNAL_URL`, `APP_ORIGIN`, `FRONTEND_ORIGIN` | CORS/allowed origins | No | Render Dashboard |

### 2) Vite frontend

Required frontend env vars today: none.

- The frontend does not currently consume `import.meta.env.VITE_*` values.
- API targeting is runtime-based in `src/app.js` (`window.location` + `RENDER_API_BASE_URL` constant).
- If a `VITE_*` key is introduced, update `.env.example`, this file, and `docs/environment-contract.md` in the same change.

### 3) Postgres + SSH tunnel (required)

| Key | Purpose | Secret | Owner |
| --- | --- | --- | --- |
| `SQL_SE` | Direct Postgres URL mode | Yes | Render Dashboard |
| `DBI`, `DB_PORT`, `DB_NAME`, `DB_USER`, `DB_PASSWORD`, `DB_SSL` | Split connection mode | `DB_PASSWORD` only | Render Dashboard |
| `PGHOST`, `PGPORT`, `PGDATABASE`, `PGUSER`, `PGPASSWORD` | Optional aliases | `PGPASSWORD` only | Render Dashboard |
| `DB_POOL_MAX`, `DB_IDLE_TIMEOUT`, `DB_CONNECT_TIMEOUT` | Pool tuning | No | Render Dashboard |
| `SSH_DB_TUNNEL_ENABLED` | Enable SSH tunnel | No | Render Dashboard |
| `SSH_DB_TUNNEL_HOST`, `SSH_DB_TUNNEL_PORT`, `SSH_DB_TUNNEL_USER` | Bastion target | No | Render Dashboard |
| `SSH_DB_TUNNEL_LOCAL_PORT`, `SSH_DB_TUNNEL_REMOTE_HOST`, `SSH_DB_TUNNEL_REMOTE_PORT` | Port forwarding | No | Render Dashboard |
| `SSH_DB_TUNNEL_IDENTITY_FILE` | Key file path | Potentially sensitive path | Render Dashboard |
| `SSH_DB_TUNNEL_PRIVATE_KEY`, `SSH_DB_TUNNEL_PRIVATE_KEY_B64` | Inline private key material | Yes | Render Dashboard |
| `SSH_DB_TUNNEL_READY_TIMEOUT_MS` | Tunnel startup timeout | No | Render Dashboard |

### 4) AppFolio API integrations (required)

| Key | Purpose | Secret | Owner |
| --- | --- | --- | --- |
| `AF_DEVELOPER_ID` | v0 developer header | No | Render Dashboard |
| `AF_CLIENT_ID`, `AF_CLIENT_SECRET` | Primary API credentials | Yes | Render Dashboard |
| `AF_REPORTS_CLIENT_ID`, `AF_REPORTS_CLIENT_SECRET` | Dedicated reports credentials | Yes | Render Dashboard |
| `AF_SUBDOMAIN` / `AF_VHOST` / `APPFOLIO_SUBDOMAIN` | Customer subdomain selection | No | Render Dashboard |
| `AF_DB_BASE`, `AF_REPORTS_BASE` | Base URL overrides | No | Render Dashboard |
| `AF_MAX_PER_SECOND`, `AF_MAX_PER_MINUTE`, `AF_MAX_PER_HOUR` | Rate limiter controls | No | Render Dashboard |
| `AF_V2_PRE_FETCH_DELAY_MS` | Optional fetch pacing | No | Render Dashboard |

Legacy aliases still recognized by code (`AF_V0_*`, `AF_V2_*`, `DEV`, `GUI_MANAGER_ID`, `GUI_MANAGER_PW`) should remain in Render until fully removed from backend logic.

### 5) RingCentral JWT / SMS (required if OTP SMS enabled)

| Key | Purpose | Secret | Owner |
| --- | --- | --- | --- |
| `RC_SERVER_URL` | RingCentral platform URL | No | Render Dashboard |
| `RC_ACCESS_TOKEN` | Direct bearer token mode | Yes | Render Dashboard |
| `RC_JWT` | JWT mode credential | Yes | Render Dashboard |
| `RC_CLIENT_ID`, `RC_CLIENT_SECRET` | JWT client credentials | Yes | Render Dashboard |
| `RC_FROM_NUMBER` | Sender number | No | Render Dashboard |

Alternate names (`RINGCENTRAL_*`) are also read by backend and should be kept consistent if used.

### 6) OTP, session signing, and roles

| Key | Purpose | Secret | Owner |
| --- | --- | --- | --- |
| `DEVICE_SETUP_PIN` | Device bootstrap PIN | Yes | Render Dashboard |
| `OTP_ALLOWED_DOMAIN` | PM login domain restriction | No | Render Dashboard |
| `DEVICE_OTP_TTL_MINUTES` | OTP expiration | No | Render Dashboard |
| `FRONTEND_PROXY_SECRET` | Session token signing key | Yes | Render Dashboard |
| `GUI_ADMIN`, `GUI_GM`, `GUI_VENDORS` | Legacy role checks/compat | Yes | Render Dashboard |

### 7) Webhook and sync controls

| Key | Purpose | Secret | Owner |
| --- | --- | --- | --- |
| `WEBHOOK_VERIFY_SIGNATURE` | Enable AppFolio signature verification | No | Render Dashboard |
| `WEBHOOK_JWKS_TTL_MS` | JWKS cache TTL | No | Render Dashboard |
| `INTERNAL_SYNC_TOKEN` | Protect internal sync routes | Yes | Render Dashboard |
| `SYNC_SCHEDULER_ENABLED` | Toggle scheduler | No | Render Dashboard |
| `SYNC_SCHEDULER_INTERVAL_MINUTES` | Run interval | No | Render Dashboard |
| `SYNC_SCHEDULER_ENDPOINTS` | Endpoint allowlist | No | Render Dashboard |
| `SYNC_SCHEDULER_MAX_PAGES` | Page cap per run | No | Render Dashboard |
| `SYNC_SCHEDULER_RUN_ON_BOOT` | Startup run toggle | No | Render Dashboard |
| `WO_DETAIL_CACHE_TTL_MS` | Work order detail cache TTL | No | Render Dashboard |

## Drift prevention checklist

1. Add or modify env keys in code.
2. Update `.env.example` in the same PR.
3. Update `docs/environment-contract.md` and `CONFIG.md` in the same PR.
4. Update Render Dashboard values.
5. If the key should be exported, mirror it in `render.yaml`.
6. Validate via `/api/health/providers` and `/api/local/sync/status` after deploy.

## Live reconciliation status (2026-07-12)

Render dashboard key audit for service `handymgr2` is now reflected in `docs/environment-contract.md`.

- Confirmed active backend keys are present for DB/tunnel, AppFolio v2, OTP/session, sync token, and RingCentral JWT mode.
- Confirmed legacy AppFolio fallback keys (`AF_V0_CLIENT_ID_store`, `AF_V0_CLIENT_SECRET_store`) are still in production and still read by code.
- Found dashboard keys not used by `backend/server.ts` (mostly legacy/`afproxy`/admin keys). Keep only if another active service depends on them.
- `render.yaml` remains a partial snapshot and should not be treated as complete env inventory.

## Security policy

- Never commit real secrets to git.
- Keep production credentials only in Render Dashboard secret values.
- `.env.example` must contain placeholders only.
