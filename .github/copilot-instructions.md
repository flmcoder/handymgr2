# HandyManager — Workspace Instructions

## Project Identity

- **App**: HandyManager — internal maintenance cockpit for Fort Lowell Realty & Property Management
- **Architecture**: Monolithic single-file frontend (`index.html`) + Val Town serverless proxy (`afproxy/`) + Turso (libSQL/SQLite) cache
- **External APIs**: AppFolio Database API v0 (CRUD, webhooks) · AppFolio Reports API v2 (bulk pulls)

## Frontend Stack — ZERO TOLERANCE

**Stack is HTML5 + CSS3 + Vanilla JavaScript (ES6+). Nothing else.**

NEVER add or reference: React, Vue, Angular, Svelte, jQuery, Lodash, Axios, Tailwind, Bootstrap, any npm package, any build tool (Webpack, Vite, Rollup), or TypeScript on the frontend.

- Use the app's `$()` and `$$()` aliases for DOM queries — never `document.getElementById()` or `document.getElementsByClassName()`
- All colors must use CSS custom variables (`--bg-primary`, `--text-primary`, `--accent`, `--color-warn`, etc.) — never hardcode hex/rgb
- Every UI component must work in both light and dark mode
- Use event delegation on stable parents — never attach per-element listeners in render loops
- Global state lives in window-scoped variables (`WORK_ORDERS`, `BILLS`, `VENDORS`, `TURNS`, `PROPERTIES`, `currentPropertyGroup`, `forcedPropertyGroupUuid`, etc.)
- Never use `innerHTML` with unsanitized user input — use `textContent` for user-supplied values
- Never store credentials, tokens, or subdomains in frontend JS, localStorage, or sessionStorage

## Backend Stack (afproxy/)

- **Runtime**: Deno (NOT Node.js) · **Language**: TypeScript strict
- All credentials in Val Town environment variables — never hardcoded
- All proxy requests use `?action=action_name` pattern — no per-resource REST routes
- CORS handled at proxy layer only — never in frontend fetch calls
- Database: Turso via `@libsql/client` — raw parameterized SQL only, no ORMs

## AppFolio API — Critical Rules

- **Rate limits**: 8 req/s · 256 req/min · 4096 req/hr — all bulk fetches are server-side
- **On 429**: read `Retry-After` header, wait that duration, retry ONCE — never retry immediately
- **Field casing**: Database API v0 = PascalCase (`Id`, `VendorId`), Reports API v2 = snake_case (`vendor_id`, `property_id`) — always coalesce both when data source varies
- **Dates**: strictly ISO 8601 (`YYYY-MM-DDTHH:mm:ssZ`) — never bare `YYYY-MM-DD` to datetime fields, never `MM/DD/YYYY`
- **History endpoints** (`bills_history`, `work_orders_completed_history`): always require `from_date` — never use `paginate_results=false` without date bounds
- **PATCH concurrency**: AppFolio fails the second concurrent PATCH to the same resource — use in-flight guards
- **UUIDs**: validate with `/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i` before use

## Database Rules (Turso/SQLite)

- All queries use parameterized placeholders (`?` or `:name`) — never concatenate values
- AppFolio IDs as TEXT (UUID strings) — never INTEGER PKs for AF entities
- Every cache table has `cached_at` (TEXT ISO 8601)
- Schema changes are additive only — never `DROP TABLE` or `DROP COLUMN`
- Store dates as ISO 8601 TEXT or Unix epoch INTEGER

## Timeout Map — Do Not Reduce

```js
var timeoutByAction = {
  turn_work_orders:              120000,
  bills_history:                 120000,
  properties:                     90000,
  bills:                          90000,
  turns:                          90000,
  unit_turns:                     90000,
  work_orders_completed_history:  90000,
};
var baseTimeoutMs = timeoutByAction[action] || 60000;
```

Do not reduce any value. History endpoints: 120s. Paginated endpoints: 90s. Default: 60s.

## Hard Stops — Reject Immediately

- Adding any framework/library to the frontend
- Calling AppFolio API directly from frontend JS
- Using `innerHTML` with unsanitized input
- Storing credentials in client-side storage
- Using `paginate_results=false` on unbounded history queries
- Firing simultaneous PATCHes to the same resource
- Ignoring `Retry-After` on 429 or retrying without backoff
- Hardcoding property group names, UUIDs, or subdomains
- Using `DROP TABLE`/`DROP COLUMN` in migrations
- Using `alert()`, `confirm()`, or `prompt()`
- Leaving any table body empty without an empty-state row
- Renaming section IDs without grepping all JS references
