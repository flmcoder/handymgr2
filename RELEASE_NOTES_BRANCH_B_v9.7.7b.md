# Branch B Release Notes — v9.7.7b

## Summary
- Improved login/session UX so **Resume Session** is the primary action when a resumable session exists.
- Added responsive/layout refinements to reduce UI clashes across topbar, dashboard chart hosts, and dispatch surfaces.
- Improved dashboard insight behavior with stronger chart lifecycle handling, first-load empty states, and source context messaging.
- Added a new **Estimate & Spend Flow** insight with range/vendor filters and click-through to scoped work-order view.
- Improved PM message experience by treating operational alerts as unread/seen message items instead of repeating persistent noise.
- Added a dispatch **24/7 Simulation Readiness** summary panel for mock-trial preparation.
- Improved v0 sync cron behavior by using configurable lookback-derived `from_date` rather than a fixed date.

## Visual QA Matrix (manual verification targets)
| Theme | Width | Target |
|---|---:|---|
| Light | 390px | Vault/login CTAs, topbar wrapping, drawer/fab overlap, dashboard insight card stacking |
| Dark | 390px | Same as above + icon/button contrast |
| Light | 768px | Dashboard chart panel stacking, bento insight host sizing, dispatch controls wrap |
| Dark | 768px | Same as above + badges/alerts legibility |
| Light | 1366px | Topbar right cluster spacing, chart host sizing consistency, drawer positioning |
| Dark | 1366px | Same as above |
| Light | 1920px | No chart compression/overlap, dispatch table and panel spacing |
| Dark | 1920px | Same as above |

## SQLite → PostgreSQL Compatibility Checklist (prep)
- Keep UI contracts DB-agnostic (no SQLite-specific assumptions in dashboard/message rendering).
- Maintain UUID-first linkage for WO notes/attachments and modal enrichment paths.
- Preserve upsert semantics used by proxy sync/config operations.
- Validate date/time parsing and numeric coercion consistency for chart aggregations.
- Ensure indexed lookup strategy exists for WO UUID resolution and vendor-spend aggregations.

## Known Non-Security Technical Debt / Follow-Up
- Typecheck/build environment currently requires dependency repair in this sandbox before green validation.
- Some dashboard charts still rely on mixed ownership across `app.js` and `dashboard.ts`; further modular consolidation is recommended.
- Dispatch simulation currently surfaces readiness diagnostics and blockers, but not an end-to-end dry-run replay engine.
- Spend-flow analytics relies on available WO fields; proxy-side enrichment for approval-state fidelity should be expanded.
