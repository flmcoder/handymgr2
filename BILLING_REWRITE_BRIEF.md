# Billing Rewrite Brief (Fresh LLM Handoff)

## Purpose

This document is a clean handoff for rewriting the Billing/AP section with a fresh perspective.

Current status:
- PM group filtering is now working and must be preserved.
- UX and data flow are still brittle.
- PM users are effectively under-served because only a partial bill set is being loaded in common flows.

Primary symptom to solve:
- A PM only sees a subset of bills/work orders because the filter flow requests only `limit: 50` from route endpoints.

---

## Stack and Constraints (Must Keep)

Frontend:
- HTML + CSS + vanilla JavaScript only.
- Main app logic is in `js/app.js`.
- UI shell is in `index.html`.

Backend:
- Deno TypeScript proxy in `afproxy/`.
- Billing handler: `afproxy/handlers/bills.ts`.
- PM/group scope helper: `afproxy/lib/groupUtils.ts`.

Guardrails:
- No frontend frameworks.
- No direct AppFolio calls from frontend.
- Preserve PM scope and group UUID enforcement.
- Additive schema changes only.

---

## Current Billing Architecture (As-Is)

Frontend paths (`js/app.js`):
- `fetchBills(days, opts)` pulls billing data via proxy actions.
- `wireBillingFilters()` triggers route-style filter calls.
- `renderBillsSection()` renders list + local paging.
- `showBillDetailModal(billId)` opens bill detail modal with line items + payload.
- `runBillHistorySearch()` fetches history rows and paginates client-side.

Backend paths (`afproxy/handlers/bills.ts`):
- `handleBillsList()` standard list route with page/per-page semantics.
- `handleBillsRoute()` route-style actions (`bills_list`, `bills_by_vendor`, etc.) with `limit`/`offset` semantics.
- `handleBillDetail()` detail fetch.
- `handleBillsHistory()` history fetch (capped slice).

---

## Known Logic Problems

### 1) Route filter path hard-caps results at 50 in frontend

In `wireBillingFilters()` the apply action sends:
- `limit: 50`
- `offset: 0`

That means the backend can return at most 50 rows for that request, even when the PM scope has hundreds of bills.

Impact:
- PM sees partial bill list.
- Work-order-to-bill coverage appears low.
- User trust degrades because totals and visible rows do not align with reality.

### 2) Mixed pagination models

The current system mixes two models:
- Server-side route calls using `limit/offset`.
- Client-side pagination using `BILLS_PAGE_SIZE` after data arrives.

When the initial fetch is truncated to 50, local pagination is irrelevant because it paginates an already incomplete dataset.

### 3) Multiple fallback paths with inconsistent semantics

`fetchBills()` can call:
- Route actions (`bills_list`, `bills_by_*`)
- Legacy `bills` action fallback

These paths do not always align in pagination and filtering contract, making bugs hard to reason about.

### 4) History path may still cap large ranges

History has improved, but there are still hard slices/caps in backend or frontend flow that can silently truncate large time windows.

---

## Rewrite Goals

### Product goals
- PM sees complete in-scope bills for selected filters.
- Billing list, KPIs, and bill-history numbers are internally consistent.
- Bill detail modal opens from every row path and shows full detail payload.

### Technical goals
- Single query model for billing list.
- Explicit server-driven pagination contract.
- No hidden row caps.
- Deterministic filter state and request state.

### Non-goals
- Do not redesign auth model.
- Do not change PM scope rules.
- Do not introduce frameworks.

---

## Recommended Fresh Design

## A) Standardize one list API contract

Use one primary list contract for the billing table:
- `action=bills` (or dedicated `action=bills_list_v2`)
- Query includes:
  - `page`
  - `per_page`
  - `group_uuid` (from PM/session scope)
  - `status`, `vendor_id`, `property_id`, `wo_id`, `invoice_number`, `due_from`, `due_to`

Response should always include:
- `results`
- `total`
- `page`
- `per_page`
- `total_pages`
- `from_cache`
- `cached_at`

Important:
- Remove fixed frontend `limit: 50` behavior from default table flow.
- Keep page size configurable (example defaults: 50 or 100), but always paged over full dataset.

## B) Billing state machine on frontend

Create a single Billing state object in `js/app.js`:
- `query`: all active filters + paging
- `rows`: current page results
- `total`, `totalPages`, `loading`, `error`
- `source`: live/cached/legacy

All UI actions mutate `query` then call one loader:
- search input
- status filter
- route/filter selectors
- next/prev page

This removes split logic between route-only and fallback-only code paths.

## C) Separate list queries from KPI queries

KPI cards should infer some data totals from currently rendered page rows.

Use a dedicated stats endpoint call for KPI data (already available in concept):
- pending approval count/amount
- approved not paid
- paid this period
- vendor count

## D) Keep bill detail independent

`showBillDetailModal(billId)` should always fetch detail by bill id and not depend on whether full row payload is in current page cache.

## E) History should be explicitly scoped and paged

For history:
- Keep date range required.
- Return paged results and total count for large windows.
- Avoid silent slices when possible.
- If capped intentionally, show explicit warning in UI.

---

## Minimal Fix (If Full Rewrite Is Deferred)

If you need immediate relief before full redesign:
1. Remove hardcoded `limit: 50` from billing filter apply flow.
2. Add server-side paging controls in Billing table footer.
3. Request first page with a larger page size (e.g., 100) and fetch subsequent pages.
4. Ensure footer reflects server `total` not local array length only.

This does not solve architecture debt, but removes the most visible PM pain quickly.

---

## Acceptance Criteria

A rewrite is done only when all are true:

1. PM scope correctness
- PM can only see rows in their scoped property group.
- No out-of-scope leakage.

2. Completeness
- PM with >50 in-scope bills can page through full set.
- UI clearly shows page and total.

3. Consistency
- KPI values match scoped server stats, not just current page rows.

4. Detail reliability
- Clicking any bill row opens populated detail modal.
- Detail includes line items, associations, payload tools, and attachments section.

5. Performance
- Initial load under normal cached conditions is responsive.
- No runaway multi-fetch loops.

6. Error clarity
- 4xx/5xx errors surface user-readable messages.
- Logs include action + key params (without secrets).

---

## Suggested Test Matrix

Core:
- Group-scoped PM with 10 bills.
- Group-scoped PM with 500+ bills.
- Full admin with all groups.

Filters:
- By vendor
- By property
- By WO UUID and WO number
- By invoice
- Due date range
- Status combinations

Pagination:
- page 1, middle page, last page
- filter change resets to page 1
- total/pages remain correct after filter changes

History:
- small range (7 days)
- large range (180/365 days)
- zero-result range

Failure modes:
- proxy 401
- proxy 429
- proxy 500
- malformed payload from backend

---

## Implementation Notes for New LLM

Where to start:
1. Read billing UI flow in `js/app.js` around:
- `fetchBills`
- `wireBillingFilters`
- `renderBillsSection`
- `showBillDetailModal`
- `runBillHistorySearch`

2. Read backend billing handler in `afproxy/handlers/bills.ts`:
- list/stats/detail/history/route actions
- understand route vs list pagination differences

3. Preserve PM scope behavior from `group_uuid`/`property_group_id` injection and group resolution helpers.

4. Implement one list-query abstraction first, then migrate UI handlers to use it.

---

## Definition of Success

A PM should be able to open Billing and reliably answer:
- "How many bills are in my portfolio right now?"
- "Am I seeing all of them, not just the first 50?"
- "Can I open any row and trust the detail data?"

If those answers are not clearly yes, the rewrite is not complete.
