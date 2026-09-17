# HandyManager v9.8.36 — User Guide

**HandyManager** is a property operations management dashboard built on top of AppFolio. It provides real-time visibility into work orders, turns, inspections, billing, vendor management, and automated dispatch — all scoped to your property portfolio.

---

## Table of Contents

1. [Getting Started](#getting-started)
2. [Authentication & Login](#authentication--login)
3. [Navigation & Layout](#navigation--layout)
4. [Dashboard](#dashboard)
5. [Work Orders](#work-orders)
6. [PM Routing Monitor](#pm-routing-monitor)
7. [Dispatch Control](#dispatch-control)
8. [Turn Board](#turn-board)
9. [Inspections](#inspections)
10. [Billing](#billing)
11. [Manager Review](#manager-review)
12. [Occupancy](#occupancy)
13. [Properties](#properties)
14. [Vendors](#vendors)
15. [Communication Templates](#communication-templates)
16. [Database Admin](#database-admin)
17. [Error Log](#error-log)
18. [Global Property Group Filter](#global-property-group-filter)
19. [Settings & PWA](#settings--pwa)

---

## Getting Started

### What is HandyManager?

HandyManager is a browser-based Progressive Web App (PWA) that syncs data from AppFolio and presents it in an operations-focused dashboard. It is designed for property managers, maintenance teams, and administrators at Fort Lowell Realty & Property Management (FLRAZ).

### Requirements

- A modern browser (Chrome, Edge, Safari, Firefox)
- Internet connection for initial data sync
- Once loaded, core features work offline via PWA caching

### Installing as PWA

1. Open HandyManager in your browser
2. Click the install icon in the address bar (or use the browser menu → "Install")
3. The app will appear as a standalone desktop/mobile application
4. Offline support is automatic once installed

---

## Authentication & Login

### Manager Login

1. Navigate to the HandyManager login screen (the "Vault")
2. Enter your **email** and **password**
3. Click **Login**

### PM Login via OTP (One-Time Password)

1. Select **PM Login** on the vault screen
2. Enter your **email address** or **phone number**
3. A 6-digit OTP code will be sent to you via SMS (RingCentral)
4. Enter the OTP code within 10 minutes
5. You will be authenticated and assigned a role based on your PM profile

### Device Setup (First Time)

1. On first login from a new device, you may be prompted for a **Device Setup PIN**
2. Enter the PIN provided by your administrator
3. The device will be registered as a trusted device
4. Future logins from this device will not require re-verification

### Resuming a Session

If you previously had an active session, you can click **Resume Previous Session** on the vault screen to reconnect without re-authenticating.

### Role-Based Access

| Role | Access |
|------|--------|
| **vendors** | Vendors tab only |
| **pm_readonly** | Dashboard, Work Orders, Billing, Occupancy, Properties, Turn Board, Vendors, Inspections, Error Log |
| **manager** | All of the above + Routing, Manager Review, Database Admin |
| **admin** | Full access including Dispatch |

### Logging Out

Click the **Lock** button in the top bar to wipe your credentials and return to the vault screen.

---

## Navigation & Layout

### Sidebar Navigation

The left sidebar contains all navigation tabs organized into four groups:

**Overview**
- Dashboard
- Routing (GM/Admin only)
- Dispatch (Admin only)

**Operations**
- Work Orders
- Turn Board
- Inspections
- Billing
- Manager Review (GM only)

**Resources**
- Occupancy
- Properties
- Vendors
- Templates

**System**
- Database (GM/Admin only)
- Error Log

### Top Bar

- **Property Group Filter** — Global filter that scopes all data across the app
- **Rate Badge** — Shows current API request rate
- **Cache Badge** — Indicates data source (live vs cached)
- **Sync Timestamp** — When data was last synced from AppFolio
- **Theme Toggle** — Switch between light and dark mode
- **Notifications** — Floating action button for messages and alerts
- **Webhook Feed** — Slide-out drawer showing real-time AppFolio events

### Badge Counts

Navigation tabs display colored badge counts showing the number of items requiring attention in each section. Counts are scoped to your current Property Group filter.

---

## Dashboard

The Dashboard is your command center — a single-page overview of portfolio health.

### KPI Cards (Top Rail)

| Card | Description |
|------|-------------|
| **Open Work Orders** | Total count of active work orders with sub-metrics |
| **Urgent / Critical** | Red-highlighted count of urgent work orders |
| **Active Turns** | Blue-highlighted count of turns currently in progress |
| **Upcoming Move-Outs** | Green-highlighted, shows move-outs within 60 days |
| **Flagged for Follow-Up** | Purple, count of work orders flagged for follow-up |
| **Pending Bill Approvals** | Amber, billing items awaiting approval |

> **Tip:** Click any KPI card to drill down into the relevant data.

### Manager Overview

Secondary metrics below the KPI rail:
- Average Turnover Completion Time
- Average Days Since Inspection
- Average Work Order Open Days
- Recent Completed Turns
- Vacancies by Group

### Turn Progress

- Toggle between **Card** and **Compact** views
- **TV Mode** for lobby/display screens (auto-refreshing)
- Filter by PM and paginate through active turns
- Sync button to refresh data with timestamp

### Portfolio Insights (Charts)

| Chart | Type | Shows |
|-------|------|-------|
| PM Workload Comparison | Bar chart | Work order distribution per PM |
| Open WO Type Mix | Pie/donut chart | Breakdown by work order type |
| Urgent Pressure by PM | Bar chart | Urgent WOs per PM |

### Advanced Operations & Portfolio Flow (Charts)

| Chart | Type | Shows |
|-------|------|-------|
| Turnover Pipeline | Funnel | Turn stage progression |
| Portfolio Sunburst | Sunburst | Hierarchical property/group/unit view |
| Work Order Flow | Sankey diagram | WO lifecycle movement between states |
| Inspection Map | Effect scatter | Geocoded inspection locations |

### Attention Required Board

Five summary cards flagging items needing action:
1. **Stalled Turns** — Turns stuck at a stage
2. **Overdue Inspections** — Past-due inspections
3. **Vendor Alerts** — Vendor compliance or availability issues
4. **Urgent / Overdue WOs** — High-priority or aging work orders
5. **Pending Bill Approvals** — Bills waiting for review

### Two-Column Support

- **Upcoming Move-Outs Table** — Property, Unit, Tenant, Date, Days Left (60-day window)
- **Recent Activity Feed** — Filterable by: All, Work Orders, Turns, Inspections, Urgent Only

---

## Work Orders

The Work Orders section is the core operational hub for managing maintenance tasks.

### Subtabs

1. **Open (Active)** — All current work orders
2. **Completed / Inactive** — Historical closed work orders
3. **WO Closure Assistant** — Helps close old work orders
4. **WO Follow-Up Queue** — Flagged items needing follow-up

### Active Work Orders View

#### View Modes

- **List (Queue)** — Standard table view with sortable columns
- **Kanban Board** — Visual board grouping WOs by status/stage

#### KPI Strip

Dynamic count badges at the top showing totals by priority, type, and status.

#### Filters

| Filter | Options |
|--------|---------|
| Text search | Search across WO descriptions, addresses, etc. |
| Priority | Urgent, Normal, Low |
| Type | Internal, Tenant Requested, Unit Turn |
| Vendor | Dynamic dropdown of all vendors |
| Property | Dynamic dropdown of all properties |
| Age | 7, 14, 30, 60, 90+ days |
| Sort | Ops Queue, Oldest, Newest, Priority |
| Quick Filters | Flagged, Estimate Requested, Estimated |
| Aging Threshold | Customizable Yellow/Orange/Red day thresholds |

#### Creating a New Work Order

1. Click **New Work Order**
2. Select the **Property**
3. Enter the **Unit** number
4. Choose **Priority** (Normal / Urgent / Low)
5. Enter a **Description** of the issue
6. Search for and assign a **Vendor** (optional)
7. Click **Save**

#### Work Order Detail

Click any work order row to open the detail modal showing:
- Full work order details (description, priority, vendor, dates)
- **Flag** toggle for follow-up
- Associated turn information (if applicable)
- **Save Changes** to update

### WO Closure Assistant

Helps identify and close old work orders:
1. Select an age threshold (14, 30, 60, 180+ days)
2. The system shows matching WOs with **AP (Accounts Payable) evidence** matching
3. A **confidence score** indicates how likely the WO can be safely closed
4. Navigate through results with pagination

### WO Follow-Up Queue

Shows all flagged work orders with:
- Current status and notes
- Actions available per WO
- Ability to unflag or take follow-up action

### Completed Work Orders

- Configurable history lookback period
- Pagination for older completed work
- Same filtering and search capabilities

---

## PM Routing Monitor

> **Access:** GM and Admin roles only

Flags work orders that may be incorrectly assigned to the wrong PM or should be handled by in-house maintenance.

### KPI Strip

Flagged / Pending / High Confidence / Reviewed / Last Scan

### PM Scorecard

A leaderboard table showing per-PM routing metrics:
- Flagged count
- High confidence count
- Pending count
- Approved count
- Reassigned count
- Potential FLM (Facilities Maintenance) Bypass count

### Flagged Routing Events

Table columns: WO, Property, Group, PM, Vendor, Trade Match, WO Type, Confidence, Status

#### Filters

| Filter | Options |
|--------|---------|
| Status | Pending, All, Approved Vendor, Reassign In-House, Dismissed |
| PM | Filter by specific PM |
| Days Range | 30, 60, 90, 180 days |
| Work Type | Filter by maintenance trade |
| WO Type | Filter by work order type |
| Confidence | Low, Medium, High threshold |
| Text Search | Search across all fields |

### PM Map

Configure the Portfolio Group → PM mapping with auto-fill capabilities.

### Work Type Toggles

Configure which work types are handled in-house:
- Add new work types
- Edit existing types
- Enable/disable work types

### Manual Scan

Click **Scan Loaded WOs** to trigger a manual routing analysis scan.

---

## Dispatch Control

> **Access:** Admin role only

Automated work order reassignment system with technician management.

### Top Bar Controls

| Control | Description |
|---------|-------------|
| Branch Selector | Filter by All / Phoenix / Tucson |
| Refresh | Reload dispatch data |
| Pause Automation | Temporarily stop auto-reassignment |
| Run Warning Pass | Execute warning notification cycle |
| Run Reassign Pass | Execute reassignment cycle |

### Stats Bar

In Queue / Warned / Escalated / Exempt / Tier 1 Active / Tier 2 Pool / Open Blasts / Tenant SMS

### Subpanels

#### 1. Grades (Performance Leaderboard)

Technician performance ranking:
- Tech name, Score, WO Share, Active WOs
- Go-Back % (revisit rate), Reassign % (reassignment rate)
- Zone, Status

#### 2. Queue (Reassignment Queue)

Work orders in the reassignment pipeline:
- Status, WO#, Address, Tech, Reassigns
- First Seen, Last Action, Actions
- Search and status filter

#### 3. Roster (Tech Roster)

Manage your technician roster:
- **Sync From AppFolio** — Pull latest tech list from AppFolio
- **Add Tech** — Manually add new technicians
- Columns: Tier, Name, Phone (E.164), Branch, Score, Active WOs, WO Share, Status
- **Tech Add/Edit Modal:**
  - AppFolio UUID, Name, Phone
  - Tier: 1-Primary or 2-Deep Bench
  - Branch: Phoenix or Tucson
  - Property Group UUID
  - Visibility and Status settings

#### 4. Config (Automation Configuration)

Configure dispatch automation:
- Cron schedule settings for automated passes
- Manual trigger buttons for:
  - Noon Warning Pass
  - Midnight Reassign Pass

#### 5. Audit Trail

System audit events log:
- **Event Filter:** Auto Reassigned, Warnings, Grace Periods, Exemptions, Blasts, Tenant SMS
- **WO ID Filter:** Search by specific work order

#### 6. Blasts (Tier 2 Blast Monitor)

Fires when no Tier 1 tech is available for a job:
- Shows blast notifications sent to Tier 2 pool
- First tech to accept claims the job

#### 7. Comms Log (Tenant Communications)

- **Generate + Send** magic-link portal SMS to tenants
- **Test** portal link sending
- Full communications log feed

---

## Turn Board

Tracks the full lifecycle of unit turns from move-out to move-in ready.

### KPIs

| KPI | Description |
|-----|-------------|
| Confirmed Active | Turns currently in progress |
| On Radar | Awaiting confirmation |
| Avg Days Elapsed | Average days per turn |
| Awaiting Estimates | Turns pending cost estimates |
| Turns In Scope | Turns matching your property group filter |

### Insights Charts

| Chart | Type | Shows |
|-------|------|-------|
| Turn Mix | Pie chart | Status distribution of all turns |
| Stage Pressure | Bar chart | Pipeline bottleneck visualization |
| Age Buckets | Bar chart | Turn aging analysis |

### Turn Pipeline

#### View Modes

- **List View** — Tabular display of all turns
- **Kanban View** — Visual board with drag-and-drop stages

#### Pipeline Stages

```
UPC (Upcoming) → MO (Move-Out) → INS (Inspection) → WO (Work Order Created)
→ REQ (Bidding) → EST (Estimated) → ASN (Approved/Assigned) → DONE (Work Done)
```

#### Filters

| Filter | Options |
|--------|---------|
| Status | Active, On Radar, Upcoming, Stalled (7d+), All Pipeline, Completed, Closed |
| Property Group | Filter by property group |
| Text Search | Search across turn data |

### Turn Detail

Click any turn to open the detail modal showing:
- Move-out and move-in dates
- Elapsed time
- Associated work orders table
- Current stage badge

---

## Inspections

Tracks property inspection schedules, compliance, and completion status.

### KPIs

| KPI | Description |
|-----|-------------|
| Overdue | Inspections past their due date |
| Due Soon (90d) | Inspections due within 90 days |
| Current | On-schedule inspections |
| Turn-Linked | Inspections connected to active turns |

### Insights Charts

| Chart | Type | Shows |
|-------|------|-------|
| Inspection Mix | Pie chart | Status distribution |
| Inspection Age | Bar chart | Aging analysis |
| Turn Coverage | Donut chart | Percentage of turns with linked inspections |

### Inspection Grid

Powered by AG Grid for high-performance data display:

- **Status Filters:** All, Overdue, Due Soon, Current, Turn-Linked
- **Text Search:** Search across inspection data
- **Chart-Driven Filters:** Click chart segments to filter the grid
- **Missing Move-In Inspection** count (highlighted in severe badge)
- **Evidence Policy Tooltip:** Shows move-in/move-out photo requirements

---

## Billing

Comprehensive financial visibility for work orders, payables, and vendor spend.

### KPIs

| KPI | Description |
|-----|-------------|
| Pending Approval | Bills awaiting approval |
| Outstanding ($) | Total outstanding amount |
| Paid This Period | Payments made in current period |
| Vendors (Open Bills) | Number of vendors with open bills |

### Insights Charts

| Chart | Type | Shows |
|-------|------|-------|
| WO Status Mix | Pie chart | Work order billing status breakdown |
| Top Vendors by Spend | Bar chart + Top 5 list | Highest-spend vendors |
| Aging Snapshot | Bar chart | Bill aging distribution |

### Subtabs

#### a) Bills Home

Main billing records table:
- **Columns:** Billing ID, Property, WO Description, Vendor, Amount, WO Status, Created, Last Billed
- **Search:** By WO, property, vendor, billing ID, address
- **Filters:** Date range, WO status
- **Deep Search:** Access bills older than 90 days (with warning modal)
- **Pagination** for large result sets

#### b) Payables (Aged Payables)

Accounts payable aging view:
- **As-of Date:** Select a reference date
- **Party Type:** Vendor, Owner, Occupancy, Management Company
- **Group By:** Detail Rows, Property→Vendor, Vendor→Property
- **Aging Buckets:** Not Yet Due, 0-30, 30-60, 60-90, 90+, 30+, 60+

#### c) Charge Detail

Detailed charge-level data:
- **Date Range** filter
- **Payment Status:** Paid, Unpaid, Partially Paid, All
- **Search:** By tenant, property, account, party ID

#### d) Bill Detail

Bill-level transaction data:
- **Date Range** with date type selector (Bill Date / Due Date / Payment Date)
- **Approval Filter:** Approved, Pending, Denied
- **Search:** By payee, property, account, transaction ID

---

## Manager Review

> **Access:** GM role only

Unified review hub combining tenant tickler events, renewals, GL evidence, and estimate approvals.

### KPIs

Tickler Events / Renewal Rows / GL Leasing Fee Rows / Estimate Approval Pipeline

### Charts

| Chart | Type | Shows |
|-------|------|-------|
| Tickler Event Mix | Pie chart | Event type distribution |
| Renewal Status Mix | Pie chart | Lease renewal status breakdown |
| Estimate Stage Mix | Pie chart | Estimate approval pipeline stages |

### Tables

#### Tenant Tickler

Columns: Date, Event, Property, Tenant, Lease From/To

#### Renewal Summary (No $ Difference)

Columns: Status, Property, Unit, Tenant, Lease Dates, Previous Rent, Rent, Turn

> Filters to show only renewals where rent amount hasn't changed.

#### Leasing Fee Evidence (Renewal + Move-In)

Columns: Date, Property, Unit, Tenant, GL Account, Description, Amount, Charge Type

#### Materials Sent to Estimate for Approval

Columns: Estimate ID, WO#, Source, Status, Property Group, Vendor, Updated

---

## Occupancy

Eight subtabs covering the full tenant lifecycle.

### Subtabs

| Tab | Description | Key Columns |
|-----|-------------|-------------|
| **Transactions** | Tenant Transaction Summary | Charges, Payments, Balance, Last Activity |
| **Tenant Directory** | Full tenant roster | Email, Phone, Lease Start/End, Move-In Date |
| **Vacancies** | Current vacancies | Status, Rent Ready, Days Vacant, Market Rent |
| **Move-Outs** | Upcoming and recent move-outs | Move-out Date, Days Left, Phone, Email |
| **Delinquency** | Outstanding balances | Balance Due, Days Past Due, Last Payment, Status |
| **Applications** | Rental applications | Applicant, Status, Applied On, Decision, Agent |
| **Showings** | Scheduled showings | Prospect, Scheduled, Status, Agent, Notes |
| **Guest Cards** | Walk-in/phone inquiries | Prospect, Interested In, Source, Created, Status |

---

## Properties

Six subtabs for property-level data management.

### Subtabs

#### 1. Directory

Full property listing with:
- Property Name, Address, City/State
- Group, Site Manager, Units
- Bills, Notes, Listings

#### 2. Performance

Property performance metrics:
- Occupancy rate
- Average rent
- Total rent per property

#### 3. Vacancies

Detailed vacant unit list:
- Beds, Baths, Sq Ft
- Market Rent
- Days Vacant

#### 4. Renewals

Lease renewal tracking:
- Status (Renewed / Did Not Renew / MTM / Pending / Canceled)
- Rent changes
- Agent

#### 5. Owner Reports

Owner distribution reports

#### 6. Bulk Notes

Mass update property notes:
1. Apply filters to narrow properties
2. Click **Bulk Note Composer**
3. Compose your note
4. Apply to all filtered properties

### Features

- Search and group filter
- Pagination (25 / 50 / 100 rows per page)

---

## Vendors

Comprehensive vendor directory with compliance tracking.

### Filters

| Filter | Options |
|--------|---------|
| Category | Employee, In-House Tech, Vendor, Subcontractor, Utilities, HOA, Insurance, Uncategorized |
| Trade | Dynamic dropdown of maintenance trades |
| Property | Filter by vendors who have worked at a specific property |
| Sort Mode | Group by Trade, Name A-Z, Most Recent Work |
| Text Search | Search across all vendor fields |

### Vendor Grid

High-performance AG Grid display with pagination.

### Vendor Detail

Click any vendor to open detail modal showing:
- **Contact:** Phone, Email
- **Compliance:** Liability Insurance, Workers Comp
- **Warnings:** Expired document alerts, "DO NOT USE" banners
- **Quick Links:** Open WOs, Completed WOs, Bills, AppFolio link

---

## Communication Templates

Create and manage reusable message templates for automated notes.

### Features

- Template grid with all existing templates
- **New Template** button to create templates
- Auto-fill with API attributes (dynamic placeholders)
- PostCreate targets:
  - Work Order Note
  - Tenant Note
  - Owner Note

---

## Database Admin

> **Access:** GM and Admin roles only

Direct database access and system configuration tools.

### Postgres Search

- **Table Selector:** Work Orders, Properties, Units, Vendors, Tech Roster, Dispatch Queue
- **Text Search:** Search across all columns in selected table
- **Shortcut Buttons:** Quick access to common table views
- **Read-only:** All queries are read-only with a 200-row cap

### Turn Closure History

View historical turn closures and their data.

### PM OTP Accounts Manager

Manage PM user accounts:
- **Add/Edit PM Users:**
  - Email, Full Name, Phone
  - Property Group UUIDs (multi-scope support)
  - Active status

### OTP & Auth Policy Settings

| Setting | Description |
|---------|-------------|
| OTP Login Enabled | Toggle OTP authentication on/off |
| Require PM Membership | Require PM record for OTP login |
| Allowed Email Domain | Restrict OTP to specific email domain |
| OTP TTL Minutes | Code expiration (3-30 minutes) |

---

## Error Log

System monitoring and troubleshooting tools.

### API Log

#### Systems Active Checker

Monitor system health across:
- SQL, Middleware, API v0/v2, Magic Links, Reassignment, Routing, PM Users

#### Shortlink Rate Limit Manager

- IP-based rate limiting for shortlinks
- Unblock tools for blocked IPs
- Rate limit configuration

#### Error Log

- Exponential backoff tracking (533 errors)
- Retry-After monitoring (429 errors)
- Semantic validation errors (422)
- Webhook configuration modal
- **Clear Resolved** button to clean up resolved errors

### Email Delivery

Monitor email delivery failures from AppFolio:

| KPI | Description |
|-----|-------------|
| Errors In Range | Total failures in selected date range |
| Invalid Address | Bounced due to invalid email |
| Rejected/Spam | Rejected or marked as spam |
| Repeat Recipients | Tenants/owners with multiple failures |

**Filters:** Date range, Party Type (Tenants/Owners), Failure Type, Search

---

## Global Property Group Filter

The property group filter is a persistent bar below the top bar that scopes **all data across every section** of the application.

### How It Works

1. Select a property group from the dropdown, or choose "All Properties"
2. Every section (Dashboard, Work Orders, Turn Board, Billing, etc.) automatically filters to show only data for that group
3. Badge counts on navigation tabs update to reflect scoped totals
4. Charts and KPIs recalculate for the selected scope

### Syncing Groups

Click **Sync Groups** to pull the latest property group list from AppFolio.

### PM Scoped Sessions

PM users are automatically scoped to their assigned property groups. They can only see data for the groups they manage.

---

## Settings & PWA

### App Settings (Gear Icon)

| Setting | Description |
|---------|-------------|
| Cache Export | Export all cached data as a JSON or HMC (AES-GCM encrypted) file |
| Cache Import | Import a previously exported cache file |
| Issue Reporting | Submit bug reports or feature requests |

### Theme System

- **Light Mode** — Clean white/gray design
- **Dark Mode** — Easy on the eyes for low-light environments
- Toggle from the top bar or login screen
- Preference persists in browser storage

### PWA Features

| Feature | Description |
|---------|-------------|
| Installable | Add to home screen / desktop as a standalone app |
| Offline Support | Core navigation and cached data work without internet |
| Background Sync | Mutations made offline are queued and synced when connectivity returns |
| Auto-Update | Service worker checks for updates and prompts to refresh |

### Maintenance Mode

- A scrolling maintenance banner may appear when the system is under maintenance
- Post-login maintenance dialogs notify you of planned downtime
- The banner is collapsible and its state is remembered

---

## Keyboard Shortcuts & Tips

- **Click KPI cards** on the Dashboard to drill into detailed views
- **Use the Property Group filter** to focus on a specific portfolio section
- **TV Mode** on the Turn Progress dashboard for lobby/display screens
- **Kanban views** available on Work Orders and Turn Board for visual management
- **Click chart segments** to filter underlying data tables
- **Export cache** regularly for backup purposes

---

## Data Sync

HandyManager syncs data from AppFolio automatically. The following data types are synced:

| Data Type | Source |
|-----------|--------|
| Properties | AppFolio API v0 |
| Property Groups | AppFolio API v0 |
| Units | AppFolio API v0 |
| Work Orders | AppFolio API v0 |
| Vendors | AppFolio API v0 |
| Bills | AppFolio API v0 |
| Estimates | AppFolio API v2 |
| Inspections | AppFolio API v2 |
| Tenant Directory | AppFolio API v2 |
| Vacancies | AppFolio API v2 |
| Move-Outs | AppFolio API v2 |
| Tickler Events | AppFolio API v2 |
| Renewals | AppFolio API v2 |
| GL Ledger | AppFolio API v2 |
| Email Errors | AppFolio API v2 |
| Payables | AppFolio API v2 |
| Charge Detail | AppFolio API v2 |
| Bill Detail | AppFolio API v2 |

The sync timestamp is displayed in the top bar. Real-time webhook events from AppFolio appear in the webhook feed drawer.

---

*HandyManager v9.8.36 — Fort Lowell Realty & Property Management (FLRAZ) — handymgr.app*
