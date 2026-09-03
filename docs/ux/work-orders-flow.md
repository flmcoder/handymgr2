# User Flow: Work Order Operations

## Entry Point
User opens Work Order Operations tab from primary navigation.

## Flow Steps
1. Work Orders landing (default queue/list)
- Sees `+ New Work Order`, Total WOs, and Urgent / Critical in one compact KPI row.
- KPI cards include a numeric value, textual trend, and a 40px sparkline; the value remains understandable when charts are unavailable.
- Sees Active and Completed / Inactive tabs with counts.
- Primary action: identify and route the next urgent or unassigned Work Order.

2. Visual triage
- Aging Buckets and Status by Owner charts sit directly above the operational workspace.
- Selecting a chart value applies a named AG Grid filter and adds a removable filter chip.
- Keyboard users can focus each chart and select the same values through an adjacent accessible summary/list.

3. Advanced controls
- Apply status, priority, type, vendor, property, owner, and age filters from one toolbar.
- Search, sort, clear all, and saved-view controls remain visible above the grid.
- Filter state persists through row selection and details-panel actions.

4. View mode switch
- Default: Queue view (high-density triage).
- Optional: Kanban view for status-flow management.
- View preference persists for returning users.

5. Action execution
- Select a row without navigating away from the queue.
- Review WO ID, assignee, urgency, description, and service-history timeline in the contextual panel.
- Update status/priority and coordinate vendor or internal assignee.
- Use closure assistant/follow-up queue for unresolved items.

## Exit Points
- Success: high-priority work is assigned or advanced.
- Partial: filters saved and queue narrowed for next pass.
- Blocked: missing owner/vendor context triggers follow-up path.

## Design Principles
1. Operational Clarity First
- Prefer scannable queue defaults for high-volume triage.

2. Kanban As Specialized Mode
- Preserve board for stage-flow analysis without making it mandatory.

3. Shape Consistency
- Use subtle squared corners (6px range) for controls and tabs.

4. Dense, Legible Controls
- Keep controls compact but clearly clickable.

5. Analytics Must Be Operational
- Charts are controls, not decoration; every selectable mark maps to a documented grid filter.
- Always show the resulting filter as text and provide a one-step clear action.

6. Context Without Disorientation
- Keep the selected row visible while details are open.
- Preserve scroll, sort, and filters when the panel closes or the user selects another row.

## Figma Layout Specification
### Theme Tokens
- Light: `bg-base #f1f5f9`, `bg-surface #ffffff`, `border-color #e2e8f0`, `text-primary #0f172a`, `text-muted #64748b`.
- Dark: `bg-base #0f172a`, `bg-surface #1e293b`, `border-color #334155`, `text-primary #f8fafc`, `text-muted #94a3b8`.
- Semantic: `status-critical #ef4444`, `status-progress #0ea5e9`, `status-assigned #f59e0b`.
- Scope boundary: apply these aliases within `#sec-workorders`; do not restyle `#vaultScreen`, login, auth, vault effects, or global settings.

### Desktop, 1280px And Wider
- KPI row: horizontal flex, 12px gap; New Work Order command first, followed by equal-width KPI cards.
- KPI card: minimum 168px wide, 64px high, 12px padding, 6px radius, 1px border.
- KPI text occupies the left half; sparkline host occupies the right half and is exactly 40px high.
- Analytics: 12-column grid directly above the workspace; Aging and Owner charts each span 6 columns.
- Main workspace: `minmax(0, 1fr) 350px`, 16px gap.
- Details panel: sticky below the app header, viewport-bounded, internally scrollable.

### Tablet, 768px To 1279px
- KPI row wraps; primary command remains first and visible.
- Analytics stack to one chart per row when either chart would fall below 360px.
- Main workspace becomes one column; details open as a right drawer at 420px maximum width.

### Mobile, Below 768px
- KPI row horizontally scrolls with a visible next-item hint; do not shrink cards below 156px.
- Filters open from one labeled filter command; active filters remain visible as horizontally scrollable chips.
- Grid uses a reduced essential column set: priority/status, WO ID, property/unit, age.
- Details open as a full-width bottom sheet or full-height drawer with a persistent close control.

## Component Specification
### KPI Cards And Sparklines
- Required DOM roles: heading/label, value, textual trend, and chart host.
- Chart host class: `wo-chart-host`; sparkline variant must have `height: 40px; width: 100%` and a stable parent width.
- Sparkline has no axis labels or interaction; its trend is repeated in text for accessibility.

### Visual Analytics
- Chart cards use a 6px radius, 1px border, surface background, compact title, and optional collapse control.
- Chart initialization reads the active theme and uses ECharts `dark` only in dark mode.
- Theme changes dispose and recreate instances; do not call `setOption` across incompatible themes.
- One `ResizeObserver` observes every `.wo-chart-host` and calls the associated instance's `resize()`.
- Before re-wiring, remove prior click listeners to prevent duplicate filter application.
- Chart click mapping:
	- Aging bucket -> AG Grid age bucket filter.
	- Status by Owner bar/pie -> AG Grid owner and optional status filter.
- Chart selection must update `aria-pressed` or an adjacent textual state and announce the applied filter.

### Grid Workspace
- Toolbar order: search, status, priority, type, owner/vendor, property, age, saved view, clear all.
- AG Grid theme aliases map background, foreground, border, row-border, header-background, odd-row-background, selected-row, and input colors to Work Orders tokens.
- Status uses pill badges with text plus an icon/shape; progress may use `rgba(14, 165, 233, 0.15)` and `#0ea5e9`.
- Row selection opens details but never changes filters or scroll position.

### Contextual Details Panel
- Header: WO ID, urgency badge, close command.
- Summary: assignee, property/unit, issue description, age, and current status.
- Service History: semantic `ul`; 2px left border using `border-color`; each `li` has a positioned bullet, timestamp, actor, and action text.
- Empty selection: concise prompt to select a Work Order, without decorative illustration.

## Accessibility Requirements
### Keyboard Navigation
- All filters, tabs, and view toggles reachable by tab order.
- Active state must be visible on focus and selection.
- Charts require a keyboard-equivalent list or controls; canvas click handling alone is insufficient.
- Escape closes the mobile/tablet details drawer and returns focus to the selected row.

### Screen Reader Support
- Toggle buttons include clear labels (Queue, Kanban, Active, Inactive).
- Counters are exposed as text, not color-only indicators.
- Dynamic chart/grid filter changes are announced through a polite live region.
- Service-history timestamps use semantic `time` elements.

### Visual Accessibility
- Maintain AA contrast for active/inactive states.
- Minimum control height near 34px for mixed mouse/touch environments.
- Touch-first commands and drawer controls use a minimum 44px target.
- Do not rely on color alone; include labels, icons, and count badges.

## Figma Handoff Notes
- Build one desktop-first triage frame for Queue default.
- Build one alternate frame for Kanban mode with same control language.
- Keep tabs, badges, and filter geometry tokenized for reuse.
- Build desktop, tablet drawer, and mobile bottom-sheet frames.
- Prototype chart-to-filter, row-to-details, clear-filter, and focus-return interactions.
- Validate urgent/unassigned discovery with dispatcher and property-manager participants before introducing role-specific defaults.

## Existing Implementation Map
- Scope root: `#sec-workorders`. All new selectors and DOM mounts should remain descendants of this section.
- Existing primary command: `#btnNewWO`; move it into the KPI strip rather than creating a second command.
- Existing filters: `#woFilters`, `#woPriorityFilter`, `#woTypeFilter`, `#woVendorFilter`, `#woPropertyFilter`, `#woAgeFilter`, and `#woSort`.
- Existing applied-filter summary: `#woChartFilterBadges`.
- Existing analytics: `#woAgingChart`, `#woOwnerChart`, `#woStatusChart`, and `#woVendorSpendChart`, all already using `.wo-chart-host`.
- Existing grid: `#woGridHost.ag-theme-quartz.hm-ag-theme--wo`.
- Existing details behavior should be extended with one contextual panel; do not add another Work Order grid or analytics shell.
- Production currently expresses visual theme through body classes such as `theme-neo`, not a `data-theme` attribute. Engineering should use the existing theme toggle as the source of truth and observe body class changes when reinitializing ECharts.
- The production Work Orders section is hidden until selected. Initialize or resize charts when the section becomes visible, not only at page load.
- Explicitly excluded: `#vaultScreen`, login panels, authentication modules, vault effects, and global settings markup or behavior.

## Role Personalization Follow-Up
- Add role metadata only through the existing user database and administrative user-management workflow.
- Dispatcher default: all authorized groups, active queue, urgent/unassigned first.
- Property-manager default: server-authorized assigned group, active queue, operational priority sort.
- Portfolio-manager default: authorized groups, analytics expanded, saved comparison view.
- Users may override display preferences; role defaults never broaden server authorization.