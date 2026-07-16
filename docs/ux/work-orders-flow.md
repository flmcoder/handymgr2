# User Flow: Work Order Operations

## Entry Point
User opens Work Order Operations tab from primary navigation.

## Flow Steps
1. Work Orders landing (default queue/list)
- Sees active/inactive tabs with counts.
- Sees search and compact filter controls.
- Primary action: identify next urgent/stale WO.

2. Triage controls
- Apply priority, type, vendor, property, and age filters.
- Sort by oldest/newest/priority.
- Save aging thresholds.

3. View mode switch
- Default: Queue view (high-density triage).
- Optional: Kanban view for status-flow management.
- View preference persists for returning users.

4. Action execution
- Open WO detail.
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

## Accessibility Requirements
### Keyboard Navigation
- All filters, tabs, and view toggles reachable by tab order.
- Active state must be visible on focus and selection.

### Screen Reader Support
- Toggle buttons include clear labels (Queue, Kanban, Active, Inactive).
- Counters are exposed as text, not color-only indicators.

### Visual Accessibility
- Maintain AA contrast for active/inactive states.
- Minimum control height near 34px for mixed mouse/touch environments.
- Do not rely on color alone; include labels, icons, and count badges.

## Figma Handoff Notes
- Build one desktop-first triage frame for Queue default.
- Build one alternate frame for Kanban mode with same control language.
- Keep tabs, badges, and filter geometry tokenized for reuse.