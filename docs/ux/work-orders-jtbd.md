# Work Orders JTBD

## User Personas
- Primary: Maintenance dispatchers triaging, assigning, and monitoring work across teams
- Primary: Property managers managing assigned properties and tenant outcomes
- Secondary: Portfolio managers reviewing risk and operational performance across groups
- Skill level: Mixed, from frequent operational users to occasional reviewers
- Devices: Responsive across desktop, tablet, and mobile; desktop remains the highest-density mode
- Accessibility priorities: Keyboard-only operation and forgiving touch/motor targets
- Risk of failure: High; a missed or incorrectly filtered Work Order can affect safety, habitability, cost, and tenant trust

## Job Statement
When I am managing daily maintenance demand, I want to triage, assign, investigate, and monitor Work Orders from one trustworthy workspace, so I can keep urgent jobs moving without losing context or missing safety-critical follow-up.

## Role-Specific Jobs
- Dispatcher: When new demand arrives, I want urgent and unassigned work surfaced first, so I can route it within minutes.
- Property manager: When I review my properties, I want current status, assignee, issue context, and service history together, so I can resolve tenant-impacting work without switching screens.
- Portfolio manager: When I assess operations, I want comparable aging, ownership, and status analytics, so I can find systemic risk and overloaded teams.

## Current Solution And Pain Points
- Current: Kanban-first view with useful context, plus list option
- Pain: Board is excellent for status flow, but not always the fastest for high-volume triage
- Pain: Control density and button geometry are inconsistent, which slows scanning and decision speed
- Pain: Inline styles and mixed affordances make state changes less obvious
- Pain: Charts can become passive decoration if they do not filter the operational queue
- Pain: Opening details in a separate context interrupts triage and makes comparison harder
- Pain: Desktop-only density can make the same workflow unusable on touch devices

## Incumbent Workflow
- Open work orders tab
- Search/filter by priority/vendor/property
- Jump between active/inactive and status columns
- Open details and issue follow-up actions

## Success Criteria
- Default experience prioritizes queue triage speed
- Kanban remains one click away for status-flow management
- Control language and shape system feel consistent across work-order actions
- Users can process the next high-priority action with less visual friction
- Every chart selection produces a visible, removable grid filter
- Selecting a row preserves queue position while showing service context
- Keyboard and touch users can perform the same filtering and detail-review tasks
- Future role metadata may tailor defaults, but it must not change authorization scope