# User Journey: Work Order Operations Queue

## User Persona
- Who: Maintenance dispatcher or property manager coordinating properties, internal technicians, and vendors
- Goal: Triage, assign, investigate, and monitor Work Orders without leaving the operational workspace
- Context: Frequent, interruption-heavy use across desktop and field devices
- Success metric: Critical WOs are identified and routed in minutes, with no missed item caused by filtering or inaccessible interaction

## Stage 1: Intake
What user is doing: Opens Work Orders and scans current workload.
What user is thinking: What needs attention first right now?
What user is feeling: Alert, time-pressured.
Pain points:
- Too much visual noise when board is default for volume triage.
- Active/inactive context switching is not always obvious.
Opportunity:
- Queue-first default with clear active/inactive segmentation.
- KPI strip surfaces total active volume and urgent/unassigned risk without requiring chart interpretation.

## Stage 2: Triage
What user is doing: Uses search and filters (priority, vendor, property, age).
What user is thinking: Can I isolate only high-risk items quickly?
What user is feeling: Focused but impatient.
Pain points:
- Small, inconsistent controls increase click precision cost.
- Varying corner radii create weak visual hierarchy.
Opportunity:
- Standardized control sizing and subtle squared corners for operational clarity.
- Interactive analytics apply named filters to the grid and expose equivalent keyboard controls.

## Stage 3: Action
What user is doing: Selects a grid row, reviews its side panel, assigns or updates the Work Order, and checks service history.
What user is thinking: Is this the right owner, and what happened before I act?
What user is feeling: Decisive, but checking for completeness.
Pain points:
- Context changes between queue and detail can feel abrupt.
Opportunity:
- Keep the grid visible while a sticky details panel supplies issue and service-history context.
- On narrow screens, preserve the same context in a dismissible bottom sheet or full-height drawer.

## Stage 4: Monitoring
What user is doing: Uses closure assistant and follow-up queue.
What user is thinking: What is still stale or blocked?
What user is feeling: Cautiously confident.
Pain points:
- Follow-up work can get buried when visual cues are inconsistent.
Opportunity:
- Stable status tabs, compact counters, and persistent removable filter chips keep backlog visibility high.

## Cross-Role Tension
- Dispatchers need speed and breadth; property managers need scoped context; portfolio managers need comparison.
- Use one shared information architecture, then tailor default filters and saved views when reliable user-role metadata exists.
- Never treat a saved view or role default as authorization. Server-enforced property-group scope remains authoritative.