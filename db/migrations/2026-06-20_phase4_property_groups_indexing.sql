-- Migration: 2026-06-20_phase4_property_groups_indexing
-- Phase 4 hardening for property-group scoped local reads.
-- Adds indexes that match the new local query patterns for work orders,
-- properties, units, and turn-linked work order lookups.

-- Work orders: date-window scans + optional property_group_id scope.
CREATE INDEX IF NOT EXISTS appfolio_work_orders_group_updated_idx
  ON appfolio_work_orders(property_group_id, updated_at DESC);

CREATE INDEX IF NOT EXISTS appfolio_work_orders_property_updated_idx
  ON appfolio_work_orders(property_id, updated_at DESC);

CREATE INDEX IF NOT EXISTS appfolio_work_orders_effective_updated_idx
  ON appfolio_work_orders((COALESCE(updated_at, created_at)));

-- Properties: group filter + name sort.
CREATE INDEX IF NOT EXISTS appfolio_properties_group_name_idx
  ON appfolio_properties(property_group_id, name);

-- Units: scoped property reads sorted by unit number.
CREATE INDEX IF NOT EXISTS appfolio_units_property_unit_number_idx
  ON appfolio_units(property_id, unit_number);

-- Turn linkage: reduce join/filter cost in /api/local/turn_work_orders.
CREATE INDEX IF NOT EXISTS unit_turn_work_orders_wo_db_uuid_idx
  ON unit_turn_work_orders(wo_db_uuid);

CREATE INDEX IF NOT EXISTS unit_turn_work_orders_active_created_idx
  ON unit_turn_work_orders(created_at DESC)
  WHERE removed IS NOT TRUE;
