-- Migration: 2026-06-20_work_orders_group_backfill
-- Backfill work-order property_group_id from properties when missing.
-- This improves property-group filtering for historical work orders.

-- Speed up one-time/periodic backfill scans for missing work-order groups.
CREATE INDEX IF NOT EXISTS appfolio_work_orders_property_missing_group_idx
  ON appfolio_work_orders(property_id)
  WHERE property_group_id IS NULL OR btrim(property_group_id) = '';

-- Backfill using canonical property group from appfolio_properties.
UPDATE appfolio_work_orders wo
SET property_group_id = p.property_group_id
FROM appfolio_properties p
WHERE wo.property_id = p.id
  AND p.property_group_id IS NOT NULL
  AND btrim(p.property_group_id) <> ''
  AND (wo.property_group_id IS NULL OR btrim(wo.property_group_id) = '');
