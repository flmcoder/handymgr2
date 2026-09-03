-- Support unrestricted active Work Order reads without scanning inactive history.
CREATE INDEX CONCURRENTLY IF NOT EXISTS appfolio_work_orders_active_updated_idx
  ON appfolio_work_orders((COALESCE(updated_at, created_at)) DESC)
  WHERE COALESCE(LOWER(status), '') NOT LIKE '%completed%'
    AND COALESCE(LOWER(status), '') NOT LIKE '%cancel%'
    AND COALESCE(LOWER(status), '') NOT LIKE '%no need to bill%';

CREATE INDEX CONCURRENTLY IF NOT EXISTS appfolio_work_orders_active_group_updated_idx
  ON appfolio_work_orders(property_group_id, (COALESCE(updated_at, created_at)) DESC)
  WHERE COALESCE(LOWER(status), '') NOT LIKE '%completed%'
    AND COALESCE(LOWER(status), '') NOT LIKE '%cancel%'
    AND COALESCE(LOWER(status), '') NOT LIKE '%no need to bill%';