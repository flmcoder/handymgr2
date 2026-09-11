-- Inspections scope-join support. The tenant directory only carries the numeric
-- AppFolio property id (e.g. '6349'); it maps to appfolio_properties through the
-- numeric tail of raw_json->>'Link', and property-group scope membership follows
-- the raw_json->>'PropertyGroupIds' array (a property can belong to several
-- groups). These indexes make the scoped/lateral nav-badge + grid queries fast
-- even on a cold cache.

create index if not exists appfolio_unit_inspections_occupancy_id_idx
  on appfolio_unit_inspections(occupancy_id);

create index if not exists appfolio_properties_link_numeric_idx
  on appfolio_properties ((raw_json ->> 'Link'));