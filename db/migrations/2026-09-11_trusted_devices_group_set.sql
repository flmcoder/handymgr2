-- Union scope support: a PM session can carry every assigned property group,
-- not just one. The legacy single column stays as primary/first for compat.
alter table trusted_devices
  add column if not exists property_group_uuids text not null default '';