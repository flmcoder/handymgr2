export const TURN_ENGINE_SQL = String.raw`
with native_turns as (
  select
    'native:' || d.turn_id as turn_key,
    d.turn_id as unit_turn_id,
    d.unit_id,
    d.property_id,
    d.unit_name,
    d.property_name,
    d.move_out_date,
    d.expected_move_in_date,
    d.turn_end_date,
    d.unit_turn_status,
    true as is_native_turn
  from appfolio_unit_turn_details d
  where d.unit_id is not null
    and d.move_out_date is not null
),
notice_occupancies as (
  select distinct on (
    td.unit_id,
    coalesce(td.move_out_date, td.lease_to, current_date::timestamptz)
  )
    td.record_id,
    td.unit_id,
    td.property_id,
    td.unit_name,
    td.property_name,
    td.tenant_name,
    coalesce(td.move_out_date, td.lease_to, current_date::timestamptz) as move_out_date
  from appfolio_tenant_directory td
  where td.unit_id is not null
    and (
      td.move_out_date is not null
      or lower(coalesce(td.status, '')) = 'notice'
    )
  order by
    td.unit_id,
    coalesce(td.move_out_date, td.lease_to, current_date::timestamptz),
    td.last_updated_at desc nulls last,
    td.cached_at desc
),
shadow_turns as (
  select
    'shadow:' || o.unit_id || ':' || to_char(o.move_out_date at time zone 'UTC', 'YYYY-MM-DD') as turn_key,
    null::text as unit_turn_id,
    o.unit_id,
    o.property_id,
    o.unit_name,
    o.property_name,
    o.move_out_date,
    null::timestamptz as expected_move_in_date,
    null::timestamptz as turn_end_date,
    'Shadow Turn'::text as unit_turn_status,
    false as is_native_turn
  from notice_occupancies o
  where not exists (
    select 1
    from native_turns n
    where n.unit_id = o.unit_id
      and n.move_out_date::date between o.move_out_date::date - 7 and o.move_out_date::date + 7
  )
),
turn_candidates as (
  select * from native_turns
  union all
  select * from shadow_turns
),
turn_bounds as (
  select
    tc.*,
    next_resident.move_in_date as next_move_in_date,
    next_resident.tenant_name as next_resident_name,
    current_resident.has_current_resident
  from turn_candidates tc
  left join lateral (
    select td.move_in_date, td.tenant_name
    from appfolio_tenant_directory td
    where td.unit_id = tc.unit_id
      and td.move_in_date is not null
      and td.move_in_date >= tc.move_out_date
    order by td.move_in_date asc
    limit 1
  ) next_resident on true
  left join lateral (
    select true as has_current_resident
    from appfolio_tenant_directory td
    where td.unit_id = tc.unit_id
      and lower(coalesce(td.status, '')) = 'current'
    limit 1
  ) current_resident on true
),
work_order_base as (
  select
    wo.id,
    wo.work_order_uuid,
    wo.wo_number,
    wo.unit_id,
    wo.status,
    wo.description,
    wo.created_at,
    coalesce(
      nullif(wo.raw_json->>'UnitTurnId', ''),
      nullif(wo.raw_json->>'unit_turn_id', '')
    ) as linked_unit_turn_id,
    coalesce(
      nullif(wo.raw_json->>'UnitTurnCategory', ''),
      nullif(wo.raw_json->>'unit_turn_category', ''),
      nullif(wo.category, '')
    ) as unit_turn_category,
    coalesce(
      nullif(wo.raw_json->>'CurrentEstimateApprovalStatus', ''),
      nullif(wo.raw_json->>'current_estimate_approval_status', '')
    ) as estimate_approval_status,
    coalesce(
      nullif(wo.vendor_name, ''),
      nullif(wo.raw_json->>'VendorName', ''),
      nullif(wo.raw_json->>'vendor_name', '')
    ) as vendor_name,
    coalesce(
      nullif(wo.raw_json->>'VendorTrade', ''),
      nullif(wo.raw_json->>'vendor_trade', '')
    ) as vendor_trade,
    (
      nullif(wo.assigned_user_id, '') is not null
      or nullif(wo.assigned_user_name, '') is not null
      or case
        when jsonb_typeof(wo.raw_json->'AssignedUsers') = 'array'
          then jsonb_array_length(wo.raw_json->'AssignedUsers') > 0
        when jsonb_typeof(wo.raw_json->'assigned_users') = 'array'
          then jsonb_array_length(wo.raw_json->'assigned_users') > 0
        else false
      end
    ) as has_assigned_user,
    coalesce(
      nullif(regexp_replace(coalesce(
        wo.raw_json->>'VendorBillAmount',
        wo.raw_json->>'vendor_bill_amount',
        ''
      ), '[^0-9.-]', '', 'g'), '')::numeric,
      0
    ) as vendor_bill_amount,
    coalesce(
      nullif(regexp_replace(coalesce(
        wo.raw_json->>'TotalCost',
        wo.raw_json->>'total_cost',
        wo.total_cost::text,
        wo.estimated_amount::text,
        ''
      ), '[^0-9.-]', '', 'g'), '')::numeric,
      0
    ) as work_order_cost
  from appfolio_work_orders wo
),
swept_wos as (
  select tb.turn_key, wo.*
  from turn_bounds tb
  join work_order_base wo
    on wo.unit_id = tb.unit_id
   and wo.created_at >= tb.move_out_date
   and wo.created_at <= coalesce(tb.next_move_in_date, current_date::timestamptz)
),
work_order_rollup as (
  select
    sw.turn_key,
    count(*)::int as work_order_count,
    count(*) filter (where sw.linked_unit_turn_id is null)::int as rogue_wos_detected,
    bool_and(lower(coalesce(sw.status, '')) = 'completed') as all_work_orders_completed,
    coalesce(sum(case
      when sw.has_assigned_user or coalesce(sw.vendor_name, '') ilike '%Fort Lowell%'
        then sw.work_order_cost
      else 0
    end), 0) as in_house_cost,
    coalesce(sum(case
      when coalesce(sw.vendor_name, '') <> ''
       and sw.vendor_name not ilike '%Fort Lowell%'
        then sw.vendor_bill_amount
      else 0
    end), 0) as third_party_cost,
    count(*) filter (where sw.unit_turn_category = '7')::int as locks_count,
    bool_and(lower(coalesce(sw.status, '')) = 'completed') filter (where sw.unit_turn_category = '7') as locks_completed,
    count(*) filter (where lower(coalesce(sw.vendor_trade, '')) like '%lock%')::int as lock_trade_count,
    bool_and(lower(coalesce(sw.status, '')) = 'completed') filter (where lower(coalesce(sw.vendor_trade, '')) like '%lock%') as lock_trade_completed,
    count(*) filter (where sw.estimate_approval_status is not null)::int as estimate_count,
    count(*) filter (where lower(coalesce(sw.estimate_approval_status, '')) = 'approved')::int as estimate_approved_count,
    count(*) filter (where sw.unit_turn_category = '1')::int as maintenance_count,
    bool_and(lower(coalesce(sw.status, '')) = 'completed') filter (where sw.unit_turn_category = '1') as maintenance_completed,
    count(*) filter (where sw.unit_turn_category = '2')::int as paint_count,
    bool_and(lower(coalesce(sw.status, '')) = 'completed') filter (where sw.unit_turn_category = '2') as paint_completed,
    count(*) filter (where sw.unit_turn_category = '4')::int as flooring_count,
    bool_and(lower(coalesce(sw.status, '')) = 'completed') filter (where sw.unit_turn_category = '4') as flooring_completed,
    count(*) filter (where sw.unit_turn_category = '6')::int as cleaning_count,
    bool_and(lower(coalesce(sw.status, '')) = 'completed') filter (where sw.unit_turn_category = '6') as cleaning_completed,
    count(*) filter (where sw.unit_turn_category = '3')::int as appliance_count,
    bool_and(lower(coalesce(sw.status, '')) = 'completed') filter (where sw.unit_turn_category = '3') as appliance_completed,
    jsonb_agg(jsonb_build_object(
      'id', sw.id,
      'wo_number', sw.wo_number,
      'work_order_uuid', sw.work_order_uuid,
      'status', sw.status,
      'description', sw.description,
      'created_at', sw.created_at,
      'unit_turn_id', sw.linked_unit_turn_id,
      'unit_turn_category', sw.unit_turn_category,
      'vendor_name', sw.vendor_name,
      'vendor_trade', sw.vendor_trade,
      'vendor_bill_amount', sw.vendor_bill_amount,
      'cost', sw.work_order_cost
    ) order by sw.created_at, sw.id) as swept_work_orders
  from swept_wos sw
  group by sw.turn_key
),
turn_evidence as (
  select
    tb.*,
    coalesce(wr.work_order_count, 0) as work_order_count,
    coalesce(wr.rogue_wos_detected, 0) as rogue_wos_detected,
    coalesce(wr.all_work_orders_completed, false) as all_work_orders_completed,
    coalesce(wr.in_house_cost, 0) as in_house_cost,
    coalesce(wr.third_party_cost, 0) as third_party_cost,
    coalesce(wr.swept_work_orders, '[]'::jsonb) as swept_work_orders,
    coalesce(wr.locks_count, 0) + coalesce(wr.lock_trade_count, 0) as locks_count,
    coalesce(wr.locks_completed, true) and coalesce(wr.lock_trade_completed, true) as locks_completed,
    coalesce(wr.estimate_count, 0) as estimate_count,
    coalesce(wr.estimate_approved_count, 0) as estimate_approved_count,
    coalesce(wr.maintenance_count, 0) as maintenance_count,
    coalesce(wr.maintenance_completed, false) as maintenance_completed,
    coalesce(wr.paint_count, 0) as paint_count,
    coalesce(wr.paint_completed, false) as paint_completed,
    coalesce(wr.flooring_count, 0) as flooring_count,
    coalesce(wr.flooring_completed, false) as flooring_completed,
    coalesce(wr.cleaning_count, 0) as cleaning_count,
    coalesce(wr.cleaning_completed, false) as cleaning_completed,
    coalesce(wr.appliance_count, 0) as appliance_count,
    coalesce(wr.appliance_completed, false) as appliance_completed,
    inspection.inspection_date,
    coalesce(vacancy.rent_ready, false) as rent_ready,
    vacancy.ready_for_showing_on
  from turn_bounds tb
  left join work_order_rollup wr on wr.turn_key = tb.turn_key
  left join lateral (
    select ui.last_inspection_date as inspection_date
    from appfolio_unit_inspections ui
    where ui.unit_id = tb.unit_id
      and ui.last_inspection_date >= tb.move_out_date
      and ui.last_inspection_date <= coalesce(tb.next_move_in_date, now())
    order by ui.last_inspection_date asc
    limit 1
  ) inspection on true
  left join lateral (
    select
      lower(coalesce(
        uv.raw_json->>'RentReady',
        uv.raw_json->>'rent_ready',
        'false'
      )) in ('true', 't', 'yes', '1') as rent_ready,
      case
        when coalesce(uv.raw_json->>'ReadyForShowingOn', uv.raw_json->>'ready_for_showing_on', '') ~ '^\d{4}-\d{2}-\d{2}'
          then left(coalesce(uv.raw_json->>'ReadyForShowingOn', uv.raw_json->>'ready_for_showing_on'), 10)::date
        else null
      end as ready_for_showing_on
    from appfolio_unit_vacancies uv
    where uv.unit_id = tb.unit_id
    order by coalesce(uv.last_updated_at, uv.cached_at) desc
    limit 1
  ) vacancy on true
),
milestone_rows as (
  select
    te.*,
    jsonb_build_array(
      jsonb_build_object('key', 'move_out_recorded', 'label', 'Move-Out Recorded', 'status', case when te.move_out_date <= now() then 'completed' else 'in_progress' end),
      jsonb_build_object('key', 'move_out_inspection', 'label', 'Move-Out Inspection', 'status', case when te.inspection_date is not null then 'completed' when te.move_out_date <= now() then 'in_progress' else 'not_started' end),
      jsonb_build_object('key', 'locks_rekeyed', 'label', 'Locks Rekeyed', 'status', case when te.locks_count = 0 then 'not_started' when te.locks_completed then 'completed' else 'in_progress' end),
      jsonb_build_object('key', 'estimates_approved', 'label', 'Estimates Approved', 'status', case when te.estimate_approved_count > 0 then 'completed' when te.estimate_count > 0 then 'in_progress' else 'not_started' end),
      jsonb_build_object('key', 'maintenance_repair', 'label', 'Maintenance / Repair', 'status', case when te.maintenance_count = 0 then 'not_started' when te.maintenance_completed then 'completed' else 'in_progress' end),
      jsonb_build_object('key', 'paint', 'label', 'Paint', 'status', case when te.paint_count = 0 then 'not_started' when te.paint_completed then 'completed' else 'in_progress' end),
      jsonb_build_object('key', 'flooring_carpet', 'label', 'Flooring / Carpet', 'status', case when te.flooring_count = 0 then 'not_started' when te.flooring_completed then 'completed' else 'in_progress' end),
      jsonb_build_object('key', 'cleaning_housekeeping', 'label', 'Cleaning / Housekeeping', 'status', case when te.cleaning_count = 0 then 'not_started' when te.cleaning_completed then 'completed' else 'in_progress' end),
      jsonb_build_object('key', 'appliances', 'label', 'Appliances', 'status', case when te.appliance_count = 0 then 'not_started' when te.appliance_completed then 'completed' else 'in_progress' end),
      jsonb_build_object('key', 'rent_ready', 'label', 'Rent Ready', 'status', case when te.rent_ready then 'completed' when te.move_out_date <= now() then 'in_progress' else 'not_started' end),
      jsonb_build_object('key', 'marketing_active', 'label', 'Marketing Active', 'status', case when te.ready_for_showing_on is not null then 'completed' when te.move_out_date <= now() then 'in_progress' else 'not_started' end)
    ) as milestones,
    (
      te.work_order_count > 0
      and te.all_work_orders_completed
      and coalesce(te.has_current_resident, false)
    ) as strict_completed
  from turn_evidence te
),
final_rows as (
  select
    mr.*,
    (select count(*)::int from jsonb_array_elements(mr.milestones) milestone where milestone->>'status' = 'completed') as milestones_completed
  from milestone_rows mr
)
select
  fr.turn_key,
  fr.unit_turn_id,
  fr.unit_id,
  fr.property_id,
  fr.unit_name,
  fr.property_name,
  fr.move_out_date,
  fr.next_move_in_date,
  fr.next_resident_name,
  fr.inspection_date,
  fr.turn_end_date,
  case
    when fr.strict_completed then 'Completed'
    when fr.move_out_date > now() then 'Upcoming'
    else 'In Progress'
  end as status,
  fr.unit_turn_status,
  fr.milestones,
  fr.milestones_completed,
  case when fr.strict_completed then 100 else least(99, round((fr.milestones_completed::numeric / 11) * 100)::int) end as completion_percent,
  fr.work_order_count,
  fr.swept_work_orders,
  fr.in_house_cost,
  fr.third_party_cost,
  fr.all_work_orders_completed,
  coalesce(fr.has_current_resident, false) as has_current_resident,
  fr.strict_completed,
  jsonb_build_object(
    'is_native_turn', fr.is_native_turn,
    'rogue_wos_detected', fr.rogue_wos_detected
  ) as compliance_status
from final_rows fr
where fr.move_out_date >= current_date - ($1::int * interval '1 day')
  and ($3::text is null or exists (
    select 1
    from appfolio_properties p_scope
    where p_scope.id = fr.property_id
      and p_scope.property_group_id = $3::text
  ))
  and ($4::text = '' or lower(case
    when fr.strict_completed then 'Completed'
    when fr.move_out_date > now() then 'Upcoming'
    else 'In Progress'
  end) like '%' || lower($4::text) || '%')
order by fr.strict_completed asc, fr.move_out_date desc
limit $2::int
`;
