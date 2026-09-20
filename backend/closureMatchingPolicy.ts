/**
 * Multi-Variable Confidence Matching Engine for Aged Work Order Closure
 *
 * Uses a phased CTE approach to evaluate work order ↔ bill matches across
 * multiple independent variables. Each phase adds weighted match flags without
 * overwriting prior phases (via COALESCE + jsonb_build_object merge).
 *
 * Phases:
 *   1. Direct WO→Bill link (50 pts)
 *   2. Strict match: vendor_id (20) + property_id (15) + amount exact (5) or within 10% (15)
 *   3. Address cross-ref: vendor_id + property_address (10 pts)
 *   4. Fuzzy vendor name (10 pts) + multi-unit property sweep (12 pts)
 *   5. Date overlap: bill invoice/service dates within WO created→completed window (10 pts)
 *
 * Confidence tiers:
 *   very_high: 80+
 *   high:      60-79
 *   medium:    40-59
 *   low:       20-39
 *   none:      <20
 */

import { queryClient } from './db.ts';

export interface MatchingConfig {
  pipelineRunId: string;
  minAgeDays?: number;
  amountTolerancePct?: number;
}

const DEFAULT_CONFIG = {
  minAgeDays: 15,
  amountTolerancePct: 0.10,
};

export async function runConfidenceMatch(config: MatchingConfig): Promise<number> {
  const { pipelineRunId } = config;
  const minAgeDays = config.minAgeDays ?? DEFAULT_CONFIG.minAgeDays;
  const tolerance = config.amountTolerancePct ?? DEFAULT_CONFIG.amountTolerancePct;

  await queryClient`
    DELETE FROM aged_wo_closure_candidates
    WHERE pipeline_run_id = ${pipelineRunId}
  `;

  await queryClient`
    WITH aged_wos AS (
      SELECT
        wo.id AS wo_id,
        wo.work_order_uuid,
        wo.wo_number,
        wo.vendor_id,
        wo.vendor_name,
        wo.property_id,
        wo.unit_id,
        wo.status AS wo_status,
        wo.total_cost AS wo_total_cost,
        wo.created_at AS wo_created_at,
        COALESCE(wo.completed_on, wo.work_completed_on) AS wo_completed_on,
        wp.name AS property_name,
        CONCAT_WS(', ',
          NULLIF(wp.street, ''),
          NULLIF(wp.city, ''),
          NULLIF(wp.state, ''),
          NULLIF(wp.zip, '')
        ) AS property_address
      FROM appfolio_work_orders wo
      LEFT JOIN appfolio_properties wp ON wp.id = wo.property_id
      WHERE wo.created_at < NOW() - make_interval(days => ${minAgeDays})
        AND LOWER(COALESCE(wo.status, '')) NOT IN ('completed', 'closed', 'canceled', 'cancelled')
        AND wo.vendor_id IS NOT NULL
        AND wo.property_id IS NOT NULL
    ),

    recent_bills AS (
      SELECT
        b.id AS bill_id,
        b.bill_number,
        b.vendor_id,
        b.vendor_name,
        b.property_id,
        b.unit_id,
        b.bill_total_amount,
        b.invoice_date AS bill_invoice_date,
        b.paid_at AS bill_paid_at,
        b.work_order_id AS bill_work_order_id,
        (b.raw_json->>'service_from')::date AS bill_service_from,
        (b.raw_json->>'service_to')::date AS bill_service_to,
        b.property_name,
        CONCAT_WS(', ',
          NULLIF(bp.street, ''),
          NULLIF(bp.city, ''),
          NULLIF(bp.state, ''),
          NULLIF(bp.zip, '')
        ) AS bill_property_address
      FROM appfolio_bills b
      LEFT JOIN appfolio_properties bp ON bp.id = b.property_id
      WHERE b.invoice_date > NOW() - INTERVAL '180 days'
        AND LOWER(COALESCE(b.status, '')) LIKE '%paid%'
    ),

    candidate_pairs AS (
      SELECT DISTINCT wo.wo_id, b.bill_id
      FROM aged_wos wo
      JOIN recent_bills b
        ON wo.vendor_id = b.vendor_id
        OR wo.property_id = b.property_id
        OR NULLIF(b.bill_work_order_id, '') = wo.wo_id
    ),

    phase1 AS (
      SELECT
        cp.wo_id,
        cp.bill_id,
        CASE WHEN NULLIF(b.bill_work_order_id, '') = wo.wo_id
          THEN true ELSE false
        END AS match_direct_link
      FROM candidate_pairs cp
      JOIN aged_wos wo ON wo.wo_id = cp.wo_id
      JOIN recent_bills b ON b.bill_id = cp.bill_id
    ),

    phase2 AS (
      SELECT
        p1.wo_id,
        p1.bill_id,
        p1.match_direct_link,
        CASE WHEN wo.vendor_id = b.vendor_id
          THEN true ELSE false
        END AS match_vendor_id,
        CASE WHEN wo.property_id = b.property_id
          THEN true ELSE false
        END AS match_property_id,
        CASE
          WHEN wo.total_cost > 0
               AND ABS(wo.total_cost - b.bill_total_amount) = 0
          THEN true ELSE false
        END AS match_amount_exact,
        CASE
          WHEN wo.total_cost > 0
               AND ABS(wo.total_cost - b.bill_total_amount) > 0
               AND ABS(wo.total_cost - b.bill_total_amount) <= (wo.total_cost * ${tolerance})
          THEN true ELSE false
        END AS match_amount_tolerance
      FROM phase1 p1
      JOIN aged_wos wo ON wo.wo_id = p1.wo_id
      JOIN recent_bills b ON b.bill_id = p1.bill_id
    ),

    phase3 AS (
      SELECT
        p2.*,
        CASE
          WHEN wo.vendor_id = b.vendor_id
               AND wo.property_address IS NOT NULL
               AND b.bill_property_address IS NOT NULL
               AND wo.property_address = b.bill_property_address
          THEN true ELSE false
        END AS match_property_address
      FROM phase2 p2
      JOIN aged_wos wo ON wo.wo_id = p2.wo_id
      JOIN recent_bills b ON b.bill_id = p2.bill_id
    ),

    phase4 AS (
      SELECT
        p3.*,
        CASE
          WHEN REPLACE(REPLACE(LOWER(COALESCE(wo.vendor_name, '')), ' ', ''), '.', '')
               = REPLACE(REPLACE(LOWER(COALESCE(b.vendor_name, '')), ' ', ''), '.', '')
               AND wo.vendor_name IS NOT NULL
               AND b.vendor_name IS NOT NULL
          THEN true ELSE false
        END AS match_vendor_name_fuzzy,
        CASE
          WHEN wo.property_id = b.property_id
               AND (b.unit_id IS NULL OR b.unit_id != wo.unit_id)
          THEN true ELSE false
        END AS match_multi_unit
      FROM phase3 p3
      JOIN aged_wos wo ON wo.wo_id = p3.wo_id
      JOIN recent_bills b ON b.bill_id = p3.bill_id
    ),

    phase5 AS (
      SELECT
        p4.*,
        CASE
          WHEN wo.wo_created_at IS NOT NULL
               AND (
                 (b.bill_invoice_date >= wo.wo_created_at::date
                  AND b.bill_invoice_date <= COALESCE(wo.wo_completed_on, CURRENT_DATE)::date)
                 OR
                 (b.bill_service_from IS NOT NULL
                  AND b.bill_service_from >= wo.wo_created_at::date
                  AND b.bill_service_from <= COALESCE(wo.wo_completed_on, CURRENT_DATE)::date)
                 OR
                 (b.bill_service_to IS NOT NULL
                  AND b.bill_service_to >= wo.wo_created_at::date
                  AND b.bill_service_to <= COALESCE(wo.wo_completed_on, CURRENT_DATE)::date)
               )
          THEN true ELSE false
        END AS match_date_overlap
      FROM phase4 p4
      JOIN aged_wos wo ON wo.wo_id = p4.wo_id
      JOIN recent_bills b ON b.bill_id = p4.bill_id
    ),

    aggregated AS (
      SELECT
        p5.wo_id,
        p5.bill_id,
        p5.match_direct_link,
        p5.match_vendor_id,
        p5.match_property_id,
        p5.match_amount_exact,
        p5.match_amount_tolerance,
        p5.match_property_address,
        p5.match_vendor_name_fuzzy,
        p5.match_multi_unit,
        p5.match_date_overlap,
        jsonb_build_object(
          'direct_link',       p5.match_direct_link,
          'vendor_id',         p5.match_vendor_id,
          'property_id',       p5.match_property_id,
          'amount_exact',      p5.match_amount_exact,
          'amount_tolerance',  p5.match_amount_tolerance,
          'property_address',  p5.match_property_address,
          'vendor_name_fuzzy', p5.match_vendor_name_fuzzy,
          'multi_unit',        p5.match_multi_unit,
          'date_overlap',      p5.match_date_overlap
        ) AS match_flags
      FROM phase5 p5
    ),

    scored AS (
      SELECT
        a.*,
        (CASE WHEN a.match_direct_link      THEN 50 ELSE 0 END)
        + (CASE WHEN a.match_vendor_id       THEN 20 ELSE 0 END)
        + (CASE WHEN a.match_property_id     THEN 15 ELSE 0 END)
        + (CASE WHEN a.match_amount_exact    THEN  5 ELSE 0 END)
        + (CASE WHEN a.match_amount_tolerance THEN 15 ELSE 0 END)
        + (CASE WHEN a.match_property_address THEN 10 ELSE 0 END)
        + (CASE WHEN a.match_vendor_name_fuzzy THEN 10 ELSE 0 END)
        + (CASE WHEN a.match_multi_unit      THEN 12 ELSE 0 END)
        + (CASE WHEN a.match_date_overlap    THEN 10 ELSE 0 END)
        AS confidence_score
      FROM aggregated a
    ),

    tiered AS (
      SELECT
        s.*,
        CASE
          WHEN s.confidence_score >= 80 THEN 'very_high'
          WHEN s.confidence_score >= 60 THEN 'high'
          WHEN s.confidence_score >= 40 THEN 'medium'
          WHEN s.confidence_score >= 20 THEN 'low'
          ELSE 'none'
        END AS confidence_tier
      FROM scored s
    )

    INSERT INTO aged_wo_closure_candidates (
      id, work_order_id, wo_number, work_order_uuid,
      bill_id, bill_number,
      vendor_id, vendor_name,
      property_id, property_name, unit_id,
      wo_status, wo_total_cost, bill_total_amount, amount_delta,
      wo_created_at, wo_completed_on,
      bill_invoice_date, bill_paid_at, bill_service_from, bill_service_to,
      property_address,
      match_flags, confidence_score, confidence_tier,
      status, pipeline_run_id, match_reason
    )
    SELECT
      wo.wo_id || ':' || b.bill_id,
      wo.wo_id,
      wo.wo_number,
      wo.work_order_uuid,
      b.bill_id,
      b.bill_number,
      wo.vendor_id,
      wo.vendor_name,
      wo.property_id,
      wo.property_name,
      wo.unit_id,
      wo.wo_status,
      wo.wo_total_cost,
      b.bill_total_amount,
      ABS(COALESCE(wo.wo_total_cost, 0) - COALESCE(b.bill_total_amount, 0)),
      wo.wo_created_at,
      wo.wo_completed_on,
      b.bill_invoice_date,
      b.bill_paid_at,
      b.bill_service_from,
      b.bill_service_to,
      wo.property_address,
      t.match_flags,
      t.confidence_score,
      t.confidence_tier,
      'pending_review',
      ${pipelineRunId},
      'multi_variable_confidence_match'
    FROM tiered t
    JOIN aged_wos wo ON wo.wo_id = t.wo_id
    JOIN recent_bills b ON b.bill_id = t.bill_id
    WHERE t.confidence_score > 0
    ON CONFLICT (id) DO UPDATE SET
      amount_delta = EXCLUDED.amount_delta,
      match_flags = EXCLUDED.match_flags,
      confidence_score = EXCLUDED.confidence_score,
      confidence_tier = EXCLUDED.confidence_tier,
      updated_at = NOW()
  `;

  const countResult = await queryClient`
    SELECT COUNT(*)::int AS cnt
    FROM aged_wo_closure_candidates
    WHERE pipeline_run_id = ${pipelineRunId}
  `;

  return countResult[0]?.cnt ?? 0;
}
