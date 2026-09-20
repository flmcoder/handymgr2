/**
 * Work Order Auto-Closure Pipeline — Stage 1: Data Ingestion & SQL Matching.
 *
 * Orchestrates a bulk aged-work-order closure operation:
 *  1. Fetches aged open work orders via AppFolio GET /api/v0/work_orders
 *  2. Fetches recent paid bills via AppFolio GET /api/v0/bills
 *  3. Runs a strict 3-point SQL match in PostgreSQL (vendor + property + amount)
 *  4. Inserts matches into aged_wo_closure_candidates with status 'pending_review'
 *  5. HALTS — requires explicit human approval before Stage 2 (mutation).
 *
 * Rate-limit aware: uses the shared afFetch worker which enforces sliding windows
 * and Retry-After backoff. Paginates with page[size]=1000 to minimize API calls.
 */

import { randomUUID } from 'node:crypto';
import { queryClient } from './db.ts';
import { afFetch, AfFetchError } from './sync/fetchWorker.ts';
import { afBaseUrl } from './sync/afCredentials.ts';

const PAGE_SIZE = 1000;
const WO_AGE_DAYS = 90;
const BILL_RECENT_DAYS = 180;
const AMOUNT_TOLERANCE_PCT = 0.10;

export interface PipelineStage1Result {
  pipelineRunId: string;
  status: 'stage1_complete' | 'failed';
  workOrdersFetched: number;
  billsFetched: number;
  candidatesInserted: number;
  error?: string;
}

export interface PipelineStatus {
  pipelineRunId: string | null;
  stage: 'idle' | 'stage1_running' | 'stage1_complete' | 'stage2_approved' | 'failed';
  startedAt: string | null;
  candidatesCount: number;
  pendingReviewCount: number;
  approvedCount: number;
  closedCount: number;
  error: string | null;
}

const pipelineState: {
  currentRunId: string | null;
  stage: PipelineStatus['stage'];
  startedAt: string | null;
  error: string | null;
} = {
  currentRunId: null,
  stage: 'idle',
  startedAt: null,
  error: null,
};

export function getPipelineStatus(): PipelineStatus {
  return {
    pipelineRunId: pipelineState.currentRunId,
    stage: pipelineState.stage,
    startedAt: pipelineState.startedAt,
    candidatesCount: 0,
    pendingReviewCount: 0,
    approvedCount: 0,
    closedCount: 0,
    error: pipelineState.error,
  };
}

export async function getPipelineStatusWithCounts(): Promise<PipelineStatus> {
  const status = getPipelineStatus();

  if (!pipelineState.currentRunId) return status;

  try {
    const counts = await queryClient`
      SELECT
        COUNT(*)::int AS total,
        COUNT(*) FILTER (WHERE status = 'pending_review')::int AS pending_review,
        COUNT(*) FILTER (WHERE status = 'approved')::int AS approved,
        COUNT(*) FILTER (WHERE status = 'closed')::int AS closed
      FROM aged_wo_closure_candidates
      WHERE pipeline_run_id = ${pipelineState.currentRunId}
    `;

    status.candidatesCount = counts[0]?.total ?? 0;
    status.pendingReviewCount = counts[0]?.pending_review ?? 0;
    status.approvedCount = counts[0]?.approved ?? 0;
    status.closedCount = counts[0]?.closed ?? 0;
  } catch {
    // counts are best-effort
  }

  return status;
}

export async function runStage1(): Promise<PipelineStage1Result> {
  if (pipelineState.stage === 'stage1_running') {
    throw new Error('Stage 1 is already running. Wait for completion or check status.');
  }

  const pipelineRunId = randomUUID();
  pipelineState.currentRunId = pipelineRunId;
  pipelineState.stage = 'stage1_running';
  pipelineState.startedAt = new Date().toISOString();
  pipelineState.error = null;

  console.log(`[autoClosure] Stage 1 START runId=${pipelineRunId}`);

  try {
    await clearPreviousCandidates(pipelineRunId);

    const woCount = await fetchAndStageWorkOrders(pipelineRunId);
    const billCount = await fetchAndStageBills(pipelineRunId);

    const matched = await runStrictMatchQuery(pipelineRunId);

    pipelineState.stage = 'stage1_complete';

    console.log(`[autoClosure] Stage 1 COMPLETE runId=${pipelineRunId} wo=${woCount} bills=${billCount} candidates=${matched}`);

    return {
      pipelineRunId,
      status: 'stage1_complete',
      workOrdersFetched: woCount,
      billsFetched: billCount,
      candidatesInserted: matched,
    };
  } catch (err) {
    const errorText = err instanceof AfFetchError
      ? `${err.kind}: ${err.message}`
      : String((err as any)?.message ?? err);

    pipelineState.stage = 'failed';
    pipelineState.error = errorText;

    console.error(`[autoClosure] Stage 1 FAILED runId=${pipelineRunId}`, errorText);

    return {
      pipelineRunId,
      status: 'failed',
      workOrdersFetched: 0,
      billsFetched: 0,
      candidatesInserted: 0,
      error: errorText,
    };
  }
}

async function clearPreviousCandidates(pipelineRunId: string): Promise<void> {
  await queryClient`
    DELETE FROM aged_wo_closure_candidates
    WHERE pipeline_run_id = ${pipelineRunId}
  `;
}

async function fetchAndStageWorkOrders(pipelineRunId: string): Promise<number> {
  const baseUrl = afBaseUrl('v0');
  const fromDate = new Date(Date.now() - WO_AGE_DAYS * 86400_000).toISOString().slice(0, 19) + 'Z';

  let url: string | null =
    `${baseUrl}/api/v0/work_orders` +
    `?filters%5BLastUpdatedAtFrom%5D=${encodeURIComponent(fromDate)}` +
    `&page%5Bsize%5D=${PAGE_SIZE}`;

  let totalFetched = 0;

  while (url) {
    const result = await afFetch(url, {
      endpointKey: 'auto_closure:work_orders',
      apiVersion: 'v0',
      runId: pipelineRunId,
    });

    const rows = result.data?.data ?? result.data?.results ?? [];
    if (!Array.isArray(rows) || rows.length === 0) break;

    await stageWorkOrderRows(rows, pipelineRunId);
    totalFetched += rows.length;

    if (result.cursorOut) {
      url = `${baseUrl}${result.cursorOut}`;
    } else {
      url = null;
    }
  }

  console.log(`[autoClosure] Fetched ${totalFetched} aged work orders`);
  return totalFetched;
}

async function stageWorkOrderRows(rows: any[], pipelineRunId: string): Promise<void> {
  for (const row of rows) {
    const id = String(row.Id ?? row.id ?? row.work_order_id ?? '').trim();
    if (!id) continue;

    const status = String(row.Status ?? row.status ?? '').trim().toLowerCase();
    if (status.includes('closed') || status.includes('completed') || status.includes('cancel')) continue;

    const vendorId = String(row.VendorId ?? row.vendor_id ?? '').trim() || null;
    const propertyId = String(row.PropertyId ?? row.property_id ?? '').trim() || null;

    if (!vendorId || !propertyId) continue;

    const totalCost = parseFloat(String(row.TotalCost ?? row.total_cost ?? '0').replace(/[^0-9.-]/g, ''));
    if (!Number.isFinite(totalCost) || totalCost <= 0) continue;

    await queryClient`
      INSERT INTO aged_wo_closure_candidates
        (id, work_order_id, wo_number, work_order_uuid, bill_id, vendor_id, vendor_name,
         property_id, wo_status, wo_total_cost, status, pipeline_run_id, match_reason)
      VALUES (
        ${id + ':_wo'},
        ${id},
        ${String(row.WorkOrderNumber ?? row.work_order_number ?? row.Number ?? '').trim() || null},
        ${String(row.UUID ?? row.uuid ?? row.work_order_uuid ?? '').trim() || null},
        ${''},
        ${vendorId},
        ${String(row.VendorName ?? row.vendor_name ?? '').trim() || null},
        ${propertyId},
        ${status},
        ${totalCost},
        'wo_staged',
        ${pipelineRunId},
        'work_order_staged'
      )
      ON CONFLICT (id) DO UPDATE SET
        wo_status = EXCLUDED.wo_status,
        wo_total_cost = EXCLUDED.wo_total_cost,
        updated_at = NOW()
    `;
  }
}

async function fetchAndStageBills(pipelineRunId: string): Promise<number> {
  const baseUrl = afBaseUrl('v0');
  const fromDate = new Date(Date.now() - BILL_RECENT_DAYS * 86400_000).toISOString().slice(0, 19) + 'Z';

  let url: string | null =
    `${baseUrl}/api/v0/bills` +
    `?filters%5BLastUpdatedAtFrom%5D=${encodeURIComponent(fromDate)}` +
    `&page%5Bsize%5D=${PAGE_SIZE}`;

  let totalFetched = 0;

  while (url) {
    const result = await afFetch(url, {
      endpointKey: 'auto_closure:bills',
      apiVersion: 'v0',
      runId: pipelineRunId,
    });

    const rows = result.data?.data ?? result.data?.results ?? [];
    if (!Array.isArray(rows) || rows.length === 0) break;

    await stageBillRows(rows, pipelineRunId);
    totalFetched += rows.length;

    if (result.cursorOut) {
      url = `${baseUrl}${result.cursorOut}`;
    } else {
      url = null;
    }
  }

  console.log(`[autoClosure] Fetched ${totalFetched} recent bills`);
  return totalFetched;
}

async function stageBillRows(rows: any[], pipelineRunId: string): Promise<void> {
  for (const row of rows) {
    const id = String(row.Id ?? row.id ?? row.BillId ?? row.bill_id ?? '').trim();
    if (!id) continue;

    const status = String(row.Status ?? row.status ?? '').trim().toLowerCase();
    if (!status.includes('paid')) continue;

    const vendorId = String(row.VendorId ?? row.vendor_id ?? row.PayeeId ?? row.payee_id ?? '').trim() || null;
    const propertyId = String(row.PropertyId ?? row.property_id ?? '').trim() || null;

    if (!vendorId || !propertyId) continue;

    const billAmount = parseFloat(String(row.BillTotalAmount ?? row.bill_total_amount ?? row.TotalAmount ?? row.total_amount ?? '0').replace(/[^0-9.-]/g, ''));
    if (!Number.isFinite(billAmount) || billAmount <= 0) continue;

    // Extract work_order_id for direct WO-to-bill linking (Phase 5 enhancement)
    const workOrderId = String(row.WorkOrderId ?? row.work_order_id ?? '').trim() || null;

    await queryClient`
      INSERT INTO aged_wo_closure_candidates
        (id, work_order_id, bill_id, bill_number, vendor_id, vendor_name,
         property_id, property_name, unit_id, bill_total_amount, status, pipeline_run_id, match_reason)
      VALUES (
        ${id + ':_bill'},
        ${workOrderId ?? ''},
        ${id},
        ${String(row.BillNumber ?? row.bill_number ?? row.Reference ?? row.reference ?? '').trim() || null},
        ${vendorId},
        ${String(row.VendorName ?? row.vendor_name ?? row.PayeeName ?? row.payee_name ?? '').trim() || null},
        ${propertyId},
        ${String(row.PropertyName ?? row.property_name ?? '').trim() || null},
        ${String(row.UnitId ?? row.unit_id ?? '').trim() || null},
        ${billAmount},
        'bill_staged',
        ${pipelineRunId},
        'bill_staged'
      )
      ON CONFLICT (id) DO UPDATE SET
        work_order_id = EXCLUDED.work_order_id,
        bill_total_amount = EXCLUDED.bill_total_amount,
        bill_number = EXCLUDED.bill_number,
        updated_at = NOW()
    `;
  }
}

async function runStrictMatchQuery(pipelineRunId: string): Promise<number> {
  // Phase 5: Use direct work_order_id link from bills for highest-confidence matches
  await queryClient`
    WITH staged_wos AS (
      SELECT *
      FROM aged_wo_closure_candidates
      WHERE pipeline_run_id = ${pipelineRunId}
        AND status = 'wo_staged'
    ),
    staged_bills AS (
      SELECT *
      FROM aged_wo_closure_candidates
      WHERE pipeline_run_id = ${pipelineRunId}
        AND status = 'bill_staged'
    ),
    direct_matches AS (
      SELECT
        wo.id AS wo_row_id,
        bill.id AS bill_row_id,
        wo.work_order_id,
        wo.wo_number,
        wo.work_order_uuid,
        bill.bill_id,
        bill.bill_number,
        wo.vendor_id,
        wo.vendor_name,
        wo.property_id,
        bill.property_name,
        bill.unit_id,
        wo.wo_status,
        wo.wo_total_cost,
        bill.bill_total_amount,
        ABS(wo.wo_total_cost - bill.bill_total_amount) AS amount_delta,
        1.0 AS match_score
      FROM staged_wos wo
      JOIN staged_bills bill
        ON bill.work_order_id = wo.work_order_id
       AND bill.work_order_id IS NOT NULL
       AND bill.work_order_id <> ''
    )
    INSERT INTO aged_wo_closure_candidates
      (id, work_order_id, wo_number, work_order_uuid, bill_id, bill_number,
       vendor_id, vendor_name, property_id, property_name, unit_id,
       wo_status, wo_total_cost, bill_total_amount, amount_delta, match_score,
       status, pipeline_run_id, match_reason)
    SELECT
      work_order_id || ':' || bill_id,
      work_order_id,
      wo_number,
      work_order_uuid,
      bill_id,
      bill_number,
      vendor_id,
      vendor_name,
      property_id,
      property_name,
      unit_id,
      wo_status,
      wo_total_cost,
      bill_total_amount,
      amount_delta,
      match_score,
      'pending_review',
      ${pipelineRunId},
      'direct_link: bill.work_order_id = work_order.id'
    FROM direct_matches
    ON CONFLICT (id) DO UPDATE SET
      match_score = EXCLUDED.match_score,
      amount_delta = EXCLUDED.amount_delta,
      match_reason = EXCLUDED.match_reason,
      status = EXCLUDED.status,
      updated_at = NOW()
  `;

  // Phase 5: Fallback to 3-point match for bills without direct work_order_id links
  await queryClient`
    WITH staged_wos AS (
      SELECT *
      FROM aged_wo_closure_candidates
      WHERE pipeline_run_id = ${pipelineRunId}
        AND status = 'wo_staged'
    ),
    staged_bills AS (
      SELECT *
      FROM aged_wo_closure_candidates
      WHERE pipeline_run_id = ${pipelineRunId}
        AND status = 'bill_staged'
        AND (work_order_id IS NULL OR work_order_id = '')
    ),
    strict_matches AS (
      SELECT
        wo.id AS wo_row_id,
        bill.id AS bill_row_id,
        wo.work_order_id,
        wo.wo_number,
        wo.work_order_uuid,
        bill.bill_id,
        bill.bill_number,
        wo.vendor_id,
        wo.vendor_name,
        wo.property_id,
        bill.property_name,
        bill.unit_id,
        wo.wo_status,
        wo.wo_total_cost,
        bill.bill_total_amount,
        ABS(wo.wo_total_cost - bill.bill_total_amount) AS amount_delta,
        CASE
          WHEN wo.wo_total_cost > 0
          THEN 1.0 - LEAST(1.0, ABS(wo.wo_total_cost - bill.bill_total_amount) / wo.wo_total_cost)
          ELSE 0.0
        END AS match_score
      FROM staged_wos wo
      JOIN staged_bills bill
        ON bill.vendor_id = wo.vendor_id
       AND bill.property_id = wo.property_id
      WHERE ABS(wo.wo_total_cost - bill.bill_total_amount) <= (wo.wo_total_cost * ${AMOUNT_TOLERANCE_PCT})
    )
    INSERT INTO aged_wo_closure_candidates
      (id, work_order_id, wo_number, work_order_uuid, bill_id, bill_number,
       vendor_id, vendor_name, property_id, property_name, unit_id,
       wo_status, wo_total_cost, bill_total_amount, amount_delta, match_score,
       status, pipeline_run_id, match_reason)
    SELECT
      work_order_id || ':' || bill_id,
      work_order_id,
      wo_number,
      work_order_uuid,
      bill_id,
      bill_number,
      vendor_id,
      vendor_name,
      property_id,
      property_name,
      unit_id,
      wo_status,
      wo_total_cost,
      bill_total_amount,
      amount_delta,
      match_score,
      'pending_review',
      ${pipelineRunId},
      'strict_3pt: vendor + property + amount within ' || (${AMOUNT_TOLERANCE_PCT} * 100)::int || '%'
    FROM strict_matches
    ON CONFLICT (id) DO UPDATE SET
      match_score = EXCLUDED.match_score,
      amount_delta = EXCLUDED.amount_delta,
      match_reason = EXCLUDED.match_reason,
      status = EXCLUDED.status,
      updated_at = NOW()
  `;

  const countResult = await queryClient`
    SELECT COUNT(*)::int AS cnt
    FROM aged_wo_closure_candidates
    WHERE pipeline_run_id = ${pipelineRunId}
      AND status = 'pending_review'
  `;

  await queryClient`
    DELETE FROM aged_wo_closure_candidates
    WHERE pipeline_run_id = ${pipelineRunId}
      AND status IN ('wo_staged', 'bill_staged')
  `;

  return countResult[0]?.cnt ?? 0;
}

export async function approveCandidates(
  pipelineRunId: string,
  candidateIds: string[],
  reviewedBy: string,
): Promise<{ approved: number }> {
  if (pipelineState.stage !== 'stage1_complete') {
    throw new Error('Cannot approve: Stage 1 is not complete. Run Stage 1 first.');
  }

  if (candidateIds.length === 0) {
    return { approved: 0 };
  }

  await queryClient`
    UPDATE aged_wo_closure_candidates
    SET status = 'approved',
        reviewed_by = ${reviewedBy},
        reviewed_at = NOW(),
        updated_at = NOW()
    WHERE pipeline_run_id = ${pipelineRunId}
      AND id = ANY(${candidateIds}::text[])
      AND status = 'pending_review'
  `;

  pipelineState.stage = 'stage2_approved';

  const countResult = await queryClient`
    SELECT COUNT(*)::int AS cnt
    FROM aged_wo_closure_candidates
    WHERE pipeline_run_id = ${pipelineRunId}
      AND status = 'approved'
  `;

  return { approved: countResult[0]?.cnt ?? 0 };
}

export async function rejectCandidates(
  pipelineRunId: string,
  candidateIds: string[],
  reviewedBy: string,
  notes: string,
): Promise<{ rejected: number }> {
  if (candidateIds.length === 0) return { rejected: 0 };

  await queryClient`
    UPDATE aged_wo_closure_candidates
    SET status = 'rejected',
        reviewed_by = ${reviewedBy},
        reviewed_at = NOW(),
        review_notes = ${notes},
        updated_at = NOW()
    WHERE pipeline_run_id = ${pipelineRunId}
      AND id = ANY(${candidateIds}::text[])
      AND status = 'pending_review'
  `;

  const countResult = await queryClient`
    SELECT COUNT(*)::int AS cnt
    FROM aged_wo_closure_candidates
    WHERE pipeline_run_id = ${pipelineRunId}
      AND status = 'rejected'
  `;

  return { rejected: countResult[0]?.cnt ?? 0 };
}

export async function getCandidates(
  pipelineRunId: string,
  opts: { status?: string; limit?: number; offset?: number } = {},
): Promise<{ candidates: any[]; total: number }> {
  const { status, limit = 100, offset = 0 } = opts;

  const rows = status
    ? await queryClient`
        SELECT *
        FROM aged_wo_closure_candidates
        WHERE pipeline_run_id = ${pipelineRunId}
          AND status = ${status}
        ORDER BY match_score DESC NULLS LAST, created_at DESC
        LIMIT ${limit}
        OFFSET ${offset}
      `
    : await queryClient`
        SELECT *
        FROM aged_wo_closure_candidates
        WHERE pipeline_run_id = ${pipelineRunId}
        ORDER BY match_score DESC NULLS LAST, created_at DESC
        LIMIT ${limit}
        OFFSET ${offset}
      `;

  const countRows = status
    ? await queryClient`
        SELECT COUNT(*)::int AS cnt
        FROM aged_wo_closure_candidates
        WHERE pipeline_run_id = ${pipelineRunId}
          AND status = ${status}
      `
    : await queryClient`
        SELECT COUNT(*)::int AS cnt
        FROM aged_wo_closure_candidates
        WHERE pipeline_run_id = ${pipelineRunId}
      `;

  return { candidates: rows, total: countRows[0]?.cnt ?? 0 };
}

export async function updateSingleCandidate(
  candidateId: string,
  action: 'approve' | 'reject',
  reviewedBy: string,
  notes?: string,
): Promise<{ id: string; work_order_id: string; status: string } | null> {
  const statusMap = { approve: 'approved', reject: 'rejected' };
  const newStatus = statusMap[action];

  const rows = await queryClient`
    UPDATE aged_wo_closure_candidates
    SET status = ${newStatus},
        reviewed_by = ${reviewedBy},
        reviewed_at = NOW(),
        review_notes = ${notes ?? null},
        updated_at = NOW()
    WHERE id = ${candidateId}
      AND status = 'pending_review'
    RETURNING id, work_order_id, status
  `;

  const row = rows[0] as { id: string; work_order_id: string; status: string } | undefined;
  return row ?? null;
}

export async function resetPipeline(): Promise<void> {
  if (pipelineState.currentRunId) {
    await queryClient`
      DELETE FROM aged_wo_closure_candidates
      WHERE pipeline_run_id = ${pipelineState.currentRunId}
    `;
  }

  pipelineState.currentRunId = null;
  pipelineState.stage = 'idle';
  pipelineState.startedAt = null;
  pipelineState.error = null;
}
