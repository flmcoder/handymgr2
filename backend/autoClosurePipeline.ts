/**
 * Work Order Auto-Closure Pipeline — Stage 1: Confidence Matching.
 *
 * Orchestrates a bulk aged-work-order closure operation:
 *  1. Delegates to closureMatchingPolicy.runConfidenceMatch() which queries
 *     local appfolio_work_orders + appfolio_bills caches via phased CTE.
 *  2. Inserts scored candidates into aged_wo_closure_candidates with
 *     confidence_score, confidence_tier, and match_flags.
 *  3. HALTS — requires explicit human approval before Stage 2 (mutation).
 *
 * No live AppFolio API calls — all matching runs against the local Postgres cache
 * populated by the sync pipeline.
 */

import { randomUUID } from 'node:crypto';
import { queryClient } from './db.ts';
import { runConfidenceMatch } from './closureMatchingPolicy.ts';

const DEFAULT_MIN_AGE_DAYS = 15;

export interface PipelineStage1Result {
  pipelineRunId: string;
  status: 'stage1_complete' | 'failed';
  candidatesInserted: number;
  minAgeDays: number;
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

export async function runStage1(minAgeDays?: number): Promise<PipelineStage1Result> {
  if (pipelineState.stage === 'stage1_running') {
    throw new Error('Stage 1 is already running. Wait for completion or check status.');
  }

  const pipelineRunId = randomUUID();
  const ageThreshold = minAgeDays ?? DEFAULT_MIN_AGE_DAYS;

  pipelineState.currentRunId = pipelineRunId;
  pipelineState.stage = 'stage1_running';
  pipelineState.startedAt = new Date().toISOString();
  pipelineState.error = null;

  console.log(`[autoClosure] Stage 1 START runId=${pipelineRunId} minAgeDays=${ageThreshold}`);

  try {
    const candidatesInserted = await runConfidenceMatch({
      pipelineRunId,
      minAgeDays: ageThreshold,
    });

    pipelineState.stage = 'stage1_complete';

    console.log(`[autoClosure] Stage 1 COMPLETE runId=${pipelineRunId} candidates=${candidatesInserted}`);

    return {
      pipelineRunId,
      status: 'stage1_complete',
      candidatesInserted,
      minAgeDays: ageThreshold,
    };
  } catch (err) {
    const errorText = String((err as any)?.message ?? err);

    pipelineState.stage = 'failed';
    pipelineState.error = errorText;

    console.error(`[autoClosure] Stage 1 FAILED runId=${pipelineRunId}`, errorText);

    return {
      pipelineRunId,
      status: 'failed',
      candidatesInserted: 0,
      minAgeDays: ageThreshold,
      error: errorText,
    };
  }
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
        ORDER BY confidence_score DESC NULLS LAST, created_at DESC
        LIMIT ${limit}
        OFFSET ${offset}
      `
    : await queryClient`
        SELECT *
        FROM aged_wo_closure_candidates
        WHERE pipeline_run_id = ${pipelineRunId}
        ORDER BY confidence_score DESC NULLS LAST, created_at DESC
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
