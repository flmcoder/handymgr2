import { postgresErrorDetails } from './workOrderUpsertPolicy.ts';
import { isTransientDbError } from '../dbResiliencePolicy.ts';

export type TrackerRow = { trackingUuid: string; turnKey: string };

export type TrackerIdentityResolution = {
  /** tracking_uuid every write for this record must use (tracker row + children). */
  trackingUuid: string;
  /** Existing row to update, or null to insert. */
  existing: TrackerRow | null;
  /** True when an existing row found by tracking_uuid must adopt the new turn_key. */
  rekey: boolean;
  /** Set when the incoming tracking_uuid and turn_key point at two different rows. */
  conflictingRow: TrackerRow | null;
};

/**
 * unit_turn_tracker has two unique identities: tracking_uuid (PK) and turn_key.
 * Upserting on only one of them raises 23505 on the other when move-out identity
 * normalization changes how a turn is keyed. turn_key is the logical identity,
 * so it wins; tracking_uuid matches are only used to re-key a legacy row.
 */
export function resolveTrackerIdentity(
  incoming: TrackerRow,
  byTurnKey: TrackerRow | null,
  byTrackingUuid: TrackerRow | null,
): TrackerIdentityResolution {
  if (byTurnKey) {
    const conflictingRow = byTrackingUuid && byTrackingUuid.trackingUuid !== byTurnKey.trackingUuid
      ? byTrackingUuid
      : null;
    return { trackingUuid: byTurnKey.trackingUuid, existing: byTurnKey, rekey: false, conflictingRow };
  }
  if (byTrackingUuid) {
    return {
      trackingUuid: byTrackingUuid.trackingUuid,
      existing: byTrackingUuid,
      rekey: byTrackingUuid.turnKey !== incoming.turnKey,
      conflictingRow: null,
    };
  }
  return { trackingUuid: incoming.trackingUuid, existing: null, rekey: false, conflictingRow: null };
}

export type TurnTrackerFailure = { turnKey: string; error: unknown };

/**
 * Logs a failed tracker write with SQLSTATE/constraint metadata and the
 * identifying payload (never raw_json-style bulk data), then rethrows.
 */
export async function writeTurnTrackerWithDiagnostics<T>(
  record: Record<string, unknown>,
  write: () => Promise<T>,
): Promise<T> {
  try {
    return await write();
  } catch (error) {
    console.error('[unit_turn_tracker:upsert] Database write failed', {
      ...postgresErrorDetails(error),
      transient: isTransientDbError(error),
      message: String((error as { message?: unknown })?.message ?? error),
      turn_key: record.turn_key,
      tracking_uuid: record.tracking_uuid,
      unit_turn_id: record.unit_turn_id,
      unit_id: record.unit_id,
      property_id: record.property_id,
      status: record.status,
    });
    throw error;
  }
}

/**
 * Decides what to do with per-record failures after a batch: connection-level
 * failures abort the batch (the caller retries it), data errors are skipped,
 * and a batch where every record failed is surfaced instead of reported as success.
 */
export function summarizeTurnTrackerFailures(
  total: number,
  upserted: number,
  failures: TurnTrackerFailure[],
): { skipped: number; fatal: unknown | null } {
  const transient = failures.find((failure) => isTransientDbError(failure.error));
  if (transient) return { skipped: failures.length, fatal: transient.error };
  if (failures.length > 0 && upserted === 0 && total > 0) {
    return { skipped: failures.length, fatal: failures[0].error };
  }
  return { skipped: failures.length, fatal: null };
}
