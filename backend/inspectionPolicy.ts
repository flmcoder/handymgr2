export type InspectionPage = {
  limit: number;
  offset: number;
  page: number;
};

function positiveInteger(value: unknown, fallback: number): number {
  const parsed = Number.parseInt(String(value ?? ''), 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

export function parseInspectionPage(query: Record<string, unknown>): InspectionPage {
  const limit = Math.min(100, positiveInteger(query.limit, 50));
  const requestedPage = positiveInteger(query.page, 1);
  const requestedOffset = Number.parseInt(String(query.offset ?? ''), 10);
  const offset = Number.isFinite(requestedOffset) && requestedOffset >= 0
    ? Math.min(250_000, requestedOffset)
    : Math.min(250_000, (requestedPage - 1) * limit);
  return { limit, offset, page: Math.floor(offset / limit) + 1 };
}

export function inspectionMissingIndicators(row: {
  move_in_date?: string | null;
  move_out_date?: string | null;
  last_inspection_date?: string | null;
  move_out_inspection_date?: string | null;
}, today = new Date()): { missing_move_in_inspection: boolean; missing_move_out_inspection: boolean } {
  const moveIn = row.move_in_date ? new Date(row.move_in_date) : null;
  const moveOut = row.move_out_date ? new Date(row.move_out_date) : null;
  const lastInspection = row.last_inspection_date ? new Date(row.last_inspection_date) : null;
  const moveOutInspection = row.move_out_inspection_date ? new Date(row.move_out_inspection_date) : null;
  const valid = (value: Date | null): value is Date => !!value && !Number.isNaN(value.getTime());

  return {
    missing_move_in_inspection: valid(moveIn) && moveIn <= today && (!valid(lastInspection) || lastInspection < moveIn),
    missing_move_out_inspection: valid(moveOut) && moveOut >= today && !valid(moveOutInspection),
  };
}