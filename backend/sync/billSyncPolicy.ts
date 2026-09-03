export type NormalizedBillSyncRow = {
  id: string;
  billNumber: string | null;
  vendorId: string | null;
  vendorName: string | null;
  propertyId: string | null;
  propertyName: string | null;
  unitId: string | null;
  status: string | null;
  statusLabel: string | null;
  billTotalAmount: number | null;
  invoiceDate: Date | null;
  dueDate: Date | null;
  paidAt: Date | null;
  updatedAt: Date | null;
};

function optionalString(value: unknown): string | null {
  const normalized = String(value ?? '').trim();
  return normalized || null;
}

function optionalNumber(value: unknown): number | null {
  const normalized = Number.parseFloat(String(value ?? '').replace(/[^0-9.-]/g, ''));
  return Number.isFinite(normalized) ? normalized : null;
}

function optionalDate(value: unknown): Date | null {
  if (!value) return null;
  const normalized = new Date(String(value));
  return Number.isNaN(normalized.getTime()) ? null : normalized;
}

function resolvePropertyId(row: Record<string, any>): string | null {
  const direct = optionalString(row.PropertyId ?? row.property_id ?? row.propertyId);
  if (direct) return direct;

  const lineItems = [row.LineItems, row.line_items].find(Array.isArray) || [];
  const propertyIds = Array.from(new Set(
    lineItems
      .flatMap((item: any) => [item?.PropertyId, item?.property_id, item?.property?.Id, item?.property?.id])
      .map(optionalString)
      .filter((value: string | null): value is string => Boolean(value)),
  ));
  return propertyIds.length === 1 ? propertyIds[0] : null;
}

export function normalizeBillSyncRow(value: unknown): NormalizedBillSyncRow | null {
  if (!value || typeof value !== 'object') return null;
  const row = value as Record<string, any>;
  const id = optionalString(row.Id ?? row.id ?? row.BillId ?? row.bill_id);
  if (!id) return null;

  return {
    id,
    billNumber: optionalString(row.BillNumber ?? row.bill_number ?? row.Reference ?? row.reference),
    vendorId: optionalString(row.VendorId ?? row.vendor_id ?? row.PayeeId ?? row.payee_id),
    vendorName: optionalString(row.VendorName ?? row.vendor_name ?? row.PayeeName ?? row.payee_name),
    propertyId: resolvePropertyId(row),
    propertyName: optionalString(row.PropertyName ?? row.property_name),
    unitId: optionalString(row.UnitId ?? row.unit_id),
    status: optionalString(row.Status ?? row.status),
    statusLabel: optionalString(row.ApprovalStatus ?? row.approval_status ?? row.StatusLabel ?? row.status_label),
    billTotalAmount: optionalNumber(row.BillTotalAmount ?? row.bill_total_amount ?? row.TotalAmount ?? row.total_amount ?? row.Amount ?? row.amount),
    invoiceDate: optionalDate(row.InvoiceDate ?? row.invoice_date ?? row.BillDate ?? row.bill_date),
    dueDate: optionalDate(row.DueDate ?? row.due_date),
    paidAt: optionalDate(row.PaidAt ?? row.paid_at ?? row.PaymentDate ?? row.payment_date),
    updatedAt: optionalDate(row.LastUpdatedAt ?? row.last_updated_at ?? row.UpdatedAt ?? row.updated_at),
  };
}