const UPSERT_BATCH = 50;

type SqlClient = {
  execute: (stmt: string | { sql: string; args?: unknown[] }) => Promise<unknown>;
};

function asString(value: unknown, max = 500): string | null {
  const normalized = String(value ?? "").trim();
  return normalized ? normalized.substring(0, max) : null;
}

function asNumber(value: unknown): number | null {
  const normalized = parseFloat(String(value ?? "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(normalized) ? normalized : null;
}

function isUuidLike(value: unknown): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(value ?? "").trim());
}

export async function upsertPropertyRowsToDb(
  sqlite: SqlClient,
  rows: any[],
): Promise<void> {
  const now = Date.now();
  for (let index = 0; index < rows.length; index += UPSERT_BATCH) {
    const chunk = rows.slice(index, index + UPSERT_BATCH);
    const placeholders: string[] = [];
    const values: unknown[] = [];
    for (const row of chunk) {
      const id = asString(row.Id || row.id || row.property_id);
      if (!id) continue;
      const groupIds: string[] = Array.isArray(row.PropertyGroupIds || row.property_group_ids)
        ? (row.PropertyGroupIds || row.property_group_ids)
          .map((entry: unknown) => String(entry ?? "").trim())
          .filter(Boolean)
        : [];
      placeholders.push("(?,?,?,?,?,?,?,?,?,?)");
      values.push(
        id,
        id,
        null,
        0,
        groupIds[0] ?? null,
        asString(row.Name || row.PropertyName || row.property_name),
        null,
        asString(row.Address1 || row.StreetAddress || row.address),
        asString(row.City || row.city),
        now,
      );
    }
    if (placeholders.length === 0) continue;
    try {
      await sqlite.execute({
        sql: `INSERT OR REPLACE INTO property_map
              (id, property_id, unit_id, is_unit, property_group_id,
               property_name, unit_name, address, city, cached_at)
              VALUES ${placeholders.join(",")}`,
        args: values,
      });
    } catch {
      // Non-fatal.
    }
  }
}

export async function upsertVendorRowsToDb(
  sqlite: SqlClient,
  rows: any[],
): Promise<void> {
  const now = Date.now();
  for (let index = 0; index < rows.length; index += UPSERT_BATCH) {
    const chunk = rows.slice(index, index + UPSERT_BATCH);
    const placeholders: string[] = [];
    const values: unknown[] = [];
    for (const row of chunk) {
      const id = asString(row.vendor_id || row.VendorId || row.Id || row.id);
      if (!id) continue;
      const name = asString(
        row.company_name ||
          row.CompanyName ||
          [row.first_name || row.FirstName, row.last_name || row.LastName]
            .filter(Boolean)
            .join(" ") ||
          row.name ||
          row.Name,
      ) ?? id;
      placeholders.push("(?,?,?,?,?,?,?,?)");
      values.push(
        id,
        name,
        asString(row.company_name || row.CompanyName),
        asString(row.email || row.Email),
        asString(row.phone || row.Phone || row.phone_numbers),
        asString(row.license_number || row.LicenseNumber),
        asString(row.insurance_expiry || row.InsuranceExpiry),
        now,
      );
    }
    if (placeholders.length === 0) continue;
    try {
      await sqlite.execute({
        sql: `INSERT OR REPLACE INTO vendor_map
              (id, name, company_name, email, phone, license_number, insurance_expiry, cached_at)
              VALUES ${placeholders.join(",")}`,
        args: values,
      });
    } catch {
      // Non-fatal.
    }
  }
}

export async function upsertWorkOrderRowsToDb(
  sqlite: SqlClient,
  rows: any[],
): Promise<void> {
  const now = Date.now();
  for (let index = 0; index < rows.length; index += UPSERT_BATCH) {
    const chunk = rows.slice(index, index + UPSERT_BATCH);
    const placeholders: string[] = [];
    const values: unknown[] = [];
    for (const row of chunk) {
      const rawId = asString(row.work_order_id || row.WorkOrderId || row.Id || row.id);
      const v0Uuid = asString(row.db_api_id || row.dbApiId || row.v0_uuid || row.v0_id || row.uuid || row.UUID);
      const id = isUuidLike(v0Uuid) ? v0Uuid : rawId;
      if (!id) continue;
      const workOrderNumber = asString(
        row.work_order_number || row.WorkOrderNumber || row.Number,
      );
      if (!workOrderNumber) continue;
      const propertyId = asString(row.property_id || row.PropertyId) ?? "";
      const propertyMapId = propertyId || id;
      const assignedRaw = row.assigned_users ?? row.AssignedUsers ?? [];
      placeholders.push("(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
      values.push(
        id,
        workOrderNumber,
        propertyMapId,
        propertyId || null,
        asString(row.unit_id || row.UnitId),
        asString(row.vendor_id || row.VendorId),
        asString(row.status || row.Status),
        asString(row.priority || row.Priority),
        asString(row.category || row.Category || row.work_order_type || row.WorkOrderType),
        asString(row.description || row.Description || row.subject || row.Subject, 500),
        typeof assignedRaw === "string" ? assignedRaw : JSON.stringify(assignedRaw),
        asString(row.created_date || row.CreatedDate || row.created_at || row.CreatedAt)?.slice(0, 10) ?? null,
        asString(row.completed_date || row.CompletedDate || row.closed_date)?.slice(0, 10) ?? null,
        asString(row.last_updated_at || row.LastUpdatedAt)?.slice(0, 24) ?? null,
        now,
      );
    }
    if (placeholders.length === 0) continue;
    try {
      await sqlite.execute({
        sql: `INSERT OR REPLACE INTO work_order_map
              (id, work_order_number, property_map_id, property_id, unit_id, vendor_id,
               status, priority, category, description, assigned_users_json,
               created_date, completed_date, last_updated_at, cached_at)
              VALUES ${placeholders.join(",")}`,
        args: values,
      });
    } catch {
      // Non-fatal.
    }
  }
}

export async function upsertBillingRowsToDb(
  sqlite: SqlClient,
  rows: any[],
): Promise<void> {
  const now = Date.now();

  for (let index = 0; index < rows.length; index += UPSERT_BATCH) {
    const chunk = rows.slice(index, index + UPSERT_BATCH);
    const placeholders: string[] = [];
    const values: unknown[] = [];
    for (const row of chunk) {
      const id = asString(row.Id || row.id || row.BillId || row.bill_id);
      if (!id) continue;
      const lineItems = Array.isArray(row.LineItems)
        ? row.LineItems
        : (Array.isArray(row.line_items) ? row.line_items : []);
      const firstLine = lineItems[0] || {};
      const propertyId = asString(
        row.PropertyId || row.property_id ||
          firstLine.PropertyId || firstLine.property_id ||
          firstLine.PropertyUuid || firstLine.property_uuid,
      );
      const unitId = asString(
        row.UnitId || row.unit_id ||
          firstLine.UnitId || firstLine.unit_id,
      );
      placeholders.push("(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
      values.push(
        id,
        asString(row.VendorId || row.vendor_id || row.PayeeId || row.payee_id),
        null,
        propertyId,
        unitId,
        asString(row.WorkOrderId || row.work_order_id),
        asString(row.WorkOrderNumber || row.work_order_number),
        asString(row.InvoiceDate || row.invoice_date)?.slice(0, 10) ?? null,
        asString(row.DueDate || row.due_date)?.slice(0, 10) ?? null,
        asNumber(row.TotalAmount || row.total_amount || row.Amount || row.amount),
        asString(row.CheckMemo || row.check_memo || row.Remarks || row.remarks, 500),
        asString(row.ApprovalStatus || row.approval_status || row.Status || row.status),
        asString(row.Reference || row.reference || id),
        lineItems.length > 0 ? JSON.stringify(lineItems) : null,
        asString(row.VendorName || row.vendor_name || row.PayeeName || row.payee_name),
        asString(row.PropertyName || row.property_name),
        asString(row.PropertyGroupId || row.property_group_id || row.PropertyGroupUuid || row.property_group_uuid),
        asString(row.PropertyGroup || row.property_group || row.PropertyGroupName || row.property_group_name),
        asString(row.LastUpdatedAt || row.last_updated_at)?.slice(0, 24) ?? null,
        now,
      );
    }
    if (placeholders.length === 0) continue;
    try {
      await sqlite.execute({
        sql: `INSERT OR REPLACE INTO billing_map
              (id, vendor_id, property_map_id, property_id, unit_id, work_order_id, work_order_number,
               invoice_date, due_date, total_amount, check_memo, approval_status, bill_number,
               line_items_json, vendor_name, property_name, property_group_id, property_group_name,
               last_updated_at, cached_at)
              VALUES ${placeholders.join(",")}`,
        args: values,
      });
    } catch {
      // Non-fatal.
    }
  }

  for (const row of rows) {
    const billId = asString(row.Id || row.id || row.BillId || row.bill_id);
    if (!billId) continue;
    const lineItems = Array.isArray(row.LineItems)
      ? row.LineItems
      : (Array.isArray(row.line_items) ? row.line_items : []);
    if (lineItems.length === 0) continue;
    try {
      await sqlite.execute({ sql: `DELETE FROM bill_line_items WHERE bill_id = ?`, args: [billId] });
      for (let index = 0; index < lineItems.length; index += UPSERT_BATCH) {
        const chunk = lineItems.slice(index, index + UPSERT_BATCH);
        const placeholders: string[] = [];
        const values: unknown[] = [];
        for (const lineItem of chunk) {
          placeholders.push("(?,?,?,?,?,?,?,?,?)");
          values.push(
            billId,
            asString(lineItem.UnitId || lineItem.unit_id || lineItem.UnitUuid || lineItem.unit_uuid),
            asString(lineItem.PropertyId || lineItem.property_id || lineItem.PropertyUuid || lineItem.property_uuid),
            asString(lineItem.GlAccountId || lineItem.gl_account_id || lineItem.GlAccount || lineItem.gl_account),
            asNumber(lineItem.Amount || lineItem.amount),
            asString(lineItem.Description || lineItem.description, 500),
            asString(lineItem.LineItemType || lineItem.line_item_type || lineItem.Type || lineItem.type),
            asNumber(lineItem.Quantity || lineItem.quantity),
            asNumber(lineItem.UnitPrice || lineItem.unit_price || lineItem.Rate || lineItem.rate),
          );
        }
        await sqlite.execute({
          sql: `INSERT INTO bill_line_items
                (bill_id, unit_id, property_id, gl_account_id, amount, description,
                 line_item_type, quantity, unit_price)
                VALUES ${placeholders.join(",")}`,
          args: values,
        });
      }
    } catch {
      // Non-fatal.
    }
  }
}

export async function upsertUnitsToDb(
  sqlite: SqlClient,
  rows: any[],
): Promise<void> {
  const now = new Date().toISOString();
  for (let index = 0; index < rows.length; index += UPSERT_BATCH) {
    const chunk = rows.slice(index, index + UPSERT_BATCH);
    const placeholders: string[] = [];
    const values: unknown[] = [];
    for (const row of chunk) {
      const unitId = asString(row.Id || row.id || row.UnitId || row.unit_id);
      const propertyId = asString(row.PropertyId || row.property_id);
      if (!unitId || !propertyId) continue;
      placeholders.push("(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
      values.push(
        unitId,
        propertyId,
        asString(row.Name || row.name),
        asString(row.Status || row.status),
        row.Bedrooms != null ? Number(row.Bedrooms) : (row.bedrooms != null ? Number(row.bedrooms) : null),
        asString(row.Bathrooms || row.bathrooms),
        asString(row.Address1 || row.address1),
        asString(row.City || row.city),
        asString(row.State || row.state),
        asString(row.Zip || row.zip),
        asString(row.LeasingType || row.leasing_type),
        row.RentReady != null ? (row.RentReady ? 1 : 0) : (row.rent_ready ? 1 : 0),
        asString(row.HiddenAt || row.hidden_at),
        asString(row.LastUpdatedAt || row.last_updated_at),
        now,
      );
    }
    if (placeholders.length === 0) continue;
    try {
      await sqlite.execute({
        sql: `INSERT OR REPLACE INTO units
              (unit_id, property_id, name, status, bedrooms, bathrooms,
               address1, city, state, zip, leasing_type, rent_ready,
               hidden_at, last_updated_at, cached_at)
              VALUES ${placeholders.join(",")}`,
        args: values,
      });
    } catch {
      // Non-fatal.
    }
  }
}