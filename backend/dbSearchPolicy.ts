/**
 * Postgres-backed table search for the Database admin section.
 * Replaces the legacy SQLite (sql_query / sql_execute) console, which is
 * disabled. Only a small allow-list of read-only tables may be searched.
 */

export type SearchableTable = {
  key: string;
  table: string;
  // Columns rendered in the results grid, in order.
  columns: string[];
  // Columns matched with ILIKE against the search term.
  searchColumns: string[];
};

export const SEARCHABLE_TABLES: SearchableTable[] = [
  {
    key: 'work_orders',
    table: 'appfolio_work_orders',
    columns: ['wo_number', 'property_id', 'unit_id', 'status', 'priority', 'vendor_name', 'created_at'],
    searchColumns: ['wo_number::text', 'description', 'status', 'vendor_name', 'assigned_user_name'],
  },
  {
    key: 'properties',
    table: 'appfolio_properties',
    columns: ['id', 'name', 'street', 'city', 'state', 'zip'],
    searchColumns: ['name', 'street', 'city'],
  },
  {
    key: 'units',
    table: 'appfolio_units',
    columns: ['unit_id', 'property_id', 'name', 'status'],
    searchColumns: ['name', 'unit_id::text'],
  },
  {
    key: 'vendors',
    table: 'vendor_directory',
    columns: ['vendor_key', 'vendor_name', 'trade_category', 'phone_numbers', 'email'],
    searchColumns: ['vendor_name', 'trade_category', 'email', 'phone_numbers'],
  },
  {
    key: 'tech_roster',
    table: 'tech_grades',
    columns: ['tech_id', 'tech_name', 'tier', 'performance_score', 'active', 'property_group_uuid'],
    searchColumns: ['tech_name', 'tech_id'],
  },
  {
    key: 'reassignment_queue',
    table: 'reassignment_queue',
    columns: ['wo_id', 'wo_number', 'property_address', 'status', 'assigned_tech_name', 'reassignment_count'],
    searchColumns: ['wo_number', 'property_address', 'assigned_tech_name', 'wo_id'],
  },
];

const TABLE_BY_KEY = new Map(SEARCHABLE_TABLES.map((t) => [t.key, t]));

export function resolveSearchableTable(key: unknown): SearchableTable | null {
  return TABLE_BY_KEY.get(String(key || '').trim()) || null;
}

/** Clamp a requested row cap to a safe bound. */
export function clampSearchLimit(value: unknown): number {
  const n = Number.parseInt(String(value), 10);
  if (!Number.isFinite(n)) return 50;
  if (n < 1) return 1;
  return Math.min(200, n);
}

export type BuiltSearchQuery = { sql: string; params: any[] };

/**
 * Build a parameterized ILIKE search against an allow-listed table. Returns
 * null when the table key is not allow-listed so the route can 400 out.
 */
export function buildTableSearchQuery(tableKey: unknown, rawTerm: unknown, rawLimit: unknown): BuiltSearchQuery | null {
  const table = resolveSearchableTable(tableKey);
  if (!table) return null;
  const limit = clampSearchLimit(rawLimit);
  const term = String(rawTerm || '').trim();
  const params: any[] = [];
  let whereSql = '';
  if (term) {
    const like = `%${term}%`;
    const ors = table.searchColumns.map((col) => {
      params.push(like);
      return `coalesce(${col}::text, '') ilike $${params.length}`;
    });
    whereSql = `where ${ors.join(' or ')}`;
  }
  const cols = table.columns.map((c) => `${c}`).join(', ');
  params.push(limit);
  const sql = `select ${cols} from ${table.table} ${whereSql} limit $${params.length}`;
  return { sql, params };
}
