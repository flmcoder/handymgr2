// ============================================================================
// db.ts — Drizzle ORM PostgreSQL Connection (via PgBouncer tunnel)
// Replaces legacy SQLite with PostgreSQL via postgres-js
// ============================================================================

import postgres from 'postgres';
import { drizzle } from 'drizzle-orm/postgres-js';
import * as schema from './schema';

// ── Environment Variables ───────────────────────────────────────────────────
// Prefer SQL_SE (connection string) first, fallback to individual vars
const connectionString = process.env.SQL_SE;
const dbHost = process.env.DBI || process.env.PGHOST || 'localhost';
const dbPort = parseInt(process.env.DB_PORT || process.env.PGPORT || '6432', 10);
const dbName = process.env.DB_NAME || process.env.PGDATABASE || 'handymgr';
const dbUser = process.env.PGUSER || 'postgres';
const dbPassword = process.env.PGPASSWORD || process.env.SQL_SE || '';

// ── Connection String Builder ───────────────────────────────────────────────
function buildConnectionString(): string {
  if (connectionString && connectionString.startsWith('postgres')) {
    return connectionString;
  }
  
  // Build from individual components
  const userPass = dbPassword ? `${dbUser}:${dbPassword}` : dbUser;
  return `postgres://${userPass}@${dbHost}:${dbPort}/${dbName}`;
}

// ── postgres-js Client ──────────────────────────────────────────────────────
// Configure with tunnel-friendly settings (PgBouncer + Localtonet)
const finalConnectionString = buildConnectionString();

console.log('[db] Connecting to PostgreSQL via:', finalConnectionString.replace(/:[^:@]+@/, ':***@'));

export const queryClient = postgres(finalConnectionString, {
  max: 10, // Connection pool size (keep small for tunneled connections)
  idle_timeout: 20, // Close idle connections after 20s
  connect_timeout: 10, // Fail fast if tunnel is down
  prepare: false, // Disable prepared statements for PgBouncer compatibility
  onnotice: () => {}, // Suppress PostgreSQL notices
  debug: process.env.NODE_ENV === 'development', // Enable query logging in dev
});

// ── Drizzle ORM Instance ────────────────────────────────────────────────────
export const db = drizzle(queryClient, { schema });

// ── Connection Health Check ─────────────────────────────────────────────────
export async function testConnection(): Promise<boolean> {
  try {
    const result = await queryClient`SELECT 1 as health`;
    if (result && result.length > 0 && result[0].health === 1) {
      console.log('[db] ✅ PostgreSQL connection healthy');
      return true;
    }
    console.error('[db] ❌ PostgreSQL health check failed: unexpected result');
    return false;
  } catch (err: any) {
    console.error('[db] ❌ PostgreSQL connection error:', err.message);
    console.error('[db] Stack trace:', err.stack);
    return false;
  }
}

// ── Graceful Shutdown ───────────────────────────────────────────────────────
export async function closeConnection(): Promise<void> {
  try {
    await queryClient.end();
    console.log('[db] PostgreSQL connection closed');
  } catch (err: any) {
    console.error('[db] Error closing connection:', err.message);
  }
}

// Test connection on module load
testConnection().catch((err) => {
  console.error('[db] Failed to establish initial connection:', err);
});
