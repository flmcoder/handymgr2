/**
 * Apply a SQL migration file to the Postgres database.
 *
 * Usage:
 *   node scripts/apply-pg-sql.mjs <path-to-sql-file>
 *
 * Reads connection config from the same env vars the backend uses:
 *   DBI, DB_PORT, DB_NAME, DB_USER, DB_PASSWORD, DB_SSL
 *   (or a full SQL_SE postgres:// URL if set)
 *
 * Example:
 *   DBI=hmsh.localto.net DB_PORT=6432 DB_NAME=handymgr DB_USER=Administrator \
 *     DB_PASSWORD=secret node scripts/apply-pg-sql.mjs \
 *     db/migrations/2026-06-17_sync_control_plane.sql
 */

import { readFile } from 'node:fs/promises';
import process from 'node:process';
import postgres from 'postgres';

const sqlFile = process.argv[2];
if (!sqlFile) {
  console.error('Usage: node scripts/apply-pg-sql.mjs <sql-file>');
  process.exit(1);
}

// Resolve connection config (mirrors backend/db.ts logic).
const SQL_SE     = String(process.env.SQL_SE      || '').trim();
const DBI        = String(process.env.DBI         || process.env.PGHOST     || '127.0.0.1').trim();
const DB_PORT    = Number(process.env.DB_PORT      || process.env.PGPORT     || 6432);
const DB_NAME    = String(process.env.DB_NAME      || process.env.PGDATABASE || '').trim();
const DB_USER    = String(process.env.DB_USER      || process.env.PGUSER     || 'Administrator').trim();
const DB_PASSWORD = String(process.env.DB_PASSWORD || process.env.PGPASSWORD || '').trim();
const DB_SSL     = String(process.env.DB_SSL       || '').trim().toLowerCase();

let config;
if (SQL_SE && /^postgres(ql)?:\/\//i.test(SQL_SE)) {
  const u = new URL(SQL_SE);
  config = {
    host: u.hostname, port: Number(u.port) || DB_PORT,
    database: u.pathname.replace(/^\//, '') || DB_NAME,
    username: decodeURIComponent(u.username) || DB_USER,
    password: decodeURIComponent(u.password) || DB_PASSWORD,
    ssl: DB_SSL === 'true' ? 'require' : false,
  };
} else {
  if (!DB_NAME) { console.error('DB_NAME is required.'); process.exit(1); }
  config = { host: DBI, port: DB_PORT, database: DB_NAME, username: DB_USER, password: DB_PASSWORD, ssl: DB_SSL === 'true' ? 'require' : false };
}

console.log(`[apply-pg-sql] Connecting to ${config.host}:${config.port}/${config.database} as ${config.username}`);

const sql = postgres({ ...config, max: 1, connect_timeout: 15, prepare: false });

const raw = await readFile(sqlFile, 'utf8');

// Split on statement boundaries (semicolons not inside string literals).
// Simple split is safe for our migration files which contain no PL/pgSQL blocks.
const statements = raw
  .split(/;\s*\n/)
  .map(s => s.trim())
  .filter(s => s && !s.startsWith('--'));

let applied = 0;
for (const stmt of statements) {
  try {
    await sql.unsafe(stmt);
    applied++;
    // Print first 80 chars as progress indicator.
    console.log(`  ✓ ${stmt.slice(0, 80).replace(/\n/g, ' ')}…`);
  } catch (err) {
    console.error(`  ✗ FAILED: ${stmt.slice(0, 120)}`);
    console.error(`    ${err.message}`);
    await sql.end();
    process.exit(1);
  }
}

await sql.end();
console.log(`\n[apply-pg-sql] Done — ${applied} statements applied from ${sqlFile}`);
