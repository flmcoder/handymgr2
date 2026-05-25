import { readFile } from 'node:fs/promises';
import process from 'node:process';
import { createClient } from '@libsql/client';

const sqlFile = process.argv[2];

if (!sqlFile) {
  console.error('Usage: node scripts/apply-turso-sql.mjs <sql-file-path>');
  process.exit(1);
}

if (!process.env.TURSO_DATABASE_URL || !process.env.TURSO_AUTH_TOKEN) {
  console.error('Missing TURSO_DATABASE_URL or TURSO_AUTH_TOKEN environment variables.');
  process.exit(1);
}

const sql = await readFile(sqlFile, 'utf8');
const db = createClient({
  url: process.env.TURSO_DATABASE_URL,
  authToken: process.env.TURSO_AUTH_TOKEN,
});

await db.batch(
  sql
    .split(';')
    .map((s) => s.trim())
    .filter(Boolean)
    .map((statement) => ({ sql: statement })),
  'write'
);

console.log(`Applied SQL file: ${sqlFile}`);