import 'dotenv/config';
import postgres from 'postgres';
import { drizzle } from 'drizzle-orm/postgres-js';
import * as schema from './schema';

const SQL_SE = String(process.env.SQL_SE || '').trim();
const DBI = String(process.env.DBI || process.env.PGHOST || '127.0.0.1').trim();
const DB_PORT = Number(process.env.DB_PORT || process.env.PGPORT || 6432);
const DB_NAME = String(process.env.DB_NAME || process.env.PGDATABASE || '').trim();
const DB_USER = String(process.env.DB_USER || process.env.PGUSER || '').trim();
const DB_PASSWORD = String(process.env.DB_PASSWORD || process.env.PGPASSWORD || '').trim();

function buildConnectionString(): string {
  if (SQL_SE) return SQL_SE;
  if (!DB_NAME) {
    throw new Error('Missing DB_NAME (or SQL_SE) for PostgreSQL connection');
  }

  const auth = DB_USER
    ? `${encodeURIComponent(DB_USER)}${DB_PASSWORD ? `:${encodeURIComponent(DB_PASSWORD)}` : ''}@`
    : '';

  return `postgres://${auth}${DBI}:${DB_PORT}/${DB_NAME}`;
}

const connectionString = buildConnectionString();

export const queryClient = postgres(connectionString, {
  max: Number(process.env.DB_POOL_MAX || 10),
  idle_timeout: Number(process.env.DB_IDLE_TIMEOUT || 20),
  connect_timeout: Number(process.env.DB_CONNECT_TIMEOUT || 10),
  prepare: false,
  ssl: String(process.env.DB_SSL || '').toLowerCase() === 'true' ? 'require' : undefined,
});

export const db = drizzle(queryClient, { schema });

export async function pingDatabase(): Promise<boolean> {
  try {
    await queryClient`select 1 as ok`;
    return true;
  } catch (error) {
    console.error('[db] PostgreSQL/PgBouncer connection failure:', error);
    return false;
  }
}

export async function closeDatabase(): Promise<void> {
  await queryClient.end({ timeout: 5 });
}
