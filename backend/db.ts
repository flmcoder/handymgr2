import 'dotenv/config';
import postgres from 'postgres';
import { drizzle } from 'drizzle-orm/postgres-js';
import * as schema from './schema';

const SQL_SE = String(process.env.SQL_SE || '').trim();
const DBI = String(process.env.DBI || process.env.PGHOST || '127.0.0.1').trim();
const DB_PORT = Number(process.env.DB_PORT || process.env.PGPORT || 6432);
const DB_NAME = String(process.env.DB_NAME || process.env.PGDATABASE || '').trim();
const DB_USER = String(process.env.DB_USER || process.env.PGUSER || 'Administrator').trim();
const DB_PASSWORD = String(process.env.DB_PASSWORD || process.env.PGPASSWORD || '').trim();
const DB_SSL = String(process.env.DB_SSL || '').trim().toLowerCase();

type DbConfig = {
  host: string;
  port: number;
  database: string;
  username: string;
  password: string;
  ssl: 'require' | false;
};

function resolveDbConfig(): DbConfig {
  if (SQL_SE && /^postgres(ql)?:\/\//i.test(SQL_SE)) {
    try {
      const parsed = new URL(SQL_SE);
      return {
        host: parsed.hostname || DBI,
        port: parsed.port ? Number(parsed.port) : DB_PORT,
        database: (parsed.pathname || '').replace(/^\//, '') || DB_NAME,
        username: decodeURIComponent(parsed.username || DB_USER),
        password: decodeURIComponent(parsed.password || DB_PASSWORD),
        ssl: DB_SSL === 'true' ? 'require' : false,
      };
    } catch (err) {
      console.warn('[db] SQL_SE is present but invalid; falling back to discrete env vars:', String((err as Error).message || err));
    }
  }

  if (!DB_NAME) {
    throw new Error('Missing DB_NAME (or valid SQL_SE) for PostgreSQL connection');
  }

  return {
    host: DBI,
    port: DB_PORT,
    database: DB_NAME,
    username: DB_USER,
    password: DB_PASSWORD,
    ssl: DB_SSL === 'true' ? 'require' : false,
  };
}

const dbConfig = resolveDbConfig();

console.log('[db] PostgreSQL config resolved:', {
  host: dbConfig.host,
  port: dbConfig.port,
  database: dbConfig.database,
  username: dbConfig.username,
  ssl: dbConfig.ssl,
});

export const queryClient = postgres({
  host: dbConfig.host,
  port: dbConfig.port,
  database: dbConfig.database,
  username: dbConfig.username,
  password: dbConfig.password,
  ssl: dbConfig.ssl,
  max: Number(process.env.DB_POOL_MAX || 10),
  idle_timeout: Number(process.env.DB_IDLE_TIMEOUT || 20),
  connect_timeout: Number(process.env.DB_CONNECT_TIMEOUT || 10),
  prepare: false,
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
