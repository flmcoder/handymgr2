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
 * Optional SSH tunnel env vars for Render runtime:
 *   SSH_DB_TUNNEL_ENABLED=true
 *   SSH_DB_TUNNEL_HOST=hmsh.localto.net
 *   SSH_DB_TUNNEL_USER=Administrator
 *   SSH_DB_TUNNEL_PRIVATE_KEY_B64=<base64-private-key>
 */

import { spawn } from 'node:child_process';
import { chmod, mkdir, readFile, writeFile } from 'node:fs/promises';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';
import postgres from 'postgres';

const sqlFile = process.argv[2];
if (!sqlFile) {
  console.error('Usage: node scripts/apply-pg-sql.mjs <sql-file>');
  process.exit(1);
}

function envFlag(name) {
  return /^(1|true|yes|on)$/i.test(String(process.env[name] || '').trim());
}

function readEnv(name) {
  return String(process.env[name] || '').trim();
}

function getTunnelTarget() {
  if (!envFlag('SSH_DB_TUNNEL_ENABLED')) {
    return null;
  }

  return {
    host: '127.0.0.1',
    port: Number(readEnv('SSH_DB_TUNNEL_LOCAL_PORT') || 15432),
  };
}

async function ensureTunnelIdentityFile() {
  const identityFile = readEnv('SSH_DB_TUNNEL_IDENTITY_FILE');
  if (identityFile) {
    return identityFile;
  }

  const keyB64 = readEnv('SSH_DB_TUNNEL_PRIVATE_KEY_B64');
  const rawKey = keyB64
    ? Buffer.from(keyB64, 'base64').toString('utf8').trim()
    : readEnv('SSH_DB_TUNNEL_PRIVATE_KEY').replace(/\\n/g, '\n').trim();

  if (!rawKey) {
    return undefined;
  }

  const dir = path.join(os.tmpdir(), 'handymgr2-ssh');
  const file = path.join(dir, 'render-db-tunnel.key');
  await mkdir(dir, { recursive: true });
  await writeFile(file, `${rawKey}\n`, { mode: 0o600 });
  await chmod(file, 0o600);
  return file;
}

function waitForLocalPort(port, timeoutMs) {
  const startedAt = Date.now();

  return new Promise((resolve, reject) => {
    const tryConnect = () => {
      const socket = net.createConnection({ host: '127.0.0.1', port });
      socket.once('connect', () => {
        socket.end();
        resolve();
      });
      socket.once('error', () => {
        socket.destroy();
        if (Date.now() - startedAt >= timeoutMs) {
          reject(new Error(`Timed out waiting for local SSH tunnel on 127.0.0.1:${port}`));
          return;
        }
        setTimeout(tryConnect, 250);
      });
    };

    tryConnect();
  });
}

async function ensureDbSshTunnel() {
  if (!envFlag('SSH_DB_TUNNEL_ENABLED')) {
    return;
  }

  const host = readEnv('SSH_DB_TUNNEL_HOST');
  const user = readEnv('SSH_DB_TUNNEL_USER');
  if (!host || !user) {
    throw new Error('SSH_DB_TUNNEL_HOST and SSH_DB_TUNNEL_USER are required when SSH_DB_TUNNEL_ENABLED=true');
  }

  const localPort = Number(readEnv('SSH_DB_TUNNEL_LOCAL_PORT') || 15432);
  const remoteHost = readEnv('SSH_DB_TUNNEL_REMOTE_HOST') || '127.0.0.1';
  const remotePort = Number(readEnv('SSH_DB_TUNNEL_REMOTE_PORT') || 6432);
  const sshPort = Number(readEnv('SSH_DB_TUNNEL_PORT') || 22);
  const identityFile = await ensureTunnelIdentityFile();
  const args = [
    '-N',
    '-L', `127.0.0.1:${localPort}:${remoteHost}:${remotePort}`,
    '-o', 'ExitOnForwardFailure=yes',
    '-o', 'ServerAliveInterval=30',
    '-o', 'ServerAliveCountMax=3',
    '-o', 'StrictHostKeyChecking=no',
    '-o', 'UserKnownHostsFile=/dev/null',
    '-o', 'IdentitiesOnly=yes',
  ];

  if (identityFile) {
    args.push('-i', identityFile);
  }

  if (sshPort !== 22) {
    args.push('-p', String(sshPort));
  }

  args.push(`${user}@${host}`);

  console.log(`[apply-pg-sql] Starting SSH tunnel via ${user}@${host}:${sshPort} -> 127.0.0.1:${localPort}`);

  const child = spawn('ssh', args, {
    stdio: ['ignore', 'pipe', 'pipe'],
    env: process.env,
  });

  let stderr = '';
  child.stderr.on('data', (chunk) => {
    stderr += String(chunk);
  });

  child.once('exit', (code, signal) => {
    if (code !== null && code !== 0) {
      console.error(`[apply-pg-sql] SSH tunnel exited (code=${code} signal=${signal ?? 'null'}) ${stderr.trim()}`.trim());
    }
  });

  await waitForLocalPort(localPort, Number(readEnv('SSH_DB_TUNNEL_READY_TIMEOUT_MS') || 15000));

  const cleanup = async () => {
    child.kill('SIGTERM');
  };

  process.once('exit', cleanup);
  process.once('SIGINT', () => {
    void cleanup();
    process.exit(130);
  });
  process.once('SIGTERM', () => {
    void cleanup();
    process.exit(143);
  });
}

await ensureDbSshTunnel();

const SQL_SE = String(process.env.SQL_SE || '').trim();
const DBI = String(process.env.DBI || process.env.PGHOST || '127.0.0.1').trim();
const DB_PORT = Number(process.env.DB_PORT || process.env.PGPORT || 6432);
const DB_NAME = String(process.env.DB_NAME || process.env.PGDATABASE || '').trim();
const DB_USER = String(process.env.DB_USER || process.env.PGUSER || 'Administrator').trim();
const DB_PASSWORD = String(process.env.DB_PASSWORD || process.env.PGPASSWORD || '').trim();
const DB_SSL = String(process.env.DB_SSL || '').trim().toLowerCase();

function applyTunnelTarget(config) {
  const target = getTunnelTarget();
  return target ? { ...config, host: target.host, port: target.port } : config;
}

let config;
if (SQL_SE && /^postgres(ql)?:\/\//i.test(SQL_SE)) {
  const u = new URL(SQL_SE);
  config = applyTunnelTarget({
    host: u.hostname,
    port: Number(u.port) || DB_PORT,
    database: u.pathname.replace(/^\//, '') || DB_NAME,
    username: decodeURIComponent(u.username) || DB_USER,
    password: decodeURIComponent(u.password) || DB_PASSWORD,
    ssl: DB_SSL === 'true' ? 'require' : false,
  });
} else {
  if (!DB_NAME) {
    console.error('DB_NAME is required.');
    process.exit(1);
  }
  config = applyTunnelTarget({
    host: DBI,
    port: DB_PORT,
    database: DB_NAME,
    username: DB_USER,
    password: DB_PASSWORD,
    ssl: DB_SSL === 'true' ? 'require' : false,
  });
}

console.log(`[apply-pg-sql] Connecting to ${config.host}:${config.port}/${config.database} as ${config.username}`);

const sql = postgres({ ...config, max: 1, connect_timeout: 15, prepare: false });

const raw = await readFile(sqlFile, 'utf8');
const statements = raw
  .split(/;\s*\n/)
  .map((s) => s.trim())
  .filter((s) => s && !s.startsWith('--'));

let applied = 0;
for (const stmt of statements) {
  try {
    await sql.unsafe(stmt);
    applied++;
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
