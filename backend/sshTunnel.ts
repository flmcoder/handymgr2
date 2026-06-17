import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process';
import { chmod, mkdir, writeFile } from 'node:fs/promises';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';

type TunnelConfig = {
  host: string;
  port: number;
  user: string;
  localPort: number;
  remoteHost: string;
  remotePort: number;
  identityFile?: string;
  privateKeyText?: string;
};

type TunnelRuntime = {
  child: ChildProcessWithoutNullStreams;
  config: TunnelConfig;
};

let runtime: TunnelRuntime | null = null;
let startPromise: Promise<void> | null = null;
let restartTimer: NodeJS.Timeout | null = null;
let shutdown = false;
let exitHooksInstalled = false;

function envFlag(name: string): boolean {
  return /^(1|true|yes|on)$/i.test(String(process.env[name] || '').trim());
}

function readEnv(name: string): string {
  return String(process.env[name] || '').trim();
}

function readPrivateKeyText(): string | undefined {
  const rawBase64 = readEnv('SSH_DB_TUNNEL_PRIVATE_KEY_B64');
  if (rawBase64) {
    const decoded = Buffer.from(rawBase64, 'base64').toString('utf8').trim();
    if (!decoded.includes('BEGIN OPENSSH PRIVATE KEY')) {
      throw new Error('SSH_DB_TUNNEL_PRIVATE_KEY_B64 did not decode to an OpenSSH private key');
    }
    return decoded;
  }

  const raw = readEnv('SSH_DB_TUNNEL_PRIVATE_KEY');
  if (!raw) {
    return undefined;
  }

  const decoded = raw.replace(/\\n/g, '\n').trim();
  if (!decoded.includes('BEGIN OPENSSH PRIVATE KEY')) {
    throw new Error('SSH_DB_TUNNEL_PRIVATE_KEY did not contain an OpenSSH private key');
  }
  return decoded;
}

function resolveConfig(): TunnelConfig {
  const host = readEnv('SSH_DB_TUNNEL_HOST');
  const user = readEnv('SSH_DB_TUNNEL_USER');

  if (!host || !user) {
    throw new Error('SSH_DB_TUNNEL_HOST and SSH_DB_TUNNEL_USER are required when SSH_DB_TUNNEL_ENABLED=true');
  }

  return {
    host,
    port: Number(readEnv('SSH_DB_TUNNEL_PORT') || 22),
    user,
    localPort: Number(readEnv('SSH_DB_TUNNEL_LOCAL_PORT') || 15432),
    remoteHost: readEnv('SSH_DB_TUNNEL_REMOTE_HOST') || '127.0.0.1',
    remotePort: Number(readEnv('SSH_DB_TUNNEL_REMOTE_PORT') || 6432),
    identityFile: readEnv('SSH_DB_TUNNEL_IDENTITY_FILE') || undefined,
    privateKeyText: readPrivateKeyText(),
  };
}

async function ensureIdentityFile(config: TunnelConfig): Promise<string | undefined> {
  if (config.identityFile) {
    return config.identityFile;
  }

  if (!config.privateKeyText) {
    return undefined;
  }

  const dir = path.join(os.tmpdir(), 'handymgr2-ssh');
  const file = path.join(dir, 'render-db-tunnel.key');
  await mkdir(dir, { recursive: true });
  await writeFile(file, `${config.privateKeyText}\n`, { mode: 0o600 });
  await chmod(file, 0o600);
  console.log('[ssh-tunnel] Created ephemeral SSH identity file from env key');
  return file;
}

function installExitHooks(): void {
  if (exitHooksInstalled) {
    return;
  }

  exitHooksInstalled = true;

  const terminate = () => {
    shutdown = true;
    if (restartTimer) {
      clearTimeout(restartTimer);
      restartTimer = null;
    }
    runtime?.child.kill('SIGTERM');
  };

  process.once('exit', terminate);
  process.once('SIGINT', () => {
    terminate();
    process.exit(130);
  });
  process.once('SIGTERM', () => {
    terminate();
    process.exit(143);
  });
}

function waitForLocalPort(port: number, timeoutMs: number): Promise<void> {
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

function scheduleRestart(config: TunnelConfig): void {
  if (shutdown || restartTimer) {
    return;
  }

  restartTimer = setTimeout(() => {
    restartTimer = null;
    void startTunnel(config, true).catch((error) => {
      console.error('[ssh-tunnel] restart failed:', String((error as Error).message || error));
      scheduleRestart(config);
    });
  }, 2_000);
}

async function startTunnel(config: TunnelConfig, isRestart = false): Promise<void> {
  installExitHooks();

  const identityFile = await ensureIdentityFile(config);
  const destination = `${config.user}@${config.host}`;
  const localBinding = `127.0.0.1:${config.localPort}:${config.remoteHost}:${config.remotePort}`;
  const args = [
    '-N',
    '-L', localBinding,
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

  if (config.port !== 22) {
    args.push('-p', String(config.port));
  }

  args.push(destination);

  console.log(`[ssh-tunnel] ${isRestart ? 'Restarting' : 'Starting'} SSH tunnel: 127.0.0.1:${config.localPort} -> ${config.remoteHost}:${config.remotePort} via ${destination}:${config.port}`);

  const child = spawn('ssh', args, {
    stdio: ['ignore', 'pipe', 'pipe'],
    env: process.env,
  });

  child.stdout.on('data', (chunk) => {
    const message = String(chunk).trim();
    if (message) {
      console.log(`[ssh-tunnel] ${message}`);
    }
  });

  let stderrBuffer = '';
  child.stderr.on('data', (chunk) => {
    const message = String(chunk);
    stderrBuffer += message;
    const trimmed = message.trim();
    if (trimmed) {
      console.error(`[ssh-tunnel] ${trimmed}`);
    }
  });

  await new Promise<void>((resolve, reject) => {
    let settled = false;

    child.once('exit', (code, signal) => {
      if (!settled) {
        reject(new Error(`SSH tunnel exited before ready (code=${String(code)} signal=${String(signal)}) ${stderrBuffer.trim()}`.trim()));
        return;
      }

      runtime = null;
      console.error(`[ssh-tunnel] SSH tunnel exited unexpectedly (code=${String(code)} signal=${String(signal)})`);
      scheduleRestart(config);
    });

    void waitForLocalPort(config.localPort, Number(readEnv('SSH_DB_TUNNEL_READY_TIMEOUT_MS') || 15_000))
      .then(() => {
        settled = true;
        runtime = { child, config };
        resolve();
      })
      .catch((error) => {
        child.kill('SIGTERM');
        reject(error);
      });
  });
}

export function isSshDbTunnelEnabled(): boolean {
  return envFlag('SSH_DB_TUNNEL_ENABLED');
}

export function getSshDbTunnelTarget(): { host: string; port: number } | null {
  if (!isSshDbTunnelEnabled()) {
    return null;
  }

  return {
    host: '127.0.0.1',
    port: Number(readEnv('SSH_DB_TUNNEL_LOCAL_PORT') || 15432),
  };
}

export async function ensureDbSshTunnel(): Promise<void> {
  if (!isSshDbTunnelEnabled() || runtime) {
    return;
  }

  if (!startPromise) {
    const config = resolveConfig();
    console.log('[ssh-tunnel] Config resolved:', {
      host: config.host,
      port: config.port,
      user: config.user,
      localPort: config.localPort,
      remoteHost: config.remoteHost,
      remotePort: config.remotePort,
      hasIdentityFile: Boolean(config.identityFile),
      hasInlinePrivateKey: Boolean(config.privateKeyText),
    });
    startPromise = startTunnel(config).finally(() => {
      startPromise = null;
    });
  }

  await startPromise;
}
