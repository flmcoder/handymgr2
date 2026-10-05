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
let restartAttempts = 0;
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

/**
 * Resolves once the local port can be bound (i.e. nothing else is listening on
 * it). ssh is launched with ExitOnForwardFailure=yes, so if the port is still
 * held by a lingering ssh child or a TIME_WAIT socket, ssh exits 255 with
 * "cannot listen to port" and triggers an endless restart loop. Waiting for the
 * port to be free before spawning breaks that loop.
 */
export function waitForPortFree(port: number, timeoutMs: number): Promise<void> {
  const startedAt = Date.now();

  return new Promise((resolve, reject) => {
    const tryBind = () => {
      const tester = net.createServer();

      tester.once('error', (err: NodeJS.ErrnoException) => {
        tester.close();
        if (err.code === 'EADDRINUSE' || err.code === 'EACCES') {
          if (Date.now() - startedAt >= timeoutMs) {
            reject(new Error(`Local port 127.0.0.1:${port} still in use after ${timeoutMs}ms`));
            return;
          }
          setTimeout(tryBind, 250);
          return;
        }
        reject(err);
      });

      tester.once('listening', () => {
        tester.close(() => resolve());
      });

      tester.listen(port, '127.0.0.1');
    };

    tryBind();
  });
}

/**
 * Sends SIGTERM (then SIGKILL as a fallback) to an ssh child and resolves once
 * it has exited, so the local forward port is released before we respawn.
 */
function killChild(child: ChildProcessWithoutNullStreams): Promise<void> {
  return new Promise((resolve) => {
    if (child.exitCode !== null || child.signalCode !== null) {
      resolve();
      return;
    }

    child.once('exit', () => resolve());

    try {
      child.kill('SIGTERM');
    } catch {
      resolve();
      return;
    }

    setTimeout(() => {
      if (child.exitCode === null && child.signalCode === null) {
        try {
          child.kill('SIGKILL');
        } catch {
          /* ignore */
        }
      }
    }, 2_000);
  });
}

/**
 * Exponential backoff (capped at 30s) for tunnel restart attempts. Exported so
 * the backoff schedule can be asserted in tests without touching timers.
 */
export function computeRestartDelay(attempt: number): number {
  const safeAttempt = Math.max(1, Math.floor(attempt));
  return Math.min(2_000 * 2 ** (safeAttempt - 1), 30_000);
}

function scheduleRestart(config: TunnelConfig): void {
  if (shutdown || restartTimer) {
    return;
  }

  restartAttempts += 1;
  const delay = computeRestartDelay(restartAttempts) + Math.floor(Math.random() * 1_000);
  console.log(`[ssh-tunnel] Scheduling tunnel restart #${restartAttempts} in ${delay}ms`);

  restartTimer = setTimeout(() => {
    restartTimer = null;
    if (shutdown || startPromise || runtime) {
      return;
    }
    void requestLaunch(config).catch((error) => {
      console.error('[ssh-tunnel] restart failed:', String((error as Error).message || error));
    });
  }, delay);
}

/**
 * Single-flight launcher. Both the initial start and every restart funnel
 * through here so there is never more than one ssh child competing for the
 * local forward port at a time.
 */
function requestLaunch(config: TunnelConfig): Promise<void> {
  if (startPromise) {
    return startPromise;
  }
  startPromise = startTunnel(config, restartAttempts > 0).finally(() => {
    startPromise = null;
  });
  return startPromise;
}

async function startTunnel(config: TunnelConfig, isRestart = false): Promise<void> {
  installExitHooks();

  if (shutdown) {
    return;
  }

  // Kill any prior ssh child and wait for it to exit so it releases the port.
  if (runtime?.child) {
    const previousChild = runtime.child;
    runtime = null;
    await killChild(previousChild);
  }

  const identityFile = await ensureIdentityFile(config);

  // Ensure the local forward port is actually free before spawning, otherwise
  // ExitOnForwardFailure=yes makes ssh exit 255 ("cannot listen to port").
  await waitForPortFree(config.localPort, Number(readEnv('SSH_DB_TUNNEL_PORT_FREE_TIMEOUT_MS') || 10_000));

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
        restartAttempts = 0;
        resolve();
      })
      .catch((error) => {
        child.kill('SIGTERM');
        scheduleRestart(config);
        reject(error);
      });
  });
}

export function isSshDbTunnelEnabled(): boolean {
  if (envFlag('SKIP_TUNNEL')) {
    return false;
  }
  return envFlag('SSH_DB_TUNNEL_ENABLED');
}

export function getSshDbTunnelTarget(): { host: string; port: number } | null {
  const skipTunnel = envFlag('SKIP_TUNNEL');
  const tunnelEnabled = envFlag('SSH_DB_TUNNEL_ENABLED');
  
  if (!skipTunnel && !tunnelEnabled) {
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
    await requestLaunch(config);
    return;
  }

  await startPromise;
}
