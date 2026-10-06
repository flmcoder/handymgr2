import assert from 'node:assert/strict';
import net from 'node:net';
import test from 'node:test';

import { computeRestartDelay, getSshDbTunnelTarget, waitForPortFree } from './sshTunnel.ts';

function withEnvironment(values: Record<string, string | undefined>, run: () => void): void {
  const previous = new Map(Object.keys(values).map((key) => [key, process.env[key]]));
  for (const [key, value] of Object.entries(values)) {
    if (value === undefined) {
      delete process.env[key];
    } else {
      process.env[key] = value;
    }
  }

  try {
    run();
  } finally {
    for (const [key, value] of previous) {
      if (value === undefined) {
        delete process.env[key];
      } else {
        process.env[key] = value;
      }
    }
  }
}

test('getSshDbTunnelTarget uses localhost only when the tunnel is enabled', () => {
  withEnvironment({ SSH_DB_TUNNEL_ENABLED: 'true', SKIP_TUNNEL: undefined, SSH_DB_TUNNEL_LOCAL_PORT: '15432' }, () => {
    assert.deepEqual(getSshDbTunnelTarget(), { host: '127.0.0.1', port: 15432 });
  });

  withEnvironment({ SSH_DB_TUNNEL_ENABLED: undefined, SKIP_TUNNEL: undefined }, () => {
    assert.equal(getSshDbTunnelTarget(), null);
  });
});

test('getSshDbTunnelTarget leaves DBI and DB_PORT untouched when SKIP_TUNNEL is enabled', () => {
  withEnvironment({ SSH_DB_TUNNEL_ENABLED: 'true', SKIP_TUNNEL: 'true' }, () => {
    assert.equal(getSshDbTunnelTarget(), null);
  });

  withEnvironment({ SSH_DB_TUNNEL_ENABLED: undefined, SKIP_TUNNEL: 'true' }, () => {
    assert.equal(getSshDbTunnelTarget(), null);
  });
});

test('computeRestartDelay grows exponentially from 2s and caps at 30s', () => {
  assert.equal(computeRestartDelay(1), 2_000);
  assert.equal(computeRestartDelay(2), 4_000);
  assert.equal(computeRestartDelay(3), 8_000);
  assert.equal(computeRestartDelay(4), 16_000);
  // 2000 * 2^4 = 32000, capped to 30000.
  assert.equal(computeRestartDelay(5), 30_000);
  assert.equal(computeRestartDelay(50), 30_000);
});

test('computeRestartDelay clamps non-positive and fractional attempts to attempt 1', () => {
  assert.equal(computeRestartDelay(0), 2_000);
  assert.equal(computeRestartDelay(-3), 2_000);
  assert.equal(computeRestartDelay(1.9), 2_000);
});

async function getFreePort(): Promise<number> {
  return new Promise((resolve, reject) => {
    const srv = net.createServer();
    srv.once('error', reject);
    srv.listen(0, '127.0.0.1', () => {
      const addr = srv.address();
      if (addr && typeof addr === 'object') {
        const { port } = addr;
        srv.close(() => resolve(port));
      } else {
        srv.close(() => reject(new Error('Could not acquire a free port')));
      }
    });
  });
}

test('waitForPortFree resolves immediately when the port is free', async () => {
  const port = await getFreePort();
  await assert.doesNotReject(() => waitForPortFree(port, 1_000));
});

test('waitForPortFree times out and rejects while the port is held', async () => {
  const port = await getFreePort();
  const holder = net.createServer();
  await new Promise<void>((resolve) => holder.listen(port, '127.0.0.1', () => resolve()));

  try {
    await assert.rejects(
      () => waitForPortFree(port, 500),
      /still in use/,
    );
  } finally {
    await new Promise<void>((resolve) => holder.close(() => resolve()));
  }
});

test('waitForPortFree resolves once a held port is released before the timeout', async () => {
  const port = await getFreePort();
  const holder = net.createServer();
  await new Promise<void>((resolve) => holder.listen(port, '127.0.0.1', () => resolve()));

  // Release the port shortly after we start waiting; the retry loop should then
  // observe it as free and resolve well within the generous timeout.
  setTimeout(() => holder.close(), 300);

  await assert.doesNotReject(() => waitForPortFree(port, 5_000));
});
