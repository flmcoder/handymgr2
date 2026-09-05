const HM_CACHE_VERSION = "hm-static-v8";
const HM_BASE_PATH = new URL(self.registration.scope).pathname;
const HM_SHELL_URL = HM_BASE_PATH + "index.html";
const HM_STATIC_ASSETS = [
  HM_BASE_PATH,
  HM_SHELL_URL,
  HM_BASE_PATH + "assets/logo.png",
  HM_BASE_PATH + "assets/favicon.svg",
  HM_BASE_PATH + "manifest.webmanifest",
];

const HM_QUEUE_DB = "hm-offline-queue";
const HM_QUEUE_STORE = "requests";
const HM_REPLAY_MIN_GAP_MS = 5000;
const HM_REPLAY_MAX_ATTEMPTS = 6;
const HM_REPLAY_BACKOFF_BASE_MS = 3000;
const HM_REPLAY_BACKOFF_MAX_MS = 300000;

let replayInProgress = false;
let replayLastStartedAt = 0;

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(HM_CACHE_VERSION)
      .then((cache) => cache.addAll(HM_STATIC_ASSETS))
      .catch(() => undefined),
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(
      keys
        .filter((k) => k !== HM_CACHE_VERSION)
        .map((k) => caches.delete(k)),
    );
    await self.clients.claim();
  })());
});

function openQueueDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(HM_QUEUE_DB, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(HM_QUEUE_STORE)) {
        db.createObjectStore(HM_QUEUE_STORE, { keyPath: "id", autoIncrement: true });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function txDone(tx) {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve(undefined);
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error);
  });
}

async function queueRequest(request) {
  const clone = request.clone();
  const bodyText = await clone.text().catch(() => "");
  const headers = {};
  for (const [key, value] of clone.headers.entries()) {
    headers[key] = value;
  }

  const item = {
    url: clone.url,
    method: clone.method,
    headers,
    bodyText,
    createdAt: new Date().toISOString(),
    attempts: 0,
    nextAttemptAt: 0,
    lastAttemptAt: 0,
    lastError: "",
  };

  const db = await openQueueDb();
  const tx = db.transaction(HM_QUEUE_STORE, "readwrite");
  tx.objectStore(HM_QUEUE_STORE).add(item);
  await txDone(tx);

  if (self.registration && self.registration.sync) {
    try {
      await self.registration.sync.register("hm-replay-queue");
    } catch (_) {
      // Browsers can deny one-off sync; online replay fallback still works.
    }
  }
}

async function readQueuedRequests() {
  const db = await openQueueDb();
  const tx = db.transaction(HM_QUEUE_STORE, "readonly");
  const store = tx.objectStore(HM_QUEUE_STORE);
  const req = store.getAll();
  const rows = await new Promise((resolve, reject) => {
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
  await txDone(tx);
  return rows;
}

async function removeQueuedRequest(id) {
  const db = await openQueueDb();
  const tx = db.transaction(HM_QUEUE_STORE, "readwrite");
  tx.objectStore(HM_QUEUE_STORE).delete(id);
  await txDone(tx);
}

async function updateQueuedRequest(item) {
  const db = await openQueueDb();
  const tx = db.transaction(HM_QUEUE_STORE, "readwrite");
  tx.objectStore(HM_QUEUE_STORE).put(item);
  await txDone(tx);
}

function isRetryableStatus(status) {
  return status === 408 || status === 425 || status === 429 || status >= 500;
}

function computeBackoffMs(attempts) {
  const exp = Math.max(0, Number(attempts || 0) - 1);
  const base = Math.min(
    HM_REPLAY_BACKOFF_MAX_MS,
    HM_REPLAY_BACKOFF_BASE_MS * Math.pow(2, exp),
  );
  const jitter = Math.floor(Math.random() * 1000);
  return Math.min(HM_REPLAY_BACKOFF_MAX_MS, base + jitter);
}

async function notifyClients(message) {
  const clients = await self.clients.matchAll({ type: "window", includeUncontrolled: true });
  for (const client of clients) {
    client.postMessage(message);
  }
}

async function replayQueue() {
  if (replayInProgress) {
    const remaining = (await readQueuedRequests()).length;
    return { sent: 0, failed: 0, remaining, skipped: true, reason: "replay_in_progress" };
  }

  const now = Date.now();
  if (now - replayLastStartedAt < HM_REPLAY_MIN_GAP_MS) {
    const remaining = (await readQueuedRequests()).length;
    return { sent: 0, failed: 0, remaining, skipped: true, reason: "replay_rate_limited" };
  }

  replayInProgress = true;
  replayLastStartedAt = now;

  const queued = await readQueuedRequests();
  if (!queued.length) {
    replayInProgress = false;
    return { sent: 0, failed: 0, remaining: 0 };
  }

  let sent = 0;
  let failed = 0;
  let deferred = 0;
  let dropped = 0;

  try {
    for (const item of queued) {
      const dueAt = Number(item.nextAttemptAt || 0) || 0;
      if (dueAt > Date.now()) {
        deferred += 1;
        continue;
      }

      try {
        const res = await fetch(item.url, {
          method: item.method,
          headers: item.headers,
          body: item.bodyText || undefined,
        });

        if (res.ok || (res.status >= 400 && res.status < 500 && !isRetryableStatus(res.status))) {
          await removeQueuedRequest(item.id);
          sent += 1;
          continue;
        }

        const attempts = Number(item.attempts || 0) + 1;
        if (attempts >= HM_REPLAY_MAX_ATTEMPTS) {
          await removeQueuedRequest(item.id);
          dropped += 1;
          continue;
        }

        item.attempts = attempts;
        item.lastAttemptAt = Date.now();
        item.lastError = "HTTP " + String(res.status || 0);
        item.nextAttemptAt = Date.now() + computeBackoffMs(attempts);
        await updateQueuedRequest(item);
        failed += 1;
      } catch (err) {
        const attempts = Number(item.attempts || 0) + 1;
        if (attempts >= HM_REPLAY_MAX_ATTEMPTS) {
          await removeQueuedRequest(item.id);
          dropped += 1;
          continue;
        }
        item.attempts = attempts;
        item.lastAttemptAt = Date.now();
        item.lastError = String((err && err.message) || err || "network_error");
        item.nextAttemptAt = Date.now() + computeBackoffMs(attempts);
        await updateQueuedRequest(item);
        failed += 1;
      }
    }

    const remaining = (await readQueuedRequests()).length;
    await notifyClients({
      type: "HM_QUEUE_REPLAY_RESULT",
      sent,
      failed,
      deferred,
      dropped,
      remaining,
    });

    if (remaining > 0 && self.registration && self.registration.sync) {
      try {
        await self.registration.sync.register("hm-replay-queue");
      } catch (_) {
        // Ignore; manual replay or next connectivity event will retry.
      }
    }

    return { sent, failed, deferred, dropped, remaining };
  } finally {
    replayInProgress = false;
  }
}

function isQueueableMutation(request) {
  if (!["POST", "PUT", "PATCH", "DELETE"].includes(request.method)) {
    return false;
  }
  const url = new URL(request.url);
  return url.searchParams.has("action") || url.pathname.includes("/api/") || url.pathname.includes(".val.run");
}

self.addEventListener("fetch", (event) => {
  const req = event.request;
  const url = new URL(req.url);

  if (req.method === "GET" && url.origin === self.location.origin) {
    if (req.mode === "navigate") {
      event.respondWith((async () => {
        try {
          const network = await fetch(req);
          const cache = await caches.open(HM_CACHE_VERSION);
          cache.put(HM_SHELL_URL, network.clone());
          return network;
        } catch (_) {
          return (await caches.match(HM_SHELL_URL)) || (await caches.match(HM_BASE_PATH)) || Response.error();
        }
      })());
      return;
    }

    event.respondWith((async () => {
      const cached = await caches.match(req);
      const networkPromise = fetch(req)
        .then(async (res) => {
          if (res && res.ok) {
            const cache = await caches.open(HM_CACHE_VERSION);
            cache.put(req, res.clone());
          }
          return res;
        })
        .catch(() => null);

      if (cached) return cached;
      const network = await networkPromise;
      return network || Response.error();
    })());
    return;
  }

  if (isQueueableMutation(req)) {
    event.respondWith((async () => {
      const reqForNetwork = req.clone();
      const reqForQueue = req.clone();
      try {
        return await fetch(reqForNetwork);
      } catch (_) {
        await queueRequest(reqForQueue);
        const remaining = (await readQueuedRequests()).length;
        await notifyClients({
          type: "HM_REQUEST_QUEUED",
          remaining,
        });
        return new Response(
          JSON.stringify({
            ok: true,
            queued: true,
            offline: true,
            message: "Request queued offline and will retry automatically.",
          }),
          {
            status: 202,
            headers: { "Content-Type": "application/json" },
          },
        );
      }
    })());
  }
});

self.addEventListener("sync", (event) => {
  if (event.tag === "hm-replay-queue") {
    event.waitUntil(replayQueue());
  }
});

self.addEventListener("message", (event) => {
  if (!event.data || typeof event.data !== "object") return;
  if (event.data.type === "SKIP_WAITING") {
    self.skipWaiting();
    return;
  }
  if (event.data.type === "HM_REPLAY_QUEUE") {
    event.waitUntil(replayQueue());
  }
});
