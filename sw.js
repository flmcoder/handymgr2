const HM_CACHE_VERSION = "hm-static-v1";
const HM_STATIC_ASSETS = [
  "/",
  "/index.html",
  "/css/app.css",
  "/js/app.js",
  "/login-effects.js",
  "/assets/logo.png",
];

const HM_QUEUE_DB = "hm-offline-queue";
const HM_QUEUE_STORE = "requests";

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(HM_CACHE_VERSION)
      .then((cache) => cache.addAll(HM_STATIC_ASSETS))
      .then(() => self.skipWaiting())
      .catch(() => self.skipWaiting()),
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

async function notifyClients(message) {
  const clients = await self.clients.matchAll({ type: "window", includeUncontrolled: true });
  for (const client of clients) {
    client.postMessage(message);
  }
}

async function replayQueue() {
  const queued = await readQueuedRequests();
  if (!queued.length) return { sent: 0, failed: 0, remaining: 0 };

  let sent = 0;
  let failed = 0;

  for (const item of queued) {
    try {
      const res = await fetch(item.url, {
        method: item.method,
        headers: item.headers,
        body: item.bodyText || undefined,
      });
      if (res.ok || (res.status >= 400 && res.status < 500)) {
        await removeQueuedRequest(item.id);
        sent += 1;
      } else {
        failed += 1;
      }
    } catch (_) {
      failed += 1;
    }
  }

  const remaining = (await readQueuedRequests()).length;
  await notifyClients({
    type: "HM_QUEUE_REPLAY_RESULT",
    sent,
    failed,
    remaining,
  });
  return { sent, failed, remaining };
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
          cache.put("/index.html", network.clone());
          return network;
        } catch (_) {
          return (await caches.match("/index.html")) || (await caches.match("/")) || Response.error();
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
      try {
        return await fetch(req);
      } catch (_) {
        await queueRequest(req);
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
  if (event.data.type === "HM_REPLAY_QUEUE") {
    event.waitUntil(replayQueue());
  }
});
