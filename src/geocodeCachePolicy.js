export function normalizeGeocodeAddress(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function isValidCoordinate(value) {
  if (!value || typeof value !== 'object' || value.lat === null || value.lat === undefined || value.lat === '' || value.lon === null || value.lon === undefined || value.lon === '') return false;
  const lat = Number(value.lat);
  const lon = Number(value.lon);
  return Number.isFinite(lat) && lat >= -90 && lat <= 90
    && Number.isFinite(lon) && lon >= -180 && lon <= 180;
}

export function createGeocodeCoordinateCache(storageProvider, storageKey, maxEntries = 2_000) {
  const entries = new Map();
  const capacity = Math.max(1, Number.parseInt(String(maxEntries), 10) || 2_000);
  let persistTimer = null;

  const persist = () => {
    if (persistTimer !== null) {
      clearTimeout(persistTimer);
      persistTimer = null;
    }
    try {
      const storage = storageProvider();
      storage?.setItem(storageKey, JSON.stringify(Array.from(entries.entries())));
    } catch {
      // Keep the in-memory cache usable when storage is unavailable or full.
    }
  };

  try {
    const storage = storageProvider();
    const saved = JSON.parse(storage?.getItem(storageKey) || '[]');
    if (Array.isArray(saved)) {
      for (const entry of saved) {
        if (!Array.isArray(entry) || entry.length !== 2 || !isValidCoordinate(entry[1])) continue;
        const key = normalizeGeocodeAddress(entry[0]);
        if (key) entries.set(key, { lat: Number(entry[1].lat), lon: Number(entry[1].lon) });
      }
    }
  } catch {
    // Ignore malformed persisted data and continue with an empty cache.
  }

  while (entries.size > capacity) entries.delete(entries.keys().next().value);

  return {
    get(address) {
      const key = normalizeGeocodeAddress(address);
      return key ? entries.get(key) || null : null;
    },
    set(address, coordinates) {
      const key = normalizeGeocodeAddress(address);
      if (!key || !isValidCoordinate(coordinates)) return false;
      if (entries.has(key)) entries.delete(key);
      entries.set(key, { lat: Number(coordinates.lat), lon: Number(coordinates.lon) });
      while (entries.size > capacity) entries.delete(entries.keys().next().value);
      if (persistTimer !== null) clearTimeout(persistTimer);
      persistTimer = setTimeout(persist, 100);
      return true;
    },
    flush: persist,
    get size() {
      return entries.size;
    },
  };
}

export async function mapWithConcurrency(items, concurrency, mapper) {
  const values = Array.from(items || []);
  if (values.length === 0) return [];
  const workerCount = Math.min(values.length, Math.max(1, Math.floor(Number(concurrency) || 1)));
  const results = new Array(values.length);
  let nextIndex = 0;

  async function worker() {
    while (true) {
      const index = nextIndex++;
      if (index >= values.length) return;
      results[index] = await mapper(values[index], index);
    }
  }

  await Promise.all(Array.from({ length: workerCount }, worker));
  return results;
}

export function createConcurrencyLimiter(limit) {
  const capacity = Math.max(1, Math.floor(Number(limit) || 1));
  let active = 0;
  const waiting = [];

  async function releaseNext() {
    active--;
    const next = waiting.shift();
    if (next) {
      active++;
      next();
    }
  }

  return {
    async run(task) {
      if (active >= capacity) {
        await new Promise((resolve) => waiting.push(resolve));
      } else {
        active++;
      }
      try {
        return await task();
      } finally {
        await releaseNext();
      }
    },
    get activeCount() {
      return active;
    },
    get pendingCount() {
      return waiting.length;
    },
  };
}