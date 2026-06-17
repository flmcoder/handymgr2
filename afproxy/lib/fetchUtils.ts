type Telemetry = {
  total: number;
  failed: number;
  timed_out: number;
  retries: number;
};

const telemetry: Telemetry = {
  total: 0,
  failed: 0,
  timed_out: 0,
  retries: 0,
};

function withTimeout<T>(promise: Promise<T>, timeoutMs: number): Promise<T> {
  if (!timeoutMs || timeoutMs <= 0) return promise;
  return new Promise<T>((resolve, reject) => {
    const id = setTimeout(() => {
      telemetry.timed_out += 1;
      reject(new Error(`Request timed out after ${timeoutMs}ms`));
    }, timeoutMs);
    promise.then(
      (v) => {
        clearTimeout(id);
        resolve(v);
      },
      (err) => {
        clearTimeout(id);
        reject(err);
      },
    );
  });
}

export async function fetchWithTimeout(
  input: string | URL | Request,
  init: RequestInit = {},
  timeoutMs = 30_000,
): Promise<Response> {
  telemetry.total += 1;
  try {
    return await withTimeout(fetch(input, init), timeoutMs);
  } catch (err) {
    telemetry.failed += 1;
    throw err;
  }
}

export function getFetchResilienceTelemetry(): Telemetry {
  return { ...telemetry };
}
