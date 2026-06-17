const env = (name: string, fallback = ""): string =>
  String(Deno.env.get(name) || fallback || "").trim();

export const PROXY_APP_VERSION = env("PROXY_APP_VERSION", "v9.7.4");

export const AF_DB = env("AF_DB", env("APPFOLIO_DB_BASE_URL", "https://api.appfolio.com"));
export const AF_REPORTS = env(
  "AF_REPORTS",
  env("APPFOLIO_REPORTS_BASE_URL", "https://api.appfolio.com"),
);

export const DEV = env("AF_DEVELOPER_ID", env("DEV", ""));
export const V2_CLIENT_ID = env("AF_V2_CLIENT_ID", env("V2_CLIENT_ID", ""));
export const V2_CLIENT_SECRET = env("AF_V2_CLIENT_SECRET", env("V2_CLIENT_SECRET", ""));

export const CRON_SECRET = env("CRON_SECRET", "");
export const PROXY_ADMIN_KEY = env("PROXY_ADMIN_KEY", env("ADMIN_SECRET", ""));
export const MAGIC_LINK_SECRET = env("MAGIC_LINK_SECRET", env("FRONTEND_PROXY_SECRET", ""));

export const CORS_HEADERS: Record<string, string> = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization,content-type,x-admin-key,x-admin-token,x-cron-secret,x-proxy-token",
  "Access-Control-Allow-Methods": "GET,POST,PUT,PATCH,DELETE,OPTIONS,HEAD",
};

function basicAuth(id: string, secret: string): string {
  const raw = `${id}:${secret}`;
  return `Basic ${btoa(raw)}`;
}

export function dbHeaders(): Record<string, string> {
  const headers: Record<string, string> = {
    Accept: "application/json",
    "Content-Type": "application/json",
  };
  if (V2_CLIENT_ID && V2_CLIENT_SECRET) {
    headers.Authorization = basicAuth(V2_CLIENT_ID, V2_CLIENT_SECRET);
  }
  if (DEV) headers["X-Developer-Id"] = DEV;
  return headers;
}

export function reportsHeaders(): Record<string, string> {
  return dbHeaders();
}
