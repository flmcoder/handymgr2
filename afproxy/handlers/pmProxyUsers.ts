// ============================================================================
// handlers/pmProxyUsers.ts — Property Manager proxy user syncing from AppFolio
// ============================================================================

import { AF_DB, dbHeaders } from "../config.ts";
import { rowsAsObjects, sqlite } from "../db.ts";

export async function handlePmProxyUsers(): Promise<any> {
  let userList: Record<string, unknown>[] = [];
  let sourceLabel = "appfolio";

  try {
    const afResp = await fetch(`${AF_DB}/api/v0/users?limit=200`, {
      headers: dbHeaders(),
    });

    if (afResp.ok) {
      const rawData: unknown = await afResp.json().catch(() => null);
      if (rawData && typeof rawData === "object") {
        userList = Array.isArray(rawData)
          ? (rawData as Record<string, unknown>[])
          : Array.isArray(
              (rawData as Record<string, unknown[]>).results,
            )
          ? ((rawData as Record<string, unknown[]>).results as Record<
            string,
            unknown
          >[])
          : Array.isArray((rawData as Record<string, unknown[]>).data)
          ? ((rawData as Record<string, unknown[]>).data as Record<
            string,
            unknown
          >[])
          : [];
      }
    } else if (afResp.status === 422) {
      console.warn(
        "[pm_proxy_users] 422 - verify AppFolio API account roles",
      );
    }
  } catch (fetchErr) {
    console.warn(
      "[pm_proxy_users] AppFolio fetch failed:",
      String(fetchErr),
    );
  }

  for (const u of userList) {
    if (!u?.id) continue;
    const fullName = String(
      u.full_name ?? u.name ??
        [u.first_name, u.last_name].filter(Boolean).join(" ") ?? "",
    );
    await sqlite.execute({
      sql: `INSERT OR REPLACE INTO pm_proxy_users
              (id, user_uuid, full_name, email, phone,
               property_group_uuid, roles, is_active, active,
               raw_json, created_at, updated_at)
            VALUES (?,?,?,?,?,?,?,?,?,?,datetime('now'),datetime('now'))`,
      args: [
        String(u.id),
        String(u.id),
        fullName,
        String(u.email ?? ""),
        String(u.phone ?? u.phone_number ?? ""),
        String(u.property_group_uuid ?? u.property_group_id ?? ""),
        JSON.stringify(Array.isArray(u.roles) ? u.roles : []),
        u.is_active !== false ? 1 : 0,
        u.is_active !== false ? 1 : 0,
        JSON.stringify(u),
      ],
    }).catch((e: unknown) =>
      console.warn("[pm_proxy_users upsert]", String(e))
    );
  }

  let cached: any;
  try {
    cached = await sqlite.execute(
      `SELECT id, user_uuid, full_name, email, phone,
              property_group_uuid, roles, is_active, active, raw_json
       FROM   pm_proxy_users
       ORDER  BY full_name ASC`,
    );
  } catch (dbErr) {
    return {
      ok: false,
      error: "db_error",
      message: String(dbErr),
    };
  }

  const cachedRows = rowsAsObjects(cached);
  if (!userList.length) {
    sourceLabel = cachedRows.length
      ? "cache_fallback"
      : "empty";
  }

  return {
    ok: true,
    source: sourceLabel,
    total: cachedRows.length,
    results: cachedRows.map((r: Record<string, unknown>) => ({
      id: String(r.id ?? r.user_uuid ?? ""),
      full_name: String(r.full_name ?? ""),
      email: String(r.email ?? ""),
      phone: String(r.phone ?? ""),
      property_group_uuid: String(r.property_group_uuid ?? ""),
      roles: safeParseJSON(String(r.roles ?? "[]"), []),
      is_active: Number(r.is_active ?? r.active ?? 0) === 1,
    })),
  };
}

// Helper to safely parse JSON
function safeParseJSON<T>(raw: string | null | undefined, fallback: T): T {
  if (!raw) return fallback;
  try {
    return JSON.parse(raw) as T;
  } catch {
    return fallback;
  }
}
