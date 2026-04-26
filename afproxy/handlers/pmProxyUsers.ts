// ============================================================================
// handlers/pmProxyUsers.ts — Property Manager proxy user syncing from AppFolio
// ============================================================================

import { AF_DB, dbHeaders } from "../config.ts";
import { rowsAsObjects, sqlite } from "../db.ts";

export async function handlePmProxyUsers(): Promise<any> {
  let userList: Record<string, unknown>[] = [];
  let sourceLabel = "appfolio";

  let tableCols = new Set<string>();
  try {
    const tableInfo = await sqlite.execute(`PRAGMA table_info(pm_proxy_users)`);
    tableCols = new Set(
      rowsAsObjects(tableInfo).map((r: any) => String(r.name || "")),
    );
  } catch (_) {
    tableCols = new Set();
  }

  const hasCol = (name: string) => tableCols.has(name);

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

    const insertCols = ["user_uuid", "email", "full_name", "phone"];
    const insertArgs: any[] = [
      String(u.id),
      String(u.email ?? ""),
      fullName,
      String(u.phone ?? u.phone_number ?? ""),
    ];
    if (hasCol("id")) {
      insertCols.unshift("id");
      insertArgs.unshift(String(u.id));
    }
    if (hasCol("property_group_uuid")) {
      insertCols.push("property_group_uuid");
      insertArgs.push(String(u.property_group_uuid ?? u.property_group_id ?? ""));
    }
    if (hasCol("roles")) {
      insertCols.push("roles");
      insertArgs.push(JSON.stringify(Array.isArray(u.roles) ? u.roles : []));
    }
    if (hasCol("is_active")) {
      insertCols.push("is_active");
      insertArgs.push(u.is_active !== false ? 1 : 0);
    }
    if (hasCol("active")) {
      insertCols.push("active");
      insertArgs.push(u.is_active !== false ? 1 : 0);
    }
    if (hasCol("raw_json")) {
      insertCols.push("raw_json");
      insertArgs.push(JSON.stringify(u));
    }
    if (hasCol("created_at")) insertCols.push("created_at");
    if (hasCol("updated_at")) insertCols.push("updated_at");

    const valuePlaceholders = insertCols.map((c) =>
      (c === "created_at" || c === "updated_at") ? "datetime('now')" : "?"
    );

    const updateClauses = [
      "full_name=excluded.full_name",
      "phone=excluded.phone",
    ];
    if (hasCol("property_group_uuid")) updateClauses.push("property_group_uuid=excluded.property_group_uuid");
    if (hasCol("roles")) updateClauses.push("roles=excluded.roles");
    if (hasCol("is_active")) updateClauses.push("is_active=excluded.is_active");
    if (hasCol("active")) updateClauses.push("active=excluded.active");
    if (hasCol("raw_json")) updateClauses.push("raw_json=excluded.raw_json");
    if (hasCol("id")) updateClauses.push("id=COALESCE(NULLIF(pm_proxy_users.id,''),excluded.id)");
    if (hasCol("updated_at")) updateClauses.push("updated_at=datetime('now')");

    await sqlite.execute({
      sql: `INSERT INTO pm_proxy_users (${insertCols.join(", ")})
            VALUES (${valuePlaceholders.join(", ")})
            ON CONFLICT(email) DO UPDATE SET ${updateClauses.join(", ")}`,
      args: insertArgs,
    }).catch((e: unknown) =>
      console.warn("[pm_proxy_users upsert]", String(e))
    );
  }

  let cached: any;
  try {
    const selectCols = ["user_uuid", "full_name", "email", "phone"];
    if (hasCol("id")) selectCols.unshift("id");
    if (hasCol("property_group_uuid")) selectCols.push("property_group_uuid");
    if (hasCol("roles")) selectCols.push("roles");
    if (hasCol("is_active")) selectCols.push("is_active");
    if (hasCol("active")) selectCols.push("active");
    if (hasCol("raw_json")) selectCols.push("raw_json");
    cached = await sqlite.execute(
      `SELECT ${selectCols.join(", ")}
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
