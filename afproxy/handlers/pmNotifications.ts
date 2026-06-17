import { rowsAsObjects, sqlite } from "../db.ts";

type SessionContext = {
  role?: string;
  user_name?: string;
  property_group_uuid?: string;
} | null;

function isUuidLike(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(
    value,
  );
}

function getRole(session: SessionContext): string {
  return String(session?.role || "full").trim().toLowerCase() || "full";
}

function getScopeUuid(session: SessionContext): string {
  return String(session?.property_group_uuid || "").trim();
}

export async function handlePmNotificationsInbox(
  params: Record<string, string>,
  session: SessionContext,
  deviceToken: string,
): Promise<any> {
  const role = getRole(session);
  const sessionScopeUuid = getScopeUuid(session);
  const limit = Math.max(
    1,
    Math.min(200, parseInt(String(params.limit || "75"), 10) || 75),
  );
  const offset = Math.max(0, parseInt(String(params.offset || "0"), 10) || 0);

  const where: string[] = [];
  const args: any[] = [];
  if (role === "pm_readonly" || role === "manager") {
    if (sessionScopeUuid && isUuidLike(sessionScopeUuid)) {
      where.push("(n.scope_group_uuid = '' OR n.scope_group_uuid IS NULL OR n.scope_group_uuid = ?)");
      args.push(sessionScopeUuid);
    } else {
      where.push("(n.scope_group_uuid = '' OR n.scope_group_uuid IS NULL)");
    }
  }

  const whereSql = where.length ? `WHERE ${where.join(" AND ")}` : "";
  const token = String(deviceToken || "").trim();

  const unreadSql = token
    ? `SELECT COUNT(1) AS cnt
       FROM pm_notifications n
       LEFT JOIN pm_notification_reads r
         ON r.notification_uuid = n.uuid AND r.device_token = ?
       ${whereSql} ${whereSql ? "AND" : "WHERE"} r.notification_uuid IS NULL`
    : `SELECT COUNT(1) AS cnt
       FROM pm_notifications n
       ${whereSql}`;
  const unreadArgs = token ? [token, ...args] : [...args];

  const listSql = token
    ? `SELECT
         n.uuid,
         n.message,
         n.scope_group_uuid,
         n.created_by_role,
         n.created_by_user,
         n.created_at,
         CASE WHEN r.notification_uuid IS NULL THEN 0 ELSE 1 END AS read_by_me,
         r.read_at AS read_at
       FROM pm_notifications n
       LEFT JOIN pm_notification_reads r
         ON r.notification_uuid = n.uuid AND r.device_token = ?
       ${whereSql}
       ORDER BY n.created_at DESC
       LIMIT ? OFFSET ?`
    : `SELECT
         n.uuid,
         n.message,
         n.scope_group_uuid,
         n.created_by_role,
         n.created_by_user,
         n.created_at,
         0 AS read_by_me,
         NULL AS read_at
       FROM pm_notifications n
       ${whereSql}
       ORDER BY n.created_at DESC
       LIMIT ? OFFSET ?`;
  const listArgs = token
    ? [token, ...args, limit, offset]
    : [...args, limit, offset];

  const [listRes, unreadRes] = await Promise.all([
    sqlite.execute({ sql: listSql, args: listArgs }),
    sqlite.execute({ sql: unreadSql, args: unreadArgs }),
  ]);

  const rows = rowsAsObjects(listRes);
  const unreadRows = rowsAsObjects(unreadRes);
  const unread = Number(unreadRows[0]?.cnt || 0) || 0;

  return {
    ok: true,
    role,
    unread,
    count: rows.length,
    notifications: rows.map((r: any) => ({
      uuid: String(r.uuid || ""),
      message: String(r.message || ""),
      scope_group_uuid: String(r.scope_group_uuid || ""),
      created_by_role: String(r.created_by_role || ""),
      created_by_user: String(r.created_by_user || ""),
      created_at: String(r.created_at || ""),
      read_by_me: Number(r.read_by_me || 0) === 1,
      read_at: String(r.read_at || ""),
    })),
  };
}

export async function handlePmNotificationPost(
  req: Request,
  session: SessionContext,
): Promise<any> {
  const role = getRole(session);
  if (role !== "manager" && role !== "full") {
    return {
      ok: false,
      status: 403,
      error: "Only manager or full sessions can post inbox notifications",
    };
  }

  let body: any = {};
  try {
    body = await req.json();
  } catch {
    body = {};
  }

  const message = String(body.message || body.body || "").trim();
  if (!message) {
    return { ok: false, status: 400, error: "message is required" };
  }
  if (message.length > 1200) {
    return { ok: false, status: 400, error: "message too long (max 1200)" };
  }

  const providedScope = String(
    body.scope_group_uuid || body.group_uuid || "",
  ).trim();
  const sessionScopeUuid = getScopeUuid(session);
  let scopeGroupUuid = "";

  if (role === "manager") {
    scopeGroupUuid = sessionScopeUuid;
  } else if (providedScope && isUuidLike(providedScope)) {
    scopeGroupUuid = providedScope;
  }

  const uuid = crypto.randomUUID();
  await sqlite.execute({
    sql: `INSERT INTO pm_notifications
      (uuid, message, scope_group_uuid, created_by_role, created_by_user, created_at)
      VALUES (?, ?, ?, ?, ?, datetime('now'))`,
    args: [
      uuid,
      message,
      scopeGroupUuid,
      role,
      String(session?.user_name || "manager"),
    ],
  });

  return {
    ok: true,
    uuid,
    message,
    scope_group_uuid: scopeGroupUuid,
    created_by_role: role,
    created_by_user: String(session?.user_name || "manager"),
  };
}

export async function handlePmNotificationAck(
  req: Request,
  session: SessionContext,
  deviceToken: string,
): Promise<any> {
  const role = getRole(session);
  if (!deviceToken) {
    return { ok: false, status: 401, error: "device token missing" };
  }

  let body: any = {};
  try {
    body = await req.json();
  } catch {
    body = {};
  }

  const uuid = String(body.uuid || body.notification_uuid || "").trim();
  if (!uuid) {
    return { ok: false, status: 400, error: "notification uuid is required" };
  }

  const sessionScopeUuid = getScopeUuid(session);
  const visibleSql =
    role === "pm_readonly" || role === "manager"
      ? `SELECT uuid FROM pm_notifications
         WHERE uuid = ?
           AND (
             scope_group_uuid = '' OR scope_group_uuid IS NULL OR scope_group_uuid = ?
           )
         LIMIT 1`
      : `SELECT uuid FROM pm_notifications WHERE uuid = ? LIMIT 1`;
  const visibleArgs =
    role === "pm_readonly" || role === "manager"
      ? [uuid, sessionScopeUuid]
      : [uuid];
  const visible = rowsAsObjects(await sqlite.execute({
    sql: visibleSql,
    args: visibleArgs,
  }));
  if (!visible.length) {
    return {
      ok: false,
      status: 404,
      error: "notification not found or outside session scope",
    };
  }

  await sqlite.execute({
    sql: `INSERT OR REPLACE INTO pm_notification_reads
      (notification_uuid, device_token, read_at)
      VALUES (?, ?, datetime('now'))`,
    args: [uuid, deviceToken],
  });

  return {
    ok: true,
    uuid,
    acknowledged: true,
    acknowledged_at: new Date().toISOString(),
  };
}