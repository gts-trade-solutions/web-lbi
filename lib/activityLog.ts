/* eslint-disable @typescript-eslint/no-explicit-any */
// Central activity / audit log. Records who did what and when — logins,
// creates, edits, deletes, bulk imports, exports — so data changes can be
// traced to a user + time instead of being blamed on "the app". Logging is
// ALWAYS best-effort: any failure here is swallowed so it can never break a
// real request.
import { v4 as uuidv4 } from "uuid";
import pool from "./db";
import { getAuthUser } from "./auth";

let tableReady = false;

export async function ensureActivityLogTable() {
  if (tableReady) return;
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS activity_log (
        id VARCHAR(36) PRIMARY KEY,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        user_id VARCHAR(36) NULL,
        user_email VARCHAR(255) NULL,
        action VARCHAR(64) NOT NULL,
        table_name VARCHAR(64) NULL,
        entity_id TEXT NULL,
        project_id VARCHAR(64) NULL,
        row_count INT NULL,
        details LONGTEXT NULL,
        ip_address VARCHAR(64) NULL,
        user_agent TEXT NULL,
        INDEX idx_activity_created (created_at),
        INDEX idx_activity_user (user_id),
        INDEX idx_activity_action (action),
        INDEX idx_activity_project (project_id)
      )
    `);
    tableReady = true;
  } catch (err) {
    console.error("[activityLog] ensure table failed:", err);
  }
}

// Best-effort client IP from the proxy headers (nginx sets x-forwarded-for).
function clientIp(request: Request): string | null {
  const xff = request.headers.get("x-forwarded-for");
  if (xff) return xff.split(",")[0].trim();
  return request.headers.get("x-real-ip") || null;
}

export type ActivityInput = {
  action: string; // login | login_failed | logout | create | update | delete | bulk_import | export | move_project
  table?: string | null;
  entityId?: string | string[] | null;
  projectId?: string | null;
  rowCount?: number | null;
  details?: unknown;
  // When the actor isn't derivable from the JWT (e.g. a failed login), pass
  // the email/id explicitly.
  actorEmail?: string | null;
  actorId?: string | null;
};

export async function logActivity(request: Request, input: ActivityInput) {
  try {
    await ensureActivityLogTable();

    let userId = input.actorId ?? null;
    let userEmail = input.actorEmail ?? null;
    if (!userId || !userEmail) {
      const actor = getAuthUser(request) as any;
      if (actor) {
        userId = userId || actor.id || null;
        userEmail = userEmail || actor.email || null;
      }
    }

    const entityId = Array.isArray(input.entityId)
      ? input.entityId.filter(Boolean).join(",").slice(0, 4000)
      : input.entityId
        ? String(input.entityId).slice(0, 4000)
        : null;

    let details: string | null = null;
    if (input.details != null) {
      try {
        details =
          typeof input.details === "string"
            ? input.details
            : JSON.stringify(input.details);
        if (details && details.length > 60000) details = details.slice(0, 60000);
      } catch {
        details = null;
      }
    }

    await pool.query(
      `INSERT INTO activity_log
        (id, user_id, user_email, action, table_name, entity_id, project_id, row_count, details, ip_address, user_agent)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        uuidv4(),
        userId,
        userEmail,
        String(input.action || "unknown").slice(0, 64),
        input.table ? String(input.table).slice(0, 64) : null,
        entityId,
        input.projectId ? String(input.projectId).slice(0, 64) : null,
        input.rowCount ?? null,
        details,
        clientIp(request),
        (request.headers.get("user-agent") || "").slice(0, 1000) || null,
      ]
    );
  } catch (err) {
    // Never let logging break the real request.
    console.error("[activityLog] insert failed:", err);
  }
}
