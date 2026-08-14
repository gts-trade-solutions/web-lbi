/* eslint-disable @typescript-eslint/no-explicit-any */
import pool from "../../../lib/db";
import { requireAuth } from "../../../lib/auth";
import { ensureActivityLogTable } from "../../../lib/activityLog";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

function unauthorized(error: unknown) {
  return (error as { message?: string })?.message === "Unauthorized";
}

// GET /api/activity — read the audit log with optional filters:
//   ?action=delete            only that action
//   ?email=someone@x.com      only that user (partial match)
//   ?project_id=<uuid>        only actions on that project
//   ?date=2026-07-02          only that calendar day
//   ?q=<text>                 search email / action / entity / details
//   ?limit=500                cap rows (default 300, max 2000)
export async function GET(request: Request) {
  try {
    requireAuth(request);
    await ensureActivityLogTable();

    const { searchParams } = new URL(request.url);
    const where: string[] = [];
    const args: any[] = [];

    const action = String(searchParams.get("action") || "").trim();
    if (action) {
      where.push("al.action = ?");
      args.push(action);
    }
    const email = String(searchParams.get("email") || "").trim();
    if (email) {
      where.push("al.user_email LIKE ?");
      args.push(`%${email}%`);
    }
    const projectId = String(searchParams.get("project_id") || "").trim();
    if (projectId) {
      where.push("al.project_id = ?");
      args.push(projectId);
    }
    const date = String(searchParams.get("date") || "").trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      where.push("DATE(al.created_at) = ?");
      args.push(date);
    }
    const q = String(searchParams.get("q") || "").trim();
    if (q) {
      where.push(
        "(al.user_email LIKE ? OR al.action LIKE ? OR al.table_name LIKE ? OR al.entity_id LIKE ? OR al.details LIKE ? OR al.ip_address LIKE ? OR p.name LIKE ?)"
      );
      const like = `%${q}%`;
      args.push(like, like, like, like, like, like, like);
    }

    const limitRaw = Number(searchParams.get("limit") || 300);
    const limit = Number.isFinite(limitRaw) ? Math.min(Math.max(1, limitRaw), 2000) : 300;

    const whereSql = where.length ? ` WHERE ${where.join(" AND ")}` : "";
    // LEFT JOIN projects so the log shows the readable project NAME (so a
    // wrong-project save is obvious), and can be searched by project name.
    const [rows] = await pool.query(
      `SELECT al.id, al.created_at, al.user_id, al.user_email, al.action, al.table_name,
              al.entity_id, al.project_id, p.name AS project_name, al.row_count,
              al.details, al.ip_address, al.user_agent
       FROM activity_log al
       LEFT JOIN projects p ON p.id = al.project_id${whereSql}
       ORDER BY al.created_at DESC
       LIMIT ?`,
      [...args, limit]
    );

    // Small summary so the UI can show top actors / actions at a glance.
    const [summary] = await pool.query(
      `SELECT al.action AS action, COUNT(*) AS n
       FROM activity_log al
       LEFT JOIN projects p ON p.id = al.project_id${whereSql}
       GROUP BY al.action ORDER BY n DESC`,
      args
    );

    return Response.json({
      rows: Array.isArray(rows) ? rows : [],
      summary: Array.isArray(summary) ? summary : [],
    });
  } catch (error) {
    if (unauthorized(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/activity] GET error:", error);
    return Response.json({ error: "Failed to fetch activity log" }, { status: 500 });
  }
}
