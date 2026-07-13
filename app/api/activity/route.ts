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
      where.push("action = ?");
      args.push(action);
    }
    const email = String(searchParams.get("email") || "").trim();
    if (email) {
      where.push("user_email LIKE ?");
      args.push(`%${email}%`);
    }
    const projectId = String(searchParams.get("project_id") || "").trim();
    if (projectId) {
      where.push("project_id = ?");
      args.push(projectId);
    }
    const date = String(searchParams.get("date") || "").trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      where.push("DATE(created_at) = ?");
      args.push(date);
    }
    const q = String(searchParams.get("q") || "").trim();
    if (q) {
      where.push(
        "(user_email LIKE ? OR action LIKE ? OR table_name LIKE ? OR entity_id LIKE ? OR details LIKE ? OR ip_address LIKE ?)"
      );
      const like = `%${q}%`;
      args.push(like, like, like, like, like, like);
    }

    const limitRaw = Number(searchParams.get("limit") || 300);
    const limit = Number.isFinite(limitRaw) ? Math.min(Math.max(1, limitRaw), 2000) : 300;

    const whereSql = where.length ? ` WHERE ${where.join(" AND ")}` : "";
    const [rows] = await pool.query(
      `SELECT id, created_at, user_id, user_email, action, table_name, entity_id, project_id, row_count, details, ip_address, user_agent
       FROM activity_log${whereSql}
       ORDER BY created_at DESC
       LIMIT ?`,
      [...args, limit]
    );

    // Small summary so the UI can show top actors / actions at a glance.
    const [summary] = await pool.query(
      `SELECT action, COUNT(*) AS n FROM activity_log${whereSql} GROUP BY action ORDER BY n DESC`,
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
