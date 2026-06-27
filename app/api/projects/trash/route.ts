import pool from "../../../../lib/db";
import { requireAuth } from "../../../../lib/auth";
import {
  ensureTrashColumns,
  purgeExpiredTrash,
  TRASH_RETENTION_DAYS,
} from "../../../../lib/projects-trash";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

function isUnauthorizedError(e: unknown) {
  return (e as { message?: string })?.message === "Unauthorized";
}

// GET /api/projects/trash — list soft-deleted projects in the recycle bin.
// Anything older than the retention window is purged first.
export async function GET(request: Request) {
  try {
    requireAuth(request);

    const hasTrash = await ensureTrashColumns();
    if (!hasTrash) {
      return Response.json({ projects: [], retentionDays: TRASH_RETENTION_DAYS });
    }

    await purgeExpiredTrash();

    const [rows] = await pool.query(
      "SELECT * FROM projects WHERE deleted_at IS NOT NULL ORDER BY deleted_at DESC"
    );
    const projects = Array.isArray(rows) ? rows : [];

    return Response.json({ projects, retentionDays: TRASH_RETENTION_DAYS });
  } catch (error) {
    if (isUnauthorizedError(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/projects/trash] GET error:", error);
    return Response.json({ error: "Failed to fetch recycle bin" }, { status: 500 });
  }
}
