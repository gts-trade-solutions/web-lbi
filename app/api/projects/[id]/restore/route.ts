import pool from "../../../../../lib/db";
import { requireAuth } from "../../../../../lib/auth";
import { ensureTrashColumns } from "../../../../../lib/projects-trash";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type RouteContext = {
  params: {
    id: string;
  };
};

function isUnauthorizedError(e: unknown) {
  return (e as { message?: string })?.message === "Unauthorized";
}

// POST /api/projects/:id/restore — bring a project back out of the recycle bin.
export async function POST(request: Request, context: RouteContext) {
  try {
    requireAuth(request);
    const projectId = String(context.params?.id || "").trim();

    if (!projectId) {
      return Response.json({ error: "Project id is required" }, { status: 400 });
    }

    const hasTrash = await ensureTrashColumns();
    if (!hasTrash) {
      return Response.json(
        { error: "Recycle bin is not available" },
        { status: 400 }
      );
    }

    const [result] = await pool.query(
      "UPDATE projects SET deleted_at = NULL, deleted_by = NULL WHERE id = ? AND deleted_at IS NOT NULL",
      [projectId]
    );
    const affectedRows = Number((result as { affectedRows?: number })?.affectedRows || 0);

    if (!affectedRows) {
      return Response.json(
        { error: "Project not found in recycle bin" },
        { status: 404 }
      );
    }

    return Response.json({ ok: true, restored: true });
  } catch (error) {
    if (isUnauthorizedError(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/projects/:id/restore] POST error:", error);
    return Response.json({ error: "Failed to restore project" }, { status: 500 });
  }
}
