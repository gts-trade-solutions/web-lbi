import pool from "../../../../lib/db";
import { requireAuth } from "../../../../lib/auth";
import { ensureTrashColumns, ensureStatusColumn } from "../../../../lib/projects-trash";
import { logActivity } from "../../../../lib/activityLog";
import { signS3Url } from "../../../../lib/s3";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type DbColumnRow = { Field?: string };

type RouteContext = {
  params: {
    id: string;
  };
};

async function getProjectColumns() {
  const [rows] = await pool.query("SHOW COLUMNS FROM projects");
  const cols = new Set(
    (Array.isArray(rows) ? rows : []).map((r) =>
      String((r as DbColumnRow).Field || "").toLowerCase()
    )
  );
  return cols;
}

// Lazily add the optional client-logo column so a project can carry a client
// logo (shown top-right of the exported report). Safe to call repeatedly.
async function ensureClientLogoColumn() {
  try {
    const [rows] = await pool.query("SHOW COLUMNS FROM projects LIKE 'client_logo_url'");
    if (!Array.isArray(rows) || rows.length === 0) {
      await pool.query("ALTER TABLE projects ADD COLUMN client_logo_url TEXT NULL");
    }
  } catch (e) {
    console.error("[api/projects/:id] ensureClientLogoColumn failed:", e);
  }
}

function isUnauthorizedError(error: unknown) {
  return (error as { message?: string })?.message === "Unauthorized";
}

function getProjectId(context: RouteContext) {
  return String(context.params?.id || "").trim();
}

export async function GET(request: Request, context: RouteContext) {
  try {
    requireAuth(request);
    const projectId = getProjectId(context);

    if (!projectId) {
      return Response.json({ error: "Project id is required" }, { status: 400 });
    }

    const hasTrash = await ensureTrashColumns();
    const notDeleted = hasTrash ? " AND deleted_at IS NULL" : "";

    const [rows] = await pool.query(
      `SELECT * FROM projects WHERE id = ?${notDeleted} LIMIT 1`,
      [projectId]
    );
    const project = Array.isArray(rows) && rows.length ? rows[0] : null;

    if (!project) {
      return Response.json({ error: "Project not found" }, { status: 404 });
    }

    // The client logo lives in the private bucket; sign it so the browser can
    // display it. The DB keeps the canonical URL (what the export reads).
    const clientLogo = (project as { client_logo_url?: string | null }).client_logo_url;
    if (clientLogo) {
      try {
        (project as { client_logo_url?: string | null }).client_logo_url = await signS3Url(clientLogo);
      } catch {
        /* fall back to the stored URL */
      }
    }

    return Response.json({ project });
  } catch (error) {
    if (isUnauthorizedError(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/projects/:id] GET error:", error);
    return Response.json({ error: "Failed to fetch project" }, { status: 500 });
  }
}

export async function PUT(request: Request, context: RouteContext) {
  try {
    const authUser = requireAuth(request);
    const projectId = getProjectId(context);

    if (!projectId) {
      return Response.json({ error: "Project id is required" }, { status: 400 });
    }

    const body = await request.json().catch(() => ({} as any));
    // Make sure the status column exists before an "mark finished" update.
    if (body?.status !== undefined) await ensureStatusColumn();
    // Make sure the client-logo column exists before setting / clearing it.
    if (body?.client_logo_url !== undefined) await ensureClientLogoColumn();
    const columns = await getProjectColumns();

    const updates: Record<string, unknown> = {};

    if (body?.name !== undefined && columns.has("name")) {
      const name = String(body?.name || "").trim();
      if (!name) {
        return Response.json({ error: "Project name is required" }, { status: 400 });
      }
      updates.name = name;
    }

    if (body?.description !== undefined && columns.has("description")) {
      const description = String(body?.description || "").trim();
      updates.description = description || null;
    }

    if (body?.client_name !== undefined && columns.has("client_name")) {
      const clientName = String(body?.client_name || "").trim();
      updates.client_name = clientName || null;
    }

    if (body?.location !== undefined && columns.has("location")) {
      const location = String(body?.location || "").trim();
      updates.location = location || null;
    }

    if (body?.status !== undefined && columns.has("status")) {
      const status = String(body?.status || "").trim();
      updates.status = status || "active";
    }

    if (body?.client_logo_url !== undefined && columns.has("client_logo_url")) {
      const clientLogoUrl = String(body?.client_logo_url || "").trim();
      updates.client_logo_url = clientLogoUrl || null; // empty string clears it
    }

    if (columns.has("last_modified_by")) {
      updates.last_modified_by = authUser.id;
    }

    const updateKeys = Object.keys(updates);
    if (!updateKeys.length) {
      return Response.json({ error: "No valid fields to update" }, { status: 400 });
    }

    const hasTrash = await ensureTrashColumns();
    const notDeleted = hasTrash ? " AND deleted_at IS NULL" : "";

    const setClause = updateKeys.map((k) => `${k} = ?`).join(", ");
    const values = updateKeys.map((k) => updates[k]);
    values.push(projectId);

    const [result] = await pool.query(
      `UPDATE projects SET ${setClause} WHERE id = ?${notDeleted}`,
      values
    );
    const affectedRows = Number((result as any)?.affectedRows || 0);

    if (!affectedRows) {
      return Response.json({ error: "Project not found" }, { status: 404 });
    }

    const [rows] = await pool.query("SELECT * FROM projects WHERE id = ? LIMIT 1", [projectId]);
    const project = Array.isArray(rows) && rows.length ? rows[0] : null;

    return Response.json({ project });
  } catch (error) {
    if (isUnauthorizedError(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/projects/:id] PUT error:", error);
    return Response.json({ error: "Failed to update project" }, { status: 500 });
  }
}

export async function DELETE(request: Request, context: RouteContext) {
  try {
    const authUser = requireAuth(request);
    const projectId = getProjectId(context);

    if (!projectId) {
      return Response.json({ error: "Project id is required" }, { status: 400 });
    }

    // `?permanent=true` skips the recycle bin and deletes immediately
    // (used by "Delete forever" inside the bin). Otherwise we soft-delete.
    const url = new URL(request.url);
    const permanent = url.searchParams.get("permanent") === "true";

    const hasTrash = await ensureTrashColumns();

    if (permanent || !hasTrash) {
      const [result] = await pool.query("DELETE FROM projects WHERE id = ?", [projectId]);
      const affectedRows = Number((result as any)?.affectedRows || 0);

      if (!affectedRows) {
        return Response.json({ error: "Project not found" }, { status: 404 });
      }

      // AUDIT: project permanently deleted.
      await logActivity(request, {
        action: "delete",
        table: "projects",
        entityId: projectId,
        projectId,
        rowCount: affectedRows,
        details: { permanent: true },
      });

      return Response.json({ ok: true, permanent: true });
    }

    // Soft delete: move to the recycle bin. Only affects projects that are
    // not already binned.
    const [result] = await pool.query(
      "UPDATE projects SET deleted_at = NOW(), deleted_by = ? WHERE id = ? AND deleted_at IS NULL",
      [authUser.id, projectId]
    );
    const affectedRows = Number((result as any)?.affectedRows || 0);

    if (!affectedRows) {
      return Response.json({ error: "Project not found" }, { status: 404 });
    }

    // AUDIT: project moved to the recycle bin (soft delete).
    await logActivity(request, {
      action: "delete",
      table: "projects",
      entityId: projectId,
      projectId,
      rowCount: affectedRows,
      details: { recycled: true },
    });

    return Response.json({ ok: true, recycled: true });
  } catch (error) {
    if (isUnauthorizedError(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/projects/:id] DELETE error:", error);
    return Response.json({ error: "Failed to delete project" }, { status: 500 });
  }
}
