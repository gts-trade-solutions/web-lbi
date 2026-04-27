import pool from "../../../../lib/db";
import { requireAuth } from "../../../../lib/auth";

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

    const [rows] = await pool.query("SELECT * FROM projects WHERE id = ? LIMIT 1", [projectId]);
    const project = Array.isArray(rows) && rows.length ? rows[0] : null;

    if (!project) {
      return Response.json({ error: "Project not found" }, { status: 404 });
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

    if (columns.has("last_modified_by")) {
      updates.last_modified_by = authUser.id;
    }

    const updateKeys = Object.keys(updates);
    if (!updateKeys.length) {
      return Response.json({ error: "No valid fields to update" }, { status: 400 });
    }

    const setClause = updateKeys.map((k) => `${k} = ?`).join(", ");
    const values = updateKeys.map((k) => updates[k]);
    values.push(projectId);

    const [result] = await pool.query(`UPDATE projects SET ${setClause} WHERE id = ?`, values);
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
    requireAuth(request);
    const projectId = getProjectId(context);

    if (!projectId) {
      return Response.json({ error: "Project id is required" }, { status: 400 });
    }

    const [result] = await pool.query("DELETE FROM projects WHERE id = ?", [projectId]);
    const affectedRows = Number((result as any)?.affectedRows || 0);

    if (!affectedRows) {
      return Response.json({ error: "Project not found" }, { status: 404 });
    }

    return Response.json({ ok: true });
  } catch (error) {
    if (isUnauthorizedError(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/projects/:id] DELETE error:", error);
    return Response.json({ error: "Failed to delete project" }, { status: 500 });
  }
}
