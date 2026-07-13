import pool from "../../../../lib/db";
import { requireAuth } from "../../../../lib/auth";
import { logActivity } from "../../../../lib/activityLog";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type DbColumnRow = { Field?: string };
type Ctx = { params: { id: string } };

async function getColumns() {
  const [rows] = await pool.query("SHOW COLUMNS FROM reports");
  return new Set(
    (Array.isArray(rows) ? rows : []).map((r) =>
      String((r as DbColumnRow).Field || "").toLowerCase()
    )
  );
}

function unauthorized(error: unknown) {
  return (error as { message?: string })?.message === "Unauthorized";
}

export async function GET(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const reportId = String(context.params?.id || "").trim();
    if (!reportId) return Response.json({ error: "Report id is required" }, { status: 400 });

    const [rows] = await pool.query("SELECT * FROM reports WHERE id = ? LIMIT 1", [reportId]);
    const report = Array.isArray(rows) && rows.length ? rows[0] : null;
    if (!report) return Response.json({ error: "Report not found" }, { status: 404 });
    return Response.json({ report });
  } catch (error) {
    if (unauthorized(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/reports/:id] GET error:", error);
    return Response.json({ error: "Failed to fetch report" }, { status: 500 });
  }
}

export async function PUT(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const reportId = String(context.params?.id || "").trim();
    if (!reportId) return Response.json({ error: "Report id is required" }, { status: 400 });

    const body = await request.json().catch(() => ({} as any));
    const cols = await getColumns();
    const allowed = [
      "route_id",
      "category",
      "description",
      "remarks_action",
      "difficulty",
      "sort_order",
      "status",
      "point_key",
      "latitude",
      "longitude",
      "loc_lat",
      "loc_lon",
    ];
    const updates = allowed.filter((k) => cols.has(k) && Object.prototype.hasOwnProperty.call(body, k));
    if (!updates.length) {
      return Response.json({ error: "No valid fields to update" }, { status: 400 });
    }

    const setSql = updates.map((k) => `${k} = ?`).join(", ");
    const args = [...updates.map((k) => body[k]), reportId];
    const [result] = await pool.query(`UPDATE reports SET ${setSql} WHERE id = ?`, args);
    const affectedRows = Number((result as any)?.affectedRows || 0);
    if (!affectedRows) return Response.json({ error: "Report not found" }, { status: 404 });

    const [rows] = await pool.query("SELECT * FROM reports WHERE id = ? LIMIT 1", [reportId]);
    const report = Array.isArray(rows) && rows.length ? rows[0] : null;

    // AUDIT: report edited from the web UI (category / difficulty / etc.).
    await logActivity(request, {
      action: "update",
      table: "reports",
      entityId: reportId,
      projectId: (report as any)?.project_id ? String((report as any).project_id) : null,
      rowCount: affectedRows,
      details: { changedFields: updates },
    });

    return Response.json({ report });
  } catch (error) {
    if (unauthorized(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/reports/:id] PUT error:", error);
    return Response.json({ error: "Failed to update report" }, { status: 500 });
  }
}

export async function DELETE(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const reportId = String(context.params?.id || "").trim();
    if (!reportId) return Response.json({ error: "Report id is required" }, { status: 400 });

    // Capture the report BEFORE deletion so the audit log records what was removed.
    let preRow: any = null;
    try {
      const [pre] = await pool.query("SELECT * FROM reports WHERE id = ? LIMIT 1", [reportId]);
      preRow = Array.isArray(pre) && pre.length ? pre[0] : null;
    } catch {
      /* best-effort pre-image */
    }

    const [result] = await pool.query("DELETE FROM reports WHERE id = ?", [reportId]);
    const affectedRows = Number((result as any)?.affectedRows || 0);
    if (!affectedRows) return Response.json({ error: "Report not found" }, { status: 404 });

    // AUDIT: report deleted from the web UI.
    await logActivity(request, {
      action: "delete",
      table: "reports",
      entityId: reportId,
      projectId: preRow?.project_id ? String(preRow.project_id) : null,
      rowCount: affectedRows,
      details: preRow
        ? { category: preRow.category, point_key: preRow.point_key, description: String(preRow.description || "").slice(0, 200) }
        : null,
    });

    return Response.json({ ok: true });
  } catch (error) {
    if (unauthorized(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/reports/:id] DELETE error:", error);
    return Response.json({ error: "Failed to delete report" }, { status: 500 });
  }
}
