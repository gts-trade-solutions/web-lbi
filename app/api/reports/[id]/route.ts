import pool from "../../../../lib/db";
import { requireAuth } from "../../../../lib/auth";

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

    const [result] = await pool.query("DELETE FROM reports WHERE id = ?", [reportId]);
    const affectedRows = Number((result as any)?.affectedRows || 0);
    if (!affectedRows) return Response.json({ error: "Report not found" }, { status: 404 });
    return Response.json({ ok: true });
  } catch (error) {
    if (unauthorized(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/reports/:id] DELETE error:", error);
    return Response.json({ error: "Failed to delete report" }, { status: 500 });
  }
}
