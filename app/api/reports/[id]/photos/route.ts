import { v4 as uuidv4 } from "uuid";
import pool from "../../../../../lib/db";
import { requireAuth } from "../../../../../lib/auth";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type DbColumnRow = { Field?: string };
type Ctx = { params: { id: string } };

async function getColumns() {
  const [rows] = await pool.query("SHOW COLUMNS FROM report_photos");
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

    const [rows] = await pool.query(
      "SELECT * FROM report_photos WHERE report_id = ? ORDER BY created_at ASC",
      [reportId]
    );
    return Response.json({ photos: Array.isArray(rows) ? rows : [] });
  } catch (error) {
    if (unauthorized(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/reports/:id/photos] GET error:", error);
    return Response.json({ error: "Failed to fetch report photos" }, { status: 500 });
  }
}

export async function POST(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const reportId = String(context.params?.id || "").trim();
    if (!reportId) return Response.json({ error: "Report id is required" }, { status: 400 });

    console.log("[save report photo] reportId:", reportId);

    // Validate that the parent report row actually exists. Without this guard
    // a buggy client could insert orphan photos with a bogus report_id that
    // never joins against reports during export.
    const [reportRows] = await pool.query("SELECT id FROM reports WHERE id = ? LIMIT 1", [
      reportId,
    ]);
    if (!Array.isArray(reportRows) || reportRows.length === 0) {
      console.warn("[save report photo] report not found:", reportId);
      return Response.json(
        { error: "Report not found", reportId },
        { status: 404 }
      );
    }

    const body = await request.json().catch(() => ({} as any));
    const list = Array.isArray(body?.photos) ? body.photos : body ? [body] : [];
    const rowsIn = list.filter((x: any) => x && typeof x === "object" && String(x.url || "").trim());
    if (!rowsIn.length) {
      return Response.json({ error: "At least one photo url is required" }, { status: 400 });
    }

    const cols = await getColumns();
    const prepared = rowsIn.map((p: any) => ({
      id: p?.id || uuidv4(),
      report_id: reportId,
      url: String(p.url || "").trim(),
      file_name:
        p?.file_name ??
        p?.fileName ??
        (typeof p?.key === "string" ? p.key.split("/").pop() : null) ??
        null,
      width: p?.width ?? null,
      height: p?.height ?? null,
      bucket: p?.bucket ?? null,
      path: p?.path ?? p?.key ?? null,
    }));

    for (const row of prepared) {
      console.log("[save report photo] url:", row.url);
    }

    const keys = Object.keys(prepared[0]).filter((k) => cols.has(k));
    const valuesSql = prepared.map(() => `(${keys.map(() => "?").join(", ")})`).join(", ");
    const args = prepared.flatMap((r: Record<string, unknown>) =>
      keys.map((k) => r[k])
    );

    await pool.query(`INSERT INTO report_photos (${keys.join(", ")}) VALUES ${valuesSql}`, args);
    console.log(
      "[save report photo] inserted into report_photos count=",
      prepared.length,
      "reportId=",
      reportId
    );

    const ids = prepared.map((p: { id: string }) => p.id);
    const [rows] = await pool.query(
      `SELECT * FROM report_photos WHERE id IN (${ids.map(() => "?").join(", ")})`,
      ids
    );

    return Response.json({ photos: Array.isArray(rows) ? rows : [] }, { status: 201 });
  } catch (error) {
    if (unauthorized(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/reports/:id/photos] POST error:", error);
    return Response.json({ error: "Failed to save report photos" }, { status: 500 });
  }
}
