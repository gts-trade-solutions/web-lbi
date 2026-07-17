/* eslint-disable @typescript-eslint/no-explicit-any */
// Public: the reports for a shared project, in map order. Requires a valid
// share session (minted by /api/share/<token>/auth). Returns the SAME shape
// as /api/projects/<id>/reports so the animated map can consume it unchanged.
import pool from "../../../../../lib/db";
import {
  getShareByToken,
  parseSelectedIds,
  readShareSession,
  verifyShareSession,
} from "../../../../../lib/shareLink";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type Ctx = { params: { token: string } };

async function getColumns(table: string): Promise<Set<string>> {
  const [rows] = await pool.query(`SHOW COLUMNS FROM ${table}`);
  return new Set(
    (Array.isArray(rows) ? rows : []).map((r) => String((r as any)?.Field || "").toLowerCase())
  );
}

export async function GET(request: Request, context: Ctx) {
  try {
    const token = String(context.params?.token || "").trim();
    if (!token) return Response.json({ error: "Invalid link" }, { status: 400 });

    const session = readShareSession(request);
    if (!verifyShareSession(session, token)) {
      return Response.json({ error: "Session expired" }, { status: 401 });
    }
    const share = await getShareByToken(token);
    if (!share) return Response.json({ error: "This link is no longer active." }, { status: 404 });

    const cols = await getColumns("reports");
    const orderParts: string[] = [];
    if (cols.has("sort_order")) orderParts.push("sort_order ASC");
    if (cols.has("created_at")) orderParts.push("created_at ASC");
    if (cols.has("id")) orderParts.push("id ASC");
    const orderBy = orderParts.length ? `ORDER BY ${orderParts.join(", ")}` : "";
    const delFilter = cols.has("deleted_at") ? "AND deleted_at IS NULL" : "";

    const [rows] = await pool.query(
      `SELECT * FROM reports WHERE project_id = ? ${delFilter} ${orderBy}`,
      [share.project_id]
    );
    let reports = Array.isArray(rows) ? (rows as any[]) : [];

    // If the link was created for a specific selection, serve only those.
    const selected = parseSelectedIds(share.selected_ids);
    if (selected.length) {
      const set = new Set(selected);
      reports = reports.filter((r) => set.has(String(r.id)));
    }

    return Response.json({ reports });
  } catch (error) {
    console.error("[api/share/:token/reports] error:", error);
    return Response.json({ error: "Failed to load reports" }, { status: 500 });
  }
}
