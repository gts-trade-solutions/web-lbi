/* eslint-disable @typescript-eslint/no-explicit-any */
// Public: photos for ONE report inside a shared project. Requires a valid
// share session AND that the report actually belongs to this link's project
// (so a session can't be used to read some other project's photos).
import pool from "../../../../../../../lib/db";
import {
  getShareByToken,
  parseSelectedIds,
  readShareSession,
  verifyShareSession,
} from "../../../../../../../lib/shareLink";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type Ctx = { params: { token: string; reportId: string } };

export async function GET(request: Request, context: Ctx) {
  try {
    const token = String(context.params?.token || "").trim();
    const reportId = String(context.params?.reportId || "").trim();
    if (!token || !reportId) return Response.json({ error: "Invalid request" }, { status: 400 });

    const session = readShareSession(request);
    if (!verifyShareSession(session, token)) {
      return Response.json({ error: "Session expired" }, { status: 401 });
    }
    const share = await getShareByToken(token);
    if (!share) return Response.json({ error: "This link is no longer active." }, { status: 404 });

    // The report must belong to this share's project (and its selection).
    const [reportRows] = await pool.query(
      "SELECT id, project_id FROM reports WHERE id = ? LIMIT 1",
      [reportId]
    );
    const report = Array.isArray(reportRows) && reportRows.length ? (reportRows[0] as any) : null;
    if (!report || String(report.project_id) !== String(share.project_id)) {
      return Response.json({ photos: [] });
    }
    const selected = parseSelectedIds(share.selected_ids);
    if (selected.length && !selected.includes(String(reportId))) {
      return Response.json({ photos: [] });
    }

    const [rows] = await pool.query(
      "SELECT * FROM report_photos WHERE report_id = ? ORDER BY created_at ASC",
      [reportId]
    );
    return Response.json({ photos: Array.isArray(rows) ? rows : [] });
  } catch (error) {
    console.error("[api/share/:token/reports/:reportId/photos] error:", error);
    return Response.json({ error: "Failed to load photos" }, { status: 500 });
  }
}
