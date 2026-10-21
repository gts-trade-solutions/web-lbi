/* eslint-disable @typescript-eslint/no-explicit-any */
// Public: photos for ONE report inside a shared project. Requires a valid
// share session AND that the report actually belongs to this link's project
// (so a session can't be used to read some other project's photos).
import pool from "../../../../../../../lib/db";
import { signRowUrls } from "../../../../../../../lib/s3";
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

    // Return photos in the user's saved order (sort_order), matching the
    // authored endpoint, so the shared route map shows the same first photo /
    // carousel order the user set. Falls back to created_at when unordered.
    let orderBy = "ORDER BY created_at ASC";
    try {
      const [pc] = await pool.query("SHOW COLUMNS FROM report_photos LIKE 'sort_order'");
      if (Array.isArray(pc) && pc.length) {
        orderBy = "ORDER BY (sort_order IS NULL), sort_order ASC, created_at ASC";
      }
    } catch {
      /* keep created_at ordering if the column check fails */
    }
    const [rows] = await pool.query(
      `SELECT * FROM report_photos WHERE report_id = ? ${orderBy}`,
      [reportId]
    );
    return Response.json({ photos: await signRowUrls("report_photos", Array.isArray(rows) ? rows : []) });
  } catch (error) {
    console.error("[api/share/:token/reports/:reportId/photos] error:", error);
    return Response.json({ error: "Failed to load photos" }, { status: 500 });
  }
}
