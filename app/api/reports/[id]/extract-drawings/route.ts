import { v4 as uuidv4 } from "uuid";
import pool from "../../../../../lib/db";
import { requireAuth } from "../../../../../lib/auth";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type Ctx = { params: { id: string } };
type DbColumnRow = { Field?: string; Extra?: string; Type?: string };

function unauthorized(error: unknown) {
  return (error as { message?: string })?.message === "Unauthorized";
}

/**
 * POST /api/reports/:id/extract-drawings
 *
 * Pulls the annotation DRAWINGS (images made with the "Draw on photo" tool,
 * saved as file_name "annotated_<ts>.jpg") off this report and MOVES them onto
 * a brand-new report inserted right after it. The source report keeps its other
 * images (photos / route map / stage-summary table). No image is duplicated —
 * the drawing rows are re-parented, so each drawing becomes its own separate
 * report/observation in the grid and in the Word export.
 *
 * The new report is created by COPYING the source row's columns (so every
 * NOT NULL column already has a valid value), overriding only id / sort_order /
 * description, and letting datetime columns take NOW().
 */
export async function POST(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const sourceId = String(context.params?.id || "").trim();
    if (!sourceId) return Response.json({ error: "Report id is required" }, { status: 400 });

    // 1) Load the source report.
    const [srcRows] = await pool.query("SELECT * FROM reports WHERE id = ? LIMIT 1", [sourceId]);
    const source =
      Array.isArray(srcRows) && srcRows.length ? (srcRows[0] as Record<string, unknown>) : null;
    if (!source) return Response.json({ error: "Report not found" }, { status: 404 });
    const projectId = String(source.project_id || "");

    // 2) Find the annotation-drawing photos on this report. The "Draw on photo"
    //    tool uploads them as annotated_<timestamp>.jpg (file_name), and the S3
    //    key/url carries the same token — match either so both paths are caught.
    const [drawRows] = await pool.query(
      `SELECT id FROM report_photos
       WHERE report_id = ?
         AND (LOWER(file_name) LIKE 'annotated%' OR LOWER(url) LIKE '%annotated%')`,
      [sourceId]
    );
    const drawingIds = (Array.isArray(drawRows) ? (drawRows as Record<string, unknown>[]) : [])
      .map((d) => String(d.id || "").trim())
      .filter(Boolean);
    if (!drawingIds.length) {
      return Response.json(
        { error: "No annotation drawings found on this report." },
        { status: 400 }
      );
    }

    // 3) Column metadata so the source row can be copied safely.
    const [colRows] = await pool.query("SHOW COLUMNS FROM reports");
    const colMeta = (Array.isArray(colRows) ? colRows : []) as DbColumnRow[];
    const allCols: string[] = [];
    const nowCols = new Set<string>(); // datetime/timestamp/date -> NOW()
    const skipCols = new Set<string>(); // auto/generated -> let the DB fill
    for (const c of colMeta) {
      const name = String(c.Field || "");
      if (!name) continue;
      allCols.push(name);
      const type = String(c.Type || "").toLowerCase();
      const extra = String(c.Extra || "").toLowerCase();
      if (
        extra.includes("auto_increment") ||
        extra.includes("default_generated") ||
        extra.includes("virtual generated") ||
        extra.includes("stored generated")
      ) {
        skipCols.add(name);
      }
      if (
        type.startsWith("datetime") ||
        type.startsWith("timestamp") ||
        type.startsWith("date")
      ) {
        nowCols.add(name);
      }
    }

    const newId = uuidv4();

    // 4) Place the new report right after the source; shift the rest down.
    let newSortOrder: number | null = null;
    if (allCols.includes("sort_order")) {
      const srcSort = Number(source.sort_order);
      if (Number.isFinite(srcSort)) {
        try {
          await pool.query(
            "UPDATE reports SET sort_order = sort_order + 1 WHERE project_id = ? AND sort_order > ?",
            [projectId, srcSort]
          );
        } catch (shiftErr) {
          console.warn("[extract-drawings] sort_order shift failed:", shiftErr);
        }
        newSortOrder = srcSort + 1;
      }
    }

    // 5) Build the INSERT by copying the source row.
    const insertCols: string[] = [];
    const placeholders: string[] = [];
    const values: unknown[] = [];
    for (const name of allCols) {
      if (skipCols.has(name)) continue; // let the DB default it (e.g. auto ts)
      if (nowCols.has(name)) {
        insertCols.push(name);
        placeholders.push("NOW()");
        continue;
      }
      let val: unknown;
      if (name === "id") val = newId;
      else if (name === "sort_order") val = newSortOrder ?? (source.sort_order ?? null);
      else if (name === "description") val = ""; // the new report is drawings-only
      else val = typeof source[name] === "undefined" ? null : source[name];
      insertCols.push(name);
      placeholders.push("?");
      values.push(val);
    }
    await pool.query(
      `INSERT INTO reports (${insertCols.join(", ")}) VALUES (${placeholders.join(", ")})`,
      values
    );

    // 6) MOVE the drawing photos onto the new report (re-parent, no copy).
    const ph = drawingIds.map(() => "?").join(",");
    const [mv] = await pool.query(
      `UPDATE report_photos SET report_id = ? WHERE id IN (${ph})`,
      [newId, ...drawingIds]
    );
    const moved = Number((mv as { affectedRows?: number })?.affectedRows || 0);

    // 7) Return the new report row so the client can refresh in place.
    const [newRows] = await pool.query("SELECT * FROM reports WHERE id = ? LIMIT 1", [newId]);
    const newReport =
      Array.isArray(newRows) && newRows.length ? (newRows[0] as Record<string, unknown>) : null;

    console.log("[extract-drawings] done", {
      sourceId,
      newId,
      projectId,
      movedPhotos: moved,
    });

    return Response.json({ ok: true, report: newReport, movedPhotos: moved, sourceId });
  } catch (error) {
    if (unauthorized(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/reports/:id/extract-drawings] error:", error);
    return Response.json({ error: "Failed to extract drawings" }, { status: 500 });
  }
}
