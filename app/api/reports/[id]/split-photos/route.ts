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
 * POST /api/reports/:id/split-photos
 *
 * Splits a report that bundles several images (e.g. a "Route & Vehicle
 * Drawings" report holding 4 drawings) into ONE report PER photo. The source
 * report keeps its FIRST photo; every other photo is MOVED onto a brand-new
 * report inserted right after it, in order. Each new report keeps the SAME
 * category AND description as the source (per the user's request), plus its
 * point_key / coordinates, so the drawings become separate observations that
 * still read identically.
 *
 * Nothing is duplicated (photos are re-parented, not copied) and nothing is
 * deleted. The new reports are created by COPYING the source row's columns so
 * every NOT NULL column already has a valid value; only id and sort_order are
 * overridden.
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

    // 2) Load this report's photos in the SAME order the grid shows them.
    const [photoRows] = await pool.query(
      "SELECT id FROM report_photos WHERE report_id = ? ORDER BY created_at ASC, id ASC",
      [sourceId]
    );
    const photoIds = (Array.isArray(photoRows) ? (photoRows as Record<string, unknown>[]) : [])
      .map((p) => String(p.id || "").trim())
      .filter(Boolean);
    if (photoIds.length < 2) {
      return Response.json(
        { error: "This report has fewer than 2 photos — nothing to split." },
        { status: 400 }
      );
    }

    // The source keeps its first photo; the rest each get their own report.
    const toMove = photoIds.slice(1);
    const nNew = toMove.length;

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

    // 4) Make room after the source for nNew reports; shift the rest down.
    const hasSortOrder = allCols.includes("sort_order");
    let srcSort = Number(source.sort_order);
    if (hasSortOrder && Number.isFinite(srcSort)) {
      try {
        await pool.query(
          "UPDATE reports SET sort_order = sort_order + ? WHERE project_id = ? AND sort_order > ?",
          [nNew, projectId, srcSort]
        );
      } catch (shiftErr) {
        console.warn("[split-photos] sort_order shift failed:", shiftErr);
      }
    } else {
      srcSort = NaN;
    }

    // 5) Create one new report per moved photo and re-parent that photo onto it.
    const createdReportIds: string[] = [];
    for (let i = 0; i < toMove.length; i++) {
      const newId = uuidv4();
      const newSort =
        hasSortOrder && Number.isFinite(srcSort) ? srcSort + (i + 1) : undefined;

      const insertCols: string[] = [];
      const placeholders: string[] = [];
      const values: unknown[] = [];
      for (const name of allCols) {
        if (skipCols.has(name)) continue;
        if (nowCols.has(name)) {
          insertCols.push(name);
          placeholders.push("NOW()");
          continue;
        }
        let val: unknown;
        if (name === "id") val = newId;
        else if (name === "sort_order" && typeof newSort !== "undefined") val = newSort;
        // category & description are copied AS-IS (same category + description).
        else val = typeof source[name] === "undefined" ? null : source[name];
        insertCols.push(name);
        placeholders.push("?");
        values.push(val);
      }
      await pool.query(
        `INSERT INTO reports (${insertCols.join(", ")}) VALUES (${placeholders.join(", ")})`,
        values
      );

      await pool.query("UPDATE report_photos SET report_id = ? WHERE id = ?", [newId, toMove[i]]);
      createdReportIds.push(newId);
    }

    console.log("[split-photos] done", {
      sourceId,
      projectId,
      keptOnSource: 1,
      newReports: createdReportIds.length,
    });

    return Response.json({
      ok: true,
      sourceId,
      createdReportIds,
      newReports: createdReportIds.length,
      keptOnSource: 1,
    });
  } catch (error) {
    if (unauthorized(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/reports/:id/split-photos] error:", error);
    return Response.json({ error: "Failed to split photos into separate reports" }, { status: 500 });
  }
}
