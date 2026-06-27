import { v4 as uuidv4 } from "uuid";
import pool from "../../../../lib/db";
import { requireAuth } from "../../../../lib/auth";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type DbColumnRow = { Field?: string };

function isUnauthorizedError(e: unknown) {
  return (e as { message?: string })?.message === "Unauthorized";
}

async function getColumns(table: string): Promise<Set<string>> {
  const [rows] = await pool.query(`SHOW COLUMNS FROM ${table}`);
  return new Set(
    (Array.isArray(rows) ? rows : []).map((r) =>
      String((r as DbColumnRow).Field || "").toLowerCase()
    )
  );
}

// Like getColumns but returns null when the table does not exist on this
// install (so optional features like GA setup don't break the whole copy).
async function safeGetColumns(table: string): Promise<Set<string> | null> {
  try {
    return await getColumns(table);
  } catch {
    return null;
  }
}

/**
 * Build an INSERT that copies a source row into the same table with a fresh
 * id and caller-supplied overrides. created_at / updated_at are emitted as
 * inline NOW() so MySQL writes its own datetime. Columns absent from `cols`
 * are skipped. Any column listed in `skip` is omitted entirely.
 */
function buildCopyInsert(
  table: string,
  cols: Set<string>,
  sourceRow: Record<string, unknown>,
  overrides: Record<string, unknown>,
  skip: Set<string> = new Set()
): { sql: string; values: unknown[] } {
  const keys: string[] = [];
  const placeholders: string[] = [];
  const values: unknown[] = [];

  for (const col of cols) {
    if (skip.has(col)) continue;

    if (col === "created_at" || col === "updated_at") {
      keys.push(col);
      placeholders.push("NOW()");
      continue;
    }

    if (Object.prototype.hasOwnProperty.call(overrides, col)) {
      keys.push(col);
      placeholders.push("?");
      values.push(overrides[col]);
      continue;
    }

    if (Object.prototype.hasOwnProperty.call(sourceRow, col)) {
      keys.push(col);
      placeholders.push("?");
      values.push(sourceRow[col]);
    }
  }

  const sql = `INSERT INTO ${table} (${keys.join(", ")}) VALUES (${placeholders.join(", ")})`;
  return { sql, values };
}

// POST /api/projects/from-reports
// Body: { name, sourceProjectId, reportIds: string[] }
// Creates a NEW project and copies the selected reports (with their photos
// and path points) into it. The source project is left untouched.
export async function POST(request: Request) {
  // Connection held for the duration of the copy transaction.
  let conn: Awaited<ReturnType<typeof pool.getConnection>> | null = null;
  try {
    const authUser = requireAuth(request);
    const body = await request.json().catch(() => ({} as any));

    const name = String(body?.name || "").trim();
    const sourceProjectId = String(body?.sourceProjectId || "").trim();
    const reportIds: string[] = Array.isArray(body?.reportIds)
      ? body.reportIds.map((x: unknown) => String(x || "").trim()).filter(Boolean)
      : [];
    // Copy options (default to copying everything).
    const includePhotos = body?.includePhotos !== false;
    const includeGaSetup = body?.includeGaSetup !== false;

    if (!name) return Response.json({ error: "Project name is required" }, { status: 400 });
    if (!sourceProjectId) {
      return Response.json({ error: "Source project id is required" }, { status: 400 });
    }
    if (!reportIds.length) {
      return Response.json({ error: "Select at least one report" }, { status: 400 });
    }

    const projectCols = await getColumns("projects");
    const reportCols = await getColumns("reports");
    const photoCols = await getColumns("report_photos");
    const pointCols = await getColumns("report_path_points");

    // Everything below runs in a single transaction so a mid-way failure
    // leaves NO partial project behind.
    conn = await pool.getConnection();
    await conn.beginTransaction();

    // ---- 1. Fetch the selected source reports (scoped to the source project).
    const placeholders = reportIds.map(() => "?").join(", ");
    const [reportRows] = await conn.query(
      `SELECT * FROM reports WHERE project_id = ? AND id IN (${placeholders})`,
      [sourceProjectId, ...reportIds]
    );
    const sourceReports = Array.isArray(reportRows)
      ? (reportRows as Record<string, unknown>[])
      : [];

    if (!sourceReports.length) {
      await conn.rollback();
      conn.release();
      conn = null;
      return Response.json(
        { error: "None of the selected reports were found in this project" },
        { status: 404 }
      );
    }

    // ---- 2. Create the new project.
    const newProjectId = uuidv4();
    const projectPayload: Record<string, unknown> = {
      id: newProjectId,
      name,
      created_by: authUser.id,
      last_modified_by: authUser.id,
      user_id: authUser.id,
    };
    const pKeys = Object.keys(projectPayload).filter((k) => projectCols.has(k));
    const pPlaceholders = pKeys.map(() => "?");
    const pValues = pKeys.map((k) => projectPayload[k]);
    await conn.query(
      `INSERT INTO projects (${pKeys.join(", ")}) VALUES (${pPlaceholders.join(", ")})`,
      pValues
    );

    // Preserve the order the caller selected them in.
    const orderIndex = new Map(reportIds.map((id, i) => [id, i]));
    sourceReports.sort(
      (a, b) =>
        (orderIndex.get(String(a.id)) ?? 0) - (orderIndex.get(String(b.id)) ?? 0)
    );

    let reportsCopied = 0;
    let photosCopied = 0;
    let pointsCopied = 0;

    // ---- 3. Copy each report + its photos + its path points.
    for (let i = 0; i < sourceReports.length; i += 1) {
      const src = sourceReports[i];
      const newReportId = uuidv4();

      const reportOverrides: Record<string, unknown> = {
        id: newReportId,
        project_id: newProjectId,
        created_by: authUser.id,
        // Detach from the source project's route; the new project has its own.
        route_id: null,
        // Re-sequence cleanly (10, 20, 30, ...) so insertion still works.
        sort_order: (i + 1) * 10,
      };
      if (reportCols.has("user_id")) reportOverrides.user_id = authUser.id;
      if (reportCols.has("last_modified_by")) reportOverrides.last_modified_by = authUser.id;

      const reportInsert = buildCopyInsert("reports", reportCols, src, reportOverrides);
      await conn.query(reportInsert.sql, reportInsert.values);
      reportsCopied += 1;

      // Photos
      const [photoRows] = includePhotos
        ? await conn.query("SELECT * FROM report_photos WHERE report_id = ?", [src.id])
        : [[] as Record<string, unknown>[]];
      for (const ph of Array.isArray(photoRows) ? (photoRows as Record<string, unknown>[]) : []) {
        const photoOverrides: Record<string, unknown> = {
          id: uuidv4(),
          report_id: newReportId,
        };
        if (photoCols.has("user_id")) photoOverrides.user_id = authUser.id;
        const photoInsert = buildCopyInsert(
          "report_photos",
          photoCols,
          ph,
          photoOverrides
        );
        await conn.query(photoInsert.sql, photoInsert.values);
        photosCopied += 1;
      }

      // Path points
      const [pointRows] = await conn.query(
        "SELECT * FROM report_path_points WHERE report_id = ?",
        [src.id]
      );
      for (const pt of Array.isArray(pointRows) ? (pointRows as Record<string, unknown>[]) : []) {
        const pointOverrides: Record<string, unknown> = {
          report_id: newReportId,
        };
        if (pointCols.has("user_id")) pointOverrides.user_id = authUser.id;
        // report_path_points.id is BIGINT AUTO_INCREMENT — never copy or
        // override it (a UUID string would truncate). Let MySQL assign it.
        const pointInsert = buildCopyInsert(
          "report_path_points",
          pointCols,
          pt,
          pointOverrides,
          new Set(["id"])
        );
        await conn.query(pointInsert.sql, pointInsert.values);
        pointsCopied += 1;
      }
    }

    // ---- 4. Optionally copy the project's GA drawing + route-survey setup.
    let gaSetupCopied = false;
    if (includeGaSetup) {
      try {
        // GA drawing (one row per project).
        const gaCols = await safeGetColumns("project_ga_drawings");
        if (gaCols) {
          const [gaRows] = await conn.query(
            "SELECT * FROM project_ga_drawings WHERE project_id = ?",
            [sourceProjectId]
          );
          for (const ga of Array.isArray(gaRows) ? (gaRows as Record<string, unknown>[]) : []) {
            const overrides: Record<string, unknown> = {
              id: uuidv4(),
              project_id: newProjectId,
            };
            if (gaCols.has("created_by")) overrides.created_by = authUser.id;
            const ins = buildCopyInsert("project_ga_drawings", gaCols, ga, overrides);
            await conn.query(ins.sql, ins.values);
            gaSetupCopied = true;
          }
        }

        // Route-survey pages, plus their locations and images.
        const pageCols = await safeGetColumns("project_route_pages");
        const locCols = await safeGetColumns("project_route_page_locations");
        const imgCols = await safeGetColumns("project_route_page_images");
        if (pageCols) {
          const [pageRows] = await conn.query(
            "SELECT * FROM project_route_pages WHERE project_id = ?",
            [sourceProjectId]
          );
          for (const page of Array.isArray(pageRows) ? (pageRows as Record<string, unknown>[]) : []) {
            const newPageId = uuidv4();
            const pageOverrides: Record<string, unknown> = {
              id: newPageId,
              project_id: newProjectId,
            };
            if (pageCols.has("user_id")) pageOverrides.user_id = authUser.id;
            const pageIns = buildCopyInsert("project_route_pages", pageCols, page, pageOverrides);
            await conn.query(pageIns.sql, pageIns.values);
            gaSetupCopied = true;

            if (locCols) {
              const [locRows] = await conn.query(
                "SELECT * FROM project_route_page_locations WHERE project_page_id = ?",
                [page.id]
              );
              for (const loc of Array.isArray(locRows) ? (locRows as Record<string, unknown>[]) : []) {
                const locOverrides: Record<string, unknown> = {
                  id: uuidv4(),
                  project_page_id: newPageId,
                  project_id: newProjectId,
                };
                if (locCols.has("user_id")) locOverrides.user_id = authUser.id;
                const locIns = buildCopyInsert(
                  "project_route_page_locations",
                  locCols,
                  loc,
                  locOverrides
                );
                await conn.query(locIns.sql, locIns.values);
              }
            }

            if (imgCols) {
              const [imgRows] = await conn.query(
                "SELECT * FROM project_route_page_images WHERE project_page_id = ?",
                [page.id]
              );
              for (const img of Array.isArray(imgRows) ? (imgRows as Record<string, unknown>[]) : []) {
                const imgOverrides: Record<string, unknown> = {
                  id: uuidv4(),
                  project_page_id: newPageId,
                  project_id: newProjectId,
                };
                if (imgCols.has("user_id")) imgOverrides.user_id = authUser.id;
                const imgIns = buildCopyInsert(
                  "project_route_page_images",
                  imgCols,
                  img,
                  imgOverrides
                );
                await conn.query(imgIns.sql, imgIns.values);
              }
            }
          }
        }
      } catch (gaErr) {
        // GA setup is best-effort; never fail the whole copy because of it.
        console.error("[api/projects/from-reports] GA setup copy failed:", gaErr);
      }
    }

    // ---- 5. Commit and return the new project.
    const [newRows] = await conn.query("SELECT * FROM projects WHERE id = ? LIMIT 1", [
      newProjectId,
    ]);
    const project = Array.isArray(newRows) && newRows.length ? newRows[0] : null;

    await conn.commit();
    conn.release();
    conn = null;

    return Response.json(
      {
        ok: true,
        project,
        reportsCopied,
        photosCopied,
        pointsCopied,
        gaSetupCopied,
      },
      { status: 201 }
    );
  } catch (error) {
    // Roll back any partial copy so no half-built project is left behind.
    if (conn) {
      try {
        await conn.rollback();
      } catch (rbErr) {
        console.error("[api/projects/from-reports] rollback failed:", rbErr);
      }
      try {
        conn.release();
      } catch {
        // ignore
      }
      conn = null;
    }

    if (isUnauthorizedError(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    const e = error as { message?: string; sqlMessage?: string; code?: string };
    console.error("[api/projects/from-reports] POST error:", error);
    return Response.json(
      {
        error: "Failed to create project from reports",
        detail: e?.sqlMessage || e?.message || String(error),
        code: e?.code || null,
      },
      { status: 500 }
    );
  }
}
