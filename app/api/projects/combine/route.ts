/* eslint-disable @typescript-eslint/no-explicit-any */
// Combine 2+ existing projects into ONE new project. Non-destructive: the
// source projects are left untouched; a brand-new project is created that
// contains COPIES of every source report (and its photos), in the order the
// projects are listed, renumbered as a single continuous sequence
// (point_key 1..N, sort_order 10,20,...). Photo rows reference the same S3
// URLs, so nothing is re-uploaded. Runs in a transaction so a failure can't
// leave a half-built project.
import { v4 as uuidv4 } from "uuid";
import type { PoolConnection } from "mysql2/promise";
import pool from "../../../../lib/db";
import { requireAuth } from "../../../../lib/auth";
import { logActivity } from "../../../../lib/activityLog";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 300;

async function columnsOf(conn: PoolConnection, table: string): Promise<Set<string>> {
  const [rows] = await conn.query(`SHOW COLUMNS FROM ${table}`);
  return new Set(
    (Array.isArray(rows) ? rows : []).map((r) => String((r as any)?.Field || "").toLowerCase())
  );
}

// Insert only the columns that actually exist on this install.
async function insertRow(
  conn: PoolConnection,
  table: string,
  payload: Record<string, unknown>,
  cols: Set<string>
) {
  const keys = Object.keys(payload).filter((k) => cols.has(k.toLowerCase()));
  if (!keys.length) return;
  const placeholders = keys.map(() => "?").join(", ");
  await conn.query(
    `INSERT INTO ${table} (${keys.join(", ")}) VALUES (${placeholders})`,
    keys.map((k) => payload[k])
  );
}

// Build a copy payload from a source row: keep every real data column, drop the
// ones we set ourselves (identity, ordering, timestamps, ownership).
function copyPayload(row: Record<string, any>, drop: Set<string>, cols: Set<string>) {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(row)) {
    const lk = k.toLowerCase();
    if (drop.has(lk)) continue;
    if (cols.has(lk)) out[k] = v;
  }
  return out;
}

// Straight-line distance between two GPS points, in kilometres. Used to work
// out how far each point sits ALONG its project when the reports carry no KM of
// their own, so the combined chainage can still be continuous.
function haversineKm(lat1: number, lon1: number, lat2: number, lon2: number): number {
  const R = 6371;
  const toRad = (d: number) => (d * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
  return 2 * R * Math.asin(Math.min(1, Math.sqrt(a)));
}

const REPORT_DROP = new Set([
  "id",
  "project_id",
  "point_key",
  "sort_order",
  "created_at",
  "updated_at",
  "deleted_at",
  "deleted_by",
  "route_id", // a source route wouldn't belong to the combined project
  "user_id",
  "created_by",
]);
const PHOTO_DROP = new Set(["id", "report_id", "created_at", "point_key", "user_id"]);

export async function POST(request: Request) {
  let authUser: { id: string };
  try {
    authUser = requireAuth(request) as { id: string };
  } catch {
    return Response.json({ error: "Unauthorized" }, { status: 401 });
  }

  let body: any = {};
  try {
    body = (await request.json()) || {};
  } catch {
    body = {};
  }
  const sourceProjectIds: string[] = Array.isArray(body.sourceProjectIds)
    ? body.sourceProjectIds.map((v: unknown) => String(v || "").trim()).filter(Boolean)
    : [];
  const name = String(body.name || "").trim();

  // Per-project starting KM entered in the combine dialog (projectId -> km).
  // A project with no entry continues from where the previous one ended.
  const startKms: Record<string, number> = {};
  if (body.startKms && typeof body.startKms === "object") {
    for (const [k, v] of Object.entries(body.startKms as Record<string, unknown>)) {
      const n = Number(v);
      if (Number.isFinite(n)) startKms[String(k)] = n;
    }
  }

  // De-dupe while preserving order.
  const orderedIds = Array.from(new Set(sourceProjectIds));
  if (orderedIds.length < 2) {
    return Response.json({ error: "Select at least 2 projects to combine." }, { status: 400 });
  }
  if (!name) {
    return Response.json({ error: "A name for the combined project is required." }, { status: 400 });
  }

  const conn = await pool.getConnection();
  try {
    // Validate all sources exist.
    const placeholders = orderedIds.map(() => "?").join(",");
    const [existRows] = await conn.query(
      `SELECT id, name FROM projects WHERE id IN (${placeholders})`,
      orderedIds
    );
    const found = new Map<string, string>(
      (Array.isArray(existRows) ? existRows : []).map((r: any) => [String(r.id), String(r.name || "")])
    );
    const missing = orderedIds.filter((id) => !found.has(id));
    if (missing.length) {
      conn.release();
      return Response.json({ error: "Some projects were not found", missing }, { status: 404 });
    }

    const projectCols = await columnsOf(conn, "projects");
    const reportCols = await columnsOf(conn, "reports");
    const photoCols = await columnsOf(conn, "report_photos");

    await conn.beginTransaction();

    // 1) Create the combined project.
    const newProjectId = uuidv4();
    await insertRow(
      conn,
      "projects",
      {
        id: newProjectId,
        name,
        description: `Combined from: ${orderedIds.map((id) => found.get(id)).join(" + ")}`,
        status: "active",
        created_by: authUser.id,
        last_modified_by: authUser.id,
        user_id: authUser.id,
      },
      projectCols
    );

    // 2) Copy each source project's reports (in order), renumbered sequentially.
    let seq = 0;
    let reportsCopied = 0;
    let photosCopied = 0;
    // Build WHERE / ORDER BY only from columns this install actually has.
    const delFilter = reportCols.has("deleted_at") ? "AND deleted_at IS NULL" : "";
    const orderCols = ["sort_order", "created_at", "id"].filter((cch) => reportCols.has(cch));
    const orderBy = orderCols.length ? `ORDER BY ${orderCols.map((cch) => `${cch} ASC`).join(", ")}` : "";
    // KM (chainage) for the combined project. Each source project begins at the
    // "Start KM" entered in the combine dialog (or, if blank, continues from
    // where the previous project ended). WITHIN a project the point-to-point KM
    // is kept as-is: its own stored KM when it has one, otherwise the straight-
    // line distance travelled from that project's first point.
    const kmCols = ["kms", "km"].filter((c) => reportCols.has(c));
    const hasLatLon = reportCols.has("loc_lat") && reportCols.has("loc_lon");
    let prevEndKm = 0;
    for (const srcId of orderedIds) {
      const [reportRows] = await conn.query(
        `SELECT * FROM reports WHERE project_id = ? ${delFilter} ${orderBy}`,
        [srcId]
      );
      const reports = (Array.isArray(reportRows) ? reportRows : []) as any[];

      // Where this project's KM starts.
      const startKm = Number.isFinite(startKms[srcId]) ? startKms[srcId] : prevEndKm;
      // Each point's distance ALONG this project (own KM, else GPS cumulative).
      const within: number[] = [];
      let cum = 0;
      let prevLat: number | null = null;
      let prevLon: number | null = null;
      for (const r of reports) {
        const lat = hasLatLon ? Number(r.loc_lat) : NaN;
        const lon = hasLatLon ? Number(r.loc_lon) : NaN;
        if (prevLat != null && prevLon != null && Number.isFinite(lat) && Number.isFinite(lon)) {
          cum += haversineKm(prevLat, prevLon, lat, lon);
        }
        const own = Number(r.kms ?? r.km);
        within.push(Number.isFinite(own) ? own : cum);
        if (Number.isFinite(lat) && Number.isFinite(lon)) {
          prevLat = lat;
          prevLon = lon;
        }
      }
      const firstWithin = within.length ? within[0] : 0;
      let projMaxKm = startKm;

      for (let ri = 0; ri < reports.length; ri++) {
        const r = reports[ri];
        seq += 1;
        const newReportId = uuidv4();
        const payload = copyPayload(r, REPORT_DROP, reportCols);
        payload.id = newReportId;
        payload.project_id = newProjectId;
        payload.point_key = String(seq);
        payload.sort_order = seq * 10;
        payload.user_id = authUser.id;
        payload.created_by = authUser.id;
        // KM = this project's start + how far this point is along the project.
        if (kmCols.length) {
          const adjusted = Math.round((startKm + (within[ri] - firstWithin)) * 100) / 100;
          for (const c of kmCols) payload[c] = adjusted;
          if (adjusted > projMaxKm) projMaxKm = adjusted;
        }
        await insertRow(conn, "reports", payload, reportCols);
        reportsCopied += 1;

        // Copy this report's photos (same S3 URLs).
        const [photoRows] = await conn.query(
          "SELECT * FROM report_photos WHERE report_id = ?",
          [r.id]
        );
        for (const ph of (Array.isArray(photoRows) ? photoRows : []) as any[]) {
          const pp = copyPayload(ph, PHOTO_DROP, photoCols);
          pp.id = uuidv4();
          pp.report_id = newReportId;
          pp.point_key = String(seq);
          pp.user_id = authUser.id;
          await insertRow(conn, "report_photos", pp, photoCols);
          photosCopied += 1;
        }
      }
      // The next project (if it has no Start KM entered) continues from here.
      prevEndKm = projMaxKm;
    }

    await conn.commit();
    conn.release();

    await logActivity(request, {
      action: "combine_projects",
      table: "projects",
      entityId: newProjectId,
      projectId: newProjectId,
      rowCount: reportsCopied,
      details: { sources: orderedIds, name, reportsCopied, photosCopied },
    });

    return Response.json({
      ok: true,
      projectId: newProjectId,
      projectName: name,
      sourcesCombined: orderedIds.length,
      reportsCopied,
      photosCopied,
    });
  } catch (error: any) {
    try {
      await conn.rollback();
    } catch {
      /* ignore */
    }
    conn.release();
    console.error("[api/projects/combine] error:", error);
    return Response.json(
      { error: "Failed to combine projects", detail: error?.message || String(error) },
      { status: 500 }
    );
  }
}
