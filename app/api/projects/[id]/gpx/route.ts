/* eslint-disable @typescript-eslint/no-explicit-any */
// Server-side GPX export: builds a .gpx with (a) the recorded GPS TRACK from
// report_path_points and (b) an observation waypoint per report. Reads MySQL
// directly (fast, reliable) instead of the client shim's hundreds of round
// trips. Coordinates come from report_path_points.latitude/longitude, falling
// back to the report's own loc_lat/loc_lon.
//
// Which points join into lines is decided in lib/gpx-track.ts — each report's
// trace is its own <trkseg>, so traces are never stitched to one another.
import pool from "../../../../../lib/db";
import { requireAuth } from "../../../../../lib/auth";
import {
  buildSegments,
  gpxTime,
  validLatLon,
  type GpxReport,
  type TrackPoint,
} from "../../../../../lib/gpx-track";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 120;

type Ctx = { params: { id: string } };

function xmlEsc(s: unknown) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
async function getColumns(table: string): Promise<Set<string>> {
  const [rows] = await pool.query(`SHOW COLUMNS FROM ${table}`);
  return new Set(
    (Array.isArray(rows) ? rows : []).map((r) => String((r as any)?.Field || "").toLowerCase())
  );
}

export async function GET(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const projectId = String(context.params?.id || "").trim();
    if (!projectId) return Response.json({ error: "Project id is required" }, { status: 400 });

    const url = new URL(request.url);
    const nameParam = String(url.searchParams.get("name") || "").trim();
    const reportIds = String(url.searchParams.get("reportIds") || "")
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean);

    const [projRows] = await pool.query("SELECT id, name FROM projects WHERE id = ? LIMIT 1", [
      projectId,
    ]);
    const project = Array.isArray(projRows) && projRows.length ? (projRows[0] as any) : null;
    if (!project) return Response.json({ error: "Project not found" }, { status: 404 });
    const projectName = nameParam || String(project.name || "route");

    // Reports in map order (optionally limited to a selection).
    const cols = await getColumns("reports");
    const orderParts: string[] = [];
    if (cols.has("sort_order")) orderParts.push("sort_order ASC");
    if (cols.has("created_at")) orderParts.push("created_at ASC");
    if (cols.has("id")) orderParts.push("id ASC");
    const orderBy = orderParts.length ? `ORDER BY ${orderParts.join(", ")}` : "";
    // No deleted_at filter — mirror the Word export (which works), so installs
    // where active reports don't store deleted_at = NULL still return rows.
    const [reportRows] = await pool.query(
      `SELECT * FROM reports WHERE project_id = ? ${orderBy}`,
      [projectId]
    );
    let reports = Array.isArray(reportRows) ? (reportRows as any[]) : [];
    if (reportIds.length) {
      // Keep the order the page sent, not the database's: the line is drawn
      // point to point in this order, so it must match what is on screen.
      const byId = new Map(reports.map((r) => [String(r.id), r]));
      reports = reportIds.map((id) => byId.get(id)).filter(Boolean);
    }
    if (!reports.length) {
      // Log WHY there are no reports so a production failure is diagnosable
      // from the pm2 log (raw count ignores every filter).
      let rawCount = -1;
      try {
        const [cnt] = await pool.query(
          "SELECT COUNT(*) AS n FROM reports WHERE project_id = ?",
          [projectId]
        );
        rawCount = Number((cnt as any)?.[0]?.n ?? -1);
      } catch (e) {
        console.error("[gpx] count probe failed:", e);
      }
      console.error(
        `[gpx] NO REPORTS for project ${projectId} — raw count in reports table = ${rawCount}, reportIds filter = ${reportIds.length}`
      );
      return Response.json(
        { error: "No reports available for GPX export.", projectId, rawCount },
        { status: 400 }
      );
    }

    // Path points for all these reports, grouped by report.
    const ids = reports.map((r) => String(r.id));
    const placeholders = ids.map(() => "?").join(",");
    const byReport = new Map<string, any[]>();
    try {
      const [ptRows] = await pool.query(
        `SELECT report_id, latitude, longitude, seq, timestamp, created_at
           FROM report_path_points
          WHERE report_id IN (${placeholders})
          ORDER BY report_id ASC, seq ASC`,
        ids
      );
      for (const p of Array.isArray(ptRows) ? (ptRows as any[]) : []) {
        const k = String(p.report_id);
        if (!byReport.has(k)) byReport.set(k, []);
        byReport.get(k)!.push(p);
      }
    } catch {
      /* table may not exist on some installs — fall back to report coords */
    }

    // One waypoint per report, and the report's own data for the track.
    const wpts: string[] = [];
    const trackInput: GpxReport[] = [];
    for (const r of reports) {
      const rid = String(r.id);
      const label = `${r.point_key || ""} ${r.category || "Report"}`.trim();
      const rLat = Number(r.loc_lat ?? r.latitude);
      const rLon = Number(r.loc_lon ?? r.longitude);
      const hasObs = validLatLon(rLat, rLon);
      if (hasObs) {
        wpts.push(
          `  <wpt lat="${rLat}" lon="${rLon}"><name>${xmlEsc(label)}</name>` +
            `${r.description ? `<desc>${xmlEsc(r.description)}</desc>` : ""}</wpt>`
        );
      }

      const path: TrackPoint[] = [];
      for (const p of byReport.get(rid) || []) {
        const lat = Number(p.latitude);
        const lon = Number(p.longitude);
        if (!validLatLon(lat, lon)) continue;
        // An unparseable timestamp used to throw here and fail the whole
        // export; now it just leaves that point without a time.
        path.push({ lat, lon, time: gpxTime(p.timestamp || p.created_at) });
      }
      trackInput.push({
        id: rid,
        lat: hasObs ? rLat : null,
        lon: hasObs ? rLon : null,
        path,
      });
    }

    const segments = buildSegments(trackInput);

    if (!segments.length && !wpts.length) {
      return Response.json(
        { error: "No valid coordinates found to export GPX." },
        { status: 400 }
      );
    }

    // Several <trkseg>s in one <trk> is standard GPX: viewers draw each as its
    // own line and leave the gaps between them empty.
    const trkXml = segments.length
      ? `  <trk><name>${xmlEsc(projectName)} route</name>\n` +
        segments
          .map(
            (seg) =>
              `    <trkseg>\n` +
              seg
                .map(
                  (p) =>
                    `      <trkpt lat="${p.lat}" lon="${p.lon}">` +
                    `${p.time ? `<time>${p.time}</time>` : ""}</trkpt>`
                )
                .join("\n") +
              `\n    </trkseg>`
          )
          .join("\n") +
        `\n  </trk>\n`
      : "";

    const gpx =
      `<?xml version="1.0" encoding="UTF-8"?>\n` +
      `<gpx version="1.1" creator="LBI Web App" xmlns="http://www.topografix.com/GPX/1/1">\n` +
      `  <metadata><name>${xmlEsc(projectName)}</name></metadata>\n` +
      `${wpts.join("\n")}${wpts.length ? "\n" : ""}` +
      trkXml +
      `</gpx>\n`;

    const safeName =
      (projectName.replace(/[\r\n"\\/?*<>|:]/g, "_").slice(0, 120) || "route") + ".gpx";

    return new Response(gpx, {
      status: 200,
      headers: {
        "Content-Type": "application/gpx+xml; charset=utf-8",
        "Content-Disposition": `attachment; filename="${safeName}"; filename*=UTF-8''${encodeURIComponent(safeName)}`,
        "Cache-Control": "no-store",
      },
    });
  } catch (error) {
    if ((error as { message?: string })?.message === "Unauthorized") {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/projects/:id/gpx] error:", error);
    return Response.json({ error: "Failed to build GPX" }, { status: 500 });
  }
}
