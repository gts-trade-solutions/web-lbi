/* eslint-disable @typescript-eslint/no-explicit-any */
// Server-side GPX export: builds a .gpx with one observation WAYPOINT per
// report PLUS a ROAD-FOLLOWING track line that joins the points along the
// actual roads (not straight lines between them). Reads MySQL directly (fast,
// reliable). Coordinates come from the report's loc_lat/loc_lon (falling back
// to latitude/longitude).
//
// The track used to be a straight line from each point to the next, which — in
// report order — looked like a web of "spoke" lines. It was removed, then the
// user asked for a line that follows the roads, like a Google Maps route. So
// the points (in sort order) are now run through OpenRouteService (the same
// road-routing the animated map uses) and the returned road geometry becomes
// the <trkseg>. If routing is unavailable or a leg can't be routed, that leg
// falls back to a straight line so the track is never broken; if routing is not
// configured at all, the file is waypoints-only (never a misleading straight web).
import pool from "../../../../../lib/db";
import { requireAuth } from "../../../../../lib/auth";
import { validLatLon } from "../../../../../lib/gpx-track";
import { fetchOrsPath, ORS_MAX_COORDS } from "../../../../../lib/orsRoute";

type Pt = { lat: number; lon: number };

// Road geometry for a run of points, splitting on failure so one unroutable
// leg can't flatten the whole run to a straight line.
async function routeLeg(pts: Pt[], apiKey: string): Promise<Pt[]> {
  const r = await fetchOrsPath(pts.map((p) => [p.lat, p.lon]), apiKey);
  if (r.ok && r.path.length >= 2) return r.path.map((p) => ({ lat: p.lat, lon: p.lng }));
  if (pts.length <= 2) return pts.slice(); // a single leg that won't route → straight
  const mid = Math.floor(pts.length / 2);
  const left = await routeLeg(pts.slice(0, mid + 1), apiKey);
  const right = await routeLeg(pts.slice(mid), apiKey);
  return left.concat(right.slice(1)); // drop the shared join point
}

// Full road polyline for all points, in order. Chunked to ORS's per-request
// limit, with a one-point overlap so consecutive chunks join seamlessly.
async function buildRoadPolyline(pts: Pt[], apiKey: string): Promise<Pt[]> {
  const out: Pt[] = [];
  const step = Math.max(1, ORS_MAX_COORDS - 1);
  for (let i = 0; i < pts.length - 1; i += step) {
    const chunk = pts.slice(i, i + ORS_MAX_COORDS);
    if (chunk.length < 2) break;
    const leg = await routeLeg(chunk, apiKey);
    if (!out.length) out.push(...leg);
    else out.push(...leg.slice(1));
  }
  return out;
}

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

    // One waypoint per report — the observation pin — and the ordered points
    // that the road line will be routed through.
    const wpts: string[] = [];
    const orderedPts: Pt[] = [];
    for (const r of reports) {
      const label = `${r.point_key || ""} ${r.category || "Report"}`.trim();
      const rLat = Number(r.loc_lat ?? r.latitude);
      const rLon = Number(r.loc_lon ?? r.longitude);
      if (!validLatLon(rLat, rLon)) continue;
      wpts.push(
        `  <wpt lat="${rLat}" lon="${rLon}"><name>${xmlEsc(label)}</name>` +
          `${r.description ? `<desc>${xmlEsc(r.description)}</desc>` : ""}</wpt>`
      );
      orderedPts.push({ lat: rLat, lon: rLon });
    }

    if (!wpts.length) {
      return Response.json(
        { error: "No valid coordinates found to export GPX." },
        { status: 400 }
      );
    }

    // Road-following track through the points, in order. Skipped (waypoints
    // only) when routing isn't configured, so we never emit a misleading
    // straight-line web.
    let trkXml = "";
    const orsKey = process.env.ORS_API_KEY;
    if (orsKey && orderedPts.length >= 2) {
      try {
        const line = await buildRoadPolyline(orderedPts, orsKey);
        if (line.length >= 2) {
          const trkpts = line
            .map((p) => `      <trkpt lat="${p.lat}" lon="${p.lon}"></trkpt>`)
            .join("\n");
          trkXml =
            `  <trk><name>${xmlEsc(projectName)} route</name>\n` +
            `    <trkseg>\n${trkpts}\n    </trkseg>\n` +
            `  </trk>\n`;
        }
      } catch (e) {
        console.warn("[gpx] road routing failed, exporting waypoints only:", e);
      }
    } else if (!orsKey) {
      console.warn("[gpx] ORS_API_KEY not set — exporting waypoints only (no road line).");
    }

    const gpx =
      `<?xml version="1.0" encoding="UTF-8"?>\n` +
      `<gpx version="1.1" creator="LBI Web App" xmlns="http://www.topografix.com/GPX/1/1">\n` +
      `  <metadata><name>${xmlEsc(projectName)}</name></metadata>\n` +
      `${wpts.join("\n")}\n` +
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
