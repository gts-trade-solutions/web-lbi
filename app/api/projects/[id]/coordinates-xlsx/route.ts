import { NextResponse } from "next/server";
import * as XLSX from "xlsx";
import pool from "../../../../../lib/db";
import { requireAuth } from "../../../../../lib/auth";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type Ctx = { params: { id: string } };
type Row = Record<string, any>;

async function safeQuery(sql: string, args: unknown[] = []): Promise<Row[]> {
  try {
    const [rows] = await pool.query(sql, args);
    return Array.isArray(rows) ? (rows as Row[]) : [];
  } catch (err) {
    console.error("[coordinates-xlsx] query failed:", { sql, err });
    return [];
  }
}

/**
 * Coordinate + KM extraction mirrors lib/reenaTemplateExport.ts so the Excel
 * KM column matches the Word export exactly.
 */
function pickLat(r: Row): number | null {
  for (const c of [r.latitude, r.lat, r.ne_latitude, r.ne_lat, r.loc_lat]) {
    const n = Number(c);
    if (Number.isFinite(n)) return n;
  }
  return null;
}
function pickLng(r: Row): number | null {
  for (const c of [r.longitude, r.lng, r.ne_longitude, r.ne_lng, r.lon, r.loc_lon]) {
    const n = Number(c);
    if (Number.isFinite(n)) return n;
  }
  return null;
}
function haversineKm(lat1: number, lon1: number, lat2: number, lon2: number): number {
  const R = 6371;
  const toRad = (v: number) => (v * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function normalizeReportIds(raw: string) {
  return String(raw || "")
    .split(",")
    .map((v) => v.trim())
    .filter(Boolean)
    .slice(0, 5000);
}

export async function GET(request: Request, context: Ctx) {
  const projectId = String(context.params?.id || "").trim();

  try {
    requireAuth(request);
  } catch {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  if (!projectId) {
    return NextResponse.json({ error: "Project id is required" }, { status: 400 });
  }

  // Optional subset filter (?reportIds=a,b,c). Default = all reports.
  const url = new URL(request.url);
  const reportIdsFilter = normalizeReportIds(url.searchParams.get("reportIds") || "");

  // Project name for the file name.
  const projectRows = await safeQuery(
    "SELECT id, name, title, project_name FROM projects WHERE id = ? LIMIT 1",
    [projectId]
  );
  const project = projectRows[0] || null;
  const projectName =
    String(project?.name || project?.title || project?.project_name || "Project").trim() || "Project";

  // Same ordering as the Word export so KM accumulates identically.
  const reportSql =
    reportIdsFilter.length > 0
      ? `SELECT * FROM reports WHERE project_id = ? AND id IN (${reportIdsFilter
          .map(() => "?")
          .join(",")}) ORDER BY sort_order ASC, created_at ASC`
      : "SELECT * FROM reports WHERE project_id = ? ORDER BY sort_order ASC, created_at ASC";
  const reportArgs = reportIdsFilter.length ? [projectId, ...reportIdsFilter] : [projectId];
  const reports = await safeQuery(reportSql, reportArgs);

  // Build rows with cumulative KM (first report = 0).
  let total = 0;
  const aoa: (string | number)[][] = [["S.No", "KM", "Coordinates (Lat, Long)"]];

  reports.forEach((r, idx) => {
    const lat = pickLat(r);
    const lng = pickLng(r);

    if (idx > 0) {
      const prev = reports[idx - 1];
      const pLat = pickLat(prev);
      const pLng = pickLng(prev);
      if (lat !== null && lng !== null && pLat !== null && pLng !== null) {
        total += haversineKm(pLat, pLng, lat, lng);
      }
    }

    const coords =
      lat !== null && lng !== null ? `${lat.toFixed(6)}, ${lng.toFixed(6)}` : "";

    aoa.push([idx + 1, Number(total.toFixed(2)), coords]);
  });

  // Build the workbook.
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 6 }, { wch: 10 }, { wch: 26 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "KM & Coordinates");

  const buffer: Buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

  const safeName =
    projectName.replace(/[\r\n"\\/?*<>|:]/g, "_").slice(0, 80) + "-KM-Coordinates.xlsx";

  return new NextResponse(buffer, {
    status: 200,
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="${safeName}"`,
      "Content-Length": String(buffer.length),
      "Cache-Control": "no-store",
    },
  });
}
