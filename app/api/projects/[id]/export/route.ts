import path from "path";
import fs from "fs";
import os from "os";
import crypto from "crypto";
import { NextResponse } from "next/server";
import pool from "../../../../../lib/db";
import { requireAuth } from "../../../../../lib/auth";
import { generateReenaDocx } from "../../../../../lib/reenaTemplateExport";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const TEMPLATE_PATH = path.join(process.cwd(), "templates", "reena-all-template.docx");

/**
 * Where the POST handler stashes the rendered DOCX so the follow-up GET
 * download endpoint can stream it back without re-rendering. Lives outside
 * the bundled output so Next does not try to scan or serve it as part of
 * the build artefacts.
 */
export const TEMP_EXPORT_DIR = path.join(os.tmpdir(), "lbi-exports");

/**
 * Best-effort prune of stale temp exports. The browser download window is a
 * few seconds at most; 30 minutes is a generous safety margin.
 */
const TEMP_EXPORT_TTL_MS = 30 * 60 * 1000;

function ensureTempExportDir() {
  try {
    fs.mkdirSync(TEMP_EXPORT_DIR, { recursive: true });
  } catch (err) {
    console.error("[export] mkdir TEMP_EXPORT_DIR failed:", err);
  }
}

function cleanupOldExports() {
  try {
    if (!fs.existsSync(TEMP_EXPORT_DIR)) return;
    const cutoff = Date.now() - TEMP_EXPORT_TTL_MS;
    for (const name of fs.readdirSync(TEMP_EXPORT_DIR)) {
      const full = path.join(TEMP_EXPORT_DIR, name);
      try {
        const stat = fs.statSync(full);
        if (stat.isFile() && stat.mtimeMs < cutoff) {
          fs.unlinkSync(full);
          console.log("[export] cleaned up stale temp file:", name);
        }
      } catch {
        // ignore individual stat / unlink errors
      }
    }
  } catch (err) {
    console.warn("[export] cleanupOldExports failed:", err);
  }
}

type Ctx = { params: { id: string } };
type Row = Record<string, any>;

function normalizeReportIds(raw: string) {
  return String(raw || "")
    .split(",")
    .map((v) => v.trim())
    .filter(Boolean)
    .slice(0, 1000);
}

async function safeQuery(sql: string, args: unknown[] = []): Promise<Row[]> {
  try {
    const [rows] = await pool.query(sql, args);
    return Array.isArray(rows) ? (rows as Row[]) : [];
  } catch (err) {
    console.error("[export] safeQuery failed:", { sql, err });
    return [];
  }
}

async function ensureGaDrawingsTable() {
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS project_ga_drawings (
        id VARCHAR(36) PRIMARY KEY,
        project_id VARCHAR(36) NOT NULL UNIQUE,
        image_url TEXT NULL,
        image_key TEXT NULL,
        file_name VARCHAR(255) NULL,
        conclusion_html LONGTEXT NULL,
        created_by VARCHAR(36) NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        INDEX idx_project_ga_drawings_project_id (project_id)
      )
    `);
  } catch (err) {
    console.error("[export] ensureGaDrawingsTable failed:", err);
  }
}

export async function GET(request: Request, context: Ctx) {
  const url = new URL(request.url);
  const projectId = String(context.params?.id || "").trim();
  const isCheck = url.searchParams.get("check") === "1";

  console.error("[export] route started", { projectId, check: isCheck });

  // ---- Auth.
  try {
    requireAuth(request);
  } catch (err) {
    if ((err as { message?: string })?.message === "Unauthorized") {
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[export] auth error:", err);
    return NextResponse.json({ error: "Auth check failed" }, { status: 500 });
  }

  if (!projectId) {
    return NextResponse.json({ error: "Project id is required" }, { status: 400 });
  }

  // ---- check=1 path: NEVER render DOCX, NEVER fetch S3. Return JSON only.
  if (isCheck) {
    try {
      const projectRows = await safeQuery(
        "SELECT id, name, title, project_name FROM projects WHERE id = ? LIMIT 1",
        [projectId]
      );
      const project = projectRows[0] || null;
      console.error("[export] project found:", !!project);

      const reports = await safeQuery(
        "SELECT id FROM reports WHERE project_id = ? ORDER BY sort_order ASC, created_at ASC",
        [projectId]
      );
      console.error("[export] reports count:", reports.length);

      const reportIds = reports.map((r) => String(r.id || "").trim()).filter(Boolean);
      let photos: Row[] = [];
      if (reportIds.length > 0) {
        const placeholders = reportIds.map(() => "?").join(",");
        photos = await safeQuery(
          `SELECT id, report_id, url, file_name, path, width, height, created_at
           FROM report_photos
           WHERE report_id IN (${placeholders})
           ORDER BY created_at ASC`,
          reportIds
        );
      }
      console.error("[export] photos count:", photos.length);

      await ensureGaDrawingsTable();
      const gaRows = await safeQuery(
        "SELECT id, project_id, image_url, image_key, file_name FROM project_ga_drawings WHERE project_id = ? LIMIT 1",
        [projectId]
      );
      const gaDrawing = gaRows[0] || null;
      console.error("[export] ga drawing found:", !!gaDrawing);

      const routeRows = await safeQuery(
        "SELECT id, project_id, objective, map_file_url FROM project_route_pages WHERE project_id = ? ORDER BY created_at DESC LIMIT 1",
        [projectId]
      );
      const routePage = routeRows[0] || null;

      // The original frontend reads { hasSetup, hasImages, pageId } - keep that
      // contract while also returning richer debug fields.
      const gaImageUrl = String(gaDrawing?.image_url || "").trim();
      const routeMapUrl = String(routePage?.map_file_url || "").trim();
      const hasSetup = Boolean(routePage || gaDrawing);
      const hasImages = Boolean(gaImageUrl || routeMapUrl);

      const templateExists = fs.existsSync(TEMPLATE_PATH);
      console.error("[export] template path:", TEMPLATE_PATH, "exists:", templateExists);

      // Lightweight per-report sample for KM/location debugging.
      const reportSamples = await safeQuery(
        `SELECT id, latitude, longitude, loc_lat, loc_lon, location, address, resolved_location,
                category, description, remarks_action, kms, sort_order, created_at
         FROM reports WHERE project_id = ? ORDER BY sort_order ASC, created_at ASC LIMIT 5`,
        [projectId]
      );
      // Inline mini-haversine so we can show the cumulative km without
      // importing the full helper.
      const reportsSample = (() => {
        let total = 0;
        const toRad = (v: number) => (v * Math.PI) / 180;
        return reportSamples.map((r, idx) => {
          const lat = Number(r.latitude ?? r.loc_lat);
          const lng = Number(r.longitude ?? r.loc_lon);
          if (idx > 0) {
            const prev = reportSamples[idx - 1];
            const pLat = Number(prev.latitude ?? prev.loc_lat);
            const pLng = Number(prev.longitude ?? prev.loc_lon);
            if (
              Number.isFinite(lat) &&
              Number.isFinite(lng) &&
              Number.isFinite(pLat) &&
              Number.isFinite(pLng)
            ) {
              const dLat = toRad(lat - pLat);
              const dLon = toRad(lng - pLng);
              const a =
                Math.sin(dLat / 2) ** 2 +
                Math.cos(toRad(pLat)) * Math.cos(toRad(lat)) * Math.sin(dLon / 2) ** 2;
              total += 6371 * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
            }
          }
          return {
            id: r.id,
            latitude: r.latitude ?? r.loc_lat,
            longitude: r.longitude ?? r.loc_lon,
            calculated_km: idx === 0 ? 0 : total,
            formatted_km: Number.isFinite(total) ? total.toFixed(4) : "-",
            resolved_location: r.resolved_location ?? r.location ?? r.address ?? null,
            category: r.category ?? null,
            description: r.description ?? null,
            remarks_action: r.remarks_action ?? null,
          };
        });
      })();

      // Per-report photo counts so the caller can verify the join is correct.
      const photosByReportIdCount: Record<string, number> = {};
      for (const p of photos) {
        const k = String(p.report_id || "").trim();
        if (!k) continue;
        photosByReportIdCount[k] = (photosByReportIdCount[k] || 0) + 1;
      }

      return NextResponse.json({
        ok: true,
        projectId,
        projectFound: !!project,
        pageId: routePage?.id || null,
        hasSetup,
        hasImages,
        reportsCount: reports.length,
        reportIds,
        photosCount: photos.length,
        photos: photos.map((p) => ({
          id: p.id,
          report_id: p.report_id,
          url: p.url,
          file_name: p.file_name,
        })),
        photosSample: photos.slice(0, 3).map((p) => ({
          report_id: p.report_id,
          url: p.url,
          file_name: p.file_name,
        })),
        photosByReportIdCount,
        reportsSample,
        gaDrawingImageUrl: gaImageUrl || null,
        routeMapUrl: routeMapUrl || null,
        templateExists,
        templatePath: TEMPLATE_PATH,
      });
    } catch (err) {
      console.error("[export] check failed:", err);
      return NextResponse.json(
        {
          ok: false,
          error: "Check failed",
          message: (err as { message?: string })?.message || "unknown",
        },
        { status: 500 }
      );
    }
  }

  // ---- Full DOCX render path (GET).
  return renderAndRespond({
    projectId,
    reportIds: normalizeReportIds(url.searchParams.get("reportIds") || ""),
    includePhotos: String(url.searchParams.get("includePhotos") || "1") !== "0",
    requestedFileName: String(url.searchParams.get("fileName") || "").trim(),
  });
}

/**
 * Render handler shared by GET (precheck + small selections) and POST (large
 * selections - avoids 414/431 from long reportIds query strings hitting the
 * upstream Nginx URL-length limit).
 */
async function renderAndRespond(args: {
  projectId: string;
  reportIds: string[];
  includePhotos: boolean;
  requestedFileName: string;
}) {
  const { projectId, reportIds, includePhotos, requestedFileName } = args;
  let stage: string = "started";
  try {
    console.log("[export stage] started", {
      projectId,
      selectedReports: reportIds.length,
      includePhotos,
    });
    stage = "template_check";
    if (!fs.existsSync(TEMPLATE_PATH)) {
      console.error("[export] template missing at", TEMPLATE_PATH);
      return NextResponse.json(
        { error: "DOCX template not found", stage, templatePath: TEMPLATE_PATH },
        { status: 500 }
      );
    }

    stage = "generate";
    const result = await generateReenaDocx({
      projectId,
      reportIds: reportIds.length ? reportIds : undefined,
      includePhotos,
    });

    stage = "respond";
    const fileName = requestedFileName
      ? requestedFileName.replace(/\.docx$/i, "") + ".docx"
      : result.fileName;
    const safeFileName = String(fileName)
      .replace(/[\r\n"\\/?*<>|:]/g, "_")
      .slice(0, 200);

    // Two-step delivery: write the rendered DOCX to a server-side temp file
    // and respond with a small JSON envelope pointing at the GET download
    // endpoint. This avoids the Chrome "Failed to fetch / no data found for
    // resource" failure mode that hits some setups when streaming a large
    // (16+ MB) binary response back through fetch().
    ensureTempExportDir();
    cleanupOldExports();

    // Random token + .docx extension. The GET endpoint validates that the
    // requested file lives under TEMP_EXPORT_DIR and matches this pattern.
    const token = crypto.randomBytes(16).toString("hex");
    const storedName = `${token}.docx`;
    const storedPath = path.join(TEMP_EXPORT_DIR, storedName);
    fs.writeFileSync(storedPath, result.buffer);

    const downloadUrl =
      `/api/projects/${encodeURIComponent(projectId)}/export/download` +
      `?file=${encodeURIComponent(storedName)}` +
      `&name=${encodeURIComponent(safeFileName)}`;

    console.log("[export stage] response sent", {
      projectId,
      fileName: safeFileName,
      storedName,
      bytes: result.buffer.length,
      downloadUrl,
    });

    return NextResponse.json(
      {
        success: true,
        fileName: safeFileName,
        downloadUrl,
        bytes: result.buffer.length,
      },
      {
        status: 200,
        headers: {
          "Cache-Control": "no-store, no-cache, must-revalidate",
        },
      }
    );
  } catch (error) {
    if ((error as { message?: string })?.message === "Project not found") {
      return NextResponse.json({ error: "Project not found", stage }, { status: 404 });
    }
    console.error("[export] render failed at stage", stage, error);
    const e = error as {
      message?: string;
      step?: string;
      detail?: string;
      stack?: string;
    };
    const isDev = process.env.NODE_ENV !== "production";
    return NextResponse.json(
      {
        error: "Failed to generate export",
        stage,
        step: e?.step || "unknown",
        message: e?.message || "unknown",
        detail: e?.detail,
        stack: isDev ? e?.stack : undefined,
      },
      { status: 500 }
    );
  }
}

/**
 * POST variant: same DOCX response as GET, but reads `reportIds` (and other
 * params) from the JSON body so URL length is bounded. Use this whenever the
 * caller has many selected reports.
 *
 * Body shape:
 *   { reportIds?: string[]; includePhotos?: boolean; fileName?: string }
 */
export async function POST(request: Request, context: Ctx) {
  const projectId = String(context.params?.id || "").trim();
  console.error("[export] route started (POST)", { projectId });

  try {
    requireAuth(request);
  } catch (err) {
    if ((err as { message?: string })?.message === "Unauthorized") {
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[export] auth error:", err);
    return NextResponse.json({ error: "Auth check failed" }, { status: 500 });
  }

  if (!projectId) {
    return NextResponse.json({ error: "Project id is required" }, { status: 400 });
  }

  let body: {
    reportIds?: unknown;
    includePhotos?: unknown;
    fileName?: unknown;
  } = {};
  try {
    body = (await request.json()) || {};
  } catch {
    body = {};
  }

  const reportIds = Array.isArray(body.reportIds)
    ? body.reportIds
        .map((v) => String(v || "").trim())
        .filter(Boolean)
        .slice(0, 5000)
    : [];
  const includePhotos = body.includePhotos !== false;
  const requestedFileName = String(body.fileName || "").trim();

  return renderAndRespond({ projectId, reportIds, includePhotos, requestedFileName });
}
