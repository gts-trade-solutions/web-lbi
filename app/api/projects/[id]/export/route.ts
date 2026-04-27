import path from "path";
import fs from "fs";
import { NextResponse } from "next/server";
import pool from "../../../../../lib/db";
import { requireAuth } from "../../../../../lib/auth";
import { generateReenaDocx } from "../../../../../lib/reenaTemplateExport";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const TEMPLATE_PATH = path.join(process.cwd(), "templates", "reena-all-template.docx");

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

  // ---- Full DOCX render path.
  try {
    if (!fs.existsSync(TEMPLATE_PATH)) {
      console.error("[export] template missing at", TEMPLATE_PATH);
      return NextResponse.json(
        { error: "DOCX template not found", templatePath: TEMPLATE_PATH },
        { status: 500 }
      );
    }

    const includePhotos = String(url.searchParams.get("includePhotos") || "1") !== "0";
    const reportIds = normalizeReportIds(url.searchParams.get("reportIds") || "");
    const requestedFileName = String(url.searchParams.get("fileName") || "").trim();

    const result = await generateReenaDocx({
      projectId,
      reportIds: reportIds.length ? reportIds : undefined,
      includePhotos,
    });

    const fileName = requestedFileName
      ? requestedFileName.replace(/\.docx$/i, "") + ".docx"
      : result.fileName;

    return new NextResponse(new Uint8Array(result.buffer), {
      status: 200,
      headers: {
        "Content-Type": DOCX_MIME,
        "Content-Disposition": `attachment; filename="${fileName}"`,
      },
    });
  } catch (error) {
    if ((error as { message?: string })?.message === "Project not found") {
      return NextResponse.json({ error: "Project not found" }, { status: 404 });
    }
    console.error("[export] render failed:", error);
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
        step: e?.step || "unknown",
        message: e?.message || "unknown",
        detail: e?.detail,
        stack: isDev ? e?.stack : undefined,
      },
      { status: 500 }
    );
  }
}
