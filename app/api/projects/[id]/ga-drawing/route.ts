import { v4 as uuidv4 } from "uuid";
import pool from "../../../../../lib/db";
import { requireAuth } from "../../../../../lib/auth";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type Ctx = { params: { id: string } };

function unauthorized(error: unknown) {
  return (error as { message?: string })?.message === "Unauthorized";
}

async function ensureGaDrawingsTable() {
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
}

export async function GET(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const projectId = String(context.params?.id || "").trim();
    if (!projectId) {
      return Response.json({ error: "Project id is required" }, { status: 400 });
    }

    const { searchParams } = new URL(request.url);
    const pageIdParam = String(searchParams.get("pageId") || "").trim();

    let page: any = null;
    if (pageIdParam) {
      const [rows] = await pool.query(
        "SELECT * FROM project_route_pages WHERE id = ? AND project_id = ? LIMIT 1",
        [pageIdParam, projectId]
      );
      page = Array.isArray(rows) && rows.length ? (rows[0] as any) : null;
    } else {
      const [rows] = await pool.query(
        "SELECT * FROM project_route_pages WHERE project_id = ? ORDER BY created_at DESC LIMIT 1",
        [projectId]
      );
      page = Array.isArray(rows) && rows.length ? (rows[0] as any) : null;
    }

    const pageId = String(page?.id || "").trim();

    let locations: any[] = [];
    if (pageId) {
      const [locRows] = await pool.query(
        "SELECT id, label, pin_type, sort_order FROM project_route_page_locations WHERE project_id = ? AND project_page_id = ? ORDER BY sort_order ASC",
        [projectId, pageId]
      );
      locations = Array.isArray(locRows) ? (locRows as any[]) : [];
    }

    let images: any[] = [];
    if (pageId) {
      const [imgRows] = await pool.query(
        "SELECT * FROM project_route_page_images WHERE project_id = ? AND project_page_id = ? ORDER BY created_at ASC",
        [projectId, pageId]
      );
      images = Array.isArray(imgRows) ? (imgRows as any[]) : [];
    }

    await ensureGaDrawingsTable();
    const [gaRows] = await pool.query(
      "SELECT * FROM project_ga_drawings WHERE project_id = ? LIMIT 1",
      [projectId]
    );
    const gaDrawing = Array.isArray(gaRows) && gaRows.length ? (gaRows[0] as any) : null;

    return Response.json({
      page: page || null,
      locations,
      images,
      gaDrawing,
    });
  } catch (error) {
    if (unauthorized(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/projects/[id]/ga-drawing] GET error:", error);
    return Response.json({ error: "Failed to load GA drawing" }, { status: 500 });
  }
}

export async function POST(request: Request, context: Ctx) {
  try {
    const authUser = requireAuth(request);
    const projectId = String(context.params?.id || "").trim();
    if (!projectId) {
      return Response.json({ error: "Project id is required" }, { status: 400 });
    }

    const body = await request.json().catch(() => ({} as any));
    const objective = String(body?.objective || "").trim();
    const mapMode = String(body?.mapMode || "preset").trim();
    const presetMapKey = body?.presetMapKey ? String(body.presetMapKey).trim() : null;
    const mapFileUrl = body?.mapFileUrl ? String(body.mapFileUrl).trim() : null;
    const conclusionHtml = body?.conclusionHtml ? String(body.conclusionHtml) : null;
    const routeLocations = Array.isArray(body?.routeLocations) ? body.routeLocations : [];
    const incomingImages = Array.isArray(body?.gaImages) ? body.gaImages : [];
    const explicitPageId = body?.pageId ? String(body.pageId).trim() : "";

    let pageId = explicitPageId;
    if (pageId) {
      const [existsRows] = await pool.query(
        "SELECT id FROM project_route_pages WHERE id = ? AND project_id = ? LIMIT 1",
        [pageId, projectId]
      );
      const exists = Array.isArray(existsRows) && existsRows.length > 0;
      if (!exists) pageId = "";
    }

    if (!pageId) {
      const [latestRows] = await pool.query(
        "SELECT id FROM project_route_pages WHERE project_id = ? ORDER BY created_at DESC LIMIT 1",
        [projectId]
      );
      pageId = Array.isArray(latestRows) && latestRows.length ? String((latestRows[0] as any).id) : "";
    }

    if (pageId) {
      await pool.query(
        `UPDATE project_route_pages
         SET objective = ?, map_mode = ?, preset_map_key = ?, map_file_url = ?, conclusion_html = ?
         WHERE id = ?`,
        [objective || null, mapMode || null, presetMapKey, mapFileUrl, conclusionHtml, pageId]
      );
    } else {
      pageId = uuidv4();
      await pool.query(
        `INSERT INTO project_route_pages
         (id, project_id, user_id, objective, map_mode, preset_map_key, map_file_url, conclusion_html)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
        [pageId, projectId, authUser.id, objective || null, mapMode || null, presetMapKey, mapFileUrl, conclusionHtml]
      );
    }

    await pool.query(
      "DELETE FROM project_route_page_locations WHERE project_id = ? AND project_page_id = ?",
      [projectId, pageId]
    );

    const locRows = routeLocations
      .map((x: any, idx: number) => ({
        id: uuidv4(),
        project_page_id: pageId,
        project_id: projectId,
        user_id: authUser.id,
        label: String(x || "").trim(),
        pin_type: idx === 0 ? "start" : idx === 3 ? "end" : "mid",
        sort_order: idx,
      }))
      .filter((r: any) => r.label);

    if (locRows.length) {
      await pool.query(
        `INSERT INTO project_route_page_locations
         (id, project_page_id, project_id, user_id, label, pin_type, sort_order)
         VALUES ${locRows.map(() => "(?, ?, ?, ?, ?, ?, ?)").join(", ")}`,
        locRows.flatMap((r: any) => [r.id, r.project_page_id, r.project_id, r.user_id, r.label, r.pin_type, r.sort_order])
      );
    }

    await pool.query(
      "DELETE FROM project_route_page_images WHERE project_id = ? AND project_page_id = ?",
      [projectId, pageId]
    );

    const imageRows = incomingImages
      .map((x: any) => ({
        id: uuidv4(),
        project_page_id: pageId,
        project_id: projectId,
        user_id: authUser.id,
        file_url: String(x?.file_url || x?.imageUrl || x?.url || "").trim(),
        file_name: String(x?.file_name || x?.fileName || "").trim() || null,
        mime_type: String(x?.mime_type || x?.mimeType || "").trim() || null,
        file_size: Number.isFinite(Number(x?.file_size ?? x?.fileSize)) ? Number(x?.file_size ?? x?.fileSize) : null,
      }))
      .filter((r: any) => r.file_url);

    if (imageRows.length) {
      await pool.query(
        `INSERT INTO project_route_page_images
         (id, project_page_id, project_id, user_id, file_url, file_name, mime_type, file_size)
         VALUES ${imageRows.map(() => "(?, ?, ?, ?, ?, ?, ?, ?)").join(", ")}`,
        imageRows.flatMap((r: any) => [
          r.id,
          r.project_page_id,
          r.project_id,
          r.user_id,
          r.file_url,
          r.file_name,
          r.mime_type,
          r.file_size,
        ])
      );
    }

    await ensureGaDrawingsTable();

    const firstImage = imageRows[0] || null;
    const imageUrl = String(body?.imageUrl || firstImage?.file_url || "").trim() || null;
    const imageKey = String(body?.imageKey || "").trim() || null;
    const fileName = String(body?.fileName || firstImage?.file_name || "").trim() || null;

    await pool.query(
      `INSERT INTO project_ga_drawings
       (id, project_id, image_url, image_key, file_name, conclusion_html, created_by)
       VALUES (?, ?, ?, ?, ?, ?, ?)
       ON DUPLICATE KEY UPDATE
         image_url = VALUES(image_url),
         image_key = VALUES(image_key),
         file_name = VALUES(file_name),
         conclusion_html = VALUES(conclusion_html),
         updated_at = CURRENT_TIMESTAMP`,
      [uuidv4(), projectId, imageUrl, imageKey, fileName, conclusionHtml, authUser.id]
    );

    return Response.json({
      ok: true,
      pageId,
      imageUrl,
      imageKey,
      fileName,
      conclusionHtml,
    });
  } catch (error) {
    if (unauthorized(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/projects/[id]/ga-drawing] POST error:", error);
    return Response.json({ error: "Failed to save GA drawing" }, { status: 500 });
  }
}
