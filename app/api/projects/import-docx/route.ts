/* eslint-disable @typescript-eslint/no-explicit-any */
// Import an OLD survey report .docx and turn it into a new website project:
// one report per observation point (coords, KM, location, category,
// observation, remarks) plus the road photos uploaded to S3. Reuses the same
// tables the rest of the app reads, so an imported project behaves exactly
// like a normally-created one and can be reviewed/edited in the grid.
import { promises as fsp } from "fs";
import path from "path";
import { v4 as uuidv4 } from "uuid";
import pool from "../../../../lib/db";
import { requireAuth } from "../../../../lib/auth";
import { parseSurveyReport } from "../../../../lib/docxReportImport";
import { uploadBufferToS3 } from "../../../../lib/s3";
import { logActivity } from "../../../../lib/activityLog";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 300;

async function columnsOf(table: string): Promise<Set<string>> {
  const [rows] = await pool.query(`SHOW COLUMNS FROM ${table}`);
  return new Set(
    (Array.isArray(rows) ? rows : []).map((r) => String((r as any)?.Field || "").toLowerCase())
  );
}

// Insert only the columns that actually exist on this install.
async function insertRow(table: string, payload: Record<string, unknown>, cols: Set<string>) {
  const keys = Object.keys(payload).filter((k) => cols.has(k.toLowerCase()));
  if (!keys.length) return;
  const placeholders = keys.map(() => "?").join(", ");
  await pool.query(
    `INSERT INTO ${table} (${keys.join(", ")}) VALUES (${placeholders})`,
    keys.map((k) => payload[k])
  );
}

function projectNameFromFile(fileName: string): string {
  return (
    fileName
      .replace(/\.[^.]+$/, "")
      .replace(/detail(?:ed)?\s*(survey\s*)?report/i, "")
      .replace(/route\s*sou?rvey\s*report/i, "")
      .replace(/-all-\d+/i, "")
      .replace(/\(\s*\d+\s*\)/g, "")
      .replace(/[_]+/g, " ")
      .replace(/\s{2,}/g, " ")
      .trim() || "Imported report"
  );
}

function contentTypeFor(name: string): string {
  const ext = (name.split(".").pop() || "").toLowerCase();
  if (ext === "png") return "image/png";
  if (ext === "webp") return "image/webp";
  if (ext === "gif") return "image/gif";
  return "image/jpeg";
}

// Save a photo to public/uploads so Next.js serves it at /uploads/... . Used as
// a fallback when S3 is not configured (e.g. local dev with placeholder keys)
// or when the S3 upload fails, so the imported report still shows its images.
async function saveBufferLocally(key: string, body: Buffer): Promise<string> {
  const rel = key.replace(/^\/+/, "");
  const dest = path.join(process.cwd(), "public", "uploads", rel);
  await fsp.mkdir(path.dirname(dest), { recursive: true });
  await fsp.writeFile(dest, body);
  return `/uploads/${rel}`;
}

// Try S3 first; if it isn't configured or fails, keep the image locally.
async function storePhoto(
  key: string,
  body: Buffer,
  contentType: string
): Promise<{ url: string; local: boolean }> {
  try {
    const { url } = await uploadBufferToS3({ key, body, contentType });
    return { url, local: false };
  } catch (err) {
    console.error("[import-docx] S3 upload failed, saving locally:", key, (err as any)?.message);
    const url = await saveBufferLocally(key, body);
    return { url, local: true };
  }
}

export async function POST(request: Request) {
  try {
    const authUser = requireAuth(request);

    const form = await request.formData();
    const file = form.get("file");
    if (!(file instanceof File)) {
      return Response.json({ error: "Upload a .docx or .pptx file in the 'file' field" }, { status: 400 });
    }
    if (!/\.(docx|pptx)$/i.test(file.name)) {
      return Response.json({ error: "Only .docx or .pptx files are supported" }, { status: 400 });
    }

    const buffer = Buffer.from(await file.arrayBuffer());
    const { points, getPhoto } = parseSurveyReport(buffer, file.name);
    if (!points.length) {
      return Response.json(
        { error: "No survey points found in this file. Is it the table-style report?" },
        { status: 422 }
      );
    }

    const projectCols = await columnsOf("projects");
    const reportCols = await columnsOf("reports");
    const photoCols = await columnsOf("report_photos");

    // 1) Create the project (named from the file).
    const projectId = uuidv4();
    const projectName = projectNameFromFile(file.name);
    await insertRow(
      "projects",
      {
        id: projectId,
        name: projectName,
        description: `Imported from ${file.name}`,
        status: "active",
        created_by: authUser.id,
        last_modified_by: authUser.id,
        user_id: authUser.id,
      },
      projectCols
    );

    // 2) One report per point, + photos.
    let reportsCreated = 0;
    let photosUploaded = 0;
    let photosFailed = 0;
    let photosLocal = 0;
    let sortOrder = 10;

    for (const p of points) {
      const reportId = uuidv4();
      const kmNum = Number(String(p.km || "").replace(/[^\d.]/g, ""));
      await insertRow(
        "reports",
        {
          id: reportId,
          user_id: authUser.id,
          created_by: authUser.id,
          project_id: projectId,
          category: p.category || "Report",
          description: p.observation || null,
          remarks_action: p.remarks || null,
          difficulty: p.difficulty || "green",
          status: "active",
          sort_order: sortOrder,
          point_key: p.point_key,
          latitude: p.latitude,
          longitude: p.longitude,
          loc_lat: p.latitude,
          loc_lon: p.longitude,
          resolved_location: p.location || null,
          location: p.location || null,
          kms: Number.isFinite(kmNum) ? kmNum : null,
        },
        reportCols
      );
      reportsCreated += 1;
      sortOrder += 10;

      // Upload this point's photos.
      let photoIdx = 0;
      for (const zipPath of p.photoNames) {
        const bytes = getPhoto(zipPath);
        if (!bytes || bytes.length < 1000) continue; // skip empties/icons that slipped through
        photoIdx += 1;
        const base = zipPath.split("/").pop() || `photo_${photoIdx}`;
        const key = `reports/photos/imported/${projectId}/${reportId}/${p.point_key}_${photoIdx}_${base}`;
        try {
          const { url, local } = await storePhoto(key, bytes, contentTypeFor(base));
          await insertRow(
            "report_photos",
            {
              id: uuidv4(),
              report_id: reportId,
              url,
              file_name: base,
              point_key: p.point_key,
              image_key: `${p.point_key}.${photoIdx}`,
              user_id: authUser.id,
              include_in_export: 1,
              latitude: p.latitude,
              longitude: p.longitude,
              has_gps: p.latitude != null ? 1 : 0,
            },
            photoCols
          );
          photosUploaded += 1;
          if (local) photosLocal += 1;
        } catch (err) {
          console.error("[import-docx] photo save failed:", key, err);
          photosFailed += 1;
        }
      }
    }

    await logActivity(request, {
      action: "import_docx",
      table: "projects",
      entityId: projectId,
      projectId,
      rowCount: reportsCreated,
      details: { file: file.name, points: points.length, photosUploaded, photosFailed, photosLocal },
    });

    return Response.json({
      ok: true,
      projectId,
      projectName,
      pointsFound: points.length,
      reportsCreated,
      photosUploaded,
      photosFailed,
      photosLocal,
      withCoords: points.filter((p) => Number.isFinite(p.latitude) && Number.isFinite(p.longitude)).length,
    });
  } catch (error: any) {
    if (error?.message === "Unauthorized") {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/projects/import-docx] error:", error);
    return Response.json(
      { error: "Failed to import report", detail: error?.message || String(error) },
      { status: 500 }
    );
  }
}
