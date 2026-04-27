import { randomUUID } from "crypto";
import { PutObjectCommand } from "@aws-sdk/client-s3";
import { requireAuth } from "../../../lib/auth";
import { s3Client, S3_BUCKET_NAME, safeObjectKey } from "../../../lib/s3";
import pool from "../../../lib/db";

export const runtime = "nodejs";

function safeName(name: string) {
  return String(name || "upload")
    .replace(/[^a-zA-Z0-9._-]/g, "-")
    .replace(/-+/g, "-")
    .slice(0, 180);
}

function safeFolder(value: string) {
  return String(value || "uploads")
    .replace(/^\/+|\/+$/g, "")
    .replace(/[^a-zA-Z0-9/_-]/g, "-");
}

/**
 * Insert one row into report_photos when /api/upload is called with reportId.
 * This protects against orphaned S3 photos from clients that upload but never
 * call /api/reports/[id]/photos. Validates that the report row exists first.
 */
async function maybeInsertReportPhoto(args: {
  reportId: string;
  url: string;
  fileName: string | null;
  key: string;
  width: number | null;
  height: number | null;
}) {
  const { reportId, url, fileName, key, width, height } = args;
  try {
    const [reportRows] = await pool.query("SELECT id FROM reports WHERE id = ? LIMIT 1", [
      reportId,
    ]);
    if (!Array.isArray(reportRows) || reportRows.length === 0) {
      console.warn("[api/upload] reportId provided but report not found:", reportId);
      return { saved: false, reason: "report_not_found" as const };
    }
    const id = randomUUID();
    await pool.query(
      `INSERT INTO report_photos (id, report_id, url, file_name, path, width, height)
       VALUES (?, ?, ?, ?, ?, ?, ?)`,
      [id, reportId, url, fileName, key, width, height]
    );
    console.log("[api/upload] inserted report_photos row", { id, reportId, url });
    return { saved: true as const, photoId: id };
  } catch (err) {
    console.error("[api/upload] report_photos insert failed:", err);
    return { saved: false as const, reason: "insert_failed" };
  }
}

export async function POST(request: Request) {
  try {
    requireAuth(request);

    const formData = await request.formData();
    const folder = safeFolder(String(formData.get("folder") || "uploads"));
    const path = String(formData.get("path") || "").trim();
    const reportId = String(formData.get("reportId") || "").trim();
    const widthRaw = formData.get("width");
    const heightRaw = formData.get("height");
    const file = formData.get("file");

    if (!(file instanceof File)) {
      return Response.json({ error: "No file uploaded" }, { status: 400 });
    }

    const fileName = safeName(file.name || "upload");
    const key = path
      ? safeObjectKey(path, file.name)
      : reportId
        ? `reports/photos/${reportId}/${randomUUID()}-${fileName}`
        : `${folder}/${randomUUID()}-${fileName}`;
    const contentType = file.type || "application/octet-stream";
    const bytes = Buffer.from(await file.arrayBuffer());

    if (!S3_BUCKET_NAME) {
      return Response.json({ error: "S3 bucket is not configured" }, { status: 500 });
    }
    await s3Client.send(
      new PutObjectCommand({
        Bucket: S3_BUCKET_NAME,
        Key: key,
        Body: bytes,
        ContentType: contentType,
      })
    );

    const baseUrl = (process.env.NEXT_PUBLIC_S3_BUCKET_URL || "").replace(/\/+$/, "");
    if (!baseUrl) {
      return Response.json({ error: "S3 public URL is not configured" }, { status: 500 });
    }
    const url = `${baseUrl}/${key}`;

    let saveResult: { saved: boolean; reason?: string; photoId?: string } | null = null;
    if (reportId) {
      const w = Number(widthRaw);
      const h = Number(heightRaw);
      saveResult = await maybeInsertReportPhoto({
        reportId,
        url,
        fileName,
        key,
        width: Number.isFinite(w) ? w : null,
        height: Number.isFinite(h) ? h : null,
      });
    }

    return Response.json({
      url,
      key,
      path: key,
      fileName,
      contentType,
      size: file.size,
      reportPhoto: saveResult,
    });
  } catch (e: unknown) {
    const err = e as { message?: string };
    if (err?.message === "Unauthorized") {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/upload] error:", e);
    return Response.json({ error: "Failed to upload file" }, { status: 500 });
  }
}
