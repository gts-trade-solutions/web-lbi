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
 *
 * Hard-coded against the actual MySQL schema:
 *   id varchar(36), report_id varchar(36), url text, width int, height int,
 *   created_at timestamp DEFAULT CURRENT_TIMESTAMP, user_id varchar(36),
 *   file_name text, point_key text, image_key text
 *
 * Notes:
 *   - There is NO `path` column — never include it.
 *   - `created_at` is NEVER passed; the column default (CURRENT_TIMESTAMP)
 *     fills it on insert.
 */
async function maybeInsertReportPhoto(args: {
  reportId: string;
  url: string;
  fileName: string | null;
  width: number | null;
  height: number | null;
  imageKey: string | null;
  pointKey: string | null;
  userId: string | null;
}) {
  const { reportId, url, fileName, width, height, imageKey, pointKey, userId } = args;
  try {
    const [reportRows] = await pool.query("SELECT id FROM reports WHERE id = ? LIMIT 1", [
      reportId,
    ]);
    if (!Array.isArray(reportRows) || reportRows.length === 0) {
      console.warn("[api/upload] reportId provided but report not found:", reportId);
      return { saved: false, reason: "report_not_found" as const };
    }

    // Optional dedup: drop any prior row for the SAME (report_id, file_name)
    // so re-uploads of the same image replace cleanly. We do NOT wipe all
    // photos for the report (the previous broader DELETE could erase
    // unrelated manual photos).
    try {
      await pool.query(
        "DELETE FROM report_photos WHERE report_id = ? AND file_name = ?",
        [reportId, fileName]
      );
    } catch (delErr) {
      console.warn("[api/upload] dedup DELETE failed - continuing:", delErr);
    }

    const photoId = randomUUID();
    console.log("[api/upload] report_photos insert start", {
      photoId,
      reportId,
      publicUrl: url,
      width,
      height,
      userId,
      originalFileName: fileName,
      pointKey,
      imageKey,
    });

    // Schema-exact INSERT. Column order matches the actual report_photos
    // table; created_at is omitted because the column has DEFAULT
    // CURRENT_TIMESTAMP. NEVER add `path` here — the column does not
    // exist in this schema.
    await pool.query(
      `INSERT INTO report_photos
         (id, report_id, url, width, height, user_id, file_name, point_key, image_key)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        photoId,
        reportId,
        url,
        width ?? null,
        height ?? null,
        userId ?? null,
        fileName ?? null,
        pointKey ?? null,
        imageKey ?? null,
      ]
    );
    const id = photoId;
    console.log("[api/upload] inserted report_photos row", { id, reportId, url });

    // Spec-mandated verify SELECT. Reads back the exact column set the
    // table declares for THIS reportId, newest-first (LIMIT 5). Then a
    // second JOIN query exposes the parent project_id so the client can
    // confirm the link landed on the EXPECTED project.
    let verifiedCount = 0;
    let verifiedRows: Array<Record<string, unknown>> = [];
    let verifiedProjectId: string | null = null;
    try {
      const [vRows] = await pool.query(
        `SELECT id, report_id, url, width, height, user_id, file_name, point_key, image_key, created_at
         FROM report_photos
         WHERE report_id = ?
         ORDER BY created_at DESC
         LIMIT 5`,
        [reportId]
      );
      verifiedRows = Array.isArray(vRows) ? (vRows as Array<Record<string, unknown>>) : [];
      verifiedCount = verifiedRows.length;
      // Spec-mandated success log key. Operator greps for this exact line
      // — finding it with verifyCount: 1 proves the row landed.
      console.log("[api/upload] report_photos insert success", {
        reportId,
        verifyCount: verifiedCount,
        verifyRows: verifiedRows,
      });
    } catch (vErr) {
      console.warn("[api/upload] verify SELECT failed:", vErr);
    }
    // Separate, lightweight JOIN-back to reports so the response can prove
    // the FK landed on the correct project for the bulk-import client.
    try {
      const [pRows] = await pool.query(
        `SELECT r.project_id
         FROM report_photos rp
         JOIN reports r ON r.id = rp.report_id
         WHERE rp.report_id = ?
         LIMIT 1`,
        [reportId]
      );
      const first = Array.isArray(pRows) ? (pRows[0] as { project_id?: string } | undefined) : undefined;
      verifiedProjectId = first?.project_id ? String(first.project_id) : null;
    } catch (jErr) {
      console.warn("[api/upload] verify JOIN failed:", jErr);
    }

    return {
      saved: true as const,
      photoId: id,
      verifiedCount,
      verifiedRows,
      verifiedProjectId,
    };
  } catch (err) {
    const message = (err as { message?: string })?.message || String(err);
    console.error("[api/upload] report_photos insert failed HARD", err);
    return {
      saved: false as const,
      reason: "insert_failed",
      errorMessage: message,
    };
  }
}

export async function POST(request: Request) {
  try {
    const authUser = requireAuth(request);
    const userId = authUser?.id ? String(authUser.id) : null;

    const formData = await request.formData();
    const folder = safeFolder(String(formData.get("folder") || "uploads"));
    const path = String(formData.get("path") || "").trim();
    const reportId = String(formData.get("reportId") || "").trim();
    const widthRaw = formData.get("width");
    const heightRaw = formData.get("height");
    const imageKey = String(formData.get("imageKey") || "").trim() || null;
    const pointKey = String(formData.get("pointKey") || "").trim() || null;
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

    let saveResult: {
      saved: boolean;
      reason?: string;
      photoId?: string;
      verifiedCount?: number;
      verifiedRows?: Array<Record<string, unknown>>;
      verifiedProjectId?: string | null;
      errorMessage?: string;
    } | null = null;
    if (reportId) {
      const w = Number(widthRaw);
      const h = Number(heightRaw);
      saveResult = await maybeInsertReportPhoto({
        reportId,
        url,
        fileName,
        width: Number.isFinite(w) ? w : null,
        height: Number.isFinite(h) ? h : null,
        imageKey,
        pointKey,
        userId,
      });
    }

    // Honest failure mode: if a reportId was provided AND the DB insert
    // did NOT succeed, we MUST NOT report 200 OK silently. The S3 object
    // is already up — we tell the client it landed but flag dbInsertSuccess
    // false (with the actual SQL error) AND respond with HTTP 500 so a
    // naive client treating 2xx as "all good" still sees a failure.
    const dbInsertSuccess = !reportId || saveResult?.saved === true;
    const responseBody = {
      url,
      key,
      path: key,
      fileName,
      contentType,
      size: file.size,
      reportPhoto: saveResult,
      dbInsertSuccess,
    };
    if (reportId && !dbInsertSuccess) {
      console.error("[api/upload] DB insert failed for report_photos", {
        reportId,
        reason: saveResult?.reason,
        errorMessage: saveResult?.errorMessage,
      });
      return Response.json(
        {
          ...responseBody,
          error: "report_photos insert failed",
          reason: saveResult?.reason || "unknown",
          errorMessage: saveResult?.errorMessage || null,
        },
        { status: 500 }
      );
    }
    return Response.json(responseBody);
  } catch (e: unknown) {
    const err = e as { message?: string };
    if (err?.message === "Unauthorized") {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/upload] error:", e);
    return Response.json({ error: "Failed to upload file" }, { status: 500 });
  }
}
