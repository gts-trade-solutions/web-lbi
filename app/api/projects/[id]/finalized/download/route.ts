/* eslint-disable @typescript-eslint/no-explicit-any */
// Return a short-lived signed S3 URL for a finalized report, forcing a download
// with the original file name. The client fetches this (authenticated), then
// navigates to the URL to pull the .docx straight from S3.
import fs from "fs";
import path from "path";
import { GetObjectCommand } from "@aws-sdk/client-s3";
import { getSignedUrl } from "@aws-sdk/s3-request-presigner";
import pool from "../../../../../../lib/db";
import { requireAuth } from "../../../../../../lib/auth";
import { s3Client, getBucketName } from "../../../../../../lib/s3";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const LOCAL_ROOT = path.join(process.cwd(), ".uploads", "finalized");
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

type Ctx = { params: { id: string } };

export async function GET(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const projectId = String(context.params?.id || "").trim();
    const url = new URL(request.url);
    const fileId = String(url.searchParams.get("fileId") || "").trim();
    if (!projectId || !fileId) {
      return Response.json({ error: "projectId and fileId are required" }, { status: 400 });
    }

    const [rows] = await pool.query(
      "SELECT file_name, s3_key, content_type FROM project_files WHERE id = ? AND project_id = ? LIMIT 1",
      [fileId, projectId]
    );
    const row = Array.isArray(rows) && rows.length ? (rows[0] as any) : null;
    if (!row) return Response.json({ error: "File not found" }, { status: 404 });

    const fileName = String(row.file_name || "finalized-report.docx");
    const storageKey = String(row.s3_key || "");
    const contentType = row.content_type || DOCX_MIME;

    // Locally-stored file (S3 was unavailable at upload time): stream from disk.
    if (storageKey.startsWith("local:")) {
      const abs = path.resolve(storageKey.slice("local:".length));
      if (!abs.startsWith(LOCAL_ROOT) || !fs.existsSync(abs)) {
        return Response.json({ error: "File not found" }, { status: 404 });
      }
      const bytes = fs.readFileSync(abs);
      return new Response(bytes, {
        headers: {
          "Content-Type": contentType,
          "Content-Disposition": `attachment; filename="${fileName.replace(/"/g, "")}"`,
          "Content-Length": String(bytes.length),
        },
      });
    }

    const cmd = new GetObjectCommand({
      Bucket: getBucketName(),
      Key: String(row.s3_key),
      ResponseContentDisposition: `attachment; filename="${fileName.replace(/"/g, "")}"`,
      ResponseContentType:
        row.content_type ||
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    const signedUrl = await getSignedUrl(s3Client, cmd, { expiresIn: 600 });
    return Response.json({ url: signedUrl, fileName });
  } catch (error) {
    if ((error as { message?: string })?.message === "Unauthorized") {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/projects/:id/finalized/download] error:", error);
    return Response.json({ error: "Failed to prepare download" }, { status: 500 });
  }
}
