/* eslint-disable @typescript-eslint/no-explicit-any */
// Finalized report files for a project: upload a .docx (stored in S3), list
// them, and delete them. Download is handled by ./download. The stored file is
// whatever the user uploads (their finalized Word report) — re-downloadable
// anytime without regenerating.
import { randomUUID } from "crypto";
import fs from "fs";
import path from "path";
import pool from "../../../../../lib/db";
import { requireAuth } from "../../../../../lib/auth";
import { uploadBufferToS3, deleteS3Object } from "../../../../../lib/s3";
import { logActivity } from "../../../../../lib/activityLog";

// Files stored on disk when S3 is unavailable (e.g. local dev with placeholder
// AWS keys). Kept OUTSIDE public/ and served only through the download route.
export const LOCAL_ROOT = path.join(process.cwd(), ".uploads", "finalized");

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 300;

type Ctx = { params: { id: string } };

function unauthorized(error: unknown) {
  return (error as { message?: string })?.message === "Unauthorized";
}

async function ensureTable() {
  await pool.query(
    `CREATE TABLE IF NOT EXISTS project_files (
       id VARCHAR(64) NOT NULL PRIMARY KEY,
       project_id VARCHAR(64) NOT NULL,
       kind VARCHAR(32) NOT NULL DEFAULT 'finalized_report',
       file_name VARCHAR(255) NOT NULL,
       s3_key VARCHAR(512) NOT NULL,
       content_type VARCHAR(160) NULL,
       size_bytes BIGINT NULL,
       uploaded_by VARCHAR(64) NULL,
       created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
       KEY idx_project (project_id)
     ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4`
  );
}

function safeName(name: string) {
  return String(name || "finalized-report.docx")
    .replace(/[^a-zA-Z0-9._-]/g, "-")
    .replace(/-+/g, "-")
    .slice(0, 180);
}

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

export async function POST(request: Request, context: Ctx) {
  try {
    const user = requireAuth(request) as { id?: string };
    const projectId = String(context.params?.id || "").trim();
    if (!projectId) return Response.json({ error: "Project id is required" }, { status: 400 });

    const form = await request.formData();
    const file = form.get("file");
    if (!(file instanceof File)) {
      return Response.json({ error: "No file uploaded" }, { status: 400 });
    }
    const originalName = safeName(file.name || "finalized-report.docx");
    if (!/\.docx$/i.test(originalName)) {
      return Response.json(
        { error: "Please upload a Word (.docx) file — the same format as the export." },
        { status: 400 }
      );
    }

    // Confirm the project exists.
    const [projRows] = await pool.query("SELECT id FROM projects WHERE id = ? LIMIT 1", [projectId]);
    if (!Array.isArray(projRows) || !projRows.length) {
      return Response.json({ error: "Project not found" }, { status: 404 });
    }

    const bytes = Buffer.from(await file.arrayBuffer());
    const uuid = randomUUID();
    const key = `finalized/${projectId}/${uuid}-${originalName}`;
    // Try S3 first; if it's not available (local dev / bad creds / outage), fall
    // back to disk. storageKey is either the S3 key or "local:<absolute path>".
    let storageKey = key;
    try {
      await uploadBufferToS3({ key, body: bytes, contentType: file.type || DOCX_MIME });
    } catch (e: any) {
      try {
        const dir = path.join(LOCAL_ROOT, projectId);
        fs.mkdirSync(dir, { recursive: true });
        const localPath = path.join(dir, `${uuid}-${originalName}`);
        fs.writeFileSync(localPath, bytes);
        storageKey = "local:" + localPath;
        console.warn("[api/projects/:id/finalized] S3 unavailable — stored on disk:", localPath);
      } catch (localErr) {
        console.error("[api/projects/:id/finalized] S3 AND local storage failed:", e, localErr);
        return Response.json(
          { error: "Could not store the file (S3 and disk both failed).", detail: e?.message },
          { status: 502 }
        );
      }
    }

    await ensureTable();
    const id = randomUUID();
    await pool.query(
      `INSERT INTO project_files
         (id, project_id, kind, file_name, s3_key, content_type, size_bytes, uploaded_by)
       VALUES (?, ?, 'finalized_report', ?, ?, ?, ?, ?)`,
      [id, projectId, originalName, storageKey, file.type || DOCX_MIME, bytes.length, user?.id || null]
    );

    await logActivity(request, {
      action: "upload_finalized",
      table: "project_files",
      entityId: id,
      projectId,
      details: { fileName: originalName, size: bytes.length },
    }).catch(() => {});

    return Response.json({
      ok: true,
      file: { id, fileName: originalName, size: bytes.length, createdAt: new Date().toISOString() },
    });
  } catch (error) {
    if (unauthorized(error)) return Response.json({ error: "Unauthorized" }, { status: 401 });
    console.error("[api/projects/:id/finalized] POST error:", error);
    return Response.json({ error: "Failed to upload finalized report" }, { status: 500 });
  }
}

export async function GET(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const projectId = String(context.params?.id || "").trim();
    if (!projectId) return Response.json({ error: "Project id is required" }, { status: 400 });

    await ensureTable();
    const [rows] = await pool.query(
      `SELECT id, file_name, size_bytes, content_type, uploaded_by, created_at
         FROM project_files
        WHERE project_id = ? AND kind = 'finalized_report'
        ORDER BY created_at DESC`,
      [projectId]
    );
    const files = (Array.isArray(rows) ? rows : []).map((r: any) => ({
      id: r.id,
      fileName: r.file_name,
      size: Number(r.size_bytes || 0),
      contentType: r.content_type,
      createdAt: r.created_at,
    }));
    return Response.json({ files });
  } catch (error) {
    if (unauthorized(error)) return Response.json({ error: "Unauthorized" }, { status: 401 });
    console.error("[api/projects/:id/finalized] GET error:", error);
    return Response.json({ error: "Failed to list finalized reports" }, { status: 500 });
  }
}

export async function DELETE(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const projectId = String(context.params?.id || "").trim();
    const url = new URL(request.url);
    const fileId = String(url.searchParams.get("fileId") || "").trim();
    if (!fileId) return Response.json({ error: "fileId is required" }, { status: 400 });

    await ensureTable();
    const [rows] = await pool.query(
      "SELECT s3_key FROM project_files WHERE id = ? AND project_id = ? LIMIT 1",
      [fileId, projectId]
    );
    const row = Array.isArray(rows) && rows.length ? (rows[0] as any) : null;
    if (!row) return Response.json({ error: "File not found" }, { status: 404 });

    const storageKey = String(row.s3_key || "");
    try {
      if (storageKey.startsWith("local:")) {
        fs.unlinkSync(storageKey.slice("local:".length));
      } else {
        await deleteS3Object(storageKey);
      }
    } catch (e) {
      console.warn("[api/projects/:id/finalized] storage delete failed (removing row anyway):", e);
    }
    await pool.query("DELETE FROM project_files WHERE id = ? AND project_id = ?", [fileId, projectId]);
    return Response.json({ ok: true });
  } catch (error) {
    if (unauthorized(error)) return Response.json({ error: "Unauthorized" }, { status: 401 });
    console.error("[api/projects/:id/finalized] DELETE error:", error);
    return Response.json({ error: "Failed to delete finalized report" }, { status: 500 });
  }
}
