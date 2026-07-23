/* eslint-disable @typescript-eslint/no-explicit-any */
// Every project that has an uploaded finalized report, grouped by project, so
// the "Finished projects" page can list them all in one place. Static sibling
// of the [id] segment (like /api/projects/combine), so it never collides with
// /api/projects/<id>.
import pool from "../../../../lib/db";
import { requireAuth } from "../../../../lib/auth";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  try {
    requireAuth(request);

    // The table may not exist yet if nothing has been uploaded — create it so
    // the query returns cleanly instead of erroring.
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

    const [rows] = await pool.query(
      `SELECT pf.id, pf.project_id, pf.file_name, pf.size_bytes, pf.created_at,
              p.name AS project_name
         FROM project_files pf
         JOIN projects p ON p.id = pf.project_id
        WHERE pf.kind = 'finalized_report'
        ORDER BY pf.created_at DESC`
    );

    const byProject = new Map<
      string,
      { projectId: string; projectName: string; files: any[]; latest: string }
    >();
    for (const r of Array.isArray(rows) ? (rows as any[]) : []) {
      const pid = String(r.project_id);
      if (!byProject.has(pid)) {
        byProject.set(pid, {
          projectId: pid,
          projectName: String(r.project_name || "Untitled project"),
          files: [],
          latest: String(r.created_at || ""),
        });
      }
      byProject.get(pid)!.files.push({
        id: r.id,
        fileName: r.file_name,
        size: Number(r.size_bytes || 0),
        createdAt: r.created_at,
      });
    }

    // Most-recently-finalized projects first.
    const projects = Array.from(byProject.values()).sort((a, b) =>
      String(b.latest).localeCompare(String(a.latest))
    );

    return Response.json({ projects });
  } catch (error) {
    if ((error as { message?: string })?.message === "Unauthorized") {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/projects/finalized] GET error:", error);
    return Response.json({ error: "Failed to list finalized projects" }, { status: 500 });
  }
}
