import pool from "../../../../../../lib/db";
import { requireAuth } from "../../../../../../lib/auth";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type Ctx = { params: { id: string; key: string } };

function unauthorized(error: unknown) {
  return (error as { message?: string })?.message === "Unauthorized";
}

// Section keys the editor is allowed to write. Keeps the generic store tidy and
// prevents arbitrary rows.
const ALLOWED_KEYS = new Set(["observation-summary", "hindrance", "stage-image"]);

// One generic, lazily-created table holds the editable extras for a project's
// report (Summary of Observations, Hindrance Solutions, Stage-Summary image).
// Created on first use so existing databases upgrade themselves — no migration.
async function ensureTable() {
  await pool.query(
    `CREATE TABLE IF NOT EXISTS project_report_sections (
       project_id VARCHAR(64) NOT NULL,
       section_key VARCHAR(40) NOT NULL,
       data_json LONGTEXT NULL,
       updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
       PRIMARY KEY (project_id, section_key)
     )`
  );
}

async function readSection(projectId: string, key: string): Promise<any | null> {
  const [rows] = await pool.query(
    "SELECT data_json FROM project_report_sections WHERE project_id = ? AND section_key = ? LIMIT 1",
    [projectId, key]
  );
  const raw =
    Array.isArray(rows) && rows.length ? (rows[0] as { data_json?: string }).data_json : null;
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

// GET → { data } (the saved JSON, or null if nothing saved yet).
export async function GET(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const projectId = String(context.params?.id || "").trim();
    const key = String(context.params?.key || "").trim();
    if (!projectId) return Response.json({ error: "Project id is required" }, { status: 400 });
    if (!ALLOWED_KEYS.has(key)) return Response.json({ error: "Unknown section" }, { status: 400 });

    await ensureTable();
    const data = await readSection(projectId, key);
    return Response.json({ data, overridden: data != null });
  } catch (error) {
    if (unauthorized(error)) return Response.json({ error: "Unauthorized" }, { status: 401 });
    console.error("[api/projects/:id/report-sections/:key] GET error:", error);
    return Response.json({ error: "Failed to load section" }, { status: 500 });
  }
}

// PUT { data } → saves the section JSON verbatim.
export async function PUT(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const projectId = String(context.params?.id || "").trim();
    const key = String(context.params?.key || "").trim();
    if (!projectId) return Response.json({ error: "Project id is required" }, { status: 400 });
    if (!ALLOWED_KEYS.has(key)) return Response.json({ error: "Unknown section" }, { status: 400 });

    const body = await request.json().catch(() => ({} as any));
    const data = body?.data ?? null;

    await ensureTable();
    if (data == null) {
      await pool.query(
        "DELETE FROM project_report_sections WHERE project_id = ? AND section_key = ?",
        [projectId, key]
      );
      return Response.json({ ok: true, data: null, overridden: false });
    }
    const json = JSON.stringify(data);
    await pool.query(
      `INSERT INTO project_report_sections (project_id, section_key, data_json) VALUES (?, ?, ?)
       ON DUPLICATE KEY UPDATE data_json = VALUES(data_json)`,
      [projectId, key, json]
    );
    return Response.json({ ok: true, data, overridden: true });
  } catch (error) {
    if (unauthorized(error)) return Response.json({ error: "Unauthorized" }, { status: 401 });
    console.error("[api/projects/:id/report-sections/:key] PUT error:", error);
    return Response.json({ error: "Failed to save section" }, { status: 500 });
  }
}

// DELETE → removes the override so the export falls back to its default.
export async function DELETE(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const projectId = String(context.params?.id || "").trim();
    const key = String(context.params?.key || "").trim();
    if (!projectId) return Response.json({ error: "Project id is required" }, { status: 400 });
    if (!ALLOWED_KEYS.has(key)) return Response.json({ error: "Unknown section" }, { status: 400 });

    await ensureTable();
    await pool.query(
      "DELETE FROM project_report_sections WHERE project_id = ? AND section_key = ?",
      [projectId, key]
    );
    return Response.json({ ok: true, data: null, overridden: false });
  } catch (error) {
    if (unauthorized(error)) return Response.json({ error: "Unauthorized" }, { status: 401 });
    console.error("[api/projects/:id/report-sections/:key] DELETE error:", error);
    return Response.json({ error: "Failed to reset section" }, { status: 500 });
  }
}
