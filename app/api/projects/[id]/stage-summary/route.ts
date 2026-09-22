import pool from "../../../../../lib/db";
import { requireAuth } from "../../../../../lib/auth";
import { summarizeCategories, type StageSummaryRow } from "../../../../../lib/categoryDisplay";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type Ctx = { params: { id: string } };

function unauthorized(error: unknown) {
  return (error as { message?: string })?.message === "Unauthorized";
}

// Per-project Stage Summary override, stored as a JSON array of {label,count}.
// Created lazily so existing databases upgrade themselves on first use.
async function ensureTable() {
  await pool.query(
    `CREATE TABLE IF NOT EXISTS project_stage_summary (
       project_id VARCHAR(64) NOT NULL PRIMARY KEY,
       data_json LONGTEXT NULL,
       updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
     )`
  );
}

async function readOverride(projectId: string): Promise<StageSummaryRow[] | null> {
  const [rows] = await pool.query(
    "SELECT data_json FROM project_stage_summary WHERE project_id = ? LIMIT 1",
    [projectId]
  );
  const raw =
    Array.isArray(rows) && rows.length ? (rows[0] as { data_json?: string }).data_json : null;
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return null;
    return parsed
      .map((r) => ({ label: String(r?.label ?? "").trim(), count: Number(r?.count) || 0 }))
      .filter((r) => r.label);
  } catch {
    return null;
  }
}

// Auto-compute the Stage Summary from the project's reports (same grouping the
// export uses). No deleted_at filter — on this install deleted_at is set on
// live rows, so filtering would count nothing.
async function autoSummary(projectId: string): Promise<StageSummaryRow[]> {
  const [rows] = await pool.query(
    "SELECT category FROM reports WHERE project_id = ?",
    [projectId]
  );
  const cats = (Array.isArray(rows) ? rows : []).map((r) => (r as { category?: unknown }).category);
  return summarizeCategories(cats);
}

// GET → { rows, overridden }. Returns the saved override if present, else the
// auto-computed rows so the editor can start from them.
export async function GET(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const projectId = String(context.params?.id || "").trim();
    if (!projectId) return Response.json({ error: "Project id is required" }, { status: 400 });

    await ensureTable();
    const override = await readOverride(projectId);
    if (override && override.length) {
      return Response.json({ rows: override, overridden: true });
    }
    return Response.json({ rows: await autoSummary(projectId), overridden: false });
  } catch (error) {
    if (unauthorized(error)) return Response.json({ error: "Unauthorized" }, { status: 401 });
    console.error("[api/projects/:id/stage-summary] GET error:", error);
    return Response.json({ error: "Failed to load stage summary" }, { status: 500 });
  }
}

// PUT { rows: [{label,count}] } → saves the override.
export async function PUT(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const projectId = String(context.params?.id || "").trim();
    if (!projectId) return Response.json({ error: "Project id is required" }, { status: 400 });

    const body = await request.json().catch(() => ({} as any));
    const rowsIn: unknown[] = Array.isArray(body?.rows) ? body.rows : [];
    const rows: StageSummaryRow[] = rowsIn
      .map((r) => ({
        label: String((r as { label?: unknown })?.label ?? "").trim(),
        count: Math.max(0, Math.round(Number((r as { count?: unknown })?.count) || 0)),
      }))
      .filter((r) => r.label);
    if (!rows.length) {
      return Response.json({ error: "At least one row with a category is required" }, { status: 400 });
    }

    await ensureTable();
    const json = JSON.stringify(rows);
    await pool.query(
      `INSERT INTO project_stage_summary (project_id, data_json) VALUES (?, ?)
       ON DUPLICATE KEY UPDATE data_json = VALUES(data_json)`,
      [projectId, json]
    );
    return Response.json({ ok: true, rows, overridden: true });
  } catch (error) {
    if (unauthorized(error)) return Response.json({ error: "Unauthorized" }, { status: 401 });
    console.error("[api/projects/:id/stage-summary] PUT error:", error);
    return Response.json({ error: "Failed to save stage summary" }, { status: 500 });
  }
}

// DELETE → removes the override so the export goes back to auto-computed rows.
export async function DELETE(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const projectId = String(context.params?.id || "").trim();
    if (!projectId) return Response.json({ error: "Project id is required" }, { status: 400 });

    await ensureTable();
    await pool.query("DELETE FROM project_stage_summary WHERE project_id = ?", [projectId]);
    // Return the fresh auto rows so the editor can show them immediately.
    return Response.json({ ok: true, rows: await autoSummary(projectId), overridden: false });
  } catch (error) {
    if (unauthorized(error)) return Response.json({ error: "Unauthorized" }, { status: 401 });
    console.error("[api/projects/:id/stage-summary] DELETE error:", error);
    return Response.json({ error: "Failed to reset stage summary" }, { status: 500 });
  }
}
