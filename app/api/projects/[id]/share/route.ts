/* eslint-disable @typescript-eslint/no-explicit-any */
// Manage password-protected share links for a project's animated route report.
//   POST   -> create a link (returns /share/<token>)
//   GET    -> list this project's links
//   DELETE -> revoke a link (?shareId=...)
import { v4 as uuidv4 } from "uuid";
import bcrypt from "bcryptjs";
import pool from "../../../../../lib/db";
import { requireAuth } from "../../../../../lib/auth";
import { logActivity } from "../../../../../lib/activityLog";
import {
  ensureShareTable,
  getShareByToken,
  makeShareToken,
  parseSelectedIds,
} from "../../../../../lib/shareLink";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type Ctx = { params: { id: string } };

function unauthorized(error: unknown) {
  return (error as { message?: string })?.message === "Unauthorized";
}

export async function POST(request: Request, context: Ctx) {
  try {
    const user = requireAuth(request) as { id?: string };
    const projectId = String(context.params?.id || "").trim();
    if (!projectId) return Response.json({ error: "Project id is required" }, { status: 400 });

    const body = await request.json().catch(() => ({} as any));
    const password = String(body?.password || "");
    if (password.length < 4) {
      return Response.json(
        { error: "Password must be at least 4 characters." },
        { status: 400 }
      );
    }
    const title = String(body?.title || "").trim().slice(0, 240) || null;
    const selectedIds = Array.from(
      new Set(
        (Array.isArray(body?.selectedIds) ? body.selectedIds : [])
          .map((x: unknown) => String(x || "").trim())
          .filter(Boolean)
      )
    );

    // Confirm the project exists (and grab its name for a friendly response).
    const [projRows] = await pool.query("SELECT id, name FROM projects WHERE id = ? LIMIT 1", [
      projectId,
    ]);
    const project = Array.isArray(projRows) && projRows.length ? (projRows[0] as any) : null;
    if (!project) return Response.json({ error: "Project not found" }, { status: 404 });

    await ensureShareTable();
    let token = makeShareToken();
    if (await getShareByToken(token)) token = makeShareToken(); // astronomically unlikely retry

    const hash = await bcrypt.hash(password, 10);
    const id = uuidv4();
    await pool.query(
      `INSERT INTO report_shares
         (id, token, project_id, password_hash, title, selected_ids, created_by, revoked)
       VALUES (?, ?, ?, ?, ?, ?, ?, 0)`,
      [
        id,
        token,
        projectId,
        hash,
        title,
        selectedIds.length ? JSON.stringify(selectedIds) : null,
        user?.id || null,
      ]
    );

    await logActivity(request, {
      action: "create_share",
      table: "report_shares",
      entityId: id,
      projectId,
      details: { title, reportCount: selectedIds.length || "all" },
    }).catch(() => {});

    return Response.json({
      ok: true,
      token,
      path: `/share/${token}`,
      projectName: String(project.name || ""),
    });
  } catch (error) {
    if (unauthorized(error)) return Response.json({ error: "Unauthorized" }, { status: 401 });
    console.error("[api/projects/:id/share] POST error:", error);
    return Response.json({ error: "Failed to create share link" }, { status: 500 });
  }
}

export async function GET(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const projectId = String(context.params?.id || "").trim();
    if (!projectId) return Response.json({ error: "Project id is required" }, { status: 400 });

    await ensureShareTable();
    const [rows] = await pool.query(
      `SELECT id, token, title, selected_ids, created_at, revoked
         FROM report_shares WHERE project_id = ? ORDER BY created_at DESC`,
      [projectId]
    );
    const shares = (Array.isArray(rows) ? rows : []).map((r: any) => {
      const selectedIds = parseSelectedIds(r.selected_ids);
      return {
        id: r.id,
        token: r.token,
        title: r.title,
        createdAt: r.created_at,
        revoked: !!r.revoked,
        reportCount: selectedIds.length,
        // The exact points this link shows ([] = all points in the project).
        selectedIds,
        path: `/share/${r.token}`,
      };
    });
    return Response.json({ shares });
  } catch (error) {
    if (unauthorized(error)) return Response.json({ error: "Unauthorized" }, { status: 401 });
    console.error("[api/projects/:id/share] GET error:", error);
    return Response.json({ error: "Failed to list share links" }, { status: 500 });
  }
}

// Edit a link in place (same token / URL): title, password and/or which points
// it shows. Nothing changes the link itself, so a client's URL keeps working.
export async function PATCH(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const projectId = String(context.params?.id || "").trim();
    const url = new URL(request.url);
    const shareId = String(url.searchParams.get("shareId") || "").trim();
    if (!shareId) return Response.json({ error: "shareId is required" }, { status: 400 });

    const body = await request.json().catch(() => ({} as any));
    const sets: string[] = [];
    const args: unknown[] = [];

    if (body?.title !== undefined) {
      sets.push("title = ?");
      args.push(String(body.title || "").trim().slice(0, 240) || null);
    }
    if (body?.password !== undefined && String(body.password).length > 0) {
      if (String(body.password).length < 4) {
        return Response.json({ error: "Password must be at least 4 characters." }, { status: 400 });
      }
      sets.push("password_hash = ?");
      args.push(await bcrypt.hash(String(body.password), 10));
    }
    if (Array.isArray(body?.selectedIds)) {
      const ids = Array.from(
        new Set(body.selectedIds.map((x: unknown) => String(x || "").trim()).filter(Boolean))
      );
      sets.push("selected_ids = ?");
      args.push(ids.length ? JSON.stringify(ids) : null);
    }
    if (!sets.length) {
      return Response.json({ error: "Nothing to update." }, { status: 400 });
    }

    await ensureShareTable();
    args.push(shareId, projectId);
    await pool.query(
      `UPDATE report_shares SET ${sets.join(", ")} WHERE id = ? AND project_id = ?`,
      args
    );
    await logActivity(request, {
      action: "edit_share",
      table: "report_shares",
      entityId: shareId,
      projectId,
    }).catch(() => {});
    return Response.json({ ok: true });
  } catch (error) {
    if (unauthorized(error)) return Response.json({ error: "Unauthorized" }, { status: 401 });
    console.error("[api/projects/:id/share] PATCH error:", error);
    return Response.json({ error: "Failed to update share link" }, { status: 500 });
  }
}

export async function DELETE(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const projectId = String(context.params?.id || "").trim();
    const url = new URL(request.url);
    const shareId = String(url.searchParams.get("shareId") || "").trim();
    if (!shareId) return Response.json({ error: "shareId is required" }, { status: 400 });

    await ensureShareTable();
    // Permanently remove the link (the client's URL stops working).
    await pool.query(
      "DELETE FROM report_shares WHERE id = ? AND project_id = ?",
      [shareId, projectId]
    );
    await logActivity(request, {
      action: "delete_share",
      table: "report_shares",
      entityId: shareId,
      projectId,
    }).catch(() => {});
    return Response.json({ ok: true });
  } catch (error) {
    if (unauthorized(error)) return Response.json({ error: "Unauthorized" }, { status: 401 });
    console.error("[api/projects/:id/share] DELETE error:", error);
    return Response.json({ error: "Failed to delete share link" }, { status: 500 });
  }
}
