import { v4 as uuidv4 } from "uuid";
import pool from "../../../lib/db";
import { requireAuth } from "../../../lib/auth";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type DbColumnRow = { Field?: string };

async function getProjectColumns() {
  const [rows] = await pool.query("SHOW COLUMNS FROM projects");
  const cols = new Set(
    (Array.isArray(rows) ? rows : []).map((r) =>
      String((r as DbColumnRow).Field || "").toLowerCase()
    )
  );
  return cols;
}

function isUnauthorizedError(e: unknown) {
  return (e as { message?: string })?.message === "Unauthorized";
}

export async function GET(request: Request) {
  try {
    requireAuth(request);
    const columns = await getProjectColumns();

    const orderBy = columns.has("created_at")
      ? " ORDER BY created_at DESC"
      : columns.has("updated_at")
        ? " ORDER BY updated_at DESC"
        : "";

    const [rows] = await pool.query(`SELECT * FROM projects${orderBy}`);
    const projects = Array.isArray(rows) ? rows : [];

    return Response.json({ projects });
  } catch (error) {
    if (isUnauthorizedError(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/projects] GET error:", error);
    return Response.json({ error: "Failed to fetch projects" }, { status: 500 });
  }
}

export async function POST(request: Request) {
  try {
    const authUser = requireAuth(request);
    const body = await request.json().catch(() => ({} as any));
    const name = String(body?.name || "").trim();
    const description = String(body?.description || "").trim() || null;
    const clientName = String(body?.client_name || "").trim() || null;
    const location = String(body?.location || "").trim() || null;
    const status = String(body?.status || "").trim() || "active";

    if (!name) {
      return Response.json({ error: "Project name is required" }, { status: 400 });
    }

    const columns = await getProjectColumns();
    const projectId = uuidv4();

    const payload: Record<string, unknown> = {
      id: projectId,
      name,
      description,
      client_name: clientName,
      location,
      status,
      created_by: authUser.id,
      last_modified_by: authUser.id,
      user_id: authUser.id,
    };

    const keys = Object.keys(payload).filter((k) => columns.has(k.toLowerCase()));
    const values = keys.map((k) => payload[k]);

    if (!keys.includes("id")) {
      return Response.json({ error: "Projects table is missing id column" }, { status: 500 });
    }

    const placeholders = keys.map(() => "?").join(", ");
    await pool.query(
      `INSERT INTO projects (${keys.join(", ")}) VALUES (${placeholders})`,
      values
    );

    const [rows] = await pool.query("SELECT * FROM projects WHERE id = ? LIMIT 1", [projectId]);
    const project = Array.isArray(rows) && rows.length ? rows[0] : null;

    return Response.json({ project }, { status: 201 });
  } catch (error) {
    if (isUnauthorizedError(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/projects] POST error:", error);
    return Response.json({ error: "Failed to create project" }, { status: 500 });
  }
}

