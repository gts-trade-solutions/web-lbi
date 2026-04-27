import { v4 as uuidv4 } from "uuid";
import pool from "../../../../../lib/db";
import { requireAuth } from "../../../../../lib/auth";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type DbColumnRow = { Field?: string };
type Ctx = { params: { id: string } };

async function getColumns(table: string) {
  const [rows] = await pool.query(`SHOW COLUMNS FROM ${table}`);
  return new Set(
    (Array.isArray(rows) ? rows : []).map((r) =>
      String((r as DbColumnRow).Field || "").toLowerCase()
    )
  );
}

function has(cols: Set<string>, col: string) {
  return cols.has(col.toLowerCase());
}

function toDirection(value: string | null) {
  return String(value || "").toLowerCase() === "desc" ? "DESC" : "ASC";
}

function unauthorized(error: unknown) {
  return (error as { message?: string })?.message === "Unauthorized";
}

export async function GET(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const projectId = String(context.params?.id || "").trim();
    if (!projectId) {
      return Response.json({ error: "Project id is required" }, { status: 400 });
    }

    const cols = await getColumns("reports");
    const url = new URL(request.url);
    const search = String(url.searchParams.get("search") || "").trim();
    const difficulty = String(url.searchParams.get("difficulty") || "").trim().toLowerCase();
    const sort = toDirection(url.searchParams.get("sort"));
    const limitRaw = Number(url.searchParams.get("limit") || 0);
    const offsetRaw = Number(url.searchParams.get("offset") || 0);
    const limit = Number.isFinite(limitRaw) && limitRaw > 0 ? Math.min(limitRaw, 1000) : 0;
    const offset = Number.isFinite(offsetRaw) && offsetRaw > 0 ? offsetRaw : 0;

    const where: string[] = ["project_id = ?"];
    const args: unknown[] = [projectId];

    if (difficulty === "unset") {
      if (has(cols, "difficulty")) where.push("(difficulty IS NULL OR difficulty = '')");
    } else if (difficulty && difficulty !== "all" && has(cols, "difficulty")) {
      where.push("difficulty = ?");
      args.push(difficulty);
    }

    if (search) {
      const searchable = ["category", "description", "remarks_action", "point_key", "id"].filter((c) =>
        has(cols, c)
      );
      if (searchable.length) {
        const like = `%${search}%`;
        where.push(`(${searchable.map((c) => `LOWER(${c}) LIKE LOWER(?)`).join(" OR ")})`);
        searchable.forEach(() => args.push(like));
      }
    }

    const orderParts: string[] = [];
    if (has(cols, "sort_order")) orderParts.push(`sort_order ${sort}`);
    if (has(cols, "created_at")) orderParts.push(`created_at ${sort}`);
    if (!orderParts.length) orderParts.push("id ASC");

    const limitSql = limit ? " LIMIT ? OFFSET ?" : "";
    if (limit) {
      args.push(limit);
      args.push(offset);
    }

    const [rows] = await pool.query(
      `SELECT * FROM reports WHERE ${where.join(" AND ")} ORDER BY ${orderParts.join(", ")}${limitSql}`,
      args
    );

    return Response.json({ reports: Array.isArray(rows) ? rows : [] });
  } catch (error) {
    if (unauthorized(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/projects/[id]/reports] GET error:", error);
    return Response.json({ error: "Failed to fetch reports" }, { status: 500 });
  }
}

export async function POST(request: Request, context: Ctx) {
  try {
    const authUser = requireAuth(request);
    const projectId = String(context.params?.id || "").trim();
    if (!projectId) {
      return Response.json({ error: "Project id is required" }, { status: 400 });
    }

    const body = await request.json().catch(() => ({} as any));
    const cols = await getColumns("reports");
    const reportId = uuidv4();

    const payload: Record<string, unknown> = {
      id: reportId,
      project_id: projectId,
      route_id: body?.route_id ?? null,
      category: body?.category ?? null,
      description: body?.description ?? null,
      remarks_action: body?.remarks_action ?? null,
      difficulty: body?.difficulty ?? null,
      sort_order: body?.sort_order ?? null,
      status: body?.status ?? "active",
      created_by: authUser.id,
      user_id: body?.user_id ?? authUser.id,
      created_at: body?.created_at ?? undefined,
    };

    const keys = Object.keys(payload).filter((k) => has(cols, k) && typeof payload[k] !== "undefined");
    const values = keys.map((k) => payload[k]);

    if (!keys.includes("id") || !keys.includes("project_id")) {
      return Response.json({ error: "Reports table schema is invalid" }, { status: 500 });
    }

    await pool.query(
      `INSERT INTO reports (${keys.join(", ")}) VALUES (${keys.map(() => "?").join(", ")})`,
      values
    );

    const [rows] = await pool.query("SELECT * FROM reports WHERE id = ? LIMIT 1", [reportId]);
    const report = Array.isArray(rows) && rows.length ? rows[0] : null;

    return Response.json({ report }, { status: 201 });
  } catch (error) {
    if (unauthorized(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/projects/[id]/reports] POST error:", error);
    return Response.json({ error: "Failed to create report" }, { status: 500 });
  }
}
