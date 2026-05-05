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

/**
 * Schema-aware introspection. Captures every column on `reports` so
 * we can satisfy NOT NULL no-default columns even on installs that
 * have shape drift from the canonical schema.
 */
type ColumnMeta = {
  name: string;
  isNullable: boolean;
  hasDefault: boolean;
  isAuto: boolean;
  type: string;
};
async function getReportColumnsMeta(): Promise<ColumnMeta[]> {
  const [rows] = await pool.query("SHOW COLUMNS FROM reports");
  const out: ColumnMeta[] = [];
  for (const r of Array.isArray(rows) ? rows : []) {
    const row = r as {
      Field?: string;
      Null?: string;
      Default?: string | null;
      Extra?: string;
      Type?: string;
    };
    const name = String(row.Field || "").toLowerCase();
    if (!name) continue;
    const extra = String(row.Extra || "").toLowerCase();
    out.push({
      name,
      isNullable: String(row.Null || "").toUpperCase() === "YES",
      hasDefault: row.Default !== null && typeof row.Default !== "undefined",
      isAuto: extra.includes("auto_increment") || extra.includes("default_generated"),
      type: String(row.Type || "").toLowerCase(),
    });
  }
  return out;
}

function placeholderForType(type: string): unknown {
  // Per-type default for NOT NULL no-default columns we don't have
  // explicit data for. Numeric → 0; everything else → "" (MySQL
  // coerces appropriately for varchar/text). Date/time columns are
  // NEVER routed through this function — they get the inline NOW()
  // treatment in the INSERT builder so MySQL writes its own DATETIME.
  if (
    type.startsWith("int") ||
    type.startsWith("tinyint") ||
    type.startsWith("smallint") ||
    type.startsWith("mediumint") ||
    type.startsWith("bigint") ||
    type.startsWith("decimal") ||
    type.startsWith("numeric") ||
    type.startsWith("float") ||
    type.startsWith("double") ||
    type.startsWith("real")
  ) {
    return 0;
  }
  return "";
}

export async function POST(request: Request, context: Ctx) {
  const projectId = String(context.params?.id || "").trim();
  let body: any = {};
  let payload: Record<string, unknown> | null = null;
  try {
    const authUser = requireAuth(request);
    if (!projectId) {
      return Response.json({ error: "Project id is required" }, { status: 400 });
    }

    body = await request.json().catch(() => ({} as any));
    const colsMeta = await getReportColumnsMeta();
    const cols = new Set(colsMeta.map((c) => c.name));

    const afterReportId = String(body?.afterReportId || "").trim();
    const reportId = uuidv4();

    // ---- Resolve the "after" row when the caller used afterReportId
    // mode. We use it both as a defaults source (to satisfy NOT NULL
    // columns the body doesn't carry) AND to compute sort_order.
    let prevReport: Record<string, unknown> | null = null;
    if (afterReportId && cols.has("id") && cols.has("project_id")) {
      try {
        const [prevRows] = await pool.query(
          "SELECT * FROM reports WHERE id = ? AND project_id = ? LIMIT 1",
          [afterReportId, projectId]
        );
        if (Array.isArray(prevRows) && prevRows.length > 0) {
          prevReport = prevRows[0] as Record<string, unknown>;
        }
      } catch (err) {
        console.warn("[api/projects/[id]/reports] previous-row lookup failed:", err);
      }
    }
    const cp = (k: string) =>
      prevReport ? (prevReport as Record<string, unknown>)[k] : undefined;

    // ---- Resolve sort_order. Three modes, in priority:
    //   1. Body already includes sort_order (existing client path) →
    //      use it verbatim.
    //   2. afterReportId given AND prev row found → SHIFT every row
    //      strictly after prev.sort_order by +1, take prev+1 ourselves.
    //   3. Neither → leave null (DB will keep current default).
    let resolvedSortOrder: number | null = null;
    if (cols.has("sort_order")) {
      const bodySortOrder = Number(body?.sort_order);
      if (Number.isFinite(bodySortOrder)) {
        resolvedSortOrder = bodySortOrder;
      } else if (prevReport && typeof prevReport.sort_order !== "undefined") {
        const afterSortOrderRaw = Number(prevReport.sort_order);
        if (Number.isFinite(afterSortOrderRaw)) {
          try {
            await pool.query(
              "UPDATE reports SET sort_order = sort_order + 1 WHERE project_id = ? AND sort_order > ?",
              [projectId, afterSortOrderRaw]
            );
          } catch (shiftErr) {
            console.warn("[api/projects/[id]/reports] sort_order shift failed:", shiftErr);
          }
          resolvedSortOrder = afterSortOrderRaw + 1;
        }
      }
    }

    // ---- Build the INSERT payload from columns that ACTUALLY exist
    // on this install. Defaults copied from the previous report (when
    // the caller used afterReportId) so NOT NULL columns the body
    // doesn't carry don't blow the insert up.
    payload = {
      id: reportId,
      project_id: projectId,
      route_id: body?.route_id ?? cp("route_id") ?? null,
      category: body?.category ?? "Unknown",
      description: body?.description ?? "",
      remarks_action: body?.remarks_action ?? "",
      difficulty: body?.difficulty ?? "green",
      vehicle_movement: body?.vehicle_movement ?? cp("vehicle_movement") ?? null,
      status: body?.status ?? "active",
      created_by: authUser.id,
      user_id: body?.user_id ?? cp("user_id") ?? authUser.id,
      sort_order: resolvedSortOrder,
      // point_key & coords carry forward from the previous row when
      // present so geographic continuity is preserved; the user can
      // edit the new row in the UI afterwards.
      point_key:
        body?.point_key ??
        (cp("point_key") != null ? String(cp("point_key")) : null),
      latitude: body?.latitude ?? cp("latitude") ?? null,
      longitude: body?.longitude ?? cp("longitude") ?? null,
      loc_lat: body?.loc_lat ?? cp("loc_lat") ?? null,
      loc_lon: body?.loc_lon ?? cp("loc_lon") ?? null,
      location: body?.location ?? "",
      // created_at / updated_at are NEVER passed as parameters here.
      // MySQL DATETIME truncates ISO strings ("2026-05-04T13:53:18.523Z")
      // → ER_TRUNCATED_WRONG_VALUE. The columns are emitted INLINE as
      // NOW() further below so the DB writes its own correct datetime.
    };

    // Only emit columns that exist in the live schema AND that we have
    // a defined value for. EXCLUDE created_at / updated_at — they get
    // appended as inline NOW() placeholders (NOT parameter values).
    const keys = Object.keys(payload).filter(
      (k) =>
        cols.has(k.toLowerCase()) &&
        typeof payload![k] !== "undefined" &&
        k !== "created_at" &&
        k !== "updated_at"
    );
    const values = keys.map((k) => payload![k]);
    const placeholders: string[] = keys.map(() => "?");

    // For every NOT NULL no-default column we did NOT already fill,
    // emit a per-type placeholder so the INSERT doesn't throw with
    // "Field 'X' doesn't have a default value".
    // Skip every datetime/timestamp column here too — those go in via
    // the inline NOW() block below, never as a parameter value.
    const filled = new Set(keys.map((k) => k.toLowerCase()));
    for (const c of colsMeta) {
      if (filled.has(c.name)) continue;
      if (c.isNullable) continue;
      if (c.hasDefault) continue;
      if (c.isAuto) continue;
      if (
        c.name === "created_at" ||
        c.name === "updated_at" ||
        c.name === "timestamp"
      )
        continue;
      // Skip ALL date-ish types from the parameter path — they always
      // get the inline NOW() treatment below to avoid datetime
      // truncation.
      if (
        c.type.startsWith("datetime") ||
        c.type.startsWith("timestamp") ||
        c.type.startsWith("date")
      ) {
        continue;
      }
      const ph = placeholderForType(c.type);
      console.warn(
        `[api/projects/[id]/reports] filling NOT NULL column "${c.name}" with placeholder`,
        { type: c.type, value: ph }
      );
      keys.push(c.name);
      values.push(ph);
      placeholders.push("?");
    }

    // Inline NOW() for created_at / updated_at / timestamp when the
    // schema declares them. NO parameter value pushed — the SQL placeholder
    // IS the function call, so MySQL writes its own DATETIME without
    // any format negotiation.
    if (cols.has("created_at")) {
      keys.push("created_at");
      placeholders.push("NOW()");
    }
    if (cols.has("updated_at")) {
      keys.push("updated_at");
      placeholders.push("NOW()");
    }
    if (cols.has("timestamp")) {
      keys.push("timestamp");
      placeholders.push("NOW()");
    }
    // Datetime columns that are NOT NULL with no default and aren't
    // one of the well-known names above also need NOW() inline so the
    // INSERT succeeds.
    for (const c of colsMeta) {
      if (filled.has(c.name)) continue;
      if (keys.includes(c.name)) continue;
      if (c.isNullable) continue;
      if (c.hasDefault) continue;
      if (c.isAuto) continue;
      if (
        c.type.startsWith("datetime") ||
        c.type.startsWith("timestamp") ||
        c.type.startsWith("date")
      ) {
        keys.push(c.name);
        placeholders.push("NOW()");
        console.warn(
          `[api/projects/[id]/reports] filling NOT NULL date column "${c.name}" with NOW()`,
          { type: c.type }
        );
      }
    }

    if (!keys.includes("id") || !keys.includes("project_id")) {
      return Response.json({ error: "Reports table schema is invalid" }, { status: 500 });
    }

    const sql = `INSERT INTO reports (${keys.join(", ")}) VALUES (${placeholders.join(", ")})`;
    console.log("[ADD REPORT INSERT]", {
      projectId,
      afterReportId: afterReportId || null,
      reportId,
      resolvedSortOrder,
      columns: keys,
    });
    // Spec-mandated insert log. valuesCount MUST equal the number of "?"
    // placeholders (NOT the total placeholder count, which includes
    // NOW() literals). If they diverge, the param binding is wrong.
    console.log("[ADD REPORT INSERT SQL]", {
      keys,
      placeholders,
      valuesCount: values.length,
      questionMarkCount: placeholders.filter((p) => p === "?").length,
      values,
    });

    await pool.query(sql, values);

    const [rows] = await pool.query("SELECT * FROM reports WHERE id = ? LIMIT 1", [reportId]);
    const report = Array.isArray(rows) && rows.length ? rows[0] : null;

    return Response.json({ ok: true, report }, { status: 201 });
  } catch (error) {
    if (unauthorized(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    const e = error as {
      message?: string;
      code?: string;
      errno?: number;
      sqlState?: string;
      sqlMessage?: string;
      sql?: string;
      stack?: string;
    };
    // Spec-mandated detailed server log so the operator can see the
    // exact SQL error in the terminal.
    console.error("[ADD REPORT CREATE FAILED]", {
      projectId,
      afterReportId: body?.afterReportId || null,
      payload,
      message: e?.message,
      code: e?.code,
      errno: e?.errno,
      sqlState: e?.sqlState,
      sqlMessage: e?.sqlMessage,
      sql: e?.sql,
      stack: e?.stack,
    });
    return Response.json(
      {
        error: "Failed to create report",
        detail: e?.sqlMessage || e?.message || String(error),
        code: e?.code || null,
        sqlState: e?.sqlState || null,
      },
      { status: 500 }
    );
  }
}
