import { v4 as uuidv4 } from "uuid";
import pool from "../../../../../lib/db";
import { requireAuth } from "../../../../../lib/auth";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type DbColumnRow = {
  Field?: string;
  Null?: string;
  Default?: string | null;
  Extra?: string;
  Type?: string;
};
type Ctx = { params: { id: string } };

type ColumnMeta = {
  name: string;
  isNullable: boolean;
  hasDefault: boolean;
  isAuto: boolean;
};

async function getColumns() {
  const [rows] = await pool.query("SHOW COLUMNS FROM report_photos");
  return new Set(
    (Array.isArray(rows) ? rows : []).map((r) =>
      String((r as DbColumnRow).Field || "").toLowerCase()
    )
  );
}

async function getColumnsMeta(): Promise<ColumnMeta[]> {
  const [rows] = await pool.query("SHOW COLUMNS FROM report_photos");
  const out: ColumnMeta[] = [];
  for (const r of Array.isArray(rows) ? rows : []) {
    const row = r as DbColumnRow;
    const name = String(row.Field || "").toLowerCase();
    if (!name) continue;
    const extra = String(row.Extra || "").toLowerCase();
    out.push({
      name,
      isNullable: String(row.Null || "").toUpperCase() === "YES",
      hasDefault: row.Default !== null && typeof row.Default !== "undefined",
      isAuto: extra.includes("auto_increment") || extra.includes("default_generated"),
    });
  }
  return out;
}

function unauthorized(error: unknown) {
  return (error as { message?: string })?.message === "Unauthorized";
}

export async function GET(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const reportId = String(context.params?.id || "").trim();
    if (!reportId) return Response.json({ error: "Report id is required" }, { status: 400 });

    const [rows] = await pool.query(
      "SELECT * FROM report_photos WHERE report_id = ? ORDER BY created_at ASC",
      [reportId]
    );
    return Response.json({ photos: Array.isArray(rows) ? rows : [] });
  } catch (error) {
    if (unauthorized(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/reports/:id/photos] GET error:", error);
    return Response.json({ error: "Failed to fetch report photos" }, { status: 500 });
  }
}

export async function POST(request: Request, context: Ctx) {
  const reportId = String(context.params?.id || "").trim();
  let body: any = {};
  let lastInsertedRow: Record<string, unknown> | null = null;
  try {
    const authUser = requireAuth(request);
    if (!reportId) return Response.json({ error: "Report id is required" }, { status: 400 });

    console.log("[save report photo] reportId:", reportId);

    // Validate that the parent report row actually exists. Without this guard
    // a buggy client could insert orphan photos with a bogus report_id that
    // never joins against reports during export.
    const [reportRows] = await pool.query(
      "SELECT id, project_id, point_key, user_id FROM reports WHERE id = ? LIMIT 1",
      [reportId]
    );
    const parentReport =
      Array.isArray(reportRows) && reportRows.length
        ? (reportRows[0] as Record<string, unknown>)
        : null;
    if (!parentReport) {
      console.warn("[save report photo] report not found:", reportId);
      return Response.json({ error: "Report not found", reportId }, { status: 404 });
    }

    body = await request.json().catch(() => ({} as any));
    const list = Array.isArray(body?.photos) ? body.photos : body ? [body] : [];
    const rowsIn = list.filter(
      (x: any) => x && typeof x === "object" && String(x.url || "").trim()
    );

    // Spec: empty file list is NOT an error. The caller may have just
    // added a blank report with no images selected — return ok with
    // empty photos so the UI doesn't show a misleading failure.
    if (!rowsIn.length) {
      console.log("[save report photo] no photos provided — returning ok with empty list");
      return Response.json({ ok: true, message: "No photos to save", photos: [] });
    }

    // Schema-aware: only emit columns that exist on this install. The
    // canonical schema is (id, report_id, url, width, height, created_at,
    // user_id, file_name, point_key, image_key) — NO `path`, NO `bucket`,
    // NO `updated_at`, NO `created_by`.
    const colsMeta = await getColumnsMeta();
    const cols = new Set(colsMeta.map((c) => c.name));
    const colByName = new Map(colsMeta.map((c) => [c.name, c]));

    // Should we inline NOW() for created_at? Only when the column
    // exists AND has no default AND is NOT NULL. Otherwise omit and
    // let MySQL fill from DEFAULT CURRENT_TIMESTAMP.
    const createdAtCol = colByName.get("created_at");
    const createdAtNeedsNow =
      !!createdAtCol &&
      !createdAtCol.hasDefault &&
      !createdAtCol.isAuto &&
      !createdAtCol.isNullable;

    // Build per-row inserts (NOT a multi-row insert) so we can identify
    // which row failed if only one violates a constraint.
    const inserted: Record<string, unknown>[] = [];
    const failed: Array<{
      index: number;
      file_name: string | null;
      url: string;
      error: string;
      code?: string | null;
      sqlState?: string | null;
    }> = [];

    for (let i = 0; i < rowsIn.length; i += 1) {
      const p = rowsIn[i];
      const row: Record<string, unknown> = {
        id: p?.id || uuidv4(),
        report_id: reportId,
        url: String(p.url || "").trim(),
      };
      if (cols.has("file_name")) {
        row.file_name =
          p?.file_name ??
          p?.fileName ??
          (typeof p?.key === "string" ? p.key.split("/").pop() : null) ??
          (typeof row.url === "string" ? row.url.split("/").pop() : null) ??
          null;
      }
      if (cols.has("width")) row.width = p?.width ?? null;
      if (cols.has("height")) row.height = p?.height ?? null;
      if (cols.has("user_id")) {
        row.user_id =
          p?.user_id ??
          authUser?.id ??
          (parentReport && (parentReport as { user_id?: string }).user_id) ??
          null;
      }
      if (cols.has("point_key")) {
        row.point_key =
          p?.point_key ??
          p?.pointKey ??
          (parentReport && (parentReport as { point_key?: string }).point_key) ??
          null;
      }
      if (cols.has("image_key")) {
        row.image_key = p?.image_key ?? p?.imageKey ?? null;
      }

      // Build keys+placeholders. created_at gets inline NOW() (NOT a
      // parameter) so MySQL never sees an ISO string for a DATETIME.
      const keys = Object.keys(row);
      const placeholders: string[] = keys.map(() => "?");
      const values = keys.map((k) => row[k]);
      if (createdAtNeedsNow) {
        keys.push("created_at");
        placeholders.push("NOW()");
      }

      const sql = `INSERT INTO report_photos (${keys.join(", ")}) VALUES (${placeholders.join(", ")})`;
      console.log("[SAVE REPORT PHOTO INSERT START]", {
        reportId,
        index: i,
        fileName: row.file_name ?? null,
        url: row.url,
        pointKey: row.point_key ?? null,
        imageKey: row.image_key ?? null,
        keys,
        placeholders,
        valuesCount: values.length,
        questionMarkCount: placeholders.filter((p2) => p2 === "?").length,
      });
      try {
        await pool.query(sql, values);
        inserted.push(row);
        lastInsertedRow = row;
      } catch (insertErr) {
        const e = insertErr as {
          message?: string;
          code?: string;
          sqlState?: string;
          sqlMessage?: string;
          sql?: string;
        };
        console.error("[SAVE REPORT PHOTO INSERT FAILED]", {
          reportId,
          index: i,
          file_name: row.file_name ?? null,
          url: row.url,
          message: e?.message,
          code: e?.code,
          sqlState: e?.sqlState,
          sqlMessage: e?.sqlMessage,
          sql: e?.sql,
        });
        failed.push({
          index: i,
          file_name: (row.file_name as string) ?? null,
          url: row.url as string,
          error: e?.sqlMessage || e?.message || String(insertErr),
          code: e?.code || null,
          sqlState: e?.sqlState || null,
        });
      }
    }

    // Verify what landed.
    const [verifyRows] = await pool.query(
      `SELECT id, report_id, url${cols.has("file_name") ? ", file_name" : ""}${
        cols.has("point_key") ? ", point_key" : ""
      }${cols.has("image_key") ? ", image_key" : ""}${
        cols.has("created_at") ? ", created_at" : ""
      }
       FROM report_photos
       WHERE report_id = ?
       ORDER BY ${cols.has("created_at") ? "created_at" : "id"} ASC`,
      [reportId]
    );
    console.log("[SAVE REPORT PHOTOS VERIFY]", {
      reportId,
      attemptedCount: rowsIn.length,
      insertedCount: inserted.length,
      failedCount: failed.length,
      totalAfterInsert: Array.isArray(verifyRows) ? verifyRows.length : 0,
    });

    // If EVERY row failed, propagate the first failure as 500 with detail.
    if (inserted.length === 0 && failed.length > 0) {
      return Response.json(
        {
          error: "Failed to save report photos",
          detail: failed[0].error,
          code: failed[0].code,
          sqlState: failed[0].sqlState,
          failedCount: failed.length,
          failed,
        },
        { status: 500 }
      );
    }

    // Partial success returns 207-style payload but HTTP 200 so the
    // client can decide what to do with it.
    return Response.json(
      {
        ok: true,
        photos: Array.isArray(verifyRows) ? verifyRows : [],
        insertedCount: inserted.length,
        failedCount: failed.length,
        failed: failed.length ? failed : undefined,
      },
      { status: 201 }
    );
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
    console.error("[SAVE REPORT PHOTOS FAILED]", {
      reportId,
      payload: body,
      lastInsertedRow,
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
        error: "Failed to save report photos",
        detail: e?.sqlMessage || e?.message || String(error),
        code: e?.code || null,
        sqlState: e?.sqlState || null,
      },
      { status: 500 }
    );
  }
}
