/* eslint-disable @typescript-eslint/no-explicit-any */
import { v4 as uuidv4 } from "uuid";
import type { ResultSetHeader } from "mysql2";
import pool from "../../../../../lib/db";
import { requireAuth } from "../../../../../lib/auth";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type Ctx = { params: { id: string } };

type IncomingRow = {
  point_key: string;
  latitude: number | string | null;
  longitude: number | string | null;
  category?: string | null;
  description?: string | null;
  difficulty?: string | null;
  remarks_action?: string | null;
  file_name?: string | null;
  image_key?: string | null;
};

type DbColumnRow = {
  Field?: string;
  Null?: string;
  Default?: string | null;
  Extra?: string;
};

type ReportColumns = {
  names: Set<string>;
  // Columns that are NOT NULL, have no default, and aren't auto-incremented —
  // we must supply a value for each of these on INSERT or MySQL throws
  // "Field 'X' doesn't have a default value".
  requiredOnInsert: Set<string>;
  // True when reports.id is auto_increment — in that case the INSERT must
  // OMIT id and we read result.insertId for the new row's id.
  idIsAutoIncrement: boolean;
};

function unauthorized(error: unknown) {
  return (error as { message?: string })?.message === "Unauthorized";
}

function toNumOrNull(v: unknown): number | null {
  if (v === null || typeof v === "undefined" || v === "") return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

async function getReportColumns(): Promise<ReportColumns> {
  const [rows] = await pool.query("SHOW COLUMNS FROM reports");
  const names = new Set<string>();
  const requiredOnInsert = new Set<string>();
  let idIsAutoIncrement = false;
  for (const r of Array.isArray(rows) ? rows : []) {
    const row = r as DbColumnRow;
    const name = String(row.Field || "").toLowerCase();
    if (!name) continue;
    names.add(name);
    const isNullable = String(row.Null || "").toUpperCase() === "YES";
    const hasDefault = row.Default !== null && typeof row.Default !== "undefined";
    const extra = String(row.Extra || "").toLowerCase();
    const isAuto =
      extra.includes("auto_increment") || extra.includes("default_generated");
    if (name === "id" && extra.includes("auto_increment")) {
      idIsAutoIncrement = true;
    }
    if (!isNullable && !hasDefault && !isAuto) {
      requiredOnInsert.add(name);
    }
  }
  return { names, requiredOnInsert, idIsAutoIncrement };
}

type PathPointColumns = {
  names: Set<string>;
  // Same NOT NULL / no-default / not-auto-incremented set as reports — needed
  // because report_path_points has its own user_id column on this install.
  requiredOnInsert: Set<string>;
  idIsAutoIncrement: boolean;
};

async function getPathPointColumns(): Promise<PathPointColumns> {
  try {
    const [rows] = await pool.query("SHOW COLUMNS FROM report_path_points");
    const names = new Set<string>();
    const requiredOnInsert = new Set<string>();
    let idIsAutoIncrement = false;
    for (const r of Array.isArray(rows) ? rows : []) {
      const row = r as DbColumnRow;
      const name = String(row.Field || "").toLowerCase();
      if (!name) continue;
      names.add(name);
      const isNullable = String(row.Null || "").toUpperCase() === "YES";
      const hasDefault = row.Default !== null && typeof row.Default !== "undefined";
      const extra = String(row.Extra || "").toLowerCase();
      const isAuto =
        extra.includes("auto_increment") || extra.includes("default_generated");
      if (name === "id" && extra.includes("auto_increment")) {
        idIsAutoIncrement = true;
      }
      if (!isNullable && !hasDefault && !isAuto) {
        requiredOnInsert.add(name);
      }
    }
    return { names, requiredOnInsert, idIsAutoIncrement };
  } catch {
    return { names: new Set(), requiredOnInsert: new Set(), idIsAutoIncrement: false };
  }
}

export async function POST(request: Request, context: Ctx) {
  try {
    const authUser = requireAuth(request);
    const userId = authUser?.id ? String(authUser.id) : "";
    console.log("[bulk import auth]", { userId });
    if (!userId) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }

    const projectId = String(context.params?.id || "").trim();
    if (!projectId) {
      return Response.json({ error: "Project id is required" }, { status: 400 });
    }

    const body = await request.json().catch(() => ({} as any));
    const rowsIn: IncomingRow[] = Array.isArray(body?.rows) ? body.rows : [];

    console.log("[bulk import api] projectId:", projectId, "rows:", rowsIn.length);

    if (!rowsIn.length) {
      return Response.json({ error: "rows[] required" }, { status: 400 });
    }

    // Validate the project exists.
    const [projectRows] = await pool.query(
      "SELECT id FROM projects WHERE id = ? LIMIT 1",
      [projectId]
    );
    if (!Array.isArray(projectRows) || projectRows.length === 0) {
      return Response.json({ error: "Project not found", projectId }, { status: 404 });
    }

    // Introspect the reports table so the route works on schemas where the
    // legacy `latitude`/`longitude` columns are absent (only `loc_lat`/
    // `loc_lon` exist) and vice versa, AND so we can satisfy any NOT NULL
    // columns that have no default (e.g., user_id) without hardcoding them.
    const colsInfo = await getReportColumns();
    const cols = colsInfo.names;
    const has = (c: string) => cols.has(c.toLowerCase());
    console.log("[bulk import api] reports columns:", Array.from(cols).join(","));
    console.log(
      "[bulk import api] required-on-insert columns:",
      Array.from(colsInfo.requiredOnInsert).join(",")
    );
    console.log(
      "[bulk import api] reports.id auto_increment:",
      colsInfo.idIsAutoIncrement
    );

    const pathPointCols = await getPathPointColumns();

    let insertedCount = 0;
    let updatedCount = 0;
    const reportsOut: Array<{
      point_key: string;
      report_id: string;
      file_name: string | null;
      image_key: string | null;
      action: "inserted" | "updated";
    }> = [];

    for (let i = 0; i < rowsIn.length; i += 1) {
      const r = rowsIn[i];
      const point_key = String(r.point_key || "").trim();
      if (!point_key) continue;

      const lat = toNumOrNull(r.latitude);
      const lng = toNumOrNull(r.longitude);
      const category = r.category ? String(r.category).trim() : null;
      const description = r.description ? String(r.description).trim() : null;
      // The action/actions/difficulty source column is saved into BOTH
      // reports.difficulty AND reports.remarks_action per the import spec, so
      // downstream consumers (DOCX export, project page) can read either field.
      const action =
        (r.remarks_action ? String(r.remarks_action).trim() : null) ||
        (r.difficulty ? String(r.difficulty).trim() : null);

      console.log("[bulk import save row]", {
        projectId,
        userId,
        point_key,
        loc_lat: lat,
        loc_lon: lng,
        category,
        description,
        difficulty: action,
        remarks_action: action,
      });

      const [existing] = await pool.query(
        "SELECT id FROM reports WHERE project_id = ? AND point_key = ? LIMIT 1",
        [projectId, point_key]
      );
      const existingRow = Array.isArray(existing) ? (existing[0] as any) : null;
      const existingReportId = existingRow?.id ?? null;
      console.log("[bulk import existing check]", {
        projectId,
        point_key,
        existingReportId,
      });

      // Build the per-column write list dynamically. Only columns present in
      // the live schema are touched, so the same code runs on schemas that
      // have just loc_lat/loc_lon, just latitude/longitude, both, or neither.
      const fields: Array<{ col: string; value: unknown }> = [];
      if (has("category")) fields.push({ col: "category", value: category });
      if (has("description")) fields.push({ col: "description", value: description });
      if (has("difficulty")) fields.push({ col: "difficulty", value: action });
      if (has("remarks_action")) fields.push({ col: "remarks_action", value: action });
      if (has("loc_lat")) fields.push({ col: "loc_lat", value: lat });
      if (has("loc_lon")) fields.push({ col: "loc_lon", value: lng });
      if (has("latitude")) fields.push({ col: "latitude", value: lat });
      if (has("longitude")) fields.push({ col: "longitude", value: lng });
      if (has("sort_order")) fields.push({ col: "sort_order", value: i + 1 });
      if (has("status")) fields.push({ col: "status", value: "active" });
      // Insert-only columns (don't reassign on UPDATE — they're set once at
      // creation). Built into a separate list so the UPDATE branch ignores
      // them and won't reset ownership / created_by on re-import.
      const insertOnlyFields: Array<{ col: string; value: unknown }> = [];
      if (has("user_id")) insertOnlyFields.push({ col: "user_id", value: userId });
      if (has("created_by")) insertOnlyFields.push({ col: "created_by", value: userId });
      // Any remaining NOT NULL no-default column the live schema declares.
      // We satisfy it with a safe placeholder so MySQL doesn't reject the
      // INSERT. The list of columns we can naturally fill is hard-coded above;
      // anything else will surface in the [bulk import api] required-on-insert
      // log so we can extend the list.
      const knownFilled = new Set([
        "id",
        "project_id",
        "point_key",
        ...fields.map((f) => f.col),
        ...insertOnlyFields.map((f) => f.col),
      ]);
      for (const reqCol of colsInfo.requiredOnInsert) {
        if (knownFilled.has(reqCol)) continue;
        // Best-effort: empty string for varchar/text, 0 for numeric. We won't
        // know the column type from SHOW COLUMNS without parsing Type, so use
        // empty string — MySQL will coerce or surface a clearer error.
        console.warn(
          `[bulk import api] unhandled NOT NULL column "${reqCol}" — defaulting to empty string`
        );
        insertOnlyFields.push({ col: reqCol, value: "" });
      }

      // reportId is `string | number` because the live schema decides:
      // auto_increment INT → number from result.insertId, otherwise UUID.
      let reportId: string | number;
      if (existingReportId !== null && typeof existingReportId !== "undefined") {
        reportId =
          typeof existingReportId === "number" ? existingReportId : String(existingReportId);
        // UPDATE never touches user_id / created_by — those stay with the
        // original creator across re-imports.
        const setClause = fields.map((f) => `${f.col} = ?`).join(", ");
        const args = [...fields.map((f) => f.value), reportId, projectId];
        await pool.query(
          `UPDATE reports SET ${setClause} WHERE id = ? AND project_id = ?`,
          args
        );
        updatedCount += 1;
        reportsOut.push({
          point_key,
          report_id: String(reportId),
          file_name: r.file_name || null,
          image_key: r.image_key || null,
          action: "updated",
        });
      } else {
        const allFields = [...fields, ...insertOnlyFields];
        // When reports.id is auto_increment, OMIT it from the INSERT — passing
        // a UUID into an INT column triggers "Incorrect integer value: '...'".
        const baseCols = colsInfo.idIsAutoIncrement
          ? ["project_id", "point_key"]
          : ["id", "project_id", "point_key"];
        const baseArgs: Array<string | number> = colsInfo.idIsAutoIncrement
          ? [projectId, point_key]
          : [uuidv4(), projectId, point_key];
        const insertCols = [...baseCols, ...allFields.map((f) => f.col)];
        const insertArgs = [...baseArgs, ...allFields.map((f) => f.value)];
        const placeholders = insertCols.map(() => "?").join(", ");
        console.log("[bulk import insert]", {
          userId,
          projectId,
          point_key,
          loc_lat: lat,
          loc_lon: lng,
          category,
          insertCols,
          autoIncrementId: colsInfo.idIsAutoIncrement,
        });
        const [result] = await pool.query(
          `INSERT INTO reports (${insertCols.join(", ")}) VALUES (${placeholders})`,
          insertArgs
        );
        const header = result as ResultSetHeader;
        if (colsInfo.idIsAutoIncrement) {
          reportId = header.insertId;
          if (!reportId) {
            throw new Error("Bulk import: INSERT did not return an insertId");
          }
        } else {
          // String/UUID id path — we passed the UUID in baseArgs.
          reportId = String(baseArgs[0]);
        }
        console.log("[bulk import inserted report]", {
          reportId,
          projectId,
          point_key,
        });
        insertedCount += 1;
        reportsOut.push({
          point_key,
          report_id: String(reportId),
          file_name: r.file_name || null,
          image_key: r.image_key || null,
          action: "inserted",
        });
      }

      // Always (re)write a single seq=1 path point so the per-report KM /
      // location pipeline sees the coordinates the master file declared.
      // Columns are picked dynamically: id only when not auto_increment;
      // user_id when the table has it (NOT NULL on this install); the
      // coordinate column pair that actually exists; and any other NOT NULL
      // no-default column the live schema declares.
      if (lat !== null && lng !== null) {
        await pool.query("DELETE FROM report_path_points WHERE report_id = ? AND seq = 1", [
          reportId,
        ]);

        const ppHas = (c: string) => pathPointCols.names.has(c.toLowerCase());
        const ppFields: Array<{ col: string; value: unknown }> = [];
        ppFields.push({ col: "report_id", value: reportId });
        if (ppHas("seq")) ppFields.push({ col: "seq", value: 1 });
        if (ppHas("latitude")) ppFields.push({ col: "latitude", value: lat });
        if (ppHas("longitude")) ppFields.push({ col: "longitude", value: lng });
        if (ppHas("loc_lat")) ppFields.push({ col: "loc_lat", value: lat });
        if (ppHas("loc_lon")) ppFields.push({ col: "loc_lon", value: lng });
        if (ppHas("user_id")) ppFields.push({ col: "user_id", value: userId });
        if (ppHas("created_by")) ppFields.push({ col: "created_by", value: userId });

        // id column: skip when auto_increment, supply UUID otherwise.
        if (ppHas("id") && !pathPointCols.idIsAutoIncrement) {
          ppFields.unshift({ col: "id", value: uuidv4() });
        }

        // Fill any remaining NOT NULL no-default column we didn't already
        // satisfy. timestamp is supplied via NOW() in SQL below if needed.
        const ppFilled = new Set(ppFields.map((f) => f.col));
        for (const reqCol of pathPointCols.requiredOnInsert) {
          if (ppFilled.has(reqCol)) continue;
          if (reqCol === "timestamp" || reqCol === "created_at" || reqCol === "updated_at") {
            // Will be filled with NOW() right below.
            continue;
          }
          console.warn(
            `[bulk import api] unhandled NOT NULL column on report_path_points "${reqCol}" — defaulting to empty string`
          );
          ppFields.push({ col: reqCol, value: "" });
        }

        // Compose INSERT. Any timestamp / created_at column gets NOW() inline
        // because mysql2 placeholders can't carry a SQL function.
        const cols = ppFields.map((f) => f.col);
        const placeholders = ppFields.map(() => "?");
        if (ppHas("timestamp")) {
          cols.push("timestamp");
          placeholders.push("NOW()");
        }
        if (ppHas("created_at") && !pathPointCols.requiredOnInsert.has("created_at")) {
          // Already has DEFAULT CURRENT_TIMESTAMP — skip.
        } else if (ppHas("created_at")) {
          cols.push("created_at");
          placeholders.push("NOW()");
        }

        console.log("[bulk import path point insert]", {
          userId,
          reportId,
          seq: 1,
          latitude: lat,
          longitude: lng,
          cols,
        });

        await pool.query(
          `INSERT INTO report_path_points (${cols.join(", ")}) VALUES (${placeholders.join(", ")})`,
          ppFields.map((f) => f.value)
        );
      }
    }

    console.log(
      "[bulk import api] inserted:",
      insertedCount,
      "updated:",
      updatedCount
    );

    return Response.json({
      ok: true,
      projectId,
      insertedCount,
      updatedCount,
      reports: reportsOut,
    });
  } catch (error) {
    if (unauthorized(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/projects/:id/bulk-import] POST error:", error);
    return Response.json(
      {
        error: "Failed to bulk import",
        message: (error as { message?: string })?.message || "unknown",
      },
      { status: 500 }
    );
  }
}
