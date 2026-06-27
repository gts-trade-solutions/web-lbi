/* eslint-disable @typescript-eslint/no-explicit-any */
import { v4 as uuidv4 } from "uuid";
import type { ResultSetHeader } from "mysql2";
import { ListObjectsV2Command, PutObjectCommand } from "@aws-sdk/client-s3";
import pool from "../../../../../lib/db";
import { requireAuth } from "../../../../../lib/auth";
import { s3Client, S3_BUCKET_NAME, getPublicS3Url } from "../../../../../lib/s3";

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
  // image_refs lets a single point map to multiple files when the master
  // sheet has more than one row per point_key with file references.
  image_refs?: Array<{ file_name?: string | null; image_key?: string | null }>;
};

type S3LookupOptions = {
  enabled: boolean;
  // List S3 under one or more prefixes when searching for filenames.
  prefixes: string[];
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

/**
 * Spec-mandated filename normaliser. Keep behaviourally identical to the
 * frontend's normalizeFileName so a file picked client-side matches the
 * server-side uploadedImageMap by the same key shape.
 */
function normalizeFileName(value: unknown): string {
  if (!value) return "";
  let text = String(value).trim();
  try {
    text = decodeURIComponent(text);
  } catch {
    /* ignore */
  }
  return (
    text
      .replace(/\\/g, "/")
      .split("/")
      .pop()
      ?.trim()
      .toLowerCase() || ""
  );
}

/**
 * Strip directory components, decode URL-encoding, lowercase, collapse spaces.
 * Used as the canonical key for matching master-file file_name strings to
 * actual S3 object basenames.
 */
function normalizeFileKey(value: unknown): string {
  let s = String(value ?? "")
    .trim()
    .replace(/\\/g, "/");
  try {
    s = decodeURIComponent(s);
  } catch {
    // ignore malformed percent-escapes
  }
  s = s.split("/").pop() || "";
  return s.toLowerCase().replace(/\s+/g, " ").trim();
}

/**
 * Walk one or more S3 prefixes and build a basename->Key map. Result is
 * normalised (lower-case basename) so we can match master-file references
 * regardless of folder, case, or encoding. Returns null on any S3 failure
 * (caller treats lookup as disabled).
 */
async function buildS3FileIndex(prefixes: string[]): Promise<Map<string, string> | null> {
  if (!S3_BUCKET_NAME) return null;
  const index = new Map<string, string>();
  try {
    for (const rawPrefix of prefixes) {
      const prefix = String(rawPrefix || "").replace(/^\/+/, "");
      let token: string | undefined;
      let pages = 0;
      // Cap pages to avoid pathological scans on huge buckets.
      while (pages < 50) {
        const out = await s3Client.send(
          new ListObjectsV2Command({
            Bucket: S3_BUCKET_NAME,
            Prefix: prefix || undefined,
            ContinuationToken: token,
            MaxKeys: 1000,
          })
        );
        for (const obj of out.Contents || []) {
          const key = obj.Key || "";
          if (!key) continue;
          const base = normalizeFileKey(key);
          if (!base) continue;
          // First match wins so deeper/duplicate copies don't override the
          // shallower one.
          if (!index.has(base)) index.set(base, key);
        }
        if (!out.IsTruncated || !out.NextContinuationToken) break;
        token = out.NextContinuationToken;
        pages += 1;
      }
    }
    return index;
  } catch (err) {
    console.error("[bulk import s3 lookup] ListObjectsV2 failed:", err);
    return null;
  }
}

/**
 * Resolve a single (file_name, image_key) reference to a canonical S3 key
 * using the prebuilt index. Falls back to direct image_key match.
 */
function resolveS3KeyForRef(
  ref: { file_name?: string | null; image_key?: string | null },
  index: Map<string, string>
): string | null {
  // Prefer matching image_key directly when it looks like a real key/path.
  const ikRaw = String(ref.image_key || "").trim();
  if (ikRaw) {
    const ikBase = normalizeFileKey(ikRaw);
    if (ikBase && index.has(ikBase)) return index.get(ikBase) || null;
    // image_key may already be the full S3 key ("reports/photos/xxx/y.jpg").
    const looksLikeKey = ikRaw.includes("/") && /\.[a-z0-9]+$/i.test(ikRaw);
    if (looksLikeKey) return ikRaw.replace(/^\/+/, "");
  }
  const fnBase = normalizeFileKey(ref.file_name);
  if (fnBase && index.has(fnBase)) return index.get(fnBase) || null;
  return null;
}

/**
 * Best-effort: ensure reports has file_name / image_key / point_key columns
 * so bulk import can persist them straight onto the report row even when no
 * report_photos row is created. Failure is logged but never aborts the
 * import - schemas without ALTER privileges silently skip.
 */
async function ensureReportImageColumns(existing: Set<string>) {
  const adds: string[] = [];
  if (!existing.has("file_name")) adds.push("file_name TEXT NULL");
  if (!existing.has("image_key")) adds.push("image_key TEXT NULL");
  for (const sql of adds) {
    try {
      await pool.query(`ALTER TABLE reports ADD COLUMN ${sql}`);
      console.log("[bulk import] added reports column:", sql);
    } catch (err) {
      console.error("[bulk import] ALTER reports failed:", sql, err);
    }
  }
}

async function getReportPhotoColumns(): Promise<Set<string>> {
  try {
    const [rows] = await pool.query("SHOW COLUMNS FROM report_photos");
    const set = new Set<string>();
    for (const r of Array.isArray(rows) ? rows : []) {
      const name = String((r as DbColumnRow).Field || "").toLowerCase();
      if (name) set.add(name);
    }
    return set;
  } catch {
    return new Set();
  }
}

/**
 * Best-effort: add point_key / image_key / file_name columns to report_photos
 * if they are missing. Failure is logged but never aborts the import.
 */
async function ensureReportPhotoColumns(existing: Set<string>) {
  const adds: string[] = [];
  if (!existing.has("file_name")) adds.push("file_name VARCHAR(255) NULL");
  if (!existing.has("image_key")) adds.push("image_key TEXT NULL");
  if (!existing.has("point_key")) adds.push("point_key VARCHAR(255) NULL");
  for (const sql of adds) {
    try {
      await pool.query(`ALTER TABLE report_photos ADD COLUMN ${sql}`);
      console.log("[bulk import] added report_photos column:", sql);
    } catch (err) {
      console.error("[bulk import] ALTER report_photos failed:", sql, err);
    }
  }
}

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

    // Accept BOTH application/json and multipart/form-data. When the
    // caller sends multipart, we read the row data from a JSON-encoded
    // "rows" / "payload" field AND collect every File from the well-known
    // image keys so we can match them locally and insert report_photos
    // ourselves — instead of relying on a follow-up /api/upload chain.
    const contentType = String(request.headers.get("content-type") || "").toLowerCase();
    const isMultipart = contentType.includes("multipart/form-data");

    let body: any = {};
    let localImages: File[] = [];
    if (isMultipart) {
      const fd = await request.formData();
      console.log("[BULK FORM KEYS]", Array.from(fd.keys()));
      const rowsField =
        String(fd.get("rows") || fd.get("payload") || "").trim();
      if (rowsField) {
        try {
          const parsed = JSON.parse(rowsField);
          body = parsed && typeof parsed === "object" ? parsed : { rows: parsed };
        } catch (e) {
          console.warn("[bulk import] failed to JSON.parse rows field:", e);
        }
      }
      // Some callers send fields one-by-one alongside images.
      if (!body.s3Lookup) {
        const lookupRaw = String(fd.get("s3Lookup") || "").trim();
        if (lookupRaw) {
          try {
            body.s3Lookup = JSON.parse(lookupRaw);
          } catch {
            /* ignore */
          }
        }
      }
      // Spec: support every plausible image FormData key. We grab Files
      // only — strings under these keys are silently ignored.
      localImages = [
        ...fd.getAll("images"),
        ...fd.getAll("files"),
        ...fd.getAll("photos"),
        ...fd.getAll("imageFiles"),
        ...fd.getAll("attachments"),
      ].filter((v): v is File => v instanceof File);
    } else {
      body = await request.json().catch(() => ({} as any));
    }

    const rowsIn: IncomingRow[] = Array.isArray(body?.rows) ? body.rows : [];

    console.log("[BULK API IMAGE RECEIVE]", {
      projectId,
      mode: isMultipart ? "multipart" : "json",
      imagesCount: localImages.length,
      sampleImages: localImages.slice(0, 10).map((f) => ({
        name: f.name,
        size: f.size,
        type: f.type,
      })),
    });

    // Spec-mandated first-image diagnostics. Surfaces the EXACT shape of
    // the first 3 parsed rows AND the first 5 uploaded files so the
    // operator can prove the first row's filename matches a file.
    console.log("[FIRST IMAGE DEBUG parsed first rows]", {
      parsedRowsCount: rowsIn.length,
      first3Rows: rowsIn.slice(0, 3).map((r, i) => ({
        index: i,
        point_key: r.point_key,
        file_name: r.file_name,
        // Defensive: surface every aliased filename column even though
        // the parser collapses them into r.file_name on the client.
        filename: (r as Record<string, unknown>).filename ?? null,
        image: (r as Record<string, unknown>).image ?? null,
        image_name: (r as Record<string, unknown>).image_name ?? null,
        photo: (r as Record<string, unknown>).photo ?? null,
        photo_name: (r as Record<string, unknown>).photo_name ?? null,
        image_key: r.image_key,
        photo_key: (r as Record<string, unknown>).photo_key ?? null,
      })),
    });
    console.log("[FIRST IMAGE DEBUG uploaded first files]", {
      imagesCount: localImages.length,
      first5Files: localImages.slice(0, 5).map((f, i) => ({
        index: i,
        name: f.name,
        size: f.size,
        type: f.type,
        normalizedName: normalizeFileName(f.name),
      })),
    });
    if (rowsIn[0]) {
      console.log("[FIRST IMAGE DEBUG first parsed row included]", {
        index: 0,
        point_key: rowsIn[0].point_key,
        file_name: rowsIn[0].file_name,
        image_key: rowsIn[0].image_key,
      });
    } else {
      console.error("[FIRST IMAGE DEBUG no first parsed row]");
    }

    // Optional server-side S3 lookup. When enabled, the route walks the given
    // prefixes once, builds a normalised filename->key index, and uses that to
    // create report_photos rows pointing at the canonical public URL. This is
    // the path used by the "S3 (configured server-side)" Storage option in
    // the Manual Bulk Upload modal - users do not have to re-select files
    // that already exist in the bucket.
    const s3Lookup: S3LookupOptions = (() => {
      const raw = body?.s3Lookup;
      if (!raw || typeof raw !== "object") {
        return { enabled: false, prefixes: [] };
      }
      const prefixesIn = Array.isArray(raw.prefixes)
        ? raw.prefixes
        : raw.prefix
          ? [raw.prefix]
          : [];
      const prefixes = prefixesIn
        .map((p: unknown) => String(p || "").trim())
        .filter(Boolean);
      return { enabled: !!raw.enabled, prefixes };
    })();
    if (s3Lookup.enabled && !s3Lookup.prefixes.length) {
      // Sensible defaults: the layouts the web frontend currently uses.
      s3Lookup.prefixes = [
        `reports/photos/`,
        `reports/${projectId}/`,
        `reports/photos/${projectId}/`,
      ];
    }

    console.log("[bulk import api] projectId:", projectId, "rows:", rowsIn.length, {
      s3Lookup: { enabled: s3Lookup.enabled, prefixCount: s3Lookup.prefixes.length },
    });

    if (!rowsIn.length) {
      return Response.json({ error: "rows[] required" }, { status: 400 });
    }

    // Spec-mandated trace: log the shape of the rows we just parsed so we can
    // confirm point_key / file_name / image_key / category survived JSON
    // transport intact.
    console.log("[BULK PHOTO TRACE parsed rows]", {
      totalRows: rowsIn.length,
      sampleRows: rowsIn.slice(0, 10).map((r) => ({
        point_key: r.point_key,
        file_name: r.file_name,
        image_key: r.image_key,
        category: r.category,
        image_refs_count: Array.isArray(r.image_refs) ? r.image_refs.length : 0,
      })),
    });

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

    // Pre-flight: ensure reports has file_name/image_key columns so we can
    // persist the master-file refs onto the report row. Refresh the column
    // set so the field-builder later in the loop sees the new columns.
    await ensureReportImageColumns(cols);
    {
      const refreshed = await getReportColumns();
      // Replace cols.has() lookups by re-syncing the local Set with the live
      // schema after ALTER. (Safer than mutating the original Set object.)
      for (const n of refreshed.names) cols.add(n);
    }

    // Pre-flight: prepare the S3 filename index once and ensure
    // report_photos has the optional file_name/image_key/point_key columns we
    // want to populate. Failures are logged but never abort the import.
    const photoCols = await getReportPhotoColumns();
    await ensureReportPhotoColumns(photoCols);
    // Refresh the column set in case ALTER added some.
    const photoColsLatest = await getReportPhotoColumns();
    const photoHas = (c: string) => photoColsLatest.has(c.toLowerCase());

    // Build the in-request image map from File objects sent in the
    // multipart payload. The map is keyed by the SAME normaliser the
    // frontend uses, so a master-file file_name "IMG_001.JPG" matches an
    // uploaded "img_001.jpg".
    const uploadedImageMap = new Map<string, File>();
    for (const file of localImages) {
      const k = normalizeFileName(file.name);
      if (k) uploadedImageMap.set(k, file);
    }
    if (localImages.length) {
      console.log("[BULK IMAGE MAP]", {
        imagesCount: localImages.length,
        keysSample: Array.from(uploadedImageMap.keys()).slice(0, 20),
      });
    }

    const totalRefs = rowsIn.reduce((sum, r) => {
      const refs = Array.isArray(r.image_refs) ? r.image_refs : [];
      const total =
        refs.length + (r.file_name || r.image_key ? 1 : 0);
      return sum + total;
    }, 0);

    let s3Index: Map<string, string> | null = null;
    if (s3Lookup.enabled && totalRefs > 0) {
      console.log("[bulk import api] building S3 index for prefixes:", s3Lookup.prefixes);
      s3Index = await buildS3FileIndex(s3Lookup.prefixes);
      console.log(
        "[bulk import api] S3 index size:",
        s3Index ? s3Index.size : "(disabled)"
      );
    }

    let insertedCount = 0;
    let updatedCount = 0;
    let photosMatched = 0;
    let photosMissing = 0;
    let rowsWithFileName = 0;
    let rowsWithImageKey = 0;
    let reportPhotosInserted = 0;
    let reportPhotosVerified = 0;
    const photosMissingSamples: string[] = [];
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
      // Persist the master-file image refs straight onto the report row so
      // the export helper can fall back to them when report_photos has no
      // join row yet. The columns are auto-added above when missing.
      const rowFileName = r.file_name ? String(r.file_name).trim() : null;
      const rowImageKey = r.image_key ? String(r.image_key).trim() : null;
      if (rowFileName) rowsWithFileName += 1;
      if (rowImageKey) rowsWithImageKey += 1;
      console.log("[IMPORT photo parsed]", {
        point_key,
        file_name: rowFileName,
        image_key: rowImageKey,
      });
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
      if (has("file_name")) fields.push({ col: "file_name", value: rowFileName });
      if (has("image_key")) fields.push({ col: "image_key", value: rowImageKey });
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

      // Spec-required confirmation that the report row is saved AND the id
      // we are about to use for report_photos is the same id the export will
      // query against later. The savedReportId logged here MUST be identical
      // to the report_id used in every subsequent report_photos INSERT for
      // this row.
      console.log("[BULK PHOTO TRACE report saved]", {
        point_key,
        savedReportId: String(reportId),
        projectId,
        category,
      });

      // ---- Per-row photo resolution. Combines top-level file_name/image_key
      // and image_refs[] so callers can pass either shape. Each ref is matched
      // against the prebuilt S3 index; on hit we INSERT into report_photos
      // with the canonical public URL keyed off the EXACT reports.id we just
      // saved. Existing rows for the same (report_id, normalised file_name)
      // are removed first so re-imports stay idempotent.
      // Spec: top-level r.file_name may be a comma-separated string of
      // multiple filenames for the same observation. Expand here so the
      // local-upload + s3-lookup branches each iterate over the full
      // list. r.image_key follows the same shape.
      const refs: Array<{ file_name?: string | null; image_key?: string | null }> = [];
      if (r.file_name || r.image_key) {
        const splitFn = String(r.file_name || "")
          .split(",")
          .map((s) => s.trim())
          .filter(Boolean);
        const splitIk = String(r.image_key || "")
          .split(",")
          .map((s) => s.trim())
          .filter(Boolean);
        if (splitFn.length <= 1 && splitIk.length <= 1) {
          refs.push({ file_name: r.file_name || null, image_key: r.image_key || null });
        } else {
          const max = Math.max(splitFn.length, splitIk.length, 1);
          for (let s = 0; s < max; s += 1) {
            refs.push({
              file_name: splitFn[s] || splitFn[0] || null,
              image_key: splitIk[s] || splitIk[0] || null,
            });
          }
        }
      }
      if (Array.isArray(r.image_refs)) {
        for (const x of r.image_refs) {
          if (!x) continue;
          if (x.file_name || x.image_key) {
            refs.push({ file_name: x.file_name || null, image_key: x.image_key || null });
          }
        }
      }

      // ---- LOCAL UPLOAD BRANCH. When the multipart request carried image
      // files, we resolve refs against the in-request map FIRST, upload
      // matched files to S3, and INSERT the report_photos row right here
      // using the EXACT savedReport.id. This branch must NOT depend on
      // s3Lookup / s3Index being enabled — it succeeds purely from the
      // files the caller sent in this same request. Refs that fail to
      // match a local file fall through to the existing s3Lookup branch
      // below (which itself falls back to the legacy "missing" warning).
      const refsResolvedLocally = new Set<number>();
      if (refs.length && uploadedImageMap.size > 0) {
        for (let refIdx = 0; refIdx < refs.length; refIdx += 1) {
          const ref = refs[refIdx];
          const rawFileName =
            (ref.file_name as string | null) ||
            (r.file_name as string | null) ||
            "";
          const normalizedFileName = normalizeFileName(rawFileName);
          const matchedFile = normalizedFileName
            ? uploadedImageMap.get(normalizedFileName) || null
            : null;

          console.log("[BULK LOCAL IMAGE MATCH]", {
            projectId,
            reportId: String(reportId),
            point_key,
            rawFileName,
            normalizedFileName,
            matched: !!matchedFile,
            matchedFileName: matchedFile?.name || null,
          });

          // Spec-mandated extra trace for the FIRST data row (i === 0).
          // Use `i >= 0` semantics — explicitly numeric so index 0 is
          // never treated as falsy.
          if (i === 0) {
            console.log("[FIRST IMAGE DEBUG row match]", {
              rowIndex: i,
              point_key,
              rawFileName,
              normalizedFileName,
              matched: !!matchedFile,
              matchedFileName: matchedFile?.name || null,
            });
          }

          if (!matchedFile) continue;

          if (!S3_BUCKET_NAME) {
            console.error("[BULK LOCAL IMAGE UPLOAD] S3 bucket not configured — skipping");
            continue;
          }

          const safeBaseName =
            String(matchedFile.name || "upload")
              .replace(/[^a-zA-Z0-9._-]/g, "-")
              .replace(/-+/g, "-")
              .slice(0, 180);
          const key = `reports/photos/${reportId}/${uuidv4()}-${safeBaseName}`;

          let publicUrl = "";
          try {
            const bytes = Buffer.from(await matchedFile.arrayBuffer());
            await s3Client.send(
              new PutObjectCommand({
                Bucket: S3_BUCKET_NAME,
                Key: key,
                Body: bytes,
                ContentType: matchedFile.type || "application/octet-stream",
              })
            );
            publicUrl = getPublicS3Url(key);
            console.log("[BULK LOCAL IMAGE UPLOAD]", {
              projectId,
              reportId: String(reportId),
              point_key,
              key,
              publicUrl,
              size: bytes.length,
            });
          } catch (uErr) {
            console.error("[BULK LOCAL IMAGE UPLOAD] S3 PutObject failed:", uErr);
            continue;
          }

          // Spec: dedup by (report_id, file_name) so re-imports replace
          // cleanly without wiping unrelated photos for the same report.
          try {
            await pool.query(
              "DELETE FROM report_photos WHERE report_id = ? AND file_name = ?",
              [reportId, rawFileName || matchedFile.name]
            );
          } catch (delErr) {
            console.warn("[BULK LOCAL PHOTO INSERT] dedup DELETE failed - continuing:", delErr);
          }

          // Schema-safe INSERT against the actual report_photos columns.
          // Mandatory: id, report_id, url. Optional: width/height/user_id/
          // file_name/point_key/image_key — emitted only when the live
          // schema declares them. NEVER emit `path` (column does not
          // exist on this install).
          const photoFields: Array<{ col: string; value: unknown }> = [
            { col: "id", value: uuidv4() },
            { col: "report_id", value: reportId },
            { col: "url", value: publicUrl },
          ];
          if (photoHas("width")) photoFields.push({ col: "width", value: null });
          if (photoHas("height")) photoFields.push({ col: "height", value: null });
          if (photoHas("user_id")) photoFields.push({ col: "user_id", value: userId || null });
          if (photoHas("file_name"))
            photoFields.push({ col: "file_name", value: rawFileName || matchedFile.name });
          if (photoHas("point_key")) photoFields.push({ col: "point_key", value: point_key || null });
          if (photoHas("image_key"))
            photoFields.push({ col: "image_key", value: ref.image_key || r.image_key || null });

          const colsList = photoFields.map((f) => f.col);
          const placeholders = photoFields.map(() => "?");

          try {
            await pool.query(
              `INSERT INTO report_photos (${colsList.join(", ")}) VALUES (${placeholders.join(", ")})`,
              photoFields.map((f) => f.value)
            );
            photosMatched += 1;
            reportPhotosInserted += 1;
            refsResolvedLocally.add(refIdx);

            // Verify the link landed.
            try {
              const [verifyRows] = await pool.query(
                `SELECT id, report_id, url, file_name, image_key, point_key
                 FROM report_photos WHERE report_id = ?`,
                [reportId]
              );
              const verified = Array.isArray(verifyRows) ? verifyRows : [];
              if (verified.length) reportPhotosVerified += 1;
              console.log("[BULK LOCAL PHOTO INSERT VERIFY]", {
                reportId: String(reportId),
                point_key,
                count: verified.length,
                rows: verified,
              });

              // Spec-mandated extra verification for the FIRST data row.
              // Re-runs the spec's exact JOIN query so the operator sees
              // the parent project_id alongside the photo row — proving
              // (or disproving) that the first row's photo landed.
              if (i === 0) {
                try {
                  const [firstVerify] = await pool.query(
                    `SELECT rp.id, rp.report_id, rp.url, rp.file_name,
                            rp.point_key, rp.image_key,
                            r.project_id, r.point_key AS report_point_key
                     FROM report_photos rp
                     JOIN reports r ON r.id = rp.report_id
                     WHERE rp.report_id = ?`,
                    [reportId]
                  );
                  console.log("[FIRST IMAGE DEBUG first report photo verify]", {
                    projectId,
                    reportId: String(reportId),
                    point_key,
                    rows: firstVerify,
                  });
                } catch (fvErr) {
                  console.warn(
                    "[FIRST IMAGE DEBUG first report photo verify] failed:",
                    fvErr
                  );
                }
              }
            } catch (vErr) {
              console.warn("[BULK LOCAL PHOTO INSERT VERIFY] SELECT failed:", vErr);
            }
          } catch (iErr) {
            console.error("[BULK LOCAL PHOTO INSERT] failed:", iErr);
          }
        }
      }

      // Strip refs that the local branch already handled so the s3Lookup
      // fallback below does not double-insert for the same file.
      const remainingRefs = refs.filter((_, idx) => !refsResolvedLocally.has(idx));
      const _refsBackup = refs.length;
      // Reuse the existing branch logic by swapping the refs array in-place.
      refs.length = 0;
      for (const x of remainingRefs) refs.push(x);
      if (refsResolvedLocally.size > 0) {
        console.log("[BULK LOCAL PHOTO BRANCH SUMMARY]", {
          reportId: String(reportId),
          point_key,
          totalRefs: _refsBackup,
          resolvedLocally: refsResolvedLocally.size,
          remainingForS3Lookup: remainingRefs.length,
        });
      }

      if (refs.length && s3Lookup.enabled && s3Index) {
        // Spec-mandated idempotency: wipe ALL prior photos for THIS report
        // before inserting anything new. Runs ONCE per row, BEFORE the refs
        // loop, so multi-photo reports do not lose earlier inserts. Note:
        // this also drops manually-uploaded photos on re-import, which is
        // the contract the spec asks for ("DELETE FROM report_photos WHERE
        // report_id = ?").
        try {
          await pool.query("DELETE FROM report_photos WHERE report_id = ?", [reportId]);
        } catch (delErr) {
          console.warn("[bulk import photo] per-report DELETE failed - continuing:", delErr);
        }

        for (const ref of refs) {
          const key = resolveS3KeyForRef(ref, s3Index);
          const normalizedFileName = normalizeFileKey(ref.file_name || "");

          // Spec-mandated trace #1: did this row's file_name / image_key
          // match anything in the S3 index? When matched=false the photo
          // never reaches the INSERT step.
          console.log("[BULK PHOTO TRACE row match]", {
            point_key,
            file_name: ref.file_name,
            normalizedFileName,
            image_key: ref.image_key,
            matchedKey: key,
            matched: !!key,
            s3IndexSize: s3Index.size,
          });

          if (!key) {
            console.warn("[BULK PHOTO TRACE photo missing]", {
              savedReportId: String(reportId),
              point_key,
              file_name: ref.file_name,
              image_key: ref.image_key,
              reason: "no match in S3 lookup index",
              s3IndexSize: s3Index.size,
            });
            photosMissing += 1;
            if (photosMissingSamples.length < 10) {
              photosMissingSamples.push(
                String(ref.file_name || ref.image_key || "(unknown)")
              );
            }
            continue;
          }
          let publicUrl: string;
          try {
            publicUrl = getPublicS3Url(key);
          } catch (err) {
            console.error("[bulk import photo] getPublicS3Url failed:", key, err);
            console.warn("[BULK PHOTO TRACE photo missing]", {
              savedReportId: String(reportId),
              point_key,
              file_name: ref.file_name,
              image_key: ref.image_key,
              reason: "getPublicS3Url threw",
            });
            photosMissing += 1;
            continue;
          }

          // Spec-mandated trace #2: which savedReportId is this S3 URL about
          // to be linked to? If this id is wrong, the report_photos row will
          // be orphaned from the master-file report.
          console.log("[BULK PHOTO TRACE s3 uploaded]", {
            point_key,
            savedReportId: String(reportId),
            matchedFileName: normalizedFileName || normalizeFileKey(key),
            finalImageUrl: publicUrl,
          });

          const baseName =
            normalizedFileName || normalizeFileKey(key);

          // Build INSERT only with columns that actually exist in this DB.
          // The id column carries a generated UUID for schemas where the
          // PRIMARY KEY has no DEFAULT (UUID()) — newer schemas would auto-
          // populate it but we keep it explicit for safety.
          const photoId = uuidv4();
          const photoFields: Array<{ col: string; value: unknown }> = [
            { col: "id", value: photoId },
            { col: "report_id", value: reportId },
            { col: "url", value: publicUrl },
          ];
          if (photoHas("path")) photoFields.push({ col: "path", value: key });
          if (photoHas("file_name"))
            photoFields.push({ col: "file_name", value: ref.file_name || baseName || null });
          if (photoHas("image_key"))
            photoFields.push({ col: "image_key", value: ref.image_key || null });
          if (photoHas("point_key"))
            photoFields.push({ col: "point_key", value: point_key });

          const colsList = photoFields.map((f) => f.col);
          const placeholders = photoFields.map(() => "?").join(", ");

          // Spec-mandated trace #3: log the EXACT id values being written
          // into report_photos. report_id_used MUST equal savedReport.id.
          console.log("[BULK PHOTO TRACE insert report_photos]", {
            report_id_used: String(reportId),
            project_id: projectId,
            point_key,
            file_name: ref.file_name,
            image_key: ref.image_key,
            url: publicUrl,
          });

          try {
            await pool.query(
              `INSERT INTO report_photos (${colsList.join(", ")}) VALUES (${placeholders})`,
              photoFields.map((f) => f.value)
            );
            photosMatched += 1;
            reportPhotosInserted += 1;

            // Spec-mandated trace #4: read back the row keyed by the EXACT
            // report_id we just inserted with. count > 0 proves the FK link
            // landed; count = 0 proves the bug we are tracking.
            try {
              const [verifyRows] = await pool.query(
                `SELECT id, report_id, url, file_name, image_key, point_key
                 FROM report_photos WHERE report_id = ?`,
                [reportId]
              );
              const verified = Array.isArray(verifyRows) ? verifyRows : [];
              if (verified.length) reportPhotosVerified += 1;
              console.log("[BULK PHOTO TRACE verify report_photos]", {
                savedReportId: String(reportId),
                point_key,
                count: verified.length,
                rows: verified,
              });
            } catch (vErr) {
              console.warn("[BULK PHOTO TRACE verify report_photos] SELECT failed:", vErr);
            }
          } catch (err) {
            console.error("[bulk import photo] INSERT failed:", err);
            console.warn("[BULK PHOTO TRACE photo missing]", {
              savedReportId: String(reportId),
              point_key,
              file_name: ref.file_name,
              image_key: ref.image_key,
              reason: "INSERT into report_photos failed",
            });
            photosMissing += 1;
          }
        }
      } else if (refs.length && (!s3Lookup.enabled || !s3Index)) {
        // S3 lookup disabled or index unavailable; the frontend's manual
        // upload path will create the report_photos rows. Count them as
        // missing here so the summary still reflects unresolved refs.
        for (const ref of refs) {
          console.warn("[BULK PHOTO TRACE photo missing]", {
            savedReportId: String(reportId),
            point_key,
            file_name: ref.file_name,
            image_key: ref.image_key,
            reason: !s3Lookup.enabled
              ? "s3Lookup disabled in request"
              : "s3 index could not be built (check bucket permissions)",
          });
        }
        photosMissing += refs.length;
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

    console.log("[bulk import api] photos matched:", photosMatched, "missing:", photosMissing);

    // Spec-mandated final verification: run the SAME query the user runs
    // manually after import, so the API response itself reports whether the
    // report ↔ report_photos join landed for this project. If
    // reports_with_photos is 0 here, the photo INSERT path silently failed
    // and the per-row [BULK PHOTO TRACE ...] logs above will pinpoint where.
    let projectVerify: {
      total_reports: number;
      reports_with_photos: number;
      reports_with_photo_url: number;
    } = { total_reports: 0, reports_with_photos: 0, reports_with_photo_url: 0 };
    try {
      const [vRows] = await pool.query(
        `SELECT
           COUNT(*) AS total_reports,
           SUM(CASE WHEN rp.id IS NOT NULL THEN 1 ELSE 0 END) AS reports_with_photos,
           SUM(CASE WHEN rp.url IS NOT NULL AND rp.url <> '' THEN 1 ELSE 0 END) AS reports_with_photo_url
         FROM reports r
         LEFT JOIN report_photos rp ON rp.report_id = r.id
         WHERE r.project_id = ?`,
        [projectId]
      );
      const vRow = Array.isArray(vRows) ? (vRows[0] as any) : null;
      if (vRow) {
        projectVerify = {
          total_reports: Number(vRow.total_reports) || 0,
          reports_with_photos: Number(vRow.reports_with_photos) || 0,
          reports_with_photo_url: Number(vRow.reports_with_photo_url) || 0,
        };
      }
      // Spec Part 4 log key: emits the raw DB COUNT row so the operator
      // can grep for this exact line and confirm the per-project link
      // landed for the bulk import that just ran.
      console.log("[bulk] final photo DB check", vRows);
    } catch (vErr) {
      console.warn("[BULK PHOTO TRACE project verify] failed:", vErr);
    }

    console.log("[BULK PHOTO TRACE summary]", {
      projectId,
      totalRows: rowsIn.length,
      rowsWithFileName,
      rowsWithImageKey,
      reportsSaved: insertedCount + updatedCount,
      insertedCount,
      updatedCount,
      // Local-upload tallies (in-request multipart files):
      localImagesReceived: localImages.length,
      localImagesMatched: photosMatched,
      localImagesUploaded: photosMatched,
      // Combined photo tallies:
      matchedImagesCount: photosMatched,
      uploadedToS3Count: photosMatched,
      reportPhotosInserted,
      reportPhotosVerified,
      photosMissing,
      photosMissingCount: photosMissing,
      s3LookupEnabled: s3Lookup.enabled,
      s3IndexSize: s3Index ? s3Index.size : 0,
      projectVerify,
    });

    return Response.json({
      ok: true,
      projectId,
      insertedCount,
      updatedCount,
      reportsSaved: insertedCount + updatedCount,
      photosMatched,
      photosMissing,
      photosMissingSamples,
      rowsWithFileName,
      rowsWithImageKey,
      reportPhotosInserted,
      reportPhotosVerified,
      s3LookupEnabled: s3Lookup.enabled,
      s3IndexSize: s3Index ? s3Index.size : 0,
      projectVerify,
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
