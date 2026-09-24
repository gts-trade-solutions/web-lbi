import { v4 as uuidv4 } from "uuid";
import pool from "../../../../../lib/db";
import { requireAuth } from "../../../../../lib/auth";
import { canonicalS3Url, signRowUrls, getPublicS3Url, s3KeyFromUrl } from "../../../../../lib/s3";

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

// Per-photo Word-export selection flag. 1 (default) = photo is included in
// DOCX exports, 0 = user unticked it in the photo picker. Added lazily so
// existing databases upgrade themselves on first use.
async function ensureIncludeInExportColumn() {
  try {
    const cols = await getColumns();
    if (!cols.has("include_in_export")) {
      await pool.query(
        "ALTER TABLE report_photos ADD COLUMN include_in_export TINYINT(1) NOT NULL DEFAULT 1"
      );
      console.log("[api/reports/:id/photos] added include_in_export column");
    }
  } catch (err) {
    console.error("[api/reports/:id/photos] ensure include_in_export failed:", err);
  }
}

// Re-editable drawings. When the user draws on a photo we flatten the strokes
// into the saved image (for display/export) but ALSO keep the vector strokes
// (`anno_json`) and the untouched base image they were drawn over
// (`anno_base_url`) so "Draw on photo" can reopen and edit them later instead
// of drawing on top of a baked-in picture. Added lazily so existing databases
// upgrade themselves on first use.
async function ensureAnnoColumns() {
  try {
    const cols = await getColumns();
    if (!cols.has("anno_json")) {
      await pool.query("ALTER TABLE report_photos ADD COLUMN anno_json LONGTEXT NULL");
      console.log("[api/reports/:id/photos] added anno_json column");
    }
    if (!cols.has("anno_base_url")) {
      await pool.query("ALTER TABLE report_photos ADD COLUMN anno_base_url VARCHAR(2048) NULL");
      console.log("[api/reports/:id/photos] added anno_base_url column");
    }
    // Write-once copy of the UNTOUCHED original image. Captured the first time
    // a photo is drawn on / cropped, and never overwritten after, so the clean
    // original is always kept and can be viewed or restored.
    if (!cols.has("original_url")) {
      await pool.query("ALTER TABLE report_photos ADD COLUMN original_url VARCHAR(2048) NULL");
      console.log("[api/reports/:id/photos] added original_url column");
    }
  } catch (err) {
    console.error("[api/reports/:id/photos] ensure anno columns failed:", err);
  }
}

export async function GET(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const reportId = String(context.params?.id || "").trim();
    if (!reportId) return Response.json({ error: "Report id is required" }, { status: 400 });

    // ?resolveOriginal=<photoId> — return the untouched original image URL for
    // that photo (or null), so the UI can show the original before a drawing.
    const resolveId = new URL(request.url).searchParams.get("resolveOriginal");
    if (resolveId) {
      const orig = await findOriginalUrl(reportId, String(resolveId).trim());
      let originalUrl: string | null = null;
      if (orig) {
        const signed = await signRowUrls("report_photos", [{ url: orig } as Record<string, unknown>]);
        originalUrl =
          Array.isArray(signed) && signed[0] && typeof signed[0].url === "string"
            ? (signed[0].url as string)
            : orig;
      }
      return Response.json({ originalUrl });
    }

    await ensureIncludeInExportColumn();
    await ensureAnnoColumns();
    const [rows] = await pool.query(
      "SELECT * FROM report_photos WHERE report_id = ? ORDER BY created_at ASC",
      [reportId]
    );
    return Response.json({ photos: await signRowUrls("report_photos", Array.isArray(rows) ? rows : []) });
  } catch (error) {
    if (unauthorized(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/reports/:id/photos] GET error:", error);
    return Response.json({ error: "Failed to fetch report photos" }, { status: 500 });
  }
}

/** Only links a browser can actually show — never javascript: and friends. */
function isImageUrl(url: string): boolean {
  return url.length <= 2048 && (/^https?:\/\//i.test(url) || /^\/(?!\/)/.test(url));
}

/**
 * Swap the picture behind an existing photo, keeping everything else.
 *
 * The draw and crop tools used to save their result as a *new* photo (and
 * crop then deleted the original). Photos are ordered by created_at, so the
 * edited one jumped to the end — Photo 1 became Photo 3 in the Word export —
 * and the new row lost the photo's export tick and image_key. Updating the
 * row in place keeps its position, its number and its settings, and only the
 * picture changes.
 */
async function replacePhotoImage(reportId: string, body: Record<string, unknown>) {
  const photoId = String(body?.photoId || body?.id || "").trim();
  const url = canonicalS3Url(String(body?.url || "").trim());
  if (!photoId) return Response.json({ error: "photoId is required" }, { status: 400 });
  if (!isImageUrl(url)) return Response.json({ error: "A valid image url is required" }, { status: 400 });

  await ensureAnnoColumns();
  const cols = await getColumns();
  const sets: string[] = ["url = ?"];
  const values: unknown[] = [url];
  if (cols.has("file_name") && body?.file_name) {
    sets.push("file_name = ?");
    values.push(String(body.file_name).slice(0, 255));
  }
  for (const dim of ["width", "height"] as const) {
    const n = Number(body?.[dim]);
    if (cols.has(dim) && Number.isFinite(n) && n > 0) {
      sets.push(`${dim} = ?`);
      values.push(Math.round(n));
    }
  }
  // Re-editable drawings: the draw tool sends the vector strokes (anno_json)
  // and the untouched base image (anno_base_url). Any other image replacement
  // (e.g. crop) sends neither, which CLEARS them — the strokes no longer match
  // the new picture, so re-editing must start fresh.
  if (cols.has("anno_json")) {
    sets.push("anno_json = ?");
    values.push(typeof body?.anno_json === "string" && body.anno_json ? String(body.anno_json) : null);
  }
  if (cols.has("anno_base_url")) {
    const base = body?.anno_base_url ? canonicalS3Url(String(body.anno_base_url)) : "";
    sets.push("anno_base_url = ?");
    values.push(base && isImageUrl(base) ? base : null);
  }
  // Capture the pre-edit image as the WRITE-ONCE original the first time this
  // photo is modified, so the clean original is always kept for view/restore.
  if (cols.has("original_url")) {
    try {
      const [curRows] = await pool.query(
        "SELECT url, original_url FROM report_photos WHERE id = ? AND report_id = ? LIMIT 1",
        [photoId, reportId]
      );
      const curRow =
        Array.isArray(curRows) && curRows[0] ? (curRows[0] as Record<string, unknown>) : null;
      const existingOriginal = curRow?.original_url ? String(curRow.original_url).trim() : "";
      const curUrl = curRow?.url ? canonicalS3Url(String(curRow.url)) : "";
      if (!existingOriginal && curUrl && curUrl !== url) {
        sets.push("original_url = ?");
        values.push(curUrl);
      }
    } catch (err) {
      console.error("[api/reports/:id/photos] original_url capture failed:", err);
    }
  }
  values.push(photoId, reportId);

  const [result] = await pool.query(
    `UPDATE report_photos SET ${sets.join(", ")} WHERE id = ? AND report_id = ?`,
    values
  );
  if (!Number((result as { affectedRows?: number })?.affectedRows || 0)) {
    return Response.json({ error: "Photo not found" }, { status: 404 });
  }
  const [rows] = await pool.query("SELECT * FROM report_photos WHERE id = ? LIMIT 1", [photoId]);
  return Response.json({
    ok: true,
    photo: Array.isArray(rows) && rows[0] ? await signRowUrls("report_photos", rows[0]) : null,
  });
}

/**
 * Erase a drawing (or crop) and bring back the ORIGINAL photo. The draw & crop
 * tools only swap the row's `url` (they upload a new object and never delete the
 * old one) and never touch `image_key`, so the untouched original is often still
 * reachable — via `anno_base_url` (drawings made after that feature) or via
 * `image_key` (the key the photo was first uploaded under). We verify the
 * candidate actually exists before swapping it in, then clear any stored strokes.
 */
// Resolve the untouched original image URL for a photo (canonical), or null if
// none is recoverable. Tries, in order: the write-once `original_url`, the
// drawing base (`anno_base_url`), then the first-upload `image_key`. Each
// candidate must differ from the current (edited) image and actually exist.
async function findOriginalUrl(reportId: string, photoId: string): Promise<string | null> {
  await ensureAnnoColumns();
  const cols = await getColumns();
  const [rows] = await pool.query(
    "SELECT * FROM report_photos WHERE id = ? AND report_id = ? LIMIT 1",
    [photoId, reportId]
  );
  const row = Array.isArray(rows) && rows[0] ? (rows[0] as Record<string, unknown>) : null;
  if (!row) return null;
  const currentUrl = canonicalS3Url(String(row.url || ""));
  const candidates: string[] = [];
  if (cols.has("original_url") && row.original_url) candidates.push(canonicalS3Url(String(row.original_url)));
  if (cols.has("anno_base_url") && row.anno_base_url) candidates.push(canonicalS3Url(String(row.anno_base_url)));
  if (cols.has("image_key") && row.image_key) {
    const ik = String(row.image_key).trim();
    if (/^https?:\/\//i.test(ik)) candidates.push(canonicalS3Url(ik));
    else if (ik) {
      try {
        candidates.push(getPublicS3Url(ik));
      } catch {
        /* bucket url not configured — skip */
      }
    }
  }
  for (const c of candidates) {
    if (!c || canonicalS3Url(c) === currentUrl) continue;
    if (!s3KeyFromUrl(c) && !/^https?:\/\//i.test(c)) continue;
    try {
      const head = await fetch(c, { method: "HEAD" });
      if (head.ok) return canonicalS3Url(c) as string;
    } catch {
      /* not reachable — try the next candidate */
    }
  }
  return null;
}

async function restoreOriginalPhoto(reportId: string, body: Record<string, unknown>) {
  const photoId = String(body?.photoId || body?.id || "").trim();
  if (!photoId) return Response.json({ error: "photoId is required" }, { status: 400 });

  const restored = await findOriginalUrl(reportId, photoId);
  if (!restored) {
    return Response.json(
      {
        error:
          "No clean original is stored for this photo, so the drawing can't be removed automatically. Re-upload the original photo to replace it.",
      },
      { status: 409 }
    );
  }

  const cols = await getColumns();
  const sets: string[] = ["url = ?"];
  const values: unknown[] = [canonicalS3Url(restored)];
  if (cols.has("file_name")) {
    sets.push("file_name = ?");
    values.push(String(restored).split("/").pop()?.slice(0, 255) || null);
  }
  if (cols.has("anno_json")) sets.push("anno_json = NULL");
  if (cols.has("anno_base_url")) sets.push("anno_base_url = NULL");
  values.push(photoId, reportId);
  await pool.query(
    `UPDATE report_photos SET ${sets.join(", ")} WHERE id = ? AND report_id = ?`,
    values
  );
  const [after] = await pool.query("SELECT * FROM report_photos WHERE id = ? LIMIT 1", [photoId]);
  return Response.json({
    ok: true,
    restored: true,
    photo: Array.isArray(after) && after[0] ? await signRowUrls("report_photos", after[0]) : null,
  });
}

// PATCH /api/reports/:id/photos
//   { photoId, url, width?, height?, file_name? } — replace a photo's image
//     in place (draw / crop tools).
//   { photoId, include } or { selections: [{ id, include }, ...] } — update
//     the per-photo Word-export selection.
export async function PATCH(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const reportId = String(context.params?.id || "").trim();
    if (!reportId) return Response.json({ error: "Report id is required" }, { status: 400 });

    const body = await request.json().catch(() => ({} as any));
    if (body?.restoreOriginal) return await restoreOriginalPhoto(reportId, body);
    if (typeof body?.url === "string") return await replacePhotoImage(reportId, body);

    const selections: Array<{ id: string; include: boolean }> = Array.isArray(body?.selections)
      ? body.selections
          .map((s: any) => ({ id: String(s?.id || "").trim(), include: !!s?.include }))
          .filter((s: any) => s.id)
      : [];
    const singleId = String(body?.photoId || body?.id || "").trim();
    if (!selections.length && singleId) {
      selections.push({ id: singleId, include: !!body?.include });
    }
    if (!selections.length) {
      return Response.json({ error: "photoId (or selections[]) is required" }, { status: 400 });
    }

    await ensureIncludeInExportColumn();
    let updated = 0;
    for (const s of selections) {
      const [result] = await pool.query(
        "UPDATE report_photos SET include_in_export = ? WHERE id = ? AND report_id = ?",
        [s.include ? 1 : 0, s.id, reportId]
      );
      updated += Number((result as { affectedRows?: number })?.affectedRows || 0);
    }
    return Response.json({ ok: true, updated });
  } catch (error) {
    if (unauthorized(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/reports/:id/photos] PATCH error:", error);
    return Response.json({ error: "Failed to update photo" }, { status: 500 });
  }
}

// DELETE /api/reports/:id/photos?photoId=... — remove a single photo from a
// report. The photo id can be passed as a query param or in the JSON body.
export async function DELETE(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const reportId = String(context.params?.id || "").trim();
    if (!reportId) return Response.json({ error: "Report id is required" }, { status: 400 });

    const url = new URL(request.url);
    let photoId = String(url.searchParams.get("photoId") || "").trim();
    if (!photoId) {
      const body = await request.json().catch(() => ({} as any));
      photoId = String(body?.photoId || body?.id || "").trim();
    }
    if (!photoId) return Response.json({ error: "Photo id is required" }, { status: 400 });

    const [result] = await pool.query(
      "DELETE FROM report_photos WHERE id = ? AND report_id = ?",
      [photoId, reportId]
    );
    const affectedRows = Number((result as { affectedRows?: number })?.affectedRows || 0);
    if (!affectedRows) {
      return Response.json({ error: "Photo not found" }, { status: 404 });
    }

    return Response.json({ ok: true, deleted: photoId });
  } catch (error) {
    if (unauthorized(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/reports/:id/photos] DELETE error:", error);
    return Response.json({ error: "Failed to delete report photo" }, { status: 500 });
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
        url: canonicalS3Url(String(p.url || "").trim()),
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
        photos: await signRowUrls("report_photos", Array.isArray(verifyRows) ? verifyRows : []),
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
