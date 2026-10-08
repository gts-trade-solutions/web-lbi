/* eslint-disable @typescript-eslint/no-explicit-any */
// Bulk-set each report's KM from an uploaded Excel sheet. Pairs with the
// "KM + Coords (Excel)" download (coordinates-xlsx): download it, fix the KM
// column, upload it back here. Rows are matched to points by S.No (the point's
// position / point_key), and the number in the KM column is written to the
// report's `kms`. Blank KM cells are left unchanged.
import { NextResponse } from "next/server";
import * as XLSX from "xlsx";
import pool from "../../../../../lib/db";
import { requireAuth } from "../../../../../lib/auth";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 120;

type Ctx = { params: { id: string } };

async function reportsHasKms(): Promise<boolean> {
  try {
    const [rows] = await pool.query("SHOW COLUMNS FROM reports LIKE 'kms'");
    return Array.isArray(rows) && rows.length > 0;
  } catch {
    return false;
  }
}

function parseKm(v: unknown): number | null {
  if (v === null || typeof v === "undefined") return null;
  const s = String(v).trim();
  if (s === "") return null;
  // Keep digits, dot, minus (handles "774.24", "774.24 km", "KM 12.5").
  const cleaned = s.replace(/[^0-9.\-]/g, "");
  if (cleaned === "" || cleaned === "-" || cleaned === ".") return null;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

export async function POST(request: Request, context: Ctx) {
  const projectId = String(context.params?.id || "").trim();
  try {
    requireAuth(request);
  } catch {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }
  if (!projectId) {
    return NextResponse.json({ error: "Project id is required" }, { status: 400 });
  }

  // Read the uploaded spreadsheet.
  let file: File | null = null;
  try {
    const form = await request.formData();
    const f = form.get("file");
    if (f instanceof File) file = f;
  } catch {
    /* ignore */
  }
  if (!file) {
    return NextResponse.json({ error: "Upload an .xlsx / .xls / .csv file in the 'file' field" }, { status: 400 });
  }
  if (!/\.(xlsx|xls|csv)$/i.test(file.name)) {
    return NextResponse.json({ error: "Only .xlsx, .xls or .csv files are supported" }, { status: 400 });
  }

  // Make sure the KM column exists (added lazily like elsewhere).
  if (!(await reportsHasKms())) {
    try {
      await pool.query("ALTER TABLE reports ADD COLUMN kms DOUBLE NULL");
    } catch (e) {
      console.error("[km-upload] could not add kms column:", e);
      return NextResponse.json({ error: "The KM column could not be set up on the database." }, { status: 500 });
    }
  }

  // Parse the sheet into rows.
  let aoa: any[][] = [];
  try {
    const buffer = Buffer.from(await file.arrayBuffer());
    const wb = XLSX.read(buffer, { type: "buffer" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    aoa = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: "" }) as any[][];
  } catch (e) {
    console.error("[km-upload] parse failed:", e);
    return NextResponse.json({ error: "Could not read the spreadsheet." }, { status: 400 });
  }
  if (!aoa.length) {
    return NextResponse.json({ error: "The spreadsheet is empty." }, { status: 400 });
  }

  // The sheet MUST have a proper header row with a KM column (as produced by
  // the "KM + Coords (Excel)" download). This rejects random / dummy files.
  const header = (aoa[0] || []).map((h) => String(h || "").toLowerCase().trim());
  const BAD_SHEET =
    "This doesn't look like a KM sheet. Click \"KM + Coords (Excel)\" to download it, edit the KM column, then upload that same file.";
  const kmCol = header.findIndex((h) => h === "km" || h === "kms" || /\bkms?\b/.test(h) || /kilom/.test(h));
  if (kmCol < 0) {
    return NextResponse.json({ error: BAD_SHEET }, { status: 422 });
  }
  // Point-number column (S.No). Default to the first column (the download puts
  // S.No there) only when a header row is clearly present.
  let snoCol = header.findIndex((h) => /s\.?\s*no|serial|point|key|^no\.?$|^#$/.test(h));
  if (snoCol < 0) snoCol = 0;

  // Build pointNumber -> km from the data rows (skip the header row).
  const kmByNum = new Map<number, number>();
  for (let i = 1; i < aoa.length; i += 1) {
    const row = aoa[i] || [];
    const num = Number(String(row[snoCol] ?? "").trim());
    const km = parseKm(row[kmCol]);
    if (Number.isFinite(num) && km != null) kmByNum.set(num, km);
  }
  if (!kmByNum.size) {
    return NextResponse.json(
      { error: "No KM values found in the sheet. Put the point number in the S.No column and the KM value in the KM column." },
      { status: 422 }
    );
  }

  // Load this project's reports in the SAME order as the Excel download so the
  // S.No lines up, and update each matched report's KM.
  let reports: any[] = [];
  try {
    const [rows] = await pool.query(
      "SELECT id, point_key FROM reports WHERE project_id = ? ORDER BY sort_order ASC, created_at ASC",
      [projectId]
    );
    reports = Array.isArray(rows) ? (rows as any[]) : [];
  } catch (e) {
    console.error("[km-upload] load reports failed:", e);
    return NextResponse.json({ error: "Could not load the project's points." }, { status: 500 });
  }

  let updated = 0;
  for (let i = 0; i < reports.length; i += 1) {
    const r = reports[i];
    // Match by position (S.No = i+1) OR by the point's own point_key number.
    const pkNum = Number(String(r.point_key ?? "").trim());
    const km = kmByNum.has(i + 1)
      ? kmByNum.get(i + 1)!
      : Number.isFinite(pkNum) && kmByNum.has(pkNum)
        ? kmByNum.get(pkNum)!
        : null;
    if (km == null) continue;
    try {
      await pool.query("UPDATE reports SET kms = ? WHERE id = ?", [km, r.id]);
      updated += 1;
    } catch (e) {
      console.error("[km-upload] update failed for", r.id, e);
    }
  }

  if (updated === 0) {
    return NextResponse.json(
      {
        error:
          "No points were updated — the point numbers (S.No) in the file don't match this project's points. Make sure you downloaded the sheet from THIS project.",
      },
      { status: 422 }
    );
  }

  return NextResponse.json({
    ok: true,
    updated,
    totalPoints: reports.length,
    valuesInFile: kmByNum.size,
  });
}
