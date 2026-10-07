/* eslint-disable @typescript-eslint/no-explicit-any */
// Convert an OLD survey report PDF into the SAME structured shape the .docx
// importer produces (DocxReport), so the existing project-creation machinery in
// app/api/projects/import-docx/route.ts works unchanged.
//
// Handles the survey-report table layouts seen so far:
//   • "Kms | Location | Description | Remarks"            (coords on a
//      "Co Ordinates :- lat, lon" line; one or two points per page)
//   • "GPS LOCATION | KM | LOCATION | CATEGORY | OBSERVATION | REMARKS/ACTION"
//      (coords in the GPS column, an explicit category column, wrapped headers,
//      and "Satellite Image" continuation pages that belong to the point above)
//
// Text + table geometry come from pdfjs (positions preserved). The PHOTOS are
// captured EXACTLY as the PDF shows them — including the yellow measurement
// arrows / "H-6.2m" labels, which are drawn ON TOP of the raster as PDF
// vectors/text — by rendering each page with Ghostscript and cropping the photo
// region. When Ghostscript isn't available we fall back to extracting the bare
// raster via pdfjs (loses the drawn overlays).
import sharp from "sharp";
import { execFile, execFileSync } from "child_process";
import { promisify } from "util";
import { promises as fsp } from "fs";
import os from "os";
import path from "path";
import { randomUUID } from "crypto";
import type { DocxPoint, DocxReport, FrontImage } from "./docxReportImport";
import { inferCategory, parseCoord, splitCoordPair } from "./docxReportImport";

const execFileAsync = promisify(execFile);

// Column header -> field. Order-independent; matched against each header word.
const FIELD_RE: { f: Field; re: RegExp }[] = [
  { f: "coord", re: /^(gps(\s*location)?|co[\s-]?o?rdinates?.*)$/i },
  { f: "km", re: /^kms?\.?$/i },
  { f: "location", re: /^location$/i },
  { f: "category", re: /^category$/i },
  { f: "observation", re: /^(observations?|description|details?)$/i },
  { f: "remarks", re: /^(remarks?|action|remarks?\s*\/?\s*action)$/i },
];
type Field = "coord" | "km" | "location" | "category" | "observation" | "remarks";
// When several header words stack at one x (a wrapped header like "GPS" over
// "LOCATION"), keep the highest-priority field so "GPS LOCATION" is a coord
// column, not a stray location column.
const PRIO: Field[] = ["coord", "km", "category", "observation", "remarks", "location"];

type Item = { x: number; y: number; w: number; cx: number; s: string };
type Row = { y: number; items: Item[] };
type Rect = { id: string; cx: number; cy: number; rw: number; rh: number };
type PhotoTask = { page: number; pointIdx: number; rect: Rect; pageH: number };

const COORD_LABEL_RE = /co[\s-]*ordinate/i;

function groupByY(items: Item[]): Row[] {
  const rows: Row[] = [];
  for (const it of items) {
    let row = rows.find((r) => Math.abs(r.y - it.y) < 4);
    if (!row) {
      row = { y: it.y, items: [] };
      rows.push(row);
    }
    row.items.push(it);
  }
  for (const r of rows) r.items.sort((a, b) => a.x - b.x);
  return rows.sort((a, b) => b.y - a.y);
}

// Difficulty from observation + remarks wording. An action that must be taken
// (shutdown / open / remove / caution / feasibility concern …) -> YELLOW; an
// outright block -> RED; otherwise GREEN ("Normal Pass").
function inferDifficulty(text: string): string {
  const t = String(text || "").toLowerCase();
  if (/(don'?t|do not)\s*go|not\s*(feasible|possible)|cannot|\bstop\b|blocked|no\s*movement|dead\s*end/.test(t))
    return "red";
  if (
    /shut\s*down|shutdown|need\s*to\s*be|needs?\s*to\s*be|required|require\b|remove|removed|caution|careful|restrict|diversion|detour|shift|\blift\b|regulate|work\s*(is\s*)?in\s*progress|develop|compact|open(ed)?\b|height\s*restriction|escort|puller|trans-?shipment|bypass|feasibility\s*concern|obstruct/.test(
      t
    )
  )
    return "yellow";
  return "green";
}

// Parse "N22 30.870 E88 18.222" or "23.01, 70.21" -> decimal lat/lon.
function coordFromText(s: string): { lat: number | null; lon: number | null } {
  const raw = String(s || "").replace(COORD_LABEL_RE, "").replace(/^[\s:–—-]+/, "").trim();
  const dec = raw.match(/(-?\d{1,3}(?:\.\d+)?)\s*[,/]\s*(-?\d{1,3}(?:\.\d+)?)/);
  if (dec) {
    const lat = Number(dec[1]);
    const lon = Number(dec[2]);
    if (Number.isFinite(lat) && Number.isFinite(lon) && Math.abs(lat) <= 90 && Math.abs(lon) <= 180)
      return { lat, lon };
  }
  const pair = splitCoordPair(raw);
  if (pair) {
    const lat = parseCoord(pair[0]);
    const lon = parseCoord(pair[1]);
    if (lat != null && lon != null && Math.abs(lat) <= 90 && Math.abs(lon) <= 180) return { lat, lon };
  }
  return { lat: null, lon: null };
}

// pdfjs Util.transform(m1, m2): compose two 2-D affine matrices.
function mul(m1: number[], m2: number[]): number[] {
  return [
    m1[0] * m2[0] + m1[2] * m2[1],
    m1[1] * m2[0] + m1[3] * m2[1],
    m1[0] * m2[2] + m1[2] * m2[3],
    m1[1] * m2[2] + m1[3] * m2[3],
    m1[0] * m2[4] + m1[2] * m2[5] + m1[4],
    m1[1] * m2[4] + m1[3] * m2[5] + m1[5],
  ];
}

// ---- Ghostscript (exact page render) ----
let _gsBin: string | null | undefined;
function resolveGs(): string | null {
  if (_gsBin !== undefined) return _gsBin;
  const candidates = [
    "gs",
    "/usr/bin/gs",
    "gswin64c",
    "gswin64c.exe",
    "C:/Program Files/gs/gs10.05.1/bin/gswin64c.exe",
  ];
  for (const c of candidates) {
    try {
      execFileSync(c, ["--version"], { stdio: "ignore", timeout: 15000 });
      _gsBin = c;
      return c;
    } catch {
      /* try next */
    }
  }
  _gsBin = null;
  return null;
}

const RENDER_DPI = 130;

// Render the needed pages (in concurrent chunks) and crop each photo's rect,
// filling photoMap. Returns the key assigned to each task (in task order).
async function renderAndCrop(
  gsBin: string,
  pdfPath: string,
  tasks: PhotoTask[],
  photoMap: Map<string, Buffer>
): Promise<string[]> {
  const scale = RENDER_DPI / 72;
  const keys: (string | null)[] = tasks.map(() => null);
  const tasksByPage = new Map<number, number[]>(); // page -> task indexes
  tasks.forEach((t, i) => {
    if (!tasksByPage.has(t.page)) tasksByPage.set(t.page, []);
    tasksByPage.get(t.page)!.push(i);
  });
  const pages = [...tasksByPage.keys()].sort((a, b) => a - b);
  const CHUNK = 20;
  const chunks: number[][] = [];
  for (let i = 0; i < pages.length; i += CHUNK) chunks.push(pages.slice(i, i + CHUNK));

  const tmpRoot = path.join(os.tmpdir(), `lbi-pdf-${randomUUID()}`);
  await fsp.mkdir(tmpRoot, { recursive: true });

  const doChunk = async (ci: number) => {
    const chunk = chunks[ci];
    const dir = path.join(tmpRoot, `c${ci}`);
    await fsp.mkdir(dir, { recursive: true });
    try {
      await execFileAsync(
        gsBin,
        [
          "-q",
          "-dNOPAUSE",
          "-dBATCH",
          "-dSAFER",
          "-sDEVICE=jpeg",
          "-dJPEGQ=90",
          `-r${RENDER_DPI}`,
          `-sPageList=${chunk.join(",")}`,
          "-o",
          path.join(dir, "%d.jpg"),
          pdfPath,
        ],
        { timeout: 120000, maxBuffer: 1 << 24 }
      );
      for (let k = 0; k < chunk.length; k++) {
        const page = chunk[k];
        const img = path.join(dir, `${k + 1}.jpg`);
        let meta: sharp.Metadata;
        try {
          meta = await sharp(img).metadata();
        } catch {
          continue; // page not rendered
        }
        const mw = meta.width || 0;
        const mh = meta.height || 0;
        for (const ti of tasksByPage.get(page)!) {
          const { rect: d, pageH } = tasks[ti];
          const padW = d.rw * 0.015;
          const padH = d.rh * 0.02;
          const x0 = d.cx - d.rw / 2 - padW;
          const x1 = d.cx + d.rw / 2 + padW;
          const yTop = pageH - (d.cy + d.rh / 2 + padH); // PDF y is bottom-up
          const yBot = pageH - (d.cy - d.rh / 2 - padH);
          const left = Math.max(0, Math.round(x0 * scale));
          const top = Math.max(0, Math.round(yTop * scale));
          let w = Math.round((x1 - x0) * scale);
          let h = Math.round((yBot - yTop) * scale);
          w = Math.min(w, mw - left);
          h = Math.min(h, mh - top);
          if (w < 40 || h < 40) continue;
          try {
            const buf = await sharp(img)
              .extract({ left, top, width: w, height: h })
              .jpeg({ quality: 85 })
              .toBuffer();
            if (buf.length < 1200) continue;
            const key = `pdf-photo-p${page}-${ti}.jpg`;
            photoMap.set(key, buf);
            keys[ti] = key;
          } catch {
            /* skip this crop */
          }
        }
      }
    } finally {
      await fsp.rm(dir, { recursive: true, force: true }).catch(() => {});
    }
  };

  // Run chunks with limited concurrency (3 Ghostscript processes at a time).
  const CONC = 3;
  let next = 0;
  const worker = async () => {
    while (next < chunks.length) {
      const my = next++;
      await doChunk(my).catch((e) => console.warn("[pdf-import] render chunk failed:", e?.message || e));
    }
  };
  await Promise.all(Array.from({ length: Math.min(CONC, chunks.length) }, () => worker()));
  await fsp.rm(tmpRoot, { recursive: true, force: true }).catch(() => {});
  return keys.map((k) => k || "");
}

// ---- Raster fallback (no Ghostscript): extract the bare image bytes ----
function getImageObj(page: any, id: string): Promise<any> {
  return new Promise((resolve) => {
    let done = false;
    const fin = (o: any) => {
      if (!done) {
        done = true;
        resolve(o);
      }
    };
    try {
      if (page.objs?.has?.(id)) return fin(page.objs.get(id));
    } catch {
      /* not ready */
    }
    try {
      page.objs.get(id, (o: any) => fin(o));
    } catch {
      fin(null);
    }
    setTimeout(() => fin(null), 40);
  });
}
async function encodeRaster(obj: any): Promise<Buffer | null> {
  const { width, height, kind, data } = obj || {};
  if (!width || !height || !data) return null;
  const channels = kind === 3 ? 4 : kind === 2 ? 3 : null;
  if (!channels) return null;
  const expected = width * height * channels;
  if (data.length < expected) return null;
  try {
    const raw = Buffer.from(data.buffer, data.byteOffset, expected);
    const out = await sharp(raw, { raw: { width, height, channels } }).jpeg({ quality: 85 }).toBuffer();
    return out.length >= 1200 ? out : null;
  } catch {
    return null;
  }
}

export async function parsePdfReport(buffer: Buffer): Promise<DocxReport> {
  const pdfjs: any = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const OPS = pdfjs.OPS;
  const doc = await pdfjs.getDocument({
    data: new Uint8Array(buffer),
    isEvalSupported: false,
    disableFontFace: true,
    useSystemFonts: false,
  }).promise;

  const gsBin = resolveGs();
  const points: DocxPoint[] = [];
  const photoMap = new Map<string, Buffer>();
  const photoTasks: PhotoTask[] = [];
  let objective = "";
  let started = false;
  let lastPointIdx = -1;

  // ---- PASS 1: text -> points, and collect photo rects (or raster bytes) ----
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    try {
      const pageH = page.getViewport({ scale: 1 }).height;
      const tc = await page.getTextContent();
      const items: Item[] = tc.items
        .filter((i: any) => i.str && i.str.trim())
        .map((i: any) => ({
          x: i.transform[4],
          y: i.transform[5],
          w: i.width || 0,
          cx: i.transform[4] + (i.width || 0) / 2,
          s: i.str.trim(),
        }));
      const rows = groupByY(items);

      // Objective — "OBJECTIVE:" or an "Executive Summary" page (not the TOC).
      if (
        !objective &&
        (items.some((i) => /objective\s*:/i.test(i.s)) ||
          (items.some((i) => /executive\s*summary/i.test(i.s)) &&
            !items.some((i) => /table\s*of\s*content/i.test(i.s))))
      ) {
        const txt = rows
          .map((r) => r.items.map((i) => i.s).join(" "))
          .filter((l) => !/report by|raceinnovations|www\.|^dated\b|^rev\b|abnormal movements|page\s*(no\.?\s*)?\d/i.test(l))
          .join(" ")
          .replace(/.*objective\s*:?\s*/i, "")
          .replace(/\s*route\s*map.*/i, "")
          .replace(/^\s*executive\s*summary\s*/i, "")
          .replace(/\s+/g, " ")
          .trim();
        if (txt.length > 40) objective = txt.slice(0, 1500);
      }

      // Find the header row (the one with the most column labels).
      let mainRow: Row | null = null;
      let mainCount = 0;
      for (const row of rows) {
        const found = new Set<Field>();
        for (const it of row.items) for (const { f, re } of FIELD_RE) if (re.test(it.s)) found.add(f);
        const ok =
          (found.has("km") && found.has("observation")) ||
          (found.has("km") && found.has("location") && found.has("coord"));
        if (ok && found.size > mainCount) {
          mainCount = found.size;
          mainRow = row;
        }
      }

      // Image rectangles on this page (track the CTM; no decode yet).
      const ops = await page.getOperatorList();
      let ctm = [1, 0, 0, 1, 0, 0];
      const stack: number[][] = [];
      const imgs: Rect[] = [];
      for (let i = 0; i < ops.fnArray.length; i++) {
        const fn = ops.fnArray[i];
        if (fn === OPS.save) stack.push(ctm.slice());
        else if (fn === OPS.restore) ctm = stack.pop() || ctm;
        else if (fn === OPS.transform) ctm = mul(ctm, ops.argsArray[i]);
        else if (
          fn === OPS.paintImageXObject ||
          fn === OPS.paintJpegXObject ||
          fn === OPS.paintImageXObjectRepeat
        ) {
          const cx = ctm[0] * 0.5 + ctm[2] * 0.5 + ctm[4];
          const cy = ctm[1] * 0.5 + ctm[3] * 0.5 + ctm[5];
          const rw = Math.hypot(ctm[0], ctm[1]);
          const rh = Math.hypot(ctm[2], ctm[3]);
          imgs.push({ id: ops.argsArray[i][0], cx, cy, rw, rh });
        }
      }
      // Body photos: drop the repeated header/footer logos (tiny, banner-shaped,
      // or in the top/bottom margin).
      const bodyPhotos = imgs.filter((d) => {
        const minR = Math.min(d.rw, d.rh);
        const aspect = minR > 0 ? Math.max(d.rw, d.rh) / minR : 99;
        return !(minR < 110 || aspect > 4 || d.cy > pageH * 0.83 || d.cy < pageH * 0.08);
      });

      // queue a photo for a point (render in pass 2, or raster-decode now).
      const queuePhoto = async (rect: Rect, pointIdx: number) => {
        if (gsBin) {
          photoTasks.push({ page: p, pointIdx, rect, pageH });
        } else {
          const obj = await getImageObj(page, rect.id);
          const buf = obj ? await encodeRaster(obj) : null;
          if (buf) {
            const key = `pdf-photo-p${p}-${photoTasks.length}.jpg`;
            photoMap.set(key, buf);
            points[pointIdx].photoNames.push(key);
            photoTasks.push({ page: p, pointIdx, rect, pageH }); // keep count/order
          }
        }
      };

      if (!mainRow) {
        // Continuation page (e.g. "Satellite Image:") -> its photos belong to
        // the most recent point.
        if (started && lastPointIdx >= 0 && bodyPhotos.length) {
          for (const d of bodyPhotos.sort((a, b) => b.cy - a.cy || a.cx - b.cx)) await queuePhoto(d, lastPointIdx);
        }
        continue;
      }
      started = true;

      // Gather the header words within a small band around the main row, then
      // cluster them by x into one field per column.
      const band = 16;
      const rawLabels: { f: Field; cx: number; x: number; y: number }[] = [];
      for (const row of rows) {
        if (Math.abs(row.y - mainRow.y) > band) continue;
        for (const it of row.items)
          for (const { f, re } of FIELD_RE) if (re.test(it.s)) rawLabels.push({ f, cx: it.cx, x: it.x, y: it.y });
      }
      const headerBottomY = Math.min(...rawLabels.map((l) => l.y));
      rawLabels.sort((a, b) => a.cx - b.cx);
      const labelled: { f: Field; cx: number; x: number }[] = [];
      for (const l of rawLabels) {
        const near = labelled.find((k) => Math.abs(k.cx - l.cx) < 28);
        if (!near) {
          labelled.push({ f: l.f, cx: l.cx, x: l.x });
        } else if (PRIO.indexOf(l.f) < PRIO.indexOf(near.f)) {
          near.f = l.f;
          near.cx = l.cx;
          near.x = l.x;
        }
      }
      const kmLefts = labelled.filter((l) => l.f === "km").map((l) => l.x).sort((a, b) => a - b);
      if (!kmLefts.length) continue;

      const tables = kmLefts.map((left, i) => {
        const l = i === 0 ? -Infinity : left;
        const right = i + 1 < kmLefts.length ? kmLefts[i + 1] : Infinity;
        return {
          left: l,
          right,
          cols: labelled.filter((c) => c.cx >= (l === -Infinity ? -Infinity : l - 2) && c.cx < right),
          buckets: { km: [], location: [], category: [], observation: [], remarks: [], coord: [] } as Record<Field, string[]>,
          coordRaw: "",
          photos: [] as Rect[],
        };
      });
      const tableFor = (cx: number) => {
        for (let i = tables.length - 1; i >= 0; i--)
          if (cx >= (tables[i].left === -Infinity ? -Infinity : tables[i].left - 2)) return tables[i];
        return tables[0];
      };
      const inPhoto = (it: Item) =>
        bodyPhotos.some(
          (d) => it.cx >= d.cx - d.rw / 2 && it.cx <= d.cx + d.rw / 2 && it.y >= d.cy - d.rh / 2 && it.y <= d.cy + d.rh / 2
        );
      const isFooter = (s: string) =>
        /report by|raceinnovations|www\.|^dated\b|^rev\b|abnormal movements|page\s*(no\.?\s*)?\d/i.test(s);

      const below = rows.filter((r) => r.y < headerBottomY - 1 && !isFooter(r.items.map((i) => i.s).join(" ")));
      for (const row of below) {
        const joined = row.items.map((i) => i.s).join(" ");
        if (COORD_LABEL_RE.test(joined)) {
          for (const t of tables) {
            const sub = row.items
              .filter((it) => tableFor(it.cx) === t)
              .map((it) => it.s)
              .join(" ");
            if (COORD_LABEL_RE.test(sub) && !t.coordRaw) t.coordRaw = sub;
          }
          continue;
        }
        for (const it of row.items) {
          if (inPhoto(it)) continue;
          const t = tableFor(it.cx);
          let best: Field | null = null;
          let bestD = Infinity;
          for (const c of t.cols) {
            const d = Math.abs(c.cx - it.cx);
            if (d < bestD) {
              bestD = d;
              best = c.f;
            }
          }
          if (best) t.buckets[best].push(it.s);
        }
      }

      for (const d of bodyPhotos) tableFor(d.cx).photos.push(d);

      const sortedTables = tables.sort(
        (a, b) => (a.left === -Infinity ? -1 : a.left) - (b.left === -Infinity ? -1 : b.left)
      );
      for (const t of sortedTables) {
        const km = (t.buckets.km.join(" ").match(/-?\d+(?:\.\d+)?/) || [""])[0];
        const location = t.buckets.location.join(" ").replace(/\s+/g, " ").trim();
        const observation = t.buckets.observation.join(" ").replace(/\s+/g, " ").trim();
        const remarks = t.buckets.remarks.join(" ").replace(/\s+/g, " ").trim();
        const catCol = t.buckets.category.join(" ").replace(/\s+/g, " ").trim();
        let coord = { lat: null as number | null, lon: null as number | null };
        if (t.buckets.coord.length) coord = coordFromText(t.buckets.coord.join(" "));
        if (coord.lat == null && t.coordRaw) coord = coordFromText(t.coordRaw);
        if (!km && !location && !observation && !remarks && !catCol && !t.photos.length) continue;
        const category = catCol || inferCategory(`${location} ${observation} ${remarks}`);
        points.push({
          point_key: String(points.length + 1),
          coordRaw: t.coordRaw,
          latitude: coord.lat,
          longitude: coord.lon,
          km,
          location,
          category,
          observation,
          remarks,
          difficulty: inferDifficulty(`${observation} ${remarks}`),
          photoNames: [],
        });
        lastPointIdx = points.length - 1;
        for (const d of t.photos.sort((a, b) => b.cy - a.cy || a.cx - b.cx)) await queuePhoto(d, lastPointIdx);
      }
    } finally {
      try {
        page.cleanup();
      } catch {
        /* ignore */
      }
    }
  }

  try {
    await doc.destroy();
  } catch {
    /* ignore */
  }

  // ---- PASS 2: render the pages with Ghostscript and crop each photo ----
  if (gsBin && photoTasks.length) {
    let pdfPath = "";
    try {
      pdfPath = path.join(os.tmpdir(), `lbi-src-${randomUUID()}.pdf`);
      await fsp.writeFile(pdfPath, buffer);
      const keys = await renderAndCrop(gsBin, pdfPath, photoTasks, photoMap);
      keys.forEach((key, i) => {
        if (key) points[photoTasks[i].pointIdx].photoNames.push(key);
      });
    } catch (err) {
      console.warn("[pdf-import] Ghostscript render failed; photos may be missing:", (err as any)?.message || err);
    } finally {
      if (pdfPath) await fsp.rm(pdfPath, { force: true }).catch(() => {});
    }
  }

  const getPhoto = async (key: string): Promise<Buffer | null> => photoMap.get(key) || null;
  const frontImages: FrontImage[] = [];
  return { points, frontImages, objective, getPhoto };
}
