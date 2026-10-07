/* eslint-disable @typescript-eslint/no-explicit-any */
// Convert an OLD survey report PDF into the SAME structured shape the .docx
// importer produces (DocxReport), so the existing project-creation machinery in
// app/api/projects/import-docx/route.ts works unchanged.
//
// These PDFs lay out ONE survey point per small table ("Kms | Location |
// Description | Remarks" header, a data row, a "Co Ordinates :- lat, lon" line),
// with the road photos below it. Some pages carry TWO such tables side by side
// (two points). We read the text with pdfjs (positions preserved), split each
// page into tables by each "Kms" header, bucket the data into columns by their
// center-x (the cells are centre-aligned), and pull the embedded photos out of
// the page — dropping the repeated header/footer logos by position + size.
//
// pdfjs-dist is already a dependency; sharp re-encodes the extracted bitmaps.
import sharp from "sharp";
import type { DocxPoint, DocxReport, FrontImage } from "./docxReportImport";
import { inferCategory, parseCoord, splitCoordPair } from "./docxReportImport";

// Column header -> field. The survey tables always use these four.
const FIELD_RE: { f: keyof typeof EMPTY_BUCKET; re: RegExp }[] = [
  { f: "km", re: /^kms?\.?$/i },
  { f: "location", re: /^location$/i },
  { f: "observation", re: /^(description|details?)$/i },
  { f: "remarks", re: /^(remarks?|action)$/i },
];
const EMPTY_BUCKET = { km: "", location: "", observation: "", remarks: "" };

type Item = { x: number; y: number; w: number; cx: number; s: string };
type Row = { y: number; items: Item[] };
type HeaderCol = { f: "km" | "location" | "observation" | "remarks"; cx: number; x: number };

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

// Difficulty from the observation + remarks wording. An action that must be
// taken (shutdown / open / remove / develop / regulate …) makes a point YELLOW;
// an outright block makes it RED; otherwise GREEN ("Normal Pass"). "Go straight"
// is only a driving instruction, so it does NOT by itself mean green.
function inferDifficulty(text: string): string {
  const t = String(text || "").toLowerCase();
  if (/(don'?t|do not)\s*go|not\s*(feasible|possible)|cannot|\bstop\b|blocked|no\s*movement|dead\s*end/.test(t))
    return "red";
  if (
    /shut\s*down|shutdown|need\s*to\s*be|needs?\s*to\s*be|required|require\b|remove|removed|caution|careful|restrict|diversion|detour|shift|\blift\b|regulate|work\s*(is\s*)?in\s*progress|develop|compact|open(ed)?\b|height\s*restriction|escort|puller|trans-?shipment|bypass/.test(
      t
    )
  )
    return "yellow";
  return "green";
}

// A "Co Ordinates :- 23.01, 70.21" or "Co-ordinates: N23 54.2 E78 45.3" value.
const COORD_LABEL_RE = /co[\s-]*ordinate/i;
function coordFromText(s: string): { lat: number | null; lon: number | null; raw: string } {
  const raw = String(s || "").replace(COORD_LABEL_RE, "").replace(/^[\s:–—-]+/, "").trim();
  // Plain decimal pair (the common case): "23.013210, 70.213991".
  const dec = raw.match(/(-?\d{1,3}(?:\.\d+)?)\s*[,/]\s*(-?\d{1,3}(?:\.\d+)?)/);
  if (dec) {
    const lat = Number(dec[1]);
    const lon = Number(dec[2]);
    if (Number.isFinite(lat) && Number.isFinite(lon) && Math.abs(lat) <= 90 && Math.abs(lon) <= 180)
      return { lat, lon, raw };
  }
  // Hemisphere / DMS fallback, reusing the docx coordinate logic.
  const pair = splitCoordPair(raw);
  if (pair) {
    const lat = parseCoord(pair[0]);
    const lon = parseCoord(pair[1]);
    if (lat != null && lon != null) return { lat, lon, raw };
  }
  return { lat: null, lon: null, raw };
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

// Resolve a pdfjs image object; grouped images (drawn inside a transparency
// group) never settle without a full render, so give up quickly on those.
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

async function encodeImage(obj: any): Promise<Buffer | null> {
  const { width, height, kind, data } = obj || {};
  if (!width || !height || !data) return null;
  const channels = kind === 3 ? 4 : kind === 2 ? 3 : null; // 2=RGB, 3=RGBA
  if (!channels) return null;
  const expected = width * height * channels;
  if (data.length < expected) return null;
  try {
    const raw = Buffer.from(data.buffer, data.byteOffset, expected);
    const out = await sharp(raw, { raw: { width, height, channels } })
      .jpeg({ quality: 82 })
      .toBuffer();
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

  const points: DocxPoint[] = [];
  const photoMap = new Map<string, Buffer>();
  let objective = "";

  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    try {
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
      if (!items.length) continue;
      const rows = groupByY(items);

      // Objective — the Executive Summary paragraph (first time we see it). Skip
      // the Table-of-Contents page, which also lists "Executive Summary".
      if (
        !objective &&
        items.some((i) => /executive\s*summary/i.test(i.s)) &&
        !items.some((i) => /table\s*of\s*content|^\s*contents?\s*$/i.test(i.s))
      ) {
        const txt = rows
          .map((r) => r.items.map((i) => i.s).join(" "))
          .filter((l) => !/^\s*executive\s*summary\s*$|feasibility study report|abnormal movements|rev\s*\d|page\s*no\./i.test(l))
          .join(" ")
          .replace(/\s+/g, " ")
          .trim();
        if (txt.length > 300) objective = txt.slice(0, 1500);
      }

      // Find the header row (may hold 1 or 2 tables' headers at the same y).
      let headerRow: Row | null = null;
      for (const row of rows) {
        const found = new Set<string>();
        for (const it of row.items) for (const { f, re } of FIELD_RE) if (re.test(it.s)) found.add(f);
        if (found.has("km") && found.has("observation") && (found.has("remarks") || found.has("location"))) {
          headerRow = row;
          break;
        }
      }
      if (!headerRow) continue; // not a survey-point page

      // Build the per-table column sets. Each "Kms" label starts a table.
      const labelled: (HeaderCol & { any: true })[] = [];
      for (const it of headerRow.items) {
        for (const { f, re } of FIELD_RE)
          if (re.test(it.s)) labelled.push({ f, cx: it.cx, x: it.x, any: true });
      }
      const kmLefts = labelled.filter((l) => l.f === "km").map((l) => l.x).sort((a, b) => a - b);
      if (!kmLefts.length) continue;
      const tableBounds = kmLefts.map((left, i) => ({
        left,
        right: i + 1 < kmLefts.length ? kmLefts[i + 1] : Infinity,
      }));
      const tables = tableBounds.map((b) => ({
        ...b,
        cols: labelled.filter((l) => l.cx >= b.left - 2 && l.cx < b.right).map((l) => ({ f: l.f, cx: l.cx })),
        buckets: { km: [] as string[], location: [] as string[], observation: [] as string[], remarks: [] as string[] },
        coordRaw: "",
        lat: null as number | null,
        lon: null as number | null,
        photos: [] as string[],
      }));
      const tableFor = (cx: number) => {
        for (let i = tables.length - 1; i >= 0; i--) if (cx >= tables[i].left - 2) return tables[i];
        return tables[0];
      };

      // Data rows: directly below the header, stop at the first big vertical gap
      // (which is where the photos / dimension labels begin).
      const below = rows.filter((r) => r.y < headerRow!.y - 1);
      let lastY = headerRow.y;
      for (const row of below) {
        if (lastY - row.y > 30) break;
        lastY = row.y;
        const joined = row.items.map((i) => i.s).join(" ");
        if (COORD_LABEL_RE.test(joined)) {
          // One coord line per table — split the row into per-table coord
          // strings by x and assign each to its table.
          for (const t of tables) {
            const sub = row.items
              .filter((it) => tableFor(it.cx) === t)
              .map((it) => it.s)
              .join(" ");
            if (COORD_LABEL_RE.test(sub) && !t.coordRaw) {
              const c = coordFromText(sub);
              t.coordRaw = c.raw;
              t.lat = c.lat;
              t.lon = c.lon;
            }
          }
          continue;
        }
        for (const it of row.items) {
          const t = tableFor(it.cx);
          let best: HeaderCol["f"] | null = null;
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

      // Photos: walk the operator list, track the CTM, keep real photos (large,
      // in the body) and drop the repeated header/footer logos by position/size.
      const ops = await page.getOperatorList();
      let ctm = [1, 0, 0, 1, 0, 0];
      const stack: number[][] = [];
      const imgDraws: { id: string; cx: number; cy: number; rw: number; rh: number }[] = [];
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
          imgDraws.push({ id: ops.argsArray[i][0], cx, cy, rw, rh });
        }
      }
      let imgIdx = 0;
      for (const d of imgDraws) {
        // Branding: in/above the table header band, in the footer, tiny, or a
        // very wide/tall strip (a logo banner).
        const minR = Math.min(d.rw, d.rh);
        const aspect = minR > 0 ? Math.max(d.rw, d.rh) / minR : 99;
        if (d.cy >= headerRow.y - 6 || d.cy <= 80 || minR < 40 || aspect > 4) continue;
        const obj = await getImageObj(page, d.id);
        const buf = obj ? await encodeImage(obj) : null;
        if (!buf) continue;
        const key = `pdf-photo-p${p}-${imgIdx++}`;
        photoMap.set(key, buf);
        tableFor(d.cx).photos.push(key);
      }

      // Emit one point per table, left -> right.
      for (const t of tables.sort((a, b) => a.left - b.left)) {
        const km = (t.buckets.km.join(" ").match(/-?\d+(?:\.\d+)?/) || [""])[0];
        const location = t.buckets.location.join(" ").replace(/\s+/g, " ").trim();
        const observation = t.buckets.observation.join(" ").replace(/\s+/g, " ").trim();
        const remarks = t.buckets.remarks.join(" ").replace(/\s+/g, " ").trim();
        // Skip an empty table (header with no data row under it).
        if (!km && !location && !observation && !remarks && !t.photos.length) continue;
        const allText = `${location} ${observation} ${remarks}`;
        points.push({
          point_key: String(points.length + 1),
          coordRaw: t.coordRaw,
          latitude: t.lat,
          longitude: t.lon,
          km,
          location,
          category: inferCategory(allText),
          observation,
          remarks,
          difficulty: inferDifficulty(`${observation} ${remarks}`),
          photoNames: t.photos,
        });
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

  const getPhoto = async (key: string): Promise<Buffer | null> => photoMap.get(key) || null;
  const frontImages: FrontImage[] = [];
  return { points, frontImages, objective, getPhoto };
}
