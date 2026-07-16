/* eslint-disable @typescript-eslint/no-explicit-any */
// Convert an OLD survey report .docx into structured points + photos, ready to
// import as a new website project. Each observation point in these reports is a
// small 2-row table (header row + data row) with columns like GPS
// LOCATION / KM / LOCATION / CATEGORY / OBSERVATION / REMARKS. The category
// cell holds a tiny category ICON (excluded); the real road photos are the
// images that sit between one point's table and the next.
//
// Reads the .docx with pizzip (already a dependency via docxtemplater). No new
// packages required. Handles the coordinate formats these reports use, incl.
// "N12 59.312E80 10.350" (lat and lon run together) and DMS.
import PizZip from "pizzip";

export type DocxPoint = {
  point_key: string;
  coordRaw: string;
  latitude: number | null;
  longitude: number | null;
  km: string;
  location: string;
  category: string;
  observation: string;
  remarks: string;
  difficulty: string; // green | yellow | red
  photoNames: string[]; // zip paths like "word/media/image3.png"
};

export type DocxReport = {
  points: DocxPoint[];
  // Read a photo's bytes out of the same .docx.
  getPhoto: (zipPath: string) => Buffer | null;
};

// Build the relationship-id → media-target map from a .docx/.pptx _rels XML.
// Extracts Id and Target per <Relationship> element ORDER-INDEPENDENTLY — some
// Word/PPT writers emit `Target="..." Id="..."` (Target first), which an
// Id-then-Target regex silently misses, dropping every image (0 photos).
function buildRelMap(relsXml: string): Record<string, string> {
  const relMap: Record<string, string> = {};
  for (const m of Array.from(relsXml.matchAll(/<Relationship\b[^>]*>/g))) {
    const el = m[0];
    const id = el.match(/\bId="([^"]+)"/);
    const target = el.match(/\bTarget="([^"]+)"/);
    if (id && target) relMap[id[1]] = target[1];
  }
  return relMap;
}

// Column header → field. Order-independent; matched against each header cell.
const HEADER_KEYS: Record<string, RegExp> = {
  // Match a real coordinate column but NOT a "GPS No" point-number column.
  // Tolerates the real header spellings seen across reports: "NE COORDINATE",
  // "NE CO Ordinates" (space), "NE CORDINATE" (typo, one O), "Co-ordinates",
  // "GPS LOCATION".
  coord: /co[\s-]?o?rdinate|gps\s*location/i,
  km: /^km|kms/i,
  place: /^location$/i,
  category: /category/i,
  obs: /observation|detail/i,
  remarks: /remark|action/i,
  photo: /^photo/i,
  vm: /vehicle|movement/i,
};

function decodeXml(s: string): string {
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function cellText(tc: string | undefined): string {
  if (!tc) return "";
  // Paragraph ends, line breaks and tabs are turned into the U+0001 sentinel,
  // read back as a space, so text on separate lines is not glued together —
  // e.g. "Signboard crossing" + "height 7m" must read "...crossing height 7m",
  // not "...crossingheight 7m". Runs WITHIN one line still join with no gap so a
  // word split across runs is not broken apart. Only <w:t> text is taken.
  const marked = tc.replace(/<\/w:p>|<w:br\b[^>]*\/?>|<w:tab\b[^>]*\/?>|<w:cr\b[^>]*\/?>/g, "\u0001");
  let result = "";
  for (const m of Array.from(marked.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>|\u0001/g))) {
    result += m[1] !== undefined ? m[1] : " ";
  }
  return decodeXml(result).replace(/\s+/g, " ").trim();
}

// A single coordinate token → decimal degrees. Handles N/S/E/W (either end),
// degrees+decimal-minutes, and degrees-minutes-seconds.
function parseCoord(v: string): number | null {
  let raw = String(v || "").trim().replace(/^[\s"',]+|[\s"',]+$/g, "");
  if (!raw) return null;
  let sign = 1;
  const dir = raw.match(/[NSEW]/i);
  if (dir) {
    const c = dir[0].toUpperCase();
    if (c === "S" || c === "W") sign = -1;
    raw = raw.replace(/[NSEW]/gi, " ");
  }
  const nums = raw
    .replace(/[°'"′″]/g, " ")
    .split(/\s+/)
    .map((x) => x.trim())
    .filter((x) => x !== "" && Number.isFinite(Number(x)))
    .map(Number);
  if (!nums.length) return null;
  const deg = Math.abs(nums[0]);
  const dec =
    nums.length === 1 ? deg : nums.length === 2 ? deg + nums[1] / 60 : deg + nums[1] / 60 + nums[2] / 3600;
  return dir ? sign * dec : nums[0] < 0 ? -dec : dec;
}

// Split a coordinate cell into [lat, lon]. Copes with:
//   • lat & lon run together, dir-first: "N12 59.312E80 10.350"
//   • space-separated dir-first: "N12 53.397 E79 54.775"
//   • DMS dir-last: "12°55'21.5\"N 79°52'29.2\"E"
function splitCoordPair(raw: string): [string, string] | null {
  const s0 = String(raw || "").trim();
  if (!s0) return null;
  // A: break BEFORE a hemisphere letter that follows a digit (start of 2nd coord).
  let a = s0.replace(/(\d)\s*([EWNS])/gi, "$1|$2");
  // B: break AFTER a hemisphere letter that precedes a digit (dir-last style).
  a = a.replace(/([NSEW])\s+(?=\d)/gi, "$1|");
  let parts = a.split("|").map((x) => x.trim()).filter(Boolean);
  if (parts.length !== 2) parts = s0.split(/[,;]+/).map((x) => x.trim()).filter(Boolean);
  return parts.length === 2 ? [parts[0], parts[1]] : null;
}

// TPI-style reports carry no CATEGORY column — infer it from the detail text so
// the imported report still gets a sensible category (user can fix in the grid).
function inferCategory(text: string): string {
  const c = String(text || "").toLowerCase();
  if (/side\s*sign/.test(c)) return "Side Signboard";
  if (/electric\s*sign/.test(c)) return "Electric Sign Board";
  if (/sign\s*board|signboard/.test(c)) return "Signboard";
  if (/signal\s*pole|signal/.test(c)) return "Signal Pole";
  if (/underpass/.test(c)) return "Underpass Bridge";
  if (/footpath/.test(c)) return "Footpath Bridge";
  if (/river\s*bridge/.test(c)) return "River Bridge";
  if (/railway|rail\s*over|level\s*cross/.test(c)) return "Railway Level Crossing";
  if (/flyover|over\s*bridge|overpass/.test(c)) return "Flyover";
  if (/\bbridge\b|girder|rcc\s*slab/.test(c)) return "Bridge";
  if (/high\s*tension|ht\s*cable/.test(c)) return "High Tension Cable";
  if (/low\s*tension|lt\s*cable/.test(c)) return "Low Tension Cable";
  if (/tower\s*line|towerline/.test(c)) return "Towerline Cable";
  if (/toll/.test(c)) return "Toll Plaza";
  if (/junction/.test(c)) return "Junction";
  if (/diversion|detour/.test(c)) return "Take Diversion";
  if (/\bbend\b|curve/.test(c)) return "Bend";
  if (/petrol|fuel/.test(c)) return "Petrol bunk";
  if (/parking/.test(c)) return "Parking Place";
  if (/service\s*(road|lane)/.test(c)) return "Service road";
  if (/metro/.test(c)) return "Metro";
  return "Report";
}

// Classify a 6-hex colour as green/yellow/red (survey difficulty), or null for
// white/black/greys and other colours (e.g. blue headers, black text) we ignore.
// Copes with the real shades these reports use: pure & pale yellow (FFFF00,
// FFFFCC, FFF2CC), greens (00B050, 56BE9F) and reds (FF0000, EE0000).
function hexToDifficulty(hex: string): string | null {
  const h = hex.toUpperCase();
  if (!/^[0-9A-F]{6}$/.test(h)) return null;
  const r = parseInt(h.slice(0, 2), 16), g = parseInt(h.slice(2, 4), 16), b = parseInt(h.slice(4, 6), 16);
  if (Math.max(r, g, b) - Math.min(r, g, b) < 30) return null; // white / black / grey → no signal
  if (r >= 150 && r - g > 60 && r - b > 60 && g < 170) return "red";
  if (r >= 170 && g >= 140 && g - b > 30 && r - g < 90) return "yellow";
  if (g >= 110 && g >= r && g - b > 30) return "green";
  return null;
}

// Difficulty for one table row (Word) or slide-table row (PPT). The colour that
// marks a point green/yellow/red lives in different places across the report
// templates, so we look at ALL of them and let the most severe win
// (red > yellow > green):
//   • Word cell shading        <w:shd w:fill="RRGGBB"/>
//   • Word text highlight       <w:highlight w:val="yellow|red|green"/>
//   • DrawingML shape/box fill  <a:srgbClr val="RRGGBB"/>   (Word shapes & PPT)
//   • VML shape fill            fillcolor="#RRGGBB"
// Returns null when the row carries no colour signal (caller then defaults to
// green = "normal pass").
function difficultyFromXml(xml: string): string | null {
  const votes: Record<string, boolean> = {};
  const add = (d: string | null) => { if (d) votes[d] = true; };
  for (const m of Array.from(xml.matchAll(/<w:highlight w:val="([^"]+)"/g))) {
    const v = m[1].toLowerCase();
    if (v === "yellow" || v === "red" || v === "green") votes[v] = true;
  }
  for (const m of Array.from(xml.matchAll(/<w:shd\b[^>]*w:fill="([0-9A-Fa-f]{6})"/g))) add(hexToDifficulty(m[1]));
  for (const m of Array.from(xml.matchAll(/<a:srgbClr val="([0-9A-Fa-f]{6})"/g))) add(hexToDifficulty(m[1]));
  for (const m of Array.from(xml.matchAll(/fillcolor="#?([0-9A-Fa-f]{6})"/g))) add(hexToDifficulty(m[1]));
  return votes.red ? "red" : votes.yellow ? "yellow" : votes.green ? "green" : null;
}

// Slice a block of XML into chunks that each start at one occurrence of a tag.
function sliceByTag(text: string, tagRe: RegExp): string[] {
  const re = new RegExp(tagRe.source, tagRe.flags.includes("g") ? tagRe.flags : tagRe.flags + "g");
  const starts: number[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) starts.push(m.index);
  return starts.map((s, k) => text.slice(s, k + 1 < starts.length ? starts[k + 1] : text.length));
}

export function parseDocxReport(buffer: Buffer): DocxReport {
  const zip = new PizZip(buffer);
  const xml = zip.file("word/document.xml")?.asText() || "";
  const relsXml = zip.file("word/_rels/document.xml.rels")?.asText() || "";

  // relationship id -> media target (relative to word/)
  const relMap = buildRelMap(relsXml);
  const imagesIn = (s: string): string[] => {
    const ids: string[] = [];
    for (const m of Array.from(s.matchAll(/r:embed="([^"]+)"/g))) ids.push(m[1]);
    for (const m of Array.from(s.matchAll(/r:id="([^"]+)"/g))) ids.push(m[1]);
    return ids
      .map((id) => relMap[id])
      .filter((t) => t && /media\//.test(t))
      .map((t) => "word/" + t.replace(/^\/*/, ""));
  };

  // Top-level tables (nesting-aware).
  type Tbl = { start: number; end: number; xml: string };
  const tables: Tbl[] = [];
  let idx = 0;
  while (true) {
    const start = xml.indexOf("<w:tbl>", idx);
    if (start < 0) break;
    let depth = 0;
    let end = -1;
    const re = /<w:tbl>|<\/w:tbl>/g;
    re.lastIndex = start;
    let mm: RegExpExecArray | null;
    while ((mm = re.exec(xml))) {
      if (mm[0] === "<w:tbl>") depth++;
      else {
        depth--;
        if (!depth) {
          end = mm.index + mm[0].length;
          break;
        }
      }
    }
    if (end < 0) break;
    tables.push({ start, end, xml: xml.slice(start, end) });
    idx = end;
  }

  const rowsOf = (t: string): string[] => {
    const starts: number[] = [];
    const re = /<w:tr[ >]/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(t))) starts.push(m.index);
    return starts.map((s, k) => t.slice(s, k + 1 < starts.length ? starts[k + 1] : t.length));
  };
  const cellsOf = (r: string): string[] => {
    const starts: number[] = [];
    const re = /<w:tc>/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(r))) starts.push(m.index);
    return starts.map((s, k) => r.slice(s, k + 1 < starts.length ? starts[k + 1] : r.length));
  };

  const points: DocxPoint[] = [];
  let lastPointEnd = 0;

  const cleanLocation = (s: string) =>
    s.replace(/^\s*(location\s*[-:]\s*)+/i, "").replace(/\s{2,}/g, " ").trim();

  for (const t of tables) {
    const rows = rowsOf(t.xml);
    if (rows.length < 2) continue;
    const header = cellsOf(rows[0]).map((c) => cellText(c));
    const headerJoined = header.join(" ");
    const isData = /gps|coordinate/i.test(headerJoined) && /km/i.test(headerJoined);
    if (!isData) continue;

    // Column map from the header.
    const map: Record<string, number> = {};
    header.forEach((h, i) => {
      for (const k in HEADER_KEYS) if (HEADER_KEYS[k].test(h) && map[k] == null) map[k] = i;
    });

    // A "between-table" photo (Tambaram layout: one point per 2-row table, photo
    // sits after the table) belongs to the PREVIOUS point.
    if (points.length) {
      points[points.length - 1].photoNames.push(...imagesIn(xml.slice(lastPointEnd, t.start)));
    }

    // Emit one point PER data row (handles TPI's single big table with many
    // rows AND Tambaram's 2-row tables alike).
    for (let ri = 1; ri < rows.length; ri++) {
      const dataCells = cellsOf(rows[ri]);
      // Skip stray empty rows (spacer rows).
      if (!dataCells.length || dataCells.every((c) => !cellText(c) && !imagesIn(c).length)) continue;

      // A merged full-width row (far fewer cells than the header) is a
      // continuation / detail block of the point ABOVE it — e.g. TPI's
      // "TOPOGRAPHY / TURNING CIRCLE DRAWING / CIVIL WORK" row that carries the
      // junction's extra drawings. It is NOT a new point: folding its images
      // (and text) into the previous point keeps those photos with their point
      // and stops every later point's number from shifting by one.
      const isMergedRow = dataCells.length < header.length && dataCells.length <= 2;
      if (isMergedRow) {
        if (points.length) {
          const prev = points[points.length - 1];
          for (const c of dataCells) prev.photoNames.push(...imagesIn(c));
          const extra = dataCells.map((c) => cellText(c)).join(" ").replace(/\s+/g, " ").trim();
          if (extra) prev.observation = prev.observation ? `${prev.observation} — ${extra}` : extra;
        }
        continue;
      }

      const get = (k: string) => (map[k] != null ? cellText(dataCells[map[k]]) : "");

      const coordRaw = get("coord");
      const pair = splitCoordPair(coordRaw);
      const latitude = pair ? parseCoord(pair[0]) : null;
      const longitude = pair ? parseCoord(pair[1]) : null;

      // Location: from the LOCATION column, else scan cells for "Location-…".
      let location = cleanLocation(get("place"));
      if (!location) {
        for (const c of dataCells) {
          const tx = cellText(c);
          if (/^location\s*[-:]/i.test(tx)) {
            location = cleanLocation(tx);
            break;
          }
        }
      }

      // Category: use the column if present, else infer from the detail text.
      const observation = get("obs");
      let category = get("category");
      if (!category) category = inferCategory(observation);

      // Photos: every image in this data row EXCEPT the category-icon cell.
      const rowPhotos: string[] = [];
      dataCells.forEach((c, ci) => {
        if (ci === map.category || ci === map.coord) return;
        rowPhotos.push(...imagesIn(c));
      });

      // Difficulty: read the colour ONLY from the "VEHICLE MOVEMENT" column's
      // green/yellow/red square. Scanning the WHOLE row wrongly picked up the
      // yellow highlight on photo labels (e.g. "W-7m"/"W-14m"), turning green
      // points yellow. When there's no such column (e.g. Tambaram, whose colour
      // is uniform cell shading) fall back to scanning the whole row.
      const vmCols = header
        .map((h, ci) => (/vehicle|movement/i.test(h) ? ci : -1))
        .filter((ci) => ci >= 0);
      // Some layouts have TWO "VEHICLE MOVEMENT" columns — one holds the PHOTOS
      // (whose yellow "W-7m" labels must NOT count), the other the colour
      // square. Prefer VM cells that contain NO image (the square). Fall back to
      // all VM cells, then to the whole row (Tambaram-style uniform shading).
      const vmSquareCells = vmCols
        .map((ci) => dataCells[ci] || "")
        .filter((cell) => imagesIn(cell).length === 0);
      const diffXml = vmSquareCells.length
        ? vmSquareCells.join("")
        : vmCols.length
          ? vmCols.map((ci) => dataCells[ci] || "").join("")
          : rows[ri];
      const difficulty = difficultyFromXml(diffXml) || "green";

      points.push({
        point_key: String(points.length + 1),
        coordRaw,
        latitude,
        longitude,
        km: get("km"),
        location,
        category,
        observation,
        remarks: get("remarks"),
        difficulty,
        photoNames: rowPhotos,
      } as DocxPoint);
    }
    lastPointEnd = t.end;
  }
  // Trailing photos after the last point's table.
  if (points.length) points[points.length - 1].photoNames.push(...imagesIn(xml.slice(lastPointEnd)));

  // De-duplicate photo names per point.
  for (const p of points) p.photoNames = Array.from(new Set(p.photoNames));

  const getPhoto = (zipPath: string): Buffer | null => {
    const f = zip.file(zipPath);
    if (!f) return null;
    try {
      return Buffer.from(f.asUint8Array());
    } catch {
      return null;
    }
  };

  return { points, getPhoto };
}

// Same reports are sometimes delivered as PowerPoint (.pptx): one survey point
// per slide, each slide carrying its own little table (GPS Location / KM /
// LOCATION / OBSERVATIONS / REMARKS) plus 1–3 road photos as picture shapes.
// The company logo/legend repeats on every slide, so images used on many slides
// are treated as branding and excluded; the per-slide unique images are the
// real photos.
export function parsePptxReport(buffer: Buffer): DocxReport {
  const zip = new PizZip(buffer);

  const slideNames = Object.keys((zip as any).files)
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort(
      (a, b) => parseInt(a.match(/slide(\d+)/)![1], 10) - parseInt(b.match(/slide(\d+)/)![1], 10)
    );

  // First pass: read each slide's XML + its photo targets, and count how often
  // each image is used across all slides (logos repeat → branding).
  type SlideInfo = { xml: string; media: string[] };
  const slides: SlideInfo[] = [];
  const freq: Record<string, number> = {};
  for (const sn of slideNames) {
    const sx = zip.file(sn)?.asText() || "";
    const relsPath = sn.replace(/slides\/(slide\d+)\.xml/, "slides/_rels/$1.xml.rels");
    const relsXml = zip.file(relsPath)?.asText() || "";
    const relMap = buildRelMap(relsXml);
    const ids: string[] = [];
    for (const m of Array.from(sx.matchAll(/r:embed="([^"]+)"/g))) ids.push(m[1]);
    const media = Array.from(
      new Set(
        ids
          .map((id) => relMap[id])
          .filter((t) => t && /media\//.test(t))
          .map((t) => "ppt/" + t.replace(/^\.\.\//, "").replace(/^\/*/, ""))
      )
    );
    media.forEach((m) => (freq[m] = (freq[m] || 0) + 1));
    slides.push({ xml: sx, media });
  }
  const brandingCut = Math.max(3, Math.floor(slideNames.length * 0.1));
  const isBranding = (m: string) => (freq[m] || 0) > brandingCut;

  const aCellText = (tc: string): string => {
    const parts: string[] = [];
    for (const m of Array.from(tc.matchAll(/<a:t>([^<]*)<\/a:t>/g))) parts.push(m[1]);
    return decodeXml(parts.join(" ")).replace(/\s+/g, " ").trim();
  };
  const cleanLocation = (s: string) =>
    s
      .replace(/^\s*(location\s*[-:]\s*)+/i, "")
      .replace(/^[\s,]+/, "")
      .replace(/\s+([,;])/g, "$1")
      .replace(/\s{2,}/g, " ")
      .trim();

  const points: DocxPoint[] = [];

  for (const slide of slides) {
    const ts = slide.xml.indexOf("<a:tbl>");
    if (ts < 0) continue; // title / section slide, no data table
    const teRaw = slide.xml.indexOf("</a:tbl>", ts);
    const tbl = teRaw < 0 ? slide.xml.slice(ts) : slide.xml.slice(ts, teRaw + 8);

    const rows = sliceByTag(tbl, /<a:tr[ >]/);
    if (rows.length < 2) continue;
    const header = sliceByTag(rows[0], /<a:tc[ >]/).map(aCellText);
    const headerJoined = header.join(" ");
    if (!(/gps|coordinate/i.test(headerJoined) && /km/i.test(headerJoined))) continue;

    const map: Record<string, number> = {};
    header.forEach((h, i) => {
      for (const k in HEADER_KEYS) if (HEADER_KEYS[k].test(h) && map[k] == null) map[k] = i;
    });

    // Photos for this slide = every image on it that isn't branding. They go on
    // the FIRST real data row of the slide (tracked with a flag so a blank
    // leading row doesn't cause the photos to be dropped).
    const slidePhotos = slide.media.filter((m) => !isBranding(m));
    let slidePhotosAssigned = false;

    for (let ri = 1; ri < rows.length; ri++) {
      const cells = sliceByTag(rows[ri], /<a:tc[ >]/);
      if (!cells.length || cells.every((c) => !aCellText(c))) continue;
      const get = (k: string) => (map[k] != null ? aCellText(cells[map[k]]) : "");

      const coordRaw = get("coord");
      const pair = splitCoordPair(coordRaw);
      const latitude = pair ? parseCoord(pair[0]) : null;
      const longitude = pair ? parseCoord(pair[1]) : null;

      const observation = get("obs");
      let category = get("category");
      if (!category) category = inferCategory(observation);

      points.push({
        point_key: String(points.length + 1),
        coordRaw,
        latitude,
        longitude,
        km: get("km"),
        location: cleanLocation(get("place")),
        category,
        observation,
        remarks: get("remarks"),
        // Colour comes from the row's own cells only — the yellow highlight on
        // photo labels (H-7m etc.) sits OUTSIDE the table, so it can't leak in.
        difficulty: difficultyFromXml(rows[ri]) || "green",
        // The slide's photos go on its first real data row only, so a rare
        // multi-row slide doesn't duplicate them onto every point.
        photoNames: slidePhotosAssigned ? [] : slidePhotos.slice(),
      } as DocxPoint);
      slidePhotosAssigned = true;
    }
  }

  const getPhoto = (zipPath: string): Buffer | null => {
    const f = zip.file(zipPath);
    if (!f) return null;
    try {
      return Buffer.from(f.asUint8Array());
    } catch {
      return null;
    }
  };

  return { points, getPhoto };
}

// Pick the right parser from the file name / bytes.
export function parseSurveyReport(buffer: Buffer, fileName: string): DocxReport {
  return /\.pptx$/i.test(fileName) ? parsePptxReport(buffer) : parseDocxReport(buffer);
}
