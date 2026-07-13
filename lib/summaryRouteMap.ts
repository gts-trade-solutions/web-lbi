import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  AlignmentType,
  PageOrientation,
} from "docx";
import pool from "./db";

// ------- sharp (lazy; SVG->PNG needs libvips+librsvg) -------
// eslint-disable-next-line @typescript-eslint/no-explicit-any
let _sharp: any = null;
let _sharpTried = false;
function getSharp() {
  if (_sharpTried) return _sharp;
  _sharpTried = true;
  try {
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    _sharp = require("sharp");
  } catch (err) {
    console.warn("[summaryRouteMap] sharp unavailable:", err);
    _sharp = null;
  }
  return _sharp;
}

type Row = Record<string, any>;

export type SummaryRouteMapOptions = {
  projectId: string;
  reportIds?: string[];
};

export type SummaryRouteMapResult = {
  buffer: Buffer;
  fileName: string;
  projectName: string;
};

// ---- helpers ----------------------------------------------------------

function xmlEscape(s: string): string {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function hslToHex(h: number, s: number, l: number): string {
  h = ((h % 360) + 360) % 360;
  s /= 100;
  l /= 100;
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
  const m = l - c / 2;
  let r = 0;
  let g = 0;
  let b = 0;
  if (h < 60) [r, g, b] = [c, x, 0];
  else if (h < 120) [r, g, b] = [x, c, 0];
  else if (h < 180) [r, g, b] = [0, c, x];
  else if (h < 240) [r, g, b] = [0, x, c];
  else if (h < 300) [r, g, b] = [x, 0, c];
  else [r, g, b] = [c, 0, x];
  const to = (v: number) =>
    Math.round((v + m) * 255)
      .toString(16)
      .padStart(2, "0");
  return `#${to(r)}${to(g)}${to(b)}`;
}

const INDIAN_STATES = new Set([
  "andhra pradesh", "arunachal pradesh", "assam", "bihar", "chhattisgarh", "goa",
  "gujarat", "haryana", "himachal pradesh", "jharkhand", "karnataka", "kerala",
  "madhya pradesh", "maharashtra", "manipur", "meghalaya", "mizoram", "nagaland",
  "odisha", "punjab", "rajasthan", "sikkim", "tamil nadu", "telangana", "tripura",
  "uttar pradesh", "uttarakhand", "west bengal", "delhi", "jammu and kashmir",
  "ladakh", "puducherry", "chandigarh", "andaman and nicobar islands",
  "dadra and nagar haveli and daman and diu", "lakshadweep",
]);

// Parse a resolved-location string into a short place label + its state.
// Handles both the short "Alandur, Tamil Nadu" form and the long Nominatim
// form "St. Thomas Mount, Chennai, Tamil Nadu, 600016, India". Drops the
// postcode and country, and prefers a recognised Indian state for the band.
function parseLocation(value: unknown): { label: string; state: string } {
  const t = String(value ?? "").trim();
  if (!t) return { label: "", state: "" };

  const parts = t
    .split(",")
    .map((x) => x.trim())
    .filter(Boolean);

  // Coordinate-only fallback like "12.97, 77.59" — not a usable place name.
  if (parts.length <= 2 && parts.every((p) => /^-?\d+(\.\d+)?$/.test(p))) {
    return { label: "", state: "" };
  }

  // Drop postcodes and the country.
  const cleaned = parts.filter(
    (p) => !/^\d{4,6}$/.test(p) && p.toLowerCase() !== "india"
  );

  const label = cleaned[0] || parts[0] || "";
  // Prefer a recognised state name; otherwise use the last cleaned part.
  const stateMatch = cleaned.find((p) => INDIAN_STATES.has(p.toLowerCase()));
  const state = stateMatch || (cleaned.length > 1 ? cleaned[cleaned.length - 1] : "");
  return { label, state };
}

function firstNonEmpty(...vals: unknown[]): string {
  for (const v of vals) {
    const s = String(v ?? "").trim();
    if (s) return s;
  }
  return "";
}

// Wrap a label to at most 2 lines of ~maxChars, ellipsis when longer.
function wrapLabel(label: string, maxChars = 15): string[] {
  const words = label.split(/\s+/).filter(Boolean);
  const lines: string[] = [];
  let cur = "";
  for (const w of words) {
    if (!cur) cur = w;
    else if ((cur + " " + w).length <= maxChars) cur += " " + w;
    else {
      lines.push(cur);
      cur = w;
      if (lines.length === 2) break;
    }
  }
  if (cur && lines.length < 2) lines.push(cur);
  if (lines.length === 2 && words.join(" ").length > lines.join(" ").length) {
    let last = lines[1];
    if (last.length > maxChars - 1) last = last.slice(0, maxChars - 1);
    lines[1] = last + "…";
  }
  return lines.length ? lines : [label.slice(0, maxChars)];
}

function chunk<T>(arr: T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

// ---- SVG diagram ------------------------------------------------------

type Cell =
  | { kind: "start"; label: string }
  | { kind: "point"; n: number; label: string; state: string; hue: number };

const NAVY = "#16306b";
const NAVY_DARK = "#0c1f4d";
const LINE = "#1f3a93";

const PAD = 28;
const TITLE_H = 54;
const BOX_W = 188;
const BOX_H = 84;
const H_GAP = 46;
const HEADING_H = 30;
const ROW_V_GAP = 54;

function buildRouteMapSvg(
  title: string,
  startLabel: string,
  points: Array<{ label: string; state: string }>
): { svg: string; width: number; height: number } {
  const total = points.length;
  const cells: Cell[] = [
    { kind: "start", label: startLabel || "START" },
    ...points.map((p, i) => ({
      kind: "point" as const,
      n: i + 1,
      label: p.label || `Point ${i + 1}`,
      state: p.state,
      hue: 210 + (total > 1 ? (i / (total - 1)) * 300 : 0),
    })),
  ];

  const perRow = Math.min(8, Math.max(3, cells.length));
  const rows = chunk(cells, perRow);

  const width = PAD * 2 + perRow * BOX_W + (perRow - 1) * H_GAP;
  const contentTop = PAD + TITLE_H;
  const rowBlock = HEADING_H + BOX_H + ROW_V_GAP;
  const height = contentTop + rows.length * rowBlock - ROW_V_GAP + PAD;

  const cellX = (c: number) => PAD + c * (BOX_W + H_GAP);
  const rowBoxTop = (r: number) => contentTop + r * rowBlock + HEADING_H;

  const parts: string[] = [];
  parts.push(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">`
  );
  parts.push(
    `<defs><marker id="arrow" markerWidth="9" markerHeight="9" refX="6" refY="3" orient="auto" markerUnits="userSpaceOnUse"><path d="M0,0 L6,3 L0,6 Z" fill="${LINE}"/></marker></defs>`
  );
  parts.push(`<rect x="0" y="0" width="${width}" height="${height}" fill="#ffffff"/>`);

  // Title
  parts.push(
    `<text x="${width / 2}" y="${PAD + 30}" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-size="26" font-weight="700" fill="${NAVY}">${xmlEscape(
      title
    )}</text>`
  );

  const drawBox = (c: Cell, x: number, y: number) => {
    const isStart = c.kind === "start";
    const fill = isStart ? NAVY : hslToHex(c.hue, 58, 60);
    const stroke = isStart ? NAVY_DARK : hslToHex(c.hue, 58, 40);
    const numColor = stroke;
    parts.push(
      `<rect x="${x}" y="${y}" width="${BOX_W}" height="${BOX_H}" rx="12" ry="12" fill="${fill}" stroke="${stroke}" stroke-width="2"/>`
    );
    // number circle (points only)
    if (!isStart) {
      const cx = x + 20;
      const cy = y + 20;
      parts.push(`<circle cx="${cx}" cy="${cy}" r="12" fill="#ffffff"/>`);
      parts.push(
        `<text x="${cx}" y="${cy + 4}" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-size="12" font-weight="700" fill="${numColor}">${c.n}</text>`
      );
    }
    // label (white), wrapped to 2 lines, centered
    const lines = wrapLabel(c.label, isStart ? 16 : 15);
    const cxText = x + BOX_W / 2;
    const cyText = y + BOX_H / 2 + (isStart ? 0 : 8);
    if (lines.length === 1) {
      parts.push(
        `<text x="${cxText}" y="${cyText + 5}" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-size="14" font-weight="700" fill="#ffffff">${xmlEscape(
          lines[0]
        )}</text>`
      );
    } else {
      parts.push(
        `<text x="${cxText}" y="${cyText - 4}" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-size="13" font-weight="700" fill="#ffffff">${xmlEscape(
          lines[0]
        )}</text>`
      );
      parts.push(
        `<text x="${cxText}" y="${cyText + 14}" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-size="13" font-weight="700" fill="#ffffff">${xmlEscape(
          lines[1]
        )}</text>`
      );
    }
  };

  rows.forEach((row, r) => {
    const boxTop = rowBoxTop(r);

    // State heading: unique states across this row's point cells.
    const states: string[] = [];
    for (const c of row) {
      if (c.kind === "point" && c.state && !states.includes(c.state)) states.push(c.state);
    }
    if (states.length) {
      parts.push(
        `<text x="${width / 2}" y="${boxTop - 9}" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-size="15" font-weight="700" fill="${NAVY}" letter-spacing="1">${xmlEscape(
          states.join(" / ").toUpperCase()
        )}</text>`
      );
    }

    // boxes + intra-row arrows
    row.forEach((c, ci) => {
      const x = cellX(ci);
      drawBox(c, x, boxTop);
      if (ci < row.length - 1) {
        const x1 = x + BOX_W;
        const x2 = cellX(ci + 1);
        const y = boxTop + BOX_H / 2;
        parts.push(
          `<line x1="${x1 + 4}" y1="${y}" x2="${x2 - 4}" y2="${y}" stroke="${LINE}" stroke-width="2.5" marker-end="url(#arrow)"/>`
        );
      }
    });

    // elbow connector to next row (from last cell -> first cell of next row)
    if (r < rows.length - 1) {
      const lastCi = row.length - 1;
      const lx = cellX(lastCi) + BOX_W / 2;
      const by = boxTop + BOX_H;
      const nextTop = rowBoxTop(r + 1);
      const fx = cellX(0) + BOX_W / 2;
      const midY = by + ROW_V_GAP / 2;
      parts.push(
        `<path d="M ${lx} ${by} L ${lx} ${midY} L ${fx} ${midY} L ${fx} ${nextTop - 4}" fill="none" stroke="${LINE}" stroke-width="2.5" marker-end="url(#arrow)"/>`
      );
    }
  });

  parts.push(`</svg>`);
  return { svg: parts.join(""), width, height };
}

// ---- main -------------------------------------------------------------

export async function generateSummaryRouteMapDocx(
  options: SummaryRouteMapOptions
): Promise<SummaryRouteMapResult> {
  const projectId = options.projectId;
  const reportIds = (options.reportIds || []).filter(Boolean);

  // Project
  const [projRows] = await pool.query("SELECT * FROM projects WHERE id = ? LIMIT 1", [
    projectId,
  ]);
  const project = (Array.isArray(projRows) && projRows[0]) as Row | undefined;
  if (!project) throw new Error("Project not found");
  const projectName = String(
    project.name || project.title || project.project_name || "PROJECT"
  ).trim();

  // Reports (route order)
  let reports: Row[] = [];
  if (reportIds.length) {
    const ph = reportIds.map(() => "?").join(",");
    const [rows] = await pool.query(
      `SELECT * FROM reports WHERE project_id = ? AND id IN (${ph})
       ORDER BY sort_order ASC, created_at ASC`,
      [projectId, ...reportIds]
    );
    reports = Array.isArray(rows) ? (rows as Row[]) : [];
  } else {
    const [rows] = await pool.query(
      `SELECT * FROM reports WHERE project_id = ?
       ORDER BY sort_order ASC, created_at ASC`,
      [projectId]
    );
    reports = Array.isArray(rows) ? (rows as Row[]) : [];
  }

  // Skip the bulk-import NO_GPS holding report if present.
  reports = reports.filter((r) => String(r.point_key || "") !== "__NO_GPS__");

  if (!reports.length) throw new Error("No reports to summarize for this project");

  const points = reports.map((r, i) => {
    const locRaw = firstNonEmpty(
      r.resolved_location,
      r.location,
      r.location_name,
      r.address
    );
    const { label, state } = parseLocation(locRaw);
    return {
      label: label || firstNonEmpty(r.point_key, `Point ${i + 1}`),
      state,
    };
  });

  const title = `${projectName} – ROUTE MAPPING`.toUpperCase();
  const { svg, width, height } = buildRouteMapSvg(title, projectName, points);

  // Render SVG -> PNG
  const sharp = getSharp();
  if (!sharp) {
    throw new Error(
      "Image rendering is unavailable on the server (sharp/libvips with SVG support is required) — cannot build the summary route map."
    );
  }
  let png: Buffer;
  try {
    png = await sharp(Buffer.from(svg)).png().toBuffer();
  } catch (err) {
    console.error("[summaryRouteMap] SVG->PNG failed:", err);
    throw new Error(
      "Failed to render the route-map image. The server's image library may lack SVG support."
    );
  }

  // Fit the image onto one landscape A4 page (content ~ 10.2in x 7in @96dpi).
  const MAX_W = 980;
  const MAX_H = 660;
  const scale = Math.min(MAX_W / width, MAX_H / height, 1);
  const dispW = Math.round(width * scale);
  const dispH = Math.round(height * scale);

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size: { orientation: PageOrientation.LANDSCAPE },
            margin: { top: 480, bottom: 480, left: 480, right: 480 },
          },
        },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new ImageRun({
                type: "png",
                data: png,
                transformation: { width: dispW, height: dispH },
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: `${points.length} location(s)`,
                size: 18,
                color: "667085",
              }),
            ],
          }),
        ],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  const safeBase =
    projectName.replace(/[<>:"/\\|?*\x00-\x1F]/g, "_").replace(/\s+/g, " ").trim() ||
    "PROJECT";
  const fileName = `${safeBase}-SUMMARY.docx`;

  return { buffer, fileName, projectName };
}
