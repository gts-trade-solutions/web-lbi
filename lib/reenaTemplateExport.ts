import path from "path";
import { promises as fs } from "fs";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import pool from "./db";
import { getReadSignedUrl } from "./s3";

// docxtemplater-image-module-free has no TypeScript types shipped.
// eslint-disable-next-line @typescript-eslint/no-var-requires
const ImageModule = require("docxtemplater-image-module-free");

type Row = Record<string, any>;

export type ExportOptions = {
  projectId: string;
  reportIds?: string[];
  includePhotos?: boolean;
  debug?: boolean;
};

export type ExportResult = {
  buffer: Buffer;
  fileName: string;
  projectName: string;
};

const TEMPLATE_PATH = path.join(process.cwd(), "templates", "reena-all-template.docx");

const ROUTE_MAP_SIZE: [number, number] = [933, 700];
const GA_DRAWING_SIZE: [number, number] = [867, 650];
const OBSERVATION_PHOTO_SIZE: [number, number] = [747, 560];

async function safeQuery(sql: string, args: unknown[] = []): Promise<Row[]> {
  try {
    const [rows] = await pool.query(sql, args);
    return Array.isArray(rows) ? (rows as Row[]) : [];
  } catch (err) {
    console.error("[reenaTemplateExport] safeQuery failed:", { sql, err });
    return [];
  }
}

/**
 * Extracts the S3 object key from a public URL pointing at our bucket.
 * Examples:
 *   https://bucket.s3.region.amazonaws.com/key/path.jpg -> "key/path.jpg"
 *   https://cdn.domain.com/key/path.jpg                 -> "key/path.jpg" (when host matches NEXT_PUBLIC_S3_BUCKET_URL)
 * Returns null if the URL does not match either form.
 */
function extractS3Key(url: string): string | null {
  if (!url) return null;
  try {
    const u = new URL(url);
    const publicBase = (process.env.NEXT_PUBLIC_S3_BUCKET_URL || "").replace(/\/+$/, "");
    if (publicBase) {
      const baseUrl = new URL(publicBase);
      if (u.host === baseUrl.host) {
        const basePath = baseUrl.pathname.replace(/\/+$/, "");
        let p = u.pathname;
        if (basePath && p.startsWith(basePath)) p = p.slice(basePath.length);
        return p.replace(/^\/+/, "") || null;
      }
    }
    const bucket = process.env.AWS_S3_BUCKET_NAME || "";
    const region = process.env.AWS_S3_REGION || "";
    if (
      bucket &&
      (u.host === `${bucket}.s3.${region}.amazonaws.com` ||
        u.host === `${bucket}.s3.amazonaws.com` ||
        u.host === `s3.${region}.amazonaws.com` ||
        u.host === `s3.amazonaws.com`)
    ) {
      let p = u.pathname.replace(/^\/+/, "");
      if (u.host.startsWith("s3.")) {
        // s3.region.amazonaws.com/bucket/key form
        const prefix = `${bucket}/`;
        if (p.startsWith(prefix)) p = p.slice(prefix.length);
      }
      return p || null;
    }
  } catch {
    return null;
  }
  return null;
}

/**
 * Inspect the first few bytes of `buf` and infer an image MIME type from the
 * format magic-bytes. Returns null when the buffer is not a recognized image
 * (in which case we know the response was an error page or non-image asset).
 */
function detectImageMimeFromBytes(buf: Buffer): string | null {
  if (!buf || buf.length < 4) return null;
  // PNG  : 89 50 4E 47
  if (buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47) {
    return "image/png";
  }
  // JPEG : FF D8 FF
  if (buf[0] === 0xff && buf[1] === 0xd8 && buf[2] === 0xff) {
    return "image/jpeg";
  }
  // GIF  : 47 49 46 38
  if (buf[0] === 0x47 && buf[1] === 0x49 && buf[2] === 0x46 && buf[3] === 0x38) {
    return "image/gif";
  }
  // WEBP : "RIFF"...."WEBP"
  if (
    buf.length >= 12 &&
    buf[0] === 0x52 &&
    buf[1] === 0x49 &&
    buf[2] === 0x46 &&
    buf[3] === 0x46 &&
    buf[8] === 0x57 &&
    buf[9] === 0x45 &&
    buf[10] === 0x42 &&
    buf[11] === 0x50
  ) {
    return "image/webp";
  }
  // BMP  : 42 4D
  if (buf[0] === 0x42 && buf[1] === 0x4d) {
    return "image/bmp";
  }
  return null;
}

/**
 * Normalise a candidate value into a usable absolute http(s) URL.
 * Accepts:
 *   - "https://..." or "http://..."  -> returned as-is
 *   - "key/path.jpg"                 -> joined onto NEXT_PUBLIC_S3_BUCKET_URL
 * Returns null when the input cannot be turned into a fetchable URL.
 */
function normalizeS3Url(raw: string | null | undefined): string | null {
  const v = String(raw || "").trim();
  if (!v) return null;
  if (/^https?:\/\//i.test(v)) return v;
  const base = (process.env.NEXT_PUBLIC_S3_BUCKET_URL || "").replace(/\/+$/, "");
  if (!base) return null;
  return `${base}/${v.replace(/^\/+/, "")}`;
}

/**
 * Fetch a URL and validate that the body is an actual image. Strategy:
 *   1. direct fetch
 *   2. accept the body if Content-Type starts with "image/" OR (the
 *      Content-Type is missing/octet-stream AND the buffer's magic bytes
 *      are a known image format)
 *   3. on 4xx/5xx OR non-image body, retry once with a presigned S3 URL
 *      derived from the original URL's object key (handles private buckets).
 *
 * Never throws - all failures return null and are logged.
 */
async function fetchImageBuffer(
  url: string,
  label: string
): Promise<{ buffer: Buffer; contentType: string; url: string } | null> {
  const normalized = normalizeS3Url(url);
  if (!normalized) {
    console.warn(`[image fetch] ${label} could not normalise url:`, url);
    return null;
  }

  const tryFetch = async (target: string, attempt: string) => {
    try {
      const res = await fetch(target);
      const contentType = res.headers.get("content-type") || "";
      const contentLength = res.headers.get("content-length") || "";
      console.log(`[image fetch] ${label} ${attempt}`, {
        url: target,
        ok: res.ok,
        status: res.status,
        contentType,
        contentLength,
      });
      console.log(`[DOCX image fetch]`, {
        label,
        attempt,
        url: target,
        ok: res.ok,
        status: res.status,
        contentType,
        contentLength,
      });
      if (!res.ok) return null;
      const ab = await res.arrayBuffer();
      const buf = Buffer.from(ab);
      if (!buf.length) {
        console.warn(`[image fetch] ${label} ${attempt} empty body`);
        return null;
      }

      // Reject obvious HTML/text error pages outright.
      if (contentType.startsWith("text/") || contentType.includes("html")) {
        console.warn(
          `[image fetch] ${label} ${attempt} non-image content-type:`,
          contentType
        );
        return null;
      }

      // Strict check: server explicitly claims an image.
      if (contentType.startsWith("image/")) {
        console.log(`[DOCX image buffer]`, {
          label,
          attempt,
          bufferSize: buf.length,
          contentType,
        });
        return { buffer: buf, contentType, url: target };
      }

      // Loose check: server returned octet-stream or missing header. Sniff
      // the magic bytes - this is the case where S3 stored the photo without
      // a Content-Type set.
      const sniffed = detectImageMimeFromBytes(buf);
      if (sniffed) {
        console.log(
          `[image fetch] ${label} ${attempt} sniffed image type from bytes:`,
          sniffed
        );
        console.log(`[DOCX image buffer]`, {
          label,
          attempt,
          bufferSize: buf.length,
          contentType: sniffed,
          sniffed: true,
        });
        return { buffer: buf, contentType: sniffed, url: target };
      }

      console.warn(
        `[image fetch] ${label} ${attempt} not image (header="${contentType}" and bytes don't match any image format)`
      );
      return null;
    } catch (err) {
      console.error(`[image fetch] ${label} ${attempt} threw`, target, err);
      return null;
    }
  };

  // Attempt 1: direct fetch.
  const direct = await tryFetch(normalized, "direct");
  if (direct) return direct;

  // Attempt 2: signed URL fallback for private buckets.
  const key = extractS3Key(normalized);
  if (key) {
    try {
      const signed = await getReadSignedUrl(key, 600);
      const second = await tryFetch(signed, "signed");
      if (second) return second;
    } catch (err) {
      console.error(`[image fetch] ${label} signed-url generation failed`, err);
    }
  } else {
    console.warn(`[image fetch] ${label} could not extract S3 key from`, normalized);
  }

  return null;
}

function stripHtml(input: string) {
  return String(input || "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|li|h[1-6])>/gi, "\n")
    .replace(/<li>/gi, "- ")
    .replace(/<[^>]*>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function formatDateDDMMYYYYDot(d = new Date()) {
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${dd}.${mm}.${d.getFullYear()}`;
}

function formatDateDDMMYYYYDash(d = new Date()) {
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${dd}-${mm}-${d.getFullYear()}`;
}

function formatLatLine(prefix: "N" | "S" | "E" | "W", v: unknown): string {
  const n = Number(v);
  if (!Number.isFinite(n)) return "-";
  const abs = Math.abs(n);
  const deg = Math.floor(abs);
  const minutes = ((abs - deg) * 60).toFixed(3);
  return `${prefix}${deg} ${minutes}`;
}

function formatGpsLat(lat: unknown): string {
  const n = Number(lat);
  if (!Number.isFinite(n)) return "-";
  return formatLatLine(n >= 0 ? "N" : "S", lat);
}

function formatGpsLon(lng: unknown): string {
  const n = Number(lng);
  if (!Number.isFinite(n)) return "-";
  return formatLatLine(n >= 0 ? "E" : "W", lng);
}

function valueOrDash(v: unknown): string {
  if (v === null || typeof v === "undefined") return "-";
  const s = String(v).trim();
  return s ? s : "-";
}

function valueOrEmDash(v: unknown): string {
  if (v === null || typeof v === "undefined") return "—";
  const s = String(v).trim();
  return s ? s : "—";
}

function cleanFileName(name: string, fallback: string) {
  const s = String(name || "")
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/[. ]+$/g, "");
  return (s || fallback).slice(0, 120);
}

async function readTemplate(): Promise<Buffer> {
  return fs.readFile(TEMPLATE_PATH);
}

type ImageEntry = { buffer: Buffer; contentType: string; url?: string; path?: string };

type ObservationData = {
  gpsLat: string;
  gpsLon: string;
  km: string;
  location: string;
  category: string;
  observation: string;
  remarks: string;
  photo: string;
  photoKey: string;
  hasObservationPhoto: boolean;
  photoFallback: string;
  categoryIconKey: string;
  hasCategoryIcon: boolean;
  // Per-row difficulty styling — feeds either rawXml placeholders for cell
  // shading or the {#isRedDifficulty}/{#isYellowDifficulty}/{#isGreenDifficulty}
  // conditional blocks that triplicate the observation table in the template.
  difficultyValue: string;
  difficultyKey: "red" | "yellow" | "green";
  headerFillColor: string;
  headerTextColor: string;
  bodyFillColor: string;
  bodyTextColor: string;
  isRedDifficulty: boolean;
  isYellowDifficulty: boolean;
  isGreenDifficulty: boolean;
};

function normalizeDifficulty(value: unknown) {
  return String(value || "")
    .trim()
    .toLowerCase();
}

function getDifficultyTableColors(value: unknown): {
  key: "red" | "yellow" | "green";
  headerFillColor: string;
  headerTextColor: string;
  bodyFillColor: string;
  bodyTextColor: string;
} {
  const d = normalizeDifficulty(value);
  if (d.includes("red")) {
    return {
      key: "red",
      headerFillColor: "C00000",
      headerTextColor: "FFFFFF",
      bodyFillColor: "F4CCCC",
      bodyTextColor: "0B3D2E",
    };
  }
  if (d.includes("yellow")) {
    return {
      key: "yellow",
      headerFillColor: "FFFF00",
      headerTextColor: "000000",
      bodyFillColor: "FFF2CC",
      bodyTextColor: "0B3D2E",
    };
  }
  return {
    key: "green",
    headerFillColor: "4CAF50",
    headerTextColor: "FFFFFF",
    bodyFillColor: "D9EAD3",
    bodyTextColor: "0B3D2E",
  };
}

const CATEGORY_ICONS_DIR = path.join(process.cwd(), "public", "images", "report-icons");

function normalizeCategory(value: unknown) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function getCategoryDisplayName(category: unknown): string {
  const c = normalizeCategory(category);
  if (!c) return String(category || "-");
  if (c.includes("footpath bridge")) return "Footpath Bridge";
  if (c.includes("low tension")) return "Low Tension Cable";
  if (c.includes("high tension")) return "High Tension Cable";
  if (c.includes("tower")) return "Tower Line Cable";
  if (c.includes("underpass")) return "Underpass Bridge";
  if (c.includes("tree")) return "Tree Branches";
  if (c.includes("river bridge")) return "River Bridge";
  if (
    c.includes("signal pole") ||
    c.includes("speed pole") ||
    c.includes("side signboard")
  ) {
    return "Side Signboard / Signal Pole / Speed Pole";
  }
  if (
    c.includes("signboard") ||
    c.includes("camera pole") ||
    c.includes("electric sign")
  ) {
    return "Signboard / Electric Signboard / Camera Pole";
  }
  if (c.includes("toll")) return "Toll Plaza";
  if (c.includes("narrow")) return "Narrow Road";
  if (c.includes("gate")) return "Gate";
  if (c.includes("bend")) return "Bend";
  if (c.includes("petrol")) return "Petrol Bunk";
  if (c.includes("railway")) return "Railway Level Crossing";
  return String(category || "-");
}

// Underscore-key normaliser: turns "Footpath Bridge" → "footpath_bridge",
// "Side Signboard / Signal Pole" → "side_signboard_signal_pole", etc.
// Then runs alias rules so "LT Cable" / "Low Tension Cable" / "low_tension"
// all collapse to a single canonical key the CATEGORY_ICON_MAP knows.
function normalizeCategoryKey(value: unknown): string {
  const c = String(value || "")
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
  // Side variants must be checked BEFORE the broader signboard / camera_pole
  // rules so "Side Signboard" doesn't resolve to "signboard" (ca-9).
  if (
    c.includes("side_signboard") ||
    c.includes("signal_pole") ||
    c.includes("speed_pole") ||
    c.includes("electric_side_signboard")
  ) {
    return "side_signboard";
  }
  if (c.includes("footpath_bridge")) return "footpath_bridge";
  if (c.includes("low_tension") || c === "lt_cable") return "lt_cable";
  if (c.includes("high_tension") || c === "ht_cable") return "ht_cable";
  if (c.includes("tower")) return "towerline_cable";
  if (c.includes("underpass")) return "underpass";
  if (c.includes("tree")) return "tree";
  if (c.includes("river")) return "river_bridge";
  if (
    c.includes("signboard") ||
    c.includes("electric_sign") ||
    c.includes("camera_pole")
  ) {
    return "signboard";
  }
  if (c.includes("toll")) return "toll";
  if (c.includes("narrow")) return "narrow_road";
  if (c.includes("gate")) return "gate";
  if (c.includes("bend")) return "bend";
  if (c.includes("petrol")) return "petrol";
  if (c.includes("railway")) return "railway_level_crossing";
  if (c.includes("diversion")) return "diversion";
  return "fallback";
}

// Server-side fs paths (process.cwd() is the project root). The browser-side
// equivalent lives in lib/download.ts and uses URL-style paths instead.
const CATEGORY_ICON_MAP: Record<string, string> = {
  footpath_bridge: "public/images/report-icons/image.png",
  lt_cable: "public/images/report-icons/ca-3.png",
  low_tension_cable: "public/images/report-icons/ca-3.png",
  ht_cable: "public/images/report-icons/ca-4.png",
  high_tension_cable: "public/images/report-icons/ca-4.png",
  towerline_cable: "public/images/report-icons/ca-5.png",
  towerline: "public/images/report-icons/ca-5.png",
  tower_line: "public/images/report-icons/ca-5.png",
  tower_line_cable: "public/images/report-icons/ca-5.png",
  underpass: "public/images/report-icons/ca-6.png",
  underpass_bridge: "public/images/report-icons/ca-6.png",
  tree: "public/images/report-icons/ca-7.png",
  tree_branches: "public/images/report-icons/ca-7.png",
  river_bridge: "public/images/report-icons/ca-8.png",
  signboard: "public/images/report-icons/ca-9.png",
  electric_sign: "public/images/report-icons/ca-9.png",
  electric_signboard: "public/images/report-icons/ca-9.png",
  camera_pole: "public/images/report-icons/ca-9.png",
  toll: "public/images/report-icons/ca-10.png",
  toll_plaza: "public/images/report-icons/ca-10.png",
  narrow_road: "public/images/report-icons/ca-11.png",
  gate: "public/images/report-icons/ca-12.png",
  side_signboard: "public/images/report-icons/ca-13.png",
  signal_pole: "public/images/report-icons/ca-13.png",
  speed_pole: "public/images/report-icons/ca-13.png",
  electric_side_signboard: "public/images/report-icons/ca-13.png",
  bend: "public/images/report-icons/ca-14.png",
  petrol: "public/images/report-icons/ca-15.png",
  petrol_bunk: "public/images/report-icons/ca-15.png",
  railway_level_crossing: "public/images/report-icons/ca-16.png",
  diversion: "public/images/report-icons/diversion.jpeg",
  fallback: "public/images/report-icons/image.png",
};

function getCategoryIconFile(category: unknown): string {
  const key = normalizeCategoryKey(category);
  const rel = CATEGORY_ICON_MAP[key] || CATEGORY_ICON_MAP.fallback;
  console.log("[category mapping check]", {
    rawCategory: category,
    normalizedKey: key,
    iconRelativePath: rel,
  });
  // Strip the public/images/report-icons/ prefix so the existing
  // CATEGORY_ICONS_DIR-based loader (path.join(CATEGORY_ICONS_DIR, fileName))
  // continues to point at the right file.
  return rel.replace(/^public[\\/]images[\\/]report-icons[\\/]/, "");
}

const categoryIconCache = new Map<string, ImageEntry | null>();

async function loadCategoryIcon(category: unknown): Promise<ImageEntry | null> {
  const fileName = getCategoryIconFile(category);
  const fullPath = path.join(CATEGORY_ICONS_DIR, fileName);
  if (categoryIconCache.has(fileName)) {
    const cached = categoryIconCache.get(fileName) || null;
    console.log("[category icon debug]", {
      category,
      categoryIconPath: fullPath,
      cached: true,
      hasCategoryIcon: !!cached?.buffer,
      iconBufferSize: cached?.buffer?.length || 0,
    });
    return cached;
  }
  try {
    const buffer = await fs.readFile(fullPath);
    if (!buffer.length) {
      console.log("[category icon debug]", {
        category,
        categoryIconPath: fullPath,
        exists: true,
        empty: true,
        hasCategoryIcon: false,
        iconBufferSize: 0,
      });
      categoryIconCache.set(fileName, null);
      return null;
    }
    const contentType = fileName.toLowerCase().endsWith(".jpeg") || fileName.toLowerCase().endsWith(".jpg")
      ? "image/jpeg"
      : "image/png";
    const entry: ImageEntry = { buffer, contentType, path: fullPath };
    categoryIconCache.set(fileName, entry);
    console.log("[category icon debug]", {
      category,
      categoryIconPath: fullPath,
      exists: true,
      hasCategoryIcon: true,
      iconBufferSize: buffer.length,
      contentType,
    });
    return entry;
  } catch (err) {
    console.warn("[category icon debug]", {
      category,
      categoryIconPath: fullPath,
      exists: false,
      hasCategoryIcon: false,
      iconBufferSize: 0,
      error: (err as { code?: string })?.code || String(err),
    });
    categoryIconCache.set(fileName, null);
    return null;
  }
}

/**
 * If the row stores a key but no public URL, build the URL from the key.
 */
function resolveImageUrl(row: Row | null | undefined, urlCol: string, keyCol: string) {
  if (!row) return "";
  const url = String(row[urlCol] || "").trim();
  if (url) return url;
  const key = String(row[keyCol] || "").trim();
  if (!key) return "";
  const base = (process.env.NEXT_PUBLIC_S3_BUCKET_URL || "").replace(/\/+$/, "");
  if (!base) return "";
  return `${base}/${key.replace(/^\/+/, "")}`;
}

/**
 * Pull the first defined coordinate from a report row.
 * `reports` table has both `latitude/longitude` and `loc_lat/loc_lon` columns.
 */
function pickLat(r: Row): number | null {
  const candidates = [r.latitude, r.lat, r.ne_latitude, r.ne_lat, r.loc_lat];
  for (const c of candidates) {
    const n = Number(c);
    if (Number.isFinite(n)) return n;
  }
  return null;
}
function pickLng(r: Row): number | null {
  const candidates = [r.longitude, r.lng, r.ne_longitude, r.ne_lng, r.lon, r.loc_lon];
  for (const c of candidates) {
    const n = Number(c);
    if (Number.isFinite(n)) return n;
  }
  return null;
}

/**
 * Great-circle distance between two lat/lng points, in kilometres.
 */
function haversineKm(lat1: number, lon1: number, lat2: number, lon2: number): number {
  const R = 6371;
  const toRad = (v: number) => (v * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

/**
 * Annotate each report (already sorted by sort_order ASC, created_at ASC) with
 * a `calculated_km` field equal to the cumulative haversine distance walked
 * along the ordered route. The first report is always 0 km.
 */
function calculateCumulativeKms(reports: Row[]): Row[] {
  let total = 0;
  return reports.map((r, idx) => {
    if (idx === 0) {
      return { ...r, calculated_km: 0 };
    }
    const prev = reports[idx - 1];
    const lat = pickLat(r);
    const lng = pickLng(r);
    const prevLat = pickLat(prev);
    const prevLng = pickLng(prev);
    if (
      lat !== null &&
      lng !== null &&
      prevLat !== null &&
      prevLng !== null
    ) {
      total += haversineKm(prevLat, prevLng, lat, lng);
    }
    return { ...r, calculated_km: total };
  });
}

/**
 * Format a numeric km value to 4 decimal places, or "-" if not finite.
 */
function formatKmValue(value: unknown): string {
  const n = Number(value);
  if (!Number.isFinite(n)) return "-";
  return n.toFixed(4);
}

/**
 * Ensure reports.resolved_location exists so we can cache reverse-geocoded
 * place names without a separate migration step.
 */
async function ensureResolvedLocationColumn() {
  try {
    const [cols] = await pool.query("SHOW COLUMNS FROM reports LIKE 'resolved_location'");
    if (Array.isArray(cols) && cols.length === 0) {
      await pool.query("ALTER TABLE reports ADD COLUMN resolved_location TEXT NULL");
      console.log("[export location] added reports.resolved_location column");
    }
  } catch (err) {
    console.error("[export location] ensureResolvedLocationColumn failed:", err);
  }
}

/**
 * Pick the first non-empty stored location string from a report row.
 */
function getExistingLocation(r: Row): string | null {
  const candidates = [
    r.location,
    r.address,
    r.resolved_location,
    r.location_name,
    r.location_text,
    r.place,
    r.landmark,
    r.loc_text,
  ];
  for (const c of candidates) {
    if (c === null || typeof c === "undefined") continue;
    const s = String(c).trim();
    if (s) return s;
  }
  return null;
}

/**
 * Reverse-geocode lat/lng to a human-readable place name. Uses OpenStreetMap
 * Nominatim. Fails silently and returns null on any error/timeout. Callers
 * MUST cache the result in MySQL so we don't hammer the public service.
 */
async function reverseGeocodeLocation(lat: number, lng: number): Promise<string | null> {
  try {
    const url =
      `https://nominatim.openstreetmap.org/reverse?format=jsonv2` +
      `&lat=${encodeURIComponent(String(lat))}&lon=${encodeURIComponent(String(lng))}&zoom=14`;
    const ac = new AbortController();
    const timer = setTimeout(() => ac.abort(), 5000);
    let res: Response;
    try {
      res = await fetch(url, {
        headers: {
          "User-Agent": "lbi-web-export/1.0 (race route export)",
          Accept: "application/json",
        },
        signal: ac.signal,
      });
    } finally {
      clearTimeout(timer);
    }
    if (!res.ok) {
      console.warn("[export location] nominatim non-200:", res.status);
      return null;
    }
    const json = (await res.json()) as { display_name?: unknown };
    const name = typeof json?.display_name === "string" ? json.display_name.trim() : "";
    return name || null;
  } catch (err) {
    console.warn("[export location] reverseGeocodeLocation failed:", err);
    return null;
  }
}

/**
 * Resolve the LOCATION cell value for one report. Order:
 *   1. existing stored value (location, address, resolved_location, ...)
 *   2. reverse-geocoded name from coordinates (cached back into MySQL)
 *   3. coordinate fallback "lat, lng"
 *   4. "-"
 */
async function resolveReportLocation(r: Row, hasResolvedColumn: boolean): Promise<string> {
  const existing = getExistingLocation(r);
  if (existing) return existing;
  const lat = pickLat(r);
  const lng = pickLng(r);
  if (lat === null || lng === null) return "-";
  const resolved = await reverseGeocodeLocation(lat, lng);
  if (resolved) {
    if (hasResolvedColumn && r.id) {
      try {
        await pool.execute("UPDATE reports SET resolved_location = ? WHERE id = ?", [
          resolved,
          r.id,
        ]);
      } catch (err) {
        console.warn("[export location] cache update failed:", err);
      }
    }
    return resolved;
  }
  // Coordinate fallback so the cell never says "-" when a coordinate exists.
  return `${lat.toFixed(6)}, ${lng.toFixed(6)}`;
}

async function ensureGaDrawingsTable() {
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS project_ga_drawings (
        id VARCHAR(36) PRIMARY KEY,
        project_id VARCHAR(36) NOT NULL UNIQUE,
        image_url TEXT NULL,
        image_key TEXT NULL,
        file_name VARCHAR(255) NULL,
        conclusion_html LONGTEXT NULL,
        created_by VARCHAR(36) NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        INDEX idx_project_ga_drawings_project_id (project_id)
      )
    `);
  } catch (err) {
    console.error("[reenaTemplateExport] ensureGaDrawingsTable failed:", err);
  }
}

/**
 * Wrap any thrown value in a labelled error so the caller can tell the user
 * exactly which export step crashed.
 */
class ExportStepError extends Error {
  step: string;
  cause?: unknown;
  detail?: string;
  constructor(step: string, message: string, cause?: unknown, detail?: string) {
    super(message);
    this.name = "ExportStepError";
    this.step = step;
    this.cause = cause;
    this.detail = detail;
  }
}

export async function generateReenaDocx(options: ExportOptions): Promise<ExportResult> {
  const projectId = options.projectId;
  const includePhotos = options.includePhotos !== false;
  const reportIdsFilter = (options.reportIds || []).filter(Boolean);

  console.log("[export actual] started", {
    projectId,
    reportIdsCount: reportIdsFilter.length,
    includePhotos,
    BUILD_FINGERPRINT: "obs-photo-fix-2026-04-27-A",
  });

  // ----- Step 1: Template existence sanity check.
  let templateExists = false;
  try {
    const stat = await fs.stat(TEMPLATE_PATH);
    templateExists = stat.isFile();
  } catch (err) {
    console.error("[export actual] template stat failed:", err);
  }
  console.log("[export actual] template exists:", templateExists, "path:", TEMPLATE_PATH);
  if (!templateExists) {
    throw new ExportStepError(
      "template_load",
      `DOCX template not found at ${TEMPLATE_PATH}`
    );
  }

  // ----- Step 2: Project.
  let projectRow: Row | null = null;
  try {
    const rows = await safeQuery("SELECT * FROM projects WHERE id = ? LIMIT 1", [projectId]);
    projectRow = rows[0] || null;
  } catch (err) {
    throw new ExportStepError("data_prepare", "Project query failed", err);
  }
  if (!projectRow) throw new Error("Project not found");
  console.log("[export actual] project found:", !!projectRow);

  const projectName =
    String(projectRow.name || projectRow.title || projectRow.project_name || "PROJECT").trim() ||
    "PROJECT";

  // ----- Step 3: Reports.
  let reports: Row[] = [];
  try {
    const reportSql =
      reportIdsFilter.length > 0
        ? `SELECT * FROM reports WHERE project_id = ? AND id IN (${reportIdsFilter
            .map(() => "?")
            .join(",")}) ORDER BY sort_order ASC, created_at ASC`
        : "SELECT * FROM reports WHERE project_id = ? ORDER BY sort_order ASC, created_at ASC";
    const reportArgs = reportIdsFilter.length ? [projectId, ...reportIdsFilter] : [projectId];
    reports = await safeQuery(reportSql, reportArgs);
  } catch (err) {
    throw new ExportStepError("data_prepare", "Reports query failed", err);
  }
  const reportIds = reports.map((r) => String(r.id || "").trim()).filter(Boolean);

  console.log("[export actual] selected reports:", reports.length);
  console.log("[export actual] report ids:", reportIds);

  // ----- Step 3b: cumulative KM + resolved-location column setup.
  reports = calculateCumulativeKms(reports);
  await ensureResolvedLocationColumn();
  let hasResolvedLocationColumn = false;
  try {
    const [cols] = await pool.query("SHOW COLUMNS FROM reports LIKE 'resolved_location'");
    hasResolvedLocationColumn = Array.isArray(cols) && cols.length > 0;
  } catch {
    hasResolvedLocationColumn = false;
  }

  // ----- Step 4: Photos.
  let photos: Row[] = [];
  let photosQueryError: unknown = null;
  console.error("[PHOTOS_DEBUG_2026_04_27_B] entering photos query block", {
    reportIdsLength: reportIds.length,
    includePhotos,
    sample: reportIds.slice(0, 3),
  });
  if (includePhotos && reportIds.length > 0) {
    try {
      const placeholders = reportIds.map(() => "?").join(",");
      const sql = `SELECT id, report_id, url, file_name, created_at
         FROM report_photos
         WHERE report_id IN (${placeholders})
         ORDER BY created_at ASC`;
      console.error("[PHOTOS_DEBUG_2026_04_27_B] sql", { sql, args: reportIds });
      const [photoRows] = await pool.execute(sql, reportIds);
      photos = Array.isArray(photoRows) ? (photoRows as Row[]) : [];
      console.error("[PHOTOS_DEBUG_2026_04_27_B] query returned rows:", photos.length);
    } catch (err) {
      photosQueryError = err;
      console.error("[PHOTOS_DEBUG_2026_04_27_B] photos query THREW:", err);
      photos = [];
    }
  } else {
    console.error("[PHOTOS_DEBUG_2026_04_27_B] skipping photos query", {
      includePhotos,
      reportIdsLength: reportIds.length,
    });
  }

  console.error("[PHOTOS_DEBUG_2026_04_27_B] report photo query result", {
    includePhotos,
    reportIdsForPhotos: reportIds,
    photosCount: photos.length,
    queryFailed: !!photosQueryError,
    photos: photos.map((p) => ({
      id: p.id,
      report_id: p.report_id,
      url: p.url,
    })),
  });

  console.log("[export observation photos query]", {
    reportIds,
    photosCount: photos.length,
    queryFailed: !!photosQueryError,
    photos: photos.map((p) => ({
      id: p.id,
      report_id: p.report_id,
      url: p.url,
    })),
  });

  console.log("[export actual] photos:", photos.length);
  console.log(
    "[export actual] photos sample:",
    photos.slice(0, 3).map((p) => ({
      report_id: p.report_id,
      url: p.url,
      file_name: p.file_name,
    }))
  );

  // Spec-mandated diagnostic dumps so we can confirm the join is correct.
  console.log("[export photos] reports:", reports.map((r) => r.id));
  console.log(
    "[export photos] photos:",
    photos.map((p) => ({ id: p.id, report_id: p.report_id, url: p.url }))
  );

  const photosByReportId = new Map<string, Row[]>();
  for (const p of photos) {
    const key = String(p.report_id || "").trim();
    if (!key) continue;
    if (!photosByReportId.has(key)) photosByReportId.set(key, []);
    photosByReportId.get(key)!.push(p);
  }

  // ----- Step 5: GA drawing.
  await ensureGaDrawingsTable();
  let gaRow: Row | null = null;
  try {
    const rows = await safeQuery(
      "SELECT * FROM project_ga_drawings WHERE project_id = ? LIMIT 1",
      [projectId]
    );
    gaRow = rows[0] || null;
  } catch (err) {
    console.error("[export actual] ga query failed - continuing without:", err);
  }
  const gaImageUrl = resolveImageUrl(gaRow, "image_url", "image_key");
  console.log("[export actual] gaDrawingUrl:", gaImageUrl || "(none)");

  // ----- Step 6: Route page (objective + route map).
  let routePageRow: Row | null = null;
  try {
    const rows = await safeQuery(
      "SELECT * FROM project_route_pages WHERE project_id = ? ORDER BY created_at DESC LIMIT 1",
      [projectId]
    );
    routePageRow = rows[0] || null;
  } catch (err) {
    console.error("[export actual] route page query failed - continuing without:", err);
  }
  const routeMapUrl = String(routePageRow?.map_file_url || "").trim();
  console.log("[export actual] routeMapUrl:", routeMapUrl || "(none)");

  // ----- Resolve scalar template values.
  const objective = valueOrDash(
    routePageRow?.objective ?? projectRow.objective ?? projectRow.description
  );
  const conclusionSource = gaRow?.conclusion_html || routePageRow?.conclusion_html || "";
  const conclusion = stripHtml(String(conclusionSource)) || "-";
  const dateDot = formatDateDDMMYYYYDot();
  const dateDash = formatDateDDMMYYYYDash();

  // ----- Step 7: Fetch image buffers. Each fetch must NEVER throw - it
  // returns null on any failure so the export can continue with fallbacks.
  // Reuses the module-level ImageEntry type so category-icon entries (loaded
  // from disk, no url) and S3-fetched entries (no path) share a single shape.
  const imageMap = new Map<string, ImageEntry>();
  try {
    const gaDrawingFetched = gaImageUrl
      ? await fetchImageBuffer(gaImageUrl, "gaDrawing")
      : null;
    if (gaDrawingFetched && Buffer.isBuffer(gaDrawingFetched.buffer)) {
      imageMap.set("gaDrawing", gaDrawingFetched);
    }
  } catch (err) {
    console.error("[export actual] gaDrawing fetch step threw - ignoring:", err);
  }
  try {
    const routeMapFetched = routeMapUrl ? await fetchImageBuffer(routeMapUrl, "routeMap") : null;
    if (routeMapFetched && Buffer.isBuffer(routeMapFetched.buffer)) {
      imageMap.set("routeMap", routeMapFetched);
    }
  } catch (err) {
    console.error("[export actual] routeMap fetch step threw - ignoring:", err);
  }

  // ----- Observations.
  console.log("[includePhotos check]", { includePhotos });
  console.log("[photosByReportId keys]", Array.from(photosByReportId.keys()));

  const observations: ObservationData[] = [];
  for (let i = 0; i < reports.length; i += 1) {
    const r = reports[i];
    const rid = String(r.id || "").trim();
    const lat = pickLat(r);
    const lng = pickLng(r);
    const hasLat = lat !== null;
    const hasLon = lng !== null;

    let photoKey = "";
    if (!includePhotos) {
      console.log("[export observation photo prepared]", {
        index: i,
        reportId: rid,
        skipped: "includePhotos=false",
      });
    } else {
      const reportPhotos = photosByReportId.get(rid) || [];
      const firstPhotoUrl = normalizeS3Url(reportPhotos[0]?.url) || null;

      console.error("[PHOTOS_DEBUG_2026_04_27_B] photo mapping", {
        reportId: rid,
        reportPhotosCount: reportPhotos.length,
        firstPhotoUrl,
        photoUrls: reportPhotos.map((p: Row) => ({ url: p.url })),
      });

      console.log("[DOCX photo URL]", {
        reportId: rid,
        firstPhotoUrl,
      });

      // Iterate every saved photo for this report; first one that fetches OK
      // wins. This handles the case where the first row was orphaned/deleted
      // in S3 but a later row is still valid.
      for (const p of reportPhotos) {
        const candidate = normalizeS3Url(p?.url);
        if (!candidate) continue;
        try {
          const fetched = await fetchImageBuffer(candidate, `photo[${rid}]`);
          if (fetched && Buffer.isBuffer(fetched.buffer) && fetched.buffer.length > 0) {
            photoKey = `obsPhoto_${i}`;
            imageMap.set(photoKey, fetched);
            break;
          }
        } catch (err) {
          console.error(`[export actual] photo fetch threw for ${candidate} - skipping:`, err);
        }
      }

      console.log("[export observation photo prepared]", {
        index: i,
        reportId: rid,
        reportPhotosCount: reportPhotos.length,
        firstPhotoUrl,
        photoKey,
        hasObservationPhoto: !!photoKey,
        imageMapHasKey: photoKey ? imageMap.has(photoKey) : false,
      });

      console.log("[observation image prepared]", {
        reportId: rid,
        hasPhotoUrl: !!firstPhotoUrl,
        hasImageBuffer: !!photoKey && imageMap.has(photoKey),
      });
    }

    // KM is the cumulative haversine distance along the ordered route. The
    // database `kms`/`km` columns are NOT used because the user spec says KM
    // must be derived from coordinates so the first row is always 0.0000.
    const kmText = formatKmValue(r.calculated_km);

    // LOCATION uses (1) any stored value, then (2) reverse-geocoded name
    // (cached into reports.resolved_location), then (3) coordinate fallback.
    const locationText = await resolveReportLocation(r, hasResolvedLocationColumn);

    const hasObservationPhoto = !!photoKey && imageMap.has(photoKey);
    const entry = photoKey ? imageMap.get(photoKey) : undefined;

    console.error("[PHOTOS_DEBUG_2026_04_27_B] observation photo prepared", {
      index: i,
      reportId: rid,
      photoKey,
      hasObservationPhoto,
      imageMapHasKey: !!photoKey && imageMap.has(photoKey),
      bufferSize: entry?.buffer?.length || 0,
    });

    const categoryIcon = await loadCategoryIcon(r.category);
    let categoryIconKey = "";
    let hasCategoryIcon = false;
    if (categoryIcon?.buffer && categoryIcon.buffer.length > 0) {
      categoryIconKey = `catIcon_${i}`;
      imageMap.set(categoryIconKey, categoryIcon);
      hasCategoryIcon = true;
    }

    console.log("[category icon prepared]", {
      index: i,
      category: r.category,
      categoryIconPath: categoryIcon?.path || null,
      categoryIconKey,
      hasCategoryIcon,
      imageMapHasKey: categoryIconKey ? imageMap.has(categoryIconKey) : false,
      bufferSize: categoryIconKey
        ? imageMap.get(categoryIconKey)?.buffer?.length || 0
        : 0,
    });

    const difficultyValue = String(
      r.difficulty ||
        r.remarks_action ||
        r.status ||
        r.vehicle_movement ||
        r.movement ||
        ""
    );
    const tableColors = getDifficultyTableColors(difficultyValue);

    observations.push({
      gpsLat: hasLat ? formatGpsLat(lat) : "-",
      gpsLon: hasLon ? formatGpsLon(lng) : "-",
      km: kmText,
      location: locationText || "-",
      category: valueOrDash(r.category),
      observation: valueOrEmDash(r.description ?? r.observation),
      remarks: valueOrEmDash(r.remarks_action ?? r.difficulty ?? r.status),
      photo: photoKey,
      photoKey,
      hasObservationPhoto,
      photoFallback: hasObservationPhoto ? "" : "Photo not available.",
      categoryIconKey,
      hasCategoryIcon,
      difficultyValue,
      difficultyKey: tableColors.key,
      headerFillColor: tableColors.headerFillColor,
      headerTextColor: tableColors.headerTextColor,
      bodyFillColor: tableColors.bodyFillColor,
      bodyTextColor: tableColors.bodyTextColor,
      isRedDifficulty: tableColors.key === "red",
      isYellowDifficulty: tableColors.key === "yellow",
      isGreenDifficulty: tableColors.key === "green",
    });
  }

  console.log(
    "[difficulty table color]",
    observations.map((o, idx) => ({
      index: idx,
      difficultyValue: o.difficultyValue,
      difficultyKey: o.difficultyKey,
      bodyFillColor: o.bodyFillColor,
    }))
  );

  // ----- Category count summary. Group reports by display-name and attach
  // the same icon used in the per-observation CATEGORY cell. Reuses the
  // loadCategoryIcon cache so we don't re-read PNGs from disk.
  const categoryCountMap = new Map<string, number>();
  const categorySampleByLabel = new Map<string, unknown>();
  for (const r of reports) {
    const label = getCategoryDisplayName(r.category);
    categoryCountMap.set(label, (categoryCountMap.get(label) || 0) + 1);
    if (!categorySampleByLabel.has(label)) {
      categorySampleByLabel.set(label, r.category);
    }
  }
  const categorySummary = await Promise.all(
    Array.from(categoryCountMap.entries()).map(async ([label, count]) => {
      const sample = categorySampleByLabel.get(label);
      const icon = await loadCategoryIcon(sample);
      const hasIcon = !!icon?.buffer && icon.buffer.length > 0;
      return {
        categoryLabel: label,
        count,
        countText: String(count),
        categoryIcon: icon,
        hasCategoryIcon: hasIcon,
      };
    })
  );
  console.log(
    "[category summary]",
    categorySummary.map((c) => ({
      categoryLabel: c.categoryLabel,
      count: c.count,
      hasCategoryIcon: c.hasCategoryIcon,
      iconBufferSize: c.categoryIcon?.buffer?.length || 0,
    }))
  );

  console.log(
    "[export km/location]",
    reports.map((r) => ({
      id: r.id,
      lat: pickLat(r),
      lng: pickLng(r),
      calculated_km: r.calculated_km,
      location: r.resolved_location || r.location || r.address || null,
    }))
  );

  console.log("[export actual] image map keys:", Array.from(imageMap.keys()));
  console.log("[export actual] observations built:", observations.length);

  // ----- Step 8: Load + parse template.
  let templateBuf: Buffer;
  try {
    templateBuf = await readTemplate();
  } catch (err) {
    console.error("[export actual] template load failed:", err);
    throw new ExportStepError("template_load", `Failed to read DOCX template at ${TEMPLATE_PATH}`, err);
  }

  let zip: PizZip;
  try {
    zip = new PizZip(templateBuf);
  } catch (err) {
    console.error("[export actual] template parse (pizzip) failed:", err);
    throw new ExportStepError("template_load", "DOCX template is corrupt or not a valid ZIP", err);
  }

  // ----- Step 9: Build image module + docxtemplater. getImage MUST always
  // return a Buffer (the image module crashes inside getRenderedPart if it
  // ever receives null/undefined). For the "no real image" case we fall back
  // to a 1x1 transparent PNG; the visible fallback text is handled by sibling
  // {photoFallback} placeholders in the template.
  const TRANSPARENT_PIXEL_PNG = Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
    "base64"
  );
  const imageModule = new ImageModule({
    centered: true,
    fileType: "docx",
    getImage: (tagValue: unknown, tagName: string) => {
      try {
        // Two supported tagValue shapes:
        //   - string key: look up imageMap (gaDrawing, routeMap, obsPhoto_*).
        //   - object {buffer}: pass through directly (categoryIcon).
        let buf: Buffer | undefined;
        let key = "";
        if (typeof tagValue === "string") {
          key = tagValue;
          const entry = key ? imageMap.get(key) : undefined;
          buf = entry?.buffer;
        } else if (tagValue && typeof tagValue === "object") {
          const maybe = (tagValue as { buffer?: unknown }).buffer;
          if (Buffer.isBuffer(maybe)) buf = maybe;
        }
        const hasBuffer = !!buf && Buffer.isBuffer(buf) && buf.length > 0;
        console.log("[DOCX getImage called]", {
          tagName,
          key,
          tagValueType: typeof tagValue,
          hasDirectBuffer: !!(tagValue && typeof tagValue === "object"),
          hasImageMapBuffer: hasBuffer && typeof tagValue === "string",
          hasBuffer,
          bufferSize: hasBuffer ? buf!.length : 0,
        });
        if (!hasBuffer) {
          console.warn(
            `[imageModule.getImage] ${tagName} (key="${key}") missing/empty buffer - using transparent fallback`
          );
          return TRANSPARENT_PIXEL_PNG;
        }
        return buf!;
      } catch (err) {
        console.error(`[imageModule.getImage] ${tagName} threw:`, err);
        return TRANSPARENT_PIXEL_PNG;
      }
    },
    getSize: (_img: unknown, tagValue: unknown, tagName: string): [number, number] => {
      const key = typeof tagValue === "string" ? tagValue : "";
      console.log("[DOCX getSize called]", {
        tagName,
        key,
        tagValueType: typeof tagValue,
      });
      if (tagName === "routeMap" || key === "routeMap") return ROUTE_MAP_SIZE;
      if (tagName === "gaDrawing" || key === "gaDrawing") return GA_DRAWING_SIZE;
      if (tagName === "categoryIconKey" || key.startsWith("catIcon_")) return [40, 26];
      if (tagName === "categorySummaryIcon") return [180, 70];
      if (
        tagName === "photo" ||
        tagName === "photoKey" ||
        tagName === "observationPhoto" ||
        key.startsWith("obsPhoto_") ||
        key.startsWith("photo-")
      ) {
        return OBSERVATION_PHOTO_SIZE;
      }
      // Critical: never return undefined - the image module reads size[0]
      // unconditionally and crashes on undefined.
      return [120, 80];
    },
  });

  let doc: Docxtemplater;
  try {
    doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      modules: [imageModule],
      nullGetter: () => "",
    });
  } catch (err) {
    console.error("[export actual] docxtemplater init failed:", err);
    const e = err as { properties?: { errors?: unknown[] } };
    if (e?.properties?.errors) {
      console.error(
        "[export actual] init error details:",
        JSON.stringify(e.properties.errors, null, 2)
      );
    }
    throw new ExportStepError(
      "docx_render",
      "Docxtemplater init failed",
      err,
      JSON.stringify(e?.properties?.errors || null)
    );
  }

  // ----- Step 10: Render.
  // categorySummary entries expose `categoryIcon` (the buffer object) under the
  // template tag name `categorySummaryIcon` so the image-module getSize can
  // distinguish summary icons from per-row icons. categoryIcon stays as-is for
  // backward-compat with anyone reading the old key.
  const categorySummaryForRender = categorySummary.map((c) => ({
    categoryLabel: c.categoryLabel,
    count: c.count,
    countText: c.countText,
    hasCategoryIcon: c.hasCategoryIcon,
    categoryIcon: c.categoryIcon,
    categorySummaryIcon: c.categoryIcon,
  }));

  const renderData = {
    projectNameUpper: projectName.toUpperCase(),
    objective,
    conclusion,
    dateDot,
    dateDash,
    routeMap: imageMap.has("routeMap") ? "routeMap" : "",
    gaDrawing: imageMap.has("gaDrawing") ? "gaDrawing" : "",
    observations,
    categorySummary: categorySummaryForRender,
    hasCategorySummary: categorySummaryForRender.length > 0,
    categorySummaryTotal: categorySummaryForRender.reduce(
      (acc, c) => acc + c.count,
      0
    ),
  };

  console.log(
    "[template observations photos]",
    observations.map((o) => {
      const entry = o.photo ? imageMap.get(o.photo) : undefined;
      const buf = entry?.buffer;
      return {
        photoKey: o.photo,
        hasObservationPhoto: !!o.photo && !!buf && buf.length > 0,
        hasBuffer: !!buf && Buffer.isBuffer(buf) && buf.length > 0,
        bufferSize: buf?.length || 0,
        photoFallback: o.photoFallback,
      };
    })
  );

  console.log(
    "[category icon]",
    observations.map((o) => {
      const entry = o.categoryIconKey ? imageMap.get(o.categoryIconKey) : undefined;
      return {
        category: o.category,
        categoryIconKey: o.categoryIconKey,
        hasCategoryIcon: o.hasCategoryIcon,
        imageMapHasKey: !!o.categoryIconKey && imageMap.has(o.categoryIconKey),
        iconBufferSize: entry?.buffer?.length || 0,
        iconPath: entry?.path || null,
      };
    })
  );

  try {
    doc.render(renderData);
  } catch (err) {
    console.error("[export actual] docx render failed:", err);
    const e = err as { properties?: { errors?: unknown[] }; message?: string };
    let detail: string | undefined;
    if (e?.properties?.errors) {
      detail = JSON.stringify(e.properties.errors, null, 2);
      console.error("[export actual] docxtemplater errors:", detail);
    }
    throw new ExportStepError(
      "docx_render",
      e?.message || "DOCX render failed",
      err,
      detail
    );
  }

  // ----- Step 10b: Per-row header + body shading.
  // docxtemplater placeholders cannot dynamically rewrite cell shading
  // attributes, so we patch the rendered XML directly. The template's
  // observation cells use two anchor fills:
  //   - "4CAF50" → 6 header cells per observation (one per column)
  //   - "DDE8D7" → 6 body cells per observation (one per column)
  // Rewriting each in groups of 6 lets us apply per-row red/yellow/green
  // colors to BOTH the header and the body without touching the layout.
  try {
    const renderedZip = doc.getZip();
    const docFile = renderedZip.file("word/document.xml");
    if (docFile) {
      let xml = docFile.asText();

      const TEMPLATE_HEADER_FILL_RE = /w:fill="4CAF50"/g;
      const TEMPLATE_BODY_FILL_RE = /w:fill="DDE8D7"/g;

      const headerFillSequence = observations.flatMap((o) =>
        Array.from({ length: 6 }, () => String(o.headerFillColor || "4CAF50"))
      );
      const bodyFillSequence = observations.flatMap((o) =>
        Array.from({ length: 6 }, () => String(o.bodyFillColor || "D9EAD3"))
      );

      let headerIdx = 0;
      xml = xml.replace(TEMPLATE_HEADER_FILL_RE, () => {
        const fill = headerFillSequence[headerIdx] ?? "4CAF50";
        headerIdx += 1;
        return `w:fill="${fill}"`;
      });

      let bodyIdx = 0;
      xml = xml.replace(TEMPLATE_BODY_FILL_RE, () => {
        const fill = bodyFillSequence[bodyIdx] ?? "D9EAD3";
        bodyIdx += 1;
        return `w:fill="${fill}"`;
      });

      console.log("[difficulty table shading] applied", {
        observations: observations.length,
        headerReplacements: headerIdx,
        bodyReplacements: bodyIdx,
        expected: observations.length * 6,
      });
      if (
        headerIdx !== observations.length * 6 ||
        bodyIdx !== observations.length * 6
      ) {
        console.warn(
          "[difficulty table shading] replacement count differs from expected — header (4CAF50) or body (DDE8D7) anchor may have drifted in the template"
        );
      }

      renderedZip.file("word/document.xml", xml);
    } else {
      console.warn("[difficulty table shading] word/document.xml not found in rendered zip - skipping shading patch");
    }
  } catch (err) {
    console.error("[difficulty table shading] patch failed - leaving default fills:", err);
  }

  // ----- Step 11: Generate output buffer.
  let outBuf: Buffer;
  try {
    const generated = doc
      .getZip()
      .generate({ type: "nodebuffer", compression: "DEFLATE" }) as Buffer | Uint8Array;
    outBuf = Buffer.isBuffer(generated) ? generated : Buffer.from(generated);
  } catch (err) {
    console.error("[export actual] buffer generation failed:", err);
    throw new ExportStepError("buffer_generate", "DOCX zip generation failed", err);
  }

  const fileName = `${cleanFileName(projectName, "project")}-ALL.docx`;
  console.log("[export actual] success:", { fileName, bytes: outBuf.length });
  return { buffer: outBuf, fileName, projectName };
}

/**
 * Diagnostics shape used by GET /api/projects/[id]/export-debug.
 */
export type ExportDebugReport = {
  projectId: string;
  reportsCount: number;
  reportIds: string[];
  photosCount: number;
  photosByReportId: Record<string, Array<{ url: string; file_name: string | null; path: string | null }>>;
  gaDrawingImageUrl: string;
  routeMapUrl: string;
  fetchChecks: Array<{
    label: string;
    url: string;
    direct: { ok: boolean; status: number; contentType: string };
    signed?: { ok: boolean; status: number; contentType: string };
  }>;
};

export async function buildExportDebug(projectId: string): Promise<ExportDebugReport> {
  const reports = await safeQuery(
    "SELECT id FROM reports WHERE project_id = ? ORDER BY sort_order ASC, created_at ASC",
    [projectId]
  );
  const reportIds = reports.map((r) => String(r.id || "").trim()).filter(Boolean);

  let photos: Row[] = [];
  if (reportIds.length) {
    const placeholders = reportIds.map(() => "?").join(",");
    const [rows] = await pool.query(
      `SELECT id, report_id, url, file_name, path FROM report_photos WHERE report_id IN (${placeholders}) ORDER BY created_at ASC`,
      reportIds
    );
    photos = Array.isArray(rows) ? (rows as Row[]) : [];
  }

  const photosByReportId: Record<string, Array<{ url: string; file_name: string | null; path: string | null }>> = {};
  for (const p of photos) {
    const k = String(p.report_id || "").trim();
    (photosByReportId[k] ||= []).push({
      url: String(p.url || ""),
      file_name: p.file_name ?? null,
      path: p.path ?? null,
    });
  }

  const [gaRow] = await safeQuery(
    "SELECT * FROM project_ga_drawings WHERE project_id = ? LIMIT 1",
    [projectId]
  );
  const gaDrawingImageUrl = resolveImageUrl(gaRow, "image_url", "image_key");

  const [routePageRow] = await safeQuery(
    "SELECT * FROM project_route_pages WHERE project_id = ? ORDER BY created_at DESC LIMIT 1",
    [projectId]
  );
  const routeMapUrl = String(routePageRow?.map_file_url || "").trim();

  const checkUrls: Array<{ label: string; url: string }> = [];
  if (gaDrawingImageUrl) checkUrls.push({ label: "gaDrawing", url: gaDrawingImageUrl });
  if (routeMapUrl) checkUrls.push({ label: "routeMap", url: routeMapUrl });
  for (const k of Object.keys(photosByReportId)) {
    const first = photosByReportId[k][0];
    if (first?.url) checkUrls.push({ label: `photo[${k}]`, url: first.url });
  }

  const fetchChecks: ExportDebugReport["fetchChecks"] = [];
  for (const c of checkUrls.slice(0, 12)) {
    const direct = await (async () => {
      try {
        const r = await fetch(c.url);
        return {
          ok: r.ok,
          status: r.status,
          contentType: r.headers.get("content-type") || "",
        };
      } catch (err) {
        return { ok: false, status: -1, contentType: String(err) };
      }
    })();
    let signed: { ok: boolean; status: number; contentType: string } | undefined;
    if (!direct.ok) {
      const key = extractS3Key(c.url);
      if (key) {
        try {
          const signedUrl = await getReadSignedUrl(key, 600);
          const r = await fetch(signedUrl);
          signed = {
            ok: r.ok,
            status: r.status,
            contentType: r.headers.get("content-type") || "",
          };
        } catch (err) {
          signed = { ok: false, status: -1, contentType: String(err) };
        }
      }
    }
    fetchChecks.push({ label: c.label, url: c.url, direct, signed });
  }

  return {
    projectId,
    reportsCount: reports.length,
    reportIds,
    photosCount: photos.length,
    photosByReportId,
    gaDrawingImageUrl,
    routeMapUrl,
    fetchChecks,
  };
}
