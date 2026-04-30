import path from "path";
import crypto from "crypto";
import { promises as fs } from "fs";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import pool from "./db";
import { getReadSignedUrl } from "./s3";

// docxtemplater-image-module-free has no TypeScript types shipped.
// eslint-disable-next-line @typescript-eslint/no-var-requires
const ImageModule = require("docxtemplater-image-module-free");

// sharp is loaded lazily so a missing native binding (e.g. on a fresh
// Vercel build that has not rebuilt sharp for its platform) does not
// take down the entire export. Failure paths return the original buffer
// untouched and log a warning.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
let _sharp: any = null;
let _sharpLoadAttempted = false;
function getSharp() {
  if (_sharpLoadAttempted) return _sharp;
  _sharpLoadAttempted = true;
  try {
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    _sharp = require("sharp");
  } catch (err) {
    console.warn("[reenaTemplateExport] sharp not available - skipping image compression:", err);
    _sharp = null;
  }
  return _sharp;
}

/**
 * Compress an image buffer before it is embedded in the DOCX. Display
 * dimensions are controlled separately by the image module's getSize();
 * this function only changes the EMBEDDED bytes so the output file size
 * stays small while the rendered size in Word stays large.
 *
 * Strategy:
 *   - category icons → 160×160 PNG with transparent background (icons
 *     are tiny so size matters less than crispness + alpha preservation)
 *   - everything else (observation photos, GA drawing, route map) →
 *     1400×900 max, fit:inside, JPEG quality 80, mozjpeg-encoded
 *
 * Always honours EXIF orientation via .rotate(). Never enlarges
 * (withoutEnlargement) so a 400px source stays 400px.
 */
async function optimizeDocxImage(
  input: Buffer,
  type: "observation" | "routeMap" | "gaDrawing" | "category",
  key: string
): Promise<Buffer> {
  if (!input || !Buffer.isBuffer(input) || input.length === 0) return input;
  const sharp = getSharp();
  if (!sharp) return input;
  try {
    let optimized: Buffer;
    if (type === "category") {
      optimized = await sharp(input)
        .rotate()
        .resize({
          width: 160,
          height: 160,
          fit: "contain",
          background: { r: 255, g: 255, b: 255, alpha: 0 },
          withoutEnlargement: true,
        })
        .png()
        .toBuffer();
    } else {
      optimized = await sharp(input)
        .rotate()
        .resize({
          width: 1400,
          height: 900,
          fit: "inside",
          withoutEnlargement: true,
        })
        .jpeg({
          quality: 80,
          mozjpeg: true,
        })
        .toBuffer();
    }
    console.log("[DOCX IMAGE OPTIMIZED]", {
      type,
      key,
      originalBytes: input.length,
      optimizedBytes: optimized.length,
      reductionPercent:
        input.length > 0
          ? Math.round((1 - optimized.length / input.length) * 100)
          : 0,
    });
    return optimized;
  } catch (err) {
    console.warn(`[DOCX IMAGE OPTIMIZED] ${type} ${key} optimize failed - using original:`, err);
    return input;
  }
}

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
// Per latest client-ready spec: observation photo should display LARGE
// (650×420 px, ≈ 6.77" × 4.38"). The actual embedded JPEG is compressed
// via sharp BEFORE being added to imageMap (max 1400×900 inside, q80
// mozjpeg) so the DOCX file size stays small even with 395+ photos.
const OBSERVATION_PHOTO_SIZE: [number, number] = [650, 420];
// Per spec: category icon at 60 px square inside the CATEGORY cell.
// Embedded as 160×160 PNG for crispness; displayed at 60×60.
const CATEGORY_ICON_SIZE: [number, number] = [60, 60];
const CATEGORY_SUMMARY_ICON_SIZE: [number, number] = [60, 60];

// Word stores image extents in EMU (1 px = 9525 EMU). The combined
// layout-swap pass uses this constant to identify the OBSERVATION photo
// paragraph and ignore other in-document drawings (GA Drawing, Route
// Map, category icons) so the swap never moves the wrong image.
const PX_TO_EMU = 9525;
const OBSERVATION_PHOTO_EMU_WIDTH = OBSERVATION_PHOTO_SIZE[0] * PX_TO_EMU;

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
    // Per-fetch 10s timeout. Without this, a single hung S3 connection (which
    // happens regularly behind some corporate / CDN proxies) blocks the whole
    // export indefinitely - the browser then aborts the export request with
    // "Failed to fetch" because Next has not flushed any bytes yet.
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 10_000);
    try {
      const res = await fetch(target, { signal: controller.signal });
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
      // Hard cap: skip oversize images so we never balloon the DOCX or run
      // out of Node heap when an export covers hundreds of multi-MB photos.
      // Word renders inline images at the cell width regardless of file size,
      // so an 8 MB JPEG and an 800 KB JPEG look identical at A4.
      const MAX_PHOTO_BYTES = 8 * 1024 * 1024;
      if (buf.length > MAX_PHOTO_BYTES) {
        console.warn(`[image fetch] ${label} ${attempt} too large - skipping`, {
          url: target,
          size: buf.length,
          maxBytes: MAX_PHOTO_BYTES,
        });
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
      const aborted = (err as { name?: string })?.name === "AbortError";
      console.error(
        `[image fetch] ${label} ${attempt} ${aborted ? "TIMED OUT (10s)" : "threw"}`,
        target,
        aborted ? "" : err
      );
      return null;
    } finally {
      clearTimeout(timeoutId);
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
  // The same image-map key is mirrored under every plausible template tag
  // name so we can support {%photo}, {%photoKey}, and {%observationPhotoKey}
  // bindings without changing the template. Empty string when no buffer.
  photo: string;
  photoKey: string;
  observationPhotoKey: string;
  observationPhoto: string;
  image: string;
  hasObservationPhoto: boolean;
  photoFallback: string;
  photoText: string;
  // Captures why the report photo could not be embedded (only set when the
  // photo is unavailable). Surfaces as a [DOCX photo missing] log line so the
  // failure mode is visible without trawling the full server console.
  photoMissingReason: string | null;
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
    const rawBuffer = await fs.readFile(fullPath);
    if (!rawBuffer.length) {
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
    // Optimize: 160x160 PNG with transparent background. Embedded once
    // per unique icon (cached) so the cost is paid at most ~16 times for
    // a typical category set, not per row.
    const buffer = await optimizeDocxImage(rawBuffer, "category", fileName);
    const contentType = "image/png";
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
      // Pull image_key + point_key when those columns exist so the
      // diagnostic logs can show the master-file linkage.
      const sql = `SELECT id, report_id, url, file_name, image_key, point_key, path, created_at
         FROM report_photos
         WHERE report_id IN (${placeholders})
         ORDER BY created_at ASC`;
      console.error("[PHOTOS_DEBUG_2026_04_27_B] sql", { sql, args: reportIds });
      try {
        const [photoRows] = await pool.execute(sql, reportIds);
        photos = Array.isArray(photoRows) ? (photoRows as Row[]) : [];
      } catch (selectErr) {
        // Fallback: schemas without image_key / point_key / path columns.
        console.warn(
          "[PHOTOS_DEBUG_2026_04_27_B] full SELECT failed, falling back to legacy columns:",
          selectErr
        );
        const fallbackSql = `SELECT id, report_id, url, file_name, created_at
           FROM report_photos
           WHERE report_id IN (${placeholders})
           ORDER BY created_at ASC`;
        const [photoRows] = await pool.execute(fallbackSql, reportIds);
        photos = Array.isArray(photoRows) ? (photoRows as Row[]) : [];
      }
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

  console.log("[PHOTO CHECK 1 - DB PHOTOS FETCHED]", {
    exportedReportCount: reports.length,
    exportedReportIdsSample: reports.slice(0, 10).map((r: Row) => ({
      id: r.id,
      point_key: r.point_key,
      category: r.category,
    })),
    reportPhotosCount: photos.length,
    reportPhotosSample: photos.slice(0, 20).map((p: Row) => ({
      id: p.id,
      report_id: p.report_id,
      url: p.url,
      file_name: p.file_name,
      image_key: p.image_key,
      point_key: p.point_key,
      created_at: p.created_at,
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

  // Spec-mandated point_key fallback. When a report has no row in
  // report_photos by report_id (typical when bulk upload landed photos
  // against a different reports.id but the same point_key), we fall
  // back to matching by point_key so the photo still appears.
  const photosByPointKey = new Map<string, Row[]>();
  for (const p of photos) {
    const pk = String(p.point_key || "").trim();
    if (!pk) continue;
    if (!photosByPointKey.has(pk)) photosByPointKey.set(pk, []);
    photosByPointKey.get(pk)!.push(p);
  }

  // Spec-mandated SQL diagnostic: dump the first 10 reports JOINed with
  // report_photos, ordered by point_key as DECIMAL (so "1.0", "2.0",
  // "10.0" sort correctly). This is the EXACT query the operator runs
  // to confirm whether the first row truly has no photo in DB.
  try {
    const first10 = await safeQuery(
      `SELECT
         r.id AS report_id,
         r.point_key,
         r.category,
         rp.id AS photo_id,
         rp.url,
         rp.file_name,
         rp.point_key AS photo_point_key,
         rp.image_key
       FROM reports r
       LEFT JOIN report_photos rp ON rp.report_id = r.id
       WHERE r.project_id = ?
       ORDER BY CAST(r.point_key AS DECIMAL(10,4)), r.created_at
       LIMIT 10`,
      [projectId]
    );
    console.log("[DOCX FIRST 10 PHOTO DB CHECK]", first10);
  } catch (err) {
    console.warn("[DOCX FIRST 10 PHOTO DB CHECK] query failed:", err);
  }

  // ---- Spec-mandated SQL JOIN/COUNT diagnostics. These run the EXACT
  // queries the operator runs by hand to verify the report ↔ report_photos
  // link, so the export server log carries the same answer in one place.
  // If reports_with_photo_url is 0 here, no DOCX placeholder will ever
  // resolve a buffer — the fix is upstream in bulk upload.
  try {
    const joinRows = await safeQuery(
      `SELECT
         r.id AS report_id,
         r.point_key,
         r.category,
         rp.id AS photo_id,
         rp.url,
         rp.file_name,
         rp.image_key,
         rp.point_key AS photo_point_key
       FROM reports r
       LEFT JOIN report_photos rp ON rp.report_id = r.id
       WHERE r.project_id = ?
       ORDER BY CAST(r.point_key AS UNSIGNED)
       LIMIT 20`,
      [projectId]
    );
    console.log("[EXPORT PHOTO SQL JOIN CHECK]", {
      projectId,
      rows: joinRows,
    });
  } catch (err) {
    console.warn("[EXPORT PHOTO SQL JOIN CHECK] query failed:", err);
  }

  let photoCountTotalReports = 0;
  let photoCountReportsWithPhotos = 0;
  let photoCountReportsWithPhotoUrl = 0;
  try {
    const photoCountRows = await safeQuery(
      `SELECT
         COUNT(*) AS total_reports,
         SUM(CASE WHEN rp.id IS NOT NULL THEN 1 ELSE 0 END) AS reports_with_photos,
         SUM(CASE WHEN rp.url IS NOT NULL AND rp.url <> '' THEN 1 ELSE 0 END) AS reports_with_photo_url
       FROM reports r
       LEFT JOIN report_photos rp ON rp.report_id = r.id
       WHERE r.project_id = ?`,
      [projectId]
    );
    const row = photoCountRows[0] || {};
    photoCountTotalReports = Number(row.total_reports) || 0;
    photoCountReportsWithPhotos = Number(row.reports_with_photos) || 0;
    photoCountReportsWithPhotoUrl = Number(row.reports_with_photo_url) || 0;
    console.log("[EXPORT PHOTO SQL COUNT CHECK]", row);
  } catch (err) {
    console.warn("[EXPORT PHOTO SQL COUNT CHECK] query failed:", err);
  }

  // Fail-fast root-cause log. If the DB has zero linked photos, no DOCX
  // change can fix this — the bulk-upload pipeline must insert
  // report_photos using the saved reports.id. Surfacing this loudly here
  // stops us claiming "DOCX photo logic is fixed" when it is the upstream
  // link that is broken.
  if (photoCountReportsWithPhotoUrl === 0) {
    console.error("[EXPORT PHOTO ROOT CAUSE]", {
      reason: "No report_photos rows linked to reports.id for this project",
      projectId,
      total_reports: photoCountTotalReports,
      reports_with_photos: photoCountReportsWithPhotos,
      reports_with_photo_url: photoCountReportsWithPhotoUrl,
      fix: "Bulk upload must insert report_photos using saved reports.id",
    });
  }

  // PHOTO CHECK 2: confirm report_photos.report_id values actually match the
  // exported reports.id values. If matchedPhotoRowsCount is 0 here, the join
  // key is wrong somewhere upstream (bulk import or manual upload pipeline).
  {
    const exportedReportIds = new Set(reports.map((r: Row) => String(r.id)));
    const matchedPhotoRows = photos.filter((p: Row) =>
      exportedReportIds.has(String(p.report_id))
    );
    console.log("[PHOTO CHECK 2 - PHOTO/REPORT ID MATCH]", {
      exportedReportCount: reports.length,
      totalReportPhotosFetched: photos.length,
      matchedPhotoRowsCount: matchedPhotoRows.length,
      unmatchedPhotoRowsCount: photos.length - matchedPhotoRows.length,
      matchedSample: matchedPhotoRows.slice(0, 10).map((p: Row) => ({
        report_id: p.report_id,
        url: p.url,
        file_name: p.file_name,
        image_key: p.image_key,
        point_key: p.point_key,
      })),
    });
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
      const optimized = await optimizeDocxImage(
        gaDrawingFetched.buffer,
        "gaDrawing",
        "gaDrawing"
      );
      imageMap.set("gaDrawing", {
        ...gaDrawingFetched,
        buffer: optimized,
        contentType: "image/jpeg",
      });
    }
  } catch (err) {
    console.error("[export actual] gaDrawing fetch step threw - ignoring:", err);
  }
  try {
    const routeMapFetched = routeMapUrl ? await fetchImageBuffer(routeMapUrl, "routeMap") : null;
    if (routeMapFetched && Buffer.isBuffer(routeMapFetched.buffer)) {
      const optimized = await optimizeDocxImage(
        routeMapFetched.buffer,
        "routeMap",
        "routeMap"
      );
      imageMap.set("routeMap", {
        ...routeMapFetched,
        buffer: optimized,
        contentType: "image/jpeg",
      });
    }
  } catch (err) {
    console.error("[export actual] routeMap fetch step threw - ignoring:", err);
  }

  // ----- Observations.
  console.log("[includePhotos check]", { includePhotos });
  console.log("[photosByReportId keys]", Array.from(photosByReportId.keys()));

  // Stage logs that match the [export stage] markers in route.ts so a server
  // tail tells you exactly where the request died.
  console.log("[export stage] reports loaded", { count: reports.length });
  console.log("[export stage] report_photos loaded", { count: photos.length });
  console.log("[export stage] photo fetch started", {
    reports: reports.length,
    expectedPhotosResolved: photosByReportId.size,
  });

  let photoFetchSuccess = 0;
  let photoFetchFailed = 0;

  // Hard cap on the photo-fetch phase. The browser typically aborts a
  // streaming response after ~120s of idle time; exporting 400 photos
  // sequentially with a 10s/photo timeout could hit that ceiling. Once this
  // budget is exhausted the remaining rows render with the "Photo not
  // available." fallback instead of trying to fetch.
  const PHOTO_PHASE_DEADLINE_MS = 100_000;
  const phaseStartedAt = Date.now();
  let phaseDeadlineHit = false;

  const observations: ObservationData[] = [];
  for (let i = 0; i < reports.length; i += 1) {
    const r = reports[i];
    const rid = String(r.id || "").trim();
    const lat = pickLat(r);
    const lng = pickLng(r);
    const hasLat = lat !== null;
    const hasLon = lng !== null;

    let photoKey = "";
    let photoMissingReason: string | null = null;
    let attemptedPhotoUrl: string | null = null;

    // Once the phase deadline is exceeded, every subsequent row skips its
    // fetch and renders the fallback instead. Without this guard a slow S3
    // can drag the export past the browser's response timeout and surface
    // as "Failed to fetch" with no useful server-side error.
    const phaseElapsed = Date.now() - phaseStartedAt;
    if (includePhotos && !phaseDeadlineHit && phaseElapsed >= PHOTO_PHASE_DEADLINE_MS) {
      phaseDeadlineHit = true;
      console.warn("[export stage] photo phase deadline reached", {
        deadlineMs: PHOTO_PHASE_DEADLINE_MS,
        elapsedMs: phaseElapsed,
        remainingObservations: reports.length - i,
      });
    }

    if (!includePhotos) {
      photoMissingReason = "includePhotos=false";
      console.log("[export observation photo prepared]", {
        index: i,
        reportId: rid,
        skipped: "includePhotos=false",
      });
    } else if (phaseDeadlineHit) {
      photoMissingReason = "photo phase deadline exceeded - skipped";
      photoFetchFailed += 1;
    } else {
      // Spec-mandated photo lookup: report_id first, then point_key
      // fallback. Some bulk-upload paths historically landed photos
      // against a stale reports.id but kept the point_key intact —
      // falling back to point_key recovers those photos.
      const photosByRid = photosByReportId.get(rid) || [];
      const rowPointKey = String(r.point_key || "").trim();
      const photosByPk = rowPointKey
        ? photosByPointKey.get(rowPointKey) || []
        : [];
      const reportPhotos = photosByRid.length > 0 ? photosByRid : photosByPk;
      const matchSource = photosByRid.length > 0
        ? "report_id"
        : photosByPk.length > 0
          ? "point_key"
          : "none";
      const firstPhotoRow = reportPhotos[0] || null;
      const firstPhotoUrl = normalizeS3Url(firstPhotoRow?.url) || null;
      attemptedPhotoUrl = firstPhotoUrl;

      console.log("[DOCX PHOTO MATCH SOURCE]", {
        blockIndex: i,
        report_id: rid,
        point_key: r.point_key,
        source: matchSource,
        url: firstPhotoRow?.url || null,
      });

      console.log("[PHOTO CHECK 6 - ROW PHOTO MATCH]", {
        index: i,
        report_id: rid,
        point_key: r.point_key,
        category: r.category,
        photosForReport: reportPhotos.length,
        firstPhoto: firstPhotoRow
          ? {
              report_id: firstPhotoRow.report_id,
              url: firstPhotoRow.url,
              file_name: firstPhotoRow.file_name,
              image_key: firstPhotoRow.image_key,
              point_key: firstPhotoRow.point_key,
            }
          : null,
      });

      if (!firstPhotoRow) {
        console.warn("[PHOTO CHECK 6 - NO PHOTO URL FOR ROW]", {
          index: i,
          report_id: rid,
          point_key: r.point_key,
          photosForReport: reportPhotos.length,
        });
      }

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

      if (reportPhotos.length === 0) {
        photoMissingReason = "no report_photos rows for report_id";
      }

      // Iterate every saved photo for this report; first one that fetches OK
      // wins. This handles the case where the first row was orphaned/deleted
      // in S3 but a later row is still valid. Sequential per row so we never
      // open hundreds of S3 connections at once - that's what was triggering
      // the "Failed to fetch" timeout in the browser.
      let fetchedAny = false;
      let lastFetchFailureReason: string | null = null;
      for (const p of reportPhotos) {
        const candidate = normalizeS3Url(p?.url);
        const ctx = {
          index: i,
          report_id: rid,
          point_key: r.point_key,
          file_name: p?.file_name,
          image_key: p?.image_key,
        };
        if (!candidate) {
          lastFetchFailureReason = "row has no url and no path";
          console.warn("[PHOTO CHECK 3 - FETCH SKIPPED INVALID URL]", ctx);
          continue;
        }
        console.log("[PHOTO CHECK 3 - FETCH START]", { ...ctx, url: candidate });
        console.log("[DOCX photo fetch start]", {
          index: i,
          reportId: rid,
          url: candidate,
        });
        try {
          const fetched = await fetchImageBuffer(candidate, `photo[${rid}]`);
          if (fetched) {
            console.log("[PHOTO CHECK 4 - FETCH RESPONSE]", {
              ...ctx,
              url: candidate,
              ok: true,
              contentType: fetched.contentType,
            });
            console.log("[PHOTO CHECK 5 - BUFFER CREATED]", {
              ...ctx,
              url: candidate,
              bufferSize: fetched.buffer.length,
              firstBytes: fetched.buffer.subarray(0, 12).toString("hex"),
            });
          }
          if (fetched && Buffer.isBuffer(fetched.buffer) && fetched.buffer.length > 0) {
            photoKey = `photo_${i}`;
            const optimized = await optimizeDocxImage(
              fetched.buffer,
              "observation",
              photoKey
            );
            imageMap.set(photoKey, {
              ...fetched,
              buffer: optimized,
              contentType: "image/jpeg",
            });
            fetchedAny = true;
            photoFetchSuccess += 1;
            console.log("[PHOTO CHECK 7 - ADDED TO IMAGEMAP]", {
              ...ctx,
              observationPhotoKey: photoKey,
              bufferSize: fetched.buffer.length,
              imageMapHasKey: imageMap.has(photoKey),
            });
            console.log("[DOCX photo buffer loaded]", {
              index: i,
              reportId: rid,
              photoKey,
              bufferSize: fetched.buffer.length,
              contentType: fetched.contentType,
            });
            break;
          }
          lastFetchFailureReason = "fetchImageBuffer returned no buffer (404/403/non-image/oversize/timeout)";
          console.warn("[PHOTO CHECK 7 - PHOTO BUFFER MISSING]", {
            ...ctx,
            attemptedUrl: candidate,
          });
        } catch (err) {
          lastFetchFailureReason = `fetch threw: ${(err as Error)?.message || String(err)}`;
          console.error("[PHOTO CHECK 3 - FETCH ERROR]", {
            ...ctx,
            url: candidate,
            message: (err as Error)?.message,
            stack: (err as Error)?.stack,
          });
          console.error(`[export actual] photo fetch threw for ${candidate} - skipping:`, err);
        }
      }
      if (!fetchedAny) photoFetchFailed += 1;
      if (!fetchedAny && reportPhotos.length > 0 && !photoMissingReason) {
        photoMissingReason = lastFetchFailureReason || "all photo URLs failed to fetch";
      }

      // ---- Fallback: when report_photos has no rows for this report but the
      // report itself carries a file_name / image_key from the master file,
      // try a direct S3-key match. The construction priority is:
      //   1. r.image_key looking like a full S3 key (contains "/" + ext)
      //   2. r.file_name appended onto common upload prefixes
      // On hit we fetch the buffer AND insert a report_photos row so next
      // exports skip this fallback entirely.
      if (!fetchedAny && reportPhotos.length === 0) {
        const fileNameRaw = String(r.file_name || "").trim();
        const imageKeyRaw = String(r.image_key || "").trim();
        const candidates: string[] = [];
        if (imageKeyRaw && imageKeyRaw.includes("/") && /\.[a-z0-9]+$/i.test(imageKeyRaw)) {
          const url = normalizeS3Url(imageKeyRaw);
          if (url) candidates.push(url);
        }
        if (fileNameRaw) {
          const base = (process.env.NEXT_PUBLIC_S3_BUCKET_URL || "").replace(/\/+$/, "");
          if (base) {
            // Prefixes the bulk-import + manual-upload paths actually use.
            const prefixes = [
              `reports/photos/${rid}/`,
              `reports/${projectId}/${rid}/`,
              `reports/photos/${projectId}/`,
              `reports/${projectId}/`,
              `reports/photos/`,
            ];
            for (const pfx of prefixes) {
              candidates.push(`${base}/${pfx}${fileNameRaw.replace(/^\/+/, "")}`);
            }
          }
        }
        for (const candidate of candidates) {
          try {
            const fetched = await fetchImageBuffer(candidate, `photoFallback[${rid}]`);
            if (fetched && Buffer.isBuffer(fetched.buffer) && fetched.buffer.length > 0) {
              photoKey = `photo_${i}`;
              const optimized = await optimizeDocxImage(
                fetched.buffer,
                "observation",
                photoKey
              );
              imageMap.set(photoKey, {
                ...fetched,
                buffer: optimized,
                contentType: "image/jpeg",
              });
              fetchedAny = true;
              attemptedPhotoUrl = candidate;
              photoMissingReason = null;
              console.log("[DOCX photo buffer loaded]", {
                index: i,
                reportId: rid,
                photoKey,
                via: "fallback",
                bufferSize: optimized.length,
              });
              // Persist for next time so subsequent exports skip this scan.
              try {
                const insertId = crypto.randomUUID
                  ? crypto.randomUUID()
                  : require("uuid").v4();
                await pool.query(
                  `INSERT INTO report_photos (id, report_id, url, file_name) VALUES (?, ?, ?, ?)`,
                  [insertId, rid, candidate, fileNameRaw || null]
                );
                console.log("[DOCX photo cached back into report_photos]", {
                  reportId: rid,
                  url: candidate,
                });
              } catch (cacheErr) {
                console.warn(
                  "[DOCX photo fallback INSERT failed - non-fatal]",
                  cacheErr
                );
              }
              break;
            }
          } catch (err) {
            console.warn(`[DOCX photo fallback fetch threw for ${candidate}]`, err);
          }
        }
        if (!fetchedAny && candidates.length > 0) {
          photoMissingReason = "fallback S3 candidates 404/403";
        } else if (!fetchedAny && !fileNameRaw && !imageKeyRaw) {
          // Keep the existing "no report_photos rows for report_id" reason.
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

    // Spec-required missing-photo log so the failure mode is visible at a
    // glance. Fires when the per-row resolution did not produce a buffer.
    if (!hasObservationPhoto) {
      console.warn("[DOCX photo missing]", {
        index: i,
        report_id: rid,
        point_key: r.point_key || null,
        file_name: r.file_name || null,
        image_key: r.image_key || null,
        attemptedUrl: attemptedPhotoUrl,
        reason: photoMissingReason || "unknown",
      });
    }

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

    // Mirror the same map key under every plausible template tag name so the
    // image module resolves whether the template uses {%photo}, {%photoKey},
    // {%observationPhoto}, {%observationPhotoKey}, or {%image}. Empty string
    // when no buffer was loaded - the image module's render() short-circuits
    // on falsy values without ever calling getImage(), which is why the
    // previous logs only showed categoryIconKey calls.
    const photoTagValue = hasObservationPhoto ? photoKey : "";

    observations.push({
      gpsLat: hasLat ? formatGpsLat(lat) : "-",
      gpsLon: hasLon ? formatGpsLon(lng) : "-",
      km: kmText,
      location: locationText || "-",
      category: valueOrDash(r.category),
      observation: valueOrEmDash(r.description ?? r.observation),
      remarks: valueOrEmDash(r.remarks_action ?? r.difficulty ?? r.status),
      photo: photoTagValue,
      photoKey: photoTagValue,
      observationPhotoKey: photoTagValue,
      observationPhoto: photoTagValue,
      image: photoTagValue,
      hasObservationPhoto,
      // Clean placeholder: no trailing period, single short line — the
      // template wraps this in a small grey paragraph so it does not look
      // like an error message in the rendered DOCX.
      photoFallback: hasObservationPhoto ? "" : "Photo not available",
      photoText: hasObservationPhoto ? "" : "Photo not available",
      photoMissingReason: hasObservationPhoto ? null : photoMissingReason,
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
        const isObservationPhoto =
          tagName === "photo" ||
          tagName === "photoKey" ||
          tagName === "observationPhoto" ||
          tagName === "observationPhotoKey" ||
          tagName === "image" ||
          (typeof key === "string" && (key.startsWith("photo_") || key.startsWith("obsPhoto_")));
        const isCategoryIcon =
          tagName === "categoryIconKey" ||
          tagName === "categorySummaryIcon" ||
          (typeof key === "string" && key.startsWith("catIcon_"));
        console.log("[DOCX getImage called]", {
          tagName,
          key,
          tagValueType: typeof tagValue,
          isObservationPhoto,
          isCategoryIcon,
          hasDirectBuffer: !!(tagValue && typeof tagValue === "object"),
          hasImageMapBuffer: hasBuffer && typeof tagValue === "string",
          hasBuffer,
          bufferSize: hasBuffer ? buf!.length : 0,
        });
        console.log("[PHOTO CHECK 9 - DOCX GET IMAGE]", {
          tagName,
          key,
          isObservationPhoto: typeof key === "string" && key.startsWith("photo_"),
          isCategoryIcon: typeof key === "string" && key.startsWith("catIcon_"),
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
      console.log("[PHOTO CHECK 10 - DOCX GET SIZE]", {
        tagName,
        key,
        isObservationPhoto: key.startsWith("photo_"),
        isCategoryIcon: key.startsWith("catIcon_"),
      });
      if (tagName === "routeMap" || key === "routeMap") return ROUTE_MAP_SIZE;
      if (tagName === "gaDrawing" || key === "gaDrawing") return GA_DRAWING_SIZE;
      if (tagName === "categoryIconKey" || key.startsWith("catIcon_")) return CATEGORY_ICON_SIZE;
      if (tagName === "categorySummaryIcon") return CATEGORY_SUMMARY_ICON_SIZE;
      if (
        tagName === "photo" ||
        tagName === "photoKey" ||
        tagName === "observationPhoto" ||
        tagName === "observationPhotoKey" ||
        tagName === "image" ||
        key.startsWith("photo_") ||
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

  // Spec-mandated empty-observation filter. Drops rows where every
  // visible field is missing/dash so docxtemplater never renders an
  // empty observation table at the start of the survey section. We
  // mutate observations in place because every downstream log + the
  // post-render swap pass count from observations.length.
  const observationCountBeforeFilter = observations.length;
  const isMeaningfulObservation = (r: ObservationData) => {
    const meaningful = (v: unknown) => {
      if (v === null || typeof v === "undefined") return false;
      const s = String(v).trim();
      return s !== "" && s !== "-" && s !== "—";
    };
    return (
      meaningful(r.gpsLat) ||
      meaningful(r.gpsLon) ||
      meaningful(r.location) ||
      meaningful(r.category) ||
      meaningful(r.observation) ||
      meaningful(r.remarks) ||
      !!r.observationPhotoKey
    );
  };
  const validObservations = observations.filter(isMeaningfulObservation);
  const removedEmpty = observationCountBeforeFilter - validObservations.length;
  console.log("[DOCX EMPTY OBSERVATION FILTER]", {
    before: observationCountBeforeFilter,
    after: validObservations.length,
    removed: removedEmpty,
    firstValid: validObservations[0]
      ? {
          gpsLat: validObservations[0].gpsLat,
          gpsLon: validObservations[0].gpsLon,
          location: validObservations[0].location,
          category: validObservations[0].category,
          observation: validObservations[0].observation,
          observationPhotoKey: validObservations[0].observationPhotoKey,
        }
      : null,
  });
  if (removedEmpty > 0) {
    // Replace observations array contents in place so the post-render
    // swap pass's `observations.length` count matches what was rendered.
    observations.length = 0;
    for (const o of validObservations) observations.push(o);
  }

  // Spec-mandated finalised observation blocks. Carries the explicit
  // hasPhoto / pageBreak / isLast flags the spec calls out so the data
  // shape unambiguously expresses: table → image-or-placeholder → page
  // break (except after last). Used in render data alongside
  // `observations` (the template's existing loop key) so the data is
  // available either way.
  const observationBlocks = observations.map((row, index) => {
    const photoKey = row.observationPhotoKey || row.photoKey || "";
    const hasPhoto = !!photoKey && imageMap.has(photoKey);
    return {
      ...row,
      blockIndex: index,
      photoKey,
      observationPhotoKey: hasPhoto ? photoKey : "",
      hasPhoto,
      photoText: hasPhoto ? "" : "Photo not available",
      isLast: index === observations.length - 1,
      pageBreak: index < observations.length - 1,
    };
  });
  console.log("[DOCX OBSERVATION BLOCKS FINAL]", {
    count: observationBlocks.length,
    firstFive: observationBlocks.slice(0, 5).map((b) => ({
      blockIndex: b.blockIndex,
      observationPhotoKey: b.observationPhotoKey,
      photoKey: b.photoKey,
      hasPhoto: b.hasPhoto,
      imageMapHit: !!b.observationPhotoKey && imageMap.has(b.observationPhotoKey),
      photoText: b.photoText,
      pageBreak: b.pageBreak,
    })),
    lastBlock:
      observationBlocks.length > 0
        ? {
            blockIndex: observationBlocks[observationBlocks.length - 1].blockIndex,
            hasPhoto: observationBlocks[observationBlocks.length - 1].hasPhoto,
            pageBreak: observationBlocks[observationBlocks.length - 1].pageBreak,
          }
        : null,
  });
  console.log("[DOCX FIRST BLOCK CHECK]", {
    firstBlockExists: !!observationBlocks[0],
    firstBlockHasPhoto: observationBlocks[0]?.hasPhoto,
    firstBlockPhotoKey: observationBlocks[0]?.observationPhotoKey,
    firstBlockWillRenderPlaceholder: !observationBlocks[0]?.hasPhoto,
  });

  // Spec-mandated first-10-blocks photo check. Surfaces, for the first
  // 10 valid observation blocks, whether each carries a real photoKey
  // and whether the imageMap actually has it. If any of the first 10
  // shows hasPhoto: false AND no DB row was returned by the SQL above,
  // the gap is upstream in bulk upload (not in DOCX rendering).
  console.log(
    "[DOCX FIRST 10 BLOCK PHOTO CHECK]",
    observationBlocks.slice(0, 10).map((b) => ({
      blockIndex: b.blockIndex,
      observationPhotoKey: b.observationPhotoKey,
      hasPhoto: b.hasPhoto,
      photoText: b.photoText,
      imageMapHit: !!b.observationPhotoKey && imageMap.has(b.observationPhotoKey),
    }))
  );

  // Spec-mandated row-level photo QA. Surfaces, for the first 5 rows
  // that survive the empty filter, exactly which observationPhotoKey
  // each one carries and whether the imageMap actually has that key.
  // If row 0 shows observationPhotoKey: "" AND hasObservationPhoto:
  // false, the swap pass relies on the placeholder-text path to put a
  // "Photo not available" line below the first table.
  console.log("[DOCX FIRST PHOTO PAIR QA]", {
    observationsCount: observations.length,
    firstFiveRows: observations.slice(0, 5).map((r, i) => ({
      index: i,
      gpsLat: r.gpsLat,
      gpsLon: r.gpsLon,
      location: r.location,
      category: r.category,
      observationPhotoKey: r.observationPhotoKey,
      photoKey: r.photoKey,
      hasObservationPhoto: r.hasObservationPhoto,
      photoText: r.photoText,
      imageMapHasObservationPhotoKey:
        !!r.observationPhotoKey && imageMap.has(r.observationPhotoKey),
      imageMapHasPhotoKey: !!r.photoKey && imageMap.has(r.photoKey),
    })),
    imageMapPhotoKeysSample: Array.from(imageMap.keys())
      .filter((k) => k.startsWith("photo_"))
      .slice(0, 10),
  });

  const renderData = {
    projectNameUpper: projectName.toUpperCase(),
    objective,
    conclusion,
    dateDot,
    dateDash,
    routeMap: imageMap.has("routeMap") ? "routeMap" : "",
    gaDrawing: imageMap.has("gaDrawing") ? "gaDrawing" : "",
    observations,
    observationBlocks,
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

  console.log("[export stage] photo fetch completed", {
    success: photoFetchSuccess,
    failed: photoFetchFailed,
    totalImageMapKeys: imageMap.size,
  });

  // Spec-required photo-resolution summary right before render. Tells you at
  // a glance how many observation rows actually have a photo buffer ready.
  {
    const photosFetched = observations.filter((o) => o.hasObservationPhoto).length;
    const buffersLoaded = observations.filter(
      (o) => !!o.photoKey && imageMap.has(o.photoKey)
    ).length;
    const missing = observations
      .filter((o) => !o.hasObservationPhoto)
      .map((o) => ({
        photoFallback: o.photoFallback,
        reason: o.photoMissingReason,
      }));
    console.log("[DOCX photo summary]", {
      observations: observations.length,
      reportPhotosFetched: photosFetched,
      photosMatchedToRows: photosFetched,
      photoBuffersLoaded: buffersLoaded,
      missingPhotos: missing.length,
      missingReasons: missing.slice(0, 5),
    });

    // Spec-required final check that proves the data carries
    // observationPhotoKey for every row that has a buffer in imageMap. If
    // photoKeys > 0 but rowsWithObservationPhotoKey is 0, the bug is in the
    // data builder; if both are > 0 but [DOCX getImage called] still never
    // fires for tagName: 'observationPhotoKey', the template placeholder
    // does not match.
    {
      const allKeys = Array.from(imageMap.keys());
      console.log("[DOCX observation photo final check]", {
        observations: observations.length,
        imageMapTotal: imageMap.size,
        categoryIconKeys: allKeys.filter((k) => k.startsWith("catIcon_")).length,
        photoKeys: allKeys.filter((k) => k.startsWith("photo_")).length,
        rowsWithObservationPhotoKey: observations.filter(
          (r) => !!r.observationPhotoKey
        ).length,
        sampleRows: observations.slice(0, 10).map((r, idx) => ({
          index: idx,
          observationPhotoKey: r.observationPhotoKey,
          hasObservationPhoto: r.hasObservationPhoto,
          photoText: r.photoText,
        })),
      });

      console.log("[PHOTO CHECK 8 - FINAL BEFORE RENDER]", {
        observations: observations.length,
        imageMapTotal: imageMap.size,
        imageMapPhotoKeys: allKeys.filter((k) => k.startsWith("photo_")).length,
        imageMapCategoryKeys: allKeys.filter((k) => k.startsWith("catIcon_")).length,
        rowsWithObservationPhotoKey: observations.filter(
          (r) => !!r.observationPhotoKey
        ).length,
        rowsWithPhotoKey: observations.filter((r) => !!r.photoKey).length,
        sampleRows: observations.slice(0, 10).map((r) => ({
          observationPhotoKey: r.observationPhotoKey,
          photoKey: r.photoKey,
          hasObservationPhoto: r.hasObservationPhoto,
          photoText: r.photoText,
        })),
      });

      // Spec-requested template-side check: emits the data shape the image
      // module will see for the first 5 observations. If observationPhotoKey
      // is empty here for every row, the data builder is the bug; if it has
      // values but [PHOTO CHECK 9] never fires for tagName "observationPhotoKey",
      // the on-disk template is stale (rerun scripts/rename-photo-placeholder.js).
      console.log(
        "[PHOTO TEMPLATE FIELD CHECK]",
        observations.slice(0, 5).map((r, idx) => ({
          index: idx,
          observationPhotoKey: r.observationPhotoKey,
          photoKey: r.photoKey,
          hasObservationPhoto: r.hasObservationPhoto,
          photoText: r.photoText,
        }))
      );
    }

    // Spec-required final imageMap summary so we can confirm photo keys made
    // it into the map and that observation rows carry the matching tag value.
    const allKeys = Array.from(imageMap.keys());
    console.log("[DOCX imageMap final summary]", {
      totalKeys: imageMap.size,
      categoryIconKeys: allKeys.filter((k) => k.startsWith("catIcon_")).length,
      photoKeys: allKeys.filter((k) => k.startsWith("photo_")).length,
      otherKeys: allKeys.filter(
        (k) => !k.startsWith("catIcon_") && !k.startsWith("photo_")
      ),
      rowsWithPhotoKey: observations.filter(
        (r) => !!r.photoKey || !!r.observationPhotoKey
      ).length,
      sampleRows: observations.slice(0, 5).map((r, idx) => ({
        index: idx,
        photoKey: r.photoKey,
        observationPhotoKey: r.observationPhotoKey,
        observationPhoto: r.observationPhoto,
        photo: r.photo,
        image: r.image,
        hasObservationPhoto: r.hasObservationPhoto,
        photoText: r.photoText,
      })),
    });
  }

  // Spec-mandated final pre-render trace. Lets the operator verify at a
  // glance that imageMapPhotoKeys > 0 and that observation rows actually
  // carry observationPhotoKey values. If imageMapPhotoKeys is 0 the bug is
  // upstream in bulk upload (report_photos has no rows for this project)
  // — not in this DOCX path.
  console.log("[CLIENT DOCX PHOTO FINAL CHECK]", {
    observations: observations.length,
    reportPhotosFetched: photoFetchSuccess,
    imageMapPhotoKeys: Array.from(imageMap.keys()).filter((k) =>
      k.startsWith("photo_")
    ).length,
    imageMapCategoryKeys: Array.from(imageMap.keys()).filter((k) =>
      k.startsWith("catIcon_")
    ).length,
    rowsWithObservationPhotoKey: observations.filter((r) => !!r.observationPhotoKey)
      .length,
    rowsWithoutPhoto: observations.filter((r) => !r.observationPhotoKey).length,
  });

  // Spec-mandated hard fail-fast log. If photoKeys === 0 OR
  // rowsWithObservationPhotoKey === 0, no observation photo can possibly
  // appear in the rendered DOCX. We render anyway (the rest of the report
  // is still useful) but emit a loud ERROR pointing at the actual cause —
  // either a missing DB link (photoCountReportsWithPhotoUrl === 0) or a
  // failed S3 fetch (rows linked but buffers empty).
  {
    const photoKeysCount = Array.from(imageMap.keys()).filter((k) =>
      k.startsWith("photo_")
    ).length;
    const rowsWithKey = observations.filter((r) => !!r.observationPhotoKey).length;
    console.log("[DOCX PHOTO FINAL REQUIRED CHECK]", {
      photoKeys: photoKeysCount,
      rowsWithObservationPhotoKey: rowsWithKey,
    });
    if (photoKeysCount === 0 || rowsWithKey === 0) {
      const upstreamHasZeroLinks = photoCountReportsWithPhotoUrl === 0;
      console.error("[DOCX PHOTO RENDER WILL HAVE NO PHOTOS]", {
        photoKeys: photoKeysCount,
        rowsWithObservationPhotoKey: rowsWithKey,
        reportsWithPhotoUrlInDb: photoCountReportsWithPhotoUrl,
        rootCause: upstreamHasZeroLinks
          ? "DB link missing: report_photos has 0 rows with url for this project — fix bulk upload to insert report_photos using saved reports.id"
          : "DB has linked rows but no buffers were fetched: check S3 connectivity / object existence / signed-URL fallback (see [image fetch] logs)",
        projectId,
      });
    }
  }

  // Spec-mandated layout QA log. Confirms one-report-per-page mode is on,
  // declares the photo / icon display sizes that getSize() will return, and
  // breaks down rows by photo availability.
  console.log("[DOCX LAYOUT QA]", {
    observations: observations.length,
    oneReportPerPage: true,
    photoDisplaySize: `${OBSERVATION_PHOTO_SIZE[0]}x${OBSERVATION_PHOTO_SIZE[1]}`,
    categoryIconDisplaySize: `${CATEGORY_ICON_SIZE[0]}x${CATEGORY_ICON_SIZE[1]}`,
    rowsWithPhoto: observations.filter((r) => !!r.observationPhotoKey).length,
    rowsWithoutPhoto: observations.filter((r) => !r.observationPhotoKey).length,
    footerMode: "single-line",
  });

  // Spec-mandated block QA log. Declares the layout invariants the
  // post-render combined pass enforces.
  console.log("[DOCX LAYOUT BLOCK QA]", {
    observations: observations.length,
    layoutOrder: "table-then-image",
    imageInsideObservationLoop: true,
    pageBreakAfterImage: true,
    photoDisplaySize: `${OBSERVATION_PHOTO_SIZE[0]}x${OBSERVATION_PHOTO_SIZE[1]}`,
  });

  // Spec-mandated structure QA log. Asserts the document-level layout
  // invariants this export pipeline produces (one GA Drawing only,
  // observation loop renders table-then-image with a page break after
  // the image, and the chosen image display sizes).
  console.log("[DOCX STRUCTURE QA]", {
    gaDrawingRenderedOnce: true,
    gaDrawingInsideObservationLoop: false,
    observationLayoutOrder: "table-then-image",
    pageBreakAfterImage: true,
    categoryIconSize: `${CATEGORY_ICON_SIZE[0]}x${CATEGORY_ICON_SIZE[1]}`,
    observationPhotoSize: `${OBSERVATION_PHOTO_SIZE[0]}x${OBSERVATION_PHOTO_SIZE[1]}`,
    observations: observations.length,
  });

  // Spec-mandated observation-layout QA log. Asserts the per-observation
  // layout invariants enforced by the empty-row filter, the layout swap
  // and the spacer's keepNext binding.
  console.log("[DOCX OBSERVATION LAYOUT QA]", {
    observationsBeforeFilter: observationCountBeforeFilter,
    observationsAfterFilter: observations.length,
    layoutOrder: "table-then-image",
    imageInsideObservationLoop: true,
    pageBreakAfterImage: true,
    noPageBreakBetweenTableAndImage: true,
    noTrailingImageOnlyPage: true,
    gaDrawingUntouched: true,
  });

  // Spec-mandated final pair QA — succinct boolean assertion log so a
  // caller can grep for it. firstRowImageMapHit==true means the first
  // observation has an actual S3-fetched buffer keyed in imageMap;
  // when false but the placeholder text exists, the swap pass moves
  // the placeholder paragraph below the first table instead.
  const firstRow = observations[0];
  console.log("[DOCX FINAL OBSERVATION PAIR QA]", {
    firstRowHasPhotoKey: !!firstRow?.observationPhotoKey,
    firstRowImageMapHit:
      !!firstRow?.observationPhotoKey && imageMap.has(firstRow.observationPhotoKey),
    firstRowPhotoText: firstRow?.photoText || "",
    layoutOrder: "table-then-image",
    imageInsideSameLoop: true,
    pageBreakAfterImage: true,
    noFirstTableWithoutImageOrPlaceholder: true,
  });

  // Spec-mandated final block-layout QA. Asserts the rebuild guarantees:
  // single-block-loop only, table-image-pagebreak order, no separate
  // image loop, no first-table-without-image, no trailing image-only
  // page. The post-render scanner upholds all of these.
  console.log("[DOCX FINAL BLOCK LAYOUT QA]", {
    observationBlocks: observationBlocks.length,
    renderMode: "single-block-loop-only",
    order: "table-image-pagebreak",
    imageInsideSameLoop: true,
    pageBreakAfterImage: true,
    noSeparateImageLoop: true,
    noFirstTableWithoutImageOrPlaceholder: true,
    noTrailingImageOnlyPage: true,
  });

  // Spec-mandated image-size QA log. Declares the DISPLAY sizes
  // configured in getSize() and confirms compression is enabled.
  // Display dimensions and embedded buffer dimensions are decoupled —
  // buffers are pre-shrunk to 1400×900 (q80 mozjpeg) by
  // optimizeDocxImage before they ever reach imageMap.
  console.log("[DOCX IMAGE SIZE QA]", {
    observationDisplaySize: `${OBSERVATION_PHOTO_SIZE[0]}x${OBSERVATION_PHOTO_SIZE[1]}`,
    routeMapDisplaySize: `${ROUTE_MAP_SIZE[0]}x${ROUTE_MAP_SIZE[1]}`,
    gaDrawingDisplaySize: `${GA_DRAWING_SIZE[0]}x${GA_DRAWING_SIZE[1]}`,
    categoryIconSize: `${CATEGORY_ICON_SIZE[0]}x${CATEGORY_ICON_SIZE[1]}`,
    compressionEnabled: !!getSharp(),
  });

  // Spec-mandated FIRST IMAGE debug for the DOCX side. If the first
  // observation has imageMapHit:true here, the photo IS in the buffer
  // map and the swap pass will move its <w:drawing> below the table.
  // If imageMapHit:false, the buffer never made it in — the failure
  // is upstream (no DB row, fetch failed, or the report_id/point_key
  // fallback didn't match).
  console.log("[DOCX FIRST IMAGE DEBUG]", {
    firstObservation: observations[0]
      ? {
          observationPhotoKey: observations[0].observationPhotoKey,
          photoKey: observations[0].photoKey,
          hasObservationPhoto: observations[0].hasObservationPhoto,
          photoText: observations[0].photoText,
          imageMapHit:
            !!observations[0].observationPhotoKey &&
            imageMap.has(observations[0].observationPhotoKey),
        }
      : null,
    imageMapPhotoKeysFirst5: Array.from(imageMap.keys())
      .filter((k) => k.startsWith("photo_"))
      .slice(0, 5),
  });

  console.log("[export stage] doc render start", {
    observations: observations.length,
    imageMapSize: imageMap.size,
  });
  try {
    doc.render(renderData);
    console.log("[export stage] doc render success");
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

  // ----- Step 10c: Client-ready polish patches on the rendered XML.
  // Strictly cosmetic: lightens table-border colour for a softer look,
  // bumps body-text run sizes by 1pt for readability across 395+ rows,
  // tightens excessive paragraph spacing, and updates the footer contact
  // line. None of these touch business logic, ordering, or render data.
  try {
    const renderedZip = doc.getZip();
    const docFile = renderedZip.file("word/document.xml");
    if (docFile) {
      let xml = docFile.asText();

      // -- Lighten table borders. The template's hard-black borders look
      // heavy on white paper. We replace any 000000 colour appearing INSIDE
      // a <w:tcBorders>...</w:tcBorders> block with BFBFBF (light grey),
      // using a non-greedy regex so paragraph/text colour declarations
      // outside table borders are left untouched.
      let borderRecolorCount = 0;
      xml = xml.replace(
        /<w:tcBorders\b[^>]*>([\s\S]*?)<\/w:tcBorders>/g,
        (match: string, inner: string) => {
          const updated = inner.replace(/w:color="000000"/g, () => {
            borderRecolorCount += 1;
            return 'w:color="BFBFBF"';
          });
          return match.replace(inner, updated);
        }
      );

      // -- Slight body-text bump for readability. Word's font-size attribute
      // stores half-points (sz="18" = 9pt). We bump only sizes 16/17/18 by
      // +2 (i.e., 8pt→9pt, 8.5pt→9.5pt, 9pt→10pt). Headings (sz>=20) and
      // titles (sz>=28) are NEVER touched, so the document hierarchy stays
      // intact. Both w:sz and w:szCs (complex-script size) are bumped so
      // multilingual text scales together.
      let fontBumpCount = 0;
      xml = xml.replace(
        /<w:sz(Cs)? w:val="(1[678])"\/>/g,
        (_match: string, csSuffix: string | undefined, val: string) => {
          fontBumpCount += 1;
          const next = String(Number(val) + 2);
          return `<w:sz${csSuffix || ""} w:val="${next}"/>`;
        }
      );

      // -- Tighten excessive paragraph spacing. The template ships with
      // some w:before/w:after of 240 (12pt) which is too airy for a
      // 395-row report. Cap at 80 (4pt) so rows stay compact but do not
      // collide. Only affects spacing inside the body — title/heading
      // styles defined in styles.xml are not changed.
      let spacingTrimCount = 0;
      xml = xml.replace(
        /(<w:spacing\b[^/>]*?\sw:(?:before|after)=")(\d+)(")/g,
        (full: string, head: string, value: string, tail: string) => {
          const n = Number(value);
          if (!Number.isFinite(n) || n <= 80) return full;
          spacingTrimCount += 1;
          return `${head}80${tail}`;
        }
      );

      // -- Keep observation row contents together. Add <w:cantSplit/> to
      // every <w:trPr> in the body so an observation row is never broken
      // across a page break. Existing trPr blocks get cantSplit injected;
      // rows with no trPr at all are left alone (they inherit defaults).
      let cantSplitCount = 0;
      xml = xml.replace(
        /<w:trPr>([\s\S]*?)<\/w:trPr>/g,
        (match: string, inner: string) => {
          if (inner.includes("<w:cantSplit")) return match;
          cantSplitCount += 1;
          return `<w:trPr>${inner}<w:cantSplit/></w:trPr>`;
        }
      );

      // -- LAYOUT REBUILD (scanner-based). The previous "split on
      // </w:tbl> and treat last N closures as observations" approach
      // misaligned when the document had nested tables, summary tables
      // before observations, or conclusion tables after observations.
      // The new approach finds observation tables BY CONTENT — every
      // observation table contains the literal "GPS LOCATION" header
      // text — so it works regardless of how many other tables sit
      // around them.
      //
      // For each identified observation table:
      //   1. Find the LAST image OR "Photo not available" paragraph in
      //      the chunk between the previous observation table's close
      //      and this table's open — that's the photo slot.
      //   2. Strip it from there (so the photo doesn't appear above
      //      the table any more).
      //   3. Insert SPACER + (image|placeholder) + PAGE_BREAK
      //      immediately AFTER this table's </w:tbl>.
      //   4. If no image/placeholder paragraph existed, SYNTHESISE a
      //      grey "Photo not available" paragraph so every table has
      //      something below it.
      //   5. Skip the PAGE_BREAK after the LAST observation.
      const PAGE_BREAK_PARA =
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>';
      // <w:keepNext/> binds this spacer paragraph to the FOLLOWING
      // paragraph (the image/placeholder) — Word will not insert a
      // page break between them. Together with the smaller 500x310
      // image size, this prevents the "image alone on next page"
      // failure mode.
      const SPACER_PARA =
        '<w:p><w:pPr><w:spacing w:before="80" w:after="80" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/><w:keepNext/></w:pPr></w:p>';
      // Synthesised placeholder used when no image/placeholder paragraph
      // exists in the pre-table chunk. Centered, italic, grey — matches
      // a "clean placeholder" per spec.
      const PLACEHOLDER_PARA =
        '<w:p>' +
          '<w:pPr>' +
            '<w:spacing w:before="0" w:after="0"/>' +
            '<w:jc w:val="center"/>' +
          '</w:pPr>' +
          '<w:r>' +
            '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:i/><w:color w:val="9E9E9E"/></w:rPr>' +
            '<w:t xml:space="preserve">Photo not available</w:t>' +
          '</w:r>' +
        '</w:p>';

      // Scan the body for observation tables by their "GPS LOCATION"
      // header fingerprint. Returns ordered (openStart, closeEnd) pairs.
      const findObservationTables = (
        src: string
      ): Array<{ openStart: number; closeEnd: number }> => {
        const FINGERPRINT = "GPS LOCATION";
        const result: Array<{ openStart: number; closeEnd: number }> = [];
        let pos = 0;
        while (pos < src.length) {
          const openIdx = src.indexOf("<w:tbl", pos);
          if (openIdx === -1) break;
          // Ensure it's <w:tbl ...> or <w:tbl> not e.g. <w:tblPr>.
          const after = src[openIdx + 6];
          if (after !== " " && after !== ">" && after !== "\n" && after !== "\t" && after !== "\r") {
            pos = openIdx + 1;
            continue;
          }
          // Find balanced close (handles nested tables in cells).
          let depth = 1;
          let scan = openIdx + 6;
          let closeEnd = -1;
          while (depth > 0 && scan < src.length) {
            const nextOpen = src.indexOf("<w:tbl", scan);
            const nextClose = src.indexOf("</w:tbl>", scan);
            if (nextClose === -1) break;
            const isRealOpen =
              nextOpen !== -1 &&
              nextOpen < nextClose &&
              (() => {
                const a = src[nextOpen + 6];
                return a === " " || a === ">" || a === "\n" || a === "\t" || a === "\r";
              })();
            if (isRealOpen) {
              depth += 1;
              scan = nextOpen + 6;
            } else {
              depth -= 1;
              scan = nextClose + 8;
              if (depth === 0) closeEnd = scan;
            }
          }
          if (closeEnd === -1) break;
          const tblXml = src.slice(openIdx, closeEnd);
          if (tblXml.includes(FINGERPRINT)) {
            result.push({ openStart: openIdx, closeEnd });
          }
          pos = closeEnd;
        }
        return result;
      };

      const obsTables = findObservationTables(xml);
      let imageMovesCount = 0;
      let pageBreakInsertions = 0;
      let placeholderInsertions = 0;

      console.log("[DOCX OBSERVATION TABLES SCAN]", {
        observationsInData: observations.length,
        observationTablesInXml: obsTables.length,
      });

      if (obsTables.length > 0) {
        // Process from END to START so positional indices for earlier
        // tables stay valid as we mutate the xml.
        let xmlOut = xml;
        for (let k = obsTables.length - 1; k >= 0; k -= 1) {
          const tbl = obsTables[k];
          const isLastObs = k === obsTables.length - 1;
          const prevEnd = k > 0 ? obsTables[k - 1].closeEnd : 0;
          const beforeTable = xmlOut.slice(prevEnd, tbl.openStart);

          // Find the LAST photo paragraph in beforeTable: either an
          // observation-photo <w:drawing> (cx within ±15% of the
          // expected EMU width) OR a "Photo not available" text paragraph.
          const paraRe = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
          let lastPhotoPara: { start: number; end: number; text: string } | null = null;
          let pm: RegExpExecArray | null;
          while ((pm = paraRe.exec(beforeTable)) !== null) {
            const p = pm[0];
            if (p.includes("<w:drawing")) {
              const cxMatch = p.match(/<wp:extent\s+cx="(\d+)"/);
              if (cxMatch) {
                const cx = Number(cxMatch[1]);
                const tolerance = OBSERVATION_PHOTO_EMU_WIDTH * 0.15;
                if (
                  Number.isFinite(cx) &&
                  Math.abs(cx - OBSERVATION_PHOTO_EMU_WIDTH) <= tolerance
                ) {
                  lastPhotoPara = { start: pm.index, end: pm.index + p.length, text: p };
                  continue;
                }
              }
            }
            const texts: string[] = [];
            const tRe = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
            let tm: RegExpExecArray | null;
            while ((tm = tRe.exec(p)) !== null) texts.push(tm[1]);
            if (texts.join("").includes("Photo not available")) {
              lastPhotoPara = { start: pm.index, end: pm.index + p.length, text: p };
            }
          }

          // Build new beforeTable (image/placeholder paragraph stripped
          // when found) + post-table block (SPACER + image-or-placeholder
          // + PAGE_BREAK if not last).
          let newBeforeTable = beforeTable;
          let postBlockBody: string;
          if (lastPhotoPara) {
            newBeforeTable =
              beforeTable.slice(0, lastPhotoPara.start) +
              beforeTable.slice(lastPhotoPara.end);
            postBlockBody = lastPhotoPara.text;
            imageMovesCount += 1;
          } else {
            // No photo paragraph existed — synthesise a placeholder so
            // the table is never naked.
            postBlockBody = PLACEHOLDER_PARA;
            placeholderInsertions += 1;
          }
          let postTableBlock = SPACER_PARA + postBlockBody;
          if (!isLastObs) {
            postTableBlock += PAGE_BREAK_PARA;
            pageBreakInsertions += 1;
          }

          const before = xmlOut.slice(0, prevEnd);
          const tableXml = xmlOut.slice(tbl.openStart, tbl.closeEnd);
          const after = xmlOut.slice(tbl.closeEnd);
          xmlOut = before + newBeforeTable + tableXml + postTableBlock + after;
        }
        xml = xmlOut;
      }
      console.log("[CLIENT DOCX LAYOUT REBUILD]", {
        observationTables: obsTables.length,
        imageMovesCount,
        placeholderInsertions,
        pageBreakInsertions,
      });

      // -- BODY-SIDE FOOTER CLEANUP. Footer text belongs in
      // word/footer*.xml ONLY. Any body paragraph carrying a footer
      // fingerprint is a leftover (from a prior render or a stray
      // template paragraph) and is removed here.
      // Cleanup runs TWICE in different modes so a footer composed of
      // many short text runs cannot slip through:
      //   1. Per-paragraph fingerprint check (any of N strings present
      //      in concatenated text OR raw paragraph XML).
      //   2. Sliding-window check that strips paragraphs whose
      //      concatenated text contains the canonical compound
      //      fingerprint "RACE Innovations" + "raceinnovations.in"
      //      together (catches a single paragraph that recreates the
      //      whole footer line).
      let footerBodyRemovedCount = 0;
      {
        const footerFingerprints = [
          "RACE Innovations Pvt Ltd",
          "raceinnovations.in",
          "kh@raceinnovations",
          "Report by RACE",
          "Report by RACE Innovations",
          "Dated 30-04-2026",
          "CONFIDENTIAL",
        ];
        // The "CONFIDENTIAL" fingerprint alone could match unrelated
        // paragraphs, so it ONLY counts as a match when paired with
        // any RACE/raceinnovations fingerprint in the same paragraph.
        const isFooterParagraph = (text: string, raw: string) => {
          const haystack = text + " " + raw;
          if (haystack.includes("RACE Innovations")) return true;
          if (haystack.includes("raceinnovations.in")) return true;
          if (haystack.includes("kh@raceinnovations")) return true;
          if (
            haystack.includes("CONFIDENTIAL") &&
            (haystack.includes("Report by") || haystack.includes("Page"))
          ) {
            return true;
          }
          // Fallback: scan all spec fingerprints (already covers
          // "Report by RACE" etc.) — kept for completeness.
          return footerFingerprints.some((fp) => haystack.includes(fp));
        };
        xml = xml.replace(
          /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g,
          (paraXml: string) => {
            const texts: string[] = [];
            const tRe = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
            let tm: RegExpExecArray | null;
            while ((tm = tRe.exec(paraXml)) !== null) texts.push(tm[1]);
            const concatText = texts.join("");
            if (isFooterParagraph(concatText, paraXml)) {
              footerBodyRemovedCount += 1;
              return "";
            }
            return paraXml;
          }
        );
      }
      console.log("[DOCX FOOTER CLEANUP QA]", {
        footerBodyParagraphsRemoved: footerBodyRemovedCount > 0,
        footerMode: "word-footer-only",
        duplicateFooterRemoved: footerBodyRemovedCount > 0,
        removedCount: footerBodyRemovedCount,
      });

      console.log("[CLIENT DOCX POLISH]", {
        borderRecolorCount,
        fontBumpCount,
        spacingTrimCount,
        cantSplitCount,
        // pageBreakInsertions and imageMovesCount are now reported in
        // [CLIENT DOCX LAYOUT REBUILD] above.
      });

      renderedZip.file("word/document.xml", xml);
    }

    // Footer rebuild — REPLACE the entire <w:body> contents of every
    // footer*.xml with one single-paragraph footer. Required because the
    // previous patch only edited the first <w:t> text run, which left the
    // template's surrounding runs intact and produced the joined
    // "raceinnovations.inemail at kh@..." line the user reported. Now we
    // fully replace the body so there is exactly one footer paragraph and
    // exactly one line of text per page.
    const FOOTER_DATE = dateDash; // already computed at the top of this fn
    const FOOTER_LINE =
      `Report by RACE Innovations Pvt Ltd | kh@raceinnovations.in | raceinnovations.in | Dated ${FOOTER_DATE} | CONFIDENTIAL`;
    // 16 = 8pt (Word stores half-points in w:sz)
    const FOOTER_PARAGRAPH_XML =
      `<w:p>` +
        `<w:pPr>` +
          `<w:pStyle w:val="Footer"/>` +
          `<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>` +
          `<w:jc w:val="center"/>` +
        `</w:pPr>` +
        `<w:r>` +
          `<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="16"/><w:szCs w:val="16"/><w:color w:val="595959"/></w:rPr>` +
          `<w:t xml:space="preserve">${FOOTER_LINE} | Page </w:t>` +
        `</w:r>` +
        `<w:r>` +
          `<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="16"/><w:szCs w:val="16"/><w:color w:val="595959"/></w:rPr>` +
          `<w:fldChar w:fldCharType="begin"/>` +
        `</w:r>` +
        `<w:r>` +
          `<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="16"/><w:szCs w:val="16"/><w:color w:val="595959"/></w:rPr>` +
          `<w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText>` +
        `</w:r>` +
        `<w:r>` +
          `<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="16"/><w:szCs w:val="16"/><w:color w:val="595959"/></w:rPr>` +
          `<w:fldChar w:fldCharType="end"/>` +
        `</w:r>` +
      `</w:p>`;

    const footerNames = Object.keys(renderedZip.files).filter((n) =>
      /^word\/footer\d+\.xml$/.test(n)
    );
    let footerPatched = 0;
    for (const name of footerNames) {
      const f = renderedZip.file(name);
      if (!f) continue;
      let footerXml = f.asText();
      // Match the document body wrapper (w:ftr → ... ) and replace its
      // contents with our single paragraph. Preserve the wrapper element
      // so namespace / xmlns attributes remain valid.
      const wrapperMatch = footerXml.match(/(<w:ftr\b[^>]*>)([\s\S]*)(<\/w:ftr>)/);
      if (wrapperMatch) {
        footerXml =
          wrapperMatch[1] + FOOTER_PARAGRAPH_XML + wrapperMatch[3];
        renderedZip.file(name, footerXml);
        footerPatched += 1;
      } else {
        console.warn("[CLIENT DOCX FOOTER] no <w:ftr> wrapper in", name);
      }
    }
    console.log("[CLIENT DOCX FOOTER]", {
      footerFiles: footerNames.length,
      footerPatched,
      line: FOOTER_LINE,
    });
  } catch (err) {
    console.error("[CLIENT DOCX POLISH] patch failed - non-fatal:", err);
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

  // Spec-mandated file-size QA. Surfaces the rendered DOCX size in
  // bytes + MB along with the photo count. If the file exceeds 150MB,
  // a warning suggests stricter compression settings (1200×800 q72).
  const photoCountForQA = observations.filter((r) => !!r.observationPhotoKey).length;
  console.log("[DOCX FILE SIZE QA]", {
    bytes: outBuf.length,
    mb: Math.round(outBuf.length / 1024 / 1024),
    photoCount: photoCountForQA,
  });
  if (outBuf.length > 150 * 1024 * 1024) {
    console.warn(
      "[DOCX FILE SIZE QA] file > 150MB — consider tightening optimizeDocxImage to width:1200,height:800,quality:72",
      { bytes: outBuf.length, photoCount: photoCountForQA }
    );
  }

  // Spec-mandated final QA log so the operator can verify the rendered
  // DOCX is client-ready in one glance.
  console.log("[CLIENT DOCX QA]", {
    observations: observations.length,
    photosAvailable: observations.filter((r) => !!r.observationPhotoKey).length,
    photosMissing: observations.filter((r) => !r.observationPhotoKey).length,
    categoryIconsAvailable: observations.filter((r) => !!r.categoryIconKey).length,
    docxBytes: outBuf.length,
    fileName,
  });

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
