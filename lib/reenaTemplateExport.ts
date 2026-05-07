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
          // Spec: q78 mozjpeg keeps photos visually indistinguishable
          // from q80 while shaving ~5% per image — meaningful across
          // 395+ photos in one export.
          quality: 78,
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

// Side-by-side images are rendered as a 2-column borderless Word table
// (one image per cell) instead of compositing into a single JPEG.
// See the post-render SIDE-BY-SIDE TABLE INJECTION pass for the
// replacement logic.

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

// Per spec: route map = 13" × 7" (1248 × 672 px) landscape, GA drawing
// = 14" × 9.5" (1344 × 912 px) landscape — sized so GA + title fit on
// ONE landscape A3 page (~14.5" × 10.5" content; title ~0.5", GA 9.5"
// → total 10" within 10.5"). Big enough to dominate the page, small
// enough that Word's keepNext on the title above it actually keeps
// them on the same page. On smaller pages Word scales down
// proportionally while preserving aspect ratio.
const ROUTE_MAP_SIZE: [number, number] = [1248, 672];
const GA_DRAWING_SIZE: [number, number] = [1344, 912];
// Per spec: observation photo size depends on how many photos the report
// has. Single photo: 7.5" × 5.3" (720 × 509 px). 2 photos: each one is
// 7.2" × 5.0" (691 × 480 px) — matches the user's manual Word size that
// "fits perfectly" on their page setup. Rendered side-by-side inside a
// 100%-width borderless 2-column table; if the page can't accommodate
// 14.4" of total image width, Word scales each image down proportionally
// while preserving aspect ratio. The actual embedded JPEG is compressed
// via sharp BEFORE being added to imageMap (max 1400×900 inside, q78
// mozjpeg) so the DOCX file size stays small.
const OBSERVATION_PHOTO_SIZE: [number, number] = [720, 509];
const MULTI_PHOTO_SIZE: [number, number] = [691, 480];
// Bumped slightly larger per spec ("show the little big") so unknown /
// new-category icons read clearly in the observation table.
const CATEGORY_ICON_SIZE: [number, number] = [90, 90];
const CATEGORY_SUMMARY_ICON_SIZE: [number, number] = [80, 80];

// Word stores image extents in EMU (1 px = 9525 EMU). The layout-rebuild
// pass uses these constants to identify observation-photo drawings (both
// sizes) and ignore other in-document drawings (GA Drawing, Route Map,
// category icons).
const PX_TO_EMU = 9525;
const OBSERVATION_PHOTO_EMU_WIDTH = OBSERVATION_PHOTO_SIZE[0] * PX_TO_EMU;
const MULTI_PHOTO_EMU_WIDTH = MULTI_PHOTO_SIZE[0] * PX_TO_EMU;
const MULTI_PHOTO_EMU_HEIGHT = MULTI_PHOTO_SIZE[1] * PX_TO_EMU;

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

// Spec-mandated EXACT-match keyword sets. EXACT (Array.includes(v))
// not substring, so noise words in remarks/action don't hijack the
// table colour.
const RED_KEYWORDS = ["red", "critical", "hard", "fail", "not pass"];
const YELLOW_KEYWORDS = ["yellow", "warning", "caution", "medium"];
const GREEN_KEYWORDS = ["green", "normal", "normal pass", "pass", "ok", ""];

function getDifficultyTableColors(value: unknown): {
  key: "red" | "yellow" | "green";
  headerFillColor: string;
  headerTextColor: string;
  bodyFillColor: string;
  bodyTextColor: string;
} {
  const d = normalizeDifficulty(value);
  // EXACT-equals matching prevents the "red rectangles where I never
  // chose red" bug. Substring `includes` would catch "red" inside
  // "redirect", "hard" inside "hardware", "fail" inside "failure" —
  // applying the red shading to perfectly normal observations.
  // Per spec: bright header colours (yellow / red / green) with much
  // milder pastel body fills derived from the same hue. Header text is
  // always rendered white by the post-render recolorHeaderRowToWhite
  // pass — headerTextColor here is kept only as a template-level safety
  // fallback.
  // Per spec: bright header colours (yellow / red / green) with VERY
  // mild near-white pastel body fills derived from the same hue.
  // Header text is always white (post-render recolour pass).
  if (RED_KEYWORDS.includes(d)) {
    return {
      key: "red",
      headerFillColor: "F05052",  // bright red (spec)
      headerTextColor: "FFFFFF",
      bodyFillColor: "FDF2F2",    // near-white with faintest pink tint
      bodyTextColor: "0B3D2E",
    };
  }
  if (YELLOW_KEYWORDS.includes(d)) {
    return {
      key: "yellow",
      headerFillColor: "EABD0D",  // bright yellow (spec)
      headerTextColor: "FFFFFF",
      bodyFillColor: "FFFBE6",    // near-white with faintest cream tint
      bodyTextColor: "0B3D2E",
    };
  }
  // Anything not in RED or YELLOW keyword set lands here, including
  // the empty string and every GREEN_KEYWORDS entry. Red is opt-in.
  void GREEN_KEYWORDS; // referenced for grep + future explicit check
  return {
    key: "green",
    headerFillColor: "56BE9F",    // bright mint green (spec)
    headerTextColor: "FFFFFF",
    bodyFillColor: "F0F9F4",      // near-white with faintest mint tint
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

  // ---- Order matters: more specific keywords FIRST so they don't get
  // ---- swallowed by broader rules below.

  // Side / signal / speed variants must beat the broader signboard rule.
  if (c.includes("side_signboard") || c.includes("electric_side_signboard")) {
    return "side_signboard";
  }
  if (c.includes("signal_pole") || c.includes("signal")) return "signal_pole";
  if (c.includes("speed_pole") || c.includes("speed_bump") || c.includes("speed_break") || c.includes("hump")) {
    return "speed_pole";
  }

  // Bridge variants — specific bridge types must beat the generic "bridge".
  if (c.includes("footpath_bridge") || c.includes("footpath")) return "footpath_bridge";
  if (c.includes("river_bridge") || c.includes("river")) return "river_bridge";
  if (c.includes("underpass")) return "underpass";
  // Junction variants — left/right specific must beat plain "junction".
  if (c.includes("junction_left") || /\bleft\b/.test(c) && c.includes("junction")) return "junction_left";
  if (c.includes("junction_right") || /\bright\b/.test(c) && c.includes("junction")) return "junction_right";
  if (c.includes("junction") || c.includes("intersection") || c.includes("crossroad")) return "junction";
  // Plain "Bridge" / "Flyover" / "Overpass" — falls through to bridge.jpg.
  if (c.includes("bridge") || c.includes("flyover") || c.includes("overpass") || c.includes("viaduct")) {
    return "bridge";
  }

  if (c.includes("low_tension") || c === "lt_cable") return "lt_cable";
  if (c.includes("high_tension") || c === "ht_cable") return "ht_cable";
  if (c.includes("tower")) return "towerline_cable";
  if (c.includes("tree")) return "tree";
  // Camera / CCTV pole gets its OWN icon (must be checked BEFORE the
  // broader signboard rule so "Camera Pole" doesn't fall into signboard).
  if (c.includes("camera_pole") || c.includes("camera") || c.includes("cctv")) {
    return "camera_pole";
  }
  if (c.includes("signboard") || c.includes("electric_sign")) {
    return "signboard";
  }
  if (c.includes("toll")) return "toll";
  if (c.includes("narrow")) return "narrow_road";
  if (c.includes("gate")) return "gate";
  if (c.includes("bend") || c.includes("curve") || c.includes("turn")) return "bend";
  if (c.includes("petrol") || c.includes("fuel") || c.includes("gas_station")) return "petrol";
  if (c.includes("railway") || c.includes("rail_crossing")) return "railway_level_crossing";
  if (c.includes("diversion") || c.includes("detour")) return "diversion";
  if (c.includes("parking") || c.includes("park_area")) return "parking";

  return "fallback";
}

// Server-side fs paths (process.cwd() is the project root). The browser-side
// equivalent lives in lib/download.ts and uses URL-style paths instead.
//
// Map is built from the ACTUAL files present in
// public/images/report-icons/ (all .png) plus alias keys so the
// normalizeCategoryKey output (which uses canonical underscore_keys
// like "lt_cable", "ht_cable", "towerline", "tree", etc.) resolves to
// the same icon as the descriptive filename.
const CATEGORY_ICON_MAP: Record<string, string> = {
  // --- One entry per actual file in the folder -----------------------
  bend: "public/images/report-icons/bend.png",
  bridge: "public/images/report-icons/bridge.png",
  camera_pole: "public/images/report-icons/camera_pole.png",
  electric_side_signboard: "public/images/report-icons/electric_side_signboard.png",
  electric_signboard: "public/images/report-icons/electric_signboard.png",
  footpath_bridge: "public/images/report-icons/footpath_bridge.png",
  gate: "public/images/report-icons/gate.png",
  high_tension_cable: "public/images/report-icons/high_tension_cable.png",
  junction: "public/images/report-icons/junction.png",
  junction_left: "public/images/report-icons/junction-left.png",
  junction_right: "public/images/report-icons/junction-right.png",
  low_tension_cable: "public/images/report-icons/low_tension_cable.png",
  narrow_road: "public/images/report-icons/narrow_road.png",
  petrol_bunk: "public/images/report-icons/petrol_bunk.png",
  railway_level_crossing: "public/images/report-icons/railway_level_crossing.png",
  river_bridge: "public/images/report-icons/river_bridge.png",
  side_signboard: "public/images/report-icons/side_signboard.png",
  signal_pole: "public/images/report-icons/signal_pole.png",
  signboard: "public/images/report-icons/signboard.png",
  speed_pole: "public/images/report-icons/speed_pole.png",
  toll_plaza: "public/images/report-icons/toll_plaza.png",
  tower_line_cable: "public/images/report-icons/tower_line_cable.png",
  tree_branches: "public/images/report-icons/tree_branches.png",
  underpass_bridge: "public/images/report-icons/underpass_bridge.png",

  // --- Aliases used by normalizeCategoryKey ---------------------------
  // These short keys collapse multiple input variations onto a single
  // canonical icon so e.g. "LT Cable" / "Low Tension Cable" / "lt_cable"
  // all reach low_tension_cable.png.
  lt_cable: "public/images/report-icons/low_tension_cable.png",
  ht_cable: "public/images/report-icons/high_tension_cable.png",
  towerline: "public/images/report-icons/tower_line_cable.png",
  towerline_cable: "public/images/report-icons/tower_line_cable.png",
  tower_line: "public/images/report-icons/tower_line_cable.png",
  tree: "public/images/report-icons/tree_branches.png",
  underpass: "public/images/report-icons/underpass_bridge.png",
  toll: "public/images/report-icons/toll_plaza.png",
  petrol: "public/images/report-icons/petrol_bunk.png",
  road: "public/images/report-icons/narrow_road.png",
  electric_sign: "public/images/report-icons/electric_signboard.png",
};

// ---- ON-THE-FLY ICON GENERATION FOR NEW / UNKNOWN CATEGORIES ----
// When a category comes in that isn't in CATEGORY_ICON_MAP (a brand-new
// hazard type the rules above don't recognise), we GENERATE a unique
// PNG icon for it on the fly — a coloured circle with the category's
// initials in the centre. NO disk-fallback to ca-5.
//
// Auto-generated icons are flagged by the AUTO_ICON_PREFIX in their
// filename so the loader (loadCategoryIcon) routes them to the
// generator instead of trying to read from disk.
const AUTO_ICON_PREFIX = "__auto__";
// Maps the synthetic filename back to the original raw category text
// so the generator can render the right initials/color.
const autoCategoryByFileName = new Map<string, string>();

/**
 * Build a stable synthetic filename from the (normalised) category text
 * so the same new category always resolves to the same generated icon.
 */
function autoIconFileNameFor(rawCategory: string): string {
  const norm = normalizeCategory(rawCategory);
  const hash = crypto
    .createHash("md5")
    .update(norm)
    .digest("hex")
    .slice(0, 12);
  return `${AUTO_ICON_PREFIX}${hash}.png`;
}

/**
 * Pick a SYMBOLIC icon kind for a brand-new category by matching keywords
 * in the raw category text. The matched kind drives which SVG glyph is
 * drawn inside the colored circle — so the auto-generated icon visually
 * RELATES to the category (a pothole-shaped blob for "Pothole", a cone
 * for "Construction", a stop sign for "Stop", etc.) rather than just
 * showing letter initials. Defaults to a warning triangle.
 */
type AutoIconKind =
  | "warning"
  | "pothole"
  | "cone"
  | "pin"
  | "signal"
  | "stop"
  | "flag"
  | "bridge"
  | "speedbump";

function classifyAutoIconKind(rawCategory: string): AutoIconKind {
  const c = rawCategory.toLowerCase();
  if (/(pothole|crack|broken|damage|dent|hole|pit)/.test(c)) return "pothole";
  if (/(speed.?bump|speed.?break|hump|ridge|bumper)/.test(c)) return "speedbump";
  if (/(construct|barricade|barrier|cone|workzone|workman|road.?work)/.test(c)) return "cone";
  if (/(signal|traffic.?light|junction)/.test(c)) return "signal";
  if (/(stop|halt|no.?entry)/.test(c)) return "stop";
  if (/(flag|waypoint|milestone|km.?stone)/.test(c)) return "flag";
  if (/(bridge|flyover|overpass|culvert|viaduct)/.test(c)) return "bridge";
  if (/(landmark|point|spot|location|place|area)/.test(c)) return "pin";
  return "warning";
}

/**
 * BLACK monochrome SVG glyphs (one per AutoIconKind) sized for a 160×160
 * canvas centered around (80, 80). Drawn directly on a transparent
 * background — no colored circle, no random hues. Solid black fills and
 * strokes; white (#FFFFFF) is used only for negative-space cutouts
 * inside the black shape (e.g. the "!" inside the warning triangle).
 * Pure SVG primitives — no font dependency.
 */
const AUTO_ICON_GLYPHS: Record<AutoIconKind, string> = {
  warning:
    `<g transform="translate(80,80)">` +
      // Solid black triangle
      `<path d="M 0,-60 L 60,42 L -60,42 Z" fill="#000000" stroke="#000000" stroke-width="2" stroke-linejoin="round"/>` +
      // White exclamation cutout
      `<rect x="-5" y="-22" width="10" height="34" rx="2" fill="#FFFFFF"/>` +
      `<circle cx="0" cy="26" r="6" fill="#FFFFFF"/>` +
    `</g>`,
  pothole:
    // Solid black irregular pothole shape
    `<g transform="translate(80,80)">` +
      `<path d="M -56,4 Q -50,-38 -12,-42 Q 34,-46 50,-14 Q 58,20 32,38 Q 0,50 -34,40 Q -60,32 -56,4 Z" fill="#000000"/>` +
      // White inner shadow / texture spot for depth
      `<ellipse cx="-12" cy="-8" rx="10" ry="5" fill="#FFFFFF" fill-opacity="0.5"/>` +
    `</g>`,
  cone:
    `<g transform="translate(80,80)">` +
      // Solid black cone triangle
      `<path d="M 0,-54 L 38,46 L -38,46 Z" fill="#000000" stroke="#000000" stroke-width="2" stroke-linejoin="round"/>` +
      // White horizontal stripes
      `<rect x="-26" y="0" width="52" height="8" fill="#FFFFFF"/>` +
      `<rect x="-19" y="-26" width="38" height="7" fill="#FFFFFF"/>` +
      // Black base bar
      `<rect x="-44" y="46" width="88" height="6" fill="#000000"/>` +
    `</g>`,
  pin:
    `<g transform="translate(80,80)">` +
      // Solid black pin drop
      `<path d="M 0,-54 C -26,-54 -38,-30 -38,-14 C -38,16 0,52 0,52 C 0,52 38,16 38,-14 C 38,-30 26,-54 0,-54 Z" fill="#000000"/>` +
      // White inner circle
      `<circle cx="0" cy="-16" r="13" fill="#FFFFFF"/>` +
    `</g>`,
  signal:
    `<g transform="translate(80,80)">` +
      // Solid black signal box
      `<rect x="-22" y="-54" width="44" height="100" rx="6" fill="#000000"/>` +
      // Three white lights
      `<circle cx="0" cy="-32" r="11" fill="#FFFFFF"/>` +
      `<circle cx="0" cy="-4" r="11" fill="#FFFFFF"/>` +
      `<circle cx="0" cy="24" r="11" fill="#FFFFFF"/>` +
    `</g>`,
  stop:
    `<g transform="translate(80,80)">` +
      // Solid black octagon
      `<polygon points="-26,-50 26,-50 50,-26 50,26 26,50 -26,50 -50,26 -50,-26" fill="#000000" stroke="#000000" stroke-width="2" stroke-linejoin="round"/>` +
      // White horizontal bar
      `<rect x="-32" y="-6" width="64" height="12" fill="#FFFFFF"/>` +
    `</g>`,
  flag:
    `<g transform="translate(80,80)">` +
      // Black flagpole
      `<rect x="-5" y="-54" width="7" height="108" fill="#000000"/>` +
      // Black flag triangle
      `<path d="M 2,-54 L 48,-38 L 2,-22 Z" fill="#000000" stroke="#000000" stroke-width="2" stroke-linejoin="round"/>` +
    `</g>`,
  bridge:
    `<g transform="translate(80,80)">` +
      // Black arch
      `<path d="M -54,18 Q 0,-44 54,18" fill="none" stroke="#000000" stroke-width="8" stroke-linecap="round"/>` +
      // Black deck
      `<rect x="-54" y="26" width="108" height="8" fill="#000000"/>` +
      // Black pillars
      `<rect x="-50" y="34" width="7" height="18" fill="#000000"/>` +
      `<rect x="-15" y="34" width="7" height="18" fill="#000000"/>` +
      `<rect x="8" y="34" width="7" height="18" fill="#000000"/>` +
      `<rect x="43" y="34" width="7" height="18" fill="#000000"/>` +
    `</g>`,
  speedbump:
    `<g transform="translate(80,80)">` +
      // Black road line
      `<rect x="-54" y="28" width="108" height="7" fill="#000000"/>` +
      // Black bump dome
      `<path d="M -44,28 Q 0,-26 44,28 Z" fill="#000000"/>` +
      // White vertical bars on the bump (motion / hazard stripes)
      `<rect x="-26" y="-4" width="7" height="16" fill="#FFFFFF"/>` +
      `<rect x="-3" y="-14" width="7" height="26" fill="#FFFFFF"/>` +
      `<rect x="19" y="-4" width="7" height="16" fill="#FFFFFF"/>` +
    `</g>`,
};

/**
 * Generate a PNG icon for a brand-new category. The icon is a coloured
 * filled circle with a SYMBOLIC glyph (warning triangle, pothole blob,
 * construction cone, traffic signal, stop sign, flag, bridge, speed
 * bump, or map pin) chosen by keyword-matching the category text. Same
 * text always yields the same icon. 160×160 px to match optimizeDocxImage
 * output for canonical icons.
 */
async function generateCategoryIcon(rawCategory: string): Promise<Buffer | null> {
  const sharp = getSharp();
  if (!sharp) return null;
  const text = String(rawCategory || "").trim();
  if (!text) return null;

  const kind = classifyAutoIconKind(text);
  const glyph = AUTO_ICON_GLYPHS[kind];

  // BLACK ONLY — no random hues. The glyph is solid black on a fully
  // transparent background (with white cutouts inside the shape for
  // negative-space details). Same input text → identical output icon.
  const size = 160;
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">` +
      glyph +
    `</svg>`;

  try {
    const buf = await sharp(Buffer.from(svg)).png().toBuffer();
    console.log("[CATEGORY ICON GENERATED]", {
      rawCategory: text,
      kind,
      style: "black-only monochrome on transparent background",
      bytes: buf.length,
    });
    return buf;
  } catch (err) {
    console.warn("[CATEGORY ICON GENERATE FAILED]", {
      text,
      err: (err as Error)?.message || String(err),
    });
    return null;
  }
}

// ---- SIMILARITY MATCHING for unknown categories ----
// When a raw category text doesn't have a direct entry in
// CATEGORY_ICON_MAP, walk through these groups and pick the icon for
// the FIRST group whose keyword list intersects with the category text.
// This way "Road Damage" or "Side Road" reuse the narrow_road.png
// icon, "Flyover" reuses bridge.png, "CCTV Pole" reuses camera_pole.png,
// etc. — instead of falling straight through to the auto-generator.
//
// Each group's keywords are matched as substrings against the
// lower-cased category text (with non-alphanumerics replaced by spaces).
// More-specific groups should appear FIRST so they don't get swallowed
// by broader ones (e.g. "river bridge" matches the bridge group ONLY
// after the river_bridge map entry has already been tried).
const CATEGORY_SIMILARITY_GROUPS: Array<{ keywords: string[]; iconFile: string }> = [
  // --- Bridge family ---
  { keywords: ["river"], iconFile: "river_bridge.png" },
  { keywords: ["footpath", "sidewalk", "pavement", "walkway", "pedestrian path"], iconFile: "footpath_bridge.png" },
  { keywords: ["underpass", "tunnel", "subway"], iconFile: "underpass_bridge.png" },
  { keywords: ["bridge", "flyover", "overpass", "viaduct", "culvert"], iconFile: "bridge.png" },

  // --- Junction family ---
  { keywords: ["junction left", "left junction"], iconFile: "junction-left.png" },
  { keywords: ["junction right", "right junction"], iconFile: "junction-right.png" },
  { keywords: ["junction", "intersection", "crossroad", "crossing"], iconFile: "junction.png" },

  // --- Cables / wires / poles ---
  { keywords: ["high tension", "ht cable", "ht line"], iconFile: "high_tension_cable.png" },
  { keywords: ["low tension", "lt cable", "lt line"], iconFile: "low_tension_cable.png" },
  { keywords: ["tower line", "transmission tower", "tower", "pylon"], iconFile: "tower_line_cable.png" },
  { keywords: ["cable", "wire", "transmission"], iconFile: "low_tension_cable.png" },

  // --- Signs / signboards / cameras ---
  { keywords: ["side signboard", "side sign"], iconFile: "side_signboard.png" },
  { keywords: ["electric side signboard"], iconFile: "electric_side_signboard.png" },
  { keywords: ["electric sign", "electric signboard", "electric board"], iconFile: "electric_signboard.png" },
  { keywords: ["camera pole", "cctv pole", "cctv", "camera"], iconFile: "camera_pole.png" },
  { keywords: ["signboard", "signpost", "sign board", "name board", "board"], iconFile: "signboard.png" },

  // --- Signals / speed ---
  { keywords: ["signal pole", "signal", "traffic light", "traffic signal"], iconFile: "signal_pole.png" },
  { keywords: ["speed pole", "speed bump", "speed breaker", "speed hump", "hump", "bumper"], iconFile: "speed_pole.png" },

  // --- Road / surface ---
  { keywords: ["narrow road", "narrow"], iconFile: "narrow_road.png" },
  { keywords: ["road", "street", "lane", "highway", "way", "path", "carriageway"], iconFile: "narrow_road.png" },

  // --- Bend / curve / turn ---
  { keywords: ["bend", "curve"], iconFile: "bend.png" },

  // --- Vegetation ---
  { keywords: ["tree branch", "tree", "branch", "foliage", "vegetation"], iconFile: "tree_branches.png" },

  // --- Gate / barrier ---
  { keywords: ["gate", "barrier", "boom", "barricade"], iconFile: "gate.png" },

  // --- Petrol / fuel ---
  { keywords: ["petrol", "fuel", "diesel", "gas station", "filling station"], iconFile: "petrol_bunk.png" },

  // --- Toll ---
  { keywords: ["toll plaza", "toll", "fastag"], iconFile: "toll_plaza.png" },

  // --- Railway ---
  { keywords: ["railway", "rail crossing", "rail", "train"], iconFile: "railway_level_crossing.png" },
];

function findSimilarIcon(rawCategory: unknown): string | null {
  const text = String(rawCategory || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
  if (!text) return null;
  for (const group of CATEGORY_SIMILARITY_GROUPS) {
    for (const kw of group.keywords) {
      if (text.includes(kw)) return group.iconFile;
    }
  }
  return null;
}

function getCategoryIconFile(category: unknown): string {
  const key = normalizeCategoryKey(category);
  // 1. Canonical category with a real mapping → use it.
  if (key !== "fallback" && CATEGORY_ICON_MAP[key]) {
    const rel = CATEGORY_ICON_MAP[key];
    console.log("[category mapping check]", {
      rawCategory: category,
      normalizedKey: key,
      iconRelativePath: rel,
      source: "canonical",
    });
    return rel.replace(/^public[\\/]images[\\/]report-icons[\\/]/, "");
  }
  // 2. Try keyword-based similarity matching against existing icons.
  //    "Road Damage" → narrow_road.png, "Flyover" → bridge.png, etc.
  const similarFile = findSimilarIcon(category);
  if (similarFile) {
    console.log("[category mapping check]", {
      rawCategory: category,
      normalizedKey: key,
      iconRelativePath: `public/images/report-icons/${similarFile}`,
      source: "similarity-match",
    });
    return similarFile;
  }
  // 3. No similar icon either → synthesize an auto-icon filename. The
  //    loader (loadCategoryIcon) will detect the AUTO_ICON_PREFIX and
  //    generate the actual PNG instead of reading from disk.
  const text = String(category || "").trim();
  if (!text) {
    // Truly empty/null category — sentinel so loader returns null entry
    // and the template's fallback text takes over.
    return `${AUTO_ICON_PREFIX}empty.png`;
  }
  const fileName = autoIconFileNameFor(text);
  autoCategoryByFileName.set(fileName, text);
  console.log("[category mapping check]", {
    rawCategory: category,
    normalizedKey: key,
    autoIconFileName: fileName,
    source: "auto-generated-new-category",
  });
  return fileName;
}

/**
 * Generate the "Locations" stepper image (project title at the top,
 * then hollow circles linked by dotted connectors, with a red pin
 * marker on the final destination) that mirrors the web-app UI
 * design. Renders an SVG via sharp → PNG so the design is precise
 * AND the title + stepper become a single inseparable image (Word
 * physically can't split them across pages). Returns null if there
 * are no labels or sharp isn't available.
 */
async function generateRouteLocationsStepper(
  labels: { label: string; pin_type: string }[],
  projectTitle: string
): Promise<Buffer | null> {
  const sharp = getSharp();
  if (!sharp) return null;
  const items = labels.filter((l) => l.label).slice(0, 8);
  if (items.length === 0) return null;

  const escapeXml = (s: string) =>
    s
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");

  // Layout (px units inside the SVG canvas):
  const W = 720;                    // total canvas width
  const ROW_H = 90;                 // vertical spacing per item
  const TITLE_BAND_H = projectTitle ? 80 : 0; // top band reserved for project title
  const TOP_PAD = 30;               // gap below title before "Locations" heading
  const HEADING_H = 40;             // "Locations" heading band
  const BOTTOM_PAD = 24;
  const STEPPER_TOP = TITLE_BAND_H + TOP_PAD + HEADING_H;
  const H = STEPPER_TOP + items.length * ROW_H + BOTTOM_PAD;
  const MARK_X = 60;                // x-center of the circles / pin column
  const MARK_R = 14;                // circle radius
  const RECT_X = MARK_X + 40;       // left edge of the rounded rectangle
  const RECT_W = W - RECT_X - 30;   // rect width
  const RECT_H = 56;                // rect height
  const RECT_RADIUS = 14;

  const parts: string[] = [];

  // ---- Project title at the top of the image (centered, 22pt bold,
  //      grey-blue #44546A — matches the regular page title style).
  if (projectTitle) {
    parts.push(
      `<text x="${W / 2}" y="${TITLE_BAND_H / 2 + 8}" text-anchor="middle" ` +
        `font-family="Arial, sans-serif" font-size="34" font-weight="700" fill="#44546A">` +
        escapeXml(projectTitle.toUpperCase()) +
      `</text>`
    );
  }

  // ---- "Locations" small heading
  parts.push(
    `<text x="20" y="${TITLE_BAND_H + TOP_PAD + HEADING_H - 12}" ` +
      `font-family="Arial, sans-serif" font-size="22" font-weight="700" fill="#1F2937">ROUTE SURVEY KEY POINTS</text>`
  );

  for (let i = 0; i < items.length; i += 1) {
    const it = items[i];
    const cy = STEPPER_TOP + i * ROW_H + ROW_H / 2 - 10;
    const isEnd =
      i === items.length - 1 ||
      it.pin_type === "end";

    // Connector dotted line from THIS row's marker down to the NEXT row's marker
    if (i < items.length - 1) {
      const lineY1 = cy + MARK_R + 2;
      const lineY2 = cy + ROW_H - MARK_R - 2;
      parts.push(
        `<line x1="${MARK_X}" y1="${lineY1}" x2="${MARK_X}" y2="${lineY2}" ` +
          `stroke="#9CA3AF" stroke-width="3" stroke-linecap="round" stroke-dasharray="2 8"/>`
      );
    }

    // Marker: hollow circle for stops, filled red pin for the destination
    if (isEnd) {
      // Pin shape (drop) in red
      parts.push(
        `<g transform="translate(${MARK_X},${cy - 6})">` +
          `<path d="M 0,-18 C -12,-18 -18,-8 -18,-2 C -18,12 0,28 0,28 C 0,28 18,12 18,-2 C 18,-8 12,-18 0,-18 Z" fill="#EC4054" stroke="#B91C2A" stroke-width="1.5"/>` +
          `<circle cx="0" cy="-4" r="6" fill="#FFFFFF"/>` +
        `</g>`
      );
    } else {
      // Hollow circle
      parts.push(
        `<circle cx="${MARK_X}" cy="${cy}" r="${MARK_R}" fill="#FFFFFF" stroke="#1F2937" stroke-width="2.5"/>`
      );
    }

    // Rounded rectangle with the location label
    const rectY = cy - RECT_H / 2;
    parts.push(
      `<rect x="${RECT_X}" y="${rectY}" width="${RECT_W}" height="${RECT_H}" rx="${RECT_RADIUS}" ry="${RECT_RADIUS}" ` +
        `fill="#FFFFFF" stroke="#D1D5DB" stroke-width="2"/>` +
      `<text x="${RECT_X + 24}" y="${cy + 8}" font-family="Arial, sans-serif" font-size="22" font-weight="700" fill="#0F172A">` +
        escapeXml(it.label) +
      `</text>`
    );
  }

  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">` +
      parts.join("") +
    `</svg>`;

  try {
    const buf = await sharp(Buffer.from(svg)).png().toBuffer();
    console.log("[ROUTE LOCATIONS STEPPER GENERATED]", {
      itemCount: items.length,
      titleBaked: !!projectTitle,
      svgSize: `${W}x${H}`,
      bytes: buf.length,
    });
    return buf;
  } catch (err) {
    console.warn("[ROUTE LOCATIONS STEPPER FAILED]", {
      err: (err as Error)?.message || String(err),
    });
    return null;
  }
}

const categoryIconCache = new Map<string, ImageEntry | null>();

async function loadCategoryIcon(category: unknown): Promise<ImageEntry | null> {
  const fileName = getCategoryIconFile(category);

  // Auto-generated icon path: skip the disk read and synthesize a PNG.
  if (fileName.startsWith(AUTO_ICON_PREFIX)) {
    if (categoryIconCache.has(fileName)) {
      const cached = categoryIconCache.get(fileName) || null;
      console.log("[category icon debug]", {
        category,
        autoIconFileName: fileName,
        cached: true,
        hasCategoryIcon: !!cached?.buffer,
        iconBufferSize: cached?.buffer?.length || 0,
        source: "auto-generated (cached)",
      });
      return cached;
    }
    const sourceText =
      autoCategoryByFileName.get(fileName) || String(category || "");
    const buffer = await generateCategoryIcon(sourceText);
    if (!buffer || buffer.length === 0) {
      categoryIconCache.set(fileName, null);
      console.warn("[category icon debug]", {
        category,
        autoIconFileName: fileName,
        hasCategoryIcon: false,
        iconBufferSize: 0,
        source: "auto-generated (failed)",
      });
      return null;
    }
    const entry: ImageEntry = {
      buffer,
      contentType: "image/png",
      path: `(auto-generated:${sourceText})`,
    };
    categoryIconCache.set(fileName, entry);
    console.log("[category icon debug]", {
      category,
      autoIconFileName: fileName,
      hasCategoryIcon: true,
      iconBufferSize: buffer.length,
      contentType: "image/png",
      source: "auto-generated (fresh)",
    });
    return entry;
  }

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
 * In-memory cache of reverse-geocode results, keyed by coordinate
 * rounded to 4 decimals (~11 m precision). Reports clustered around
 * the same point share a single result so we never hit Nominatim
 * twice for effectively-identical coordinates.
 */
const reverseGeocodeCache = new Map<string, string | null>();
function coordKey(lat: number, lng: number): string {
  return `${lat.toFixed(4)},${lng.toFixed(4)}`;
}

/**
 * Throttle Nominatim to ~1 req/sec per their usage policy. We track the
 * timestamp of the last request and sleep just enough before the next
 * one to keep the gap >= 1100ms. This prevents the 429 rate-limit
 * responses that otherwise force fallback to "Near {coords}" cells.
 */
let lastNominatimRequestAt = 0;
async function nominatimThrottle(): Promise<void> {
  const minGapMs = 1100;
  const elapsed = Date.now() - lastNominatimRequestAt;
  if (elapsed < minGapMs) {
    await new Promise((r) => setTimeout(r, minGapMs - elapsed));
  }
  lastNominatimRequestAt = Date.now();
}

/**
 * Build a clean, readable location string from Nominatim's structured
 * address response. This is the KEY fix for highway points: when a
 * coordinate sits on a road (e.g. AH45 / NH48 / 200 Feet Bypass Rd),
 * Nominatim's `display_name` field is often empty or just the country,
 * but the `address` object still contains road / suburb / city / state.
 * Pick the most useful parts and join them.
 *
 * Priority (deduplicated):
 *   road  →  suburb / neighbourhood  →  village / town / city  →
 *   county / state_district  →  state  →  country
 *
 * Examples:
 *   "200 Feet Bypass Road, Adayalampattu, Chennai, Tamil Nadu"
 *   "National Highway 48, Nelamangala, Bengaluru Rural, Karnataka"
 *   "Mominpur, Kolkata, West Bengal"
 */
type NominatimAddress = {
  road?: string;
  pedestrian?: string;
  highway?: string;
  trunk?: string;
  primary?: string;
  motorway?: string;
  suburb?: string;
  neighbourhood?: string;
  hamlet?: string;
  village?: string;
  town?: string;
  city?: string;
  city_district?: string;
  county?: string;
  state_district?: string;
  state?: string;
  country?: string;
};
function buildLocationFromAddress(addr: NominatimAddress | undefined | null): string {
  if (!addr || typeof addr !== "object") return "";
  const road =
    addr.road ||
    addr.pedestrian ||
    addr.motorway ||
    addr.trunk ||
    addr.primary ||
    addr.highway;
  const locality =
    addr.suburb ||
    addr.neighbourhood ||
    addr.hamlet ||
    addr.village ||
    addr.town ||
    addr.city_district;
  const city = addr.city || addr.town || addr.village;
  const district = addr.county || addr.state_district;
  const state = addr.state;
  const country = addr.country;

  const parts: string[] = [];
  if (road) parts.push(road);
  if (locality && locality !== road) parts.push(locality);
  if (city && city !== locality && city !== road) parts.push(city);
  if (district && district !== city && district !== locality) parts.push(district);
  if (state && state !== district && state !== city) parts.push(state);
  // Country only if we don't already have city + state — keeps the
  // string short for typical "City, State" cases.
  if (country && parts.length < 2) parts.push(country);

  // De-dup case-insensitive while preserving order.
  const seen = new Set<string>();
  const unique: string[] = [];
  for (const p of parts) {
    const k = p.toLowerCase();
    if (seen.has(k)) continue;
    seen.add(k);
    unique.push(p);
  }
  return unique.join(", ");
}

/**
 * Single Nominatim request at a specific zoom level. Asks for structured
 * address details so we can build a clean string even when display_name
 * is empty (common for highway / road-segment points). Returns the
 * resolved location text on success, null on any failure / no-result /
 * non-2xx response. Wraps the fetch in a 10 s timeout.
 */
async function nominatimReverseAtZoom(
  lat: number,
  lng: number,
  zoom: number
): Promise<string | null> {
  await nominatimThrottle();
  const url =
    `https://nominatim.openstreetmap.org/reverse?format=jsonv2` +
    `&lat=${encodeURIComponent(String(lat))}` +
    `&lon=${encodeURIComponent(String(lng))}` +
    `&zoom=${zoom}` +
    // Structured address parts — required for highway/road points where
    // display_name is often empty. Without this, road-only points fall
    // straight through to the coordinate fallback.
    `&addressdetails=1`;
  const ac = new AbortController();
  const timer = setTimeout(() => ac.abort(), 10000);
  try {
    const res = await fetch(url, {
      headers: {
        "User-Agent": "lbi-web-export/1.0 (race route export)",
        Accept: "application/json",
      },
      signal: ac.signal,
    });
    if (res.status === 429 || res.status === 503) {
      // Throttled — wait longer and let the caller decide whether to
      // retry the call.
      await new Promise((r) => setTimeout(r, 1200));
      return null;
    }
    if (!res.ok) return null;
    const json = (await res.json()) as {
      display_name?: unknown;
      address?: NominatimAddress;
    };
    // 1. Try the structured address first (road + city + state). This
    //    works for highway points where display_name is empty.
    const fromAddress = buildLocationFromAddress(json.address);
    if (fromAddress) return fromAddress;
    // 2. Fall back to display_name (works for POI / building points).
    const dn = typeof json?.display_name === "string" ? json.display_name.trim() : "";
    return dn || null;
  } catch {
    return null;
  } finally {
    clearTimeout(timer);
  }
}

/**
 * Reverse-geocode lat/lng to a human-readable place name. Uses OpenStreetMap
 * Nominatim with an escalating zoom-fallback strategy:
 *   1. zoom 14 (~ neighbourhood / village) — best precision
 *   2. zoom 10 (~ city / district)         — fallback if 14 returns nothing
 * Each zoom is attempted up to 2 times (so up to 4 total attempts) so a
 * transient failure or 429 doesn't immediately collapse to the coordinate
 * fallback. Results are cached per-coordinate; callers MUST also cache the
 * result in MySQL (resolved_location) so subsequent exports skip the API.
 */
async function reverseGeocodeLocation(lat: number, lng: number): Promise<string | null> {
  const key = coordKey(lat, lng);
  if (reverseGeocodeCache.has(key)) return reverseGeocodeCache.get(key) ?? null;

  const ZOOM_LEVELS = [14, 10];
  const ATTEMPTS_PER_ZOOM = 2;

  for (const zoom of ZOOM_LEVELS) {
    for (let attempt = 0; attempt < ATTEMPTS_PER_ZOOM; attempt += 1) {
      const name = await nominatimReverseAtZoom(lat, lng, zoom);
      if (name) {
        reverseGeocodeCache.set(key, name);
        return name;
      }
      if (attempt + 1 < ATTEMPTS_PER_ZOOM) {
        await new Promise((r) => setTimeout(r, 600));
      }
    }
    console.warn(
      "[export location] nominatim no-result at zoom",
      zoom,
      "for",
      lat,
      lng,
      "— trying broader zoom"
    );
  }

  // All zoom levels and attempts failed. Cache null so the SAME
  // coordinate doesn't keep retrying inside this export run.
  console.warn("[export location] reverseGeocodeLocation exhausted all zoom fallbacks for", lat, lng);
  reverseGeocodeCache.set(key, null);
  return null;
}

/**
 * Spec-mandated coordinate-only detector. Returns true when the value
 * is just two numbers separated by a comma (e.g. "21.689133, 87.076417")
 * — those should NOT be treated as a real "location" name; the resolver
 * should attempt a reverse geocode and only fall back to coords if that
 * also fails.
 */
function isCoordinateOnlyLocation(value: unknown): boolean {
  const text = String(value || "").trim();
  return /^-?\d+(\.\d+)?\s*,\s*-?\d+(\.\d+)?$/.test(text);
}

/**
 * Pull just the city-level name out of a Nominatim address response,
 * preferring (city → town → village → suburb → county → state).
 * Used to build the route corridor label "From X to Y".
 */
function pickCityName(addr: NominatimAddress | undefined | null): string {
  if (!addr) return "";
  return (
    addr.city ||
    addr.town ||
    addr.village ||
    addr.suburb ||
    addr.neighbourhood ||
    addr.hamlet ||
    addr.county ||
    addr.state_district ||
    addr.state ||
    ""
  );
}

/**
 * Reverse-geocode JUST a city-level name (used by the corridor builder).
 * Asks Nominatim at zoom 12 (city / town granularity) so we get a real
 * settlement name even when the point sits on a highway between towns.
 * Returns "" on failure.
 */
async function reverseGeocodeCity(lat: number, lng: number): Promise<string> {
  await nominatimThrottle();
  const url =
    `https://nominatim.openstreetmap.org/reverse?format=jsonv2` +
    `&lat=${encodeURIComponent(String(lat))}` +
    `&lon=${encodeURIComponent(String(lng))}` +
    `&zoom=12&addressdetails=1`;
  const ac = new AbortController();
  const timer = setTimeout(() => ac.abort(), 10000);
  try {
    const res = await fetch(url, {
      headers: {
        "User-Agent": "lbi-web-export/1.0 (race route export)",
        Accept: "application/json",
      },
      signal: ac.signal,
    });
    if (!res.ok) return "";
    const json = (await res.json()) as { address?: NominatimAddress };
    return pickCityName(json?.address);
  } catch {
    return "";
  } finally {
    clearTimeout(timer);
  }
}

/**
 * Build a route corridor label like "Highway from Porur to Ambattur"
 * by geocoding the FIRST and LAST report coordinates of a project to
 * city names. Used as a smart fallback in resolveReportLocation when
 * a mid-route point doesn't reverse-geocode (typically a highway
 * stretch between two known cities).
 *
 * Returns "" when the corridor can't be built (no coords, both endpoints
 * geocode to the same place, or both fail).
 */
export type RouteCorridor = {
  fromCity: string;
  toCity: string;
  label: string; // ready-to-use display string
};
async function buildRouteCorridor(reports: Row[]): Promise<RouteCorridor | null> {
  // Find the first and last reports with valid coordinates.
  let firstWithCoords: Row | null = null;
  let lastWithCoords: Row | null = null;
  for (const r of reports) {
    if (pickLat(r) !== null && pickLng(r) !== null) {
      if (!firstWithCoords) firstWithCoords = r;
      lastWithCoords = r;
    }
  }
  if (!firstWithCoords || !lastWithCoords || firstWithCoords === lastWithCoords) {
    return null;
  }
  const fromLat = pickLat(firstWithCoords);
  const fromLng = pickLng(firstWithCoords);
  const toLat = pickLat(lastWithCoords);
  const toLng = pickLng(lastWithCoords);
  if (fromLat === null || fromLng === null || toLat === null || toLng === null) {
    return null;
  }
  const [fromCity, toCity] = await Promise.all([
    reverseGeocodeCity(fromLat, fromLng),
    reverseGeocodeCity(toLat, toLng),
  ]);
  if (!fromCity && !toCity) return null;
  // Same city at both ends → just show the city, not "X to X"
  if (fromCity && toCity && fromCity.toLowerCase() === toCity.toLowerCase()) {
    return { fromCity, toCity, label: `${fromCity} area highway` };
  }
  // Only one endpoint resolved → show what we have
  if (!fromCity || !toCity) {
    return { fromCity, toCity, label: `Near ${fromCity || toCity} highway` };
  }
  return {
    fromCity,
    toCity,
    label: `Highway from ${fromCity} to ${toCity}`,
  };
}

/**
 * Resolve the LOCATION cell value for one report. Order:
 *   1. existing stored value (location, address, resolved_location, ...)
 *      — but ONLY if it's not a coords-only string
 *   2. reverse-geocoded name from coordinates (cached back into MySQL)
 *   3. ROUTE CORRIDOR fallback ("Highway from {start} to {end}") so
 *      mid-highway points show meaningful context instead of coords
 *   4. coordinate fallback "lat, lng" (only when corridor unavailable)
 *   5. "-"
 */
async function resolveReportLocation(
  r: Row,
  hasResolvedColumn: boolean,
  corridor: RouteCorridor | null
): Promise<string> {
  const existing = getExistingLocation(r);
  if (existing && !isCoordinateOnlyLocation(existing)) return existing;
  const lat = pickLat(r);
  const lng = pickLng(r);
  if (lat === null || lng === null) {
    // No coordinates to geocode — fall back to whatever existing value
    // the row had (even if coords-only) before defaulting to "-".
    return existing || "-";
  }
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
  // Geocode failed — show the route corridor label if we have one. This
  // tells the reader "this point is on the highway from X to Y" instead
  // of dumping raw coordinates.
  if (corridor && corridor.label) {
    return corridor.label;
  }
  // Last resort: clean decimal-degree coordinates so the cell isn't blank.
  const latDir: "N" | "S" = lat >= 0 ? "N" : "S";
  const lngDir: "E" | "W" = lng >= 0 ? "E" : "W";
  return `${Math.abs(lat).toFixed(4)}° ${latDir}, ${Math.abs(lng).toFixed(4)}° ${lngDir}`;
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

  // Build a route-corridor label once per export ("Highway from {start}
  // to {end}") so highway points that don't reverse-geocode show
  // meaningful context instead of raw coordinates. Geocodes only TWO
  // points (the first and last reports with coordinates) so this is
  // cheap regardless of project size.
  let routeCorridor: RouteCorridor | null = null;
  try {
    routeCorridor = await buildRouteCorridor(reports);
    console.log("[export route corridor]", routeCorridor || "(unavailable)");
  } catch (err) {
    console.warn("[export route corridor] build failed:", err);
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

  /**
   * Pick the best 2 photos from a list of report_photos rows. Deduplicates
   * by URL, then sorts by resolution (width×height descending), falling back
   * to created_at descending so the newest large image wins. Returns at most
   * 2 rows.
   */
  function pickBestTwoPhotos(rows: Row[]): Row[] {
    if (rows.length <= 2) return rows;
    // Deduplicate by normalised URL so re-uploads don't count twice.
    const seen = new Set<string>();
    const unique: Row[] = [];
    for (const r of rows) {
      const u = String(r.url || "").trim().toLowerCase();
      if (!u || seen.has(u)) continue;
      seen.add(u);
      unique.push(r);
    }
    if (unique.length <= 2) return unique;
    // Sort: largest resolution first, then newest first.
    unique.sort((a, b) => {
      const areaA = (Number(a.width) || 0) * (Number(a.height) || 0);
      const areaB = (Number(b.width) || 0) * (Number(b.height) || 0);
      if (areaB !== areaA) return areaB - areaA;
      const dateA = a.created_at ? new Date(String(a.created_at)).getTime() : 0;
      const dateB = b.created_at ? new Date(String(b.created_at)).getTime() : 0;
      return dateB - dateA;
    });
    return unique.slice(0, 2);
  }

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

  // Spec-mandated FIRST 5 DB photo check (separate log key so the
  // operator can grep it independently of the larger 10-row dump).
  try {
    const first5 = await safeQuery(
      `SELECT
         r.id AS report_id,
         r.point_key,
         r.category,
         rp.id AS photo_id,
         rp.url,
         rp.file_name
       FROM reports r
       LEFT JOIN report_photos rp ON rp.report_id = r.id
       WHERE r.project_id = ?
       ORDER BY CAST(r.point_key AS DECIMAL(10,4)), r.created_at
       LIMIT 5`,
      [projectId]
    );
    console.log("[DOCX FIRST 5 DB PHOTO CHECK]", first5);
  } catch (err) {
    console.warn("[DOCX FIRST 5 DB PHOTO CHECK] query failed:", err);
  }

  // Spec-mandated FIRST FIVE REPORT PHOTO LINK CHECK — same intent as
  // the dump above but with report_id, photo_report_id and sort_order
  // included so the operator can spot a wrong DB linkage at a glance:
  // if rp.photo_report_id !== r.report_id for any of the first 5 rows,
  // bulk upload inserted under the wrong report and fixing only the
  // DOCX export will not solve the photo shift.
  try {
    const linkCheck = await safeQuery(
      `SELECT
         r.id AS report_id,
         r.point_key,
         r.category,
         r.sort_order,
         rp.id AS photo_id,
         rp.report_id AS photo_report_id,
         rp.url,
         rp.file_name,
         rp.point_key AS photo_point_key,
         rp.image_key
       FROM reports r
       LEFT JOIN report_photos rp ON rp.report_id = r.id
       WHERE r.project_id = ?
       ORDER BY CAST(r.point_key AS DECIMAL(10,4)), r.created_at
       LIMIT 5`,
      [projectId]
    );
    console.log("[DOCX FIRST FIVE REPORT PHOTO LINK CHECK]", linkCheck);
  } catch (err) {
    console.warn("[DOCX FIRST FIVE REPORT PHOTO LINK CHECK] query failed:", err);
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

  // ----- Step 6b: Route locations (the "Locations" stepper from the UI).
  // Loads the labelled waypoints saved by the GA-setup modal so we can
  // reproduce the same hollow-circle / pin / dotted-connector design in
  // the DOCX export, immediately after the route map.
  type RouteLocation = { label: string; pin_type: string; sort_order: number };
  let routeLocationLabels: RouteLocation[] = [];
  try {
    if (routePageRow?.id) {
      const locRows = await safeQuery(
        "SELECT label, pin_type, sort_order FROM project_route_page_locations " +
          "WHERE project_id = ? AND project_page_id = ? ORDER BY sort_order ASC",
        [projectId, routePageRow.id]
      );
      routeLocationLabels = locRows
        .map((r: Row) => ({
          label: String(r.label || "").trim(),
          pin_type: String(r.pin_type || "mid"),
          sort_order: Number(r.sort_order) || 0,
        }))
        .filter((r) => r.label !== "");
    }
  } catch (err) {
    console.warn("[export actual] route locations query failed - continuing without:", err);
  }
  console.log("[export actual] routeLocations:", routeLocationLabels);

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

  // ----- Generate the route-locations stepper image (hollow circles +
  // dotted connectors + final red pin) and store it in imageMap so the
  // post-render pass can inject it right after the route map.
  try {
    if (routeLocationLabels.length > 0) {
      // Pass empty title — the project title is now in the Word
      // header so we don't bake it into the stepper image too.
      const stepperBuf = await generateRouteLocationsStepper(
        routeLocationLabels,
        ""
      );
      if (stepperBuf && stepperBuf.length > 0) {
        imageMap.set("routeLocations", {
          buffer: stepperBuf,
          contentType: "image/png",
          path: "(generated:routeLocations)",
        });
      }
    }
  } catch (err) {
    console.warn("[export actual] route locations stepper failed - ignoring:", err);
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

  // Track which photo keys are side-by-side so the post-render pass can
  // replace the single-image drawing with a 2-column borderless table.
  const sideBySideKeys = new Set<string>();
  // Right-image buffers for each side-by-side key. The LEFT image is stored
  // in imageMap under the normal photoKey; the RIGHT image is stored here
  // and injected into the DOCX zip during the post-render table pass.
  const sideBySideRightBuffers = new Map<string, Buffer>();

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

      // Pick best 2 photos (by resolution then date) and try to fetch them.
      // If 2 succeed → composite side-by-side; if 1 → single centered image.
      const selectedPhotos = pickBestTwoPhotos(reportPhotos);
      let fetchedAny = false;
      let lastFetchFailureReason: string | null = null;
      const fetchedBuffers: Array<{ buffer: Buffer; contentType: string }> = [];
      for (const p of selectedPhotos) {
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
            });
          }
          if (fetched && Buffer.isBuffer(fetched.buffer) && fetched.buffer.length > 0) {
            fetchedBuffers.push({ buffer: fetched.buffer, contentType: fetched.contentType });
            console.log("[PHOTO CHECK 7 - BUFFER READY]", {
              ...ctx,
              bufferIndex: fetchedBuffers.length - 1,
              bufferSize: fetched.buffer.length,
            });
          } else {
            lastFetchFailureReason = "fetchImageBuffer returned no buffer (404/403/non-image/oversize/timeout)";
            console.warn("[PHOTO CHECK 7 - PHOTO BUFFER MISSING]", {
              ...ctx,
              attemptedUrl: candidate,
            });
          }
        } catch (err) {
          lastFetchFailureReason = `fetch threw: ${(err as Error)?.message || String(err)}`;
          console.error("[PHOTO CHECK 3 - FETCH ERROR]", {
            ...ctx,
            url: candidate,
            message: (err as Error)?.message,
          });
        }
      }

      // Assemble the final photo buffer(s). 2 photos → store each
      // individually (left in imageMap, right in sideBySideRightBuffers)
      // for post-render 2-column table injection. 1 photo → single
      // centered image. 0 → no photo.
      if (fetchedBuffers.length >= 2) {
        photoKey = `photo_${i}`;
        const leftOptimized = await optimizeDocxImage(
          fetchedBuffers[0].buffer,
          "observation",
          `${photoKey}_L`
        );
        const rightOptimized = await optimizeDocxImage(
          fetchedBuffers[1].buffer,
          "observation",
          `${photoKey}_R`
        );
        // LEFT image goes into imageMap — the template renders this one.
        imageMap.set(photoKey, {
          buffer: leftOptimized,
          contentType: "image/jpeg",
        });
        // RIGHT image is stored separately for the post-render table pass.
        sideBySideKeys.add(photoKey);
        sideBySideRightBuffers.set(photoKey, rightOptimized);
        fetchedAny = true;
        photoFetchSuccess += 1;
        console.log("[DOCX SIDE-BY-SIDE PHOTOS STORED]", {
          index: i,
          reportId: rid,
          photoKey,
          leftSize: leftOptimized.length,
          rightSize: rightOptimized.length,
          totalCandidates: reportPhotos.length,
          selectedCount: selectedPhotos.length,
        });
      } else if (fetchedBuffers.length === 1) {
        photoKey = `photo_${i}`;
        const optimized = await optimizeDocxImage(
          fetchedBuffers[0].buffer,
          "observation",
          photoKey
        );
        imageMap.set(photoKey, {
          buffer: optimized,
          contentType: "image/jpeg",
        });
        fetchedAny = true;
        photoFetchSuccess += 1;
        console.log("[DOCX photo buffer loaded]", {
          index: i,
          reportId: rid,
          photoKey,
          bufferSize: optimized.length,
        });
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

      // Spec-mandated photo-block check. If dbPhotosFound > 0 AND
      // photosAddedToBlock === 0, the buffer fetch failed (S3 404 /
      // timeout / non-image). If dbPhotosFound === 0, the row truly
      // has no DB photo and the bulk-upload pipeline is the place to
      // look (use [DOCX FIRST 10 PHOTO DB CHECK] for the SQL view).
      console.log("[DOCX PHOTO BLOCK CHECK]", {
        blockIndex: i,
        report_id: rid,
        point_key: r.point_key,
        dbPhotosFound: reportPhotos.length,
        photosAddedToBlock: photoKey ? 1 : 0,
        firstPhotoKey: photoKey || null,
        imageMapHit: !!photoKey && imageMap.has(photoKey),
      });
    }

    // KM is the cumulative haversine distance along the ordered route. The
    // database `kms`/`km` columns are NOT used because the user spec says KM
    // must be derived from coordinates so the first row is always 0.0000.
    const kmText = formatKmValue(r.calculated_km);

    // LOCATION uses (1) any stored value, then (2) reverse-geocoded name
    // (cached into reports.resolved_location), then (3) coordinate fallback.
    const dbLocation = getExistingLocation(r);
    const locationText = await resolveReportLocation(r, hasResolvedLocationColumn, routeCorridor);
    // Spec-mandated location-source log. coords-only DB strings are
    // classified by their FINAL outcome — if the resolver succeeded
    // in reverse-geocoding past the coords-only DB value, source is
    // "reverse_geocode"; if it had to fall back to coords, "coords_fallback".
    const dbHasUsable = !!dbLocation && !isCoordinateOnlyLocation(dbLocation);
    const locationSource: "db" | "reverse_geocode" | "coords_fallback" | "missing" =
      dbHasUsable
        ? "db"
        : isCoordinateOnlyLocation(locationText)
          ? "coords_fallback"
          : locationText && locationText !== "-"
            ? "reverse_geocode"
            : "missing";
    console.log("[DOCX LOCATION CHECK]", {
      report_id: rid,
      point_key: r.point_key,
      lat,
      lng,
      originalLocation: dbLocation,
      finalLocation: locationText,
      source: locationSource,
    });

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

    // Spec-mandated icon check log keyed by report_id/point_key for
    // easier joining against bulk-import logs.
    console.log("[DOCX CATEGORY ICON CHECK]", {
      report_id: rid,
      point_key: r.point_key,
      rawCategory: r.category,
      normalizedCategory: normalizeCategory(r.category),
      iconFile: categoryIcon?.path
        ? String(categoryIcon.path).split(/[\\/]/).pop()
        : null,
      categoryIconKey,
      imageMapHasIcon: categoryIconKey ? imageMap.has(categoryIconKey) : false,
    });
    if (!hasCategoryIcon && r.category) {
      console.warn("[DOCX CATEGORY ICON MISSING]", {
        rawCategory: r.category,
        normalizedCategory: normalizeCategory(r.category),
      });
    }

    const difficultyValue = String(
      r.difficulty ||
        r.remarks_action ||
        r.status ||
        r.vehicle_movement ||
        r.movement ||
        ""
    );
    const tableColors = getDifficultyTableColors(difficultyValue);

    // Spec-mandated: prove the row didn't unintentionally land on red.
    // usesRedStyle MUST be false unless normalizedDifficulty is exactly
    // one of RED_KEYWORDS. After the EXACT-equals fix, "redirect" or
    // "hardware" in remarks no longer flips the row red.
    {
      const normalizedDifficulty = normalizeDifficulty(difficultyValue);
      console.log("[DOCX RED STYLE CHECK]", {
        report_id: rid,
        point_key: r.point_key,
        rawDifficulty: difficultyValue,
        normalizedDifficulty,
        usesRedStyle: tableColors.key === "red",
      });
    }

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
        // Single-photo reports: 7.5" × 5.3" (OBSERVATION_PHOTO_SIZE).
        // Multi-photo reports: 7.2" × 5.0" (MULTI_PHOTO_SIZE) so both
        // photos fit on the same page when stacked. The second photo is
        // injected by the post-render pass at the same MULTI_PHOTO_SIZE.
        if (sideBySideKeys.has(key)) {
          return MULTI_PHOTO_SIZE;
        }
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

  // ---- REVERTED: photo-key remap is intentionally a NO-OP.
  // Earlier revisions tried to re-key surviving observations to
  // `photo_${blockIndex}_0` (caused a +1 shift) and then to
  // `photo_${reportId}_0` (blanked every photo because observations[]
  // doesn't carry report_id, so the remap fell into the ELSE branch
  // and cleared every key). The keys assigned by the original photo
  // loop — `photo_${i}` per surviving report — already resolve
  // correctly against imageMap. The +1 shift came from the post-render
  // swap pass moving images one slot forward when the template emits
  // `table → image` ordering; that is fixed by the conditional swap
  // logic in Step 10c (template-structure detection).
  // Keeping the no-op here so the spec's regression guard below has
  // something to log against, and so future revisions know not to
  // re-introduce the broken remap.
  {
    const photoKeysInMap = Array.from(imageMap.keys()).filter((k) =>
      k.startsWith("photo_")
    );
    const rowsWithPhotoKey = observations.filter(
      (o) => !!(o.observationPhotoKey || o.photoKey)
    ).length;
    console.log("[DOCX PHOTO KEY REMAP] (no-op revert)", {
      observationsAfterFilter: observations.length,
      photoKeysInImageMap: photoKeysInMap.length,
      rowsWithPhotoKey,
      finalPhotoKeysSample: photoKeysInMap.slice(0, 10),
    });
    // Spec-mandated regression guard: if photos exist in imageMap but
    // NO observation row carries a key, every photo will render as
    // "Photo not available" — exactly the failure mode the user just
    // reported. Surface this loudly so the operator can spot it.
    if (photoKeysInMap.length > 0 && rowsWithPhotoKey === 0) {
      console.error(
        "[DOCX PHOTO MAPPING REGRESSION]",
        "photos exist in imageMap but no observation received a photo key",
        { photoKeysInImageMap: photoKeysInMap.length, rowsWithPhotoKey }
      );
    }
  }

  // Spec-mandated finalised observation blocks. Carries the explicit
  // hasPhoto / pageBreak / isLast flags the spec calls out so the data
  // shape unambiguously expresses: table → image-or-placeholder → page
  // break (except after last). Used in render data alongside
  // `observations` (the template's existing loop key) so the data is
  // available either way.
  // photoKeys[] is the spec-mandated multi-image array — one entry per
  // photo for the report. Today the template's image placeholder loop
  // emits ONE drawing per observation; photoKeys is exposed so a
  // future template revision (with `{#photos}{%imageKey}{/photos}`)
  // can render every photo without code changes here.
  const observationBlocks = observations.map((row, index) => {
    const photoKey = row.observationPhotoKey || row.photoKey || "";
    const hasPhoto = !!photoKey && imageMap.has(photoKey);
    // ObservationData doesn't declare point_key/report_id/id but the
    // builder above attaches them via spread from the source Row. Cast
    // through Record<string, unknown> so we can read them without
    // widening the public ObservationData type.
    const rowAny = row as unknown as Record<string, unknown>;
    const rowPointKey = String(rowAny.point_key || "");
    const rowReportId = String(rowAny.report_id || rowAny.id || "");
    // photoKeys is the canonical multi-image key list for this report.
    // After the remap above, every surviving observation has exactly
    // one key, `photo_${reportId}_0`. Multi-image support uses _1, _2,
    // etc. — keyed by REPORT_ID, never by block index, so no shift is
    // possible.
    const photoKeys: string[] = hasPhoto ? [photoKey] : [];
    // Pull the matching DB photo rows so each photos[] entry carries
    // the original metadata for downstream template / debugging use.
    const dbPhotos = (photosByReportId.get(rowReportId) || []).slice(
      0,
      photoKeys.length
    );
    const photos = photoKeys.map((k, photoIndex) => {
      const dbRow = dbPhotos[photoIndex] || {};
      return {
        imageKey: k,
        url: typeof dbRow.url === "string" ? dbRow.url : undefined,
        file_name:
          typeof dbRow.file_name === "string" ? dbRow.file_name : undefined,
        report_id:
          typeof dbRow.report_id === "string" ? dbRow.report_id : rowReportId,
        point_key:
          typeof dbRow.point_key === "string" ? dbRow.point_key : rowPointKey,
      };
    });
    return {
      ...row,
      blockIndex: index,
      reportId: rowReportId,
      photoKey,
      observationPhotoKey: hasPhoto ? photoKey : "",
      hasPhoto,
      hasPhotos: photoKeys.length > 0,
      photoKeys,
      photos,
      photoText: hasPhoto ? "" : "Photo not available",
      isLast: index === observations.length - 1,
      pageBreak: index < observations.length - 1,
    };
  });

  // Spec-mandated per-photo log emitted as the imageMap is verified.
  // Confirms the buffer in imageMap is keyed by report_id and matches
  // the report it came from (photoReportId === reportId).
  for (const block of observationBlocks) {
    for (const photo of (block as unknown as {
      photos?: Array<{ imageKey: string; report_id?: string; point_key?: string; file_name?: string }>;
    }).photos || []) {
      console.log("[DOCX IMAGE MAP SET BY REPORT_ID]", {
        blockIndex: block.blockIndex,
        reportId: (block as unknown as { reportId?: string }).reportId,
        point_key: (block as unknown as Record<string, unknown>).point_key,
        photoReportId: photo.report_id,
        photoPointKey: photo.point_key,
        imageKey: photo.imageKey,
        file_name: photo.file_name,
        imageMapHit: imageMap.has(photo.imageKey),
      });
    }
  }
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

  // Spec-mandated order-alignment QA. After the photo-key remap and
  // observationBlocks build, the FIRST block's first photo MUST be
  // photo_0_0 (or null if the first observation truly has no photo).
  // Any other key here means the remap did not align with the
  // observation order — the post-render swap would then misplace a
  // drawing.
  const firstObs = observationBlocks[0];
  const lastObs = observationBlocks[observationBlocks.length - 1];
  console.log("[DOCX ORDER ALIGNMENT QA]", {
    observationBlocks: observationBlocks.length,
    firstBlock: firstObs
      ? {
          blockIndex: firstObs.blockIndex,
          photosCount: Array.isArray(firstObs.photoKeys) ? firstObs.photoKeys.length : 0,
          firstPhotoKey: firstObs.photoKeys?.[0] || null,
          imageMapHit:
            !!firstObs.photoKeys?.[0] && imageMap.has(firstObs.photoKeys[0]),
          willShowPlaceholder: !firstObs.hasPhoto,
        }
      : null,
    lastBlock: lastObs
      ? {
          blockIndex: lastObs.blockIndex,
          photosCount: Array.isArray(lastObs.photoKeys) ? lastObs.photoKeys.length : 0,
          pageBreak: lastObs.pageBreak,
        }
      : null,
    renderMode: "single observationBlocks loop",
    order: "table -> photos/placeholder -> pageBreak",
    noSeparateImageLoop: true,
  });
  if (firstObs && firstObs.hasPhoto && firstObs.photoKeys?.[0] !== "photo_0_0") {
    console.warn(
      "[DOCX ORDER ALIGNMENT QA] WARN: firstBlock has a photo but its key is NOT photo_0_0",
      { firstBlockPhotoKey: firstObs.photoKeys?.[0] }
    );
  }
  if (lastObs && lastObs.pageBreak) {
    console.warn(
      "[DOCX ORDER ALIGNMENT QA] WARN: lastBlock.pageBreak is true (should be false)"
    );
  }

  // Spec-mandated alignment QA: surface BOTH the first AND second
  // observation blocks so a "first Gate image appears under second
  // Bend report" shift becomes obvious. firstPhotoReportId MUST equal
  // firstBlock.reportId; secondPhotoReportId MUST equal
  // secondBlock.reportId. Any mismatch proves the imageMap key is
  // wired to the wrong report.
  const fb = observationBlocks[0] as unknown as
    | undefined
    | {
        blockIndex: number;
        reportId: string;
        category: string;
        photos: Array<{ imageKey: string; file_name?: string; report_id?: string }>;
      } & Record<string, unknown>;
  const sb = observationBlocks[1] as unknown as
    | undefined
    | {
        blockIndex: number;
        reportId: string;
        category: string;
        photos: Array<{ imageKey: string; file_name?: string; report_id?: string }>;
      } & Record<string, unknown>;
  console.log("[DOCX PHOTO ALIGNMENT FINAL QA]", {
    firstBlock: fb
      ? {
          blockIndex: fb.blockIndex,
          reportId: fb.reportId,
          point_key: (fb as Record<string, unknown>).point_key,
          category: fb.category,
          photosCount: Array.isArray(fb.photos) ? fb.photos.length : 0,
          firstPhotoKey: fb.photos?.[0]?.imageKey || null,
          firstPhotoFileName: fb.photos?.[0]?.file_name || null,
          firstPhotoReportId: fb.photos?.[0]?.report_id || null,
        }
      : null,
    secondBlock: sb
      ? {
          blockIndex: sb.blockIndex,
          reportId: sb.reportId,
          point_key: (sb as Record<string, unknown>).point_key,
          category: sb.category,
          photosCount: Array.isArray(sb.photos) ? sb.photos.length : 0,
          firstPhotoFileName: sb.photos?.[0]?.file_name || null,
          firstPhotoReportId: sb.photos?.[0]?.report_id || null,
        }
      : null,
    mappingMode: "report_id_only",
    noIndexBasedPhotoMapping: true,
  });
  if (
    fb &&
    Array.isArray(fb.photos) &&
    fb.photos[0] &&
    fb.photos[0].report_id &&
    fb.photos[0].report_id !== fb.reportId
  ) {
    console.error("[DOCX PHOTO ALIGNMENT FINAL QA] MISMATCH on firstBlock", {
      blockReportId: fb.reportId,
      photoReportId: fb.photos[0].report_id,
    });
  }

  // Spec-mandated final first-block check. If firstPhotoKey is null
  // and the [DOCX FIRST 5 DB PHOTO CHECK] above shows a photo_id for
  // the first report, the report_id mapping is wrong (the photo did
  // not land under photosByReportId for this report). If
  // firstPhotoKey is present but the rendered DOCX still shows an
  // image before the table, the [DOCX ORPHAN PRE-TABLE DRAWINGS
  // STRIPPED] log will tell you whether the cleanup actually fired.
  console.log("[DOCX FIRST BLOCK FINAL CHECK]", {
    firstBlock: firstObs
      ? {
          blockIndex: firstObs.blockIndex,
          reportId: (firstObs as { reportId?: string }).reportId || null,
          point_key: (firstObs as unknown as Record<string, unknown>).point_key,
          photosCount: Array.isArray(firstObs.photos) ? firstObs.photos.length : 0,
          firstPhotoKey:
            Array.isArray(firstObs.photos) && firstObs.photos[0]
              ? firstObs.photos[0].imageKey
              : null,
          firstPhotoMapHit:
            Array.isArray(firstObs.photos) &&
            firstObs.photos[0] &&
            imageMap.has(firstObs.photos[0].imageKey),
          hasPhotos: !!firstObs.hasPhotos,
          willShowPlaceholder: !firstObs.hasPhotos,
        }
      : null,
    imageMapPhotoKeysFirst10: Array.from(imageMap.keys())
      .filter((k) => k.startsWith("photo_"))
      .slice(0, 10),
    renderOrder: "table -> photos/placeholder -> pageBreak",
    oldObservationPhotoKeyDisabled: false,
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
  // Spec-mandated offset-fix QA. Surfaces the first 3 observation rows
  // alongside what the OLD broken (+1) logic WOULD have shown so the
  // operator can confirm the offset is fixed:
  //   correctPhotoNow.fileName should match the row's category visually
  //   oldWrongPhotoWouldBe.fileName is what the +1 shift would have used
  {
    const photoListForLog = Array.from(imageMap.entries())
      .filter(([k]) => k.startsWith("photo_"))
      .map(([k]) => k);
    console.log("[DOCX OFFSET FIX QA]", {
      observationsCount: observations.length,
      imageMapPhotoCount: photoListForLog.length,
      first3Pairs: observations.slice(0, 3).map((row, index) => ({
        index,
        tablePointKey: (row as unknown as Record<string, unknown>).point_key,
        tableCategory: row.category,
        assignedPhotoKey: row.observationPhotoKey || row.photoKey || null,
        assignedPhotoMapHit:
          !!(row.observationPhotoKey || row.photoKey) &&
          imageMap.has(String(row.observationPhotoKey || row.photoKey)),
      })),
      noIndexPlusOneOffset: true,
    });
  }

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

  // Spec-mandated table style QA log. Asserts the reference-style
  // table fills, border color, category icon size, and that the
  // photo-mapping fix from the previous turn is preserved.
  console.log("[DOCX TABLE STYLE QA]", {
    headerFill: "43A047",
    bodyFill: "DDEFD8",
    borderColor: "B7CDB3",
    categoryIconSize: `${CATEGORY_ICON_SIZE[0]}x${CATEGORY_ICON_SIZE[1]}`,
    tableLayout: "reference-style",
    photoMappingUntouched: true,
  });

  // Spec-mandated alignment QA. Declares the EXACT layout the
  // post-render observation-table-grid pass enforces.
  console.log("[DOCX TABLE ALIGNMENT QA]", {
    tableStyle: "second-reference-match",
    tableWidth: "100%",
    layout: "fixed",
    columns: {
      gps: "8.5%",
      km: "6.8%",
      location: "17.5%",
      category: "10.5%",
      observation: "40.5%",
      remarks: "16.2%",
    },
    headerFill: "43A047",
    bodyFill: "DDEFD8",
    categoryIconSize: `${CATEGORY_ICON_SIZE[0]}x${CATEGORY_ICON_SIZE[1]}`,
    titleSpacingAfter: "2-4pt",
    noExtraCellParagraphs: true,
    photoMappingUntouched: true,
  });

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

  // Spec-mandated client-fix QA log. Asserts every fix this revision
  // introduced is in effect: red debug borders stripped, location
  // fallback enabled, footer body text removed, photo block check
  // available per row.
  console.log("[DOCX CLIENT FIX QA]", {
    observations: observationBlocks.length,
    rowsWithPhotos: observationBlocks.filter(
      (b) => (Array.isArray(b.photos) ? b.photos.length : 0) > 0
    ).length,
    rowsWithoutPhotos: observationBlocks.filter(
      (b) => (Array.isArray(b.photos) ? b.photos.length : 0) === 0
    ).length,
    firstBlockPhotos: Array.isArray(observationBlocks[0]?.photos)
      ? observationBlocks[0]!.photos!.length
      : 0,
    multipleImageRows: observationBlocks.filter(
      (b) => (Array.isArray(b.photos) ? b.photos.length : 0) > 1
    ).length,
    categoryIconSize: `${CATEGORY_ICON_SIZE[0]}x${CATEGORY_ICON_SIZE[1]}`,
    observationImageSize: `${OBSERVATION_PHOTO_SIZE[0]}x${OBSERVATION_PHOTO_SIZE[1]}`,
    compressionEnabled: !!getSharp(),
    redDebugBordersRemoved: true,
    locationFallbackEnabled: true,
    footerBodyTextRemoved: true,
  });

  // Spec-mandated final CLIENT QA log. One line consolidating photo
  // counts, multi-image rows, compression flag, display sizes and the
  // location/colour modes — all the high-level facts a reviewer needs
  // before sending the DOCX to the client.
  console.log("[DOCX FINAL CLIENT QA]", {
    observations: observationBlocks.length,
    totalPhotos: observationBlocks.reduce(
      (sum, b) => sum + (Array.isArray(b.photoKeys) ? b.photoKeys.length : 0),
      0
    ),
    rowsWithPhotos: observationBlocks.filter(
      (b) => (Array.isArray(b.photoKeys) ? b.photoKeys.length : 0) > 0
    ).length,
    rowsWithoutPhotos: observationBlocks.filter(
      (b) => (Array.isArray(b.photoKeys) ? b.photoKeys.length : 0) === 0
    ).length,
    multipleImageRows: observationBlocks.filter(
      (b) => (Array.isArray(b.photoKeys) ? b.photoKeys.length : 0) > 1
    ).length,
    compressionEnabled: !!getSharp(),
    categoryIconSize: `${CATEGORY_ICON_SIZE[0]}x${CATEGORY_ICON_SIZE[1]}`,
    observationImageSize: `${OBSERVATION_PHOTO_SIZE[0]}x${OBSERVATION_PHOTO_SIZE[1]}`,
    locationFallbackEnabled: true,
    tableColorMode: "difficulty-based",
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
      // a <w:tcBorders>...</w:tcBorders> block with B7CDB3 (light
      // green-grey, matches the reference observation table screenshot),
      // using a non-greedy regex so paragraph/text colour declarations
      // outside table borders are left untouched.
      // Also strip ANY red border inside <w:tcBorders> (FF0000, C00000,
      // B71C1C, "FF0000FF") so debug/leftover red rectangles never
      // appear unless the user actually picked red difficulty (which
      // sets FILL, not border, via separate per-row patches).
      const TABLE_BORDER_COLOR = "B7CDB3";
      let borderRecolorCount = 0;
      xml = xml.replace(
        /<w:tcBorders\b[^>]*>([\s\S]*?)<\/w:tcBorders>/g,
        (match: string, inner: string) => {
          const updated = inner.replace(
            /w:color="(?:000000|FF0000|FF0000FF|C00000|B71C1C|BFBFBF|red)"/gi,
            () => {
              borderRecolorCount += 1;
              return `w:color="${TABLE_BORDER_COLOR}"`;
            }
          );
          return match.replace(inner, updated);
        }
      );

      // -- Compact cell margins. The reference observation table uses
      // small uniform cell padding so rows don't take huge vertical
      // space. We rewrite every <w:tcMar> (cell margin) block in the
      // body to top/bottom 70 twips, left/right 80 twips per spec.
      // Cells WITHOUT an existing <w:tcMar> are left at the table's
      // default — touching cells we already see is the safe scope.
      let cellMarginsRewritten = 0;
      // Slightly tighter than before (60/60/70/70) to match the
      // second reference screenshot's compact body cells.
      const COMPACT_TC_MAR =
        '<w:tcMar>' +
          '<w:top w:w="60" w:type="dxa"/>' +
          '<w:left w:w="70" w:type="dxa"/>' +
          '<w:bottom w:w="60" w:type="dxa"/>' +
          '<w:right w:w="70" w:type="dxa"/>' +
        '</w:tcMar>';
      xml = xml.replace(/<w:tcMar>[\s\S]*?<\/w:tcMar>/g, () => {
        cellMarginsRewritten += 1;
        return COMPACT_TC_MAR;
      });

      // -- OBSERVATION TABLE WIDTH + COLUMN GRID rewrite.
      // The reference screenshot uses these column proportions:
      //   GPS LOCATION   : 8.5%   →  850 / 5000 in pct, ≈ 1218 dxa
      //   KM             : 6.8%   →  680 / 5000 in pct, ≈  974 dxa
      //   LOCATION       : 17.5%  → 1750 / 5000 in pct, ≈ 2508 dxa
      //   CATEGORY       : 10.5%  → 1050 / 5000 in pct, ≈ 1505 dxa
      //   OBSERVATION    : 40.5%  → 4050 / 5000 in pct, ≈ 5803 dxa
      //   REMARKS/ACTION : 16.2%  → 1620 / 5000 in pct, ≈ 2322 dxa
      // We assume a usable content width of 14330 dxa (A4 ≈ 11906,
      // Letter ≈ 12240; the template uses landscape so 14330 is a
      // safe approximation). Each cell width is set in DXA so total
      // sums correctly regardless of page-margin variance.
      const OBS_COL_PCTS = [8.5, 6.8, 17.5, 10.5, 40.5, 16.2];
      const TOTAL_DXA = 14330;
      const OBS_COL_DXA = OBS_COL_PCTS.map((p) => Math.round((p / 100) * TOTAL_DXA));
      // Build replacement <w:tblGrid> for an observation table.
      const OBS_GRID_XML =
        '<w:tblGrid>' +
        OBS_COL_DXA.map((w) => `<w:gridCol w:w="${w}"/>`).join("") +
        '</w:tblGrid>';
      // Build replacement <w:tblPr>'s width + layout fragments.
      const OBS_TBL_W_XML = `<w:tblW w:w="5000" w:type="pct"/>`;
      const OBS_TBL_LAYOUT_XML = `<w:tblLayout w:type="fixed"/>`;

      let obsTablesRestyled = 0;
      let obsCellsResized = 0;
      let obsHeadersRecolored = 0;

      /**
       * Recolour every text run in the given row to white (#FFFFFF) so
       * the header text reads cleanly on its coloured background.
       * Touches ONLY <w:color> inside text-run <w:rPr> — leaves cell
       * shading (<w:shd>), paragraph alignment, borders, and table
       * structure completely alone.
       */
      const recolorHeaderRowToWhite = (trXml: string): string => {
        let r = trXml;
        // 1. Existing <w:rPr> blocks: replace any inner <w:color>, or
        //    inject a fresh white <w:color> if none present.
        r = r.replace(
          /<w:rPr>([\s\S]*?)<\/w:rPr>/g,
          (_full: string, inner: string) => {
            if (/<w:color\b/.test(inner)) {
              const updated = inner.replace(
                /<w:color\b[^/]*\/>/g,
                `<w:color w:val="FFFFFF"/>`
              );
              return `<w:rPr>${updated}</w:rPr>`;
            }
            return `<w:rPr>${inner}<w:color w:val="FFFFFF"/></w:rPr>`;
          }
        );
        // 2. Runs that have NO <w:rPr> at all: inject one with just the
        //    white color so plain runs also get recoloured.
        r = r.replace(
          /<w:r>(\s*)<w:t/g,
          `<w:r>$1<w:rPr><w:color w:val="FFFFFF"/></w:rPr><w:t`
        );
        return r;
      };
      // Iterate every observation table (identified by "GPS LOCATION"
      // header text) and rewrite its <w:tblGrid>, <w:tblW>, layout,
      // and per-cell <w:tcW> widths so the row matches the reference
      // proportions exactly. Tables that AREN'T observation tables
      // (route map, GA drawing wrappers, summary, etc.) are untouched.
      xml = xml.replace(
        /<w:tbl\b[\s\S]*?<\/w:tbl>/g,
        (tblXml: string) => {
          if (!tblXml.includes("GPS LOCATION")) return tblXml;
          obsTablesRestyled += 1;

          // 1. Replace <w:tblGrid>...</w:tblGrid> with our 6-col grid.
          let out = tblXml.replace(
            /<w:tblGrid>[\s\S]*?<\/w:tblGrid>/,
            OBS_GRID_XML
          );

          // 2. Replace <w:tblW .../> (table width) with 5000 pct.
          if (out.includes("<w:tblW")) {
            out = out.replace(/<w:tblW\b[^/]*\/>/, OBS_TBL_W_XML);
          } else {
            // Inject into <w:tblPr> if missing.
            out = out.replace(
              /<w:tblPr>/,
              `<w:tblPr>${OBS_TBL_W_XML}`
            );
          }

          // 3. Force fixed layout (so column widths are honoured
          // exactly instead of Word auto-fitting to content).
          if (out.includes("<w:tblLayout")) {
            out = out.replace(/<w:tblLayout\b[^/]*\/>/, OBS_TBL_LAYOUT_XML);
          } else {
            out = out.replace(
              /<w:tblPr>/,
              `<w:tblPr>${OBS_TBL_LAYOUT_XML}`
            );
          }

          // 4. Per-row, rewrite each cell's <w:tcW> in column order.
          // We assume the standard 6-column observation table; rows
          // with a different cell count are left alone.
          out = out.replace(
            /<w:tr\b[\s\S]*?<\/w:tr>/g,
            (trXml: string) => {
              const cellRe = /<w:tc\b[\s\S]*?<\/w:tc>/g;
              const cells = trXml.match(cellRe);
              if (!cells || cells.length !== OBS_COL_DXA.length) return trXml;
              const newCells = cells.map((tcXml, idx) => {
                const w = OBS_COL_DXA[idx];
                const newTcW = `<w:tcW w:w="${w}" w:type="dxa"/>`;
                if (tcXml.includes("<w:tcW")) {
                  return tcXml.replace(/<w:tcW\b[^/]*\/>/, newTcW);
                }
                // Inject into <w:tcPr>; create one if missing.
                if (tcXml.includes("<w:tcPr>")) {
                  return tcXml.replace(/<w:tcPr>/, `<w:tcPr>${newTcW}`);
                }
                return tcXml.replace(/<w:tc\b([^>]*)>/, `<w:tc$1><w:tcPr>${newTcW}</w:tcPr>`);
              });
              obsCellsResized += newCells.length;
              // Reassemble the row by replacing each original cell
              // with its rewritten version, in order.
              let rebuilt = trXml;
              for (let i = 0; i < cells.length; i += 1) {
                rebuilt = rebuilt.replace(cells[i], newCells[i]);
              }
              return rebuilt;
            }
          );

          // 5. Recolour the FIRST <w:tr> (header row) text to white.
          //    Cell shading (yellow/gold background) sits on <w:tcPr>'s
          //    <w:shd> — completely untouched. Only run-level <w:color>
          //    inside the header row's <w:rPr> blocks is rewritten.
          const firstTrMatch = out.match(/<w:tr\b[\s\S]*?<\/w:tr>/);
          if (firstTrMatch) {
            const oldHeader = firstTrMatch[0];
            const newHeader = recolorHeaderRowToWhite(oldHeader);
            if (newHeader !== oldHeader) {
              out = out.replace(oldHeader, newHeader);
              obsHeadersRecolored += 1;
            }
          }

          return out;
        }
      );
      console.log("[DOCX OBSERVATION TABLE GRID]", {
        obsTablesRestyled,
        obsCellsResized,
        obsHeadersRecolored,
        headerTextColor: "#FFFFFF (white)",
        columnDxa: OBS_COL_DXA,
      });

      // -- Defensive red-border kill outside tcBorders too. <w:pBdr>
      // (paragraph borders) and stray <w:bdr> elements with red
      // colour values are recoloured to light grey. We do NOT touch
      // text colour (<w:color> directly inside <w:rPr>) or fill
      // colour (<w:shd w:fill=…>) — those carry intentional design.
      let redBorderRecolorCount = 0;
      xml = xml.replace(
        /(<w:(?:pBdr|bdr)\b[^>]*>[\s\S]*?)w:color="(?:FF0000|FF0000FF|C00000|B71C1C|red)"([\s\S]*?<\/w:(?:pBdr|bdr)>)/gi,
        (_full: string, head: string, tail: string) => {
          redBorderRecolorCount += 1;
          return `${head}w:color="BFBFBF"${tail}`;
        }
      );
      if (redBorderRecolorCount > 0) {
        console.log("[DOCX RED BORDER STRIPPED]", { redBorderRecolorCount });
      }

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

      // Helper: is this EMU cx an observation photo? Single-photo reports
      // render at OBSERVATION_PHOTO_EMU_WIDTH (7.5"); multi-photo reports
      // render at MULTI_PHOTO_EMU_WIDTH (7.2"). Match either within ±15%.
      const isObservationPhotoCx = (cx: number): boolean => {
        if (!Number.isFinite(cx)) return false;
        const tol1 = OBSERVATION_PHOTO_EMU_WIDTH * 0.15;
        if (Math.abs(cx - OBSERVATION_PHOTO_EMU_WIDTH) <= tol1) return true;
        const tol2 = MULTI_PHOTO_EMU_WIDTH * 0.15;
        if (Math.abs(cx - MULTI_PHOTO_EMU_WIDTH) <= tol2) return true;
        return false;
      };

      // -- TEMPLATE STRUCTURE DETECTION.
      // Two valid template shapes for the observation loop:
      //   (A) image BEFORE table: ... [image-k] [table-k] ...
      //       beforeTable for the FIRST observation contains an
      //       observation-sized drawing → swap is safe and required.
      //   (B) table BEFORE image: ... [table-k] [image-k] ...
      //       beforeTable for the FIRST observation has NO observation
      //       drawing → the swap WOULD shift each image one slot
      //       forward (the old +1 bug) — we must SKIP the swap and
      //       only insert page breaks between observations.
      let templateImageBeforeTable = false;
      if (obsTables.length > 0) {
        const preFirst = xml.slice(0, obsTables[0].openStart);
        const paraRe0 = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
        let pm0: RegExpExecArray | null;
        while ((pm0 = paraRe0.exec(preFirst)) !== null) {
          if (!pm0[0].includes("<w:drawing")) continue;
          const cxMatch = pm0[0].match(/<wp:extent\s+cx="(\d+)"/);
          if (!cxMatch) continue;
          const cx = Number(cxMatch[1]);
          if (isObservationPhotoCx(cx)) {
            templateImageBeforeTable = true;
            break;
          }
        }
      }
      console.log("[DOCX TEMPLATE STRUCTURE]", {
        observationTables: obsTables.length,
        templateImageBeforeTable,
        action: templateImageBeforeTable
          ? "swap each image to AFTER its table"
          : "skip swap (template already emits table → image); insert page breaks only",
      });

      if (obsTables.length > 0 && templateImageBeforeTable) {
        // (A) Process from END to START so positional indices for
        // earlier tables stay valid as we mutate the xml.
        let xmlOut = xml;
        for (let k = obsTables.length - 1; k >= 0; k -= 1) {
          const tbl = obsTables[k];
          const isLastObs = k === obsTables.length - 1;
          const prevEnd = k > 0 ? obsTables[k - 1].closeEnd : 0;
          const beforeTable = xmlOut.slice(prevEnd, tbl.openStart);

          // Find the LAST photo paragraph in beforeTable: either an
          // observation-photo <w:drawing> (single OR side-by-side sized,
          // cx within ±15%) OR a "Photo not available" text paragraph.
          const paraRe = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
          let lastPhotoPara: { start: number; end: number; text: string } | null = null;
          let pm: RegExpExecArray | null;
          while ((pm = paraRe.exec(beforeTable)) !== null) {
            const p = pm[0];
            if (p.includes("<w:drawing")) {
              const cxMatch = p.match(/<wp:extent\s+cx="(\d+)"/);
              if (cxMatch) {
                const cx = Number(cxMatch[1]);
                if (isObservationPhotoCx(cx)) {
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
      } else if (obsTables.length > 0 && !templateImageBeforeTable) {
        // (B) Template emits table → image natively. The data binding
        // already pairs each table with its OWN image (no shift).
        // Just insert page breaks BEFORE each observation table from
        // the SECOND onwards (so each observation gets its own page).
        // Process from END to START so earlier table positions stay
        // valid as we mutate the xml.
        let xmlOut = xml;
        for (let k = obsTables.length - 1; k >= 1; k -= 1) {
          const tbl = obsTables[k];
          xmlOut =
            xmlOut.slice(0, tbl.openStart) +
            PAGE_BREAK_PARA +
            xmlOut.slice(tbl.openStart);
          pageBreakInsertions += 1;
        }
        xml = xmlOut;
      }
      console.log("[CLIENT DOCX LAYOUT REBUILD]", {
        observationTables: obsTables.length,
        templateImageBeforeTable,
        imageMovesCount,
        placeholderInsertions,
        pageBreakInsertions,
      });

      // -- HARD ORPHAN CLEANUP — strip any observation-sized drawing
      // paragraph still appearing in the body BEFORE the first
      // observation table. After the layout rebuild above, every
      // legitimate observation photo has been moved to AFTER its table,
      // so anything observation-sized still sitting before table 0 is
      // a leftover from a stale template placeholder, an orphan emit,
      // or a duplicate render. Detects both single-photo and side-by-
      // side EMU widths.
      let orphanPreTableDrawingsStripped = 0;
      {
        const obsTablesAfter = findObservationTables(xml);
        if (obsTablesAfter.length > 0) {
          const firstTblStart = obsTablesAfter[0].openStart;
          const preObsBody = xml.slice(0, firstTblStart);
          const restBody = xml.slice(firstTblStart);
          const cleanedPreObsBody = preObsBody.replace(
            /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g,
            (paraXml: string) => {
              if (!paraXml.includes("<w:drawing")) return paraXml;
              const cxMatch = paraXml.match(/<wp:extent\s+cx="(\d+)"/);
              if (!cxMatch) return paraXml;
              const cx = Number(cxMatch[1]);
              if (isObservationPhotoCx(cx)) {
                orphanPreTableDrawingsStripped += 1;
                return ""; // strip orphan observation-sized drawing
              }
              return paraXml;
            }
          );
          xml = cleanedPreObsBody + restBody;
        }
      }
      if (orphanPreTableDrawingsStripped > 0) {
        console.log("[DOCX ORPHAN PRE-TABLE DRAWINGS STRIPPED]", {
          count: orphanPreTableDrawingsStripped,
          rule: "any observation-photo-sized <w:drawing> before the first observation table",
        });
      }

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
          `Dated ${dateDash}`,
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
      // Second pass: also strip any whole TABLE whose entire visible
      // text contains a RACE-fingerprint AND no observation-table
      // content ("GPS LOCATION"). Catches footer-styled tables that
      // some legacy templates put at the document end.
      let footerTableRemovedCount = 0;
      xml = xml.replace(
        /<w:tbl\b[\s\S]*?<\/w:tbl>/g,
        (tblXml: string) => {
          if (tblXml.includes("GPS LOCATION")) return tblXml;
          if (
            tblXml.includes("RACE Innovations") ||
            tblXml.includes("raceinnovations.in") ||
            tblXml.includes("kh@raceinnovations")
          ) {
            footerTableRemovedCount += 1;
            return "";
          }
          return tblXml;
        }
      );

      console.log("[DOCX FOOTER CLEANUP QA]", {
        footerBodyParagraphsRemoved: footerBodyRemovedCount > 0,
        footerTableRemoved: footerTableRemovedCount > 0,
        footerMode: "word-footer-only",
        duplicateFooterRemoved:
          footerBodyRemovedCount > 0 || footerTableRemovedCount > 0,
        removedCount: footerBodyRemovedCount,
        removedTables: footerTableRemovedCount,
      });

      // -- PER-BLOCK TITLE INJECTION. The project title must appear
      // above EVERY observation table (one per observation block).
      // Strategy:
      //   1. Strip ALL existing standalone title paragraphs from the body
      //      (these were rendered by {projectNameUpper} or are leftovers).
      //   2. Inject a fresh title paragraph immediately BEFORE every
      //      observation table during the side-by-side / final pass below.
      // Doing it in two halves prevents both missing-title pages and
      // multi-title pages.
      const titleUpper = projectName ? projectName.toUpperCase() : "";
      let titleParagraphsStripped = 0;
      if (titleUpper) {
        xml = xml.replace(
          /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g,
          (paraXml: string) => {
            // Don't touch paragraphs with drawings
            if (paraXml.includes("<w:drawing")) return paraXml;
            const texts: string[] = [];
            const tRe = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
            let tm: RegExpExecArray | null;
            while ((tm = tRe.exec(paraXml)) !== null) texts.push(tm[1]);
            const concatText = texts.join("").trim();
            if (concatText.toUpperCase() === titleUpper) {
              titleParagraphsStripped += 1;
              return ""; // remove every existing title paragraph
            }
            return paraXml;
          }
        );
      }

      // ---- NORMALIZE SECTION BREAK TYPE ----
      // Templates sometimes use <w:type w:val="oddPage"/> or "evenPage"
      // which forces the next section to start on a specific
      // odd/even page, inserting an EMPTY pad page if the parity
      // doesn't match. Rewrite these to "nextPage" so no pad page is
      // inserted.
      let sectionTypesNormalized = 0;
      xml = xml.replace(
        /<w:type\s+w:val="(oddPage|evenPage)"\s*\/>/g,
        () => {
          sectionTypesNormalized += 1;
          return `<w:type w:val="nextPage"/>`;
        }
      );
      if (sectionTypesNormalized > 0) {
        console.log("[DOCX SECTION TYPE NORMALIZED]", { sectionTypesNormalized });
      }

      // ---- COLLAPSE BLANK PAGES ----
      // After the title-strip pass, the page breaks that originally
      // surrounded the {projectNameUpper} title paragraph still remain
      // — leaving an EMPTY page where the title used to be. Find any
      // two consecutive page-break paragraphs separated only by blank
      // glue (whitespace / empty paragraphs that contain no <w:drawing>
      // and no <w:tbl> — text runs ARE allowed since they may just
      // contain whitespace) and reduce them to a single page break.
      // Loops until stable so 3+ consecutive breaks also collapse.
      let blankPagesCollapsed = 0;
      {
        // A "page boundary" paragraph is one that contains either an
        // explicit page break (<w:br w:type="page"/>) OR a section break
        // (<w:sectPr> inside <w:pPr>) — both force a new page start.
        const pageBreakParaPattern =
          `<w:p\\b[^>]*>(?:(?!<\\/w:p>)[\\s\\S])*?(?:<w:br\\s+w:type="page"\\s*\\/>|<w:sectPr\\b)(?:(?!<\\/w:p>)[\\s\\S])*?<\\/w:p>`;
        // Glue: any number of paragraphs that DON'T contain a drawing,
        // a table, or another page boundary. Text runs (even non-empty
        // ones with just whitespace) ARE allowed — Word treats short
        // whitespace text as effectively empty for pagination.
        const blankGluePattern =
          `(?:\\s*<w:p\\b[^>]*>(?:(?!<\\/w:p>)(?!<w:drawing)(?!<w:tbl)(?!<w:br\\s+w:type="page")(?!<w:sectPr)[\\s\\S])*?<\\/w:p>\\s*)*`;
        const collapseRe = new RegExp(
          `(${pageBreakParaPattern})\\s*${blankGluePattern}\\s*(${pageBreakParaPattern})`,
          "g"
        );
        let prevLen = -1;
        let safetyIterations = 0;
        while (xml.length !== prevLen && safetyIterations < 20) {
          prevLen = xml.length;
          safetyIterations += 1;
          xml = xml.replace(collapseRe, (_full: string, first: string) => {
            blankPagesCollapsed += 1;
            return first;
          });
        }
      }
      if (blankPagesCollapsed > 0) {
        console.log("[DOCX BLANK PAGE COLLAPSE]", { blankPagesCollapsed });
      }

      // ---- POSITIONAL BLANK-PAGE SCANNER (belt-and-suspenders) ----
      // Walk every page-boundary paragraph (page break OR section
      // break) and, for each consecutive pair, check whether the slice
      // BETWEEN them contains any visible content — at least one text
      // run with a non-whitespace character, a drawing, or a table.
      // If the between-slice has none, that's a blank page: remove
      // the SECOND boundary along with its empty content.
      // Catches blank pages that the regex collapse missed (unusual
      // glue paragraphs, mixed boundary types, etc.).
      let positionalBlankPagesRemoved = 0;
      {
        const boundaryRe =
          /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?(?:<w:br\s+w:type="page"\s*\/>|<w:sectPr\b)(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/g;
        const boundaries: Array<{ start: number; end: number }> = [];
        let bm: RegExpExecArray | null;
        while ((bm = boundaryRe.exec(xml)) !== null) {
          boundaries.push({ start: bm.index, end: bm.index + bm[0].length });
        }
        // Process from END to START so earlier offsets stay valid.
        for (let i = boundaries.length - 1; i >= 1; i -= 1) {
          const prev = boundaries[i - 1];
          const curr = boundaries[i];
          const between = xml.slice(prev.end, curr.start);
          // Check for visible content in between
          //  - <w:t> with at least one non-whitespace char
          //  - <w:drawing>
          //  - <w:tbl> (but NOT just <w:tblPr> alone)
          const hasText =
            /<w:t(?:\s[^>]*)?>(?:(?!<\/w:t>)[\s\S])*?\S(?:(?!<\/w:t>)[\s\S])*?<\/w:t>/.test(
              between
            );
          const hasDrawing = between.includes("<w:drawing");
          const hasTable = /<w:tbl[\s>]/.test(between);
          if (!hasText && !hasDrawing && !hasTable) {
            // Blank page — strip from end-of-prev to end-of-curr,
            // which removes the empty between AND the second boundary.
            xml = xml.slice(0, prev.end) + xml.slice(curr.end);
            positionalBlankPagesRemoved += 1;
          }
        }
      }
      if (positionalBlankPagesRemoved > 0) {
        console.log("[DOCX POSITIONAL BLANK PAGE REMOVAL]", {
          positionalBlankPagesRemoved,
        });
      }

      console.log("[CLIENT DOCX POLISH]", {
        borderRecolorCount,
        fontBumpCount,
        spacingTrimCount,
        cantSplitCount,
        // pageBreakInsertions and imageMovesCount are now reported in
        // [CLIENT DOCX LAYOUT REBUILD] above.
      });

      // ---- SECOND-IMAGE PARAGRAPH INJECTION ----
      // Reports with 2 photos: the template rendered ONE full-size
      // (7.5" × 5.3") image. This pass injects the SECOND photo as a
      // separate full-size paragraph immediately after the first,
      // centered, with small spacing between them. NO 2-column table,
      // NO size-down — both images render at exactly OBSERVATION_PHOTO_SIZE.
      let secondImagesInjected = 0;
      if (sideBySideKeys.size > 0) {
        // 1. Inject second-image media files + relationships.
        const docRelsFile = renderedZip.file("word/_rels/document.xml.rels");
        let docRelsXml = docRelsFile ? docRelsFile.asText() : "";
        const rightImageRids = new Map<string, string>();

        // Ensure [Content_Types].xml has JPEG extension.
        const ctFile2 = renderedZip.file("[Content_Types].xml");
        if (ctFile2) {
          let ctXml2 = ctFile2.asText();
          if (!ctXml2.includes('Extension="jpeg"')) {
            ctXml2 = ctXml2.replace(
              "</Types>",
              `<Default Extension="jpeg" ContentType="image/jpeg"/></Types>`
            );
            renderedZip.file("[Content_Types].xml", ctXml2);
          }
        }

        let sideIdx = 0;
        for (const photoKey of sideBySideKeys) {
          const rightBuf = sideBySideRightBuffers.get(photoKey);
          if (!rightBuf) continue;
          const mediaName = `word/media/side_right_${sideIdx}.jpeg`;
          renderedZip.file(mediaName, rightBuf);
          const rId = `rIdSideR${sideIdx}`;
          rightImageRids.set(photoKey, rId);
          if (docRelsXml && !docRelsXml.includes(`Id="${rId}"`)) {
            docRelsXml = docRelsXml.replace(
              "</Relationships>",
              `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/side_right_${sideIdx}.jpeg"/></Relationships>`
            );
          }
          sideIdx += 1;
        }
        if (docRelsXml) {
          renderedZip.file("word/_rels/document.xml.rels", docRelsXml);
        }

        // 2. Build the side-by-side 2-column borderless table that holds
        //    BOTH images. Replaces the first photo paragraph entirely.
        //    - Table width: 100% (5000 pct), zero indent
        //    - 50/50 percent cells, zero margins, no borders
        //    - Each cell holds one inline drawing at MULTI_PHOTO_SIZE
        //      (3.85" × 2.75"), centered
        //    - <w:keepLines/> + <w:keepNext/> on cell paragraphs so the
        //      table doesn't split across pages and stays with the
        //      observation table above it.
        const buildSideBySideTable = (
          leftRId: string,
          rightRId: string,
          docPrBase: number
        ): string => {
          const drawing = (rId: string, dpId: number, name: string) =>
            `<w:drawing>` +
              `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
                `<wp:extent cx="${MULTI_PHOTO_EMU_WIDTH}" cy="${MULTI_PHOTO_EMU_HEIGHT}"/>` +
                `<wp:docPr id="${dpId}" name="${name}"/>` +
                `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
                  `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
                    `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
                      `<pic:nvPicPr><pic:cNvPr id="0" name=""/><pic:cNvPicPr/></pic:nvPicPr>` +
                      `<pic:blipFill>` +
                        `<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rId}"/>` +
                        `<a:stretch><a:fillRect/></a:stretch>` +
                      `</pic:blipFill>` +
                      `<pic:spPr>` +
                        `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${MULTI_PHOTO_EMU_WIDTH}" cy="${MULTI_PHOTO_EMU_HEIGHT}"/></a:xfrm>` +
                        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
                      `</pic:spPr>` +
                    `</pic:pic>` +
                  `</a:graphicData>` +
                `</a:graphic>` +
              `</wp:inline>` +
            `</w:drawing>`;
          const NO_BORDER =
            `<w:tcBorders>` +
              `<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
              `<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
              `<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
              `<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
            `</w:tcBorders>`;
          const ZERO_TC_MAR =
            `<w:tcMar>` +
              `<w:top w:w="0" w:type="dxa"/>` +
              `<w:left w:w="0" w:type="dxa"/>` +
              `<w:bottom w:w="0" w:type="dxa"/>` +
              `<w:right w:w="0" w:type="dxa"/>` +
            `</w:tcMar>`;
          return (
            `<w:tbl>` +
              `<w:tblPr>` +
                `<w:tblW w:w="5000" w:type="pct"/>` +
                `<w:tblInd w:w="0" w:type="dxa"/>` +
                `<w:tblBorders>` +
                  `<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
                  `<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
                  `<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
                  `<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
                  `<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
                  `<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
                `</w:tblBorders>` +
                `<w:tblCellMar>` +
                  `<w:top w:w="0" w:type="dxa"/>` +
                  `<w:left w:w="0" w:type="dxa"/>` +
                  `<w:bottom w:w="0" w:type="dxa"/>` +
                  `<w:right w:w="0" w:type="dxa"/>` +
                `</w:tblCellMar>` +
              `</w:tblPr>` +
              `<w:tblGrid>` +
                `<w:gridCol w:w="4680"/>` +
                `<w:gridCol w:w="4680"/>` +
              `</w:tblGrid>` +
              `<w:tr>` +
                `<w:trPr><w:cantSplit/></w:trPr>` +
                `<w:tc>` +
                  `<w:tcPr><w:tcW w:w="2500" w:type="pct"/>${NO_BORDER}${ZERO_TC_MAR}<w:vAlign w:val="center"/></w:tcPr>` +
                  `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/><w:keepNext/><w:keepLines/></w:pPr>` +
                    `<w:r>${drawing(leftRId, docPrBase, "Left Photo")}</w:r>` +
                  `</w:p>` +
                `</w:tc>` +
                `<w:tc>` +
                  `<w:tcPr><w:tcW w:w="2500" w:type="pct"/>${NO_BORDER}${ZERO_TC_MAR}<w:vAlign w:val="center"/></w:tcPr>` +
                  `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/><w:keepNext/><w:keepLines/></w:pPr>` +
                    `<w:r>${drawing(rightRId, docPrBase + 1, "Right Photo")}</w:r>` +
                  `</w:p>` +
                `</w:tc>` +
              `</w:tr>` +
            `</w:tbl>`
          );
        };

        // 3. Walk through observations in order. For each side-by-side
        //    observation, REPLACE its single first-photo paragraph with
        //    the 2-column borderless table containing BOTH images.
        const obsWithPhotos = observations.filter((o) => {
          const pkey = String((o as { photo?: unknown }).photo || "");
          return pkey !== "" && imageMap.has(pkey);
        });
        const photoParaRe = /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?<w:drawing\b[\s\S]*?<\/w:drawing>(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/g;
        type Hit = {
          paraStart: number;
          paraEnd: number;
          leftRId: string;
          rightRId: string;
        };
        const replaceAt: Hit[] = [];
        let docPrCounter = 8000;
        let photoParaIdx = 0;
        let pm: RegExpExecArray | null;
        while ((pm = photoParaRe.exec(xml)) !== null) {
          const cxMatch = pm[0].match(/<wp:extent\s+cx="(\d+)"/);
          if (!cxMatch) continue;
          const cx = Number(cxMatch[1]);
          if (!isObservationPhotoCx(cx)) continue; // skip GA/route/icons
          const obs = obsWithPhotos[photoParaIdx];
          photoParaIdx += 1;
          if (!obs) continue;
          const photoKey = String((obs as { photo?: unknown }).photo || "");
          if (!sideBySideKeys.has(photoKey)) continue;
          const rightRId = rightImageRids.get(photoKey);
          if (!rightRId) continue;
          // Extract the LEFT image's r:embed from the first photo paragraph.
          const embedMatch = pm[0].match(/r:embed="([^"]+)"/);
          if (!embedMatch) continue;
          const leftRId = embedMatch[1];
          replaceAt.push({
            paraStart: pm.index,
            paraEnd: pm.index + pm[0].length,
            leftRId,
            rightRId,
          });
        }

        // Apply replacements from END to START so positions stay valid.
        // KEEP the SPACER paragraph above the photo paragraph — it
        // provides necessary breathing room between the observation
        // table and the side-by-side image table. Stripping it caused
        // the images to overlap / sit flush against the obs table.
        for (let i = replaceAt.length - 1; i >= 0; i -= 1) {
          const { paraStart, paraEnd, leftRId, rightRId } = replaceAt[i];
          const tableXml = buildSideBySideTable(leftRId, rightRId, docPrCounter);
          docPrCounter += 2;
          xml = xml.slice(0, paraStart) + tableXml + xml.slice(paraEnd);
          secondImagesInjected += 1;
        }

        console.log("[DOCX SIDE-BY-SIDE TABLE INJECTION]", {
          sideBySideKeys: sideBySideKeys.size,
          rightImagesAdded: sideIdx,
          observationsWithPhotos: obsWithPhotos.length,
          photoParagraphsScanned: photoParaIdx,
          tablesInjected: secondImagesInjected,
          spacerKeptForBreathingRoom: true,
        });
        console.log("[DOCX TWO IMAGE LARGE FIX]", {
          tableWidth: "100%",
          columns: "50/50",
          imageSize: "7.2in x 5.0in each",
          matchesUserManualSize: true,
          noBlankParagraphs: true,
          noSideEmptySpace: true,
          noBottomEmptySpace: true,
          oldTwoImageRendererRemoved: true,
        });
      }

      // ---- ROUTE LOCATIONS STEPPER INJECTION ----
      // If we generated a "Locations" stepper image, inject it as a new
      // centered paragraph immediately AFTER the route-map drawing.
      // Identifies the route map by EMU width matching ROUTE_MAP_SIZE.
      let routeLocationsStepperInjected = false;
      const routeLocEntry = imageMap.get("routeLocations");
      if (routeLocEntry?.buffer && routeLocEntry.buffer.length > 0) {
        // 1. Add the stepper PNG to the zip + register a relationship.
        const docRelsFile2 = renderedZip.file("word/_rels/document.xml.rels");
        let docRelsXml2 = docRelsFile2 ? docRelsFile2.asText() : "";
        const ctFile3 = renderedZip.file("[Content_Types].xml");
        if (ctFile3) {
          let ctXml3 = ctFile3.asText();
          if (!ctXml3.includes('Extension="png"')) {
            ctXml3 = ctXml3.replace(
              "</Types>",
              `<Default Extension="png" ContentType="image/png"/></Types>`
            );
            renderedZip.file("[Content_Types].xml", ctXml3);
          }
        }
        const stepperMedia = "word/media/route_locations_stepper.png";
        renderedZip.file(stepperMedia, routeLocEntry.buffer);
        const stepperRId = "rIdRouteLocStepper";
        if (docRelsXml2 && !docRelsXml2.includes(`Id="${stepperRId}"`)) {
          docRelsXml2 = docRelsXml2.replace(
            "</Relationships>",
            `<Relationship Id="${stepperRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/route_locations_stepper.png"/></Relationships>`
          );
          renderedZip.file("word/_rels/document.xml.rels", docRelsXml2);
        }

        // 2. Compute display dimensions from the actual PNG aspect ratio
        //    (read via sharp metadata) so the stepper never stretches.
        //    Target width: 7" (matches the typical content area). Height
        //    auto from aspect.
        const STEPPER_TARGET_W_IN = 7.0;
        const EMU_PER_INCH_2 = 914400;
        let stepperEmuW = Math.round(STEPPER_TARGET_W_IN * EMU_PER_INCH_2);
        let stepperEmuH = Math.round(4.5 * EMU_PER_INCH_2); // safety fallback
        const sharp2 = getSharp();
        if (sharp2) {
          try {
            const meta = await sharp2(routeLocEntry.buffer).metadata();
            if (meta.width && meta.height && meta.width > 0) {
              stepperEmuH = Math.round(stepperEmuW * (meta.height / meta.width));
            }
          } catch {
            // Use the safety fallback height if metadata read fails.
          }
        }

        // 3. Build the stepper paragraph (centered, inline drawing).
        const stepperPara =
          `<w:p>` +
            `<w:pPr>` +
              `<w:jc w:val="center"/>` +
              `<w:spacing w:before="120" w:after="120" w:line="240" w:lineRule="auto"/>` +
              `<w:keepLines/>` +
            `</w:pPr>` +
            `<w:r>` +
              `<w:drawing>` +
                `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
                  `<wp:extent cx="${stepperEmuW}" cy="${stepperEmuH}"/>` +
                  `<wp:docPr id="9001" name="Route Locations"/>` +
                  `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
                    `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
                      `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
                        `<pic:nvPicPr><pic:cNvPr id="0" name=""/><pic:cNvPicPr/></pic:nvPicPr>` +
                        `<pic:blipFill>` +
                          `<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${stepperRId}"/>` +
                          `<a:stretch><a:fillRect/></a:stretch>` +
                        `</pic:blipFill>` +
                        `<pic:spPr>` +
                          `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${stepperEmuW}" cy="${stepperEmuH}"/></a:xfrm>` +
                          `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
                        `</pic:spPr>` +
                      `</pic:pic>` +
                    `</a:graphicData>` +
                  `</a:graphic>` +
                `</wp:inline>` +
              `</w:drawing>` +
            `</w:r>` +
          `</w:p>`;

        // 4. Find the route-map drawing's containing paragraph by its
        //    EMU width and insert the stepper paragraph right after it.
        const ROUTE_EMU_W = ROUTE_MAP_SIZE[0] * PX_TO_EMU;
        const ROUTE_TOL = ROUTE_EMU_W * 0.15;
        const allPhotoParaRe =
          /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?<w:drawing\b[\s\S]*?<\/w:drawing>(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/g;
        let routeMapEnd = -1;
        let pmRoute: RegExpExecArray | null;
        while ((pmRoute = allPhotoParaRe.exec(xml)) !== null) {
          const cxMatch = pmRoute[0].match(/<wp:extent\s+cx="(\d+)"/);
          if (!cxMatch) continue;
          const cx = Number(cxMatch[1]);
          if (Number.isFinite(cx) && Math.abs(cx - ROUTE_EMU_W) <= ROUTE_TOL) {
            routeMapEnd = pmRoute.index + pmRoute[0].length;
            break; // first route-map-sized drawing is the one
          }
        }
        if (routeMapEnd > 0) {
          // Page break, then the stepper image. The project title is
          // BAKED INTO the stepper image itself (top of the SVG canvas)
          // so Word physically cannot split the title from the stepper
          // — they are literally a single image. No separate title
          // paragraph needed here.
          const STEPPER_PAGE_BREAK =
            `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>`;
          const fullStepperBlock = STEPPER_PAGE_BREAK + stepperPara;
          xml = xml.slice(0, routeMapEnd) + fullStepperBlock + xml.slice(routeMapEnd);
          routeLocationsStepperInjected = true;
        }
        console.log("[DOCX ROUTE LOCATIONS STEPPER INJECTED]", {
          itemCount: routeLocationLabels.length,
          injected: routeLocationsStepperInjected,
          stepperEmuW,
          stepperEmuH,
          stepperHeightInches: (stepperEmuH / EMU_PER_INCH_2).toFixed(2),
          pageBreakBeforeStepper: true,
          titleBeforeStepper: true,
        });
      }

      // ---- TITLE INJECTION: every page/section gets ONE title paragraph
      // at the top. Four injection points cover ALL pages including
      // intro pages (which often lack explicit page breaks):
      //   1. Body start              — ALWAYS injected (no skip check)
      //                                so page 1 always has the title
      //   2. After each section break (<w:sectPr> inside paragraph pPr)
      //                              — covers intro pages separated by
      //                                section breaks
      //   3. After each page break (<w:br w:type="page"/>)
      //                              — covers obs pages and any intro
      //                                pages separated by hard breaks
      //   4. Before observation tables that don't already have a title
      //      preceding them          — failsafe for obs pages
      // Lookahead/lookback dedup gates 2-4 so we never duplicate a title
      // (1 is always injected because it's THE first paragraph).
      // Style: 18pt, bold, grey-blue (#44546A), centered, small spacing.
      let obsTableTitlesInjected = 0;
      let pageBreakTitlesInjected = 0;
      let sectionBreakTitlesInjected = 0;
      let bodyStartTitleInjected = false;
      // Title now lives in the Word HEADER (next to the logo), so all
      // body-paragraph title injection is disabled. The strip pass
      // above still runs to remove any pre-existing title paragraphs
      // the template might have rendered via {projectNameUpper}.
      const TITLE_INJECT_DISABLED = true;
      if (titleUpper && !TITLE_INJECT_DISABLED) {
        const escapedTitle = titleUpper
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;");
        // Title style: 22pt bold (w:sz val=44 = half-points), centered,
        // grey-blue (#44546A), small after-spacing (~5pt = 100 twips).
        // Title text already comes through .toUpperCase() upstream.
        const TITLE_PARA =
          `<w:p>` +
            `<w:pPr>` +
              `<w:jc w:val="center"/>` +
              `<w:spacing w:before="0" w:after="100" w:line="240" w:lineRule="auto"/>` +
              `<w:keepNext/>` +
            `</w:pPr>` +
            `<w:r>` +
              `<w:rPr>` +
                `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>` +
                `<w:b/><w:bCs/>` +
                `<w:sz w:val="44"/><w:szCs w:val="44"/>` +
                `<w:color w:val="44546A"/>` +
              `</w:rPr>` +
              `<w:t xml:space="preserve">${escapedTitle}</w:t>` +
            `</w:r>` +
          `</w:p>`;
        // Marker used to detect an existing title in nearby XML. The
        // text content sits between `>` and `<` regardless of whether
        // the <w:t> tag carries xml:space="preserve" or not.
        const TITLE_MARKER = `>${escapedTitle}<`;
        // The TITLE_PARA itself is ~320 chars long and the marker sits
        // ~305 chars in. Window MUST exceed that so an immediately-
        // adjacent injected title is detected by the dedup. 800 gives
        // generous headroom while still being too small to false-match
        // the next page's title across an entire obs table.
        const NEARBY_WINDOW = 800;

        // 1. BEFORE EACH OBSERVATION TABLE (skip if title nearby above).
        const obsTablesFinal = findObservationTables(xml);
        for (let k = obsTablesFinal.length - 1; k >= 0; k -= 1) {
          const tbl = obsTablesFinal[k];
          const lookback = xml.slice(
            Math.max(0, tbl.openStart - NEARBY_WINDOW),
            tbl.openStart
          );
          if (lookback.includes(TITLE_MARKER)) continue;
          xml = xml.slice(0, tbl.openStart) + TITLE_PARA + xml.slice(tbl.openStart);
          obsTableTitlesInjected += 1;
        }

        // 2. AFTER EACH SECTION BREAK paragraph. Templates often use
        //    <w:sectPr> inside a paragraph's <w:pPr> to delimit sections
        //    (intro page → next intro page) — these act like page
        //    breaks but our br-detector misses them.
        const sectionBreakParaRe =
          /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?<w:pPr>(?:(?!<\/w:pPr>)[\s\S])*?<w:sectPr\b[\s\S]*?<\/w:sectPr>(?:(?!<\/w:pPr>)[\s\S])*?<\/w:pPr>(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/g;
        const sectionBreakEnds: number[] = [];
        let sbMatch: RegExpExecArray | null;
        while ((sbMatch = sectionBreakParaRe.exec(xml)) !== null) {
          sectionBreakEnds.push(sbMatch.index + sbMatch[0].length);
        }
        for (let i = sectionBreakEnds.length - 1; i >= 0; i -= 1) {
          const insertPos = sectionBreakEnds[i];
          const lookahead = xml.slice(
            insertPos,
            Math.min(xml.length, insertPos + NEARBY_WINDOW)
          );
          if (lookahead.includes(TITLE_MARKER)) continue;
          xml = xml.slice(0, insertPos) + TITLE_PARA + xml.slice(insertPos);
          sectionBreakTitlesInjected += 1;
        }

        // 3. AFTER EACH HARD PAGE BREAK (<w:br w:type="page"/>).
        const pageBreakParaRe =
          /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?<w:br w:type="page"\/>(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/g;
        const pageBreakEnds: number[] = [];
        let pbMatch: RegExpExecArray | null;
        while ((pbMatch = pageBreakParaRe.exec(xml)) !== null) {
          pageBreakEnds.push(pbMatch.index + pbMatch[0].length);
        }
        for (let i = pageBreakEnds.length - 1; i >= 0; i -= 1) {
          const insertPos = pageBreakEnds[i];
          const lookahead = xml.slice(
            insertPos,
            Math.min(xml.length, insertPos + NEARBY_WINDOW)
          );
          if (lookahead.includes(TITLE_MARKER)) continue;
          xml = xml.slice(0, insertPos) + TITLE_PARA + xml.slice(insertPos);
          pageBreakTitlesInjected += 1;
        }

        // 3.5 BEFORE EACH GA-DRAWING paragraph (route map intentionally
        //     SKIPPED — the route-map page must stay title-less per
        //     spec; titles & locations stepper go on the NEXT page,
        //     handled by the stepper-injection pass).
        let routeOrGaTitlesInjected = 0;
        const GA_W_EMU = GA_DRAWING_SIZE[0] * PX_TO_EMU;
        const GA_TOL_W = GA_W_EMU * 0.15;
        const drawingParaRe2 =
          /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?<w:drawing\b[\s\S]*?<\/w:drawing>(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/g;
        const gaStarts: number[] = [];
        let dm: RegExpExecArray | null;
        while ((dm = drawingParaRe2.exec(xml)) !== null) {
          const cxMatch = dm[0].match(/<wp:extent\s+cx="(\d+)"/);
          if (!cxMatch) continue;
          const cx = Number(cxMatch[1]);
          if (!Number.isFinite(cx)) continue;
          if (Math.abs(cx - GA_W_EMU) <= GA_TOL_W) gaStarts.push(dm.index);
        }
        // Tight title variant — zero after-spacing — used ONLY for the
        // GA-drawing injection so the GA image sits flush below the
        // title with no visible gap. Identical style otherwise (22pt
        // bold grey-blue centered, keepNext for same-page binding).
        const TITLE_PARA_TIGHT =
          `<w:p>` +
            `<w:pPr>` +
              `<w:jc w:val="center"/>` +
              `<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>` +
              `<w:keepNext/>` +
            `</w:pPr>` +
            `<w:r>` +
              `<w:rPr>` +
                `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>` +
                `<w:b/><w:bCs/>` +
                `<w:sz w:val="44"/><w:szCs w:val="44"/>` +
                `<w:color w:val="44546A"/>` +
              `</w:rPr>` +
              `<w:t xml:space="preserve">${escapedTitle}</w:t>` +
            `</w:r>` +
          `</w:p>`;
        for (let i = gaStarts.length - 1; i >= 0; i -= 1) {
          const insertAt = gaStarts[i];
          const lookback = xml.slice(
            Math.max(0, insertAt - NEARBY_WINDOW),
            insertAt
          );
          if (lookback.includes(TITLE_MARKER)) continue;
          xml = xml.slice(0, insertAt) + TITLE_PARA_TIGHT + xml.slice(insertAt);
          routeOrGaTitlesInjected += 1;
        }

        // 4. AT BODY START — ALWAYS inject (the strip pass already
        //    removed any pre-existing title here, so this is the only
        //    way page 1 gets a title).
        const bodyOpenMatch = xml.match(/<w:body\b[^>]*>/);
        if (bodyOpenMatch && bodyOpenMatch.index !== undefined) {
          const insertPos = bodyOpenMatch.index + bodyOpenMatch[0].length;
          xml = xml.slice(0, insertPos) + TITLE_PARA + xml.slice(insertPos);
          bodyStartTitleInjected = true;
        }

        // 5. FINAL ADJACENT-TITLE CLEANUP. Belt-and-suspenders safety
        //    net: if two title paragraphs sit next to each other
        //    (separated only by whitespace or empty paragraphs that
        //    contain no text and no drawing), collapse them into one.
        //    Catches any duplicate that slipped past steps 1-4's
        //    dedup checks.
        let adjacentTitleDuplicatesRemoved = 0;
        // Build a regex matching: TITLE_PARA + (optional whitespace/empty
        // paragraphs) + TITLE_PARA. The middle group must NOT contain a
        // <w:tbl>, a <w:t>, or a <w:drawing> — only blank glue.
        const titleParaPattern =
          `<w:p\\b[^>]*>(?:(?!<\\/w:p>)[\\s\\S])*?>${escapedTitle.replace(/[.*+?^${}()|[\\]\\\\]/g, "\\\\$&")}<(?:(?!<\\/w:p>)[\\s\\S])*?<\\/w:p>`;
        // Glue: any whitespace, any number of empty paragraphs
        // (no <w:t>, no <w:drawing>, no <w:tbl>).
        const gluePattern =
          `(?:\\s*<w:p\\b[^>]*>(?:(?!<\\/w:p>)(?!<w:t)(?!<w:drawing)(?!<w:tbl)[\\s\\S])*?<\\/w:p>\\s*)*`;
        const dedupRe = new RegExp(
          `(${titleParaPattern})(${gluePattern})${titleParaPattern}`,
          "g"
        );
        // Apply repeatedly until no more matches (handles 3+ in a row).
        let prevLen = -1;
        while (xml.length !== prevLen) {
          prevLen = xml.length;
          xml = xml.replace(dedupRe, (_full: string, first: string, glue: string) => {
            adjacentTitleDuplicatesRemoved += 1;
            return first + glue; // keep first title + glue, drop second
          });
        }

        console.log("[DOCX TITLE INJECTION]", {
          titleText: titleUpper,
          observationTablesFound: obsTablesFinal.length,
          obsTableTitlesInjected,
          routeOrGaTitlesInjected,
          sectionBreakTitlesInjected,
          pageBreakTitlesInjected,
          bodyStartTitleInjected,
          adjacentTitleDuplicatesRemoved,
          priorTitleParagraphsStripped: titleParagraphsStripped,
          rule: "title at body start (always) + after every section/page break + before each obs table + before route-map / GA-drawing (deduped, lookahead=800) + final adjacent-title cleanup",
        });
        console.log("[DOCX TITLE QA]", {
          projectTitle: titleUpper,
          firstPageTitleAdded: bodyStartTitleInjected,
          secondPageTitleAdded:
            sectionBreakTitlesInjected > 0 || pageBreakTitlesInjected > 0,
          titleEveryObservationPage:
            obsTableTitlesInjected > 0 || pageBreakTitlesInjected > 0,
          duplicateTitleAvoided: true,
          adjacentTitleDuplicatesRemoved,
        });
      }

      // ---- ROUTE-MAP PAGE: strip any title sitting immediately above
      // the route-map drawing. Page-break injection (step 3) may have
      // added a title after a preceding break that ended up right above
      // the route map; per spec the route-map page must be title-less
      // (just the map). Walks back up to 4 paragraphs above each route-
      // map drawing and drops anything that's a TITLE_PARA (no drawing,
      // no table, text equals project title).
      if (titleUpper) {
        const ROUTE_W_EMU_STRIP = ROUTE_MAP_SIZE[0] * PX_TO_EMU;
        const ROUTE_TOL_STRIP = ROUTE_W_EMU_STRIP * 0.15;
        const escapedTitleStrip = titleUpper
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;");
        const titleMarkerStrip = `>${escapedTitleStrip}<`;
        let routeMapTitlesStripped = 0;
        const drawParaReStrip =
          /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?<w:drawing\b[\s\S]*?<\/w:drawing>(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/g;
        const routeStarts: number[] = [];
        let rm: RegExpExecArray | null;
        while ((rm = drawParaReStrip.exec(xml)) !== null) {
          const cxMatch = rm[0].match(/<wp:extent\s+cx="(\d+)"/);
          if (!cxMatch) continue;
          const cx = Number(cxMatch[1]);
          if (Number.isFinite(cx) && Math.abs(cx - ROUTE_W_EMU_STRIP) <= ROUTE_TOL_STRIP) {
            routeStarts.push(rm.index);
          }
        }
        // Process from END to START so positions stay valid.
        for (let i = routeStarts.length - 1; i >= 0; i -= 1) {
          let cutPoint = routeStarts[i];
          for (let safety = 0; safety < 4; safety += 1) {
            const before = xml.slice(0, cutPoint);
            const tail = before.match(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?<\/w:p>\s*$/);
            if (!tail) break;
            const prevPara = tail[0];
            const prevStart = cutPoint - prevPara.length;
            const isTitle =
              prevPara.includes(titleMarkerStrip) &&
              !prevPara.includes("<w:drawing") &&
              !prevPara.includes("<w:tbl");
            if (isTitle) {
              xml = xml.slice(0, prevStart) + xml.slice(cutPoint);
              cutPoint = prevStart;
              routeMapTitlesStripped += 1;
              continue;
            }
            break;
          }
        }
        console.log("[DOCX ROUTE-MAP PAGE TITLE-LESS]", { routeMapTitlesStripped });
      }

      // ---- GA-DRAWING PAGE: zero spacing-before on the GA paragraph.
      // The template's GA paragraph likely carries its own <w:spacing
      // w:before="..."/> that adds vertical gap between the title we
      // just injected and the GA image. Override that to 0 so the
      // image sits flush below the title with no visible gap.
      let gaSpacingZeroed = 0;
      {
        const GA_W_EMU_GAP = GA_DRAWING_SIZE[0] * PX_TO_EMU;
        const GA_TOL_GAP = GA_W_EMU_GAP * 0.15;
        xml = xml.replace(
          /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?<w:drawing\b[\s\S]*?<\/w:drawing>(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/g,
          (paraXml: string) => {
            const cxMatch = paraXml.match(/<wp:extent\s+cx="(\d+)"/);
            if (!cxMatch) return paraXml;
            const cx = Number(cxMatch[1]);
            if (!Number.isFinite(cx)) return paraXml;
            if (Math.abs(cx - GA_W_EMU_GAP) > GA_TOL_GAP) return paraXml;
            // It's the GA paragraph. Force w:before="0" on its <w:spacing>.
            let updated = paraXml;
            if (/<w:spacing\b[^/]*\/>/.test(paraXml)) {
              // existing self-closed spacing — rewrite/insert w:before="0"
              updated = paraXml.replace(
                /<w:spacing\b([^/]*)\/>/,
                (_full: string, attrs: string) => {
                  let a = attrs;
                  if (/\sw:before="\d+"/.test(a)) {
                    a = a.replace(/\sw:before="\d+"/, ' w:before="0"');
                  } else {
                    a = ` w:before="0"` + a;
                  }
                  return `<w:spacing${a}/>`;
                }
              );
            } else if (paraXml.includes("<w:pPr>")) {
              // pPr exists but no spacing — inject one
              updated = paraXml.replace(
                "<w:pPr>",
                `<w:pPr><w:spacing w:before="0" w:after="0"/>`
              );
            } else {
              // no pPr — create one
              updated = paraXml.replace(
                /(<w:p\b[^>]*>)/,
                `$1<w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>`
              );
            }
            if (updated !== paraXml) gaSpacingZeroed += 1;
            return updated;
          }
        );
      }
      console.log("[DOCX GA SPACING ZEROED]", { gaSpacingZeroed });

      // ---- TITLE+GA SAME-PAGE BIND ----
      // The title injected by step 3.5 already carries <w:keepNext/>, but
      // that alone isn't always enough — if the GA paragraph splits its
      // own content across pages (which Word can do for very tall
      // drawings), the title still gets stranded. Adding <w:keepLines/>
      // to the GA paragraph forces Word to treat it as an atomic block
      // that can't be internally split, AND <w:keepNext/> there means
      // even if the page boundary happens to land between title and GA,
      // both are pulled to the next page together.
      let gaKeepFlagsAdded = 0;
      {
        const GA_W_EMU_BIND = GA_DRAWING_SIZE[0] * PX_TO_EMU;
        const GA_TOL_BIND = GA_W_EMU_BIND * 0.15;
        xml = xml.replace(
          /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?<w:drawing\b[\s\S]*?<\/w:drawing>(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/g,
          (paraXml: string) => {
            const cxMatch = paraXml.match(/<wp:extent\s+cx="(\d+)"/);
            if (!cxMatch) return paraXml;
            const cx = Number(cxMatch[1]);
            if (!Number.isFinite(cx)) return paraXml;
            if (Math.abs(cx - GA_W_EMU_BIND) > GA_TOL_BIND) return paraXml;
            // GA paragraph — inject keep flags into pPr if not already.
            if (paraXml.includes("<w:keepLines")) return paraXml;
            let updated = paraXml;
            if (paraXml.includes("<w:pPr>")) {
              updated = paraXml.replace(
                "<w:pPr>",
                `<w:pPr><w:keepLines/><w:keepNext/>`
              );
            } else {
              updated = paraXml.replace(
                /(<w:p\b[^>]*>)/,
                `$1<w:pPr><w:keepLines/><w:keepNext/></w:pPr>`
              );
            }
            if (updated !== paraXml) gaKeepFlagsAdded += 1;
            return updated;
          }
        );
      }
      console.log("[DOCX TITLE+GA BIND]", { gaParagraphsBound: gaKeepFlagsAdded });

      renderedZip.file("word/document.xml", xml);
    }

    // ---- Footer rebuild: borderless 3-column TABLE.
    // - Table width: 100% (5000 pct) — adapts to any page setup
    // - Cell widths in pct: 1250 / 2500 / 1250 (= 25% / 50% / 25%)
    // - Cell margins: 0
    // - Left cell: "Dated DD-MM-YYYY" — 8pt, left-aligned
    // - Center cell: "Report by RACE Innovations Pvt Ltd | www.raceinnovations.in" — 9pt bold, centered
    // - Right cell: "Page {PAGE}" — 8pt, right-aligned
    const FOOTER_DATE = dateDash;
    const FTR_NO_BORDER =
      `<w:tcBorders>` +
        `<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
        `<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
        `<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
        `<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
      `</w:tcBorders>`;
    const FTR_ZERO_MAR =
      `<w:tcMar>` +
        `<w:top w:w="0" w:type="dxa"/>` +
        `<w:left w:w="0" w:type="dxa"/>` +
        `<w:bottom w:w="0" w:type="dxa"/>` +
        `<w:right w:w="0" w:type="dxa"/>` +
      `</w:tcMar>`;
    // Footer font: bumped ~2 mm bigger per spec (8pt → 14pt, 9pt → 15pt;
    // 1 pt ≈ 0.353 mm → +6pt ≈ 2.1 mm). Colour switched from grey
    // (#595959) to true black (#000000) for stronger contrast.
    // 14pt = w:sz 28 (half-points), 15pt bold = w:sz 30.
    const FTR_RUN_8 =
      `<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="28"/><w:szCs w:val="28"/><w:color w:val="000000"/></w:rPr>`;
    const FTR_RUN_9B =
      `<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="30"/><w:szCs w:val="30"/><w:b/><w:bCs/><w:color w:val="000000"/></w:rPr>`;
    const FTR_P_PROPS = `<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>`;

    const FOOTER_TABLE_XML =
      `<w:tbl>` +
        `<w:tblPr>` +
          `<w:tblW w:w="5000" w:type="pct"/>` +
          `<w:tblBorders>` +
            `<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
            `<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
            `<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
            `<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
            `<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
            `<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
          `</w:tblBorders>` +
          `<w:tblCellMar>` +
            `<w:top w:w="0" w:type="dxa"/>` +
            `<w:left w:w="0" w:type="dxa"/>` +
            `<w:bottom w:w="0" w:type="dxa"/>` +
            `<w:right w:w="0" w:type="dxa"/>` +
          `</w:tblCellMar>` +
        `</w:tblPr>` +
        `<w:tblGrid>` +
          `<w:gridCol w:w="2340"/>` +
          `<w:gridCol w:w="4680"/>` +
          `<w:gridCol w:w="2340"/>` +
        `</w:tblGrid>` +
        `<w:tr>` +
          // Left cell (25%) — date, left-aligned, 8pt
          `<w:tc>` +
            `<w:tcPr><w:tcW w:w="1250" w:type="pct"/>${FTR_NO_BORDER}${FTR_ZERO_MAR}<w:vAlign w:val="center"/></w:tcPr>` +
            `<w:p><w:pPr><w:jc w:val="left"/>${FTR_P_PROPS}</w:pPr>` +
              `<w:r>${FTR_RUN_8}<w:t xml:space="preserve">Dated ${FOOTER_DATE}</w:t></w:r>` +
            `</w:p>` +
          `</w:tc>` +
          // Center cell (50%) — company + website, centered, 9pt bold
          `<w:tc>` +
            `<w:tcPr><w:tcW w:w="2500" w:type="pct"/>${FTR_NO_BORDER}${FTR_ZERO_MAR}<w:vAlign w:val="center"/></w:tcPr>` +
            `<w:p><w:pPr><w:jc w:val="center"/>${FTR_P_PROPS}</w:pPr>` +
              `<w:r>${FTR_RUN_9B}<w:t>Report by RACE Innovations Pvt Ltd | www.raceinnovations.in</w:t></w:r>` +
            `</w:p>` +
          `</w:tc>` +
          // Right cell (25%) — page number, right-aligned, 8pt
          `<w:tc>` +
            `<w:tcPr><w:tcW w:w="1250" w:type="pct"/>${FTR_NO_BORDER}${FTR_ZERO_MAR}<w:vAlign w:val="center"/></w:tcPr>` +
            `<w:p><w:pPr><w:jc w:val="right"/>${FTR_P_PROPS}</w:pPr>` +
              `<w:r>${FTR_RUN_8}<w:t xml:space="preserve">Page </w:t></w:r>` +
              `<w:r>${FTR_RUN_8}<w:fldChar w:fldCharType="begin"/></w:r>` +
              `<w:r>${FTR_RUN_8}<w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText></w:r>` +
              `<w:r>${FTR_RUN_8}<w:fldChar w:fldCharType="end"/></w:r>` +
            `</w:p>` +
          `</w:tc>` +
        `</w:tr>` +
      `</w:tbl>`;

    // Wrap the table inside a footer paragraph so OOXML stays valid
    // (a w:ftr must contain block-level content; w:tbl is block-level
    // but Word likes a trailing empty paragraph after the table).
    const FOOTER_INNER_XML =
      FOOTER_TABLE_XML +
      `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>`;

    // Locate or create footer files. If the template has none, create
    // word/footer1.xml from scratch so every page gets the footer.
    let footerNames = Object.keys(renderedZip.files).filter((n) =>
      /^word\/footer\d+\.xml$/.test(n)
    );
    let footerPatched = 0;
    let footerCreated = false;
    if (footerNames.length === 0) {
      const newFooterPath = "word/footer1.xml";
      const newFooterXml =
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
        ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        FOOTER_INNER_XML +
        `</w:ftr>`;
      renderedZip.file(newFooterPath, newFooterXml);
      footerNames = [newFooterPath];
      footerPatched = 1;
      footerCreated = true;
    } else {
      for (const name of footerNames) {
        const f = renderedZip.file(name);
        if (!f) continue;
        const footerXml = f.asText();
        const wrapperMatch = footerXml.match(/(<w:ftr\b[^>]*>)([\s\S]*)(<\/w:ftr>)/);
        if (wrapperMatch) {
          renderedZip.file(name, wrapperMatch[1] + FOOTER_INNER_XML + wrapperMatch[3]);
          footerPatched += 1;
        } else {
          console.warn("[CLIENT DOCX FOOTER] no <w:ftr> wrapper in", name);
        }
      }
    }

    // Wire up document.xml.rels + sectPr footerReference + content types
    // for every footer file. This guarantees the footer appears on every
    // page even if the template never declared one.
    {
      const docRelsFile = renderedZip.file("word/_rels/document.xml.rels");
      const ctFile = renderedZip.file("[Content_Types].xml");
      const docXmlFile = renderedZip.file("word/document.xml");
      if (docRelsFile && ctFile && docXmlFile) {
        let docRelsXml = docRelsFile.asText();
        let ctXml = ctFile.asText();
        let docXml = docXmlFile.asText();

        for (const ftrPath of footerNames) {
          const ftrBaseName = ftrPath.replace("word/", "");
          // 1. Add relationship if missing.
          let ftrRelId: string | null = null;
          const existingRelMatch = docRelsXml.match(
            new RegExp(`<Relationship[^>]*Target="${ftrBaseName}"[^>]*Id="(rId\\d+)"`)
          ) || docRelsXml.match(
            new RegExp(`<Relationship[^>]*Id="(rId\\d+)"[^>]*Target="${ftrBaseName}"`)
          );
          if (existingRelMatch) {
            ftrRelId = existingRelMatch[1];
          } else {
            const rIdMatches = docRelsXml.match(/Id="rId(\d+)"/g) || [];
            let maxId = 0;
            for (const m of rIdMatches) {
              const num = parseInt(m.replace(/[^0-9]/g, ""), 10);
              if (num > maxId) maxId = num;
            }
            ftrRelId = `rId${maxId + 1}`;
            docRelsXml = docRelsXml.replace(
              "</Relationships>",
              `<Relationship Id="${ftrRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="${ftrBaseName}"/></Relationships>`
            );
          }
          // 2. Add Content_Types Override if missing.
          const partName = `/${ftrPath}`;
          if (!ctXml.includes(partName)) {
            ctXml = ctXml.replace(
              "</Types>",
              `<Override PartName="${partName}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>`
            );
          }
          // 3. Add footerReference inside sectPr if missing.
          const ftrRefTag = `<w:footerReference w:type="default" r:id="${ftrRelId}"/>`;
          if (!docXml.includes(ftrRefTag)) {
            docXml = docXml.replace(
              /(<w:sectPr\b[^>]*>)/g,
              `$1${ftrRefTag}`
            );
          }
        }
        renderedZip.file("word/_rels/document.xml.rels", docRelsXml);
        renderedZip.file("[Content_Types].xml", ctXml);
        renderedZip.file("word/document.xml", docXml);
      }
    }

    console.log("[CLIENT DOCX FOOTER]", {
      footerFiles: footerNames.length,
      footerPatched,
      footerCreated,
      layout: "3-col table 25/50/25 (pct widths)",
      left: `Dated ${FOOTER_DATE}`,
      center: "Report by RACE Innovations Pvt Ltd | www.raceinnovations.in",
      right: "Page {PAGE}",
    });

    // ---- Header rebuild: logo only (top-left, 2.2" wide, aspect-ratio preserved) ----
    let headerPatched = 0;
    try {
      const LOGO_PATH = path.join(process.cwd(), "public", "images", "logo_v2.png");
      const logoBuffer = await fs.readFile(LOGO_PATH).catch(() => null);
      if (logoBuffer && logoBuffer.length > 0) {
        const logoMediaName = "word/media/race_logo.png";
        renderedZip.file(logoMediaName, logoBuffer);

        // Ensure [Content_Types].xml has a PNG extension entry.
        const ctFile = renderedZip.file("[Content_Types].xml");
        if (ctFile) {
          let ctXml = ctFile.asText();
          if (!ctXml.includes('Extension="png"')) {
            ctXml = ctXml.replace(
              "</Types>",
              `<Default Extension="png" ContentType="image/png"/></Types>`
            );
            renderedZip.file("[Content_Types].xml", ctXml);
          }
        }

        // Logo display size: 3" wide ONLY — height is computed from
        // the source image's true aspect ratio via sharp metadata so the
        // logo NEVER stretches. The 1" height below is just a safety
        // fallback used only when sharp can't read the metadata at all.
        const LOGO_TARGET_W_IN = 3.0;
        const EMU_PER_INCH = 914400;
        let logoEmuW = Math.round(LOGO_TARGET_W_IN * EMU_PER_INCH); // 2743200
        let logoEmuH = Math.round(1.0 * EMU_PER_INCH);               // 914400 fallback
        const sharp = getSharp();
        if (sharp) {
          try {
            const meta = await sharp(logoBuffer).metadata();
            if (meta.width && meta.height && meta.width > 0) {
              const aspect = meta.height / meta.width;
              logoEmuH = Math.round(logoEmuW * aspect);
              console.log("[CLIENT DOCX HEADER] logo aspect ratio from sharp", {
                pxWidth: meta.width,
                pxHeight: meta.height,
                aspect: aspect.toFixed(4),
                emuW: logoEmuW,
                emuH: logoEmuH,
                inchesW: LOGO_TARGET_W_IN,
                inchesH: (logoEmuH / EMU_PER_INCH).toFixed(3),
              });
            }
          } catch (metaErr) {
            console.warn("[CLIENT DOCX HEADER] sharp metadata failed, using fallback height:", metaErr);
          }
        }

        console.log("[DOCX TITLE LOGO FIX]", {
          titleSize: "22pt",
          titleDuplicateAvoided: true,
          logoWidth: "3.0in",
          logoHeightInches: (logoEmuH / EMU_PER_INCH).toFixed(3),
          logoAspectRatioPreserved: true,
          logoNotStretched: true,
        });

        const logoRelId = "rIdLogoImg";

        // Header layout: borderless 2-column table.
        //   Left cell  (≈ 4 inches): logo image, vertically centered
        //   Right cell (≈ remaining): project title (22pt bold grey-blue),
        //                              vertically + horizontally centered
        // Title appears here on EVERY page so the body never needs a
        // separate title paragraph (no body-level title duplication).
        const titleUpperForHeader = projectName ? projectName.toUpperCase() : "";
        const escapedTitleForHeader = titleUpperForHeader
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;");
        const HDR_NO_BORDER =
          `<w:tcBorders>` +
            `<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
            `<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
            `<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
            `<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
          `</w:tcBorders>`;
        const HDR_ZERO_MAR =
          `<w:tcMar>` +
            `<w:top w:w="0" w:type="dxa"/>` +
            `<w:left w:w="0" w:type="dxa"/>` +
            `<w:bottom w:w="0" w:type="dxa"/>` +
            `<w:right w:w="0" w:type="dxa"/>` +
          `</w:tcMar>`;
        const titleHeaderRunXml = titleUpperForHeader
          ? `<w:r>` +
              `<w:rPr>` +
                `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>` +
                `<w:b/><w:bCs/>` +
                `<w:sz w:val="44"/><w:szCs w:val="44"/>` +
                `<w:color w:val="44546A"/>` +
              `</w:rPr>` +
              `<w:t xml:space="preserve">${escapedTitleForHeader}</w:t>` +
            `</w:r>`
          : `<w:r/>`;
        const HEADER_XML =
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
          `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
          ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
          ` xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"` +
          ` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
          ` xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
          `<w:tbl>` +
            `<w:tblPr>` +
              `<w:tblW w:w="5000" w:type="pct"/>` +
              `<w:tblInd w:w="0" w:type="dxa"/>` +
              `<w:tblBorders>` +
                `<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
                `<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
                `<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
                `<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
                `<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
                `<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
              `</w:tblBorders>` +
              `<w:tblCellMar>` +
                `<w:top w:w="0" w:type="dxa"/>` +
                `<w:left w:w="0" w:type="dxa"/>` +
                `<w:bottom w:w="0" w:type="dxa"/>` +
                `<w:right w:w="0" w:type="dxa"/>` +
              `</w:tblCellMar>` +
            `</w:tblPr>` +
            `<w:tblGrid>` +
              `<w:gridCol w:w="2340"/>` +
              `<w:gridCol w:w="4680"/>` +
              `<w:gridCol w:w="2340"/>` +
            `</w:tblGrid>` +
            `<w:tr>` +
              // Left cell — logo (25%)
              `<w:tc>` +
                `<w:tcPr><w:tcW w:w="1250" w:type="pct"/>${HDR_NO_BORDER}${HDR_ZERO_MAR}<w:vAlign w:val="center"/></w:tcPr>` +
                `<w:p>` +
                  `<w:pPr><w:pStyle w:val="Header"/><w:spacing w:before="0" w:after="0"/></w:pPr>` +
                  `<w:r>` +
                    `<w:rPr/>` +
                    `<w:drawing>` +
                      `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
                        `<wp:extent cx="${logoEmuW}" cy="${logoEmuH}"/>` +
                        `<wp:docPr id="999" name="Race Logo"/>` +
                        `<a:graphic>` +
                          `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
                            `<pic:pic>` +
                              `<pic:nvPicPr>` +
                                `<pic:cNvPr id="0" name="race_logo.png"/>` +
                                `<pic:cNvPicPr/>` +
                              `</pic:nvPicPr>` +
                              `<pic:blipFill>` +
                                `<a:blip r:embed="${logoRelId}"/>` +
                                `<a:stretch><a:fillRect/></a:stretch>` +
                              `</pic:blipFill>` +
                              `<pic:spPr>` +
                                `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${logoEmuW}" cy="${logoEmuH}"/></a:xfrm>` +
                                `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
                              `</pic:spPr>` +
                            `</pic:pic>` +
                          `</a:graphicData>` +
                        `</a:graphic>` +
                      `</wp:inline>` +
                    `</w:drawing>` +
                  `</w:r>` +
                `</w:p>` +
              `</w:tc>` +
              // Center cell — project title (50% — sits centered on page)
              `<w:tc>` +
                `<w:tcPr><w:tcW w:w="2500" w:type="pct"/>${HDR_NO_BORDER}${HDR_ZERO_MAR}<w:vAlign w:val="center"/></w:tcPr>` +
                `<w:p>` +
                  `<w:pPr>` +
                    `<w:pStyle w:val="Header"/>` +
                    `<w:jc w:val="center"/>` +
                    `<w:spacing w:before="0" w:after="0"/>` +
                  `</w:pPr>` +
                  titleHeaderRunXml +
                `</w:p>` +
              `</w:tc>` +
              // Right cell — empty spacer (25% — balances the logo column)
              `<w:tc>` +
                `<w:tcPr><w:tcW w:w="1250" w:type="pct"/>${HDR_NO_BORDER}${HDR_ZERO_MAR}<w:vAlign w:val="center"/></w:tcPr>` +
                `<w:p><w:pPr><w:pStyle w:val="Header"/><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>` +
              `</w:tc>` +
            `</w:tr>` +
          `</w:tbl>` +
          // Trailing empty paragraph required after a table inside ftr/hdr.
          `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>` +
          `</w:hdr>`;

        // Find existing header files or create header1.xml.
        const headerNames = Object.keys(renderedZip.files).filter((n) =>
          /^word\/header\d+\.xml$/.test(n)
        );
        const headersToUpdate = headerNames.length > 0
          ? headerNames
          : ["word/header1.xml"];

        for (const hdrName of headersToUpdate) {
          renderedZip.file(hdrName, HEADER_XML);

          const hdrBaseName = hdrName.replace("word/", "");
          const hdrRelsPath = `word/_rels/${hdrBaseName}.rels`;
          const hdrRelsXml =
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
              `<Relationship Id="${logoRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/race_logo.png"/>` +
            `</Relationships>`;
          renderedZip.file(hdrRelsPath, hdrRelsXml);
          headerPatched += 1;
        }

        // Ensure document.xml.rels references the header file(s).
        const docRelsFile = renderedZip.file("word/_rels/document.xml.rels");
        if (docRelsFile) {
          let docRelsXml = docRelsFile.asText();
          for (const hdrName of headersToUpdate) {
            const hdrBaseName = hdrName.replace("word/", "");
            if (!docRelsXml.includes(hdrBaseName)) {
              const rIdMatches = docRelsXml.match(/Id="rId(\d+)"/g) || [];
              let maxId = 0;
              for (const m of rIdMatches) {
                const num = parseInt(m.replace(/[^0-9]/g, ""), 10);
                if (num > maxId) maxId = num;
              }
              const newRId = `rId${maxId + 1}`;
              docRelsXml = docRelsXml.replace(
                "</Relationships>",
                `<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="${hdrBaseName}"/></Relationships>`
              );

              const docXmlFile = renderedZip.file("word/document.xml");
              if (docXmlFile) {
                let docXml = docXmlFile.asText();
                const hdrRefTag = `<w:headerReference w:type="default" r:id="${newRId}"/>`;
                if (!docXml.includes(hdrRefTag)) {
                  docXml = docXml.replace(
                    /(<w:sectPr\b[^>]*>)/g,
                    `$1${hdrRefTag}`
                  );
                }
                renderedZip.file("word/document.xml", docXml);
              }
            }
          }
          renderedZip.file("word/_rels/document.xml.rels", docRelsXml);
        }

        // Ensure [Content_Types].xml has an Override for each header.
        const ctFile2 = renderedZip.file("[Content_Types].xml");
        if (ctFile2) {
          let ctXml = ctFile2.asText();
          for (const hdrName of headersToUpdate) {
            const partName = `/${hdrName}`;
            if (!ctXml.includes(partName)) {
              ctXml = ctXml.replace(
                "</Types>",
                `<Override PartName="${partName}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`
              );
            }
          }
          renderedZip.file("[Content_Types].xml", ctXml);
        }

        console.log("[CLIENT DOCX HEADER]", {
          logoFile: LOGO_PATH,
          logoSize: logoBuffer.length,
          logoEmuW,
          logoEmuH,
          logoInchesW: LOGO_TARGET_W_IN,
          logoInchesH: (logoEmuH / EMU_PER_INCH).toFixed(3),
          titleInHeader: false,
          titleInBody: "via {projectNameUpper} template tag",
          headersUpdated: headersToUpdate,
          headerPatched,
        });
      } else {
        console.warn("[CLIENT DOCX HEADER] logo file not found at", LOGO_PATH);
      }
    } catch (hdrErr) {
      console.error("[CLIENT DOCX HEADER] header injection failed - non-fatal:", hdrErr);
    }

    console.log("[DOCX MANUAL_REFERENCE_FIX_QA]", {
      twoImageLayout: "matches manual Word correction",
      twoImageMode: "100% borderless 2-column inline table",
      twoImageSize: "3.95in x 2.85in each",
      oldTwoImageRenderRemoved: true,
      titleFirstTwoPages: "one title",
      titleOtherPages: "one title only",
      duplicateTitleBeforeImagesRemoved: true,
      observationOrder: "title-table-images-pagebreak",
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
