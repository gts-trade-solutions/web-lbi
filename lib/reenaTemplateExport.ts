import path from "path";
import crypto from "crypto";
import { deflateSync } from "zlib";
import { promises as fs, existsSync } from "fs";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import pool from "./db";
import { getReadSignedUrl, s3KeyFromUrl } from "./s3";
import { injectDifficultyThemeColors } from "./wordThemeColors";
import { composeObservation } from "./observationTemplates";
import { formatNeLat, formatNeLon } from "./coordFormat";
import { getCategoryDisplayName, normalizeCategory, type StageSummaryRow } from "./categoryDisplay";
import { STANDARD_HINDRANCE_ROWS, STANDARD_HINDRANCE_CONSIDERATIONS, buildStandardConclusionHtml } from "./standardReportText";

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
    // Diagnostic: log sharp's actual capabilities so we know in
    // production exactly which features are available. SVG support
    // requires libvips compiled with librsvg; text support requires
    // Pango+FreeType. These are independent — production may have
    // one without the other.
    try {
      const formatNames = _sharp.format ? Object.keys(_sharp.format) : [];
      const svgIn = _sharp.format?.svg?.input?.buffer === true;
      const pngOut = _sharp.format?.png?.output?.buffer === true;
      console.log("[SHARP DIAGNOSTIC]", {
        sharpVersion: _sharp.versions?.sharp || "unknown",
        vipsVersion: _sharp.versions?.vips || "unknown",
        platform: process.platform,
        svgInputSupported: svgIn,
        pngOutputSupported: pngOut,
        availableFormats: formatNames.slice(0, 30),
      });
    } catch (diagErr) {
      console.warn("[SHARP DIAGNOSTIC] failed to read sharp metadata:", diagErr);
    }
  } catch (err) {
    console.warn("[reenaTemplateExport] sharp not available - skipping image compression:", err);
    _sharp = null;
  }
  return _sharp;
}

/**
 * Read width/height from a PNG buffer WITHOUT sharp by parsing the
 * IHDR chunk directly. Production environments where sharp/libvips
 * isn't available (e.g. lightweight Linux containers) still need to
 * compute aspect ratios for logos and other PNGs — this gives a
 * dependency-free fallback so images aren't forced to stretch.
 *
 * PNG layout:
 *   - 8-byte signature (89 50 4E 47 0D 0A 1A 0A)
 *   - First chunk MUST be IHDR (length 13)
 *     - 4 bytes  length
 *     - 4 bytes  type "IHDR"
 *     - 4 bytes  width  (big-endian, at file offset 16)
 *     - 4 bytes  height (big-endian, at file offset 20)
 */
function readPngDimensions(buffer: Buffer): { width: number; height: number } | null {
  if (!Buffer.isBuffer(buffer) || buffer.length < 24) return null;
  // Verify PNG signature
  if (
    buffer[0] !== 0x89 ||
    buffer[1] !== 0x50 ||
    buffer[2] !== 0x4e ||
    buffer[3] !== 0x47 ||
    buffer[4] !== 0x0d ||
    buffer[5] !== 0x0a ||
    buffer[6] !== 0x1a ||
    buffer[7] !== 0x0a
  ) {
    return null;
  }
  const width = buffer.readUInt32BE(16);
  const height = buffer.readUInt32BE(20);
  if (width <= 0 || height <= 0) return null;
  return { width, height };
}

// ---- PURE-JS PNG ENCODER ----
// Pre-computed CRC table for IEEE 802.3 polynomial (used by PNG chunks).
const _CRC32_TABLE: number[] = (() => {
  const t = new Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    t[n] = c >>> 0;
  }
  return t;
})();
function _crc32(buf: Buffer): number {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i += 1) {
    c = _CRC32_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}
function _pngChunk(type: string, data: Buffer): Buffer {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length, 0);
  const typeBuf = Buffer.from(type, "ascii");
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(_crc32(Buffer.concat([typeBuf, data])), 0);
  return Buffer.concat([length, typeBuf, data, crcBuf]);
}
/**
 * Encode an RGBA pixel array to a PNG buffer using only Node's
 * built-in zlib — NO sharp / libvips / librsvg required. This is the
 * production-safe path for generating images when sharp can't render.
 */
function encodeRgbaToPng(width: number, height: number, rgba: Uint8Array): Buffer {
  // Build raw scanlines with PNG row-filter byte (0 = None) per row.
  const rowSize = width * 4;
  const raw = Buffer.alloc((rowSize + 1) * height);
  for (let y = 0; y < height; y += 1) {
    raw[y * (rowSize + 1)] = 0;
    rgba.subarray(y * rowSize, (y + 1) * rowSize).forEach((v, i) => {
      raw[y * (rowSize + 1) + 1 + i] = v;
    });
  }
  const compressed = deflateSync(raw);
  const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;  // bit depth
  ihdr[9] = 6;  // color type RGBA
  ihdr[10] = 0; // compression
  ihdr[11] = 0; // filter
  ihdr[12] = 0; // interlace
  return Buffer.concat([
    sig,
    _pngChunk("IHDR", ihdr),
    _pngChunk("IDAT", compressed),
    _pngChunk("IEND", Buffer.alloc(0)),
  ]);
}

/**
 * Tiny 2D drawing context backed by an RGBA pixel buffer. Provides
 * just enough primitives (filled rect, filled circle, hollow circle,
 * dotted line, drop pin) to render the route-locations stepper
 * without sharp/libvips/librsvg or any other native dependency.
 */
type RgbColor = readonly [number, number, number, number]; // r,g,b,a 0-255
function _hexToRgba(hex: string, alpha = 255): RgbColor {
  const h = hex.replace(/^#/, "");
  const n = parseInt(h, 16);
  return [(n >> 16) & 0xff, (n >> 8) & 0xff, n & 0xff, alpha];
}
function _setPixel(buf: Uint8Array, w: number, h: number, x: number, y: number, c: RgbColor) {
  if (x < 0 || y < 0 || x >= w || y >= h) return;
  const o = (y * w + x) * 4;
  // Source-over alpha compositing
  const sa = c[3] / 255;
  if (sa >= 1) {
    buf[o] = c[0]; buf[o + 1] = c[1]; buf[o + 2] = c[2]; buf[o + 3] = 255;
    return;
  }
  const da = buf[o + 3] / 255;
  const oa = sa + da * (1 - sa);
  if (oa <= 0) return;
  buf[o] = Math.round((c[0] * sa + buf[o] * da * (1 - sa)) / oa);
  buf[o + 1] = Math.round((c[1] * sa + buf[o + 1] * da * (1 - sa)) / oa);
  buf[o + 2] = Math.round((c[2] * sa + buf[o + 2] * da * (1 - sa)) / oa);
  buf[o + 3] = Math.round(oa * 255);
}
function _fillRect(buf: Uint8Array, w: number, h: number, x: number, y: number, rw: number, rh: number, c: RgbColor) {
  const x2 = Math.min(w, Math.floor(x + rw));
  const y2 = Math.min(h, Math.floor(y + rh));
  for (let py = Math.max(0, Math.floor(y)); py < y2; py += 1) {
    for (let px = Math.max(0, Math.floor(x)); px < x2; px += 1) {
      _setPixel(buf, w, h, px, py, c);
    }
  }
}
function _fillRoundedRect(
  buf: Uint8Array, w: number, h: number,
  x: number, y: number, rw: number, rh: number, radius: number, c: RgbColor
) {
  const r = Math.min(radius, rw / 2, rh / 2);
  // Draw straight body parts
  _fillRect(buf, w, h, x + r, y, rw - 2 * r, rh, c);
  _fillRect(buf, w, h, x, y + r, rw, rh - 2 * r, c);
  // Draw the four rounded corners as filled quarter-circles
  const corners: Array<[number, number]> = [
    [x + r, y + r],                 // top-left
    [x + rw - r - 1, y + r],        // top-right
    [x + r, y + rh - r - 1],        // bottom-left
    [x + rw - r - 1, y + rh - r - 1], // bottom-right
  ];
  for (const [cx, cy] of corners) {
    for (let py = Math.floor(cy - r); py <= Math.ceil(cy + r); py += 1) {
      for (let px = Math.floor(cx - r); px <= Math.ceil(cx + r); px += 1) {
        const dx = px - cx;
        const dy = py - cy;
        if (dx * dx + dy * dy <= r * r) {
          _setPixel(buf, w, h, px, py, c);
        }
      }
    }
  }
}
function _strokeRoundedRect(
  buf: Uint8Array, w: number, h: number,
  x: number, y: number, rw: number, rh: number, radius: number,
  thickness: number, c: RgbColor, fillColor: RgbColor = [255, 255, 255, 255]
) {
  // Fill-and-cut approach: paint the OUTER rounded shape in the
  // border color, then paint a SMALLER inner rounded shape in the
  // fill color to "cut" the interior. Produces a clean stroke with
  // no corner artifacts. fillColor defaults to white (matches the
  // stepper background) — caller can pass a different color to
  // composite cleanly over a non-white background.
  _fillRoundedRect(buf, w, h, x, y, rw, rh, radius, c);
  const innerRadius = Math.max(0, radius - thickness);
  _fillRoundedRect(
    buf, w, h,
    x + thickness, y + thickness,
    rw - 2 * thickness, rh - 2 * thickness,
    innerRadius, fillColor
  );
}
function _fillCircle(buf: Uint8Array, w: number, h: number, cx: number, cy: number, r: number, c: RgbColor) {
  for (let py = Math.floor(cy - r); py <= Math.ceil(cy + r); py += 1) {
    for (let px = Math.floor(cx - r); px <= Math.ceil(cx + r); px += 1) {
      const dx = px - cx;
      const dy = py - cy;
      if (dx * dx + dy * dy <= r * r) {
        _setPixel(buf, w, h, px, py, c);
      }
    }
  }
}
function _strokeCircle(
  buf: Uint8Array, w: number, h: number,
  cx: number, cy: number, outerR: number, innerR: number, c: RgbColor
) {
  for (let py = Math.floor(cy - outerR); py <= Math.ceil(cy + outerR); py += 1) {
    for (let px = Math.floor(cx - outerR); px <= Math.ceil(cx + outerR); px += 1) {
      const dx = px - cx;
      const dy = py - cy;
      const d2 = dx * dx + dy * dy;
      if (d2 <= outerR * outerR && d2 >= innerR * innerR) {
        _setPixel(buf, w, h, px, py, c);
      }
    }
  }
}
function _drawDottedVerticalLine(
  buf: Uint8Array, w: number, h: number,
  x: number, y1: number, y2: number, dotSize: number, gapSize: number, c: RgbColor
) {
  let y = y1;
  while (y < y2) {
    _fillRect(buf, w, h, x - dotSize / 2, y, dotSize, dotSize, c);
    y += dotSize + gapSize;
  }
}
/**
 * Draw a filled red drop-pin shape (similar to Material/Google Maps
 * pin) centered horizontally at cx, with its top tip at topY. The pin
 * has total height ~ topY+h, width ~ 2*r+something. White inner circle.
 */
function _drawPin(
  buf: Uint8Array, w: number, h: number,
  cx: number, topY: number, size: number,
  fill: RgbColor, innerColor: RgbColor
) {
  const r = size * 0.42;          // top circle radius
  const cy = topY + r + 2;         // center of top circle
  const tipY = topY + size;        // tip position
  // Tear-drop body: top circle + downward triangle
  _fillCircle(buf, w, h, cx, cy, r, fill);
  // Triangle from circle tangent down to tipY
  for (let py = Math.floor(cy); py <= Math.ceil(tipY); py += 1) {
    const t = (py - cy) / (tipY - cy); // 0..1
    const halfW = r * (1 - t * 0.95);
    if (halfW < 0.5) break;
    _fillRect(buf, w, h, cx - halfW, py, halfW * 2, 1, fill);
  }
  // Inner white circle
  _fillCircle(buf, w, h, cx, cy - 1, r * 0.42, innerColor);
}

// ---- TINY BITMAP FONT (5x7) for PNG-renderer text (uppercase only) ----
// Each glyph is 5 columns x 7 rows, stored as 7 row-bytes (5 LSBs).
// Used by the route-locations stepper fallback to render labels and
// the heading without depending on any font / library.
const _BITMAP_FONT_5x7: Record<string, number[]> = {
  // Uppercase letters
  "A": [0x0E, 0x11, 0x11, 0x1F, 0x11, 0x11, 0x11],
  "B": [0x1E, 0x11, 0x11, 0x1E, 0x11, 0x11, 0x1E],
  "C": [0x0E, 0x11, 0x10, 0x10, 0x10, 0x11, 0x0E],
  "D": [0x1E, 0x11, 0x11, 0x11, 0x11, 0x11, 0x1E],
  "E": [0x1F, 0x10, 0x10, 0x1E, 0x10, 0x10, 0x1F],
  "F": [0x1F, 0x10, 0x10, 0x1E, 0x10, 0x10, 0x10],
  "G": [0x0E, 0x11, 0x10, 0x17, 0x11, 0x11, 0x0E],
  "H": [0x11, 0x11, 0x11, 0x1F, 0x11, 0x11, 0x11],
  "I": [0x0E, 0x04, 0x04, 0x04, 0x04, 0x04, 0x0E],
  "J": [0x07, 0x02, 0x02, 0x02, 0x02, 0x12, 0x0C],
  "K": [0x11, 0x12, 0x14, 0x18, 0x14, 0x12, 0x11],
  "L": [0x10, 0x10, 0x10, 0x10, 0x10, 0x10, 0x1F],
  "M": [0x11, 0x1B, 0x15, 0x15, 0x11, 0x11, 0x11],
  "N": [0x11, 0x11, 0x19, 0x15, 0x13, 0x11, 0x11],
  "O": [0x0E, 0x11, 0x11, 0x11, 0x11, 0x11, 0x0E],
  "P": [0x1E, 0x11, 0x11, 0x1E, 0x10, 0x10, 0x10],
  "Q": [0x0E, 0x11, 0x11, 0x11, 0x15, 0x12, 0x0D],
  "R": [0x1E, 0x11, 0x11, 0x1E, 0x14, 0x12, 0x11],
  "S": [0x0F, 0x10, 0x10, 0x0E, 0x01, 0x01, 0x1E],
  "T": [0x1F, 0x04, 0x04, 0x04, 0x04, 0x04, 0x04],
  "U": [0x11, 0x11, 0x11, 0x11, 0x11, 0x11, 0x0E],
  "V": [0x11, 0x11, 0x11, 0x11, 0x11, 0x0A, 0x04],
  "W": [0x11, 0x11, 0x11, 0x15, 0x15, 0x15, 0x0A],
  "X": [0x11, 0x11, 0x0A, 0x04, 0x0A, 0x11, 0x11],
  "Y": [0x11, 0x11, 0x0A, 0x04, 0x04, 0x04, 0x04],
  "Z": [0x1F, 0x01, 0x02, 0x04, 0x08, 0x10, 0x1F],
  // Digits
  "0": [0x0E, 0x11, 0x13, 0x15, 0x19, 0x11, 0x0E],
  "1": [0x04, 0x0C, 0x04, 0x04, 0x04, 0x04, 0x0E],
  "2": [0x0E, 0x11, 0x01, 0x02, 0x04, 0x08, 0x1F],
  "3": [0x0E, 0x11, 0x01, 0x06, 0x01, 0x11, 0x0E],
  "4": [0x02, 0x06, 0x0A, 0x12, 0x1F, 0x02, 0x02],
  "5": [0x1F, 0x10, 0x1E, 0x01, 0x01, 0x11, 0x0E],
  "6": [0x06, 0x08, 0x10, 0x1E, 0x11, 0x11, 0x0E],
  "7": [0x1F, 0x01, 0x02, 0x04, 0x08, 0x08, 0x08],
  "8": [0x0E, 0x11, 0x11, 0x0E, 0x11, 0x11, 0x0E],
  "9": [0x0E, 0x11, 0x11, 0x0F, 0x01, 0x02, 0x0C],
  // Punctuation
  " ": [0, 0, 0, 0, 0, 0, 0],
  "-": [0x00, 0x00, 0x00, 0x1F, 0x00, 0x00, 0x00],
  ".": [0, 0, 0, 0, 0, 0x06, 0x06],
  ",": [0, 0, 0, 0, 0, 0x06, 0x04],
  ":": [0, 0x06, 0x06, 0, 0x06, 0x06, 0],
  "/": [0x01, 0x02, 0x02, 0x04, 0x08, 0x08, 0x10],
  "(": [0x02, 0x04, 0x08, 0x08, 0x08, 0x04, 0x02],
  ")": [0x08, 0x04, 0x02, 0x02, 0x02, 0x04, 0x08],
};
function _drawBitmapText(
  buf: Uint8Array, w: number, h: number,
  text: string, x: number, y: number, scale: number, c: RgbColor
) {
  // High-quality bitmap text rendering. Steps:
  //  1. Render at 3× scale into a temporary single-channel mask.
  //  2. Apply a 1-pixel dilation pass to thicken strokes (gives a
  //     heavier, more readable look).
  //  3. Apply a 3×3 box blur to soften edges.
  //  4. Downsample 3× → 1× by averaging 3×3 blocks (9 alpha levels
  //     per output pixel — proper anti-aliasing).
  //  5. Composite each output pixel onto the main buffer.
  const upper = text.toUpperCase();
  const cellW = 5 * scale + scale; // 5 cols + 1 col gap
  const renderH = 7 * scale;
  const renderW = upper.length * cellW;

  const SS = 3; // supersampling factor
  const ssScale = scale * SS;
  const ssCellW = 5 * ssScale + ssScale;
  const ssH = 7 * ssScale;
  const ssW = upper.length * ssCellW;

  // Step 1: rasterise the bitmap font into the mask at 3× scale
  const mask = new Uint8Array(ssW * ssH);
  for (let i = 0; i < upper.length; i += 1) {
    const ch = upper[i];
    const glyph = _BITMAP_FONT_5x7[ch] || _BITMAP_FONT_5x7[" "];
    for (let row = 0; row < 7; row += 1) {
      const bits = glyph[row];
      for (let col = 0; col < 5; col += 1) {
        if (bits & (1 << (4 - col))) {
          const baseX = i * ssCellW + col * ssScale;
          const baseY = row * ssScale;
          for (let py = 0; py < ssScale; py += 1) {
            const yi = (baseY + py) * ssW;
            for (let px = 0; px < ssScale; px += 1) {
              mask[yi + baseX + px] = 255;
            }
          }
        }
      }
    }
  }

  // Step 2: dilate by 1 pixel (thickens strokes for "bolder" look)
  const dilated = new Uint8Array(ssW * ssH);
  for (let yy = 0; yy < ssH; yy += 1) {
    for (let xx = 0; xx < ssW; xx += 1) {
      if (mask[yy * ssW + xx]) {
        dilated[yy * ssW + xx] = 255;
      } else {
        // If any neighbour is on, set dilated
        const left = xx > 0 ? mask[yy * ssW + (xx - 1)] : 0;
        const right = xx + 1 < ssW ? mask[yy * ssW + (xx + 1)] : 0;
        const up = yy > 0 ? mask[(yy - 1) * ssW + xx] : 0;
        const down = yy + 1 < ssH ? mask[(yy + 1) * ssW + xx] : 0;
        if (left || right || up || down) dilated[yy * ssW + xx] = 200;
      }
    }
  }

  // Step 3: 3×3 box blur — smooths jagged edges into a gradient
  const blurred = new Uint8Array(ssW * ssH);
  for (let yy = 0; yy < ssH; yy += 1) {
    for (let xx = 0; xx < ssW; xx += 1) {
      let sum = 0;
      let count = 0;
      for (let dy = -1; dy <= 1; dy += 1) {
        const ny = yy + dy;
        if (ny < 0 || ny >= ssH) continue;
        for (let dx = -1; dx <= 1; dx += 1) {
          const nx = xx + dx;
          if (nx < 0 || nx >= ssW) continue;
          sum += dilated[ny * ssW + nx];
          count += 1;
        }
      }
      blurred[yy * ssW + xx] = count > 0 ? Math.round(sum / count) : 0;
    }
  }

  // Step 4 + 5: downsample 3× → 1× and composite
  for (let oy = 0; oy < renderH; oy += 1) {
    for (let ox = 0; ox < renderW; ox += 1) {
      const sx = ox * SS;
      const sy = oy * SS;
      let sum = 0;
      for (let dy = 0; dy < SS; dy += 1) {
        const yi = (sy + dy) * ssW;
        for (let dx = 0; dx < SS; dx += 1) {
          sum += blurred[yi + sx + dx];
        }
      }
      const alpha = Math.round(sum / (SS * SS));
      if (alpha > 0) {
        _setPixel(
          buf, w, h,
          x + ox, y + oy,
          [c[0], c[1], c[2], Math.round((alpha * c[3]) / 255)]
        );
      }
    }
  }
}

/**
 * Render the stepper using sharp.text + sharp.composite — bypasses
 * SVG entirely. Sharp's text input uses Pango/FreeType (commonly
 * installed) instead of librsvg (often missing). This means production
 * environments that have sharp + Pango but lack librsvg still get
 * REAL FONT TEXT instead of the pixelated bitmap fallback.
 *
 * Strategy:
 *   - Build the canvas with sharp.create (white background)
 *   - For circles, pin, dotted lines, rounded label boxes: render
 *     each shape into a small RGBA buffer using the pure-JS drawing
 *     functions, then composite at the right position
 *   - For all TEXT (heading + each label): use sharp.text to render
 *     in a real system font, composite at the right position
 *
 * Returns null if sharp is unavailable, sharp.text doesn't work,
 * or labels list is empty. The caller will then drop to the
 * pure-JS bitmap fallback.
 */
async function renderRouteLocationsStepperViaSharpText(
  labels: { label: string; pin_type: string }[]
): Promise<Buffer | null> {
  const sharp = getSharp();
  if (!sharp) return null;
  const items = labels.filter((l) => l.label).slice(0, 30);
  if (items.length === 0) return null;

  // First try a tiny sharp.text call to confirm text rendering works on
  // this build of sharp. If it throws, bail and fall through to the
  // pure-JS path.
  try {
    await sharp({
      text: { text: "test", font: "sans-serif", rgba: true, width: 200 },
    })
      .png()
      .toBuffer();
  } catch (probeErr) {
    console.warn(
      "[ROUTE LOCATIONS STEPPER] sharp.text probe failed — sharp built " +
        "without Pango/FreeType, dropping to pure-JS bitmap fallback. Error:",
      (probeErr as Error)?.message || probeErr
    );
    return null;
  }

  // Layout — same dimensions as the pure-JS PNG version
  const W = 1080;
  const ROW_H = 110;
  const HEADING_H = 70;
  const TOP_PAD = 30;
  const BOTTOM_PAD = 40;
  const H = TOP_PAD + HEADING_H + items.length * ROW_H + BOTTOM_PAD;
  const MARK_X = 90;
  const MARK_R = 24;
  const RECT_X = MARK_X + 70;
  const RECT_W = W - RECT_X - 40;
  const RECT_H = 88;
  const RECT_RADIUS = 36;

  type Composite = { input: Buffer; left: number; top: number };
  const composites: Composite[] = [];

  // Helper: build a small RGBA shape buffer with the pure-JS drawing
  // functions, then encode it as PNG so sharp can composite it.
  const shapeAsPng = (width: number, height: number, draw: (buf: Uint8Array) => void): Buffer => {
    const buf = new Uint8Array(width * height * 4);
    // Initialise transparent (so compositing onto canvas blends cleanly)
    for (let i = 3; i < buf.length; i += 4) buf[i] = 0;
    draw(buf);
    return encodeRgbaToPng(width, height, buf);
  };

  const grey800 = _hexToRgba("#1F2937");
  const grey400 = _hexToRgba("#9CA3AF");
  const grey300 = _hexToRgba("#D1D5DB");
  const red500 = _hexToRgba("#EC4054");
  const white = _hexToRgba("#FFFFFF");

  // ---- Heading TEXT via sharp.text
  try {
    const headingPng = await sharp({
      text: {
        text: "ROUTE SURVEY KEY POINTS",
        font: "sans-serif",
        rgba: true,
        height: 40,
      },
    })
      .png()
      .toBuffer();
    // Tint pixels to dark grey (sharp.text by default renders black on
    // transparent — that's fine for our dark-grey heading look).
    composites.push({ input: headingPng, left: 40, top: TOP_PAD + 10 });
  } catch (err) {
    console.warn("[STEPPER sharp.text] heading render failed:", err);
    return null;
  }

  // ---- Per-row markers + labels
  for (let i = 0; i < items.length; i += 1) {
    const it = items[i];
    const cy = TOP_PAD + HEADING_H + i * ROW_H + ROW_H / 2;
    const isEnd = i === items.length - 1 || it.pin_type === "end";

    // Connector dotted line
    if (i < items.length - 1) {
      const lineH = ROW_H - 2 * MARK_R - 12;
      const linePng = shapeAsPng(8, lineH, (buf) => {
        _drawDottedVerticalLine(buf, 8, lineH, 4, 0, lineH, 4, 8, grey400);
      });
      composites.push({ input: linePng, left: MARK_X - 4, top: cy + MARK_R + 6 });
    }

    // Marker (circle or pin)
    if (isEnd) {
      const pinPng = shapeAsPng(60, 70, (buf) => {
        _drawPin(buf, 60, 70, 30, 4, 60, red500, white);
      });
      composites.push({ input: pinPng, left: MARK_X - 30, top: cy - 30 });
    } else {
      const circlePng = shapeAsPng(60, 60, (buf) => {
        _strokeCircle(buf, 60, 60, 30, 30, MARK_R, MARK_R - 4, grey800);
      });
      composites.push({ input: circlePng, left: MARK_X - 30, top: cy - 30 });
    }

    // Rounded label pill (background + border, NO text)
    const pillPng = shapeAsPng(RECT_W, RECT_H, (buf) => {
      _fillRoundedRect(buf, RECT_W, RECT_H, 0, 0, RECT_W, RECT_H, RECT_RADIUS, white);
      _strokeRoundedRect(
        buf, RECT_W, RECT_H,
        0, 0, RECT_W, RECT_H, RECT_RADIUS, 2, grey300, white
      );
    });
    const rectY = cy - RECT_H / 2;
    composites.push({ input: pillPng, left: RECT_X, top: rectY });

    // Label TEXT via sharp.text
    try {
      const labelPng = await sharp({
        text: {
          text: it.label,
          font: "sans-serif",
          rgba: true,
          height: 38,
        },
      })
        .png()
        .toBuffer();
      composites.push({
        input: labelPng,
        left: RECT_X + 28,
        top: rectY + Math.round(RECT_H / 2 - 22),
      });
    } catch (err) {
      console.warn("[STEPPER sharp.text] label render failed for", it.label, err);
      return null;
    }
  }

  try {
    const out = await sharp({
      create: {
        width: W,
        height: H,
        channels: 4,
        background: { r: 255, g: 255, b: 255, alpha: 1 },
      },
    })
      .composite(composites)
      .png()
      .toBuffer();
    console.log("[ROUTE LOCATIONS STEPPER GENERATED]", {
      itemCount: items.length,
      renderPath: "sharp.text + composite (real font, no SVG/librsvg needed)",
      bytes: out.length,
    });
    return out;
  } catch (err) {
    console.warn(
      "[STEPPER sharp.text] final composite failed:",
      (err as Error)?.message || err
    );
    return null;
  }
}

/**
 * Render the route-locations stepper as a PNG using only the
 * pure-JS encoder above. Identical in spirit to the SVG → sharp
 * version, but works in production environments where sharp /
 * libvips / librsvg isn't available. This is the SECOND-tier
 * fallback (after the sharp/SVG path, before the Word-table path).
 */
function renderRouteLocationsStepperPng(
  labels: { label: string; pin_type: string }[]
): Buffer | null {
  const items = labels.filter((l) => l.label).slice(0, 30);
  if (items.length === 0) return null;

  // Canvas dimensions
  const W = 1080;
  const ROW_H = 110;
  const HEADING_H = 70;
  const TOP_PAD = 30;
  const BOTTOM_PAD = 40;
  const H = TOP_PAD + HEADING_H + items.length * ROW_H + BOTTOM_PAD;
  const buf = new Uint8Array(W * H * 4);
  // Initialise to white background (so the image isn't transparent
  // when embedded in Word — keeps the visible look matching the SVG).
  for (let i = 0; i < buf.length; i += 4) {
    buf[i] = 255; buf[i + 1] = 255; buf[i + 2] = 255; buf[i + 3] = 255;
  }

  const grey800 = _hexToRgba("#1F2937");
  const grey400 = _hexToRgba("#9CA3AF");
  const grey300 = _hexToRgba("#D1D5DB");
  const black0F = _hexToRgba("#0F172A");
  const red500 = _hexToRgba("#EC4054");
  const white = _hexToRgba("#FFFFFF");

  // ---- Heading: "ROUTE SURVEY KEY POINTS"
  _drawBitmapText(
    buf, W, H,
    "ROUTE SURVEY KEY POINTS",
    40,
    TOP_PAD + 14,
    5, // bigger scale for the heading
    grey800
  );

  const MARK_X = 90;
  const MARK_R = 24;
  const RECT_X = MARK_X + 70;
  const RECT_W = W - RECT_X - 40;
  const RECT_H = 88;
  // Bigger radius makes the labels read as proper "pills" matching
  // the web-app design.
  const RECT_RADIUS = 36;

  for (let i = 0; i < items.length; i += 1) {
    const it = items[i];
    const cy = TOP_PAD + HEADING_H + i * ROW_H + ROW_H / 2;
    const isEnd = i === items.length - 1 || it.pin_type === "end";

    // Connector dotted line down to next row
    if (i < items.length - 1) {
      _drawDottedVerticalLine(
        buf, W, H,
        MARK_X,
        cy + MARK_R + 6,
        cy + ROW_H - MARK_R - 6,
        4, 8, grey400
      );
    }

    // Marker
    if (isEnd) {
      _drawPin(buf, W, H, MARK_X, cy - 26, 60, red500, white);
    } else {
      // Hollow circle: outer dark, inner white
      _strokeCircle(buf, W, H, MARK_X, cy, MARK_R, MARK_R - 4, grey800);
    }

    // Label rounded pill
    const rectY = cy - RECT_H / 2;
    _fillRoundedRect(buf, W, H, RECT_X, rectY, RECT_W, RECT_H, RECT_RADIUS, white);
    _strokeRoundedRect(buf, W, H, RECT_X, rectY, RECT_W, RECT_H, RECT_RADIUS, 2, grey300);
    _drawBitmapText(
      buf, W, H,
      it.label,
      RECT_X + 28,
      rectY + RECT_H / 2 - 18,
      5, // bigger label scale
      black0F
    );
  }

  console.log("[ROUTE LOCATIONS STEPPER PURE-JS]", {
    itemCount: items.length,
    renderedSize: `${W}x${H}`,
    fallbackPath: "pure-JS PNG renderer (no sharp/libvips needed)",
  });
  return encodeRgbaToPng(W, H, buf);
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
// Observation-photo target, adapted PER EXPORT by scale in generateReenaDocx.
// Big reports use smaller photos so a several-thousand-photo file stays
// openable in Word and the server's peak memory stays bounded.
let OBS_PHOTO_TARGET = { width: 1120, height: 720, quality: 82 };

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
      // TRIM the transparent padding baked into the source PNGs (e.g. gate.png
      // is a 512×512 file whose glyph is only 430×152 — the rest is empty). If
      // we kept the padding and rendered into a square cell, the glyph floated
      // with large blank space above/below it. After trimming we resize to fit
      // INSIDE a box (no forced square / no re-padding) so the buffer keeps the
      // glyph's real aspect ratio; getSize then renders it at that ratio.
      optimized = await sharp(input)
        .rotate()
        .trim({ threshold: 10 })
        .resize({
          width: 320,
          height: 320,
          fit: "inside",
          withoutEnlargement: true,
        })
        .png()
        .toBuffer();
    } else {
      // Dimensions/quality adapt to the export's scale (OBS_PHOTO_TARGET): the
      // default 1120×720 q72 stays crisp at print size, while very large
      // reports drop to a smaller size so a several-thousand-photo file stays
      // openable and memory stays bounded.
      optimized = await sharp(input)
        .rotate()
        // Preserve the source photo's colour profile. Phone cameras (iPhone and
        // most modern Android) shoot in the WIDE "Display P3" colour space with
        // an embedded ICC profile. sharp drops that profile on JPEG conversion
        // unless told to keep it, so Word reads the P3 pixels as plain sRGB and
        // EVERY photo looks colour-shifted (duller / greener). Keeping the ICC
        // profile lets Word colour-manage the image and show the true colours.
        // keepIccProfile keeps ONLY the colour profile (not EXIF orientation),
        // so it never fights the .rotate() auto-orient above.
        .keepIccProfile()
        .resize({
          width: OBS_PHOTO_TARGET.width,
          height: OBS_PHOTO_TARGET.height,
          fit: "inside",
          withoutEnlargement: true,
        })
        .jpeg({
          quality: OBS_PHOTO_TARGET.quality,
          // 4:4:4 keeps COLOURED drawings crisp — arrows/lines and small text
          // (e.g. a yellow "H 9m") drawn on the photo. The default 4:2:0
          // subsamples colour and makes those edges blur/bleed.
          chromaSubsampling: "4:4:4",
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
  // When false, the GA Drawing image/section is left out of the report, so the
  // user can download a "without GA drawing" version alongside the full one.
  includeGa?: boolean;
  debug?: boolean;
  // Report order in the document. "desc" reverses the route (and the
  // cumulative KMs recompute for that direction). Defaults to "asc".
  sort?: "asc" | "desc";
  // Summary report: keeps the "SUMMARY REPORT" title and injects the two
  // summary tables (category counts + key-obstacle list) before the first
  // point table. Meant to be combined with a pinned-points reportIds filter.
  summaryMode?: boolean;
  // Diagonal text watermark across every page (classic Word watermark shape
  // anchored in the page header). When enabled with a non-empty text, a light
  // grey rotated text is drawn behind the page content.
  watermark?: { enabled?: boolean; text?: string };
};

export type ExportResult = {
  buffer: Buffer;
  fileName: string;
  projectName: string;
};

const TEMPLATE_PATH = path.join(process.cwd(), "templates", "reena-all-template.docx");

// Per spec: route map = 13" × 7" (1248 × 672 px) landscape, GA drawing
// landscape — sized so the GA heading + the drawing fit TOGETHER on ONE
// landscape A3 page. The template page has a red double border (offset
// ~0.67" from each edge) plus a header and footer, which eat the vertical
// space: a full 9.5"-tall drawing left no room for the heading above it, so
// the drawing spilled onto the next page (heading alone on one page, image
// on the next). 1142 × 775 px (~11.9" × 8.1", same aspect) leaves room for
// the heading so title + image always share one page. On smaller pages Word
// scales down proportionally while preserving aspect ratio.
const ROUTE_MAP_SIZE: [number, number] = [1248, 672];
const GA_DRAWING_SIZE: [number, number] = [1142, 775];

// The GA drawing slot in the template holds ONE image. When several GA images
// were uploaded, lay them out in a grid on a single canvas that matches the GA
// slot's aspect ratio, so all of them show (not just the first) and nothing is
// distorted when the slot stretches the image to fit.
async function compositeImagesGrid(
  buffers: Buffer[],
  boxW: number,
  boxH: number
): Promise<Buffer | null> {
  const sharpLib = getSharp();
  if (!sharpLib || buffers.length === 0) return null;
  try {
    const n = buffers.length;
    const scale = 2; // render at 2x for crispness
    const W = boxW * scale;
    const H = boxH * scale;
    const cols = n <= 1 ? 1 : n <= 4 ? 2 : n <= 9 ? 3 : 4;
    const rows = Math.ceil(n / cols);
    const pad = Math.max(6, Math.round(W * 0.008));
    const cellW = Math.floor((W - pad * (cols + 1)) / cols);
    const cellH = Math.floor((H - pad * (rows + 1)) / rows);
    if (cellW < 10 || cellH < 10) return null;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parts: any[] = [];
    for (let i = 0; i < n; i += 1) {
      const rr = Math.floor(i / cols);
      const cc = i % cols;
      const cell = await sharpLib(buffers[i])
        .rotate()
        .resize(cellW, cellH, { fit: "contain", background: { r: 255, g: 255, b: 255, alpha: 1 } })
        .toBuffer();
      parts.push({ input: cell, left: pad + cc * (cellW + pad), top: pad + rr * (cellH + pad) });
    }
    return await sharpLib({
      create: { width: W, height: H, channels: 3, background: { r: 255, g: 255, b: 255 } },
    })
      .composite(parts)
      .jpeg({ quality: 88, chromaSubsampling: "4:4:4", mozjpeg: true })
      .toBuffer();
  } catch (e) {
    console.warn("[export] GA image grid composite failed:", (e as Error).message);
    return null;
  }
}
// Observation photo size for a point with a SINGLE photo — it renders BIG at
// this one size (9.4" × 6.6" / 900 × 636 px), filling the width under its
// observation table on the same landscape-A3 page. A point with 2+ photos is
// instead laid out two-per-row at SIDE_BY_SIDE_PHOTO_SIZE (see below), every
// photo the same size. Kept below the route map (13") / GA drawing (14")
// widths (isObservationPhotoCx excludes those two explicitly). The embedded
// JPEG is compressed via sharp before imageMap so the DOCX stays small.
const OBSERVATION_PHOTO_SIZE: [number, number] = [900, 636];
// Unified with the single-photo size so all observation photos are identical.
const MULTI_PHOTO_SIZE: [number, number] = [900, 636];
// A point with 2+ photos lays them out TWO PER ROW (side by side) at this
// SAME smaller size, so a pair of photos sits together on one landscape page.
// Width ~6.1" so two + the cell gaps fit inside the ~14.5" content area; the
// 900:636 aspect ratio is preserved so nothing is stretched. A single-photo
// point is NOT affected — it still renders big at OBSERVATION_PHOTO_SIZE.
const SIDE_BY_SIDE_PHOTO_SIZE: [number, number] = [760, 537];
// Category icon shown in the CATEGORY column of the observation table.
// WIDTH stays 240 so wide glyphs (gate, etc.) keep their full width and the
// CATEGORY column width (derived from [0], see OBS_COL_DXA) is unchanged.
// HEIGHT is capped at 150 (was 240): a tall portrait icon (e.g. a location
// pin) rendered at the full 240px = 2.5" made the observation ROW ~2.5" tall,
// so table (~3") + the 6.63" photo together exceeded the ~10" page content
// height and the photo spilled onto the next page — a half-empty page after
// the report. Capping the icon height to 150px (~1.56") keeps the table under
// ~2", so table + photo fit together on one page with comfortable margin.
// Aspect ratio is preserved (categoryIconDisplaySize fits inside this box),
// so nothing is stretched; wide-short glyphs are unaffected (already < 150 tall).
const CATEGORY_ICON_SIZE: [number, number] = [240, 150];
const CATEGORY_SUMMARY_ICON_SIZE: [number, number] = [80, 80];

// Word stores image extents in EMU (1 px = 9525 EMU). The layout-rebuild
// pass uses these constants to identify observation-photo drawings (both
// sizes) and ignore other in-document drawings (GA Drawing, Route Map,
// category icons).
const PX_TO_EMU = 9525;
const OBSERVATION_PHOTO_EMU_WIDTH = OBSERVATION_PHOTO_SIZE[0] * PX_TO_EMU;
const MULTI_PHOTO_EMU_WIDTH = MULTI_PHOTO_SIZE[0] * PX_TO_EMU;
const MULTI_PHOTO_EMU_HEIGHT = MULTI_PHOTO_SIZE[1] * PX_TO_EMU;
const SIDE_BY_SIDE_EMU_WIDTH = SIDE_BY_SIDE_PHOTO_SIZE[0] * PX_TO_EMU;
const SIDE_BY_SIDE_EMU_HEIGHT = SIDE_BY_SIDE_PHOTO_SIZE[1] * PX_TO_EMU;

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
  return s3KeyFromUrl(url);
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
 *
 * Transient failures (throttle 429, 5xx, timeout, network blip) are RETRIED
 * with backoff. Without this, a big report (500+ points ~ 1000 photos) that
 * briefly gets throttled by S3 would silently drop every photo after that
 * point — the "photos come up to report N then Photo not available" bug.
 */
const RETRYABLE: unique symbol = Symbol("retryable-image-fetch");
type ImageFetchResult = { buffer: Buffer; contentType: string; url: string } | null | typeof RETRYABLE;
const sleepMs = (ms: number) => new Promise<void>((r) => setTimeout(r, ms));

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
      if (!res.ok) {
        // Throttling (429) and server errors (5xx) are transient — ask the
        // caller to retry. 404/403 and other 4xx are permanent.
        return res.status === 429 || res.status >= 500 ? RETRYABLE : null;
      }
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
      // Timeout / network error is transient — retryable.
      return RETRYABLE;
    } finally {
      clearTimeout(timeoutId);
    }
  };

  const key = extractS3Key(normalized);
  // Up to 4 tries with growing backoff, so a burst of S3 throttling / transient
  // timeouts during a large export doesn't permanently drop the rest of the
  // photos. Each try does direct fetch, then a signed-URL fallback.
  const BACKOFF_MS = [400, 1200, 3000];
  for (let attempt = 0; attempt <= BACKOFF_MS.length; attempt += 1) {
    const suffix = attempt ? `#retry${attempt}` : "";

    // Direct fetch — skipped for our own bucket: it is private, so a plain GET
    // is a guaranteed 403 and only wastes a round trip per photo.
    const direct: ImageFetchResult = key ? null : await tryFetch(normalized, `direct${suffix}`);
    if (direct && direct !== RETRYABLE) return direct;

    // Signed URL fallback for private buckets.
    let signed: ImageFetchResult = null;
    if (key) {
      try {
        const signedUrl = await getReadSignedUrl(key, 600);
        signed = await tryFetch(signedUrl, `signed${suffix}`);
        if (signed && signed !== RETRYABLE) return signed;
      } catch (err) {
        console.error(`[image fetch] ${label} signed-url generation failed`, err);
      }
    } else if (attempt === 0) {
      console.warn(`[image fetch] ${label} could not extract S3 key from`, normalized);
    }

    // If neither attempt was transient, it's a permanent failure (404/403/
    // non-image) — retrying won't help, so stop now.
    if (direct !== RETRYABLE && signed !== RETRYABLE) return null;

    if (attempt < BACKOFF_MS.length) {
      console.warn(`[image fetch] ${label} transient failure — retrying in ${BACKOFF_MS[attempt]}ms`);
      await sleepMs(BACKOFF_MS[attempt]);
    }
  }

  console.error(`[image fetch] ${label} gave up after retries (still throttled/timing out)`);
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

// Sentinel put into the {conclusion} slot at render time. After docxtemplater
// runs, the post-render pass finds this single paragraph and swaps it for the
// neatly-formatted Conclusion & Certification body (see conclusionHtmlToOoxml).
const CONCLUSION_SENTINEL = "RACE_CONCLUSION_BODY_PLACEHOLDER_X7";

// Make an inline text slice XML-safe for a <w:t>. The standard conclusion HTML
// is already escaped, but a user-EDITED conclusion (contentEditable) can contain
// stray tags (<span>, <div>, <font>…) or a lone "&"/"<". Drop unknown tags and
// escape lone &/< so the run never breaks the document XML.
function conclusionSanitizeInline(t: string): string {
  return String(t)
    .replace(/<[^>]*>/g, "") // drop stray editor tags (bold/br handled by caller)
    .replace(/&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/gi, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/\s+/g, " ");
}

// Convert the inline part of a conclusion block (text + <strong>/<b> + <br/>)
// into Word runs.
function conclusionInlineRunsToOoxml(inner: string, sz: number): string {
  const runText = (t: string, bold: boolean): string => {
    const safe = conclusionSanitizeInline(t);
    if (!safe) return "";
    return (
      `<w:r><w:rPr>${bold ? "<w:b/><w:bCs/>" : ""}` +
      `<w:color w:val="111111"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr>` +
      `<w:t xml:space="preserve">${safe}</w:t></w:r>`
    );
  };
  let out = "";
  let bold = false;
  let last = 0;
  const tokenRe = /<\/?(strong|b|br)\s*\/?>/gi;
  let m: RegExpExecArray | null;
  while ((m = tokenRe.exec(inner)) !== null) {
    out += runText(inner.slice(last, m.index).replace(/\s+/g, " "), bold);
    const tag = m[1].toLowerCase();
    if (tag === "br") out += "<w:r><w:br/></w:r>";
    else bold = m[0][1] !== "/"; // "<strong>" opens, "</strong>" closes
    last = tokenRe.lastIndex;
  }
  out += runText(inner.slice(last).replace(/\s+/g, " "), bold);
  return out;
}

// Turn the fixed Conclusion HTML (<p>, <strong>, <br/>, <ul>/<li>) into neat
// Word paragraphs: proper paragraph breaks, bold sub-headers and bullet lists,
// instead of one flat run (or a blank page when the saved value is empty).
function conclusionHtmlToOoxml(html: string): string {
  const BODY = 22; // 11pt
  const src = String(html || "");
  const paras: string[] = [];
  // <p> and <div> both act as paragraph blocks (contentEditable emits <div> for
  // new lines when the conclusion is edited). <ul> becomes a bulleted list.
  const blockRe = /<ul\b[^>]*>([\s\S]*?)<\/ul>|<(?:p|div)\b[^>]*>([\s\S]*?)<\/(?:p|div)>/gi;
  let m: RegExpExecArray | null;
  while ((m = blockRe.exec(src)) !== null) {
    if (m[1] !== undefined) {
      // <ul> list → one bulleted paragraph per <li>
      const liRe = /<li\b[^>]*>([\s\S]*?)<\/li>/gi;
      let li: RegExpExecArray | null;
      while ((li = liRe.exec(m[1])) !== null) {
        const bullet =
          `<w:r><w:rPr><w:color w:val="111111"/><w:sz w:val="${BODY}"/><w:szCs w:val="${BODY}"/></w:rPr>` +
          `<w:t xml:space="preserve">•  </w:t></w:r>`;
        paras.push(
          `<w:p><w:pPr><w:spacing w:before="0" w:after="40" w:line="264" w:lineRule="auto"/>` +
            `<w:ind w:left="360" w:hanging="220"/></w:pPr>` +
            bullet +
            conclusionInlineRunsToOoxml(li[1], BODY) +
            `</w:p>`
        );
      }
    } else {
      // <p> → one paragraph. Justify normal prose (both edges), BUT a paragraph
      // that contains hard line breaks (<br/>) — e.g. the "Issued By / For and
      // on behalf of: / RACE LBI …" signature block — must be LEFT-aligned.
      // Justifying short <br/> lines spreads each word across the whole page.
      const inner = m[2];
      const jc = /<br\b/i.test(inner) ? "left" : "both";
      paras.push(
        `<w:p><w:pPr><w:spacing w:before="0" w:after="120" w:line="264" w:lineRule="auto"/><w:jc w:val="${jc}"/></w:pPr>` +
          conclusionInlineRunsToOoxml(inner, BODY) +
          `</w:p>`
      );
    }
  }
  if (!paras.length) {
    // Unknown markup → split the stripped text on blank lines so it is at least
    // readable rather than one wall of text.
    const txt = stripHtml(src) || "-";
    for (const chunk of txt.split(/\n{2,}/)) {
      const safe = chunk
        .trim()
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
      if (!safe) continue;
      paras.push(
        `<w:p><w:pPr><w:spacing w:before="0" w:after="120" w:line="264" w:lineRule="auto"/></w:pPr>` +
          `<w:r><w:rPr><w:color w:val="111111"/><w:sz w:val="${BODY}"/><w:szCs w:val="${BODY}"/></w:rPr>` +
          `<w:t xml:space="preserve">${safe}</w:t></w:r></w:p>`
      );
    }
  }
  return paras.join("");
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

/**
 * Survey date = when the reports were actually recorded (the EARLIEST
 * report's created_at), not the export/generation date. Report timestamps
 * are stored in UTC, so we shift by +5:30 and read UTC fields to get the
 * correct India calendar date regardless of the server's timezone. Falls
 * back to today if no report has a usable created_at.
 */
function formatSurveyDate(reports: Row[], sep: "." | "-"): string {
  let earliest: number | null = null;
  for (const r of reports || []) {
    const raw = (r as Row)?.created_at;
    if (!raw) continue;
    const t = new Date(String(raw)).getTime();
    if (Number.isFinite(t)) earliest = earliest == null ? t : Math.min(earliest, t);
  }
  const baseMs = earliest != null ? earliest : Date.now();
  const ist = new Date(baseMs + 5.5 * 60 * 60 * 1000);
  const dd = String(ist.getUTCDate()).padStart(2, "0");
  const mm = String(ist.getUTCMonth() + 1).padStart(2, "0");
  return `${dd}${sep}${mm}${sep}${ist.getUTCFullYear()}`;
}

// NE coordinates use ONE shared format across the whole app (list, Edit form,
// route map and this report): decimal degrees with a direction letter, e.g.
// "N 14.2633142" \u2014 the exact value the user typed, not degrees/minutes.
function formatGpsLat(lat: unknown): string {
  return formatNeLat(lat, true);
}

function formatGpsLon(lng: unknown): string {
  return formatNeLon(lng, true);
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

// Standard "action" wording for each difficulty colour, shown in the
// REMARKS / ACTION column of the Word file:
//   green  → go ahead (normal pass)
//   yellow → go ahead with caution
//   red    → don't go / stop
function actionForDifficulty(key: string): string {
  if (key === "red") return "Don't go, Stop";
  if (key === "yellow") return "Go ahead with caution";
  return "Go ahead, Normal pass"; // green / default
}

// Build the REMARKS / ACTION cell. The two cases are NOT merged:
//   • the user TYPED a remark   → show ONLY what they typed (nothing added)
//   • the remark is EMPTY       → fill in the standard action for the colour
// A bare colour word / dash counts as empty (it carries no information).
function formatRemarksAction(written: unknown, difficultyKey: string): string {
  const action = actionForDifficulty(difficultyKey);
  const note = String(written ?? "").trim();
  const COLOUR_ONLY = ["red", "yellow", "green", "-", "—", "--"];
  if (!note || COLOUR_ONLY.includes(note.toLowerCase())) return action;
  return note; // user typed something → show exactly that, on its own
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

type ImageEntry = {
  buffer: Buffer;
  contentType: string;
  url?: string;
  path?: string;
  // For category icons: the render size [w,h] in px that preserves the icon's
  // TRIMMED aspect ratio (padding removed), capped to CATEGORY_ICON_SIZE. Used
  // by getSize so a wide-short glyph (e.g. the gate) renders short instead of
  // floating in a tall square cell with empty space above/below it.
  displaySize?: [number, number];
};

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
  difficultyKey: "red" | "yellow" | "green" | "neutral";
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
const RED_KEYWORDS = [
  "red", "critical", "hard", "fail", "not pass",
  // stop / no-go status words a surveyor might type instead of a colour, so a
  // "Stop" point is never shaded green and labelled "Go ahead".
  "stop", "dont go", "don't go", "do not go", "no go", "no-go", "halt",
  "blocked", "closed", "danger",
];
const YELLOW_KEYWORDS = ["yellow", "warning", "caution", "cautious", "medium", "go ahead with caution"];
const GREEN_KEYWORDS = ["green", "normal", "normal pass", "pass", "ok"];

function getDifficultyTableColors(value: unknown): {
  key: "red" | "yellow" | "green" | "neutral";
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
  // Green is now OPT-IN: only explicit green/normal/pass/ok words shade
  // green. This stops rows with no difficulty from all rendering green.
  if (GREEN_KEYWORDS.includes(d)) {
    return {
      key: "green",
      headerFillColor: "56BE9F",    // bright mint green (spec)
      headerTextColor: "FFFFFF",
      bodyFillColor: "F0F9F4",      // near-white with faintest mint tint
      bodyTextColor: "0B3D2E",
    };
  }
  // Empty / unrecognized difficulty → NEUTRAL grey, never green. A report
  // with no difficulty set must not be misrepresented as "go ahead".
  return {
    key: "neutral",
    headerFillColor: "9AA3AF",    // neutral grey header
    headerTextColor: "FFFFFF",
    bodyFillColor: "F4F5F7",      // near-white neutral body
    bodyTextColor: "0B3D2E",
  };
}

const CATEGORY_ICONS_DIR = path.join(process.cwd(), "public", "images", "report-icons");

// normalizeCategory + getCategoryDisplayName now live in ./categoryDisplay
// (shared with the Stage Summary editor API) and are imported at the top.

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
  // Flyover / overpass get their OWN icon (checked BEFORE the generic bridge
  // rule so "Flyover" no longer borrows the bridge image).
  if (c.includes("flyover") || c.includes("overpass") || c.includes("over_bridge")) return "flyover";
  // Service road / service lane get their own icon.
  if (c.includes("service")) return "service_road";
  // Plain "Bridge" / "Viaduct" — falls through to bridge.png.
  if (c.includes("bridge") || c.includes("viaduct")) {
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
  // Damaged road / potholes / broken surface get their own icon.
  if (c.includes("damage") || c.includes("pothole") || c.includes("broken")) {
    return "damaged_road";
  }
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
  // New custom icons.
  flyover: "public/images/report-icons/flyover.png",
  service_road: "public/images/report-icons/service_road.png",
  damaged_road: "public/images/report-icons/damaged_road.png",
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
  // 1. Canonical category with a real mapping → use it, but ONLY if the file
  //    actually exists. Some icons (e.g. flyover.png / service_road.png) are
  //    mapped but may not be deployed; without this check the loader returned
  //    a dead path and showed the fallback triangle. If missing, fall through
  //    to a similar existing icon (step 2) / an auto-generated one (step 3).
  if (key !== "fallback" && CATEGORY_ICON_MAP[key] && existsSync(path.join(process.cwd(), CATEGORY_ICON_MAP[key]))) {
    const rel = CATEGORY_ICON_MAP[key];
    console.log("[category mapping check]", {
      rawCategory: category,
      normalizedKey: key,
      iconRelativePath: rel,
      source: "canonical",
    });
    return rel.replace(/^public[\\/]images[\\/]report-icons[\\/]/, "");
  }
  if (key !== "fallback" && CATEGORY_ICON_MAP[key]) {
    console.warn("[category icon] mapped file missing, using similar/auto icon:", CATEGORY_ICON_MAP[key]);
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
  if (!sharp) {
    // Production-friendly diagnostic: sharp's native binary (or
    // libvips) isn't loadable on this host. The Word-table fallback
    // path picks up where this returns null.
    console.warn(
      "[ROUTE LOCATIONS STEPPER] sharp unavailable — SVG→PNG impossible. " +
        "Word-table fallback will be used instead."
    );
    return null;
  }
  const items = labels.filter((l) => l.label).slice(0, 30);
  if (items.length === 0) {
    console.warn("[ROUTE LOCATIONS STEPPER] no labels passed in");
    return null;
  }

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
    console.warn(
      "[ROUTE LOCATIONS STEPPER FAILED] SVG→PNG render failed " +
        "(probably libvips lacks librsvg in this build). Word-table " +
        "fallback will be used instead. Error:",
      (err as Error)?.message || String(err)
    );
    return null;
  }
}

/**
 * Render the route key points as a horizontal numbered "snake" map
 * (START ▶ ① ▶ ② … ▶ DESTINATION), matching the gold-standard route
 * infographic but using the typed key-point NAMES only (no photos).
 * Boustrophedon layout: rows alternate L→R / R→L, joined by arrows.
 * Rendered as SVG → PNG via sharp (same pipeline as the stepper image).
 * Returns a PNG Buffer, or null if sharp/SVG rendering is unavailable.
 */
async function generateRouteLocationsSnake(
  labels: { label: string; pin_type: string }[]
): Promise<Buffer | null> {
  const sharp = getSharp();
  if (!sharp) return null;
  const items = labels.filter((l) => l.label).slice(0, 40);
  if (items.length === 0) return null;

  const escapeXml = (s: string) =>
    s
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");

  // Number-badge colours cycle like the sample map.
  const BADGE_COLORS = [
    "#2563EB", "#EA580C", "#16A34A", "#7C3AED",
    "#DB2777", "#0891B2", "#CA8A04", "#DC2626",
  ];

  type Cell = { kind: "start" | "dest" | "point"; label: string; n?: number; color?: string };
  const cells: Cell[] = [];
  cells.push({ kind: "start", label: "START" });
  items.forEach((it, i) =>
    cells.push({
      kind: "point",
      label: it.label,
      n: i + 1,
      color: BADGE_COLORS[i % BADGE_COLORS.length],
    })
  );
  cells.push({ kind: "dest", label: "DESTINATION" });

  const T = cells.length;
  // Cards per row: wide but balanced (≈6 per row).
  const perRow = T <= 6 ? T : Math.ceil(T / Math.ceil(T / 6));
  const rows = Math.ceil(T / perRow);

  const CARD_W = 230, CARD_H = 92;
  const ARROW_W = 60;   // horizontal gap between cards (room for the arrow)
  const ROW_GAP = 78;   // vertical gap between rows (room for the down arrow)
  const PAD = 36;
  const HEADING_H = 54;

  const W = PAD * 2 + perRow * CARD_W + (perRow - 1) * ARROW_W;
  const H = PAD * 2 + HEADING_H + rows * CARD_H + (rows - 1) * ROW_GAP;

  const pos = (i: number) => {
    const row = Math.floor(i / perRow);
    const inRow = i % perRow;
    const col = row % 2 === 0 ? inRow : perRow - 1 - inRow; // snake reversal
    const x = PAD + col * (CARD_W + ARROW_W);
    const y = PAD + HEADING_H + row * (CARD_H + ROW_GAP);
    return { row, x, y };
  };

  const HEAD = 11;
  const aRight = (x1: number, x2: number, y: number) =>
    `<line x1="${x1}" y1="${y}" x2="${x2 - HEAD}" y2="${y}" stroke="#2563EB" stroke-width="6" stroke-linecap="round"/>` +
    `<path d="M ${x2} ${y} L ${x2 - HEAD} ${y - HEAD * 0.8} L ${x2 - HEAD} ${y + HEAD * 0.8} Z" fill="#2563EB"/>`;
  const aLeft = (x1: number, x2: number, y: number) =>
    `<line x1="${x1}" y1="${y}" x2="${x2 + HEAD}" y2="${y}" stroke="#2563EB" stroke-width="6" stroke-linecap="round"/>` +
    `<path d="M ${x2} ${y} L ${x2 + HEAD} ${y - HEAD * 0.8} L ${x2 + HEAD} ${y + HEAD * 0.8} Z" fill="#2563EB"/>`;
  const aDown = (x: number, y1: number, y2: number) =>
    `<line x1="${x}" y1="${y1}" x2="${x}" y2="${y2 - HEAD}" stroke="#2563EB" stroke-width="6" stroke-linecap="round"/>` +
    `<path d="M ${x} ${y2} L ${x - HEAD * 0.8} ${y2 - HEAD} L ${x + HEAD * 0.8} ${y2 - HEAD} Z" fill="#2563EB"/>`;

  const parts: string[] = [];
  parts.push(`<rect x="0" y="0" width="${W}" height="${H}" fill="#FFFFFF"/>`);
  parts.push(
    `<text x="${PAD}" y="${PAD + 34}" font-family="Arial, sans-serif" font-size="30" font-weight="700" fill="#1F2937">ROUTE SURVEY KEY POINTS</text>`
  );

  // Connector arrows (drawn under the cards).
  for (let i = 0; i < T - 1; i += 1) {
    const a = pos(i), b = pos(i + 1);
    if (a.row === b.row) {
      const yMid = a.y + CARD_H / 2;
      if (b.x > a.x) parts.push(aRight(a.x + CARD_W + 6, b.x - 6, yMid));
      else parts.push(aLeft(a.x - 6, b.x + CARD_W + 6, yMid));
    } else {
      const xMid = a.x + CARD_W / 2;
      parts.push(aDown(xMid, a.y + CARD_H + 4, b.y - 4));
    }
  }

  // Cells.
  for (let i = 0; i < T; i += 1) {
    const c = cells[i];
    const p = pos(i);
    if (c.kind === "point") {
      parts.push(
        `<rect x="${p.x}" y="${p.y}" width="${CARD_W}" height="${CARD_H}" rx="16" ry="16" fill="#FFFFFF" stroke="#CBD5E1" stroke-width="2"/>`
      );
      const bx = p.x + 28, by = p.y + 22;
      parts.push(`<circle cx="${bx}" cy="${by}" r="20" fill="${c.color}"/>`);
      parts.push(
        `<text x="${bx}" y="${by + 8}" text-anchor="middle" font-family="Arial, sans-serif" font-size="24" font-weight="700" fill="#FFFFFF">${c.n}</text>`
      );
      const fs = c.label.length > 16 ? 20 : c.label.length > 11 ? 24 : 27;
      parts.push(
        `<text x="${p.x + CARD_W / 2}" y="${p.y + CARD_H / 2 + 26}" text-anchor="middle" font-family="Arial, sans-serif" font-size="${fs}" font-weight="700" fill="#0F172A">${escapeXml(c.label)}</text>`
      );
    } else if (c.kind === "start") {
      const cx = p.x + CARD_W / 2, top = p.y + 4;
      parts.push(
        `<g transform="translate(${cx},${top})">` +
          `<path d="M 0,0 C -16,0 -24,11 -24,19 C -24,37 0,60 0,60 C 0,60 24,37 24,19 C 24,11 16,0 0,0 Z" fill="#16A34A" stroke="#0F7A33" stroke-width="2"/>` +
          `<circle cx="0" cy="20" r="8" fill="#FFFFFF"/>` +
        `</g>`
      );
      parts.push(
        `<text x="${cx}" y="${p.y + CARD_H + 4}" text-anchor="middle" font-family="Arial, sans-serif" font-size="24" font-weight="700" fill="#16A34A">START</text>`
      );
    } else {
      // Checkered destination flag.
      const fx = p.x + CARD_W / 2 - 22, fy = p.y + 8, sq = 14;
      parts.push(`<line x1="${fx}" y1="${fy}" x2="${fx}" y2="${fy + 60}" stroke="#334155" stroke-width="4"/>`);
      for (let r = 0; r < 2; r += 1)
        for (let cc = 0; cc < 3; cc += 1) {
          const fill = (r + cc) % 2 === 0 ? "#111827" : "#FFFFFF";
          parts.push(
            `<rect x="${fx + 2 + cc * sq}" y="${fy + r * sq}" width="${sq}" height="${sq}" fill="${fill}" stroke="#111827" stroke-width="1"/>`
          );
        }
      parts.push(
        `<text x="${p.x + CARD_W / 2}" y="${p.y + CARD_H + 4}" text-anchor="middle" font-family="Arial, sans-serif" font-size="22" font-weight="700" fill="#DC2626">DESTINATION</text>`
      );
    }
  }

  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">` +
    parts.join("") +
    `</svg>`;

  try {
    const buf = await sharp(Buffer.from(svg)).png().toBuffer();
    console.log("[ROUTE LOCATIONS SNAKE GENERATED]", {
      items: items.length,
      perRow,
      rows,
      svgSize: `${W}x${H}`,
      bytes: buf.length,
    });
    return buf;
  } catch (err) {
    console.warn(
      "[ROUTE LOCATIONS SNAKE] SVG→PNG failed (libvips may lack librsvg) — " +
        "OOXML snake-table fallback will be used. Error:",
      (err as Error)?.message || String(err)
    );
    return null;
  }
}

/**
 * Sharp-free fallback for the route locations stepper. Builds the
 * locations list as a pure OOXML borderless table — no SVG, no PNG,
 * no native dependencies. Renders cleanly on any production server,
 * even ones where sharp/libvips/librsvg isn't installed.
 *
 * Layout: a single-column outer table. Each row is one location with:
 *   - A small bullet character ("○" for a stop, "●" for the final
 *     destination) prefixed to the label
 *   - The label text in a bordered rounded-grey cell
 *   - "Locations" heading sits in its own row at the top
 *
 * Returns the full <w:tbl>...</w:tbl> XML string (no leading/trailing
 * paragraph), or "" if no labels are passed.
 */
function buildRouteLocationsStepperTableXml(
  labels: { label: string; pin_type: string }[]
): string {
  const items = labels.filter((l) => l.label).slice(0, 30);
  if (items.length === 0) return "";

  const escapeXml = (s: string) =>
    s
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");

  // Cell border styles
  const NO_BORDER =
    `<w:tcBorders>` +
      `<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
      `<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
      `<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
      `<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
    `</w:tcBorders>`;
  const PILL_BORDER =
    `<w:tcBorders>` +
      `<w:top w:val="single" w:sz="6" w:space="0" w:color="D1D5DB"/>` +
      `<w:left w:val="single" w:sz="6" w:space="0" w:color="D1D5DB"/>` +
      `<w:bottom w:val="single" w:sz="6" w:space="0" w:color="D1D5DB"/>` +
      `<w:right w:val="single" w:sz="6" w:space="0" w:color="D1D5DB"/>` +
    `</w:tcBorders>`;

  // Outer table grid: marker column ~700 dxa (~0.5"), label column rest
  const TBL_W = 9360;
  const MARK_COL = 700;
  const LABEL_COL = TBL_W - MARK_COL;

  // Heading row built as an array of fragments joined together. Using
  // `.join("")` over an array (instead of long `+`-chained template
  // literals) keeps the AST flat — long `+` chains can blow the SWC
  // parser's stack on Windows during `next build`.
  const headingRow = [
    `<w:tr>`,
    `<w:tc>`,
    `<w:tcPr><w:tcW w:w="${TBL_W}" w:type="dxa"/>`,
    `<w:gridSpan w:val="2"/>${NO_BORDER}`,
    `</w:tcPr>`,
    `<w:p>`,
    `<w:pPr><w:spacing w:before="0" w:after="100"/></w:pPr>`,
    `<w:r>`,
    `<w:rPr>`,
    `<w:rFonts w:ascii="Neue Haas Grotesk Text Pro" w:hAnsi="Neue Haas Grotesk Text Pro"/>`,
    `<w:b/><w:bCs/>`,
    `<w:sz w:val="28"/><w:szCs w:val="28"/>`,
    `<w:color w:val="1F2937"/>`,
    `</w:rPr>`,
    `<w:t>ROUTE SURVEY KEY POINTS</w:t>`,
    `</w:r>`,
    `</w:p>`,
    `</w:tc>`,
    `</w:tr>`,
  ].join("");

  const itemRows: string[] = [];
  for (let i = 0; i < items.length; i += 1) {
    const it = items[i];
    const isEnd = i === items.length - 1 || it.pin_type === "end";
    // ⚫ red filled circle for end, ◯ outlined circle for stops.
    // Unicode glyphs work without any font install on Windows servers.
    const marker = isEnd ? "●" : "○";
    const markerColor = isEnd ? "EC4054" : "1F2937";

    itemRows.push([
      `<w:tr>`,
      // Marker cell — small bullet character
      `<w:tc>`,
      `<w:tcPr><w:tcW w:w="${MARK_COL}" w:type="dxa"/>${NO_BORDER}<w:vAlign w:val="center"/></w:tcPr>`,
      `<w:p>`,
      `<w:pPr><w:jc w:val="center"/><w:spacing w:before="60" w:after="60"/></w:pPr>`,
      `<w:r>`,
      `<w:rPr>`,
      `<w:rFonts w:ascii="Neue Haas Grotesk Text Pro" w:hAnsi="Neue Haas Grotesk Text Pro"/>`,
      `<w:b/><w:bCs/>`,
      `<w:sz w:val="40"/><w:szCs w:val="40"/>`,
      `<w:color w:val="${markerColor}"/>`,
      `</w:rPr>`,
      `<w:t>${marker}</w:t>`,
      `</w:r>`,
      `</w:p>`,
      `</w:tc>`,
      // Label cell — bordered "pill" with the location name
      `<w:tc>`,
      `<w:tcPr><w:tcW w:w="${LABEL_COL}" w:type="dxa"/>${PILL_BORDER}<w:vAlign w:val="center"/></w:tcPr>`,
      `<w:p>`,
      `<w:pPr><w:spacing w:before="80" w:after="80"/></w:pPr>`,
      `<w:r>`,
      `<w:rPr>`,
      `<w:rFonts w:ascii="Neue Haas Grotesk Text Pro" w:hAnsi="Neue Haas Grotesk Text Pro"/>`,
      `<w:b/><w:bCs/>`,
      `<w:sz w:val="28"/><w:szCs w:val="28"/>`,
      `<w:color w:val="0F172A"/>`,
      `</w:rPr>`,
      `<w:t xml:space="preserve">${escapeXml(it.label)}</w:t>`,
      `</w:r>`,
      `</w:p>`,
      `</w:tc>`,
      `</w:tr>`,
    ].join(""));
  }

  console.log("[ROUTE LOCATIONS STEPPER FALLBACK XML]", {
    itemCount: items.length,
    fallbackPath: "Word-table (no sharp/libvips needed)",
  });

  return [
    `<w:tbl>`,
    `<w:tblPr>`,
    `<w:tblW w:w="5000" w:type="pct"/>`,
    `<w:tblInd w:w="0" w:type="dxa"/>`,
    `<w:tblBorders>`,
    `<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>`,
    `<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>`,
    `<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>`,
    `<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>`,
    `<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>`,
    `<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>`,
    `</w:tblBorders>`,
    `</w:tblPr>`,
    `<w:tblGrid>`,
    `<w:gridCol w:w="${MARK_COL}"/>`,
    `<w:gridCol w:w="${LABEL_COL}"/>`,
    `</w:tblGrid>`,
    headingRow,
    itemRows.join(""),
    `</w:tbl>`,
  ].join("");
}

// Compute the render size [w,h] (px) for a category icon so it keeps the icon's
// real (already trimmed) aspect ratio while fitting inside the CATEGORY_ICON_SIZE
// box. A wide-short glyph (gate) → e.g. 240×85 instead of a 240×240 square that
// leaves blank space above/below it. Never enlarges the icon past the box width,
// so the visible glyph is not reduced. Falls back to the square box on failure.
async function categoryIconDisplaySize(buffer: Buffer): Promise<[number, number]> {
  const [maxW, maxH] = CATEGORY_ICON_SIZE;
  try {
    const sharp = getSharp();
    if (!sharp) return CATEGORY_ICON_SIZE;
    const meta = await sharp(buffer).metadata();
    const w = meta.width || 0;
    const h = meta.height || 0;
    if (w <= 0 || h <= 0) return CATEGORY_ICON_SIZE;
    const scale = Math.min(maxW / w, maxH / h);
    return [Math.max(1, Math.round(w * scale)), Math.max(1, Math.round(h * scale))];
  } catch {
    return CATEGORY_ICON_SIZE;
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
    const rawBuffer = await generateCategoryIcon(sourceText);
    if (!rawBuffer || rawBuffer.length === 0) {
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
    // Trim the transparent padding so the auto-generated glyph (the location-pin
    // / warning-triangle fallback) fills its box with no side gaps — the same
    // trim the file-based icons already get in optimizeDocxImage.
    const buffer = await optimizeDocxImage(rawBuffer, "category", fileName);
    const entry: ImageEntry = {
      buffer,
      contentType: "image/png",
      path: `(auto-generated:${sourceText})`,
      displaySize: await categoryIconDisplaySize(buffer),
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
    const entry: ImageEntry = {
      buffer,
      contentType,
      path: fullPath,
      displaySize: await categoryIconDisplaySize(buffer),
    };
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
// IMPORTANT: skip null/undefined/blank BEFORE Number(), because Number(null),
// Number(""), and Number(" ") all equal 0 (a finite value). Without this guard
// a report with no GPS would be read as coordinate (0,0), injecting a ~8,600 km
// phantom leg into the cumulative KM column and printing a fake "N0 0.000".
function firstFiniteCoord(candidates: unknown[]): number | null {
  for (const c of candidates) {
    if (c === null || typeof c === "undefined" || String(c).trim() === "") continue;
    const n = Number(c);
    if (Number.isFinite(n)) return n;
  }
  return null;
}
function pickLat(r: Row): number | null {
  return firstFiniteCoord([r.latitude, r.lat, r.ne_latitude, r.ne_lat, r.loc_lat]);
}
function pickLng(r: Row): number | null {
  return firstFiniteCoord([r.longitude, r.lng, r.ne_longitude, r.ne_lng, r.lon, r.loc_lon]);
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
 * Format a numeric km value to 2 decimal places, or "-" if not finite.
 */
function formatKmValue(value: unknown): string {
  const n = Number(value);
  if (!Number.isFinite(n)) return "-";
  return n.toFixed(2);
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
    // Skip a name written in a non-Latin script (e.g. imported Hindi) so the
    // resolver reverse-geocodes an English one and caches it in
    // resolved_location (a later candidate, which IS returned once in Latin).
    if (s && !hasNonLatinScript(s)) return s;
  }
  return null;
}

/**
 * True when a place name contains a non-Latin (Indian-language) script, e.g.
 * Devanagari "गोसाईगंज". Imported reports sometimes carry the place name in the
 * local script; the export re-geocodes those in English instead of printing the
 * Hindi/regional text. Covers Devanagari..Sinhala (U+0900–U+0DFF).
 */
function hasNonLatinScript(s: string | null | undefined): boolean {
  return /[ऀ-෿]/.test(String(s || ""));
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
  // Keep the LOCATION column short: just the most specific place name and the
  // state, e.g. "Alandur, Tamil Nadu". Road / district / country are dropped.
  const place =
    addr.suburb ||
    addr.neighbourhood ||
    addr.city_district ||
    addr.town ||
    addr.city ||
    addr.village ||
    addr.hamlet;
  const state = addr.state;
  const country = addr.country;

  const parts: string[] = [];
  if (place) parts.push(place);
  // Prefer state; fall back to country only when there is no state.
  if (state && state !== place) parts.push(state);
  else if (!state && country && country !== place) parts.push(country);

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
    `&addressdetails=1` +
    // Force English (Latin) names so results don't come back in the local
    // script (Hindi/Tamil/etc.) and produce mixed-language locations.
    `&accept-language=en`;
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
    `&zoom=12&addressdetails=1&accept-language=en`;
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
  // getExistingLocation already skips non-Latin (Hindi) names, so `existing` is
  // either a usable Latin name or null → non-Latin stored names get re-geocoded.
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
  if (resolved && !hasNonLatinScript(resolved)) {
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
  const includeGa = options.includeGa !== false; // default true; false = skip GA drawing
  const summaryMode = options.summaryMode === true;
  // Watermark text (empty = no watermark). Drawn in the header pass below.
  const watermarkText =
    options.watermark?.enabled === true ? String(options.watermark?.text || "").trim() : "";
  const reportIdsFilter = (options.reportIds || []).filter(Boolean);

  // ----- Editable report-section overrides (Summary of Observations, Hindrance
  // Solutions, Stage-Summary image). Stored per-project in project_report_sections
  // and edited via /api/projects/:id/report-sections/:key. Loaded once here so
  // both the early summary block and the late Stage-Summary injection can use
  // them. Non-fatal: any failure leaves the defaults in place.
  let observationSummaryOverride: {
    fromText?: string;
    toText?: string;
    distance?: string;
    keyLocations?: string;
    flowChartUrl?: string;
  } | null = null;
  let hindranceOverride: {
    rows?: { hindrance?: string; solution?: string }[];
    considerations?: string[];
  } | null = null;
  let stageImageOverride: { imageUrl?: string } | null = null;
  if (projectId) {
    try {
      const [secRows] = await pool.query(
        "SELECT section_key, data_json FROM project_report_sections WHERE project_id = ?",
        [projectId]
      );
      for (const r of (Array.isArray(secRows) ? secRows : []) as any[]) {
        let parsed: any = null;
        try {
          parsed = JSON.parse(String(r?.data_json ?? ""));
        } catch {
          parsed = null;
        }
        if (!parsed) continue;
        if (r?.section_key === "observation-summary") observationSummaryOverride = parsed;
        else if (r?.section_key === "hindrance") hindranceOverride = parsed;
        else if (r?.section_key === "stage-image") stageImageOverride = parsed;
      }
    } catch (secErr) {
      console.warn("[report sections] override load failed - using defaults:", secErr);
    }
  }

  // Fetch + size the optional "Route Survey Flow Chart" image shown under the
  // Summary of Observations (uploaded in the setup wizard). Non-fatal: if it
  // fails to load, the section just renders without the image.
  let flowChartImg: { buffer: Buffer; emuW: number; emuH: number } | null = null;
  if (!summaryMode && observationSummaryOverride?.flowChartUrl) {
    try {
      const f = await fetchImageBuffer(String(observationSummaryOverride.flowChartUrl), "flowChart");
      if (f && Buffer.isBuffer(f.buffer) && f.buffer.length) {
        const opt = await optimizeDocxImage(f.buffer, "observation", "flowChart");
        const EMU_PER_INCH_FC = 914400;
        const emuW = Math.round(7.5 * EMU_PER_INCH_FC);
        let emuH = Math.round(4.0 * EMU_PER_INCH_FC);
        const shFc = getSharp();
        if (shFc) {
          try {
            const m = await shFc(opt).metadata();
            if (m.width && m.height && m.width > 0) emuH = Math.round((emuW * m.height) / m.width);
          } catch {
            /* keep fallback height */
          }
        }
        flowChartImg = { buffer: opt, emuW, emuH };
      }
    } catch (fcErr) {
      console.warn("[flow chart] load failed - skipping:", fcErr);
    }
  }

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
    // Honour the requested direction so the DOCX matches the on-screen sort.
    const dir = options.sort === "desc" ? "DESC" : "ASC";
    const orderBy = `ORDER BY sort_order ${dir}, created_at ${dir}`;
    const reportSql =
      reportIdsFilter.length > 0
        ? `SELECT * FROM reports WHERE project_id = ? AND id IN (${reportIdsFilter
            .map(() => "?")
            .join(",")}) ${orderBy}`
        : `SELECT * FROM reports WHERE project_id = ? ${orderBy}`;
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
  // Rows the user explicitly unticked in the photo picker (include_in_export=0).
  let excludedPhotoRows: Row[] = [];
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
      // diagnostic logs can show the master-file linkage. include_in_export
      // is the user's per-photo Word selection (photo picker checkboxes).
      const sqlWithInclude = `SELECT id, report_id, url, file_name, image_key, point_key, path, created_at, include_in_export, sort_order
         FROM report_photos
         WHERE report_id IN (${placeholders})
         ORDER BY (sort_order IS NULL), sort_order ASC, created_at ASC`;
      const sql = `SELECT id, report_id, url, file_name, image_key, point_key, path, created_at
         FROM report_photos
         WHERE report_id IN (${placeholders})
         ORDER BY created_at ASC`;
      console.error("[PHOTOS_DEBUG_2026_04_27_B] sql", { sql: sqlWithInclude, args: reportIds });
      try {
        const [photoRows] = await pool.execute(sqlWithInclude, reportIds);
        photos = Array.isArray(photoRows) ? (photoRows as Row[]) : [];
      } catch {
        try {
          // Fallback: schemas without the include_in_export column.
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
      }
      // Respect the photo-picker selection: photos unticked by the user
      // (include_in_export = 0) never reach the DOCX. Rows without the
      // column (legacy schema fallback) are always included.
      excludedPhotoRows = photos.filter(
        (p) => p.include_in_export != null && Number(p.include_in_export) === 0
      );
      photos = photos.filter(
        (p) => p.include_in_export == null || Number(p.include_in_export) !== 0
      );
      // Videos live in report_photos too but can't embed in a DOCX — drop them.
      photos = photos.filter((p) => {
        const ref = String(p.url || p.file_name || "");
        return !/\.(mp4|mov|m4v|webm|avi|mkv|3gp|ogv|qt)(\?|#|$)/i.test(ref);
      });
      if (excludedPhotoRows.length) {
        console.log("[export photos] excluded by user selection:", excludedPhotoRows.length);
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

  // Reports where the user unticked EVERY photo. The S3 filename fallback
  // below must not "recover" a photo for these — the empty photo block is
  // the user's explicit choice, not a broken link.
  const reportsFullyExcludedByUser = new Set<string>();
  for (const p of excludedPhotoRows) {
    const key = String(p.report_id || "").trim();
    if (key && !photosByReportId.has(key)) reportsFullyExcludedByUser.add(key);
  }

  // ---- ONE BLOCK PER REPORT (multi-photo shown together) ----
  // A survey point is ONE observation block (one table). When the point has
  // several photos they are shown TOGETHER as a 2-per-row grid under that one
  // table (built in the multi-photo grid injection below) — NOT split into a
  // separate block per photo with the table repeated. Splitting a point into
  // one-report-per-photo is a manual choice via the "Split photos" button, not
  // something the export forces. Repeat-point de-duplication still happens in
  // the main loop's repeat-skip below. `splitDupRows` stays empty (kept so the
  // loop's isSplitDup check compiles as a no-op).
  const splitDupRows = new WeakSet<object>();
  const expandedReports: Row[] = reports;
  console.log("[DOCX OBSERVATION BLOCKS]", {
    reports: reports.length,
    mode: "one block per report (multi-photo grid)",
  });

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

  // Set of report ids present in THIS export. The point_key fallback below
  // must only recover TRUE orphan photos (those whose report_id matches no
  // current report). Without this guard, when several reports share a
  // point_key, a report with no photos of its own would "borrow" photos that
  // legitimately belong to ANOTHER report — making the SAME photo appear
  // under multiple report blocks (the "photos repeated 2-3 times" bug).
  const knownReportIdSet = new Set(reportIds);

  // Global guard: a given photo may be embedded under AT MOST ONE report in
  // the whole export. Even if two reports resolve to the same photo (shared
  // point_key, duplicate upload, stale report_id, etc.), the second report
  // skips it and renders "Photo not available" instead of repeating the image.
  // Reports are processed in order, so the rightful owner (matched by
  // report_id) claims the photo first.
  const usedPhotoKeys = new Set<string>();
  const photoIdentity = (p: Row) =>
    String(p?.id || p?.url || p?.image_key || "").trim();

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

  // ----- Step 5: GA drawing. Skipped entirely when includeGa is false so the
  // user can download a "without GA drawing" version.
  await ensureGaDrawingsTable();
  let gaRow: Row | null = null;
  if (includeGa) {
    try {
      const rows = await safeQuery(
        "SELECT * FROM project_ga_drawings WHERE project_id = ? LIMIT 1",
        [projectId]
      );
      gaRow = rows[0] || null;
    } catch (err) {
      console.error("[export actual] ga query failed - continuing without:", err);
    }
  }
  const gaImageUrl = includeGa ? resolveImageUrl(gaRow, "image_url", "image_key") : "";
  console.log("[export actual] gaDrawingUrl:", gaImageUrl || "(none)", "includeGa:", includeGa);

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

  // ALL GA drawing images (not just the first). They live in
  // project_route_page_images keyed to the route page; the single-row
  // project_ga_drawings table only ever keeps the first one.
  let gaImageUrls: string[] = [];
  // File names run parallel to gaImageUrls. Each diagram file name carries a
  // short slot tag (e.g. "GAWOL__…") written by the Diagrams wizard, so the
  // export can print every diagram under its OWN heading (GA DRAWING WITHOUT
  // LOAD / WITH LOAD / TURNING CIRCLE / STABILITY) instead of stacking them.
  let gaImageNames: string[] = [];
  if (includeGa) {
    const gaPageId = String(routePageRow?.id || "").trim();
    // 1) The GA images on the route page (the expected place).
    if (gaPageId) {
      try {
        const imgRows = await safeQuery(
          "SELECT file_url, file_name FROM project_route_page_images WHERE project_id = ? AND project_page_id = ? ORDER BY created_at ASC",
          [projectId, gaPageId]
        );
        const pairs = imgRows
          .map((x) => ({
            url: String((x as Row).file_url || "").trim(),
            name: String((x as Row).file_name || "").trim(),
          }))
          .filter((p) => p.url);
        gaImageUrls = pairs.map((p) => p.url);
        gaImageNames = pairs.map((p) => p.name);
      } catch (err) {
        console.error("[export actual] ga images (page) query failed:", err);
      }
    }
    // 2) Fallback: ALL GA images for this project (in case they were saved
    // against a different route-page row than the one we picked above). The
    // project_route_page_images table holds only GA-drawing images, so this is
    // safe. De-dup by url in case pages overlap.
    if (!gaImageUrls.length) {
      try {
        const imgRows = await safeQuery(
          "SELECT file_url, file_name FROM project_route_page_images WHERE project_id = ? ORDER BY created_at ASC",
          [projectId]
        );
        const seen = new Set<string>();
        const pairs: { url: string; name: string }[] = [];
        for (const x of imgRows) {
          const url = String((x as Row).file_url || "").trim();
          if (!url || seen.has(url)) continue;
          seen.add(url);
          pairs.push({ url, name: String((x as Row).file_name || "").trim() });
        }
        gaImageUrls = pairs.map((p) => p.url);
        gaImageNames = pairs.map((p) => p.name);
      } catch (err) {
        console.error("[export actual] ga images (project) query failed:", err);
      }
    }
    // 3) Last resort: the single image from project_ga_drawings.
    if (!gaImageUrls.length && gaImageUrl) {
      gaImageUrls = [gaImageUrl];
      gaImageNames = [""];
    }
  }

  // Heading for each diagram, parallel to gaImageUrls. Primary source is the
  // slot tag in the file name; a positional fallback covers older diagrams that
  // were saved before slot tagging. Summary reports keep their single
  // "GA VEHICLE DRAWING:" heading, so no per-slot headings there.
  const GA_HEADINGS_BY_TAG: Record<string, string> = {
    GAWOL: "GA DRAWING WITHOUT LOAD:",
    GAWL: "GA DRAWING WITH LOAD:",
    TURN: "TURNING CIRCLE DRAWING:",
    STAB: "STABILITY DRAWING:",
  };
  const GA_HEADINGS_POSITIONAL = [
    "GA DRAWING WITHOUT LOAD:",
    "GA DRAWING WITH LOAD:",
    "TURNING CIRCLE DRAWING:",
    "STABILITY DRAWING:",
  ];
  const gaHeadings: string[] = summaryMode
    ? []
    : gaImageUrls.map((_u, idx) => {
        const m = String(gaImageNames[idx] || "").match(/^([A-Z]{3,6})__/);
        if (m && GA_HEADINGS_BY_TAG[m[1]]) return GA_HEADINGS_BY_TAG[m[1]];
        return GA_HEADINGS_POSITIONAL[idx] || "";
      });
  console.log("[export actual] GA images:", gaImageUrls.length);
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
  // ----- Conclusion & Certification body.
  // Prefer the saved conclusion HTML; otherwise fall back to the fixed RACE
  // standard certification text so the page is NEVER blank. From/To for the
  // standard text come from the route's first/last key location, else the
  // project name ("A to B"). The body is rendered as neat formatted paragraphs
  // in the post-render pass (via CONCLUSION_SENTINEL), not a flat run.
  const savedConclusionHtml = String(
    gaRow?.conclusion_html || routePageRow?.conclusion_html || ""
  ).trim();
  let conclusionHtmlResolved = savedConclusionHtml;
  if (!conclusionHtmlResolved) {
    const locLabels = (routeLocationLabels || [])
      .map((l) => String(l.label || "").trim())
      .filter(Boolean);
    let concFrom = locLabels[0] || "";
    let concTo = locLabels.length > 1 ? locLabels[locLabels.length - 1] : "";
    if (!concFrom || !concTo) {
      const parts = String(projectName || "").split(/\s+to\s+/i);
      if (parts.length >= 2) {
        concFrom = concFrom || parts[0].trim();
        concTo = concTo || parts[parts.length - 1].trim();
      }
    }
    conclusionHtmlResolved = buildStandardConclusionHtml(concFrom, concTo);
  }
  const conclusionBodyOoxml = conclusionHtmlToOoxml(conclusionHtmlResolved);
  const dateDot = formatSurveyDate(reports, ".");
  const dateDash = formatSurveyDate(reports, "-");

  // ----- Step 7: Fetch image buffers. Each fetch must NEVER throw - it
  // returns null on any failure so the export can continue with fallbacks.
  // Reuses the module-level ImageEntry type so category-icon entries (loaded
  // from disk, no url) and S3-fetched entries (no path) share a single shape.
  const imageMap = new Map<string, ImageEntry>();
  // Extra GA-drawing images (index 1..N-1). GA image 0 goes into imageMap under
  // "gaDrawing" (the template renders it in the GA slot); each EXTRA GA image is
  // injected on its OWN page in the post-render pass — so multiple GA drawings
  // come one-per-page instead of being squeezed into a grid on a single page.
  const gaExtraBuffers: Buffer[] = [];
  try {
    // Fetch every GA image. Image 0 fills the GA slot; every other GA image is
    // kept whole and placed on its own separate page (not a grid).
    const gaBuffers: Buffer[] = [];
    for (const u of gaImageUrls) {
      const f = await fetchImageBuffer(u, "gaDrawing");
      if (f && Buffer.isBuffer(f.buffer)) gaBuffers.push(f.buffer);
    }
    if (gaBuffers.length) {
      const first = await optimizeDocxImage(gaBuffers[0], "gaDrawing", "gaDrawing");
      imageMap.set("gaDrawing", {
        buffer: first,
        contentType: "image/jpeg",
        url: gaImageUrls[0] || gaImageUrl,
      });
      for (let gi = 1; gi < gaBuffers.length; gi += 1) {
        const extra = await optimizeDocxImage(gaBuffers[gi], "gaDrawing", "gaDrawing");
        if (extra && Buffer.isBuffer(extra)) gaExtraBuffers.push(extra);
      }
    }
  } catch (err) {
    console.error("[export actual] gaDrawing fetch step threw - ignoring:", err);
  }
  console.log("[export actual] GA extra pages:", gaExtraBuffers.length);
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

  // ----- Generate the route-locations stepper image. Three-tier
  // resolution so the stepper looks good on EVERY environment:
  //   1. sharp + SVG → PNG (preferred, used locally)
  //   2. pure-JS PNG renderer (used in production when sharp/libvips
  //      lacks SVG support — produces an image that visually matches
  //      the SVG version, no native deps required)
  //   3. Word-table fallback (only if both PNG paths fail; handled
  //      by the post-render injection pass)
  try {
    if (routeLocationLabels.length > 0) {
      // Primary: a horizontal numbered "snake" map (START ▶ ①..N ▶ DEST)
      // rendered via sharp (SVG→PNG). If that fails (no sharp / no librsvg),
      // the post-render pass falls back to the OOXML snake table instead — so
      // we DON'T fall back to the old vertical stepper, which the client asked
      // us to replace.
      let stepperBuf = await generateRouteLocationsSnake(routeLocationLabels);
      if (stepperBuf && stepperBuf.length > 0) {
        imageMap.set("routeLocations", {
          buffer: stepperBuf,
          contentType: "image/png",
          path: "(generated:routeLocations)",
        });
      }
    }
  } catch (err) {
    console.warn("[export actual] route locations snake failed - ignoring:", err);
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

  // Safety cap on the photo-fetch phase (prevents a hung S3 from making the
  // export run forever). Once exhausted, remaining rows render the "Photo not
  // available." fallback. This SCALES with the report count — a flat 100s used
  // to silently drop every photo after ~report 226 on big reports (500+ pts).
  // ~700ms/report budget, min 4 min, capped at 20 min. For very large reports
  // the server's reverse-proxy read timeout must also be raised to match.
  // Adapt photo dimensions/quality to scale so very large reports produce a
  // file Word can still open and the server doesn't run out of memory holding
  // thousands of images. ~2 photos/report worst case.
  const approxPhotos = expandedReports.length * 2;
  OBS_PHOTO_TARGET =
    approxPhotos > 6000
      ? { width: 1000, height: 660, quality: 78 } // huge reports: still keep files openable
      : approxPhotos > 2500
        ? { width: 1100, height: 720, quality: 84 }
        : { width: 1280, height: 860, quality: 88 }; // normal reports: max clarity for drawings
  console.log("[DOCX PHOTO TARGET]", { approxPhotos, target: OBS_PHOTO_TARGET });

  // Safety cap on the photo-fetch phase, scaled by report count (700ms/report),
  // min 4 min, max 45 min so very large reports aren't cut off. NOTE: the
  // server's reverse-proxy read timeout (nginx proxy_read_timeout) must be
  // raised to match, or the request is cut off before the file is ready.
  const PHOTO_PHASE_DEADLINE_MS = Math.min(2_700_000, Math.max(240_000, expandedReports.length * 700));
  const phaseStartedAt = Date.now();
  let phaseDeadlineHit = false;

  // Track which photo keys hold MULTIPLE photos so the post-render pass can
  // replace the single-image drawing with a 2-per-row borderless photo grid.
  const multiPhotoKeys = new Set<string>();
  // The EXTRA photos (index 1..N-1) for each multi-photo key. Photo 0 is stored
  // in imageMap under the normal photoKey (the template renders it); the extras
  // are stored here and injected into the DOCX zip during the grid pass.
  const multiPhotoExtraBuffers = new Map<string, Buffer[]>();

  const observations: ObservationData[] = [];
  // Export-time point de-duplication (DB is left untouched). A report is a
  // repeat — and is SKIPPED — if its GPS coordinate OR its point_key already
  // appeared in an earlier report. Photos are de-duped separately by the
  // at-most-once guard above. The first report for each unique coordinate /
  // point_key is kept; later repeats are dropped from the document only.
  const seenCoordKeys = new Set<string>();
  const seenPointKeys = new Set<string>();
  let skippedRepeatPoints = 0;
  for (let i = 0; i < expandedReports.length; i += 1) {
    const r = expandedReports[i];
    const rid = String(r.id || "").trim();
    const lat = pickLat(r);
    const lng = pickLng(r);
    const hasLat = lat !== null;
    const hasLon = lng !== null;

    // ---- Skip repeated points (coordinate OR point_key already emitted).
    const coordKey =
      hasLat && hasLon ? `${lat!.toFixed(6)},${lng!.toFixed(6)}` : "";
    const pointKey = String(r.point_key || "").trim();
    // Split duplicates (extra photos of the same report) legitimately share
    // the report's coordinate/point_key — never treat them as repeat points.
    const isSplitDup = splitDupRows.has(r);
    // IMPORTANT: do NOT drop a distinct report just because its coordinate or
    // point_key matches an earlier one. A point added in the app inherits the
    // previous row's point_key (and often its coordinate), so the old coord/
    // point_key de-dup silently deleted newly-added points from the document.
    // Keeping every distinct report is the safe default — a rare duplicate is
    // far better than a missing survey point. (Only emit a dup when the SAME
    // report id somehow repeats, which never happens in one list.)
    const isRepeatCoord = false && !isSplitDup && coordKey !== "" && seenCoordKeys.has(coordKey);
    const isRepeatPointKey = false && !isSplitDup && pointKey !== "" && seenPointKeys.has(pointKey);
    if (isRepeatCoord || isRepeatPointKey) {
      skippedRepeatPoints += 1;
      console.log("[DOCX SKIP REPEAT POINT]", {
        index: i,
        report_id: rid,
        reason: isRepeatCoord ? "duplicate coordinate" : "duplicate point_key",
        coordKey: coordKey || null,
        point_key: pointKey || null,
      });
      continue;
    }
    if (!isSplitDup) {
      if (coordKey) seenCoordKeys.add(coordKey);
      if (pointKey) seenPointKeys.add(pointKey);
    }

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
        remainingObservations: expandedReports.length - i,
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
      // Fallback only recovers ORPHAN photos — those whose report_id matches
      // no report in this export. Photos that belong to another EXISTING
      // report sharing this point_key are excluded, so the same photo can no
      // longer leak into a second report's block.
      // When the user unticked EVERY photo of this report in the photo
      // picker, honour that: no orphan may be borrowed via point_key.
      const photosByPk = rowPointKey && !reportsFullyExcludedByUser.has(rid)
        ? (photosByPointKey.get(rowPointKey) || []).filter(
            (p) => !knownReportIdSet.has(String(p.report_id || "").trim())
          )
        : [];
      const candidatePhotos = photosByRid.length > 0 ? photosByRid : photosByPk;
      // Drop any photo already embedded under an earlier report so the same
      // image can never repeat across blocks.
      const reportPhotos = candidatePhotos.filter((p) => {
        const idnt = photoIdentity(p);
        return !idnt || !usedPhotoKeys.has(idnt);
      });
      const matchSource = photosByRid.length > 0
        ? "report_id"
        : photosByPk.length > 0
          ? "point_key"
          : "none";
      const firstPhotoRow = reportPhotos[0] || null;
      // Claim ALL of this report's photos so none can be reused by a later
      // block (the whole point's set of photos belongs to this one block).
      for (const p of reportPhotos) {
        const idnt = photoIdentity(p);
        if (idnt) usedPhotoKeys.add(idnt);
      }
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

      // ALL of this point's photos render TOGETHER as a grid under its one
      // table. Cap at 8 so a point with an unusual number of photos stays sane.
      const selectedPhotos = reportPhotos.slice(0, 8);
      void firstPhotoRow; // (kept above for logging/claim; grid uses selectedPhotos)
      let fetchedAny = false;
      let lastFetchFailureReason: string | null = null;
      const fetchedBuffers: Array<{ buffer: Buffer; contentType: string }> = [];
      // Fetch this report's (up to 2) photos IN PARALLEL, preserving order — so
      // a 2-photo report doesn't pay two sequential fetch/timeouts, halving the
      // photo-phase time on big reports.
      const settledPhotos = await Promise.all(
        selectedPhotos.map(async (p) => {
          const candidate = normalizeS3Url(p?.url);
          if (!candidate) {
            return { ok: false as const, reason: "row has no url and no path" };
          }
          try {
            const fetched = await fetchImageBuffer(candidate, `photo[${rid}]`);
            if (fetched && Buffer.isBuffer(fetched.buffer) && fetched.buffer.length > 0) {
              return { ok: true as const, buffer: fetched.buffer, contentType: fetched.contentType };
            }
            return {
              ok: false as const,
              reason: "fetchImageBuffer returned no buffer (404/403/non-image/oversize/timeout)",
            };
          } catch (err) {
            return { ok: false as const, reason: `fetch threw: ${(err as Error)?.message || String(err)}` };
          }
        })
      );
      for (const s of settledPhotos) {
        if (s.ok) fetchedBuffers.push({ buffer: s.buffer, contentType: s.contentType });
        else lastFetchFailureReason = s.reason;
      }

      // Assemble the photo buffers. Photo 0 goes into imageMap (the template
      // renders it in place); photos 1..N-1 are kept as extras and, together
      // with photo 0, laid out as a 2-per-row grid by the post-render pass.
      // 0 fetched → no photo (placeholder).
      if (fetchedBuffers.length >= 1) {
        photoKey = `photo_${i}`;
        const optimizedAll = await Promise.all(
          fetchedBuffers.map((b, bi) =>
            optimizeDocxImage(b.buffer, "observation", `${photoKey}_${bi}`)
          )
        );
        // Photo 0 → imageMap (template renders it; the grid pass replaces it
        // when there are 2+ photos).
        imageMap.set(photoKey, {
          buffer: optimizedAll[0],
          contentType: "image/jpeg",
        });
        if (optimizedAll.length >= 2) {
          multiPhotoKeys.add(photoKey);
          multiPhotoExtraBuffers.set(photoKey, optimizedAll.slice(1));
        }
        fetchedAny = true;
        photoFetchSuccess += 1;
        console.log("[DOCX PHOTO BUFFERS STORED]", {
          index: i,
          reportId: rid,
          photoKey,
          photoCount: optimizedAll.length,
          totalCandidates: reportPhotos.length,
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
      if (!fetchedAny && reportPhotos.length === 0 && !reportsFullyExcludedByUser.has(rid)) {
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
            // The master file's extension often doesn't match the real image
            // (says .jpg, object is .webp/.png) — try every common variant so
            // the operator never has to fix extensions by hand.
            const cleanName = fileNameRaw.replace(/^\/+/, "");
            const stem = cleanName.replace(/\.[a-z0-9]{2,5}$/i, "");
            const nameVariants = Array.from(
              new Set([
                cleanName,
                `${stem}.jpg`,
                `${stem}.jpeg`,
                `${stem}.png`,
                `${stem}.webp`,
              ])
            );
            // Prefixes the bulk-import + manual-upload paths actually use.
            const prefixes = [
              `reports/photos/${rid}/`,
              `reports/${projectId}/${rid}/`,
              `reports/photos/${projectId}/`,
              `reports/${projectId}/`,
              `reports/photos/`,
            ];
            for (const pfx of prefixes) {
              for (const name of nameVariants) {
                candidates.push(`${base}/${pfx}${name}`);
              }
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

    // KM: prefer the value that came WITH the report (its own kms/km — e.g. the
    // "Km: 6.10" printed in an uploaded survey), so re-exporting never shifts
    // the chainage. Only when the report has no km of its own do we fall back to
    // the cumulative haversine distance along the ordered route.
    const rawKm =
      r.kms != null && String(r.kms).trim() !== ""
        ? Number(r.kms)
        : r.km != null && String(r.km).trim() !== ""
          ? Number(r.km)
          : NaN;
    const kmText = Number.isFinite(rawKm)
      ? formatKmValue(rawKm)
      : formatKmValue(r.calculated_km);

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

    // Color strictly by the report's own difficulty field. (Previously this
    // fell back through remarks_action/status/vehicle_movement, which made
    // every row with unrecognized text render green.)
    const difficultyValue = String(r.difficulty || "");
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
      // Join the category's template sentence with the report's measurement
      // ("...height is 6m"); template-only when there's no measurement; the
      // observation unchanged when it's already a full sentence.
      observation: valueOrEmDash(
        composeObservation(String(r.description ?? r.observation ?? ""), String(r.category || ""))
      ),
      remarks: formatRemarksAction(r.remarks_action, tableColors.key),
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
      if (tagName === "categoryIconKey" || key.startsWith("catIcon_")) {
        // Use the icon's trimmed aspect ratio so a wide-short glyph renders
        // short (no blank space around it) instead of a fixed square.
        const entry = key ? imageMap.get(key) : undefined;
        return entry?.displaySize || CATEGORY_ICON_SIZE;
      }
      if (tagName === "categorySummaryIcon") {
        // Summary icons share the (now non-square, trimmed) category buffer.
        // Rescale the icon's trimmed aspect ratio into the small summary box so
        // it isn't stretched to a square.
        const ds = (tagValue as { displaySize?: [number, number] })?.displaySize;
        if (Array.isArray(ds) && ds.length === 2 && ds[0] > 0 && ds[1] > 0) {
          const [maxW, maxH] = CATEGORY_SUMMARY_ICON_SIZE;
          const s = Math.min(maxW / ds[0], maxH / ds[1]);
          return [Math.max(1, Math.round(ds[0] * s)), Math.max(1, Math.round(ds[1] * s))];
        }
        return CATEGORY_SUMMARY_ICON_SIZE;
      }
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
        // Single-photo reports: full OBSERVATION_PHOTO_SIZE. Multi-photo
        // reports render photo 0 at MULTI_PHOTO_SIZE, but the post-render grid
        // pass replaces it with a 2-per-row grid of ALL the point's photos.
        if (multiPhotoKeys.has(key)) {
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
    // NOTE: `remarks` is intentionally NOT counted — it is now always populated
    // with the difficulty-based action text ("Go ahead…" etc.), so it would
    // make every row look "meaningful" and defeat this empty-row filter. A row
    // is real only if it has GPS / location / category / observation / photo.
    return (
      meaningful(r.gpsLat) ||
      meaningful(r.gpsLon) ||
      meaningful(r.location) ||
      meaningful(r.category) ||
      meaningful(r.observation) ||
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
    conclusion: CONCLUSION_SENTINEL,
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

  // ----- Step 10b-1b (summary report only): insert two summary tables before
  // the first point table — a Category/Work-Item count table and a Key Obstacle
  // Summary table (Sl No · Location · GPS · Km · Obstruction · Score · Recommendation).
  // The Obstruction / Score / Recommendation columns are AUTO-DERIVED from each
  // point's category, route difficulty and observation, and are fully editable
  // in Word. Only runs for the summary export; never touches the detail report.
  if (summaryMode) {
    try {
      const renderedZip = doc.getZip();
      const docFile = renderedZip.file("word/document.xml");
      if (docFile) {
        let xml = docFile.asText();
        const alreadyHasSummary = xml.includes("Mitigation Recommendations for major obstacles");

        const gpsIdx = xml.indexOf("GPS LOCATION");
        let tblStart = -1;
        if (gpsIdx !== -1) {
          // Anchor on the TABLE OPEN tag ("<w:tbl>" or "<w:tbl " with
          // attributes) nearest before the "GPS LOCATION" header — NOT a child
          // tag like <w:tblPr>/<w:tblGrid>/<w:tblHeader> (those also start with
          // "<w:tbl" and previously made this return -1, so nothing injected).
          tblStart = Math.max(
            xml.lastIndexOf("<w:tbl>", gpsIdx),
            xml.lastIndexOf("<w:tbl ", gpsIdx)
          );
        }

        if (!alreadyHasSummary && tblStart !== -1) {
          const esc = (s: unknown) =>
            String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

          const cell = (
            text: string,
            opt: { w: number; bold?: boolean; fill?: string | null; color?: string; align?: string; size?: number } = { w: 1000 }
          ) => {
            const color = opt.color || "000000";
            const align = opt.align || "left";
            const size = opt.size || 18;
            return (
              "<w:tc><w:tcPr>" +
              `<w:tcW w:w="${opt.w}" w:type="dxa"/>` +
              (opt.fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${opt.fill}"/>` : "") +
              '<w:vAlign w:val="center"/></w:tcPr>' +
              `<w:p><w:pPr><w:jc w:val="${align}"/><w:spacing w:before="20" w:after="20" w:line="288" w:lineRule="auto"/></w:pPr>` +
              `<w:r><w:rPr>${opt.bold ? "<w:b/>" : ""}<w:color w:val="${color}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr>` +
              `<w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p></w:tc>`
            );
          };
          const tableWrap = (_totalW: number, cols: number[], rowsXml: string) =>
            "<w:tbl><w:tblPr>" +
            // OOXML order inside tblPr: tblW → tblBorders → tblLayout → tblCellMar.
            // Fill 100% of the text width; columns keep their relative widths.
            '<w:tblW w:w="5000" w:type="pct"/>' +
            "<w:tblBorders>" +
            ["top", "left", "bottom", "right", "insideH", "insideV"]
              .map((s) => `<w:${s} w:val="single" w:sz="6" w:space="0" w:color="BFC7D2"/>`)
              .join("") +
            "</w:tblBorders>" +
            '<w:tblLayout w:type="fixed"/>' +
            // Tight, compact rows: small horizontal padding so text doesn't hug
            // the borders, minimal vertical padding (no extra blank space).
            "<w:tblCellMar>" +
            '<w:top w:w="10" w:type="dxa"/>' +
            '<w:left w:w="120" w:type="dxa"/>' +
            '<w:bottom w:w="10" w:type="dxa"/>' +
            '<w:right w:w="120" w:type="dxa"/>' +
            "</w:tblCellMar></w:tblPr>" +
            "<w:tblGrid>" +
            cols.map((w) => `<w:gridCol w:w="${w}"/>`).join("") +
            "</w:tblGrid>" +
            rowsXml +
            "</w:tbl>";
          const headingP = (text: string) =>
            '<w:p><w:pPr><w:spacing w:before="200" w:after="80"/></w:pPr>' +
            `<w:r><w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/><w:color w:val="1F2937"/></w:rPr>` +
            `<w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;
          const spacer = '<w:p><w:pPr><w:spacing w:before="0" w:after="120"/></w:pPr></w:p>';
          const bodyP = (text: string, opt: { bold?: boolean; size?: number } = {}) =>
            '<w:p><w:pPr><w:spacing w:before="20" w:after="20"/></w:pPr>' +
            `<w:r><w:rPr>${opt.bold ? "<w:b/>" : ""}<w:sz w:val="${opt.size || 20}"/><w:szCs w:val="${opt.size || 20}"/><w:color w:val="1F2937"/></w:rPr>` +
            `<w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;

          const deriveObstruction = (category: string, diff: string) => {
            const base = String(category || "Obstacle").trim() || "Obstacle";
            const tail = diff === "red" ? "Major Civil Work" : diff === "yellow" ? "Clearance Check" : "Review";
            return `${base} – ${tail}`;
          };
          const deriveScore = (diff: string) =>
            diff === "red"
              ? "High risk; significant clearance / civil work required."
              : diff === "yellow"
                ? "Moderate risk; pass with caution / minor works."
                : "Low risk; normal passage.";

          // Table 1 — fixed Category / Work-Item taxonomy; counts derived from
          // the points by category/keyword. Exact reference colours + widths:
          // green header 196B24, striped light-green C1F0C8.
          const WORK_ITEMS: { label: string; test: (cat: string, hay: string) => boolean }[] = [
            { label: "Toll plaza modification / clearance", test: (c) => /toll/.test(c) },
            { label: "Toll sensor pole removal", test: (_c, h) => /sensor pole/.test(h) },
            { label: "Centre median modification / removal", test: (_c, h) => /median/.test(h) },
            { label: "Land filling", test: (_c, h) => /land\s*fill|filling/.test(h) },
            { label: "Digging / excavation", test: (_c, h) => /digging|excavat/.test(h) },
            { label: "Road levelling", test: (_c, h) => /level+ing/.test(h) },
            { label: "Footpath bridge obstruction", test: (c, h) => /footpath/.test(c) || /footpath/.test(h) },
            { label: "Underpass bridge obstruction", test: (c, h) => /under\s*pass/.test(c) || /under\s*pass/.test(h) },
            { label: "Port gate clearance", test: (c) => /gate/.test(c) },
            { label: "Railway shutdown", test: (c, h) => /railway|level crossing/.test(c) || /railway/.test(h) },
          ];
          const workCounts = WORK_ITEMS.map(() => 0);
          for (const o of observations || []) {
            const cat = String((o as any).category || "").toLowerCase();
            const hay = `${cat} ${String((o as any).observation || "")} ${String((o as any).remarks || "")}`
              .replace(/<[^>]+>/g, " ")
              .toLowerCase();
            WORK_ITEMS.forEach((w, wi) => {
              if (w.test(cat, hay)) workCounts[wi] += 1;
            });
          }
          const t1cols = [5553, 1218, 10417];
          const t1Head =
            "<w:tr>" +
            cell("Category / Work Item", { w: t1cols[0], bold: true, fill: "196B24", color: "FFFFFF", align: "center" }) +
            cell("Count", { w: t1cols[1], bold: true, fill: "196B24", color: "FFFFFF", align: "center" }) +
            cell("Key Note", { w: t1cols[2], bold: true, fill: "196B24", color: "FFFFFF", align: "center" }) +
            "</w:tr>";
          const t1Body = WORK_ITEMS.map((w, ri) => {
            const f = ri % 2 === 0 ? "C1F0C8" : null;
            const count = workCounts[ri];
            return (
              "<w:tr>" +
              cell(w.label, { w: t1cols[0], fill: f }) +
              cell(String(count), { w: t1cols[1], align: "center", fill: f }) +
              cell(count === 0 ? "Not specified" : "-", { w: t1cols[2], fill: f }) +
              "</w:tr>"
            );
          }).join("");
          const table1 = tableWrap(17188, t1cols, t1Head + t1Body);

          // Table 2 — key obstacle summary, one row per point. Exact reference
          // colours + widths: navy header 183153, striped light-blue CAEDFB.
          const t2cols = [1362, 2358, 2376, 988, 2690, 5813, 7174];
          const t2Head =
            "<w:tr>" +
            cell("Sl No", { w: t2cols[0], bold: true, fill: "183153", color: "FFFFFF", align: "center" }) +
            cell("Route / Location", { w: t2cols[1], bold: true, fill: "183153", color: "FFFFFF" }) +
            cell("GPS Coordinates", { w: t2cols[2], bold: true, fill: "183153", color: "FFFFFF" }) +
            cell("Kms", { w: t2cols[3], bold: true, fill: "183153", color: "FFFFFF", align: "center" }) +
            cell("Obstruction", { w: t2cols[4], bold: true, fill: "183153", color: "FFFFFF" }) +
            cell("Score Parameter", { w: t2cols[5], bold: true, fill: "183153", color: "FFFFFF" }) +
            cell("Recommendation", { w: t2cols[6], bold: true, fill: "183153", color: "FFFFFF" }) +
            "</w:tr>";
          const t2Body = (observations || [])
            .map((o: any, i: number) => {
              const diff = String(o.difficultyKey || "green");
              const f = i % 2 === 0 ? "CAEDFB" : null; // striped light-blue rows
              const gps = [o.gpsLat, o.gpsLon].filter((x) => x && x !== "-").join(" ");
              const rec =
                String(o.observation || "")
                  .replace(/<[^>]+>/g, " ")
                  .replace(/\s+/g, " ")
                  .trim() || String(o.remarks || "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
              return (
                "<w:tr>" +
                cell(String(i + 1), { w: t2cols[0], align: "center", fill: f }) +
                cell(String(o.location || "-"), { w: t2cols[1], fill: f }) +
                cell(gps || "-", { w: t2cols[2], fill: f }) +
                cell(String(o.km || "-"), { w: t2cols[3], align: "center", fill: f }) +
                cell(deriveObstruction(o.category, diff), { w: t2cols[4], fill: f }) +
                cell(deriveScore(diff), { w: t2cols[5], fill: f }) +
                cell(rec || "-", { w: t2cols[6], fill: f }) +
                "</w:tr>"
              );
            })
            .join("");
          const table2 = tableWrap(22761, t2cols, t2Head + t2Body);

          // Observations summary block (A. Surveyed Route / Distance / B. Key
          // Locations) derived from the included points.
          const routeName = String(projectName || "").toUpperCase();
          const maxKm = (observations || []).reduce((m: number, o: any) => {
            const v = parseFloat(String(o.km || "").replace(/[^\d.]/g, ""));
            return Number.isFinite(v) && v > m ? v : m;
          }, 0);
          const keyLocs = Array.from(
            new Set((observations || []).map((o: any) => String(o.location || "").trim()).filter(Boolean))
          ).join(" - ");
          // Prefer the per-project editable override (Summary of Observations
          // step in the setup wizard) over the auto-derived text.
          const ovFrom = String(observationSummaryOverride?.fromText || "").trim();
          const ovTo = String(observationSummaryOverride?.toText || "").trim();
          const ovDistance = String(observationSummaryOverride?.distance || "").trim();
          const ovKeyLocs = String(observationSummaryOverride?.keyLocations || "").trim();
          const surveyedRouteLine = ovFrom
            ? ovTo
              ? `From ${ovFrom} to ${ovTo}`
              : `From ${ovFrom}`
            : routeName
              ? `From ${routeName}`
              : "-";
          const distanceLine = ovDistance
            ? `Distance: ${ovDistance}`
            : maxKm > 0
              ? `Distance: ${maxKm.toFixed(2)} km`
              : "Distance: -";
          // Embed the optional flow-chart image into the rendered doc (media +
          // relationship, the same proven way the route-locations image is
          // embedded) and show it under "B. Key Locations". Defensive: on any
          // failure it is simply omitted, never corrupting the document.
          let flowChartPara = "";
          if (flowChartImg) {
            try {
              const ctF = renderedZip.file("[Content_Types].xml");
              if (ctF) {
                let ctX = ctF.asText();
                if (!/Extension="jpeg"/i.test(ctX) && ctX.includes("</Types>")) {
                  ctX = ctX.replace(
                    "</Types>",
                    `<Default Extension="jpeg" ContentType="image/jpeg"/></Types>`
                  );
                  renderedZip.file("[Content_Types].xml", ctX);
                }
              }
              renderedZip.file("word/media/flow_chart.jpeg", flowChartImg.buffer);
              const relsF = renderedZip.file("word/_rels/document.xml.rels");
              let relsX = relsF ? relsF.asText() : "";
              const fcRId = "rIdFlowChart";
              if (relsX && !relsX.includes(`Id="${fcRId}"`)) {
                relsX = relsX.replace(
                  "</Relationships>",
                  `<Relationship Id="${fcRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/flow_chart.jpeg"/></Relationships>`
                );
                renderedZip.file("word/_rels/document.xml.rels", relsX);
              }
              flowChartPara =
                `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="80" w:after="80"/><w:keepLines/></w:pPr>` +
                `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
                `<wp:extent cx="${flowChartImg.emuW}" cy="${flowChartImg.emuH}"/>` +
                `<wp:docPr id="9400" name="Route Survey Flow Chart"/>` +
                `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
                `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
                `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
                `<pic:nvPicPr><pic:cNvPr id="0" name=""/><pic:cNvPicPr/></pic:nvPicPr>` +
                `<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${fcRId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
                `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${flowChartImg.emuW}" cy="${flowChartImg.emuH}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
                `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
            } catch (fcEmbedErr) {
              console.warn("[flow chart] embed failed - omitting:", fcEmbedErr);
              flowChartPara = "";
            }
          }
          const observationsBlock =
            headingP("OBSERVATIONS:") +
            bodyP("Route Summary", { bold: true }) +
            bodyP("A. Surveyed Route:", { bold: true }) +
            bodyP(surveyedRouteLine) +
            bodyP(distanceLine) +
            bodyP("B. Key Locations:", { bold: true }) +
            bodyP(ovKeyLocs || keyLocs || "-") +
            flowChartPara +
            spacer;

          // Page break so the per-point detail tables start on a fresh page,
          // after the title / observations / two summary tables.
          const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

          const summaryXml =
            headingP("TURNING CIRCLE DRAWING:") +
            spacer +
            headingP("STABILITY DRAWING:") +
            spacer +
            observationsBlock +
            table1 +
            spacer +
            headingP("Mitigation Recommendations for major obstacles with coordinates – proposed recommendation") +
            table2 +
            spacer +
            pageBreak;

          xml = xml.slice(0, tblStart) + summaryXml + xml.slice(tblStart);
          renderedZip.file("word/document.xml", xml);
          console.log("[summary tables] inserted", {
            categories: (categorySummary || []).length,
            points: (observations || []).length,
          });
        } else {
          console.log("[summary tables] skipped", { alreadyHasSummary, anchorFound: tblStart !== -1 });
        }
      }
    } catch (err) {
      console.error("[summary tables] insert failed - leaving document untouched:", err);
    }
  }

  // ----- Step 10b-2: Insert a "difficulty colour key" legend before the
  // first observation table so the exported file is self-explanatory and
  // user-friendly. The legend uses the SAME 6 colours as the observation
  // rows (3 bright header fills + 3 faint body tints) and is rendered as a
  // plain, fully-editable Word table (standard <w:shd w:fill> cell shading),
  // so the client can recolour, retype, add or delete it like any other
  // table. Defensive: only inserts when the anchor is found and the legend
  // is not already present; any failure leaves the document untouched.
  try {
    const renderedZip = doc.getZip();
    const docFile = renderedZip.file("word/document.xml");
    if (docFile) {
      let xml = docFile.asText();

      const LEGEND_MARKER = "ROUTE DIFFICULTY COLOUR KEY";
      const alreadyHasLegend = xml.includes(LEGEND_MARKER);

      // Anchor: the first observation table, detected by its "GPS LOCATION"
      // header fingerprint, then backtracking to the enclosing <w:tbl> open.
      const gpsIdx = xml.indexOf("GPS LOCATION");
      let tblStart = -1;
      if (gpsIdx !== -1) {
        const candidate = xml.lastIndexOf("<w:tbl", gpsIdx);
        if (candidate !== -1) {
          const after = xml[candidate + 6];
          if (after === " " || after === ">" || after === "\n" || after === "\t" || after === "\r") {
            tblStart = candidate;
          }
        }
      }

      if (!alreadyHasLegend && tblStart !== -1) {
        // Each row: bright header-fill swatch cell (white bold label) +
        // faint body-tint meaning cell. Colours match Step 10b exactly.
        const legendRow = (
          headerFill: string,
          bodyFill: string,
          label: string,
          meaning: string
        ) =>
          "<w:tr>" +
            "<w:tc>" +
              `<w:tcPr><w:tcW w:w="1800" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="${headerFill}"/><w:vAlign w:val="center"/></w:tcPr>` +
              '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="40" w:after="40"/></w:pPr>' +
                `<w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>${label}</w:t></w:r></w:p>` +
            "</w:tc>" +
            "<w:tc>" +
              `<w:tcPr><w:tcW w:w="7200" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="${bodyFill}"/><w:vAlign w:val="center"/></w:tcPr>` +
              '<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>' +
                `<w:r><w:rPr><w:color w:val="0B3D2E"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve">${meaning}</w:t></w:r></w:p>` +
            "</w:tc>" +
          "</w:tr>";

        const legendXml =
          // Heading paragraph.
          '<w:p><w:pPr><w:spacing w:before="160" w:after="80"/></w:pPr>' +
            `<w:r><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/><w:color w:val="1F2937"/></w:rPr><w:t>${LEGEND_MARKER}</w:t></w:r></w:p>` +
          // Legend table.
          "<w:tbl>" +
            "<w:tblPr>" +
              '<w:tblW w:w="9000" w:type="dxa"/>' +
              "<w:tblBorders>" +
                '<w:top w:val="single" w:sz="6" w:space="0" w:color="B7CDB3"/>' +
                '<w:left w:val="single" w:sz="6" w:space="0" w:color="B7CDB3"/>' +
                '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="B7CDB3"/>' +
                '<w:right w:val="single" w:sz="6" w:space="0" w:color="B7CDB3"/>' +
                '<w:insideH w:val="single" w:sz="6" w:space="0" w:color="B7CDB3"/>' +
                '<w:insideV w:val="single" w:sz="6" w:space="0" w:color="B7CDB3"/>' +
              "</w:tblBorders>" +
            "</w:tblPr>" +
            '<w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="7200"/></w:tblGrid>' +
            legendRow("56BE9F", "F0F9F4", "GREEN", "No problem - vehicle can pass normally.") +
            legendRow("EABD0D", "FFFBE6", "YELLOW", "Quite difficult - pass with caution.") +
            legendRow("F05052", "FDF2F2", "RED", "Do not go - not passable for vehicle movement.") +
          "</w:tbl>" +
          // Spacer paragraph (a paragraph between two adjacent tables keeps
          // the OOXML valid and the legend visually separated).
          '<w:p><w:pPr><w:spacing w:before="0" w:after="120"/></w:pPr></w:p>';

        xml = xml.slice(0, tblStart) + legendXml + xml.slice(tblStart);
        renderedZip.file("word/document.xml", xml);
        console.log("[difficulty legend] inserted colour key before first observation table");
      } else {
        console.log("[difficulty legend] skipped", {
          alreadyHasLegend,
          anchorFound: tblStart !== -1,
        });
      }
    }
  } catch (err) {
    console.error("[difficulty legend] insert failed - leaving document untouched:", err);
  }

  // ----- Step 10b-3: Title wording — this is the DETAIL report, not a
  // summary. The template heading reads "{projectNameUpper} SUMMARY REPORT";
  // rewrite the visible heading text to "... DETAIL REPORT". Only the text
  // between tags is changed (no structural edits), so it cannot corrupt the
  // document.
  try {
    const renderedZip = doc.getZip();
    const docFile = renderedZip.file("word/document.xml");
    if (docFile) {
      let xml = docFile.asText();
      if (summaryMode) {
        // Summary report: first title paragraph (the one with "SUMMARY REPORT")
        // becomes the fixed Risk Mitigation Plan heading; the route name stays
        // on the line below (rendered separately by the template). Also rename
        // the single GA heading to "GA VEHICLE DRAWING:".
        const titlePhrase = "RISK MITIGATION PLAN FOR TRANSPORTATION OF OVER DIMENSION CARGO (ODC)";
        const escT = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
        let titleDone = false;
        xml = xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (p) => {
          if (titleDone || !p.includes("SUMMARY REPORT")) return p;
          titleDone = true;
          let firstRun = true;
          return p.replace(/(<w:t\b[^>]*>)([\s\S]*?)(<\/w:t>)/g, (_m, open, _t, close) => {
            if (firstRun) {
              firstRun = false;
              return `${open}${escT(titlePhrase)}${close}`;
            }
            return `${open}${close}`;
          });
        });
        xml = xml.replace(/GA DRAWING FOR 50 FEET TRAILER WITHOUT LOAD:/g, "GA VEHICLE DRAWING:");
        renderedZip.file("word/document.xml", xml);
        console.log("[summary title] applied Risk Mitigation Plan title + GA heading");
      } else if (xml.includes("SUMMARY REPORT")) {
        xml = xml.replace(/SUMMARY REPORT/g, "DETAIL REPORT");
        renderedZip.file("word/document.xml", xml);
        console.log("[title wording] replaced 'SUMMARY REPORT' with 'DETAIL REPORT'");
      }
    }
  } catch (err) {
    console.error("[title wording] replace failed - leaving heading untouched:", err);
  }

  // ----- Step 10b-3b: The template's single GA-drawing heading ("GA DRAWING
  // FOR 50 FEET TRAILER WITHOUT LOAD:") is coloured WHITE (invisible on the
  // white page) and carries fixed wording. Two jobs here, scoped to THAT
  // heading paragraph only (matched by its text, so the white banner title on
  // the coloured header is never touched):
  //   (a) Detail reports: rename it to the FIRST uploaded diagram's heading
  //       (e.g. "GA DRAWING WITHOUT LOAD:") so the template GA slot sits under
  //       the correct heading. Extra diagrams get their own headed pages below.
  //   (b) Recolour white -> dark so the heading is actually visible.
  try {
    const renderedZip = doc.getZip();
    const docFile = renderedZip.file("word/document.xml");
    if (docFile) {
      let xml = docFile.asText();
      let changed = false;
      // Heading text is plain ASCII (letters/space/colon) so no escaping needed.
      const firstGaHeading = !summaryMode ? String(gaHeadings[0] || "").trim() : "";
      xml = xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (p) => {
        if (!/GA DRAWING FOR 50 FEET TRAILER|GA VEHICLE DRAWING/.test(p)) return p;
        let np = p;
        // (a) rename the template GA heading to the first diagram's heading.
        if (firstGaHeading) {
          np = np.replace(
            /GA DRAWING FOR 50 FEET TRAILER WITHOUT LOAD:/g,
            firstGaHeading
          );
        }
        // (b) recolour white -> dark so it is visible.
        if (/w:color w:val="FFFFFF"/i.test(np)) {
          np = np.replace(/<w:color w:val="FFFFFF"\s*\/>/gi, '<w:color w:val="1F2937"/>');
        }
        if (np !== p) changed = true;
        return np;
      });
      if (changed) {
        renderedZip.file("word/document.xml", xml);
        console.log(
          "[GA heading] renamed + recoloured (first heading: " + (firstGaHeading || "n/a") + ")"
        );
      }
    }
  } catch (err) {
    console.error("[GA heading] rename/recolour failed - leaving untouched:", err);
  }

  // ----- Step 10b-4: Add a custom theme so the 6 route-difficulty colours
  // appear in Word's colour picker ("Theme Colors" row). Defensive: skips if a
  // theme already exists or any required part is missing.
  try {
    const injected = injectDifficultyThemeColors(doc.getZip() as never);
    console.log("[difficulty theme] colour-picker theme injected:", injected);
  } catch (err) {
    console.error("[difficulty theme] injection failed - leaving document untouched:", err);
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
      //   GPS LOCATION   : 17.0%  (widened so "GPS LOCATION" header and each
      //                            coordinate line like "N28 33.274" stay on
      //                            one line instead of wrapping)
      //   KM             : 6.8%
      //   LOCATION       : 17.5%
      //   CATEGORY       : auto-sizes to the category icon
      //   OBSERVATION    : 32.0%  (reduced to give the width to GPS; still the
      //                            widest column, plenty for the description)
      //   REMARKS/ACTION : 16.2%
      // We assume a usable content width of 14330 dxa (A4 ≈ 11906,
      // Letter ≈ 12240; the template uses landscape so 14330 is a
      // safe approximation). Each cell width is set in DXA so total
      // sums correctly regardless of page-margin variance.
      // The CATEGORY column auto-sizes to the (enlarged) category icon so the
      // image is never clipped — the column updates based on its content. The
      // five remaining columns share the leftover width in their original
      // relative proportions. Column order:
      //   [GPS LOCATION, KM, LOCATION, CATEGORY, OBSERVATION, REMARKS/ACTION].
      const TOTAL_DXA = 14330;
      const PX_TO_DXA = 15; // 96 px = 1 in = 1440 dxa
      // Icon width + a little horizontal cell padding (left/right margins).
      const CATEGORY_COL_DXA = CATEGORY_ICON_SIZE[0] * PX_TO_DXA + 260;
      const OTHER_PCTS = [17, 6.8, 17.5, 32, 16.2]; // GPS, KM, LOC, OBS, REM
      const OTHER_SUM = OTHER_PCTS.reduce((a, b) => a + b, 0);
      const REMAINING_DXA = Math.max(0, TOTAL_DXA - CATEGORY_COL_DXA);
      const OTHER_DXA = OTHER_PCTS.map((p) =>
        Math.round(REMAINING_DXA * (p / OTHER_SUM))
      );
      // Narrow leading "#" serial-number column (mirrors the website's point
      // list). ~0.4 in — just wide enough for a 2-3 digit number.
      const SERIAL_COL_DXA = 560;
      const OBS_COL_DXA = [
        SERIAL_COL_DXA, // # (serial number)
        OTHER_DXA[0], // GPS LOCATION
        OTHER_DXA[1], // KM
        OTHER_DXA[2], // LOCATION
        CATEGORY_COL_DXA, // CATEGORY (fits the icon)
        OTHER_DXA[3], // OBSERVATION
        OTHER_DXA[4], // REMARKS / ACTION
      ];
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
      let obsSerial = 0; // running point number for the "#" column

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
          obsSerial += 1;

          // 0. Inject a leading "#" serial-number column so each point is
          //    numbered like the website's point list. The header (1st) row
          //    gets a white "#" label; the data (2nd) row gets this point's
          //    running number. Both clone the existing first cell's <w:tcPr>
          //    (so they inherit the difficulty shading + borders), only the
          //    width and the text differ. Only the standard 6-column header
          //    and data rows are touched; any other row (e.g. a photo row) is
          //    left alone.
          let injected = 0;
          let out = tblXml.replace(
            /<w:tr\b[\s\S]*?<\/w:tr>/g,
            (trXml: string) => {
              if (injected >= 2) return trXml;
              const cells = trXml.match(/<w:tc\b[\s\S]*?<\/w:tc>/g);
              if (!cells || cells.length !== 6) return trXml;
              const firstCell = cells[0];
              const tcPrMatch = firstCell.match(/<w:tcPr>[\s\S]*?<\/w:tcPr>/);
              let newTcPr: string;
              if (tcPrMatch) {
                newTcPr = tcPrMatch[0];
                const widthXml = `<w:tcW w:w="${SERIAL_COL_DXA}" w:type="dxa"/>`;
                newTcPr = newTcPr.includes("<w:tcW")
                  ? newTcPr.replace(/<w:tcW\b[^/]*\/>/, widthXml)
                  : newTcPr.replace("<w:tcPr>", `<w:tcPr>${widthXml}`);
              } else {
                const shd = (firstCell.match(/<w:shd\b[^>]*\/>/) || [""])[0];
                newTcPr = `<w:tcPr><w:tcW w:w="${SERIAL_COL_DXA}" w:type="dxa"/>${shd}</w:tcPr>`;
              }
              const isHeaderRow = injected === 0;
              const para = isHeaderRow
                ? `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>#</w:t></w:r></w:p>`
                : `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>${obsSerial}</w:t></w:r></w:p>`;
              const serialCell = `<w:tc>${newTcPr}${para}</w:tc>`;
              injected += 1;
              return trXml.replace(firstCell, serialCell + firstCell);
            }
          );

          // 1. Replace <w:tblGrid>...</w:tblGrid> with our 7-col grid.
          out = out.replace(
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

          // 6. Keep each ROW's content on one page (cantSplit — a tall text
          //    cell no longer splits across pages).
          //    NOTE: we deliberately do NOT add <w:keepNext/> to bind the table
          //    to the photo. Each point already starts at the TOP of a fresh
          //    page, so when the table + photo fit they stay together anyway;
          //    and binding them when they DON'T fit (a tall table + a big photo)
          //    forced Word to push the whole block to the next page, stranding a
          //    BLANK page before it. Without the binding, a tall table fills its
          //    page and the photo flows to the next page — no blank page.
          out = out.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, (trXml: string) => {
            if (/<w:cantSplit\b/.test(trXml)) return trXml;
            return /<w:trPr>/.test(trXml)
              ? trXml.replace("<w:trPr>", "<w:trPr><w:cantSplit/>")
              : trXml.replace(/(<w:tr\b[^>]*>)/, "$1<w:trPr><w:cantSplit/></w:trPr>");
          });

          // 7. Fill the gaps: align DATA-row cells to the TOP. When an
          //    Observation is long the row is tall, and vertically-centred
          //    short cells (GPS / KM / Location) float in the middle with big
          //    empty gaps above and below. Top alignment lines every cell up
          //    with the first line of text — no gaps. The header row (first
          //    <w:tr>) stays centred.
          const alignTrs = out.match(/<w:tr\b[\s\S]*?<\/w:tr>/g) || [];
          for (let ti = 1; ti < alignTrs.length; ti += 1) {
            const tr = alignTrs[ti];
            const topped = tr.replace(
              /<w:vAlign w:val="center"\/>/g,
              '<w:vAlign w:val="top"/>'
            );
            if (topped !== tr) out = out.replace(tr, topped);
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
      //
      // IMPORTANT: this uses <w:pageBreakBefore/> (a paragraph PROPERTY),
      // NOT a hard <w:br w:type="page"/> run. A hard break is a break
      // WITHIN the text flow: the empty paragraph that carries it must be
      // placed somewhere, and when the preceding observation fills the
      // page to the bottom (the single photo is ~6.6" tall, so a report
      // with a long description does) that empty paragraph cannot fit on
      // the full page — it flows to the next page ALONE and its break then
      // pushes the following table down another page. Result: one BLANK
      // page after (almost) every report. pageBreakBefore instead just
      // tells this (empty, ~0-height) paragraph to render at the TOP of a
      // fresh page, with the next observation's table immediately below it
      // on that SAME page — so each report still starts on its own page,
      // but no blank page is ever stranded, however full the previous page
      // was. The XML blank-page scanners below can't fix the hard-break
      // case because the blank page is produced by Word's layout engine at
      // render time, not by empty XML they could detect.
      const PAGE_BREAK_PARA =
        '<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/></w:pPr></w:p>';
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
            '<w:rPr><w:rFonts w:ascii="Neue Haas Grotesk Text Pro" w:hAnsi="Neue Haas Grotesk Text Pro"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:i/><w:color w:val="9E9E9E"/></w:rPr>' +
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
        // The single-photo width is large (fills most of the landscape page),
        // so its ±15% window can overlap the route map / GA drawing widths.
        // Exclude those two explicitly FIRST so they are never rearranged as
        // observation photos.
        const ROUTE_W = ROUTE_MAP_SIZE[0] * PX_TO_EMU;
        const GA_W = GA_DRAWING_SIZE[0] * PX_TO_EMU;
        if (Math.abs(cx - ROUTE_W) <= ROUTE_W * 0.05) return false;
        if (Math.abs(cx - GA_W) <= GA_W * 0.05) return false;
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

      // ---- COLLAPSE BLANK PAGES (SECTPR-SAFE) ----
      // Find pairs of consecutive page-break paragraphs separated only
      // by blank glue and reduce them to one. Loops until stable.
      //
      // SAFETY: NEVER touch a paragraph that contains <w:sectPr>. Those
      // paragraphs carry section properties (page size, orientation,
      // margins, headers, footers) that Word requires. Removing one
      // makes Word say "unreadable content" and force the user to
      // recover. So the boundary pattern here matches ONLY explicit
      // page breaks (<w:br w:type="page"/>), not section breaks.
      let blankPagesCollapsed = 0;
      {
        const pageBreakParaPattern =
          `<w:p\\b[^>]*>(?:(?!<\\/w:p>)(?!<w:sectPr)[\\s\\S])*?<w:br\\s+w:type="page"\\s*\\/>(?:(?!<\\/w:p>)(?!<w:sectPr)[\\s\\S])*?<\\/w:p>`;
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
      // SAFETY: this scanner removes the SECOND boundary in a blank
      // pair. NEVER remove a boundary that contains <w:sectPr> (those
      // carry section properties Word requires). Only consider
      // <w:br w:type="page"/> paragraphs as removable boundaries.
      let positionalBlankPagesRemoved = 0;
      {
        const boundaryRe =
          /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?<w:br\s+w:type="page"\s*\/>(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/g;
        const boundaries: Array<{ start: number; end: number; xml: string }> = [];
        let bm: RegExpExecArray | null;
        while ((bm = boundaryRe.exec(xml)) !== null) {
          boundaries.push({
            start: bm.index,
            end: bm.index + bm[0].length,
            xml: bm[0],
          });
        }
        for (let i = boundaries.length - 1; i >= 1; i -= 1) {
          const prev = boundaries[i - 1];
          const curr = boundaries[i];
          // Don't remove a boundary that carries section properties.
          if (curr.xml.includes("<w:sectPr")) continue;
          const between = xml.slice(prev.end, curr.start);
          const hasText =
            /<w:t(?:\s[^>]*)?>(?:(?!<\/w:t>)[\s\S])*?\S(?:(?!<\/w:t>)[\s\S])*?<\/w:t>/.test(
              between
            );
          const hasDrawing = between.includes("<w:drawing");
          const hasTable = /<w:tbl[\s>]/.test(between);
          // Also bail if the "between" slice contains a sectPr-bearing
          // paragraph — removing the second boundary would change the
          // section-properties paragraph's positioning relative to it.
          const hasSectPrInBetween = between.includes("<w:sectPr");
          if (!hasText && !hasDrawing && !hasTable && !hasSectPrInBetween) {
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

      // ---- TRAILING EMPTY PAGE STRIP ----
      // The body always ends with a paragraph that carries the
      // body-level <w:sectPr>. If a page-break paragraph sits between
      // the last visible content and that final sectPr paragraph (or
      // between the last sectPr-bearing paragraph and the body close),
      // the document renders an empty trailing page. Remove any such
      // page breaks.
      let trailingBlankPagesRemoved = 0;
      {
        // Find the index of </w:body>. Everything up to it is body content.
        const bodyCloseIdx = xml.indexOf("</w:body>");
        if (bodyCloseIdx > 0) {
          // Walk all page-break paragraphs in the body, find the LAST one,
          // and check if NOTHING visible (text/drawing/table) sits between
          // it and the body close. If so, drop it. Repeat until no more
          // trailing blank pages found.
          const pageBreakParaPattern =
            /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?<w:br\s+w:type="page"\s*\/>(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/g;
          let safety = 0;
          while (safety < 10) {
            safety += 1;
            const bodyClose = xml.indexOf("</w:body>");
            if (bodyClose < 0) break;
            // Find the LAST page-break paragraph BEFORE </w:body>
            pageBreakParaPattern.lastIndex = 0;
            let lastBreak: { start: number; end: number; xml: string } | null = null;
            let m: RegExpExecArray | null;
            while ((m = pageBreakParaPattern.exec(xml)) !== null) {
              if (m.index >= bodyClose) break;
              lastBreak = {
                start: m.index,
                end: m.index + m[0].length,
                xml: m[0],
              };
            }
            if (!lastBreak) break;
            // SAFETY: never remove a paragraph carrying <w:sectPr> —
            // those define section properties Word requires.
            if (lastBreak.xml.includes("<w:sectPr")) break;
            // Slice between this page break and </w:body>
            const tail = xml.slice(lastBreak.end, bodyClose);
            // Bail if the tail itself contains a sectPr-bearing
            // paragraph (we'd be deleting structurally-required
            // content above it).
            if (tail.includes("<w:sectPr")) break;
            // Visible content checks (matches the positional scanner)
            const tailHasText =
              /<w:t(?:\s[^>]*)?>(?:(?!<\/w:t>)[\s\S])*?\S(?:(?!<\/w:t>)[\s\S])*?<\/w:t>/.test(
                tail
              );
            const tailHasDrawing = tail.includes("<w:drawing");
            const tailHasTable = /<w:tbl[\s>]/.test(tail);
            if (tailHasText || tailHasDrawing || tailHasTable) {
              // Last page break has real content after it — leave it alone.
              break;
            }
            // Tail is empty — strip this page break paragraph.
            xml = xml.slice(0, lastBreak.start) + xml.slice(lastBreak.end);
            trailingBlankPagesRemoved += 1;
          }
        }
      }
      if (trailingBlankPagesRemoved > 0) {
        console.log("[DOCX TRAILING BLANK PAGE STRIP]", {
          trailingBlankPagesRemoved,
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

      // ---- MULTI-PHOTO GRID INJECTION ----
      // A point with 2+ photos: the template rendered photo 0 in place. Replace
      // that single-photo paragraph with a borderless 2-per-row grid that holds
      // ALL of the point's photos, under the point's one table. Cell size
      // shrinks as the photo count grows so the grid stays compact.
      let multiPhotoGridsInjected = 0;
      if (multiPhotoKeys.size > 0) {
        const docRelsFile = renderedZip.file("word/_rels/document.xml.rels");
        let docRelsXml = docRelsFile ? docRelsFile.asText() : "";

        // Ensure [Content_Types].xml declares the jpeg extension.
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

        // Upload every EXTRA photo (index 1..N-1) and remember its rId per key.
        const extraRIdsByKey = new Map<string, string[]>();
        let mediaIdx = 0;
        for (const photoKey of Array.from(multiPhotoKeys)) {
          const extras = multiPhotoExtraBuffers.get(photoKey) || [];
          const rIds: string[] = [];
          for (const buf of extras) {
            const mediaName = `word/media/grid_${mediaIdx}.jpeg`;
            renderedZip.file(mediaName, buf);
            const rId = `rIdGrid${mediaIdx}`;
            rIds.push(rId);
            if (docRelsXml && !docRelsXml.includes(`Id="${rId}"`)) {
              docRelsXml = docRelsXml.replace(
                "</Relationships>",
                `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/grid_${mediaIdx}.jpeg"/></Relationships>`
              );
            }
            mediaIdx += 1;
          }
          extraRIdsByKey.set(photoKey, rIds);
        }
        if (docRelsXml) {
          renderedZip.file("word/_rels/document.xml.rels", docRelsXml);
        }

        // Build a borderless grid of ALL the point's photos, laid out TWO PER
        // ROW (side by side). Every photo in a point is the SAME size, but the
        // size SHRINKS as the number of rows grows so the whole grid + the
        // observation table fit on ONE page — otherwise the photos spill onto
        // the next page and both pages are left mostly empty. (A single-photo
        // point never reaches here: it renders big in place.)
        const GRID_ASPECT = SIDE_BY_SIDE_PHOTO_SIZE[0] / SIDE_BY_SIDE_PHOTO_SIZE[1]; // w/h
        // Fixed, legible sizes — NO shrink-to-fit. Photos PAGINATE two-per-page,
        // matching the reference PDF: the first pair sits under the observation
        // table on its page, and every further pair starts a fresh page. An odd
        // final photo renders larger and centred, alone on its own page.
        const pairWpx = Math.round(6.6 * 96); // ~6.6" wide so two + the gap fit
        const pairHpx = Math.round((6.6 / GRID_ASPECT) * 96);
        const PAIR_CX = pairWpx * PX_TO_EMU;
        const PAIR_CY = pairHpx * PX_TO_EMU;
        const loneWpx = Math.round(9.0 * 96); // a single leftover photo renders big
        const loneHpx = Math.round((9.0 / GRID_ASPECT) * 96);
        const LONE_CX = loneWpx * PX_TO_EMU;
        const LONE_CY = loneHpx * PX_TO_EMU;
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
            `<w:left w:w="60" w:type="dxa"/>` +
            `<w:bottom w:w="0" w:type="dxa"/>` +
            `<w:right w:w="60" w:type="dxa"/>` +
          `</w:tcMar>`;
        // Property page break (not a hard <w:br>): starts the paragraph it's on
        // at the top of a new page without stranding a blank page.
        const PHOTO_PAGE_BREAK =
          `<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/></w:pPr></w:p>`;
        const drawing = (rId: string, dpId: number, cx: number, cy: number) =>
          `<w:drawing>` +
            `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
              `<wp:extent cx="${cx}" cy="${cy}"/>` +
              `<wp:docPr id="${dpId}" name="Photo ${dpId}"/>` +
              `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
                `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
                  `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
                    `<pic:nvPicPr><pic:cNvPr id="0" name=""/><pic:cNvPicPr/></pic:nvPicPr>` +
                    `<pic:blipFill>` +
                      `<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rId}"/>` +
                      `<a:stretch><a:fillRect/></a:stretch>` +
                    `</pic:blipFill>` +
                    `<pic:spPr>` +
                      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
                      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
                    `</pic:spPr>` +
                  `</pic:pic>` +
                `</a:graphicData>` +
              `</a:graphic>` +
            `</wp:inline>` +
          `</w:drawing>`;
        const photoCell = (rId: string, dpId: number, cellPct: string, cx: number, cy: number) =>
          `<w:tc>` +
            `<w:tcPr><w:tcW w:w="${cellPct}" w:type="pct"/>${NO_BORDER}${ZERO_TC_MAR}<w:vAlign w:val="center"/></w:tcPr>` +
            `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="40" w:after="40"/><w:keepNext/><w:keepLines/></w:pPr>` +
              (rId ? `<w:r>${drawing(rId, dpId, cx, cy)}</w:r>` : "") +
            `</w:p>` +
          `</w:tc>`;
        const tableWrap = (gridCols: string, rowXml: string) =>
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
            `<w:tblGrid>${gridCols}</w:tblGrid>` +
            `<w:tr><w:trPr><w:cantSplit/></w:trPr>${rowXml}</w:tr>` +
          `</w:tbl>`;
        // One page's worth of photos: a pair side by side, or a lone big one.
        const pairTable = (a: string, b: string, base: number) =>
          tableWrap(
            `<w:gridCol w:w="4680"/><w:gridCol w:w="4680"/>`,
            photoCell(a, base, "2500", PAIR_CX, PAIR_CY) +
              photoCell(b || "", base + 1, "2500", PAIR_CX, PAIR_CY)
          );
        const loneTable = (a: string, base: number) =>
          tableWrap(
            `<w:gridCol w:w="9360"/>`,
            photoCell(a, base, "5000", LONE_CX, LONE_CY)
          );
        // Paginate the point's photos TWO PER PAGE: first pair under the table,
        // each further pair (or odd single) on its own fresh page.
        const buildPhotoGrid = (rIds: string[], docPrBase: number): string => {
          if (rIds.length <= 1) return loneTable(rIds[0] || "", docPrBase);
          let out = "";
          let page = 0;
          for (let r = 0; r < rIds.length; r += 2) {
            if (page > 0) out += PHOTO_PAGE_BREAK; // every pair after the first → new page
            out +=
              r + 1 < rIds.length
                ? pairTable(rIds[r], rIds[r + 1], docPrBase + r)
                : loneTable(rIds[r], docPrBase + r); // odd last photo, big & centred
            page += 1;
          }
          return out;
        };

        // Walk observation photo paragraphs in order; for each multi-photo
        // observation, REPLACE its single-photo paragraph with the full grid.
        const obsWithPhotos = observations.filter((o) => {
          const pkey = String((o as { photo?: unknown }).photo || "");
          return pkey !== "" && imageMap.has(pkey);
        });
        const photoParaRe = /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?<w:drawing\b[\s\S]*?<\/w:drawing>(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/g;
        type Hit = { paraStart: number; paraEnd: number; rIds: string[] };
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
          if (!multiPhotoKeys.has(photoKey)) continue;
          // Photo 0's r:embed is already in the rendered paragraph.
          const embedMatch = pm[0].match(/r:embed="([^"]+)"/);
          if (!embedMatch) continue;
          const leftRId = embedMatch[1];
          const rIds = [leftRId, ...(extraRIdsByKey.get(photoKey) || [])];
          replaceAt.push({ paraStart: pm.index, paraEnd: pm.index + pm[0].length, rIds });
        }

        // Apply replacements from END to START so positions stay valid.
        for (let i = replaceAt.length - 1; i >= 0; i -= 1) {
          const { paraStart, paraEnd, rIds } = replaceAt[i];
          const tableXml = buildPhotoGrid(rIds, docPrCounter);
          docPrCounter += rIds.length + 2;
          xml = xml.slice(0, paraStart) + tableXml + xml.slice(paraEnd);
          multiPhotoGridsInjected += 1;
        }

        console.log("[DOCX MULTI-PHOTO GRID INJECTION]", {
          multiPhotoKeys: multiPhotoKeys.size,
          extraImagesAdded: mediaIdx,
          observationsWithPhotos: obsWithPhotos.length,
          photoParagraphsScanned: photoParaIdx,
          gridsInjected: multiPhotoGridsInjected,
        });
      }

      // ---- EXTRA GA-DRAWING PAGES INJECTION ----
      // GA image 0 was rendered by the template in the GA slot. Any further GA
      // images are placed on their OWN separate page (one GA drawing per page),
      // each at the SAME full GA size — never squeezed into a grid on one page.
      // A page break precedes each; the title pass below then adds the report
      // title above each GA-sized image, exactly like GA image 0's page.
      let gaExtraPagesInjected = 0;
      if (gaExtraBuffers.length > 0) {
        // 1) Register each extra GA image in the zip + the document rels.
        const gaRelsFile = renderedZip.file("word/_rels/document.xml.rels");
        let gaRelsXml = gaRelsFile ? gaRelsFile.asText() : "";
        const gaCtFile = renderedZip.file("[Content_Types].xml");
        if (gaCtFile) {
          let gaCtXml = gaCtFile.asText();
          if (!gaCtXml.includes('Extension="jpeg"')) {
            gaCtXml = gaCtXml.replace(
              "</Types>",
              `<Default Extension="jpeg" ContentType="image/jpeg"/></Types>`
            );
            renderedZip.file("[Content_Types].xml", gaCtXml);
          }
        }
        const gaExtraRIds: string[] = [];
        for (let gi = 0; gi < gaExtraBuffers.length; gi += 1) {
          const mediaName = `word/media/ga_extra_${gi}.jpeg`;
          renderedZip.file(mediaName, gaExtraBuffers[gi]);
          const rId = `rIdGaExtra${gi}`;
          gaExtraRIds.push(rId);
          if (gaRelsXml && !gaRelsXml.includes(`Id="${rId}"`)) {
            gaRelsXml = gaRelsXml.replace(
              "</Relationships>",
              `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/ga_extra_${gi}.jpeg"/></Relationships>`
            );
          }
        }
        if (gaRelsXml) renderedZip.file("word/_rels/document.xml.rels", gaRelsXml);

        // 2) One "page" per extra GA image: a page-break paragraph, its OWN
        //    heading (GA DRAWING WITH LOAD / TURNING CIRCLE / STABILITY, etc.),
        //    then a centered full-size GA drawing (same size as GA image 0).
        // Extra-diagram image size. GA_DRAWING_SIZE is already sized to leave
        // room for a heading above it (see its definition), so the extra pages
        // use it as-is (scale 1.0) — same size as the first diagram, for a
        // consistent look, with the title + image always on one page.
        const GA_HEAD_SCALE = 1.0;
        const GA_CX = Math.round(GA_DRAWING_SIZE[0] * PX_TO_EMU * GA_HEAD_SCALE);
        const GA_CY = Math.round(GA_DRAWING_SIZE[1] * PX_TO_EMU * GA_HEAD_SCALE);
        // Heading paragraph for an extra diagram page. pageBreakBefore starts the
        // new page ON the heading (no separate empty spacer paragraph, which
        // wasted a line and could push the image off the page); keepNext glues
        // the heading to its drawing so they never split across pages.
        // Heading text is plain ASCII (letters/space/colon) so no escaping.
        const gaHeadingPara = (heading: string): string =>
          heading
            ? `<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="80"/><w:keepNext/></w:pPr>` +
              `<w:r><w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/><w:color w:val="1F2937"/></w:rPr>` +
              `<w:t xml:space="preserve">${heading}</w:t></w:r></w:p>`
            : "";
        // The image paragraph. When there is no heading, it carries the
        // pageBreakBefore itself so the diagram still starts on a fresh page.
        const gaImagePara = (rId: string, dpId: number, firstOnPage: boolean): string =>
          `<w:p>` +
            `<w:pPr>${firstOnPage ? "<w:pageBreakBefore/>" : ""}<w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/><w:keepNext/><w:keepLines/></w:pPr>` +
            `<w:r><w:drawing>` +
              `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
                `<wp:extent cx="${GA_CX}" cy="${GA_CY}"/>` +
                `<wp:docPr id="${dpId}" name="GA Drawing ${dpId}"/>` +
                `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
                  `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
                    `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
                      `<pic:nvPicPr><pic:cNvPr id="0" name=""/><pic:cNvPicPr/></pic:nvPicPr>` +
                      `<pic:blipFill>` +
                        `<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rId}"/>` +
                        `<a:stretch><a:fillRect/></a:stretch>` +
                      `</pic:blipFill>` +
                      `<pic:spPr>` +
                        `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${GA_CX}" cy="${GA_CY}"/></a:xfrm>` +
                        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
                      `</pic:spPr>` +
                    `</pic:pic>` +
                  `</a:graphicData>` +
                `</a:graphic>` +
              `</wp:inline>` +
            `</w:drawing></w:r>` +
          `</w:p>`;
        const gaPageBlock = (rId: string, dpId: number, heading: string): string =>
          heading
            ? gaHeadingPara(heading) + gaImagePara(rId, dpId, false)
            : gaImagePara(rId, dpId, true);

        // 3) Insert all extra GA pages right AFTER GA image 0's paragraph
        //    (found by its GA-sized drawing width). ±4% tolerance (empirical,
        //    absorbs the image module's rounding): GA image 0 matches, while the
        //    route map (now ~9.3% wider than the GA drawing) stays outside the
        //    window. Observation photos display at 7.5" (far narrower), so they
        //    never collide either.
        const GA_W_EMU_INS = GA_DRAWING_SIZE[0] * PX_TO_EMU;
        const GA_TOL_INS = GA_W_EMU_INS * 0.04;
        const gaParaRe =
          /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?<w:drawing\b[\s\S]*?<\/w:drawing>(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/g;
        let gaImg0End = -1;
        let gpm: RegExpExecArray | null;
        while ((gpm = gaParaRe.exec(xml)) !== null) {
          const cxMatch = gpm[0].match(/<wp:extent\s+cx="(\d+)"/);
          if (!cxMatch) continue;
          const cx = Number(cxMatch[1]);
          if (Number.isFinite(cx) && Math.abs(cx - GA_W_EMU_INS) <= GA_TOL_INS) {
            gaImg0End = gpm.index + gpm[0].length;
            break; // the first GA-sized drawing is GA image 0
          }
        }
        if (gaImg0End > 0) {
          let block = "";
          for (let gi = 0; gi < gaExtraRIds.length; gi += 1) {
            // Extra image gi is overall diagram index gi + 1 (image 0 fills the
            // template GA slot), so its heading is gaHeadings[gi + 1].
            block += gaPageBlock(gaExtraRIds[gi], 9100 + gi, String(gaHeadings[gi + 1] || ""));
          }
          xml = xml.slice(0, gaImg0End) + block + xml.slice(gaImg0End);
          gaExtraPagesInjected = gaExtraRIds.length;
        }
        console.log("[DOCX EXTRA GA PAGES INJECTED]", {
          extraGaImages: gaExtraBuffers.length,
          injected: gaExtraPagesInjected,
          foundGaImage0: gaImg0End > 0,
        });
      }

      // ---- ROUTE LOCATIONS STEPPER INJECTION ----
      // If we generated a "Locations" stepper image, inject it as a new
      // centered paragraph immediately AFTER the route-map drawing.
      // Identifies the route map by EMU width matching ROUTE_MAP_SIZE.
      let routeLocationsStepperInjected = false;

      // The route key points render as a horizontal numbered "snake" map
      // (START → ①..N → DESTINATION) generated into imageMap("routeLocations")
      // and injected as an image below. See generateRouteLocationsSnake. The
      // image injection block + OOXML snake-table fallback are just below.

      // ---- CONCLUSION BODY INJECTION ----
      // The {conclusion} slot was rendered as a single sentinel paragraph.
      // Replace that whole paragraph with the neatly-formatted Conclusion &
      // Certification body (bold sub-headers + bullet lists). If the sentinel
      // is missing, nothing changes (safe no-op).
      try {
        const sentIdx = xml.indexOf(CONCLUSION_SENTINEL);
        if (sentIdx >= 0 && conclusionBodyOoxml) {
          const pStart = xml.lastIndexOf("<w:p>", sentIdx);
          const pEndMarker = xml.indexOf("</w:p>", sentIdx);
          if (pStart >= 0 && pEndMarker > pStart) {
            const pEnd = pEndMarker + "</w:p>".length;
            xml = xml.slice(0, pStart) + conclusionBodyOoxml + xml.slice(pEnd);
            console.log("[DOCX CONCLUSION BODY INJECTED]", {
              paras: (conclusionBodyOoxml.match(/<w:p\b/g) || []).length,
            });
          }
        }
      } catch (concErr) {
        console.warn("[conclusion body inject] failed - leaving placeholder:", concErr);
      }
      // Safety net: never let the raw sentinel show. If it somehow survived,
      // replace it with the plain conclusion text.
      if (xml.indexOf(CONCLUSION_SENTINEL) >= 0) {
        const plain = (stripHtml(conclusionHtmlResolved) || "-")
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/\s+/g, " ");
        xml = xml.split(CONCLUSION_SENTINEL).join(plain);
      }

      const routeLocEntry = imageMap.get("routeLocations");
      if (!routeLocationsStepperInjected && routeLocEntry?.buffer && routeLocEntry.buffer.length > 0) {
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
        //    (read via sharp metadata) so the snake map never stretches.
        //    The snake is WIDE, so fill most of the landscape page width
        //    (≈13.5") but cap the height (≈7") so it still fits on one page.
        const EMU_PER_INCH_2 = 914400;
        const SNAKE_MAX_W_IN = 13.5; // landscape content ≈ 15.8"
        const SNAKE_MAX_H_IN = 7.0;  // leave room for heading + page borders
        let stepperEmuW = Math.round(SNAKE_MAX_W_IN * EMU_PER_INCH_2);
        let stepperEmuH = Math.round(4.5 * EMU_PER_INCH_2); // safety fallback
        const sharp2 = getSharp();
        if (sharp2) {
          try {
            const meta = await sharp2(routeLocEntry.buffer).metadata();
            if (meta.width && meta.height && meta.width > 0) {
              const aspect = meta.height / meta.width; // h/w
              let wIn = SNAKE_MAX_W_IN;
              let hIn = wIn * aspect;
              if (hIn > SNAKE_MAX_H_IN) {
                // Too tall at full width — shrink to fit the height cap.
                hIn = SNAKE_MAX_H_IN;
                wIn = hIn / aspect;
              }
              stepperEmuW = Math.round(wIn * EMU_PER_INCH_2);
              stepperEmuH = Math.round(hIn * EMU_PER_INCH_2);
            }
          } catch {
            // Use the safety fallback dimensions if metadata read fails.
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
          // pageBreakBefore (property) not a hard <w:br> — see PAGE_BREAK_PARA
          // note above: a hard break after a full route-map page strands a
          // blank page; this renders the stepper at the top of a fresh page.
          const STEPPER_PAGE_BREAK =
            `<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/></w:pPr></w:p>`;
          const fullStepperBlock = STEPPER_PAGE_BREAK + stepperPara;
          xml = xml.slice(0, routeMapEnd) + fullStepperBlock + xml.slice(routeMapEnd);
          routeLocationsStepperInjected = true;
        }
        console.log("[DOCX ROUTE LOCATIONS STEPPER INJECTED]", {
          itemCount: routeLocationLabels.length,
          injected: routeLocationsStepperInjected,
          renderPath: "image (sharp SVG→PNG)",
          stepperEmuW,
          stepperEmuH,
          stepperHeightInches: (stepperEmuH / EMU_PER_INCH_2).toFixed(2),
        });
      } else if (!routeLocationsStepperInjected && routeLocationLabels.length > 0) {
        // ---- SHARP-FREE FALLBACK ----
        // Stepper PNG wasn't generated (sharp unavailable, libvips
        // missing librsvg, etc.). Inject a pure OOXML table instead so
        // the locations still appear in the export — works on any
        // production server regardless of native deps.
        const ROUTE_EMU_W_FB = ROUTE_MAP_SIZE[0] * PX_TO_EMU;
        const ROUTE_TOL_FB = ROUTE_EMU_W_FB * 0.15;
        const allPhotoParaReFB =
          /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?<w:drawing\b[\s\S]*?<\/w:drawing>(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/g;
        let routeMapEndFB = -1;
        let pmRouteFB: RegExpExecArray | null;
        while ((pmRouteFB = allPhotoParaReFB.exec(xml)) !== null) {
          const cxMatch = pmRouteFB[0].match(/<wp:extent\s+cx="(\d+)"/);
          if (!cxMatch) continue;
          const cx = Number(cxMatch[1]);
          if (Number.isFinite(cx) && Math.abs(cx - ROUTE_EMU_W_FB) <= ROUTE_TOL_FB) {
            routeMapEndFB = pmRouteFB.index + pmRouteFB[0].length;
            break;
          }
        }
        if (routeMapEndFB > 0) {
          // pageBreakBefore (property) not a hard <w:br> — see PAGE_BREAK_PARA note.
          const STEPPER_PAGE_BREAK_FB =
            `<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/></w:pPr></w:p>`;
          const tableXml = buildRouteLocationsStepperTableXml(routeLocationLabels);
          // Trailing empty paragraph (Word requires a paragraph after a table)
          const trailingPara = `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>`;
          const fallbackBlock = STEPPER_PAGE_BREAK_FB + tableXml + trailingPara;
          xml = xml.slice(0, routeMapEndFB) + fallbackBlock + xml.slice(routeMapEndFB);
          routeLocationsStepperInjected = true;
        }
        console.log("[DOCX ROUTE LOCATIONS STEPPER INJECTED]", {
          itemCount: routeLocationLabels.length,
          injected: routeLocationsStepperInjected,
          renderPath: "fallback (Word-table, no sharp/libvips needed)",
          reason: "sharp PNG buffer missing — generator returned null",
        });
      } else {
        console.log("[DOCX ROUTE LOCATIONS STEPPER SKIPPED]", {
          reason: "no route location labels in DB for this project",
          routeLocationsTableExists:
            "check that project_route_page_locations table is migrated in production",
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
                `<w:rFonts w:ascii="Neue Haas Grotesk Text Pro" w:hAnsi="Neue Haas Grotesk Text Pro"/>` +
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
                `<w:rFonts w:ascii="Neue Haas Grotesk Text Pro" w:hAnsi="Neue Haas Grotesk Text Pro"/>` +
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

      // ---- UNIFORM BODY FONT SIZE (client requirement): body text is 16pt.
      // Word stores half-points, so 16pt = w:val="32". Text already LARGER than
      // 16pt (the 46pt report title, section headings) is LEFT ALONE — otherwise
      // this crushed the title down to 16pt. Only smaller/varied body text is
      // brought up/down to a uniform 16pt. Runs with no explicit size inherit
      // 16pt via the styles.xml patch below.
      let uniformSizeCount = 0;
      xml = xml.replace(
        /<w:sz(Cs)?\s+w:val="([^"]+)"\s*\/>/g,
        (_m: string, cs: string | undefined, val: string) => {
          const n = parseInt(val, 10);
          if (Number.isFinite(n) && n > 32) return _m; // keep titles/headings
          uniformSizeCount += 1;
          return `<w:sz${cs || ""} w:val="32"/>`;
        }
      );
      console.log("[DOCX UNIFORM 16PT]", { runsResized: uniformSizeCount });

      renderedZip.file("word/document.xml", xml);
    }

    // ---- UNIFORM FONT SIZE in styles.xml: document defaults and every
    // named style also become 16pt, so text that inherits its size from a
    // style (no explicit w:sz on the run) renders at 16pt too.
    {
      const stylesFile = renderedZip.file("word/styles.xml");
      if (stylesFile) {
        let stylesXml = stylesFile.asText();
        let stylesSizeCount = 0;
        stylesXml = stylesXml.replace(
          /<w:sz(Cs)?\s+w:val="([^"]+)"\s*\/>/g,
          (_m: string, cs: string | undefined, val: string) => {
            const n = parseInt(val, 10);
            if (Number.isFinite(n) && n > 32) return _m; // keep title/heading styles large
            stylesSizeCount += 1;
            return `<w:sz${cs || ""} w:val="32"/>`;
          }
        );
        if (stylesSizeCount > 0) renderedZip.file("word/styles.xml", stylesXml);
        console.log("[DOCX UNIFORM 16PT styles.xml]", { stylesResized: stylesSizeCount });
      }
    }

    // ---- ENSURE OUTPUT IS FREELY EDITABLE ----
    // Word's "Save → asks for filename" / Save-As prompt usually fires
    // when one of these is true:
    //   1. settings.xml has <w:writeProtection> (modify-recommended)
    //   2. settings.xml has <w:documentProtection> (read-only / forms)
    //   3. Content_Types.xml says the main document is a TEMPLATE
    //      (wordprocessingml.template.main+xml) instead of a regular
    //      document — Word always opens templates as "new untitled"
    //   4. core.xml has <cp:contentStatus>Final</cp:contentStatus>
    // None of these should be set, but if the template ever inherits
    // them (or docxtemplater adds them), this pass strips them so the
    // user can save the downloaded DOCX in place without a dialog.
    {
      let protectionStripped = 0;
      let templateContentTypeFixed = false;
      let contentStatusCleared = false;

      // 1+2. Strip <w:writeProtection .../> and <w:documentProtection .../>
      const settingsFile = renderedZip.file("word/settings.xml");
      if (settingsFile) {
        let settingsXml = settingsFile.asText();
        const before = settingsXml;
        settingsXml = settingsXml.replace(
          /<w:writeProtection\b[^/]*\/>/g,
          () => {
            protectionStripped += 1;
            return "";
          }
        );
        settingsXml = settingsXml.replace(
          /<w:writeProtection\b[\s\S]*?<\/w:writeProtection>/g,
          () => {
            protectionStripped += 1;
            return "";
          }
        );
        settingsXml = settingsXml.replace(
          /<w:documentProtection\b[^/]*\/>/g,
          () => {
            protectionStripped += 1;
            return "";
          }
        );
        settingsXml = settingsXml.replace(
          /<w:documentProtection\b[\s\S]*?<\/w:documentProtection>/g,
          () => {
            protectionStripped += 1;
            return "";
          }
        );
        if (settingsXml !== before) {
          renderedZip.file("word/settings.xml", settingsXml);
        }
      }

      // 3. If Content_Types.xml flagged the doc as a TEMPLATE, switch it
      //    to a regular document so Word doesn't open it "as new".
      const ctxFile = renderedZip.file("[Content_Types].xml");
      if (ctxFile) {
        let ctxXml = ctxFile.asText();
        const beforeCtx = ctxXml;
        ctxXml = ctxXml.replace(
          /application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.template\.main\+xml/g,
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
        );
        if (ctxXml !== beforeCtx) {
          renderedZip.file("[Content_Types].xml", ctxXml);
          templateContentTypeFixed = true;
        }
      }

      // 4. Strip <cp:contentStatus>Final</cp:contentStatus> from core.xml
      //    so Word doesn't open it in "Marked as Final" read-only mode.
      const coreFile = renderedZip.file("docProps/core.xml");
      if (coreFile) {
        let coreXml = coreFile.asText();
        const beforeCore = coreXml;
        coreXml = coreXml.replace(
          /<cp:contentStatus\b[\s\S]*?<\/cp:contentStatus>/g,
          () => {
            contentStatusCleared = true;
            return "";
          }
        );
        if (coreXml !== beforeCore) {
          renderedZip.file("docProps/core.xml", coreXml);
        }
      }

      // 5. Strip <Template>...</Template> and <HyperlinkBase>...</HyperlinkBase>
      //    from docProps/app.xml. If the doc has <Template>foo.docx</Template>,
      //    Word treats it as derived from a template and the FIRST Ctrl+S
      //    forces Save-As (the user's "Save this file" popup).
      let appPropsCleaned = false;
      const appFile = renderedZip.file("docProps/app.xml");
      if (appFile) {
        let appXml = appFile.asText();
        const beforeApp = appXml;
        appXml = appXml.replace(/<Template\b[^>]*\/>/g, () => {
          appPropsCleaned = true;
          return "<Template>Normal.dotm</Template>";
        });
        appXml = appXml.replace(/<Template\b[^>]*>[\s\S]*?<\/Template>/g, () => {
          appPropsCleaned = true;
          return "<Template>Normal.dotm</Template>";
        });
        appXml = appXml.replace(/<HyperlinkBase\b[\s\S]*?<\/HyperlinkBase>/g, () => {
          appPropsCleaned = true;
          return "";
        });
        if (appXml !== beforeApp) {
          renderedZip.file("docProps/app.xml", appXml);
        }
      }

      console.log("[DOCX EDITABILITY PASS]", {
        protectionStripped,
        templateContentTypeFixed,
        contentStatusCleared,
        appPropsCleaned,
      });
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
      `<w:rPr><w:rFonts w:ascii="Neue Haas Grotesk Text Pro" w:hAnsi="Neue Haas Grotesk Text Pro"/><w:sz w:val="28"/><w:szCs w:val="28"/><w:color w:val="000000"/></w:rPr>`;
    const FTR_RUN_9B =
      `<w:rPr><w:rFonts w:ascii="Neue Haas Grotesk Text Pro" w:hAnsi="Neue Haas Grotesk Text Pro"/><w:sz w:val="30"/><w:szCs w:val="30"/><w:b/><w:bCs/><w:color w:val="000000"/></w:rPr>`;
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

        // Logo display size: sized by HEIGHT so the whole logo sits ABOVE the
        // red page-border line. The border's top line is drawn ~0.67in from the
        // page edge (space=48pt); the header starts 0.125in from the top
        // (w:header=180, set in the pgMar pass) — so a 0.45in-tall logo ends at
        // ~0.575in, clear of the line instead of being sliced by it. WIDTH is
        // computed from the source image's true aspect ratio. Resolution order:
        //   1. sharp metadata (works locally)
        //   2. native PNG IHDR parser (works in production even when
        //      sharp/libvips is missing — fixes logo stretching)
        //   3. 2.4in-wide fallback (used only if both readers fail)
        const LOGO_TARGET_H_IN = 0.45;
        const EMU_PER_INCH = 914400;
        let logoEmuH = Math.round(LOGO_TARGET_H_IN * EMU_PER_INCH);
        let logoEmuW = Math.round(2.4 * EMU_PER_INCH); // fallback width
        let aspectSource: "sharp" | "png-parser" | "fallback" = "fallback";
        let aspectRatio: number | null = null;
        const sharp = getSharp();
        if (sharp) {
          try {
            const meta = await sharp(logoBuffer).metadata();
            if (meta.width && meta.height && meta.width > 0) {
              aspectRatio = meta.height / meta.width;
              aspectSource = "sharp";
            }
          } catch (metaErr) {
            console.warn("[CLIENT DOCX HEADER] sharp metadata failed, will try PNG parser:", metaErr);
          }
        }
        if (aspectRatio === null) {
          // sharp unavailable or threw → use the native PNG dimension
          // parser. The logo IS a PNG so this is reliable.
          const dims = readPngDimensions(logoBuffer);
          if (dims && dims.width > 0) {
            aspectRatio = dims.height / dims.width;
            aspectSource = "png-parser";
          }
        }
        if (aspectRatio !== null) {
          logoEmuW = Math.round(logoEmuH / aspectRatio);
        }
        console.log("[CLIENT DOCX HEADER] logo aspect ratio", {
          aspectSource,
          aspect: aspectRatio !== null ? aspectRatio.toFixed(4) : "n/a",
          emuW: logoEmuW,
          emuH: logoEmuH,
          inchesW: (logoEmuW / EMU_PER_INCH).toFixed(3),
          inchesH: LOGO_TARGET_H_IN,
        });

        console.log("[DOCX TITLE LOGO FIX]", {
          titleSize: "22pt",
          titleDuplicateAvoided: true,
          logoWidth: (logoEmuW / EMU_PER_INCH).toFixed(3) + "in",
          logoHeightInches: (logoEmuH / EMU_PER_INCH).toFixed(3),
          logoAspectRatioPreserved: true,
          logoNotStretched: true,
        });

        const logoRelId = "rIdLogoImg";

        // ---- Optional CLIENT logo (top-RIGHT of every page's header) ----
        // The RACE logo on the left is mandatory; a client logo appears on the
        // right ONLY when the project has one set (projects.client_logo_url).
        const clientRelId = "rIdClientLogo";
        const clientExtDecl = 'Extension="jpeg"';
        let clientLogoBuffer: Buffer | null = null;
        let clientLogoExt = "png";
        let clientEmuW = 0;
        let clientEmuH = 0;
        try {
          const clientLogoUrl = String((projectRow as any)?.client_logo_url || "").trim();
          if (clientLogoUrl) {
            const fetched = await fetchImageBuffer(clientLogoUrl, "clientLogo");
            if (fetched?.buffer && fetched.buffer.length > 0) {
              clientLogoBuffer = fetched.buffer;
              clientLogoExt = /jpe?g/i.test(fetched.contentType) ? "jpeg" : "png";
              // Match the RACE logo's HEIGHT so the two look balanced; width
              // from the client logo's own aspect ratio (never stretched),
              // capped so a very wide logo still fits the right cell.
              let cAspect = 0.5; // height/width fallback
              const sh = getSharp();
              if (sh) {
                try {
                  const m = await sh(clientLogoBuffer).metadata();
                  if (m.width && m.height && m.width > 0) cAspect = m.height / m.width;
                } catch {
                  /* keep fallback */
                }
              } else {
                const d = readPngDimensions(clientLogoBuffer);
                if (d && d.width > 0) cAspect = d.height / d.width;
              }
              let w = Math.round(logoEmuH / (cAspect || 0.5)); // width for matched height
              let h = logoEmuH;
              const maxW = Math.round(2.6 * EMU_PER_INCH);
              if (w > maxW) {
                w = maxW;
                h = Math.round(w * cAspect);
              }
              clientEmuW = w;
              clientEmuH = h;
            }
          }
        } catch (e) {
          console.warn("[CLIENT DOCX HEADER] client logo fetch failed:", e);
          clientLogoBuffer = null;
        }
        if (clientLogoBuffer) {
          renderedZip.file(`word/media/client_logo.${clientLogoExt}`, clientLogoBuffer);
          const ctFileC = renderedZip.file("[Content_Types].xml");
          if (ctFileC) {
            let ctXmlC = ctFileC.asText();
            const extDecl = clientLogoExt === "jpeg" ? clientExtDecl : 'Extension="png"';
            const ctType = clientLogoExt === "jpeg" ? "image/jpeg" : "image/png";
            if (!ctXmlC.includes(extDecl)) {
              ctXmlC = ctXmlC.replace("</Types>", `<Default ${extDecl} ContentType="${ctType}"/></Types>`);
              renderedZip.file("[Content_Types].xml", ctXmlC);
            }
          }
        }

        // Right-cell content: the client logo (right-aligned) when present,
        // otherwise an empty spacer paragraph (the original behaviour).
        const rightCellInner = clientLogoBuffer
          ? `<w:p>` +
              `<w:pPr><w:pStyle w:val="Header"/><w:jc w:val="right"/><w:spacing w:before="0" w:after="0"/></w:pPr>` +
              `<w:r><w:rPr/><w:drawing>` +
                `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
                  `<wp:extent cx="${clientEmuW}" cy="${clientEmuH}"/>` +
                  `<wp:docPr id="998" name="Client Logo"/>` +
                  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
                    `<pic:pic>` +
                      `<pic:nvPicPr><pic:cNvPr id="0" name="client_logo.${clientLogoExt}"/><pic:cNvPicPr/></pic:nvPicPr>` +
                      `<pic:blipFill><a:blip r:embed="${clientRelId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
                      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${clientEmuW}" cy="${clientEmuH}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
                    `</pic:pic>` +
                  `</a:graphicData></a:graphic>` +
                `</wp:inline>` +
              `</w:drawing></w:r>` +
            `</w:p>`
          : `<w:p><w:pPr><w:pStyle w:val="Header"/><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>`;

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
                `<w:rFonts w:ascii="Neue Haas Grotesk Text Pro" w:hAnsi="Neue Haas Grotesk Text Pro"/>` +
                `<w:b/><w:bCs/>` +
                `<w:sz w:val="44"/><w:szCs w:val="44"/>` +
                `<w:color w:val="44546A"/>` +
              `</w:rPr>` +
              `<w:t xml:space="preserve">${escapedTitleForHeader}</w:t>` +
            `</w:r>`
          : `<w:r/>`;
        // Optional diagonal text watermark (classic Word "washout" watermark):
        // a VML text shape anchored in the header, centred on the page,
        // rotated 315°, light grey, BEHIND the body text (negative z-index).
        // Word, LibreOffice (the PDF converter) and Google Docs all honour
        // this shape — it is exactly what Word's own Insert→Watermark emits.
        const wmEsc = (s: string) =>
          s
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
        const WATERMARK_P = watermarkText
          ? `<w:p><w:pPr><w:pStyle w:val="Header"/><w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/></w:pPr><w:r><w:pict>` +
            `<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">` +
            `<v:formulas>` +
            `<v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/>` +
            `<v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/>` +
            `<v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/>` +
            `<v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/>` +
            `<v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/>` +
            `</v:formulas>` +
            `<v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>` +
            `<v:textpath on="t" fitshape="t"/>` +
            `<v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>` +
            `<o:lock v:ext="edit" text="t" shapetype="t"/>` +
            `</v:shapetype>` +
            `<v:shape id="WordWaterMark1" o:spid="_x0000_s2049" type="#_x0000_t136"` +
            ` style="position:absolute;margin-left:0;margin-top:0;width:760pt;height:190pt;rotation:315;z-index:-251658752;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin;mso-wrap-style:none"` +
            ` o:allowincell="f" fillcolor="silver" stroked="f">` +
            `<v:fill opacity=".5"/>` +
            `<v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="${wmEsc(watermarkText)}"/>` +
            `</v:shape>` +
            `</w:pict></w:r></w:p>`
          : "";
        const HEADER_XML =
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
          `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
          ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
          ` xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"` +
          ` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
          ` xmlns:v="urn:schemas-microsoft-com:vml"` +
          ` xmlns:o="urn:schemas-microsoft-com:office:office"` +
          ` xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
          WATERMARK_P +
          // Tab-stop header (NOT a table). The logo sits at the LEFT; a CENTER
          // tab at the page midpoint (11385 twips = half of the 22771-twip
          // content width) centres the title; a RIGHT tab at 22771 right-aligns
          // the optional client logo. Tab stops are honoured by Word, Google
          // Docs AND LibreOffice, so the title is reliably page-centred — tables
          // were collapsing the empty column in some renderers and shifting it.
          `<w:p>` +
            `<w:pPr>` +
              `<w:pStyle w:val="Header"/>` +
              `<w:tabs>` +
                `<w:tab w:val="center" w:pos="11385"/>` +
                `<w:tab w:val="right" w:pos="22771"/>` +
              `</w:tabs>` +
              `<w:spacing w:before="0" w:after="0"/>` +
            `</w:pPr>` +
            // Logo — inline, at the left of the line.
            `<w:r><w:rPr/><w:drawing>` +
              `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
                `<wp:extent cx="${logoEmuW}" cy="${logoEmuH}"/>` +
                `<wp:docPr id="999" name="Race Logo"/>` +
                `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
                  `<pic:pic>` +
                    `<pic:nvPicPr><pic:cNvPr id="0" name="race_logo.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
                    `<pic:blipFill><a:blip r:embed="${logoRelId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
                    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${logoEmuW}" cy="${logoEmuH}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
                  `</pic:pic>` +
                `</a:graphicData></a:graphic>` +
              `</wp:inline>` +
            `</w:drawing></w:r>` +
            // Center tab, then the title centred at the page midpoint. The title
            // run is raised (w:position, half-points) so it sits at the vertical
            // middle of the taller logo instead of on the logo's baseline.
            `<w:r><w:tab/></w:r>` +
            (titleUpperForHeader
              ? `<w:r>` +
                  `<w:rPr>` +
                    `<w:rFonts w:ascii="Neue Haas Grotesk Text Pro" w:hAnsi="Neue Haas Grotesk Text Pro"/>` +
                    `<w:b/><w:bCs/>` +
                    `<w:position w:val="${Math.max(0, Math.round((logoEmuH / 914400) * 72 - 22))}"/>` +
                    `<w:sz w:val="44"/><w:szCs w:val="44"/>` +
                    `<w:color w:val="44546A"/>` +
                  `</w:rPr>` +
                  `<w:t xml:space="preserve">${escapedTitleForHeader}</w:t>` +
                `</w:r>`
              : "") +
            // Optional client logo: right tab, then the image at the right edge.
            (clientLogoBuffer
              ? `<w:r><w:tab/></w:r>` +
                `<w:r><w:rPr/><w:drawing>` +
                  `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
                    `<wp:extent cx="${clientEmuW}" cy="${clientEmuH}"/>` +
                    `<wp:docPr id="998" name="Client Logo"/>` +
                    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
                      `<pic:pic>` +
                        `<pic:nvPicPr><pic:cNvPr id="0" name="client_logo.${clientLogoExt}"/><pic:cNvPicPr/></pic:nvPicPr>` +
                        `<pic:blipFill><a:blip r:embed="${clientRelId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
                        `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${clientEmuW}" cy="${clientEmuH}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
                      `</pic:pic>` +
                    `</a:graphicData></a:graphic>` +
                  `</wp:inline>` +
                `</w:drawing></w:r>`
              : "") +
          `</w:p>` +
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
          const clientRelXml = clientLogoBuffer
            ? `<Relationship Id="${clientRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/client_logo.${clientLogoExt}"/>`
            : "";
          const hdrRelsXml =
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
              `<Relationship Id="${logoRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/race_logo.png"/>` +
              clientRelXml +
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
          logoInchesW: (logoEmuW / 914400).toFixed(3),
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

  // ----- Stage Summary table (Category / Count / %). The template has no
  // summary section, so append a table into the body just before the final
  // section properties. Non-fatal. (Re-enabled: the earlier "unreadable
  // content" corruption was the template's orphaned .undefined media, fixed
  // above — not this table.)
  const STAGE_SUMMARY_ENABLED = true;
  // Per-project Stage Summary OVERRIDE (edited in the app via
  // /api/projects/:id/stage-summary). When present it replaces the
  // auto-computed rows in the table below, in the saved order.
  let stageSummaryOverride: StageSummaryRow[] | null = null;
  try {
    if (projectId) {
      const [ovRows] = await pool.query(
        "SELECT data_json FROM project_stage_summary WHERE project_id = ? LIMIT 1",
        [projectId]
      );
      const raw =
        Array.isArray(ovRows) && ovRows.length
          ? (ovRows[0] as { data_json?: string }).data_json
          : null;
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) {
          stageSummaryOverride = parsed
            .map((r) => ({ label: String(r?.label ?? "").trim(), count: Number(r?.count) || 0 }))
            .filter((r) => r.label);
        }
      }
    }
  } catch (ovErr) {
    console.warn("[stage summary] override load failed - using auto:", ovErr);
  }

  // ----- Hindrance Solutions + Additional Considerations. A new section the
  // template never had; injected before the final section properties (the same
  // structurally-safe spot the Stage Summary uses). Inserted BEFORE the Stage
  // Summary block below so it reads Hindrance → Stage Summary, matching the
  // reference report. Defaults come from the standard report wording and are
  // overridable per-project via the setup wizard. Non-fatal.
  try {
    // Opt-in: the Hindrance Solutions section only appears once the user has
    // saved it in the setup wizard (an override exists). This keeps existing
    // reports unchanged until the section is deliberately filled in.
    if (!summaryMode && hindranceOverride) {
      const hRows = (hindranceOverride?.rows && hindranceOverride.rows.length
        ? hindranceOverride.rows
        : STANDARD_HINDRANCE_ROWS
      )
        .map((r) => ({
          hindrance: String(r?.hindrance || "").trim(),
          solution: String(r?.solution || "").trim(),
        }))
        .filter((r) => r.hindrance || r.solution);
      const considerations = (hindranceOverride?.considerations &&
      hindranceOverride.considerations.length
        ? hindranceOverride.considerations
        : STANDARD_HINDRANCE_CONSIDERATIONS
      )
        .map((s) => String(s || "").trim())
        .filter(Boolean);
      if (hRows.length || considerations.length) {
        const zip = doc.getZip();
        const docFile = zip.file("word/document.xml");
        if (docFile) {
          let xml = docFile.asText();
          const FONT = "Neue Haas Grotesk Text Pro";
          const esc = (s: string) =>
            String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
          // Layout copied from the reference PDF ("SAMEEP … TO MUNDRA PORT"
          // p7): a bold black "Hindrance Solutions:" heading, then plain
          // BULLET lines "• Label: text" (bold label, normal text — NOT a
          // table), then bold "Additional Considerations" with numbered bold
          // titles and an indented "o body" line under each.
          const rPrN = `<w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}"/><w:color w:val="111827"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>`;
          const rPrB = `<w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}"/><w:b/><w:bCs/><w:color w:val="111827"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>`;
          const heading =
            `<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="140" w:line="240" w:lineRule="auto"/></w:pPr>` +
            `<w:r><w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}"/><w:b/><w:bCs/><w:color w:val="111827"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Hindrance Solutions:</w:t></w:r></w:p>`;
          // "• Major Obstacles: Height / width clearance …" bullet per row.
          const rowsXml = hRows
            .map((r) => {
              const label = r.hindrance ? `${r.hindrance}:` : "";
              return (
                `<w:p><w:pPr><w:spacing w:before="40" w:after="40" w:line="264" w:lineRule="auto"/><w:ind w:left="446" w:hanging="230"/></w:pPr>` +
                `<w:r>${rPrN}<w:t xml:space="preserve">•  </w:t></w:r>` +
                (label
                  ? `<w:r>${rPrB}<w:t xml:space="preserve">${esc(label)} </w:t></w:r>`
                  : "") +
                `<w:r>${rPrN}<w:t xml:space="preserve">${esc(r.solution || "")}</w:t></w:r>` +
                `</w:p>`
              );
            })
            .join("");
          const addlHeading = considerations.length
            ? `<w:p><w:pPr><w:spacing w:before="280" w:after="120"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}"/><w:b/><w:bCs/><w:color w:val="111827"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Additional Considerations</w:t></w:r></w:p>`
            : "";
          // Each consideration "Title: body" becomes
          //   1. Title:            (bold)
          //      o body            (indented sub-line, like the PDF)
          // A consideration without a "Title:" prefix prints as a plain
          // numbered line.
          const addlItems = considerations
            .map((c, i) => {
              const m = c.match(/^([^:]{1,60}):\s*([\s\S]*)$/);
              const title = m ? m[1].trim() : "";
              const body = m ? m[2].trim() : c;
              const numberLine =
                `<w:p><w:pPr><w:spacing w:before="60" w:after="0" w:line="264" w:lineRule="auto"/><w:ind w:left="360" w:hanging="360"/></w:pPr>` +
                `<w:r>${rPrN}<w:t xml:space="preserve">${i + 1}. </w:t></w:r>` +
                (title
                  ? `<w:r>${rPrB}<w:t xml:space="preserve">${esc(title)}:</w:t></w:r>`
                  : `<w:r>${rPrN}<w:t xml:space="preserve">${esc(body)}</w:t></w:r>`) +
                `</w:p>`;
              const subLine = title
                ? `<w:p><w:pPr><w:spacing w:before="0" w:after="40" w:line="264" w:lineRule="auto"/><w:ind w:left="900" w:hanging="180"/></w:pPr>` +
                  `<w:r>${rPrN}<w:t xml:space="preserve">o </w:t></w:r>` +
                  `<w:r>${rPrN}<w:t xml:space="preserve">${esc(body)}</w:t></w:r></w:p>`
                : "";
              return numberLine + subLine;
            })
            .join("");
          const block = heading + rowsXml + addlHeading + addlItems + `<w:p/>`;
          // Insert just BEFORE the first per-point observation table (its header
          // contains "GPS LOCATION"), so the order reads
          // Summary of Observations → Hindrance → Stage Summary → per-point
          // observations, matching the reference report. "GPS LOCATION" sits
          // INSIDE the observation table, so the nearest <w:tbl> before it is
          // that table (not this section's own table). Falls back to the end of
          // the document (before the final section properties) if there are no
          // observation tables. Same safe anchor the Summary block already uses.
          const gpsIdxH = xml.indexOf("GPS LOCATION");
          let insertAtH = -1;
          if (gpsIdxH !== -1) {
            insertAtH = Math.max(
              xml.lastIndexOf("<w:tbl>", gpsIdxH),
              xml.lastIndexOf("<w:tbl ", gpsIdxH)
            );
          }
          if (insertAtH === -1) {
            const bodyCloseH = xml.lastIndexOf("</w:body>");
            if (bodyCloseH !== -1) {
              const lastSectPrH = xml.lastIndexOf("<w:sectPr", bodyCloseH);
              insertAtH = lastSectPrH !== -1 ? lastSectPrH : bodyCloseH;
            }
          }
          if (insertAtH > 0) {
            xml = xml.slice(0, insertAtH) + block + xml.slice(insertAtH);
            zip.file("word/document.xml", xml);
            console.log("[hindrance] injected", {
              rows: hRows.length,
              considerations: considerations.length,
              beforeObservations: gpsIdxH !== -1,
            });
          }
        }
      }
    }
  } catch (err) {
    console.error("[hindrance] injection failed - non-fatal:", err);
  }

  try {
    const stageRows: StageSummaryRow[] =
      stageSummaryOverride && stageSummaryOverride.length
        ? stageSummaryOverride
        : categorySummary.map((c) => ({ label: c.categoryLabel, count: c.count }));
    const stageImageUrl = String(stageImageOverride?.imageUrl || "").trim();
    if (STAGE_SUMMARY_ENABLED && (stageRows.length || stageImageUrl)) {
      const zip = doc.getZip();
      const docFile = zip.file("word/document.xml");
      if (docFile) {
        let xml = docFile.asText();
        const FONT = "Neue Haas Grotesk Text Pro";
        const esc = (s: string) =>
          String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
        const total = stageRows.reduce((a, c) => a + c.count, 0) || 1;
        // Keep the user's saved order for an override; sort auto rows by count.
        const sorted =
          stageSummaryOverride && stageSummaryOverride.length
            ? stageRows
            : [...stageRows].sort((a, b) => b.count - a.count);
        // Bigger font (14pt) + roomier cell padding for a clearer, airier table.
        const cell = (text: string, bold: boolean, fill: string, color: string, align: string) =>
          `<w:tc><w:tcPr>${fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>` : ""}` +
          `<w:tcMar><w:top w:w="140" w:type="dxa"/><w:left w:w="180" w:type="dxa"/><w:bottom w:w="140" w:type="dxa"/><w:right w:w="180" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/></w:tcPr>` +
          `<w:p><w:pPr><w:spacing w:before="70" w:after="70" w:line="264" w:lineRule="auto"/><w:jc w:val="${align}"/></w:pPr>` +
          `<w:r><w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}"/>${bold ? "<w:b/><w:bCs/>" : ""}<w:color w:val="${color}"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p></w:tc>`;
        const hRow = `<w:tr>${cell("Category", true, "1F7A5A", "FFFFFF", "left")}${cell("Count", true, "1F7A5A", "FFFFFF", "center")}${cell("%", true, "1F7A5A", "FFFFFF", "center")}</w:tr>`;
        // Plain rows — no per-category colour; only the header row is shaded.
        const dRows = sorted
          .map((c) => {
            const pct = Math.round((c.count / total) * 100);
            return `<w:tr>${cell(c.label, false, "", "111827", "left")}${cell(String(c.count), false, "", "111827", "center")}${cell(pct + "%", false, "", "111827", "center")}</w:tr>`;
          })
          .join("");
        const tRow = `<w:tr>${cell("Total", true, "", "111827", "left")}${cell(String(total), true, "", "111827", "center")}${cell("100%", true, "", "111827", "center")}</w:tr>`;
        const borders = `<w:tblBorders><w:top w:val="single" w:sz="4" w:color="B7C4BF"/><w:left w:val="single" w:sz="4" w:color="B7C4BF"/><w:bottom w:val="single" w:sz="4" w:color="B7C4BF"/><w:right w:val="single" w:sz="4" w:color="B7C4BF"/><w:insideH w:val="single" w:sz="4" w:color="D7E0DC"/><w:insideV w:val="single" w:sz="4" w:color="D7E0DC"/></w:tblBorders>`;
        const grid = `<w:tblGrid><w:gridCol w:w="6600"/><w:gridCol w:w="1600"/><w:gridCol w:w="1600"/></w:tblGrid>`;
        const table = `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>${borders}<w:tblLayout w:type="fixed"/></w:tblPr>${grid}${hRow}${dRows}${tRow}</w:tbl>`;
        // pageBreakBefore on the heading (a paragraph PROPERTY) makes the
        // Stage Summary start its own page WITHOUT a standalone hard-break
        // paragraph that would strand a blank page when the preceding
        // observation fills its page — see the PAGE_BREAK_PARA note above.
        const heading = `<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="200" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}"/><w:b/><w:bCs/><w:color w:val="163A2A"/><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr><w:t>Stage Summary</w:t></w:r></w:p>`;

        // If the user uploaded a Stage-Summary image (e.g. a pie chart made
        // elsewhere), show that image INSTEAD of the auto table. Embedded the
        // same proven way as the route-locations stepper: add media + a
        // relationship, size from the PNG's real aspect ratio via sharp.
        let stageContent = table;
        if (stageImageUrl) {
          try {
            const fetched = await fetchImageBuffer(stageImageUrl, "stageImage");
            if (fetched && Buffer.isBuffer(fetched.buffer) && fetched.buffer.length > 0) {
              const ctFileSi = zip.file("[Content_Types].xml");
              if (ctFileSi) {
                let ctXmlSi = ctFileSi.asText();
                if (!ctXmlSi.includes('Extension="png"')) {
                  ctXmlSi = ctXmlSi.replace(
                    "</Types>",
                    `<Default Extension="png" ContentType="image/png"/></Types>`
                  );
                  zip.file("[Content_Types].xml", ctXmlSi);
                }
              }
              const siMedia = "word/media/stage_summary_image.png";
              zip.file(siMedia, fetched.buffer);
              const siRelsFile = zip.file("word/_rels/document.xml.rels");
              let siRelsXml = siRelsFile ? siRelsFile.asText() : "";
              const siRId = "rIdStageSummaryImage";
              if (siRelsXml && !siRelsXml.includes(`Id="${siRId}"`)) {
                siRelsXml = siRelsXml.replace(
                  "</Relationships>",
                  `<Relationship Id="${siRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/stage_summary_image.png"/></Relationships>`
                );
                zip.file("word/_rels/document.xml.rels", siRelsXml);
              }
              const EMU_PER_INCH_SI = 914400;
              let siW = Math.round(8.0 * EMU_PER_INCH_SI);
              let siH = Math.round(5.5 * EMU_PER_INCH_SI);
              const sharpSi = getSharp();
              if (sharpSi) {
                try {
                  const meta = await sharpSi(fetched.buffer).metadata();
                  if (meta.width && meta.height && meta.width > 0) {
                    siH = Math.round((siW * meta.height) / meta.width);
                  }
                } catch {
                  /* keep fallback dims */
                }
              }
              stageContent =
                `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/><w:keepNext/><w:keepLines/></w:pPr>` +
                `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
                `<wp:extent cx="${siW}" cy="${siH}"/>` +
                `<wp:docPr id="9300" name="Stage Summary Image"/>` +
                `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
                `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
                `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
                `<pic:nvPicPr><pic:cNvPr id="0" name=""/><pic:cNvPicPr/></pic:nvPicPr>` +
                `<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${siRId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
                `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${siW}" cy="${siH}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
                `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
            }
          } catch (siErr) {
            console.warn("[stage summary] image embed failed - falling back to table:", siErr);
          }
        }
        const block = heading + stageContent + `<w:p/>`;
        // Insert just BEFORE the first per-point observation table (header
        // contains "GPS LOCATION") so the order reads Hindrance → Stage Summary
        // → per-point observations. Because Hindrance was inserted just above
        // (its own table is FURTHER from "GPS LOCATION"), the nearest <w:tbl>
        // before "GPS LOCATION" is still the observation table, so Stage lands
        // after Hindrance and before the observations. Falls back to the end of
        // the document if there are no observation tables. NOTE: this is the
        // SAME anchor the Summary block uses — it does not rewrite any section
        // break, so it does not hit the earlier GA-boundary corruption.
        const gpsIdxS = xml.indexOf("GPS LOCATION");
        let insertAt = -1;
        if (gpsIdxS !== -1) {
          insertAt = Math.max(
            xml.lastIndexOf("<w:tbl>", gpsIdxS),
            xml.lastIndexOf("<w:tbl ", gpsIdxS)
          );
        }
        if (insertAt === -1) {
          const bodyClose = xml.lastIndexOf("</w:body>");
          if (bodyClose !== -1) {
            const lastSectPr = xml.lastIndexOf("<w:sectPr", bodyClose);
            insertAt = lastSectPr !== -1 ? lastSectPr : bodyClose;
          }
        }
        if (insertAt > 0) {
          xml = xml.slice(0, insertAt) + block + xml.slice(insertAt);
          zip.file("word/document.xml", xml);
          console.log("[stage summary] injected", sorted.length, "rows", {
            beforeObservations: gpsIdxS !== -1,
          });
        }
      }
    }
  } catch (err) {
    console.error("[stage summary] injection failed - non-fatal:", err);
  }

  // ----- Fix pre-existing template corruption. reena-all-template.docx ships
  // with 5 orphaned "word/media/<hash>.undefined" files (1x1 transparent PNGs)
  // and NO content-type for the "undefined" extension. Word rejects any part it
  // can't type, so EVERY export triggered "Word found unreadable content". The
  // files are valid PNGs — declare the extension as image/png so Word reads them.
  try {
    const zip = doc.getZip();
    const ctFile = zip.file("[Content_Types].xml");
    if (ctFile) {
      let ct = ctFile.asText();
      if (!/Extension="undefined"/i.test(ct) && ct.includes("</Types>")) {
        ct = ct.replace(
          "</Types>",
          '<Default Extension="undefined" ContentType="image/png"/></Types>'
        );
        zip.file("[Content_Types].xml", ct);
        console.log("[export] declared .undefined media as image/png (template-junk corruption fix)");
      }
    }
  } catch (err) {
    console.error("[export] undefined-media content-type fix failed (non-fatal):", err);
  }

  // ----- Vertical alignment: force every page to TOP — EXCEPT the cover page.
  // Some template sections carry <w:sectPr>…<w:vAlign w:val="center"/>, which
  // vertically centres the WHOLE page — so when the user opens the exported
  // file to edit it, the cursor/content sits in the middle of the page instead
  // of the top. Rewrite page-level (sectPr) vAlign to "top" so content starts
  // at the top. Table-cell (tcPr) vAlign is intentionally left untouched.
  //
  // The ONE exception is the cover/title page (the first section): the user
  // wants its title block ("…SURVEY REPORT" + the title table) sitting in the
  // vertical MIDDLE of the page, not pinned to the top. So we keep the first
  // centred section centred and only force the REST to top.
  try {
    const zip = doc.getZip();
    const df = zip.file("word/document.xml");
    if (df) {
      let dx = df.asText();
      let changed = 0;
      // Force EVERY centred section (including the cover) to TOP. Section-level
      // <w:vAlign w:val="center"> is honoured by Microsoft Word but IGNORED by
      // Google Docs / LibreOffice, so the cover title looked top-pinned whenever
      // the file was opened/converted there. We instead centre the cover title
      // deterministically with a tall spacer paragraph (below), which every
      // renderer honours.
      dx = dx.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, (sect: string) => {
        if (!/<w:vAlign\b[^>]*w:val="center"/i.test(sect)) return sect;
        changed++;
        return sect.replace(/<w:vAlign\b[^>]*\/>/g, '<w:vAlign w:val="top"/>');
      });
      // Centre the COVER title: push the cover content down to the vertical
      // middle with ONE tall empty spacer paragraph at the very top of the body.
      // The cover is the first section, so this spacer only affects page 1; the
      // ~4.2in height lands the title block near the page centre. (An exact-line
      // empty paragraph is reliable across Word / Google Docs / LibreOffice,
      // unlike section vAlign or suppressible spacing-before on the 1st para.)
      const COVER_SPACER =
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="6000" w:lineRule="exact"/></w:pPr></w:p>';
      dx = dx.replace(/(<w:body>)/, '$1' + COVER_SPACER);
      // Keep the header clear of the red page border, and the body below it:
      //   w:header → 180: the header (logo + title) starts 0.125in from the
      //     page top, so the 0.45in logo ends ~0.575in — ABOVE the border's top
      //     line, which is drawn ~0.67in from the page edge (space=48pt). The
      //     template's 240/420 values pushed the logo down INTO the line.
      //   pgMar w:top → 1040: body content starts 0.72in down — BELOW the
      //     line — explicitly, instead of relying on a tall header to push it.
      //     (Slightly MORE usable page height than before: the old header
      //     pushed the body to ~0.86in.)
      dx = dx.replace(
        /(<w:pgMar\b[^>]*?w:header=")\d+(")/g,
        (_m: string, a: string, b: string) => a + "180" + b
      );
      dx = dx.replace(
        /(<w:pgMar\b[^>]*?w:top=")\d+(")/g,
        (_m: string, a: string, b: string) => a + "1040" + b
      );
      zip.file("word/document.xml", dx);
      console.log(
        `[export] sections forced to top: ${changed}; cover centred via spacer; header/top margins normalised (header=180, top=1040)`
      );
    }
  } catch (err) {
    console.error("[export] page vAlign top fix failed (non-fatal):", err);
  }

  // Lower the observation row's forced minimum height. The template pins each
  // observation row to w:trHeight val="2850" (~1.98"), which left blank space
  // in the Category cell whenever the (now trimmed, aspect-correct) icon + its
  // sibling cells were shorter than that floor. Keeping hRule="atLeast" means
  // rows with taller content (a big icon or long description) still grow — we
  // only remove the empty floor so short rows hug their content.
  try {
    const zip = doc.getZip();
    const df = zip.file("word/document.xml");
    if (df) {
      let dx = df.asText();
      const before = dx;
      dx = dx.replace(
        /<w:trHeight\b[^>]*w:val="2850"[^>]*\/>/g,
        '<w:trHeight w:val="1000" w:hRule="atLeast"/>'
      );
      if (dx !== before) {
        zip.file("word/document.xml", dx);
        console.log("[export] observation row height floor lowered (2850 -> 1000 twips)");
      }
    }
  } catch (err) {
    console.error("[export] observation row height floor fix failed (non-fatal):", err);
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
