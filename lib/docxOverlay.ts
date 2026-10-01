// Drawings that were made ON TOP of a photo in Word, rendered onto the photo.
//
// Surveyors often open a photo in Word and draw over it — a yellow arrow
// pointing at an obstruction, a box around a signboard, a "H 6.5m" label. Those
// marks are NOT part of the photo file. Word stores them as separate vector
// shapes (<wps:wsp> connectors / lines / rectangles / text boxes) floating over
// the picture. The importer reads only the picture's bytes, so every one of
// those drawings used to vanish on import — the recurring "I'm not getting the
// drawings on the image" complaint.
//
// This module finds those shapes, works out where each one sits ON the photo,
// and paints them back on (as an SVG composited onto the image) before upload.
//
// How the position is recovered (it is recoverable because the photo and its
// shapes live in the SAME paragraph and share one coordinate system):
//   - A shape's <wp:anchor> gives an absolute offset in EMU from the column
//     (horizontal) and the paragraph (vertical).
//   - The photo is usually inline and centred, so its left edge is
//     (contentWidth - pictureWidth) / 2 and its top edge is the paragraph top.
//     An anchored photo gives its own offset directly.
//   - Subtract the photo's origin from the shape's offset and divide by the
//     photo's size → the shape's position as a fraction (0..1) of the photo,
//     which is resolution-independent and survives any later resize/crop.
//
// Best-effort by design: anything unexpected is skipped (and on any error the
// original photo is returned untouched), because a photo without one arrow is
// far better than a failed import.

import sharp from "sharp";

const EMU_PER_INCH = 914400;
const EMU_PER_TWIP = 635; // 1 twip = 1/1440 inch, 1 inch = 914400 EMU
const EMU_PER_POINT = 12700;

/** A single mark to paint on a photo. All coordinates are fractions (0..1). */
export type OverlayShape =
  | {
      kind: "line";
      x1: number;
      y1: number;
      x2: number;
      y2: number;
      color: string; // CSS color, e.g. "#FFFF00"
      widthFrac: number; // line width as a fraction of the photo width
      head: boolean; // arrowhead at (x1,y1)
      tail: boolean; // arrowhead at (x2,y2)
    }
  | {
      kind: "rect";
      x: number;
      y: number;
      w: number;
      h: number;
      color: string;
      widthFrac: number;
    }
  | {
      kind: "text";
      x: number;
      y: number;
      w: number;
      h: number;
      text: string;
      color: string;
      sizeFrac: number; // font size as a fraction of the photo height
    };

function num(s: string | undefined, fallback = 0): number {
  if (s == null) return fallback;
  const n = Number(s);
  return Number.isFinite(n) ? n : fallback;
}

/** Usable content width of the page in EMU (page width minus side margins). */
function contentWidthEmu(xml: string): number | null {
  const pg = xml.match(/<w:pgSz\b[^>]*\bw:w="(\d+)"/);
  const mar = xml.match(/<w:pgMar\b[^>]*\/>/);
  if (!pg) return null;
  const w = num(pg[1]);
  const left = num(mar?.[0].match(/w:left="(\d+)"/)?.[1]);
  const right = num(mar?.[0].match(/w:right="(\d+)"/)?.[1]);
  const twips = w - left - right;
  return twips > 0 ? twips * EMU_PER_TWIP : null;
}

/** A drawing's displayed size (EMU). */
function extentEmu(drawing: string): { cx: number; cy: number } | null {
  const m = drawing.match(/<wp:extent\b[^>]*\bcx="(\d+)"\s+cy="(\d+)"/);
  if (!m) return null;
  return { cx: num(m[1]), cy: num(m[2]) };
}

/**
 * A floating drawing's top-left offset (EMU) and size, resolving BOTH ways Word
 * places a shape: an explicit <wp:posOffset>, or a <wp:align> (left/center/
 * right) which we turn into an offset using the container width / shape size.
 * Returns null if the shape can't be placed (e.g. a vertical align we can't
 * resolve without the container height).
 */
function shapeBox(
  drawing: string,
  containerWidth: number | null
): { h: number; v: number; cx: number; cy: number } | null {
  const ext = extentEmu(drawing);
  if (!ext) return null;

  const hOff = drawing.match(/<wp:positionH\b[^>]*>\s*<wp:posOffset>(-?\d+)<\/wp:posOffset>/);
  let h: number | null;
  if (hOff) {
    h = num(hOff[1]);
  } else {
    const hAlign = drawing.match(/<wp:positionH\b[^>]*>\s*<wp:align>(\w+)<\/wp:align>/)?.[1];
    const cw = containerWidth;
    if (hAlign && cw != null) {
      h = hAlign === "center" ? (cw - ext.cx) / 2 : hAlign === "right" || hAlign === "outside" ? cw - ext.cx : 0;
    } else h = null;
  }

  const vOff = drawing.match(/<wp:positionV\b[^>]*>\s*<wp:posOffset>(-?\d+)<\/wp:posOffset>/);
  let v: number | null;
  if (vOff) {
    v = num(vOff[1]);
  } else {
    // A vertical align needs the paragraph/container height we don't have;
    // only "top" is safe to resolve (0).
    const vAlign = drawing.match(/<wp:positionV\b[^>]*>\s*<wp:align>(\w+)<\/wp:align>/)?.[1];
    v = vAlign === "top" ? 0 : null;
  }

  if (h == null || v == null) return null;
  return { h, v, cx: ext.cx, cy: ext.cy };
}

/** First srgbClr inside a given block, as "#RRGGBB". */
function colorIn(block: string, fallback: string): string {
  const m = block.match(/srgbClr val="([0-9A-Fa-f]{6})"/);
  return m ? `#${m[1]}` : fallback;
}

/** A table cell's width in EMU (dxa = twips), or null. */
function cellWidthEmu(tc: string): number | null {
  const m = tc.match(/<w:tcW\b[^>]*\bw:w="(\d+)"/);
  if (!m) return null;
  const twips = num(m[1]);
  return twips > 0 ? twips * EMU_PER_TWIP : null;
}

/**
 * Process every photo paragraph in one layout scope (a table cell, or the page
 * body) using that scope's width as the container the photo is centred in.
 * `handled` stops a photo already placed in a more specific scope (its cell)
 * from being reprocessed by the broader body scope.
 */
function processScope(
  scopeMasked: string,
  allDrawings: string[],
  containerWidth: number | null,
  resolve: (relId: string) => string | null,
  handled: Set<string>,
  out: Map<string, OverlayShape[]>
): void {
  // scopeMasked has every <w:drawing> replaced by a \x00D<i>\x00 placeholder, so
  // the <w:p> inside a drawing's text box can't corrupt the paragraph split.
  const paragraphs = scopeMasked.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g) || [];
  for (const para of paragraphs) {
    const drawings = Array.from(para.matchAll(/\u0000D(\d+)\u0000/g))
      .map((m) => allDrawings[Number(m[1])])
      .filter(Boolean);
    if (drawings.length < 2) continue; // need a photo plus at least one mark

    const picDraws = drawings.filter((d) => d.includes("<pic:pic"));
    if (picDraws.length !== 1) continue; // only simple one-photo paragraphs
    const picDraw = picDraws[0];

    const picExt = extentEmu(picDraw);
    const embed = picDraw.match(/r:embed="([^"]+)"/)?.[1];
    if (!picExt || !embed || picExt.cx <= 0 || picExt.cy <= 0) continue;
    const zipPath = resolve(embed);
    if (!zipPath || handled.has(zipPath)) continue;
    handled.add(zipPath);

    // Photo origin (EMU) in the column/paragraph coordinate system.
    const picAnchor = picDraw.includes("<wp:anchor") ? shapeBox(picDraw, containerWidth) : null;
    let picLeft: number;
    let picTop: number;
    if (picAnchor) {
      picLeft = picAnchor.h;
      picTop = picAnchor.v;
    } else {
      // Inline: horizontally placed per the paragraph's justification within its
      // container (a table cell, or the page body), vertically at paragraph top.
      const jc = para.match(/<w:jc w:val="([^"]+)"/)?.[1] || "left";
      const cw = containerWidth ?? picExt.cx;
      picLeft = jc === "center" ? (cw - picExt.cx) / 2 : jc === "right" ? cw - picExt.cx : 0;
      picTop = 0;
    }

    {
      const shapes: OverlayShape[] = [];
      for (const dr of drawings) {
        if (dr === picDraw || dr.includes("<pic:pic")) continue;
        const box = shapeBox(dr, containerWidth);
        if (!box) continue;
        const off = { h: box.h, v: box.v };
        const ext = { cx: box.cx, cy: box.cy };

        // Fractions of the photo.
        const fx = (off.h - picLeft) / picExt.cx;
        const fy = (off.v - picTop) / picExt.cy;
        const fw = ext.cx / picExt.cx;
        const fh = ext.cy / picExt.cy;
        // Drop shapes that fall well outside the photo (stray page graphics).
        if (fx > 1.3 || fy > 1.3 || fx + fw < -0.3 || fy + fh < -0.3) continue;

        const prst = dr.match(/prst="([^"]+)"/)?.[1] || "";
        const lnBlock = dr.match(/<a:ln\b[^>]*>[\s\S]*?<\/a:ln>/)?.[0] || dr.match(/<a:ln\b[^>]*\/>/)?.[0] || "";
        const lnW = num(lnBlock.match(/\bw="(\d+)"/)?.[1], 38100); // default ~3pt

        if (prst === "straightConnector1" || prst === "line") {
          const flipH = /flipH="1"/.test(dr);
          const flipV = /flipV="1"/.test(dr);
          const head = /<a:headEnd\b[^>]*type="(?!none)[^"]+"/.test(dr);
          const tail = /<a:tailEnd\b[^>]*type="(?!none)[^"]+"/.test(dr);
          shapes.push({
            kind: "line",
            x1: flipH ? fx + fw : fx,
            y1: flipV ? fy + fh : fy,
            x2: flipH ? fx : fx + fw,
            y2: flipV ? fy : fy + fh,
            color: colorIn(lnBlock, "#FFFF00"),
            widthFrac: lnW / picExt.cx,
            head,
            tail,
          });
        } else if (/<a:t>/.test(dr)) {
          const text = (dr.match(/<a:t>([\s\S]*?)<\/a:t>/g) || [])
            .map((s) => s.replace(/<\/?a:t>/g, ""))
            .join(" ")
            .replace(/\s+/g, " ")
            .trim();
          if (!text) continue;
          const sz = num(dr.match(/<a:rPr\b[^>]*\bsz="(\d+)"/)?.[1], 1800); // hundredths of a point
          shapes.push({
            kind: "text",
            x: fx,
            y: fy,
            w: fw,
            h: fh,
            text,
            // text runs carry their own fill; fall back to red (common for labels)
            color: colorIn(dr, "#FF0000"),
            sizeFrac: ((sz / 100) * EMU_PER_POINT) / picExt.cy,
          });
        } else if (prst === "rect") {
          // Only a rectangle with a visible outline is a drawn box; a plain
          // picture frame has no separate stroke here.
          const strokeClr = lnBlock.match(/srgbClr val="([0-9A-Fa-f]{6})"/)?.[1];
          if (!strokeClr || /<a:noFill\s*\/>/.test(lnBlock)) continue;
          shapes.push({
            kind: "rect",
            x: fx,
            y: fy,
            w: fw,
            h: fh,
            color: `#${strokeClr}`,
            widthFrac: lnW / picExt.cx,
          });
        }
      }

      if (shapes.length) out.set(zipPath, (out.get(zipPath) || []).concat(shapes));
    }
  }
}

/**
 * Read every drawn-on-photo shape in the document, keyed by the photo's zip
 * path (the same key the crop reader and the photo getter use). `resolve` turns
 * a relationship id into that zip path.
 *
 * Table cells are processed first with their own width as the container, then
 * the page body — so a photo that fills a narrow cell and one centred on the
 * full page both get their shapes placed correctly.
 */
export function collectOverlays(
  xml: string,
  resolve: (relId: string) => string | null
): Map<string, OverlayShape[]> {
  const out = new Map<string, OverlayShape[]>();
  try {
    // Pull every drawing out and leave a placeholder behind, so paragraph (and
    // cell) splitting never trips over the <w:p> that lives inside a text box.
    const drawings: string[] = [];
    const masked = xml.replace(/<w:drawing>[\s\S]*?<\/w:drawing>/g, (m) => {
      const i = drawings.length;
      drawings.push(m);
      return `\u0000D${i}\u0000`;
    });

    const contentW = contentWidthEmu(xml);
    const handled = new Set<string>();
    // Table cells first (their own width is the container), then the page body.
    const cells = masked.match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || [];
    for (const tc of cells) processScope(tc, drawings, cellWidthEmu(tc) ?? contentW, resolve, handled, out);
    processScope(masked, drawings, contentW, resolve, handled, out);
  } catch (err) {
    console.warn("[import] collectOverlays failed, skipping overlays:", (err as Error).message);
  }
  return out;
}

function xmlEsc(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Build the SVG that paints the shapes onto a W×H photo. */
function buildSvg(shapes: OverlayShape[], W: number, H: number): string {
  const markers: string[] = [];
  const body: string[] = [];

  shapes.forEach((s, i) => {
    if (s.kind === "line") {
      const sw = Math.max(1, s.widthFrac * W);
      const x1 = s.x1 * W;
      const y1 = s.y1 * H;
      const x2 = s.x2 * W;
      const y2 = s.y2 * H;
      const m = Math.max(6, sw * 3.2); // arrowhead size
      const id = `a${i}`;
      if (s.head || s.tail) {
        markers.push(
          `<marker id="${id}" markerWidth="${m}" markerHeight="${m}" refX="${(m * 0.9).toFixed(2)}" refY="${(m / 2).toFixed(2)}" orient="auto-start-reverse" markerUnits="userSpaceOnUse"><path d="M0,0 L${m.toFixed(2)},${(m / 2).toFixed(2)} L0,${m.toFixed(2)} z" fill="${s.color}"/></marker>`
        );
      }
      const ms = s.head ? ` marker-start="url(#${id})"` : "";
      const me = s.tail ? ` marker-end="url(#${id})"` : "";
      body.push(
        `<line x1="${x1.toFixed(1)}" y1="${y1.toFixed(1)}" x2="${x2.toFixed(1)}" y2="${y2.toFixed(1)}" stroke="${s.color}" stroke-width="${sw.toFixed(1)}" stroke-linecap="round"${ms}${me}/>`
      );
    } else if (s.kind === "rect") {
      const sw = Math.max(1, s.widthFrac * W);
      body.push(
        `<rect x="${(s.x * W).toFixed(1)}" y="${(s.y * H).toFixed(1)}" width="${(s.w * W).toFixed(1)}" height="${(s.h * H).toFixed(1)}" fill="none" stroke="${s.color}" stroke-width="${sw.toFixed(1)}"/>`
      );
    } else {
      const fs = Math.max(8, s.sizeFrac * H);
      const cx = (s.x + s.w / 2) * W;
      const cy = (s.y + s.h / 2) * H;
      body.push(
        `<text x="${cx.toFixed(1)}" y="${cy.toFixed(1)}" font-family="Arial, sans-serif" font-size="${fs.toFixed(1)}" font-weight="700" fill="${s.color}" text-anchor="middle" dominant-baseline="central" stroke="#ffffff" stroke-width="${(fs * 0.06).toFixed(2)}" paint-order="stroke">${xmlEsc(s.text)}</text>`
      );
    }
  });

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}"><defs>${markers.join("")}</defs>${body.join("")}</svg>`;
}

/**
 * Paint the shapes onto the photo. Returns the original bytes untouched if
 * there is nothing to draw or anything goes wrong.
 */
export async function applyOverlay(bytes: Buffer, shapes: OverlayShape[] | undefined): Promise<Buffer> {
  if (!shapes || !shapes.length) return bytes;
  try {
    const base = sharp(bytes).rotate(); // honour EXIF orientation, like the crop path
    const meta = await base.metadata();
    const W = meta.width;
    const H = meta.height;
    if (!W || !H) return bytes;

    const svg = buildSvg(shapes, W, H);
    const flattened = await base.toBuffer(); // upright pixels to composite onto
    const out = sharp(flattened).composite([{ input: Buffer.from(svg), top: 0, left: 0 }]);

    const fmt = String(meta.format || "").toLowerCase();
    if (fmt === "png") return await out.png().toBuffer();
    if (fmt === "webp") return await out.webp({ quality: 95 }).toBuffer();
    // keep drawn lines/text crisp (no chroma blur), like the crop re-encode
    return await out.jpeg({ quality: 95, chromaSubsampling: "4:4:4", mozjpeg: true }).toBuffer();
  } catch (err) {
    console.warn("[import] could not paint overlay, keeping original:", (err as Error).message);
    return bytes;
  }
}
