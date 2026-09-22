// Crops that Word and PowerPoint apply to pictures, applied for real.
//
// Cropping a picture in Word or PowerPoint does not cut it. The full original
// stays inside the file (word/media/image3.jpeg) and the crop is only a note
// beside it saying how much of each edge to hide:
//
//   <pic:blipFill>
//     <a:blip r:embed="rId7"/>
//     <a:srcRect l="12000" t="5000" r="8000" b="20000"/>   ← the crop
//   </pic:blipFill>
//
// Anything that reads the media file directly — as the importer does — gets
// the uncropped original. So people cropped their photos in Word, imported
// the file, and saw the portal show the whole picture again.
//
// This reads those notes and cuts the image to match, before it is uploaded.
//
// (Word's own "Compress Pictures → Delete cropped areas" does the cut inside
// the file, which is why a compressed document already imported correctly.)

import sharp from "sharp";

/** Fraction of each edge to remove, 0–1. */
export interface Crop {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/** srcRect values are thousandths of a percent: 100000 = the whole side. */
const SRC_RECT_UNIT = 100000;

function edge(value: string | undefined, unit: number): number {
  if (!value) return 0;
  const n = Number(value);
  // Negative means the picture was *extended* past its edge (padding), which
  // is nothing to cut. Clamp so a malformed value can never exceed the side.
  return Number.isFinite(n) ? Math.min(Math.max(n / unit, 0), 1) : 0;
}

/**
 * Legacy VML (older .doc files re-saved as .docx) writes crops as either a
 * plain fraction ("0.1") or a count of 1/65536ths with an "f" suffix ("6554f").
 */
function vmlEdge(value: string | undefined): number {
  if (!value) return 0;
  const v = value.trim();
  const n = v.endsWith("f") ? Number(v.slice(0, -1)) / 65536 : Number(v);
  return Number.isFinite(n) ? Math.min(Math.max(n, 0), 1) : 0;
}

function attr(tag: string, name: string): string | undefined {
  return tag.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1];
}

function isCropped(c: Crop): boolean {
  return c.left > 0 || c.top > 0 || c.right > 0 || c.bottom > 0;
}

/**
 * Every picture's crop in one XML part, keyed by the zip path of the image
 * file. `resolve` turns a relationship id into that zip path, the way each
 * parser already does for photos.
 *
 * If the same image file is placed twice with different crops, the first
 * placement wins — survey reports use each photo once, so this does not
 * arise in practice, and the alternative would mean changing how every
 * parser identifies a photo.
 */
export function collectCrops(
  xml: string,
  resolve: (relId: string) => string | null
): Map<string, Crop> {
  const crops = new Map<string, Crop>();
  const record = (relId: string | undefined, crop: Crop) => {
    if (!relId || !isCropped(crop)) return;
    const path = resolve(relId);
    if (path && !crops.has(path)) crops.set(path, crop);
  };

  // DrawingML — Word 2007+ and PowerPoint. The crop sits in the same
  // blipFill as the image reference, in either order.
  for (const m of Array.from(xml.matchAll(/<(?:pic|p|a):blipFill\b[\s\S]*?<\/(?:pic|p|a):blipFill>/g))) {
    const block = m[0];
    const relId = block.match(/<a:blip\b[^>]*\br:embed="([^"]+)"/)?.[1];
    const rect = block.match(/<a:srcRect\b[^>]*\/?>/)?.[0];
    if (!rect) continue;
    record(relId, {
      left: edge(attr(rect, "l"), SRC_RECT_UNIT),
      top: edge(attr(rect, "t"), SRC_RECT_UNIT),
      right: edge(attr(rect, "r"), SRC_RECT_UNIT),
      bottom: edge(attr(rect, "b"), SRC_RECT_UNIT),
    });
  }

  // VML — pictures in documents that began life as .doc files.
  for (const m of Array.from(xml.matchAll(/<v:imagedata\b[^>]*\/?>/g))) {
    const tag = m[0];
    record(attr(tag, "r:id"), {
      left: vmlEdge(attr(tag, "cropleft")),
      top: vmlEdge(attr(tag, "croptop")),
      right: vmlEdge(attr(tag, "cropright")),
      bottom: vmlEdge(attr(tag, "cropbottom")),
    });
  }

  return crops;
}

/**
 * Cut an image down to its crop.
 *
 * The photo is turned upright first. Phone pictures are often stored
 * sideways with a flag saying "rotate me", and Word shows them upright — so
 * the crop was drawn on the upright picture and has to be applied to it.
 *
 * Returns the original bytes untouched if anything about the image is
 * unexpected: a picture that arrives uncropped is far better than one that
 * does not arrive at all.
 */
export async function applyCrop(bytes: Buffer, crop: Crop): Promise<Buffer> {
  if (!isCropped(crop)) return bytes;
  try {
    const upright = await sharp(bytes).rotate().toBuffer({ resolveWithObject: true });
    const { width, height } = upright.info;
    if (!width || !height) return bytes;

    const left = Math.round(width * crop.left);
    const top = Math.round(height * crop.top);
    const w = width - left - Math.round(width * crop.right);
    const h = height - top - Math.round(height * crop.bottom);
    // A crop that removes everything is a malformed file, not a request.
    if (w < 1 || h < 1) return bytes;

    return await sharp(upright.data)
      .extract({ left, top, width: w, height: h })
      .toFormat(upright.info.format as keyof sharp.FormatEnum)
      .toBuffer();
  } catch (err) {
    console.warn("[import] could not apply picture crop, keeping original:", (err as Error).message);
    return bytes;
  }
}

/**
 * The importer's photo reader, cropping as it goes. Shared by the Word and
 * PowerPoint parsers so all three read pictures the same way.
 */
export function makePhotoGetter(
  read: (zipPath: string) => Buffer | null,
  crops: Map<string, Crop>
): (zipPath: string) => Promise<Buffer | null> {
  return async (zipPath) => {
    const bytes = read(zipPath);
    if (!bytes) return null;
    const crop = crops.get(zipPath);
    return crop ? applyCrop(bytes, crop) : bytes;
  };
}
