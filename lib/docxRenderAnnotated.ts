// Render photos EXACTLY as Word shows them — with every drawn mark baked in.
//
// The vector approach (lib/docxOverlay.ts) re-draws straight arrows/boxes/text
// itself, so it can never match Word pixel-for-pixel and cannot reproduce
// freehand / curved / grouped drawings at all. The client wants the photos to
// look *identical to the Word file*. The only reliable way to do that is to let
// a real Word-compatible renderer draw them — here, LibreOffice (headless) on
// the server, the same tool the earlier proof-of-concept verified.
//
// For each photo that has a drawn shape over it, we build a tiny .docx holding
// just that photo's paragraph (the inline picture + its anchored shapes), sized
// to the same width it had in the report, convert it to PDF with LibreOffice,
// rasterise page 1 with Ghostscript, and trim the surrounding whitespace. The
// result is the photo with the arrows/lines/labels rendered exactly as Word
// draws them.
//
// Entirely best-effort: if LibreOffice/Ghostscript aren't present (e.g. local
// dev on Windows) or anything fails, it returns an empty map and the importer
// falls back to lib/docxOverlay.ts. It never throws into the import.

import fs from "fs";
import os from "os";
import path from "path";
import { execFileSync } from "child_process";
import PizZip from "pizzip";
import sharp from "sharp";

// EMU → twips (page/section sizes are in twips; 1 twip = 635 EMU).
const EMU_PER_TWIP = 635;

/** Locate the LibreOffice binary, or null if it isn't installed. */
function findSoffice(): string | null {
  for (const c of ["/usr/bin/soffice", "/usr/bin/libreoffice", "soffice", "libreoffice"]) {
    try {
      execFileSync(c, ["--version"], { stdio: "ignore", timeout: 15000 });
      return c;
    } catch {
      /* try next */
    }
  }
  return null;
}

function hasGhostscript(): boolean {
  try {
    execFileSync("gs", ["--version"], { stdio: "ignore", timeout: 15000 });
    return true;
  } catch {
    return false;
  }
}

/** Does this drawing (not the picture) look like a real drawn mark? */
function isDrawnShape(dr: string): boolean {
  if (dr.includes("<pic:pic")) return false;
  return (
    /<a:(?:tail|head)End\s+type="(?!none)/.test(dr) ||
    /<a:custGeom\b/.test(dr) || // freeform / scribbles
    /prst="(?:straightConnector1|line|bentConnector\d|curvedConnector\d)"/.test(dr) ||
    /<wps:wsp\b/.test(dr) ||
    /<v:(?:shape|line|group|curve|polyline)\b/.test(dr)
  );
}

function contentWidthTwips(xml: string): number {
  const pg = xml.match(/<w:pgSz\b[^>]*\bw:w="(\d+)"/);
  const mar = xml.match(/<w:pgMar\b[^>]*\/>/);
  const w = Number(pg?.[1] || 0);
  const l = Number(mar?.[0].match(/w:left="(\d+)"/)?.[1] || 0);
  const r = Number(mar?.[0].match(/w:right="(\d+)"/)?.[1] || 0);
  const cw = w - l - r;
  return cw > 0 ? cw : 12240 - 1440; // A4-ish fallback
}

type Target = { zipPath: string; paraXml: string; pageWidthTwips: number; aspect: number };

/** relationship id → media zip path ("word/media/imageN.ext"), order-independent. */
function buildResolver(relsXml: string): (relId: string) => string | null {
  const map: Record<string, string> = {};
  for (const m of Array.from(relsXml.matchAll(/<Relationship\b[^>]*>/g))) {
    const tag = m[0];
    const id = tag.match(/\bId="([^"]+)"/)?.[1];
    const target = tag.match(/\bTarget="([^"]+)"/)?.[1];
    if (id && target) map[id] = target;
  }
  return (relId) => {
    const t = map[relId];
    return t && /media\//.test(t) ? "word/" + t.replace(/^\/*/, "") : null;
  };
}

/**
 * Render every drawn-on photo in the document with LibreOffice.
 * Returns a map of photo zip path → rendered PNG bytes (the photo with its
 * drawings baked in). Empty when rendering isn't available.
 */
export async function renderAnnotatedPhotos(buffer: Buffer): Promise<Map<string, Buffer>> {
  const out = new Map<string, Buffer>();

  const soffice = findSoffice();
  if (!soffice) return out; // not on this host → caller falls back
  const gs = hasGhostscript();

  let xml: string;
  let docOpen: string;
  let resolve: (relId: string) => string | null;
  try {
    const zip = new PizZip(buffer);
    xml = zip.file("word/document.xml")?.asText() || "";
    docOpen = xml.match(/<w:document[^>]*>/)?.[0] || "";
    resolve = buildResolver(zip.file("word/_rels/document.xml.rels")?.asText() || "");
    if (!xml || !docOpen) return out;
  } catch {
    return out;
  }

  // Mask drawings so the <w:p> inside a text box can't corrupt paragraph/cell
  // splitting; keep the originals to rebuild the real paragraph.
  const drawings: string[] = [];
  const masked = xml.replace(/<w:drawing>[\s\S]*?<\/w:drawing>/g, (m) => {
    drawings.push(m);
    return `\u0000D${drawings.length - 1}\u0000`;
  });

  const contentW = contentWidthTwips(xml);
  const seen = new Set<string>();
  const targets: Target[] = [];

  const scan = (scopeMasked: string, widthTwips: number) => {
    const paras = scopeMasked.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g) || [];
    for (const pm of paras) {
      const refs = Array.from(pm.matchAll(/\u0000D(\d+)\u0000/g)).map((m) => Number(m[1]));
      if (refs.length < 2) continue;
      const real = refs.map((i) => drawings[i]).filter(Boolean);
      const pics = real.filter((d) => d.includes("<pic:pic"));
      if (pics.length !== 1) continue;
      if (!real.some(isDrawnShape)) continue;
      const embed = pics[0].match(/r:embed="([^"]+)"/)?.[1];
      const zipPath = embed ? resolve(embed) : null;
      if (!zipPath || seen.has(zipPath)) continue;
      const ext = pics[0].match(/<wp:extent\b[^>]*\bcx="(\d+)"\s+cy="(\d+)"/);
      const cx = Number(ext?.[1] || 0);
      const cy = Number(ext?.[2] || 0);
      if (!cx || !cy) continue;
      seen.add(zipPath);
      const paraXml = pm.replace(/\u0000D(\d+)\u0000/g, (_, i) => drawings[Number(i)] || "");
      targets.push({ zipPath, paraXml, pageWidthTwips: Math.max(2000, widthTwips), aspect: cx / cy });
    }
  };

  // Table cells first (their own width), then the page body.
  for (const tc of masked.match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || []) {
    const w = Number(tc.match(/<w:tcW\b[^>]*\bw:w="(\d+)"/)?.[1] || 0) || contentW;
    scan(tc, w);
  }
  scan(masked, contentW);

  if (!targets.length) return out;

  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "lbi-annot-"));
  try {
    // One tiny .docx per drawn photo: the original zip (media + rels intact) with
    // its body replaced by just this paragraph, on a page as wide as the photo's
    // container so the shapes keep their positions.
    const miniPaths: string[] = [];
    targets.forEach((t, idx) => {
      try {
        const z = new PizZip(buffer);
        const sect =
          `<w:sectPr><w:pgSz w:w="${t.pageWidthTwips}" w:h="31680"/>` +
          `<w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>`;
        const body = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${docOpen}<w:body>${t.paraXml}${sect}</w:body></w:document>`;
        z.file("word/document.xml", body);
        const p = path.join(tmp, `m${idx}.docx`);
        fs.writeFileSync(p, z.generate({ type: "nodebuffer", compression: "DEFLATE" }));
        miniPaths.push(p);
      } catch {
        miniPaths.push(""); // keep index alignment
      }
    });

    const valid = miniPaths.filter(Boolean);
    if (!valid.length) return out;

    // One LibreOffice call converts them all (amortises its slow start-up).
    try {
      execFileSync(
        soffice,
        [
          "--headless",
          `-env:UserInstallation=file://${path.join(tmp, "lo")}`,
          "--convert-to",
          "pdf",
          "--outdir",
          tmp,
          ...valid,
        ],
        { stdio: "ignore", timeout: 240000 }
      );
    } catch {
      return out; // conversion failed wholesale → fall back
    }

    for (let idx = 0; idx < targets.length; idx++) {
      if (!miniPaths[idx]) continue;
      const pdf = path.join(tmp, `m${idx}.pdf`);
      if (!fs.existsSync(pdf)) continue;
      try {
        let rawPng = "";
        if (gs) {
          rawPng = path.join(tmp, `r${idx}.png`);
          execFileSync(
            "gs",
            [
              "-dNOPAUSE",
              "-dBATCH",
              "-sDEVICE=png16m",
              "-r150",
              "-dFirstPage=1",
              "-dLastPage=1",
              `-sOutputFile=${rawPng}`,
              pdf,
            ],
            { stdio: "ignore", timeout: 120000 }
          );
        } else {
          // No Ghostscript: let LibreOffice render the mini-doc straight to PNG.
          execFileSync(
            soffice,
            ["--headless", `-env:UserInstallation=file://${path.join(tmp, "lo")}`, "--convert-to", "png", "--outdir", tmp, miniPaths[idx]],
            { stdio: "ignore", timeout: 120000 }
          );
          rawPng = path.join(tmp, `m${idx}.png`);
        }
        if (!rawPng || !fs.existsSync(rawPng)) continue;
        const trimmed = await sharp(rawPng).trim({ threshold: 12 }).png().toBuffer();
        if (!trimmed || trimmed.length <= 1000) continue;
        // Safety: only trust the render if its shape matches the photo's (within
        // 35%). A wildly different aspect means trim grabbed the wrong region or
        // the render failed — fall back to the extracted photo instead.
        const meta = await sharp(trimmed).metadata();
        if (!meta.width || !meta.height) continue;
        const renderedAspect = meta.width / meta.height;
        const want = targets[idx].aspect;
        if (want > 0 && Math.abs(renderedAspect - want) / want > 0.35) {
          console.warn(
            `[annot-render] ${targets[idx].zipPath}: aspect ${renderedAspect.toFixed(2)} vs expected ${want.toFixed(2)} — skipping, using fallback.`
          );
          continue;
        }
        out.set(targets[idx].zipPath, trimmed);
      } catch {
        /* skip this one; importer falls back for it */
      }
    }
  } finally {
    try {
      fs.rmSync(tmp, { recursive: true, force: true });
    } catch {
      /* ignore */
    }
  }

  return out;
}

export const __emuPerTwip = EMU_PER_TWIP; // (kept for tests / future sizing use)
