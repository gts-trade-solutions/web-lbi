// Standalone proof-of-concept: render observation photos WITH their drawn
// arrows + text labels (e.g. "H-6.2m") flattened in, using LibreOffice.
//
// Why: the old import extracts the clean photo from the .docx zip, which loses
// the arrows/labels drawn on top (they're vector shapes, not part of the
// image). This script proves we can recover them by isolating each annotated
// observation table into a tiny .docx and having LibreOffice render it to an
// image (arrows + text baked in), then trimming the whitespace.
//
// Usage (on the server):
//   node scripts/render-annotated-test.mjs "/tmp/test.docx" /tmp/annot-out 3
//     arg1 = source .docx   arg2 = output dir   arg3 = how many to render
//
// Requires: libreoffice (soffice), ghostscript (gs) — both confirmed present.
import fs from "fs";
import path from "path";
import { execFileSync } from "child_process";
import PizZip from "pizzip";
import sharp from "sharp";

const SRC = process.argv[2];
const OUT = process.argv[3] || "/tmp/annot-out";
const LIMIT = Number(process.argv[4] || 3);

if (!SRC || !fs.existsSync(SRC)) {
  console.error("Pass a valid .docx path as arg1. Got:", SRC);
  process.exit(1);
}
fs.mkdirSync(OUT, { recursive: true });

// Find the soffice / libreoffice binary.
function findSoffice() {
  for (const c of ["/usr/bin/soffice", "/usr/bin/libreoffice", "soffice", "libreoffice"]) {
    try { execFileSync(c, ["--version"], { stdio: "ignore" }); return c; } catch {}
  }
  throw new Error("LibreOffice (soffice) not found");
}
const SOFFICE = findSoffice();
console.log("soffice:", SOFFICE);

const buf = fs.readFileSync(SRC);
const zip = new PizZip(buf);
const xml = zip.file("word/document.xml").asText();
const docOpen = xml.match(/<w:document[^>]*>/)[0];
const sectPr = (xml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) || []).pop() || "";

// Top-level tables (nesting-aware).
function topTables(x) {
  const t = [];
  let i = 0;
  while (true) {
    const s = x.indexOf("<w:tbl>", i);
    if (s < 0) break;
    let d = 0, e = -1;
    const re = /<w:tbl>|<\/w:tbl>/g; re.lastIndex = s;
    let m;
    while ((m = re.exec(x))) { if (m[0] === "<w:tbl>") d++; else { d--; if (!d) { e = m.index + m[0].length; break; } } }
    if (e < 0) break;
    t.push(x.slice(s, e));
    i = e;
  }
  return t;
}

// A table is an "annotated photo" table if it has an embedded image AND a drawn
// arrow/shape (floating anchor + arrowhead/custGeom/wps shape).
function isAnnotated(t) {
  const hasPic = /r:embed=/.test(t);
  const hasShape =
    /<a:(?:tail|head)End\s+type="(?!none)/.test(t) ||
    /<a:custGeom\b/.test(t) ||
    /<wps:wsp\b/.test(t) ||
    /<v:(shape|line|group|curve|polyline)\b/.test(t);
  return hasPic && hasShape;
}

const tables = topTables(xml);
const annotated = tables.filter(isAnnotated).slice(0, LIMIT);
console.log(`tables: ${tables.length}, annotated: ${tables.filter(isAnnotated).length}, rendering first ${annotated.length}`);

// Build a mini .docx per annotated table: clone the original zip, replace the
// body with just this table + the section (page size). Keeps all media + rels
// so the images + shapes resolve. (Media isn't stripped here — fine for a
// few-table test; the real import will strip to keep temp size down.)
const miniPaths = [];
annotated.forEach((table, idx) => {
  const z = new PizZip(buf); // fresh clone
  const body = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${docOpen}<w:body>${table}${sectPr}</w:body></w:document>`;
  z.file("word/document.xml", body);
  const out = z.generate({ type: "nodebuffer", compression: "DEFLATE" });
  const p = path.join(OUT, `mini_${idx}.docx`);
  fs.writeFileSync(p, out);
  miniPaths.push(p);
});

// One LibreOffice call converts ALL mini-docs to PDF (amortises startup).
console.log("converting to PDF via LibreOffice...");
execFileSync(
  SOFFICE,
  ["--headless", "-env:UserInstallation=file:///tmp/lo_profile_test",
   "--convert-to", "pdf", "--outdir", OUT, ...miniPaths],
  { stdio: "inherit", timeout: 180000 }
);

// Rasterize page 1 of each PDF, then trim whitespace → the annotated photo.
for (let idx = 0; idx < miniPaths.length; idx++) {
  const pdf = path.join(OUT, `mini_${idx}.pdf`);
  if (!fs.existsSync(pdf)) { console.log(`mini_${idx}.pdf MISSING`); continue; }
  const rawPng = path.join(OUT, `raw_${idx}.png`);
  execFileSync("gs", ["-dNOPAUSE", "-dBATCH", "-sDEVICE=png16m", "-r110",
    "-dFirstPage=1", "-dLastPage=1", `-sOutputFile=${rawPng}`, pdf], { stdio: "ignore" });
  const finalPng = path.join(OUT, `annotated_${idx}.png`);
  await sharp(rawPng).trim({ threshold: 10 }).toFile(finalPng);
  const st = fs.statSync(finalPng);
  const meta = await sharp(finalPng).metadata();
  console.log(`annotated_${idx}.png  ${meta.width}x${meta.height}  ${(st.size / 1024).toFixed(0)}KB`);
}
console.log("DONE. Annotated images are in:", OUT);
