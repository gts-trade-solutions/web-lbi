// Render a .docx to page PNGs so the report layout can be reviewed on the
// server (LibreOffice → PDF → per-page PNG via ghostscript). Writes the pages
// as annotated_0.png, annotated_1.png … into public/uploads/annot-test/, which
// the existing /api/annot-preview route already streams (?n=0, ?n=1, …).
//
// Usage (on the server):
//   node scripts/render-docx-preview.mjs "$(ls -t /tmp/lbi-exports/*.docx | head -1)" 10
//     arg1 = .docx path (default: newest export)   arg2 = how many pages
import fs from "fs";
import path from "path";
import { execFileSync } from "child_process";

const SRC = process.argv[2];
const COUNT = Math.min(Math.max(1, Number(process.argv[3] || 10)), 40);

if (!SRC || !fs.existsSync(SRC)) {
  console.error("Pass a valid .docx path. Got:", SRC);
  process.exit(1);
}

const SOFFICE =
  ["/usr/bin/soffice", "/usr/bin/libreoffice"].find((p) => fs.existsSync(p)) || "soffice";
const OUTDIR = "/tmp/lbi-preview"; // served by /api/annot-preview
fs.mkdirSync(OUTDIR, { recursive: true });

// Clear old preview pages so stale ones aren't mistaken for the new export.
for (const f of fs.readdirSync(OUTDIR)) {
  if (/^annotated_\d+\.png$/.test(f)) fs.unlinkSync(path.join(OUTDIR, f));
}

console.log("rendering", path.basename(SRC), "via", SOFFICE);
execFileSync(
  SOFFICE,
  ["--headless", "-env:UserInstallation=file:///tmp/lo_profile_exp", "--convert-to", "pdf", "--outdir", "/tmp", SRC],
  { stdio: "inherit", timeout: 240000 }
);
const pdf = path.join("/tmp", path.basename(SRC).replace(/\.docx$/i, ".pdf"));
if (!fs.existsSync(pdf)) {
  console.error("PDF not produced:", pdf);
  process.exit(1);
}

let made = 0;
for (let i = 1; i <= COUNT; i++) {
  const out = path.join(OUTDIR, `annotated_${i - 1}.png`);
  try {
    execFileSync(
      "gs",
      ["-dNOPAUSE", "-dBATCH", "-sDEVICE=png16m", "-r90", `-dFirstPage=${i}`, `-dLastPage=${i}`, `-sOutputFile=${out}`, pdf],
      { stdio: "ignore" }
    );
  } catch {
    break; // ran past the last page
  }
  if (!fs.existsSync(out) || fs.statSync(out).size < 200) {
    if (fs.existsSync(out)) fs.unlinkSync(out);
    break;
  }
  made++;
  console.log(`page ${i} -> annotated_${i - 1}.png  ${(fs.statSync(out).size / 1024).toFixed(0)}KB`);
}
console.log(`DONE — ${made} page(s). View at /api/annot-preview?n=0 … ?n=${made - 1}`);
