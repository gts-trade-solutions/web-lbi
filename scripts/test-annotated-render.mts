// Validate the EXACT (LibreOffice) photo rendering on the server, before
// relying on it in a real import.
//
// Run on the server (LibreOffice + Ghostscript must be installed there):
//   npx tsx scripts/test-annotated-render.mts "/path/to/REPORT.docx"
//
// It renders every drawn-on photo and writes them to /tmp/lbi-preview, which the
// app already serves for viewing. Open each in a browser:
//   https://lbi-web.raceinnovations.in/api/annot-preview?n=0   (then n=1, 2, …)
//
// If they look exactly like the Word file, the real import (which uses the same
// code) will produce the same result.

import fs from "fs";
import path from "path";
import { renderAnnotatedPhotos } from "../lib/docxRenderAnnotated";

const src = process.argv[2];
if (!src || !fs.existsSync(src)) {
  console.error("Pass a .docx path, e.g.:  npx tsx scripts/test-annotated-render.mts \"/tmp/report.docx\"");
  process.exit(1);
}

const buf = fs.readFileSync(src);
console.log("Rendering annotated photos with LibreOffice… (this can take a minute)");
const map = await renderAnnotatedPhotos(buf);

const dest = "/tmp/lbi-preview";
fs.mkdirSync(dest, { recursive: true });
let i = 0;
for (const [zipPath, png] of map) {
  fs.writeFileSync(path.join(dest, `annotated_${i}.png`), png);
  console.log(`  annotated_${i}.png  <-  ${zipPath}  (${(png.length / 1024).toFixed(0)} KB)`);
  i++;
}

if (!map.size) {
  console.log(
    "\nNo photos were rendered. Either this file has no drawn-on photos, or LibreOffice/Ghostscript isn't installed. Check: soffice --version ; gs --version"
  );
} else {
  console.log(
    `\nRendered ${map.size} photo(s) to ${dest}.\nView them at:  https://lbi-web.raceinnovations.in/api/annot-preview?n=0  (then n=1 … n=${map.size - 1})`
  );
}
