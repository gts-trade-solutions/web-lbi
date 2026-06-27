/**
 * Offline smoke test for the Reena DOCX template pipeline.
 *
 * Loads the prepared template, runs docxtemplater with the same image module
 * and data shape used by lib/reenaTemplateExport.ts, and writes the result to
 * smoke-output.docx. The data is fully synthetic - no MySQL or S3 access.
 *
 * Run with: node scripts/smoke-test-reena-template.js
 */

const path = require("path");
const fs = require("fs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const ImageModule = require("docxtemplater-image-module-free");

const ROOT = path.resolve(__dirname, "..");
const TEMPLATE = path.join(ROOT, "templates", "reena-all-template.docx");
const OUT = path.join(ROOT, "smoke-output.docx");

// 1x1 PNG (white pixel) used as a stand-in for every image slot.
const FAKE_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgAAIAAAUAAen63NgAAAAASUVORK5CYII=",
  "base64"
);

function formatLatLine(prefix, v) {
  const n = Number(v);
  if (!Number.isFinite(n)) return "-";
  const abs = Math.abs(n);
  const deg = Math.floor(abs);
  const minutes = ((abs - deg) * 60).toFixed(3);
  return `${prefix}${deg} ${minutes}`;
}

function gpsLat(lat) {
  const n = Number(lat);
  return Number.isFinite(n) ? formatLatLine(n >= 0 ? "N" : "S", lat) : "-";
}
function gpsLon(lng) {
  const n = Number(lng);
  return Number.isFinite(n) ? formatLatLine(n >= 0 ? "E" : "W", lng) : "-";
}

const imageMap = new Map();
imageMap.set("routeMap", FAKE_PNG);
imageMap.set("gaDrawing", FAKE_PNG);

function makeObservation(i, hasPhoto = true) {
  const lat = 13.0 + i * 0.001;
  const lng = 80.197 + i * 0.001;
  let photoKey = "";
  if (hasPhoto) {
    photoKey = `photo-${i}`;
    imageMap.set(photoKey, FAKE_PNG);
  }
  return {
    gpsLat: gpsLat(lat),
    gpsLon: gpsLon(lng),
    km: (i * 0.0055).toFixed(4),
    location: `Sample location ${i + 1}, Tamil Nadu`,
    category: i % 2 === 0 ? "Take Diversion" : "Normal",
    observation: i === 0 ? "—" : `Observation entry ${i + 1}`,
    remarks: i % 3 === 0 ? "Green" : "Normal pass",
    photo: photoKey,
    photoFallback: photoKey ? "" : "Photo not available.",
  };
}

const data = {
  projectNameUpper: "ACME ROUTE FEASIBILITY",
  objective: "Validate route feasibility for 50-foot trailer movement.",
  conclusion:
    "Based on the findings of the route feasibility study, route is feasible with the noted modifications.",
  dateDot: "25.04.2026",
  dateDash: "25-04-2026",
  routeMap: "routeMap",
  gaDrawing: "gaDrawing",
  observations: [makeObservation(0, true), makeObservation(1, false), makeObservation(2, true)],
};

function main() {
  const buf = fs.readFileSync(TEMPLATE);
  const zip = new PizZip(buf);

  const imageModule = new ImageModule({
    centered: true,
    fileType: "docx",
    getImage: (tagValue) =>
      typeof tagValue === "string" && tagValue ? imageMap.get(tagValue) || null : null,
    getSize: (_img, _value, tagName) => {
      if (tagName === "routeMap") return [933, 700];
      if (tagName === "gaDrawing") return [867, 650];
      return [747, 560];
    },
  });

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    modules: [imageModule],
    nullGetter: () => "",
  });

  doc.render(data);

  const outBuf = doc.getZip().generate({ type: "nodebuffer", compression: "DEFLATE" });
  fs.writeFileSync(OUT, outBuf);
  console.log(`Wrote ${OUT} (${outBuf.length} bytes).`);
}

main();
