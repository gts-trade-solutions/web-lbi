/**
 * Rename the observation-photo placeholder inside the prepared template:
 *   {%photo}        -> {%observationPhotoKey}
 *   {photoFallback} -> {photoText}
 *
 * The helper code emits BOTH `observationPhotoKey` + `photoText` AND legacy
 * `photo` + `photoFallback` fields on every observation, so this rename can
 * be re-run safely if the template is regenerated. Run with:
 *   node scripts/rename-photo-placeholder.js
 */

const path = require("path");
const fs = require("fs");
const PizZip = require("pizzip");

const ROOT = path.resolve(__dirname, "..");
const TPL = path.join(ROOT, "templates", "reena-all-template.docx");

function rename(buf) {
  const zip = new PizZip(buf);
  const entry = zip.file("word/document.xml");
  if (!entry) throw new Error("word/document.xml missing from template");
  let xml = entry.asText();
  let changed = false;
  if (xml.includes("{%photo}") && !xml.includes("{%observationPhotoKey}")) {
    xml = xml.replace(/\{%photo\}/g, "{%observationPhotoKey}");
    changed = true;
    console.log("Renamed {%photo} -> {%observationPhotoKey}");
  }
  if (xml.includes("{photoFallback}") && !xml.includes("{photoText}")) {
    xml = xml.replace(/\{photoFallback\}/g, "{photoText}");
    changed = true;
    console.log("Renamed {photoFallback} -> {photoText}");
  }
  if (!changed) {
    console.log("No changes needed - placeholders already match.");
    return null;
  }
  zip.file("word/document.xml", xml);
  return zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
}

const out = rename(fs.readFileSync(TPL));
if (out) {
  fs.writeFileSync(TPL, out);
  console.log("Wrote", TPL, "(" + out.length + " bytes).");
}
