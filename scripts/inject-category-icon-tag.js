/* eslint-disable @typescript-eslint/no-var-requires */
// One-shot template patch: inject {%categoryIcon} (gated by hasCategoryIcon)
// into the CATEGORY cell of templates/reena-all-template.docx, immediately
// before the existing {category} text. Idempotent — re-running is a no-op.
//
// Usage: node scripts/inject-category-icon-tag.js
const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");

const TEMPLATE = path.join(__dirname, "..", "templates", "reena-all-template.docx");
const BACKUP = TEMPLATE + ".bak";

// The exact run sequence that opens the {category} paragraph in the current
// template. Captured by inspecting word/document.xml. If any of these bytes
// drift we must re-capture rather than blindly inject.
const CATEGORY_PARAGRAPH_OPEN =
  '<w:p><w:pPr><w:spacing w:before="28" w:after="28" w:line="320"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:bCs/><w:color w:val="163A2A"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t xml:space="preserve">{category}</w:t></w:r></w:p>';

// Three paragraphs — each tag in its own paragraph so paragraphLoop drops the
// section markers cleanly and leaves just the image paragraph (or nothing) in
// the rendered output. Centered, zero spacing, single-line at a tiny font so
// the image paragraph adds minimal vertical room next to the category text.
function buildIconBlock() {
  const compactPPr =
    '<w:pPr>' +
    '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>' +
    '<w:jc w:val="center"/>' +
    '</w:pPr>';
  const compactRPr = '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>';
  const wrap = (text) =>
    "<w:p>" + compactPPr +
    "<w:r>" + compactRPr + '<w:t xml:space="preserve">' + text + "</w:t></w:r>" +
    "</w:p>";
  return wrap("{#hasCategoryIcon}") + wrap("{%categoryIconKey}") + wrap("{/hasCategoryIcon}");
}

function main() {
  if (!fs.existsSync(TEMPLATE)) {
    console.error("Template not found:", TEMPLATE);
    process.exit(1);
  }

  const buf = fs.readFileSync(TEMPLATE);
  const zip = new PizZip(buf);
  const docFile = zip.file("word/document.xml");
  if (!docFile) {
    console.error("word/document.xml not found inside the docx");
    process.exit(1);
  }

  let xml = docFile.asText();

  // If a previous run inserted {%categoryIcon} (the direct-object variant
  // that crashed the image module), upgrade it in place to the imageMap-key
  // variant {%categoryIconKey} which routes through the same code path as
  // {%routeMap} / {%gaDrawing}.
  if (xml.includes("{%categoryIcon}") && !xml.includes("{%categoryIconKey}")) {
    if (!fs.existsSync(BACKUP)) {
      fs.copyFileSync(TEMPLATE, BACKUP);
      console.log("Wrote backup:", BACKUP);
    }
    xml = xml.replace(/\{%categoryIcon\}/g, "{%categoryIconKey}");
    zip.file("word/document.xml", xml);
    const out = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
    fs.writeFileSync(TEMPLATE, out);
    console.log("Upgraded {%categoryIcon} → {%categoryIconKey} in template.");
    return;
  }

  if (xml.includes("{%categoryIconKey}")) {
    // Per design pivot: full category drawings are too wide for the
    // observation CATEGORY cell. Remove the entire `{#hasCategoryIcon} …
    // {%categoryIconKey} … {/hasCategoryIcon}` block from the cell so it shows
    // text-only ({category}). The drawings will only appear on the dedicated
    // CATEGORY COUNT SUMMARY section via {%categorySummaryIcon}.
    const blockRe =
      /<w:p[^>]*>(?:(?!<\/w:p>).)*\{#hasCategoryIcon\}<\/w:t><\/w:r><\/w:p><w:p[^>]*>(?:(?!<\/w:p>).)*\{%categoryIconKey\}<\/w:t><\/w:r><\/w:p><w:p[^>]*>(?:(?!<\/w:p>).)*\{\/hasCategoryIcon\}<\/w:t><\/w:r><\/w:p>/;
    if (!blockRe.test(xml)) {
      console.log(
        "Template contains {%categoryIconKey} but block signature drifted — leaving as-is."
      );
      return;
    }
    if (!fs.existsSync(BACKUP)) {
      fs.copyFileSync(TEMPLATE, BACKUP);
      console.log("Wrote backup:", BACKUP);
    }
    xml = xml.replace(blockRe, "");
    zip.file("word/document.xml", xml);
    const out = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
    fs.writeFileSync(TEMPLATE, out);
    console.log(
      "Removed observation-cell category-icon block. CATEGORY cell now shows {category} text only."
    );
    return;
  }

  if (!xml.includes(CATEGORY_PARAGRAPH_OPEN)) {
    console.error(
      "Could not locate the {category} paragraph signature.\n" +
        "The template XML has drifted from the captured snapshot. Re-inspect word/document.xml."
    );
    process.exit(2);
  }

  // Backup once.
  if (!fs.existsSync(BACKUP)) {
    fs.copyFileSync(TEMPLATE, BACKUP);
    console.log("Wrote backup:", BACKUP);
  }

  const before = xml.length;
  xml = xml.replace(CATEGORY_PARAGRAPH_OPEN, buildIconBlock() + CATEGORY_PARAGRAPH_OPEN);
  console.log("Patched document.xml:", before, "→", xml.length, "bytes");

  zip.file("word/document.xml", xml);
  const out = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
  fs.writeFileSync(TEMPLATE, out);
  console.log("Wrote patched template:", TEMPLATE, "(", out.length, "bytes )");
}

main();
