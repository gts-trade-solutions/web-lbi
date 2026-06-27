/**
 * Build templates/reena-all-template.docx from templates/_source/reena-all-source.docx.
 *
 * Reads the original Reena-ALL DOCX (the user's reference file) and injects
 * docxtemplater placeholders, then strips the original embedded images so the
 * runtime image module can replace them.
 *
 * Output keeps every Reena style/section/header/footer untouched - only text
 * runs and image placeholders are mutated. Re-run this script when the
 * Reena reference file changes.
 */

const path = require("path");
const fs = require("fs");
const PizZip = require("pizzip");

const ROOT = path.resolve(__dirname, "..");
const SRC = path.join(ROOT, "templates", "_source", "reena-all-source.docx");
const OUT = path.join(ROOT, "templates", "reena-all-template.docx");

function readSourceZip() {
  const buf = fs.readFileSync(SRC);
  return new PizZip(buf);
}

function getXml(zip, name) {
  const entry = zip.file(name);
  if (!entry) throw new Error(`Template entry missing: ${name}`);
  return entry.asText();
}

function setXml(zip, name, content) {
  zip.file(name, content);
}

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Replace exact text inside <w:t> nodes only.
 */
function replaceVisibleText(xml, exact, replacement) {
  const re = new RegExp(`(<w:t[^>]*>)${escapeRegExp(exact)}(</w:t>)`, "g");
  return xml.replace(re, `$1${replacement}$2`);
}

/**
 * Walk backwards from `from` and return the index of the most recent occurrence
 * of `needle`. Returns -1 if not found.
 */
function lastIndexBefore(haystack, needle, from) {
  return haystack.lastIndexOf(needle, from);
}

/**
 * Find the matching closing tag of a balanced XML element starting at openIdx
 * (which must point at the start of an open tag like "<w:tbl>"). Returns the
 * index immediately AFTER the matching close tag.
 */
function findMatchingClose(xml, openIdx, openTag, closeTag) {
  let depth = 0;
  let i = openIdx;
  while (i < xml.length) {
    const nextOpen = xml.indexOf(openTag, i);
    const nextClose = xml.indexOf(closeTag, i);
    if (nextClose === -1) return -1;
    if (nextOpen !== -1 && nextOpen < nextClose) {
      depth++;
      i = nextOpen + openTag.length;
    } else {
      depth--;
      if (depth === 0) return nextClose + closeTag.length;
      i = nextClose + closeTag.length;
    }
  }
  return -1;
}

function transformDocumentXml(xml) {
  // ----- 1) Replace the nested CATEGORY table that wraps an icon + label.
  // The original cell contains <w:tbl>...icon + Take Diversion...</w:tbl><w:p/>.
  // We collapse it to a single paragraph holding {category}.
  {
    const takeIdx = xml.indexOf("Take Diversion");
    if (takeIdx === -1) throw new Error("Could not find 'Take Diversion' marker.");
    const innerTblOpen = lastIndexBefore(xml, "<w:tbl>", takeIdx);
    if (innerTblOpen === -1) throw new Error("Could not find inner <w:tbl> for category cell.");
    const innerTblEnd = findMatchingClose(xml, innerTblOpen, "<w:tbl>", "</w:tbl>");
    if (innerTblEnd === -1) throw new Error("Could not match inner </w:tbl>.");
    // The category cell ends the nested table with a trailing <w:p/>. Consume it too.
    let cutEnd = innerTblEnd;
    if (xml.startsWith("<w:p/>", cutEnd)) cutEnd += "<w:p/>".length;
    const categoryParagraph =
      `<w:p><w:pPr><w:spacing w:before="28" w:after="28" w:line="320"/><w:jc w:val="center"/></w:pPr>` +
      `<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman"/>` +
      `<w:b/><w:bCs/><w:color w:val="163A2A"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>` +
      `<w:t xml:space="preserve">{category}</w:t></w:r></w:p>`;
    xml = xml.slice(0, innerTblOpen) + categoryParagraph + xml.slice(cutEnd);
  }

  // ----- 2) Cover page text and project-name substitutions.
  xml = replaceVisibleText(xml, "REENA SUMMARY REPORT", "{projectNameUpper} SUMMARY REPORT");
  xml = replaceVisibleText(xml, "REENA", "{projectNameUpper}");

  // ----- 3) Objective text (first "test") and conclusion text (second "test").
  let firstTestReplaced = false;
  let secondTestReplaced = false;
  xml = xml.replace(/(<w:t[^>]*>)test(<\/w:t>)/g, (match, open, close) => {
    if (!firstTestReplaced) {
      firstTestReplaced = true;
      return `${open}{objective}${close}`;
    }
    if (!secondTestReplaced) {
      secondTestReplaced = true;
      return `${open}{conclusion}${close}`;
    }
    return match;
  });

  // ----- 4) Observation row data placeholders (categorical text already done).
  xml = replaceVisibleText(xml, "N13 0.478", "{gpsLat}");
  xml = replaceVisibleText(xml, "E80 11.820", "{gpsLon}");
  xml = replaceVisibleText(xml, "0.0000", "{km}");
  xml = replaceVisibleText(
    xml,
    "CMWSSB Division 201, Alandur, St. Thomas Mount Cantonment, Tamil Nadu",
    "{location}"
  );
  xml = replaceVisibleText(xml, "Normal pass", "{remarks}");

  // The OBSERVATION cell has the dash em-dash run "—". Replace just the first
  // occurrence (the OBSERVATION column). The REMARKS column already became
  // {remarks} in the prior call, so the only "—" left is the one in OBSERVATION.
  // Note: it could appear elsewhere in the doc, so target only the first.
  xml = xml.replace(/(<w:t[^>]*>)—(<\/w:t>)/, `$1{observation}$2`);

  // ----- 5) Replace the route-map drawing with {%routeMap}.
  {
    const re = /<w:drawing>(?:(?!<\/w:drawing>)[\s\S])*?cx="8886825"[\s\S]*?<\/w:drawing>/;
    if (!re.test(xml)) throw new Error("Route-map drawing not located.");
    xml = xml.replace(re, `<w:t xml:space="preserve">{%routeMap}</w:t>`);
  }

  // ----- 6) Replace the GA-drawing image with {%gaDrawing}.
  {
    const re = /<w:drawing>(?:(?!<\/w:drawing>)[\s\S])*?cx="8258175"[\s\S]*?<\/w:drawing>/;
    if (!re.test(xml)) throw new Error("GA drawing not located.");
    xml = xml.replace(re, `<w:t xml:space="preserve">{%gaDrawing}</w:t>`);
  }

  // ----- 7) Replace the per-observation photo with {%photo}, and add a sibling
  // fallback paragraph that prints "Photo not available." when the photo is
  // empty. The photoFallback value is set by the helper: "" when a photo is
  // present (so the paragraph renders blank) and "Photo not available." when
  // the photo is missing. Combined with {%photo}, exactly one of the two
  // renders visibly per observation.
  {
    const re = /<w:drawing>(?:(?!<\/w:drawing>)[\s\S])*?cx="7115175"[\s\S]*?<\/w:drawing>/;
    if (!re.test(xml)) throw new Error("Observation photo not located.");
    xml = xml.replace(re, `<w:t xml:space="preserve">{%photo}</w:t>`);

    // Insert the fallback text paragraph immediately after the photo paragraph.
    const photoTagIdx = xml.indexOf("{%photo}");
    const photoParaCloseIdx = xml.indexOf("</w:p>", photoTagIdx);
    if (photoParaCloseIdx === -1) throw new Error("Cannot find photo paragraph close.");
    const insertAt = photoParaCloseIdx + "</w:p>".length;
    const fallbackPara =
      `<w:p><w:pPr><w:spacing w:before="60" w:after="60"/><w:jc w:val="center"/></w:pPr>` +
      `<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman"/>` +
      `<w:i/><w:iCs/><w:color w:val="666666"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>` +
      `<w:t xml:space="preserve">{photoFallback}</w:t></w:r></w:p>`;
    xml = xml.slice(0, insertAt) + fallbackPara + xml.slice(insertAt);
  }

  // ----- 8) Wrap the observation block (table + spacer + photo paragraph)
  // with {#observations}/{/observations}. This relies on docxtemplater's
  // paragraphLoop option to repeat the wrapped content per item.
  {
    const obsHeaderIdx = xml.indexOf("GPS LOCATION");
    if (obsHeaderIdx === -1) throw new Error("GPS LOCATION marker not found.");
    const tableOpen = lastIndexBefore(xml, "<w:tbl>", obsHeaderIdx);
    if (tableOpen === -1) throw new Error("Outer observation <w:tbl> not found.");
    const tableEnd = findMatchingClose(xml, tableOpen, "<w:tbl>", "</w:tbl>");
    if (tableEnd === -1) throw new Error("Outer observation </w:tbl> not matched.");

    // The block ends at the close of the photoFallback paragraph (which sits
    // immediately after the photo paragraph - both must be inside the loop).
    const fallbackTagIdx = xml.indexOf("{photoFallback}", tableEnd);
    if (fallbackTagIdx === -1) throw new Error("{photoFallback} placeholder missing.");
    const fallbackParaClose = xml.indexOf("</w:p>", fallbackTagIdx);
    if (fallbackParaClose === -1) throw new Error("Cannot close photoFallback paragraph.");
    const afterPhotoPara = fallbackParaClose + "</w:p>".length;

    const openLoopPara =
      `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>` +
      `<w:r><w:t xml:space="preserve">{#observations}</w:t></w:r></w:p>`;
    const closeLoopPara =
      `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>` +
      `<w:r><w:t xml:space="preserve">{/observations}</w:t></w:r></w:p>`;

    xml =
      xml.slice(0, tableOpen) +
      openLoopPara +
      xml.slice(tableOpen, afterPhotoPara) +
      closeLoopPara +
      xml.slice(afterPhotoPara);
  }

  return xml;
}

function transformHeader(xml) {
  return replaceVisibleText(xml, "REENA", "{projectNameUpper}");
}

function transformFooter1(xml) {
  return xml.replace(/Dated\s+\d{2}-\d{2}-\d{4}/g, "Dated {dateDash}");
}

function transformFooter2to5(xml) {
  return xml.replace(/Date\s*:\s*\d{2}\.\d{2}\.\d{4}/g, "Date : {dateDot}");
}

/**
 * Strip the bulky embedded images from word/media. They will be replaced at
 * render time by the docxtemplater image module which generates fresh image
 * relationships and media entries. We replace each existing media file with a
 * 1x1 placeholder PNG so the committed template remains under a few hundred KB.
 */
function stripEmbeddedMedia(zip) {
  const placeholderPng = Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgAAIAAAUAAen63NgAAAAASUVORK5CYII=",
    "base64"
  );
  const files = zip.file(/^word\/media\//);
  for (const f of files) {
    zip.file(f.name, placeholderPng);
  }
}

function main() {
  const zip = readSourceZip();

  setXml(zip, "word/document.xml", transformDocumentXml(getXml(zip, "word/document.xml")));

  for (const name of [
    "word/header1.xml",
    "word/header2.xml",
    "word/header3.xml",
    "word/header4.xml",
  ]) {
    const entry = zip.file(name);
    if (entry) setXml(zip, name, transformHeader(entry.asText()));
  }

  const f1 = zip.file("word/footer1.xml");
  if (f1) setXml(zip, "word/footer1.xml", transformFooter1(f1.asText()));
  for (const name of [
    "word/footer2.xml",
    "word/footer3.xml",
    "word/footer4.xml",
    "word/footer5.xml",
  ]) {
    const entry = zip.file(name);
    if (entry) setXml(zip, name, transformFooter2to5(entry.asText()));
  }

  stripEmbeddedMedia(zip);

  const outBuf = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
  fs.writeFileSync(OUT, outBuf);
  console.log(`Wrote ${OUT} (${outBuf.length} bytes).`);
}

main();
