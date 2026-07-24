/**
 * Injects a custom Office theme into a rendered DOCX so the six route-difficulty
 * colours appear in Word's colour picker ("Theme Colors" row).
 *
 * Word populates the colour picker's top "Theme Colors" row from the theme's
 * <a:clrScheme> accent1..accent6. The export template ships WITHOUT a theme
 * part, so Word falls back to the built-in Office theme. We add a complete,
 * valid theme part, wire it up via [Content_Types].xml + document.xml.rels, and
 * map accent1..6 to our 6 colours (3 bright header fills + 3 faint body tints):
 *
 *   accent1 = 56BE9F  GREEN  (header)        accent2 = F0F9F4  GREEN  (body tint)
 *   accent3 = EABD0D  YELLOW (header)        accent4 = FFFBE6  YELLOW (body tint)
 *   accent5 = F05052  RED    (header)        accent6 = FDF2F2  RED    (body tint)
 *
 * The fontScheme / fmtScheme are the standard Office defaults so the theme is a
 * valid OOXML part (an incomplete theme triggers Word's "repair" prompt).
 */

// The 6 colours, in picker order, grouped header+body per difficulty.
export const DIFFICULTY_PICKER_COLORS = {
  accent1: "56BE9F", // green header
  accent2: "F0F9F4", // green body
  accent3: "EABD0D", // yellow header
  accent4: "FFFBE6", // yellow body
  accent5: "F05052", // red header
  accent6: "FDF2F2", // red body
} as const;

const THEME_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Route Difficulty Theme">' +
    "<a:themeElements>" +
      '<a:clrScheme name="Route Difficulty">' +
        '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
        '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
        '<a:dk2><a:srgbClr val="1F2937"/></a:dk2>' +
        '<a:lt2><a:srgbClr val="EEECE1"/></a:lt2>' +
        `<a:accent1><a:srgbClr val="${DIFFICULTY_PICKER_COLORS.accent1}"/></a:accent1>` +
        `<a:accent2><a:srgbClr val="${DIFFICULTY_PICKER_COLORS.accent2}"/></a:accent2>` +
        `<a:accent3><a:srgbClr val="${DIFFICULTY_PICKER_COLORS.accent3}"/></a:accent3>` +
        `<a:accent4><a:srgbClr val="${DIFFICULTY_PICKER_COLORS.accent4}"/></a:accent4>` +
        `<a:accent5><a:srgbClr val="${DIFFICULTY_PICKER_COLORS.accent5}"/></a:accent5>` +
        `<a:accent6><a:srgbClr val="${DIFFICULTY_PICKER_COLORS.accent6}"/></a:accent6>` +
        '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>' +
        '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
      "</a:clrScheme>" +
      '<a:fontScheme name="Office">' +
        '<a:majorFont><a:latin typeface="Neue Haas Grotesk Text Pro"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
        '<a:minorFont><a:latin typeface="Neue Haas Grotesk Text Pro"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
      "</a:fontScheme>" +
      '<a:fmtScheme name="Office">' +
        "<a:fillStyleLst>" +
          '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
          '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>' +
          '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>' +
        "</a:fillStyleLst>" +
        "<a:lnStyleLst>" +
          '<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>' +
          '<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>' +
          '<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>' +
        "</a:lnStyleLst>" +
        "<a:effectStyleLst>" +
          "<a:effectStyle><a:effectLst/></a:effectStyle>" +
          "<a:effectStyle><a:effectLst/></a:effectStyle>" +
          '<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>' +
        "</a:effectStyleLst>" +
        "<a:bgFillStyleLst>" +
          '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
          '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>' +
          '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>' +
        "</a:bgFillStyleLst>" +
      "</a:fmtScheme>" +
    "</a:themeElements>" +
    "<a:objectDefaults/>" +
    "<a:extraClrSchemeLst/>" +
  "</a:theme>";

const THEME_PART_NAME = "/word/theme/theme1.xml";
const THEME_PART_PATH = "word/theme/theme1.xml";
const THEME_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.theme+xml";
const THEME_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

type ZipLike = {
  file(path: string): { asText(): string } | null;
  file(path: string, data: string): unknown;
};

/**
 * Add the custom theme part to a rendered DOCX zip and wire it up. Idempotent
 * and defensive: if a theme already exists, or any required part is missing, it
 * leaves the document untouched rather than risk producing an invalid file.
 *
 * Returns true if the theme was injected, false if it was skipped.
 */
export function injectDifficultyThemeColors(zip: ZipLike): boolean {
  // If the document already has a theme part, do not add a second one.
  if (zip.file(THEME_PART_PATH)) return false;

  const ctFile = zip.file("[Content_Types].xml");
  const relsFile = zip.file("word/_rels/document.xml.rels");
  if (!ctFile || !relsFile) return false;

  let contentTypes = ctFile.asText();
  let rels = relsFile.asText();

  // 1) Content type override for the theme part.
  if (!contentTypes.includes(THEME_PART_NAME)) {
    const override = `<Override PartName="${THEME_PART_NAME}" ContentType="${THEME_CONTENT_TYPE}"/>`;
    const closeIdx = contentTypes.lastIndexOf("</Types>");
    if (closeIdx === -1) return false;
    contentTypes = contentTypes.slice(0, closeIdx) + override + contentTypes.slice(closeIdx);
  }

  // 2) Relationship from the document part to the theme part. Pick an rId that
  //    does not collide with any existing one.
  if (!rels.includes(THEME_REL_TYPE)) {
    const existingIds = Array.from(rels.matchAll(/Id="rId(\d+)"/g)).map((m) => Number(m[1]));
    const nextId = (existingIds.length ? Math.max(...existingIds) : 0) + 1;
    const relId = `rId${nextId}`;
    const rel = `<Relationship Id="${relId}" Type="${THEME_REL_TYPE}" Target="theme/theme1.xml"/>`;
    const closeIdx = rels.lastIndexOf("</Relationships>");
    if (closeIdx === -1) return false;
    rels = rels.slice(0, closeIdx) + rel + rels.slice(closeIdx);
  }

  // 3) Write everything back (theme part + the two wiring files).
  zip.file(THEME_PART_PATH, THEME_XML);
  zip.file("[Content_Types].xml", contentTypes);
  zip.file("word/_rels/document.xml.rels", rels);
  return true;
}
