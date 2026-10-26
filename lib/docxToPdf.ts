import path from "path";
import fs from "fs";
import os from "os";
import { execFileSync } from "child_process";

/** Locate the LibreOffice binary, or null if it isn't installed. */
function findSoffice(): string | null {
  for (const c of ["/usr/bin/soffice", "/usr/bin/libreoffice", "soffice", "libreoffice"]) {
    try {
      execFileSync(c, ["--version"], { stdio: "ignore", timeout: 15000 });
      return c;
    } catch {
      /* try the next candidate */
    }
  }
  return null;
}

/** True when a PDF can be produced on this host (LibreOffice is present). */
export function canConvertToPdf(): boolean {
  return findSoffice() !== null;
}

/**
 * Convert a DOCX buffer to PDF with headless LibreOffice — so the PDF is the
 * EXACT same report as the Word file, just a different format. Returns the PDF
 * buffer, or null if LibreOffice is missing or the conversion fails (the caller
 * then tells the user PDF isn't available and offers the .docx instead).
 */
export async function convertDocxToPdf(docx: Buffer): Promise<Buffer | null> {
  const soffice = findSoffice();
  if (!soffice) {
    console.warn("[docxToPdf] LibreOffice not found — PDF conversion unavailable");
    return null;
  }
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "lbi-pdf-"));
  const inPath = path.join(tmp, "report.docx");
  const outPath = path.join(tmp, "report.pdf");
  try {
    fs.writeFileSync(inPath, docx);
    execFileSync(
      soffice,
      [
        "--headless",
        // Isolated profile dir so concurrent conversions don't clash.
        `-env:UserInstallation=file://${path.join(tmp, "lo")}`,
        "--convert-to",
        "pdf",
        "--outdir",
        tmp,
        inPath,
      ],
      { stdio: "ignore", timeout: 300000 }
    );
    if (!fs.existsSync(outPath)) {
      console.error("[docxToPdf] soffice produced no PDF");
      return null;
    }
    return fs.readFileSync(outPath);
  } catch (e) {
    console.error("[docxToPdf] conversion failed:", e);
    return null;
  } finally {
    try {
      fs.rmSync(tmp, { recursive: true, force: true });
    } catch {
      /* best-effort cleanup */
    }
  }
}
