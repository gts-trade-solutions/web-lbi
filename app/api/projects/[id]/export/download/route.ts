import path from "path";
import fs from "fs";
import { NextResponse } from "next/server";
import { requireAuth } from "../../../../../../lib/auth";
import { TEMP_EXPORT_DIR } from "../route";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

type Ctx = { params: { id: string } };

/**
 * Validates that `requested` resolves to a regular file directly under
 * TEMP_EXPORT_DIR. Rejects path traversal (../), absolute paths, anything
 * with a directory separator, and any name that doesn't match the
 * <32-hex>.docx pattern produced by the POST handler.
 */
function resolveSafeTempPath(requested: string): string | null {
  const name = String(requested || "").trim();
  if (!name) return null;
  if (!/^[a-f0-9]{32}\.docx$/i.test(name)) return null;
  // basename() guarantees no path components survive even if the regex were
  // ever loosened. Final realpath check ensures the resolved path lives
  // strictly inside TEMP_EXPORT_DIR.
  const baseName = path.basename(name);
  const target = path.join(TEMP_EXPORT_DIR, baseName);
  const root = path.resolve(TEMP_EXPORT_DIR);
  const resolved = path.resolve(target);
  if (!resolved.startsWith(root + path.sep) && resolved !== root) {
    return null;
  }
  return resolved;
}

export async function GET(request: Request, _context: Ctx) {
  try {
    requireAuth(request);
  } catch (err) {
    if ((err as { message?: string })?.message === "Unauthorized") {
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[export download] auth error:", err);
    return NextResponse.json({ error: "Auth check failed" }, { status: 500 });
  }

  const url = new URL(request.url);
  const fileParam = url.searchParams.get("file") || "";
  const downloadName = String(url.searchParams.get("name") || "export.docx")
    .replace(/[\r\n"\\/?*<>|:]/g, "_")
    .slice(0, 200);

  const safePath = resolveSafeTempPath(fileParam);
  if (!safePath) {
    console.warn("[export download] rejected file param:", fileParam);
    return NextResponse.json({ error: "Invalid file parameter" }, { status: 400 });
  }

  let stat;
  try {
    stat = fs.statSync(safePath);
  } catch {
    return NextResponse.json({ error: "Export file not found or expired" }, { status: 404 });
  }
  if (!stat.isFile()) {
    return NextResponse.json({ error: "Export file is not a regular file" }, { status: 400 });
  }

  let buffer: Buffer;
  try {
    buffer = await fs.promises.readFile(safePath);
  } catch (err) {
    console.error("[export download] readFile failed:", err);
    return NextResponse.json({ error: "Failed to read export file" }, { status: 500 });
  }

  console.log("[export download] sending file", {
    storedName: path.basename(safePath),
    downloadName,
    bytes: buffer.length,
  });

  return new NextResponse(new Uint8Array(buffer), {
    status: 200,
    headers: {
      "Content-Type": DOCX_MIME,
      "Content-Disposition":
        `attachment; filename="${downloadName}"; filename*=UTF-8''${encodeURIComponent(downloadName)}`,
      // Content-Length intentionally omitted - same reason as the POST path:
      // any gzip layer above us would invalidate it. Chunked framing is fine.
      "Cache-Control": "no-store, no-cache, must-revalidate",
      Pragma: "no-cache",
      Expires: "0",
      "X-Content-Type-Options": "nosniff",
    },
  });
}
