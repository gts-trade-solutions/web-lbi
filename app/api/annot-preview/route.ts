// TEMPORARY debug route: stream a rendered annotation-test PNG from disk so we
// can view the LibreOffice output (Next.js doesn't serve runtime-added public/
// files). Remove after the annotation-render feature is validated.
import fs from "fs";
import path from "path";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  const url = new URL(request.url);
  const n = (url.searchParams.get("n") || "0").replace(/[^0-9]/g, "").slice(0, 4) || "0";
  // Fixed absolute dir so it works regardless of the server's working directory
  // (Next standalone runs from a different cwd than the repo root).
  const file = path.join("/tmp", "lbi-preview", `annotated_${n}.png`);
  try {
    const buf = fs.readFileSync(file);
    return new Response(new Uint8Array(buf), {
      headers: { "Content-Type": "image/png", "Cache-Control": "no-store" },
    });
  } catch {
    return new Response("not found", { status: 404 });
  }
}
