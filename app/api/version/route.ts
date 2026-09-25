import fs from "fs";
import path from "path";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

// The running server serves one fixed build, so read its BUILD_ID once and
// cache it. When the app is rebuilt + restarted, the new process reports a new
// id — the client compares against the id it first saw and offers to refresh.
let cachedBuildId: string | null = null;
function getBuildId(): string {
  if (cachedBuildId !== null) return cachedBuildId;
  try {
    cachedBuildId = fs
      .readFileSync(path.join(process.cwd(), ".next", "BUILD_ID"), "utf8")
      .trim();
  } catch {
    cachedBuildId = "";
  }
  return cachedBuildId;
}

export async function GET() {
  return Response.json(
    { buildId: getBuildId() },
    { headers: { "Cache-Control": "no-store, max-age=0" } }
  );
}
