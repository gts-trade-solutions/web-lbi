// Login-only download of the Tracker Android app (APK).
//
// The APK lives on the server's disk OUTSIDE public/ (default
// .uploads/app/lbi-tracker.apk, override with APP_APK_PATH), so Next.js never
// serves it statically — this route is the only way to get it.
//
//   POST  (logged in)       -> { url, fileName, size } — url carries a ticket
//                              valid for 2 minutes
//   GET   ?ticket=...       -> streams the APK
//   GET   (Bearer / cookie) -> streams the APK (API clients)
//
// A ticket is needed because a browser download is a plain navigation: it can't
// send the Authorization header, and the Secure auth cookie isn't stored while
// the site runs on plain HTTP.
import fs from "fs";
import path from "path";
import { Readable } from "stream";
import { getAuthUser, requireAuth, signDownloadTicket, verifyDownloadTicket } from "../../../../lib/auth";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const PURPOSE = "app-apk";
const APK_MIME = "application/vnd.android.package-archive";

function apkPath() {
  const configured = (process.env.APP_APK_PATH || "").trim();
  return configured
    ? path.resolve(configured)
    : path.join(process.cwd(), ".uploads", "app", "lbi-tracker.apk");
}

function apkStat() {
  try {
    const st = fs.statSync(apkPath());
    return st.isFile() ? st : null;
  } catch {
    return null;
  }
}

function downloadFileName(mtime: Date) {
  return `lbi-tracker-${mtime.toISOString().slice(0, 10)}.apk`;
}

export async function POST(request: Request) {
  try {
    const user = requireAuth(request);
    const st = apkStat();
    if (!st) {
      return Response.json({ error: "The app file has not been uploaded to the server yet." }, { status: 404 });
    }
    const ticket = signDownloadTicket(String(user.id), PURPOSE);
    return Response.json({
      url: `/api/app/download?ticket=${encodeURIComponent(ticket)}`,
      fileName: downloadFileName(st.mtime),
      size: st.size,
      updatedAt: st.mtime.toISOString(),
    });
  } catch (err: any) {
    if (err?.message === "Unauthorized") {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/app/download] POST error:", err);
    return Response.json({ error: "Could not prepare the download" }, { status: 500 });
  }
}

export async function GET(request: Request) {
  const ticket = new URL(request.url).searchParams.get("ticket") || "";
  const allowed = ticket ? !!verifyDownloadTicket(ticket, PURPOSE) : !!getAuthUser(request);
  if (!allowed) {
    return new Response("Unauthorized — sign in to the web app and use Download app.", {
      status: 401,
      headers: { "Content-Type": "text/plain; charset=utf-8", "Cache-Control": "no-store" },
    });
  }

  const st = apkStat();
  if (!st) return new Response("App file not found", { status: 404 });

  const body = Readable.toWeb(fs.createReadStream(apkPath())) as unknown as ReadableStream;
  return new Response(body, {
    headers: {
      "Content-Type": APK_MIME,
      "Content-Length": String(st.size),
      "Content-Disposition": `attachment; filename="${downloadFileName(st.mtime)}"`,
      "Cache-Control": "private, no-store",
      "X-Content-Type-Options": "nosniff",
    },
  });
}
