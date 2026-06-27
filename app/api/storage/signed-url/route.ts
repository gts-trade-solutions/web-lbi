import { requireAuth } from "../../../../lib/auth";
import { getPublicS3Url, getReadSignedUrl, safeObjectKey } from "../../../../lib/s3";

export const runtime = "nodejs";

export async function POST(request: Request) {
  try {
    requireAuth(request);
    const body = await request.json().catch(() => ({}));
    const path = safeObjectKey(String(body?.path || ""));
    const expiresIn = Number(body?.expiresIn || 600);

    // Generate signed URL; if it fails, gracefully fallback to public URL.
    let url = "";
    try {
      url = await getReadSignedUrl(path, Number.isFinite(expiresIn) ? expiresIn : 600);
    } catch {
      url = getPublicS3Url(path);
    }

    return Response.json({ url, path });
  } catch (e: unknown) {
    const err = e as { message?: string };
    if (err?.message === "Unauthorized") {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    return Response.json({ error: "Failed to generate signed URL" }, { status: 500 });
  }
}
