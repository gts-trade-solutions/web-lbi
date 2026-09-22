// Same-origin image proxy. The annotation canvas needs to READ photo pixels
// (canvas.toBlob) to flatten a drawing onto a photo. Photos served from S3
// don't send CORS headers, so loading them cross-origin taints the canvas and
// export fails. Fetching them through this same-origin route removes the taint.
//
// Login required. Only objects in OUR bucket are served, read with the S3 SDK
// (server credentials), so it works with a private bucket and can't be used as
// an open proxy / SSRF vector. Clients call it with fetch() + the Bearer token
// and turn the response into a blob: URL (see lib/authedImage.ts).
import { GetObjectCommand } from "@aws-sdk/client-s3";
import { requireAuth } from "../../../lib/auth";
import { getBucketName, s3Client, s3KeyFromUrl } from "../../../lib/s3";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  try {
    requireAuth(request);
  } catch {
    return new Response("Unauthorized", { status: 401 });
  }

  const { searchParams } = new URL(request.url);
  const target = searchParams.get("url") || "";
  if (!target) return new Response("Missing url", { status: 400 });

  const key = s3KeyFromUrl(target);
  if (!key) return new Response("Host not allowed", { status: 403 });

  try {
    const obj = await s3Client.send(new GetObjectCommand({ Bucket: getBucketName(), Key: key }));
    if (!obj.Body) return new Response("Not found", { status: 404 });
    return new Response(obj.Body.transformToWebStream(), {
      status: 200,
      headers: {
        "Content-Type": obj.ContentType || "image/jpeg",
        "Cache-Control": "private, max-age=3600",
      },
    });
  } catch (err: any) {
    const status = err?.$metadata?.httpStatusCode;
    if (status === 404 || err?.name === "NoSuchKey") return new Response("Not found", { status: 404 });
    console.error("[api/image-proxy] S3 read failed:", key, err?.name || err);
    return new Response("Fetch failed", { status: 502 });
  }
}
