// Same-origin image proxy. The annotation canvas needs to READ photo pixels
// (canvas.toBlob) to flatten a drawing onto a photo. Photos served from S3
// don't send CORS headers, so loading them cross-origin taints the canvas and
// export fails. Fetching them through this same-origin route removes the taint.
//
// Locked to https + AWS S3 hosts (plus the configured bucket host) so it can't
// be abused as an open proxy / SSRF vector.
export const runtime = "nodejs";
export const dynamic = "force-dynamic";

function allowedHost(host: string): boolean {
  const h = host.toLowerCase();
  // Never proxy raw IPs (blocks 169.254.169.254 metadata etc.).
  if (/^\d{1,3}(\.\d{1,3}){3}$/.test(h)) return false;
  if (h.endsWith(".amazonaws.com")) return true;
  try {
    const bucket = process.env.NEXT_PUBLIC_S3_BUCKET_URL || "";
    if (bucket) {
      const bh = new URL(bucket).host.toLowerCase();
      if (bh && h === bh) return true;
    }
  } catch {
    /* ignore malformed env */
  }
  return false;
}

export async function GET(request: Request) {
  const { searchParams } = new URL(request.url);
  const target = searchParams.get("url") || "";
  if (!target) return new Response("Missing url", { status: 400 });

  let parsed: URL;
  try {
    parsed = new URL(target);
  } catch {
    return new Response("Bad url", { status: 400 });
  }
  if (parsed.protocol !== "https:" || !allowedHost(parsed.host)) {
    return new Response("Host not allowed", { status: 403 });
  }

  try {
    const upstream = await fetch(parsed.toString(), { cache: "no-store" });
    if (!upstream.ok) return new Response("Upstream error", { status: upstream.status });
    const buf = await upstream.arrayBuffer();
    const ct = upstream.headers.get("content-type") || "image/jpeg";
    return new Response(buf, {
      status: 200,
      headers: {
        "Content-Type": ct,
        "Cache-Control": "public, max-age=3600",
      },
    });
  } catch {
    return new Response("Fetch failed", { status: 502 });
  }
}
