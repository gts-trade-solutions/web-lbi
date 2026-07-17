/* eslint-disable @typescript-eslint/no-explicit-any */
// Server-side proxy to OpenRouteService for road-following geometry. Keeps the
// ORS_API_KEY private (never shipped to the browser). Callable by a logged-in
// user OR by a valid share-link session, so the animated map works in both the
// authored view and a client's password-protected share.
import { requireAuth } from "../../../lib/auth";
import { verifyShareSession } from "../../../lib/shareLink";
import { fetchOrsPath } from "../../../lib/orsRoute";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request) {
  const body = await request.json().catch(() => ({} as any));

  // Gate: a logged-in user, or a valid share session for the supplied token.
  let allowed = false;
  try {
    requireAuth(request);
    allowed = true;
  } catch {
    const token = String(body?.token || "").trim();
    const session = request.headers.get("x-share-session") || "";
    if (token && verifyShareSession(session, token)) allowed = true;
  }
  if (!allowed) return Response.json({ error: "Unauthorized" }, { status: 401 });

  const key = process.env.ORS_API_KEY;
  if (!key) {
    return Response.json(
      { error: "ORS_API_KEY not configured", code: "UNCONFIGURED" },
      { status: 503 }
    );
  }

  const coordsIn = Array.isArray(body?.coordinates) ? body.coordinates : [];
  const latlngs: Array<[number, number]> = coordsIn
    .map((c: any) => [Number(c?.[0]), Number(c?.[1])] as [number, number])
    .filter((c: [number, number]) => Number.isFinite(c[0]) && Number.isFinite(c[1]));
  if (latlngs.length < 2) {
    return Response.json({ error: "need >= 2 coordinates" }, { status: 400 });
  }

  const result = await fetchOrsPath(latlngs, key);
  if (!result.ok) {
    console.warn("[api/route-directions] ORS error:", result.status, result.detail);
    return Response.json(
      { error: "Directions provider error", status: result.status, detail: result.detail },
      { status: 502 }
    );
  }
  return Response.json({ path: result.path });
}
