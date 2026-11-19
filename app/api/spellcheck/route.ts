import { requireAuth } from "../../../lib/auth";
import { checkText, type Misspelling } from "../../../lib/spellcheck";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

function isUnauthorizedError(e: unknown) {
  return (e as { message?: string })?.message === "Unauthorized";
}

/**
 * POST { fields: { <name>: text, ... }, ignore?: string[] }
 * →    { results: { <name>: Misspelling[] }, total: number }
 *
 * Runs the self-contained English spell checker (no external service) on each
 * field's text. `ignore` holds the user's taught words (sent from the client).
 */
export async function POST(request: Request) {
  try {
    requireAuth(request);
    const body = await request.json().catch(() => ({} as any));
    const fields = (body?.fields && typeof body.fields === "object") ? body.fields : {};
    const ignore: string[] = Array.isArray(body?.ignore)
      ? body.ignore.map((w: unknown) => String(w || "")).filter(Boolean)
      : [];

    const results: Record<string, Misspelling[]> = {};
    let total = 0;
    for (const [name, value] of Object.entries(fields)) {
      const miss = checkText(String(value ?? ""), ignore);
      results[name] = miss;
      total += miss.length;
    }

    return Response.json({ results, total });
  } catch (error) {
    if (isUnauthorizedError(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/spellcheck] POST error:", error);
    return Response.json({ error: "Spell check failed" }, { status: 500 });
  }
}
