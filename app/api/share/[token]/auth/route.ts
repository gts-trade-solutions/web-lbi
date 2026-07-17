/* eslint-disable @typescript-eslint/no-explicit-any */
// Public: exchange the share password for a short-lived session token.
// No login required — this is the door a client walks through.
import bcrypt from "bcryptjs";
import pool from "../../../../../lib/db";
import { getShareByToken, signShareSession } from "../../../../../lib/shareLink";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type Ctx = { params: { token: string } };

export async function POST(request: Request, context: Ctx) {
  try {
    const token = String(context.params?.token || "").trim();
    if (!token) return Response.json({ error: "Invalid link" }, { status: 400 });

    const body = await request.json().catch(() => ({} as any));
    const password = String(body?.password || "");

    const share = await getShareByToken(token);
    // Same generic message whether the link is unknown or the password is
    // wrong, so a probe can't tell valid tokens from invalid ones.
    const deny = () =>
      Response.json({ error: "Incorrect password, or this link is no longer active." }, { status: 401 });

    if (!share) {
      // Constant-ish time: still run one bcrypt so timing doesn't leak
      // "token exists" vs "token doesn't".
      await bcrypt.compare(password, "$2a$10$0000000000000000000000000000000000000000000000000000").catch(() => false);
      return deny();
    }

    const ok = await bcrypt.compare(password, share.password_hash).catch(() => false);
    if (!ok) return deny();

    // Friendly project name for the client's header.
    let projectName = "";
    try {
      const [rows] = await pool.query("SELECT name FROM projects WHERE id = ? LIMIT 1", [
        share.project_id,
      ]);
      projectName = Array.isArray(rows) && rows.length ? String((rows[0] as any)?.name || "") : "";
    } catch {
      /* ignore */
    }

    const session = signShareSession(token, share.project_id);
    return Response.json({
      ok: true,
      session,
      projectId: share.project_id,
      projectName,
      title: share.title || "",
    });
  } catch (error) {
    console.error("[api/share/:token/auth] error:", error);
    return Response.json({ error: "Could not verify this link right now." }, { status: 500 });
  }
}
