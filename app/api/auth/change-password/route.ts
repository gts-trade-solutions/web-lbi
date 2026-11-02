import bcrypt from "bcryptjs";
import pool from "../../../../lib/db";
import { extractTokenFromRequest, verifyToken } from "../../../../lib/auth";
import { logActivity } from "../../../../lib/activityLog";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type Body = {
  currentPassword?: string;
  newPassword?: string;
};

type UserRow = {
  id: string;
  email: string;
  name: string | null;
  password_hash: string | null;
  password: string | null;
};

export async function POST(request: Request) {
  try {
    // Must be logged in — change the password of whoever owns this token.
    const token = extractTokenFromRequest(request);
    const payload = token ? verifyToken(token) : null;
    if (!payload?.id) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }

    const { currentPassword, newPassword } = (await request
      .json()
      .catch(() => ({}))) as Body;
    const current = String(currentPassword || "");
    const next = String(newPassword || "");

    if (!current || !next) {
      return Response.json(
        { error: "Current and new password are required" },
        { status: 400 }
      );
    }
    if (next.length < 6) {
      return Response.json(
        { error: "New password must be at least 6 characters" },
        { status: 400 }
      );
    }
    if (next === current) {
      return Response.json(
        { error: "New password must be different from the current one" },
        { status: 400 }
      );
    }

    // The users table schema varies across installs: it may store the hash in
    // `password_hash`, in `password`, or both. Detect what exists (same as
    // /api/auth/login) and keep every present column in sync.
    const [columnRows] = await pool.query("SHOW COLUMNS FROM users");
    const columnSet = new Set(
      (Array.isArray(columnRows) ? columnRows : []).map((row) =>
        String((row as Record<string, unknown>)?.Field || "").toLowerCase()
      )
    );
    const hasPasswordHash = columnSet.has("password_hash");
    const hasPassword = columnSet.has("password");
    const hasUpdatedAt = columnSet.has("updated_at");

    const [rows] = await pool.query(
      `SELECT id, email, name, ${
        hasPasswordHash ? "password_hash" : "NULL AS password_hash"
      }, ${hasPassword ? "password" : "NULL AS password"} FROM users WHERE id = ? LIMIT 1`,
      [payload.id]
    );
    const user =
      Array.isArray(rows) && rows.length > 0 ? (rows[0] as UserRow) : null;
    if (!user) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }

    const existingHash =
      (hasPasswordHash ? user.password_hash : null) ||
      (hasPassword ? user.password : null) ||
      null;
    if (!existingHash) {
      return Response.json(
        { error: "No password is set for this account" },
        { status: 400 }
      );
    }

    const ok = await bcrypt.compare(current, existingHash);
    if (!ok) {
      await logActivity(request, {
        action: "password_change_failed",
        actorId: user.id,
        actorEmail: user.email,
        details: { reason: "wrong current password" },
      });
      return Response.json(
        { error: "Current password is incorrect" },
        { status: 401 }
      );
    }

    const newHash = await bcrypt.hash(next, 10);

    // Write the new hash into every password column the table has, so login
    // keeps working whichever column it reads from.
    const setParts: string[] = [];
    const args: unknown[] = [];
    if (hasPasswordHash) {
      setParts.push("password_hash = ?");
      args.push(newHash);
    }
    if (hasPassword) {
      setParts.push("password = ?");
      args.push(newHash);
    }
    if (hasUpdatedAt) setParts.push("updated_at = NOW()");
    args.push(user.id);

    await pool.query(
      `UPDATE users SET ${setParts.join(", ")} WHERE id = ?`,
      args
    );

    await logActivity(request, {
      action: "password_changed",
      actorId: user.id,
      actorEmail: user.email,
      details: { name: user.name || null },
    });

    return Response.json({ ok: true });
  } catch (error) {
    console.error("[auth/change-password] error:", error);
    return Response.json({ error: "Could not change password" }, { status: 500 });
  }
}
