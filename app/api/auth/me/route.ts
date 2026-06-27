import pool from "../../../../lib/db";
import { extractTokenFromRequest, verifyToken } from "../../../../lib/auth";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type UserRow = {
  id: string;
  email: string;
  name: string | null;
  role: string | null;
};

export async function GET(request: Request) {
  try {
    const token = extractTokenFromRequest(request);
    if (!token) return Response.json({ error: "Unauthorized" }, { status: 401 });

    const payload = verifyToken(token);
    if (!payload?.id) return Response.json({ error: "Unauthorized" }, { status: 401 });

    const [columnRows] = await pool.query("SHOW COLUMNS FROM users");
    const columnSet = new Set(
      (Array.isArray(columnRows) ? columnRows : []).map((row) =>
        String((row as Record<string, unknown>)?.Field || "").toLowerCase()
      )
    );
    const hasRole = columnSet.has("role");

    const [rows] = await pool.query(
      `SELECT id, email, name, ${hasRole ? "role" : "NULL AS role"} FROM users WHERE id = ? LIMIT 1`,
      [payload.id]
    );

    const user = Array.isArray(rows) && rows.length > 0 ? (rows[0] as UserRow) : null;
    if (!user) return Response.json({ error: "Unauthorized" }, { status: 401 });

    return Response.json({
      user: {
        id: user.id,
        email: user.email,
        name: user.name,
        role: user.role || "user",
      },
    });
  } catch (error) {
    console.error("[auth/me] error:", error);
    return Response.json({ error: "Unauthorized" }, { status: 401 });
  }
}
