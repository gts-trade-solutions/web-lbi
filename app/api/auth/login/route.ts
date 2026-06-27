import bcrypt from "bcryptjs";
import pool from "../../../../lib/db";
import { makeAuthCookieHeader, signToken } from "../../../../lib/auth";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type LoginBody = {
  email?: string;
  password?: string;
};

type UserRow = {
  id: string;
  email: string;
  name: string | null;
  role: string | null;
  password_hash: string | null;
  password: string | null;
};

export async function POST(request: Request) {
  try {
    const { email: rawEmail, password } = (await request.json().catch(() => ({}))) as LoginBody;
    const email = String(rawEmail || "").trim().toLowerCase();
    const plainPassword = String(password || "");

    console.log("[auth/login] MYSQL env check:", {
      host: Boolean(process.env.MYSQL_HOST),
      port: Boolean(process.env.MYSQL_PORT),
      user: Boolean(process.env.MYSQL_USER),
      database: Boolean(process.env.MYSQL_DATABASE),
      passwordExists: Boolean(process.env.MYSQL_PASSWORD),
    });

    if (!email || !plainPassword) {
      return Response.json({ error: "Email and password are required" }, { status: 400 });
    }

    const [columnRows] = await pool.query("SHOW COLUMNS FROM users");
    const columnSet = new Set(
      (Array.isArray(columnRows) ? columnRows : []).map((row) =>
        String((row as Record<string, unknown>)?.Field || "").toLowerCase()
      )
    );

    const hasRole = columnSet.has("role");
    const hasPasswordHash = columnSet.has("password_hash");
    const hasPassword = columnSet.has("password");

    const [rows] = await pool.query(
      `SELECT id, email, name, ${
        hasRole ? "role" : "NULL AS role"
      }, ${
        hasPasswordHash ? "password_hash" : "NULL AS password_hash"
      }, ${
        hasPassword ? "password" : "NULL AS password"
      } FROM users WHERE email = ? LIMIT 1`,
      [email]
    );

    const user = Array.isArray(rows) && rows.length > 0 ? (rows[0] as UserRow) : null;
    console.log("[auth/login] user found:", Boolean(user));

    if (!user) {
      return Response.json({ error: "Invalid email or password" }, { status: 401 });
    }

    const hashToCompare =
      (hasPasswordHash ? user.password_hash : null) ||
      (hasPassword ? user.password : null) ||
      null;
    console.log("[auth/login] password hash exists:", Boolean(hashToCompare));

    if (!hashToCompare) {
      return Response.json({ error: "Invalid email or password" }, { status: 401 });
    }

    const passwordOk = await bcrypt.compare(plainPassword, hashToCompare);
    console.log("[auth/login] bcrypt compare:", passwordOk ? "passed" : "failed");

    if (!passwordOk) {
      return Response.json({ error: "Invalid email or password" }, { status: 401 });
    }

    const safeUser = {
      id: user.id,
      email: user.email,
      name: user.name,
      role: user.role || "user",
    };

    const token = signToken(safeUser);
    const cookieHeader = makeAuthCookieHeader(token);

    return new Response(
      JSON.stringify({
        user: safeUser,
        token,
      }),
      {
        status: 200,
        headers: {
          "Content-Type": "application/json",
          "Set-Cookie": cookieHeader,
        },
      }
    );
  } catch (error) {
    console.error("[auth/login] error:", error);
    return Response.json({ error: "Login server error" }, { status: 500 });
  }
}
