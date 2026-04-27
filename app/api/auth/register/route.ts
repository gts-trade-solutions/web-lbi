import bcrypt from "bcryptjs";
import { v4 as uuidv4 } from "uuid";
import pool from "../../../../lib/db";
import { makeAuthCookieHeader, signToken } from "../../../../lib/auth";

export const runtime = "nodejs";

export async function POST(request: Request) {
  try {
    const body = await request.json().catch(() => ({}));
    const email = String(body?.email || "").trim().toLowerCase();
    const password = String(body?.password || "");
    const name = String(body?.name || "").trim() || null;

    if (!email || !password) {
      return Response.json({ error: "Email and password are required" }, { status: 400 });
    }

    const [existing] = await pool.query("SELECT id FROM users WHERE email = ? LIMIT 1", [email]);
    if (Array.isArray(existing) && existing.length > 0) {
      return Response.json({ error: "User already exists" }, { status: 409 });
    }

    const userId = uuidv4();
    const hash = await bcrypt.hash(password, 10);

    await pool.query(
      "INSERT INTO users (id, email, name, password_hash, created_at, updated_at) VALUES (?, ?, ?, ?, NOW(), NOW())",
      [userId, email, name, hash]
    );

    const user = { id: userId, email, name };
    const token = signToken(user);
    const cookieHeader = makeAuthCookieHeader(token);

    return new Response(JSON.stringify({ user, token }), {
      status: 201,
      headers: {
        "Content-Type": "application/json",
        "Set-Cookie": cookieHeader,
      },
    });
  } catch {
    return Response.json({ error: "Failed to register user" }, { status: 500 });
  }
}
