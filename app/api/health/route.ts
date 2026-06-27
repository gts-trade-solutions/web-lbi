import pool from "../../../lib/db";

export const runtime = "nodejs";

export async function GET() {
  try {
    await pool.query("SELECT 1");
    return Response.json({
      ok: true,
      database: "connected",
    });
  } catch {
    return Response.json(
      {
        ok: false,
        database: "disconnected",
      },
      { status: 500 }
    );
  }
}
