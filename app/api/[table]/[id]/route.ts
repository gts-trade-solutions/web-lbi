/* eslint-disable @typescript-eslint/no-explicit-any */
import pool from "../../../../lib/db";
import { requireAuth } from "../../../../lib/auth";
import { isAllowedTable, isValidIdentifier, quoteIdentifier } from "../../../../lib/tableConfig";

export const runtime = "nodejs";

function getTableOrThrow(table: string) {
  if (!isAllowedTable(table)) throw new Error("Table is not allowed");
  return quoteIdentifier(table);
}

export async function GET(request: Request, { params }: { params: { table: string; id: string } }) {
  try {
    requireAuth(request);
    const table = getTableOrThrow(params.table);
    const [rows] = await pool.query(`SELECT * FROM ${table} WHERE id = ? LIMIT 1`, [params.id]);
    const row = Array.isArray(rows) ? (rows[0] as any) : null;
    if (!row) return Response.json({ error: "Not found" }, { status: 404 });
    return Response.json({ data: row });
  } catch (e: any) {
    if (e?.message === "Unauthorized") return Response.json({ error: "Unauthorized" }, { status: 401 });
    return Response.json({ error: "Failed to fetch row" }, { status: 400 });
  }
}

export async function PATCH(request: Request, { params }: { params: { table: string; id: string } }) {
  try {
    requireAuth(request);
    const table = getTableOrThrow(params.table);
    const body = await request.json().catch(() => ({}));
    const keys = Object.keys(body || {}).filter(isValidIdentifier);
    if (!keys.length) return Response.json({ error: "No valid fields to update" }, { status: 400 });

    const setSql = keys.map((k) => `${quoteIdentifier(k)} = ?`).join(", ");
    const args = [...keys.map((k) => body[k]), params.id];
    await pool.query(`UPDATE ${table} SET ${setSql} WHERE id = ?`, args);
    return Response.json({ ok: true });
  } catch (e: any) {
    if (e?.message === "Unauthorized") return Response.json({ error: "Unauthorized" }, { status: 401 });
    return Response.json({ error: "Failed to update row" }, { status: 400 });
  }
}

export async function DELETE(request: Request, { params }: { params: { table: string; id: string } }) {
  try {
    requireAuth(request);
    const table = getTableOrThrow(params.table);
    await pool.query(`DELETE FROM ${table} WHERE id = ?`, [params.id]);
    return Response.json({ ok: true });
  } catch (e: any) {
    if (e?.message === "Unauthorized") return Response.json({ error: "Unauthorized" }, { status: 401 });
    return Response.json({ error: "Failed to delete row" }, { status: 400 });
  }
}
