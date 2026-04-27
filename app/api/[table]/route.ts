/* eslint-disable @typescript-eslint/no-explicit-any */
import pool from "../../../lib/db";
import { requireAuth } from "../../../lib/auth";
import { isAllowedTable, isValidIdentifier, quoteIdentifier } from "../../../lib/tableConfig";

export const runtime = "nodejs";

function getTableOrThrow(table: string) {
  if (!isAllowedTable(table)) throw new Error("Table is not allowed");
  return quoteIdentifier(table);
}

export async function GET(request: Request, { params }: { params: { table: string } }) {
  try {
    requireAuth(request);
    const table = getTableOrThrow(params.table);
    const { searchParams } = new URL(request.url);
    const limit = Number(searchParams.get("limit") || 100);

    const [rows] = await pool.query(`SELECT * FROM ${table} LIMIT ?`, [Math.max(1, Math.min(limit, 500))]);
    return Response.json({ data: Array.isArray(rows) ? rows : [] });
  } catch (e: any) {
    if (e?.message === "Unauthorized") return Response.json({ error: "Unauthorized" }, { status: 401 });
    return Response.json({ error: "Failed to fetch rows" }, { status: 400 });
  }
}

export async function POST(request: Request, { params }: { params: { table: string } }) {
  try {
    requireAuth(request);
    const tableName = params.table;
    const table = getTableOrThrow(tableName);
    const body = await request.json().catch(() => ({}));
    const rows = Array.isArray(body) ? body : [body];
    const validRows = rows.filter((r) => r && typeof r === "object");
    if (!validRows.length) return Response.json({ error: "Invalid payload" }, { status: 400 });

    const columns = Array.from(new Set(validRows.flatMap((r) => Object.keys(r)).filter(isValidIdentifier)));
    if (!columns.length) return Response.json({ error: "No valid columns" }, { status: 400 });

    const colsSql = columns.map((c) => quoteIdentifier(c)).join(", ");
    const valuesSql = validRows.map(() => `(${columns.map(() => "?").join(", ")})`).join(", ");
    const args = validRows.flatMap((row) => columns.map((c) => (c in row ? row[c] : null)));

    await pool.query(`INSERT INTO ${table} (${colsSql}) VALUES ${valuesSql}`, args);
    return Response.json({ ok: true });
  } catch (e: any) {
    if (e?.message === "Unauthorized") return Response.json({ error: "Unauthorized" }, { status: 401 });
    return Response.json({ error: "Failed to insert rows" }, { status: 400 });
  }
}
