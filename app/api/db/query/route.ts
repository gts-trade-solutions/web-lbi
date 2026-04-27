/* eslint-disable @typescript-eslint/no-explicit-any */
import { v4 as uuidv4 } from "uuid";
import pool from "../../../../lib/db";
import { requireAuth } from "../../../../lib/auth";
import { isAllowedTable, isValidIdentifier, quoteIdentifier } from "../../../../lib/tableConfig";

export const runtime = "nodejs";

type Filter =
  | { type: "eq"; column: string; value: any }
  | { type: "in"; column: string; value: any[] }
  | { type: "is"; column: string; value: any }
  | { type: "textSearch"; column: string; value: string };

type QueryBody = {
  table: string;
  action: "select" | "insert" | "update" | "delete";
  select?: string;
  payload?: any;
  filters?: Filter[];
  orders?: Array<{ column: string; ascending?: boolean; nullsFirst?: boolean }>;
  limit?: number;
  single?: boolean;
  maybeSingle?: boolean;
  head?: boolean;
  count?: "exact" | "planned" | "estimated" | null;
};

const SEARCHABLE_COLUMNS: Record<string, string[]> = {
  projects: ["name", "title", "project_name", "description", "id"],
  reports: ["category", "description", "remarks_action", "point_key", "id"],
  report_path_points: ["details", "location_text", "vehicle_movement"],
};

function parseSelect(selectRaw: string | undefined) {
  const raw = (selectRaw || "*").trim();
  if (raw === "*") return "*";
  const cols = raw
    .split(",")
    .map((c) => c.trim())
    .filter(Boolean);
  if (!cols.length) return "*";
  return cols.map((c) => quoteIdentifier(c)).join(", ");
}

function buildWhere(table: string, filters: Filter[] = []) {
  const clauses: string[] = [];
  const args: any[] = [];

  for (const f of filters) {
    if (!isValidIdentifier(f.column)) continue;
    const col = quoteIdentifier(f.column);

    if (f.type === "eq") {
      clauses.push(`${col} = ?`);
      args.push(f.value);
      continue;
    }

    if (f.type === "in") {
      const arr = Array.isArray(f.value) ? f.value : [];
      if (!arr.length) {
        clauses.push("1 = 0");
        continue;
      }
      clauses.push(`${col} IN (${arr.map(() => "?").join(", ")})`);
      args.push(...arr);
      continue;
    }

    if (f.type === "is") {
      if (f.value === null || typeof f.value === "undefined") {
        clauses.push(`${col} IS NULL`);
      } else {
        clauses.push(`${col} IS ?`);
        args.push(f.value);
      }
      continue;
    }

    if (f.type === "textSearch") {
      const q = String(f.value || "").trim();
      if (!q) continue;
      const cols = SEARCHABLE_COLUMNS[table] || [f.column];
      const validCols = cols.filter(isValidIdentifier);
      if (!validCols.length) continue;
      const like = `%${q}%`;
      clauses.push(`(${validCols.map((c) => `LOWER(${quoteIdentifier(c)}) LIKE LOWER(?)`).join(" OR ")})`);
      for (let i = 0; i < validCols.length; i += 1) args.push(like);
    }
  }

  if (!clauses.length) return { sql: "", args };
  return {
    sql: ` WHERE ${clauses.join(" AND ")}`,
    args,
  };
}

function buildOrder(orders: QueryBody["orders"] = []) {
  const chunks: string[] = [];
  for (const o of orders || []) {
    if (!o || !isValidIdentifier(o.column)) continue;
    const dir = o.ascending === false ? "DESC" : "ASC";
    chunks.push(`${quoteIdentifier(o.column)} ${dir}`);
  }
  if (!chunks.length) return "";
  return ` ORDER BY ${chunks.join(", ")}`;
}

function sanitizeTable(table: string) {
  if (!isAllowedTable(table)) throw new Error("Table is not allowed");
  return quoteIdentifier(table);
}

async function runSelect(body: QueryBody) {
  const table = String(body.table || "").trim();
  const tableSql = sanitizeTable(table);
  const selectSql = parseSelect(body.select);
  const where = buildWhere(table, body.filters || []);
  const order = buildOrder(body.orders || []);
  const limit = Number.isFinite(body.limit) ? Math.max(0, Number(body.limit)) : undefined;

  let count: number | null = null;
  if (body.count === "exact") {
    const [countRows] = await pool.query(
      `SELECT COUNT(*) AS count FROM ${tableSql}${where.sql}`,
      where.args
    );
    count = Array.isArray(countRows) ? Number((countRows[0] as any)?.count || 0) : 0;
  }

  if (body.head) return { data: null, count };

  const limitSql = limit || body.single || body.maybeSingle ? ` LIMIT ${limit || 1}` : "";
  const [rows] = await pool.query(
    `SELECT ${selectSql} FROM ${tableSql}${where.sql}${order}${limitSql}`,
    where.args
  );

  const dataRows = Array.isArray(rows) ? rows : [];
  if (body.single) {
    if (!dataRows.length) {
      return { error: "Row not found", status: 406 };
    }
    return { data: dataRows[0], count };
  }
  if (body.maybeSingle) return { data: dataRows[0] || null, count };
  return { data: dataRows, count };
}

async function runInsert(body: QueryBody) {
  const table = String(body.table || "").trim();
  const tableSql = sanitizeTable(table);
  const rawRows = Array.isArray(body.payload) ? body.payload : [body.payload];
  const rows = rawRows.filter((r) => r && typeof r === "object");
  if (!rows.length) return { error: "Insert payload is required", status: 400 };

  const prepared = rows.map((row) => {
    const clone: Record<string, any> = { ...row };
    if (!clone.id) clone.id = uuidv4();
    return clone;
  });

  const columns = Array.from(
    new Set(
      prepared
        .flatMap((r) => Object.keys(r))
        .filter(isValidIdentifier)
    )
  );

  if (!columns.length) return { error: "No valid columns in payload", status: 400 };
  const colsSql = columns.map((c) => quoteIdentifier(c)).join(", ");
  const placeholders = `(${columns.map(() => "?").join(", ")})`;
  const valuesSql = prepared.map(() => placeholders).join(", ");
  const args = prepared.flatMap((row) => columns.map((c) => (c in row ? row[c] : null)));

  await pool.query(`INSERT INTO ${tableSql} (${colsSql}) VALUES ${valuesSql}`, args);

  const ids = prepared.map((r) => r.id).filter(Boolean);
  if (ids.length) {
    const [insertedRows] = await pool.query(
      `SELECT ${parseSelect(body.select)} FROM ${tableSql} WHERE id IN (${ids.map(() => "?").join(", ")})`,
      ids
    );
    const arr = Array.isArray(insertedRows) ? insertedRows : [];
    if (body.single) return { data: arr[0] || null };
    if (body.maybeSingle) return { data: arr[0] || null };
    return { data: arr };
  }

  return { data: prepared };
}

async function runUpdate(body: QueryBody) {
  const table = String(body.table || "").trim();
  const tableSql = sanitizeTable(table);
  const where = buildWhere(table, body.filters || []);
  const payload = body.payload && typeof body.payload === "object" ? body.payload : null;
  if (!payload) return { error: "Update payload is required", status: 400 };

  const columns = Object.keys(payload).filter(isValidIdentifier);
  if (!columns.length) return { error: "No valid columns in update payload", status: 400 };

  const setSql = columns.map((c) => `${quoteIdentifier(c)} = ?`).join(", ");
  const args = [...columns.map((c) => payload[c]), ...where.args];

  await pool.query(`UPDATE ${tableSql} SET ${setSql}${where.sql}`, args);

  if (body.select || body.single || body.maybeSingle) {
    return runSelect({ ...body, action: "select" });
  }

  return { data: [] };
}

async function runDelete(body: QueryBody) {
  const table = String(body.table || "").trim();
  const tableSql = sanitizeTable(table);
  const where = buildWhere(table, body.filters || []);
  await pool.query(`DELETE FROM ${tableSql}${where.sql}`, where.args);
  return { data: [] };
}

export async function POST(request: Request) {
  try {
    requireAuth(request);
    const body = (await request.json().catch(() => ({}))) as QueryBody;
    if (!body?.table || !body?.action) {
      return Response.json({ error: "Invalid request" }, { status: 400 });
    }

    let result: any;
    if (body.action === "select") result = await runSelect(body);
    else if (body.action === "insert") result = await runInsert(body);
    else if (body.action === "update") result = await runUpdate(body);
    else if (body.action === "delete") result = await runDelete(body);
    else return Response.json({ error: "Unsupported action" }, { status: 400 });

    if (result?.error) {
      return Response.json({ error: result.error }, { status: result.status || 400 });
    }

    return Response.json({
      data: typeof result?.data === "undefined" ? null : result.data,
      count: result?.count ?? null,
    });
  } catch (e: any) {
    if (e?.message === "Unauthorized") {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    return Response.json({ error: "Query failed" }, { status: 500 });
  }
}
