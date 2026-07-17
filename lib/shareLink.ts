/* eslint-disable @typescript-eslint/no-explicit-any */
// Password-protected public share links for the animated route report.
//
// A survey team creates a link from the route-map ("Share") which stores a
// bcrypt hash of a chosen password plus the exact set of reports being shown.
// A client opens /share/<token>, types the password, and is handed a short
// JWT "share session" (signed with the same JWT_SECRET_KEY the app already
// uses). Every public data request carries that session so the read-only map
// endpoints can serve reports/photos WITHOUT any logged-in user — but nobody
// without the password can mint a session, and a session only unlocks the one
// project the link points at.
import crypto from "crypto";
import jwt from "jsonwebtoken";
import pool from "./db";

// How long a client stays "unlocked" after typing the password once.
const SHARE_SESSION_TTL = "12h";

function secret(): string {
  const s = process.env.JWT_SECRET_KEY;
  if (!s) throw new Error("JWT_SECRET_KEY is not configured");
  return s;
}

// URL-safe, unguessable public id that lives in the /share/<token> URL.
export function makeShareToken(): string {
  return crypto.randomBytes(18).toString("base64url");
}

export function signShareSession(token: string, projectId: string): string {
  return jwt.sign(
    { kind: "share", share: token, project: projectId },
    secret(),
    { expiresIn: SHARE_SESSION_TTL }
  );
}

// Returns true only for a live session token that belongs to THIS share.
export function verifyShareSession(session: string, token: string): boolean {
  if (!session || !token) return false;
  try {
    const p = jwt.verify(session, secret()) as any;
    return p?.kind === "share" && p?.share === token;
  } catch {
    return false;
  }
}

// Pull the session out of the request: Authorization: Bearer, an
// X-Share-Session header, or a ?s= query param (whichever is present).
export function readShareSession(request: Request): string {
  const h = request.headers.get("authorization") || "";
  const bearer = /^bearer\s+/i.test(h) ? h.replace(/^bearer\s+/i, "").trim() : "";
  if (bearer) return bearer;
  const header = request.headers.get("x-share-session");
  if (header) return header.trim();
  try {
    return new URL(request.url).searchParams.get("s") || "";
  } catch {
    return "";
  }
}

export type ShareRow = {
  id: string;
  token: string;
  project_id: string;
  password_hash: string;
  title: string | null;
  selected_ids: string | null;
  revoked: number;
};

// Create the table on first use so existing installs upgrade themselves.
export async function ensureShareTable(): Promise<void> {
  await pool.query(
    `CREATE TABLE IF NOT EXISTS report_shares (
       id VARCHAR(64) NOT NULL PRIMARY KEY,
       token VARCHAR(80) NOT NULL,
       project_id VARCHAR(64) NOT NULL,
       password_hash VARCHAR(255) NOT NULL,
       title VARCHAR(255) NULL,
       selected_ids MEDIUMTEXT NULL,
       created_by VARCHAR(64) NULL,
       created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
       revoked TINYINT NOT NULL DEFAULT 0,
       UNIQUE KEY uniq_token (token),
       KEY idx_project (project_id)
     ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4`
  );
}

// Live (non-revoked) share row for a token, or null.
export async function getShareByToken(token: string): Promise<ShareRow | null> {
  if (!token) return null;
  await ensureShareTable();
  const [rows] = await pool.query(
    "SELECT * FROM report_shares WHERE token = ? AND revoked = 0 LIMIT 1",
    [token]
  );
  const r = Array.isArray(rows) && rows.length ? (rows[0] as ShareRow) : null;
  return r;
}

// Parse the stored selected_ids JSON into a string[] ([] = whole project).
export function parseSelectedIds(raw: string | null | undefined): string[] {
  if (!raw) return [];
  try {
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr.map((x) => String(x || "").trim()).filter(Boolean) : [];
  } catch {
    return [];
  }
}
