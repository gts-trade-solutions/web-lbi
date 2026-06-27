import pool from "./db";

// How long a soft-deleted project stays in the recycle bin before it is
// permanently purged.
export const TRASH_RETENTION_DAYS = 30;

// The soft-delete columns are added lazily so existing databases (migrated
// from Supabase) don't need a manual ALTER before the recycle bin works.
// Cached so we only hit SHOW COLUMNS / ALTER TABLE once per process.
let ensured: Promise<boolean> | null = null;

async function columnExists(column: string): Promise<boolean> {
  const [rows] = await pool.query("SHOW COLUMNS FROM projects LIKE ?", [column]);
  return Array.isArray(rows) && rows.length > 0;
}

/**
 * Ensure the `projects` table has the `deleted_at` / `deleted_by` columns the
 * recycle bin relies on. Returns true once the columns are present (or were
 * added). If the ALTER fails (e.g. insufficient privileges) it resolves false
 * so callers can degrade to hard-delete instead of crashing.
 */
export function ensureTrashColumns(): Promise<boolean> {
  if (!ensured) {
    ensured = (async () => {
      try {
        if (!(await columnExists("deleted_at"))) {
          await pool.query(
            "ALTER TABLE projects ADD COLUMN deleted_at DATETIME NULL DEFAULT NULL"
          );
        }
        if (!(await columnExists("deleted_by"))) {
          await pool.query(
            "ALTER TABLE projects ADD COLUMN deleted_by VARCHAR(36) NULL DEFAULT NULL"
          );
        }
        // Index keeps the "active projects" and "trash" filters cheap.
        try {
          await pool.query(
            "ALTER TABLE projects ADD INDEX idx_projects_deleted_at (deleted_at)"
          );
        } catch {
          // Index probably already exists — ignore.
        }
        return true;
      } catch (error) {
        console.error("[projects-trash] failed to ensure trash columns:", error);
        ensured = null; // allow a retry on the next request
        return false;
      }
    })();
  }
  return ensured;
}

/**
 * Permanently remove any project that has been in the recycle bin longer than
 * the retention window. Safe to call on every trash read.
 */
export async function purgeExpiredTrash(): Promise<number> {
  const hasColumns = await ensureTrashColumns();
  if (!hasColumns) return 0;

  const [result] = await pool.query(
    "DELETE FROM projects WHERE deleted_at IS NOT NULL AND deleted_at < (NOW() - INTERVAL ? DAY)",
    [TRASH_RETENTION_DAYS]
  );
  return Number((result as { affectedRows?: number })?.affectedRows || 0);
}
