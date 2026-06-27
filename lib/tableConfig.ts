export const ALLOWED_TABLES = new Set([
  "users",
  "profiles",
  "projects",
  "reports",
  "report_path_points",
  "report_photos",
  "bulk_import_history",
  "project_route_pages",
  "project_route_page_images",
  "project_route_page_locations",
  "routes",
]);

export function isAllowedTable(table: string) {
  return ALLOWED_TABLES.has(table);
}

export function isValidIdentifier(v: string) {
  return /^[a-zA-Z_][a-zA-Z0-9_]*$/.test(v);
}

export function quoteIdentifier(v: string) {
  if (!isValidIdentifier(v)) throw new Error("Invalid identifier");
  return `\`${v}\``;
}

