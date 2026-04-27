# Supabase to MySQL + S3 Migration Notes

## Table mapping
- `auth.users` (Supabase Auth) -> `users` (custom auth table with `password_hash`)
- `profiles` -> `profiles`
- `projects` -> `projects`
- `routes` -> `routes`
- `reports` -> `reports`
- `report_path_points` -> `report_path_points`
- `report_photos` -> `report_photos`
- `bulk_import_history` -> `bulk_import_history`
- `project_route_pages` -> `project_route_pages`
- `project_route_page_images` -> `project_route_page_images`
- `project_route_page_locations` -> `project_route_page_locations`

## Data import approach
1. Export Supabase table data (CSV/JSON) table-by-table.
2. Transform UUID values directly into MySQL `VARCHAR(36)` fields.
3. Convert timestamps to MySQL `YYYY-MM-DD HH:MM:SS` where needed.
4. Import in FK-safe order:
   - `users`, `profiles`
   - `projects`, `routes`
   - `reports`
   - `report_path_points`, `report_photos`
   - `project_route_pages`
   - `project_route_page_locations`, `project_route_page_images`
   - `bulk_import_history`

## Image/storage migration (Supabase Storage -> S3)
1. Export objects from Supabase buckets used by the app (notably report photos and GA/map files).
2. Upload files to S3 while preserving path shape where possible, for example:
   - `reports/{projectId}/{reportId}/{fileName}`
   - `projects/{projectId}/{pageId}/{fileName}`
3. Update DB URL fields:
   - Replace Supabase public URLs with `NEXT_PUBLIC_S3_BUCKET_URL/<key>`.
4. If old rows stored `{bucket,path}` only, keep those fields and regenerate URLs using the new bucket base URL.

## Assumptions
- S3 bucket objects are publicly readable or can be accessed by signed URL.
- Existing frontend expects mostly unchanged field names and nullable columns.
- Full-text Postgres `textSearch` behavior is approximated by MySQL `LIKE` search over key text columns.
- JWT auth is now app-managed (`users` + `password_hash`) instead of Supabase Auth.

