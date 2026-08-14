-- ---------------------------------------------------------------------------
-- Audit triggers: log EVERY report/photo insert into activity_log, regardless
-- of which app wrote it.
--
-- WHY: the website's logActivity() only runs inside the website's own API
-- routes. The MOBILE tracker app writes reports/photos straight to the database
-- (via the separate lbi-backend), so its captures NEVER reached the audit log.
-- That is why a wrong-project capture (58 Mumbai points saved into "Test" on
-- 2026-08-13) was invisible in /activity. A DB-level trigger catches all of
-- them — web AND mobile — at the source.
--
-- Safe to re-run (drops + recreates). Requires the TRIGGER privilege — run as
-- root if the app DB user is not allowed to create triggers.
-- ---------------------------------------------------------------------------

DROP TRIGGER IF EXISTS trg_reports_audit_ins;
CREATE TRIGGER trg_reports_audit_ins
AFTER INSERT ON reports
FOR EACH ROW
INSERT INTO activity_log
  (id, user_id, user_email, action, table_name, entity_id, project_id, details)
VALUES (
  UUID(),
  NEW.user_id,
  (SELECT a.email FROM app_auth_users a WHERE a.id = NEW.user_id LIMIT 1),
  'create',
  'reports',
  NEW.id,
  NEW.project_id,
  JSON_OBJECT('source', 'db_trigger',
              'category', NEW.category,
              'lat', NEW.loc_lat,
              'lon', NEW.loc_lon)
);

-- Optional: also log photo uploads (one row per photo). Enable if you want the
-- audit to show photo activity too — it is noisier (many rows per report).
-- DROP TRIGGER IF EXISTS trg_report_photos_audit_ins;
-- CREATE TRIGGER trg_report_photos_audit_ins
-- AFTER INSERT ON report_photos
-- FOR EACH ROW
-- INSERT INTO activity_log
--   (id, user_id, user_email, action, table_name, entity_id, project_id, details)
-- VALUES (
--   UUID(), NEW.user_id,
--   (SELECT a.email FROM app_auth_users a WHERE a.id = NEW.user_id LIMIT 1),
--   'upload', 'report_photos', NEW.id, NULL,
--   JSON_OBJECT('source', 'db_trigger', 'report_id', NEW.report_id, 'url', NEW.url)
-- );
