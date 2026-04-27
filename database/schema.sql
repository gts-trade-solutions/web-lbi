CREATE TABLE IF NOT EXISTS users (
  id VARCHAR(36) PRIMARY KEY,
  email VARCHAR(255) NOT NULL UNIQUE,
  name VARCHAR(255) NULL,
  role VARCHAR(64) NOT NULL DEFAULT 'user',
  password_hash VARCHAR(255) NOT NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS profiles (
  id VARCHAR(36) PRIMARY KEY DEFAULT (UUID()),
  user_id VARCHAR(36) NOT NULL UNIQUE,
  email VARCHAR(255) NULL,
  name VARCHAR(255) NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  INDEX idx_profiles_email (email),
  CONSTRAINT fk_profiles_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS projects (
  id VARCHAR(36) PRIMARY KEY DEFAULT (UUID()),
  user_id VARCHAR(36) NULL,
  name VARCHAR(255) NULL,
  title VARCHAR(255) NULL,
  project_name VARCHAR(255) NULL,
  description TEXT NULL,
  created_by VARCHAR(36) NULL,
  last_modified_by VARCHAR(36) NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  INDEX idx_projects_user_id (user_id),
  INDEX idx_projects_created_at (created_at),
  INDEX idx_projects_name (name)
);

CREATE TABLE IF NOT EXISTS routes (
  id VARCHAR(36) PRIMARY KEY DEFAULT (UUID()),
  project_id VARCHAR(36) NOT NULL,
  user_id VARCHAR(36) NULL,
  name VARCHAR(255) NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  INDEX idx_routes_project_id (project_id),
  CONSTRAINT fk_routes_project FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS reports (
  id VARCHAR(36) PRIMARY KEY DEFAULT (UUID()),
  project_id VARCHAR(36) NOT NULL,
  route_id VARCHAR(36) NULL,
  created_by VARCHAR(36) NULL,
  point_key VARCHAR(255) NULL,
  category VARCHAR(255) NULL,
  description TEXT NULL,
  remarks_action TEXT NULL,
  difficulty VARCHAR(64) NULL,
  vehicle_movement VARCHAR(64) NULL,
  latitude DECIMAL(12,8) NULL,
  longitude DECIMAL(12,8) NULL,
  sort_order INT NULL,
  status VARCHAR(50) NULL DEFAULT 'active',
  loc_lat DOUBLE NULL,
  loc_lon DOUBLE NULL,
  loc_acc DOUBLE NULL,
  loc_time DATETIME NULL,
  kms DECIMAL(12,4) NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  INDEX idx_reports_project_id (project_id),
  INDEX idx_reports_route_id (route_id),
  INDEX idx_reports_point_key (point_key),
  INDEX idx_reports_category (category),
  INDEX idx_reports_difficulty (difficulty),
  INDEX idx_reports_status (status),
  INDEX idx_reports_sort_order (sort_order),
  CONSTRAINT fk_reports_project FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE,
  CONSTRAINT fk_reports_route FOREIGN KEY (route_id) REFERENCES routes(id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS report_path_points (
  id VARCHAR(36) PRIMARY KEY DEFAULT (UUID()),
  report_id VARCHAR(36) NOT NULL,
  user_id VARCHAR(36) NULL,
  seq INT NULL,
  km DECIMAL(12,4) NULL,
  latitude DOUBLE NULL,
  longitude DOUBLE NULL,
  elevation DOUBLE NULL,
  accuracy DOUBLE NULL,
  timestamp DATETIME NULL,
  details TEXT NULL,
  location_text TEXT NULL,
  vehicle_movement VARCHAR(64) NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  INDEX idx_report_path_points_report_id (report_id),
  INDEX idx_report_path_points_seq (seq),
  CONSTRAINT fk_report_path_points_report FOREIGN KEY (report_id) REFERENCES reports(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS report_photos (
  id VARCHAR(36) PRIMARY KEY DEFAULT (UUID()),
  report_id VARCHAR(36) NOT NULL,
  url TEXT NULL,
  file_name VARCHAR(255) NULL,
  bucket VARCHAR(255) NULL,
  path TEXT NULL,
  width INT NULL,
  height INT NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  INDEX idx_report_photos_report_id (report_id),
  CONSTRAINT fk_report_photos_report FOREIGN KEY (report_id) REFERENCES reports(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS bulk_import_history (
  id VARCHAR(36) PRIMARY KEY DEFAULT (UUID()),
  project_id VARCHAR(36) NOT NULL,
  master_file_name VARCHAR(255) NULL,
  master_file_hash VARCHAR(255) NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  INDEX idx_bulk_import_project_id (project_id),
  INDEX idx_bulk_import_hash (master_file_hash),
  CONSTRAINT fk_bulk_import_project FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS project_route_pages (
  id VARCHAR(36) PRIMARY KEY DEFAULT (UUID()),
  project_id VARCHAR(36) NOT NULL,
  user_id VARCHAR(36) NULL,
  objective TEXT NULL,
  map_mode VARCHAR(64) NULL,
  preset_map_key VARCHAR(255) NULL,
  map_file_url TEXT NULL,
  conclusion_html LONGTEXT NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  INDEX idx_project_route_pages_project_id (project_id),
  CONSTRAINT fk_project_route_pages_project FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS project_route_page_locations (
  id VARCHAR(36) PRIMARY KEY DEFAULT (UUID()),
  project_page_id VARCHAR(36) NOT NULL,
  project_id VARCHAR(36) NOT NULL,
  user_id VARCHAR(36) NULL,
  label VARCHAR(255) NULL,
  pin_type VARCHAR(64) NULL,
  sort_order INT NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  INDEX idx_project_route_page_locations_project_page_id (project_page_id),
  INDEX idx_project_route_page_locations_project_id (project_id),
  INDEX idx_project_route_page_locations_sort_order (sort_order),
  CONSTRAINT fk_project_route_page_locations_page FOREIGN KEY (project_page_id) REFERENCES project_route_pages(id) ON DELETE CASCADE,
  CONSTRAINT fk_project_route_page_locations_project FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS project_route_page_images (
  id VARCHAR(36) PRIMARY KEY DEFAULT (UUID()),
  project_page_id VARCHAR(36) NOT NULL,
  project_id VARCHAR(36) NOT NULL,
  user_id VARCHAR(36) NULL,
  file_url TEXT NULL,
  file_name VARCHAR(255) NULL,
  mime_type VARCHAR(255) NULL,
  file_size BIGINT NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  INDEX idx_project_route_page_images_project_page_id (project_page_id),
  INDEX idx_project_route_page_images_project_id (project_id),
  CONSTRAINT fk_project_route_page_images_page FOREIGN KEY (project_page_id) REFERENCES project_route_pages(id) ON DELETE CASCADE,
  CONSTRAINT fk_project_route_page_images_project FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS project_ga_drawings (
  id VARCHAR(36) PRIMARY KEY DEFAULT (UUID()),
  project_id VARCHAR(36) NOT NULL UNIQUE,
  image_url TEXT NULL,
  image_key TEXT NULL,
  file_name VARCHAR(255) NULL,
  conclusion_html LONGTEXT NULL,
  created_by VARCHAR(36) NULL,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  INDEX idx_project_ga_drawings_project_id (project_id)
);
