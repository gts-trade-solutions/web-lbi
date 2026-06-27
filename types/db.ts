export type Project = {
  id: string;
  name?: string | null;
  created_at?: string | null;
};

export type Report = {
  id: string;
  project_id?: string | null;
  category?: string | null;
  description?: string | null;
  created_at?: string | null;
  location_lat?: number | null;
  location_lng?: number | null;
};

export type ReportPhoto = {
  id: string;
  report_id: string;
  // You might have one of these in DB:
  url?: string | null;     // full http url
  path?: string | null;    // storage path
  created_at?: string | null;
};

export type ReportPathPoint = {
  id: string;
  report_id: string;
  seq?: number | null;
  km?: number | null;
  latitude?: number | null;
  longitude?: number | null;
  details?: string | null;
  location_text?: string | null;
  vehicle_movement?: string | null;
};
