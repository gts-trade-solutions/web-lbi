"use client";

import React, { useEffect, useMemo, useRef, useState } from "react";
import { useRouter } from "next/navigation";
import * as XLSX from "xlsx";
import { downloadProjectsCSV } from "../../lib/download";

type ProjectRow = {
  id: string;
  user_id?: string | null;
  name?: string | null;
  title?: string | null;
  project_name?: string | null;
  description?: string | null;
  created_at?: string | null;
  updated_at?: string | null;
  last_modified_by?: string | null;
  created_by?: string | null;
};

type ParsedPointRow = {
  point_key: string;
  latitude: number;
  longitude: number;
  category: string;
  normalizedCategory?: string | null;
  description?: string | null;
  difficulty?: string | null;
  remarks_action?: string | null;
};

type ParsedImageMapRow = {
  file_name: string;
  point_key: string | null;
  image_key?: string | null;
};

type ParsedCombinedRow = {
  point_key: string;
  latitude: number;
  longitude: number;
  category: string;
  description?: string | null;
  file_name?: string | null;
  image_key?: string | null;
  difficulty?: string | null;
  remarks_action?: string | null;
};

type ImportSummary = {
  pointsRead: number;
  reportsCreatedOrUsed: number;
  imagesSelected: number;
  imagesUploaded: number;
  photosInserted: number;
  // Optional fields populated when the bulk-import API also performed a
  // server-side S3 lookup. Older callers ignore these.
  photosMatchedServer?: number;
  photosMissingServer?: number;
  photosMissingSamples?: string[];
  s3IndexSize?: number;
  totalPhotosResolved?: number;
  masterRowsReferencingImages?: number;
  reportsInserted?: number;
  reportsUpdated?: number;
  noGpsImages: string[];
  missingFilesInUpload: string[];
  extraFilesNotInMap: string[];
  errors: string[];
  duplicateSelectedImages: string[];
  duplicateMappingFiles: string[];
  missingPointKeysInPointsCsv: string[];
  invalidCategories: string[];
};

const CATEGORY_OPTIONS = [
  "Footpath Bridge",
  "Low Tension Cable",
  "High Tension Cable",
  "Towerline Cable",
  "Take Diversion",
  "Towerline",
  "Underpass Bridge",
  "Tree Branches",
  "Bridge",
  "Petrol bunk",
  "Signboard",
  "Electric Sign Board",
  "Camera Pole",
  "Toll Plaza",
  "Junction left",
  "Bend",
  "Junction right",
] as const;

const CATEGORY_ALIAS_MAP: Record<string, string> = {
  "footpath bridge": "Footpath Bridge",
  "footpathbridge": "Footpath Bridge",
  "foot bridge": "Footpath Bridge",
  "pedestrian bridge": "Footpath Bridge",

  "low tension cable": "Low Tension Cable",
  "low tension cables": "Low Tension Cable",
  "lt cable": "Low Tension Cable",
  "lt cables": "Low Tension Cable",
  "lt line": "Low Tension Cable",
  "lt lines": "Low Tension Cable",

  "high tension cable": "High Tension Cable",
  "high tension cables": "High Tension Cable",
  "ht cable": "High Tension Cable",
  "ht cables": "High Tension Cable",
  "ht line": "High Tension Cable",
  "ht lines": "High Tension Cable",

  "towerline cable": "Towerline Cable",
  "towerline cables": "Towerline Cable",
  "tower line cable": "Towerline Cable",
  "tower line cables": "Towerline Cable",

  "take diversion": "Take Diversion",
  diversion: "Take Diversion",
  diversions: "Take Diversion",
  "take diversions": "Take Diversion",

  towerline: "Towerline",
  "tower line": "Towerline",
  "tower lines": "Towerline",
  "transmission tower": "Towerline",

  "underpass bridge": "Underpass Bridge",
  underpass: "Underpass Bridge",
  "under bridge": "Underpass Bridge",

  "tree branches": "Tree Branches",
  "tree branch": "Tree Branches",
  branches: "Tree Branches",
  tree: "Tree Branches",

  bridge: "Bridge",
  bridges: "Bridge",
  flyover: "Bridge",

  "petrol bunk": "Petrol bunk",
  "petrol bunks": "Petrol bunk",
  "petrol pump": "Petrol bunk",
  "fuel station": "Petrol bunk",
  "fuel bunk": "Petrol bunk",

  signboard: "Signboard",
  signboards: "Signboard",
  "sign board": "Signboard",
  "sign boards": "Signboard",
  "road sign": "Signboard",

  "electric sign board": "Electric Sign Board",
  "electric signboard": "Electric Sign Board",
  "electric sign boards": "Electric Sign Board",
  "electrical sign board": "Electric Sign Board",
  "illuminated sign board": "Electric Sign Board",

  "camera pole": "Camera Pole",
  "camera poles": "Camera Pole",
  "cctv pole": "Camera Pole",
  "surveillance pole": "Camera Pole",

  "toll plaza": "Toll Plaza",
  "toll plazas": "Toll Plaza",
  toll: "Toll Plaza",

  "junction left": "Junction left",
  "left junction": "Junction left",
  "left turn junction": "Junction left",
  "left turn": "Junction left",

  bend: "Bend",
  bends: "Bend",
  curve: "Bend",
  curves: "Bend",

  "junction right": "Junction right",
  "right junction": "Junction right",
  "right turn junction": "Junction right",
  "right turn": "Junction right",
};

function normalizeHeader(h: string) {
  return h.trim().toLowerCase().replace(/\s+/g, "_");
}

function detectDelimiter(line: string) {
  return line.includes("\t") ? "\t" : ",";
}

function stripBom(text: string) {
  return text.replace(/^\uFEFF/, "");
}

function splitDelimitedLine(line: string, delim: string) {
  const out: string[] = [];
  let cur = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];

    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        cur += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (ch === delim && !inQuotes) {
      out.push(cur.trim());
      cur = "";
      continue;
    }

    cur += ch;
  }

  out.push(cur.trim());
  return out;
}

function normalizeCategoryName(input: string) {
  const raw = String(input || "").trim();
  if (!raw) return "Unknown";

  const lowered = raw
    .toLowerCase()
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const exact = CATEGORY_OPTIONS.find((x) => x.toLowerCase() === lowered);
  if (exact) return exact;

  return CATEGORY_ALIAS_MAP[lowered] || raw;
}

function normalizeDifficulty(input: string) {
  return String(input || "").trim() || "green";
}

function parseCSVLike(text: string) {
  const raw = stripBom(text).replace(/\r/g, "").trim();
  if (!raw) return { headers: [] as string[], rows: [] as string[][] };

  const lines = raw.split("\n").filter((l) => l.trim());
  if (!lines.length) return { headers: [] as string[], rows: [] as string[][] };

  const delim = detectDelimiter(lines[0]);
  const headers = splitDelimitedLine(lines[0], delim).map((h) => normalizeHeader(h));
  const rows = lines
    .slice(1)
    .map((ln) => splitDelimitedLine(ln, delim).map((c) => c.trim()));

  return { headers, rows };
}

function colIndex(headers: string[], names: string[]) {
  for (const n of names) {
    const idx = headers.indexOf(n);
    if (idx >= 0) return idx;
  }
  return -1;
}

function isExcelFile(file: File) {
  const name = file.name.toLowerCase();
  return name.endsWith(".xlsx") || name.endsWith(".xls");
}

function parseNEToDecimalDecimalPart(value: string, kind: "lat" | "lon"): number | null {
  const raw = String(value || "")
    .replace(/\r/g, " ")
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!raw) return null;

  const regex =
    kind === "lat"
      ? /^([NS])\s*(\d{1,3}(?:\.\d+)?)$/i
      : /^([EW])\s*(\d{1,3}(?:\.\d+)?)$/i;

  const m = raw.match(regex);
  if (!m) return null;

  const dir = m[1].toUpperCase();
  let num = Number(m[2]);
  if (!Number.isFinite(num)) return null;

  if (dir === "S" || dir === "W") num = -num;
  return num;
}

function parseSeparateDMPart(value: string, kind: "lat" | "lon"): number | null {
  const raw = String(value || "")
    .replace(/\r/g, " ")
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!raw) return null;

  const regex =
    kind === "lat"
      ? /^([NS])\s*(\d{1,3})\s+(\d{1,2}(?:\.\d+)?)$/i
      : /^([EW])\s*(\d{1,3})\s+(\d{1,2}(?:\.\d+)?)$/i;

  const m = raw.match(regex);
  if (!m) return null;

  const dir = m[1].toUpperCase();
  const deg = Number(m[2]);
  const min = Number(m[3]);

  if (!Number.isFinite(deg) || !Number.isFinite(min)) return null;

  let valueNum = deg + min / 60;
  if (dir === "S" || dir === "W") valueNum = -valueNum;
  return valueNum;
}

function parseSingleCoordinate(value: string, kind: "lat" | "lon"): number | null {
  const raw = String(value || "").trim();
  if (!raw) return null;

  const plain = Number(raw);
  if (Number.isFinite(plain)) return plain;

  const dm = parseSeparateDMPart(raw, kind);
  if (dm != null) return dm;

  const neDecimal = parseNEToDecimalDecimalPart(raw, kind);
  if (neDecimal != null) return neDecimal;

  return null;
}

function haversineKm(lat1: number, lon1: number, lat2: number, lon2: number) {
  const toRad = (v: number) => (v * Math.PI) / 180;
  const R = 6371;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) *
      Math.cos(toRad(lat2)) *
      Math.sin(dLon / 2) *
      Math.sin(dLon / 2);
  return 2 * R * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

async function getRoadDistanceKm(
  fromLat: number,
  fromLon: number,
  toLat: number,
  toLon: number
): Promise<number> {
  try {
    const url = `https://router.project-osrm.org/route/v1/driving/${encodeURIComponent(
      fromLon
    )},${encodeURIComponent(fromLat)};${encodeURIComponent(
      toLon
    )},${encodeURIComponent(toLat)}?overview=false`;

    const res = await fetch(url, {
      headers: {
        Accept: "application/json",
      },
    });

    if (!res.ok) {
      return haversineKm(fromLat, fromLon, toLat, toLon);
    }

    const data = await res.json();
    const meters = data?.routes?.[0]?.distance;

    if (!Number.isFinite(meters)) {
      return haversineKm(fromLat, fromLon, toLat, toLon);
    }

    return Number(meters) / 1000;
  } catch {
    return haversineKm(fromLat, fromLon, toLat, toLon);
  }
}

async function readWorkbookRows(file: File): Promise<string[][]> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });

  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) return [];

  const sheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: "",
  }) as any[][];

  return rows.map((row) => row.map((cell) => String(cell ?? "").trim()));
}

async function readTextFile(file: File): Promise<string> {
  if (!file) throw new Error("No file selected.");

  try {
    return await file.text();
  } catch {
    return await new Promise((resolve, reject) => {
      try {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result || ""));
        reader.onerror = () =>
          reject(
            new Error(
              `Failed to read file "${file.name}". Please re-select the file from a normal local folder and try again.`
            )
          );
        reader.readAsText(file);
      } catch (err: any) {
        reject(
          new Error(
            `Failed to read file "${file.name}". ${err?.message || String(err)}`
          )
        );
      }
    });
  }
}

async function readStructuredFile(file: File) {
  if (isExcelFile(file)) {
    const rows = await readWorkbookRows(file);
    if (!rows.length) return { headers: [] as string[], rows: [] as string[][] };

    const headers = rows[0].map((h) => normalizeHeader(String(h || "")));
    const dataRows = rows
      .slice(1)
      .filter((r) => r.some((c) => String(c || "").trim() !== ""))
      .map((r) => r.map((c) => String(c || "").trim()));

    return { headers, rows: dataRows };
  }

  const text = await readTextFile(file);
  return parseCSVLike(text);
}

async function parseCombinedFile(file: File): Promise<ParsedCombinedRow[]> {
  const { headers, rows } = await readStructuredFile(file);
  if (!headers.length) return [];

  const iKey = colIndex(headers, ["point_key", "point", "seq", "key"]);
  const iLat = colIndex(headers, ["latitude", "lat"]);
  const iLon = colIndex(headers, ["longitude", "lon", "lng"]);
  const iCat = colIndex(headers, ["category", "report_category"]);
  const iDesc = colIndex(headers, ["description", "report_description", "desc", "observations"]);
  // Spec: support every filename column alias the master file may carry
  // so a customer's column heading (e.g. "photo", "image_file") still
  // resolves to the row's image filename.
  const iFile = colIndex(headers, [
    "file_name",
    "filename",
    "file",
    "name",
    "image",
    "image_name",
    "photo",
    "photo_name",
    "image_file",
    "photo_file",
  ]);
  // Spec: image_key column may also be labelled photo_key.
  const iImgKey = colIndex(headers, [
    "image_key",
    "imagekey",
    "img_key",
    "imgkey",
    "photo_key",
    "photokey",
  ]);
  const iAction = colIndex(headers, [
    "action",
    "actions",
    "difficulty",
    "route_difficulty",
    "remarks_action",
    "remarks/action",
    "remarks_/_action",
    "vehicle_movement",
    "movement",
    "remarks_action_status",
  ]);

  if (iKey < 0 || iCat < 0 || iLat < 0 || iLon < 0) {
    throw new Error(
      "Master file must include separate columns: point_key, latitude, longitude, and category"
    );
  }

  const out: ParsedCombinedRow[] = [];

  for (const r of rows) {
    const point_key = (r[iKey] ?? "").trim();
    if (!point_key) continue;

    const latRaw = (r[iLat] ?? "").trim();
    const lonRaw = (r[iLon] ?? "").trim();

    const latitude = parseSingleCoordinate(latRaw, "lat");
    const longitude = parseSingleCoordinate(lonRaw, "lon");

    if (!Number.isFinite(latitude) || !Number.isFinite(longitude)) continue;

    const rawAction = iAction >= 0 ? (r[iAction] ?? "").trim() : "";

    out.push({
      point_key,
      latitude: latitude as number,
      longitude: longitude as number,
      category: normalizeCategoryName((r[iCat] ?? "").trim()),
      description: iDesc >= 0 ? ((r[iDesc] ?? "").trim() || null) : null,
      file_name: iFile >= 0 ? ((r[iFile] ?? "").trim() || null) : null,
      image_key: iImgKey >= 0 ? ((r[iImgKey] ?? "").trim() || null) : null,
      difficulty: normalizeDifficulty(rawAction),
      remarks_action: rawAction || null,
    });
  }

  return out;
}

async function getImageSize(
  file: File
): Promise<{ width: number | null; height: number | null }> {
  try {
    const bmp = await createImageBitmap(file);
    const width = bmp.width ?? null;
    const height = bmp.height ?? null;
    bmp.close?.();
    return { width, height };
  } catch {
    return { width: null, height: null };
  }
}

// Replaces the old Supabase Storage uploader. Posts the file (and reportId)
// to /api/upload — the route streams to S3 and ALSO inserts a row into
// report_photos in the same transaction, so callers do NOT need a second
// /api/reports/[id]/photos call. Reuses session cookies for auth.
async function uploadFileToS3(args: {
  file: File;
  reportId: string;
  width?: number | null;
  height?: number | null;
  pointKey?: string | null;
  imageKey?: string | null;
  projectId?: string | null;
  fileName?: string | null;
}) {
  // Spec Part 2: log BEFORE the request so we can correlate the upload
  // attempt with whatever /api/upload reports back.
  console.log("[bulk] calling /api/upload", {
    projectId: args.projectId || null,
    reportId: args.reportId,
    pointKey: args.pointKey || null,
    imageKey: args.imageKey || null,
    fileName: args.fileName || args.file.name,
  });
  const fd = new FormData();
  fd.append("file", args.file);
  fd.append("reportId", args.reportId);
  if (args.width != null) fd.append("width", String(args.width));
  if (args.height != null) fd.append("height", String(args.height));
  if (args.pointKey) fd.append("pointKey", String(args.pointKey));
  if (args.imageKey) fd.append("imageKey", String(args.imageKey));
  if (args.projectId) fd.append("projectId", String(args.projectId));
  // fileName is sent so the server has the original (unsanitised) name
  // even if the multipart filename gets mangled by an intermediate proxy.
  fd.append("fileName", String(args.fileName || args.file.name));

  const res = await fetch("/api/upload", {
    method: "POST",
    body: fd,
    credentials: "include",
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`/api/upload ${res.status}: ${txt || res.statusText}`);
  }
  const json = (await res.json()) as {
    url: string;
    key: string;
    fileName?: string;
    reportPhoto?: {
      saved: boolean;
      reason?: string;
      verifiedCount?: number;
      verifiedRows?: Array<Record<string, unknown>>;
      verifiedProjectId?: string | null;
    };
  };
  return { path: json.key, publicUrl: json.url, fileName: json.fileName, reportPhoto: json.reportPhoto };
}

async function mapLimit<T>(
  items: T[],
  limit: number,
  worker: (item: T, index: number) => Promise<void>
) {
  let nextIndex = 0;

  async function runner() {
    while (true) {
      const i = nextIndex++;
      if (i >= items.length) return;
      await worker(items[i], i);
    }
  }

  await Promise.all(Array.from({ length: Math.max(1, limit) }, () => runner()));
}

function getDuplicates(values: string[]) {
  const seen = new Set<string>();
  const dup = new Set<string>();

  for (const v of values) {
    const key = v.trim();
    if (!key) continue;
    if (seen.has(key)) dup.add(key);
    else seen.add(key);
  }

  return Array.from(dup).sort((a, b) => a.localeCompare(b));
}


/**
 * Spec-mandated multi-image splitter. A single Excel cell may carry
 * `IMG_001.jpg, IMG_002.jpg, IMG_003.jpg` for one observation row.
 * This helper splits on comma, trims, and drops empty entries so the
 * downstream uploader can iterate every filename and insert a separate
 * `report_photos` row for each.
 */
function splitMultiImageRefs(value: unknown): string[] {
  return String(value || "")
    .split(",")
    .map((v) => v.trim())
    .filter(Boolean);
}

/**
 * Spec-mandated normalisation for matching master-file `file_name` against
 * uploaded image filenames. trim → URI-decode → strip folder path → lowercase.
 * Returns "" for empty/invalid input so callers can use the result as a Map
 * key without false-positive matches.
 */
function normalizeFileName(value: unknown): string {
  if (!value) return "";
  let v = String(value);
  try {
    v = decodeURIComponent(v);
  } catch {
    // ignore malformed percent escapes
  }
  return (
    v
      .trim()
      .replace(/\\/g, "/")
      .split("/")
      .pop()
      ?.trim()
      .toLowerCase() || ""
  );
}

function normalizeFileKey(value: string) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\\/g, "/")
    .split("/")
    .pop()
    ?.replace(/\s+/g, " ")
    .trim() || "";
}


export default function ProjectsPage() {
  const router = useRouter();

  const [projects, setProjects] = useState<ProjectRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [q, setQ] = useState("");
  const [exporting, setExporting] = useState(false);

  const [lastModifiedMap, setLastModifiedMap] = useState<Record<string, string>>({});

  const [newOpen, setNewOpen] = useState(false);
  const [newName, setNewName] = useState("");
  const [newDesc, setNewDesc] = useState("");
  const [creating, setCreating] = useState(false);

  const [editOpen, setEditOpen] = useState(false);
  const [editingProjectId, setEditingProjectId] = useState<string>("");
  const [editName, setEditName] = useState("");
  const [editDesc, setEditDesc] = useState("");
  const [savingEdit, setSavingEdit] = useState(false);
  const [deletingProjectId, setDeletingProjectId] = useState<string>("");

  const [bulkOpen, setBulkOpen] = useState(false);
  const [bulkProjectId, setBulkProjectId] = useState<string>("");
  const [importing, setImporting] = useState(false);

  const [masterFile, setMasterFile] = useState<File | null>(null);
  const [imageFiles, setImageFiles] = useState<File[]>([]);

  const masterInputRef = useRef<HTMLInputElement | null>(null);
  const imagesInputRef = useRef<HTMLInputElement | null>(null);

  const [summary, setSummary] = useState<ImportSummary | null>(null);
  const [masterPreview, setMasterPreview] = useState<string>("");
  const [previewRows, setPreviewRows] = useState<Record<string, any>[]>([]);
  const [previewFileName, setPreviewFileName] = useState<string>("master_preview.xlsx");
  const [previewReady, setPreviewReady] = useState(false);

  const filtered = useMemo(() => {
    const s = q.trim().toLowerCase();
    if (!s) return projects;
    return projects.filter((p) => JSON.stringify(p).toLowerCase().includes(s));
  }, [projects, q]);

  const safeName = (p: ProjectRow) =>
    p.name || p.title || p.project_name || "Untitled Project";

  const redirectToLogin = () => {
    router.replace("/login");
  };

  const getAuthUser = async () => {
    const token =
      typeof window !== "undefined" ? localStorage.getItem("auth_token") : null;
    const res = await fetch("/api/auth/me", {
      method: "GET",
      credentials: "include",
      headers: token ? { Authorization: `Bearer ${token}` } : {},
    });
    if (!res.ok) return null;
    const data = await res.json().catch(() => ({} as any));
    return data?.user || null;
  };

  const authHeaders = () => {
    const token =
      typeof window !== "undefined" ? localStorage.getItem("auth_token") : null;
    return token ? { Authorization: `Bearer ${token}` } : {};
  };

  const hydrateLastModifiedNames = async (rows: ProjectRow[]) => {
    const fallback: Record<string, string> = {};
    rows.forEach((p) => {
      if (p.last_modified_by) fallback[p.id] = p.last_modified_by.slice(0, 8);
    });
    setLastModifiedMap(fallback);
  };

  const load = async () => {
    setLoading(true);
    setLoadError(null);
    try {
      const res = await fetch("/api/projects", {
        method: "GET",
        credentials: "include",
        headers: authHeaders(),
      });
      const data = await res.json().catch(() => ({} as any));

      if (!res.ok) {
        throw new Error(data?.error || "Failed to fetch projects");
      }

      const rows = ((data?.projects || data || []) as ProjectRow[]).slice();
      rows.sort((a, b) => {
        const ta = a?.created_at ? new Date(a.created_at).getTime() : 0;
        const tb = b?.created_at ? new Date(b.created_at).getTime() : 0;
        return tb - ta;
      });
      setProjects(rows);
      await hydrateLastModifiedNames(rows);

      if (!bulkProjectId && rows.length) setBulkProjectId(rows[0].id);
    } catch (e: any) {
      const msg = String(e?.message || e || "").toLowerCase();
      if (msg.includes("auth") || msg.includes("session") || msg.includes("jwt")) {
        redirectToLogin();
        return;
      }
      setLoadError(e?.message || "Failed to fetch projects");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    const init = async () => {
      try {
        const user = await getAuthUser();
        if (!user) {
          redirectToLogin();
          return;
        }

        await load();
      } catch {
        redirectToLogin();
      }
    };

    init();
  }, []);

  const logout = async () => {
    try {
      await fetch("/api/auth/logout", {
        method: "POST",
        credentials: "include",
      });
    } catch {}
    if (typeof window !== "undefined") {
      localStorage.removeItem("auth_token");
    }
    redirectToLogin();
  };

  const exportCSV = async () => {
    if (!filtered.length) return;
    try {
      setExporting(true);
      await downloadProjectsCSV(filtered);
    } finally {
      setExporting(false);
    }
  };

  const createProject = async () => {
    const name = newName.trim();
    if (!name) return alert("Project name is required");

    try {
      setCreating(true);
      const res = await fetch("/api/projects", {
        method: "POST",
        credentials: "include",
        headers: {
          "Content-Type": "application/json",
          ...authHeaders(),
        },
        body: JSON.stringify({
          name,
          description: newDesc.trim() || null,
        }),
      });
      const data = await res.json().catch(() => ({} as any));
      if (!res.ok) throw new Error(data?.error || "Failed to create project");

      setNewOpen(false);
      setNewName("");
      setNewDesc("");
      await load();
    } catch (e: any) {
      const msg = String(e?.message || e || "").toLowerCase();
      if (msg.includes("auth") || msg.includes("session") || msg.includes("jwt")) {
        redirectToLogin();
        return;
      }
      alert(e?.message || String(e));
    } finally {
      setCreating(false);
    }
  };

  const openEditModal = (project: ProjectRow) => {
    setEditingProjectId(project.id);
    setEditName(safeName(project));
    setEditDesc(project.description || "");
    setEditOpen(true);
  };

  const updateProject = async () => {
    const name = editName.trim();
    if (!editingProjectId) return;
    if (!name) return alert("Project name is required");

    try {
      setSavingEdit(true);
      const res = await fetch(`/api/projects/${encodeURIComponent(editingProjectId)}`, {
        method: "PUT",
        credentials: "include",
        headers: {
          "Content-Type": "application/json",
          ...authHeaders(),
        },
        body: JSON.stringify({
          name,
          description: editDesc.trim() || null,
        }),
      });
      const data = await res.json().catch(() => ({} as any));
      if (!res.ok) throw new Error(data?.error || "Failed to update project");

      setEditOpen(false);
      setEditingProjectId("");
      setEditName("");
      setEditDesc("");
      await load();
    } catch (e: any) {
      const msg = String(e?.message || e || "").toLowerCase();
      if (msg.includes("auth") || msg.includes("session") || msg.includes("jwt")) {
        redirectToLogin();
        return;
      }
      alert(e?.message || String(e));
    } finally {
      setSavingEdit(false);
    }
  };

  const deleteProject = async (project: ProjectRow) => {
    const projectName = safeName(project);
    const confirmed = window.confirm(
      `Delete project \"${projectName}\"? This will also delete related reports, photos, path points, and bulk import history for this project.`
    );

    if (!confirmed) return;

    try {
      setDeletingProjectId(project.id);
      const res = await fetch(`/api/projects/${encodeURIComponent(project.id)}`, {
        method: "DELETE",
        credentials: "include",
        headers: authHeaders(),
      });
      const data = await res.json().catch(() => ({} as any));
      if (!res.ok) throw new Error(data?.error || "Failed to delete project");

      if (bulkProjectId === project.id) {
        setBulkProjectId("");
      }

      await load();
      alert("Project deleted successfully.");
    } catch (e: any) {
      const msg = String(e?.message || e || "").toLowerCase();
      if (msg.includes("auth") || msg.includes("session") || msg.includes("jwt")) {
        redirectToLogin();
        return;
      }
      alert(e?.message || String(e));
    } finally {
      setDeletingProjectId("");
    }
  };

  const resetBulk = () => {
    setMasterFile(null);
    setImageFiles([]);
    setSummary(null);
    setMasterPreview("");
    setPreviewRows([]);
    setPreviewFileName("master_preview.xlsx");
    setPreviewReady(false);

    if (masterInputRef.current) masterInputRef.current.value = "";
    if (imagesInputRef.current) imagesInputRef.current.value = "";
  };

  const getLocationName = async (lat: number, lng: number): Promise<string> => {
    try {
      const res = await fetch(
        `https://nominatim.openstreetmap.org/reverse?format=jsonv2&lat=${encodeURIComponent(
          lat
        )}&lon=${encodeURIComponent(
          lng
        )}&addressdetails=1&zoom=18`,
        {
          headers: {
            Accept: "application/json",
          },
        }
      );

      if (!res.ok) return `${lat}, ${lng}`;

      const data = await res.json();
      const addr = data?.address || {};

      const houseNumber = addr.house_number || "";
      const road =
        addr.road ||
        addr.pedestrian ||
        addr.footway ||
        addr.cycleway ||
        addr.path ||
        addr.residential ||
        "";
      const suburb =
        addr.suburb ||
        addr.neighbourhood ||
        addr.hamlet ||
        addr.village ||
        "";
      const city =
        addr.city ||
        addr.town ||
        addr.municipality ||
        addr.county ||
        addr.state_district ||
        "";
      const state = addr.state || "";
      const postcode = addr.postcode || "";

      const line1 = [houseNumber, road].filter(Boolean).join(" ").trim();

      const parts = [line1, suburb, city, state, postcode].filter(
        (value, index, arr) => value && arr.indexOf(value) === index
      );

      if (parts.length) return parts.join(", ");

      return data?.display_name || `${lat}, ${lng}`;
    } catch {
      return `${lat}, ${lng}`;
    }
  };

  const buildPreviewRows = async (rows: ParsedCombinedRow[]) => {
    const sorted = [...rows].sort((a, b) => {
      const na = Number(a.point_key);
      const nb = Number(b.point_key);
      if (Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
      return String(a.point_key).localeCompare(String(b.point_key));
    });

    let cumulativeKm = 0;
    const previewRowsResult: Array<Record<string, any>> = [];

    for (let index = 0; index < sorted.length; index += 1) {
      const row = sorted[index];

      let legKm = 0;

      if (index > 0) {
        const prev = sorted[index - 1];
        legKm = await getRoadDistanceKm(
          prev.latitude,
          prev.longitude,
          row.latitude,
          row.longitude
        );
        cumulativeKm += legKm;
      }

      const location = await getLocationName(row.latitude, row.longitude);

      previewRowsResult.push({
        sl_no: index + 1,
        point_key: row.point_key,
        latitude: row.latitude,
        longitude: row.longitude,
        location,
        leg_kms: index === 0 ? "0.0000" : legKm.toFixed(4),
        kms: cumulativeKm.toFixed(4),
        category: row.category || "",
        description: row.description || "",
        file_name: row.file_name || "",
        image_key: row.image_key || "",
        difficulty: row.difficulty || "",
        remarks_action: row.remarks_action || "",
      });
    }

    return previewRowsResult;
  };

  const downloadPreviewExcel = () => {
    if (!previewRows.length) {
      alert("No preview data available.");
      return;
    }

    const worksheet = XLSX.utils.json_to_sheet(previewRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Preview");
    XLSX.writeFile(workbook, previewFileName || "master_preview.xlsx");
  };

  const handleMasterFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] ?? null;
    setMasterFile(file);
    setSummary(null);
    setPreviewRows([]);
    setPreviewReady(false);

    if (!file) {
      setMasterPreview("");
      setPreviewFileName("master_preview.xlsx");
      return;
    }

    try {
      if (isExcelFile(file)) {
        const { headers, rows } = await readStructuredFile(file);
        const previewLines = [
          headers.join(", "),
          ...rows.slice(0, 5).map((r) => r.join(", ")),
        ];
        setMasterPreview(previewLines.join("\n"));
      } else {
        const text = await readTextFile(file);
        setMasterPreview(text.split(/\r?\n/).slice(0, 6).join("\n"));
      }

      const combinedRows = await parseCombinedFile(file);
      const rowsForPreview = await buildPreviewRows(combinedRows);

      if (!rowsForPreview.length) {
        throw new Error(
          "No valid rows found to build preview. Check point_key, latitude, longitude, and category columns."
        );
      }

      setPreviewRows(rowsForPreview);
      setPreviewReady(true);
      setPreviewFileName(
        `${file.name.replace(/\.[^.]+$/, "") || "master"}_preview.xlsx`
      );
    } catch (err: any) {
      setMasterPreview("");
      setPreviewRows([]);
      setPreviewReady(false);
      alert(err?.message || String(err));
    }
  };

  const handleImagesChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    setImageFiles(files);
    setSummary(null);
  };

  const runImportSingleMaster = async () => {
    if (!bulkProjectId) return alert("Select a project.");
    if (!masterFile) return alert("Select master file.");
    // imageFiles MAY be empty: the server-side S3 lookup will then try to
    // resolve the master file's file_name / image_key references against
    // images already uploaded to the bucket. The post-import summary will
    // tell the user how many references could not be resolved.

    // Spec-mandated entry log so we can verify, before anything else runs,
    // exactly which master file + image files the user has selected for
    // THIS project. If imageFilesCount is 0 here, the per-row /api/upload
    // step further down will have nothing to send — the only path to
    // report_photos for this project would then be the server-side S3
    // lookup.
    console.log("[BULK FRONTEND IMAGE SEND]", {
      projectId: bulkProjectId,
      masterFile: masterFile?.name || null,
      imageFilesCount: imageFiles?.length || 0,
      sampleImages: Array.from(imageFiles || []).slice(0, 10).map((f) => ({
        name: f.name,
        size: f.size,
        type: f.type,
      })),
    });

    setImporting(true);
    setSummary(null);

    const errors: string[] = [];

    try {
      const user = await getAuthUser();
      if (!user) {
        redirectToLogin();
        return;
      }

      // Master-file dedup history was a Supabase-only safety net. Skip it on
      // the MySQL stack — the bulk-import API upserts by (project_id, point_key)
      // so re-importing the same file is now safe (each point updates in place).

      const combinedRows = await parseCombinedFile(masterFile);
      console.log("[bulk import] parsed rows:", combinedRows.length);
      console.log("[bulk import] selected images:", imageFiles.length);
      console.log("[bulk import] importing project:", bulkProjectId);

      // Spec-mandated trace: dump the first 10 parsed master rows so we can
      // confirm point_key / file_name / image_key / category survived parse.
      console.log("[BULK PHOTO TRACE parsed rows]", {
        totalRows: combinedRows.length,
        sampleRows: combinedRows.slice(0, 10).map((r: any) => ({
          point_key: r.point_key,
          file_name: r.file_name,
          image_key: r.image_key,
          category: r.category,
        })),
      });

      // Spec-mandated trace: dump the first 20 manually selected image files
      // so we can correlate them against master-file file_name values when
      // matching fails.
      console.log("[BULK PHOTO TRACE uploaded files]", {
        uploadedFilesCount: imageFiles.length,
        sampleFiles: imageFiles.slice(0, 20).map((f) => ({
          name: f.name,
          size: f.size,
          type: f.type,
          normalizedName: normalizeFileName(f.name),
        })),
      });

      if (!combinedRows.length) {
        throw new Error(
          "No valid rows found in master file. Check point_key, latitude, longitude, category, and file values."
        );
      }

      const incomingPointKeys = Array.from(
        new Set(
          combinedRows
            .map((row) => String(row.point_key || "").trim())
            .filter(Boolean)
        )
      );

      if (!incomingPointKeys.length) {
        throw new Error("No valid point_key values found in master file.");
      }

      // Build the per-point summary row (one report per unique point_key) and
      // the file→point_key image map (multiple files can map to one point).
      const pointsMap = new Map<string, ParsedPointRow>();
      const imageMap: ParsedImageMapRow[] = [];

      combinedRows.forEach((row) => {
        if (!pointsMap.has(row.point_key)) {
          pointsMap.set(row.point_key, {
            point_key: row.point_key,
            latitude: row.latitude,
            longitude: row.longitude,
            category: row.category || "Unknown",
            normalizedCategory: row.category || "Unknown",
            description: row.description || null,
            difficulty: row.difficulty || "green",
            remarks_action: row.remarks_action || null,
          });
        }

        if (row.file_name) {
          // Spec: one Excel cell may carry comma-separated filenames.
          // Expand into one imageMap entry per filename so multi-image
          // rows produce one report_photos row per file.
          const splitNames = splitMultiImageRefs(row.file_name);
          const splitKeys = splitMultiImageRefs(row.image_key || "");
          if (splitNames.length === 0) {
            imageMap.push({
              file_name: row.file_name,
              point_key: row.point_key,
              image_key: row.image_key || null,
            });
          } else {
            for (let nIdx = 0; nIdx < splitNames.length; nIdx += 1) {
              imageMap.push({
                file_name: splitNames[nIdx],
                point_key: row.point_key,
                image_key: splitKeys[nIdx] || splitKeys[0] || null,
              });
            }
          }
        }
      });

      const points = Array.from(pointsMap.values());

      const duplicateSelectedImages = getDuplicates(imageFiles.map((f) => f.name));
      const duplicateMappingFiles = getDuplicates(imageMap.map((r) => r.file_name));
      const invalidCategories: string[] = [];

      if (duplicateSelectedImages.length) {
        errors.push(
          `Duplicate selected image names found: ${duplicateSelectedImages.join(", ")}`
        );
      }

      if (duplicateMappingFiles.length) {
        errors.push(
          `Duplicate file_name values found in master file: ${duplicateMappingFiles.join(", ")}`
        );
      }

      const pointByKey = new Map<string, ParsedPointRow>();
      points.forEach((p) => {
        pointByKey.set(p.point_key, {
          ...p,
          category: normalizeCategoryName(p.category),
          difficulty: normalizeDifficulty(p.difficulty || "green"),
          remarks_action: p.remarks_action || null,
        });
      });

      const missingPointKeysInPointsCsv = Array.from(
        new Set(
          imageMap
            .map((m) => m.point_key)
            .filter((pk): pk is string => !!pk && !pointByKey.has(pk))
        )
      ).sort((a, b) => {
        const na = Number(a);
        const nb = Number(b);
        if (Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
        return a.localeCompare(b);
      });

      if (missingPointKeysInPointsCsv.length) {
        throw new Error(
          `These point_key values are used in master file image mapping but missing in points data: ${missingPointKeysInPointsCsv.join(", ")}`
        );
      }

      // ---- Step A: send all unique points to the bulk-import API in one call.
      const sortedKeys = Array.from(pointByKey.keys()).sort((a, b) => {
        const na = Number(a);
        const nb = Number(b);
        if (Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
        return a.localeCompare(b);
      });

      // Build per-point image refs from imageMap so the server can do its
      // own S3 lookup when the user did not manually select files. Multiple
      // file rows in the master can map to one point_key.
      const refsByPointKey = new Map<
        string,
        Array<{ file_name: string | null; image_key: string | null }>
      >();
      for (const m of imageMap) {
        if (!m.point_key) continue;
        if (!m.file_name && !m.image_key) continue;
        const list = refsByPointKey.get(m.point_key) || [];
        list.push({ file_name: m.file_name || null, image_key: m.image_key || null });
        refsByPointKey.set(m.point_key, list);
      }

      const apiRows = sortedKeys.map((key) => {
        const p = pointByKey.get(key)!;
        const refs = refsByPointKey.get(key) || [];
        return {
          point_key: key,
          latitude: p.latitude,
          longitude: p.longitude,
          category: p.category.trim() || "Unknown",
          description: p.description?.trim() || null,
          difficulty: normalizeDifficulty(p.difficulty || "green"),
          remarks_action: p.remarks_action?.trim() || null,
          // Pass the master-file file/key references so the server can resolve
          // them against S3 directly. file_name is the first ref's filename
          // (kept for the existing manual-upload code path), image_refs[]
          // covers all of them.
          file_name: refs[0]?.file_name || null,
          image_key: refs[0]?.image_key || null,
          image_refs: refs,
        };
      });

      const importRes = await fetch(
        `/api/projects/${encodeURIComponent(bulkProjectId)}/bulk-import`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          credentials: "include",
          body: JSON.stringify({
            rows: apiRows,
            // "S3 (configured server-side)" Storage option => always ask the
            // server to scan the bucket for any references the user did not
            // upload manually in this run.
            s3Lookup: {
              enabled: true,
              prefixes: [
                "reports/photos/",
                `reports/${bulkProjectId}/`,
                `reports/photos/${bulkProjectId}/`,
              ],
            },
          }),
        }
      );
      if (!importRes.ok) {
        const txt = await importRes.text().catch(() => "");
        throw new Error(`Bulk import API ${importRes.status}: ${txt || importRes.statusText}`);
      }
      const importJson = (await importRes.json()) as {
        ok: boolean;
        insertedCount: number;
        updatedCount: number;
        photosMatched?: number;
        photosMissing?: number;
        photosMissingSamples?: string[];
        s3LookupEnabled?: boolean;
        s3IndexSize?: number;
        reports: Array<{ point_key: string; report_id: string }>;
      };

      const reportIdByPointKey = new Map<string, string>();
      for (const r of importJson.reports) {
        reportIdByPointKey.set(r.point_key, r.report_id);
      }
      console.log("[bulk import] api result:", {
        inserted: importJson.insertedCount,
        updated: importJson.updatedCount,
        reportCount: reportIdByPointKey.size,
      });

      // ---- Step B: NO_GPS holding report. Created lazily via the same
      // bulk-import API with a sentinel point_key so it survives re-runs.
      const NO_GPS_KEY = "__NO_GPS__";
      let noGpsReportId: string | null = null;
      async function getOrCreateNoGpsReport() {
        if (noGpsReportId) return noGpsReportId;
        const r = await fetch(
          `/api/projects/${encodeURIComponent(bulkProjectId)}/bulk-import`,
          {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            credentials: "include",
            body: JSON.stringify({
              rows: [
                {
                  point_key: NO_GPS_KEY,
                  latitude: null,
                  longitude: null,
                  category: "NO_GPS Images",
                  description:
                    "Images that do not have GPS point mapping (bulk import).",
                  difficulty: "NO_GPS",
                },
              ],
            }),
          }
        );
        if (!r.ok) {
          const t = await r.text().catch(() => "");
          throw new Error(`NO_GPS report create failed: ${r.status} ${t}`);
        }
        const j = (await r.json()) as {
          reports: Array<{ point_key: string; report_id: string }>;
        };
        const found = j.reports.find((x) => x.point_key === NO_GPS_KEY);
        if (!found) throw new Error("NO_GPS report not returned by API");
        noGpsReportId = found.report_id;
        return noGpsReportId;
      }

      // ---- Step C: per-row chain. For every parsed master row:
      //    1. resolve the saved reportId from reportIdByPointKey,
      //    2. normalise the master row's file_name and look it up in
      //       uploadedImageMap (keyed by normalised filename),
      //    3. upload the matched File to S3 via /api/upload, which inserts
      //       the report_photos row using the EXACT reportId we pass.
      // The bulk-import API also performs server-side S3 lookup for any rows
      // we did NOT manually upload here, so this loop only needs to handle
      // rows whose images live in `imageFiles` for this run.
      const uploadedImageMap = new Map<string, File>();
      for (const file of imageFiles) {
        const normalized = normalizeFileName(file.name);
        if (normalized && !uploadedImageMap.has(normalized)) {
          uploadedImageMap.set(normalized, file);
        }
      }

      // Diagnostics retained for the existing UI summary.
      const mapByFileName = new Map<string, ParsedImageMapRow>();
      imageMap.forEach((r) => {
        const k = normalizeFileName(r.file_name);
        if (k && !mapByFileName.has(k)) mapByFileName.set(k, r);
      });
      const missingFilesInUpload: string[] = [];
      for (const [name] of mapByFileName) {
        if (!uploadedImageMap.has(name)) missingFilesInUpload.push(name);
      }
      const extraFilesNotInMap: string[] = [];
      imageFiles.forEach((f) => {
        if (!mapByFileName.has(normalizeFileName(f.name))) {
          extraFilesNotInMap.push(f.name);
        }
      });

      // Spec-required tallies.
      let rowsWithFileName = 0;
      let matchedImagesCount = 0;
      let uploadedToS3Count = 0;
      let reportPhotosInsertedClient = 0;
      let missingImagesCount = 0;
      const noGpsImages: string[] = [];
      let imagesUploaded = 0;
      let photosInserted = 0; // legacy: still surfaced in the existing UI summary

      // Iterate the parsed master rows (one per point_key) so the chain is
      // strictly: save report → match image → upload → insert report_photos
      // keyed by THIS reportId. Concurrency-limited so we don't open hundreds
      // of S3 connections at once on a 400-row import.
      await mapLimit(combinedRows, 5, async (row) => {
        const point_key = row.point_key;
        // Spec Part 3: support every filename column alias the master file
        // may carry. parseCombinedFile already collapses these into
        // row.file_name, but we re-cascade defensively in case a row was
        // assembled from a different parser path.
        const r = row as Record<string, unknown>;
        const rawFileName =
          (r.file_name as string | null) ||
          (r.filename as string | null) ||
          (r.image as string | null) ||
          (r.image_name as string | null) ||
          (r.photo as string | null) ||
          (r.photo_name as string | null) ||
          (r.image_file as string | null) ||
          (r.photo_file as string | null) ||
          "";
        const normalizedFileName = normalizeFileName(rawFileName);
        if (rawFileName) rowsWithFileName += 1;

        console.log("[UPLOAD parsed row]", {
          point_key,
          file_name: rawFileName,
          normalizedFileName,
          image_key: row.image_key,
        });

        // Spec: split comma-separated filenames so multi-image rows
        // upload one file per piece.
        const splitFileNames = splitMultiImageRefs(rawFileName);
        const splitImageKeys = splitMultiImageRefs(row.image_key || "");
        const fileNamesToTry = splitFileNames.length > 0 ? splitFileNames : [rawFileName];

        // Match each split filename against the uploaded files map.
        const matches: Array<{
          rawName: string;
          normName: string;
          file: File;
          imageKey: string | null;
        }> = [];
        for (let mIdx = 0; mIdx < fileNamesToTry.length; mIdx += 1) {
          const candidateRaw = fileNamesToTry[mIdx];
          const candidateNorm = normalizeFileName(candidateRaw);
          const candidateFile = candidateNorm
            ? uploadedImageMap.get(candidateNorm) || null
            : null;

          // Spec Part 3 log key — fires per split filename.
          console.log("[bulk] image match", {
            pointKey: point_key,
            rawFileName: candidateRaw,
            normalized: candidateNorm,
            matched: !!candidateFile,
            matchedFileName: candidateFile?.name || null,
          });

          // Spec-mandated multi-image match log.
          console.log("[BULK MULTI IMAGE MATCH]", {
            projectId: bulkProjectId,
            point_key,
            imgIndex: mIdx,
            rawFileName: candidateRaw,
            rawImageKey: splitImageKeys[mIdx] || splitImageKeys[0] || null,
            matched: !!candidateFile,
            matchedFileName: candidateFile?.name || null,
          });

          if (candidateFile) {
            matches.push({
              rawName: candidateRaw,
              normName: candidateNorm,
              file: candidateFile,
              imageKey: splitImageKeys[mIdx] || splitImageKeys[0] || null,
            });
          }
        }

        if (matches.length > 0) matchedImagesCount += matches.length;
        const matchedFile = matches[0]?.file || null;

        // Resolve the EXACT reports.id for this point. If the master row had
        // no GPS we fall through to the NO_GPS placeholder report so the
        // photo still has a parent.
        let reportId: string | undefined = reportIdByPointKey.get(point_key);
        if (!reportId) {
          if (matchedFile) {
            try {
              reportId = await getOrCreateNoGpsReport();
              noGpsImages.push(matchedFile.name);
            } catch (err) {
              console.error("[UPLOAD report saved] NO_GPS fallback failed:", err);
              return;
            }
          } else {
            // No file AND no report saved - the server-side S3 lookup will
            // try to find this row's file_name in the bucket. Nothing more
            // for the client to do here.
            if (row.file_name) missingImagesCount += 1;
            return;
          }
        }

        console.log("[UPLOAD report saved]", {
          point_key,
          reportId,
        });

        if (matches.length === 0) {
          if (row.file_name) {
            console.warn("[UPLOAD photo missing - no manual file]", {
              point_key,
              reportId,
              file_name: row.file_name,
            });
            // Spec-mandated skip log: makes "no match" failures visible
            // in the same key-space as the success path.
            console.warn("[BULK PHOTO SKIPPED]", {
              projectId: bulkProjectId,
              point_key,
              reportId,
              rawFileName: row.file_name,
              reason: "no matched uploaded file",
            });
            missingImagesCount += 1;
          }
          return;
        }

        // Upload every matched file. /api/upload dedups by
        // (report_id, file_name) so distinct filenames produce distinct
        // report_photos rows for the same report — multi-image support
        // without losing previously uploaded files.
        for (let upIdx = 0; upIdx < matches.length; upIdx += 1) {
          const m = matches[upIdx];
          // Spec-mandated per-row debug. Prove BEFORE the upload runs that
          // the file we matched and the savedReportId we are about to link
          // it to are exactly what we expect.
          console.log("[BULK PHOTO ROW DEBUG]", {
            point_key,
            file_name: m.rawName,
            normalizedFileName: m.normName,
            matchedFileName: m.file.name,
            reportId,
            multiImageIndex: upIdx,
            multiImageTotal: matches.length,
          });

          const { width, height } = await getImageSize(m.file);
          try {
            const uploaded = await uploadFileToS3({
              file: m.file,
              reportId,
              width,
              height,
              pointKey: point_key,
              imageKey: m.imageKey,
              projectId: bulkProjectId,
              fileName: m.rawName || m.file.name,
            });
          imagesUploaded += 1;
          uploadedToS3Count += 1;
          if (uploaded.reportPhoto?.saved) {
            photosInserted += 1;
            reportPhotosInsertedClient += 1;
          }

          console.log("[UPLOAD report_photo inserted]", {
            report_id: reportId,
            point_key,
            file_name: row.file_name,
            image_key: row.image_key,
            url: uploaded.publicUrl,
          });

          // Spec-mandated post-insert verify trace. Echoes the verify rows
          // /api/upload returned so the operator can confirm count > 0
          // for THIS reportId AND that the parent report belongs to the
          // current project (verifiedProjectId === bulkProjectId).
          const verifiedProjectId = uploaded.reportPhoto?.verifiedProjectId ?? null;
          const projectIdMatches =
            verifiedProjectId !== null &&
            String(verifiedProjectId) === String(bulkProjectId);
          console.log("[BULK REPORT_PHOTOS VERIFY]", {
            expectedProjectId: bulkProjectId,
            verifiedProjectId,
            projectIdMatches,
            reportId,
            point_key,
            count: uploaded.reportPhoto?.verifiedCount ?? 0,
            rows: uploaded.reportPhoto?.verifiedRows ?? null,
          });
          if (verifiedProjectId !== null && !projectIdMatches) {
            console.error("[BULK REPORT_PHOTOS VERIFY] PROJECT MISMATCH", {
              expectedProjectId: bulkProjectId,
              actualProjectId: verifiedProjectId,
              reportId,
              point_key,
            });
          }
        } catch (upErr: any) {
          errors.push(`${m.file.name}: upload failed - ${upErr?.message || String(upErr)}`);
          missingImagesCount += 1;
          console.error("[UPLOAD photo missing - upload threw]", {
            point_key,
            reportId,
            file_name: m.rawName,
            multiImageIndex: upIdx,
            error: upErr?.message,
          });
        }
        } // end for matches
      });

      console.log("[UPLOAD import summary]", {
        totalRows: combinedRows.length,
        rowsWithFileName,
        uploadedImagesCount: imageFiles.length,
        matchedImagesCount,
        uploadedToS3Count,
        reportPhotosInserted: reportPhotosInsertedClient,
        missingImagesCount,
      });

      // Spec-mandated final summary in the [BULK PHOTO ...] key-space.
      // Aggregates the per-row outcomes so a single log line tells the
      // operator whether the import populated report_photos for THIS
      // project. If insertedReportPhotos is 0 here, the post-import SQL
      // JOIN check will also be 0 and the DOCX export cannot show photos.
      console.log("[BULK PHOTO FINAL SUMMARY]", {
        projectId: bulkProjectId,
        totalRows: combinedRows.length,
        imagesReceived: imageFiles.length,
        rowsWithFileName,
        matchedImages: matchedImagesCount,
        uploadedToS3: uploadedToS3Count,
        insertedReportPhotos: reportPhotosInsertedClient,
        skippedNoMatch: missingImagesCount,
      });

      if (imageFiles.length > 0 && photosInserted === 0) {
        errors.push(
          "No report_photos rows were inserted. Most likely the uploaded file names do not exactly match the master file file_name values."
        );
      }

      // Server-side S3 lookup outcome.
      const refRows = combinedRows.filter((row) => row.file_name || row.image_key).length;
      const photosMatchedServer = importJson.photosMatched || 0;
      const photosMissingServer = importJson.photosMissing || 0;
      const totalPhotosResolved = photosInserted + photosMatchedServer;

      if (refRows > 0 && imageFiles.length === 0 && photosMatchedServer === 0) {
        // Hard warning per spec: master file references images, user selected
        // none, and the server-side S3 lookup matched none either.
        errors.push(
          `${refRows} rows reference image filenames, but 0 images were selected. Please select the image files or upload them to S3 first so the server-side lookup can find them.`
        );
      } else if (photosMissingServer > 0) {
        const sample = (importJson.photosMissingSamples || []).slice(0, 5).join(", ");
        errors.push(
          `${photosMissingServer} image reference${photosMissingServer === 1 ? "" : "s"} from the master file could not be found in S3${sample ? ` (e.g. ${sample})` : ""}.`
        );
      }

      setSummary({
        pointsRead: points.length,
        reportsCreatedOrUsed: sortedKeys.length + (noGpsReportId ? 1 : 0),
        imagesSelected: imageFiles.length,
        imagesUploaded,
        photosInserted,
        // New diagnostic fields - rendered alongside the existing summary,
        // never replacing it.
        photosMatchedServer,
        photosMissingServer,
        photosMissingSamples: importJson.photosMissingSamples || [],
        s3IndexSize: importJson.s3IndexSize || 0,
        totalPhotosResolved,
        masterRowsReferencingImages: refRows,
        reportsInserted: importJson.insertedCount,
        reportsUpdated: importJson.updatedCount,
        noGpsImages,
        missingFilesInUpload,
        extraFilesNotInMap,
        errors,
        duplicateSelectedImages,
        duplicateMappingFiles,
        missingPointKeysInPointsCsv,
        invalidCategories,
      });

      alert(
        errors.length
          ? "Bulk import completed with some warnings/errors. Check summary."
          : "Bulk import completed."
      );

      setBulkOpen(false);
      resetBulk();
      await load();
    } catch (e: any) {
      const msg = String(e?.message || e || "").toLowerCase();
      if (msg.includes("auth") || msg.includes("session") || msg.includes("jwt")) {
        redirectToLogin();
        return;
      }
      alert(e?.message || String(e));
    } finally {
      setImporting(false);
    }
  };

  return (
    <div style={styles.page}>
      <div style={styles.header}>
        <div>
          <div style={styles.title}>Projects</div>
          <div style={styles.subtitle}>
            Total: <b>{projects.length}</b> • Showing: <b>{filtered.length}</b>
          </div>
        </div>

        <div style={styles.headerRight}>
          <button
            style={{ ...styles.btnGhost, opacity: loading || exporting ? 0.7 : 1 }}
            onClick={load}
            disabled={loading || exporting}
          >
            {loading ? "Refreshing..." : "Refresh"}
          </button>

          <button style={styles.btnPrimary} onClick={() => setNewOpen(true)}>
            + New Project
          </button>

          <button style={styles.btnGhost} onClick={() => setBulkOpen(true)}>
            Bulk Import (Single Master File)
          </button>

          <div style={styles.exportGroup}>
            <button
              style={{ ...styles.btnPrimary, opacity: exporting ? 0.7 : 1 }}
              onClick={exportCSV}
              disabled={!filtered.length || exporting}
            >
              {exporting ? "Exporting..." : "Export CSV"}
            </button>
          </div>

          <button style={styles.btnDanger} onClick={logout}>
            Logout
          </button>
        </div>
      </div>

      <div style={styles.searchBar}>
        <div style={styles.searchWrap}>
          <span style={styles.searchIcon}>⌕</span>
          <input
            style={styles.searchInput}
            placeholder="Search projects by name, id, any field..."
            value={q}
            onChange={(e) => setQ(e.target.value)}
          />
          {q ? (
            <button style={styles.clearBtn} onClick={() => setQ("")}>
              Clear
            </button>
          ) : null}
        </div>
      </div>

      {loading ? (
        <div style={styles.stateBox}>Loading projects...</div>
      ) : loadError ? (
        <div style={styles.stateBox}>{loadError}</div>
      ) : filtered.length === 0 ? (
        <div style={styles.stateBox}>No projects found</div>
      ) : (
        <div style={styles.grid}>
          {filtered.map((p) => {
            const name = safeName(p);
            const dt = p.created_at ? new Date(p.created_at).toLocaleString() : "";
            const modifiedBy = lastModifiedMap[p.id] || "—";
            const isDeleting = deletingProjectId === p.id;

            return (
              <div key={p.id} style={styles.cardWrap}>
                <div
                  style={styles.card}
                  role="button"
                  tabIndex={0}
                  onClick={() => router.push(`/projects/${p.id}`)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter" || e.key === " ") {
                      e.preventDefault();
                      router.push(`/projects/${p.id}`);
                    }
                  }}
                >
                  <div style={styles.cardTop}>
                    <div style={styles.cardTitle}>{name}</div>
                    <span style={styles.badge}>Open</span>
                  </div>

                  <div style={styles.metaRow}>
                    <span style={styles.metaLabel}>Project ID</span>
                    <span style={styles.metaValue} title={p.id}>
                      {p.id}
                    </span>
                  </div>

                  <div style={styles.metaRow}>
                    <span style={styles.metaLabel}>Created</span>
                    <span style={styles.metaValue}>{dt || "—"}</span>
                  </div>

                  <div style={styles.metaRow}>
                    <span style={styles.metaLabel}>Last modified by</span>
                    <span style={styles.metaValue} title={modifiedBy}>
                      {modifiedBy}
                    </span>
                  </div>

                  <div style={styles.cardHint}>Click to view reports →</div>
                </div>

                <div style={styles.cardActions}>
                  <button
                    type="button"
                    style={styles.btnGhostSmall}
                    onClick={(e) => {
                      e.stopPropagation();
                      openEditModal(p);
                    }}
                  >
                    Edit
                  </button>

                  <button
                    type="button"
                    style={{ ...styles.btnDangerSmall, opacity: isDeleting ? 0.7 : 1 }}
                    disabled={isDeleting}
                    onClick={(e) => {
                      e.stopPropagation();
                      deleteProject(p);
                    }}
                  >
                    {isDeleting ? "Deleting..." : "Delete"}
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {newOpen && (
        <div style={styles.modalOverlay} onClick={() => setNewOpen(false)}>
          <div style={styles.modal} onClick={(e) => e.stopPropagation()}>
            <div style={{ fontSize: 18, fontWeight: 800, marginBottom: 10 }}>
              Create Project
            </div>

            <div style={styles.formRow}>
              <div style={styles.formLabel}>Name *</div>
              <input
                style={styles.input}
                value={newName}
                onChange={(e) => setNewName(e.target.value)}
                placeholder="Project name"
              />
            </div>

            <div style={styles.formRow}>
              <div style={styles.formLabel}>Description</div>
              <textarea
                style={styles.textarea}
                value={newDesc}
                onChange={(e) => setNewDesc(e.target.value)}
                placeholder="Optional description"
              />
            </div>

            <div style={styles.modalActions}>
              <button style={styles.btnGhost} onClick={() => setNewOpen(false)}>
                Cancel
              </button>
              <button
                style={{ ...styles.btnPrimary, opacity: creating ? 0.7 : 1 }}
                onClick={createProject}
                disabled={creating}
              >
                {creating ? "Creating..." : "Create"}
              </button>
            </div>
          </div>
        </div>
      )}

      {editOpen && (
        <div
          style={styles.modalOverlay}
          onClick={() => {
            if (!savingEdit) setEditOpen(false);
          }}
        >
          <div style={styles.modal} onClick={(e) => e.stopPropagation()}>
            <div style={{ fontSize: 18, fontWeight: 800, marginBottom: 10 }}>
              Edit Project
            </div>

            <div style={styles.formRow}>
              <div style={styles.formLabel}>Project name *</div>
              <input
                style={styles.input}
                value={editName}
                onChange={(e) => setEditName(e.target.value)}
                placeholder="Project name"
              />
            </div>

            <div style={styles.formRow}>
              <div style={styles.formLabel}>Description</div>
              <textarea
                style={styles.textarea}
                value={editDesc}
                onChange={(e) => setEditDesc(e.target.value)}
                placeholder="Optional description"
              />
            </div>

            <div style={styles.modalActions}>
              <button
                style={styles.btnGhost}
                onClick={() => setEditOpen(false)}
                disabled={savingEdit}
              >
                Cancel
              </button>
              <button
                style={{ ...styles.btnPrimary, opacity: savingEdit ? 0.7 : 1 }}
                onClick={updateProject}
                disabled={savingEdit}
              >
                {savingEdit ? "Saving..." : "Save Changes"}
              </button>
            </div>
          </div>
        </div>
      )}

      {bulkOpen && (
        <div style={styles.modalOverlay} onClick={() => setBulkOpen(false)}>
          <div style={styles.modalWide} onClick={(e) => e.stopPropagation()}>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                gap: 10,
              }}
            >
              <div style={{ fontSize: 18, fontWeight: 800 }}>
                Bulk Import (Single Master File + Images)
              </div>
              <button style={styles.btnGhost} onClick={resetBulk} disabled={importing}>
                Clear
              </button>
            </div>

            <div style={{ marginTop: 10, fontSize: 13, color: "#475467", lineHeight: 1.6 }}>
              <b>Single master file format:</b>
              <br />
              Required columns:
              <br />
              <b>point_key</b>, <b>latitude</b>, <b>longitude</b>, <b>category</b>
              <br />
              Optional columns:
              <br />
              <b>description</b>, <b>file_name</b>, <b>image_key</b>, <b>action</b> / <b>actions</b> / <b>difficulty</b>
              <br />
              Supports <b>CSV, TXT, XLSX, XLS</b>
              <br />
              <br />
              <b>Accepted latitude / longitude examples:</b>
              <br />
              <b>19.17189</b> and <b>72.54298</b>
              <br />
              <b>N19.17189</b> and <b>E72.54298</b>
              <br />
              <b>N19 10.313</b> and <b>E72 32.578</b>
              <br />
              <br />
              Preview file now includes:
              <br />
              <b>location</b> = improved reverse geocoded address
              <br />
              <b>leg_kms</b> = distance from previous point
              <br />
              <b>kms</b> = cumulative route distance
              <br />
              <br />
              One <b>point_key</b> inside one project can be imported only once.
            </div>

            <div
              style={{
                display: "grid",
                gridTemplateColumns: "1fr 1fr",
                gap: 10,
                marginTop: 12,
              }}
            >
              <div>
                <div style={styles.formLabel}>Select Project</div>
                <select
                  style={styles.input as any}
                  value={bulkProjectId}
                  onChange={(e) => setBulkProjectId(e.target.value)}
                >
                  {projects.map((p) => (
                    <option key={p.id} value={p.id}>
                      {safeName(p)} ({p.id.slice(0, 8)}…)
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <div style={styles.formLabel}>Storage</div>
                <div style={{ ...styles.input, display: "flex", alignItems: "center", color: "#666" }}>
                  S3 (configured server-side)
                </div>
              </div>
            </div>

            <div style={{ marginTop: 10 }}>
              <div style={styles.formLabel}>Category / Difficulty Handling</div>
              <div style={styles.categoryHelpBox}>
                Known categories are normalized automatically. New/custom categories are allowed. The <b>action / actions / difficulty</b> column is saved exactly as provided into both the <b>difficulty</b> field and the <b>remarks_action</b> field in <b>reports</b>.
              </div>
            </div>

            <div
              style={{
                display: "grid",
                gridTemplateColumns: "1fr 1fr",
                gap: 10,
                marginTop: 12,
              }}
            >
              <div>
                <div style={styles.formLabel}>Master file</div>
                <input
                  ref={masterInputRef}
                  type="file"
                  accept=".csv,.txt,.xlsx,.xls"
                  onChange={handleMasterFileChange}
                />
                {masterFile && (
                  <div style={styles.fileMeta}>
                    {masterFile.name} • {masterFile.size} bytes
                  </div>
                )}
                {masterPreview && <pre style={styles.previewBox}>{masterPreview}</pre>}
                {previewReady && (
                  <div style={{ marginTop: 10 }}>
                    <button
                      type="button"
                      style={styles.btnPrimary}
                      onClick={downloadPreviewExcel}
                    >
                      Download Preview Excel
                    </button>
                    <div style={{ fontSize: 12, color: "#667085", marginTop: 6 }}>
                      Download this file, correct location / kms if needed, then upload the corrected Excel again.
                    </div>
                  </div>
                )}
              </div>

              <div>
                <div style={styles.formLabel}>Select Images (bulk)</div>
                <input
                  ref={imagesInputRef}
                  type="file"
                  multiple
                  accept="image/*"
                  onChange={handleImagesChange}
                />
                <div style={{ fontSize: 12, color: "#667085", marginTop: 6 }}>
                  Selected: <b>{imageFiles.length}</b> images
                </div>
              </div>
            </div>

            {summary && (
              <div style={{ ...styles.stateBox, marginTop: 12 }}>
                <div style={{ fontWeight: 800, marginBottom: 8 }}>Import Summary</div>

                <div style={{ fontSize: 13, lineHeight: 1.7 }}>
                  Points read: <b>{summary.pointsRead}</b>
                  <br />
                  Reports used: <b>{summary.reportsCreatedOrUsed}</b>
                  <br />
                  Images selected: <b>{summary.imagesSelected}</b>
                  <br />
                  Images uploaded: <b>{summary.imagesUploaded}</b>
                  <br />
                  report_photos inserted: <b>{summary.photosInserted}</b>
                </div>

                {summary.duplicateSelectedImages.length > 0 && (
                  <div style={{ marginTop: 10, color: "#B42318", fontSize: 13 }}>
                    <b>Duplicate selected image names:</b>
                    <div style={styles.scrollBox}>
                      {summary.duplicateSelectedImages.map((f) => (
                        <div key={f}>{f}</div>
                      ))}
                    </div>
                  </div>
                )}

                {summary.duplicateMappingFiles.length > 0 && (
                  <div style={{ marginTop: 10, color: "#B42318", fontSize: 13 }}>
                    <b>Duplicate file_name values in master file:</b>
                    <div style={styles.scrollBox}>
                      {summary.duplicateMappingFiles.map((f) => (
                        <div key={f}>{f}</div>
                      ))}
                    </div>
                  </div>
                )}

                {summary.missingPointKeysInPointsCsv.length > 0 && (
                  <div style={{ marginTop: 10, color: "#B42318", fontSize: 13 }}>
                    <b>point_key used in master file images but missing in points:</b>
                    <div style={styles.scrollBox}>
                      {summary.missingPointKeysInPointsCsv.map((f) => (
                        <div key={f}>{f}</div>
                      ))}
                    </div>
                  </div>
                )}

                {summary.noGpsImages.length > 0 && (
                  <div style={{ marginTop: 10, color: "#7A5AF8", fontSize: 13 }}>
                    <b>NO_GPS images:</b>
                    <div style={styles.scrollBox}>
                      {summary.noGpsImages.map((f) => (
                        <div key={f}>{f}</div>
                      ))}
                    </div>
                  </div>
                )}

                {summary.missingFilesInUpload.length > 0 && (
                  <div style={{ marginTop: 10, color: "#B42318", fontSize: 13 }}>
                    <b>In master file but not selected:</b>
                    <div style={styles.scrollBox}>
                      {summary.missingFilesInUpload.map((f) => (
                        <div key={f}>{f}</div>
                      ))}
                    </div>
                  </div>
                )}

                {summary.extraFilesNotInMap.length > 0 && (
                  <div style={{ marginTop: 10, color: "#475467", fontSize: 13 }}>
                    <b>Selected images not present in master file (skipped):</b>
                    <div style={styles.scrollBox}>
                      {summary.extraFilesNotInMap.map((f) => (
                        <div key={f}>{f}</div>
                      ))}
                    </div>
                  </div>
                )}

                {summary.errors.length > 0 && (
                  <div style={{ marginTop: 10, color: "#B42318", fontSize: 13 }}>
                    <b>Errors / Warnings:</b>
                    <div style={styles.scrollBox}>
                      {summary.errors.map((e, idx) => (
                        <div key={idx}>{e}</div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            )}

            <div style={styles.modalActions}>
              <button
                style={styles.btnGhost}
                onClick={() => setBulkOpen(false)}
                disabled={importing}
              >
                Close
              </button>
              <button
                style={{ ...styles.btnPrimary, opacity: importing ? 0.7 : 1 }}
                onClick={runImportSingleMaster}
                disabled={importing}
              >
                {importing ? "Importing..." : "Run Import"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  page: {
    padding: 24,
    fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif",
    background: "#F7F8FA",
    minHeight: "100vh",
    maxWidth: 1400,
    margin: "0 auto",
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    gap: 12,
    flexWrap: "wrap",
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 16,
    padding: 16,
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
  },
  title: { fontSize: 22, fontWeight: 800, color: "#101828", lineHeight: 1.2 },
  subtitle: { fontSize: 13, color: "#667085", marginTop: 6 },
  headerRight: {
    display: "flex",
    gap: 10,
    alignItems: "center",
    flexWrap: "wrap",
    justifyContent: "flex-end",
  },
  exportGroup: { display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" },
  btnPrimary: {
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid #111",
    background: "#111",
    color: "#fff",
    cursor: "pointer",
    fontWeight: 700,
    fontSize: 13,
  },
  btnGhost: {
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid #EAECF0",
    background: "#fff",
    cursor: "pointer",
    fontWeight: 700,
    fontSize: 13,
    color: "#344054",
  },
  btnDanger: {
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid #FDA29B",
    background: "#FEF3F2",
    cursor: "pointer",
    fontWeight: 700,
    fontSize: 13,
    color: "#B42318",
  },
  btnGhostSmall: {
    flex: 1,
    padding: "9px 10px",
    borderRadius: 10,
    border: "1px solid #EAECF0",
    background: "#fff",
    cursor: "pointer",
    fontWeight: 700,
    fontSize: 12,
    color: "#344054",
  },
  btnDangerSmall: {
    flex: 1,
    padding: "9px 10px",
    borderRadius: 10,
    border: "1px solid #FDA29B",
    background: "#FEF3F2",
    cursor: "pointer",
    fontWeight: 700,
    fontSize: 12,
    color: "#B42318",
  },
  searchBar: { marginTop: 14, marginBottom: 14 },
  searchWrap: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 14,
    padding: "10px 12px",
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
  },
  searchIcon: { color: "#667085", fontSize: 14 },
  searchInput: { flex: 1, border: "none", outline: "none", fontSize: 14, color: "#101828" },
  clearBtn: {
    padding: "8px 10px",
    borderRadius: 10,
    border: "1px solid #EAECF0",
    background: "#fff",
    cursor: "pointer",
    fontWeight: 700,
    fontSize: 12,
    color: "#344054",
  },
  stateBox: {
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 16,
    padding: 18,
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
  },
  grid: { display: "grid", gridTemplateColumns: "repeat(4, minmax(0, 1fr))", gap: 12 },
  cardWrap: {
    display: "flex",
    flexDirection: "column",
    gap: 8,
  },
  card: {
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 16,
    padding: 14,
    color: "#101828",
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
    cursor: "pointer",
  },
  cardActions: {
    display: "flex",
    gap: 8,
  },
  cardTop: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    gap: 10,
    marginBottom: 10,
  },
  cardTitle: { fontSize: 16, fontWeight: 800, lineHeight: 1.2 },
  badge: {
    fontSize: 12,
    fontWeight: 800,
    padding: "4px 10px",
    borderRadius: 999,
    border: "1px solid #D0D5DD",
    background: "#F9FAFB",
    color: "#344054",
    whiteSpace: "nowrap",
  },
  metaRow: {
    display: "flex",
    justifyContent: "space-between",
    gap: 10,
    padding: "8px 0",
    borderTop: "1px dashed #EAECF0",
  },
  metaLabel: { fontSize: 12, color: "#667085", fontWeight: 700 },
  metaValue: {
    fontSize: 12,
    color: "#101828",
    fontWeight: 700,
    maxWidth: 190,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  cardHint: { marginTop: 10, fontSize: 12, color: "#475467", fontWeight: 700 },
  modalOverlay: {
    position: "fixed",
    inset: 0,
    background: "rgba(0,0,0,0.35)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: 16,
    zIndex: 9999,
  },
  modal: {
    width: "min(640px, 100%)",
    background: "#fff",
    borderRadius: 16,
    border: "1px solid #EAECF0",
    padding: 16,
    boxShadow: "0 10px 30px rgba(0,0,0,0.15)",
  },
  modalWide: {
    width: "min(980px, 100%)",
    background: "#fff",
    borderRadius: 16,
    border: "1px solid #EAECF0",
    padding: 16,
    boxShadow: "0 10px 30px rgba(0,0,0,0.15)",
    maxHeight: "90vh",
    overflow: "auto",
  },
  formRow: { marginTop: 10 },
  formLabel: { fontSize: 12, fontWeight: 800, color: "#344054", marginBottom: 6 },
  input: {
    width: "100%",
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid #D0D5DD",
    outline: "none",
    fontSize: 14,
  },
  textarea: {
    width: "100%",
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid #D0D5DD",
    outline: "none",
    fontSize: 14,
    minHeight: 90,
    resize: "vertical",
  },
  modalActions: { display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 14 },
  scrollBox: {
    maxHeight: 160,
    overflow: "auto",
    marginTop: 6,
    border: "1px solid #EAECF0",
    borderRadius: 10,
    padding: 8,
    background: "#fff",
  },
  fileMeta: { marginTop: 6, fontSize: 12, color: "#667085" },
  previewBox: {
    marginTop: 8,
    background: "#F8FAFC",
    border: "1px solid #EAECF0",
    borderRadius: 10,
    padding: 10,
    fontSize: 12,
    lineHeight: 1.5,
    color: "#101828",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    maxHeight: 140,
    overflow: "auto",
  },
  categoryHelpBox: {
    marginTop: 4,
    background: "#F8FAFC",
    border: "1px solid #EAECF0",
    borderRadius: 10,
    padding: 10,
    fontSize: 12,
    lineHeight: 1.6,
    color: "#101828",
  },
};
