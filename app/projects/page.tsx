"use client";

import React, { useEffect, useMemo, useRef, useState } from "react";
import { useRouter } from "next/navigation";
import * as XLSX from "xlsx";
import { supabase } from "../../lib/supabaseClient";
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
  const iFile = colIndex(headers, ["file_name", "filename", "file", "name", "image_name"]);
  const iImgKey = colIndex(headers, ["image_key", "imagekey", "img_key", "imgkey"]);
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

async function uploadToBucket(bucket: string, path: string, file: File) {
  const { data, error } = await supabase.storage.from(bucket).upload(path, file, {
    cacheControl: "3600",
    upsert: true,
    contentType: file.type || undefined,
  });
  if (error) throw error;

  const { data: pub } = supabase.storage.from(bucket).getPublicUrl(data.path);
  return { path: data.path, publicUrl: pub.publicUrl };
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

async function sha256File(file: File): Promise<string> {
  const buffer = await file.arrayBuffer();
  const hashBuffer = await crypto.subtle.digest("SHA-256", buffer);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map((b) => b.toString(16).padStart(2, "0")).join("");
}

export default function ProjectsPage() {
  const router = useRouter();

  const [projects, setProjects] = useState<ProjectRow[]>([]);
  const [loading, setLoading] = useState(true);
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
  const [bucketName, setBucketName] = useState<string>("reports");
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

  const hydrateLastModifiedNames = async (rows: ProjectRow[]) => {
    try {
      const userIds = Array.from(
        new Set(rows.map((r) => r.last_modified_by).filter(Boolean) as string[])
      );

      if (!userIds.length) {
        setLastModifiedMap({});
        return;
      }

      const { data: profiles, error } = await supabase
        .from("profiles")
        .select("user_id, name, email")
        .in("user_id", userIds);

      if (error) throw error;

      const userIdToName: Record<string, string> = {};
      (profiles || []).forEach((u: any) => {
        const uid = String(u?.user_id || "");
        if (!uid) return;
        userIdToName[uid] = u?.name || u?.email || uid.slice(0, 8);
      });

      const map: Record<string, string> = {};
      rows.forEach((p) => {
        if (p.last_modified_by) map[p.id] = userIdToName[p.last_modified_by] || p.last_modified_by.slice(0, 8);
      });
      setLastModifiedMap(map);
    } catch {
      const fallback: Record<string, string> = {};
      rows.forEach((p) => {
        if (p.last_modified_by) fallback[p.id] = p.last_modified_by.slice(0, 8);
      });
      setLastModifiedMap(fallback);
    }
  };

  const load = async () => {
    setLoading(true);
    try {
      const { data, error } = await supabase
        .from("projects")
        .select("*")
        .order("created_at", { ascending: false });

      if (error) throw error;

      const rows = (data || []) as ProjectRow[];
      setProjects(rows);
      await hydrateLastModifiedNames(rows);

      if (!bulkProjectId && rows.length) setBulkProjectId(rows[0].id);
    } catch (e: any) {
      const msg = String(e?.message || e || "").toLowerCase();
      if (msg.includes("auth") || msg.includes("session") || msg.includes("jwt")) {
        redirectToLogin();
        return;
      }
      alert(e?.message || String(e));
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    const init = async () => {
      try {
        const {
          data: { session },
          error,
        } = await supabase.auth.getSession();

        if (error) throw error;

        if (!session?.user) {
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

  useEffect(() => {
    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange((event) => {
      if (event === "SIGNED_OUT") {
        redirectToLogin();
      }
    });

    return () => {
      subscription.unsubscribe();
    };
  }, []);

  const logout = async () => {
    await supabase.auth.signOut();
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

      const {
        data: { session },
        error: sessionErr,
      } = await supabase.auth.getSession();

      if (sessionErr) throw sessionErr;

      if (!session?.user) {
        redirectToLogin();
        return;
      }

      const user = session.user;

      const { error } = await supabase.from("projects").insert([
        {
          user_id: user.id,
          name,
          description: newDesc.trim() || null,
          created_by: user.id,
          last_modified_by: user.id,
        },
      ]);

      if (error) throw error;

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

      const {
        data: { session },
        error: sessionErr,
      } = await supabase.auth.getSession();

      if (sessionErr) throw sessionErr;
      if (!session?.user) {
        redirectToLogin();
        return;
      }

      const { error } = await supabase
        .from("projects")
        .update({
          name,
          description: editDesc.trim() || null,
          last_modified_by: session.user.id,
          updated_at: new Date().toISOString(),
        })
        .eq("id", editingProjectId);

      if (error) throw error;

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

      const { data: reportRows, error: reportFetchError } = await supabase
        .from("reports")
        .select("id")
        .eq("project_id", project.id);

      if (reportFetchError) throw reportFetchError;

      const reportIds = (reportRows || []).map((row: any) => row.id).filter(Boolean);

      if (reportIds.length) {
        const { error: pathDeleteError } = await supabase
          .from("report_path_points")
          .delete()
          .in("report_id", reportIds);

        if (pathDeleteError) throw pathDeleteError;

        const { error: photosDeleteError } = await supabase
          .from("report_photos")
          .delete()
          .in("report_id", reportIds);

        if (photosDeleteError) throw photosDeleteError;
      }

      const { error: reportsDeleteError } = await supabase
        .from("reports")
        .delete()
        .eq("project_id", project.id);

      if (reportsDeleteError) throw reportsDeleteError;

      const { error: historyDeleteError } = await supabase
        .from("bulk_import_history")
        .delete()
        .eq("project_id", project.id);

      if (historyDeleteError) throw historyDeleteError;

      const { error: projectDeleteError } = await supabase
        .from("projects")
        .delete()
        .eq("id", project.id);

      if (projectDeleteError) throw projectDeleteError;

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
    if (!bucketName.trim()) return alert("Bucket name is required.");
    if (!masterFile) return alert("Select master file.");
    if (!imageFiles.length) return alert("Select image files (bulk).");

    setImporting(true);
    setSummary(null);

    const errors: string[] = [];

    try {
      const {
        data: { session },
        error: sessionErr,
      } = await supabase.auth.getSession();

      if (sessionErr) throw sessionErr;
      if (!session?.user) {
        redirectToLogin();
        return;
      }

      const masterFileHash = await sha256File(masterFile);

      const { data: existingImportRows, error: existingImportErr } = await supabase
        .from("bulk_import_history")
        .select("id")
        .eq("project_id", bulkProjectId)
        .eq("master_file_hash", masterFileHash)
        .limit(1);

      if (existingImportErr) throw existingImportErr;

      if (existingImportRows && existingImportRows.length > 0) {
        alert("This master file has already been imported for this project.");
        return;
      }

      const combinedRows = await parseCombinedFile(masterFile);

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

      const { data: existingPointRows, error: existingPointErr } = await supabase
        .from("reports")
        .select("point_key")
        .eq("project_id", bulkProjectId)
        .in("point_key", incomingPointKeys);

      if (existingPointErr) throw existingPointErr;

      const existingPointKeys = Array.from(
        new Set(
          (existingPointRows || [])
            .map((row: any) => String(row.point_key || "").trim())
            .filter(Boolean)
        )
      );

      if (existingPointKeys.length > 0) {
        alert(
          `Import blocked. These point_key values already exist in this project: ${existingPointKeys.join(", ")}`
        );
        return;
      }

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
          imageMap.push({
            file_name: row.file_name,
            point_key: row.point_key,
            image_key: row.image_key || null,
          });
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

      let noGpsReportId: string | null = null;

      async function getOrCreateNoGpsReport() {
        if (noGpsReportId) return noGpsReportId;

        const noGpsCategory = "NO_GPS Images";
        const nowIso = new Date().toISOString();

        const { data: existing, error: exErr } = await supabase
          .from("reports")
          .select("id")
          .eq("project_id", bulkProjectId)
          .eq("category", noGpsCategory)
          .order("created_at", { ascending: false })
          .limit(1);

        if (exErr) throw exErr;

        if (existing && existing.length) {
          noGpsReportId = existing[0].id;
          return noGpsReportId;
        }

        const { data: created, error: crErr } = await supabase
          .from("reports")
          .insert([
            {
              project_id: bulkProjectId,
              point_key: null,
              category: noGpsCategory,
              description: "Images that do not have GPS point mapping (bulk import).",
              route_id: null,
              difficulty: "NO_GPS",
              loc_lat: null,
              loc_lon: null,
              loc_acc: null,
              loc_time: nowIso,
            },
          ])
          .select("id")
          .single();

        if (crErr) throw crErr;

        noGpsReportId = created?.id ?? null;
        if (!noGpsReportId) throw new Error("Failed to create NO_GPS report.");

        return noGpsReportId;
      }

      const reportIdByPointKey = new Map<string, string>();
      const sortedKeys = Array.from(pointByKey.keys()).sort((a, b) => {
        const na = Number(a);
        const nb = Number(b);
        if (Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
        return a.localeCompare(b);
      });

      for (const key of sortedKeys) {
        const p = pointByKey.get(key)!;
        const category = p.category.trim() || "Unknown";
        const description = p.description?.trim() || null;
        const difficulty = normalizeDifficulty(p.difficulty || "green");
        const nowIso = new Date().toISOString();

        const { data: created, error: cErr } = await supabase
          .from("reports")
          .insert([
            {
              project_id: bulkProjectId,
              point_key: key,
              category,
              description,
              route_id: null,
              difficulty,
              remarks_action: p.remarks_action?.trim() || null,
              loc_lat: p.latitude,
              loc_lon: p.longitude,
              loc_acc: null,
              loc_time: nowIso,
            },
          ])
          .select("id")
          .single();

        if (cErr) throw cErr;

        const reportId = created?.id ?? null;
        if (!reportId) throw new Error(`Failed to create report for point_key=${key}`);

        reportIdByPointKey.set(key, reportId);

        const { error: insErr } = await supabase
          .from("report_path_points")
          .insert([
            {
              report_id: reportId,
              seq: 1,
              latitude: p.latitude,
              longitude: p.longitude,
              elevation: null,
              accuracy: null,
              timestamp: nowIso,
            },
          ]);

        if (insErr) throw insErr;
      }

      const mapByFileName = new Map<string, ParsedImageMapRow>();
      imageMap.forEach((r) => {
        const key = normalizeFileKey(r.file_name);
        if (key && !mapByFileName.has(key)) mapByFileName.set(key, r);
      });

      const selectedByName = new Map<string, File>();
      imageFiles.forEach((f) => {
        const key = normalizeFileKey(f.name);
        if (key && !selectedByName.has(key)) selectedByName.set(key, f);
      });

      const missingFilesInUpload: string[] = [];
      for (const [name] of mapByFileName) {
        if (!selectedByName.has(name)) missingFilesInUpload.push(name);
      }

      const extraFilesNotInMap: string[] = [];
      imageFiles.forEach((f) => {
        if (!mapByFileName.has(normalizeFileKey(f.name))) extraFilesNotInMap.push(f.name);
      });

      const noGpsImages: string[] = [];
      let imagesUploaded = 0;
      let photosInserted = 0;

      await mapLimit(imageFiles, 3, async (file) => {
        const mapping = mapByFileName.get(normalizeFileKey(file.name));
        if (!mapping) return;

        const point_key = mapping.point_key ?? null;

        let reportId = point_key ? reportIdByPointKey.get(point_key) : null;
        if (!reportId) {
          reportId = await getOrCreateNoGpsReport();
          noGpsImages.push(file.name);
        }

        const safeFileName = file.name.replace(/[^\w.\-]+/g, "_");
        const storagePath = `${bulkProjectId}/${reportId}/${safeFileName}`;

        let uploaded: { path: string; publicUrl: string };
        try {
          uploaded = await uploadToBucket(bucketName.trim(), storagePath, file);
        } catch (upErr: any) {
          errors.push(`${file.name}: upload failed - ${upErr?.message || String(upErr)}`);
          return;
        }

        imagesUploaded++;

        const url = uploaded.publicUrl;
        if (!url) {
          errors.push(`${file.name}: could not generate public URL`);
          return;
        }

        const { width, height } = await getImageSize(file);

        const { data: existingPhotoRows, error: existingPhotoErr } = await supabase
          .from("report_photos")
          .select("id")
          .eq("report_id", reportId)
          .eq("url", url)
          .limit(1);

        if (existingPhotoErr) {
          errors.push(`${file.name}: photo check failed - ${existingPhotoErr.message}`);
          return;
        }

        const alreadyExists =
          Array.isArray(existingPhotoRows) && existingPhotoRows.length > 0;

        if (!alreadyExists) {
          const { error: insErr } = await supabase.from("report_photos").insert([
            {
              report_id: reportId,
              url,
              width,
              height,
            },
          ]);

          if (insErr) {
            errors.push(`${file.name}: report_photos insert failed - ${insErr.message}`);
            return;
          }

          photosInserted++;
        }
      });

      const { error: historyErr } = await supabase
        .from("bulk_import_history")
        .insert([
          {
            project_id: bulkProjectId,
            master_file_name: masterFile.name,
            master_file_hash: masterFileHash,
          },
        ]);

      if (historyErr) throw historyErr;

      if (imageFiles.length > 0 && photosInserted === 0) {
        errors.push(
          "No report_photos rows were inserted. Most likely the uploaded file names do not exactly match the master file file_name values."
        );
      }

      setSummary({
        pointsRead: points.length,
        reportsCreatedOrUsed: sortedKeys.length + (noGpsReportId ? 1 : 0),
        imagesSelected: imageFiles.length,
        imagesUploaded,
        photosInserted,
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
                <div style={styles.formLabel}>Storage Bucket</div>
                <input
                  style={styles.input}
                  value={bucketName}
                  onChange={(e) => setBucketName(e.target.value)}
                />
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
