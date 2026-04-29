"use client";

import React, { useEffect, useMemo, useRef, useState } from "react";
import Link from "next/link";
import { useParams } from "next/navigation";
import { Editor } from "@tinymce/tinymce-react";

import { dbClient } from "../../../lib/dbClient";
import {
  generateProjectGPX,
  generateProjectGPXByReportIds,
} from "../../../lib/download";

type VehicleMovement = "green" | "yellow" | "red" | "";
type VMFilter = "all" | "green" | "yellow" | "red" | "unset";

type ExportFormat = "docx" | "gpx";
type ExportMode = "listed" | "selectedOne" | "selectedSplit" | "all";

type ProjectRow = {
  id: string;
  name?: string | null;
  title?: string | null;
  project_name?: string | null;
  created_at?: string | null;
};

type ReportRow = {
  id: string;
  project_id: string;
  route_id?: string | null;
  category?: string | null;
  description?: string | null;
  remarks_action?: string | null;
  created_at: string;
  difficulty?: VehicleMovement | null; // ✅ DB column
  sort_order?: number | null; // ✅ NEW (for inserting at exact position)
};

type WatermarkOpts = { enabled: boolean; text: string };

type ManualPointRow = {
  report_id: string;
  user_id: string;
  seq: number;
  latitude: number;
  longitude: number;
  elevation?: number | null;
  accuracy?: number | null;
  timestamp?: string | null;
};
type PreparedFile = { fileName: string; blob: Blob };

// ========= ROUTE SURVEY TYPES =========
type PresetMap = { id: string; label: string; src: string };
type PickMode = "preset" | "upload";

// ========= LOCATIONS TABLE ROW (OPTIONAL) =========
type RouteLocationRow = {
  id: string;
  label?: string | null;
  pin_type?: string | null;
  sort_order?: number | null;
};

function projectNameOf(p: ProjectRow | null) {
  return p?.name || p?.title || p?.project_name || "Project";
}

function sanitizeFileBaseName(name: string) {
  const cleaned = String(name || "")
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/[. ]+$/g, "");
  return cleaned.slice(0, 80) || "Export";
}

function displayDescription(raw: string) {
  const t = (raw || "").trim();
  if (!t) return "";
  return t.replace(/^considered\s*/i, "").trim();
}

function normalizeVM(v: any): VehicleMovement {
  const t = String(v ?? "").trim().toLowerCase();
  if (t === "green") return "green";
  if (t === "yellow" || t === "amber") return "yellow";
  if (t === "red") return "red";
  return "";
}

function vmDisplayToDb(v: string): VehicleMovement {
  const t = String(v || "").trim().toLowerCase();
  if (t === "green") return "green";
  if (t === "yellow") return "yellow";
  if (t === "red") return "red";
  return "";
}

async function getAuthUserFromApi() {
  const token =
    typeof window !== "undefined" ? localStorage.getItem("auth_token") : null;
  const res = await fetch("/api/auth/me", {
    method: "GET",
    credentials: "include",
    headers: token ? { Authorization: `Bearer ${token}` } : {},
  });
  if (!res.ok) throw new Error("Not logged in.");
  const data = await res.json().catch(() => ({} as any));
  const user = data?.user;
  if (!user?.id) throw new Error("Not logged in.");
  return user as { id: string };
}

function authHeaders() {
  const token =
    typeof window !== "undefined" ? localStorage.getItem("auth_token") : null;
  return token ? { Authorization: `Bearer ${token}` } : {};
}

async function parseJsonSafe(res: Response) {
  return res.json().catch(() => ({} as any));
}

async function apiRequestJson(url: string, init: RequestInit = {}) {
  const mergedHeaders: Record<string, string> = {
    ...(init.body ? { "Content-Type": "application/json" } : {}),
    ...authHeaders(),
    ...(init.headers as Record<string, string> | undefined),
  };
  const res = await fetch(url, {
    ...init,
    credentials: "include",
    headers: mergedHeaders,
  });
  const data = await parseJsonSafe(res);
  if (!res.ok) throw new Error(data?.error || "Request failed");
  return data;
}

async function updateReportVM(reportId: string, next: VehicleMovement) {
  const payload: any = { difficulty: next ? next : null };
  await apiRequestJson(`/api/reports/${encodeURIComponent(reportId)}`, {
    method: "PUT",
    body: JSON.stringify(payload),
  });
}


async function uploadToBucket(bucket: string, path: string, file: File) {
  const normalizedPath = String(path || "").replace(/^\/+/, "").replace(/\\/g, "/");
  const slash = normalizedPath.lastIndexOf("/");
  const folderFromPath = slash > 0 ? normalizedPath.slice(0, slash) : normalizedPath || "uploads";
  const formData = new FormData();
  formData.append("folder", folderFromPath);
  formData.append("file", file);

  const res = await fetch("/api/upload", {
    method: "POST",
    credentials: "include",
    headers: authHeaders(),
    body: formData,
  });
  const data = await parseJsonSafe(res);
  if (!res.ok) throw new Error(data?.error || "Upload failed");
  return { path: String(data?.key || data?.path || path), publicUrl: String(data?.url || "") };
}

async function uploadReportPhotos(projectId: string, reportId: string, files: File[]) {
  if (!files.length) return;

  // helper to read width/height
  async function getImageSize(file: File): Promise<{ width: number; height: number } | null> {
    try {
      const url = URL.createObjectURL(file);
      const img = new Image();
      const p = new Promise<{ width: number; height: number }>((resolve, reject) => {
        img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
        img.onerror = reject;
      });
      img.src = url;
      const size = await p;
      URL.revokeObjectURL(url);
      return size;
    } catch {
      return null;
    }
  }

  // Upload in parallel (faster) then insert DB rows once
  const uploads = await Promise.all(
    files.map(async (f, index) => {
      const safeName = f.name.replace(/[^\w.\-]+/g, "_");
      const uniquePrefix = `${Date.now()}_${index}_${Math.random().toString(36).slice(2, 8)}`;

      // Storage layout: reports/photos/<reportId>/<fileName>. The folder MUST
      // be keyed off the actual reports.id (NOT user.id, NOT projectId) so the
      // export query `WHERE report_id IN (...)` joins against report_photos.
      const storagePath = `reports/photos/${reportId}/${uniquePrefix}_${safeName}`;

      const size = await getImageSize(f);
      const uploaded = await uploadToBucket(REPORT_PHOTO_BUCKET, storagePath, f);

      return {
        report_id: reportId,
        url: uploaded.publicUrl,
        width: size?.width ?? null,
        height: size?.height ?? null,
      };
    })
  );

  const uniqueUploads = uploads.filter(
    (row, index, arr) => index === arr.findIndex((x) => x.url === row.url)
  );

  const res = await fetch(`/api/reports/${encodeURIComponent(reportId)}/photos`, {
    method: "POST",
    credentials: "include",
    headers: {
      "Content-Type": "application/json",
      ...authHeaders(),
    },
    body: JSON.stringify({
      photos: uniqueUploads.map((x) => ({
        url: x.url,
        width: x.width,
        height: x.height,
        file_name: x.url.split("/").pop() || null,
      })),
    }),
  });
  const data = await parseJsonSafe(res);
  if (!res.ok) throw new Error(data?.error || "Failed to save report photos");
}

function vmFilterLabel(f: VMFilter) {
  if (f === "all") return "ALL";
  if (f === "unset") return "UNSET";
  return f.toUpperCase();
}

function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

async function fetchProjectExportDocx(params: {
  projectId: string;
  reportIds?: string[];
  includePhotos?: boolean;
  fileName?: string;
}) {
  const token =
    typeof window !== "undefined" ? localStorage.getItem("auth_token") : null;
  const authHdrs: Record<string, string> = token ? { Authorization: `Bearer ${token}` } : {};
  const exportUrl = `/api/projects/${encodeURIComponent(params.projectId)}/export`;

  // Use POST with a JSON body so the upstream Nginx never sees a multi-KB
  // ?reportIds=... query string (which triggers 414/431 above ~30 UUIDs).
  // Fall back to GET only when there is no selection at all.
  const useGet = !params.reportIds?.length;
  let res: Response;
  if (useGet) {
    const sp = new URLSearchParams();
    if (typeof params.includePhotos === "boolean") {
      sp.set("includePhotos", params.includePhotos ? "1" : "0");
    }
    if (params.fileName?.trim()) {
      sp.set("fileName", params.fileName.trim());
    }
    const suffix = sp.toString() ? `?${sp.toString()}` : "";
    res = await fetch(`${exportUrl}${suffix}`, {
      method: "GET",
      credentials: "include",
      headers: authHdrs,
    });
  } else {
    res = await fetch(exportUrl, {
      method: "POST",
      credentials: "include",
      headers: { "Content-Type": "application/json", ...authHdrs },
      body: JSON.stringify({
        reportIds: params.reportIds,
        includePhotos: params.includePhotos,
        fileName: params.fileName?.trim() || undefined,
      }),
    });
  }

  if (!res.ok) {
    const data = await res.json().catch(() => ({} as any));
    throw new Error(data?.error || "Failed to export");
  }

  return res.blob();
}

function clamp(n: number, a: number, b: number) {
  return Math.max(a, Math.min(b, n));
}

function readAvgSecondsPerReport(includePhotos: boolean) {
  const key = includePhotos ? "docx_avg_sec_photo" : "docx_avg_sec_nophoto";
  const v = Number(localStorage.getItem(key) || "");
  return Number.isFinite(v) && v > 0 ? v : includePhotos ? 1.4 : 0.6;
}

function estimateSeconds(mode: ExportMode, count: number, includePhotos: boolean) {
  const per = readAvgSecondsPerReport(includePhotos);
  const base =
    mode === "all"
      ? 6
      : mode === "listed"
        ? 4
        : mode === "selectedOne"
          ? 3
          : mode === "selectedSplit"
            ? 4
            : 4;

  const est = Math.round(base + count * per);
  return clamp(est, 6, 10 * 60);
}

function parseStageRanges(input: string, total: number) {
  const out: Array<{ from: number; to: number; label: string }> = [];
  const parts = String(input || "")
    .split(",")
    .map((x) => x.trim())
    .filter(Boolean);

  const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

  for (let i = 0; i < parts.length; i++) {
    const p = parts[i];
    const m = p.match(/^(\d+)\s*-\s*(\d+)$/) || p.match(/^(\d+)$/);
    if (!m) continue;

    let from = Number(m[1]);
    let to = Number(m[2] ?? m[1]);

    if (!Number.isFinite(from) || !Number.isFinite(to)) continue;
    if (from < 1) from = 1;
    if (to < 1) to = 1;
    if (from > total) from = total;
    if (to > total) to = total;
    if (to < from) [from, to] = [to, from];

    const label = letters[i] || `S${i + 1}`;
    out.push({ from, to, label });
  }

  return out;
}

// ✅ fixed preview heights (ONLY sizing change; UI remains same)
const MAP_PREVIEW_HEIGHT = 320;
const GA_PREVIEW_HEIGHT = 220;

// ✅ spacing used for sort_order
const SORT_STEP = 10;
// ✅ if gap becomes too small, we renumber everything to 10,20,30...
const MIN_GAP = 1;


// ✅ Category options for new report creation (dropdown + custom)
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


// ✅ report photos upload
const REPORT_PHOTO_BUCKET = "reports";
const REPORT_IMAGE_TABLE = "report_photos";

export default function ProjectReportsPage() {
  const params = useParams();
  const projectId = useMemo(() => {
    const id = (params as any)?.id;
    return Array.isArray(id) ? id[0] : id;
  }, [(params as any)?.id]);

  const [project, setProject] = useState<ProjectRow | null>(null);
  const [reports, setReports] = useState<ReportRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);
  const fetchingRef = useRef(false);


  const [vmSaving, setVmSaving] = useState<Record<string, boolean>>({});

  // Search + filter
  const [q, setQ] = useState("");
  const [vmFilter, setVmFilter] = useState<VMFilter>("all");

  // ✅ sorting UI label only (we still always keep "in-list order" by sort_order)
  const [sortDir, setSortDir] = useState<"asc" | "desc">("asc");

  // selection
  const [selected, setSelected] = useState<Record<string, boolean>>({});
  const [selMenuOpen, setSelMenuOpen] = useState(false);
  const selMenuRef = useRef<HTMLDivElement>(null);

  // watermark
  const [wmEnabled, setWmEnabled] = useState(true);
  const [wmText, setWmText] = useState("");
  const [wmDirty, setWmDirty] = useState(false);

  // Export modal
  const [exportModalOpen, setExportModalOpen] = useState(false);
  const [exportFormat, setExportFormat] = useState<ExportFormat>("docx");
  const [exportMode, setExportMode] = useState<ExportMode>("listed");
  const [exportName, setExportName] = useState("");
  const [stageRanges, setStageRanges] = useState("1-12,13-14,15-25");
  const [includePhotos, setIncludePhotos] = useState(true);

  // Download progress modal
  const [dlOpen, setDlOpen] = useState(false);
  const [dlTitle, setDlTitle] = useState("");
  const [dlError, setDlError] = useState<string | null>(null);
  const [dlDone, setDlDone] = useState(false);
  const [dlSecondsLeft, setDlSecondsLeft] = useState(0);
  const [preparedFiles, setPreparedFiles] = useState<PreparedFile[]>([]);

  // ✅ GA setup modal gate before export
  const [gaSetupOpen, setGaSetupOpen] = useState(false);
  const [gaSetupReason, setGaSetupReason] = useState<string>("");

  // ✅ if GA already exists -> ask Edit or Skip
  const [gaChoiceOpen, setGaChoiceOpen] = useState(false);
  const [gaExistingPageId, setGaExistingPageId] = useState<string | null>(null);

  // ✅ Insert report modal (between rows)
  const [insertOpen, setInsertOpen] = useState(false);
  const [insertAfterId, setInsertAfterId] = useState<string | null>(null);

  // ✅ Manual points modal (report_path_points)
  const [manualPointsOpen, setManualPointsOpen] = useState(false);
  const [manualPointsReportId, setManualPointsReportId] = useState<string | null>(null);

  // ✅ Upload photos (for selected report)
  const [photoOpen, setPhotoOpen] = useState(false);
  const [photoReportId, setPhotoReportId] = useState<string | null>(null);

  const projectName = projectNameOf(project);

  useEffect(() => {
    const def = `CONFIDENTIAL REPORT: ${projectName}`;
    setWmText((prev) => (wmDirty ? prev : prev?.trim() ? prev : def));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [projectName]);

  const watermarkOpts: WatermarkOpts = useMemo(() => {
    const def = `CONFIDENTIAL REPORT: ${projectName}`;
    return { enabled: wmEnabled, text: (wmText || def).trim() };
  }, [wmEnabled, wmText, projectName]);

  // ✅ NOTE:
  // We ALWAYS order the list by sort_order so insertion works exactly at a position.
  // sortDir controls ascending/descending display only.
  const fetchReports = async (searchText: string, dir: "asc" | "desc", filter: VMFilter) => {
    if (!projectId) return;
    if (fetchingRef.current) return;
    fetchingRef.current = true;

    try {
      const params = new URLSearchParams();
      const sText = searchText.trim();
      if (sText) params.set("search", sText);
      if (filter && filter !== "all") params.set("difficulty", filter);
      params.set("sort", dir);

      const res = await fetch(`/api/projects/${encodeURIComponent(projectId)}/reports?${params.toString()}`, {
        method: "GET",
        credentials: "include",
        headers: authHeaders(),
      });
      const data = await parseJsonSafe(res);
      if (!res.ok) throw new Error(data?.error || "Failed to fetch reports");
      setReports(((data?.reports || data || []) as ReportRow[]) || []);
      setLoadError(null);
    } finally {
      fetchingRef.current = false;
    }
  };

  const load = async () => {
    if (!projectId) return;
    setLoading(true);
    setLoadError(null);
    try {
      const pRes = await fetch(`/api/projects/${encodeURIComponent(projectId)}`, {
        method: "GET",
        credentials: "include",
        headers: authHeaders(),
      });
      const pData = await parseJsonSafe(pRes);
      if (!pRes.ok) throw new Error(pData?.error || "Failed to fetch project");
      setProject((pData?.project || pData?.data || pData) as ProjectRow);

      await fetchReports(q, sortDir, vmFilter);
      setSelected({});
    } catch (e: any) {
      setLoadError(e?.message || "Failed to fetch reports");
    } finally {
      setLoading(false);
    }
  };

  const onChangeVM = async (reportId: string, value: string) => {
    const next = vmDisplayToDb(value);
    setVmSaving((p) => ({ ...p, [reportId]: true }));
    setReports((prev) => prev.map((r) => (r.id === reportId ? { ...r, difficulty: next || null } : r)));

    try {
      await updateReportVM(reportId, next);
    } catch (e: any) {
      await fetchReports(q, sortDir, vmFilter);
      alert(e?.message || String(e));
    } finally {
      setVmSaving((p) => ({ ...p, [reportId]: false }));
    }
  };

  // ✅ On page open, FORCE ascending fetch
  useEffect(() => {
    if (!projectId) return;

    (async () => {
      setLoading(true);
      setLoadError(null);
      try {
        setSortDir("asc");

        const pRes = await fetch(`/api/projects/${encodeURIComponent(projectId)}`, {
          method: "GET",
          credentials: "include",
          headers: authHeaders(),
        });
        const pData = await parseJsonSafe(pRes);
        if (!pRes.ok) throw new Error(pData?.error || "Failed to fetch project");
        setProject((pData?.project || pData?.data || pData) as ProjectRow);

        await fetchReports(q, "asc", vmFilter);
        setSelected({});
      } catch (e: any) {
        setLoadError(e?.message || "Failed to fetch reports");
      } finally {
        setLoading(false);
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [projectId]);

  useEffect(() => {
    if (!projectId) return;

    const t = setTimeout(() => {
      fetchReports(q, sortDir, vmFilter).catch((e: any) => {
        setLoadError(e?.message || "Failed to fetch reports");
      });
      setSelected({});
    }, 250);

    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [q, sortDir, vmFilter, projectId]);

  // close selection menu on outside click
  useEffect(() => {
    function onDown(e: MouseEvent) {
      if (!selMenuRef.current) return;
      if (!selMenuRef.current.contains(e.target as Node)) setSelMenuOpen(false);
    }
    function onKey(e: KeyboardEvent) {
      if (e.key === "Escape") setSelMenuOpen(false);
    }
    if (selMenuOpen) {
      document.addEventListener("mousedown", onDown);
      document.addEventListener("keydown", onKey);
    }
    return () => {
      document.removeEventListener("mousedown", onDown);
      document.removeEventListener("keydown", onKey);
    };
  }, [selMenuOpen]);

  const filteredSortedReports = useMemo(() => reports, [reports]);

  const stats = useMemo(() => {
    const shown = filteredSortedReports.length;

    const last = filteredSortedReports.length
      ? new Date(
          [...filteredSortedReports].sort(
            (a, b) => new Date(b.created_at).getTime() - new Date(a.created_at).getTime()
          )[0].created_at
        ).toLocaleString()
      : "—";

    const selectedCount = filteredSortedReports.filter((r) => selected[r.id]).length;
    return { shown, last, selectedCount };
  }, [filteredSortedReports, selected]);

  const selectedIdsInOrder = useMemo(() => {
    return filteredSortedReports.filter((r) => selected[r.id]).map((r) => r.id);
  }, [filteredSortedReports, selected]);

  const toggleOne = (id: string) => setSelected((prev) => ({ ...prev, [id]: !prev[id] }));

  const selectAllVisible = () => {
    const next: Record<string, boolean> = { ...selected };
    for (const r of filteredSortedReports) next[r.id] = true;
    setSelected(next);
    setSelMenuOpen(false);
  };

  const clearSelection = () => {
    setSelected({});
    setSelMenuOpen(false);
  };

  // ========= Export modal helpers =========
  const openExportModal = () => {
    const baseListed = `${projectName}-${vmFilterLabel(vmFilter)}-${stats.shown}`;
    const baseSelectedOne = `${projectName}-SELECTED-${stats.selectedCount}`;
    const baseSelectedSplit = `${projectName}`;
    const baseAll = `${projectName}-ALL-REPORTS`;

    const defName =
      exportMode === "listed"
        ? baseListed
        : exportMode === "selectedOne"
          ? baseSelectedOne
          : exportMode === "selectedSplit"
            ? baseSelectedSplit
            : baseAll;

    setExportName((p) => (p?.trim() ? p : defName));
    setExportModalOpen(true);
  };

  // ✅ Check GA drawing setup before export, else open GA modal
  const onClickExport = async () => {
    if (!projectId) return;

    try {
      const token =
        typeof window !== "undefined" ? localStorage.getItem("auth_token") : null;
      const res = await fetch(`/api/projects/${encodeURIComponent(projectId)}/export?check=1`, {
        method: "GET",
        credentials: "include",
        headers: token ? { Authorization: `Bearer ${token}` } : {},
      });
      const data = await res.json().catch(() => ({} as any));
      if (!res.ok) throw new Error(data?.error || "Failed to export");

      const pageId = data?.pageId ? String(data.pageId) : null;
      const hasSetup = Boolean(data?.hasSetup);
      const hasImages = Boolean(data?.hasImages);

      if (!hasSetup) {
        setGaSetupReason("GA drawing setup not found for this project. Please fill it before exporting.");
        setGaExistingPageId(null);
        setGaSetupOpen(true);
        return;
      }

      if (!hasImages) {
        setGaSetupReason("GA drawing files are not added for this project. Please upload at least 1 image or PDF.");
        setGaExistingPageId(pageId);
        setGaSetupOpen(true);
        return;
      }

      setGaExistingPageId(pageId);
      setGaChoiceOpen(true);
    } catch (e: any) {
      console.error("[export] failed:", e);
      setDlTitle("Export");
      setDlError("Failed to export");
      setDlDone(true);
      setDlOpen(true);
    }
  };

  // If user switches to GPX, force allowed modes
  useEffect(() => {
    if (!exportModalOpen) return;

    if (exportFormat === "gpx") {
      if (exportMode === "listed" || exportMode === "selectedSplit") {
        setExportMode(stats.selectedCount ? "selectedOne" : "all");
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [exportFormat, exportModalOpen]);

  const closeDl = () => {
    setDlOpen(false);
    setDlTitle("");
    setDlError(null);
    setDlDone(false);
    setDlSecondsLeft(0);
    setPreparedFiles([]);
  };

  // countdown timer (reverse)
  useEffect(() => {
    if (!dlOpen) return;
    if (dlDone) return;

    const t = setInterval(() => {
      setDlSecondsLeft((p) => (p > 0 ? p - 1 : 0));
    }, 1000);

    return () => clearInterval(t);
  }, [dlOpen, dlDone]);

  const startDlUI = (title: string, estSeconds: number) => {
    setDlTitle(title);
    setDlError(null);
    setDlDone(false);
    setPreparedFiles([]);
    setDlSecondsLeft(estSeconds);
    setDlOpen(true);
  };

  
  // ✅ Ensure reports have route_id before DOCX/GPX export (prevents "Points not found" for newly added reports)
  async function ensureRouteIdForMissingReports(reportIds: string[]) {
    if (!projectId) return;
    if (!reportIds.length) return;

    const missing = filteredSortedReports.filter((r) => reportIds.includes(r.id) && !r.route_id);
    if (!missing.length) return;

    const routeData = await apiRequestJson(
      `/api/projects/${encodeURIComponent(projectId)}/reports?sort=desc&limit=1`,
      {
        method: "GET",
      }
    );
    const routeId = String((routeData?.reports || [])[0]?.route_id || "").trim() || undefined;
    if (!routeId) return; // no route exists

    const ids = missing.map((m) => m.id);
    await Promise.all(
      ids.map((id) =>
        apiRequestJson(`/api/reports/${encodeURIComponent(id)}`, {
          method: "PUT",
          body: JSON.stringify({ route_id: routeId }),
        })
      )
    );

    setReports((prev) => prev.map((r) => (ids.includes(r.id) ? { ...r, route_id: routeId } : r)));
  }

  const runExport = async () => {
    if (!projectId) return;

    if (exportMode === "selectedOne" && stats.selectedCount === 0) {
      alert("Please select at least 1 report.");
      return;
    }

    setExportModalOpen(false);

    try {
      if (exportFormat === "gpx") {
        const base = sanitizeFileBaseName(exportName || projectName);
        const fileName = `${base}.gpx`;

        const countForEst = exportMode === "selectedOne" ? stats.selectedCount : Math.max(stats.shown, 1);
        const est = estimateSeconds(exportMode, countForEst, false);
        startDlUI("Preparing GPX export…", est);

        // ✅ attach route_id for any newly added reports included in export
        if (exportMode === "all") {
          await ensureRouteIdForMissingReports(filteredSortedReports.map((r) => r.id));
        } else {
          await ensureRouteIdForMissingReports(selectedIdsInOrder);
        }

        if (exportMode === "all") {
          const { blob, fileName: fn } = await generateProjectGPX(dbClient, projectId, {
            name: exportName || projectName,
            fileName,
          });
          setPreparedFiles([{ fileName: fn, blob }]);
          setDlDone(true);
          return;
        }

        if (exportMode === "selectedOne") {
          const ids = selectedIdsInOrder;
          const { blob, fileName: fn } = await generateProjectGPXByReportIds(dbClient, projectId, ids, {
            name: exportName || projectName,
            fileName,
          });
          setPreparedFiles([{ fileName: fn, blob }]);
          setDlDone(true);
          return;
        }

        throw new Error("GPX supports only: Selected reports or All reports.");
      }

      const count =
        exportMode === "listed"
          ? stats.shown
          : exportMode === "selectedOne" || exportMode === "selectedSplit"
            ? stats.selectedCount
            : Math.max(stats.shown, 1);

      const est = estimateSeconds(exportMode, count, includePhotos);
      startDlUI("Preparing DOCX export…", est);

      // ✅ attach route_id for any newly added reports included in export
      if (exportMode === "all" || exportMode === "listed") {
        await ensureRouteIdForMissingReports(filteredSortedReports.map((r) => r.id));
      } else {
        await ensureRouteIdForMissingReports(selectedIdsInOrder);
      }

      const wm = watermarkOpts.enabled ? watermarkOpts : { enabled: false, text: "" };
      void wm;

      if (exportMode === "all") {
        const fileName = `${sanitizeFileBaseName(exportName || `${projectName}-ALL-REPORTS`)}.docx`;
        const blob = await fetchProjectExportDocx({
          projectId,
          includePhotos,
          fileName,
        });
        setPreparedFiles([{ fileName, blob }]);
        setDlDone(true);
        return;
      }

      if (exportMode === "listed") {
        const ids = filteredSortedReports.map((r) => r.id);
        if (!ids.length) throw new Error("No reports available to export.");

        const fileName = `${sanitizeFileBaseName(
          exportName || `${projectName}-${vmFilterLabel(vmFilter)}-${ids.length}`
        )}.docx`;

        const blob = await fetchProjectExportDocx({
          projectId,
          reportIds: ids,
          includePhotos,
          fileName,
        });

        setPreparedFiles([{ fileName, blob }]);
        setDlDone(true);
        return;
      }

      if (exportMode === "selectedOne") {
        const ids = selectedIdsInOrder;
        const fileName = `${sanitizeFileBaseName(exportName || `${projectName}-SELECTED-${ids.length}`)}.docx`;

        const blob = await fetchProjectExportDocx({
          projectId,
          reportIds: ids,
          includePhotos,
          fileName,
        });

        setPreparedFiles([{ fileName, blob }]);
        setDlDone(true);
        return;
      }

      if (exportMode === "selectedSplit") {
        const ids = selectedIdsInOrder;
        const stages = parseStageRanges(stageRanges, ids.length);

        if (!stages.length) {
          throw new Error(`Invalid stage ranges.\nExample: "1-12,13-14,15-25"\nTotal selected: ${ids.length}`);
        }

        const base = sanitizeFileBaseName(exportName || projectName);
        const files: PreparedFile[] = [];

        for (const st of stages) {
          const subset = ids.slice(st.from - 1, st.to);
          if (!subset.length) continue;

          const fileName = `${base}-${st.label}.docx`;
          const blob = await fetchProjectExportDocx({
            projectId,
            reportIds: subset,
            includePhotos,
            fileName,
          });

          files.push({ fileName, blob });
        }

        if (!files.length) throw new Error("No stage files generated (check your stage ranges).");

        setPreparedFiles(files);
        setDlDone(true);
      }
    } catch (e: any) {
      setDlError(e?.message || String(e));
      setDlDone(true);
    }
  };


  const openManualPoints = (reportId: string) => {
    setManualPointsReportId(reportId);
    setManualPointsOpen(true);
  };

  const closeManualPoints = () => {
    setManualPointsOpen(false);
    setManualPointsReportId(null);
  };

  const saveManualPoints = async (reportId: string, text: string) => {
    const raw = String(text || "").trim();
    if (!raw) throw new Error("Please enter at least 1 point. Format: lat,lon (one per line).");

    const lines = raw
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean);

    const parsed: Array<{ lat: number; lon: number }> = [];
    for (const line of lines) {
      const parts = line.split(/[, \t]+/).map((p) => p.trim()).filter(Boolean);
      if (parts.length < 2) continue;
      const lat = Number(parts[0]);
      const lon = Number(parts[1]);
      if (!Number.isFinite(lat) || !Number.isFinite(lon)) continue;
      if (Math.abs(lat) > 90 || Math.abs(lon) > 180) continue;
      parsed.push({ lat, lon });
    }
    if (!parsed.length) throw new Error('No valid points found. Use format: "12.9716,77.5946" (one per line).');

    const lastRowsRes = await apiRequestJson("/api/db/query", {
      method: "POST",
      body: JSON.stringify({
        table: "report_path_points",
        action: "select",
        select: "seq",
        filters: [{ type: "eq", column: "report_id", value: reportId }],
        orders: [{ column: "seq", ascending: false }],
        limit: 1,
        maybeSingle: true,
      }),
    });
    const lastRow = lastRowsRes?.data || null;

    const authUser = await getAuthUserFromApi();
    const userId = authUser.id;

    const startSeq = (lastRow as any)?.seq ? Number((lastRow as any).seq) + 1 : 1;
    const nowIso = new Date().toISOString();

    const rows: ManualPointRow[] = parsed.map((p, idx) => ({
      report_id: reportId,
      user_id: userId,
      seq: startSeq + idx,
      latitude: p.lat,
      longitude: p.lon,
      elevation: null,
      accuracy: null,
      timestamp: nowIso,
    }));

    await apiRequestJson("/api/[table]".replace("[table]", "report_path_points"), {
      method: "POST",
      body: JSON.stringify(rows),
    });
  };

  // ✅ open insert modal after a row (between rows)
  const openInsertModal = (afterReportId: string) => {
    setInsertAfterId(afterReportId);
    setInsertOpen(true);
  };


  // ✅ open photo upload modal for exactly 1 selected report
  const openPhotoUpload = () => {
    if (stats.selectedCount !== 1) {
      alert("Please select exactly 1 report to upload photos.");
      return;
    }
    const rid = selectedIdsInOrder[0];
    setPhotoReportId(rid);
    setPhotoOpen(true);
  };

  // ✅ re-number all reports in this project to have safe gaps (10,20,30...)
  const renumberSortOrders = async () => {
    if (!projectId) return;

    const data = await apiRequestJson(`/api/projects/${encodeURIComponent(projectId)}/reports?sort=asc`, {
      method: "GET",
    });
    const rows = ((data?.reports || []) as ReportRow[]).slice();
    for (let i = 0; i < rows.length; i++) {
      const id = rows[i].id;
      const nextOrder = (i + 1) * SORT_STEP;
      // update only if different (avoid extra writes)
      if ((rows[i].sort_order ?? null) !== nextOrder) {
        await apiRequestJson(`/api/reports/${encodeURIComponent(id)}`, {
          method: "PUT",
          body: JSON.stringify({ sort_order: nextOrder }),
        });
      }
    }
  };

  // ✅ insert a report exactly after 'afterId'
// ✅ Insert manual points into report_path_points (used by Add Report modal too)
  async function insertManualPoints(reportId: string, text: string) {
    const raw = String(text || "").trim();
    if (!raw) return;

    const lines = raw
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean);

    const parsed: Array<{ lat: number; lon: number }> = [];
    for (const line of lines) {
      const parts = line.split(/[, \t]+/).map((p) => p.trim()).filter(Boolean);
      if (parts.length < 2) continue;
      const lat = Number(parts[0]);
      const lon = Number(parts[1]);
      if (!Number.isFinite(lat) || !Number.isFinite(lon)) continue;
      if (Math.abs(lat) > 90 || Math.abs(lon) > 180) continue;
      parsed.push({ lat, lon });
    }
    if (!parsed.length) return;

    const lastRowsRes = await apiRequestJson("/api/db/query", {
      method: "POST",
      body: JSON.stringify({
        table: "report_path_points",
        action: "select",
        select: "seq",
        filters: [{ type: "eq", column: "report_id", value: reportId }],
        orders: [{ column: "seq", ascending: false }],
        limit: 1,
        maybeSingle: true,
      }),
    });
    const lastRow = lastRowsRes?.data || null;

    const startSeq = (lastRow as any)?.seq ? Number((lastRow as any).seq) + 1 : 1;
    const nowIso = new Date().toISOString();

    const rows = parsed.map((p, idx) => ({
      report_id: reportId,
      seq: startSeq + idx,
      latitude: p.lat,
      longitude: p.lon,
      elevation: null,
      accuracy: null,
      timestamp: nowIso,
    }));

    await apiRequestJson("/api/report_path_points", {
      method: "POST",
      body: JSON.stringify(rows),
    });
  }

  const insertReportAfter = async (afterId: string, payload: { category: string; description: string; remarksAction: string; difficulty: VehicleMovement; files: File[]; pointsText?: string }) => {
    if (!projectId) return;

    // make sure we have latest list order in memory
    const list = filteredSortedReports;
    const idx = list.findIndex((x) => x.id === afterId);
    if (idx === -1) throw new Error("Insert position not found.");

    // Ensure current row has sort_order; if not, renumber first
    const after = list[idx];
    const next = list[idx + 1] || null;

    const afterOrder = Number(after.sort_order);
    const nextOrder = next ? Number(next.sort_order) : NaN;

    if (!Number.isFinite(afterOrder)) {
      await renumberSortOrders();
      await fetchReports(q, sortDir, vmFilter);
      return insertReportAfter(afterId, payload);
    }

    let newOrder: number;

    if (next && Number.isFinite(nextOrder)) {
      const gap = nextOrder - afterOrder;
      if (gap <= 1) {
        // too tight -> renumber then retry
        await renumberSortOrders();
        await fetchReports(q, sortDir, vmFilter);
        return insertReportAfter(afterId, payload);
      }
      newOrder = afterOrder + Math.floor(gap / 2);
    } else {
      // inserting at end
      newOrder = afterOrder + SORT_STEP;
    }

    const authUser = await getAuthUserFromApi();
    const userId = authUser.id;

    const nowIso = new Date().toISOString();

    const ins = await apiRequestJson(`/api/projects/${encodeURIComponent(projectId)}/reports`, {
      method: "POST",
      body: JSON.stringify({
        user_id: userId,
        category: payload.category || "Report",
        description: payload.description || null,
        remarks_action: payload.remarksAction?.trim() || null,
        difficulty: payload.difficulty ? payload.difficulty : "green",
        created_at: nowIso,
        sort_order: newOrder,
      }),
    });

    const newId = String(ins?.report?.id || "").trim();

    // ✅ Photos (optional)
    if (newId && payload.files?.length) {
      await uploadReportPhotos(projectId, newId, payload.files);
    }

    // ✅ Manual points -> report_path_points (optional)
    if (newId && payload.pointsText?.trim()) {
      await insertManualPoints(newId, payload.pointsText);
    }

    // refresh list
    await fetchReports(q, sortDir, vmFilter);
    setSelected({});
  };

  return (
    <div style={styles.containerFluid}>
      <div style={styles.pageInner}>
        {/* ✅ INSERT REPORT MODAL */}
        {insertOpen && projectId && insertAfterId && (
          <InsertReportModal
            afterIndex={filteredSortedReports.findIndex((x) => x.id === insertAfterId) + 1}
            onClose={() => setInsertOpen(false)}
            onCreate={async (p) => {
              try {
                await insertReportAfter(insertAfterId, p);
                setInsertOpen(false);
              } catch (e: any) {
                alert(e?.message || String(e));
              }
            }}
          />
        )}


        {/* ✅ PHOTO UPLOAD MODAL (for selected report) */}
        {photoOpen && projectId && photoReportId && (
          <PhotoUploadModal
            reportId={photoReportId}
            onClose={() => setPhotoOpen(false)}
            onUploaded={async () => {
              setPhotoOpen(false);
              await fetchReports(q, sortDir, vmFilter);
            }}
          />
        )}

        {/* ✅ MANUAL POINTS MODAL */}
        {manualPointsOpen && manualPointsReportId && (
          <ManualPointsModal
            reportId={manualPointsReportId}
            onClose={closeManualPoints}
            onSave={async (text) => {
              try {
                await saveManualPoints(manualPointsReportId, text);
                closeManualPoints();
                alert("Points saved.");
              } catch (e: any) {
                alert(e?.message || String(e));
              }
            }}
          />

        )}

        {/* ✅ GA CHOICE MODAL (Edit or Skip) */}
        {gaChoiceOpen && projectId && (
          <div style={styles.modalOverlay} onMouseDown={() => setGaChoiceOpen(false)}>
            <div
              style={{ ...styles.modalCard, width: "min(560px, 96vw)" }}
              onMouseDown={(e) => e.stopPropagation()}
              role="dialog"
              aria-modal="true"
              aria-label="GA exists"
            >
              <div style={styles.modalTitle}>GA Setup Already Exists</div>

              <div style={styles.modalHint}>
                GA setup already exists for this project. Do you want to <b>Edit</b> it (Objective + Map + Locations +
                GA images + Conclusion) or <b>Skip</b> and continue export?
              </div>


              <div style={styles.modalActions}>
                <button
                  style={styles.btnGhost}
                  onClick={() => {
                    setGaChoiceOpen(false);
                    openExportModal();
                  }}
                >
                  Skip
                </button>

                <button
                  style={styles.btnPrimary}
                  onClick={() => {
                    setGaChoiceOpen(false);
                    setGaSetupReason("Edit GA setup (Objective, Map, Locations, GA images, Conclusion).");
                    setGaSetupOpen(true);
                  }}
                >
                  Edit
                </button>
              </div>

              <div style={styles.modalNote}>Locations inputs are near “Upload map” (4 inputs).</div>
            </div>
          </div>
        )}

        {/* ✅ GA SETUP MODAL (before export) */}
        {gaSetupOpen && projectId && (
          <RouteSetupModal
            projectId={projectId}
            reason={gaSetupReason}
            existingPageId={gaExistingPageId}
            onClose={() => setGaSetupOpen(false)}
            onSaved={() => {
              setGaSetupOpen(false);
              openExportModal();
            }}
          />
        )}

        {/* ========= EXPORT MODAL ========= */}
        {exportModalOpen && (
          <div style={styles.modalOverlay} onMouseDown={() => setExportModalOpen(false)}>
            <div
              style={{ ...styles.modalCard, width: "min(720px, 96vw)" }}
              onMouseDown={(e) => e.stopPropagation()}
              role="dialog"
              aria-modal="true"
              aria-label="Export"
            >
              <div style={styles.modalTitle}>Export</div>

              <div style={styles.modalHint}>
                Generate file first (with an estimated timer), then click <b>Download now</b>.
              </div>

              <div style={{ marginTop: 4 }}>
                <div style={styles.routeLabel}>Format</div>
                <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
                  <label style={styles.radioRow}>
                    <input
                      type="radio"
                      name="fmt"
                      checked={exportFormat === "docx"}
                      onChange={() => setExportFormat("docx")}
                    />
                    <div>
                      <div style={styles.radioTitle}>DOCX</div>
                      <div style={styles.radioSub}>Table export with photos + watermark</div>
                    </div>
                  </label>

                  <label style={styles.radioRow}>
                    <input type="radio" name="fmt" checked={exportFormat === "gpx"} onChange={() => setExportFormat("gpx")} />
                    <div>
                      <div style={styles.radioTitle}>GPX</div>
                      <div style={styles.radioSub}>NE coordinate track points only</div>
                    </div>
                  </label>
                </div>
              </div>

              {/* Mode + options */}
              <div style={styles.exportGrid}>
                <div style={styles.routeLabel}>What to export</div>

                <div style={{ display: "grid", gap: 10 }}>
                  {exportFormat === "gpx" ? (
                    <>
                      <label style={{ ...styles.radioRow, opacity: stats.selectedCount ? 1 : 0.5 }}>
                        <input
                          type="radio"
                          name="mode"
                          disabled={!stats.selectedCount}
                          checked={exportMode === "selectedOne"}
                          onChange={() => setExportMode("selectedOne")}
                        />
                        <div>
                          <div style={styles.radioTitle}>Selected reports (one GPX)</div>
                          <div style={styles.radioSub}>Exports {stats.selectedCount} selected reports</div>
                        </div>
                      </label>

                      <label style={styles.radioRow}>
                        <input type="radio" name="mode" checked={exportMode === "all"} onChange={() => setExportMode("all")} />
                        <div>
                          <div style={styles.radioTitle}>All reports (Project)</div>
                          <div style={styles.radioSub}>Exports every report in this project</div>
                        </div>
                      </label>
                    </>
                  ) : (
                    <>
                      <label style={styles.radioRow}>
                        <input type="radio" name="mode" checked={exportMode === "listed"} onChange={() => setExportMode("listed")} />
                        <div>
                          <div style={styles.radioTitle}>Listed (current filter/search)</div>
                          <div style={styles.radioSub}>Exports {stats.shown} reports currently shown</div>
                        </div>
                      </label>

                      <label style={{ ...styles.radioRow, opacity: stats.selectedCount ? 1 : 0.5 }}>
                        <input
                          type="radio"
                          name="mode"
                          disabled={!stats.selectedCount}
                          checked={exportMode === "selectedOne"}
                          onChange={() => setExportMode("selectedOne")}
                        />
                        <div>
                          <div style={styles.radioTitle}>Selected (one DOCX)</div>
                          <div style={styles.radioSub}>Exports {stats.selectedCount} selected reports into one file</div>
                        </div>
                      </label>

                      <label style={{ ...styles.radioRow, opacity: stats.selectedCount ? 1 : 0.5 }}>
                        <input
                          type="radio"
                          name="mode"
                          disabled={!stats.selectedCount}
                          checked={exportMode === "selectedSplit"}
                          onChange={() => setExportMode("selectedSplit")}
                        />
                        <div>
                          <div style={styles.radioTitle}>Selected (split by stages)</div>
                          <div style={styles.radioSub}>Generates multiple DOCX files (A, B, C…)</div>
                        </div>
                      </label>

                      <label style={styles.radioRow}>
                        <input type="radio" name="mode" checked={exportMode === "all"} onChange={() => setExportMode("all")} />
                        <div>
                          <div style={styles.radioTitle}>All reports (Project)</div>
                          <div style={styles.radioSub}>Exports every report in this project</div>
                        </div>
                      </label>
                    </>
                  )}
                </div>

                {/* Name */}
                <div style={styles.routeLabel}>File name</div>
                <input
                  style={styles.input}
                  value={exportName}
                  onChange={(e) => setExportName(e.target.value)}
                  placeholder="Example: TSPL to Nallur"
                />

                {/* DOCX only: Split ranges */}
                <div style={styles.routeLabel}>Stage split</div>
                <input
                  style={{
                    ...styles.input,
                    opacity: exportFormat === "docx" && exportMode === "selectedSplit" ? 1 : 0.5,
                  }}
                  disabled={exportFormat !== "docx" || exportMode !== "selectedSplit"}
                  value={stageRanges}
                  onChange={(e) => setStageRanges(e.target.value)}
                  placeholder='Example: "1-12,13-14,15-25"'
                />

                {/* DOCX only: Options */}
                <div style={styles.routeLabel}>Options</div>
                {exportFormat === "docx" ? (
                  <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
                    <label style={{ display: "inline-flex", gap: 8, alignItems: "center", fontWeight: 800 }}>
                      <input
                        type="checkbox"
                        checked={includePhotos}
                        onChange={(e) => setIncludePhotos(e.target.checked)}
                        style={{ width: 16, height: 16 }}
                      />
                      Include photos
                    </label>

                    <label style={{ display: "inline-flex", gap: 8, alignItems: "center", fontWeight: 800 }}>
                      <input
                        type="checkbox"
                        checked={wmEnabled}
                        onChange={(e) => setWmEnabled(e.target.checked)}
                        style={{ width: 16, height: 16 }}
                      />
                      Watermark
                    </label>
                  </div>
                ) : (
                  <div style={{ fontSize: 12, fontWeight: 800, color: "#667085" }}>
                    GPX exports only NE coordinate points (no photos / watermark).
                  </div>
                )}

                {/* DOCX only: Watermark text */}
                <div style={styles.routeLabel}>Watermark text</div>
                <input
                  style={{ ...styles.input, opacity: exportFormat === "docx" && wmEnabled ? 1 : 0.5 }}
                  value={wmText}
                  disabled={exportFormat !== "docx" || !wmEnabled}
                  placeholder={`CONFIDENTIAL REPORT: ${projectName}`}
                  onChange={(e) => {
                    setWmDirty(true);
                    setWmText(e.target.value);
                  }}
                />
              </div>

              <div style={styles.modalActions}>
                <button style={styles.btnGhost} onClick={() => setExportModalOpen(false)}>
                  Cancel
                </button>
                <button style={styles.btnPrimary} onClick={runExport}>
                  Generate
                </button>
              </div>

              {exportFormat === "docx" ? (
                <div style={styles.modalNote}>
                  Tip: If export is slow, try disabling <b>Include photos</b>.
                </div>
              ) : (
                <div style={styles.modalNote}>Tip: Select only needed reports for smaller GPX.</div>
              )}
            </div>
          </div>
        )}

        {/* ========= DOWNLOAD PROGRESS MODAL ========= */}
        {dlOpen && (
          <div style={styles.modalOverlay} onMouseDown={() => (dlDone ? closeDl() : null)}>
            <div
              style={{ ...styles.modalCard, width: "min(640px, 96vw)" }}
              onMouseDown={(e) => e.stopPropagation()}
              role="dialog"
              aria-modal="true"
              aria-label="Download progress"
            >
              <div style={styles.modalTitle}>{dlTitle || "Working…"}</div>

              {!dlDone ? (
                <>
                  <div style={styles.modalHint}>
                    Estimated remaining: <b>{dlSecondsLeft}s</b>
                    {dlSecondsLeft === 0 ? (
                      <span style={{ marginLeft: 6, color: "#b42318" }}>(still working…)</span>
                    ) : null}
                  </div>
                  <div style={styles.progressBarOuter}>
                    <div style={styles.progressBarInner} />
                  </div>
                  <div style={styles.modalNote}>Please keep this tab open until generation finishes.</div>
                </>
              ) : dlError ? (
                <>
                  <div style={{ ...styles.modalHint, color: "#b42318" }}>{dlError}</div>
                  <div style={styles.modalActions}>
                    <button style={styles.btnPrimary} onClick={closeDl}>
                      Close
                    </button>
                  </div>
                </>
              ) : (
                <>
                  <div style={styles.modalHint}>Ready. Click download below.</div>

                  <div style={{ display: "grid", gap: 10, marginTop: 8 }}>
                    {preparedFiles.map((f) => (
                      <button
                        key={f.fileName}
                        style={styles.btnPrimary}
                        onClick={() => downloadBlob(f.blob, f.fileName)}
                        title="Download now"
                      >
                        Download: {f.fileName}
                      </button>
                    ))}
                  </div>

                  <div style={styles.modalActions}>
                    <button style={styles.btnGhost} onClick={closeDl}>
                      Close
                    </button>
                  </div>

                  <div style={styles.modalNote}>
                    If your browser blocks multiple downloads (split stages), click each file button one-by-one.
                  </div>
                </>
              )}
            </div>
          </div>
        )}

        {/* ========= HEADER ========= */}
        <div style={styles.headerCard}>
          <div style={{ display: "grid", gap: 6, minWidth: 240 }}>
            <Link href="/projects" style={styles.backLink}>
              ← Back to Projects
            </Link>

            <div style={styles.title}>{projectName}</div>

            <div style={styles.metaRow}>
              <span style={styles.pill}>Project ID: {projectId}</span>
              <span style={styles.pill}>Showing: {stats.shown}</span>
              <span style={styles.pill}>Filter: {vmFilterLabel(vmFilter)}</span>
              <span style={styles.pill}>Last: {stats.last}</span>
              <span style={styles.pill}>Selected: {stats.selectedCount}</span>
            </div>
          </div>

          <div style={styles.actions}>
            <button style={styles.btnGhost} onClick={load} disabled={loading}>
              {loading ? "Refreshing..." : "Refresh"}
            </button>

            <button
              style={styles.btnGhost}
              onClick={() => setSortDir((p) => (p === "asc" ? "desc" : "asc"))}
              disabled={loading}
              title="Toggle ascending/descending"
            >
              Sort: {sortDir === "asc" ? "Ascending" : "Descending"}
            </button>

            {/* ✅ only change: Export button now uses gate */}
            <button style={styles.btnPrimary} onClick={onClickExport} disabled={loading}>
              Export
            </button>
          </div>
        </div>

        {/* ========= REPORTS CONTROLS ========= */}
        <div style={styles.card}>
          <div style={styles.cardHeader}>
            <div style={styles.cardTitle}>Reports</div>
            <div style={styles.cardHint}>{q.trim() ? `Results for “${q.trim()}”` : `Showing ${stats.shown} reports`}</div>
          </div>

          <div style={styles.controlsRow}>
            <div style={{ position: "relative", flex: 1, minWidth: 280 }}>
              <input
                style={{ ...styles.input, paddingRight: 38 }}
                placeholder='Search (FTS): try "bridge", "culvert", "tspl", "red"...'
                value={q}
                onChange={(e) => setQ(e.target.value)}
              />
              {q ? (
                <button style={styles.inputClearBtn} onClick={() => setQ("")} title="Clear search">
                  ×
                </button>
              ) : null}
            </div>

            <select
              value={vmFilter}
              onChange={(e) => setVmFilter(e.target.value as VMFilter)}
              style={styles.select}
              title="Filter by route difficulty"
            >
              <option value="all">Difficulty: All</option>
              <option value="green">Green</option>
              <option value="yellow">Yellow</option>
              <option value="red">Red</option>
              <option value="unset">Not set</option>
            </select>

            <button
              style={styles.btnGhost}
              onClick={() => {
                setQ("");
                setVmFilter("all");
              }}
              disabled={!q && vmFilter === "all"}
              title="Clear search + filter"
            >
              Clear all
            </button>

            {/* Selection dropdown */}
            <div ref={selMenuRef} style={{ position: "relative" }}>
              <button style={styles.btnGhost} onClick={() => setSelMenuOpen((v) => !v)} title="Selection actions">
                Selection ▾
              </button>

              {selMenuOpen && (
                <div style={styles.menu}>
                  <button style={styles.menuItem} onClick={selectAllVisible} disabled={!stats.shown}>
                    Select all listed
                  </button>
                  <button style={styles.menuItem} onClick={clearSelection} disabled={!stats.selectedCount}>
                    Clear selection
                  </button>
                </div>
              )}
            </div>

            <button
              style={styles.btnGhost}
              onClick={openPhotoUpload}
              disabled={stats.selectedCount !== 1}
              title={stats.selectedCount !== 1 ? "Select exactly 1 report" : "Upload photos to selected report"}
            >
              Upload photos
            </button>
          </div>
        </div>

        {/* ========= TABLE ========= */}
        {loading ? (
          <div style={styles.stateCard}>Loading...</div>
        ) : loadError ? (
          <div style={styles.stateCard}>{loadError}</div>
        ) : filteredSortedReports.length === 0 ? (
          <div style={styles.stateCard}>
            <div style={{ fontWeight: 900, color: "#101828" }}>No reports found</div>
            <div style={{ marginTop: 6, color: "#667085", fontWeight: 700 }}>
              Try a different search keyword or filter.
            </div>
          </div>
        ) : (
          <div style={styles.tableCard}>
            <div style={styles.tableWrapNoScroll}>
              <table style={styles.table}>
                <thead>
                  <tr>
                    <th className="col-idx" style={styles.th}>
                      #
                    </th>
                    <th className="col-sel" style={styles.th}>
                      Select
                    </th>
                    <th className="col-cat" style={styles.th}>
                      Category
                    </th>
                    <th className="col-desc" style={styles.th}>
                      Description
                    </th>
                    <th className="col-created" style={styles.th}>
                      Created
                    </th>
                    <th className="col-id" style={styles.th}>
                      Report ID
                    </th>
                    <th className="col-vm" style={styles.th}>
                      Route difficulty
                    </th>
                    <th className="col-act" style={{ ...styles.th, textAlign: "right" }}>
                      Actions
                    </th>
                  </tr>
                </thead>

                <tbody>
                  {filteredSortedReports.map((r, i) => {
                    const created = r.created_at ? new Date(r.created_at).toLocaleString() : "—";
                    const desc = displayDescription((r.description || "").trim());
                    const shortId = r.id ? `${r.id.slice(0, 8)}...` : "—";
                    const vmValue = normalizeVM(r.difficulty);

                    return (
                      <React.Fragment key={r.id}>
                        <tr>
                          <td className="col-idx" style={styles.td}>
                            {i + 1}
                          </td>

                          <td className="col-sel" style={styles.td}>
                            <input
                              type="checkbox"
                              checked={!!selected[r.id]}
                              onChange={() => toggleOne(r.id)}
                              style={{ width: 18, height: 18, cursor: "pointer" }}
                            />
                          </td>

                          <td className="col-cat" style={styles.td}>
                            <div style={styles.catTitle}>{r.category || "Report"}</div>
                            <div style={styles.subtle}>Includes photos</div>
                          </td>

                          <td className="col-desc" style={styles.td}>
                            <div style={styles.descCell}>
                              {desc ? desc : <span style={{ color: "#98A2B3", fontWeight: 800 }}>No description</span>}
                            </div>
                          </td>

                          <td className="col-created" style={styles.td}>
                            <span style={styles.mutedWrap}>{created}</span>
                          </td>

                          <td className="col-id" style={styles.td}>
                            <span style={styles.codePillWrap} title={r.id}>
                              {shortId}
                            </span>
                          </td>

                          <td className="col-vm" style={styles.td}>
                            <select
                              value={vmValue || ""}
                              disabled={!!vmSaving[r.id]}
                              onChange={(e) => onChangeVM(r.id, e.target.value)}
                              style={{
                                height: 38,
                                borderRadius: 14,
                                border: "1px solid #EAECF0",
                                padding: "0 12px",
                                fontWeight: 950,
                                background:
                                  vmValue === "green"
                                    ? "#EAFBF0"
                                    : vmValue === "yellow"
                                      ? "#FFFBEB"
                                      : vmValue === "red"
                                        ? "#FEF2F2"
                                        : "#F2F4F7",
                                color:
                                  vmValue === "green"
                                    ? "#067647"
                                    : vmValue === "yellow"
                                      ? "#92400E"
                                      : vmValue === "red"
                                        ? "#B42318"
                                        : "#475467",
                                cursor: vmSaving[r.id] ? "not-allowed" : "pointer",
                                outline: "none",
                                minWidth: 130,
                                opacity: vmSaving[r.id] ? 0.7 : 1,
                              }}
                              title="Update route difficulty"
                            >
                              <option value="">Not set</option>
                              <option value="green">Green</option>
                              <option value="yellow">Yellow</option>
                              <option value="red">Red</option>
                            </select>
                          </td>

                          <td className="col-act" style={{ ...styles.td, textAlign: "right" }}>
                            <div style={{ display: "inline-flex", gap: 8, justifyContent: "flex-end", flexWrap: "wrap" }}>
                              <button
                                type="button"
                                onClick={() => openManualPoints(r.id)}
                                style={{
                                  padding: "9px 12px",
                                  borderRadius: 12,
                                  border: "1px solid #EAECF0",
                                  background: "#fff",
                                  color: "#344054",
                                  fontWeight: 900,
                                  fontSize: 12,
                                  height: 36,
                                  cursor: "pointer",
                                  whiteSpace: "nowrap",
                                }}
                                title="Add manual lat/long points"
                              >
                                Points
                              </button>

                              <Link href={`/reports/${r.id}`} style={styles.btnOpen} title="Open report">
                                Open
                              </Link>
                            </div>
                          </td>
                        </tr>

                        {/* ✅ BUTTON BETWEEN ROWS (Insert after this row) */}
                        <tr>
                          <td colSpan={8} style={{ padding: 0, borderBottom: "1px solid #F2F4F7" }}>
                            <div style={{ padding: "8px 12px", display: "flex", justifyContent: "center" }}>
                              <button
                                type="button"
                                style={{
                                  padding: "8px 12px",
                                  borderRadius: 12,
                                  border: "1px dashed #D0D5DD",
                                  background: "#fff",
                                  cursor: "pointer",
                                  fontWeight: 900,
                                  fontSize: 12,
                                  color: "#344054",
                                }}
                                onClick={() => openInsertModal(r.id)}
                                title="Insert a new report after this row"
                              >
                                + Add report here (after #{i + 1})
                              </button>
                            </div>
                          </td>
                        </tr>
                      </React.Fragment>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <style jsx>{`
              table {
                width: 100%;
              }
              .col-desc,
              .col-created,
              .col-id,
              .col-vm {
                word-break: break-word;
              }
              .col-idx {
                width: 44px;
              }
              .col-sel {
                width: 72px;
              }
              .col-cat {
                width: 180px;
              }
              .col-created {
                width: 190px;
              }
              .col-id {
                width: 120px;
              }
              .col-vm {
                width: 150px;
              }
              .col-act {
                width: 220px;
              }
              @media (max-width: 1200px) {
                .col-desc {
                  display: none;
                }
              }
              @media (max-width: 992px) {
                .col-id {
                  display: none;
                }
              }
              @media (max-width: 820px) {
                .col-created {
                  display: none;
                }
              }
              @media (max-width: 640px) {
                .col-idx {
                  display: none;
                }
                .col-sel {
                  width: 64px;
                }
                .col-cat {
                  width: 150px;
                }
                .col-vm {
                  width: auto;
                }
                .col-act {
                  width: 160px;
                }
              }
            `}</style>
          </div>
        )}
      </div>
    </div>
  );
}

/** ✅ Small modal to create a new report exactly at a position */
function InsertReportModal({
  afterIndex,
  onClose,
  onCreate,
}: {
  afterIndex: number;
  onClose: () => void;
  onCreate: (payload: { category: string; description: string; remarksAction: string; difficulty: VehicleMovement; files: File[]; pointsText?: string }) => void | Promise<void>;
}) {
  const [category, setCategory] = useState("");
  const [customCategory, setCustomCategory] = useState("");
  const [description, setDescription] = useState("");
  const [remarksAction, setRemarksAction] = useState("");
  const [pointsText, setPointsText] = useState("");
  const [difficulty, setDifficulty] = useState<VehicleMovement>("");
  const [files, setFiles] = useState<File[]>([]);
  const fileRef = useRef<HTMLInputElement | null>(null);
  const [saving, setSaving] = useState(false);

  const addFiles = (list: FileList | null) => {
    if (!list?.length) return;
    const next = Array.from(list);
    const existingKeys = new Set(files.map((f) => `${f.name}__${f.size}`));
    const filtered = next.filter((f) => !existingKeys.has(`${f.name}__${f.size}`));
    setFiles((p) => [...p, ...filtered]);
  };

  const submit = async () => {
    setSaving(true);
    try {
      const finalCategory =
        category === "__custom__"
          ? (customCategory.trim() || "Report")
          : (category.trim() || "Report");

      await onCreate({
      category: finalCategory,
      description: description.trim(),
      remarksAction: remarksAction.trim(),
      difficulty,
      files,
      pointsText: pointsText.trim(),
    });
    } finally {
      setSaving(false);
    }
  };

  return (
    <div style={styles.modalOverlay} onMouseDown={onClose}>
      <div
        style={{ ...styles.modalCard, width: "min(620px, 96vw)" }}
        onMouseDown={(e) => e.stopPropagation()}
        role="dialog"
        aria-modal="true"
        aria-label="Insert report"
      >
        <div style={styles.modalTitle}>Add Report (insert after #{afterIndex})</div>
        <div style={styles.modalHint}>
          This will create a new report and place it immediately after the selected row.
        </div>

        <div style={{ display: "grid", gap: 8 }}>
          <div style={styles.routeLabel}>Category</div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <select
              value={category}
              onChange={(e) => {
                const v = e.target.value;
                setCategory(v);
                if (v !== "__custom__") setCustomCategory("");
              }}
              style={{
                height: 44,
                borderRadius: 14,
                border: "1px solid #D0D5DD",
                padding: "0 14px",
                fontWeight: 900,
                outline: "none",
                background: "#fff",
              }}
            >
              <option value="">Select category…</option>
              {CATEGORY_OPTIONS.map((c) => (
                <option key={c} value={c}>
                  {c}
                </option>
              ))}
              <option value="__custom__">+ Custom…</option>
            </select>

            <input
              value={customCategory}
              onChange={(e) => setCustomCategory(e.target.value)}
              disabled={category !== "__custom__"}
              placeholder="Type custom category"
              style={{
                height: 44,
                borderRadius: 14,
                border: "1px solid #D0D5DD",
                padding: "0 14px",
                fontWeight: 900,
                outline: "none",
                background: "#fff",
                opacity: category === "__custom__" ? 1 : 0.6,
              }}
            />
          </div>
        </div>

        <div style={{ display: "grid", gap: 8 }}>
          <div style={styles.routeLabel}>Description</div>
          <textarea
            value={description}
            onChange={(e) => setDescription(e.target.value)}
            placeholder="Optional description..."
            style={{
              width: "100%",
              minHeight: 90,
              borderRadius: 14,
              border: "1px solid #D0D5DD",
              padding: "12px 12px",
              fontSize: 14,
              fontWeight: 700,
              outline: "none",
              resize: "vertical",
            }}
          />
        </div>


        <div style={{ display: "grid", gap: 8 }}>
          <div style={styles.routeLabel}>Remarks / Action</div>
          <textarea
            value={remarksAction}
            onChange={(e) => setRemarksAction(e.target.value)}
            placeholder="Type remarks / action"
            style={{
              width: "100%",
              minHeight: 90,
              borderRadius: 14,
              border: "1px solid #D0D5DD",
              padding: "12px 14px",
              fontSize: 14,
              fontWeight: 700,
              outline: "none",
              resize: "vertical",
            }}
          />
        </div>

        <div style={{ display: "grid", gap: 8 }}>
          <div style={styles.routeLabel}>Route difficulty</div>
          <select
            value={difficulty || ""}
            onChange={(e) => setDifficulty(vmDisplayToDb(e.target.value))}
            style={{
              height: 44,
              borderRadius: 14,
              border: "1px solid #D0D5DD",
              padding: "0 14px",
              fontWeight: 950,
              background:
                difficulty === "green"
                  ? "#EAFBF0"
                  : difficulty === "yellow"
                    ? "#FFFBEB"
                    : difficulty === "red"
                      ? "#FEF2F2"
                      : "#fff",
              color:
                difficulty === "green"
                  ? "#067647"
                  : difficulty === "yellow"
                    ? "#92400E"
                    : difficulty === "red"
                      ? "#B42318"
                      : "#344054",
              outline: "none",
              cursor: "pointer",
            }}
            title="Set difficulty for new report"
          >
            <option value="">Not set</option>
            <option value="green">Green</option>
            <option value="yellow">Yellow</option>
            <option value="red">Red</option>
          </select>
        </div>

        <div style={{ display: "grid", gap: 8 }}>
          <div style={styles.routeLabel}>Photos (optional)</div>

          <input
            ref={fileRef}
            type="file"
            accept="image/*,.pdf,application/pdf"
            multiple
            style={{ display: "none" }}
            onChange={(e) => {
              addFiles(e.target.files);
              if (fileRef.current) fileRef.current.value = "";
            }}
          />

          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <button type="button" style={styles.btnPrimary} onClick={() => fileRef.current?.click()} disabled={saving}>
              Choose photos
            </button>
            <button type="button" style={styles.btnGhost} onClick={() => setFiles([])} disabled={!files.length || saving}>
              Clear
            </button>
          </div>

          {files.length ? (
            <div style={{ fontSize: 12, fontWeight: 800, color: "#667085" }}>
              Selected: {files.length} file(s)
            </div>
          ) : (
            <div style={{ fontSize: 12, fontWeight: 800, color: "#98A2B3" }}>No photos selected.</div>
          )}
        </div>


        <div style={{ display: "grid", gap: 8 }}>
          <div style={styles.routeLabel}>Points (optional)</div>
          <div style={styles.modalHint}>
            Paste <b>lat,lon</b> (one per line). Example:{" "}
            <span style={{ fontFamily: "ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace" }}>
              12.9716,77.5946
            </span>
          </div>

          <textarea
            value={pointsText}
            onChange={(e) => setPointsText(e.target.value)}
            placeholder="lat,lon (one per line)"
            style={{
              width: "100%",
              minHeight: 120,
              borderRadius: 14,
              border: "1px solid #D0D5DD",
              padding: "12px 12px",
              fontSize: 14,
              fontWeight: 700,
              outline: "none",
              resize: "vertical",
            }}
          />
        </div>

        <div style={styles.modalActions}>
          <button style={styles.btnGhost} onClick={onClose} disabled={saving}>
            Cancel
          </button>
          <button style={styles.btnPrimary} onClick={submit} disabled={saving}>
            {saving ? "Adding..." : "Add Report"}
          </button>
        </div>

        <div style={styles.modalNote}>
          Note: This requires <b>reports.sort_order</b> column (integer/number). Existing reports should be backfilled with
          10,20,30... for stable insertion.
        </div>
      </div>
    </div>
  );
}

/** ✅ Modal to upload photos for one report (from list page) */
function PhotoUploadModal({
  reportId,
  onClose,
  onUploaded,
}: {
  reportId: string;
  onClose: () => void;
  onUploaded: () => void | Promise<void>;
}) {
  const params = useParams();
  const projectId = useMemo(() => {
    const id = (params as any)?.id;
    return Array.isArray(id) ? id[0] : id;
  }, [(params as any)?.id]);

  const [files, setFiles] = useState<File[]>([]);
  const [saving, setSaving] = useState(false);

  const inputRef = useRef<HTMLInputElement | null>(null);

  const addFiles = (list: FileList | null) => {
    if (!list?.length) return;
    const next = Array.from(list);
    const existingKeys = new Set(files.map((f) => `${f.name}__${f.size}`));
    const filtered = next.filter((f) => !existingKeys.has(`${f.name}__${f.size}`));
    setFiles((p) => [...p, ...filtered]);
  };

  const removeAt = (idx: number) => setFiles((p) => p.filter((_, i) => i !== idx));

  const upload = async () => {
    if (!projectId) return;
    if (!files.length) {
      alert("Please choose at least 1 photo.");
      return;
    }

    setSaving(true);
    try {
      await uploadReportPhotos(projectId, reportId, files);
      await onUploaded();
    } catch (e: any) {
      alert(e?.message || String(e));
    } finally {
      setSaving(false);
    }
  };

  return (
    <div style={styles.modalOverlay} onMouseDown={onClose}>
      <div
        style={{ ...styles.modalCard, width: "min(760px, 96vw)" }}
        onMouseDown={(e) => e.stopPropagation()}
        role="dialog"
        aria-modal="true"
        aria-label="Upload photos"
      >
        <div style={styles.modalTitle}>Upload Photos</div>
        <div style={styles.modalHint}>
          Upload photos for this report. (Bucket: <b>{REPORT_PHOTO_BUCKET}</b>, Table: <b>{REPORT_IMAGE_TABLE}</b>)
        </div>

        <input
          ref={inputRef}
          type="file"
          accept="image/*"
          multiple
          style={{ display: "none" }}
          onChange={(e) => {
            addFiles(e.target.files);
            if (inputRef.current) inputRef.current.value = "";
          }}
        />

        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <button type="button" style={styles.btnPrimary} onClick={() => inputRef.current?.click()} disabled={saving}>
            Choose photos
          </button>
          <button type="button" style={styles.btnGhost} onClick={() => setFiles([])} disabled={!files.length || saving}>
            Clear
          </button>
        </div>

        {!files.length ? (
          <div
            style={{
              padding: 14,
              borderRadius: 14,
              border: "1px dashed #D0D5DD",
              background: "#F9FAFB",
              fontWeight: 800,
              color: "#667085",
              marginTop: 6,
            }}
          >
            No photos selected.
          </div>
        ) : (
          <div style={{ display: "grid", gap: 10, marginTop: 6 }}>
            {files.map((f, idx) => (
              <div
                key={`${f.name}-${idx}`}
                style={{
                  border: "1px solid #EAECF0",
                  borderRadius: 14,
                  padding: 10,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  gap: 10,
                }}
              >
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontWeight: 950, color: "#101828", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                    {f.name}
                  </div>
                  <div style={{ fontSize: 12, fontWeight: 800, color: "#667085" }}>{(f.size / 1024).toFixed(1)} KB</div>
                </div>
                <button
                  type="button"
                  onClick={() => removeAt(idx)}
                  style={{
                    padding: "8px 10px",
                    borderRadius: 12,
                    border: "1px solid #FECDD6",
                    background: "#FFF1F3",
                    color: "#B42318",
                    fontWeight: 900,
                    cursor: "pointer",
                  }}
                  disabled={saving}
                >
                  Remove
                </button>
              </div>
            ))}
          </div>
        )}

        <div style={styles.modalActions}>
          <button style={styles.btnGhost} onClick={onClose} disabled={saving}>
            Cancel
          </button>
          <button style={styles.btnPrimary} onClick={upload} disabled={saving}>
            {saving ? "Uploading..." : "Upload"}
          </button>
        </div>
      </div>
    </div>
  );
}


/** ✅ Modal that collects Objective + Map + GA images + Locations + Conclusion and saves to dbClient */
function RouteSetupModal({
  projectId,
  reason,
  existingPageId,
  onClose,
  onSaved,
}: {
  projectId: string;
  reason: string;
  existingPageId?: string | null;
  onClose: () => void;
  onSaved: () => void;
}) {
  const [saving, setSaving] = useState(false);
  const [err, setErr] = useState<string | null>(null);

  const [objective, setObjective] = useState("");

  const [loadingExisting, setLoadingExisting] = useState(false);
  const [existingMapUrl, setExistingMapUrl] = useState<string>("");

  // ✅ Conclusion HTML (TinyMCE)
  const [conclusionHtml, setConclusionHtml] = useState<string>("");
  const conclusionRef = useRef<any>(null);

  // ✅ Locations (4 inputs)
  const [routeLocations, setRouteLocations] = useState<string[]>(["", "", "", ""]);
  const updateLoc = (idx: number, val: string) => {
    setRouteLocations((p) => {
      const copy = [...p];
      copy[idx] = val;
      return copy;
    });
  };

  const presetMaps: PresetMap[] = useMemo(
    () => [
      // If you want preset maps, add here:
      // { id: "route1", label: "Route Map 1", src: "/maps/route1.png" },
    ],
    []
  );

  const [mapPickMode, setMapPickMode] = useState<PickMode>("preset");
  const [selectedPresetMapId, setSelectedPresetMapId] = useState(presetMaps[0]?.id || "");
  const [uploadedMapFile, setUploadedMapFile] = useState<File | null>(null);
  const mapInputRef = useRef<HTMLInputElement | null>(null);

  const uploadedMapPreview = useMemo(() => {
    if (!uploadedMapFile) return "";
    return URL.createObjectURL(uploadedMapFile);
  }, [uploadedMapFile]);

  const selectedPresetMap = useMemo(() => {
    return presetMaps.find((m) => m.id === selectedPresetMapId) || null;
  }, [presetMaps, selectedPresetMapId]);

  const finalMapPreview = mapPickMode === "upload" ? (uploadedMapPreview || existingMapUrl) : selectedPresetMap?.src || "";

  const [gaFiles, setGaFiles] = useState<File[]>([]);
  const gaInputRef = useRef<HTMLInputElement | null>(null);

  const gaPreviews = useMemo(() => {
    return gaFiles.map((f) => ({
      name: f.name,
      size: f.size,
      url: URL.createObjectURL(f),
    }));
  }, [gaFiles]);

  const addGaFiles = (files: FileList | null) => {
    if (!files?.length) return;
    const next = Array.from(files);
    const existingKeys = new Set(gaFiles.map((f) => `${f.name}__${f.size}`));
    const filtered = next.filter((f) => !existingKeys.has(`${f.name}__${f.size}`));
    setGaFiles((p) => [...p, ...filtered]);
  };

  const removeGaAt = (idx: number) => setGaFiles((p) => p.filter((_, i) => i !== idx));

  const isPdfFile = (file: File) => {
    const name = String(file?.name || "").toLowerCase();
    const type = String(file?.type || "").toLowerCase();
    return type === "application/pdf" || name.endsWith(".pdf");
  };

  const PDF_JS_PROMISE_KEY = "__gaUploadPdfJsPromise";
  const PDF_JS_SCRIPT_ID = "ga-upload-pdfjs-script";
  const PDF_JS_CDN = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
  const PDF_JS_WORKER_CDN = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

  async function getPdfJsForGaUpload() {
    if (typeof window === "undefined") throw new Error("PDF conversion works only in browser.");

    const win = window as any;
    if (win.pdfjsLib) {
      if (!win.pdfjsLib.GlobalWorkerOptions.workerSrc) {
        win.pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_JS_WORKER_CDN;
      }
      return win.pdfjsLib;
    }

    if (!win[PDF_JS_PROMISE_KEY]) {
      win[PDF_JS_PROMISE_KEY] = new Promise((resolve, reject) => {
        const existing = document.getElementById(PDF_JS_SCRIPT_ID) as HTMLScriptElement | null;

        const finish = () => {
          if (win.pdfjsLib) {
            if (!win.pdfjsLib.GlobalWorkerOptions.workerSrc) {
              win.pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_JS_WORKER_CDN;
            }
            resolve(win.pdfjsLib);
          } else {
            reject(new Error("PDF library failed to load."));
          }
        };

        if (existing) {
          if ((existing as any).dataset.loaded === "true") {
            finish();
            return;
          }
          existing.addEventListener("load", finish, { once: true });
          existing.addEventListener("error", () => reject(new Error("Failed to load PDF library.")), { once: true });
          return;
        }

        const script = document.createElement("script");
        script.id = PDF_JS_SCRIPT_ID;
        script.src = PDF_JS_CDN;
        script.async = true;
        script.onload = () => {
          (script as any).dataset.loaded = "true";
          finish();
        };
        script.onerror = () => reject(new Error("Failed to load PDF library."));
        document.head.appendChild(script);
      });
    }

    return win[PDF_JS_PROMISE_KEY];
  }

  async function convertPdfFirstPageToPngFile(file: File): Promise<File> {
    const pdfjsLib = await getPdfJsForGaUpload();
    const buffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
    const page = await pdf.getPage(1);
    const viewport = page.getViewport({ scale: 2.2 });

    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");
    if (!context) throw new Error("Unable to create canvas for PDF conversion.");

    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);

    await page.render({ canvasContext: context, viewport }).promise;

    const pngBlob: Blob = await new Promise((resolve, reject) => {
      canvas.toBlob((blob) => {
        if (blob) resolve(blob);
        else reject(new Error("Failed to convert PDF to PNG."));
      }, "image/png");
    });

    const baseName = file.name.replace(/\.[^.]+$/i, "") || "ga_drawing";
    return new File([pngBlob], `${baseName}.png`, {
      type: "image/png",
      lastModified: Date.now(),
    });
  }


  // ✅ Load existing setup when editing (best-effort)
  useEffect(() => {
    if (!existingPageId) return;

    (async () => {
      setLoadingExisting(true);
      setErr(null);

      try {
        const token =
          typeof window !== "undefined" ? localStorage.getItem("auth_token") : null;
        const res = await fetch(
          `/api/projects/${encodeURIComponent(projectId)}/ga-drawing?pageId=${encodeURIComponent(existingPageId)}`,
          {
            method: "GET",
            credentials: "include",
            headers: token ? { Authorization: `Bearer ${token}` } : {},
          }
        );
        const data = await res.json().catch(() => ({} as any));
        if (!res.ok) throw new Error(data?.error || "Failed to load GA drawing");
        const page = data?.page || null;

        if (page) {
          setObjective((page as any).objective || "");
          setMapPickMode(((page as any).map_mode as PickMode) || "preset");
          setSelectedPresetMapId((page as any).preset_map_key || presetMaps[0]?.id || "");
          setExistingMapUrl((page as any).map_file_url || "");
          setConclusionHtml((page as any).conclusion_html || "");
        }

        const locs = Array.isArray(data?.locations) ? (data.locations as RouteLocationRow[]) : [];
        if (locs.length) {
          const labels = locs.map((x) => String(x.label || "").trim());
          setRouteLocations([labels[0] || "", labels[1] || "", labels[2] || "", labels[3] || ""]);
        }
      } catch (e: any) {
        setErr(e?.message || String(e));
      } finally {
        setLoadingExisting(false);
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [existingPageId]);

  async function uploadToBucket(bucket: string, path: string, file: File) {
    const normalizedPath = String(path || "").replace(/^\/+/, "").replace(/\\/g, "/");
    const slash = normalizedPath.lastIndexOf("/");
    const folderFromPath = slash > 0 ? normalizedPath.slice(0, slash) : normalizedPath || "uploads";
    const formData = new FormData();
    formData.append("folder", folderFromPath);
    formData.append("file", file);

    const token =
      typeof window !== "undefined" ? localStorage.getItem("auth_token") : null;

    const res = await fetch("/api/upload", {
      method: "POST",
      credentials: "include",
      headers: token ? { Authorization: `Bearer ${token}` } : {},
      body: formData,
    });
    const data = await res.json().catch(() => ({} as any));
    if (!res.ok) throw new Error(data?.error || "Failed to upload image");
    return {
      path: String(data?.key || data?.path || path),
      publicUrl: String(data?.url || ""),
    };
  }

  const saveAll = async () => {
    setSaving(true);
    setErr(null);

    try {
      // Must have at least 1 GA image to proceed
      if (!gaFiles.length) {
        throw new Error("Please upload at least 1 GA drawing file (image or PDF).");
      }

      let preset_map_key: string | null = null;
      let map_file_url: string | null = null;

      if (mapPickMode === "preset") {
        preset_map_key = selectedPresetMapId || null;
        map_file_url = null;
      } else {
        preset_map_key = null;

        if (uploadedMapFile) {
          const ext = uploadedMapFile.name.split(".").pop() || "png";
          const storagePath = `projects/${projectId}/${Date.now()}_map.${ext}`;
          const uploaded = await uploadToBucket("route-maps", storagePath, uploadedMapFile);
          map_file_url = uploaded.publicUrl;
        } else if (existingMapUrl) {
          map_file_url = existingMapUrl;
        } else {
          throw new Error("Please upload a route map image.");
        }
      }

      const currentConclusion = (conclusionRef.current?.getContent?.() ?? conclusionHtml ?? "").trim();

      // Upload GA files + insert rows
      // ✅ Important fix:
      // If user uploads a PDF GA drawing, convert page 1 to PNG first and store that PNG in storage + DB.
      // This makes project_route_page_images behave like normal image rows, so DOCX export can retrieve them directly.
      const rows: any[] = [];
      for (const originalFile of gaFiles) {
        const storageFile = isPdfFile(originalFile)
          ? await convertPdfFirstPageToPngFile(originalFile)
          : originalFile;

        const safeName = storageFile.name.replace(/[^\w.\-]+/g, "_");
        const storagePath = `projects/${projectId}/${Date.now()}_${safeName}`;
        const uploaded = await uploadToBucket("ga-drawings", storagePath, storageFile);

        rows.push({
          file_url: uploaded.publicUrl,
          imageUrl: uploaded.publicUrl,
          imageKey: uploaded.path,
          file_name: storageFile.name,
          fileName: storageFile.name,
          mime_type: storageFile.type || "image/png",
          mimeType: storageFile.type || "image/png",
          file_size: storageFile.size || null,
          fileSize: storageFile.size || null,
        });
      }

      const token =
        typeof window !== "undefined" ? localStorage.getItem("auth_token") : null;
      const saveRes = await fetch(`/api/projects/${encodeURIComponent(projectId)}/ga-drawing`, {
        method: "POST",
        credentials: "include",
        headers: {
          "Content-Type": "application/json",
          ...(token ? { Authorization: `Bearer ${token}` } : {}),
        },
        body: JSON.stringify({
          pageId: existingPageId || null,
          objective,
          mapMode: mapPickMode,
          presetMapKey: preset_map_key,
          mapFileUrl: map_file_url,
          conclusionHtml: currentConclusion || null,
          routeLocations,
          gaImages: rows,
          imageUrl: rows[0]?.file_url || null,
          imageKey: rows[0]?.imageKey || null,
          fileName: rows[0]?.file_name || null,
        }),
      });
      const saveData = await saveRes.json().catch(() => ({} as any));
      if (!saveRes.ok) {
        throw new Error(saveData?.error || "Failed to save GA drawing");
      }

      onSaved();
    } catch (e: any) {
      setErr(e?.message || "Failed to save GA drawing");
    } finally {
      setSaving(false);
    }
  };

  // Use same overlay style as your page (NO change to existing styles object)
  return (
    <div style={styles.modalOverlay} onMouseDown={onClose}>
      <div
        style={{ ...styles.modalCard, width: "min(920px, 96vw)", maxHeight: "92vh", overflow: "auto" as any }}
        onMouseDown={(e) => e.stopPropagation()}
        role="dialog"
        aria-modal="true"
        aria-label="GA Drawing Setup"
      >
        <div style={styles.modalTitle}>GA Drawing Setup Required</div>

        <div style={{ ...styles.modalHint, color: "#b42318", fontWeight: 900 }}>
          {reason || "Please complete GA Drawing setup to export."}
        </div>

        {loadingExisting ? <div style={styles.modalHint}>Loading existing setup…</div> : null}

        {err ? <div style={{ ...styles.modalHint, color: "#b42318", fontWeight: 900 }}>{err}</div> : null}

        {/* Objective */}
        <div style={{ display: "grid", gap: 8 }}>
          <div style={styles.routeLabel}>Objective</div>
          <textarea
            value={objective}
            onChange={(e) => setObjective(e.target.value)}
            placeholder="Type objective..."
            style={{
              width: "100%",
              minHeight: 90,
              borderRadius: 14,
              border: "1px solid #EAECF0",
              padding: "12px 12px",
              fontSize: 14,
              fontWeight: 700,
              outline: "none",
            }}
          />
        </div>

        {/* Map picker */}
        <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
          <div style={styles.routeLabel}>Route Map</div>

          <button
            type="button"
            style={{
              ...styles.btnGhost,
              height: 36,
              borderRadius: 12,
              background: mapPickMode === "preset" ? "#111" : "#fff",
              color: mapPickMode === "preset" ? "#fff" : "#344054",
              borderColor: mapPickMode === "preset" ? "#111" : "#EAECF0",
            }}
            onClick={() => setMapPickMode("preset")}
          >
            Preset
          </button>

          <button
            type="button"
            style={{
              ...styles.btnGhost,
              height: 36,
              borderRadius: 12,
              background: mapPickMode === "upload" ? "#111" : "#fff",
              color: mapPickMode === "upload" ? "#fff" : "#344054",
              borderColor: mapPickMode === "upload" ? "#111" : "#EAECF0",
            }}
            onClick={() => setMapPickMode("upload")}
          >
            Upload
          </button>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "340px 1fr", gap: 12, alignItems: "start" }}>
          <div style={{ border: "1px solid #EAECF0", borderRadius: 16, padding: 12 }}>
            {mapPickMode === "preset" ? (
              <>
                <div style={styles.routeLabel}>Select preset map</div>
                <select
                  value={selectedPresetMapId}
                  onChange={(e) => setSelectedPresetMapId(e.target.value)}
                  disabled={!presetMaps.length}
                  style={{
                    marginTop: 8,
                    width: "100%",
                    height: 40,
                    borderRadius: 12,
                    border: "1px solid #EAECF0",
                    padding: "0 10px",
                    fontWeight: 900,
                    background: "#fff",
                  }}
                >
                  {presetMaps.length ? (
                    presetMaps.map((m) => (
                      <option key={m.id} value={m.id}>
                        {m.label}
                      </option>
                    ))
                  ) : (
                    <option value="">No preset maps</option>
                  )}
                </select>
              </>
            ) : (
              <>
                <div style={styles.routeLabel}>Upload map</div>
                <input
                  ref={mapInputRef}
                  type="file"
                  accept="image/*"
                  style={{ display: "none" }}
                  onChange={(e) => setUploadedMapFile(e.target.files?.[0] || null)}
                />
                <button type="button" style={{ ...styles.btnPrimary, width: "100%" }} onClick={() => mapInputRef.current?.click()}>
                  Choose map file
                </button>
                {uploadedMapFile ? (
                  <div style={{ marginTop: 8, fontSize: 12, fontWeight: 800, color: "#475467" }}>
                    Selected: {uploadedMapFile.name}
                  </div>
                ) : null}
              </>
            )}

            {/* ✅ Locations (4 inputs) */}
            <div style={{ display: "grid", gap: 8, marginTop: 10 }}>
              <div style={styles.routeLabel}>Locations</div>

              <div style={{ display: "grid", gridTemplateColumns: "36px 1fr", gap: 12, alignItems: "start" }}>
                <div style={{ display: "grid", gap: 14, paddingTop: 8, justifyItems: "center" }}>
                  <div style={locStyles.circle} />
                  <div style={locStyles.dots} />
                  <div style={locStyles.circle} />
                  <div style={locStyles.dots} />
                  <div style={locStyles.circle} />
                  <div style={locStyles.dots} />
                  <div style={locStyles.pin}>📍</div>
                </div>

                <div style={{ display: "grid", gap: 12 }}>
                  {routeLocations.map((val, idx) => (
                    <input
                      key={idx}
                      value={val}
                      onChange={(e) => updateLoc(idx, e.target.value)}
                      placeholder={idx === 0 ? "Start location" : idx === 3 ? "End location" : "Stop location"}
                      style={{
                        height: 44,
                        borderRadius: 14,
                        border: "1px solid #D0D5DD",
                        padding: "0 14px",
                        fontWeight: 900,
                        outline: "none",
                        background: "#fff",
                      }}
                    />
                  ))}
                </div>
              </div>

              <div style={styles.modalNote}>Example: Kappalur → Salem → Krishnagiri → Hosur</div>
            </div>
          </div>

          <div style={{ border: "1px solid #EAECF0", borderRadius: 16, padding: 12, background: "#F9FAFB" }}>
            <div style={styles.routeLabel}>Map preview</div>
            <div
              style={{
                marginTop: 8,
                border: "1px solid #EAECF0",
                borderRadius: 14,
                overflow: "hidden",
                background: "#fff",
                height: MAP_PREVIEW_HEIGHT,
              }}
            >
              {finalMapPreview ? (
                // eslint-disable-next-line @next/next/no-img-element
                <img
                  src={finalMapPreview}
                  alt="Map preview"
                  style={{ width: "100%", height: "100%", objectFit: "contain" }}
                />
              ) : (
                <div
                  style={{
                    height: MAP_PREVIEW_HEIGHT,
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    fontWeight: 800,
                    color: "#667085",
                  }}
                >
                  No map selected
                </div>
              )}
            </div>
          </div>
        </div>

        {/* GA files */}
        <div style={{ display: "grid", gap: 8 }}>
          <div style={styles.routeLabel}>GA Drawing Files (required)</div>

          <input
            ref={gaInputRef}
            type="file"
            accept="image/*,.pdf,application/pdf"
            multiple
            style={{ display: "none" }}
            onChange={(e) => {
              const files = e.target.files;
              if (!files?.length) return;
              addGaFiles(files);
              if (gaInputRef.current) gaInputRef.current.value = "";
            }}
          />

          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <button type="button" style={styles.btnPrimary} onClick={() => gaInputRef.current?.click()}>
              Add GA files
            </button>
            <button type="button" style={styles.btnGhost} onClick={() => setGaFiles([])} disabled={!gaFiles.length}>
              Clear
            </button>
          </div>

          {!gaFiles.length ? (
            <div
              style={{
                padding: 14,
                borderRadius: 14,
                border: "1px dashed #D0D5DD",
                background: "#F9FAFB",
                fontWeight: 800,
                color: "#667085",
              }}
            >
              No GA files selected.
            </div>
          ) : (
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              {gaPreviews.map((p, idx) => {
                const file = gaFiles[idx];
                const isPdf = file ? isPdfFile(file) : false;

                return (
                  <div key={`${p.name}-${idx}`} style={{ border: "1px solid #EAECF0", borderRadius: 16, padding: 12 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 10 }}>
                      <div style={{ minWidth: 0 }}>
                        <div
                          style={{
                            fontWeight: 950,
                            color: "#101828",
                            whiteSpace: "nowrap",
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                          }}
                        >
                          {p.name}
                        </div>
                        <div style={{ fontSize: 12, fontWeight: 800, color: "#667085" }}>
                          {(p.size / 1024).toFixed(1)} KB {isPdf ? "• PDF" : "• Image"}
                        </div>
                      </div>
                      <button
                        type="button"
                        onClick={() => removeGaAt(idx)}
                        style={{
                          padding: "8px 10px",
                          borderRadius: 12,
                          border: "1px solid #FECDD6",
                          background: "#FFF1F3",
                          color: "#B42318",
                          fontWeight: 900,
                          cursor: "pointer",
                        }}
                      >
                        Remove
                      </button>
                    </div>

                    <div
                      style={{
                        marginTop: 10,
                        borderRadius: 14,
                        overflow: "hidden",
                        border: "1px solid #EAECF0",
                        background: "#F9FAFB",
                        height: GA_PREVIEW_HEIGHT,
                      }}
                    >
                      {isPdf ? (
                        <div
                          style={{
                            width: "100%",
                            height: "100%",
                            display: "grid",
                            placeItems: "center",
                            padding: 16,
                            textAlign: "center",
                            gap: 10,
                          }}
                        >
                          <div style={{ fontSize: 44 }}>📄</div>
                          <div style={{ fontWeight: 900, color: "#101828" }}>PDF selected</div>
                          <a
                            href={p.url}
                            target="_blank"
                            rel="noreferrer"
                            style={{
                              textDecoration: "none",
                              padding: "10px 12px",
                              borderRadius: 12,
                              border: "1px solid #D0D5DD",
                              color: "#344054",
                              fontWeight: 900,
                              background: "#fff",
                            }}
                          >
                            Open PDF Preview
                          </a>
                        </div>
                      ) : (
                        <>
                          {/* eslint-disable-next-line @next/next/no-img-element */}
                          <img src={p.url} alt="GA preview" style={{ width: "100%", height: "100%", objectFit: "contain" }} />
                        </>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        {/* Conclusion */}
        <div style={{ display: "grid", gap: 8 }}>
          <div style={styles.routeLabel}>Conclusion & Certification</div>

          <div style={{ border: "1px solid #EAECF0", borderRadius: 16, overflow: "hidden", background: "#fff" }}>
            <Editor
              apiKey="3fr142nwyhd2jop9d509ekq6i2ks2u6dmrbgm8c74gu5xrml"
              onInit={(_evt, editor) => (conclusionRef.current = editor)}
              value={conclusionHtml}
              onEditorChange={(v) => setConclusionHtml(v)}
              init={{
                height: 260,
                menubar: false,
                branding: false,
                statusbar: false,
                plugins: "lists",
                toolbar:
                  "undo redo | fontfamily fontsize | bold italic underline | alignleft aligncenter alignright | bullist numlist | removeformat",
                fontsize_formats: "10pt 12pt 14pt 16pt 18pt 20pt 24pt 30pt 36pt 48pt",
                content_style:
                  "body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; font-size: 14pt; }",
              }}
            />
          </div>
        </div>

        <div style={styles.modalActions}>
          <button style={styles.btnGhost} onClick={onClose} disabled={saving}>
            Cancel
          </button>
          <button style={styles.btnPrimary} onClick={saveAll} disabled={saving}>
            {saving ? "Saving..." : "Save GA Drawing"}
          </button>
        </div>

        <div style={styles.modalNote}>After saving, Export will open automatically.</div>
      </div>
    </div>
  );
}

/** ✅ Manual points modal that writes to report_path_points (lat,lon per line) */
function ManualPointsModal({
  reportId,
  onClose,
  onSave,
}: {
  reportId: string;
  onClose: () => void;
  onSave: (text: string) => void | Promise<void>;
}) {
  const [text, setText] = useState("");
  const [saving, setSaving] = useState(false);

  const submit = async () => {
    setSaving(true);
    try {
      await onSave(text);
    } finally {
      setSaving(false);
    }
  };

  return (
    <div style={styles.modalOverlay} onMouseDown={onClose}>
      <div
        style={{ ...styles.modalCard, width: "min(720px, 96vw)" }}
        onMouseDown={(e) => e.stopPropagation()}
        role="dialog"
        aria-modal="true"
        aria-label="Manual points"
      >
        <div style={styles.modalTitle}>Add Points (Manual)</div>

        <div style={styles.modalHint}>
          Enter points as <b>lat,lon</b> (one per line). Example:
          <div style={{ marginTop: 6, fontFamily: "ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace" }}>
            12.9716,77.5946<br />
            12.9719,77.5952<br />
            12.9723,77.5960
          </div>
        </div>

        <textarea
          value={text}
          onChange={(e) => setText(e.target.value)}
          placeholder="lat,lon (one per line)"
          style={{
            width: "100%",
            minHeight: 180,
            borderRadius: 14,
            border: "1px solid #D0D5DD",
            padding: "12px 12px",
            fontSize: 14,
            fontWeight: 700,
            outline: "none",
            resize: "vertical",
          }}
        />

        <div style={styles.modalActions}>
          <button style={styles.btnGhost} onClick={onClose} disabled={saving}>
            Cancel
          </button>
          <button style={styles.btnPrimary} onClick={submit} disabled={saving}>
            {saving ? "Saving..." : "Save Points"}
          </button>
        </div>

        <div style={styles.modalNote}>
          Saved to <b>report_path_points</b> for report: <span style={{ fontWeight: 900 }}>{reportId.slice(0, 8)}...</span>
        </div>
      </div>
    </div>
  );
}

const locStyles: Record<string, React.CSSProperties> = {
  circle: { width: 12, height: 12, borderRadius: 999, border: "2px solid #111", background: "#fff" },
  dots: {
    width: 2,
    height: 18,
    borderRadius: 2,
    background: "repeating-linear-gradient(to bottom, #111 0 3px, transparent 3px 7px)",
  },
  pin: { fontSize: 18, lineHeight: "18px", color: "#B42318" },
};

// ✅ keep your existing styles object EXACTLY as-is below (unchanged)
const styles: Record<string, React.CSSProperties> = {
  containerFluid: { background: "#F7F8FA", minHeight: "100vh", padding: "18px 18px" },
  pageInner: {
    width: "100%",
    margin: 0,
    display: "grid",
    gap: 14,
    fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif",
  },

  headerCard: {
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 18,
    padding: 16,
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    gap: 14,
    flexWrap: "wrap",
  },
  backLink: { textDecoration: "none", color: "#344054", fontWeight: 800, fontSize: 13 },
  title: { fontSize: 22, fontWeight: 900, color: "#101828", lineHeight: 1.2 },
  metaRow: { display: "flex", gap: 8, flexWrap: "wrap", marginTop: 2 },
  pill: {
    fontSize: 12,
    fontWeight: 800,
    color: "#475467",
    background: "#F2F4F7",
    border: "1px solid #EAECF0",
    borderRadius: 999,
    padding: "6px 10px",
  },

  actions: { display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" },
  btnPrimary: {
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid #111",
    background: "#111",
    color: "#fff",
    cursor: "pointer",
    fontWeight: 900,
    fontSize: 13,
    height: 40,
    whiteSpace: "nowrap",
  },
  btnGhost: {
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid #EAECF0",
    background: "#fff",
    cursor: "pointer",
    fontWeight: 900,
    fontSize: 13,
    color: "#344054",
    height: 40,
    whiteSpace: "nowrap",
  },

  select: {
    height: 40,
    borderRadius: 12,
    border: "1px solid #EAECF0",
    padding: "0 10px",
    fontWeight: 900,
    color: "#101828",
    background: "#fff",
    minWidth: 180,
  },

  card: {
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 18,
    padding: 14,
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
  },
  cardHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "baseline",
    gap: 10,
    marginBottom: 10,
  },
  cardTitle: { fontSize: 14, fontWeight: 900, color: "#101828" },
  cardHint: { fontSize: 12, fontWeight: 800, color: "#667085" },

  controlsRow: { display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" },
  input: {
    flex: 1,
    minWidth: 220,
    padding: "12px 14px",
    borderRadius: 12,
    border: "1px solid #EAECF0",
    outline: "none",
    fontSize: 14,
    background: "#fff",
  },
  inputClearBtn: {
    position: "absolute",
    right: 8,
    top: "50%",
    transform: "translateY(-50%)",
    width: 28,
    height: 28,
    borderRadius: 10,
    border: "1px solid #EAECF0",
    background: "#fff",
    cursor: "pointer",
    fontWeight: 900,
    color: "#667085",
    lineHeight: "26px",
  },

  stateCard: {
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 18,
    padding: 18,
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
  },

  tableCard: {
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 18,
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
    overflow: "hidden",
  },
  tableWrapNoScroll: { width: "100%", overflowX: "hidden" },
  table: { width: "100%", borderCollapse: "separate", borderSpacing: 0, tableLayout: "fixed" },
  th: {
    textAlign: "left",
    fontSize: 12,
    letterSpacing: 0.2,
    fontWeight: 900,
    color: "#475467",
    background: "#F9FAFB",
    borderBottom: "1px solid #EAECF0",
    padding: "12px 12px",
    position: "sticky",
    top: 0,
    zIndex: 1,
    whiteSpace: "nowrap",
  },
  td: {
    padding: "12px 12px",
    borderBottom: "1px solid #F2F4F7",
    verticalAlign: "top",
    fontSize: 13,
    color: "#101828",
    overflow: "hidden",
  },
  catTitle: { fontWeight: 900, fontSize: 13, color: "#101828" },
  descCell: { color: "#475467", lineHeight: 1.45, fontWeight: 700, whiteSpace: "normal", wordBreak: "break-word" },
  subtle: { fontSize: 12, fontWeight: 800, color: "#667085" },
  mutedWrap: { fontSize: 12, fontWeight: 800, color: "#667085", whiteSpace: "normal", wordBreak: "break-word" },
  codePillWrap: {
    display: "inline-block",
    fontSize: 12,
    fontWeight: 900,
    color: "#344054",
    background: "#F2F4F7",
    border: "1px solid #EAECF0",
    borderRadius: 999,
    padding: "6px 10px",
    whiteSpace: "nowrap",
  },
  btnOpen: {
    textDecoration: "none",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    padding: "9px 12px",
    borderRadius: 12,
    border: "1px solid #B2DDFF",
    background: "#EFF8FF",
    color: "#175CD3",
    fontWeight: 900,
    fontSize: 12,
    height: 36,
    whiteSpace: "nowrap",
  },

  // menus
  menu: {
    position: "absolute",
    top: "calc(100% + 8px)",
    right: 0,
    width: 220,
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 14,
    boxShadow: "0 12px 32px rgba(16,24,40,0.12)",
    padding: 6,
    zIndex: 50,
  },
  menuItem: {
    width: "100%",
    textAlign: "left",
    padding: "10px 10px",
    borderRadius: 12,
    border: "none",
    background: "transparent",
    cursor: "pointer",
    fontWeight: 900,
    fontSize: 13,
    color: "#101828",
  },

  // modal
  modalOverlay: {
    position: "fixed",
    inset: 0,
    background: "rgba(16,24,40,0.45)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: 14,
    zIndex: 9999,
  },
  modalCard: {
    width: "min(520px, 96vw)",
    background: "#fff",
    borderRadius: 18,
    border: "1px solid #EAECF0",
    boxShadow: "0 20px 60px rgba(16,24,40,0.25)",
    padding: 16,
    display: "grid",
    gap: 10,
  },
  modalTitle: { fontSize: 16, fontWeight: 950, color: "#101828" },
  modalHint: { fontSize: 12, fontWeight: 800, color: "#667085", lineHeight: 1.35 },
  modalActions: { display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 6 },
  modalNote: { fontSize: 12, fontWeight: 750, color: "#667085" },

  exportGrid: { display: "grid", gridTemplateColumns: "160px 1fr", gap: 10, alignItems: "start", marginTop: 6 },
  routeLabel: { fontSize: 12, fontWeight: 900, color: "#475467" },

  radioRow: {
    display: "flex",
    gap: 10,
    alignItems: "flex-start",
    padding: "10px 10px",
    borderRadius: 14,
    border: "1px solid #EAECF0",
    background: "#fff",
    cursor: "pointer",
  },
  radioTitle: { fontWeight: 950, color: "#101828", fontSize: 13 },
  radioSub: { fontWeight: 800, color: "#667085", fontSize: 12, marginTop: 2 },

  progressBarOuter: {
    height: 10,
    borderRadius: 999,
    background: "#F2F4F7",
    border: "1px solid #EAECF0",
    overflow: "hidden",
    marginTop: 6,
  },
  progressBarInner: {
    height: "100%",
    width: "65%",
    borderRadius: 999,
    background: "#111",
    animation: "pulse 1.1s ease-in-out infinite",
  },
};
