"use client";

import { useEffect, useMemo, useState } from "react";
import { supabase } from "../../../lib/supabase";
import type { Report, ReportPhoto, ReportPathPoint } from "@/types/db";

type VehicleMovement = "green" | "yellow" | "red" | null;

const VM_OPTIONS: { value: VehicleMovement; label: string; bg: string; fg: string }[] = [
  { value: null, label: "— Select —", bg: "#F2F4F7", fg: "#475467" },
  { value: "green", label: "No problem", bg: "#12B76A", fg: "#FFFFFF" },
  { value: "yellow", label: "Quite difficulties", bg: "#F79009", fg: "#FFFFFF" },
  { value: "red", label: "Don't go (vehicle movement)", bg: "#F04438", fg: "#FFFFFF" },
];

function vmPill(value: VehicleMovement) {
  const opt = VM_OPTIONS.find((o) => o.value === value) || VM_OPTIONS[0];
  return (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        padding: "6px 10px",
        borderRadius: 999,
        background: opt.bg,
        color: opt.fg,
        fontWeight: 900,
        fontSize: 12,
        lineHeight: 1,
        border: "1px solid rgba(16,24,40,0.08)",
      }}
    >
      {opt.label}
    </span>
  );
}

export default function ReportDetailPage({ params }: { params: { id: string } }) {
  const reportId = params.id;

  const [report, setReport] = useState<Report | null>(null);
  const [photos, setPhotos] = useState<ReportPhoto[]>([]);
  const [pathPoints, setPathPoints] = useState<ReportPathPoint[]>([]);
  const [loading, setLoading] = useState(true);
  const [exporting, setExporting] = useState(false);

  const [resolvedPhotoUrls, setResolvedPhotoUrls] = useState<string[]>([]);
  const [savingVM, setSavingVM] = useState(false);

  useEffect(() => {
    (async () => {
      setLoading(true);

      const rep = await supabase.from("reports").select("*").eq("id", reportId).maybeSingle();
      if (rep.error || !rep.data) {
        setReport(null);
        setLoading(false);
        return;
      }

      setReport(rep.data as Report);

      const ph = await supabase
        .from("report_photos")
        .select("*")
        .eq("report_id", reportId)
        .order("created_at", { ascending: true });
      setPhotos((ph.data || []) as ReportPhoto[]);

      const pts = await supabase
        .from("report_path_points")
        .select("*")
        .eq("report_id", reportId)
        .order("seq", { ascending: true });
      setPathPoints((pts.data || []) as ReportPathPoint[]);

      setLoading(false);
    })();
  }, [reportId]);

  const center = useMemo(() => {
    const lat = (report as any)?.loc_lat;
    const lon = (report as any)?.loc_lon;
    if (typeof lat === "number" && typeof lon === "number") return { lat, lng: lon };

    const p0: any = pathPoints[0];
    const plat = p0?.latitude ?? p0?.lat;
    const plon = p0?.longitude ?? p0?.lon ?? p0?.lng;
    if (typeof plat === "number" && typeof plon === "number") return { lat: plat, lng: plon };

    return { lat: 12.9716, lng: 77.5946 };
  }, [report, pathPoints]);

  const getPhotoUrl = async (p: ReportPhoto) => {
    const anyP: any = p;
    if (anyP.url) return anyP.url as string;
    if (anyP.public_url) return anyP.public_url as string;

    if (anyP.bucket && anyP.path) {
      const { data, error } = await supabase.storage.from(anyP.bucket).createSignedUrl(anyP.path, 60);
      if (!error && data?.signedUrl) return data.signedUrl;
    }
    return "";
  };

  useEffect(() => {
    (async () => {
      const urls: string[] = [];
      for (const p of photos) {
        const u = await getPhotoUrl(p);
        if (u) urls.push(u);
      }
      setResolvedPhotoUrls(urls);
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [photos]);

  const saveVehicleMovement = async (value: VehicleMovement) => {
    if (!report) return;
    try {
      setSavingVM(true);
      const { error } = await supabase.from("reports").update({ vehicle_movement: value }).eq("id", reportId);
      if (error) throw error;

      setReport({ ...(report as any), vehicle_movement: value } as any);
    } catch (e: any) {
      alert("Failed to save vehicle movement: " + (e?.message || String(e)));
    } finally {
      setSavingVM(false);
    }
  };

  const downloadWord = async () => {
    try {
      setExporting(true);
      const res = await fetch(`/api/reports/${encodeURIComponent(reportId)}/docx`, {
        method: "GET",
        credentials: "include",
      });
      const data = await res.json().catch(() => ({} as any));
      if (!res.ok) throw new Error(data?.error || "Failed to generate DOCX");
      if (!data?.base64) throw new Error("Invalid DOCX API response");

      const binary = atob(data.base64);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
      const blob = new Blob([bytes], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      const filename = data.filename || `report-${reportId}.docx`;
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (e: any) {
      alert("Word export error: " + (e?.message || String(e)));
    } finally {
      setExporting(false);
    }
  };

  if (loading) return <div style={styles.page}><div style={styles.state}>Loading...</div></div>;
  if (!report) return <div style={styles.page}><div style={styles.state}>Report not found</div></div>;

  const title = (report as any).category || "Report";
  const created = (report as any).created_at ? new Date((report as any).created_at).toLocaleString() : "—";
  const vmValue = ((report as any).vehicle_movement ?? null) as VehicleMovement;

  return (
    <div style={styles.page}>
      {/* Header Card */}
      <div style={styles.headerCard}>
        <div>
          <div style={styles.title}>{title}</div>
          <div style={styles.subTitle}>Created: {created}</div>
          <div style={styles.smallMeta}>Report ID: {(report as any).id}</div>
        </div>

        <div style={styles.actions}>
          <button style={styles.btnGhost} onClick={() => history.back()}>
            ← Back
          </button>

          {/* <button style={styles.btnPrimary} onClick={downloadWord} disabled={exporting}>
            {exporting ? "Generating..." : "Download DOCX"}
          </button> */}
        </div>
      </div>

      {/* Vehicle Movement (NEW) */}
      {/* <div style={styles.card}>
        <div style={styles.cardHeader}>
          <div style={styles.cardTitle}>Route difficulty / Vehicle movement</div>
          <div style={styles.cardHint}>{vmPill(vmValue)}</div>
        </div>

        <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
          <select
            value={vmValue ?? ""}
            onChange={(e) => saveVehicleMovement((e.target.value || null) as VehicleMovement)}
            disabled={savingVM}
            style={{
              height: 42,
              borderRadius: 12,
              border: "1px solid #EAECF0",
              padding: "0 12px",
              fontWeight: 800,
              color: "#101828",
              background: "#fff",
              minWidth: 260,
            }}
          >
            <option value="">— Select route condition —</option>
            <option value="green">Green — Clear (no issues)</option>
            <option value="yellow">Yellow — Moderate difficulty (use caution)</option>
            <option value="red">Red — Not passable (do not proceed)</option>

          </select>

          <div style={{ fontSize: 12, color: "#667085", fontWeight: 700 }}>
            {savingVM ? "Saving..." : "Saved to DB and will show in Word export."}
          </div>
        </div>
      </div> */}

      {/* Map */}
      <div style={styles.card}>
        <div style={styles.cardHeader}>
          <div style={styles.cardTitle}>Location</div>
          <div style={styles.cardHint}>
            {center.lat.toFixed(5)}, {center.lng.toFixed(5)}
          </div>
        </div>

        <div style={styles.mapWrap}>
          <iframe
            width="100%"
            height="320"
            loading="lazy"
            style={{ border: 0 }}
            referrerPolicy="no-referrer-when-downgrade"
            src={`https://www.google.com/maps?q=${center.lat},${center.lng}&z=16&output=embed`}
          />
        </div>
      </div>

      {/* Photos */}
      <div style={styles.card}>
        <div style={styles.cardHeader}>
          <div style={styles.cardTitle}>Photos</div>
          <div style={styles.cardHint}>{resolvedPhotoUrls.length} image(s)</div>
        </div>

        {resolvedPhotoUrls.length === 0 ? (
          <div style={styles.empty}>No photos available.</div>
        ) : (
          <div style={styles.photoGrid}>
            {resolvedPhotoUrls.map((src, idx) => (
              <a key={src + idx} href={src} target="_blank" rel="noreferrer" style={styles.photoLink}>
                <img src={src} alt="" style={styles.photo} />
              </a>
            ))}
          </div>
        )}
      </div>

      {/* Description */}
      <div style={styles.card}>
        <div style={styles.cardHeader}>
          <div style={styles.cardTitle}>Description</div>
        </div>
        <div style={styles.desc}>{(report as any).description || "—"}</div>
      </div>

      {/* Bottom actions */}
      {/* <div style={styles.bottomRow}>
        <button style={styles.btnGhost} onClick={() => alert("GPX/KML/KMZ next step")}>
          Export GPX/KML/KMZ
        </button>
        <button style={styles.btnGhost} onClick={() => history.back()}>
          Back
        </button>
      </div> */}
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  page: {
    padding: 24,
    fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif",
    background: "#F7F8FA",
    minHeight: "100vh",
    maxWidth: 1200,
    margin: "0 auto",
    display: "grid",
    gap: 14,
  },
  state: {
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 16,
    padding: 18,
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
    fontWeight: 700,
    color: "#101828",
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
  title: { fontSize: 22, fontWeight: 900, color: "#101828", lineHeight: 1.2 },
  subTitle: { marginTop: 6, fontSize: 13, color: "#667085", fontWeight: 700 },
  smallMeta: { marginTop: 6, fontSize: 12, color: "#98A2B3", fontWeight: 700 },
  actions: { display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" },
  btnPrimary: {
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid #111",
    background: "#111",
    color: "#fff",
    cursor: "pointer",
    fontWeight: 800,
    fontSize: 13,
  },
  btnGhost: {
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid #EAECF0",
    background: "#fff",
    cursor: "pointer",
    fontWeight: 800,
    fontSize: 13,
    color: "#344054",
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
  mapWrap: { overflow: "hidden", borderRadius: 14, border: "1px solid #EAECF0" },
  photoGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))",
    gap: 10,
  },
  photoLink: { textDecoration: "none" },
  photo: {
    width: "100%",
    height: 180,
    objectFit: "cover",
    borderRadius: 14,
    border: "1px solid #EAECF0",
    display: "block",
  },
  empty: {
    padding: 14,
    borderRadius: 14,
    border: "1px dashed #D0D5DD",
    color: "#667085",
    fontWeight: 700,
    background: "#FCFCFD",
  },
  desc: {
    whiteSpace: "pre-wrap",
    color: "#475467",
    fontSize: 14,
    lineHeight: 1.5,
    paddingTop: 6,
  },
  bottomRow: {
    display: "flex",
    gap: 10,
    justifyContent: "flex-end",
    flexWrap: "wrap",
  },
};
