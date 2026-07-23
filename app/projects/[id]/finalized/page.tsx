"use client";

// Finalized reports for a project: upload your finished .docx to S3 and
// re-download it anytime — no regenerating. Same Word format as the export.
import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { useParams, useRouter } from "next/navigation";
import { toast } from "../../../../components/Toast";
import { confirmDialog } from "../../../../components/ConfirmDialog";

type FinalFile = { id: string; fileName: string; size: number; createdAt: string };

function authHeaders(): Record<string, string> {
  const t = typeof window !== "undefined" ? localStorage.getItem("auth_token") : null;
  return t ? { Authorization: `Bearer ${t}` } : {};
}
function fmtSize(n: number) {
  if (!n) return "";
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(0)} KB`;
  return `${(n / 1024 / 1024).toFixed(1)} MB`;
}
function fmtDate(s: string) {
  try {
    return new Date(s).toLocaleString();
  } catch {
    return s;
  }
}

export default function FinalizedPage() {
  const params = useParams();
  const router = useRouter();
  const projectId = useMemo(() => {
    const id = (params as any)?.id;
    return Array.isArray(id) ? id[0] : String(id || "");
  }, [params]);

  const [projectName, setProjectName] = useState("");
  const [files, setFiles] = useState<FinalFile[]>([]);
  const [loading, setLoading] = useState(true);
  const [uploading, setUploading] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  const load = useCallback(async () => {
    try {
      const res = await fetch(`/api/projects/${encodeURIComponent(projectId)}/finalized`, {
        headers: authHeaders(),
        credentials: "include",
      });
      const data = await res.json().catch(() => ({}));
      if (res.status === 401) {
        router.replace("/login");
        return;
      }
      setFiles(Array.isArray(data?.files) ? data.files : []);
    } catch {
      /* ignore */
    } finally {
      setLoading(false);
    }
  }, [projectId, router]);

  useEffect(() => {
    if (!projectId) return;
    // Best-effort project name for the header.
    fetch(`/api/projects/${encodeURIComponent(projectId)}`, { headers: authHeaders(), credentials: "include" })
      .then((r) => r.json())
      .then((d) => setProjectName(String(d?.name || d?.project?.name || "")))
      .catch(() => {});
    load();
  }, [projectId, load]);

  const onPick = () => fileRef.current?.click();

  const onFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = ""; // allow re-uploading the same file
    if (!file) return;
    if (!/\.docx$/i.test(file.name)) {
      toast("Please choose a Word (.docx) file — the same format as the export.", "warning");
      return;
    }
    setUploading(true);
    try {
      const fd = new FormData();
      fd.append("file", file);
      const res = await fetch(`/api/projects/${encodeURIComponent(projectId)}/finalized`, {
        method: "POST",
        headers: authHeaders(),
        credentials: "include",
        body: fd,
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok || !data?.ok) throw new Error(data?.error || "Upload failed");
      toast(`Uploaded "${data.file.fileName}"`, "success");
      await load();
    } catch (err: any) {
      toast(err?.message || "Upload failed", "error");
    } finally {
      setUploading(false);
    }
  };

  const download = async (f: FinalFile) => {
    try {
      const res = await fetch(
        `/api/projects/${encodeURIComponent(projectId)}/finalized/download?fileId=${encodeURIComponent(f.id)}`,
        { headers: authHeaders(), credentials: "include" }
      );
      const data = await res.json().catch(() => ({}));
      if (!res.ok || !data?.url) throw new Error(data?.error || "Could not prepare download");
      const a = document.createElement("a");
      a.href = data.url;
      a.rel = "noopener";
      document.body.appendChild(a);
      a.click();
      a.remove();
    } catch (err: any) {
      toast(err?.message || "Download failed", "error");
    }
  };

  const remove = async (f: FinalFile) => {
    const ok = await confirmDialog(`Delete "${f.fileName}"? This removes it from storage.`, {
      title: "Delete finalized report?",
      confirmText: "Delete",
      danger: true,
    });
    if (!ok) return;
    try {
      const res = await fetch(
        `/api/projects/${encodeURIComponent(projectId)}/finalized?fileId=${encodeURIComponent(f.id)}`,
        { method: "DELETE", headers: authHeaders(), credentials: "include" }
      );
      const data = await res.json().catch(() => ({}));
      if (!res.ok || !data?.ok) throw new Error(data?.error || "Delete failed");
      toast("Deleted", "success");
      setFiles((prev) => prev.filter((x) => x.id !== f.id));
    } catch (err: any) {
      toast(err?.message || "Delete failed", "error");
    }
  };

  return (
    <div style={S.page}>
      <div style={S.bar}>
        <button style={S.back} onClick={() => router.push(`/projects/${encodeURIComponent(projectId)}`)}>
          ← Back to project
        </button>
        <div style={S.title}>📎 Finalized report{projectName ? ` — ${projectName}` : ""}</div>
        <div style={{ width: 120 }} />
      </div>

      <div style={S.wrap}>
        <div style={S.card}>
          <div style={S.h2}>Upload finalized Word report</div>
          <p style={S.p}>
            Store your finished <b>.docx</b> (same format as the export) here. It&apos;s kept in
            secure storage and you can download it again anytime — no need to regenerate.
          </p>
          <input
            ref={fileRef}
            type="file"
            accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            style={{ display: "none" }}
            onChange={onFile}
          />
          <button style={S.upload} onClick={onPick} disabled={uploading}>
            {uploading ? "Uploading…" : "⬆ Choose .docx to upload"}
          </button>
        </div>

        <div style={S.card}>
          <div style={S.h2}>Stored files {files.length ? `(${files.length})` : ""}</div>
          {loading ? (
            <div style={S.empty}>Loading…</div>
          ) : !files.length ? (
            <div style={S.empty}>No finalized reports uploaded yet.</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {files.map((f) => (
                <div key={f.id} style={S.row}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={S.fname}>📄 {f.fileName}</div>
                    <div style={S.meta}>
                      {fmtSize(f.size)} · {fmtDate(f.createdAt)}
                    </div>
                  </div>
                  <button style={S.dl} onClick={() => download(f)}>
                    ⬇ Download
                  </button>
                  <button style={S.del} onClick={() => remove(f)} title="Delete">
                    ✕
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

const S: Record<string, React.CSSProperties> = {
  page: { minHeight: "100vh", background: "#F7F8FA", fontFamily: "system-ui, Segoe UI, Arial" },
  bar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "12px 18px",
    background: "#0f172a",
    color: "#fff",
  },
  back: {
    padding: "8px 12px",
    borderRadius: 10,
    border: "1px solid #334155",
    background: "#1e293b",
    color: "#fff",
    fontWeight: 800,
    cursor: "pointer",
  },
  title: { fontSize: 16, fontWeight: 900 },
  wrap: { maxWidth: 760, margin: "0 auto", padding: 20, display: "flex", flexDirection: "column", gap: 16 },
  card: {
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 16,
    padding: 20,
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
  },
  h2: { fontSize: 16, fontWeight: 900, color: "#101828", marginBottom: 8 },
  p: { fontSize: 13.5, lineHeight: 1.5, color: "#667085", margin: "0 0 16px" },
  upload: {
    padding: "12px 18px",
    borderRadius: 12,
    border: "none",
    background: "#2563eb",
    color: "#fff",
    fontWeight: 900,
    fontSize: 14,
    cursor: "pointer",
  },
  empty: { color: "#98A2B3", fontSize: 13.5, padding: "8px 0" },
  row: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    padding: "10px 12px",
    border: "1px solid #EAECF0",
    borderRadius: 12,
    background: "#FCFCFD",
  },
  fname: { fontWeight: 800, color: "#101828", fontSize: 14, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  meta: { fontSize: 12, color: "#98A2B3", marginTop: 2 },
  dl: {
    padding: "8px 14px",
    borderRadius: 10,
    border: "none",
    background: "#16a34a",
    color: "#fff",
    fontWeight: 800,
    cursor: "pointer",
    whiteSpace: "nowrap",
  },
  del: {
    padding: "8px 11px",
    borderRadius: 10,
    border: "1px solid #FDA29B",
    background: "#fff",
    color: "#B42318",
    fontWeight: 900,
    cursor: "pointer",
  },
};
