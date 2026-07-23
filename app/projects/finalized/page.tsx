"use client";

// Central list of every project that has an uploaded finalized report, with
// a download for each stored .docx.
import React, { useCallback, useEffect, useState } from "react";
import { useRouter } from "next/navigation";
import { toast } from "../../../components/Toast";

type FinalFile = { id: string; fileName: string; size: number; createdAt: string };
type FinalProject = { projectId: string; projectName: string; files: FinalFile[] };

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

export default function FinishedProjectsPage() {
  const router = useRouter();
  const [projects, setProjects] = useState<FinalProject[]>([]);
  const [loading, setLoading] = useState(true);

  const load = useCallback(async () => {
    try {
      const res = await fetch(`/api/projects/finalized`, { headers: authHeaders(), credentials: "include" });
      if (res.status === 401) {
        router.replace("/login");
        return;
      }
      const data = await res.json().catch(() => ({}));
      setProjects(Array.isArray(data?.projects) ? data.projects : []);
    } catch {
      /* ignore */
    } finally {
      setLoading(false);
    }
  }, [router]);

  useEffect(() => {
    load();
  }, [load]);

  const download = async (projectId: string, f: FinalFile) => {
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

  const totalFiles = projects.reduce((n, p) => n + p.files.length, 0);

  return (
    <div style={S.page}>
      <div style={S.bar}>
        <button style={S.back} onClick={() => router.push("/projects")}>
          ← Projects
        </button>
        <div style={S.title}>📎 Finished projects{projects.length ? ` (${projects.length})` : ""}</div>
        <div style={{ width: 110 }} />
      </div>

      <div style={S.wrap}>
        <p style={S.lead}>
          Projects that have a finalized Word report uploaded. Download any of them here, or open a
          project to add / replace its finalized file.
        </p>

        {loading ? (
          <div style={S.empty}>Loading…</div>
        ) : !projects.length ? (
          <div style={S.emptyCard}>
            No finished projects yet. Open a project → <b>📎 Finalized</b> → upload its final Word
            report, and it will appear here.
          </div>
        ) : (
          <>
            <div style={S.count}>
              {projects.length} project{projects.length === 1 ? "" : "s"} · {totalFiles} file
              {totalFiles === 1 ? "" : "s"}
            </div>
            {projects.map((p) => (
              <div key={p.projectId} style={S.card}>
                <div style={S.head}>
                  <div style={S.pname}>📁 {p.projectName}</div>
                  <button
                    style={S.open}
                    onClick={() => router.push(`/projects/${encodeURIComponent(p.projectId)}/finalized`)}
                    title="Open this project's finalized manager"
                  >
                    Manage ▸
                  </button>
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 8, marginTop: 10 }}>
                  {p.files.map((f) => (
                    <div key={f.id} style={S.row}>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={S.fname}>📄 {f.fileName}</div>
                        <div style={S.meta}>
                          {fmtSize(f.size)} · {fmtDate(f.createdAt)}
                        </div>
                      </div>
                      <button style={S.dl} onClick={() => download(p.projectId, f)}>
                        ⬇ Download
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </>
        )}
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
  wrap: { maxWidth: 820, margin: "0 auto", padding: 20 },
  lead: { fontSize: 13.5, lineHeight: 1.5, color: "#667085", margin: "0 0 14px" },
  count: { fontSize: 12.5, fontWeight: 800, color: "#475467", margin: "0 0 12px" },
  empty: { color: "#98A2B3", fontSize: 14, padding: "24px 0", textAlign: "center" },
  emptyCard: {
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 16,
    padding: 24,
    color: "#667085",
    fontSize: 14,
    lineHeight: 1.6,
    textAlign: "center",
  },
  card: {
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 16,
    padding: 16,
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
    marginBottom: 12,
  },
  head: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 },
  pname: { fontSize: 15, fontWeight: 900, color: "#101828", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  open: {
    padding: "6px 12px",
    borderRadius: 9,
    border: "1px solid #D0D5DD",
    background: "#fff",
    color: "#344054",
    fontWeight: 800,
    cursor: "pointer",
    fontSize: 12.5,
    whiteSpace: "nowrap",
  },
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
};
