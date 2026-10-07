"use client";

// Central list of every project that has an uploaded finalized report, with
// a download for each stored .docx.
import React, { useCallback, useEffect, useState } from "react";
import { useRouter } from "next/navigation";
import { toast } from "../../../components/Toast";

type FinalFile = { id: string; fileName: string; size: number; createdAt: string };
type FinalProject = { projectId: string; projectName: string; files: FinalFile[]; finished?: boolean };

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
  const [movingId, setMovingId] = useState("");
  const [dlId, setDlId] = useState("");

  // Generate and download the project's Word report on the fly (same as the
  // export on the project page), so a finished project can be downloaded here
  // even when no finalized file was uploaded.
  const downloadReport = async (projectId: string, name: string) => {
    try {
      setDlId(projectId);
      toast("Preparing the Word report… this can take a minute for large projects.");
      const res = await fetch(`/api/projects/${encodeURIComponent(projectId)}/export`, {
        headers: authHeaders(),
        credentials: "include",
      });
      if (!res.ok) {
        const d = await res.json().catch(() => ({}));
        throw new Error(d?.error || "Could not generate the report");
      }
      const blob = await res.blob();
      const cd = res.headers.get("content-disposition") || "";
      const m = cd.match(/filename\*?=(?:UTF-8''|")?([^";]+)/i);
      const fileName = m ? decodeURIComponent(m[1].replace(/"/g, "")) : `${name}.docx`;
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 5000);
      toast("Report downloaded.", "success");
    } catch (e: any) {
      toast(e?.message || "Download failed", "error");
    } finally {
      setDlId("");
    }
  };

  const moveBack = async (projectId: string, name: string) => {
    if (!window.confirm(`Move "${name}" back to active projects?`)) return;
    try {
      setMovingId(projectId);
      const res = await fetch(`/api/projects/${encodeURIComponent(projectId)}`, {
        method: "PUT",
        credentials: "include",
        headers: { "Content-Type": "application/json", ...authHeaders() },
        body: JSON.stringify({ status: "active" }),
      });
      if (!res.ok) {
        const d = await res.json().catch(() => ({}));
        throw new Error(d?.error || "Failed to move project back");
      }
      setProjects((prev) => prev.filter((p) => p.projectId !== projectId));
      toast(`Moved "${name}" back to active projects.`, "success");
    } catch (e: any) {
      toast(e?.message || "Failed to move project back", "error");
    } finally {
      setMovingId("");
    }
  };

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
      if (!res.ok) {
        const d = await res.json().catch(() => ({}));
        throw new Error(d?.error || "Could not prepare download");
      }
      const ct = res.headers.get("content-type") || "";
      const a = document.createElement("a");
      if (ct.includes("application/json")) {
        const data = await res.json();
        if (!data?.url) throw new Error("No download URL");
        a.href = data.url;
        a.rel = "noopener";
      } else {
        const blob = await res.blob();
        a.href = URL.createObjectURL(blob);
        a.download = f.fileName;
      }
      document.body.appendChild(a);
      a.click();
      a.remove();
      if (a.href.startsWith("blob:")) setTimeout(() => URL.revokeObjectURL(a.href), 5000);
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
          Projects you&apos;ve marked <b>finished</b> (hidden from the main list), plus any with a
          finalized Word report uploaded. Use <b>↩ Move back to active</b> to return one to the main
          list.
        </p>

        {loading ? (
          <div style={S.empty}>Loading…</div>
        ) : !projects.length ? (
          <div style={S.emptyCard}>
            No finished projects yet. On the Projects page, click <b>✓ Finished</b> on any project to
            move it here.
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
                  <div style={{ display: "flex", gap: 8, flexShrink: 0, flexWrap: "wrap", justifyContent: "flex-end" }}>
                    <button
                      style={{ ...S.dl, opacity: dlId === p.projectId ? 0.6 : 1 }}
                      disabled={dlId === p.projectId}
                      onClick={() => downloadReport(p.projectId, p.projectName)}
                      title="Generate and download this project's Word report"
                    >
                      {dlId === p.projectId ? "Preparing…" : "⬇ Download report"}
                    </button>
                    {p.finished && (
                      <button
                        style={{ ...S.open, borderColor: "#B54708", color: "#B54708", opacity: movingId === p.projectId ? 0.6 : 1 }}
                        disabled={movingId === p.projectId}
                        onClick={() => moveBack(p.projectId, p.projectName)}
                        title="Move this project back to the active projects list"
                      >
                        {movingId === p.projectId ? "Moving…" : "↩ Move back to active"}
                      </button>
                    )}
                    <button
                      style={S.open}
                      onClick={() => router.push(`/projects/${encodeURIComponent(p.projectId)}`)}
                      title="Open this project"
                    >
                      Open ▸
                    </button>
                  </div>
                </div>
                {p.files.length ? (
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
                ) : (
                  <div style={{ ...S.meta, marginTop: 10 }}>
                    Marked finished — no uploaded file. Use <b>⬇ Download report</b> to generate the
                    Word report.
                  </div>
                )}
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
