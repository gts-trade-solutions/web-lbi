"use client";

// Download the Tracker Android app. Behind AuthGate (not a public path), and
// the file itself is only served by /api/app/download to a logged-in user.
import { useEffect, useState } from "react";
import { useRouter } from "next/navigation";

type ApkInfo = { url: string; fileName: string; size: number; updatedAt: string };

function authHeaders(): Record<string, string> {
  const t = typeof window !== "undefined" ? localStorage.getItem("auth_token") : null;
  return t ? { Authorization: `Bearer ${t}` } : {};
}

// Each call returns a fresh ticket (valid 2 minutes), so ask right before downloading.
async function requestTicket(): Promise<ApkInfo> {
  const res = await fetch("/api/app/download", {
    method: "POST",
    headers: authHeaders(),
    credentials: "include",
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data?.error || "Could not prepare the download");
  return data as ApkInfo;
}

export default function AppDownloadPage() {
  const router = useRouter();
  const [info, setInfo] = useState<ApkInfo | null>(null);
  const [err, setErr] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    requestTicket()
      .then(setInfo)
      .catch((e) => setErr(e?.message || "Could not load app details"));
  }, []);

  const download = async () => {
    setErr(null);
    setBusy(true);
    try {
      const fresh = await requestTicket();
      setInfo(fresh);
      const a = document.createElement("a");
      a.href = fresh.url;
      a.download = fresh.fileName;
      document.body.appendChild(a);
      a.click();
      a.remove();
    } catch (e: any) {
      setErr(e?.message || "Download failed");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div style={styles.page}>
      <div style={styles.card}>
        <h2 style={styles.title}>Tracker Android app</h2>
        <div style={styles.subtitle}>Only signed-in users can download the app.</div>

        {info ? (
          <div style={styles.meta}>
            <div>
              Size: <b>{(info.size / (1024 * 1024)).toFixed(1)} MB</b>
            </div>
            <div>
              Uploaded: <b>{new Date(info.updatedAt).toLocaleString()}</b>
            </div>
          </div>
        ) : null}

        {err ? <div style={styles.error}>{err}</div> : null}

        <button onClick={download} disabled={busy || (!info && !err)} style={styles.btn}>
          {busy ? "Preparing…" : "Download APK"}
        </button>

        <div style={styles.hint}>
          On the phone, open the downloaded file and allow “Install unknown apps” for your browser if Android
          asks.
        </div>

        <button onClick={() => router.push("/projects")} style={styles.btnGhost}>
          ← Back to projects
        </button>
      </div>
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  page: {
    minHeight: "100vh",
    display: "grid",
    placeItems: "center",
    padding: 16,
    background: "#F7F8FA",
    fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif",
  },
  card: {
    width: "100%",
    maxWidth: 420,
    display: "grid",
    gap: 12,
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 18,
    padding: 18,
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
  },
  title: { margin: 0, fontSize: 22, fontWeight: 900, color: "#101828", lineHeight: 1.2 },
  subtitle: { fontSize: 13, fontWeight: 700, color: "#667085" },
  meta: { display: "grid", gap: 4, fontSize: 13, color: "#344054" },
  error: {
    background: "#FEF3F2",
    border: "1px solid #FECDCA",
    color: "#B42318",
    borderRadius: 12,
    padding: "10px 12px",
    fontWeight: 800,
    fontSize: 13,
  },
  btn: {
    padding: "12px 14px",
    borderRadius: 12,
    border: "1px solid #111",
    background: "#111",
    color: "#fff",
    fontWeight: 900,
    fontSize: 14,
    cursor: "pointer",
  },
  btnGhost: {
    padding: "10px 14px",
    borderRadius: 12,
    border: "1px solid #EAECF0",
    background: "#fff",
    color: "#344054",
    fontWeight: 800,
    fontSize: 13,
    cursor: "pointer",
  },
  hint: { fontSize: 12, fontWeight: 700, color: "#667085" },
};
