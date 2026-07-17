"use client";

// Public landing for a password-protected animated report.
// Client types the password -> we mint a short share session -> hand off to
// the existing route map in "share mode". No app login involved.
import React, { useEffect, useMemo, useState } from "react";
import { useParams, useRouter } from "next/navigation";

export default function SharePage() {
  const params = useParams();
  const router = useRouter();
  const token = useMemo(() => {
    const t = (params as any)?.token;
    return Array.isArray(t) ? t[0] : String(t || "");
  }, [params]);

  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [checking, setChecking] = useState(true);

  const goToMap = (projectId: string) => {
    router.replace(
      `/projects/${encodeURIComponent(projectId)}/route3d?share=${encodeURIComponent(token)}`
    );
  };

  // If this browser already unlocked the link (within the session window),
  // skip the password box entirely.
  useEffect(() => {
    if (!token) {
      setChecking(false);
      return;
    }
    try {
      const sess = sessionStorage.getItem(`share_sess_${token}`);
      const proj = sessionStorage.getItem(`share_proj_${token}`);
      if (sess && proj) {
        goToMap(proj);
        return;
      }
    } catch {
      /* ignore */
    }
    setChecking(false);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [token]);

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (busy) return;
    setBusy(true);
    setError(null);
    try {
      const res = await fetch(`/api/share/${encodeURIComponent(token)}/auth`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ password }),
      });
      const data = await res.json().catch(() => ({} as any));
      if (!res.ok || !data?.ok) {
        throw new Error(data?.error || "Incorrect password.");
      }
      try {
        sessionStorage.setItem(`share_sess_${token}`, String(data.session || ""));
        sessionStorage.setItem(`share_proj_${token}`, String(data.projectId || ""));
      } catch {
        /* ignore */
      }
      goToMap(String(data.projectId || ""));
    } catch (err: any) {
      setError(err?.message || "Incorrect password.");
      setBusy(false);
    }
  };

  if (checking) {
    return (
      <div style={styles.page}>
        <div style={styles.card}>
          <div style={styles.spinner}>Loading…</div>
        </div>
      </div>
    );
  }

  return (
    <div style={styles.page}>
      <form style={styles.card} onSubmit={submit}>
        <div style={styles.logo}>🛰️</div>
        <div style={styles.title}>Protected report</div>
        <div style={styles.subtitle}>
          This animated survey report is password protected. Enter the password shared with you to
          view it.
        </div>

        <label style={styles.label} htmlFor="share-pw">
          Password
        </label>
        <input
          id="share-pw"
          type="password"
          autoFocus
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          placeholder="Enter password"
          style={styles.input}
          autoComplete="off"
        />

        {error ? <div style={styles.error}>{error}</div> : null}

        <button type="submit" disabled={busy || !password} style={styles.button}>
          {busy ? "Checking…" : "View report"}
        </button>

        <div style={styles.footer}>Secured link · do not share the password publicly</div>
      </form>
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  page: {
    minHeight: "100vh",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    background: "linear-gradient(160deg, #0b1220 0%, #0f2436 60%, #0b1220 100%)",
    padding: 20,
  },
  card: {
    width: "100%",
    maxWidth: 380,
    background: "#0f172a",
    border: "1px solid #1e293b",
    borderRadius: 18,
    padding: "30px 26px",
    color: "#e2e8f0",
    boxShadow: "0 20px 60px rgba(0,0,0,0.45)",
    display: "flex",
    flexDirection: "column",
  },
  logo: { fontSize: 40, textAlign: "center", marginBottom: 6 },
  title: { fontSize: 22, fontWeight: 900, textAlign: "center", color: "#fff" },
  subtitle: {
    fontSize: 13.5,
    lineHeight: 1.5,
    textAlign: "center",
    color: "#94a3b8",
    margin: "10px 0 22px",
  },
  label: { fontSize: 12, fontWeight: 800, color: "#cbd5e1", marginBottom: 6, letterSpacing: 0.3 },
  input: {
    padding: "12px 14px",
    borderRadius: 12,
    border: "1px solid #334155",
    background: "#0b1220",
    color: "#fff",
    fontSize: 15,
    outline: "none",
    marginBottom: 14,
  },
  error: {
    background: "#3f1d1d",
    border: "1px solid #7f1d1d",
    color: "#fecaca",
    borderRadius: 10,
    padding: "9px 12px",
    fontSize: 13,
    marginBottom: 14,
  },
  button: {
    padding: "12px 16px",
    borderRadius: 12,
    border: "none",
    background: "#2563eb",
    color: "#fff",
    fontWeight: 900,
    fontSize: 15,
    cursor: "pointer",
  },
  footer: { fontSize: 11.5, color: "#64748b", textAlign: "center", marginTop: 18 },
  spinner: { textAlign: "center", padding: "30px 0", color: "#94a3b8", fontWeight: 700 },
};
