"use client";

import { useState } from "react";
import { useRouter } from "next/navigation";

export default function LoginPage() {
  const router = useRouter();

  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState<string | null>(null);

  const onLogin = async () => {
    setErr(null);
    setLoading(true);

    try {
      const res = await fetch("/api/auth/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        credentials: "include",
        body: JSON.stringify({
          email: email.trim(),
          password,
        }),
      });

      const data = await res.json().catch(() => ({}));

      if (!res.ok) {
        setErr(data?.error || "Login failed");
        return;
      }

      if (data?.token && typeof window !== "undefined") {
        localStorage.setItem("auth_token", data.token);
      }

      router.push("/projects");
    } catch (error) {
      console.error("[login] request failed:", error);
      setErr("Login failed");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={styles.page}>
      <div style={styles.card}>
        <div style={{ display: "grid", gap: 6, textAlign: "center" }}>
         
          <h2 style={styles.title}>Sign in</h2>
          <div style={styles.subtitle}>Use your email and password to continue</div>
        </div>

        <div style={styles.form}>
          <label style={styles.label}>Email</label>
          <input
            placeholder="name@company.com"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            style={styles.input}
            autoComplete="email"
            inputMode="email"
          />

          <label style={styles.label}>Password</label>
          <input
            placeholder="••••••••"
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            style={styles.input}
            autoComplete="current-password"
            onKeyDown={(e) => {
              if (e.key === "Enter") onLogin();
            }}
          />

          {err ? <div style={styles.error}>{err}</div> : null}

          <button onClick={onLogin} disabled={loading} style={styles.btn}>
            {loading ? "Signing in..." : "Login"}
          </button>

          <div style={styles.footerHint}>
            Trouble signing in? Check your credentials.
          </div>
        </div>
      </div>
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  page: {
    minHeight: "100vh",
    display: "grid",
    placeItems: "center",
    padding: 20,
    background: "#F7F8FA",
    fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif",
  },
  card: {
    width: "100%",
    maxWidth: 420,
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 18,
    padding: 18,
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
  },
  brand: {
    fontWeight: 900,
    fontSize: 13,
    color: "#475467",
    background: "#F2F4F7",
    border: "1px solid #EAECF0",
    borderRadius: 999,
    padding: "6px 10px",
    display: "inline-block",
    margin: "0 auto",
  },
  title: {
    margin: 0,
    fontSize: 22,
    fontWeight: 900,
    color: "#101828",
    lineHeight: 1.2,
  },
  subtitle: {
    fontSize: 13,
    fontWeight: 700,
    color: "#667085",
  },
  form: {
    display: "grid",
    gap: 10,
    marginTop: 14,
  },
  label: {
    fontSize: 12,
    fontWeight: 900,
    color: "#344054",
    marginTop: 4,
  },
  input: {
    padding: "12px 14px",
    border: "1px solid #EAECF0",
    borderRadius: 12,
    outline: "none",
    fontSize: 14,
    background: "#fff",
  },
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
    marginTop: 6,
    padding: "12px 14px",
    borderRadius: 12,
    border: "1px solid #111",
    background: "#111",
    color: "#fff",
    fontWeight: 900,
    fontSize: 14,
    cursor: "pointer",
  },
  footerHint: {
    marginTop: 10,
    fontSize: 12,
    fontWeight: 700,
    color: "#667085",
    textAlign: "center",
  },
};
