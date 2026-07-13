"use client";

import { useCallback, useEffect, useState } from "react";
import Link from "next/link";

type ActivityRow = {
  id: string;
  created_at: string;
  user_id: string | null;
  user_email: string | null;
  action: string;
  table_name: string | null;
  entity_id: string | null;
  project_id: string | null;
  row_count: number | null;
  details: string | null;
  ip_address: string | null;
  user_agent: string | null;
};

type Summary = { action: string; n: number };

function authHeaders(): Record<string, string> {
  const token = typeof window !== "undefined" ? localStorage.getItem("auth_token") : null;
  return token ? { Authorization: `Bearer ${token}` } : {};
}

// Colour per action so deletes/logins stand out at a glance.
const ACTION_STYLE: Record<string, { bg: string; fg: string }> = {
  login: { bg: "#EFF8FF", fg: "#175CD3" },
  login_failed: { bg: "#FEF3F2", fg: "#B42318" },
  logout: { bg: "#F2F4F7", fg: "#475467" },
  create: { bg: "#ECFDF3", fg: "#027A48" },
  update: { bg: "#FFFAEB", fg: "#B54708" },
  move_project: { bg: "#FDF4FF", fg: "#9F1AB1" },
  bulk_import: { bg: "#EEF4FF", fg: "#3538CD" },
  delete: { bg: "#FEF3F2", fg: "#B42318" },
  export: { bg: "#F2F4F7", fg: "#475467" },
};

export default function ActivityLogPage() {
  const [rows, setRows] = useState<ActivityRow[]>([]);
  const [summary, setSummary] = useState<Summary[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const [q, setQ] = useState("");
  const [action, setAction] = useState("");
  const [date, setDate] = useState("");

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const sp = new URLSearchParams();
      if (q.trim()) sp.set("q", q.trim());
      if (action) sp.set("action", action);
      if (date) sp.set("date", date);
      sp.set("limit", "500");
      const res = await fetch(`/api/activity?${sp.toString()}`, {
        credentials: "include",
        headers: authHeaders(),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data?.error || "Failed to load activity log");
      setRows(Array.isArray(data.rows) ? data.rows : []);
      setSummary(Array.isArray(data.summary) ? data.summary : []);
    } catch (e: any) {
      setError(e?.message || "Failed to load activity log");
      setRows([]);
    } finally {
      setLoading(false);
    }
  }, [q, action, date]);

  useEffect(() => {
    load();
  }, [load]);

  const fmtTime = (s: string) => {
    try {
      return new Date(s).toLocaleString();
    } catch {
      return s;
    }
  };

  const prettyDetails = (d: string | null) => {
    if (!d) return "";
    try {
      const o = JSON.parse(d);
      return Object.entries(o)
        .map(([k, v]) => `${k}: ${typeof v === "object" ? JSON.stringify(v) : String(v)}`)
        .join("  •  ");
    } catch {
      return d;
    }
  };

  return (
    <div style={{ padding: 24, fontFamily: "system-ui, Segoe UI, Arial", background: "#F7F8FA", minHeight: "100vh" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <Link href="/projects" style={{ ...btnGhost, textDecoration: "none" }}>
          ← Projects
        </Link>
        <h1 style={{ fontSize: 22, fontWeight: 950, color: "#101828", margin: 0 }}>Activity Log</h1>
        <button style={btnGhost} onClick={load} disabled={loading}>
          {loading ? "Loading…" : "↻ Refresh"}
        </button>
      </div>
      <p style={{ color: "#667085", fontWeight: 700, marginTop: 6, fontSize: 13 }}>
        Every login, create, edit, delete, project move and bulk import — with who, when, and from which
        device/IP. Use this to see exactly what happened and when.
      </p>

      {/* Filters */}
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", margin: "12px 0" }}>
        <input
          placeholder="Search email, IP, ids, details…"
          value={q}
          onChange={(e) => setQ(e.target.value)}
          style={{ ...input, minWidth: 260, flex: 1 }}
        />
        <select value={action} onChange={(e) => setAction(e.target.value)} style={input}>
          <option value="">All actions</option>
          <option value="login">login</option>
          <option value="login_failed">login_failed</option>
          <option value="create">create</option>
          <option value="update">update</option>
          <option value="move_project">move_project</option>
          <option value="delete">delete</option>
          <option value="bulk_import">bulk_import</option>
        </select>
        <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={input} />
        {(q || action || date) && (
          <button
            style={btnGhost}
            onClick={() => {
              setQ("");
              setAction("");
              setDate("");
            }}
          >
            Clear
          </button>
        )}
      </div>

      {/* Summary chips */}
      {summary.length > 0 && (
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 12 }}>
          {summary.map((s) => {
            const st = ACTION_STYLE[s.action] || { bg: "#F2F4F7", fg: "#475467" };
            return (
              <span
                key={s.action}
                style={{ background: st.bg, color: st.fg, borderRadius: 999, padding: "4px 12px", fontWeight: 900, fontSize: 12 }}
              >
                {s.action}: {s.n}
              </span>
            );
          })}
        </div>
      )}

      {error ? (
        <div style={{ ...card, color: "#B42318" }}>{error}</div>
      ) : loading ? (
        <div style={card}>Loading…</div>
      ) : !rows.length ? (
        <div style={card}>No activity recorded yet for this filter.</div>
      ) : (
        <div style={{ ...card, padding: 0, overflowX: "auto" }}>
          <table style={{ borderCollapse: "collapse", width: "100%", minWidth: 980 }}>
            <thead>
              <tr>
                {["When", "Who", "Action", "Table", "Rows", "Project", "Details", "IP"].map((h) => (
                  <th key={h} style={th}>
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => {
                const st = ACTION_STYLE[r.action] || { bg: "#F2F4F7", fg: "#475467" };
                return (
                  <tr key={r.id}>
                    <td style={{ ...td, whiteSpace: "nowrap" }}>{fmtTime(r.created_at)}</td>
                    <td style={td}>{r.user_email || r.user_id || "—"}</td>
                    <td style={td}>
                      <span style={{ background: st.bg, color: st.fg, borderRadius: 999, padding: "3px 10px", fontWeight: 900, fontSize: 12 }}>
                        {r.action}
                      </span>
                    </td>
                    <td style={td}>{r.table_name || "—"}</td>
                    <td style={{ ...td, textAlign: "right", fontWeight: 800 }}>{r.row_count ?? "—"}</td>
                    <td style={{ ...td, fontFamily: "monospace", fontSize: 11 }}>
                      {r.project_id ? r.project_id.slice(0, 8) + "…" : "—"}
                    </td>
                    <td style={{ ...td, maxWidth: 360, fontSize: 12, color: "#475467" }}>{prettyDetails(r.details)}</td>
                    <td style={{ ...td, fontSize: 11, color: "#667085" }}>{r.ip_address || "—"}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

const card: React.CSSProperties = {
  background: "#fff",
  border: "1px solid #EAECF0",
  borderRadius: 14,
  padding: 16,
  fontWeight: 700,
  color: "#101828",
};
const input: React.CSSProperties = {
  height: 38,
  borderRadius: 10,
  border: "1px solid #D0D5DD",
  padding: "0 12px",
  fontWeight: 700,
  color: "#101828",
  background: "#fff",
};
const btnGhost: React.CSSProperties = {
  height: 38,
  borderRadius: 10,
  border: "1px solid #D0D5DD",
  padding: "0 14px",
  background: "#fff",
  color: "#344054",
  fontWeight: 900,
  cursor: "pointer",
};
const th: React.CSSProperties = {
  textAlign: "left",
  padding: "10px 12px",
  fontSize: 12,
  fontWeight: 900,
  color: "#475467",
  background: "#F9FAFB",
  borderBottom: "1px solid #EAECF0",
  whiteSpace: "nowrap",
};
const td: React.CSSProperties = {
  padding: "9px 12px",
  fontSize: 13,
  color: "#101828",
  borderBottom: "1px solid #F2F4F7",
  verticalAlign: "top",
};
