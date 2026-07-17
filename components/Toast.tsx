"use client";

// App-wide toast notifications, a drop-in replacement for alert(). `toast()` is
// a plain function (not a hook) so it can be called from anywhere — event
// handlers AND module-level helpers. It dispatches a window event that the
// single <ToastHost/> (mounted in the root layout) renders.
import React, { useEffect, useState } from "react";

export type ToastType = "success" | "error" | "warning" | "info";
type ToastItem = { id: number; message: string; type: ToastType };

// When the caller doesn't pass a type, pick one from the wording so a blanket
// alert()->toast() swap still colours failures red, confirmations green and
// "please select…" prompts amber.
function classify(message: string): ToastType {
  const m = message.toLowerCase();
  if (/\b(fail|failed|error|invalid|cannot|can'?t|couldn'?t|unable|denied|not found|no data|went wrong)\b/.test(m))
    return "error";
  if (/\b(success|saved|created|updated|deleted|imported|exported|copied|moved|combined|done|complete|completed|added)\b/.test(m))
    return "success";
  if (/\b(select|pick|required|at least|choose|enter|provide)\b/.test(m)) return "warning";
  return "info";
}

let seq = 0;
const EVENT = "app-toast";

export function toast(message: string, type?: ToastType): void {
  if (typeof window === "undefined") return;
  const msg = String(message ?? "");
  if (!msg.trim()) return;
  window.dispatchEvent(
    new CustomEvent(EVENT, { detail: { id: ++seq, message: msg, type: type || classify(msg) } })
  );
}

const COLORS: Record<ToastType, { bg: string; icon: string }> = {
  success: { bg: "#16a34a", icon: "✓" },
  error: { bg: "#dc2626", icon: "✕" },
  warning: { bg: "#d97706", icon: "!" },
  info: { bg: "#2563eb", icon: "i" },
};

export default function ToastHost() {
  const [items, setItems] = useState<ToastItem[]>([]);

  useEffect(() => {
    const onToast = (e: Event) => {
      const d = (e as CustomEvent).detail as ToastItem;
      setItems((prev) => [...prev, d]);
      const ttl = d.type === "error" ? 6000 : 4000;
      setTimeout(() => setItems((prev) => prev.filter((t) => t.id !== d.id)), ttl);
    };
    window.addEventListener(EVENT, onToast);
    return () => window.removeEventListener(EVENT, onToast);
  }, []);

  const dismiss = (id: number) => setItems((prev) => prev.filter((t) => t.id !== id));

  return (
    <div style={wrap} aria-live="polite">
      <style>{`@keyframes toastIn{from{opacity:0;transform:translateX(24px)}to{opacity:1;transform:none}}`}</style>
      {items.map((t) => {
        const c = COLORS[t.type];
        return (
          <div
            key={t.id}
            style={{ ...card, background: c.bg }}
            onClick={() => dismiss(t.id)}
            role="alert"
            title="Click to dismiss"
          >
            <span style={iconStyle}>{c.icon}</span>
            <span style={msgStyle}>{t.message}</span>
            <button
              style={closeBtn}
              onClick={(e) => {
                e.stopPropagation();
                dismiss(t.id);
              }}
              aria-label="Dismiss"
            >
              ×
            </button>
          </div>
        );
      })}
    </div>
  );
}

const wrap: React.CSSProperties = {
  position: "fixed",
  top: 16,
  right: 16,
  zIndex: 100000,
  display: "flex",
  flexDirection: "column",
  gap: 10,
  maxWidth: "min(92vw, 400px)",
  pointerEvents: "none",
};
const card: React.CSSProperties = {
  pointerEvents: "auto",
  display: "flex",
  alignItems: "flex-start",
  gap: 10,
  color: "#fff",
  padding: "12px 14px",
  borderRadius: 12,
  boxShadow: "0 10px 30px rgba(0,0,0,0.25)",
  fontSize: 14,
  fontWeight: 600,
  lineHeight: 1.4,
  cursor: "pointer",
  animation: "toastIn .2s ease",
  whiteSpace: "pre-line",
  wordBreak: "break-word",
};
const iconStyle: React.CSSProperties = {
  flex: "0 0 auto",
  width: 20,
  height: 20,
  borderRadius: 999,
  background: "rgba(255,255,255,0.25)",
  display: "inline-flex",
  alignItems: "center",
  justifyContent: "center",
  fontWeight: 900,
  fontSize: 13,
  marginTop: 1,
};
const msgStyle: React.CSSProperties = { flex: 1, minWidth: 0 };
const closeBtn: React.CSSProperties = {
  flex: "0 0 auto",
  background: "transparent",
  border: "none",
  color: "#fff",
  fontSize: 18,
  lineHeight: 1,
  cursor: "pointer",
  opacity: 0.85,
  padding: 0,
  marginLeft: 4,
};
