"use client";

// App-wide async confirm dialog — an in-page replacement for window.confirm().
// `confirmDialog()` returns a Promise<boolean>; a single <ConfirmHost/> in the
// root layout renders the modal and resolves the promise on OK / Cancel.
import React, { useEffect, useState } from "react";

type Req = {
  message: string;
  title?: string;
  confirmText?: string;
  cancelText?: string;
  danger?: boolean;
  resolve: (v: boolean) => void;
};

let register: ((r: Req | null) => void) | null = null;

export function confirmDialog(
  message: string,
  opts?: { title?: string; confirmText?: string; cancelText?: string; danger?: boolean }
): Promise<boolean> {
  return new Promise((resolve) => {
    if (typeof window === "undefined" || !register) {
      // Host not mounted (shouldn't happen once the app has rendered) — fall
      // back to the native dialog so an action is never silently lost.
      resolve(typeof window !== "undefined" ? window.confirm(message) : false);
      return;
    }
    register({ message, ...(opts || {}), resolve });
  });
}

export default function ConfirmHost() {
  const [req, setReq] = useState<Req | null>(null);

  useEffect(() => {
    register = setReq;
    return () => {
      register = null;
    };
  }, []);

  useEffect(() => {
    if (!req) return;
    const finish = (v: boolean) => {
      req.resolve(v);
      setReq(null);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") finish(false);
      else if (e.key === "Enter") finish(true);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [req]);

  if (!req) return null;
  const done = (v: boolean) => {
    req.resolve(v);
    setReq(null);
  };

  return (
    <div style={overlay} onClick={() => done(false)}>
      <div style={card} onClick={(e) => e.stopPropagation()} role="dialog" aria-modal="true">
        <div style={title}>{req.title || "Please confirm"}</div>
        <div style={message}>{req.message}</div>
        <div style={actions}>
          <button style={cancelBtn} onClick={() => done(false)}>
            {req.cancelText || "Cancel"}
          </button>
          <button
            style={{ ...confirmBtn, background: req.danger ? "#dc2626" : "#2563eb" }}
            autoFocus
            onClick={() => done(true)}
          >
            {req.confirmText || "OK"}
          </button>
        </div>
      </div>
    </div>
  );
}

const overlay: React.CSSProperties = {
  position: "fixed",
  inset: 0,
  background: "rgba(15,23,42,0.55)",
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  zIndex: 100001,
  padding: 16,
};
const card: React.CSSProperties = {
  width: "100%",
  maxWidth: 420,
  background: "#fff",
  borderRadius: 16,
  padding: "22px 22px 18px",
  boxShadow: "0 24px 70px rgba(0,0,0,0.35)",
  color: "#101828",
};
const title: React.CSSProperties = { fontSize: 17, fontWeight: 900, marginBottom: 8 };
const message: React.CSSProperties = {
  fontSize: 14,
  lineHeight: 1.5,
  color: "#475467",
  marginBottom: 20,
  whiteSpace: "pre-line",
};
const actions: React.CSSProperties = { display: "flex", justifyContent: "flex-end", gap: 10 };
const cancelBtn: React.CSSProperties = {
  padding: "10px 16px",
  borderRadius: 10,
  border: "1px solid #D0D5DD",
  background: "#fff",
  color: "#344054",
  fontWeight: 800,
  cursor: "pointer",
};
const confirmBtn: React.CSSProperties = {
  padding: "10px 18px",
  borderRadius: 10,
  border: "none",
  color: "#fff",
  fontWeight: 900,
  cursor: "pointer",
};
