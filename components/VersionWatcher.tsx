"use client";

// Detects when the app has been redeployed and offers a one-tap refresh, so
// users (and the client's demos) always run the latest build instead of a stale
// cached one. The first poll records the build the tab loaded with; any later
// change means a new deploy is live.
import { useEffect, useState } from "react";

export default function VersionWatcher() {
  const [stale, setStale] = useState(false);

  useEffect(() => {
    let baseline: string | null = null;
    let stopped = false;

    const check = async () => {
      try {
        const res = await fetch("/api/version", { cache: "no-store" });
        if (!res.ok) return;
        const data = (await res.json().catch(() => null)) as { buildId?: string } | null;
        const v = String(data?.buildId || "");
        if (!v || stopped) return;
        if (baseline === null) {
          baseline = v; // first successful poll = the build this tab is running
          return;
        }
        if (v !== baseline) setStale(true);
      } catch {
        /* offline / transient — ignore */
      }
    };

    check();
    const iv = setInterval(check, 3 * 60 * 1000); // every 3 minutes
    const onVisible = () => {
      if (document.visibilityState === "visible") check();
    };
    document.addEventListener("visibilitychange", onVisible);

    return () => {
      stopped = true;
      clearInterval(iv);
      document.removeEventListener("visibilitychange", onVisible);
    };
  }, []);

  if (!stale) return null;

  return (
    <div
      role="alert"
      style={{
        position: "fixed",
        left: 0,
        right: 0,
        bottom: 0,
        zIndex: 100000,
        background: "#101828",
        color: "#fff",
        padding: "10px 16px",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        gap: 12,
        flexWrap: "wrap",
        fontWeight: 800,
        boxShadow: "0 -6px 20px rgba(16,24,40,0.28)",
      }}
    >
      <span>A new version of the app is available.</span>
      <button
        type="button"
        onClick={() => window.location.reload()}
        style={{
          background: "#12B76A",
          color: "#fff",
          border: "none",
          borderRadius: 10,
          padding: "8px 18px",
          fontWeight: 900,
          cursor: "pointer",
        }}
      >
        Refresh now
      </button>
    </div>
  );
}
