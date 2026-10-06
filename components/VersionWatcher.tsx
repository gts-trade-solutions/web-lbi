"use client";

// Detects when the app has been redeployed and offers a one-tap refresh, so
// users (and the client's demos) always run the latest build instead of a stale
// cached one. The first poll records the build the tab loaded with; any later
// change means a new deploy is live.
import { useEffect, useRef, useState } from "react";

export default function VersionWatcher() {
  const [stale, setStale] = useState(false);
  const newBuildRef = useRef<string | null>(null);
  const dismissedRef = useRef<string | null>(null);

  useEffect(() => {
    // Only watch in production. In dev the build id churns and would pop the
    // banner constantly.
    if (process.env.NODE_ENV !== "production") return;

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
        // Show once per NEW build; stay quiet for a build the user dismissed.
        if (v !== baseline && v !== dismissedRef.current) {
          newBuildRef.current = v;
          setStale(true);
        }
      } catch {
        /* offline / transient — ignore */
      }
    };

    check();
    // Poll on an interval only. (Re-checking on tab focus / visibilitychange
    // used to fire right after the OS file picker when adding/changing a photo,
    // which popped this banner immediately after a photo edit.)
    const iv = setInterval(check, 3 * 60 * 1000); // every 3 minutes

    return () => {
      stopped = true;
      clearInterval(iv);
    };
  }, []);

  if (!stale) return null;

  const dismiss = () => {
    dismissedRef.current = newBuildRef.current;
    setStale(false);
  };

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
      <button
        type="button"
        onClick={dismiss}
        style={{
          background: "transparent",
          color: "#fff",
          border: "1px solid rgba(255,255,255,0.5)",
          borderRadius: 10,
          padding: "8px 14px",
          fontWeight: 800,
          cursor: "pointer",
        }}
      >
        Later
      </button>
    </div>
  );
}
