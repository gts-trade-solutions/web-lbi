"use client";

// Crop a report photo and save. The user drags a crop box (move + 8 resize
// handles) over the image; "Apply crop" renders the selected region at full
// resolution, uploads it as a new photo on the report (via /api/upload), and
// the parent removes the original so the CROPPED version replaces it.
//
// Cross-origin (S3) photos are loaded through /api/image-proxy so the canvas
// stays exportable (toBlob doesn't taint on a same-origin image).
import React, { useEffect, useRef, useState } from "react";

function authHeaders(): Record<string, string> {
  const t = typeof window !== "undefined" ? localStorage.getItem("auth_token") : null;
  return t ? { Authorization: `Bearer ${t}` } : {};
}

// Crop rect stored as fractions (0..1) of the image so it's resolution-free.
type Crop = { x: number; y: number; w: number; h: number };
type Handle = "move" | "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";

const MIN_FRAC = 0.05; // crop can't get smaller than 5% of the image

export default function PhotoCropper({
  photoUrl,
  reportId,
  onClose,
  onSaved,
}: {
  photoUrl: string;
  reportId: string;
  onClose: () => void;
  // Called after the cropped image uploaded successfully. The parent should
  // delete the ORIGINAL photo and refresh so the crop replaces it.
  onSaved?: (newUrl: string) => void;
}) {
  const imgRef = useRef<HTMLImageElement | null>(null);
  const boxRef = useRef<HTMLDivElement | null>(null); // the image's on-screen box
  const [loaded, setLoaded] = useState(false);
  const [failed, setFailed] = useState(false);
  const [saving, setSaving] = useState(false);
  const [crop, setCrop] = useState<Crop>({ x: 0.1, y: 0.1, w: 0.8, h: 0.8 });

  const proxied = (() => {
    const origin = typeof window !== "undefined" ? window.location.origin : "";
    const isExternal = /^https?:\/\//i.test(photoUrl) && !photoUrl.startsWith(origin);
    return isExternal ? `/api/image-proxy?url=${encodeURIComponent(photoUrl)}` : photoUrl;
  })();

  // Preload so we have the natural dimensions for the export crop.
  useEffect(() => {
    const img = new Image();
    img.onload = () => {
      imgRef.current = img;
      setLoaded(true);
    };
    img.onerror = () => setFailed(true);
    img.src = proxied;
  }, [proxied]);

  // ---- Drag/resize interaction (pointer based) ----
  const dragRef = useRef<{
    handle: Handle;
    startX: number;
    startY: number;
    start: Crop;
  } | null>(null);

  const clamp01 = (v: number) => Math.max(0, Math.min(1, v));

  const onPointerDown = (handle: Handle) => (e: React.PointerEvent) => {
    e.preventDefault();
    e.stopPropagation();
    (e.target as HTMLElement).setPointerCapture?.(e.pointerId);
    dragRef.current = { handle, startX: e.clientX, startY: e.clientY, start: crop };
  };

  const onPointerMove = (e: React.PointerEvent) => {
    const d = dragRef.current;
    const box = boxRef.current;
    if (!d || !box) return;
    const rect = box.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) return;
    const dx = (e.clientX - d.startX) / rect.width;
    const dy = (e.clientY - d.startY) / rect.height;

    let { x, y, w, h } = d.start;

    if (d.handle === "move") {
      x = clamp01(x + dx > 1 - w ? 1 - w : x + dx);
      y = clamp01(y + dy > 1 - h ? 1 - h : y + dy);
      x = Math.max(0, Math.min(x, 1 - w));
      y = Math.max(0, Math.min(y, 1 - h));
    } else {
      // Resize: adjust the touched edges, keeping the opposite edge fixed.
      let left = x;
      let top = y;
      let right = x + w;
      let bottom = y + h;
      if (d.handle.includes("w")) left = clamp01(Math.min(x + dx, right - MIN_FRAC));
      if (d.handle.includes("e")) right = clamp01(Math.max(x + w + dx, left + MIN_FRAC));
      if (d.handle.includes("n")) top = clamp01(Math.min(y + dy, bottom - MIN_FRAC));
      if (d.handle.includes("s")) bottom = clamp01(Math.max(y + h + dy, top + MIN_FRAC));
      x = left;
      y = top;
      w = right - left;
      h = bottom - top;
    }
    setCrop({ x, y, w, h });
  };

  const onPointerUp = (e: React.PointerEvent) => {
    (e.target as HTMLElement).releasePointerCapture?.(e.pointerId);
    dragRef.current = null;
  };

  const applyCrop = async () => {
    const img = imgRef.current;
    if (!img || saving) return;
    setSaving(true);
    try {
      const natW = img.naturalWidth || 0;
      const natH = img.naturalHeight || 0;
      const sx = Math.round(crop.x * natW);
      const sy = Math.round(crop.y * natH);
      const sw = Math.max(1, Math.round(crop.w * natW));
      const sh = Math.max(1, Math.round(crop.h * natH));

      const canvas = document.createElement("canvas");
      canvas.width = sw;
      canvas.height = sh;
      const ctx = canvas.getContext("2d");
      if (!ctx) throw new Error("Canvas not supported in this browser.");
      ctx.drawImage(img, sx, sy, sw, sh, 0, 0, sw, sh);

      const blob: Blob | null = await new Promise((resolve) =>
        canvas.toBlob((b) => resolve(b), "image/jpeg", 0.92)
      );
      if (!blob) {
        throw new Error("Couldn't export the crop — the photo host is blocking canvas export.");
      }

      const fd = new FormData();
      fd.append("file", blob, `cropped_${Date.now()}.jpg`);
      fd.append("folder", "uploads");
      fd.append("reportId", reportId);
      fd.append("width", String(sw));
      fd.append("height", String(sh));

      const res = await fetch("/api/upload", {
        method: "POST",
        credentials: "include",
        headers: authHeaders(),
        body: fd,
      });
      const data = await res.json().catch(() => ({} as Record<string, unknown>));
      if (!res.ok) throw new Error((data as { error?: string })?.error || "Upload failed");
      const newUrl = String(
        (data as { url?: string; photo?: { url?: string } })?.url ||
          (data as { photo?: { url?: string } })?.photo?.url ||
          ""
      );
      onSaved?.(newUrl);
      onClose();
    } catch (err) {
      alert((err as { message?: string })?.message || "Failed to save the cropped image.");
    } finally {
      setSaving(false);
    }
  };

  const pct = (n: number) => `${n * 100}%`;
  const handleStyle: React.CSSProperties = {
    position: "absolute",
    width: 16,
    height: 16,
    background: "#fff",
    border: "2px solid #7C3AED",
    borderRadius: 4,
    boxShadow: "0 1px 4px rgba(0,0,0,0.4)",
    touchAction: "none",
  };

  return (
    <div
      style={{
        position: "fixed",
        inset: 0,
        zIndex: 1400,
        background: "rgba(8,12,20,0.82)",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        padding: 16,
        gap: 14,
      }}
      onMouseDown={onClose}
    >
      <div
        style={{ display: "grid", gap: 12, maxWidth: "96vw" }}
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <strong style={{ color: "#fff", fontSize: 16 }}>✂️ Crop photo</strong>
          <span style={{ color: "#98A2B3", fontSize: 12, fontWeight: 700 }}>
            Drag the box to move · drag a handle to resize
          </span>
          <button
            type="button"
            onClick={onClose}
            style={{
              marginLeft: "auto",
              width: 34,
              height: 34,
              borderRadius: 10,
              border: "1px solid #475467",
              background: "#101828",
              color: "#fff",
              fontWeight: 900,
              cursor: "pointer",
            }}
            aria-label="Close"
          >
            ✕
          </button>
        </div>

        {failed ? (
          <div style={{ color: "#FDA29B", fontWeight: 800, padding: 40 }}>
            Couldn&apos;t load this image for cropping.
          </div>
        ) : !loaded ? (
          <div style={{ color: "#fff", fontWeight: 800, padding: 40 }}>Loading photo…</div>
        ) : (
          <div
            ref={boxRef}
            style={{
              position: "relative",
              lineHeight: 0,
              maxWidth: "min(1100px, 92vw)",
              maxHeight: "72vh",
              userSelect: "none",
              touchAction: "none",
            }}
            onPointerMove={onPointerMove}
            onPointerUp={onPointerUp}
          >
            {/* eslint-disable-next-line @next/next/no-img-element */}
            <img
              src={proxied}
              alt="Crop source"
              draggable={false}
              style={{
                display: "block",
                maxWidth: "min(1100px, 92vw)",
                maxHeight: "72vh",
                width: "auto",
                height: "auto",
                borderRadius: 8,
              }}
            />

            {/* Dim everything outside the crop rect (4 contained panels). */}
            <div style={{ position: "absolute", left: 0, top: 0, width: "100%", height: pct(crop.y), background: "rgba(0,0,0,0.5)", pointerEvents: "none" }} />
            <div style={{ position: "absolute", left: 0, top: pct(crop.y + crop.h), width: "100%", bottom: 0, background: "rgba(0,0,0,0.5)", pointerEvents: "none" }} />
            <div style={{ position: "absolute", left: 0, top: pct(crop.y), width: pct(crop.x), height: pct(crop.h), background: "rgba(0,0,0,0.5)", pointerEvents: "none" }} />
            <div style={{ position: "absolute", left: pct(crop.x + crop.w), top: pct(crop.y), right: 0, height: pct(crop.h), background: "rgba(0,0,0,0.5)", pointerEvents: "none" }} />

            {/* Crop rectangle (drag body = move) */}
            <div
              onPointerDown={onPointerDown("move")}
              style={{
                position: "absolute",
                left: pct(crop.x),
                top: pct(crop.y),
                width: pct(crop.w),
                height: pct(crop.h),
                border: "2px solid #fff",
                cursor: "move",
                touchAction: "none",
                boxSizing: "border-box",
              }}
            >
              {/* rule-of-thirds guides */}
              <div style={{ position: "absolute", left: "33.33%", top: 0, bottom: 0, width: 1, background: "rgba(255,255,255,0.4)" }} />
              <div style={{ position: "absolute", left: "66.66%", top: 0, bottom: 0, width: 1, background: "rgba(255,255,255,0.4)" }} />
              <div style={{ position: "absolute", top: "33.33%", left: 0, right: 0, height: 1, background: "rgba(255,255,255,0.4)" }} />
              <div style={{ position: "absolute", top: "66.66%", left: 0, right: 0, height: 1, background: "rgba(255,255,255,0.4)" }} />

              {/* 8 resize handles */}
              <div onPointerDown={onPointerDown("nw")} style={{ ...handleStyle, left: -8, top: -8, cursor: "nwse-resize" }} />
              <div onPointerDown={onPointerDown("n")} style={{ ...handleStyle, left: "calc(50% - 8px)", top: -8, cursor: "ns-resize" }} />
              <div onPointerDown={onPointerDown("ne")} style={{ ...handleStyle, right: -8, top: -8, cursor: "nesw-resize" }} />
              <div onPointerDown={onPointerDown("e")} style={{ ...handleStyle, right: -8, top: "calc(50% - 8px)", cursor: "ew-resize" }} />
              <div onPointerDown={onPointerDown("se")} style={{ ...handleStyle, right: -8, bottom: -8, cursor: "nwse-resize" }} />
              <div onPointerDown={onPointerDown("s")} style={{ ...handleStyle, left: "calc(50% - 8px)", bottom: -8, cursor: "ns-resize" }} />
              <div onPointerDown={onPointerDown("sw")} style={{ ...handleStyle, left: -8, bottom: -8, cursor: "nesw-resize" }} />
              <div onPointerDown={onPointerDown("w")} style={{ ...handleStyle, left: -8, top: "calc(50% - 8px)", cursor: "ew-resize" }} />
            </div>
          </div>
        )}

        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button
            type="button"
            onClick={onClose}
            disabled={saving}
            style={{
              padding: "11px 18px",
              borderRadius: 12,
              border: "1px solid #475467",
              background: "#101828",
              color: "#fff",
              fontWeight: 900,
              cursor: saving ? "not-allowed" : "pointer",
            }}
          >
            Cancel
          </button>
          <button
            type="button"
            onClick={applyCrop}
            disabled={saving || !loaded}
            style={{
              padding: "11px 18px",
              borderRadius: 12,
              border: "none",
              background: "#7C3AED",
              color: "#fff",
              fontWeight: 900,
              cursor: saving || !loaded ? "not-allowed" : "pointer",
              opacity: saving || !loaded ? 0.6 : 1,
            }}
          >
            {saving ? "Saving…" : "Apply crop & save"}
          </button>
        </div>
      </div>
    </div>
  );
}
