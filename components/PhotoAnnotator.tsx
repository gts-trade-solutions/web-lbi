"use client";

// Draw arrows / lines / shapes / text labels (e.g. "H-6.2m") directly onto a
// report photo and Save — the flattened image is uploaded as a new photo on the
// report (via /api/upload), so it flows into the grid and the Word export.
//
// This is the same drawing tool used in the Route Map (route3d) view, packaged
// as a reusable full-screen editor so the reports grid photo viewer can use it
// too. Cross-origin (S3) photos are loaded through /api/image-proxy so the
// canvas stays exportable (toBlob doesn't taint).
import React, { useCallback, useEffect, useRef, useState } from "react";

type Stroke = {
  tool: "arrow" | "darrow" | "line" | "pen" | "rect" | "ellipse" | "left" | "right" | "uturn" | "x" | "text";
  color: string;
  width: number;
  points: { x: number; y: number }[];
  text?: string;
  fontSize?: number;
};

type TextDraft = {
  dispX: number;
  dispY: number;
  cx: number;
  cy: number;
  font: number;
  naturalFont: number;
  value: string;
};

function authHeaders(): Record<string, string> {
  const t = typeof window !== "undefined" ? localStorage.getItem("auth_token") : null;
  return t ? { Authorization: `Bearer ${t}` } : {};
}

export default function PhotoAnnotator({
  photoUrl,
  reportId,
  onClose,
  onSaved,
}: {
  photoUrl: string;
  reportId: string;
  onClose: () => void;
  onSaved?: (newUrl: string) => void;
}) {
  const [drawTool, setDrawTool] = useState<Stroke["tool"] | "move">("arrow");
  const [drawColor, setDrawColor] = useState("#FFD400");
  const [drawWidth, setDrawWidth] = useState(6);
  const [strokes, setStrokes] = useState<Stroke[]>([]);
  const [saving, setSaving] = useState(false);
  const [textDraft, setTextDraft] = useState<TextDraft | null>(null);
  const [selectedIdx, setSelectedIdx] = useState<number | null>(null);

  const imgWrapRef = useRef<HTMLDivElement | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const drawingRef = useRef<Stroke | null>(null);
  const baseImgRef = useRef<HTMLImageElement | null>(null);
  const moveDragRef = useRef<{ x: number; y: number } | null>(null);

  // ---- Drawing primitives (copied from the Route Map annotator) ----
  const drawTurnArrow = (
    ctx: CanvasRenderingContext2D,
    type: "left" | "right" | "uturn",
    a: { x: number; y: number },
    b: { x: number; y: number },
    color: string,
    width: number
  ) => {
    const L = Math.min(a.x, b.x), R = Math.max(a.x, b.x);
    const T = Math.min(a.y, b.y), B = Math.max(a.y, b.y);
    const W = R - L, H = B - T;
    if (W < 4 || H < 4) return;
    const cx = (L + R) / 2;
    const head = Math.max(14, width * 3.5);
    ctx.strokeStyle = color;
    ctx.fillStyle = color;
    ctx.lineWidth = width;
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
    const headAt = (x: number, y: number, angle: number) => {
      ctx.beginPath();
      ctx.moveTo(x, y);
      ctx.lineTo(x - head * Math.cos(angle - Math.PI / 7), y - head * Math.sin(angle - Math.PI / 7));
      ctx.lineTo(x - head * Math.cos(angle + Math.PI / 7), y - head * Math.sin(angle + Math.PI / 7));
      ctx.closePath();
      ctx.fill();
    };
    if (type === "uturn") {
      const legOffset = Math.min(W * 0.28, H * 0.4);
      const lx = cx - legOffset, rx = cx + legOffset, topY = T + legOffset;
      ctx.beginPath();
      ctx.moveTo(lx, B);
      ctx.lineTo(lx, topY);
      ctx.arc(cx, topY, legOffset, Math.PI, 2 * Math.PI, false);
      ctx.lineTo(rx, B - head);
      ctx.stroke();
      headAt(rx, B, Math.PI / 2);
      return;
    }
    const side = type === "left" ? L : R;
    ctx.beginPath();
    ctx.moveTo(cx, B);
    ctx.lineTo(cx, T + H * 0.5);
    ctx.quadraticCurveTo(cx, T, side, T);
    ctx.stroke();
    headAt(side, T, type === "left" ? Math.PI : 0);
  };

  const drawOneStroke = (ctx: CanvasRenderingContext2D, s: Stroke) => {
    ctx.strokeStyle = s.color;
    ctx.fillStyle = s.color;
    ctx.lineWidth = s.width;
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
    const pts = s.points;
    if (!pts.length) return;
    if (s.tool === "text") {
      ctx.textBaseline = "top";
      ctx.font = `bold ${s.fontSize || 32}px system-ui, Segoe UI, Arial, sans-serif`;
      ctx.lineWidth = Math.max(2, (s.fontSize || 32) * 0.06);
      ctx.strokeStyle = "rgba(0,0,0,0.55)";
      ctx.strokeText(s.text || "", pts[0].x, pts[0].y);
      ctx.fillStyle = s.color;
      ctx.fillText(s.text || "", pts[0].x, pts[0].y);
      return;
    }
    if (s.tool === "pen") {
      ctx.beginPath();
      ctx.moveTo(pts[0].x, pts[0].y);
      for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i].x, pts[i].y);
      ctx.stroke();
      return;
    }
    const a = pts[0], b = pts[pts.length - 1];
    if (s.tool === "rect") {
      ctx.strokeRect(Math.min(a.x, b.x), Math.min(a.y, b.y), Math.abs(b.x - a.x), Math.abs(b.y - a.y));
      return;
    }
    if (s.tool === "ellipse") {
      ctx.beginPath();
      ctx.ellipse((a.x + b.x) / 2, (a.y + b.y) / 2, Math.abs(b.x - a.x) / 2, Math.abs(b.y - a.y) / 2, 0, 0, Math.PI * 2);
      ctx.stroke();
      return;
    }
    if (s.tool === "left" || s.tool === "right" || s.tool === "uturn") {
      drawTurnArrow(ctx, s.tool, a, b, s.color, s.width);
      return;
    }
    if (s.tool === "x") {
      ctx.beginPath();
      ctx.moveTo(a.x, a.y);
      ctx.lineTo(b.x, b.y);
      ctx.moveTo(b.x, a.y);
      ctx.lineTo(a.x, b.y);
      ctx.stroke();
      return;
    }
    // line / arrow / double-arrow: shaft, plus a filled head on one or both ends.
    ctx.beginPath();
    ctx.moveTo(a.x, a.y);
    ctx.lineTo(b.x, b.y);
    ctx.stroke();
    const angle = Math.atan2(b.y - a.y, b.x - a.x);
    const head = Math.max(14, s.width * 3.5);
    const arrowHead = (hx: number, hy: number, ang: number) => {
      ctx.beginPath();
      ctx.moveTo(hx, hy);
      ctx.lineTo(hx - head * Math.cos(ang - Math.PI / 7), hy - head * Math.sin(ang - Math.PI / 7));
      ctx.lineTo(hx - head * Math.cos(ang + Math.PI / 7), hy - head * Math.sin(ang + Math.PI / 7));
      ctx.closePath();
      ctx.fill();
    };
    if (s.tool === "arrow" || s.tool === "darrow") arrowHead(b.x, b.y, angle); // head at end
    if (s.tool === "darrow") arrowHead(a.x, a.y, angle + Math.PI); // head at start too
    // "line" → no heads
  };

  const strokeBBox = (s: Stroke) => {
    if (s.tool === "text") {
      const fs = s.fontSize || 32;
      const w = Math.max(20, (s.text?.length || 1) * fs * 0.6);
      return { minX: s.points[0].x, minY: s.points[0].y, maxX: s.points[0].x + w, maxY: s.points[0].y + fs * 1.2 };
    }
    const xs = s.points.map((p) => p.x);
    const ys = s.points.map((p) => p.y);
    return { minX: Math.min(...xs), minY: Math.min(...ys), maxX: Math.max(...xs), maxY: Math.max(...ys) };
  };

  const redrawCanvas = useCallback(
    (extra?: Stroke | null) => {
      const canvas = canvasRef.current;
      if (!canvas) return;
      const ctx = canvas.getContext("2d");
      if (!ctx) return;
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      const base = baseImgRef.current;
      if (base) ctx.drawImage(base, 0, 0, canvas.width, canvas.height);
      for (const s of strokes) drawOneStroke(ctx, s);
      if (extra) drawOneStroke(ctx, extra);
      if (drawTool === "move" && selectedIdx != null && strokes[selectedIdx]) {
        const b = strokeBBox(strokes[selectedIdx]);
        ctx.save();
        ctx.setLineDash([10, 7]);
        ctx.strokeStyle = "#22D3EE";
        ctx.lineWidth = 2;
        ctx.strokeRect(b.minX - 8, b.minY - 8, b.maxX - b.minX + 16, b.maxY - b.minY + 16);
        ctx.restore();
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [strokes, drawTool, selectedIdx]
  );

  // Load the photo into the canvas on mount. External (S3) URLs go through the
  // same-origin proxy so the canvas stays exportable.
  useEffect(() => {
    if (!photoUrl) return;
    const origin = typeof window !== "undefined" ? window.location.origin : "";
    const isExternal = /^https?:\/\//i.test(photoUrl) && !photoUrl.startsWith(origin);
    const url = isExternal ? `/api/image-proxy?url=${encodeURIComponent(photoUrl)}` : photoUrl;
    let cancelled = false;
    const img = new Image();
    img.onload = () => {
      if (cancelled) return;
      baseImgRef.current = img;
      const w = img.naturalWidth || 1200;
      const h = img.naturalHeight || 800;
      const canvas = canvasRef.current;
      if (canvas) {
        canvas.width = w;
        canvas.height = h;
        const maxW = Math.min(window.innerWidth * 0.9, 1280);
        const maxH = window.innerHeight * 0.72;
        const scale = Math.min(maxW / w, maxH / h, 1);
        canvas.style.width = `${Math.round(w * scale)}px`;
        canvas.style.height = `${Math.round(h * scale)}px`;
      }
      redrawCanvas();
    };
    img.onerror = () => {
      if (!cancelled) baseImgRef.current = null;
    };
    img.src = url;
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [photoUrl]);

  useEffect(() => {
    redrawCanvas();
  }, [strokes, selectedIdx, drawTool, redrawCanvas]);

  const canvasPoint = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current!;
    const rect = canvas.getBoundingClientRect();
    return {
      x: ((e.clientX - rect.left) / rect.width) * canvas.width,
      y: ((e.clientY - rect.top) / rect.height) * canvas.height,
    };
  };
  const hitTest = (p: { x: number; y: number }): number | null => {
    const pad = 12;
    for (let i = strokes.length - 1; i >= 0; i--) {
      const b = strokeBBox(strokes[i]);
      if (p.x >= b.minX - pad && p.x <= b.maxX + pad && p.y >= b.minY - pad && p.y <= b.maxY + pad) return i;
    }
    return null;
  };
  const resizeSelected = (factor: number) => {
    if (selectedIdx == null) return;
    setStrokes((prev) =>
      prev.map((s, i) => {
        if (i !== selectedIdx) return s;
        if (s.tool === "text") return { ...s, fontSize: Math.max(10, Math.round((s.fontSize || 32) * factor)) };
        const b = strokeBBox(s);
        const cx = (b.minX + b.maxX) / 2;
        const cy = (b.minY + b.maxY) / 2;
        return {
          ...s,
          width: Math.min(60, Math.max(1, s.width * factor)),
          points: s.points.map((p) => ({ x: cx + (p.x - cx) * factor, y: cy + (p.y - cy) * factor })),
        };
      })
    );
  };
  const deleteSelected = () => {
    if (selectedIdx == null) return;
    setStrokes((prev) => prev.filter((_, i) => i !== selectedIdx));
    setSelectedIdx(null);
  };

  const onDrawDown = (e: React.PointerEvent<HTMLCanvasElement>) => {
    e.preventDefault();
    if (textDraft) {
      const d = textDraft;
      if (d.value.trim()) {
        setStrokes((prev) => [
          ...prev,
          { tool: "text", color: drawColor, width: drawWidth, points: [{ x: d.cx, y: d.cy }], text: d.value.trim(), fontSize: d.naturalFont },
        ]);
      }
      setTextDraft(null);
      if (drawTool !== "text") return;
    }
    const p = canvasPoint(e);
    if (drawTool === "move") {
      (e.target as HTMLCanvasElement).setPointerCapture(e.pointerId);
      const idx = hitTest(p);
      setSelectedIdx(idx);
      moveDragRef.current = idx != null ? { x: p.x, y: p.y } : null;
      redrawCanvas();
      return;
    }
    if (drawTool === "text") {
      const canvas = canvasRef.current;
      const wrap = imgWrapRef.current;
      if (!canvas || !wrap) return;
      const wrapRect = wrap.getBoundingClientRect();
      const canvasRect = canvas.getBoundingClientRect();
      const naturalFont = Math.max(28, drawWidth * 7);
      const scale = canvas.width ? canvasRect.width / canvas.width : 1;
      setTextDraft({
        dispX: e.clientX - wrapRect.left,
        dispY: e.clientY - wrapRect.top,
        cx: p.x,
        cy: p.y,
        naturalFont,
        font: naturalFont * scale,
        value: "",
      });
      return;
    }
    (e.target as HTMLCanvasElement).setPointerCapture(e.pointerId);
    drawingRef.current = { tool: drawTool as Stroke["tool"], color: drawColor, width: drawWidth, points: [p, p] };
    redrawCanvas(drawingRef.current);
  };
  const onDrawMove = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const p = canvasPoint(e);
    if (drawTool === "move") {
      if (selectedIdx == null || !moveDragRef.current) return;
      const dx = p.x - moveDragRef.current.x;
      const dy = p.y - moveDragRef.current.y;
      moveDragRef.current = { x: p.x, y: p.y };
      setStrokes((prev) => prev.map((s, i) => (i === selectedIdx ? { ...s, points: s.points.map((pt) => ({ x: pt.x + dx, y: pt.y + dy })) } : s)));
      return;
    }
    if (!drawingRef.current) return;
    if (drawTool === "pen") drawingRef.current.points.push(p);
    else drawingRef.current.points[1] = p;
    redrawCanvas(drawingRef.current);
  };
  const onDrawUp = () => {
    moveDragRef.current = null;
    if (!drawingRef.current) return;
    const s = drawingRef.current;
    drawingRef.current = null;
    setStrokes((prev) => [...prev, s]);
  };

  const commitText = () => {
    const d = textDraft;
    if (d && d.value.trim()) {
      setStrokes((prev) => [
        ...prev,
        { tool: "text", color: drawColor, width: drawWidth, points: [{ x: d.cx, y: d.cy }], text: d.value.trim(), fontSize: d.naturalFont },
      ]);
    }
    setTextDraft(null);
  };

  const saveAnnotated = async () => {
    const canvas = canvasRef.current;
    if (!canvas || saving) return;
    setSaving(true);
    try {
      if (textDraft && textDraft.value.trim()) {
        const ctx = canvas.getContext("2d");
        if (ctx)
          drawOneStroke(ctx, { tool: "text", color: drawColor, width: drawWidth, points: [{ x: textDraft.cx, y: textDraft.cy }], text: textDraft.value.trim(), fontSize: textDraft.naturalFont });
        setTextDraft(null);
      }
      const blob: Blob | null = await new Promise((resolve) => canvas.toBlob((b) => resolve(b), "image/jpeg", 0.92));
      if (!blob) throw new Error("Couldn't export the annotated image — the photo host is blocking cross-origin canvas export.");
      const fd = new FormData();
      fd.append("file", blob, `annotated_${Date.now()}.jpg`);
      fd.append("folder", "uploads");
      fd.append("reportId", reportId);
      fd.append("width", String(canvas.width));
      fd.append("height", String(canvas.height));
      const res = await fetch("/api/upload", { method: "POST", credentials: "include", headers: authHeaders(), body: fd });
      const data = await res.json().catch(() => ({} as Record<string, unknown>));
      if (!res.ok) throw new Error((data as { error?: string })?.error || "Upload failed");
      const newUrl = String((data as { url?: string; photo?: { url?: string } })?.url || (data as { photo?: { url?: string } })?.photo?.url || "");
      onSaved?.(newUrl);
      onClose();
    } catch (err) {
      alert((err as { message?: string })?.message || "Failed to save the annotated image.");
    } finally {
      setSaving(false);
    }
  };

  const total = strokes.length;

  return (
    <div style={S.overlay} onClick={onClose}>
      <div style={S.card} onClick={(e) => e.stopPropagation()}>
        <div style={S.head}>
          <strong style={{ fontSize: 15 }}>✏️ Draw on photo</strong>
          <button style={S.close} onClick={onClose} aria-label="Close">✕</button>
        </div>

        <div style={S.imgWrap} ref={imgWrapRef}>
          <canvas
            ref={canvasRef}
            style={S.canvas}
            onPointerDown={onDrawDown}
            onPointerMove={onDrawMove}
            onPointerUp={onDrawUp}
            onPointerLeave={onDrawUp}
          />
          {textDraft ? (
            <input
              autoFocus
              value={textDraft.value}
              placeholder="Short label…"
              maxLength={40}
              onChange={(e) => setTextDraft((d) => (d ? { ...d, value: e.target.value } : d))}
              onKeyDown={(e) => {
                if (e.key === "Enter") { e.preventDefault(); commitText(); }
                else if (e.key === "Escape") setTextDraft(null);
              }}
              onBlur={commitText}
              style={{
                position: "absolute",
                left: textDraft.dispX,
                top: textDraft.dispY,
                font: `bold ${Math.max(12, textDraft.font)}px system-ui, Segoe UI, Arial, sans-serif`,
                color: drawColor,
                background: "rgba(0,0,0,0.35)",
                border: "1px dashed rgba(255,255,255,0.85)",
                outline: "none",
                padding: "0 2px",
                margin: 0,
                lineHeight: 1.1,
                minWidth: 60,
                maxWidth: "90%",
                caretColor: drawColor,
                zIndex: 5,
              }}
            />
          ) : null}
        </div>

        <div style={S.tools}>
          {([
            ["move", "✥", "Move / select"],
            ["arrow", "➔", "Arrow"],
            ["darrow", "↔", "Double arrow (measure H / W)"],
            ["line", "—", "Straight line"],
            ["pen", "✎", "Free draw (scribble anything)"],
            ["rect", "▭", "Rectangle"],
            ["ellipse", "◯", "Ellipse"],
            ["left", "↰", "Left turn"],
            ["right", "↱", "Right turn"],
            ["uturn", "↩", "U-turn"],
            ["x", "✕", "X mark"],
            ["text", "T", "Text"],
          ] as const).map(([t, icon, label]) => (
            <button key={t} onClick={() => setDrawTool(t)} title={label} style={{ ...S.toolBtn, ...(drawTool === t ? S.toolBtnActive : {}) }}>
              {icon}
            </button>
          ))}
          {([
            ["#FFD400", "Yellow"],
            ["#FF7A00", "Orange"],
            ["#1E3A8A", "Navy blue"],
            ["#E11D2E", "Red"],
          ] as const).map(([c, label]) => (
            <button
              key={c}
              onClick={() => setDrawColor(c)}
              title={label}
              aria-label={label}
              style={{ ...S.swatch, background: c, outline: drawColor.toLowerCase() === c.toLowerCase() ? "2px solid #0f172a" : "2px solid transparent" }}
            />
          ))}
          <input type="color" value={drawColor} onChange={(e) => setDrawColor(e.target.value)} style={S.colorInput} title="Custom colour" />
          <select value={drawWidth} onChange={(e) => setDrawWidth(Number(e.target.value))} style={S.widthSelect} title="Line width">
            <option value={3}>Thin</option>
            <option value={6}>Medium</option>
            <option value={10}>Thick</option>
          </select>
        </div>

        {drawTool === "move" ? (
          <div style={{ fontSize: 11, fontWeight: 700, color: "#64748b" }}>
            {selectedIdx == null ? "Tap any drawing or text to select, then drag to move." : "Drag to move · − / + to resize · Delete to remove."}
          </div>
        ) : null}
        {drawTool === "move" && selectedIdx != null ? (
          <div style={S.actions}>
            <button style={S.ghost} onClick={() => resizeSelected(0.85)}>− Smaller</button>
            <button style={S.ghost} onClick={() => resizeSelected(1.18)}>Larger +</button>
            <button style={S.ghost} onClick={deleteSelected}>Delete</button>
          </div>
        ) : null}

        <div style={S.actions}>
          <button style={S.ghost} onClick={() => setStrokes((s) => s.slice(0, -1))} disabled={!total}>Undo</button>
          <button style={S.ghost} onClick={() => setStrokes([])} disabled={!total}>Clear</button>
          <button style={S.ghost} onClick={onClose}>Cancel</button>
          <button style={S.primary} onClick={saveAnnotated} disabled={saving || !total}>
            {saving ? "Saving…" : "Save"}
          </button>
        </div>
      </div>
    </div>
  );
}

const S: Record<string, React.CSSProperties> = {
  overlay: { position: "fixed", inset: 0, background: "rgba(2,6,23,0.72)", zIndex: 4000, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 },
  card: { background: "#fff", borderRadius: 14, padding: 14, maxWidth: "min(96vw, 1360px)", maxHeight: "94vh", overflow: "auto", display: "flex", flexDirection: "column", gap: 10, boxShadow: "0 24px 60px rgba(0,0,0,0.4)" },
  head: { display: "flex", alignItems: "center", justifyContent: "space-between", color: "#0f172a" },
  close: { width: 34, height: 34, borderRadius: 8, border: "1px solid #e2e8f0", background: "#f8fafc", cursor: "pointer", fontSize: 15, color: "#0f172a" },
  imgWrap: { position: "relative", alignSelf: "center", lineHeight: 0 },
  canvas: { touchAction: "none", cursor: "crosshair", borderRadius: 8, display: "block", background: "#0b1220" },
  tools: { display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" },
  toolBtn: { width: 38, height: 38, borderRadius: 8, border: "1px solid #d7dbe0", background: "#fff", fontSize: 18, cursor: "pointer", color: "#0f172a" },
  toolBtnActive: { background: "#0f172a", color: "#fff", borderColor: "#0f172a" },
  colorInput: { width: 38, height: 38, border: "1px solid #d7dbe0", borderRadius: 8, cursor: "pointer", padding: 2, background: "#fff" },
  swatch: { width: 30, height: 30, borderRadius: "50%", border: "1px solid rgba(0,0,0,0.15)", cursor: "pointer", outlineOffset: 2, flexShrink: 0 },
  widthSelect: { height: 38, borderRadius: 8, border: "1px solid #d7dbe0", padding: "0 8px", fontSize: 14, cursor: "pointer" },
  actions: { display: "flex", gap: 6, flexWrap: "wrap" },
  ghost: { flex: 1, padding: "9px 10px", borderRadius: 8, border: "1px solid #d7dbe0", background: "#fff", fontWeight: 600, fontSize: 13, cursor: "pointer", color: "#0f172a" },
  primary: { flex: 1, padding: "9px 10px", borderRadius: 8, border: "none", background: "#16a34a", color: "#fff", fontWeight: 700, fontSize: 13, cursor: "pointer" },
};
