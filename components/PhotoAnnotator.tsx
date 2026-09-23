"use client";

// Draw arrows / lines / shapes / text labels (e.g. "H-6.2m") directly onto a
// report photo and Save — the flattened image REPLACES the photo it was drawn
// on (same position, same export settings), so it flows into the grid and the
// Word export without adding a second copy.
//
// This is the same drawing tool used in the Route Map (route3d) view, packaged
// as a reusable full-screen editor so the reports grid photo viewer can use it
// too. Cross-origin (S3) photos are loaded through /api/image-proxy so the
// canvas stays exportable (toBlob doesn't taint).
import React, { useCallback, useEffect, useRef, useState } from "react";
import { replacePhotoImage } from "../lib/replacePhoto";
import { loadCanvasSafeImage } from "../lib/authedImage";

type Stroke = {
  tool:
    | "arrow" | "darrow" | "line" | "pen"
    | "rect" | "rrect" | "ellipse" | "triangle" | "diamond" | "star"
    | "pentagon" | "hexagon" | "callout" | "uparrow" | "downarrow"
    | "left" | "right" | "uturn" | "x" | "text";
  color: string;
  width: number;
  points: { x: number; y: number }[];
  text?: string;
  fontSize?: number;
  fill?: boolean;
  rot?: number; // rotation in radians, around the shape's bbox centre
  dash?: "solid" | "dashed" | "dotted";
};

// Closed shapes that can be outlined or filled.
const CLOSED_SHAPES = new Set([
  "rect", "rrect", "ellipse", "triangle", "diamond", "star",
  "pentagon", "hexagon", "callout", "uparrow", "downarrow",
]);

// Rounded-rectangle path (manual, so it works without ctx.roundRect support).
function roundRectPath(ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number, r: number) {
  const rr = Math.max(0, Math.min(r, w / 2, h / 2));
  ctx.moveTo(x + rr, y);
  ctx.arcTo(x + w, y, x + w, y + h, rr);
  ctx.arcTo(x + w, y + h, x, y + h, rr);
  ctx.arcTo(x, y + h, x, y, rr);
  ctx.arcTo(x, y, x + w, y, rr);
  ctx.closePath();
}

// N-point star path centred at (cx,cy), outer radius R (inner = R*0.5).
function starPath(ctx: CanvasRenderingContext2D, cx: number, cy: number, R: number, points = 5) {
  const inner = R * 0.5;
  for (let i = 0; i < points * 2; i++) {
    const rad = i % 2 === 0 ? R : inner;
    const ang = (Math.PI / points) * i - Math.PI / 2;
    const x = cx + rad * Math.cos(ang);
    const y = cy + rad * Math.sin(ang);
    if (i === 0) ctx.moveTo(x, y);
    else ctx.lineTo(x, y);
  }
  ctx.closePath();
}

// Regular N-gon (pentagon=5, hexagon=6) fitted to the bbox, flat-ish top.
function polygonPath(ctx: CanvasRenderingContext2D, cx: number, cy: number, rx: number, ry: number, sides: number) {
  for (let i = 0; i < sides; i++) {
    const ang = (2 * Math.PI * i) / sides - Math.PI / 2;
    const x = cx + rx * Math.cos(ang);
    const y = cy + ry * Math.sin(ang);
    if (i === 0) ctx.moveTo(x, y);
    else ctx.lineTo(x, y);
  }
  ctx.closePath();
}

// Speech-bubble callout: rounded box with a small tail at the bottom-left.
function calloutPath(ctx: CanvasRenderingContext2D, L: number, T: number, W: number, H: number) {
  const bodyH = H * 0.78;
  const B = T + bodyH;
  const r = Math.min(W, bodyH) * 0.16;
  roundRectPath(ctx, L, T, W, bodyH, r);
  ctx.moveTo(L + W * 0.22, B);
  ctx.lineTo(L + W * 0.12, T + H);
  ctx.lineTo(L + W * 0.42, B);
}

// Block arrow (up or down) fitted to the bbox.
function blockArrowPath(ctx: CanvasRenderingContext2D, dir: "up" | "down", L: number, T: number, W: number, H: number) {
  const shaftW = W * 0.44;
  const sx0 = L + (W - shaftW) / 2;
  const sx1 = sx0 + shaftW;
  const headH = H * 0.45;
  if (dir === "up") {
    const tip = T, neck = T + headH, bot = T + H;
    ctx.moveTo(L + W / 2, tip);
    ctx.lineTo(L + W, neck);
    ctx.lineTo(sx1, neck);
    ctx.lineTo(sx1, bot);
    ctx.lineTo(sx0, bot);
    ctx.lineTo(sx0, neck);
    ctx.lineTo(L, neck);
  } else {
    const tip = T + H, neck = T + H - headH, top = T;
    ctx.moveTo(L + W / 2, tip);
    ctx.lineTo(L + W, neck);
    ctx.lineTo(sx1, neck);
    ctx.lineTo(sx1, top);
    ctx.lineTo(sx0, top);
    ctx.lineTo(sx0, neck);
    ctx.lineTo(L, neck);
  }
  ctx.closePath();
}

type TextDraft = {
  dispX: number;
  dispY: number;
  cx: number;
  cy: number;
  font: number;
  naturalFont: number;
  value: string;
  // When set, this draft is EDITING an existing text stroke at that index
  // (re-type it) rather than placing a new one.
  editIdx?: number | null;
};

export default function PhotoAnnotator({
  photoUrl,
  reportId,
  photoId,
  onClose,
  onSaved,
}: {
  photoUrl: string;
  reportId: string;
  /** The photo being drawn on — saving replaces it rather than adding one. */
  photoId: string;
  onClose: () => void;
  onSaved?: (newUrl: string) => void;
}) {
  const [drawTool, setDrawTool] = useState<Stroke["tool"] | "move">("arrow");
  const [drawColor, setDrawColor] = useState("#FFD400");
  const [drawWidth, setDrawWidth] = useState(6);
  const [drawFill, setDrawFill] = useState(false);
  const [drawDash, setDrawDash] = useState<"solid" | "dashed" | "dotted">("solid");
  const [strokes, setStrokes] = useState<Stroke[]>([]);
  const [redoStack, setRedoStack] = useState<Stroke[]>([]);
  const [saving, setSaving] = useState(false);
  const [textDraft, setTextDraft] = useState<TextDraft | null>(null);
  const [selectedIdx, setSelectedIdx] = useState<number | null>(null);

  const imgWrapRef = useRef<HTMLDivElement | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const drawingRef = useRef<Stroke | null>(null);
  const baseImgRef = useRef<HTMLImageElement | null>(null);
  const moveDragRef = useRef<{ x: number; y: number } | null>(null);
  // Which endpoint of the selected shape is being dragged to reshape it
  // (point index in stroke.points), or null when moving/idle.
  const resizeHandleRef = useRef<number | null>(null);

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

  // Draws a stroke WITHOUT rotation. The wrapper below applies rotation.
  const drawOneStrokeRaw = (ctx: CanvasRenderingContext2D, s: Stroke) => {
    ctx.strokeStyle = s.color;
    ctx.fillStyle = s.color;
    ctx.lineWidth = s.width;
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
    if (s.dash === "dashed") ctx.setLineDash([Math.max(6, s.width * 3), Math.max(4, s.width * 2)]);
    else if (s.dash === "dotted") ctx.setLineDash([Math.max(1, s.width), Math.max(3, s.width * 2)]);
    else ctx.setLineDash([]);
    const pts = s.points;
    if (!pts.length) return;
    if (s.tool === "text") {
      ctx.setLineDash([]);
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
    if (CLOSED_SHAPES.has(s.tool)) {
      const L = Math.min(a.x, b.x), R = Math.max(a.x, b.x);
      const T = Math.min(a.y, b.y), B = Math.max(a.y, b.y);
      const W = R - L, H = B - T, cx = (L + R) / 2, cy = (T + B) / 2;
      ctx.beginPath();
      if (s.tool === "rect") ctx.rect(L, T, W, H);
      else if (s.tool === "rrect") roundRectPath(ctx, L, T, W, H, Math.min(W, H) * 0.18);
      else if (s.tool === "ellipse") ctx.ellipse(cx, cy, W / 2, H / 2, 0, 0, Math.PI * 2);
      else if (s.tool === "triangle") { ctx.moveTo(cx, T); ctx.lineTo(R, B); ctx.lineTo(L, B); ctx.closePath(); }
      else if (s.tool === "diamond") { ctx.moveTo(cx, T); ctx.lineTo(R, cy); ctx.lineTo(cx, B); ctx.lineTo(L, cy); ctx.closePath(); }
      else if (s.tool === "star") starPath(ctx, cx, cy, Math.min(W, H) / 2, 5);
      else if (s.tool === "pentagon") polygonPath(ctx, cx, cy, W / 2, H / 2, 5);
      else if (s.tool === "hexagon") polygonPath(ctx, cx, cy, W / 2, H / 2, 6);
      else if (s.tool === "callout") calloutPath(ctx, L, T, W, H);
      else if (s.tool === "uparrow") blockArrowPath(ctx, "up", L, T, W, H);
      else if (s.tool === "downarrow") blockArrowPath(ctx, "down", L, T, W, H);
      if (s.fill) {
        ctx.save();
        ctx.globalAlpha = 0.35; // translucent so the photo underneath stays visible
        ctx.fillStyle = s.color;
        ctx.fill();
        ctx.restore();
      }
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

  // Applies per-shape rotation (around the bbox centre), then draws.
  const drawOneStroke = (ctx: CanvasRenderingContext2D, s: Stroke) => {
    const rot = s.rot || 0;
    if (!rot) {
      drawOneStrokeRaw(ctx, s);
      return;
    }
    const b = strokeBBox(s);
    const cx = (b.minX + b.maxX) / 2;
    const cy = (b.minY + b.maxY) / 2;
    ctx.save();
    ctx.translate(cx, cy);
    ctx.rotate(rot);
    ctx.translate(-cx, -cy);
    drawOneStrokeRaw(ctx, s);
    ctx.restore();
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
        const sel = strokes[selectedIdx];
        const b = strokeBBox(sel);
        const dispScale = canvas.width ? (parseFloat(canvas.style.width || "0") / canvas.width) || 1 : 1;
        ctx.save();
        ctx.setLineDash([10 / dispScale, 7 / dispScale]);
        ctx.strokeStyle = "#22D3EE";
        ctx.lineWidth = Math.max(2, 2 / dispScale);
        ctx.strokeRect(b.minX - 8, b.minY - 8, b.maxX - b.minX + 16, b.maxY - b.minY + 16);
        ctx.restore();
        // Draggable endpoint handles for reshaping (2-point shapes only, and
        // only when not rotated — after rotating, resize with − / +).
        if (sel.tool !== "pen" && sel.tool !== "text" && !sel.rot) {
          const hs = Math.max(6, 9 / dispScale);
          const hpts = [sel.points[0], sel.points[sel.points.length - 1]];
          ctx.save();
          ctx.setLineDash([]);
          ctx.fillStyle = "#ffffff";
          ctx.strokeStyle = "#0f172a";
          ctx.lineWidth = Math.max(2, 2 / dispScale);
          for (const hp of hpts) {
            ctx.beginPath();
            ctx.rect(hp.x - hs, hp.y - hs, hs * 2, hs * 2);
            ctx.fill();
            ctx.stroke();
          }
          ctx.restore();
        }
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [strokes, drawTool, selectedIdx]
  );

  // Load the photo into the canvas on mount. External (S3) URLs go through the
  // same-origin (login-only) proxy so the canvas stays exportable.
  useEffect(() => {
    if (!photoUrl) return;
    let cancelled = false;
    let revoke = () => {};
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
    loadCanvasSafeImage(photoUrl)
      .then((r) => {
        revoke = r.revoke;
        if (cancelled) return revoke();
        img.src = r.src;
      })
      .catch(() => {
        if (!cancelled) baseImgRef.current = null;
      });
    return () => {
      cancelled = true;
      revoke();
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
      const s = strokes[i];
      const b = strokeBBox(s);
      let q = p;
      if (s.rot) {
        // Un-rotate the click into the shape's local frame before the bbox test.
        const cx = (b.minX + b.maxX) / 2;
        const cy = (b.minY + b.maxY) / 2;
        const cos = Math.cos(-s.rot);
        const sin = Math.sin(-s.rot);
        const dx = p.x - cx;
        const dy = p.y - cy;
        q = { x: cx + dx * cos - dy * sin, y: cy + dx * sin + dy * cos };
      }
      if (q.x >= b.minX - pad && q.x <= b.maxX + pad && q.y >= b.minY - pad && q.y <= b.maxY + pad) return i;
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
      setStrokes(strokesWithDraft(textDraft));
      setTextDraft(null);
      if (drawTool !== "text") return;
    }
    const p = canvasPoint(e);
    if (drawTool === "move") {
      (e.target as HTMLCanvasElement).setPointerCapture(e.pointerId);
      // If a shape is already selected, grabbing one of its white endpoint
      // handles reshapes it (drag the corner) instead of moving the whole shape.
      if (selectedIdx != null) {
        const sel = strokes[selectedIdx];
        if (sel && sel.tool !== "pen" && sel.tool !== "text" && !sel.rot) {
          const canvas = canvasRef.current!;
          const rect = canvas.getBoundingClientRect();
          const scale = rect.width / canvas.width || 1;
          const tol = 18 / scale;
          const iLast = sel.points.length - 1;
          if (Math.hypot(p.x - sel.points[0].x, p.y - sel.points[0].y) <= tol) {
            resizeHandleRef.current = 0;
            moveDragRef.current = null;
            return;
          }
          if (Math.hypot(p.x - sel.points[iLast].x, p.y - sel.points[iLast].y) <= tol) {
            resizeHandleRef.current = iLast;
            moveDragRef.current = null;
            return;
          }
        }
      }
      const idx = hitTest(p);
      setSelectedIdx(idx);
      moveDragRef.current = idx != null ? { x: p.x, y: p.y } : null;
      resizeHandleRef.current = null;
      redrawCanvas();
      return;
    }
    if (drawTool === "text") {
      const canvas = canvasRef.current;
      const wrap = imgWrapRef.current;
      if (!canvas || !wrap) return;
      const wrapRect = wrap.getBoundingClientRect();
      const canvasRect = canvas.getBoundingClientRect();
      const scale = canvas.width ? canvasRect.width / canvas.width : 1;
      // Tapping an existing text label re-opens it for editing (re-type it),
      // keeping its position — instead of only being able to place new text.
      const hitIdx = hitTest(p);
      const hitStroke = hitIdx != null ? strokes[hitIdx] : null;
      if (hitStroke && hitStroke.tool === "text") {
        const nf = hitStroke.fontSize || 32;
        setSelectedIdx(hitIdx);
        setTextDraft({
          dispX: e.clientX - wrapRect.left,
          dispY: e.clientY - wrapRect.top,
          cx: hitStroke.points[0].x,
          cy: hitStroke.points[0].y,
          naturalFont: nf,
          font: nf * scale,
          value: hitStroke.text || "",
          editIdx: hitIdx,
        });
        return;
      }
      const naturalFont = Math.max(28, drawWidth * 7);
      setTextDraft({
        dispX: e.clientX - wrapRect.left,
        dispY: e.clientY - wrapRect.top,
        cx: p.x,
        cy: p.y,
        naturalFont,
        font: naturalFont * scale,
        value: "",
        editIdx: null,
      });
      return;
    }
    (e.target as HTMLCanvasElement).setPointerCapture(e.pointerId);
    drawingRef.current = {
      tool: drawTool as Stroke["tool"],
      color: drawColor,
      width: drawWidth,
      points: [p, p],
      fill: drawFill && CLOSED_SHAPES.has(drawTool),
      dash: drawDash,
    };
    redrawCanvas(drawingRef.current);
  };
  const onDrawMove = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const p = canvasPoint(e);
    if (drawTool === "move") {
      // Reshaping: drag one endpoint handle.
      if (resizeHandleRef.current != null && selectedIdx != null) {
        const hi = resizeHandleRef.current;
        setStrokes((prev) =>
          prev.map((s, i) =>
            i === selectedIdx ? { ...s, points: s.points.map((pt, pi) => (pi === hi ? { x: p.x, y: p.y } : pt)) } : s
          )
        );
        return;
      }
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
    resizeHandleRef.current = null;
    if (!drawingRef.current) return;
    const s = drawingRef.current;
    drawingRef.current = null;
    setStrokes((prev) => [...prev, s]);
    setRedoStack([]); // a new drawing invalidates the redo history
  };

  // ---- Undo / Redo (add-level history, like Word's basic undo stack) ----
  const undo = () => {
    if (!strokes.length) return;
    setRedoStack((r) => [...r, strokes[strokes.length - 1]]);
    setStrokes(strokes.slice(0, -1));
    setSelectedIdx(null);
  };
  const redo = () => {
    if (!redoStack.length) return;
    setStrokes([...strokes, redoStack[redoStack.length - 1]]);
    setRedoStack(redoStack.slice(0, -1));
  };

  // ---- Selected-shape editing: rotate, duplicate, layer order ----
  const rotateSelected = (deltaDeg: number) => {
    if (selectedIdx == null) return;
    setStrokes((prev) => prev.map((s, i) => (i === selectedIdx ? { ...s, rot: (s.rot || 0) + (deltaDeg * Math.PI) / 180 } : s)));
  };
  const duplicateSelected = () => {
    if (selectedIdx == null) return;
    const s = strokes[selectedIdx];
    const clone: Stroke = { ...s, points: s.points.map((p) => ({ x: p.x + 24, y: p.y + 24 })) };
    setStrokes((prev) => [...prev, clone]);
    setSelectedIdx(strokes.length); // select the new clone
    setRedoStack([]);
  };
  const bringToFront = () => {
    if (selectedIdx == null) return;
    setStrokes((prev) => {
      const copy = prev.slice();
      const [s] = copy.splice(selectedIdx, 1);
      copy.push(s);
      return copy;
    });
    setSelectedIdx(strokes.length - 1);
  };
  const sendToBack = () => {
    if (selectedIdx == null) return;
    setStrokes((prev) => {
      const copy = prev.slice();
      const [s] = copy.splice(selectedIdx, 1);
      copy.unshift(s);
      return copy;
    });
    setSelectedIdx(0);
  };

  // Apply a text draft to the strokes: EDIT the existing stroke in place when
  // editIdx is set (empty text removes it), otherwise append a new text stroke.
  const strokesWithDraft = (d: TextDraft | null): Stroke[] => {
    if (!d) return strokes;
    const val = d.value.trim();
    if (d.editIdx != null) {
      if (!val) return strokes.filter((_, i) => i !== d.editIdx);
      return strokes.map((s, i) => (i === d.editIdx ? { ...s, text: val } : s));
    }
    return val
      ? [
          ...strokes,
          { tool: "text", color: drawColor, width: drawWidth, points: [{ x: d.cx, y: d.cy }], text: val, fontSize: d.naturalFont },
        ]
      : strokes;
  };

  const commitText = () => {
    setStrokes(strokesWithDraft(textDraft));
    setTextDraft(null);
  };

  const saveAnnotated = async () => {
    const canvas = canvasRef.current;
    if (!canvas || saving) return;
    setSaving(true);
    try {
      if (textDraft) {
        // Bake a pending text edit/placement into the image before exporting.
        // Redraw base + all strokes (with the draft applied) so an EDIT replaces
        // the old text instead of drawing on top of it, and no selection box
        // leaks into the saved image.
        const finalStrokes = strokesWithDraft(textDraft);
        const ctx = canvas.getContext("2d");
        if (ctx) {
          ctx.clearRect(0, 0, canvas.width, canvas.height);
          const base = baseImgRef.current;
          if (base) ctx.drawImage(base, 0, 0, canvas.width, canvas.height);
          for (const s of finalStrokes) drawOneStroke(ctx, s);
        }
        setStrokes(finalStrokes);
        setTextDraft(null);
      }
      const blob: Blob | null = await new Promise((resolve) => canvas.toBlob((b) => resolve(b), "image/jpeg", 0.92));
      if (!blob) throw new Error("Couldn't export the annotated image — the photo host is blocking cross-origin canvas export.");
      const newUrl = await replacePhotoImage({
        reportId,
        photoId,
        blob,
        fileName: `annotated_${Date.now()}.jpg`,
        width: canvas.width,
        height: canvas.height,
      });
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
            ["rrect", "▢", "Rounded rectangle"],
            ["ellipse", "◯", "Ellipse / circle"],
            ["triangle", "△", "Triangle"],
            ["diamond", "◇", "Diamond"],
            ["star", "★", "Star"],
            ["pentagon", "⬠", "Pentagon"],
            ["hexagon", "⬡", "Hexagon"],
            ["callout", "💬", "Callout / speech bubble"],
            ["uparrow", "⬆", "Up block arrow"],
            ["downarrow", "⬇", "Down block arrow"],
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
          <select
            value={drawDash}
            onChange={(e) => setDrawDash(e.target.value as "solid" | "dashed" | "dotted")}
            style={S.widthSelect}
            title="Line style"
          >
            <option value="solid">Solid</option>
            <option value="dashed">Dashed</option>
            <option value="dotted">Dotted</option>
          </select>
          <label
            style={{ display: "inline-flex", alignItems: "center", gap: 5, fontSize: 12, fontWeight: 800, color: "#0f172a", cursor: "pointer", padding: "0 4px" }}
            title="Fill closed shapes (translucent, so the photo stays visible)"
          >
            <input type="checkbox" checked={drawFill} onChange={(e) => setDrawFill(e.target.checked)} style={{ width: 16, height: 16, cursor: "pointer" }} />
            Fill
          </label>
        </div>

        {drawTool === "move" ? (
          <div style={{ fontSize: 11, fontWeight: 700, color: "#64748b" }}>
            {selectedIdx == null
              ? "Tap any drawing or text to select it."
              : "Drag inside to move · drag a white corner to reshape · − / + to resize · Delete to remove."}
          </div>
        ) : null}
        {drawTool === "text" ? (
          <div style={{ fontSize: 11, fontWeight: 700, color: "#64748b" }}>
            Tap to place a label · tap an existing label to re-edit it · Enter to confirm.
          </div>
        ) : null}
        {drawTool === "move" && selectedIdx != null ? (
          <div style={S.actions}>
            <button style={S.ghost} onClick={() => rotateSelected(-15)} title="Rotate left 15°">⟲ Rotate</button>
            <button style={S.ghost} onClick={() => rotateSelected(15)} title="Rotate right 15°">Rotate ⟳</button>
            <button style={S.ghost} onClick={() => resizeSelected(0.85)}>− Smaller</button>
            <button style={S.ghost} onClick={() => resizeSelected(1.18)}>Larger +</button>
            <button style={S.ghost} onClick={duplicateSelected}>Duplicate</button>
            <button style={S.ghost} onClick={bringToFront} title="Bring to front">Front</button>
            <button style={S.ghost} onClick={sendToBack} title="Send to back">Back</button>
            <button style={S.ghost} onClick={deleteSelected}>Delete</button>
          </div>
        ) : null}

        <div style={S.actions}>
          <button style={S.ghost} onClick={undo} disabled={!total}>↶ Undo</button>
          <button style={S.ghost} onClick={redo} disabled={!redoStack.length}>↷ Redo</button>
          <button
            style={S.ghost}
            onClick={() => { setStrokes([]); setRedoStack([]); setSelectedIdx(null); }}
            disabled={!total}
          >
            Clear
          </button>
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
