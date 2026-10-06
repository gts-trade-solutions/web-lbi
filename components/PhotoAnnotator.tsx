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
import { confirmDialog } from "./ConfirmDialog";

type Stroke = {
  tool:
    | "arrow" | "darrow" | "carrow" | "line" | "pen"
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

type Pt = { x: number; y: number };

// Perpendicular distance from p to the line through a–b.
function perpDist(p: Pt, a: Pt, b: Pt): number {
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  const len2 = dx * dx + dy * dy;
  if (len2 === 0) return Math.hypot(p.x - a.x, p.y - a.y);
  const t = ((p.x - a.x) * dx + (p.y - a.y) * dy) / len2;
  const cx = a.x + t * dx;
  const cy = a.y + t * dy;
  return Math.hypot(p.x - cx, p.y - cy);
}

// Ramer–Douglas–Peucker: drop points that don't change the shape, which removes
// the hand's jitter and leaves only the real turns of the gesture.
function simplifyRDP(pts: Pt[], tol: number): Pt[] {
  if (pts.length < 3) return pts.slice();
  const a = pts[0];
  const b = pts[pts.length - 1];
  let maxD = 0;
  let idx = 0;
  for (let i = 1; i < pts.length - 1; i++) {
    const d = perpDist(pts[i], a, b);
    if (d > maxD) {
      maxD = d;
      idx = i;
    }
  }
  if (maxD > tol) {
    const left = simplifyRDP(pts.slice(0, idx + 1), tol);
    const right = simplifyRDP(pts.slice(idx), tol);
    return left.slice(0, -1).concat(right);
  }
  return [a, b];
}

// Clean up a shaky free-hand path so it reads like a Word vector line rather
// than a pencil scribble. Chaikin corner-cutting rounds every corner a few
// times; the endpoints are kept so the line still starts/ends where drawn.
function chaikinSmooth(pts: { x: number; y: number }[], iterations = 2): { x: number; y: number }[] {
  let out = pts;
  for (let it = 0; it < iterations; it++) {
    if (out.length < 3) break;
    const next: { x: number; y: number }[] = [out[0]];
    for (let i = 0; i < out.length - 1; i++) {
      const p0 = out[i];
      const p1 = out[i + 1];
      next.push({ x: p0.x * 0.75 + p1.x * 0.25, y: p0.y * 0.75 + p1.y * 0.25 });
      next.push({ x: p0.x * 0.25 + p1.x * 0.75, y: p0.y * 0.25 + p1.y * 0.75 });
    }
    next.push(out[out.length - 1]);
    out = next;
  }
  return out;
}

// Word's Shift-constrain while drawing: straight lines snap to 0/45/90°, and
// boxes/circles become square. `a` is the start point, `b` the current point.
function constrainEnd(a: Pt, b: Pt, tool: string): Pt {
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  if (tool === "line" || tool === "arrow" || tool === "darrow") {
    const step = Math.PI / 4;
    const snap = Math.round(Math.atan2(dy, dx) / step) * step;
    const len = Math.hypot(dx, dy);
    return { x: a.x + len * Math.cos(snap), y: a.y + len * Math.sin(snap) };
  }
  const s = Math.max(Math.abs(dx), Math.abs(dy));
  return { x: a.x + (dx < 0 ? -s : s), y: a.y + (dy < 0 ? -s : s) };
}

// Rotate point p by `ang` radians around centre c.
function rotateAround(p: Pt, ang: number, c: Pt): Pt {
  const cos = Math.cos(ang);
  const sin = Math.sin(ang);
  const dx = p.x - c.x;
  const dy = p.y - c.y;
  return { x: c.x + dx * cos - dy * sin, y: c.y + dx * sin + dy * cos };
}

// What a drag of the selection box is doing (Word-style move / resize / rotate).
type BoxDrag =
  | { kind: "move" }
  | { kind: "rotate"; c: Pt; startAngle: number; startRot: number }
  | {
      kind: "resize";
      handle: number; // 0..7, clockwise from top-left
      pts0: Pt[];
      fontSize0: number;
      rot: number;
      c0: Pt;
      anchorLocal: Pt; // opposite handle, fixed during the drag
      draggedLocal: Pt; // the handle being dragged
    };

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

function annoAuthHeaders(): Record<string, string> {
  const t = typeof window !== "undefined" ? localStorage.getItem("auth_token") : null;
  return t ? { Authorization: `Bearer ${t}` } : {};
}

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
  const [drawColor, setDrawColor] = useState("#FFFF00"); // Word's standard yellow
  const [drawWidth, setDrawWidth] = useState(6);
  const [drawFill, setDrawFill] = useState(false);
  const [drawDash, setDrawDash] = useState<"solid" | "dashed" | "dotted">("solid");
  const [textSize, setTextSize] = useState(40); // font size (px) for new text labels
  const [strokes, setStrokes] = useState<Stroke[]>([]);
  // True once this photo loaded with a SAVED drawing — so after "Clear all" the
  // Save button stays enabled and saving removes the drawing (clean photo).
  const [hadDrawing, setHadDrawing] = useState(false);
  const [redoStack, setRedoStack] = useState<Stroke[]>([]);
  const [saving, setSaving] = useState(false);
  const [textDraft, setTextDraft] = useState<TextDraft | null>(null);
  const [selectedIdx, setSelectedIdx] = useState<number | null>(null);

  const imgWrapRef = useRef<HTMLDivElement | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const drawingRef = useRef<Stroke | null>(null);
  const baseImgRef = useRef<HTMLImageElement | null>(null);
  // The URL the strokes are drawn OVER. On a first drawing this is the photo
  // itself; after a save it's the preserved original (anno_base_url), so
  // reopening edits the strokes over the clean base instead of a baked image.
  const baseUrlRef = useRef<string>(photoUrl);
  const moveDragRef = useRef<{ x: number; y: number } | null>(null);
  // Which endpoint of the selected shape is being dragged to reshape it
  // (point index in stroke.points), or null when moving/idle.
  const resizeHandleRef = useRef<number | null>(null);
  // Active Word-style selection-box drag (move / resize via a handle / rotate).
  const boxDragRef = useRef<BoxDrag | null>(null);
  // Dragging the text label while it's being typed (via its move grip).
  const textDragRef = useRef<
    { sx: number; sy: number; dX: number; dY: number; cx: number; cy: number; scale: number } | null
  >(null);

  // Advanced shapes (rectangles, stars, turn arrows…) are hidden behind "More"
  // so the everyday toolbar stays simple — the #1 "too many buttons" complaint.
  const [showMore, setShowMore] = useState(false);

  // Size the photo as large as the window allows (minus room for the toolbar),
  // so there is a big area to draw on — and keep it right on window resize /
  // device rotation. Never upscales past 1:1 (keeps small photos crisp).
  const fitCanvas = useCallback(() => {
    const canvas = canvasRef.current;
    const img = baseImgRef.current;
    if (!canvas) return;
    const w = img?.naturalWidth || canvas.width || 1200;
    const h = img?.naturalHeight || canvas.height || 800;
    // Size from the window (reliable at any time), leaving room for the toolbar.
    const maxW = Math.max(260, window.innerWidth - 24);
    const maxH = Math.max(260, window.innerHeight - 220);
    const scale = Math.min(maxW / w, maxH / h, 1);
    canvas.style.width = `${Math.round(w * scale)}px`;
    canvas.style.height = `${Math.round(h * scale)}px`;
  }, []);

  useEffect(() => {
    const onResize = () => fitCanvas();
    window.addEventListener("resize", onResize);
    window.addEventListener("orientationchange", onResize);
    return () => {
      window.removeEventListener("resize", onResize);
      window.removeEventListener("orientationchange", onResize);
    };
  }, [fitCanvas]);

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
    if (s.tool === "pen" || s.tool === "carrow") {
      // Clean, Word-like free-hand: strongly smooth the shaky path (Chaikin),
      // then draw a quadratic curve through it. Used on screen AND in the
      // exported image, so saved drawings are clean too. "carrow" adds a filled
      // arrowhead at the end — a Word-style curved arrow.
      // Remove the hand's jitter (RDP) first, then round what's left (Chaikin):
      // a shaky stroke becomes a clean, flowing curve that follows the gesture.
      const simp = simplifyRDP(pts, Math.max(5, s.width * 1.2));
      const sm = chaikinSmooth(simp, 4);
      ctx.beginPath();
      ctx.moveTo(sm[0].x, sm[0].y);
      if (sm.length < 3) {
        for (let i = 1; i < sm.length; i++) ctx.lineTo(sm[i].x, sm[i].y);
      } else {
        for (let i = 1; i < sm.length - 1; i++) {
          const mx = (sm[i].x + sm[i + 1].x) / 2;
          const my = (sm[i].y + sm[i + 1].y) / 2;
          ctx.quadraticCurveTo(sm[i].x, sm[i].y, mx, my);
        }
        ctx.quadraticCurveTo(
          sm[sm.length - 2].x,
          sm[sm.length - 2].y,
          sm[sm.length - 1].x,
          sm[sm.length - 1].y
        );
      }
      ctx.stroke();
      if (s.tool === "carrow" && sm.length >= 2) {
        // Arrowhead pointing along the final direction of the curve.
        const tip = sm[sm.length - 1];
        let back = sm[sm.length - 2];
        const minBack = s.width * 2 + 6;
        for (let i = sm.length - 2; i >= 0; i--) {
          if (Math.hypot(tip.x - sm[i].x, tip.y - sm[i].y) >= minBack) {
            back = sm[i];
            break;
          }
        }
        const ang = Math.atan2(tip.y - back.y, tip.x - back.x);
        const head = Math.max(14, s.width * 3.5);
        ctx.setLineDash([]);
        ctx.beginPath();
        ctx.moveTo(tip.x, tip.y);
        ctx.lineTo(tip.x - head * Math.cos(ang - Math.PI / 7), tip.y - head * Math.sin(ang - Math.PI / 7));
        ctx.lineTo(tip.x - head * Math.cos(ang + Math.PI / 7), tip.y - head * Math.sin(ang + Math.PI / 7));
        ctx.closePath();
        ctx.fill();
      }
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

  // Word-style selection geometry: the 8 resize handles (clockwise from
  // top-left), the rotation handle above the shape, the box corners and the
  // centre — all in canvas coords with the shape's rotation applied.
  const handleGeom = (s: Stroke) => {
    const b = strokeBBox(s);
    const pad = 6;
    const minX = b.minX - pad;
    const minY = b.minY - pad;
    const maxX = b.maxX + pad;
    const maxY = b.maxY + pad;
    const c = { x: (minX + maxX) / 2, y: (minY + maxY) / 2 };
    const rot = s.rot || 0;
    // local (un-rotated) positions
    const local = [
      { x: minX, y: minY }, // 0 TL
      { x: c.x, y: minY }, // 1 T
      { x: maxX, y: minY }, // 2 TR
      { x: maxX, y: c.y }, // 3 R
      { x: maxX, y: maxY }, // 4 BR
      { x: c.x, y: maxY }, // 5 B
      { x: minX, y: maxY }, // 6 BL
      { x: minX, y: c.y }, // 7 L
    ];
    const handles = local.map((h) => rotateAround(h, rot, c));
    const corners = [handles[0], handles[2], handles[4], handles[6]];
    const rotLocal = { x: c.x, y: minY - Math.max(26, (maxY - minY) * 0.14) };
    const rotHandle = rotateAround(rotLocal, rot, c);
    return { b: { minX, minY, maxX, maxY }, c, rot, local, handles, corners, rotHandle };
  };

  // Is canvas point p inside the selected shape's (rotated) box?
  const insideBox = (s: Stroke, p: Pt) => {
    const g = handleGeom(s);
    const lp = rotateAround(p, -g.rot, g.c);
    return lp.x >= g.b.minX && lp.x <= g.b.maxX && lp.y >= g.b.minY && lp.y <= g.b.maxY;
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
      // Hide the text stroke currently being edited so its live editor input
      // isn't doubled up by the committed label underneath it.
      const editingIdx = textDraft?.editIdx ?? null;
      for (let i = 0; i < strokes.length; i++) {
        if (editingIdx === i) continue;
        drawOneStroke(ctx, strokes[i]);
      }
      if (extra) drawOneStroke(ctx, extra);
      if (drawTool === "move" && selectedIdx != null && strokes[selectedIdx]) {
        const sel = strokes[selectedIdx];
        const g = handleGeom(sel);
        const dispScale = canvas.width ? (parseFloat(canvas.style.width || "0") / canvas.width) || 1 : 1;
        const lw = Math.max(1.5, 1.5 / dispScale);
        ctx.save();
        // Selection box (follows rotation).
        ctx.setLineDash([]);
        ctx.strokeStyle = "#2b6cff";
        ctx.lineWidth = lw;
        ctx.beginPath();
        ctx.moveTo(g.corners[0].x, g.corners[0].y);
        for (let k = 1; k < g.corners.length; k++) ctx.lineTo(g.corners[k].x, g.corners[k].y);
        ctx.closePath();
        ctx.stroke();
        // Line up to the rotation handle.
        ctx.beginPath();
        ctx.moveTo(g.handles[1].x, g.handles[1].y);
        ctx.lineTo(g.rotHandle.x, g.rotHandle.y);
        ctx.stroke();
        // Handles — white circles with a blue ring, like Word.
        const hr = Math.max(5, 7 / dispScale);
        const drawHandle = (pt: Pt, r: number) => {
          ctx.beginPath();
          ctx.arc(pt.x, pt.y, r, 0, Math.PI * 2);
          ctx.fillStyle = "#ffffff";
          ctx.fill();
          ctx.lineWidth = lw;
          ctx.strokeStyle = "#2b6cff";
          ctx.stroke();
        };
        // Line/arrow-type shapes (2 points) only need the two end handles;
        // closed shapes & text get the full 8-handle box.
        const twoPoint =
          sel.tool === "arrow" ||
          sel.tool === "darrow" ||
          sel.tool === "line";
        if (twoPoint) {
          drawHandle(rotateAround(sel.points[0], g.rot, g.c), hr);
          drawHandle(rotateAround(sel.points[sel.points.length - 1], g.rot, g.c), hr);
        } else if (sel.tool === "pen" || sel.tool === "carrow") {
          // free-hand: just the box (dragging handles would distort it oddly)
        } else {
          for (const h of g.handles) drawHandle(h, hr);
        }
        drawHandle(g.rotHandle, hr * 1.1);
        ctx.restore();
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [strokes, drawTool, selectedIdx, textDraft]
  );

  // Load the photo into the canvas on mount. If this photo carries a SAVED
  // drawing (anno_json + anno_base_url from a previous "Draw on photo"), restore
  // the strokes and draw them over the PRESERVED base image so they can be
  // edited again — instead of drawing on top of the already-flattened picture.
  // External (S3) URLs go through the same-origin (login-only) proxy so the
  // canvas stays exportable.
  useEffect(() => {
    if (!photoUrl) return;
    let cancelled = false;
    let revoke = () => {};

    const loadBase = (url: string, parsedStrokes: Stroke[] | null) => {
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
          fitCanvas();
        }
        // Base image is ready. Draw it, THEN apply the saved strokes — so the
        // strokes re-render happens with the image already loaded. Setting the
        // strokes BEFORE the image finished loading raced with this onload and
        // the base-only redraw here wiped the drawing (it "disappeared").
        redrawCanvas();
        if (parsedStrokes && parsedStrokes.length) {
          setStrokes(parsedStrokes);
          setHadDrawing(true);
        }
      };
      img.onerror = () => {
        if (!cancelled) baseImgRef.current = null;
      };
      loadCanvasSafeImage(url)
        .then((r) => {
          revoke = r.revoke;
          if (cancelled) return revoke();
          img.src = r.src;
        })
        .catch(() => {
          if (!cancelled) baseImgRef.current = null;
        });
    };

    (async () => {
      let baseUrl = photoUrl;
      let parsedStrokes: Stroke[] | null = null;
      try {
        const res = await fetch(`/api/reports/${encodeURIComponent(reportId)}/photos`, {
          headers: annoAuthHeaders(),
          credentials: "include",
        });
        if (res.ok) {
          const data = await res.json().catch(() => null);
          const row =
            data && Array.isArray(data.photos)
              ? data.photos.find((p: any) => String(p?.id) === String(photoId))
              : null;
          const rawBase = row?.anno_base_url;
          const rawAnno = row?.anno_json;
          if (rawBase && typeof rawBase === "string") baseUrl = rawBase;
          if (rawAnno && typeof rawAnno === "string") {
            try {
              const parsed = JSON.parse(rawAnno);
              if (Array.isArray(parsed)) parsedStrokes = parsed;
            } catch {
              /* corrupt anno — ignore, start clean */
            }
          }
        }
      } catch {
        /* no saved strokes / offline — fall back to the plain photo */
      }
      if (cancelled) return;
      baseUrlRef.current = baseUrl;
      loadBase(baseUrl, parsedStrokes);
    })();

    return () => {
      cancelled = true;
      revoke();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [photoUrl, reportId, photoId]);

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

  // Pick a colour: set it for new shapes AND recolour the shape currently being
  // edited (a text being retyped) or selected (Move tool), so "change the
  // colour of this arrow / text" works, not just the next one you draw.
  const pickColor = (c: string) => {
    setDrawColor(c);
    const editIdx = textDraft?.editIdx ?? null;
    const targetIdx =
      editIdx != null ? editIdx : drawTool === "move" ? selectedIdx : null;
    if (targetIdx != null) {
      setStrokes((prev) =>
        prev.map((s, i) => (i === targetIdx ? { ...s, color: c } : s))
      );
    }
  };

  const canvasDispScale = () => {
    const c = canvasRef.current;
    return c && c.width ? (parseFloat(c.style.width || "0") / c.width) || 1 : 1;
  };

  // Drag the text label (by its grip) while it's being typed — move it without
  // leaving the keyboard / clicking a separate tool.
  const startTextDrag = (e: React.PointerEvent) => {
    if (!textDraft) return;
    e.preventDefault();
    try {
      (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
    } catch {
      /* best effort */
    }
    textDragRef.current = {
      sx: e.clientX,
      sy: e.clientY,
      dX: textDraft.dispX,
      dY: textDraft.dispY,
      cx: textDraft.cx,
      cy: textDraft.cy,
      scale: canvasDispScale(),
    };
  };
  const onTextDragMove = (e: React.PointerEvent) => {
    const d = textDragRef.current;
    if (!d) return;
    const ddx = e.clientX - d.sx;
    const ddy = e.clientY - d.sy;
    setTextDraft((t) =>
      t ? { ...t, dispX: d.dX + ddx, dispY: d.dY + ddy, cx: d.cx + ddx / d.scale, cy: d.cy + ddy / d.scale } : t
    );
  };
  const endTextDrag = (e: React.PointerEvent) => {
    textDragRef.current = null;
    try {
      (e.currentTarget as HTMLElement).releasePointerCapture(e.pointerId);
    } catch {
      /* ignore */
    }
  };

  // Make text bigger / smaller. Applies to the label being typed, the selected
  // text label, and the default size for the next new label.
  const changeTextSize = (factor: number) => {
    const clamp = (v: number) => Math.max(10, Math.min(240, Math.round(v * factor)));
    const canvas = canvasRef.current;
    const scale = canvas && canvas.width ? (parseFloat(canvas.style.width || "0") / canvas.width) || 1 : 1;
    if (textDraft) {
      const newNat = clamp(textDraft.naturalFont);
      setTextDraft((d) => (d ? { ...d, naturalFont: newNat, font: newNat * scale } : d));
      setTextSize(newNat);
      if (textDraft.editIdx != null) {
        const ei = textDraft.editIdx;
        setStrokes((prev) => prev.map((s, i) => (i === ei ? { ...s, fontSize: newNat } : s)));
      }
      return;
    }
    if (drawTool === "move" && selectedIdx != null && strokes[selectedIdx]?.tool === "text") {
      setStrokes((prev) =>
        prev.map((s, i) => (i === selectedIdx ? { ...s, fontSize: clamp(s.fontSize || textSize) } : s))
      );
      setTextSize((v) => clamp(v));
      return;
    }
    setTextSize((v) => clamp(v));
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
      boxDragRef.current = null;
      resizeHandleRef.current = null;
      moveDragRef.current = null;
      // If a shape is already selected, a tap on one of its handles / the
      // rotation handle / inside the box drives resize / rotate / move — exactly
      // like grabbing a shape in Word.
      if (selectedIdx != null && strokes[selectedIdx]) {
        const sel = strokes[selectedIdx];
        const canvas = canvasRef.current!;
        const rect = canvas.getBoundingClientRect();
        const scale = rect.width / canvas.width || 1;
        const tol = 16 / scale; // finger/mouse hit radius for a handle
        const g = handleGeom(sel);

        // Rotation handle.
        if (Math.hypot(p.x - g.rotHandle.x, p.y - g.rotHandle.y) <= tol) {
          boxDragRef.current = {
            kind: "rotate",
            c: g.c,
            startAngle: Math.atan2(p.y - g.c.y, p.x - g.c.x),
            startRot: g.rot,
          };
          return;
        }

        const twoPoint = sel.tool === "arrow" || sel.tool === "darrow" || sel.tool === "line";
        const freeHand = sel.tool === "pen" || sel.tool === "carrow";

        if (twoPoint) {
          // Line/arrow: the two handles ARE the endpoints — drag to re-aim.
          const ends = [
            { i: 0, pt: rotateAround(sel.points[0], g.rot, g.c) },
            { i: sel.points.length - 1, pt: rotateAround(sel.points[sel.points.length - 1], g.rot, g.c) },
          ];
          for (const en of ends) {
            if (Math.hypot(p.x - en.pt.x, p.y - en.pt.y) <= tol) {
              resizeHandleRef.current = en.i;
              return;
            }
          }
        } else if (!freeHand) {
          // Closed shapes & text: 8-handle box resize.
          for (let hi = 0; hi < 8; hi++) {
            if (Math.hypot(p.x - g.handles[hi].x, p.y - g.handles[hi].y) <= tol) {
              const anchorIdx = (hi + 4) % 8;
              boxDragRef.current = {
                kind: "resize",
                handle: hi,
                pts0: sel.points.map((pt) => ({ ...pt })),
                fontSize0: sel.fontSize || 32,
                rot: g.rot,
                c0: g.c,
                anchorLocal: g.local[anchorIdx],
                draggedLocal: g.local[hi],
              };
              return;
            }
          }
        }

        // Inside the (rotated) box → move the whole shape.
        if (insideBox(sel, p)) {
          boxDragRef.current = { kind: "move" };
          moveDragRef.current = { x: p.x, y: p.y };
          return;
        }
      }

      // Otherwise pick whatever is under the pointer (or clear the selection).
      const idx = hitTest(p);
      setSelectedIdx(idx);
      boxDragRef.current = idx != null ? { kind: "move" } : null;
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
      const scale = canvas.width ? canvasRect.width / canvas.width : 1;
      // Tapping an existing text label re-opens it for editing (re-type it),
      // keeping its position — instead of only being able to place new text.
      const hitIdx = hitTest(p);
      const hitStroke = hitIdx != null ? strokes[hitIdx] : null;
      if (hitStroke && hitStroke.tool === "text") {
        const nf = hitStroke.fontSize || 32;
        setSelectedIdx(hitIdx);
        // Reflect the text's current colour in the picker so it shows the right
        // colour while editing, and changing it recolours this text.
        if (hitStroke.color) setDrawColor(hitStroke.color);
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
      const naturalFont = textSize;
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
      if (selectedIdx == null) return;
      const drag = boxDragRef.current;

      // Rotate via the rotation handle.
      if (drag?.kind === "rotate") {
        const ang = Math.atan2(p.y - drag.c.y, p.x - drag.c.x);
        let newRot = drag.startRot + (ang - drag.startAngle);
        // Hold Shift to snap rotation to 15° steps, like Word.
        if (e.shiftKey) {
          const step = Math.PI / 12;
          newRot = Math.round(newRot / step) * step;
        }
        setStrokes((prev) => prev.map((s, i) => (i === selectedIdx ? { ...s, rot: newRot } : s)));
        return;
      }

      // Resize via a box handle (keeps the opposite handle fixed).
      if (drag?.kind === "resize") {
        const lp = rotateAround(p, -drag.rot, drag.c0); // pointer in the shape's local frame
        const corner = drag.handle % 2 === 0;
        const affectsX = corner || drag.handle === 3 || drag.handle === 7;
        const affectsY = corner || drag.handle === 1 || drag.handle === 5;
        let sx = 1;
        let sy = 1;
        if (affectsX) {
          const denom = drag.draggedLocal.x - drag.anchorLocal.x;
          if (Math.abs(denom) > 0.5) sx = (lp.x - drag.anchorLocal.x) / denom;
        }
        if (affectsY) {
          const denom = drag.draggedLocal.y - drag.anchorLocal.y;
          if (Math.abs(denom) > 0.5) sy = (lp.y - drag.anchorLocal.y) / denom;
        }
        const clamp = (v: number) => (v >= 0 ? Math.max(0.05, v) : Math.min(-0.05, v));
        sx = clamp(sx);
        sy = clamp(sy);
        const a = drag.anchorLocal;
        if (strokes[selectedIdx]?.tool === "text") {
          const f = corner ? Math.max(Math.abs(sx), Math.abs(sy)) : affectsX ? Math.abs(sx) : Math.abs(sy);
          const newFont = Math.max(10, Math.round(drag.fontSize0 * f));
          const p0 = drag.pts0[0];
          const np0 = { x: a.x + (p0.x - a.x) * sx, y: a.y + (p0.y - a.y) * sy };
          setStrokes((prev) => prev.map((s, i) => (i === selectedIdx ? { ...s, fontSize: newFont, points: [np0] } : s)));
        } else {
          const np = drag.pts0.map((pt) => ({ x: a.x + (pt.x - a.x) * sx, y: a.y + (pt.y - a.y) * sy }));
          setStrokes((prev) => prev.map((s, i) => (i === selectedIdx ? { ...s, points: np } : s)));
        }
        return;
      }

      // Reshaping a line/arrow by dragging one endpoint.
      if (resizeHandleRef.current != null) {
        const hi = resizeHandleRef.current;
        setStrokes((prev) =>
          prev.map((s, i) =>
            i === selectedIdx ? { ...s, points: s.points.map((pt, pi) => (pi === hi ? { x: p.x, y: p.y } : pt)) } : s
          )
        );
        return;
      }

      // Move the whole shape.
      if (!moveDragRef.current) return;
      const dx = p.x - moveDragRef.current.x;
      const dy = p.y - moveDragRef.current.y;
      moveDragRef.current = { x: p.x, y: p.y };
      setStrokes((prev) => prev.map((s, i) => (i === selectedIdx ? { ...s, points: s.points.map((pt) => ({ x: pt.x + dx, y: pt.y + dy })) } : s)));
      return;
    }
    if (!drawingRef.current) return;
    if (drawTool === "pen" || drawTool === "carrow") {
      // Free-hand (plain line, or a curved arrow): collect the path. Ignore tiny
      // jitter between samples so the smoothed curve stays clean.
      const last = drawingRef.current.points[drawingRef.current.points.length - 1];
      if (!last || Math.hypot(p.x - last.x, p.y - last.y) >= 2.5) {
        drawingRef.current.points.push(p);
      }
    } else {
      // Hold Shift to constrain: straight lines snap to 45°, boxes/circles
      // become square — exactly like Word.
      drawingRef.current.points[1] = e.shiftKey
        ? constrainEnd(drawingRef.current.points[0], p, drawTool)
        : p;
    }
    redrawCanvas(drawingRef.current);
  };
  const onDrawUp = () => {
    moveDragRef.current = null;
    resizeHandleRef.current = null;
    boxDragRef.current = null;
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

  // ---- Remove the current drawing & start fresh ----
  // Brings back the clean ORIGINAL photo (before any drawing) as the base and
  // wipes every stroke, so you can draw a brand-new drawing on the clean photo.
  // Works for new drawings (original kept) and older ones the server can still
  // recover; if no clean original is found it just clears the strokes.
  const [resetting, setResetting] = useState(false);
  const removeAndRedraw = async () => {
    if (saving || resetting) return;
    setResetting(true);
    let targetUrl = baseUrlRef.current || photoUrl;
    try {
      const res = await fetch(
        `/api/reports/${encodeURIComponent(reportId)}/photos?resolveOriginal=${encodeURIComponent(photoId)}`,
        { headers: annoAuthHeaders(), credentials: "include" }
      );
      if (res.ok) {
        const d = (await res.json().catch(() => null)) as { originalUrl?: string } | null;
        if (d?.originalUrl) targetUrl = String(d.originalUrl);
      }
    } catch {
      /* offline / not found — fall back to the current base */
    }
    const finish = () => {
      setStrokes([]);
      setRedoStack([]);
      setSelectedIdx(null);
      setTextDraft(null);
      setResetting(false);
    };
    try {
      const loaded = await loadCanvasSafeImage(targetUrl);
      const img = new Image();
      img.onload = () => {
        baseImgRef.current = img;
        baseUrlRef.current = targetUrl;
        const canvas = canvasRef.current;
        if (canvas) {
          const w = img.naturalWidth || 1200;
          const h = img.naturalHeight || 800;
          canvas.width = w;
          canvas.height = h;
          const maxW = Math.min(window.innerWidth * 0.9, 1280);
          const maxH = window.innerHeight * 0.72;
          const scale = Math.min(maxW / w, maxH / h, 1);
          canvas.style.width = `${Math.round(w * scale)}px`;
          canvas.style.height = `${Math.round(h * scale)}px`;
        }
        finish(); // setStrokes([]) triggers the redraw effect with the new base
      };
      img.onerror = finish;
      img.src = loaded.src;
    } catch {
      finish();
    }
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

  // Word-style keyboard shortcuts for the selected shape: Delete removes it,
  // arrow keys nudge it (Shift = bigger steps), Esc deselects, Ctrl/Cmd+Z / +Y
  // undo/redo, Ctrl/Cmd+D duplicates. Ignored while typing a text label.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (textDraft) return;
      const tag = (e.target as HTMLElement | null)?.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA") return;
      const mod = e.ctrlKey || e.metaKey;
      if (mod && (e.key === "z" || e.key === "Z")) {
        e.preventDefault();
        if (e.shiftKey) redo();
        else undo();
        return;
      }
      if (mod && (e.key === "y" || e.key === "Y")) {
        e.preventDefault();
        redo();
        return;
      }
      if (mod && (e.key === "d" || e.key === "D")) {
        e.preventDefault();
        duplicateSelected();
        return;
      }
      if (e.key === "Escape") {
        setSelectedIdx(null);
        return;
      }
      if (selectedIdx == null) return;
      if (e.key === "Delete" || e.key === "Backspace") {
        e.preventDefault();
        deleteSelected();
        return;
      }
      const step = e.shiftKey ? 10 : 1;
      let dx = 0;
      let dy = 0;
      if (e.key === "ArrowUp") dy = -step;
      else if (e.key === "ArrowDown") dy = step;
      else if (e.key === "ArrowLeft") dx = -step;
      else if (e.key === "ArrowRight") dx = step;
      if (dx || dy) {
        e.preventDefault();
        setStrokes((prev) =>
          prev.map((s, i) =>
            i === selectedIdx ? { ...s, points: s.points.map((pt) => ({ x: pt.x + dx, y: pt.y + dy })) } : s
          )
        );
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedIdx, strokes, redoStack, textDraft]);

  // Apply a text draft to the strokes: EDIT the existing stroke in place when
  // editIdx is set (empty text removes it), otherwise append a new text stroke.
  const strokesWithDraft = (d: TextDraft | null): Stroke[] => {
    if (!d) return strokes;
    const val = d.value.trim();
    if (d.editIdx != null) {
      if (!val) return strokes.filter((_, i) => i !== d.editIdx);
      // Apply the picker's current colour too — it was seeded from this text's
      // colour when the edit started, so re-typing keeps the colour and picking
      // a new colour recolours the text.
      return strokes.map((s, i) =>
        i === d.editIdx ? { ...s, text: val, color: drawColor } : s
      );
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
      // Strokes to persist + flatten. A pending text draft is applied here so an
      // EDIT replaces the old text instead of drawing on top of it, and no
      // selection box leaks into the saved image.
      const strokesToSave = strokesWithDraft(textDraft);
      if (textDraft) {
        const ctx = canvas.getContext("2d");
        if (ctx) {
          ctx.clearRect(0, 0, canvas.width, canvas.height);
          const base = baseImgRef.current;
          if (base) ctx.drawImage(base, 0, 0, canvas.width, canvas.height);
          for (const s of strokesToSave) drawOneStroke(ctx, s);
        }
        setStrokes(strokesToSave);
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
        // Persist the vector strokes + the untouched base image so this drawing
        // can be reopened and edited later instead of drawn over.
        annoJson: JSON.stringify(strokesToSave),
        annoBaseUrl: baseUrlRef.current || photoUrl,
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
          {textDraft ? (
            // Inline controls over the label being typed: drag to move, A−/A+ to
            // resize — no need to click a button or switch tools.
            <div
              onMouseDown={(e) => e.preventDefault()}
              style={{
                position: "absolute",
                left: Math.max(2, textDraft.dispX),
                top: Math.max(2, textDraft.dispY - 40),
                display: "flex",
                gap: 4,
                alignItems: "center",
                zIndex: 6,
              }}
            >
              <div
                onPointerDown={startTextDrag}
                onPointerMove={onTextDragMove}
                onPointerUp={endTextDrag}
                onPointerCancel={endTextDrag}
                title="Drag to move the text"
                style={{
                  width: 36,
                  height: 32,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  background: "#0f172a",
                  color: "#fff",
                  borderRadius: 6,
                  cursor: "grab",
                  touchAction: "none",
                  fontWeight: 900,
                  fontSize: 15,
                }}
              >
                ✥
              </div>
              <button
                type="button"
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => changeTextSize(1 / 1.2)}
                title="Smaller text"
                style={{ height: 32, minWidth: 34, borderRadius: 6, border: "none", background: "#0f172a", color: "#fff", fontWeight: 900, fontSize: 13, cursor: "pointer" }}
              >
                A−
              </button>
              <button
                type="button"
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => changeTextSize(1.2)}
                title="Bigger text"
                style={{ height: 32, minWidth: 34, borderRadius: 6, border: "none", background: "#0f172a", color: "#fff", fontWeight: 900, fontSize: 17, cursor: "pointer" }}
              >
                A+
              </button>
            </div>
          ) : null}
        </div>

        <div style={S.tools}>
          {/* Everyday tools — the few surveyors actually use, named like Word. */}
          {([
            ["move", "✥", "Move"],
            ["arrow", "➔", "Arrow"],
            ["carrow", "⤷", "Curve arrow"],
            ["darrow", "↔", "Measure"],
            ["line", "—", "Line"],
            ["pen", "✎", "Draw"],
            ["text", "T", "Text"],
          ] as const).map(([t, icon, label]) => (
            <button key={t} onClick={() => setDrawTool(t)} title={label} style={{ ...S.toolBtn, ...(drawTool === t ? S.toolBtnActive : {}) }}>
              <span style={{ fontSize: 19, lineHeight: 1 }}>{icon}</span>
              <span style={S.toolLabel}>{label}</span>
            </button>
          ))}
          <button
            onClick={() => setShowMore((v) => !v)}
            title="More shapes"
            style={{ ...S.toolBtn, ...(showMore ? S.toolBtnActive : {}) }}
          >
            <span style={{ fontSize: 19, lineHeight: 1 }}>⋯</span>
            <span style={S.toolLabel}>{showMore ? "Less" : "More"}</span>
          </button>
          {showMore &&
            ([
              ["rect", "▭", "Box"],
              ["rrect", "▢", "Round box"],
              ["ellipse", "◯", "Circle"],
              ["triangle", "△", "Triangle"],
              ["diamond", "◇", "Diamond"],
              ["star", "★", "Star"],
              ["pentagon", "⬠", "Pentagon"],
              ["hexagon", "⬡", "Hexagon"],
              ["callout", "💬", "Callout"],
              ["uparrow", "⬆", "Up arrow"],
              ["downarrow", "⬇", "Down arrow"],
              ["left", "↰", "Left turn"],
              ["right", "↱", "Right turn"],
              ["uturn", "↩", "U-turn"],
              ["x", "✕", "X mark"],
            ] as const).map(([t, icon, label]) => (
              <button key={t} onClick={() => setDrawTool(t)} title={label} style={{ ...S.toolBtn, ...(drawTool === t ? S.toolBtnActive : {}) }}>
                <span style={{ fontSize: 19, lineHeight: 1 }}>{icon}</span>
                <span style={S.toolLabel}>{label}</span>
              </button>
            ))}
          {([
            ["#FFFF00", "Yellow"],
            ["#FF0000", "Red"],
            ["#00B050", "Green"],
            ["#0070C0", "Blue"],
          ] as const).map(([c, label]) => (
            <button
              key={c}
              // Keep the text-edit input focused when a swatch is clicked, so it
              // isn't committed/closed on blur before the colour is applied.
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => pickColor(c)}
              title={label}
              aria-label={label}
              style={{ ...S.swatch, background: c, outline: drawColor.toLowerCase() === c.toLowerCase() ? "2px solid #0f172a" : "2px solid transparent" }}
            />
          ))}
          <input
            type="color"
            value={drawColor}
            onMouseDown={(e) => e.preventDefault()}
            onChange={(e) => pickColor(e.target.value)}
            style={S.colorInput}
            title="Custom colour"
          />
          <select
            value={drawWidth}
            onChange={(e) => setDrawWidth(Number(e.target.value))}
            style={S.widthSelect}
            title="Thickness of arrows / shapes / lines"
          >
            <option value={3}>Thin</option>
            <option value={6}>Medium</option>
            <option value={10}>Thick</option>
          </select>
          {showMore && (
            <>
              <select
                value={drawDash}
                onChange={(e) => setDrawDash(e.target.value as "solid" | "dashed" | "dotted")}
                style={S.widthSelect}
                title="Line style — for arrows / shapes / lines (not text)"
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
            </>
          )}
        </div>

        {/* Text size for a SELECTED text label (while typing, use the floating
            A− / A+ over the label instead). */}
        {drawTool === "move" && selectedIdx != null && strokes[selectedIdx]?.tool === "text" ? (
          <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
            <span style={{ fontSize: 12, fontWeight: 800, color: "#0f172a" }}>Text size</span>
            <button type="button" onClick={() => changeTextSize(1 / 1.2)} title="Smaller text" style={{ minWidth: 44, height: 40, borderRadius: 8, border: "1px solid #d7dbe0", background: "#fff", fontWeight: 900, fontSize: 15, cursor: "pointer", color: "#0f172a" }}>A−</button>
            <button type="button" onClick={() => changeTextSize(1.2)} title="Bigger text" style={{ minWidth: 44, height: 40, borderRadius: 8, border: "1px solid #d7dbe0", background: "#fff", fontWeight: 900, fontSize: 19, cursor: "pointer", color: "#0f172a" }}>A+</button>
          </div>
        ) : null}

        {drawTool === "move" ? (
          <div style={{ fontSize: 11, fontWeight: 700, color: "#64748b" }}>
            {selectedIdx == null
              ? "Tap any drawing or text to select it."
              : "Drag inside to move · drag a corner/edge handle to resize · drag the top handle to rotate · Delete to remove."}
          </div>
        ) : null}
        {drawTool === "text" ? (
          <div style={{ fontSize: 11, fontWeight: 700, color: "#64748b" }}>
            Tap to place a label · tap an existing label to re-edit it · Enter to confirm.
          </div>
        ) : null}
        {drawTool !== "move" && drawTool !== "text" && drawTool !== "pen" && drawTool !== "carrow" ? (
          <div style={{ fontSize: 11, fontWeight: 700, color: "#64748b" }}>
            Hold <b>Shift</b> for a straight line / perfect square or circle (like Word).
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
            onClick={async () => {
              if (!total) return;
              const ok = await confirmDialog(
                "Clear the whole drawing? This removes every arrow and label on this photo.",
                { confirmText: "Clear all", cancelText: "Keep it", danger: true }
              );
              if (!ok) return;
              setStrokes([]);
              setRedoStack([]);
              setSelectedIdx(null);
              setTextDraft(null);
            }}
            disabled={!total}
            title="Remove ALL arrows and labels from this photo"
          >
            Clear all
          </button>
          <button
            style={{ ...S.ghost, borderColor: "#F79009", color: "#B54708" }}
            onClick={removeAndRedraw}
            disabled={resetting}
            title="Remove the current drawing and bring back the clean photo, so you can draw a new one"
          >
            {resetting ? "Removing…" : "🧹 Remove drawing & redraw"}
          </button>
          <button style={S.ghost} onClick={onClose}>Cancel</button>
          <button
            style={S.primary}
            onClick={saveAnnotated}
            // Enabled when there's something to save OR the photo had a drawing
            // that was cleared (so Save writes the clean photo back).
            disabled={saving || (total === 0 && !hadDrawing && !textDraft)}
          >
            {saving ? "Saving…" : "Save"}
          </button>
        </div>
      </div>
    </div>
  );
}

const S: Record<string, React.CSSProperties> = {
  overlay: { position: "fixed", inset: 0, background: "rgba(2,6,23,0.82)", zIndex: 4000, display: "flex", alignItems: "center", justifyContent: "center", padding: 0 },
  card: { background: "#fff", borderRadius: 0, padding: 10, width: "100vw", height: "100dvh", maxWidth: "100vw", maxHeight: "100dvh", overflow: "hidden", display: "flex", flexDirection: "column", gap: 8, boxShadow: "none" },
  head: { display: "flex", alignItems: "center", justifyContent: "space-between", color: "#0f172a", flexShrink: 0 },
  close: { width: 38, height: 38, borderRadius: 8, border: "1px solid #e2e8f0", background: "#f8fafc", cursor: "pointer", fontSize: 16, color: "#0f172a" },
  imgWrap: { position: "relative", alignSelf: "center", lineHeight: 0, flex: 1, display: "flex", alignItems: "center", justifyContent: "center", minHeight: 0, overflow: "hidden" },
  canvas: { touchAction: "none", cursor: "crosshair", borderRadius: 8, display: "block", background: "#0b1220", flexShrink: 0 },
  tools: { display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap", flexShrink: 0, maxHeight: "30vh", overflowY: "auto" },
  toolBtn: { minWidth: 48, height: 48, padding: "2px 6px", borderRadius: 10, border: "1px solid #d7dbe0", background: "#fff", cursor: "pointer", color: "#0f172a", display: "inline-flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 1 },
  toolLabel: { fontSize: 9, fontWeight: 800, lineHeight: 1, color: "inherit" },
  toolBtnActive: { background: "#0f172a", color: "#fff", borderColor: "#0f172a" },
  colorInput: { width: 38, height: 38, border: "1px solid #d7dbe0", borderRadius: 8, cursor: "pointer", padding: 2, background: "#fff" },
  swatch: { width: 30, height: 30, borderRadius: "50%", border: "1px solid rgba(0,0,0,0.15)", cursor: "pointer", outlineOffset: 2, flexShrink: 0 },
  widthSelect: { height: 38, borderRadius: 8, border: "1px solid #d7dbe0", padding: "0 8px", fontSize: 14, cursor: "pointer" },
  actions: { display: "flex", gap: 6, flexWrap: "wrap", flexShrink: 0 },
  ghost: { flex: 1, padding: "9px 10px", borderRadius: 8, border: "1px solid #d7dbe0", background: "#fff", fontWeight: 600, fontSize: 13, cursor: "pointer", color: "#0f172a" },
  primary: { flex: 1, padding: "9px 10px", borderRadius: 8, border: "none", background: "#16a34a", color: "#fff", fontWeight: 700, fontSize: 13, cursor: "pointer" },
};
