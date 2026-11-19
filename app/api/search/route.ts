import pool from "../../../lib/db";
import { requireAuth } from "../../../lib/auth";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type DbColumnRow = { Field?: string };

async function getColumns(table: string): Promise<Set<string>> {
  try {
    const [rows] = await pool.query(`SHOW COLUMNS FROM ${table}`);
    return new Set(
      (Array.isArray(rows) ? rows : []).map((r) =>
        String((r as DbColumnRow).Field || "").toLowerCase()
      )
    );
  } catch {
    return new Set();
  }
}

function has(cols: Set<string>, col: string) {
  return cols.has(col.toLowerCase());
}

function isUnauthorizedError(e: unknown) {
  return (e as { message?: string })?.message === "Unauthorized";
}

// ---- Query parsing -------------------------------------------------------

const STOPWORDS = new Set([
  "i", "need", "want", "reports", "report", "for", "the", "a", "an", "of",
  "route", "routes", "show", "me", "find", "get", "give", "please", "road",
  "survey", "project", "projects", "any", "is", "there", "do", "we", "have",
  "which", "what", "where", "has", "map", "between", "from", "near", "around",
]);

function splitEndpoints(raw: string): { from: string; to: string } | null {
  const seps = [" to ", " - ", " – ", " — ", " → ", "->", " towards ", " till "];
  const lower = raw.toLowerCase();
  for (const sep of seps) {
    const idx = lower.indexOf(sep);
    if (idx > 0) {
      const from = raw.slice(0, idx).trim();
      const to = raw.slice(idx + sep.length).trim();
      if (from && to) return { from: cleanPlace(from), to: cleanPlace(to) };
    }
  }
  return null;
}

function cleanPlace(s: string): string {
  const words = s.split(/\s+/).map((w) => w.trim()).filter(Boolean);
  while (words.length && STOPWORDS.has(words[0].toLowerCase())) words.shift();
  while (words.length && STOPWORDS.has(words[words.length - 1].toLowerCase())) words.pop();
  return words.join(" ").trim();
}

function placeTokens(s: string): string[] {
  return s
    .toLowerCase()
    .split(/[^a-z0-9ऀ-෿]+/i)
    .map((w) => w.trim())
    .filter((w) => w.length >= 2 && !STOPWORDS.has(w));
}

function placeInText(place: string, text: string): boolean {
  if (!place || !text) return false;
  const p = place.toLowerCase().trim();
  const t = text.toLowerCase();
  if (!p) return false;
  if (t.includes(p)) return true;
  const words = p.split(/\s+/).filter((w) => w.length >= 3);
  if (!words.length) return false;
  return words.every((w) => t.includes(w));
}

function titleCase(s: string): string {
  return s.replace(/\b\w/g, (c) => c.toUpperCase());
}

// ---- Place-name labelling ------------------------------------------------

const INDIAN_STATES = new Set([
  "andhra pradesh", "arunachal pradesh", "assam", "bihar", "chhattisgarh",
  "goa", "gujarat", "haryana", "himachal pradesh", "jharkhand", "karnataka",
  "kerala", "madhya pradesh", "maharashtra", "manipur", "meghalaya", "mizoram",
  "nagaland", "odisha", "punjab", "rajasthan", "sikkim", "tamil nadu",
  "telangana", "tripura", "uttar pradesh", "uttarakhand", "west bengal",
  "delhi", "jammu and kashmir", "ladakh", "puducherry", "chandigarh",
]);

function placeLabel(value: unknown): string {
  const t = String(value ?? "").trim();
  if (!t) return "";
  const parts = t.split(",").map((x) => x.trim()).filter(Boolean);
  if (parts.length <= 2 && parts.every((p) => /^-?\d+(\.\d+)?$/.test(p))) return "";
  const cleaned = parts.filter((p) => !/^\d{4,6}$/.test(p) && p.toLowerCase() !== "india");
  const first = cleaned[0] || parts[0] || "";
  const state = cleaned.find((p) => INDIAN_STATES.has(p.toLowerCase())) || "";
  if (first && state && first.toLowerCase() !== state.toLowerCase()) return `${first}, ${state}`;
  return first || state;
}

type Pt = Record<string, unknown>;

function bestLoc(r: Pt): string {
  return (
    String(r.resolved_location || "").trim() ||
    String(r.location || "").trim() ||
    String((r as any).address || "").trim() ||
    String((r as any).location_name || "").trim() ||
    ""
  );
}

function coordOf(r: Pt): { lat: number; lng: number } | null {
  const lat = Number(r.latitude ?? r.loc_lat);
  const lng = Number(r.longitude ?? r.loc_lon);
  if (Number.isFinite(lat) && Number.isFinite(lng) && (lat !== 0 || lng !== 0)) return { lat, lng };
  return null;
}

function diffBucket(value: unknown): "red" | "amber" | "green" | "" {
  const v = String(value || "").trim().toLowerCase();
  if (!v) return "";
  if (["red", "difficult", "hard", "high", "critical"].includes(v)) return "red";
  if (["amber", "orange", "yellow", "medium", "moderate"].includes(v)) return "amber";
  if (["green", "easy", "low", "clear", "ok"].includes(v)) return "green";
  return "";
}

// ---- Geometry (offline corridor test, km) --------------------------------

const DEG_KM = 111;

function haversineKm(a: { lat: number; lng: number }, b: { lat: number; lng: number }): number {
  const R = 6371;
  const toRad = (d: number) => (d * Math.PI) / 180;
  const dLat = toRad(b.lat - a.lat);
  const dLng = toRad(b.lng - a.lng);
  const x =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(toRad(a.lat)) * Math.cos(toRad(b.lat)) * Math.sin(dLng / 2) ** 2;
  return 2 * R * Math.asin(Math.min(1, Math.sqrt(x)));
}

// Project point P onto corridor A→B. Returns t (0=at A, 1=at B) and the
// perpendicular distance in km.
function projectOnCorridor(
  A: { lat: number; lng: number },
  B: { lat: number; lng: number },
  P: { lat: number; lng: number }
): { t: number; dist: number } {
  const lat0 = (A.lat + B.lat) / 2;
  const kx = Math.cos((lat0 * Math.PI) / 180) * DEG_KM;
  const Bx = (B.lng - A.lng) * kx;
  const By = (B.lat - A.lat) * DEG_KM;
  const Px = (P.lng - A.lng) * kx;
  const Py = (P.lat - A.lat) * DEG_KM;
  const len2 = Bx * Bx + By * By;
  if (len2 === 0) return { t: 0, dist: Math.hypot(Px, Py) };
  const t = (Px * Bx + Py * By) / len2;
  const perpX = Px - t * Bx;
  const perpY = Py - t * By;
  return { t, dist: Math.hypot(perpX, perpY) };
}

const BUFFER_KM = 60; // how far off the straight A→B line a point can sit
// (generous, because real roads bend away from the straight line — e.g. the
// Bangalore–Chennai highway dips ~50 km south at Krishnagiri. Over-inclusion is
// safe here: the user un-ticks any wrong report before combining.)
const COVER_FRACTION = 0.5; // share of a report's points that must be on-route

// ---- Types returned ------------------------------------------------------

type Segment = {
  id: string;
  name: string;
  source: string | null;
  status: string | null;
  from: string;
  to: string;
  points: number;
  onRoutePoints: number;
  tStart: number;
  tEnd: number;
  coveragePct: number;
  years: string;
  red: number;
};

type Gap = { pctStart: number; pctEnd: number; fromLabel: string; toLabel: string };

type SmartHit = {
  id: string;
  name: string;
  source: string | null;
  status: string | null;
  from: string;
  to: string;
  points: number;
  years: string;
  red: number;
  amber: number;
  green: number;
  keyObstacles: string[];
  matchedPlaces: string[];
  score: number;
  summaryLine: string;
};

// ---- Per-project rollups (labels, years, difficulty) ---------------------

function projectRollup(pts: Pt[]) {
  const usable = pts.filter((r) => String(r.point_key || "") !== "__NO_GPS__");
  const first = usable[0];
  const last = usable[usable.length - 1];
  const from = (first && placeLabel(bestLoc(first))) || "";
  const to = (last && placeLabel(bestLoc(last))) || "";
  let red = 0, amber = 0, green = 0;
  const years = new Set<string>();
  const catCount = new Map<string, number>();
  const redCatCount = new Map<string, number>();
  for (const r of usable) {
    const b = diffBucket(r.difficulty);
    if (b === "red") red++;
    else if (b === "amber") amber++;
    else if (b === "green") green++;
    const cat = String(r.category || "").trim();
    if (cat && !/^unknown$/i.test(cat)) {
      catCount.set(cat, (catCount.get(cat) || 0) + 1);
      if (b === "red") redCatCount.set(cat, (redCatCount.get(cat) || 0) + 1);
    }
    const d = r.created_at ? new Date(String(r.created_at)) : null;
    if (d && !isNaN(d.getTime())) years.add(String(d.getFullYear()));
  }
  const keyObstacles = Array.from(
    new Set([
      ...Array.from(redCatCount.entries()).sort((a, b) => b[1] - a[1]).map((e) => e[0]),
      ...Array.from(catCount.entries()).sort((a, b) => b[1] - a[1]).map((e) => e[0]),
    ])
  ).slice(0, 4);
  const yl = Array.from(years).sort();
  const yearStr = yl.length ? (yl.length === 1 ? yl[0] : `${yl[0]}–${yl[yl.length - 1]}`) : "";
  return { usable, from, to, red, amber, green, keyObstacles, yearStr };
}

export async function GET(request: Request) {
  try {
    requireAuth(request);
    const url = new URL(request.url);
    const q = String(url.searchParams.get("q") || "").trim();
    if (!q || q.length < 2) {
      return Response.json({ query: q, mode: "simple", results: [], hint: "Type a place or a route like “Chennai to Bangalore”." });
    }

    const projCols = await getColumns("projects");
    const repCols = await getColumns("reports");

    const endpoints = splitEndpoints(q);
    const fromText = endpoints?.from || "";
    const toText = endpoints?.to || "";
    const allTokens = Array.from(
      new Set(placeTokens(endpoints ? `${fromText} ${toText}` : cleanPlace(q)))
    );
    if (!allTokens.length) {
      return Response.json({ query: q, mode: "simple", results: [], hint: "Could not find a place name in that. Try “Chennai to Bangalore”." });
    }

    // ---- Load all live projects.
    const pSelCols = ["id", "name", "title", "project_name", "location", "description", "client_name", "source", "status", "created_at"]
      .filter((c) => has(projCols, c));
    if (!pSelCols.includes("id")) return Response.json({ error: "Projects table has no id column" }, { status: 500 });
    const pWhere = has(projCols, "deleted_at") ? " WHERE deleted_at IS NULL" : "";
    const [pRows] = await pool.query(`SELECT ${pSelCols.join(", ")} FROM projects${pWhere}`);
    const projects = (Array.isArray(pRows) ? pRows : []) as Pt[];
    if (!projects.length) return Response.json({ query: q, mode: "simple", results: [] });
    const projById = new Map<string, Pt>();
    for (const p of projects) projById.set(String(p.id), p);

    // ---- Load the points of all live projects (one query).
    const rSelCols = ["project_id", "location", "resolved_location", "address", "location_name",
      "category", "difficulty", "point_key", "latitude", "longitude", "loc_lat", "loc_lon", "sort_order", "created_at"]
      .filter((c) => has(repCols, c));
    const orderParts = ["project_id", ...(has(repCols, "sort_order") ? ["sort_order"] : []), ...(has(repCols, "created_at") ? ["created_at"] : [])];
    const [rRows] = await pool.query(`SELECT ${rSelCols.join(", ")} FROM reports ORDER BY ${orderParts.join(", ")}`);
    const reports = (Array.isArray(rRows) ? rRows : []) as Pt[];
    const byProject = new Map<string, Pt[]>();
    for (const r of reports) {
      const pid = String(r.project_id || "").trim();
      if (!pid || !projById.has(pid)) continue;
      if (!byProject.has(pid)) byProject.set(pid, []);
      byProject.get(pid)!.push(r);
    }

    // ===== Route-coverage mode (needs two endpoints we can locate on the map) =====
    if (endpoints) {
      const anchor = (text: string): { lat: number; lng: number } | null => {
        let sumLat = 0, sumLng = 0, n = 0;
        for (const r of reports) {
          const c = coordOf(r);
          if (!c) continue;
          if (placeInText(text, bestLoc(r))) { sumLat += c.lat; sumLng += c.lng; n++; }
        }
        return n ? { lat: sumLat / n, lng: sumLng / n } : null;
      };
      const A = anchor(fromText);
      const B = anchor(toText);

      if (A && B && haversineKm(A, B) > 1) {
        const segments: Segment[] = [];
        for (const [pid, pts] of byProject) {
          const roll = projectRollup(pts);
          const coords = roll.usable.map((r) => ({ r, c: coordOf(r) })).filter((x) => x.c) as { r: Pt; c: { lat: number; lng: number } }[];
          if (coords.length < 2) continue;
          let tMin = Infinity, tMax = -Infinity, onRoute = 0;
          for (const { c } of coords) {
            const { t, dist } = projectOnCorridor(A, B, c);
            if (dist <= BUFFER_KM && t >= -0.05 && t <= 1.05) {
              onRoute++;
              if (t < tMin) tMin = t;
              if (t > tMax) tMax = t;
            }
          }
          const frac = onRoute / coords.length;
          if (onRoute < 2 || frac < COVER_FRACTION) continue;
          const tStart = Math.max(0, Math.min(1, tMin));
          const tEnd = Math.max(0, Math.min(1, tMax));
          if (tEnd - tStart < 0.01) continue;
          const p = projById.get(pid)!;
          segments.push({
            id: pid,
            name: String(p.name || p.title || p.project_name || "Untitled Project"),
            source: (p.source as string) ?? null,
            status: (p.status as string) ?? null,
            from: roll.from, to: roll.to,
            points: roll.usable.length,
            onRoutePoints: onRoute,
            tStart, tEnd,
            coveragePct: Math.round((tEnd - tStart) * 100),
            years: roll.yearStr, red: roll.red,
          });
        }

        if (segments.length) {
          segments.sort((a, b) => a.tStart - b.tStart || b.tEnd - a.tEnd);

          // Merge covered intervals → find gaps on [0,1].
          const merged: [number, number][] = [];
          for (const s of [...segments].sort((a, b) => a.tStart - b.tStart)) {
            if (!merged.length || s.tStart > merged[merged.length - 1][1] + 0.001) merged.push([s.tStart, s.tEnd]);
            else merged[merged.length - 1][1] = Math.max(merged[merged.length - 1][1], s.tEnd);
          }
          const fromLabel = titleCase(fromText);
          const toLabel = titleCase(toText);
          const anchors: { t: number; label: string }[] = [{ t: 0, label: fromLabel }, { t: 1, label: toLabel }];
          for (const s of segments) {
            if (s.from) anchors.push({ t: s.tStart, label: s.from });
            if (s.to) anchors.push({ t: s.tEnd, label: s.to });
          }
          anchors.sort((a, b) => a.t - b.t);
          const labelBefore = (t: number) => { let L = fromLabel; for (const a of anchors) if (a.t <= t + 0.02 && a.label) L = a.label; return L; };
          const labelAfter = (t: number) => { for (const a of anchors) if (a.t >= t - 0.02 && a.label) return a.label; return toLabel; };

          const gaps: Gap[] = [];
          let cursor = 0;
          for (const [s0, s1] of merged) {
            if (s0 - cursor > 0.04) gaps.push({ pctStart: Math.round(cursor * 100), pctEnd: Math.round(s0 * 100), fromLabel: labelBefore(cursor), toLabel: labelAfter(s0) });
            cursor = Math.max(cursor, s1);
          }
          if (1 - cursor > 0.04) gaps.push({ pctStart: Math.round(cursor * 100), pctEnd: 100, fromLabel: labelBefore(cursor), toLabel: toLabel });

          const coveredPct = Math.round(merged.reduce((s, [a, b]) => s + (b - a), 0) * 100);

          return Response.json({
            query: q,
            mode: "coverage",
            from: fromText, to: toText,
            corridor: { from: { label: fromLabel, lat: A.lat, lng: A.lng }, to: { label: toLabel, lat: B.lat, lng: B.lng }, km: Math.round(haversineKm(A, B)) },
            coveredPct,
            segments,
            gaps,
          });
        }
      }
      // else: fall through to simple text search below (couldn't locate the
      // endpoints on the map, or nothing lies along the corridor).
    }

    // ===== Simple text search (single place, or route we couldn't map) =====
    const hits: SmartHit[] = [];
    for (const [pid, pts] of byProject) {
      const p = projById.get(pid)!;
      const roll = projectRollup(pts);
      const name = String(p.name || p.title || p.project_name || "Untitled Project");
      const projFieldText = [p.name, p.title, p.project_name, p.location, p.description].map((x) => String(x ?? "")).join(" | ");
      const pointText = roll.usable.map((r) => bestLoc(r)).join(" | ");
      const allText = `${projFieldText} | ${pointText}`.toLowerCase();

      let score = 0;
      const matchedPlaces: string[] = [];
      if (endpoints) {
        const fromHit = placeInText(fromText, allText);
        const toHit = placeInText(toText, allText);
        if (fromHit) matchedPlaces.push(fromText);
        if (toHit) matchedPlaces.push(toText);
        if (fromHit && toHit) score += 100;
        else if (fromHit || toHit) score += 40;
      } else {
        let matched = 0;
        for (const tok of allTokens) if (allText.includes(tok)) { matched++; matchedPlaces.push(tok); }
        score += allTokens.length ? Math.round((matched / allTokens.length) * 100) : 0;
      }
      if (name.toLowerCase().includes(q.toLowerCase())) score += 30;
      if (score <= 0) continue;

      const routeStr = roll.from && roll.to ? `${roll.from} → ${roll.to}` : roll.from || roll.to || name;
      const summaryLine = [routeStr, `${roll.usable.length} point${roll.usable.length === 1 ? "" : "s"}`, roll.yearStr, roll.red ? `${roll.red} difficult` : ""].filter(Boolean).join(" · ");
      hits.push({
        id: pid, name, source: (p.source as string) ?? null, status: (p.status as string) ?? null,
        from: roll.from, to: roll.to, points: roll.usable.length, years: roll.yearStr,
        red: roll.red, amber: roll.amber, green: roll.green, keyObstacles: roll.keyObstacles,
        matchedPlaces: Array.from(new Set(matchedPlaces)), score, summaryLine,
      });
    }
    hits.sort((a, b) => b.score - a.score || b.points - a.points);

    return Response.json({ query: q, mode: "simple", from: fromText || null, to: toText || null, results: hits.slice(0, 25) });
  } catch (error) {
    if (isUnauthorizedError(error)) return Response.json({ error: "Unauthorized" }, { status: 401 });
    console.error("[api/search] GET error:", error);
    return Response.json({ error: "Search failed" }, { status: 500 });
  }
}
