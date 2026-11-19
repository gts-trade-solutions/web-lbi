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

// Words that carry no place meaning in a query like
// "i need reports for chennai to bangalore".
const STOPWORDS = new Set([
  "i", "need", "want", "reports", "report", "for", "the", "a", "an", "of",
  "route", "routes", "show", "me", "find", "get", "give", "please", "road",
  "survey", "project", "projects", "any", "is", "there", "do", "we", "have",
  "which", "what", "where", "has", "map", "between", "from", "near", "around",
]);

// Split a query into route endpoints when it reads like "X to Y".
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

// Strip leading filler ("i need reports for chennai") down to the place part.
function cleanPlace(s: string): string {
  const words = s
    .split(/\s+/)
    .map((w) => w.trim())
    .filter(Boolean);
  // Drop leading stopwords only (keep internal words so "new delhi" survives).
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

// Does `place` (e.g. "new delhi") appear inside `text`? Match the whole
// phrase first, then fall back to "all words present".
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

// ---- Place-name labelling (compact form of summaryRouteMap.parseLocation) --

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
  if (first && state && first.toLowerCase() !== state.toLowerCase()) {
    return `${first}, ${state}`;
  }
  return first || state;
}

function bestLoc(r: Record<string, unknown>): string {
  return (
    String(r.resolved_location || "").trim() ||
    String(r.location || "").trim() ||
    String((r as any).address || "").trim() ||
    String((r as any).location_name || "").trim() ||
    ""
  );
}

function diffBucket(value: unknown): "red" | "amber" | "green" | "" {
  const v = String(value || "").trim().toLowerCase();
  if (!v) return "";
  if (["red", "difficult", "hard", "high", "critical"].includes(v)) return "red";
  if (["amber", "orange", "yellow", "medium", "moderate"].includes(v)) return "amber";
  if (["green", "easy", "low", "clear", "ok"].includes(v)) return "green";
  return "";
}

type ProjectHit = {
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

export async function GET(request: Request) {
  try {
    requireAuth(request);

    const url = new URL(request.url);
    const q = String(url.searchParams.get("q") || "").trim();
    if (!q || q.length < 2) {
      return Response.json({ query: q, results: [], hint: "Type a place or a route like “Chennai to Bangalore”." });
    }

    const projCols = await getColumns("projects");
    const repCols = await getColumns("reports");

    const endpoints = splitEndpoints(q);
    const fromText = endpoints?.from || "";
    const toText = endpoints?.to || "";
    // Every place token we care about (for free-text and candidate finding).
    const allTokens = Array.from(
      new Set([
        ...placeTokens(endpoints ? `${fromText} ${toText}` : cleanPlace(q)),
      ])
    );
    if (!allTokens.length) {
      return Response.json({ query: q, results: [], hint: "Could not find a place name in that. Try “Chennai to Bangalore”." });
    }

    // ---- 1. Load all live projects (exclude recycle bin only; keep finished
    // so old/closed routes are still findable).
    const pSelCols = [
      "id", "name", "title", "project_name", "location", "description",
      "client_name", "source", "status", "created_at",
    ].filter((c) => has(projCols, c));
    if (!pSelCols.includes("id")) {
      return Response.json({ error: "Projects table has no id column" }, { status: 500 });
    }
    const pWhere: string[] = [];
    if (has(projCols, "deleted_at")) pWhere.push("deleted_at IS NULL");
    const [pRows] = await pool.query(
      `SELECT ${pSelCols.join(", ")} FROM projects${pWhere.length ? ` WHERE ${pWhere.join(" AND ")}` : ""}`
    );
    const projects = (Array.isArray(pRows) ? pRows : []) as Record<string, unknown>[];
    if (!projects.length) return Response.json({ query: q, results: [] });

    const projById = new Map<string, Record<string, unknown>>();
    for (const p of projects) projById.set(String(p.id), p);

    // ---- 2. Find candidate projects whose POINTS mention any token.
    const repTextCols = [
      "location", "resolved_location", "address", "location_name", "point_key",
    ].filter((c) => has(repCols, c));
    const candidateIds = new Set<string>();

    if (repTextCols.length && has(repCols, "project_id")) {
      const likeClauses: string[] = [];
      const likeArgs: unknown[] = [];
      for (const tok of allTokens) {
        for (const col of repTextCols) {
          likeClauses.push(`LOWER(${col}) LIKE ?`);
          likeArgs.push(`%${tok}%`);
        }
      }
      if (likeClauses.length) {
        const [cRows] = await pool.query(
          `SELECT DISTINCT project_id FROM reports WHERE ${likeClauses.join(" OR ")}`,
          likeArgs
        );
        for (const r of Array.isArray(cRows) ? cRows : []) {
          const pid = String((r as { project_id?: unknown }).project_id || "").trim();
          if (pid && projById.has(pid)) candidateIds.add(pid);
        }
      }
    }

    // Projects whose OWN fields (name / location) match a token are candidates too.
    for (const p of projects) {
      const hay = [
        p.name, p.title, p.project_name, p.location, p.description, p.client_name,
      ]
        .map((x) => String(x ?? "").toLowerCase())
        .join(" | ");
      if (allTokens.some((t) => hay.includes(t))) candidateIds.add(String(p.id));
    }

    if (!candidateIds.size) return Response.json({ query: q, results: [] });

    // ---- 3. Pull the points of candidate projects (one query) for summaries.
    const repSelCols = [
      "project_id", "location", "resolved_location", "address", "location_name",
      "category", "difficulty", "remarks_action", "description", "point_key",
      "sort_order", "created_at",
    ].filter((c) => has(repCols, c));
    const idList = Array.from(candidateIds);
    const placeholders = idList.map(() => "?").join(",");
    const orderParts: string[] = ["project_id"];
    if (has(repCols, "sort_order")) orderParts.push("sort_order");
    if (has(repCols, "created_at")) orderParts.push("created_at");
    const [rRows] = await pool.query(
      `SELECT ${repSelCols.join(", ")} FROM reports WHERE project_id IN (${placeholders}) ORDER BY ${orderParts.join(", ")}`,
      idList
    );
    const reports = (Array.isArray(rRows) ? rRows : []) as Record<string, unknown>[];
    const byProject = new Map<string, Record<string, unknown>[]>();
    for (const r of reports) {
      const pid = String(r.project_id || "").trim();
      if (!pid) continue;
      if (!byProject.has(pid)) byProject.set(pid, []);
      byProject.get(pid)!.push(r);
    }

    // ---- 4. Score + summarise each candidate.
    const hits: ProjectHit[] = [];
    for (const pid of candidateIds) {
      const p = projById.get(pid)!;
      const pts = (byProject.get(pid) || []).filter(
        (r) => String(r.point_key || "") !== "__NO_GPS__"
      );
      const name = String(p.name || p.title || p.project_name || "Untitled Project");

      const first = pts[0];
      const last = pts[pts.length - 1];
      const fromLabel = (first && placeLabel(bestLoc(first))) || "";
      const toLabel = (last && placeLabel(bestLoc(last))) || "";

      // All place text for this project (name + location + every point).
      const projFieldText = [p.name, p.title, p.project_name, p.location, p.description]
        .map((x) => String(x ?? ""))
        .join(" | ");
      const pointText = pts.map((r) => bestLoc(r)).join(" | ");
      const allText = `${projFieldText} | ${pointText}`;

      // Difficulty + category rollups.
      let red = 0, amber = 0, green = 0;
      const catCount = new Map<string, number>();
      const redCatCount = new Map<string, number>();
      const years = new Set<string>();
      for (const r of pts) {
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
      // Key obstacles: red categories first, then most common overall.
      const keyObstacles = Array.from(
        new Set([
          ...Array.from(redCatCount.entries()).sort((a, b) => b[1] - a[1]).map((e) => e[0]),
          ...Array.from(catCount.entries()).sort((a, b) => b[1] - a[1]).map((e) => e[0]),
        ])
      ).slice(0, 4);

      // ---- scoring
      let score = 0;
      const matchedPlaces: string[] = [];
      if (endpoints) {
        const fromHit = placeInText(fromText, allText);
        const toHit = placeInText(toText, allText);
        if (fromHit) matchedPlaces.push(fromText);
        if (toHit) matchedPlaces.push(toText);
        if (fromHit && toHit) {
          score += 100;
          // Direction bonus: "from" shows up near the start, "to" near the end.
          const headText = pts.slice(0, Math.max(1, Math.ceil(pts.length * 0.3)))
            .map((r) => bestLoc(r)).join(" | ") + " | " + projFieldText;
          const tailText = pts.slice(Math.floor(pts.length * 0.7))
            .map((r) => bestLoc(r)).join(" | ") + " | " + projFieldText;
          if (placeInText(fromText, headText) && placeInText(toText, tailText)) score += 25;
        } else if (fromHit || toHit) {
          score += 40;
        }
      } else {
        let matched = 0;
        for (const tok of allTokens) {
          if (allText.toLowerCase().includes(tok)) {
            matched++;
            matchedPlaces.push(tok);
          }
        }
        score += allTokens.length ? Math.round((matched / allTokens.length) * 100) : 0;
      }
      // Whole-query substring in the project name is a strong signal.
      if (name.toLowerCase().includes(q.toLowerCase())) score += 30;
      if (score <= 0) continue;

      const yearList = Array.from(years).sort();
      const yearStr = yearList.length
        ? yearList.length === 1
          ? yearList[0]
          : `${yearList[0]}–${yearList[yearList.length - 1]}`
        : "";

      const routeStr = fromLabel && toLabel ? `${fromLabel} → ${toLabel}` : fromLabel || toLabel || name;
      const summaryLine = [
        routeStr,
        `${pts.length} point${pts.length === 1 ? "" : "s"}`,
        yearStr,
        red ? `${red} difficult` : "",
      ].filter(Boolean).join(" · ");

      hits.push({
        id: pid,
        name,
        source: (p.source as string) ?? null,
        status: (p.status as string) ?? null,
        from: fromLabel,
        to: toLabel,
        points: pts.length,
        years: yearStr,
        red, amber, green,
        keyObstacles,
        matchedPlaces: Array.from(new Set(matchedPlaces)),
        score,
        summaryLine,
      });
    }

    hits.sort((a, b) => b.score - a.score || b.points - a.points);

    return Response.json({
      query: q,
      from: fromText || null,
      to: toText || null,
      results: hits.slice(0, 25),
    });
  } catch (error) {
    if (isUnauthorizedError(error)) {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/search] GET error:", error);
    return Response.json({ error: "Search failed" }, { status: 500 });
  }
}
