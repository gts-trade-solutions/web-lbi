"use client";

import React, { useEffect, useMemo, useRef, useState } from "react";
import { useParams, useSearchParams, useRouter } from "next/navigation";

type ReportPoint = {
  id: string;
  n: number;
  lat: number;
  lng: number;
  category: string;
  description: string;
  remarks: string;
  location: string;
  state: string;
};

const INDIAN_STATES = new Set([
  "andhra pradesh", "arunachal pradesh", "assam", "bihar", "chhattisgarh", "goa",
  "gujarat", "haryana", "himachal pradesh", "jharkhand", "karnataka", "kerala",
  "madhya pradesh", "maharashtra", "manipur", "meghalaya", "mizoram", "nagaland",
  "odisha", "punjab", "rajasthan", "sikkim", "tamil nadu", "telangana", "tripura",
  "uttar pradesh", "uttarakhand", "west bengal", "delhi", "jammu and kashmir",
  "ladakh", "puducherry", "chandigarh",
]);

// Short place name from a resolved-location string ("Alandur, Chennai, TN" -> "Alandur").
function placeName(value: unknown): string {
  let t = String(value ?? "").trim();
  if (!t) return "";
  // Old imported survey reports label the place cell as "Location - <name>"
  // (e.g. "Location – Ekalaghogra"); strip that leading label so the card
  // shows the place only. Requires a dash/colon so real names like
  // "Location Road" are left untouched.
  t = t.replace(/^location\s*[-–—:]+\s*/i, "").trim();
  const parts = t.split(",").map((x) => x.trim()).filter(Boolean);
  if (parts.length <= 2 && parts.every((p) => /^-?\d+(\.\d+)?$/.test(p))) return "";
  return parts[0] || "";
}

// State name from a resolved-location string.
function stateName(value: unknown): string {
  const t = String(value ?? "").trim();
  if (!t) return "";
  const parts = t.split(",").map((x) => x.trim()).filter(Boolean);
  const cleaned = parts.filter((p) => !/^\d{4,6}$/.test(p) && p.toLowerCase() !== "india");
  const match = cleaned.find((p) => INDIAN_STATES.has(p.toLowerCase()));
  return match || (cleaned.length > 1 ? cleaned[cleaned.length - 1] : "");
}

const GMAPS_KEY = process.env.NEXT_PUBLIC_GOOGLE_MAPS_API_KEY || "";

function authHeaders(): Record<string, string> {
  const token =
    typeof window !== "undefined" ? localStorage.getItem("auth_token") : null;
  return token ? { Authorization: `Bearer ${token}` } : {};
}

// Inject the Google Maps JS API "inline bootstrap loader" exactly once.
let gmapsLoader: Promise<void> | null = null;
function loadGoogleMaps(key: string): Promise<void> {
  if (typeof window === "undefined") return Promise.resolve();
  if ((window as any).google?.maps?.importLibrary) return Promise.resolve();
  if (gmapsLoader) return gmapsLoader;

  gmapsLoader = new Promise<void>((resolve, reject) => {
    try {
      ((g: any) => {
        let h: any;
        const c = "google";
        const m = document;
        const b = window as any;
        const d = (b[c] || (b[c] = {})).maps || ((b[c] = b[c] || {}).maps = {});
        const r = new Set<string>();
        const e = new URLSearchParams();
        const u = () =>
          h ||
          (h = new Promise<void>(async (res, rej) => {
            const a = m.createElement("script");
            e.set("libraries", Array.from(r) + "");
            for (const k in g)
              e.set(k.replace(/[A-Z]/g, (t) => "_" + t[0].toLowerCase()), g[k]);
            e.set("callback", c + ".maps.__ib__");
            a.src = `https://maps.${c}apis.com/maps/api/js?` + e;
            (d as any).__ib__ = res;
            a.onerror = () => (h = rej(Error("Google Maps could not load.")));
            m.head.append(a);
          }));
        (d as any).importLibrary = (f: string, ...n: any[]) =>
          r.add(f) && u().then(() => (d as any).importLibrary(f, ...n));
      })({ key, v: "weekly" });

      const check = () => {
        if ((window as any).google?.maps?.importLibrary) resolve();
        else setTimeout(check, 50);
      };
      check();
    } catch (err) {
      reject(err as Error);
    }
  });
  return gmapsLoader;
}

export default function RouteMapPage() {
  const params = useParams();
  const search = useSearchParams();
  const router = useRouter();
  const projectId = useMemo(() => {
    const id = (params as any)?.id;
    return Array.isArray(id) ? id[0] : id;
  }, [params]);

  const selectedIds = useMemo(() => {
    // Preferred: selection stashed in localStorage (avoids huge URLs / HTTP 431).
    if (typeof window !== "undefined" && search?.get("sel") === "1" && projectId) {
      try {
        const raw = localStorage.getItem(`routemap_sel_${projectId}`);
        const arr = raw ? JSON.parse(raw) : [];
        if (Array.isArray(arr)) return arr.map((x: unknown) => String(x)).filter(Boolean);
      } catch {
        /* ignore */
      }
      return [];
    }
    // Backward-compatible: ids in the query string (only safe for small sets).
    const raw = search?.get("reports") || "";
    return raw.split(",").map((s) => s.trim()).filter(Boolean);
  }, [search, projectId]);

  // ---- Share mode ----
  // When the map is opened via a password-protected share link it carries
  // ?share=<token>. In that mode there is NO logged-in user: data comes from
  // the public /api/share/<token>/* endpoints, authorised by a short session
  // token the client got by entering the password on /share/<token>.
  const shareToken = useMemo(() => String(search?.get("share") || "").trim(), [search]);
  const isShare = !!shareToken;
  // Read the unlocked session synchronously on mount (lazy init). Reading it
  // in an effect instead would make the first render see no session and bounce
  // back to the password gate even when we DO have one — an endless loop.
  const [shareSession] = useState<string>(() => {
    if (typeof window === "undefined") return "";
    const t = new URLSearchParams(window.location.search).get("share") || "";
    if (!t) return "";
    try {
      return sessionStorage.getItem(`share_sess_${t}`) || "";
    } catch {
      return "";
    }
  });

  // "Share this report" dialog (authored side only, never in share view).
  const [shareOpen, setShareOpen] = useState(false);
  const [sharePassword, setSharePassword] = useState("");
  const [shareTitle, setShareTitle] = useState("");
  const [shareBusy, setShareBusy] = useState(false);
  const [shareErr, setShareErr] = useState<string | null>(null);
  const [shareLink, setShareLink] = useState<string>("");
  const [shareCopied, setShareCopied] = useState(false);

  // Fetch road-following geometry for a batch of stops via our server proxy
  // (OpenRouteService). Sends the auth header when logged in, or the share
  // session + token in a client's share view. Returns null on any failure so
  // the caller can fall back to a straight segment.
  const fetchRoadPath = async (
    coords: Array<{ lat: number; lng: number }>
  ): Promise<Array<{ lat: number; lng: number }> | null> => {
    try {
      const res = await fetch("/api/route-directions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          ...(isShare ? { "X-Share-Session": shareSession } : authHeaders()),
        },
        ...(isShare ? {} : { credentials: "include" as RequestCredentials }),
        body: JSON.stringify({
          coordinates: coords.map((c) => [c.lat, c.lng]),
          token: isShare ? shareToken : undefined,
        }),
      });
      if (!res.ok) return null;
      const data = await res.json().catch(() => null);
      return Array.isArray(data?.path) && data.path.length >= 2 ? data.path : null;
    } catch {
      return null;
    }
  };

  const [points, setPoints] = useState<ReportPoint[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [mapError, setMapError] = useState<string | null>(null);
  const [routeHint, setRouteHint] = useState<string | null>(null);
  const [ready, setReady] = useState(false);
  const [current, setCurrent] = useState<number>(-1);
  const [playing, setPlaying] = useState(false);
  const [banner, setBanner] = useState<string>("");
  const [lightbox, setLightbox] = useState<{ url: string; reportId: string } | null>(null);
  // Carousel state for the lightbox: all photos for the open report + which
  // one is showing. Manual prev/next only — never auto-advances.
  const [lbPhotos, setLbPhotos] = useState<string[]>([]);
  const [lbIndex, setLbIndex] = useState(0);

  // ---- Draw-on-image annotation state (authored view only) ----
  type Stroke = {
    tool: "arrow" | "pen" | "rect" | "ellipse" | "left" | "right" | "uturn" | "x" | "text";
    color: string;
    width: number;
    points: { x: number; y: number }[]; // pen: path; others: [start, end]; text: [pos]
    text?: string;
    fontSize?: number;
  };
  const [drawMode, setDrawMode] = useState(false);
  const [drawTool, setDrawTool] = useState<Stroke["tool"]>("arrow");
  const [drawColor, setDrawColor] = useState("#FFD400");
  const [drawWidth, setDrawWidth] = useState(6);
  const [strokes, setStrokes] = useState<Stroke[]>([]);
  const [savingAnno, setSavingAnno] = useState(false);
  const drawCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const drawingRef = useRef<Stroke | null>(null);
  const baseImgRef = useRef<HTMLImageElement | null>(null);
  const naturalSizeRef = useRef<{ w: number; h: number }>({ w: 0, h: 0 });
  const [view, setView] = useState<"map" | "locations">("map");
  // Locations map tour is prepared (not auto-played) after the flowchart.
  const [locReady, setLocReady] = useState(false);
  const [locMenuOpen, setLocMenuOpen] = useState(false);

  const containerRef = useRef<HTMLDivElement | null>(null);
  const mapRef = useRef<any>(null);
  const markersRef = useRef<any[]>([]);
  const infoRef = useRef<any>(null);
  const photoCacheRef = useRef<Record<string, string | null>>({});
  const focusRef = useRef<((i: number) => void) | null>(null);
  const focusLocationRef = useRef<((i: number, pan?: boolean) => void) | null>(null);
  const fitAllRef = useRef<(() => void) | null>(null);
  const focusedRef = useRef<number>(-1);
  const playingRef = useRef(false);
  const animTimerRef = useRef<any>(null);
  const gRef = useRef<any>(null);
  const routePathRef = useRef<any[]>([]);
  const routeLinesRef = useRef<any[]>([]);
  const activeTourRef = useRef<"report" | "location" | null>(null);
  const locAnimRef = useRef<any>(null);
  // Bumped on every stop/start so a tile warm-up that was aborted mid-flight
  // (user paused / restarted during the ~1.6s warm) can be told apart from the
  // current one and won't kick off a duplicate animation loop.
  const warmGenRef = useRef(0);

  // ---- Animated tours: focus each stop in order, 3 seconds per stop ----
  const SECONDS_PER_POINT = 3;

  const stopAnim = () => {
    playingRef.current = false;
    setPlaying(false);
    setBanner("");
    activeTourRef.current = null;
    warmGenRef.current += 1; // invalidate any in-flight tile warm-up

    if (animTimerRef.current) {
      clearTimeout(animTimerRef.current);
      animTimerRef.current = null;
    }
    if (locAnimRef.current?.raf) {
      cancelAnimationFrame(locAnimRef.current.raf);
      locAnimRef.current.raf = 0;
    }
  };

  // Pre-load the satellite tiles for a spot before the tour reveals it, so play
  // never starts on the bare (grey/empty) map. Centres+zooms the camera there
  // and resolves once the map fires `tilesloaded` — or after a short timeout so
  // a slow network can never stall the tour.
  const warmTiles = (lat: number, lng: number, zoom: number) =>
    new Promise<void>((resolve) => {
      const g = (window as any).google;
      const map = mapRef.current;
      if (!g?.maps?.event || !map) return resolve();
      let done = false;
      let listener: any = null;
      const finish = () => {
        if (done) return;
        done = true;
        try { listener?.remove?.(); } catch { /* ignore */ }
        resolve();
      };
      // Attach the listener BEFORE moving so the resulting tile load is caught.
      listener = g.maps.event.addListenerOnce(map, "tilesloaded", finish);
      try {
        map.setCenter({ lat, lng });
        map.setZoom(zoom);
      } catch { /* ignore */ }
      setTimeout(finish, 1600);
    });

  // Run an animated tour over a list of point indices.
  //   mode "report"   -> full report popup (category, observation, remarks, photo)
  //   mode "location" -> just the place-NAME popup (locations, not reports)
  const runTour = (indices: number[], mode: "report" | "location" = "report", from = 0) => {
    if (!indices.length) return;
    const fn = mode === "location" ? focusLocationRef.current : focusRef.current;
    if (!fn) return;
    stopAnim();
    activeTourRef.current = mode;
    playingRef.current = true;
    setPlaying(true);
    // Announce where the tour starts and ends so the direction is clear.
    const nameOf = (idx: number) =>
      points[idx]?.location || points[idx]?.category || `Point ${idx + 1}`;
    if (mode === "report") {
      setBanner(`▶ Start: ${nameOf(indices[0])}   →   End: ${nameOf(indices[indices.length - 1])}`);
    }
    let k = from < 0 ? 0 : from;
    const step = () => {
      if (!playingRef.current) return;
      fn(indices[k]);
      k += 1;
      if (k >= indices.length) {
        // Reached the end — STOP (no looping; auto-replay was confusing). Let
        // the last point dwell, then mark the tour complete.
        animTimerRef.current = setTimeout(() => {
          if (!playingRef.current) return;
          playingRef.current = false;
          setPlaying(false);
          activeTourRef.current = null;
          setBanner(
            `🏁 Tour complete — reached the end at ${nameOf(indices[indices.length - 1])}. Press ▶ to replay.`
          );
        }, SECONDS_PER_POINT * 1000);
        return;
      }
      animTimerRef.current = setTimeout(step, SECONDS_PER_POINT * 1000);
    };
    step();
  };

  // Build a pausable location tour: precompute the route geometry + the
  // distance at which the line reaches each location, draw the full line, and
  // frame the whole route. Does NOT start playing.
  const prepareLocationTour = () => {
    const g = gRef.current;
    const map = mapRef.current;
    const rawPath = routePathRef.current;
    const stops = locationIndices();
    locAnimRef.current = null;
    setLocReady(false);
    if (!(g && map && rawPath && rawPath.length > 1) || !stops.length) return;

    const path = rawPath.map((p: any) =>
      typeof p.lat === "function" ? { lat: p.lat(), lng: p.lng() } : { lat: p.lat, lng: p.lng }
    );
    const seg: number[] = [];
    let total = 0;
    for (let i = 1; i < path.length; i++) {
      const dy = path[i].lat - path[i - 1].lat;
      const dx = (path[i].lng - path[i - 1].lng) * Math.cos(((path[i].lat + path[i - 1].lat) / 2) * (Math.PI / 180));
      const d = Math.sqrt(dx * dx + dy * dy);
      seg.push(d);
      total += d;
    }
    if (total <= 0) return;
    const cum: number[] = [0];
    for (let i = 0; i < seg.length; i++) cum.push(cum[i] + seg[i]);
    const coords = stops.map((i) => ({ lat: points[i].lat, lng: points[i].lng }));
    const reachDist = coords.map((lc) => {
      let best = 0;
      let bd = Infinity;
      for (let i = 0; i < path.length; i++) {
        const dy = path[i].lat - lc.lat;
        const dx = (path[i].lng - lc.lng) * Math.cos(((path[i].lat + lc.lat) / 2) * (Math.PI / 180));
        const d = dx * dx + dy * dy;
        if (d < bd) { bd = d; best = i; }
      }
      return cum[best];
    });

    // Fresh polylines (full line visible). Reuse the shared store.
    for (const l of routeLinesRef.current) { try { l.setMap(null); } catch { /* ignore */ } }
    routeLinesRef.current = [];
    const casing = new g.maps.Polyline({ map, path, strokeColor: LINE_CASING, strokeWeight: 9, strokeOpacity: 1, zIndex: 1 });
    let icons: any;
    try {
      const ap = g?.maps?.SymbolPath?.FORWARD_CLOSED_ARROW;
      if (ap != null) icons = [{ icon: { path: ap, scale: 3, strokeColor: LINE_CASING, strokeWeight: 1, fillColor: "#fff", fillOpacity: 1 }, offset: "0", repeat: "90px" }];
    } catch { /* ignore */ }
    const lineOpts: any = { map, path, strokeColor: LINE_MAIN, strokeWeight: 5, strokeOpacity: 1, zIndex: 2 };
    if (icons) lineOpts.icons = icons;
    const line = new g.maps.Polyline(lineOpts);
    routeLinesRef.current = [casing, line];

    fitAllRef.current?.();
    locAnimRef.current = {
      path, seg, total, reachDist, fired: reachDist.map(() => false),
      // Slower crawl so each location can be read comfortably as the line
      // reaches it (was 12s / 2.6s-per-stop, which felt too fast).
      stops, casing, line, progress: 0, duration: Math.max(22000, stops.length * 5000),
      lastTs: 0, raf: 0,
    };
    setLocReady(true);
  };

  // Play / resume the location tour (pausable). The line draws slowly and each
  // location box pops when the line reaches it.
  const playLocationTour = () => {
    if (!locAnimRef.current) prepareLocationTour();
    const a = locAnimRef.current;
    if (!a) {
      runTour(locationIndices(), "location", 0); // fallback
      return;
    }
    stopAnim();
    activeTourRef.current = "location";
    playingRef.current = true;
    setPlaying(true);
    if (a.progress >= 1) {
      a.progress = 0;
      a.fired = a.fired.map(() => false);
    }
    const map = mapRef.current;
    if (a.progress === 0) {
      a.casing.setPath([a.path[0]]);
      a.line.setPath([a.path[0]]);
    }
    a.lastTs = 0;
    const tick = (ts: number) => {
      if (!playingRef.current || activeTourRef.current !== "location") return;
      if (!a.lastTs) a.lastTs = ts;
      const dt = ts - a.lastTs;
      a.lastTs = ts;
      a.progress = Math.min(1, a.progress + dt / a.duration);
      const target = a.progress * a.total;
      const out: Array<{ lat: number; lng: number }> = [a.path[0]];
      let acc = 0;
      for (let i = 0; i < a.seg.length; i++) {
        if (acc + a.seg[i] < target) { out.push(a.path[i + 1]); acc += a.seg[i]; }
        else {
          const fr = a.seg[i] > 0 ? (target - acc) / a.seg[i] : 0;
          out.push({ lat: a.path[i].lat + (a.path[i + 1].lat - a.path[i].lat) * fr, lng: a.path[i].lng + (a.path[i + 1].lng - a.path[i].lng) * fr });
          break;
        }
      }
      a.casing.setPath(out);
      a.line.setPath(out);
      // Smoothly follow the head of the line (live map).
      const head = out[out.length - 1];
      if (head && map) {
        try { map.setCenter(head); } catch { /* ignore */ }
      }
      for (let k = 0; k < a.reachDist.length; k++) {
        if (!a.fired[k] && target >= a.reachDist[k]) {
          a.fired[k] = true;
          focusLocationRef.current?.(a.stops[k], false);
        }
      }
      if (a.progress < 1) {
        a.raf = requestAnimationFrame(tick);
      } else {
        a.casing.setPath(a.path);
        a.line.setPath(a.path);
        playingRef.current = false;
        setPlaying(false);
        activeTourRef.current = null;
      }
    };
    // On a fresh start, warm the start tiles so the line doesn't begin drawing
    // over an empty grey map; otherwise resume ticking immediately.
    const beginTick = () => {
      a.lastTs = 0;
      a.raf = requestAnimationFrame(tick);
    };
    if (a.progress === 0 && map) {
      const startZoom = Math.min(16, Math.max((map.getZoom() || 12) + 2, 14));
      const gen = warmGenRef.current;
      warmTiles(a.path[0].lat, a.path[0].lng, startZoom).then(() => {
        if (warmGenRef.current === gen && playingRef.current && activeTourRef.current === "location") {
          beginTick();
        }
      });
    } else {
      beginTick();
    }
  };

  const toggleLocations = () => {
    if (playingRef.current && activeTourRef.current === "location") stopAnim();
    else playLocationTour();
  };

  const allIndices = () => points.map((_, i) => i);

  // Distinct main locations: collapse consecutive points with the same place
  // name into one stop (falls back to all points if no names are present).
  const locationIndices = () => {
    const out: number[] = [];
    let last = "";
    points.forEach((p, i) => {
      const loc = (p.location || "").toLowerCase();
      if (!loc) {
        out.push(i);
      } else if (loc !== last) {
        out.push(i);
        last = loc;
      }
    });
    return out.length ? out : allIndices();
  };

  const togglePlay = () => {
    if (playingRef.current && activeTourRef.current === "report") {
      stopAnim();
      return;
    }
    const idxs = allIndices();
    if (!idxs.length) return;
    // Resume from the current point, but if the tour already finished (we're at
    // or past the last point), replay from the start instead of the end.
    let start = current < 0 ? 0 : current;
    if (start >= idxs.length - 1) start = 0;
    const firstPt = points[idxs[start]] ?? points[idxs[0]];
    if (!firstPt) {
      runTour(idxs, "report", start);
      return;
    }
    // Flip the button to "playing" right away, warm the first point's tiles,
    // then start the tour so it never opens on an empty grey map. If the user
    // pauses during the (brief) warm-up, don't start.
    stopAnim();
    playingRef.current = true;
    setPlaying(true);
    activeTourRef.current = "report";
    const gen = warmGenRef.current;
    warmTiles(firstPt.lat, firstPt.lng, 16).then(() => {
      if (warmGenRef.current === gen && playingRef.current && activeTourRef.current === "report") {
        runTour(idxs, "report", start);
      }
    });
  };

  // Manual stepping: pause any auto-play and jump straight to one report.
  const jumpTo = (i: number) => {
    if (!points.length) return;
    const idx = Math.max(0, Math.min(points.length - 1, i));
    stopAnim();
    focusRef.current?.(idx);
  };
  const stepBack = () => jumpTo((current < 0 ? 0 : current) - 1);
  const stepForward = () => jumpTo((current < 0 ? -1 : current) + 1);
  const rewindStart = () => jumpTo(0);

  // ---- 1. Load report points ----
  useEffect(() => {
    // Share mode: wait for the session; bounce back to the password gate if
    // it's missing/expired. Normal mode: wait for the project id.
    if (isShare) {
      if (!shareSession) {
        router.replace(`/share/${encodeURIComponent(shareToken)}`);
        return;
      }
    } else if (!projectId) {
      return;
    }
    (async () => {
      setLoading(true);
      setError(null);
      try {
        const res = isShare
          ? await fetch(`/api/share/${encodeURIComponent(shareToken)}/reports`, {
              method: "GET",
              headers: { "X-Share-Session": shareSession },
            })
          : await fetch(
              `/api/projects/${encodeURIComponent(projectId)}/reports?sort=asc`,
              { method: "GET", credentials: "include", headers: authHeaders() }
            );
        const data = await res.json().catch(() => ({} as any));
        if (res.status === 401 && isShare) {
          router.replace(`/share/${encodeURIComponent(shareToken)}`);
          return;
        }
        if (!res.ok) throw new Error(data?.error || "Failed to load reports");

        let rows: any[] = Array.isArray(data?.reports) ? data.reports : [];
        // In share mode the server already returns exactly the shared set.
        if (!isShare && selectedIds.length) {
          const set = new Set(selectedIds);
          rows = rows.filter((r) => set.has(String(r.id)));
        }

        const pts: ReportPoint[] = [];
        let n = 0;
        for (const r of rows) {
          const lat = Number(r.latitude ?? r.loc_lat);
          const lng = Number(r.longitude ?? r.loc_lon);
          if (!Number.isFinite(lat) || !Number.isFinite(lng)) continue;
          if (Math.abs(lat) > 90 || Math.abs(lng) > 180) continue;
          if (String(r.point_key || "") === "__NO_GPS__") continue;
          n += 1;
          pts.push({
            id: String(r.id),
            n,
            lat,
            lng,
            category: String(r.category || "Report"),
            description: String(r.description || ""),
            remarks: String(r.remarks_action || r.remarks || ""),
            location: placeName(
              r.resolved_location || r.location || r.location_name || r.address
            ),
            state: stateName(
              r.resolved_location || r.location || r.location_name || r.address
            ),
          });
        }
        setPoints(pts);
        if (!pts.length) setError("None of these reports have GPS coordinates to map.");
        // Preload each point's first photo (concurrency-limited) so the popup
        // shows the image instantly during the flythrough.
        preloadPhotos(pts.map((p) => p.id));
      } catch (e: any) {
        setError(e?.message || "Failed to load reports");
      } finally {
        setLoading(false);
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [projectId, selectedIds, isShare, shareToken, shareSession]);

  // Preload first photos for many reports with limited concurrency.
  const preloadPhotos = async (ids: string[]) => {
    const queue = [...ids];
    const worker = async () => {
      while (queue.length) {
        const id = queue.shift()!;
        // eslint-disable-next-line no-await-in-loop
        await fetchFirstPhoto(id);
      }
    };
    await Promise.all(Array.from({ length: 6 }, () => worker()));
  };

  // Lazy-load the first photo for a report (for the info window).
  const fetchFirstPhoto = async (reportId: string): Promise<string | null> => {
    if (reportId in photoCacheRef.current) return photoCacheRef.current[reportId];
    try {
      const res = isShare
        ? await fetch(
            `/api/share/${encodeURIComponent(shareToken)}/reports/${encodeURIComponent(
              reportId
            )}/photos`,
            { method: "GET", headers: { "X-Share-Session": shareSession } }
          )
        : await fetch(`/api/reports/${encodeURIComponent(reportId)}/photos`, {
            method: "GET",
            credentials: "include",
            headers: authHeaders(),
          });
      const data = await res.json().catch(() => ({} as any));
      const url = String(data?.photos?.[0]?.url || "") || null;
      photoCacheRef.current[reportId] = url;
      return url;
    } catch {
      photoCacheRef.current[reportId] = null;
      return null;
    }
  };

  // Fetch ALL photo URLs for a report (for the lightbox carousel).
  const fetchAllPhotos = async (reportId: string): Promise<string[]> => {
    try {
      const res = isShare
        ? await fetch(
            `/api/share/${encodeURIComponent(shareToken)}/reports/${encodeURIComponent(
              reportId
            )}/photos`,
            { method: "GET", headers: { "X-Share-Session": shareSession } }
          )
        : await fetch(`/api/reports/${encodeURIComponent(reportId)}/photos`, {
            method: "GET",
            credentials: "include",
            headers: authHeaders(),
          });
      const data = await res.json().catch(() => ({} as any));
      const list: string[] = Array.isArray(data?.photos)
        ? data.photos.map((x: any) => String(x?.url || "")).filter(Boolean)
        : [];
      return list;
    } catch {
      return [];
    }
  };

  // When the lightbox opens, load every photo for that report so the user can
  // page through them with prev/next. Starts on the photo that was clicked.
  useEffect(() => {
    if (!lightbox) {
      setLbPhotos([]);
      setLbIndex(0);
      return;
    }
    let cancelled = false;
    (async () => {
      const urls = await fetchAllPhotos(lightbox.reportId);
      if (cancelled) return;
      const list = urls.length ? urls : lightbox.url ? [lightbox.url] : [];
      const clickedAt = list.indexOf(lightbox.url);
      setLbPhotos(list);
      setLbIndex(clickedAt >= 0 ? clickedAt : 0);
    })();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lightbox]);

  // ---- Annotation drawing ----
  // Reset annotations whenever the shown photo changes or the lightbox closes.
  useEffect(() => {
    setStrokes([]);
    drawingRef.current = null;
    setDrawMode(false);
  }, [lightbox, lbIndex]);

  // Directional turn arrows (left / right / U-turn) drawn to fit the dragged
  // box, like a road-survey turn marker.
  const drawTurnArrow = (
    ctx: CanvasRenderingContext2D,
    type: "left" | "right" | "uturn",
    a: { x: number; y: number },
    b: { x: number; y: number },
    color: string,
    width: number
  ) => {
    const L = Math.min(a.x, b.x);
    const R = Math.max(a.x, b.x);
    const T = Math.min(a.y, b.y);
    const B = Math.max(a.y, b.y);
    const W = R - L;
    const H = B - T;
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
      const lx = cx - legOffset;
      const rx = cx + legOffset;
      const topY = T + legOffset;
      ctx.beginPath();
      ctx.moveTo(lx, B);
      ctx.lineTo(lx, topY);
      ctx.arc(cx, topY, legOffset, Math.PI, 2 * Math.PI, false); // up and over the top
      ctx.lineTo(rx, B - head);
      ctx.stroke();
      headAt(rx, B, Math.PI / 2); // arrowhead points down
      return;
    }
    // 90° left / right turn: vertical shaft, then curve out to the side.
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
      // Subtle dark outline so light text stays readable on a bright photo.
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
    const a = pts[0];
    const b = pts[pts.length - 1];
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
      // Cross drawn across the dragged box.
      ctx.beginPath();
      ctx.moveTo(a.x, a.y);
      ctx.lineTo(b.x, b.y);
      ctx.moveTo(b.x, a.y);
      ctx.lineTo(a.x, b.y);
      ctx.stroke();
      return;
    }
    // arrow: shaft + filled head
    ctx.beginPath();
    ctx.moveTo(a.x, a.y);
    ctx.lineTo(b.x, b.y);
    ctx.stroke();
    const angle = Math.atan2(b.y - a.y, b.x - a.x);
    const head = Math.max(14, s.width * 3.5);
    ctx.beginPath();
    ctx.moveTo(b.x, b.y);
    ctx.lineTo(b.x - head * Math.cos(angle - Math.PI / 7), b.y - head * Math.sin(angle - Math.PI / 7));
    ctx.lineTo(b.x - head * Math.cos(angle + Math.PI / 7), b.y - head * Math.sin(angle + Math.PI / 7));
    ctx.closePath();
    ctx.fill();
  };

  const redrawCanvas = (extra?: Stroke | null) => {
    const canvas = drawCanvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    if (!ctx) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    const base = baseImgRef.current;
    if (base) ctx.drawImage(base, 0, 0, canvas.width, canvas.height);
    for (const s of strokes) drawOneStroke(ctx, s);
    if (extra) drawOneStroke(ctx, extra);
  };

  // Load the current photo into the canvas when draw mode opens. crossOrigin
  // lets us export the canvas if the host allows CORS; if not, Save surfaces
  // a clear error rather than silently failing.
  useEffect(() => {
    if (!drawMode || !lightbox) return;
    const rawUrl = lbPhotos[lbIndex] || lightbox.url;
    if (!rawUrl) return;
    // Cross-origin photos (S3) taint the canvas and block toBlob on Save. Route
    // them through the same-origin proxy so the photo loads AND stays exportable.
    const origin = typeof window !== "undefined" ? window.location.origin : "";
    const isExternal = /^https?:\/\//i.test(rawUrl) && !rawUrl.startsWith(origin);
    const url = isExternal ? `/api/image-proxy?url=${encodeURIComponent(rawUrl)}` : rawUrl;
    let cancelled = false;
    const img = new Image();
    img.onload = () => {
      if (cancelled) return;
      baseImgRef.current = img;
      const w = img.naturalWidth || 1200;
      const h = img.naturalHeight || 800;
      naturalSizeRef.current = { w, h };
      const canvas = drawCanvasRef.current;
      if (canvas) {
        canvas.width = w;
        canvas.height = h;
        // Fit within the viewport, preserving aspect, so client→canvas pointer
        // math has an exact uniform scale with no letterbox offset.
        const maxW = Math.min(window.innerWidth * 0.62, 1080);
        const maxH = window.innerHeight * 0.85;
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
  }, [drawMode, lbIndex, lightbox]);

  // Redraw when committed strokes change while drawing.
  useEffect(() => {
    if (drawMode) redrawCanvas();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [strokes, drawMode]);

  const canvasPoint = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const canvas = drawCanvasRef.current!;
    const rect = canvas.getBoundingClientRect();
    return {
      x: ((e.clientX - rect.left) / rect.width) * canvas.width,
      y: ((e.clientY - rect.top) / rect.height) * canvas.height,
    };
  };
  const onDrawDown = (e: React.PointerEvent<HTMLCanvasElement>) => {
    if (!drawMode) return;
    e.preventDefault();
    const p = canvasPoint(e);
    if (drawTool === "text") {
      // Click to place text; prompt for the string. Size scales with the
      // chosen width so it reads on a high-resolution photo.
      const txt = window.prompt("Text to add:");
      if (txt && txt.trim()) {
        const fontSize = Math.max(28, drawWidth * 7);
        setStrokes((prev) => [
          ...prev,
          { tool: "text", color: drawColor, width: drawWidth, points: [p], text: txt.trim(), fontSize },
        ]);
      }
      return; // no drag stroke for text
    }
    (e.target as HTMLCanvasElement).setPointerCapture(e.pointerId);
    drawingRef.current = { tool: drawTool, color: drawColor, width: drawWidth, points: [p, p] };
    redrawCanvas(drawingRef.current);
  };
  const onDrawMove = (e: React.PointerEvent<HTMLCanvasElement>) => {
    if (!drawMode || !drawingRef.current) return;
    const p = canvasPoint(e);
    if (drawTool === "pen") drawingRef.current.points.push(p);
    else drawingRef.current.points[1] = p;
    redrawCanvas(drawingRef.current);
  };
  const onDrawUp = () => {
    if (!drawingRef.current) return;
    const s = drawingRef.current;
    drawingRef.current = null;
    setStrokes((prev) => [...prev, s]);
  };

  const saveAnnotated = async () => {
    const canvas = drawCanvasRef.current;
    if (!canvas || !lightbox || savingAnno) return;
    setSavingAnno(true);
    try {
      redrawCanvas(); // bake in the latest strokes
      const blob: Blob | null = await new Promise((resolve) =>
        canvas.toBlob((b) => resolve(b), "image/jpeg", 0.92)
      );
      if (!blob)
        throw new Error(
          "Couldn't export the annotated image — the photo host is blocking cross-origin canvas export."
        );
      const fd = new FormData();
      fd.append("file", blob, `annotated_${Date.now()}.jpg`);
      fd.append("folder", "uploads");
      fd.append("reportId", lightbox.reportId);
      fd.append("width", String(canvas.width));
      fd.append("height", String(canvas.height));
      const res = await fetch("/api/upload", {
        method: "POST",
        credentials: "include",
        headers: authHeaders(),
        body: fd,
      });
      const data = await res.json().catch(() => ({} as any));
      if (!res.ok) throw new Error(data?.error || "Upload failed");
      const newUrl = String(data?.url || data?.photo?.url || "");
      const urls = await fetchAllPhotos(lightbox.reportId);
      setStrokes([]);
      setDrawMode(false);
      if (urls.length) {
        setLbPhotos(urls);
        const at = newUrl ? urls.indexOf(newUrl) : -1;
        setLbIndex(at >= 0 ? at : urls.length - 1);
      }
    } catch (err) {
      alert((err as { message?: string })?.message || "Failed to save the annotated image.");
    } finally {
      setSavingAnno(false);
    }
  };

  // Bridge clicks on the InfoWindow image (raw HTML) to the React lightbox.
  useEffect(() => {
    (window as any).__route3dLightbox = (url: string, reportId: string) => {
      // Pause the tour and close Google's info window so the lightbox (with
      // photo + details) isn't hidden behind it.
      stopAnim();
      infoRef.current?.close?.();
      setLightbox({ url, reportId });
    };
    return () => {
      delete (window as any).__route3dLightbox;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const infoHtml = (p: ReportPoint, photoUrl: string | null) => {
    // onerror hides the image if the URL is broken — otherwise a blank
    // 150px white box is left at the bottom of the popup.
    const img = photoUrl
      ? `<img src="${photoUrl}" onclick="window.__route3dLightbox && window.__route3dLightbox('${photoUrl}','${p.id}')" onerror="this.style.display='none'" title="Click to enlarge" style="width:100%;max-width:260px;height:150px;object-fit:cover;border-radius:8px;margin-top:6px;cursor:pointer;background:#f1f5f9"/>`
      : "";
    return `
      <div style="font-family:system-ui,Segoe UI,Arial;max-width:280px">
        <div style="display:flex;gap:8px;align-items:center">
          <span style="background:#d50000;color:#fff;border-radius:999px;width:22px;height:22px;display:inline-flex;align-items:center;justify-content:center;font-weight:800;font-size:12px">${p.n}</span>
          <b style="font-size:14px;color:#0f172a">${escapeHtml(p.category)}</b>
        </div>
        ${p.description ? `<div style="font-size:13px;font-weight:700;color:#0f172a;margin-top:6px"><b>Observation:</b> ${escapeHtml(p.description)}</div>` : ""}
        ${p.remarks ? `<div style="font-size:13px;font-weight:700;color:#0f172a;margin-top:4px"><b>Remarks / Action:</b> ${escapeHtml(p.remarks)}</div>` : ""}
        <div style="font-size:11px;color:#64748b;margin-top:6px">${p.lat.toFixed(5)}, ${p.lng.toFixed(5)}</div>
        ${img}
      </div>`;
  };

  // ---- 2. Init the satellite map once points are loaded ----
  useEffect(() => {
    if (!GMAPS_KEY || !points.length || !containerRef.current) return;
    let cancelled = false;

    (async () => {
      try {
        await loadGoogleMaps(GMAPS_KEY);
        if (cancelled) return;
        const g = (window as any).google;
        await g.maps.importLibrary("maps");
        if (cancelled) return;

        const bounds = new g.maps.LatLngBounds();
        points.forEach((p) => bounds.extend({ lat: p.lat, lng: p.lng }));

        const map = new g.maps.Map(containerRef.current, {
          mapTypeId: "hybrid", // satellite + labels, like the reference image
          center: bounds.getCenter(),
          zoom: 10,
          tilt: 0,
          mapTypeControl: true,
          streetViewControl: false,
          fullscreenControl: true,
          // Colour shown while satellite tiles are still loading (when the
          // camera pans/zooms to a new point). A dark backdrop blends with the
          // imagery instead of flashing an ugly light-grey box.
          backgroundColor: "#0f1e2b",
        });
        map.fitBounds(bounds, 60);
        mapRef.current = map;

        // Re-fit to the whole route (end-to-end overview) on demand.
        fitAllRef.current = () => {
          const b = new g.maps.LatLngBounds();
          points.forEach((p) => b.extend({ lat: p.lat, lng: p.lng }));
          map.fitBounds(b, 60);
        };

        // disableAutoPan keeps the map steady — otherwise opening a popup
        // pans/"jumps" the camera to fit the info window.
        const info = new g.maps.InfoWindow({ disableAutoPan: true });
        infoRef.current = info;

        // Focus helper: open details + zoom in to the point.
        const focus = async (i: number) => {
          const p = points[i];
          if (!p) return;
          focusedRef.current = i;
          setCurrent(i);
          map.panTo({ lat: p.lat, lng: p.lng });
          map.setZoom(Math.max(map.getZoom() || 0, 16));
          // Shift the view so the marker sits below centre — the tall popup
          // (photo + remarks) then fits on screen instead of being clipped
          // at the top edge.
          map.panBy(0, -140);
          info.setContent(infoHtml(p, photoCacheRef.current[p.id] ?? null));
          info.open(map, markersRef.current[i]);
          const url = await fetchFirstPhoto(p.id);
          // Only update if this point is still the one being shown.
          if (!cancelled && url && focusedRef.current === i) {
            info.setContent(infoHtml(p, url));
          }
        };
        focusRef.current = focus;

        // Location focus: pan/zoom and show only the place NAME (no report
        // details). Used by the Locations tour after the flowchart.
        const focusLocation = (i: number, pan = true) => {
          const p = points[i];
          if (!p) return;
          focusedRef.current = i;
          setCurrent(i);
          if (pan) {
            map.panTo({ lat: p.lat, lng: p.lng });
            map.setZoom(Math.max(map.getZoom() || 0, 14));
            // Small downward shift so the name box isn't clipped at the top.
            map.panBy(0, -60);
          }
          // Show the real place name; if a point has none, fall back to its
          // category, then a neutral "Stop N" — never the bare word "Location".
          const name = p.location || p.category || `Stop ${p.n}`;
          // Render the popup as a flowchart-style box: coloured rounded card
          // with a numbered white circle + the place name in white. Colour +
          // number come from the SAME per-city scheme as the flowchart, so each
          // city shows its one matching colour.
          const { fill, border, num } = cityInfoForPoint(points, i);
          info.setContent(
            `<div style="display:flex;gap:10px;align-items:center;background:${fill};border:2px solid ${border};padding:10px 16px;border-radius:14px;font-family:system-ui,Segoe UI,Arial;white-space:nowrap;line-height:1.1">
               <span style="background:#fff;color:${border};width:26px;height:26px;border-radius:999px;display:inline-flex;align-items:center;justify-content:center;font-weight:800;font-size:13px">${num}</span>
               <b style="color:#fff;font-size:16px;font-weight:800">${escapeHtml(name)}</b>
             </div>`
          );
          info.open(map, markersRef.current[i]);
        };
        focusLocationRef.current = focusLocation;

        // Numbered markers.
        markersRef.current = points.map((p, i) => {
          const isStart = i === 0;
          const isEnd = i === points.length - 1;
          const opts: any = {
            position: { lat: p.lat, lng: p.lng },
            map,
            label: { text: String(p.n), color: "#fff", fontWeight: "800", fontSize: "12px" },
            title: isStart
              ? `START — ${p.category}`
              : isEnd
                ? `END — ${p.category}`
                : p.category,
          };
          // Mark the route's start (green) and end (red) with a distinct
          // coloured circle so it's obvious where the tour begins and finishes.
          if (isStart || isEnd) {
            opts.icon = {
              path: g.maps.SymbolPath.CIRCLE,
              scale: 13,
              fillColor: isStart ? "#16a34a" : "#dc2626",
              fillOpacity: 1,
              strokeColor: "#ffffff",
              strokeWeight: 3,
            };
            opts.zIndex = 1000;
          }
          const marker = new g.maps.Marker(opts);
          marker.addListener("click", () => {
            stopAnim();
            focus(i);
          });
          return marker;
        });

        // Route line through the points — road-following via Directions when
        // possible. Store g/path/lines so the line can be redrawn slowly
        // during the Locations tour.
        gRef.current = g;
        const routeInfo = await drawRoute(
          g,
          map,
          points,
          routeLinesRef,
          (p) => {
            routePathRef.current = p;
          },
          fetchRoadPath
        );

        if (cancelled) return;
        // Explain a straight route when it happens. Roads come from Google when
        // its billing is on, otherwise from the free OpenRouteService proxy.
        if (!routeInfo.road) {
          setRouteHint(
            "Showing a straight route — road directions aren’t available yet. Add your free OpenRouteService key (ORS_API_KEY) on the server to follow roads, or enable Google billing."
          );
        } else if (routeInfo.straight) {
          setRouteHint("Most of the route follows roads; a few sections with no drivable road are shown straight.");
        } else {
          setRouteHint(null);
        }
        setReady(true);
      } catch (e: any) {
        console.error("[routemap] init failed:", e);
        if (!cancelled) {
          setMapError(
            e?.message ||
              "Failed to load Google Maps. Check the API key has Maps JavaScript API enabled with billing on."
          );
        }
      }
    })();

    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [points]);

  // Stop any running animation on unmount. We do NOT auto-play — the map
  // opens fitted to the whole route end-to-end, and the user presses Play.
  useEffect(() => {
    return () => stopAnim();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Create a password-protected share link for exactly the points now on the
  // map, so a client can watch the animated report without an app login.
  const createShare = async () => {
    if (shareBusy) return;
    if (!projectId) {
      setShareErr("No project to share.");
      return;
    }
    if (sharePassword.trim().length < 4) {
      setShareErr("Password must be at least 4 characters.");
      return;
    }
    setShareBusy(true);
    setShareErr(null);
    try {
      const res = await fetch(`/api/projects/${encodeURIComponent(projectId)}/share`, {
        method: "POST",
        headers: { "Content-Type": "application/json", ...authHeaders() },
        credentials: "include",
        body: JSON.stringify({
          password: sharePassword,
          title: shareTitle,
          // Lock the link to exactly the points currently shown.
          selectedIds: points.map((p) => p.id),
        }),
      });
      const data = await res.json().catch(() => ({} as any));
      if (!res.ok || !data?.ok) throw new Error(data?.error || "Failed to create link");
      const link =
        typeof window !== "undefined" ? `${window.location.origin}${data.path}` : String(data.path);
      setShareLink(link);
      setShareCopied(false);
    } catch (e: any) {
      setShareErr(e?.message || "Failed to create link");
    } finally {
      setShareBusy(false);
    }
  };

  const copyShareLink = async () => {
    if (!shareLink) return;
    try {
      await navigator.clipboard.writeText(shareLink);
      setShareCopied(true);
      setTimeout(() => setShareCopied(false), 2000);
    } catch {
      /* clipboard may be blocked; the field is selectable as a fallback */
    }
  };

  const resetShareDialog = () => {
    setShareOpen(false);
    setSharePassword("");
    setShareTitle("");
    setShareErr(null);
    setShareLink("");
    setShareCopied(false);
  };

  // ---- Render ----
  if (!GMAPS_KEY) {
    return (
      <div style={styles.page}>
        <TopBar onBack={isShare ? undefined : () => router.back()} title="Route Map (satellite)" />
        <div style={styles.card}>
          <div style={styles.h2}>Google Maps API key required</div>
          <p style={styles.p}>
            Add a key to <code>.env.local</code> and restart the dev server:
          </p>
          <pre style={styles.pre}>NEXT_PUBLIC_GOOGLE_MAPS_API_KEY=your_key_here</pre>
          <p style={styles.p}>
            In Google Cloud Console enable <b>Maps JavaScript API</b> (and{" "}
            <b>Directions API</b> for road-following routes), enable billing, and
            restrict the key to your domain.
          </p>
        </div>
      </div>
    );
  }

  return (
    <div style={styles.page}>
      <style>{`
        .gm-style-iw.gm-style-iw-c { padding: 8px !important; border-radius: 12px !important; }
        .gm-style-iw-d { overflow: hidden !important; }
        .gm-style-iw-tail, .gm-style-iw-t::after { }
      `}</style>
      <TopBar
        onBack={isShare ? undefined : () => router.back()}
        title="Route Map (satellite)"
        right={
          isShare ? null : (
            <button
              style={styles.shareBtn}
              onClick={() => setShareOpen(true)}
              title="Create a password-protected link to send to a client"
            >
              🔗 Share
            </button>
          )
        }
      />

      <div style={styles.layout}>
        <div style={styles.mapWrap}>
          {loading ? (
            <div style={styles.overlay}>Loading reports…</div>
          ) : error ? (
            <div style={styles.overlay}>{error}</div>
          ) : mapError ? (
            <div style={styles.overlay}>{mapError}</div>
          ) : null}
          <div ref={containerRef} style={{ width: "100%", height: "100%", background: "#0f1e2b" }} />

          {ready && banner ? (
            <div style={styles.banner}>{banner}</div>
          ) : null}

          {ready && routeHint ? (
            <div style={styles.routeHint}>
              <span style={{ flex: 1, lineHeight: 1.35 }}>🛣️ {routeHint}</span>
              <button
                style={styles.routeHintClose}
                onClick={() => setRouteHint(null)}
                title="Dismiss"
              >
                ✕
              </button>
            </div>
          ) : null}

          {ready && (
            <div style={styles.controls}>
              <button style={styles.ctrlBtn} onClick={togglePlay}>
                {playing && activeTourRef.current === "report" ? "⏸ Pause" : "▶ Play reports"}
              </button>

              {/* Manual playback: rewind to start, step back, step forward. */}
              <button
                style={styles.ctrlIconBtn}
                onClick={rewindStart}
                title="Rewind to the first report"
                aria-label="Rewind to first report"
              >
                ⏮
              </button>
              <button
                style={styles.ctrlIconBtn}
                onClick={stepBack}
                disabled={current <= 0}
                title="Previous report"
                aria-label="Previous report"
              >
                ◀
              </button>
              <button
                style={styles.ctrlIconBtn}
                onClick={stepForward}
                disabled={current >= 0 && current >= points.length - 1}
                title="Next report"
                aria-label="Next report"
              >
                ▶
              </button>

              {playing && activeTourRef.current === "location" ? (
                <button
                  style={{ ...styles.ctrlBtn, background: "#0e7490" }}
                  onClick={stopAnim}
                  title="Pause the locations tour"
                >
                  ⏸ Pause
                </button>
              ) : (
                <div style={{ position: "relative" }}>
                  <button
                    style={{ ...styles.ctrlBtn, background: "#0e7490" }}
                    onClick={() => setLocMenuOpen((v) => !v)}
                    title="Locations options"
                  >
                    📍 Locations ▾
                  </button>
                  {locMenuOpen && (
                    <div style={styles.locMenu}>
                      <button
                        style={styles.locMenuItem}
                        onClick={() => {
                          setLocMenuOpen(false);
                          stopAnim();
                          setView("locations");
                        }}
                      >
                        🗺 Flowchart
                      </button>
                      <button
                        style={styles.locMenuItem}
                        onClick={() => {
                          setLocMenuOpen(false);
                          playLocationTour();
                        }}
                      >
                        ▶ Play on map
                      </button>
                    </div>
                  )}
                </div>
              )}

              <button
                style={styles.ctrlBtn}
                onClick={() => runTour(allIndices(), "report", 0)}
                title="Restart the reports tour from the first point"
              >
                ⟲ Restart
              </button>
              <button
                style={{ ...styles.ctrlBtn, background: "#334155" }}
                onClick={() => {
                  stopAnim();
                  setCurrent(-1);
                  infoRef.current?.close?.();
                  fitAllRef.current?.();
                }}
                title="Show the whole route end-to-end"
              >
                ⤢ Fit route
              </button>
              <span style={styles.ctrlInfo}>
                {current >= 0 ? `Point ${current + 1} / ${points.length}` : `${points.length} points`}
                {playing ? " • 3s each" : ""}
              </span>
            </div>
          )}

          {view === "locations" && (
            <LocationsFlow
              points={points}
              onClose={() => {
                setView("map");
                fitAllRef.current?.();
                infoRef.current?.close?.();
              }}
              onComplete={() => {
                // After the flowchart animation, return to the map and PREPARE
                // the locations tour (line drawn, route framed) but DON'T auto-
                // play — the user presses "Play locations" to start it.
                setView("map");
                setTimeout(() => prepareLocationTour(), 120);
              }}
            />
          )}
        </div>

        <div style={styles.side}>
          <div style={styles.sideHead}>Points ({points.length})</div>
          <div style={styles.list}>
            {points.map((p, i) => (
              <button
                key={p.id}
                style={{
                  ...styles.listItem,
                  background: i === current ? "#EFF6FF" : "#fff",
                  borderColor: i === current ? "#1d4ed8" : "#EAECF0",
                }}
                onClick={() => {
                  stopAnim();
                  focusRef.current?.(i);
                }}
                disabled={!ready}
              >
                <span style={styles.badge}>{p.n}</span>
                <span style={styles.listText}>
                  <b>{p.category}</b>
                  {p.description ? (
                    <span style={styles.listSub}>{p.description.slice(0, 60)}</span>
                  ) : null}
                </span>
              </button>
            ))}
          </div>
        </div>
      </div>

      {lightbox && (() => {
        const lp = points.find((p) => p.id === lightbox.reportId) || null;
        const total = lbPhotos.length;
        const curUrl = lbPhotos[lbIndex] || lightbox.url;
        return (
          <div style={styles.lightbox} onClick={() => setLightbox(null)}>
            <div style={styles.lightboxCard} onClick={(e) => e.stopPropagation()}>
              <div style={styles.lightboxImgWrap}>
                {drawMode ? (
                  <canvas
                    ref={drawCanvasRef}
                    style={styles.drawCanvas}
                    onPointerDown={onDrawDown}
                    onPointerMove={onDrawMove}
                    onPointerUp={onDrawUp}
                    onPointerLeave={onDrawUp}
                  />
                ) : (
                  <img src={curUrl} alt="" style={styles.lightboxImg} />
                )}
                {total > 1 && !drawMode ? (
                  <>
                    <button
                      style={{ ...styles.lbNav, ...styles.lbNavPrev }}
                      onClick={(e) => {
                        e.stopPropagation();
                        setLbIndex((i) => (i - 1 + total) % total);
                      }}
                      aria-label="Previous photo"
                    >
                      ‹
                    </button>
                    <button
                      style={{ ...styles.lbNav, ...styles.lbNavNext }}
                      onClick={(e) => {
                        e.stopPropagation();
                        setLbIndex((i) => (i + 1) % total);
                      }}
                      aria-label="Next photo"
                    >
                      ›
                    </button>
                    <div style={styles.lbCounter}>
                      {lbIndex + 1} / {total}
                    </div>
                  </>
                ) : null}
              </div>

              <div style={styles.lightboxInfo}>
                <div style={styles.lightboxHeader}>
                  {lp ? <span style={styles.lightboxBadge}>{lp.n}</span> : null}
                  <div style={styles.lightboxTitle}>{lp?.category || "Report"}</div>
                </div>

                <div style={styles.lightboxBody}>
                  {lp?.description ? (
                    <div style={styles.lightboxField}>
                      <div style={styles.lightboxLabel}>Observation</div>
                      <div style={styles.lightboxValue}>{lp.description}</div>
                    </div>
                  ) : null}
                  {lp?.remarks ? (
                    <div style={styles.lightboxField}>
                      <div style={styles.lightboxLabel}>Remarks / Action</div>
                      <div style={styles.lightboxValue}>{lp.remarks}</div>
                    </div>
                  ) : null}
                  {lp ? (
                    <div style={styles.lightboxCoords}>
                      📍 {lp.lat.toFixed(5)}, {lp.lng.toFixed(5)}
                    </div>
                  ) : null}
                </div>

                {!isShare ? (
                  <div style={styles.drawPanel}>
                    {!drawMode ? (
                      <button style={styles.drawBtn} onClick={() => setDrawMode(true)}>
                        ✏️ Draw on photo
                      </button>
                    ) : (
                      <>
                        <div style={styles.drawTools}>
                          {([
                            ["arrow", "➔", "Arrow"],
                            ["pen", "✎", "Freehand"],
                            ["rect", "▭", "Rectangle"],
                            ["ellipse", "◯", "Ellipse"],
                            ["left", "↰", "Left turn"],
                            ["right", "↱", "Right turn"],
                            ["uturn", "↩", "U-turn"],
                            ["x", "✕", "X mark"],
                            ["text", "T", "Text"],
                          ] as const).map(([t, icon, label]) => (
                            <button
                              key={t}
                              onClick={() => setDrawTool(t)}
                              style={{
                                ...styles.toolBtn,
                                ...(drawTool === t ? styles.toolBtnActive : {}),
                              }}
                              title={label}
                            >
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
                              style={{
                                ...styles.swatch,
                                background: c,
                                outline:
                                  drawColor.toLowerCase() === c.toLowerCase()
                                    ? "2px solid #0f172a"
                                    : "2px solid transparent",
                              }}
                            />
                          ))}
                          <input
                            type="color"
                            value={drawColor}
                            onChange={(e) => setDrawColor(e.target.value)}
                            style={styles.colorInput}
                            title="Custom colour"
                          />
                          <select
                            value={drawWidth}
                            onChange={(e) => setDrawWidth(Number(e.target.value))}
                            style={styles.widthSelect}
                            title="Line width"
                          >
                            <option value={3}>Thin</option>
                            <option value={6}>Medium</option>
                            <option value={10}>Thick</option>
                          </select>
                        </div>
                        <div style={styles.drawActions}>
                          <button
                            style={styles.drawBtnGhost}
                            onClick={() => setStrokes((s) => s.slice(0, -1))}
                            disabled={!strokes.length}
                          >
                            Undo
                          </button>
                          <button
                            style={styles.drawBtnGhost}
                            onClick={() => setStrokes([])}
                            disabled={!strokes.length}
                          >
                            Clear
                          </button>
                          <button
                            style={styles.drawBtnGhost}
                            onClick={() => {
                              setStrokes([]);
                              setDrawMode(false);
                            }}
                          >
                            Cancel
                          </button>
                          <button
                            style={styles.drawBtnPrimary}
                            onClick={saveAnnotated}
                            disabled={savingAnno || !strokes.length}
                          >
                            {savingAnno ? "Saving…" : "Save"}
                          </button>
                        </div>
                      </>
                    )}
                  </div>
                ) : null}
              </div>

              <button style={styles.lightboxClose} onClick={() => setLightbox(null)}>
                ✕
              </button>
            </div>
          </div>
        );
      })()}

      {shareOpen && !isShare && (
        <div style={styles.modalOverlay} onClick={resetShareDialog}>
          <div style={styles.modalCard} onClick={(e) => e.stopPropagation()}>
            <div style={styles.modalTitle}>🔗 Share animated report</div>
            {!shareLink ? (
              <>
                <div style={styles.modalSub}>
                  Creates a password-protected link. The client opens it, types the password, and
                  watches this report ({points.length} points) — no app login needed.
                </div>
                <label style={styles.modalLabel}>Label (optional)</label>
                <input
                  style={styles.modalInput}
                  value={shareTitle}
                  onChange={(e) => setShareTitle(e.target.value)}
                  placeholder="e.g. Delhi–Bhubaneswar route"
                />
                <label style={styles.modalLabel}>Password for the client</label>
                <input
                  style={styles.modalInput}
                  type="text"
                  value={sharePassword}
                  onChange={(e) => setSharePassword(e.target.value)}
                  placeholder="Set a password (min 4 characters)"
                />
                {shareErr ? <div style={styles.modalErr}>{shareErr}</div> : null}
                <div style={styles.modalActions}>
                  <button style={styles.modalCancel} onClick={resetShareDialog}>
                    Cancel
                  </button>
                  <button style={styles.modalCreate} onClick={createShare} disabled={shareBusy}>
                    {shareBusy ? "Creating…" : "Create link"}
                  </button>
                </div>
              </>
            ) : (
              <>
                <div style={styles.modalSub}>
                  Send this link to your client, and give them the password separately (phone / SMS).
                  It stays private until the password is entered.
                </div>
                <div style={styles.linkRow}>
                  <input
                    style={styles.linkInput}
                    readOnly
                    value={shareLink}
                    onFocus={(e) => e.currentTarget.select()}
                  />
                  <button style={styles.copyBtn} onClick={copyShareLink}>
                    {shareCopied ? "✓ Copied" : "Copy"}
                  </button>
                </div>
                <div style={styles.pwReminder}>
                  Password: <b>{sharePassword}</b>
                </div>
                <div style={styles.modalActions}>
                  <button style={styles.modalCreate} onClick={resetShareDialog}>
                    Done
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

// Google-directions style: a darker blue casing under a lighter blue line.
const LINE_CASING = "#1a46c8"; // darker outline
const LINE_MAIN = "#8ab4f8"; // light blue top

// Animate a blue polyline drawing smoothly from start to end along `path`.
// Interpolates *within* each segment (not vertex-by-vertex) and eases the
// progress so the line grows continuously rather than jumping.
function animateLineDraw(
  g: any,
  map: any,
  rawPath: any[],
  durationMs = 4500,
  store?: { current: any[] },
  // Optional: fire onReach(k) exactly when the drawn line reaches
  // locationCoords[k]. isActive() lets a caller stop the draw (pause).
  locationCoords?: Array<{ lat: number; lng: number }>,
  onReach?: (k: number) => void,
  isActive?: () => boolean
) {
  // Normalize to plain {lat,lng} (overview_path gives LatLng objects).
  const path = rawPath.map((p: any) =>
    typeof p.lat === "function" ? { lat: p.lat(), lng: p.lng() } : { lat: p.lat, lng: p.lng }
  );
  if (path.length < 2) return;

  // Clear any previously drawn route lines so we can redraw (e.g. slowly).
  if (store?.current) {
    for (const l of store.current) {
      try { l.setMap(null); } catch { /* ignore */ }
    }
    store.current = [];
  }

  // Per-segment lengths (equirectangular approx) + cumulative total.
  const seg: number[] = [];
  let totalDist = 0;
  for (let i = 1; i < path.length; i++) {
    const dy = path[i].lat - path[i - 1].lat;
    const dx =
      (path[i].lng - path[i - 1].lng) *
      Math.cos(((path[i].lat + path[i - 1].lat) / 2) * (Math.PI / 180));
    const d = Math.sqrt(dx * dx + dy * dy);
    seg.push(d);
    totalDist += d;
  }
  if (totalDist <= 0) return;

  // Cumulative distance at each vertex, and the distance at which the line
  // "reaches" each location (nearest vertex), for synced box popups.
  const cum: number[] = [0];
  for (let i = 0; i < seg.length; i++) cum.push(cum[i] + seg[i]);
  let reachDist: number[] | null = null;
  let fired: boolean[] = [];
  if (locationCoords && onReach) {
    reachDist = locationCoords.map((lc) => {
      let best = 0;
      let bestD = Infinity;
      for (let i = 0; i < path.length; i++) {
        const dy = path[i].lat - lc.lat;
        const dx = (path[i].lng - lc.lng) * Math.cos(((path[i].lat + lc.lat) / 2) * (Math.PI / 180));
        const d = dx * dx + dy * dy;
        if (d < bestD) { bestD = d; best = i; }
      }
      return cum[best];
    });
    fired = reachDist.map(() => false);
  }

  const casing = new g.maps.Polyline({
    map,
    path: [path[0]],
    strokeColor: LINE_CASING,
    strokeWeight: 9,
    strokeOpacity: 1,
    zIndex: 1,
  });
  // Repeating direction arrows along the line (shows travel direction + turns).
  // Guarded so a missing SymbolPath can never break the animation.
  let icons: any = undefined;
  try {
    const arrowPath = g?.maps?.SymbolPath?.FORWARD_CLOSED_ARROW;
    if (arrowPath != null) {
      icons = [
        {
          icon: {
            path: arrowPath,
            scale: 3,
            strokeColor: LINE_CASING,
            strokeWeight: 1,
            fillColor: "#ffffff",
            fillOpacity: 1,
          },
          offset: "0",
          repeat: "90px",
        },
      ];
    }
  } catch {
    icons = undefined;
  }

  const lineOpts: any = {
    map,
    path: [path[0]],
    strokeColor: LINE_MAIN,
    strokeWeight: 5,
    strokeOpacity: 1,
    zIndex: 2,
  };
  if (icons) lineOpts.icons = icons;
  const line = new g.maps.Polyline(lineOpts);

  if (store) store.current = [casing, line];

  // Linear progress when synced to location reach (so distance == time and
  // boxes line up); eased otherwise for a pretty solo draw.
  const linear = !!reachDist;

  let startTs: number | null = null;
  const tick = (ts: number) => {
    if (isActive && !isActive()) return; // paused/stopped
    if (startTs == null) startTs = ts;
    const t = Math.min(1, (ts - startTs) / durationMs);
    const te = linear ? t : t < 0.5 ? 4 * t * t * t : 1 - Math.pow(-2 * t + 2, 3) / 2;
    const target = te * totalDist;

    const out: Array<{ lat: number; lng: number }> = [path[0]];
    let acc = 0;
    for (let i = 0; i < seg.length; i++) {
      if (acc + seg[i] < target) {
        out.push(path[i + 1]);
        acc += seg[i];
      } else {
        const frac = seg[i] > 0 ? (target - acc) / seg[i] : 0;
        out.push({
          lat: path[i].lat + (path[i + 1].lat - path[i].lat) * frac,
          lng: path[i].lng + (path[i + 1].lng - path[i].lng) * frac,
        });
        break;
      }
    }
    casing.setPath(out);
    line.setPath(out);

    // Fire location-reached callbacks as the line passes each location.
    if (reachDist && onReach) {
      for (let k = 0; k < reachDist.length; k++) {
        if (!fired[k] && target >= reachDist[k]) {
          fired[k] = true;
          try { onReach(k); } catch { /* ignore */ }
        }
      }
    }

    if (t < 1) requestAnimationFrame(tick);
    else {
      casing.setPath(path);
      line.setPath(path);
      if (reachDist && onReach) {
        for (let k = 0; k < reachDist.length; k++) {
          if (!fired[k]) { fired[k] = true; try { onReach(k); } catch { /* ignore */ } }
        }
      }
    }
  };
  requestAnimationFrame(tick);
}

type RouteBuild = {
  road: boolean; // did ANY section follow real roads?
  straight: boolean; // did ANY section fall back to a straight line?
  status: string; // last non-OK Directions status (for a user hint)
  via: "google" | "ors" | "straight"; // where the road geometry came from
};

// ORS free plan caps a directions request at 50 stops.
const ORS_MAX = 50;

// Build a road path from the free OpenRouteService proxy, in batches of up to
// ORS_MAX stops (overlapping by one so segments join). Any batch that fails
// falls back to a straight segment through its own points.
async function orsRoute(
  points: ReportPoint[],
  fetchRoad: (coords: Array<{ lat: number; lng: number }>) => Promise<Array<{ lat: number; lng: number }> | null>
): Promise<{ path: Array<{ lat: number; lng: number }>; road: boolean; straight: boolean }> {
  const path: Array<{ lat: number; lng: number }> = [];
  let road = false;
  let straight = false;
  for (let start = 0; start < points.length - 1; start += ORS_MAX - 1) {
    const batch = points.slice(start, start + ORS_MAX);
    if (batch.length < 2) break;
    // eslint-disable-next-line no-await-in-loop
    const seg = await fetchRoad(batch.map((p) => ({ lat: p.lat, lng: p.lng })));
    if (seg && seg.length >= 2) {
      for (const c of seg) path.push(c);
      road = true;
    } else {
      for (const p of batch) path.push({ lat: p.lat, lng: p.lng });
      straight = true;
    }
  }
  return { path, road, straight };
}

// Ask Directions for ONE batch, retrying transient failures and shrinking the
// batch if it has too many waypoints. Returns the road path or null.
async function routeBatch(
  g: any,
  service: any,
  batch: ReportPoint[],
  onStatus: (s: string) => void,
  depth = 0
): Promise<any[] | null> {
  if (batch.length < 2) return null;
  const origin = { lat: batch[0].lat, lng: batch[0].lng };
  const destination = { lat: batch[batch.length - 1].lat, lng: batch[batch.length - 1].lng };
  const waypoints = batch.slice(1, -1).map((p) => ({
    location: { lat: p.lat, lng: p.lng },
    stopover: true,
  }));

  const attempt = (): Promise<{ ok: boolean; res?: any; status?: string }> =>
    new Promise((resolve) => {
      let tries = 0;
      const go = () => {
        service.route(
          { origin, destination, waypoints, travelMode: g.maps.TravelMode.DRIVING },
          (res: any, status: string) => {
            if (status === "OK" && res) return resolve({ ok: true, res });
            // Transient — back off and retry a few times.
            if ((status === "OVER_QUERY_LIMIT" || status === "UNKNOWN_ERROR") && tries < 3) {
              tries += 1;
              setTimeout(go, 500 * tries);
              return;
            }
            resolve({ ok: false, status });
          }
        );
      };
      go();
    });

  const r = await attempt();
  if (r.ok) {
    const overview = r.res?.routes?.[0]?.overview_path || [];
    return overview.length >= 2 ? overview : null;
  }
  onStatus(r.status || "UNKNOWN");

  // Too many stops for this account/plan — split the batch and try each half.
  if (r.status === "MAX_WAYPOINTS_EXCEEDED" && batch.length > 2 && depth < 4) {
    const mid = Math.ceil(batch.length / 2);
    const left = await routeBatch(g, service, batch.slice(0, mid + 1), onStatus, depth + 1);
    const right = await routeBatch(g, service, batch.slice(mid), onStatus, depth + 1);
    if (left || right) {
      return [...(left || batch.slice(0, mid + 1).map((p) => ({ lat: p.lat, lng: p.lng }))), ...(right || [])];
    }
  }
  return null;
}

// Draw a road-following route (Directions) through the points in batches, then
// animate it drawing start->end. Each batch falls back to a STRAIGHT segment on
// its own if roads can't be fetched, so one bad section never wipes out
// road-following for the whole route.
async function drawRoute(
  g: any,
  map: any,
  points: ReportPoint[],
  store?: { current: any[] },
  onPath?: (path: any[]) => void,
  fetchRoad?: (
    coords: Array<{ lat: number; lng: number }>
  ) => Promise<Array<{ lat: number; lng: number }> | null>
): Promise<RouteBuild> {
  const build: RouteBuild = { road: false, straight: false, status: "", via: "straight" };
  if (points.length < 2) return build;

  let fullPath: any[] = [];
  const pushStraight = (batch: ReportPoint[]) => {
    for (const p of batch) fullPath.push({ lat: p.lat, lng: p.lng });
    build.straight = true;
  };

  // 1) Prefer Google Directions (only works when the key has billing on).
  try {
    await g.maps.importLibrary("routes");
    const service = new g.maps.DirectionsService();
    const BATCH = 25; // stops per request (origin + up to 23 waypoints + destination)
    const onStatus = (s: string) => {
      build.status = s;
    };

    for (let start = 0; start < points.length - 1; start += BATCH - 1) {
      const batch = points.slice(start, start + BATCH);
      if (batch.length < 2) break;
      // eslint-disable-next-line no-await-in-loop
      const road = await routeBatch(g, service, batch, onStatus);
      if (road && road.length >= 2) {
        for (const ll of road) fullPath.push(ll);
        build.road = true;
        build.via = "google";
      } else {
        pushStraight(batch); // this section only
      }
    }
  } catch (e) {
    console.warn("[routemap] directions unavailable:", e);
    if (!build.status) build.status = "LIBRARY";
    fullPath.length = 0;
  }

  // 2) Google gave no roads (billing off / denied) -> free OpenRouteService.
  if (!build.road && fetchRoad) {
    try {
      const ors = await orsRoute(points, fetchRoad);
      if (ors.road) {
        fullPath = ors.path;
        build.road = true;
        build.straight = ors.straight;
        build.via = "ors";
      }
    } catch (e) {
      console.warn("[routemap] ORS routing failed:", e);
    }
  }

  // 3) Nothing usable at all -> straight through every point.
  if (fullPath.length < 2) {
    fullPath.length = 0;
    for (const p of points) fullPath.push({ lat: p.lat, lng: p.lng });
    build.road = false;
    build.straight = true;
    build.via = "straight";
  }

  // Ensure SymbolPath (for the arrows) is loaded; ignore if unavailable.
  try {
    await g.maps.importLibrary("core");
  } catch {
    // ignore
  }

  onPath?.(fullPath);

  try {
    animateLineDraw(g, map, fullPath, 4000, store);
  } catch (e) {
    console.warn("[routemap] line animation failed:", e);
  }

  return build;
}

function escapeHtml(s: string): string {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ---- Animated route-mapping flowchart (the "Locations" view) ----
function flowHsl(h: number, s: number, l: number): string {
  h = ((h % 360) + 360) % 360; s /= 100; l /= 100;
  const c = (1 - Math.abs(2 * l - 1)) * s, x = c * (1 - Math.abs(((h / 60) % 2) - 1)), m = l - c / 2;
  let r = 0, g = 0, b = 0;
  if (h < 60) [r, g, b] = [c, x, 0]; else if (h < 120) [r, g, b] = [x, c, 0];
  else if (h < 180) [r, g, b] = [0, c, x]; else if (h < 240) [r, g, b] = [0, x, c];
  else if (h < 300) [r, g, b] = [x, 0, c]; else [r, g, b] = [c, 0, x];
  const t = (v: number) => Math.round((v + m) * 255).toString(16).padStart(2, "0");
  return `#${t(r)}${t(g)}${t(b)}`;
}

// Per-city colour + number, matching the flowchart EXACTLY. "Cities" are
// consecutive-unique locations; the hue is spread 210..510 across the cities
// and rendered with flowHsl(58,60) fill / flowHsl(58,40) border. So a point's
// "Play on map" box gets the SAME one colour its city has in the flowchart.
function cityInfoForPoint(
  pts: Array<{ location?: string | null }>,
  i: number
): { fill: string; border: string; num: number } {
  let total = 0;
  let last = "";
  for (const q of pts) {
    if (!q?.location) continue;
    const k = q.location.toLowerCase();
    if (k !== last) {
      total++;
      last = k;
    }
  }
  let cityIdx = -1;
  last = "";
  for (let j = 0; j <= i && j < pts.length; j++) {
    const q = pts[j];
    if (!q?.location) continue;
    const k = q.location.toLowerCase();
    if (k !== last) {
      cityIdx++;
      last = k;
    }
  }
  const hue = 210 + (total > 1 && cityIdx >= 0 ? (cityIdx / (total - 1)) * 300 : 0);
  return {
    fill: flowHsl(hue, 58, 60),
    border: flowHsl(hue, 58, 40),
    num: (cityIdx >= 0 ? cityIdx : i) + 1,
  };
}

type FlowCell =
  | { kind: "start"; label: string }
  | { kind: "point"; n: number; label: string; state: string; hue: number };

function LocationsFlow({
  points,
  onClose,
  onComplete,
}: {
  points: ReportPoint[];
  onClose: () => void;
  onComplete: () => void;
}) {
  const doneRef = useRef(false);
  const scrollRef = useRef<HTMLDivElement>(null);
  const svgRef = useRef<SVGSVGElement>(null);
  // Distinct main locations (collapse consecutive same place names).
  const locs = useMemo(() => {
    const out: Array<{ label: string; state: string }> = [];
    let last = "";
    points.forEach((p) => {
      if (!p.location) return;
      const k = p.location.toLowerCase();
      if (k !== last) {
        out.push({ label: p.location, state: p.state });
        last = k;
      }
    });
    if (!out.length) points.forEach((p) => out.push({ label: p.category || "Point", state: p.state }));
    return out;
  }, [points]);

  const total = locs.length;
  const cells: FlowCell[] = [
    { kind: "start", label: "START" },
    ...locs.map((l, i) => ({
      kind: "point" as const,
      n: i + 1,
      label: l.label,
      state: l.state,
      hue: 210 + (total > 1 ? (i / (total - 1)) * 300 : 0),
    })),
  ];

  const PAD = 24, TITLE_H = 40, BOX_W = 168, BOX_H = 66, H_GAP = 40, HEADING_H = 26, ROW_V_GAP = 46;
  const perRow = Math.min(6, Math.max(3, cells.length));
  const rows: FlowCell[][] = [];
  for (let i = 0; i < cells.length; i += perRow) rows.push(cells.slice(i, i + perRow));

  const W = PAD * 2 + perRow * BOX_W + (perRow - 1) * H_GAP;
  const contentTop = PAD + TITLE_H;
  const blk = HEADING_H + BOX_H + ROW_V_GAP;
  const H = contentTop + rows.length * blk - ROW_V_GAP + PAD;
  const cellX = (c: number) => PAD + c * (BOX_W + H_GAP);
  const rowTop = (r: number) => contentTop + r * blk + HEADING_H;
  const flatIndex = (r: number, ci: number) => r * perRow + ci;

  // Reveal one cell at a time.
  const [revealed, setRevealed] = useState(1);
  useEffect(() => {
    if (revealed >= cells.length) return;
    const t = setTimeout(() => setRevealed((v) => v + 1), 1000);
    return () => clearTimeout(t);
  }, [revealed, cells.length]);

  // Crawl: auto-scroll the flowchart so the card being revealed stays in view
  // (keeps the "moving" chart following itself down the page for long routes).
  useEffect(() => {
    const cont = scrollRef.current;
    const svg = svgRef.current;
    if (!cont || !svg) return;
    const idx = Math.min(revealed - 1, cells.length - 1);
    const currentRow = Math.floor(idx / perRow);
    const scale = (svg.clientWidth || W) / W; // rendered px per viewBox unit
    const rowCenter = (rowTop(currentRow) + BOX_H / 2) * scale;
    const target = rowCenter - cont.clientHeight / 2;
    cont.scrollTo({ top: Math.max(0, target), behavior: "smooth" });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [revealed]);

  // When the whole flowchart has been revealed, hand off to the map tour.
  useEffect(() => {
    if (revealed >= cells.length && !doneRef.current) {
      doneRef.current = true;
      const t = setTimeout(() => onComplete(), 1400);
      return () => clearTimeout(t);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [revealed, cells.length]);

  const NAVY = "#16306b", LINE = "#1f3a93";

  return (
    <div style={styles.flowOverlay}>
      <div style={styles.flowBar}>
        <button style={styles.backBtn} onClick={onClose}>← Reports</button>
        <div style={{ fontWeight: 900 }}>Locations — Route Mapping</div>
        <button
          style={styles.backBtn}
          onClick={() => {
            doneRef.current = false;
            setRevealed(1);
          }}
          title="Replay"
        >
          ⟲ Replay
        </button>
      </div>
      <div style={styles.flowScroll} ref={scrollRef}>
        <svg ref={svgRef} viewBox={`0 0 ${W} ${H}`} width="100%" style={{ maxWidth: W, display: "block", margin: "0 auto" }}>
          <defs>
            <marker id="fa" markerWidth="9" markerHeight="9" refX="6" refY="3" orient="auto" markerUnits="userSpaceOnUse">
              <path d="M0,0 L6,3 L0,6 Z" fill={LINE} />
            </marker>
          </defs>
          <text x={W / 2} y={PAD + 26} textAnchor="middle" fontFamily="Segoe UI, Arial" fontSize={24} fontWeight={700} fill={NAVY}>
            ROUTE MAPPING
          </text>

          {rows.map((row, r) => {
            const bt = rowTop(r);
            const states: string[] = [];
            row.forEach((c) => { if (c.kind === "point" && c.state && !states.includes(c.state)) states.push(c.state); });
            const rowRevealed = row.some((_, ci) => flatIndex(r, ci) < revealed);
            return (
              <g key={r}>
                {states.length ? (
                  <text x={W / 2} y={bt - 8} textAnchor="middle" fontFamily="Segoe UI, Arial" fontSize={14} fontWeight={700}
                        fill={NAVY} opacity={rowRevealed ? 1 : 0.15} style={{ transition: "opacity .5s" }}>
                    {states.join(" / ").toUpperCase()}
                  </text>
                ) : null}

                {row.map((c, ci) => {
                  const x = cellX(ci);
                  const fi = flatIndex(r, ci);
                  const shown = fi < revealed;
                  const isStart = c.kind === "start";
                  const fill = isStart ? NAVY : flowHsl((c as any).hue, 58, 60);
                  const stroke = isStart ? "#0c1f4d" : flowHsl((c as any).hue, 58, 40);
                  return (
                    <g key={ci} opacity={shown ? 1 : 0.12} style={{ transition: "opacity .5s" }}>
                      <rect x={x} y={bt} width={BOX_W} height={BOX_H} rx={12} fill={fill} stroke={stroke} strokeWidth={2} />
                      {!isStart && (
                        <>
                          <circle cx={x + 18} cy={bt + 18} r={11} fill="#fff" />
                          <text x={x + 18} y={bt + 22} textAnchor="middle" fontFamily="Arial" fontSize={11} fontWeight={700} fill={stroke}>
                            {(c as any).n}
                          </text>
                        </>
                      )}
                      <text x={x + BOX_W / 2} y={bt + BOX_H / 2 + (isStart ? 4 : 12)} textAnchor="middle"
                            fontFamily="Segoe UI, Arial" fontSize={13} fontWeight={700} fill="#fff">
                        {String(c.label).length > 18 ? String(c.label).slice(0, 17) + "…" : c.label}
                      </text>
                      {ci < row.length - 1 && (
                        <line x1={x + BOX_W + 4} y1={bt + BOX_H / 2} x2={cellX(ci + 1) - 4} y2={bt + BOX_H / 2}
                              stroke={LINE} strokeWidth={2.5} markerEnd="url(#fa)"
                              opacity={flatIndex(r, ci + 1) < revealed ? 1 : 0.12} style={{ transition: "opacity .5s" }} />
                      )}
                    </g>
                  );
                })}

                {r < rows.length - 1 && (() => {
                  const lastCi = row.length - 1;
                  const lx = cellX(lastCi) + BOX_W / 2;
                  const by = bt + BOX_H;
                  const nextTop = rowTop(r + 1);
                  const fx = cellX(0) + BOX_W / 2;
                  const midY = by + ROW_V_GAP / 2;
                  const on = flatIndex(r + 1, 0) < revealed;
                  return (
                    <path d={`M ${lx} ${by} L ${lx} ${midY} L ${fx} ${midY} L ${fx} ${nextTop - 4}`}
                          fill="none" stroke={LINE} strokeWidth={2.5} markerEnd="url(#fa)"
                          opacity={on ? 1 : 0.12} style={{ transition: "opacity .5s" }} />
                  );
                })()}
              </g>
            );
          })}
        </svg>
      </div>
    </div>
  );
}

function TopBar({
  onBack,
  title,
  right,
}: {
  onBack?: () => void;
  title: string;
  right?: React.ReactNode;
}) {
  return (
    <div style={styles.topbar}>
      {onBack ? (
        <button style={styles.backBtn} onClick={onBack}>
          ← Back
        </button>
      ) : (
        <div style={{ width: 80 }} />
      )}
      <div style={styles.title}>{title}</div>
      <div style={{ minWidth: 80, display: "flex", justifyContent: "flex-end" }}>
        {right || null}
      </div>
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  page: { height: "100vh", display: "flex", flexDirection: "column", background: "#0b1220" },
  shareBtn: {
    padding: "8px 12px",
    borderRadius: 10,
    border: "1px solid #1d4ed8",
    background: "#2563eb",
    color: "#fff",
    fontWeight: 800,
    fontSize: 13,
    cursor: "pointer",
    whiteSpace: "nowrap",
  },
  routeHint: {
    position: "absolute",
    left: 12,
    bottom: 12,
    maxWidth: 420,
    display: "flex",
    alignItems: "flex-start",
    gap: 8,
    background: "rgba(15,23,42,0.92)",
    color: "#fde68a",
    border: "1px solid #a16207",
    borderRadius: 10,
    padding: "9px 12px",
    fontSize: 12.5,
    fontWeight: 600,
    zIndex: 40,
    boxShadow: "0 8px 24px rgba(0,0,0,0.35)",
  },
  routeHintClose: {
    background: "transparent",
    border: "none",
    color: "#fde68a",
    cursor: "pointer",
    fontWeight: 900,
    fontSize: 13,
    lineHeight: 1,
    padding: 2,
  },
  modalOverlay: {
    position: "fixed",
    inset: 0,
    background: "rgba(2,6,23,0.72)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 5000,
    padding: 16,
  },
  modalCard: {
    width: "100%",
    maxWidth: 440,
    background: "#0f172a",
    border: "1px solid #1e293b",
    borderRadius: 16,
    padding: "22px 22px 18px",
    color: "#e2e8f0",
    boxShadow: "0 24px 70px rgba(0,0,0,0.5)",
  },
  modalTitle: { fontSize: 18, fontWeight: 900, color: "#fff", marginBottom: 8 },
  modalSub: { fontSize: 13, lineHeight: 1.5, color: "#94a3b8", marginBottom: 16 },
  modalLabel: {
    display: "block",
    fontSize: 12,
    fontWeight: 800,
    color: "#cbd5e1",
    margin: "10px 0 6px",
    letterSpacing: 0.3,
  },
  modalInput: {
    width: "100%",
    boxSizing: "border-box",
    padding: "11px 13px",
    borderRadius: 10,
    border: "1px solid #334155",
    background: "#0b1220",
    color: "#fff",
    fontSize: 14,
    outline: "none",
  },
  modalErr: {
    background: "#3f1d1d",
    border: "1px solid #7f1d1d",
    color: "#fecaca",
    borderRadius: 9,
    padding: "8px 11px",
    fontSize: 12.5,
    marginTop: 12,
  },
  modalActions: { display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 },
  modalCancel: {
    padding: "10px 16px",
    borderRadius: 10,
    border: "1px solid #334155",
    background: "#1e293b",
    color: "#e2e8f0",
    fontWeight: 800,
    cursor: "pointer",
  },
  modalCreate: {
    padding: "10px 18px",
    borderRadius: 10,
    border: "none",
    background: "#2563eb",
    color: "#fff",
    fontWeight: 900,
    cursor: "pointer",
  },
  linkRow: { display: "flex", gap: 8, marginBottom: 12 },
  linkInput: {
    flex: 1,
    minWidth: 0,
    padding: "11px 13px",
    borderRadius: 10,
    border: "1px solid #334155",
    background: "#0b1220",
    color: "#93c5fd",
    fontSize: 13,
    outline: "none",
  },
  copyBtn: {
    padding: "10px 16px",
    borderRadius: 10,
    border: "none",
    background: "#0e7490",
    color: "#fff",
    fontWeight: 800,
    cursor: "pointer",
    whiteSpace: "nowrap",
  },
  pwReminder: {
    fontSize: 12.5,
    color: "#cbd5e1",
    background: "#0b1220",
    border: "1px dashed #334155",
    borderRadius: 9,
    padding: "9px 12px",
  },
  topbar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "10px 16px",
    background: "#0f172a",
    color: "#fff",
  },
  backBtn: {
    padding: "8px 12px",
    borderRadius: 10,
    border: "1px solid #334155",
    background: "#1e293b",
    color: "#fff",
    fontWeight: 800,
    cursor: "pointer",
  },
  title: { fontSize: 16, fontWeight: 900, letterSpacing: 0.3 },
  layout: { flex: 1, display: "flex", minHeight: 0 },
  mapWrap: { position: "relative", flex: 1, minWidth: 0, background: "#000" },
  side: {
    width: 300,
    background: "#0f172a",
    color: "#e2e8f0",
    display: "flex",
    flexDirection: "column",
    borderLeft: "1px solid #1e293b",
  },
  sideHead: { padding: "12px 14px", fontWeight: 900, borderBottom: "1px solid #1e293b" },
  list: { overflowY: "auto", padding: 10, display: "flex", flexDirection: "column", gap: 8 },
  listItem: {
    display: "flex",
    gap: 10,
    alignItems: "center",
    textAlign: "left",
    padding: "10px 12px",
    borderRadius: 10,
    border: "1px solid #EAECF0",
    cursor: "pointer",
  },
  badge: {
    flexShrink: 0,
    width: 26,
    height: 26,
    borderRadius: 999,
    background: "#d50000",
    color: "#fff",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    fontWeight: 900,
    fontSize: 12,
  },
  listText: { display: "flex", flexDirection: "column", color: "#0f172a", minWidth: 0 },
  listSub: { fontSize: 11, color: "#64748b" },
  controls: {
    position: "absolute",
    left: 16,
    bottom: 16,
    display: "flex",
    gap: 8,
    alignItems: "center",
    background: "rgba(15,23,42,0.85)",
    padding: "8px 10px",
    borderRadius: 12,
    zIndex: 6,
  },
  ctrlBtn: {
    padding: "8px 12px",
    borderRadius: 10,
    border: "none",
    background: "#1d4ed8",
    color: "#fff",
    fontWeight: 800,
    cursor: "pointer",
  },
  ctrlIconBtn: {
    padding: "8px 11px",
    borderRadius: 10,
    border: "none",
    background: "#334155",
    color: "#fff",
    fontWeight: 800,
    fontSize: 14,
    lineHeight: 1,
    cursor: "pointer",
    minWidth: 38,
  },
  ctrlInfo: { color: "#cbd5e1", fontWeight: 700, fontSize: 13, paddingLeft: 4 },
  locMenu: {
    position: "absolute",
    bottom: "calc(100% + 8px)",
    left: 0,
    background: "#0f172a",
    border: "1px solid #334155",
    borderRadius: 12,
    padding: 6,
    display: "flex",
    flexDirection: "column",
    gap: 4,
    minWidth: 170,
    boxShadow: "0 8px 24px rgba(0,0,0,0.4)",
    zIndex: 8,
  },
  locMenuItem: {
    textAlign: "left",
    padding: "10px 12px",
    borderRadius: 8,
    border: "none",
    background: "transparent",
    color: "#e2e8f0",
    fontWeight: 800,
    fontSize: 13,
    cursor: "pointer",
  },
  banner: {
    position: "absolute",
    top: 16,
    left: "50%",
    transform: "translateX(-50%)",
    background: "rgba(15,23,42,0.9)",
    color: "#fff",
    fontWeight: 900,
    fontSize: 20,
    letterSpacing: 0.5,
    padding: "10px 22px",
    borderRadius: 999,
    zIndex: 6,
    maxWidth: "80%",
    textAlign: "center",
    boxShadow: "0 6px 24px rgba(0,0,0,0.35)",
  },
  lightbox: {
    position: "fixed",
    inset: 0,
    background: "rgba(0,0,0,0.85)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 2147483647,
    cursor: "zoom-out",
  },
  lightboxCard: {
    position: "relative",
    background: "#fff",
    borderRadius: 16,
    overflow: "hidden",
    width: "min(1080px, 94vw)",
    maxHeight: "88vh",
    display: "flex",
    flexDirection: "row",
    boxShadow: "0 20px 60px rgba(0,0,0,0.5)",
    cursor: "default",
  },
  lightboxImgWrap: {
    flex: 1,
    minWidth: 0,
    background: "#0b1220",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    position: "relative",
  },
  lbNav: {
    position: "absolute",
    top: "50%",
    transform: "translateY(-50%)",
    width: 44,
    height: 44,
    borderRadius: "50%",
    border: "none",
    background: "rgba(0,0,0,0.55)",
    color: "#fff",
    fontSize: 26,
    lineHeight: "44px",
    fontWeight: 700,
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 3,
  },
  lbNavPrev: { left: 12 },
  lbNavNext: { right: 12 },
  drawCanvas: {
    maxWidth: "100%",
    maxHeight: "85vh",
    background: "#0b1220",
    cursor: "crosshair",
    touchAction: "none",
    display: "block",
  },
  drawPanel: {
    marginTop: "auto",
    paddingTop: 14,
    borderTop: "1px solid #eef0f3",
    display: "flex",
    flexDirection: "column",
    gap: 10,
  },
  drawBtn: {
    padding: "10px 14px",
    borderRadius: 10,
    border: "1px solid #d7dbe0",
    background: "#f8fafc",
    fontWeight: 700,
    fontSize: 14,
    cursor: "pointer",
    color: "#0f172a",
  },
  drawTools: { display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" },
  toolBtn: {
    width: 38,
    height: 38,
    borderRadius: 8,
    border: "1px solid #d7dbe0",
    background: "#fff",
    fontSize: 18,
    cursor: "pointer",
    color: "#0f172a",
  },
  toolBtnActive: { background: "#0f172a", color: "#fff", borderColor: "#0f172a" },
  colorInput: { width: 38, height: 38, border: "1px solid #d7dbe0", borderRadius: 8, cursor: "pointer", padding: 2, background: "#fff" },
  swatch: {
    width: 30,
    height: 30,
    borderRadius: "50%",
    border: "1px solid rgba(0,0,0,0.15)",
    cursor: "pointer",
    outlineOffset: 2,
    flexShrink: 0,
  },
  widthSelect: { height: 38, borderRadius: 8, border: "1px solid #d7dbe0", padding: "0 8px", fontSize: 14, cursor: "pointer" },
  drawActions: { display: "flex", gap: 6, flexWrap: "wrap" },
  drawBtnGhost: {
    flex: 1,
    padding: "9px 10px",
    borderRadius: 8,
    border: "1px solid #d7dbe0",
    background: "#fff",
    fontWeight: 600,
    fontSize: 13,
    cursor: "pointer",
    color: "#0f172a",
  },
  drawBtnPrimary: {
    flex: 1,
    padding: "9px 10px",
    borderRadius: 8,
    border: "none",
    background: "#16a34a",
    color: "#fff",
    fontWeight: 700,
    fontSize: 13,
    cursor: "pointer",
  },
  lbCounter: {
    position: "absolute",
    bottom: 12,
    left: "50%",
    transform: "translateX(-50%)",
    background: "rgba(0,0,0,0.6)",
    color: "#fff",
    fontSize: 13,
    fontWeight: 600,
    padding: "4px 12px",
    borderRadius: 999,
    zIndex: 3,
  },
  lightboxImg: {
    width: "100%",
    height: "100%",
    maxHeight: "88vh",
    objectFit: "contain",
    display: "block",
  },
  lightboxInfo: {
    width: 360,
    flexShrink: 0,
    padding: "20px 22px",
    overflowY: "auto",
    borderLeft: "1px solid #eef0f3",
    display: "flex",
    flexDirection: "column",
    gap: 14,
  },
  lightboxHeader: { display: "flex", gap: 12, alignItems: "center" },
  lightboxBadge: {
    flexShrink: 0,
    width: 34,
    height: 34,
    borderRadius: 999,
    background: "#d50000",
    color: "#fff",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    fontWeight: 900,
    fontSize: 15,
  },
  lightboxTitle: { fontSize: 20, fontWeight: 900, color: "#0f172a", lineHeight: 1.2 },
  lightboxBody: { display: "flex", flexDirection: "column", gap: 14 },
  lightboxField: {
    background: "#f8fafc",
    border: "1px solid #eef0f3",
    borderRadius: 12,
    padding: "10px 12px",
  },
  lightboxLabel: {
    fontSize: 11,
    fontWeight: 900,
    letterSpacing: 0.6,
    textTransform: "uppercase",
    color: "#64748b",
    marginBottom: 4,
  },
  lightboxValue: { fontSize: 14, fontWeight: 600, color: "#0f172a", lineHeight: 1.5 },
  lightboxCoords: { fontSize: 12, fontWeight: 700, color: "#94a3b8" },
  flowOverlay: {
    position: "absolute",
    inset: 0,
    background: "#f8fafc",
    zIndex: 7,
    display: "flex",
    flexDirection: "column",
  },
  flowBar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "10px 14px",
    background: "#0f172a",
    color: "#fff",
    flexShrink: 0,
  },
  flowScroll: { flex: 1, overflow: "auto", padding: 16 },
  lightboxClose: {
    position: "absolute",
    top: 12,
    right: 12,
    width: 40,
    height: 40,
    borderRadius: 999,
    border: "none",
    background: "rgba(15,23,42,0.55)",
    color: "#fff",
    fontSize: 18,
    fontWeight: 900,
    cursor: "pointer",
    zIndex: 2,
  },
  overlay: {
    position: "absolute",
    inset: 0,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    color: "#fff",
    fontWeight: 800,
    background: "rgba(0,0,0,0.45)",
    zIndex: 5,
    padding: 24,
    textAlign: "center",
  },
  card: { margin: 24, padding: 20, background: "#fff", borderRadius: 14, maxWidth: 680 },
  h2: { fontSize: 18, fontWeight: 900, marginBottom: 8, color: "#0f172a" },
  p: { color: "#334155", lineHeight: 1.6, fontSize: 14 },
  pre: {
    background: "#0f172a",
    color: "#e2e8f0",
    padding: "10px 12px",
    borderRadius: 10,
    overflowX: "auto",
    fontSize: 13,
  },
};
