// Builds the <trk> of a project GPX: which points join up into lines.
//
// A GPX viewer draws a straight line between every consecutive point inside
// one <trkseg>. The old export put every report's recorded path into a single
// segment, so the end of one report's trace was joined to the start of the
// next — dozens of long straight "spokes" across the map, radiating from
// wherever the traces began.
//
// The rule here: points are only joined when they really are one continuous
// movement.
//   - each report's own recorded trace is its own line, never stitched to
//     another report's trace;
//   - reports with no trace (a single observation point, e.g. bulk-imported)
//     are chained together in report order, which is how they form a route;
//   - any step longer than a real movement could be starts a new line.
//
// Pure — no database — so it can be checked on its own.

export interface TrackPoint {
  lat: number;
  lon: number;
  /** ISO 8601, already validated. */
  time?: string | null;
}

export interface GpxReport {
  id: string;
  /** Observation point, when valid. */
  lat: number | null;
  lon: number | null;
  /** Recorded trace in seq order, already filtered to valid coordinates. */
  path: TrackPoint[];
}

/**
 * Longest believable step between two consecutive fixes of one GPS
 * recording. Fixes are seconds apart, so anything beyond a kilometre is a
 * glitch or a stale first fix, not movement.
 */
export const TRACE_GAP_KM = 1;

/**
 * Longest believable step between two consecutive observation points. They
 * can legitimately be a few kilometres apart on a highway survey; much more
 * than this and the line would cut across country rather than follow a road.
 */
export const ROUTE_GAP_KM = 5;

const EARTH_KM = 6371;

/** Great-circle distance in kilometres. */
export function distanceKm(a: TrackPoint, b: TrackPoint): number {
  const rad = Math.PI / 180;
  const dLat = (b.lat - a.lat) * rad;
  const dLon = (b.lon - a.lon) * rad;
  const h =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(a.lat * rad) * Math.cos(b.lat * rad) * Math.sin(dLon / 2) ** 2;
  return 2 * EARTH_KM * Math.asin(Math.min(1, Math.sqrt(h)));
}

export function validLatLon(lat: number, lon: number): boolean {
  return (
    Number.isFinite(lat) &&
    Number.isFinite(lon) &&
    Math.abs(lat) <= 90 &&
    Math.abs(lon) <= 180 &&
    // 0,0 is what a GPS reports before it has a fix.
    !(lat === 0 && lon === 0)
  );
}

/** Split a run of points wherever one step is longer than `maxKm`. */
function splitAtJumps(points: TrackPoint[], maxKm: number): TrackPoint[][] {
  const out: TrackPoint[][] = [];
  let seg: TrackPoint[] = [];
  for (const p of points) {
    const prev = seg[seg.length - 1];
    if (prev && distanceKm(prev, p) > maxKm) {
      out.push(seg);
      seg = [];
    }
    seg.push(p);
  }
  if (seg.length) out.push(seg);
  return out;
}

/**
 * The line segments of the track, in report order. A segment of one point
 * draws nothing, so those are dropped — the report still has its waypoint.
 */
export function buildSegments(reports: GpxReport[]): TrackPoint[][] {
  const segments: TrackPoint[][] = [];
  let chain: TrackPoint[] = [];

  const keep = (seg: TrackPoint[]) => {
    if (seg.length >= 2) segments.push(seg);
  };
  const endChain = () => {
    for (const seg of splitAtJumps(chain, ROUTE_GAP_KM)) keep(seg);
    chain = [];
  };

  for (const r of reports) {
    if (r.path.length >= 2) {
      // A real recording: its own line (or lines, where the GPS jumped).
      endChain();
      for (const seg of splitAtJumps(r.path, TRACE_GAP_KM)) keep(seg);
      continue;
    }

    // One point only — the single path point, or else the observation.
    const point =
      r.path[0] ??
      (r.lat !== null && r.lon !== null && validLatLon(r.lat, r.lon)
        ? { lat: r.lat, lon: r.lon }
        : null);
    if (point) chain.push(point);
  }
  endChain();

  return segments;
}

/** An ISO time for GPX, or null for anything unparseable. */
export function gpxTime(value: unknown): string | null {
  if (value === null || value === undefined || value === "") return null;
  const d = value instanceof Date ? value : new Date(String(value));
  return Number.isNaN(d.getTime()) ? null : d.toISOString();
}
