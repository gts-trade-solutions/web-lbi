/* eslint-disable @typescript-eslint/no-explicit-any */
// OpenRouteService (openrouteservice.org) road-following directions — a free,
// OpenStreetMap-based alternative to Google Directions. Used server-side only
// so the API key never reaches the browser.

export const ORS_DIRECTIONS_URL =
  "https://api.openrouteservice.org/v2/directions/driving-car/geojson";

// ORS free plan caps a single directions request at 50 locations.
export const ORS_MAX_COORDS = 50;

export type LatLng = { lat: number; lng: number };

// Our points are [lat, lng]; ORS wants [lng, lat]. Drop anything non-finite.
export function toOrsCoordinates(latlngs: Array<[number, number]>): Array<[number, number]> {
  const out: Array<[number, number]> = [];
  for (const c of latlngs) {
    const lat = Number(c?.[0]);
    const lng = Number(c?.[1]);
    if (Number.isFinite(lat) && Number.isFinite(lng)) out.push([lng, lat]);
  }
  return out;
}

// Pull the road geometry out of an ORS GeoJSON response as [{lat,lng}, ...].
export function orsGeometryToPath(orsJson: any): LatLng[] {
  const coords = orsJson?.features?.[0]?.geometry?.coordinates;
  if (!Array.isArray(coords)) return [];
  const path: LatLng[] = [];
  for (const c of coords) {
    const lng = Number(c?.[0]);
    const lat = Number(c?.[1]);
    if (Number.isFinite(lat) && Number.isFinite(lng)) path.push({ lat, lng });
  }
  return path;
}

export type OrsResult =
  | { ok: true; path: LatLng[] }
  | { ok: false; status: number; detail?: string };

// Fetch the road path for a set of [lat,lng] stops from ORS. Returns a
// discriminated result so the caller can react to config/quota errors.
export async function fetchOrsPath(
  latlngs: Array<[number, number]>,
  apiKey: string
): Promise<OrsResult> {
  const coordinates = toOrsCoordinates(latlngs);
  if (coordinates.length < 2) return { ok: false, status: 400, detail: "need >= 2 coordinates" };

  let res: Response;
  try {
    res = await fetch(ORS_DIRECTIONS_URL, {
      method: "POST",
      headers: { Authorization: apiKey, "Content-Type": "application/json" },
      // "recommended" favours the main road / highway (like Google's default);
      // it never picks the smaller-road "shortest" path.
      body: JSON.stringify({ coordinates, preference: "recommended" }),
    });
  } catch (e: any) {
    return { ok: false, status: 502, detail: e?.message || "request failed" };
  }

  const data = await res.json().catch(() => null);
  if (!res.ok) {
    const detail =
      data?.error?.message || (typeof data?.error === "string" ? data.error : "") || `HTTP ${res.status}`;
    return { ok: false, status: res.status, detail };
  }
  return { ok: true, path: orsGeometryToPath(data) };
}
