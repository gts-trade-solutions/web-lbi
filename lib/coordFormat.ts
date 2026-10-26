// One shared coordinate format used EVERYWHERE (project list, Edit form, route
// map and the Word report) so a point's NE coordinates always read the same.
// Decimal degrees with a direction letter, e.g. "N 14.2633142" / "E 80.0704166".
// The exact number the user typed is preserved (no DMS conversion), trimmed to
// at most 7 decimals (~1 cm) with trailing zeros removed.

function fmt(v: unknown, posLetter: "N" | "E", negLetter: "S" | "W", nbsp: boolean): string {
  const n = Number(v);
  if (!Number.isFinite(n)) return "-";
  const letter = n >= 0 ? posLetter : negLetter;
  let s = Math.abs(n).toFixed(7);
  if (s.includes(".")) s = s.replace(/0+$/, "").replace(/\.$/, "");
  // Non-breaking space keeps "N 14.26" from splitting across a line in Word.
  return `${letter}${nbsp ? " " : " "}${s}`;
}

/** "N 14.2633142" (or "S …" for a negative latitude). `nbsp` uses U+00A0. */
export function formatNeLat(v: unknown, nbsp = false): string {
  return fmt(v, "N", "S", nbsp);
}

/** "E 80.0704166" (or "W …" for a negative longitude). */
export function formatNeLon(v: unknown, nbsp = false): string {
  return fmt(v, "E", "W", nbsp);
}

/** "N 14.2633142, E 80.0704166" — both together. */
export function formatNePair(lat: unknown, lon: unknown, nbsp = false): string {
  return `${formatNeLat(lat, nbsp)}, ${formatNeLon(lon, nbsp)}`;
}
