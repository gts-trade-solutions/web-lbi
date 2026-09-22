// Shared category display-name grouping. Used by BOTH the Word export
// (lib/reenaTemplateExport.ts) and the Stage Summary editor API so the labels
// stay identical in the app and the report.

export function normalizeCategory(value: unknown): string {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

/** Collapse a raw report category into the label used in the Stage Summary. */
export function getCategoryDisplayName(category: unknown): string {
  const c = normalizeCategory(category);
  if (!c) return String(category || "-");
  if (c.includes("footpath bridge")) return "Footpath Bridge";
  if (c.includes("low tension")) return "Low Tension Cable";
  if (c.includes("high tension")) return "High Tension Cable";
  if (c.includes("tower")) return "Tower Line Cable";
  if (c.includes("underpass")) return "Underpass Bridge";
  if (c.includes("tree")) return "Tree Branches";
  if (c.includes("river bridge")) return "River Bridge";
  if (
    c.includes("signal pole") ||
    c.includes("speed pole") ||
    c.includes("side signboard")
  ) {
    return "Side Signboard / Signal Pole / Speed Pole";
  }
  if (
    c.includes("signboard") ||
    c.includes("camera pole") ||
    c.includes("electric sign")
  ) {
    return "Signboard / Electric Signboard / Camera Pole";
  }
  if (c.includes("toll")) return "Toll Plaza";
  if (c.includes("damage") || c.includes("pothole")) return "Damaged Road";
  if (c.includes("narrow")) return "Narrow Road";
  if (c.includes("gate")) return "Gate";
  if (c.includes("bend")) return "Bend";
  if (c.includes("petrol")) return "Petrol Bunk";
  if (c.includes("railway")) return "Railway Level Crossing";
  return String(category || "-");
}

export type StageSummaryRow = { label: string; count: number };

/**
 * Auto-compute the Stage Summary rows from a list of raw report categories,
 * grouped by display name and sorted by count (desc). This is the same shape
 * an editor override stores, so the two are interchangeable.
 */
export function summarizeCategories(categories: unknown[]): StageSummaryRow[] {
  const counts = new Map<string, number>();
  for (const cat of categories) {
    const label = getCategoryDisplayName(cat);
    counts.set(label, (counts.get(label) || 0) + 1);
  }
  return Array.from(counts.entries())
    .map(([label, count]) => ({ label, count }))
    .sort((a, b) => b.count - a.count);
}
