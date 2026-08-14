// Standard observation/description text per report category. When a report's
// category is chosen, its observation is auto-filled with the matching template
// (trailing spaces are intentional — the surveyor types the measured value
// right after). Shared by the projects-list bulk grid AND the project-detail
// Add/Edit report forms so the behaviour is identical everywhere.
export const DESCRIPTION_TEMPLATES: Record<string, string[]> = {
  "low tension cable": ["Low tension cable cross above the road height is "],
  "high tension cable": ["High tension cable cross above the road height is "],
  "tree branches": ["Tree branches crossing above the road height is "],
  signboard: ["Signboard crossing above the road Height is "],
  "side signboard": ["Side Signboard crossing above the road Height is "],
  "electric sign board": ["Signboard crossing above the road Height is "],
  "electric signboard": ["Signboard crossing above the road Height is "],
  "camera pole": ["Camera Pole crossing above the road Height is "],
  "signal pole": ["Signal Pole Crossing above the road height is "],
  "speed pole": ["Signal Pole Crossing above the road height is "],
  "towerline cable": ["Tower line cross above the road height is "],
  towerline: ["Tower line cross above the road height is "],
  "tower line": ["Tower line cross above the road height is "],
  "tower line cable": ["Tower line cross above the road height is "],
  "railway level crossing": ["Railway Barrier Height is m, Railway Crossing Cable Height is "],
  "take diversion": ["Take diversion Due to road work under Progress"],
  diversion: ["Take diversion Due to road work under Progress"],
  "toll plaza": ["Toll Plaza, Height is , Width is "],
  toll: ["Toll Plaza, Height is , Width is "],
  "underpass bridge": ["Underpass bridge crossing above the road height is "],
  underpass: ["Underpass bridge crossing above the road height is "],
  "footpath bridge": ["Footpath Bridge crossing above the road height is "],
  parking: ["Parking place is available on LHS", "Parking place is available on RHS"],
  "petrol bunk": ["Petrol Bunk available in LHS", "Petrol Bunk available in RHS"],
  petrol: ["Petrol Bunk available in LHS", "Petrol Bunk available in RHS"],
  tunnel: ["Tunnel crossing above the road height is "],
  bridge: [
    "Bridge,\nspans of m length each,\n•i – beams girder structure,\n•No of beams – 4,\n•Width is 0.5m,\n•Height is 1.5m,\n•Road width is m,\n•Side wall height is 1m on both sides",
    "Bridge,\nspan of m length,\n•RCC Box slab structure\n•Width is 0.6m,\nRoad width is m.\nSide wall height is 1m on both sides",
  ],
};

function normKey(category: string): string {
  return String(category || "")
    .toLowerCase()
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

// Every known template string (trimmed) — lets us tell "still a default
// template" apart from a description the surveyor actually typed.
export const ALL_DESCRIPTION_TEMPLATES = new Set(
  Object.values(DESCRIPTION_TEMPLATES)
    .flat()
    .map((s) => s.trim())
);

// The template list for a category (all templates if the category is unknown).
export function descriptionTemplatesFor(category: string): string[] {
  const exact = DESCRIPTION_TEMPLATES[normKey(category)];
  if (exact?.length) return exact;
  return Object.values(DESCRIPTION_TEMPLATES).flat();
}

// Auto-fill the description for a category — ONLY when the current description
// is empty or still a known default template. Never overwrites real edits.
export function descriptionForCategory(currentDescription: string, category: string): string {
  const cur = String(currentDescription || "").trim();
  const stillDefault = cur === "" || ALL_DESCRIPTION_TEMPLATES.has(cur);
  if (!stillDefault) return currentDescription;
  const exact = DESCRIPTION_TEMPLATES[normKey(category)];
  return exact && exact.length ? exact[0] : currentDescription;
}
