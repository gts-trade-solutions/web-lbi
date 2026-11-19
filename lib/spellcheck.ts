import fs from "fs";
import path from "path";

/**
 * Self-contained English spell checker (no external service, no npm deps).
 *
 * - Dictionary + word frequencies come from a bundled 50k word list
 *   (lib/data/en_50k.txt, "word count" per line).
 * - Suggestions use Peter Norvig's edit-distance-1/2 approach, ranked by
 *   word frequency.
 * - A seeded domain allow-list (survey / transport / logistics terms,
 *   British spellings) keeps real work vocabulary from being flagged.
 * - Callers may pass their own extra "known" words (the user's taught words).
 */

// ---- Dictionary load (once per server process) --------------------------

let FREQ: Map<string, number> | null = null;

function loadDict(): Map<string, number> {
  if (FREQ) return FREQ;
  const freq = new Map<string, number>();
  try {
    const file = path.join(process.cwd(), "lib", "data", "en_50k.txt");
    const text = fs.readFileSync(file, "utf8");
    for (const line of text.split("\n")) {
      const sp = line.indexOf(" ");
      if (sp <= 0) continue;
      const word = line.slice(0, sp).trim().toLowerCase();
      const count = Number(line.slice(sp + 1).trim()) || 1;
      // Keep purely-alphabetic words only (drop contractions / junk).
      if (!word || !/^[a-z]+$/.test(word)) continue;
      if (!freq.has(word)) freq.set(word, count);
    }
  } catch (e) {
    console.error("[spellcheck] dictionary load failed:", e);
  }
  // Seed domain + British-spelling vocabulary with a modest frequency so it
  // counts as "known" and can also surface as a suggestion.
  for (const w of EXTRA_WORDS) {
    const k = w.toLowerCase();
    if (/^[a-z]+$/.test(k) && !freq.has(k)) freq.set(k, 500);
  }
  FREQ = freq;
  return freq;
}

// Survey / road / transport / logistics vocabulary + common British
// spellings that a 50k subtitle list often lacks. Lowercase, letters only.
const EXTRA_WORDS: string[] = [
  // road / survey engineering
  "chainage", "culvert", "culverts", "gantry", "gantries", "cantilever",
  "underpass", "overpass", "flyover", "flyovers", "median", "carriageway",
  "carriageways", "kerb", "kerbs", "kerbside", "verge", "verges", "embankment",
  "embankments", "overhang", "overhangs", "clearance", "clearances",
  "obstruction", "obstructions", "bollard", "bollards", "guardrail",
  "guardrails", "signage", "hoarding", "hoardings", "pylon", "pylons",
  "substation", "feeder", "conductor", "conductors", "encroachment",
  "encroachments", "demolish", "demolition", "dismantle", "dismantling",
  "realign", "realignment", "widen", "widening", "divert", "diversion",
  "diversions", "axle", "axles", "trailer", "trailers", "lowbed", "lowboy",
  "oversize", "oversized", "consignment", "consignments", "cargo", "convoy",
  "escort", "permit", "permits", "toll", "tollgate", "tollbooth",
  "roundabout", "roundabouts", "junction", "junctions", "intersection",
  "intersections", "bypass", "footpath", "drainage", "utility", "utilities",
  "overhead", "catenary", "railway", "crossing", "crossings", "parapet",
  "parapets", "retaining", "gradient", "gradients", "incline", "decline",
  "hairpin", "metalled", "bituminous", "tarmac", "laterite", "abutment",
  "abutments", "pier", "piers", "span", "spans", "girder", "girders", "deck",
  "approach", "approaches", "shoulder", "shoulders", "formation", "alignment",
  "turning", "radius", "swept", "swing", "tracking", "obstacle", "obstacles",
  "navigable", "traversable", "manoeuvre", "manoeuvrability", "maneuver",
  "maneuverability", "overswing", "offtake", "pavement", "potholes", "pothole",
  "widths", "heights", "terminal", "terminals", "gateway", "gateways",
  // India context place words
  "nagar", "chowk", "naka", "phata", "ghat", "maidan", "colony", "sector",
  "taluka", "tehsil", "mandal", "highway", "expressway", "marg",
  // British spellings
  "authorise", "authorised", "authorises", "authorising", "authorisation",
  "organise", "organised", "organises", "organisation", "colour", "colours",
  "coloured", "metre", "metres", "kilometre", "kilometres", "centre",
  "centres", "centred", "labour", "neighbour", "programme", "licence",
  "defence", "offence", "travelling", "travelled", "cancelled", "modelling",
  "levelling", "signalling", "recognise", "recognised", "utilise", "utilised",
  "analyse", "analysed", "favour", "favourable", "behaviour",
];

const ALPHABET = "abcdefghijklmnopqrstuvwxyz";

function edits1(word: string): Set<string> {
  const splits: [string, string][] = [];
  for (let i = 0; i <= word.length; i++) splits.push([word.slice(0, i), word.slice(i)]);
  const out = new Set<string>();
  for (const [L, R] of splits) {
    if (R) out.add(L + R.slice(1)); // delete
    if (R.length > 1) out.add(L + R[1] + R[0] + R.slice(2)); // transpose
    for (const c of ALPHABET) {
      if (R) out.add(L + c + R.slice(1)); // replace
      out.add(L + c + R); // insert
    }
  }
  return out;
}

function known(words: Iterable<string>, dict: Map<string, number>): string[] {
  const out: string[] = [];
  for (const w of words) if (dict.has(w)) out.push(w);
  return out;
}

/** Ranked suggestions (best first) for a lowercase word. */
export function suggest(word: string, extra?: Set<string>): string[] {
  const dict = loadDict();
  const w = word.toLowerCase();
  const isKnown = (x: string) => dict.has(x) || (extra ? extra.has(x) : false);
  if (isKnown(w)) return [];

  const e1 = edits1(w);
  let cands = new Set<string>(known(e1, dict));
  if (extra) for (const c of e1) if (extra.has(c)) cands.add(c);

  if (cands.size === 0) {
    // edit-distance 2
    for (const a of e1) {
      for (const b of edits1(a)) {
        if (dict.has(b) || (extra && extra.has(b))) cands.add(b);
      }
    }
  }
  return Array.from(cands)
    .sort((a, b) => (dict.get(b) || 0) - (dict.get(a) || 0))
    .slice(0, 6);
}

// ---- Full-text check -----------------------------------------------------

export type Misspelling = { word: string; suggestions: string[]; count: number };

// Match a word plus any trailing apostrophe-suffix (port's, don't).
const WORD_RE = /[A-Za-z][A-Za-z'’]*/g;
const SENTENCE_END = new Set([".", "!", "?", ":", ";", "\n", "\r", "•", "·", "-", "—", "–"]);

function isAllUpper(s: string): boolean {
  return /[A-Z]/.test(s) && s === s.toUpperCase();
}

/**
 * Find likely-misspelled words in `text`.
 * Skips: acronyms (ALL CAPS), numbers, proper nouns (a capitalised word that
 * is NOT at the start of a sentence), words < 3 letters, and anything in the
 * dictionary, the domain allow-list, or the caller's `ignore` set.
 */
export function checkText(text: string, ignore?: string[]): Misspelling[] {
  const dict = loadDict();
  const extra = new Set<string>((ignore || []).map((w) => w.toLowerCase().trim()).filter(Boolean));
  const raw = String(text || "");
  if (!raw.trim()) return [];

  const found = new Map<string, Misspelling>();
  let m: RegExpExecArray | null;
  WORD_RE.lastIndex = 0;
  while ((m = WORD_RE.exec(raw)) !== null) {
    const token = m[0];
    const start = m.index;
    const core = token.replace(/['’]/g, "").toLowerCase();
    if (core.length < 3) continue;
    if (!/^[a-z]+$/.test(core)) continue;
    if (isAllUpper(token)) continue; // acronym like JNPT, LHS, RECOMMEND

    // Proper-noun heuristic: capitalised word not at sentence start.
    const firstUpper = /[A-Z]/.test(token[0]);
    if (firstUpper) {
      let j = start - 1;
      while (j >= 0 && (raw[j] === " " || raw[j] === "\t" || raw[j] === '"' || raw[j] === "'" || raw[j] === "(")) j--;
      const sentenceStart = j < 0 || SENTENCE_END.has(raw[j]);
      if (!sentenceStart) continue;
    }

    if (dict.has(core) || extra.has(core)) continue;
    // Accept simple plural/possessive of a known word.
    if (core.endsWith("s") && (dict.has(core.slice(0, -1)) || extra.has(core.slice(0, -1)))) continue;

    const existing = found.get(token);
    if (existing) {
      existing.count += 1;
    } else {
      found.set(token, { word: token, suggestions: suggest(core, extra), count: 1 });
    }
  }
  return Array.from(found.values());
}
