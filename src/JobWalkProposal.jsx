import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import * as math from "mathjs";

// ============================================================================
// CONFIG / CONSTANTS
// ============================================================================

const PRICING_TIERS = {
  standard:   { label: "Standard",        multiplier: 1.00, color: "#4A90D9", prepTier: "standard" },
  production: { label: "Production",      multiplier: 0.85, color: "#7B68EE", prepTier: "standard" },
  highend:    { label: "High-End",        multiplier: 1.35, color: "#C8963E", prepTier: "high_end" },
  prevailing: { label: "Prevailing Wage", multiplier: 1.65, color: "#E05C5C", prepTier: "high_end" },
};

const CATALOG_IDS = {
  walls:     { "1coat": "22PWT9nDL9HP", "2coats": "22PWiSS293E2", "prime+2": "22PWiSgjqkaq", code: "22PWmcdKrCnn" },
  ceilings:  { "1coat": "22PWickkTi46", "2coats": "22PWictZ3V84", "prime+2": "22PWicwN2pxL", code: "22PWmcdMcqbZ" },
  baseboard: { "1coat": "22PWiYQt5Pq8", "2coats": "22PWih8XPvPm", code: "22PWmcdjtJ9C" },
  doors:     { "1coat": "22PXFmQz3HEd", "2coats": "22PXFmRbHAgn", "prime+2": "22PXFmSG2zeV", code: "22PWmcdpBtSV" },
};

// T&M Catalog — canonical hourly line items, mapped to cost groups
const TM_CATALOG = [
  { id: "22PWTAs6vVPw", name: "Time & Materials",                          code: "1000", codeName: "Interior Walls & Ceilings",       unitCost: 43.33, unitPrice: 65, fitsCostGroups: ["drywall_walls_ceilings"] },
  { id: "22PWmfVWsyZ9", name: "Time & Materials – Doors",                  code: "2000", codeName: "Doors & Windows",                 unitCost: 43.34, unitPrice: 65, fitsCostGroups: ["doors_frames"] },
  { id: "22PWmfVZLfua", name: "Time & Materials – Windows",                code: "2000", codeName: "Doors & Windows",                 unitCost: 43.34, unitPrice: 65, fitsCostGroups: ["doors_frames"] },
  { id: "22PWmfVaqmfL", name: "Time & Materials – Trim & Beams",           code: "4000", codeName: "Trim & Beams",                    unitCost: 43.34, unitPrice: 65, fitsCostGroups: ["baseboard_trim"] },
  { id: "22PWmfVcQeW9", name: "Time & Materials – Exterior Walls & Ceilings", code: "5000", codeName: "Exterior Walls & Ceilings",    unitCost: 43.34, unitPrice: 65, fitsCostGroups: ["exterior_stucco_siding"] },
  { id: "22PWmfVec8ks", name: "Time & Materials – Exterior Miscellaneous", code: "6000", codeName: "Exterior Miscellaneous",          unitCost: 43.34, unitPrice: 65, fitsCostGroups: ["exterior_wood_trim"] },
  { id: "22PWmfVgD8nW", name: "Time & Materials – Millwork & Specialty",   code: "7000", codeName: "Millwork & Specialty Coatings",   unitCost: 43.34, unitPrice: 65, fitsCostGroups: ["cabinets"] },
  { id: "22PWmfVhkPTk", name: "Time & Materials – Repairs & Cleaning",     code: "8000", codeName: "Cleaning & Repairs",              unitCost: 43.34, unitPrice: 65, fitsCostGroups: ["universal"] },
  { id: "22PWmfVpaFLF", name: "Time & Materials – Cleaning Services",      code: "8200", codeName: "Cleaning",                        unitCost: 36.67, unitPrice: 55, fitsCostGroups: ["universal"] },
  { id: "22PWmfVmyWGZ", name: "Time & Materials – Materials & Supplies",   code: "9000", codeName: "Non-Labor Costs",                 unitCost: 43.34, unitPrice: 65, fitsCostGroups: ["universal"] },
];

// ============================================================================
// DEFAULT SCOPE LIBRARY (v3 — refined, cost-group-linked, tier-aware)
// ============================================================================

const DEFAULT_LIBRARY = {
  version: "v3",
  blanket_exclusions: [
    "Any items not listed on this estimate",
    "Prefinished or factory-finished items (unless specified)",
    "Lead, asbestos, or mold testing or abatement",
    "Permits and HOA approvals",
    "Plumbing, electrical, or HVAC work",
    "Structural repairs and dry rot beyond minor patching",
    "Color changes after work has started",
    "Materials provided by others unless specified",
    "Window glass, screens, and hardware cleaning",
    "Furniture moving (owner to clear work areas)",
  ],
  universal_prep: {
    standard: [
      "Mask floors with paper and adjacent surfaces with plastic",
      "Cover and protect furniture, fixtures, and finished surfaces",
      "Daily site cleanup and debris removal",
      "Final walkthrough with owner upon completion",
    ],
    high_end: [
      "Mask floors with rosin paper and adjacent surfaces with plastic and tape",
      "Cover and protect furniture, fixtures, flooring, and all finished surfaces with care",
      "Daily site cleanup; tools and materials stored neatly off-site or in designated area",
      "Detailed final walkthrough with owner; punch list addressed before sign-off",
    ],
  },
  cost_groups: {
    drywall_walls_ceilings: {
      label: "Drywall – Walls & Ceilings",
      triggers: ["walls", "ceiling", "drywall"],
      standard: [
        "Patch nail holes, dings, and minor cracks",
        "Sand patches smooth and feather edges",
        "Caulk gaps at trim, baseboard, and wall transitions",
        "Spot prime patched and repaired areas",
      ],
      high_end: [
        "Patch nail holes, dings, cracks, and minor drywall imperfections",
        "Skim coat patched areas where needed for a flat finish",
        "Sand all patches smooth and feather edges by hand",
        "Caulk gaps at trim, baseboard, crown, and all wall transitions",
        "Spot prime all bare, patched, and repaired areas",
        "Inspect under raking light before final coat",
      ],
    },
    doors_frames: {
      label: "Doors & Frames",
      triggers: ["door", "frame", "jamb", "casing"],
      standard: [
        "Mask door hardware in place",
        "Scuff sand existing enamel to promote adhesion",
        "Caulk gaps at jambs and casings",
        "Spot prime bare wood and any stained or repaired areas",
      ],
      high_end: [
        "Remove door hardware (knobs, hinges, strike plates) and label for reinstall",
        "Hand sand all surfaces to promote adhesion and smooth existing finish",
        "Fill nail holes and minor dings, sand flush",
        "Caulk gaps at jambs, casings, and where trim meets wall",
        "Spot prime bare wood, stains, and repaired areas with appropriate primer",
        "Sand lightly between coats for a smooth enamel finish",
        "Reinstall hardware upon completion",
      ],
    },
    baseboard_trim: {
      label: "Baseboard & Interior Trim",
      triggers: ["baseboard", "trim", "crown", "casing", "molding"],
      standard: [
        "Fill nail holes and minor gaps",
        "Caulk top edge of baseboard and inside corners",
        "Scuff sand existing enamel to promote adhesion",
        "Spot prime bare wood and repaired areas",
      ],
      high_end: [
        "Fill all nail holes, sand flush",
        "Caulk top edge of baseboard, inside corners, and all trim-to-wall transitions",
        "Hand sand all enameled surfaces to break the gloss",
        "Spot prime bare wood, stains, and repaired areas",
        "Sand lightly between coats for smooth enamel finish",
      ],
    },
    cabinets: {
      label: "Cabinets",
      triggers: ["cabinet", "vanity", "millwork"],
      standard: [
        "Remove doors, drawers, and hardware; label for reinstall",
        "Degrease and TSP wash all surfaces",
        "Scuff sand all surfaces to break existing finish",
        "Spot prime bare wood and stained areas",
        "Reinstall doors, drawers, and hardware upon completion",
      ],
      high_end: [
        "Remove doors, drawers, and hardware; label and store off-site for spray finish",
        "Degrease and TSP wash all surfaces",
        "Fill grain on open-grain woods (oak, ash) where smooth finish is specified",
        "Hand sand all surfaces, including profiles and detail areas",
        "Prime all surfaces with bonding primer",
        "Sand lightly between coats for furniture-grade finish",
        "Spray apply finish coats in controlled environment",
        "Reinstall doors, drawers, and hardware; adjust as needed",
      ],
    },
    exterior_stucco_siding: {
      label: "Exterior – Stucco & Siding",
      triggers: ["stucco", "siding", "exterior wall", "exterior body"],
      standard: [
        "Pressure wash all surfaces to remove dirt, chalk, and loose material",
        "Scrape loose or failing paint",
        "Caulk open gaps at penetrations, transitions, and corners",
        "Spot prime bare areas with appropriate primer",
      ],
      high_end: [
        "Pressure wash all surfaces thoroughly to remove dirt, chalk, mildew, and loose material",
        "Scrape and sand loose or failing paint to a sound edge",
        "Bridging elastomeric or patching compound at hairline cracks",
        "Caulk all open gaps at penetrations, transitions, corners, and trim",
        "Spot prime bare areas with appropriate masonry or bonding primer",
        "Mask windows, doors, lighting, and adjacent surfaces with care",
      ],
    },
    exterior_wood_trim: {
      label: "Exterior – Wood Trim, Fascia, Eaves",
      triggers: ["fascia", "eaves", "exterior trim", "exterior wood", "soffit", "rafter"],
      standard: [
        "Scrape loose paint and sand transitions smooth",
        "Fill nail holes and minor gaps with exterior-grade filler",
        "Caulk gaps at joints and where trim meets siding",
        "Spot prime bare wood with exterior primer",
      ],
      high_end: [
        "Scrape loose paint and hand sand all transitions to a sound edge",
        "Fill all nail holes and minor gaps with exterior-grade filler, sand flush",
        "Replace failed glazing at exterior windows (limited to minor scope)",
        "Caulk all gaps at joints, miters, and where trim meets siding",
        "Spot prime bare wood with exterior-grade oil or alkyd primer for stain blocking",
        "Sand lightly between coats on flat surfaces",
      ],
    },
    fence_deck_stain: {
      label: "Fence & Deck – Stain & Seal",
      triggers: ["fence", "deck", "stain", "seal", "cabot", "sikkens"],
      standard: [
        "Pressure wash to remove dirt, mildew, and loose material",
        "Allow surfaces to dry fully before application",
        "Mask adjacent surfaces and landscaping",
        "Apply stain/sealer per manufacturer specifications",
      ],
      high_end: [
        "Pressure wash thoroughly to remove dirt, mildew, weathered fibers, and loose material",
        "Brighten and neutralize wood with appropriate cleaner if needed",
        "Sand rough or weathered areas smooth",
        "Allow surfaces to dry fully (moisture meter check on premium jobs)",
        "Mask adjacent surfaces, landscaping, and hardscape with care",
        "Apply stain/sealer per manufacturer specifications, back-brushing as needed",
      ],
    },
  },
};

// ============================================================================
// STORAGE HELPERS
// ============================================================================

const STORAGE_KEYS = {
  library: "scope_library_v3",
  catalog: "catalog_ids_v1",
  tiers: "tier_multipliers_v1",
};

async function loadFromStorage(key, fallback) {
  try {
    const result = await window.storage.get(key);
    return result ? JSON.parse(result.value) : fallback;
  } catch {
    return fallback;
  }
}

async function saveToStorage(key, value) {
  try {
    await window.storage.set(key, JSON.stringify(value));
    return true;
  } catch (e) {
    console.error("Storage save failed:", e);
    return false;
  }
}

// ============================================================================
// ITEM NORMALIZATION (handles legacy flat substrate format)
// ============================================================================

function normalizeSubstrate(value, defaultCoats = "2coats", extra = {}) {
  // Already an object — pass through (filling missing fields)
  if (value && typeof value === "object") {
    return { enabled: value.enabled !== false, coats: value.coats || defaultCoats, ...extra, ...value };
  }
  // Legacy boolean (true/false/undefined) — convert
  return { enabled: value !== false, coats: defaultCoats, ...extra };
}

function normalizeItem(item) {
  const type = item.type || "room";
  if (type !== "room") return { ...item, type };

  const fallbackCoats = item.coats || "2coats"; // legacy top-level coats field
  return {
    type: "room",
    name: item.name || "Unnamed Room",
    length: item.length || 0,
    width: item.width || 0,
    height: item.height || 9,
    walls: normalizeSubstrate(item.walls, fallbackCoats),
    ceiling: normalizeSubstrate(item.ceiling, fallbackCoats),
    baseboard: normalizeSubstrate(item.baseboard, fallbackCoats),
    doors: normalizeSubstrate(item.doors, fallbackCoats, { count: item.doorCount || (item.doors && item.doors.count) || 0 }),
    notes: item.notes || "",
  };
}

// Apply a patch to an item, evaluating any string formulas against the item's existing values
// using mathjs. Numeric formulas (e.g. quantity: "210 * 2.5") get evaluated and stored as numbers.
// Non-numeric strings pass through unchanged.
//
// Supports nested merges for room substrate fields. A patch like { doors: { coats: "1coat" } }
// merges into item.doors rather than replacing the whole substrate object.
const ROOM_SUBSTRATE_FIELDS = new Set(["walls", "ceiling", "baseboard", "doors"]);

function applyPatchToItem(item, patch) {
  if (!patch || typeof patch !== "object") return item;
  const next = { ...item };
  for (const [key, value] of Object.entries(patch)) {
    if (key.startsWith("_")) {
      // Carry over meta fields like _formula as plain strings
      next[key] = value;
      continue;
    }
    // Nested merge for room substrate fields: { doors: { coats: "1coat" } }
    // merges with existing item.doors instead of replacing it.
    if (ROOM_SUBSTRATE_FIELDS.has(key) && value && typeof value === "object" && !Array.isArray(value)) {
      const existing = item[key] || {};
      const merged = { ...existing };
      for (const [subKey, subVal] of Object.entries(value)) {
        // Evaluate formulas inside nested patches too
        if (typeof subVal === "string" && /[\+\-\*\/\(\)]/.test(subVal)) {
          try {
            const evaluated = math.evaluate(subVal);
            if (typeof evaluated === "number" && isFinite(evaluated)) {
              merged[subKey] = evaluated;
              continue;
            }
          } catch { /* fall through */ }
        }
        merged[subKey] = subVal;
      }
      next[key] = merged;
      continue;
    }
    if (typeof value === "string" && /[\+\-\*\/\(\)]/.test(value)) {
      // Looks like a formula — try to evaluate it
      try {
        const evaluated = math.evaluate(value);
        if (typeof evaluated === "number" && isFinite(evaluated)) {
          next[key] = evaluated;
          continue;
        }
      } catch {
        // Fall through and store as plain string
      }
    }
    next[key] = value;
  }
  // Re-normalize rooms after applying
  return normalizeItem(next);
}

// ============================================================================
// CLAUDE API HELPERS (with timeout + retry + clear errors)
// ============================================================================

async function callClaude({ model, maxTokens, prompt, timeoutMs = 60000, label = "request" }) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal: controller.signal,
      body: JSON.stringify({
        model,
        max_tokens: maxTokens,
        messages: [{ role: "user", content: prompt }],
      }),
    });

    if (!response.ok) {
      const errText = await response.text().catch(() => "");
      throw new Error(`${label} HTTP ${response.status}: ${errText.slice(0, 200)}`);
    }

    const data = await response.json();

    if (data.type === "error" || data.error) {
      throw new Error(`${label} API error: ${data.error?.message || JSON.stringify(data.error)}`);
    }

    const textBlock = data.content?.find((c) => c.type === "text");
    if (!textBlock) throw new Error(`${label}: no text content in response`);

    const cleaned = textBlock.text.replace(/```json|```/g, "").trim();
    const wasTruncated = data.stop_reason === "max_tokens";

    // Try direct parse first
    try {
      return JSON.parse(cleaned);
    } catch (jsonErr) {
      // If truncated, attempt graceful recovery — keep as many complete array items as possible.
      if (wasTruncated) {
        const salvaged = salvageTruncatedJSON(cleaned);
        if (salvaged) {
          // Mark the result so callers can show a soft warning
          salvaged._truncated = true;
          return salvaged;
        }
        throw new Error(`${label}: response was truncated and could not be recovered. The job might be too complex — try splitting it into smaller sections, or simplifying the notes.`);
      }
      throw new Error(`${label}: invalid JSON — ${jsonErr.message}. First 200 chars: ${cleaned.slice(0, 200)}`);
    }
  } catch (e) {
    if (e.name === "AbortError") {
      throw new Error(`${label} timed out after ${timeoutMs / 1000}s. Try a shorter transcript, or check your network.`);
    }
    throw e;
  } finally {
    clearTimeout(timeout);
  }
}

// Attempt to recover usable content from a truncated JSON response.
// Strategy: walk the string, track bracket/brace depth, and find the last position
// where we could close all open arrays/objects. Keep everything up to there, close them,
// and try to parse. If items[] or questions[] have a half-finished trailing entry,
// drop that entry and try again.
function salvageTruncatedJSON(text) {
  // Quick try: just close any open brackets/braces at the end
  let depth = { brace: 0, bracket: 0 };
  let inString = false;
  let escape = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (escape) { escape = false; continue; }
    if (ch === '\\') { escape = true; continue; }
    if (ch === '"') { inString = !inString; continue; }
    if (inString) continue;
    if (ch === '{') depth.brace++;
    else if (ch === '}') depth.brace--;
    else if (ch === '[') depth.bracket++;
    else if (ch === ']') depth.bracket--;
  }

  // Walk back from the end, drop the trailing incomplete item, then re-close
  // For each "depth", try a few cut-points: the last comma, the last closing brace, etc.
  const candidates = [];
  // Find positions of commas at depth 1 (between top-level array items)
  // Easier heuristic: try cutting at the last `},` or `]` we see, then close.
  for (let cut = text.length; cut > 0; cut--) {
    const ch = text[cut - 1];
    if (ch !== ',' && ch !== '}' && ch !== ']' && ch !== '"' && !/\s/.test(ch)) continue;
    // Compute depth at this cut
    let b = 0, k = 0, ins = false, esc = false;
    for (let i = 0; i < cut; i++) {
      const c = text[i];
      if (esc) { esc = false; continue; }
      if (c === '\\') { esc = true; continue; }
      if (c === '"') { ins = !ins; continue; }
      if (ins) continue;
      if (c === '{') b++;
      else if (c === '}') b--;
      else if (c === '[') k++;
      else if (c === ']') k--;
    }
    if (b < 0 || k < 0) continue;
    // Build closing
    let closing = "";
    // Strip trailing comma if present
    let candidate = text.slice(0, cut).replace(/,\s*$/, "");
    for (let i = 0; i < b; i++) closing += "}";
    for (let i = 0; i < k; i++) closing += "]";
    candidates.push(candidate + closing);
    if (candidates.length > 12) break;  // don't try forever
  }

  for (const c of candidates) {
    try {
      return JSON.parse(c);
    } catch { /* try next */ }
  }
  return null;
}

function buildJobInfoPrompt(transcript) {
  return `Extract job metadata from this painting job walk transcript. Return ONLY a JSON object with shape { "jobInfo": {...}, "flags": {...} }, no preamble or markdown fences.

jobInfo:
{
  customerName: { value: string, confidence: "high"|"medium"|"low"|"unknown" },
  address:      { value: string, confidence: "high"|"medium"|"low"|"unknown" },
  email:        { value: string, confidence: "high"|"medium"|"low"|"unknown" },
  phone:        { value: string, confidence: "high"|"medium"|"low"|"unknown" },
  tier:         { value: "standard"|"production"|"highend"|"prevailing", confidence: "high"|"medium"|"low"|"unknown" }
}

Tier signals:
- "production" / "PM" / "property management" / "rental" → "production"
- "custom home" / "designer" / "luxury" / "Emerald Urethane" → "highend"
- "DIR" / "prevailing wage" / "public works" / "school district" → "prevailing"
- Otherwise → "standard" confidence "unknown"

flags (infer from transcript):
{
  vacant:         { value: bool, confidence: "high"|"medium"|"low"|"unknown" },
  petsKids:       { value: bool, confidence: "high"|"medium"|"low"|"unknown" },
  leadPaint:      { value: bool, confidence: "high"|"medium"|"low"|"unknown" },
  activeBusiness: { value: bool, confidence: "high"|"medium"|"low"|"unknown" }
}

If a field isn't in the transcript, set value to "" (or false for flags) and confidence to "unknown".

Transcript:
${transcript}`;
}

function buildItemsPrompt(transcript) {
  return `Parse this painting job walk transcript into line items and clarifying questions. Return ONLY a JSON object with shape { "items": [...], "questions": [...] }, no preamble or markdown fences.

ITEM TYPES:

"room" (physical room being painted) — fields:
{
  type: "room", name, length, width, height (default 9),
  walls:     { enabled: bool, coats: "1coat"|"2coats"|"prime+2" },
  ceiling:   { enabled: bool, coats: "1coat"|"2coats"|"prime+2" },
  baseboard: { enabled: bool, coats: "1coat"|"2coats"|"prime+2" },
  doors:     { enabled: bool, coats: "1coat"|"2coats"|"prime+2", count: number },
  notes
}
COATS ARE PER SUBSTRATE. "one coat ceiling, two coats walls" → ceiling.coats="1coat", walls.coats="2coats". Default unstated substrates to enabled=true and 2 coats. Always include all four substrate objects.

"scope" (substrate bucket, not a physical room — Doors throughout, Cabinets, Exterior Stucco, Fence, Window Trim, Metal Beams, etc.) — fields:
{ type: "scope", name, substrate: "doors"|"baseboard"|"trim"|"cabinets"|"exterior_stucco"|"exterior_trim"|"fence_deck"|"metal", quantity, unit: "EA"|"LF"|"SF", coats, notes }

"tm" (Time & Materials hourly work — ONLY for true touch-up / spot work that has no measurable quantity) — fields:
{ type: "tm", name, hours, category: "interior_walls"|"doors"|"windows"|"trim"|"exterior_walls"|"exterior_misc"|"millwork"|"repairs"|"cleaning"|"materials", notes }

═══════════════════════════════════════════════
PURPLE PAINTING CONVENTIONS (use these defaults — do NOT guess)
═══════════════════════════════════════════════

UNITS by substrate:
- Wood/window trim, baseboard, crown molding, casing → LF (linear feet)
- Metal beams, columns, posts, railings → LF (use the height as the length)
- Doors, cabinet doors, drawers → EA (each)
- Cabinets (boxes), built-ins, vanities → EA with size note (small/medium/large)
- Walls, ceilings, stucco, siding → SF (square feet)
- Fence, deck → LF or SF depending on context

COATS by scenario:
- Color change → 2 coats
- Refresh / same color recoat → 1 coat
- New unprimed wood, drywall, or stain-to-paint conversion → prime+2
- Touch-ups → 1 coat
- Default if unstated → 2 coats

ITEM TYPE CHOICES (Scope vs T&M):
- ANY measurable substrate (trim, beams, doors, cabinets, walls) → "scope" with quantity + unit
- Use "tm" ONLY for hourly punch-list work like "touch-ups throughout" or "repair drywall damage as needed"
- A metal beam, even if you don't know its exact length, is "scope" (LF) with notes about needing measurement — NOT "tm"
- Window wood trim is "scope" (LF) — NOT "tm"
- Cabinets are "scope" (EA) — NOT "tm"

WINDOW TRIM QUANTITY CALCULATION:
- An 8ft tall × 6ft wide window has a perimeter of 28 LF (8+8+6+6) if all four sides have trim
- If transcript says "window with wood trim" without specifying sides, propose 28 LF (perimeter) as default

METAL BEAMS / LINEAR ELEMENTS:
- If transcript states height/length, use that as the LF quantity
- "12ft tall metal beam" → scope, metal, 12 LF
- Default to 1 coat unless transcript says color change or refresh

CABINETS:
- Cabinet items are sized: Small, Medium, Large
- A "small bathroom vanity" or "small cabinet" → Small
- A stain-to-paint conversion → coats = "prime+2", note "stain to paint conversion"
- A "small cabinet, stain-to-paint, 2 coats" → quantity=1, unit=EA, coats="prime+2", note about conversion

═══════════════════════════════════════════════
EXTERIOR STUCCO / SIDING / BODY — CRITICAL QUANTITY RULE
═══════════════════════════════════════════════

Exterior body (stucco, siding, exterior walls) is ALWAYS measured in SQUARE FEET = perimeter × wall height.
A bare perimeter number is NOT the SF quantity. NEVER store the raw perimeter as the SF quantity for exterior walls.

When you see "exterior body", "stucco", "siding", "exterior walls" — three numbers matter:
1. PERIMETER (in linear feet around the building, including walls behind any garage/casita)
2. WALL HEIGHT (typically 9-10 ft for single-story, 18-20 ft for two-story)
3. Resulting SF = perimeter × wall height

If the transcript gives a perimeter but NOT a wall height, you MUST emit a clarifying question with "field": "dimensions" and confidence "high" asking for the wall height. Default options should offer common heights with the formula pre-filled. Example:

If the transcript said "210 LF perimeter, exterior body" but did not specify wall height:
{
  id: "ext-body-height-1",
  itemIndex: <index of the exterior body scope item>,
  field: "dimensions",
  prompt: "Wall height for the exterior body? Quantity will be calculated as perimeter × wall height.",
  appliesTo: "Exterior Body / Stucco",
  confidence: "high",
  options: [
    { label: "9 ft (single story)",  value: { quantity: "210 * 9",  _formula: "210 LF perimeter × 9 ft wall height",  unit: "SF" } },
    { label: "10 ft (single story, taller ceilings)", value: { quantity: "210 * 10", _formula: "210 LF perimeter × 10 ft wall height", unit: "SF" } },
    { label: "18 ft (two story)",    value: { quantity: "210 * 18", _formula: "210 LF perimeter × 18 ft wall height", unit: "SF" } },
    { label: "Other / custom",       value: "__custom__" }
  ]
}

Until that question is answered, set the scope item's quantity to the perimeter as a placeholder string AND mark the unit as "SF (needs height)" in notes so the user sees it's incomplete. Example item:
{ type: "scope", name: "Exterior Body", substrate: "exterior_stucco", quantity: 210, unit: "SF", coats: "2coats", notes: "PLACEHOLDER: 210 is perimeter — multiply by wall height to get actual SF" }

If the transcript ALREADY gave both perimeter and height (e.g. "210 LF around the house, 10ft walls"), compute it yourself and emit the scope item with quantity as a formula string: quantity: "210 * 10", unit: "SF", and DO NOT generate a clarifying question.

═══════════════════════════════════════════════
EXTERIOR EAVES / SOFFITS / OVERHANGS — separate formula
═══════════════════════════════════════════════

Exterior eaves / soffits / overhangs (the horizontal surface under the roof overhang as you look up from below) = eave length × eave width.

  - EAVE LENGTH (LF, sometimes called "eave run" or sum of "elevation lengths"): how far the overhang runs along the roof edges. NOT the same as the body perimeter — some elevations may have no overhang (e.g. a wall against a neighbor, a flat parapet wall, a shed addition). If the estimator hasn't specified, the default is to use the body perimeter as an approximation, but flag it.
  - EAVE WIDTH (ft, sometimes called "overhang depth" or "soffit depth"): how far the overhang extends out from the wall. Typical: 1.5 ft, 2 ft, 2.5 ft, 3 ft.
  - Quantity formula: eave_length × eave_width, NOT perimeter × wall_height.

If exterior eaves are mentioned with NEITHER eave length nor eave width specified, emit a high-confidence clarifying question that asks for BOTH at once (use a single combined question rather than two separate ones). Example:

{
  id: "eaves-dims-1",
  itemIndex: <index of eaves scope item>,
  field: "dimensions",
  prompt: "Eave dimensions? Quantity = eave length × eave width. If some elevations have no overhang, use just the lengths that do.",
  appliesTo: "Exterior Eaves / Soffits",
  confidence: "high",
  options: [
    { label: "210 LF × 2 ft (use body perimeter, 24\" overhang)",  value: { quantity: "210 * 2",   _formula: "210 LF eave length × 2 ft eave width",   unit: "SF" } },
    { label: "210 LF × 2.5 ft (use body perimeter, 30\" overhang)", value: { quantity: "210 * 2.5", _formula: "210 LF eave length × 2.5 ft eave width", unit: "SF" } },
    { label: "210 LF × 3 ft (use body perimeter, 36\" overhang)",   value: { quantity: "210 * 3",   _formula: "210 LF eave length × 3 ft eave width",   unit: "SF" } },
    { label: "Other / custom",                                      value: "__custom__" }
  ]
}

If only the eave length is given (e.g. "180 LF of eaves") but no width, ask only for the width with options 2 ft, 2.5 ft, 3 ft. If only the width is given, ask only for the length and offer body-perimeter as the suggested default.

NEVER use wall height as an eave dimension. NEVER use a single perimeter number as the SF of eaves on its own.

═══════════════════════════════════════════════
CLARIFYING QUESTIONS
═══════════════════════════════════════════════

Generate questions for anything you had to guess. Each question MUST include 2-4 clickable options so the user can resolve it with one tap. Shape:

{
  id: string,
  itemIndex: number|null,
  field: "dimensions"|"doorCount"|"coats"|"quantity"|"unit"|"type"|"size"|"ceiling_height"|"substrate"|"other",
  prompt: "short question text",
  appliesTo: "<which item or area>",
  confidence: "high"|"medium"|"low",
  options: [
    { label: "Display text user sees on the button", value: <object with fields to apply to the item> },
    { label: "...", value: {...} },
    { label: "Other / custom", value: "__custom__" }  // ALWAYS include this last option
  ]
}

The "value" in each option is the patch to apply to the item. Examples:

Coats question:
{
  options: [
    { label: "1 coat (refresh)", value: { coats: "1coat" } },
    { label: "2 coats (color change)", value: { coats: "2coats" } },
    { label: "Prime + 2", value: { coats: "prime+2" } },
    { label: "Other / custom", value: "__custom__" }
  ]
}

Unit/type question for window trim:
{
  options: [
    { label: "28 LF (full perimeter)", value: { type: "scope", substrate: "trim", quantity: 28, unit: "LF" } },
    { label: "14 LF (sides + top only)", value: { type: "scope", substrate: "trim", quantity: 14, unit: "LF" } },
    { label: "Keep as T&M", value: { type: "tm", hours: 2 } },
    { label: "Other / custom", value: "__custom__" }
  ]
}

Size question for cabinets:
{
  options: [
    { label: "Small", value: { notes: "Size: Small" } },
    { label: "Medium", value: { notes: "Size: Medium" } },
    { label: "Large", value: { notes: "Size: Large" } },
    { label: "Other / custom", value: "__custom__" }
  ]
}

Dimensions question for a room where one wall length is unclear:
{
  options: [
    { label: "12 × 14 (12 length, 14 width)", value: { length: 12, width: 14 } },
    { label: "14 × 12 (14 length, 12 width)", value: { length: 14, width: 12 } },
    { label: "Other / custom", value: "__custom__" }
  ]
}

═══════════════════════════════════════════════
ADVANCED OPTION SHAPES — compute formulas + multi-item updates
═══════════════════════════════════════════════

When the answer should COMPUTE a quantity from a formula, use this option shape. Note the DIFFERENT formulas: body uses WALL HEIGHT, eaves use EAVE WIDTH (overhang depth).

{
  label: "210 LF body perimeter, 10ft walls, 180 LF eave run × 2.5ft eave width",
  value: {
    _multiPatch: [
      { itemNameMatch: "exterior body",        patch: { quantity: "210 * 10",  _formula: "body perimeter × wall height",      unit: "SF" } },
      { itemNameMatch: "exterior fascia",      patch: { quantity: 180,         _formula: "eave run length",                   unit: "LF" } },
      { itemNameMatch: "exterior eaves",       patch: { quantity: "180 * 2.5", _formula: "eave run × eave width",             unit: "SF" } }
    ]
  }
}

When the answer applies a setting ACROSS A SUBSTRATE TYPE (e.g. "1 coat for doors throughout the whole house"), use substrateMatch:

{
  label: "1 coat for all doors",
  value: {
    _multiPatch: [
      { substrateMatch: "doors", patch: { coats: "1coat" } }
    ]
  }
}

This applies { coats: "1coat" } to EVERY room's doors substrate AND every scope item where substrate="doors", in one click. For multi-substrate answers, include one entry per substrate:

{
  label: "1 coat for doors, trim, and ceilings",
  value: {
    _multiPatch: [
      { substrateMatch: "doors",    patch: { coats: "1coat" } },
      { substrateMatch: "trim",     patch: { coats: "1coat" } },
      { substrateMatch: "ceilings", patch: { coats: "1coat" } }
    ]
  }
}

substrateMatch values: "walls" | "ceilings" | "baseboard" | "doors" | "trim" | "cabinets" | "exterior_stucco" | "exterior_trim" | "fence_deck" | "metal"

Rules:
- "_multiPatch" applies the same answer to multiple items in one click
- Each entry targets items via "itemNameMatch" (case-insensitive substring of item name) AND/OR "substrateMatch" (substrate type)
- For substrateMatch on a room, the patch is auto-wrapped into the substrate field (e.g. coats: "1coat" → doors.coats: "1coat")
- "quantity" can be a STRING formula (e.g. "210 * 2.5") — the app evaluates it safely with mathjs
- Include "_formula" with a short human-readable note (shown as a hint after the value)
- Use this whenever an answer logically updates more than one item, OR when the answer is a calculation, OR when the answer applies "throughout" / "across all" / "everywhere"

WHEN TO USE substrateMatch vs itemNameMatch:
- "1 coat for doors throughout" → substrateMatch: "doors"
- "Doors in the office only" → itemNameMatch: "office", substrateMatch: "doors"
- "Master bedroom only" → itemNameMatch: "master bedroom"
- "Exterior body and eaves" → substrateMatch: "exterior_stucco" + "exterior_trim"

When the answer ONLY updates the single item the question is about, keep the simple shape (value is a plain patch object).

═══════════════════════════════════════════════
MANDATORY QUESTIONS (always emit, even if it pushes the budget)
═══════════════════════════════════════════════

Some gaps are NEVER acceptable to leave silently. For each of these, emit a clarifying question with confidence "high" no matter what:

1. **Missing room dimensions.** Any "room" item where length or width is 0, null, undefined, or missing → MUST ask. Use field "dimensions". Offer 2-3 common sizes for that room type as quick-pick options plus "Other / custom". Examples:
   - Bathroom missing dims → options "5 × 8 (small)", "7 × 10 (typical)", "10 × 12 (master)", "Other / custom"
   - Bedroom missing dims → options "10 × 12 (small)", "12 × 14 (typical)", "14 × 16 (primary)", "Other / custom"
   - Closet missing dims → options "4 × 6 (reach-in)", "6 × 8 (walk-in)", "Other / custom"
   - Generic room with unknown type → options "10 × 12", "12 × 14", "14 × 16", "Other / custom"
   Option value shape: { length: <n>, width: <n> }

2. **Exterior body without wall height** (covered in EXTERIOR BODY section above).

3. **Exterior eaves without eave length OR eave width** (covered in EXTERIOR EAVES section above).

4. **Door count = 0 or missing for a room that has doors enabled.** Ask for door count.

These four MANDATORY question types take priority over the soft 4-question budget. If you have 5 mandatory issues, emit 5 questions. The soft limit only applies to discretionary clarifications (coats, refresh vs color change, etc).

═══════════════════════════════════════════════
RULES FOR OPTIONS:
═══════════════════════════════════════════════
- ALWAYS provide 2-3 concrete proposed options, not "Unknown" placeholders
- ALWAYS include "Other / custom" as the last option
- Base the proposed options on Purple Painting conventions above
- Make labels short and informative so the user knows the consequence at a glance
- Use _multiPatch + formula strings whenever the answer derives multiple quantities from a measurement

Transcript:
${transcript}`;
}

function buildNotesPrompt(notes) {
  return `Parse terse estimator notes into line items. Return ONLY JSON: { "items": [...], "questions": [...] }, no preamble or fences.

ITEM SHAPES:
"room": { type:"room", name, length, width, height (default 9),
  walls:{enabled,coats}, ceiling:{enabled,coats}, baseboard:{enabled,coats}, doors:{enabled,coats,count}, notes }
  — coats: "1coat" | "2coats" | "prime+2"
"scope": { type:"scope", name, substrate, quantity, unit, coats, notes }
  — substrate: doors | baseboard | trim | cabinets | exterior_stucco | exterior_trim | fence_deck | metal | walls | ceilings
  — unit: EA | LF | SF
"tm":    { type:"tm", name, hours, category, notes }
  — category: interior_walls | doors | windows | trim | exterior_walls | exterior_misc | millwork | repairs | cleaning | materials

DEFAULTS (apply silently):
- Trim/baseboard → LF · Doors → EA · Walls/ceilings/stucco → SF
- Coats unstated → 2coats · Height unstated → 9 · Substrate enabled → true if mentioned

EXAMPLES:
- "3BR ~14x12" → 3 room items "Bedroom 1/2/3", 14x12x9, all substrates enabled at 2coats, doors count=1
- "doors and trim only in hall" → room "Hall" with walls/ceiling disabled, baseboard + doors enabled
- "exterior body + fascia, 1 coat refresh" → 2 scope items: exterior_stucco SF 1coat, exterior_trim LF 1coat

QUESTIONS: only for true ambiguities ("the bathroom" → which/how many/size?). Format:
{ id, itemIndex, field, prompt, appliesTo, confidence, options:[{label,value}, ..., {label:"Other / custom", value:"__custom__"}] }

OPTION VALUE — usually a plain patch object like { coats:"1coat" }. For answers that update multiple items, use _multiPatch:

Substrate-wide ("1 coat for doors throughout"):
{ _multiPatch: [{ substrateMatch:"doors", patch:{ coats:"1coat" } }] }

Multi-substrate ("1 coat for doors, trim, ceilings"):
{ _multiPatch: [
  { substrateMatch:"doors", patch:{ coats:"1coat" } },
  { substrateMatch:"trim", patch:{ coats:"1coat" } },
  { substrateMatch:"ceilings", patch:{ coats:"1coat" } }
] }

Formula-based ("210 LF perimeter × 2.5ft eaves"):
{ _multiPatch: [{ itemNameMatch:"exterior body", patch:{ quantity:"210 * 2.5", _formula:"perimeter × eave depth", unit:"SF" }}] }

- substrateMatch: "walls" | "ceilings" | "baseboard" | "doors" | "trim" | "cabinets" | "exterior_stucco" | "exterior_trim" | "fence_deck" | "metal"
- itemNameMatch: case-insensitive substring of item name
- patches with substrateMatch auto-wrap into the substrate field on rooms (coats:"1coat" → doors.coats:"1coat")
- quantity can be a string formula (mathjs evaluated)
- Use substrateMatch for "throughout"/"all"/"every", itemNameMatch for specific rooms

Trust shorthand. Generate AT MOST 3 questions. Be aggressive with defaults.

Notes:
${notes}`;
}

// Combined prompt: returns jobInfo + flags + items + questions in ONE call.
// Used in Notes mode to eliminate the round-trip of two parallel API calls.
function buildCombinedNotesPrompt(notes) {
  return `Parse terse painting estimator notes. Return ONLY JSON with this exact shape:
{ "jobInfo": {...}, "flags": {...}, "items": [...], "questions": [...] }

No preamble, no markdown fences.

jobInfo: { customerName, address, email, phone, tier } — each is { value, confidence: "high"|"medium"|"low"|"unknown" }
  tier value: "standard" | "production" | "highend" | "prevailing"
  tier signals: "PM"/"rental"→production · "custom home"/"luxury"/"designer"/"Emerald"→highend · "DIR"/"prevailing"/"public works"/"school"→prevailing · else→standard
  If a field isn't in the notes, value:"" and confidence:"unknown"

flags: { vacant, petsKids, leadPaint, activeBusiness } — each { value:bool, confidence }

items: ITEM SHAPES
"room": { type:"room", name, length, width, height (default 9), walls:{enabled,coats}, ceiling:{enabled,coats}, baseboard:{enabled,coats}, doors:{enabled,coats,count}, notes }
  coats: "1coat" | "2coats" | "prime+2"
"scope": { type:"scope", name, substrate, quantity, unit, coats, notes }
  substrate: doors|baseboard|trim|cabinets|exterior_stucco|exterior_trim|fence_deck|metal|walls|ceilings
  unit: EA|LF|SF
"tm": { type:"tm", name, hours, category, notes }
  category: interior_walls|doors|windows|trim|exterior_walls|exterior_misc|millwork|repairs|cleaning|materials

DEFAULTS (apply silently):
- Trim/baseboard→LF · Doors→EA · Walls/ceilings/stucco→SF
- Coats unstated→2coats · Height unstated→9 · Substrate enabled→true if mentioned

EXAMPLES:
"3BR ~14x12" → 3 rooms "Bedroom 1/2/3", 14x12x9, all substrates enabled at 2coats, doors count=1
"doors and trim only in hall" → room "Hall", walls/ceiling disabled, baseboard+doors enabled
"exterior body + fascia, 1 coat refresh" → 2 scope items: exterior_stucco SF 1coat, exterior_trim LF 1coat

EXTERIOR BODY / STUCCO — CRITICAL:
exterior_stucco quantity in SF = body perimeter (LF) × wall height (ft). Perimeter alone is NOT the SF quantity.
If transcript gives perimeter but NO wall height, emit the scope item with quantity=perimeter and notes "PLACEHOLDER — needs wall height", AND emit a high-confidence question asking for wall height with options "9 ft (1 story)", "10 ft (1 story tall)", "18 ft (2 story)", "Other/custom". The option's patch should be { quantity:"<perimeter> * <height>", _formula:"<perimeter> × <height>", unit:"SF" }.
If transcript gives BOTH (e.g. "210 LF around, 10ft walls"), compute it: quantity:"210 * 10", unit:"SF". No question needed.

EXTERIOR EAVES — different formula:
exterior eaves/soffits SF = eave run length (LF) × eave width (ft, the overhang depth). NOT wall height. NOT just perimeter alone.
Eave run length is NOT always the same as body perimeter — some elevations may have no overhang. Default: use body perimeter as approximation but flag it.
Eave width is typically 1.5, 2, 2.5, or 3 ft. If unspecified, ask for it; offer options "210 × 2 ft", "210 × 2.5 ft", "210 × 3 ft" where 210 = body perimeter default.
NEVER use wall height for eaves. NEVER store a bare perimeter as SF for eaves.

questions: shape { id, itemIndex, field, prompt, appliesTo, confidence, options:[{label,value}, ..., {label:"Other / custom", value:"__custom__"}] }

MANDATORY questions (ALWAYS emit, no budget limit):
1. Any room with missing/0 length OR width → ask dimensions. Bathroom defaults: "5×8 small", "7×10 typical", "10×12 master". Bedroom: "10×12", "12×14", "14×16". Closet: "4×6", "6×8". Generic: "10×12", "12×14", "14×16". Option value shape: { length:<n>, width:<n> }.
2. Exterior body without wall height (see CRITICAL above).
3. Exterior eaves without eave length OR eave width (see EAVES above).
4. Room with doors enabled but doorCount=0/missing → ask count.

DISCRETIONARY questions (apply 4-question budget): coats ambiguity, color-change vs refresh, etc.

OPTION VALUE — usually a plain patch like { coats:"1coat" }. For multi-item answers, use _multiPatch:
Substrate-wide: { _multiPatch:[{ substrateMatch:"doors", patch:{ coats:"1coat" }}] }
Multi-substrate: { _multiPatch:[{ substrateMatch:"doors", patch:{ coats:"1coat" }},{ substrateMatch:"trim", patch:{ coats:"1coat" }}] }
Wall-height formula (body): { _multiPatch:[{ itemNameMatch:"exterior body", patch:{ quantity:"210 * 10", _formula:"body perimeter × wall height", unit:"SF" }}] }
Eave-width formula (eaves/soffits): { _multiPatch:[{ itemNameMatch:"exterior eaves", patch:{ quantity:"180 * 2.5", _formula:"eave run × eave width", unit:"SF" }}] }

Use substrateMatch for "throughout"/"all"/"every", itemNameMatch for specific rooms. Patches with substrateMatch auto-wrap into the room substrate field. quantity can be a string formula (mathjs).

Trust shorthand for the discretionary 4-question budget. Mandatory questions (missing room dims, body wall height, eaves dims, door counts) ALWAYS get emitted regardless.

Notes:
${notes}`;
}

function buildTakeoffPrompt(takeoff) {
  // No longer used — takeoff mode parses XLSX deterministically in-browser.
  // Kept here as a fallback in case we need to handle a non-standard pasted format.
  return `Parse this digital takeoff into scope items. Return ONLY JSON: { "items": [...], "questions": [] }. Each item: { type: "scope", name, substrate, quantity, unit, coats: "2coats", notes }.

Takeoff:
${takeoff}`;
}

// ============================================================================
// INPUT AUTO-DETECT (classifier + heuristic fallback) + NOTES OVERLAY
// ============================================================================
//
// Used by parseTranscript() to figure out, for the text-only path, whether
// the user pasted a voice transcript or terse estimator notes. The classifier
// is a tiny Haiku call (~200 tokens, 10s timeout). If the call fails, we fall
// back to a length + filler-word heuristic so the app never gets stuck.

function classifyHeuristic(text) {
  const t = text || "";
  const len = t.length;
  const fillerRe = /\b(um|uh|so|okay|ok|like|y'?know|i mean|basically|right\?|alright)\b/gi;
  const fillerCount = (t.match(fillerRe) || []).length;
  if (len > 800 && fillerCount >= 2) {
    return { mode: "voice", rationale: `length ${len} chars, ${fillerCount} filler words — looks like a transcript` };
  }
  return { mode: "notes", rationale: `length ${len} chars, ${fillerCount} filler words — looks like terse notes` };
}

function buildClassifierPrompt(text) {
  return `Classify this painting estimator input as either a "voice" transcript or terse "notes". Return ONLY JSON: {"mode":"voice"|"notes","rationale":"short reason"}, no preamble or markdown fences.

"voice" indicators: full sentences, filler words (um, uh, so, okay, like), narrative flow, conversational tone. Usually >800 chars. The estimator was talking through a job walk.

"notes" indicators: terse shorthand, abbreviations, measurements inline (e.g. "3BR ~14x12, 9ft, 2 coats walls"), bullet-like fragments, no filler words. Usually <800 chars.

Input:
${(text || "").slice(0, 2000)}`;
}

async function classifyInput(text) {
  return callClaude({
    model: "claude-haiku-4-5-20251001",
    maxTokens: 200,
    timeoutMs: 10000,
    prompt: buildClassifierPrompt(text),
    label: "classify",
  });
}

// Notes overlay: when the user uploads a takeoff AND pastes accompanying scope
// notes, this prompts Claude to adjust the parsed takeoff against the notes.
// Returns { overrides, additions, exclusions, questions }.
function buildOverlayPrompt(items, notesText) {
  const summary = items.map((it, i) => ({
    index: i,
    name: it.name,
    substrate: it.substrate || it.type,
    quantity: it.quantity,
    unit: it.unit,
    coats: it.coats,
  }));
  return `You are adjusting a digital takeoff with accompanying estimator notes. Return ONLY JSON: { "overrides": [...], "additions": [...], "exclusions": [...], "questions": [...] }, no preamble or markdown fences.

Existing takeoff items (DO NOT duplicate these — adjust via overrides):
${JSON.stringify(summary, null, 2)}

Estimator notes overlay:
${notesText}

Output shape:
- "overrides":  [{ "itemIndex": <number from above>, "patch": { ...fields to change on that item, e.g. coats, quantity, notes } }]
- "additions":  [<full new scope item objects>] — ONLY when notes describe substrates not already in the takeoff
- "exclusions": [<string>, ...] — exclusion lines mentioned in the notes
- "questions":  [<clarifying question objects, same shape as Notes mode>] — only true ambiguities

Be CONSERVATIVE with additions — most overlay notes adjust existing items (e.g. "1 coat for doors throughout") which should be overrides, not additions. patch.quantity may be a string formula (evaluated with mathjs).`;
}

async function applyNotesOverlay(items, notesText) {
  return callClaude({
    model: "claude-haiku-4-5-20251001",
    maxTokens: 4000,
    timeoutMs: 45000,
    prompt: buildOverlayPrompt(items, notesText),
    label: "notes overlay",
  });
}

// ============================================================================
// XLSX TAKEOFF PARSER (deterministic — no AI)
// ============================================================================
//
// Handles Purple Painting's standard BoQ format:
//   - Sheet named "BoQ" (with or without trailing space), or first sheet
//   - Header row found by looking for cells containing "description"+"unit"+"quantity"
//     (also matches "item"+"measure"+"qty" for the school-district variant)
//   - Section headers (text-only rows like "Walls", "Ceiling", "Door and Frames",
//     "Exterior") provide context for downstream substrate mapping
//   - Skips: blank rows, "sub total", "GRAND TOTAL", "BASE BID", "Painting with..."
//   - Captures items under "Exclusions" section into a separate list
//
// Returns: { items: [...], exclusions: [...], warnings: [...] }

const SECTION_TO_SUBSTRATE = {
  // direct section name → substrate key
  "walls": "walls",
  "wall": "walls",
  "ceiling": "ceilings",
  "ceilings": "ceilings",
  "base": "baseboard",
  "baseboard": "baseboard",
  "door and frames": "doors",
  "doors and frames": "doors",
  "interior doors": "doors",
  "trim casing": "trim",
  "window frames / casing": "trim",
  "window frame (paint interior side)": "trim",
  "exterior": "exterior_stucco",
  "miscellaneous item paint": null,    // ambiguous — let description-based mapping decide
  "vinyl wall cover": "walls",
  "wallcovering": "walls",
  "wood staining": "trim",
  "cabinet": "cabinets",
  "cabinets": "cabinets",
};

// Description-based substrate inference (used when section name is generic or missing)
const DESCRIPTION_KEYWORDS = [
  { pattern: /\b(door|frame).*\bunit\b/i, substrate: "doors" },
  { pattern: /\bdoor\b/i,                  substrate: "doors" },
  { pattern: /\bbaseboard|\bbase\b/i,      substrate: "baseboard" },
  { pattern: /\bcasing|window trim|crown|trim\b/i, substrate: "trim" },
  { pattern: /\bcabinet|vanity|millwork\b/i, substrate: "cabinets" },
  { pattern: /\bstucco|siding|exterior.*wall|exterior.*body\b/i, substrate: "exterior_stucco" },
  { pattern: /\bfascia|eaves|soffit|exterior.*wood|exterior.*trim\b/i, substrate: "exterior_trim" },
  { pattern: /\bfence|deck\b/i,            substrate: "fence_deck" },
  { pattern: /\bbeam|column|railing|metal\b/i, substrate: "metal" },
  { pattern: /\bceiling|soffit\b/i,        substrate: "ceilings" },
  { pattern: /\bwall|gypsum|drywall|plaster\b/i, substrate: "walls" },
];

function inferSubstrate(description, currentSection) {
  const sect = (currentSection || "").trim().toLowerCase();
  if (sect && SECTION_TO_SUBSTRATE.hasOwnProperty(sect)) {
    const mapped = SECTION_TO_SUBSTRATE[sect];
    if (mapped) return mapped;
  }
  // Combined section like "Walls and ceilings", "Door and Casing"
  if (sect) {
    const combined = inferSubstrateFromCombined(sect);
    if (combined) return combined;
  }
  // Fall back to description keyword scan
  const desc = (description || "").toLowerCase();
  for (const { pattern, substrate } of DESCRIPTION_KEYWORDS) {
    if (pattern.test(desc)) return substrate;
  }
  // Last resort
  return "walls";
}

function normalizeUnit(rawUnit) {
  if (!rawUnit) return "EA";
  const u = String(rawUnit).trim().toLowerCase().replace(/\./g, "");
  if (["sf", "sqft", "sq ft"].includes(u)) return "SF";
  if (["lf", "linft", "lin ft", "ft"].includes(u)) return "LF";
  if (["ea", "each", "unit", "units", "qty"].includes(u)) return "EA";
  if (["hr", "hour", "hours"].includes(u)) return "HR";
  if (["gal", "gallon", "gallons"].includes(u)) return "GAL";
  return "EA"; // safe fallback
}

function isSectionHeader(text) {
  if (!text) return false;
  const t = String(text).trim().toLowerCase();
  if (!t) return false;
  // Common skip patterns
  if (t.startsWith("sub total") || t.startsWith("grand total") || t.startsWith("base bid")) return false;
  if (t.startsWith("painting with") || t.startsWith("no.") || t === "description") return false;
  if (t.startsWith("project:") || t.startsWith("division:") || t.startsWith("rev:") || t.startsWith("date:")) return false;
  // Section headers are typically short. If it's > 60 chars OR contains size-like patterns
  // (e.g. "3'x7'", "(paint interior side)"), it's probably a stray item description with
  // missing unit/qty, not a real section header.
  if (t.length > 60) return false;
  if (/\d['"]?\s*x\s*\d/i.test(t)) return false;  // size pattern like 3'x7' or 5x7
  if (/\(paint\b/i.test(t)) return false;          // descriptive paint annotation
  return true;
}

// Add fallback substrate mapping for combined section names like "Walls and ceilings"
function inferSubstrateFromCombined(sect) {
  const s = sect.toLowerCase();
  if (s.includes("wall") && s.includes("ceiling")) return null; // ambiguous, let description decide
  if (s.includes("door") && s.includes("casing")) return "doors";
  if (s.includes("window") && s.includes("frame")) return "trim";
  if (s.includes("trellis") || s.includes("railing")) return "metal";
  if (s.includes("storefront")) return "metal";
  if (s.includes("facia") || s.includes("fascia")) return "exterior_trim";
  return null;
}

function findHeaderRow(sheetData) {
  // sheetData is an array of row arrays.
  // Look for a row containing description/name/item AND unit/measure AND quantity/qty
  for (let r = 0; r < Math.min(15, sheetData.length); r++) {
    const row = sheetData[r] || [];
    const rowText = row.map(c => (c == null ? "" : String(c).toLowerCase())).join(" | ");
    const hasDesc = /\b(description|name|item)\b/.test(rowText);
    const hasUnit = /\b(unit|units|measure)\b/.test(rowText);
    const hasQty  = /\b(quantity|qty)\b/.test(rowText);
    if (hasDesc && hasUnit && hasQty) {
      return { rowIndex: r, headers: row.map(c => (c == null ? "" : String(c).trim().toLowerCase())) };
    }
  }
  return null;
}

function findColumns(headers) {
  // Returns { desc, unit, qty, rate, amount, note } as 0-based column indices, or null if not found
  const cols = { desc: -1, unit: -1, qty: -1, rate: -1, amount: -1, note: -1 };
  headers.forEach((h, i) => {
    if (h === "description" || h === "name" || h === "item") cols.desc = i;
    else if (h === "unit" || h === "units" || h === "measure") cols.unit = i;
    else if (h === "quantity" || h === "qty") {
      // Prefer the FIRST quantity column (revised columns come after)
      if (cols.qty === -1) cols.qty = i;
    }
    else if (h === "rate") {
      if (cols.rate === -1) cols.rate = i;
    }
    else if (h === "amount" || h === "price") {
      if (cols.amount === -1) cols.amount = i;
    }
    else if (h === "note" || h === "notes") cols.note = i;
  });
  return cols;
}

function parseTakeoffSheet(sheetData, sheetName) {
  const result = { items: [], exclusions: [], warnings: [] };
  const header = findHeaderRow(sheetData);
  if (!header) {
    result.warnings.push(`Sheet "${sheetName}": no header row found (need columns: description, unit, quantity)`);
    return result;
  }
  const cols = findColumns(header.headers);
  if (cols.desc === -1 || cols.unit === -1 || cols.qty === -1) {
    result.warnings.push(`Sheet "${sheetName}": missing required columns (need description, unit, quantity)`);
    return result;
  }

  let currentSection = "";
  let inExclusions = false;

  for (let r = header.rowIndex + 1; r < sheetData.length; r++) {
    const row = sheetData[r] || [];
    const desc = row[cols.desc];
    const unit = row[cols.unit];
    const qty = row[cols.qty];
    const note = cols.note >= 0 ? row[cols.note] : "";

    // Has a numeric quantity AND a unit → it's a data row
    const hasNumericQty = typeof qty === "number" && qty > 0;
    const hasUnit = unit != null && String(unit).trim() !== "";

    if (hasNumericQty && hasUnit && desc) {
      if (inExclusions) {
        // Numeric exclusion (unusual) — still capture as text
        result.exclusions.push(String(desc).trim());
        continue;
      }
      const substrate = inferSubstrate(String(desc), currentSection);
      const normalizedUnit = normalizeUnit(unit);
      const item = {
        type: "scope",
        name: String(desc).trim(),
        substrate,
        quantity: Number(qty),
        unit: normalizedUnit,
        coats: "2coats",
        notes: note ? String(note).trim() : (currentSection ? `Section: ${currentSection}` : ""),
        // Carry the source rate/amount for the chat-side flow to reference (not used in pricing here —
        // the tier multiplier is applied to catalog rates, not the spreadsheet's own rates)
        _sourceRate: typeof row[cols.rate] === "number" ? row[cols.rate] : null,
        _sourceAmount: typeof row[cols.amount] === "number" ? row[cols.amount] : null,
        _sourceSection: currentSection,
        _sourceSheet: sheetName,
      };
      result.items.push(item);
    } else if (desc && !hasNumericQty && isSectionHeader(desc)) {
      // It's a section header row — update context, OR detect Exclusions transition
      const cleanDesc = String(desc).trim();
      const lower = cleanDesc.toLowerCase();
      if (lower === "exclusions" || lower.startsWith("exclusion")) {
        inExclusions = true;
        currentSection = "";
      } else if (inExclusions) {
        // Inside the exclusions block — every text row is an exclusion line
        result.exclusions.push(cleanDesc);
      } else {
        // Regular section header (e.g. "Walls", "Ceiling", "Door and Frames")
        currentSection = cleanDesc;
      }
    }
    // Otherwise: blank row or skip row — ignore
  }

  return result;
}

async function parseXlsxFile(file) {
  const arrayBuffer = await file.arrayBuffer();
  const wb = XLSX.read(arrayBuffer, { type: "array" });

  // Find the best sheet: prefer one named "BoQ" (with optional trailing whitespace),
  // else use any sheet that has a valid header row.
  const candidates = wb.SheetNames.filter(n => /^boq\b/i.test(n.trim()));
  const sheetsToTry = candidates.length > 0 ? candidates : wb.SheetNames;

  const aggregate = { items: [], exclusions: [], warnings: [], sheetsParsed: [] };
  for (const sn of sheetsToTry) {
    const ws = wb.Sheets[sn];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const result = parseTakeoffSheet(data, sn);
    if (result.items.length > 0) {
      aggregate.items.push(...result.items);
      aggregate.sheetsParsed.push(sn);
    }
    aggregate.exclusions.push(...result.exclusions);
    aggregate.warnings.push(...result.warnings);
  }

  // Dedupe exclusions (case-insensitive)
  const seen = new Set();
  aggregate.exclusions = aggregate.exclusions.filter(e => {
    const k = e.toLowerCase();
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });

  return aggregate;
}

// ============================================================================
// COST GROUP DETECTION (from parsed rooms)
// ============================================================================

function detectCostGroups(items, library) {
  const groups = new Set();
  items.forEach((item) => {
    const t = item.type || "room";

    if (t === "room") {
      const hasWalls = item.walls?.enabled !== false;
      const hasCeiling = item.ceiling?.enabled !== false;
      const hasBaseboard = item.baseboard?.enabled !== false;
      const hasDoors = item.doors?.enabled !== false && (item.doors?.count || 0) > 0;
      if (hasWalls || hasCeiling) groups.add("drywall_walls_ceilings");
      if (hasBaseboard) groups.add("baseboard_trim");
      if (hasDoors) groups.add("doors_frames");
    } else if (t === "scope") {
      const sub = (item.substrate || "").toLowerCase();
      if (sub.includes("door")) groups.add("doors_frames");
      if (sub.includes("baseboard") || sub.includes("trim")) groups.add("baseboard_trim");
      if (sub.includes("cabinet")) groups.add("cabinets");
      if (sub.includes("stucco")) groups.add("exterior_stucco_siding");
      if (sub.includes("exterior_trim")) groups.add("exterior_wood_trim");
      if (sub.includes("fence") || sub.includes("deck")) groups.add("fence_deck_stain");
      if (sub === "walls" || sub === "ceilings") groups.add("drywall_walls_ceilings");
      if (sub === "metal") groups.add("baseboard_trim");
    } else if (t === "tm") {
      const cat = (item.category || "").toLowerCase();
      if (cat.includes("door") || cat.includes("window")) groups.add("doors_frames");
      if (cat.includes("trim")) groups.add("baseboard_trim");
      if (cat.includes("millwork")) groups.add("cabinets");
      if (cat.includes("exterior_walls")) groups.add("exterior_stucco_siding");
      if (cat.includes("exterior_misc")) groups.add("exterior_wood_trim");
      if (cat.includes("interior_walls")) groups.add("drywall_walls_ceilings");
    }

    // Substrate-based detection from notes/name (universal)
    const text = ((item.notes || "") + " " + (item.name || "")).toLowerCase();
    Object.entries(library.cost_groups).forEach(([key, group]) => {
      if (group.triggers.some((t) => text.includes(t))) groups.add(key);
    });
  });
  return [...groups];
}

// Map the cost-group NAMES that BuildChat created in JobTread onto scope-library
// cost_group keys. BuildChat's subgroup names are fixed by its system prompt
// (Stage C step 8a: "Drywall Walls", "Drywall Ceilings", "Wood Baseboard",
// "Doors+Frames"). Scope-bucket and T&M groups use other names; fall back to
// trigger-keyword matching against the library for those.
function detectGroupsFromBuiltBudget(builtBudget, library) {
  if (!builtBudget || !builtBudget.costGroupIds) return [];
  const keys = new Set();
  const names = Object.keys(builtBudget.costGroupIds);
  const NAME_MAP = [
    { match: /drywall (walls|ceilings)/i, key: "drywall_walls_ceilings" },
    { match: /wood baseboard|baseboard|trim|crown|molding/i, key: "baseboard_trim" },
    { match: /doors?\+?frames?|\bdoor\b|jamb|casing/i, key: "doors_frames" },
    { match: /cabinet|vanity|millwork/i, key: "cabinets" },
    { match: /stucco|siding/i, key: "exterior_stucco_siding" },
    { match: /fascia|eaves|soffit|exterior (wood|trim)/i, key: "exterior_wood_trim" },
    { match: /fence|deck|stain|seal/i, key: "fence_deck_stain" },
  ];
  names.forEach((n) => {
    // Skip pure structural parents that carry no prep meaning.
    if (/^(interior|exterior)$/i.test(n.trim())) return;
    let matched = false;
    for (const { match, key } of NAME_MAP) {
      if (match.test(n)) { keys.add(key); matched = true; }
    }
    // Fallback: trigger-keyword match against the library (covers room-name
    // parents and any custom group names).
    if (!matched) {
      Object.entries(library.cost_groups).forEach(([key, group]) => {
        if (group.triggers.some((t) => n.toLowerCase().includes(t))) keys.add(key);
      });
    }
  });
  return [...keys];
}

// ============================================================================
// MAIN APP
// ============================================================================

export default function App() {
  const [library, setLibrary] = useState(DEFAULT_LIBRARY);
  const [step, setStep] = useState(1);
  const [showSettings, setShowSettings] = useState(false);

  // Pricing tier (used by scope builder to pick prep language).
  // Customer/job identity (name, address, email, phone) is handled at the
  // chat level when the payload is pasted — NOT collected in the artifact.
  const [tier, setTier] = useState("standard");
  const [inferredJobInfo, setInferredJobInfo] = useState({}); // retained for tier hint only

  // Transcript & rooms
  const [inputMode, setInputMode] = useState("voice"); // "voice" | "notes" | "takeoff"
  const [transcript, setTranscript] = useState("");
  const [takeoffFile, setTakeoffFile] = useState(null);          // File object
  const [takeoffExclusions, setTakeoffExclusions] = useState([]); // strings pulled from sheet
  const [takeoffWarnings, setTakeoffWarnings] = useState([]);
  const [notesOverlayText, setNotesOverlayText] = useState(null); // text used for file+text overlay (debug)
  const [detection, setDetection] = useState(null);                // { mode, rationale } for text-only auto-detect
  const [rooms, setRooms] = useState([]);
  const [parsing, setParsing] = useState(false);
  const [parseError, setParseError] = useState("");

  // Clarifications from parser
  const [clarifications, setClarifications] = useState([]);
  const [inferredFlags, setInferredFlags] = useState({});
  const [answeredQuestions, setAnsweredQuestions] = useState({}); // { id: answer | null }
  const [clarifyRemaining, setClarifyRemaining] = useState(0); // ClarifyChat queue size — gates Build Scope

  // Context flags
  const [flags, setFlags] = useState({
    vacant: false,
    petsKids: false,
    leadPaint: false,
    activeBusiness: false,
  });

  // Scope selections
  const [scopeSelections, setScopeSelections] = useState({}); // { groupKey: { itemText: bool } }
  const [exclusionSelections, setExclusionSelections] = useState({}); // { itemText: bool }
  const [aiCategorization, setAiCategorization] = useState(null); // { groupKey: { recommended: [], suggested: [], reasons: {} } }
  const [categorizing, setCategorizing] = useState(false);
  const [customScope, setCustomScope] = useState({}); // { groupKey: [strings] }
  const [customExclusions, setCustomExclusions] = useState([]);
  const [buildStarted, setBuildStarted] = useState(false);          // Step 3 — gates the BuildChat panel
  const [builtBudget, setBuiltBudget] = useState(null);             // Snapshot from BuildChat.onBuilt — feeds Scope + DocumentChat
  const [budgetConfirmed, setBudgetConfirmed] = useState(false);    // "Budget looks good" gate clicked
  const [scopeGroups, setScopeGroups] = useState([]);               // Cost-group keys resolved at categorize time (budget → fallback to parsed)
  const [documentStarted, setDocumentStarted] = useState(false);    // Step 5 — gates the DocumentChat panel

  // Load library from storage on mount
  useEffect(() => {
    (async () => {
      const stored = await loadFromStorage(STORAGE_KEYS.library, DEFAULT_LIBRARY);
      setLibrary(stored);
    })();
  }, []);

  const tierMeta = PRICING_TIERS[tier];
  const detectedGroups = useMemo(() => detectCostGroups(rooms, library), [rooms, library]);

  // ============================================================================
  // CLAUDE API CALLS
  // ============================================================================

  async function parseTranscript(forceMode = null) {
    const hasFile = !!takeoffFile;
    const hasText = !!transcript.trim();

    if (!hasFile && !hasText) {
      setParseError("Upload a takeoff file or paste transcript/notes first.");
      return;
    }

    setParsing(true);
    setParseError("");

    try {
      // ── FILE PRESENT: takeoff path (with optional notes overlay) ──────
      // Whether or not there's accompanying text, a file means "takeoff" —
      // text is treated as an overlay that adjusts the parsed items.
      if (hasFile) {
        const result = await parseXlsxFile(takeoffFile);

        if (result.items.length === 0) {
          const msg = result.warnings.length
            ? "No line items found. " + result.warnings.join(" | ")
            : "No line items found in the spreadsheet. Check that it has columns: Description, Unit, Quantity.";
          setParseError(msg);
          setParsing(false);
          return;
        }

        let normalized = result.items.map(normalizeItem);
        let exclusions = [...(result.exclusions || [])];
        let overlayClarifications = [];
        let warnings = [...(result.warnings || [])];
        const overlayText = hasText ? transcript : null;

        if (hasText) {
          try {
            const overlay = await applyNotesOverlay(normalized, transcript);
            (overlay.overrides || []).forEach(({ itemIndex, patch }) => {
              if (Number.isInteger(itemIndex) && normalized[itemIndex] && patch) {
                normalized[itemIndex] = applyPatchToItem(normalized[itemIndex], patch);
              }
            });
            const additions = (overlay.additions || []).map(normalizeItem);
            normalized = normalized.concat(additions);
            exclusions = exclusions.concat(overlay.exclusions || []);
            overlayClarifications = overlay.questions || [];
          } catch (overlayErr) {
            // Overlay failure is non-fatal — surface as a warning, keep the takeoff
            warnings = warnings.concat([`Notes overlay failed: ${overlayErr.message}`]);
          }
        }

        setInputMode("takeoff");
        setDetection(null);                       // no banner — file path is unambiguous
        setNotesOverlayText(overlayText);
        setRooms(normalized);
        setClarifications(overlayClarifications);
        setTakeoffExclusions(exclusions);
        setTakeoffWarnings(warnings);
        setInferredJobInfo({});
        setInferredFlags({});
        setStep(2);
        return;
      }

      // ── TEXT ONLY: classifier decides voice vs notes ──────────────────
      let resolvedMode = forceMode;
      let rationale = forceMode ? `Manual override → ${forceMode}` : "";
      if (!resolvedMode) {
        try {
          const classified = await classifyInput(transcript);
          resolvedMode = classified.mode === "voice" ? "voice" : "notes";
          rationale = classified.rationale || "classifier verdict";
        } catch (clsErr) {
          const h = classifyHeuristic(transcript);
          resolvedMode = h.mode;
          rationale = `classifier unavailable — ${h.rationale}`;
        }
      }
      setInputMode(resolvedMode);
      setDetection({ mode: resolvedMode, rationale });
      setNotesOverlayText(null);

      // Notes mode: SINGLE combined call (jobInfo + items in one response on Haiku).
      // This is 2x faster than two parallel calls because we save the round-trip overhead
      // and the model only has to context-switch once. Token budget is generous enough
      // to handle a 10-room job with 3 clarifications.
      if (resolvedMode === "notes") {
        const combined = await callClaude({
          model: "claude-haiku-4-5-20251001",
          maxTokens: 8000,
          timeoutMs: 60000,
          prompt: buildCombinedNotesPrompt(transcript),
          label: "notes parse",
        });

        const jobInfo = combined.jobInfo || {};
        const inferredFlagsResult = combined.flags || {};
        setInferredJobInfo(jobInfo);
        setInferredFlags(inferredFlagsResult);

        if (jobInfo.tier?.value && PRICING_TIERS[jobInfo.tier.value]) setTier(jobInfo.tier.value);

        const newFlags = { ...flags };
        Object.entries(inferredFlagsResult).forEach(([k, v]) => {
          if (v && typeof v === "object" && (v.confidence === "high" || v.confidence === "medium")) {
            newFlags[k] = !!v.value;
          }
        });
        setFlags(newFlags);

        const items = combined.items || [];
        const questions = combined.questions || [];
        const normalized = items.map(normalizeItem);
        setRooms(normalized);
        setClarifications(questions);

        setTakeoffExclusions([]);
        setTakeoffWarnings(combined._truncated
          ? [`Response was truncated and recovered. ${normalized.length} items + ${questions.length} questions parsed — but some items at the end may be missing. Add any missing items manually on the Clarify + Items screen.`]
          : []);

        setStep(2);
        return;
      }

      // Voice mode: two parallel calls (Haiku for job info, Sonnet for items).
      // Kept separate because voice transcripts benefit from Sonnet's deeper parsing
      // and the parallel calls overlap to reduce wall time.
      const [jobInfoResult, itemsResult] = await Promise.all([
        callClaude({
          model: "claude-haiku-4-5-20251001",
          maxTokens: 1000,
          timeoutMs: 30000,
          prompt: buildJobInfoPrompt(transcript),
          label: "job info",
        }),
        callClaude({
          model: "claude-sonnet-4-6",
          maxTokens: 6000,
          timeoutMs: 90000,
          prompt: buildItemsPrompt(transcript),
          label: "items",
        }),
      ]);

      const jobInfo = jobInfoResult.jobInfo || {};
      const inferredFlagsResult = jobInfoResult.flags || {};
      setInferredJobInfo(jobInfo);
      setInferredFlags(inferredFlagsResult);

      if (jobInfo.tier?.value && PRICING_TIERS[jobInfo.tier.value]) setTier(jobInfo.tier.value);

      const newFlags = { ...flags };
      Object.entries(inferredFlagsResult).forEach(([k, v]) => {
        if (v && typeof v === "object" && (v.confidence === "high" || v.confidence === "medium")) {
          newFlags[k] = !!v.value;
        }
      });
      setFlags(newFlags);

      const items = Array.isArray(itemsResult) ? itemsResult : (itemsResult.items || []);
      const questions = Array.isArray(itemsResult) ? [] : (itemsResult.questions || []);
      const normalized = items.map(normalizeItem);
      setRooms(normalized);
      setClarifications(questions);

      // Clear any takeoff-mode-only state
      setTakeoffExclusions([]);
      setTakeoffWarnings([]);

      setStep(2);
    } catch (e) {
      setParseError("Parse failed: " + e.message);
    } finally {
      setParsing(false);
    }
  }

  async function categorizeScopeItems() {
    setCategorizing(true);
    try {
      // Re-source Scope from the BUILT budget (in-memory snapshot from BuildChat).
      // Fall back to parsed-room detection if the snapshot is somehow empty.
      const budgetGroups = detectGroupsFromBuiltBudget(builtBudget, library);
      const groupsForScope = budgetGroups.length ? budgetGroups : detectedGroups;
      setScopeGroups(groupsForScope);

      const context = {
        tier,
        prepTier: tierMeta.prepTier,
        rooms,
        flags,
        detectedGroups: groupsForScope,
      };

      const groupsForApi = {};
      groupsForScope.forEach((key) => {
        const grp = library.cost_groups[key];
        if (grp) groupsForApi[key] = grp[tierMeta.prepTier] || [];
      });

      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 3000,
          messages: [{
            role: "user",
            content: `You are categorizing prep items for a painting proposal. Given the job context and a library of prep items per cost group, classify each item into one of three buckets:

- "recommended" — high confidence this applies to this specific job
- "suggested" — plausible but not default; needs estimator judgment
- (everything else stays in the Library bucket, not in your response)

Job context:
${JSON.stringify(context, null, 2)}

Library items per cost group (all at "${tierMeta.prepTier}" tier):
${JSON.stringify(groupsForApi, null, 2)}

Also classify these blanket exclusions:
${JSON.stringify(library.blanket_exclusions, null, 2)}

Universal prep items at "${tierMeta.prepTier}" tier:
${JSON.stringify(library.universal_prep[tierMeta.prepTier], null, 2)}

Return ONLY a JSON object with this shape (no preamble or markdown fences):
{
  "universal": {
    "recommended": ["item text", ...],
    "suggested": ["item text", ...],
    "reasons": { "item text": "short reason why suggested" }
  },
  "groups": {
    "<groupKey>": {
      "recommended": ["item text", ...],
      "suggested": ["item text", ...],
      "reasons": { "item text": "short reason why suggested" }
    }
  },
  "exclusions": {
    "recommended": ["item text", ...],
    "suggested": ["item text", ...],
    "reasons": {}
  }
}

Reasoning hints:
- If flags.vacant is true, "Cover and protect furniture" is NOT recommended (move to suggested or skip)
- If flags.leadPaint is true, recommend lead-safe practices
- If flags.activeBusiness is true, recommend extra care/cleanup language
- High-end tier should recommend more thorough prep
- Doors+Frames: "Remove hardware and label" vs "Mask hardware in place" are mutually exclusive — pick one for recommended based on tier (high-end = remove, standard = mask)`,
          }],
        }),
      });
      const data = await response.json();
      const text = data.content.find((c) => c.type === "text")?.text || "";
      const cleaned = text.replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(cleaned);
      setAiCategorization(parsed);

      // Pre-check the recommended items
      const newScopeSel = {};
      // Universal
      newScopeSel["_universal"] = {};
      (parsed.universal?.recommended || []).forEach((item) => {
        newScopeSel["_universal"][item] = true;
      });
      // Each group
      groupsForScope.forEach((gk) => {
        newScopeSel[gk] = {};
        (parsed.groups?.[gk]?.recommended || []).forEach((item) => {
          newScopeSel[gk][item] = true;
        });
      });
      setScopeSelections(newScopeSel);

      const newExclSel = {};
      (parsed.exclusions?.recommended || []).forEach((item) => {
        newExclSel[item] = true;
      });
      setExclusionSelections(newExclSel);

      setStep(4);
    } catch (e) {
      alert("Categorization failed: " + e.message);
    } finally {
      setCategorizing(false);
    }
  }

  // ============================================================================
  // ROOM EDITING
  // ============================================================================

  function updateRoom(idx, field, value) {
    const next = [...rooms];
    next[idx] = { ...next[idx], [field]: value };
    setRooms(next);
  }

  function deleteRoom(idx) {
    setRooms(rooms.filter((_, i) => i !== idx));
  }

  function addRoom() {
    setRooms([...rooms, normalizeItem({ type: "room", name: "New Room", length: 10, width: 10, height: 9 })]);
  }

  function addScope() {
    setRooms([...rooms, { type: "scope", name: "New Scope", substrate: "doors", quantity: 1, unit: "EA", coats: "2coats", notes: "" }]);
  }

  function addTM() {
    setRooms([...rooms, { type: "tm", name: "T&M Touch-ups", hours: 4, category: "repairs", notes: "" }]);
  }

  // ============================================================================
  // SCOPE CHECKBOX HANDLING
  // ============================================================================

  function toggleScopeItem(groupKey, itemText) {
    setScopeSelections((prev) => ({
      ...prev,
      [groupKey]: { ...(prev[groupKey] || {}), [itemText]: !prev[groupKey]?.[itemText] },
    }));
  }

  function toggleExclusion(itemText) {
    setExclusionSelections((prev) => ({ ...prev, [itemText]: !prev[itemText] }));
  }

  function addCustomScope(groupKey, text) {
    if (!text.trim()) return;
    setCustomScope((prev) => ({
      ...prev,
      [groupKey]: [...(prev[groupKey] || []), text.trim()],
    }));
    setScopeSelections((prev) => ({
      ...prev,
      [groupKey]: { ...(prev[groupKey] || {}), [text.trim()]: true },
    }));
  }

  function addCustomExclusion(text) {
    if (!text.trim()) return;
    setCustomExclusions((prev) => [...prev, text.trim()]);
    setExclusionSelections((prev) => ({ ...prev, [text.trim()]: true }));
  }

  // ============================================================================
  // EXPORT PAYLOAD
  // ============================================================================

  function buildPayload() {
    const selectedScope = {};
    Object.entries(scopeSelections).forEach(([gk, items]) => {
      const picked = Object.entries(items).filter(([, v]) => v).map(([k]) => k);
      if (picked.length) selectedScope[gk] = picked;
    });
    const selectedExclusions = Object.entries(exclusionSelections).filter(([, v]) => v).map(([k]) => k);

    // Split line items by type for clarity in the payload
    const physicalRooms = rooms.filter((r) => (r.type || "room") === "room");
    const scopeBuckets = rooms.filter((r) => r.type === "scope");
    const tmItems = rooms.filter((r) => r.type === "tm");

    // Map each T&M item to its catalog ID
    const tmWithCatalog = tmItems.map((tm) => {
      const cat = (tm.category || "").toLowerCase();
      const match = TM_CATALOG.find((c) =>
        c.codeName.toLowerCase().includes(cat.replace(/_/g, " ")) ||
        c.name.toLowerCase().includes(cat.replace(/_/g, " "))
      );
      return {
        ...tm,
        catalogId: match?.id,
        catalogName: match?.name,
        costCode: match?.code,
        costCodeName: match?.codeName,
        unitCost: match?.unitCost,
        unitPrice: match?.unitPrice,
      };
    });

    // bidOrigin maps the artifact's inputMode to JobTread's "How Did Bid Come In"
    // values. Chat-level Claude uses this to set custom field 22PWsDkAPRYB.
    const bidOrigin =
      inputMode === "takeoff" ? "Digital Takeoff" : "Job Walk";

    return {
      // NO `job` block — customer name / address / email / phone are collected
      // by chat-level Claude when the payload is pasted. The artifact only
      // produces the scope of work, not the job identity.
      tier: { key: tier, label: tierMeta.label, multiplier: tierMeta.multiplier },
      inputMode,                             // "voice" | "notes" | "takeoff"
      bidOrigin,                             // "Job Walk" | "Digital Takeoff"
      flags,
      rooms: physicalRooms,
      scopeBuckets,
      tmItems: tmWithCatalog,
      scope: selectedScope,
      exclusions: selectedExclusions,
      catalog: CATALOG_IDS,
      tmCatalog: TM_CATALOG,
      ...(inputMode === "takeoff" && {
        takeoffMeta: {
          sourceFile: takeoffFile?.name || null,
          parsedExclusions: takeoffExclusions,
          warnings: takeoffWarnings,
          notesOverlay: notesOverlayText,
        },
      }),
    };
  }

  // Manual-copy modal state — shown when clipboard APIs fail in the sandboxed iframe
  const [manualCopyText, setManualCopyText] = useState(null);
  // Last copy result — for inline "Copied ✓" feedback on buttons
  const [lastCopyResult, setLastCopyResult] = useState(null); // { label, method, at }

  // Robust clipboard write — tries modern API, then execCommand fallback, then
  // shows a textarea modal as a last resort. Artifacts iframes often have
  // restricted clipboard access, so we can't trust navigator.clipboard alone.
  async function copyToClipboard(text, label = "Payload") {
    // Path 1: modern Clipboard API
    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(text);
        setLastCopyResult({ label, method: "clipboard-api", at: Date.now() });
        return { ok: true, method: "clipboard-api" };
      }
    } catch (_) { /* fall through to legacy */ }

    // Path 2: legacy execCommand("copy") via hidden textarea
    try {
      const ta = document.createElement("textarea");
      ta.value = text;
      ta.style.position = "fixed";
      ta.style.top = "-9999px";
      ta.style.opacity = "0";
      ta.setAttribute("readonly", "");
      document.body.appendChild(ta);
      ta.select();
      ta.setSelectionRange(0, text.length);
      const ok = document.execCommand("copy");
      document.body.removeChild(ta);
      if (ok) {
        setLastCopyResult({ label, method: "execCommand", at: Date.now() });
        return { ok: true, method: "execCommand" };
      }
    } catch (_) { /* fall through */ }

    // Path 3: show modal with the text so the user can copy manually
    setManualCopyText({ text, label });
    return { ok: false, method: "manual" };
  }

  function copyPayload() {
    const payload = buildPayload();
    const text = `BUILD THIS JOB IN JOBTREAD:

\`\`\`json
${JSON.stringify(payload, null, 2)}
\`\`\`

Before building the cost groups, ASK THE USER for the following (one batch, using ask_user_input_v0):

1. **Customer / Job** — Which customer and job does this apply to? Search JobTread for the customer first. If new, gather: customer name, primary contact name, email, phone, address (job site). If existing, confirm the job (or create a new one under that customer).

2. **Pricing tier** — The artifact inferred \`${payload.tier.label}\` (${payload.tier.multiplier}x). Confirm or override. Options: Standard (1.0x), Production (0.85x), High-End (1.35x), Prevailing Wage (1.65x). This multiplier gets applied to BOTH unitCost AND unitPrice on every cost item (so margin % stays intact).

3. **Estimate preparer** — Who prepared this estimate? (Peter, Lyric, Kareem, etc.) Needed for the document.

4. **Recipient** — Default to job primary contact. Confirm before generating the customer-facing proposal.

Then, when creating the JobTread job, set these custom fields:
- **Job Type** (22PWsEVVW4aj): Based on tier. Standard→"Residential - Standard Home"; Production→"Property Management / Production"; High-End→"Residential - Custom House"; Prevailing Wage→"Prevailing Wage". Confirm with user if ambiguous (e.g. commercial vs residential).
- **How Did Bid Come In** (22PWsDkAPRYB): \`${payload.bidOrigin}\` (artifact-determined from inputMode = "${payload.inputMode}").

Once the above is confirmed, use the JobTread MCP to:
1. Find or create the customer and job
2. Check for an existing document template; use it if found
3. For each \`rooms[]\` entry: build a parent cost group (room name) with subgroups for ceiling/walls/baseboard/doors, applying the tier multiplier to unitCost AND unitPrice
4. For each \`scopeBuckets[]\` entry: build a single cost group named by substrate, with the quantity and unit specified, using catalog items at the requested coats
5. For each \`tmItems[]\` entry: use the \`catalogId\` (an existingCostItem reference) to add a Time & Materials line under the matching costCode, with \`hours\` as quantity (Unit = HR)
6. Push selected \`scope\` items to document description (one section per group); push \`exclusions\` to document footer`;
    copyToClipboard(text, "Full Payload");
  }

  function copyScopeText() {
    const lines = [];
    Object.entries(scopeSelections).forEach(([gk, items]) => {
      const picked = Object.entries(items).filter(([, v]) => v).map(([k]) => k);
      if (picked.length) {
        const label = gk === "_universal" ? "General" : (library.cost_groups[gk]?.label || gk);
        lines.push(`\n${label}:`);
        picked.forEach((p) => lines.push(`• ${p}`));
      }
    });
    copyToClipboard(lines.join("\n").trim(), "Scope Text");
  }

  function copyExclusionText() {
    const picked = Object.entries(exclusionSelections).filter(([, v]) => v).map(([k]) => `• ${k}`);
    copyToClipboard(picked.join("\n"), "Exclusions");
  }

  // ============================================================================
  // RENDER
  // ============================================================================

  const wrap = { fontFamily: "var(--font-sans, system-ui)", maxWidth: 980, margin: "0 auto", padding: "1rem 1.5rem", color: "var(--color-text-primary, #222)" };
  const card = { background: "var(--color-background-primary, #fff)", border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", borderRadius: 12, padding: "1rem 1.25rem", marginBottom: 16 };
  const btn = { padding: "8px 14px", borderRadius: 8, border: "0.5px solid var(--color-border-secondary, rgba(0,0,0,0.3))", background: "transparent", cursor: "pointer", fontSize: 14, fontFamily: "inherit", color: "inherit" };
  const btnPrimary = { ...btn, background: "#2C1654", color: "#fff", border: "0.5px solid #2C1654" };

  return (
    <div style={wrap}>
      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 24 }}>
        <div>
          <h1 style={{ fontSize: 22, fontWeight: 500, margin: 0, color: "#2C1654" }}>Job Walk → Proposal</h1>
          <p style={{ fontSize: 13, color: "var(--color-text-secondary, #666)", margin: "4px 0 0" }}>Purple Painting Co. — combined estimating tool</p>
        </div>
        <button style={btn} onClick={() => setShowSettings(!showSettings)}>
          {showSettings ? "← Back" : "Settings"}
        </button>
      </div>

      {showSettings ? (
        <SettingsPanel library={library} setLibrary={setLibrary} />
      ) : (
        <>
          {/* Progress steps */}
          <div style={{ display: "flex", gap: 8, marginBottom: 24, flexWrap: "wrap" }}>
            {(() => {
              // Build the visible step list per mode. Each entry: { realStep, label }
              // Step 2 is the combined Clarify + Line Items screen (single view).
              // In takeoff mode there's nothing to clarify, so it shows as "Confirm".
              const fullSteps = [
                { realStep: 1, label: "Input" },
                { realStep: 2, label: inputMode === "takeoff" ? "Confirm" : "Clarify + Items" },
                { realStep: 3, label: "Build" },
                { realStep: 4, label: "Scope", hideIn: ["takeoff"] },
                { realStep: 5, label: "Export" },
              ].filter((s) => !s.hideIn || !s.hideIn.includes(inputMode));

              return fullSteps.map((s, visibleIdx) => {
                const isActive = step === s.realStep;
                const isPast = step > s.realStep;
                return (
                  <div key={s.realStep} style={{
                    padding: "6px 12px", borderRadius: 16, fontSize: 12,
                    background: isActive ? "#2C1654" : isPast ? "#C8963E" : "var(--color-background-secondary, #f3f3f0)",
                    color: isActive || isPast ? "#fff" : "var(--color-text-secondary, #666)",
                    cursor: "pointer",
                  }} onClick={() => setStep(s.realStep)}>
                    {visibleIdx + 1}. {s.label}
                  </div>
                );
              });
            })()}
          </div>

          {/* Step 1: Input (auto-detect — no mode picker)
              ─────────────────────────────────────────────
              File only        → takeoff (deterministic XLSX parse)
              File + text      → takeoff + notes overlay (Haiku adjusts items)
              Text only        → classifier decides voice vs notes
              Nothing          → error
          */}
          {step === 1 && (
            <div style={card}>
              <h2 style={{ fontSize: 18, fontWeight: 500, margin: "0 0 12px" }}>Input</h2>

              <p style={{ fontSize: 13, color: "var(--color-text-secondary, #666)", marginBottom: 14 }}>
                Upload a digital takeoff, paste a job-walk transcript or estimator notes, or both. Mode is auto-detected — you can override after parsing.
              </p>

              <div style={{ display: "flex", gap: 12, alignItems: "stretch", flexWrap: "wrap", marginBottom: 8 }}>
                {/* File uploader */}
                <label
                  style={{
                    flex: "1 1 280px",
                    display: "flex",
                    flexDirection: "column",
                    alignItems: "center",
                    justifyContent: "center",
                    minHeight: 220,
                    padding: "24px 16px",
                    border: `1.5px dashed ${takeoffFile ? "#2C1654" : "var(--color-border-secondary, rgba(0,0,0,0.3))"}`,
                    borderRadius: 10,
                    cursor: "pointer",
                    background: takeoffFile ? "rgba(44, 22, 84, 0.04)" : "transparent",
                  }}
                >
                  <input
                    type="file"
                    accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    onChange={(e) => {
                      const f = e.target.files?.[0];
                      if (f) {
                        setTakeoffFile(f);
                        setParseError("");
                      }
                    }}
                    style={{ display: "none" }}
                  />
                  <div style={{ fontSize: 24, marginBottom: 6 }}>📐</div>
                  <div style={{ fontSize: 14, fontWeight: 500, marginBottom: 4, textAlign: "center" }}>
                    {takeoffFile ? takeoffFile.name : "Click to upload takeoff .xlsx"}
                  </div>
                  <div style={{ fontSize: 12, color: "var(--color-text-secondary, #666)", textAlign: "center" }}>
                    {takeoffFile
                      ? `${(takeoffFile.size / 1024).toFixed(1)} KB — click to replace`
                      : "Optional. Purple Painting BoQ format."}
                  </div>
                  {takeoffFile && (
                    <button
                      type="button"
                      onClick={(e) => { e.preventDefault(); e.stopPropagation(); setTakeoffFile(null); }}
                      style={{ marginTop: 10, fontSize: 11, background: "transparent", border: "0.5px solid rgba(0,0,0,0.3)", borderRadius: 4, padding: "2px 8px", cursor: "pointer" }}
                    >
                      Clear file
                    </button>
                  )}
                </label>

                {/* Text input */}
                <textarea
                  value={transcript}
                  onChange={(e) => setTranscript(e.target.value)}
                  rows={10}
                  style={{ flex: "1 1 320px", minHeight: 220, padding: 12, borderRadius: 8, border: "0.5px solid var(--color-border-secondary, rgba(0,0,0,0.3))", fontSize: 13, fontFamily: "var(--font-mono, monospace)", background: "transparent", color: "inherit", resize: "vertical" }}
                  placeholder={
                    takeoffFile
                      ? "Optional scope notes that adjust this takeoff (e.g. '1 coat for doors throughout, exclude basement')."
                      : "Paste a job-walk transcript or estimator notes."
                  }
                />
              </div>

              {parseError && <div style={{ color: "var(--color-text-danger, #c33)", fontSize: 13, marginTop: 8 }}>{parseError}</div>}

              <div style={{ display: "flex", gap: 8, marginTop: 16, alignItems: "center" }}>
                <button
                  style={btnPrimary}
                  onClick={() => parseTranscript()}
                  disabled={parsing || (!takeoffFile && !transcript.trim())}
                >
                  {parsing
                    ? (takeoffFile
                        ? (transcript.trim() ? "Parsing takeoff + overlay…" : "Reading spreadsheet…")
                        : "Parsing… (10–20 sec)")
                    : "Parse →"}
                </button>
                {parsing && !takeoffFile && (
                  <div style={{ fontSize: 12, color: "var(--color-text-secondary, #666)", fontStyle: "italic" }}>
                    Auto-detecting voice vs notes, then parsing.
                  </div>
                )}
              </div>
            </div>
          )}

          {/* Step 2: Combined Clarify + Line Items
              ─────────────────────────────────────────────
              Clarifications + tier + context flags live ABOVE the line items list.
              As clarifications are answered, the line items below update in real
              time — same `rooms` state powers both, so the wiring is automatic.

              Customer name, address, email, phone are NOT collected here. Those
              are handled at the chat level when the payload is pasted.
          */}
          {step === 2 && (
            <div>
              {/* Auto-detect banner — only shown for the text-only path */}
              {detection && (
                <div style={{ ...card, background: "rgba(44, 22, 84, 0.05)", border: "0.5px solid #2C1654", marginBottom: 12 }}>
                  <div style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
                    <div style={{ fontSize: 13, color: "#2C1654", flex: "1 1 auto" }}>
                      <strong>Detected:</strong> {detection.mode === "voice" ? "Voice transcript" : "Estimator notes"} — <span style={{ opacity: 0.8 }}>{detection.rationale}</span>
                    </div>
                    <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                      <span style={{ fontSize: 12, color: "var(--color-text-secondary, #666)" }}>Wrong? Re-parse as:</span>
                      <button
                        style={{ ...btn, fontSize: 12, padding: "4px 10px" }}
                        disabled={parsing || detection.mode === "voice"}
                        onClick={() => parseTranscript("voice")}
                      >
                        Voice
                      </button>
                      <button
                        style={{ ...btn, fontSize: 12, padding: "4px 10px" }}
                        disabled={parsing || detection.mode === "notes"}
                        onClick={() => parseTranscript("notes")}
                      >
                        Notes
                      </button>
                    </div>
                  </div>
                </div>
              )}

              {/* Truncation warning (Notes mode salvage) */}
              {inputMode !== "takeoff" && takeoffWarnings.length > 0 && (
                <div style={{ ...card, background: "rgba(199, 150, 62, 0.08)", border: "0.5px solid #C8963E" }}>
                  <div style={{ fontSize: 13, color: "#a70", display: "flex", gap: 8, alignItems: "flex-start" }}>
                    <span style={{ fontSize: 16 }}>⚠️</span>
                    <div>{takeoffWarnings.join(" · ")}</div>
                  </div>
                </div>
              )}

              {/* ── CLARIFICATIONS PANEL (voice / notes only) ────────────────── */}
              {inputMode !== "takeoff" && (
                <div style={card}>
                  <h2 style={{ fontSize: 18, fontWeight: 500, margin: "0 0 8px" }}>Clarify</h2>
                  <p style={{ fontSize: 13, color: "var(--color-text-secondary, #666)", marginBottom: 16 }}>
                    Chat with the estimator to refine the items below — change coats, dimensions, add or remove rooms, anything.
                    {clarifications.length > 0 && ` ${clarifications.length} item${clarifications.length === 1 ? "" : "s"} flagged by the parser; suggestions appear in the chat.`}
                  </p>

                  {/* Pricing tier — needed in-artifact only because the Scope step
                      uses it to pick standard vs high-end prep wording. Chat-level
                      Claude will re-confirm the multiplier when applying pricing. */}
                  <div style={{ marginBottom: 16 }}>
                    <label style={{ display: "block", fontSize: 13, fontWeight: 500, marginBottom: 6 }}>
                      Pricing tier
                      {inferredJobInfo.tier?.confidence && inferredJobInfo.tier.confidence !== "unknown" && (
                        <span style={{ fontSize: 11, color: "var(--color-text-info, #46c)", fontStyle: "italic", marginLeft: 8, fontWeight: 400 }}>
                          (AI inferred: {PRICING_TIERS[inferredJobInfo.tier.value]?.label || inferredJobInfo.tier.value}, {inferredJobInfo.tier.confidence} confidence)
                        </span>
                      )}
                    </label>
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 8 }}>
                      {Object.entries(PRICING_TIERS).map(([k, t]) => (
                        <button key={k} style={{
                          ...btn,
                          padding: "8px 10px",
                          background: tier === k ? t.color : "transparent",
                          color: tier === k ? "#fff" : "inherit",
                          border: `0.5px solid ${tier === k ? t.color : "var(--color-border-secondary, rgba(0,0,0,0.3))"}`,
                        }} onClick={() => setTier(k)}>
                          <div style={{ fontWeight: 500 }}>{t.label}</div>
                          <div style={{ fontSize: 11, opacity: 0.8 }}>{t.multiplier}x</div>
                        </button>
                      ))}
                    </div>
                  </div>

                  {/* Open chat — always available, even with zero parser questions */}
                  <div style={{ marginBottom: 16 }}>
                    <ClarifyChat
                      initialHints={clarifications}
                      rooms={rooms}
                      library={library}
                      onQueueChange={setClarifyRemaining}
                      applyToItem={(itemIndex, updates) => {
                        if (itemIndex == null) return;
                        const next = [...rooms];
                        next[itemIndex] = applyPatchToItem(next[itemIndex], updates);
                        setRooms(next);
                      }}
                      applyMultiPatch={(directives) => {
                        // Each directive has ONE of:
                        //   - _add: <newItem>                              → append new item
                        //   - _remove: true, _itemIndex|itemNameMatch     → delete matching items
                        //   - _itemIndex: <n>, patch: {...}                → patch by exact index
                        //   - itemNameMatch: "kitchen", patch: {...}      → patch by name substring
                        //   - substrateMatch: "doors", patch: {...}       → patch substrate across all rooms
                        //   - both name & substrate                        → substrate sub-target within matched item
                        // Returns count of items added/removed/changed.
                        let next = [...rooms];
                        const changedIndexes = new Set();
                        let addedCount = 0;
                        let removedCount = 0;

                        for (const entry of directives) {
                          // ── Add a new item
                          if (entry._add && typeof entry._add === "object") {
                            next.push(normalizeItem(entry._add));
                            addedCount += 1;
                            continue;
                          }

                          // ── Remove items
                          if (entry._remove) {
                            if (entry._itemIndex != null && next[entry._itemIndex]) {
                              next = next.filter((_, idx) => idx !== entry._itemIndex);
                              removedCount += 1;
                              continue;
                            }
                            const nameMatch = (entry.itemNameMatch || "").toLowerCase();
                            if (nameMatch) {
                              const before = next.length;
                              next = next.filter((item) => !((item.name || "").toLowerCase().includes(nameMatch)));
                              removedCount += before - next.length;
                            }
                            continue;
                          }

                          // ── Patch by exact index
                          if (entry._itemIndex != null && next[entry._itemIndex]) {
                            next[entry._itemIndex] = applyPatchToItem(next[entry._itemIndex], entry.patch);
                            changedIndexes.add(entry._itemIndex);
                            continue;
                          }

                          // ── Patch by name/substrate selectors
                          const nameMatcher = (entry.itemNameMatch || "").toLowerCase();
                          const substrateMatch = (entry.substrateMatch || "").toLowerCase();
                          if (!nameMatcher && !substrateMatch) continue;

                          next.forEach((item, idx) => {
                            const itemName = (item.name || "").toLowerCase();
                            const itemType = item.type || "room";
                            const nameOk = !nameMatcher || itemName.includes(nameMatcher);
                            if (!nameOk) return;

                            if (substrateMatch) {
                              if (itemType === "room") {
                                const roomSubKey =
                                  substrateMatch === "doors"     ? "doors" :
                                  substrateMatch === "walls"     ? "walls" :
                                  substrateMatch === "ceilings" || substrateMatch === "ceiling" ? "ceiling" :
                                  substrateMatch === "baseboard" || substrateMatch === "trim"   ? "baseboard" :
                                  null;
                                if (!roomSubKey) return;
                                const sub = item[roomSubKey];
                                if (!sub || sub.enabled === false) return;
                                next[idx] = applyPatchToItem(next[idx], { [roomSubKey]: entry.patch });
                                changedIndexes.add(idx);
                              } else if (itemType === "scope") {
                                const scopeSub = (item.substrate || "").toLowerCase();
                                if (scopeSub === substrateMatch ||
                                    (substrateMatch === "trim" && scopeSub === "baseboard")) {
                                  next[idx] = applyPatchToItem(next[idx], entry.patch);
                                  changedIndexes.add(idx);
                                }
                              }
                            } else {
                              next[idx] = applyPatchToItem(next[idx], entry.patch);
                              changedIndexes.add(idx);
                            }
                          });
                        }
                        setRooms(next);
                        return changedIndexes.size + addedCount + removedCount;
                      }}
                    />
                  </div>

                  {/* Context flags */}
                  <div style={{ borderTop: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", paddingTop: 14 }}>
                    <h3 style={{ fontSize: 13, fontWeight: 500, margin: "0 0 8px" }}>Context Flags</h3>
                    <p style={{ fontSize: 11, color: "var(--color-text-secondary, #666)", marginBottom: 8 }}>These help recommend the right prep wording.</p>
                    {[
                      ["vacant", "Vacant property (no furniture to cover)"],
                      ["petsKids", "Children or pets on site"],
                      ["leadPaint", "Lead paint present (pre-1978 home)"],
                      ["activeBusiness", "Active business / occupied commercial"],
                    ].map(([key, label]) => {
                      const inferred = inferredFlags[key];
                      const showHint = inferred && inferred.confidence && inferred.confidence !== "unknown";
                      return (
                        <label key={key} style={{ display: "flex", alignItems: "center", gap: 10, padding: "4px 0", cursor: "pointer", fontSize: 12 }}>
                          <input type="checkbox" checked={flags[key]} onChange={(e) => setFlags({ ...flags, [key]: e.target.checked })} style={{ width: 14, height: 14 }} />
                          <span>{label}</span>
                          {showHint && (
                            <span style={{ fontSize: 10, color: "var(--color-text-info, #46c)", fontStyle: "italic" }}>
                              (AI inferred: {inferred.value ? "yes" : "no"}, {inferred.confidence} confidence)
                            </span>
                          )}
                        </label>
                      );
                    })}
                  </div>
                </div>
              )}

              {/* ── LINE ITEMS (live-updated by clarifications above) ────────── */}
              {inputMode === "takeoff" ? (
                // Takeoff mode: read-only confirmation table
                <div style={card}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
                    <h2 style={{ fontSize: 18, fontWeight: 500, margin: 0 }}>Confirm Takeoff Items ({rooms.length})</h2>
                    {takeoffFile && (
                      <div style={{ fontSize: 12, color: "var(--color-text-secondary, #666)" }}>
                        Source: {takeoffFile.name}
                      </div>
                    )}
                  </div>
                  <p style={{ fontSize: 12, color: "var(--color-text-secondary, #666)", margin: "0 0 12px" }}>
                    Parsed directly from your spreadsheet. Review and click Export when ready.
                    {takeoffWarnings.length > 0 && (
                      <span style={{ color: "var(--color-text-warning, #a70)", display: "block", marginTop: 4 }}>
                        {takeoffWarnings.length} warning{takeoffWarnings.length === 1 ? "" : "s"}: {takeoffWarnings.join(" · ")}
                      </span>
                    )}
                  </p>

                  <TakeoffConfirmTable items={rooms} exclusions={takeoffExclusions} />
                </div>
              ) : (
                // Voice / Notes mode: editable groups (live-updates from clarifications above)
                <GroupedLineItemsEditor
                  rooms={rooms}
                  updateRoom={updateRoom}
                  deleteRoom={deleteRoom}
                  addRoom={addRoom}
                  addScope={addScope}
                  addTM={addTM}
                />
              )}

              <div style={{ display: "flex", gap: 8 }}>
                <button style={btn} onClick={() => setStep(1)}>← Back to Input</button>
                {inputMode === "takeoff" ? (
                  <button
                    style={btnPrimary}
                    onClick={() => {
                      // Skip AI categorization entirely — pre-populate scope + exclusions
                      // with the takeoff defaults. Build comes next; takeoff's Scope step
                      // is hidden, so the confirm gate (step 3) will jump straight to Export.
                      const universalSel = { "Per Plans and Specifications": true };
                      setScopeSelections({ _universal: universalSel });
                      const exclSel = {};
                      takeoffExclusions.forEach((e) => { exclSel[e] = true; });
                      setExclusionSelections(exclSel);
                      setCustomScope({ _universal: ["Per Plans and Specifications"] });
                      setCustomExclusions([...takeoffExclusions]);
                      setStep(3);
                    }}
                    disabled={!rooms.length}
                  >
                    Confirm → Build →
                  </button>
                ) : (
                  <button
                    style={btnPrimary}
                    onClick={() => setStep(3)}
                    disabled={!rooms.length || clarifyRemaining > 0}
                    title={clarifyRemaining > 0 ? `Answer or skip ${clarifyRemaining} remaining clarification${clarifyRemaining === 1 ? "" : "s"} above first.` : undefined}
                  >
                    {clarifyRemaining > 0
                      ? `Build Budget → (${clarifyRemaining} left)`
                      : "Build Budget →"}
                  </button>
                )}
              </div>
            </div>
          )}

          {/* Step 3: Build in JobTread + confirm gate */}
          {step === 3 && (
            <div style={card}>
              <h2 style={{ fontSize: 18, fontWeight: 500, margin: "0 0 8px" }}>Build in JobTread</h2>
              <p style={{ fontSize: 13, color: "var(--color-text-secondary, #666)", marginBottom: 16 }}>Build the budget directly. Talk through any issues that come up.</p>
              {!buildStarted ? (
                <button style={btnPrimary} onClick={() => setBuildStarted(true)}>Start Build</button>
              ) : (
                <BuildChat payload={buildPayload()} onBuilt={(snapshot) => setBuiltBudget(snapshot)} />
              )}

              <hr style={{ margin: "32px 0", border: "none", borderTop: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))" }} />

              {/* Confirm gate — feeds the in-memory snapshot into Scope (or jumps straight to Export for takeoff). */}
              {!builtBudget ? (
                <p style={{ fontSize: 13, color: "var(--color-text-secondary, #666)", margin: 0 }}>
                  Build the budget above. Once it's done, you'll confirm it here before {inputMode === "takeoff" ? "exporting" : "building the scope"}.
                </p>
              ) : (
                <div>
                  <p style={{ fontSize: 13, color: "var(--color-text-secondary, #666)", margin: "0 0 12px" }}>
                    Budget built: {Object.keys(builtBudget.costGroupIds || {}).length} cost group{Object.keys(builtBudget.costGroupIds || {}).length === 1 ? "" : "s"}.
                    Confirm to {inputMode === "takeoff" ? "continue to export" : "derive the scope from this budget"}.
                  </p>
                  <button
                    style={btnPrimary}
                    onClick={() => {
                      setBudgetConfirmed(true);
                      if (inputMode === "takeoff") {
                        setStep(5);
                      } else {
                        categorizeScopeItems();
                      }
                    }}
                    disabled={categorizing}
                  >
                    {categorizing
                      ? "AI categorizing..."
                      : inputMode === "takeoff"
                        ? "✓ Budget looks good — Continue to Export"
                        : "✓ Budget looks good — Build Scope"}
                  </button>
                </div>
              )}

              <div style={{ marginTop: 20 }}>
                <button style={btn} onClick={() => setStep(2)}>← Back</button>
              </div>
            </div>
          )}

          {/* Step 4: Scope Builder (derived from the built budget) */}
          {step === 4 && aiCategorization && (
            <div>
              <div style={card}>
                <h2 style={{ fontSize: 18, fontWeight: 500, margin: "0 0 8px" }}>Scope of Work</h2>
                <p style={{ fontSize: 13, color: "var(--color-text-secondary, #666)", marginBottom: 16 }}>
                  ✅ Recommended (pre-checked) · 💡 Suggested (review) · 📚 Library (expand to browse)
                </p>

                {/* Universal */}
                <ScopeSection
                  groupKey="_universal"
                  title="Universal Prep"
                  icon="🌐"
                  allItems={library.universal_prep[tierMeta.prepTier]}
                  categorization={aiCategorization.universal}
                  selections={scopeSelections["_universal"] || {}}
                  toggle={toggleScopeItem}
                  customItems={customScope["_universal"] || []}
                  addCustom={(t) => addCustomScope("_universal", t)}
                />

                {/* Each scope group resolved from the built budget */}
                {scopeGroups.map((gk) => {
                  const group = library.cost_groups[gk];
                  if (!group) return null;
                  return (
                    <ScopeSection
                      key={gk}
                      groupKey={gk}
                      title={group.label}
                      allItems={group[tierMeta.prepTier]}
                      categorization={aiCategorization.groups?.[gk]}
                      selections={scopeSelections[gk] || {}}
                      toggle={toggleScopeItem}
                      customItems={customScope[gk] || []}
                      addCustom={(t) => addCustomScope(gk, t)}
                    />
                  );
                })}
              </div>

              {/* Exclusions */}
              <div style={card}>
                <h2 style={{ fontSize: 18, fontWeight: 500, margin: "0 0 16px" }}>🛡️ Exclusions</h2>
                <ExclusionSection
                  allItems={library.blanket_exclusions}
                  categorization={aiCategorization.exclusions}
                  selections={exclusionSelections}
                  toggle={toggleExclusion}
                  customItems={customExclusions}
                  addCustom={addCustomExclusion}
                />
              </div>

              <div style={{ display: "flex", gap: 8 }}>
                <button style={btn} onClick={() => setStep(3)}>← Back</button>
                <button style={btnPrimary} onClick={() => setStep(5)}>Next: Export →</button>
              </div>
            </div>
          )}

          {/* Step 5: Export — DocumentChat panel + manual-copy fallback */}
          {step === 5 && (
            <div style={card}>
              <h2 style={{ fontSize: 18, fontWeight: 500, margin: "0 0 8px" }}>Create Proposal Document</h2>
              <p style={{ fontSize: 13, color: "var(--color-text-secondary, #666)", marginBottom: 16 }}>
                Build the customer-facing document directly in JobTread. The chat picks up the budget you just built and mirrors it as line items.
              </p>
              {!documentStarted ? (
                <button style={btnPrimary} onClick={() => setDocumentStarted(true)} disabled={!builtBudget}>
                  {builtBudget ? "Start Document" : "Start Document (build the budget first)"}
                </button>
              ) : (
                <DocumentChat payload={buildPayload()} builtBudget={builtBudget} />
              )}

              <hr style={{ margin: "32px 0", border: "none", borderTop: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))" }} />

              {/* Manual export fallback — kept available as an escape hatch. */}
              <h2 style={{ fontSize: 16, fontWeight: 500, margin: "0 0 8px" }}>Manual Export (fallback)</h2>
              <p style={{ fontSize: 13, color: "var(--color-text-secondary, #666)", marginBottom: 16 }}>If the chat hits an issue, you can still copy the JSON payload and paste it into a separate chat.</p>
              <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 20 }}>
                <button style={btnPrimary} onClick={copyPayload}>
                  📋 Copy Full Payload (for chat)
                  {lastCopyResult?.label === "Full Payload" && Date.now() - lastCopyResult.at < 3000 && (
                    <span style={{ marginLeft: 8, fontSize: 11, opacity: 0.9 }}>✓ copied</span>
                  )}
                </button>
                <button style={btn} onClick={copyScopeText}>
                  Copy Scope Text Only
                  {lastCopyResult?.label === "Scope Text" && Date.now() - lastCopyResult.at < 3000 && (
                    <span style={{ marginLeft: 8, fontSize: 11, color: "var(--color-text-success, #2a7)" }}>✓ copied</span>
                  )}
                </button>
                <button style={btn} onClick={copyExclusionText}>
                  Copy Exclusions Only
                  {lastCopyResult?.label === "Exclusions" && Date.now() - lastCopyResult.at < 3000 && (
                    <span style={{ marginLeft: 8, fontSize: 11, color: "var(--color-text-success, #2a7)" }}>✓ copied</span>
                  )}
                </button>
              </div>
              <details style={{ marginTop: 12 }}>
                <summary style={{ cursor: "pointer", fontSize: 13, color: "var(--color-text-secondary, #666)" }}>Preview payload</summary>
                <pre style={{ marginTop: 8, padding: 12, background: "var(--color-background-secondary, #f3f3f0)", borderRadius: 8, fontSize: 11, fontFamily: "var(--font-mono, monospace)", overflow: "auto", maxHeight: 400 }}>
                  {JSON.stringify(buildPayload(), null, 2)}
                </pre>
              </details>
              <button style={btn} onClick={() => setStep(inputMode === "takeoff" ? 3 : 4)}>← Back to {inputMode === "takeoff" ? "Build" : "Scope"}</button>
            </div>
          )}

          {/* Manual-copy modal — appears when clipboard write fails in iframe */}
          {manualCopyText && (
            <ManualCopyModal
              text={manualCopyText.text}
              label={manualCopyText.label}
              onClose={() => setManualCopyText(null)}
            />
          )}
        </>
      )}
    </div>
  );
}

// ============================================================================
// SUB-COMPONENTS
// ============================================================================

// Classify a line item into Interior / Exterior / T&M for visual grouping
function classifyItemSection(item) {
  const type = item.type || "room";
  if (type === "tm") return "tm";
  if (type === "room") return "interior";   // rooms are always interior
  // type === "scope"
  const sub = (item.substrate || "").toLowerCase();
  const name = (item.name || "").toLowerCase();
  const notes = (item.notes || "").toLowerCase();
  const combined = `${name} ${notes}`;

  // Substrate is the most reliable signal
  if (sub === "exterior_stucco" || sub === "exterior_trim" || sub === "fence_deck") return "exterior";
  if (sub.startsWith("exterior")) return "exterior";

  // Name/notes signals — "exterior" anywhere, or common exterior-only items
  const EXTERIOR_HINTS = /\b(exterior|outside|outdoor|stucco|fascia|eaves|soffit|fence|deck|patio|porch|garage|slider|sliding door|carport|trellis|exterior gate|roof|chimney)\b/;
  if (EXTERIOR_HINTS.test(combined)) return "exterior";

  // Metal is ambiguous; default interior unless above caught it
  // Everything else (walls, ceilings, doors, baseboard, cabinets, trim) → interior
  return "interior";
}

function GroupedLineItemsEditor({ rooms, updateRoom, deleteRoom, addRoom, addScope, addTM }) {
  // Build (idx, item) pairs, then group while preserving original indexes (so updateRoom works correctly)
  const groups = { interior: [], exterior: [], tm: [] };
  rooms.forEach((item, idx) => {
    const section = classifyItemSection(item);
    groups[section].push({ idx, item });
  });

  const card = { background: "var(--color-background-primary, #fff)", border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", borderRadius: 12, padding: "1rem 1.25rem", marginBottom: 16 };
  const btn = { padding: "8px 14px", borderRadius: 8, border: "0.5px solid var(--color-border-secondary, rgba(0,0,0,0.3))", background: "transparent", cursor: "pointer", fontSize: 14, fontFamily: "inherit", color: "inherit" };

  const SectionHeader = ({ icon, title, count, color }) => (
    <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "10px 0 8px", borderBottom: `1px solid ${color}`, marginBottom: 12 }}>
      <span style={{ fontSize: 16 }}>{icon}</span>
      <h3 style={{ fontSize: 14, fontWeight: 600, color, margin: 0, textTransform: "uppercase", letterSpacing: 0.5 }}>{title}</h3>
      <span style={{ fontSize: 11, color: "var(--color-text-secondary, #666)", marginLeft: "auto" }}>{count} item{count === 1 ? "" : "s"}</span>
    </div>
  );

  return (
    <div style={card}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 8 }}>
        <h2 style={{ fontSize: 18, fontWeight: 500, margin: 0 }}>Review Line Items ({rooms.length})</h2>
        <div style={{ display: "flex", gap: 6 }}>
          <button style={{ ...btn, fontSize: 12, padding: "6px 10px" }} onClick={addRoom}>+ Room</button>
          <button style={{ ...btn, fontSize: 12, padding: "6px 10px" }} onClick={addScope}>+ Scope</button>
          <button style={{ ...btn, fontSize: 12, padding: "6px 10px" }} onClick={addTM}>+ T&M</button>
        </div>
      </div>
      <p style={{ fontSize: 12, color: "var(--color-text-secondary, #666)", margin: "0 0 12px" }}>
        <strong>Room</strong> = physical space · <strong>Scope</strong> = substrate bucket · <strong>T&M</strong> = hourly work
      </p>

      {groups.interior.length > 0 && (
        <div style={{ marginBottom: 20 }}>
          <SectionHeader icon="🏠" title="Interior" count={groups.interior.length} color="#2C1654" />
          {groups.interior.map(({ idx, item }) => (
            <RoomCard key={idx} room={item} idx={idx} update={updateRoom} del={deleteRoom} />
          ))}
        </div>
      )}

      {groups.exterior.length > 0 && (
        <div style={{ marginBottom: 20 }}>
          <SectionHeader icon="🏡" title="Exterior" count={groups.exterior.length} color="#C8963E" />
          {groups.exterior.map(({ idx, item }) => (
            <RoomCard key={idx} room={item} idx={idx} update={updateRoom} del={deleteRoom} />
          ))}
        </div>
      )}

      {groups.tm.length > 0 && (
        <div style={{ marginBottom: 8 }}>
          <SectionHeader icon="⏱️" title="Time & Materials" count={groups.tm.length} color="#a70" />
          {groups.tm.map(({ idx, item }) => (
            <RoomCard key={idx} room={item} idx={idx} update={updateRoom} del={deleteRoom} />
          ))}
        </div>
      )}

      {rooms.length === 0 && (
        <div style={{ padding: 24, textAlign: "center", color: "var(--color-text-secondary, #666)", fontSize: 13 }}>
          No items yet. Use the buttons above to add rooms, scope items, or T&M.
        </div>
      )}
    </div>
  );
}

function TakeoffConfirmTable({ items, exclusions }) {
  // Group items by Interior / Exterior so the table reads cleanly
  const groups = { interior: [], exterior: [], tm: [] };
  items.forEach((item) => {
    groups[classifyItemSection(item)].push(item);
  });

  const cellStyle = { padding: "6px 10px", fontSize: 12, borderBottom: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.12))" };
  const headStyle = { ...cellStyle, fontWeight: 600, background: "var(--color-background-secondary, #f3f3f0)", textAlign: "left" };

  const renderSection = (label, color, sectionItems) => {
    if (sectionItems.length === 0) return null;
    return (
      <div key={label} style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 12, fontWeight: 600, color, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 6 }}>
          {label} ({sectionItems.length})
        </div>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
          <thead>
            <tr>
              <th style={{ ...headStyle, width: "40%" }}>Description</th>
              <th style={{ ...headStyle, width: "18%" }}>Substrate</th>
              <th style={{ ...headStyle, width: "12%", textAlign: "right" }}>Qty</th>
              <th style={{ ...headStyle, width: "8%" }}>Unit</th>
              <th style={{ ...headStyle, width: "10%" }}>Coats</th>
              <th style={{ ...headStyle, width: "12%" }}>Section</th>
            </tr>
          </thead>
          <tbody>
            {sectionItems.map((item, i) => (
              <tr key={i}>
                <td style={cellStyle}>{item.name}</td>
                <td style={cellStyle}>{item.substrate || "—"}</td>
                <td style={{ ...cellStyle, textAlign: "right", fontVariantNumeric: "tabular-nums" }}>{item.quantity?.toLocaleString() || "—"}</td>
                <td style={cellStyle}>{item.unit || "—"}</td>
                <td style={cellStyle}>{item.coats || "2coats"}</td>
                <td style={{ ...cellStyle, color: "var(--color-text-secondary, #666)", fontStyle: "italic" }}>{item._sourceSection || "—"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  };

  return (
    <div>
      {renderSection("Interior", "#2C1654", groups.interior)}
      {renderSection("Exterior", "#C8963E", groups.exterior)}
      {renderSection("Time & Materials", "#a70", groups.tm)}

      {exclusions && exclusions.length > 0 && (
        <div style={{ marginTop: 12, padding: 12, background: "var(--color-background-secondary, #f3f3f0)", borderRadius: 8 }}>
          <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 6 }}>
            🛡️ Exclusions from spreadsheet ({exclusions.length})
          </div>
          {exclusions.map((e, i) => (
            <div key={i} style={{ fontSize: 12, padding: "2px 0", color: "var(--color-text-secondary, #444)" }}>• {e}</div>
          ))}
        </div>
      )}
    </div>
  );
}

function RoomCard({ room, idx, update, del }) {
  const type = room.type || "room";
  const TM_CATEGORIES = ["interior_walls", "doors", "windows", "trim", "exterior_walls", "exterior_misc", "millwork", "repairs", "cleaning", "materials"];
  const SCOPE_SUBSTRATES = ["walls", "ceilings", "doors", "baseboard", "trim", "cabinets", "exterior_stucco", "exterior_trim", "fence_deck", "metal"];

  const typeBg = type === "tm" ? "#FFF4E5" : type === "scope" ? "#EEF4FF" : "transparent";
  const typeLabel = type === "tm" ? "T&M (hourly)" : type === "scope" ? "Scope (substrate)" : "Room";

  return (
    <div style={{ border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", borderRadius: 8, padding: 12, marginBottom: 8, background: typeBg }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
        <select value={type} onChange={(e) => update(idx, "type", e.target.value)} style={{ fontSize: 11, padding: "4px 6px", background: "transparent", color: "inherit", border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", borderRadius: 4, fontFamily: "inherit" }}>
          <option value="room">Room</option>
          <option value="scope">Scope</option>
          <option value="tm">T&M</option>
        </select>
        <input value={room.name || ""} onChange={(e) => update(idx, "name", e.target.value)} style={{ flex: 1, padding: 6, fontSize: 14, fontWeight: 500, border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", borderRadius: 6, background: "transparent", color: "inherit", fontFamily: "inherit" }} />
        <button onClick={() => del(idx)} style={{ background: "transparent", border: "none", cursor: "pointer", fontSize: 16, color: "var(--color-text-secondary, #666)" }}>✕</button>
      </div>

      {type === "room" && (
        <>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 8, fontSize: 12, marginBottom: 10 }}>
            {["length", "width", "height"].map((f) => (
              <div key={f}>
                <label style={{ display: "block", color: "var(--color-text-secondary, #666)", marginBottom: 2 }}>{f}</label>
                <input type="number" value={room[f] || 0} onChange={(e) => update(idx, f, parseFloat(e.target.value) || 0)} style={{ width: "100%", padding: 4, fontSize: 13, border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", borderRadius: 4, background: "transparent", color: "inherit", fontFamily: "inherit" }} />
              </div>
            ))}
          </div>

          {/* "Set all to" shortcut */}
          <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 8, fontSize: 11, color: "var(--color-text-secondary, #666)" }}>
            <span>Set all enabled substrates to:</span>
            {["1coat", "2coats", "prime+2"].map((c) => (
              <button key={c} onClick={() => {
                ["walls", "ceiling", "baseboard", "doors"].forEach((sub) => {
                  const cur = room[sub] || { enabled: false, coats: c };
                  if (cur.enabled !== false) update(idx, sub, { ...cur, coats: c });
                });
              }} style={{ padding: "2px 6px", fontSize: 11, borderRadius: 4, border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", background: "transparent", cursor: "pointer", color: "inherit", fontFamily: "inherit" }}>
                {c}
              </button>
            ))}
          </div>

          {/* Per-substrate rows */}
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            {[
              { key: "walls", label: "walls" },
              { key: "ceiling", label: "ceiling" },
              { key: "baseboard", label: "baseboard" },
              { key: "doors", label: "doors", hasCount: true },
            ].map(({ key, label, hasCount }) => {
              const sub = room[key] || { enabled: false, coats: "2coats" };
              return (
                <div key={key} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13 }}>
                  <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", minWidth: 90 }}>
                    <input type="checkbox" checked={sub.enabled !== false} onChange={(e) => update(idx, key, { ...sub, enabled: e.target.checked })} />
                    {label}
                  </label>
                  {hasCount && (
                    <input type="number" min="0" value={sub.count || 0} onChange={(e) => update(idx, key, { ...sub, count: parseInt(e.target.value) || 0 })} placeholder="count" disabled={sub.enabled === false} style={{ width: 60, padding: 3, fontSize: 12, border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", borderRadius: 4, background: "transparent", color: "inherit", fontFamily: "inherit", opacity: sub.enabled === false ? 0.4 : 1 }} />
                  )}
                  <select value={sub.coats || "2coats"} onChange={(e) => update(idx, key, { ...sub, coats: e.target.value })} disabled={sub.enabled === false} style={{ fontSize: 12, padding: "2px 4px", background: "transparent", color: "inherit", border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", borderRadius: 4, opacity: sub.enabled === false ? 0.4 : 1 }}>
                    <option value="1coat">1 coat</option>
                    <option value="2coats">2 coats</option>
                    <option value="prime+2">Prime + 2</option>
                  </select>
                </div>
              );
            })}
          </div>
        </>
      )}

      {type === "scope" && (
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 8, fontSize: 12 }}>
          <div>
            <label style={{ display: "block", color: "var(--color-text-secondary, #666)", marginBottom: 2 }}>substrate</label>
            <select value={room.substrate || "doors"} onChange={(e) => update(idx, "substrate", e.target.value)} style={{ width: "100%", padding: 4, fontSize: 12, background: "transparent", color: "inherit", border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", borderRadius: 4 }}>
              {SCOPE_SUBSTRATES.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
          <div>
            <label style={{ display: "block", color: "var(--color-text-secondary, #666)", marginBottom: 2 }}>qty</label>
            <input type="number" value={room.quantity || 0} onChange={(e) => update(idx, "quantity", parseFloat(e.target.value) || 0)} style={{ width: "100%", padding: 4, fontSize: 13, border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", borderRadius: 4, background: "transparent", color: "inherit", fontFamily: "inherit" }} />
          </div>
          <div>
            <label style={{ display: "block", color: "var(--color-text-secondary, #666)", marginBottom: 2 }}>unit</label>
            <select value={room.unit || "EA"} onChange={(e) => update(idx, "unit", e.target.value)} style={{ width: "100%", padding: 4, fontSize: 12, background: "transparent", color: "inherit", border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", borderRadius: 4 }}>
              <option value="EA">EA</option>
              <option value="LF">LF</option>
              <option value="SF">SF</option>
            </select>
          </div>
          <div>
            <label style={{ display: "block", color: "var(--color-text-secondary, #666)", marginBottom: 2 }}>coats</label>
            <select value={room.coats || "2coats"} onChange={(e) => update(idx, "coats", e.target.value)} style={{ width: "100%", padding: 4, fontSize: 12, background: "transparent", color: "inherit", border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", borderRadius: 4 }}>
              <option value="1coat">1 coat</option>
              <option value="2coats">2 coats</option>
              <option value="prime+2">Prime + 2</option>
            </select>
          </div>
        </div>
      )}

      {type === "tm" && (
        <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr", gap: 8, fontSize: 12 }}>
          <div>
            <label style={{ display: "block", color: "var(--color-text-secondary, #666)", marginBottom: 2 }}>hours</label>
            <input type="number" step="0.5" value={room.hours || 0} onChange={(e) => update(idx, "hours", parseFloat(e.target.value) || 0)} style={{ width: "100%", padding: 4, fontSize: 13, border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", borderRadius: 4, background: "transparent", color: "inherit", fontFamily: "inherit" }} />
          </div>
          <div>
            <label style={{ display: "block", color: "var(--color-text-secondary, #666)", marginBottom: 2 }}>category (maps to T&M code)</label>
            <select value={room.category || "repairs"} onChange={(e) => update(idx, "category", e.target.value)} style={{ width: "100%", padding: 4, fontSize: 12, background: "transparent", color: "inherit", border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", borderRadius: 4 }}>
              {TM_CATEGORIES.map((c) => {
                const tm = TM_CATALOG.find((t) => t.category === c) || TM_CATALOG.find((t) => t.codeName.toLowerCase().includes(c.replace("_", " ")));
                return <option key={c} value={c}>{c} {tm ? `→ ${tm.code}` : ""}</option>;
              })}
            </select>
          </div>
        </div>
      )}

      {room.notes && <div style={{ fontSize: 12, color: "var(--color-text-secondary, #666)", marginTop: 6, fontStyle: "italic" }}>Note: {room.notes}</div>}
    </div>
  );
}

function ScopeSection({ groupKey, title, icon, allItems, categorization, selections, toggle, customItems, addCustom }) {
  const [showLibrary, setShowLibrary] = useState(false);
  const [customInput, setCustomInput] = useState("");

  const recommended = categorization?.recommended || [];
  const suggested = categorization?.suggested || [];
  const reasons = categorization?.reasons || {};

  // "Library" = everything in allItems that's not in recommended or suggested
  const inRec = new Set(recommended);
  const inSug = new Set(suggested);
  const libraryItems = (allItems || []).filter((i) => !inRec.has(i) && !inSug.has(i));

  return (
    <div style={{ marginBottom: 20, paddingBottom: 16, borderBottom: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))" }}>
      <h3 style={{ fontSize: 15, fontWeight: 500, margin: "0 0 10px" }}>{icon ? `${icon} ` : ""}{title}</h3>

      {/* Recommended */}
      {recommended.length > 0 && (
        <div style={{ marginBottom: 10 }}>
          <div style={{ fontSize: 11, fontWeight: 500, color: "var(--color-text-success, #2a7)", marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.5 }}>✅ Recommended</div>
          {recommended.map((item) => (
            <CheckboxRow key={item} item={item} checked={!!selections[item]} onChange={() => toggle(groupKey, item)} />
          ))}
        </div>
      )}

      {/* Suggested */}
      {suggested.length > 0 && (
        <div style={{ marginBottom: 10 }}>
          <div style={{ fontSize: 11, fontWeight: 500, color: "var(--color-text-warning, #a70)", marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.5 }}>💡 Suggested</div>
          {suggested.map((item) => (
            <CheckboxRow key={item} item={item} checked={!!selections[item]} onChange={() => toggle(groupKey, item)} reason={reasons[item]} />
          ))}
        </div>
      )}

      {/* Custom items already added */}
      {customItems.length > 0 && (
        <div style={{ marginBottom: 10 }}>
          <div style={{ fontSize: 11, fontWeight: 500, color: "var(--color-text-secondary, #666)", marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.5 }}>✏️ Custom</div>
          {customItems.map((item) => (
            <CheckboxRow key={item} item={item} checked={!!selections[item]} onChange={() => toggle(groupKey, item)} />
          ))}
        </div>
      )}

      {/* Library (collapsible) */}
      {libraryItems.length > 0 && (
        <div>
          <button style={{ background: "transparent", border: "none", color: "var(--color-text-secondary, #666)", fontSize: 12, cursor: "pointer", padding: 4, fontFamily: "inherit" }} onClick={() => setShowLibrary(!showLibrary)}>
            {showLibrary ? "▼" : "▶"} 📚 Library ({libraryItems.length} more)
          </button>
          {showLibrary && libraryItems.map((item) => (
            <CheckboxRow key={item} item={item} checked={!!selections[item]} onChange={() => toggle(groupKey, item)} />
          ))}
        </div>
      )}

      {/* Add custom */}
      <div style={{ display: "flex", gap: 6, marginTop: 8 }}>
        <input value={customInput} onChange={(e) => setCustomInput(e.target.value)} placeholder="+ Add custom prep item..." style={{ flex: 1, padding: 6, fontSize: 12, border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", borderRadius: 6, background: "transparent", color: "inherit", fontFamily: "inherit" }} />
          <button style={{ padding: "4px 10px", fontSize: 12, borderRadius: 6, border: "0.5px solid var(--color-border-secondary, rgba(0,0,0,0.3))", background: "transparent", cursor: "pointer", fontFamily: "inherit", color: "inherit" }} onClick={() => { addCustom(customInput); setCustomInput(""); }}>Add</button>
      </div>
    </div>
  );
}

function ExclusionSection({ allItems, categorization, selections, toggle, customItems, addCustom }) {
  const [customInput, setCustomInput] = useState("");
  const recommended = categorization?.recommended || allItems || [];
  const suggested = categorization?.suggested || [];
  const reasons = categorization?.reasons || {};

  return (
    <div>
      <div style={{ marginBottom: 10 }}>
        <div style={{ fontSize: 11, fontWeight: 500, color: "var(--color-text-success, #2a7)", marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.5 }}>✅ Recommended</div>
        {recommended.map((item) => (
          <CheckboxRow key={item} item={item} checked={selections[item] !== false} onChange={() => toggle(item)} />
        ))}
      </div>
      {suggested.length > 0 && (
        <div style={{ marginBottom: 10 }}>
          <div style={{ fontSize: 11, fontWeight: 500, color: "var(--color-text-warning, #a70)", marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.5 }}>💡 Suggested</div>
          {suggested.map((item) => (
            <CheckboxRow key={item} item={item} checked={!!selections[item]} onChange={() => toggle(item)} reason={reasons[item]} />
          ))}
        </div>
      )}
      {customItems.length > 0 && (
        <div style={{ marginBottom: 10 }}>
          <div style={{ fontSize: 11, fontWeight: 500, color: "var(--color-text-secondary, #666)", marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.5 }}>✏️ Custom</div>
          {customItems.map((item) => (
            <CheckboxRow key={item} item={item} checked={!!selections[item]} onChange={() => toggle(item)} />
          ))}
        </div>
      )}
      <div style={{ display: "flex", gap: 6, marginTop: 8 }}>
        <input value={customInput} onChange={(e) => setCustomInput(e.target.value)} placeholder="+ Add custom exclusion..." style={{ flex: 1, padding: 6, fontSize: 12, border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", borderRadius: 6, background: "transparent", color: "inherit", fontFamily: "inherit" }} />
        <button style={{ padding: "4px 10px", fontSize: 12, borderRadius: 6, border: "0.5px solid var(--color-border-secondary, rgba(0,0,0,0.3))", background: "transparent", cursor: "pointer", fontFamily: "inherit", color: "inherit" }} onClick={() => { addCustom(customInput); setCustomInput(""); }}>Add</button>
      </div>
    </div>
  );
}

function CheckboxRow({ item, checked, onChange, reason }) {
  const [showReason, setShowReason] = useState(false);
  return (
    <div style={{ display: "flex", alignItems: "flex-start", gap: 8, padding: "4px 0", fontSize: 13 }}>
      <input type="checkbox" checked={checked} onChange={onChange} style={{ width: 16, height: 16, marginTop: 2, flexShrink: 0 }} />
      <div style={{ flex: 1 }}>
        <span style={{ cursor: "pointer" }} onClick={onChange}>{item}</span>
        {reason && (
          <>
            <button onClick={() => setShowReason(!showReason)} style={{ background: "transparent", border: "none", cursor: "pointer", fontSize: 11, color: "var(--color-text-info, #46c)", marginLeft: 6, padding: 0, fontFamily: "inherit" }}>?</button>
            {showReason && <div style={{ fontSize: 11, color: "var(--color-text-secondary, #666)", marginTop: 2, fontStyle: "italic" }}>{reason}</div>}
          </>
        )}
      </div>
    </div>
  );
}

function SettingsPanel({ library, setLibrary }) {
  const [editing, setEditing] = useState("blanket_exclusions");
  const [editText, setEditText] = useState("");

  useEffect(() => {
    if (editing === "blanket_exclusions") {
      setEditText(library.blanket_exclusions.join("\n"));
    } else if (editing.startsWith("universal_")) {
      const tier = editing.replace("universal_", "");
      setEditText((library.universal_prep[tier] || []).join("\n"));
    } else if (editing.includes(":")) {
      const [gk, tier] = editing.split(":");
      setEditText((library.cost_groups[gk]?.[tier] || []).join("\n"));
    }
  }, [editing, library]);

  async function saveEdit() {
    const lines = editText.split("\n").map((l) => l.trim()).filter(Boolean);
    const next = { ...library };
    if (editing === "blanket_exclusions") {
      next.blanket_exclusions = lines;
    } else if (editing.startsWith("universal_")) {
      const tier = editing.replace("universal_", "");
      next.universal_prep = { ...next.universal_prep, [tier]: lines };
    } else if (editing.includes(":")) {
      const [gk, tier] = editing.split(":");
      next.cost_groups = {
        ...next.cost_groups,
        [gk]: { ...next.cost_groups[gk], [tier]: lines },
      };
    }
    setLibrary(next);
    const saved = await saveToStorage(STORAGE_KEYS.library, next);
    alert(saved ? "Saved" : "Save failed");
  }

  async function resetToDefault() {
    if (!confirm("Reset library to default? Custom changes will be lost.")) return;
    setLibrary(DEFAULT_LIBRARY);
    await saveToStorage(STORAGE_KEYS.library, DEFAULT_LIBRARY);
    alert("Library reset to default");
  }

  const card = { background: "var(--color-background-primary, #fff)", border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", borderRadius: 12, padding: "1rem 1.25rem", marginBottom: 16 };
  const btn = { padding: "8px 14px", borderRadius: 8, border: "0.5px solid var(--color-border-secondary, rgba(0,0,0,0.3))", background: "transparent", cursor: "pointer", fontSize: 13, fontFamily: "inherit", color: "inherit" };

  return (
    <div style={card}>
      <h2 style={{ fontSize: 18, fontWeight: 500, margin: "0 0 16px" }}>Scope Library Settings</h2>
      <p style={{ fontSize: 13, color: "var(--color-text-secondary, #666)", marginBottom: 16 }}>Edit the master prep / exclusion lists. Changes save to your browser and apply to all future proposals.</p>

      <div style={{ display: "grid", gridTemplateColumns: "220px 1fr", gap: 16 }}>
        <div style={{ display: "flex", flexDirection: "column", gap: 4, fontSize: 13 }}>
          <SettingsButton current={editing} value="blanket_exclusions" onClick={setEditing}>Blanket Exclusions</SettingsButton>
          <div style={{ fontSize: 11, fontWeight: 500, color: "var(--color-text-secondary, #666)", marginTop: 12, marginBottom: 4, textTransform: "uppercase" }}>Universal Prep</div>
          <SettingsButton current={editing} value="universal_standard" onClick={setEditing}>Standard tier</SettingsButton>
          <SettingsButton current={editing} value="universal_high_end" onClick={setEditing}>High-End tier</SettingsButton>
          <div style={{ fontSize: 11, fontWeight: 500, color: "var(--color-text-secondary, #666)", marginTop: 12, marginBottom: 4, textTransform: "uppercase" }}>Cost Groups</div>
          {Object.entries(library.cost_groups).map(([gk, grp]) => (
            <div key={gk} style={{ marginBottom: 6 }}>
              <div style={{ fontSize: 12, fontWeight: 500, marginBottom: 2 }}>{grp.label}</div>
              <SettingsButton current={editing} value={`${gk}:standard`} onClick={setEditing}>· Standard</SettingsButton>
              <SettingsButton current={editing} value={`${gk}:high_end`} onClick={setEditing}>· High-End</SettingsButton>
            </div>
          ))}
        </div>
        <div>
          <p style={{ fontSize: 12, color: "var(--color-text-secondary, #666)", marginBottom: 8 }}>One item per line. Blank lines ignored.</p>
          <textarea value={editText} onChange={(e) => setEditText(e.target.value)} rows={16} style={{ width: "100%", padding: 10, borderRadius: 8, border: "0.5px solid var(--color-border-secondary, rgba(0,0,0,0.3))", fontSize: 13, fontFamily: "var(--font-mono, monospace)", background: "transparent", color: "inherit", resize: "vertical" }} />
          <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
            <button style={{ ...btn, background: "#2C1654", color: "#fff", border: "0.5px solid #2C1654" }} onClick={saveEdit}>Save Changes</button>
            <button style={btn} onClick={resetToDefault}>Reset to Default</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function ClarificationCard({ question, answer, setAnswer, applyToItem, applyMultiPatch }) {
  const [showCustom, setShowCustom] = useState(false);
  const [customText, setCustomText] = useState("");

  const confidenceColor = {
    high: "var(--color-text-success, #2a7)",
    medium: "var(--color-text-warning, #a70)",
    low: "var(--color-text-danger, #c33)",
  }[question.confidence] || "var(--color-text-secondary, #666)";

  const isResolved = answer !== null && answer !== undefined;
  const options = question.options || [];

  function selectOption(option) {
    if (option.value === "__custom__") {
      setShowCustom(true);
      return;
    }
    // Multi-item patch (with optional formula strings)
    if (option.value && typeof option.value === "object" && Array.isArray(option.value._multiPatch)) {
      const count = applyMultiPatch ? applyMultiPatch(option.value._multiPatch) : 0;
      // Build a friendly answer label including count + any formula hints
      const hints = option.value._multiPatch
        .map((p) => p.patch?._formula)
        .filter(Boolean);
      const baseLabel = option.label;
      const suffix = count > 0
        ? ` — updated ${count} item${count === 1 ? "" : "s"}${hints.length ? ` (${hints[0]})` : ""}`
        : (hints.length ? ` — ${hints[0]}` : "");
      setAnswer(baseLabel + suffix);
      return;
    }
    // Single-item patch
    if (question.itemIndex != null && typeof option.value === "object") {
      applyToItem(option.value);
    }
    setAnswer(option.label);
  }

  async function submitCustom() {
    if (!customText.trim()) return;
    setAnswer(customText.trim());
    setShowCustom(false);

    // Try to analyze the custom text and apply intelligently
    if (question.itemIndex != null) {
      const patch = parseCustomAnswer(question.field, customText.trim());
      if (patch) applyToItem(patch);
    }
  }

  return (
    <div style={{ border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", borderRadius: 8, padding: 12, marginBottom: 8, background: isResolved ? "var(--color-background-secondary, #f3f3f0)" : "transparent" }}>
      <div style={{ fontSize: 13, fontWeight: 500, marginBottom: 4 }}>
        {question.prompt}
        {question.appliesTo && <span style={{ fontWeight: 400, color: "var(--color-text-secondary, #666)" }}> — {question.appliesTo}</span>}
      </div>
      <div style={{ fontSize: 11, color: confidenceColor, marginBottom: 10 }}>
        Confidence: {question.confidence || "unknown"}
      </div>

      {isResolved ? (
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ fontSize: 13, padding: "4px 10px", background: "var(--color-background-primary, #fff)", border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", borderRadius: 6 }}>
            ✓ {answer}
          </span>
          <button
            style={{ padding: "4px 10px", fontSize: 12, borderRadius: 6, border: "0.5px solid var(--color-border-secondary, rgba(0,0,0,0.3))", background: "transparent", cursor: "pointer", color: "inherit", fontFamily: "inherit" }}
            onClick={() => { setAnswer(null); setShowCustom(false); setCustomText(""); }}
          >
            Change
          </button>
        </div>
      ) : showCustom ? (
        <div>
          <div style={{ fontSize: 11, color: "var(--color-text-secondary, #666)", marginBottom: 6 }}>
            Describe in your own words — Claude will figure out how to apply it.
          </div>
          <div style={{ display: "flex", gap: 6 }}>
            <input
              value={customText}
              onChange={(e) => setCustomText(e.target.value)}
              autoFocus
              style={{ flex: 1, padding: 6, fontSize: 13, border: "0.5px solid var(--color-border-secondary, rgba(0,0,0,0.3))", borderRadius: 6, background: "transparent", color: "inherit", fontFamily: "inherit" }}
              placeholder="e.g. 22 LF, 2 coats, color change"
              onKeyDown={(e) => { if (e.key === "Enter") submitCustom(); }}
            />
            <button
              style={{ padding: "4px 10px", fontSize: 12, borderRadius: 6, border: "0.5px solid #2C1654", background: "#2C1654", color: "#fff", cursor: "pointer", fontFamily: "inherit" }}
              onClick={submitCustom}
            >
              Apply
            </button>
            <button
              style={{ padding: "4px 10px", fontSize: 12, borderRadius: 6, border: "0.5px solid var(--color-border-secondary, rgba(0,0,0,0.3))", background: "transparent", cursor: "pointer", color: "inherit", fontFamily: "inherit" }}
              onClick={() => { setShowCustom(false); setCustomText(""); }}
            >
              Cancel
            </button>
          </div>
        </div>
      ) : options.length > 0 ? (
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
          {options.map((opt, i) => (
            <button
              key={i}
              onClick={() => selectOption(opt)}
              style={{
                padding: "6px 12px",
                fontSize: 12,
                borderRadius: 6,
                border: opt.value === "__custom__"
                  ? "0.5px dashed var(--color-border-secondary, rgba(0,0,0,0.3))"
                  : "0.5px solid var(--color-border-secondary, rgba(0,0,0,0.3))",
                background: "transparent",
                cursor: "pointer",
                color: "inherit",
                fontFamily: "inherit",
                fontStyle: opt.value === "__custom__" ? "italic" : "normal",
              }}
            >
              {opt.label}
            </button>
          ))}
        </div>
      ) : (
        // Fallback: no options provided (legacy questions)
        <div style={{ display: "flex", gap: 6 }}>
          <input
            value={customText}
            onChange={(e) => setCustomText(e.target.value)}
            style={{ flex: 1, padding: 6, fontSize: 13, border: "0.5px solid var(--color-border-secondary, rgba(0,0,0,0.3))", borderRadius: 6, background: "transparent", color: "inherit", fontFamily: "inherit" }}
            placeholder="Type answer..."
          />
          <button
            style={{ padding: "4px 10px", fontSize: 12, borderRadius: 6, border: "0.5px solid #2C1654", background: "#2C1654", color: "#fff", cursor: "pointer", fontFamily: "inherit" }}
            onClick={submitCustom}
          >
            Save
          </button>
        </div>
      )}
    </div>
  );
}

// ============================================================================
// CLARIFY CHAT — one-question-at-a-time conversational interface. Builds an
// ordered queue of clarifying questions (mandatory dim/door gaps first, then
// parser hints) and walks through them as chat turns. Each active question
// renders with quick-pick chips (instant apply via the existing pattern-match
// path), a free-text input (routed through the existing Claude API engine),
// and a Skip button (rotates the question to the end). Power-use freeform
// still works: typing several answers at once applies them all, and the queue
// silently advances past any items those answers resolved.
// ============================================================================

// Common-size chip options for a room with missing dimensions, hinted by the
// room's name. Each option's value is a single-item patch ({length, width})
// applied via applyToItem(itemIndex, value).
function clarifyDimChips(room) {
  const name = (room?.name || "").toLowerCase();
  if (name.includes("bath")) {
    return [
      { label: "5 × 8 (small)", value: { length: 5, width: 8 } },
      { label: "7 × 10 (typical)", value: { length: 7, width: 10 } },
      { label: "10 × 12 (master)", value: { length: 10, width: 12 } },
    ];
  }
  if (name.includes("closet")) {
    return [
      { label: "4 × 6 (reach-in)", value: { length: 4, width: 6 } },
      { label: "6 × 8 (walk-in)", value: { length: 6, width: 8 } },
    ];
  }
  if (name.includes("bed") || name.includes("master") || name.includes("guest")) {
    return [
      { label: "10 × 12", value: { length: 10, width: 12 } },
      { label: "12 × 14", value: { length: 12, width: 14 } },
      { label: "14 × 16 (primary)", value: { length: 14, width: 16 } },
    ];
  }
  if (name.includes("kitchen") || name.includes("dining") || name.includes("living") || name.includes("family")) {
    return [
      { label: "12 × 14", value: { length: 12, width: 14 } },
      { label: "14 × 16", value: { length: 14, width: 16 } },
      { label: "16 × 20", value: { length: 16, width: 20 } },
    ];
  }
  return [
    { label: "10 × 12", value: { length: 10, width: 12 } },
    { label: "12 × 14", value: { length: 12, width: 14 } },
    { label: "14 × 16", value: { length: 14, width: 16 } },
  ];
}

// Build the ordered question queue: mandatory dim/door gaps first, then
// parser hints. Each entry carries an optional predicate(rooms)=>bool that's
// true while the question is still unanswered — used to silently drop items
// from the queue when a freeform answer resolved more than just the active
// one. dim/door predicates match by room name (robust to index shifts from
// adds/removes), parser hints have no predicate (stay until user touches them).
function buildClarifyQueue(rooms, hints) {
  const queue = [];

  rooms.forEach((r, idx) => {
    if ((r?.type || "room") !== "room") return;
    if (r.length && r.width && r.length !== 0 && r.width !== 0) return;
    const name = r.name || "Unnamed room";
    queue.push({
      id: `dim-${idx}-${name}`,
      prompt: `"${name}" is missing length × width — what are the dimensions?`,
      options: clarifyDimChips(r),
      itemIndex: idx,
      roomName: name,
      source: "dim",
      predicate: (cur) => {
        const room = (cur || []).find((rr) => (rr?.name || "") === name && (rr?.type || "room") === "room");
        if (!room) return false;
        return !room.length || !room.width || room.length === 0 || room.width === 0;
      },
    });
  });

  rooms.forEach((r, idx) => {
    if ((r?.type || "room") !== "room") return;
    if (r.doors?.enabled === false) return;
    if (r.doors?.count > 0) return;
    const name = r.name || "Unnamed room";
    queue.push({
      id: `door-${idx}-${name}`,
      prompt: `"${name}" has doors enabled but no door count — how many?`,
      options: [
        { label: "1 door", value: { doors: { count: 1 } } },
        { label: "2 doors", value: { doors: { count: 2 } } },
        { label: "3 doors", value: { doors: { count: 3 } } },
      ],
      itemIndex: idx,
      roomName: name,
      source: "door",
      predicate: (cur) => {
        const room = (cur || []).find((rr) => (rr?.name || "") === name && (rr?.type || "room") === "room");
        if (!room) return false;
        if (room.doors?.enabled === false) return false;
        return !(room.doors?.count > 0);
      },
    });
  });

  (hints || []).slice(0, 6).forEach((h, i) => {
    const opts = (h.options || []).filter((o) => o.value !== "__custom__").slice(0, 4);
    queue.push({
      id: `hint-${h.id || i}`,
      prompt: h.prompt,
      options: opts,
      itemIndex: h.itemIndex,
      source: "parserHint",
      hint: h,
    });
  });

  return queue;
}

function ClarifyChat({ initialHints, rooms, library, applyToItem, applyMultiPatch, onQueueChange }) {
  // Past Q&A exchanges live in `messages`. The currently-active question lives
  // separately at queue[0] and renders as a special last bubble with chips +
  // Skip. When the user answers (chip / text / skip), the exchange gets pushed
  // to messages and the queue advances.
  const [messages, setMessages] = useState([]);
  const [queue, setQueue] = useState(() => buildClarifyQueue(rooms, initialHints));
  const [inputText, setInputText] = useState("");
  const [thinking, setThinking] = useState(false);
  const [error, setError] = useState("");
  const scrollerRef = useRef(null);

  // Welcome line when there are no gaps to walk through.
  useEffect(() => {
    if (queue.length === 0 && messages.length === 0) {
      setMessages([{
        role: "assistant",
        text: `I parsed your input cleanly. Tell me anything you want to change — "1 coat on all doors", "remove the master bath", "kitchen is 14x16", "add a hall closet 6x4", etc.`,
      }]);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Tell the parent how many questions remain — gates the Build Scope button.
  useEffect(() => {
    if (typeof onQueueChange === "function") onQueueChange(queue.length);
  }, [queue.length, onQueueChange]);

  // Auto-scroll on any visible change.
  useEffect(() => {
    if (scrollerRef.current) {
      scrollerRef.current.scrollTop = scrollerRef.current.scrollHeight;
    }
  }, [messages, queue.length, thinking]);

  // After any rooms change, prune queue items whose predicate says they're
  // resolved. This is what powers "type several answers at once, advance past
  // multiple questions." Parser hints (no predicate) stick around.
  useEffect(() => {
    setQueue((prev) => prev.filter((q) => !q.predicate || q.predicate(rooms)));
  }, [rooms]);

  const active = queue[0] || null;

  // Push the answered exchange into the message log and pop the active item.
  function pushAnsweredAndPop(activeItem, userText, replyText) {
    setMessages((prev) => [
      ...prev,
      { role: "assistant", text: activeItem.prompt },
      { role: "user", text: userText },
      { role: "assistant", text: replyText },
    ]);
    setQueue((prev) => prev.slice(1));
  }

  // ── Chip click — apply via the existing pattern-match path. Does NOT route
  //    through the API engine; the chip already carries a patch directive.
  function handleChip(option) {
    if (!active || thinking) return;
    if (option.value === "__custom__") return;

    // Resolve the current index by name for dim/door questions in case rooms
    // shifted since the queue was built. Parser hints fall back to the stored
    // itemIndex from the hint payload.
    let itemIdx = active.itemIndex;
    if (active.roomName) {
      const found = rooms.findIndex((r) => (r?.name || "") === active.roomName);
      if (found >= 0) itemIdx = found;
    }

    let changedCount = 0;
    const val = option.value;
    if (val && typeof val === "object" && Array.isArray(val._multiPatch)) {
      changedCount = applyMultiPatch ? applyMultiPatch(val._multiPatch) : 0;
    } else if (val && typeof val === "object" && itemIdx != null && itemIdx >= 0) {
      applyToItem(itemIdx, val);
      changedCount = 1;
    } else if (val && typeof val === "object" && applyMultiPatch) {
      changedCount = applyMultiPatch([{ patch: val }]);
    }

    const confirm = changedCount > 0
      ? `Updated ${changedCount} item${changedCount === 1 ? "" : "s"}.`
      : "Noted.";
    pushAnsweredAndPop(active, option.label, confirm);
  }

  // ── Skip — rotate the active question to the end of the queue.
  function handleSkip() {
    if (!active || thinking) return;
    const skipped = active;
    setMessages((prev) => [
      ...prev,
      { role: "assistant", text: skipped.prompt },
      { role: "user", text: "Skip for now." },
      { role: "assistant", text: "Coming back to it." },
    ]);
    setQueue((prev) => [...prev.slice(1), skipped]);
  }

  // ── Send the current input as a user message. Calls Claude to interpret +
  //    apply, then renders the model's natural-language reply.
  async function sendMessage() {
    if (!inputText.trim() || thinking) return;
    const userText = inputText.trim();
    setInputText("");
    setError("");

    setThinking(true);
    try {
      // Build a compact room digest. Limit to 40 items for prompt size.
      const roomDigest = rooms.slice(0, 40).map((r, idx) => formatItemDigest(r, idx)).join("\n");

      // Send the conversation as text-only history. Include the active prompt
      // (if any) as the latest assistant turn so the model has context.
      const apiHistory = [
        ...messages.map((m) => ({ role: m.role, content: m.text })),
        ...(active ? [{ role: "assistant", content: active.prompt }] : []),
        { role: "user", content: userText },
      ];

      const systemContext = `You are an interactive estimating assistant inside a painting takeoff tool. You're talking to a contractor reviewing items parsed from their job walk. Your job: interpret what they say and translate it into action directives that the app will execute, then reply naturally about what you did.

CURRENT LINE ITEMS (index in brackets — use _itemIndex to target precisely):
${roomDigest}

Respond ONLY with JSON of shape:
{
  "actions": [<action>, ...],   // can be empty if no change needed
  "reply": "<short natural reply describing what you did or asking for clarification>"
}

No preamble, no markdown fences. Just JSON.

ACTION SHAPES (each action has exactly ONE selector):
  { "_itemIndex": <int>, "patch": {...} }                — exact index
  { "itemNameMatch": "<substr>", "patch": {...} }        — case-insensitive name substring
  { "substrateMatch": "doors|walls|ceilings|baseboard|trim", "patch": {...} }
  { "itemNameMatch": "...", "substrateMatch": "...", "patch": {...} }  — both; substrate sub-targets within matched item
  { "_remove": true, "_itemIndex": <int> }               — delete an item
  { "_remove": true, "itemNameMatch": "<substr>" }       — delete all matching
  { "_add": <newItem> }                                  — append a new line item

patch fields (apply to relevant items):
  coats: "1coat" | "2coats" | "prime+2"
  quantity: number OR mathjs formula string
  unit: "SF" | "LF" | "EA" | "HR"
  length, width, height: number
  enabled: bool   — toggle substrate on/off in a room
  name: string    — rename
  doorCount: int  — number of doors in a room
  hours: number   — T&M only

_add newItem shape examples:
  Room: { "type":"room", "name":"Hall Closet", "length":6, "width":4, "height":9, "walls":{"enabled":true,"coats":"2coats"}, "ceiling":{"enabled":true,"coats":"2coats"}, "baseboard":{"enabled":true,"coats":"2coats"}, "doors":{"enabled":true,"coats":"2coats","count":1} }
  Scope: { "type":"scope", "name":"Exterior Fascia", "substrate":"exterior_trim", "quantity":210, "unit":"LF", "coats":"2coats" }
  T&M:   { "type":"tm", "name":"Repairs", "hours":4, "category":"repairs" }

EXTERIOR BODY / STUCCO — CRITICAL RULE:
exterior_stucco quantity in SF = body perimeter (LF) × wall height (ft). A bare perimeter is NOT the SF.
If you see an exterior body item with quantity that looks like a raw perimeter (e.g. 210 SF for a whole house — that's way too low), point it out and ASK for the wall height before changing it. Then patch with quantity: "<perimeter> * <height>" as a formula string.

EXTERIOR EAVES / SOFFITS — DIFFERENT RULE:
exterior eaves quantity in SF = eave run length (LF) × eave width (ft, the overhang depth — typically 1.5-3 ft).
Eave run length is NOT necessarily the same as body perimeter — some elevations may have no overhang. If the user hasn't told you which is which, default to body perimeter as the eave run but flag it. NEVER use wall height for eaves. NEVER store a bare perimeter as the SF for eaves.

EXAMPLES:
User: "all walls 2 coats not 1"
→ {"actions":[{"substrateMatch":"walls","patch":{"coats":"2coats"}}],"reply":"Switched all walls to 2 coats."}

User: "remove the dining room"
→ {"actions":[{"_remove":true,"itemNameMatch":"dining"}],"reply":"Removed the dining room."}

User: "kitchen is actually 16x14"
→ {"actions":[{"itemNameMatch":"kitchen","patch":{"length":16,"width":14}}],"reply":"Updated kitchen to 16×14."}

User: "add a hall closet 6x4 doors and trim only"
→ {"actions":[{"_add":{"type":"room","name":"Hall Closet","length":6,"width":4,"height":9,"walls":{"enabled":false},"ceiling":{"enabled":false},"baseboard":{"enabled":true,"coats":"2coats"},"doors":{"enabled":true,"coats":"2coats","count":1}}}],"reply":"Added Hall Closet (6×4) — doors and trim only."}

User: "exterior body is 210 perimeter, 10ft walls"
→ {"actions":[{"itemNameMatch":"exterior body","patch":{"quantity":"210 * 10","unit":"SF"}}],"reply":"Exterior body set to 210 × 10 = 2100 SF."}

User: "the walls are 18ft tall, it's a two story"  (when exterior body already exists with a placeholder perimeter)
→ {"actions":[{"itemNameMatch":"exterior body","patch":{"quantity":"<existing perimeter> * 18","unit":"SF"}}],"reply":"Updated exterior body for 18ft walls."}

User: "eaves are 180 LF run, 2.5 ft overhang"
→ {"actions":[{"itemNameMatch":"exterior eaves","patch":{"quantity":"180 * 2.5","unit":"SF"}}],"reply":"Eaves set to 180 × 2.5 = 450 SF."}

User: "eaves are 24 inches"  (when eaves item exists, length unknown)
→ {"actions":[],"reply":"Got the 2 ft overhang width. What's the total eave run length (LF)? Default would be body perimeter (210 LF) if all elevations have eaves."}

User: "what's the door count on the bedrooms?"
→ {"actions":[],"reply":"Bedroom 1 has 1 door, Bedroom 2 has 1 door, Bedroom 3 has 2 doors."}

User: "scratch that, undo"
→ {"actions":[],"reply":"I can't undo previous actions automatically — tell me what to change back and I'll do it. E.g. 'put walls back to 1 coat'."}

If the user asks something unclear, return {"actions":[], "reply":"<ask a clarifying question>"}.

Keep replies short — 1-2 sentences. The user can see the items update live below.`;

      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-haiku-4-5-20251001",
          max_tokens: 1500,
          system: systemContext,
          messages: apiHistory,
        }),
      });

      if (!response.ok) {
        const errText = await response.text().catch(() => "");
        throw new Error(`HTTP ${response.status}: ${errText.slice(0, 200)}`);
      }
      const data = await response.json();
      if (data.type === "error" || data.error) {
        throw new Error(data.error?.message || "API error");
      }

      const textBlock = data.content?.find((c) => c.type === "text");
      if (!textBlock) throw new Error("No text in response");

      const cleaned = textBlock.text.replace(/```json|```/g, "").trim();
      let parsed;
      try {
        parsed = JSON.parse(cleaned);
      } catch {
        // Model didn't return valid JSON — treat the whole thing as a reply
        parsed = { actions: [], reply: cleaned.slice(0, 400) };
      }

      const actions = Array.isArray(parsed.actions) ? parsed.actions : [];
      const reply = parsed.reply || "Done.";

      // Apply actions in order. Split into add/remove/patch for the right handler.
      let changedCount = 0;
      if (actions.length > 0) {
        const adds = actions.filter((a) => a._add);
        const removes = actions.filter((a) => a._remove);
        const patches = actions.filter((a) => !a._add && !a._remove);

        if (patches.length > 0) {
          changedCount += applyMultiPatch(patches);
        }
        if (removes.length > 0 && typeof applyMultiPatch === "function") {
          // Convert removes to a special directive the parent understands.
          changedCount += applyMultiPatch(removes);
        }
        if (adds.length > 0 && typeof applyMultiPatch === "function") {
          changedCount += applyMultiPatch(adds);
        }
      }

      // Push the answered exchange and advance the queue. The rooms-changed
      // useEffect will prune any other items the user implicitly resolved.
      if (active) {
        pushAnsweredAndPop(active, userText, reply);
      } else {
        setMessages((prev) => [
          ...prev,
          { role: "user", text: userText },
          { role: "assistant", text: reply },
        ]);
      }
    } catch (e) {
      setError(e.message || "Something went wrong");
      setMessages((prev) => [
        ...prev,
        { role: "user", text: userText },
        { role: "assistant", text: `(Error: ${e.message}. Try again or use a chip above.)` },
      ]);
    } finally {
      setThinking(false);
    }
  }

  // ── Reset the conversation. Re-derives the queue from current rooms +
  //    initialHints so already-applied changes don't re-surface as questions.
  function resetChat() {
    setMessages([]);
    setQueue(buildClarifyQueue(rooms, initialHints));
    setInputText("");
    setError("");
  }

  return (
    <div style={{
      border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))",
      borderRadius: 8,
      background: "var(--color-background-primary, #fff)",
      overflow: "hidden",
      display: "flex",
      flexDirection: "column",
    }}>
      {/* Header */}
      <div style={{
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        padding: "8px 12px",
        background: "var(--color-background-secondary, #f3f3f0)",
        borderBottom: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))",
        fontSize: 11,
        color: "var(--color-text-secondary, #666)",
      }}>
        <span>
          {queue.length > 0
            ? `Clarify — ${queue.length} question${queue.length === 1 ? "" : "s"} remaining`
            : "Clarify — all answered. Items update below as you go."}
        </span>
        <button
          onClick={resetChat}
          style={{
            padding: "2px 8px",
            fontSize: 11,
            borderRadius: 4,
            border: "0.5px solid var(--color-border-secondary, rgba(0,0,0,0.3))",
            background: "transparent",
            cursor: "pointer",
            color: "inherit",
            fontFamily: "inherit",
          }}
        >
          Clear
        </button>
      </div>

      {/* Message log + active question card */}
      <div ref={scrollerRef} style={{ padding: 12, maxHeight: 360, overflowY: "auto", display: "flex", flexDirection: "column", gap: 8 }}>
        {messages.map((m, i) => (
          <ChatBubble key={i} message={m} />
        ))}

        {/* Active question — chips + Skip rendered as the latest assistant turn. */}
        {active && (
          <div>
            <ChatBubble message={{ role: "assistant", text: active.prompt }} />
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 8, marginLeft: 4 }}>
              {active.options.map((opt, j) => (
                <button
                  key={j}
                  onClick={() => handleChip(opt)}
                  disabled={thinking}
                  style={{
                    padding: "6px 12px",
                    fontSize: 12,
                    borderRadius: 14,
                    border: "0.5px solid var(--color-border-secondary, rgba(0,0,0,0.3))",
                    background: "transparent",
                    cursor: thinking ? "not-allowed" : "pointer",
                    color: "inherit",
                    fontFamily: "inherit",
                    opacity: thinking ? 0.5 : 1,
                  }}
                >
                  {opt.label}
                </button>
              ))}
              <button
                onClick={handleSkip}
                disabled={thinking}
                style={{
                  padding: "6px 12px",
                  fontSize: 12,
                  borderRadius: 14,
                  border: "0.5px dashed var(--color-border-secondary, rgba(0,0,0,0.3))",
                  background: "transparent",
                  cursor: thinking ? "not-allowed" : "pointer",
                  color: "var(--color-text-secondary, #666)",
                  fontFamily: "inherit",
                  opacity: thinking ? 0.5 : 1,
                }}
              >
                Skip for now
              </button>
            </div>
          </div>
        )}

        {thinking && (
          <div style={{ alignSelf: "flex-start", fontSize: 12, color: "var(--color-text-secondary, #666)", fontStyle: "italic", padding: "4px 8px" }}>
            Thinking…
          </div>
        )}
      </div>

      {/* Input row */}
      <div style={{ borderTop: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", padding: 12 }}>
        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          <input
            value={inputText}
            onChange={(e) => setInputText(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter" && !thinking) sendMessage(); }}
            disabled={thinking}
            placeholder={active
              ? "Type your answer (or tap a chip above)…"
              : `Say anything — "all walls 2 coats", "remove dining room", "add a closet 6x4"`}
            style={{
              flex: 1,
              padding: "8px 10px",
              fontSize: 13,
              border: "0.5px solid var(--color-border-secondary, rgba(0,0,0,0.3))",
              borderRadius: 6,
              background: "transparent",
              color: "inherit",
              fontFamily: "inherit",
            }}
          />
          <button
            onClick={sendMessage}
            disabled={thinking || !inputText.trim()}
            style={{
              padding: "8px 14px",
              fontSize: 12,
              borderRadius: 6,
              border: "0.5px solid #2C1654",
              background: inputText.trim() && !thinking ? "#2C1654" : "transparent",
              color: inputText.trim() && !thinking ? "#fff" : "var(--color-text-secondary, #888)",
              cursor: thinking || !inputText.trim() ? "not-allowed" : "pointer",
              fontFamily: "inherit",
              fontWeight: 500,
            }}
          >
            Send
          </button>
        </div>
        {error && (
          <div style={{ marginTop: 6, fontSize: 11, color: "var(--color-text-danger, #c33)" }}>
            {error}
          </div>
        )}
      </div>
    </div>
  );
}

// ============================================================================
// BuildChat — Phase 1: in-app budget creation via /api/jobtread (native tool use)
// ============================================================================
// Switched from a brittle JSON-action protocol to Anthropic's native tool use.
// Each user turn runs an internal loop: call /api/chat with messages + tools →
// if stop_reason === "tool_use", dispatch each tool_use block to /api/jobtread,
// return the results as tool_result content blocks, loop up to 8 times. When
// stop_reason === "end_turn", control returns to the user. Tools mirror the
// api/jobtread.js executors 1:1 (find_customer, create_job, …). A small build
// context (customerId, jobId, costGroupIds) is still tracked in a ref and
// serialized into the system prompt as a status summary.

const BUILD_CHAT_TOOLS = [
  {
    name: "find_customer",
    description: "Search JobTread for an existing customer by name fragment. The proxy tokenizes the query and OR-matches each word, so multi-word queries like 'Purple Painting Co' work even when the stored name is just 'Purple Painting'.",
    input_schema: {
      type: "object",
      properties: {
        query: { type: "string", description: "Name fragment to search for." },
      },
      required: ["query"],
    },
  },
  {
    name: "get_customer_jobs",
    description: "List jobs for a given customer.",
    input_schema: {
      type: "object",
      properties: {
        customerId: { type: "string", description: "The customer's JobTread ID." },
      },
      required: ["customerId"],
    },
  },
  {
    name: "find_job",
    description: "Search jobs across the organization by name fragment.",
    input_schema: {
      type: "object",
      properties: {
        query: { type: "string", description: "Name fragment to search for." },
      },
      required: ["query"],
    },
  },
  {
    name: "create_customer",
    description: "Create a new customer account in JobTread plus its primary contact. Returns the created account + contact IDs.",
    input_schema: {
      type: "object",
      properties: {
        name: { type: "string", description: "Customer / account name." },
        contactName: { type: "string", description: "Primary contact's full name. Defaults to the account name if omitted." },
        email: { type: "string", description: "Primary contact email." },
        phone: { type: "string", description: "Primary contact phone number." },
        address: { type: "string", description: "Customer mailing address (the job site address is set separately via create_job)." },
        leadSource: { type: "string", description: "How this lead came in (stored as a custom field on the account)." },
      },
      required: ["name"],
    },
  },
  {
    name: "create_job",
    description: "Create a job under a customer. The Job Type and How Did Bid Come In custom fields populate automatically from tier + bidOrigin.",
    input_schema: {
      type: "object",
      properties: {
        customerId: { type: "string", description: "The customer's JobTread ID." },
        name: { type: "string", description: "Job name." },
        address: { type: "string", description: "Job site address." },
        tier: { type: "string", description: "Pricing tier — 'standard' | 'production' | 'highend' | 'prevailing'. Maps to the Job Type custom field." },
        bidOrigin: { type: "string", description: "How the bid came in — 'Job Walk' | 'Digital Takeoff' | 'Partner Work Order'. Maps to the How Did Bid Come In custom field." },
      },
      required: ["customerId", "name", "address"],
    },
  },
  {
    name: "create_cost_group",
    description: "Create a cost group on a job. Omit parentCostGroupId for top-level groups (Interior, Exterior). Use quantityFormula for derived quantities or quantity for literal numbers.",
    input_schema: {
      type: "object",
      properties: {
        jobId: { type: "string", description: "The job's JobTread ID." },
        name: { type: "string", description: "Cost group name." },
        parentCostGroupId: { type: "string", description: "Parent cost group ID. Omit for top-level groups." },
        quantityFormula: { type: "string", description: "Formula expression (e.g. '2 * (12 + 14) * 9' for room wall SF)." },
        unitId: { type: "string", description: "Unit ID — see system prompt for SF/LF/EA/HR IDs." },
        quantity: { type: "number", description: "Literal numeric quantity if no formula." },
      },
      required: ["jobId", "name"],
    },
  },
  {
    name: "create_cost_item",
    description: "Create a cost item under a cost group. organizationCostItemId links to the org catalog. The cost item carries its OWN dimensional quantityFormula (e.g. \"2 * (12 + 14) * 9\" for walls); the parent subgroup stays empty of quantity so the item's quantity is what job.costItems exposes via the API.",
    input_schema: {
      type: "object",
      properties: {
        jobId: { type: "string", description: "The job's JobTread ID." },
        costGroupId: { type: "string", description: "Parent cost group ID." },
        name: { type: "string", description: "Cost item name (e.g. 'Drywall Walls - Existing - 2-Coats')." },
        organizationCostItemId: { type: "string", description: "Org catalog item ID for the substrate+coats combo." },
        costCodeId: { type: "string", description: "Cost code ID (the 'code' field from the catalog)." },
        costTypeId: { type: "string", description: "Labor or Materials — see system prompt." },
        unitId: { type: "string", description: "Unit ID." },
        quantityFormula: { type: "string", description: "Dimensional formula for this item, e.g. '2 * (12 + 14) * 9' for walls or '12 * 14' for a ceiling. The item carries its own quantity; do not inherit from the parent." },
        quantity: { type: "number", description: "Literal numeric quantity if no formula." },
        unitCost: { type: "number", description: "Per-unit cost (catalog default × tier multiplier)." },
        unitPrice: { type: "number", description: "Per-unit price (catalog default × tier multiplier)." },
      },
      required: ["jobId", "costGroupId", "name"],
    },
  },
  {
    name: "get_catalog_prices",
    description: "Look up the live unitCost and unitPrice for one or more org catalog items by their organizationCostItemId. Batch all IDs you need into a single call. Returns the catalog values — multiply each by the tier multiplier before writing onto a cost item.",
    input_schema: {
      type: "object",
      properties: {
        ids: {
          type: "array",
          items: { type: "string" },
          description: "Array of organizationCostItemIds (e.g. ['22PWiSS293E2', '22PWickkTi46']).",
        },
      },
      required: ["ids"],
    },
  },
];

// Extract user-visible text from a BuildChat API-shape message. Returns null
// when the message has no displayable text (pure tool_use turn or tool_result
// echo) so the chat log skips it.
function buildChatDisplayText(m) {
  if (typeof m.content === "string") return m.content;
  if (!Array.isArray(m.content)) return null;
  const text = m.content
    .filter((b) => b.type === "text")
    .map((b) => b.text)
    .join("\n")
    .trim();
  return text || null;
}

function BuildChat({ payload, onBuilt }) {
  // Each message is in API-compatible shape: { role, content } where content
  // is either a string OR an array of typed blocks (text / tool_use / tool_result).
  // The opening assistant message is display-only — it's sliced off when sending
  // to /api/chat because Anthropic requires the first message to be from the user.
  const [messages, setMessages] = useState([
    {
      role: "assistant",
      content: "Ready to build this in JobTread. What's the customer name and job address? I'll search for an existing customer first — if not found, I'll create one along with the job.",
    },
  ]);
  const [inputText, setInputText] = useState("");
  const [thinking, setThinking] = useState(false);
  const [error, setError] = useState("");
  const scrollerRef = useRef(null);

  // Build context — IDs accumulated as tools execute. Kept in a ref because
  // updates happen mid-turn; the system prompt re-reads the latest values on
  // each callClaude. Note: with native tool use, Claude also sees real tool_result
  // blocks in the conversation, so this is mainly a status convenience for the prompt.
  const contextRef = useRef({
    customerId: null,
    jobId: null,
    costGroupIds: {},      // { groupName: id }
    completed: { rooms: [], scopeBuckets: [], tmItems: [] },
    errors: [],
  });

  useEffect(() => {
    if (scrollerRef.current) {
      scrollerRef.current.scrollTop = scrollerRef.current.scrollHeight;
    }
  }, [messages, thinking]);

  function buildSystemPrompt() {
    return `You are a JobTread build assistant for Purple Painting Co. Your job is to build a budget in JobTread by calling the provided tools.

PRINCIPLES
═══════════════════
- Call tools to do REAL work. NEVER claim an action succeeded without calling the corresponding tool — the tool result is the only ground truth.
- Tools available: find_customer, get_customer_jobs, find_job, create_customer, create_job, create_cost_group, create_cost_item. Each tool's input schema describes its required and optional fields.
- When you need a decision from the user (e.g. picking among multiple search hits, filling in missing details), reply with text and NO tool calls. Otherwise prefer calling tools over describing what you would do.
- Read each tool result before deciding the next step. Tool results are real JobTread API responses — examine them for IDs, errors, and confirmation.
- If a tool returns an error, report it and decide whether to retry, ask the user for clarification, or move on.
- CUSTOMER SEARCH IS MANDATORY. ALWAYS call find_customer before ever suggesting, mentioning, or asking about creating a customer. Never claim a customer "doesn't exist", "wasn't found", or "isn't in JobTread" unless find_customer has actually returned an empty nodes array on THIS turn. Treat ANY name-like input — including a single word or a short partial fragment such as "purple" — as a valid search query and pass it to find_customer as-is. Do NOT dismiss short or fragmentary input as a "voice fragment", typo, or incomplete utterance, and do NOT ask the user to retype it before searching. The customer list is small and names are reasonably unique, so partial matches are expected to work. Only AFTER find_customer returns zero matches may you offer to create a new customer.

IDS YOU'LL NEED
═══════════════════
ORG: 22PWNY9u7qZd
UNITS: SF=22PWNYKLitQU, LF=22PWNYKLfXKr, EA=22PWNYKLcv3n, HR=22PWNYKLfARY
COST TYPES: Labor=22PWNYKLxD4B, Materials=22PWNYKLxwqk
CUSTOM FIELDS:
  Job Type=22PWsEVVW4aj (values: "Residential - Standard Home" | "Residential - Custom House" | "Commercial" | "Prevailing Wage" | "Property Management / Production")
  How Did Bid Come In=22PWsDkAPRYB (values: "Received Digital Bid Invite" | "Requested Job Walk" | "Partner Work Order")
  Lead Source (customer)=22PWNYKhTU6S
  Phone (customerContact)=22PWNYKhWqBE
  Lead Stage (customerContact)=22PWsFuiWzEq

INTERIOR CATALOG (organizationCostItemId per substrate+coats; "code" is the costCodeId for the substrate):
  Walls:     { "1coat":22PWT9nDL9HP, "2coats":22PWiSS293E2, "prime+2":22PWiSgjqkaq, code:22PWmcdKrCnn }
  Ceilings:  { "1coat":22PWickkTi46, "2coats":22PWictZ3V84, "prime+2":22PWicwN2pxL, code:22PWmcdMcqbZ }
  Baseboard: { "1coat":22PWiYQt5Pq8, "2coats":22PWih8XPvPm,                          code:22PWmcdjtJ9C }
  Doors:     { "1coat":22PXFmQz3HEd, "2coats":22PXFmRbHAgn, "prime+2":22PXFmSG2zeV, code:22PWmcdpBtSV }

TIER MULTIPLIERS (apply to BOTH unitCost AND unitPrice — margin % preserved):
  standard=1.00, production=0.85, highend=1.35, prevailing=1.65

CRITICAL JT RULES
═══════════════════
1. Cost item catalog link = organizationCostItemId (NOT sourceCostItemId).
2. Tier multiplier applies to unitCost AND unitPrice. Margin stays constant.
3. Cost items carry their own dimensional quantityFormula (e.g. "2 * (12 + 14) * 9") — do NOT use "{Parent Quantity}". The subgroup above them stays empty of quantity; the API's job.costItems endpoint reads quantity off the item directly.
4. Flat creates only: parent → subgroup → item. Never nest lineItems on createCostGroup.
5. Picklist custom fields can't be empty — only set them when you have a value.
6. Job Type from tier: standard→"Residential - Standard Home", production→"Property Management / Production", highend→"Residential - Custom House", prevailing→"Prevailing Wage".
7. How Did Bid Come In: payload.bidOrigin "Job Walk"→"Requested Job Walk", "Digital Takeoff"→"Received Digital Bid Invite", "Partner Work Order"→"Partner Work Order".

BUILD SEQUENCE
═══════════════════
Stage A — Resolve customer + job:
1. If you don't have the customer name + job address, ask the user.
2. Call find_customer with whatever name fragment the user gave — single words and partial names are fine and expected to work. If find_customer returns exactly one match, confirm with the user and proceed with that customer. If it returns multiple matches, list them and let the user pick. ONLY if find_customer returns zero matches (empty nodes array) may you ask for full details (contact name, email, phone, address) and then call create_customer. Never skip the find_customer step.
3. If an existing customer is chosen, call get_customer_jobs. If a job already matches, use it. Else call create_job.
4. When calling create_job, always include tier (from payload.tier) and bidOrigin (from payload.bidOrigin) so the custom fields populate.

Stage B — Top-level structure (create ONLY the side(s) that will hold work):
5. Create the "Interior" top-level cost group (no parentCostGroupId) ONLY if payload.rooms has at least one room OR payload.scopeBuckets has at least one bucket whose substrate does NOT start with "exterior_". On an exterior-only job (no rooms and no interior buckets), SKIP this step — and skip Stage C entirely since there are no rooms to place.
6. Create the "Exterior" top-level cost group (no parentCostGroupId) ONLY if payload.scopeBuckets has at least one bucket whose substrate starts with "exterior_". On an interior-only job (no exterior buckets — the common case), SKIP this step. Never create an empty top-level group; each side's parentCostGroupId must exist before any Stage C/D items reference it, and if a side has no work, it simply doesn't exist.

Stage C — Rooms (interior). BEFORE creating any cost items in Stages C or D, call get_catalog_prices ONCE with every organizationCostItemId you will use across the whole build (batch them — do NOT call per item). The returned unitCost/unitPrice are the live catalog values; multiply each by the tier multiplier and write the result onto every createCostItem. The catalog is the single source of truth — do NOT hardcode prices, do NOT ask the user for them, and do NOT leave unitCost/unitPrice null (JT shows $0 on the budget if they aren't written on the item itself).

For each room in payload.rooms:
7. Create parent cost group with parentCostGroupId=<Interior.id>.
8. For each enabled substrate, create subgroup + cost item:
   a. Subgroup is a plain organizational container — create it with NO quantityFormula and NO quantity. Its only job is grouping the cost item underneath it.
      Subgroup names: "Drywall Walls", "Drywall Ceilings", "Wood Baseboard", "Doors+Frames".
      Unit IDs (for display on the subgroup): walls/ceiling=SF, baseboard=LF, doors=EA.
   b. Cost item under that subgroup carries the dimensional formula ITSELF — do NOT use "{Parent Quantity}" inheritance. This is what makes the item's quantity readable via the job.costItems API.
      - organizationCostItemId = catalog entry by substrate+coats
      - name = "<Substrate label> - Existing - <Coats label>" (e.g. "Drywall Walls - Existing - 2-Coats")
      - costCodeId = catalog "code" for that substrate
      - costTypeId = Labor
      - unitId = matching unit
      - quantityFormula / quantity (substitute the room's actual L, W, H, doorCount as numbers):
          walls cost item:     quantityFormula = "2 * (<L> + <W>) * <H>"
          ceiling cost item:   quantityFormula = "<L> * <W>"
          baseboard cost item: quantityFormula = "2 * (<L> + <W>)"
          doors cost item:     quantity = <doorCount>  (literal EA, no formula)
      - unitCost  = (live catalog unitCost for this organizationCostItemId, from get_catalog_prices) × tier multiplier
      - unitPrice = (live catalog unitPrice for this organizationCostItemId, from get_catalog_prices) × tier multiplier

Stage D — Scope buckets:
For each item in payload.scopeBuckets:
9. parent = Exterior if substrate starts with "exterior_", else Interior.
10. Create cost group with literal quantity + unitId by item.unit.
11. Create cost item. If substrate NOT in interior catalog above, ASK the user for organizationCostItemId + costCodeId. Set unitCost / unitPrice the same way as Stage C — call get_catalog_prices and multiply by the tier multiplier. Do NOT leave unitCost/unitPrice null.

Stage E — T&M items:
For each item in payload.tmItems:
12. payload.tmItems already include catalogId, costCode, unitCost, unitPrice — use directly.
13. T&M items belong under their cost-code group from payload.tmCatalog (find/create the group as needed). Don't nest under Interior or Exterior.

Stage F — Summary:
14. After all built, reply with a clean markdown summary in EXACTLY this structure. Do NOT use pipe "|" table syntax. Do NOT add extra prose or rehash the build narrative. Use the heading, the labeled lines, and one bullet per cost item on its own line. Substitute the bracketed placeholders with real values from payload + the build context (use payload.tier.label and payload.tier.multiplier for the Tier line; use the actual jobId in the link):

## ✅ Build Complete

**Customer:** <customer name>
**Job:** <job name>
**Tier:** <tier label> (<multiplier>x)

**Cost Groups (<count>):** <comma-separated group names>

**Cost Items (<count>):**
- <item name> — <quantity> <unit> @ <coats>
- <item name> — <quantity> <unit> @ <coats>
(one bullet per cost item, each on its own line)

**[Open Job in JobTread](https://app.jobtread.com/jobs/<jobId>)**

CURRENT BUILD CONTEXT
═══════════════════
${JSON.stringify({
  customerId: contextRef.current.customerId,
  jobId: contextRef.current.jobId,
  costGroupIds: contextRef.current.costGroupIds,
  completed: contextRef.current.completed,
  recentErrors: contextRef.current.errors.slice(-3),
}, null, 2)}

PAYLOAD (what you're building)
═══════════════════
${JSON.stringify(payload, null, 2)}`;
  }

  // Execute one tool call via /api/jobtread. Returns the raw JT response (or
  // an error object). Side effect: updates contextRef with any returned IDs.
  async function executeAction(name, input) {
    let result;
    try {
      const resp = await fetch("/api/jobtread", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: name, payload: input }),
      });
      result = await resp.json();
      if (!resp.ok) {
        contextRef.current.errors.push({ action: name, status: resp.status, error: result });
      }
    } catch (e) {
      result = { error: "fetch_failed", message: e.message };
      contextRef.current.errors.push({ action: name, error: e.message });
    }

    const ctx = contextRef.current;
    if (name === "create_customer") {
      const id = result?.createAccount?.createdAccount?.id;
      if (id) ctx.customerId = id;
    } else if (name === "create_job") {
      const id = result?.createJob?.createdJob?.id;
      if (id) ctx.jobId = id;
    } else if (name === "create_cost_group") {
      const id = result?.createCostGroup?.createdCostGroup?.id;
      const groupName = result?.createCostGroup?.createdCostGroup?.name || input?.name;
      if (id && groupName) ctx.costGroupIds[groupName] = id;
    }

    return result;
  }

  // Strip the display-only opening assistant message (Anthropic requires the
  // first message to be from the user).
  function toApiMessages(history) {
    return history.slice(1).map((m) => ({ role: m.role, content: m.content }));
  }

  async function callClaude(history) {
    const resp = await fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 4000,
        system: buildSystemPrompt(),
        tools: BUILD_CHAT_TOOLS,
        messages: toApiMessages(history),
      }),
    });
    if (!resp.ok) {
      const errText = await resp.text().catch(() => "");
      throw new Error(`/api/chat HTTP ${resp.status}: ${errText.slice(0, 200)}`);
    }
    const data = await resp.json();
    if (data.type === "error" || data.error) {
      throw new Error(data.error?.message || "API error");
    }
    return data;  // full response: content[], stop_reason, ...
  }

  async function sendMessage() {
    if (!inputText.trim() || thinking) return;
    const userText = inputText.trim();
    setInputText("");
    setError("");

    let history = [...messages, { role: "user", content: userText }];
    setMessages(history);
    setThinking(true);

    const MAX_STEPS = 20;
    try {
      for (let step = 0; step < MAX_STEPS; step++) {
        const data = await callClaude(history);
        // Echo the assistant turn verbatim — Claude needs to see its own previous
        // content (text + tool_use blocks) on the next turn for tool_use_id linkage.
        history = [...history, { role: "assistant", content: data.content }];
        setMessages(history);

        if (data.stop_reason !== "tool_use") {
          // end_turn / max_tokens / stop_sequence — yield to the user.
          // Lift a snapshot of what's been built so the parent can gate Scope.
          if (
            typeof onBuilt === "function" &&
            contextRef.current.jobId &&
            Object.keys(contextRef.current.costGroupIds).length > 0
          ) {
            onBuilt({
              ...contextRef.current,
              costGroupIds: { ...contextRef.current.costGroupIds },
            });
          }
          return;
        }

        // Execute every tool_use block from this turn. All tool_result blocks
        // go back together as ONE user turn — Anthropic groups them like this.
        const toolUseBlocks = (data.content || []).filter((b) => b.type === "tool_use");
        const toolResultBlocks = [];
        for (const block of toolUseBlocks) {
          // eslint-disable-next-line no-await-in-loop
          const result = await executeAction(block.name, block.input);
          toolResultBlocks.push({
            type: "tool_result",
            tool_use_id: block.id,
            content: JSON.stringify(result),
          });
        }
        history = [...history, { role: "user", content: toolResultBlocks }];
        setMessages(history);
      }
      setMessages((prev) => [...prev, { role: "assistant", content: "Hit max steps — pausing. What would you like to do?" }]);
    } catch (e) {
      setError(e.message || "Something went wrong");
      setMessages((prev) => [...prev, { role: "assistant", content: `(Error: ${e.message})` }]);
    } finally {
      setThinking(false);
    }
  }

  return (
    <div style={{
      border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))",
      borderRadius: 8,
      background: "var(--color-background-primary, #fff)",
      overflow: "hidden",
      display: "flex",
      flexDirection: "column",
    }}>
      {/* Header */}
      <div style={{
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        padding: "8px 12px",
        background: "var(--color-background-secondary, #f3f3f0)",
        borderBottom: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))",
        fontSize: 11,
        color: "var(--color-text-secondary, #666)",
      }}>
        <span>Building in JobTread</span>
      </div>

      {/* Message log — tool_use / tool_result blocks render no bubble; only text shows. */}
      <div ref={scrollerRef} style={{ padding: 12, maxHeight: 420, overflowY: "auto", display: "flex", flexDirection: "column", gap: 8 }}>
        {messages.map((m, i) => {
          const text = buildChatDisplayText(m);
          if (!text) return null;
          return <ChatBubble key={i} message={{ role: m.role, text }} />;
        })}
        {thinking && (
          <div style={{ alignSelf: "flex-start", fontSize: 12, color: "var(--color-text-secondary, #666)", fontStyle: "italic", padding: "4px 8px" }}>
            Thinking…
          </div>
        )}
      </div>

      {/* Input row */}
      <div style={{ borderTop: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", padding: 12 }}>
        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          <input
            value={inputText}
            onChange={(e) => setInputText(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter" && !thinking) sendMessage(); }}
            disabled={thinking}
            placeholder='e.g. "Customer is Smith Family, job at 123 Oak Street"'
            style={{
              flex: 1,
              padding: "8px 10px",
              fontSize: 13,
              border: "0.5px solid var(--color-border-secondary, rgba(0,0,0,0.3))",
              borderRadius: 6,
              background: "transparent",
              color: "inherit",
              fontFamily: "inherit",
            }}
          />
          <button
            onClick={sendMessage}
            disabled={thinking || !inputText.trim()}
            style={{
              padding: "8px 14px",
              fontSize: 12,
              borderRadius: 6,
              border: "0.5px solid #2C1654",
              background: inputText.trim() && !thinking ? "#2C1654" : "transparent",
              color: inputText.trim() && !thinking ? "#fff" : "var(--color-text-secondary, #888)",
              cursor: thinking || !inputText.trim() ? "not-allowed" : "pointer",
              fontFamily: "inherit",
              fontWeight: 500,
            }}
          >
            Send
          </button>
        </div>
        {error && (
          <div style={{ marginTop: 6, fontSize: 11, color: "var(--color-text-danger, #c33)" }}>
            {error}
          </div>
        )}
      </div>
    </div>
  );
}

// ============================================================================
// DocumentChat — Phase 2: in-app proposal document via /api/jobtread (native tool use)
// ============================================================================
// Clone of BuildChat's pattern (same /api/chat loop, contextRef + system-prompt
// status convenience, same /api/jobtread tool dispatch) but with two
// document-shaped executors. Fed by the in-memory builtBudget snapshot
// (jobId + costGroupIds) so the model can mirror what BuildChat created.

const DOCUMENT_CHAT_TOOLS = [
  {
    name: "find_user",
    description: "Look up an internal Purple Painting team member by name fragment. Returns matching users with { name, emailAddress, phoneNumber } — use these as fromName / fromEmailAddress / fromPhoneNumber on create_document. The proxy filters to @purplepainting.net users and excludes machine accounts.",
    input_schema: {
      type: "object",
      properties: {
        name: { type: "string", description: "Name fragment to match (e.g. 'Kareem', 'Peter')." },
      },
      required: ["name"],
    },
  },
  {
    name: "get_job_cost_items",
    description: "List the JOB's cost items AND the job's location in one call. Returns { job: { location: { name, address }, costItems: { nodes: [{ id, name, quantity, costCode { id }, costType { id }, unit { id }, unitCost, unitPrice, costGroup { id, name } }, ...] } } }. Each costItems.nodes[].id is the jobCostItemId you pass to update_document's costItem children. job.location.name and job.location.address are what you pass to create_document as jobLocationName / jobLocationAddress (createDocument does NOT inherit the job's location).",
    input_schema: {
      type: "object",
      properties: {
        jobId: { type: "string", description: "The job's JobTread ID (from builtBudget.jobId)." },
      },
      required: ["jobId"],
    },
  },
  {
    name: "create_document",
    description: "Create a JobTread document. Barebones — line items are added afterwards via update_document. `name` MUST be one of JT's TEMPLATE names verbatim ('Proposal' | 'Selections' | 'Change Order' | 'Interior Scope of work proposal'); the descriptive title goes in `subject`. createDocument does NOT inherit the job's location — always pass jobLocationName + jobLocationAddress from get_job_cost_items' job.location. type is 'customerOrder' for proposals. dueDays XOR dueDate — never set both.",
    input_schema: {
      type: "object",
      properties: {
        jobId: { type: "string", description: "The job's JobTread ID (from builtBudget.jobId)." },
        name: { type: "string", description: "Document TEMPLATE name (must match a JT template verbatim): 'Proposal' | 'Selections' | 'Change Order' | 'Interior Scope of work proposal'. Default: 'Proposal'. Do NOT put the customer/address here — JT rejects non-template names." },
        subject: { type: "string", description: "Descriptive document title shown to the customer (e.g. 'Smith Family — 123 Main St — Interior Repaint Proposal'). This is where the human-readable label goes, not `name`." },
        type: { type: "string", description: "Document type enum: 'bidRequest' | 'customerInvoice' | 'customerOrder' | 'vendorBill' | 'vendorOrder'. Use 'customerOrder' for proposals." },
        fromName: { type: "string", description: "Preparer's display name (the result of find_user)." },
        toName: { type: "string", description: "Recipient's display name — typically the customer's primary contact." },
        taxRate: { type: "number", description: "Sales tax rate as a decimal (0..1). Use 0 for paint proposals." },
        description: { type: "string", description: "Scope-of-work text (≤32768 chars). Goes into the document body." },
        footer: { type: "string", description: "Exclusions text (≤65536 chars). Goes at the bottom of the document." },
        dueDays: { type: "number", description: "Days until the document is due (integer ≥ 0). Use 30 by default. XOR with dueDate — pick one." },
        dueDate: { type: "string", description: "Absolute due date in ISO format YYYY-MM-DD. XOR with dueDays — pick one." },
        issueDate: { type: "string", description: "Date the document is issued, ISO YYYY-MM-DD. Default to today." },
        jobLocationName: { type: "string", description: "Document location name — pass job.location.name from get_job_cost_items." },
        jobLocationAddress: { type: "string", description: "Document location address — pass job.location.address from get_job_cost_items." },
        fromEmailAddress: { type: "string", description: "Preparer's email (from find_user)." },
        fromPhoneNumber: { type: "string", description: "Preparer's phone (from find_user, if present)." },
        toEmailAddress: { type: "string", description: "Recipient's email — from the customer's primaryContact." },
      },
      required: ["jobId", "name", "type", "fromName", "toName", "taxRate"],
    },
  },
  {
    name: "update_document",
    description: "Update a JobTread document. lineItems use the PROVEN costGroup-wrapper structure (verified against a live $8,219 document): one or more top-level costGroup entries, each carrying costItem children with the full field set { _type:'costItem', jobCostItemId, name, quantity, costCodeId, costTypeId, unitId, unitCost, unitPrice }. Discriminator is 'costGroup' / 'costItem' (NOT 'newCostGroup' / 'newCostItem'). Pull costCodeId/costTypeId/unitId/unitCost/unitPrice from the matching get_job_cost_items node's costCode.id / costType.id / unit.id / unitCost / unitPrice — leaving any of them out makes JT show $0 or wrong codes.",
    input_schema: {
      type: "object",
      properties: {
        id: { type: "string", description: "The document's JobTread ID (from create_document's createdDocument.id)." },
        name: { type: "string", description: "Updated document TEMPLATE name (same constraint as create_document — must match a JT template verbatim)." },
        subject: { type: "string", description: "Updated descriptive subject." },
        fromName: { type: "string", description: "Updated preparer name." },
        toName: { type: "string", description: "Updated recipient name." },
        taxRate: { type: "number", description: "Updated tax rate (0..1)." },
        dueDays: { type: "number", description: "Updated days until due. XOR with dueDate." },
        dueDate: { type: "string", description: "Updated absolute due date (YYYY-MM-DD). XOR with dueDays." },
        issueDate: { type: "string", description: "Updated issue date (YYYY-MM-DD)." },
        jobLocationName: { type: "string", description: "Updated document location name." },
        jobLocationAddress: { type: "string", description: "Updated document location address." },
        description: { type: "string", description: "Updated scope-of-work text." },
        footer: { type: "string", description: "Updated exclusions text." },
        fromEmailAddress: { type: "string", description: "Updated preparer email." },
        fromPhoneNumber: { type: "string", description: "Updated preparer phone." },
        toEmailAddress: { type: "string", description: "Updated recipient email." },
        lineItems: {
          type: "array",
          description: "PROVEN costGroup-wrapper structure. Each entry is a costGroup containing costItem children — see this tool's top-level description for the exact field set.",
          items: {
            type: "object",
            description: "A costGroup wrapper. _type must be 'costGroup'.",
            properties: {
              _type: { type: "string", description: "Discriminator. Must be 'costGroup'." },
              name: { type: "string", description: "Group label shown on the document (e.g. 'Full Interior Painting')." },
              lineItems: {
                type: "array",
                description: "costItem children inside this group. _type on each must be 'costItem'.",
                items: {
                  type: "object",
                  properties: {
                    _type: { type: "string", description: "Discriminator. Must be 'costItem'." },
                    jobCostItemId: { type: "string", description: "The matching get_job_cost_items node's id. NOT sourceCostItemId, NOT organizationCostItemId." },
                    name: { type: "string", description: "Line display name (typically the cost item's name)." },
                    quantity: { type: "number", description: "Resolved numeric quantity from get_job_cost_items (e.g. 396, 120). NOT a formula string." },
                    costCodeId: { type: "string", description: "From the get_job_cost_items node's costCode.id." },
                    costTypeId: { type: "string", description: "From the get_job_cost_items node's costType.id." },
                    unitId: { type: "string", description: "From the get_job_cost_items node's unit.id." },
                    unitCost: { type: "number", description: "From the get_job_cost_items node's unitCost." },
                    unitPrice: { type: "number", description: "From the get_job_cost_items node's unitPrice." },
                  },
                  required: ["_type", "jobCostItemId"],
                },
              },
            },
            required: ["_type", "name", "lineItems"],
          },
        },
      },
      required: ["id"],
    },
  },
];

function DocumentChat({ payload, builtBudget }) {
  // Same API-compatible shape as BuildChat — { role, content } where content
  // is a string OR array of typed blocks (text / tool_use / tool_result).
  const [messages, setMessages] = useState([
    {
      role: "assistant",
      content: "Ready to create the customer-facing proposal document. Who's preparing this one? Default is Kareem — say a different name if it's someone else (Peter, Lyric, etc.).",
    },
  ]);
  const [inputText, setInputText] = useState("");
  const [thinking, setThinking] = useState(false);
  const [error, setError] = useState("");
  const scrollerRef = useRef(null);

  // Document context — IDs accumulated as tools execute. Kept in a ref so the
  // system prompt re-reads the latest values on each callClaude.
  const contextRef = useRef({
    documentId: null,
    preparer: null,        // { name, emailAddress, phoneNumber } — from find_user
    recipient: null,       // { name, emailAddress } — typically customer primaryContact
    errors: [],
  });

  useEffect(() => {
    if (scrollerRef.current) {
      scrollerRef.current.scrollTop = scrollerRef.current.scrollHeight;
    }
  }, [messages, thinking]);

  function buildSystemPrompt() {
    return `You are a JobTread document assistant for Purple Painting Co. Your job is to create and edit the customer-facing proposal document by calling the provided tools.

PRINCIPLES
═══════════════════
- Call tools to do REAL work. NEVER claim an action succeeded without calling the corresponding tool — the tool result is the only ground truth.
- Tools available: find_user, get_job_cost_items, create_document, update_document. Each tool's input schema describes its required and optional fields.
- When you need a decision from the user (preparer confirmation, recipient confirmation, conversational edits), reply with text and NO tool calls. Otherwise prefer calling tools.
- Read each tool result before deciding the next step. Tool results are real JobTread API responses — examine them for IDs, errors, and confirmation.
- If a tool returns an error, report it and decide whether to retry, ask the user for clarification, or move on.

IDS YOU'LL NEED
═══════════════════
ORG: 22PWNY9u7qZd
JOB: ${builtBudget?.jobId || "(not built yet — ask the user for the job ID)"}
CUSTOMER: ${builtBudget?.customerId || "(unknown — ask the user)"}
TYPE: use "customerOrder" for customer-facing proposals.

CRITICAL DOCUMENT RULES
═══════════════════
1. update_document.lineItems uses the PROVEN costGroup-wrapper structure (verified against a live $8,219 document). One or more top-level entries with { _type:"costGroup", name:<group label>, lineItems:[<costItem children>] }. Each costItem child MUST include { _type:"costItem", jobCostItemId, name, quantity, costCodeId, costTypeId, unitId, unitCost, unitPrice } — pull costCodeId/costTypeId/unitId/unitCost/unitPrice from the matching get_job_cost_items node. Discriminator is "costGroup"/"costItem" (NOT "newCostGroup"/"newCostItem"). jobCostItemId is the node's id — NOT sourceCostItemId, NOT organizationCostItemId.
2. create_document.name MUST exactly match a JT TEMPLATE name: "Proposal" | "Selections" | "Change Order" | "Interior Scope of work proposal". Default "Proposal". The descriptive title (e.g. "Smith Family — 123 Main St — Interior Repaint Proposal") goes in 'subject', NEVER in 'name' — JT rejects non-template names.
3. createDocument does NOT inherit the job's location. ALWAYS pass jobLocationName + jobLocationAddress (read from get_job_cost_items' job.location.name / job.location.address).
4. dueDays XOR dueDate — never set both. dueDays defaults to 30; use dueDate (YYYY-MM-DD) when an absolute date is needed.
5. taxRate defaults to 0 (Purple Painting proposals are typically tax-exempt at this stage). It is a decimal — 0 means no tax, 0.0825 = 8.25%.
6. Document type is "customerOrder" for proposals (the enum also has bidRequest, customerInvoice, vendorBill, vendorOrder — don't use those here).
7. description = scope of work (compose from payload.scope). footer = exclusions text (payload.exclusions).
8. create_document REQUIRES: jobId, name, type, fromName, toName, taxRate. All required — don't omit any.

BUILD SEQUENCE
═══════════════════
Stage A — Identify the preparer:
1. The user's opening reply will name the preparer (default Kareem if unspecified — always confirm). Call find_user with that name. The proxy filters to @purplepainting.net non-machine users and returns matches with { name, emailAddress, phoneNumber }. If multiple match, ask the user to pick. Once chosen, use those exact values as fromName / fromEmailAddress / fromPhoneNumber.

Stage B — Confirm the recipient:
2. Default recipient = the customer's primary contact (from payload.customer or payload.recipient if present, otherwise ask the user for name + email). Confirm with the user, then use toName / toEmailAddress.

Stage C — Fetch job context, then create the barebones document:
3. Call get_job_cost_items with jobId FIRST. This returns BOTH the job's location and all cost-item nodes you'll need in Stage D. Remember job.location.name and job.location.address for the next step.
4. Call create_document with ALL required fields PLUS:
   - name: exactly "Proposal" (or another template name verbatim — see Rule 2). Do NOT put the customer/address here.
   - subject: the descriptive title (e.g. "Smith Family — 123 Main St — Interior Repaint Proposal").
   - type: "customerOrder"
   - fromName, fromEmailAddress, fromPhoneNumber (from find_user)
   - toName, toEmailAddress (recipient)
   - taxRate: 0
   - dueDays: 30 (or dueDate if absolute — never both)
   - issueDate: today (YYYY-MM-DD) when known
   - jobLocationName: job.location.name from step 3
   - jobLocationAddress: job.location.address from step 3
   - description: composed scope-of-work text
   - footer: composed exclusions text
   Capture createdDocument.id from the result.

Stage D — Mirror the budget as line items (PROVEN structure):
5. From the get_job_cost_items result you already have, build ONE costGroup wrapper containing every cost item as a costItem child, then call update_document with this lineItems array:
     lineItems: [
       {
         _type: "costGroup",
         name: "Full Interior Painting",
         lineItems: [
           {
             _type: "costItem",
             jobCostItemId: <node.id>,
             name: <node.name>,
             quantity: <node.quantity>,        // REAL number from the node (e.g. 396, 120) — not a formula
             costCodeId: <node.costCode.id>,
             costTypeId: <node.costType.id>,
             unitId: <node.unit.id>,
             unitCost: <node.unitCost>,
             unitPrice: <node.unitPrice>,
           },
           ... one per cost item ...
         ],
       },
     ]
   Every costItem child MUST carry the full field set above — omitting costCodeId/costTypeId/unitId/unitCost/unitPrice is what causes JT to show $0 or wrong codes on the document. The discriminators are exactly "costGroup" / "costItem", not "newCostGroup" / "newCostItem".

Stage E — Conversational edits:
5. Honor user-driven edits via update_document. Examples:
   - "change the footer" → update_document { id, footer: "<new text>" }
   - "drop that line" → update_document { id, lineItems: [<reduced list>] }
   - "different preparer — Peter" → re-run find_user("Peter"), then update_document { id, fromName, fromEmailAddress, fromPhoneNumber }

Stage F — Summary:
6. After the document is created and any edits applied, reply with EXACTLY this structure (no pipe "|" table syntax, no rehash):

## ✅ Document Ready

**Job:** <job name>
**Document:** <document name>
**Preparer:** <preparer name>
**Recipient:** <recipient name>

**Line Items (<count>):**
- <item name> — <quantity> @ <unitPrice>
- <item name> — <quantity> @ <unitPrice>
(one bullet per line item)

**[Open document in JobTread](https://app.jobtread.com/documents/<documentId>)**

CURRENT DOCUMENT CONTEXT
═══════════════════
${JSON.stringify({
  documentId: contextRef.current.documentId,
  preparer: contextRef.current.preparer,
  recipient: contextRef.current.recipient,
  recentErrors: contextRef.current.errors.slice(-3),
}, null, 2)}

BUILT BUDGET (from BuildChat)
═══════════════════
${JSON.stringify(builtBudget || null, null, 2)}

PAYLOAD (scope + exclusions + tier + line items)
═══════════════════
${JSON.stringify(payload, null, 2)}`;
  }

  async function executeAction(name, input) {
    let result;
    try {
      const resp = await fetch("/api/jobtread", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: name, payload: input }),
      });
      result = await resp.json();
      if (!resp.ok) {
        contextRef.current.errors.push({ action: name, status: resp.status, error: result });
      }
    } catch (e) {
      result = { error: "fetch_failed", message: e.message };
      contextRef.current.errors.push({ action: name, error: e.message });
    }

    const ctx = contextRef.current;
    if (name === "create_document") {
      const id = result?.createDocument?.createdDocument?.id;
      if (id) ctx.documentId = id;
    } else if (name === "find_user") {
      // Single unambiguous match → remember as the current preparer so the
      // system prompt's status block carries it forward. Multiple matches
      // wait for the model to pick before populating context.
      const matches = Array.isArray(result?.matches) ? result.matches : [];
      if (matches.length === 1) ctx.preparer = matches[0];
    }

    return result;
  }

  function toApiMessages(history) {
    return history.slice(1).map((m) => ({ role: m.role, content: m.content }));
  }

  async function callClaude(history) {
    const resp = await fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 4000,
        system: buildSystemPrompt(),
        tools: DOCUMENT_CHAT_TOOLS,
        messages: toApiMessages(history),
      }),
    });
    if (!resp.ok) {
      const errText = await resp.text().catch(() => "");
      throw new Error(`/api/chat HTTP ${resp.status}: ${errText.slice(0, 200)}`);
    }
    const data = await resp.json();
    if (data.type === "error" || data.error) {
      throw new Error(data.error?.message || "API error");
    }
    return data;
  }

  async function sendMessage() {
    if (!inputText.trim() || thinking) return;
    const userText = inputText.trim();
    setInputText("");
    setError("");

    let history = [...messages, { role: "user", content: userText }];
    setMessages(history);
    setThinking(true);

    const MAX_STEPS = 20;
    try {
      for (let step = 0; step < MAX_STEPS; step++) {
        const data = await callClaude(history);
        history = [...history, { role: "assistant", content: data.content }];
        setMessages(history);

        if (data.stop_reason !== "tool_use") {
          return;
        }

        const toolUseBlocks = (data.content || []).filter((b) => b.type === "tool_use");
        const toolResultBlocks = [];
        for (const block of toolUseBlocks) {
          // eslint-disable-next-line no-await-in-loop
          const result = await executeAction(block.name, block.input);
          toolResultBlocks.push({
            type: "tool_result",
            tool_use_id: block.id,
            content: JSON.stringify(result),
          });
        }
        history = [...history, { role: "user", content: toolResultBlocks }];
        setMessages(history);
      }
      setMessages((prev) => [...prev, { role: "assistant", content: "Hit max steps — pausing. What would you like to do?" }]);
    } catch (e) {
      setError(e.message || "Something went wrong");
      setMessages((prev) => [...prev, { role: "assistant", content: `(Error: ${e.message})` }]);
    } finally {
      setThinking(false);
    }
  }

  return (
    <div style={{
      border: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))",
      borderRadius: 8,
      background: "var(--color-background-primary, #fff)",
      overflow: "hidden",
      display: "flex",
      flexDirection: "column",
    }}>
      <div style={{
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        padding: "8px 12px",
        background: "var(--color-background-secondary, #f3f3f0)",
        borderBottom: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))",
        fontSize: 11,
        color: "var(--color-text-secondary, #666)",
      }}>
        <span>Creating document in JobTread</span>
      </div>

      <div ref={scrollerRef} style={{ padding: 12, maxHeight: 420, overflowY: "auto", display: "flex", flexDirection: "column", gap: 8 }}>
        {messages.map((m, i) => {
          const text = buildChatDisplayText(m);
          if (!text) return null;
          return <ChatBubble key={i} message={{ role: m.role, text }} />;
        })}
        {thinking && (
          <div style={{ alignSelf: "flex-start", fontSize: 12, color: "var(--color-text-secondary, #666)", fontStyle: "italic", padding: "4px 8px" }}>
            Thinking…
          </div>
        )}
      </div>

      <div style={{ borderTop: "0.5px solid var(--color-border-tertiary, rgba(0,0,0,0.15))", padding: 12 }}>
        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          <input
            value={inputText}
            onChange={(e) => setInputText(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter" && !thinking) sendMessage(); }}
            disabled={thinking}
            placeholder='e.g. "Peter is preparing, send to the primary contact"'
            style={{
              flex: 1,
              padding: "8px 10px",
              fontSize: 13,
              border: "0.5px solid var(--color-border-secondary, rgba(0,0,0,0.3))",
              borderRadius: 6,
              background: "transparent",
              color: "inherit",
              fontFamily: "inherit",
            }}
          />
          <button
            onClick={sendMessage}
            disabled={thinking || !inputText.trim()}
            style={{
              padding: "8px 14px",
              fontSize: 12,
              borderRadius: 6,
              border: "0.5px solid #2C1654",
              background: inputText.trim() && !thinking ? "#2C1654" : "transparent",
              color: inputText.trim() && !thinking ? "#fff" : "var(--color-text-secondary, #888)",
              cursor: thinking || !inputText.trim() ? "not-allowed" : "pointer",
              fontFamily: "inherit",
              fontWeight: 500,
            }}
          >
            Send
          </button>
        </div>
        {error && (
          <div style={{ marginTop: 6, fontSize: 11, color: "var(--color-text-danger, #c33)" }}>
            {error}
          </div>
        )}
      </div>
    </div>
  );
}

// Compact one-line representation of a line item for prompt context.
function formatItemDigest(r, idx) {
  if (r.type === "tm") return `[${idx}] T&M "${r.name}" ${r.hours || 0}hr (${r.category})`;
  if (r.type === "scope") return `[${idx}] scope "${r.name}" — ${r.substrate} ${r.quantity}${r.unit} ${r.coats || "2coats"}`;
  // room
  const subs = [];
  if (r.walls?.enabled !== false) subs.push(`walls:${r.walls?.coats || "2coats"}`);
  if (r.ceiling?.enabled !== false) subs.push(`ceiling:${r.ceiling?.coats || "2coats"}`);
  if (r.baseboard?.enabled !== false) subs.push(`baseboard:${r.baseboard?.coats || "2coats"}`);
  if (r.doors?.enabled !== false) subs.push(`doors:${r.doors?.coats || "2coats"}x${r.doors?.count || 0}`);
  return `[${idx}] room "${r.name}" ${r.length || "?"}x${r.width || "?"}x${r.height || 9} (${subs.join(", ")})`;
}

function ManualCopyModal({ text, label, onClose }) {
  const taRef = useRef(null);

  // Auto-select the text on mount so the user can just Cmd+C / Ctrl+C
  useEffect(() => {
    if (taRef.current) {
      taRef.current.focus();
      taRef.current.select();
    }
  }, []);

  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(0,0,0,0.5)",
        zIndex: 9999,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 20,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: "var(--color-background-primary, #fff)",
          color: "var(--color-text-primary, #222)",
          borderRadius: 12,
          padding: 20,
          maxWidth: 720,
          width: "100%",
          maxHeight: "90vh",
          display: "flex",
          flexDirection: "column",
          gap: 12,
          boxShadow: "0 10px 40px rgba(0,0,0,0.3)",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <h3 style={{ margin: 0, fontSize: 16, fontWeight: 500 }}>Copy {label} manually</h3>
          <button
            onClick={onClose}
            style={{ padding: "4px 10px", fontSize: 12, border: "0.5px solid rgba(0,0,0,0.3)", borderRadius: 6, background: "transparent", cursor: "pointer", color: "inherit" }}
          >
            Close
          </button>
        </div>
        <p style={{ fontSize: 12, color: "var(--color-text-secondary, #666)", margin: 0 }}>
          The clipboard isn't accessible from inside this artifact's iframe — common in sandboxed environments. The text is selected below — press <strong>Cmd+C</strong> (Mac) or <strong>Ctrl+C</strong> (Windows) to copy it.
        </p>
        <textarea
          ref={taRef}
          value={text}
          readOnly
          style={{
            flex: 1,
            width: "100%",
            minHeight: 300,
            padding: 12,
            fontSize: 11,
            fontFamily: "var(--font-mono, monospace)",
            border: "0.5px solid rgba(0,0,0,0.2)",
            borderRadius: 6,
            background: "var(--color-background-secondary, #f3f3f0)",
            color: "inherit",
            resize: "vertical",
          }}
        />
        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
          <button
            onClick={() => {
              if (taRef.current) {
                taRef.current.focus();
                taRef.current.select();
              }
            }}
            style={{ padding: "8px 14px", fontSize: 13, border: "0.5px solid rgba(0,0,0,0.3)", borderRadius: 6, background: "transparent", cursor: "pointer", color: "inherit", fontFamily: "inherit" }}
          >
            Re-select text
          </button>
          <button
            onClick={onClose}
            style={{ padding: "8px 14px", fontSize: 13, border: "0.5px solid #2C1654", borderRadius: 6, background: "#2C1654", color: "#fff", cursor: "pointer", fontFamily: "inherit", fontWeight: 500 }}
          >
            Done
          </button>
        </div>
      </div>
    </div>
  );
}

function ChatBubble({ message }) {
  if (message.role === "system") {
    return (
      <div style={{
        alignSelf: "center",
        fontSize: 11,
        color: "var(--color-text-secondary, #888)",
        fontStyle: "italic",
        padding: "2px 8px",
      }}>
        {message.text}
      </div>
    );
  }
  const isUser = message.role === "user";
  return (
    <div style={{
      alignSelf: isUser ? "flex-end" : "flex-start",
      maxWidth: "85%",
      background: isUser ? "#2C1654" : "var(--color-background-secondary, #f3f3f0)",
      color: isUser ? "#fff" : "inherit",
      padding: "8px 12px",
      borderRadius: 12,
      borderTopRightRadius: isUser ? 2 : 12,
      borderTopLeftRadius: isUser ? 12 : 2,
      fontSize: 13,
      lineHeight: 1.4,
    }}>
      {message.text}
      {message.appliesTo && !isUser && (
        <div style={{ fontSize: 11, opacity: 0.7, marginTop: 2 }}>
          re: {message.appliesTo}
        </div>
      )}
    </div>
  );
}

// Parse a freeform custom answer and turn it into a structured patch.
// e.g. "22 LF" → { quantity: 22, unit: "LF" }
//      "2 coats" → { coats: "2coats" }
//      "20 × 14 × 9" → { length: 20, width: 14, height: 9 }
function parseCustomAnswer(field, text) {
  const lower = text.toLowerCase();
  const patch = {};

  // Quantity + unit pattern: "22 LF", "150 SF", "4 EA"
  const qtyMatch = lower.match(/(\d+(?:\.\d+)?)\s*(lf|sf|ea|hr)\b/);
  if (qtyMatch) {
    patch.quantity = parseFloat(qtyMatch[1]);
    patch.unit = qtyMatch[2].toUpperCase();
  }

  // Coats pattern
  if (lower.includes("prime") && lower.includes("2")) patch.coats = "prime+2";
  else if (lower.match(/\b2\s*coats?\b/) || lower.includes("color change")) patch.coats = "2coats";
  else if (lower.match(/\b1\s*coats?\b/) || lower.includes("refresh") || lower.includes("recoat")) patch.coats = "1coat";

  // Dimensions: "20 × 14 × 9" or "20x14x9" or "20 by 14 by 9"
  const dimMatch = text.match(/(\d+(?:\.\d+)?)\s*[×x\*]\s*(\d+(?:\.\d+)?)(?:\s*[×x\*]\s*(\d+(?:\.\d+)?))?/i);
  if (dimMatch) {
    patch.length = parseFloat(dimMatch[1]);
    patch.width = parseFloat(dimMatch[2]);
    if (dimMatch[3]) patch.height = parseFloat(dimMatch[3]);
  }

  // Door count
  const doorMatch = lower.match(/(\d+)\s*doors?/);
  if (doorMatch && field === "doorCount") {
    patch.doors = { enabled: true, count: parseInt(doorMatch[1]), coats: "2coats" };
  }

  // Size keywords for cabinets
  if (field === "size") {
    if (lower.includes("small")) patch.notes = "Size: Small";
    else if (lower.includes("medium")) patch.notes = "Size: Medium";
    else if (lower.includes("large")) patch.notes = "Size: Large";
  }

  // Hours for T&M
  const hoursMatch = lower.match(/(\d+(?:\.\d+)?)\s*(?:hours?|hrs?)/);
  if (hoursMatch) patch.hours = parseFloat(hoursMatch[1]);

  return Object.keys(patch).length > 0 ? patch : null;
}

function SettingsButton({ current, value, onClick, children }) {
  return (
    <button
      onClick={() => onClick(value)}
      style={{
        textAlign: "left",
        padding: "6px 8px",
        fontSize: 12,
        borderRadius: 6,
        border: "none",
        background: current === value ? "var(--color-background-secondary, #f3f3f0)" : "transparent",
        cursor: "pointer",
        color: "inherit",
        fontFamily: "inherit",
        fontWeight: current === value ? 500 : 400,
      }}
    >
      {children}
    </button>
  );
}
