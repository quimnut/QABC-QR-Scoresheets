import type { StyleSet, StyleSetConfidence, StyleSetDetectionResult, EntryType, Entry } from "./types";

// ─── Embedded style set data (ported from style_sets.py) ─────────────────

export const STYLE_SETS: Record<string, StyleSet> = {
  BJCP2008: {
    id: 0,
    name: "BJCP2008",
    long_name: "Beer Judge Certification Program 2008",
    short_name: "BJCP 2008",
    beer_end: 23,
    mead: ["24", "25", "26"],
    cider: ["27", "28"],
    categories: {
      "1": "Light Lager", "2": "Pilsner", "3": "European Amber Lager",
      "4": "Dark Lager", "5": "Bock", "6": "Light Hybrid Beer",
      "7": "Amber Hybrid Beer", "8": "English Pale Ale",
      "9": "Scottish and Irish Ale", "10": "American Ale",
      "11": "English Brown Ale", "12": "Porter", "13": "Stout",
      "14": "India Pale Ale (IPA)", "15": "German Wheat and Rye Beer",
      "16": "Belgian and French Ale", "17": "Sour Ale",
      "18": "Belgian Strong Ale", "19": "Strong Ale", "20": "Fruit Beer",
      "21": "Spice/Herb/Vegetable Beer",
      "22": "Smoke-Flavored and Wood-Aged Beer", "23": "Specialty Beer",
      "24": "Traditional Mead", "25": "Melomel (Fruit Mead)", "26": "Other Mead",
      "27": "Standard Cider and Perry", "28": "Specialty Cider and Perry",
    },
  },
  BJCP2015: {
    id: 1,
    name: "BJCP2015",
    long_name: "BJCP Beer, Mead, and Cider 2015",
    short_name: "BJCP 2015",
    beer_end: 34,
    mead: ["M1", "M2", "M3", "M4"],
    cider: ["C1", "C2"],
    categories: {
      "1": "Standard American Beer", "2": "International Lager",
      "3": "Czech Lager", "4": "Pale Malty European Lager",
      "5": "Pale Bitter European Beer", "6": "Amber Malty European Lager",
      "7": "Amber Bitter European Beer", "8": "Dark European Lager",
      "9": "Strong European Beer", "10": "German Wheat Beer",
      "11": "British Bitter", "12": "Pale Commonwealth Beer",
      "13": "Brown British Beer", "14": "Scottish Ale", "15": "Irish Beer",
      "16": "Dark British Beer", "17": "Strong British Ale",
      "18": "Pale American Ale", "19": "Amber and Brown American Beer",
      "20": "American Porter and Stout", "21": "IPA",
      "22": "Strong American Ale", "23": "European Sour Ale",
      "24": "Belgian Ale", "25": "Strong Belgian Ale", "26": "Monastic Ale",
      "27": "Historical Beer", "28": "American Wild Ale", "29": "Fruit Beer",
      "30": "Spiced Beer", "31": "Alternative Fermentables Beer",
      "32": "Smoked Beer", "33": "Wood Beer", "34": "Specialty Beer",
      M1: "Traditional Mead", M2: "Fruit Mead", M3: "Spiced Mead", M4: "Specialty Mead",
      C1: "Standard Cider and Perry", C2: "Specialty Cider and Perry",
    },
  },
  BJCP2021: {
    id: 2,
    name: "BJCP2021",
    long_name: "BJCP Beer 2021, Mead and Cider 2015",
    short_name: "BJCP 2015 / 2021",
    beer_end: 34,
    mead: ["M1", "M2", "M3", "M4"],
    cider: ["C1", "C2"],
    categories: {
      "1": "Standard American Beer", "2": "International Lager",
      "3": "Czech Lager", "4": "Pale Malty European Lager",
      "5": "Pale Bitter European Beer", "6": "Amber Malty European Lager",
      "7": "Amber Bitter European Beer", "8": "Dark European Lager",
      "9": "Strong European Beer", "10": "German Wheat Beer",
      "11": "British Bitter", "12": "Pale Commonwealth Beer",
      "13": "Brown British Beer", "14": "Scottish Ale", "15": "Irish Beer",
      "16": "Dark British Beer", "17": "Strong British Ale",
      "18": "Pale American Ale", "19": "Amber and Brown American Beer",
      "20": "American Porter and Stout", "21": "IPA",
      "22": "Strong American Ale", "23": "European Sour Ale",
      "24": "Belgian Ale", "25": "Strong Belgian Ale", "26": "Monastic Ale",
      "27": "Historical Beer", "28": "American Wild Ale", "29": "Fruit Beer",
      "30": "Spiced Beer", "31": "Alternative Fermentables Beer",
      "32": "Smoked Beer", "33": "Wood Beer", "34": "Specialty Beer",
      M1: "Traditional Mead", M2: "Fruit Mead", M3: "Spiced Mead", M4: "Specialty Mead",
      C1: "Standard Cider and Perry", C2: "Specialty Cider and Perry",
    },
  },
  BJCP2025: {
    id: 3,
    name: "BJCP2025",
    long_name: "BJCP 2021 / 2025",
    short_name: "BJCP 2021 / 2025",
    beer_end: 34,
    mead: ["M1", "M2", "M3", "M4"],
    cider: ["C1", "C2", "C3", "C4"],
    categories: {
      "1": "Standard American Beer", "2": "International Lager",
      "3": "Czech Lager", "4": "Pale Malty European Lager",
      "5": "Pale Bitter European Beer", "6": "Amber Malty European Lager",
      "7": "Amber Bitter European Beer", "8": "Dark European Lager",
      "9": "Strong European Beer", "10": "German Wheat Beer",
      "11": "British Bitter", "12": "Pale Commonwealth Beer",
      "13": "Brown British Beer", "14": "Scottish Ale", "15": "Irish Beer",
      "16": "Dark British Beer", "17": "Strong British Ale",
      "18": "Pale American Ale", "19": "Amber and Brown American Beer",
      "20": "American Porter and Stout", "21": "IPA",
      "22": "Strong American Ale", "23": "European Sour Ale",
      "24": "Belgian Ale", "25": "Strong Belgian Ale", "26": "Monastic Ale",
      "27": "Historical Beer", "28": "American Wild Ale", "29": "Fruit Beer",
      "30": "Spiced Beer", "31": "Alternative Fermentables Beer",
      "32": "Smoked Beer", "33": "Wood Beer", "34": "Specialty Beer",
      LS: "Local Styles",
      M1: "Traditional Mead", M2: "Fruit Mead", M3: "Spiced Mead", M4: "Specialty Mead",
      C1: "Traditional Cider", C2: "Strong Cider", C3: "Specialty Cider", C4: "Perry",
    },
  },
  BA: {
    id: 4,
    name: "BA",
    long_name: "Brewers Association",
    short_name: "BA",
    beer_end: 11,
    mead: ["12"],
    cider: ["12"],
    categories: {
      "1": "British Origin Ales", "2": "Irish Origin Ales",
      "3": "North American Origin Ales", "4": "German Origin Ales",
      "5": "Belgian And French Origin Ales", "6": "International Ale Styles",
      "7": "European-Germanic Origin Lagers", "8": "North American Origin Lagers",
      "9": "Other Origin Lagers", "10": "International Lagers",
      "11": "Hybrid/Mixed Beer", "12": "Mead, Cider, & Perry",
    },
  },
  AABC2022: {
    id: 5,
    name: "AABC2022",
    long_name: "Australian Amateur Brewing Championship 2022",
    short_name: "AABC 2022",
    beer_end: 18,
    mead: ["19"],
    cider: ["20"],
    categories: {
      "1": "Pale Lager", "2": "Pilsner", "3": "Amber & Dark Lager",
      "4": "Pale Ale", "5": "IPA", "6": "Dark Ale", "7": "Wheat Beer",
      "8": "Belgian & French Ale", "9": "Sour & Wild Ale", "10": "Strong Ale",
      "11": "British Bitter & Scottish Ale", "12": "Porter & Stout",
      "13": "Strong & Dark Lager", "14": "Wood & Smoke Flavoured Beer",
      "15": "Specialty Beer", "16": "Fruit & Spice",
      "17": "Alternative Grain, Sugar & Starch", "18": "New World Beer",
      "19": "Mead", "20": "Cider & Perry",
    },
  },
  AABC2025: {
    id: 7,
    name: "AABC2025",
    long_name: "Australian Amateur Brewing Championship 2025",
    short_name: "AABC 2025",
    beer_end: 18,
    mead: ["19"],
    cider: ["20"],
    categories: {
      "1": "Pale Lager", "2": "Pilsner", "3": "Amber & Dark Lager",
      "4": "Pale Ale", "5": "IPA", "6": "Dark Ale", "7": "Wheat Beer",
      "8": "Belgian & French Ale", "9": "Sour & Wild Ale", "10": "Strong Ale",
      "11": "British Bitter & Scottish Ale", "12": "Porter & Stout",
      "13": "Strong & Dark Lager", "14": "Wood & Smoke Flavoured Beer",
      "15": "Specialty Beer", "16": "Fruit & Spice",
      "17": "Alternative Grain, Sugar & Starch", "18": "New World Beer",
      "19": "Mead", "20": "Cider & Perry",
    },
  },
};

// Return a style set by numeric id (used by workers that receive serialised id)
export function getStyleSetById(id: number): StyleSet {
  return Object.values(STYLE_SETS).find((s) => s.id === id) ?? STYLE_SETS.BJCP2021;
}

// ─── Detection ────────────────────────────────────────────────────────────

export function detectStyleSet(entries: Entry[]): StyleSetDetectionResult {
  if (!entries.length) {
    return {
      styleSet: STYLE_SETS.BJCP2021,
      confidence: "low",
      warnings: ["No entries to analyse — defaulting to BJCP2021"],
      name: STYLE_SETS.BJCP2021.long_name,
    };
  }

  const categories = new Set<string>();
  for (const e of entries) {
    const cat = String(e.Category ?? "").trim().toUpperCase();
    if (cat) categories.add(cat);
  }

  if (!categories.size) {
    return {
      styleSet: STYLE_SETS.BJCP2021,
      confidence: "low",
      warnings: ["No categories found — defaulting to BJCP2021"],
      name: STYLE_SETS.BJCP2021.long_name,
    };
  }

  const hasAlpha = [...categories].some((c) => !/^\d+$/.test(c));
  const numericCats = [...categories].filter((c) => /^\d+$/.test(c)).map(Number);
  const maxNum = numericCats.length ? Math.max(...numericCats) : 0;
  const minNum = numericCats.length ? Math.min(...numericCats) : 0;

  if (hasAlpha) {
    if (categories.has("C3") || categories.has("C4")) {
      return { styleSet: STYLE_SETS.BJCP2025, confidence: "high", warnings: [], name: STYLE_SETS.BJCP2025.long_name };
    }
    if (categories.has("M1") || categories.has("C1")) {
      return { styleSet: STYLE_SETS.BJCP2021, confidence: "high", warnings: [], name: STYLE_SETS.BJCP2021.long_name };
    }
  }

  if (maxNum <= 20 && minNum >= 1 && (numericCats.includes(19) || numericCats.includes(20))) {
    return { styleSet: STYLE_SETS.AABC2025, confidence: "high", warnings: [], name: STYLE_SETS.AABC2025.long_name };
  }

  if (maxNum <= 28 && maxNum > 23 && [24, 25, 26, 27, 28].some((n) => numericCats.includes(n))) {
    return { styleSet: STYLE_SETS.BJCP2008, confidence: "high", warnings: [], name: STYLE_SETS.BJCP2008.long_name };
  }

  if (maxNum <= 34 && !hasAlpha) {
    return {
      styleSet: STYLE_SETS.BJCP2021,
      confidence: "medium",
      warnings: ["Detected numeric BJCP categories (1-34) — assuming BJCP2021"],
      name: STYLE_SETS.BJCP2021.long_name,
    };
  }

  if (maxNum <= 12 && minNum >= 1) {
    return {
      styleSet: STYLE_SETS.BA,
      confidence: "medium",
      warnings: ["Detected small category range (1-12) — possibly BA"],
      name: STYLE_SETS.BA.long_name,
    };
  }

  return {
    styleSet: STYLE_SETS.BJCP2021,
    confidence: "low",
    warnings: [`Unknown style set — categories found: ${[...categories].sort().join(", ")}`, "Defaulting to BJCP2021"],
    name: STYLE_SETS.BJCP2021.long_name,
  };
}

// ─── Category / entry type ────────────────────────────────────────────────

export function getCategoryType(category: string | number, styleSet: StyleSet): EntryType {
  const cat = String(category).trim().toUpperCase();
  if (styleSet.mead.map((m) => m.toUpperCase()).includes(cat)) return "mead";
  if (styleSet.cider.map((c) => c.toUpperCase()).includes(cat)) return "cider";
  return "beer";
}

export function getEntryType(entry: Entry, styleSet: StyleSet): EntryType {
  // Tier 1 — explicit Style Type field
  const st = (entry["Style Type"] ?? "").trim().toLowerCase();
  if (st === "mead") return "mead";
  if (st === "cider" || st === "perry") return "cider";
  if (st === "beer") return "beer";

  // Tier 2 — category lookup
  const cat = entry.Category ?? "";
  if (cat) {
    const t = getCategoryType(cat, styleSet);
    if (t !== "beer") return t;
  }

  // Tier 3 — keyword scan across text fields
  const combined = [
    entry.Style ?? "",
    entry.Subcategory ?? "",
    entry["Sub Category Name"] ?? "",
    entry.SubCategoryName ?? "",
  ]
    .join(" ")
    .toLowerCase();

  if (combined.includes("mead")) return "mead";
  if (["cider", "perry", "applewine"].some((kw) => combined.includes(kw))) return "cider";

  return "beer";
}
